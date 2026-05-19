from __future__ import annotations

import os
import time
from concurrent.futures import ThreadPoolExecutor, TimeoutError as FutureTimeoutError
from threading import Thread

from flask import Flask, jsonify, request
from flask_cors import CORS

import argostranslate.package
import argostranslate.translate


app = Flask(__name__)
CORS(app)
_EXECUTOR = ThreadPoolExecutor(max_workers=2)


def _normalize_lang(code: str | None) -> str:
    value = (code or "").strip().lower()
    if not value:
        return "auto"
    return value.split("-")[0]


def _ensure_package(source: str, target: str) -> None:
    installed = argostranslate.translate.get_installed_languages()
    from_lang = next((lang for lang in installed if lang.code == source), None)
    to_lang = next((lang for lang in installed if lang.code == target), None)
    if from_lang and to_lang and from_lang.get_translation(to_lang):
        return

    argostranslate.package.update_package_index()
    available = argostranslate.package.get_available_packages()
    pkg = next(
        (p for p in available if p.from_code == source and p.to_code == target),
        None,
    )
    if not pkg:
        raise ValueError(f"No Argos package available for {source} -> {target}")
    path = pkg.download()
    argostranslate.package.install_from_path(path)


def _installed_codes() -> set[str]:
    return {lang.code for lang in argostranslate.translate.get_installed_languages()}


def _ensure_common_pairs() -> None:
    # Pairs can be overridden by env var, example: "en-hi,hi-en,en-es,es-en"
    configured = os.getenv("ARGOS_PREWARM_PAIRS", "en-hi,hi-en,en-es,es-en,en-fr,fr-en")
    pairs = []
    for raw in configured.split(","):
        raw = raw.strip()
        if not raw or "-" not in raw:
            continue
        src, dst = raw.split("-", 1)
        src = _normalize_lang(src)
        dst = _normalize_lang(dst)
        if src == "auto" or dst == "auto":
            continue
        pairs.append((src, dst))

    for src, dst in pairs:
        try:
            _ensure_package(src, dst)
        except Exception:
            # Keep startup resilient even if one package fails.
            continue


def _start_prewarm_thread() -> None:
    t = Thread(target=_ensure_common_pairs, daemon=True)
    t.start()


def _detect_lang(text: str) -> str:
    installed = argostranslate.translate.get_installed_languages()
    if not installed:
        raise ValueError("No Argos language packages installed.")

    # Prefer languages that can translate to English for simple detection.
    candidates = []
    english = next((lang for lang in installed if lang.code == "en"), None)
    for lang in installed:
        if english and lang.code != "en":
            try:
                tr = lang.get_translation(english)
                candidates.append((lang.code, tr.translate(text[:300])))
            except Exception:
                continue
    if not candidates:
        # Fallback to first installed language when detection cannot be inferred.
        return installed[0].code

    # Heuristic: translation output with highest ASCII ratio is likely closest.
    def score(item: tuple[str, str]) -> float:
        output = item[1] or ""
        if not output:
            return 0.0
        ascii_count = sum(1 for ch in output if ord(ch) < 128)
        return ascii_count / len(output)

    return max(candidates, key=score)[0]


@app.get("/api/health")
def health():
    return jsonify({"ok": True, "engine": "argos-translate"})


@app.get("/api/languages")
def languages():
    installed = argostranslate.translate.get_installed_languages()
    rows = []
    for lang in installed:
        targets = []
        for other in installed:
            if other.code == lang.code:
                continue
            try:
                lang.get_translation(other)
                targets.append(other.code)
            except Exception:
                continue
        rows.append(
            {
                "code": lang.code,
                "name": lang.name or lang.code,
                "targets": sorted(targets),
            }
        )
    rows.sort(key=lambda row: row["name"].lower())
    return jsonify({"languages": rows, "count": len(rows)})


def _translate_once(text: str, source: str, target: str) -> str:
    _ensure_package(source, target)
    return argostranslate.translate.translate(text, source, target)


@app.post("/api/translate")
def translate():
    payload = request.get_json(silent=True) or {}
    q = str(payload.get("q", "")).strip()
    source = _normalize_lang(payload.get("source"))
    target = _normalize_lang(payload.get("target"))

    if not q:
        return jsonify({"error": "Missing 'q' text"}), 400
    if not target or target == "auto":
        return jsonify({"error": "Target language is required"}), 400

    timeout_sec = int(os.getenv("ARGOS_TRANSLATE_TIMEOUT_SEC", "45"))
    retries = int(os.getenv("ARGOS_TRANSLATE_RETRIES", "2"))

    try:
        if source == "auto":
            source = _detect_lang(q)

        translated = ""
        last_error = None
        for attempt in range(retries + 1):
            try:
                future = _EXECUTOR.submit(_translate_once, q, source, target)
                translated = future.result(timeout=timeout_sec)
                break
            except FutureTimeoutError as exc:
                last_error = exc
                if attempt < retries:
                    time.sleep(1.1 * (attempt + 1))
                    continue
                raise TimeoutError(
                    f"Translation timed out after {timeout_sec}s (attempt {attempt + 1})."
                )
            except Exception as exc:
                last_error = exc
                if attempt < retries:
                    time.sleep(0.7 * (attempt + 1))
                    continue
                raise

        if not translated and last_error is not None:
            raise RuntimeError(str(last_error))

        return jsonify(
            {
                "translatedText": translated,
                "detectedLanguage": {"language": source},
                "engine": "argos-translate",
            }
        )
    except Exception as exc:
        return jsonify({"error": str(exc)}), 500


if __name__ == "__main__":
    _start_prewarm_thread()
    app.run(host="0.0.0.0", port=5000, debug=False)
