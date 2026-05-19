from __future__ import annotations

import os

import argostranslate.package
import argostranslate.translate


def normalize(code: str) -> str:
    return (code or "").strip().lower().split("-")[0]


def ensure_package(source: str, target: str) -> None:
    installed = argostranslate.translate.get_installed_languages()
    from_lang = next((lang for lang in installed if lang.code == source), None)
    to_lang = next((lang for lang in installed if lang.code == target), None)
    if from_lang and to_lang:
        try:
            from_lang.get_translation(to_lang)
            return
        except Exception:
            pass

    available = argostranslate.package.get_available_packages()
    pkg = next((p for p in available if p.from_code == source and p.to_code == target), None)
    if not pkg:
        raise RuntimeError(f"No package found for {source}->{target}")
    path = pkg.download()
    argostranslate.package.install_from_path(path)


def main() -> None:
    pairs_env = os.getenv("ARGOS_PREWARM_BUILD_PAIRS", "en-hi,hi-en,en-es,es-en,en-fr,fr-en")
    pairs = []
    for raw in pairs_env.split(","):
        raw = raw.strip()
        if not raw or "-" not in raw:
            continue
        src, dst = raw.split("-", 1)
        src = normalize(src)
        dst = normalize(dst)
        if src and dst and src != "auto" and dst != "auto":
            pairs.append((src, dst))

    if not pairs:
        return

    argostranslate.package.update_package_index()
    for src, dst in pairs:
        try:
            ensure_package(src, dst)
            print(f"Prewarmed {src}->{dst}")
        except Exception as exc:
            print(f"Skipped {src}->{dst}: {exc}")


if __name__ == "__main__":
    main()
