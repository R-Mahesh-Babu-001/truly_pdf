import { useMemo, useState, useEffect, useRef } from 'react'
import { PDFDocument, ParseSpeeds, StandardFonts, degrees, rgb } from 'pdf-lib'
import PptxGenJS from 'pptxgenjs'
import * as XLSX from 'xlsx'
import JSZip from 'jszip'
import mammoth from 'mammoth/mammoth.browser.js'
import fileSaver from 'file-saver'
import * as pdfjsLib from 'pdfjs-dist/build/pdf'
import pdfjsWorker from 'pdfjs-dist/build/pdf.worker?url'

pdfjsLib.GlobalWorkerOptions.workerSrc = pdfjsWorker

const CHATIFY_MESSAGE_TYPE = 'chatify:pdf-result'
const CHATIFY_ALLOWED_ORIGINS = new Set([
  'http://localhost:5173',
  'http://127.0.0.1:5173',
  'https://chatify-1-6nr8.onrender.com',
  'https://chatify-frontend-3uzh.onrender.com',
])

const getChatifyBridge = () => {
  const params = new URLSearchParams(window.location.search)
  const targetOrigin = params.get('chatifyOrigin')
  const operationId = params.get('chatifyOperationId')

  if (!window.opener || !targetOrigin || !operationId) return null

  try {
    const normalizedOrigin = new URL(targetOrigin).origin
    if (!CHATIFY_ALLOWED_ORIGINS.has(normalizedOrigin)) return null
    return { targetOrigin: normalizedOrigin, operationId }
  } catch {
    return null
  }
}

const getInitialPathname = () => {
  const params = new URLSearchParams(window.location.search)
  const chatifyTool = params.get('chatifyTool')

  if (chatifyTool?.startsWith('/tool/')) {
    const slug = chatifyTool.slice('/tool/'.length)
    if (tools.some((tool) => tool.slug === slug)) {
      return chatifyTool
    }
  }

  return window.location.pathname
}

const categoryRoutes = [
  { label: 'All PDF Tools', path: '/', key: 'all' },
  { label: 'Convert', path: '/convert', key: 'convert' },
  { label: 'Organize', path: '/organize', key: 'organize' },
  { label: 'Optimize', path: '/optimize', key: 'optimize' },
  { label: 'Security', path: '/security', key: 'security' },
  { label: 'eSign', path: '/esign', key: 'esign' },
]

const navRoutes = [
  { label: 'Merge PDF', path: '/tool/merge-pdf' },
  { label: 'Split PDF', path: '/tool/split-pdf' },
  { label: 'Compress PDF', path: '/tool/compress-pdf' },
  { label: 'Convert PDF', path: '/convert' },
  { label: 'All PDF Tools', path: '/' },
]

const tools = [
  { title: 'Merge PDF', slug: 'merge-pdf', description: 'Combine PDFs in the order you want with the easiest PDF merger available.', icon: 'merge', category: 'organize' },
  { title: 'Split PDF', slug: 'split-pdf', description: 'Separate one page or a whole set for easy conversion into independent PDF files.', icon: 'split', category: 'organize' },
  { title: 'Compress PDF', slug: 'compress-pdf', description: 'Reduce file size while optimizing for maximal PDF quality.', icon: 'compress', category: 'optimize' },
  { title: 'PDF to Word', slug: 'pdf-to-word', description: 'Convert PDF files into editable DOC and DOCX documents in seconds.', icon: 'convert', category: 'convert' },
  { title: 'PDF to PowerPoint', slug: 'pdf-to-powerpoint', description: 'Turn PDF files into editable PPT and PPTX slideshows.', icon: 'convert', category: 'convert' },
  { title: 'PDF to Excel', slug: 'pdf-to-excel', description: 'Pull data straight from PDFs into Excel spreadsheets.', icon: 'convert', category: 'convert' },
  { title: 'Word to PDF', slug: 'word-to-pdf', description: 'Make DOC and DOCX files easy to read by converting them to PDF.', icon: 'convert', category: 'convert' },
  { title: 'PowerPoint to PDF', slug: 'powerpoint-to-pdf', description: 'Make PPT and PPTX slideshows easy to share by converting them to PDF.', icon: 'convert', category: 'convert' },
  { title: 'Excel to PDF', slug: 'excel-to-pdf', description: 'Make spreadsheets easy to read by converting them to PDF.', icon: 'convert', category: 'convert' },
  { title: 'Edit PDF', slug: 'edit-pdf', description: 'Add text, images, shapes, and annotations to a PDF document.', icon: 'edit', category: 'organize' },
  { title: 'PDF to JPG', slug: 'pdf-to-jpg', description: 'Convert each PDF page into a JPG or extract all images from a PDF.', icon: 'image', category: 'convert' },
  { title: 'JPG to PDF', slug: 'jpg-to-pdf', description: 'Convert JPG images to PDF in seconds with margin and orientation controls.', icon: 'image', category: 'convert' },
  { title: 'Sign PDF', slug: 'sign-pdf', description: 'Sign yourself or request electronic signatures from others.', icon: 'sign', category: 'esign' },
  { title: 'Watermark', slug: 'watermark-pdf', description: 'Stamp image or text over your PDF in seconds.', icon: 'watermark', category: 'security' },
  { title: 'Rotate PDF', slug: 'rotate-pdf', description: 'Rotate PDFs exactly the way you need.', icon: 'rotate', category: 'organize' },
  { title: 'HTML to PDF', slug: 'html-to-pdf', description: 'Convert webpages in HTML to PDF by using a URL.', icon: 'convert', category: 'convert' },
  { title: 'Unlock PDF', slug: 'unlock-pdf', description: 'Remove PDF password security and unlock files.', icon: 'unlock', category: 'security' },
  { title: 'Protect PDF', slug: 'protect-pdf', description: 'Protect PDF files with a password and encryption.', icon: 'lock', category: 'security' },
  { title: 'Organize PDF', slug: 'organize-pdf', description: 'Sort, delete, and add pages inside your PDF document.', icon: 'organize', category: 'organize' },
  { title: 'PDF to PDF/A', slug: 'pdf-to-pdfa', description: 'Transform your PDF to PDF/A for long-term archiving.', icon: 'archive', category: 'optimize' },
  { title: 'Repair PDF', slug: 'repair-pdf', description: 'Repair damaged PDF files and recover data.', icon: 'repair', category: 'optimize' },
  { title: 'Page numbers', slug: 'page-numbers', description: 'Add page numbers into PDFs with full position controls.', icon: 'numbers', category: 'organize' },
  { title: 'Scan to PDF', slug: 'scan-to-pdf', description: 'Capture document scans from your device and send them to your browser.', icon: 'scan', category: 'convert' },
  { title: 'OCR PDF', slug: 'ocr-pdf', description: 'Convert scanned PDF files into searchable and selectable documents.', icon: 'ocr', category: 'optimize' },
  { title: 'Compare PDF', slug: 'compare-pdf', description: 'Show side-by-side document comparison and spot all changes.', icon: 'compare', category: 'organize' },
  { title: 'Redact PDF', slug: 'redact-pdf', description: 'Permanently remove sensitive information from a PDF.', icon: 'redact', category: 'security' },
  { title: 'Crop PDF', slug: 'crop-pdf', description: 'Crop PDF margins or select specific page areas.', icon: 'crop', category: 'organize' },
  { title: 'AI Summarizer', slug: 'ai-summarizer', description: 'Generate concise summaries from long PDFs and reports.', icon: 'spark', category: 'optimize', isNew: true },
  { title: 'Translate PDF', slug: 'translate-pdf', description: 'Translate PDF files while preserving structure and readability.', icon: 'translate', category: 'convert', isNew: true },
]

const fallbackTranslateLanguages = [
  { code: 'auto', label: 'Detect language' },
  { code: 'en', label: 'English' },
  { code: 'es', label: 'Spanish' },
  { code: 'fr', label: 'French' },
  { code: 'de', label: 'German' },
  { code: 'it', label: 'Italian' },
  { code: 'pt', label: 'Portuguese' },
  { code: 'ru', label: 'Russian' },
  { code: 'ar', label: 'Arabic' },
  { code: 'hi', label: 'Hindi' },
  { code: 'bn', label: 'Bengali' },
  { code: 'ja', label: 'Japanese' },
  { code: 'ko', label: 'Korean' },
  { code: 'zh', label: 'Chinese' },
  { code: 'tr', label: 'Turkish' },
  { code: 'nl', label: 'Dutch' },
  { code: 'pl', label: 'Polish' },
  { code: 'uk', label: 'Ukrainian' },
]

function ToolGlyph({ type }) {
  const glyphs = {
    merge: '⇄',
    split: '✂',
    compress: '⇣',
    convert: '↔',
    edit: '✎',
    image: '▣',
    sign: '✍',
    watermark: '◎',
    rotate: '↻',
    unlock: '◔',
    lock: '●',
    organize: '☰',
    archive: '⌂',
    repair: '✚',
    numbers: '#',
    scan: '▤',
    ocr: '◉',
    compare: '⇆',
    redact: '■',
    crop: '⌗',
    spark: '✦',
    translate: '⌘',
  }

  return (
    <div className="glyph-wrap">
      <span className="glyph-page" />
      <span className="glyph-symbol">{glyphs[type] || '•'}</span>
    </div>
  )
}

function getRouteState(pathname) {
  const categoryRoute = categoryRoutes.find((route) => route.path === pathname)
  if (categoryRoute) {
    return {
      title: categoryRoute.label,
      chipPath: categoryRoute.path,
      filter: (tool) => categoryRoute.key === 'all' || tool.category === categoryRoute.key,
    }
  }

  const prefix = '/tool/'
  if (pathname.startsWith(prefix)) {
    const slug = pathname.slice(prefix.length)
    const tool = tools.find((item) => item.slug === slug)
    if (tool) {
      return {
        title: tool.title,
        chipPath: '/',
        tool,
        filter: (item) => item.slug === slug,
      }
    }
  }

  return {
    title: 'All PDF Tools',
    chipPath: '/',
    filter: () => true,
  }
}

function App() {
  const [pathname, setPathname] = useState(getInitialPathname)
  const [selectedFiles, setSelectedFiles] = useState([])
  const [isDragOver, setIsDragOver] = useState(false)
  const [draggingIndex, setDraggingIndex] = useState(null)
  const [dragOverIndex, setDragOverIndex] = useState(null)
  const [isMerging, setIsMerging] = useState(false)
  const [mergedFileUrl, setMergedFileUrl] = useState('')
  const [mergedFileName, setMergedFileName] = useState('')
  const [splitMode, setSplitMode] = useState('range')
  const [splitRangeType, setSplitRangeType] = useState('custom')
  const [splitFromPage, setSplitFromPage] = useState(1)
  const [splitToPage, setSplitToPage] = useState(2)
  const [splitBusy, setSplitBusy] = useState(false)
  const [splitResultUrl, setSplitResultUrl] = useState('')
  const [splitResultName, setSplitResultName] = useState('')
  const [splitTotalPages, setSplitTotalPages] = useState(0)
  const [pageNumFileIndex, setPageNumFileIndex] = useState(0)
  const [pageNumMode, setPageNumMode] = useState('single')
  const [pageNumPosition, setPageNumPosition] = useState('br')
  const [pageNumMargin, setPageNumMargin] = useState('recommended')
  const [pageNumFirst, setPageNumFirst] = useState(1)
  const [pageNumTextMode, setPageNumTextMode] = useState('number')
  const [pageNumFontFamily, setPageNumFontFamily] = useState('helvetica')
  const [pageNumBold, setPageNumBold] = useState(false)
  const [pageNumItalic, setPageNumItalic] = useState(false)
  const [pageNumUnderline, setPageNumUnderline] = useState(false)
  const [pageNumBusy, setPageNumBusy] = useState(false)
  const [pageNumStatus, setPageNumStatus] = useState('')
  const [compareMode, setCompareMode] = useState('semantic')
  const [compareScrollSync, setCompareScrollSync] = useState(true)
  const [compareZoom, setCompareZoom] = useState(100)
  const [compareLeftFileIndex, setCompareLeftFileIndex] = useState(0)
  const [compareRightFileIndex, setCompareRightFileIndex] = useState(1)
  const [compareLeftPage, setCompareLeftPage] = useState(1)
  const [compareRightPage, setCompareRightPage] = useState(1)
  const [compareSearch, setCompareSearch] = useState('')
  const [compareReport, setCompareReport] = useState([])
  const [compareStatus, setCompareStatus] = useState('')
  const [compareBusy, setCompareBusy] = useState(false)
  const [compareLeftProxy, setCompareLeftProxy] = useState(null)
  const [compareRightProxy, setCompareRightProxy] = useState(null)
  const [compareLeftPages, setCompareLeftPages] = useState(1)
  const [compareRightPages, setCompareRightPages] = useState(1)
  const [cropPdfProxy, setCropPdfProxy] = useState(null)
  const [cropPage, setCropPage] = useState(1)
  const [cropPageCount, setCropPageCount] = useState(1)
  const [cropZoom, setCropZoom] = useState(100)
  const [cropScope, setCropScope] = useState('all')
  const [cropRect, setCropRect] = useState(null)
  const [cropBusy, setCropBusy] = useState(false)
  const [cropStatus, setCropStatus] = useState('')
  const [jpgExtractMode, setJpgExtractMode] = useState('page')
  const [jpgQualityMode, setJpgQualityMode] = useState('normal')
  const [scanPages, setScanPages] = useState([])
  const [scanSelectedId, setScanSelectedId] = useState('')
  const [scanCameraOn, setScanCameraOn] = useState(false)
  const [scanBusy, setScanBusy] = useState(false)
  const [scanStatus, setScanStatus] = useState('')
  const [htmlToPdfUrl, setHtmlToPdfUrl] = useState('')
  const [htmlToPdfStatus, setHtmlToPdfStatus] = useState('')
  const [cropCanvasSize, setCropCanvasSize] = useState({ width: 1, height: 1 })
  const [organizePages, setOrganizePages] = useState([])
  const [organizeBusy, setOrganizeBusy] = useState(false)
  const [organizeStatus, setOrganizeStatus] = useState('')
  const [organizeDragIndex, setOrganizeDragIndex] = useState(null)
  const [organizeBaseSize, setOrganizeBaseSize] = useState({ width: 595, height: 842 })
  const [organizeSortDirection, setOrganizeSortDirection] = useState('asc')
  const [aiSummaryPrompt, setAiSummaryPrompt] = useState('Summarize this file in simple language.')
  const [aiSummaryMessages, setAiSummaryMessages] = useState([
    {
      id: `assistant-${Date.now()}`,
      role: 'assistant',
      text: 'Upload a file and ask for the summary style you want.',
    },
  ])
  const [translateSourceLang, setTranslateSourceLang] = useState('auto')
  const [translateTargetLang, setTranslateTargetLang] = useState('en')
  const [translateStatus, setTranslateStatus] = useState('')
  const [translatedText, setTranslatedText] = useState('')
  const [translateLanguages, setTranslateLanguages] = useState(fallbackTranslateLanguages)
  const [rotateAnglesByFile, setRotateAnglesByFile] = useState({})
  const [rotateSelectedIndex, setRotateSelectedIndex] = useState(0)
  const [rotateBusy, setRotateBusy] = useState(false)
  const [rotateStatus, setRotateStatus] = useState('')
  const [editTool, setEditTool] = useState('annotate')
  const [editSelectedIndex, setEditSelectedIndex] = useState(0)
  const [editPage, setEditPage] = useState(1)
  const [editZoom, setEditZoom] = useState(85)
  const [editFitScale, setEditFitScale] = useState(1)
  const [editPageCount, setEditPageCount] = useState(1)
  const [editPdfData, setEditPdfData] = useState(null)
  const [editPdfProxy, setEditPdfProxy] = useState(null)
  const [editCanvasMetaByPage, setEditCanvasMetaByPage] = useState({})
  const [editObjectsByPage, setEditObjectsByPage] = useState({})
  const [editSelectedObjectId, setEditSelectedObjectId] = useState('')
  const [editPendingImageData, setEditPendingImageData] = useState('')
  const [editStatus, setEditStatus] = useState('')
  const fileInputRef = useRef(null)
  const editImageInputRef = useRef(null)
  const editBaseCanvasRef = useRef(null)
  const editOverlayCanvasRef = useRef(null)
  const editCanvasWrapRef = useRef(null)
  const editImageCacheRef = useRef({})
  const editActivePathRef = useRef(null)
  const editDragRef = useRef(null)
  const editObjectsRef = useRef({})
  const compareLeftCanvasRef = useRef(null)
  const compareRightCanvasRef = useRef(null)
  const compareOverlayCanvasRef = useRef(null)
  const compareLeftScrollRef = useRef(null)
  const compareRightScrollRef = useRef(null)
  const cropCanvasRef = useRef(null)
  const cropDragRef = useRef(null)
  const scanVideoRef = useRef(null)
  const scanCaptureCanvasRef = useRef(null)
  const scanUploadInputRef = useRef(null)
  const scanStreamRef = useRef(null)
  const chatifyBridgeRef = useRef(getChatifyBridge())
  const addFilesRef = useRef(null)

  const downloadBlob = (blobLike, filename) => {
    const blob = blobLike instanceof Blob ? blobLike : new Blob([blobLike])
    const bridge = chatifyBridgeRef.current
    if (bridge) {
      window.opener.postMessage(
        {
          type: CHATIFY_MESSAGE_TYPE,
          operationId: bridge.operationId,
          toolSlug: activeTool?.slug || '',
          filename,
          fileType: blob.type || 'application/octet-stream',
          fileSize: blob.size,
          blob,
        },
        bridge.targetOrigin,
      )
    }
    if (typeof fileSaver?.saveAs === 'function') {
      fileSaver.saveAs(blob, filename)
      return
    }
    const url = URL.createObjectURL(blob)
    const link = document.createElement('a')
    link.href = url
    link.download = filename
    document.body.appendChild(link)
    link.click()
    document.body.removeChild(link)
    window.setTimeout(() => URL.revokeObjectURL(url), 60_000)
  }

  const downloadBlobUrl = async (url, filename) => {
    const response = await fetch(url)
    const blob = await response.blob()
    downloadBlob(blob, filename)
  }

  useEffect(() => {
    const bridge = chatifyBridgeRef.current
    if (!bridge) return

    window.opener.postMessage(
      {
        type: 'chatify:pdf-ready',
        operationId: bridge.operationId,
      },
      bridge.targetOrigin,
    )
  }, [])

  useEffect(() => {
    const onPopState = () => setPathname(window.location.pathname)
    window.addEventListener('popstate', onPopState)
    return () => window.removeEventListener('popstate', onPopState)
  }, [])

  const navigate = (event, path) => {
    event.preventDefault()
    if (window.location.pathname !== path) {
      window.history.pushState({}, '', path)
      setPathname(path)
    }
    const toolsSection = document.getElementById('tools')
    if (toolsSection) {
      toolsSection.scrollIntoView({ behavior: 'smooth', block: 'start' })
    }
  }

  const goToPath = (path) => {
    if (window.location.pathname !== path) {
      window.history.pushState({}, '', path)
      setPathname(path)
    }
  }

  const routeState = useMemo(() => getRouteState(pathname), [pathname])
  const visibleTools = useMemo(
    () => tools.filter(routeState.filter),
    [routeState],
  )
  const activeTool = routeState.tool
  const selectedScanPage = useMemo(
    () => scanPages.find((page) => page.id === scanSelectedId) || scanPages[0] || null,
    [scanPages, scanSelectedId],
  )
  const previewUrls = useMemo(
    () => selectedFiles.map((file) => URL.createObjectURL(file)),
    [selectedFiles],
  )

  useEffect(() => {
    return () => {
      previewUrls.forEach((url) => URL.revokeObjectURL(url))
    }
  }, [previewUrls])

  useEffect(() => {
    return () => {
      if (mergedFileUrl) {
        URL.revokeObjectURL(mergedFileUrl)
      }
    }
  }, [mergedFileUrl])

  useEffect(() => {
    return () => {
      if (splitResultUrl) {
        URL.revokeObjectURL(splitResultUrl)
      }
    }
  }, [splitResultUrl])

  useEffect(() => {
    setSelectedFiles([])
    setIsDragOver(false)
    setDraggingIndex(null)
    setDragOverIndex(null)
    setIsMerging(false)
    if (mergedFileUrl) {
      URL.revokeObjectURL(mergedFileUrl)
    }
    setMergedFileUrl('')
    setMergedFileName('')
    setSplitMode('range')
    setSplitRangeType('custom')
    setSplitFromPage(1)
    setSplitToPage(2)
    setSplitBusy(false)
    if (splitResultUrl) {
      URL.revokeObjectURL(splitResultUrl)
    }
    setSplitResultUrl('')
    setSplitResultName('')
    setSplitTotalPages(0)
    setPageNumFileIndex(0)
    setPageNumMode('single')
    setPageNumPosition('br')
    setPageNumMargin('recommended')
    setPageNumFirst(1)
    setPageNumTextMode('number')
    setPageNumFontFamily('helvetica')
    setPageNumBold(false)
    setPageNumItalic(false)
    setPageNumUnderline(false)
    setPageNumBusy(false)
    setPageNumStatus('')
    setCompareMode('semantic')
    setCompareScrollSync(true)
    setCompareZoom(100)
    setCompareLeftFileIndex(0)
    setCompareRightFileIndex(1)
    setCompareLeftPage(1)
    setCompareRightPage(1)
    setCompareSearch('')
    setCompareReport([])
    setCompareStatus('')
    setCompareBusy(false)
    setCompareLeftProxy(null)
    setCompareRightProxy(null)
    setCompareLeftPages(1)
    setCompareRightPages(1)
    setCropPdfProxy(null)
    setCropPage(1)
    setCropPageCount(1)
    setCropZoom(100)
    setCropScope('all')
    setCropRect(null)
    setCropBusy(false)
    setCropStatus('')
    setJpgExtractMode('page')
    setJpgQualityMode('normal')
    stopScanCamera()
    setScanPages([])
    setScanSelectedId('')
    setScanCameraOn(false)
    setScanBusy(false)
    setScanStatus('')
    setHtmlToPdfUrl('')
    setHtmlToPdfStatus('')
    setCropCanvasSize({ width: 1, height: 1 })
    setOrganizePages([])
    setOrganizeBusy(false)
    setOrganizeStatus('')
    setOrganizeDragIndex(null)
    setOrganizeBaseSize({ width: 595, height: 842 })
    setOrganizeSortDirection('asc')
    setRotateAnglesByFile({})
    setRotateSelectedIndex(0)
    setRotateBusy(false)
    setRotateStatus('')
    setEditTool('annotate')
    setEditSelectedIndex(0)
    setEditPage(1)
    setEditZoom(85)
    setEditFitScale(1)
    setEditPageCount(1)
    setEditPdfData(null)
    setEditPdfProxy(null)
    setEditCanvasMetaByPage({})
    setEditObjectsByPage({})
    editObjectsRef.current = {}
    setEditSelectedObjectId('')
    setEditPendingImageData('')
    setEditStatus('')
    setAiSummaryPrompt('Summarize this file in simple language.')
    setAiSummaryMessages([
      {
        id: `assistant-${Date.now()}`,
        role: 'assistant',
        text: 'Upload a file and ask for the summary style you want.',
      },
    ])
    setTranslateSourceLang('auto')
    setTranslateTargetLang('en')
    setTranslateStatus('')
    setTranslatedText('')
    setTranslateLanguages(fallbackTranslateLanguages)
  }, [activeTool?.slug])

  useEffect(() => {
    if (activeTool?.slug !== 'translate-pdf') return
    let cancelled = false
    ;(async () => {
      try {
        const res = await fetch(`${getTranslateApiBase()}/languages`)
        if (!res.ok) return
        const data = await res.json()
        const rows = Array.isArray(data?.languages) ? data.languages : []
        if (!rows.length || cancelled) return

        const mapped = rows.map((lang) => ({
          code: String(lang.code || '').trim(),
          label: String(lang.name || lang.code || '').trim(),
          targets: Array.isArray(lang.targets) ? lang.targets.map((t) => String(t).trim()) : [],
        })).filter((lang) => lang.code && lang.label)

        const options = [
          { code: 'auto', label: 'Detect language', targets: mapped.map((lang) => lang.code) },
          ...mapped,
        ]
        setTranslateLanguages(options)

        if (!options.some((lang) => lang.code === translateTargetLang)) {
          const firstTarget = mapped.find((lang) => lang.code !== 'auto')?.code || 'en'
          setTranslateTargetLang(firstTarget)
        }
      } catch {
        // Keep fallback list when API is unavailable.
      }
    })()
    return () => {
      cancelled = true
    }
  }, [activeTool?.slug, translateTargetLang])

  useEffect(() => {
    if (!scanPages.length) {
      if (scanSelectedId) setScanSelectedId('')
      return
    }
    const hasSelected = scanPages.some((page) => page.id === scanSelectedId)
    if (!hasSelected) {
      setScanSelectedId(scanPages[0].id)
    }
  }, [scanPages, scanSelectedId])

  useEffect(() => {
    if (activeTool?.slug !== 'split-pdf' || !selectedFiles.length) {
      setSplitTotalPages(0)
      return
    }

    let mounted = true
    ;(async () => {
      try {
        const data = await selectedFiles[0].arrayBuffer()
        const pdf = await PDFDocument.load(data)
        const total = pdf.getPageCount()
        if (!mounted) return
        setSplitTotalPages(total)
        setSplitFromPage(1)
        setSplitToPage(Math.min(2, total))
      } catch {
        if (mounted) {
          setSplitTotalPages(0)
        }
      }
    })()

    return () => {
      mounted = false
    }
  }, [activeTool?.slug, selectedFiles])

  useEffect(() => {
    if (splitRangeType === 'fixed' && splitTotalPages > 0) {
      setSplitFromPage(1)
      setSplitToPage(splitTotalPages)
    }
  }, [splitRangeType, splitTotalPages])

  useEffect(() => {
    if (activeTool?.slug !== 'organize-pdf' || !selectedFiles.length) {
      setOrganizePages([])
      return
    }

    let cancelled = false
    ;(async () => {
      try {
        const data = await selectedFiles[0].arrayBuffer()
        const pdf = await PDFDocument.load(data, { ignoreEncryption: true })
        if (cancelled) return
        const pages = pdf.getPages()
        if (pages.length) {
          const first = pages[0].getSize()
          setOrganizeBaseSize({ width: first.width, height: first.height })
        }
        setOrganizePages(
          pages.map((_, idx) => ({
            id: crypto.randomUUID(),
            type: 'source',
            sourcePageIndex: idx,
            rotation: 0,
            label: String(idx + 1),
          })),
        )
        setOrganizeStatus('')
        setOrganizeSortDirection('asc')
      } catch {
        if (!cancelled) {
          setOrganizePages([])
          setOrganizeStatus('Could not read PDF pages.')
        }
      }
    })()

    return () => {
      cancelled = true
    }
  }, [activeTool?.slug, selectedFiles])

  useEffect(() => {
    if (editSelectedIndex >= selectedFiles.length) {
      setEditSelectedIndex(0)
    }
  }, [editSelectedIndex, selectedFiles.length])

  useEffect(() => {
    editObjectsRef.current = editObjectsByPage
  }, [editObjectsByPage])

  const getCanvasPoint = (event) => {
    const canvas = editOverlayCanvasRef.current
    if (!canvas) return { x: 0, y: 0 }
    const rect = canvas.getBoundingClientRect()
    const scaleX = canvas.width / rect.width
    const scaleY = canvas.height / rect.height
    return {
      x: (event.clientX - rect.left) * scaleX,
      y: (event.clientY - rect.top) * scaleY,
    }
  }

  const updateEditFitScale = () => {
    const wrap = editCanvasWrapRef.current
    const canvas = editBaseCanvasRef.current
    if (!wrap || !canvas?.width || !canvas?.height) return
    const availableWidth = Math.max(1, wrap.clientWidth - 16)
    const availableHeight = Math.max(1, wrap.clientHeight - 16)
    const fit = Math.min(availableWidth / canvas.width, availableHeight / canvas.height, 1)
    setEditFitScale(fit)
  }

  const hitTestEditObject = (objects, x, y) => {
    for (let i = objects.length - 1; i >= 0; i -= 1) {
      const obj = objects[i]
      if (obj.type === 'text') {
        const w = obj.w || 140
        const h = obj.h || 30
        if (x >= obj.x && x <= obj.x + w && y >= obj.y - h && y <= obj.y) return obj.id
      }
      if (obj.type === 'rect' || obj.type === 'image') {
        if (x >= obj.x && x <= obj.x + obj.w && y >= obj.y && y <= obj.y + obj.h) return obj.id
      }
      if (obj.type === 'path' && obj.points?.length) {
        const xs = obj.points.map((p) => p.x)
        const ys = obj.points.map((p) => p.y)
        const minX = Math.min(...xs) - 8
        const maxX = Math.max(...xs) + 8
        const minY = Math.min(...ys) - 8
        const maxY = Math.max(...ys) + 8
        if (x >= minX && x <= maxX && y >= minY && y <= maxY) return obj.id
      }
    }
    return ''
  }

  const formatSize = (bytes) => {
    if (!bytes) {
      return '0 KB'
    }
    const kb = bytes / 1024
    if (kb < 1024) {
      return `${kb.toFixed(1)} KB`
    }
    return `${(kb / 1024).toFixed(2)} MB`
  }

  const isPdfFile = (file) => {
    if (!file) {
      return false
    }
    const hasPdfMime = file.type === 'application/pdf'
    const hasPdfExt = file.name?.toLowerCase().endsWith('.pdf')
    return hasPdfMime || hasPdfExt
  }

  const isWordFile = (file) => {
    if (!file) {
      return false
    }
    const lowerName = file.name?.toLowerCase() || ''
    const mime = file.type || ''
    return (
      lowerName.endsWith('.docx') ||
      lowerName.endsWith('.doc') ||
      mime === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' ||
      mime === 'application/msword'
    )
  }

  const isPowerPointFile = (file) => {
    if (!file) return false
    const lowerName = file.name?.toLowerCase() || ''
    const mime = file.type || ''
    return (
      lowerName.endsWith('.pptx') ||
      lowerName.endsWith('.ppt') ||
      mime === 'application/vnd.openxmlformats-officedocument.presentationml.presentation' ||
      mime === 'application/vnd.ms-powerpoint'
    )
  }

  const isExcelFile = (file) => {
    if (!file) return false
    const lowerName = file.name?.toLowerCase() || ''
    const mime = file.type || ''
    return (
      lowerName.endsWith('.xlsx') ||
      lowerName.endsWith('.xls') ||
      lowerName.endsWith('.csv') ||
      mime === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' ||
      mime === 'application/vnd.ms-excel' ||
      mime === 'text/csv'
    )
  }

  const isImageFile = (file) => {
    if (!file) return false
    const lowerName = file.name?.toLowerCase() || ''
    const mime = file.type || ''
    return (
      mime.startsWith('image/') ||
      lowerName.endsWith('.jpg') ||
      lowerName.endsWith('.jpeg') ||
      lowerName.endsWith('.png')
    )
  }

  const isSummarizerFile = (file) => {
    if (!file) return false
    const lowerName = file.name?.toLowerCase() || ''
    const mime = file.type || ''
    return (
      isPdfFile(file) ||
      isWordFile(file) ||
      isPowerPointFile(file) ||
      isExcelFile(file) ||
      mime.startsWith('text/') ||
      mime === 'application/json' ||
      lowerName.endsWith('.txt') ||
      lowerName.endsWith('.md') ||
      lowerName.endsWith('.rtf') ||
      lowerName.endsWith('.json')
    )
  }

  const allowedInputAccept = activeTool?.slug === 'word-to-pdf'
    ? '.doc,.docx,application/msword,application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    : activeTool?.slug === 'powerpoint-to-pdf'
      ? '.ppt,.pptx,application/vnd.ms-powerpoint,application/vnd.openxmlformats-officedocument.presentationml.presentation'
      : activeTool?.slug === 'excel-to-pdf'
        ? '.xls,.xlsx,.csv,application/vnd.ms-excel,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,text/csv'
        : activeTool?.slug === 'jpg-to-pdf'
          ? '.jpg,.jpeg,.png,image/jpeg,image/png'
        : activeTool?.slug === 'ai-summarizer'
          ? '.pdf,.txt,.md,.doc,.docx,.ppt,.pptx,.xls,.xlsx,.csv,.json,text/*,application/json,application/pdf,application/msword,application/vnd.openxmlformats-officedocument.wordprocessingml.document,application/vnd.ms-powerpoint,application/vnd.openxmlformats-officedocument.presentationml.presentation,application/vnd.ms-excel,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,text/csv'
        : '.pdf,application/pdf'

  const selectFileLabel = activeTool?.slug === 'word-to-pdf'
    ? 'Select WORD file'
    : activeTool?.slug === 'powerpoint-to-pdf'
      ? 'Select POWERPOINT file'
      : activeTool?.slug === 'excel-to-pdf'
        ? 'Select EXCEL file'
        : activeTool?.slug === 'jpg-to-pdf'
          ? 'Select JPG file'
        : activeTool?.slug === 'ai-summarizer'
          ? 'Upload file to summarize'
        : 'Select PDF file'
  const dropHintLabel = activeTool?.slug === 'word-to-pdf'
    ? 'or drop WORD here'
    : activeTool?.slug === 'powerpoint-to-pdf'
      ? 'or drop POWERPOINT here'
      : activeTool?.slug === 'excel-to-pdf'
        ? 'or drop EXCEL here'
        : activeTool?.slug === 'jpg-to-pdf'
          ? 'or drop JPG here'
        : activeTool?.slug === 'ai-summarizer'
          ? 'or drop PDF/DOC/PPT/XLS/TXT here'
        : 'or drop PDF here'

  const addFiles = (incomingFiles) => {
    if (!incomingFiles?.length) {
      return
    }

    if (activeTool?.slug === 'scan-to-pdf') {
      Array.from(incomingFiles).forEach(async (file) => {
        if (!isImageFile(file)) return
        const dataUrl = await new Promise((resolve, reject) => {
          const reader = new FileReader()
          reader.onload = () => resolve(reader.result)
          reader.onerror = reject
          reader.readAsDataURL(file)
        })
        pushScanImage(String(dataUrl), file.name)
      })
      return
    }

    const candidates = Array.from(incomingFiles).filter((file) => {
      if (activeTool?.slug === 'word-to-pdf') return isWordFile(file)
      if (activeTool?.slug === 'powerpoint-to-pdf') return isPowerPointFile(file)
      if (activeTool?.slug === 'excel-to-pdf') return isExcelFile(file)
      if (activeTool?.slug === 'jpg-to-pdf') return isImageFile(file)
      if (activeTool?.slug === 'ai-summarizer') return isSummarizerFile(file)
      return isPdfFile(file)
    })
    if (!candidates.length) {
      return
    }

    setSelectedFiles((prev) => {
      const next = [...prev]
      candidates.forEach((file) => {
        const duplicate = next.some(
          (existing) =>
            existing.name === file.name &&
            existing.size === file.size &&
            existing.lastModified === file.lastModified,
        )
        if (!duplicate) {
          next.push(file)
        }
      })
      return next
    })
  }

  addFilesRef.current = addFiles

  useEffect(() => {
    const bridge = chatifyBridgeRef.current
    if (!bridge) return undefined

    const onChatifyMessage = (event) => {
      if (event.origin !== bridge.targetOrigin) return
      if (event.data?.type !== 'chatify:pdf-input') return
      if (event.data?.operationId !== bridge.operationId) return
      if (!(event.data?.file instanceof File)) return
      addFilesRef.current?.([event.data.file])
    }

    window.addEventListener('message', onChatifyMessage)
    return () => window.removeEventListener('message', onChatifyMessage)
  }, [])

  const onInputChange = (event) => {
    addFiles(event.target.files)
    event.target.value = ''
  }

  const openFilePicker = () => {
    fileInputRef.current?.click()
  }

  const onDragOver = (event) => {
    event.preventDefault()
    setIsDragOver(true)
  }

  const onDragLeave = (event) => {
    event.preventDefault()
    setIsDragOver(false)
  }

  const onDrop = (event) => {
    event.preventDefault()
    setIsDragOver(false)
    addFiles(event.dataTransfer.files)
  }

  const stopScanCamera = () => {
    const stream = scanStreamRef.current
    if (stream) {
      stream.getTracks().forEach((track) => track.stop())
      scanStreamRef.current = null
    }
    if (scanVideoRef.current) {
      scanVideoRef.current.srcObject = null
    }
    setScanCameraOn(false)
  }

  const startScanCamera = async () => {
    try {
      stopScanCamera()
      const stream = await navigator.mediaDevices.getUserMedia({
        video: { facingMode: { ideal: 'environment' } },
        audio: false,
      })
      scanStreamRef.current = stream
      if (scanVideoRef.current) {
        scanVideoRef.current.srcObject = stream
      }
      setScanCameraOn(true)
      setScanStatus('')
    } catch {
      setScanStatus('Camera access denied or unavailable. You can still add images from files.')
      setScanCameraOn(false)
    }
  }

  const pushScanImage = (dataUrl, label = '') => {
    const id = crypto.randomUUID()
    setScanPages((prev) => [
      ...prev,
      {
        id,
        dataUrl,
        label,
        crop: { left: 0, top: 0, right: 0, bottom: 0 },
      },
    ])
    setScanSelectedId(id)
  }

  const captureScanPage = () => {
    const video = scanVideoRef.current
    const canvas = scanCaptureCanvasRef.current
    if (!video || !canvas || !video.videoWidth || !video.videoHeight) {
      setScanStatus('Camera frame is not ready yet.')
      return
    }
    canvas.width = video.videoWidth
    canvas.height = video.videoHeight
    const context = canvas.getContext('2d', { alpha: false })
    if (!context) {
      setScanStatus('Could not capture camera frame.')
      return
    }
    context.drawImage(video, 0, 0, canvas.width, canvas.height)
    const dataUrl = canvas.toDataURL('image/jpeg', 0.92)
    pushScanImage(dataUrl, `Scan ${scanPages.length + 1}`)
  }

  const onScanFileInput = async (event) => {
    const files = Array.from(event.target.files || [])
    for (const file of files) {
      if (!file.type.startsWith('image/')) continue
      const dataUrl = await new Promise((resolve, reject) => {
        const reader = new FileReader()
        reader.onload = () => resolve(reader.result)
        reader.onerror = reject
        reader.readAsDataURL(file)
      })
      pushScanImage(String(dataUrl), file.name)
    }
    event.target.value = ''
  }

  const updateSelectedScanCrop = (key, value) => {
    setScanPages((prev) =>
      prev.map((page) => {
        if (page.id !== scanSelectedId) return page
        const next = { ...page.crop, [key]: Math.max(0, Math.min(45, Number(value))) }
        if (next.left + next.right > 90) {
          if (key === 'left') next.right = Math.max(0, 90 - next.left)
          else next.left = Math.max(0, 90 - next.right)
        }
        if (next.top + next.bottom > 90) {
          if (key === 'top') next.bottom = Math.max(0, 90 - next.top)
          else next.top = Math.max(0, 90 - next.bottom)
        }
        return { ...page, crop: next }
      }),
    )
  }

  const removeScanPage = (id) => {
    setScanPages((prev) => {
      const next = prev.filter((page) => page.id !== id)
      setScanSelectedId((current) => (current === id ? (next[0]?.id || '') : current))
      return next
    })
  }

  const exportScanToPdf = async () => {
    if (!scanPages.length) {
      setScanStatus('Add at least one scan page first.')
      return
    }
    try {
      setScanBusy(true)
      const outputPdf = await PDFDocument.create()
      for (let idx = 0; idx < scanPages.length; idx += 1) {
        const pageItem = scanPages[idx]
        const image = await new Promise((resolve, reject) => {
          const img = new Image()
          img.onload = () => resolve(img)
          img.onerror = reject
          img.src = pageItem.dataUrl
        })

        const sourceCanvas = document.createElement('canvas')
        sourceCanvas.width = image.width
        sourceCanvas.height = image.height
        const sourceCtx = sourceCanvas.getContext('2d', { alpha: false })
        sourceCtx.drawImage(image, 0, 0)

        const { left, top, right, bottom } = pageItem.crop
        const sx = Math.floor((left / 100) * sourceCanvas.width)
        const sy = Math.floor((top / 100) * sourceCanvas.height)
        const sw = Math.max(1, Math.floor(((100 - left - right) / 100) * sourceCanvas.width))
        const sh = Math.max(1, Math.floor(((100 - top - bottom) / 100) * sourceCanvas.height))

        const cropCanvas = document.createElement('canvas')
        cropCanvas.width = sw
        cropCanvas.height = sh
        const cropCtx = cropCanvas.getContext('2d', { alpha: false })
        cropCtx.drawImage(sourceCanvas, sx, sy, sw, sh, 0, 0, sw, sh)

        const blob = await new Promise((resolve) => cropCanvas.toBlob(resolve, 'image/jpeg', 0.92))
        if (!blob) continue
        const bytes = await blob.arrayBuffer()
        const embedded = await outputPdf.embedJpg(bytes)

        const pageWidth = 595.28
        const pageHeight = 841.89
        const margin = 20
        const maxW = pageWidth - margin * 2
        const maxH = pageHeight - margin * 2
        const scale = Math.min(maxW / embedded.width, maxH / embedded.height)
        const drawW = embedded.width * scale
        const drawH = embedded.height * scale
        const x = (pageWidth - drawW) / 2
        const y = (pageHeight - drawH) / 2
        const pdfPage = outputPdf.addPage([pageWidth, pageHeight])
        pdfPage.drawImage(embedded, { x, y, width: drawW, height: drawH })
      }

      const pdfBytes = await outputPdf.save()
      const filename = `scanned-${Date.now()}.pdf`
      downloadBlob(new File([pdfBytes], filename, { type: 'application/pdf' }), filename)
      setScanStatus('PDF exported.')
    } catch {
      setScanStatus('Could not export scanned PDF.')
    } finally {
      setScanBusy(false)
    }
  }

  useEffect(() => {
    if (activeTool?.slug === 'scan-to-pdf') {
      startScanCamera()
      return () => {
        stopScanCamera()
      }
    }
    stopScanCamera()
  }, [activeTool?.slug])

  const onAddHtmlSource = () => {
    const raw = htmlToPdfUrl.trim()
    if (!raw) {
      setHtmlToPdfStatus('Please enter a website URL.')
      return
    }
    const normalized = /^https?:\/\//i.test(raw) ? raw : `https://${raw}`
    try {
      // Validate URL input for consistent UX.
      // eslint-disable-next-line no-new
      new URL(normalized)
      setHtmlToPdfUrl(normalized)
      setHtmlToPdfStatus('URL added.')
    } catch {
      setHtmlToPdfStatus('Please enter a valid URL.')
    }
  }

  const removeFileAt = (indexToRemove) => {
    setSelectedFiles((prev) => prev.filter((_, index) => index !== indexToRemove))
  }

  const moveFile = (fromIndex, toIndex) => {
    if (fromIndex === toIndex || fromIndex < 0 || toIndex < 0) {
      return
    }
    setSelectedFiles((prev) => {
      const next = [...prev]
      const [moved] = next.splice(fromIndex, 1)
      next.splice(toIndex, 0, moved)
      return next
    })
  }

  const onTileDragStart = (event, index) => {
    setDraggingIndex(index)
    event.dataTransfer.effectAllowed = 'move'
    event.dataTransfer.setData('text/plain', String(index))
  }

  const onTileDragOver = (event, index) => {
    event.preventDefault()
    event.dataTransfer.dropEffect = 'move'
    if (dragOverIndex !== index) {
      setDragOverIndex(index)
    }
  }

  const onTileDrop = (event, index) => {
    event.preventDefault()
    const raw = event.dataTransfer.getData('text/plain')
    const fromIndex = Number(raw)
    if (!Number.isNaN(fromIndex)) {
      moveFile(fromIndex, index)
    }
    setDraggingIndex(null)
    setDragOverIndex(null)
  }

  const onTileDragEnd = () => {
    setDraggingIndex(null)
    setDragOverIndex(null)
  }

  const toolShortName = activeTool?.title
    .replace('PDF to ', '')
    .replace(' to PDF', '')
    .replace('PDF', '')
    .trim()

  const convertLabel = activeTool?.slug === 'merge-pdf'
    ? 'Merge PDF'
    : activeTool?.slug === 'ai-summarizer'
      ? 'Generate summary'
    : activeTool?.slug === 'translate-pdf'
      ? 'Translate'
    : activeTool?.slug === 'word-to-pdf' || activeTool?.slug === 'powerpoint-to-pdf' || activeTool?.slug === 'excel-to-pdf' || activeTool?.slug === 'jpg-to-pdf'
      ? 'Convert to PDF'
    : toolShortName
      ? `Convert to ${toolShortName.toUpperCase()}`
      : 'Convert'

  const resetMergeResult = () => {
    if (mergedFileUrl) {
      URL.revokeObjectURL(mergedFileUrl)
    }
    setMergedFileUrl('')
    setMergedFileName('')
  }

  const decodeXmlEntities = (text) =>
    String(text || '')
      .replace(/&amp;/g, '&')
      .replace(/&lt;/g, '<')
      .replace(/&gt;/g, '>')
      .replace(/&quot;/g, '"')
      .replace(/&apos;/g, "'")
      .replace(/&#(\d+);/g, (_, num) => String.fromCharCode(Number(num)))
      .replace(/&#x([0-9a-fA-F]+);/g, (_, hex) => String.fromCharCode(parseInt(hex, 16)))

  const extractTextForSummarizer = async (file) => {
    const raw = await file.arrayBuffer()
    const bytes = new Uint8Array(raw)
    const lowerName = file.name?.toLowerCase() || ''

    if (isPdfFile(file)) {
      const pdf = await pdfjsLib.getDocument({ data: bytes }).promise
      const pages = []
      for (let pageNum = 1; pageNum <= pdf.numPages; pageNum += 1) {
        const page = await pdf.getPage(pageNum)
        const content = await page.getTextContent()
        const pageText = content.items
          .map((item) => (typeof item.str === 'string' ? item.str.trim() : ''))
          .filter(Boolean)
          .join(' ')
        if (pageText) pages.push(`Page ${pageNum}: ${pageText}`)
      }
      return pages.join('\n')
    }

    if (isWordFile(file)) {
      const isZipDocx = bytes.length >= 4
        && bytes[0] === 0x50
        && bytes[1] === 0x4b
        && (bytes[2] === 0x03 || bytes[2] === 0x05 || bytes[2] === 0x07)
        && (bytes[3] === 0x04 || bytes[3] === 0x06 || bytes[3] === 0x08)
      if (isZipDocx) {
        const { value } = await mammoth.extractRawText({ arrayBuffer: raw })
        return String(value || '')
      }
      let decoded = new TextDecoder('utf-8', { fatal: false }).decode(bytes)
      if (decoded.includes('\u0000')) {
        decoded = new TextDecoder('utf-16le', { fatal: false }).decode(bytes)
      }
      const container = document.createElement('div')
      container.innerHTML = decoded
      return String(container.textContent || decoded || '')
    }

    if (isPowerPointFile(file)) {
      const isZip = bytes.length >= 4
        && bytes[0] === 0x50
        && bytes[1] === 0x4b
        && (bytes[2] === 0x03 || bytes[2] === 0x05 || bytes[2] === 0x07)
        && (bytes[3] === 0x04 || bytes[3] === 0x06 || bytes[3] === 0x08)
      if (isZip) {
        const zip = await JSZip.loadAsync(raw)
        const slideEntries = Object.keys(zip.files)
          .filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name))
          .sort((a, b) => Number(a.match(/slide(\d+)\.xml$/)?.[1] || 0) - Number(b.match(/slide(\d+)\.xml$/)?.[1] || 0))
        const slides = []
        for (const entry of slideEntries) {
          const xml = await zip.file(entry).async('text')
          const paragraphMatches = xml.match(/<a:p[\s\S]*?<\/a:p>/g) || []
          const lines = paragraphMatches
            .map((paragraph) => [...paragraph.matchAll(/<a:t[^>]*>([\s\S]*?)<\/a:t>/g)].map((m) => decodeXmlEntities(m[1])).join(' '))
            .map((line) => line.replace(/\s+/g, ' ').trim())
            .filter(Boolean)
          if (lines.length) slides.push(lines.join('\n'))
        }
        return slides.join('\n\n')
      }
    }

    if (isExcelFile(file)) {
      const workbook = XLSX.read(raw, { type: 'array' })
      const sheetNames = workbook.SheetNames || []
      const chunks = []
      sheetNames.forEach((sheetName) => {
        const sheet = workbook.Sheets[sheetName]
        const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false, defval: '' })
        const rowText = rows
          .map((row) => row.map((cell) => String(cell ?? '').trim()).filter(Boolean).join(' | '))
          .filter(Boolean)
          .join('\n')
        if (rowText) chunks.push(`Sheet: ${sheetName}\n${rowText}`)
      })
      return chunks.join('\n\n')
    }

    const text = new TextDecoder('utf-8', { fatal: false }).decode(bytes)
    if (lowerName.endsWith('.json')) {
      try {
        return JSON.stringify(JSON.parse(text), null, 2)
      } catch {
        return text
      }
    }
    return text
  }

  const requestGeminiSummary = async ({ apiKey, prompt }) => {
    const models = [
      'gemini-1.5-pro',
      'gemini-1.5-flash',
      'gemini-2.5-flash',
    ]
    let lastError = 'Unknown Gemini error.'
    for (const model of models) {
      for (let attempt = 1; attempt <= 3; attempt += 1) {
        const response = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${encodeURIComponent(apiKey)}`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            contents: [{ parts: [{ text: prompt }] }],
          }),
        })

        if (response.ok) {
          const data = await response.json()
          const summary = (data?.candidates?.[0]?.content?.parts || [])
            .map((part) => (typeof part?.text === 'string' ? part.text : ''))
            .join('\n')
            .trim()
          if (summary) {
            return { summary, model }
          }
          lastError = `Model ${model} returned empty output.`
          break
        }

        const errText = await response.text()
        lastError = `Gemini request failed (${response.status}) on ${model}: ${errText || 'unknown error'}`
        if (response.status !== 503) {
          break
        }
        await new Promise((resolve) => setTimeout(resolve, 600 * attempt))
      }
    }
    throw new Error(lastError)
  }

  const downloadSummaryChat = () => {
    const chatText = aiSummaryMessages
      .map((msg) => `${msg.role === 'user' ? 'User' : 'Assistant'}:\n${msg.text}`)
      .join('\n\n')
      .trim()
    if (!chatText) {
      window.alert('No chat messages to download yet.')
      return
    }
    const fileName = `ai-summary-chat-${Date.now()}.txt`
    downloadBlob(new Blob([chatText], { type: 'text/plain;charset=utf-8' }), fileName)
  }

  const getTranslateApiBase = () => {
    const endpoint = (import.meta.env.VITE_TRANSLATE_API_URL || 'http://localhost:5000/api/translate').trim()
    return endpoint.replace(/\/translate\/?$/i, '')
  }

  const downloadTranslatedResult = () => {
    if (!translatedText.trim()) {
      window.alert('No translated text available yet.')
      return
    }
    const sourceFile = selectedFiles[0]
    const baseName = sourceFile?.name?.replace(/\.[^/.]+$/i, '') || 'translated'
    const fileName = `${baseName}-${translateTargetLang}.txt`
    downloadBlob(new Blob([translatedText], { type: 'text/plain;charset=utf-8' }), fileName)
  }

  const splitTextForTranslate = (text, maxLen = 1800) => {
    const normalized = String(text || '').replace(/\r/g, '').trim()
    if (!normalized) return []
    if (normalized.length <= maxLen) return [normalized]
    const paragraphs = normalized.split(/\n{2,}/).map((part) => part.trim()).filter(Boolean)
    const chunks = []
    let bucket = ''
    for (const para of paragraphs) {
      if (!bucket) {
        bucket = para
        continue
      }
      if ((bucket.length + 2 + para.length) <= maxLen) {
        bucket += `\n\n${para}`
      } else {
        chunks.push(bucket)
        bucket = para
      }
    }
    if (bucket) chunks.push(bucket)
    return chunks
  }

  const translateWithArgos = async ({ text, source, target }) => {
    const endpoint = `${getTranslateApiBase()}/translate`
    const chunks = splitTextForTranslate(text)
    if (!chunks.length) return ''

    const translatedChunks = []
    for (let i = 0; i < chunks.length; i += 1) {
      setTranslateStatus(`Translating chunk ${i + 1}/${chunks.length}...`)
      const payload = {
        q: chunks[i],
        source,
        target,
      }

      let data = null
      let lastErr = ''
      for (let attempt = 1; attempt <= 3; attempt += 1) {
        const controller = new AbortController()
        const timeout = window.setTimeout(() => controller.abort(), 60_000)
        try {
          const response = await fetch(endpoint, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(payload),
            signal: controller.signal,
          })
          if (!response.ok) {
            const err = await response.text()
            lastErr = `Argos translation failed (${response.status}): ${err || 'unknown error'}`
            if (response.status >= 500 && attempt < 3) {
              await new Promise((resolve) => setTimeout(resolve, 700 * attempt))
              continue
            }
            throw new Error(lastErr)
          }
          data = await response.json()
          break
        } catch (error) {
          lastErr = error?.name === 'AbortError'
            ? `Argos translation timed out (attempt ${attempt}/3).`
            : (error?.message || `Argos request failed (attempt ${attempt}/3).`)
          if (attempt < 3) {
            await new Promise((resolve) => setTimeout(resolve, 700 * attempt))
            continue
          }
          throw new Error(lastErr)
        } finally {
          window.clearTimeout(timeout)
        }
      }

      if (!data) {
        throw new Error(lastErr || 'Argos translation failed.')
      }

      const chunkText = String(data?.translatedText || '').trim()
      if (!chunkText) {
        throw new Error('Argos translation returned empty text.')
      }
      translatedChunks.push(chunkText)
    }
    return translatedChunks.join('\n\n')
  }

  const onConvert = async () => {
    if (!selectedFiles.length) return

    try {
      setIsMerging(true)

      if (activeTool?.slug === 'merge-pdf') {
        const outputPdf = await PDFDocument.create()

        for (const file of selectedFiles) {
          const data = await file.arrayBuffer()
          const srcPdf = await PDFDocument.load(data)
          const copiedPages = await outputPdf.copyPages(srcPdf, srcPdf.getPageIndices())
          copiedPages.forEach((page) => outputPdf.addPage(page))
        }

        const mergedBytes = await outputPdf.save()
        const blob = new Blob([mergedBytes], { type: 'application/pdf' })
        const url = URL.createObjectURL(blob)
        const filename = `merged-${Date.now()}.pdf`

        if (mergedFileUrl) {
          URL.revokeObjectURL(mergedFileUrl)
        }
        setMergedFileUrl(url)
        setMergedFileName(filename)
        return
      }

      if (activeTool?.slug === 'pdf-to-word') {
        const sourceFile = selectedFiles[0]
        const raw = await sourceFile.arrayBuffer()
        const pdf = await pdfjsLib.getDocument({ data: raw }).promise
        const escapeHtml = (value) =>
          String(value || '')
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&#39;')

        const pages = []
        for (let pageNum = 1; pageNum <= pdf.numPages; pageNum += 1) {
          const page = await pdf.getPage(pageNum)
          const viewport = page.getViewport({ scale: 1 })
          const pageWidth = viewport.width
          const pageHeight = viewport.height
          const content = await page.getTextContent()
          const tokens = content.items
            .filter((item) => typeof item.str === 'string' && item.str.trim())
            .map((item) => {
              const t = item.transform || []
              const x = Number(t[4] || 0)
              const y = Number(t[5] || 0)
              const fontSize = Math.max(
                9,
                Math.min(
                  64,
                  Math.abs(Number(t[3] || 0)) ||
                    (typeof item.height === 'number' ? item.height : 12),
                ),
              )
              const top = Math.max(0, pageHeight - y - fontSize)
              return { text: item.str, x, top, fontSize }
            })
          pages.push({ width: pageWidth, height: pageHeight, tokens })
        }

        const hasText = pages.some((page) => page.tokens.length > 0)
        if (!hasText) {
          throw new Error('No extractable text found. This PDF is likely scanned/image-based and needs OCR.')
        }

        const html = `<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8" />
  <title>${escapeHtml(sourceFile.name)}</title>
  <style>
    @page { margin: 0.4in; size: auto; }
    body { margin: 0; background: #fff; }
    .pdf-page {
      position: relative;
      margin: 0 auto 16px;
      border: 1px solid #e7e7e7;
      box-sizing: border-box;
      background: #ffffff;
      page-break-after: always;
    }
    .pdf-page:last-child { page-break-after: auto; }
    .txt {
      position: absolute;
      white-space: pre;
      color: #111;
      font-family: Arial, Helvetica, sans-serif;
      line-height: 1;
    }
  </style>
</head>
<body>
${pages
  .map(
    (page) => `<div class="pdf-page" style="width:${page.width}px;height:${page.height}px;">
${page.tokens
  .map(
    (token) =>
      `<div class="txt" style="left:${token.x}px;top:${token.top}px;font-size:${token.fontSize}px;">${escapeHtml(token.text)}</div>`,
  )
  .join('\n')}
</div>`,
  )
  .join('\n')}
</body>
</html>`

        const blob = new Blob(['\ufeff', html], { type: 'application/msword;charset=utf-8' })
        const filename = `${sourceFile.name.replace(/\.pdf$/i, '') || 'converted'}-pdf-to-word.doc`
        downloadBlob(blob, filename)
        return
      }

      if (activeTool?.slug === 'pdf-to-powerpoint') {
        const sourceFile = selectedFiles[0]
        const raw = await sourceFile.arrayBuffer()
        const pdfBytes = new Uint8Array(raw)
        const pdf = await pdfjsLib.getDocument({ data: pdfBytes }).promise

        const pages = []
        for (let pageNum = 1; pageNum <= pdf.numPages; pageNum += 1) {
          const page = await pdf.getPage(pageNum)
          const viewport = page.getViewport({ scale: 1 })
          const pageWidth = viewport.width
          const pageHeight = viewport.height
          const content = await page.getTextContent()
          const tokens = content.items
            .filter((item) => typeof item.str === 'string' && item.str.trim())
            .map((item) => {
              const t = item.transform || []
              const x = Number(t[4] || 0)
              const y = Number(t[5] || 0)
              const fontSize = Math.max(
                9,
                Math.min(
                  64,
                  Math.abs(Number(t[3] || 0)) ||
                    (typeof item.height === 'number' ? item.height : 12),
                ),
              )
              const top = Math.max(0, pageHeight - y - fontSize)
              return { text: item.str, x, top, fontSize }
            })
          pages.push({ width: pageWidth, height: pageHeight, tokens })
        }

        const hasText = pages.some((page) => page.tokens.length > 0)
        if (!hasText) {
          throw new Error('No extractable text found. This PDF is likely scanned/image-based and needs OCR.')
        }

        const pptx = new PptxGenJS()
        const baseWidthInches = Math.max(4, pages[0].width / 72)
        const baseHeightInches = Math.max(3, pages[0].height / 72)
        pptx.defineLayout({
          name: 'PDF_LAYOUT',
          width: baseWidthInches,
          height: baseHeightInches,
        })
        pptx.layout = 'PDF_LAYOUT'
        pptx.author = 'TRULY PDF'
        pptx.subject = 'PDF to PowerPoint conversion'
        pptx.title = sourceFile.name.replace(/\.pdf$/i, '')

        for (const page of pages) {
          const slide = pptx.addSlide()
          const pageWidthInches = page.width / 72
          const pageHeightInches = page.height / 72
          const scale = Math.min(
            baseWidthInches / pageWidthInches,
            baseHeightInches / pageHeightInches,
          )
          const offsetX = (baseWidthInches - pageWidthInches * scale) / 2
          const offsetY = (baseHeightInches - pageHeightInches * scale) / 2

          for (const token of page.tokens) {
            const x = offsetX + (token.x / 72) * scale
            const y = offsetY + (token.top / 72) * scale
            const fontSize = Math.max(8, Math.min(56, token.fontSize * scale))
            const width = Math.max(0.2, (token.text.length * fontSize * 0.5) / 72)
            const height = Math.max(0.15, (fontSize * 1.25) / 72)
            slide.addText(token.text, {
              x,
              y,
              w: width,
              h: height,
              fontFace: 'Arial',
              fontSize,
              color: '111111',
              breakLine: false,
            })
          }
        }

        const pptBlob = await pptx.write({ outputType: 'blob' })
        const filename = `${sourceFile.name.replace(/\.pdf$/i, '') || 'converted'}-pdf-to-powerpoint.pptx`
        downloadBlob(pptBlob, filename)
        return
      }

      if (activeTool?.slug === 'pdf-to-excel') {
        const sourceFile = selectedFiles[0]
        const raw = await sourceFile.arrayBuffer()
        const pdfBytes = new Uint8Array(raw)
        const pdf = await pdfjsLib.getDocument({ data: pdfBytes }).promise
        const workbook = XLSX.utils.book_new()
        let totalTokens = 0

        for (let pageNum = 1; pageNum <= pdf.numPages; pageNum += 1) {
          const page = await pdf.getPage(pageNum)
          const viewport = page.getViewport({ scale: 1 })
          const pageHeight = viewport.height
          const content = await page.getTextContent()
          const tokens = content.items
            .filter((item) => typeof item.str === 'string' && item.str.trim())
            .map((item) => {
              const t = item.transform || []
              const x = Number(t[4] || 0)
              const y = Number(t[5] || 0)
              const top = Math.max(0, pageHeight - y)
              return { text: item.str.trim(), x, top }
            })

          totalTokens += tokens.length

          const rowTolerance = 8
          const colTolerance = 20
          const rows = []
          for (const token of tokens) {
            let row = rows.find((candidate) => Math.abs(candidate.top - token.top) <= rowTolerance)
            if (!row) {
              row = { top: token.top, cells: [] }
              rows.push(row)
            }
            row.cells.push(token)
          }

          rows.sort((a, b) => a.top - b.top)
          rows.forEach((row) => row.cells.sort((a, b) => a.x - b.x))

          const colAnchors = []
          const aoa = []

          for (const row of rows) {
            const values = []
            for (const cell of row.cells) {
              let colIndex = colAnchors.findIndex((anchor) => Math.abs(anchor - cell.x) <= colTolerance)
              if (colIndex === -1) {
                colAnchors.push(cell.x)
                colAnchors.sort((a, b) => a - b)
                colIndex = colAnchors.findIndex((anchor) => anchor === cell.x)
              }
              values[colIndex] = values[colIndex] ? `${values[colIndex]} ${cell.text}` : cell.text
            }

            let lastUsed = values.length - 1
            while (lastUsed >= 0 && (values[lastUsed] === undefined || values[lastUsed] === '')) {
              lastUsed -= 1
            }
            aoa.push(lastUsed >= 0 ? values.slice(0, lastUsed + 1) : [''])
          }

          const worksheet = XLSX.utils.aoa_to_sheet(aoa.length ? aoa : [['']])
          XLSX.utils.book_append_sheet(workbook, worksheet, `Page ${pageNum}`)
        }

        if (!totalTokens) {
          throw new Error('No extractable text found. This PDF is likely scanned/image-based and needs OCR.')
        }

        const xlsxData = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' })
        const blob = new Blob([xlsxData], {
          type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        })
        const filename = `${sourceFile.name.replace(/\.pdf$/i, '') || 'converted'}-pdf-to-excel.xlsx`
        downloadBlob(blob, filename)
        return
      }

      if (activeTool?.slug === 'pdf-to-jpg') {
        const renderScale = jpgQualityMode === 'high' ? 2.4 : 1.8
        const jpegQuality = jpgQualityMode === 'high' ? 0.94 : 0.82
        const zip = new JSZip()
        let outputCount = 0

        for (const sourceFile of selectedFiles) {
          const raw = await sourceFile.arrayBuffer()
          const pdf = await pdfjsLib.getDocument({ data: new Uint8Array(raw) }).promise
          const baseName = sourceFile.name.replace(/\.pdf$/i, '') || 'pdf'

          for (let pageNum = 1; pageNum <= pdf.numPages; pageNum += 1) {
            const page = await pdf.getPage(pageNum)
            const viewport = page.getViewport({ scale: renderScale })
            const canvas = document.createElement('canvas')
            const context = canvas.getContext('2d', { alpha: false })
            if (!context) {
              throw new Error('Could not create image canvas.')
            }

            canvas.width = Math.max(1, Math.floor(viewport.width))
            canvas.height = Math.max(1, Math.floor(viewport.height))
            await page.render({ canvasContext: context, viewport }).promise

            const imageBlob = await new Promise((resolve) =>
              canvas.toBlob(resolve, 'image/jpeg', jpegQuality),
            )
            if (!imageBlob) {
              throw new Error('Could not export JPG image.')
            }

            const filename = `${baseName}-page-${pageNum}.jpg`
            outputCount += 1

            if (selectedFiles.length === 1 && pdf.numPages === 1) {
              downloadBlob(imageBlob, filename)
              return
            }

            const imageBytes = await imageBlob.arrayBuffer()
            zip.file(filename, imageBytes)
          }
        }

        if (!outputCount) {
          throw new Error('No pages were converted.')
        }

        const zipBlob = await zip.generateAsync({ type: 'blob' })
        const zipName = jpgExtractMode === 'extract'
          ? `extracted-jpg-images-${Date.now()}.zip`
          : `pdf-to-jpg-${Date.now()}.zip`
        downloadBlob(zipBlob, zipName)
        return
      }

      if (activeTool?.slug === 'ai-summarizer') {
        const sourceFile = selectedFiles[0]
        const apiKey = (import.meta.env.VITE_GEMINI_API_KEY || '').trim()
        if (!apiKey) {
          throw new Error('Gemini API key missing. Set VITE_GEMINI_API_KEY in .env.local.')
        }

        const userPrompt = aiSummaryPrompt.trim() || 'Summarize this file in simple language.'
        setAiSummaryMessages((prev) => [
          ...prev,
          {
            id: `user-${Date.now()}`,
            role: 'user',
            text: `${userPrompt}\n\nFile: ${sourceFile.name}`,
          },
        ])
        const extractedText = (await extractTextForSummarizer(sourceFile)).trim()
        if (!extractedText) {
          throw new Error('No extractable text found in this file for summarization.')
        }

        const cappedText = extractedText.slice(0, 120000)
        const prompt = [
          'You are a document summarization assistant.',
          `User instruction: ${userPrompt}`,
          'Summarize the content below and return these sections:',
          '1. TL;DR (max 3 bullet points)',
          '2. Key points',
          '3. Action items',
          '4. Risks or open questions',
          '',
          cappedText,
        ].join('\n')

        const { summary } = await requestGeminiSummary({ apiKey, prompt })

        if (!summary) {
          throw new Error('Gemini returned an empty summary.')
        }

        setAiSummaryMessages((prev) => [
          ...prev,
          {
            id: `assistant-${Date.now()}`,
            role: 'assistant',
            text: summary,
          },
        ])
        return
      }

      if (activeTool?.slug === 'word-to-pdf') {
        const sourceFile = selectedFiles[0]
        const raw = await sourceFile.arrayBuffer()
        const bytes = new Uint8Array(raw)
        const isZipDocx = bytes.length >= 4
          && bytes[0] === 0x50
          && bytes[1] === 0x4b
          && (bytes[2] === 0x03 || bytes[2] === 0x05 || bytes[2] === 0x07)
          && (bytes[3] === 0x04 || bytes[3] === 0x06 || bytes[3] === 0x08)

        let extractedText = ''
        if (isZipDocx) {
          const { value: rawText } = await mammoth.extractRawText({ arrayBuffer: raw })
          extractedText = String(rawText || '')
        } else {
          // Fallback for HTML/text-based .doc files (or renamed files).
          let decoded = new TextDecoder('utf-8', { fatal: false }).decode(bytes)
          if (decoded.includes('\u0000')) {
            decoded = new TextDecoder('utf-16le', { fatal: false }).decode(bytes)
          }
          const container = document.createElement('div')
          container.innerHTML = decoded
          extractedText = (container.textContent || decoded || '')
        }

        const blocks = String(extractedText || '')
          .replace(/\r/g, '\n')
          .split(/\n{2,}/)
          .map((chunk) => chunk.replace(/\s+/g, ' ').trim())
          .filter(Boolean)

        if (!blocks.length) {
          blocks.push(`Converted from ${sourceFile.name}`)
        }

        // pdf-lib StandardFonts (WinAnsi) cannot encode many Unicode glyphs.
        // Replace unsupported symbols with close ASCII alternatives to avoid hard failures.
        const sanitizeForWinAnsi = (text) =>
          String(text || '')
            .replace(/[\u2018\u2019\u2032]/g, "'")
            .replace(/[\u201C\u201D\u2033]/g, '"')
            .replace(/[\u2013\u2014\u2212]/g, '-')
            .replace(/\u2026/g, '...')
            .replace(/\u00A0/g, ' ')
            .replace(/[^\x20-\x7E\xA0-\xFF]/g, '?')

        const toEncodableText = (text, pdfFont) => {
          const normalized = sanitizeForWinAnsi(text)
          let safe = ''
          for (const ch of Array.from(normalized)) {
            try {
              pdfFont.encodeText(ch)
              safe += ch
            } catch {
              safe += '?'
            }
          }
          return safe
        }

        const outputPdf = await PDFDocument.create()
        const font = await outputPdf.embedFont(StandardFonts.Helvetica)
        const pageWidth = 595.28
        const pageHeight = 841.89
        const margin = 48
        const maxWidth = pageWidth - margin * 2
        const fontSize = 12
        const lineHeight = 16

        const splitLines = (text) => {
          const safeText = toEncodableText(text, font)
          const words = safeText.split(/\s+/).filter(Boolean)
          const lines = []
          let current = ''
          for (const word of words) {
            const trial = current ? `${current} ${word}` : word
            const width = font.widthOfTextAtSize(trial, fontSize)
            if (width <= maxWidth || !current) {
              current = trial
            } else {
              lines.push(current)
              current = word
            }
          }
          if (current) lines.push(current)
          return lines.length ? lines : ['']
        }

        let page = outputPdf.addPage([pageWidth, pageHeight])
        let y = pageHeight - margin

        for (const block of blocks) {
          const lines = splitLines(block)
          for (const line of lines) {
            if (y <= margin) {
              page = outputPdf.addPage([pageWidth, pageHeight])
              y = pageHeight - margin
            }
            page.drawText(line, {
              x: margin,
              y,
              size: fontSize,
              font,
              color: rgb(0.08, 0.08, 0.08),
            })
            y -= lineHeight
          }
          y -= 6
        }

        const pdfBytes = await outputPdf.save()
        const outputName = `${sourceFile.name.replace(/\.(docx|doc)$/i, '') || 'converted'}-word-to-pdf.pdf`
        const outputFile = new File([pdfBytes], outputName, { type: 'application/pdf' })
        downloadBlob(outputFile, outputName)
        return
      }

      if (activeTool?.slug === 'powerpoint-to-pdf') {
        const sourceFile = selectedFiles[0]
        const raw = await sourceFile.arrayBuffer()
        const bytes = new Uint8Array(raw)
        const isZip = bytes.length >= 4
          && bytes[0] === 0x50
          && bytes[1] === 0x4b
          && (bytes[2] === 0x03 || bytes[2] === 0x05 || bytes[2] === 0x07)
          && (bytes[3] === 0x04 || bytes[3] === 0x06 || bytes[3] === 0x08)

        const decodeXmlEntities = (text) =>
          String(text || '')
            .replace(/&amp;/g, '&')
            .replace(/&lt;/g, '<')
            .replace(/&gt;/g, '>')
            .replace(/&quot;/g, '"')
            .replace(/&apos;/g, "'")
            .replace(/&#(\d+);/g, (_, num) => String.fromCharCode(Number(num)))
            .replace(/&#x([0-9a-fA-F]+);/g, (_, hex) => String.fromCharCode(parseInt(hex, 16)))

        const slides = []
        if (isZip) {
          const zip = await JSZip.loadAsync(raw)
          const slideEntries = Object.keys(zip.files)
            .filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name))
            .sort((a, b) => {
              const aNum = Number(a.match(/slide(\d+)\.xml$/)?.[1] || 0)
              const bNum = Number(b.match(/slide(\d+)\.xml$/)?.[1] || 0)
              return aNum - bNum
            })

          for (const entry of slideEntries) {
            const xml = await zip.file(entry).async('text')
            const paragraphMatches = xml.match(/<a:p[\s\S]*?<\/a:p>/g) || []
            const lines = paragraphMatches
              .map((paragraph) => {
                const textRuns = [...paragraph.matchAll(/<a:t[^>]*>([\s\S]*?)<\/a:t>/g)]
                  .map((m) => decodeXmlEntities(m[1]))
                return textRuns.join('').replace(/\s+/g, ' ').trim()
              })
              .filter(Boolean)
            slides.push(lines)
          }
        }

        if (!slides.length) {
          const fallbackText = new TextDecoder('utf-8', { fatal: false }).decode(bytes)
          const lines = fallbackText
            .replace(/\r/g, '\n')
            .split('\n')
            .map((line) => line.replace(/\s+/g, ' ').trim())
            .filter(Boolean)
          slides.push(lines.length ? lines : [`Converted from ${sourceFile.name}`])
        }

        const outputPdf = await PDFDocument.create()
        const font = await outputPdf.embedFont(StandardFonts.Helvetica)
        const pageWidth = 960
        const pageHeight = 540
        const margin = 36
        const maxWidth = pageWidth - margin * 2
        const headingSize = 18
        const bodySize = 14
        const lineHeight = 19

        const wrapLine = (text, size) => {
          const words = String(text || '').split(/\s+/).filter(Boolean)
          const lines = []
          let current = ''
          for (const word of words) {
            const trial = current ? `${current} ${word}` : word
            const width = font.widthOfTextAtSize(trial, size)
            if (width <= maxWidth || !current) current = trial
            else {
              lines.push(current)
              current = word
            }
          }
          if (current) lines.push(current)
          return lines.length ? lines : ['']
        }

        slides.forEach((slideLines, slideIdx) => {
          const page = outputPdf.addPage([pageWidth, pageHeight])
          let y = pageHeight - margin
          page.drawText(`Slide ${slideIdx + 1}`, {
            x: margin,
            y,
            size: headingSize,
            font,
            color: rgb(0.1, 0.1, 0.1),
          })
          y -= 28
          for (const line of slideLines) {
            const wrapped = wrapLine(line, bodySize)
            for (const segment of wrapped) {
              if (y <= margin) break
              page.drawText(segment, {
                x: margin,
                y,
                size: bodySize,
                font,
                color: rgb(0.13, 0.13, 0.13),
              })
              y -= lineHeight
            }
            if (y <= margin) break
          }
        })

        const pdfBytes = await outputPdf.save()
        const outputName = `${sourceFile.name.replace(/\.(pptx|ppt)$/i, '') || 'converted'}-powerpoint-to-pdf.pdf`
        downloadBlob(new File([pdfBytes], outputName, { type: 'application/pdf' }), outputName)
        return
      }

      if (activeTool?.slug === 'excel-to-pdf') {
        const sourceFile = selectedFiles[0]
        const raw = await sourceFile.arrayBuffer()
        const workbook = XLSX.read(raw, { type: 'array' })
        const sheetNames = workbook.SheetNames || []
        if (!sheetNames.length) {
          throw new Error('No worksheet found in this file.')
        }

        const outputPdf = await PDFDocument.create()
        const font = await outputPdf.embedFont(StandardFonts.Helvetica)
        const pageWidth = 842
        const pageHeight = 595
        const margin = 26
        const maxWidth = pageWidth - margin * 2
        const titleSize = 15
        const bodySize = 9
        const lineHeight = 12

        const wrap = (text) => {
          const words = String(text || '').split(/\s+/).filter(Boolean)
          const lines = []
          let current = ''
          for (const word of words) {
            const trial = current ? `${current} ${word}` : word
            const width = font.widthOfTextAtSize(trial, bodySize)
            if (width <= maxWidth || !current) current = trial
            else {
              lines.push(current)
              current = word
            }
          }
          if (current) lines.push(current)
          return lines.length ? lines : ['']
        }

        for (const sheetName of sheetNames) {
          const sheet = workbook.Sheets[sheetName]
          const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false, defval: '' })
          const page = outputPdf.addPage([pageWidth, pageHeight])
          let y = pageHeight - margin

          page.drawText(`Sheet: ${sheetName}`, {
            x: margin,
            y,
            size: titleSize,
            font,
            color: rgb(0.08, 0.08, 0.08),
          })
          y -= 22

          const printableRows = rows.length ? rows : [['(empty sheet)']]
          for (const row of printableRows) {
            const rowText = row
              .map((cell) => String(cell ?? '').replace(/\s+/g, ' ').trim())
              .join(' | ')
              .trim() || '(empty row)'
            const wrapped = wrap(rowText)
            for (const line of wrapped) {
              if (y <= margin) break
              page.drawText(line, {
                x: margin,
                y,
                size: bodySize,
                font,
                color: rgb(0.12, 0.12, 0.12),
              })
              y -= lineHeight
            }
            if (y <= margin) break
          }
        }

        const pdfBytes = await outputPdf.save()
        const outputName = `${sourceFile.name.replace(/\.(xlsx|xls|csv)$/i, '') || 'converted'}-excel-to-pdf.pdf`
        downloadBlob(new File([pdfBytes], outputName, { type: 'application/pdf' }), outputName)
        return
      }

      if (activeTool?.slug === 'jpg-to-pdf') {
        const outputPdf = await PDFDocument.create()
        const margin = 24

        for (const imageFile of selectedFiles) {
          const raw = await imageFile.arrayBuffer()
          const lowerName = imageFile.name?.toLowerCase() || ''
          const isPng = imageFile.type === 'image/png' || lowerName.endsWith('.png')
          const embeddedImage = isPng
            ? await outputPdf.embedPng(raw)
            : await outputPdf.embedJpg(raw)
          const imgW = embeddedImage.width
          const imgH = embeddedImage.height
          const pageWidth = Math.max(595.28, imgW + margin * 2)
          const pageHeight = Math.max(841.89, imgH + margin * 2)
          const page = outputPdf.addPage([pageWidth, pageHeight])
          const maxW = pageWidth - margin * 2
          const maxH = pageHeight - margin * 2
          const scale = Math.min(maxW / imgW, maxH / imgH)
          const drawW = imgW * scale
          const drawH = imgH * scale
          const x = (pageWidth - drawW) / 2
          const y = (pageHeight - drawH) / 2
          page.drawImage(embeddedImage, { x, y, width: drawW, height: drawH })
        }

        const pdfBytes = await outputPdf.save()
        const outputName = `jpg-to-pdf-${Date.now()}.pdf`
        downloadBlob(new File([pdfBytes], outputName, { type: 'application/pdf' }), outputName)
        return
      }

      if (activeTool?.slug === 'translate-pdf') {
        if (!translateTargetOptions.length) {
          throw new Error('No installed Argos target languages available for this source language.')
        }
        if (translateSourceLang !== 'auto' && translateSourceLang === translateTargetLang) {
          throw new Error('Source and target languages must be different.')
        }
        const sourceFile = selectedFiles[0]
        const extractedText = (await extractTextForSummarizer(sourceFile)).trim()
        if (!extractedText) {
          throw new Error('No extractable text found in this file for translation.')
        }
        setTranslatedText('')
        setTranslateStatus('Translating with Argos...')
        const translated = await translateWithArgos({
          text: extractedText.slice(0, 180000),
          source: translateSourceLang,
          target: translateTargetLang,
        })
        setTranslatedText(translated)
        setTranslateStatus('Translation ready.')
        return
      }

      throw new Error(`Conversion not implemented yet for ${activeTool?.title || 'this tool'}.`)
    } catch (error) {
      if (activeTool?.slug === 'pdf-to-word') {
        window.alert(error?.message || 'Could not convert PDF to editable Word format.')
      } else if (activeTool?.slug === 'pdf-to-powerpoint') {
        window.alert(error?.message || 'Could not convert PDF to editable PowerPoint format.')
      } else if (activeTool?.slug === 'pdf-to-excel') {
        window.alert(error?.message || 'Could not convert PDF to editable Excel format.')
      } else if (activeTool?.slug === 'pdf-to-jpg') {
        window.alert(error?.message || 'Could not convert PDF to JPG.')
      } else if (activeTool?.slug === 'word-to-pdf') {
        window.alert(error?.message || 'Could not convert Word to PDF.')
      } else if (activeTool?.slug === 'powerpoint-to-pdf') {
        window.alert(error?.message || 'Could not convert PowerPoint to PDF.')
      } else if (activeTool?.slug === 'excel-to-pdf') {
        window.alert(error?.message || 'Could not convert Excel to PDF.')
      } else if (activeTool?.slug === 'jpg-to-pdf') {
        window.alert(error?.message || 'Could not convert JPG to PDF.')
      } else if (activeTool?.slug === 'ai-summarizer') {
        setAiSummaryMessages((prev) => [
          ...prev,
          { id: `assistant-${Date.now()}`, role: 'assistant', text: `Error: ${error?.message || 'Could not generate summary.'}` },
        ])
        window.alert(error?.message || 'Could not generate AI summary.')
      } else if (activeTool?.slug === 'translate-pdf') {
        setTranslateStatus('')
        window.alert(error?.message || 'Could not translate file.')
      } else {
        window.alert(error?.message || 'Could not complete conversion.')
      }
    } finally {
      setIsMerging(false)
    }
  }

  const mergeSuggestions = [
    { label: 'Compress PDF', path: '/tool/compress-pdf' },
    { label: 'Split PDF', path: '/tool/split-pdf' },
    { label: 'Add page numbers', path: '/tool/page-numbers' },
    { label: 'Add watermark', path: '/tool/watermark-pdf' },
    { label: 'Rotate PDF', path: '/tool/rotate-pdf' },
    { label: 'Protect PDF', path: '/tool/protect-pdf' },
  ]

  const pageNumPositions = [
    ['tl', 'tc', 'tr'],
    ['cl', 'cc', 'cr'],
    ['bl', 'bc', 'br'],
  ]

  const clampPage = (value, totalPages) => {
    if (!totalPages) return 1
    return Math.min(Math.max(value, 1), totalPages)
  }

  const handleSplit = async () => {
    if (!selectedFiles.length) return

    try {
      setSplitBusy(true)
      const sourceFile = selectedFiles[0]
      const data = await sourceFile.arrayBuffer()
      const sourcePdf = await PDFDocument.load(data)
      const total = sourcePdf.getPageCount()
      const from = clampPage(splitFromPage, total)
      const to = clampPage(splitToPage, total)

      let indices = []
      if (splitMode === 'range') {
        const start = Math.min(from, to) - 1
        const end = Math.max(from, to) - 1
        for (let page = start; page <= end; page += 1) {
          indices.push(page)
        }
      } else if (splitMode === 'pages') {
        indices = Array.from(new Set([from - 1, to - 1])).sort((a, b) => a - b)
      } else {
        const chunkSize = Math.max(1, from)
        const chunkNumber = Math.max(1, to)
        const start = (chunkNumber - 1) * chunkSize
        const end = Math.min(start + chunkSize, total) - 1
        if (start >= total) {
          indices = [total - 1]
        } else {
          for (let page = start; page <= end; page += 1) {
            indices.push(page)
          }
        }
      }

      const out = await PDFDocument.create()
      const pages = await out.copyPages(sourcePdf, indices)
      pages.forEach((page) => out.addPage(page))
      const bytes = await out.save()
      const blob = new Blob([bytes], { type: 'application/pdf' })
      const url = URL.createObjectURL(blob)
      if (splitResultUrl) {
        URL.revokeObjectURL(splitResultUrl)
      }
      setSplitResultUrl(url)
      setSplitResultName(`split-${sourceFile.name || Date.now()}.pdf`)
    } finally {
      setSplitBusy(false)
    }
  }

  const resetSplitResult = () => {
    if (splitResultUrl) {
      URL.revokeObjectURL(splitResultUrl)
    }
    setSplitResultUrl('')
    setSplitResultName('')
  }

  const downloadSourceFile = () => {
    if (!selectedFiles.length) return
    const source = selectedFiles[0]
    const url = URL.createObjectURL(source)
    const link = document.createElement('a')
    link.href = url
    link.download = source.name || 'source.pdf'
    document.body.appendChild(link)
    link.click()
    document.body.removeChild(link)
    window.setTimeout(() => URL.revokeObjectURL(url), 60_000)
  }

  const copyPageLink = async (pageIdx) => {
    if (!previewUrls[0]) return
    const value = `${previewUrls[0]}#page=${pageIdx + 1}`
    try {
      await navigator.clipboard.writeText(value)
    } catch {
      // no-op fallback for unsupported clipboard contexts
    }
  }

  const selectedEditFile = selectedFiles[editSelectedIndex] || null
  const selectedEditPreviewUrl = previewUrls[editSelectedIndex] || ''
  const currentPageObjects = editObjectsByPage[editPage] || []
  const selectedPageNumFile = selectedFiles[pageNumFileIndex] || null
  const selectedPageNumPreviewUrl = previewUrls[pageNumFileIndex] || ''
  const compareLeftFile = selectedFiles[compareLeftFileIndex] || null
  const compareRightFile = selectedFiles[compareRightFileIndex] || null
  const compareLeftPreviewUrl = previewUrls[compareLeftFileIndex] || ''
  const compareRightPreviewUrl = previewUrls[compareRightFileIndex] || ''
  const selectedCropFile = selectedFiles[0] || null
  const getRotateFileKey = (file, index) => `${file.name}-${file.size}-${file.lastModified}-${index}`

  useEffect(() => {
    setRotateAnglesByFile((prev) => {
      const next = {}
      selectedFiles.forEach((file, index) => {
        const key = getRotateFileKey(file, index)
        next[key] = prev[key] || 0
      })
      return next
    })
  }, [selectedFiles])

  useEffect(() => {
    if (rotateSelectedIndex >= selectedFiles.length) {
      setRotateSelectedIndex(0)
    }
  }, [rotateSelectedIndex, selectedFiles.length])

  useEffect(() => {
    if (pageNumFileIndex >= selectedFiles.length) {
      setPageNumFileIndex(0)
    }
  }, [pageNumFileIndex, selectedFiles.length])

  useEffect(() => {
    const len = selectedFiles.length
    if (len === 0) {
      setCompareLeftFileIndex(0)
      setCompareRightFileIndex(1)
      return
    }
    if (compareLeftFileIndex >= len) {
      setCompareLeftFileIndex(0)
    }
    if (compareRightFileIndex >= len) {
      setCompareRightFileIndex(len > 1 ? 1 : 0)
    }
    if (len > 1 && compareLeftFileIndex === compareRightFileIndex) {
      setCompareRightFileIndex(compareLeftFileIndex === 0 ? 1 : 0)
    }
  }, [compareLeftFileIndex, compareRightFileIndex, selectedFiles.length])

  const setCompareFileForSide = (side, nextIndexRaw) => {
    const nextIndex = Number(nextIndexRaw)
    if (!Number.isFinite(nextIndex) || nextIndex < 0 || nextIndex >= selectedFiles.length) return
    if (side === 'left') {
      setCompareLeftFileIndex(nextIndex)
      if (selectedFiles.length > 1 && nextIndex === compareRightFileIndex) {
        const fallback = selectedFiles.findIndex((_, idx) => idx !== nextIndex)
        if (fallback >= 0) setCompareRightFileIndex(fallback)
      }
      setCompareLeftPage(1)
      return
    }
    setCompareRightFileIndex(nextIndex)
    if (selectedFiles.length > 1 && nextIndex === compareLeftFileIndex) {
      const fallback = selectedFiles.findIndex((_, idx) => idx !== nextIndex)
      if (fallback >= 0) setCompareLeftFileIndex(fallback)
    }
    setCompareRightPage(1)
  }

  useEffect(() => {
    if (activeTool?.slug !== 'compare-pdf' || selectedFiles.length < 2 || !compareLeftFile || !compareRightFile) {
      if (activeTool?.slug === 'compare-pdf') {
        setCompareStatus(selectedFiles.length === 1 ? 'Add one more PDF to compare.' : '')
      }
      setCompareLeftProxy(null)
      setCompareRightProxy(null)
      setCompareLeftPages(1)
      setCompareRightPages(1)
      return
    }

    let cancelled = false
    ;(async () => {
      try {
        setCompareBusy(true)
        setCompareStatus('Loading PDF files for comparison...')
        const [leftRaw, rightRaw] = await Promise.all([
          compareLeftFile.arrayBuffer(),
          compareRightFile.arrayBuffer(),
        ])
        const [leftTask, rightTask] = [
          pdfjsLib.getDocument({ data: leftRaw }),
          pdfjsLib.getDocument({ data: rightRaw }),
        ]
        const [leftProxy, rightProxy] = await Promise.all([leftTask.promise, rightTask.promise])
        if (cancelled) return
        setCompareLeftProxy(leftProxy)
        setCompareRightProxy(rightProxy)
        setCompareLeftPages(leftProxy.numPages)
        setCompareRightPages(rightProxy.numPages)
        setCompareLeftPage((prev) => Math.min(Math.max(prev, 1), leftProxy.numPages))
        setCompareRightPage((prev) => Math.min(Math.max(prev, 1), rightProxy.numPages))
      } catch (error) {
        if (!cancelled) {
          setCompareStatus(`Could not load files for compare: ${error?.message || 'unknown error'}`)
          setCompareLeftProxy(null)
          setCompareRightProxy(null)
        }
      } finally {
        if (!cancelled) setCompareBusy(false)
      }
    })()

    return () => {
      cancelled = true
    }
  }, [activeTool?.slug, selectedFiles, compareLeftFile, compareRightFile])

  useEffect(() => {
    if (!compareScrollSync) return
    const leftNode = compareLeftScrollRef.current
    const rightNode = compareRightScrollRef.current
    if (!leftNode || !rightNode) return

    let lock = false
    const onLeft = () => {
      if (lock) return
      lock = true
      rightNode.scrollTop = leftNode.scrollTop
      rightNode.scrollLeft = leftNode.scrollLeft
      lock = false
    }
    const onRight = () => {
      if (lock) return
      lock = true
      leftNode.scrollTop = rightNode.scrollTop
      leftNode.scrollLeft = rightNode.scrollLeft
      lock = false
    }
    leftNode.addEventListener('scroll', onLeft)
    rightNode.addEventListener('scroll', onRight)
    return () => {
      leftNode.removeEventListener('scroll', onLeft)
      rightNode.removeEventListener('scroll', onRight)
    }
  }, [compareScrollSync, compareMode, compareLeftPage, compareRightPage, compareZoom])

  const renderComparePage = async (proxy, pageNumber, canvas, scale = 1) => {
    if (!proxy || !canvas) return null
    const page = await proxy.getPage(pageNumber)
    const viewport = page.getViewport({ scale })
    canvas.width = Math.max(1, Math.floor(viewport.width))
    canvas.height = Math.max(1, Math.floor(viewport.height))
    const ctx = canvas.getContext('2d')
    if (!ctx) return { width: canvas.width, height: canvas.height }
    await page.render({ canvasContext: ctx, viewport }).promise
    return { width: canvas.width, height: canvas.height }
  }

  const getPageTextLines = async (proxy, pageNumber) => {
    if (!proxy) return []
    const page = await proxy.getPage(pageNumber)
    const content = await page.getTextContent()
    const text = content.items
      .map((item) => (typeof item.str === 'string' ? item.str.trim() : ''))
      .filter(Boolean)
      .join('\n')
    return text.split('\n').map((line) => line.trim()).filter(Boolean)
  }

  const buildCompareReport = (leftLines, rightLines) => {
    const maxLen = Math.max(leftLines.length, rightLines.length)
    const changes = []
    for (let idx = 0; idx < maxLen; idx += 1) {
      const leftLine = leftLines[idx] || ''
      const rightLine = rightLines[idx] || ''
      if (leftLine !== rightLine) {
        changes.push({
          id: `${idx}-${leftLine}-${rightLine}`,
          line: idx + 1,
          left: leftLine,
          right: rightLine,
          type: leftLine && rightLine ? 'changed' : leftLine ? 'removed' : 'added',
        })
      }
    }
    return changes
  }

  useEffect(() => {
    if (activeTool?.slug !== 'compare-pdf') return
    if (!compareLeftProxy || !compareRightProxy) return

    let cancelled = false
    ;(async () => {
      try {
        const scale = Math.max(0.5, Math.min(2.2, compareZoom / 100))
        if (compareMode === 'semantic') {
          await Promise.all([
            renderComparePage(compareLeftProxy, compareLeftPage, compareLeftCanvasRef.current, scale),
            renderComparePage(compareRightProxy, compareRightPage, compareRightCanvasRef.current, scale),
          ])
        } else {
          const leftCanvas = document.createElement('canvas')
          const rightCanvas = document.createElement('canvas')
          const leftMeta = await renderComparePage(compareLeftProxy, compareLeftPage, leftCanvas, scale)
          const rightMeta = await renderComparePage(compareRightProxy, compareRightPage, rightCanvas, scale)
          if (!leftMeta || !rightMeta) return

          const width = Math.max(leftMeta.width, rightMeta.width)
          const height = Math.max(leftMeta.height, rightMeta.height)
          const target = compareOverlayCanvasRef.current
          if (!target) return
          target.width = width
          target.height = height
          const ctx = target.getContext('2d')
          if (!ctx) return

          ctx.fillStyle = '#f0f0f3'
          ctx.fillRect(0, 0, width, height)
          ctx.drawImage(leftCanvas, 0, 0)

          const leftCtx = leftCanvas.getContext('2d')
          const rightCtx = rightCanvas.getContext('2d')
          if (!leftCtx || !rightCtx) return
          const leftData = leftCtx.getImageData(0, 0, Math.min(width, leftCanvas.width), Math.min(height, leftCanvas.height))
          const rightData = rightCtx.getImageData(0, 0, Math.min(width, rightCanvas.width), Math.min(height, rightCanvas.height))
          const merged = ctx.getImageData(0, 0, width, height)
          const maxPx = Math.min(leftData.data.length, rightData.data.length, merged.data.length)
          for (let i = 0; i < maxPx; i += 4) {
            const dr = Math.abs(leftData.data[i] - rightData.data[i])
            const dg = Math.abs(leftData.data[i + 1] - rightData.data[i + 1])
            const db = Math.abs(leftData.data[i + 2] - rightData.data[i + 2])
            const diff = dr + dg + db
            if (diff > 85) {
              merged.data[i] = 255
              merged.data[i + 1] = 80
              merged.data[i + 2] = 80
              merged.data[i + 3] = 220
            }
          }
          ctx.putImageData(merged, 0, 0)
        }

        const [leftLines, rightLines] = await Promise.all([
          getPageTextLines(compareLeftProxy, compareLeftPage),
          getPageTextLines(compareRightProxy, compareRightPage),
        ])
        if (cancelled) return
        const report = buildCompareReport(leftLines, rightLines)
        setCompareReport(report)
        setCompareStatus(report.length ? 'Changes detected in selected pages.' : 'No text changes found in selected pages.')
      } catch (error) {
        if (!cancelled) setCompareStatus(`Compare failed: ${error?.message || 'unknown error'}`)
      }
    })()
    return () => {
      cancelled = true
    }
  }, [
    activeTool?.slug,
    compareLeftProxy,
    compareRightProxy,
    compareLeftPage,
    compareRightPage,
    compareMode,
    compareZoom,
  ])

  useEffect(() => {
    if (activeTool?.slug !== 'crop-pdf' || !selectedCropFile) {
      setCropPdfProxy(null)
      setCropPageCount(1)
      return
    }
    let cancelled = false
    ;(async () => {
      try {
        const raw = await selectedCropFile.arrayBuffer()
        const pdf = await pdfjsLib.getDocument({ data: raw }).promise
        if (cancelled) return
        setCropPdfProxy(pdf)
        setCropPageCount(pdf.numPages)
        setCropPage(1)
        setCropRect(null)
        setCropStatus('')
      } catch (error) {
        if (!cancelled) {
          setCropStatus(`Could not load PDF: ${error?.message || 'unknown error'}`)
          setCropPdfProxy(null)
          setCropPageCount(1)
        }
      }
    })()
    return () => {
      cancelled = true
    }
  }, [activeTool?.slug, selectedCropFile])

  useEffect(() => {
    if (cropPage > cropPageCount) {
      setCropPage(Math.max(1, cropPageCount))
    }
  }, [cropPage, cropPageCount])

  useEffect(() => {
    if (activeTool?.slug !== 'crop-pdf') return
    if (!cropPdfProxy || !cropCanvasRef.current) return
    let cancelled = false
    ;(async () => {
      const page = await cropPdfProxy.getPage(cropPage)
      if (cancelled) return
      const scale = Math.max(0.5, Math.min(2, cropZoom / 100))
      const viewport = page.getViewport({ scale })
      const canvas = cropCanvasRef.current
      canvas.width = Math.max(1, Math.floor(viewport.width))
      canvas.height = Math.max(1, Math.floor(viewport.height))
      setCropCanvasSize({ width: canvas.width, height: canvas.height })
      const ctx = canvas.getContext('2d')
      await page.render({ canvasContext: ctx, viewport }).promise
      if (!cancelled && !cropRect) {
        const marginX = canvas.width * 0.12
        const marginY = canvas.height * 0.12
        setCropRect({
          x: marginX,
          y: marginY,
          w: canvas.width - marginX * 2,
          h: canvas.height - marginY * 2,
        })
      }
    })()
    return () => {
      cancelled = true
    }
  }, [activeTool?.slug, cropPdfProxy, cropPage, cropZoom])

  const clampCropRect = (rect, canvasWidth, canvasHeight) => {
    const minW = 24
    const minH = 24
    const next = { ...rect }
    next.w = Math.max(minW, Math.min(next.w, canvasWidth))
    next.h = Math.max(minH, Math.min(next.h, canvasHeight))
    next.x = Math.max(0, Math.min(next.x, canvasWidth - next.w))
    next.y = Math.max(0, Math.min(next.y, canvasHeight - next.h))
    return next
  }

  const getCropPointer = (event) => {
    const canvas = cropCanvasRef.current
    if (!canvas) return { x: 0, y: 0 }
    const rect = canvas.getBoundingClientRect()
    const scaleX = canvas.width / rect.width
    const scaleY = canvas.height / rect.height
    return {
      x: (event.clientX - rect.left) * scaleX,
      y: (event.clientY - rect.top) * scaleY,
    }
  }

  const getCropDragMode = (point, rect) => {
    if (!rect) return 'new'
    const handleSize = 12
    const corners = [
      { key: 'resize-nw', x: rect.x, y: rect.y },
      { key: 'resize-ne', x: rect.x + rect.w, y: rect.y },
      { key: 'resize-sw', x: rect.x, y: rect.y + rect.h },
      { key: 'resize-se', x: rect.x + rect.w, y: rect.y + rect.h },
    ]
    const hitCorner = corners.find(
      (corner) => Math.abs(point.x - corner.x) <= handleSize && Math.abs(point.y - corner.y) <= handleSize,
    )
    if (hitCorner) return hitCorner.key
    const inRect =
      point.x >= rect.x &&
      point.x <= rect.x + rect.w &&
      point.y >= rect.y &&
      point.y <= rect.y + rect.h
    if (inRect) return 'move'
    return 'new'
  }

  const onCropMouseDown = (event) => {
    const point = getCropPointer(event)
    const mode = getCropDragMode(point, cropRect)
    const baseRect = cropRect || { x: point.x, y: point.y, w: 1, h: 1 }
    cropDragRef.current = { mode, start: point, baseRect }
    if (mode === 'new') {
      setCropRect({ x: point.x, y: point.y, w: 1, h: 1 })
    }
  }

  const onCropMouseMove = (event) => {
    if (!cropDragRef.current) return
    const drag = cropDragRef.current
    const point = getCropPointer(event)
    const dx = point.x - drag.start.x
    const dy = point.y - drag.start.y
    const canvas = cropCanvasRef.current
    if (!canvas) return
    const { width, height } = canvas
    let next = { ...drag.baseRect }

    if (drag.mode === 'new') {
      next = {
        x: Math.min(drag.start.x, point.x),
        y: Math.min(drag.start.y, point.y),
        w: Math.abs(point.x - drag.start.x),
        h: Math.abs(point.y - drag.start.y),
      }
    } else if (drag.mode === 'move') {
      next.x = drag.baseRect.x + dx
      next.y = drag.baseRect.y + dy
    } else if (drag.mode === 'resize-nw') {
      next.x = drag.baseRect.x + dx
      next.y = drag.baseRect.y + dy
      next.w = drag.baseRect.w - dx
      next.h = drag.baseRect.h - dy
    } else if (drag.mode === 'resize-ne') {
      next.y = drag.baseRect.y + dy
      next.w = drag.baseRect.w + dx
      next.h = drag.baseRect.h - dy
    } else if (drag.mode === 'resize-sw') {
      next.x = drag.baseRect.x + dx
      next.w = drag.baseRect.w - dx
      next.h = drag.baseRect.h + dy
    } else if (drag.mode === 'resize-se') {
      next.w = drag.baseRect.w + dx
      next.h = drag.baseRect.h + dy
    }

    setCropRect(clampCropRect(next, width, height))
  }

  const onCropMouseUp = () => {
    cropDragRef.current = null
  }

  const resetCrop = () => {
    const canvas = cropCanvasRef.current
    if (!canvas) {
      setCropRect(null)
      return
    }
    const marginX = canvas.width * 0.12
    const marginY = canvas.height * 0.12
    setCropRect({
      x: marginX,
      y: marginY,
      w: canvas.width - marginX * 2,
      h: canvas.height - marginY * 2,
    })
    setCropStatus('Crop area reset.')
  }

  const handleCropDownload = async () => {
    if (!selectedCropFile || !cropRect || !cropCanvasRef.current) return
    try {
      setCropBusy(true)
      setCropStatus('Cropping PDF...')
      const raw = await selectedCropFile.arrayBuffer()
      const pdf = await PDFDocument.load(raw, { ignoreEncryption: true })
      const pages = pdf.getPages()
      const targetIndexes =
        cropScope === 'current'
          ? [Math.max(0, Math.min(cropPage - 1, pages.length - 1))]
          : pages.map((_, idx) => idx)

      const ratio = {
        x: cropRect.x / cropCanvasSize.width,
        y: cropRect.y / cropCanvasSize.height,
        w: cropRect.w / cropCanvasSize.width,
        h: cropRect.h / cropCanvasSize.height,
      }

      targetIndexes.forEach((idx) => {
        const page = pages[idx]
        const { width, height } = page.getSize()
        const x = width * ratio.x
        const w = width * ratio.w
        const h = height * ratio.h
        const yFromTop = height * ratio.y
        const y = height - yFromTop - h
        page.setCropBox(x, y, w, h)
        page.setMediaBox(x, y, w, h)
      })

      const bytes = await pdf.save({ useObjectStreams: false })
      const blob = new Blob([bytes], { type: 'application/pdf' })
      const filename = `cropped-${Date.now()}-${selectedCropFile.name || 'document.pdf'}`
      downloadBlob(blob, filename)
      setCropStatus('Cropped PDF downloaded.')
    } catch (error) {
      setCropStatus(`Could not crop PDF: ${error?.message || 'unknown error'}`)
    } finally {
      setCropBusy(false)
    }
  }

  const hexToRgbTuple = (hex) => {
    const clean = hex.replace('#', '').trim()
    if (clean.length !== 6) return [1, 1, 1]
    const r = parseInt(clean.slice(0, 2), 16) / 255
    const g = parseInt(clean.slice(2, 4), 16) / 255
    const b = parseInt(clean.slice(4, 6), 16) / 255
    return [r, g, b]
  }

  const getPageNumFontName = () => {
    const isB = pageNumBold
    const isI = pageNumItalic
    if (pageNumFontFamily === 'times') {
      if (isB && isI) return StandardFonts.TimesRomanBoldItalic
      if (isB) return StandardFonts.TimesRomanBold
      if (isI) return StandardFonts.TimesRomanItalic
      return StandardFonts.TimesRoman
    }
    if (pageNumFontFamily === 'courier') {
      if (isB && isI) return StandardFonts.CourierBoldOblique
      if (isB) return StandardFonts.CourierBold
      if (isI) return StandardFonts.CourierOblique
      return StandardFonts.Courier
    }
    if (isB && isI) return StandardFonts.HelveticaBoldOblique
    if (isB) return StandardFonts.HelveticaBold
    if (isI) return StandardFonts.HelveticaOblique
    return StandardFonts.Helvetica
  }

  const getMirroredPosition = (pos) => {
    if (pos.endsWith('l')) return `${pos[0]}r`
    if (pos.endsWith('r')) return `${pos[0]}l`
    return pos
  }

  const getPageNumCoords = (page, textWidth, textHeight, marginPx, basePos) => {
    const { width, height } = page.getSize()
    let x = marginPx
    let y = marginPx

    if (basePos.endsWith('c')) x = (width - textWidth) / 2
    if (basePos.endsWith('r')) x = width - textWidth - marginPx
    if (basePos.startsWith('c')) y = (height - textHeight) / 2
    if (basePos.startsWith('t')) y = height - textHeight - marginPx
    if (basePos.startsWith('b')) y = marginPx

    return { x, y }
  }

  const handleAddPageNumbers = async () => {
    const targetFile = selectedFiles[pageNumFileIndex]
    if (!targetFile) return
    try {
      setPageNumBusy(true)
      setPageNumStatus('Adding page numbers...')
      const raw = await targetFile.arrayBuffer()
      const pdf = await PDFDocument.load(raw, { ignoreEncryption: true })
      const font = await pdf.embedFont(getPageNumFontName())
      const pages = pdf.getPages()
      const marginPx = pageNumMargin === 'none' ? 10 : pageNumMargin === 'large' ? 40 : 24
      const [r, g, b] = hexToRgbTuple('#000000')
      const size = 16

      pages.forEach((page, idx) => {
        const displayNumber = pageNumFirst + idx
        const text =
          pageNumTextMode === 'full'
            ? `Page ${displayNumber} of ${pages.length}`
            : `${displayNumber}`
        const textWidth = font.widthOfTextAtSize(text, size)
        const textHeight = size
        const position =
          pageNumMode === 'facing' && (idx + 1) % 2 === 0
            ? getMirroredPosition(pageNumPosition)
            : pageNumPosition
        const { x, y } = getPageNumCoords(page, textWidth, textHeight, marginPx, position)
        page.drawText(text, {
          x,
          y,
          size,
          font,
          color: rgb(r, g, b),
        })
        if (pageNumUnderline) {
          page.drawLine({
            start: { x, y: y - 2 },
            end: { x: x + textWidth, y: y - 2 },
            thickness: 1.2,
            color: rgb(r, g, b),
          })
        }
      })

      const bytes = await pdf.save({ useObjectStreams: false })
      const blob = new Blob([bytes], { type: 'application/pdf' })
      const filename = `numbered-${Date.now()}-${targetFile.name || 'document.pdf'}`
      downloadBlob(blob, filename)
      setPageNumStatus('Page numbers burned into PDF and downloaded.')
    } catch (error) {
      setPageNumStatus(`Could not add page numbers: ${error?.message || 'unknown error'}`)
    } finally {
      setPageNumBusy(false)
    }
  }

  const filteredCompareReport = useMemo(() => {
    const needle = compareSearch.trim().toLowerCase()
    if (!needle) return compareReport
    return compareReport.filter((item) =>
      `${item.left} ${item.right} ${item.type}`.toLowerCase().includes(needle),
    )
  }, [compareReport, compareSearch])

  const downloadCompareReport = () => {
    const lines = filteredCompareReport.map(
      (item) => `Line ${item.line} | ${item.type.toUpperCase()} | Left: ${item.left || '-'} | Right: ${item.right || '-'}`,
    )
    const body = [
      'TRULYPDF Compare Report',
      `Generated: ${new Date().toLocaleString()}`,
      `Mode: ${compareMode}`,
      `Left file: ${compareLeftFile?.name || '-'}`,
      `Right file: ${compareRightFile?.name || '-'}`,
      `Total changes: ${filteredCompareReport.length}`,
      '',
      ...lines,
    ].join('\n')
    const blob = new Blob([body], { type: 'text/plain;charset=utf-8' })
    downloadBlob(blob, `compare-report-${Date.now()}.txt`)
  }

  const rotateFile = (index, delta) => {
    const file = selectedFiles[index]
    if (!file) return
    const key = getRotateFileKey(file, index)
    setRotateAnglesByFile((prev) => {
      const current = prev[key] || 0
      const next = ((current + delta) % 360 + 360) % 360
      return { ...prev, [key]: next }
    })
  }

  const resetRotateAll = () => {
    const next = {}
    selectedFiles.forEach((file, index) => {
      next[getRotateFileKey(file, index)] = 0
    })
    setRotateAnglesByFile(next)
    setRotateStatus('Rotation reset for all files.')
  }

  const handleRotateConvert = async () => {
    if (!selectedFiles.length) return
    try {
      setRotateBusy(true)
      setRotateStatus('Rotating PDF...')
      for (let fileIndex = 0; fileIndex < selectedFiles.length; fileIndex += 1) {
        const file = selectedFiles[fileIndex]
        const key = getRotateFileKey(file, fileIndex)
        const angle = rotateAnglesByFile[key] || 0
        const raw = await file.arrayBuffer()
        const pdf = await PDFDocument.load(raw, { ignoreEncryption: true })
        if (angle !== 0) {
          pdf.getPages().forEach((page) => {
            const current = page.getRotation()?.angle || 0
            page.setRotation(degrees((current + angle) % 360))
          })
        }
        const bytes = await pdf.save({ useObjectStreams: false })
        const blob = new Blob([bytes], { type: 'application/pdf' })
        const filename = `rotated-${Date.now()}-${file.name || 'document.pdf'}`
        downloadBlob(blob, filename)
      }
      setRotateStatus('Rotated PDF downloaded.')
    } catch (error) {
      setRotateStatus(`Could not rotate PDF: ${error?.message || 'unknown error'}`)
    } finally {
      setRotateBusy(false)
    }
  }

  const organizeAddBlank = (afterIndex = organizePages.length - 1) => {
    setOrganizePages((prev) => {
      const next = [...prev]
      const insertAt = Math.min(Math.max(afterIndex + 1, 0), next.length)
      next.splice(insertAt, 0, {
        id: crypto.randomUUID(),
        type: 'blank',
        rotation: 0,
        label: 'blank',
      })
      return next
    })
    setOrganizeStatus('Blank page added.')
  }

  const organizeRotate = (pageId, delta) => {
    setOrganizePages((prev) =>
      prev.map((page) =>
        page.id === pageId
          ? { ...page, rotation: ((page.rotation + delta) % 360 + 360) % 360 }
          : page,
      ),
    )
  }

  const organizeRemove = (pageId) => {
    setOrganizePages((prev) => {
      if (prev.length <= 1) return prev
      return prev.filter((page) => page.id !== pageId)
    })
  }

  const organizeResetAll = async () => {
    if (!selectedFiles.length) return
    try {
      const data = await selectedFiles[0].arrayBuffer()
      const pdf = await PDFDocument.load(data, { ignoreEncryption: true })
      const pages = pdf.getPages()
      setOrganizePages(
        pages.map((_, idx) => ({
          id: crypto.randomUUID(),
          type: 'source',
          sourcePageIndex: idx,
          rotation: 0,
          label: String(idx + 1),
        })),
      )
      setOrganizeStatus('Pages reset.')
      setOrganizeSortDirection('asc')
    } catch {
      setOrganizeStatus('Could not reset pages.')
    }
  }

  const onOrganizeDragStart = (event, index) => {
    setOrganizeDragIndex(index)
    event.dataTransfer.effectAllowed = 'move'
    event.dataTransfer.setData('text/plain', String(index))
  }

  const onOrganizeDrop = (event, index) => {
    event.preventDefault()
    const from = Number(event.dataTransfer.getData('text/plain'))
    if (Number.isNaN(from) || from === index) {
      setOrganizeDragIndex(null)
      return
    }
    setOrganizePages((prev) => {
      const next = [...prev]
      const [moved] = next.splice(from, 1)
      next.splice(index, 0, moved)
      return next
    })
    setOrganizeStatus('Page order updated.')
    setOrganizeDragIndex(null)
  }

  const onOrganizeGridDrop = (event) => {
    event.preventDefault()
    const from = Number(event.dataTransfer.getData('text/plain'))
    if (Number.isNaN(from)) {
      setOrganizeDragIndex(null)
      return
    }
    setOrganizePages((prev) => {
      if (from < 0 || from >= prev.length) return prev
      const next = [...prev]
      const [moved] = next.splice(from, 1)
      next.push(moved)
      return next
    })
    setOrganizeStatus('Page moved.')
    setOrganizeDragIndex(null)
  }

  const toggleOrganizeOrder = () => {
    setOrganizePages((prev) => [...prev].reverse())
    setOrganizeSortDirection((prev) => (prev === 'asc' ? 'desc' : 'asc'))
    setOrganizeStatus('Document arrangement changed.')
  }

  const handleOrganizeDownload = async () => {
    if (!selectedFiles.length || !organizePages.length) return
    try {
      setOrganizeBusy(true)
      setOrganizeStatus('Organizing PDF...')
      const srcData = await selectedFiles[0].arrayBuffer()
      const srcPdf = await PDFDocument.load(srcData, { ignoreEncryption: true })
      const out = await PDFDocument.create()

      for (const page of organizePages) {
        if (page.type === 'blank') {
          const newPage = out.addPage([organizeBaseSize.width, organizeBaseSize.height])
          if (page.rotation) {
            newPage.setRotation(degrees(page.rotation))
          }
          continue
        }
        const [copied] = await out.copyPages(srcPdf, [page.sourcePageIndex])
        if (page.rotation) {
          copied.setRotation(degrees(page.rotation))
        }
        out.addPage(copied)
      }

      const bytes = await out.save({ useObjectStreams: false })
      const blob = new Blob([bytes], { type: 'application/pdf' })
      const filename = `organized-${Date.now()}-${selectedFiles[0].name || 'document.pdf'}`
      downloadBlob(blob, filename)
      setOrganizeStatus('Organized PDF downloaded.')
    } catch (error) {
      setOrganizeStatus(`Could not organize PDF: ${error?.message || 'unknown error'}`)
    } finally {
      setOrganizeBusy(false)
    }
  }

  const setCurrentPageObjects = (updater) => {
    const prevAll = editObjectsRef.current
    const current = prevAll[editPage] || []
    const next = typeof updater === 'function' ? updater(current) : updater
    const nextAll = { ...prevAll, [editPage]: next }
    editObjectsRef.current = nextAll
    setEditObjectsByPage(nextAll)
  }

  const drawEditObjects = (ctx, objects, selectedId = '', activePath = null) => {
    ctx.clearRect(0, 0, ctx.canvas.width, ctx.canvas.height)

    objects.forEach((obj) => {
      if (obj.type === 'path' && obj.points?.length > 1) {
        ctx.strokeStyle = '#ff3b3b'
        ctx.lineWidth = 2
        ctx.lineCap = 'round'
        ctx.beginPath()
        ctx.moveTo(obj.points[0].x, obj.points[0].y)
        obj.points.forEach((pt) => ctx.lineTo(pt.x, pt.y))
        ctx.stroke()
      }

      if (obj.type === 'text') {
        ctx.fillStyle = '#1c2a43'
        ctx.font = `${obj.size || 28}px Inter`
        ctx.fillText(obj.text, obj.x, obj.y)
      }

      if (obj.type === 'rect') {
        ctx.strokeStyle = '#1f7dff'
        ctx.lineWidth = 2
        ctx.strokeRect(obj.x, obj.y, obj.w, obj.h)
      }

      if (obj.type === 'image' && obj.dataUrl) {
        const cached = editImageCacheRef.current[obj.dataUrl]
        if (cached?.complete) {
          ctx.drawImage(cached, obj.x, obj.y, obj.w, obj.h)
        } else {
          const image = new Image()
          image.src = obj.dataUrl
          image.onload = () => {
            editImageCacheRef.current[obj.dataUrl] = image
            ctx.drawImage(image, obj.x, obj.y, obj.w, obj.h)
          }
          editImageCacheRef.current[obj.dataUrl] = image
        }
      }

      if (selectedId && obj.id === selectedId) {
        ctx.strokeStyle = '#ffffff'
        ctx.setLineDash([4, 4])
        if (obj.type === 'text') {
          ctx.strokeRect(obj.x - 6, obj.y - 34, (obj.w || 160) + 12, (obj.h || 34) + 8)
        } else if (obj.type === 'path' && obj.points?.length) {
          const xs = obj.points.map((p) => p.x)
          const ys = obj.points.map((p) => p.y)
          const minX = Math.min(...xs)
          const maxX = Math.max(...xs)
          const minY = Math.min(...ys)
          const maxY = Math.max(...ys)
          ctx.strokeRect(minX - 6, minY - 6, maxX - minX + 12, maxY - minY + 12)
        } else {
          ctx.strokeRect(obj.x - 4, obj.y - 4, obj.w + 8, obj.h + 8)
        }
        ctx.setLineDash([])
      }
    })

    if (activePath?.length > 1) {
      ctx.strokeStyle = '#ff3b3b'
      ctx.lineWidth = 2
      ctx.lineCap = 'round'
      ctx.beginPath()
      ctx.moveTo(activePath[0].x, activePath[0].y)
      activePath.forEach((pt) => ctx.lineTo(pt.x, pt.y))
      ctx.stroke()
    }
  }

  useEffect(() => {
    if (activeTool?.slug !== 'edit-pdf' || !selectedEditFile) {
      setEditPdfProxy(null)
      setEditPdfData(null)
      setEditPageCount(1)
      return
    }

    let cancelled = false
    ;(async () => {
      const rawBuffer = await selectedEditFile.arrayBuffer()
      const bytes = new Uint8Array(rawBuffer)
      if (cancelled) return
      setEditPdfData(bytes)
      const pdf = await pdfjsLib.getDocument({ data: bytes.slice() }).promise
      if (cancelled) return
      setEditPdfProxy(pdf)
      setEditPageCount(pdf.numPages)
      setEditPage(1)
      setEditObjectsByPage({})
      editObjectsRef.current = {}
      setEditSelectedObjectId('')
      setEditStatus('')
    })()

    return () => {
      cancelled = true
    }
  }, [activeTool?.slug, selectedEditFile])

  useEffect(() => {
    if (activeTool?.slug !== 'edit-pdf') return
    if (!editPdfProxy || !editBaseCanvasRef.current || !editOverlayCanvasRef.current) return

    let cancelled = false
    ;(async () => {
      const page = await editPdfProxy.getPage(editPage)
      if (cancelled) return
      const viewport = page.getViewport({ scale: 1.15 })
      const base = editBaseCanvasRef.current
      const overlay = editOverlayCanvasRef.current
      base.width = viewport.width
      base.height = viewport.height
      overlay.width = viewport.width
      overlay.height = viewport.height
      const ctx = base.getContext('2d')
      await page.render({ canvasContext: ctx, viewport }).promise
      if (cancelled) return
      setEditCanvasMetaByPage((prev) => ({
        ...prev,
        [editPage]: { width: viewport.width, height: viewport.height },
      }))
      const overlayCtx = overlay.getContext('2d')
      drawEditObjects(overlayCtx, editObjectsByPage[editPage] || [], editSelectedObjectId, editActivePathRef.current)
      requestAnimationFrame(updateEditFitScale)
    })()

    return () => {
      cancelled = true
    }
  }, [activeTool?.slug, editPdfProxy, editPage])

  useEffect(() => {
    if (activeTool?.slug !== 'edit-pdf') return
    const onResize = () => updateEditFitScale()
    window.addEventListener('resize', onResize)
    return () => window.removeEventListener('resize', onResize)
  }, [activeTool?.slug])

  useEffect(() => {
    if (activeTool?.slug !== 'edit-pdf' || !editOverlayCanvasRef.current) return
    const ctx = editOverlayCanvasRef.current.getContext('2d')
    drawEditObjects(ctx, currentPageObjects, editSelectedObjectId, editActivePathRef.current)
  }, [activeTool?.slug, currentPageObjects, editSelectedObjectId, editPage])

  const onEditOverlayMouseDown = (event) => {
    const point = getCanvasPoint(event)

    if (editTool === 'annotate') {
      editActivePathRef.current = [point]
      return
    }

    if (editTool === 'text') {
      const text = window.prompt('Enter text')
      if (!text) return
      const next = {
        id: crypto.randomUUID(),
        type: 'text',
        text,
        x: point.x,
        y: point.y,
        size: 28,
        w: Math.max(80, text.length * 14),
        h: 34,
      }
      setCurrentPageObjects((prev) => [...prev, next])
      setEditSelectedObjectId(next.id)
      return
    }

    if (editTool === 'draw') {
      const next = {
        id: crypto.randomUUID(),
        type: 'rect',
        x: point.x - 80,
        y: point.y - 50,
        w: 160,
        h: 100,
      }
      setCurrentPageObjects((prev) => [...prev, next])
      setEditSelectedObjectId(next.id)
      return
    }

    if (editTool === 'image') {
      if (!editPendingImageData) {
        editImageInputRef.current?.click()
        return
      }
      const next = {
        id: crypto.randomUUID(),
        type: 'image',
        dataUrl: editPendingImageData,
        x: point.x - 90,
        y: point.y - 65,
        w: 180,
        h: 130,
      }
      setCurrentPageObjects((prev) => [...prev, next])
      setEditSelectedObjectId(next.id)
      setEditPendingImageData('')
      return
    }

    if (editTool === 'edit') {
      const hitId = hitTestEditObject(currentPageObjects, point.x, point.y)
      setEditSelectedObjectId(hitId)
      if (hitId) {
        editDragRef.current = { id: hitId, start: point }
      }
    }
  }

  const onEditOverlayMouseMove = (event) => {
    const point = getCanvasPoint(event)

    if (editTool === 'annotate' && editActivePathRef.current) {
      editActivePathRef.current = [...editActivePathRef.current, point]
      if (editOverlayCanvasRef.current) {
        const ctx = editOverlayCanvasRef.current.getContext('2d')
        drawEditObjects(ctx, currentPageObjects, editSelectedObjectId, editActivePathRef.current)
      }
      return
    }

    if (editTool === 'edit' && editDragRef.current) {
      const { id, start } = editDragRef.current
      const dx = point.x - start.x
      const dy = point.y - start.y
      editDragRef.current = { id, start: point }
      setCurrentPageObjects((prev) =>
        prev.map((obj) => (obj.id === id ? { ...obj, x: (obj.x || 0) + dx, y: (obj.y || 0) + dy } : obj)),
      )
    }
  }

  const onEditOverlayMouseUp = () => {
    if (editTool === 'annotate' && editActivePathRef.current?.length > 1) {
      const next = {
        id: crypto.randomUUID(),
        type: 'path',
        points: editActivePathRef.current,
      }
      setCurrentPageObjects((prev) => [...prev, next])
      setEditSelectedObjectId(next.id)
    }
    editActivePathRef.current = null
    editDragRef.current = null
  }

  const onEditImagePick = (event) => {
    const file = event.target.files?.[0]
    if (!file) return
    const reader = new FileReader()
    reader.onload = () => {
      setEditPendingImageData(String(reader.result || ''))
      setEditStatus('Image selected. Click on PDF to place it.')
    }
    reader.readAsDataURL(file)
    event.target.value = ''
  }

  const onEditSave = async () => {
    if (!selectedEditFile || !editPdfData) {
      setEditStatus('No PDF selected to save.')
      return
    }

    try {
      setEditStatus('Saving edited PDF...')
      const objectsByPageSnapshot = JSON.parse(JSON.stringify(editObjectsRef.current || {}))

      if (editActivePathRef.current?.length > 1) {
        const draftPath = {
          id: crypto.randomUUID(),
          type: 'path',
          points: [...editActivePathRef.current],
        }
        const currentObjects = objectsByPageSnapshot[editPage] || []
        objectsByPageSnapshot[editPage] = [...currentObjects, draftPath]
        setCurrentPageObjects((prev) => [...prev, draftPath])
        setEditSelectedObjectId(draftPath.id)
        editActivePathRef.current = null
      }

      const out = await PDFDocument.load(editPdfData.slice(), {
        ignoreEncryption: true,
        parseSpeed: ParseSpeeds.Fastest,
      })
      const pageKeys = Object.keys(objectsByPageSnapshot)
      let appliedPages = 0
      let skippedPages = 0
      let appliedObjects = 0

      const renderProxy = editPdfProxy || await pdfjsLib.getDocument({ data: editPdfData.slice() }).promise
      const canvasToPngBytes = async (canvas) => {
        const blob = await new Promise((resolve, reject) => {
          canvas.toBlob((value) => {
            if (value) {
              resolve(value)
            } else {
              reject(new Error('Canvas export failed'))
            }
          }, 'image/png')
        })
        const arr = await blob.arrayBuffer()
        return new Uint8Array(arr)
      }

      for (const key of pageKeys) {
        const pageNum = Number(key)
        if (!Number.isFinite(pageNum) || pageNum < 1 || pageNum > out.getPageCount()) {
          skippedPages += 1
          continue
        }
        const objects = objectsByPageSnapshot[pageNum] || []
        if (!objects.length) continue
        try {
          const srcPage = await renderProxy.getPage(pageNum)
          const baseViewport = srcPage.getViewport({ scale: 1 })
          const meta = editCanvasMetaByPage[pageNum]
          const scale = meta?.width && baseViewport.width
            ? meta.width / baseViewport.width
            : 1.15
          const viewport = srcPage.getViewport({ scale })
          const canvas = document.createElement('canvas')
          canvas.width = Math.ceil(viewport.width)
          canvas.height = Math.ceil(viewport.height)
          const ctx = canvas.getContext('2d')
          if (!ctx) {
            skippedPages += 1
            continue
          }
          await srcPage.render({ canvasContext: ctx, viewport }).promise

          const overlay = document.createElement('canvas')
          overlay.width = canvas.width
          overlay.height = canvas.height
          const overlayCtx = overlay.getContext('2d')
          if (!overlayCtx) {
            skippedPages += 1
            continue
          }
          drawEditObjects(overlayCtx, objects)
          ctx.drawImage(overlay, 0, 0)

          const pngBytes = await canvasToPngBytes(canvas)
          const png = await out.embedPng(pngBytes)

          const outPage = out.getPage(pageNum - 1)
          const { width: outW, height: outH } = outPage.getSize()
          outPage.drawImage(png, { x: 0, y: 0, width: outW, height: outH })

          appliedPages += 1
          appliedObjects += objects.length
        } catch (pageError) {
          skippedPages += 1
          console.error(`Failed to burn edits on page ${pageNum}`, pageError)
        }
      }

      const bytes = await out.save({ useObjectStreams: false })
      const blob = new Blob([bytes], { type: 'application/pdf' })
      const filename = `edited-${Date.now()}-${selectedEditFile.name || 'document.pdf'}`
      downloadBlob(blob, filename)
      if (appliedObjects === 0) {
        setEditStatus('Downloaded PDF, but no edits were detected to save.')
      } else if (skippedPages > 0) {
        setEditStatus(`Saved ${appliedObjects} edit(s). Applied on ${appliedPages} page(s), skipped ${skippedPages}.`)
      } else {
        setEditStatus(`Saved edited PDF with ${appliedObjects} edit(s).`)
      }
    } catch (error) {
      console.error('Edit PDF save failed:', error)
      setEditStatus(`Could not save edited PDF: ${error?.message || 'unknown error'}`)
    }
  }

  const editTools = [
    { key: 'annotate', label: 'Annotate' },
    { key: 'edit', label: 'Edit' },
    { key: 'text', label: 'Text' },
    { key: 'image', label: 'Image' },
    { key: 'draw', label: 'Draw' },
  ]

  const splitPreviewPages = useMemo(() => {
    if (!splitTotalPages) {
      return [1, 2]
    }

    const from = clampPage(splitFromPage, splitTotalPages)
    const to = clampPage(splitToPage, splitTotalPages)

    if (splitMode === 'size') {
      const chunkSize = Math.max(1, from)
      const chunkNumber = Math.max(1, to)
      const start = clampPage((chunkNumber - 1) * chunkSize + 1, splitTotalPages)
      const end = clampPage(start + chunkSize - 1, splitTotalPages)
      return [start, end]
    }

    return [from, to]
  }, [splitFromPage, splitMode, splitToPage, splitTotalPages])

  const translateTargetOptions = useMemo(() => {
    if (translateSourceLang === 'auto') {
      return translateLanguages.filter((lang) => lang.code !== 'auto')
    }
    const src = translateLanguages.find((lang) => lang.code === translateSourceLang)
    const supported = new Set(src?.targets || [])
    return translateLanguages.filter((lang) => lang.code !== 'auto' && supported.has(lang.code))
  }, [translateLanguages, translateSourceLang])

  useEffect(() => {
    if (activeTool?.slug !== 'translate-pdf') return
    if (!translateTargetOptions.length) return
    if (!translateTargetOptions.some((lang) => lang.code === translateTargetLang)) {
      setTranslateTargetLang(translateTargetOptions[0].code)
    }
  }, [activeTool?.slug, translateTargetLang, translateTargetOptions])

  return (
    <div className="page">
      <header className="topbar">
        <a className="brand" href="/" aria-label="TRULYPDF home" onClick={(event) => navigate(event, '/')}>
          <span className="brand-mark">TRULY</span>
          <span className="brand-name">PDF</span>
        </a>

        <nav className="menu" aria-label="Main navigation">
          {navRoutes.map((item) => (
            <a
              key={item.path}
              href={item.path}
              onClick={(event) => navigate(event, item.path)}
              className={pathname === item.path ? 'active' : ''}
            >
              {item.label}
            </a>
          ))}
        </nav>

      </header>

      <main className={`content-wrap ${activeTool ? 'tool-layout' : ''}`}>
        {activeTool ? (
          <section
            className={`converter-page ${isDragOver ? 'drag-over' : ''} ${
              activeTool?.slug === 'rotate-pdf'
                ? 'rotate-mode'
                : activeTool?.slug === 'organize-pdf'
                  ? 'organize-mode'
                  : activeTool?.slug === 'split-pdf'
                    ? 'split-mode'
                    : activeTool?.slug === 'edit-pdf'
                      ? 'edit-mode'
                      : activeTool?.slug === 'page-numbers'
                        ? 'page-number-mode'
                        : activeTool?.slug === 'compare-pdf'
                          ? 'compare-mode'
                          : activeTool?.slug === 'crop-pdf'
                            ? 'crop-mode'
                          : activeTool?.slug === 'scan-to-pdf'
                              ? 'scan-mode'
                            : activeTool?.slug === 'pdf-to-jpg'
                              ? 'pdf-jpg-mode'
                            : activeTool?.slug === 'html-to-pdf'
                              ? 'html-mode'
                            : activeTool?.slug === 'ai-summarizer'
                              ? 'ai-mode'
                        : ''
            }`}
            aria-label={`${activeTool.title} page`}
            onDragOver={onDragOver}
            onDragLeave={onDragLeave}
            onDrop={onDrop}
          >
            <div className={`converter-shell ${selectedFiles.length === 0 ? 'upload-mode' : 'workspace-mode'}`}>
              {selectedFiles.length === 0 && (
                <>
                  <h1>{activeTool.title} Converter</h1>
                </>
              )}

              <input
                ref={fileInputRef}
                type="file"
                accept={allowedInputAccept}
                multiple
                className="hidden-file-input"
                onChange={onInputChange}
              />

              {activeTool.slug === 'ai-summarizer' ? (
                <section className="ai-chat-page" aria-label="AI Summarizer chat">
                  <div className="ai-chat-header">
                    <h2>AI Summarizer Chat</h2>
                    <p>Upload a file, describe what you want, and generate a summary.</p>
                  </div>

                  <div className="ai-chat-messages">
                    {aiSummaryMessages.map((msg) => (
                      <article key={msg.id} className={`ai-chat-bubble ${msg.role}`}>
                        <p>{msg.text}</p>
                      </article>
                    ))}
                  </div>

                  <div className="ai-chat-composer">
                    <div className="ai-chat-inputbar">
                      <button type="button" className="ai-plus-btn" onClick={openFilePicker} aria-label="Upload file">
                        +
                      </button>
                      <input
                        type="text"
                        className="ai-chat-inline-input"
                        value={aiSummaryPrompt}
                        onChange={(event) => setAiSummaryPrompt(event.target.value)}
                        placeholder="Ask anything"
                      />
                      <button type="button" className="ai-send-btn" onClick={onConvert} disabled={isMerging || !selectedFiles.length}>
                        {isMerging ? '...' : 'Summarize'}
                      </button>
                    </div>
                    <div className="ai-chat-actions">
                      <div className="ai-file-list">
                        {selectedFiles.length ? selectedFiles.map((file, index) => (
                          <div key={`${file.name}-${index}`} className="ai-file-chip">
                            <span>{file.name}</span>
                            <button type="button" onClick={() => removeFileAt(index)} aria-label={`Remove ${file.name}`}>
                              x
                            </button>
                          </div>
                        )) : <p className="ai-file-empty">No file selected.</p>}
                      </div>
                      <div className="ai-chat-action-buttons">
                        <button type="button" className="ai-secondary-btn" onClick={downloadSummaryChat} disabled={!aiSummaryMessages.length}>
                          Download Chat
                        </button>
                      </div>
                    </div>
                  </div>
                </section>
              ) : activeTool.slug === 'merge-pdf' && mergedFileUrl ? (
                <section className="merge-result-page" aria-label="Merge complete">
                  <h2>PDFs have been merged!</h2>
                  <div className="merge-result-actions">
                    <button type="button" className="back-round-btn" onClick={resetMergeResult} aria-label="Back">
                      ←
                    </button>
                    <button
                      type="button"
                      className="download-merged-btn"
                      onClick={() => downloadBlobUrl(mergedFileUrl, mergedFileName)}
                    >
                      Download merged PDF
                    </button>
                    <div className="result-mini-actions">
                      <button
                        type="button"
                        className="result-delete-btn"
                        aria-label="Delete merged file"
                        onClick={() => {
                          resetMergeResult()
                          setSelectedFiles([])
                        }}
                      >
                        🗑
                      </button>
                    </div>
                  </div>

                  <div className="continue-box">
                    <p>Continue to...</p>
                    <div className="continue-grid">
                      {mergeSuggestions.map((item) => (
                        <button
                          key={item.label}
                          type="button"
                          className="continue-item"
                          onClick={() => goToPath(item.path)}
                        >
                          <span>{item.label}</span>
                          <span>›</span>
                        </button>
                      ))}
                    </div>
                  </div>
                </section>
              ) : activeTool.slug === 'scan-to-pdf' ? (
                <section className="scan-workspace-page" aria-label="Scan to PDF workspace">
                  <input
                    ref={scanUploadInputRef}
                    type="file"
                    accept="image/*"
                    capture="environment"
                    multiple
                    className="hidden-file-input"
                    onChange={onScanFileInput}
                  />
                  <canvas ref={scanCaptureCanvasRef} className="hidden-file-input" />

                  <div className="scan-left">
                    <div className="scan-camera-card">
                      <div className="scan-camera-feed">
                        <video ref={scanVideoRef} autoPlay muted playsInline />
                        {!scanCameraOn && (
                          <div className="scan-camera-placeholder">
                            Camera is off. Add pages from files or start camera.
                          </div>
                        )}
                      </div>
                      <div className="scan-camera-actions">
                        <button type="button" onClick={startScanCamera}>Start camera</button>
                        <button type="button" onClick={stopScanCamera}>Stop camera</button>
                        <button type="button" onClick={captureScanPage} disabled={!scanCameraOn}>Capture page</button>
                        <button type="button" onClick={() => scanUploadInputRef.current?.click()}>Add pages</button>
                      </div>
                    </div>

                    <div className="scan-pages-grid">
                      {scanPages.map((page, index) => (
                        <article
                          key={page.id}
                          className={`scan-page-tile ${page.id === selectedScanPage?.id ? 'active' : ''}`}
                          onClick={() => setScanSelectedId(page.id)}
                          onKeyDown={(event) => {
                            if (event.key === 'Enter' || event.key === ' ') {
                              event.preventDefault()
                              setScanSelectedId(page.id)
                            }
                          }}
                          role="button"
                          tabIndex={0}
                        >
                          <img src={page.dataUrl} alt={page.label || `Scan page ${index + 1}`} />
                          <p>{page.label || `Page ${index + 1}`}</p>
                          <button
                            type="button"
                            onClick={(event) => {
                              event.stopPropagation()
                              removeScanPage(page.id)
                            }}
                          >
                            Remove
                          </button>
                        </article>
                      ))}
                    </div>
                  </div>

                  <aside className="scan-sidebar">
                    <h2>Scan to PDF</h2>
                    <p>{scanStatus || 'Capture, crop, reorder by selection, and export pages as one PDF.'}</p>

                    <div className="scan-crop-preview">
                      {selectedScanPage ? (
                        <div className="scan-crop-preview-frame">
                          <img
                            src={selectedScanPage.dataUrl}
                            alt="Selected scan crop preview"
                            style={{
                              clipPath: `inset(${selectedScanPage.crop.top}% ${selectedScanPage.crop.right}% ${selectedScanPage.crop.bottom}% ${selectedScanPage.crop.left}%)`,
                            }}
                          />
                        </div>
                      ) : (
                        <div className="scan-crop-empty">No pages added</div>
                      )}
                    </div>

                    <div className="scan-crop-controls">
                      <label>
                        Left crop ({selectedScanPage?.crop.left ?? 0}%)
                        <input
                          type="range"
                          min="0"
                          max="45"
                          value={selectedScanPage?.crop.left ?? 0}
                          onChange={(event) => updateSelectedScanCrop('left', event.target.value)}
                          disabled={!selectedScanPage}
                        />
                      </label>
                      <label>
                        Top crop ({selectedScanPage?.crop.top ?? 0}%)
                        <input
                          type="range"
                          min="0"
                          max="45"
                          value={selectedScanPage?.crop.top ?? 0}
                          onChange={(event) => updateSelectedScanCrop('top', event.target.value)}
                          disabled={!selectedScanPage}
                        />
                      </label>
                      <label>
                        Right crop ({selectedScanPage?.crop.right ?? 0}%)
                        <input
                          type="range"
                          min="0"
                          max="45"
                          value={selectedScanPage?.crop.right ?? 0}
                          onChange={(event) => updateSelectedScanCrop('right', event.target.value)}
                          disabled={!selectedScanPage}
                        />
                      </label>
                      <label>
                        Bottom crop ({selectedScanPage?.crop.bottom ?? 0}%)
                        <input
                          type="range"
                          min="0"
                          max="45"
                          value={selectedScanPage?.crop.bottom ?? 0}
                          onChange={(event) => updateSelectedScanCrop('bottom', event.target.value)}
                          disabled={!selectedScanPage}
                        />
                      </label>
                    </div>

                    <button
                      type="button"
                      className="scan-main-btn"
                      onClick={exportScanToPdf}
                      disabled={scanBusy || !scanPages.length}
                    >
                      {scanBusy ? 'Exporting...' : 'Export to PDF'}
                    </button>
                  </aside>
                </section>
              ) : selectedFiles.length === 0 ? (
                <>
                  {activeTool.slug === 'html-to-pdf' ? (
                    <section className="html-modal-overlay" aria-label="Add HTML source">
                      <article className="html-modal-card">
                        <div className="html-modal-head">
                          <h2>Add HTML to convert from</h2>
                          <button
                            type="button"
                            className="html-modal-close"
                            aria-label="Close dialog"
                            onClick={() => {
                              setHtmlToPdfUrl('')
                              setHtmlToPdfStatus('')
                              goToPath('/convert')
                            }}
                          >
                            ×
                          </button>
                        </div>
                        <div className="html-modal-body">
                          <div className="html-tab-row">
                            <button type="button" className="html-tab active">
                              Url
                            </button>
                          </div>
                          <label className="html-url-label" htmlFor="html-url-input">
                            Write the website URL
                          </label>
                          <div className="html-url-field">
                            <span>◎</span>
                            <input
                              id="html-url-input"
                              type="url"
                              placeholder="Example: https://ilovepdf.com"
                              value={htmlToPdfUrl}
                              onChange={(event) => {
                                setHtmlToPdfUrl(event.target.value)
                                if (htmlToPdfStatus) setHtmlToPdfStatus('')
                              }}
                            />
                          </div>
                          {htmlToPdfStatus && <p className="html-url-status">{htmlToPdfStatus}</p>}
                        </div>
                        <div className="html-modal-foot">
                          <button type="button" className="html-add-btn" onClick={onAddHtmlSource}>
                            Add
                          </button>
                        </div>
                      </article>
                    </section>
                  ) : (
                    <>
                      <div className="converter-uploader">
                        <button type="button" className="select-file-btn" onClick={openFilePicker}>
                          {selectFileLabel}
                        </button>
                      </div>
                      <p className="drop-hint">{dropHintLabel}</p>
                    </>
                  )}
                </>
              ) : (
                activeTool.slug === 'edit-pdf' ? (
                  <section className="edit-workspace-page" aria-label="Edit PDF workspace">
                    <input
                      ref={editImageInputRef}
                      type="file"
                      accept="image/*"
                      className="hidden-file-input"
                      onChange={onEditImagePick}
                    />
                    <aside className="edit-thumbnails">
                      {selectedFiles.map((file, index) => (
                        <button
                          type="button"
                          key={`${file.name}-${index}`}
                          className={`edit-thumb ${index === editSelectedIndex ? 'active' : ''}`}
                          onClick={() => setEditSelectedIndex(index)}
                        >
                          <iframe
                            className="edit-thumb-frame"
                            src={`${previewUrls[index]}#page=1&view=FitH`}
                            title={`Thumbnail ${file.name}`}
                          />
                          <span>{index + 1}</span>
                        </button>
                      ))}
                    </aside>

                    <div className="edit-main">
                      <div className="edit-topbar">
                        {editTools.map((tool) => (
                          <button
                            key={tool.key}
                            type="button"
                            className={editTool === tool.key ? 'active' : ''}
                            onClick={() => {
                              setEditTool(tool.key)
                              setEditStatus(`${tool.label} mode active`)
                              if (tool.key === 'image') {
                                editImageInputRef.current?.click()
                              }
                            }}
                          >
                            {tool.label}
                          </button>
                        ))}
                        {editSelectedObjectId && (
                          <button
                            type="button"
                            onClick={() =>
                              {
                                setCurrentPageObjects((prev) =>
                                  prev.filter((obj) => obj.id !== editSelectedObjectId),
                                )
                                setEditSelectedObjectId('')
                              }
                            }
                          >
                            Delete
                          </button>
                        )}
                      </div>

                      <div className="edit-canvas-wrap">
                        <div className="edit-zoom-readout">{Math.round(editZoom * editFitScale)}%</div>
                        {selectedEditPreviewUrl && editPdfProxy ? (
                          <div
                            ref={editCanvasWrapRef}
                            className="edit-canvas-fit-wrap"
                          >
                            <div
                            className="edit-canvas-stage"
                              style={{ transform: `scale(${(editZoom / 100) * editFitScale})` }}
                          >
                            <canvas ref={editBaseCanvasRef} className="edit-canvas-base" />
                            <canvas
                              ref={editOverlayCanvasRef}
                              className="edit-canvas-overlay"
                              onMouseDown={onEditOverlayMouseDown}
                              onMouseMove={onEditOverlayMouseMove}
                              onMouseUp={onEditOverlayMouseUp}
                              onMouseLeave={onEditOverlayMouseUp}
                            />
                          </div>
                          </div>
                        ) : (
                          <div className="edit-canvas-empty">No file selected</div>
                        )}
                      </div>
                    </div>

                    <aside className="edit-sidebar">
                      <h2>Edit PDF</h2>
                      <p>
                        {editStatus ||
                          'Use the toolbar to add text, images, shapes, and annotations to your document.'}
                      </p>

                      <div className="edit-controls">
                        <button
                          type="button"
                          onClick={() => setEditPage((prev) => Math.max(1, prev - 1))}
                        >
                          Prev page
                        </button>
                        <button
                          type="button"
                          onClick={() => setEditPage((prev) => Math.min(editPageCount, prev + 1))}
                        >
                          Next page
                        </button>
                        <button
                          type="button"
                          onClick={() => setEditZoom((prev) => Math.max(40, prev - 10))}
                        >
                          Zoom -
                        </button>
                        <button
                          type="button"
                          onClick={() => setEditZoom((prev) => Math.min(130, prev + 10))}
                        >
                          Zoom +
                        </button>
                        <button type="button" onClick={openFilePicker}>
                          Add PDF
                        </button>
                        <button type="button" onClick={() => setEditStatus(`Page ${editPage} / ${editPageCount}`)}>
                          {editPage} / {editPageCount}
                        </button>
                      </div>

                      <button type="button" className="edit-save-btn" onClick={onEditSave}>
                        Save changes
                      </button>
                    </aside>
                  </section>
                ) : activeTool.slug === 'split-pdf' ? (
                  splitResultUrl ? (
                    <section className="split-result-page" aria-label="Split complete">
                      <h2>PDF has been split!</h2>
                      <div className="split-result-actions">
                        <button type="button" className="back-round-btn" onClick={resetSplitResult} aria-label="Back">
                          ←
                        </button>
                        <button
                          type="button"
                          className="download-merged-btn"
                          onClick={() => downloadBlobUrl(splitResultUrl, splitResultName)}
                        >
                          Download split PDF
                        </button>
                      </div>
                    </section>
                  ) : (
                  <section className="split-workspace-page" aria-label="Split PDF workspace">
                    <div className="split-left">
                      <p className="split-range-title">Range 1</p>
                      <div className="split-range-box">
                        {splitPreviewPages.map((pageNo, cardIdx) => (
                          <article key={`${pageNo}-${cardIdx}`} className="split-page-card">
                            <div className="split-preview-toolbar">
                              <button
                                type="button"
                                onClick={() => {
                                  if (splitMode === 'size') {
                                    setSplitToPage((prev) => prev + 1)
                                  } else {
                                    setSplitToPage((prev) => clampPage(prev + 1, splitTotalPages || 999))
                                  }
                                }}
                                aria-label="Expand range"
                              >
                                △+
                              </button>
                              <button type="button" onClick={downloadSourceFile} aria-label="Download source">
                                ↓
                              </button>
                              <button type="button" onClick={() => copyPageLink(pageNo - 1)} aria-label="Copy page link">
                                ⎘
                              </button>
                              <button type="button" onClick={openFilePicker} aria-label="Add another file">
                                ⋮
                              </button>
                            </div>
                            <div className="split-page-preview">
                              <iframe
                                className="file-preview-frame"
                                src={`${previewUrls[0]}#page=${pageNo}&view=FitH`}
                                title={`Split preview ${pageNo}`}
                              />
                            </div>
                            <p className="split-page-number">{pageNo}</p>
                          </article>
                        ))}
                      </div>
                    </div>

                    <aside className="split-sidebar">
                      <h2>Split</h2>

                      <div className="split-mode-tabs">
                        <button
                          type="button"
                          className={splitMode === 'range' ? 'active' : ''}
                          onClick={() => setSplitMode('range')}
                        >
                          <span className="split-tab-icon">◫</span>
                          <span>Range</span>
                        </button>
                          <button
                            type="button"
                            className={splitMode === 'pages' ? 'active' : ''}
                            onClick={() => {
                              setSplitMode('pages')
                              if (splitTotalPages > 1) {
                                setSplitFromPage(1)
                                setSplitToPage(2)
                              }
                            }}
                          >
                          <span className="split-tab-icon">◧</span>
                          <span>Pages</span>
                        </button>
                          <button
                            type="button"
                            className={splitMode === 'size' ? 'active' : ''}
                            onClick={() => {
                              setSplitMode('size')
                              setSplitFromPage(1)
                              setSplitToPage(1)
                            }}
                          >
                          <span className="split-tab-icon">◨</span>
                          <span>Size</span>
                        </button>
                      </div>

                      <div className="split-settings">
                        <p>Range mode:</p>
                        <div className="split-range-type">
                          <button
                            type="button"
                            className={splitRangeType === 'custom' ? 'active' : ''}
                            onClick={() => setSplitRangeType('custom')}
                          >
                            Custom
                          </button>
                          <button
                            type="button"
                            className={splitRangeType === 'fixed' ? 'active' : ''}
                            onClick={() => setSplitRangeType('fixed')}
                          >
                            Fixed
                          </button>
                        </div>

                        <div className="split-range-inputs">
                          <label>
                            from page
                            <input
                              type="number"
                              min="1"
                              max={splitTotalPages || undefined}
                              value={splitFromPage}
                              onChange={(event) => setSplitFromPage(Number(event.target.value || 1))}
                              disabled={splitRangeType === 'fixed'}
                            />
                          </label>
                          <label>
                            to page
                            <input
                              type="number"
                              min="1"
                              max={splitTotalPages || undefined}
                              value={splitToPage}
                              onChange={(event) => setSplitToPage(Number(event.target.value || 1))}
                              disabled={splitRangeType === 'fixed'}
                            />
                          </label>
                        </div>
                      </div>

                      <button type="button" className="split-main-btn" onClick={handleSplit} disabled={splitBusy}>
                        {splitBusy ? 'Splitting...' : 'Split PDF'}
                      </button>
                    </aside>
                  </section>
                  )
                ) : activeTool.slug === 'organize-pdf' ? (
                  <section className="organize-workspace-page" aria-label="Organize PDF workspace">
                    <div className="organize-left">
                      <button type="button" className="organize-add-btn" onClick={openFilePicker} aria-label="Add PDFs">
                        +
                        <span>{selectedFiles.length}</span>
                      </button>
                      <button
                        type="button"
                        className="organize-blank-btn"
                        onClick={() => organizeAddBlank(organizePages.length - 1)}
                        aria-label="Add blank page"
                      >
                        + blank
                      </button>
                      <button
                        type="button"
                        className="organize-sort-btn"
                        onClick={toggleOrganizeOrder}
                        aria-label="Change arrangement"
                        title="Change arrangement of entire document"
                      >
                        <span>{organizeSortDirection === 'asc' ? '↓' : '↑'}</span>
                        <span>1</span>
                        <span>9</span>
                      </button>

                      <div
                        className="organize-pages-grid"
                        onDragOver={(event) => event.preventDefault()}
                        onDrop={onOrganizeGridDrop}
                      >
                        {organizePages.map((page, index) => (
                          <article
                            key={page.id}
                            className={`organize-page-card ${organizeDragIndex === index ? 'dragging' : ''}`}
                            draggable
                            onDragStart={(event) => onOrganizeDragStart(event, index)}
                            onDragOver={(event) => event.preventDefault()}
                            onDrop={(event) => onOrganizeDrop(event, index)}
                            onDragEnd={() => setOrganizeDragIndex(null)}
                          >
                            <div className="organize-card-actions">
                              <button type="button" onClick={() => organizeRotate(page.id, 90)} aria-label="Rotate page">
                                ↻
                              </button>
                              <button type="button" onClick={() => organizeRemove(page.id)} aria-label="Delete page">
                                ×
                              </button>
                            </div>

                            <div className="organize-page-preview">
                              {page.type === 'source' ? (
                                <iframe
                                  className="organize-page-frame"
                                  style={{ transform: `rotate(${page.rotation}deg)` }}
                                  src={`${previewUrls[0]}#page=${page.sourcePageIndex + 1}&view=FitH`}
                                  title={`Organize page ${index + 1}`}
                                />
                              ) : (
                                <div className="organize-blank-preview" style={{ transform: `rotate(${page.rotation}deg)` }}>
                                  blank
                                </div>
                              )}
                            </div>
                            <p>{page.type === 'source' ? page.label : 'blank'}</p>
                          </article>
                        ))}
                      </div>
                    </div>

                    <aside className="organize-sidebar">
                      <h2>Organize PDF</h2>
                      <div className="organize-file-row">
                        <span>Files:</span>
                        <button type="button" onClick={organizeResetAll}>Reset all</button>
                      </div>
                      <div className="organize-file-box">
                        A: {selectedFiles[0]?.name || 'No file'}
                      </div>
                      <p className="organize-status">{organizeStatus}</p>
                      <button
                        type="button"
                        className="organize-main-btn"
                        onClick={handleOrganizeDownload}
                        disabled={organizeBusy}
                      >
                        {organizeBusy ? 'Organizing...' : 'Organize'}
                      </button>
                    </aside>
                  </section>
                ) : activeTool.slug === 'rotate-pdf' ? (
                  <section className="rotate-workspace-page" aria-label="Rotate PDF workspace">
                    <div className="rotate-left">
                      <button type="button" className="rotate-add-btn" onClick={openFilePicker} aria-label="Add PDFs">
                        +
                        <span>{selectedFiles.length}</span>
                      </button>
                      <div className="rotate-files-grid">
                        {selectedFiles.map((file, index) => {
                          const key = getRotateFileKey(file, index)
                          const angle = rotateAnglesByFile[key] || 0
                          return (
                            <button
                              type="button"
                              key={key}
                              className={`rotate-file-card ${index === rotateSelectedIndex ? 'active' : ''}`}
                              onClick={() => setRotateSelectedIndex(index)}
                            >
                              <div className="rotate-file-frame-wrap">
                                <iframe
                                  className="rotate-file-frame"
                                  style={{ transform: `rotate(${angle}deg)` }}
                                  src={`${previewUrls[index]}#page=1&view=FitH`}
                                  title={`${file.name} preview`}
                                />
                              </div>
                              <p>{file.name}</p>
                            </button>
                          )
                        })}
                      </div>
                    </div>

                    <aside className="rotate-sidebar">
                      <h2>Rotate PDF</h2>
                      <p>
                        {rotateStatus ||
                          'Use right or left controls to rotate, then click Rotate PDF to download.'}
                      </p>
                      <button type="button" className="rotate-reset-link" onClick={resetRotateAll}>
                        Reset all
                      </button>
                      <div className="rotate-controls-label">Rotation</div>
                      <div className="rotate-controls">
                        <button type="button" onClick={() => rotateFile(rotateSelectedIndex, 90)}>
                          <span>↻</span>
                          <span>RIGHT</span>
                        </button>
                        <button type="button" onClick={() => rotateFile(rotateSelectedIndex, -90)}>
                          <span>↺</span>
                          <span>LEFT</span>
                        </button>
                      </div>
                      <button
                        type="button"
                        className="rotate-main-btn"
                        onClick={handleRotateConvert}
                        disabled={rotateBusy}
                      >
                        {rotateBusy ? 'Rotating...' : 'Rotate PDF'}
                      </button>
                    </aside>
                  </section>
                ) : activeTool.slug === 'page-numbers' ? (
                  <section className="page-number-workspace-page" aria-label="Page numbers workspace">
                    <div className="page-number-left">
                      <button type="button" className="page-number-add-btn" onClick={openFilePicker} aria-label="Add PDFs">
                        +
                        <span>{selectedFiles.length}</span>
                      </button>

                      <div className="page-number-file-strip">
                        <select
                          value={pageNumFileIndex}
                          onChange={(event) => setPageNumFileIndex(Number(event.target.value))}
                        >
                          {selectedFiles.map((file, index) => (
                            <option key={`${file.name}-${file.lastModified}-${index}`} value={index}>
                              {file.name}
                            </option>
                          ))}
                        </select>
                        <button
                          type="button"
                          onClick={() => removeFileAt(pageNumFileIndex)}
                          aria-label="Remove selected file"
                        >
                          Delete
                        </button>
                      </div>

                      <div className="page-number-preview-wrap">
                        <div className="page-number-preview-card">
                          {selectedPageNumPreviewUrl ? (
                            <iframe
                              className="page-number-preview-frame"
                              src={`${selectedPageNumPreviewUrl}#page=1&view=FitH`}
                              title={`${selectedPageNumFile?.name || 'PDF'} preview`}
                            />
                          ) : (
                            <div className="page-number-preview-empty">PDF</div>
                          )}
                          <span className={`page-number-dot ${pageNumPosition}`} />
                        </div>
                      </div>
                    </div>

                    <aside className="page-number-sidebar">
                      <h2>Page Number options</h2>
                      <div className="page-number-sidebar-body">
                        {pageNumStatus && <p className="page-number-status">{pageNumStatus}</p>}

                        <div className="page-number-block">
                          <p>Page mode</p>
                          <div className="page-number-inline">
                            <button
                              type="button"
                              className={pageNumMode === 'single' ? 'active' : ''}
                              onClick={() => setPageNumMode('single')}
                            >
                              Single page
                            </button>
                            <button
                              type="button"
                              className={pageNumMode === 'facing' ? 'active' : ''}
                              onClick={() => setPageNumMode('facing')}
                            >
                              Facing pages
                            </button>
                          </div>
                        </div>

                        <div className="page-number-grid">
                          <div className="page-number-block">
                            <p>Position</p>
                            <div className="page-number-position-picker">
                              {pageNumPositions.flat().map((pos) => (
                                <button
                                  key={pos}
                                  type="button"
                                  className={pageNumPosition === pos ? 'active' : ''}
                                  onClick={() => setPageNumPosition(pos)}
                                  aria-label={`Set position ${pos}`}
                                />
                              ))}
                            </div>
                          </div>

                          <div className="page-number-block">
                            <p>Margin</p>
                            <select value={pageNumMargin} onChange={(event) => setPageNumMargin(event.target.value)}>
                              <option value="recommended">Recommended</option>
                              <option value="large">Large</option>
                              <option value="none">No margin</option>
                            </select>
                          </div>
                        </div>

                        <div className="page-number-block">
                          <p>Pages</p>
                          <label className="page-number-label">
                            First number
                            <input
                              type="number"
                              min="1"
                              value={pageNumFirst}
                              onChange={(event) => setPageNumFirst(Math.max(1, Number(event.target.value || 1)))}
                            />
                          </label>
                        </div>

                        <div className="page-number-block">
                          <p>Text</p>
                          <select value={pageNumTextMode} onChange={(event) => setPageNumTextMode(event.target.value)}>
                            <option value="number">Insert only page number</option>
                            <option value="full">Page X of Y</option>
                          </select>
                        </div>

                        <div className="page-number-block">
                          <p>Text format</p>
                          <div className="page-number-format-row">
                            <select value={pageNumFontFamily} onChange={(event) => setPageNumFontFamily(event.target.value)}>
                              <option value="helvetica">Helvetica</option>
                              <option value="times">Times</option>
                              <option value="courier">Courier</option>
                            </select>
                            <button
                              type="button"
                              className={pageNumBold ? 'active' : ''}
                              onClick={() => setPageNumBold((prev) => !prev)}
                            >
                              B
                            </button>
                            <button
                              type="button"
                              className={pageNumItalic ? 'active' : ''}
                              onClick={() => setPageNumItalic((prev) => !prev)}
                            >
                              I
                            </button>
                            <button
                              type="button"
                              className={pageNumUnderline ? 'active' : ''}
                              onClick={() => setPageNumUnderline((prev) => !prev)}
                            >
                              U
                            </button>
                          </div>
                        </div>

                <button
                  type="button"
                  className="page-number-main-btn"
                  onClick={handleAddPageNumbers}
                  disabled={pageNumBusy}
                >
                  {pageNumBusy ? 'Adding numbers...' : 'Add page numbers'}
                </button>
              </div>
            </aside>
          </section>
                ) : activeTool.slug === 'pdf-to-jpg' ? (
                  <section className="pdf-jpg-workspace-page" aria-label="PDF to JPG workspace">
                    <div className="pdf-jpg-left">
                      <button type="button" className="pdf-jpg-add-btn" onClick={openFilePicker} aria-label="Add PDFs">
                        <span>{selectedFiles.length}</span>+
                      </button>
                      <div className="pdf-jpg-files-grid">
                        {selectedFiles.map((file, index) => (
                          <article key={`${file.name}-${index}`} className="pdf-jpg-file-card">
                            <div className="pdf-jpg-file-preview">
                              <iframe
                                className="pdf-jpg-file-frame"
                                src={`${previewUrls[index]}#page=1&view=FitH`}
                                title={`Preview ${file.name}`}
                              />
                            </div>
                            <p>{file.name}</p>
                          </article>
                        ))}
                      </div>
                    </div>

                    <aside className="pdf-jpg-sidebar">
                      <h2>PDF to JPG options</h2>
                      <div className="pdf-jpg-option-row">
                        <button
                          type="button"
                          className={jpgExtractMode === 'page' ? 'active' : ''}
                          onClick={() => setJpgExtractMode('page')}
                        >
                          <strong>Page to JPG</strong>
                          <span>Convert every page of this PDF into JPG images.</span>
                        </button>
                        <button
                          type="button"
                          className={jpgExtractMode === 'extract' ? 'active' : ''}
                          onClick={() => setJpgExtractMode('extract')}
                        >
                          <strong>Extract images</strong>
                          <span>Export all pages as JPG files inside a ZIP.</span>
                        </button>
                      </div>

                      <div className="pdf-jpg-quality">
                        <p>Image quality</p>
                        <div>
                          <button
                            type="button"
                            className={jpgQualityMode === 'normal' ? 'active' : ''}
                            onClick={() => setJpgQualityMode('normal')}
                          >
                            Normal
                          </button>
                          <button
                            type="button"
                            className={jpgQualityMode === 'high' ? 'active' : ''}
                            onClick={() => setJpgQualityMode('high')}
                          >
                            High
                          </button>
                        </div>
                      </div>

                      <button
                        type="button"
                        className="pdf-jpg-main-btn"
                        onClick={onConvert}
                        disabled={isMerging}
                      >
                        {isMerging ? 'Converting...' : 'Convert to JPG'}
                      </button>
                    </aside>
                  </section>
                ) : activeTool.slug === 'compare-pdf' ? (
                  <section className="compare-workspace-page" aria-label="Compare PDF workspace">
                    <div className="compare-left">
                      <div className="compare-toolbar">
                        <button type="button" title="Move mode">✋</button>
                        <button
                          type="button"
                          className={compareScrollSync ? 'active' : ''}
                          onClick={() => setCompareScrollSync((prev) => !prev)}
                          title="Toggle scroll sync"
                        >
                          ⇅ Scroll sync
                        </button>
                        <button
                          type="button"
                          onClick={() => setCompareZoom((prev) => Math.max(50, prev - 10))}
                          title="Zoom out"
                        >
                          −
                        </button>
                        <button
                          type="button"
                          onClick={() => setCompareZoom((prev) => Math.min(200, prev + 10))}
                          title="Zoom in"
                        >
                          +
                        </button>
                        <button type="button" onClick={openFilePicker} title="Add more PDFs">
                          + Add PDF
                        </button>
                      </div>

                      {compareMode === 'semantic' ? (
                        <div className="compare-split">
                          <div className="compare-pane" ref={compareLeftScrollRef}>
                            <button
                              type="button"
                              className="compare-pane-close"
                              onClick={() => removeFileAt(compareLeftFileIndex)}
                              aria-label="Remove left file"
                            >
                              ×
                            </button>
                            <canvas ref={compareLeftCanvasRef} className="compare-canvas" />
                            <div className="compare-bottom-strip">
                              <span>{compareZoom}%</span>
                              <span>{compareLeftFile?.name || 'Left file'}</span>
                            </div>
                          </div>
                          <div className="compare-divider" />
                          <div className="compare-pane" ref={compareRightScrollRef}>
                            <button
                              type="button"
                              className="compare-pane-close"
                              onClick={() => removeFileAt(compareRightFileIndex)}
                              aria-label="Remove right file"
                            >
                              ×
                            </button>
                            <canvas ref={compareRightCanvasRef} className="compare-canvas" />
                            <div className="compare-bottom-strip">
                              <span>{compareZoom}%</span>
                              <span>{compareRightFile?.name || 'Right file'}</span>
                            </div>
                          </div>
                        </div>
                      ) : (
                        <div className="compare-overlay-wrap">
                          <div className="compare-pane compare-overlay-pane">
                            <canvas ref={compareOverlayCanvasRef} className="compare-canvas" />
                            <div className="compare-bottom-strip">
                              <span>{compareZoom}%</span>
                              <span>{compareLeftFile?.name || 'File A'} vs {compareRightFile?.name || 'File B'}</span>
                            </div>
                          </div>
                        </div>
                      )}
                    </div>

                    <aside className="compare-sidebar">
                      <h2>Compare PDF</h2>
                      <div className="compare-tabs">
                        <button
                          type="button"
                          className={compareMode === 'semantic' ? 'active' : ''}
                          onClick={() => setCompareMode('semantic')}
                        >
                          Semantic Text
                        </button>
                        <button
                          type="button"
                          className={compareMode === 'overlay' ? 'active' : ''}
                          onClick={() => setCompareMode('overlay')}
                        >
                          Content Overlay
                        </button>
                      </div>

                      <p className="compare-hint">
                        {compareMode === 'semantic'
                          ? 'Compare text changes between two PDFs.'
                          : 'Overlay both files and highlight visual differences.'}
                      </p>

                      <input
                        className="compare-search"
                        value={compareSearch}
                        onChange={(event) => setCompareSearch(event.target.value)}
                        placeholder="Search text"
                      />

                      <div className="compare-file-rows">
                        <div className="compare-file-row">
                          <span className="compare-file-name">File A</span>
                          <div className="compare-file-preview-row">
                            <iframe
                              className="compare-file-thumb"
                              src={compareLeftPreviewUrl ? `${compareLeftPreviewUrl}#page=1&view=FitH` : undefined}
                              title="File A preview"
                            />
                            <div className="compare-file-controls">
                              <select
                                value={compareLeftFileIndex}
                                onChange={(event) => setCompareFileForSide('left', event.target.value)}
                              >
                                {selectedFiles.map((file, idx) => (
                                  <option key={`left-file-${file.name}-${idx}`} value={idx}>
                                    {file.name}
                                  </option>
                                ))}
                              </select>
                              <select
                                value={compareLeftPage}
                                onChange={(event) => setCompareLeftPage(Number(event.target.value))}
                              >
                                {Array.from({ length: compareLeftPages }, (_, idx) => idx + 1).map((page) => (
                                  <option key={`left-page-${page}`} value={page}>
                                    Page {page}
                                  </option>
                                ))}
                              </select>
                            </div>
                          </div>
                        </div>
                        <div className="compare-file-row">
                          <span className="compare-file-name">File B</span>
                          <div className="compare-file-preview-row">
                            <iframe
                              className="compare-file-thumb"
                              src={compareRightPreviewUrl ? `${compareRightPreviewUrl}#page=1&view=FitH` : undefined}
                              title="File B preview"
                            />
                            <div className="compare-file-controls">
                              <select
                                value={compareRightFileIndex}
                                onChange={(event) => setCompareFileForSide('right', event.target.value)}
                              >
                                {selectedFiles.map((file, idx) => (
                                  <option key={`right-file-${file.name}-${idx}`} value={idx}>
                                    {file.name}
                                  </option>
                                ))}
                              </select>
                              <select
                                value={compareRightPage}
                                onChange={(event) => setCompareRightPage(Number(event.target.value))}
                              >
                                {Array.from({ length: compareRightPages }, (_, idx) => idx + 1).map((page) => (
                                  <option key={`right-page-${page}`} value={page}>
                                    Page {page}
                                  </option>
                                ))}
                              </select>
                            </div>
                          </div>
                        </div>
                      </div>

                      <h3>Change report ({filteredCompareReport.length})</h3>
                      <div className="compare-report-list">
                        {filteredCompareReport.slice(0, 80).map((item) => (
                          <article key={item.id} className={`compare-report-item ${item.type}`}>
                            <strong>Line {item.line}</strong>
                            <p>L: {item.left || '—'}</p>
                            <p>R: {item.right || '—'}</p>
                          </article>
                        ))}
                        {!filteredCompareReport.length && (
                          <p className="compare-empty">{compareStatus || (compareBusy ? 'Comparing...' : 'No changes')}</p>
                        )}
                      </div>

                      <button type="button" className="compare-main-btn" onClick={downloadCompareReport}>
                        Download report
                      </button>
                    </aside>
                  </section>
                ) : activeTool.slug === 'crop-pdf' ? (
                  <section className="crop-workspace-page" aria-label="Crop PDF workspace">
                    <div className="crop-left">
                      <div className="crop-canvas-wrap">
                        <div className="crop-canvas-stage">
                          <canvas
                            ref={cropCanvasRef}
                            className="crop-canvas"
                            onMouseDown={onCropMouseDown}
                            onMouseMove={onCropMouseMove}
                            onMouseUp={onCropMouseUp}
                            onMouseLeave={onCropMouseUp}
                          />
                          {cropRect && (
                            <div
                              className="crop-selection"
                              style={{
                                left: `${cropRect.x}px`,
                                top: `${cropRect.y}px`,
                                width: `${cropRect.w}px`,
                                height: `${cropRect.h}px`,
                              }}
                            >
                              <span className="crop-handle tl" />
                              <span className="crop-handle tr" />
                              <span className="crop-handle bl" />
                              <span className="crop-handle br" />
                            </div>
                          )}
                        </div>
                      </div>

                      <div className="crop-bottom-toolbar">
                        <button type="button" onClick={() => setCropZoom((prev) => Math.max(50, prev - 10))}>-</button>
                        <span>{cropZoom}%</span>
                        <button type="button" onClick={() => setCropZoom((prev) => Math.min(200, prev + 10))}>+</button>
                        <button type="button" onClick={() => setCropPage((prev) => Math.max(1, prev - 1))}>Prev</button>
                        <span>{cropPage} / {cropPageCount}</span>
                        <button type="button" onClick={() => setCropPage((prev) => Math.min(cropPageCount, prev + 1))}>Next</button>
                      </div>
                    </div>

                    <aside className="crop-sidebar">
                      <h2>Crop PDF</h2>
                      <p>
                        {cropStatus || 'Click and drag to select the area you want to keep. Drag inside to move and use corners to resize.'}
                      </p>
                      <button type="button" className="crop-reset-link" onClick={resetCrop}>
                        Reset all
                      </button>

                      <div className="crop-scope">
                        <span>Pages:</span>
                        <label>
                          <input
                            type="radio"
                            name="crop-scope"
                            checked={cropScope === 'all'}
                            onChange={() => setCropScope('all')}
                          />
                          All pages
                        </label>
                        <label>
                          <input
                            type="radio"
                            name="crop-scope"
                            checked={cropScope === 'current'}
                            onChange={() => setCropScope('current')}
                          />
                          Current page
                        </label>
                      </div>

                      <button
                        type="button"
                        className="crop-main-btn"
                        onClick={handleCropDownload}
                        disabled={cropBusy || !selectedCropFile || !cropRect}
                      >
                        {cropBusy ? 'Cropping...' : 'Crop PDF'}
                      </button>
                    </aside>
                  </section>
                ) : (
                  <section className="workspace-page" aria-label="Selected files workspace">
                    <div className="workspace-left">
                      <div className="workspace-files-grid">
                        {selectedFiles.map((file, index) => (
                          <article
                            key={`${file.name}-${file.lastModified}-${index}`}
                            className={`file-tile ${draggingIndex === index ? 'dragging' : ''} ${dragOverIndex === index ? 'drag-over' : ''}`}
                            draggable
                            onDragStart={(event) => onTileDragStart(event, index)}
                            onDragOver={(event) => onTileDragOver(event, index)}
                            onDrop={(event) => onTileDrop(event, index)}
                            onDragEnd={onTileDragEnd}
                          >
                            <div className="file-preview-box">
                              {activeTool?.slug === 'jpg-to-pdf' ? (
                                <img
                                  className="file-preview-frame"
                                  src={previewUrls[index]}
                                  alt={`Preview ${file.name}`}
                                />
                              ) : (
                                <>
                                  <iframe
                                    className="file-preview-frame"
                                    src={`${previewUrls[index]}#page=1&view=FitH`}
                                    title={`Preview ${file.name}`}
                                  />
                                  <span className="file-preview-fallback">PDF</span>
                                </>
                              )}
                            </div>
                            <p className="file-tile-name">{file.name}</p>
                            <p className="file-tile-size">{formatSize(file.size)}</p>
                            <button
                              type="button"
                              className="remove-file-btn"
                              onClick={() => removeFileAt(index)}
                              aria-label={`Remove ${file.name}`}
                            >
                              Remove
                            </button>
                          </article>
                        ))}
                      </div>
                      <button type="button" className="add-more-btn" onClick={openFilePicker}>
                        <span className="file-count-badge">{selectedFiles.length}</span>
                        +
                      </button>
                    </div>

                    <aside className="workspace-sidebar">
                      <h2>{activeTool.title}</h2>
                      {activeTool.slug === 'translate-pdf' ? (
                        <div className="translate-panel">
                          <label htmlFor="translate-source">Source language</label>
                          <select
                            id="translate-source"
                            value={translateSourceLang}
                            onChange={(event) => setTranslateSourceLang(event.target.value)}
                          >
                            {translateLanguages.map((lang) => (
                              <option key={`src-${lang.code}`} value={lang.code}>
                                {lang.label}
                              </option>
                            ))}
                          </select>

                          <label htmlFor="translate-target">Target language</label>
                          <select
                            id="translate-target"
                            value={translateTargetLang}
                            onChange={(event) => setTranslateTargetLang(event.target.value)}
                            disabled={!translateTargetOptions.length}
                          >
                            {translateTargetOptions.map((lang) => (
                              <option key={`target-${lang.code}`} value={lang.code}>
                                {lang.label}
                              </option>
                            ))}
                          </select>

                          <p className="translate-status">{translateStatus || 'Ready to translate'}</p>
                          <textarea
                            readOnly
                            className="translate-output"
                            value={translatedText}
                            placeholder="Translated text will appear here."
                          />
                          <button
                            type="button"
                            className="translate-download-btn"
                            onClick={downloadTranslatedResult}
                            disabled={!translatedText.trim()}
                          >
                            Download translation
                          </button>
                        </div>
                      ) : (
                        <div className="workspace-empty" />
                      )}
                      <button type="button" className="convert-main-btn" onClick={onConvert} disabled={isMerging}>
                        {isMerging ? 'Processing...' : convertLabel}
                      </button>
                    </aside>
                  </section>
                )
              )}

              {activeTool?.slug !== 'rotate-pdf' && activeTool?.slug !== 'organize-pdf' && activeTool?.slug !== 'split-pdf' && activeTool?.slug !== 'edit-pdf' && activeTool?.slug !== 'page-numbers' && activeTool?.slug !== 'compare-pdf' && activeTool?.slug !== 'crop-pdf' && activeTool?.slug !== 'pdf-to-jpg' && activeTool?.slug !== 'scan-to-pdf' && <p className="status-text">@urstrulypdf</p>}
            </div>
          </section>
        ) : (
          <>
            <section className="hero-block">
              <p className="eyebrow">Online PDF tools</p>
              <h1>Everything you need to work with PDFs in one place</h1>
            </section>

            <section className="filters" aria-label="Tool categories">
              {categoryRoutes.map((category) => (
                <a
                  key={category.path}
                  href={category.path}
                  onClick={(event) => navigate(event, category.path)}
                  className={`chip ${routeState.chipPath === category.path ? 'active' : ''}`}
                >
                  {category.label}
                </a>
              ))}
            </section>

            <section className="route-heading" aria-live="polite">
              <p>{routeState.title}</p>
            </section>

            <section id="tools" className="cards-grid" aria-label="All PDF tools">
              {visibleTools.map((tool) => (
                <a
                  key={tool.title}
                  href={`/tool/${tool.slug}`}
                  onClick={(event) => navigate(event, `/tool/${tool.slug}`)}
                  className="tool-card-link"
                >
                  <article className="tool-card">
                    {tool.isNew && <span className="new-badge">New!</span>}
                    <div className="tool-icon" aria-hidden="true">
                      <ToolGlyph type={tool.icon} />
                    </div>
                    <h3>{tool.title}</h3>
                    <p>{tool.description}</p>
                  </article>
                </a>
              ))}
            </section>
          </>
        )}
      </main>
    </div>
  )
}

export default App
