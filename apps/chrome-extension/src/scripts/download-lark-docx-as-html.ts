import { fileSave, supported } from 'browser-fs-access'
import { fs } from '@zip.js/zip.js'
import normalizeFileName from 'filenamify/browser'
import { legacyFileSave } from '../common/legacy'
import {
  WindowMessageType,
  type FetchAssetResponse,
  type WindowFetchAssetResponse,
} from '../common/message'
import { DownloadMethod, SettingKey, getSettings } from '../common/settings'

const KSTACK_ARTICLES_HOST = 'kstack.corp.kuaishou.com'
const ASSET_FETCH_TIMEOUT = 8000
const MAX_FILENAME_LENGTH = 100
const VODKA_SCROLL_SETTLE_TIME = 250

const normalizeText = (value: string): string =>
  value
    .replace(/\u00a0/g, ' ')
    .replace(/[ \t]+/g, ' ')
    .trim()

const firstNonEmpty = (...values: (string | null | undefined)[]): string =>
  values.map(value => normalizeText(value ?? '')).find(Boolean) ?? ''

const cleanDocumentTitle = (value: string): string =>
  normalizeText(
    value
      .replace(/\s+-\s+云文档$/, '')
      .replace(/\s+-\s+轻雀文档$/, '')
      .replace(/\s+-\s+文章\s+-\s+KStack$/, ''),
  )

const normalizeDocumentFileName = (value: string): string =>
  normalizeFileName(
    (cleanDocumentTitle(value) || 'doc')
      .replace(/^\d+(?=\p{Script=Han})/u, '')
      .slice(0, MAX_FILENAME_LENGTH),
  )

const escapeHtml = (value: string): string =>
  value.replace(/[&<>"']/g, char => {
    switch (char) {
      case '&':
        return '&amp;'
      case '<':
        return '&lt;'
      case '>':
        return '&gt;'
      case '"':
        return '&quot;'
      default:
        return '&#39;'
    }
  })

const getDocumentLanguage = (): string =>
  firstNonEmpty(document.documentElement.lang, navigator.language, 'zh-CN')

const isHiddenElement = (element: HTMLElement): boolean => {
  const style = getComputedStyle(element)
  return (
    style.display === 'none' ||
    style.visibility === 'hidden' ||
    element.getAttribute('aria-hidden') === 'true'
  )
}

const selectDocumentRoot = (): HTMLElement => {
  const vodkaRoot = document.querySelector<HTMLElement>(
    '#vodka-paginateddocumentplugin, .vodka-page-content-wrapper',
  )
  if (vodkaRoot) return vodkaRoot

  if (location.hostname === KSTACK_ARTICLES_HOST) {
    const kstackRoot = document.querySelector<HTMLElement>(
      '.ck-content, .ArticleContent_wrapper__vmEAq',
    )
    if (kstackRoot && !isHiddenElement(kstackRoot)) return kstackRoot
  }

  const selectors = [
    'article',
    'main',
    '[role="main"]',
    '[contenteditable="true"]',
    '[class*="editor" i]',
    '[class*="document" i]',
    '[class*="doc" i]',
    '[class*="reader" i]',
    '[class*="content" i]',
  ]

  const candidates = selectors
    .flatMap(selector =>
      Array.from(document.querySelectorAll<HTMLElement>(selector)),
    )
    .filter(element => !isHiddenElement(element))

  return (
    candidates
      .map(element => ({
        element,
        textLength: normalizeText(element.innerText).length,
      }))
      .filter(({ textLength }) => textLength > 0)
      .sort((a, b) => b.textLength - a.textLength)
      .at(0)?.element ?? document.body
  )
}

const requestAssetFetchViaExtension = async (
  src: string,
): Promise<FetchAssetResponse | null> => {
  const id = `${Date.now().toFixed()}-${Math.random().toString(36).slice(2)}`

  return await new Promise(resolve => {
    const timeout = window.setTimeout(() => {
      window.removeEventListener('message', onMessage)
      resolve(null)
    }, ASSET_FETCH_TIMEOUT)

    const onMessage = (event: MessageEvent<unknown>) => {
      if (event.source !== window || event.origin !== location.origin) return

      const data = event.data as Partial<WindowFetchAssetResponse>
      if (
        data.type !== WindowMessageType.FetchAssetResponse ||
        data.id !== id
      ) {
        return
      }

      window.clearTimeout(timeout)
      window.removeEventListener('message', onMessage)
      resolve(data.response ?? null)
    }

    window.addEventListener('message', onMessage)
    window.postMessage(
      {
        type: WindowMessageType.FetchAssetRequest,
        id,
        src,
      },
      location.origin,
    )
  })
}

const dataUrlToBlob = async (dataUrl: string): Promise<Blob> =>
  await (await fetch(dataUrl)).blob()

const fetchImageBlob = async (
  src: string,
): Promise<{ blob: Blob; contentType: string | null } | null> => {
  try {
    const response = await fetch(src, { credentials: 'include' })
    if (response.ok) {
      const contentType = response.headers.get('content-type')
      const blob = await response.blob()

      return {
        blob,
        contentType: blob.type || contentType,
      }
    }
  } catch {
    // Some document CDNs allow rendering images but block direct page fetches.
  }

  const response = await requestAssetFetchViaExtension(src)
  if (!response?.ok) return null

  const blob = await dataUrlToBlob(response.dataUrl)
  return {
    blob,
    contentType: blob.type || response.contentType,
  }
}

const extensionFromContentType = (
  contentType: string | null,
): string | null => {
  if (!contentType) return null
  if (contentType.includes('webp')) return '.webp'
  if (contentType.includes('png')) return '.png'
  if (contentType.includes('jpeg') || contentType.includes('jpg')) return '.jpg'
  if (contentType.includes('gif')) return '.gif'
  if (contentType.includes('svg')) return '.svg'
  return null
}

const extensionFromUrl = (src: string): string => {
  try {
    const pathname = new URL(src, location.href).pathname
    const extension = pathname
      .split('/')
      .pop()
      ?.match(/\.[a-z0-9]{1,8}$/i)?.[0]
    return extension ?? '.png'
  } catch {
    return '.png'
  }
}

const replacePathExtension = (path: string, extension: string): string =>
  /\.[a-z0-9]{1,8}$/i.test(path)
    ? path.replace(/\.[a-z0-9]{1,8}$/i, extension)
    : `${path}${extension}`

const inlineStyleProperties = [
  'background-color',
  'border-bottom-color',
  'border-bottom-style',
  'border-bottom-width',
  'border-collapse',
  'border-left-color',
  'border-left-style',
  'border-left-width',
  'border-right-color',
  'border-right-style',
  'border-right-width',
  'border-top-color',
  'border-top-style',
  'border-top-width',
  'box-sizing',
  'color',
  'display',
  'font-family',
  'font-size',
  'font-style',
  'font-weight',
  'height',
  'line-height',
  'list-style-position',
  'list-style-type',
  'margin-bottom',
  'margin-left',
  'margin-right',
  'margin-top',
  'max-width',
  'min-height',
  'min-width',
  'padding-bottom',
  'padding-left',
  'padding-right',
  'padding-top',
  'text-align',
  'text-decoration',
  'vertical-align',
  'white-space',
  'width',
] as const

const copyInlineStyles = (source: HTMLElement, target: HTMLElement): void => {
  const style = getComputedStyle(source)

  inlineStyleProperties.forEach(property => {
    const value = style.getPropertyValue(property)
    if (value) target.style.setProperty(property, value)
  })
}

const absolutizeAttribute = (
  element: Element,
  attribute: 'href' | 'src',
): void => {
  const value = element.getAttribute(attribute)
  if (!value || value.startsWith('data:') || value.startsWith('#')) return

  element.setAttribute(attribute, new URL(value, location.href).toString())
}

interface HtmlAsset {
  filename: string
  src: string
}

interface HtmlExportContext {
  assets: HtmlAsset[]
  srcToFilename: Map<string, string>
}

const createHtmlExportContext = (): HtmlExportContext => ({
  assets: [],
  srcToFilename: new Map(),
})

const registerImageAsset = (
  src: string,
  context: HtmlExportContext,
): string => {
  const absoluteSrc = new URL(src, location.href).toString()
  const cached = context.srcToFilename.get(absoluteSrc)
  if (cached) return cached

  const filename = `images/image-${(context.assets.length + 1)
    .toFixed()
    .padStart(3, '0')}${extensionFromUrl(absoluteSrc)}`
  context.srcToFilename.set(absoluteSrc, filename)
  context.assets.push({
    filename,
    src: absoluteSrc,
  })

  return filename
}

interface CloneNodeForExportOptions {
  context: HtmlExportContext
}

const cloneNodeForExport = async (
  node: Node,
  options: CloneNodeForExportOptions,
): Promise<Node | null> => {
  const { context } = options

  if (node.nodeType === Node.TEXT_NODE) {
    return document.createTextNode(node.textContent ?? '')
  }

  if (!(node instanceof HTMLElement) || isHiddenElement(node)) return null

  const tagName = node.tagName.toLowerCase()
  if (['script', 'style', 'noscript'].includes(tagName)) return null

  if (node instanceof HTMLCanvasElement) {
    try {
      const image = document.createElement('img')
      image.src = node.toDataURL('image/png')
      image.alt = 'canvas'
      copyInlineStyles(node, image)
      return image
    } catch {
      return null
    }
  }

  const clone = node.cloneNode(false) as HTMLElement
  copyInlineStyles(node, clone)
  clone.removeAttribute('contenteditable')
  clone.removeAttribute('spellcheck')
  clone.removeAttribute('data-CDC-button-type')

  if (clone instanceof HTMLAnchorElement) {
    absolutizeAttribute(clone, 'href')
  }

  if (clone instanceof HTMLIFrameElement) {
    absolutizeAttribute(clone, 'src')
  }

  if (node instanceof HTMLImageElement && clone instanceof HTMLImageElement) {
    const rawSrc = node.getAttribute('src')
    const src = node.currentSrc || node.src
    if (rawSrc?.startsWith('images/')) {
      clone.setAttribute('src', rawSrc)
    } else if (src && !src.startsWith('data:')) {
      clone.src = registerImageAsset(src, context)
    }

    clone.alt = node.alt
    clone.removeAttribute('srcset')
  }

  for (const child of Array.from(node.childNodes)) {
    const clonedChild = await cloneNodeForExport(child, options)
    if (clonedChild) clone.appendChild(clonedChild)
  }

  return clone
}

interface VodkaExportItem {
  key: string
  y: number
  x: number
  node: Node
}

const getElementPagePosition = (
  element: HTMLElement,
  scroller: HTMLElement,
): { y: number; x: number } => {
  const rect = element.getBoundingClientRect()
  const scrollerRect = scroller.getBoundingClientRect()

  return {
    y: Math.round(scroller.scrollTop + rect.top - scrollerRect.top),
    x: Math.round(rect.left - scrollerRect.left),
  }
}

const wait = async (timeout: number): Promise<void> => {
  await new Promise(resolve => {
    window.setTimeout(resolve, timeout)
  })
}

const addVodkaExportItem = async (
  items: Map<string, VodkaExportItem>,
  element: HTMLElement,
  scroller: HTMLElement,
  context: HtmlExportContext,
  type: string,
): Promise<void> => {
  if (isHiddenElement(element)) return

  const text = normalizeText(element.innerText)
  const imageSrc =
    element instanceof HTMLImageElement ? element.currentSrc || element.src : ''
  if (!text && !imageSrc) return

  const { y, x } = getElementPagePosition(element, scroller)
  const key =
    type === 'table'
      ? `${type}:${text}`
      : type === 'image'
        ? `${type}:${imageSrc}`
        : `${type}:${y.toFixed()}:${x.toFixed()}:${text || imageSrc}`
  if (items.has(key)) return

  const node = await cloneNodeForExport(element, { context })
  if (!node) return

  items.set(key, {
    key,
    y,
    x,
    node,
  })
}

const collectVodkaDocumentRoot = async (
  context: HtmlExportContext,
): Promise<HTMLElement | null> => {
  const scroller = document.querySelector<HTMLElement>('#vodka-appview-editor')
  const root = document.querySelector<HTMLElement>(
    '#vodka-paginateddocumentplugin, .vodka-page-content-wrapper',
  )
  if (!scroller || !root) return null

  const initialScrollTop = scroller.scrollTop
  const items = new Map<string, VodkaExportItem>()
  const step = Math.max(300, Math.floor(scroller.clientHeight * 0.6))
  const maxScrollTop = Math.max(
    0,
    scroller.scrollHeight - scroller.clientHeight,
  )

  for (let top = 0; top <= maxScrollTop + step; top += step) {
    scroller.scrollTop = Math.min(top, maxScrollTop)
    await wait(VODKA_SCROLL_SETTLE_TIME)

    await Promise.all(
      Array.from(root.querySelectorAll<HTMLElement>('table')).map(table =>
        addVodkaExportItem(items, table, scroller, context, 'table'),
      ),
    )

    await Promise.all(
      Array.from(root.querySelectorAll<HTMLImageElement>('img'))
        .filter(image => !image.closest('table'))
        .map(image =>
          addVodkaExportItem(items, image, scroller, context, 'image'),
        ),
    )

    await Promise.all(
      Array.from(root.querySelectorAll<HTMLElement>('.vodka-lineview-content'))
        .filter(line => !line.closest('table'))
        .map(line =>
          addVodkaExportItem(items, line, scroller, context, 'line'),
        ),
    )
  }

  scroller.scrollTop = initialScrollTop

  const output = document.createElement('div')
  output.className = 'cdc-vodka-export'
  Array.from(items.values())
    .sort((a, b) => a.y - b.y || a.x - b.x || a.key.localeCompare(b.key))
    .forEach(item => {
      const wrapper = document.createElement('div')
      wrapper.className = 'cdc-vodka-export-item'
      wrapper.appendChild(item.node)
      output.appendChild(wrapper)
    })

  return output
}

const createHtmlDocument = async (
  root: HTMLElement,
  displayTitle: string,
  context: HtmlExportContext,
): Promise<string> => {
  const clonedRoot = await cloneNodeForExport(root, { context })
  const body = clonedRoot instanceof HTMLElement ? clonedRoot.outerHTML : ''
  const rootText = normalizeText(root.innerText)
  const titleMarkup = rootText.startsWith(displayTitle)
    ? ''
    : `<h1>${escapeHtml(displayTitle)}</h1>`

  return `<!doctype html>
<html lang="${escapeHtml(getDocumentLanguage())}">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>${escapeHtml(displayTitle)}</title>
  <style>
    :root { color-scheme: light; }
    body {
      margin: 0;
      background: #f6f7f9;
      color: #1f2329;
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
      line-height: 1.65;
    }
    .cdc-html-export {
      max-width: 980px;
      margin: 0 auto;
      padding: 32px 24px 48px;
      background: #fff;
      min-height: 100vh;
    }
    table { border-collapse: collapse; max-width: 100%; }
    th, td { border: 1px solid #d0d7de; padding: 6px 8px; vertical-align: top; }
    img, video, canvas { max-width: 100%; height: auto; }
    pre { overflow: auto; }
    .cdc-vodka-export-item { margin: 0 0 10px; }
  </style>
</head>
<body>
  <main class="cdc-html-export">
    ${titleMarkup}
    ${body}
  </main>
</body>
</html>
`
}

const saveHtmlZip = async (
  html: string,
  title: string,
  assets: HtmlAsset[],
): Promise<void> => {
  const zipFs = new fs.FS()
  let htmlContent = html

  await Promise.allSettled(
    assets.map(async asset => {
      const result = await fetchImageBlob(asset.src)

      if (!result) {
        htmlContent = htmlContent.replaceAll(asset.filename, asset.src)
        return
      }

      const extension =
        extensionFromContentType(result.contentType) ??
        extensionFromContentType(result.blob.type) ??
        extensionFromUrl(asset.src)
      const filename = replacePathExtension(asset.filename, extension)

      zipFs.addBlob(`${title}/${filename}`, result.blob)

      if (filename !== asset.filename) {
        htmlContent = htmlContent.replaceAll(asset.filename, filename)
      }
    }),
  )

  assets.forEach(asset => {
    if (htmlContent.includes(asset.filename)) {
      htmlContent = htmlContent.replaceAll(asset.filename, asset.src)
    }
  })

  zipFs.addText(`${title}/${title}.html`, htmlContent)
  const blob = await zipFs.exportBlob()
  const settings = await getSettings([SettingKey.DownloadMethod])

  if (
    settings[SettingKey.DownloadMethod] === DownloadMethod.ShowSaveFilePicker &&
    supported &&
    navigator.userActivation.isActive
  ) {
    await fileSave(blob, {
      fileName: `${title}.zip`,
      extensions: ['.zip'],
    })
    return
  }

  legacyFileSave(blob, { fileName: `${title}.zip` })
}

const main = async (): Promise<void> => {
  const context = createHtmlExportContext()
  const root = (await collectVodkaDocumentRoot(context)) ?? selectDocumentRoot()
  const displayTitle = firstNonEmpty(
    cleanDocumentTitle(
      firstNonEmpty(document.title, root.querySelector('h1')?.textContent),
    ),
    'doc',
  )
  const title = normalizeDocumentFileName(displayTitle)
  const html = await createHtmlDocument(root, displayTitle, context)

  await saveHtmlZip(html, title, context.assets)
}

main().catch((error: unknown) => {
  console.error(error)
  alert(`Unknown error during download: ${String(error)}`)
})
