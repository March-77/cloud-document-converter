import i18next from 'i18next'
import { Toast, Docx, docx, type mdast } from '@dolphin/lark'
import { Minute, OneHundred, Second, waitFor } from '@dolphin/common'
import { fileSave, supported } from 'browser-fs-access'
import { fs } from '@zip.js/zip.js'
import normalizeFileName from 'filenamify/browser'
import { cluster } from 'radash'
import { CommonTranslationKey, en, Namespace, zh } from '../common/i18n'
import {
  WindowMessageType,
  type FetchAssetResponse,
  type WindowFetchAssetResponse,
} from '../common/message'
import { confirm } from '../common/notification'
import { legacyFileSave } from '../common/legacy'
import { reportBug } from '../common/issue'
import {
  transformMentionUsers,
  UniqueFileName,
  withSignal,
  transformTableBySettings,
} from '../common/utils'
import { getSettings, Grid } from '../common/settings'
import { DownloadMethod, SettingKey } from '@/common/settings'
import {
  codeMirrorToMarkdown,
  fencedCodeBlock,
  htmlTableToMarkdown,
} from './vodka-markdown'

const uniqueFileName = new UniqueFileName()

const DOWNLOAD_ABORTED = 'Download aborted'

const enum TranslationKey {
  CONTENT_LOADING = 'content_loading',
  UNKNOWN_ERROR = 'unknown_error',
  NOT_SUPPORT = 'not_support',
  NOT_SUPPORT_DOC_1_0 = 'not_support_doc_1_0',
  DOWNLOADING_FILE = 'downloading_file',
  FAILED_TO_DOWNLOAD = 'failed_to_download',
  DOWNLOAD_PROGRESS = 'download_progress',
  DOWNLOAD_COMPLETE = 'download_complete',
  STILL_SAVING = 'still_saving',
  IMAGE = 'image',
  FILE = 'file',
  CANCEL = 'cancel',
  SCROLL_DOCUMENT = 'scroll_document',
}

enum ToastKey {
  DOWNLOADING = 'downloading',
  REPORT_BUG = 'report_bug',
}

i18next
  .init({
    lng: docx.language,
    resources: {
      en: {
        translation: {
          [TranslationKey.CONTENT_LOADING]:
            'Part of the content is still loading and cannot be downloaded at the moment. Please wait for loading to complete and retry',
          [TranslationKey.UNKNOWN_ERROR]: 'Unknown error during download',
          [TranslationKey.NOT_SUPPORT]:
            'This is not a lark document page and cannot be downloaded as Markdown',
          [TranslationKey.NOT_SUPPORT_DOC_1_0]:
            'This is a old version lark document page and cannot be downloaded as Markdown',
          [TranslationKey.DOWNLOADING_FILE]:
            'Download {{name}} in: {{progress}}% (please do not refresh or close the page)',
          [TranslationKey.FAILED_TO_DOWNLOAD]: 'Failed to download {{name}}',
          [TranslationKey.STILL_SAVING]:
            'Still saving (please do not refresh or close the page)',
          [TranslationKey.DOWNLOAD_PROGRESS]:
            '{{name}} download progress: {{progress}} %',
          [TranslationKey.DOWNLOAD_COMPLETE]: 'Download complete',
          [TranslationKey.IMAGE]: 'Image',
          [TranslationKey.FILE]: 'File',
          [TranslationKey.CANCEL]: 'Cancel',
          [TranslationKey.SCROLL_DOCUMENT]: 'Scrolling to load document',
        },
        ...en,
      },
      zh: {
        translation: {
          [TranslationKey.CONTENT_LOADING]:
            '部分内容仍在加载中，暂时无法下载。请等待加载完成后重试',
          [TranslationKey.UNKNOWN_ERROR]: '下载过程中出现未知错误',
          [TranslationKey.NOT_SUPPORT]:
            '这不是一个飞书文档页面，无法下载为 Markdown',
          [TranslationKey.NOT_SUPPORT_DOC_1_0]:
            '这是一个旧版飞书文档页面，无法下载为 Markdown',
          [TranslationKey.DOWNLOADING_FILE]:
            '下载 {{name}} 中：{{progress}}%（请不要刷新或关闭页面）',
          [TranslationKey.FAILED_TO_DOWNLOAD]: '下载 {{name}} 失败',
          [TranslationKey.STILL_SAVING]: '仍在保存中（请不要刷新或关闭页面）',
          [TranslationKey.DOWNLOAD_PROGRESS]: '{{name}}下载进度：{{progress}}%',
          [TranslationKey.DOWNLOAD_COMPLETE]: '下载完成',
          [TranslationKey.IMAGE]: '图片',
          [TranslationKey.FILE]: '文件',
          [TranslationKey.CANCEL]: '取消',
          [TranslationKey.SCROLL_DOCUMENT]: '滚动中，以便加载文档',
        },
        ...zh,
      },
    },
  })
  .catch(console.error)

interface ProgressOptions {
  onProgress?: (progress: number) => void
  onComplete?: () => void
}

async function toBlob(
  response: Response,
  options: ProgressOptions = {},
): Promise<Blob> {
  if (!response.ok) {
    throw new Error(`HTTP error! status: ${response.status.toFixed()}`)
  }

  if (!response.body) {
    throw new Error('This request has no response body.')
  }

  const { onProgress, onComplete } = options

  const reader = response.body.getReader()
  const contentLength = parseInt(
    response.headers.get('Content-Length') ?? '0',
    10,
  )

  let receivedLength = 0
  const chunks = []

  let _done = false
  while (!_done) {
    const { done, value } = await reader.read()

    _done = done

    if (done) {
      onComplete?.()

      break
    }

    chunks.push(value)
    receivedLength += value.length

    onProgress?.(receivedLength / contentLength)
  }

  const blob = new Blob(chunks)

  return blob
}

const downloadImage = async (
  image: mdast.Image,
  options: {
    signal?: AbortSignal
    useUUID?: boolean
    markdownFileName?: string
  } = {},
): Promise<DownloadResult | null> => {
  if (!image.data) return null

  const { signal, useUUID = false, markdownFileName = '' } = options

  const { name: originName, fetchSources, fetchBlob } = image.data

  const result = await withSignal(
    async isAborted => {
      try {
        // whiteboard
        if (fetchBlob) {
          if (isAborted()) {
            return null
          }

          const content = await fetchBlob()
          if (!content) return null

          const baseName = markdownFileName
            ? `${markdownFileName}-diagram.png`
            : 'diagram.png'
          const name = useUUID
            ? uniqueFileName.generateWithUUID(baseName)
            : uniqueFileName.generate(baseName)
          const filename = `images/${name}`

          image.url = filename

          return {
            filename,
            content,
          }
        }

        // image
        if (originName && fetchSources) {
          if (isAborted()) {
            return null
          }
          const sources = await fetchSources()
          if (!sources) return null

          const baseName = markdownFileName
            ? `${markdownFileName}-${originName}`
            : originName
          const name = useUUID
            ? uniqueFileName.generateWithUUID(baseName)
            : uniqueFileName.generate(baseName)
          const filename = `images/${name}`

          const { src } = sources
          if (isAborted()) {
            return null
          }
          const response = await fetch(src, {
            signal,
          })

          try {
            if (isAborted()) {
              return null
            }
            const blob = await toBlob(response, {
              onProgress: progress => {
                if (isAborted()) {
                  Toast.remove(filename)

                  return
                }

                Toast.loading({
                  content: i18next.t(TranslationKey.DOWNLOADING_FILE, {
                    name,
                    progress: Math.floor(progress * OneHundred),
                  }),
                  keepAlive: true,
                  key: filename,
                })
              },
            })

            image.url = filename

            return {
              filename,
              content: blob,
            }
          } finally {
            Toast.remove(filename)
          }
        }

        return null
      } catch (error) {
        const isAbortError =
          isAborted() ||
          (error instanceof DOMException && error.name === 'AbortError')

        if (!isAbortError) {
          Toast.error({
            content: i18next.t(TranslationKey.FAILED_TO_DOWNLOAD, {
              name: originName,
            }),
            actionText: i18next.t(CommonTranslationKey.CONFIRM_REPORT_BUG, {
              ns: Namespace.COMMON,
            }),
            onActionClick: () => {
              reportBug(error)
            },
          })
        }

        return null
      }
    },
    { signal },
  )

  return result
}

const downloadFile = async (
  file: mdast.Link,
  options: {
    signal?: AbortSignal
    useUUID?: boolean
    markdownFileName?: string
  } = {},
): Promise<DownloadResult | null> => {
  if (!file.data?.name || !file.data.fetchFile) return null

  const { signal, useUUID = false, markdownFileName = '' } = options

  const { name, fetchFile } = file.data

  let controller = new AbortController()

  const cancel = () => {
    controller.abort()
  }

  const result = await withSignal(
    async () => {
      try {
        const baseName = markdownFileName ? `${markdownFileName}-${name}` : name
        const filename = `files/${
          useUUID
            ? uniqueFileName.generateWithUUID(baseName)
            : uniqueFileName.generate(baseName)
        }`

        const response = await fetchFile({ signal: controller.signal })
        try {
          const blob = await toBlob(response, {
            onProgress: progress => {
              Toast.loading({
                content: i18next.t(TranslationKey.DOWNLOADING_FILE, {
                  name,
                  progress: Math.floor(progress * OneHundred),
                }),
                keepAlive: true,
                key: filename,
                actionText: i18next.t(TranslationKey.CANCEL),
                onActionClick: cancel,
              })
            },
          })

          file.url = filename

          return {
            filename,
            content: blob,
          }
        } finally {
          Toast.remove(filename)
        }
      } catch (error) {
        if (error instanceof DOMException && error.name === 'AbortError') {
          return null
        }

        Toast.error({
          content: i18next.t(TranslationKey.FAILED_TO_DOWNLOAD, {
            name,
          }),
          actionText: i18next.t(CommonTranslationKey.CONFIRM_REPORT_BUG, {
            ns: Namespace.COMMON,
          }),
          onActionClick: () => {
            reportBug(error)
          },
        })

        return null
      }
    },
    { signal, onAbort: cancel },
  )

  // @ts-expect-error remove reference
  controller = null

  return result
}

interface DownloadResult {
  filename: string
  content: Blob
}

interface GenericImage {
  alt: string
  filename: string
  src: string
  y?: number
}

const SUPPORTED_GENERIC_DOCUMENT_HOSTS = new Set([
  'docs.corp.kuaishou.com',
  'docs.qingque.cn',
  'kstack.corp.kuaishou.com',
])

const KSTACK_ARTICLES_HOST = 'kstack.corp.kuaishou.com'

const isSupportedGenericDocument = (): boolean =>
  SUPPORTED_GENERIC_DOCUMENT_HOSTS.has(location.hostname)

const escapeMarkdown = (value: string): string =>
  value.replace(/[\\`*_{}[\]()#+\-.!|>]/g, '\\$&')

const normalizeText = (value: string): string =>
  value
    .replace(/\u00a0/g, ' ')
    .replace(/[ \t]+/g, ' ')
    .trim()

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
      .slice(0, OneHundred),
  )

const getVisibleText = (node: Node): string =>
  Array.from(node.childNodes)
    .map(child =>
      child.nodeType === Node.TEXT_NODE
        ? (child.textContent ?? '')
        : child instanceof HTMLElement && !isHiddenElement(child)
          ? getVisibleText(child)
          : '',
    )
    .join('')

const isHiddenElement = (element: HTMLElement): boolean => {
  const style = getComputedStyle(element)
  return (
    style.display === 'none' ||
    style.visibility === 'hidden' ||
    element.getAttribute('aria-hidden') === 'true'
  )
}

const selectGenericDocumentRoot = (): HTMLElement => {
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
    .map(selector =>
      Array.from(document.querySelectorAll<HTMLElement>(selector)),
    )
    .flat(1)
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

const normalizeVodkaLineText = (value: string): string => {
  const lines = value.split('\n').map(normalizeText).filter(Boolean)

  return lines.filter((line, index) => line !== lines[index - 1]).join('\n')
}

interface VodkaTextRecord {
  text: string
  y: number
  x: number
}

const extensionFromContentType = (
  contentType: string | null,
): string | null => {
  if (!contentType) return null
  if (contentType.includes('webp')) return '.webp'
  if (contentType.includes('png')) return '.png'
  if (contentType.includes('jpeg') || contentType.includes('jpg')) return '.jpg'
  if (contentType.includes('gif')) return '.gif'
  return null
}

const requestAssetFetchViaExtension = async (
  src: string,
): Promise<FetchAssetResponse | null> => {
  const id = `${Date.now().toFixed()}-${Math.random().toString(36).slice(2)}`

  return await new Promise(resolve => {
    const timeout = window.setTimeout(() => {
      window.removeEventListener('message', onMessage)
      resolve(null)
    }, 30 * Second)

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

const fetchGenericImageBlob = async (
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
    // Some article CDNs allow <img> but block page fetch with CORS.
  }

  const response = await requestAssetFetchViaExtension(src)
  if (!response?.ok) return null

  const blob = await dataUrlToBlob(response.dataUrl)
  return {
    blob,
    contentType: blob.type || response.contentType,
  }
}

const replacePathExtension = (path: string, extension: string): string =>
  /\.[a-z0-9]{1,8}$/i.test(path)
    ? path.replace(/\.[a-z0-9]{1,8}$/i, extension)
    : `${path}${extension}`

const collectVodkaDocument = async (): Promise<{
  markdown: string
  images: GenericImage[]
  title: string
} | null> => {
  const scroller = document.querySelector<HTMLElement>('#vodka-appview-editor')
  const root = document.querySelector<HTMLElement>(
    '#vodka-paginateddocumentplugin, .vodka-page-content-wrapper',
  )
  if (!scroller || !root) return null

  const initialScrollTop = scroller.scrollTop
  const records: VodkaTextRecord[] = []
  const images = new Map<string, GenericImage>()
  const step = Math.max(300, Math.floor(scroller.clientHeight * 0.45))
  const maxScrollTop = Math.max(
    0,
    scroller.scrollHeight - scroller.clientHeight,
  )

  const collectVisible = () => {
    const scrollTop = scroller.scrollTop

    Array.from(root.querySelectorAll<HTMLElement>('table'))
      .filter(table => !table.closest('table table'))
      .forEach(table => {
        const text = htmlTableToMarkdown(table)
        if (!text) return

        const rect = table.getBoundingClientRect()
        records.push({
          text,
          y: Math.round(scrollTop + rect.top),
          x: Math.round(rect.left),
        })
      })

    Array.from(root.querySelectorAll<HTMLElement>('.vodka-lineview-content'))
      .filter(line => !line.closest('table'))
      .map(line => {
        const rect = line.getBoundingClientRect()
        const codeBlock = line.querySelector<HTMLElement>(
          '.vodka-embeddedobject-code-block-wrapper, .CodeMirror',
        )
        return {
          text: codeBlock
            ? (codeMirrorToMarkdown(codeBlock) ??
              normalizeVodkaLineText(line.innerText))
            : normalizeVodkaLineText(line.innerText),
          y: Math.round(scrollTop + rect.top),
          x: Math.round(rect.left),
        }
      })
      .filter(({ text }) => text)
      .forEach(record => {
        records.push(record)
      })

    Array.from(root.querySelectorAll<HTMLImageElement>('img'))
      .filter(image => !image.closest('table'))
      .map(image => {
        const src = image.currentSrc || image.src
        const rect = image.getBoundingClientRect()
        return {
          alt: normalizeText(image.alt),
          filename: '',
          src,
          y: Math.round(scrollTop + rect.top),
        }
      })
      .filter(image => image.src && !image.src.startsWith('data:'))
      .forEach(image => {
        if (!images.has(image.src)) {
          images.set(image.src, {
            ...image,
            filename: `__vodka_image_${images.size.toFixed()}__`,
          })
        }
      })
  }

  for (let top = 0; top <= maxScrollTop + step; top += step) {
    scroller.scrollTop = Math.min(top, maxScrollTop)
    await waitFor(0.45 * Second)
    collectVisible()
  }

  scroller.scrollTop = initialScrollTop

  const uniqueLines = Array.from(
    new Map(
      records.map(record => [`${record.y.toFixed()}:${record.text}`, record]),
    ).values(),
  )
    .sort((a, b) => a.y - b.y || a.x - b.x)
    .filter((line, index, lines) => line.text !== lines[index - 1]?.text)

  const title = normalizeDocumentFileName(document.title)

  const body =
    uniqueLines.at(0)?.text.replace(/^\d+/, '') === title
      ? uniqueLines.slice(1)
      : uniqueLines

  const items: { text: string; y: number; x?: number }[] = [
    { text: `# ${title}`, y: Number.NEGATIVE_INFINITY },
    ...body,
    ...Array.from(images.values()).map(image => ({
      text: `![${image.alt || 'image'}](${image.filename})`,
      y: image.y ?? Number.MAX_SAFE_INTEGER,
      x: 0,
    })),
  ]

  return {
    markdown:
      items
        .sort((a, b) => a.y - b.y || (a.x ?? 0) - (b.x ?? 0))
        .map(item => item.text)
        .join('\n\n') + '\n',
    images: Array.from(images.values()),
    title,
  }
}

const exportVodkaDocumentAsZip = async (): Promise<boolean> => {
  const result = await collectVodkaDocument()
  if (!result) return false

  const { title, images } = result
  const zipFs = new fs.FS()

  let markdown = result.markdown
  for (let index = 0; index < images.length; index++) {
    const image = images[index]
    const asset = await fetchGenericImageBlob(image.src)
    if (!asset) {
      markdown = markdown.replace(`](${image.filename})`, `](${image.src})`)
      continue
    }

    const extension =
      extensionFromContentType(asset.contentType) ??
      extensionFromContentType(asset.blob.type) ??
      extensionFromUrl(image.src)
    const filename = `image-${(index + 1).toFixed().padStart(3, '0')}${extension}`
    zipFs.addBlob(`${title}/images/${filename}`, asset.blob)
    markdown = markdown.replace(`](${image.filename})`, `](images/${filename})`)
  }

  zipFs.addText(`${title}/${title}.md`, markdown)
  legacyFileSave(await zipFs.exportBlob(), {
    fileName: `${title}.zip`,
  })

  return true
}

const createGenericDomTransformer = () => {
  const images: GenericImage[] = []
  const imageSrcToFilename = new Map<string, string>()

  const imageMarkdown = (image: HTMLImageElement): string => {
    const src = image.currentSrc || image.src
    if (!src || src.startsWith('data:')) return ''

    let filename = imageSrcToFilename.get(src)
    if (!filename) {
      filename = `images/image-${(images.length + 1)
        .toFixed()
        .padStart(3, '0')}${extensionFromUrl(src)}`
      imageSrcToFilename.set(src, filename)
      images.push({
        alt: normalizeText(image.alt),
        filename,
        src: new URL(src, location.href).toString(),
      })
    }

    return `![${escapeMarkdown(normalizeText(image.alt))}](${filename})`
  }

  const inline = (node: Node): string => {
    if (node.nodeType === Node.TEXT_NODE) {
      return escapeMarkdown(normalizeText(node.textContent ?? ''))
    }

    if (!(node instanceof HTMLElement) || isHiddenElement(node)) {
      return ''
    }

    const children = () => Array.from(node.childNodes).map(inline).join(' ')

    switch (node.tagName.toLowerCase()) {
      case 'br':
        return '  \n'
      case 'strong':
      case 'b':
        return `**${children()}**`
      case 'em':
      case 'i':
        return `*${children()}*`
      case 'code':
        return `\`${normalizeText(node.textContent)}\``
      case 'a': {
        const href = node.getAttribute('href')
        const text = children() || escapeMarkdown(normalizeText(href ?? ''))
        if (!href) return text
        return `[${text}](${new URL(href, location.href).toString()})`
      }
      case 'img':
        return imageMarkdown(node as HTMLImageElement)
      default:
        return children()
    }
  }

  const block = (node: Node, depth = 0): string[] => {
    if (node.nodeType === Node.TEXT_NODE) {
      const text = inline(node)
      return text ? [text] : []
    }

    if (!(node instanceof HTMLElement) || isHiddenElement(node)) {
      return []
    }

    const tag = node.tagName.toLowerCase()
    const childBlocks = () =>
      Array.from(node.childNodes)
        .map(child => block(child, depth))
        .flat(1)
    const plain = () => normalizeText(getVisibleText(node))

    if (tag === 'script' || tag === 'style' || tag === 'noscript') return []
    if (/^h[1-6]$/.test(tag)) {
      const level = Number(tag.slice(1))
      const text = plain()
      return text ? [`${'#'.repeat(level)} ${escapeMarkdown(text)}`] : []
    }
    if (tag === 'p') {
      const text = inline(node)
      return text ? [text] : []
    }
    if (tag === 'pre') {
      const className = [
        node.className,
        node.querySelector('code')?.className ?? '',
      ].join(' ')
      const language = /language-([\w-]+)/.exec(className)?.[1] ?? ''

      return [fencedCodeBlock(node.textContent.trimEnd(), language)]
    }
    if (tag === 'blockquote') {
      const text = childBlocks().join('\n\n') || inline(node)
      return text ? [text.replace(/^/gm, '> ').replace(/^> $/gm, '>')] : []
    }
    if (tag === 'img') {
      const text = imageMarkdown(node as HTMLImageElement)
      return text ? [text] : []
    }
    if (tag === 'li') {
      const text = childBlocks().join('\n\n') || inline(node)
      const prefix = `${'  '.repeat(depth)}- `
      return text
        ? [prefix + text.replace(/\n/g, `\n${'  '.repeat(depth + 1)}`)]
        : []
    }
    if (tag === 'ul' || tag === 'ol') {
      return Array.from(node.children)
        .map(child => block(child, depth))
        .flat(1)
    }
    if (tag === 'table') {
      const markdown = htmlTableToMarkdown(node)
      return markdown ? [markdown] : []
    }

    const blocks = childBlocks()
    if (blocks.length > 0) return blocks

    const text = plain()
    return text ? [escapeMarkdown(text)] : []
  }

  return {
    transform(root: HTMLElement): { markdown: string; images: GenericImage[] } {
      return {
        markdown: block(root).filter(Boolean).join('\n\n') + '\n',
        images,
      }
    },
  }
}

const exportGenericDocumentAsMarkdown = async (): Promise<void> => {
  if (await exportVodkaDocumentAsZip()) {
    return
  }

  const root = selectGenericDocumentRoot()
  const transformer = createGenericDomTransformer()
  const result = transformer.transform(root)
  const sourceTitle = [
    document.title,
    root.querySelector('h1')?.textContent,
  ].find(value => value)
  const cleanedTitle = cleanDocumentTitle(sourceTitle ?? '')
  const displayTitle = cleanedTitle ? cleanedTitle : 'doc'
  const title = normalizeDocumentFileName(displayTitle)
  const markdownBody = result.markdown.trim()
  let markdown =
    markdownBody.startsWith(`# ${escapeMarkdown(displayTitle)}`) ||
    markdownBody.startsWith(`# ${displayTitle}`)
      ? `${markdownBody}\n`
      : `# ${escapeMarkdown(displayTitle)}\n\n${markdownBody}\n`
  const zipFs = new fs.FS()

  await Promise.allSettled(
    result.images.map(async image => {
      const asset = await fetchGenericImageBlob(image.src)

      if (!asset) {
        markdown = markdown.replaceAll(
          `](${image.filename})`,
          `](${image.src})`,
        )
        return
      }

      const extension =
        extensionFromContentType(asset.contentType) ??
        extensionFromContentType(asset.blob.type) ??
        extensionFromUrl(image.src)
      const filename = replacePathExtension(image.filename, extension)

      zipFs.addBlob(`${title}/${filename}`, asset.blob)

      if (filename !== image.filename) {
        markdown = markdown.replaceAll(`](${image.filename})`, `](${filename})`)
      }
    }),
  )

  result.images.forEach(image => {
    if (markdown.includes(`](${image.filename})`)) {
      markdown = markdown.replaceAll(`](${image.filename})`, `](${image.src})`)
    }
  })

  zipFs.addText(`${title}/${title}.md`, markdown)
  legacyFileSave(await zipFs.exportBlob(), {
    fileName: `${title}.zip`,
  })
}

type File = mdast.Image | mdast.Link

const downloadFiles = async (
  files: File[],
  options: ProgressOptions & {
    /**
     * @default 3
     */
    batchSize?: number
    signal?: AbortSignal
    useUUID?: boolean
    markdownFileName?: string
  } = {},
): Promise<DownloadResult[]> => {
  const {
    onProgress,
    onComplete,
    batchSize = 3,
    signal,
    useUUID = false,
    markdownFileName = '',
  } = options

  let completeEventCalled = false
  const onCompleteOnce = () => {
    if (!completeEventCalled) {
      completeEventCalled = true
      onComplete?.()
    }
  }

  const results = await withSignal(
    async isAborted => {
      const _results: DownloadResult[] = []

      const totalSize = files.length
      let downloadedSize = 0

      for (const batch of cluster(files, batchSize)) {
        if (isAborted()) {
          break
        }

        await Promise.allSettled(
          batch.map(async file => {
            if (isAborted()) {
              return
            }

            try {
              const result =
                file.type === 'image'
                  ? await downloadImage(file, {
                      signal,
                      useUUID,
                      markdownFileName,
                    })
                  : await downloadFile(file, {
                      signal,
                      useUUID,
                      markdownFileName,
                    })

              if (result) {
                _results.push(result)
              }
            } finally {
              downloadedSize++

              if (!isAborted()) {
                onProgress?.(downloadedSize / totalSize)
              }
            }
          }),
        )
      }

      onCompleteOnce()

      return _results
    },
    {
      signal,
      onAbort: onCompleteOnce,
    },
  )

  return results ?? []
}

interface PrepareResult {
  isReady: boolean
  recoverScrollTop?: () => void
}

const prepare = async (): Promise<PrepareResult> => {
  const checkIsReady = () => docx.isReady({ checkWhiteboard: true })

  let recoverScrollTop

  if (!checkIsReady()) {
    const initialScrollTop = docx.container?.scrollTop ?? 0
    recoverScrollTop = () => {
      docx.scrollTo({
        top: initialScrollTop,
        behavior: 'instant',
      })
    }

    let top = 0

    docx.scrollTo({
      top,
      behavior: 'instant',
    })

    const maxTryTimes = OneHundred
    let tryTimes = 0

    Toast.loading({
      content: i18next.t(TranslationKey.SCROLL_DOCUMENT),
      keepAlive: true,
      key: TranslationKey.SCROLL_DOCUMENT,
      actionText: i18next.t(TranslationKey.CANCEL),
      onActionClick: () => {
        tryTimes = maxTryTimes
      },
    })

    while (!checkIsReady() && tryTimes <= maxTryTimes) {
      docx.scrollTo({
        top,
        behavior: 'smooth',
      })

      await waitFor(0.4 * Second)

      tryTimes++

      top = docx.container?.scrollHeight ?? 0
    }

    Toast.remove(TranslationKey.SCROLL_DOCUMENT)
  }

  return {
    isReady: checkIsReady(),
    recoverScrollTop,
  }
}

const main = async (options: { signal?: AbortSignal } = {}) => {
  const { signal } = options

  if (docx.isDoc) {
    Toast.warning({ content: i18next.t(TranslationKey.NOT_SUPPORT_DOC_1_0) })

    throw new Error(DOWNLOAD_ABORTED)
  }

  if (!docx.isDocx) {
    if (isSupportedGenericDocument()) {
      await exportGenericDocumentAsMarkdown()

      return
    }

    Toast.warning({ content: i18next.t(TranslationKey.NOT_SUPPORT) })

    throw new Error(DOWNLOAD_ABORTED)
  }

  const { isReady, recoverScrollTop } = await prepare()

  if (!isReady) {
    Toast.warning({
      content: i18next.t(TranslationKey.CONTENT_LOADING),
    })

    throw new Error(DOWNLOAD_ABORTED)
  }

  const settings = await getSettings([
    SettingKey.DownloadMethod,
    SettingKey.Table,
    SettingKey.Grid,
    SettingKey.TextHighlight,
    SettingKey.DownloadFileWithUniqueName,
  ])

  const { root, images, files, tableWithParents, mentionUsers } =
    docx.intoMarkdownAST({
      whiteboard: true,
      diagram: true,
      file: true,
      highlight: settings[SettingKey.TextHighlight],
      flatGrid: settings[SettingKey.Grid] === Grid.Flatten,
    })

  await transformMentionUsers(mentionUsers)

  const recommendName = docx.pageTitle
    ? normalizeFileName(docx.pageTitle.slice(0, OneHundred))
    : 'doc'
  const isZip = images.length > 0 || files.length > 0
  const ext = isZip ? '.zip' : '.md'
  const filename = `${recommendName}${ext}`

  const toBlob = async () => {
    Toast.loading({
      content: i18next.t(TranslationKey.STILL_SAVING),
      keepAlive: true,
      key: ToastKey.DOWNLOADING,
    })

    const singleFileContent = () => {
      transformTableBySettings(tableWithParents, settings)

      const markdown = Docx.stringify(root)

      return new Blob([markdown])
    }

    const zipFileContent = async () => {
      const zipFs = new fs.FS()

      const imgs = images.filter(image => image.data?.fetchSources)
      const diagrams = images.filter(image => image.data?.fetchBlob)

      const results = await Promise.all([
        downloadFiles(imgs, {
          batchSize: 15,
          onProgress: progress => {
            Toast.loading({
              content: i18next.t(TranslationKey.DOWNLOAD_PROGRESS, {
                name: i18next.t(TranslationKey.IMAGE),
                progress: Math.floor(progress * OneHundred),
              }),
              keepAlive: true,
              key: TranslationKey.IMAGE,
            })
          },
          onComplete: () => {
            Toast.remove(TranslationKey.IMAGE)
          },
          signal,
          useUUID: settings[SettingKey.DownloadFileWithUniqueName],
          markdownFileName: recommendName,
        }),
        // Diagrams must be downloaded one by one
        downloadFiles(diagrams, {
          batchSize: 1,
          signal,
          useUUID: settings[SettingKey.DownloadFileWithUniqueName],
          markdownFileName: recommendName,
        }),
        downloadFiles(files, {
          onProgress: progress => {
            Toast.loading({
              content: i18next.t(TranslationKey.DOWNLOAD_PROGRESS, {
                name: i18next.t(TranslationKey.FILE),
                progress: Math.floor(progress * OneHundred),
              }),
              keepAlive: true,
              key: TranslationKey.FILE,
            })
          },
          onComplete: () => {
            Toast.remove(TranslationKey.FILE)
          },
          signal,
          useUUID: settings[SettingKey.DownloadFileWithUniqueName],
          markdownFileName: recommendName,
        }),
      ])
      results.flat(1).forEach(({ filename, content }) => {
        zipFs.addBlob(filename, content)
      })

      transformTableBySettings(tableWithParents, settings)

      const markdown = Docx.stringify(root)

      zipFs.addText(`${recommendName}.md`, markdown)

      return await zipFs.exportBlob()
    }

    const content = isZip ? await zipFileContent() : singleFileContent()

    recoverScrollTop?.()

    return content
  }

  if (
    settings[SettingKey.DownloadMethod] === DownloadMethod.ShowSaveFilePicker &&
    supported
  ) {
    if (!navigator.userActivation.isActive) {
      const confirmed = await confirm()
      if (!confirmed) {
        throw new Error(DOWNLOAD_ABORTED)
      }
    }

    await fileSave(toBlob(), {
      fileName: filename,
      extensions: [ext],
    })
  } else {
    const blob = await toBlob()

    legacyFileSave(blob, {
      fileName: filename,
    })
  }
}

let controller = new AbortController()
main({
  signal: controller.signal,
})
  .then(() => {
    Toast.success({
      content: i18next.t(TranslationKey.DOWNLOAD_COMPLETE),
    })
  })
  .catch((error: unknown) => {
    const aborted =
      error instanceof Error &&
      (error.name === 'AbortError' || error.message === DOWNLOAD_ABORTED)

    if (aborted) {
      controller.abort()
    } else {
      Toast.error({
        key: ToastKey.REPORT_BUG,
        content: String(error),
        actionText: i18next.t(CommonTranslationKey.CONFIRM_REPORT_BUG, {
          ns: Namespace.COMMON,
        }),
        duration: Minute,
        onActionClick: () => {
          reportBug(error)

          Toast.remove(ToastKey.REPORT_BUG)
        },
      })
    }
  })
  .finally(() => {
    Toast.remove(ToastKey.DOWNLOADING)

    // @ts-expect-error remove reference
    controller = null
  })
