import {
  Flag,
  RuntimeMessageType,
  type FetchAssetMessage,
  type FetchAssetResponse,
} from './common/message'

const larkDocumentUrlPatterns: string[] = [
  'https://*.feishu.cn/*',
  'https://*.feishu.net/*',
  'https://*.larksuite.com/*',
  'https://*.feishu-pre.net/*',
  'https://*.larkoffice.com/*',
  'https://*.larkenterprise.com/*',
]

const sharedDocumentUrlPatterns: string[] = [
  ...larkDocumentUrlPatterns,
  'https://docs.corp.kuaishou.com/*',
  'https://docs.qingque.cn/*',
  'https://kstack.corp.kuaishou.com/*',
]

const genericDocumentHosts = new Set([
  'docs.corp.kuaishou.com',
  'docs.qingque.cn',
  'kstack.corp.kuaishou.com',
])

const isAllowedAssetFetch = (src: string, senderUrl?: string): boolean => {
  try {
    const assetUrl = new URL(src)
    const documentUrl = senderUrl ? new URL(senderUrl) : null
    const documentHost = documentUrl?.hostname

    if (!documentHost || !genericDocumentHosts.has(documentHost)) {
      return false
    }

    if (assetUrl.protocol !== 'https:') {
      return false
    }

    if (assetUrl.hostname === documentHost) {
      return true
    }

    return (
      documentHost === 'kstack.corp.kuaishou.com' &&
      assetUrl.hostname === 'static.yximgs.com'
    )
  } catch {
    return false
  }
}

const arrayBufferToBase64 = (buffer: ArrayBuffer): string => {
  const bytes = new Uint8Array(buffer)
  const chunkSize = 0x8000
  const chunks: string[] = []

  for (let offset = 0; offset < bytes.length; offset += chunkSize) {
    chunks.push(
      String.fromCharCode(...bytes.subarray(offset, offset + chunkSize)),
    )
  }

  return btoa(chunks.join(''))
}

const blobToDataUrl = async (
  blob: Blob,
  contentType: string | null,
): Promise<string> => {
  const mimeType =
    [blob.type, contentType].find(value => value) ?? 'application/octet-stream'
  return `data:${mimeType};base64,${arrayBufferToBase64(
    await blob.arrayBuffer(),
  )}`
}

const fetchAsset = async (
  message: FetchAssetMessage,
  sender: chrome.runtime.MessageSender,
): Promise<FetchAssetResponse> => {
  try {
    if (!isAllowedAssetFetch(message.src, sender.url)) {
      throw new Error('Asset fetch is not allowed for this page')
    }

    const response = await fetch(message.src, { credentials: 'include' })
    if (!response.ok) {
      throw new Error(`HTTP error! status: ${response.status.toFixed()}`)
    }

    const contentType = response.headers.get('content-type')
    const blob = await response.blob()

    return {
      ok: true,
      dataUrl: await blobToDataUrl(blob, contentType),
      contentType: blob.type || contentType,
    }
  } catch (error) {
    return {
      ok: false,
      error: String(error),
    }
  }
}

const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === 'object' && value !== null

const isFetchAssetMessage = (message: unknown): message is FetchAssetMessage =>
  isRecord(message) && message['type'] === RuntimeMessageType.FetchAsset

const isExecuteScriptMessage = (message: unknown): message is { flag: Flag } =>
  isRecord(message) && 'flag' in message

enum MenuItemId {
  DOWNLOAD_DOCX_AS_MARKDOWN = 'download_docx_as_markdown',
  DOWNLOAD_DOCX_AS_HTML = 'download_docx_as_html',
  COPY_DOCX_AS_MARKDOWN = 'copy_docx_as_markdown',
  VIEW_DOCX_AS_MARKDOWN = 'view_docx_as_markdown',
}

chrome.runtime.onInstalled.addListener(() => {
  chrome.contextMenus.removeAll(() => {
    chrome.contextMenus.create({
      id: MenuItemId.DOWNLOAD_DOCX_AS_MARKDOWN,
      title: chrome.i18n.getMessage('download_docx_as_markdown'),
      documentUrlPatterns: sharedDocumentUrlPatterns,
      contexts: ['page', 'editable'],
    })

    chrome.contextMenus.create({
      id: MenuItemId.DOWNLOAD_DOCX_AS_HTML,
      title: chrome.i18n.getMessage('download_docx_as_html'),
      documentUrlPatterns: sharedDocumentUrlPatterns,
      contexts: ['page', 'editable'],
    })

    chrome.contextMenus.create({
      id: MenuItemId.COPY_DOCX_AS_MARKDOWN,
      title: chrome.i18n.getMessage('copy_docx_as_markdown'),
      documentUrlPatterns: sharedDocumentUrlPatterns,
      contexts: ['page', 'editable'],
    })

    chrome.contextMenus.create({
      id: MenuItemId.VIEW_DOCX_AS_MARKDOWN,
      title: chrome.i18n.getMessage('view_docx_as_markdown'),
      documentUrlPatterns: sharedDocumentUrlPatterns,
      contexts: ['page', 'editable'],
    })
  })
})

const executeScriptByFlag = async (flag: string | number, tabId: number) => {
  switch (flag) {
    case MenuItemId.DOWNLOAD_DOCX_AS_MARKDOWN:
      await chrome.scripting.executeScript({
        files: ['bundles/scripts/download-lark-docx-as-markdown.js'],
        target: { tabId },
        world: 'MAIN',
      })
      break
    case MenuItemId.DOWNLOAD_DOCX_AS_HTML:
      await chrome.scripting.executeScript({
        files: ['bundles/scripts/download-lark-docx-as-html.js'],
        target: { tabId },
        world: 'MAIN',
      })
      break
    case MenuItemId.COPY_DOCX_AS_MARKDOWN:
      await chrome.scripting.executeScript({
        files: ['bundles/scripts/copy-lark-docx-as-markdown.js'],
        target: { tabId },
        world: 'MAIN',
      })
      break
    case MenuItemId.VIEW_DOCX_AS_MARKDOWN:
      await chrome.scripting.executeScript({
        files: ['bundles/scripts/view-lark-docx-as-markdown.js'],
        target: { tabId },
        world: 'MAIN',
      })
      break
    default:
      break
  }
}

chrome.contextMenus.onClicked.addListener(({ menuItemId }, tab) => {
  if (tab?.id !== undefined) {
    executeScriptByFlag(menuItemId, tab.id).catch(console.error)
  }
})

chrome.runtime.onMessage.addListener(
  (message: unknown, sender, sendResponse) => {
    if (isFetchAssetMessage(message)) {
      fetchAsset(message, sender).then(sendResponse).catch(console.error)

      return true
    }

    if (!isExecuteScriptMessage(message)) {
      return false
    }

    const executeScript = async () => {
      const activeTabs = await chrome.tabs.query({
        currentWindow: true,
        active: true,
      })

      const activeTabId = activeTabs.at(0)?.id

      if (activeTabs.length === 1 && activeTabId !== undefined) {
        await executeScriptByFlag(message.flag, activeTabId)
      }
    }

    executeScript().then(sendResponse).catch(console.error)

    return true
  },
)
