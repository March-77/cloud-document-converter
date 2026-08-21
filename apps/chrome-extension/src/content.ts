import { chunk } from 'es-toolkit/array'
import {
  EventName,
  RuntimeMessageType,
  WindowMessageType,
  portImpl,
  type FetchAssetResponse,
  type WindowFetchAssetRequest,
} from './common/message'

const COMMENT_BUTTON_CLASS = '.docx-comment__first-comment-btn'
const HELP_BLOCK_CLASS = '.help-block'
const DOWNLOAD_ONLY_HOSTS = new Set([
  'docs.corp.kuaishou.com',
  'docs.qingque.cn',
  'kstack.corp.kuaishou.com',
])
const KSTACK_STATIC_IMAGE_HOST = 'static.yximgs.com'

let disposables: (() => void)[] = []

const dispose = (): void => {
  disposables.forEach(disposable => {
    disposable()
  })

  disposables = []
}

interface Button {
  element: HTMLElement
  width: number
  height: number
}

const isDownloadOnlyHost = (): boolean =>
  DOWNLOAD_ONLY_HOSTS.has(location.hostname)

const isWindowFetchAssetRequest = (
  data: unknown,
): data is WindowFetchAssetRequest =>
  typeof data === 'object' &&
  data !== null &&
  'type' in data &&
  data.type === WindowMessageType.FetchAssetRequest &&
  'id' in data &&
  typeof data.id === 'string' &&
  'src' in data &&
  typeof data.src === 'string'

const isAllowedAssetBridgeRequest = (src: string): boolean => {
  try {
    const url = new URL(src, location.href)
    if (url.protocol !== 'https:') return false
    if (url.hostname === location.hostname) return true

    return (
      location.hostname === 'kstack.corp.kuaishou.com' &&
      url.hostname === KSTACK_STATIC_IMAGE_HOST
    )
  } catch {
    return false
  }
}

const initAssetFetchBridge = (): void => {
  if (!isDownloadOnlyHost()) return

  const handleMessage = async (event: MessageEvent<unknown>) => {
    if (event.source !== window || event.origin !== location.origin) return
    if (!isWindowFetchAssetRequest(event.data)) return

    const { id, src } = event.data
    let response: FetchAssetResponse

    try {
      response = isAllowedAssetBridgeRequest(src)
        ? ((await chrome.runtime.sendMessage({
            type: RuntimeMessageType.FetchAsset,
            src: new URL(src, location.href).toString(),
          })) as FetchAssetResponse)
        : {
            ok: false,
            error: 'Asset fetch is not allowed for this page',
          }
    } catch (error) {
      response = {
        ok: false,
        error: String(error),
      }
    }

    window.postMessage(
      {
        type: WindowMessageType.FetchAssetResponse,
        id,
        response,
      },
      location.origin,
    )
  }

  const onMessage = (event: MessageEvent<unknown>): void => {
    handleMessage(event).catch(console.error)
  }

  window.addEventListener('message', onMessage)
  disposables.push(() => {
    window.removeEventListener('message', onMessage)
  })
}

const initButtons = (): void => {
  const root = document.body
  const isDownloadOnly = isDownloadOnlyHost()

  const isReady = () => {
    if (isDownloadOnly) return true

    // Comment button may not be displayed
    for (const selector of [HELP_BLOCK_CLASS]) {
      if (!root.querySelector(selector)) {
        return false
      }
    }
    return true
  }

  const render = () => {
    const style = document.createElement('style')
    style.innerHTML = `
  [data-CDC-button-type] {
    position: fixed;
    display: flex;
    justify-content: center;
    align-items: center;
    border: 1px solid var(--line-border-card);
    border-radius: 50%;
    background-color: var(--bg-body);
    box-shadow: var(--shadow-s4-down);
    cursor: pointer;
    text-align: center;
    color: var(--text-title);
    z-index: 3;
  }

  [data-CDC-button-type]:hover {
    color: var(--colorful-blue);
  }
  `

    const operates = [
      {
        type: 'copy',
        innerHtml: `<svg aria-hidden="true" focusable="false" role="img" class="octicon octicon-copy" viewBox="0 0 16 16" width="16"
          height="16" fill="currentColor">
          <path
              d="M0 6.75C0 5.784.784 5 1.75 5h1.5a.75.75 0 0 1 0 1.5h-1.5a.25.25 0 0 0-.25.25v7.5c0 .138.112.25.25.25h7.5a.25.25 0 0 0 .25-.25v-1.5a.75.75 0 0 1 1.5 0v1.5A1.75 1.75 0 0 1 9.25 16h-7.5A1.75 1.75 0 0 1 0 14.25Z">
          </path>
          <path
              d="M5 1.75C5 .784 5.784 0 6.75 0h7.5C15.216 0 16 .784 16 1.75v7.5A1.75 1.75 0 0 1 14.25 11h-7.5A1.75 1.75 0 0 1 5 9.25Zm1.75-.25a.25.25 0 0 0-.25.25v7.5c0 .138.112.25.25.25h7.5a.25.25 0 0 0 .25-.25v-7.5a.25.25 0 0 0-.25-.25Z">
          </path>
          </svg>`,
        action: () => {
          chrome.runtime
            .sendMessage({ flag: 'copy_docx_as_markdown' })
            .catch(console.error)
        },
      },
      {
        type: 'view',
        innerHtml: `<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" 
        stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" 
        class="octicon octicon-view" fill="currentColor">
          <path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8"/>
          <circle cx="12" cy="12" r="3"/>
        </svg>
        `,
        action: () => {
          chrome.runtime
            .sendMessage({ flag: 'view_docx_as_markdown' })
            .catch(console.error)
        },
      },
      {
        type: 'download',
        innerHtml: `<svg aria-hidden="true" focusable="false" role="img" class="octicon octicon-download" viewBox="0 0 16 16"
        width="16" height="16" fill="currentColor">
        <path
          d="M2.75 14A1.75 1.75 0 0 1 1 12.25v-2.5a.75.75 0 0 1 1.5 0v2.5c0 .138.112.25.25.25h10.5a.25.25 0 0 0 .25-.25v-2.5a.75.75 0 0 1 1.5 0v2.5A1.75 1.75 0 0 1 13.25 14Z">
        </path>
        <path
          d="M7.25 7.689V2a.75.75 0 0 1 1.5 0v5.689l1.97-1.969a.749.749 0 1 1 1.06 1.06l-3.25 3.25a.749.749 0 0 1-1.06 0L4.22 6.78a.749.749 0 1 1 1.06-1.06l1.97 1.969Z">
        </path>
      </svg>`,
        action: () => {
          chrome.runtime
            .sendMessage({ flag: 'download_docx_as_markdown' })
            .catch(console.error)
        },
      },
      {
        type: 'download-html',
        innerHtml: `<svg aria-hidden="true" focusable="false" role="img" class="octicon octicon-file-code" viewBox="0 0 16 16"
        width="16" height="16" fill="currentColor">
        <path d="M2 1.75C2 .784 2.784 0 3.75 0h5.586c.464 0 .909.184 1.237.513l2.914 2.914c.329.328.513.773.513 1.237v9.586A1.75 1.75 0 0 1 12.25 16h-8.5A1.75 1.75 0 0 1 2 14.25Zm1.75-.25a.25.25 0 0 0-.25.25v12.5c0 .138.112.25.25.25h8.5a.25.25 0 0 0 .25-.25V6h-2.75A1.75 1.75 0 0 1 8 4.25V1.5Zm5.75.062V4.25c0 .138.112.25.25.25h2.688Z"></path>
        <path d="M6.03 7.72a.75.75 0 0 1 0 1.06L4.81 10l1.22 1.22a.75.75 0 1 1-1.06 1.06l-1.75-1.75a.75.75 0 0 1 0-1.06l1.75-1.75a.75.75 0 0 1 1.06 0Zm3.94 0a.75.75 0 0 1 1.06 0l1.75 1.75a.75.75 0 0 1 0 1.06l-1.75 1.75a.75.75 0 1 1-1.06-1.06L11.19 10 9.97 8.78a.75.75 0 0 1 0-1.06ZM8.77 7.19a.75.75 0 0 1 .54.91l-1 4a.75.75 0 0 1-1.46-.36l1-4a.75.75 0 0 1 .92-.55Z"></path>
      </svg>`,
        action: () => {
          chrome.runtime
            .sendMessage({ flag: 'download_docx_as_html' })
            .catch(console.error)
        },
      },
    ]

    const buttons = operates
      .filter(
        ({ type }) =>
          !isDownloadOnly || type === 'download' || type === 'download-html',
      )
      .map<Button>(({ type, innerHtml, action }) => {
        const btn = document.createElement('button')
        btn.type = 'button'
        btn.setAttribute('data-CDC-button-type', type)
        btn.innerHTML = innerHtml

        btn.style.width = '36px'
        btn.style.height = '36px'

        btn.addEventListener('click', action)

        return {
          element: btn,
          width: 36,
          height: 36,
        }
      })

    const getOriginalBtnPos = (buttons: Button[]) => {
      const defaultBtnHeight = 36
      const defaultGap = 14
      const helpBlock: HTMLDivElement | null =
        root.querySelector(HELP_BLOCK_CLASS)
      if (!helpBlock) {
        if (isDownloadOnly) {
          return buttons.map((_, index) => ({
            right: 24,
            bottom: 24 + index * (defaultGap + defaultBtnHeight),
          }))
        }

        return
      }

      const windowWidth = window.innerWidth
      const windowHeight = window.innerHeight

      const helpBlockRect = helpBlock.getBoundingClientRect()

      const commentButton: HTMLDivElement | null =
        root.querySelector(COMMENT_BUTTON_CLASS)

      // Comment button may not be displayed
      if (!commentButton) {
        const startBottom = windowHeight - helpBlockRect.bottom
        const right = windowWidth - helpBlockRect.right

        return buttons.map((_, index) => ({
          right,
          bottom: startBottom + (index + 1) * (defaultGap + defaultBtnHeight),
        }))
      }

      const commentButtonRect = commentButton.getBoundingClientRect()

      if (commentButtonRect.right === helpBlockRect.right) {
        const btnHeight = commentButtonRect.height
        const gap =
          Math.abs(helpBlockRect.bottom - commentButtonRect.bottom) - btnHeight
        const initialBottom = Math.max(
          windowHeight - commentButtonRect.bottom,
          windowHeight - helpBlockRect.bottom,
        )

        return buttons.map((_, index) => ({
          right: windowWidth - commentButtonRect.right,
          bottom: initialBottom + (index + 1) * (gap + btnHeight),
        }))
      } else if (commentButtonRect.bottom === helpBlockRect.bottom) {
        let bottom = windowHeight - commentButtonRect.bottom

        return chunk(buttons, 2)
          .map(items => {
            bottom += defaultGap + defaultBtnHeight

            if (items.length === 2) {
              return [
                { right: windowWidth - commentButtonRect.right, bottom },
                { right: windowWidth - helpBlockRect.right, bottom },
              ]
            }

            const minRight =
              windowWidth -
              Math.max(commentButtonRect.right, helpBlockRect.right)

            return [
              {
                right: minRight,
                bottom,
              },
            ]
          })
          .flat(1)
      }

      return
    }

    const layout = (buttons: Button[]) => {
      const pos = getOriginalBtnPos(buttons)
      if (!pos) return

      buttons.forEach((button, index) => {
        button.element.style.right = pos[index].right.toFixed() + 'px'
        button.element.style.bottom = pos[index].bottom.toFixed() + 'px'
      })
    }

    // When the width of the docx's visible content is too narrow, the position of the comment button changes.
    const autoLayoutObserver = new MutationObserver(mutations => {
      for (const mutation of mutations) {
        if (
          mutation.type === 'attributes' &&
          mutation.attributeName === 'class'
        ) {
          layout(buttons)
        }
      }
    })
    const autoLayout = () => {
      const commentButton = root.querySelector(COMMENT_BUTTON_CLASS)
      if (!commentButton) {
        return
      }

      autoLayoutObserver.observe(commentButton, {
        attributes: true,
        attributeFilter: ['class'],
      })
    }

    layout(buttons)
    autoLayout()

    root.appendChild(style)
    buttons.forEach(button => {
      root.appendChild(button.element)
    })

    const unmount = () => {
      autoLayoutObserver.disconnect()

      buttons.forEach(button => {
        if (root.contains(button.element)) {
          root.removeChild(button.element)
        }
      })

      if (root.contains(style)) {
        root.removeChild(style)
      }
    }

    disposables.push(unmount)

    return unmount
  }

  const checkCapture = async () => {
    try {
      const { manualCaptureActive } = await chrome.storage.local.get(
        'manualCaptureActive',
      )
      if (manualCaptureActive) {
        chrome.runtime
          .sendMessage({
            type: 'CDC_CAPTURED_DOCUMENT',
            url: window.location.href,
            title: document.title || 'Untitled Document',
          })
          .catch(console.error)
      }
    } catch (e) {
      console.error('Failed to report capture state:', e)
    }
  }

  let unmount: (() => void) | null = null
  const init = () => {
    // Rendering may be called multiple times
    if (unmount) {
      unmount()

      unmount = null
    }

    unmount = render()
    void checkCapture()
  }

  const initObserver = new MutationObserver(mutations => {
    for (const mutation of mutations) {
      if (mutation.type === 'childList') {
        for (const node of mutation.addedNodes) {
          if (
            node instanceof HTMLElement &&
            (node.matches(COMMENT_BUTTON_CLASS) ||
              node.matches(HELP_BLOCK_CLASS)) &&
            isReady()
          ) {
            init()
          }
        }
      }
    }
  })

  if (isReady()) {
    init()
  }

  initObserver.observe(root, {
    childList: true,
    subtree: true,
  })

  disposables.push(() => {
    initObserver.disconnect()
  })
}

function initContent(): void {
  initAssetFetchBridge()
  initButtons()
}

window.addEventListener('load', initContent, false)

// For SPA, some page content updates do not trigger page reloads
let lastPathname: string = location.pathname
const urlChangeObserver: MutationObserver = new MutationObserver(() => {
  const pathname = location.pathname

  if (pathname !== lastPathname) {
    lastPathname = pathname

    dispose()

    initContent()
  }
})
urlChangeObserver.observe(document.body, { childList: true })

portImpl.receiver.on(EventName.GetSettings, async keys => {
  return await chrome.storage.sync.get(keys)
})

if (import.meta.env.DEV) {
  portImpl.receiver.on(EventName.Console, data => {
    console.log('MAIN World Console:', ...data)
  })
}

const isFeishuDocUrl = (urlStr: string): boolean => {
  try {
    const url = new URL(urlStr)
    const hostname = url.hostname
    const pathname = url.pathname

    const isFeishuDomain =
      hostname.endsWith('.feishu.cn') ||
      hostname.endsWith('.feishu.net') ||
      hostname.endsWith('.larksuite.com') ||
      hostname.endsWith('.feishu-pre.net') ||
      hostname.endsWith('.larkoffice.com') ||
      hostname.endsWith('.larkenterprise.com')

    if (!isFeishuDomain) return false

    return (
      pathname.includes('/docx/') ||
      pathname.includes('/wiki/') ||
      pathname.includes('/doc/')
    )
  } catch {
    return false
  }
}

document.addEventListener(
  'click',
  e => {
    const handle = async () => {
      try {
        const { manualCaptureActive } = await chrome.storage.local.get(
          'manualCaptureActive',
        )
        if (!manualCaptureActive) return

        const target = e.target as HTMLElement
        const anchor = target.closest('a')
        if (anchor?.href) {
          const url = anchor.href
          if (isFeishuDocUrl(url)) {
            if (anchor.target !== '_blank') {
              window.open(url, '_blank')
              e.preventDefault()
              e.stopPropagation()
            }
          }
        }
      } catch (err) {
        console.error('Error in manual capture click handler:', err)
      }
    }
    void handle()
  },
  { capture: true },
)

window.addEventListener('message', event => {
  if (event.source !== window) return
  const data = event.data as Record<string, unknown> | null
  if (data && data['type'] === 'CDC_EXTRACTED_DATA_INTERNAL') {
    void chrome.runtime.sendMessage({
      type: data['success'] ? 'CDC_EXTRACTION_SUCCESS' : 'CDC_EXTRACTION_ERROR',
      data: data['data'],
      error: data['error'],
    })
  }
})
