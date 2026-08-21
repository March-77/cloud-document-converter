const normalizeLanguage = (value: string): string => {
  const language = value.trim().toLowerCase()
  if (!language) return ''

  const aliases: Record<string, string> = {
    'c++': 'cpp',
    'c#': 'csharp',
    'plain text': 'text',
    plaintext: 'text',
    shell: 'bash',
  }

  return aliases[language] ?? language.replace(/\s+/g, '-')
}

export const fencedCodeBlock = (value: string, language: string): string => {
  const longestBacktickRun = Math.max(
    0,
    ...Array.from(value.matchAll(/`+/g), match => match[0].length),
  )
  const fence = '`'.repeat(Math.max(3, longestBacktickRun + 1))

  return `${fence}${language}\n${value}\n${fence}`
}

export const codeMirrorToMarkdown = (root: HTMLElement): string | null => {
  const lines = Array.from(
    root.querySelectorAll<HTMLElement>('.CodeMirror-line'),
  )
  if (lines.length === 0) return null

  const language = normalizeLanguage(
    root.querySelector<HTMLElement>('.code-block-react-language-label')
      ?.textContent ?? '',
  )
  const code = lines
    .map(line => line.textContent.replace(/\u200b/g, ''))
    .join('\n')
    .replace(/\n+$/, '')

  return fencedCodeBlock(code, language)
}

const escapeTableCell = (value: string): string =>
  value
    .replace(/\u00a0/g, ' ')
    .replace(/\r?\n+/g, '<br>')
    .replace(/\\/g, '\\\\')
    .replace(/\|/g, '\\|')
    .replace(/[ \t]+/g, ' ')
    .trim()

const tableCellText = (node: Node): string =>
  node.nodeType === Node.TEXT_NODE
    ? (node.textContent ?? '')
    : node instanceof HTMLElement && node.tagName === 'BR'
      ? '\n'
      : Array.from(node.childNodes).map(tableCellText).join('')

export const htmlTableToMarkdown = (table: HTMLElement): string | null => {
  const rows = Array.from(table.querySelectorAll('tr')).map(row =>
    Array.from(row.children)
      .filter(
        (cell): cell is HTMLElement =>
          cell instanceof HTMLElement &&
          (cell.tagName === 'TH' || cell.tagName === 'TD'),
      )
      .map(cell => escapeTableCell(tableCellText(cell))),
  )
  if (rows.length === 0) return null

  const columnCount = Math.max(...rows.map(row => row.length))
  if (columnCount === 0) return null

  const normalizedRows = rows.map(row =>
    Array.from({ length: columnCount }, (_, index) => row[index] ?? ''),
  )
  const [header, ...body] = normalizedRows

  return [
    `| ${header.join(' | ')} |`,
    `| ${header.map(() => '---').join(' | ')} |`,
    ...body.map(row => `| ${row.join(' | ')} |`),
  ].join('\n')
}
