import { describe, expect, it } from 'vitest'
import {
  codeMirrorToMarkdown,
  fencedCodeBlock,
  htmlTableToMarkdown,
} from '../src/scripts/vodka-markdown'

describe('codeMirrorToMarkdown', () => {
  it('extracts code without the toolbar or line numbers', () => {
    const root = document.createElement('div')
    root.innerHTML = `
      <div class="code-block-react-header">
        <a class="code-block-react-language-label">Markdown</a>
        <span>自动换行</span><span>折叠</span>
      </div>
      <div class="CodeMirror-code">
        <div><span class="CodeMirror-linenumber">1</span><pre class="CodeMirror-line"><span>swift sft \\</span></pre></div>
        <div><span class="CodeMirror-linenumber">2</span><pre class="CodeMirror-line"><span>  --model Qwen3</span></pre></div>
      </div>
    `

    expect(codeMirrorToMarkdown(root)).toBe(
      '```markdown\nswift sft \\\n  --model Qwen3\n```',
    )
  })

  it('uses a longer fence when the code contains backticks', () => {
    const root = document.createElement('div')
    root.innerHTML = `
      <a class="code-block-react-language-label">Plain Text</a>
      <pre class="CodeMirror-line">\`\`\`</pre>
    `

    expect(codeMirrorToMarkdown(root)).toBe('````text\n```\n````')
  })
})

describe('fencedCodeBlock', () => {
  it('keeps the fence and code on consecutive lines', () => {
    expect(fencedCodeBlock('const value = 1', 'javascript')).toBe(
      '```javascript\nconst value = 1\n```',
    )
  })
})

describe('htmlTableToMarkdown', () => {
  it('preserves rows and escapes pipes and line breaks', () => {
    const table = document.createElement('table')
    table.innerHTML = `
      <tbody>
        <tr><td>名称</td><td>说明</td></tr>
        <tr><td>A | B</td><td>第一行<br>第二行</td></tr>
      </tbody>
    `

    expect(htmlTableToMarkdown(table)).toBe(
      [
        '| 名称 | 说明 |',
        '| --- | --- |',
        '| A \\| B | 第一行<br>第二行 |',
      ].join('\n'),
    )
  })
})
