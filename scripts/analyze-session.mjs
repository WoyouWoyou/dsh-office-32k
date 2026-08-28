// Analyze a DSH session log (plaintext jsonl) for tool-call statistics.
// Usage: node analyze-session.mjs <session-dir>
import { readdirSync, readFileSync, existsSync } from 'node:fs'
import { join } from 'node:path'

const dir = process.argv[2]
if (!dir) { console.error('usage: node analyze-session.mjs <session-dir>'); process.exit(1) }

// find the session log (plaintext .jsonl preferred)
let logPath = join(dir, 'session.jsonl')
if (!existsSync(logPath)) logPath = join(dir, 'session.jsonl.zstd')
if (!existsSync(logPath)) { console.error('no log found in', dir); process.exit(1) }

const text = readFileSync(logPath, 'utf8')
const lines = text.split('\n').filter(Boolean)
const types = {}
let toolCalls = 0, toolErrors = 0, invalidArgs = 0
const toolNames = {}
const errorMsgs = []
const retrySeq = []

for (const line of lines) {
  let e
  try { e = JSON.parse(line) } catch { continue }
  const t = e.type || '?'
  types[t] = (types[t] || 0) + 1
  const d = e.data || {}
  // tool call events: look for name + either start/result shapes
  if (t.includes('tool') || d.name || d.toolName || d.result?.tool) {
    const name = d.name || d.toolName || d.tool || (d.result && d.result.tool) || (d.request && d.request.name) || '?'
    if (name && name !== '?') {
      toolCalls++
      toolNames[name] = (toolNames[name] || 0) + 1
      // retry detection: consecutive identical tool names
      retrySeq.push(name)
    }
    const isErr = d.error !== undefined || d.ok === false
      || (d.result && d.result.error) || (d.result && d.result.ok === false)
    if (isErr) {
      toolErrors++
      const msg = JSON.stringify(d.error || d.result || '')
      if (msg.toLowerCase().includes('invalid')) invalidArgs++
      if (errorMsgs.length < 8) errorMsgs.push(msg.slice(0, 160))
    }
  }
}
// consecutive retry runs (same tool name repeated adjacently)
let retries = 0
for (let i = 1; i < retrySeq.length; i++) if (retrySeq[i] === retrySeq[i - 1]) retries++

console.log('log:', logPath)
console.log('event types:', JSON.stringify(types, null, 0))
console.log('toolCalls:', toolCalls, '| toolErrors:', toolErrors, '| invalidArgs:', invalidArgs, '| adjacent retries:', retries)
console.log('tool names:', JSON.stringify(toolNames))
if (errorMsgs.length) { console.log('first errors:'); errorMsgs.forEach(m => console.log('  -', m)) }
