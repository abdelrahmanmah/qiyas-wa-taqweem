import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'
import { readdir, readFile } from 'fs/promises'
import { existsSync } from 'fs'
import path from 'path'
import { load as yamlLoad } from 'js-yaml'

function yamlLoaderPlugin() {
  return {
    name: 'yaml-loader',
    transform(code, id) {
      if (!id.endsWith('.yaml') && !id.endsWith('.yml')) return;
      const obj = yamlLoad(code);
      return { code: `export default ${JSON.stringify(obj)}`, map: null };
    },
  };
}

// ── Format converters (OpenAI ↔ Anthropic) ────────────────────────────────────
function toAnthropicMessages(messages) {
  return messages
    .filter(m => m.role !== 'system')
    .map(msg => {
      if (msg.role === 'assistant') {
        const content = [];
        if (msg.content) content.push({ type: 'text', text: msg.content });
        if (msg.tool_calls) {
          for (const tc of msg.tool_calls) {
            content.push({
              type: 'tool_use', id: tc.id, name: tc.function.name,
              input: JSON.parse(tc.function.arguments || '{}'),
            });
          }
        }
        return { role: 'assistant', content };
      }
      if (msg.role === 'tool') {
        return {
          role: 'user',
          content: [{ type: 'tool_result', tool_use_id: msg.tool_call_id, content: msg.content }],
        };
      }
      return { role: msg.role, content: msg.content };
    });
}

function toAnthropicTools(tools) {
  return tools.map(t => ({
    name: t.function.name,
    description: t.function.description,
    input_schema: t.function.parameters,
  }));
}

function fromAnthropicToOpenAI(data) {
  const content = data.content || [];
  const text = content.filter(c => c.type === 'text').map(c => c.text).join('');
  const toolUses = content.filter(c => c.type === 'tool_use');
  const message = { role: 'assistant', content: text || null };
  if (toolUses.length > 0) {
    message.tool_calls = toolUses.map(tu => ({
      id: tu.id, type: 'function',
      function: { name: tu.name, arguments: JSON.stringify(tu.input) },
    }));
  }
  return {
    choices: [{
      message,
      finish_reason: data.stop_reason === 'tool_use' ? 'tool_calls' : 'stop',
    }],
  };
}

const PROVIDER_URLS = {
  groq:   'https://api.groq.com/openai/v1/chat/completions',
  gemini: 'https://generativelanguage.googleapis.com/v1beta/openai/chat/completions',
};

async function callOpenAICompat(provider, apiKey, model, messages, tools, systemPrompt) {
  const allMessages = systemPrompt
    ? [{ role: 'system', content: systemPrompt }, ...messages]
    : messages;
  const body = { model, messages: allMessages, max_tokens: 4096, temperature: 0.3 };
  if (tools?.length) {
    body.tools = tools;
    body.tool_choice = 'auto';
    body.parallel_tool_calls = false; // prevents malformed multi-tool output on Groq/Llama
  }

  const resp = await fetch(PROVIDER_URLS[provider], {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${apiKey}` },
    body: JSON.stringify(body),
  });
  if (!resp.ok) {
    const t = await resp.text();
    throw new Error(`${provider} API error (${resp.status}): ${t.slice(0, 400)}`);
  }
  return resp.json();
}

async function callClaude(apiKey, model, messages, tools, systemPrompt) {
  const body = { model, max_tokens: 4096, messages: toAnthropicMessages(messages) };
  if (systemPrompt) body.system = systemPrompt;
  if (tools?.length) body.tools = toAnthropicTools(tools);

  const resp = await fetch('https://api.anthropic.com/v1/messages', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'x-api-key': apiKey,
      'anthropic-version': '2023-06-01',
    },
    body: JSON.stringify(body),
  });
  if (!resp.ok) {
    const t = await resp.text();
    throw new Error(`Claude API error (${resp.status}): ${t.slice(0, 400)}`);
  }
  return fromAnthropicToOpenAI(await resp.json());
}

// ── File system helpers ───────────────────────────────────────────────────────
async function listSurveyFiles(baseFolder, subfolder) {
  const target = subfolder ? path.join(baseFolder, subfolder) : baseFolder;
  if (!existsSync(target)) return { error: `المجلد غير موجود: ${target}` };

  const files = [];
  async function walk(dir, sub) {
    let entries;
    try { entries = await readdir(dir, { withFileTypes: true }); } catch { return; }
    for (const e of entries) {
      const full = path.join(dir, e.name);
      if (e.isDirectory()) {
        await walk(full, e.name);
      } else if (/\.(xlsx?|xls)$/i.test(e.name)) {
        files.push({ path: full, name: e.name, subfolder: sub || '' });
      }
    }
  }

  await walk(target, subfolder || '');
  return { files };
}

function readSurveyFile(filePath, baseFolder) {
  const resolved    = path.resolve(filePath);
  const resolvedBase = path.resolve(baseFolder);
  if (!resolved.startsWith(resolvedBase)) throw new Error('Access denied: path outside surveys folder');
  if (!existsSync(resolved)) throw new Error(`الملف غير موجود: ${resolved}`);
  return readFile(resolved); // returns Buffer
}

// ── Vite plugin ───────────────────────────────────────────────────────────────
function chatApiPlugin() {
  return {
    name: 'chat-api',
    configureServer(server) {

      // POST /api/chat — proxy to AI provider
      server.middlewares.use('/api/chat', (req, res) => {
        res.setHeader('Content-Type', 'application/json');
        if (req.method !== 'POST') { res.statusCode = 405; return res.end('{}'); }

        let body = '';
        req.on('data', c => { body += c.toString(); });
        req.on('end', async () => {
          try {
            const { provider, model, apiKey: clientKey, messages, tools, systemPrompt } = JSON.parse(body);
            const envKey = process.env[`${String(provider).toUpperCase()}_API_KEY`];
            const apiKey = envKey || clientKey;

            if (!apiKey) {
              res.statusCode = 401;
              return res.end(JSON.stringify({
                error: `مفتاح API غير موجود. أضفه في الإعدادات أو عرّف ${String(provider).toUpperCase()}_API_KEY في .env`,
              }));
            }

            const result = provider === 'claude'
              ? await callClaude(apiKey, model, messages, tools, systemPrompt)
              : await callOpenAICompat(provider, apiKey, model, messages, tools, systemPrompt);

            res.end(JSON.stringify(result));
          } catch (err) {
            res.statusCode = 500;
            res.end(JSON.stringify({ error: err.message }));
          }
        });
      });

      // POST /api/list-files — list xlsx files in surveys folder
      server.middlewares.use('/api/list-files', (req, res) => {
        res.setHeader('Content-Type', 'application/json');
        if (req.method !== 'POST') { res.statusCode = 405; return res.end('{}'); }

        let body = '';
        req.on('data', c => { body += c.toString(); });
        req.on('end', async () => {
          try {
            const { surveysFolder, subfolder } = JSON.parse(body);
            if (!surveysFolder) return res.end(JSON.stringify({ error: 'surveysFolder not configured' }));
            const result = await listSurveyFiles(surveysFolder, subfolder);
            res.end(JSON.stringify(result));
          } catch (err) {
            res.statusCode = 500;
            res.end(JSON.stringify({ error: err.message }));
          }
        });
      });

      // POST /api/read-file — read xlsx file and return as base64
      server.middlewares.use('/api/read-file', (req, res) => {
        res.setHeader('Content-Type', 'application/json');
        if (req.method !== 'POST') { res.statusCode = 405; return res.end('{}'); }

        let body = '';
        req.on('data', c => { body += c.toString(); });
        req.on('end', async () => {
          try {
            const { filePath, surveysFolder } = JSON.parse(body);
            const buf = await readSurveyFile(filePath, surveysFolder);
            res.end(JSON.stringify({ data: buf.toString('base64'), name: path.basename(filePath) }));
          } catch (err) {
            res.statusCode = 500;
            res.end(JSON.stringify({ error: err.message }));
          }
        });
      });

    },
  };
}

export default defineConfig({
  plugins: [yamlLoaderPlugin(), react(), chatApiPlugin()],
  server: { port: 3000, open: true },
})
