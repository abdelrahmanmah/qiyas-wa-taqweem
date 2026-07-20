// Vercel serverless function — proxies chat requests to the selected AI provider.
// Mirrors the dev-only middleware in vite.config.js so /api/chat works in production too.

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

export default async function handler(req, res) {
  if (req.method !== 'POST') {
    res.status(405).json({ error: 'Method not allowed' });
    return;
  }

  try {
    const { provider, model, apiKey: clientKey, messages, tools, systemPrompt } = req.body;
    const envKey = process.env[`${String(provider).toUpperCase()}_API_KEY`];
    const apiKey = envKey || clientKey;

    if (!apiKey) {
      res.status(401).json({
        error: `مفتاح API غير موجود. أضفه في الإعدادات أو عرّف ${String(provider).toUpperCase()}_API_KEY في متغيرات البيئة على Vercel`,
      });
      return;
    }

    const result = provider === 'claude'
      ? await callClaude(apiKey, model, messages, tools, systemPrompt)
      : await callOpenAICompat(provider, apiKey, model, messages, tools, systemPrompt);

    res.status(200).json(result);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
}
