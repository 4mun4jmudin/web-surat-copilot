/**
 * functions/api/ai.js - Cloudflare Pages Function
 * Proxy aman ke API AI (Gemini / OpenAI compatible)
 * Mencegah kebocoran API key di client-side, mendukung model normalization, dan validasi response.
 */

export async function onRequestPost(context) {
  const { request, env } = context;

  // 1. CORS Headers
  const corsHeaders = {
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'POST, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, Authorization, X-Custom-Key',
    'Content-Type': 'application/json'
  };

  // Batas ukuran payload request: 500 KB
  const contentLength = request.headers.get('content-length');
  if (contentLength && parseInt(contentLength, 10) > 500 * 1024) {
    return new Response(JSON.stringify({ error: 'Ukuran payload melebihi batas 500 KB.' }), {
      status: 413,
      headers: corsHeaders
    });
  }

  let body;
  try {
    body = await request.json();
  } catch (err) {
    return new Response(JSON.stringify({ error: 'Format JSON request tidak valid.' }), {
      status: 400,
      headers: corsHeaders
    });
  }

  let { prompt, systemPrompt, stream = false, model = 'gemini-2.5-flash', provider = 'gemini' } = body;
  if (!prompt) {
    return new Response(JSON.stringify({ error: 'Parameter "prompt" wajib disertakan.' }), {
      status: 400,
      headers: corsHeaders
    });
  }

  // Normalisasi nama model Gemini untuk mencegah 404 & 429 omni (Google 2026 update)
  const m = (model || '').toLowerCase();
  if (!m || m.includes('omni') || m.includes('2.0-flash') || m.includes('2-flash') || m.includes('1.5-pro') || m === 'gemini-pro') {
    model = 'gemini-2.5-flash';
  }

  // Gunakan API Key dari Environment Variable Cloudflare Pages atau custom header
  const apiKey = env.AI_API_KEY || request.headers.get('X-Custom-Key') || body.apiKey;

  if (!apiKey) {
    return new Response(JSON.stringify({
      error: 'API Key AI tidak ditemukan di server. Masukkan API Key Anda di menu Pengaturan / Gudang API (BYOK Mode).'
    }), {
      status: 401,
      headers: corsHeaders
    });
  }

  try {
    // Panggilan ke Google Gemini API v1beta
    const geminiUrl = `https://generativelanguage.googleapis.com/v1beta/models/${model}:${stream ? 'streamGenerateContent?alt=sse&' : 'generateContent?'}key=${apiKey}`;

    const geminiPayload = {
      contents: [
        {
          role: 'user',
          parts: [{ text: prompt }]
        }
      ],
      generationConfig: {
        temperature: 0.2,
        responseMimeType: 'application/json'
      }
    };

    if (systemPrompt) {
      geminiPayload.systemInstruction = {
        parts: [{ text: systemPrompt }]
      };
    }

    const aiResponse = await fetch(geminiUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(geminiPayload)
    });

    if (!aiResponse.ok) {
      const errText = await aiResponse.text();
      return new Response(JSON.stringify({ error: `AI API Error (${aiResponse.status}): ${errText}` }), {
        status: aiResponse.status,
        headers: corsHeaders
      });
    }

    if (stream) {
      return new Response(aiResponse.body, {
        headers: {
          'Access-Control-Allow-Origin': '*',
          'Content-Type': 'text/event-stream; charset=utf-8',
          'Cache-Control': 'no-cache',
          'Connection': 'keep-alive'
        }
      });
    }

    const resJson = await aiResponse.json();
    const generatedText = resJson.candidates?.[0]?.content?.parts?.[0]?.text || '';

    return new Response(JSON.stringify({ text: generatedText }), {
      status: 200,
      headers: corsHeaders
    });
  } catch (err) {
    return new Response(JSON.stringify({ error: `Serverless Proxy Error: ${err.message}` }), {
      status: 500,
      headers: corsHeaders
    });
  }
}

export async function onRequestOptions() {
  return new Response(null, {
    status: 204,
    headers: {
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'POST, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type, Authorization, X-Custom-Key'
    }
  });
}
