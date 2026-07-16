// Cloudflare Worker: proxy seguro entre el Dashboard Comercial Open y la API de Anthropic.
// La ANTHROPIC_API_KEY vive como secret en Cloudflare (Settings > Variables and Secrets),
// nunca en el código del dashboard ni en este archivo.

const ALLOWED_ORIGINS = [
  'https://jvargas-cso.github.io',
  'http://localhost:8934',
];

const SYSTEM_PROMPT = `Eres un analista de datos comerciales para "Dashboard Comercial Open", una empresa de medios OOH (espectaculares, pantallas, vallas, transporte público, centros comerciales, etc.). Se te da un resumen de los datos de ventas actualmente filtrados en el dashboard, y debes responder preguntas del usuario sobre esos datos: tendencias, comparativos, quién vende más, dónde hay riesgo de concentración, estacionalidad, márgenes, etc.

Responde en español, de forma directa y concisa, citando cifras concretas del resumen cuando existan. Si el resumen no trae la información necesaria para responder algo, dilo claramente en vez de inventar datos.`;

export default {
  async fetch(request, env) {
    const origin = request.headers.get('Origin') || '';
    const corsOrigin = ALLOWED_ORIGINS.includes(origin) ? origin : ALLOWED_ORIGINS[0];
    const corsHeaders = {
      'Access-Control-Allow-Origin': corsOrigin,
      'Access-Control-Allow-Methods': 'POST, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type',
    };

    if (request.method === 'OPTIONS') {
      return new Response(null, { headers: corsHeaders });
    }
    if (request.method !== 'POST') {
      return new Response('Method not allowed', { status: 405, headers: corsHeaders });
    }

    let body;
    try {
      body = await request.json();
    } catch {
      return json({ error: 'JSON inválido' }, 400, corsHeaders);
    }

    const question = (body.question || '').toString().slice(0, 2000).trim();
    const context = (body.context || '').toString().slice(0, 30000);
    if (!question) {
      return json({ error: 'Falta la pregunta' }, 400, corsHeaders);
    }

    const anthropicRes = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-api-key': env.ANTHROPIC_API_KEY,
        'anthropic-version': '2023-06-01',
      },
      body: JSON.stringify({
        model: 'claude-sonnet-5',
        max_tokens: 1024,
        thinking: { type: 'disabled' },
        system: SYSTEM_PROMPT,
        messages: [
          { role: 'user', content: `DATOS ACTUALES DEL DASHBOARD (ya filtrados):\n${context}\n\nPREGUNTA: ${question}` },
        ],
      }),
    });

    if (!anthropicRes.ok) {
      const detail = await anthropicRes.text();
      return json({ error: 'Error al consultar Claude', detail }, 502, corsHeaders);
    }

    const data = await anthropicRes.json();
    const answer = (data.content || [])
      .filter((b) => b.type === 'text')
      .map((b) => b.text)
      .join('\n');

    return json({ answer }, 200, corsHeaders);
  },
};

function json(obj, status, corsHeaders) {
  return new Response(JSON.stringify(obj), {
    status,
    headers: { ...corsHeaders, 'Content-Type': 'application/json' },
  });
}
