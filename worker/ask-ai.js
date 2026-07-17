// Cloudflare Worker: proxy seguro entre el Dashboard Comercial Open y la API de Anthropic.
// La ANTHROPIC_API_KEY vive como secret en Cloudflare (Settings > Variables and Secrets),
// nunca en el código del dashboard ni en este archivo.
//
// Este Worker es deliberadamente "tonto": solo agrega la API key y el system prompt,
// y reenvía { messages, tools } a la API de Claude tal cual, regresando la respuesta cruda.
// El loop de tool use (ejecutar cada tool contra los datos cargados en el navegador y
// volver a preguntar) vive en app.js, porque los datos del dashboard solo existen ahí.

const ALLOWED_ORIGINS = [
  'https://jvargas-cso.github.io',
  'http://localhost:8934',
];

const SYSTEM_PROMPT = `Eres un analista de datos comerciales para "Dashboard Comercial Open", una empresa de medios OOH (espectaculares, pantallas, vallas, transporte público, centros comerciales, etc.).

No ves los datos directamente: tienes herramientas para consultarlos (agregaciones por dimensión, series mensuales, totales, búsqueda de líneas específicas). Usa las herramientas cuantas veces haga falta, encadenando varias si la pregunta lo requiere, antes de responder.

Los datos que consultes ya respetan los filtros que el usuario tiene activos en el dashboard (año, vendedor, cliente, etc.) — no necesitas pedir eso, ya viene aplicado.

Responde en español, de forma directa y concisa, citando cifras concretas. Si después de consultar las herramientas disponibles no encuentras la información pedida (por ejemplo, una columna que no existe en los datos), dilo claramente en vez de inventar.`;

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

    const messages = Array.isArray(body.messages) ? body.messages : [];
    const tools = Array.isArray(body.tools) ? body.tools : [];
    if (!messages.length) {
      return json({ error: 'Falta messages' }, 400, corsHeaders);
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
        max_tokens: 1536,
        thinking: { type: 'disabled' },
        system: SYSTEM_PROMPT,
        tools,
        messages,
      }),
    });

    const data = await anthropicRes.json();
    if (!anthropicRes.ok) {
      return json({ error: data?.error?.message || 'Error al consultar Claude' }, 502, corsHeaders);
    }

    // Reenviamos la respuesta cruda de Claude (content, stop_reason, etc.) — el
    // navegador decide si hay que ejecutar tool calls y volver a preguntar.
    return json(data, 200, corsHeaders);
  },
};

function json(obj, status, corsHeaders) {
  return new Response(JSON.stringify(obj), {
    status,
    headers: { ...corsHeaders, 'Content-Type': 'application/json' },
  });
}
