// =====================================================================
// verify-recaptcha — confere um token do Google reCAPTCHA v2 no servidor.
// A SECRET nunca vai para o front; fica no segredo RECAPTCHA_SECRET da função.
//
// Deploy (painel): Edge Functions → Deploy a new function → nome
//   "verify-recaptcha" → cole este arquivo. IMPORTANTE: desligue "Verify JWT"
//   (o usuário ainda não está logado quando resolve o captcha).
// Segredo: Edge Functions → Secrets → RECAPTCHA_SECRET = <sua secret key v2>.
// =====================================================================
const cors = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
};

function json(obj: unknown, status = 200): Response {
  return new Response(JSON.stringify(obj), {
    status,
    headers: { ...cors, "Content-Type": "application/json" },
  });
}

Deno.serve(async (req: Request) => {
  if (req.method === "OPTIONS") return new Response("ok", { headers: cors });
  if (req.method !== "POST") return json({ success: false, error: "method not allowed" }, 405);

  try {
    const { token } = await req.json().catch(() => ({}));
    if (!token) return json({ success: false, error: "missing token" }, 400);

    const secret = Deno.env.get("RECAPTCHA_SECRET");
    if (!secret) return json({ success: false, error: "RECAPTCHA_SECRET not set" }, 500);

    const body = new URLSearchParams({ secret, response: String(token) });
    const r = await fetch("https://www.google.com/recaptcha/api/siteverify", {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body,
    });
    const data = await r.json();
    return json({
      success: !!data.success,
      errors: data["error-codes"] ?? null,
    });
  } catch (e) {
    return json({ success: false, error: String(e) }, 500);
  }
});
