// =====================================================================
// admin-create-user — cria um usuário (e-mail + senha + nome + perfil).
// SOMENTE um admin logado consegue chamar. Usa a service_role NO SERVIDOR
// (nunca no front) para criar a conta no Auth.
//
// Deploy (painel): Edge Functions → Deploy → nome "admin-create-user".
//   Mantenha "Verify JWT" LIGADO (só usuário logado chama).
// Não precisa de segredo extra: SUPABASE_URL / SUPABASE_ANON_KEY /
//   SUPABASE_SERVICE_ROLE_KEY já vêm prontos nas Edge Functions.
// =====================================================================
import { createClient } from "jsr:@supabase/supabase-js@2";

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
  if (req.method !== "POST") return json({ error: "method not allowed" }, 405);

  try {
    const url = Deno.env.get("SUPABASE_URL")!;
    const anonKey = Deno.env.get("SUPABASE_ANON_KEY")!;
    const serviceKey = Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!;

    const token = (req.headers.get("Authorization") || "").replace("Bearer ", "").trim();
    if (!token) return json({ error: "sem autenticação" }, 401);

    // 1) Quem está chamando? (valida o JWT do usuário)
    const asUser = createClient(url, anonKey, {
      global: { headers: { Authorization: `Bearer ${token}` } },
    });
    const { data: ures, error: uerr } = await asUser.auth.getUser();
    if (uerr || !ures?.user) return json({ error: "token inválido" }, 401);

    // 2) O chamador é admin? (consulta perfis com service_role)
    const admin = createClient(url, serviceKey, { auth: { persistSession: false } });
    const { data: perfil } = await admin
      .from("perfis").select("role").eq("id", ures.user.id).maybeSingle();
    if (!perfil || perfil.role !== "admin") {
      return json({ error: "apenas administradores podem criar usuários" }, 403);
    }

    // 3) Cria o usuário
    const body = await req.json().catch(() => ({}));
    const email = String(body.email || "").trim().toLowerCase();
    const password = String(body.password || "");
    const nome = String(body.nome || "").trim();
    const role = body.role === "admin" ? "admin" : "operador";
    if (!email || !password) return json({ error: "e-mail e senha são obrigatórios" }, 400);
    if (password.length < 6) return json({ error: "a senha precisa de ao menos 6 caracteres" }, 400);

    const { data: created, error: cerr } = await admin.auth.admin.createUser({
      email,
      password,
      email_confirm: true,
      user_metadata: { nome, role },
    });
    if (cerr) return json({ error: cerr.message }, 400);

    // Reforça o perfil com nome/role (o trigger handle_new_user também cria)
    await admin.from("perfis").upsert({
      id: created.user!.id,
      nome: nome || email.split("@")[0],
      role,
    });

    return json({ ok: true, id: created.user!.id, email, role });
  } catch (e) {
    return json({ error: String(e) }, 500);
  }
});
