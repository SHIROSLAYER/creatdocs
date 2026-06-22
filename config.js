/* =====================================================================
   Configuração PÚBLICA do Supabase (login + dados compartilhados).
   ---------------------------------------------------------------------
   A chave publishable/anon é segura no cliente: o acesso é controlado
   por RLS no banco. Pode ir para o repositório público.
   (Dados pessoais ficam em config.local.js, que NÃO vai para o git.)
   ===================================================================== */
window.SUPABASE_CONFIG = {
  url: 'https://wxiansccenhbgjoejnyh.supabase.co',
  key: 'sb_publishable_C6M4qvmzYnE0DfOrp4w33A_HMYBHKS0',
};

/* reCAPTCHA v2 ("Não sou robô") no login. Deixe a siteKey vazia para desligar
   o captcha. Site key é pública (pode ir no repo); a SECRET fica só na edge
   function (segredo RECAPTCHA_SECRET), nunca aqui. */
window.RECAPTCHA = {
  siteKey: '6LerRS4tAAAAAMuPsIg1nLMVKdoAXgYWRptNdKY_',
};
