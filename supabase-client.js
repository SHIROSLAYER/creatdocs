/* =====================================================================
   Cliente Supabase — inicializado a partir de config.js (SUPABASE_CONFIG).
   Sem config válida → window.sb = null e o app cai no modo LOCAL
   (login/dados pelo servidor /__store). Assim nada quebra offline.
   ===================================================================== */
(function () {
  const cfg = window.SUPABASE_CONFIG;
  const lib = window.supabase;
  if (!cfg || !cfg.url || !cfg.key || !lib || typeof lib.createClient !== 'function') {
    window.sb = null;
    window.hasSupabase = () => false;
    return;
  }

  /* "Lembrar login neste dispositivo":
     marcado  → token no localStorage (continua logado mesmo fechando o navegador);
     desmarcado → sessionStorage (sai ao fechar a aba/navegador). */
  const REMEMBER_KEY = 'infobarra_lembrar';
  window.sbRemember = function (on) {
    if (on === undefined) return localStorage.getItem(REMEMBER_KEY) === '1';
    localStorage.setItem(REMEMBER_KEY, on ? '1' : '0');
  };
  const primary = () => (window.sbRemember() ? window.localStorage : window.sessionStorage);
  const authStorage = {
    getItem: (k) => {
      const v = localStorage.getItem(k);
      return v !== null ? v : sessionStorage.getItem(k);
    },
    setItem: (k, v) => {
      primary().setItem(k, v);
      (window.sbRemember() ? sessionStorage : localStorage).removeItem(k);
    },
    removeItem: (k) => {
      localStorage.removeItem(k);
      sessionStorage.removeItem(k);
    },
  };

  window.sb = lib.createClient(cfg.url, cfg.key, {
    auth: {
      persistSession: true,
      autoRefreshToken: true,
      detectSessionInUrl: false,
      storage: authStorage,
      storageKey: 'infobarra-auth',
    },
  });
  window.hasSupabase = () => true;
})();
