/* Service worker mínimo — habilita instalar como PWA (janela própria + ícone).
   Sem cache do código: o app sempre carrega a versão atual (atualizada pela
   pasta de rede / servidor local). */
self.addEventListener('install', () => self.skipWaiting());
self.addEventListener('activate', (e) => e.waitUntil(self.clients.claim()));
self.addEventListener('fetch', () => { /* deixa o navegador buscar normalmente */ });
