# Supabase — Gerador de Documentos Infobarra

Backend de **login + dados compartilhados** (o app continua rodando local pelo `INICIAR.bat`;
o Word/Access e o PDF nítido seguem funcionando — só os dados/login vão para a nuvem).

## Passo a passo (conta nova)

1. **Criar o projeto** no Supabase (região: South America / São Paulo).
2. **Authentication → Sign In / Up:** **DESATIVE** "Allow new users to sign up"
   (`Enable email signups` = OFF) → ninguém cria conta sozinho; só você cadastra.
3. **SQL Editor:** cole e rode o `schema.sql` (cria tabelas, RLS, trigger de perfil e o bucket `assinados`).
4. **Cadastrar os operadores** (Authentication → Users → **Add user**): informe **email + senha**
   de cada pessoa (usuários criados no painel já vêm confirmados). Para tornar alguém **admin**,
   ajuste o `role` dele na tabela `perfis` (ou o 1º usuário já vira admin pelo trigger).
5. Em **Settings → API**, copie e me envie:
   - **Project URL** (`https://xxxxx.supabase.co`)
   - **anon / public key** (segura para o app — protegida por RLS; pode ir no repo).
   - ⚠️ a **service_role** é secreta — não compartilhe.

## Como o login vai funcionar

- Supabase **Auth** (sessão/JWT real), login por **email + senha**.
- **Contas são criadas por você no painel do Supabase** (Authentication → Users). O app **não tem
  cadastro** — só login. Cadastro aberto fica **desativado**.
- `perfis` (ligado a `auth.users`) guarda `nome` + `role` (admin/operador); o trigger cria o perfil
  ao criar o usuário (1º vira admin). RLS: qualquer autenticado **vê**; só admin **gerencia** perfis.
- Tirar acesso de alguém = remover/desativar no painel do Supabase (vale na hora, em todas as máquinas).

## Modelo de dados

| Tabela | Para quê |
|---|---|
| `perfis` | operador ↔ `auth.users` (nome, role) |
| `clientes` | base de clientes (`tipo`, `contrato`, `nome`, `cpf`, `dados` jsonb) |
| `operacoes` | termos gerados (quem criou / assinado / anexo) — alimenta os Relatórios |
| `estoque` | estoque de equipamentos (marca/modelo/qtd) |
| `catalogo` | catálogo Marca→Modelos aprendido |
| storage `assinados` | PDFs assinados anexados |

## O que falta (eu faço quando você enviar URL + anon key)

- Vendorizar/`<script>` o **supabase-js**; criar `config.js` com URL + anon key (a anon é pública, pode ir no git).
- Trocar `storeGet/storeSet`, `loadData/saveLocal`, `loadEstoque/saveEstoque`, `loadCatalog/learnCatalog`
  e o login (`gateLogin/loginOperador/criarOperador`) para falarem com o Supabase.
- **Importar uma vez** os 385 + 116 clientes atuais (dos JSON locais) para a tabela `clientes`.
- Detecção: se o Supabase não estiver configurado, o app cai no modo local atual (servidor `/__store`).
