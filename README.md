# Gerador de Documentos — Infobarra

Aplicação web (local) que substitui o fluxo de **mala direta do Word + Access** para
emitir **Contrato de Comodato de Equipamentos** e **Recibo de Equipamento**, gerando
os documentos preenchidos em **DOCX** e **PDF** direto no navegador.

---

## ▶ Como usar

1. Dê **duplo clique em `INICIAR.bat`**.
   - Abre uma janela do servidor (deixe aberta) e o navegador em `http://localhost:8777/`.
2. À esquerda, **selecione um cliente** (busca por nome, CPF/CNPJ ou contrato) ou clique **“+ Novo”**.
3. Escolha o documento no topo da barra lateral: **Contrato de Comodato** ou **Recibo de Equipamento**.
4. Ajuste os campos e os equipamentos (marque as caixas, preencha marca/modelo/qtd).
5. Clique em **⬇ Baixar DOCX** (editável, idêntico ao modelo Word) ou **⬇ Baixar PDF** (pronto para imprimir).
6. **💾 Salvar cliente** grava/edita o cliente na base local do navegador.

> Não precisa de internet para gerar os documentos. A fonte “Inter” usa internet só para o visual; o resto funciona offline.

---

## 📁 Estrutura

```
projeto word/
├─ INICIAR.bat            ← inicia o app (duplo clique)
├─ index.html             ← interface
├─ styles.css             ← estilo
├─ app.js                 ← lógica (formulário, geração DOCX/PDF)
├─ server.ps1             ← servidor estático local (PowerShell, sem admin)
├─ data/
│  ├─ clientes_comodato.json   ← 385 registros importados do Access
│  ├─ clientes_recibo.json     ← 116 registros importados do Access
│  └─ catalogo.json            ← Marca → Modelos dos comboboxes (EDITÁVEL)
├─ templates/
│  ├─ template_contrato.docx   ← modelo Word com tokens {{CAMPO}}
│  └─ template_recibo.docx     ← modelo Word com tokens {{CAMPO}}
├─ assets/                     ← logos do cabeçalho do PDF (ISO, INFOBARRA, ANATEL, ABRINQ)
├─ vendor/
│  ├─ jszip.min.js             ← gera o DOCX
│  └─ html2pdf.bundle.min.js   ← gera o PDF
└─ _source/              ← arquivos de origem e scripts de build (não usados em runtime)
   ├─ contrato_original.docx / recibo_original.docx   ← modelos Word originais
   ├─ _comodato.mdb / _recibo.mdb                     ← bancos Access originais
   ├─ _export_access.ps1                              ← exporta .mdb → JSON
   ├─ conv.pl                                         ← converte MERGEFIELD → {{token}}
   └─ _repackage.ps1                                  ← reempacota o .docx com o XML convertido
```

---

## 🔧 Como funciona

- **DOCX:** o template Word original teve cada campo de mala direta (`MERGEFIELD`) trocado por
  um token `{{NOME}}` (preservando a formatação). Em runtime, o `app.js` abre o `.docx` (que é um ZIP)
  com JSZip, substitui os tokens pelos valores do formulário e baixa o arquivo — **fidelidade total** ao modelo.
- **PDF (alta qualidade / vetorial):** o app preenche o DOCX e envia ao servidor local (`/__pdf`), que usa o
  **Microsoft Word** (`ExportAsFixedFormat`) para converter em PDF — texto nítido/selecionável, idêntico ao
  documento oficial. **Requer o Word instalado e o app rodando pelo `INICIAR.bat`** (servidor local).
  - *Fallback:* se o Word/servidor não estiver disponível (ex.: abrindo o `index.html` direto), o app gera um
    PDF alternativo via html2pdf (imagem) — funciona, mas com qualidade inferior.
- **Dados:** os `.mdb` (formato Access/ACE) foram exportados para JSON. Adições/edições ficam no
  `localStorage` do navegador e se mesclam com a base importada (registro local vence por nº de contrato).

---

## 🔄 Reimportar dados do Access (quando os .mdb forem atualizados)

Requer o **Microsoft Access** instalado (usa automação COM). No PowerShell:

```powershell
# Comodato
& ".\_source\_export_access.ps1" -mdb "CAMINHO\MB COMODATO.mdb" -out ".\data\clientes_comodato.json"
# Recibo
& ".\_source\_export_access.ps1" -mdb "CAMINHO\MD RECIBO.mdb"  -out ".\data\clientes_recibo.json"
```

A tabela lida é **`Office Address List`** (padrão da mala direta do Word).

---

## 🔐 Login (obrigatório para criar/editar)

Ao abrir o app aparece a **tela de login** — só é possível **criar ou editar documentos logado**.
- **Primeiro acesso:** crie o operador **administrador** (nome + senha).
- Depois, cada operador entra com **nome + senha**. A sessão fica salva (loga uma vez e permanece) — para
  trocar de operador, use **Sair** (no canto da barra lateral, abaixo do menu).
- Todo documento gerado/impresso é **carimbado com o operador** (aparece nos Relatórios).
- **Exige rodar pelo `INICIAR.bat`** (o login usa o servidor).

## 🖨️ Imprimir (sem baixar)

No rodapé, ao lado de **Baixar DOCX**, o botão **🖨️ Imprimir** gera o documento e **manda direto para a
impressora**, sem precisar baixar o `.docx`/PDF antes. (Usa o Word via servidor — rode pelo `INICIAR.bat`.)

## 📊 Aba Relatórios

Abaixo do Estoque.

- Cada operador entra com **nome + senha**; a partir daí, todo documento gerado é **carimbado com o operador**.
- O relatório lista os **termos** de todos os operadores: data, operador, tipo, contrato, cliente e **status**
  (Assinado / Não assinado), com filtros e busca.
- **Status de assinatura:** marque manualmente (✔) **ou** anexe o **PDF assinado** (📎) — anexar marca como assinado.
- **Administrador** cadastra/remove operadores; operador comum só vê o relatório.
- Senhas são guardadas em **hash** (não em texto). Dados em `data/store/` — pronto para migrar para nuvem (Supabase).

## 📦 Aba Estoque

No menu lateral, **Estoque** abre o controle de equipamentos: escolha **Marca** e **Modelo**
(o Modelo só lista após escolher a Marca) e a quantidade, clique **+ Adicionar**. Na tabela:
- a **quantidade é editável** — clique no número, digite e clique fora;
- **− / +** ajustam de 1 em 1; **🗑** remove (com confirmação);
- **baixa automática:** ao **gerar um documento** (DOCX/PDF), os equipamentos marcados são
  descontados do estoque (não desconta 2× se você gerar DOCX e PDF do mesmo documento).

## 🧩 Equipamento "Outros"

Na seção *Objeto do Comodato*, a linha **Outros** é preenchível como as demais: marque a caixa
e preencha quantidade/marca/modelo. Sai no DOCX e no PDF.

## 🔌 Marca e Modelo (comboboxes dependentes)

Os campos **Marca** e **Modelo** dos equipamentos são listas suspensas: escolha a marca e o
campo Modelo passa a sugerir apenas os modelos daquela marca. Ainda é possível **digitar** um
valor novo que não esteja na lista.

**Catálogo que aprende sozinho:** quando você digita uma **marca ou modelo novo** (na aba Estoque,
ao salvar um cliente ou ao gerar um documento), ele é **memorizado** automaticamente. Da próxima vez,
ao escolher aquela marca no Comodato/Recibo, o novo modelo já aparece na lista. Os aprendidos ficam
salvos no navegador (`infobarra_catalogo_user_v1`) e somam-se ao catálogo base.

O catálogo base fica em **`data/catalogo.json`** (editável):
```json
{ "marcas": { "FIBERHOME": ["GPON ONU AN5506", "..."], "TP-LINK": ["..."] } }
```

## ✍️ Testemunhas do contrato

As testemunhas saem fixas no Contrato, preenchidas a partir de um arquivo **local** que **não vai para
o repositório** (dados pessoais ficam só na máquina):

1. Copie `config.example.js` para **`config.local.js`**.
2. Preencha `witnesses` com nome + CPF das duas testemunhas.

O template (`templates/template_contrato.docx`) usa os tokens `{{TEST1_NOME}}`, `{{TEST1_CPF}}`,
`{{TEST2_NOME}}`, `{{TEST2_CPF}}`, que o app substitui na geração. Sem `config.local.js`, o contrato
sai sem testemunhas (nenhum dado pessoal no repositório público).

## 💾 Backup / Restauração

- **⬇ Backup** (rodapé): abre um menu para escolher o formato — **JSON** (edições locais) ou
  **.mdb** (Access) da base de **Comodato** ou de **Recibo** (recria a tabela `Office Address List`).
- **⬆ Restaurar base**: importa um **JSON** de backup **ou** um banco **Microsoft Access (.mdb)**.
  - Ao escolher um `.mdb`, o servidor lê a tabela `Office Address List` (via Access) e o app **detecta
    automaticamente** se é Comodato ou Recibo (pelas colunas) e atualiza a base daquele tipo.
  - O `.mdb` exige rodar pelo `INICIAR.bat` (usa o Access instalado).
- Os botões Backup/Restaurar ficam ocultos na aba **Estoque**.

> A base **importada original** (385 + 116) está nos arquivos `data/*.json`. O backup cobre só o que você
> adicionou/editou dentro do app.
