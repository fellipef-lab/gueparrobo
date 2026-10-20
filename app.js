/* =====================================================================
   Controle de Atividades e Produtividade — app online (Firebase)
   Baseado na planilha "Controle de Atividades": cada atividade tem data,
   colaborador(es), tipo, local, início, término, itens da lista,
   itens abastecidos e observação. O app calcula tempo, pendentes e
   itens/hora sozinho e monta o resumo por local e por colaborador.
   Perfis: Supervisor (tudo) e Líder (lança e acompanha atividades).
   ===================================================================== */
import * as FA from "https://www.gstatic.com/firebasejs/10.12.2/firebase-app.js";
import * as AU from "https://www.gstatic.com/firebasejs/10.12.2/firebase-auth.js";
import * as FS from "https://www.gstatic.com/firebasejs/10.12.2/firebase-firestore.js";

window.__ativOk = true;

/* ================== CONFIGURAÇÃO ================== */
const C = self.ATIV_CONFIG || {};
const CONFIGURADO = !!(C.firebase && C.firebase.apiKey && !String(C.firebase.apiKey).startsWith("COLE"));
// Versão offline (pasta "offline"): tudo roda dentro do navegador, com o "Firebase de mentira" do demo.js.
// A versão online (pasta "app", que vai para o GitHub) usa sempre o Firebase de verdade.
const DEMO = !!self.ATIV_OFFLINE;
const FK = DEMO ? criarFalso() : null;
const { initializeApp, getApps } = DEMO ? FK.app : FA;
const {
  getAuth, onAuthStateChanged, signInWithEmailAndPassword, signOut, createUserWithEmailAndPassword,
  EmailAuthProvider, reauthenticateWithCredential, updatePassword, setPersistence, inMemoryPersistence,
} = DEMO ? FK.auth : AU;
const {
  initializeFirestore, persistentLocalCache, persistentMultipleTabManager,
  doc, collection, getDoc, getDocs, setDoc, updateDoc, deleteDoc, onSnapshot,
  query, where, serverTimestamp, writeBatch, Timestamp,
} = DEMO ? FK.fs : FS;
const HORAS_ALERTA = +C.alertaHorasAndamento || 4;
const XLSX_URL = "https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js";

let fb, auth, db;
if (DEMO) {
  fb = initializeApp({}); auth = getAuth(fb); db = initializeFirestore(fb, {});
} else {
  fb = initializeApp(C.firebase);
  auth = getAuth(fb);
  auth.languageCode = "pt-BR";
  try {
    // Cache local: o app abre e registra atividades mesmo sem internet, e sincroniza depois
    db = initializeFirestore(fb, { localCache: persistentLocalCache({ tabManager: persistentMultipleTabManager() }) });
  } catch (e) {
    db = initializeFirestore(fb, {});
  }
}

/* ================== UTILIDADES ================== */
const $ = (s, el = document) => el.querySelector(s);
const $$ = (s, el = document) => [...el.querySelectorAll(s)];
const app = $("#app");
const esc = s => String(s ?? "").replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
const pad = n => String(n).padStart(2, "0");
const ms = t => t == null ? null : typeof t === "number" ? t : t instanceof Date ? t.getTime() : t.toMillis ? t.toMillis() : null;
const iso = d => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
const hojeISO = () => iso(new Date());
const deISO = s => { const [a, m, d] = String(s).split("-").map(Number); return new Date(a, m - 1, d); };
const somaDias = (s, n) => { const d = deISO(s); d.setDate(d.getDate() + n); return iso(d); };
const dataBR = s => s ? String(s).split("-").reverse().join("/") : "—";
const DIAS = ["domingo", "segunda-feira", "terça-feira", "quarta-feira", "quinta-feira", "sexta-feira", "sábado"];
const diaSemana = s => DIAS[deISO(s).getDay()];
const hhmm = t => { const v = ms(t); if (v == null) return "—"; const d = new Date(v); return `${pad(d.getHours())}:${pad(d.getMinutes())}`; };
const fmtDur = min => { if (min == null || isNaN(min)) return "—"; const t = Math.round(min), h = Math.floor(t / 60), m = t % 60; return h ? `${h} h ${pad(m)}` : `${m} min`; };
const fmtN = (n, d = 0) => n == null || isNaN(n) ? "—" : Number(n).toLocaleString("pt-BR", { minimumFractionDigits: d, maximumFractionDigits: d });
const pct = x => x == null || isNaN(x) ? "—" : fmtN(x * 100) + "%";
const num = v => v === "" || v == null || isNaN(Number(v)) ? null : Number(v);
const norm = s => String(s ?? "").normalize("NFD").replace(/[̀-ͯ]/g, "").trim().toUpperCase().replace(/\s+/g, " ");
const normLogin = s => String(s || "").trim().toLowerCase().normalize("NFD").replace(/[̀-ͯ]/g, "").replace(/\s+/g, ".").replace(/[^a-z0-9._-]/g, "");
const primeiroNome = n => String(n || "").trim().split(" ")[0];
const porNome = (a, b) => String(a.nome || "").localeCompare(String(b.nome || ""), "pt-BR");
const porOrdem = (a, b) => ((a.ordem ?? 999) - (b.ordem ?? 999)) || porNome(a, b);
const dados = d => ({ id: d.id, ...d.data({ serverTimestamps: "estimate" }) });
const lista = snap => snap.docs.map(dados);
const getLS = k => { try { return localStorage.getItem(k); } catch (e) { return null; } };
const setLS = (k, v) => { try { v == null ? localStorage.removeItem(k) : localStorage.setItem(k, v); } catch (e) {} };
// Junta a data (AAAA-MM-DD) com a hora (HH:MM) do campo
const quando = (dataIso, hm) => { if (!hm) return null; const [h, m] = hm.split(":").map(Number); const d = deISO(dataIso); d.setHours(h, m, 0, 0); return d; };
const agoraHM = () => { const d = new Date(); return `${pad(d.getHours())}:${pad(d.getMinutes())}`; };
function hash(s) { let h = 2166136261; for (let i = 0; i < s.length; i++) { h ^= s.charCodeAt(i); h = Math.imul(h, 16777619); } return (h >>> 0).toString(36); }

function toast(msg, tempo = 2800) {
  const t = $("#toast"); t.textContent = msg; t.classList.add("show");
  clearTimeout(toast.t); toast.t = setTimeout(() => t.classList.remove("show"), tempo);
}
function ocupado(msg) {
  let o = $(".ocupado");
  if (!msg) { o?.remove(); return; }
  if (!o) { o = document.createElement("div"); o.className = "ocupado"; o.innerHTML = "<div></div>"; document.body.appendChild(o); }
  o.firstChild.textContent = msg;
}
function erroMsg(e, padrao) {
  const c = String(e?.code || ""), m = String(e?.message || "");
  if (c.includes("unavailable") || /network|offline|failed to fetch/i.test(c + m)) return "Sem internet. Tente de novo.";
  if (c.includes("permission-denied")) return "Sem permissão para essa ação.";
  if (/invalid-credential|wrong-password|user-not-found|invalid-login/.test(c)) return "Usuário ou senha incorretos.";
  if (/too-many-requests/.test(c)) return "Muitas tentativas. Aguarde alguns minutos.";
  if (/weak-password/.test(c)) return "Senha fraca: use pelo menos 6 caracteres.";
  if (/operation-not-allowed|admin-restricted/.test(c)) return "O login por e-mail/senha não está ativado no Firebase (veja o LEIA-ME).";
  if (/requires-recent-login/.test(c)) return "Por segurança, saia e entre de novo antes de fazer isso.";
  if (m && !m.startsWith("Firebase")) return m;
  return padrao || "Algo deu errado. Tente de novo.";
}
async function tentar(fn, padrao) {
  try { return await fn(); }
  catch (e) { console.error(e); toast(erroMsg(e, padrao), 4500); return undefined; }
  finally { ocupado(false); }
}
// Gravações simples não esperam o servidor: aparecem na hora e sincronizam quando houver internet
const semEsperar = p => p.catch(e => { console.error(e); toast(erroMsg(e), 4500); });

function modal(html) {
  const m = document.createElement("div");
  m.className = "modal";
  m.innerHTML = `<div class="sheet">${html}</div>`;
  document.body.appendChild(m);
  const fechar = () => m.remove();
  m.addEventListener("click", e => { if (e.target === m || e.target.closest("[data-fechar]")) fechar(); });
  return { el: m, sheet: m.firstChild, fechar };
}
function olhoSenha(id, raiz = document) {   // 👁 mostrar/esconder senha
  const i = $("#" + id, raiz); if (!i) return;
  const w = document.createElement("div"); w.className = "senha";
  i.parentNode.insertBefore(w, i); w.appendChild(i);
  const b = document.createElement("button"); b.type = "button"; b.className = "olho"; b.textContent = "👁"; b.title = "Mostrar senha";
  b.onclick = () => { const m = i.type === "password"; i.type = m ? "text" : "password"; b.textContent = m ? "🙈" : "👁"; };
  w.appendChild(b);
}

/* ================== ESTADO E NAVEGAÇÃO ================== */
const S = {
  user: null, perfil: null, config: null,
  colaboradores: [], locais: [], usuarios: [],
  dia: hojeISO(), atividades: [], aba: "dia", filtroTipo: "todos", sub: "colaboradores",
  periodo: { tipo: "7d" }, filtroH: { texto: "", pessoa: "", local: "", status: "" },
  subs: [], subDia: null, view: null, live: false, cache: null, criandoConta: false,
  metas: null,   // { valores: {localId: {iph, n, listaMedia, listaDow}}, geral, t } — calculado do histórico
  alertas: [],   // alertas do dia (sininho)
};
// Perfis: "supervisor" = Administrador (faz tudo) · "restrito" = só o que estiver marcado em "perms"
const supervisor = () => S.perfil?.papel === "supervisor";
const PERMISSOES = [
  ["lancar", "Lançar e editar atividades", "criar, editar, copiar escala"],
  ["executar", "Iniciar, pausar e finalizar", "e marcar os itens da lista"],
  ["excluir", "Excluir atividades", ""],
  ["relatorios", "Ver Resumo e Histórico", "e baixar Excel / relatório do dia"],
  ["planejar", "Planejar o dia", "criar as atividades do dia seguinte"],
  ["colab", "Cadastrar colaboradores", ""],
  ["locais", "Cadastrar e editar locais", "inclusive a meta de cada local"],
  ["importar", "Importar planilha", ""],
  ["alertas", "Receber alertas", "sininho e notificações"],
  ["tv", "Modo TV", ""],
];
const MODELOS_PERM = {
  lider: { nome: "Líder", perms: ["lancar", "executar", "relatorios", "planejar", "colab", "alertas", "tv"] },
  operador: { nome: "Operador", perms: ["executar"] },
  consulta: { nome: "Só consulta", perms: ["relatorios", "tv"] },
};
// Usuários antigos com perfil "Líder" (sem lista de permissões) continuam com as permissões de líder
const permsDe = u => u?.papel === "supervisor" ? Object.fromEntries(PERMISSOES.map(p => [p[0], true]))
  : u?.perms || (u?.papel === "lider" ? Object.fromEntries(MODELOS_PERM.lider.perms.map(k => [k, true])) : {});
const pode = k => supervisor() || !!permsDe(S.perfil)[k];
const nomePapel = u => u?.papel === "supervisor" ? "Administrador" : (Object.values(MODELOS_PERM).find(m => { const p = permsDe(u); return m.perms.length === Object.keys(p).filter(k => p[k]).length && m.perms.every(k => p[k]); })?.nome || "Restrito");

function go(fn, ...args) { S.view = { fn, args }; S.live = false; window.scrollTo(0, 0); fn(...args); }
let tRender;
function reRender() {
  clearTimeout(tRender);
  // com uma janela aberta, espera ela fechar para redesenhar
  tRender = setTimeout(function f() { if ($(".modal")) { tRender = setTimeout(f, 400); return; } if (S.live && S.view) S.view.fn(...S.view.args); }, 150);
}
function topo(titulo, botoes = []) {
  app.classList.remove("largo");
  $("#hTitle").textContent = titulo;
  const box = $("#hBtns"); box.innerHTML = "";
  botoes.forEach(([txt, fn]) => { const b = document.createElement("button"); b.textContent = txt; b.onclick = fn; box.appendChild(b); });
}
function pararTudo() {
  S.subs.forEach(f => f()); S.subs = [];
  S.subDia?.(); S.subDia = null;
  S.subHoje?.(); S.subHoje = null; S.ativHojeDia = null;
  clearInterval(RELOGIO); RELOGIO = null;
  S.alertasIniciados = false; S.alertasDia = null; S.metas = null;
  document.body.classList.remove("tv");
}
const voltarPainel = () => go(telaPainel);

/* ================== INÍCIO E LOGIN ================== */
// Cada usuário (supervisor ou líder) entra com "usuário" e senha. Por dentro, o usuário vira um
// e-mail técnico, guardado em /logins (o Firebase grátis só aceita e-mail).
const emailDe = login => `u.${login}.${Date.now().toString(36)}@usuarios.atividades.app`;
const loginRef = login => doc(db, "logins", login);

function iniciar() {
  if (!CONFIGURADO && !DEMO) return telaSemConfig();
  if (DEMO) {   // faixa avisando que é demonstração
    const f = document.createElement("div");
    f.className = "demo-bar no-print";
    f.innerHTML = self.ATIV_OFFLINE
      ? `💻 <b>Versão offline</b> · os dados ficam só neste computador (neste navegador) <button type="button">Recomeçar do zero</button>`
      : `👀 <b>Demonstração</b> · dados de exemplo, ficam só neste aparelho${CONFIGURADO ? "" : " · o Firebase ainda não foi configurado"} <button type="button">Recomeçar do zero</button>`;
    f.querySelector("button").onclick = () => { if (confirm("Apagar tudo o que foi feito na demonstração e voltar aos dados de exemplo?")) recomecarDemo(); };
    $("header").insertAdjacentElement("afterend", f);
  }
  onAuthStateChanged(auth, u => {
    S.user = u;
    if (S.criandoConta) return;   // primeiro acesso: entra depois de gravar o cadastro
    if (!u) { pararTudo(); S.perfil = null; return telaEntrada(); }
    entrar();
  });
}

function telaSemConfig() {
  topo("📋 Controle de Atividades");
  app.innerHTML = `<div class="card"><h2>Falta configurar o Firebase</h2>
    <p class="muted">Abra o arquivo <b>config.js</b> e cole os dados do projeto Firebase deste app (passo a passo no LEIA-ME.txt). Depois suba o config.js de novo no GitHub.</p></div>`;
}

async function telaEntrada(msg) {
  msg = msg || S.msgSaida; S.msgSaida = null;
  topo("📋 Controle de Atividades");
  app.innerHTML = `<div class="empty">Carregando…</div>`;
  let cfg = null, erro = false;
  try { const s = await getDoc(doc(db, "config", "sistema")); cfg = s.exists() ? s.data() : null; } catch (e) { erro = true; }
  if (S.user) return;   // entrou enquanto carregava
  S.config = cfg;
  if (!cfg && !erro) return telaPrimeiroAcesso();
  telaLogin(msg);
}

function telaLogin(msg) {
  S.view = null; S.live = false;
  topo(`📋 ${S.config?.empresa || "Controle de Atividades"}`);
  app.innerHTML = `
    ${msg ? `<div class="alert">${esc(msg)}</div>` : ""}
    ${DEMO ? `<div class="alert info" style="max-width:440px;margin:16px auto 0"><b>Demonstração</b>Supervisor: usuário <b>demo</b> · senha <b>123456</b><br>Líder: usuário <b>carlos</b> · senha <b>123456</b></div>` : ""}
    <div class="card" style="max-width:440px;margin:24px auto 14px">
      <div class="center" style="font-size:44px;line-height:1.2">📋</div>
      <h2 class="center">Controle de Atividades</h2>
      <label for="lUser">Usuário</label>
      <input id="lUser" autocomplete="username" autocapitalize="none" autocorrect="off" spellcheck="false">
      <label for="lSenha">Senha</label>
      <input id="lSenha" type="password" autocomplete="current-password">
      <button class="btn" id="lEntrar">Entrar</button>
      <button class="btn sec" id="lEsqueci">Esqueci minha senha</button>
    </div>`;
  olhoSenha("lSenha");
  $("#lUser").oninput = e => { const p = e.target.selectionStart; e.target.value = normLogin(e.target.value); e.target.setSelectionRange?.(p, p); };
  $("#lEsqueci").onclick = () => modal(`<h2>Esqueci minha senha</h2>
    <p class="muted">Peça ao supervisor para criar uma senha nova para você em <b>Cadastros → Usuários do app</b>.</p>
    <button class="btn sec" data-fechar>Fechar</button>`);
  const entrar = async () => {
    const login = normLogin($("#lUser").value), senha = $("#lSenha").value;
    if (!login || !senha) return toast("Preencha usuário e senha");
    ocupado("Entrando…");
    await tentar(async () => {
      const lk = await getDoc(loginRef(login));
      if (!lk.exists()) throw new Error("Usuário não encontrado.");
      await signInWithEmailAndPassword(auth, lk.data().email, senha);   // o resto acontece em onAuthStateChanged
    });
  };
  $("#lEntrar").onclick = entrar;
  $("#lSenha").onkeydown = e => { if (e.key === "Enter") entrar(); };
  $("#lUser").onkeydown = e => { if (e.key === "Enter") $("#lSenha").focus(); };
}

// Primeiro acesso: quem configura o app vira o primeiro supervisor
function telaPrimeiroAcesso() {
  S.view = null; S.live = false;
  topo("📋 Primeiro acesso");
  app.innerHTML = `
    <div class="card" style="max-width:520px;margin:16px auto">
      <h2>👋 Bem-vindo! Vamos configurar o app</h2>
      <p class="muted">Você será o <b>administrador</b>: cadastra colaboradores, locais e os outros usuários (administradores ou com restrições).</p>
      <label for="pEmpresa">Nome da empresa / setor</label>
      <input id="pEmpresa" maxlength="60" placeholder="Ex.: Recebimento - CD">
      <label for="pNome">Seu nome</label>
      <input id="pNome" maxlength="60" autocomplete="name">
      <label for="pUser">Crie seu usuário</label>
      <input id="pUser" maxlength="30" autocapitalize="none" autocorrect="off" spellcheck="false" placeholder="Ex.: fellippe">
      <label for="pSenha">Crie uma senha (mínimo 6 caracteres)</label>
      <input id="pSenha" type="password" autocomplete="new-password">
      <label for="pSenha2">Repita a senha</label>
      <input id="pSenha2" type="password" autocomplete="new-password">
      <label style="display:flex;gap:8px;align-items:center;color:var(--text)"><input type="checkbox" id="pLocais" checked style="width:auto;margin:0">
        Já criar os locais da planilha (Rua A a I, 1º a 3º Jirau e Sala Nobre)</label>
      <button class="btn" id="pOk" style="margin-top:12px">Começar</button>
    </div>`;
  olhoSenha("pSenha"); olhoSenha("pSenha2");
  $("#pUser").oninput = e => { e.target.value = normLogin(e.target.value); };
  $("#pOk").onclick = async () => {
    const empresa = $("#pEmpresa").value.trim(), nome = $("#pNome").value.trim().replace(/\s+/g, " ");
    const login = normLogin($("#pUser").value), senha = $("#pSenha").value;
    if (nome.length < 2) return toast("Digite seu nome");
    if (login.length < 3) return toast("O usuário precisa ter pelo menos 3 letras/números");
    if (senha.length < 6) return toast("A senha precisa ter pelo menos 6 caracteres");
    if (senha !== $("#pSenha2").value) return toast("As duas senhas não são iguais");
    const semear = $("#pLocais").checked;
    ocupado("Configurando…");
    S.criandoConta = true;
    const ok = await tentar(async () => {
      const email = emailDe(login);
      const u = (await createUserWithEmailAndPassword(auth, email, senha)).user;
      const b = writeBatch(db);
      b.set(doc(db, "usuarios", u.uid), { nome, login, papel: "supervisor", ativo: true, criadoEm: serverTimestamp() });
      b.set(loginRef(login), { email, uid: u.uid });
      b.set(doc(db, "config", "sistema"), { dono: u.uid, empresa, criadoEm: serverTimestamp() });
      await b.commit();
      if (semear) await semearLocais();
      return true;
    }, "Não foi possível configurar. Tente de novo.");
    S.criandoConta = false;
    if (ok && S.user) { toast(`Pronto! Seu usuário é: ${login}`, 5000); entrar(); }
  };
}

async function entrar() {
  ocupado("Carregando…");
  let snap;
  try { snap = await getDoc(doc(db, "usuarios", S.user.uid)); }
  catch (e) { ocupado(false); return telaSemConexao(); }
  ocupado(false);
  if (!snap.exists() || snap.data().ativo === false) return sair("Seu acesso não está ativo. Fale com o supervisor.");
  S.perfil = { id: snap.id, ...snap.data() };
  ouvirBase();
  ouvirDia();
  S.aba = "dia";
  go(telaPainel);
  setTimeout(atualizarMetas, 1500);   // supervisor recalcula as metas 1x por dia
  iniciarRelogio();
}
function telaSemConexao() {
  topo("📋 Controle de Atividades");
  app.innerHTML = `<div class="card center"><h2>Sem conexão</h2><p class="muted">Não foi possível carregar seus dados. Confira a internet.</p>
    <button class="btn" id="tNovo">Tentar de novo</button></div>`;
  $("#tNovo").onclick = () => entrar();
}
async function sair(msg) {
  pararTudo();
  S.perfil = null; S.msgSaida = msg || null; S.cache = null;
  await signOut(auth).catch(() => {});   // onAuthStateChanged leva para a tela de login
}

/* ---------- Minha conta (botão com o nome, no topo) ---------- */
function minhaConta() {
  const p = S.perfil || {};
  const m = modal("");
  const menu = () => {
    m.sheet.innerHTML = `<h2>👤 ${esc(p.nome)}</h2>
      <p class="muted">Usuário <b>${esc(p.login)}</b> · ${p.papel === "supervisor" ? "👑 Administrador" : "🔒 " + esc(nomePapel(p))}</p>
      ${pode("alertas") ? `<label style="display:flex;gap:8px;align-items:center;color:var(--text);margin:6px 0 12px"><input type="checkbox" id="mcAlertas" ${recebeAlertas() ? "checked" : ""} style="width:auto;margin:0"> 🔔 Receber alertas neste aparelho</label>` : ""}
      ${"Notification" in window && Notification.permission === "default" ? `<button class="btn sec" id="mcNotif">📲 Ativar notificações do aparelho</button>` : ""}
      <button class="btn sec" id="mcSenha">🔑 Trocar minha senha</button>
      <button class="btn bad" id="mcSair">Sair</button>
      <button class="btn sec" data-fechar>Fechar</button>`;
    if ($("#mcAlertas", m.sheet)) $("#mcAlertas", m.sheet).onchange = e => { setLS("ativ_alertas", e.target.checked ? "1" : "0"); toast(e.target.checked ? "Alertas ligados neste aparelho" : "Alertas desligados neste aparelho"); checarAlertas(); if (S.view?.fn === telaPainel) telaPainel(); };
    $("#mcNotif", m.sheet) && ($("#mcNotif", m.sheet).onclick = () => Notification.requestPermission().then(p => toast(p === "granted" ? "Notificações ativadas" : "As notificações foram bloqueadas no navegador")));
    $("#mcSenha", m.sheet).onclick = formSenha;
    $("#mcSair", m.sheet).onclick = () => { if (confirm("Sair do app neste aparelho?")) { m.fechar(); sair(); } };
  };
  const formSenha = () => {
    m.sheet.innerHTML = `<h2>🔑 Trocar senha</h2>
      <label for="mcAtual">Senha atual</label><input id="mcAtual" type="password" autocomplete="current-password">
      <label for="mcNova">Nova senha (mínimo 6 caracteres)</label><input id="mcNova" type="password" autocomplete="new-password">
      <label for="mcNova2">Repita a nova senha</label><input id="mcNova2" type="password" autocomplete="new-password">
      <button class="btn" id="mcOk">Salvar nova senha</button>
      <button class="btn sec" id="mcVoltar">Voltar</button>`;
    ["mcAtual", "mcNova", "mcNova2"].forEach(id => olhoSenha(id, m.sheet));
    $("#mcVoltar", m.sheet).onclick = menu;
    $("#mcOk", m.sheet).onclick = async () => {
      const atual = $("#mcAtual", m.sheet).value, nova = $("#mcNova", m.sheet).value;
      if (!atual) return toast("Digite a senha atual");
      if (nova.length < 6) return toast("A nova senha precisa ter pelo menos 6 caracteres");
      if (nova !== $("#mcNova2", m.sheet).value) return toast("As duas senhas novas não são iguais");
      ocupado("Trocando a senha…");
      const ok = await tentar(async () => {
        await reauthenticateWithCredential(auth.currentUser, EmailAuthProvider.credential(auth.currentUser.email, atual));
        await updatePassword(auth.currentUser, nova);
        return true;
      });
      if (ok) { m.fechar(); toast("Senha alterada!", 4000); }
    };
  };
  menu();
}

/* ================== DADOS EM TEMPO REAL ================== */
function ouvirBase() {
  S.subs.forEach(f => f()); S.subs = [];
  const ouvir = (ref, fn) => S.subs.push(onSnapshot(ref, fn, e => {
    console.warn("tempo real:", e);
    if (e?.code === "permission-denied" && S.perfil) sair("Seu acesso foi alterado pelo supervisor. Entre de novo.");
  }));
  ouvir(doc(db, "usuarios", S.user.uid), s => {
    if (!s.exists() || s.data().ativo === false) { if (!s.metadata.fromCache) sair("Seu acesso foi alterado ou desativado pelo supervisor."); return; }
    S.perfil = { id: s.id, ...s.data() }; reRender();
  });
  ouvir(doc(db, "config", "sistema"), s => { S.config = s.exists() ? s.data() : null; reRender(); });
  ouvir(collection(db, "colaboradores"), s => { S.colaboradores = lista(s).sort(porNome); reRender(); });
  ouvir(collection(db, "locais"), s => { S.locais = lista(s).sort(porOrdem); S.locaisOk = true; reRender(); });
  ouvir(collection(db, "usuarios"), s => { S.usuarios = lista(s).sort(porNome); reRender(); });
  ouvir(doc(db, "config", "metas"), s => { S.metas = s.exists() ? s.data() : null; reRender(); });
}
function ouvirDia() {
  S.subDia?.();
  const dia = S.dia;
  S.subDia = onSnapshot(query(collection(db, "atividades"), where("data", "==", dia)), s => {
    if (dia !== S.dia) return;
    S.atividades = lista(s); reRender();
  }, e => console.warn("tempo real (dia):", e));
}
function mudarDia(novo) {
  if (!novo || novo === S.dia) return;
  S.dia = novo; S.atividades = [];
  ouvirDia();
  go(telaPainel);
}

/* ================== CÁLCULOS (as fórmulas da planilha) ================== */
// Tempo = término − início − pausas · Pendentes = máx(lista − abastecidos, 0) · Itens/hora = abastecidos ÷ horas trabalhadas
const MOTIVOS_PAUSA = ["Almoço / intervalo", "Aguardando material", "Sem equipamento", "Reunião / DDS", "Ajudando outro local", "Outro"];
const MOTIVOS_FALTA = ["Sem estoque", "Sem espaço no local", "Avaria", "Não encontrado", "Outro"];
const pausaAberta = a => { const p = (a.pausas || []).at(-1); return p && p.fim == null ? p : null; };
// minutos parados (pausas) dentro da atividade, até "ate" (término ou agora)
const minutosParado = (a, ate) => (a.pausas || []).reduce((s, p) => {
  const i = ms(p.ini), f = ms(p.fim) ?? ate;
  return i != null && f != null && f > i ? s + (f - i) / 60000 : s;
}, 0);
// Meta de itens/hora do local: a do cadastro (se o supervisor definiu) ou a calculada pelo histórico
function metaDe(localId) {
  const l = S.locais.find(x => x.id === localId);
  if (num(l?.metaIph) > 0) return num(l.metaIph);
  return S.metas?.valores?.[localId]?.iph || null;
}
function calc(a) {
  const ini = ms(a.inicio), fim = ms(a.fim);
  const hoje = a.data === hojeISO();
  const parado = ini != null ? minutosParado(a, fim ?? (hoje ? Date.now() : ini)) : 0;
  const bruto = ini != null && fim != null && fim > ini ? (fim - ini) / 60000 : null;
  const min = bruto != null ? Math.max(bruto - parado, 0) : null;   // tempo trabalhado (sem as pausas)
  const lista = num(a.itensLista), abast = num(a.itensAbastecidos);
  const pend = lista != null && abast != null ? Math.max(lista - abast, 0) : null;
  const extra = lista != null && abast != null ? Math.max(abast - lista, 0) : null;
  const iph = min && abast != null ? abast / (min / 60) : null;
  const status = fim != null ? "concluida" : ini != null ? (pausaAberta(a) ? "pausada" : "andamento") : "planejada";
  const meta = metaDe(a.localId);
  const pctMeta = iph != null && meta ? iph / meta : null;
  // em andamento hoje: tempo corrido trabalhado até agora
  const corrido = ini != null && fim == null && hoje ? Math.max((Date.now() - ini) / 60000 - parado, 0) : null;
  const faltas = (a.itens || []).filter(i => i.status === "falta");
  return { ini, fim, min, bruto, parado, lista, abast, pend, extra, iph, status, meta, pctMeta, corrido, faltas };
}
function totais(arr) {
  const t = { registros: arr.length, concluidas: 0, andamento: 0, pausadas: 0, planejadas: 0, lista: 0, abast: 0, pend: 0, extra: 0, min: 0, parado: 0, abastT: 0, minT: 0, listaCob: 0, cobertos: 0, esperado: 0, realMeta: 0 };
  const pessoas = new Set(), dias = new Set(), porMotivo = {};
  arr.forEach(a => {
    const c = calc(a);
    if (c.status === "concluida") t.concluidas++; else if (c.status === "andamento") t.andamento++; else if (c.status === "pausada") t.pausadas++; else t.planejadas++;
    if (c.lista != null) t.lista += c.lista;
    if (c.abast != null) t.abast += c.abast;
    if (c.pend) t.pend += c.pend;
    if (c.extra) t.extra += c.extra;
    if (c.min) t.min += c.min;
    t.parado += c.parado || 0;
    (a.pausas || []).forEach(p => { const f = ms(p.fim) ?? (c.fim ?? (a.data === hojeISO() ? Date.now() : null)); const i = ms(p.ini); if (i != null && f > i) porMotivo[p.motivo || "Outro"] = (porMotivo[p.motivo || "Outro"] || 0) + (f - i) / 60000; });
    // produtividade: só atividades com tempo E itens abastecidos
    if (c.min && c.abast != null) {
      t.abastT += c.abast; t.minT += c.min;
      if (c.meta) { t.esperado += c.meta * c.min / 60; t.realMeta += c.abast; }   // itens que a meta pedia × o que foi feito
    }
    // % da lista: itens da lista que foram atendidos (o que passa da lista não conta)
    if (c.lista && c.abast != null) { t.listaCob += c.lista; t.cobertos += Math.min(c.abast, c.lista); }
    (a.pessoas || []).forEach(p => pessoas.add(p.id || norm(p.nome)));
    dias.add(a.data);
  });
  t.iph = t.minT ? t.abastT / (t.minT / 60) : null;
  t.cobertura = t.listaCob ? t.cobertos / t.listaCob : null;
  t.pctMeta = t.esperado ? t.realMeta / t.esperado : null;
  t.porMotivo = porMotivo;
  t.nPessoas = pessoas.size; t.nDias = dias.size;
  return t;
}
const corMeta = p => p == null ? "var(--muted)" : p >= 1 ? "var(--ok)" : p >= 0.8 ? "var(--warn)" : "var(--bad)";
const pillMeta = p => p == null ? "" : `<span class="pill ${p >= 1 ? "ok" : p >= 0.8 ? "warn" : "bad"}">🎯 ${pct(p)} da meta</span>`;

/* ---------- Metas automáticas (histórico das últimas 4 semanas) ---------- */
// O supervisor recalcula 1x por dia e grava em config/metas; os outros só leem esse resumo (economiza leituras)
async function atualizarMetas(forcar) {
  if (!supervisor() || !S.perfil) return;
  if (!forcar && S.metas?.t && Date.now() - ms(S.metas.t) < 20 * 3600000) return;
  const ate = somaDias(hojeISO(), -1), de = somaDias(hojeISO(), -28);
  const snap = await getDocs(query(collection(db, "atividades"), where("data", ">=", de), where("data", "<=", ate))).catch(() => null);
  if (!snap) return;
  const acc = {}, geral = { abast: 0, min: 0 };
  lista(snap).forEach(a => {
    const c = calc(a), k = a.localId;
    if (!k) return;
    const r = acc[k] = acc[k] || { abast: 0, min: 0, n: 0, listas: [], dows: {} };
    if (c.min && c.abast != null) { r.abast += c.abast; r.min += c.min; r.n++; geral.abast += c.abast; geral.min += c.min; }
    if (c.lista != null) { r.listas.push(c.lista); const d = deISO(a.data).getDay(); (r.dows[d] = r.dows[d] || []).push(c.lista); }
  });
  const media = v => v.length ? Math.round(v.reduce((s, x) => s + x, 0) / v.length) : null;
  const valores = {};
  Object.entries(acc).forEach(([k, r]) => {
    valores[k] = {
      iph: r.n >= 3 && r.min ? Math.round(r.abast / (r.min / 60) * 100) / 100 : null,   // precisa de pelo menos 3 atividades
      n: r.n, listaMedia: media(r.listas),
      listaDow: Object.fromEntries(Object.entries(r.dows).map(([d, v]) => [d, media(v)])),
    };
  });
  await setDoc(doc(db, "config", "metas"), {
    t: Timestamp.now(), de, ate, valores,
    geral: geral.min ? Math.round(geral.abast / (geral.min / 60) * 100) / 100 : null,
  }).catch(e => console.warn("metas:", e));
}
const localDe = a => S.locais.find(l => l.id === a.localId);
const ordemDe = a => localDe(a)?.ordem ?? 999;
function porLocal(arr) {
  const m = new Map();
  arr.forEach(a => {
    const k = a.localId || norm(a.localNome);
    if (!m.has(k)) m.set(k, { k, nome: a.localNome, tipo: a.tipo, ordem: ordemDe(a), local: localDe(a), itens: [] });
    m.get(k).itens.push(a);
  });
  return [...m.values()].map(g => ({ ...g, t: totais(g.itens) })).sort((a, b) => (a.ordem - b.ordem) || String(a.nome).localeCompare(b.nome, "pt-BR"));
}
// Por colaborador: quando a atividade tem mais de uma pessoa, os itens são divididos entre elas
function porColaborador(arr) {
  const m = new Map();
  arr.forEach(a => {
    const c = calc(a), ps = a.pessoas || [], n = ps.length || 1;
    ps.forEach(p => {
      const k = p.id || norm(p.nome);
      if (!m.has(k)) m.set(k, { nome: p.nome, ativ: 0, abast: 0, pend: 0, min: 0, abastT: 0, minT: 0, locais: new Set(), dias: new Set() });
      const r = m.get(k);
      r.ativ++; r.locais.add(a.localNome); r.dias.add(a.data);
      if (c.abast != null) r.abast += c.abast / n;
      if (c.pend) r.pend += c.pend / n;
      if (c.min) r.min += c.min;
      if (c.min && c.abast != null) { r.abastT += c.abast / n; r.minT += c.min; }
    });
  });
  return [...m.values()].map(r => ({ ...r, iph: r.minT ? r.abastT / (r.minT / 60) : null })).sort((a, b) => b.abast - a.abast || porNome(a, b));
}
function porDia(arr) {
  const m = new Map();
  arr.forEach(a => { if (!m.has(a.data)) m.set(a.data, []); m.get(a.data).push(a); });
  return [...m.entries()].sort((a, b) => a[0].localeCompare(b[0])).map(([d, itens]) => ({ data: d, t: totais(itens) }));
}

/* ================== PAINEL ================== */
const ABAS = [["dia", "📅", "Dia"], ["resumo", "📊", "Resumo"], ["planejar", "🗓️", "Planejar"], ["historico", "🗂️", "Histórico"], ["cadastros", "⚙️", "Cadastros"]];
function telaPainel() {
  const nAlertas = S.alertas.filter(x => !x.visto).length;
  topo(`📋 ${S.config?.empresa || "Controle de Atividades"}`, [
    ...(recebeAlertas() ? [[`🔔${nAlertas ? " " + nAlertas : ""}`, abrirAlertas]] : []),
    ["👤 " + primeiroNome(S.perfil?.nome || "Conta"), minhaConta],
  ]);
  app.classList.add("largo");
  // cada usuário só vê as abas que pode usar
  const temCadastros = supervisor() || pode("colab") || pode("locais") || pode("importar");
  const abas = ABAS.filter(([k]) => k === "dia" || ((k === "resumo" || k === "historico") && pode("relatorios")) || (k === "planejar" && pode("planejar")) || (k === "cadastros" && temCadastros));
  if (!abas.some(a => a[0] === S.aba)) S.aba = "dia";
  S.live = S.aba === "dia" || S.aba === "cadastros";
  app.innerHTML = `<div class="painel">
    <nav class="menu">${abas.map(([k, ic, t]) => `<button data-aba="${k}" class="${S.aba === k ? "on" : ""}"><span class="ic">${ic}</span><span>${t}</span></button>`).join("")}
      ${pode("tv") ? `<button data-tv="1" class="so-pc"><span class="ic">📺</span><span>Modo TV</span></button>` : ""}</nav>
    <section id="aba"></section></div>`;
  $$(".menu button[data-aba]").forEach(b => b.onclick = () => { S.aba = b.dataset.aba; go(telaPainel); });
  $(".menu [data-tv]") && ($(".menu [data-tv]").onclick = abrirTV);
  ({ dia: abaDia, resumo: abaResumo, planejar: abaPlanejar, historico: abaHistorico, cadastros: abaCadastros }[S.aba] || abaDia)($("#aba"));
}

const kpi = (n, t, s, cor) => `<div class="kpi" style="--cor:${cor}"><div class="n">${n}</div><div class="t">${t}</div>${s ? `<div class="s">${s}</div>` : ""}</div>`;
function kpisHtml(t, periodo) {
  return `<div class="kpis">
    ${kpi(t.registros, "Registros", periodo ? `${t.nDias} dia(s) com registro` : `${t.concluidas} concluída(s) · ${t.andamento + t.pausadas} em andamento`, "var(--primary)")}
    ${kpi(t.nPessoas, "Pessoas", "colaboradores diferentes", "#8250df")}
    ${kpi(fmtN(t.lista), "Itens da lista", "", "var(--info)")}
    ${kpi(fmtN(t.abast), "Itens abastecidos", t.extra ? `+${fmtN(t.extra)} além da lista` : "", "var(--ok)")}
    ${kpi(fmtN(t.pend), "Itens pendentes", t.pend ? "ficaram faltando" : "nada pendente", t.pend ? "var(--bad)" : "var(--ok)")}
    ${kpi(fmtN(t.iph, 1), "Itens por hora", "produtividade média", "var(--warn)")}
    ${kpi(pct(t.pctMeta), "Da meta", t.pctMeta == null ? "sem meta ainda" : t.pctMeta >= 1 ? "acima da meta 🎉" : "abaixo da meta", corMeta(t.pctMeta))}
    ${kpi(fmtDur(t.min), "Tempo trabalhado", t.parado >= 1 ? `+ ${fmtDur(t.parado)} parado` : "sem pausas", "#8250df")}
    ${kpi(pct(t.cobertura), "Lista abastecida", "itens da lista atendidos", "var(--primary)")}
  </div>`;
}

/* ---------- Aba Dia ---------- */
const COR_ST = { planejada: "var(--line)", andamento: "var(--info)", pausada: "var(--warn)", concluida: "var(--ok)" };
function pillStatus(c, a) {
  if (c.status === "concluida") return `<span class="pill ok">Concluída</span>`;
  if (c.status === "pausada") return `<span class="pill warn">⏸ Pausada</span>`;
  if (c.status === "andamento") return a.data < hojeISO() ? `<span class="pill warn">Sem término</span>` : `<span class="pill azul">Em andamento</span>`;
  return `<span class="pill off">Planejada</span>`;
}
function cardAtiv(a) {
  const c = calc(a);
  const nomes = (a.pessoas || []).map(p => esc(p.nome)).join(", ") || `<i>sem colaborador</i>${a.sugeridas ? ` · 👥 ${a.sugeridas} sugerida(s)` : ""}`;
  const hoje = a.data === hojeISO();
  const pa = pausaAberta(a);
  const nItens = (a.itens || []).length, okItens = (a.itens || []).filter(i => i.status === "ok").length;
  return `<div class="ativ" style="--cor:${c.pend ? "var(--bad)" : COR_ST[c.status]}">
    <div class="topo"><div><span class="local">${esc(a.localNome)}</span> · <span class="pessoas">${nomes}</span></div>${pillStatus(c, a)}</div>
    <div class="nums">
      <span>🕒 <b>${hhmm(c.ini)}</b> → <b>${c.fim != null ? hhmm(c.fim) : "…"}</b>${c.min ? ` · ${fmtDur(c.min)}` : ""}${c.corrido != null ? ` · ${fmtDur(c.corrido)} trabalhando` : ""}</span>
      ${c.parado >= 1 ? `<span>⏸ ${fmtDur(c.parado)} parado</span>` : ""}
      ${c.lista != null ? `<span>Lista <b>${c.lista}</b></span>` : ""}
      ${c.abast != null ? `<span>Abastecidos <b>${c.abast}</b></span>` : ""}
      ${c.pend ? `<span style="color:var(--bad)">Pendentes <b style="color:var(--bad)">${c.pend}</b></span>` : ""}
      ${c.extra ? `<span>+${c.extra} além da lista</span>` : ""}
      ${c.iph != null ? `<span><b>${fmtN(c.iph, 1)}</b> itens/h</span>` : c.meta ? `<span>meta ${fmtN(c.meta, 1)}/h</span>` : ""}
      ${pillMeta(c.pctMeta)}
    </div>
    ${pa ? `<div class="obs" style="background:var(--warn-bg);color:var(--warn)">⏸ Parado desde ${hhmm(pa.ini)} · <b>${esc(pa.motivo)}</b></div>` : ""}
    ${c.faltas.length ? `<div class="obs" style="background:var(--bad-bg);color:var(--bad)">🔻 Faltou: ${c.faltas.map(i => `${esc(i.desc)} (${esc(i.motivo || "—")})`).join(", ")}</div>` : a.motivoFalta && c.pend ? `<div class="obs" style="background:var(--bad-bg);color:var(--bad)">🔻 Motivo da falta: ${esc(a.motivoFalta)}</div>` : ""}
    ${a.obs ? `<div class="obs">💬 ${esc(a.obs)}</div>` : ""}
    <div class="acts">
      ${pode("executar") ? `
      ${c.status === "planejada" ? `<button class="mini pri" data-iniciar="${a.id}">▶ Iniciar</button>` : ""}
      ${c.status === "andamento" && hoje ? `<button class="mini" data-pausar="${a.id}">⏸ Pausar</button>` : ""}
      ${c.status === "pausada" ? `<button class="mini pri" data-retomar="${a.id}">▶ Retomar</button>` : ""}
      ${c.status === "andamento" || c.status === "pausada" ? `<button class="mini ok" data-finalizar="${a.id}">■ Finalizar</button>` : ""}` : ""}
      ${nItens && (pode("executar") || pode("lancar")) ? `<button class="mini" data-itens="${a.id}">☑️ Itens ${okItens}/${nItens}</button>` : ""}
      ${pode("lancar") || pode("excluir") ? `<button class="mini" data-editar="${a.id}">✏️ Editar</button>` : ""}
    </div></div>`;
}
function abaDia(el) {
  const hoje = S.dia === hojeISO();
  const tipos = [...new Set([...S.locais.filter(l => l.ativo !== false).map(l => l.tipo), ...S.atividades.map(a => a.tipo)].filter(Boolean))];
  if (S.filtroTipo !== "todos" && !tipos.includes(S.filtroTipo)) S.filtroTipo = "todos";
  const doFiltro = a => S.filtroTipo === "todos" || a.tipo === S.filtroTipo;
  const ativs = S.atividades.filter(doFiltro).sort((a, b) => (ordemDe(a) - ordemDe(b)) || String(a.localNome).localeCompare(b.localNome, "pt-BR") || ((ms(a.inicio) ?? 9e15) - (ms(b.inicio) ?? 9e15)));
  const t = totais(S.atividades);
  const longas = hoje ? S.atividades.filter(a => { const c = calc(a); return (c.status === "andamento" || c.status === "pausada") && (Date.now() - c.ini) / 3600000 > HORAS_ALERTA; }) : [];
  const usados = new Set(S.atividades.map(a => a.localId));
  const semReg = S.locais.filter(l => l.ativo !== false && !usados.has(l.id) && doFiltro(l));

  // agrupa por tipo, na ordem dos locais
  const grupos = new Map();
  ativs.forEach(a => { const k = a.tipo || "Outros"; if (!grupos.has(k)) grupos.set(k, []); grupos.get(k).push(a); });

  el.innerHTML = `
    <div class="navdia">
      <button class="mini" id="dAnt" title="Dia anterior">◀</button>
      <input type="date" id="dData" value="${S.dia}">
      <button class="mini" id="dProx" title="Próximo dia">▶</button>
      ${hoje ? "" : `<button class="mini pri" id="dHoje">Hoje</button>`}
    </div>
    <div class="muted" style="margin:-6px 0 12px">${diaSemana(S.dia)}, ${dataBR(S.dia)}${hoje ? " · hoje" : ""}</div>
    ${S.locais.length || !S.locaisOk ? "" : `<div class="alert info"><b>Comece cadastrando os locais</b>${pode("locais") ? `Vá em <b>Cadastros → Locais</b> (tem um botão que já cria os locais da planilha).` : "Peça ao administrador para cadastrar os locais."}</div>`}
    ${kpisHtml(t)}
    ${longas.length ? `<div class="alert"><b>⏰ ${longas.length} atividade(s) em andamento há mais de ${HORAS_ALERTA} h</b>${longas.map(a => `${esc(a.localNome)} · ${esc((a.pessoas || []).map(p => p.nome).join(", "))} (desde ${hhmm(a.inicio)})`).join("<br>")}</div>` : ""}
    ${pode("lancar") || pode("relatorios") ? `<div class="linha-btns">
      ${pode("lancar") ? `<button class="btn" id="dNova">＋ Nova atividade</button>
      <button class="btn sec" id="dCopiar">📋 Copiar escala de outro dia</button>` : ""}
      ${pode("relatorios") ? `<button class="btn sec" id="dRel">📤 Relatório do dia</button>` : ""}
    </div>` : ""}
    ${tipos.length > 1 ? `<div class="filtros">${["todos", ...tipos].map(k => `<button data-tipo="${esc(k)}" class="${S.filtroTipo === k ? "on" : ""}">${k === "todos" ? "Todos" : esc(k)}</button>`).join("")}</div>` : ""}
    ${ativs.length ? [...grupos.entries()].map(([tipo, itens]) => `<div class="grupo-tit">${esc(tipo)} · ${itens.length}</div><div class="grade-ativ">${itens.map(cardAtiv).join("")}</div>`).join("")
      : `<div class="card empty">Nenhuma atividade ${hoje ? "hoje" : "neste dia"} ainda.<br>Toque em <b>＋ Nova atividade</b>, em <b>Copiar escala</b> ou monte o dia em <b>🗓️ Planejar</b>.</div>`}
    ${semReg.length && pode("lancar") ? `<div class="card" style="margin-top:14px"><h2>Locais sem registro ${hoje ? "hoje" : "neste dia"}</h2>
      <div class="sem-reg">${semReg.map(l => `<button class="mini" data-novo-local="${l.id}">＋ ${esc(l.nome)}</button>`).join("")}</div></div>` : ""}`;

  const achar = id => S.atividades.find(a => a.id === id);
  $("#dAnt").onclick = () => mudarDia(somaDias(S.dia, -1));
  $("#dProx").onclick = () => mudarDia(somaDias(S.dia, 1));
  $("#dData").onchange = e => mudarDia(e.target.value);
  $("#dHoje") && ($("#dHoje").onclick = () => mudarDia(hojeISO()));
  $("#dNova") && ($("#dNova").onclick = () => go(telaAtividade, null, { data: S.dia }));
  $("#dCopiar") && ($("#dCopiar").onclick = copiarEscala);
  $("#dRel") && ($("#dRel").onclick = () => modalRelatorio(S.dia, S.atividades));
  $$("[data-tipo]", el).forEach(b => b.onclick = () => { S.filtroTipo = b.dataset.tipo; telaPainel(); });
  $$("[data-novo-local]", el).forEach(b => b.onclick = () => go(telaAtividade, null, { data: S.dia, localId: b.dataset.novoLocal }));
  $$("[data-editar]", el).forEach(b => b.onclick = () => go(telaAtividade, achar(b.dataset.editar)));
  $$("[data-finalizar]", el).forEach(b => b.onclick = () => modalFinalizar(achar(b.dataset.finalizar)));
  $$("[data-pausar]", el).forEach(b => b.onclick = () => modalPausar(achar(b.dataset.pausar)));
  $$("[data-retomar]", el).forEach(b => b.onclick = () => retomar(achar(b.dataset.retomar)));
  $$("[data-itens]", el).forEach(b => b.onclick = () => go(telaItens, b.dataset.itens));
  $$("[data-iniciar]", el).forEach(b => b.onclick = () => {
    const a = achar(b.dataset.iniciar);
    if (a.data !== hojeISO()) return pode("lancar") ? go(telaAtividade, a) : toast("Só dá para iniciar atividades de hoje");   // dia passado/futuro: informa a hora no formulário
    if (!(a.pessoas || []).length) { toast("Escolha quem vai fazer esta atividade"); return pode("lancar") ? go(telaAtividade, a) : null; }
    semEsperar(updateDoc(doc(db, "atividades", a.id), { inicio: Timestamp.now(), atualizadoEm: serverTimestamp(), atualizadoPor: S.perfil.nome }));
    S.cache = null;
    toast(`▶ Início registrado às ${agoraHM()}`);
  });
}

/* ---------- Pausas com motivo ---------- */
function modalPausar(a) {
  if (!a) return;
  const m = modal(`<h2>⏸ Pausar · ${esc(a.localNome)}</h2>
    <p class="muted">Por que a atividade parou? O tempo parado não conta na produtividade.</p>
    <div class="chips" id="pMot">${MOTIVOS_PAUSA.map(x => `<button type="button" class="chip" data-m="${esc(x)}">${esc(x)}</button>`).join("")}</div>
    <input id="pOutro" class="hidden" maxlength="60" placeholder="Qual o motivo?">
    <button class="btn warn" id="pOk">Pausar</button>
    <button class="btn sec" data-fechar>Cancelar</button>`);
  let motivo = "";
  $$("#pMot .chip", m.sheet).forEach(c => c.onclick = () => {
    motivo = c.dataset.m;
    $$("#pMot .chip", m.sheet).forEach(x => x.classList.toggle("on", x === c));
    $("#pOutro", m.sheet).classList.toggle("hidden", motivo !== "Outro");
  });
  $("#pOk", m.sheet).onclick = () => {
    if (!motivo) return toast("Escolha o motivo");
    const texto = motivo === "Outro" ? ($("#pOutro", m.sheet).value.trim() || "Outro") : motivo;
    const pausas = [...(a.pausas || []), { ini: Timestamp.now(), fim: null, motivo: texto }];
    semEsperar(updateDoc(doc(db, "atividades", a.id), { pausas, atualizadoEm: serverTimestamp(), atualizadoPor: S.perfil.nome }));
    S.cache = null;
    m.fechar();
    toast(`⏸ Pausada: ${texto}`);
  };
}
// fecha a pausa aberta (no horário informado ou agora)
const fecharPausa = (a, quandoMs) => (a.pausas || []).map(p => p.fim == null ? { ...p, fim: Timestamp.fromMillis(Math.max(quandoMs, ms(p.ini))) } : p);
function retomar(a) {
  if (!a) return;
  semEsperar(updateDoc(doc(db, "atividades", a.id), { pausas: fecharPausa(a, Date.now()), atualizadoEm: serverTimestamp(), atualizadoPor: S.perfil.nome }));
  S.cache = null;
  toast("▶ Atividade retomada");
}

// Quadro com as contas feitas na hora (igual às colunas calculadas da planilha)
function calcBox(c) {
  return `<div><b>${fmtDur(c.min)}</b><span>Tempo${c.parado >= 1 ? ` (−${fmtDur(c.parado)} parado)` : ""}</span></div>
    <div><b style="${c.pend ? "color:var(--bad)" : ""}">${c.pend ?? "—"}</b><span>Pendentes</span></div>
    <div><b>${c.extra ?? "—"}</b><span>Além da lista</span></div>
    <div><b style="color:${corMeta(c.pctMeta)}">${fmtN(c.iph, 1)}</b><span>Itens/hora${c.meta ? ` · meta ${fmtN(c.meta, 1)}` : ""}</span></div>`;
}
// Lê um número inteiro ≥ 0 do campo (vazio = null; inválido = NaN)
const lerInt = v => { v = String(v).trim(); if (!v) return null; const n = Number(v); return Number.isInteger(n) && n >= 0 ? n : NaN; };

function modalFinalizar(a) {
  if (!a) return;
  const c0 = calc(a);
  const temItens = (a.itens || []).length > 0;
  const okItens = (a.itens || []).filter(i => i.status === "ok").length;
  const m = modal(`<h2>■ Finalizar · ${esc(a.localNome)}</h2>
    <p class="muted">${esc((a.pessoas || []).map(p => p.nome).join(", "))} · início ${hhmm(c0.ini)}${c0.status === "pausada" ? " · a pausa será encerrada" : ""}</p>
    <label for="fFim">Término</label>
    <div class="campo-agora"><input type="time" id="fFim" value="${a.data === hojeISO() ? agoraHM() : ""}"><button class="mini" id="fAgora" type="button">Agora</button></div>
    ${temItens ? `<div class="alert info"><b>Lista item a item</b>${okItens} abastecido(s), ${c0.faltas.length} falta(s), ${(a.itens || []).filter(i => i.status === "pendente").length} sem marcar (contam como pendentes).</div>` : ""}
    <div class="row">
      <div><label for="fLista">Qtd. itens da lista</label><input id="fLista" type="number" inputmode="numeric" min="0" step="1" value="${c0.lista ?? ""}" ${temItens ? "disabled" : ""}></div>
      <div><label for="fAbast">Itens abastecidos</label><input id="fAbast" type="number" inputmode="numeric" min="0" step="1" value="${temItens ? okItens : (c0.abast ?? c0.lista ?? "")}" ${temItens ? "disabled" : ""}></div>
    </div>
    <div id="fFaltaBox" class="hidden"><label for="fFalta">Por que faltaram itens?</label>
      <select id="fFalta"><option value="">Escolha…</option>${MOTIVOS_FALTA.map(x => `<option ${a.motivoFalta === x ? "selected" : ""}>${x}</option>`).join("")}</select></div>
    <div class="calc" id="fCalc"></div>
    <label for="fObs">Observação</label>
    <textarea id="fObs" maxlength="300" placeholder="Ex.: acabou e foi para a Rua F">${esc(a.obs || "")}</textarea>
    <button class="btn ok" id="fOk">Finalizar</button>
    <button class="btn sec" data-fechar>Cancelar</button>`);
  const s = m.sheet;
  const ler = () => {
    let fim = quando(a.data, $("#fFim", s).value);
    if (fim && c0.ini != null && fim.getTime() <= c0.ini) fim = new Date(fim.getTime() + 864e5);
    const fimMs = fim ? fim.getTime() : null;
    return { ...a, fim: fimMs, pausas: fimMs ? fecharPausa(a, fimMs) : a.pausas, itensLista: lerInt($("#fLista", s).value), itensAbastecidos: lerInt($("#fAbast", s).value) };
  };
  const atualizar = () => {
    const c = calc(ler());
    $("#fCalc", s).innerHTML = calcBox(c);
    $("#fFaltaBox", s).classList.toggle("hidden", !(c.pend > 0) || temItens);
  };
  ["fFim", "fLista", "fAbast"].forEach(id => $("#" + id, s).oninput = atualizar);
  $("#fAgora", s).onclick = () => { $("#fFim", s).value = agoraHM(); atualizar(); };
  atualizar();
  $("#fOk", s).onclick = () => {
    const hm = $("#fFim", s).value;
    if (!hm) return toast("Informe o horário de término");
    const lista = lerInt($("#fLista", s).value), abast = lerInt($("#fAbast", s).value);
    if (Number.isNaN(lista) || Number.isNaN(abast)) return toast("As quantidades precisam ser números inteiros");
    let fim = quando(a.data, hm);
    if (c0.ini != null && fim.getTime() <= c0.ini) {
      if (!confirm("O término está antes do início. A atividade terminou no dia seguinte?")) return;
      fim = new Date(fim.getTime() + 864e5);
    }
    const faltou = !temItens && lista != null && abast != null && abast < lista;
    const falta = $("#fFalta", s).value;
    if (faltou && !falta) return toast("Escolha o motivo da falta");
    semEsperar(updateDoc(doc(db, "atividades", a.id), {
      fim: Timestamp.fromDate(fim), pausas: fecharPausa(a, fim.getTime()),
      itensLista: lista, itensAbastecidos: abast, motivoFalta: faltou ? falta : "",
      obs: $("#fObs", s).value.trim(), atualizadoEm: serverTimestamp(), atualizadoPor: S.perfil.nome,
    }));
    S.cache = null;
    m.fechar();
    toast("■ Atividade finalizada");
  };
}

/* ---------- Itens da lista, um por um ---------- */
// Cada item: { desc, status: "pendente" | "ok" | "falta", motivo, resolvido }
const ITEM_ST = { pendente: ["⬜", "Pendente"], ok: ["✅", "Abastecido"], falta: ["🔻", "Faltou"] };
// Mantém os totais da atividade iguais à lista (para o resumo, o Excel e as regras)
const totaisItens = itens => {
  const ok = itens.filter(i => i.status === "ok").length;
  return { itens, itensLista: itens.length || null, itensAbastecidos: itens.some(i => i.status !== "pendente") ? ok : null };
};
function lerListaColada(texto, antigos = []) {
  const vistos = new Map(antigos.map(i => [norm(i.desc), i]));
  return String(texto).split(/\r?\n/).map(l => l.replace(/\t+/g, " · ").trim()).filter(Boolean).slice(0, 500)
    .map(desc => vistos.get(norm(desc)) || { desc: desc.slice(0, 120), status: "pendente", motivo: "" });
}
function modalColarItens(atuais, aoSalvar) {
  const m = modal(`<h2>📝 Itens da lista</h2>
    <p class="muted">Um item por linha. Dá para copiar a coluna (ou as colunas) do Excel/sistema e colar aqui.</p>
    <textarea id="cItens" style="min-height:220px" placeholder="7891234 - CAIXA DE PAPELÃO&#10;7894567 - FITA ADESIVA&#10;…">${esc(atuais.map(i => i.desc).join("\n"))}</textarea>
    <p class="muted small" id="cConta"></p>
    <button class="btn" id="cOk">Usar esta lista</button>
    ${atuais.length ? `<button class="btn sec" id="cLimpar" style="color:var(--bad)">Tirar a lista item a item</button>` : ""}
    <button class="btn sec" data-fechar>Cancelar</button>`);
  const conta = () => { $("#cConta", m.sheet).textContent = `${lerListaColada($("#cItens", m.sheet).value).length} item(ns)`; };
  $("#cItens", m.sheet).oninput = conta; conta();
  $("#cOk", m.sheet).onclick = () => { aoSalvar(lerListaColada($("#cItens", m.sheet).value, atuais)); m.fechar(); };
  $("#cLimpar", m.sheet) && ($("#cLimpar", m.sheet).onclick = () => { if (confirm("Tirar a lista item a item? Ficam só os totais.")) { aoSalvar([]); m.fechar(); } });
}
function telaItens(id) {
  const a = S.atividades.find(x => x.id === id);
  if (!a) return voltarPainel();
  let itens = (a.itens || []).map(i => ({ ...i }));
  let filtro = S.filtroItens || "todos", tGravar = null;
  const marca = pode("executar");   // quem pode marcar abastecido / faltou
  // junta toques rápidos numa gravação só; "forcar" grava na hora
  const gravarItens = modo => {
    const fn = () => { tGravar = null; semEsperar(updateDoc(doc(db, "atividades", a.id), { ...totaisItens(itens), atualizadoEm: serverTimestamp(), atualizadoPor: S.perfil.nome })); S.cache = null; };
    if (modo === "pendente") { if (tGravar !== null) { clearTimeout(tGravar); fn(); } return; }
    clearTimeout(tGravar);
    if (modo === "forcar") fn(); else tGravar = setTimeout(fn, 1200);
  };
  topo(`☑️ ${a.localNome}`, [["Voltar", () => { gravarItens("pendente"); voltarPainel(); }]]);
  const desenhar = () => {
    const ok = itens.filter(i => i.status === "ok").length, falta = itens.filter(i => i.status === "falta").length, pend = itens.length - ok - falta;
    const vis = itens.map((i, k) => [i, k]).filter(([i]) => filtro === "todos" || i.status === filtro);
    app.innerHTML = `
      <div class="card">
        <div class="muted">${esc((a.pessoas || []).map(p => p.nome).join(", "))} · ${dataBR(a.data)}</div>
        <div class="barra-prog"><i style="width:${itens.length ? ok / itens.length * 100 : 0}%;background:var(--ok)"></i><i style="width:${itens.length ? falta / itens.length * 100 : 0}%;background:var(--bad)"></i></div>
        <div class="muted"><b style="color:var(--ok)">${ok} abastecido(s)</b> · <b style="color:var(--bad)">${falta} falta(s)</b> · ${pend} pendente(s) · total ${itens.length}</div>
      </div>
      <div class="filtros">${[["todos", "Todos"], ["pendente", "Pendentes"], ["ok", "Abastecidos"], ["falta", "Faltas"]].map(([k, t]) => `<button data-f="${k}" class="${filtro === k ? "on" : ""}">${t}</button>`).join("")}</div>
      ${pend && marca ? `<div class="linha-btns"><button class="btn sec" id="iTodos">✅ Marcar os ${pend} pendentes como abastecidos</button></div>` : ""}
      <div class="card"><ul class="list">${vis.map(([i, k]) => `<li class="item">
          <div class="info"><div class="title">${ITEM_ST[i.status][0]} ${esc(i.desc)}</div>${i.status === "falta" ? `<div class="muted" style="color:var(--bad)">${esc(i.motivo || "sem motivo")}${i.resolvido ? " · ✔ resolvido" : ""}</div>` : ""}</div>
          ${marca ? `<div class="acts"><button class="mini ${i.status === "ok" ? "ok" : ""}" data-ok="${k}" title="Abastecido">✔</button><button class="mini ${i.status === "falta" ? "del" : ""}" data-falta="${k}" title="Faltou">✖</button></div>` : ""}</li>`).join("") || `<div class="empty">Nenhum item aqui.</div>`}</ul></div>
      <div class="linha-btns">${pode("lancar") ? `<button class="btn sec" id="iEditar">📝 Editar a lista</button>` : ""}
        ${marca && calc(a).status !== "concluida" && a.inicio ? `<button class="btn ok" id="iFinalizar">■ Finalizar a atividade</button>` : ""}</div>`;
    $$("[data-f]").forEach(b => b.onclick = () => { filtro = S.filtroItens = b.dataset.f; desenhar(); });
    $$("[data-ok]").forEach(b => b.onclick = () => { const i = itens[+b.dataset.ok]; i.status = i.status === "ok" ? "pendente" : "ok"; i.motivo = ""; gravarItens(); desenhar(); });
    $$("[data-falta]").forEach(b => b.onclick = () => {
      const i = itens[+b.dataset.falta];
      if (i.status === "falta") { i.status = "pendente"; i.motivo = ""; gravarItens(); return desenhar(); }
      const mm = modal(`<h2>🔻 Faltou</h2><p>${esc(i.desc)}</p><div class="chips">${MOTIVOS_FALTA.map(x => `<button type="button" class="chip" data-mf="${esc(x)}">${esc(x)}</button>`).join("")}</div><button class="btn sec" data-fechar>Cancelar</button>`);
      $$("[data-mf]", mm.sheet).forEach(c => c.onclick = () => { i.status = "falta"; i.motivo = c.dataset.mf; i.resolvido = false; mm.fechar(); gravarItens(); desenhar(); });
    });
    $("#iTodos") && ($("#iTodos").onclick = () => { itens.forEach(i => { if (i.status === "pendente") i.status = "ok"; }); gravarItens(); desenhar(); });
    $("#iEditar") && ($("#iEditar").onclick = () => modalColarItens(itens, novos => { itens = novos; gravarItens("forcar"); desenhar(); }));
    $("#iFinalizar") && ($("#iFinalizar").onclick = () => {
      gravarItens("forcar");
      const base = { ...(S.atividades.find(x => x.id === a.id) || a), ...totaisItens(itens) };
      voltarPainel();
      setTimeout(() => modalFinalizar(base), 300);
    });
  };
  desenhar();
}

/* ---------- Nova / editar atividade ---------- */
function telaAtividade(a, preset = {}) {
  const novo = !a;
  topo(novo ? "＋ Nova atividade" : "✏️ Editar atividade", [["Voltar", voltarPainel]]);
  const d = a || { data: preset.data || S.dia, localId: preset.localId || "", pessoas: [], itensLista: null, itensAbastecidos: null, obs: "" };
  const sel = new Map((d.pessoas || []).map(p => [p.id || "n:" + norm(p.nome), p]));
  let itensF = (d.itens || []).map(i => ({ ...i })), pausasF = [...(d.pausas || [])];
  // locais agrupados por tipo (inclui o local da atividade mesmo se foi desativado)
  const locais = S.locais.filter(l => l.ativo !== false || l.id === d.localId);
  const tipos = [...new Set(locais.map(l => l.tipo || "Outros"))];
  const orfao = d.localId && !locais.some(l => l.id === d.localId);
  app.innerHTML = `
    <div class="card">
      <div class="row">
        <div><label for="aData">Data</label><input type="date" id="aData" value="${esc(d.data)}"></div>
        <div><label for="aLocal">Local</label>
          <select id="aLocal"><option value="">Escolha…</option>
            ${tipos.map(t => `<optgroup label="${esc(t)}">${locais.filter(l => (l.tipo || "Outros") === t).map(l => `<option value="${l.id}" ${l.id === d.localId ? "selected" : ""}>${esc(l.nome)}</option>`).join("")}</optgroup>`).join("")}
            ${orfao ? `<option value="${esc(d.localId)}" selected>${esc(d.localNome)}</option>` : ""}
          </select></div>
      </div>
      <label for="aBusca">Colaboradores <span class="muted" id="aQtd"></span></label>
      <input id="aBusca" placeholder="🔎 Procurar ou digitar um nome novo" autocomplete="off">
      <div class="chips" id="aChips"></div>
      <div class="row horas">
        <div><label for="aIni">Início</label><div class="campo-agora"><input type="time" id="aIni" value="${d.inicio ? hhmm(d.inicio) : ""}"><button class="mini" type="button" data-agora="aIni">Agora</button></div></div>
        <div><label for="aFim">Término</label><div class="campo-agora"><input type="time" id="aFim" value="${d.fim ? hhmm(d.fim) : ""}"><button class="mini" type="button" data-agora="aFim">Agora</button></div></div>
      </div>
      <div class="row">
        <div><label for="aLista">Qtd. itens da lista</label><input id="aLista" type="number" inputmode="numeric" min="0" step="1" value="${d.itensLista ?? ""}"></div>
        <div><label for="aAbast">Itens abastecidos</label><input id="aAbast" type="number" inputmode="numeric" min="0" step="1" value="${d.itensAbastecidos ?? ""}"></div>
      </div>
      <div class="caixa-itens"><span id="aItensTxt"></span><button class="mini" type="button" id="aItensBtn">📝 Lista item a item</button></div>
      <div id="aPausas"></div>
      <div class="calc" id="aCalc"></div>
      <label for="aObs">Observação</label>
      <textarea id="aObs" maxlength="300" placeholder="Ex.: ficou organizando a rua depois do abastecimento">${esc(d.obs || "")}</textarea>
      ${pode("lancar") ? `<button class="btn" id="aSalvar">${novo ? "Salvar atividade" : "Salvar alterações"}</button>` : `<div class="alert info">Você pode ver, mas não alterar esta atividade.</div>`}
      ${!novo && pode("excluir") ? `<button class="btn sec" id="aExcluir" style="color:var(--bad)">🗑️ Excluir atividade</button>` : ""}
      ${!novo ? `<p class="muted small">Criada por ${esc(a.criadoPor || "—")}${a.atualizadoPor ? ` · última alteração por ${esc(a.atualizadoPor)}` : ""}</p>` : ""}
    </div>`;

  const chips = () => {
    const b = norm($("#aBusca").value);
    const ativos = S.colaboradores.filter(c => c.ativo !== false);
    const visiveis = ativos.filter(c => !b || norm(c.nome).includes(b) || sel.has(c.id));
    const extras = [...sel.entries()].filter(([k]) => !ativos.some(c => c.id === k));   // selecionados que não estão na lista ativa
    const exato = ativos.some(c => norm(c.nome) === b);
    $("#aChips").innerHTML = [
      ...extras.map(([k, p]) => `<button type="button" class="chip on" data-k="${esc(k)}">✓ ${esc(p.nome)}</button>`),
      ...visiveis.map(c => `<button type="button" class="chip ${sel.has(c.id) ? "on" : ""}" data-k="${c.id}">${sel.has(c.id) ? "✓ " : ""}${esc(c.nome)}</button>`),
      b && !exato && pode("colab") ? `<button type="button" class="chip" data-novo="1">＋ Cadastrar “${esc($("#aBusca").value.trim())}”</button>` : "",
    ].join("") || `<span class="muted">${pode("colab") ? "Nenhum colaborador cadastrado. Digite um nome acima para cadastrar." : "Nenhum colaborador encontrado."}</span>`;
    $("#aQtd").textContent = sel.size ? `· ${sel.size} selecionado(s)` : "";
    $$("#aChips [data-k]").forEach(x => x.onclick = () => {
      const k = x.dataset.k;
      if (sel.has(k)) sel.delete(k);
      else { const c = S.colaboradores.find(c => c.id === k); if (c) sel.set(k, { id: c.id, nome: c.nome }); }
      chips();
    });
    $("#aChips [data-novo]") && ($("#aChips [data-novo]").onclick = () => {
      const nome = $("#aBusca").value.trim().replace(/\s+/g, " ").toUpperCase();
      if (nome.length < 2) return toast("Digite o nome");
      const ref = doc(collection(db, "colaboradores"));
      semEsperar(setDoc(ref, { nome, ativo: true, criadoEm: serverTimestamp() }));
      S.colaboradores.push({ id: ref.id, nome, ativo: true }); S.colaboradores.sort(porNome);
      sel.set(ref.id, { id: ref.id, nome });
      $("#aBusca").value = "";
      toast(`Colaborador ${nome} cadastrado`);
      chips();
    });
  };
  $("#aBusca").oninput = chips;
  $("#aBusca").onkeydown = e => { if (e.key === "Enter") { e.preventDefault(); const p = $("#aChips .chip:not(.on)"); p?.click(); } };
  chips();

  const ler = () => {
    const dataIso = $("#aData").value || d.data;
    let ini = quando(dataIso, $("#aIni").value), fim = quando(dataIso, $("#aFim").value);
    if (ini && fim && fim <= ini) fim = new Date(fim.getTime() + 864e5);   // terminou depois da meia-noite
    return { dataIso, ini, fim, lista: lerInt($("#aLista").value), abast: lerInt($("#aAbast").value) };
  };
  const atualizar = () => {
    const v = ler();
    $("#aCalc").innerHTML = calcBox(calc({ data: v.dataIso, localId: $("#aLocal").value, inicio: v.ini?.getTime(), fim: v.fim?.getTime(), itensLista: v.lista, itensAbastecidos: v.abast, pausas: pausasF }));
  };
  // lista item a item: quando existe, as quantidades vêm dela
  const desenharItens = () => {
    const tem = itensF.length > 0, t = totaisItens(itensF);
    $("#aItensTxt").innerHTML = tem ? `<b>${itensF.length}</b> item(ns) na lista · ${itensF.filter(i => i.status === "ok").length} abastecido(s) · ${itensF.filter(i => i.status === "falta").length} falta(s)` : `<span class="muted">Opcional: colar os itens da lista para marcar um por um</span>`;
    $("#aItensBtn").textContent = tem ? "📝 Editar a lista" : "📝 Lista item a item";
    $("#aLista").disabled = $("#aAbast").disabled = tem;
    if (tem) { $("#aLista").value = t.itensLista ?? ""; $("#aAbast").value = t.itensAbastecidos ?? ""; }
    atualizar();
  };
  $("#aItensBtn").onclick = () => modalColarItens(itensF, novos => { itensF = novos; desenharItens(); });
  const desenharPausas = () => {
    $("#aPausas").innerHTML = pausasF.length ? `<div class="muted small" style="margin-bottom:8px"><b>Pausas:</b> ${pausasF.map((p, k) => `<span class="pausa-tag">${hhmm(p.ini)}–${p.fim ? hhmm(p.fim) : "agora"} ${esc(p.motivo)} <button type="button" data-tira="${k}" title="Tirar esta pausa">✕</button></span>`).join(" ")}</div>` : "";
    $$("[data-tira]").forEach(b => b.onclick = () => { if (confirm("Tirar esta pausa? O tempo dela volta a contar como trabalhado.")) { pausasF.splice(+b.dataset.tira, 1); desenharPausas(); atualizar(); } });
  };
  ["aIni", "aFim", "aLista", "aAbast", "aData", "aLocal"].forEach(id => $("#" + id).addEventListener("input", atualizar));
  $$("[data-agora]").forEach(b => b.onclick = () => { $("#" + b.dataset.agora).value = agoraHM(); atualizar(); });
  desenharPausas(); desenharItens();

  if ($("#aSalvar")) $("#aSalvar").onclick = () => {
    const v = ler();
    const localId = $("#aLocal").value;
    const loc = S.locais.find(l => l.id === localId) || (localId && localId === d.localId ? { id: d.localId, nome: d.localNome, tipo: d.tipo } : null);
    if (!loc) return toast("Escolha o local");
    if (!sel.size) return toast("Escolha pelo menos um colaborador");
    if (v.fim && !v.ini) return toast("Preencha o horário de início");
    if (Number.isNaN(v.lista) || Number.isNaN(v.abast)) return toast("As quantidades precisam ser números inteiros");
    if ($("#aIni").value && $("#aFim").value && $("#aFim").value <= $("#aIni").value
      && !confirm("O término está antes do início. A atividade terminou no dia seguinte?")) return;
    const pessoas = [...sel.values()].map(p => ({ id: p.id || null, nome: p.nome }));
    const reg = {
      data: v.dataIso, localId: loc.id, localNome: loc.nome, tipo: loc.tipo || "",
      pessoas, pessoaIds: pessoas.map(p => p.id).filter(Boolean),
      inicio: v.ini ? Timestamp.fromDate(v.ini) : null, fim: v.fim ? Timestamp.fromDate(v.fim) : null,
      itensLista: v.lista, itensAbastecidos: v.abast, obs: $("#aObs").value.trim(),
      pausas: pausasF, ...(itensF.length ? totaisItens(itensF) : (d.itens || []).length ? { itens: [] } : {}),
      atualizadoEm: serverTimestamp(), atualizadoPor: S.perfil.nome,
    };
    if (novo) semEsperar(setDoc(doc(collection(db, "atividades")), { ...reg, criadoEm: serverTimestamp(), criadoPor: S.perfil.nome }));
    else semEsperar(updateDoc(doc(db, "atividades", a.id), reg));
    S.cache = null;
    toast(novo ? "Atividade salva" : "Alterações salvas");
    if (S.aba === "dia" && v.dataIso !== S.dia) { S.dia = v.dataIso; S.atividades = []; ouvirDia(); }
    voltarPainel();
  };
  $("#aExcluir") && ($("#aExcluir").onclick = () => {
    if (!confirm(`Excluir a atividade de ${a.localNome} (${dataBR(a.data)})? Não dá para desfazer.`)) return;
    semEsperar(deleteDoc(doc(db, "atividades", a.id)));
    S.cache = null;
    toast("Atividade excluída");
    voltarPainel();
  });
}

// Copia locais e colaboradores de outro dia (sem horários), para só tocar em ▶ Iniciar
function copiarEscala() {
  const m = modal(`<h2>📋 Copiar escala</h2>
    <p class="muted">Cria para <b>${dataBR(S.dia)}</b> as mesmas atividades (local e colaboradores) de outro dia, ainda sem horários. Depois é só tocar em ▶ Iniciar em cada uma.</p>
    <label for="cOrig">Copiar do dia</label>
    <input type="date" id="cOrig" value="${somaDias(S.dia, -1)}">
    <label style="display:flex;gap:8px;align-items:center;color:var(--text)"><input type="checkbox" id="cLista" style="width:auto;margin:0"> Copiar também a quantidade de itens da lista</label>
    <button class="btn" id="cOk" style="margin-top:12px">Copiar</button>
    <button class="btn sec" data-fechar>Cancelar</button>`);
  $("#cOk", m.sheet).onclick = async () => {
    const orig = $("#cOrig", m.sheet).value, comLista = $("#cLista", m.sheet).checked;
    if (!orig || orig === S.dia) return toast("Escolha outro dia");
    ocupado("Copiando…");
    const r = await tentar(async () => {
      const snap = await getDocs(query(collection(db, "atividades"), where("data", "==", orig)));
      const origem = lista(snap);
      if (!origem.length) throw new Error(`Não há atividades em ${dataBR(orig)}.`);
      const chave = a => a.localId + "|" + (a.pessoas || []).map(p => p.id || norm(p.nome)).sort().join(",");
      const ja = new Set(S.atividades.map(chave));
      const novas = origem.filter(a => !ja.has(chave(a)));
      const b = writeBatch(db);
      novas.forEach(a => b.set(doc(collection(db, "atividades")), {
        data: S.dia, localId: a.localId, localNome: a.localNome, tipo: a.tipo || "", pessoas: a.pessoas || [], pessoaIds: a.pessoaIds || [],
        inicio: null, fim: null, itensLista: comLista ? (a.itensLista ?? null) : null, itensAbastecidos: null, obs: "",
        criadoEm: serverTimestamp(), criadoPor: S.perfil.nome, atualizadoEm: serverTimestamp(), atualizadoPor: S.perfil.nome,
      }));
      if (novas.length) await b.commit();
      return { novas: novas.length, iguais: origem.length - novas.length };
    });
    if (!r) return;
    m.fechar();
    S.cache = null;
    toast(r.novas ? `${r.novas} atividade(s) criada(s)${r.iguais ? ` · ${r.iguais} já existiam` : ""}` : "Todas essas atividades já existem neste dia", 4500);
  };
}

/* ---------- Período (Resumo e Histórico) ---------- */
const PERIODOS = [["hoje", "Hoje"], ["ontem", "Ontem"], ["7d", "7 dias"], ["mes", "Este mês"], ["mesant", "Mês passado"], ["pers", "Escolher…"]];
function intervalo(p) {
  const h = hojeISO();
  if (p.tipo === "hoje") return [h, h];
  if (p.tipo === "ontem") { const o = somaDias(h, -1); return [o, o]; }
  if (p.tipo === "7d") return [somaDias(h, -6), h];
  if (p.tipo === "mes") return [h.slice(0, 8) + "01", h];
  if (p.tipo === "mesant") { const d = deISO(h); return [iso(new Date(d.getFullYear(), d.getMonth() - 1, 1)), iso(new Date(d.getFullYear(), d.getMonth(), 0))]; }
  let de = p.de || h, ate = p.ate || h;
  if (de > ate) [de, ate] = [ate, de];
  return [de, ate];
}
const textoPeriodo = (de, ate) => de === ate ? `${diaSemana(de)}, ${dataBR(de)}` : `${dataBR(de)} a ${dataBR(ate)}`;
function periodoHtml() {
  const [de, ate] = intervalo(S.periodo);
  return `<div class="filtros">${PERIODOS.map(([k, t]) => `<button data-per="${k}" class="${S.periodo.tipo === k ? "on" : ""}">${t}</button>`).join("")}</div>
    ${S.periodo.tipo === "pers" ? `<div class="row no-print" style="max-width:420px"><div><label>De</label><input type="date" id="pDe" value="${de}"></div><div><label>Até</label><input type="date" id="pAte" value="${ate}"></div></div>` : ""}`;
}
function ligarPeriodo(el, redesenhar) {
  $$("[data-per]", el).forEach(b => b.onclick = () => {
    const [de, ate] = intervalo(S.periodo);
    S.periodo = b.dataset.per === "pers" ? { tipo: "pers", de, ate } : { tipo: b.dataset.per };
    redesenhar();
  });
  $("#pDe", el) && ($("#pDe", el).onchange = e => { S.periodo.de = e.target.value; redesenhar(); });
  $("#pAte", el) && ($("#pAte", el).onchange = e => { S.periodo.ate = e.target.value; redesenhar(); });
}
// Busca as atividades do período (guarda por 2 min para não gastar leituras à toa)
async function buscarPeriodo(de, ate) {
  const chave = de + "|" + ate;
  if (S.cache && S.cache.chave === chave && Date.now() - S.cache.t < 120000) return S.cache.dados;
  if ((deISO(ate) - deISO(de)) / 864e5 > 400) throw new Error("Escolha um período de até 1 ano.");
  const snap = await getDocs(query(collection(db, "atividades"), where("data", ">=", de), where("data", "<=", ate)));
  const dadosP = lista(snap).sort((a, b) => a.data.localeCompare(b.data) || (ordemDe(a) - ordemDe(b)) || ((ms(a.inicio) ?? 9e15) - (ms(b.inicio) ?? 9e15)));
  S.cache = { chave, t: Date.now(), dados: dadosP };
  return dadosP;
}

/* ---------- Aba Resumo ---------- */
async function abaResumo(el) {
  el.innerHTML = `${periodoHtml()}<div id="rCorpo"><div class="empty">Carregando…</div></div>`;
  ligarPeriodo(el, () => abaResumo(el));
  const [de, ate] = intervalo(S.periodo);
  const dadosP = await tentar(() => buscarPeriodo(de, ate));
  const corpo = $("#rCorpo", el);
  if (!dadosP || !corpo || !corpo.isConnected) return;
  if (!dadosP.length) { corpo.innerHTML = `<div class="card empty">Nenhuma atividade em ${textoPeriodo(de, ate)}.</div>`; return; }
  const T = totais(dadosP), umDia = de === ate;
  const locaisG = porLocal(dadosP), pessoas = porColaborador(dadosP), dias = porDia(dadosP);
  const maxIph = Math.max(1, ...locaisG.map(g => Math.max(g.t.iph || 0, metaDe(g.k) || 0)));
  // faltas: itens marcados como "faltou" (lista item a item) e atividades finalizadas com falta
  const faltasItens = dadosP.flatMap(a => (a.itens || []).map((i, k) => ({ a, i, k })).filter(x => x.i.status === "falta"));
  const faltasAtiv = dadosP.filter(a => !(a.itens || []).length && calc(a).pend > 0);
  const porMotivoFalta = {};
  faltasItens.forEach(x => { const m = x.i.motivo || "Sem motivo"; porMotivoFalta[m] = (porMotivoFalta[m] || 0) + 1; });
  faltasAtiv.forEach(a => { const m = a.motivoFalta || "Sem motivo"; porMotivoFalta[m] = (porMotivoFalta[m] || 0) + calc(a).pend; });
  const barrasDe = (obj, fmt) => { const e = Object.entries(obj).sort((a, b) => b[1] - a[1]); const mx = Math.max(1, ...e.map(x => x[1])); return e.length ? `<div class="barras">${e.map(([k, v]) => `<div class="b"><span>${esc(k)}</span><div class="trilho"><i style="width:${Math.round(v / mx * 100)}%"></i></div><b>${fmt(v)}</b></div>`).join("")}</div>` : ""; };
  const linhaLocal = g => { const meta = metaDe(g.k), pm = g.t.pctMeta; return `<tr>
      <td>${esc(g.nome)}</td><td class="esq">${esc(g.tipo || "")}</td>
      ${umDia ? `<td>${g.local?.previstas ?? "—"}</td>` : ""}
      <td>${g.t.nPessoas}</td><td>${g.t.registros}</td><td>${fmtN(g.t.lista)}</td><td>${fmtN(g.t.abast)}</td>
      <td class="${g.t.pend ? "pend" : ""}">${fmtN(g.t.pend)}</td><td>${fmtDur(g.t.min)}</td><td>${g.t.parado >= 1 ? fmtDur(g.t.parado) : "—"}</td>
      <td>${fmtN(g.t.iph, 1)}</td><td>${fmtN(meta, 1)}</td><td style="color:${corMeta(pm)};font-weight:700">${pct(pm)}</td><td>${pct(g.t.cobertura)}</td></tr>`; };
  const prevTotal = locaisG.reduce((s, g) => s + (num(g.local?.previstas) || 0), 0);
  corpo.innerHTML = `
    <div class="so-print"><h2 style="margin:0">RESUMO DE PRODUTIVIDADE · ${esc(S.config?.empresa || "")}</h2><p class="muted">${textoPeriodo(de, ate)} · gerado em ${dataBR(hojeISO())} ${agoraHM()}</p></div>
    <p class="muted no-print" style="margin-top:0">${textoPeriodo(de, ate)}</p>
    ${kpisHtml(T, true)}
    <div class="linha-btns no-print">
      <button class="btn" id="rExcel">📥 Baixar Excel</button>
      ${umDia ? `<button class="btn sec" id="rRel">📤 Relatório do dia</button>` : ""}
      <button class="btn sec" id="rImprimir">🖨️ Imprimir / PDF</button>
    </div>
    <div class="card"><h2>📍 Por local</h2><div class="tabela-wrap"><table class="tab">
      <thead><tr><th>Local</th><th class="esq">Tipo</th>${umDia ? "<th>Previstas</th>" : ""}<th>Pessoas</th><th>Registros</th><th>Itens lista</th><th>Abastecidos</th><th>Pendentes</th><th>Tempo</th><th>Parado</th><th>Itens/h</th><th>Meta/h</th><th>% meta</th><th>% lista</th></tr></thead>
      <tbody>${locaisG.map(linhaLocal).join("")}
        <tr class="total"><td>Total</td><td></td>${umDia ? `<td>${prevTotal}</td>` : ""}<td>${T.nPessoas}</td><td>${T.registros}</td><td>${fmtN(T.lista)}</td><td>${fmtN(T.abast)}</td><td class="${T.pend ? "pend" : ""}">${fmtN(T.pend)}</td><td>${fmtDur(T.min)}</td><td>${T.parado >= 1 ? fmtDur(T.parado) : "—"}</td><td>${fmtN(T.iph, 1)}</td><td></td><td style="color:${corMeta(T.pctMeta)}">${pct(T.pctMeta)}</td><td>${pct(T.cobertura)}</td></tr>
      </tbody></table></div>
      <p class="muted small">${umDia ? "Previstas = pessoas previstas no cadastro do local · Pessoas = quem trabalhou de fato. " : ""}Meta/h = média das últimas 4 semanas (ou a meta fixa do cadastro do local). Tempo = trabalhado, sem as pausas.</p></div>
    <div class="duas">
      <div class="card"><h2>⚡ Itens por hora × meta</h2><div class="barras">${locaisG.map(g => { const meta = metaDe(g.k); return `<div class="b"><span>${esc(g.nome)}</span><div class="trilho" style="position:relative"><i style="width:${Math.round((g.t.iph || 0) / maxIph * 100)}%;background:${corMeta(g.t.pctMeta)}"></i>${meta ? `<em class="marca-meta" style="left:${Math.round(meta / maxIph * 100)}%" title="meta ${fmtN(meta, 1)}"></em>` : ""}</div><b>${fmtN(g.t.iph, 1)}</b></div>`; }).join("")}</div>
        <p class="muted small">A barra mostra os itens/hora; o tracinho preto é a meta.</p></div>
      <div class="card"><h2>⏸ Tempo parado por motivo · ${fmtDur(T.parado)}</h2>${barrasDe(T.porMotivo, v => fmtDur(v)) || `<div class="empty">Nenhuma pausa registrada no período.</div>`}</div>
    </div>
    <div class="card"><h2>🔻 Faltas e pendências</h2>
      ${Object.keys(porMotivoFalta).length ? `<h3 style="margin-top:0">Por motivo (itens)</h3>${barrasDe(porMotivoFalta, v => fmtN(v))}` : ""}
      ${faltasItens.length || faltasAtiv.length ? `<ul class="list">
        ${faltasItens.map(x => `<li class="item"><div class="info"><div class="title">${x.i.resolvido ? "✔ " : ""}${esc(x.i.desc)}</div><div class="muted">${dataBR(x.a.data)} · ${esc(x.a.localNome)} · <b style="color:var(--bad)">${esc(x.i.motivo || "sem motivo")}</b>${x.i.resolvido ? " · resolvido" : ""}</div></div>
          ${pode("lancar") ? `<div class="acts"><button class="mini" data-resolve="${x.a.id}|${x.k}">${x.i.resolvido ? "Reabrir" : "✔ Resolvido"}</button></div>` : ""}</li>`).join("")}
        ${faltasAtiv.map(a => { const c = calc(a); return `<li class="item"><div class="info"><div class="title">${esc(a.localNome)} · faltaram ${c.pend} de ${c.lista}</div><div class="muted">${dataBR(a.data)} · ${esc((a.pessoas || []).map(p => p.nome).join(", "))} · <b style="color:var(--bad)">${esc(a.motivoFalta || "sem motivo")}</b>${a.obs ? ` · ${esc(a.obs)}` : ""}</div></div></li>`; }).join("")}
      </ul>` : `<div class="empty">Nenhum item pendente no período 🎉</div>`}</div>
    <div class="card"><h2>👷 Por colaborador</h2><div class="tabela-wrap"><table class="tab">
      <thead><tr><th>Colaborador</th><th>Atividades</th>${umDia ? "" : "<th>Dias</th>"}<th>Itens abastecidos*</th><th>Pendentes*</th><th>Tempo</th><th>Itens/h</th><th class="esq">Locais</th></tr></thead>
      <tbody>${pessoas.map((p, i) => `<tr><td>${i < 3 ? ["🥇", "🥈", "🥉"][i] + " " : ""}${esc(p.nome)}</td><td>${p.ativ}</td>${umDia ? "" : `<td>${p.dias.size}</td>`}<td>${fmtN(p.abast, Number.isInteger(p.abast) ? 0 : 1)}</td><td>${fmtN(p.pend, Number.isInteger(p.pend) ? 0 : 1)}</td><td>${fmtDur(p.min)}</td><td>${fmtN(p.iph, 1)}</td><td class="esq">${esc([...p.locais].join(", "))}</td></tr>`).join("")}</tbody>
    </table></div><p class="muted small">* Quando a atividade tem mais de uma pessoa, os itens são divididos igualmente entre elas.</p></div>
    ${umDia ? "" : `<div class="card"><h2>📅 Por dia</h2><div class="tabela-wrap"><table class="tab">
      <thead><tr><th>Data</th><th>Registros</th><th>Pessoas</th><th>Itens lista</th><th>Abastecidos</th><th>Pendentes</th><th>Tempo</th><th>Parado</th><th>Itens/h</th><th>% meta</th><th>% lista</th></tr></thead>
      <tbody>${dias.map(x => `<tr class="click" data-ir-dia="${x.data}"><td>${dataBR(x.data)} <span class="muted small">${diaSemana(x.data).slice(0, 3)}</span></td><td>${x.t.registros}</td><td>${x.t.nPessoas}</td><td>${fmtN(x.t.lista)}</td><td>${fmtN(x.t.abast)}</td><td class="${x.t.pend ? "pend" : ""}">${fmtN(x.t.pend)}</td><td>${fmtDur(x.t.min)}</td><td>${x.t.parado >= 1 ? fmtDur(x.t.parado) : "—"}</td><td>${fmtN(x.t.iph, 1)}</td><td style="color:${corMeta(x.t.pctMeta)}">${pct(x.t.pctMeta)}</td><td>${pct(x.t.cobertura)}</td></tr>`).join("")}</tbody>
    </table></div></div>`}`;
  $("#rExcel", corpo).onclick = async () => { ocupado("Gerando o Excel…"); await tentar(() => exportarExcel(dadosP, de, ate)); };
  $("#rImprimir", corpo).onclick = () => window.print();
  $("#rRel", corpo) && ($("#rRel", corpo).onclick = () => modalRelatorio(de, dadosP));
  $$("[data-ir-dia]", corpo).forEach(tr => tr.onclick = () => { S.aba = "dia"; S.dia = tr.dataset.irDia; S.atividades = []; ouvirDia(); go(telaPainel); });
  $$("[data-resolve]", corpo).forEach(b => b.onclick = () => {
    const [id, k] = b.dataset.resolve.split("|");
    const a = dadosP.find(x => x.id === id);
    const itens = (a.itens || []).map((i, j) => j === +k ? { ...i, resolvido: !i.resolvido } : i);
    a.itens = itens;
    semEsperar(updateDoc(doc(db, "atividades", id), { itens, atualizadoEm: serverTimestamp(), atualizadoPor: S.perfil.nome }));
    abaResumo(el);
  });
}

/* ---------- Aba Histórico ---------- */
async function abaHistorico(el) {
  el.innerHTML = `${periodoHtml()}<div id="hCorpo"><div class="empty">Carregando…</div></div>`;
  ligarPeriodo(el, () => abaHistorico(el));
  const [de, ate] = intervalo(S.periodo);
  const dadosP = await tentar(() => buscarPeriodo(de, ate));
  const corpo = $("#hCorpo", el);
  if (!dadosP || !corpo || !corpo.isConnected) return;
  const F = S.filtroH;
  const nomes = [...new Set(dadosP.flatMap(a => (a.pessoas || []).map(p => p.nome)))].sort((a, b) => a.localeCompare(b, "pt-BR"));
  const locs = [...new Set(dadosP.map(a => a.localNome))].sort((a, b) => a.localeCompare(b, "pt-BR"));
  corpo.innerHTML = `
    <div class="card no-print">
      <input id="hTexto" placeholder="🔎 Buscar (nome, local, observação)" value="${esc(F.texto)}">
      <div class="row" style="flex-wrap:wrap">
        <select id="hPessoa"><option value="">Todos os colaboradores</option>${nomes.map(n => `<option ${F.pessoa === n ? "selected" : ""}>${esc(n)}</option>`).join("")}</select>
        <select id="hLocal"><option value="">Todos os locais</option>${locs.map(n => `<option ${F.local === n ? "selected" : ""}>${esc(n)}</option>`).join("")}</select>
        <select id="hStatus">${[["", "Todas"], ["concluida", "Concluídas"], ["andamento", "Sem término"], ["planejada", "Planejadas"], ["pend", "Com pendência"]].map(([v, t]) => `<option value="${v}" ${F.status === v ? "selected" : ""}>${t}</option>`).join("")}</select>
      </div>
      <button class="btn sec" id="hExcel">📥 Baixar Excel do que está filtrado</button>
    </div>
    <div class="card"><div id="hLista"></div></div>`;
  let filtrados = [];
  const desenhar = () => {
    const tx = norm(F.texto);
    filtrados = dadosP.filter(a => {
      const c = calc(a);
      if (F.pessoa && !(a.pessoas || []).some(p => p.nome === F.pessoa)) return false;
      if (F.local && a.localNome !== F.local) return false;
      if (F.status === "pend" ? !(c.pend > 0) : F.status && c.status !== F.status) return false;
      if (tx && !norm(`${a.localNome} ${a.tipo} ${(a.pessoas || []).map(p => p.nome).join(" ")} ${a.obs}`).includes(tx)) return false;
      return true;
    }).reverse();   // mais recentes primeiro
    $("#hLista", corpo).innerHTML = filtrados.length ? `<p class="muted" style="margin-top:0">${filtrados.length} atividade(s) · ${textoPeriodo(de, ate)}</p>
      <div class="tabela-wrap"><table class="tab">
      <thead><tr><th>Data</th><th class="esq">Local</th><th class="esq">Colaborador(es)</th><th>Início</th><th>Término</th><th>Tempo</th><th>Lista</th><th>Abast.</th><th>Pend.</th><th>Itens/h</th><th class="esq">Observação</th></tr></thead>
      <tbody>${filtrados.map(a => { const c = calc(a); return `<tr class="click" data-h="${a.id}">
        <td>${dataBR(a.data)}</td><td class="esq">${esc(a.localNome)}</td><td class="esq">${esc((a.pessoas || []).map(p => p.nome).join(", "))}</td>
        <td>${hhmm(c.ini)}</td><td>${c.fim != null ? hhmm(c.fim) : "—"}</td><td>${fmtDur(c.min)}</td><td>${c.lista ?? "—"}</td><td>${c.abast ?? "—"}</td>
        <td class="${c.pend ? "pend" : ""}">${c.pend ?? "—"}</td><td>${fmtN(c.iph, 1)}</td><td class="esq" style="white-space:normal;min-width:180px">${esc(a.obs || "")}</td></tr>`; }).join("")}</tbody></table></div>`
      : `<div class="empty">Nenhuma atividade encontrada.</div>`;
    $$("[data-h]", corpo).forEach(tr => tr.onclick = () => go(telaAtividade, dadosP.find(a => a.id === tr.dataset.h)));
  };
  $("#hTexto", corpo).oninput = e => { F.texto = e.target.value; desenhar(); };
  $("#hPessoa", corpo).onchange = e => { F.pessoa = e.target.value; desenhar(); };
  $("#hLocal", corpo).onchange = e => { F.local = e.target.value; desenhar(); };
  $("#hStatus", corpo).onchange = e => { F.status = e.target.value; desenhar(); };
  $("#hExcel", corpo).onclick = async () => {
    if (!filtrados.length) return toast("Nada para exportar");
    ocupado("Gerando o Excel…");
    await tentar(() => exportarExcel([...filtrados].reverse(), de, ate));
  };
  desenhar();
}

/* ---------- Aba Cadastros ---------- */
function abaCadastros(el) {
  const subs = [
    ...(pode("colab") ? [["colaboradores", "👷 Colaboradores"]] : []),
    ...(pode("locais") ? [["locais", "📍 Locais"]] : []),
    ...(supervisor() ? [["usuarios", "🔑 Usuários do app"]] : []),
    ...(pode("importar") ? [["importar", "📤 Importar planilha"]] : []),
    ...(supervisor() ? [["empresa", "🏢 Empresa"]] : []),
  ];
  if (!subs.length) { el.innerHTML = `<div class="card empty">Você não tem acesso aos cadastros.</div>`; return; }
  if (!subs.some(s => s[0] === S.sub)) S.sub = subs[0][0];
  el.innerHTML = `<div class="sub-abas">${subs.map(([k, t]) => `<button data-sub="${k}" class="${S.sub === k ? "on" : ""}">${t}</button>`).join("")}</div><div id="cCorpo"></div>`;
  $$("[data-sub]", el).forEach(b => b.onclick = () => { S.sub = b.dataset.sub; telaPainel(); });
  S.live = S.sub !== "importar" && S.sub !== "empresa";
  if (S.sub === "empresa") S.live = false;   // não apaga o que está sendo digitado
  ({ colaboradores: subColaboradores, locais: subLocais, usuarios: subUsuarios, importar: subImportar, empresa: subEmpresa }[S.sub])($("#cCorpo", el));
}

function subColaboradores(el) {
  const ativos = S.colaboradores.filter(c => c.ativo !== false), inativos = S.colaboradores.filter(c => c.ativo === false);
  const item = c => `<li class="item"><div class="info"><div class="title">${esc(c.nome)} ${c.ativo === false ? `<span class="pill off">Inativo</span>` : ""}</div></div>
    <div class="acts"><button class="mini" data-ren="${c.id}">Renomear</button>
      <button class="mini" data-tog="${c.id}">${c.ativo === false ? "Reativar" : "Inativar"}</button>
      ${supervisor() ? `<button class="mini del" data-del="${c.id}" title="Excluir">✕</button>` : ""}</div></li>`;
  el.innerHTML = `
    <div class="card"><h2>Novo colaborador</h2>
      <div class="row"><input id="cNome" maxlength="60" placeholder="Nome (ex.: JOÃO SILVA)"><button class="mini pri" id="cAdd" style="flex:0 0 auto;margin-bottom:10px">Adicionar</button></div>
      <p class="muted small" style="margin:0">Dica: na tela da atividade também dá para cadastrar digitando um nome novo.</p></div>
    <div class="card"><h2>Colaboradores ativos (${ativos.length})</h2>${ativos.length ? `<ul class="list">${ativos.map(item).join("")}</ul>` : `<div class="empty">Nenhum ainda.</div>`}</div>
    ${inativos.length ? `<div class="card"><h2>Inativos (${inativos.length})</h2><p class="muted">Não aparecem para escolher nas atividades, mas o histórico deles continua.</p><ul class="list">${inativos.map(item).join("")}</ul></div>` : ""}`;
  const add = () => {
    const nome = $("#cNome").value.trim().replace(/\s+/g, " ").toUpperCase();
    if (nome.length < 2) return toast("Digite o nome");
    if (S.colaboradores.some(c => norm(c.nome) === norm(nome))) return toast("Esse colaborador já está cadastrado");
    semEsperar(setDoc(doc(collection(db, "colaboradores")), { nome, ativo: true, criadoEm: serverTimestamp() }));
    toast(`${nome} cadastrado`);
    $("#cNome").value = "";
  };
  $("#cAdd").onclick = add;
  $("#cNome").onkeydown = e => { if (e.key === "Enter") add(); };
  $$("[data-ren]", el).forEach(b => b.onclick = () => {
    const c = S.colaboradores.find(x => x.id === b.dataset.ren);
    const nome = (prompt("Novo nome:", c.nome) || "").trim().replace(/\s+/g, " ").toUpperCase();
    if (nome.length < 2 || nome === c.nome) return;
    semEsperar(updateDoc(doc(db, "colaboradores", c.id), { nome }));
    toast("Nome alterado (as atividades antigas mantêm o nome da época)", 4000);
  });
  $$("[data-tog]", el).forEach(b => b.onclick = () => {
    const c = S.colaboradores.find(x => x.id === b.dataset.tog);
    semEsperar(updateDoc(doc(db, "colaboradores", c.id), { ativo: c.ativo === false }));
  });
  $$("[data-del]", el).forEach(b => b.onclick = () => {
    const c = S.colaboradores.find(x => x.id === b.dataset.del);
    if (confirm(`Excluir ${c.nome} do cadastro? As atividades antigas continuam com o nome dele. (Se ele só saiu da equipe, prefira "Inativar".)`)) semEsperar(deleteDoc(doc(db, "colaboradores", c.id)));
  });
}

// Locais da planilha original, com as pessoas previstas de cada um
const LOCAIS_PADRAO = [
  ["Rua A", "Rua", 1], ["Rua B", "Rua", 1], ["Rua C", "Rua", 0], ["Rua D", "Rua", 1], ["Rua E", "Rua", 1], ["Rua F", "Rua", 0],
  ["Rua G", "Rua", 1], ["Rua H", "Rua", 1], ["Rua I", "Rua", 1], ["1º Jirau", "Jirau", 2], ["2º Jirau", "Jirau", 3], ["3º Jirau", "Jirau", 2],
  ["Sala Nobre", "Sala Nobre", 3],
];
async function semearLocais() {
  const b = writeBatch(db);
  LOCAIS_PADRAO.forEach(([nome, tipo, previstas], i) => b.set(doc(collection(db, "locais")), { nome, tipo, previstas, ordem: (i + 1) * 10, ativo: true }));
  await b.commit();
}
function subLocais(el) {
  const sup = pode("locais");
  el.innerHTML = `
    ${sup ? `<div class="linha-btns"><button class="btn" id="lNovo">＋ Novo local</button>${S.locais.length ? "" : `<button class="btn sec" id="lPadrao">Criar os locais da planilha</button>`}</div>` : ""}
    <div class="card"><h2>Locais (${S.locais.length})</h2>
      ${S.locais.length ? `<ul class="list">${S.locais.map((l, i) => `<li class="item"><div class="info">
          <div class="title">${esc(l.nome)} <span class="pill azul">${esc(l.tipo || "—")}</span> ${l.ativo === false ? `<span class="pill off">Inativo</span>` : ""}</div>
          <div class="muted">Pessoas previstas: ${l.previstas ?? 0} · Meta: ${num(l.metaIph) > 0 ? `<b>${fmtN(l.metaIph, 1)} itens/h</b> (fixa)` : S.metas?.valores?.[l.id]?.iph ? `${fmtN(S.metas.valores[l.id].iph, 1)} itens/h (histórico)` : "ainda sem histórico"}</div></div>
        ${sup ? `<div class="acts">
          <button class="mini" data-sobe="${i}" ${i === 0 ? "disabled" : ""} title="Subir">↑</button>
          <button class="mini" data-desce="${i}" ${i === S.locais.length - 1 ? "disabled" : ""} title="Descer">↓</button>
          <button class="mini" data-edit="${l.id}">Editar</button>
          <button class="mini" data-tog="${l.id}">${l.ativo === false ? "Reativar" : "Inativar"}</button>
          <button class="mini del" data-del="${l.id}" title="Excluir">✕</button></div>` : ""}</li>`).join("")}</ul>`
        : `<div class="empty">Nenhum local cadastrado.${sup ? "" : " Peça ao supervisor para cadastrar."}</div>`}
      ${sup ? `<p class="muted small">A ordem aqui é a ordem em que os locais aparecem no dia e no resumo.</p>` : ""}</div>`;
  if (!sup) return;
  $("#lNovo").onclick = () => formLocal(null);
  $("#lPadrao") && ($("#lPadrao").onclick = async () => { ocupado("Criando…"); if (await tentar(async () => { await semearLocais(); return true; })) toast("Locais criados"); });
  $$("[data-edit]", el).forEach(b => b.onclick = () => formLocal(S.locais.find(l => l.id === b.dataset.edit)));
  $$("[data-tog]", el).forEach(b => b.onclick = () => { const l = S.locais.find(x => x.id === b.dataset.tog); semEsperar(updateDoc(doc(db, "locais", l.id), { ativo: l.ativo === false })); });
  $$("[data-del]", el).forEach(b => b.onclick = () => {
    const l = S.locais.find(x => x.id === b.dataset.del);
    if (confirm(`Excluir o local ${l.nome}? As atividades antigas continuam com o nome dele. (Para só esconder, use "Inativar".)`)) semEsperar(deleteDoc(doc(db, "locais", l.id)));
  });
  const trocar = (i, j) => {
    const b = writeBatch(db), L = S.locais;
    // renumera tudo para a ordem ficar sempre limpa
    const nova = [...L]; [nova[i], nova[j]] = [nova[j], nova[i]];
    nova.forEach((l, k) => { if (l.ordem !== (k + 1) * 10) b.update(doc(db, "locais", l.id), { ordem: (k + 1) * 10 }); });
    semEsperar(b.commit());
  };
  $$("[data-sobe]", el).forEach(b => b.onclick = () => trocar(+b.dataset.sobe, +b.dataset.sobe - 1));
  $$("[data-desce]", el).forEach(b => b.onclick = () => trocar(+b.dataset.desce, +b.dataset.desce + 1));
}
function formLocal(l) {
  const tipos = [...new Set(["Rua", "Jirau", "Sala Nobre", ...S.locais.map(x => x.tipo).filter(Boolean)])];
  const m = modal(`<h2>${l ? "Editar local" : "Novo local"}</h2>
    <label for="flNome">Nome</label><input id="flNome" maxlength="40" value="${esc(l?.nome || "")}" placeholder="Ex.: Rua J">
    <label for="flTipo">Tipo</label>
    <select id="flTipo">${tipos.map(t => `<option ${t === (l?.tipo || "Rua") ? "selected" : ""}>${esc(t)}</option>`).join("")}<option value="__novo">Outro tipo…</option></select>
    <input id="flTipoNovo" class="hidden" maxlength="30" placeholder="Nome do novo tipo">
    <label for="flPrev">Pessoas previstas</label><input id="flPrev" type="number" min="0" step="1" inputmode="numeric" value="${l?.previstas ?? 1}">
    <label for="flMeta">Meta fixa de itens por hora (opcional)</label><input id="flMeta" type="number" min="0" step="0.1" inputmode="decimal" value="${num(l?.metaIph) > 0 ? l.metaIph : ""}" placeholder="${l && S.metas?.valores?.[l.id]?.iph ? "histórico: " + fmtN(S.metas.valores[l.id].iph, 1) : "vazio = calcula pelo histórico"}">
    <p class="muted small" style="margin-top:-4px">Deixe vazio para o app calcular sozinho pela média das últimas 4 semanas.</p>
    <button class="btn" id="flOk">Salvar</button><button class="btn sec" data-fechar>Cancelar</button>`);
  const s = m.sheet;
  $("#flTipo", s).onchange = e => $("#flTipoNovo", s).classList.toggle("hidden", e.target.value !== "__novo");
  $("#flOk", s).onclick = () => {
    const nome = $("#flNome", s).value.trim().replace(/\s+/g, " ");
    const tipo = $("#flTipo", s).value === "__novo" ? $("#flTipoNovo", s).value.trim() : $("#flTipo", s).value;
    const previstas = lerInt($("#flPrev", s).value) ?? 0;
    if (nome.length < 1) return toast("Digite o nome do local");
    if (!tipo) return toast("Digite o tipo");
    if (Number.isNaN(previstas)) return toast("Pessoas previstas: use um número inteiro");
    if (S.locais.some(x => x.id !== l?.id && norm(x.nome) === norm(nome))) return toast("Já existe um local com esse nome");
    const metaTxt = $("#flMeta", s).value.trim().replace(",", ".");
    const metaIph = metaTxt ? Number(metaTxt) : null;
    if (metaTxt && !(metaIph > 0)) return toast("Meta: use um número maior que zero (ou deixe vazio)");
    if (l) semEsperar(updateDoc(doc(db, "locais", l.id), { nome, tipo, previstas, metaIph }));
    else semEsperar(setDoc(doc(collection(db, "locais")), { nome, tipo, previstas, metaIph, ordem: (Math.max(0, ...S.locais.map(x => x.ordem || 0)) + 10), ativo: true }));
    m.fechar();
    toast("Local salvo");
  };
}

/* ---------- Usuários do app (supervisor) ---------- */
// Cria a conta de login sem desconectar o supervisor (usa um 2º "app" só na memória)
let AUTH_CRIADOR = null;
async function criarConta(login, senha) {
  if (!AUTH_CRIADOR) {
    const a = getApps().find(x => x.name === "criador") || initializeApp(C.firebase, "criador");
    AUTH_CRIADOR = getAuth(a);
    await setPersistence(AUTH_CRIADOR, inMemoryPersistence);
  }
  const email = emailDe(login);
  try { return { uid: (await createUserWithEmailAndPassword(AUTH_CRIADOR, email, senha)).user.uid, email }; }
  finally { signOut(AUTH_CRIADOR).catch(() => {}); }
}
function subUsuarios(el) {
  el.innerHTML = `
    <div class="linha-btns"><button class="btn" id="uNovo">＋ Novo usuário</button></div>
    <div class="card"><h2>Quem pode usar o app (${S.usuarios.length})</h2>
      <ul class="list">${S.usuarios.map(u => `<li class="item"><div class="info">
        <div class="title">${esc(u.nome)} ${u.id === S.user.uid ? `<span class="pill ok">Você</span>` : ""}
          <span class="pill ${u.papel === "supervisor" ? "azul" : "off"}">${u.papel === "supervisor" ? "👑 Administrador" : "🔒 " + esc(nomePapel(u))}</span>
          ${u.ativo === false ? `<span class="pill bad">Desativado</span>` : ""}</div>
        <div class="muted">👤 ${esc(u.login)}${u.papel !== "supervisor" ? ` · pode: ${esc(PERMISSOES.filter(p => permsDe(u)[p[0]]).map(p => p[1].toLowerCase()).join(", ") || "só ver o dia")}` : ""}</div></div>
        <div class="acts">${u.id === S.user.uid ? `<button class="mini" data-eu="1">Minha senha</button>` : `<button class="mini" data-uedit="${u.id}">✏️ Editar / senha</button>`}</div></li>`).join("")}</ul>
      <p class="muted small"><b>Administrador</b>: faz tudo, inclusive usuários e ajustes da empresa. <b>Restrito</b>: só o que você marcar (há modelos prontos: Líder, Operador e Só consulta).</p></div>`;
  $("#uNovo").onclick = () => formUsuario(null);
  $$("[data-uedit]", el).forEach(b => b.onclick = () => formUsuario(S.usuarios.find(u => u.id === b.dataset.uedit)));
  $$("[data-eu]", el).forEach(b => b.onclick = minhaConta);
}
function formUsuario(u) {
  const permsIni = u ? permsDe(u) : Object.fromEntries(MODELOS_PERM.lider.perms.map(k => [k, true]));
  const m = modal(`<h2>${u ? "Editar usuário" : "Novo usuário"}</h2>
    <label for="fuNome">Nome</label><input id="fuNome" maxlength="60" value="${esc(u?.nome || "")}">
    <label>Tipo de acesso</label>
    <div class="tipo-acesso">
      <label class="opcao"><input type="radio" name="fuPapel" value="supervisor" ${u?.papel === "supervisor" ? "checked" : ""}><span><b>👑 Administrador</b><small>Faz tudo: atividades, relatórios, cadastros, usuários e ajustes.</small></span></label>
      <label class="opcao"><input type="radio" name="fuPapel" value="restrito" ${u?.papel !== "supervisor" ? "checked" : ""}><span><b>🔒 Com restrições</b><small>Só faz o que você marcar abaixo.</small></span></label>
    </div>
    <div id="fuPerms">
      <div class="filtros" style="margin-top:4px">${Object.entries(MODELOS_PERM).map(([k, mo]) => `<button type="button" data-modelo="${k}">${mo.nome}</button>`).join("")}</div>
      <div class="perms">${PERMISSOES.map(([k, t, d]) => `<label class="perm"><input type="checkbox" data-perm="${k}" ${permsIni[k] ? "checked" : ""}><span>${t}${d ? `<small>${d}</small>` : ""}</span></label>`).join("")}</div>
      <p class="muted small">Todo usuário vê a tela do Dia. Os botões que ele não pode usar não aparecem, e o banco de dados também bloqueia.</p>
    </div>
    ${u ? `<label style="display:flex;gap:8px;align-items:center;color:var(--text)"><input type="checkbox" id="fuAtivo" ${u.ativo !== false ? "checked" : ""} style="width:auto;margin:0"> Acesso ativo</label>
      <button class="btn" id="fuSalvar" style="margin-top:12px">Salvar</button>
      <h3>🔑 Usuário e senha</h3><p class="muted">Usuário atual: <b>${esc(u.login)}</b>. Preencha para criar um acesso novo (ex.: esqueceu a senha). O acesso antigo para de funcionar.</p>` : ""}
    <label for="fuUser">Usuário</label><input id="fuUser" maxlength="30" autocapitalize="none" autocorrect="off" spellcheck="false" value="${esc(u?.login || "")}" placeholder="Ex.: joao.lider">
    <label for="fuSenha">${u ? "Nova senha" : "Senha"} (mínimo 6 caracteres)</label><input id="fuSenha" type="password" autocomplete="new-password">
    <button class="btn ${u ? "sec" : ""}" id="fuAcesso">${u ? "Salvar novo acesso" : "Criar usuário"}</button>
    ${u ? `<button class="btn sec" id="fuExcluir" style="color:var(--bad)">Remover usuário</button>` : ""}
    <button class="btn sec" data-fechar>Fechar</button>`);
  const s = m.sheet;
  olhoSenha("fuSenha", s);
  $("#fuUser", s).oninput = e => { e.target.value = normLogin(e.target.value); };
  const nomeOk = () => { const n = $("#fuNome", s).value.trim().replace(/\s+/g, " "); if (n.length < 2) { toast("Digite o nome"); return null; } return n; };
  // tipo de acesso + permissões marcadas
  const mostrarPerms = () => $("#fuPerms", s).classList.toggle("hidden", $('input[name="fuPapel"]:checked', s).value === "supervisor");
  $$('input[name="fuPapel"]', s).forEach(r => r.onchange = mostrarPerms); mostrarPerms();
  $$("[data-modelo]", s).forEach(b => b.onclick = () => { const mo = MODELOS_PERM[b.dataset.modelo].perms; $$("[data-perm]", s).forEach(c => { c.checked = mo.includes(c.dataset.perm); }); });
  const lerPerfil = () => {
    const papel = $('input[name="fuPapel"]:checked', s).value;
    const perms = papel === "supervisor" ? {} : Object.fromEntries($$("[data-perm]", s).map(c => [c.dataset.perm, c.checked]));
    return { papel, perms };
  };
  $("#fuSalvar", s) && ($("#fuSalvar", s).onclick = () => {
    const nome = nomeOk(); if (!nome) return;
    semEsperar(updateDoc(doc(db, "usuarios", u.id), { nome, ...lerPerfil(), ativo: $("#fuAtivo", s).checked }));
    m.fechar(); toast("Usuário atualizado");
  });
  $("#fuExcluir", s) && ($("#fuExcluir", s).onclick = () => {
    if (!confirm(`Remover o acesso de ${u.nome}? Ele não consegue mais entrar.`)) return;
    const b = writeBatch(db);
    b.delete(doc(db, "usuarios", u.id));
    if (u.login) b.delete(loginRef(u.login));
    semEsperar(b.commit());
    m.fechar(); toast("Usuário removido");
  });
  $("#fuAcesso", s).onclick = async () => {
    const nome = nomeOk(); if (!nome) return;
    const login = normLogin($("#fuUser", s).value), senha = $("#fuSenha", s).value, { papel, perms } = lerPerfil();
    if (login.length < 3) return toast("O usuário precisa ter pelo menos 3 letras/números");
    if (senha.length < 6) return toast("A senha precisa ter pelo menos 6 caracteres");
    if (u && !confirm(`Criar um acesso novo para ${u.nome}? O usuário/senha antigos deixam de funcionar.`)) return;
    ocupado(u ? "Trocando o acesso…" : "Criando…");
    const ok = await tentar(async () => {
      const lk = await getDoc(loginRef(login));
      if (lk.exists() && lk.data().uid !== u?.id) throw new Error("Esse usuário já existe. Escolha outro.");
      const conta = await criarConta(login, senha);
      const b = writeBatch(db);
      const { id, ...resto } = u || {};
      b.set(doc(db, "usuarios", conta.uid), u
        ? { ...resto, nome, papel, perms, ativo: $("#fuAtivo", s).checked, login }
        : { nome, login, papel, perms, ativo: true, criadoEm: serverTimestamp() });
      b.set(loginRef(login), { email: conta.email, uid: conta.uid });
      if (u) { b.delete(doc(db, "usuarios", u.id)); if (u.login && u.login !== login) b.delete(loginRef(u.login)); }
      await b.commit();
      return true;
    }, "Não foi possível salvar o acesso. Tente de novo.");
    if (!ok) return;
    m.sheet.innerHTML = `<h2>✅ Acesso pronto</h2>
      <p>${esc(nome)}<br>Usuário: <b>${esc(login)}</b><br>Senha: <b>${esc(senha)}</b></p>
      <p class="muted">Anote ou envie agora: a senha não aparece de novo. Endereço do app: ${esc(location.origin + location.pathname)}</p>
      <button class="btn" id="fuCopiar">📋 Copiar</button><button class="btn sec" data-fechar>Concluir</button>`;
    $("#fuCopiar", m.sheet).onclick = async () => {
      try { await navigator.clipboard.writeText(`Acesso ao Controle de Atividades\n${location.origin + location.pathname}\nUsuário: ${login}\nSenha: ${senha}`); toast("Copiado"); }
      catch (e) { toast("Não deu para copiar: anote o usuário e a senha"); }
    };
  };
}

function subEmpresa(el) {
  el.innerHTML = `<div class="card"><h2>🏢 Empresa / setor</h2>
    <label for="eNome">Nome que aparece no topo e nos relatórios</label>
    <input id="eNome" maxlength="60" value="${esc(S.config?.empresa || "")}">
    <div class="row"><div><label for="eIni">Início do turno</label><input type="time" id="eIni" value="${esc(cfg("turnoInicio"))}"></div>
      <div><label for="eFim">Fim do turno (relatório do dia)</label><input type="time" id="eFim" value="${esc(cfg("turnoFim"))}"></div></div>
    <label for="eWhats">WhatsApp do gestor (recebe o relatório do dia)</label>
    <input id="eWhats" type="tel" inputmode="tel" maxlength="20" placeholder="Ex.: 11 91234-5678" value="${esc(cfg("whatsGestor"))}">
    <h3>🔔 Quando avisar o supervisor</h3>
    <div class="row" style="flex-wrap:wrap">
      <div><label for="ePausa">Pausa maior que (min)</label><input id="ePausa" type="number" min="5" max="240" value="${cfg("limitePausa")}"></div>
      <div><label for="eAlmoco">Almoço maior que (min)</label><input id="eAlmoco" type="number" min="15" max="240" value="${cfg("limiteAlmoco")}"></div>
      <div><label for="eMeta">Terminar abaixo de (% da meta)</label><input id="eMeta" type="number" min="10" max="100" value="${cfg("metaMinPct")}"></div>
    </div>
    <p class="muted small">Também avisa: item que faltou, local planejado que não começou 1 h depois do início do turno, atividade esquecida em andamento e, no fim do turno, que o relatório está pronto.</p>
    <button class="btn" id="eSalvar">Salvar</button></div>
    <div class="card"><h2>🎯 Metas</h2>
      <p class="muted">${S.metas?.t ? `Calculadas em ${dataBR(iso(new Date(ms(S.metas.t))))} com as atividades de ${dataBR(S.metas.de)} a ${dataBR(S.metas.ate)}. Média geral: <b>${fmtN(S.metas.geral, 1)} itens/h</b>.` : "Ainda não calculadas."} O app recalcula sozinho 1 vez por dia. Para fixar a meta de um local, edite o local em Cadastros → Locais.</p>
      <button class="btn sec" id="eMetas">🔄 Recalcular as metas agora</button></div>
    <div class="card"><h2>ℹ️ Sobre</h2><p class="muted">Controle de Atividades e Produtividade · versão 1.1<br>
      Alerta de atividade esquecida em andamento: ${HORAS_ALERTA} h (muda no config.js).</p></div>`;
  $("#eSalvar").onclick = () => {
    const n = (id, min, max) => { const v = Math.round(+$(id).value); return v >= min && v <= max ? v : null; };
    const dados_ = { empresa: $("#eNome").value.trim(), turnoInicio: $("#eIni").value || "07:00", turnoFim: $("#eFim").value || "17:00",
      whatsGestor: $("#eWhats").value.trim(), limitePausa: n("#ePausa", 5, 240), limiteAlmoco: n("#eAlmoco", 15, 240), metaMinPct: n("#eMeta", 10, 100) };
    if (Object.values(dados_).some(v => v === null)) return toast("Confira os números dos alertas");
    semEsperar(updateDoc(doc(db, "config", "sistema"), dados_));
    toast("Salvo");
  };
  $("#eMetas").onclick = async () => { ocupado("Calculando as metas…"); await tentar(() => atualizarMetas(true)); toast("Metas recalculadas"); };
}

/* ================== AJUSTES DO SETOR (turno, limites dos alertas) ================== */
const PADRAO_CFG = { turnoInicio: "07:00", turnoFim: "17:00", limitePausa: 30, limiteAlmoco: 75, metaMinPct: 70, whatsGestor: "" };
const cfg = k => S.config?.[k] ?? PADRAO_CFG[k];
const recebeAlertas = () => { if (!pode("alertas")) return false; const v = getLS("ativ_alertas"); return v ? v === "1" : true; };
const minutosDoDia = hm => { const [h, m] = String(hm || "0:0").split(":").map(Number); return h * 60 + (m || 0); };

/* ================== PLANEJAR O DIA SEGUINTE ================== */
// Horas necessárias = itens previstos ÷ meta de itens/hora do local; pessoas = horas ÷ horas produtivas por pessoa
function abaPlanejar(el) {
  const data = S.planData || somaDias(hojeISO(), 1);
  const horasPessoa = +getLS("ativ_horas_pessoa") || 7;
  const dow = deISO(data).getDay();
  const prev = S.planPrev = S.planPrev || {};
  const geral = S.metas?.geral || null;
  const locais = S.locais.filter(l => l.ativo !== false);
  const sugestao = l => { const v = S.metas?.valores?.[l.id]; return v?.listaDow?.[dow] ?? v?.listaMedia ?? ""; };
  const linhaCalc = l => {
    const itens = lerInt(prev[l.id] ?? sugestao(l)), meta = metaDe(l.id) || geral;
    const horas = itens && meta ? itens / meta : null;
    const pessoas = itens ? Math.max(1, Math.ceil((horas ?? 0) / horasPessoa)) : 0;
    return { itens: Number.isNaN(itens) ? null : itens, meta, horas, pessoas };
  };
  const ativos = S.colaboradores.filter(c => c.ativo !== false).length;
  el.innerHTML = `
    <div class="card">
      <h2>🗓️ Planejar o dia</h2>
      <p class="muted">Informe quantos itens cada local deve ter na lista. O app calcula as horas e quantas pessoas colocar, pela produtividade (meta) de cada local.</p>
      <div class="row" style="flex-wrap:wrap">
        <div><label for="plData">Dia</label><input type="date" id="plData" value="${data}"></div>
        <div><label for="plHoras">Horas produtivas por pessoa</label><input type="number" id="plHoras" min="1" max="12" step="0.5" value="${horasPessoa}"></div>
      </div>
      ${S.metas?.valores ? `<p class="muted small" style="margin:0">Sugestões = média dos itens da lista nas últimas 4 semanas (${diaSemana(data)}). Metas atualizadas em ${dataBR(iso(new Date(ms(S.metas.t))))}.</p>`
        : `<div class="alert info" style="margin:0"><b>Ainda sem histórico</b>Depois de algumas semanas de registros o app sugere os itens e a meta de cada local. Por enquanto, digite os itens; sem meta, use a meta fixa no cadastro do local.</div>`}
    </div>
    <div class="card"><div class="tabela-wrap"><table class="tab" id="plTab">
      <thead><tr><th>Local</th><th class="esq">Tipo</th><th>Itens previstos</th><th>Meta itens/h</th><th>Horas</th><th>Pessoas</th></tr></thead>
      <tbody>${locais.map(l => { const r = linhaCalc(l); return `<tr data-l="${l.id}"><td>${esc(l.nome)}</td><td class="esq">${esc(l.tipo || "")}</td>
        <td><input class="pl-in" type="number" min="0" step="1" inputmode="numeric" data-prev="${l.id}" value="${esc(prev[l.id] ?? sugestao(l))}" placeholder="0"></td>
        <td>${r.meta ? fmtN(r.meta, 1) + (metaDe(l.id) ? "" : "*") : "—"}</td><td data-h>${r.horas != null ? fmtN(r.horas, 1) + " h" : "—"}</td><td data-p><b>${r.pessoas || "—"}</b></td></tr>`; }).join("")}</tbody>
      <tfoot><tr class="total"><td>Total</td><td></td><td id="plTI"></td><td></td><td id="plTH"></td><td id="plTP"></td></tr></tfoot>
    </table></div>
    <p class="muted small">${geral ? `* local sem meta própria: usando a média geral (${fmtN(geral, 1)} itens/h).` : ""} Colaboradores ativos cadastrados: <b>${ativos}</b>.</p>
    <div id="plAviso"></div></div>
    <div class="linha-btns">
      <button class="btn" id="plCriar">Criar as atividades de ${dataBR(data)}</button>
      <button class="btn sec" id="plCopiar">📋 Copiar o texto do planejamento</button>
    </div>`;
  const totaisPl = () => {
    let ti = 0, th = 0, tp = 0;
    locais.forEach(l => {
      const r = linhaCalc(l), tr = $(`tr[data-l="${l.id}"]`, el);
      $("[data-h]", tr).textContent = r.horas != null ? fmtN(r.horas, 1) + " h" : "—";
      $("[data-p]", tr).innerHTML = `<b>${r.pessoas || "—"}</b>`;
      ti += r.itens || 0; th += r.horas || 0; tp += r.pessoas;
    });
    $("#plTI").textContent = fmtN(ti); $("#plTH").textContent = fmtN(th, 1) + " h"; $("#plTP").textContent = tp;
    $("#plAviso").innerHTML = tp > ativos ? `<div class="alert"><b>Faltam pessoas</b>O plano pede ${tp} pessoas e há ${ativos} colaboradores ativos. Priorize os locais com mais itens ou junte locais pequenos para a mesma pessoa.</div>` : "";
    return { ti, th, tp };
  };
  $$("[data-prev]", el).forEach(i => i.oninput = () => { prev[i.dataset.prev] = i.value; totaisPl(); });
  $("#plData").onchange = e => { S.planData = e.target.value; S.planPrev = {}; abaPlanejar(el); };
  $("#plHoras").onchange = e => { const v = +e.target.value; if (v > 0) setLS("ativ_horas_pessoa", String(v)); abaPlanejar(el); };
  totaisPl();
  $("#plCopiar").onclick = async () => {
    const t = totaisPl();
    const linhas = locais.map(l => ({ l, r: linhaCalc(l) })).filter(x => x.r.itens).map(x => `• ${x.l.nome}: ${x.r.itens} itens · ${x.r.pessoas} pessoa(s) (~${fmtN(x.r.horas, 1)} h)`);
    const txt = `🗓️ *Planejamento ${diaSemana(data)}, ${dataBR(data)}*\n${linhas.join("\n")}\n\n*Total:* ${fmtN(t.ti)} itens · ${fmtN(t.th, 1)} h · ${t.tp} pessoa(s)`;
    try { await navigator.clipboard.writeText(txt); toast("Copiado: é só colar no WhatsApp"); } catch (e) { modal(`<h2>Planejamento</h2><textarea style="min-height:240px">${esc(txt)}</textarea><button class="btn sec" data-fechar>Fechar</button>`); }
  };
  $("#plCriar").onclick = async () => {
    const plano = locais.map(l => ({ l, r: linhaCalc(l) })).filter(x => x.r.itens);
    if (!plano.length) return toast("Informe os itens previstos de pelo menos um local");
    if (!confirm(`Criar ${plano.length} atividade(s) planejada(s) em ${dataBR(data)}? Depois é só escolher as pessoas e tocar em ▶ Iniciar.`)) return;
    ocupado("Criando…");
    const r = await tentar(async () => {
      const ja = new Set(lista(await getDocs(query(collection(db, "atividades"), where("data", "==", data)))).map(a => a.localId));
      const novas = plano.filter(x => !ja.has(x.l.id));
      const b = writeBatch(db);
      novas.forEach(({ l, r }) => b.set(doc(collection(db, "atividades")), {
        data, localId: l.id, localNome: l.nome, tipo: l.tipo || "", pessoas: [], pessoaIds: [], sugeridas: r.pessoas,
        inicio: null, fim: null, itensLista: r.itens, itensAbastecidos: null, obs: "", pausas: [],
        criadoEm: serverTimestamp(), criadoPor: S.perfil.nome, atualizadoEm: serverTimestamp(), atualizadoPor: S.perfil.nome,
      }));
      if (novas.length) await b.commit();
      return { novas: novas.length, ja: plano.length - novas.length };
    });
    if (!r) return;
    S.cache = null;
    toast(`${r.novas} atividade(s) criada(s)${r.ja ? ` · ${r.ja} local(is) já tinham atividade` : ""}`, 4500);
    S.aba = "dia"; S.dia = data; S.atividades = []; ouvirDia(); go(telaPainel);
  };
}

/* ================== MODO TV (telão do setor) ================== */
let WAKE = null;
function abrirTV() {
  if (S.dia !== hojeISO()) { S.dia = hojeISO(); S.atividades = []; ouvirDia(); }
  document.body.classList.add("tv");
  document.documentElement.requestFullscreen?.().catch(() => {});
  navigator.wakeLock?.request("screen").then(w => { WAKE = w; }).catch(() => {});   // não deixa a tela apagar
  go(telaTV);
}
function sairTV() {
  document.body.classList.remove("tv");
  if (document.fullscreenElement) document.exitFullscreen?.().catch(() => {});
  WAKE?.release?.().catch(() => {}); WAKE = null;
  go(telaPainel);
}
function telaTV() {
  S.live = true;
  if (S.dia !== hojeISO()) { S.dia = hojeISO(); S.atividades = []; ouvirDia(); }   // virou o dia
  const A = S.atividades, T = totais(A), agora = new Date();
  const trabalhando = new Set(A.filter(a => calc(a).status === "andamento").flatMap(a => (a.pessoas || []).map(p => p.id || p.nome))).size;
  const paradas = A.filter(a => calc(a).status === "pausada").length;
  const prog = T.lista ? Math.min(1, T.cobertos / T.listaCob || T.abast / T.lista) : 0;
  const card = l => {
    const as = A.filter(a => a.localId === l.id), t = totais(as), cs = as.map(calc);
    const st = !as.length ? ["Sem atividade", "off"] : cs.some(c => c.status === "pausada") ? ["⏸ " + (pausaAberta(as[cs.findIndex(c => c.status === "pausada")])?.motivo || "Parado"), "warn"]
      : cs.some(c => c.status === "andamento") ? ["Em andamento", "azul"] : cs.every(c => c.status === "concluida") ? ["Concluído", "ok"] : ["A começar", "off"];
    const p = t.lista ? Math.min(100, Math.round((t.abast || 0) / t.lista * 100)) : 0;
    return `<div class="tv-local st-${st[1]}">
      <div class="tv-l1"><b>${esc(l.nome)}</b><span class="pill ${st[1]}">${esc(st[0])}</span></div>
      <div class="tv-pess">${esc([...new Set(as.flatMap(a => (a.pessoas || []).map(x => primeiroNome(x.nome))))].join(", ") || "—")}</div>
      <div class="barra-prog"><i style="width:${p}%;background:${t.pend ? "var(--bad)" : "var(--ok)"}"></i></div>
      <div class="tv-l3"><span>${fmtN(t.abast)} / ${fmtN(t.lista)} itens</span><span style="color:${corMeta(t.pctMeta)}">${t.iph != null ? fmtN(t.iph, 1) + "/h" : ""} ${t.pctMeta != null ? "· " + pct(t.pctMeta) : ""}</span></div>
    </div>`;
  };
  const tipos = [...new Set(S.locais.filter(l => l.ativo !== false).map(l => l.tipo || "Outros"))];
  app.innerHTML = `<div class="tv-tela">
    <div class="tv-topo"><div><b>📋 ${esc(S.config?.empresa || "Controle de Atividades")}</b><span>${diaSemana(S.dia)}, ${dataBR(S.dia)}</span></div>
      <div class="tv-hora">${pad(agora.getHours())}:${pad(agora.getMinutes())}</div><button class="tv-sair" id="tvSair">✕ Sair</button></div>
    <div class="tv-kpis">
      <div><span>Abastecidos</span><b>${fmtN(T.abast)} <small>/ ${fmtN(T.lista)}</small></b><div class="barra-prog"><i style="width:${Math.round(prog * 100)}%;background:var(--ok)"></i></div></div>
      <div><span>Itens por hora</span><b style="color:${corMeta(T.pctMeta)}">${fmtN(T.iph, 1)}</b><small>${T.pctMeta != null ? pct(T.pctMeta) + " da meta" : "sem meta"}</small></div>
      <div><span>Pendentes</span><b style="color:${T.pend ? "var(--bad)" : "var(--ok)"}">${fmtN(T.pend)}</b><small>${T.pend ? "faltaram" : "nada pendente"}</small></div>
      <div><span>Trabalhando agora</span><b>${trabalhando}</b><small>${paradas ? `⏸ ${paradas} parado(s)` : "ninguém parado"}</small></div>
    </div>
    ${tipos.map(tp => `<div class="tv-grupo"><div class="tv-gtit">${esc(tp)} ${(() => { const t = totais(A.filter(a => (a.tipo || "Outros") === tp)); return t.pctMeta != null ? `<span style="color:${corMeta(t.pctMeta)}">equipe ${pct(t.pctMeta)} da meta</span>` : ""; })()}</div>
      <div class="tv-grade">${S.locais.filter(l => l.ativo !== false && (l.tipo || "Outros") === tp).map(card).join("")}</div></div>`).join("")}
  </div>`;
  $("#tvSair").onclick = sairTV;
}

/* ================== ALERTAS AO SUPERVISOR ================== */
// Rodam no aparelho de quem recebe alertas, com o app aberto ou minimizado (plano grátis: sem servidor)
let RELOGIO = null;
function iniciarRelogio() {
  if (RELOGIO) return;
  S.alertasDesde = Date.now();
  ouvirHoje();
  RELOGIO = setInterval(() => {
    if (!S.perfil) return;
    checarAlertas();
    if (S.view?.fn === telaTV) telaTV();
    else if (S.view?.fn === telaPainel && S.aba === "dia" && S.dia === hojeISO()) reRender();   // atualiza os tempos correndo
  }, 60000);
}
// Atividades de hoje sempre ouvidas (mesmo olhando outro dia), para os alertas
function ouvirHoje() {
  S.subHoje?.();
  const dia = hojeISO();
  S.subHoje = onSnapshot(query(collection(db, "atividades"), where("data", "==", dia)), s => {
    S.ativHoje = lista(s); S.ativHojeDia = dia; checarAlertas();
  }, e => console.warn("alertas:", e));
}
const chaveAlertas = () => "ativ_alertas_" + hojeISO();
function carregarAlertas() {
  if (S.alertasDia === hojeISO()) return;
  S.alertasDia = hojeISO();
  try { S.alertas = JSON.parse(getLS(chaveAlertas()) || "[]"); } catch (e) { S.alertas = []; }
}
function checarAlertas() {
  if (!S.perfil || !recebeAlertas()) return;
  if (S.ativHojeDia !== hojeISO()) { ouvirHoje(); return; }
  carregarAlertas();
  const agora = Date.now(), minAgora = new Date().getHours() * 60 + new Date().getMinutes();
  const novos = [];
  const add = (k, txt, acao) => { if (!S.alertas.some(x => x.k === k)) novos.push({ k, txt, t: agora, acao: acao || "", visto: false }); };
  (S.ativHoje || []).forEach(a => {
    const c = calc(a), pa = pausaAberta(a);
    const quem = (a.pessoas || []).map(p => primeiroNome(p.nome)).join(", ");
    if (pa) {
      const lim = /almo|interv/i.test(pa.motivo) ? +cfg("limiteAlmoco") : +cfg("limitePausa");
      const parado = (agora - ms(pa.ini)) / 60000;
      if (parado > lim) add(`pausa:${a.id}:${ms(pa.ini)}`, `⏸ ${a.localNome} parado há ${fmtDur(parado)} (${pa.motivo}) · ${quem}`);
    }
    if (c.status === "andamento" && c.corrido / 60 > HORAS_ALERTA) add(`longa:${a.id}`, `⏰ ${a.localNome} em andamento há mais de ${HORAS_ALERTA} h · ${quem}`);
    if (c.status === "concluida" && c.pctMeta != null && c.pctMeta * 100 < +cfg("metaMinPct"))
      add(`meta:${a.id}`, `📉 ${a.localNome} terminou com ${pct(c.pctMeta)} da meta (${fmtN(c.iph, 1)} de ${fmtN(c.meta, 1)} itens/h) · ${quem}`);
    c.faltas.forEach(i => add(`falta:${a.id}:${norm(i.desc)}`, `🔻 Faltou na ${a.localNome}: ${i.desc} (${i.motivo || "sem motivo"})`));
    if (c.status === "concluida" && a.motivoFalta && c.pend) add(`faltaA:${a.id}`, `🔻 ${a.localNome}: faltaram ${c.pend} item(ns) (${a.motivoFalta})`);
    if (c.status === "planejada" && minAgora > minutosDoDia(cfg("turnoInicio")) + 60) add(`semini:${a.id}`, `🕗 ${a.localNome} ainda não começou`);
  });
  if (minAgora >= minutosDoDia(cfg("turnoFim")) && (S.ativHoje || []).length) add(`rel:${hojeISO()}`, "📤 O relatório do dia está pronto. Toque para enviar.", "rel");
  if (!novos.length) return;
  // ao abrir o app (primeiros 30 s), o que já estava acontecendo entra na lista sem tocar alarme
  const abrindo = Date.now() - (S.alertasDesde || 0) < 30000;
  S.alertas.push(...novos);
  setLS(chaveAlertas(), JSON.stringify(S.alertas.slice(-100)));
  if (!abrindo) {
    toast(novos[0].txt, 6000);
    try { navigator.vibrate?.([200, 100, 200]); } catch (e) {}
    novos.forEach(n => notificar(n.txt));
  }
  atualizarSino();
}
function notificar(txt) {
  try {
    if (!("Notification" in window) || Notification.permission !== "granted") return;
    const op = { body: txt, icon: "icon-192.png", tag: txt.slice(0, 40) };
    if (navigator.serviceWorker?.controller) navigator.serviceWorker.ready.then(r => r.showNotification("Controle de Atividades", op)).catch(() => {});
    else new Notification("Controle de Atividades", op);
  } catch (e) {}
}
function atualizarSino() {
  const b = $$("#hBtns button").find(x => x.textContent.startsWith("🔔"));
  const n = S.alertas.filter(x => !x.visto).length;
  if (b) { b.textContent = `🔔${n ? " " + n : ""}`; b.classList.toggle("sino-on", n > 0); }
  else if (S.view?.fn === telaPainel) telaPainel();
}
function abrirAlertas() {
  carregarAlertas();
  const lista_ = [...S.alertas].reverse();
  const m = modal(`<h2>🔔 Alertas de hoje</h2>
    ${"Notification" in window && Notification.permission === "default" ? `<button class="btn sec" id="alNotif">📲 Avisar também nas notificações do aparelho</button>` : ""}
    ${lista_.length ? `<ul class="list">${lista_.map((x, i) => `<li class="item ${x.acao ? "click" : ""}" data-al="${i}"><div class="info"><div class="title" style="font-weight:${x.visto ? 400 : 700}">${esc(x.txt)}</div><div class="muted small">${hhmm(x.t)}</div></div></li>`).join("")}</ul>`
      : `<div class="empty">Nenhum alerta hoje 👍</div>`}
    <p class="muted small">Os alertas chegam com o app aberto ou minimizado. Limites em Cadastros → Empresa.</p>
    <button class="btn sec" data-fechar>Fechar</button>`);
  S.alertas.forEach(x => { x.visto = true; });
  setLS(chaveAlertas(), JSON.stringify(S.alertas));
  atualizarSino();
  $("#alNotif", m.sheet) && ($("#alNotif", m.sheet).onclick = () => Notification.requestPermission().then(p => { toast(p === "granted" ? "Notificações ativadas" : "As notificações foram bloqueadas no navegador"); m.fechar(); }));
  $$("[data-al]", m.sheet).forEach(li => li.onclick = () => { const x = lista_[+li.dataset.al]; if (x.acao === "rel") { m.fechar(); modalRelatorio(hojeISO(), S.ativHoje || []); } });
}

/* ================== RELATÓRIO DO DIA (WhatsApp) ================== */
function textoRelatorio(dataIso, ativs) {
  const T = totais(ativs), G = porLocal(ativs), P = porColaborador(ativs);
  const faltas = ativs.flatMap(a => (a.itens || []).filter(i => i.status === "falta").map(i => `• ${a.localNome}: ${i.desc} (${i.motivo || "sem motivo"})`))
    .concat(ativs.filter(a => !(a.itens || []).length && calc(a).pend > 0).map(a => `• ${a.localNome}: faltaram ${calc(a).pend} de ${calc(a).lista} (${a.motivoFalta || "sem motivo"})`));
  const paradas = Object.entries(T.porMotivo).sort((a, b) => b[1] - a[1]).map(([k, v]) => `${k} ${fmtDur(v)}`);
  const equipes = [...new Set(ativs.map(a => a.tipo || "Outros"))].map(tp => ({ tp, t: totais(ativs.filter(a => (a.tipo || "Outros") === tp)) })).filter(x => x.t.iph != null).sort((a, b) => (b.t.pctMeta ?? 0) - (a.t.pctMeta ?? 0));
  const abertas = T.andamento + T.pausadas + T.planejadas;
  return [
    `📋 *Relatório do dia${S.config?.empresa ? " – " + S.config.empresa : ""}*`,
    `📅 ${diaSemana(dataIso)}, ${dataBR(dataIso)}`,
    ``,
    `✅ *Abastecidos:* ${fmtN(T.abast)} de ${fmtN(T.lista)} da lista (${pct(T.cobertura)})${T.extra ? ` · +${fmtN(T.extra)} além da lista` : ""}`,
    `${T.pend ? "⚠️" : "👍"} *Pendentes:* ${fmtN(T.pend)} item(ns)`,
    `⚡ *Produtividade:* ${fmtN(T.iph, 1)} itens/h${T.pctMeta != null ? ` (${pct(T.pctMeta)} da meta)` : ""}`,
    `⏱️ *Tempo trabalhado:* ${fmtDur(T.min)}${T.parado >= 1 ? ` · parado ${fmtDur(T.parado)}` : ""}`,
    `👷 *Pessoas:* ${T.nPessoas} · *Atividades:* ${T.registros} (${T.concluidas} concluídas${abertas ? `, ${abertas} em aberto` : ""})`,
    ``,
    `📍 *Por local*`,
    ...G.map(g => {
      const t = g.t, aberto = t.andamento + t.pausadas, novo = t.planejadas;
      if (!t.concluidas && t.iph == null) return `• ${g.nome}: ${aberto ? `⏳ em andamento${t.pausadas ? " (parado)" : ""}` : "🕗 não começou"}${t.lista ? ` · lista ${fmtN(t.lista)}` : ""}`;
      return `• ${g.nome}: ${fmtN(t.abast)}/${fmtN(t.lista)} · ${fmtN(t.iph, 1)} it/h${t.pctMeta != null ? ` (${pct(t.pctMeta)})` : ""}${t.pend ? ` · ⚠️ ${t.pend} pend.` : ""}${aberto || novo ? " · ainda em aberto" : ""}`;
    }),
    ...(faltas.length ? [``, `🔻 *Faltas (${faltas.length})*`, ...faltas.slice(0, 15), ...(faltas.length > 15 ? [`… e mais ${faltas.length - 15}`] : [])] : []),
    ...(paradas.length ? [``, `⏸ *Paradas:* ${paradas.join(" · ")}`] : []),
    ...(equipes.length ? [``, `🏆 *Destaque:* equipe ${equipes[0].tp} – ${fmtN(equipes[0].t.iph, 1)} it/h${equipes[0].t.pctMeta != null ? ` (${pct(equipes[0].t.pctMeta)} da meta)` : ""}`] : []),
    ...(P.length ? [`👏 Mais itens: ${P.slice(0, 3).map(p => `${primeiroNome(p.nome)} (${fmtN(p.abast)})`).join(", ")}`] : []),
  ].join("\n");
}
function modalRelatorio(dataIso, ativs) {
  if (!ativs || !ativs.length) return toast("Não há atividades nesse dia");
  const txt = textoRelatorio(dataIso, ativs);
  const fone = soDigitosBR(cfg("whatsGestor"));
  const m = modal(`<h2>📤 Relatório de ${dataBR(dataIso)}</h2>
    <textarea id="relTxt" style="min-height:300px;font-size:14px">${esc(txt)}</textarea>
    <a class="btn ok" target="_blank" rel="noopener" href="https://wa.me/${fone}?text=${encodeURIComponent(txt)}">💬 Enviar no WhatsApp${fone ? " do gestor" : ""}</a>
    <button class="btn sec" id="relCopiar">📋 Copiar</button>
    ${navigator.share ? `<button class="btn sec" id="relShare">📲 Compartilhar…</button>` : ""}
    <button class="btn sec" data-fechar>Fechar</button>
    <p class="muted small">${fone ? "" : "Dica: cadastre o WhatsApp do gestor em Cadastros → Empresa para o botão já abrir a conversa dele. "}Dá para editar o texto antes de enviar.</p>`);
  const atual = () => $("#relTxt", m.sheet).value;
  $("#relTxt", m.sheet).oninput = () => { $("a.btn", m.sheet).href = `https://wa.me/${fone}?text=${encodeURIComponent(atual())}`; };
  $("#relCopiar", m.sheet).onclick = async () => { try { await navigator.clipboard.writeText(atual()); toast("Copiado"); } catch (e) { $("#relTxt", m.sheet).select(); toast("Selecionei o texto: use Ctrl+C"); } };
  $("#relShare", m.sheet) && ($("#relShare", m.sheet).onclick = () => navigator.share({ text: atual() }).catch(() => {}));
}
const soDigitosBR = t => { let d = String(t || "").replace(/\D/g, ""); if (d && d.length <= 11) d = "55" + d; return d; };

/* ================== EXCEL (baixar e importar) ================== */
let XLSX_P = null;
function carregarXLSX() {
  if (window.XLSX) return Promise.resolve(window.XLSX);
  if (!XLSX_P) XLSX_P = new Promise((ok, erro) => {
    const s = document.createElement("script");
    s.src = XLSX_URL;
    s.onload = () => ok(window.XLSX);
    s.onerror = () => { XLSX_P = null; erro(new Error("Não foi possível carregar o Excel. Confira a internet.")); };
    document.head.appendChild(s);
  });
  return XLSX_P;
}
// Datas e horas no formato do Excel (número de dias desde 30/12/1899)
const serialData = isoS => { const [a, m, d] = isoS.split("-").map(Number); return (Date.UTC(a, m - 1, d) - Date.UTC(1899, 11, 30)) / 864e5; };
const fracHora = t => { const v = ms(t); if (v == null) return null; const d = new Date(v); return (d.getHours() * 60 + d.getMinutes()) / 1440; };
const r2 = n => n == null || isNaN(n) ? null : Math.round(n * 100) / 100;

async function exportarExcel(arr, de, ate) {
  const X = await carregarXLSX();
  const fmt = (ws, r, c, z) => { const ref = X.utils.encode_cell({ r, c }); if (ws[ref] && ws[ref].v != null && ws[ref].v !== "") ws[ref].z = z; };
  const periodo = de === ate ? dataBR(de) : `${dataBR(de)} a ${dataBR(ate)}`;

  // 1) Controle Diário — mesmas colunas da planilha original
  // (as 12 primeiras colunas são iguais às da planilha original; as novas vêm depois)
  const cab = ["Data", "Colaborador", "Tipo", "Local", "Início", "Término", "Tempo na atividade", "Qtd. itens da lista", "Itens abastecidos", "Itens pendentes", "Itens/hora", "Observação", "Tempo parado", "Meta itens/hora", "% da meta", "Motivo da falta"];
  const linhas = arr.map(a => { const c = calc(a); return [serialData(a.data), (a.pessoas || []).map(p => p.nome).join(" E "), a.tipo || "", a.localNome || "", fracHora(c.ini), fracHora(c.fim), c.min != null ? c.min / 1440 : null, c.lista, c.abast, c.pend, r2(c.iph), a.obs || "",
    c.parado >= 1 ? c.parado / 1440 : null, r2(c.meta), c.pctMeta, a.motivoFalta || [...new Set(c.faltas.map(i => i.motivo))].join(", ")]; });
  const ws1 = X.utils.aoa_to_sheet([["CONTROLE DE ATIVIDADES E PRODUTIVIDADE" + (S.config?.empresa ? " · " + S.config.empresa : "")], [`Período: ${periodo}`], cab, ...linhas]);
  linhas.forEach((_, i) => { const r = i + 3; fmt(ws1, r, 0, "dd/mm/yyyy"); fmt(ws1, r, 4, "hh:mm"); fmt(ws1, r, 5, "hh:mm"); fmt(ws1, r, 6, "[h]:mm"); fmt(ws1, r, 10, "0.00"); fmt(ws1, r, 12, "[h]:mm"); fmt(ws1, r, 13, "0.00"); fmt(ws1, r, 14, "0%"); });
  ws1["!merges"] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 11 } }];
  ws1["!cols"] = [11, 28, 11, 13, 8, 8, 10, 10, 10, 10, 9, 45, 10, 10, 9, 18].map(wch => ({ wch }));
  ws1["!autofilter"] = { ref: `A3:P${3 + linhas.length}` };

  // 2) Resumo — por local + indicadores gerais (com as contas corrigidas)
  const T = totais(arr), G = porLocal(arr), umDia = de === ate;
  const res = [["RESUMO DE PRODUTIVIDADE"], [`Período: ${periodo}`], [],
    ["Local", "Tipo", "Pessoas previstas", "Pessoas", "Registros", "Itens da lista", "Itens abastecidos", "Itens pendentes", "Tempo trabalhado", "Itens/hora", "% da lista abastecida", "Tempo parado", "Meta itens/hora", "% da meta"],
    ...G.map(g => [g.nome, g.tipo || "", umDia ? (g.local?.previstas ?? null) : null, g.t.nPessoas, g.t.registros, g.t.lista, g.t.abast, g.t.pend, g.t.min / 1440, r2(g.t.iph), g.t.cobertura, g.t.parado / 1440, r2(metaDe(g.k)), g.t.pctMeta]),
    ["TOTAL", "", umDia ? G.reduce((s, g) => s + (num(g.local?.previstas) || 0), 0) : null, T.nPessoas, T.registros, T.lista, T.abast, T.pend, T.min / 1440, r2(T.iph), T.cobertura, T.parado / 1440, null, T.pctMeta],
    [], ["INDICADORES GERAIS"],
    ["Total de pessoas", T.nPessoas], ["Total de locais", G.length], ["Total de registros", T.registros],
    ["Total de itens da lista", T.lista], ["Total de itens abastecidos", T.abast], ["Itens abastecidos além da lista", T.extra],
    ["Total de itens pendentes", T.pend], ["Tempo trabalhado (sem pausas)", T.min / 1440],
    ["Produtividade média (itens/hora)", r2(T.iph)], ["% da lista abastecida", T.cobertura],
    ["Tempo parado (pausas)", T.parado / 1440], ["% da meta", T.pctMeta]];
  const ws2 = X.utils.aoa_to_sheet(res);
  for (let r = 4; r <= 4 + G.length; r++) { fmt(ws2, r, 8, "[h]:mm"); fmt(ws2, r, 9, "0.00"); fmt(ws2, r, 10, "0%"); fmt(ws2, r, 11, "[h]:mm"); fmt(ws2, r, 12, "0.00"); fmt(ws2, r, 13, "0%"); }
  const base = 4 + G.length + 3;
  fmt(ws2, base + 7, 1, "[h]:mm"); fmt(ws2, base + 8, 1, "0.00"); fmt(ws2, base + 9, 1, "0%"); fmt(ws2, base + 10, 1, "[h]:mm"); fmt(ws2, base + 11, 1, "0%");
  ws2["!cols"] = [32, 12, 10, 9, 10, 12, 12, 12, 11, 10, 12, 11, 11, 9].map(wch => ({ wch }));

  // 3) Por colaborador
  const P = porColaborador(arr);
  const ws3 = X.utils.aoa_to_sheet([["PRODUTIVIDADE POR COLABORADOR"], [`Período: ${periodo} · itens divididos igualmente quando a atividade tem mais de uma pessoa`], [],
    ["Colaborador", "Atividades", "Dias", "Itens abastecidos", "Itens pendentes", "Tempo", "Itens/hora", "Locais"],
    ...P.map(p => [p.nome, p.ativ, p.dias.size, r2(p.abast), r2(p.pend), p.min / 1440, r2(p.iph), [...p.locais].join(", ")])]);
  P.forEach((_, i) => { fmt(ws3, i + 4, 5, "[h]:mm"); fmt(ws3, i + 4, 6, "0.00"); });
  ws3["!cols"] = [28, 10, 7, 14, 12, 9, 10, 40].map(wch => ({ wch }));

  // 4) Paradas (cada pausa com o motivo)
  const pausas = arr.flatMap(a => (a.pausas || []).map(p => {
    const i = ms(p.ini), f = ms(p.fim);
    return [serialData(a.data), a.localNome || "", (a.pessoas || []).map(x => x.nome).join(" E "), p.motivo || "", fracHora(i), f != null ? fracHora(f) : null, f != null && i != null ? (f - i) / 86400000 : null];
  }));
  const ws4 = X.utils.aoa_to_sheet([["PARADAS"], [`Período: ${periodo}`], [], ["Data", "Local", "Colaborador", "Motivo", "Início", "Fim", "Duração"], ...pausas]);
  pausas.forEach((_, i) => { fmt(ws4, i + 4, 0, "dd/mm/yyyy"); fmt(ws4, i + 4, 4, "hh:mm"); fmt(ws4, i + 4, 5, "hh:mm"); fmt(ws4, i + 4, 6, "[h]:mm"); });
  ws4["!cols"] = [11, 14, 28, 24, 8, 8, 9].map(wch => ({ wch }));

  // 5) Faltas (itens que faltaram, com o motivo)
  const faltas = arr.flatMap(a => (a.itens || []).filter(i => i.status === "falta").map(i => [serialData(a.data), a.localNome || "", i.desc, i.motivo || "", i.resolvido ? "Sim" : "Não"]))
    .concat(arr.filter(a => !(a.itens || []).length && calc(a).pend > 0).map(a => [serialData(a.data), a.localNome || "", `${calc(a).pend} item(ns) da lista`, a.motivoFalta || "", ""]));
  const ws5 = X.utils.aoa_to_sheet([["FALTAS"], [`Período: ${periodo}`], [], ["Data", "Local", "Item", "Motivo", "Resolvido"], ...faltas]);
  faltas.forEach((_, i) => fmt(ws5, i + 4, 0, "dd/mm/yyyy"));
  ws5["!cols"] = [11, 14, 40, 20, 10].map(wch => ({ wch }));

  const wb = X.utils.book_new();
  X.utils.book_append_sheet(wb, ws1, "Controle Diário");
  X.utils.book_append_sheet(wb, ws2, "Resumo");
  X.utils.book_append_sheet(wb, ws3, "Por colaborador");
  X.utils.book_append_sheet(wb, ws4, "Paradas");
  X.utils.book_append_sheet(wb, ws5, "Faltas");
  const nomeArq = `Controle_de_Atividades_${de.split("-").reverse().join("-")}${de !== ate ? "_a_" + ate.split("-").reverse().join("-") : ""}.xlsx`;
  X.writeFile(wb, nomeArq);
  toast("📥 Excel baixado");
}

// Lê a planilha no formato "Controle de Atividades" (aba com a coluna "Colaborador")
async function lerPlanilha(arquivo) {
  const X = await carregarXLSX();
  const wb = X.read(await arquivo.arrayBuffer(), { type: "array" });
  for (const nomeAba of wb.SheetNames) {
    const rows = X.utils.sheet_to_json(wb.Sheets[nomeAba], { header: 1, raw: true, defval: null });
    const hi = rows.findIndex(r => Array.isArray(r) && r.some(c => norm(c) === "COLABORADOR"));
    if (hi < 0) continue;
    const cab = rows[hi].map(norm);
    const col = (...nomes) => cab.findIndex(c => nomes.includes(c));
    const ix = {
      data: col("DATA"), colab: col("COLABORADOR", "COLABORADORES"), tipo: col("TIPO"), local: col("LOCAL"),
      ini: col("INICIO", "HORA INICIO"), fim: col("TERMINO", "FIM", "HORA FIM"),
      lista: col("QTD. ITENS DA LISTA", "QTD ITENS DA LISTA", "ITENS DA LISTA"), abast: col("ITENS ABASTECIDOS", "ABASTECIDOS"),
      obs: col("OBSERVACAO", "OBSERVACOES", "OBS"),
    };
    return converterLinhas(X, nomeAba, rows.slice(hi + 1), ix);
  }
  throw new Error("Não encontrei a coluna \"Colaborador\" em nenhuma aba da planilha.");
}
function converterLinhas(X, nomeAba, rows, ix) {
  const v = (r, i) => i >= 0 ? r[i] : null;
  const vazio = x => x == null || String(x).trim() === "" || /^-+$/.test(String(x).trim());
  const dataDe = x => {
    if (typeof x === "number") { const d = X.SSF.parse_date_code(x); return d ? `${d.y}-${pad(d.m)}-${pad(d.d)}` : null; }
    const s = String(x ?? "");
    let m = s.match(/(\d{1,2})\/(\d{1,2})\/(\d{2,4})/);
    if (m) { let y = +m[3]; if (y < 100) y += 2000; return `${y}-${pad(+m[2])}-${pad(+m[1])}`; }
    m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
    return m ? m[0] : null;
  };
  const minutosDe = x => {   // hora → minutos do dia (o Excel guarda hora como fração do dia)
    if (typeof x === "number") return Math.round((x - Math.floor(x)) * 1440) % 1440;
    const m = String(x ?? "").match(/(\d{1,2})\s*[:h]\s*(\d{2})/i);
    return m ? (+m[1]) * 60 + (+m[2]) : null;
  };
  const inteiro = x => { if (vazio(x)) return null; const n = Number(String(x).replace(",", ".")); return isNaN(n) || n < 0 ? null : Math.round(n); };
  const tipoDe = local => /^RUA\b/.test(norm(local)) ? "Rua" : /JIRAU/.test(norm(local)) ? "Jirau" : String(local).trim();
  const out = [], ignoradas = [];
  let ultimaData = null;
  rows.forEach((r, i) => {
    if (!Array.isArray(r) || r.every(vazio)) return;
    const data = dataDe(v(r, ix.data)) || ultimaData;
    if (data) ultimaData = data;
    const colab = v(r, ix.colab), local = v(r, ix.local);
    if (vazio(colab) || vazio(local) || !data) { ignoradas.push(i); return; }
    const nomes = String(colab).split(/\s+E\s+|,|\/|\+|;/).map(s => s.trim().replace(/\s+/g, " ").toUpperCase()).filter(s => s && !/^-+$/.test(s));
    if (!nomes.length) { ignoradas.push(i); return; }
    const ini = vazio(v(r, ix.ini)) ? null : minutosDe(v(r, ix.ini));
    const fim = vazio(v(r, ix.fim)) ? null : minutosDe(v(r, ix.fim));
    const localNome = String(local).trim().replace(/\s+/g, " ");
    out.push({
      data, nomes, localNome, tipo: vazio(v(r, ix.tipo)) ? tipoDe(localNome) : String(v(r, ix.tipo)).trim(),
      ini, fim, lista: inteiro(v(r, ix.lista)), abast: inteiro(v(r, ix.abast)),
      obs: vazio(v(r, ix.obs)) ? "" : String(v(r, ix.obs)).trim(),
    });
  });
  return { aba: nomeAba, linhas: out, ignoradas: ignoradas.length };
}

function subImportar(el) {
  el.innerHTML = `<div class="card"><h2>📤 Importar planilha de controle</h2>
    <p class="muted">Traga para o app as atividades das planilhas que vocês já usam (formato "Controle de Atividades": Data, Colaborador, Tipo, Local, Início, Término, Qtd. itens da lista, Itens abastecidos, Observação).</p>
    <ul class="muted">
      <li>Linhas com "-" ou sem colaborador são ignoradas.</li>
      <li>"ANDREI E PAULO" vira uma atividade com 2 colaboradores.</li>
      <li>Colaboradores e locais que não existem são cadastrados.</li>
      <li>Importar a mesma planilha de novo <b>não duplica</b>: as linhas iguais são atualizadas.</li>
      <li>Tempo, pendentes e itens/hora são recalculados pelo app (as fórmulas da planilha não são usadas).</li>
    </ul>
    <input type="file" id="iArq" accept=".xlsx,.xls,.xlsm,.csv">
  </div>`;
  $("#iArq").onchange = async e => {
    const arq = e.target.files[0]; if (!arq) return;
    ocupado("Lendo a planilha…");
    const r = await tentar(() => lerPlanilha(arq));
    e.target.value = "";
    if (r) previaImportacao(r);
  };
}
function previaImportacao(r) {
  if (!r.linhas.length) return toast("Não achei nenhuma atividade válida na planilha.", 4500);
  const colabPorNome = new Map(S.colaboradores.map(c => [norm(c.nome), c]));
  const localPorNome = new Map(S.locais.map(l => [norm(l.nome), l]));
  const novosColab = [...new Set(r.linhas.flatMap(l => l.nomes).filter(n => !colabPorNome.has(norm(n))))];
  const novosLocais = [...new Map(r.linhas.filter(l => !localPorNome.has(norm(l.localNome))).map(l => [norm(l.localNome), l])).values()];
  const datas = r.linhas.map(l => l.data).sort();
  const m = modal(`<h2>📤 Conferir importação</h2>
    <p>Aba <b>${esc(r.aba)}</b> · <b>${r.linhas.length}</b> atividade(s) de ${dataBR(datas[0])}${datas[0] !== datas.at(-1) ? ` a ${dataBR(datas.at(-1))}` : ""}.</p>
    ${r.ignoradas ? `<p class="muted">${r.ignoradas} linha(s) ignorada(s) (sem colaborador, sem local ou com "-").</p>` : ""}
    ${novosColab.length ? `<p><b>${novosColab.length} colaborador(es) novo(s):</b><br><span class="muted">${esc(novosColab.join(", "))}</span></p>` : ""}
    ${novosLocais.length ? `<p><b>${novosLocais.length} local(is) novo(s):</b><br><span class="muted">${esc(novosLocais.map(l => l.localNome).join(", "))}</span></p>` : ""}
    <div class="alert info">Confira os nomes: se a mesma pessoa aparece escrita de jeitos diferentes (ex.: "THIAGO C." e "THIAGO CATIZANO"), ela vira dois cadastros. Dá para corrigir depois em Cadastros.</div>
    <button class="btn" id="impOk">Importar ${r.linhas.length} atividade(s)</button>
    <button class="btn sec" data-fechar>Cancelar</button>`);
  $("#impOk", m.sheet).onclick = async () => {
    m.fechar();
    ocupado("Importando…");
    const ok = await tentar(async () => {
      const ops = [];   // [ref, dados]
      novosColab.forEach(nome => { const ref = doc(collection(db, "colaboradores")); colabPorNome.set(norm(nome), { id: ref.id, nome }); ops.push([ref, { nome, ativo: true, criadoEm: serverTimestamp() }]); });
      let ordem = Math.max(0, ...S.locais.map(l => l.ordem || 0));
      novosLocais.forEach(l => { const ref = doc(collection(db, "locais")); ordem += 10; localPorNome.set(norm(l.localNome), { id: ref.id, nome: l.localNome, tipo: l.tipo }); ops.push([ref, { nome: l.localNome, tipo: l.tipo, previstas: 1, ordem, ativo: true }]); });
      r.linhas.forEach(l => {
        const loc = localPorNome.get(norm(l.localNome));
        const pessoas = l.nomes.map(n => { const c = colabPorNome.get(norm(n)); return { id: c.id, nome: c.nome }; });
        const base = deISO(l.data);
        const ini = l.ini != null ? new Date(base.getTime() + l.ini * 60000) : null;
        let fim = l.fim != null ? new Date(base.getTime() + l.fim * 60000) : null;
        if (ini && fim && fim <= ini) fim = new Date(fim.getTime() + 864e5);
        // id fixo pelo conteúdo da linha: importar de novo atualiza em vez de duplicar
        const chave = [l.data, l.nomes.map(norm).sort().join("+"), norm(l.localNome), l.ini, l.fim, l.lista, l.abast].join("|");
        const ref = doc(db, "atividades", `imp_${hash(chave)}${hash(chave + "#")}`);
        ops.push([ref, {
          data: l.data, localId: loc.id, localNome: loc.nome, tipo: l.tipo || loc.tipo || "",
          pessoas, pessoaIds: pessoas.map(p => p.id),
          inicio: ini ? Timestamp.fromDate(ini) : null, fim: fim ? Timestamp.fromDate(fim) : null,
          itensLista: l.lista, itensAbastecidos: l.abast, obs: l.obs,
          origem: "planilha", criadoEm: serverTimestamp(), criadoPor: S.perfil.nome, atualizadoEm: serverTimestamp(), atualizadoPor: S.perfil.nome,
        }]);
      });
      for (let i = 0; i < ops.length; i += 400) {
        ocupado(`Importando… ${Math.min(i + 400, ops.length)} de ${ops.length}`);
        const b = writeBatch(db);
        ops.slice(i, i + 400).forEach(([ref, d]) => b.set(ref, d));
        await b.commit();
      }
      return true;
    }, "A importação parou no meio. Importe de novo: o que já entrou não duplica.");
    if (!ok) return;
    S.cache = null;
    toast(`✅ ${r.linhas.length} atividade(s) importada(s)`, 4500);
    const ultima = r.linhas.map(l => l.data).sort().at(-1);
    S.aba = "dia"; S.dia = ultima; S.atividades = []; ouvirDia(); go(telaPainel);
  };
}

iniciar();
