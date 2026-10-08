/*!
 * <listado-participantes> — listado embebible de participantes de un evento.
 * Lee una planilla (Google Sheets publicada como CSV, o cualquier CSV) y arma
 * tarjetas con foto circular, búsqueda, filtros desplegables y biografía desplegable.
 * Interfaz en español, inglés o portugués: cada página de idioma tiene su propia planilla
 * y su propio código con idioma="es", "en" o "pt".
 *
 * Uso mínimo:
 *   <listado-participantes src="URL_DE_LA_PLANILLA" idioma="es"></listado-participantes>
 *   <script src="participantes.js" defer></script>
 *
 * Atributos:
 *   src       URL del CSV. Acepta el link "Publicar en la web" (CSV) de Google Sheets
 *             o el link normal de la planilla (compartida "cualquiera con el link").
 *   idioma    es | en | pt. Si falta, usa el idioma de la página (<html lang>) o español.
 *   filtro    Columnas que arman los desplegables, separadas por coma.
 *             Defecto: "organizacion". Un filtro con menos de 2 opciones no se muestra.
 *   columnas  Máximo de columnas en pantallas anchas: 3, 4 (defecto) o 5.
 *   orden     "planilla" (defecto: columna "orden" o el orden de las filas) o "alfabetico".
 *   titulo    Texto opcional arriba del buscador.
 *   fuente    "sitio" para usar la tipografía de la página en lugar de Archivo.
 *
 * Encabezados de la planilla: se aceptan en español, inglés o portugués
 * (nombre/name/nome, organizacion/organization/organizacao, cargo/role, foto/photo, bio...).
 *
 * Estilo: variables CSS, por ejemplo
 *   listado-participantes { --participantes-acento: #E53357; }
 */
(() => {
  if (customElements.get('listado-participantes')) return;

  /* ---------- textos de la interfaz ---------- */
  const I18N = {
    es: {
      buscar: 'Nombre u organización',
      buscarAria: 'Buscar participantes',
      filtrosAria: 'Filtros',
      todos: { categoria: 'Todas las categorías', organizacion: 'Todas las organizaciones', pais: 'Todos los países', cargo: 'Todos los cargos', _: 'Todos' },
      etiqueta: { categoria: 'Categoría', organizacion: 'Organización', pais: 'País', cargo: 'Cargo' },
      total: (n) => `${n} participante${n === 1 ? '' : 's'}`,
      mostrando: (a, b) => `Mostrando ${a} de ${b}`,
      vacio: (q) => `Ningún participante coincide${q ? ` con «${q}»` : ''}.`,
      limpiar: 'Limpiar filtros',
      perfil: 'Ver perfil',
      bioDe: (n) => `Biografía de ${n}`,
      verBio: 'Ver biografía',
      cargando: 'Cargando participantes',
      error: 'No se pudo cargar el listado.',
      sinSrc: 'Falta el atributo src con la URL de la planilla.',
      noCsv: 'La URL devolvió una página web, no un CSV. Publicá la hoja como CSV (Archivo › Compartir › Publicar en la web).',
      httpError: (s) => `La planilla respondió con error ${s}.`,
    },
    en: {
      buscar: 'Name or organization',
      buscarAria: 'Search participants',
      filtrosAria: 'Filters',
      todos: { categoria: 'All categories', organizacion: 'All organizations', pais: 'All countries', cargo: 'All roles', _: 'All' },
      etiqueta: { categoria: 'Category', organizacion: 'Organization', pais: 'Country', cargo: 'Role' },
      total: (n) => `${n} participant${n === 1 ? '' : 's'}`,
      mostrando: (a, b) => `Showing ${a} of ${b}`,
      vacio: (q) => `No participants match${q ? ` “${q}”` : ' these filters'}.`,
      limpiar: 'Clear filters',
      perfil: 'View profile',
      bioDe: (n) => `Biography of ${n}`,
      verBio: 'Show biography',
      cargando: 'Loading participants',
      error: 'The list could not be loaded.',
      sinSrc: 'The src attribute with the spreadsheet URL is missing.',
      noCsv: 'The URL returned a web page, not a CSV. Publish the sheet as CSV (File › Share › Publish to web).',
      httpError: (s) => `The spreadsheet returned error ${s}.`,
    },
    pt: {
      buscar: 'Nome ou organização',
      buscarAria: 'Buscar participantes',
      filtrosAria: 'Filtros',
      todos: { categoria: 'Todas as categorias', organizacion: 'Todas as organizações', pais: 'Todos os países', cargo: 'Todos os cargos', _: 'Todos' },
      etiqueta: { categoria: 'Categoria', organizacion: 'Organização', pais: 'País', cargo: 'Cargo' },
      total: (n) => `${n} participante${n === 1 ? '' : 's'}`,
      mostrando: (a, b) => `Mostrando ${a} de ${b}`,
      vacio: (q) => `Nenhum participante corresponde${q ? ` a “${q}”` : ' a estes filtros'}.`,
      limpiar: 'Limpar filtros',
      perfil: 'Ver perfil',
      bioDe: (n) => `Biografia de ${n}`,
      verBio: 'Ver biografia',
      cargando: 'Carregando participantes',
      error: 'Não foi possível carregar a lista.',
      sinSrc: 'Falta o atributo src com a URL da planilha.',
      noCsv: 'A URL retornou uma página web, não um CSV. Publique a planilha como CSV (Arquivo › Compartilhar › Publicar na web).',
      httpError: (s) => `A planilha retornou o erro ${s}.`,
    },
  };

  /* ---------- utilidades ---------- */
  const norm = (s) =>
    (s ?? '').toString().normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().trim();

  const esc = (s) =>
    (s ?? '').toString().replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));

  // CSV robusto: comillas, saltos de línea dentro de celdas, separador "," o ";" (Excel en español).
  function parseCSV(text) {
    text = text.replace(/^﻿/, '');
    const nl = text.indexOf('\n');
    const firstLine = nl === -1 ? text : text.slice(0, nl);
    const sep = (firstLine.match(/;/g) || []).length > (firstLine.match(/,/g) || []).length ? ';' : ',';
    const rows = [];
    let row = [], f = '', q = false;
    for (let i = 0; i < text.length; i++) {
      const c = text[i];
      if (q) {
        if (c === '"') { if (text[i + 1] === '"') { f += '"'; i++; } else q = false; }
        else f += c;
      } else if (c === '"') q = true;
      else if (c === sep) { row.push(f); f = ''; }
      else if (c === '\n') { row.push(f); rows.push(row); row = []; f = ''; }
      else if (c !== '\r') f += c;
    }
    if (f !== '' || row.length) { row.push(f); rows.push(row); }
    return rows.filter((r) => r.some((x) => x.trim() !== ''));
  }

  // Encabezados aceptados para cada campo (sin tildes, en minúscula).
  const ALIAS = {
    nombre: ['nombre', 'nombre y apellido', 'participante', 'name', 'full name', 'nome', 'nome completo'],
    organizacion: ['organizacion', 'org', 'medio', 'institucion', 'empresa', 'organization', 'organisation', 'organizacao', 'instituicao', 'veiculo'],
    cargo: ['cargo', 'rol', 'puesto', 'role', 'position', 'title', 'funcao'],
    foto: ['foto', 'imagen', 'foto url', 'url foto', 'photo', 'image', 'picture', 'imagem'],
    bio: ['bio', 'biografia', 'descripcion', 'perfil', 'biography', 'about', 'descricao'],
    categoria: ['categoria', 'tipo', 'grupo', 'category', 'type', 'group'],
    orden: ['orden', 'order', 'posicion', 'ordem', 'posicao'],
    link: ['link', 'url', 'web', 'redes', 'perfil url', 'website', 'profile'],
    pais: ['pais', 'country'],
  };

  // Devuelve la clave canónica (nombre, cargo...) de un nombre de columna, o el nombre normalizado.
  const canon = (name) => {
    const n = norm(name);
    for (const [k, names] of Object.entries(ALIAS)) if (names.includes(n)) return k;
    return n;
  };

  function rowsToPeople(rows, filtros) {
    if (!rows.length) return [];
    const head = rows[0].map(norm);
    const find = (names) => head.findIndex((h) => names.includes(h));
    const col = {};
    for (const [k, names] of Object.entries(ALIAS)) col[k] = find(names);
    // Columnas de filtro que no son campos conocidos (p. ej. "dia", "eje")
    for (const f of filtros) if (!(f in col)) col[f] = find([f]);
    const get = (r, k) => (col[k] >= 0 ? (r[col[k]] ?? '').trim() : '');
    return rows.slice(1)
      .map((r, i) => ({
        nombre: get(r, 'nombre'),
        organizacion: get(r, 'organizacion'),
        cargo: get(r, 'cargo'),
        foto: get(r, 'foto'),
        bio: get(r, 'bio'),
        link: get(r, 'link'),
        pais: get(r, 'pais'),
        f: Object.fromEntries(filtros.map((k) => [k, get(r, k)])),
        orden: parseFloat(get(r, 'orden').replace(',', '.')),
        _fila: i,
      }))
      .filter((p) => p.nombre);
  }

  // Convierte links de Google Sheets a su versión CSV.
  function toCsvUrl(u) {
    u = (u || '').trim();
    if (/docs\.google\.com\/spreadsheets\/d\/e\//.test(u)) {
      if (/output=csv/.test(u)) return u;
      const gid = (u.match(/[#&?]gid=(\d+)/) || [])[1];
      return u.replace(/\/pub(html)?.*$/, '/pub?output=csv' + (gid ? '&gid=' + gid : ''));
    }
    const m = u.match(/docs\.google\.com\/spreadsheets\/d\/([\w-]+)/);
    if (m && !/format=csv/.test(u)) {
      const gid = (u.match(/[#&?]gid=(\d+)/) || [])[1] || '0';
      return `https://docs.google.com/spreadsheets/d/${m[1]}/export?format=csv&gid=${gid}`;
    }
    return u;
  }

  // Convierte links de Google Drive en URLs de imagen directas.
  function driveId(u) {
    const m = (u || '').match(/drive\.google\.com\/(?:file\/d\/|open\?id=|uc\?(?:[^#]*&)?id=|thumbnail\?(?:[^#]*&)?id=)([\w-]+)/);
    return m ? m[1] : null;
  }
  function photoUrl(u) {
    u = (u || '').trim();
    const id = driveId(u);
    return id ? `https://lh3.googleusercontent.com/d/${id}=w300` : u;
  }

  function initials(name) {
    const parts = name.split(/\s+/).filter(Boolean);
    return ((parts[0]?.[0] || '') + (parts.length > 1 ? parts[parts.length - 1][0] : '')).toUpperCase();
  }

  function cargarFuente() {
    if (document.querySelector('link[data-participantes-fuente]')) return;
    const l = document.createElement('link');
    l.rel = 'stylesheet';
    l.href = 'https://fonts.googleapis.com/css2?family=Archivo:wght@400;500;600;700&display=swap';
    l.setAttribute('data-participantes-fuente', '');
    document.head.appendChild(l);
  }

  const ICON_SEARCH = '<svg viewBox="0 0 20 20" aria-hidden="true"><circle cx="8.5" cy="8.5" r="5.5" fill="none" stroke="currentColor" stroke-width="1.8"/><path d="M13 13l4.5 4.5" stroke="currentColor" stroke-width="1.8" stroke-linecap="round"/></svg>';
  const ICON_CHEV = '<svg viewBox="0 0 16 16" aria-hidden="true"><path d="M4 6l4 4 4-4" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"/></svg>';

  const CSS = `
:host{
  /* Identidad LatamChequea: rojo #E53357, azul oscuro #384051, tipografía Archivo */
  --lp-a: var(--participantes-acento, #E53357);
  --lp-at: var(--participantes-acento-texto, #D02A4C);
  --lp-card: var(--participantes-card, #ffffff);
  --lp-fg: var(--participantes-texto, #384051);
  --lp-muted: var(--participantes-texto-suave, #6B7385);
  --lp-line: var(--participantes-borde, #E2E4EA);
  --lp-soft: var(--participantes-suave, #F2F3F6);
  --lp-ini: var(--participantes-iniciales, #FDE8EC);
  --lp-shadow: var(--participantes-sombra, 0 1px 2px rgba(53,64,82,.06), 0 3px 10px rgba(53,64,82,.07));
  --lp-radius: var(--participantes-radio, 6px);
  --lp-foto: 64px;
  display:block; container-type:inline-size;
  font-family: var(--participantes-fuente, 'Archivo', system-ui, -apple-system, 'Segoe UI', Roboto, sans-serif);
  color: var(--lp-fg); line-height:1.4;
}
:host([fuente="sitio"]){font-family:var(--participantes-fuente, inherit)}
*,*::before,*::after{box-sizing:border-box}
[hidden]{display:none!important}
.titulo{margin:0 0 12px;font-size:1.375rem;font-weight:700;letter-spacing:-.01em}
.barra{display:flex;flex-wrap:wrap;gap:10px 12px;align-items:center;margin-bottom:14px}
.buscar{position:relative;flex:1 1 260px;min-width:0;max-width:420px}
.buscar svg{position:absolute;left:12px;top:50%;width:18px;height:18px;transform:translateY(-50%);color:var(--lp-muted);pointer-events:none}
.campo{width:100%;font:inherit;font-size:1rem;color:var(--lp-fg);background:var(--lp-card);
  border:1px solid var(--lp-line);border-radius:var(--lp-radius);min-height:44px}
.buscar .campo{padding:10px 12px 10px 38px}
.campo::placeholder{color:var(--lp-muted)}
.campo:focus-visible{outline:2px solid var(--lp-a);outline-offset:1px;border-color:transparent}
.sel{position:relative;flex:0 1 250px;min-width:0}
.sel .campo{appearance:none;-webkit-appearance:none;padding:10px 38px 10px 12px;cursor:pointer;
  text-overflow:ellipsis;white-space:nowrap;overflow:hidden;font-size:.9375rem}
.sel svg{position:absolute;right:12px;top:50%;width:16px;height:16px;transform:translateY(-50%);color:var(--lp-muted);pointer-events:none}
.sel.activo .campo{border-color:var(--lp-a);box-shadow:inset 0 0 0 1px var(--lp-a);font-weight:600}
.sel.activo svg{color:var(--lp-at)}
@container (max-width:599px){.buscar{flex-basis:100%;max-width:none}.sel{flex:1 1 calc(50% - 6px)}}
@container (max-width:380px){.sel{flex-basis:100%}}
.estado{display:flex;flex-wrap:wrap;gap:4px 12px;align-items:baseline;margin:0 0 12px;font-size:.8125rem;color:var(--lp-muted);font-variant-numeric:tabular-nums}
.limpiar{font:inherit;font-size:.8125rem;font-weight:600;color:var(--lp-at);background:none;border:0;cursor:pointer;text-decoration:underline;text-underline-offset:2px;padding:2px}
.cab:focus-visible,.limpiar:focus-visible{outline:2px solid var(--lp-a);outline-offset:2px}
.grid{list-style:none;margin:0;padding:0;display:grid;gap:14px;grid-template-columns:minmax(0,1fr);align-items:start}
@container (min-width:500px){.grid{grid-template-columns:repeat(2,minmax(0,1fr))}}
@container (min-width:740px){.grid{grid-template-columns:repeat(3,minmax(0,1fr))}}
@container (min-width:960px){:host(:not([columnas="3"])) .grid{grid-template-columns:repeat(4,minmax(0,1fr))}}
@container (min-width:1200px){:host([columnas="5"]) .grid{grid-template-columns:repeat(5,minmax(0,1fr))}}
@container (max-width:499px){:host{--lp-foto:56px}}
.card{background:var(--lp-card);border:1px solid var(--lp-line);border-radius:var(--lp-radius);box-shadow:var(--lp-shadow);
  overflow:hidden;transition:box-shadow .2s, border-color .2s}
.card.abierta{border-color:color-mix(in srgb, var(--lp-a) 45%, var(--lp-line))}
.cab{display:flex;gap:14px;align-items:center;width:100%;min-height:calc(var(--lp-foto) + 28px);padding:14px;
  background:none;border:0;font:inherit;color:inherit;text-align:left}
button.cab{cursor:pointer}
button.cab:hover .nombre{color:var(--lp-at)}
.foto{flex:none;width:var(--lp-foto);height:var(--lp-foto);border-radius:50%;object-fit:cover;background:var(--lp-soft);display:block}
.ini{flex:none;width:var(--lp-foto);height:var(--lp-foto);border-radius:50%;display:grid;place-items:center;
  background:var(--lp-ini);color:var(--lp-at);font-weight:700;font-size:1.25rem;letter-spacing:.02em}
.txt{min-width:0;flex:1;display:flex;flex-direction:column;gap:3px}
.nombre{font-size:1.0625rem;font-weight:700;color:var(--lp-fg);letter-spacing:-.005em;line-height:1.25;overflow-wrap:anywhere;transition:color .15s}
.cargo{font-size:.8125rem;color:var(--lp-muted);line-height:1.3}
.org{font-size:.8125rem;font-weight:600;color:var(--lp-at);line-height:1.3}
.chev{flex:none;width:28px;height:28px;display:grid;place-items:center;border-radius:50%;color:var(--lp-muted);
  transition:transform .25s, background .15s}
.chev svg{width:16px;height:16px}
button.cab:hover .chev{background:var(--lp-soft)}
.abierta .chev{transform:rotate(180deg);color:var(--lp-at)}
.bio{display:grid;grid-template-rows:0fr;transition:grid-template-rows .25s ease}
.abierta .bio{grid-template-rows:1fr}
.bio>div{overflow:hidden}
.bio-in{border-top:1px solid var(--lp-line);margin:0 14px;padding:12px 0 14px;font-size:.9rem;color:var(--lp-fg)}
.bio-in p{margin:0 0 .6em}
.bio-in p:last-child{margin:0}
.bio-in a{color:var(--lp-at);font-weight:600;font-size:.8125rem;text-decoration:none}
.bio-in a:hover{text-decoration:underline}
.vacio{padding:32px 16px;text-align:center;color:var(--lp-muted);border:1px dashed var(--lp-line);border-radius:var(--lp-radius)}
.error{padding:16px;border-radius:var(--lp-radius);background:var(--lp-soft);color:var(--lp-fg);font-size:.9rem}
.skel .foto,.skel .linea{background:var(--lp-soft);animation:pulso 1.2s ease-in-out infinite}
.skel .linea{height:12px;border-radius:4px}
@keyframes pulso{50%{opacity:.5}}
@media (prefers-reduced-motion:reduce){*{transition:none!important;animation:none!important}}
`;

  class ListadoParticipantes extends HTMLElement {
    static get observedAttributes() { return ['src', 'filtro', 'titulo', 'orden', 'idioma']; }

    constructor() {
      super();
      this.root = this.attachShadow({ mode: 'open' });
      this.people = [];
      this.q = '';
      this.sel = {};
      this.open = new Set();
      this.uid = 'lp' + Math.random().toString(36).slice(2, 7);
    }

    get lang() {
      const l = (this.getAttribute('idioma') || document.documentElement.lang || 'es').slice(0, 2).toLowerCase();
      return I18N[l] ? l : 'es';
    }
    get t() { return I18N[this.lang]; }
    get filtros() {
      return (this.getAttribute('filtro') ?? 'organizacion')
        .split(',').map((s) => s.trim()).filter(Boolean).map(canon);
    }

    connectedCallback() {
      if (this._init) return;
      this._init = true;
      if (this.getAttribute('fuente') !== 'sitio') cargarFuente();
      this.root.innerHTML = `<style>${CSS}</style><div class="w"></div>`;
      this.w = this.root.querySelector('.w');
      this.load();
    }

    attributeChangedCallback(name, oldV, newV) {
      if (this._init && oldV !== newV) this.load();
    }

    /** Carga datos desde JS: el.setData([{nombre, organizacion, bio, bio_en, ...}]) */
    setData(list) {
      const cols = [...new Set(list.flatMap((o) => Object.keys(o)))];
      this._rows = [cols, ...list.map((o) => cols.map((c) => (o[c] ?? '').toString()))];
      if (this._init) this.load();
    }

    error(msg) {
      this.w.innerHTML = `<div class="error" role="alert"><strong>${esc(this.t.error)}</strong><br>${esc(msg)}</div>`;
    }

    async load() {
      const t = this.t;
      this.renderSkeleton();
      try {
        let rows;
        if (this._rows) rows = this._rows;
        else {
          const inline = this.querySelector('script[type="text/csv"]');
          if (inline) rows = parseCSV(inline.textContent.trim());
          else {
            const src = this.getAttribute('src');
            if (!src) {
              // Los datos pueden llegar después por setData(); esperamos antes de mostrar el error.
              clearTimeout(this._espera);
              this._espera = setTimeout(() => {
                if (!this._rows && !this.getAttribute('src')) this.error(t.sinSrc);
              }, 2500);
              return;
            }
            const res = await fetch(toCsvUrl(src), { cache: 'no-cache' });
            if (!res.ok) throw new Error(t.httpError(res.status));
            const text = await res.text();
            if (/^\s*<(!doctype|html)/i.test(text)) throw new Error(t.noCsv);
            rows = parseCSV(text);
          }
        }
        const filtros = this.filtros;
        const people = rowsToPeople(rows, filtros);
        const locale = this.lang === 'en' ? 'en' : this.lang === 'pt' ? 'pt' : 'es';
        if ((this.getAttribute('orden') || '') === 'alfabetico') {
          people.sort((a, b) => a.nombre.localeCompare(b.nombre, locale));
        } else {
          people.sort((a, b) => {
            const ao = isNaN(a.orden) ? Infinity : a.orden, bo = isNaN(b.orden) ? Infinity : b.orden;
            return ao - bo || a._fila - b._fila;
          });
        }
        people.forEach((p) => {
          p._busca = norm([p.nombre, p.organizacion, p.cargo, p.bio, p.pais, ...Object.values(p.f)].join(' '));
          p._id = this.uid + '-' + p._fila;
        });
        this.people = people;
        // Opciones de cada filtro, ordenadas alfabéticamente y con cantidad
        this.opciones = filtros
          .map((k) => {
            const cuenta = new Map();
            people.forEach((p) => p.f[k] && cuenta.set(p.f[k], (cuenta.get(p.f[k]) || 0) + 1));
            const vals = [...cuenta.keys()].sort((a, b) => a.localeCompare(b, locale, { sensitivity: 'base' }));
            return { k, vals, cuenta };
          })
          .filter((o) => o.vals.length > 1);
        this.opciones.forEach((o) => { if (this.sel[o.k] && !o.cuenta.has(this.sel[o.k])) this.sel[o.k] = ''; });
        this.renderShell();
        this.renderList();
        this.dispatchEvent(new CustomEvent('participantes:cargado', { detail: { total: people.length }, bubbles: true }));
      } catch (err) {
        this.error(err.message);
        console.error('[listado-participantes]', err);
      }
    }

    renderSkeleton() {
      const card = `<li class="card skel"><div class="cab"><span class="foto"></span><span class="txt"><span class="linea" style="width:70%"></span><span class="linea" style="width:45%"></span></span></div></li>`;
      this.w.innerHTML = `<ul class="grid" aria-busy="true" aria-label="${esc(this.t.cargando)}">${card.repeat(8)}</ul>`;
    }

    renderShell() {
      const t = this.t;
      const titulo = this.getAttribute('titulo');
      const selects = this.opciones.map((o) => {
        const todos = t.todos[o.k] || t.todos._;
        const etiqueta = t.etiqueta[o.k] || o.k;
        const v = this.sel[o.k] || '';
        return `<label class="sel${v ? ' activo' : ''}">
            <select class="campo" id="${this.uid}-f-${esc(o.k)}" data-k="${esc(o.k)}" aria-label="${esc(etiqueta)}">
              <option value="">${esc(todos)}</option>
              ${o.vals.map((x) => `<option value="${esc(x)}"${x === v ? ' selected' : ''}>${esc(x)} (${o.cuenta.get(x)})</option>`).join('')}
            </select>${ICON_CHEV}
          </label>`;
      }).join('');
      this.w.innerHTML = `
        ${titulo ? `<h2 class="titulo">${esc(titulo)}</h2>` : ''}
        <div class="barra" role="search" aria-label="${esc(t.filtrosAria)}">
          <label class="buscar">
            ${ICON_SEARCH}
            <input class="campo" type="search" id="${this.uid}-q" placeholder="${esc(t.buscar)}" aria-label="${esc(t.buscarAria)}" autocomplete="off" value="${esc(this.q)}">
          </label>
          ${selects}
        </div>
        <p class="estado"><span class="cuenta" aria-live="polite"></span><button type="button" class="limpiar" hidden>${esc(t.limpiar)}</button></p>
        <ul class="grid"></ul>
        <div class="vacio" hidden></div>`;
      this.grid = this.w.querySelector('.grid');
      this.cuenta = this.w.querySelector('.cuenta');
      this.vacio = this.w.querySelector('.vacio');
      this.btnLimpiar = this.w.querySelector('.limpiar');
      this.input = this.w.querySelector('input');

      let timer;
      this.input.addEventListener('input', () => {
        clearTimeout(timer);
        timer = setTimeout(() => { this.q = this.input.value; this.renderList(); }, 120);
      });
      this.w.querySelectorAll('select').forEach((s) =>
        s.addEventListener('change', () => {
          this.sel[s.dataset.k] = s.value;
          s.parentElement.classList.toggle('activo', !!s.value);
          this.renderList();
        })
      );
      this.grid.addEventListener('click', (e) => {
        const btn = e.target.closest('button.cab');
        if (!btn) return;
        const card = btn.closest('.card');
        const abierta = !card.classList.contains('abierta');
        card.classList.toggle('abierta', abierta);
        btn.setAttribute('aria-expanded', abierta);
        abierta ? this.open.add(card.dataset.id) : this.open.delete(card.dataset.id);
      });
      const limpiar = () => {
        this.q = ''; this.sel = {};
        this.input.value = '';
        this.w.querySelectorAll('select').forEach((s) => { s.value = ''; s.parentElement.classList.remove('activo'); });
        this.renderList();
        this.input.focus();
      };
      this.btnLimpiar.addEventListener('click', limpiar);
      this.vacio.addEventListener('click', (e) => e.target.closest('.limpiar') && limpiar());
    }

    card(p) {
      const t = this.t;
      const foto = photoUrl(p.foto);
      const avatar = foto
        ? `<img class="foto" src="${esc(foto)}" alt="" loading="lazy" decoding="async" referrerpolicy="no-referrer" data-ini="${esc(initials(p.nombre))}">`
        : `<span class="ini" aria-hidden="true">${esc(initials(p.nombre))}</span>`;
      const texto = `<span class="txt">
          <span class="nombre">${esc(p.nombre)}</span>
          ${p.cargo ? `<span class="cargo">${esc(p.cargo)}</span>` : ''}
          ${p.organizacion || p.pais ? `<span class="org">${esc([p.organizacion, p.pais].filter(Boolean).join(' · '))}</span>` : ''}
        </span>`;
      const abierta = this.open.has(p._id);
      if (!p.bio && !p.link) return `<li class="card" data-id="${p._id}"><div class="cab">${avatar}${texto}</div></li>`;
      const parrafos = p.bio.split(/\n+/).filter(Boolean).map((x) => `<p>${esc(x)}</p>`).join('');
      const link = /^https?:\/\//.test(p.link) ? `<p><a href="${esc(p.link)}" target="_blank" rel="noopener">${esc(t.perfil)} ↗</a></p>` : '';
      return `<li class="card${abierta ? ' abierta' : ''}" data-id="${p._id}">
        <button type="button" class="cab" aria-expanded="${abierta}" aria-controls="${p._id}-bio" title="${esc(t.verBio)}">
          ${avatar}${texto}<span class="chev">${ICON_CHEV}</span>
        </button>
        <div class="bio" id="${p._id}-bio" role="region" aria-label="${esc(t.bioDe(p.nombre))}"><div><div class="bio-in">${parrafos}${link}</div></div></div>
      </li>`;
    }

    renderList() {
      const t = this.t;
      const terms = norm(this.q).split(/\s+/).filter(Boolean);
      const activos = Object.entries(this.sel).filter(([, v]) => v);
      const list = this.people.filter((p) =>
        activos.every(([k, v]) => p.f[k] === v) && terms.every((x) => p._busca.includes(x))
      );
      this.grid.innerHTML = list.map((p) => this.card(p)).join('');
      this.grid.querySelectorAll('img.foto').forEach((img) =>
        img.addEventListener('error', () => {
          // Si falla la imagen de Drive, probamos con la miniatura; si no, iniciales.
          const id = driveId(img.src) || (img.src.match(/googleusercontent\.com\/d\/([\w-]+)/) || [])[1];
          if (id && !img.dataset.retry) {
            img.dataset.retry = '1';
            img.src = `https://drive.google.com/thumbnail?id=${id}&sz=w300`;
            return;
          }
          const s = document.createElement('span');
          s.className = 'ini'; s.setAttribute('aria-hidden', 'true'); s.textContent = img.dataset.ini;
          img.replaceWith(s);
        })
      );
      const total = this.people.length;
      const filtrado = list.length !== total || activos.length > 0 || terms.length > 0;
      this.cuenta.textContent = list.length === total ? t.total(total) : t.mostrando(list.length, total);
      this.btnLimpiar.hidden = !filtrado || !list.length;
      this.vacio.hidden = list.length > 0;
      if (!list.length) {
        this.vacio.innerHTML = `${esc(t.vacio(this.q.trim()))} <button type="button" class="limpiar">${esc(t.limpiar)}</button>`;
      }
    }
  }

  customElements.define('listado-participantes', ListadoParticipantes);
})();
