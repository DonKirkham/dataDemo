// ABOUTME: Builds a space-themed PowerPoint deck from the content in slides.md.
// ABOUTME: Run with `node presentation/build-deck.mjs` from the repo root.

import PptxGenJS from 'pptxgenjs';
import { fileURLToPath } from 'url';
import { dirname, join } from 'path';

const __dirname = dirname(fileURLToPath(import.meta.url));
const OUT_PATH = join(__dirname, 'data-access-spfx.pptx');

// ---------- THEME (deep space) ----------
const C = {
  bg:        '0A0E27',  // deep space navy
  bgPanel:   '141937',  // slightly lighter for code panels
  bgPanelHi: '1B2348',  // panel hover-ish, used for table headers
  ink:       'F0F4FF',  // off-white body
  inkDim:    'A4ADCE',  // muted body for secondary text
  cyan:      '00E5FF',  // hero accent (nebula cyan)
  violet:    'B580FF',  // secondary accent (nebula violet)
  amber:     'FFB347',  // for warnings / "missing" cells
  good:      '7EE787',  // for "free upgrades" / success cells
  bad:       'FF6B6B',  // for ❌ in pivot table
  rule:      '2A3258',  // hairline rule
};

const FONT = {
  body: 'Calibri',
  mono: 'Consolas',
};

// ---------- SLIDE BUILDER ----------
const pptx = new PptxGenJS();
pptx.layout = 'LAYOUT_WIDE'; // 13.333 x 7.5
pptx.title = 'Reading and Writing SharePoint Data from SPFx';
pptx.author = 'Don Kirkham';
pptx.company = 'Solution Foundry';

// ---------- MASTERS ----------
pptx.defineSlideMaster({
  title: 'SPACE_BASE',
  background: { color: C.bg },
  objects: [
    // Top-left tiny "constellation" mark
    { rect: { x: 0.3, y: 0.3, w: 0.08, h: 0.08, fill: { color: C.cyan }, line: { color: C.cyan } } },
    { rect: { x: 0.45, y: 0.42, w: 0.05, h: 0.05, fill: { color: C.violet }, line: { color: C.violet } } },
    { rect: { x: 0.55, y: 0.32, w: 0.04, h: 0.04, fill: { color: C.cyan }, line: { color: C.cyan } } },
    // Bottom hairline rule
    { line: { x: 0.5, y: 7.05, w: 12.33, h: 0, line: { color: C.rule, width: 0.75 } } },
    // Footer
    {
      text: {
        text: 'Don Kirkham · Solution Foundry · Reading and Writing SharePoint Data from SPFx',
        options: { x: 0.5, y: 7.1, w: 10.5, h: 0.3, fontFace: FONT.body, fontSize: 9, color: C.inkDim },
      },
    },
    // Page number (right of footer)
    {
      text: {
        text: 'PAGE_NUM',
        options: { x: 11.5, y: 7.1, w: 1.3, h: 0.3, fontFace: FONT.body, fontSize: 9, color: C.cyan, align: 'right' },
      },
    },
  ],
  slideNumber: { x: 11.5, y: 7.1, w: 1.3, h: 0.3, fontFace: FONT.body, fontSize: 9, color: C.cyan, align: 'right' },
});

pptx.defineSlideMaster({
  title: 'SPACE_SECTION',
  background: { color: C.bg },
});

// ---------- HELPERS ----------
function addTitle(slide, text, opts = {}) {
  slide.addText(text, {
    x: 0.5,
    y: 0.5,
    w: 12.33,
    h: 0.7,
    fontFace: FONT.body,
    fontSize: 30,
    bold: true,
    color: C.cyan,
    ...opts,
  });
  // Accent bar under title
  slide.addShape('rect', {
    x: 0.5,
    y: 1.2,
    w: 0.6,
    h: 0.05,
    fill: { color: C.violet },
    line: { color: C.violet, width: 0 },
  });
}

function addSubtitle(slide, text, opts = {}) {
  slide.addText(text, {
    x: 0.5,
    y: 1.32,
    w: 12.33,
    h: 0.4,
    fontFace: FONT.body,
    fontSize: 16,
    color: C.inkDim,
    italic: true,
    ...opts,
  });
}

function addBody(slide, runs, x = 0.5, y = 1.9, w = 12.33, h = 5.0, fontSize = 18) {
  slide.addText(runs, {
    x, y, w, h,
    fontFace: FONT.body,
    fontSize,
    color: C.ink,
    valign: 'top',
    paraSpaceAfter: 6,
  });
}

function bullet(text, opts = {}) {
  return { text, options: { bullet: { code: '25CF' }, color: C.ink, ...opts } };
}

function plain(text, opts = {}) {
  return { text, options: { color: C.ink, ...opts } };
}

function addCode(slide, code, x, y, w, h, fontSize = 13) {
  // Dark panel under the code
  slide.addShape('roundRect', {
    x, y, w, h,
    fill: { color: C.bgPanel },
    line: { color: C.rule, width: 0.75 },
    rectRadius: 0.08,
  });
  slide.addText(code, {
    x: x + 0.18, y: y + 0.12, w: w - 0.36, h: h - 0.24,
    fontFace: FONT.mono,
    fontSize,
    color: C.ink,
    valign: 'top',
    paraSpaceAfter: 0,
  });
}

function addNotes(slide, notes) {
  if (notes && notes.length) slide.addNotes(notes);
}

// ---------- SLIDES ----------

// SLIDE 1 — TITLE PLACEHOLDER (Don will replace)
{
  const s = pptx.addSlide({ masterName: 'SPACE_SECTION' });
  s.addText('TITLE SLIDE', {
    x: 0.5, y: 2.6, w: 12.33, h: 1.2,
    fontFace: FONT.body, fontSize: 56, bold: true, color: C.cyan, align: 'center',
  });
  s.addText('(replace with your own)', {
    x: 0.5, y: 3.9, w: 12.33, h: 0.6,
    fontFace: FONT.body, fontSize: 20, color: C.inkDim, italic: true, align: 'center',
  });
  s.addText('Don Kirkham · Solution Foundry', {
    x: 0.5, y: 4.6, w: 12.33, h: 0.5,
    fontFace: FONT.body, fontSize: 18, color: C.violet, align: 'center',
  });
}

// SLIDE 2 — ABOUT-ME PLACEHOLDER (Don will replace)
{
  const s = pptx.addSlide({ masterName: 'SPACE_SECTION' });
  s.addText('ABOUT ME', {
    x: 0.5, y: 2.6, w: 12.33, h: 1.2,
    fontFace: FONT.body, fontSize: 56, bold: true, color: C.cyan, align: 'center',
  });
  s.addText('(replace with your own)', {
    x: 0.5, y: 3.9, w: 12.33, h: 0.6,
    fontFace: FONT.body, fontSize: 20, color: C.inkDim, italic: true, align: 'center',
  });
}

// SLIDE 3 — Why this talk
{
  const s = pptx.addSlide({ masterName: 'SPACE_BASE' });
  addTitle(s, 'Why this talk');
  addSubtitle(s, 'You have four ways to get data into an SPFx web part');
  addBody(s, [
    bullet('SP REST via SPHttpClient'),
    bullet('MS Graph via MSGraphClientV3'),
    bullet('Anonymous external APIs via HttpClient'),
    bullet('PnPjs (@pnp/sp, @pnp/graph, @pnp/queryable) — wraps the first three'),
    plain(' ', { bullet: false }),
    plain('Most tutorials pick one and move on. Today we compare them on the same problem.', { bullet: false, color: C.violet, italic: true, fontSize: 18 }),
  ]);
  addNotes(s,
    `The "same problem" is a Speaking Events list — Title, Speaker, Session, SessionDate, SessionType, EventLogo, EventSite, SessionLink. Same CRUD operations against the same list. Six implementations.`
  );
}

// SLIDE 4 — The roadmap
{
  const s = pptx.addSlide({ masterName: 'SPACE_BASE' });
  addTitle(s, 'The roadmap');
  // Three columns for the three passes
  const colY = 1.9;
  const colH = 4.6;
  const colW = 4.0;
  const gutter = 0.13;
  const x0 = 0.5;
  const cols = [
    {
      x: x0, w: colW,
      header: 'PASS 1 — SPFx native',
      time: '≈20 min',
      bullets: ['SP REST against the list', 'MS Graph against the same list (and Graph Explorer)', 'Anonymous calls to a public API'],
      accent: C.cyan,
    },
    {
      x: x0 + colW + gutter, w: colW,
      header: 'PIVOT',
      time: '≈2 min',
      bullets: ['What\'s missing?'],
      accent: C.amber,
    },
    {
      x: x0 + (colW + gutter) * 2, w: colW,
      header: 'PASS 2 — PnPjs',
      time: '≈20 min',
      bullets: ['Same SP operations, fluent', 'Same Graph operations, fluent', 'The free upgrades: logging, caching, batching'],
      accent: C.violet,
    },
  ];
  for (const c of cols) {
    s.addShape('roundRect', { x: c.x, y: colY, w: c.w, h: colH, fill: { color: C.bgPanel }, line: { color: c.accent, width: 1.5 }, rectRadius: 0.1 });
    s.addText(c.header, { x: c.x + 0.2, y: colY + 0.18, w: c.w - 0.4, h: 0.5, fontFace: FONT.body, fontSize: 16, bold: true, color: c.accent });
    s.addText(c.time, { x: c.x + 0.2, y: colY + 0.7, w: c.w - 0.4, h: 0.35, fontFace: FONT.body, fontSize: 11, color: C.inkDim, italic: true });
    s.addText(c.bullets.map((b) => bullet(b, { fontSize: 14 })),
      { x: c.x + 0.2, y: colY + 1.1, w: c.w - 0.4, h: colH - 1.3, fontFace: FONT.body, fontSize: 14, color: C.ink, valign: 'top', paraSpaceAfter: 6 });
  }
  s.addText('Wrap (≈8 min) — when to still use SPFx-native, decision cheat sheet, Q&A', {
    x: 0.5, y: colY + colH + 0.15, w: 12.33, h: 0.4, fontFace: FONT.body, fontSize: 13, color: C.inkDim, italic: true, align: 'center',
  });
}

// SECTION DIVIDER — PASS 1
{
  const s = pptx.addSlide({ masterName: 'SPACE_SECTION' });
  s.addText('PASS 1', { x: 0.5, y: 2.5, w: 12.33, h: 0.8, fontFace: FONT.body, fontSize: 18, color: C.violet, align: 'center', charSpacing: 12 });
  s.addText('The SPFx Way', { x: 0.5, y: 3.3, w: 12.33, h: 1.4, fontFace: FONT.body, fontSize: 64, bold: true, color: C.cyan, align: 'center' });
  s.addShape('rect', { x: 6.17, y: 4.85, w: 1.0, h: 0.05, fill: { color: C.violet }, line: { color: C.violet, width: 0 } });
}

// SLIDE — SPHttpClient: the mechanics
{
  const s = pptx.addSlide({ masterName: 'SPACE_BASE' });
  addTitle(s, 'SPHttpClient: the mechanics');
  addSubtitle(s, 'A fetch wrapper from @microsoft/sp-http, already authenticated to your site');

  // Two columns: free vs not-free
  s.addText('What you get for free', { x: 0.5, y: 1.9, w: 6.0, h: 0.4, fontFace: FONT.body, fontSize: 16, bold: true, color: C.good });
  s.addText([
    bullet('Auth headers (cookies + bearer)', { fontSize: 16 }),
    bullet('Form digest on writes', { fontSize: 16 }),
    bullet('Sensible defaults', { fontSize: 16 }),
    bullet('  Accept: application/json;odata=nometadata', { fontSize: 13, color: C.inkDim, fontFace: FONT.mono, bullet: false }),
  ], { x: 0.5, y: 2.4, w: 6.0, h: 4.3, fontFace: FONT.body, fontSize: 16, color: C.ink, valign: 'top', paraSpaceAfter: 6 });

  s.addText('What you don\'t', { x: 6.83, y: 1.9, w: 6.0, h: 0.4, fontFace: FONT.body, fontSize: 16, bold: true, color: C.bad });
  s.addText([
    bullet('Logging', { fontSize: 16 }),
    bullet('Caching', { fontSize: 16 }),
    bullet('Batching', { fontSize: 16 }),
    bullet('Retry', { fontSize: 16 }),
    bullet('Type-safe results', { fontSize: 16 }),
  ], { x: 6.83, y: 2.4, w: 6.0, h: 4.3, fontFace: FONT.body, fontSize: 16, color: C.ink, valign: 'top', paraSpaceAfter: 6 });

  addNotes(s, `"If you've ever written fetch against /_api/web, you know the headers dance. SPFx hides that. That's already a win over raw fetch."`);
}

// SLIDE — SPFx + SP REST: read
{
  const s = pptx.addSlide({ masterName: 'SPACE_BASE' });
  addTitle(s, 'SPFx + SP REST: read');
  addCode(s,
`public async getItems(list: IListIdentifier): Promise<IEventItem[]> {
  const url = \`\${this.siteUrl}/_api/web/lists/getbytitle('\${list.title}')/items\`
    + \`?$select=Id,Title,Session,SessionDate,SessionType,EventSite,SessionLink,\`
    +   \`Speaker/Id,Speaker/Title,Speaker/EMail\`
    + \`&$expand=Speaker\`
    + \`&$filter=\${encodeURIComponent(\`SessionDate ge datetime'\${startOfTodayIso()}'\`)}\`
    + \`&$orderby=SessionDate asc\`;
  const response = await this.spHttpClient.get(url, SPHttpClient.configurations.v1);
  if (!response.ok) throw new Error(\`Failed to get items: \${response.statusText}\`);
  const data = await response.json();
  return data.value as IEventItem[];
}`,
    0.5, 1.9, 12.33, 4.0, 12);
  s.addText('It works. It\'s also a string-concatenated URL, a manual cast, and zero observability.',
    { x: 0.5, y: 6.05, w: 12.33, h: 0.5, fontFace: FONT.body, fontSize: 16, color: C.violet, italic: true });

  addNotes(s, 'Point out the URL — $select, $expand, $filter, $orderby. Memorize this URL shape; it matters in 30 minutes. Highlight the as IEventItem[] cast — TypeScript trusts you, SharePoint doesn\'t validate.');
}

// SLIDE — SPFx + SP REST: write
{
  const s = pptx.addSlide({ masterName: 'SPACE_BASE' });
  addTitle(s, 'SPFx + SP REST: write');
  addCode(s,
`public async updateItem(list, itemId, item): Promise<IEventItem> {
  const url = \`\${this.siteUrl}/_api/web/lists/getbytitle('\${list.title}')/items(\${itemId})\`;
  const options = {
    headers: { 'IF-MATCH': '*', 'X-HTTP-Method': 'MERGE' },
    body: JSON.stringify(toSpWritePayload(item))
  };
  const response = await this.spHttpClient.post(url, SPHttpClient.configurations.v1, options);
  if (!response.ok) throw new Error(\`Failed to update item \${itemId}: \${response.statusText}\`);
  return { ...item, Id: itemId };
}`,
    0.5, 1.9, 12.33, 3.2, 12);

  s.addText('Note the ceremony', { x: 0.5, y: 5.25, w: 12.33, h: 0.4, fontFace: FONT.body, fontSize: 14, bold: true, color: C.amber });
  s.addText([
    bullet("IF-MATCH: '*' — opt out of ETag concurrency", { fontSize: 14 }),
    bullet("X-HTTP-Method: MERGE — because it's a POST pretending to be a PATCH", { fontSize: 14 }),
    bullet('A separate toSpWritePayload(item) mapper for SharePoint field shapes', { fontSize: 14 }),
  ], { x: 0.5, y: 5.6, w: 12.33, h: 1.3, fontFace: FONT.body, fontSize: 14, color: C.ink, valign: 'top', paraSpaceAfter: 4 });

  s.addText('▶ Demo 1', { x: 11.0, y: 6.55, w: 1.83, h: 0.35, fontFace: FONT.body, fontSize: 12, bold: true, color: C.cyan, align: 'right' });
}

// SLIDE — When SP REST stops being enough
{
  const s = pptx.addSlide({ masterName: 'SPACE_BASE' });
  addTitle(s, 'When SP REST stops being enough');
  addSubtitle(s, 'Reach for MS Graph when…');
  addBody(s, [
    bullet('You need data outside SharePoint (Outlook, Teams, OneDrive, users, groups)'),
    bullet('You need to span sites or tenants without site-collection-scoped tokens'),
    bullet('You want one API instead of three'),
    plain(' ', { bullet: false }),
    plain('Caveat: Graph for SharePoint list items is awkward. You\'ll see why in a minute.', { bullet: false, color: C.amber, italic: true, fontSize: 18 }),
  ]);
  addNotes(s, 'This is where the "use Graph for everything" advice falls down. Graph is great for non-SP workloads. For SP lists, REST is often cleaner.');
}

// SLIDE — SPFx + Graph: read
{
  const s = pptx.addSlide({ masterName: 'SPACE_BASE' });
  addTitle(s, 'SPFx + Graph: read list items');
  addCode(s,
`const qs = [
  \`expand=fields(select=\${GRAPH_FIELD_SELECT})\`,            // <-- no $ on expand
  \`$filter=fields/SessionDate ge '\${startOfTodayIsoNoMs()}'\`,// <-- single-quoted ISO
  \`$orderby=fields/SessionDate asc\`
].join('&');

const response = await this.graphClient
  .api(\`/sites/\${this.siteId}/lists/\${list.id}/items\`)
  .version('v1.0')
  .header('Prefer', 'HonorNonIndexedQueriesWarningMayFailRandomly')
  .query(qs)
  .get();`,
    0.5, 1.9, 12.33, 3.0, 12);

  s.addText('Three Graph-specific gotchas already on this slide', { x: 0.5, y: 5.05, w: 12.33, h: 0.4, fontFace: FONT.body, fontSize: 14, bold: true, color: C.amber });
  s.addText([
    { text: '1. expand=fields(select=...) — no $. With $expand you get an OData parser error.', options: { color: C.ink, fontSize: 14 } },
    { text: "2. Datetime literal must be a single-quoted ISO 8601 string. datetime'...' returns 400.", options: { color: C.ink, fontSize: 14 } },
    { text: '3. $filter / $orderby against fields/<column> requires the column to be indexed.', options: { color: C.ink, fontSize: 14 } },
  ], { x: 0.5, y: 5.4, w: 12.33, h: 1.3, fontFace: FONT.body, fontSize: 14, color: C.ink, valign: 'top', paraSpaceAfter: 4 });

  addNotes(s, 'The Prefer header is a fallback for non-indexed columns — unreliable. Indexing is the real fix. "If you\'ve never seen these errors, that\'s because you\'ve been using PnPjs and you didn\'t know it was hiding them."');
}

// SLIDE — Graph Explorer detour
{
  const s = pptx.addSlide({ masterName: 'SPACE_BASE' });
  addTitle(s, 'The Graph Explorer detour');
  addSubtitle(s, 'Same SPFx MSGraphClientV3, no list at all — just path → JSON');
  addCode(s,
`public async runQuery(path: string): Promise<unknown> {
  return await this.graphClient.api(path).version('v1.0').get();
}`,
    0.5, 2.0, 12.33, 1.4, 14);

  s.addText('Use this in real life when prototyping or exploring a Graph endpoint before you commit to a service implementation. Cheaper than Postman because the auth is already there.',
    { x: 0.5, y: 3.7, w: 12.33, h: 1.5, fontFace: FONT.body, fontSize: 16, color: C.ink, valign: 'top' });

  s.addText('▶ Demo 2', { x: 11.0, y: 6.55, w: 1.83, h: 0.35, fontFace: FONT.body, fontSize: 12, bold: true, color: C.cyan, align: 'right' });
}

// SLIDE — SPFx + anonymous APIs
{
  const s = pptx.addSlide({ masterName: 'SPACE_BASE' });
  addTitle(s, 'SPFx + anonymous APIs');
  addSubtitle(s, 'The third HTTP client in @microsoft/sp-http: HttpClient. No auth. Plain fetch.');
  addCode(s,
`const response = await this.httpClient.get(
  'https://official-joke-api.appspot.com/random_joke',
  HttpClient.configurations.v1
);
const joke: IJokeResponse = await response.json();
return { id: joke.id, setup: joke.setup, punchline: joke.punchline };`,
    0.5, 2.0, 12.33, 2.2, 14);

  s.addText('Use this for any external API that doesn\'t need your tenant\'s auth.',
    { x: 0.5, y: 4.5, w: 12.33, h: 0.5, fontFace: FONT.body, fontSize: 16, color: C.violet, italic: true });

  s.addText('▶ Demo 3', { x: 11.0, y: 6.55, w: 1.83, h: 0.35, fontFace: FONT.body, fontSize: 12, bold: true, color: C.cyan, align: 'right' });
}

// SLIDE 11 — THE PIVOT (keystone)
{
  const s = pptx.addSlide({ masterName: 'SPACE_BASE' });
  addTitle(s, 'Pivot: notice anything?');
  addSubtitle(s, 'Look at what we wrote across three services');

  const headers = ['', 'Logging', 'Caching', 'Batching', 'Retry', 'Types'];
  const rows = [
    ['SPHttpClient', '❌', '❌', '❌', '❌', 'cast'],
    ['MSGraphClientV3', '❌', '❌', '❌', '❌', 'cast'],
    ['HttpClient', '❌', '❌', '❌', '❌', 'cast'],
  ];
  const tableRows = [
    headers.map((h, i) => ({ text: h, options: { bold: true, color: C.cyan, fill: { color: C.bgPanelHi }, fontFace: FONT.body, fontSize: 16, align: i === 0 ? 'left' : 'center' } })),
    ...rows.map((r) => r.map((cell, i) => ({
      text: cell,
      options: {
        bold: i === 0,
        color: i === 0 ? C.ink : (cell === '❌' ? C.bad : C.amber),
        fill: { color: C.bgPanel },
        fontFace: i === 0 ? FONT.body : (i === 5 ? FONT.mono : FONT.body),
        fontSize: i === 0 ? 16 : 18,
        align: i === 0 ? 'left' : 'center',
      },
    }))),
  ];
  s.addTable(tableRows, {
    x: 0.5, y: 2.0, w: 12.33,
    colW: [3.0, 1.866, 1.866, 1.866, 1.866, 1.866],
    rowH: 0.55,
    border: { type: 'solid', pt: 0.75, color: C.rule },
  });

  // Pull-quote
  s.addShape('rect', { x: 0.5, y: 4.6, w: 0.06, h: 1.4, fill: { color: C.violet }, line: { color: C.violet, width: 0 } });
  s.addText([
    { text: 'Every empty cell is a problem you\'d solve yourself.', options: { color: C.ink, fontSize: 22, bold: true, breakLine: true } },
    { text: 'PnPjs already solved them.', options: { color: C.cyan, fontSize: 22, bold: true } },
  ], { x: 0.75, y: 4.6, w: 12.0, h: 1.4, fontFace: FONT.body, valign: 'middle', paraSpaceAfter: 6 });

  addNotes(s,
    `This is the turn. Pause here. Don't rush. "We just wrote three services. Every one of them needs the same five things added before it's production-ready. We're going to write the same three services again with PnPjs, and you're going to watch all five columns light up."

Disambiguation: the [DataDemo] lines in console aren't the data layer logging — those are component-level Logger.info calls announcing user actions. The empty-cell claim is about automatic request/response logging at the data layer, which we wire up for free in the PnPjs pass. Keep DevTools focused on Network tab during Pass 1.`);
}

// SECTION DIVIDER — PASS 2
{
  const s = pptx.addSlide({ masterName: 'SPACE_SECTION' });
  s.addText('PASS 2', { x: 0.5, y: 2.5, w: 12.33, h: 0.8, fontFace: FONT.body, fontSize: 18, color: C.violet, align: 'center', charSpacing: 12 });
  s.addText('The PnPjs Way', { x: 0.5, y: 3.3, w: 12.33, h: 1.4, fontFace: FONT.body, fontSize: 64, bold: true, color: C.cyan, align: 'center' });
  s.addShape('rect', { x: 6.17, y: 4.85, w: 1.0, h: 0.05, fill: { color: C.violet }, line: { color: C.violet, width: 0 } });
}

// SLIDE — What PnPjs is
{
  const s = pptx.addSlide({ masterName: 'SPACE_BASE' });
  addTitle(s, 'What PnPjs is');
  addSubtitle(s, 'A fluent, composable HTTP pipeline for SharePoint, Graph, and arbitrary endpoints');

  s.addText('Three packages today', { x: 0.5, y: 1.9, w: 12.33, h: 0.4, fontFace: FONT.body, fontSize: 16, bold: true, color: C.violet });
  s.addText([
    bullet('@pnp/sp — SharePoint REST, fluent', { fontSize: 16 }),
    bullet('@pnp/graph — MS Graph, fluent', { fontSize: 16 }),
    bullet('@pnp/queryable — the underlying pipeline (also works against any URL)', { fontSize: 16 }),
  ], { x: 0.5, y: 2.3, w: 12.33, h: 1.5, fontFace: FONT.body, fontSize: 16, color: C.ink, valign: 'top', paraSpaceAfter: 6 });

  s.addText('Init pattern (once per service)', { x: 0.5, y: 3.85, w: 12.33, h: 0.4, fontFace: FONT.body, fontSize: 14, bold: true, color: C.violet });
  addCode(s,
`const sp    = spfi(siteUrl).using(spSPFx(this.context));
const graph = graphfi().using(graphSPFx(this.context));`,
    0.5, 4.25, 12.33, 1.2, 14);

  s.addText('Same SPFx context, same auth. PnPjs is not a separate auth story — it rides on top of the SPFx clients you already have.',
    { x: 0.5, y: 5.6, w: 12.33, h: 0.7, fontFace: FONT.body, fontSize: 14, color: C.violet, italic: true, valign: 'top' });
}

// SLIDE — The URL reveal (keystone)
{
  const s = pptx.addSlide({ masterName: 'SPACE_BASE' });
  addTitle(s, 'The URL reveal');

  s.addText('SPFx-native', { x: 0.5, y: 1.6, w: 6.0, h: 0.4, fontFace: FONT.body, fontSize: 14, bold: true, color: C.amber });
  addCode(s,
`spHttpClient.get(
  \`\${siteUrl}/_api/web/lists/getbytitle('Speaking Events')/items\`
  + \`?$select=Id,Title,SessionDate&$orderby=SessionDate desc&$top=50\`,
  SPHttpClient.configurations.v1
);`,
    0.5, 2.0, 6.16, 2.2, 11);

  s.addText('PnPjs', { x: 6.83, y: 1.6, w: 6.0, h: 0.4, fontFace: FONT.body, fontSize: 14, bold: true, color: C.good });
  addCode(s,
`sp.web.lists.getByTitle('Speaking Events').items
  .select('Id', 'Title', 'SessionDate')
  .orderBy('SessionDate', false)
  .top(50)();`,
    6.83, 2.0, 6.0, 2.2, 11);

  // Pull quote
  s.addShape('rect', { x: 0.5, y: 4.5, w: 0.06, h: 1.6, fill: { color: C.cyan }, line: { color: C.cyan, width: 0 } });
  s.addText([
    { text: 'Both produce the same network request.', options: { color: C.ink, fontSize: 22, bold: true, breakLine: true } },
    { text: ' ', options: { fontSize: 8, breakLine: true } },
    { text: 'PnPjs is not magic.', options: { color: C.cyan, fontSize: 18, italic: true, breakLine: true } },
    { text: 'It\'s the URL you\'d write anyway, just typed for you.', options: { color: C.cyan, fontSize: 18, italic: true } },
  ], { x: 0.75, y: 4.5, w: 12.0, h: 1.6, fontFace: FONT.body, valign: 'middle', paraSpaceAfter: 4 });

  s.addText('▶ Demo 4 — keystone', { x: 9.5, y: 6.55, w: 3.33, h: 0.35, fontFace: FONT.body, fontSize: 12, bold: true, color: C.cyan, align: 'right' });

  addNotes(s, 'This is the most important slide in the deck. Stay on it. The audience needs to leave this slide trusting that PnPjs isn\'t doing anything weird underneath. That trust is what makes the next slides land.');
}

// SLIDE — PnPjs + SP REST: read
{
  const s = pptx.addSlide({ masterName: 'SPACE_BASE' });
  addTitle(s, 'PnPjs + SP REST: read');
  addCode(s,
`public async getItems(list): Promise<IEventItem[]> {
  return await this.sp.web.lists
    .getByTitle(list.title)
    .items
    .select('Id', 'Title', 'Session', 'SessionDate', 'SessionType',
            'EventSite', 'SessionLink',
            'Speaker/Id', 'Speaker/Title', 'Speaker/EMail')
    .expand('Speaker')
    .filter(\`SessionDate ge datetime'\${startOfTodayIso()}'\`)
    .orderBy('SessionDate', true)() as IEventItem[];
}`,
    0.5, 1.9, 12.33, 3.6, 13);

  s.addText('Same fields. Same filter. Same order.', { x: 0.5, y: 5.7, w: 12.33, h: 0.4, fontFace: FONT.body, fontSize: 16, color: C.ink });
  s.addText('No URL string. No response.ok check. No response.json().', { x: 0.5, y: 6.05, w: 12.33, h: 0.4, fontFace: FONT.body, fontSize: 16, bold: true, color: C.cyan });

  s.addText('▶ Demo 5', { x: 11.0, y: 6.55, w: 1.83, h: 0.35, fontFace: FONT.body, fontSize: 12, bold: true, color: C.cyan, align: 'right' });
}

// SLIDE — PnPjs + SP REST: write
{
  const s = pptx.addSlide({ masterName: 'SPACE_BASE' });
  addTitle(s, 'PnPjs + SP REST: write');
  addCode(s,
`public async updateItem(list, itemId, item): Promise<IEventItem> {
  await this.sp.web.lists
    .getByTitle(list.title)
    .items
    .getById(itemId)
    .update(toSpWritePayload(item));
  return { ...item, Id: itemId };
}`,
    0.5, 1.9, 12.33, 2.6, 14);

  s.addText('Compare to SPFx', { x: 0.5, y: 4.7, w: 12.33, h: 0.4, fontFace: FONT.body, fontSize: 14, bold: true, color: C.violet });
  s.addText([
    bullet('No IF-MATCH header.', { fontSize: 16 }),
    bullet('No X-HTTP-Method: MERGE.', { fontSize: 16 }),
    bullet('No URL.', { fontSize: 16 }),
  ], { x: 0.5, y: 5.05, w: 12.33, h: 1.4, fontFace: FONT.body, fontSize: 16, color: C.ink, valign: 'top', paraSpaceAfter: 4 });

  s.addText('Same toSpWritePayload mapper — that\'s a SharePoint shape problem, not an HTTP-client problem.',
    { x: 0.5, y: 6.45, w: 12.33, h: 0.5, fontFace: FONT.body, fontSize: 13, color: C.inkDim, italic: true });
}

// SLIDE — PnPjs + Graph: same operations
{
  const s = pptx.addSlide({ masterName: 'SPACE_BASE' });
  addTitle(s, 'PnPjs + Graph: same operations');
  addCode(s,
`const itemsQuery = this.graph.sites
  .getById(this.siteId)
  .lists
  .getById(list.id)
  .items
  .using(InjectHeaders({ Prefer: 'HonorNonIndexedQueriesWarningMayFailRandomly' }));

itemsQuery.query.set('expand', \`fields(select=\${GRAPH_FIELD_SELECT})\`);
itemsQuery.query.set('$filter', \`fields/SessionDate ge '\${startOfTodayIsoNoMs()}'\`);
itemsQuery.query.set('$orderby', 'fields/SessionDate asc');

const items = await itemsQuery() as IGraphListItem[];`,
    0.5, 1.9, 12.33, 3.5, 12);

  s.addText('Honesty check', { x: 0.5, y: 5.6, w: 12.33, h: 0.4, fontFace: FONT.body, fontSize: 14, bold: true, color: C.amber });
  s.addText('The Graph SP-list awkwardness is a Graph problem, not a PnPjs problem. Both clients have to deal with it. PnPjs at least gives you InjectHeaders and a query builder.',
    { x: 0.5, y: 5.95, w: 12.33, h: 0.9, fontFace: FONT.body, fontSize: 14, color: C.ink, valign: 'top' });

  s.addText('▶ Demo 6', { x: 11.0, y: 6.55, w: 1.83, h: 0.35, fontFace: FONT.body, fontSize: 12, bold: true, color: C.cyan, align: 'right' });
}

// SLIDE — PnPjs + anonymous
{
  const s = pptx.addSlide({ masterName: 'SPACE_BASE' });
  addTitle(s, 'PnPjs + anonymous');
  addSubtitle(s, 'Yes, that works too');
  addCode(s,
`const q = new Queryable('https://official-joke-api.appspot.com/random_joke');
q.using(BrowserFetch(), RejectOnError(), ResolveOnData(), JSONParse());
const joke: IJokeResponse = await q();`,
    0.5, 2.0, 12.33, 1.7, 14);

  s.addText('The same pipeline that talks to SharePoint can talk to anything. You just compose the behaviors you need.',
    { x: 0.5, y: 4.0, w: 12.33, h: 0.7, fontFace: FONT.body, fontSize: 16, color: C.ink, valign: 'top' });

  s.addText('Real-world note: strip X-PnPjs-RequestId on external calls to avoid CORS preflight — see the source.',
    { x: 0.5, y: 4.85, w: 12.33, h: 0.5, fontFace: FONT.body, fontSize: 13, color: C.inkDim, italic: true });
}

// SLIDE — The free upgrades (combined)
{
  const s = pptx.addSlide({ masterName: 'SPACE_BASE' });
  addTitle(s, 'The free upgrades');
  addSubtitle(s, 'Three toggles. Three real problems gone.');

  addCode(s,
`// 1. Logging — @pnp/logging wrapped once; level driven by a property-pane toggle.
Logger.attachConsoleListener();
Logger.setLevel(this.properties.enhancedLogging ? LogLevel.Verbose : LogLevel.Warning);

// 2. Caching — one line on the SPFI, toggled by un/commenting in ServiceFactory.
const sp = spfi(siteUrl)
  .using(spSPFx(this.context))
  .using(Caching({ store: 'session' }));   // <-- the only line that flips

// 3. Batching — pack many reads (or writes) into one $batch envelope.
const [batchedSp, execute] = sp.batched({ maxRequests: 100 });
for (let i = 0; i < pageCount; i++) {
  pagePromises.push(
    batchedSp.web.lists.getByTitle('Speaking Events').items
      .top(pageSize).skip(i * pageSize)() as Promise<IEventItem[]>
  );
}
await execute();   // one HTTP request to /_api/$batch`,
    0.5, 1.9, 12.33, 4.5, 11);

  s.addText('▶ Demo 7', { x: 11.0, y: 6.55, w: 1.83, h: 0.35, fontFace: FONT.body, fontSize: 12, bold: true, color: C.cyan, align: 'right' });

  addNotes(s, 'Each upgrade in the demo is a single-line toggle (a property pane checkbox for logging, an uncomment for caching, a commented block swap for batching). No live editing of fluent chains under pressure. The Caching demo is the visceral one — Click → wait → click → instant. The batching demo is the architectural one.');
}

// SLIDE — What logging gets you
{
  const s = pptx.addSlide({ masterName: 'SPACE_BASE' });
  addTitle(s, 'What logging gets you');
  addCode(s,
`[DataDemo] onInit: starting (enhancedLogging=on)
[DataDemo] loadItems: requesting items from list Speaking Events
[DataDemo] loadItems: received 12 item(s)        ▶ (12) [{…}, {…}, …]`,
    0.5, 1.9, 12.33, 1.8, 13);

  s.addText('A thin wrapper over @pnp/logging (info / debug / warn / error) emits via console.* so object payloads render as collapsible trees. Listeners are pluggable — ConsoleListener today, App Insights tomorrow. Same call sites.',
    { x: 0.5, y: 4.0, w: 12.33, h: 1.5, fontFace: FONT.body, fontSize: 15, color: C.ink, valign: 'top' });

  addNotes(s, 'Tease: "If you\'re shipping into production without observability on your data calls, this is the easiest win you\'ll ever take." The wrapper isn\'t a hard requirement — Logger.subscribe(ConsoleListener()) works out of the box. The wrapper just lets us prefix [DataDemo] and route through console.info/console.debug for the collapsible-tree affordance.');
}

// SLIDE — What caching gets you
{
  const s = pptx.addSlide({ masterName: 'SPACE_BASE' });
  addTitle(s, 'What caching gets you');
  addSubtitle(s, 'Two scopes. Pick one.');
  addCode(s,
`// Global: every read on this SPFI is cached.
const sp = spfi(siteUrl)
  .using(spSPFx(this.context))
  .using(Caching({ store: 'session' }));   // 'session' | 'local'

// Per-call: opt one query out of caching when global is on.
.items.using(CacheNever())()

// Per-call: opt one query in with a custom expiry when global is off.
.items.using(Caching({ store: 'session', expireFunc: () => addMinutes(new Date(), 5) }))()`,
    0.5, 1.9, 12.33, 3.0, 12);

  s.addText('Defaults: 60-second expiry, keyed by URL. \'session\' survives a hard refresh; \'local\' survives a tab close. Default is in-memory and dies with the page.',
    { x: 0.5, y: 5.05, w: 12.33, h: 0.7, fontFace: FONT.body, fontSize: 13, color: C.violet, italic: true, valign: 'top' });

  s.addText([
    { text: 'Use it for:  ', options: { bold: true, color: C.good, fontSize: 13 } },
    { text: 'reference data · repeated reads · anything you\'d hate to re-fetch on tab switch', options: { color: C.ink, fontSize: 13, breakLine: true } },
    { text: 'Avoid for:  ', options: { bold: true, color: C.bad, fontSize: 13 } },
    { text: 'frequently-changing transactional data · anything where stale = wrong (use CacheNever() there)', options: { color: C.ink, fontSize: 13 } },
  ], { x: 0.5, y: 5.85, w: 12.33, h: 1.0, fontFace: FONT.body, valign: 'top', paraSpaceAfter: 4 });
}

// SLIDE — What batching gets you
{
  const s = pptx.addSlide({ masterName: 'SPACE_BASE' });
  addTitle(s, 'What batching gets you');
  addSubtitle(s, 'Reads, in our demo. Writes work the same way.');
  addCode(s,
`const [batchedSp, execute] = this.sp.batched({ maxRequests: 100 });
const pagePromises: Promise<IEventItem[]>[] = [];
for (let i = 0; i < pageCount; i++) {
  pagePromises.push(
    batchedSp.web.lists.getByTitle(list.title).items
      .select('Id', 'Title', 'SessionDate', /* ... */)
      .orderBy('SessionDate', true)
      .top(pageSize).skip(i * pageSize)() as Promise<IEventItem[]>
  );
}
await execute();
const items = (await Promise.all(pagePromises)).flat();`,
    0.5, 1.9, 12.33, 3.0, 12);

  s.addText('100 paginated reads → one HTTP request to /_api/$batch. The same wiring works for add/update/delete.',
    { x: 0.5, y: 5.05, w: 12.33, h: 0.5, fontFace: FONT.body, fontSize: 16, bold: true, color: C.cyan });

  s.addText('Why it matters', { x: 0.5, y: 5.65, w: 12.33, h: 0.4, fontFace: FONT.body, fontSize: 13, bold: true, color: C.violet });
  s.addText([
    bullet('Latency: one round trip beats N. Especially over slow links.', { fontSize: 13 }),
    bullet('Throttling: SharePoint counts requests, not inner operations. Your throttle escape hatch.', { fontSize: 13 }),
    bullet('Atomicity-ish: not a real transaction, but closer than N independent calls.', { fontSize: 13 }),
  ], { x: 0.5, y: 5.95, w: 12.33, h: 1.0, fontFace: FONT.body, fontSize: 13, color: C.ink, valign: 'top', paraSpaceAfter: 2 });

  addNotes(s, 'Be honest: $batch is not transactional. If page 3 fails, pages 1 and 2 still came back. PnPjs surfaces per-operation results so you can detect partial failures. The demo uses paginated reads because the Speaking Events list isn\'t big enough for a bulk-write story to land. The pattern is identical for writes.');
}

// SECTION DIVIDER — WRAP
{
  const s = pptx.addSlide({ masterName: 'SPACE_SECTION' });
  s.addText('WRAP', { x: 0.5, y: 2.5, w: 12.33, h: 0.8, fontFace: FONT.body, fontSize: 18, color: C.violet, align: 'center', charSpacing: 12 });
  s.addText('When and Why', { x: 0.5, y: 3.3, w: 12.33, h: 1.4, fontFace: FONT.body, fontSize: 64, bold: true, color: C.cyan, align: 'center' });
  s.addShape('rect', { x: 6.17, y: 4.85, w: 1.0, h: 0.05, fill: { color: C.violet }, line: { color: C.violet, width: 0 } });
}

// SLIDE — When to still reach for SPFx-native
{
  const s = pptx.addSlide({ masterName: 'SPACE_BASE' });
  addTitle(s, 'When to still reach for SPFx-native');
  addSubtitle(s, 'PnPjs is the default. But it\'s not free.');
  addBody(s, [
    bullet('Bundle size matters. PnPjs adds ~50–80 KB depending on imports. For a single-purpose tile that does one Graph call, that\'s a lot.'),
    bullet('One-off calls with no list interaction (a single /me ping, a webhook trigger).'),
    bullet('Debugging a network issue — fewer abstractions between you and the wire.'),
    bullet('Your team doesn\'t know PnPjs yet and the call is trivial.'),
    plain(' ', { bullet: false }),
    plain('Default: PnPjs. Exceptions: above.', { bullet: false, color: C.cyan, italic: true, fontSize: 18, bold: true }),
  ]);
}

// SLIDE — Decision cheat sheet
{
  const s = pptx.addSlide({ masterName: 'SPACE_BASE' });
  addTitle(s, 'Decision cheat sheet');

  const headers = ['Scenario', 'Use'];
  const rows = [
    ['List CRUD on the current site', 'PnPjs SP'],
    ['List CRUD across sites (one tenant)', 'PnPjs Graph'],
    ['Cross-tenant or app-only', 'PnPjs Graph (with custom auth)'],
    ['One-off anonymous fetch', 'SPFx HttpClient'],
    ['Many anonymous fetches, want consistency', 'PnPjs Queryable'],
    ['Prototyping a Graph endpoint', 'SPFx MSGraphClientV3 (or Graph Explorer)'],
    ['Single-purpose perf-critical tile', 'SPFx native'],
    ['You need logging/caching/batching', 'PnPjs (every time)'],
  ];
  const tableRows = [
    headers.map((h) => ({ text: h, options: { bold: true, color: C.cyan, fill: { color: C.bgPanelHi }, fontFace: FONT.body, fontSize: 14 } })),
    ...rows.map((r) => r.map((cell, i) => ({
      text: cell,
      options: { color: i === 0 ? C.ink : C.cyan, fill: { color: C.bgPanel }, fontFace: FONT.body, fontSize: 13, bold: i === 1 },
    }))),
  ];
  s.addTable(tableRows, {
    x: 0.5, y: 1.9, w: 12.33,
    colW: [6.5, 5.83],
    rowH: 0.42,
    border: { type: 'solid', pt: 0.75, color: C.rule },
  });
}

// SLIDE — Resources & takeaways
{
  const s = pptx.addSlide({ masterName: 'SPACE_BASE' });
  addTitle(s, 'Three takeaways');

  s.addText([
    { text: '1.  ', options: { color: C.cyan, fontSize: 22, bold: true } },
    { text: 'PnPjs makes the same HTTP request, just typed for you.', options: { color: C.ink, fontSize: 22, breakLine: true } },
    { text: ' ', options: { fontSize: 8, breakLine: true } },
    { text: '2.  ', options: { color: C.cyan, fontSize: 22, bold: true } },
    { text: 'Logging, caching, and batching are one toggle each.', options: { color: C.ink, fontSize: 22, breakLine: true } },
    { text: ' ', options: { fontSize: 8, breakLine: true } },
    { text: '3.  ', options: { color: C.cyan, fontSize: 22, bold: true } },
    { text: 'Reach for SPFx-native deliberately, not by default.', options: { color: C.ink, fontSize: 22 } },
  ], { x: 0.5, y: 1.9, w: 12.33, h: 3.0, fontFace: FONT.body, valign: 'top', paraSpaceAfter: 8 });

  s.addText('Resources', { x: 0.5, y: 5.05, w: 12.33, h: 0.4, fontFace: FONT.body, fontSize: 14, bold: true, color: C.violet });
  s.addText([
    bullet('PnPjs docs — https://pnp.github.io/pnpjs/', { fontSize: 13 }),
    bullet('SPFx HTTP clients — https://learn.microsoft.com/sharepoint/dev/spfx/web-parts/basics/connect-to-sharepoint', { fontSize: 13 }),
    bullet('Microsoft Graph — https://learn.microsoft.com/graph/', { fontSize: 13 }),
    bullet('This demo project — https://github.com/DonKirkham/DataDemo', { fontSize: 13 }),
  ], { x: 0.5, y: 5.4, w: 12.33, h: 1.5, fontFace: FONT.body, fontSize: 13, color: C.ink, valign: 'top', paraSpaceAfter: 2 });
}

// SLIDE — Q&A
{
  const s = pptx.addSlide({ masterName: 'SPACE_SECTION' });
  s.addText('Q&A', { x: 0.5, y: 2.8, w: 12.33, h: 1.6, fontFace: FONT.body, fontSize: 96, bold: true, color: C.cyan, align: 'center' });
  s.addText('Questions?', { x: 0.5, y: 4.6, w: 12.33, h: 0.6, fontFace: FONT.body, fontSize: 24, color: C.violet, italic: true, align: 'center' });

  addNotes(s, `Likely questions to be ready for:
- "Does PnPjs work with app-only auth?" → Yes, but not via SPFx behavior — you provide your own MSAL token via using().
- "Bundle impact?" → Tree-shakes well. Import the selectives (@pnp/sp/webs, @pnp/sp/lists, @pnp/sp/items) not the barrel.
- "PnPjs v3 vs v4?" → v4 is current. v3 → v4 was the batching API rewrite (sp.createBatch() → sp.batched()).
- "Why not just use Graph for everything?" → You saw the expand=fields mess. SP REST is cleaner for SP data.`);
}

// ---------- WRITE ----------
await pptx.writeFile({ fileName: OUT_PATH });
console.log(`Wrote ${OUT_PATH}`);
