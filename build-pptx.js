const pptxgen = require('pptxgenjs');
const pres = new pptxgen();
pres.layout = 'LAYOUT_WIDE'; // 13.333 x 7.5 in
pres.author = 'FastForward Global';
pres.title = "DDX'26 San Diego — Recap";

const W = 13.333, H = 7.5;
const px = (n) => n * (W / 1920);          // 1920px design grid → inches
const pt = (n) => n / 2;                    // px → pt at this scale
const IMG = __dirname + '/img/';
const BG = '050505', CARD = '0B0B0B', LINE = '242424', LINE_SOFT = '1A1A1A', Y = 'FFF204', WHITE = 'FFFFFF', DIM = '9A9A9A', FAINT = '5C5C5C';
const SANS = 'Arial', SERIF = 'Times New Roman';
const M = px(120), TOP = px(96);

function base(n) {
  const s = pres.addSlide();
  s.background = { color: BG };
  if (n > 1) {
    s.addImage({ path: IMG + 'ddx-wordmark.png', x: M, y: H - px(64), w: px(64), h: px(64) * (89 / 205) });
    s.addText("DDX’26 SAN DIEGO · RECAP", { x: W / 2 - 2, y: H - px(70), w: 4, h: px(30), align: 'center', fontFace: SANS, fontSize: 8, color: FAINT, charSpacing: 3, margin: 0, isTextBox: true });
    s.addText(String(n).padStart(2, '0'), { x: W - M - 1, y: H - px(70), w: 1, h: px(30), align: 'right', fontFace: SANS, fontSize: 8, color: FAINT, charSpacing: 3, margin: 0, isTextBox: true });
  }
  return s;
}
function eyebrow(s, y, a, b) {
  s.addShape(pres.shapes.RECTANGLE, { x: M, y: y + px(9), w: px(30), h: px(3), fill: { color: Y }, line: { color: Y, width: 0 } });
  const runs = [{ text: a.toUpperCase(), options: { color: Y, bold: true } }];
  if (b) runs.push({ text: '   ' + b.toUpperCase(), options: { color: 'A6A6A6', bold: true } });
  s.addText(runs, { x: M + px(44), y, w: W - 2 * M - px(44), h: px(24), fontFace: SANS, fontSize: 10, charSpacing: 5, margin: 0, isTextBox: true, valign: 'middle' });
}
function h1(s, text, y, size = 104) {
  s.addText(text.toUpperCase(), { x: M, y, w: W - 2 * M, h: px(size * 1.9), fontFace: SANS, fontSize: pt(size), bold: true, color: WHITE, charSpacing: -4, lineSpacingMultiple: 0.92, margin: 0, isTextBox: true, valign: 'top' });
}
function line(s, x, y, w) { s.addShape(pres.shapes.LINE, { x, y, w, h: 0, line: { color: LINE, width: 0.75 } }); }
function vline(s, x, y, h) { s.addShape(pres.shapes.LINE, { x, y, w: 0, h, line: { color: LINE, width: 0.75 } }); }
function circle(s, file, x, y, d) { s.addImage({ path: IMG + file, x, y, w: d, h: d, rounding: true, sizing: { type: 'cover', w: d, h: d } }); }
function photo(s, file, x, y, w, h) { s.addImage({ path: IMG + file, x, y, w, h, sizing: { type: 'cover', w, h } }); }
function evidence(s, x, y, w, k, n, t, src) {
  s.addShape(pres.shapes.RECTANGLE, { x, y, w: px(3), h: px(290), fill: { color: Y }, line: { color: Y, width: 0 } });
  const xx = x + px(24), ww = w - px(24);
  s.addText(k.toUpperCase(), { x: xx, y, w: ww, h: px(20), fontFace: SANS, fontSize: 7.5, color: Y, charSpacing: 3, margin: 0, isTextBox: true });
  s.addText(n, { x: xx, y: y + px(24), w: ww, h: px(90), fontFace: SANS, fontSize: 40, color: WHITE, charSpacing: -2, margin: 0, isTextBox: true, valign: 'top' });
  s.addText(t, { x: xx, y: y + px(120), w: ww, h: px(120), fontFace: SANS, fontSize: 9.5, color: 'CCCCCC', margin: 0, isTextBox: true, valign: 'top', lineSpacingMultiple: 1.15 });
  s.addText(src, { x: xx, y: y + px(244), w: ww, h: px(40), fontFace: SANS, fontSize: 7.5, color: FAINT, margin: 0, isTextBox: true, valign: 'top' });
}

/* ---------- 1 cover ---------- */
{
  const s = base(1);
  photo(s, 'hero.jpg', 0, 0, W, H);
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: W, h: H, fill: { color: '050505', transparency: 62 }, line: { color: '050505', width: 0 } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: H * 0.5, w: W, h: H * 0.5, fill: { color: '050505', transparency: 45 }, line: { color: '050505', width: 0 } });
  s.addImage({ path: IMG + 'ddx-wordmark.png', x: M, y: px(80), w: px(150), h: px(150) * (89 / 205) });
  eyebrow(s, px(300), 'Conference recap', '17 September 2026');
  s.addText([{ text: 'THE\nINFLECTION\n', options: { color: WHITE } }, { text: 'POINT', options: { color: Y } }],
    { x: M, y: px(330), w: W - 2 * M, h: px(560), fontFace: SANS, fontSize: 112, bold: true, charSpacing: -6, lineSpacingMultiple: 0.84, margin: 0, isTextBox: true, valign: 'top' });
  s.addText([{ text: "DDX’26 San Diego", options: { bold: true, color: WHITE } }, { text: '      UC San Diego Park & Market      Design & Product Leaders Exchange on the Future of Digital Innovation', options: { color: 'DDDDDD' } }],
    { x: M, y: px(900), w: W - 2 * M, h: px(40), fontFace: SANS, fontSize: 13, margin: 0, isTextBox: true, valign: 'middle' });
  s.addShape(pres.shapes.LINE, { x: M, y: px(960), w: W - 2 * M, h: 0, line: { color: '555555', width: 0.75 } });
  s.addText('WITH', { x: M, y: px(985), w: px(90), h: px(50), fontFace: SANS, fontSize: 7.5, color: 'A6A6A6', charSpacing: 4, margin: 0, isTextBox: true, valign: 'middle' });
  // logos: [file, natural w, natural h, target h px]
  const logos = [['freshform.png', 614, 128, 40], ['bcca.png', 345, 181, 66], ['sd-design-week.png', 248, 243, 66], ['dscout.png', 476, 91, 32], ['gain.png', 549, 78, 32], ['marvin.png', 460, 126, 40], ['askable.png', 419, 96, 32], ['mondo-robot.png', 596, 237, 40]];
  let x = M + px(150);
  for (const [f, nw, nh, th] of logos) {
    const h = px(th), w = h * nw / nh;
    s.addImage({ path: IMG + f, x, y: px(1010) - h / 2, w, h });
    x += w + px(52);
  }
}

/* ---------- 2 brief ---------- */
{
  const s = base(2);
  eyebrow(s, TOP, 'In brief');
  h1(s, 'One day, three stages,\none question', TOP + px(40));
  const y0 = TOP + px(280);
  s.addText('If AI can produce “good enough” design on demand, what is the job of a design and product leader now?', { x: M, y: y0, w: px(760), h: px(200), fontFace: SANS, fontSize: 17, color: WHITE, margin: 0, isTextBox: true, valign: 'top', lineSpacingMultiple: 1.2 });
  s.addText('Twenty-five leaders from Adobe, Google, IDEO, Webflow, Otter.ai, Target, Netflix, Careem, TaylorMade, Marvin, Dscout and the Obama White House took the question to a room that came to argue back. Don Norman opened the day; three enterprise design leaders closed it by comparing the expectations of the role with its reality in 2026.', { x: M, y: y0 + px(210), w: px(760), h: px(260), fontFace: SANS, fontSize: 11.5, color: 'BBBBBB', margin: 0, isTextBox: true, valign: 'top', lineSpacingMultiple: 1.25 });
  const cells = [['25', 'Speakers on stage', true], ['3', 'Stages in parallel'], ['3 · 6', 'Panels · Workshops'], ['9', 'Table talks'], ['300+', 'Leaders registered', true], ['3,000+', 'DDX alumni worldwide']];
  const gx = M + px(980), gy = y0, cw = px(355), ch = px(165);
  cells.forEach(([n, l, hi], i) => {
    const cx = gx + (i % 2) * cw, cy = gy + Math.floor(i / 2) * ch;
    s.addShape(pres.shapes.RECTANGLE, { x: cx, y: cy, w: cw, h: ch, fill: { color: BG }, line: { color: LINE, width: 0.75 } });
    s.addText(n, { x: cx + px(28), y: cy + px(24), w: cw - px(40), h: px(80), fontFace: SANS, fontSize: 38, bold: false, color: hi ? Y : WHITE, charSpacing: -2, margin: 0, isTextBox: true, valign: 'top' });
    s.addText(l.toUpperCase(), { x: cx + px(28), y: cy + px(112), w: cw - px(40), h: px(30), fontFace: SANS, fontSize: 8, color: DIM, charSpacing: 3, margin: 0, isTextBox: true });
  });
}

/* ---------- 3 theme index ---------- */
const THEMES = [
  ['Good is free now. ', 'Great is the job.', 'Taste, judgment and a durable point of view are what AI cannot generate.'],
  ['Research stops being a project ', 'and becomes a pulse.', 'When teams ship weekly, episodic studies arrive after the decision.'],
  ['Design leaves the screen: ', 'decisions, systems, organisations.', 'Value is measured by influence on what gets built, not by deliverables.'],
  ['The human stays irrational, ', 'and irreplaceable.', 'AI can read the deck. It cannot read the room.'],
  ['Context is the asset. ', 'The experience is the business.', 'Tools change weekly; structured knowledge compounds.'],
];
{
  const s = base(3);
  eyebrow(s, TOP, 'What the room agreed on');
  h1(s, 'The floor rose.\nThe ceiling didn’t move.', TOP + px(40));
  const y0 = TOP + px(290), cw = (W - 2 * M) / 5, ch = px(560);
  line(s, M, y0, W - 2 * M);
  THEMES.forEach(([a, b, p], i) => {
    const x = M + i * cw;
    vline(s, x, y0, ch);
    s.addText([{ text: a, options: { color: WHITE } }, { text: b, options: { color: Y } }], { x: x + px(30), y: y0 + px(34), w: cw - px(60), h: px(300), fontFace: SANS, fontSize: 17, bold: true, charSpacing: -1, margin: 0, isTextBox: true, valign: 'top', lineSpacingMultiple: 1.05 });
    s.addText(p, { x: x + px(30), y: y0 + px(420), w: cw - px(60), h: px(120), fontFace: SANS, fontSize: 9.5, color: DIM, margin: 0, isTextBox: true, valign: 'bottom', lineSpacingMultiple: 1.2 });
  });
  vline(s, W - M, y0, ch);
}

/* ---------- 4–8 theme slides ---------- */
const THEME_SLIDES = [
  { sub: 'Craft in an AI-native world', p: 'Richard Ward called the “death of design” narrative backwards: when everyone can clear the bar for good, the differentiator is knowing what great looks like. Kevin Wong went further: AI has not changed design, it has exposed how little depth most organisations were designing with. At the tables, Jaleesa Chagan framed the edge as a shift “from capability to conviction”.',
    who: [['richard-ward.png', 'Richard Ward, Otter.ai'], ['kevin-wong.png', 'Kevin Wong, Webflow'], ['jaleesa-chagan.png', 'Jaleesa Chagan, GAIN']], photo: 'DSC09616.jpg',
    ev: [['2026 evidence', '91%', 'of designers use AI for design tasks every week, up from 54% a year earlier.', 'Designer Fund, AI in Design 2026, May 2026'], ['And yet', '49%', 'of US consumers say generative AI has made content quality worse.', 'Gartner, June 2026']] },
  { sub: 'Research at the speed of shipping', p: 'Prayag Narula made the case for always-on research. Harriet Swan and Katie Niles showed it in practice with TaylorMade’s AI caddie, CHIP: exploratory research to decide what it should say and do, A/B tests to find what moved engagement, and evaluative research to catch what the tests could not, inconclusive results included. Bhavik Gandecha’s table put it in one line: AI moves fast, research still has to lead.',
    who: [['prayag-narula.png', 'Prayag Narula, Marvin'], ['harriet-swan.png', 'Harriet Swan, GAIN'], ['katie-niles.jpg', 'Katie Niles, TaylorMade'], ['bhavik-gandecha.png', 'Bhavik Gandecha, Dscout']], photo: 'DSC09633.jpg',
    ev: [['2026 evidence', '8 → 22%', 'of organisations call research “essential to all levels of business strategy”, nearly tripling in a year.', 'Maze, Future of User Research 2026, March 2026'], ['The gap to close', '37%', 'say AI has contributed to enterprise EBIT, unchanged, while 80% say it improved their own productivity.', 'McKinsey, State of AI 2026, August 2026']] },
  { sub: 'The expanding brief', p: 'Chris Gielow asked designers to stop designing screens and start designing decisions, and to measure their value by influence on what gets built. Twisha Shah-Brandenburg took it a level up: AI transformation succeeds or fails on fragmented systems, invisible labour and unclear ownership, so the brief is now the organisation itself. The closing panel compared the expectations placed on enterprise design leaders with the reality of the role.',
    who: [['chris-gielow.png', 'Chris Gielow, UX Speakeasy'], ['twisha-shah-brandenburg.png', 'Twisha Shah-Brandenburg, Target'], ['diana-tobey.png', 'Diana Tobey, IDEO']], photo: 'DSC09693.jpg',
    ev: [['2026 evidence', '21 → 41%', 'of designers now do development work; developers doing design rose from 44% to 60%.', 'Figma, 2026 AI Report, June 2026'], ['At the top', '72%', 'of CEOs say they are now the main AI decision-maker, double the prior year.', 'BCG, AI Radar 2026, January 2026']] },
  { sub: 'Humanity-centred, still', p: 'Don Norman opened on the state of humanity-centred innovation. Rowan Salama reminded the room that the rational user most products are designed for does not exist, and that AI’s best use is understanding the “beautifully irrational” ways people decide. Stephanie Mencarelli argued that as machines generate more, human originality becomes the scarce resource, which is why joy, not productivity, was her metric. Kara Fitzpatrick: AI can read the deck, it cannot read the room.',
    who: [['don-norman.png', 'Don Norman'], ['rowan-salama.png', 'Rowan Salama, Careem'], ['stephanie-mencarelli.png', 'Stephanie Mencarelli, Adobe'], ['kara-fitzpatrick.jpg', 'Kara Fitzpatrick']], photo: 'DSC09522.jpg',
    ev: [['2026 evidence', '82%', 'of research practitioners say interpreting nuance and emotion remains a job for humans.', 'Maze, March 2026'], ['The trust problem', '50%', 'of US consumers prefer brands that do not use generative AI in customer-facing content.', 'Gartner, March 2026']] },
  { sub: 'Agents, context and the experience economy', p: 'Diana Wolosin’s “context engineers” argued that models change weekly while well-structured knowledge compounds; AJ Azzi’s MCP roundtable showed the same from the tool side. Scott Robinson asked who we design for when an agent is the one browsing and buying. The Experience Economy panel drew the conclusion: digital experiences are no longer a layer on the business, they are the business.',
    who: [['diana-wolosin.png', 'Diana Wolosin, Netflix'], ['aj-azzi.png', 'AJ Azzi, Askable'], ['scott-robinson.png', 'Scott Robinson, FreshForm'], ['ann-kostopanagiotou.jpg', 'Ann Kostopanagiotou, Mondo Robot']], photo: 'DSC09661.jpg',
    ev: [['2026 evidence', '20 → 77%', 'AI agents’ success rate on real-world tasks, up from 20% in 2025.', 'Stanford HAI, 2026 AI Index, April 2026'], ['The experience gap', '43%', 'of customers are willing to interact with a brand’s AI agent; 78% of organisations expect agents to handle half of support within 18 months.', 'Adobe, 2026 Digital Trends, February 2026']] },
];
THEME_SLIDES.forEach((t, i) => {
  const s = base(4 + i);
  eyebrow(s, TOP, 'Theme', t.sub);
  const [a, b] = THEMES[i];
  const lw = px(800), y0 = TOP + px(60);
  s.addText([{ text: a, options: { color: WHITE } }, { text: b, options: { color: Y } }], { x: M, y: y0, w: lw, h: px(260), fontFace: SANS, fontSize: 36, bold: true, charSpacing: -2, margin: 0, isTextBox: true, valign: 'top', lineSpacingMultiple: 0.98 });
  s.addText(t.p, { x: M, y: y0 + px(280), w: lw, h: px(330), fontFace: SANS, fontSize: 11.5, color: 'C4C4C4', margin: 0, isTextBox: true, valign: 'top', lineSpacingMultiple: 1.25 });
  // chips
  let cx = M, cy = y0 + px(640);
  const chipH = px(56), d = px(40);
  t.who.forEach(([f, label]) => {
    const w = px(label.length * 10.5 + 80);
    if (cx + w > M + lw) { cx = M; cy += chipH + px(10); }
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: cx, y: cy, w, h: chipH, rectRadius: chipH / 2, fill: { color: BG }, line: { color: LINE, width: 0.75 } });
    circle(s, f, cx + px(8), cy + px(8), d);
    s.addText(label, { x: cx + px(60), y: cy, w: w - px(66), h: chipH, fontFace: SANS, fontSize: 9.5, color: DIM, margin: 0, isTextBox: true, valign: 'middle' });
    cx += w + px(10);
  });
  // right column
  const rx = M + px(890), rw = W - M - rx;
  photo(s, t.photo, rx, y0, rw, px(380));
  const ey = y0 + px(414), ew = (rw - px(40)) / 2;
  evidence(s, rx, ey, ew, ...t.ev[0]);
  evidence(s, rx + ew + px(40), ey, ew, ...t.ev[1]);
});

/* ---------- 9 big quote ---------- */
{
  const s = base(9);
  eyebrow(s, TOP, 'In their words');
  s.addText('“The next era of innovation will not belong to the companies that automate the fastest, but to those that intentionally design systems in which humans and intelligent technology can learn, adapt, and make better decisions together.”', { x: M, y: px(300), w: px(1100), h: px(520), fontFace: SERIF, italic: true, fontSize: 27, color: WHITE, margin: 0, isTextBox: true, valign: 'middle', lineSpacingMultiple: 1.2 });
  circle(s, 'twisha-shah-brandenburg.png', M + px(1280), px(360), px(260));
  s.addText('Twisha Shah-Brandenburg', { x: M + px(1280), y: px(650), w: px(420), h: px(40), fontFace: SANS, fontSize: 15, bold: true, color: WHITE, margin: 0, isTextBox: true });
  s.addText('Principal, Target\nWorkshop “From Optimizing Interfaces to Designing Organizations”', { x: M + px(1280), y: px(695), w: px(420), h: px(90), fontFace: SANS, fontSize: 10, color: DIM, margin: 0, isTextBox: true, valign: 'top', lineSpacingMultiple: 1.2 });
}

/* ---------- 10–12 quote grids ---------- */
const QUOTES = [
  [['richard-ward.png', 'Richard Ward', 'Chief Design Officer, Otter.ai', "Democratized tools didn't lower the bar for great design. They raised it, because now everyone can clear the bar for good."],
   ['stephanie-mencarelli.png', 'Stephanie Mencarelli', 'Vice President of Design, Adobe', 'As AI becomes more capable and generates more machine output, human originality becomes a scarce resource. We need to protect and foster the conditions where humans can continue to create original ideas.'],
   ['kara-fitzpatrick.jpg', 'Kara Fitzpatrick', 'Director of Design & Product, Obama White House', 'While AI can read a deck, a doc, a website, your emails… it can’t read the room. AI is a tool, just like the internet was a tool. It may augment us, but it can never replace us.'],
   ['prayag-narula.png', 'Prayag Narula', 'Co-founder and CEO, Marvin', 'Research must shift from slow, episodic projects to an always-on practice that stays one step ahead, helping teams see around corners, make better decisions, and know where to invest next.']],
  [['diana-wolosin.png', 'Diana Wolosin', 'Sr. AI Systems Architect, Ads · Netflix', 'The winners of this era will be those who invest in curating and structuring their context and their knowledge, because tools change every week, but context never gets old.'],
   ['rowan-salama.png', 'Rowan Salama', 'Lead Product Researcher, Careem (Uber)', 'We can use AI and emerging technology not to remove the human from the process, but to understand human behavior better and design for the beautifully irrational ways people actually think, behave, and make decisions.'],
   ['jaleesa-chagan.png', 'Jaleesa Chagan', 'Senior UX Researcher, GAIN', 'The real competitive edge shifts from capability to conviction. Having a clear point of view, and the trust it builds, becomes the actual differentiator.'],
   ['ann-kostopanagiotou.jpg', 'Ann Kostopanagiotou', 'Managing Director, Mondo Robot', 'Digital experiences are no longer a layer on top of the business — they are the business.']],
  [['chris-gielow.png', 'Chris Gielow', 'President, UX Speakeasy', 'The future of design belongs to those who move beyond creating interfaces and toward shaping decisions.'],
   ['harriet-swan.png', 'Harriet Swan', 'Director of UX Research, GAIN', 'We are starting to think about designing webpages not just for users, but for LLMs.'],
   ['brandon-burlington.png', 'Brandon Burlington', 'Staff UX Designer, Google', 'Embracing the AI transformation is incredibly exciting because the book on how we design for it isn’t finished yet — we’re the ones actively writing the chapters.'],
   ['lauren-hughes.jpg', 'Lauren Hughes', 'Lead Product Designer, Dscout', 'Good design has always been about staying close to real people. The future of innovation is about embracing speed without losing that.']],
];
QUOTES.forEach((qs, i) => {
  const s = base(10 + i);
  eyebrow(s, TOP, 'In their words', 'Speakers on the future of innovation');
  const y0 = TOP + px(70), cw = (W - 2 * M - px(34)) / 2, ch = px(360);
  qs.forEach(([f, name, role, q], k) => {
    const x = M + (k % 2) * (cw + px(34)), y = y0 + Math.floor(k / 2) * (ch + px(34));
    s.addShape(pres.shapes.RECTANGLE, { x, y, w: cw, h: ch, fill: { color: CARD }, line: { color: LINE, width: 0.75 } });
    s.addText('“' + q + '”', { x: x + px(40), y: y + px(36), w: cw - px(80), h: px(200), fontFace: SERIF, italic: true, fontSize: 15, color: WHITE, margin: 0, isTextBox: true, valign: 'top', lineSpacingMultiple: 1.15 });
    circle(s, f, x + px(40), y + ch - px(112), px(72));
    s.addText(name, { x: x + px(130), y: y + ch - px(112), w: cw - px(170), h: px(36), fontFace: SANS, fontSize: 11, bold: true, color: WHITE, margin: 0, isTextBox: true, valign: 'bottom' });
    s.addText(role, { x: x + px(130), y: y + ch - px(76), w: cw - px(170), h: px(36), fontFace: SANS, fontSize: 9, color: DIM, margin: 0, isTextBox: true, valign: 'top' });
  });
});

/* ---------- 13 programme ---------- */
{
  const s = base(13);
  eyebrow(s, TOP, 'The day', 'Thursday 17 September 2026');
  const cols = [
    ['Guggenheim Theatre', [['10:00', 'Welcome', 'Tad Parzen (BCCA), Sebastian Gier, Scott Robinson'], ['10:15', 'The State of Humanity-Centric Innovation', 'Don Norman · fireside chat & Q&A'], ['10:45', 'Industry insights from the global DDX series', 'Sebastian Gier'], ['11:30', 'F*** Productivity: Make Space for Creative Joy', 'Stephanie Mencarelli, Adobe'], ['12:15', 'Drive, Play, Eat: designers shaping our culture', 'Panel · Scott Robinson, Elizabeth Yeongmin, Diana Tobey, Kara Fitzpatrick'], ['14:00', 'Design Isn’t Dead, It’s Just Getting Harder to Be Great', 'Richard Ward, Otter.ai'], ['14:30', 'Design Didn’t Change – AI Just Exposed What Was Missing', 'Kevin Wong, Webflow'], ['15:30', 'Closing the Design-to-Dev Gap', 'Scott Robinson, Alexander Danilowicz'], ['16:00', 'The Rise of the Experience Economy', 'Panel · Ann Kostopanagiotou, Twisha Shah-Brandenburg, Diana Wolosin, Dave Rowley'], ['17:00', 'Expectations and realities of being an enterprise design leader today', 'Panel · Lauren Hughes, Stephanie Mencarelli, Kevin Wong']]],
    ['Forum Space', [['11:30', 'The User You Designed For Doesn’t Exist', 'Rowan Salama, Careem (Uber)'], ['12:15', 'Always-On Customer Research for the AI Era', 'Prayag Narula, Marvin'], ['14:00', 'From AI Experimentation to Agentic Transformation', 'Brandon Burlington, Google'], ['14:30', 'Ship an AI feature that actually works for your users', 'Harriet Swan (GAIN), Katie Niles (TaylorMade)'], ['15:30', 'Stop Designing Screens. Start Designing Decisions.', 'Chris Gielow · workshop']]],
    ['Executive Space', [['11:30', 'Service Design / Service Jam', 'Brian LeDuc, Grace Rieger LeDuc · workshop'], ['15:30', 'Playful Career Experiments for Thriving in the AI Era', 'Diana Tobey, IDEO · workshop'], ['17:00', 'From Optimizing Interfaces to Designing Organizations', 'Twisha Shah-Brandenburg, Target · workshop'], ['18:00', 'Happy Hour', 'Hosted by Marvin & Dscout · Lobby']]],
    ['Table talks in the breaks', [['', 'Your Customer Is Not Human', 'Scott Robinson, FreshForm · 11:00'], ['', 'The Rise of Context Engineers', 'Diana Wolosin, Netflix · 11:00'], ['', 'AI: Enhancing Experiences or Adding Complexity?', 'Brianna Koch, SDXD · 11:00'], ['', 'Are Relationships Part of the Work?', 'Alex Diener, Dexcom · 15:00'], ['', 'AI Interviews: Hype or Here to Stay?', 'Prayag Narula, Marvin · 15:00'], ['', 'Shipping AI features that work for users and business', 'Jaleesa Chagan, GAIN · 15:00'], ['', 'Play Your Hand: an MCP Roundtable', 'AJ Azzi, Askable · 16:30'], ['', 'Human-Centered Design in the Age of AI', 'Jordi Morillo, Insulet · 16:30'], ['', 'AI Moves Fast. Research Still Has to Lead.', 'Bhavik Gandecha, Dscout · 16:30']]],
  ];
  const gap = px(44), widths = [1.3, 1, 1, 1], total = W - 2 * M - 3 * gap, unit = total / 4.3;
  let x = M;
  cols.forEach(([title, rows], ci) => {
    const cw = unit * widths[ci], y0 = TOP + px(60);
    s.addText(title.toUpperCase(), { x, y: y0, w: cw, h: px(30), fontFace: SANS, fontSize: 8, bold: true, color: Y, charSpacing: 3, margin: 0, isTextBox: true, valign: 'bottom' });
    line(s, x, y0 + px(40), cw);
    let y = y0 + px(52);
    const tw = ci < 3 ? px(64) : 0;
    rows.forEach(([t, b, sub]) => {
      const twoLine = b.length > (ci === 0 ? 46 : 32);
      const rh = px(twoLine ? 84 : 66);
      if (t) s.addText(t, { x, y, w: tw, h: px(24), fontFace: SANS, fontSize: 8, color: DIM, margin: 0, isTextBox: true });
      s.addText([{ text: b + '\n', options: { bold: true, color: WHITE, fontSize: 8.5 } }, { text: sub, options: { color: DIM, fontSize: 7.5 } }], { x: x + tw, y, w: cw - tw, h: rh - px(6), fontFace: SANS, margin: 0, isTextBox: true, valign: 'top', lineSpacingMultiple: 1.1 });
      y += rh;
      s.addShape(pres.shapes.LINE, { x, y: y - px(4), w: cw, h: 0, line: { color: LINE_SOFT, width: 0.5 } });
    });
    x += cw + gap;
  });
}

/* ---------- 14 photos ---------- */
{
  const s = base(14);
  eyebrow(s, TOP, 'The room', 'Photos by Desmond Chua and Nicole Gosé');
  const y0 = TOP + px(70), gw = W - 2 * M, g = px(12), rh = px(384), cu = (gw - 5 * g) / 6;
  const tiles = [['DSC09601.jpg', 3, 'Drive, Play, Eat panel'], ['DSC09472.jpg', 2, 'Park & Market atrium'], ['DSC09562.jpg', 1, 'Rowan Salama'], ['DSC09649.jpg', 2, 'Table talks in the breaks'], ['DSC09551.jpg', 2, 'Don Norman signing'], ['DSC09732.jpg', 2, 'Happy hour by Marvin & Dscout']];
  let x = M, y = y0, used = 0;
  tiles.forEach(([f, span, cap]) => {
    if (used + span > 6) { used = 0; x = M; y += rh + g; }
    const w = cu * span + g * (span - 1);
    photo(s, f, x, y, w, rh);
    s.addShape(pres.shapes.RECTANGLE, { x, y: y + rh - px(60), w, h: px(60), fill: { color: '000000', transparency: 45 }, line: { color: '000000', width: 0 } });
    s.addText(cap, { x: x + px(18), y: y + rh - px(56), w: w - px(36), h: px(50), fontFace: SANS, fontSize: 8.5, color: WHITE, margin: 0, isTextBox: true, valign: 'middle' });
    x += w + g; used += span;
  });
}

/* ---------- 15 partners + next ---------- */
{
  const s = base(15);
  eyebrow(s, TOP, 'Made possible by');
  const logos = [['freshform.png', 614, 128, 40], ['bcca.png', 345, 181, 66], ['sd-design-week.png', 248, 243, 66], ['dscout.png', 476, 91, 32], ['gain.png', 549, 78, 32], ['marvin.png', 460, 126, 40], ['askable.png', 419, 96, 32], ['mondo-robot.png', 596, 237, 40]];
  let x = M;
  for (const [f, nw, nh, th] of logos) { const h = px(th), w = h * nw / nh; s.addImage({ path: IMG + f, x, y: TOP + px(90) - h / 2, w, h }); x += w + px(60); }
  const chips = ['Don Norman Design Award', 'J.S. Held', 'Edenspiekermann', 'Appsfactory', 'San Diego Design Week', 'Friends of Figma San Diego', 'SDXD', 'UX Speakeasy'];
  x = M; let cy2 = TOP + px(170);
  chips.forEach((c) => { const w = px(c.length * 11.5 + 44); if (x + w > W - M) { x = M; cy2 += px(62); } s.addShape(pres.shapes.RECTANGLE, { x, y: cy2, w, h: px(50), fill: { color: BG }, line: { color: LINE, width: 0.75 } }); s.addText(c, { x, y: cy2, w, h: px(50), align: 'center', fontFace: SANS, fontSize: 10, color: 'D9D9D9', margin: 0, isTextBox: true, valign: 'middle' }); x += w + px(12); });
  eyebrow(s, TOP + px(330), 'The series continues');
  h1(s, 'Same exchange, next city', TOP + px(370), 92);
  const cities = [['Miami', '25 September 2026'], ['London', 'Late 2026'], ['Tokyo', '12 February 2027'], ['Dubai', 'Early 2027'], ['San Diego ’27', 'Early-bird tickets: luma.com/ddx-sd27', true]];
  const cw = (W - 2 * M) / 5, cyy = px(760), ch = px(210);
  cities.forEach(([n, d, cta], i) => {
    const cx = M + i * cw;
    s.addShape(pres.shapes.RECTANGLE, { x: cx, y: cyy, w: cw, h: ch, fill: { color: cta ? Y : BG }, line: { color: cta ? Y : LINE, width: 0.75 } });
    s.addText(n.toUpperCase(), { x: cx + px(30), y: cyy + px(30), w: cw - px(50), h: px(100), fontFace: SANS, fontSize: 22, bold: true, color: cta ? '000000' : WHITE, charSpacing: -1, margin: 0, isTextBox: true, valign: 'top' });
    s.addText(d, { x: cx + px(30), y: cyy + px(135), w: cw - px(50), h: px(60), fontFace: SANS, fontSize: 9.5, color: cta ? '404000' : DIM, margin: 0, isTextBox: true, valign: 'top' });
  });
}

/* ---------- 16 sources ---------- */
{
  const s = base(16);
  eyebrow(s, TOP, 'Sources', 'All external figures published in 2026');
  const src = [
    ['Designer Fund & Foundation Capital', 'AI in Design 2026: The Inflection Point Is Here, May 2026', 'designerfund.substack.com/p/ai-in-design-2026-the-inflection'],
    ['Figma', 'Figma’s 2026 AI Report, June 2026', 'figma.com/blog/2026-ai-report'],
    ['Maze', 'The Future of User Research Report 2026, March 2026', 'maze.co/blog/future-user-research-2026'],
    ['McKinsey & QuantumBlack', 'The State of AI in 2026: On the Road to ROI, August 2026', 'mckinsey.com/capabilities/quantumblack/our-insights/the-state-of-ai'],
    ['Gartner', 'Consumer surveys on GenAI in brand content, March and June 2026', 'gartner.com/en/newsroom'],
    ['BCG', 'AI Radar 2026: As AI Investments Surge, CEOs Take the Lead, January 2026', 'bcg.com/publications/2026/as-ai-investments-surge-ceos-take-the-lead'],
    ['Stanford HAI', 'The 2026 AI Index Report, April 2026', 'hai.stanford.edu/news/inside-the-ai-index-12-takeaways-from-the-2026-report'],
    ['Adobe', '2026 AI and Digital Trends Report, February 2026', 'business.adobe.com/resources/digital-trends-report.html'],
    ['DDX', 'Programme and roles as published at ddxconference.com/sd-agenda; speaker statements from their DDX submissions.', ''],
  ];
  const cw = (W - 2 * M - px(80)) / 2, rh = px(96);
  src.forEach(([o, t, u], i) => {
    const col = i < 5 ? 0 : 1, row = i < 5 ? i : i - 5;
    const x = M + col * (cw + px(80)), y = TOP + px(80) + row * rh;
    s.addText([{ text: o + ', ', options: { bold: true, color: WHITE } }, { text: t + '\n', options: { color: 'C4C4C4' } }, { text: u, options: { color: FAINT, fontSize: 7.5 } }], { x, y, w: cw, h: rh - px(10), fontFace: SANS, fontSize: 9.5, margin: 0, isTextBox: true, valign: 'top', lineSpacingMultiple: 1.15 });
    s.addShape(pres.shapes.LINE, { x, y: y + rh - px(6), w: cw, h: 0, line: { color: LINE_SOFT, width: 0.5 } });
  });
}

pres.writeFile({ fileName: process.argv[2] || 'deck.pptx' }).then((f) => console.log('wrote', f));
