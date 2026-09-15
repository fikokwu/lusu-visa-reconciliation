// Builds visa-rec-training-slideshow.html: a full-screen web player for the
// same slides as the design canvas. The .dc.html files and canvas.json stay the
// single source; re-run `node build-slideshow.mjs` after editing a slide.
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const dir = path.dirname(fileURLToPath(import.meta.url));
const canvas = JSON.parse(fs.readFileSync(path.join(dir, 'canvas.json'), 'utf8'));

// Each slide's <helmet> styles are scoped to its own section, because slides
// reuse class names (.th, .td, .pin ...) with different values.
function scopeCss(css, scope) {
  return css.replace(/([^{}]+)\{([^{}]*)\}/g, (m, selectors, decls) =>
    selectors.split(',').map((s) => s.trim()).filter(Boolean)
      .map((s) => (s === 'body' ? scope : scope + ' ' + s)).join(', ') + ' {' + decls + '}');
}

const slides = canvas.artboards.map((board, i) => {
  const src = fs.readFileSync(path.join(dir, board.file), 'utf8');
  const css = (src.match(/<helmet>\s*<style>([\s\S]*?)<\/style>\s*<\/helmet>/) || [, ''])[1];
  const body = src.match(/<\/helmet>([\s\S]*?)<\/x-dc>/);
  if (!body) throw new Error('No slide markup in ' + board.file);
  const id = 's' + (i + 1);
  return { id, title: board.title.replace(/^\d+\s+/, ''), css: scopeCss(css, '#' + id), html: body[1].trim() };
});

const total = slides.length;
const sections = slides.map((s, i) =>
  `<section class="slide" id="${s.id}" data-title="${s.title.replace(/"/g, '&quot;')}" aria-roledescription="slide" aria-label="${i + 1} of ${total}: ${s.title.replace(/"/g, '&quot;')}"${i ? ' hidden' : ''}>\n${s.html}\n</section>`
).join('\n');

const page = `<title>Visa Rec Training Slideshow</title>
<style>
  /* Single deliberate look: a dark stage that frames the light app-styled slides. */
  :root {
    --stage: #0B1830;
    --bar: #102447;
    --ink: #E6EEF8;
    --muted: #8FA3BF;
    --accent: #1CA0E3;
    --control: rgba(230, 238, 248, .08);
    --control-hover: rgba(230, 238, 248, .16);
  }
  html, body { height: 100%; }
  body { margin: 0; background: var(--stage); color: var(--ink); font-family: system-ui, -apple-system, "Segoe UI", Roboto, sans-serif; overflow: hidden; }
  .player { height: 100dvh; min-height: 420px; display: flex; flex-direction: column; }
  .stage { flex: 1; min-height: 0; display: flex; align-items: center; justify-content: center; padding: 16px; cursor: pointer; }
  .sizer { position: relative; flex-shrink: 0; box-shadow: 0 18px 60px rgba(0, 0, 0, .45); border-radius: 6px; overflow: hidden; background: #F4F6F9; }
  .deck { position: absolute; top: 0; left: 0; width: 1280px; height: 720px; transform-origin: 0 0; }
  .slide { position: absolute; inset: 0; }
  .slide.enter { animation: fade .22s ease-out; }
  @keyframes fade { from { opacity: 0; } to { opacity: 1; } }
  @media (prefers-reduced-motion: reduce) { .slide.enter { animation: none; } }
  .progress { height: 3px; background: var(--control); }
  .progress-fill { height: 100%; width: 0; background: var(--accent); transition: width .22s ease-out; }
  .bar { display: flex; align-items: center; gap: 10px; flex-wrap: wrap; padding: 10px 16px; background: var(--bar); font-size: 14px; transition: opacity .3s; }
  .bar.idle { opacity: 0; pointer-events: none; }
  .bar button, .bar select { font: inherit; color: var(--ink); background: var(--control); border: 0; border-radius: 8px; cursor: pointer; }
  .bar button { width: 40px; height: 40px; display: inline-flex; align-items: center; justify-content: center; }
  .bar select { height: 40px; padding: 0 10px; }
  .bar button:hover, .bar select:hover { background: var(--control-hover); }
  .bar button:focus-visible, .bar select:focus-visible, .bar input:focus-visible { outline: 2px solid var(--accent); outline-offset: 2px; }
  .bar button:disabled { opacity: .35; cursor: default; }
  .count { min-width: 64px; text-align: center; font-variant-numeric: tabular-nums; font-weight: 600; }
  .slide-title { flex: 1; min-width: 120px; color: var(--muted); white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
  .auto { display: inline-flex; align-items: center; gap: 8px; color: var(--ink); cursor: pointer; }
  .auto input { width: 18px; height: 18px; accent-color: var(--accent); }
${slides.map((s) => s.css).join('\n')}
</style>

<main class="player">
  <div class="stage" id="stage" title="Click to go forward">
    <div class="sizer" id="sizer">
      <div class="deck" id="deck">
${sections}
      </div>
    </div>
  </div>
  <div class="progress" aria-hidden="true"><div class="progress-fill" id="progressFill"></div></div>
  <nav class="bar" id="bar" aria-label="Slideshow controls">
    <button type="button" id="prev" aria-label="Previous slide"><svg width="20" height="20" viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12.5 4.5L7 10l5.5 5.5"/></svg></button>
    <span class="count" id="count" aria-live="polite">1 / ${total}</span>
    <button type="button" id="next" aria-label="Next slide"><svg width="20" height="20" viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M7.5 4.5L13 10l-5.5 5.5"/></svg></button>
    <span class="slide-title" id="slideTitle">${slides[0].title}</span>
    <label class="auto" for="autoplay"><input type="checkbox" id="autoplay"> Auto-advance</label>
    <select id="interval" aria-label="Seconds per slide">
      <option value="8">8 s</option>
      <option value="15" selected>15 s</option>
      <option value="30">30 s</option>
      <option value="60">60 s</option>
    </select>
    <button type="button" id="fullscreen" aria-label="Full screen"><svg width="20" height="20" viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M3.5 7.5v-4h4M16.5 7.5v-4h-4M3.5 12.5v4h4M16.5 12.5v4h-4"/></svg></button>
  </nav>
</main>

<script>
(function () {
  var slides = Array.prototype.slice.call(document.querySelectorAll('.slide'));
  var total = slides.length;
  var stage = document.getElementById('stage');
  var sizer = document.getElementById('sizer');
  var deck = document.getElementById('deck');
  var bar = document.getElementById('bar');
  var autoBox = document.getElementById('autoplay');
  var intervalSel = document.getElementById('interval');
  var fsBtn = document.getElementById('fullscreen');
  var current = 0;
  var timer = null;
  var idleTimer = null;

  function fit() {
    var style = getComputedStyle(stage);
    var w = stage.clientWidth - parseFloat(style.paddingLeft) - parseFloat(style.paddingRight);
    var h = stage.clientHeight - parseFloat(style.paddingTop) - parseFloat(style.paddingBottom);
    var scale = Math.max(0.1, Math.min(w / 1280, h / 720));
    sizer.style.width = Math.floor(1280 * scale) + 'px';
    sizer.style.height = Math.floor(720 * scale) + 'px';
    deck.style.transform = 'scale(' + scale + ')';
  }

  function show(index) {
    index = Math.max(0, Math.min(total - 1, index));
    if (index !== current) {
      slides[current].hidden = true;
      slides[current].classList.remove('enter');
    }
    current = index;
    var slide = slides[current];
    slide.hidden = false;
    slide.classList.remove('enter');
    void slide.offsetWidth;
    slide.classList.add('enter');
    document.getElementById('count').textContent = (current + 1) + ' / ' + total;
    document.getElementById('slideTitle').textContent = slide.getAttribute('data-title');
    document.getElementById('progressFill').style.width = ((current + 1) / total * 100) + '%';
    document.getElementById('prev').disabled = current === 0;
    document.getElementById('next').disabled = current === total - 1;
    try { history.replaceState(null, '', '#' + (current + 1)); } catch (e) {}
    if (autoBox.checked) schedule();
  }

  // Auto-advance plays through once and stops on the last slide.
  function schedule() {
    clearTimeout(timer);
    if (!autoBox.checked) return;
    if (current === total - 1) { autoBox.checked = false; return; }
    timer = setTimeout(function () { show(current + 1); }, Number(intervalSel.value) * 1000);
  }

  function next() { show(current + 1); }
  function prev() { show(current - 1); }

  document.getElementById('next').addEventListener('click', next);
  document.getElementById('prev').addEventListener('click', prev);
  stage.addEventListener('click', function (e) {
    var rect = stage.getBoundingClientRect();
    if (e.clientX - rect.left < rect.width / 3) prev(); else next();
  });
  autoBox.addEventListener('change', schedule);
  intervalSel.addEventListener('change', schedule);

  document.addEventListener('keydown', function (e) {
    var tag = e.target && e.target.tagName;
    if (tag === 'SELECT' || tag === 'INPUT') return;
    if (e.key === 'ArrowRight' || e.key === 'ArrowDown' || e.key === 'PageDown' || e.key === ' ' || e.key === 'Enter') { e.preventDefault(); next(); }
    else if (e.key === 'ArrowLeft' || e.key === 'ArrowUp' || e.key === 'PageUp' || e.key === 'Backspace') { e.preventDefault(); prev(); }
    else if (e.key === 'Home') { e.preventDefault(); show(0); }
    else if (e.key === 'End') { e.preventDefault(); show(total - 1); }
    else if (e.key === 'f' || e.key === 'F') { toggleFullscreen(); }
  });

  function toggleFullscreen() {
    if (!document.fullscreenEnabled) return;
    if (document.fullscreenElement) document.exitFullscreen();
    else document.documentElement.requestFullscreen().catch(function () {});
  }
  if (document.fullscreenEnabled) fsBtn.addEventListener('click', toggleFullscreen);
  else fsBtn.hidden = true;

  // In full screen the controls fade out until the mouse moves.
  function wake() {
    bar.classList.remove('idle');
    clearTimeout(idleTimer);
    if (document.fullscreenElement) idleTimer = setTimeout(function () { bar.classList.add('idle'); }, 2500);
  }
  document.addEventListener('mousemove', wake);
  document.addEventListener('fullscreenchange', function () { wake(); fit(); });

  window.addEventListener('resize', fit);
  fit();
  var start = parseInt((location.hash || '').slice(1), 10);
  show(start >= 1 && start <= total ? start - 1 : 0);
})();
</script>
`;

const out = path.join(dir, 'visa-rec-training-slideshow.html');
fs.writeFileSync(out, page);
console.log('wrote ' + path.basename(out) + ' — ' + total + ' slides, ' + Math.round(page.length / 1024) + ' KB');
