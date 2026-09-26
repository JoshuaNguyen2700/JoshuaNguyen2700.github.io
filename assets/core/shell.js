// Shared page frame: fills <header data-shell="nav"> and <footer data-shell="footer"> on any page
// that includes them. Change the nav or footer here and every page picks it up.
import { SITE } from './site.js';
import { esc } from './util.js';

const THEME_ICONS =
  '<svg class="i-moon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M20.5 14.2A8.5 8.5 0 1 1 9.8 3.5a6.6 6.6 0 0 0 10.7 10.7z"/></svg>' +
  '<svg class="i-sun" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" aria-hidden="true"><circle cx="12" cy="12" r="4"/><path d="M12 2v2M12 20v2M4.9 4.9l1.4 1.4M17.7 17.7l1.4 1.4M2 12h2M20 12h2M4.9 19.1l1.4-1.4M17.7 6.3l1.4-1.4"/></svg>';

const nav = document.querySelector('[data-shell="nav"]');
if (nav){
  nav.innerHTML = '<nav class="nav-inner" aria-label="Site">' +
    '<a class="brand" href="/">' + esc(SITE.name) + '</a>' +
    '<div class="nav-links"><a href="/#tools">Tools</a><a href="' + esc(SITE.github) + '">GitHub</a>' +
      '<button type="button" class="theme-btn" aria-label="Switch light or dark mode" title="Light / dark mode">' + THEME_ICONS + '</button></div>' +
  '</nav>';
  // Flip the current look (saved choice, or the system setting) and remember it for every page
  nav.querySelector('.theme-btn').addEventListener('click', () => {
    const root = document.documentElement;
    const current = root.dataset.theme || (matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light');
    const next = current === 'dark' ? 'light' : 'dark';
    root.dataset.theme = next;
    try { localStorage.setItem('theme', next); } catch (e) {}
  });
}

// Any <a data-shell="contact"> links to the site owner's contact page
document.querySelectorAll('a[data-shell="contact"]').forEach(a => { a.href = SITE.contact; });

const foot = document.querySelector('[data-shell="footer"]');
if (foot){
  foot.innerHTML = '<div class="foot-inner">' +
    '<span>© ' + new Date().getFullYear() + ' ' + esc(SITE.name) + '</span>' +
    '<a href="' + esc(SITE.github) + '">' + esc(SITE.github.replace(/^https?:\/\//, '')) + '</a>' +
  '</div>';
}
