// Renders tool tiles from the registry. Each tile is built on its own,
// so one broken entry is skipped instead of breaking the whole grid.
import { esc } from './util.js';
import { ICONS } from './icons.js';
import { toolUrl } from './registry.js';

const STATUS = {
  live: {eyebrow:'Available now', cta:'Open'},
  soon: {eyebrow:'Coming soon', cta:'Learn more'}
};

function iconEl(t){
  // A tool's own icon file, drawn through a CSS mask so it follows light/dark mode
  if (/\.svg$/i.test(t.icon || '')){
    const el = document.createElement('span');
    el.className = 'icon icon-mask';
    el.setAttribute('aria-hidden', 'true');
    el.style.setProperty('--icon', 'url(' + JSON.stringify(toolUrl(t) + t.icon) + ')');
    return el;
  }
  if (t.icon && !ICONS[t.icon]) console.warn('Unknown icon "' + t.icon + '" for tool "' + t.id + '"');
  const tpl = document.createElement('template');
  tpl.innerHTML = '<svg class="icon" viewBox="0 0 64 64" fill="none" stroke="currentColor" aria-hidden="true">' + (ICONS[t.icon] || ICONS.plus) + '</svg>';
  return tpl.content.firstChild;
}

function tile(t){
  const status = STATUS[t.status] ? t.status : 'soon';
  const s = STATUS[status];
  const a = document.createElement('a');
  a.className = 'tile ' + status;
  a.href = toolUrl(t);
  a.innerHTML = '<div class="eyebrow">' + s.eyebrow + '</div>' +
    '<h2>' + esc(t.name) + '</h2>' +
    (t.blurb ? '<p>' + esc(t.blurb) + '</p>' : '') +
    '<span class="cta">' + s.cta + '</span>';
  a.prepend(iconEl(t));
  return a;
}

export function renderTiles(container, tools){
  container.replaceChildren();
  for (const t of tools){
    try { container.append(tile(t)); }
    catch (e){ console.error('Could not show the tile for "' + t.id + '"', e); }
  }
  if (!container.children.length) container.innerHTML = '<p class="empty">No tools to show right now.</p>';
}
