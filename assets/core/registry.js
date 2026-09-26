// Reads the tool registry (/assets/tools.json): the one list of tools the whole site uses.
// Bad entries are skipped with a console warning, so one mistake can't take down the page.

const ID = /^[a-z0-9][a-z0-9-]*$/;

// Every tool lives at /tools/<id>/
export const toolUrl = t => '/tools/' + t.id + '/';

export async function loadTools(){
  try {
    const res = await fetch('/assets/tools.json', {cache:'no-cache'});
    if (!res.ok) throw new Error('HTTP ' + res.status);
    const data = await res.json();
    return (data.tools || []).filter((t, i) => {
      const ok = t && typeof t.id === 'string' && ID.test(t.id) && typeof t.name === 'string' && t.name.trim();
      if (!ok) console.warn('tools.json: skipping invalid entry ' + (i + 1), t);
      return ok;
    });
  } catch (e){
    console.error('Could not load the tool registry', e);
    return [];
  }
}
