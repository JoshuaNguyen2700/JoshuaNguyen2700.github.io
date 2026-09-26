// Small helpers any page can import. Add new ones freely; don't change what existing ones return.

// Escape text before putting it into HTML.
export const esc = s => String(s == null ? '' : s).replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
