// Applies the visitor's saved light/dark choice before the page draws, so there is no flash.
// Every page loads it in <head> as a plain script:  <script src="/assets/core/theme.js"></script>
// With no saved choice, the page follows the system setting. The toggle button lives in shell.js.
(function(){
  try {
    var t = localStorage.getItem('theme');
    if (t === 'light' || t === 'dark') document.documentElement.setAttribute('data-theme', t);
  } catch (e) {}
})();
