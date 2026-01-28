// assets/js/includes.js
/*
  Loads shared partials (navbar/footer) into pages.
  This avoids duplicating markup across HTML files.
*/

async function injectPartial(selector, url) {
  const mount = document.querySelector(selector);
  if (!mount) return;

  try {
    const res = await fetch(url, { cache: "no-cache" });
    mount.innerHTML = await res.text();
  } catch (err) {
    console.error("Partial load failed:", url, err);
  }
}

(async function () {
  await injectPartial("#navbarMount", "partials/navbar.html");
  await injectPartial("#footerMount", "partials/footer.html");
})();
