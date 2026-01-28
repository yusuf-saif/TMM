/* 
  The Market Masters (TMM) - main.js
  - Navbar behavior (glass + shrink)
  - Active nav link based on data-page attr
  - Reveal-on-scroll animation
  - Events filtering + render (events.html)
  - Back-to-top button

  Keep it vanilla: fewer dependencies, easier maintenance.
*/

(function () {
  "use strict";

  // ---------- Helpers ----------
  const $ = (sel, root = document) => root.querySelector(sel);
  const $$ = (sel, root = document) => [...root.querySelectorAll(sel)];

  // ---------- Navbar: glass on load + shrink on scroll ----------
  const navbar = $(".navbar");
  const applyNavbarState = () => {
    if (!navbar) return;
    const scrolled = window.scrollY > 10;
    navbar.classList.toggle("glass", true); // always glass for a modern look
    navbar.classList.toggle("py-3", !scrolled);
    navbar.classList.toggle("py-2", scrolled);
  };
  window.addEventListener("scroll", applyNavbarState, { passive: true });
  applyNavbarState();

  // ---------- Active nav ----------
  // Each page sets <body data-page="home">, etc.
  const pageKey = document.body?.dataset?.page;
  if (pageKey) {
    $$("[data-nav]").forEach((link) => {
      link.classList.toggle("active", link.dataset.nav === pageKey);
      link.setAttribute("aria-current", link.dataset.nav === pageKey ? "page" : "false");
    });
  }

  // ---------- Reveal on scroll ----------
  const revealEls = $$(".reveal");
  if ("IntersectionObserver" in window && revealEls.length) {
    const io = new IntersectionObserver(
      (entries, obs) => {
        entries.forEach((e) => {
          if (!e.isIntersecting) return;
          e.target.classList.add("is-visible");
          obs.unobserve(e.target);
        });
      },
      { threshold: 0.12 }
    );
    revealEls.forEach((el) => io.observe(el));
  } else {
    revealEls.forEach((el) => el.classList.add("is-visible"));
  }

  // ---------- Back to top ----------
  const backToTop = $("#backToTop");
  const toggleBackToTop = () => {
    if (!backToTop) return;
    backToTop.style.display = window.scrollY > 450 ? "inline-flex" : "none";
  };
  window.addEventListener("scroll", toggleBackToTop, { passive: true });
  toggleBackToTop();

  // ---------- Events page: render + filter ----------
  // Only runs if we find #eventsGrid on the page.
  const eventsGrid = $("#eventsGrid");
  const filterWrap = $("#eventFilters");

  const eventsData = [
    {
      title: "Ramadan Souq",
      date: "Seasonal",
      location: "Maiduguri, Borno",
      tag: "Community",
      desc:
        "A curated marketplace experience that brings vendors, families, and communities together during Ramadan.",
    },
    {
      title: "End-of-Year Market Sales",
      date: "Q4 (Annual)",
      location: "Maiduguri, Borno",
      tag: "Marketplace",
      desc:
        "A sales-driven activation to boost visibility for vendors and help customers discover standout brands.",
    },
    {
      title: "Funfair Activation",
      date: "Periodic",
      location: "Maiduguri, Borno",
      tag: "Experiential",
      desc:
        "High-energy brand experiences—music, games, pop-ups—designed to create memorable brand moments.",
    },
    {
      title: "Social Media Exclusives",
      date: "Always-on",
      location: "Digital",
      tag: "Digital",
      desc:
        "Online-first campaigns that keep brands top-of-mind and drive engagement across social platforms.",
    },
  ];

  const renderEvents = (items) => {
    if (!eventsGrid) return;
    eventsGrid.innerHTML = items
      .map(
        (ev) => `
        <div class="col-md-6 col-lg-4 reveal">
          <div class="card card--lift h-100">
            <div class="card-body p-4">
              <div class="d-flex justify-content-between align-items-start gap-3">
                <div>
                  <span class="tag mb-3">${ev.tag}</span>
                  <h3 class="h5 mb-2">${ev.title}</h3>
                  <p class="text-muted-2 mb-3">${ev.desc}</p>
                </div>
                <div class="icon-chip flex-shrink-0" aria-hidden="true">
                  <i class="bi bi-calendar-event"></i>
                </div>
              </div>
              <div class="d-flex flex-wrap gap-2 text-muted-2 small">
                <span><i class="bi bi-clock me-1"></i>${ev.date}</span>
                <span class="mx-1">•</span>
                <span><i class="bi bi-geo-alt me-1"></i>${ev.location}</span>
              </div>
            </div>
          </div>
        </div>
      `
      )
      .join("");

    // Re-attach reveal observer for newly injected elements
    $$(".reveal", eventsGrid).forEach((el) => el.classList.add("is-visible"));
  };

  const uniqueTags = () => {
    const tags = new Set(eventsData.map((e) => e.tag));
    return ["All", ...tags];
  };

  const renderFilters = () => {
    if (!filterWrap) return;
    filterWrap.innerHTML = uniqueTags()
      .map(
        (tag, idx) => `
        <button 
          type="button"
          class="btn ${idx === 0 ? "btn-primary" : "btn-outline-primary"} btn-sm"
          data-filter="${tag}">
          ${tag}
        </button>`
      )
      .join("");

    // Click handling
    filterWrap.addEventListener("click", (e) => {
      const btn = e.target.closest("[data-filter]");
      if (!btn) return;

      // Update active style
      $$("[data-filter]", filterWrap).forEach((b) => {
        b.classList.toggle("btn-primary", b === btn);
        b.classList.toggle("btn-outline-primary", b !== btn);
      });

      const tag = btn.dataset.filter;
      const filtered = tag === "All" ? eventsData : eventsData.filter((x) => x.tag === tag);
      renderEvents(filtered);
    });
  };

  if (eventsGrid) {
    renderFilters();
    renderEvents(eventsData);
  }

})();
