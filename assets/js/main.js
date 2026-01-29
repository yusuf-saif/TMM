/* 
  The Market Masters (TMM) - main.js
  - Navbar behavior (glass + shrink)
  - Active nav link based on data-page attr
  - Reveal-on-scroll animation
  - Events: single source of truth (Home preview + events.html grid)
  - Events filtering + render (events.html)
  - Home: events preview render (index.html)
  - Home: dynamic headline rotating word
  - Back-to-top button
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
      const isActive = link.dataset.nav === pageKey;
      link.classList.toggle("active", isActive);
      if (isActive) link.setAttribute("aria-current", "page");
      else link.removeAttribute("aria-current");
    });
  }

  // ---------- Reveal on scroll ----------
  // NOTE: We keep this + add a small helper to reveal injected content too.
  const revealEls = $$(".reveal");
  let revealObserver = null;

  const revealNow = (root = document) => {
    // For injected HTML (events cards), we force reveal.
    $$(".reveal", root).forEach((el) => el.classList.add("is-visible"));
  };

  if ("IntersectionObserver" in window && revealEls.length) {
    revealObserver = new IntersectionObserver(
      (entries, obs) => {
        entries.forEach((e) => {
          if (!e.isIntersecting) return;
          e.target.classList.add("is-visible");
          obs.unobserve(e.target);
        });
      },
      { threshold: 0.12 }
    );
    revealEls.forEach((el) => revealObserver.observe(el));
  } else {
    revealNow(document);
  }

  // ---------- Back to top ----------
  const backToTop = $("#backToTop");
  const toggleBackToTop = () => {
    if (!backToTop) return;
    backToTop.style.display = window.scrollY > 450 ? "inline-flex" : "none";
  };
  window.addEventListener("scroll", toggleBackToTop, { passive: true });
  toggleBackToTop();

  // ============================================================
  // EVENTS: Single source of truth
  // - Used on events.html (grid + filters)
  // - Used on index.html (3-card preview)
  // ============================================================

  /**
   * Tip: Add "status" so we can show Upcoming/Past if you want later.
   * For now we keep your original fields to avoid breaking anything.
   */
  const eventsData = [
    {
      title: "TMM Summit",
      date: "Annual",
      location: "Maiduguri, Borno",
      tag: "Conference",
      desc:
        "An annual conference bringing together industry leaders, innovators, and stakeholders to discuss trends and strategies."
    },
    {
      title: "Funfair Activation",
      date: "Periodic",
      location: "Maiduguri, Borno",
      tag: "Experiential",
      desc:
        "High-energy brand experiences—music, games, pop-ups—designed to create memorable brand moments."
    },
    {
      title: "Social Media Exclusives",
      date: "Always-on",
      location: "Digital",
      tag: "Digital",
      desc:
        "Online-first campaigns that keep brands top-of-mind and drive engagement across social platforms."
    },
    {
      title: "TMM Webinar Series",
      date: "Always-on",
      location: "Hybrid",
      tag: "Webinar",
      desc:
        "Live sessions and practical learning for business owners, creators, and brand builders."
    }
  ];

  // Expose globally (optional, useful for debugging or future pages)
  window.TMM_EVENTS = eventsData;

  // ---------- Shared card renderer ----------
  const eventCard = (ev) => `
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
  `;

  // ---------- Home page: Events preview (must exist: #homeEventsGrid) ----------
  const homeEventsGrid = $("#homeEventsGrid");
  const renderHomeEvents = () => {
    if (!homeEventsGrid) return;

    // Strategy:
    // - If you later add status, we can show Upcoming first.
    // - For now: show first 3 items.
    const top = eventsData.slice(0, 3);

    homeEventsGrid.innerHTML = top.map(eventCard).join("");

    // Since these are injected, mark them visible immediately.
    revealNow(homeEventsGrid);
  };

  // ---------- Events page: render + filter ----------
  // Only runs if we find #eventsGrid on the page.
  const eventsGrid = $("#eventsGrid");
  const filterWrap = $("#eventFilters");

  const renderEvents = (items) => {
    if (!eventsGrid) return;

    eventsGrid.innerHTML = items.map(eventCard).join("");

    // injected => reveal immediately
    revealNow(eventsGrid);
  };

  const uniqueTags = () => {
    // Normalize tags a bit (avoid duplicates due to typos/case differences)
    const tags = new Set(eventsData.map((e) => (e.tag || "").trim()));
    return ["All", ...tags].filter(Boolean);
  };

  const renderFilters = () => {
    if (!filterWrap) return;

    const tags = uniqueTags();
    filterWrap.innerHTML = tags
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
      const filtered =
        tag === "All" ? eventsData : eventsData.filter((x) => x.tag === tag);

      renderEvents(filtered);
    });
  };

  if (eventsGrid) {
    renderFilters();
    renderEvents(eventsData);
  }

  // ============================================================
  // HOME: Dynamic headline rotating phrase
  // Requirement: "Ideas that move markets" -> rotate 3 phrases
  // ============================================================
  const initHeroRotator = () => {
    // Works with your existing HTML if you add:
    // <span id="heroRotateWord" data-rotate='["move markets","build trust","drive growth"]'>move markets</span>
    const el = $("#heroRotateWord");
    if (!el) return;

    let items = [];
    try {
      items = JSON.parse(el.getAttribute("data-rotate") || "[]");
    } catch (err) {
      items = [];
    }
    if (!items.length) return;

    const interval = parseInt(el.getAttribute("data-interval") || "2400", 10);
    let index = 0;

    const swap = () => {
      index = (index + 1) % items.length;

      // smooth fade down/up (CSS class already used in my suggestion)
      el.classList.add("is-changing");
      setTimeout(() => {
        el.textContent = items[index];
        el.classList.remove("is-changing");
      }, 220);
    };

    setInterval(swap, interval);
  };

  // Run page-level renders
  renderHomeEvents();
  initHeroRotator();

})();
