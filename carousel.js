(() => {
  if (!Array.isArray(window.featuredImages) || window.featuredImages.length === 0) return;
  const items = [...window.featuredImages];
  for (let i = items.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [items[i], items[j]] = [items[j], items[i]];
  }

  const section = document.querySelector(".featured-work");
  const link = document.getElementById("featured-link");
  const image = document.getElementById("featured-image");
  const dots = document.getElementById("carousel-dots");
  if (!section || !link || !image || !dots) return;

  let current = 0;
  let timer;
  const reducedMotion = window.matchMedia("(prefers-reduced-motion: reduce)");
  const dotButtons = items.map((item, index) => {
    const button = document.createElement("button");
    button.type = "button";
    button.className = "carousel-dot";
    button.setAttribute("aria-label", `Show ${item.caption}`);
    button.addEventListener("click", () => { show(index); start(); });
    dots.append(button);
    return button;
  });

  function show(index) {
    current = (index + items.length) % items.length;
    const item = items[current];
    image.src = item.image;
    image.alt = item.alt;
    link.href = item.paper;
    dotButtons.forEach((button, dotIndex) => {
      button.setAttribute("aria-pressed", String(dotIndex === current));
    });
  }

  function stop() {
    window.clearInterval(timer);
    timer = undefined;
  }

  function start() {
    stop();
    if (!reducedMotion.matches && !document.hidden && !section.matches(":hover") && !section.contains(document.activeElement)) {
      timer = window.setInterval(() => show(current + 1), 7000);
    }
  }

  section.addEventListener("mouseenter", stop);
  section.addEventListener("mouseleave", start);
  section.addEventListener("focusin", stop);
  section.addEventListener("focusout", () => window.setTimeout(start, 0));
  document.addEventListener("visibilitychange", start);
  reducedMotion.addEventListener("change", start);

  show(0);
  start();
})();
