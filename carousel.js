(() => {
  const items = window.featuredImages;
  if (!Array.isArray(items) || items.length === 0) return;

  const section = document.querySelector(".featured-work");
  const link = document.getElementById("featured-link");
  const image = document.getElementById("featured-image");
  const caption = document.getElementById("featured-caption");
  const position = document.getElementById("carousel-position");
  const previous = document.getElementById("carousel-prev");
  const next = document.getElementById("carousel-next");
  if (!section || !link || !image || !caption || !position || !previous || !next) return;

  let current = 0;
  let timer;
  const reducedMotion = window.matchMedia("(prefers-reduced-motion: reduce)");

  function show(index) {
    current = (index + items.length) % items.length;
    const item = items[current];
    image.src = item.image;
    image.alt = item.alt;
    link.href = item.paper;
    caption.textContent = `${item.caption}. Click the image to view it in the paper.`;
    position.textContent = `${current + 1} of ${items.length}`;
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

  previous.addEventListener("click", () => { show(current - 1); start(); });
  next.addEventListener("click", () => { show(current + 1); start(); });
  section.addEventListener("mouseenter", stop);
  section.addEventListener("mouseleave", start);
  section.addEventListener("focusin", stop);
  section.addEventListener("focusout", () => window.setTimeout(start, 0));
  document.addEventListener("visibilitychange", start);
  reducedMotion.addEventListener("change", start);

  show(0);
  start();
})();
