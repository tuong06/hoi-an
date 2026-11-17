(() => {
  'use strict';
  const $ = (s, r = document) => r.querySelector(s);
  const $$ = (s, r = document) => [...r.querySelectorAll(s)];

  /* ---------- Ảnh lộ dần khi cuộn tới ---------- */
  const plates = $$('.pl');
  if ('IntersectionObserver' in window) {
    document.documentElement.classList.add('js');
    const io = new IntersectionObserver((es) => {
      es.forEach((e) => {
        if (e.isIntersecting) { e.target.classList.add('in'); io.unobserve(e.target); }
      });
    }, { rootMargin: '0px 0px -8% 0px' });
    plates.forEach((p) => io.observe(p));
  }

  /* ---------- Xem ảnh lớn ---------- */
  const box = $('.lightbox');
  const big = $('img', box);
  const count = $('.lb-count');
  const pins = $$('.pl button');
  let at = 0;

  function show(i) {
    at = (i + pins.length) % pins.length;
    const thumb = $('img', pins[at]);
    big.src = thumb.getAttribute('src').replace('-s.jpg', '.jpg');
    big.alt = thumb.alt;
    count.textContent = `${at + 1} / ${pins.length}`;
  }

  pins.forEach((b, i) => b.addEventListener('click', () => { show(i); box.showModal(); }));
  $('.lb-prev').addEventListener('click', () => show(at - 1));
  $('.lb-next').addEventListener('click', () => show(at + 1));
  $('.lb-close').addEventListener('click', () => box.close());
  box.addEventListener('click', (e) => { if (e.target === box || e.target === big) box.close(); });
  box.addEventListener('keydown', (e) => {
    if (e.key === 'ArrowLeft') show(at - 1);
    if (e.key === 'ArrowRight') show(at + 1);
  });
  box.addEventListener('close', () => pins[at].focus());

  let x0 = null;
  box.addEventListener('touchstart', (e) => { x0 = e.touches[0].clientX; }, { passive: true });
  box.addEventListener('touchend', (e) => {
    if (x0 === null) return;
    const dx = e.changedTouches[0].clientX - x0;
    if (Math.abs(dx) > 50) show(at + (dx < 0 ? 1 : -1));
    x0 = null;
  }, { passive: true });

  /* ---------- Phim: chọn đoạn nào thì phát đoạn đó ---------- */
  const player = $('#player');
  const clips = $$('.clips button');
  clips.forEach((b) => b.addEventListener('click', () => {
    clips.forEach((o) => o.setAttribute('aria-pressed', String(o === b)));
    const n = b.dataset.v;
    player.poster = `assets/video/v${n}.jpg`;
    player.src = `assets/video/v${n}.mp4`;
    player.setAttribute('aria-label', $('.clip-t', b).textContent);
    player.play().catch(() => {});
  }));
  player.addEventListener('ended', () => {
    const i = clips.findIndex((b) => b.getAttribute('aria-pressed') === 'true');
    if (i < clips.length - 1) clips[i + 1].click();
  });
  if ('IntersectionObserver' in window) {
    new IntersectionObserver((es) => {
      es.forEach((e) => { if (!e.isIntersecting) player.pause(); });
    }, { threshold: 0.15 }).observe(player);
  }
})();
