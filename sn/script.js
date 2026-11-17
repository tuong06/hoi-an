(() => {
  'use strict';
  const $ = (s, r = document) => r.querySelector(s);
  const $$ = (s, r = document) => [...r.querySelectorAll(s)];
  const calm = matchMedia('(prefers-reduced-motion: reduce)').matches;

  /* ---------- In dải ảnh: chớp flash, giấy chạy ra, hoa giấy ---------- */
  const shutter = $('.shutter');
  const strip = $('#strip');
  const flash = $('.flash');
  const canvas = $('.confetti');
  const ctx = canvas.getContext('2d');
  const COLORS = ['#f2c4cd', '#17130f', '#f6efe3', '#c8a35a', '#7b2d3a'];
  let bits = [];
  let raf = 0;

  function sizeCanvas() {
    const dpr = Math.min(devicePixelRatio || 1, 2);
    canvas.width = innerWidth * dpr;
    canvas.height = innerHeight * dpr;
    ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
  }

  function burst() {
    if (calm) return;
    sizeCanvas();
    const n = Math.min(160, Math.round(innerWidth / 5));
    for (let i = 0; i < n; i++) {
      bits.push({
        x: Math.random() * innerWidth,
        y: -20 - Math.random() * innerHeight * 0.4,
        // mảnh giấy hình dải ảnh: hẹp và dài
        w: 5 + Math.random() * 5,
        h: 12 + Math.random() * 14,
        vx: -1 + Math.random() * 2,
        vy: 2 + Math.random() * 3.5,
        a: Math.random() * Math.PI,
        va: -0.12 + Math.random() * 0.24,
        c: COLORS[i % COLORS.length],
      });
    }
    if (!raf) raf = requestAnimationFrame(tick);
  }

  function tick() {
    ctx.clearRect(0, 0, innerWidth, innerHeight);
    bits = bits.filter((b) => b.y < innerHeight + 40);
    for (const b of bits) {
      b.x += b.vx + Math.sin(b.a) * 0.8;
      b.y += b.vy;
      b.a += b.va;
      ctx.save();
      ctx.translate(b.x, b.y);
      ctx.rotate(b.a);
      ctx.scale(1, Math.cos(b.a * 1.7));
      ctx.fillStyle = b.c;
      ctx.fillRect(-b.w / 2, -b.h / 2, b.w, b.h);
      ctx.restore();
    }
    raf = bits.length ? requestAnimationFrame(tick) : 0;
    if (!raf) ctx.clearRect(0, 0, innerWidth, innerHeight);
  }

  function printStrip(scroll) {
    flash.classList.remove('on');
    strip.classList.remove('printing');
    strip.hidden = false;
    void strip.offsetWidth; // khởi động lại animation
    flash.classList.add('on');
    strip.classList.add('printing');
    shutter.setAttribute('aria-expanded', 'true');
    $('.shutter-label').textContent = 'In thêm một dải nữa';
    burst();
    if (scroll) {
      $('.booth').scrollIntoView({ behavior: calm ? 'auto' : 'smooth', block: 'center' });
    }
  }

  shutter.addEventListener('click', () => printStrip(false));
  $('.again').addEventListener('click', () => printStrip(true));
  addEventListener('resize', () => { if (raf) sizeCanvas(); });

  /* ---------- Xem ảnh lớn ---------- */
  const box = $('.lightbox');
  const big = $('img', box);
  const count = $('.lb-count');
  const pins = $$('.pins button');
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

  /* ---------- Video: mỗi lúc chỉ một đoạn phát ---------- */
  const vids = $$('.film video');
  vids.forEach((v) => v.addEventListener('play', () => {
    vids.forEach((o) => { if (o !== v) o.pause(); });
  }));
  if ('IntersectionObserver' in window) {
    const io = new IntersectionObserver((es) => {
      es.forEach((e) => { if (!e.isIntersecting) e.target.pause(); });
    }, { threshold: 0.2 });
    vids.forEach((v) => io.observe(v));
  }
})();
