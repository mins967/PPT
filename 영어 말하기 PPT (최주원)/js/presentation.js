/**
 * Portugal PPT — Presentation Controller
 * Morph-style transitions + keyboard navigation
 */

(function () {
  'use strict';

  const slides = document.querySelectorAll('.slide');
  const totalSlides = slides.length;
  let currentSlide = 0;
  let isTransitioning = false;

  const prevBtn = document.getElementById('prevBtn');
  const nextBtn = document.getElementById('nextBtn');
  const navDots = document.getElementById('navDots');
  const navCounter = document.getElementById('navCounter');
  const progressBar = document.getElementById('progressBar');

  /* ─── Init ─── */
  function init() {
    buildDots();
    updateUI();
    bindEvents();
    resetAnimations(slides[0]);
  }

  function buildDots() {
    navDots.innerHTML = '';
    for (let i = 0; i < totalSlides; i++) {
      const dot = document.createElement('button');
      dot.className = 'nav-dot' + (i === 0 ? ' active' : '');
      dot.setAttribute('aria-label', `Slide ${i + 1}`);
      dot.addEventListener('click', () => goToSlide(i));
      navDots.appendChild(dot);
    }
  }

  function bindEvents() {
    prevBtn.addEventListener('click', () => goToSlide(currentSlide - 1));
    nextBtn.addEventListener('click', () => goToSlide(currentSlide + 1));

    document.addEventListener('keydown', (e) => {
      if (e.key === 'ArrowRight' || e.key === ' ' || e.key === 'PageDown') {
        e.preventDefault();
        goToSlide(currentSlide + 1);
      } else if (e.key === 'ArrowLeft' || e.key === 'PageUp') {
        e.preventDefault();
        goToSlide(currentSlide - 1);
      } else if (e.key === 'Home') {
        goToSlide(0);
      } else if (e.key === 'End') {
        goToSlide(totalSlides - 1);
      }
    });

    /* Touch swipe */
    let touchStartX = 0;
    document.addEventListener('touchstart', (e) => {
      touchStartX = e.changedTouches[0].screenX;
    }, { passive: true });

    document.addEventListener('touchend', (e) => {
      const diff = touchStartX - e.changedTouches[0].screenX;
      if (Math.abs(diff) > 50) {
        goToSlide(diff > 0 ? currentSlide + 1 : currentSlide - 1);
      }
    }, { passive: true });
  }

  /* ─── Navigation ─── */
  function goToSlide(index) {
    if (isTransitioning) return;
    if (index < 0 || index >= totalSlides || index === currentSlide) return;

    isTransitioning = true;
    const prev = slides[currentSlide];
    const next = slides[index];
    const direction = index > currentSlide ? 1 : -1;

    /* Morph-style exit */
    prev.classList.remove('active');
    prev.classList.add('exiting');

    resetAnimations(prev);

    /* Prepare next slide */
    next.classList.add('entering');
    next.style.transform = direction > 0 ? 'scale(1.04)' : 'scale(0.96)';

    requestAnimationFrame(() => {
      next.classList.add('active');
      next.style.transform = 'scale(1)';
    });

    setTimeout(() => {
      prev.classList.remove('exiting');
      next.classList.remove('entering');
      next.style.transform = '';
      resetAnimations(next);
      triggerMorphElements(next);
      currentSlide = index;
      updateUI();
      isTransitioning = false;
    }, 850);
  }

  /* Reset animation states when leaving a slide */
  function resetAnimations(slide) {
    slide.querySelectorAll('[data-animate]').forEach((el) => {
      el.style.transition = 'none';
      el.style.opacity = '0';
      el.style.transform = '';
      void el.offsetHeight;
      el.style.transition = '';
    });
  }

  /* Re-trigger morph shared elements */
  function triggerMorphElements(slide) {
    slide.querySelectorAll('.morph-target').forEach((el) => {
      el.style.animation = 'none';
      void el.offsetHeight;
      el.style.animation = '';
    });
  }

  function updateUI() {
    const num = String(currentSlide + 1).padStart(2, '0');
    const total = String(totalSlides).padStart(2, '0');
    navCounter.textContent = `${num} / ${total}`;

    prevBtn.disabled = currentSlide === 0;
    nextBtn.disabled = currentSlide === totalSlides - 1;

    progressBar.style.width = `${((currentSlide + 1) / totalSlides) * 100}%`;

    navDots.querySelectorAll('.nav-dot').forEach((dot, i) => {
      dot.classList.toggle('active', i === currentSlide);
    });
  }

  init();
})();
