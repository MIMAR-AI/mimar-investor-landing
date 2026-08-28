const reducedMotion = window.matchMedia('(prefers-reduced-motion: reduce)').matches;

if (!reducedMotion) {
  document.documentElement.classList.add('motion-ready');

  const revealObserver = new IntersectionObserver((entries, observer) => {
    entries.forEach((entry) => {
      if (!entry.isIntersecting) return;
      entry.target.classList.add('is-visible');
      observer.unobserve(entry.target);
    });
  }, { threshold: 0.12, rootMargin: '0px 0px -40px' });

  document.querySelectorAll('[data-reveal]').forEach((element) => revealObserver.observe(element));

  const counterObserver = new IntersectionObserver((entries, observer) => {
    entries.forEach((entry) => {
      if (!entry.isIntersecting) return;
      const element = entry.target;
      const target = Number(element.dataset.count);
      const prefix = element.dataset.prefix ?? '';
      const suffix = element.dataset.suffix ?? '';
      const decimals = Number(element.dataset.decimals ?? 0);
      const numberFormatter = new Intl.NumberFormat('en-US', { minimumFractionDigits: decimals, maximumFractionDigits: decimals });
      const startedAt = performance.now();
      const duration = 850;

      const animate = (now) => {
        const progress = Math.min((now - startedAt) / duration, 1);
        const eased = 1 - Math.pow(1 - progress, 3);
        const value = decimals ? target * eased : Math.round(target * eased);
        element.textContent = `${prefix}${numberFormatter.format(value)}${suffix}`;
        if (progress < 1) requestAnimationFrame(animate);
      };

      requestAnimationFrame(animate);
      observer.unobserve(element);
    });
  }, { threshold: 0.65 });

  document.querySelectorAll('[data-count]').forEach((element) => counterObserver.observe(element));
}
