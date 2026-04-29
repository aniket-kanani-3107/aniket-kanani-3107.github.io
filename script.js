const navToggle = document.querySelector('.nav-toggle');
const navLinks = document.querySelector('.nav-links');

if (navToggle && navLinks) {
  navToggle.addEventListener('click', () => {
    const isOpen = navLinks.classList.toggle('open');
    navToggle.setAttribute('aria-expanded', String(isOpen));
  });
}

const tabs = document.querySelectorAll('.tab');
const panels = document.querySelectorAll('.auth-panel');

tabs.forEach((tab) => {
  tab.addEventListener('click', () => {
    tabs.forEach((btn) => btn.classList.remove('active'));
    panels.forEach((panel) => panel.classList.remove('active'));

    tab.classList.add('active');
    const target = tab.dataset.panel;
    const activePanel = document.getElementById(target);
    if (activePanel) {
      activePanel.classList.add('active');
    }
  });
});

const year = document.getElementById('year');
if (year) {
  year.textContent = new Date().getFullYear();
}
