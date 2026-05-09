const toggleActive = (selector) => {
  const items = document.querySelectorAll(selector);
  items.forEach((item) => {
    item.addEventListener('click', () => {
      items.forEach((el) => el.classList.remove('active'));
      item.classList.add('active');
    });
  });
};

toggleActive('.nav-item');
toggleActive('.chip');
