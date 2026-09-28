document.addEventListener('DOMContentLoaded', () => {
  const searchInput = document.getElementById('testSearch');
  const filterPills = document.querySelectorAll('.filter-pill');
  const testCards = document.querySelectorAll('.test-card');
  const noResults = document.getElementById('noResults');
  const visibleCountEl = document.getElementById('visibleCount');

  let currentCategory = 'all';
  let currentSearchQuery = '';

  function filterCards() {
    let visibleCount = 0;

    testCards.forEach(card => {
      const name = (card.getAttribute('data-name') || '').toLowerCase();
      const category = (card.getAttribute('data-category') || '').toLowerCase();
      const tags = (card.getAttribute('data-tags') || '').toLowerCase();
      const standard = (card.getAttribute('data-standard') || '').toLowerCase();
      const desc = (card.getAttribute('data-desc') || '').toLowerCase();

      const matchesSearch = !currentSearchQuery || 
        name.includes(currentSearchQuery) || 
        tags.includes(currentSearchQuery) || 
        standard.includes(currentSearchQuery) || 
        desc.includes(currentSearchQuery);

      const matchesCategory = (currentCategory === 'all') || (category === currentCategory.toLowerCase());

      if (matchesSearch && matchesCategory) {
        card.style.display = 'flex';
        visibleCount++;
      } else {
        card.style.display = 'none';
      }
    });

    if (visibleCountEl) {
      visibleCountEl.textContent = visibleCount;
    }

    if (noResults) {
      noResults.style.display = visibleCount === 0 ? 'block' : 'none';
    }
  }

  if (searchInput) {
    searchInput.addEventListener('input', (e) => {
      currentSearchQuery = e.target.value.trim().toLowerCase();
      filterCards();
    });
  }

  filterPills.forEach(pill => {
    pill.addEventListener('click', () => {
      filterPills.forEach(p => p.classList.remove('active'));
      pill.classList.add('active');
      currentCategory = pill.getAttribute('data-filter') || 'all';
      filterCards();
    });
  });
});
