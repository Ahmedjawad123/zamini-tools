// Smooth scroll
document.querySelectorAll('a[href^="#"]').forEach(anchor => {
  anchor.addEventListener('click', function(e) {
    e.preventDefault();
    const target = document.querySelector(this.getAttribute('href'));
    if(target) target.scrollIntoView({ behavior: 'smooth' });
  });
});

// Tabs functionality
const tabs = document.querySelectorAll('.tab');
const contents = document.querySelectorAll('.tab-content');
tabs.forEach(tab => {
  tab.addEventListener('click', () => {
    tabs.forEach(t => t.classList.remove('active'));
    tab.classList.add('active');
    const target = document.getElementById(tab.dataset.tab);
    contents.forEach(c => c.classList.remove('active'));
    if(target) target.classList.add('active');
  });
});

// Chat toggle
const chatBtn = document.getElementById('chatBtn');
const chatBox = document.getElementById('chatBox');
chatBtn.addEventListener('click', () => {
  chatBox.style.display = chatBox.style.display === 'block' ? 'none' : 'block';
});

// Forms alert
document.querySelectorAll('form').forEach(f => {
  f.addEventListener('submit', e => {
    e.preventDefault();
    alert('Thank you! Form submitted.');
    f.reset();
  });
});
