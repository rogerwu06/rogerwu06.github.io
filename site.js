document.addEventListener('DOMContentLoaded', () => {
  const page = document.body.dataset.page;
  document.querySelectorAll('.nav a[data-nav]').forEach(a => {
    if (a.dataset.nav === page) a.classList.add('active');
  });

  // Keep the public site focused on the projects rather than resume-style profile text.
  document.querySelectorAll('.role, .institution, .profile-focus, .profile-links .github, .hello').forEach(el => el.remove());

  document.querySelectorAll('.profile-links .email').forEach(a => {
    a.href = 'mailto:wurogerwu0@gmail.com';
    a.textContent = 'wurogerwu0@gmail.com';
  });
});
