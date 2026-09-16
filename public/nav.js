async function renderNav(activePage) {
  const res = await fetch('/api/auth/me');
  if (res.status === 401) { window.location.href = '/login.html'; return null; }
  const user = await res.json();

  const links = [
    { href: 'index.html', label: 'หน้าแรก' },
    { href: 'time-correction.html', label: 'ขอเพิ่มเวลา' },
    { href: 'ot.html', label: 'ขอ OT' },
    { href: 'leave.html', label: 'ขอลางาน' },
  ];
  if (user.role === 'admin') {
    links.push({ href: 'history-leave.html', label: 'ประวัติวันลา' });
    links.push({ href: 'history-time-correction.html', label: 'ประวัติขอเพิ่มเวลา' });
    links.push({ href: 'history-ot.html', label: 'ประวัติ OT' });
  }

  const linksHtml = links.map(l =>
    `<a href="${l.href}" class="${activePage === l.href ? 'active' : ''}">${l.label}</a>`
  ).join('');

  const header = document.createElement('div');
  header.className = 'header';
  header.innerHTML = `
    <div class="brand">ระบบลา / OT</div>
    <div class="links">${linksHtml}</div>
    <div class="user">
      <span>${user.name} ${user.role === 'admin' ? '(Admin)' : ''}</span>
      <button class="logout" onclick="logout()">ออกจากระบบ</button>
    </div>
  `;
  document.body.insertBefore(header, document.body.firstChild);
  return user;
}

async function logout() {
  await fetch('/api/auth/logout', { method: 'POST' });
  window.location.href = '/login.html';
}