const fs = require('fs');
const content = <!DOCTYPE html>
<html lang="ar">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>منصة الأعضاء</title>
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link href="https://fonts.googleapis.com/css2?family=Cairo:wght@300;400;600;700&display=swap" rel="stylesheet">
  <link rel="stylesheet" href="styles.css" />
</head>
<body>
  <div id="login-page" class="login-container">
    <div class="login-box">
      <div class="login-header">
        <div class="logo-circle"></div>
        <h1>منصة الأعضاء</h1>
        <p>لوحة الأعضاء والتدريبات</p>
      </div>

      <form id="member-login-form" onsubmit="handleMemberLogin(event)">
        <div class="form-group">
          <label for="member-name">اسم العضو</label>
          <input
            type="text"
            id="member-name"
            placeholder="أدخل اسم العضو"
            required
          />
        </div>

        <div class="form-group">
          <label for="member-password">الرقم السري</label>
          <input
            type="password"
            id="member-password"
            placeholder="أدخل الرقم السري"
            required
          />
        </div>

        <button type="submit" class="login-button">دخول</button>
      </form>

      <div class="login-footer">
        <p>أو <a href="admin.html" class="admin-link">دخول الإدارة</a></p>
      </div>
    </div>

    <div class="login-accent"></div>
  </div>

  <div id="main-page" class="app-shell" style="display: none;">
    <header class="main-header">
      <div class="brand">
        <div class="logo-placeholder"></div>
        <h1>منصة الأعضاء</h1>
      </div>
    </header>

    <main id="main-content" class="main-content"></main>
    <button class="logout-button" onclick="handleMemberLogout()">تسجيل الخروج</button>
  </div>

  <script src="https://cdn.sheetjs.com/xlsx-0.19.2/package/dist/xlsx.full.min.js"></script>
  <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
  <script src="app.js"></script>
  <script>
    initSite();
  </script>
</body>
</html>;
fs.writeFileSync('d:/ai web kpis/index.html', content, 'utf8');
