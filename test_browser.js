const puppeteer = require('puppeteer');

(async () => {
  const browser = await puppeteer.launch({ headless: true, args: ['--no-sandbox'] });
  const page = await browser.newPage();
  
  page.on('console', msg => console.log('BROWSER CONSOLE:', msg.text()));
  page.on('pageerror', error => console.log('BROWSER ERROR:', error.message));
  
  // Set localStorage to bypass auth
  await page.goto('http://localhost:3000/outsource_manage.html');
  await page.evaluate(() => {
    localStorage.setItem('currentUser', JSON.stringify({ userCode: 'wangjian', role: 'admin' }));
  });
  
  await page.goto('http://localhost:3000/outsource_manage.html');
  await page.waitForTimeout(2000);
  
  await browser.close();
})();
