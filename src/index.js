const puppeteer = require('puppeteer');
const { exportExcel } = require('./excel');

function sleep(interval) {
  return new Promise((resolve) => {
    setTimeout(resolve, interval);
  });
}

const main = async () => {
  const browser = await puppeteer.launch({
    // executablePath: '/usr/bin/chromium-browser',
    // args: ['--no-sandbox', '--headless', '--disable-gpu'],
    headless: false // 是否使用无头浏览器
  });

  const page = await browser.newPage();
  // 设置 UA
  await page.setUserAgent(
    'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/105.0.0.0 Safari/537.36'
  );

  await page.goto(
    'https://work.weixin.qq.com/wework_admin/loginpage_wx?from=myhome'
  );
  // 扫码登录
  await page.waitForNavigation({ timeout: 1e9 });
  // 判断是否需要短信验证码
  if (!(await page.$('#menu_contacts'))) {
    console.log('扫码成功，等待短信验证码');
    await page.waitForNavigation({ timeout: 1e9 });
  }
  console.log('登录成功');

  await page.click('#menu_contacts');
  console.log('进入通讯录页面');
  console.log('获取所有部门 id');
  await page.waitForSelector('.member_colRight', { timeout: 1e9 });
  await page.waitForTimeout(200);
  // 获取部门 id
  // 展开所有部门，以及子部门
  for (;;) {
    const anchors = await page.$$(
      '.member_colLeft_bottom .jstree-1 .jstree-closed'
    );
    if (anchors.length === 0) {
      break;
    }
    // eslint-disable-next-line no-restricted-syntax
    for (const anchor of anchors) {
      await anchor.click();
      await anchor.click();
    }
    await sleep(300);
  }
  const departsArr = await page.$$eval(
    '.member_colLeft_bottom .jstree-1 .jstree-anchor',
    (anchors) => {
      return anchors.map((anchor) => [
        anchor.innerText,
        anchor.getAttribute('id').replace('_anchor', '')
      ]);
    }
  );
  const departs = Object.fromEntries(departsArr);
  console.log(`获取所有 ${departsArr.length} 部门 id 完成`);
  // // 判断页数
  const pages = await page.evaluate(() => {
    // eslint-disable-next-line no-undef
    const text = document.querySelector('.ww_pageNav_info_text')?.innerText;
    return Number((text || '1/1').split('/')?.[1]) || 1;
  });
  console.log(`用户数据共计 ${pages} 页`);
  // 循环查找用户
  const users = [];
  for (let p = 1; p <= pages; p += 1) {
    // 用户表格遍历
    let memberListElm = await page.$$('.member_colRight_memberTable tbody tr');
    // 当前页用户数
    const membersOnPage = memberListElm.length;
    console.log(`当前页 ${p} 用户数： ${membersOnPage}`);
    for (let index = 0; index < membersOnPage; index += 1) {
      // 逐个进入用户详情页
      memberListElm = await page.$$('.member_colRight_memberTable tbody tr');
      await memberListElm[index].click();
      await page.waitForSelector('.member_colRight_view', { timeout: 1e9 });
      // 处理用户数据
      const user = {};
      // externalId
      const userIdElm = await page.$(
        '.member_display_cover_detail_bottom:last-child'
      );
      const externalId = await page.evaluate(
        (elm) => (elm?.innerText || '').replace('帐号：', ''),
        userIdElm
      );
      if (externalId) {
        user.externalId = externalId;
      }
      // 性别
      const womanElm = await page.$('.ww_commonImg_Woman');
      user.gender = womanElm ? 'F' : 'M';
      // 姓名
      const nameElm = await page.$('.member_display_cover_detail_name');
      const name = await page.evaluate((elm) => elm?.innerText, nameElm);
      if (name) {
        user.name = name;
      }
      // 通用字段处理
      const fields = ['Phone', 'Email'];
      // eslint-disable-next-line no-restricted-syntax
      for (const field of fields) {
        const Elm = await page.$(
          `.member_display_item_${field} .member_display_item_right`
        );
        const value = await page.evaluate((elm) => elm?.innerText, Elm);
        if (value && value !== '未设置') {
          user[field.toLowerCase()] = value;
        }
      }
      // 部门
      // eslint-disable-next-line no-loop-func
      const departments = await page.evaluate(() => {
        // eslint-disable-next-line no-undef
        const elm = document.querySelectorAll('.member_display_item');
        for (let i = 0; i < elm.length; i += 1) {
          if ((elm[i].innerText || '').includes('部门')) {
            const text = elm[i].innerText.replace('部门：\t', '').trim();
            return text.replaceAll(' / ', '/');
          }
        }
      });
      if (departments) {
        user.departments = departments;
      }
      // 部门负责人
      // eslint-disable-next-line no-loop-func
      const department_leader = await page.evaluate(() => {
        // eslint-disable-next-line no-undef
        const elm = document.querySelectorAll('.member_display_item');
        for (let i = 0; i < elm.length; i += 1) {
          if ((elm[i].innerText || '').includes('部门负责人')) {
            return (elm[i].innerText || '').includes('是');
          }
        }
        return false;
      });
      user.department_leader = department_leader;

      users.push(user);
      // 返回
      await page.click('.ww_btn_Back');
      await page.waitForSelector('.member_colRight', { timeout: 1e9 });
      await page.waitForTimeout(200);
    }
    // 下一页
    if (p < pages) {
      console.log('进入下一页');
      await page.click('.js_next_page');
      await page.waitForSelector('.member_colRight', { timeout: 1e9 });
      await page.waitForTimeout(200);
    }
  }

  exportExcel(users, departs);
  process.exit(0);
};

main().then(console.log).catch(console.error);
