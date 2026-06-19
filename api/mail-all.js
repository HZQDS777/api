const Imap = require('node-imap');
const { simpleParser } = require("mailparser");
const { setTimeout } = require('timers/promises');

function generateEmailsHtml(emailsData) {
  const escapeHtml = (str) => {
    if (!str) return '';
    return String(str)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#039;');
  };

  const renderEmailItem = (email, index) => {
    const { send, subject, text, html: emailHtml, date } = email;
    const escapedText = escapeHtml(text || '无内容');
    const escapedHtml = emailHtml || `<p>${escapedText.replace(/\n/g, '<br>')}</p>`;
    let formattedDate = '未知日期';
    if (date) {
      try {
        formattedDate = new Date(date).toLocaleString('zh-CN');
      } catch (e) {}
    }

    return `
      <div class="email-item" id="email-${index}">
        <div class="email-header" onclick="toggleEmailContent(${index})">
          <h3 class="email-subject">${escapeHtml(subject || '无主题')}</h3>
          <div class="email-meta">
            <span>发件人：${escapeHtml(send || '未知')}</span>
            <span>日期：${formattedDate}</span>
            <span class="toggle-btn">${index === 0 ? '收起' : '展开'}</span>
          </div>
        </div>
        <div class="email-content" id="content-${index}" style="${index === 0 ? 'display:block' : 'display:none'}">
          ${escapedHtml || `<p>${escapedText}</p>`}
        </div>
      </div>
    `;
  };

  const emailsHtml = Array.isArray(emailsData)
    ? emailsData.map((email, index) => renderEmailItem(email, index)).join('')
    : '';

  return `
    <!DOCTYPE html>
    <html lang="zh-CN">
      <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>邮件列表 - 共${emailsData.length}封</title>
        <style>
          * { margin: 0; padding: 0; box-sizing: border-box; }
          body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; line-height: 1.6; background: #f8f9fa; padding: 20px; }
          .page-title { text-align: center; color: #2d3748; margin-bottom: 30px; font-size: 1.8em; }
          .email-list { max-width: 1000px; margin: 0 auto; gap: 15px; display: flex; flex-direction: column; }
          .email-item { background: white; border-radius: 8px; box-shadow: 0 2px 8px rgba(0,0,0,0.08); overflow: hidden; }
          .email-header { padding: 15px 20px; background: #f5fafe; cursor: pointer; border-bottom: 1px solid #eee; }
          .email-subject { color: #2d3748; margin-bottom: 8px; font-size: 1.1em; }
          .email-meta { display: flex; gap: 20px; color: #4a5568; font-size: 0.9em; flex-wrap: wrap; }
          .toggle-btn { margin-left: auto; color: #4299e1; font-weight: 500; }
          .email-content { padding: 20px; color: #1a202c; line-height: 1.8; word-break: break-all; }
          .email-content p { margin-bottom: 10px; }
          .email-content img { max-width: 100%; height: auto; }
          @media (max-width: 768px) {
            .email-meta { flex-direction: column; gap: 5px; }
            .toggle-btn { margin-left: 0; margin-top: 5px; }
          }
        </style>
        <script>
          function toggleEmailContent(index) {
            const content = document.getElementById(\`content-\${index}\`);
            const btn = document.querySelector(\`#email-\${index} .toggle-btn\`);
            if (!content || !btn) return;
            if (content.style.display === 'none') {
              content.style.display = 'block';
              btn.textContent = '收起';
            } else {
              content.style.display = 'none';
              btn.textContent = '展开';
            }
          }
        </script>
      </head>
      <body>
        <h1 class="page-title">邮件列表（共${emailsData.length}封）</h1>
        <div class="email-list">
          ${emailsHtml || '<div style="text-align:center; padding:30px; color:#718096;">未获取到邮件</div>'}
        </div>
      </body>
    </html>
  `;
}

// IMAP Token
async function get_imap_token(refresh_token, client_id) {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), 8000);
  try {
    const response = await fetch('https://login.microsoftonline.com/consumers/oauth2/v2.0/token', {
      method: 'POST',
      signal: controller.signal,
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded'
      },
      body: new URLSearchParams({
        'client_id': client_id,
        'grant_type': 'refresh_token',
        'refresh_token': refresh_token,
        'scope': 'https://outlook.office.com/IMAP.AccessAsUser.All offline_access'
      }).toString()
    });
    clearTimeout(timer);
    if (!response.ok) throw new Error('IMAP token 获取失败');
    const data = await response.json();
    return data.access_token;
  } catch (e) {
    clearTimeout(timer);
    return null;
  }
}

const generateAuthString = (user, accessToken) => {
  const authString = `user=${user}\x01auth=Bearer ${accessToken}\x01\x01`;
  return Buffer.from(authString).toString('base64');
}

// 新版 Graph Token
async function get_new_gr_token(refresh_token, client_id) {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), 8000);
  try {
    const response = await fetch('https://login.microsoftonline.com/consumers/oauth2/v2.0/token', {
      method: 'POST',
      signal: controller.signal,
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded'
      },
      body: new URLSearchParams({
        'client_id': client_id,
        'grant_type': 'refresh_token',
        'refresh_token': refresh_token,
        'scope': 'User.Read Mail.Read offline_access'
      }).toString()
    });
    clearTimeout(timer);
    if (!response.ok) return null;
    const data = await response.json();
    return data.access_token || null;
  } catch {
    clearTimeout(timer);
    return null;
  }
}

// 旧版 Graph Token
async function get_old_gr_token(refresh_token, client_id) {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), 8000);
  try {
    const response = await fetch('https://login.microsoftonline.com/consumers/oauth2/v2.0/token', {
      method: 'POST',
      signal: controller.signal,
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded'
      },
      body: new URLSearchParams({
        'client_id': client_id,
        'grant_type': 'refresh_token',
        'refresh_token': refresh_token,
        'scope': 'https://graph.microsoft.com/.default'
      }).toString()
    });
    clearTimeout(timer);
    if (!response.ok) return null;
    const data = await response.json();
    return data.access_token || null;
  } catch {
    clearTimeout(timer);
    return null;
  }
}

// Graph 获取邮件（修复top上限100）
async function get_emails(access_token, mailbox) {
  if (!access_token) return null;
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), 10000);
  try {
    const response = await fetch(`https://graph.microsoft.com/v1.0/me/mailFolders/${mailbox}/messages?$top=100`, {
      method: 'GET',
      signal: controller.signal,
      headers: {
        "Authorization": `Bearer ${access_token}`
      },
    });
    clearTimeout(timer);
    if (!response.ok) return null;
    const responseData = await response.json();
    const emails = responseData.value;
    if (!Array.isArray(emails)) return null;

    return emails.map(item => ({
      send: item.from?.emailAddress?.address || '未知',
      subject: item.subject || '无主题',
      text: item.bodyPreview || '',
      html: item.body?.content || '',
      date: item.createdDateTime,
    }));
  } catch (error) {
    clearTimeout(timer);
    return null;
  }
}

// 封装 IMAP 获取邮件 Promise 化，解决回调异步混乱
async function fetchImapMails(emailAddr, mailboxName, xoauthBase64) {
  return new Promise((resolve, reject) => {
    const emailList = [];
    let finishedCount = 0;
    let totalMsgCount = 0;
    let isDone = false;

    const imap = new Imap({
      user: emailAddr,
      xoauth2: xoauthBase64,
      host: 'outlook.office365.com',
      port: 993,
      tls: true,
      tlsOptions: { rejectUnauthorized: false }
    });

    function safeClose(err) {
      if (isDone) return;
      isDone = true;
      imap.end();
      if (err) reject(err);
      else resolve(emailList);
    }

    imap.once('error', err => safeClose(err));
    imap.once('end', () => safeClose());

    imap.once("ready", () => {
      imap.openBox(mailboxName, true, (err) => {
        if (err) return safeClose(err);
        imap.search(["ALL"], (err, results) => {
          if (err) return safeClose(err);
          totalMsgCount = results.length;
          if (totalMsgCount === 0) return safeClose();

          const fetchStream = imap.fetch(results, { bodies: "" });
          fetchStream.on("message", (msg) => {
            msg.on("body", (stream) => {
              simpleParser(stream, (err, mail) => {
                if (!err) {
                  emailList.push({
                    send: mail.from?.text || '未知',
                    subject: mail.subject || '无主题',
                    text: mail.text || '',
                    html: mail.html || '',
                    date: mail.date || new Date()
                  });
                }
                finishedCount++;
                if (finishedCount >= totalMsgCount) {
                  safeClose();
                }
              });
            });
          });
          fetchStream.once("error", err => safeClose(err));
          fetchStream.once("end", () => {
            // 防止解析慢，超时兜底3秒
            setTimeout(() => safeClose(), 3000);
          });
        });
      });
    });
    imap.connect();
  });
}

module.exports = async (req, res) => {
  // 防重复返回
  let responded = false;
  const sendJson = (code, data) => {
    if (responded) return;
    responded = true;
    res.status(code).json(data);
  };
  const sendHtml = (html) => {
    if (responded) return;
    responded = true;
    res.send(html);
  };

  const { password } = req.method === 'GET' ? req.query : req.body;
  const expectedPassword = process.env.PASSWORD;

  if (expectedPassword && password !== expectedPassword) {
    return sendJson(401, { error: 'Authentication failed' });
  }

  const params = req.method === 'GET' ? req.query : req.body;
  let { refresh_token, client_id, email, mailbox, response_type = 'json' } = params;

  if (!refresh_token || !client_id || !email || !mailbox) {
    return sendJson(400, { error: 'Missing required parameters' });
  }
  if (!['json', 'html'].includes(response_type)) {
    return sendJson(400, { error: 'Invalid response_type' });
  }

  // 文件夹名称兼容转换
  let targetMailbox = mailbox;
  let graphMailbox = mailbox;
  if (targetMailbox === 'INBOX') targetMailbox = 'INBOX';
  if (targetMailbox === 'Junk') targetMailbox = 'Junk';
  if (graphMailbox === 'INBOX') graphMailbox = 'inbox';
  if (graphMailbox === 'Junk') graphMailbox = 'junkemail';

  try {
    // ========== 优先级调整：优先 IMAP，失败再走Graph ==========
    let mailList = null;

    // 第一步：尝试 IMAP
    const imapToken = await get_imap_token(refresh_token, client_id);
    if (imapToken) {
      const authStr = generateAuthString(email, imapToken);
      try {
        mailList = await fetchImapMails(email, targetMailbox, authStr);
      } catch (imapErr) {
        mailList = null;
      }
    }

    // IMAP 无结果再走 Graph 两套兜底
    if (!mailList || !Array.isArray(mailList)) {
      // 新Graph
      let grToken = await get_new_gr_token(refresh_token, client_id);
      if (grToken) {
        mailList = await get_emails(grToken, graphMailbox);
      }
      // 新Graph失败走旧Graph
      if (!mailList || !Array.isArray(mailList)) {
        grToken = await get_old_gr_token(refresh_token, client_id);
        if (grToken) {
          mailList = await get_emails(grToken, graphMailbox);
        }
      }
    }

    // 兜底空数组
    if (!Array.isArray(mailList)) mailList = [];

    // 返回对应格式
    if (response_type === 'html') {
      return sendHtml(generateEmailsHtml(mailList));
    } else {
      return sendJson(200, mailList);
    }

  } catch (error) {
    sendJson(500, { error: error.message });
  }
};
