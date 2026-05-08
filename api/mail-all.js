const Imap = require('node-imap');
const simpleParser = require("mailparser").simpleParser;

function generateEmailsHtml(emailsData) {
  const escapeHtml = (str) => str
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');

  const renderEmailItem = (email, index) => {
    const { send, subject, text, html: emailHtml, date } = email;
    const escapedText = escapeHtml(text || '无内容');
    const escapedHtml = emailHtml || `<p>${escapedText.replace(/\n/g, '<br>')}</p>`;
    const formattedDate = new Date(date).toLocaleString() || '未知日期';

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

  const emailsHtml = emailsData.map((email, index) => renderEmailItem(email, index)).join('');

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
          .email-meta { display: flex; gap: 20px; color: #4a5568; font-size: 0.9em; }
          .toggle-btn { margin-left: auto; color: #4299e1; font-weight: 500; }
          .email-content { padding: 20px; color: #1a202c; line-height: 1.8; }
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

// IMAP 获取Token 专用Scope
async function get_imap_token(refresh_token, client_id) {
    const response = await fetch('https://login.microsoftonline.com/consumers/oauth2/v2.0/token', {
        method: 'POST',
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

    if (!response.ok) throw new Error('IMAP token get fail');
    const data = await response.json();
    return data.access_token;
}

const generateAuthString = (user, accessToken) => {
    const authString = `user=${user}\x01auth=Bearer ${accessToken}\x01\x01`;
    return Buffer.from(authString).toString('base64');
}

// 1. 新GR：User.Read Mail.Read offline_access
async function get_new_gr_token(refresh_token, client_id) {
    const response = await fetch('https://login.microsoftonline.com/consumers/oauth2/v2.0/token', {
        method: 'POST',
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
    if (!response.ok) return null;
    const data = await response.json();
    return data.access_token ? data.access_token : null;
}

// 2. 老GR：https://graph.microsoft.com/.default
async function get_old_gr_token(refresh_token, client_id) {
    const response = await fetch('https://login.microsoftonline.com/consumers/oauth2/v2.0/token', {
        method: 'POST',
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
    if (!response.ok) return null;
    const data = await response.json();
    return data.access_token ? data.access_token : null;
}

// Graph 拉取邮件
async function get_emails(access_token, mailbox) {
    if (!access_token) return null;
    try {
        const response = await fetch(`https://graph.microsoft.com/v1.0/me/mailFolders/${mailbox}/messages?$top=10000`, {
            method: 'GET',
            headers: {
                "Authorization": `Bearer ${access_token}`
            },
        });
        if (!response.ok) return null;
        const responseData = await response.json();
        const emails = responseData.value;
        if (!emails) return null;

        return emails.map(item => ({
            send: item.from?.emailAddress?.address || '未知',
            subject: item.subject || '无主题',
            text: item.bodyPreview || '',
            html: item.body?.content || '',
            date: item.createdDateTime,
        }));
    } catch (error) {
        return null;
    }
}

module.exports = async (req, res) => {
    const { password } = req.method === 'GET' ? req.query : req.body;
    const expectedPassword = process.env.PASSWORD;

    if (password !== expectedPassword && expectedPassword) {
        return res.status(401).json({
            error: 'Authentication failed'
        });
    }

    const params = req.method === 'GET' ? req.query : req.body;
    let { refresh_token, client_id, email, mailbox, response_type = 'json' } = params;

    if (!refresh_token || !client_id || !email || !mailbox) {
        return res.status(400).json({ error: 'Missing required parameters' });
    }
    if (!['json', 'html'].includes(response_type)) {
        return res.status(400).json({ error: 'Invalid response_type' });
    }

    // 文件夹兼容
    let mb = mailbox;
    if (mb === 'INBOX') mb = 'inbox';
    if (mb === 'Junk') mb = 'junkemail';

    try {
        // 第一步：尝试 新GR
        let token = await get_new_gr_token(refresh_token, client_id);
        if (token) {
            let list = await get_emails(token, mb);
            if (list) {
                if (response_type === 'html') {
                    return res.send(generateEmailsHtml(list));
                } else {
                    return res.json(list);
                }
            }
        }

        // 第二步：尝试 老GR
        token = await get_old_gr_token(refresh_token, client_id);
        if (token) {
            let list = await get_emails(token, mb);
            if (list) {
                if (response_type === 'html') {
                    return res.send(generateEmailsHtml(list));
                } else {
                    return res.json(list);
                }
            }
        }

        // 第三步：全部失败 走 IMAP 兜底
        const imapToken = await get_imap_token(refresh_token, client_id);
        const authString = generateAuthString(email, imapToken);
        const emailList = [];

        const imap = new Imap({
            user: email,
            xoauth2: authString,
            host: 'outlook.office365.com',
            port: 993,
            tls: true,
            tlsOptions: { rejectUnauthorized: false }
        });

        imap.once("ready", async () => {
            try {
                await new Promise((resolve, reject) => {
                    imap.openBox(mailbox, true, (err) => err ? reject(err) : resolve());
                });
                const results = await new Promise((resolve, reject) => {
                    imap.search(["ALL"], (err, r) => err ? reject(err) : resolve(r));
                });

                const f = imap.fetch(results, { bodies: "" });
                f.on("message", (msg) => {
                    msg.on("body", (stream) => {
                        simpleParser(stream, (err, mail) => {
                            if (!err) {
                                emailList.push({
                                    send: mail.from.text,
                                    subject: mail.subject,
                                    text: mail.text,
                                    html: mail.html,
                                    date: mail.date,
                                });
                            }
                        });
                    });
                });
                f.once("end", () => imap.end());
            } catch (err) {
                imap.end();
                res.status(500).json({ error: err.message });
            }
        });

        imap.once('error', (err) => {
            res.status(500).json({ error: err.message });
        });

        imap.once('end', () => {
            if (response_type === 'html') {
                res.send(generateEmailsHtml(emailList));
            } else {
                res.json(emailList);
            }
        });

        imap.connect();

    } catch (error) {
        res.status(500).json({ error: error.message });
    }
};
