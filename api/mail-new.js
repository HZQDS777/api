const Imap = require('node-imap');
const simpleParser = require("mailparser").simpleParser;

// ===================== 全局配置与工具函数 =====================
const CONFIG = {
  OAUTH_TOKEN_URL: 'https://login.microsoftonline.com/consumers/oauth2/v2.0/token',
  GRAPH_API_BASE_URL: 'https://graph.microsoft.com/v1.0/me/mailFolders',
  IMAP_CONFIG: {
    host: 'outlook.office365.com',
    port: 993,
    tls: true,
    tlsOptions: { rejectUnauthorized: false },
    connTimeout: 6000,
    authTimeout: 6000
  },
  MAILBOX_MAP: {
    '收件箱': 'inbox',
    'inbox': 'inbox',
    '已发送': 'sentitems',
    'sentitems': 'sentitems',
    '草稿': 'draft',
    'drafts': 'draft',
    '删除邮件': 'deleteditems',
    'deleteditems': 'deleteditems',
    '垃圾邮件': 'junkemail',
    'junk': 'junkemail'
  },
  REQUEST_TIMEOUT: 8000,
  SUPPORTED_METHODS: ['GET', 'POST'],
  REQUIRED_PARAMS: ['refresh_token', 'client_id', 'email', 'mailbox'],
  TARGET_FOLDERS: {
    graph: ['inbox', 'junkemail'],
    imap: ['INBOX', 'Junk'],
    chineseName: {
      'inbox': '收件箱',
      'junkemail': '垃圾箱',
      'INBOX': '收件箱',
      'Junk': '垃圾箱'
    }
  }
};

const SCOPE = {
  NEW_GR: "User.Read Mail.Read offline_access",
  OLD_GR: "https://graph.microsoft.com/.default",
  IMAP: "https://outlook.office.com/IMAP.AccessAsUser.All offline_access"
};

async function fetchWithTimeout(url, options = {}, timeout = CONFIG.REQUEST_TIMEOUT) {
  const controller = new AbortController();
  const timeoutId = setTimeout(() => controller.abort(), timeout);
  try {
    const response = await fetch(url, { ...options, signal: controller.signal });
    clearTimeout(timeoutId);
    return response;
  } catch (error) {
    clearTimeout(timeoutId);
    throw new Error(error.name === "AbortError" ? "请求超时" : error.message);
  }
}

async function getNewGrToken(refresh_token, client_id) {
  try {
    const res = await fetchWithTimeout(CONFIG.OAUTH_TOKEN_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({
        client_id,
        grant_type: 'refresh_token',
        refresh_token,
        scope: SCOPE.NEW_GR
      }).toString()
    });
    if (!res.ok) return null;
    const data = await res.json();
    return data.access_token || null;
  } catch (e) {
    return null;
  }
}

async function getOldGrToken(refresh_token, client_id) {
  try {
    const res = await fetchWithTimeout(CONFIG.OAUTH_TOKEN_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({
        client_id,
        grant_type: 'refresh_token',
        refresh_token,
        scope: SCOPE.OLD_GR
      }).toString()
    });
    if (!res.ok) return null;
    const data = await res.json();
    return data.access_token || null;
  } catch (e) {
    return null;
  }
}

async function getImapToken(refresh_token, client_id) {
  const res = await fetchWithTimeout(CONFIG.OAUTH_TOKEN_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({
      client_id,
      grant_type: 'refresh_token',
      refresh_token,
      scope: SCOPE.IMAP
    }).toString()
  });
  if (!res.ok) throw new Error("IMAP令牌获取失败");
  const data = await res.json();
  return data.access_token;
}

function escapeHtml(str) {
  if (!str) return '';
  return str
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

function escapeJson(str) {
  if (!str) return str;
  return str.replace(/\\/g, '\\\\').replace(/"/g, '\\"').replace(/\n/g, '\\n');
}

function getLatestEmail(email1, email2) {
  if (!email1) return email2;
  if (!email2) return email1;
  const time1 = new Date(email1.date).getTime() || 0;
  const time2 = new Date(email2.date).getTime() || 0;
  return time1 > time2 ? email1 : email2;
}

function validateParams(params) {
  const { email } = params;
  const emailReg = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  if (!emailReg.test(email)) return new Error("邮箱格式无效");
  if (params.refresh_token?.length < 50) return new Error("refresh_token无效");
  if (params.client_id?.length < 10) return new Error("client_id无效");
  return null;
}

function preprocessText(rawText) {
  if (!rawText) return '';
  const withoutHtml = rawText.replace(/<[^>]+>/g, ' ');
  const mergeDigitSeparators = withoutHtml.replace(/(\d)[\s-_]+(\d)/g, '$1$2');
  const normalized = mergeDigitSeparators.replace(/\s+/g, ' ').trim();
  return normalized.toLowerCase();
}

const VERIFY_CODE_RULES = [
  { regex: /(\d{6})\s*is your/i, desc: "TikTok专属格式", extractGroup: 1, confidence: 100 },
  { regex: /(verify|verification|validation|auth|security).*code\s*(\d{6})/i, desc: "英文语义6位验证码", extractGroup: 2, confidence: 100 },
  { regex: /(验证码|校验码|动态码|登录码|安全码|短信码|授权码|动态口令|登录口令)[:：\s]*([0-9]{6})/i, desc: "中文语义6位验证码", extractGroup: 2, confidence: 100 },
  { regex: /(验证码|校验码)[:：\s]*([0-9]{3}[-_][0-9]{3})/i, desc: "带分隔符6位验证码", extractGroup: 2, confidence: 95 },
  { regex: /\b[0-9]{6}\b/, desc: "纯6位数字验证码", extractGroup: 0, confidence: 90 },
  { regex: /(验证码|校验码)[:：\s]*([0-9]{4})/i, desc: "中文4位兜底", extractGroup: 2, confidence: 10 },
  { regex: /\b[0-9]{4}\b/, desc: "纯4位兜底", extractGroup: 0, confidence: 5 }
];

function extractVerifyCode(text) {
  const cleanText = preprocessText(text);
  if (!cleanText) return { code: '', rule: '无有效文本', confidence: 0 };
  const matchedResults = [];
  for (const rule of VERIFY_CODE_RULES) {
    const matches = cleanText.match(rule.regex);
    if (matches) {
      const rawCode = matches[rule.extractGroup].trim();
      const finalCode = rawCode.replace(/\D/g, '');
      if (finalCode.length === 6) {
        matchedResults.push({ code: finalCode, rule: rule.desc, confidence: rule.confidence });
      } else if (finalCode.length === 4 && rule.confidence < 90) {
        matchedResults.push({ code: finalCode, rule: rule.desc, confidence: rule.confidence });
      }
    }
  }
  if (matchedResults.length === 0) return { code: '', rule: '无匹配规则', confidence: 0 };
  const uniqueResults = Array.from(new Map(matchedResults.map(item => [item.code, item])).values());
  uniqueResults.sort((a, b) => b.confidence - a.confidence);
  return uniqueResults[0];
}

function extractVerifyCodeWithLog(text, emailSubject = '未知主题') {
  const result = extractVerifyCode(text);
  console.log(`【验证码】主题：${emailSubject} | 码：${result.code} | ${result.rule} | ${result.confidence}`);
  return result;
}

function getVerifyCodeFromEmail(emailData, emailSubject = '未知主题') {
  const targetText = emailData.text || emailData.html || '';
  return extractVerifyCodeWithLog(targetText, emailSubject);
}

function generateEmailHtml(emailData) {
  const { send, subject, text, html: emailHtml, date, folderSource, verifyCode } = emailData;
  const escapedText = escapeHtml(text || '');
  const escapedHtml = emailHtml || `<p>${escapedText.replace(/\n/g, '<br>')}</p>`;
  const folderCN = folderSource || '未知文件夹';
  const codeDisplay = verifyCode.code ? `<span style="color: #e53e3e; font-weight: bold; font-size: 1.2em;">${verifyCode.code}</span>（${verifyCode.rule}，置信度：${verifyCode.confidence}%）` : '未提取到6位验证码';

  return `<!DOCTYPE html><html lang="zh-CN"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1.0"><title>${escapeHtml(subject || '无主题邮件')}</title></head><body style="font-family:system-ui;background:#f5f5f5;margin:0;padding:20px"><div style="max-width:800px;margin:0 auto;background:white;padding:30px;border-radius:8px;box-shadow:0 2px 10px #0000001a"><div style="margin-bottom:20px;padding-bottom:15px;border-bottom:1px solid #eee"><h1 style="margin:0 0 15px;color:#2d3748">${escapeHtml(subject || '无主题')}</h1><div style="color:#4a5568;font-size:.9em"><span><strong>发件人：</strong>${escapeHtml(send || '未知发件人')}</span><span><strong>发送日期：</strong>${new Date(date).toLocaleString() || '未知日期'}</span><div style="margin-top:10px;padding:10px;background:#fef7fb;border-left:3px solid #e53e3e"><strong>提取的6位验证码：</strong>${codeDisplay}</div></div></div><div style="color:#1a202c">${escapedHtml}</div></div></body></html>`;
}

async function get_single_folder_email(access_token, mailbox) {
  try {
    const url = `${CONFIG.GRAPH_API_BASE_URL}/${mailbox}/messages?$top=1&$orderby=receivedDateTime desc`;
    const response = await fetchWithTimeout(url, {
      method: 'GET',
      headers: { Authorization: `Bearer ${access_token}` }
    });
    if (!response.ok) return null;
    const responseData = await response.json();
    const email = responseData.value?.[0];
    if (!email) return null;
    const verifyCode = getVerifyCodeFromEmail({ text: email.bodyPreview, html: email.body?.content }, email.subject);
    return {
      send: email.from?.emailAddress?.address || '未知发件人',
      subject: email.subject || '无主题',
      text: email.bodyPreview || '',
      html: email.body?.content || '',
      date: email.createdDateTime || new Date().toISOString(),
      folderSource: CONFIG.TARGET_FOLDERS.chineseName[mailbox] || '未知文件夹',
      verifyCode
    };
  } catch (error) {
    return null;
  }
}

async function get_dual_folder_latest_email_graph(access_token) {
  const [inbox, junk] = await Promise.all([
    get_single_folder_email(access_token, 'inbox'),
    get_single_folder_email(access_token, 'junkemail')
  ]);
  return getLatestEmail(inbox, junk);
}

async function get_dual_folder_latest_email_imap(imapConfig) {
  return new Promise((resolve) => {
    const imap = new Imap(imapConfig);
    let inboxEmail = null, junkEmail = null;
    let finished = false;

    const finish = () => {
      if (finished) return;
      finished = true;
      imap.end();
      resolve(getLatestEmail(inboxEmail, junkEmail));
    };

    imap.once('error', () => finish());
    imap.once('ready', async () => {
      const readBox = async (boxName, folderName) => {
        return new Promise(async (cb) => {
          try {
            await new Promise(r => imap.openBox(boxName, true, r));
            const ids = await new Promise((r, j) => imap.search(['ALL'], (err, res) => err ? j(err) : r(res)));
            if (!ids.length) return cb(null);
            const targetIds = ids.slice(-3);
            const f = imap.fetch(targetIds, { bodies: '' });
            let targetMail = null;
            f.on('message', async (msg) => {
              const stream = await Promise.race([
                new Promise(r => msg.on('body', r)),
                new Promise((_, rej) => setTimeout(() => rej(new Error('流读取超时')), 3000))
              ]);
              const mail = await simpleParser(stream);
              const mailObj = {
                send: escapeJson(mail.from?.text || '未知发件人'),
                subject: escapeJson(mail.subject || '无主题'),
                text: escapeJson(mail.text || ''),
                html: mail.html || `<p>${escapeHtml(mail.text || '').replace(/\n/g, '<br>')}</p>`,
                date: mail.date || new Date().toISOString(),
                folderSource: folderName,
                verifyCode: getVerifyCodeFromEmail({ text: mail.text, html: mail.html }, mail.subject)
              };
              if (!targetMail || new Date(mailObj.date) > new Date(targetMail.date)) {
                targetMail = mailObj;
              }
            });
            f.once('end', () => cb(targetMail));
          } catch (err) {
            cb(null);
          }
        });
      };

      Promise.all([
        readBox('INBOX', '收件箱'),
        readBox('Junk', '垃圾箱')
      ]).then(([inboxRes, junkRes]) => {
        inboxEmail = inboxRes;
        junkEmail = junkRes;
        finish();
      }).catch(() => finish());
    });
    imap.connect();
  });
}

const generateAuthString = (user, accessToken) => {
  const authString = `user=${user}\x01auth=Bearer ${accessToken}\x01\x01`;
  return Buffer.from(authString).toString('base64');
};

// ===================== 主入口：单次IMAP，无重试 =====================
module.exports = async (req, res) => {
  try {
    if (!CONFIG.SUPPORTED_METHODS.includes(req.method)) {
      return res.status(405).json({ code: 405, error: `仅支持 ${CONFIG.SUPPORTED_METHODS.join('/')}` });
    }

    const isGet = req.method === 'GET';
    const { password } = isGet ? req.query : req.body;
    if (process.env.PASSWORD && password !== process.env.PASSWORD) {
      return res.status(401).json({ code: 4010, error: '认证失败' });
    }

    const params = isGet ? req.query : req.body;
    const { refresh_token, client_id, email, response_type = 'json' } = params;
    const missing = CONFIG.REQUIRED_PARAMS.filter(k => !params[k]);
    if (missing.length) return res.status(400).json({ code: 4001, error: `缺少参数：${missing.join('、')}` });

    const paramErr = validateParams(params);
    if (paramErr) return res.status(400).json({ code: 4002, error: paramErr.message });

    let mail = null;
    let imapToken = null;
    try {
      imapToken = await getImapToken(refresh_token, client_id);
    } catch (e) {
      console.log("IMAP Token获取失败，直接走Graph");
    }

    // 仅单次IMAP，移除重试逻辑
    if (imapToken) {
      console.log("🔹 尝试IMAP取件");
      const imapConfig = { ...CONFIG.IMAP_CONFIG, user: email, xoauth2: generateAuthString(email, imapToken) };
      mail = await get_dual_folder_latest_email_imap(imapConfig);
    }

    // IMAP无结果，切换新Graph
    if (!mail) {
      console.log("🔹 IMAP无邮件，切换新Graph");
      const token = await getNewGrToken(refresh_token, client_id);
      if (token) mail = await get_dual_folder_latest_email_graph(token);
    }
    // 新Graph失败，切换老Graph
    if (!mail) {
      console.log("🔹 新Graph失败，切换老Graph");
      const token = await getOldGrToken(refresh_token, client_id);
      if (token) mail = await get_dual_folder_latest_email_graph(token);
    }

    if (mail) {
      return response_type === 'html'
        ? res.send(generateEmailHtml(mail))
        : res.json({ code: 200, message: '成功', data: [mail] });
    }
    return res.status(200).json({ code: 2001, message: '无邮件', data: null });

  } catch (error) {
    console.error(error);
    res.status(500).json({ code: 5000, error: '服务器异常：' + error.message });
  }
};
