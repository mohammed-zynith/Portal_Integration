import nodemailer from 'nodemailer';
import { SESv2Client, SendEmailCommand } from '@aws-sdk/client-sesv2';
import axios from 'axios';
import dotenv from 'dotenv';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// Load .env from current dir and backend dir if not already loaded
dotenv.config();
dotenv.config({ path: path.resolve(__dirname, '../.env') });
dotenv.config({ path: path.resolve(__dirname, '.env') });

// Helper to get sender email address
export const getSenderEmail = () => {
  return (
    process.env.DEFAULT_FROM_EMAIL ||
    process.env.EMAIL_HOST_USER ||
    process.env.EMAIL_FROM ||
    process.env.EMAIL_USER ||
    'admin@zynith-it.com'
  );
};

// Check if an email address domain is allowed by ALLOWED_EMAIL_DOMAINS whitelist
export const isEmailDomainAllowed = (email) => {
  if (!email || typeof email !== 'string') return false;
  const rawAllowed = process.env.ALLOWED_EMAIL_DOMAINS;
  if (!rawAllowed || rawAllowed.trim() === '' || rawAllowed.trim() === '*') {
    return true;
  }
  const allowedList = rawAllowed
    .split(',')
    .map(d => d.trim().toLowerCase())
    .filter(Boolean);

  if (allowedList.length === 0) return true;

  const domain = email.split('@')[1]?.toLowerCase();
  return domain ? allowedList.includes(domain) : false;
};

// Determine the active email provider
export const getActiveEmailProvider = () => {
  const provider = (process.env.EMAIL_PROVIDER || '').toLowerCase();
  if (provider === 'office365' || provider === 'azure' || provider === 'microsoft') {
    return 'office365';
  }
  if (provider === 'ses' || provider === 'aws') {
    return 'ses';
  }
  if (provider === 'smtp' || provider === 'gmail') {
    return 'smtp';
  }

  // If EMAIL_BACKEND is set to Office365 backend
  if (process.env.EMAIL_BACKEND && process.env.EMAIL_BACKEND.toLowerCase().includes('office365')) {
    return 'office365';
  }

  // Auto-detect Office 365 / Azure AD credentials
  if (
    process.env.AZURE_TENANT_ID &&
    process.env.AZURE_CLIENT_ID &&
    process.env.AZURE_CLIENT_SECRET &&
    !process.env.AZURE_TENANT_ID.includes('###')
  ) {
    return 'office365';
  }

  return 'smtp';
};

// In-memory cache for Azure AD access token
let cachedAzureToken = null;
let azureTokenExpiresAt = 0;

// Acquire access token from Azure AD using OAuth2 Client Credentials
export const getAzureAccessToken = async () => {
  const now = Date.now();
  // Reuse valid token with 60s safety buffer
  if (cachedAzureToken && azureTokenExpiresAt > now + 60000) {
    return cachedAzureToken;
  }

  const tenantId = process.env.AZURE_TENANT_ID;
  const clientId = process.env.AZURE_CLIENT_ID;
  const clientSecret = process.env.AZURE_CLIENT_SECRET;

  if (!tenantId || !clientId || !clientSecret) {
    throw new Error('Azure AD credentials missing: AZURE_TENANT_ID, AZURE_CLIENT_ID, and AZURE_CLIENT_SECRET are required.');
  }

  if (tenantId.includes('###') || clientId.includes('###') || clientSecret.includes('###')) {
    throw new Error('Azure AD credentials in .env contain placeholder values (###). Please provide actual credentials.');
  }

  const tokenUrl = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;

  const bodyParams = new URLSearchParams();
  bodyParams.append('client_id', clientId);
  bodyParams.append('client_secret', clientSecret);
  bodyParams.append('grant_type', 'client_credentials');
  bodyParams.append('scope', 'https://graph.microsoft.com/.default');

  const response = await axios.post(tokenUrl, bodyParams.toString(), {
    headers: {
      'Content-Type': 'application/x-www-form-urlencoded'
    },
    timeout: 10000
  });

  if (!response.data || !response.data.access_token) {
    throw new Error('Failed to acquire access token from Azure AD: response missing access_token');
  }

  cachedAzureToken = response.data.access_token;
  azureTokenExpiresAt = now + ((response.data.expires_in || 3600) * 1000);
  return cachedAzureToken;
};

// Send email via Microsoft Graph API (Office 365)
export const sendViaOffice365Graph = async (mailOptions) => {
  const token = await getAzureAccessToken();
  const senderEmail = process.env.EMAIL_HOST_USER || process.env.DEFAULT_FROM_EMAIL || getSenderEmail();

  // Normalize recipient(s)
  let recipients = [];
  if (Array.isArray(mailOptions.to)) {
    recipients = mailOptions.to;
  } else if (typeof mailOptions.to === 'string') {
    recipients = mailOptions.to.split(',').map(s => s.trim()).filter(Boolean);
  }

  const toRecipients = recipients.map(address => ({
    emailAddress: { address }
  }));

  const message = {
    subject: mailOptions.subject,
    body: {
      contentType: mailOptions.html ? 'HTML' : 'Text',
      content: mailOptions.html || mailOptions.text || ''
    },
    toRecipients
  };

  // Support attachments if provided
  if (mailOptions.attachments && Array.isArray(mailOptions.attachments) && mailOptions.attachments.length > 0) {
    message.attachments = mailOptions.attachments.map(att => {
      let contentBytes = '';
      if (Buffer.isBuffer(att.content)) {
        contentBytes = att.content.toString('base64');
      } else if (typeof att.content === 'string') {
        contentBytes = Buffer.from(att.content).toString('base64');
      }
      return {
        '@odata.type': '#microsoft.graph.fileAttachment',
        name: att.filename || 'attachment',
        contentType: att.contentType || 'application/octet-stream',
        contentBytes
      };
    });
  }

  const graphUrl = `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(senderEmail)}/sendMail`;

  const response = await axios.post(
    graphUrl,
    {
      message,
      saveToSentItems: true
    },
    {
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json'
      },
      timeout: 15000
    }
  );

  const messageId = response.headers?.['client-request-id'] || `office365-${Date.now()}`;
  console.log(`✅ Office 365 Graph email sent successfully (Message ID: ${messageId})`);
  return { success: true, messageId };
};

// Create transporter (AWS SES when EMAIL_PROVIDER=ses, otherwise standard SMTP/Gmail)
export const createTransporter = () => {
  const isSES = process.env.EMAIL_PROVIDER === 'ses';

  if (isSES) {
    const ses = new SESv2Client({
      region: process.env.AWS_SES_REGION || process.env.AWS_REGION || 'ap-south-1',
      credentials: {
        accessKeyId: process.env.AWS_ACCESS_KEY_ID,
        secretAccessKey: process.env.AWS_SECRET_ACCESS_KEY,
      },
    });

    return nodemailer.createTransport({
      SES: { sesClient: ses, SendEmailCommand },
    });
  }

  // Fallback to standard SMTP (development/local)
  // Gmail Service (Recommended for Gmail)
  if (process.env.EMAIL_SERVICE === 'gmail') {
    return nodemailer.createTransport({
      service: 'gmail',
      auth: {
        user: process.env.EMAIL_USER,
        pass: process.env.EMAIL_PASS,
      },
    });
  }

  // Fallback to custom SMTP host & port
  return nodemailer.createTransport({
    service: process.env.EMAIL_SERVICE,
    host: process.env.EMAIL_HOST || 'smtp.gmail.com',
    port: process.env.EMAIL_PORT ? parseInt(process.env.EMAIL_PORT, 10) : 587,
    secure: process.env.EMAIL_PORT === '465',
    auth: {
      user: process.env.EMAIL_USER,
      pass: process.env.EMAIL_PASS,
    },
  });
};

// Unified sendMail function that routes to the active provider (Office 365, SES, or SMTP)
export const sendMail = async (mailOptions) => {
  const recipient = Array.isArray(mailOptions.to) ? mailOptions.to[0] : mailOptions.to;

  // Domain whitelist check
  if (recipient && !isEmailDomainAllowed(recipient)) {
    const domain = recipient.split('@')[1];
    const errMsg = `Email domain "${domain}" for ${recipient} is not permitted in ALLOWED_EMAIL_DOMAINS`;
    console.warn(`⚠️ [Email Blocked]: ${errMsg}`);
    return { success: false, error: errMsg };
  }

  const provider = getActiveEmailProvider();
  console.log(`📧 Dispatching email to "${recipient}" via [${provider.toUpperCase()}] | Subject: "${mailOptions.subject}"`);

  if (provider === 'office365') {
    try {
      return await sendViaOffice365Graph(mailOptions);
    } catch (graphError) {
      const details = graphError.response?.data?.error?.message || graphError.message;
      console.error(`❌ Office 365 Graph send failed:`, details);
      return { success: false, error: `Office 365 send error: ${details}` };
    }
  }

  // SES or standard SMTP
  try {
    const transporter = createTransporter();
    const result = await transporter.sendMail(mailOptions);
    console.log(`✅ Email sent successfully via [${provider.toUpperCase()}]:`, result.messageId);
    return { success: true, messageId: result.messageId };
  } catch (smtpError) {
    console.error(`❌ Email send failed via [${provider.toUpperCase()}]:`, smtpError.message);
    return { success: false, error: smtpError.message };
  }
};

// Email template for employee credentials
const createEmployeeEmailTemplate = (employeeData, loginUrl) => {
  return `
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="utf-8">
        <style>
            body {
                font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
                line-height: 1.6;
                color: #333;
                max-width: 600px;
                margin: 0 auto;
                padding: 20px;
            }
            .header {
                background: linear-gradient(135deg, #4361ee, #7209b7);
                color: white;
                padding: 30px;
                text-align: center;
                border-radius: 10px 10px 0 0;
            }
            .content {
                background: #f8f9fa;
                padding: 30px;
                border-radius: 0 0 10px 10px;
            }
            .credentials {
                background: white;
                padding: 20px;
                border-radius: 8px;
                margin: 20px 0;
                border-left: 4px solid #4361ee;
            }
            .login-btn {
                display: inline-block;
                background: linear-gradient(135deg, #4361ee, #7209b7);
                color: white;
                padding: 12px 30px;
                text-decoration: none;
                border-radius: 5px;
                margin: 20px 0;
                font-weight: bold;
            }
            .footer {
                text-align: center;
                margin-top: 30px;
                color: #666;
                font-size: 14px;
            }
        </style>
    </head>
    <body>
        <div class="header">
            <h1>Welcome to Our Team! 🎉</h1>
        </div>
        <div class="content">
            <h2>Hello ${employeeData.personId},</h2>
            <p>Your employee account has been created successfully. Here are your login credentials:</p>
            
            <div class="credentials">
                <h3>Your Login Details:</h3>
                <p><strong>Employee ID:</strong> ${employeeData.personId}</p>
                <p><strong>Email:</strong> ${employeeData.email}</p>
                <p><strong>Password:</strong> ${employeeData.password}</p>
                <p><strong>Role:</strong> Employee</p>
            </div>

            <p>You can login using the button below:</p>
            <a href="${loginUrl}" class="login-btn">Login to Your Account</a>

            <div class="security-note">
                <h4>🔒 Security Tips:</h4>
                <ul>
                    <li>Keep your credentials secure</li>
                    <li>Change your password after first login</li>
                    <li>Do not share your login details with anyone</li>
                </ul>
            </div>
        </div>
        <div class="footer">
            <p>If you have any questions, please contact your administrator.</p>
            <p>© ${new Date().getFullYear()} Company Name. All rights reserved.</p>
        </div>
    </body>
    </html>
  `;
};

// Send employee credentials email
export const sendEmployeeCredentials = async (employeeData, plainPassword) => {
  try {
    const loginUrl = `${process.env.FRONTEND_URL}/login`;
    
    const mailOptions = {
      from: {
        name: process.env.EMAIL_FROM_NAME || 'Company HR System',
        address: getSenderEmail()
      },
      to: employeeData.email,
      subject: 'Your Employee Account Credentials',
      html: createEmployeeEmailTemplate(
        { ...employeeData, password: plainPassword },
        loginUrl
      ),
      text: `Welcome to Our Team!
      
Your employee account has been created successfully.

Login Details:
- Employee ID: ${employeeData.personId}
- Email: ${employeeData.email}
- Password: ${plainPassword}
- Role: Employee

Login URL: ${loginUrl}

Please keep your credentials secure and change your password after first login.

Best regards,
Company HR Team`
    };

    return await sendMail(mailOptions);
  } catch (error) {
    console.error('Error sending employee credentials email:', error);
    return { success: false, error: error.message };
  }
};

// Test email configuration
export const testEmailConfig = async () => {
  const provider = getActiveEmailProvider();
  console.log(`🔍 Testing email configuration using provider: [${provider.toUpperCase()}]`);

  if (provider === 'office365') {
    try {
      console.log('1. Testing Azure AD OAuth2 Token Acquisition...');
      const token = await getAzureAccessToken();
      if (!token) throw new Error('Acquired Azure AD token was empty');
      console.log('✅ Azure AD authentication successful (Token acquired)');

      const sender = process.env.EMAIL_HOST_USER || process.env.DEFAULT_FROM_EMAIL || getSenderEmail();
      console.log(`2. Verifying sending mailbox (${sender}) on Microsoft 365...`);
      try {
        const userCheck = await axios.get(
          `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(sender)}?$select=id,mail,userPrincipalName`,
          {
            headers: { Authorization: `Bearer ${token}` },
            timeout: 5000
          }
        );
        console.log(`✅ Mailbox verified on Microsoft 365: ${userCheck.data?.mail || userCheck.data?.userPrincipalName}`);
      } catch (userErr) {
        console.warn(`⚠️ Mailbox lookup note (${userErr.response?.data?.error?.message || userErr.message}). (If permissions are restricted to Mail.Send only, mailbox lookup is expected to be restricted, but sending will still work).`);
      }
      return true;
    } catch (error) {
      console.error('❌ Office 365 / Azure AD config test error:', error.response?.data || error.message);
      return false;
    }
  }

  try {
    const transporter = createTransporter();
    await transporter.verify();
    console.log(`✅ Email configuration is correct for [${provider.toUpperCase()}]`);
    return true;
  } catch (error) {
    console.error(`❌ Email configuration error for [${provider.toUpperCase()}]:`, error);
    return false;
  }
};

// Send payslip email
export const sendPayslipEmail = async (payslip) => {
  try {
    const mailOptions = {
      from: {
        name: process.env.EMAIL_FROM_NAME || 'Zynith IT Solutions - HR',
        address: getSenderEmail()
      },
      to: payslip.email,
      subject: `Salary Payslip - ${payslip.month} ${payslip.year}`,
      html: createPayslipEmailTemplate(payslip),
      text: `Dear ${payslip.name},

Your salary for ${payslip.month} ${payslip.year} has been processed.

Employee ID: ${payslip.employeeId}
Basic Salary: ₹${payslip.basicSalary}
Gross Earnings: ₹${payslip.grossEarnings}
Total Deductions: ₹${payslip.totalDeductions}
Net Pay: ₹${payslip.netPay}

Please find the attached payslip for detailed information.

Best regards,
Zynith IT Solutions HR Team`
    };

    return await sendMail(mailOptions);
  } catch (error) {
    console.error('Error sending payslip email:', error);
    return { success: false, error: error.message };
  }
};

// Payslip email template
const createPayslipEmailTemplate = (payslip) => {
  return `
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="utf-8">
        <style>
            body {
                font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
                line-height: 1.6;
                color: #333;
                max-width: 600px;
                margin: 0 auto;
                padding: 20px;
            }
            .header {
                background: linear-gradient(135deg, #4361ee, #7209b7);
                color: white;
                padding: 30px;
                text-align: center;
                border-radius: 10px 10px 0 0;
            }
            .content {
                background: #f8f9fa;
                padding: 30px;
                border-radius: 0 0 10px 10px;
            }
            .salary-summary {
                background: white;
                padding: 20px;
                border-radius: 8px;
                margin: 20px 0;
                border-left: 4px solid #4361ee;
            }
            .details-grid {
                display: grid;
                grid-template-columns: 1fr 1fr;
                gap: 1rem;
                margin: 1rem 0;
            }
            .net-pay {
                background: #dcfce7;
                padding: 1rem;
                border-radius: 8px;
                text-align: center;
                font-weight: bold;
                font-size: 1.2rem;
                margin: 1rem 0;
            }
            .footer {
                text-align: center;
                margin-top: 30px;
                color: #666;
                font-size: 14px;
            }
        </style>
    </head>
    <body>
        <div class="header">
            <h1>Zynith IT Solutions</h1>
            <h2>Salary Payslip</h2>
        </div>
        <div class="content">
            <h2>Hello ${payslip.name},</h2>
            <p>Your salary for <strong>${payslip.month} ${payslip.year}</strong> has been processed successfully.</p>
            
            <div class="salary-summary">
                <h3>Salary Details</h3>
                <div class="details-grid">
                    <div>
                        <strong>Employee ID:</strong><br>
                        ${payslip.employeeId}
                    </div>
                    <div>
                        <strong>Pay Date:</strong><br>
                        ${new Date(payslip.payDate).toLocaleDateString()}
                    </div>
                </div>
                
                <h4>Earnings</h4>
                <p><strong>Basic Salary:</strong> ₹${payslip.basicSalary.toFixed(2)}</p>
                ${payslip.earnings.map(earning => `
                    <p><strong>${earning.type}:</strong> ₹${earning.amount.toFixed(2)}</p>
                `).join('')}
                <p><strong>Gross Earnings:</strong> ₹${payslip.grossEarnings.toFixed(2)}</p>
                
                <h4>Deductions</h4>
                ${payslip.deductions.map(deduction => `
                    <p><strong>${deduction.type}:</strong> ₹${deduction.amount.toFixed(2)}</p>
                `).join('')}
                <p><strong>Total Deductions:</strong> ₹${payslip.totalDeductions.toFixed(2)}</p>
                
                <div class="net-pay">
                    Net Pay: ₹${payslip.netPay.toFixed(2)}
                </div>
                
                <p><strong>Paid Days:</strong> ${payslip.paidDays}</p>
                <p><strong>LOP Days:</strong> ${payslip.lopDays}</p>
            </div>
            
            <p>This is an automated email. Please do not reply to this message.</p>
        </div>
        <div class="footer">
            <p>If you have any questions, please contact the HR department.</p>
            <p>© ${new Date().getFullYear()} Zynith IT Solutions. All rights reserved.</p>
        </div>
    </body>
    </html>
  `;
};

// Password reset email template
const createPasswordResetEmailTemplate = (resetUrl, user) => {
  const displayName = user.personId || user.email;
  return `
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="utf-8">
        <style>
            body {
                font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
                line-height: 1.6;
                color: #333;
                max-width: 600px;
                margin: 0 auto;
                padding: 20px;
            }
            .header {
                background: linear-gradient(135deg, #06510f, #3a9440);
                color: white;
                padding: 30px;
                text-align: center;
                border-radius: 10px 10px 0 0;
            }
            .content {
                background: #f8f9fa;
                padding: 30px;
                border-radius: 0 0 10px 10px;
            }
            .reset-card {
                background: white;
                padding: 25px;
                border-radius: 8px;
                margin: 20px 0;
                border-left: 4px solid #3a9440;
                text-align: center;
            }
            .reset-btn {
                display: inline-block;
                background: linear-gradient(135deg, #3a9440, #5acd62);
                color: white !important;
                padding: 14px 35px;
                text-decoration: none;
                border-radius: 25px;
                margin: 20px 0;
                font-weight: bold;
                font-size: 16px;
                box-shadow: 0 4px 15px rgba(58, 148, 64, 0.3);
            }
            .link-fallback {
                word-break: break-all;
                font-size: 13px;
                color: #666;
                background: #eee;
                padding: 10px;
                border-radius: 5px;
                margin-top: 15px;
            }
            .security-note {
                background: #fff3cd;
                border: 1px solid #ffeeba;
                color: #856404;
                padding: 15px;
                border-radius: 6px;
                margin-top: 20px;
                font-size: 14px;
            }
            .footer {
                text-align: center;
                margin-top: 30px;
                color: #666;
                font-size: 14px;
            }
        </style>
    </head>
    <body>
        <div class="header">
            <h1>Zynith IT Solutions</h1>
            <h2>Password Reset Request</h2>
        </div>
        <div class="content">
            <h2>Hello ${displayName},</h2>
            <p>We received a request to reset the password for your account (<strong>${user.email}</strong>).</p>
            
            <div class="reset-card">
                <p>Click the button below to reset your password. This link is valid for <strong>1 hour</strong>.</p>
                <a href="${resetUrl}" class="reset-btn" target="_blank">Reset Password</a>
                <p style="margin-top: 15px; font-size: 13px; color: #777;">If the button doesn't work, copy and paste this link into your browser:</p>
                <div class="link-fallback">${resetUrl}</div>
            </div>

            <div class="security-note">
                <strong>🔒 Security Notice:</strong>
                <p style="margin: 5px 0 0 0;">If you did not request a password reset, please ignore this email or contact support if you suspect unauthorized access. Your password will remain unchanged.</p>
            </div>
        </div>
        <div class="footer">
            <p>This is an automated message from Zynith IT Solutions Finance Portal.</p>
            <p>© ${new Date().getFullYear()} Zynith IT Solutions. All rights reserved.</p>
        </div>
    </body>
    </html>
  `;
};

// Send password reset email
export const sendPasswordResetEmail = async (user, rawToken) => {
  try {
    const frontendUrl = process.env.FRONTEND_URL || 'http://localhost:5173';
    const resetUrl = `${frontendUrl}/reset-password/${rawToken}`;

    const mailOptions = {
      from: {
        name: process.env.EMAIL_FROM_NAME || 'Zynith IT Solutions - Security',
        address: getSenderEmail()
      },
      to: user.email,
      subject: 'Password Reset Request - Zynith IT Solutions',
      html: createPasswordResetEmailTemplate(resetUrl, user),
      text: `Hello ${user.personId || user.email},

We received a request to reset your password for Zynith IT Solutions Finance Portal.

Please reset your password using the link below (valid for 1 hour):
${resetUrl}

If you did not request this password reset, please ignore this email.

Best regards,
Zynith IT Solutions Team`
    };

    return await sendMail(mailOptions);
  } catch (error) {
    console.error('Error sending password reset email:', error);
    return { success: false, error: error.message };
  }
};