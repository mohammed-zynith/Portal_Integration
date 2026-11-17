import dotenv from 'dotenv';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
dotenv.config();
dotenv.config({ path: path.resolve(__dirname, '.env') });

import { sendPasswordResetEmail, testEmailConfig, getActiveEmailProvider, getSenderEmail } from './services/emailService.js';

const provider = getActiveEmailProvider();

console.log('==============================================');
console.log(`--- Email Diagnostic Test [Provider: ${provider.toUpperCase()}] ---`);
console.log('==============================================');
console.log('NODE_ENV:', process.env.NODE_ENV);
console.log('EMAIL_PROVIDER:', process.env.EMAIL_PROVIDER || '(auto-detected: ' + provider + ')');
console.log('EMAIL_BACKEND:', process.env.EMAIL_BACKEND);
console.log('SENDER_EMAIL:', getSenderEmail());
console.log('ALLOWED_EMAIL_DOMAINS:', process.env.ALLOWED_EMAIL_DOMAINS || '(none - all allowed)');

if (provider === 'office365') {
  console.log('AZURE_TENANT_ID:', process.env.AZURE_TENANT_ID ? `${process.env.AZURE_TENANT_ID.slice(0, 6)}...` : '(not set)');
  console.log('AZURE_CLIENT_ID:', process.env.AZURE_CLIENT_ID ? `${process.env.AZURE_CLIENT_ID.slice(0, 6)}...` : '(not set)');
  console.log('AZURE_CLIENT_SECRET:', process.env.AZURE_CLIENT_SECRET ? '****** (set)' : '(not set)');
  console.log('EMAIL_HOST_USER:', process.env.EMAIL_HOST_USER || '(not set)');
  console.log('DEFAULT_FROM_EMAIL:', process.env.DEFAULT_FROM_EMAIL || '(not set)');
} else if (provider === 'ses') {
  console.log('AWS_REGION:', process.env.AWS_REGION || process.env.AWS_SES_REGION);
  console.log('AWS_ACCESS_KEY_ID present?:', Boolean(process.env.AWS_ACCESS_KEY_ID));
} else {
  console.log('EMAIL_SERVICE:', process.env.EMAIL_SERVICE);
  console.log('EMAIL_HOST:', process.env.EMAIL_HOST);
  console.log('EMAIL_PORT:', process.env.EMAIL_PORT);
  console.log('EMAIL_USER:', process.env.EMAIL_USER);
}

async function runTest() {
  console.log('\n1. Testing Email Provider Configuration / Authentication...');
  const verifyResult = await testEmailConfig();
  console.log('Verify Result:', verifyResult ? '✅ PASSED' : '❌ FAILED');

  if (process.argv.includes('--send')) {
    console.log('\n2. Testing Send Password Reset Email...');
    const targetEmail = process.argv[3] || process.env.DEFAULT_FROM_EMAIL || process.env.EMAIL_HOST_USER || process.env.EMAIL_USER || 'test@example.com';
    const fakeUser = {
      email: targetEmail,
      personId: 'TEST001'
    };
    const fakeToken = 'test-token-1234567890abcdef';

    const sendResult = await sendPasswordResetEmail(fakeUser, fakeToken);
    console.log('Send Result:', sendResult);
  } else {
    console.log('\nTip: Run `node testEmail.js --send your-email@domain.com` to test actual email delivery.');
  }
}

runTest();
