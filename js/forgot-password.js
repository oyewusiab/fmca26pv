document.addEventListener('DOMContentLoaded', () => {
  const step1 = document.getElementById('fpStep1');
  const step2 = document.getElementById('fpStep2');

  const identifierEl = document.getElementById('fpIdentifier');
  const otpEl = document.getElementById('fpOtp');
  const newPassEl = document.getElementById('fpNewPassword');
  const confirmEl = document.getElementById('fpConfirmPassword');

  document.getElementById('sendOtpBtn')?.addEventListener('click', async () => {
    const identifier = (identifierEl.value || '').trim();
    if (!identifier) return Utils.showToast('Enter email or username', 'warning');

    const res = await API.requestPasswordReset(identifier);
    if (res.success) {
      Utils.showToast(res.message || 'If the account exists, OTP has been sent.', 'success');
      step1.classList.add('hidden');
      step2.classList.remove('hidden');
    } else {
      Utils.showToast(res.error || 'Failed to send OTP', 'error');
    }
  });

  document.getElementById('resetPasswordBtn')?.addEventListener('click', async () => {
    const identifier = (identifierEl.value || '').trim();
    const otp = (otpEl.value || '').trim();
    const newPassword = (newPassEl.value || '').trim();
    const confirmPassword = (confirmEl.value || '').trim();

    if (!otp || !newPassword || !confirmPassword) return Utils.showToast('All fields are required', 'warning');
    if (newPassword.length < 6) return Utils.showToast('Password must be at least 6 characters', 'warning');
    if (newPassword !== confirmPassword) return Utils.showToast('Passwords do not match', 'error');

    const res = await API.resetPasswordWithOtp(identifier, otp, newPassword);
    if (res.success) {
      Utils.showToast(res.message || 'Password reset successful', 'success');
      window.location.href = 'index.html';
    } else {
      Utils.showToast(res.error || 'Reset failed', 'error');
    }
  });
});