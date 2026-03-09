const Profile = {
  showLoading(show) {
    document.getElementById('loadingOverlay')?.classList.toggle('hidden', !show);
  },

  async init() {
    const ok = await Auth.requireAuth();
    if (!ok) return;

    document.getElementById('sidebar').innerHTML = Components.getSidebar('profile');

    document.getElementById('menuToggle')?.addEventListener('click', () => {
      document.getElementById('sidebar')?.classList.toggle('active');
    });

    await this.loadProfile();
    this.bind();
  },

  async loadProfile() {
    this.showLoading(true);
    try {
      const res = await API.getMyProfile();
      if (!res.success) {
        Utils.showToast(res.error || 'Failed to load profile', 'error');
        return;
      }

      const p = res.profile || {};
      document.getElementById('pName').value = p.name || '';
      document.getElementById('pEmail').value = p.email || '';
      document.getElementById('pUsername').value = p.username || '';
      document.getElementById('pRole').value = p.role || '';

      document.getElementById('pPhone').value = p.phone || '';
      document.getElementById('pDepartment').value = p.department || '';
    } finally {
      this.showLoading(false);
    }
  },

  bind() {
    document.getElementById('saveProfileBtn')?.addEventListener('click', async () => {
      const profile = {
        phone: document.getElementById('pPhone').value,
        department: document.getElementById('pDepartment').value
      };

      const res = await API.updateMyProfile(profile);
      if (res.success) Utils.showToast(res.message || 'Profile updated', 'success');
      else Utils.showToast(res.error || 'Update failed', 'error');
    });

    document.getElementById('changePasswordBtn')?.addEventListener('click', async () => {
      const oldPassword = document.getElementById('oldPassword').value;
      const newPassword = document.getElementById('newPassword').value;
      const confirm = document.getElementById('confirmNewPassword').value;

      if (!oldPassword || !newPassword || !confirm) return Utils.showToast('All fields required', 'warning');
      if (newPassword.length < 6) return Utils.showToast('Password must be at least 6 characters', 'warning');
      if (newPassword !== confirm) return Utils.showToast('Passwords do not match', 'error');

      const res = await API.changePassword(oldPassword, newPassword);
      if (res.success) {
        Utils.showToast(res.message || 'Password changed', 'success');
        document.getElementById('oldPassword').value = '';
        document.getElementById('newPassword').value = '';
        document.getElementById('confirmNewPassword').value = '';
      } else {
        Utils.showToast(res.error || 'Password change failed', 'error');
      }
    });
  }
};

document.addEventListener('DOMContentLoaded', () => Profile.init());