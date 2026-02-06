// Prototype user login: verify email exists in 'users' table, accept any non-empty password
import { supabase } from './supabase-client.js';

function initUserLogin() {
    const form = document.getElementById('userLoginForm');
    const emailEl = document.getElementById('loginEmail');
    const pwdEl = document.getElementById('loginPassword');
    const submitBtn = form.querySelector('button[type="submit"]');
    const msgBox = document.querySelector('.login-status-message');

    function showMessage(text, type = 'info') {
        msgBox.textContent = text;
        msgBox.style.display = '';
        msgBox.style.padding = '12px';
        msgBox.style.borderRadius = '6px';
        msgBox.style.backgroundColor = type === 'error' ? '#fde8e8' : type === 'success' ? '#e7f6ec' : '#f3f4f6';
        msgBox.style.color = type === 'error' ? '#991b1b' : type === 'success' ? '#065f46' : '#111827';
        msgBox.setAttribute('role', 'status');
    }

    function setLoading(loading) {
        submitBtn.disabled = loading;
        submitBtn.textContent = loading ? 'Checking...' : 'Log In';
    }

    form.addEventListener('submit', async (e) => {
        e.preventDefault();
        msgBox.style.display = 'none';

        const email = (emailEl.value || '').trim().toLowerCase();
        const password = (pwdEl.value || '').trim();

        if (!email) {
            showMessage('Please enter your email.', 'error');
            return;
        }
        if (!password) {
            showMessage('Please enter a password. Any non-empty value is accepted for this prototype.', 'error');
            return;
        }

        try {
            setLoading(true);

            // Verify the user exists by email in public.users
            const { data, error } = await supabase
                .from('users')
                .select('id, name, email, organizationid, usertype')
                .eq('email', email)
                .maybeSingle();

            if (error) {
                showMessage('Unable to verify user right now. Please try again later.', 'error');
                return;
            }
            if (!data) {
                showMessage('User not found. Please check your email or contact your administrator.', 'error');
                return;
            }

            // Prototype "success": email exists and password is non-empty
            showMessage('Login successful. Redirecting...', 'success');

            // Store minimal session info
            try {
                sessionStorage.setItem('ct_user', JSON.stringify({
                    id: data.id,
                    name: data.name,
                    email: data.email,
                    organizationid: data.organizationid,
                    usertype: data.usertype,
                    ts: Date.now()
                }));
            } catch (_) {}

            // Redirect to the user portal
            setTimeout(() => {
                window.location.href = 'user-portal.html';
            }, 500);
        } catch (_) {
            showMessage('Unexpected error. Please try again.', 'error');
        } finally {
            setLoading(false);
        }
    });
}

if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', initUserLogin);
} else {
    initUserLogin();
}


