/**
 * PAYABLE VOUCHER 2026 - Activity Tracker
 * Auto-logout after inactivity
 * - Regular users: 5 minutes, immediate logout
 * - Admin users: 5 minutes, then 60-second countdown before logout
 */

const ActivityTracker = {
    // Configuration
    INACTIVITY_TIMEOUT: 10 * 60 * 1000,      // 10 minutes in milliseconds
    ADMIN_COUNTDOWN: 60,                      // 60 seconds countdown for admin
    
    // State
    lastActivity: Date.now(),
    timeoutId: null,
    countdownId: null,
    countdownModal: null,
    countdownSeconds: 60,
    isCountingDown: false,
    
    /**
     * Initialize the activity tracker
     */
    init() {
        // Only track if user is logged in
        if (!Auth.isLoggedIn()) return;
        
        // Track user activity
        this.setupActivityListeners();
        
        // Start the inactivity timer
        this.resetTimer();
        
        // Create countdown modal (for admin)
        this.createCountdownModal();
        
        console.log('Activity tracker initialized');
    },
    
    /**
     * Setup activity listeners
     */
    setupActivityListeners() {
        const events = ['mousedown', 'mousemove', 'keydown', 'scroll', 'touchstart', 'click'];
        
        events.forEach(event => {
            document.addEventListener(event, () => this.handleActivity(), { passive: true });
        });
    },
    
    /**
     * Handle user activity
     */
    handleActivity() {
        this.lastActivity = Date.now();
        
        // If countdown is active, cancel it (user came back)
        if (this.isCountingDown) {
            this.cancelCountdown();
        }
        
        // Reset the timer
        this.resetTimer();
    },
    
    /**
     * Reset the inactivity timer
     */
    resetTimer() {
        // Clear existing timeout
        if (this.timeoutId) {
            clearTimeout(this.timeoutId);
        }
        
        // Set new timeout
        this.timeoutId = setTimeout(() => {
            this.handleInactivity();
        }, this.INACTIVITY_TIMEOUT);
    },
    
    /**
     * Handle inactivity timeout
     */
    handleInactivity() {
        const user = Auth.getUser();
        
        if (!user) {
            // Not logged in, nothing to do
            return;
        }
        
        if (user.role === CONFIG.ROLES.ADMIN) {
            // Admin: Show countdown
            this.showCountdown();
        } else {
            // Regular user: Immediate logout
            this.performLogout('Session expired due to inactivity.');
        }
    },
    
    /**
     * Create the countdown modal (hidden initially)
     */
    createCountdownModal() {
        // Remove existing modal if any
        const existing = document.getElementById('inactivityModal');
        if (existing) existing.remove();
        
        // Create modal
        const modal = document.createElement('div');
        modal.id = 'inactivityModal';
        modal.className = 'inactivity-modal-overlay';
        modal.innerHTML = `
            <div class="inactivity-modal">
                <div class="inactivity-modal-icon">
                    <i class="fas fa-clock"></i>
                </div>
                <h3>Session Timeout Warning</h3>
                <p>You have been inactive for 5 minutes.</p>
                <p>You will be logged out in:</p>
                <div class="countdown-timer" id="countdownTimer">60</div>
                <p>seconds</p>
                <div class="inactivity-modal-buttons">
                    <button class="btn btn-primary" id="stayLoggedInBtn">
                        <i class="fas fa-user-check"></i> Stay Logged In
                    </button>
                    <button class="btn btn-secondary" id="logoutNowBtn">
                        <i class="fas fa-sign-out-alt"></i> Logout Now
                    </button>
                </div>
            </div>
        `;
        
        document.body.appendChild(modal);
        
        // Add styles
        if (!document.getElementById('inactivityStyles')) {
            const style = document.createElement('style');
            style.id = 'inactivityStyles';
            style.textContent = `
                .inactivity-modal-overlay {
                    position: fixed;
                    top: 0;
                    left: 0;
                    right: 0;
                    bottom: 0;
                    background: rgba(0, 0, 0, 0.7);
                    display: flex;
                    align-items: center;
                    justify-content: center;
                    z-index: 10000;
                    opacity: 0;
                    visibility: hidden;
                    transition: all 0.3s ease;
                }
                
                .inactivity-modal-overlay.active {
                    opacity: 1;
                    visibility: visible;
                }
                
                .inactivity-modal {
                    background: white;
                    border-radius: 12px;
                    padding: 40px;
                    text-align: center;
                    max-width: 400px;
                    width: 90%;
                    box-shadow: 0 10px 40px rgba(0, 0, 0, 0.3);
                    animation: modalPulse 0.5s ease;
                }
                
                @keyframes modalPulse {
                    0% { transform: scale(0.9); }
                    50% { transform: scale(1.02); }
                    100% { transform: scale(1); }
                }
                
                .inactivity-modal-icon {
                    font-size: 48px;
                    color: var(--warning-color);
                    margin-bottom: 20px;
                }
                
                .inactivity-modal h3 {
                    color: var(--text-dark);
                    margin-bottom: 15px;
                    font-size: 22px;
                }
                
                .inactivity-modal p {
                    color: var(--text-muted);
                    margin-bottom: 10px;
                }
                
                .countdown-timer {
                    font-size: 64px;
                    font-weight: 700;
                    color: var(--danger-color);
                    margin: 20px 0;
                    font-family: monospace;
                }
                
                .inactivity-modal-buttons {
                    display: flex;
                    gap: 10px;
                    justify-content: center;
                    margin-top: 25px;
                }
                
                .inactivity-modal-buttons .btn {
                    min-width: 140px;
                }
            `;
            document.head.appendChild(style);
        }
        
        // Setup button handlers
        document.getElementById('stayLoggedInBtn').addEventListener('click', () => {
            this.cancelCountdown();
        });
        
        document.getElementById('logoutNowBtn').addEventListener('click', () => {
            this.performLogout('You have been logged out.');
        });
        
        this.countdownModal = modal;
    },
    
    /**
     * Show countdown modal (Admin only)
     */
    showCountdown() {
        this.isCountingDown = true;
        this.countdownSeconds = this.ADMIN_COUNTDOWN;
        
        // Show modal
        if (this.countdownModal) {
            this.countdownModal.classList.add('active');
        }
        
        // Update timer display
        document.getElementById('countdownTimer').textContent = this.countdownSeconds;
        
        // Play alert sound (optional)
        this.playAlertSound();
        
        // Start countdown
        this.countdownId = setInterval(() => {
            this.countdownSeconds--;
            document.getElementById('countdownTimer').textContent = this.countdownSeconds;
            
            // Play tick sound in last 10 seconds
            if (this.countdownSeconds <= 10) {
                this.playTickSound();
            }
            
            if (this.countdownSeconds <= 0) {
                this.performLogout('Session expired due to inactivity.');
            }
        }, 1000);
    },
    
    /**
     * Cancel countdown and hide modal
     */
    cancelCountdown() {
        this.isCountingDown = false;
        
        // Clear countdown interval
        if (this.countdownId) {
            clearInterval(this.countdownId);
            this.countdownId = null;
        }
        
        // Hide modal
        if (this.countdownModal) {
            this.countdownModal.classList.remove('active');
        }
        
        // Reset activity timer
        this.resetTimer();
        
        Utils.showToast('Session extended. Welcome back!', 'success');
    },
    
    /**
     * Perform logout
     */
    performLogout(message) {
        // Clear all timers
        if (this.timeoutId) clearTimeout(this.timeoutId);
        if (this.countdownId) clearInterval(this.countdownId);
        
        // Show message
        alert(message);
        
        // Logout
        Auth.logout();
    },
    
    /**
     * Play alert sound
     */
    playAlertSound() {
        try {
            const audio = new Audio('data:audio/wav;base64,UklGRnoGAABXQVZFZm10IBAAAAABAAEAQB8AAEAfAAABAAgAZGF0YQoGAACBhYqFbF1fdH2Onp6WgnZnYGJuhZ2qq5+PfW1obnuQo6iunYx6amVpdIWYqLGsnIp4amRndYOUp6yonIl3ZWJkc4KSoqejnIh0Y2BhcIGQnqOhnId0YmBgb4COm5+cmIR0YWBgb3+Nm5yamYR0');
            audio.volume = 0.5;
            audio.play().catch(() => {}); // Ignore autoplay errors
        } catch (e) {}
    },
    
    /**
     * Play tick sound
     */
    playTickSound() {
        try {
            const audio = new Audio('data:audio/wav;base64,UklGRl9vAABXQVZFZm10IBAAAAABAAEAQB8AAEAfAAABAAgAZGF0YQ==');
            audio.volume = 0.3;
            audio.play().catch(() => {});
        } catch (e) {}
    },
    
    /**
     * Cleanup - call when logging out manually
     */
    cleanup() {
        if (this.timeoutId) clearTimeout(this.timeoutId);
        if (this.countdownId) clearInterval(this.countdownId);
        if (this.countdownModal) this.countdownModal.remove();
    }
};

// Initialize when DOM is ready (if user is logged in)
document.addEventListener('DOMContentLoaded', () => {
    if (Auth.isLoggedIn()) {
        ActivityTracker.init();
    }
});