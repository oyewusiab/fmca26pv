/**
 * PAYABLE VOUCHER 2026 - Footer Component
 * Adds consistent footer to all pages
 */

const Footer = {
    /**
     * Inserts the footer into the page
     */
    init() {
        // Find existing footer or create one
        let footer = document.querySelector('footer.main-footer');
        
        if (!footer) {
            footer = document.createElement('footer');
            footer.className = 'main-footer';
            
            // Find main content area to append footer
            const mainContent = document.querySelector('.main-content');
            if (mainContent) {
                mainContent.appendChild(footer);
            } else {
                document.body.appendChild(footer);
            }
        }
        
        footer.innerHTML = `
            <div class="footer-content">
                <p>© ${new Date().getFullYear()} Federal Medical Centre, Abeokuta • Finance & Accounts Department</p>
                <p class="powered-by">• Powered by ABLEBIZ @ <a href="mailto:hello@ablebiz.com.ng">hello@ablebiz.com.ng</a> •</p>
            </div>
        `;
        
        // Add footer styles if not already added
        if (!document.getElementById('footerStyles')) {
            const style = document.createElement('style');
            style.id = 'footerStyles';
            style.textContent = `
                .main-footer {
                    margin-top: 40px;
                    padding: 20px;
                    text-align: center;
                    border-top: 1px solid var(--border-color);
                    background: var(--white);
                }
                
                .main-footer .footer-content {
                    max-width: 600px;
                    margin: 0 auto;
                }
                
                .main-footer p {
                    margin: 5px 0;
                    font-size: 12px;
                    color: var(--text-muted);
                }
                
                .main-footer .powered-by {
                    font-size: 11px;
                    color: var(--primary-color);
                    font-weight: 500;
                }
                
                .main-footer .powered-by a {
                    color: var(--primary-color);
                    text-decoration: none;
                }
                
                .main-footer .powered-by a:hover {
                    text-decoration: underline;
                }
                
                @media print {
                    .main-footer {
                        position: fixed;
                        bottom: 0;
                        width: 100%;
                    }
                }
            `;
            document.head.appendChild(style);
        }
    }
};

// Auto-initialize when DOM is ready
document.addEventListener('DOMContentLoaded', () => Footer.init());