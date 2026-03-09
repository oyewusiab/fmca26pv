const ChatWidget = {
  open: false,
  selectedUser: null,
  pollId: null,
  threadPoll: null,
  allUsers: [],
  unreadMap: {},

  async init() {
    if (!Auth.isLoggedIn()) return;

    this.currentUser = Auth.getUser();
    this.injectUI();
    await this.refreshUnreadBadge();

    // Poll unread every 15s only when panel is open
    this.pollId = setInterval(() => {
      if (this.open) this.refreshUnreadBadge();
    }, 15000);
  },

  injectUI() {
    if (document.getElementById('chatFab')) return;

    const fab = document.createElement('div');
    fab.id = 'chatFab';
    fab.innerHTML = `
      <button class="chat-fab-btn" id="chatOpenBtn" title="Chat">
        <i class="fas fa-comment-dots"></i>
        <span class="chat-badge" id="chatBadge" style="display:none;">0</span>
      </button>

      <div class="chat-panel" id="chatPanel" style="display:none;">
        <div class="chat-header">
          <strong>Chat</strong>
          <button class="chat-close" id="chatCloseBtn">&times;</button>
        </div>

        <div class="chat-body">
          <div class="chat-users">
            <input type="text" placeholder="Search users..." id="chatUserSearch" />
            <div class="chat-users-list" id="chatUsers">Loading users...</div>
          </div>

          <div class="chat-thread">
            <div class="chat-thread-header" id="chatThreadHeader">Select a user</div>
            <div class="chat-messages" id="chatMessages"></div>
            <div class="chat-typing-indicator" id="chatTypingIndicator" style="display:none;">Typing...</div>
            <div class="chat-input">
              <input type="text" id="chatMessageInput" placeholder="Type a message..." />
              <button class="btn btn-sm" id="chatEmojiBtn">😊</button>
              <button class="btn btn-primary btn-sm" id="chatSendBtn"><i class="fas fa-paper-plane"></i></button>
            </div>
          </div>
        </div>
      </div>
    `;
    document.body.appendChild(fab);

    const style = document.createElement('style');
    style.textContent = `
      /* --- FAB & Badge --- */
      #chatFab{position:fixed;bottom:20px;right:20px;z-index:9999;}
      .chat-fab-btn{width:54px;height:54px;border-radius:50%;background:var(--primary-color);color:#fff;border:none;cursor:pointer;box-shadow:0 6px 12px rgba(0,0,0,0.2);position:relative;}
      .chat-badge{position:absolute;top:-6px;right:-6px;background:var(--danger-color);color:#fff;border-radius:12px;padding:2px 6px;font-size:10px;font-weight:700;}

      /* --- Chat Panel --- */
      .chat-panel{position:absolute;bottom:70px;right:0;width:400px;max-width:95vw;height:550px;background:#fff;border-radius:16px;box-shadow:0 10px 30px rgba(0,0,0,0.25);display:flex;flex-direction:column;overflow:hidden;border:1px solid #ddd;}
      .chat-header{display:flex;justify-content:space-between;align-items:center;padding:12px;background:var(--primary-color);color:#fff;font-weight:700;}
      .chat-close{background:transparent;border:none;color:#fff;font-size:20px;cursor:pointer;}
      .chat-body{display:flex;flex:1;overflow:hidden;}

      /* --- Users --- */
      .chat-users{width:180px;background:#f5f5f5;border-right:1px solid #ddd;display:flex;flex-direction:column;}
      #chatUserSearch{padding:6px 10px;border:none;border-bottom:1px solid #ddd;border-radius:0;}
      .chat-users-list{overflow-y:auto;flex:1;font-size:14px;}
      .chat-user{padding:10px;border-bottom:1px solid #eee;cursor:pointer;display:flex;flex-direction:column;gap:2px;transition:background 0.2s;border-radius:8px;position:relative;}
      .chat-user:hover{background:#e0f7fa;}
      .chat-user.active{background:#b2ebf2;font-weight:600;}
      .chat-user .chat-badge{position:absolute;top:8px;right:8px;font-size:10px;background:#f44336;color:#fff;padding:2px 5px;border-radius:8px;}
      .chat-user small{color:#666;font-size:11px;}

      /* --- Thread --- */
      .chat-thread{flex:1;display:flex;flex-direction:column;height:100%;}
      .chat-thread-header{padding:10px;border-bottom:1px solid #ddd;font-weight:600;background:#fafafa;}
      .chat-messages{flex:1;overflow-y:auto;padding:12px;background:#f0f4f8;display:flex;flex-direction:column;gap:8px;scroll-behavior:smooth;}
      .chat-msg{padding:10px 14px;border-radius:16px;max-width:75%;word-wrap:break-word;box-shadow:0 2px 6px rgba(0,0,0,0.1);}
      .chat-msg.me{align-self:flex-end;background:linear-gradient(145deg,#90caf9,#64b5f6);color:#fff;}
      .chat-msg.them{align-self:flex-start;background:#fff;border:1px solid #ddd;color:#333;}
      .chat-msg small{display:block;font-size:10px;color:#666;margin-bottom:4px;}
      .chat-typing-indicator{padding:4px 10px;font-size:12px;color:#666;animation:blink 1s infinite;}
      @keyframes blink{0%,50%,100%{opacity:1}25%,75%{opacity:0.3}}

      /* --- Input --- */
      .chat-input{display:flex;gap:6px;padding:10px;border-top:1px solid #ddd;background:#fafafa;}
      .chat-input input{flex:1;border:1px solid #ccc;border-radius:12px;padding:10px;}
      .chat-input button{border:none;background:var(--primary-color);color:#fff;border-radius:12px;padding:0 12px;cursor:pointer;box-shadow:0 2px 4px rgba(0,0,0,0.1);}
    `;
    document.head.appendChild(style);

    // Event Listeners
    document.getElementById('chatOpenBtn').addEventListener('click',()=>this.toggle(true));
    document.getElementById('chatCloseBtn').addEventListener('click',()=>this.toggle(false));
    document.getElementById('chatSendBtn').addEventListener('click',()=>this.send());
    document.getElementById('chatMessageInput').addEventListener('keydown',e=>{if(e.key==='Enter'&&!e.shiftKey)this.send();});
    document.getElementById('chatUserSearch').addEventListener('input',()=>this.renderUserList());
  },

  async toggle(show){
    this.open=show;
    document.getElementById('chatPanel').style.display=show?'flex':'none';
    if(show){
      await this.loadUsers();
      if(!this.threadPoll){
        this.threadPoll=setInterval(()=>{if(this.selectedUser)this.loadThread()},5000);
      }
    } else{
      clearInterval(this.threadPoll);
      this.threadPoll=null;
    }
  },

  async loadUsers(){
    const box=document.getElementById('chatUsers');
    box.innerHTML='Loading users...';

    const [usersRes, unreadRes]=await Promise.all([API.getChatUsers(),API.getChatUnreadSummary()]);
    if(!usersRes.success){box.innerHTML='<div class="text-danger">Failed to load users</div>';return;}

    // Filter out current user
    this.allUsers=usersRes.users.filter(u=>u.email.toLowerCase()!==this.currentUser.email.toLowerCase());
    this.unreadMap=unreadRes.success?unreadRes.summary:{};

    // Add General room
    this.allUsers.unshift({
      email:'GENERAL',
      name:'General Room',
      role:'All',
    });

    this.renderUserList();
  },

  renderUserList(){
    const box=document.getElementById('chatUsers');
    const search=document.getElementById('chatUserSearch').value.toLowerCase();

    let html=this.allUsers.filter(u=>u.name.toLowerCase().includes(search)||u.email.toLowerCase().includes(search))
      .map(u=>{
        const count=this.unreadMap[u.email]||0;
        const emoji=u.email==='GENERAL'?'📢':'👤';
        return `<div class="chat-user" data-email="${u.email}" data-name="${u.name}">
          <span>${emoji} ${u.name}</span>
          <small>${u.role}</small>
          <small>${u.email}</small>
          ${count?`<span class="chat-badge">${count}</span>`:''}
        </div>`;
      }).join('');

    box.innerHTML=html;
    box.querySelectorAll('.chat-user').forEach(el=>el.addEventListener('click',()=>this.selectUser(el.dataset.email)));
  },

  async selectUser(email){
    this.selectedUser=email;
    document.querySelectorAll('.chat-user').forEach(el=>el.classList.remove('active'));
    const active=document.querySelector(`.chat-user[data-email="${email}"]`);
    if(active)active.classList.add('active');

    // Use display name in header
    const user=this.allUsers.find(u=>u.email===email);
    document.getElementById('chatThreadHeader').textContent=user?user.name:email;

    await this.loadThread();
    if(email!=='GENERAL') await API.markChatRead(email);
    await this.refreshUnreadBadge();
  },

  async loadThread(){
    const box=document.getElementById('chatMessages');
    box.innerHTML='';
    if(!this.selectedUser){box.innerHTML='<small>Select a user</small>';return;}

    const res=await API.getChatThread(this.selectedUser,80);
    if(!res.success){box.innerHTML=`<small>${res.error}</small>`;return;}

    const myEmail=this.currentUser.email.toLowerCase();
    box.innerHTML=res.messages.map(m=>{
      const cls=(m.from||'').toLowerCase()===myEmail?'me':'them';
      return `<div class="chat-msg ${cls}">
        <small>${m.from} • ${Utils.formatDateTime(m.timestamp)}</small>
        ${this.escapeHtml(m.message)}
      </div>`;
    }).join('');
    box.scrollTop=box.scrollHeight;
  },

  async send(){
    const input=document.getElementById('chatMessageInput');
    const text=(input.value||'').trim();
    if(!this.selectedUser)return Utils.showToast('Select a user','warning');
    if(!text)return;
    input.value='';

    const res=await API.sendChatMessage(this.selectedUser,text);
    if(!res.success)return Utils.showToast(res.error||'Failed','error');
    await this.loadThread();
  },

  async refreshUnreadBadge(){
    const badge=document.getElementById('chatBadge');
    const res=await API.getChatUnreadCount();
    if(!res.success)return;
    const unread=res.unread||0;
    badge.style.display=unread?'inline-block':'none';
    badge.textContent=unread>99?'99+':unread;
  },

  escapeHtml(str){return String(str).replaceAll('&','&amp;').replaceAll('<','&lt;').replaceAll('>','&gt;')}
};

document.addEventListener('DOMContentLoaded',()=>{if(Auth.isLoggedIn())ChatWidget.init();});
