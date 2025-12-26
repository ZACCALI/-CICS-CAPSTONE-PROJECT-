import React, { useState, useRef, useEffect } from 'react';
import { useNavigate } from 'react-router-dom';
import { useAuth } from '../context/AuthContext';
import { useApp } from '../context/AppContext';

const Header = ({ toggleSidebar, user }) => {
  const navigate = useNavigate();
  const { currentUser, logout } = useAuth();
  const { notifications, markAllAsRead, clearAllNotifications, resetState } = useApp();
  
  const [showProfileMenu, setShowProfileMenu] = useState(false);
  const [showNotifications, setShowNotifications] = useState(false);
  const [showQuickActions, setShowQuickActions] = useState(false);

  const profileRef = useRef(null);
  const notifRef = useRef(null);
  const quickRef = useRef(null);

  useEffect(() => {
    const handleClickOutside = (event) => {
        if (profileRef.current && !profileRef.current.contains(event.target)) setShowProfileMenu(false);
        if (notifRef.current && !notifRef.current.contains(event.target)) setShowNotifications(false);
        if (quickRef.current && !quickRef.current.contains(event.target)) setShowQuickActions(false);
    };
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, []);

  const handleLogout = async () => {
    try {
        // Stop all audio/media before logging out
        if (resetState) resetState(currentUser?.name || 'User'); // Check if exists just in case
        
        await logout();
        navigate('/');
    } catch (e) {
        console.error("Logout failed", e);
        navigate('/'); // Force redirect anyway
    }
  };
  
  // Navigation Helper for Dropdowns to avoid prop drilling activeSection if we rely on route
  // Note: Sidebar uses parent state 'activeSection', but we are inside Header.
  // Best way is to assume DashboardLayout works on routes or we pass a callback.
  // Actually DashboardLayout handles routes? Let's check App.jsx. 
  // App.jsx renders Routes. Navigation creates URL change. 
  // But DashboardLayout.jsx expects `activeSection`. 
  // Let's rely on the Sidebar to pick up state? 
  // Or better: In AdminDashboard, we should probably listen to location?
  // Wait, the App uses simple state for sections in `AdminDashboard.jsx`. 
  // If we just navigate, we might break the "Single Page" feel if it's state-based.
  // However, I see `useNavigate` is used for logout.
  // The layout seems to control `activeSection`.
  
  // Let's check `AdminDashboard.jsx`.
  // It renders `DashboardLayout` and passes `activeSection`.
  // So changing URL won't change the section unless we have logic for it.
  // But wait! `AdminDashboard.jsx` handles the state `activeSection`.
  
  // Alternative: We can emit a custom event or use Context to change section?
  // Or just pass `setActiveSection` to Header?
  // `Header` is child of `DashboardLayout`. 
  // `DashboardLayout` doesn't pass `setActiveSection` to `Header`.
  
  // Let's modify Header to accept `onNavigate` or similar, OR
  // Since we don't have that prop drilled, let's just use `window.location.hash` or similar? 
  // No, let's look at `DashboardLayout`.
  // <Header user={user} toggleSidebar={...} />
  
  // To fix this properly, I need to pass `setActiveSection` from `DashboardLayout` to `Header`.
  // But `DashboardLayout` receives it from `AdminDashboard`.
  
  // For now, I will assume we can't easily change the props without editing 3 files.
  // However, the Quick Actions should probably just "work".
  // If I can't set state, I can maybe navigate to specific routes if they existed.
  // But this seems to be a "Tabs" based dashboard on one route `/dashboard`.
  
  // FIX: I will modify `DashboardLayout` to pass `setActiveSection` to `Header`.
  // Then I update `Header` to use it.
  
  // Wait, I am editing Header directly here.
  // I will add `setActiveSection` prop.
  
  const handleQuickAction = (section) => {
      // We need a way to tell the parent to switch tabs.
      // If we don't have the prop, we can dispatch a custom event?
      // Or we can assume the user will click the sidebar.
      
      // Let's try dispatching event as a quick fix if we don't want to refactor everything.
      // dispatchEvent(new CustomEvent('switch-section', { detail: section }));
      
      // But adding a prop is cleaner.
      if (typeof toggleSidebar === 'function' && toggleSidebar.length === 2) { 
         // Hacky check? No.
      }
      // I'll emit custom event for now to avoid prop drilling hell if not needed.
      // actually, let's just look at how `Sidebar` does it. `setActiveSection`.
      
      window.dispatchEvent(new CustomEvent('nav-change', { detail: section }));
  };

  const unreadCount = notifications.filter(n => n.unread).length;

  return (
    <header className="fixed top-0 left-0 right-0 h-16 bg-primary shadow-lg z-[70] flex items-center justify-between px-4 md:px-6 transition-colors">
      <div className="flex items-center">
        <button onClick={toggleSidebar} title="Menu" className="p-2 mr-3 text-white hover:bg-white/20 rounded-full transition-colors">
          <i className="material-icons">menu</i>
        </button>
        <div className="flex items-center space-x-2">
            <img src="/cics-logo.png" alt="Logo" className="h-8 w-8 brightness-0 invert" />
            <h1 className="text-xl font-bold text-white font-sans tracking-wide">CICS PA SYSTEM</h1>
        </div>
      </div>

      <div className="flex items-center space-x-2 md:space-x-4">
        {/* Notifications */}
        <div className="relative" ref={notifRef}>
          <button 
            title="Notifications"
            onClick={() => setShowNotifications(!showNotifications)}
            className="p-2 text-blue-100 hover:bg-white/20 rounded-full relative transition-colors focus:outline-none"
          >
            <i className="material-icons">notifications</i>
            {unreadCount > 0 && (
                <span className="absolute top-0 right-0 h-5 w-5 bg-red-500 rounded-full border-2 border-primary flex items-center justify-center text-xs text-white font-bold">{unreadCount}</span>
            )}
          </button>

          {showNotifications && (
              <div className="absolute right-0 top-full mt-2 w-80 bg-white rounded-xl shadow-2xl border border-gray-100 py-1 animation-fade-in-up z-50">
                  <div className="px-4 py-3 border-b border-gray-100 flex justify-between items-center bg-gray-50/50">
                      <span className="font-semibold text-gray-700">Notifications</span>
                      {unreadCount > 0 && <span className="text-xs bg-red-100 text-red-600 px-2 py-0.5 rounded-full font-bold">{unreadCount} New</span>}
                  </div>
                  <div className="max-h-[300px] overflow-y-auto">
                      {notifications.length > 0 ? (
                          <ul className="divide-y divide-gray-50">
                              {notifications.map(note => (
                                  <li key={note.id} className={`p-4 hover:bg-gray-50 transition-colors cursor-pointer border-l-2 ${note.unread ? 'bg-primary/5 border-primary' : 'border-transparent'}`}>
                                      <div className="flex items-start">
                                          <i className={`material-icons text-sm mr-3 mt-1 ${note.type === 'warning' ? 'text-orange-500' : 'text-primary'}`}>
                                              {note.type === 'warning' ? 'warning' : 'info'}
                                          </i>
                                          <div>
                                              <p className={`text-sm font-medium ${note.unread ? 'text-gray-900' : 'text-gray-600'}`}>{note.title}</p>
                                              <p className="text-xs text-gray-500 mt-1">{note.message}</p>
                                              <p className="text-[10px] text-gray-400 mt-2">{note.time}</p>
                                          </div>
                                      </div>
                                  </li>
                              ))}
                          </ul>
                      ) : (
                          <div className="p-8 text-center text-gray-400 text-sm">No notifications</div>
                      )}
                  </div>
                  <div className="px-4 py-2 border-t border-gray-100 flex justify-between items-center text-xs font-medium">
                      <button 
                        onClick={markAllAsRead}
                        className="text-primary hover:underline"
                      >
                        Mark all as read
                      </button>
                      <button 
                        onClick={clearAllNotifications}
                        className="text-gray-400 hover:text-red-500 hover:underline"
                      >
                        Clear all
                      </button>
                  </div>
              </div>
          )}
        </div>

        {/* Quick Actions (+) - Responsive: Desktop Header / Mobile FAB */}
        <div className="fixed bottom-6 right-6 z-50 md:static md:z-auto" ref={quickRef}>
           <button 
             title="Quick Actions"
             onClick={() => setShowQuickActions(!showQuickActions)}
             className={`bg-primary hover:bg-primary-dark text-white rounded-full p-2 w-14 h-14 md:w-10 md:h-10 flex items-center justify-center shadow-lg transition-all transform hover:scale-105 active:scale-95 ${showQuickActions ? 'rotate-45 bg-red-500 hover:bg-red-600' : ''}`}
           >
             <i className="material-icons text-[24px] md:text-[20px]">add</i>
           </button>
           
           {showQuickActions && (
               <div className="absolute right-0 bottom-full mb-2 md:bottom-auto md:top-full md:mt-2 w-48 bg-white rounded-xl shadow-2xl border border-gray-100 py-1 animation-fade-in-up z-50 origin-bottom-right md:origin-top-right">
                   <div className="px-4 py-2 text-xs font-semibold text-gray-500 uppercase tracking-wider bg-gray-50">Quick Actions</div>
                   <button 
                      onClick={() => handleQuickAction('realtime')}
                      className="w-full text-left px-4 py-2 text-sm text-gray-700 hover:bg-gray-50 flex items-center transition-colors"
                   >
                       <i className="material-icons text-lg mr-2 text-green-600">mic</i> New Broadcast
                   </button>
                   <button 
                      onClick={() => handleQuickAction('schedule')}
                      className="w-full text-left px-4 py-2 text-sm text-gray-700 hover:bg-gray-50 flex items-center transition-colors"
                   >
                       <i className="material-icons text-lg mr-2 text-primary">event</i> Schedule
                   </button>
                   <button 
                      onClick={() => handleQuickAction('upload')}
                      className="w-full text-left px-4 py-2 text-sm text-gray-700 hover:bg-gray-50 flex items-center transition-colors"
                   >
                       <i className="material-icons text-lg mr-2 text-primary">cloud_upload</i> Upload File
                   </button>
               </div>
           )}
        </div>

        {/* Profile */}
        <div className="relative" ref={profileRef}>
          <div 
            title="Profile"
            className="h-10 w-10 rounded-full overflow-hidden border-2 border-white/50 cursor-pointer hover:ring-2 hover:ring-white/30 transition-all shadow-sm"
            onClick={() => setShowProfileMenu(!showProfileMenu)}
          >
            <img 
              src={currentUser?.avatar || `https://ui-avatars.com/api/?name=${currentUser?.name || 'User'}&background=2563EB&color=fff`} 
              alt="Profile" 
              className="w-full h-full object-cover"
            />
          </div>

          {showProfileMenu && (
            <div className="absolute right-0 top-full mt-2 w-56 bg-white rounded-xl shadow-2xl border border-gray-100 py-1 animation-fade-in-up z-50">
              <div className="px-4 py-4 border-b border-gray-100 bg-gray-50/50">
                <p className="text-sm font-bold text-gray-800 truncate">{currentUser?.name}</p>
                <p className="text-xs text-secondary truncate">{currentUser?.email}</p>
                <span className={`mt-2 inline-block px-2 py-0.5 rounded text-[10px] font-bold uppercase tracking-wide ${currentUser?.role === 'admin' ? 'bg-purple-100 text-purple-700' : 'bg-primary/10 text-primary'}`}>
                    {currentUser?.role || 'Guest'}
                </span>
              </div>
              
              <div className="py-1">
                  <button 
                      onClick={() => {
                          window.dispatchEvent(new CustomEvent('nav-change', { detail: 'manage-account' }));
                          setShowProfileMenu(false);
                      }}
                      className="w-full text-left px-4 py-2 text-sm text-gray-700 hover:bg-gray-50 flex items-center transition-colors"
                  >
                      <i className="material-icons text-lg mr-3 text-gray-400">person</i> My Profile
                  </button>
              </div>

              <div className="border-t border-gray-100 pt-1">
                  <button 
                    onClick={handleLogout}
                    className="w-full text-left px-4 py-2 text-sm text-red-600 hover:bg-red-50 flex items-center transition-colors font-medium"
                  >
                    <i className="material-icons text-lg mr-3">exit_to_app</i> Logout
                  </button>
              </div>
            </div>
          )}
        </div>
      </div>
    </header>
  );
};

export default Header;
