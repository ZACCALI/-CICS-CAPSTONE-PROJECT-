import React, { useState } from 'react';
import Sidebar from './Sidebar';
import Header from './Header';

const DashboardLayout = ({ children, activeSection, setActiveSection, user }) => {
  // Mobile Sidebar
  const [mobileSidebarOpen, setMobileSidebarOpen] = useState(false);
  // Desktop Sidebar (Default Open)
  const [desktopSidebarOpen, setDesktopSidebarOpen] = useState(true);

  const toggleSidebar = () => {
      // Logic: If on mobile, toggle mobile sidebar. If desktop, toggle desktop sidebar.
      // We can use window width or simpler: just toggle both states, CSS handles visibility.
      // But better: Header calls this.
      if (window.innerWidth < 768) {
          setMobileSidebarOpen(!mobileSidebarOpen);
      } else {
          setDesktopSidebarOpen(!desktopSidebarOpen);
      }
  };

  return (
    <div className="min-h-screen bg-gray-50 flex flex-col transition-all duration-300">
      <Header 
        user={user} 
        toggleSidebar={toggleSidebar} 
      />

      <div className="flex flex-1 pt-16 relative">
        {/* Desktop Sidebar */}
        {/* We manipulate the width/transform based on desktopSidebarOpen */}
        {/* Desktop Sidebar - Hoverable Rail */}
        <div className={`hidden md:block fixed inset-y-0 left-0 pt-16 h-full bg-white border-r border-gray-100 z-20 transition-all duration-300 w-20 hover:w-64 group/sidebar shadow-lg ${desktopSidebarOpen ? 'translate-x-0' : '-translate-x-full'}`}>
             <Sidebar 
                activeSection={activeSection} 
                setActiveSection={setActiveSection} 
                role={user?.role}
                isExpanded={false}
             />
        </div>

        {mobileSidebarOpen && (
          <div 
            className="fixed inset-0 bg-transparent z-50 md:hidden"
            onClick={() => setMobileSidebarOpen(false)}
          ></div>
        )}

        {/* Mobile Sidebar */}
        <div className={`fixed top-16 bottom-0 left-0 w-64 bg-white z-[60] transform transition-transform duration-300 md:hidden ${mobileSidebarOpen ? 'translate-x-0' : '-translate-x-full'}`}>
           <Sidebar 
            activeSection={activeSection} 
            setActiveSection={(section) => {
              setActiveSection(section);
              setMobileSidebarOpen(false);
            }} 
            role={user?.role}
            isExpanded={true}
          />
        </div>

        {/* Main Content */}
        {/* Margin Left changes based on Desktop Sidebar State */}
        <main 
            className={`flex-1 p-6 overflow-y-auto h-[calc(100vh-64px)] transition-all duration-300 animate-fade-in ${desktopSidebarOpen ? 'md:ml-20' : 'md:ml-0'}`}
        >
          {children}
        </main>
      </div>
    </div>
  );
};

export default DashboardLayout;
