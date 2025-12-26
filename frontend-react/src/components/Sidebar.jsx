import React from 'react';

const Sidebar = ({ activeSection, setActiveSection, role = 'admin', isExpanded = false }) => {
  const menuItems = [
    { id: 'realtime', label: 'Broadcast', icon: 'campaign' },
    { id: 'schedule', label: 'Schedule', icon: 'schedule' },
    { id: 'emergency', label: 'Emergency', icon: 'emergency' },
    { id: 'history', label: 'Activity Logs', icon: 'history_edu' }, // Admin Only
    { id: 'upload', label: 'Uploads', icon: 'upload' },
    { id: 'manage-account', label: 'Account', icon: 'admin_panel_settings' },
  ];

  const items = role === 'admin' 
    ? menuItems 
    : menuItems.filter(i => i.id !== 'history' && i.id !== 'manage-account'); 

  // Safe handler
  const handleItemClick = (id) => {
      if (typeof setActiveSection === 'function') {
          setActiveSection(id);
      } else {
           console.error("setActiveSection is not a function");
      }
  };

  return (
    <nav className="h-full bg-white shadow-xl overflow-y-auto w-full relative z-50 animate-fade-in-left"> 
      
      <ul className="py-6 space-y-2 pb-24">
        {items.map((item) => (
          <li 
            key={item.id}
            onClick={() => handleItemClick(item.id)}
            className={`px-6 py-4 cursor-pointer flex items-center transition-all duration-300 relative group/item
              ${activeSection === item.id 
                ? 'bg-primary/10 text-primary font-bold border-r-4 border-primary' 
                : 'text-gray-600 hover:bg-gray-50 hover:text-primary'}`}
          >
            <i className={`material-icons mr-4 text-[24px] transition-colors ${activeSection === item.id ? 'text-primary' : 'text-gray-400 group-hover/item:text-primary'}`}>
                {item.icon}
            </i>
            <span className={`text-base font-bold whitespace-nowrap overflow-hidden transition-all duration-300 
                ${isExpanded 
                    ? 'opacity-100 w-auto ml-2' 
                    : 'opacity-0 w-0 group-hover/sidebar:opacity-100 group-hover/sidebar:w-auto group-hover/sidebar:ml-2 delay-75'
                }`}>
                {item.label}
            </span>
          </li>
        ))}
      </ul>
      

    </nav>
  );
};

export default Sidebar;
