import React, { useState, useEffect } from 'react';
import DashboardLayout from '../components/DashboardLayout';
import RealTime from '../components/dashboard/RealTime';
import Schedule from '../components/dashboard/Schedule';
import Emergency from '../components/dashboard/Emergency';
import Upload from '../components/dashboard/Upload';
import ManageAccount from '../components/dashboard/ManageAccount';
import HistoryLogs from '../components/dashboard/HistoryLogs';
import { useApp } from '../context/AppContext';
import { useAuth } from '../context/AuthContext';

const AdminDashboard = () => {
  // Persist active tab
  const [activeSection, setActiveSection] = useState(() => {
    return localStorage.getItem('admin_active_section') || 'realtime';
  });
  const { currentUser: user } = useAuth(); 
  const { schedules } = useApp();

  useEffect(() => {
    localStorage.setItem('admin_active_section', activeSection);
  }, [activeSection]);

  useEffect(() => {
    const handleNavChange = (e) => {
        if (e.detail) setActiveSection(e.detail);
    };

    window.addEventListener('nav-change', handleNavChange);
    return () => window.removeEventListener('nav-change', handleNavChange);
  }, []);

  const mySchedules = schedules.filter(s => s.user === user?.name);
  const totalAnnouncements = mySchedules.length;
  const pendingAnnouncements = mySchedules.filter(s => s.status === 'Pending').length;



  return (
    <DashboardLayout 
      activeSection={activeSection} 
      setActiveSection={setActiveSection} 
      user={user}
    >
      <div className="max-w-7xl mx-auto animation-fade-in pb-12">


        {/* Stats Grid: Request was for "total and upcoming is in one line and the system and devices is in one line".
            This implies a 2-column layout on most screens, rows:
            Row 1: Total | Upcoming
            Row 2: System | Devices
        */}
        <div className="grid grid-cols-1 md:grid-cols-3 gap-6 mb-8">
          
          {/* Row 1 */}
          <div className="bg-white p-6 rounded-2xl shadow-sm border border-gray-100 flex items-center hover:shadow-md transition-all duration-300">
            <div className="w-14 h-14 rounded-xl bg-blue-50 text-primary flex items-center justify-center mr-5 shadow-sm">
              <i className="material-icons text-3xl">campaign</i>
            </div>
            <div>
              <p className="text-gray-500 text-xs font-bold uppercase tracking-wider mb-1">Total Announcements</p>
              <h3 className="text-3xl font-bold text-gray-800">{totalAnnouncements}</h3>
            </div>
          </div>
          
          <div className="bg-white p-6 rounded-2xl shadow-sm border border-gray-100 flex items-center hover:shadow-md transition-all duration-300">
            <div className="w-14 h-14 rounded-xl bg-purple-50 text-purple-600 flex items-center justify-center mr-5 shadow-sm">
               <i className="material-icons text-3xl">schedule</i>
            </div>
            <div>
              <p className="text-gray-500 text-xs font-bold uppercase tracking-wider mb-1">Upcoming Schedule</p>
              <h3 className="text-3xl font-bold text-gray-800">{pendingAnnouncements}</h3>
            </div>
          </div>

          {/* Row 2 */}
          <div className="bg-white p-6 rounded-2xl shadow-sm border border-gray-100 flex items-center hover:shadow-md transition-all duration-300">
            <div className="w-14 h-14 rounded-xl bg-green-50 text-green-600 flex items-center justify-center mr-5 shadow-sm">
               <i className="material-icons text-3xl">check_circle</i>
            </div>
            <div>
              <p className="text-gray-500 text-xs font-bold uppercase tracking-wider mb-1">Device Status</p>
              <h3 className="text-2xl font-bold text-green-600">Online</h3>
            </div>
          </div>
        </div>

        {/* Content Area */}
        <div className="transition-all duration-300">
          <div className={activeSection === 'realtime' ? 'block' : 'hidden'}>
            <RealTime />
          </div>
          <div className={activeSection === 'schedule' ? 'block' : 'hidden'}>
            <Schedule />
          </div>
          <div className={activeSection === 'emergency' ? 'block' : 'hidden'}>
            <Emergency />
          </div>
          <div className={activeSection === 'history' ? 'block' : 'hidden'}>
            <HistoryLogs />
          </div>
          <div className={activeSection === 'upload' ? 'block' : 'hidden'}>
            <Upload />
          </div>
          <div className={activeSection === 'manage-account' ? 'block' : 'hidden'}>
            <ManageAccount />
          </div>
        </div>
      </div>
    </DashboardLayout>
  );
};

export default AdminDashboard;
