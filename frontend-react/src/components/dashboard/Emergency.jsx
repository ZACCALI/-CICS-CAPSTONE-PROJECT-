import React, { useState, useEffect, useRef } from 'react';
import { useApp } from '../../context/AppContext';
import { useAuth } from '../../context/AuthContext';
import Modal from '../common/Modal';

const Emergency = () => {
  const { emergencyActive, toggleEmergency, emergencyHistory, clearEmergencyHistory, logActivity, systemState } = useApp();
  const { currentUser } = useAuth();
  const [showConfirm, setShowConfirm] = useState(false);
  const [showLockModal, setShowLockModal] = useState(false);
  
  // Safety Lock: Prevent instant deactivation while AI Script is playing
  const [isScriptPlaying, setIsScriptPlaying] = useState(false);

  useEffect(() => {
      if (emergencyActive) {
          setIsScriptPlaying(true);
          // Approximate length of emergency script "Attention... This is an emergency..."
          // Set to 14 seconds to be safe
          const timer = setTimeout(() => {
              setIsScriptPlaying(false);
          }, 14000); 
          return () => clearTimeout(timer);
      } else {
          setIsScriptPlaying(false);
      }
  }, [emergencyActive]);

  // Activator Logic
  // Check if history exists and has items
  const lastEmergency = emergencyHistory && emergencyHistory.length > 0 ? emergencyHistory[0] : null;
  const activator = lastEmergency ? lastEmergency.user : 'Unknown';
  // Check ownership
  const isOwner = (currentUser?.name || 'Admin') === activator;

  const handleActivate = () => {
    if (!currentUser || !currentUser.name) {
        alert("Unable to verify user identity. Please reload.");
        return;
    }
    setShowConfirm(false);
    toggleEmergency(currentUser.name, 'ACTIVATED');
  };

  return (
    <div className="space-y-6">
      <h2 className="text-2xl font-bold text-red-600 flex items-center">
        <i className="material-icons mr-3">emergency</i> Emergency Alert
      </h2>

      <div className={`border-l-4 rounded-lg p-6 shadow-sm transition-colors duration-500 ${emergencyActive ? 'bg-red-100 border-red-600 animate-pulse' : 'bg-red-50 border-red-500'}`}>
         <div className="flex items-start">
           <div className="flex-shrink-0">
             <i className="material-icons text-red-500 text-3xl">warning</i>
           </div>
           <div className="ml-4 flex-1">
             <h3 className="text-xl font-bold text-red-800">
                 {emergencyActive 
                    ? `EMERGENCY ALERT ACTIVE${lastEmergency ? ` (Started by ${activator})` : ''}` 
                    : 'Emergency Broadcast System'}
             </h3>
             <p className="mt-2 text-red-700 leading-relaxed">
               {emergencyActive 
                 ? 'System is currently broadcasting emergency signal to ALL ZONES. Normal operations are suspended. Alarm is sounding.' 
                 : 'In case of emergency, activate the alarm system immediately. This will broadcast an alert to all zones and override any current announcements.'}
             </p>
             <div className="mt-6">
                {emergencyActive ? (
                    isOwner ? (
                        /* OWNER: Allow Deactivation */
                        <div className="flex flex-col items-center">
                            <button 
                                onClick={() => {
                                    if (currentUser?.name) {
                                        toggleEmergency(currentUser.name, 'DEACTIVATED');
                                    }
                                }}
                                disabled={isScriptPlaying} // Lock during AI Voice
                                className={`font-bold py-3 px-8 rounded-lg shadow-lg transform transition flex items-center ${isScriptPlaying ? 'bg-gray-500 cursor-not-allowed text-gray-200' : 'bg-gray-800 hover:bg-gray-900 hover:scale-105 text-white'}`}
                            >
                                <i className="material-icons mr-2">power_settings_new</i> 
                                {isScriptPlaying ? 'BROADCASTING NOTICE...' : 'DEACTIVATE ALERT'}
                            </button>
                            {isScriptPlaying && (
                                <p className="text-xs text-red-600 mt-2 font-semibold animate-pulse">
                                    Wait for voice announcement to complete...
                                </p>
                            )}
                        </div>
                    ) : (
                        /* NON-OWNER: Show Status / Disabled Look */
                        <div className="relative group">
                            <button 
                                onClick={() => setShowLockModal(true)}
                                disabled={true}
                                className="bg-gray-400 text-gray-100 font-bold py-3 px-8 rounded-lg shadow cursor-not-allowed flex items-center w-full justify-center"
                            >
                                <i className="material-icons mr-2">lock</i> EMERGENCY ONGOING
                            </button>
                            <p className="text-center text-xs text-red-800 mt-2 font-semibold">
                                Protected: Only {activator} can deactivate.
                            </p>
                        </div>
                    )
                ) : (
                    <button 
                        onClick={() => setShowConfirm(true)}
                        className="bg-red-600 hover:bg-red-700 text-white font-bold py-3 px-8 rounded-lg shadow-lg transform transition hover:scale-105 flex items-center"
                    >
                        <i className="material-icons mr-2">notifications_active</i> ACTIVATE EMERGENCY ALERT
                    </button>
                )}
             </div>
           </div>
         </div>
      </div>

        <div className="bg-white rounded-xl shadow-sm border border-gray-100 p-6">
          <div className="flex justify-between items-center mb-4">
              <h3 className="text-lg font-semibold text-gray-700 flex items-center">
                <i className="material-icons mr-2">history</i> Emergency History
              </h3>
              {emergencyHistory.filter(item => item.user === currentUser?.name).length > 0 && (
                  <button onClick={() => clearEmergencyHistory(currentUser?.name)} className="text-xs text-red-500 hover:text-red-700 font-medium hover:underline">
                      Clear History
                  </button>
              )}
          </div>
          
          {emergencyHistory.length > 0 ? (
              <div className="space-y-4">
                  {emergencyHistory.filter(item => item.user === currentUser?.name).map((item) => (
                      <div key={item.id} className="flex items-center justify-between p-3 bg-gray-50 rounded-lg border-l-4 border-red-400">
                          <div>
                              <p className="font-bold text-gray-800 text-sm">{item.action}</p>
                              <p className="text-xs text-gray-500">{item.time}</p>
                          </div>
                          <div className="text-sm font-medium text-gray-600">
                              by {item.user}
                          </div>
                      </div>
                  ))}
              </div>
          ) : (
              <p className="text-gray-500 text-sm italic">No recent emergency alerts triggered.</p>
          )}
        </div>

       <Modal
          isOpen={showConfirm}
          onClose={() => setShowConfirm(false)}
          title="Confirm Emergency Alert"
          type="danger"
          footer={
             <>
                 <button 
                    onClick={() => setShowConfirm(false)}
                    className="px-4 py-2 border border-gray-300 rounded-lg text-gray-700 font-medium hover:bg-gray-50 bg-white"
                 >
                     Cancel
                 </button>
                 <button 
                    onClick={handleActivate}
                    className="px-4 py-2 bg-red-600 text-white rounded-lg font-medium hover:bg-red-700 shadow-lg"
                 >
                     Yes, Activate
                 </button>
             </>
          }
       >
           <div className="text-center">
               <div className="w-16 h-16 bg-red-100 rounded-full flex items-center justify-center mx-auto mb-4">
                    <i className="material-icons text-3xl text-red-600">warning</i>
               </div>
               <p className="text-gray-600">
                   Are you sure you want to activate the emergency alert system? This will override all current announcements and play an alarm to all connected devices.
               </p>
           </div>
       </Modal>

       {/* Lock Modal */}
       <Modal
           isOpen={showLockModal}
           onClose={() => setShowLockModal(false)}
           title="Emergency Locked"
           type="info"
           footer={
               <button 
                  onClick={() => setShowLockModal(false)}
                  className="px-6 py-2 bg-gray-800 text-white rounded-lg font-medium shadow-sm hover:bg-gray-900"
               >
                  Understood
               </button>
           }
       >
           <div className="text-center p-4">
               <i className="material-icons text-5xl text-gray-400 mb-4">lock</i>
               <p className="text-lg text-gray-700 font-semibold mb-2">Action Blocked</p>
               <p className="text-gray-600">
                   Emergency is currently active.
                   <br/>
                   Only the activator (<span className="font-bold text-red-600">{emergencyHistory[0]?.user || 'Unknown'}</span>) can deactivate it.
               </p>
           </div>
       </Modal>
    </div>
  );
};

export default Emergency;
