import React, { useState, useEffect, useRef } from 'react';
import { useApp } from '../../context/AppContext';
import { useAuth } from '../../context/AuthContext';
import api from '../../api/axios'; // Import API for direct Text Broadcast calls
import Modal from '../common/Modal';

const RealTime = () => {
  const { addSchedule, logActivity, updateLog, broadcastActive, broadcastPreparing, startBroadcast, stopBroadcast, broadcastStream, zones, setZones, emergencyActive, stopAllAudio, systemState } = useApp();
  const { currentUser, loading: authLoading } = useAuth();
  const [currentLogId, setCurrentLogId] = useState(null); // Track session log
  
  // Lock Logic
  const isSystemLoading = !systemState || authLoading;
  const activeTask = systemState?.active_task;

  // Lock if: Loading OR Active Task belongs to someone else
  const isLockedByOther = isSystemLoading || (activeTask && 
                          (activeTask.type === 'voice' || activeTask.type === 'text') && 
                          activeTask.data?.user !== (currentUser?.name || 'Admin'));

  const lockingUser = isSystemLoading ? 'System' : (isLockedByOther ? (activeTask?.data?.user || 'Another User') : null);
  
  // Ref for readable start time
  const startTimeStrRef = useRef('');
  const startTimeRef = useRef(null);
  
  const audioContextRef = useRef(null);
  const analyserRef = useRef(null);
  const canvasRef = useRef(null);
  const animationFrameRef = useRef(null);
  
  const [textMessage, setTextMessage] = useState('');
  const [voices, setVoices] = useState([]);
  const [selectedVoice, setSelectedVoice] = useState('');
  
  const [showModal, setShowModal] = useState(false);
  const [modalMessage, setModalMessage] = useState('');
  
  // Debounce Lock
  const [isSubmitting, setIsSubmitting] = useState(false);

  // Init Voices
  useEffect(() => {
     const initVoices = () => {
         const vs = window.speechSynthesis.getVoices();
         setVoices(vs);
         if (vs.length > 0 && !selectedVoice) setSelectedVoice(vs[0].name);
     };

     initVoices();
     window.speechSynthesis.onvoiceschanged = initVoices;
  }, []);

  const handleZoneChange = (zone) => {
    // Logic for Broadasting safety (Min 1 active)
    if (broadcastActive) {
        if (zone === 'All Zones') {
             // If unchecking All Zones, but we are broadcasting, check if it clears everything
             if (zones['All Zones']) {
                 setModalMessage('Cannot unselect all zones while broadcasting.');
                 setShowModal(true);
                 return;
             }
        } else {
             // If unchecking a specific zone, ensure it's not the LAST one
             if (zones[zone]) {
                 // Count active zones (excluding 'All Zones' key logic)
                 const activeCount = Object.keys(zones).filter(k => k !== 'All Zones' && zones[k]).length;
                 if (activeCount <= 1) {
                     setModalMessage('At least one zone must remain active during broadcast.');
                     setShowModal(true);
                     return;
                 }
             }
        }
    }

    if (zone === 'All Zones') {
        const newValue = !zones['All Zones'];
        const newZones = {};
        Object.keys(zones).forEach(k => newZones[k] = newValue);
        setZones(newZones);
    } else {
        // Toggle specific zone
        const newValue = !zones[zone];
        const newZones = { ...zones, [zone]: newValue };
        
        // Sync 'All Zones' checkbox
        // If unchecking any zone, 'All Zones' must be false.
        // If checking a zone, check if ALL others are now true -> set 'All Zones' to true.
        
        if (!newValue) {
            newZones['All Zones'] = false;
        } else {
            // Check if all OTHER real zones are true
            const allOthersTrue = Object.keys(newZones).filter(k => k !== 'All Zones' && k !== zone).every(k => newZones[k]);
            if (allOthersTrue) {
                newZones['All Zones'] = true;
            }
        }
        
        setZones(newZones);
    }
  };

  // Handle Visualizer locally when stream is active
  useEffect(() => {
    if (broadcastActive && broadcastStream) {
        // Setup Visualizer
        const AudioContext = window.AudioContext || window.webkitAudioContext;
        const audioCtx = new AudioContext();
        audioContextRef.current = audioCtx;
        const analyser = audioCtx.createAnalyser();
        analyserRef.current = analyser;
        
        try {
            const source = audioCtx.createMediaStreamSource(broadcastStream);
            source.connect(analyser); // Connect for visualization
            analyser.fftSize = 256;
            drawVisualizer();
        } catch(e) { console.error(e); }

        const handleStop = () => {
            if (audioContextRef.current) audioContextRef.current.close();
            if (animationFrameRef.current) cancelAnimationFrame(animationFrameRef.current);
        };
        window.addEventListener('stop-all-audio', handleStop);

        return () => {
            window.removeEventListener('stop-all-audio', handleStop);
            if (audioContextRef.current) audioContextRef.current.close();
            if (animationFrameRef.current) cancelAnimationFrame(animationFrameRef.current);
        };
    }
  }, [broadcastActive, broadcastStream]);
  const drawVisualizer = () => {
      if (!analyserRef.current || !canvasRef.current) return;
      const canvas = canvasRef.current;
      const ctx = canvas.getContext('2d');
      const bufferLength = analyserRef.current.frequencyBinCount;
      const dataArray = new Uint8Array(bufferLength);
      
      const draw = () => {
          animationFrameRef.current = requestAnimationFrame(draw);
          analyserRef.current.getByteFrequencyData(dataArray);
          
          ctx.clearRect(0, 0, canvas.width, canvas.height); // Clear instead of fill for transparent overlay
          
          const barWidth = (canvas.width / bufferLength) * 2.5;
          let barHeight;
          let x = 0;
          
          for (let i = 0; i < bufferLength; i++) {
              barHeight = dataArray[i] / 2;
              // Gradient Color
              ctx.fillStyle = `rgba(239, 68, 68, ${barHeight / 100})`; // Red based intensity
              ctx.fillRect(x, canvas.height - barHeight, barWidth, barHeight);
              x += barWidth + 1;
          }
      };
      
      draw();
  };
  const toggleBroadcast = async () => {
    if (isSubmitting) return; // Debounce
    
    if (emergencyActive) {
        setModalMessage("Emergency Alert is currently active. All broadcasts are suspended.");
        setShowModal(true);
        return;
    }

    if (broadcastActive) {
        stopBroadcast(currentUser?.name || 'Admin');
        
        if (currentLogId) {
             const endTimeStr = new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
             
             updateLog(currentLogId, {
                 action: 'Voice Broadcast Session',
                 details: `Voice Broadcast (Start: ${startTimeStrRef.current} - End: ${endTimeStr})`
             });
             setCurrentLogId(null);
        } else {
             // Fallback
             logActivity(currentUser?.name, 'Stopped Voice Broadcast', 'Voice', 'Microphone deactivated');
        }
        return;
    }

    if (!Object.values(zones).some(z => z)) {
        setModalMessage('Please select at least one zone before broadcasting.');
        setShowModal(true);
        return;
    }
    
    // Lock Button
    setIsSubmitting(true);
    
    try {
        // Pass User and Zones to Context
        const success = await startBroadcast(currentUser?.name || 'Admin', zones);
        
        if (success) {
            startTimeRef.current = Date.now();
            startTimeStrRef.current = new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
            
            const logId = await logActivity(currentUser?.name, 'Active Voice Broadcast', 'Voice', 'Microphone is active...');
            setCurrentLogId(logId);
        } 
    } catch (e) {
        console.error("Broadcast toggle error", e);
    } finally {
        setIsSubmitting(false); // Unlock
    }
  };

  const handleTextBroadcast = async () => {
    if (isSubmitting) return; // Debounce
    
    if (emergencyActive) {
        setModalMessage("Emergency Alert is currently active. Text announcements are disabled.");
        setShowModal(true);
        return;
    }

    if (broadcastActive) {
        setModalMessage('Cannot send text announcement while voice broadcast is active.');
        setShowModal(true);
        return;
    }

    if (!textMessage.trim()) {
        setModalMessage('Please enter a message.');
        setShowModal(true);
        return;
    }

    if (!Object.values(zones).some(z => z)) {
        setModalMessage('Please select at least one zone.');
        setShowModal(true);
        return;
    }

    // Stop existing audio locally immediately for responsiveness
    stopAllAudio();
    
    // Lock Button
    setIsSubmitting(true);
    
    // Send to Backend Controller (Global Sync & Priority)
    const activeZonesList = Object.keys(zones).filter(z => zones[z]);

    try {
        await api.post('/realtime/start', {
             user: currentUser?.name || 'Admin',
             zones: activeZonesList,
             type: 'text',
             content: textMessage
        });

        // Log to Global History
        logActivity(
            currentUser?.name, 
            'Broadcasted Text', 
            'Text', 
            `Message: "${textMessage}" to ${activeZonesList.filter(z => z !== 'All Zones').join(', ')}`
        );
        
        setTextMessage('');
        
    } catch (err) {
        console.error("Text Broadcast Failed", err);
        setModalMessage(
            err.response && err.response.status === 409 
            ? "System Busy: Another broadcast is active." 
            : "Failed to broadcast message."
        );
        setShowModal(true);
    } finally {
        setIsSubmitting(false); // Unlock
    }
  };

  // Determine if I am currently broadcasting text (Server-side)
  const isSpeaking = systemState?.active_task?.type === 'text' && 
                     systemState?.active_task?.data?.user === (currentUser?.name || 'Admin');

  return (
    <div className="space-y-6">
      <h2 className="text-2xl font-bold text-gray-800 flex items-center mb-6">
        <i className="material-icons mr-3 text-primary">campaign</i> Real-Time Announcement
      </h2>

      {emergencyActive && (
          <div className="bg-red-100 border-l-4 border-red-500 text-red-700 p-4 rounded shadow-sm flex items-center animate-pulse mb-6">
              <i className="material-icons text-2xl mr-3">warning</i>
              <div>
                  <p className="font-bold">Emergency Alert Active</p>
                  <p className="text-sm">Broadcast features are temporarily disabled.</p>
              </div>
          </div>
      )}

      {isLockedByOther && !emergencyActive && (
          <div className="bg-orange-100 border-l-4 border-orange-500 text-orange-800 p-4 rounded shadow-sm flex items-center mb-6 animate-fade-in">
              <i className="material-icons text-2xl mr-3">lock</i>
              <div>
                  <p className="font-bold">System Busy</p>
                  <p className="text-sm">
                      <span className="font-semibold">{lockingUser}</span> is currently broadcasting. Please wait.
                  </p>
              </div>
          </div>
      )}

      {/* Live Broadcast */}
      <div className="bg-white rounded-xl shadow-sm border border-gray-100 p-6">
        <h3 className="text-lg font-semibold text-gray-700 mb-4 flex items-center">
          <i className="material-icons mr-2 text-green-600">mic</i> Live Broadcast (Microphone)
        </h3>
        
        <div className="flex flex-col items-center justify-center p-8 bg-gray-50 rounded-lg border border-dashed border-gray-200 relative overflow-hidden">
           
           {/* Visualizer Canvas overlay */}
           {broadcastActive && (
               <canvas ref={canvasRef} width="600" height="200" className="absolute inset-x-0 bottom-0 w-full h-[150px] opacity-50 pointer-events-none" />
           )}

           <div className={`w-24 h-24 rounded-full flex items-center justify-center mb-4 transition-all duration-300 relative z-10 ${broadcastActive ? 'bg-red-100 animate-pulse' : 'bg-gray-200'}`}>
             <i className={`material-icons text-4xl ${broadcastActive ? 'text-red-500' : 'text-gray-400'}`}>mic</i>
           </div>
           
           <button 
             onClick={toggleBroadcast}
             disabled={isLockedByOther || emergencyActive || isSubmitting || broadcastPreparing}
             className={`px-8 py-3 rounded-full font-bold shadow-lg transition-all transform hover:scale-105 relative z-10 ${broadcastActive ? 'bg-red-500 hover:bg-red-600 text-white' : ((isLockedByOther || emergencyActive || isSubmitting || broadcastPreparing) ? 'bg-gray-300 text-gray-500 cursor-not-allowed transform-none hover:scale-100 shadow-none' : 'bg-primary hover:bg-primary-dark text-white')}`}
           >
             {isSystemLoading || isSubmitting || broadcastPreparing ? (broadcastPreparing ? 'PLAYING CHIME...' : (isSubmitting ? 'PROCESSING...' : 'CONNECTING...')) : (emergencyActive ? 'EMERGENCY ACTIVE' : (isLockedByOther ? 'SYSTEM BUSY' : (broadcastActive ? 'STOP BROADCAST' : 'START BROADCAST')))}
           </button>
           
           <p className="mt-4 text-sm text-gray-500 relative z-10">
             {broadcastActive ? 'Broadcasting live...' : 'Ready to broadcast'}
           </p>
           
           <div className="mt-2 flex items-center text-xs text-gray-400 bg-white px-3 py-1 rounded-full border border-gray-100 relative z-10">
             <i className="material-icons text-sm mr-1">info</i>
             Audio monitoring enabled (You will hear yourself)
           </div>
        </div>
      </div>

      {/* Text Broadcast */}
      <div className="bg-white rounded-xl shadow-sm border border-gray-100 p-6">
        <h3 className="text-lg font-semibold text-gray-700 mb-4 flex items-center">
          <i className="material-icons mr-2 text-primary">text_fields</i> Text Announcement
        </h3>

        <textarea 
          value={textMessage}
          onChange={(e) => setTextMessage(e.target.value)}
          className="w-full p-4 border border-gray-200 rounded-lg focus:ring-2 focus:ring-primary/20 focus:border-primary outline-none transition-all resize-y min-h-[120px] mb-4"
          placeholder="Type your announcement here..."
        ></textarea>

        <div className="flex flex-col md:flex-row gap-4 justify-between items-center">
          <div className="w-full md:w-auto">
             <label className="text-sm font-medium text-gray-600 mr-3">Voice:</label>
             <select 
                value={selectedVoice}
                onChange={(e) => setSelectedVoice(e.target.value)}
                className="p-2 border border-gray-200 rounded-md bg-white text-gray-700 outline-none focus:border-primary max-w-full md:max-w-[200px]"
             >
               {voices.map(v => (
                   <option key={v.name} value={v.name}>{v.name}</option>
               ))}
               {voices.length === 0 && <option>Default AI Voice</option>}
             </select>
          </div>
          
          <button 
             onClick={handleTextBroadcast}
             disabled={isLockedByOther || broadcastActive || emergencyActive || isSubmitting || isSpeaking}
             className={`w-full md:w-auto px-6 py-2.5 rounded-lg shadow-md font-medium flex items-center justify-center transition-all ${isLockedByOther || broadcastActive || emergencyActive || isSubmitting || isSpeaking ? 'bg-gray-300 text-gray-500 cursor-not-allowed shadow-none' : 'bg-primary hover:bg-primary-dark text-white'}`}
           >
             <i className="material-icons mr-2">volume_up</i> {emergencyActive ? 'Emergency Active' : (isLockedByOther ? 'System Busy' : (broadcastActive ? 'Voice Active' : (isSubmitting ? 'Processing...' : (isSpeaking ? 'Speaking...' : 'Broadcast Text'))))}
           </button>
        </div>
      </div>

      {/* Zones */}
      <div className="bg-white rounded-xl shadow-sm border border-gray-100 p-6">
        <h3 className="text-lg font-semibold text-gray-700 mb-4">Select Zones:</h3>
        <div className="grid grid-cols-2 md:grid-cols-3 lg:grid-cols-5 gap-4">
          {Object.keys(zones).map((label, idx) => (
             <label key={idx} className="flex items-center space-x-3 p-3 border border-gray-100 rounded-lg hover:bg-gray-50 cursor-pointer transition-colors">
               <input 
                 type="checkbox" 
                 checked={zones[label]}
                 onChange={() => handleZoneChange(label)}
                 className="w-5 h-5 text-primary rounded focus:ring-primary border-gray-300" 
               />
               <span className="text-gray-700 font-medium">{label}</span>
             </label>
          ))}
        </div>
      </div>

      <Modal
        isOpen={showModal}
        onClose={() => setShowModal(false)}
        title="Notice"
        footer={
           <button 
              onClick={() => setShowModal(false)}
              className="px-6 py-2 bg-primary text-white rounded-lg font-medium shadow-sm"
           >
              OK
           </button>
        }
      >
        <p className="text-gray-600">{modalMessage}</p>
      </Modal>
    </div>
  );
};

export default RealTime;
