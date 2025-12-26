import React, { useState, useRef, useEffect } from 'react';
import { useApp } from '../../context/AppContext';
import { useAuth } from '../../context/AuthContext';
import Modal from '../common/Modal';

const Schedule = () => {
  const { schedules, addSchedule, updateSchedule, deleteSchedule, logActivity } = useApp();
  const { currentUser, getAllUsers } = useAuth();
  const [activeTab, setActiveTab] = useState('pending');
  
  // Modals
  const [showModal, setShowModal] = useState(false); // Form Modal
  const [showInfoModal, setShowInfoModal] = useState(false); // Alert Replacement
  const [infoMessage, setInfoMessage] = useState('');
  
  // Delete Modal
  const [showDeleteModal, setShowDeleteModal] = useState(false);
  const [scheduleToDelete, setScheduleToDelete] = useState(null);

  // Form State
  const [editId, setEditId] = useState(null);
  const [ownerFilter, setOwnerFilter] = useState('all'); // 'all', 'mine', 'others'
  const [userRoles, setUserRoles] = useState({}); // Map of username -> role
  const [isSubmitting, setIsSubmitting] = useState(false);
  const [isRecording, setIsRecording] = useState(false);
  const [audioBlob, setAudioBlob] = useState(null);
  const mediaRecorderRef = useRef(null);
  const recognitionRef = useRef(null);
  const audioPreviewRef = useRef(null); // Ref for audio element
  const [interimText, setInterimText] = useState('');
  
  const [formData, setFormData] = useState({
      message: '',
      date: '',
      time: '',
      repeat: 'once',
      zones: {
        'All Zones': false,
        'Admin Office': false,
        'Main Hall': false,
        'Library': false,
        'Classrooms': false
      }
  });

  // Init Speech Recognition
  useEffect(() => {
      if ('webkitSpeechRecognition' in window || 'SpeechRecognition' in window) {
          const SpeechRecognition = window.SpeechRecognition || window.webkitSpeechRecognition;
          recognitionRef.current = new SpeechRecognition();
          recognitionRef.current.continuous = true;
          recognitionRef.current.interimResults = true;
          
          
          recognitionRef.current.onresult = (event) => {
              let interim = '';
              let final = '';

              for (let i = event.resultIndex; i < event.results.length; ++i) {
                  if (event.results[i].isFinal) {
                      final += event.results[i][0].transcript;
                  } else {
                      interim += event.results[i][0].transcript;
                  }
              }
              
              if (final) {
                  setFormData(prev => ({
                      ...prev, 
                      message: (prev.message + ' ' + final).replace(/\s+/g, ' ').trim()
                  }));
                  setInterimText(''); 
              } else {
                  setInterimText(interim);
              }
          };
      }
      return () => {
          if (recognitionRef.current) {
              recognitionRef.current.abort();
          }
      };
  }, []);

  // Global Audio Stop Listener
  useEffect(() => {
      const handleStop = () => {
          if (audioPreviewRef.current) {
              audioPreviewRef.current.pause();
          }
      };
      window.addEventListener('stop-all-audio', handleStop);
      return () => window.removeEventListener('stop-all-audio', handleStop);
  }, []);

  // Fetch User Roles for Filtering (Admin Only)
  useEffect(() => {
    if (currentUser?.role === 'admin') {
        const fetchRoles = async () => {
            try {
                const users = await getAllUsers();
                if (users) {
                    const roleMap = {};
                    users.forEach(u => {
                        if (u.name) roleMap[u.name] = u.role;
                    });
                    setUserRoles(roleMap);
                }
            } catch (err) {
                console.error("Failed to fetch user roles for filtering:", err);
            }
        };
        fetchRoles();
    }
  }, [currentUser]);

  const handleZoneChange = (zone) => {
    if (zone === 'All Zones') {
        const newValue = !formData.zones['All Zones'];
        const newZones = {};
        Object.keys(formData.zones).forEach(k => newZones[k] = newValue);
        setFormData(prev => ({ ...prev, zones: newZones }));
    } else {
        setFormData(prev => ({
            ...prev, 
            zones: { ...prev.zones, [zone]: !prev.zones[zone] }
        }));
    }
  };

  const startRecording = async () => {
      try {
          // 1. Audio Recording
          const stream = await navigator.mediaDevices.getUserMedia({ audio: true });
          const mediaRecorder = new MediaRecorder(stream);
          mediaRecorderRef.current = mediaRecorder;
          const chunks = [];

          mediaRecorder.ondataavailable = (e) => chunks.push(e.data);
          mediaRecorder.onstop = () => {
              const blob = new Blob(chunks, { type: 'audio/webm' });
              setAudioBlob(blob);
              stream.getTracks().forEach(track => track.stop());
          };

          mediaRecorder.start();

          // 2. Speech Recognition (Transcription)
          if (recognitionRef.current) {
              try {
                recognitionRef.current.start();
              } catch (e) {
                console.warn("Recognition already started");
              }
          }

          setIsRecording(true);
      } catch (err) {
          console.error("Error accessing microphone:", err);
          setInfoMessage("Could not access microphone. Please ensure permissions are granted.");
          setShowInfoModal(true);
      }
  };

  const stopRecording = () => {
      if (mediaRecorderRef.current && isRecording) {
          mediaRecorderRef.current.stop();
      }
      if (recognitionRef.current) {
          recognitionRef.current.stop();
      }
      setIsRecording(false);
      setInterimText('');
  };

  const resetRecording = () => {
      setAudioBlob(null);
      // We do NOT plain clear text message, as user might want to keep the transcript
  };

  const startEdit = (schedule) => {
      setEditId(schedule.id);
      
      // Parse zones string back to object
      const zonesMap = {
          'All Zones': false,
          'Admin Office': false,
          'Main Hall': false,
          'Library': false,
          'Classrooms': false
      };
      
      if (schedule.zones) {
          schedule.zones.split(', ').forEach(z => {
              if (zonesMap.hasOwnProperty(z)) zonesMap[z] = true;
          });
      }

      setFormData({
          message: schedule.message,
          date: schedule.date,
          time: schedule.time,
          repeat: schedule.repeat || 'once',
          zones: zonesMap
      });
      setAudioBlob(schedule.audio || null);
      setShowModal(true);
  };

  const confirmDelete = (id) => {
      setScheduleToDelete(id);
      setShowDeleteModal(true);
  };

  const handleDelete = () => {
      if (scheduleToDelete) {
          deleteSchedule(scheduleToDelete, currentUser?.name);
          setShowDeleteModal(false);
          setScheduleToDelete(null);
      }
  };

  const handleSubmit = (e) => {
      e.preventDefault();
      if (!formData.date) {
          setInfoMessage("Please select a date.");
          setShowInfoModal(true);
          return;
      }

      if (!formData.time) {
          setInfoMessage("Please select a time.");
          setShowInfoModal(true);
          return;
      }

      const activeZones = Object.keys(formData.zones).filter(z => formData.zones[z] && z !== 'All Zones');
      
      if (!activeZones.length) {
          setInfoMessage("Select at least one zone");
          setShowInfoModal(true);
          return;
      }
      
      if (!formData.message && !audioBlob) {
           setInfoMessage("Please enter a message or record audio");
           setShowInfoModal(true);
           return;
      }

      const processSubmission = async () => {
          if (isSubmitting) return;
          setIsSubmitting(true);

          try {
              let audioString = null;
          if (audioBlob) {
              // Convert Blob to Base64
              const reader = new FileReader();
              const base64Promise = new Promise((resolve, reject) => {
                  reader.onloadend = () => {
                      const result = reader.result;
                      // Check size (~800KB limit for Firestore safety)
                      if (result.length > 800000) {
                          reject(new Error("Audio recording is too long (Limit: ~60 seconds)."));
                      } else {
                          resolve(result);
                      }
                  };
                  reader.onerror = reject;
                  reader.readAsDataURL(audioBlob); 
              });

              try {
                  audioString = await base64Promise;
              } catch (err) {
                  setInfoMessage(err.message);
                  setShowInfoModal(true);
                  return;
              }
          }

          const scheduleData = {
              message: formData.message, 
              date: formData.date,
              time: formData.time,
              repeat: formData.repeat,
              zones: activeZones.join(', '),
              status: 'Pending',
              type: audioBlob ? 'voice' : 'text',
              audio: audioString // Send Base64 string
          };

          if (editId) {
              await updateSchedule(editId, scheduleData, currentUser?.name);
          } else {
              await addSchedule(scheduleData, currentUser?.name);
          }
          
          setShowModal(false);
          resetForm();
      } catch (err) {
          // Handle backend 409 conflict specifically if your Context doesn't catch it
          if (err.response && err.response.status === 409) {
             setInfoMessage("Time slot busy. Please choose another time.");
             setShowInfoModal(true);
          }
      } finally {
          setIsSubmitting(false);
      }
    };
    
    processSubmission();
  };
  
  const resetForm = () => {
      setEditId(null);
      setFormData({
          message: '',
          date: '',
          time: '',
          repeat: 'once',
          zones: {
              'All Zones': false,
              'Admin Office': false,
              'Main Hall': false,
              'Library': false,
              'Classrooms': false
          }
      });
      setAudioBlob(null);
      setIsRecording(false);
  };

  const filteredSchedules = schedules.filter(s => {
      // 1. Ownership Filter
      const isAdmin = currentUser?.role === 'admin';
      const isOwner = s.user === currentUser?.name;

      if (!isAdmin && !isOwner) {
          return false;
      }
      
      // Admin Logic: Apply 'ownerFilter'
      if (isAdmin) {
          if (ownerFilter === 'mine' && s.user !== currentUser?.name) return false;
          if (ownerFilter === 'others') {
               // STRICT: Show ONLY if creator is NOT current admin AND role is 'user'
               if (s.user === currentUser?.name) return false; // Hide mine
               
               const creatorRole = userRoles[s.user];
               // If we know the role, enforce it must be 'user'. 
               // If unknown (e.g. deleted user logic?), we might default to show, but user asked for "only user schedule".
               // So if creatorRole exists and is NOT 'user', we hide it.
               if (creatorRole && creatorRole !== 'user') return false;
          }
      }

      // 2. Status/Tab Filter
      return activeTab === 'pending' ? s.status === 'Pending' : (s.status === 'Completed' || s.status === 'History');
  });

  return (
    <div className="space-y-6">
      
      {/* Header Section */}
      <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center gap-4">
          <h2 className="text-2xl font-bold text-gray-800 flex items-center">
            <i className="material-icons mr-3 text-primary">schedule</i> Scheduled Announcements
          </h2>
          
          {/* Admin: Button in Header */}
          {currentUser?.role === 'admin' && (
              <button 
                    onClick={() => setShowModal(true)}
                    className="flex items-center px-4 py-2 bg-green-600 hover:bg-green-700 text-white rounded-lg shadow font-medium text-sm transition-all w-full sm:w-auto justify-center"
                 >
                   <i className="material-icons text-sm mr-2">add</i> Add Schedule
              </button>
          )}
      </div>

      {/* Toolbar Section: Tabs & Filters */}
      <div className="flex flex-col md:flex-row justify-between items-center bg-white p-2 rounded-xl shadow-sm border border-gray-100 gap-4">
         
         {/* Left: Status Toggles */}
         <div className="flex bg-gray-100 p-1 rounded-lg w-full md:w-auto">
            <button 
              onClick={() => setActiveTab('pending')}
              className={`flex-1 md:flex-none px-6 py-2 rounded-md text-sm font-medium transition-all ${activeTab === 'pending' ? 'bg-white text-primary shadow-sm' : 'text-gray-500 hover:text-gray-700'}`}
            >
              Pending
            </button>
            <button 
              onClick={() => setActiveTab('history')}
              className={`flex-1 md:flex-none px-6 py-2 rounded-md text-sm font-medium transition-all ${activeTab === 'history' ? 'bg-white text-primary shadow-sm' : 'text-gray-500 hover:text-gray-700'}`}
            >
              History
            </button>
         </div>

         {/* Admin: Filters */ }
         {currentUser?.role === 'admin' && (
             <div className="flex bg-gray-100 p-1 rounded-lg w-full md:w-auto overflow-x-auto">
                 <button 
                    onClick={() => setOwnerFilter('all')}
                    className={`flex-1 md:flex-none px-4 py-2 text-xs font-medium rounded-md transition-all whitespace-nowrap ${ownerFilter === 'all' ? 'bg-white text-gray-800 shadow-sm' : 'text-gray-500 hover:text-gray-700'}`}
                 >
                    All Schedules
                 </button>
                 <button 
                    onClick={() => setOwnerFilter('mine')}
                    className={`flex-1 md:flex-none px-4 py-2 text-xs font-medium rounded-md transition-all whitespace-nowrap ${ownerFilter === 'mine' ? 'bg-white text-gray-800 shadow-sm' : 'text-gray-500 hover:text-gray-700'}`}
                 >
                    Created by Me
                 </button>
                 <button 
                    onClick={() => setOwnerFilter('others')}
                    className={`flex-1 md:flex-none px-4 py-2 text-xs font-medium rounded-md transition-all whitespace-nowrap ${ownerFilter === 'others' ? 'bg-white text-gray-800 shadow-sm' : 'text-gray-500 hover:text-gray-700'}`}
                 >
                    Created by Others
                 </button>
             </div>
         )}
         
         {/* User: Button in Toolbar */}
         {currentUser?.role !== 'admin' && (
              <button 
                    onClick={() => setShowModal(true)}
                    className="flex items-center px-4 py-2 bg-green-600 hover:bg-green-700 text-white rounded-lg shadow font-medium text-sm transition-all w-full sm:w-auto justify-center md:ml-auto"
                 >
                   <i className="material-icons text-sm mr-2">add</i> Add Schedule
              </button>
         )}
      </div>

      {/* Stats Section (Admin Only) */}
      {currentUser?.role === 'admin' && (
          <div className="grid grid-cols-2 md:grid-cols-4 gap-4 mb-6">
              <div className="bg-blue-50 p-4 rounded-xl border border-blue-100 flex flex-col items-center justify-center">
                  <span className="text-2xl font-bold text-blue-600">
                      {schedules.filter(s => s.user === currentUser?.name).length}
                  </span>
                  <span className="text-xs text-blue-400 uppercase font-bold tracking-wide mt-1">My Total</span>
              </div>
              <div className="bg-yellow-50 p-4 rounded-xl border border-yellow-100 flex flex-col items-center justify-center">
                  <span className="text-2xl font-bold text-yellow-600">
                      {schedules.filter(s => s.user === currentUser?.name && s.status === 'Pending').length}
                  </span>
                  <span className="text-xs text-yellow-400 uppercase font-bold tracking-wide mt-1">My Pending</span>
              </div>
               <div className="bg-purple-50 p-4 rounded-xl border border-purple-100 flex flex-col items-center justify-center">
                  <span className="text-2xl font-bold text-purple-600">
                      {schedules.length}
                  </span>
                  <span className="text-xs text-purple-400 uppercase font-bold tracking-wide mt-1">System Total</span>
              </div>
              <div className="bg-green-50 p-4 rounded-xl border border-green-100 flex flex-col items-center justify-center">
                  <span className="text-2xl font-bold text-green-600">
                      {schedules.filter(s => s.status === 'Pending').length}
                  </span>
                  <span className="text-xs text-green-400 uppercase font-bold tracking-wide mt-1">System Pending</span>
              </div>
          </div>
      )}

      <div className="bg-white rounded-xl shadow-sm border border-gray-100 overflow-hidden">
        {/* Table Header - Hidden on Mobile */}
        <div className="hidden md:grid grid-cols-12 gap-4 p-4 bg-gray-50 border-b border-gray-200 font-semibold text-gray-600 text-sm">
          <div className="col-span-1">No.</div>
          <div className={`${currentUser?.role === 'admin' ? 'col-span-3' : 'col-span-3'}`}>Message</div>
          {currentUser?.role === 'admin' && <div className="col-span-1">Created By</div>}
          <div className="col-span-2">Date & Time</div>
          <div className="col-span-1">Repeat</div>
          <div className="col-span-2">Zones</div>
          <div className="col-span-1">Status</div>
          <div className="col-span-1 text-center">Actions</div>
        </div>
        {/* List */}
        {filteredSchedules.length > 0 ? (
            <div className="divide-y divide-gray-100">
                {filteredSchedules.map((schedule, index) => (
                    <div key={schedule.id || index} className="grid grid-cols-1 md:grid-cols-12 gap-2 md:gap-4 p-4 items-start md:items-center hover:bg-gray-50 transition-colors text-sm text-gray-700">
                        {/* Mobile: Card Layout, Desktop: Table Row */}
                        <div className="md:col-span-1 font-mono text-gray-400 hidden md:block">#{index + 1}</div>
                        
                        <div className={`${currentUser?.role === 'admin' ? 'md:col-span-3' : 'md:col-span-3'} font-medium flex items-center`}>
                            <span className="md:hidden font-bold text-gray-500 mr-2">Message:</span>
                            {schedule.type === 'voice' && <i className="material-icons text-primary mr-2 text-sm">mic</i>}
                            <span className="break-words line-clamp-2 md:truncate" title={schedule.message}>
                                {schedule.message.length > 50 ? schedule.message.substring(0, 50) + '...' : schedule.message}
                            </span>
                        </div>
                        
                        {currentUser?.role === 'admin' && (
                            <div className="md:col-span-1 text-gray-500 text-xs flex md:block items-center">
                                <span className="md:hidden font-bold text-gray-500 mr-2 w-20">By:</span>
                                <span className="bg-gray-100 px-2 py-0.5 rounded text-gray-700 truncate block max-w-full" title={schedule.user}>
                                    {schedule.user === currentUser?.name ? 'Me' : schedule.user}
                                </span>
                            </div>
                        )}
                        
                        <div className="md:col-span-2 text-gray-500 flex md:block">
                            <span className="md:hidden font-bold text-gray-500 mr-2 w-20">When:</span>
                            <span>{schedule.date} <span className="text-xs ml-1 md:ml-0 md:block">{schedule.time}</span></span>
                        </div>

                         <div className="md:col-span-1 text-gray-500 flex md:block">
                            <span className="md:hidden font-bold text-gray-500 mr-2 w-20">Repeat:</span>
                            <span className={`px-2 py-0.5 rounded text-xs font-semibold border ${
                                schedule.repeat === 'daily' ? 'bg-purple-100 text-purple-700 border-purple-200' : 
                                (schedule.repeat === 'weekly' ? 'bg-indigo-100 text-indigo-700 border-indigo-200' : 'bg-gray-100 text-gray-600 border-gray-200')
                            }`}>
                                {schedule.repeat ? schedule.repeat.charAt(0).toUpperCase() + schedule.repeat.slice(1) : 'Once'}
                            </span>
                        </div>
                        
                        <div className="md:col-span-2">
                             <div className="flex md:block items-center">
                                <span className="md:hidden font-bold text-gray-500 mr-2 w-20">Zones:</span>
                                <span className="truncate text-xs bg-primary/10 text-primary px-2 py-1 rounded w-fit inline-block max-w-full">
                                    {schedule.zones}
                                </span>
                             </div>
                        </div>
                        
                        <div className="md:col-span-1 flex md:block items-center">
                             <span className="md:hidden font-bold text-gray-500 mr-2 w-20">Status:</span>
                             <span className={`px-2 py-1 rounded-full text-xs font-semibold ${schedule.status === 'Pending' ? 'bg-yellow-100 text-yellow-700' : 'bg-gray-100 text-gray-600'}`}>
                                {schedule.status}
                             </span>
                        </div>
                        
                        <div className="md:col-span-1 flex justify-end md:justify-center space-x-2 mt-2 md:mt-0 border-t md:border-t-0 pt-2 md:pt-0 border-gray-100">
                             <button 
                                onClick={() => startEdit(schedule)}
                                className="p-1 px-3 md:px-1 text-primary hover:bg-primary/10 rounded flex items-center md:inline-flex bg-primary/10 md:bg-transparent"
                             >
                                 <i className="material-icons text-sm mr-1 md:mr-0">edit</i> <span className="md:hidden text-xs">Edit</span>
                             </button>
                             <button 
                                onClick={() => confirmDelete(schedule.id)}
                                className="p-1 px-3 md:px-1 text-red-600 hover:bg-red-50 rounded flex items-center md:inline-flex bg-red-50 md:bg-transparent"
                             >
                                 <i className="material-icons text-sm mr-1 md:mr-0">delete</i> <span className="md:hidden text-xs">Delete</span>
                             </button>
                        </div>
                    </div>
                ))}
            </div>
        ) : (
            <div className="p-12 text-center text-gray-500 flex flex-col items-center">
            <i className="material-icons text-5xl mb-4 text-gray-300">event_note</i>
            <p>No {activeTab} announcements found.</p>
            </div>
        )}
      </div>

      <Modal
        isOpen={showModal}
        onClose={() => setShowModal(false)}
        title={editId ? "Edit Schedule" : "Schedule Announcement"}
        footer={
           <>
               <button 
                  onClick={() => setShowModal(false)}
                  className="px-4 py-2 text-gray-600 hover:bg-gray-100 rounded-lg font-medium"
               >
                   Cancel
               </button>
               <button 
                  onClick={handleSubmit}
                  disabled={isSubmitting}
                  className={`px-6 py-2 rounded-lg font-medium shadow-sm text-white ${isSubmitting ? 'bg-gray-400 cursor-not-allowed' : 'bg-primary hover:bg-primary-dark'}`}
               >
                   {isSubmitting ? "Scheduling..." : (editId ? "Update Schedule" : "Confirm Schedule")}
               </button>
           </>
        }
      >
        <form className="space-y-4 max-h-[60vh] md:max-h-[70vh] overflow-y-auto px-1">
            
            {/* Unified Input Section */}
            <div>
                <label className="block text-sm font-medium text-gray-700 mb-2">Message & Audio</label>
                
                <div className="relative">
                    <textarea 
                        required
                        value={formData.message + (interimText ? ' ' + interimText : '')}
                        onChange={e => setFormData({...formData, message: e.target.value})}
                        className="w-full p-3 border border-gray-200 rounded-lg focus:ring-primary focus:border-primary min-h-[100px] pr-12"
                        placeholder="Type message here or use microphone to speak..."
                    ></textarea>
                    
                    {/* Floating Mic Button inside textarea */}
                    <button
                        type="button"
                        onClick={isRecording ? stopRecording : startRecording}
                        className={`absolute bottom-3 right-3 p-2 rounded-full shadow-sm transition-all ${isRecording ? 'bg-red-500 animate-pulse text-white' : 'bg-gray-100 text-gray-600 hover:bg-primary/10 hover:text-primary'}`}
                        title={isRecording ? "Stop Recording" : "Start Recording"}
                    >
                        <i className="material-icons text-xl">{isRecording ? 'stop' : 'mic'}</i>
                    </button>
                </div>
                <p className="text-xs text-gray-500 mt-1">
                    {isRecording ? 'Listening... Speak clearly to transcribe.' : 'Click mic to record & transcribe automatically.'}
                </p>

                {/* Audio Blob Indicator */}
                {audioBlob && (
                    <div className="mt-3 flex items-center justify-between bg-green-50 p-3 rounded-lg border border-green-100">
                        <div className="flex items-center text-green-700">
                            <i className="material-icons mr-2">check_circle</i>
                            <span className="text-sm font-medium">Audio Recorded Successfully</span>
                        </div>
                        <div className="flex items-center space-x-2">
                             <audio ref={audioPreviewRef} src={audioBlob instanceof Blob ? URL.createObjectURL(audioBlob) : ''} controls className="h-8 w-32" />
                             <button type="button" onClick={resetRecording} className="text-red-500 hover:bg-red-50 p-1 rounded-full"><i className="material-icons text-sm">delete</i></button>
                        </div>
                    </div>
                )}
            </div>

            <div className="grid grid-cols-2 gap-4">
                <div>
                    <label className="block text-sm font-medium text-gray-700 mb-1">Date</label>
                    <input 
                        type="date" 
                        required
                        value={formData.date}
                        onChange={e => setFormData({...formData, date: e.target.value})}
                        className="w-full p-2 border border-gray-200 rounded-lg"
                    />
                </div>
                <div>
                    <label className="block text-sm font-medium text-gray-700 mb-1">Time</label>
                    <input 
                        type="time" 
                        required
                        value={formData.time}
                        onChange={e => setFormData({...formData, time: e.target.value})}
                        className="w-full p-2 border border-gray-200 rounded-lg"
                    />
                </div>
            </div>

            <div>
                <label className="block text-sm font-medium text-gray-700 mb-1">Repeat</label>
                <select 
                    value={formData.repeat}
                    onChange={e => setFormData({...formData, repeat: e.target.value})}
                    className="w-full p-2 border border-gray-200 rounded-lg bg-white"
                >
                    <option value="once">Once</option>
                    <option value="daily">Daily</option>
                    <option value="weekly">Weekly</option>
                </select>
            </div>

            <div>
                <label className="block text-sm font-medium text-gray-700 mb-2">Select Zones</label>
                <div className="grid grid-cols-2 gap-2 max-h-[150px] overflow-y-auto">
                    {Object.keys(formData.zones).map((zone) => (
                        <label key={zone} className="flex items-center space-x-2 text-sm text-gray-600 cursor-pointer">
                            <input 
                                type="checkbox" 
                                checked={formData.zones[zone]}
                                onChange={() => handleZoneChange(zone)}
                                className="rounded text-primary focus:ring-primary"
                            />
                            <span>{zone}</span>
                        </label>
                    ))}
                </div>
            </div>
        </form>
      </Modal>

      {/* Info/Alert Modal */}
      <Modal
        isOpen={showInfoModal}
        onClose={() => setShowInfoModal(false)}
        title="Notice"
        footer={
           <button 
              onClick={() => setShowInfoModal(false)}
              className="px-6 py-2 bg-primary text-white rounded-lg font-medium shadow-sm"
           >
               OK
           </button>
        }
      >
        <p className="text-gray-600">{infoMessage}</p>
      </Modal>

      {/* Delete Confirmation Modal */}
      <Modal
          isOpen={showDeleteModal}
          onClose={() => setShowDeleteModal(false)}
          title="Delete Schedule"
          type="danger"
          footer={
             <>
                <button 
                   onClick={() => setShowDeleteModal(false)}
                   className="px-4 py-2 border rounded-lg text-gray-700 hover:bg-gray-50 bg-white"
                >
                    Cancel
                </button>
                <button 
                   onClick={handleDelete}
                   className="px-4 py-2 bg-red-600 text-white rounded-lg font-medium hover:bg-red-700 shadow-lg"
                >
                    Delete
                </button>
             </>
          }
      >
          <p className="text-gray-600">Are you sure you want to delete this schedule? This action cannot be undone.</p>
      </Modal>
    </div>
  );
};

export default Schedule;
