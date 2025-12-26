import React, { createContext, useContext, useState, useEffect, useRef } from 'react';
import api from '../api/axios';
import { doc, onSnapshot, collection, query, orderBy, limit } from "firebase/firestore";
import { db } from '../firebase';

const AppContext = createContext();
export const useApp = () => useContext(AppContext);

export const AppProvider = ({ children }) => {
  // Announcements
  const [schedules, setSchedules] = useState([]);

  const [notifications, setNotifications] = useState([
     { id: 1, title: 'System Update Available', message: 'New version 2.3.1 is ready.', time: '2 hours ago', unread: true, type: 'info' },
     { id: 2, title: 'Library Speaker Offline', message: 'Disconnected for 45 minutes.', time: 'Yesterday', unread: true, type: 'warning' }
  ]);

  // Files
  const [files, setFiles] = useState([]); // Init empty, fetch from API

  // Emergency State
  const [emergencyActive, setEmergencyActive] = useState(false);
  const [emergencyHistory, setEmergencyHistory] = useState([]);

  // Global Activity Logs
  const [activityLogs, setActivityLogs] = useState([]);

  // Global Broadcast State (Persistent across navigation)
  const [broadcastActive, setBroadcastActive] = useState(false);
  // Global Zones (Persistent)
  const [zones, setZones] = useState({
    'All Zones': false,
    'Admin Office': false,
    'Main Hall': false,
    'Library': false,
    'Classrooms': false
  });
  
  // Track specific System State Details (Mode, Active Task User, etc)
  const [systemState, setSystemState] = useState({});

  // Initial Fetch & Listeners
  useEffect(() => {
    // 1. Emergency System Listener
    const emergencyRef = doc(db, "emergency", "status");
    const unsubEmergency = onSnapshot(emergencyRef, (docSnap) => {
        if (docSnap.exists()) {
            const data = docSnap.data();
            setEmergencyActive(data.active);
            setEmergencyHistory(data.history || []);
        } else {
            setEmergencyActive(false);
            setEmergencyHistory([]);
        }
    });


    
    // 2. Schedules Listener (Real-time)
    const schedulesQuery = query(collection(db, "schedules"));
    const unsubSchedules = onSnapshot(schedulesQuery, (snapshot) => {
        const list = snapshot.docs.map(doc => ({ id: doc.id, ...doc.data() }));
        setSchedules(list);
    }, (error) => {
        console.error("Schedules sync error:", error);
    });

    // 3. Activity Logs Listener
    const logsQuery = query(collection(db, "logs"), orderBy("timestamp", "desc"), limit(50));
    const unsubLogs = onSnapshot(logsQuery, (snapshot) => {
        const list = snapshot.docs.map(doc => {
             const data = doc.data();
             let dateObj = new Date();
             if (data.timestamp && typeof data.timestamp.toDate === 'function') {
                 dateObj = data.timestamp.toDate();
             } else if (data.time) {
                 dateObj = new Date(data.time);
             }
             
             return {
                 id: doc.id,
                 ...data,
                 time: dateObj.toLocaleString('en-US', { 
                    year: 'numeric', 
                    month: 'short', 
                    day: 'numeric', 
                    hour: '2-digit', 
                    minute: '2-digit',
                    second: '2-digit'
                 })
             };
        });
        setActivityLogs(list);
    }, (error) => {
        console.error("Logs sync error:", error);
    });

    return () => {
        unsubEmergency();
        // unsubSystem(); // We didn't fully implement it in this block
        unsubSchedules();
        unsubLogs();
    };
  }, []); // End of mount effect

  // 4. Fetch Files from Backend
  const fetchFiles = async () => {
      try {
          const res = await api.get('/files/');
          setFiles(res.data);
      } catch (e) {
          console.error("Failed to fetch files", e);
      }
  };

  useEffect(() => {
      fetchFiles();
  }, []);

  // Siren Audio Ref (Singleton)
  const emergencyAudioRef = useRef(null);
  const sirenIntervalRef = useRef(null);
  const lastBroadcastTaskId = useRef(null);
  const broadcastStartingRef = useRef(false); // Grace period flag

  const playEmergencySiren = () => {
      // Prevent double-start
      if (emergencyAudioRef.current) return;

      try {
          const AudioContext = window.AudioContext || window.webkitAudioContext;
          const ctx = new AudioContext();
          emergencyAudioRef.current = ctx;

          const oscillator = ctx.createOscillator();
          const gainNode = ctx.createGain();

          oscillator.type = 'sine';
          oscillator.frequency.value = 600; 
          
          // SAFETY: Limit volume to 10%
          gainNode.gain.value = 0.1; 
          
          oscillator.connect(gainNode);
          gainNode.connect(ctx.destination);
          
          oscillator.start();
          
          // Siren Pulse Effect
          let isHigh = false;
          sirenIntervalRef.current = setInterval(() => {
              if (ctx.state === 'closed') return;
              const now = ctx.currentTime;
              const freq = isHigh ? 600 : 900; 
              oscillator.frequency.setValueAtTime(freq, now);
              isHigh = !isHigh;
          }, 800); 

      } catch (e) {
          console.error("Siren start failed", e);
      }
  };

  const stopEmergencySiren = () => {
      if (sirenIntervalRef.current) {
          clearInterval(sirenIntervalRef.current);
          sirenIntervalRef.current = null;
      }
      
      if (emergencyAudioRef.current) {
          try {
              emergencyAudioRef.current.close();
          } catch(e) {
              console.error("Siren close error", e);
          }
          emergencyAudioRef.current = null;
      }
  };

  // Centralized Emergency Audio Effect
  useEffect(() => {
      if (emergencyActive) {
          // Play ONLY if not already playing (handled by ref check inside function but good to be explicit)
          playEmergencySiren();
      } else {
          stopEmergencySiren();
      }

      // Cleanup on unmount (App close)
      return () => stopEmergencySiren();
  }, [emergencyActive]);
  
  // Track currently playing task to prevent re-triggering
  const currentTaskIdRef = useRef(null);
  const systemAudioRef = useRef(null); // For file playback

  const stopSystemPlayback = () => {
      // 1. Stop Audio Object
      if (systemAudioRef.current) {
          systemAudioRef.current.pause();
          systemAudioRef.current = null;
      }
      // 2. Stop TTS
      if ('speechSynthesis' in window) {
          window.speechSynthesis.cancel();
      }
      currentTaskIdRef.current = null;
  };

  const playSystemTask = async (task) => {
      if (!task || !task.data) return;
      if (currentTaskIdRef.current === task.id) return; // Already playing this task

      // Stop previous if any
      stopSystemPlayback();
      currentTaskIdRef.current = task.id;

      // Resolve Type (Root takes precedence, fallback to data check)
      const type = task.type || task.data?.type || 'text';

      // Ignore Background tasks (Content is just metadata/title, handled by local player)
      if (type === 'BACKGROUND' || type === 'background') return;

      console.log("Starting System Task:", task);

      try {
          
          if (type === 'voice' && task.data.audio) {
              // Play Base64 Audio
              const audioSrc = task.data.audio.startsWith('data:') 
                  ? task.data.audio 
                  : `data:audio/webm;base64,${task.data.audio}`;
              
              const audio = new Audio(audioSrc);
              systemAudioRef.current = audio;
              
              audio.onended = () => {
                  console.log("Task Audio Ended");
                  // Notify Backend to Clear State
                  api.post('/realtime/complete', { task_id: task.id })
                     .catch(err => console.error("Failed to complete task:", err));
                  
                  // Clear local state
                  stopSystemPlayback();
              };
              
              await audio.play();
              
          } else if (task.data.message || task.data.content) {
                  // Text to Speech
              if ('speechSynthesis' in window) {
                  const utterance = new SpeechSynthesisUtterance(task.data.message || task.data.content);
                  // MUTE FRONTEND TTS (Backend PowerShell handles audio)
                  // We keep utterance running so 'onend' triggers task completion
                  utterance.volume = 0; 
                  utterance.rate = 1.0;
                  
                  utterance.onend = () => {
                      console.log("TTS Ended (Frontend Logic)");
                      api.post('/realtime/complete', { task_id: task.id })
                         .catch(err => console.error("Failed to complete task:", err));
                      stopSystemPlayback();
                  };
                  
                  // SYNC FIX: Backend plays Intro Chime (~5s) before Text.
                  // We must delay Frontend 'Completion Timer' so we don't kill backend audio early.
                  // Default Delay: 5s. Emergency: 0s.
                  const delay = (task.type === 'emergency' || task.priority === 100) ? 0 : 5500;
                  
                  console.log(`Starting Silent Completion Timer in ${delay}ms`);
                  setTimeout(() => {
                      window.speechSynthesis.speak(utterance);
                  }, delay);
              }
          }
      } catch (err) {
          console.error("Failed to play system task:", err);
      }
  };

  // Global System State Listener (The Executor)
  useEffect(() => {
      const systemRef = doc(db, "system", "state");
      const unsubSystem = onSnapshot(systemRef, (docSnap) => {
          if (docSnap.exists()) {
              const data = docSnap.data();
              setSystemState(data);
              
              // 1. EMERGENCY OVERRIDE
              if (data.mode === 'EMERGENCY') {
                  // Strictly stop all local activity (Mic, Music, System)
                  stopAllAudio();
                  return; 
              }

              // 2. Active Task Playback
              if (data.active_task) {
                  // If new task is High Priority (Voice/Text), stop any low-priority Music.
                  if (data.active_task.type !== 'BACKGROUND' && data.active_task.priority !== 10) {
                      window.dispatchEvent(new Event('stop-all-audio'));
                  }

                  // CHECK PREEMPTION: If we are broadcasting but the Active Task ID doesn't match ours, we lost the lock.
                  if (mediaStreamRef.current && lastBroadcastTaskId.current && data.active_task.id !== lastBroadcastTaskId.current) {
                       console.warn("Broadcast preempted.");
                       stopBroadcast(); 
                       // alert("Your broadcast was interrupted by another user or higher priority event."); // REMOVED: False positives
                  }

                  playSystemTask(data.active_task);
              } else {
                  // No active task
                  
                  // If we thought we were broadcasting, but system says NO task, we must have been killed/timed out.
                  // Check grace period to avoid race condition.
                  if (mediaStreamRef.current && !broadcastStartingRef.current) {
                       console.warn("Broadcast ended by system.");
                       stopBroadcast();
                  }

                  // Stop System if playing
                  if (currentTaskIdRef.current) {
                      stopSystemPlayback();
                  }
              }
          }
      });
      return () => {
          unsubSystem();
          stopSystemPlayback();
      };
  }, [emergencyActive]); // Depend on emergency to re-eval if needed, or just keep it simple. 

  // Removed manual fetchSchedules as it is now real-time

  // Removed LocalStorage Logic
  // useEffect(() => { ... }, [files]);


  // Methods
  const addSchedule = async (schedule, user = 'Admin') => {
      try {
          // Audio is now expected to be a Base64 string if present
          // We no longer strip it.
          const { ...payload } = schedule; 
          // Include user in payload for logging
          const data = { ...payload, user };
          
          const res = await api.post('/scheduled/', data);
          // Real-time listener will handle the update
          return res.data;
      } catch (e) {
          console.error("Add schedule failed", e);
          throw e; // Rethrow for UI handling (e.g. 409 Conflict)
      }
  };

  const updateSchedule = async (id, updatedData, user = 'Admin') => {
       try {
          const { ...payload } = updatedData;
          const data = { ...payload, user };
          await api.put(`/scheduled/${id}`, data);
          setSchedules(prev => prev.map(s => s.id === id ? { ...s, ...updatedData } : s));
       } catch (e) {
           console.error("Update schedule failed", e);
           throw e; // Rethrow for UI handling
       }
  };

  const deleteSchedule = async (id, user = 'Admin') => {
      try {
          await api.delete(`/scheduled/${id}?user=${encodeURIComponent(user)}`);
          setSchedules(prev => prev.filter(s => s.id !== id));
      } catch (e) {
          console.error("Delete schedule failed", e);
      }
  };

  const addFile = (fileDetails) => {
      setFiles(prev => [fileDetails, ...prev]);
  };

  const deleteFile = async (filename) => {
      // Optimistic
      setFiles(prev => prev.filter(f => f.name !== filename));
      try {
          await api.delete(`/files/${filename}?user=${encodeURIComponent('Admin')}`);
      } catch (e) {
          console.error("Delete file failed", e);
          fetchFiles(); // Revert on fail
      }
  };

  const toggleEmergency = async (user = 'Admin', action = 'TOGGLE') => {
      // OPTIMISTIC KILL: Stop everything immediately if activating
      if (action === 'ACTIVATED') {
          stopAllAudio();
      }

      try {
          const res = await api.post('/emergency/toggle', { user, action });
          setEmergencyActive(res.data.active);
          setEmergencyHistory(res.data.history);
      } catch (e) {
          console.error("Emergency toggle failed", e);
      }
  };

  const clearEmergencyHistory = async (user) => {
      // Optimistic update: Remove immediately from UI
      setEmergencyHistory(prev => user ? prev.filter(h => h.user !== user) : []);
      
      try {
          const url = user ? `/emergency/history?user=${encodeURIComponent(user)}` : '/emergency/history';
          await api.delete(url);
      } catch (e) {
          console.error("Failed to clear emergency history", e);
          // Optional: we could revert here if needed, but snapshot listener often corrects state
      }
  };

  const logActivity = async (user, action, type, details) => {
      // Optimistic local
      const tempId = Date.now() + Math.random();
      const newLog = {
          id: tempId,
          user: user || 'Unknown',
          action, 
          type, 
          details,
          time: new Date().toLocaleString()
      };
      setActivityLogs(prev => [newLog, ...prev]);
      
      // Send to backend
      try {
          const res = await api.post('/realtime/log', {
              user: user || 'Unknown',
              type,
              action,
              details
          });
          // Update local ID with real ID if needed, 
          // or ideally we just refetch logs occasionally.
          // For session logging, we need the REAL ID to update it.
          if (res.data.id) {
              return res.data.id;
          }
      } catch (e) {
          console.error("Log failed", e);
      }
      return null;
  };

  const updateLog = async (id, updateData) => {
      if (!id) return;
      
      // Optimistic update
      setActivityLogs(prev => prev.map(log => 
          log.id === id ? { ...log, ...updateData } : log
      ));

      try {
          await api.put(`/realtime/log/${id}`, updateData);
      } catch(e) {
          console.error("Update log failed", e);
      }
  };

  const deleteLog = async (id, user = 'Admin') => {
      try {
          await api.delete(`/realtime/log/${id}?user=${encodeURIComponent(user)}`);
          setActivityLogs(prev => prev.filter(log => log.id !== id));
      } catch (e) {
          console.error("Delete log failed", e);
      }
  };

  const deleteLogs = async (ids, user = 'Admin') => {
      try {
          // Parallel delete
          await Promise.all(ids.map(id => api.delete(`/realtime/log/${id}?user=${encodeURIComponent(user)}`)));
          setActivityLogs(prev => prev.filter(log => !ids.includes(log.id)));
      } catch (e) {
          console.error("Bulk delete failed", e);
      }
  };

  // Broadcast Refs
  const mediaStreamRef = useRef(null);
  const audioContextRef = useRef(null);
  const [broadcastStream, setBroadcastStream] = useState(null);
  
  // Ref to track current user for beacon
  const currentBroadcasterRef = useRef(null);
  
  // NEW: State to track "Chime Phase"
  const [broadcastPreparing, setBroadcastPreparing] = useState(false);

  // DEAD MAN SWITCH: Stop broadcast/music on refresh/close
  useEffect(() => {
    const handleUnload = () => {
        // ... (existing unload logic) ...
        const possibleUser = currentBroadcasterRef.current || 'Admin'; 
        const urlBg = `${api.defaults.baseURL || 'http://localhost:8000'}/realtime/stop?user=${encodeURIComponent(possibleUser)}&type=background`;
        fetch(urlBg, { method: 'POST', keepalive: true, headers: { 'Content-Type': 'application/json' } });
    };
    window.addEventListener('beforeunload', handleUnload);
    return () => window.removeEventListener('beforeunload', handleUnload);
  }, [broadcastActive]); 

  const startBroadcast = async (user, zonesObj) => {
      try {
          // 1. Register with Backend Controller (Start Chime)
          const zoneList = Object.keys(zonesObj).filter(k => zonesObj[k]);
          
          const res = await api.post('/realtime/start', {
              user: user || 'Unknown',
              zones: zoneList,
              type: 'voice'
          });

          // Store User for Cleanup
          currentBroadcasterRef.current = user || 'Unknown';

          // Store Task ID and Set Grace Period
          lastBroadcastTaskId.current = res.data.task_id;
          broadcastStartingRef.current = true;
          setTimeout(() => { broadcastStartingRef.current = false; }, 5000);

          // 2. Preemptively stop all other audio 
          stopSystemPlayback();
          if ('speechSynthesis' in window) window.speechSynthesis.cancel();
          window.dispatchEvent(new Event('stop-all-audio'));
          document.querySelectorAll('audio').forEach(el => {
               try { el.pause(); el.currentTime = 0; } catch (e) {}
          });
          
          // 3. Start Microphone
          const stream = await navigator.mediaDevices.getUserMedia({ audio: true });
          mediaStreamRef.current = stream;
          setBroadcastStream(stream);

          // Audio Context for monitoring
          const AudioContext = window.AudioContext || window.webkitAudioContext;
          const audioCtx = new AudioContext();
          audioContextRef.current = audioCtx;
          
          const source = audioCtx.createMediaStreamSource(stream);
          source.connect(audioCtx.destination);

          // 4. NEW: Wait for Chime (Simulated 5s)
          setBroadcastPreparing(true); // UI shows "PLAYING CHIME..."
          await new Promise(resolve => setTimeout(resolve, 5000));
          setBroadcastPreparing(false);

          setBroadcastActive(true); // UI shows "STOP BROADCAST"
          
          return true;
      } catch (err) {
          console.error("Broadcast Start Failed:", err);
          // If 409, it means system busy
          if (err.response && err.response.status === 409) {
              alert("System is busy. Another broadcast is active.");
          }
          return false;
      }
  };

  const stopBroadcast = async (user = 'System') => {
      try {
          const taskId = lastBroadcastTaskId.current;
          if (!taskId) {
               // If no task ID, we likely aren't broadcasting.
               // sending a blind /stop command kills background music!
               console.log("Local stopBroadcast called (No active task ID). Skipping API call.");
          } else {
               let url = `/realtime/stop?user=${encodeURIComponent(user)}`;
               url += `&task_id=${encodeURIComponent(taskId)}`;
               await api.post(url);
          }
      } catch (e) {
          console.error("Failed to stop backend broadcast:", e);
      }
      
      lastBroadcastTaskId.current = null;

      if (mediaStreamRef.current) {
          mediaStreamRef.current.getTracks().forEach(track => track.stop());
          mediaStreamRef.current = null;
      }
      if (audioContextRef.current) {
          audioContextRef.current.close();
          audioContextRef.current = null;
      }
      setBroadcastStream(null);
      setBroadcastActive(false);
  };

  const markAllAsRead = () => {
      setNotifications(prev => prev.map(n => ({ ...n, unread: false })));
  };

  const clearAllNotifications = () => {
      setNotifications([]);
  };

  const stopAllAudio = () => {
      // 1. Stop System Playback (Schedules/TTS)
      stopSystemPlayback();

      // 2. Stop Broadcast / Mic
      stopBroadcast();
      
      // 3. Stop Text-to-Speech
      if ('speechSynthesis' in window) {
          window.speechSynthesis.cancel();
      }

      // 4. Stop any file audio (Global Event for React Components)
      window.dispatchEvent(new Event('stop-all-audio'));

      // 4. AGGRESSIVE FAILSAFE: Kill any rogue HTML5 Audio/Video elements
      document.querySelectorAll('audio').forEach(el => {
          try { el.pause(); el.currentTime = 0; } catch (e) {}
      });
      document.querySelectorAll('video').forEach(el => {
           try { el.pause(); el.currentTime = 0; } catch (e) {}
      });
  };

  const resetState = (userName = 'System') => {
      stopAllAudio();
      
      // Force Deactivate Emergency if active
      if (emergencyActive) {
          toggleEmergency(userName, 'DEACTIVATED');
      }

      setSchedules([]);
      setActivityLogs([]);
      setNotifications([]);
  };

  const value = {
      schedules,
      addSchedule,
      updateSchedule,
      deleteSchedule,
      notifications,
      markAllAsRead, 
      clearAllNotifications,
      files,
      addFile,
      deleteFile,
      emergencyActive,
      toggleEmergency,
      emergencyHistory,
      clearEmergencyHistory, // Exported
      activityLogs, // Exported for UI

      logActivity,
      updateLog, // New export
      deleteLog,
      deleteLogs, // New export
      broadcastActive,
      broadcastPreparing, // NEW STATE EXPORT
      startBroadcast,
      stopBroadcast,
      broadcastStream,
      activeTask: systemState?.active_task,
      systemState, // Export Full State for Locking Logic
      zones, setZones,
      
      stopAllAudio, // New
      resetState    // New
  };

  return (
    <AppContext.Provider value={value}>
      {children}
    </AppContext.Provider>
  );
};

