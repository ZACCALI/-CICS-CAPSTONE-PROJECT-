import React, { useRef, useState, useEffect } from 'react';
import { useApp } from '../../context/AppContext';
import { useAuth } from '../../context/AuthContext';
import Modal from '../common/Modal';
import api from '../../api/axios';

const Upload = () => {
  const { files, addFile, deleteFile, logActivity, updateLog, emergencyActive, systemState } = useApp();
  const { currentUser } = useAuth();
  const fileInputRef = useRef(null);
  
  // Audio Player State
  const [playingId, setPlayingId] = useState(null);
  const [currentLogId, setCurrentLogId] = useState(null); 
  const [currentTime, setCurrentTime] = useState(0);
  const [duration, setDuration] = useState(0);
  
  const audioRef = useRef(new Audio());
  const startTimeRef = useRef(null);
  const logIdRef = useRef(null);

  // Sync Log ID for callbacks
  useEffect(() => { logIdRef.current = currentLogId; }, [currentLogId]);

  // Format Helper
  const formatTime = (time) => {
    if (!time || isNaN(time)) return "0:00";
    const minutes = Math.floor(time / 60);
    const seconds = Math.floor(time % 60);
    return `${minutes}:${seconds < 10 ? '0' + seconds : seconds}`;
  };

  // Player Handlers
  const handleTimeUpdate = () => {
      setCurrentTime(audioRef.current.currentTime);
  };

  const handleLoadedMetadata = () => {
      setDuration(audioRef.current.duration);
  };

  const handleEnded = () => {
      // Auto-Next Logic
      const currentIndex = files.findIndex(f => f.id === playingId);
      if (currentIndex !== -1 && currentIndex < files.length - 1) {
          playSound(files[currentIndex + 1].id);
      } else {
          stopPlayback();
      }
  };

  const stopPlayback = async () => {
      try {
          await api.post(`/realtime/stop?user=${encodeURIComponent(currentUser?.name || 'Admin')}&type=background`);
      } catch (e) { /* Ignore if already stopped */ }

      if (audioRef.current) {
          audioRef.current.pause();
          audioRef.current.currentTime = 0;
      }
      setPlayingId(null);
      setCurrentTime(0);
  };

  const playNext = () => {
      const currentIndex = files.findIndex(f => f.id === playingId);
      if (currentIndex !== -1 && currentIndex < files.length - 1) {
          playSound(files[currentIndex + 1].id);
      }
  };

  const playPrev = () => {
      const currentIndex = files.findIndex(f => f.id === playingId);
      if (currentIndex !== -1 && currentIndex > 0) {
          playSound(files[currentIndex - 1].id);
      }
  };

  const handleSeek = (e) => {
      const time = parseFloat(e.target.value);
      setCurrentTime(time);
      if (audioRef.current) {
          audioRef.current.currentTime = time;
      }
  };

  // Listeners
  useEffect(() => {
      const audio = audioRef.current;
      
      // These handlers are already defined above, but we need to attach them to the audio element.
      // Re-defining them here would create new functions on each render, which is not ideal.
      // Instead, we attach the top-level handlers.
      audio.addEventListener('timeupdate', handleTimeUpdate);
      audio.addEventListener('loadedmetadata', handleLoadedMetadata);
      audio.addEventListener('ended', handleEnded);
      
      const handleStopGlobal = () => {
          // INTERRUPTION LOGIC: Just pause, don't clear state (so we can resume)
          // But if it's a Manual Stop from another user, we might want to stop? 
          // For now, assume global stop event means "Quiet Please", so pause.
          if (audioRef.current) audioRef.current.pause();
      };
      
      window.addEventListener('stop-all-audio', handleStopGlobal);

      return () => {
          audio.removeEventListener('timeupdate', handleTimeUpdate);
          audio.removeEventListener('loadedmetadata', handleLoadedMetadata);
          audio.removeEventListener('ended', handleEnded);
          window.removeEventListener('stop-all-audio', handleStopGlobal);
          audio.pause();
      };
  }, [files, playingId]); 

  // RESUME LOGIC (Watch System State)
  useEffect(() => {
      if (!playingId || !audioRef.current) return;

      // If Emergency, stay paused
      if (emergencyActive) {
          audioRef.current.pause();
          return;
      }

      // Check System State
      if (systemState?.active_task) {
          const task = systemState.active_task;
          const currentUserName = currentUser?.name || 'Admin';
          const isMyUser = task.data?.user === currentUserName;
          
          if (task.type === 'BACKGROUND' || task.priority === 10) {
              // It's Background Mode. If it's MY file (or generic background), RESUME.
             if (isMyUser && audioRef.current.paused) {
                 console.log("[Resume Logic] Auto-Resuming Background Music for", currentUserName);
                 audioRef.current.play().catch(e => console.error("Resume failed", e));
             } else if (!isMyUser) {
                 // If background but NOT mine, ensure I am paused/stopped?
                 // We rely on standard behavior.
             }
          } else {
              // Higher priority task active (Voice/Schedule) -> PAUSE
              if (!audioRef.current.paused) {
                   console.log(`[Resume Logic] Pausing for High Priority Task (${task.type})`);
                   audioRef.current.pause();
              }
          }
      } else {
          // System IDLE. 
          // If we are playing locally but system says IDLE, it means our backend task finished/died.
          // We should probably stop? Or maybe we haven't started yet?
          // If we have a playingId, we expect to be playing. 
          // But if backend says IDLE, maybe we should re-request? 
          // Or just Pause to be safe.
          if (!audioRef.current.paused) {
               // Optional: audioRef.current.pause(); 
          }
      }
  }, [systemState, emergencyActive, playingId, currentUser]);

  // Main Play Function
  const playSound = async (id) => {
      if (emergencyActive) {
          setErrorMessage("Emergency Alert is currently active. Audio playback is disabled.");
          setShowErrorModal(true);
          return;
      }
      
      const fileToPlay = files.find(f => f.id === id);
      if (!fileToPlay) return;

      // Check Backend Lock First
      if (playingId !== id || audioRef.current.paused) {
          try {
              console.log("[Upload] Requesting playback lock for:", fileToPlay.name);
              await api.post('/realtime/start', {
                  user: currentUser?.name || 'Admin',
                  zones: ['All Zones'], 
                  type: 'background',
                  content: fileToPlay.name
              });
          } catch (err) {
              if (err.response && err.response.status === 409) {
                  setErrorMessage("System Busy: Another broadcast or audio is playing.");
                  setShowErrorModal(true);
                  return;
              }
              console.error("Lock Request Failed:", err);
          }
      }

      if (playingId === id) {
          // Toggle Pause/Play
          if (audioRef.current.paused) {
              try {
                  await audioRef.current.play();
              } catch (err) {
                  console.error("Playback failed:", err);
                  setErrorMessage("Playback failed: " + err.message);
                  setShowErrorModal(true);
              }
          } else {
              // OPTIMIZATION: Release Lock on Pause so others can use system
              console.log("[Upload] Pausing and Releasing Lock");
              audioRef.current.pause();
              try {
                  // Calls stop but keeps local PlayingId so we can resume
                  await api.post(`/realtime/stop?user=${encodeURIComponent(currentUser?.name || 'Admin')}&type=background`);
              } catch (e) {
                  console.warn("Failed to release lock on pause", e);
              }
          }
      } else {
          // Play New
          // Play New
          // New: Use 'url' from backend (Static File)
          if (fileToPlay.url) {
             const fullUrl = `${api.defaults.baseURL}${fileToPlay.url}`;
             console.log("Playing from URL:", fullUrl);
             audioRef.current.src = fullUrl;
             
             try {
                 await audioRef.current.play();
                 setPlayingId(id);
                 startTimeRef.current = Date.now();
                 
                 // Log activity
             try {
                 const newLogId = await logActivity(
                     currentUser?.name || 'Admin',
                     'Music Session',
                     'Music',
                     `${fileToPlay.name} (Start: ${new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })})`
                 );
                 setCurrentLogId(newLogId);
             } catch (logErr) {
                 console.error("Logging failed", logErr);
             }
             } catch (err) {
                 console.error("Playback load failed:", err);
                 setErrorMessage("Could not play audio: " + err.message);
                 setShowErrorModal(true);
             }
          } else if (fileToPlay.content) {
              // Fallback for legacy local files (if any persist in cache)
              audioRef.current.src = fileToPlay.content;
              // ... (Start play)
              await audioRef.current.play();
              setPlayingId(id);
          }
      }
  };
 
  // Modal
  const [showDeleteModal, setShowDeleteModal] = useState(false);
  const [showErrorModal, setShowErrorModal] = useState(false);
  const [errorMessage, setErrorMessage] = useState('');
  const [fileToDelete, setFileToDelete] = useState(null);

  const handleFileChange = async (e) => {
      const selectedFiles = Array.from(e.target.files);
      
      for (const file of selectedFiles) {
          // Check Duplication (by name, since ID is filename)
          if (files.some(f => f.name === file.name)) {
               setErrorMessage(`File "${file.name}" already exists.`);
               setShowErrorModal(true);
               continue;
          }

          const formData = new FormData();
          formData.append('file', file);
          
          try {
              const res = await api.post(`/files/upload?user=${encodeURIComponent(currentUser?.name || 'Admin')}`, formData, {
                  headers: { 'Content-Type': 'multipart/form-data' }
              });
              
              // Backend returns the file object which matches our needed structure
              addFile(res.data);
              
          } catch (err) {
              console.error("Upload failed", err);
              setErrorMessage(`Failed to upload ${file.name}: ${err.response?.data?.detail || err.message}`);
              setShowErrorModal(true);
          }
      }
      
      // Reset Input
      e.target.value = '';
  };

  const confirmDelete = (file) => {
      setFileToDelete(file);
      setShowDeleteModal(true);
  };
  
  // Bulk Selection State
  const [selectedFiles, setSelectedFiles] = useState(new Set());
  const [isBulkDelete, setIsBulkDelete] = useState(false);

  const toggleSelect = (id) => {
      const newSet = new Set(selectedFiles);
      if (newSet.has(id)) newSet.delete(id);
      else newSet.add(id);
      setSelectedFiles(newSet);
  };

  const toggleSelectAll = () => {
      if (selectedFiles.size === files.length) {
          setSelectedFiles(new Set());
      } else {
          setSelectedFiles(new Set(files.map(f => f.id)));
      }
  };

  const confirmBulkDelete = () => {
      setIsBulkDelete(true);
      setShowDeleteModal(true);
  };

  const handleDelete = () => {
      if (isBulkDelete) {
          // Bulk Delete
          selectedFiles.forEach(id => {
              const file = files.find(f => f.id === id);
              if (file) deleteFile(file.name); 

              if (playingId === id) {
                  stopPlayback();
              }
          });
          setSelectedFiles(new Set());
          setIsBulkDelete(false);
          setShowDeleteModal(false);
      } else if (fileToDelete) {
          // Single Delete
          deleteFile(fileToDelete.name);

          // If playing deleted file, stop
          if (playingId === fileToDelete.id) {
              stopPlayback();
              if (currentLogId) {
                  const endTimeStr = new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
                  updateLog(currentLogId, {
                      action: 'Music Session',
                      details: `${fileToDelete.name} (Start: ${startTimeRef.current || ''} - End: ${endTimeStr})`
                  });
                  setCurrentLogId(null);
              }
          }
          setShowDeleteModal(false);
          setFileToDelete(null);
      }
  };

  return (
    <div className="space-y-6">
      <h2 className="text-2xl font-bold text-gray-800 flex items-center">
        <i className="material-icons mr-3 text-primary">upload</i> Upload Audio
      </h2>

      {emergencyActive && (
          <div className="bg-red-100 border-l-4 border-red-500 text-red-700 p-4 rounded shadow-sm flex items-center animate-pulse">
              <i className="material-icons text-2xl mr-3">warning</i>
              <div>
                   <p className="font-bold">Emergency Alert Active</p>
                   <p className="text-sm">Audio playback is temporarily disabled.</p>
               </div>
           </div>
       )}

       {systemState?.active_task && systemState.active_task.data?.user !== (currentUser?.name || 'Admin') && (
           <div className="bg-orange-100 border-l-4 border-orange-500 text-orange-800 p-4 rounded shadow-sm flex items-center animate-fade-in">
               <i className="material-icons text-2xl mr-3">lock</i>
               <div>
                   <p className="font-bold">System Busy</p>
                   <p className="text-sm">
                        <span className="font-semibold">{systemState.active_task.data?.user}</span> is using the system ({systemState.mode}).
                   </p>
               </div>
           </div>
       )}

      <div 
        onClick={() => fileInputRef.current.click()}
        className="bg-white rounded-xl shadow-sm border border-gray-100 p-8 text-center border-dashed border-2 border-gray-300 hover:border-primary transition-colors cursor-pointer group"
      >
         <div className="w-20 h-20 bg-primary/10 rounded-full flex items-center justify-center mx-auto mb-4 group-hover:scale-110 transition-transform">
           <i className="material-icons text-primary text-4xl">cloud_upload</i>
         </div>
         <h3 className="text-lg font-semibold text-gray-700">Upload Audio Files</h3>
         <p className="text-gray-500 mt-2">Drag & drop files here or click to browse</p>
         <input 
            type="file" 
            multiple 
            accept="audio/*"
            ref={fileInputRef} 
            className="hidden" 
            onChange={handleFileChange}
         />
         <button className="mt-4 px-6 py-2 bg-white border border-gray-300 rounded-md text-gray-700 font-medium hover:bg-gray-50 transition-colors">
           Browse Files
         </button>
      </div>

       <div className="bg-white rounded-xl shadow-sm border border-gray-100 p-6 overflow-hidden">
         
         {/* PLAYER BAR - Moved to Top */}
         {playingId && (
            <div className="mb-6 pb-6 border-b border-gray-100 animate-fade-in">
                <div className="flex flex-col space-y-3">
                    <div className="text-center mb-2">
                        <span className="inline-block px-3 py-1 bg-primary/10 text-primary text-xs font-bold rounded-full uppercase tracking-wider">
                            Now Playing
                        </span>
                        <p className="text-sm text-gray-700 font-medium mt-1 truncate">
                            {files.find(f => f.id === playingId)?.name || 'Unknown Track'}
                        </p>
                    </div>

                    {/* Time & Scrubber */}
                    <div className="flex items-center space-x-3 text-xs text-gray-500 font-mono">
                        <span className="w-10 text-right">{formatTime(currentTime)}</span>
                        <input 
                            type="range" 
                            min="0" 
                            max={duration || 0} 
                            value={currentTime} 
                            onChange={handleSeek}
                            className="flex-1 h-1 bg-gray-200 rounded-lg appearance-none cursor-pointer accent-primary"
                        />
                        <span className="w-10">{formatTime(duration)}</span>
                    </div>

                    {/* Controls */}
                    <div className="flex items-center justify-center space-x-6">
                        <button onClick={playPrev} className="text-gray-400 hover:text-primary transition-colors">
                            <i className="material-icons text-2xl">skip_previous</i>
                        </button>
                        
                        <button onClick={() => playSound(playingId)} className="w-12 h-12 bg-primary text-white rounded-full flex items-center justify-center shadow-md hover:bg-primary-dark transition-all transform hover:scale-105 active:scale-95">
                            <i className="material-icons text-2xl">{audioRef.current && !audioRef.current.paused ? 'pause' : 'play_arrow'}</i>
                        </button>
                        
                        <button onClick={stopPlayback} className="text-gray-400 hover:text-red-500 transition-colors" title="Stop">
                            <i className="material-icons text-2xl">stop</i>
                        </button>

                        <button onClick={playNext} className="text-gray-400 hover:text-primary transition-colors">
                            <i className="material-icons text-2xl">skip_next</i>
                        </button>
                    </div>
                </div>
            </div>
         )}

         {/* Header & Bulk Actions */}
         <div className="flex items-center justify-between mb-4">
             <div className="flex items-center space-x-3">
                <h3 className="text-lg font-semibold text-gray-700">Uploaded Files ({files.length})</h3>
                {files.length > 0 && (
                    <div className="flex items-center space-x-2 ml-4 px-3 py-1 bg-gray-50 rounded-lg">
                        <input 
                            type="checkbox"
                            checked={selectedFiles.size === files.length && files.length > 0}
                            onChange={toggleSelectAll}
                            className="w-4 h-4 text-primary rounded border-gray-300 focus:ring-primary cursor-pointer"
                        />
                        <span className="text-xs text-gray-500 font-medium">Select All</span>
                    </div>
                )}
             </div>

             {selectedFiles.size > 0 && (
                 <button 
                    onClick={confirmBulkDelete}
                    className="flex items-center px-3 py-1.5 bg-red-50 text-red-600 rounded-lg text-sm font-medium hover:bg-red-100 transition-colors"
                 >
                     <i className="material-icons text-base mr-1">delete</i>
                     Delete ({selectedFiles.size})
                 </button>
             )}
         </div>
         
         {files.length > 0 ? (
             <div className="space-y-2 max-h-[400px] overflow-y-auto mb-4 pr-1">
                  {files.map((file) => {
                      const isLocked = systemState?.active_task && systemState.active_task.data?.user !== (currentUser?.name || 'Admin');
                      return (
                      <div 
                         key={file.id} 
                         className={`flex items-center justify-between p-3 rounded-lg transition-colors group ${isLocked ? 'cursor-not-allowed opacity-60 bg-gray-50' : 'cursor-pointer'} ${playingId === file.id ? 'bg-primary/5 border border-primary/20' : (!isLocked && 'hover:bg-gray-50 border border-transparent')} ${selectedFiles.has(file.id) ? 'bg-blue-50/50' : ''}`}
                         onClick={() => !isLocked && playSound(file.id)}
                      >
                          <div className="flex items-center overflow-hidden flex-1">
                              {/* Checkbox */}
                              <div 
                                 onClick={(e) => { e.stopPropagation(); !isLocked && toggleSelect(file.id); }}
                                 className={`mr-3 flex items-center justify-center p-1 rounded-full ${!isLocked && 'hover:bg-black/5 cursor-pointer'} z-10`}
                              >
                                 <input 
                                    type="checkbox"
                                    checked={selectedFiles.has(file.id)}
                                    onChange={() => {}} // Handled by div click
                                    className="w-4 h-4 text-primary rounded border-gray-300 focus:ring-primary pointer-events-none"
                                 />
                             </div>

                             <div className={`w-10 h-10 rounded flex items-center justify-center mr-3 flex-shrink-0 ${playingId === file.id ? 'bg-primary text-white' : 'bg-gray-100 text-gray-500'}`}>
                                 <i className="material-icons">{playingId === file.id ? 'equalizer' : 'audiotrack'}</i>
                             </div>
                             <div className="min-w-0">
                                 <h4 className={`font-medium truncate text-sm ${playingId === file.id ? 'text-primary' : 'text-gray-800'}`}>{file.name}</h4>
                                 <p className="text-xs text-gray-500">{file.size} • {file.date}</p>
                             </div>
                         </div>
                         <div className="flex items-center space-x-2 pl-2">
                             <button 
                                onClick={(e) => { e.stopPropagation(); confirmDelete(file); }}
                                className="p-1.5 text-gray-400 hover:text-red-600 hover:bg-red-50 rounded-full transition-colors z-10"
                             >
                                 <i className="material-icons text-lg">delete</i>
                             </button>
                     </div>
                     </div>
                 );
                 })}
             </div>
         ) : (
             <div className="text-center text-gray-500 py-8 border-2 border-dashed border-gray-100 rounded-lg">
               No files uploaded yet.
             </div>
         )}
       </div>

       {/* Delete Modal */}
       <Modal
          isOpen={showDeleteModal}
          onClose={() => setShowDeleteModal(false)}
          title={isBulkDelete ? "Delete Multiple Files" : "Delete File"}
          type="danger"
          footer={
             <>
                 <button 
                    onClick={() => setShowDeleteModal(false)}
                    className="px-4 py-2 border border-gray-300 rounded-lg text-gray-700 font-medium hover:bg-gray-50"
                 >
                     Cancel
                 </button>
                 <button 
                    onClick={handleDelete}
                    className="px-4 py-2 bg-red-600 text-white rounded-lg font-medium hover:bg-red-700 shadow-md"
                 >
                     Delete {isBulkDelete && `(${selectedFiles.size})`}
                 </button>
             </>
          }
       >
           <p className="text-gray-600">
               {isBulkDelete 
                   ? `Are you sure you want to delete ${selectedFiles.size} selected files? This action cannot be undone.`
                   : <>Are you sure you want to delete <span className="font-bold">{fileToDelete?.name}</span>? This action cannot be undone.</>
               }
           </p>
       </Modal>

       {/* Error Modal */}
       <Modal
            isOpen={showErrorModal}
            onClose={() => setShowErrorModal(false)}
            title="Upload Error"
            type="info"
            footer={<button onClick={() => setShowErrorModal(false)} className="px-6 py-2 bg-primary text-white rounded-lg">OK</button>}
       >
           <p className="text-gray-600">{errorMessage}</p>
       </Modal>
    </div>
  );
};

export default Upload;
