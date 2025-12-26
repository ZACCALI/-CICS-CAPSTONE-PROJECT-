import React, { useState } from 'react';
import { useApp } from '../../context/AppContext';
import { useAuth } from '../../context/AuthContext';
import Modal from '../common/Modal';

const HistoryLogs = () => {
  const { activityLogs, deleteLog, deleteLogs } = useApp();
  const { currentUser } = useAuth();
  const [filter, setFilter] = useState('All'); // Type Filter
  const [timeFilter, setTimeFilter] = useState('All'); // Time Filter (default All)
  const [showTimelineDropdown, setShowTimelineDropdown] = useState(false);
  const [showTypeDropdown, setShowTypeDropdown] = useState(false);
  const [selectedLogs, setSelectedLogs] = useState([]); // Array of IDs
  const [selectedLog, setSelectedLog] = useState(null);
  const [showModal, setShowModal] = useState(false);

  const isWithinTimeRange = (logDateStr, range) => {
      if (range === 'All') return true;
      
      const now = new Date();
      const logDate = new Date(logDateStr);
      
      if (isNaN(logDate.getTime())) return true; // Fail safe

      // Reset times for comparison
      now.setHours(0,0,0,0);
      const logDay = new Date(logDate);
      logDay.setHours(0,0,0,0);

      if (range === 'Today') {
          return logDay.getTime() === now.getTime();
      }
      
      if (range === 'Week') {
          const startOfWeek = new Date(now);
          startOfWeek.setDate(now.getDate() - now.getDay()); // Sunday
          return logDate >= startOfWeek;
      }
      
      if (range === 'Month') {
          return logDate.getMonth() === now.getMonth() && logDate.getFullYear() === now.getFullYear();
      }

      if (range === 'Year') {
          return logDate.getFullYear() === now.getFullYear();
      }
      
      return true;
  };

  const filteredLogs = activityLogs.filter(log => {
      let typeMatch = false;
      if (filter === 'All') {
          typeMatch = true;
      } else if (filter === 'System') {
          // Group 'Account' and 'Session' under 'System' filter
          typeMatch = ['System', 'Account', 'Session'].includes(log.type);
      } else {
          typeMatch = log.type === filter;
      }

      // Prefer raw timestamp for accuracy, fallback to formatted time string if needed
      const dateToUse = log.timestamp || log.time; 
      const timeMatch = isWithinTimeRange(dateToUse, timeFilter);
      return typeMatch && timeMatch;
  });

  const [showBulkDeleteModal, setShowBulkDeleteModal] = useState(false);

  const handleRowClick = (log) => {
      setSelectedLog(log);
      setShowModal(true);
  };

  const handleSelectAll = (e) => {
      if (e.target.checked) {
          setSelectedLogs(filteredLogs.map(l => l.id));
      } else {
          setSelectedLogs([]);
      }
  };

  const handleSelectOne = (e, id) => {
      e.stopPropagation();
      if (e.target.checked) {
          setSelectedLogs(prev => [...prev, id]);
      } else {
          setSelectedLogs(prev => prev.filter(lid => lid !== id));
      }
  };

  const handleBulkDeleteClick = () => {
      setShowBulkDeleteModal(true);
  };

  const confirmBulkDelete = () => {
      deleteLogs(selectedLogs, currentUser?.name);
      setSelectedLogs([]);
      setShowBulkDeleteModal(false);
  };

  return (
    <div className="space-y-6">
      <div className="flex justify-between items-center">
        <h2 className="text-2xl font-bold text-gray-800 flex items-center">
            <i className="material-icons mr-3 text-primary">history_edu</i> Activity Logs
        </h2>
        {selectedLogs.length > 0 && (
            <button 
                onClick={handleBulkDeleteClick}
                className="px-4 py-2 bg-red-600 text-white text-sm font-bold rounded-lg shadow-sm hover:bg-red-700 transition-colors flex items-center"
            >
                <i className="material-icons text-base mr-2">delete</i>
                Delete Selected ({selectedLogs.length})
            </button>
        )}
      </div>

      <div className="bg-white rounded-xl shadow-sm border border-gray-100">
        <div className="p-4 border-b border-gray-100 bg-gray-50 rounded-t-xl">
             <div className="flex flex-col md:flex-row justify-between items-center gap-4">
                {/* Mobile: Grid row for side-by-side. Desktop: Flex row */}
                <div className="grid grid-cols-2 gap-3 w-full md:w-auto md:flex md:items-center">
                    
                    {/* TIMELINE FILTER */}
                    <div className="relative">
                        <button 
                            onClick={() => setShowTimelineDropdown(!showTimelineDropdown)}
                            className="w-full px-4 py-2 bg-white border border-gray-200 rounded-lg text-sm font-medium text-gray-700 hover:bg-gray-50 flex items-center justify-between md:justify-start shadow-sm whitespace-nowrap"
                        >
                            <div className="flex items-center">
                                <i className="material-icons text-base mr-2 text-primary">calendar_today</i>
                                <span>{timeFilter === 'All' ? 'All Time' : timeFilter}</span>
                            </div>
                            <i className="material-icons text-base ml-2 text-gray-400">arrow_drop_down</i>
                        </button>
                        
                        {showTimelineDropdown && (
                            <div className="absolute top-full left-0 mt-2 w-40 bg-white border border-gray-100 rounded-lg shadow-lg z-20 max-h-60 overflow-y-auto py-1">
                                {['All', 'Today', 'Week', 'Month', 'Year'].map(t => (
                                    <button
                                        key={t}
                                        onClick={() => {
                                            setTimeFilter(t);
                                            setShowTimelineDropdown(false);
                                        }}
                                        className={`w-full text-left px-4 py-2 text-sm hover:bg-gray-50 transition-colors ${timeFilter === t ? 'text-primary font-bold bg-primary/5' : 'text-gray-600'}`}
                                    >
                                        {t === 'All' ? 'All Time' : t}
                                    </button>
                                ))}
                            </div>
                        )}
                        {/* Backdrop */}
                        {showTimelineDropdown && (
                            <div className="fixed inset-0 z-10" onClick={() => setShowTimelineDropdown(false)}></div>
                        )}
                    </div>

                    {/* TYPE FILTER (DROPDOWN - ALL SCREENS) */}
                    <div className="relative">
                        <button 
                            onClick={() => setShowTypeDropdown(!showTypeDropdown)}
                            className="w-full px-4 py-2 bg-white border border-gray-200 rounded-lg text-sm font-medium text-gray-700 hover:bg-gray-50 flex items-center justify-between shadow-sm whitespace-nowrap min-w-[140px]"
                        >   
                            <div className="flex items-center">
                                <i className="material-icons text-base mr-2 text-primary">filter_list</i>
                                <span>{filter === 'All' ? 'All Types' : filter}</span>
                            </div>
                            <i className="material-icons text-base ml-2 text-gray-400">arrow_drop_down</i>
                        </button>

                        {showTypeDropdown && (
                            <div className="absolute top-full left-0 mt-2 w-40 bg-white border border-gray-100 rounded-lg shadow-lg z-20 max-h-60 overflow-y-auto py-1">
                                {['All', 'Text', 'Voice', 'Music', 'Schedule', 'System', 'Emergency'].map(f => (
                                    <button
                                        key={f}
                                        onClick={() => {
                                            setFilter(f);
                                            setShowTypeDropdown(false);
                                        }}
                                        className={`w-full text-left px-4 py-2 text-sm hover:bg-gray-50 transition-colors ${filter === f ? 'text-primary font-bold bg-primary/5' : 'text-gray-600'}`}
                                    >
                                        {f === 'All' ? 'All Types' : f}
                                    </button>
                                ))}
                            </div>
                        )}
                         {/* Backdrop */}
                         {showTypeDropdown && (
                            <div className="fixed inset-0 z-10" onClick={() => setShowTypeDropdown(false)}></div>
                        )}
                    </div>
                </div>
             </div>
        </div>
        
        {/* Desktop Table View */}
        <div className="hidden md:block overflow-x-auto">
            <table className="w-full text-left text-sm text-gray-600">
                <thead className="bg-gray-50 border-b border-gray-200 font-semibold text-gray-700 uppercase tracking-wider text-xs">
                    <tr><th className="px-6 py-4 w-4">
                            <input 
                                type="checkbox" 
                                onChange={handleSelectAll}
                                checked={filteredLogs.length > 0 && selectedLogs.length === filteredLogs.length}
                                className="rounded border-gray-300 text-primary focus:ring-primary h-4 w-4"
                            />
                        </th><th className="px-6 py-4">Time</th><th className="px-6 py-4">User</th><th className="px-6 py-4">Action</th><th className="px-6 py-4">Type</th><th className="px-6 py-4">Details</th><th className="px-6 py-4"></th></tr>
                </thead>
                <tbody className="divide-y divide-gray-100">
                    {filteredLogs.length > 0 ? (
                        filteredLogs.map(log => (
                            <tr key={log.id} onClick={() => handleRowClick(log)} className={`hover:bg-gray-50 transition-colors cursor-pointer ${selectedLogs.includes(log.id) ? 'bg-blue-50' : ''}`}><td className="px-6 py-4" onClick={(e) => e.stopPropagation()}>
                                    <input 
                                        type="checkbox" 
                                        checked={selectedLogs.includes(log.id)}
                                        onChange={(e) => handleSelectOne(e, log.id)}
                                        className="rounded border-gray-300 text-primary focus:ring-primary h-4 w-4"
                                    />
                                </td><td className="px-6 py-4 whitespace-nowrap text-xs text-gray-500 font-mono">{log.time}</td><td className="px-6 py-4 font-medium text-gray-900">{log.user}</td><td className="px-6 py-4">{log.action}</td><td className="px-6 py-4"><span className={`px-2 py-1 text-xs rounded-full font-bold uppercase ${log.type === 'Emergency' ? 'bg-red-100 text-red-700' : log.type === 'Voice' ? 'bg-green-100 text-green-700' : log.type === 'Music' ? 'bg-yellow-100 text-yellow-700' : log.type === 'Text' ? 'bg-blue-100 text-blue-700' : log.type === 'Schedule' ? 'bg-purple-100 text-purple-700' : 'bg-gray-100 text-gray-700'}`}>{log.type}</span></td><td className="px-6 py-4 text-gray-500"><div className="max-w-xs truncate cursor-help" title={log.details}>{log.details.length > 50 ? log.details.substring(0, 50) + '...' : log.details}</div></td><td className="px-6 py-4 text-right"><i className="material-icons text-gray-300">chevron_right</i></td></tr>
                        ))
                    ) : (
                        <tr><td colSpan="7" className="px-6 py-8 text-center text-gray-400">No activity logs found.</td></tr>
                    )}
                </tbody>
            </table>
        </div>

        {/* Mobile Card View */}
        <div className="md:hidden">
            {filteredLogs.length > 0 ? (
                <div className="divide-y divide-gray-100">
                    {filteredLogs.map(log => (
                        <div 
                            key={log.id} 
                            onClick={() => handleRowClick(log)}
                            className={`p-4 hover:bg-gray-50 transition-colors cursor-pointer ${selectedLogs.includes(log.id) ? 'bg-blue-50' : ''}`}
                        >
                            <div className="flex justify-between items-start mb-2">
                                <div className="flex items-center space-x-3">
                                    <input 
                                        type="checkbox" 
                                        checked={selectedLogs.includes(log.id)}
                                        onChange={(e) => handleSelectOne(e, log.id)}
                                        onClick={(e) => e.stopPropagation()}
                                        className="rounded border-gray-300 text-primary focus:ring-primary h-5 w-5"
                                    />
                                    <div>
                                        <div className="font-bold text-gray-800 text-sm">{log.action}</div>
                                        <div className="text-xs text-gray-500">{log.user} • {log.time}</div>
                                    </div>
                                </div>
                                <span className={`px-2 py-0.5 text-[10px] rounded-full font-bold uppercase 
                                    ${log.type === 'Emergency' ? 'bg-red-100 text-red-700' : 
                                      log.type === 'Voice' ? 'bg-green-100 text-green-700' :
                                      log.type === 'Music' ? 'bg-yellow-100 text-yellow-700' :
                                      log.type === 'Text' ? 'bg-blue-100 text-blue-700' : 
                                      log.type === 'Schedule' ? 'bg-purple-100 text-purple-700' : 
                                      'bg-gray-100 text-gray-700'}`}>
                                    {log.type}
                                </span>
                            </div>
                            <p className="text-sm text-gray-600 line-clamp-2 pl-8">{log.details}</p>
                        </div>
                    ))}
                </div>
            ) : (
                <div className="p-8 text-center text-gray-400 bg-gray-50">
                    No activity logs found.
                </div>
            )}
        </div>
      </div>

      {selectedLog && (
        <Modal
            isOpen={showModal}
            onClose={() => setShowModal(false)}
            title="Log Details"
            footer={
                <div className="flex justify-end space-x-3">
                    <button 
                        onClick={() => {
                            // Trigger deletion flow using the Bulk Delete Modal logic (reused for single)
                            setSelectedLogs([selectedLog.id]);
                            setShowBulkDeleteModal(true);
                            setShowModal(false); // Close details modal
                        }} 
                        className="px-4 py-2 bg-white text-red-600 border border-red-200 rounded-lg hover:bg-red-50 transition-colors flex items-center font-medium"
                    >
                         <i className="material-icons mr-2 text-sm">delete</i> Delete
                    </button>
                    <button onClick={() => setShowModal(false)} className="px-6 py-2 bg-primary text-white rounded-lg hover:bg-primary-dark transition-colors font-medium">Close</button>
                </div>
            }
        >
            <div className="space-y-4">
                <div>
                    <label className="block text-xs font-bold text-gray-500 uppercase">Action</label>
                    <p className="text-lg font-bold text-gray-800">{selectedLog.action}</p>
                </div>
                <div className="grid grid-cols-2 gap-4">
                    <div>
                        <label className="block text-xs font-bold text-gray-500 uppercase">User</label>
                        <p className="text-gray-700">{selectedLog.user}</p>
                    </div>
                    <div>
                        <label className="block text-xs font-bold text-gray-500 uppercase">Time</label>
                        <p className="text-gray-700 font-mono text-sm">{selectedLog.time}</p>
                    </div>
                </div>
                <div>
                    <label className="block text-xs font-bold text-gray-500 uppercase">Details</label>
                    <div className="mt-1 p-3 bg-gray-50 rounded-lg border border-gray-100 text-gray-700 whitespace-pre-wrap font-mono text-sm">
                        {selectedLog.details}
                    </div>
                </div>
                <div>
                    <span className={`px-2 py-1 text-xs rounded-full font-bold uppercase 
                        ${selectedLog.type === 'Emergency' ? 'bg-red-100 text-red-700' : 
                          selectedLog.type === 'Voice' ? 'bg-green-100 text-green-700' :
                          selectedLog.type === 'Music' ? 'bg-yellow-100 text-yellow-700' :
                          selectedLog.type === 'Text' ? 'bg-blue-100 text-blue-700' : 'bg-gray-100 text-gray-600'}`}>
                        {selectedLog.type}
                    </span>
                </div>
            </div>
        </Modal>
      )}

      {/* Bulk Delete Confirmation Modal */}
      <Modal
          isOpen={showBulkDeleteModal}
          onClose={() => setShowBulkDeleteModal(false)}
          title={`Delete ${selectedLogs.length} Log${selectedLogs.length !== 1 ? 's' : ''}?`}
          type="danger"
          footer={
             <>
                <button 
                   onClick={() => setShowBulkDeleteModal(false)}
                   className="px-4 py-2 border rounded-lg text-gray-700 hover:bg-gray-50 bg-white"
                >
                    Cancel
                </button>
                <button 
                   onClick={confirmBulkDelete}
                   className="px-4 py-2 bg-red-600 text-white rounded-lg font-medium hover:bg-red-700 shadow-lg"
                >
                    Delete {selectedLogs.length} Item{selectedLogs.length !== 1 ? 's' : ''}
                </button>
             </>
          }
      >
          <div className="text-gray-600">
             <p className="mb-2">Are you sure you want to delete the selected logs?</p>
             <p className="text-sm text-gray-500">This action cannot be undone.</p>
          </div>
      </Modal>
    </div>
  );
};

export default HistoryLogs;
