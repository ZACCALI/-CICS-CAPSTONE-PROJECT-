import React, { useState, useEffect } from 'react';
import { useAuth } from '../../context/AuthContext';
import { collection, query, orderBy, onSnapshot } from "firebase/firestore";
import { db } from '../../firebase';
import Modal from '../common/Modal';

const ManageAccount = () => {
  const { currentUser, updateUser, adminUpdateProfile, getAllUsers, adminDeleteUser, adminResetPassword, adminApproveUser, adminCreateUser, getSystemLogs } = useAuth();
  
  const [formData, setFormData] = useState({
      name: '',
      email: '',
      avatar: ''
  });

  // Sync formData with currentUser when it changes
  useEffect(() => {
    if (currentUser) {
        setFormData({
            name: currentUser.name || '',
            email: currentUser.email || '',
            avatar: currentUser.avatar || ''
        });
    }
  }, [currentUser]);

  const [message, setMessage] = useState('');
  const [errorMessage, setErrorMessage] = useState('');
  const [showErrorModal, setShowErrorModal] = useState(false);
  const [successMessage, setSuccessMessage] = useState('');
  const [showSuccessModal, setShowSuccessModal] = useState(false);

  const [showPasswordModal, setShowPasswordModal] = useState(false);
  const [passForm, setPassForm] = useState({ current: '', new: '', confirm: '' });
  
  // Toggles
  const [showPass, setShowPass] = useState({ current: false, new: false, confirm: false });

  // ADMIN STATE
  const [userList, setUserList] = useState([]);
  const [logs, setLogs] = useState([]);

  // Firestore Listener for Admin
  useEffect(() => {
      let unsubscribe = () => {};

      if (currentUser?.role === 'admin') {
           try {
               const q = query(collection(db, "users"), orderBy("createdAt", "desc"));
               unsubscribe = onSnapshot(q, (snapshot) => {
                   const users = snapshot.docs.map(doc => ({
                       uid: doc.id,
                       ...doc.data()
                   }));
                   setUserList(users);
               }, (error) => {
                   console.error("Firestore sync error:", error);
                   loadUsersFallback();
               });
           } catch (e) {
               console.error("Failed to setup listener", e);
               loadUsersFallback();
           }
      }

      return () => unsubscribe();
  }, [currentUser]);

  const loadUsersFallback = async () => {
      try {
          const users = await getAllUsers();
          // Safety check: only update if component is still effectively "active" in context of this async flow
          // (React 18 handles unmount updates better, but good practice)
          setUserList(users);
      } catch (error) {
          console.error("Failed to load users (fallback)", error);
      }
  };
  
  const [selectedUser, setSelectedUser] = useState(null);
  const [showDeleteModal, setShowDeleteModal] = useState(false);
  const [showLogModal, setShowLogModal] = useState(false);
  const [showResetModal, setShowResetModal] = useState(false);
  const [showApproveModal, setShowApproveModal] = useState(false); // New state for approve confirmation
  const [resetResult, setResetResult] = useState('');
  
  // Add User State
  const [showAddUserModal, setShowAddUserModal] = useState(false);
  const [addUserForm, setAddUserForm] = useState({ name: '', email: '', password: '', role: 'user' });

  // Derived Lists
  const pendingUsers = userList.filter(u => u.status === 'pending');
  // Show all users in the main table EXCEPT pending ones (or show all if you prefer, but separation is cleaner)
  const activeUsers = userList.filter(u => u.status !== 'pending'); 

  // Actions
  const handleFileChange = (e) => {
      const file = e.target.files[0];
      if (file) {
          const reader = new FileReader();
          reader.onloadend = () => {
              setFormData(prev => ({ ...prev, avatar: reader.result }));
          };
          reader.readAsDataURL(file);
      }
  };

  const handleUpdate = async (e) => {
      e.preventDefault();
      try {
        const payload = { 
            name: formData.name, 
            email: formData.email,
            avatar: formData.avatar 
        };

        if (currentUser.role === 'admin') {
            await adminUpdateProfile(payload);
        } else {
            await updateUser(payload);
        }

        setSuccessMessage('Profile updated successfully!');
        setShowSuccessModal(true);
        window.scrollTo({ top: 0, behavior: 'smooth' });
      } catch (error) {
        setErrorMessage("Failed to update profile: " + error.message);
        setShowErrorModal(true);
      }
  };

  const handlePasswordChange = async (e) => {
      e.preventDefault();
      if (passForm.new !== passForm.confirm) {
          setErrorMessage("New passwords do not match");
          setShowErrorModal(true);
          return;
      }
      try {
        if (currentUser.role === 'admin') {
            await adminUpdateProfile({ password: passForm.new });
        } else {
            await updateUser({ password: passForm.new }); 
        } 
        setShowPasswordModal(false);
        setSuccessMessage('Password changed successfully!');
        setShowSuccessModal(true);
        setPassForm({ current: '', new: '', confirm: '' });
      } catch (error) {
        setErrorMessage("Failed to change password: " + error.message);
        setShowErrorModal(true);
      }
  };

  // ADMIN ACTIONS
  const handleAdminReset = async (uid) => {
      // Use UID now
      const newPass = await adminResetPassword(uid);
      if (newPass) {
          setResetResult(`Password reset successfully. New Password: ${newPass}`);
          setShowResetModal(true);
      }
  };

  const handleApproveClick = (user) => {
      setSelectedUser(user);
      setShowApproveModal(true);
  };

  const confirmApprove = async () => {
      if (!selectedUser) return;
      try {
          await adminApproveUser(selectedUser.uid);
          setSuccessMessage(`User ${selectedUser.name} approved successfully.`);
          setShowSuccessModal(true);
          setShowApproveModal(false);
          setSelectedUser(null);
      } catch (error) {
          setErrorMessage("Failed to approve user: " + error.message);
          setShowErrorModal(true);
      }
  };

  const handleDeleteClick = (user) => {
      setSelectedUser(user);
      setShowDeleteModal(true);
  };

  const confirmDelete = async () => {
      if (!selectedUser) return;
      
      if (selectedUser.email === currentUser.email) {
          setErrorMessage("You cannot delete your own admin account.");
          setShowErrorModal(true);
          setShowDeleteModal(false);
          return;
      }

      try {
          const success = await adminDeleteUser(selectedUser.uid); 
          if (success) {
              setSuccessMessage('User rejected/deleted successfully.');
              setShowSuccessModal(true);
              setShowDeleteModal(false);
              setSelectedUser(null);
          }
      } catch (error) {
           setErrorMessage("Failed to delete user: " + error.message);
           setShowErrorModal(true);
      }
  };

  const viewLogs = async (user) => {
      setSelectedUser(user);
      setShowLogModal(true);
      setLogs([]); // Clear previous
      try {
          const allLogs = await getSystemLogs();
          // Filter by username
          const userLogs = allLogs.filter(l => l.user === user.name || l.user === user.email);
          setLogs(userLogs);
      } catch (e) {
          console.error(e);
      }
  };

  const toggleVisibility = (field) => {
      setShowPass(prev => ({ ...prev, [field]: !prev[field] }));
  };

  const handleCreateUser = async (e) => {
      e.preventDefault();
      try {
          await adminCreateUser(addUserForm);
          setSuccessMessage(`User ${addUserForm.email} created successfully.`);
          setShowSuccessModal(true);
          setShowAddUserModal(false);
          setAddUserForm({ name: '', email: '', password: '', role: 'user' });
      } catch (error) {
          setErrorMessage("Failed to create user: " + (error.response?.data?.detail || error.message));
          setShowErrorModal(true);
      }
  };

  return (
    <div className="space-y-8 animate-fade-in">
      
      {/* 1. Profile Section */}
      <div>
        <h2 className="text-2xl font-bold text-gray-800 flex items-center mb-6">
            <i className="material-icons mr-3 text-primary">admin_panel_settings</i> Manage Account
        </h2>

        <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
            {/* Profile Card */}
            <div className="bg-white rounded-xl shadow-sm border border-gray-100 p-6 lg:col-span-1 h-fit">
                <div className="flex flex-col items-center">
                <div className="relative group w-24 h-24 mb-4">
                    <div className="w-24 h-24 rounded-full overflow-hidden border-4 border-gray-100 relative shadow-md">
                            <img 
                                src={formData.avatar || `https://ui-avatars.com/api/?name=${currentUser?.name}&background=2563EB&color=fff`} 
                                alt="Profile" 
                                className="w-full h-full object-cover rounded-full" 
                            />
                    </div>
                    <label className="absolute inset-0 bg-black/50 rounded-full flex items-center justify-center opacity-0 group-hover:opacity-100 transition-opacity cursor-pointer">
                        <i className="material-icons text-white">camera_alt</i>
                        <input type="file" accept="image/*" onChange={handleFileChange} className="hidden" />
                    </label>
                </div>
                <h3 className="text-xl font-bold text-gray-800 text-center">{currentUser?.name}</h3>
                <p className="text-gray-500 text-center break-all">{currentUser?.email}</p>
                <div className="mt-4 px-3 py-1 bg-green-100 text-green-700 rounded-full text-xs font-semibold uppercase">
                    {currentUser?.role} Active
                </div>
                </div>
            </div>

            {/* Account Form */}
            <div className="bg-white rounded-xl shadow-sm border border-gray-100 p-6 lg:col-span-2">
                <h3 className="text-lg font-semibold text-gray-700 mb-6 pb-2 border-b border-gray-100">Account Details</h3>
                {message && (
                    <div className="mb-4 p-3 bg-green-50 text-green-700 rounded-lg text-sm font-medium flex items-center animate-pulse">
                        <i className="material-icons mr-2 text-sm">check_circle</i> {message}
                    </div>
                )}
                <form onSubmit={handleUpdate} className="space-y-4">
                <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                    <div>
                    <label className="block text-sm font-medium text-gray-700 mb-1">Full Name</label>
                    <input type="text" value={formData.name} onChange={e => setFormData({...formData, name: e.target.value})} className="w-full p-2 border border-gray-200 rounded-lg focus:ring-2 focus:ring-primary/20 focus:border-primary outline-none transition-all" />
                    </div>
                    <div>
                    <label className="block text-sm font-medium text-gray-700 mb-1">Email</label>
                    <input 
                        type="email" 
                        value={formData.email} 
                        onChange={e => setFormData({...formData, email: e.target.value})} 
                        className="w-full p-2 border border-gray-200 rounded-lg focus:ring-2 focus:ring-primary/20 focus:border-primary outline-none transition-all" 
                    />
                    </div>
                </div>
                <div className="pt-6 flex flex-col sm:flex-row justify-end space-y-3 sm:space-y-0 sm:space-x-3 border-t border-gray-100 mt-6">
                    <button type="button" onClick={() => setShowPasswordModal(true)} className="px-4 py-2 text-gray-700 bg-white border border-gray-300 hover:bg-gray-50 rounded-lg font-medium transition-colors">Change Password</button>
                    <button type="submit" className="px-6 py-2 bg-primary hover:bg-primary-dark text-white rounded-lg shadow-md hover:shadow-lg transition-all font-medium">Save Changes</button>
                </div>
                </form>
            </div>
        </div>
      </div>

      {/* 2. ADMIN SECTION: User Management */}
      {currentUser?.role === 'admin' && (
        <div className="space-y-6">
            
            {/* PENDING APPROVALS SECTION */}
            {pendingUsers.length > 0 && (
                <div>
                     <h3 className="text-lg font-bold text-gray-800 flex items-center mb-4">
                        <i className="material-icons mr-2 text-orange-500">pending_actions</i> Pending Approvals
                        <span className="ml-2 bg-orange-100 text-orange-600 text-xs px-2 py-1 rounded-full">{pendingUsers.length}</span>
                    </h3>
                    
                    <div className="grid grid-cols-1 md:grid-cols-2 xl:grid-cols-3 gap-4">
                        {pendingUsers.map(u => (
                            <div key={u.uid} className="bg-white p-5 rounded-xl border border-orange-200 shadow-[0_4px_20px_rgba(251,146,60,0.1)] flex flex-col relative overflow-hidden">
                                <div className="absolute top-0 right-0 w-16 h-16 bg-orange-50 rounded-bl-full -mr-8 -mt-8 z-0"></div>
                                
                                <div className="flex items-center mb-4 z-10">
                                    <div className="w-12 h-12 rounded-full bg-orange-100 text-orange-600 flex items-center justify-center font-bold text-xl mr-4 shadow-sm">
                                        {u.name.charAt(0)}
                                    </div>
                                    <div className="overflow-hidden">
                                        <h4 className="font-bold text-gray-800 truncate">{u.name}</h4>
                                        <p className="text-xs text-gray-500 truncate">{u.email}</p>
                                    </div>
                                </div>
                                
                                <div className="flex items-center text-xs text-gray-500 mb-4 bg-gray-50 p-2 rounded-lg z-10">
                                    <i className="material-icons text-sm mr-1">access_time</i>
                                    Registered: {u.createdAt ? new Date(u.createdAt).toLocaleDateString() : 'Unknown'}
                                </div>
                                
                                <div className="flex gap-2 mt-auto z-10">
                                    <button 
                                        onClick={() => handleDeleteClick(u)}
                                        className="flex-1 py-2 text-red-600 bg-red-50 hover:bg-red-100 rounded-lg text-sm font-semibold transition-colors"
                                    >
                                        Reject
                                    </button>
                                    <button 
                                        onClick={() => handleApproveClick(u)}
                                        className="flex-[2] py-2 text-white bg-green-500 hover:bg-green-600 rounded-lg text-sm font-semibold shadow-md shadow-green-200 transition-all transform hover:-translate-y-0.5"
                                    >
                                        Approve Access
                                    </button>
                                </div>
                            </div>
                        ))}
                    </div>
                </div>
            )}

            {/* ALL USERS TABLE */}
            <div className="bg-white rounded-xl shadow-sm border border-gray-100 overflow-hidden">
                <div className="p-6 border-b border-gray-100 bg-gray-50 flex justify-between items-center">
                    <h3 className="text-lg font-bold text-gray-800 flex items-center">
                        <i className="material-icons mr-2 text-primary">people</i> Account Management
                    </h3>
                    <div className="flex items-center space-x-3">
                        <button 
                            onClick={() => setShowAddUserModal(true)}
                            className="bg-primary hover:bg-primary-dark text-white text-xs font-bold py-2 px-4 rounded-lg flex items-center shadow-md transition-all"
                        >
                            <i className="material-icons text-sm mr-1">add</i> Add User
                        </button>
                        <span className="text-xs font-semibold bg-gray-200 text-gray-600 px-2 py-1 rounded-full">{activeUsers.length} Users</span>
                    </div>
                </div>
                
                <div className="hidden md:block overflow-x-auto">
                    <table className="w-full text-left text-sm text-gray-600">
                        <thead className="bg-gray-50 border-b border-gray-200 font-semibold text-gray-700 uppercase tracking-wider text-xs">
                            <tr>
                                <th className="px-6 py-4">User</th>
                                <th className="px-6 py-4">Role</th>
                                <th className="px-6 py-4">Status</th>
                                <th className="px-6 py-4">Last Login</th>
                                <th className="px-6 py-4 text-center">Actions</th>
                            </tr>
                        </thead>
                        <tbody className="divide-y divide-gray-100">
                            {activeUsers.length === 0 ? (
                                <tr>
                                    <td colSpan="5" className="text-center py-8 text-gray-400">No active users found.</td>
                                </tr>
                            ) : (
                                activeUsers.map(u => (
                                <tr key={u.uid} className="hover:bg-gray-50 transition-colors">
                                    <td className="px-6 py-4">
                                        <div className="flex items-center">
                                            <div className="flex-shrink-0 h-8 w-8 rounded-full bg-primary/10 flex items-center justify-center text-primary font-bold mr-3">
                                                {u.name.charAt(0)}
                                            </div>
                                            <div>
                                                <div className="font-medium text-gray-900">{u.name}</div>
                                                <div className="text-xs text-gray-500">{u.email}</div>
                                            </div>
                                        </div>
                                    </td>
                                    <td className="px-6 py-4">
                                        <span className={`px-2 py-1 text-xs rounded-full font-bold uppercase ${u.role === 'admin' ? 'bg-purple-100 text-purple-700' : 'bg-blue-50 text-blue-600'}`}>
                                            {u.role}
                                        </span>
                                    </td>
                                    <td className="px-6 py-4">
                                        {u.status === 'pending' ? (
                                             <span className="px-2 py-1 text-xs rounded-full font-bold uppercase bg-yellow-100 text-yellow-700">
                                                Pending
                                             </span>
                                        ) : (
                                            <span className={`px-2 py-1 text-xs rounded-full font-bold uppercase ${u.isOnline ? 'bg-green-100 text-green-700' : 'bg-gray-100 text-gray-500'}`}>
                                                {u.isOnline ? 'Online' : 'Offline'}
                                            </span>
                                        )}
                                    </td>
                                    <td className="px-6 py-4 text-xs font-mono">
                                        {u.lastLogin ? new Date(u.lastLogin).toLocaleString() : 'Never'}
                                    </td>
                                    <td className="px-6 py-4 text-center space-x-2 flex justify-center">
                                        <button 
                                            onClick={() => viewLogs(u)}
                                            className="text-primary hover:text-primary-dark text-xs font-medium hover:underline"
                                        >
                                            Logs
                                        </button>
                                        <button 
                                            onClick={() => handleAdminReset(u.uid)}
                                            className="text-orange-600 hover:text-orange-800 text-xs font-medium hover:underline"
                                        >
                                            Reset
                                        </button>
                                        {u.email !== currentUser.email && (
                                            <button 
                                                onClick={() => handleDeleteClick(u)}
                                                className="text-red-600 hover:text-red-800 text-xs font-medium hover:underline"
                                            >
                                                Delete
                                            </button>
                                        )}
                                    </td>
                                </tr>
                            )))}
                        </tbody>
                    </table>
                </div>

                {/* Mobile Card View */}
                <div className="md:hidden">
                    {activeUsers.length === 0 ? (
                        <div className="text-center py-8 text-gray-400 bg-gray-50">No active users found.</div>
                    ) : (
                        <div className="divide-y divide-gray-100">
                             {activeUsers.map(u => (
                                 <div key={u.uid} className="p-4 hover:bg-gray-50 transition-colors">
                                     <div className="flex items-center justify-between mb-3">
                                         <div className="flex items-center">
                                            <div className="h-10 w-10 rounded-full bg-primary/10 flex items-center justify-center text-primary font-bold mr-3 shadow-sm">
                                                {u.name.charAt(0)}
                                            </div>
                                            <div>
                                                <div className="font-bold text-gray-900">{u.name}</div>
                                                <div className="text-xs text-gray-500">{u.email}</div>
                                            </div>
                                         </div>
                                         <span className={`px-2 py-1 text-[10px] rounded-full font-bold uppercase ${u.role === 'admin' ? 'bg-purple-100 text-purple-700' : 'bg-blue-50 text-blue-600'}`}>
                                            {u.role}
                                        </span>
                                     </div>
                                     
                                     <div className="flex justify-between items-center text-xs text-gray-500 mb-3 px-1">
                                         <div className="flex items-center">
                                             <i className="material-icons text-sm mr-1">history</i>
                                             {u.lastLogin ? new Date(u.lastLogin).toLocaleDateString() : 'Never'}
                                         </div>
                                         <div>
                                            {u.status === 'pending' ? (
                                                <span className="text-yellow-600 font-bold">Pending</span>
                                            ) : (
                                                <span className={`${u.isOnline ? 'text-green-600' : 'text-gray-400'} font-bold flex items-center`}>
                                                    <span className={`w-2 h-2 rounded-full mr-1 ${u.isOnline ? 'bg-green-500' : 'bg-gray-300'}`}></span>
                                                    {u.isOnline ? 'Online' : 'Offline'}
                                                </span>
                                            )}
                                         </div>
                                     </div>
                                     
                                     <div className="flex bg-gray-50 p-2 rounded-lg gap-2 text-xs font-medium">
                                         <button onClick={() => viewLogs(u)} className="flex-1 py-1.5 text-primary bg-white border border-gray-200 rounded shadow-sm hover:bg-gray-50">Logs</button>
                                         <button onClick={() => handleAdminReset(u.uid)} className="flex-1 py-1.5 text-orange-600 bg-white border border-gray-200 rounded shadow-sm hover:bg-gray-50">Reset Pass</button>
                                         {u.email !== currentUser.email && (
                                            <button onClick={() => handleDeleteClick(u)} className="flex-1 py-1.5 text-red-600 bg-white border border-gray-200 rounded shadow-sm hover:bg-gray-50">Delete</button>
                                         )}
                                     </div>
                                 </div>
                             ))}
                        </div>
                    )}
                </div>
            </div>
        </div>
      )}

      {/* Password Modal */}
      <Modal isOpen={showPasswordModal} onClose={() => setShowPasswordModal(false)} title="Change Password" 
        footer={
           <>
              <button onClick={() => setShowPasswordModal(false)} className="px-4 py-2 text-gray-600 hover:bg-gray-100 rounded-lg">Cancel</button>
              <button onClick={handlePasswordChange} className="px-6 py-2 bg-primary text-white rounded-lg">Update</button>
           </>
        }
      >
        <div className="space-y-4">
             <div className="relative">
                 <input type={showPass.current ? "text" : "password"} placeholder="Current Password" value={passForm.current} onChange={e => setPassForm({...passForm, current: e.target.value})} className="w-full p-2 border rounded-lg" />
                 <i onClick={() => toggleVisibility('current')} className="material-icons absolute right-3 top-3 text-gray-400 cursor-pointer text-sm">{showPass.current ? 'visibility' : 'visibility_off'}</i>
             </div>
             <div className="relative">
                 <input type={showPass.new ? "text" : "password"} placeholder="New Password" value={passForm.new} onChange={e => setPassForm({...passForm, new: e.target.value})} className="w-full p-2 border rounded-lg" />
                 <i onClick={() => toggleVisibility('new')} className="material-icons absolute right-3 top-3 text-gray-400 cursor-pointer text-sm">{showPass.new ? 'visibility' : 'visibility_off'}</i>
             </div>
             <div className="relative">
                <input type={showPass.confirm ? "text" : "password"} placeholder="Confirm Password" value={passForm.confirm} onChange={e => setPassForm({...passForm, confirm: e.target.value})} className="w-full p-2 border rounded-lg" />
                <i onClick={() => toggleVisibility('confirm')} className="material-icons absolute right-3 top-3 text-gray-400 cursor-pointer text-sm">{showPass.confirm ? 'visibility' : 'visibility_off'}</i>
             </div>
        </div>
      </Modal>

      {/* Approve Confirmation Modal */}
      <Modal isOpen={showApproveModal} onClose={() => setShowApproveModal(false)} title="Approve User" type="info"
        footer={
            <>
               <button onClick={() => setShowApproveModal(false)} className="px-4 py-2 border rounded-lg hover:bg-gray-50">Cancel</button>
               <button onClick={confirmApprove} className="px-4 py-2 bg-green-600 text-white rounded-lg hover:bg-green-700 shadow-md">Confirm Approve</button>
            </>
        }
      >
          <p>Are you sure you want to approve access for <strong>{selectedUser?.name}</strong>?</p>
          <p className="text-sm text-gray-500 mt-2">They will be granted access to the {selectedUser?.role || 'user'} dashboard immediately.</p>
      </Modal>

      {/* Delete/Reject User Modal */}
      <Modal isOpen={showDeleteModal} onClose={() => setShowDeleteModal(false)} title={selectedUser?.status === 'pending' ? "Reject Application" : "Delete User"} type="danger"
        footer={
            <>
               <button onClick={() => setShowDeleteModal(false)} className="px-4 py-2 border rounded-lg">Cancel</button>
               <button onClick={confirmDelete} className="px-4 py-2 bg-red-600 text-white rounded-lg">Confirm Delete</button>
            </>
        }
      >
          <p>Are you sure you want to {selectedUser?.status === 'pending' ? "reject" : "delete"} user <strong>{selectedUser?.name}</strong>? This will permanently remove them from the system.</p>
      </Modal>

      {/* Reset Result Modal */}
      <Modal isOpen={showResetModal} onClose={() => setShowResetModal(false)} title="Reset Successful" type="info"
        footer={<button onClick={() => setShowResetModal(false)} className="px-6 py-2 bg-primary text-white rounded-lg">OK</button>}
      >
          <p className="text-gray-600 font-medium">{resetResult}</p>
      </Modal>

      {/* View Logs Modal */}
      <Modal isOpen={showLogModal} onClose={() => setShowLogModal(false)} title={`Activity Logs: ${selectedUser?.name}`}
         footer={<button onClick={() => setShowLogModal(false)} className="px-4 py-2 bg-gray-100 rounded-lg">Close</button>}
      >
          <div className="max-h-[300px] overflow-y-auto space-y-3">
              <div className="flex justify-between text-xs font-bold text-gray-500 border-b pb-2">
                  <span>Action</span>
                  <span>Date/Time</span>
              </div>
              {/* Real Logs */}
              {logs.length > 0 ? (
                  logs.map((log, i) => (
                  <div key={i} className="flex justify-between text-sm text-gray-700 py-2 border-b border-gray-50">
                      <div className="flex flex-col">
                          <span className="font-semibold text-xs text-primary">{log.action || 'ACTION'}</span>
                          <span>{log.details || log.message || log.type}</span>
                      </div>
                      <span className="text-gray-500 text-xs whitespace-nowrap ml-4">
                          {log.timestamp ? new Date(log.timestamp).toLocaleString() : 'Just now'}
                      </span>
                  </div>
                  ))
              ) : (
                  <p className="text-center text-gray-400 py-4">No activity logs found for this user.</p>
              )}
          </div>
      </Modal>

      {/* Add User Modal */}
      <Modal isOpen={showAddUserModal} onClose={() => setShowAddUserModal(false)} title="Add New User"
          footer={
             <>
                <button onClick={() => setShowAddUserModal(false)} className="px-4 py-2 border rounded-lg hover:bg-gray-50">Cancel</button>
                <button onClick={handleCreateUser} className="px-6 py-2 bg-primary text-white rounded-lg hover:bg-primary-dark shadow-md">Create User</button>
             </>
          }
      >
          <form className="space-y-4">
              <div>
                  <input type="text" value={addUserForm.name} onChange={e => setAddUserForm({...addUserForm, name: e.target.value})} className="w-full p-2 border rounded-lg focus:ring-primary focus:border-primary" placeholder="Full Name" />
              </div>
              <div>
                  <input type="email" value={addUserForm.email} onChange={e => setAddUserForm({...addUserForm, email: e.target.value})} className="w-full p-2 border rounded-lg focus:ring-primary focus:border-primary" placeholder="Email (e.g., user@cics.edu)" />
              </div>
              <div>
                  <input type="password" value={addUserForm.password} onChange={e => setAddUserForm({...addUserForm, password: e.target.value})} className="w-full p-2 border rounded-lg focus:ring-primary focus:border-primary" placeholder="Password" />
              </div>
              <div>
                  <select value={addUserForm.role} onChange={e => setAddUserForm({...addUserForm, role: e.target.value})} className="w-full p-2 border rounded-lg focus:ring-primary focus:border-primary bg-white" aria-label="Role">
                      <option value="user">User</option>
                      <option value="admin">Admin</option>
                  </select>
              </div>
          </form>
      </Modal>

      {/* Error Modal */}
      <Modal isOpen={showErrorModal} onClose={() => setShowErrorModal(false)} title="Error" type="danger"
          footer={<button onClick={() => setShowErrorModal(false)} className="px-6 py-2 bg-red-600 text-white rounded-lg hover:bg-red-700">Close</button>}
      >
          <p className="text-gray-700">{errorMessage}</p>
      </Modal>

      {/* Success Modal */}
      <Modal isOpen={showSuccessModal} onClose={() => setShowSuccessModal(false)} title="Success" type="info" // Using info styling which is blue, but we can assume it's fine or customize Modal later
          footer={<button onClick={() => setShowSuccessModal(false)} className="px-6 py-2 bg-primary text-white rounded-lg">OK</button>}
      >
          <div className="flex items-center text-green-600">
             <i className="material-icons mr-2">check_circle</i>
             <p className="font-medium text-gray-800">{successMessage}</p>
          </div>
      </Modal>
    </div>
  );
};

export default ManageAccount;

