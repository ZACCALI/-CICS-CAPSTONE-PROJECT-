import React, { createContext, useContext, useState, useEffect } from 'react';
import { 
  signInWithEmailAndPassword, 
  createUserWithEmailAndPassword, 
  signOut, 
  onAuthStateChanged,
  updateProfile,
  updatePassword,
  sendPasswordResetEmail,
  setPersistence,
  browserLocalPersistence,
  browserSessionPersistence
} from "firebase/auth";
import { doc, setDoc, getDoc, updateDoc, onSnapshot, addDoc, collection, serverTimestamp } from "firebase/firestore";
import { auth, db } from '../firebase';
import api from '../api/axios';

const AuthContext = createContext();

export const useAuth = () => useContext(AuthContext);

export const AuthProvider = ({ children }) => {
  const [currentUser, setCurrentUser] = useState(null);
  const [userData, setUserData] = useState(null); // Firestore data (role, status)
  const [loading, setLoading] = useState(true);

  // Load User Session and Role
  useEffect(() => {
    const unsubscribe = onAuthStateChanged(auth, (user) => {
      if (user) {
        // Real-time listener for user data
        const userDocRef = doc(db, "users", user.uid);
        
        // Return the inner unsubscribe function to clean up this listener when auth state changes
        const unsubDoc = onSnapshot(userDocRef, (docSnap) => {
            if (docSnap.exists()) {
                const data = docSnap.data();
                setUserData(data);
                // Merge auth user and firestore data
                setCurrentUser({ ...user, ...data });
                
                // IMPORTANT: Status Check - Force logout if pending/banned
                if (data.status === 'pending') {
                    // Logic to handle pending state updates live? 
                    // Maybe we don't force logout immediately to allow UI to show "Pending" message?
                    // But typically we want them out or restricted. 
                    // Login function blocks them initially. 
                    // If they are approved while on the page, they gain access!
                }
            } else {
                 // First time user or doc missing?
                setUserData({ role: 'user', status: 'pending' });
                setCurrentUser({ ...user, role: 'user', status: 'pending' });
            }
            setLoading(false);
        }, (error) => {
            console.error("Error fetching user data:", error);
            setCurrentUser(user);
            setLoading(false);
        });
        
        return () => unsubDoc(); 
        
      } else {
        setCurrentUser(null);
        setUserData(null);
        setLoading(false);
      }
    });

    return unsubscribe;
  }, []);

  const login = async (email, password, remember = false) => {
    // Set Persistence based on user preference
    await setPersistence(auth, remember ? browserLocalPersistence : browserSessionPersistence);
    
    const result = await signInWithEmailAndPassword(auth, email, password);
    
    // Explicitly fetch doc to return robust user object immediately if needed
    // (Though onAuthStateChanged will trigger too)
    const userDocRef = doc(db, "users", result.user.uid);
    const userDoc = await getDoc(userDocRef);
    let data = {};
    if (userDoc.exists()) data = userDoc.data();

    // Block pending users
    if (data.status === 'pending') {
         await signOut(auth);
         throw new Error("Account is pending approval. Please contact Admin.");
    }
    
    // Update lastLogin and Status (Online)
    await updateDoc(userDocRef, {
        lastLogin: new Date().toISOString(),
        isOnline: true
    });
    
    // Data has changed, let's refresh our local data object
    data.lastLogin = new Date().toISOString();
    data.isOnline = true;
    
    const fullUser = { ...result.user, ...data };
    // Add redirect helper
    fullUser.redirect = data.role === 'admin' ? '/admin-dashboard' : '/user-dashboard';
    
    return fullUser;
  };

  const signup = async ({ email, password, name }) => {
    const result = await createUserWithEmailAndPassword(auth, email, password);
    const user = result.user;

    // Update Auth Profile
    await updateProfile(user, { displayName: name });

    // Create Firestore Document
    const defaultRole = 'user'; 
    const newUserDoc = {
        name,
        email,
        role: defaultRole,
        status: 'pending', 
        createdAt: new Date().toISOString(),
        isOnline: false,
    };

    try {
        await setDoc(doc(db, "users", user.uid), newUserDoc);
    } catch (firestoreError) {
        console.error("Firestore creation failed, rolling back Auth user:", firestoreError);
        // Rollback: Delete the user from Auth so they can try again
        try {
            await user.delete();
        } catch (deleteErr) {
            console.error("Rollback failed (User manual cleanup required):", deleteErr);
        }
        throw new Error("Registration failed. Please try again.");
    }

    // FORCE SIGNOUT - User is pending approval
    await signOut(auth);

    // Provide robust response
    return { 
        ...user, 
        ...newUserDoc, 
        redirect: '/user-dashboard' 
    };
  };

  const logout = async () => {
    const user = auth.currentUser;
    if (user) {
        try {
            // 1. Set offline
            await updateDoc(doc(db, "users", user.uid), {
                isOnline: false
            });

            // 2. Log Logout Activity - REMOVED per request
            /*
            const userName = user.displayName || user.email || 'Unknown User';
            try {
                await addDoc(collection(db, "logs"), {
                    user: userName,
                    action: 'Logged Out',
                    type: 'Session',
                    details: `${userName} logged out of the system.`,
                    timestamp: serverTimestamp()
                });
            } catch (logErr) {
                console.error("Failed to log logout:", logErr);
            }
            */

        } catch (e) {
            console.error("Logout cleanup failed", e);
        }
    }
    // 3. Sign Out
    await signOut(auth);
    
    // 4. Force Reload/Redirect to ensure clean state
    window.location.href = '/'; 
  };

  // Admin Methods - Now calling Backend API
  const getAllUsers = async () => {
    const response = await api.get('/account/');
    return response.data;
  };

  const adminApproveUser = async (uid) => {
      await api.put(`/account/approve/${uid}`);
      return true;
  };

  const adminCreateUser = async (data) => {
      const res = await api.post('/account/create', data);
      return res.data;
  };

  const adminResetPassword = async (uid) => {
     await api.post(`/account/reset/${uid}`);
     return "12345678"; // As defined in backend
  };

  const adminDeleteUser = async (uid) => {
      await api.delete(`/account/${uid}`);
      return true;
  };
  
  const getSystemLogs = async () => {
      try {
         const res = await api.get('/realtime/logs');
         return res.data;
      } catch(e) {
         console.error("Failed to fetch logs:", e);
         return [];
      }
  };

  // New helper for password reset (email)
  const resetPasswordEmail = (email) => {
      return sendPasswordResetEmail(auth, email);
  };
 
  const updateUser = async (data) => {
      const user = auth.currentUser;
      if (!user) return;

      try {
          // 1. Password Update
          if (data.password) {
              await updatePassword(user, data.password);
          }

          // 2. Profile Update (Auth)
          if (data.name || data.avatar) {
              await updateProfile(user, { 
                  displayName: data.name,
                  photoURL: data.avatar 
              });
          }

          // 3. Firestore Update
          // Filter out password from firestore data
          const { password, ...firestoreData } = data;
          if (Object.keys(firestoreData).length > 0) {
              await updateDoc(doc(db, "users", user.uid), firestoreData);
          }

          // 4. Update Local State
          setCurrentUser(prev => ({ ...prev, ...data }));
          return true;
      } catch (error) {
          console.error("Error updating user:", error);
          throw error;
      }
  };

  // Special Admin Update (Bypasses Auth restrictions via Backend)
  const adminUpdateProfile = async (data) => {
      try {
          await api.put('/account/profile', data);
          // Manually update local state since Auth/Firestore listeners might lag slightly
          // or if we changed email, we want UI to reflect it immediately
          setCurrentUser(prev => ({ ...prev, ...data }));
          return true;
      } catch (e) {
          console.error("Admin update failed:", e);
          throw e; // Rethrow to show error in UI
      }
  };

  const value = {
    currentUser,
    loading,
    login,
    signup,
    logout,
    updateUser, // Restored
    adminUpdateProfile, // New
    getAllUsers, 
    adminCreateUser,
    adminApproveUser,
    adminDeleteUser,
    adminResetPassword,
    getSystemLogs, // Added
    resetPasswordEmail
  };

  return (
    <AuthContext.Provider value={value}>
      {children}
    </AuthContext.Provider>
  );
};

