import { useState } from 'react'
import { BrowserRouter as Router, Routes, Route, Navigate } from 'react-router-dom'
import Login from './pages/Login'
import AdminDashboard from './pages/AdminDashboard'
import UserDashboard from './pages/UserDashboard'

import { AuthProvider, useAuth } from './context/AuthContext';
import { AppProvider } from './context/AppContext';
import OfflineIndicator from './components/common/OfflineIndicator';

import DashboardSkeleton from './components/DashboardSkeleton';

// Protected Route Wrapper
const ProtectedRoute = ({ children, allowedRole }) => {
  const { currentUser, loading } = useAuth();
  
  if (loading) {
      return <DashboardSkeleton />;
  }

  if (!currentUser) {
    return <Navigate to="/" replace />;
  }

  if (allowedRole && currentUser.role !== allowedRole) {
    return <Navigate to="/" replace />; // Or unauthorized page
  }

  return children;
};

function App() {
  return (
    <AuthProvider>
      <AppProvider>
        <OfflineIndicator />
        <Router>
          <Routes>
            <Route path="/" element={<Login />} />
            
            <Route path="/admin-dashboard" element={
              <ProtectedRoute allowedRole="admin">
                <AdminDashboard />
              </ProtectedRoute>
            } />
            
            <Route path="/user-dashboard" element={
              <ProtectedRoute allowedRole="user">
                <UserDashboard />
              </ProtectedRoute>
            } />
            
            {/* Fallback */}
            <Route path="*" element={<Navigate to="/" replace />} />
          </Routes>
        </Router>
      </AppProvider>
    </AuthProvider>
  )
}

export default App
