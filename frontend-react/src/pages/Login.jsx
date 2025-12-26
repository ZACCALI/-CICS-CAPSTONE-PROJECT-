import React, { useState, useEffect } from 'react';
import { useNavigate } from 'react-router-dom';
import { useAuth } from '../context/AuthContext';
import { useApp } from '../context/AppContext';

const Login = () => {
  const navigate = useNavigate();
  const { login, signup, currentUser, loading } = useAuth();
  const { logActivity } = useApp();

  useEffect(() => {
      if (!loading && currentUser) {
          // If pending, do NOT redirect. Let the registration success message show.
          if (currentUser.status === 'pending') {
              return;
          }

          // Safety check: ensure role is defined to avoid undefined behavior
          if (currentUser.role) {
            const target = currentUser.role === 'admin' ? '/admin-dashboard' : '/user-dashboard';
            navigate(target, { replace: true });
          } else {
            console.warn("User logged in but no role found. Waiting or handling error.");
            // Optionally could force a fetch or logout here, but letting it sit might allow subsequent updates to fix it.
          }
      }
  }, [currentUser, loading, navigate]);

  const [isLoginMode, setIsLoginMode] = useState(true);

  // Form State
  const [username, setUsername] = useState(''); // Serves as email for signup too
  const [password, setPassword] = useState('');
  const [name, setName] = useState(''); // For signup
  
  const [remember, setRemember] = useState(true);
  const [showPassword, setShowPassword] = useState(false);
  
  // UI State
  const [formLoading, setFormLoading] = useState(false); // Renamed from loading
  const [error, setError] = useState('');
  const [success, setSuccess] = useState(false);
  const [registrationSuccess, setRegistrationSuccess] = useState(false);
  const [shake, setShake] = useState(false);

  const validateEmail = (email) => /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);

  const handleSubmit = async (e) => {
    e.preventDefault();
    setError('');

    // Basic Validation
    if (!username || !password || (!isLoginMode && !name)) {
      showError('Please fill in all fields');
      return;
    }
    
    if (!validateEmail(username)) {
      showError('Please enter a valid email address');
      return;
    }
    
    if (password.length < 6) {
      showError('Password must be at least 6 characters');
      return;
    }

    setFormLoading(true);

    try {
      if (isLoginMode) {
         const user = await login(username.trim(), password.trim(), remember);
         
         // Log Login Activity - REMOVED per request
         // if (logActivity && user.name) {
         //     logActivity(user.name, 'Logged In', 'Session', `User logged in from Login Page`);
         // }

         setSuccess(true);
         setTimeout(() => navigate(user.redirect), 1000);
      } else {
         // Signup Flow
         const user = await signup({ email: username, password, name });
         
         // Log Signup Activity
         if (logActivity) {
             logActivity(name, 'Registered', 'Account', `User registered and is pending approval`);
         }
         
         setRegistrationSuccess(true);
         setFormLoading(false);
      }
    } catch (err) {
      console.error("Login Error:", err);
      if (err.code === 'auth/invalid-credential' || err.message.includes('invalid-credential')) {
          showError('Invalid email or password. Please try again.');
      } else if (err.code === 'auth/email-already-in-use') {
          showError('Email is already in use.');
      } else if (err.code === 'auth/weak-password') {
          showError('Password is too weak (min 6 chars).');
      } else {
          showError(err.message || 'An unexpected error occurred.');
      }
    } finally {
        if (!success) setFormLoading(false);
    }
  };

  const showError = (msg) => {
    setError(msg);
  };

  const toggleMode = () => {
      setIsLoginMode(!isLoginMode);
      setError('');
      setSuccess(false);
  };

  return (
    <div className="flex min-h-screen font-sans flex-col md:flex-row">
      {/* Left Section */}
      <div className="flex-1 bg-gradient-to-br from-primary to-primary-dark flex flex-col justify-center items-center p-10 text-center text-white">
        <div className="bg-white w-[100px] h-[100px] rounded-full flex items-center justify-center mb-8">
          <i className="material-icons text-[50px] text-primary">campaign</i>
        </div>
        <h1 className="text-[2rem] mb-4 font-bold font-sans">CICS PA SYSTEM</h1>
        <p className="max-w-[400px] leading-relaxed opacity-90 font-normal">
          Welcome to the Public Address System. Manage announcements and notifications with ease across your campus.
        </p>
      </div>

      {/* Right Section */}
      <div className="flex-1 flex justify-center items-center p-10 bg-white md:rounded-none rounded-t-[30px] -mt-[30px] md:mt-0 shadow-[0_-10px_30px_rgba(0,0,0,0.1)] md:shadow-none">
        <div className="w-full max-w-[380px]">
          <h2 className="text-[1.8rem] mb-8 text-primary font-bold font-sans">
              {isLoginMode ? 'Welcome Back!' : 'Create Account'}
          </h2>
          
          
          {registrationSuccess ? (
              <div className="text-center animate-fade-in p-6">
                  <div className="w-16 h-16 bg-green-100 rounded-full flex items-center justify-center mx-auto mb-4">
                      <i className="material-icons text-green-600 text-3xl">check_circle</i>
                  </div>
                  <h3 className="text-xl font-bold text-gray-800 mb-2">Account Created!</h3>
                  <p className="text-gray-600 mb-6 leading-relaxed">
                      Your account has been successfully registered. <br/>
                      <span className="font-semibold text-primary">Pending Admin Approval</span>
                  </p>
                  <p className="text-sm text-gray-500 mb-8">
                      You will be able to log in once an administrator approves your account.
                  </p>
                  <button 
                      onClick={() => {
                          setRegistrationSuccess(false);
                          setIsLoginMode(true);
                          setError('');
                          setSuccess(false);
                          setPassword('');
                      }}
                      className="w-full p-3 bg-primary text-white rounded-lg font-medium hover:bg-primary-dark transition-colors"
                  >
                      Back to Login
                  </button>
              </div>
          ) : (
          <form onSubmit={handleSubmit} className="" autoComplete="off">
            <input type="hidden" value="something" /> {/* Hack to trick Chrome */}
            
            {/* Full Name for Signup */}
            {!isLoginMode && (
                <div className="mb-6 relative animation-fade-in">
                <input 
                    type="text" 
                    value={name}
                    onChange={(e) => setName(e.target.value)}
                    required={!isLoginMode}
                    placeholder="Full Name"
                    autoComplete="name"
                    className="w-full p-4 border border-slate-200 rounded-lg text-base transition-all duration-300 font-sans bg-slate-50 focus:border-primary focus:outline-none focus:ring-2 focus:ring-primary/20 focus:bg-white"
                />
                </div>
            )}

            <div className="mb-6 relative">
              <input 
                type="email" 
                value={username}
                onChange={(e) => setUsername(e.target.value)}
                required 
                placeholder="Email Address"
                autoComplete="new-email" // browser hack
                className="w-full p-4 border border-slate-200 rounded-lg text-base transition-all duration-300 font-sans bg-slate-50 focus:border-primary focus:outline-none focus:ring-2 focus:ring-primary/20 focus:bg-white"
              />
            </div>

            <div className="mb-6 relative flex items-center">
              <input 
                type={showPassword ? "text" : "password"} 
                value={password}
                onChange={(e) => setPassword(e.target.value)}
                required 
                placeholder="Password"
                autoComplete="new-password"
                className="w-full p-4 pr-12 border border-slate-200 rounded-lg text-base transition-all duration-300 font-sans bg-slate-50 focus:border-primary focus:outline-none focus:ring-2 focus:ring-primary/20 focus:bg-white"
              />
              <button 
                type="button" 
                onClick={() => setShowPassword(!showPassword)}
                className="absolute right-4 text-gray-400 hover:text-gray-600 bg-transparent border-none cursor-pointer"
              >
                <i className="material-icons">{showPassword ? 'visibility' : 'visibility_off'}</i>
              </button>
            </div>

            {error && (
              <div className="flex items-center p-3 bg-red-50 text-red-700 rounded-lg mb-5 font-semibold">
                <i className="material-icons mr-2 text-xl">error_outline</i>
                <span>{error}</span>
              </div>
            )}

            {isLoginMode && (
                <div className="flex items-center mb-6">
                <label className="flex items-center cursor-pointer relative pl-8 select-none font-medium text-gray-700 hover:text-gray-900 group">
                    <input 
                    type="checkbox" 
                    checked={remember} 
                    onChange={(e) => setRemember(e.target.checked)}
                    className="absolute opacity-0 cursor-pointer h-0 w-0 peer"
                    />
                    <span className="absolute left-0 top-0 h-5 w-5 bg-white border border-gray-300 rounded peer-checked:bg-primary peer-checked:border-primary transition-all group-hover:bg-gray-50"></span>
                    <span className="absolute left-[7px] top-[3px] w-[5px] h-[10px] border-solid border-white border-r-2 border-b-2 rotate-45 hidden peer-checked:block"></span>
                    Remember me
                </label>
                </div>
            )}

            <button 
              type="submit" 
              disabled={formLoading}
              className={`w-full p-4 text-white rounded-lg text-base font-semibold cursor-pointer transition-all duration-300 flex items-center justify-center shadow-lg shadow-primary/30 tracking-wide 
                ${success ? 'bg-green-500 hover:bg-green-600' : 'bg-primary hover:bg-primary-dark hover:-translate-y-0.5 shadow-xl'} 
                ${formLoading && !success ? 'opacity-70 cursor-not-allowed transform-none' : ''}`}
            >
              {formLoading ? (
                <>
                  <i className="material-icons mr-2 animate-spin">hourglass_empty</i> {isLoginMode ? 'AUTHENTICATING...' : 'CREATING ACCOUNT...'}
                </>
              ) : success ? (
                <>
                  <i className="material-icons mr-2">check</i> SUCCESS!
                </>
              ) : (
                <>
                  <i className="material-icons mr-2">{isLoginMode ? 'login' : 'person_add'}</i> {isLoginMode ? 'LOG IN' : 'SIGN UP'}
                </>
              )}
            </button>

            <div className="text-center mt-6">
                <button 
                  type="button" 
                  onClick={toggleMode}
                  className="text-primary hover:text-primary-dark font-medium hover:underline focus:outline-none"
                >
                    {isLoginMode ? "Don't have an account? Sign Up" : "Already have an account? Log In"}
                </button>
            </div>

            <div className="text-center mt-8 pt-5 border-t border-gray-100 text-gray-500 text-sm">
              <p>© 2025 CICS Public Address System</p>
              <p>v2.5.0</p>
            </div>
          </form>
          )}
        </div>
      </div>
    </div>
  );
};

export default Login;
