import React, { useState, useEffect } from 'react';

const OfflineIndicator = () => {
    const [isOnline, setIsOnline] = useState(navigator.onLine);

    useEffect(() => {
        const handleOnline = () => setIsOnline(true);
        const handleOffline = () => setIsOnline(false);

        window.addEventListener('online', handleOnline);
        window.addEventListener('offline', handleOffline);

        return () => {
            window.removeEventListener('online', handleOnline);
            window.removeEventListener('offline', handleOffline);
        };
    }, []);

    if (isOnline) return null;

    return (
        <div className="fixed bottom-4 left-4 z-50 animate-bounce-in">
            <div className="bg-gray-900/90 text-white px-6 py-3 rounded-lg shadow-lg flex items-center backdrop-blur-sm border border-gray-700">
                <i className="material-icons text-amber-400 mr-3">wifi_off</i>
                <div>
                   <p className="font-bold text-sm">You are currently offline</p>
                   <p className="text-xs text-gray-300">Changes will sync when connection is restored.</p>
                </div>
            </div>
        </div>
    );
};

export default OfflineIndicator;
