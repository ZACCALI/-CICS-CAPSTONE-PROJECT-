import React, { useEffect, useState } from 'react';
import { createPortal } from 'react-dom';

const Modal = ({ isOpen, onClose, title, children, footer, type = 'default' }) => {
  const [isAnimating, setIsAnimating] = useState(false);

  useEffect(() => {
    if (isOpen) {
      setIsAnimating(true);
      document.body.style.overflow = 'hidden';
    } else {
        setTimeout(() => setIsAnimating(false), 300); // Wait for transition
        document.body.style.overflow = 'unset';
    }
    return () => {
        document.body.style.overflow = 'unset';
    };
  }, [isOpen]);

  if (!isOpen && !isAnimating) return null;

  const typeStyles = {
      default: 'border-t-4 border-t-primary border-x border-b border-gray-200',
      danger: 'border-t-4 border-t-red-500 border-x border-b border-gray-200',
      success: 'border-t-4 border-t-green-500 border-x border-b border-gray-200'
  };

  const headerColors = {
      default: 'text-gray-800',
      danger: 'text-red-600',
      success: 'text-green-600'
  };

  return createPortal(
    <div className={`fixed inset-0 z-[100] flex items-center justify-center p-4 transition-opacity duration-300 ${isOpen ? 'opacity-100' : 'opacity-0 pointer-events-none'}`}>
      
      {/* Backdrop */}
      <div 
        className="absolute inset-0 bg-black/50 backdrop-blur-sm transition-opacity duration-300" 
        onClick={onClose}
      />

      {/* Modal Content */}
      <div className={`bg-white rounded-xl shadow-2xl w-full max-w-lg transform transition-all duration-300 ${isOpen ? 'scale-100 translate-y-0' : 'scale-95 translate-y-4'} ${typeStyles[type]} overflow-hidden`}>
        
        {/* Header */}
        <div className="px-6 py-4 border-b border-gray-100 flex justify-between items-center bg-gray-50/50">
           <h3 className={`text-lg font-bold flex items-center ${headerColors[type]}`}>
               {title}
           </h3>
           <button 
              onClick={onClose} 
              className="text-gray-400 hover:text-gray-600 transition-colors p-1 hover:bg-gray-100 rounded-full"
           >
               <i className="material-icons text-xl">close</i>
           </button>
        </div>

        {/* Body */}
        <div className="p-6">
            {children}
        </div>

        {/* Footer */}
        {footer && (
            <div className="px-6 py-4 bg-gray-50 border-t border-gray-100 flex justify-end space-x-3">
                {footer}
            </div>
        )}
      </div>
    </div>,
    document.body
  );
};

export default Modal;
