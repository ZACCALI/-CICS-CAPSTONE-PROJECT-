import React from 'react';

const DashboardSkeleton = () => {
  return (
    <div className="min-h-screen bg-gray-50 flex flex-col overflow-hidden">
      {/* YouTube-style Top Loading Bar */}
      <div className="fixed top-0 left-0 w-full h-1 z-[100] bg-gray-200 overflow-hidden">
        <div className="h-full bg-primary animate-loading-bar"></div>
      </div>

      {/* Header Skeleton */}
      <div className="h-16 bg-white border-b border-gray-100 flex items-center justify-between px-4 fixed top-0 w-full z-30">
        <div className="flex items-center gap-4">
             <div className="w-8 h-8 bg-gray-200 rounded-md animate-pulse"></div>
             <div className="w-32 h-6 bg-gray-200 rounded animate-pulse hidden md:block"></div>
        </div>
        <div className="flex items-center gap-4">
             <div className="w-8 h-8 bg-gray-200 rounded-full animate-pulse"></div>
             <div className="w-10 h-10 bg-gray-200 rounded-full animate-pulse"></div>
        </div>
      </div>

      <div className="flex flex-1 pt-16 h-screen">
        {/* Sidebar Skeleton */}
        <div className="hidden md:block w-20 bg-white border-r border-gray-100 h-full fixed top-16 left-0 z-20 flex flex-col items-center py-6 gap-6">
            {[1, 2, 3, 4, 5, 6].map(i => (
                <div key={i} className="w-10 h-10 bg-gray-200 rounded-lg animate-pulse"></div>
            ))}
        </div>

        {/* Main Content Skeleton */}
        <main className="flex-1 p-6 md:ml-20 overflow-hidden">
            <div className="max-w-7xl mx-auto space-y-8">
                {/* Stats Grid Skeleton */}
                <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
                    {[1, 2, 3].map(i => (
                        <div key={i} className="bg-white p-6 rounded-2xl border border-gray-100 shadow-sm h-32 flex items-center gap-4">
                            <div className="w-14 h-14 bg-gray-200 rounded-xl animate-pulse"></div>
                            <div className="space-y-2 flex-1">
                                <div className="h-4 bg-gray-200 rounded w-1/2 animate-pulse"></div>
                                <div className="h-8 bg-gray-200 rounded w-3/4 animate-pulse"></div>
                            </div>
                        </div>
                    ))}
                </div>

                {/* Main Panel Skeleton */}
                <div className="bg-white h-[400px] rounded-2xl border border-gray-100 shadow-sm p-6 animate-pulse">
                     <div className="h-8 bg-gray-200 rounded w-1/4 mb-6"></div>
                     <div className="space-y-4">
                         <div className="h-4 bg-gray-200 rounded w-full"></div>
                         <div className="h-4 bg-gray-200 rounded w-full"></div>
                         <div className="h-4 bg-gray-200 rounded w-3/4"></div>
                     </div>
                </div>
            </div>
        </main>
      </div>
    </div>
  );
};

export default DashboardSkeleton;
