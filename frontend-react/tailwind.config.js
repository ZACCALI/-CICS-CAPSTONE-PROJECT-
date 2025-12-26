/** @type {import('tailwindcss').Config} */
export default {
  content: [
    "./index.html",
    "./src/**/*.{js,ts,jsx,tsx}",
  ],
  theme: {
    extend: {
      fontFamily: {
        'sans': ['Roboto', 'sans-serif'],
      },
      colors: {
        primary: {
            DEFAULT: '#144df7', // New Bright Blue
            dark: '#0d32a6',    
            light: '#4b7afd',   
        },
        secondary: '#64748B', // Slate 500
        dark: '#0F172A', // Slate 900
      }
    },
  },
  plugins: [],
}

