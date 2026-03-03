/** @type {import('tailwindcss').Config} */
export default {
  content: [
    "./index.html",
    "./src/**/*.{js,ts,jsx,tsx}",
  ],
  theme: {
    extend: {
      colors: {
        'strategy-blue': '#d7e4ee',
        'deep-purple': '#4d217a',
      },
    },
  },
  plugins: [],
}