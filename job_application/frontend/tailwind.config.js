/** @type {import('tailwindcss').Config} */
export default {
  darkMode: 'class',
  content: ['./index.html', './src/**/*.{js,jsx}'],
  theme: {
    extend: {
      colors: {
        base: {
          950: '#0b1120',
          900: '#0f172a',
          800: '#1e293b',
        },
        glass: 'rgba(30, 41, 59, 0.55)',
        accent: '#7c3aed',
        accentSoft: '#c4b5fd',
      },
      boxShadow: {
        card: '0 20px 40px rgba(15, 23, 42, 0.45)',
      },
    },
  },
  plugins: [],
}
