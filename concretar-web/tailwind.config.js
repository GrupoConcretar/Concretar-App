// Los estilos se generan al compilar (antes se compilaban en el navegador con el CDN de
// Tailwind cada vez que se abría la app, lo que la hacía más lenta, sobre todo en el celular).
/** @type {import('tailwindcss').Config} */
export default {
  content: ["./index.html", "./src/**/*.{js,jsx}"],
  theme: { extend: {} },
  plugins: [],
};
