import React from "react";
import ReactDOM from "react-dom/client";
import App from "./App.jsx";
import "./index.css";

// Chrome avisa una sola vez, apenas carga la página, que la app se puede instalar: se
// guarda ese aviso para ofrecer el botón "Instalar" cuando la app termine de cargar. Así
// se instala como app (con el ícono completo) y no como un acceso directo.
window.addEventListener("beforeinstallprompt", (e) => {
  e.preventDefault();
  window.__avisoInstalar = e;
  window.dispatchEvent(new Event("concretar-instalable"));
});
window.addEventListener("appinstalled", () => {
  window.__avisoInstalar = null;
  window.dispatchEvent(new Event("concretar-instalable"));
});

ReactDOM.createRoot(document.getElementById("root")).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>
);
