import React from "react";
import ReactDOM from "react-dom/client";
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

// El link privado de un cliente (?portal=...) abre solo la página de su obra: se carga
// aparte, sin bajar la app completa ni pedir usuario.
const tokenPortal = new URLSearchParams(window.location.search).get("portal");
const root = ReactDOM.createRoot(document.getElementById("root"));

if (tokenPortal) {
  import("./portalCliente.jsx").then(({ default: PortalCliente }) => {
    root.render(
      <React.StrictMode>
        <PortalCliente token={tokenPortal} />
      </React.StrictMode>
    );
  });
} else {
  import("./App.jsx").then(({ default: App }) => {
    root.render(
      <React.StrictMode>
        <App />
      </React.StrictMode>
    );
  });
}
