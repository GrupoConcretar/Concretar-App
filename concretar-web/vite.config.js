import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

// Las librerías que casi nunca cambian van en archivos aparte: así, cuando se publica un
// cambio de la app, el navegador vuelve a bajar solo el código de la app (lo demás ya lo
// tiene guardado).
const GRUPOS = [
  ["libs-react", ["/react/", "/react-dom/", "/scheduler/"]],
  ["libs-supabase", ["/@supabase/"]],
  ["libs-iconos", ["/lucide-react/"]],
];

export default defineConfig({
  plugins: [react()],
  build: {
    chunkSizeWarningLimit: 900,
    rollupOptions: {
      output: {
        manualChunks(id) {
          if (!id.includes("node_modules")) return undefined;
          for (const [nombre, partes] of GRUPOS) if (partes.some((p) => id.includes(p))) return nombre;
          return undefined;
        },
      },
    },
  },
});
