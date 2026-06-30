import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

export default defineConfig({
  plugins: [react()],
  // IMPORTANTE: troque "orcacpu" pelo nome exato do seu repositório no GitHub
  base: "/orcacpu/",
});
