import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import tailwindcss from "@tailwindcss/vite";
import path from "node:path";

export default defineConfig({
  plugins: [react(), tailwindcss()],
  resolve: {
    alias: {
      "@": path.resolve(__dirname, "src"),
    },
  },
  build: {
    outDir: path.resolve(__dirname, "../app/static/ui"),
    emptyOutDir: true,
    manifest: true,
    rollupOptions: {
      input: {
        "task-new": path.resolve(__dirname, "src/entries/task-new.tsx"),
        "task-list": path.resolve(__dirname, "src/entries/task-list.tsx"),
        "task-detail": path.resolve(__dirname, "src/entries/task-detail.tsx"),
      },
      output: {
        entryFileNames: "[name].js",
        chunkFileNames: "chunks/[name]-[hash].js",
        assetFileNames: "assets/[name]-[hash][extname]",
      },
    },
  },
});
