import { defineConfig } from "vite";
import { resolve } from "node:path";

export default defineConfig({
  root: "src",
  base: "./",
  publicDir: resolve(__dirname, "public"),
  build: {
    outDir: resolve(__dirname, "dist"),
    emptyOutDir: true,
    target: "safari15",
    rollupOptions: {
      input: {
        taskpane: resolve(__dirname, "src/taskpane.html"),
        commands: resolve(__dirname, "src/commands.html"),
      },
    },
  },
  test: { root: __dirname, include: ["test/**/*.test.ts"] },
} as any);
