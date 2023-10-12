import dts from "vite-plugin-dts";
import path from "path";
import { defineConfig } from "vite";

export default defineConfig({
  base: "./",
  plugins: [dts({ rollupTypes: true })],
  build: {
    sourcemap: true,
    lib: {
      entry: path.resolve(__dirname, "src/index.ts"),
      name: "sp-propertybag",
      formats: ["es"],
      fileName: "index.js",
    },
  },
});
