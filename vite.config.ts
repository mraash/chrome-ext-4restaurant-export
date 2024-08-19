import { defineConfig } from 'vite';

export default defineConfig({
  build: {
    rollupOptions: {
      input: {
        content: 'src/content.ts',
        background: 'src/background.ts',
      },
      output: {
        entryFileNames: '[name].js',
      },
    },
  },
  resolve: {
    alias: {
      '~root': 'src/',
    }
  }
});
