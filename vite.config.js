import { defineConfig } from 'vite';

export default defineConfig({
    base: './', // Quan trọng cho GitHub Pages
    build: {
        outDir: 'dist',
        rollupOptions: {
            input: {
                main: './index.html',
            },
        },
    },
    server: {
        port: 3000,
    },
});
