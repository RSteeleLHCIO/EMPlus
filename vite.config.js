import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import { visualizer } from 'rollup-plugin-visualizer';
import { computeBuildInfo } from './server/buildInfo.js';

// https://vitejs.dev/config/
export default defineConfig({
    plugins: [react()],
    define: {
        __CLIENT_BUILD_INFO__: JSON.stringify(computeBuildInfo()),
    },
    build: {
        rollupOptions: {
            plugins: [
                visualizer({ filename: 'dist/bundle-analysis.html', template: 'treemap', gzipSize: true, brotliSize: true })
            ]
        }
    }
});
