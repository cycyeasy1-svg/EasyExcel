import { defineConfig } from 'vitest/config';

// Standalone config: intentionally NOT reusing vite.config.ts, whose plugin
// side-builds the extension host via esbuild on startup.
export default defineConfig({
    test: {
        include: ['test/**/*.test.ts'],
        server: {
            deps: {
                // Fork ships only a browser ESM build with extensionless
                // core-js imports; must be transformed by vite, not node.
                inline: ['@cweijan/exceljs'],
            },
        },
    },
});
