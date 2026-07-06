import { build, context } from 'esbuild'

const isProd = process.argv.indexOf('--mode=production') >= 0;

function createBuildNoticePlugin() {
    return {
        name: 'build notice',
        setup(build) {
            build.onStart(() => {
                console.log('extension build start')
            })
            build.onEnd((result) => {
                if (result.errors.length === 0) {
                    console.log('extension build success')
                }
            })
        }
    };
}

async function buildDesktopExtension() {
    const options = {
        entryPoints: ['./src/extension.ts'],
        bundle: true,
        outfile: 'out/extension.js',
        external: ['vscode'],
        format: 'cjs' as const,
        platform: 'node' as const,
        minify: isProd,
        sourcemap: !isProd,
        plugins: [
            createBuildNoticePlugin(),
        ],
    };

    if (!isProd) {
        const ctx = await context(options);
        await ctx.watch();
        return;
    }

    await build(options);
}

void buildDesktopExtension().catch(() => {
    process.exit(1);
});
