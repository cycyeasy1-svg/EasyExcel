import { rmSync } from 'node:fs'
import { resolve } from 'node:path'
import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

const cwd = process.cwd()
const argv = process.argv
const isProdBuild = argv.includes('build') && argv.some((arg) => arg.includes('production'))

if (isProdBuild) {
  rmSync(resolve(cwd, 'out'), { recursive: true, force: true })
}

if (argv.join(',').includes('mode')) {
  void import('./build')
}

// https://vitejs.dev/config/
export default defineConfig(({ command, mode }) => ({
  plugins: [
    react()
  ],
  define: {
    global: 'globalThis',
  },
  resolve: {
    alias: {
      buffer: 'buffer',
      stream: resolve(cwd, 'src/react/shims/nodeStream.ts'),
      util: resolve(cwd, 'src/react/shims/nodeUtil.ts'),
    },
  },
  optimizeDeps: {
    include: ['buffer'],
  },
  server: {
    cors: {
      origin: true,
    },
    host: '127.0.0.1',
    port: 5739,
    // 扩展宿主 dev 模式硬编码从 5739 取 webview（reactApp.ts）：
    // 端口被占时必须立刻失败，静默换端口会让 webview 白屏/断言
    strictPort: true,
    fs: {
      allow: ['..'],
    },
  },
  base: '',
  build: {
    outDir: 'out/webview',
    chunkSizeWarningLimit: 2048,
  }
}))
