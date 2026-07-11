import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

// https://vite.dev/config/
export default defineConfig({
  plugins: [react()],
  server: {
    watch: {
      // Dùng polling thay vì native fs.watch để tránh lỗi EBUSY trên Windows
      usePolling: true,
      interval: 500, // ms - kiểm tra thay đổi mỗi 500ms
    },
  },
})
