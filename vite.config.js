import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';

export default defineConfig({
  plugins: [react()],
  // PENTING: ganti '/headcount-dashboard/' sesuai nama repo GitHub lo persis.
  // Kalau nama repo lo "hc-dashboard", ini jadi base: '/hc-dashboard/'.
  base: '/replacementpegawai/',
});
