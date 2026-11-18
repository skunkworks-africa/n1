import { defineConfig } from 'vite';
import {resolve} from 'node:path';
export default defineConfig({base:'./', build:{sourcemap:false,rollupOptions:{input:{main:resolve('index.html'),redirect:resolve('redirect.html')}}}});
