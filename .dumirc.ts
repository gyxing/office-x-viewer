import { defineConfig } from 'dumi';

export default defineConfig({
  outputPath: 'docs-dist',
  themeConfig: {
    name: 'office-x-viewer',
  },
  resolve: {
    docDirs: ['./docs-2'],
  },
});
