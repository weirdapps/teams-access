import { defineConfig } from 'vitest/config';

export default defineConfig({
  test: {
    include: ['test_scripts/**/*.test.ts'],
    environment: 'node',
    globals: false,
    reporters: ['default'],
    coverage: {
      provider: 'v8',
      reporter: ['text', 'lcov'],
      include: ['src/**/*.ts'],
      exclude: ['src/cli.ts', 'src/auth/browser-capture.ts'],
    },
  },
});
