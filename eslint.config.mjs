// ESLint flat config for TypeScript/JavaScript projects
// Requires: npm install --save-dev eslint @typescript-eslint/parser @typescript-eslint/eslint-plugin eslint-config-prettier

import js from '@eslint/js';
import tseslint from '@typescript-eslint/eslint-plugin';
import tsparser from '@typescript-eslint/parser';
import prettier from 'eslint-config-prettier';

// Node.js built-in globals (subset covering this project's usage)
const nodeGlobals = {
  process: 'readonly',
  Buffer: 'readonly',
  __dirname: 'readonly',
  __filename: 'readonly',
  module: 'readonly',
  require: 'readonly',
  exports: 'readonly',
  setTimeout: 'readonly',
  clearTimeout: 'readonly',
  setInterval: 'readonly',
  clearInterval: 'readonly',
  setImmediate: 'readonly',
  clearImmediate: 'readonly',
  console: 'readonly',
  URL: 'readonly',
  URLSearchParams: 'readonly',
  global: 'readonly',
};

// Fetch API globals available in Node 18+ and test environments
const fetchGlobals = {
  fetch: 'readonly',
  Headers: 'readonly',
  Request: 'readonly',
  Response: 'readonly',
};

// Vitest globals
const vitestGlobals = {
  describe: 'readonly',
  it: 'readonly',
  test: 'readonly',
  expect: 'readonly',
  beforeEach: 'readonly',
  afterEach: 'readonly',
  beforeAll: 'readonly',
  afterAll: 'readonly',
  vi: 'readonly',
};

export default [
  // Recommended base
  js.configs.recommended,

  // TypeScript files
  {
    files: ['**/*.ts', '**/*.tsx'],
    languageOptions: {
      parser: tsparser,
      parserOptions: {
        ecmaVersion: 2022,
        sourceType: 'module',
      },
      globals: {
        ...nodeGlobals,
        ...fetchGlobals,
      },
    },
    plugins: {
      '@typescript-eslint': tseslint,
    },
    rules: {
      ...tseslint.configs.recommended.rules,
      // TypeScript's compiler handles undefined-reference checks for .ts files;
      // no-undef produces false positives on TS types (NodeJS.Timeout, RequestInit, etc.)
      'no-undef': 'off',
      '@typescript-eslint/no-unused-vars': [
        'error',
        {
          argsIgnorePattern: '^_',
          varsIgnorePattern: '^_',
        },
      ],
      '@typescript-eslint/no-explicit-any': 'warn',
      '@typescript-eslint/explicit-function-return-type': 'off',
      '@typescript-eslint/explicit-module-boundary-types': 'off',
    },
  },

  // JavaScript files
  {
    files: ['**/*.js', '**/*.jsx'],
    languageOptions: {
      ecmaVersion: 2022,
      sourceType: 'module',
      globals: {
        ...nodeGlobals,
      },
    },
    rules: {
      'no-unused-vars': [
        'error',
        {
          argsIgnorePattern: '^_',
          varsIgnorePattern: '^_',
        },
      ],
    },
  },

  // Test files - relax some rules and add test runner globals
  {
    files: ['**/*.test.ts', '**/*.test.js', '**/*.spec.ts', '**/*.spec.js', 'test_scripts/**/*'],
    languageOptions: {
      globals: {
        ...vitestGlobals,
      },
    },
    rules: {
      '@typescript-eslint/no-explicit-any': 'off',
      'no-console': 'off',
    },
  },

  // Ignore patterns
  {
    ignores: [
      'node_modules/**',
      'dist/**',
      'build/**',
      'coverage/**',
      '.venv/**',
      '*.min.js',
      '.env*',
      '!.env.example',
    ],
  },

  // Prettier - must be last to override formatting rules
  prettier,
];
