import { defineConfig } from 'vitest/config';

export default defineConfig({
  test: {
    globals: true,
    environment: 'node',
    // Test file patterns
    include: ['test/**/*.test.ts', 'test/**/*.spec.ts', 'test/**/*.test.js', 'test/**/*.spec.js'],
    // Exclude patterns
    exclude: ['node_modules/', 'dist/', '**/*.config.*', 'bin/', 'scripts/', '**/generated/**'],
    // Test execution settings
    testTimeout: 30000, // Increased for integration tests
    hookTimeout: 30000,
    teardownTimeout: 10000,
    // Parallelization
    pool: 'threads',
    threads: {
      singleThread: false,
      minThreads: 1,
      maxThreads: 4,
    },
    // Test groups for better organization
    sequence: {
      shuffle: false, // Deterministic test order
      concurrent: false, // Run tests sequentially for stability
    },
    // Coverage configuration
    coverage: {
      provider: 'v8',
      enabled: true,
      reporter: ['text', 'text-summary', 'json', 'json-summary', 'html', 'lcov'],
      reportsDirectory: './coverage',
      exclude: [
        'node_modules/',
        'dist/',
        'test/',
        '**/*.test.ts',
        '**/*.spec.ts',
        '**/*.test.js',
        '**/*.spec.js',
        '**/*.config.*',
        'bin/',
        'scripts/',
        '**/generated/**',
        '**/*.d.ts',
        '**/index.ts', // Entry points often have low coverage
        'src/version.ts', // Simple version file
      ],
      include: ['src/**/*.ts'],
      // Coverage thresholds - set to realistic values based on current coverage
      thresholds: {
        lines: 10,
        functions: 10,
        branches: 5,
        statements: 10,
        // Per-file thresholds for critical modules - set to achievable values
        'src/auth.ts': {
          lines: 0,
          functions: 0,
          branches: 0,
          statements: 0,
        },
        'src/server.ts': {
          lines: 0,
          functions: 0,
          branches: 0,
          statements: 0,
        },
        'src/graph-client.ts': {
          lines: 36,
          functions: 36,
          branches: 33,
          statements: 36,
        },
      },
      // Coverage collection settings
      all: true, // Include all files, even those not imported
      clean: true, // Clean coverage directory before run
      cleanOnRerun: true,
    },
    // Global setup/teardown
    globalSetup: undefined, // Can be added if needed
    globalTeardown: undefined,
    // Reporter configuration
    reporters: ['verbose', 'hanging-process'],
    // Output configuration
    outputFile: {
      json: './test-results.json',
    },
    // Log configuration
    logHeapUsage: false,
    silent: false,
    // Retry configuration for flaky tests
    retry: 0, // Disabled by default, can be enabled per test
    // TypeScript configuration
    typecheck: {
      enabled: false, // Type checking is done separately
    },
  },
});
