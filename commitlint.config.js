/** @type {import('@commitlint/types').UserConfig} */
export default {
  extends: ['@commitlint/config-conventional'],
  rules: {
    // Type must be one of the allowed values
    'type-enum': [
      2,
      'always',
      [
        'feat', // New feature
        'fix', // Bug fix
        'docs', // Documentation only changes
        'style', // Changes that do not affect the meaning of the code
        'refactor', // Code change that neither fixes a bug nor adds a feature
        'perf', // Performance improvements
        'test', // Adding missing tests or correcting existing tests
        'build', // Changes that affect the build system or external dependencies
        'ci', // Changes to CI configuration files and scripts
        'chore', // Other changes that don't modify src or test files
        'revert', // Reverts a previous commit
        'security', // Security fixes
        'compliance', // Compliance-related changes (GDPR, ISO 27001)
      ],
    ],
    // Subject must not be empty
    'subject-empty': [2, 'never'],
    // Type must not be empty
    'type-empty': [2, 'never'],
    // Subject must be in lower case
    'subject-case': [2, 'always', 'lower-case'],
    // Header must not exceed 100 characters
    'header-max-length': [2, 'always', 100],
    // Body lines should not exceed 200 characters
    'body-max-line-length': [1, 'always', 200],
    // Scope should be in lower case
    'scope-case': [2, 'always', 'lower-case'],
  },
  helpUrl:
    'https://www.conventionalcommits.org/en/v1.0.0/#summary\n\nExamples:\n  feat: add user authentication\n  fix(auth): resolve token refresh issue\n  docs: update API documentation\n  chore(deps): update dependencies',
};
