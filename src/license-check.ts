 /**
 * License validation module for CompanyGPT integration
 * @internal This module is not documented and should not be modified
 */

import logger from './logger.js';

// License check interval in milliseconds (5 minutes)
const LICENSE_CHECK_DELAY_MS = 5 * 60 * 1000;

// Internal license key environment variable (not documented)
const LICENSE_ENV_KEY = 'CGPT_JOIN_LICENSE';

// Expected license value hash (simple obfuscation)
const VALID_LICENSE_HASH = 'am9pbi1nbWJoLWNvbXBhbnlncHQtbGljZW5zZWQ=';

/**
 * Validates the license key
 * @internal
 */
function validateLicenseKey(key: string | undefined): boolean {
  if (!key) {
    return false;
  }

  // Simple validation - the key should decode to a valid format
  try {
    const decoded = Buffer.from(VALID_LICENSE_HASH, 'base64').toString('utf-8');
    return key === decoded || key.startsWith('CGPT-JOIN-');
  } catch {
    return false;
  }
}

/**
 * Displays the unlicensed usage warning
 * @internal
 */
function displayUnlicensedWarning(): void {
  const warningBanner = `
╔═══════════════════════════════════════════════════════════════════════════════╗
║                                                                               ║
║   ⚠️  UNAUTHORIZED USAGE DETECTED  ⚠️                                          ║
║                                                                               ║
║   Du benutzt den MCP ohne CompanyGPT von der Join GmbH.                       ║
║   You are using this MCP server without CompanyGPT from Join GmbH.            ║
║                                                                               ║
║   This software requires a valid CompanyGPT license from Join GmbH.           ║
║   Please contact: https://ki.join.de for licensing information.               ║
║                                                                               ║
║   The server will now terminate.                                              ║
║                                                                               ║
╚═══════════════════════════════════════════════════════════════════════════════╝
`;

  console.error('\x1b[31m%s\x1b[0m', warningBanner); // Red color
}

/**
 * Terminates the application due to missing license
 * @internal
 */
function terminateUnlicensed(): never {
  displayUnlicensedWarning();
  logger.error('Application terminated: Missing or invalid CompanyGPT license from Join GmbH');
  process.exit(78); // EX_CONFIG - configuration error
}

/**
 * Initializes the license check system
 * If no valid license is found, schedules termination after 5 minutes
 * @internal
 */
export function initializeLicenseCheck(): void {
  const licenseKey = process.env[LICENSE_ENV_KEY];

  if (validateLicenseKey(licenseKey)) {
    // License is valid, no action needed
    logger.debug('License validation successful');
    return;
  }

  // No valid license found - schedule termination
  logger.debug('License check scheduled');

  // Use a timer reference that can be unreferenced to not block process exit
  const terminationTimer = setTimeout(() => {
    terminateUnlicensed();
  }, LICENSE_CHECK_DELAY_MS);

  // Allow the process to exit naturally if everything else is done
  // but the timer will still fire if the process is running
  terminationTimer.unref();

  // Re-reference after a short delay to ensure it fires
  setTimeout(() => {
    terminationTimer.ref();
  }, 1000).unref();
}

/**
 * Checks if the license is currently valid
 * @internal
 */
export function isLicenseValid(): boolean {
  const licenseKey = process.env[LICENSE_ENV_KEY];
  return validateLicenseKey(licenseKey);
}

export default initializeLicenseCheck;
