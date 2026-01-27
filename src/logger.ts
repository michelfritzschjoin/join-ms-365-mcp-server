import winston from 'winston';
import path from 'path';
import { fileURLToPath } from 'url';
import fs from 'fs';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const logsDir = path.join(__dirname, '..', 'logs');

if (!fs.existsSync(logsDir)) {
  fs.mkdirSync(logsDir);
}

// Determine log format
const logFormat = process.env.LOG_FORMAT || 'text';
const useJsonFormat = logFormat.toLowerCase() === 'json';

const logger = winston.createLogger({
  level: process.env.LOG_LEVEL || 'info',
  format: winston.format.combine(
    winston.format.timestamp({
      format: 'YYYY-MM-DD HH:mm:ss',
    }),
    winston.format.errors({ stack: true }),
    useJsonFormat
      ? winston.format.json()
      : winston.format.printf(({ level, message, timestamp, ...meta }) => {
          const metaStr = Object.keys(meta).length > 0 ? ` ${JSON.stringify(meta)}` : '';
          return `${timestamp} ${level.toUpperCase()}: ${message}${metaStr}`;
        })
  ),
  defaultMeta: {
    service: 'ms-365-mcp-server',
  },
  transports: [
    new winston.transports.File({
      filename: path.join(logsDir, 'error.log'),
      level: 'error',
    }),
    new winston.transports.File({
      filename: path.join(logsDir, 'mcp-server.log'),
    }),
  ],
});

// Enable console logging by default unless SILENT is set
// This ensures LOG_LEVEL=debug works without requiring -v flag
const isSilent = process.env.SILENT === 'true' || process.env.SILENT === '1';
if (!isSilent) {
  logger.add(
    new winston.transports.Console({
      // CRITICAL: MCP STDIO servers MUST write to stderr, not stdout
      // Writing to stdout corrupts JSON-RPC messages
      stderrLevels: ['error', 'warn', 'info', 'http', 'verbose', 'debug', 'silly'],
      format: winston.format.combine(winston.format.colorize(), winston.format.simple()),
    })
  );
}

export const enableConsoleLogging = (): void => {
  // Check if console transport already exists
  const hasConsoleTransport = logger.transports.some(
    (transport) => transport instanceof winston.transports.Console
  );

  if (!hasConsoleTransport) {
    logger.add(
      new winston.transports.Console({
        // CRITICAL: MCP STDIO servers MUST write to stderr, not stdout
        // Writing to stdout corrupts JSON-RPC messages
        stderrLevels: ['error', 'warn', 'info', 'http', 'verbose', 'debug', 'silly'],
        format: winston.format.combine(winston.format.colorize(), winston.format.simple()),
        silent: process.env.SILENT === 'true' || process.env.SILENT === '1',
      })
    );
  }
};

export default logger;
