// ABOUTME: Thin wrapper around @pnp/logging that exposes info/debug/warn/error (with log as
// ABOUTME: alias for info). Pre-binds console.* so DevTools attributes each line to the caller,
// ABOUTME: not this file. setLevel() rebinds the methods to either bound console.* or a no-op.

import { Logger as PnPLogger, LogLevel, ConsoleListener } from '@pnp/logging';

const PREFIX = '[DataDemo]';

type LogFn = (message: string, ...data: unknown[]) => void;

const noop: LogFn = () => { /* below threshold */ };

function bindIfActive(consoleFn: (...args: unknown[]) => void, level: LogLevel): LogFn {
  // A bound function is native; V8/Chromium skip its frame in source attribution,
  // so DevTools shows the caller of Logger.<level>(), not this file.
  return PnPLogger.activeLogLevel <= level
    ? (consoleFn.bind(console, PREFIX) as LogFn)
    : noop;
}

function rebind(): void {
  // Logger is declared below; rebind() only runs after that initializer, so the
  // forward reference is safe. Silence the rule for this intentional pattern.
  /* eslint-disable @typescript-eslint/no-use-before-define */
  Logger.info  = bindIfActive(console.info,  LogLevel.Info);
  Logger.debug = bindIfActive(console.debug, LogLevel.Verbose);
  Logger.warn  = bindIfActive(console.warn,  LogLevel.Warning);
  Logger.error = bindIfActive(console.error, LogLevel.Error);
  Logger.log   = Logger.info;
  /* eslint-enable @typescript-eslint/no-use-before-define */
}

interface ILogger {
  info: LogFn;
  debug: LogFn;
  warn: LogFn;
  error: LogFn;
  log: LogFn;
  setLevel(level: LogLevel): void;
  attachConsoleListener(): void;
}

export const Logger: ILogger = {
  info: noop,
  debug: noop,
  warn: noop,
  error: noop,
  log: noop,
  setLevel(level: LogLevel): void {
    PnPLogger.activeLogLevel = level;
    rebind();
  },
  attachConsoleListener(): void {
    PnPLogger.subscribe(ConsoleListener());
  }
};

rebind();

export { LogLevel };
