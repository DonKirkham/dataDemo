// ABOUTME: Thin wrapper around @pnp/logging that exposes info/debug/warn/error (with log as
// ABOUTME: alias for info). Emits via console so object payloads render as collapsible trees.

import { Logger as PnPLogger, LogLevel, ConsoleListener } from '@pnp/logging';

const PREFIX = '[DataDemo]';

function shouldEmit(level: LogLevel): boolean {
  return PnPLogger.activeLogLevel <= level;
}

function emit(consoleFn: (...args: unknown[]) => void, message: string, data: unknown[]): void {
  const text = `${PREFIX} ${message}`;
  if (data.length === 0) consoleFn(text);
  else consoleFn(text, ...data);
}

export const Logger = {
  info(message: string, ...data: unknown[]): void {
    if (!shouldEmit(LogLevel.Info)) return;
    // eslint-disable-next-line no-console
    emit(console.info.bind(console), message, data);
  },
  debug(message: string, ...data: unknown[]): void {
    if (!shouldEmit(LogLevel.Verbose)) return;
    // eslint-disable-next-line no-console
    emit(console.debug.bind(console), message, data);
  },
  warn(message: string, ...data: unknown[]): void {
    if (!shouldEmit(LogLevel.Warning)) return;
    // eslint-disable-next-line no-console
    emit(console.warn.bind(console), message, data);
  },
  error(message: string, ...data: unknown[]): void {
    if (!shouldEmit(LogLevel.Error)) return;
    // eslint-disable-next-line no-console
    emit(console.error.bind(console), message, data);
  },
  log(message: string, ...data: unknown[]): void {
    Logger.info(message, ...data);
  },
  setLevel(level: LogLevel): void {
    PnPLogger.activeLogLevel = level;
  },
  attachConsoleListener(): void {
    PnPLogger.subscribe(ConsoleListener());
  }
};

export { LogLevel };
