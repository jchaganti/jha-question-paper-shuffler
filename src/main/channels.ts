/** IPC channel names, shared by the main process and the preload bridge. */
export const IPC = {
  pickSourceFile: 'shuffler:pick-source-file',
  inspect: 'shuffler:inspect',
  dryRun: 'shuffler:dry-run',
  generate: 'shuffler:generate',
  revealFolder: 'shuffler:reveal-folder',
  progress: 'shuffler:progress',
  /**
   * Read synchronously by the preload script, so the appearance mode is on <html> before
   * the page paints. `invoke` would arrive a frame or two later, which is exactly long
   * enough to see the wrong colours.
   */
  readTheme: 'shuffler:read-theme',
  writeTheme: 'shuffler:write-theme',
  /**
   * The sign-in window. These are served to a window whose preload exposes nothing else:
   * until the right password is typed there is no window that can reach the channels above.
   */
  lockStatus: 'shuffler:lock-status',
  lockSet: 'shuffler:lock-set',
  lockUnlock: 'shuffler:lock-unlock',
  lockChange: 'shuffler:lock-change',
  lockReset: 'shuffler:lock-reset',
  lockProceed: 'shuffler:lock-proceed',
} as const;
