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
} as const;
