/** IPC channel names, shared by the main process and the preload bridge. */
export const IPC = {
  pickSourceFile: 'shuffler:pick-source-file',
  inspect: 'shuffler:inspect',
  dryRun: 'shuffler:dry-run',
  generate: 'shuffler:generate',
  revealFolder: 'shuffler:reveal-folder',
  progress: 'shuffler:progress',
} as const;
