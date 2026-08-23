import { contextBridge, ipcRenderer } from 'electron';
import type { GenerationRequest, ProgressEvent } from '../shared/types';
import type { IPC } from './channels';

/**
 * A sandboxed preload script cannot `require` relative modules, so the channel names
 * are inlined here. The `satisfies`-style assignment below makes the compiler fail if
 * they ever drift from `src/main/channels.ts`.
 */
const CHANNELS = {
  pickSourceFile: 'shuffler:pick-source-file',
  inspect: 'shuffler:inspect',
  dryRun: 'shuffler:dry-run',
  generate: 'shuffler:generate',
  revealFolder: 'shuffler:reveal-folder',
  progress: 'shuffler:progress',
} as const;

const channelsMatchMainProcess: typeof IPC = CHANNELS;
void channelsMatchMainProcess;

/**
 * The only bridge between the UI and Node. The renderer has no direct filesystem or
 * module access; it can call exactly these five operations.
 */
contextBridge.exposeInMainWorld('shuffler', {
  pickSourceFile: () => ipcRenderer.invoke(CHANNELS.pickSourceFile),
  inspect: (sourceFile: string) => ipcRenderer.invoke(CHANNELS.inspect, sourceFile),
  dryRun: (request: GenerationRequest) => ipcRenderer.invoke(CHANNELS.dryRun, request),
  generate: (request: GenerationRequest) => ipcRenderer.invoke(CHANNELS.generate, request),
  revealFolder: (folder: string) => ipcRenderer.invoke(CHANNELS.revealFolder, folder),
  onProgress: (listener: (event: ProgressEvent) => void) => {
    ipcRenderer.on(CHANNELS.progress, (_event, payload: ProgressEvent) => listener(payload));
  },
});
