import { BrowserWindow, app, dialog, ipcMain, shell } from 'electron';
import path from 'node:path';
import { GenerationService } from '../core/generate/GenerationService';
import type {
  DryRunReport,
  GenerationRequest,
  GenerationResult,
  PaperSummary,
  ProgressEvent,
  Result,
} from '../shared/types';
import { IPC } from './channels';

const service = new GenerationService();
let mainWindow: BrowserWindow | undefined;

function createWindow(): void {
  mainWindow = new BrowserWindow({
    width: 1040,
    height: 900,
    minWidth: 720,
    minHeight: 600,
    title: 'Question Paper Shuffler',
    backgroundColor: '#f5f6f8',
    autoHideMenuBar: true,
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
    },
  });

  void mainWindow.loadFile(path.join(__dirname, '..', 'web', 'renderer', 'index.html'));
}

/** Wraps a handler so the renderer always receives a Result instead of a thrown error. */
async function guard<T>(work: () => Promise<T>): Promise<Result<T>> {
  try {
    return { ok: true, value: await work() };
  } catch (error) {
    return { ok: false, error: error instanceof Error ? error.message : String(error) };
  }
}

ipcMain.handle(IPC.pickSourceFile, async (): Promise<string | null> => {
  const result = await dialog.showOpenDialog({
    title: 'Select the original question paper',
    filters: [{ name: 'Word document', extensions: ['docx'] }],
    properties: ['openFile'],
  });
  return result.canceled ? null : (result.filePaths[0] ?? null);
});

ipcMain.handle(IPC.inspect, (_event, sourceFile: string): Promise<Result<PaperSummary>> =>
  guard(() => service.inspect(sourceFile)),
);

ipcMain.handle(IPC.dryRun, (_event, request: GenerationRequest): Promise<Result<DryRunReport>> =>
  guard(() => service.dryRun(request)),
);

ipcMain.handle(IPC.generate, (event, request: GenerationRequest): Promise<Result<GenerationResult>> =>
  guard(() =>
    service.generate(request, (progress: ProgressEvent) => {
      if (!event.sender.isDestroyed()) event.sender.send(IPC.progress, progress);
    }),
  ),
);

ipcMain.handle(IPC.revealFolder, async (_event, folder: string): Promise<void> => {
  await shell.openPath(folder);
});

app.whenReady().then(
  () => {
    createWindow();
    app.on('activate', () => {
      if (BrowserWindow.getAllWindows().length === 0) createWindow();
    });
  },
  (error: unknown) => {
    dialog.showErrorBox('Startup failed', error instanceof Error ? error.message : String(error));
    app.quit();
  },
);

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit();
});
