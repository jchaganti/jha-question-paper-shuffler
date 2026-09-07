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
import { THEME_PAGE_COLOUR, asTheme, type Theme } from '../shared/theme';
import { IPC } from './channels';
import { PreferencesStore } from './Preferences';

const service = new GenerationService();
let mainWindow: BrowserWindow | undefined;
let preferences: PreferencesStore | undefined;

/** The store, created on first use: `app.getPath` is only valid once Electron is ready. */
function store(): PreferencesStore {
  preferences ??= new PreferencesStore(app.getPath('userData'));
  return preferences;
}

function createWindow(): void {
  // Read before the window exists, so the very first frame is painted in the colour the
  // user chose last time. Applying it afterwards would show a flash of the wrong theme.
  const theme = store().read().theme;
  mainWindow = new BrowserWindow({
    width: 1040,
    height: 900,
    minWidth: 720,
    minHeight: 600,
    title: 'Question Paper Shuffler',
    backgroundColor: THEME_PAGE_COLOUR[theme],
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

// Synchronous on purpose: the preload script needs the answer before the page paints.
ipcMain.on(IPC.readTheme, (event) => {
  event.returnValue = store().read().theme;
});

ipcMain.handle(IPC.writeTheme, (_event, value: unknown): boolean => {
  const theme: Theme | undefined = asTheme(value);
  if (theme === undefined) return false;
  const saved = store().write({ theme });
  // The window keeps its own background colour for the life of the window; setting it here
  // too means a resize or a maximise never flashes the previous theme behind the page.
  if (saved) mainWindow?.setBackgroundColor(THEME_PAGE_COLOUR[theme]);
  return saved;
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
