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
import type { LockStatus, SetOutcome, UnlockOutcome } from '../shared/lock';
import { THEME_PAGE_COLOUR, asTheme, type Theme } from '../shared/theme';
import { AppLock } from './AppLock';
import { IPC } from './channels';
import { electronVault, windowsMachineId } from './LockRuntime';
import { PreferencesStore } from './Preferences';

const service = new GenerationService();
let mainWindow: BrowserWindow | undefined;
let lockWindow: BrowserWindow | undefined;
let preferences: PreferencesStore | undefined;
let lock: AppLock | undefined;
/**
 * Whether the password has been accepted in this run.
 *
 * The sign-in window closing means two different things - "let me in" and "I have changed
 * my mind" - and this is what tells them apart. It is also the guard on every lock channel
 * below: once it is true there is no sign-in window left to call them.
 */
let unlocked = false;

/** The store, created on first use: `app.getPath` is only valid once Electron is ready. */
function store(): PreferencesStore {
  preferences ??= new PreferencesStore(app.getPath('userData'));
  return preferences;
}

/** The same, for the password this computer asks for when the app starts. */
function appLock(): AppLock {
  lock ??= new AppLock(app.getPath('userData'), electronVault, windowsMachineId);
  return lock;
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

/**
 * The sign-in window, shown before the app itself exists.
 *
 * Small, fixed and not minimisable: it is a door, not a window of the app. Its preload is
 * `lock-preload.js`, which exposes the sign-in operations and nothing else, so the page
 * cannot reach a question paper even if something were wrong with the page.
 */
function createLockWindow(): void {
  const theme = store().read().theme;
  lockWindow = new BrowserWindow({
    width: 480,
    height: 560,
    resizable: false,
    maximizable: false,
    fullscreenable: false,
    title: 'Question Paper Shuffler',
    backgroundColor: THEME_PAGE_COLOUR[theme],
    autoHideMenuBar: true,
    webPreferences: {
      preload: path.join(__dirname, 'lock-preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true,
    },
  });

  // Closing the door without opening it is the same as not starting the app. Once the
  // password has been accepted this window is closed by us, and `unlocked` says so.
  lockWindow.on('closed', () => {
    lockWindow = undefined;
    if (!unlocked) app.quit();
  });

  void lockWindow.loadFile(path.join(__dirname, '..', 'web', 'renderer', 'lock.html'));
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

/**
 * The sign-in channels.
 *
 * Every one of them refuses once the app is open. A renderer that is already past the door
 * has no business setting or checking the password, and the only window that can reach
 * these channels at all is the one that is closed the moment it succeeds.
 */
function whileLocked<T>(answer: () => T, refusal: T): T {
  return unlocked ? refusal : answer();
}

ipcMain.handle(IPC.lockStatus, (): LockStatus => appLock().status());

ipcMain.handle(IPC.lockSet, (_event, password: unknown): SetOutcome =>
  whileLocked(
    () => appLock().setPassword(asPassword(password)),
    { ok: false, message: 'The app is already open.' },
  ),
);

ipcMain.handle(IPC.lockUnlock, (_event, password: unknown): UnlockOutcome =>
  whileLocked(
    () => appLock().unlock(asPassword(password)),
    { ok: false, reason: 'unreadable', message: 'The app is already open.' },
  ),
);

ipcMain.handle(IPC.lockChange, (_event, current: unknown, next: unknown): SetOutcome =>
  whileLocked(
    () => appLock().changePassword(asPassword(current), asPassword(next)),
    { ok: false, message: 'The app is already open.' },
  ),
);

ipcMain.handle(IPC.lockReset, (_event, adminPassword: unknown, next: unknown): SetOutcome =>
  whileLocked(
    () => appLock().resetWithAdminPassword(asPassword(adminPassword), asPassword(next)),
    { ok: false, message: 'The app is already open.' },
  ),
);

/**
 * Opens the app. Only reachable from the sign-in window, and only once - a second call
 * arrives after `unlocked` is set and does nothing.
 */
ipcMain.on(IPC.lockProceed, () => {
  if (unlocked) return;
  unlocked = true;
  createWindow();
  lockWindow?.close();
});

/** Whatever came over IPC, as the string the lock expects. */
function asPassword(value: unknown): string {
  return typeof value === 'string' ? value : '';
}

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
    // The password comes first: the app's own window is not created until the sign-in
    // window says so, so there is never a moment where the UI exists but the door is shut.
    createLockWindow();
    app.on('activate', () => {
      if (BrowserWindow.getAllWindows().length > 0) return;
      if (unlocked) createWindow();
      else createLockWindow();
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
