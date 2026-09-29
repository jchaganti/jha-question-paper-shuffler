// Loads the REAL compiled preloads the way a sandboxed Electron preload is loaded: `require`
// works for Electron's own modules and nothing else. This is the failure that broke the
// window - the preload threw at load, so `window.shuffler` was never defined.
//
// Both preloads are checked. `lock-preload.js` runs in front of the sign-in window, so if it
// throws the door has no handle: the app opens a page that cannot check a password and
// cannot be got past.
const path = require('node:path');
const Module = require('node:module');

const PRELOADS = [
  { file: 'dist/main/preload.js', global: 'shuffler' },
  { file: 'dist/main/lock-preload.js', global: 'lock' },
];

function checkPreload({ file, global: name }) {
  const listeners = {};
  const html = { dataset: {} };
  global.document = {
    addEventListener: (type, fn) => (listeners[type] = fn),
    documentElement: html,
  };

  const exposed = {};
  const electronStub = {
    contextBridge: { exposeInMainWorld: (key, api) => (exposed[key] = api) },
    ipcRenderer: {
      sendSync: (channel) => (channel === 'shuffler:read-theme' ? 'dim' : undefined),
      invoke: async () => true,
      send: () => {},
      on: () => {},
    },
  };

  const target = path.resolve(file);
  const realLoad = Module._load;
  Module._load = function (request, parent, ...rest) {
    if (request === 'electron') return electronStub;
    if (parent && parent.filename === target) {
      throw new Error(`sandboxed preload cannot require "${request}"`);
    }
    return realLoad.call(this, request, parent, ...rest);
  };

  try {
    delete require.cache[target];
    require(target);
    const api = exposed[name];
    if (!api) throw new Error(`nothing was exposed as window.${name}`);
    console.log(`${file}`);
    console.log(`  loaded        : ok`);
    console.log(`  bridge        : window.${name} { ${Object.keys(api).join(', ')} }`);
    console.log(`  theme from disk: ${api.theme}`);
    listeners['DOMContentLoaded']();
    console.log(`  data-theme set: ${html.dataset.theme}`);
    return true;
  } catch (error) {
    console.log(`${file}`);
    console.log(`  PRELOAD FAILED: ${error.message}`);
    return false;
  } finally {
    Module._load = realLoad;
  }
}

let allOk = true;
for (const preload of PRELOADS) {
  allOk = checkPreload(preload) && allOk;
}
if (!allOk) process.exitCode = 1;
