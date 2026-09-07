// Loads the REAL compiled preload the way a sandboxed Electron preload is loaded: `require`
// works for Electron's own modules and nothing else. This is the failure that broke the
// window - the preload threw at load, so `window.shuffler` was never defined.
const path = require('node:path');
const Module = require('node:module');

const listeners = {};
const html = { dataset: {} };
global.document = {
  addEventListener: (type, fn) => (listeners[type] = fn),
  documentElement: html,
};

let exposed;
const electronStub = {
  contextBridge: { exposeInMainWorld: (_name, api) => (exposed = api) },
  ipcRenderer: {
    sendSync: (channel) => (channel === 'shuffler:read-theme' ? 'dim' : undefined),
    invoke: async () => true,
    on: () => {},
  },
};

const target = path.resolve('dist/main/preload.js');
const realLoad = Module._load;
Module._load = function (request, parent, ...rest) {
  if (request === 'electron') return electronStub;
  if (parent && parent.filename === target) {
    throw new Error(`sandboxed preload cannot require "${request}"`);
  }
  return realLoad.call(this, request, parent, ...rest);
};

try {
  require(target);
  console.log('preload loaded  : ok');
  console.log('bridge keys     :', Object.keys(exposed).join(', '));
  console.log('theme from disk :', exposed.theme);
  listeners['DOMContentLoaded']();
  console.log('data-theme set  :', html.dataset.theme);
} catch (error) {
  console.log('PRELOAD FAILED  :', error.message);
  process.exitCode = 1;
}
