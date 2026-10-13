// Minimal browser stand-in for running the JavaScript snippets of response_export.py in Node.
// Usage: node browser_stub.js < snippet.js   -> prints a JSON report of what the snippet did.
// The snippet is wrapped exactly like Open WebUI's "execute" event handler does it.

const report = { returned: null, clicks: [], blobs: [], printed: false, frameRemoved: false, srcdoc: null };

function element(tag) {
  const el = {
    tag,
    style: {},
    click() {
      report.clicks.push({ href: el.href, download: el.download });
    },
    remove() {
      if (tag === 'iframe') report.frameRemoved = true;
    },
  };
  if (tag === 'iframe') {
    el.contentWindow = { focus() {}, print() { report.printed = true; } };
    Object.defineProperty(el, 'srcdoc', {
      set(value) { report.srcdoc = value; setTimeout(() => el.onload && el.onload(), 0); },
    });
  }
  return el;
}

globalThis.document = { createElement: element, body: { appendChild() {} } };
const createObjectURL = URL.createObjectURL;
URL.createObjectURL = (blob) => {
  report.blobs.push({ size: blob.size, type: blob.type });
  return createObjectURL(blob);
};
// Shorten timers so the 60 s frame cleanup does not keep the test waiting.
const realSetTimeout = setTimeout;
globalThis.setTimeout = (fn, ms) => realSetTimeout(fn, Math.min(ms, 50));

let code = '';
process.stdin.on('data', (chunk) => (code += chunk));
process.stdin.on('end', async () => {
  report.returned = await new Function('return (async () => { ' + code + ' })()')();
  realSetTimeout(() => console.log(JSON.stringify(report)), 300);
});
