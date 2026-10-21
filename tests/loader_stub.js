// Minimal DOM and virtual clock for running office-file-ui/loader.js in Node.
// Usage: node loader_stub.js <scenario>   -> prints a JSON report of what the script did.
// Scenarios: missing-file, found-file, tabs, tabs-no-switch, tabs-reordered

const fs = require('fs');
const path = require('path');
const vm = require('vm');

// ------------------------------------------------------------------------------------- clock --
let now = 0;
let nextTimer = 1;
const timers = new Map();
globalThis.setTimeout = (fn, ms = 0) => {
  const id = nextTimer++;
  timers.set(id, { at: now + Math.max(0, ms || 0), fn });
  return id;
};
globalThis.clearTimeout = (id) => timers.delete(id);
globalThis.requestAnimationFrame = (fn) => setTimeout(() => fn(now), 16);
globalThis.performance = { now: () => now };
Date.now = () => 1.7e12 + now;
const flush = async () => {
  for (let i = 0; i < 20; i++) await Promise.resolve();
};
async function advance(ms) {
  const end = now + ms;
  for (;;) {
    await flush();
    let due = null;
    for (const [id, timer] of timers) if (timer.at <= end && (!due || timer.at < due[1].at)) due = [id, timer];
    if (!due) break;
    timers.delete(due[0]);
    now = Math.max(now, due[1].at);
    due[1].fn();
  }
  now = end;
  await flush();
}

// --------------------------------------------------------------------------------------- DOM --
const report = { fetches: 0, scans: 0, clicks: 0 };

function parseCompound(text) {
  const tag = text.match(/^[a-z]+/i)?.[0];
  const classes = [...text.matchAll(/\.([\w-]+)/g)].map((m) => m[1]);
  const attrs = [...text.matchAll(/\[([\w-]+)(?:(\^?=)"([^"]*)")?\]/g)].map((m) => ({ name: m[1], op: m[2], value: m[3] }));
  return { tag, classes, attrs };
}

function matchesCompound(el, { tag, classes, attrs }) {
  if (!(el instanceof Element)) return false;
  if (tag && el.tagName !== tag.toLowerCase()) return false;
  if (!classes.every((c) => el.classList.contains(c))) return false;
  return attrs.every(({ name, op, value }) => {
    const actual = el.getAttribute(name);
    if (actual === null) return false;
    return !op || (op === '=' ? actual === value : actual.startsWith(value));
  });
}

// Selectors with descendant (space) and child (>) combinators, matched from right to left.
function matches(el, selector) {
  const tokens = selector.trim().split(/\s*(>)\s*|\s+/).filter(Boolean);
  const match = (node, index) => {
    if (!matchesCompound(node, parseCompound(tokens[index]))) return false;
    if (index === 0) return true;
    if (tokens[index - 1] === '>') return node.parentElement ? match(node.parentElement, index - 2) : false;
    for (let up = node.parentElement; up; up = up.parentElement) if (match(up, index - 1)) return true;
    return false;
  };
  return match(el, tokens.length - 1);
}

class Element {
  constructor(tag, { cls = '', attrs = {}, text = '' } = {}) {
    this.tagName = tag.toLowerCase();
    this.childList = [];
    this.parentElement = null;
    this.attrs = new Map(Object.entries(attrs));
    this.classes = new Set(cls.split(/\s+/).filter(Boolean));
    this.style = {};
    this.dataset = {};
    this.text = text;
    this.listeners = {};
    const classes = this.classes;
    this.classList = {
      contains: (c) => classes.has(c),
      add: (...cs) => cs.forEach((c) => classes.add(c)),
      remove: (...cs) => cs.forEach((c) => classes.delete(c)),
    };
  }
  get children() {
    return this.childList;
  }
  get id() {
    return this.attrs.get('id') ?? '';
  }
  set className(value) {
    this.classes.clear();
    value.split(/\s+/).filter(Boolean).forEach((c) => this.classes.add(c));
  }
  set innerHTML(value) {
    this.html = value;
  }
  get textContent() {
    return this.text + this.childList.map((c) => c.textContent).join('');
  }
  set textContent(value) {
    this.text = value;
  }
  get nextSibling() {
    return null;
  }
  getAttribute(name) {
    if (name === 'class') return [...this.classes].join(' ');
    return this.attrs.has(name) ? this.attrs.get(name) : null;
  }
  setAttribute(name, value) {
    this.attrs.set(name, String(value));
  }
  hasAttribute(name) {
    return this.attrs.has(name);
  }
  set type(value) {
    this.setAttribute('type', value);
  }
  appendChild(child) {
    child.remove();
    child.parentElement = this;
    this.childList.push(child);
    return child;
  }
  append(...children) {
    children.forEach((c) => this.appendChild(c));
  }
  prepend(child) {
    child.remove();
    child.parentElement = this;
    this.childList.unshift(child);
  }
  remove() {
    if (!this.parentElement) return;
    const siblings = this.parentElement.childList;
    siblings.splice(siblings.indexOf(this), 1);
    this.parentElement = null;
  }
  descendants() {
    return this.childList.flatMap((c) => [c, ...c.descendants()]);
  }
  querySelectorAll(selector) {
    if (this === document.body && selector === '.chat-assistant') report.scans += 1;
    return this.descendants().filter((el) => matches(el, selector));
  }
  querySelector(selector) {
    return this.querySelectorAll(selector)[0] ?? null;
  }
  matches(selector) {
    return matches(this, selector);
  }
  closest(selector) {
    for (let el = this; el; el = el.parentElement) if (el instanceof Element && matches(el, selector)) return el;
    return null;
  }
  addEventListener(type, fn) {
    (this.listeners[type] ??= []).push(fn);
  }
  click() {
    for (const fn of this.listeners.click ?? []) fn({ target: this, preventDefault() {}, stopPropagation() {} });
  }
}

const el = (tag, options, ...children) => {
  const element = new Element(tag, options);
  children.forEach((c) => element.appendChild(c));
  return element;
};

const html = el('html');
const body = el('body');
html.appendChild(body);
const documentListeners = {};
globalThis.document = {
  body,
  head: el('head'),
  createElement: (tag) => new Element(tag),
  querySelectorAll: (selector) => body.querySelectorAll(selector),
  querySelector: (selector) => body.querySelector(selector),
  addEventListener: (type, fn) => (documentListeners[type] ??= []).push(fn),
};
globalThis.Node = { TEXT_NODE: 3 };
globalThis.location = { pathname: '/c/chat-1' };
globalThis.localStorage = { token: 'token', locale: 'en-US', getItem: () => null };
globalThis.navigator = { language: 'en-US' };
let observer = null;
globalThis.MutationObserver = class {
  constructor(callback) {
    observer = callback;
  }
  observe() {}
};
let chatFiles = [];
let failingFetches = 0; // the first requests fail (network blip, 5xx)
globalThis.fetch = async (url) => {
  report.fetches += 1;
  if (report.fetches <= failingFetches) return { ok: false, json: async () => ({}) };
  return { ok: true, json: async () => ({ chat: { history: { messages: { m1: { files: chatFiles } } } } }) };
};
const mutate = () => observer && observer([]);

// ---------------------------------------------------------------------------------- scenarios --
const chip = el(
  'button',
  { cls: 'h-8 rounded-xl', attrs: { type: 'button' } },
  el('div', { cls: 'icon' }),
  el('div', { cls: 'justify-between' }, el('div', { cls: 'truncate', text: 'report.docx' }))
);
body.appendChild(
  el('div', { attrs: { id: 'message-m1' } }, el('div', { cls: 'chat-assistant' }, el('div', {}, el('div', { cls: 'flex-wrap' }, el('div', {}, chip)))))
);

function tabBar(selected, onClick) {
  const tab = (index) => {
    const button = el('button', { cls: `border-b${index === selected ? '' : ' border-transparent'}`, attrs: { type: 'button' } });
    button.addEventListener('click', () => {
      report.clicks += 1;
      onClick(index);
    });
    return button;
  };
  const bar = el('div', { cls: 'flex border-b' }, tab(0), tab(1));
  const select = (index) =>
    queueMicrotask(() => bar.children.forEach((b, i) => (i === index ? b.classList.remove('border-transparent') : b.classList.add('border-transparent'))));
  return { bar, select };
}

async function main(scenario) {
  vm.runInThisContext(fs.readFileSync(path.join(__dirname, '..', 'office-file-ui', 'loader.js'), 'utf8'));
  if (scenario === 'missing-file' || scenario === 'found-file' || scenario === 'flaky-server') {
    if (scenario === 'flaky-server') failingFetches = 3;
    if (scenario !== 'missing-file') chatFiles = [{ id: 'f1', name: 'report.docx', type: 'file' }];
    for (let t = 0; t < 120000; t += 100) {
      mutate(); // a streaming answer changes the page all the time
      await advance(100);
    }
    const icon = chip.querySelector('[data-office-file-ui="download"]');
    report.icon = icon ? icon.dataset.fileId : null;
    report.seconds = now / 1000;
  } else {
    let bar;
    let select;
    const selected = scenario === 'tabs-reordered' ? 1 : 0;
    ({ bar, select } = tabBar(selected, (index) => scenario !== 'tabs-no-switch' && select(index)));
    for (const fn of documentListeners.click ?? []) fn({ target: chip });
    const modal = el(
      'div',
      { cls: 'modal', attrs: { role: 'dialog' } },
      el('a', { cls: 'line-clamp-1', text: 'report.docx' }),
      bar
    );
    body.appendChild(modal);
    mutate();
    await advance(1000);
    report.barDisplay = bar.style.display ?? '';
    report.selected = bar.children.findIndex((b) => !b.classList.contains('border-transparent'));
    // Open WebUI switches back to "Content" (e.g. it reloads the file): the bar has to come back
    if (scenario === 'tabs') {
      select(0);
      mutate();
      await advance(1000);
      report.barDisplayAfterReset = bar.style.display ?? '';
    }
  }
  console.log(JSON.stringify(report));
}

main(process.argv[2]).catch((error) => {
  console.error(error);
  process.exit(1);
});
