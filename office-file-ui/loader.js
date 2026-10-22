// Office File UI for Open WebUI (tested with 0.11.4, 0.12.0 and 0.12.1)
//
// Open WebUI loads /static/loader.js on every page. For files attached to an answer of the assistant
// (e.g. by the Office Documents tool) this script
// - shows them below the text of the answer instead of above it,
// - fits each file chip to its name, colours its icon by type (Word, Excel, PowerPoint) and adds a
//   download icon to it: the icon downloads the file, the rest of the chip opens the preview, and
// - opens Word, Excel and PowerPoint files straight in the "Preview" tab of the file view and hides
//   the "Content" tab (the extracted text).
// Files the user uploads and file views elsewhere (workspace, knowledge, chat controls) stay as they are.
// It only changes the page in the browser and depends on Open WebUI's markup: check it after updates.
// If the markup changes, the icons and the tab switch simply stop appearing; nothing else breaks.

(() => {
	'use strict';

	const PREVIEW_ONLY = /\.(docx|xlsx|pptx)$/i;
	const MARK = 'data-office-file-ui';
	const CHIP = 'button.h-8.rounded-xl'; // FileItem with small=true
	const RUN_INTERVAL_MS = 150; // page changes are handled at most this often (streaming changes it constantly)
	const RELOAD_DELAYS_MS = [2000, 8000, 32000]; // reloads of the chat for a chip without a known file
	const FAILED_LOAD_DELAY_MS = 30000; // after a failed reload, which does not use up one of them
	const GERMAN = (localStorage.locale || navigator.language || '').toLowerCase().startsWith('de');
	const LABEL = GERMAN ? 'Herunterladen' : 'Download';
	const FAILED = GERMAN ? 'Herunterladen fehlgeschlagen' : 'Download failed';
	const downloadIcon = (size) =>
		`<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 20 20" fill="currentColor" class="${size}" aria-hidden="true">` +
		'<path d="M10.75 2.75a.75.75 0 0 0-1.5 0v8.614L6.295 8.235a.75.75 0 1 0-1.09 1.03l4.25 4.5a.75.75 0 0 0 1.09 0l4.25-4.5a.75.75 0 0 0-1.09-1.03l-2.955 3.129V2.75Z"/>' +
		'<path d="M3.5 12.75a.75.75 0 0 0-1.5 0v2.5A2.75 2.75 0 0 0 4.75 18h10.5A2.75 2.75 0 0 0 18 15.25v-2.5a.75.75 0 0 0-1.5 0v2.5c0 .69-.56 1.25-1.25 1.25H4.75c-.69 0-1.25-.56-1.25-1.25v-2.5Z"/>' +
		'</svg>';

	// The chip fits its name. The limit (24rem, or the row's width on small screens) sits on the chip's box in
	// the row of files: a limit with a percentage on the chip itself is ignored when the box is sized, so the
	// box would keep the full width of the name and leave a gap. In light mode Open WebUI draws the chip white
	// on white, so it gets the grey of the user's message bubble. The size is smaller than the name: centred,
	// it floats above the name's baseline, so both share the baseline, and the line moves down 1px to centre
	// the name between the two icons.
	// Office colours for the document icon (the chip's first child), lighter in dark mode.
	const STYLE = `
		button[${MARK}="chip"] { width: auto !important; max-width: 100%; transition: background-color 0.15s; }
		[${MARK}="chip-box"] { max-width: min(24rem, 100%); min-width: 0; }
		html:not(.dark) button[${MARK}="chip"] { background-color: var(--color-gray-50, #f9f9f9); border-color: var(--color-gray-100, #ececec); }
		html:not(.dark) button[${MARK}="chip"]:hover { background-color: var(--color-gray-100, #ececec); }
		.dark button[${MARK}="chip"]:hover { background-color: var(--color-gray-800, #333); }
		button[${MARK}="chip"] .justify-between { align-items: baseline; gap: 0.375rem; position: relative; top: 1px; }
		button[${MARK}-type="docx"] { --office-file-ui-type: #185abd; }
		button[${MARK}-type="xlsx"] { --office-file-ui-type: #107c41; }
		button[${MARK}-type="pptx"] { --office-file-ui-type: #c43e1c; }
		.dark button[${MARK}-type="docx"] { --office-file-ui-type: #6ca0f6; }
		.dark button[${MARK}-type="xlsx"] { --office-file-ui-type: #4cc38a; }
		.dark button[${MARK}-type="pptx"] { --office-file-ui-type: #f08a64; }
		button[${MARK}-type] > :first-child { color: var(--office-file-ui-type); }
		[${MARK}="download"][data-busy] svg { animation: office-file-ui-pulse 0.8s ease-in-out infinite alternate; }
		[${MARK}="download"][data-failed] { color: #dc2626 !important; }
		@keyframes office-file-ui-pulse { to { opacity: 0.3; } }
		@media (pointer: coarse) { span[${MARK}="download"] { padding: 0.375rem; } }
		html.dark .office-preview { background: transparent; }
		html.dark .office-preview table th.excel-col-hdr,
		html.dark .office-preview .excel-row-num { background: var(--color-gray-850, #262626); }
	`;
	// In dark mode Open WebUI draws the spreadsheet preview in a blue-violet (#1a1a2e, headers #2a2a3e) instead
	// of the greys of the rest of the page; the rules above (one element more, so they win) use the greys.

	const decode = (text) => {
		try {
			return decodeURIComponent(text);
		} catch {
			return text;
		}
	};

	const officeType = (name) => name.match(PREVIEW_ONLY)?.[1]?.toLowerCase();

	const authHeaders = () => (localStorage.token ? { Authorization: `Bearer ${localStorage.token}` } : {});

	// ------------------------------------------------------------------ file ids from the chat --

	// The chips carry no file id in the page, so it is looked up in the saved chat.
	const chatState = (id, messages = {}, loadedAt = 0) => ({ id, messages, loading: null, loadedAt });
	let chatCache = chatState(null);
	// chips without a known file, by message, name and position: reloads so far and when the next one is due
	let misses = new Map();
	let retryTimer = null;

	const currentChatId = () => {
		const match = location.pathname.match(/^\/c\/([^/]+)/);
		return match && !match[1].startsWith('local') ? decode(match[1]) : null;
	};

	const loadChat = (chatId) => {
		chatCache.loading = (async () => {
			try {
				const response = await fetch(`/api/v1/chats/${encodeURIComponent(chatId)}`, { headers: authHeaders() });
				if (!response.ok) return false;
				const chat = await response.json();
				chatCache = chatState(chatId, chat?.chat?.history?.messages ?? {}, Date.now());
				return true;
			} catch (error) {
				console.warn('office-file-ui: loading the chat failed', error);
				return false;
			} finally {
				chatCache.loading = null;
			}
		})();
		return chatCache.loading;
	};

	const messageFiles = (messageId) =>
		(chatCache.messages[messageId]?.files ?? []).filter((file) => file?.id && file.type === 'file');

	// --------------------------------------------------------------------------- download --

	const download = async (control) => {
		const { fileId, fileName } = control.dataset;
		if (!fileId || control.dataset.busy) return;
		control.dataset.busy = '1';
		delete control.dataset.failed;
		control.title = LABEL;
		try {
			const response = await fetch(`/api/v1/files/${encodeURIComponent(fileId)}/content?attachment=true`, {
				headers: authHeaders()
			});
			if (!response.ok) throw new Error(`HTTP ${response.status}`);
			const url = URL.createObjectURL(await response.blob());
			const link = document.createElement('a');
			link.href = url;
			link.download = fileName;
			document.body.appendChild(link);
			link.click();
			link.remove();
			setTimeout(() => URL.revokeObjectURL(url), 10000);
		} catch (error) {
			console.warn('office-file-ui: download failed', error);
			control.dataset.failed = '1';
			control.title = FAILED;
			setTimeout(() => {
				delete control.dataset.failed;
				control.title = LABEL;
			}, 4000);
		} finally {
			delete control.dataset.busy;
		}
	};

	const downloadControl = (tag) => {
		const control = document.createElement(tag);
		control.setAttribute(MARK, 'download');
		control.title = LABEL;
		control.setAttribute('aria-label', LABEL);
		if (tag === 'button') control.type = 'button';
		else {
			control.setAttribute('role', 'button');
			control.tabIndex = 0;
		}
		control.addEventListener('click', (event) => {
			event.preventDefault();
			event.stopPropagation();
			download(control);
		});
		control.addEventListener('keydown', (event) => {
			if (event.key !== 'Enter' && event.key !== ' ') return;
			event.preventDefault();
			event.stopPropagation();
			download(control);
		});
		return control;
	};

	// The chip is a <button> itself, so the download icon inside it is a <span role="button">.
	// Its click does not reach the chip: the icon downloads, the rest of the chip opens the preview.
	const chipDownloadIcon = () => {
		const control = downloadControl('span');
		control.className =
			'shrink-0 -mr-1 p-1 rounded-lg text-gray-500 dark:text-gray-400 hover:text-gray-900 ' +
			'dark:hover:text-white hover:bg-black/5 dark:hover:bg-white/10 transition';
		control.innerHTML = downloadIcon('size-3.5');
		return control;
	};

	// ----------------------------------------------------------------- file chips of answers --

	const chipName = (chip) => decode(chip.querySelector('.truncate')?.textContent?.trim() ?? '');

	const decorateChip = (chip, name) => {
		if (chip.getAttribute(MARK) !== 'chip') chip.setAttribute(MARK, 'chip');
		const box = chip.parentElement;
		if (box?.children.length === 1 && box.getAttribute(MARK) !== 'chip-box') box.setAttribute(MARK, 'chip-box');
		const type = officeType(name);
		if (type && chip.getAttribute(`${MARK}-type`) !== type) chip.setAttribute(`${MARK}-type`, type);
	};

	const decorateAnswerFiles = () => {
		const chatId = currentChatId();
		if (chatId && chatCache.id !== chatId) {
			chatCache = chatState(chatId);
			misses = new Map();
		}
		const missing = [];
		for (const answer of document.querySelectorAll('.chat-assistant')) {
			const chips = [...answer.querySelectorAll(CHIP)];
			if (!chips.length) continue;
			const messageId = answer.closest('[id^="message-"]')?.id.slice('message-'.length);
			const files = chatId && messageId ? messageFiles(messageId) : [];
			const seen = {};
			for (const chip of chips) {
				const name = chipName(chip);
				decorateChip(chip, name);
				if (!chatId) continue; // temporary chat: not saved, no file ids to look up
				const index = (seen[name] = (seen[name] ?? -1) + 1);
				const file = files.filter((f) => decode(f.name ?? '') === name)[index];
				let icon = chip.querySelector(`[${MARK}="download"]`);
				if (!file) {
					missing.push(`${messageId}\n${name}\n${index}`);
					icon?.remove();
					continue;
				}
				if (!icon) {
					icon = chipDownloadIcon();
					chip.append(icon);
				}
				if (icon.dataset.fileId !== file.id) icon.dataset.fileId = file.id;
				if (icon.dataset.fileName !== (file.name ?? name)) icon.dataset.fileName = file.name ?? name;
				if (chip.dataset.officeFileId !== file.id) chip.dataset.officeFileId = file.id;
			}
		}
		// A chip without a known file: the chat was probably saved after the last load, so load it again, a few
		// times per chip with growing pauses. A chip whose file never shows up must not reload the chat forever.
		if (!missing.length || chatCache.loading) return;
		const now = Date.now();
		const due = missing.filter((key) => {
			const miss = misses.get(key) ?? { tries: 0, at: chatCache.loadedAt + RELOAD_DELAYS_MS[0] };
			misses.set(key, miss);
			return miss.tries < RELOAD_DELAYS_MS.length && now >= miss.at;
		});
		if (!due.length) {
			const next = Math.min(
				...missing.map((key) => misses.get(key)).filter((miss) => miss.tries < RELOAD_DELAYS_MS.length).map((miss) => miss.at)
			);
			clearTimeout(retryTimer); // the page may not change again before the next reload is due
			if (Number.isFinite(next)) retryTimer = setTimeout(schedule, next - now);
			return;
		}
		loadChat(chatId).then((loaded) => {
			for (const key of due) {
				const miss = misses.get(key);
				if (!miss) continue; // another chat by now
				if (loaded) {
					miss.tries += 1;
					miss.at = now + (RELOAD_DELAYS_MS[miss.tries] ?? 0);
				} else miss.at = Date.now() + FAILED_LOAD_DELAY_MS; // a failed request (network, 5xx) is no try
			}
			schedule();
		});
	};

	// Open WebUI puts the files of an answer above its text. Show them below it: the box around status,
	// files and text becomes a flex column and the files go last. Only CSS order changes; the elements
	// stay where Open WebUI put them, so its own updates keep working.
	const placeFilesBelowText = () => {
		for (const box of document.querySelectorAll('.chat-assistant > div')) {
			const files = [...box.children].find(
				(child) => !child.id && child.classList.contains('flex-wrap') && child.querySelector(CHIP)
			);
			if (!files || files.style.order) continue;
			box.style.display = 'flex';
			box.style.flexDirection = 'column';
			files.style.order = '1';
			files.style.marginTop = '0.5rem';
		}
	};

	// ----------------------------------------------------------------------- file view (modal) --

	// The same file view opens for uploads, the workspace and knowledge files. Only a view opened from
	// the chip of an answer is changed: remember that chip until the next file view appears.
	let openedFromAnswer = null;
	document.addEventListener(
		'click',
		(event) => {
			if (event.target.closest?.(`[${MARK}="download"]`)) return;
			const chip = event.target.closest?.(CHIP);
			if (!chip) return;
			openedFromAnswer = chip.closest('.chat-assistant')
				? { name: chipName(chip), fileId: chip.dataset.officeFileId ?? '' }
				: null;
		},
		true
	);

	// The line below the title reads "39 KB • 3 extracted lines • Formatting may be inconsistent from source.",
	// which is about the hidden "Content" tab. Keep the size, and for Excel the row count of the preview.
	const hideExtractedTextInfo = (modal, keepCount) => {
		const row = modal.querySelector('div.flex.flex-wrap.text-xs');
		const [size, count, ...rest] = row ? [...row.children] : [];
		if (!size || !count) return;
		for (const element of keepCount ? rest : [count, ...rest]) element.style.display = 'none';
		const separator = size.nextSibling;
		if (!keepCount && separator?.nodeType === Node.TEXT_NODE && separator.textContent) separator.textContent = '';
	};

	// Index of the selected tab of a two-tab bar: Open WebUI greys out the others with "border-transparent".
	const selectedTab = (bar) => {
		const selected = [...bar.children].map((tab) => !tab.classList.contains('border-transparent'));
		return selected.filter(Boolean).length === 1 ? selected.indexOf(true) : -1;
	};

	// The bar is hidden only while "Preview" (the second tab) is selected; otherwise it stays usable.
	const syncTabs = (bar) => {
		const display = selectedTab(bar) === 1 ? 'none' : '';
		if (bar.style.display !== display) bar.style.display = display;
	};

	const adjustFileView = () => {
		for (const modal of document.querySelectorAll('.modal[role="dialog"]')) {
			const title = modal.querySelector('a.line-clamp-1');
			if (!title) continue;
			const name = title.textContent.trim();

			if (!modal.hasAttribute(MARK)) {
				const source = openedFromAnswer?.name === name ? openedFromAnswer : null;
				openedFromAnswer = null;
				modal.setAttribute(MARK, source ? 'answer-file' : 'other');
				if (source?.fileId) modal.dataset.officeFileId = source.fileId;
			}
			if (modal.getAttribute(MARK) !== 'answer-file') continue;

			// "Content" | "Preview" bar: the first bar of exactly two tab buttons. "Content" is selected when the
			// view opens; click "Preview" and hide the bar only while "Preview" is really the selected tab, so the
			// bar stays usable if Open WebUI changes or reorders its tabs.
			const type = officeType(name);
			if (type) {
				const bar = [...modal.querySelectorAll('div.border-b')].find(
					(div) => div.children.length === 2 && [...div.children].every((c) => c.matches('button[type="button"]'))
				);
				if (bar && !bar.hasAttribute(MARK) && selectedTab(bar) === 0) {
					bar.setAttribute(MARK, 'tabs');
					bar.children[1].click(); // "Preview"
					requestAnimationFrame(() => syncTabs(bar)); // Open WebUI updates the tabs before the next frame
				} else if (bar?.getAttribute(MARK) === 'tabs') syncTabs(bar);
				hideExtractedTextInfo(modal, type === 'xlsx');
			}

			// download button next to the close button
			const header = title.closest('div.flex.items-start.justify-between');
			if (header && modal.dataset.officeFileId && !header.querySelector(`[${MARK}="download"]`)) {
				const button = downloadControl('button');
				button.dataset.fileId = modal.dataset.officeFileId;
				button.dataset.fileName = name;
				button.className = 'mr-2 text-gray-500 dark:text-gray-400 hover:text-gray-900 dark:hover:text-white transition';
				button.innerHTML = downloadIcon('size-5');
				header.lastElementChild.classList.add('flex', 'items-center');
				header.lastElementChild.prepend(button);
			}
		}
	};

	// ------------------------------------------------------------------ PowerPoint preview --

	// Open WebUI's PowerPoint preview draws every paragraph of a title or subtitle placeholder centred unless
	// it says "ctr" or "r": an explicit "l" (left) is drawn centred as well. The preview reads the slides with
	// DOMParser, so the slides are corrected right after parsing: a title or subtitle shape with its own
	// position whose paragraphs all give their alignment is marked as a body placeholder, which the preview
	// draws as the paragraphs say. The preview only uses the type for the default font size (kept by writing it
	// into each paragraph) and for the position when the shape has none (such shapes are left alone).
	const DRAWING_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';
	const TITLE_TYPES = { title: 3400, ctrTitle: 3400, subTitle: 2200 }; // the preview's default sizes

	const keepFontSize = (doc, para, size) => {
		let pPr = para.getElementsByTagName('a:pPr')[0];
		let defRPr = pPr && pPr.getElementsByTagName('a:defRPr')[0];
		if (!defRPr) {
			// the preview takes the paragraph's style from defRPr, else endParaRPr: copy that one
			defRPr = doc.createElementNS(DRAWING_NS, 'a:defRPr');
			const end = para.getElementsByTagName('a:endParaRPr')[0];
			if (end) {
				for (const attr of Array.from(end.attributes)) defRPr.setAttribute(attr.name, attr.value);
				for (const child of Array.from(end.childNodes)) defRPr.appendChild(child.cloneNode(true));
			}
			pPr.appendChild(defRPr);
		}
		if (!defRPr.hasAttribute('sz')) defRPr.setAttribute('sz', String(size));
	};

	const alignTitles = (doc) => {
		for (const ph of Array.from(doc.getElementsByTagName('p:ph'))) {
			const size = TITLE_TYPES[ph.getAttribute('type')];
			const shape = ph.closest('sp');
			const body = shape && shape.getElementsByTagName('p:txBody')[0];
			if (!size || !body || !shape.querySelector('spPr > xfrm')) continue;
			const paras = Array.from(body.getElementsByTagName('a:p'));
			const aligned = (para) => para.getElementsByTagName('a:pPr')[0]?.hasAttribute('algn');
			if (!paras.length || !paras.every(aligned)) continue;
			for (const para of paras) keepFontSize(doc, para, size);
			ph.setAttribute('type', 'body');
		}
	};

	if (typeof DOMParser !== 'undefined') {
		const parseFromString = DOMParser.prototype.parseFromString;
		DOMParser.prototype.parseFromString = function (text, type) {
			const doc = parseFromString.call(this, text, type);
			try {
				if (doc.documentElement?.nodeName === 'p:sld') alignTitles(doc);
			} catch (error) {
				console.warn('office-file-ui: slide not adjusted', error);
			}
			return doc;
		};
	}

	// ------------------------------------------------------------------------------- driver --

	// Runs once per frame at most and, while the page keeps changing (e.g. a streaming answer), at most every
	// RUN_INTERVAL_MS: every run scans all answers of the chat.
	let timer = null;
	let lastRun = -Infinity;
	function schedule() {
		if (timer !== null) return;
		const wait = Math.max(0, lastRun + RUN_INTERVAL_MS - performance.now());
		timer = setTimeout(() => requestAnimationFrame(run), wait);
	}

	function run() {
		timer = null;
		lastRun = performance.now();
		adjustFileView();
		placeFilesBelowText();
		decorateAnswerFiles();
	}

	const start = () => {
		const style = document.createElement('style');
		style.setAttribute(MARK, 'style');
		style.textContent = STYLE;
		document.head.appendChild(style);
		new MutationObserver(schedule).observe(document.body, { childList: true, subtree: true });
		schedule();
	};

	if (document.body) start();
	else document.addEventListener('DOMContentLoaded', start);
})();
