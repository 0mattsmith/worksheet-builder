/*
 * Worksheet Builder – link between the pop-out window and the task pane.
 * Office only lets the task pane change the document, so the pop-out window sends each
 * document operation to the task pane, which runs it and sends the result back.
 * Messages are JSON; long ones are split into chunks so they fit any Office message limit.
 */
(function (global) {
  'use strict';
  const CHUNK = 30000;
  let seq = 0;

  function pack(obj) {
    const s = JSON.stringify(obj);
    if (s.length <= CHUNK) return [s];
    const mid = 'm' + Date.now().toString(36) + (seq++);
    const n = Math.ceil(s.length / CHUNK);
    const parts = [];
    for (let i = 0; i < n; i++) parts.push(JSON.stringify({ t: 'part', mid, i, n, d: s.slice(i * CHUNK, (i + 1) * CHUNK) }));
    return parts;
  }

  function makeReceiver(onMessage) {
    const buf = {};
    return (str) => {
      let m;
      try { m = JSON.parse(str); } catch (e) { return; }
      if (m && m.t === 'part') {
        const b = buf[m.mid] || (buf[m.mid] = { n: m.n, got: 0, parts: [] });
        if (b.parts[m.i] === undefined) { b.parts[m.i] = m.d; b.got++; }
        if (b.got < b.n) return;
        delete buf[m.mid];
        try { m = JSON.parse(b.parts.join('')); } catch (e) { return; }
      }
      if (m) onMessage(m);
    };
  }

  const DIALOG_ERRORS = {
    12004: 'The window address is not on the add-in’s own site.',
    12005: 'The window must use https.',
    12007: 'Worksheet Builder already has a window open.',
    12009: 'The window was blocked. Allow pop-ups for Word and try again.',
    12011: 'Your browser blocked the window. Allow pop-ups for this site and try again.',
  };

  /**
   * Task pane side: open the pop-out window and serve its requests.
   * opts: { url, width, height, handle(op, args) -> Promise, onClose(reason) }
   */
  function openWindow(opts) {
    return new Promise((resolve, reject) => {
      Office.context.ui.displayDialogAsync(opts.url,
        { width: opts.width || 60, height: opts.height || 85, displayInIframe: false, promptBeforeOpen: false },
        (res) => {
          if (res.status !== Office.AsyncResultStatus.Succeeded) {
            const code = res.error && res.error.code;
            return reject(new Error(DIALOG_ERRORS[code] || ('Word could not open the window (' + (res.error ? res.error.message : code) + ').')));
          }
          const dialog = res.value;
          let closed = false;
          const finish = (reason) => { if (closed) return; closed = true; if (opts.onClose) opts.onClose(reason); };
          const send = (obj) => { if (!closed) pack(obj).forEach((s) => dialog.messageChild(s)); };
          const receive = makeReceiver(async (m) => {
            if (m.t === 'req') {
              try {
                const result = await opts.handle(m.op, m.args || {});
                send({ t: 'res', id: m.id, ok: true, result: result === undefined ? null : result });
              } catch (e) {
                send({ t: 'res', id: m.id, ok: false, error: (e && e.message) || String(e) });
              }
            } else if (m.t === 'dock') {
              try { dialog.close(); } catch (e) { /* already closed */ }
              finish('dock');
            }
          });
          dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => receive(arg.message));
          dialog.addEventHandler(Office.EventType.DialogEventReceived, (arg) => finish(arg && arg.error));
          resolve({ close() { try { dialog.close(); } catch (e) { /* ignore */ } finish('closed'); } });
        });
    });
  }

  /**
   * Pop-out window side: returns { ready, call(op, args), dock() }.
   */
  function connect() {
    const pending = {};
    let n = 0;
    const receive = makeReceiver((m) => {
      if (m.t !== 'res' || !pending[m.id]) return;
      const p = pending[m.id];
      delete pending[m.id];
      clearTimeout(p.timer);
      if (m.ok) p.resolve(m.result); else p.reject(new Error(m.error));
    });
    const ready = new Promise((resolve, reject) => {
      Office.context.ui.addHandlerAsync(Office.EventType.DialogParentMessageReceived, (arg) => receive(arg.message), (r) => {
        if (r && r.status === Office.AsyncResultStatus.Failed) reject(new Error('This window can’t talk to Word (needs Microsoft 365 or Word 2021+).'));
        else resolve(true);
      });
    });
    const post = (obj) => pack(obj).forEach((s) => Office.context.ui.messageParent(s));
    return {
      ready,
      call(op, args) {
        return ready.then(() => new Promise((resolve, reject) => {
          const id = 'r' + (++n);
          const timer = setTimeout(() => {
            delete pending[id];
            reject(new Error('Word didn’t respond. Is the Worksheet Builder panel still open in Word?'));
          }, op === 'state' ? 15000 : 300000);
          pending[id] = { resolve, reject, timer };
          post({ t: 'req', id, op, args: args || {} });
        }));
      },
      dock() { post({ t: 'dock' }); },
    };
  }

  global.WSBridge = { openWindow, connect, pack, makeReceiver };
})(typeof window !== 'undefined' ? window : globalThis);
