/* midas 応力抽出 Excel ワーカー
 * Pyodide（ブラウザ内 Python）上で make_excel.py の build() をそのまま実行する。
 * ファイルはワーカー内の仮想ファイルシステムだけで処理され、外部には送信されない。
 */
importScripts("https://cdn.jsdelivr.net/pyodide/v0.26.4/full/pyodide.js");

let pyodide = null;

async function init() {
  pyodide = await loadPyodide();
  await pyodide.loadPackage("micropip");
  const micropip = pyodide.pyimport("micropip");
  await micropip.install("openpyxl==3.1.5");
  const resp = await fetch("make_excel.py?v=" + Date.now());
  if (!resp.ok) throw new Error("make_excel.py の取得に失敗しました");
  pyodide.FS.writeFile("/home/pyodide/make_excel.py", await resp.text());
  pyodide.runPython("import sys; sys.path.insert(0, '/home/pyodide'); import make_excel");
  postMessage({ type: "ready" });
}

function convert(msg) {
  pyodide.FS.writeFile("/tmp/model.mgtx", new Uint8Array(msg.model));
  pyodide.FS.writeFile("/tmp/result.anl", new Uint8Array(msg.anl));
  pyodide.globals.set("js_log", (s) => postMessage({ type: "log", text: String(s) }));
  pyodide.globals.set("face_beam", msg.faceBeam);
  pyodide.globals.set("face_col", msg.faceCol);
  pyodide.globals.set("model_name", msg.modelName);
  pyodide.globals.set("anl_name", msg.anlName);
  try {
    pyodide.runPython(`
import make_excel
make_excel.build("/tmp/model.mgtx", "/tmp/result.anl", "/tmp/out.xlsx", face_beam, face_col,
                 log=js_log, model_label=model_name, anl_label=anl_name)
`);
    const bytes = pyodide.FS.readFile("/tmp/out.xlsx");
    postMessage({ type: "done", bytes: bytes.buffer, filename: msg.filename }, [bytes.buffer]);
  } finally {
    for (const p of ["/tmp/model.mgtx", "/tmp/result.anl", "/tmp/out.xlsx"]) {
      try { pyodide.FS.unlink(p); } catch (e) { /* noop */ }
    }
  }
}

self.onmessage = async (e) => {
  const msg = e.data;
  try {
    if (msg.type === "init") await init();
    else if (msg.type === "convert") convert(msg);
  } catch (err) {
    const text = (err && err.message) || String(err);
    const lastLine = text.trim().split(/\r?\n/).pop();
    postMessage({ type: "error", text, summary: lastLine.replace(/^(?:\w+\.)*\w*(?:Error|Exception):\s*/, "") });
  }
};
