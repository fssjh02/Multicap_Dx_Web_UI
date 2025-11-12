from flask import Flask, request, send_file, render_template_string, jsonify
import time, os, io, base64, re
from datetime import datetime
import numpy as np
from PIL import Image
import serial
import xlsxwriter

# ---------------------- App ----------------------
app = Flask(__name__, static_folder="static")

# ---------------------- Configuration ----------------------
PORT_NAME    = "COM7"
BAUDRATE     = 115200
READ_TIMEOUT = 3.0
TOTAL_WAIT   = 6.0

W, H = 160, 160
ROI_SIZE = 50
N = W * H

# ---- Cutoffs (점수 범위 0~2500) ----
CUTOFF_HIV = 97.5
CUTOFF_HBV = 195.5
CUTOFF_HCV = 134.7

# ---- ROI 중심좌표 ----
DEFAULT_ROIS = [
    {"cx": 35,  "cy": 125},  # ROI1 (Internal Control)
    {"cx": 125, "cy": 125},  # ROI2 (HIV)
    {"cx": 35,  "cy": 35},   # ROI3 (HBV)
    {"cx": 125, "cy": 35},   # ROI4 (HCV)
]

last_rotated = None

# ---------------------- Utilities ----------------------
def create_dir(path: str):
    if not os.path.exists(path):
        os.makedirs(path)

def read_pixels_from_serial(port=PORT_NAME, baud=BAUDRATE, timeout=READ_TIMEOUT, total_wait=TOTAL_WAIT):
    """160x160 grayscale frame from MCU (0~255 integers)"""
    with serial.Serial(port, baudrate=baud, timeout=timeout) as mcu:
        mcu.dtr = False; mcu.rts = False
        time.sleep(0.05)
        mcu.reset_input_buffer(); mcu.reset_output_buffer()
        mcu.write(b"99\n")
        time.sleep(0.05)
        tokens, buf = [], b""
        start = time.time()
        while len(tokens) < N:
            chunk = mcu.read(4096)
            if not chunk:
                if time.time() - start > total_wait:
                    break
                continue
            buf += chunk
            text = buf.decode("utf-8", errors="ignore")
            parts = re.split(r"[,\s]+", text.strip())
            tail = ""
            if parts and not text.endswith((",", " ", "\n", "\r", "\t")):
                tail = parts.pop()
            for p in parts:
                if p.isdigit():
                    tokens.append(int(p))
                elif p.startswith("-") and p[1:].isdigit():
                    v = max(0, min(255, int(p)))
                    tokens.append(v)
                if len(tokens) >= N:
                    break
            buf = tail.encode("utf-8")

        if len(tokens) < N:
            raise RuntimeError(f"Not enough data from MCU: {len(tokens)} of {N}")
        arr = np.array(tokens[:N], dtype=np.uint8).reshape((H, W))
        return arr

def numpy_to_png_base64_gray(array2d: np.ndarray, scale=3) -> str:
    img = Image.fromarray(array2d, mode='L')
    if scale and scale != 1:
        img = img.resize((array2d.shape[1]*scale, array2d.shape[0]*scale), resample=Image.NEAREST)
    bio = io.BytesIO()
    img.save(bio, format="PNG")
    return base64.b64encode(bio.getvalue()).decode("utf-8")

def clamp_centroid(cx, cy):
    half = ROI_SIZE // 2
    return max(half, min(W - half, int(cx))), max(half, min(H - half, int(cy)))

def centroid_to_tl(cx, cy):
    half = ROI_SIZE // 2
    return int(cx - half), int(cy - half)

# ---------------------- HTML Template ----------------------
INDEX_HTML = """
<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8" />
<title>Multicap Dx | Shafiee Lab</title>
<meta name="viewport" content="width=device-width, initial-scale=1" />
<style>
:root{
  --bg:#ffffff;
  --card:#f6f7fb;
  --line:#e6e9ef;
  --text:#0b0b10;
  --accent:#5aa2ff;
  --btn:#2f80ed;
  --pos:#1aa351;
  --neg:#d63636;
}
*{box-sizing:border-box}
body{font-family:Inter,system-ui,Arial,sans-serif;background:var(--bg);color:var(--text);margin:0}
.container{max-width:1200px;margin:0 auto;padding:18px}
.topbar{background:#ffffff;color:#000;display:flex;align-items:center;gap:20px;padding:12px 20px;border-bottom:2px solid var(--accent)}
.topbar img{height:48px}
.topbar-text{display:flex;flex-direction:column}
.topbar-text .lab{font-weight:700;font-size:18px;color:#000}
.topbar-text .inst{font-size:13px;color:#444}
.topbar-text .proj{font-size:12px;color:var(--accent);margin-top:3px}

/* layout cards */
.row{display:flex;gap:16px;flex-wrap:wrap;margin-top:16px}
.card{background:var(--card);border:1px solid var(--line);border-radius:10px;box-shadow:0 6px 18px rgba(16,24,40,0.04)}
.card .header{display:flex;align-items:center;justify-content:space-between;padding:12px 14px;border-bottom:1px solid var(--line);font-weight:700}
.card .body{padding:14px}

/* controls */
.controls {display:flex;gap:10px;align-items:center}
button.run{background:var(--btn);border:0;color:#fff;padding:10px 16px;border-radius:8px;cursor:pointer;font-weight:700}
button.extract{background:#fff;border:1px solid var(--line);color:var(--text);padding:8px 12px;border-radius:8px;cursor:pointer}
button.reset{background:#fff;border:1px solid var(--line);color:var(--text);padding:6px 10px;border-radius:6px;cursor:pointer}

/* image area */
.imgwrap{position:relative;width:480px;height:480px;background:#f0f2f6;border-radius:8px;display:flex;align-items:center;justify-content:center;overflow:hidden;border:1px dashed var(--line)}
#img{width:100%;height:100%;object-fit:contain;display:block}
#overlay{position:absolute;left:0;top:0;pointer-events:none}

/* ROI controls */
.roi-grid{display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-top:12px}
.roi-box{border:1px dashed var(--line);padding:8px;border-radius:8px;background:#fff}
.arrow-grid{display:grid;grid-template-columns:repeat(3,34px);gap:6px;justify-content:start}
.arrow-grid button{padding:6px;font-size:14px;background:#fff;border:1px solid var(--line);border-radius:6px;cursor:pointer}

/* results */
.results{margin-top:8px}
.result-row{margin:6px 0;font-size:15px}
.result-status{font-weight:800;font-size:16px}
.pos{color:var(--pos)}
.neg{color:var(--neg)}

/* footer */
.footer{text-align:center;font-size:12px;color:#666;margin-top:30px;padding:12px 0;border-top:1px solid var(--line)}

/* small helpers */
.legend{font-size:13px;color:#444;margin-top:8px}
.hint{font-size:13px;color:#333;margin-top:6px}
</style>
</head>
<body>
<header class="topbar">
  <img src="/static/lab_logo.png" alt="Shafiee Lab Logo">
  <div class="topbar-text">
    <div class="lab">Shafiee Lab</div>
    <div class="inst">Brigham and Women's Hospital | Harvard Medical School</div>
    <div class="proj">Multicap Dx — Bubble-Induced 3D Capacitance Profiling for Quantitative Multi-Viral Antigen Detection</div>
  </div>
</header>

<div class="container">
  <div class="row">
    <!-- Controls card -->
    <div class="card" style="flex:0 0 320px">
      <div class="header">Controls</div>
      <div class="body">
        <div class="controls">
          <button id="btnRun" class="run">RUN</button>
        </div>
        <div class="hint">After dispensing the sample into all chambers, press the RUN button.</div>
      </div>
    </div>

    <!-- Image & ROI card -->
    <div class="card" style="flex:1 1 620px">
      <div class="header">
        <div>Image & ROI</div>
        <div><button id="btnResetROI" class="reset">Reset ROI</button></div>
      </div>
      <div class="body">
        <div style="display:flex;gap:16px;flex-wrap:wrap">
          <div class="imgwrap" id="imgwrap">
            <div id="placeholder" style="text-align:center;color:#888;font-size:15px;">
              No image captured yet
            </div>
            <img id="img" alt="frame" style="display:none"/>
            <svg id="overlay"></svg>
          </div>

          <div style="flex:1 1 240px;min-width:220px">
            <div style="font-weight:700;margin-bottom:8px">ROIs (center coordinates)</div>
            <div class="roi-grid" id="roiGrid"></div>
            <div class="legend">Use arrows to move the ROI center by 1 pixel.</div>
          </div>
        </div>

        <!-- Extract button just under the image area -->
        <div style="margin-top:12px">
          <button id="btnExtract" class="extract">Extract & Analyze</button>
        </div>
      </div>
    </div>

    <!-- Results card -->
    <div class="card" style="flex:0 0 260px">
      <div class="header">Results</div>
      <div class="body">
        <div id="downloadArea"></div>
        <div id="results" class="results">
          <div class="result-row">No analysis yet</div>
        </div>
      </div>
    </div>
  </div>

  <div class="footer">
    <div>© 2025 Shafiee Lab, Brigham and Women's Hospital, Harvard Medical School</div>
    <div>Developed under the supervision of Dr. Hadi Shafiee (hshafiee@bwh.harvard.edu)</div>
    <div>For Research and Clinical Evaluation Purposes Only | MultiCapDx Web Interface v4.2 (Nat.Comm.)</div>
  </div>
</div>

<script>
const W=160, H=160, ROI_SIZE=50;
let scale = 3;
let rois = {{ default_rois | tojson }};

function clamp(v,min,max){ return Math.max(min, Math.min(max, v)); }
function centroidToTL(cx, cy){ const h = Math.floor(ROI_SIZE/2); return {x: cx - h, y: cy - h}; }

function drawOverlay(){
  const svg = document.getElementById("overlay");
  const imgWrap = document.getElementById("imgwrap");
  const w = W*scale, h = H*scale;
  svg.setAttribute("width", w);
  svg.setAttribute("height", h);
  svg.style.width = w + "px";
  svg.style.height = h + "px";
  svg.innerHTML = "";
  const colors = ["#00a3ff","#ff6b6b","#ffd93d","#7cff91"];
  const names = ["Internal","HIV","HBV","HCV"];

  svg.style.position = "absolute";
  svg.style.left = "0";
  svg.style.top = "0";

  rois.forEach((r,i) => {
    const tl = centroidToTL(r.cx, r.cy);
    const rect = document.createElementNS("http://www.w3.org/2000/svg","rect");
    rect.setAttribute("x", tl.x*scale);
    rect.setAttribute("y", tl.y*scale);
    rect.setAttribute("width", ROI_SIZE*scale);
    rect.setAttribute("height", ROI_SIZE*scale);
    rect.setAttribute("fill", "none");
    rect.setAttribute("stroke", colors[i%colors.length]);
    rect.setAttribute("stroke-width", "2");
    svg.appendChild(rect);

    const label = document.createElementNS("http://www.w3.org/2000/svg","text");
    label.setAttribute("x", tl.x*scale + 6);
    label.setAttribute("y", tl.y*scale + 14);
    label.setAttribute("fill", colors[i%colors.length]);
    label.setAttribute("font-size", "12px");
    label.setAttribute("font-weight", "700");
    label.textContent = names[i] + ` (${r.cx},${r.cy})`;
    svg.appendChild(label);
  });
}

// ------------------ ROI Controls 배열 재배치 ------------------
function buildROIControls(){
  const grid = document.getElementById("roiGrid");
  grid.innerHTML = "";

  const names = ["Internal","HIV","HBV","HCV"];

  // 위쪽: ROI1 + ROI2
  [2,3].forEach(idx => {
    const r = rois[idx];
    const box = document.createElement("div");
    box.className = "roi-box";
    box.innerHTML = `
      <div style="font-weight:700;margin-bottom:6px">${names[idx]}</div>
      <div style="font-size:13px;margin-bottom:6px">cx=${r.cx}, cy=${r.cy}</div>
      <div class="arrow-grid">
        <div></div>
        <button data-i="${idx}" data-dx="0" data-dy="-1">↑</button>
        <div></div>
        <button data-i="${idx}" data-dx="-1" data-dy="0">←</button>
        <div></div>
        <button data-i="${idx}" data-dx="1" data-dy="0">→</button>
        <div></div>
        <button data-i="${idx}" data-dx="0" data-dy="1">↓</button>
        <div></div>
      </div>
    `;
    grid.appendChild(box);
  });

  // 아래쪽: ROI3 + ROI4
  [0,1].forEach(idx => {
    const r = rois[idx];
    const box = document.createElement("div");
    box.className = "roi-box";
    box.innerHTML = `
      <div style="font-weight:700;margin-bottom:6px">${names[idx]}</div>
      <div style="font-size:13px;margin-bottom:6px">cx=${r.cx}, cy=${r.cy}</div>
      <div class="arrow-grid">
        <div></div>
        <button data-i="${idx}" data-dx="0" data-dy="-1">↑</button>
        <div></div>
        <button data-i="${idx}" data-dx="-1" data-dy="0">←</button>
        <div></div>
        <button data-i="${idx}" data-dx="1" data-dy="0">→</button>
        <div></div>
        <button data-i="${idx}" data-dx="0" data-dy="1">↓</button>
        <div></div>
      </div>
    `;
    grid.appendChild(box);
  });

  // 버튼 이벤트 연결
  grid.querySelectorAll("button").forEach(btn => {
    btn.addEventListener("click", () => {
      const i = +btn.dataset.i;
      const dx = +btn.dataset.dx;
      const dy = +btn.dataset.dy;
      const half = Math.floor(ROI_SIZE/2);
      rois[i].cx = clamp(rois[i].cx + dx, half, W - half);
      rois[i].cy = clamp(rois[i].cy + dy, half, H - half);
      buildROIControls();
      drawOverlay();
    });
  });
}

// ------------------ 나머지 JS 동일 ------------------
function showPlaceholder(show){
  const img = document.getElementById("img");
  const placeholder = document.getElementById("placeholder");
  if(show){
    placeholder.style.display = "block";
    img.style.display = "none";
  } else {
    placeholder.style.display = "none";
    img.style.display = "block";
  }
}

async function runCapture(){
  showPlaceholder(true);
  try{
    const res = await fetch("/api/generate", { method: "POST" });
    const data = await res.json();
    if(!data.ok) return alert("Capture error: " + (data.error || "unknown"));
    const img = document.getElementById("img");
    img.src = "data:image/png;base64," + data.image_b64;
    img.onload = () => {
      showPlaceholder(false);
      drawOverlay();
    };
  }catch(e){
    alert("Capture failed: " + e.message);
  }
}

async function doExtract(){
  try{
    const res = await fetch("/api/extract", {
      method: "POST",
      headers: {"Content-Type": "application/json"},
      body: JSON.stringify({ rois })
    });
    const data = await res.json();
    if(!data.ok) return alert("Extract error: " + (data.error || "unknown"));

    document.getElementById("downloadArea").innerHTML =
      `<a href="/download/${data.csv}" target="_blank">Download CSV (normalized)</a>`;

    const results = document.getElementById("results");
    const hivClass = data.hiv.status === "Positive" ? "pos" : "neg";
    const hbvClass = data.hbv.status === "Positive" ? "pos" : "neg";
    const hcvClass = data.hcv.status === "Positive" ? "pos" : "neg";

    results.innerHTML = `
      <div class="result-row"><b>Internal Control:</b> <span class="result-status">${data.ic_ok}</span></div>
      <div class="result-row"><b>HIV:</b> <span class="result-status ${hivClass}" style="font-size:18px">${data.hiv.status}</span>
        <div style="font-size:13px;color:#444">score=${data.hiv.score.toFixed(2)}, cutoff=${data.hiv.cutoff}</div></div>
      <div class="result-row"><b>HBV:</b> <span class="result-status ${hbvClass}" style="font-size:18px">${data.hbv.status}</span>
        <div style="font-size:13px;color:#444">score=${data.hbv.score.toFixed(2)}, cutoff=${data.hbv.cutoff}</div></div>
      <div class="result-row"><b>HCV:</b> <span class="result-status ${hcvClass}" style="font-size:18px">${data.hcv.status}</span>
        <div style="font-size:13px;color:#444">score=${data.hcv.score.toFixed(2)}, cutoff=${data.hcv.cutoff}</div></div>
    `;
  }catch(e){
    alert("Extract failed: " + e.message);
  }
}

document.addEventListener("DOMContentLoaded", () => {
  buildROIControls();
  drawOverlay();
  showPlaceholder(true);

  document.getElementById("btnRun").addEventListener("click", runCapture);
  document.getElementById("btnExtract").addEventListener("click", doExtract);
  document.getElementById("btnResetROI").addEventListener("click", () => {
    rois = {{ default_rois | tojson }};
    buildROIControls(); drawOverlay();
  });
});
</script>
</body>
</html>
"""

# ---------------------- Routes ----------------------
@app.route("/")
def index():
    return render_template_string(INDEX_HTML, default_rois=DEFAULT_ROIS)

@app.route("/api/generate", methods=["POST"])
def api_generate():
    global last_rotated
    try:
        frame = read_pixels_from_serial()
    except Exception as e:
        print("[Serial] Fallback:", e)
        frame = (np.random.rand(H, W) * 255).astype(np.uint8)
    rotated = np.rot90(frame, k=-1)
    last_rotated = rotated.copy()
    image_b64 = numpy_to_png_base64_gray(rotated, scale=3)
    return jsonify({"ok": True, "image_b64": image_b64, "scale": 3})

@app.route("/api/extract", methods=["POST"])
def api_extract():
    global last_rotated
    if last_rotated is None:
        return jsonify({"ok": False, "error": "Capture first"}), 400

    payload = request.get_json(silent=True) or {}
    rois = payload.get("rois") or []
    if len(rois) != 4:
        return jsonify({"ok": False, "error": "Need 4 ROIs"}), 400

    coords = []
    for r in rois:
        cx, cy = clamp_centroid(r.get("cx"), r.get("cy"))
        x, y = centroid_to_tl(cx, cy)
        coords.append((x, y))

    arr = last_rotated.copy()
    vals = [arr[y:y+ROI_SIZE, x:x+ROI_SIZE].reshape(-1) for x, y in coords]
    all_vals = np.concatenate(vals)
    vmin, vmax = int(all_vals.min()), int(all_vals.max())
    if vmax > vmin:
        norm = np.rint((all_vals - vmin) * (255.0 / (vmax - vmin))).astype(np.uint8)
    else:
        norm = np.zeros_like(all_vals, dtype=np.uint8)

    roi1, roi2, roi3, roi4 = [norm[i*2500:(i+1)*2500] for i in range(4)]

    ic_ok = not np.all(roi1 == 0)

    def score01_sum(roi):
        return float((roi.astype(np.float64) / 255.0).sum())

    hiv_score = score01_sum(roi2)
    hbv_score = score01_sum(roi3)
    hcv_score = score01_sum(roi4)

    hiv_status = "Positive" if hiv_score > CUTOFF_HIV else "Negative"
    hbv_status = "Positive" if hbv_score > CUTOFF_HBV else "Negative"
    hcv_status = "Positive" if hcv_score > CUTOFF_HCV else "Negative"

    now = datetime.now()
    date_now = now.strftime("%m-%d-%Y")
    time_file = now.strftime("%H-%M-%S")
    out_dir = f"roi_extract/{date_now}"
    create_dir(out_dir)
    csv_path = f"{out_dir}/{time_file}_ROI_NORMALIZED.csv"
    np.savetxt(csv_path, norm[None, :], fmt="%d", delimiter=",")

    return jsonify({
        "ok": True,
        "csv": csv_path,
        "ic_ok": bool(ic_ok),
        "hiv": {"status": hiv_status, "score": hiv_score, "cutoff": CUTOFF_HIV},
        "hbv": {"status": hbv_status, "score": hbv_score, "cutoff": CUTOFF_HBV},
        "hcv": {"status": hcv_status, "score": hcv_score, "cutoff": CUTOFF_HCV},
        "vmin": vmin, "vmax": vmax
    })

@app.route("/download/<path:filename>")
def download_file(filename):
    safe = os.path.normpath(filename)
    if not (safe.startswith("roi_extract")):
        return "Invalid path", 400
    if not os.path.exists(safe):
        return "File not found", 404
    return send_file(safe, as_attachment=True)

# ---------------------- Run ----------------------
if __name__ == "__main__":
    print("Starting Multicap Dx UI on http://127.0.0.1:5050 ...")
    app.run(host="127.0.0.1", port=5050, debug=True)
