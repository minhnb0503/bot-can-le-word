# -*- coding: utf-8 -*-
"""
Bot Căn Chỉnh Báo Cáo — Web App
Mở trình duyệt, kéo thả file .docx vào, tùy chỉnh thông số, tải kết quả.
"""

import os
import sys
import uuid
import threading
import webbrowser
import tempfile
from pathlib import Path

from flask import Flask, request, send_file, jsonify, render_template_string

# Thêm thư mục chứa format_engine vào sys.path
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
if SCRIPT_DIR not in sys.path:
    sys.path.insert(0, SCRIPT_DIR)

from format_engine import format_document

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB max

# Thư mục tạm cho file kết quả
TEMP_DIR = os.path.join(tempfile.gettempdir(), 'format_bot_output')
os.makedirs(TEMP_DIR, exist_ok=True)

# ════════════════════════════════════════════════════════════
# HTML TEMPLATE
# ════════════════════════════════════════════════════════════

HTML_PAGE = r"""
<!DOCTYPE html>
<html lang="vi">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Bot Căn Chỉnh Báo Cáo</title>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap" rel="stylesheet">
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }

        body {
            font-family: 'Inter', sans-serif;
            min-height: 100vh;
            background: linear-gradient(135deg, #0f0c29, #302b63, #24243e);
            display: flex;
            flex-direction: column;
            align-items: center;
            justify-content: center;
            color: #e0e0e0;
            overflow-x: hidden;
        }

        /* Animated background blobs */
        .bg-blob {
            position: fixed;
            border-radius: 50%;
            filter: blur(100px);
            opacity: 0.15;
            animation: float 20s ease-in-out infinite;
            z-index: 0;
        }
        .bg-blob:nth-child(1) {
            width: 500px; height: 500px;
            background: #667eea;
            top: -100px; left: -100px;
            animation-duration: 25s;
        }
        .bg-blob:nth-child(2) {
            width: 400px; height: 400px;
            background: #f093fb;
            bottom: -80px; right: -80px;
            animation-duration: 20s;
            animation-delay: -5s;
        }
        .bg-blob:nth-child(3) {
            width: 300px; height: 300px;
            background: #4facfe;
            top: 50%; left: 60%;
            animation-duration: 18s;
            animation-delay: -10s;
        }

        @keyframes float {
            0%, 100% { transform: translate(0, 0) rotate(0deg); }
            25% { transform: translate(30px, -40px) rotate(5deg); }
            50% { transform: translate(-20px, 30px) rotate(-3deg); }
            75% { transform: translate(40px, 20px) rotate(4deg); }
        }

        /* Main card */
        .card {
            position: relative;
            z-index: 1;
            background: rgba(255, 255, 255, 0.06);
            backdrop-filter: blur(24px);
            -webkit-backdrop-filter: blur(24px);
            border: 1px solid rgba(255, 255, 255, 0.1);
            border-radius: 24px;
            padding: 40px 48px;
            width: 580px;
            max-width: 94vw;
            text-align: center;
            box-shadow: 0 25px 60px rgba(0,0,0,0.3),
                        inset 0 1px 0 rgba(255,255,255,0.1);
        }

        .title {
            font-size: 1.6rem;
            font-weight: 800;
            background: linear-gradient(135deg, #667eea, #764ba2, #f093fb);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
            background-clip: text;
            margin-bottom: 8px;
            letter-spacing: -0.02em;
        }
        .subtitle {
            font-size: 0.85rem;
            color: rgba(255,255,255,0.4);
            margin-bottom: 28px;
        }

        /* Drop zone */
        .drop-zone {
            border: 2px dashed rgba(255,255,255,0.2);
            border-radius: 16px;
            padding: 44px 24px;
            cursor: pointer;
            transition: all 0.3s ease;
            background: rgba(255,255,255,0.02);
            position: relative;
        }
        .drop-zone:hover, .drop-zone.drag-over {
            border-color: #667eea;
            background: rgba(102,126,234,0.08);
            transform: scale(1.01);
        }
        .drop-zone.drag-over {
            box-shadow: 0 0 30px rgba(102,126,234,0.2);
        }

        .drop-icon { font-size: 3.5rem; margin-bottom: 16px; display: block; }
        .drop-text {
            font-size: 1.05rem; font-weight: 600;
            color: rgba(255,255,255,0.8); margin-bottom: 6px;
        }
        .drop-hint { font-size: 0.8rem; color: rgba(255,255,255,0.35); }

        #fileInput { display: none; }

        /* ═══ SETTINGS PANEL ═══ */
        .settings-toggle {
            display: inline-flex;
            align-items: center;
            gap: 6px;
            margin-top: 20px;
            padding: 8px 20px;
            border-radius: 10px;
            background: rgba(255,255,255,0.06);
            border: 1px solid rgba(255,255,255,0.1);
            color: rgba(255,255,255,0.5);
            font-size: 0.8rem;
            font-weight: 600;
            cursor: pointer;
            transition: all 0.25s ease;
            user-select: none;
        }
        .settings-toggle:hover {
            background: rgba(255,255,255,0.1);
            color: rgba(255,255,255,0.8);
        }
        .settings-toggle .arrow {
            transition: transform 0.3s ease;
            font-size: 0.65rem;
        }
        .settings-toggle.open .arrow { transform: rotate(180deg); }

        .settings-panel {
            max-height: 0;
            overflow: hidden;
            transition: max-height 0.4s ease, margin 0.3s ease, opacity 0.3s ease;
            opacity: 0;
            margin-top: 0;
        }
        .settings-panel.open {
            max-height: 500px;
            opacity: 1;
            margin-top: 16px;
        }

        .settings-grid {
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 10px;
            text-align: left;
        }

        .setting-item {
            background: rgba(255,255,255,0.04);
            border: 1px solid rgba(255,255,255,0.07);
            border-radius: 10px;
            padding: 10px 12px;
        }
        .setting-item label {
            display: block;
            font-size: 0.68rem;
            font-weight: 600;
            color: rgba(255,255,255,0.4);
            margin-bottom: 4px;
            text-transform: uppercase;
            letter-spacing: 0.04em;
        }
        .setting-item input, .setting-item select {
            width: 100%;
            background: rgba(255,255,255,0.06);
            border: 1px solid rgba(255,255,255,0.1);
            border-radius: 6px;
            padding: 6px 10px;
            color: #fff;
            font-family: 'Inter', sans-serif;
            font-size: 0.82rem;
            font-weight: 500;
            outline: none;
            transition: border-color 0.2s;
        }
        .setting-item input:focus, .setting-item select:focus {
            border-color: #667eea;
        }
        .setting-item select option {
            background: #1a1a2e;
            color: #fff;
        }

        .settings-full-row {
            grid-column: 1 / -1;
        }

        .settings-divider {
            grid-column: 1 / -1;
            height: 1px;
            background: rgba(255,255,255,0.06);
            margin: 4px 0;
        }

        .btn-reset {
            grid-column: 1 / -1;
            background: rgba(255,255,255,0.04);
            border: 1px solid rgba(255,255,255,0.08);
            border-radius: 8px;
            padding: 7px;
            color: rgba(255,255,255,0.4);
            font-size: 0.75rem;
            font-weight: 600;
            cursor: pointer;
            transition: all 0.2s;
            font-family: 'Inter', sans-serif;
        }
        .btn-reset:hover {
            background: rgba(255,255,255,0.08);
            color: rgba(255,255,255,0.7);
        }

        /* Processing state */
        .processing { display: none; padding: 32px 0; }
        .processing.active { display: block; }

        .spinner {
            width: 56px; height: 56px;
            border: 4px solid rgba(255,255,255,0.1);
            border-top: 4px solid #667eea;
            border-radius: 50%;
            animation: spin 0.8s linear infinite;
            margin: 0 auto 20px;
        }
        @keyframes spin { to { transform: rotate(360deg); } }

        .processing-text { font-size: 1rem; font-weight: 600; color: rgba(255,255,255,0.7); }
        .processing-file { font-size: 0.8rem; color: rgba(255,255,255,0.35); margin-top: 6px; word-break: break-all; }

        .progress-bar-container {
            width: 100%; height: 4px;
            background: rgba(255,255,255,0.08);
            border-radius: 2px; margin-top: 20px; overflow: hidden;
        }
        .progress-bar {
            height: 100%;
            background: linear-gradient(90deg, #667eea, #764ba2);
            border-radius: 2px; width: 0%;
            animation: indeterminate 1.5s ease-in-out infinite;
        }
        @keyframes indeterminate {
            0% { width: 0%; margin-left: 0%; }
            50% { width: 60%; margin-left: 20%; }
            100% { width: 0%; margin-left: 100%; }
        }

        /* Success state */
        .success { display: none; padding: 24px 0; }
        .success.active { display: block; }

        .success-icon {
            font-size: 3.5rem; margin-bottom: 16px; display: block;
            animation: popIn 0.4s cubic-bezier(0.175, 0.885, 0.32, 1.275);
        }
        @keyframes popIn {
            0% { transform: scale(0); opacity: 0; }
            100% { transform: scale(1); opacity: 1; }
        }

        .success-text { font-size: 1.05rem; font-weight: 700; color: #4ade80; margin-bottom: 6px; }
        .success-file { font-size: 0.78rem; color: rgba(255,255,255,0.35); margin-bottom: 24px; word-break: break-all; }

        .btn {
            display: inline-flex; align-items: center; gap: 8px;
            padding: 14px 32px; border: none; border-radius: 12px;
            font-family: 'Inter', sans-serif; font-size: 0.95rem;
            font-weight: 700; cursor: pointer; transition: all 0.25s ease;
            text-decoration: none;
        }
        .btn-primary {
            background: linear-gradient(135deg, #667eea, #764ba2);
            color: white;
            box-shadow: 0 4px 15px rgba(102,126,234,0.4);
        }
        .btn-primary:hover {
            transform: translateY(-2px);
            box-shadow: 0 8px 25px rgba(102,126,234,0.5);
        }
        .btn-secondary {
            background: rgba(255,255,255,0.08);
            color: rgba(255,255,255,0.7);
            border: 1px solid rgba(255,255,255,0.12);
            margin-top: 12px; font-size: 0.82rem; padding: 10px 24px;
        }
        .btn-secondary:hover {
            background: rgba(255,255,255,0.12); color: white;
        }

        /* Error state */
        .error { display: none; padding: 24px 0; }
        .error.active { display: block; }
        .error-icon { font-size: 3rem; margin-bottom: 12px; display: block; }
        .error-text { font-size: 0.95rem; font-weight: 600; color: #f87171; margin-bottom: 6px; }
        .error-detail { font-size: 0.78rem; color: rgba(255,255,255,0.35); margin-bottom: 20px; }

        /* Specs bar */
        .specs {
            margin-top: 24px; display: flex; flex-wrap: wrap;
            gap: 8px; justify-content: center;
        }
        .spec-tag {
            background: rgba(255,255,255,0.06);
            border: 1px solid rgba(255,255,255,0.08);
            border-radius: 8px; padding: 5px 12px;
            font-size: 0.7rem; color: rgba(255,255,255,0.4);
        }

        .footer {
            position: fixed; bottom: 16px;
            font-size: 0.7rem; color: rgba(255,255,255,0.15); z-index: 1;
        }
    </style>
</head>
<body>

    <div class="bg-blob"></div>
    <div class="bg-blob"></div>
    <div class="bg-blob"></div>

    <div class="card">
        <div class="title">Căn Chỉnh Báo Cáo</div>
        <div class="subtitle">Kéo thả file Word vào để tự động căn chỉnh theo chuẩn</div>

        <!-- DROP ZONE -->
        <div class="drop-zone" id="dropZone">
            <span class="drop-icon">📄</span>
            <div class="drop-text">Kéo thả file .docx vào đây</div>
            <div class="drop-hint">hoặc nhấn để chọn file</div>
        </div>
        <input type="file" id="fileInput" accept=".docx">

        <!-- SETTINGS TOGGLE -->
        <div class="settings-toggle" id="settingsToggle" onclick="toggleSettings()">
            ⚙️ Tùy chỉnh thông số <span class="arrow">▼</span>
        </div>

        <!-- SETTINGS PANEL -->
        <div class="settings-panel" id="settingsPanel">
            <div class="settings-grid">
                <div class="setting-item">
                    <label>Font chữ</label>
                    <select id="s_font">
                        <option value="Times New Roman" selected>Times New Roman</option>
                        <option value="Arial">Arial</option>
                        <option value="Calibri">Calibri</option>
                        <option value="Tahoma">Tahoma</option>
                    </select>
                </div>
                <div class="setting-item">
                    <label>Giãn dòng</label>
                    <input type="number" id="s_line_spacing" value="1.3" step="0.05" min="1" max="3">
                </div>
                <div class="setting-item">
                    <label>Cỡ chữ nội dung (pt)</label>
                    <input type="number" id="s_body_size" value="13" step="0.5" min="8" max="28">
                </div>
                <div class="setting-item">
                    <label>Cỡ chữ Heading 1 (pt)</label>
                    <input type="number" id="s_h1_size" value="16" step="0.5" min="10" max="36">
                </div>
                <div class="setting-item">
                    <label>Thụt đầu dòng (cm)</label>
                    <input type="number" id="s_first_indent" value="1.0" step="0.1" min="0" max="5">
                </div>

                <div class="settings-divider"></div>

                <div class="setting-item">
                    <label>Lề trái (mm)</label>
                    <input type="number" id="s_margin_left" value="30" step="1" min="5" max="50">
                </div>
                <div class="setting-item">
                    <label>Lề phải (mm)</label>
                    <input type="number" id="s_margin_right" value="15" step="1" min="5" max="50">
                </div>
                <div class="setting-item">
                    <label>Lề trên (mm)</label>
                    <input type="number" id="s_margin_top" value="20" step="1" min="5" max="50">
                </div>
                <div class="setting-item">
                    <label>Lề dưới (mm)</label>
                    <input type="number" id="s_margin_bot" value="20" step="1" min="5" max="50">
                </div>

                <button class="btn-reset" onclick="resetSettings()">↺ Khôi phục mặc định</button>
            </div>
        </div>

        <!-- PROCESSING -->
        <div class="processing" id="processing">
            <div class="spinner"></div>
            <div class="processing-text">Đang căn chỉnh...</div>
            <div class="processing-file" id="processingFile"></div>
            <div class="progress-bar-container">
                <div class="progress-bar"></div>
            </div>
        </div>

        <!-- SUCCESS -->
        <div class="success" id="success">
            <span class="success-icon">✅</span>
            <div class="success-text">Hoàn thành!</div>
            <div class="success-file" id="successFile"></div>
            <a class="btn btn-primary" id="downloadBtn" href="#">
                <span>⬇</span> Tải file đã chỉnh
            </a>
            <br>
            <button class="btn btn-secondary" onclick="resetUI()">
                📄 Chỉnh file khác
            </button>
        </div>

        <!-- ERROR -->
        <div class="error" id="error">
            <span class="error-icon">❌</span>
            <div class="error-text">Có lỗi xảy ra</div>
            <div class="error-detail" id="errorDetail"></div>
            <button class="btn btn-secondary" onclick="resetUI()">
                🔄 Thử lại
            </button>
        </div>

        <!-- SPECS -->
        <div class="specs" id="specsTags">
            <span class="spec-tag" id="specMargin">Lề: 3-1.5-2-2 cm</span>
            <span class="spec-tag" id="specFont">Times New Roman</span>
            <span class="spec-tag" id="specSize">13pt / 16pt</span>
            <span class="spec-tag" id="specSpacing">Giãn dòng 1.3</span>
            <span class="spec-tag">Mục lục tự động</span>
        </div>
    </div>

    <div class="footer">Bot Căn Chỉnh Báo Cáo v2.0</div>

    <script>
        const dropZone    = document.getElementById('dropZone');
        const fileInput   = document.getElementById('fileInput');
        const processing  = document.getElementById('processing');
        const success     = document.getElementById('success');
        const error       = document.getElementById('error');

        // ── Settings toggle ──
        function toggleSettings() {
            const panel  = document.getElementById('settingsPanel');
            const toggle = document.getElementById('settingsToggle');
            panel.classList.toggle('open');
            toggle.classList.toggle('open');
        }

        // ── Reset settings ──
        const DEFAULTS = {
            s_font: 'Times New Roman', s_line_spacing: '1.3',
            s_body_size: '13', s_h1_size: '16', s_first_indent: '1.0',
            s_margin_left: '30', s_margin_right: '15',
            s_margin_top: '20', s_margin_bot: '20',
        };

        function resetSettings() {
            for (const [id, val] of Object.entries(DEFAULTS)) {
                document.getElementById(id).value = val;
            }
            updateSpecTags();
        }

        // ── Update spec tags when settings change ──
        function updateSpecTags() {
            const ml = document.getElementById('s_margin_left').value;
            const mr = document.getElementById('s_margin_right').value;
            const mt = document.getElementById('s_margin_top').value;
            const mb = document.getElementById('s_margin_bot').value;
            document.getElementById('specMargin').textContent =
                `Lề: ${ml/10}-${mr/10}-${mt/10}-${mb/10} cm`;
            document.getElementById('specFont').textContent =
                document.getElementById('s_font').value;
            document.getElementById('specSize').textContent =
                `${document.getElementById('s_body_size').value}pt / ${document.getElementById('s_h1_size').value}pt`;
            document.getElementById('specSpacing').textContent =
                `Giãn dòng ${document.getElementById('s_line_spacing').value}`;
        }

        // Listen for settings changes
        document.querySelectorAll('.settings-grid input, .settings-grid select')
            .forEach(el => el.addEventListener('change', updateSpecTags));

        // ── Click to browse ──
        dropZone.addEventListener('click', () => fileInput.click());
        fileInput.addEventListener('change', (e) => {
            if (e.target.files.length) uploadFile(e.target.files[0]);
        });

        // ── Drag & Drop ──
        ['dragenter', 'dragover'].forEach(evt => {
            dropZone.addEventListener(evt, (e) => {
                e.preventDefault();
                dropZone.classList.add('drag-over');
            });
        });
        ['dragleave', 'drop'].forEach(evt => {
            dropZone.addEventListener(evt, (e) => {
                e.preventDefault();
                dropZone.classList.remove('drag-over');
            });
        });
        dropZone.addEventListener('drop', (e) => {
            const files = e.dataTransfer.files;
            if (files.length) uploadFile(files[0]);
        });

        // Also allow drop on entire page
        document.body.addEventListener('dragover', (e) => e.preventDefault());
        document.body.addEventListener('drop', (e) => {
            e.preventDefault();
            const files = e.dataTransfer.files;
            if (files.length) uploadFile(files[0]);
        });

        // ── Upload file with settings ──
        function uploadFile(file) {
            if (!file.name.toLowerCase().endsWith('.docx')) {
                showError('Chỉ hỗ trợ file .docx', 'Vui lòng chọn file Word (.docx)');
                return;
            }

            // Show processing
            dropZone.style.display = 'none';
            document.getElementById('settingsToggle').style.display = 'none';
            document.getElementById('settingsPanel').classList.remove('open');
            processing.classList.add('active');
            success.classList.remove('active');
            error.classList.remove('active');
            document.getElementById('processingFile').textContent = file.name;

            // Build FormData with file + settings
            const formData = new FormData();
            formData.append('file', file);
            formData.append('font_name',    document.getElementById('s_font').value);
            formData.append('line_spacing',  document.getElementById('s_line_spacing').value);
            formData.append('body_size',     document.getElementById('s_body_size').value);
            formData.append('h1_size',       document.getElementById('s_h1_size').value);
            formData.append('first_indent',  document.getElementById('s_first_indent').value);
            formData.append('margin_left',   document.getElementById('s_margin_left').value);
            formData.append('margin_right',  document.getElementById('s_margin_right').value);
            formData.append('margin_top',    document.getElementById('s_margin_top').value);
            formData.append('margin_bot',    document.getElementById('s_margin_bot').value);

            fetch('/upload', { method: 'POST', body: formData })
            .then(response => {
                if (!response.ok) {
                    return response.json().then(data => { throw new Error(data.error || 'Lỗi server'); });
                }
                const cd = response.headers.get('Content-Disposition');
                let filename = file.name.replace('.docx', ' - DA CHINH SUA.docx');
                if (cd) {
                    const match = cd.match(/filename\*=UTF-8''(.+)/);
                    if (match) filename = decodeURIComponent(match[1]);
                    else {
                        const match2 = cd.match(/filename="?(.+?)"?$/);
                        if (match2) filename = match2[1];
                    }
                }
                return response.blob().then(blob => ({ blob, filename }));
            })
            .then(({ blob, filename }) => {
                const url = URL.createObjectURL(blob);
                const downloadBtn = document.getElementById('downloadBtn');
                downloadBtn.href = url;
                downloadBtn.download = filename;

                processing.classList.remove('active');
                success.classList.add('active');
                document.getElementById('successFile').textContent = filename;

                // Auto-download
                downloadBtn.click();
            })
            .catch(err => {
                showError('Không thể xử lý file', err.message);
            });
        }

        function showError(title, detail) {
            dropZone.style.display = 'none';
            document.getElementById('settingsToggle').style.display = 'none';
            processing.classList.remove('active');
            success.classList.remove('active');
            error.classList.add('active');
            document.querySelector('.error-text').textContent = title;
            document.getElementById('errorDetail').textContent = detail;
        }

        function resetUI() {
            dropZone.style.display = '';
            document.getElementById('settingsToggle').style.display = '';
            processing.classList.remove('active');
            success.classList.remove('active');
            error.classList.remove('active');
            fileInput.value = '';
        }
    </script>
</body>
</html>
"""


# ════════════════════════════════════════════════════════════
# FLASK ROUTES
# ════════════════════════════════════════════════════════════

@app.route('/')
def index():
    return render_template_string(HTML_PAGE)


@app.route('/upload', methods=['POST'])
def upload():
    if 'file' not in request.files:
        return jsonify({'error': 'Không tìm thấy file'}), 400

    file = request.files['file']
    if not file.filename or not file.filename.lower().endswith('.docx'):
        return jsonify({'error': 'Chỉ hỗ trợ file .docx'}), 400

    # Lưu file upload vào thư mục tạm
    job_id = str(uuid.uuid4())[:8]
    job_dir = os.path.join(TEMP_DIR, job_id)
    os.makedirs(job_dir, exist_ok=True)

    input_path = os.path.join(job_dir, file.filename)
    file.save(input_path)

    # Tạo tên output
    stem = Path(file.filename).stem
    out_name = f"{stem} - DA CHINH SUA.docx"
    output_path = os.path.join(job_dir, out_name)

    # Lấy settings từ form
    settings = {}
    for key in ['font_name', 'body_size', 'h1_size', 'line_spacing',
                'first_indent', 'margin_left', 'margin_right',
                'margin_top', 'margin_bot']:
        val = request.form.get(key)
        if val is not None and val != '':
            settings[key] = val

    try:
        format_document(input_path, output_path, log=print, settings=settings)
    except Exception as e:
        return jsonify({'error': str(e)}), 500

    return send_file(
        output_path,
        as_attachment=True,
        download_name=out_name,
        mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    )


# ════════════════════════════════════════════════════════════
# ENTRY POINT
# ════════════════════════════════════════════════════════════

def open_browser():
    import time
    time.sleep(1.5)
    webbrowser.open('http://localhost:5000')


if __name__ == '__main__':
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8')

    print()
    print("  Bot Can Chinh Bao Cao - Web App v2.0")
    print("  ------------------------------------")
    print("  Dang mo trinh duyet...")
    print("  Neu khong tu mo, hay vao: http://localhost:5000")
    print("  Nhan Ctrl+C de tat server")
    print()

    threading.Thread(target=open_browser, daemon=True).start()

    app.run(host='0.0.0.0', port=5000, debug=False)
