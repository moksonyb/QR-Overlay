import csv
import io
from flask import Flask, render_template_string, request, jsonify, send_file
from PIL import Image
import PyPDF2
import qrcode
import segno
from reportlab.lib.utils import ImageReader
from reportlab.pdfgen import canvas
from reportlab.graphics import renderPDF
from reportlab.graphics.shapes import Drawing
from pdf2image import convert_from_bytes


# Units: positions are in PDF points (1 point = 1/72 inch), origin at bottom-left.

app = Flask(__name__)


def read_csv_column_by_name(file_storage, column_name: str):
    """Return a list of strings from the specified CSV column (by header name).
    
    First row is treated as header. Returns all data rows for the selected column.
    Auto-detects delimiter (comma or semicolon). Empty cells are returned as empty strings.
    """
    data = file_storage.read().decode("utf-8")
    file_storage.seek(0)
    
    # Auto-detect delimiter
    sniffer = csv.Sniffer()
    sample = data[:4096]
    try:
        delimiter = sniffer.sniff(sample).delimiter
    except csv.Error:
        delimiter = ','
    
    reader = csv.reader(io.StringIO(data), delimiter=delimiter)
    all_rows = list(reader)
    
    if not all_rows:
        raise ValueError("CSV is empty")
    
    # Parse header
    header = all_rows[0]
    cleaned_headers = [h.strip() for h in header]
    
    if column_name.strip() not in cleaned_headers:
        raise ValueError(f"Column '{column_name}' not found. Available: {cleaned_headers}")
    
    column_index = cleaned_headers.index(column_name.strip())
    
    # Extract column values from data rows, preserving empty cells as empty strings
    out = []
    for i, row in enumerate(all_rows[1:], start=2):
        if not row:
            out.append("")  # Empty row = empty string
        elif column_index >= len(row):
            out.append("")  # Missing column = empty string
        else:
            val = (row[column_index] or "").strip()
            out.append(val)  # Include even if empty
    
    return out


def make_qr(data, size_pts):
    """Create a QR PIL image sized to size_pts with high quality for sharp rendering."""
    qr = qrcode.QRCode(
        version=None,
        error_correction=qrcode.constants.ERROR_CORRECT_L,
        box_size=20,  # Large box size for better quality
        border=1,
    )
    qr.add_data(data)
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white").convert("RGB")
    # Scale up to size_pts with high-quality resampling
    return img.resize((int(size_pts), int(size_pts)), Image.Resampling.LANCZOS)


def build_overlay_vector(page_width, page_height, qr_data, qr_size, x_pos, y_pos):
    """Create a single-page PDF overlay with a vector QR code.
    
    QR is centered within the specified position and size box.
    """
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=(page_width, page_height))
    
    # Generate vector QR using segno
    qr = segno.make(qr_data, error='l')
    
    # Render as SVG
    svg_buffer = io.BytesIO()
    qr.save(svg_buffer, kind='svg', scale=1, border=0, xmldecl=False, svgns=False)
    svg_buffer.seek(0)
    
    from svglib.svglib import svg2rlg
    drawing = svg2rlg(svg_buffer)
    if drawing:
        # Get natural QR size
        natural_width = drawing.width
        natural_height = drawing.height
        
        # Calculate scale to fit within qr_size box (use min to prevent overflow)
        scale_factor = min(qr_size / natural_width, qr_size / natural_height)
        
        # Actual rendered size
        rendered_width = natural_width * scale_factor
        rendered_height = natural_height * scale_factor
        
        # Center QR within the qr_size box
        offset_x = (qr_size - rendered_width) / 2
        offset_y = (qr_size - rendered_height) / 2
        
        # Apply transformations
        drawing.width = rendered_width
        drawing.height = rendered_height
        drawing.scale(scale_factor, scale_factor)
        
        # Draw at centered position
        renderPDF.draw(drawing, c, x_pos + offset_x, y_pos + offset_y)
    
    c.save()
    buffer.seek(0)
    return buffer


def place_qrs_on_pdf_stream(pdf_stream, csv_rows, qr_size_pts, x_pos, y_pos):
    reader = PyPDF2.PdfReader(pdf_stream)
    writer = PyPDF2.PdfWriter()

    if len(csv_rows) != len(reader.pages):
        raise ValueError("CSV rows and PDF pages count must match.")

    for row_data, page in zip(csv_rows, reader.pages):
        # Only add QR if row_data is not empty
        if row_data and row_data.strip():
            media_box = page.mediabox
            width = float(media_box.width)
            height = float(media_box.height)

            overlay_stream = build_overlay_vector(width, height, row_data, qr_size_pts, x_pos, y_pos)
            overlay_pdf = PyPDF2.PdfReader(overlay_stream)
            overlay_page = overlay_pdf.pages[0]

            page.merge_page(overlay_page)
        
        # Always add the page (with or without QR)
        writer.add_page(page)

    out_buf = io.BytesIO()
    writer.write(out_buf)
    out_buf.seek(0)
    return out_buf


INDEX_HTML = """
<!doctype html>
<html>
  <head>
    <meta charset="utf-8" />
    <title>QR Bulk Adder (Web)</title>
    <style>
      body { font-family: system-ui, sans-serif; margin: 20px; max-width: 1200px; }
      .container { display: grid; grid-template-columns: 350px 1fr; gap: 20px; }
      .controls { }
      .preview-section { }
      .row { margin-bottom: 10px; }
      label { display: inline-block; width: 160px; font-weight: 500; }
      input[type="file"], input[type="text"], input[type="number"], select {
        padding: 4px 8px;
        border: 1px solid #ccc;
        border-radius: 4px;
      }
      select { min-width: 200px; }
      .actions { margin-top: 16px; }
      button {
        padding: 8px 16px;
        margin-right: 8px;
        cursor: pointer;
        background: #1e66f5;
        color: white;
        border: none;
        border-radius: 4px;
      }
      button:hover { background: #1557d6; }
      .preview-canvas-container {
        position: relative;
        display: inline-block;
        border: 1px solid #888;
      }
      #previewCanvas {
        display: block;
        max-width: 100%;
      }
      .page-controls {
        margin-bottom: 10px;
        display: flex;
        align-items: center;
        gap: 10px;
      }
      .page-controls button {
        padding: 4px 12px;
      }
      .page-info {
        font-size: 14px;
        color: #666;
      }
      .page-selector {
        padding: 4px 8px;
        border: 1px solid #ccc;
        border-radius: 4px;
        font-size: 14px;
      }
    </style>
  </head>
  <body>
    <h2>QR Bulk Adder</h2>
    <div class="container">
      <div class="controls">
        <form id="form" enctype="multipart/form-data">
          <div class="row">
            <label>PDF file:</label>
            <input type="file" name="pdf" id="pdf" accept="application/pdf" required />
          </div>
          <div class="row">
            <label>CSV file:</label>
            <input type="file" name="csv" id="csv" accept="text/csv" required />
          </div>
          <div class="row">
            <label>CSV column:</label>
            <select name="csv_column" id="csv_column" required>
              <option value="">-- Select a column --</option>
            </select>
          </div>
          <div class="row">
            <label>QR size (pt):</label>
            <input type="number" name="qr_size" id="qr_size" value="80" min="10" />
          </div>
          <div class="row">
            <label>X position (pt):</label>
            <input type="number" name="x_pos" id="x_pos" value="36" />
          </div>
          <div class="row">
            <label>Y position (pt):</label>
            <input type="number" name="y_pos" id="y_pos" value="36" />
          </div>
          <div class="actions">
            <button type="button" id="previewBtn">Preview</button>
            <button type="button" id="generateBtn">Generate</button>
          </div>
        </form>
      </div>
      
      <div class="preview-section">
        <h3>Preview</h3>
        <div class="page-controls">
          <button id="prevPageBtn" disabled>&larr; Previous</button>
          <select id="pageSelector" class="page-selector">
            <option value="1">Page 1</option>
          </select>
          <span class="page-info" id="pageInfo">of 1</span>
          <button id="nextPageBtn" disabled>Next &rarr;</button>
        </div>
        <div class="preview-canvas-container">
          <canvas id="previewCanvas" width="800" height="800"></canvas>
        </div>
      </div>
    </div>

    <script>
      const pdfInput = document.getElementById('pdf');
      const csvInput = document.getElementById('csv');
      const columnSelect = document.getElementById('csv_column');
      const qrSizeInput = document.getElementById('qr_size');
      const xPosInput = document.getElementById('x_pos');
      const yPosInput = document.getElementById('y_pos');
      const canvas = document.getElementById('previewCanvas');
      const ctx = canvas.getContext('2d');
      const pageInfo = document.getElementById('pageInfo');
      const prevPageBtn = document.getElementById('prevPageBtn');
      const nextPageBtn = document.getElementById('nextPageBtn');
      const pageSelector = document.getElementById('pageSelector');
      
      let currentPage = 1;
      let totalPages = 1;
      let pdfWidth = 612;
      let pdfHeight = 792;

      // Populate column selector when CSV is uploaded
      csvInput.addEventListener('change', async (e) => {
        const file = e.target.files[0];
        columnSelect.innerHTML = '<option value="">-- Select a column --</option>';
        if (!file) return;
        
        const fd = new FormData();
        fd.append('csv', file);
        try {
          const res = await fetch('/csv_headers', { method: 'POST', body: fd });
          if (!res.ok) {
            alert('Failed to read CSV headers');
            return;
          }
          const data = await res.json();
          if (Array.isArray(data.headers)) {
            for (const h of data.headers) {
              const opt = document.createElement('option');
              opt.value = h;
              opt.textContent = h;
              columnSelect.appendChild(opt);
            }
          }
        } catch (err) {
          alert('Error reading CSV: ' + err.message);
        }
      });

      async function getPdfInfo() {
        const file = pdfInput.files[0];
        if (!file) throw new Error('Select a PDF');
        const fd = new FormData();
        fd.append('pdf', file);
        const res = await fetch('/preview', { method: 'POST', body: fd });
        if (!res.ok) throw new Error(await res.text());
        return await res.json();
      }

      async function loadPdfPage(pageNum) {
        const file = pdfInput.files[0];
        if (!file) return null;
        
        const fd = new FormData();
        fd.append('pdf', file);
        fd.append('page_num', pageNum);
        const res = await fetch('/pdf_page_image', { method: 'POST', body: fd });
        if (!res.ok) return null;
        
        const blob = await res.blob();
        return new Promise((resolve) => {
          const img = new Image();
          img.onload = () => resolve(img);
          img.src = URL.createObjectURL(blob);
        });
      }

      async function loadQrCode(pageNum) {
        const csvFile = csvInput.files[0];
        const column = columnSelect.value;
        if (!csvFile || !column) return null;
        
        const fd = new FormData();
        fd.append('csv', csvFile);
        fd.append('csv_column', column);
        fd.append('page_num', pageNum);
        fd.append('qr_size', qrSizeInput.value);
        const res = await fetch('/preview_qr', { method: 'POST', body: fd });
        if (!res.ok) return null;
        
        const blob = await res.blob();
        return new Promise((resolve) => {
          const img = new Image();
          img.onload = () => resolve(img);
          img.src = URL.createObjectURL(blob);
        });
      }

      async function renderPreview() {
        try {
          const info = await getPdfInfo();
          pdfWidth = info.width;
          pdfHeight = info.height;
          totalPages = info.page_count;
          
          updatePageControls();
          
          // Clear canvas
          ctx.clearRect(0, 0, canvas.width, canvas.height);
          
          // Calculate scale
          const scale = Math.min(canvas.width / pdfWidth, canvas.height / pdfHeight);
          const scaledWidth = pdfWidth * scale;
          const scaledHeight = pdfHeight * scale;
          
          // Center in canvas
          const offsetX = (canvas.width - scaledWidth) / 2;
          const offsetY = (canvas.height - scaledHeight) / 2;
          
          // Load and draw PDF page
          const pdfImg = await loadPdfPage(currentPage);
          if (pdfImg) {
            ctx.drawImage(pdfImg, offsetX, offsetY, scaledWidth, scaledHeight);
          }
          
          // Draw QR code
          const qrImg = await loadQrCode(currentPage);
          if (qrImg) {
            const qrSize = parseFloat(qrSizeInput.value || '80');
            const x = parseFloat(xPosInput.value || '36');
            const y = parseFloat(yPosInput.value || '36');
            
            const qrScaled = qrSize * scale;
            const xScaled = x * scale;
            const yScaled = y * scale;
            const yFlipped = scaledHeight - yScaled - qrScaled;
            
            ctx.drawImage(qrImg, offsetX + xScaled, offsetY + yFlipped, qrScaled, qrScaled);
          }
        } catch (e) {
          alert(e.message || e);
        }
      }
      function updatePageControls() {
        pageInfo.textContent = `of ${totalPages}`;
        prevPageBtn.disabled = currentPage <= 1;
        nextPageBtn.disabled = currentPage >= totalPages;
        
        // Update page selector dropdown
        pageSelector.innerHTML = '';
        for (let i = 1; i <= totalPages; i++) {
          const opt = document.createElement('option');
          opt.value = i;
          opt.textContent = `Page ${i}`;
          if (i === currentPage) {
            opt.selected = true;
          }
          pageSelector.appendChild(opt);
        }
      }
      
      pageSelector.addEventListener('change', () => {
        currentPage = parseInt(pageSelector.value);
        renderPreview();
      });

      prevPageBtn.addEventListener('click', () => {
        if (currentPage > 1) {
          currentPage--;
          renderPreview();
        }
      });

      nextPageBtn.addEventListener('click', () => {
        if (currentPage < totalPages) {
          currentPage++;
          renderPreview();
        }
      });

      document.getElementById('previewBtn').addEventListener('click', () => {
        currentPage = 1;
        renderPreview();
      });

      // Update preview when QR parameters change
      qrSizeInput.addEventListener('change', () => {
        if (totalPages > 0) renderPreview();
      });
      xPosInput.addEventListener('change', () => {
        if (totalPages > 0) renderPreview();
      });
      yPosInput.addEventListener('change', () => {
        if (totalPages > 0) renderPreview();
      });

      async function generate() {
        const pdf = pdfInput.files[0];
        const csv = csvInput.files[0];
        const column = columnSelect.value;
        
        if (!pdf || !csv) {
          alert('Select both PDF and CSV');
          return;
        }
        if (!column) {
          alert('Select a CSV column');
          return;
        }
        
        const fd = new FormData();
        fd.append('pdf', pdf);
        fd.append('csv', csv);
        fd.append('csv_column', column);
        fd.append('qr_size', qrSizeInput.value);
        fd.append('x_pos', xPosInput.value);
        fd.append('y_pos', yPosInput.value);

        try {
          const res = await fetch('/generate', { method: 'POST', body: fd });
          if (!res.ok) {
            const err = await res.text();
            alert('Error: ' + err);
            return;
          }
          const blob = await res.blob();
          const url = URL.createObjectURL(blob);
          const a = document.createElement('a');
          a.href = url;
          a.download = 'output.pdf';
          document.body.appendChild(a);
          a.click();
          a.remove();
          URL.revokeObjectURL(url);
        } catch (err) {
          alert('Generation failed: ' + err.message);
        }
      }

      document.getElementById('generateBtn').addEventListener('click', generate);
    </script>
  </body>
  </html>
"""


@app.get("/")
def index():
    return render_template_string(INDEX_HTML)


@app.post("/preview")
def preview_endpoint():
    pdf = request.files.get("pdf")
    if not pdf:
        return ("Missing PDF", 400)
    try:
        pdf.seek(0)
        reader = PyPDF2.PdfReader(pdf)
        page_count = len(reader.pages)
        page = reader.pages[0]
        width = float(page.mediabox.width)
        height = float(page.mediabox.height)
        return jsonify({"width": width, "height": height, "page_count": page_count})
    except Exception as exc:
        return (str(exc), 400)


@app.post("/preview_qr")
def preview_qr_endpoint():
    """Generate a QR code preview image at high resolution."""
    csv_file = request.files.get("csv")
    csv_column = request.form.get("csv_column", "").strip()
    page_num = int(request.form.get("page_num", 1))
    qr_size = int(request.form.get("qr_size", 80))
    
    if not csv_file or not csv_column:
        return ("Missing CSV or column", 400)
    
    try:
        csv_rows = read_csv_column_by_name(csv_file, csv_column)
        if page_num < 1 or page_num > len(csv_rows):
            return ("Invalid page number", 400)
        
        qr_data = csv_rows[page_num - 1]
        
        # Return empty response if row is empty
        if not qr_data or not qr_data.strip():
            return ("", 204)
        
        # Generate at 4x size for sharp preview, then let browser scale down
        qr_img = make_qr(qr_data, qr_size * 4)
        
        buf = io.BytesIO()
        qr_img.save(buf, format="PNG")
        buf.seek(0)
        return send_file(buf, mimetype="image/png")
    except Exception as exc:
        return (str(exc), 400)


@app.post("/pdf_page_image")
def pdf_page_image():
    """Render a specific PDF page as an image."""
    pdf = request.files.get("pdf")
    page_num = int(request.form.get("page_num", 1))
    
    if not pdf:
        return ("Missing PDF", 400)
    
    try:
        pdf.seek(0)
        pdf_bytes = pdf.read()
        images = convert_from_bytes(pdf_bytes, first_page=page_num, last_page=page_num, dpi=150)
        
        if not images:
            return ("Failed to render page", 400)
        
        img = images[0]
        buf = io.BytesIO()
        img.save(buf, format="PNG")
        buf.seek(0)
        return send_file(buf, mimetype="image/png")
    except Exception as exc:
        return (str(exc), 400)


@app.post("/generate")
def generate_endpoint():
    pdf = request.files.get("pdf")
    csv_file = request.files.get("csv")
    if not pdf or not csv_file:
        return ("Missing files", 400)

    try:
        csv_column = (request.form.get("csv_column", "") or "").strip()
        if not csv_column:
            return ("CSV column is required", 400)
        
        qr_size = float(request.form.get("qr_size", 80))
        x_pos = float(request.form.get("x_pos", 36))
        y_pos = float(request.form.get("y_pos", 36))
    except ValueError:
        return ("Invalid numeric inputs", 400)

    try:
        csv_rows = read_csv_column_by_name(csv_file, csv_column)
        pdf.seek(0)
        out_buf = place_qrs_on_pdf_stream(pdf, csv_rows, qr_size, x_pos, y_pos)
        return send_file(
            out_buf,
            mimetype="application/pdf",
            as_attachment=True,
            download_name="output.pdf",
        )
    except Exception as exc:
        return (str(exc), 400)


@app.post("/csv_headers")
def csv_headers():
    csv_file = request.files.get("csv")
    if not csv_file:
        return ("Missing CSV", 400)
    try:
        data = csv_file.read().decode("utf-8")
        csv_file.seek(0)
        
        # Auto-detect delimiter
        sniffer = csv.Sniffer()
        sample = data[:4096]
        try:
            delimiter = sniffer.sniff(sample).delimiter
        except csv.Error:
            delimiter = ','
        
        reader = csv.reader(io.StringIO(data), delimiter=delimiter)
        header = next(reader, None)
        headers = [h.strip() for h in header] if header else []
        return jsonify({"headers": headers})
    except Exception as exc:
        return (str(exc), 400)


def main():
    app.run(host="127.0.0.1", port=5000, debug=True)


if __name__ == "__main__":
    main()
