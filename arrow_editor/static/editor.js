/**
 * Arrow Editor — Interactive Canvas Logic (Multi-View)
 * =====================================================
 * Handles image display, click-to-place arrow endpoints,
 * live preview, and save/load operations.
 *
 * Supports multiple view types (side_view, front_view, top_view)
 * with per-view arrow configurations. Dynamically resolves the
 * correct image based on the master car type + selected view.
 */

// ─── State ──────────────────────────────────────────────────────
let metrics = [];
let selectedMetric = null;
let currentImage = null;
let loadedImage = null;  // HTMLImageElement
let masterCarType = null; // e.g. "Curvy Master"
let masterCarName = null;
let viewTypes = ["side_view", "front_view", "top_view"];
let currentViewType = "side_view";

// Arrow state — stored per view
let arrowStart = null;    // {x: 0-1, y: 0-1} percentage coords
let arrowEnd = null;
let placingMode = "start"; // "start" | "end" | "done"

// Full arrow data for all views of the current metric
let metricArrowData = {};  // { view_type: "side_view", side_view: {...}, front_view: {...} }

// Canvas
let canvas = null;
let ctx = null;

// ─── Initialize ─────────────────────────────────────────────────
document.addEventListener("DOMContentLoaded", () => {
    loadMasterCarInfo();
    loadMetrics();

    // Search filter
    document.getElementById("searchBox").addEventListener("input", (e) => {
        filterMetrics(e.target.value);
    });
});


// ─── API Calls ──────────────────────────────────────────────────
async function loadMasterCarInfo() {
    try {
        const resp = await fetch("/api/master-car");
        const data = await resp.json();
        masterCarName = data.name;
        masterCarType = data.type;
        updateMasterCarBadge();
    } catch (err) {
        console.error("Failed to load master car info:", err);
    }
}

async function loadMetrics() {
    try {
        const resp = await fetch("/api/metrics");
        metrics = await resp.json();
        renderMetricList();
        updateStats();
    } catch (err) {
        showToast("Failed to load metrics", "error");
    }
}

async function loadArrowConfig(metricName) {
    try {
        const resp = await fetch(`/api/arrow/${encodeURIComponent(metricName)}`);
        const data = await resp.json();
        return data;
    } catch (err) {
        return null;
    }
}

async function saveArrowConfig(metricName, config) {
    try {
        const resp = await fetch(`/api/arrow/${encodeURIComponent(metricName)}`, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify(config),
        });
        const data = await resp.json();
        if (data.status === "saved") {
            showToast(`Saved arrow for "${metricName}" (${currentViewType})`, "success");
            // Update metric status
            const m = metrics.find(m => m.name === metricName);
            if (m) {
                m.configured = true;
                if (!m.configured_views) m.configured_views = [];
                if (!m.configured_views.includes(currentViewType)) {
                    m.configured_views.push(currentViewType);
                }
                m.view_type = config.view_type || currentViewType;
            }
            renderMetricList();
            updateStats();
        }
    } catch (err) {
        showToast("Failed to save", "error");
    }
}

async function deleteArrowConfig(metricName, viewOnly = false) {
    try {
        let url = `/api/arrow/${encodeURIComponent(metricName)}`;
        if (viewOnly) {
            url += `?view=${currentViewType}`;
        }
        const resp = await fetch(url, { method: "DELETE" });
        const data = await resp.json();
        if (data.status === "deleted") {
            const deleteMsg = viewOnly
                ? `Deleted ${currentViewType} arrow for "${metricName}"`
                : `Deleted all arrows for "${metricName}"`;
            showToast(deleteMsg, "success");

            const m = metrics.find(m => m.name === metricName);
            if (m) {
                if (viewOnly) {
                    m.configured_views = (m.configured_views || []).filter(v => v !== currentViewType);
                    if (m.configured_views.length === 0) m.configured = false;
                } else {
                    m.configured = false;
                    m.configured_views = [];
                }
            }
            renderMetricList();
            updateStats();
            resetArrow();
            // Remove from local data
            if (viewOnly) {
                delete metricArrowData[currentViewType];
            } else {
                metricArrowData = {};
            }
            drawCanvas();
        }
    } catch (err) {
        showToast("Failed to delete", "error");
    }
}

async function resolveImage(carType, viewType) {
    try {
        const resp = await fetch(`/api/resolve-image?car_type=${encodeURIComponent(carType || "")}&view_type=${encodeURIComponent(viewType)}`);
        const data = await resp.json();
        return data.image;
    } catch (err) {
        return "side_view.png";
    }
}


// ─── Metric List ────────────────────────────────────────────────
function renderMetricList(filter = "") {
    const list = document.getElementById("metricList");
    list.innerHTML = "";

    const filterLower = filter.toLowerCase();
    const filtered = metrics.filter(m =>
        m.name.toLowerCase().includes(filterLower) ||
        m.display_name.toLowerCase().includes(filterLower)
    );

    filtered.forEach(m => {
        const item = document.createElement("div");
        item.className = `metric-item ${selectedMetric && selectedMetric.name === m.name ? "active" : ""}`;

        // Build view indicator dots
        const viewDots = viewTypes.map(vt => {
            const isConfigured = (m.configured_views || []).includes(vt);
            const initial = vt === "side_view" ? "S" : vt === "front_view" ? "F" : "T";
            return `<span class="view-dot ${isConfigured ? 'configured' : 'unconfigured'}" title="${vt}">${initial}</span>`;
        }).join("");

        item.innerHTML = `
            <span class="status-dot ${m.configured ? "configured" : "unconfigured"}"></span>
            <div class="metric-info">
                <div class="metric-name" title="${m.name}">${m.display_name}</div>
                <div class="metric-unit">${m.unit ? `Unit: ${m.unit}` : ""}</div>
                <div class="metric-views">${viewDots}</div>
            </div>
        `;
        item.addEventListener("click", () => selectMetric(m));
        list.appendChild(item);
    });
}

function filterMetrics(query) {
    renderMetricList(query);
}

function updateStats() {
    const configured = metrics.filter(m => m.configured).length;
    document.getElementById("stats-configured").textContent =
        `Configured: ${configured}/${metrics.length}`;
}

function updateMasterCarBadge() {
    const badge = document.getElementById("masterCarBadge");
    if (badge && masterCarName) {
        badge.textContent = `Master: ${masterCarName} (${masterCarType || "N/A"})`;
    } else if (badge) {
        badge.textContent = "No master car";
    }
}


// ─── Select Metric ──────────────────────────────────────────────
async function selectMetric(metric) {
    selectedMetric = metric;
    currentViewType = metric.view_type || "side_view";
    renderMetricList(document.getElementById("searchBox").value);

    // Build editor UI
    const editor = document.getElementById("editorArea");
    document.getElementById("emptyState")?.remove();

    // Build view type options HTML
    const viewTypeOptions = viewTypes.map(vt => {
        const label = vt.replace("_", " ").replace(/\b\w/g, c => c.toUpperCase());
        const selected = vt === currentViewType ? "selected" : "";
        return `<option value="${vt}" ${selected}>${label}</option>`;
    }).join("");

    editor.innerHTML = `
        <div class="editor-title">
            <h2>${metric.display_name}</h2>
            <span class="metric-badge">${metric.unit || "no unit"}</span>
        </div>

        <div class="canvas-container" id="canvasContainer">
            <div class="canvas-instructions" id="canvasInstructions">
                Click to place START point (green)
            </div>
            <canvas id="editorCanvas"></canvas>
        </div>

        <div class="coords-display" id="coordsDisplay">
            <div class="coord-item">
                <span class="coord-dot start"></span>
                <span>Start: </span>
                <span class="coord-value" id="coordStart">Not set</span>
            </div>
            <div class="coord-item">
                <span class="coord-dot end"></span>
                <span>End: </span>
                <span class="coord-value" id="coordEnd">Not set</span>
            </div>
        </div>

        <div class="controls-panel">
            <div class="controls-grid">
                <div class="control-group">
                    <label>View Type</label>
                    <select id="viewTypeSelect">${viewTypeOptions}</select>
                </div>
                <div class="control-group">
                    <label>Arrow Label</label>
                    <input type="text" id="arrowLabel" value="${metric.display_name}" placeholder="Label text...">
                </div>
                <div class="control-group">
                    <label>Arrow Color</label>
                    <input type="color" id="arrowColor" value="#E31937">
                </div>
                <div class="btn-group">
                    <button class="btn btn-save" id="btnSave" disabled onclick="onSave()">Save</button>
                    <button class="btn btn-reset" onclick="onReset()">Reset</button>
                    <button class="btn btn-delete" onclick="onDelete(false)">Delete All</button>
                    <button class="btn btn-delete-view" onclick="onDelete(true)">Delete View</button>
                </div>
            </div>
        </div>

        <div class="view-status-panel" id="viewStatusPanel">
            <!-- Populated by JS -->
        </div>
    `;

    // View type change handler
    document.getElementById("viewTypeSelect").addEventListener("change", async (e) => {
        // Save current arrow state to local data before switching
        saveCurrentViewToLocal();
        currentViewType = e.target.value;
        // Load arrow data for the new view
        loadViewFromLocal();
        // Resolve and load the correct image for this view + master type
        await loadImageForCurrentView();
    });

    // Color change
    document.getElementById("arrowColor").addEventListener("input", () => drawCanvas());

    // Load full arrow data for this metric
    metricArrowData = {};
    resetArrow();

    const existing = await loadArrowConfig(metric.name);
    if (existing) {
        metricArrowData = existing;
        currentViewType = existing.view_type || "side_view";
        document.getElementById("viewTypeSelect").value = currentViewType;
        loadViewFromLocal();
    }

    updateViewStatusPanel();

    // Load image for current view + master type
    await loadImageForCurrentView();
}

function saveCurrentViewToLocal() {
    if (arrowStart && arrowEnd) {
        metricArrowData[currentViewType] = {
            start: [arrowStart.x, arrowStart.y],
            end: [arrowEnd.x, arrowEnd.y],
            label: document.getElementById("arrowLabel")?.value || selectedMetric?.display_name || "",
            color: document.getElementById("arrowColor")?.value || "#E31937",
        };
    }
}

function loadViewFromLocal() {
    const viewData = metricArrowData[currentViewType];
    if (viewData) {
        arrowStart = { x: viewData.start[0], y: viewData.start[1] };
        arrowEnd = { x: viewData.end[0], y: viewData.end[1] };
        placingMode = "done";
        document.getElementById("arrowLabel").value = viewData.label || selectedMetric?.display_name || "";
        document.getElementById("arrowColor").value = viewData.color || "#E31937";
        document.getElementById("btnSave").disabled = false;
    } else {
        resetArrow();
    }
    updateCoordDisplay();
    updateInstructions();
    updateViewStatusPanel();
}

async function loadImageForCurrentView() {
    // Resolve the correct image for master type + view
    const imageFile = await resolveImage(masterCarType, currentViewType);
    currentImage = imageFile;
    loadImageAndSetupCanvas();
}

function updateViewStatusPanel() {
    const panel = document.getElementById("viewStatusPanel");
    if (!panel) return;

    const items = viewTypes.map(vt => {
        const hasData = vt in metricArrowData;
        const isActive = vt === currentViewType;
        const isPptView = metricArrowData.view_type === vt;
        const label = vt.replace("_", " ").replace(/\b\w/g, c => c.toUpperCase());
        return `
            <div class="view-status-item ${isActive ? 'active' : ''} ${hasData ? 'has-data' : ''}">
                <span class="view-status-dot ${hasData ? 'configured' : 'unconfigured'}"></span>
                <span class="view-status-label">${label}</span>
                ${isPptView ? '<span class="ppt-badge">PPT</span>' : ''}
            </div>
        `;
    }).join("");

    panel.innerHTML = `
        <div class="view-status-title">View Configurations</div>
        <div class="view-status-list">${items}</div>
    `;
}

function resetArrow() {
    arrowStart = null;
    arrowEnd = null;
    placingMode = "start";
    updateInstructions();
    updateCoordDisplay();

    const saveBtn = document.getElementById("btnSave");
    if (saveBtn) saveBtn.disabled = true;
}

function loadImageAndSetupCanvas() {
    const container = document.getElementById("canvasContainer");
    canvas = document.getElementById("editorCanvas");
    if (!canvas || !container) return;
    ctx = canvas.getContext("2d");

    const img = new Image();
    img.crossOrigin = "anonymous";
    img.onload = () => {
        loadedImage = img;

        // Set canvas size to match image aspect ratio within container
        // Respect both container width and max-height (55vh)
        const containerWidth = container.clientWidth;
        const maxHeight = window.innerHeight * 0.55;
        const aspectRatio = img.width / img.height;

        let displayWidth = containerWidth;
        let displayHeight = containerWidth / aspectRatio;

        // If the calculated height exceeds max, scale down to fit
        if (displayHeight > maxHeight) {
            displayHeight = maxHeight;
            displayWidth = maxHeight * aspectRatio;
        }

        canvas.width = img.width;
        canvas.height = img.height;
        canvas.style.width = displayWidth + "px";
        canvas.style.height = displayHeight + "px";

        drawCanvas();
    };
    img.onerror = () => {
        showToast("Failed to load image: " + currentImage, "error");
    };
    img.src = `/images/${currentImage}`;

    // Click handler
    canvas.onclick = onCanvasClick;
}


// ─── Canvas Drawing ─────────────────────────────────────────────
function drawCanvas() {
    if (!ctx || !loadedImage) return;

    const w = canvas.width;
    const h = canvas.height;

    // Clear and draw image
    ctx.clearRect(0, 0, w, h);
    ctx.drawImage(loadedImage, 0, 0, w, h);

    const color = document.getElementById("arrowColor")?.value || "#E31937";
    const label = document.getElementById("arrowLabel")?.value || "";

    // Draw start point
    if (arrowStart) {
        const sx = arrowStart.x * w;
        const sy = arrowStart.y * h;
        drawDot(sx, sy, "#27ae60", 10);

        // Draw arrow if both points set
        if (arrowEnd) {
            const ex = arrowEnd.x * w;
            const ey = arrowEnd.y * h;
            drawDot(ex, ey, "#e74c3c", 10);
            drawArrow(sx, sy, ex, ey, color, label);
        }
    }

    // Draw view type indicator on canvas
    drawViewIndicator();
}

function drawViewIndicator() {
    if (!ctx) return;
    const label = currentViewType.replace("_", " ").replace(/\b\w/g, c => c.toUpperCase());

    ctx.save();
    ctx.font = `bold ${Math.max(14, canvas.height * 0.02)}px Inter, Arial, sans-serif`;
    const metrics = ctx.measureText(label);
    const pad = 8;
    const x = canvas.width - metrics.width - pad * 2 - 10;
    const y = 10;

    ctx.fillStyle = "rgba(0, 150, 150, 0.85)";
    ctx.roundRect(x, y, metrics.width + pad * 2, 28, 6);
    ctx.fill();

    ctx.fillStyle = "white";
    ctx.textAlign = "left";
    ctx.textBaseline = "middle";
    ctx.fillText(label, x + pad, y + 14);
    ctx.restore();
}

function drawDot(x, y, color, radius) {
    ctx.beginPath();
    ctx.arc(x, y, radius, 0, Math.PI * 2);
    ctx.fillStyle = color;
    ctx.fill();
    ctx.strokeStyle = "white";
    ctx.lineWidth = 2;
    ctx.stroke();
}

function drawArrow(x1, y1, x2, y2, color, label) {
    const headLen = 20;
    const angle = Math.atan2(y2 - y1, x2 - x1);

    // Main line
    ctx.beginPath();
    ctx.moveTo(x1, y1);
    ctx.lineTo(x2, y2);
    ctx.strokeStyle = color;
    ctx.lineWidth = 3;
    ctx.stroke();

    // Arrowhead at end
    drawArrowhead(x2, y2, angle, headLen, color);
    // Arrowhead at start (reversed)
    drawArrowhead(x1, y1, angle + Math.PI, headLen, color);

    // Label
    if (label) {
        const midX = (x1 + x2) / 2;
        const midY = (y1 + y2) / 2;

        ctx.font = `bold ${Math.max(16, canvas.height * 0.025)}px Inter, Arial, sans-serif`;
        const textMetrics = ctx.measureText(label);
        const textH = 20;
        const pad = 6;

        // Background
        ctx.fillStyle = "rgba(255,255,255,0.85)";
        ctx.fillRect(
            midX - textMetrics.width / 2 - pad,
            midY - textH - pad,
            textMetrics.width + pad * 2,
            textH + pad * 2
        );

        // Text
        ctx.fillStyle = color;
        ctx.textAlign = "center";
        ctx.textBaseline = "middle";
        ctx.fillText(label, midX, midY);
    }
}

function drawArrowhead(tipX, tipY, angle, size, color) {
    const leftAngle = angle + Math.PI * 5 / 6;
    const rightAngle = angle - Math.PI * 5 / 6;

    ctx.beginPath();
    ctx.moveTo(tipX, tipY);
    ctx.lineTo(tipX + size * Math.cos(leftAngle), tipY + size * Math.sin(leftAngle));
    ctx.lineTo(tipX + size * Math.cos(rightAngle), tipY + size * Math.sin(rightAngle));
    ctx.closePath();
    ctx.fillStyle = color;
    ctx.fill();
}


// ─── Canvas Click ───────────────────────────────────────────────
function onCanvasClick(e) {
    const rect = canvas.getBoundingClientRect();
    const scaleX = canvas.width / rect.width;
    const scaleY = canvas.height / rect.height;

    const pixelX = (e.clientX - rect.left) * scaleX;
    const pixelY = (e.clientY - rect.top) * scaleY;

    // Convert to percentage
    const pctX = pixelX / canvas.width;
    const pctY = pixelY / canvas.height;

    if (placingMode === "start") {
        arrowStart = { x: pctX, y: pctY };
        arrowEnd = null;
        placingMode = "end";
        document.getElementById("btnSave").disabled = true;
    } else if (placingMode === "end") {
        arrowEnd = { x: pctX, y: pctY };
        placingMode = "done";
        document.getElementById("btnSave").disabled = false;
    } else {
        // Already done — re-start placing
        arrowStart = { x: pctX, y: pctY };
        arrowEnd = null;
        placingMode = "end";
        document.getElementById("btnSave").disabled = true;
    }

    updateCoordDisplay();
    updateInstructions();
    drawCanvas();
}


// ─── UI Updates ─────────────────────────────────────────────────
function updateCoordDisplay() {
    const startEl = document.getElementById("coordStart");
    const endEl = document.getElementById("coordEnd");

    if (startEl) {
        startEl.textContent = arrowStart
            ? `(${arrowStart.x.toFixed(3)}, ${arrowStart.y.toFixed(3)})`
            : "Not set";
    }
    if (endEl) {
        endEl.textContent = arrowEnd
            ? `(${arrowEnd.x.toFixed(3)}, ${arrowEnd.y.toFixed(3)})`
            : "Not set";
    }
}

function updateInstructions() {
    const el = document.getElementById("canvasInstructions");
    if (!el) return;

    const viewLabel = currentViewType.replace("_", " ").replace(/\b\w/g, c => c.toUpperCase());

    if (placingMode === "start") {
        el.textContent = `[${viewLabel}] Click to place START point (green)`;
        el.style.opacity = "1";
    } else if (placingMode === "end") {
        el.textContent = `[${viewLabel}] Click to place END point (red)`;
        el.style.opacity = "1";
    } else {
        el.textContent = `[${viewLabel}] Arrow set! Click again to re-place start point`;
        el.style.opacity = "0.6";
    }
}


// ─── Button Handlers ────────────────────────────────────────────
function onSave() {
    if (!selectedMetric || !arrowStart || !arrowEnd) return;

    const config = {
        view_type: currentViewType,  // This view becomes the active PPT view
        view: currentViewType,
        start: [arrowStart.x, arrowStart.y],
        end: [arrowEnd.x, arrowEnd.y],
        label: document.getElementById("arrowLabel").value || selectedMetric.display_name,
        color: document.getElementById("arrowColor").value || "#E31937",
    };

    // Also update local data
    metricArrowData.view_type = currentViewType;
    metricArrowData[currentViewType] = {
        start: config.start,
        end: config.end,
        label: config.label,
        color: config.color,
    };

    saveArrowConfig(selectedMetric.name, config);
    updateViewStatusPanel();
}

function onReset() {
    resetArrow();
    drawCanvas();
}

function onDelete(viewOnly = false) {
    if (!selectedMetric) return;
    const msg = viewOnly
        ? `Delete ${currentViewType} arrow for "${selectedMetric.display_name}"?`
        : `Delete ALL arrow configs for "${selectedMetric.display_name}"?`;
    if (confirm(msg)) {
        deleteArrowConfig(selectedMetric.name, viewOnly);
    }
}


// ─── Toast ──────────────────────────────────────────────────────
function showToast(message, type = "success") {
    const toast = document.getElementById("toast");
    toast.textContent = message;
    toast.className = `toast ${type} show`;
    setTimeout(() => {
        toast.className = "toast";
    }, 2500);
}
