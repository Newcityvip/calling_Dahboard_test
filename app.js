const API_URL = "https://script.google.com/macros/s/AKfycbyBdtI6VnZmsLfo9KT0QfmEuMWREhlKkvZJXl8vyVdXhqxcWqtEUgLKV8SOrev1s3xgXA/exec";

const STATUS_OPTIONS = [
  "Pending",
  "Completed",
  "No Answer",
  "Not Reachable",
  "Call Back Later",
  "Wrong Number",
  "Interested",
  "Not Interested"
];

let currentUser = null;
let currentTasks = [];
let statusChart = null;
let staffChart = null;

const loginView = document.getElementById("loginView");
const adminView = document.getElementById("adminView");
const staffView = document.getElementById("staffView");
const logoutBtn = document.getElementById("logoutBtn");

const loginBtn = document.getElementById("loginBtn");
const staffCodeInput = document.getElementById("staffCode");
const loginMsg = document.getElementById("loginMsg");

const adminName = document.getElementById("adminName");
const processBtn = document.getElementById("processBtn");
const teamGroupSelect = document.getElementById("teamGroup");
const brandNameInput = document.getElementById("brandName");
const excelFileInput = document.getElementById("excelFile");
const uploadMsg = document.getElementById("uploadMsg");

const exportBtn = document.getElementById("exportBtn");
const exportMsg = document.getElementById("exportMsg");

const clearBatchBtn = document.getElementById("clearBatchBtn");
const clearBatchIdInput = document.getElementById("clearBatchId");
const clearMsg = document.getElementById("clearMsg");

const refreshSummaryBtn = document.getElementById("refreshSummaryBtn");
const summaryMsg = document.getElementById("summaryMsg");
const staffPercentGrid = document.getElementById("staffPercentGrid");

const sumTotal = document.getElementById("sumTotal");
const sumPending = document.getElementById("sumPending");
const sumCompleted = document.getElementById("sumCompleted");
const sumUpdated = document.getElementById("sumUpdated");
const sumNoAnswer = document.getElementById("sumNoAnswer");
const sumNotReachable = document.getElementById("sumNotReachable");
const sumCallBackLater = document.getElementById("sumCallBackLater");
const sumWrongNumber = document.getElementById("sumWrongNumber");
const sumInterested = document.getElementById("sumInterested");
const sumNotInterested = document.getElementById("sumNotInterested");

const taskTableBody = document.getElementById("taskTableBody");
const taskCount = document.getElementById("taskCount");
const staffTitle = document.getElementById("staffTitle");
const staffMsg = document.getElementById("staffMsg");
const searchInput = document.getElementById("searchInput");
const statusFilter = document.getElementById("statusFilter");
const refreshTasksBtn = document.getElementById("refreshTasksBtn");

loginBtn.addEventListener("click", handleLogin);
logoutBtn.addEventListener("click", logout);
processBtn.addEventListener("click", handleUploadAndDistribute);
exportBtn.addEventListener("click", handleExport);
clearBatchBtn.addEventListener("click", handleClearBatch);
refreshSummaryBtn.addEventListener("click", loadAdminSummary);
refreshTasksBtn.addEventListener("click", loadStaffTasks);
searchInput.addEventListener("input", renderTasks);
statusFilter.addEventListener("change", renderTasks);

staffCodeInput.addEventListener("keypress", function (e) {
  if (e.key === "Enter") handleLogin();
});

async function postData(payload) {
  const res = await fetch(API_URL, {
    method: "POST",
    headers: {
      "Content-Type": "text/plain;charset=utf-8"
    },
    body: JSON.stringify(payload),
    redirect: "follow"
  });

  const text = await res.text();

  let data;
  try {
    data = JSON.parse(text);
  } catch (e) {
    throw new Error(`Non-JSON response: ${text.slice(0, 300)}`);
  }

  if (!res.ok) {
    throw new Error(data.error || `HTTP ${res.status}`);
  }

  return data;
}

function setMessage(el, text, type = "info") {
  el.className = `message ${type}`;
  el.textContent = text || "";
}

function showView(view) {
  loginView.classList.add("hidden");
  adminView.classList.add("hidden");
  staffView.classList.add("hidden");

  view.classList.remove("hidden");
  logoutBtn.classList.remove("hidden");
}

function logout() {
  currentUser = null;
  currentTasks = [];
  staffCodeInput.value = "";
  loginView.classList.remove("hidden");
  adminView.classList.add("hidden");
  staffView.classList.add("hidden");
  logoutBtn.classList.add("hidden");

  setMessage(loginMsg, "", "info");
  setMessage(uploadMsg, "", "info");
  setMessage(staffMsg, "", "info");
  setMessage(exportMsg, "", "info");
  setMessage(clearMsg, "", "info");
  setMessage(summaryMsg, "", "info");

  clearBatchIdInput.value = "";
  excelFileInput.value = "";
  brandNameInput.value = "";
  staffPercentGrid.innerHTML = `<div class="empty-percent">No data yet.</div>`;

  if (statusChart) {
    statusChart.destroy();
    statusChart = null;
  }
  if (staffChart) {
    staffChart.destroy();
    staffChart = null;
  }
}

async function handleLogin() {
  const code = staffCodeInput.value.trim().toUpperCase();

  if (!code) {
    setMessage(loginMsg, "Please enter your staff code.", "error");
    return;
  }

  loginBtn.disabled = true;
  setMessage(loginMsg, "Logging in...", "info");

  try {
    const res = await postData({
      action: "login",
      code
    });

    if (!res.success) {
      setMessage(loginMsg, "Invalid or inactive staff code.", "error");
      return;
    }

    currentUser = {
      code,
      name: res.name,
      role: res.role,
      group: res.group
    };

    setMessage(loginMsg, "Login successful.", "success");

    if (res.role === "Admin") {
      adminName.textContent = `${res.name} (${code})`;
      showView(adminView);
      await loadAdminSummary();
    } else {
      staffTitle.textContent = `${res.name} - My Assigned Tasks`;
      showView(staffView);
      await loadStaffTasks();
    }
  } catch (err) {
    console.error("Login error:", err);
    setMessage(loginMsg, `Login failed: ${err.message}`, "error");
  } finally {
    loginBtn.disabled = false;
  }
}

async function handleUploadAndDistribute() {
  if (!currentUser || currentUser.role !== "Admin") return;

  const file = excelFileInput.files[0];
  const group = teamGroupSelect.value;
  const brandName = brandNameInput.value.trim();

  if (!brandName) {
    setMessage(uploadMsg, "Please enter the brand name.", "error");
    return;
  }

  if (!file) {
    setMessage(uploadMsg, "Please choose an Excel file first.", "error");
    return;
  }

  processBtn.disabled = true;
  setMessage(uploadMsg, "Reading Excel file...", "info");

  try {
    const rows = await readExcelFile(file);
    const records = mapRowsToRecords(rows);

    if (!records.length) {
      setMessage(uploadMsg, "No valid rows found in the uploaded file.", "error");
      return;
    }

    setMessage(uploadMsg, "Uploading and distributing records...", "info");

    const res = await postData({
      action: "uploadAndDistribute",
      code: currentUser.code,
      group,
      brandName,
      fileName: file.name,
      records
    });

    if (res.success) {
      const batchId = res.batchId || "";
      clearBatchIdInput.value = batchId;

      setMessage(
        uploadMsg,
        `Upload complete.
Brand: ${res.brandName}
Total rows: ${res.total}
After duplicate cleanup: ${res.cleaned}
Duplicates removed: ${res.duplicatesRemoved}
Batch ID: ${batchId}`,
        "success"
      );

      excelFileInput.value = "";
      await loadAdminSummary();
    } else {
      setMessage(uploadMsg, res.error || "Upload failed.", "error");
    }
  } catch (err) {
    console.error("Upload error:", err);
    setMessage(uploadMsg, `Upload failed: ${err.message}`, "error");
  } finally {
    processBtn.disabled = false;
  }
}

async function handleClearBatch() {
  if (!currentUser || currentUser.role !== "Admin") return;

  const batchId = clearBatchIdInput.value.trim();

  if (!batchId) {
    setMessage(clearMsg, "Please enter a Batch ID.", "error");
    return;
  }

  const ok = window.confirm(`Are you sure you want to clear batch: ${batchId} ?`);
  if (!ok) return;

  clearBatchBtn.disabled = true;
  setMessage(clearMsg, "Clearing batch...", "info");

  try {
    const res = await postData({
      action: "clearBatch",
      code: currentUser.code,
      batchId
    });

    if (res.success) {
      setMessage(
        clearMsg,
        `Batch cleared successfully.
Batch ID: ${res.batchId}
Deleted from assigned_tasks: ${res.deleted.assigned_tasks}
Deleted from status_logs: ${res.deleted.status_logs}
Deleted from upload_batches: ${res.deleted.upload_batches}`,
        "success"
      );
      await loadAdminSummary();
    } else {
      setMessage(clearMsg, res.error || "Batch clear failed.", "error");
    }
  } catch (err) {
    console.error("Clear batch error:", err);
    setMessage(clearMsg, `Batch clear failed: ${err.message}`, "error");
  } finally {
    clearBatchBtn.disabled = false;
  }
}

async function loadAdminSummary() {
  if (!currentUser || currentUser.role !== "Admin") return;

  refreshSummaryBtn.disabled = true;
  setMessage(summaryMsg, "Loading live report...", "info");

  try {
    const res = await postData({
      action: "getAdminSummary",
      code: currentUser.code
    });

    if (!res.success) {
      setMessage(summaryMsg, res.error || "Failed to load summary.", "error");
      return;
    }

    const s = res.summary || {};
    sumTotal.textContent = s.total || 0;
    sumPending.textContent = s.pending || 0;
    sumCompleted.textContent = s.completed || 0;
    sumUpdated.textContent = s.updated || 0;
    sumNoAnswer.textContent = s.noAnswer || 0;
    sumNotReachable.textContent = s.notReachable || 0;
    sumCallBackLater.textContent = s.callBackLater || 0;
    sumWrongNumber.textContent = s.wrongNumber || 0;
    sumInterested.textContent = s.interested || 0;
    sumNotInterested.textContent = s.notInterested || 0;

    renderStatusChart(s);
    renderStaffChart(res.staffBreakdown || {});
    renderStaffPercentCards(res.staffBreakdown || {});
    setMessage(summaryMsg, "Live report updated.", "success");
  } catch (err) {
    console.error("Summary error:", err);
    setMessage(summaryMsg, `Failed to load summary: ${err.message}`, "error");
  } finally {
    refreshSummaryBtn.disabled = false;
  }
}

function renderStatusChart(s) {
  const ctx = document.getElementById("statusChart");

  if (statusChart) statusChart.destroy();

  statusChart = new Chart(ctx, {
    type: "doughnut",
    data: {
      labels: [
        "Pending",
        "Completed",
        "No Answer",
        "Not Reachable",
        "Call Back Later",
        "Wrong Number",
        "Interested",
        "Not Interested"
      ],
      datasets: [{
        data: [
          s.pending || 0,
          s.completed || 0,
          s.noAnswer || 0,
          s.notReachable || 0,
          s.callBackLater || 0,
          s.wrongNumber || 0,
          s.interested || 0,
          s.notInterested || 0
        ]
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false
    }
  });
}

function renderStaffChart(staffBreakdown) {
  const ctx = document.getElementById("staffChart");

  if (staffChart) staffChart.destroy();

  const names = Object.keys(staffBreakdown);
  const totals = names.map(name => staffBreakdown[name].total || 0);
  const completed = names.map(name => staffBreakdown[name].completed || 0);

  staffChart = new Chart(ctx, {
    type: "bar",
    data: {
      labels: names,
      datasets: [
        {
          label: "Total",
          data: totals
        },
        {
          label: "Completed",
          data: completed
        }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false
    }
  });
}

function renderStaffPercentCards(staffBreakdown) {
  const names = Object.keys(staffBreakdown);

  if (!names.length) {
    staffPercentGrid.innerHTML = `<div class="empty-percent">No data yet.</div>`;
    return;
  }

  staffPercentGrid.innerHTML = names
    .map((name) => {
      const total = Number(staffBreakdown[name].total || 0);
      const completed = Number(staffBreakdown[name].completed || 0);
      const percent = total > 0 ? Math.round((completed / total) * 100) : 0;

      return `
        <div class="percent-card">
          <div class="percent-name">${escapeHtml(name)}</div>
          <div class="percent-meta">
            <span>${completed} / ${total} completed</span>
            <span>${percent}%</span>
          </div>
          <div class="percent-bar">
            <div class="percent-fill" style="width: ${percent}%"></div>
          </div>
          <div class="percent-value">${percent}%</div>
        </div>
      `;
    })
    .join("");
}

function readExcelFile(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();

    reader.onload = function (e) {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: "array" });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        const jsonRows = XLSX.utils.sheet_to_json(worksheet, { defval: "" });
        resolve(jsonRows);
      } catch (err) {
        reject(err);
      }
    };

    reader.onerror = () => reject(new Error("Could not read Excel file."));
    reader.readAsArrayBuffer(file);
  });
}

function mapRowsToRecords(rows) {
  return rows
    .map((row) => {
      const username = pickField(row, ["Affiliate Username", "affiliate username", "AffiliateUsername"]);
      const email = pickField(row, ["Email", "email"]);
      const regTimeRaw = pickField(row, ["Registration Time", "registration time", "RegistrationTime"]);
      const phone = pickField(row, ["Phone Number", "phone number", "PhoneNumber"]);

      return {
        username: String(username || "").trim(),
        email: String(email || "").trim(),
        regTime: formatRegistrationTime(regTimeRaw),
        phone: String(phone || "").trim()
      };
    })
    .filter((r) => r.username || r.email || r.regTime || r.phone);
}

function pickField(row, keys) {
  for (const key of keys) {
    if (row[key] !== undefined) return row[key];
  }
  return "";
}

function formatRegistrationTime(value) {
  if (value === null || value === undefined || value === "") return "";

  if (typeof value === "number") {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (parsed) {
      const yyyy = parsed.y;
      const mm = String(parsed.m).padStart(2, "0");
      const dd = String(parsed.d).padStart(2, "0");
      const hh = String(parsed.H || 0).padStart(2, "0");
      const min = String(parsed.M || 0).padStart(2, "0");
      return `${yyyy}-${mm}-${dd} ${hh}:${min}`;
    }
  }

  const str = String(value).trim();
  return str;
}

async function loadStaffTasks() {
  if (!currentUser || currentUser.role === "Admin") return;

  refreshTasksBtn.disabled = true;

  // 🔥 CLEAR UI FIRST (fix glitch)
  currentTasks = [];
  taskTableBody.innerHTML = `<tr><td colspan="8" class="empty-cell">Loading...</td></tr>`;
  taskCount.textContent = "0 Tasks";

  setMessage(staffMsg, "Loading tasks...", "info");

  try {
    const res = await postData({
      action: "getStaffTasks",
      code: currentUser.code
    });

    // 🔥 SAFETY: ensure array
    let data = Array.isArray(res.data) ? res.data : [];

    // 🔥 IMPORTANT FIX: only latest batch
    const latestBatch = data.length ? data[0].batchId : null;
    if (latestBatch) {
      data = data.filter(t => t.batchId === latestBatch);
    }

    currentTasks = data;

    renderTasks();

    setMessage(staffMsg, `Loaded ${currentTasks.length} tasks.`, "success");

  } catch (err) {
    console.error("Load tasks error:", err);
    setMessage(staffMsg, `Failed to load tasks: ${err.message}`, "error");
  } finally {
    refreshTasksBtn.disabled = false;
  }
}

function renderTasks() {
  const q = searchInput.value.trim().toLowerCase();
  const filterStatus = statusFilter.value;

  const pendingCount = currentTasks.filter((task) => String(task.status || "").trim() === "Pending").length;
  taskCount.textContent = `${pendingCount} Tasks`;

  const filtered = currentTasks.filter((task) => {
    return (
      (!q ||
        String(task.brandName || "").toLowerCase().includes(q) ||
        String(task.username || "").toLowerCase().includes(q) ||
        String(task.email || "").toLowerCase().includes(q) ||
        String(task.phone || "").toLowerCase().includes(q)) &&
      (!filterStatus || task.status === filterStatus)
    );
  });

  if (!filtered.length) {
    taskTableBody.innerHTML = `<tr><td colspan="8" class="empty-cell">No tasks found.</td></tr>`;
    return;
  }

  taskTableBody.innerHTML = filtered.map(task => `
    <tr>
      <td>${escapeHtml(task.brandName || "")}</td>
      <td>${escapeHtml(task.username || "")}</td>
      <td>${escapeHtml(task.email || "")}</td>
      <td>${escapeHtml(task.regTime || "")}</td>
      <td>${escapeHtml(task.phone || "")}</td>
      <td>
        <select class="status-select" data-task-id="${task.id}">
          ${STATUS_OPTIONS.map(s =>
            `<option ${task.status === s ? "selected" : ""}>${s}</option>`
          ).join("")}
        </select>
      </td>
      <td>
        <textarea class="remark-input" data-task-id="${task.id}">${task.remark || ""}</textarea>
      </td>
      <td>
        <button class="save-btn" onclick="saveTask('${task.id}', this)">Save</button>
      </td>
    </tr>
  `).join("");
}

async function saveTask(taskId, btn) {
  const statusEl = document.querySelector(`.status-select[data-task-id="${cssEscape(taskId)}"]`);
  const remarkEl = document.querySelector(`.remark-input[data-task-id="${cssEscape(taskId)}"]`);

  if (!statusEl || !remarkEl) return;

  const status = statusEl.value;
  const remark = remarkEl.value.trim();

  if (btn) btn.disabled = true;
  setMessage(staffMsg, "Saving update...", "info");

  try {
    const res = await postData({
      action: "updateTask",
      taskId,
      status,
      remark,
      code: currentUser.code
    });

    if (res.success) {
      const task = currentTasks.find((t) => t.id === taskId);
      if (task) {
        task.status = status;
        task.remark = remark;
      }
      renderTasks();
      setMessage(staffMsg, "Task updated successfully.", "success");
    } else {
      setMessage(staffMsg, res.error || "Update failed.", "error");
    }
  } catch (err) {
    console.error("Save task error:", err);
    setMessage(staffMsg, `Update failed: ${err.message}`, "error");
  } finally {
    if (btn) btn.disabled = false;
  }
}

async function handleExport() {
  if (!currentUser || currentUser.role !== "Admin") return;

  exportBtn.disabled = true;
  setMessage(exportMsg, "Preparing export...", "info");

  try {
    const res = await postData({
      action: "exportReport"
    });

    if (!res.data || !Array.isArray(res.data)) {
      setMessage(exportMsg, "No report data found.", "error");
      return;
    }

    downloadCSV(res.data, "assigned_tasks_report.csv");
    setMessage(exportMsg, "Report downloaded successfully.", "success");
  } catch (err) {
    console.error("Export error:", err);
    setMessage(exportMsg, `Export failed: ${err.message}`, "error");
  } finally {
    exportBtn.disabled = false;
  }
}

function downloadCSV(rows, fileName) {
  const csv = rows
    .map((row) =>
      row
        .map((cell) => {
          const value = cell == null ? "" : String(cell);
          return `"${value.replace(/"/g, '""')}"`;
        })
        .join(",")
    )
    .join("\n");

  const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = fileName;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

function escapeHtml(str) {
  return String(str)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function escapeJs(str) {
  return String(str).replaceAll("\\", "\\\\").replaceAll("'", "\\'");
}

function cssEscape(str) {
  if (window.CSS && CSS.escape) return CSS.escape(str);
  return String(str).replace(/"/g, '\\"');
}
