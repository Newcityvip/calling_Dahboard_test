const API_URL = "https://script.google.com/macros/s/AKfycbzXNFfkW4Pa2GQfwqTJCNwNv7-4RH4nceFxjEIIvfFv-NR_sRG9nPljIWYjTvI2Tjmn/exec";

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
const excelFileInput = document.getElementById("excelFile");
const uploadMsg = document.getElementById("uploadMsg");

const exportBtn = document.getElementById("exportBtn");
const exportMsg = document.getElementById("exportMsg");

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
refreshTasksBtn.addEventListener("click", loadStaffTasks);
searchInput.addEventListener("input", renderTasks);
statusFilter.addEventListener("change", renderTasks);
if (teamGroupSelect) teamGroupSelect.addEventListener("change", handleTeamGroupChange);

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
      fileName: file.name,
      records
    });

    if (res.success) {
      setMessage(
        uploadMsg,
        `Upload complete.
Total rows: ${res.total}
After duplicate cleanup: ${res.cleaned}
Duplicates removed: ${res.duplicatesRemoved}
Batch ID: ${res.batchId}`,
        "success"
      );
      excelFileInput.value = "";
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
      const regTime = pickField(row, ["Registration Time", "registration time", "RegistrationTime"]);
      const phone = pickField(row, ["Phone Number", "phone number", "PhoneNumber"]);

      return {
        username: String(username || "").trim(),
        email: String(email || "").trim(),
        regTime: String(regTime || "").trim(),
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

async function loadStaffTasks() {
  if (!currentUser || currentUser.role === "Admin") return;

  refreshTasksBtn.disabled = true;
  setMessage(staffMsg, "Loading tasks...", "info");

  try {
    const res = await postData({
      action: "getStaffTasks",
      code: currentUser.code
    });

    currentTasks = Array.isArray(res.data) ? res.data : [];
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

  const filtered = currentTasks.filter((task) => {
    const matchesSearch =
      !q ||
      String(task.username || "").toLowerCase().includes(q) ||
      String(task.email || "").toLowerCase().includes(q) ||
      String(task.phone || "").toLowerCase().includes(q);

    const matchesStatus = !filterStatus || task.status === filterStatus;

    return matchesSearch && matchesStatus;
  });

  taskCount.textContent = `${filtered.length} Tasks`;

  if (!filtered.length) {
    taskTableBody.innerHTML = `<tr><td colspan="7" class="empty-cell">No matching tasks found.</td></tr>`;
    return;
  }

  taskTableBody.innerHTML = filtered
    .map((task) => {
      const statusOptions = STATUS_OPTIONS.map(
        (status) =>
          `<option value="${escapeHtml(status)}" ${task.status === status ? "selected" : ""}>${escapeHtml(status)}</option>`
      ).join("");

      return `
        <tr>
          <td>${escapeHtml(task.username || "")}</td>
          <td>${escapeHtml(task.email || "")}</td>
          <td>${escapeHtml(task.regTime || "")}</td>
          <td>${escapeHtml(task.phone || "")}</td>
          <td>
            <select class="status-select" data-task-id="${escapeHtml(task.id)}">
              ${statusOptions}
            </select>
          </td>
          <td>
            <textarea class="remark-input" data-task-id="${escapeHtml(task.id)}" placeholder="Enter note...">${escapeHtml(task.remark || "")}</textarea>
          </td>
          <td>
            <button class="save-btn" onclick="saveTask('${escapeJs(task.id)}', this)">Save</button>
          </td>
        </tr>
      `;
    })
    .join("");
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
