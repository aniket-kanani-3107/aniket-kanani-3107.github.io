const storage = {
  get(key, fallback) {
    try {
      const raw = localStorage.getItem(key);
      return raw ? JSON.parse(raw) : fallback;
    } catch (error) {
      return fallback;
    }
  },
  set(key, value) {
    localStorage.setItem(key, JSON.stringify(value));
  },
};

const setText = (id, value) => {
  const el = document.getElementById(id);
  if (el) {
    el.textContent = value;
  }
};

const setValue = (id, value) => {
  const el = document.getElementById(id);
  if (el) {
    el.value = value;
  }
};

const formatDate = () =>
  new Intl.DateTimeFormat("en-GB", {
    weekday: "long",
    day: "2-digit",
    month: "short",
    year: "numeric",
  }).format(new Date());

const formatTime = (timeZone, hour12) =>
  new Intl.DateTimeFormat("en-GB", {
    timeZone,
    hour: "2-digit",
    minute: "2-digit",
    second: "2-digit",
    hour12,
  }).format(new Date());

const financeRegions = ["germany", "india"];

const loadFinanceNotes = () => {
  financeRegions.forEach((region) => {
    const text = storage.get(`finance.${region}.text`, "");
    const updated = storage.get(`finance.${region}.updated`, null);
    setValue(`finance-${region}`, text);
    setText(
      `finance-${region}-updated`,
      updated ? `Last saved: ${updated}` : "Not saved yet"
    );
  });
};

const saveFinanceNotes = (region) => {
  const textarea = document.getElementById(`finance-${region}`);
  if (!textarea) return;
  const now = new Date().toLocaleString();
  storage.set(`finance.${region}.text`, textarea.value.trim());
  storage.set(`finance.${region}.updated`, now);
  setText(`finance-${region}-updated`, `Last saved: ${now}`);
};

const parseCSV = (text) => {
  const rows = [];
  let row = [];
  let value = "";
  let inQuotes = false;

  for (let i = 0; i < text.length; i += 1) {
    const char = text[i];
    const next = text[i + 1];

    if (char === '"') {
      if (inQuotes && next === '"') {
        value += '"';
        i += 1;
      } else {
        inQuotes = !inQuotes;
      }
    } else if (char === "," && !inQuotes) {
      row.push(value);
      value = "";
    } else if ((char === "\n" || char === "\r") && !inQuotes) {
      if (char === "\r" && next === "\n") {
        i += 1;
      }
      row.push(value);
      if (row.some((cell) => cell.trim() !== "")) {
        rows.push(row);
      }
      row = [];
      value = "";
    } else {
      value += char;
    }
  }

  if (value.length || row.length) {
    row.push(value);
    if (row.some((cell) => cell.trim() !== "")) {
      rows.push(row);
    }
  }

  return rows;
};

const renderScheduleTable = (rows) => {
  const tableWrapper = document.getElementById("schedule-table");
  if (!tableWrapper) return;

  if (!rows || rows.length === 0) {
    tableWrapper.innerHTML = '<div class="empty-state">Upload a CSV file to see your schedule here.</div>';
    return;
  }

  const header = rows[0];
  const bodyRows = rows.slice(1);

  const table = document.createElement("table");
  const thead = document.createElement("thead");
  const headRow = document.createElement("tr");
  header.forEach((cell) => {
    const th = document.createElement("th");
    th.textContent = cell.trim() || "-";
    headRow.appendChild(th);
  });
  thead.appendChild(headRow);
  table.appendChild(thead);

  const tbody = document.createElement("tbody");
  bodyRows.forEach((row) => {
    const tr = document.createElement("tr");
    row.forEach((cell) => {
      const td = document.createElement("td");
      td.textContent = cell.trim();
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });
  table.appendChild(tbody);

  tableWrapper.innerHTML = "";
  tableWrapper.appendChild(table);
};

const loadSchedule = () => {
  const schedule = storage.get("schedule.rows", []);
  renderScheduleTable(schedule);
  const updated = storage.get("schedule.updated", null);
  setText(
    "schedule-status",
    updated ? `Last loaded: ${updated}` : "No schedule loaded."
  );
};

const saveSchedule = (rows) => {
  const now = new Date().toLocaleString();
  storage.set("schedule.rows", rows);
  storage.set("schedule.updated", now);
  setText("schedule-status", `Last loaded: ${now}`);
  renderScheduleTable(rows);
};

const statusLabels = {
  "not-started": "Not started",
  "in-progress": "In progress",
  done: "Done",
};

const loadTasks = () => storage.get("project.tasks", []);

const saveTasks = (tasks) => storage.set("project.tasks", tasks);

const renderTasks = (tasks) => {
  const list = document.getElementById("project-list");
  if (!list) return;

  list.innerHTML = "";
  tasks.forEach((task, index) => {
    const item = document.createElement("div");
    item.className = "task-item";

    const nameInput = document.createElement("input");
    nameInput.type = "text";
    nameInput.value = task.name;
    nameInput.addEventListener("input", () => {
      tasks[index].name = nameInput.value;
      saveTasks(tasks);
    });

    const statusSelect = document.createElement("select");
    Object.entries(statusLabels).forEach(([value, label]) => {
      const option = document.createElement("option");
      option.value = value;
      option.textContent = label;
      if (task.status === value) option.selected = true;
      statusSelect.appendChild(option);
    });
    statusSelect.addEventListener("change", () => {
      tasks[index].status = statusSelect.value;
      saveTasks(tasks);
      updateProgress(tasks);
    });

    const removeButton = document.createElement("button");
    removeButton.type = "button";
    removeButton.textContent = "✕";
    removeButton.addEventListener("click", () => {
      tasks.splice(index, 1);
      saveTasks(tasks);
      renderTasks(tasks);
      updateProgress(tasks);
    });

    item.appendChild(nameInput);
    item.appendChild(statusSelect);
    item.appendChild(removeButton);
    list.appendChild(item);
  });
};

const updateProgress = (tasks) => {
  const total = tasks.length;
  const doneCount = tasks.filter((task) => task.status === "done").length;
  const percent = total ? Math.round((doneCount / total) * 100) : 0;
  const progress = document.getElementById("project-progress");
  if (progress) {
    progress.value = percent;
  }
  setText("project-progress-text", `${percent}% complete`);
};

const loadAccounts = () =>
  storage.get("accounts.data", {
    personal: { email: "", drive: "" },
    work: { email: "", drive: "" },
    university: { email: "", drive: "" },
  });

const saveAccounts = (data) => storage.set("accounts.data", data);

const renderAccounts = (data) => {
  setValue("account-personal-email", data.personal.email || "");
  setValue("account-personal-drive", data.personal.drive || "");
  setValue("account-work-email", data.work.email || "");
  setValue("account-work-drive", data.work.drive || "");
  setValue("account-uni-email", data.university.email || "");
  setValue("account-uni-drive", data.university.drive || "");
};

const updateDateDisplay = () => setText("date-display", formatDate());

const updateTimeZones = () => {
  setText("time-de-24", formatTime("Europe/Berlin", false));
  setText("time-de-12", formatTime("Europe/Berlin", true));
  setText("time-in-24", formatTime("Asia/Kolkata", false));
  setText("time-in-12", formatTime("Asia/Kolkata", true));
};

const searchLinks = {
  chatgpt: "https://chat.openai.com/?q=",
  gemini: "https://gemini.google.com/app?prompt=",
  antigravity: "https://www.antigravity.ai/?q=",
  github: "https://github.com/search?q=",
};

const initSearch = () => {
  const queryInput = document.getElementById("search-query");
  if (!queryInput) return;
  document.querySelectorAll("[data-engine]").forEach((button) => {
    button.addEventListener("click", () => {
      const query = queryInput.value.trim();
      const engine = button.dataset.engine;
      const baseUrl = searchLinks[engine];
      if (!baseUrl) return;
      const url = query ? `${baseUrl}${encodeURIComponent(query)}` : baseUrl;
      window.open(url, "_blank", "noopener");
    });
  });
};

const init = () => {
  updateDateDisplay();
  updateTimeZones();
  setInterval(updateTimeZones, 1000);

  loadFinanceNotes();
  document.querySelectorAll("[data-action='save-finance']").forEach((button) => {
    button.addEventListener("click", () => {
      const region = button.dataset.region;
      if (region) saveFinanceNotes(region);
    });
  });

  const scheduleInput = document.getElementById("schedule-file");
  if (scheduleInput) {
    scheduleInput.addEventListener("change", (event) => {
      const file = event.target.files[0];
      if (!file) return;
      const reader = new FileReader();
      reader.onload = () => {
        const rows = parseCSV(reader.result);
        saveSchedule(rows);
      };
      reader.readAsText(file);
    });
  }

  const scheduleClear = document.getElementById("schedule-clear");
  if (scheduleClear) {
    scheduleClear.addEventListener("click", () => {
      storage.set("schedule.rows", []);
      storage.set("schedule.updated", null);
      loadSchedule();
    });
  }

  loadSchedule();

  const tasks = loadTasks();
  renderTasks(tasks);
  updateProgress(tasks);

  const addButton = document.getElementById("project-add");
  if (addButton) {
    addButton.addEventListener("click", () => {
      const nameInput = document.getElementById("project-task-name");
      const statusInput = document.getElementById("project-task-status");
      if (!nameInput || !statusInput) return;
      const name = nameInput.value.trim();
      if (!name) return;
      tasks.push({ name, status: statusInput.value });
      nameInput.value = "";
      saveTasks(tasks);
      renderTasks(tasks);
      updateProgress(tasks);
    });
  }

  const accountData = loadAccounts();
  renderAccounts(accountData);
  const accountSave = document.getElementById("account-save");
  if (accountSave) {
    accountSave.addEventListener("click", () => {
      const updated = {
        personal: {
          email: document.getElementById("account-personal-email").value.trim(),
          drive: document.getElementById("account-personal-drive").value.trim(),
        },
        work: {
          email: document.getElementById("account-work-email").value.trim(),
          drive: document.getElementById("account-work-drive").value.trim(),
        },
        university: {
          email: document.getElementById("account-uni-email").value.trim(),
          drive: document.getElementById("account-uni-drive").value.trim(),
        },
      };
      saveAccounts(updated);
      const now = new Date().toLocaleString();
      storage.set("accounts.updated", now);
      setText("account-status", `Saved: ${now}`);
    });
  }

  const accountUpdated = storage.get("accounts.updated", null);
  if (accountUpdated) {
    setText("account-status", `Saved: ${accountUpdated}`);
  }

  initSearch();
};

document.addEventListener("DOMContentLoaded", init);
