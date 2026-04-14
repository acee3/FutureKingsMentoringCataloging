const state = {
  rows: [],
  columns: [],
  filters: {},
  sort: { key: null, direction: "asc" },
  page: 1,
  pageSize: 10,
  totalRows: 0,
};

const HIDDEN_COLUMNS = new Set(["slide_texts"]);

async function fetchJson(url, options = {}) {
  const response = await fetch(url, {
    headers: { "Content-Type": "application/json" },
    ...options,
  });
  const data = await response.json().catch(() => ({}));
  if (!response.ok) {
    throw new Error(data.detail || `Request failed: ${response.status}`);
  }
  return data;
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function getTableRows() {
  const filtered = state.rows.filter((row) =>
    state.columns.every((column) => {
      const filterValue = (state.filters[column] || "").trim().toLowerCase();
      if (!filterValue) {
        return true;
      }
      return String(row[column] ?? "").toLowerCase().includes(filterValue);
    }),
  );

  const { key, direction } = state.sort;
  if (!key) {
    return filtered;
  }

  return [...filtered].sort((left, right) => {
    const leftValue = String(left[key] ?? "");
    const rightValue = String(right[key] ?? "");
    const comparison = leftValue.localeCompare(rightValue, undefined, {
      numeric: true,
      sensitivity: "base",
    });
    return direction === "asc" ? comparison : -comparison;
  });
}

function renderTable() {
  const container = document.getElementById("table-container");
  if (!state.columns.length) {
    container.innerHTML = "<p>No indexed workbook rows yet.</p>";
    return;
  }

  const sortIndicator = (column) => {
    if (state.sort.key !== column) {
      return "";
    }
    return state.sort.direction === "asc" ? " ▲" : " ▼";
  };

  const headerCells = state.columns
    .map(
      (column) => `
        <th>
          <button type="button" data-sort-key="${escapeHtml(column)}">
            ${escapeHtml(column)}${sortIndicator(column)}
          </button>
        </th>
      `,
    )
    .join("");

  const filterCells = state.columns
    .map(
      (column) => `
        <th>
          <input
            class="filter-input"
            type="text"
            data-filter-key="${escapeHtml(column)}"
            value="${escapeHtml(state.filters[column] || "")}"
            placeholder="Filter"
          />
        </th>
      `,
    )
    .join("");

  const bodyRows = getTableRows()
    .map(
      (row) => `
        <tr>
          ${state.columns
            .map(
              (column) => `
                <td><div class="cell-wrap">${escapeHtml(row[column] ?? "")}</div></td>
              `,
            )
            .join("")}
        </tr>
      `,
    )
    .join("");

  container.innerHTML = `
    <table>
      <thead>
        <tr>${headerCells}</tr>
        <tr>${filterCells}</tr>
      </thead>
      <tbody>${bodyRows}</tbody>
    </table>
  `;

  container.querySelectorAll("[data-sort-key]").forEach((button) => {
    button.addEventListener("click", () => {
      const key = button.dataset.sortKey;
      if (state.sort.key === key) {
        state.sort.direction = state.sort.direction === "asc" ? "desc" : "asc";
      } else {
        state.sort = { key, direction: "asc" };
      }
      renderTable();
    });
  });

  container.querySelectorAll("[data-filter-key]").forEach((input) => {
    input.addEventListener("input", () => {
      state.filters[input.dataset.filterKey] = input.value;
      renderTable();
    });
  });
}

function renderSearchResults(items) {
  const results = document.getElementById("search-results");
  if (!items.length) {
    results.classList.remove("hidden");
    results.innerHTML = "<p>No semantic matches found.</p>";
    return;
  }

  results.classList.remove("hidden");
  results.innerHTML = items
    .map((item) => {
      const summary = Object.entries(item.metadata || {})
        .filter(([, value]) => value)
        .slice(0, 6)
        .map(([key, value]) => `<strong>${escapeHtml(key)}:</strong> ${escapeHtml(value)}`)
        .join("<br />");

      return `
        <article class="result-card">
          <div class="result-head">
            <div>
              <h3>${escapeHtml(item.title)}</h3>
              <div class="result-meta">
                ${escapeHtml(item.workbook_name)} · ${escapeHtml(item.sheet_name)} · row ${escapeHtml(item.row_number)}
              </div>
            </div>
            <div class="result-score">${(item.score * 100).toFixed(1)}%</div>
          </div>
          <div class="result-meta">${summary}</div>
        </article>
      `;
    })
    .join("");
}

async function loadHealth() {
  const health = await fetchJson("/api/health");
  document.getElementById("indexed-count").textContent = health.indexed_rows;
}

function updatePaginationControls() {
  const totalPages = Math.max(1, Math.ceil(state.totalRows / state.pageSize));
  document.getElementById("page-status").textContent = `Page ${state.page} of ${totalPages}`;
  document.getElementById("prev-page-button").disabled = state.page <= 1;
  document.getElementById("next-page-button").disabled = state.page >= totalPages;
}

async function loadRows() {
  const offset = (state.page - 1) * state.pageSize;
  const payload = await fetchJson(`/api/presentations?limit=${state.pageSize}&offset=${offset}`);
  const items = payload.items || [];
  state.totalRows = payload.total || 0;
  state.rows = items.map((item) => ({ ...(item.metadata || {}) }));
  state.columns = (payload.columns || []).filter((key) => !HIDDEN_COLUMNS.has(key));
  updatePaginationControls();
  renderTable();
}

async function uploadWorkbook() {
  const fileInput = document.getElementById("workbook-file-input");
  const button = document.getElementById("upload-button");
  const file = fileInput.files?.[0];
  if (!file) {
    throw new Error("Choose a workbook file before uploading.");
  }

  button.disabled = true;
  button.textContent = "Uploading...";
  try {
    const formData = new FormData();
    formData.append("workbook", file);
    const response = await fetch("/api/index", {
      method: "POST",
      body: formData,
    });
    const payload = await response.json().catch(() => ({}));
    if (!response.ok) {
      throw new Error(payload.detail || `Request failed: ${response.status}`);
    }
    document.getElementById("workbook-name").textContent = file.name;
    state.page = 1;
    await loadHealth();
    await loadRows();
  } finally {
    button.disabled = false;
    button.textContent = "Upload and index";
  }
}

async function handleSearch(event) {
  event.preventDefault();
  const errorBox = document.getElementById("search-error");
  const resultsBox = document.getElementById("search-results");
  errorBox.classList.add("hidden");
  resultsBox.classList.add("hidden");

  const query = document.getElementById("query-input").value.trim();
  const topK = Number(document.getElementById("top-k-input").value || 10);
  if (!query) {
    return;
  }

  try {
    const payload = await fetchJson("/api/search", {
      method: "POST",
      body: JSON.stringify({ query, top_k: topK }),
    });
    renderSearchResults(payload.items || []);
  } catch (error) {
    errorBox.textContent = error.message;
    errorBox.classList.remove("hidden");
  }
}

function resetFilters() {
  state.filters = {};
  state.sort = { key: null, direction: "asc" };
  renderTable();
}

async function goToPreviousPage() {
  if (state.page <= 1) {
    return;
  }
  state.page -= 1;
  await loadRows();
}

async function goToNextPage() {
  const totalPages = Math.max(1, Math.ceil(state.totalRows / state.pageSize));
  if (state.page >= totalPages) {
    return;
  }
  state.page += 1;
  await loadRows();
}

async function bootstrap() {
  document.getElementById("semantic-search-form").addEventListener("submit", handleSearch);
  document.getElementById("upload-button").addEventListener("click", async () => {
    const errorBox = document.getElementById("search-error");
    errorBox.classList.add("hidden");
    try {
      await uploadWorkbook();
    } catch (error) {
      errorBox.textContent = error.message;
      errorBox.classList.remove("hidden");
    }
  });
  document.getElementById("reset-filters-button").addEventListener("click", resetFilters);
  document.getElementById("prev-page-button").addEventListener("click", goToPreviousPage);
  document.getElementById("next-page-button").addEventListener("click", goToNextPage);
  await Promise.all([loadHealth(), loadRows()]);
}

bootstrap().catch((error) => {
  const container = document.getElementById("table-container");
  container.innerHTML = `<div class="message error">${escapeHtml(error.message)}</div>`;
});
