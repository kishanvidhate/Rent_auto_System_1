/* Automatic Rent System converted from Kishankumar2.0.xlsm
   Excel editable cells:
   - D2  : Rent month date
   - H2:H5 : Today's meter readings for unit-based tenants
   - J7  : Actual light bill for Somnath More

   VBA behavior converted to JavaScript:
   When H2:H5 changes, the old H value is copied into the same row's G cell
   before storing the new H value.
*/

const STORAGE_KEY = "automatic-rent-system-v1";

const defaultState = {
  rentDate: "2026-06-01", // Excel serial 46174 from D2
  tenants: [
    {
      excelRow: 2,
      srNo: 1,
      tenant: "ABC",
      mobile: "7030705076",
      rent: 2500,
      lightType: "Unit based",
      lastReading: 100,
      todayReading: 125,
      unitRate: 12,
      lightBill: null,
      maintenance: 0
    },
    {
      excelRow: 3,
      srNo: 2,
      tenant: "Sagita Tai",
      mobile: "9307384771",
      rent: 2500,
      lightType: "Unit based",
      lastReading: 100,
      todayReading: 125,
      unitRate: 12,
      lightBill: null,
      maintenance: 0
    },
    {
      excelRow: 4,
      srNo: 3,
      tenant: "Rani Tai",
      mobile: "7821883076",
      rent: 2500,
      lightType: "Unit based",
      lastReading: 100,
      todayReading: 125,
      unitRate: 12,
      lightBill: null,
      maintenance: 0
    },
    {
      excelRow: 5,
      srNo: 4,
      tenant: "Rambhajan Kushwah",
      mobile: "8949137989",
      rent: 2500,
      lightType: "Unit based",
      lastReading: 100,
      todayReading: 125,
      unitRate: 12,
      lightBill: null,
      maintenance: 0
    },
    {
      excelRow: 6,
      srNo: 5,
      tenant: "Rabi Satpute",
      mobile: "9673305643",
      rent: 3500,
      lightType: "Fixed",
      lastReading: "NA",
      todayReading: "NA",
      unitRate: "NA",
      lightBill: 1000,
      maintenance: 0
    },
    {
      excelRow: 7,
      srNo: 6,
      tenant: "Somnath More",
      mobile: "9673305643",
      rent: 4500,
      lightType: "Actual",
      lastReading: "NA",
      todayReading: "NA",
      unitRate: "NA",
      lightBill: 1200,
      maintenance: 800
    }
  ]
};

let state = loadState();
let previousTodayReading = new Map();

const rentDateInput = document.getElementById("rentDate");
const rentTableBody = document.querySelector("#rentTable tbody");
const rentTableFoot = document.querySelector("#rentTable tfoot");
const resetBtn = document.getElementById("resetBtn");

rentDateInput.addEventListener("change", () => {
  state.rentDate = rentDateInput.value || defaultState.rentDate;
  saveState();
  render();
});

resetBtn.addEventListener("click", () => {
  localStorage.removeItem(STORAGE_KEY);
  state = structuredClone(defaultState);
  previousTodayReading = new Map();
  saveState();
  render();
});

function loadState() {
  const savedJson = localStorage.getItem(STORAGE_KEY);
  if (!savedJson) return structuredClone(defaultState);

  try {
    const saved = JSON.parse(savedJson);
    return {
      rentDate: saved.rentDate || defaultState.rentDate,
      tenants: defaultState.tenants.map((tenant, index) => ({
        ...tenant,
        ...(saved.tenants?.[index] || {})
      }))
    };
  } catch {
    return structuredClone(defaultState);
  }
}

function saveState() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
}

function toNumber(value) {
  const numberValue = Number(value);
  return Number.isFinite(numberValue) ? numberValue : 0;
}

function formatMoney(value) {
  return Math.round(toNumber(value)).toLocaleString("en-IN");
}

function formatMonthYear(dateValue) {
  const date = new Date(`${dateValue}T00:00:00`);
  if (Number.isNaN(date.getTime())) return "";

  const month = date.toLocaleString("en-US", { month: "long" });
  const year = String(date.getFullYear()).slice(-2);
  return `${month}-${year}`;
}

function getLightBill(tenant) {
  if (tenant.lightType === "Unit based") {
    return (toNumber(tenant.todayReading) - toNumber(tenant.lastReading)) * toNumber(tenant.unitRate);
  }

  return toNumber(tenant.lightBill);
}

function getTotalAmount(tenant) {
  return toNumber(tenant.rent) + getLightBill(tenant) + toNumber(tenant.maintenance);
}

function getWhatsAppMessage(tenant) {
  return [
    `Mr./Miss ${tenant.tenant.trim()},`,
    `${formatMonthYear(state.rentDate)} Month Rent Details`,
    `Rent: ${Math.round(toNumber(tenant.rent))}`,
    `Maintenance: ${Math.round(toNumber(tenant.maintenance))}`,
    `Light Bill: ${Math.round(getLightBill(tenant))}`,
    `Total Amount: ${Math.round(getTotalAmount(tenant))}`,
    "Please pay on time."
  ].join("\n");
}

function getWhatsAppUrl(tenant) {
  const mobile = tenant.mobile.replace(/\D/g, "");
  const numberWithCountryCode = mobile.length === 10 ? `91${mobile}` : mobile;
  return `https://wa.me/${numberWithCountryCode}?text=${encodeURIComponent(getWhatsAppMessage(tenant))}`;
}

function render() {
  rentDateInput.value = state.rentDate;
  rentTableBody.innerHTML = "";
  rentTableFoot.innerHTML = "";

  let totalRent = 0;
  let totalLight = 0;
  let totalMaintenance = 0;
  let grandTotal = 0;

  state.tenants.forEach((tenant, index) => {
    const lightBill = getLightBill(tenant);
    const totalAmount = getTotalAmount(tenant);

    totalRent += toNumber(tenant.rent);
    totalLight += lightBill;
    totalMaintenance += toNumber(tenant.maintenance);
    grandTotal += totalAmount;

    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td class="number">${tenant.srNo}</td>
      <td>${tenant.tenant}</td>
      <td>${tenant.mobile}</td>
      <td>${formatMonthYear(state.rentDate)}</td>
      <td class="number">${formatMoney(tenant.rent)}</td>
      <td>${tenant.lightType}</td>
      <td class="number locked">${tenant.lastReading}</td>
      <td class="number">${renderTodayReadingCell(tenant, index)}</td>
      <td class="number locked">${tenant.unitRate}</td>
      <td class="number">${renderLightBillCell(tenant, index, lightBill)}</td>
      <td class="number locked">${formatMoney(tenant.maintenance)}</td>
      <td class="number"><strong>${formatMoney(totalAmount)}</strong></td>
      <td>
        <div class="message">${escapeHtml(getWhatsAppMessage(tenant))}</div>
        <a class="whatsapp-link" href="${getWhatsAppUrl(tenant)}" target="_blank" rel="noopener">Send WhatsApp</a>
      </td>
    `;
    rentTableBody.appendChild(tr);
  });

  rentTableFoot.innerHTML = `
    <tr>
      <td></td>
      <td><strong>Total Rent</strong></td>
      <td></td>
      <td></td>
      <td class="number">${formatMoney(totalRent)}</td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td class="number">${formatMoney(totalLight)}</td>
      <td class="number">${formatMoney(totalMaintenance)}</td>
      <td class="number">${formatMoney(grandTotal)}</td>
      <td></td>
    </tr>
  `;

  document.querySelectorAll(".reading-input").forEach((input) => {
    input.addEventListener("focus", handleReadingFocus);
    input.addEventListener("change", handleReadingChange);
  });

  document.querySelectorAll(".light-input").forEach((input) => {
    input.addEventListener("change", handleActualLightChange);
  });
}

function renderTodayReadingCell(tenant, index) {
  if (tenant.lightType !== "Unit based") return `<span class="locked">${tenant.todayReading}</span>`;

  return `
    <input
      class="reading-input"
      type="number"
      step="1"
      min="0"
      data-index="${index}"
      value="${tenant.todayReading}"
      aria-label="Today reading for ${escapeHtml(tenant.tenant)}"
    />
  `;
}

function renderLightBillCell(tenant, index, lightBill) {
  if (tenant.lightType === "Actual") {
    return `
      <input
        class="light-input"
        type="number"
        step="1"
        min="0"
        data-index="${index}"
        value="${tenant.lightBill}"
        aria-label="Actual light bill for ${escapeHtml(tenant.tenant)}"
      />
    `;
  }

  return `<span class="locked">${formatMoney(lightBill)}</span>`;
}

function handleReadingFocus(event) {
  const index = Number(event.target.dataset.index);
  previousTodayReading.set(index, toNumber(event.target.value));
}

function handleReadingChange(event) {
  const index = Number(event.target.dataset.index);
  const tenant = state.tenants[index];
  const oldValue = previousTodayReading.has(index)
    ? previousTodayReading.get(index)
    : toNumber(tenant.todayReading);
  const newValue = toNumber(event.target.value);

  // Same as the VBA macro: old H value becomes G; new value remains in H.
  tenant.lastReading = oldValue;
  tenant.todayReading = newValue;

  previousTodayReading.delete(index);
  saveState();
  render();
}

function handleActualLightChange(event) {
  const index = Number(event.target.dataset.index);
  state.tenants[index].lightBill = toNumber(event.target.value);
  saveState();
  render();
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

render();
