/* global Office, Word */
// v1.0.0.34
// ------------------
// Data 
// ------------------
const salesReps = [
  { name: "Clifford Sales", email: "sales@cliffeng.com" },
  { name: "Darryn Rowbotham", email: "rowbothamd@cliffeng.com" },
  { name: "Scott Liebenberg", email: "scott@cliffeng.com" },
  { name: "Clive White", email: "clive@cliffeng.com" },
  { name: "Iain Ambler", email: "iain@cliffeng.com" },
  { name: "Graham Raynor", email: "graham@cliffeng.com" }
];

const cliffordCompany = [
  { name: "Clifford Machines and Technology" },
  { name: "Clifford North America" }
];

const currency = [
  { name: "USD" },
  { name: "EUR" },
  { name: "ZAR" }
];

let liveUpdateEnabled = false;
let isSyncing = false;

// ------------------
// Entry point
// ------------------
Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    Office.context.document.settings.set("Office.AutoShowTaskpaneWithDocument", true);
    Office.context.document.settings.saveAsync();
  }  
  initialiseUI();
});

function showSyncStatus() {
  const el = document.getElementById("syncStatus");
  if (el) el.classList.remove("hidden");
}

function hideSyncStatus() {
  const el = document.getElementById("syncStatus");
  if (el) el.classList.add("hidden");
}

// ------------------
// Init
// ------------------
let updateTimer = null;

function scheduleLiveUpdate() {
  if (!liveUpdateEnabled) return;

  clearTimeout(updateTimer);
  updateTimer = setTimeout(() => {
    generateQuote();
  }, 400);
}

function initialiseUI() {
  populateSalesRepDropdown();
  populateCompanyDropdown();
  populateCurrency();
  wireSalesRepChange();
  setDefaultQuoteDate();
  wireLiveUpdates();

  // Buttons 
  document
    .getElementById("import")
    .addEventListener("click", importDataFromDoc);

  document
    .getElementById("generateQuote")
    .addEventListener("click", generateQuote);

  // SANITY FIX: Check percentages immediately on startup
  validatePercentagesSubtle();
}

// ------------------
// Dropdowns
// ------------------
function populateSalesRepDropdown() {
  const select = document.getElementById("salesRep");
  const emailInput = document.getElementById("agentEmail");

  select.innerHTML = "";

  salesReps.forEach((rep, index) => {
    const opt = new Option(rep.name, rep.email);
    if (index === 0) opt.selected = true;
    select.append(opt);
  });

  emailInput.value = select.value;
  delete emailInput.dataset.userEdited;
}

function populateCompanyDropdown() {
  const select = document.getElementById("cliffordCompany");
  select.innerHTML = "";

  cliffordCompany.forEach((company, index) => {
    const option = new Option(company.name, company.name);
    if (index === 0) option.selected = true;
    select.append(option);
  });
}

function populateCurrency() {
  const select = document.getElementById("currency");
  select.innerHTML = "";

  currency.forEach((currency, index) => {
    const option = new Option(currency.name, currency.name);
    if (index === 0) option.selected = true;
    select.append(option);
  });
}

// ------------------
// Wiring
// ------------------
function wireSalesRepChange() {
  const select = document.getElementById("salesRep");
  const emailInput = document.getElementById("agentEmail");

  select.addEventListener("change", () => {
    if (!emailInput.dataset.userEdited) {
      emailInput.value = select.value;
    }
  });

  emailInput.addEventListener("input", () => {
    emailInput.dataset.userEdited = "true";
  });
}

function setDefaultQuoteDate() {
  const quoteDateInput = document.getElementById("quoteDate");
  if (!quoteDateInput.value) {
    quoteDateInput.value = new Date().toISOString().split("T")[0];
  }
}

function formatDateForWord(isoDate) {
  const date = new Date(isoDate);
  const day = date.getDate();
  const suffix =
    day % 10 === 1 && day !== 11 ? "st" :
    day % 10 === 2 && day !== 12 ? "nd" :
    day % 10 === 3 && day !== 13 ? "rd" : "th";

  const month = date.toLocaleString("en-GB", { month: "long" });
  const year = date.getFullYear();

  return `${day}${suffix} ${month} ${year}`;
}

function wireLiveUpdates() {
  const fields = document.querySelectorAll(
    "#quoteDate, #quoteId, #salesRep, #agentEmail, #cliffordCompany, " +
    "#customerCompany, #customerName, #address1, #address2, #address3, " +
    "#currency, #delivery"
  );

  fields.forEach(field => {
    field.addEventListener("change", scheduleLiveUpdate);
  });

  const pctFields = document.querySelectorAll("#deposit, #shipment, #signoff");
  pctFields.forEach(field => {
    field.addEventListener("blur", validatePercentagesSubtle);
    field.addEventListener("input", scheduleLiveUpdate);
  });
}

// IMPORT
async function importDataFromDoc() {
  await Word.run(async (context) => {
    const tagsToImport = [
      "quoteId", 
      "customerName", 
      "customerCompany", 
      "address1", 
      "address2", 
      "address3",
      "currency",
      "cliffordCompany",
      "salesRep",
      "agentEmail",
      "delivery", 
      "deposit",
      "shipment",
      "signoff"
    ];
    
    const controlMap = {};

    tagsToImport.forEach(tag => {
      const controls = context.document.contentControls.getByTag(tag);
      controls.load("items/text"); 
      controlMap[tag] = controls;
    });

    await context.sync();

    tagsToImport.forEach(tag => {
      const htmlElement = document.getElementById(tag);
      const docControls = controlMap[tag].items;

      if (htmlElement && docControls.length > 0) {
        const docValue = docControls[0].text;

        if (tag === "salesRep") {
          const foundRep = salesReps.find(rep => rep.name === docValue);
          if (foundRep) {
            htmlElement.value = foundRep.email;
          }
        } 
        else if (tag === "deposit" || tag === "shipment" || tag === "signoff") {
          if (!docValue || docValue.trim() === "") {
            htmlElement.value = "";
          } else {
            const match = docValue.match(/\d+/);
            htmlElement.value = match ? match[0] : "";
          }
        } 
        else {
          htmlElement.value = docValue;
        }
      }
    });

    // Run verification cycle after importing to set base status state
    validatePercentagesSubtle();
    console.log("Import complete.");
  });
}

// PERCENT VALIDATION ON FOCUS SHIFT
function validatePercentagesSubtle() {
  const errorDisplay = document.getElementById("validation-error");
  
  const depositEl = document.getElementById("deposit");
  const shipmentEl = document.getElementById("shipment");
  const signoffEl = document.getElementById("signoff");

  const depositVal = depositEl ? (depositEl.valueAsNumber || 0) : 0;
  const shipmentVal = shipmentEl ? (shipmentEl.valueAsNumber || 0) : 0;
  const signoffRaw = signoffEl ? signoffEl.value.trim() : "";

  let total = depositVal + shipmentVal;

  if (signoffRaw !== "") {
    const signoffVal = signoffEl ? (signoffEl.valueAsNumber || 0) : 0;
    total += signoffVal;
  }

  if (total !== 100) {
    if (errorDisplay) {
      errorDisplay.innerText = `Total allocations equal ${total}%. They should ideally equal exactly 100%.`;
      errorDisplay.className = "status-box error";
    }
  } else {
    if (errorDisplay) {
      errorDisplay.innerText = "";
      errorDisplay.className = "status-box";
    }
  }
}

// GENERATE
async function generateQuote() {
  if (isSyncing) return;

  isSyncing = true;
  liveUpdateEnabled = true;
  showSyncStatus();

  try {
    const depositVal = document.getElementById("deposit").value || "0";
    const shipmentVal = document.getElementById("shipment").value || "0";
    const signoffRaw = document.getElementById("signoff").value.trim();

    const formattedDepositText = `\u2022\t${depositVal}% non-refundable deposit with order.`;
    const formattedShipmentText = `\u2022\t${shipmentVal}% payable prior to shipment.`;
    const formattedSignoffText = signoffRaw !== "" ? `\u2022\t${signoffRaw}% payable at project Signoff.` : "\u200B";

    const data = {
      quoteDate: formatDateForWord(document.getElementById("quoteDate").value),
      quoteId: document.getElementById("quoteId").value,
      salesRep: document.getElementById("salesRep").selectedOptions[0].text,
      agentEmail: document.getElementById("agentEmail").value,
      cliffordCompany: document.getElementById("cliffordCompany").value,
      customerCompany: document.getElementById("customerCompany").value,
      customerName: document.getElementById("customerName").value,
      address1: document.getElementById("address1").value,
      address2: document.getElementById("address2").value,
      address3: document.getElementById("address3").value,
      currency: document.getElementById("currency").value,
      delivery: document.getElementById("delivery").value,
      deposit: formattedDepositText,
      shipment: formattedShipmentText,
      signoff: formattedSignoffText
    };

    await Word.run(async (context) => {
      const controlMap = {};

      for (const tag in data) {
        const controls = context.document.contentControls.getByTag(tag);
        controls.load("items");
        controlMap[tag] = controls;
      }

      await context.sync();

      for (const tag in controlMap) {
        controlMap[tag].items.forEach(cc => {
          cc.insertText(data[tag], Word.InsertLocation.replace);
        });
      }

      await context.sync();
    });
  }
  finally {
    isSyncing = false;
    hideSyncStatus();
  }
}