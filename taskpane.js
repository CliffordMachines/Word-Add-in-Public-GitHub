/* global Office, Word */
// v1.0.0.30 
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
  { name: "Clifford North America" },
  { name: "AWT" }
];

let liveUpdateEnabled = false;
let isSyncing = false;

// ------------------
// Entry point
// ------------------
Office.onReady((info) => {
  // Check if the script is running inside Microsoft Word
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

  // Default email
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
    "#customerCompany, #customerName, #address1, #address2, #address3, #delivery"
  );

  fields.forEach(field => {
    field.addEventListener("change", scheduleLiveUpdate);
  });
}


// IMPORT
async function importDataFromDoc() {
await Word.run(async (context) => {
    // 1. Define the tags we want to pull back from the document
    const tagsToImport = [
      "quoteId", 
      "customerName", 
      "customerCompany", 
      "address1", 
      "address2", 
      "address3",
      "delivery"
    ];
    
    const controlMap = {};

    // 2. Queue up the search and property load for each tag
    tagsToImport.forEach(tag => {
      const controls = context.document.contentControls.getByTag(tag);
      controls.load("items/text"); 
      controlMap[tag] = controls;
    });

    // 3. One sync to grab all data in a single handshake
    await context.sync();

    // 4. Map the document data back to your form inputs
    tagsToImport.forEach(tag => {
      const htmlElement = document.getElementById(tag);
      const docControls = controlMap[tag].items;

      if (htmlElement && docControls.length > 0) {
        htmlElement.value = docControls[0].text;
      }
    });

    console.log("Import complete.");
  });
}


async function generateQuote() {
  // Prevent re-entry while a sync is already running
  if (isSyncing) return;

  isSyncing = true;
  liveUpdateEnabled = true;
  showSyncStatus();

  try {
    const data = {
      quoteDate: formatDateForWord(
        document.getElementById("quoteDate").value
      ),
      quoteId: document.getElementById("quoteId").value,
      salesRep: document.getElementById("salesRep").selectedOptions[0].text,
      agentEmail: document.getElementById("agentEmail").value,
      cliffordCompany: document.getElementById("cliffordCompany").value,
      customerCompany: document.getElementById("customerCompany").value,
      customerName: document.getElementById("customerName").value,
      address1: document.getElementById("address1").value,
      address2: document.getElementById("address2").value,
      address3: document.getElementById("address3").value,
      delivery: document.getElementById("delivery").value
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

	//Set document Title 
	context.document.properties.title = data.quoteId;
	context.document.properties.subject = "";
	context.document.properties.category = "";

      await context.sync();
    });
  }
  finally {
    isSyncing = false;
    hideSyncStatus();
  }
}
