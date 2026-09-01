import { initializeApp } from
  "https://www.gstatic.com/firebasejs/12.18.0/firebase-app.js";

import {
  getFirestore,
  collection,
  getDocs,
  setDoc,
  deleteDoc,
  doc,
  writeBatch
} from
  "https://www.gstatic.com/firebasejs/12.18.0/firebase-firestore.js";


/* Firebase configuration */
const firebaseConfig = {
 apiKey: "AIzaSyBKH3gOWk_rG3_IxZuhj3SUGvN5DU_dZ9E",
  authDomain: "epaymentsystem.firebaseapp.com",
  projectId: "epaymentsystem",
  storageBucket: "epaymentsystem.firebasestorage.app",
  messagingSenderId: "627228868868",
  appId: "1:627228868868:web:83c5cddab040fdee981051",
  measurementId: "G-S9D37602ZR"
};


/* Firebase initialization */
const app =
  initializeApp(firebaseConfig);

const db =
  getFirestore(app);


/* Collections */
const COLLECTIONS = {
  enumerators: "enumerators",
  surveys: "surveys",
  payments: "payments",
  ddpoAllocations: "ddpoAllocations"
};


/* Fixed survey rate */
const SURVEY_RATE = 250;


/* Application memory */
let enumerators = [];
let surveys = [];
let payments = [];
let ddpoAllocations = [];


/* General helpers */
function clean(value) {
  return String(value ?? "").trim();
}

function normalizeMobile(value) {
  return clean(value)
    .replace(/\D/g, "")
    .slice(-10);
}
function money(value) {
  return Number(value || 0).toLocaleString(
    "en-IN",
    {
      minimumFractionDigits: 0,
      maximumFractionDigits: 2
    }
  );
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function makeId(prefix = "item") {
  if (
    window.crypto &&
    typeof crypto.randomUUID === "function"
  ) {
    return `${prefix}_${crypto.randomUUID()}`;
  }

  return `${prefix}_${Date.now()}_${Math.random()
    .toString(36)
    .slice(2)}`;
}

function parseNumber(value) {
  if (
    value === undefined ||
    value === null
  ) {
    return 0;
  }

  const text = String(value)
    .trim()
    .replace(/,/g, "")
    .replace(/[₹$]/g, "")
    .replace(/[^\d.-]/g, "");

  const number =
    Number.parseFloat(text);

  return Number.isFinite(number)
    ? number
    : 0;
}

function normalizeExcelHeader(value) {
  return clean(value)
    .toLowerCase()
    .replace(/[^a-z0-9]/g, "");
}

function normalizeMatchValue(value) {
  return clean(value)
    .toLowerCase()
    .replace(/[\s-]+/g, "");
}

function makeIdForText(value) {
  return clean(value)
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "_")
    .replace(/^_+|_+$/g, "");
}


/* Sorting */
function sortByEnumeratorId(a, b) {
  return clean(a.id).localeCompare(
    clean(b.id),
    undefined,
    {
      numeric: true,
      sensitivity: "base"
    }
  );
}

function sortByBoothNo(a, b) {
  return clean(a.boothNo).localeCompare(
    clean(b.boothNo),
    undefined,
    {
      numeric: true,
      sensitivity: "base"
    }
  );
}


/* Firebase functions */
async function saveToFirebase(
  collectionName,
  item,
  documentId = item.id
) {
  const finalId =
    clean(documentId);

  if (!finalId) {
    throw new Error(
      "Firebase document ID is missing."
    );
  }

  await setDoc(
    doc(
      db,
      collectionName,
      finalId
    ),
    {
      ...item,
      id: finalId,
      updatedAt:
        new Date().toISOString()
    },
    {
      merge: true
    }
  );
}

async function deleteFromFirebase(
  collectionName,
  id
) {
  await deleteDoc(
    doc(
      db,
      collectionName,
      String(id)
    )
  );
}

async function loadCollection(
  collectionName
) {
  const snapshot =
    await getDocs(
      collection(db, collectionName)
    );

  return snapshot.docs.map(
    (document) => ({
      id: document.id,
      ...document.data()
    })
  );
}

async function loadDataFromFirebase() {
  try {
   const [
  enumeratorData,
  surveyData,
  ddpoAllocationData,
  paymentData
] = await Promise.all([
  loadCollection(
    COLLECTIONS.enumerators
  ),
  loadCollection(
    COLLECTIONS.surveys
  ),
  loadCollection(
    COLLECTIONS.ddpoAllocations
    
  ),
  loadCollection(
    COLLECTIONS.payments
  )
]);

    enumerators =
      enumeratorData;

    surveys =
      surveyData;
    ddpoAllocations =
  ddpoAllocationData;

   payments = paymentData;

populateDdpoDateFilter();

refreshAll();
  } catch (error) {
    console.error(
      "Firebase loading error:",
      error
    );

    alert(
      "Firebase data could not be loaded. " +
      "Check Firebase configuration and rules."
    );
  }
}






/* Navigation */
function showPage(
  pageId,
  button
) {
  document.querySelectorAll(
    ".page"
  ).forEach((page) => {
    page.classList.remove(
      "active"
    );
  });

  document.querySelectorAll(
    ".tab-btn"
  ).forEach((tab) => {
    tab.classList.remove(
      "active"
    );
  });

  const page =
    document.getElementById(pageId);

  if (page) {
    page.classList.add("active");
  }

  if (button) {
    button.classList.add("active");
  }
}


/* Excel helper */
function readExcelRows(
  file,
  useHeaderRow = false
) {
  return new Promise(
    (resolve, reject) => {
      const reader =
        new FileReader();

      reader.onload =
        function (event) {
          try {
            const workbook =
              XLSX.read(
                event.target.result,
                {
                  type: "array",
                  cellText: true,
                  cellDates: false
                }
              );

            const sheetName =
              workbook.SheetNames[0];

            if (!sheetName) {
              throw new Error(
                "No worksheet found."
              );
            }

            const worksheet =
              workbook.Sheets[sheetName];

            const rows =
              XLSX.utils.sheet_to_json(
                worksheet,
                useHeaderRow
                  ? {
                      defval: "",
                      raw: false
                    }
                  : {
                      header: 1,
                      defval: "",
                      raw: false
                    }
              );

            resolve(rows);
          } catch (error) {
            reject(error);
          }
        };

      reader.onerror =
        function () {
          reject(
            new Error(
              "File could not be read."
            )
          );
        };

      reader.readAsArrayBuffer(file);
    }
  );
}


/* Enumerator Excel header value */
function getExcelValue(
  row,
  aliases
) {
  const keys =
    Object.keys(row);

  for (
    const alias
    of aliases
  ) {
    const wanted =
      normalizeExcelHeader(
        alias
      );

    const matchingKey =
      keys.find(
        (key) =>
          normalizeExcelHeader(
            key
          ) === wanted
      );

    if (
      matchingKey !== undefined &&
      row[matchingKey] !==
        undefined &&
      row[matchingKey] !== null
    ) {
      return clean(
        row[matchingKey]
      );
    }
  }

  return "";
}


/* Enumerator calculation */
function getEnumeratorData(id) {
  const enumerator =
    enumerators.find(
      (item) =>
        clean(item.id) ===
        clean(id)
    );

  if (!enumerator) {
    return null;
  }

  const relatedSurveys =
    surveys.filter(
      (item) =>
        clean(
          item.enumeratorId
        ) === clean(id)
    );

  const totalSurveys =
    relatedSurveys.reduce(
      (total, item) =>
        total +
        Number(item.count || 0),
      0
    );

  const earned =
    totalSurveys * SURVEY_RATE;

  const relatedPayments =
    payments.filter(
      (item) =>
        clean(
          item.enumeratorId
        ) === clean(id)
    );

  const totalPaid =
    relatedPayments.reduce(
      (total, item) =>
        total +
        Number(
          item.amount || 0
        ),
      0
    );

  return {
    ...enumerator,
    rate: SURVEY_RATE,
    totalSurveys,
    earned,
    totalPaid,
    balance:
      earned - totalPaid
  };
}


/* Add enumerator manually */
async function addEnumerator() {
  const id =
    clean(
      document.getElementById(
        "enumId"
      ).value
    );

  const name =
    clean(
      document.getElementById(
        "enumName"
      ).value
    );

  const mobile =
    clean(
      document.getElementById(
        "enumMobile"
      ).value
    );

  const boothNo =
    clean(
      document.getElementById(
        "enumBoothNo"
      ).value
    );

  const boothName =
    clean(
      document.getElementById(
        "enumBoothName"
      ).value
    );

  const bankAccountNo =
    clean(
      document.getElementById(
        "enumBankAccount"
      ).value
    );

  const ifscCode =
    clean(
      document.getElementById(
        "enumIfsc"
      ).value
    );

  const status =
    document.getElementById(
      "enumStatus"
    ).value;

  if (!id || !name) {
    alert(
      "Enumerator ID and Name are required."
    );
    return;
  }

  if (
    enumerators.some(
      (item) =>
        clean(item.id) === id
    )
  ) {
    alert(
      "Enumerator ID already exists."
    );
    return;
  }

  const item = {
    id,
    name,
    mobile,
    boothNo,
    boothName,
    bankAccountNo,
    ifscCode,
    status
  };

  try {
    await saveToFirebase(
      COLLECTIONS.enumerators,
      item
    );

    enumerators.push(item);

    refreshAll();

    alert(
      "Enumerator saved successfully."
    );
  } catch (error) {
    console.error(
      "Enumerator save error:",
      error
    );

    alert(
      "Enumerator could not be saved."
    );
  }
}


/* Import enumerators */
async function importEnumeratorsFromExcel() {
  const fileInput =
    document.getElementById(
      "enumeratorExcelFile"
    );

  const resultBox =
    document.getElementById(
      "importResult"
    );

  const file =
    fileInput.files[0];

  if (!file) {
    alert(
      "Please select an Enumerator Excel file."
    );
    return;
  }

  try {
    const rows =
      await readExcelRows(
        file,
        true
      );

    let added = 0;
    let updated = 0;
    let skipped = 0;
    let failed = 0;

    const errors = [];

    for (
      const [index, row]
      of rows.entries()
    ) {
      const excelRow =
        index + 2;

      const id =
        getExcelValue(
          row,
          [
            "Enumerator ID",
            "EnumeratorId",
            "ID",
            "Enum ID",
            "EnumId"
          ]
        );

      const name =
        getExcelValue(
          row,
          [
            "Name",
            "Enumerator Name",
            "EnumeratorName"
          ]
        );

      const mobile =
        getExcelValue(
          row,
          [
            "Mobile No",
            "Mobile Number",
            "Mobile",
            "Phone",
            "Phone Number"
          ]
        );

      const boothNo =
        getExcelValue(
          row,
          [
            "Booth No",
            "Booth Number",
            "BoothNo"
          ]
        );

      const boothName =
        getExcelValue(
          row,
          [
            "Booth Name",
            "BoothName"
          ]
        );

      const bankAccountNo =
        getExcelValue(
          row,
          [
            "Bank Account No",
            "Bank Account Number",
            "Account No",
            "Account Number",
            "BankAccountNo"
          ]
        );

      const ifscCode =
        getExcelValue(
          row,
          [
            "IFSC Code",
            "IFSC",
            "IFSCCode"
          ]
        );

      const statusValue =
        getExcelValue(
          row,
          [
            "Status"
          ]
        );

      if (!id || !name) {
        skipped++;

        errors.push(
          `Row ${excelRow}: ID and Name are required.`
        );

        continue;
      }

      const item = {
        id,
        name,
        mobile,
        boothNo,
        boothName,
        bankAccountNo,
        ifscCode,
        status:
          statusValue
            .toLowerCase() ===
          "inactive"
            ? "Inactive"
            : "Active"
      };

      try {
        const oldIndex =
          enumerators.findIndex(
            (oldItem) =>
              clean(oldItem.id) ===
              clean(id)
          );

        await saveToFirebase(
          COLLECTIONS.enumerators,
          item
        );

        if (oldIndex >= 0) {
          enumerators[oldIndex] =
            item;

          updated++;
        } else {
          enumerators.push(item);
          added++;
        }
      } catch (error) {
        failed++;

        errors.push(
          `Row ${excelRow}: Firebase save failed.`
        );
      }
    }

    refreshAll();

    resultBox.innerHTML = `
      <div class="import-success">
        Enumerator import completed.<br>
        Added: ${added}<br>
        Updated: ${updated}<br>
        Skipped: ${skipped}<br>
        Failed: ${failed}
      </div>

      ${
        errors.length
          ? `
            <div class="import-warning">
              ${errors
                .slice(0, 10)
                .map(
                  (error) =>
                    escapeHtml(error)
                )
                .join("<br>")}
            </div>
          `
          : ""
      }
    `;

    fileInput.value = "";
  } catch (error) {
    console.error(
      "Enumerator import error:",
      error
    );

    resultBox.innerHTML = `
      <div class="import-warning">
        Could not read Enumerator Excel file.
      </div>
    `;
  }
}


/* Enumerator Excel template */
function downloadEnumeratorTemplate() {
  const data = [
    {
      "Enumerator ID": "E001",
      Name: "Example Name",
      "Mobile No": "9876543210",
      "Booth No": "001",
      "Booth Name": "Example Booth",
      "Bank Account No": "123456789012",
      "IFSC Code": "SBIN0000001",
      Status: "Active"
    }
  ];

  const worksheet =
    XLSX.utils.json_to_sheet(data);

  const workbook =
    XLSX.utils.book_new();

  XLSX.utils.book_append_sheet(
    workbook,
    worksheet,
    "Enumerators"
  );

  XLSX.writeFile(
    workbook,
    "enumerator-template.xlsx"
  );
}


/* Find enumerator by name, mobile and booth */
function findEnumerator(
  name,
  mobile,
  boothName
) {
  const wantedName =
    normalizeMatchValue(name);

  const wantedMobile =
    normalizeMatchValue(mobile);

  const wantedBooth =
    normalizeMatchValue(boothName);

  return enumerators.find(
    (item) =>
      normalizeMatchValue(
        item.name
      ) === wantedName &&
      normalizeMatchValue(
        item.mobile
      ) === wantedMobile &&
      normalizeMatchValue(
        item.boothName
      ) === wantedBooth
  );
}

function findEnumeratorByMobileAndBooth(
  mobile,
  boothName
) {
  const wantedMobile =
    normalizeMatchValue(mobile);

  const wantedBooth =
    normalizeMatchValue(boothName);

  return enumerators.find(
    (item) =>
      normalizeMatchValue(
        item.mobile
      ) === wantedMobile &&
      normalizeMatchValue(
        item.boothName
      ) === wantedBooth
  );
}

async function importDDPOAllocations() {
  const fileInput =
    document.getElementById(
      "ddpoAllocationExcelFile"
    );

  const resultBox =
    document.getElementById(
      "ddpoAllocationImportResult"
    );

  const file =
    fileInput.files[0];

  if (!file) {
    alert(
      "Please select the DDPO list Excel file."
    );

    return;
  }

  try {
    const rows =
      await readExcelRows(
        file,
        false
      );

    if (rows.length < 2) {
      alert(
        "The DDPO list is empty."
      );

      return;
    }

    let added = 0;
    let updated = 0;
    let skipped = 0;
    let failed = 0;

    const errors = [];

    for (
      const [index, row]
      of rows.slice(1).entries()
    ) {
      const excelRow =
        index + 2;

      const date =
        clean(row[0]);

      const name =
        clean(row[1]);

      const mobile =
        clean(row[2]);

      const boothName =
        clean(row[3]);

      const releasedAmount =
        parseNumber(row[4]);

      if (
        !date &&
        !name &&
        !mobile &&
        !boothName &&
        !releasedAmount
      ) {
        continue;
      }

      if (
        !date ||
        !mobile ||
        !boothName ||
        !Number.isFinite(
          releasedAmount
        ) ||
        releasedAmount < 0
      ) {
        skipped++;

        errors.push(
          `Row ${excelRow}: Date, Mobile No, Booth Name and valid amount are required.`
        );

        continue;
      }

      const enumerator =
        findEnumeratorByMobileAndBooth(
          mobile,
          boothName
        );

      if (!enumerator) {
        skipped++;

        errors.push(
          `Row ${excelRow}: No matching enumerator found for ${mobile} and ${boothName}.`
        );

        continue;
      }

      const allocationId =
        [
          "ddpo",
          makeIdForText(date),
          makeIdForText(mobile),
          makeIdForText(boothName)
        ].join("_");

      const item = {
        id: allocationId,
        date,
        enumeratorId:
          enumerator.id,
        name:
          enumerator.name ||
          name,
        mobile:
          enumerator.mobile ||
          mobile,
        boothName:
          enumerator.boothName ||
          boothName,
        releasedAmount,
        source: "excel"
      };

      try {
        const oldIndex =
          ddpoAllocations.findIndex(
            (oldItem) =>
              oldItem.id ===
              allocationId
          );

        await saveToFirebase(
          COLLECTIONS.ddpoAllocations,
          item
        );

        if (oldIndex >= 0) {
          ddpoAllocations[oldIndex] =
            item;

          updated++;
        } else {
          ddpoAllocations.push(item);

          added++;
        }
      } catch (error) {
        failed++;

        errors.push(
          `Row ${excelRow}: Firebase save failed.`
        );
      }
    }

    refreshAll();

    resultBox.innerHTML = `
      <div class="import-success">
        DDPO list processed successfully.<br>
        Added: ${added}<br>
        Updated: ${updated}<br>
        Skipped: ${skipped}<br>
        Failed: ${failed}
      </div>

      ${
        errors.length
          ? `
            <div class="import-warning">
              ${errors
                .slice(0, 10)
                .map(
                  (error) =>
                    escapeHtml(error)
                )
                .join("<br>")}
            </div>
          `
          : ""
      }
    `;

    fileInput.value = "";
  } catch (error) {
    console.error(
      "DDPO upload error:",
      error
    );

    resultBox.innerHTML = `
      <div class="import-warning">
        Could not read the DDPO Excel file.
      </div>
    `;
  }
}

function renderDDPOAllocationTable() {
  const head =
    document.getElementById(
      "ddpoAllocationTableHead"
    );

  const body =
    document.getElementById(
      "ddpoAllocationTableBody"
    );

  const footer =
    document.getElementById(
      "ddpoAllocationTableFooter"
    );

  if (
    !head ||
    !body ||
    !footer
  ) {
    return;
  }

  if (
    !ddpoAllocations ||
    !ddpoAllocations.length
  ) {
    head.innerHTML = `
      <tr>
        <th>Name</th>
        <th>Mobile No</th>
        <th>Booth Name</th>
        <th>Total Released</th>
      </tr>
    `;

    body.innerHTML = `
      <tr>
        <td
          colspan="4"
          class="empty">
          No DDPO release records found.
        </td>
      </tr>
    `;

    footer.innerHTML = "";

    return;
  }

  const dateList =
    [...new Set(
      ddpoAllocations
        .map(
          (item) =>
            clean(item.date)
        )
        .filter(Boolean)
    )].sort();

  const grouped =
    {};

  ddpoAllocations.forEach(
    (item) => {
      const mobile =
        normalizeMatchValue(
          item.mobile
        );

      const booth =
        normalizeMatchValue(
          item.boothName
        );

      const key =
        `${mobile}__${booth}`;

      if (!grouped[key]) {
        grouped[key] = {
          name:
            item.name || "-",
          mobile:
            item.mobile || "-",
          boothName:
            item.boothName || "-",
          dates: {},
          total: 0
        };
      }

      const date =
        clean(item.date);

      const amount =
        Number(
          item.releasedAmount || 0
        );

      grouped[key].dates[date] =
        (grouped[key].dates[date] || 0) +
        amount;

      grouped[key].total +=
        amount;
    }
  );

  head.innerHTML = `
    <tr>
      <th>Name</th>
      <th>Mobile No</th>
      <th>Booth Name</th>

      ${dateList
        .map(
          (date) =>
            `<th>${escapeHtml(date)}</th>`
        )
        .join("")}

      <th>Total Released</th>
    </tr>
  `;

  body.innerHTML = "";

  const rows =
    Object.values(grouped).sort(
      (a, b) =>
        String(a.name)
          .localeCompare(
            String(b.name)
          )
    );

  const dateTotals =
    {};

  dateList.forEach(
    (date) => {
      dateTotals[date] = 0;
    }
  );

  let grandTotal = 0;

  rows.forEach(
    (item) => {
      let cells = "";

      dateList.forEach(
        (date) => {
          const amount =
            Number(
              item.dates[date] || 0
            );

          dateTotals[date] +=
            amount;

          cells += `
            <td>
              ₹${money(amount)}
            </td>
          `;
        }
      );

      grandTotal +=
        item.total;

      body.innerHTML += `
        <tr>
          <td>
            ${escapeHtml(
              item.name
            )}
          </td>

          <td>
            ${escapeHtml(
              item.mobile
            )}
          </td>

          <td>
            ${escapeHtml(
              item.boothName
            )}
          </td>

          ${cells}

          <td class="positive">
            ₹${money(item.total)}
          </td>
        </tr>
      `;
    }
  );

  footer.innerHTML = `
    <tr class="report-total">
      <td colspan="3">
        Total
      </td>

      ${dateList
        .map(
          (date) =>
            `
              <td>
                ₹${money(
                  dateTotals[date]
                )}
              </td>
            `
        )
        .join("")}

      <td>
        ₹${money(grandTotal)}
      </td>
    </tr>
  `;
}

function getDDPOReleasedForEnumerator(
  enumerator
) {
  if (
    !enumerator ||
    !Array.isArray(
      ddpoAllocations
    )
  ) {
    return 0;
  }

  const mobile =
    normalizeMatchValue(
      enumerator.mobile
    );

  const boothName =
    normalizeMatchValue(
      enumerator.boothName
    );

  return ddpoAllocations.reduce(
    (sum, item) => {
      const sameMobile =
        normalizeMatchValue(
          item.mobile
        ) === mobile;

      const sameBooth =
        normalizeMatchValue(
          item.boothName
        ) === boothName;

      if (
        sameMobile &&
        sameBooth
      ) {
        return sum +
          Number(
            item.releasedAmount || 0
          );
      }

      return sum;
    },
    0
  );
}
function calculateDDPODashboardTotals() {
  const totalEarned =
    surveys.reduce((total, item) => {
      return total +
        Number(item.count || 0) *
        SURVEY_RATE;
    }, 0);

  const totalDDPOReleased =
    ddpoAllocations.reduce((total, item) => {
      return total +
        Number(item.releasedAmount || 0);
    }, 0);

  const totalPaid =
    payments.reduce((total, item) => {
      return total +
        Number(item.amount || 0);
    }, 0);

  return {
    totalEarned,
    totalDDPOReleased,
    totalPaid,

    pendingFromEarned:
      totalEarned - totalPaid,

    pendingFromDDPO:
      totalDDPOReleased - totalPaid
  };
}
function renderDDPODashboardCards() {
  const totals =
    calculateDDPODashboardTotals();

  const earnedBox =
    document.getElementById(
      "totalEarned"
    );

  const paidBox =
    document.getElementById(
      "totalPaid"
    );

  const pendingEarnedBox =
    document.getElementById(
      "totalPending"
    );

  const releasedBox =
    document.getElementById(
      "dashboardTotalDDPOReleased"
    );

  const pendingDDPOBox =
    document.getElementById(
      "dashboardPendingDDPO"
    );

  if (earnedBox) {
    earnedBox.innerText =
      money(totals.totalEarned);
  }

  if (paidBox) {
    paidBox.innerText =
      money(totals.totalPaid);
  }

  if (pendingEarnedBox) {
    pendingEarnedBox.innerText =
      money(totals.pendingFromEarned);
  }

  if (releasedBox) {
    releasedBox.innerText =
      money(totals.totalDDPOReleased);
  }

  if (pendingDDPOBox) {
    pendingDDPOBox.innerText =
      money(totals.pendingFromDDPO);
  }
}
/* Survey Excel survey-count detector */
function getSurveyCountValue(row) {
  const keys =
    Object.keys(row);

  const key =
    keys.find(
      (item) => {
        const header =
          normalizeExcelHeader(
            item
          );

        return (
          header.includes("survey") ||
          header.includes("count") ||
          header.includes("numberof")
        );
      }
    );

  return key
    ? clean(row[key])
    : "";
}

function parseSurveyNumber(value) {
  const number =
    parseNumber(value);

  return number;
}


/* Create Survey Import ID */
function makeSurveyImportId(
  date,
  enumeratorId,
  boothName,
  surveyCount
) {
  return (
    "excel_" +
    [
      date,
      enumeratorId,
      boothName,
      surveyCount
    ]
      .map(clean)
      .join("_")
      .replace(
        /[^a-zA-Z0-9_-]/g,
        "_"
      )
  );
}


/* Import surveys */
async function importSurveysFromExcel() {
  const fileInput =
    document.getElementById(
      "surveyExcelFile"
    );

  const resultBox =
    document.getElementById(
      "surveyImportResult"
    );

  const file =
    fileInput.files[0];

  if (!file) {
    alert(
      "Please select a Survey Excel file."
    );
    return;
  }

  try {
    const rows =
      await readExcelRows(
        file,
        true
      );

    let added = 0;
    let updated = 0;
    let skipped = 0;
    let failed = 0;

    const errors = [];

    /*
     * This prevents duplicate rows in the
     * same Excel file from creating problems.
     *
     * If the same Date + Enumerator + Mobile
     * + Booth appears more than once, the
     * last Excel row replaces the previous row.
     */
    const importedSurveys =
      new Map();

    for (
      const [
        index,
        row
      ] of rows.entries()
    ) {
      const excelRow =
        index + 2;

      const date =
        getExcelValue(
          row,
          [
            "Date",
            "Survey Date"
          ]
        );

      const name =
        getExcelValue(
          row,
          [
            "Name",
            "Enumerator Name",
            "EnumeratorName"
          ]
        );

      const mobile =
        getExcelValue(
          row,
          [
            "Mobile No",
            "Mobile Number",
            "Mobile",
            "Phone",
            "Phone Number"
          ]
        );

      const boothName =
        getExcelValue(
          row,
          [
            "Booth Name",
            "BoothName",
            "Polling Booth",
            "Booth"
          ]
        );

      const surveyValue =
        getExcelValue(
          row,
          [
            "No Of Survey",
            "No. Of Survey",
            "Number Of Survey",
            "Survey Count",
            "Surveys",
            "Count"
          ]
        );

      const remarks =
        getExcelValue(
          row,
          [
            "Remarks",
            "Remark",
            "Comments"
          ]
        );

      const surveyCount =
        parseNumber(
          surveyValue
        );

      if (
        !date ||
        !name ||
        !mobile ||
        !boothName
      ) {
        skipped++;

        errors.push(
          `Row ${excelRow}: Date, Name, Mobile No and Booth Name are required.`
        );

        continue;
      }

      if (
        !Number.isFinite(
          surveyCount
        ) ||
        surveyCount < 0
      ) {
        skipped++;

        errors.push(
          `Row ${excelRow}: Survey count is invalid.`
        );

        continue;
      }

      /*
       * The survey Excel has no Enumerator ID.
       * Find the enumerator from existing data.
       */
      const normalizedName =
        normalizeMatchValue(
          name
        );

      const normalizedMobile =
        normalizeMatchValue(
          mobile
        );

      const normalizedBooth =
        normalizeMatchValue(
          boothName
        );

      const enumerator =
        enumerators.find(
          (item) =>
            normalizeMatchValue(
              item.name
            ) === normalizedName &&
            normalizeMatchValue(
              item.mobile
            ) === normalizedMobile &&
            normalizeMatchValue(
              item.boothName
            ) === normalizedBooth
        );

      if (!enumerator) {
        skipped++;

        errors.push(
          `Row ${excelRow}: Enumerator not found for Name: ${name}, Mobile: ${mobile}, Booth: ${boothName}.`
        );

        continue;
      }

      /*
       * Survey count is deliberately NOT included.
       *
       * Final ID:
       * Date_EnumeratorID_MobileNo_BoothName
       */
      const surveyId =
        [
          date,
          enumerator.id,
          mobile,
          boothName
        ]
          .map(clean)
          .join("_")
          .replace(
            /[^a-zA-Z0-9_-]/g,
            ""
          );

      /*
       * Last matching row in the Excel file
       * add the survey count of an earlier matching row.
       */
				const previousSurvey =
				importedSurveys.get(
    surveyId
  );

if (previousSurvey) {
  /*
   * Same Date + Enumerator +
   * Mobile + Booth found again
   * in this Excel file.
   *
   * Merge the survey counts.
   */
  previousSurvey.count +=
    surveyCount;

  previousSurvey.totalAmount =
    previousSurvey.count *
    SURVEY_RATE;

  previousSurvey.remarks =
    [
      previousSurvey.remarks,
      remarks
    ]
      .filter(Boolean)
      .join(" | ");

  importedSurveys.set(
    surveyId,
    previousSurvey
  );
} else {
  importedSurveys.set(
    surveyId,
    {
      id: surveyId,
      date,
      enumeratorId:
        enumerator.id,
      name:
        enumerator.name ||
        name,
      mobile:
        enumerator.mobile ||
        mobile,
      boothName:
        enumerator.boothName ||
        boothName,
      count: surveyCount,
      rate: SURVEY_RATE,
      totalAmount:
        surveyCount *
        SURVEY_RATE,
      remarks,
      source: "excel",
      excelRow
    }
  );
}
    }

    /*
     * Save one final row per survey ID.
     */
    for (
      const surveyItem of
        importedSurveys.values()
    ) {
      try {
        const existingIndex =
          surveys.findIndex(
            (item) =>
              clean(item.id) ===
              clean(surveyItem.id)
          );

        /*
         * This replaces the complete old
         * survey data. It does not add counts.
         */
        await saveToFirebase(
          COLLECTIONS.surveys,
          surveyItem
        );

        if (existingIndex >= 0) {
          surveys[existingIndex] =
            surveyItem;

          updated++;
        } else {
          surveys.push(
            surveyItem
          );

          added++;
        }
      } catch (error) {
        failed++;

        errors.push(
          `Excel row ${surveyItem.excelRow}: Firebase save failed.`
        );
      }
    }

    refreshAll();

    resultBox.innerHTML = `
      <div class="import-success">
        Survey import completed.<br>
        Added: ${added}<br>
        Replaced/Updated: ${updated}<br>
        Skipped: ${skipped}<br>
        Failed: ${failed}
      </div>

      ${
        errors.length
          ? `
            <div class="import-warning">
              ${errors
                .slice(0, 20)
                .map(
                  (error) =>
                    escapeHtml(error)
                )
                .join("<br>")}
            </div>
          `
          : ""
      }
    `;

    fileInput.value = "";
  } catch (error) {
    console.error(
      "Survey import error:",
      error
    );

    resultBox.innerHTML = `
      <div class="import-warning">
        Could not read the Survey Excel file.
      </div>
    `;
  }
}

/* Survey template */
function downloadSurveyTemplate() {
  const data = [
    {
      Date: "2026-08-27",
      Name: "Example Name",
      "Mobile No": "9876543210",
      "Booth Name": "Example Booth",
      "No. of Survey": 25,
      Remarks: ""
    }
  ];

  const worksheet =
    XLSX.utils.json_to_sheet(data);

  const workbook =
    XLSX.utils.book_new();

  XLSX.utils.book_append_sheet(
    workbook,
    worksheet,
    "Surveys"
  );

  XLSX.writeFile(
    workbook,
    "survey-template.xlsx"
  );
}


/* Add survey manually */
async function addSurvey() {
  const date =
    clean(
      document.getElementById(
        "surveyDate"
      ).value
    );

  const enumeratorId =
    clean(
      document.getElementById(
        "surveyEnumerator"
      ).value
    );

  const count =
    Number(
      document.getElementById(
        "surveyCount"
      ).value
    );

  const remarks =
    clean(
      document.getElementById(
        "surveyRemarks"
      ).value
    );

  if (
    !date ||
    !enumeratorId ||
    count <= 0
  ) {
    alert(
      "Date, Enumerator and Survey Count are required."
    );
    return;
  }

  const enumerator =
    enumerators.find(
      (item) =>
        clean(item.id) ===
        enumeratorId
    );

  if (!enumerator) {
    alert(
      "Enumerator not found."
    );
    return;
  }

  const item = {
    id: makeId("survey"),
    date,
    enumeratorId,
    name:
      enumerator.name || "",
    mobile:
      enumerator.mobile || "",
    boothName:
      enumerator.boothName || "",
    count,
    rate:
      SURVEY_RATE,
    totalAmount:
      count * SURVEY_RATE,
    remarks,
    source: "manual"
  };

  try {
    await saveToFirebase(
      COLLECTIONS.surveys,
      item
    );

    surveys.push(item);

    document.getElementById(
      "surveyCount"
    ).value = "";

    document.getElementById(
      "surveyRemarks"
    ).value = "";

    refreshAll();

    alert(
      "Survey saved successfully."
    );
  } catch (error) {
    console.error(
      "Survey save error:",
      error
    );

    alert(
      "Survey could not be saved."
    );
  }
}


/* Payment Excel position reader */
function getPaymentColumn(
  row,
  position
) {
  return clean(
    row[position] ?? ""
  );
}


/* Payment installment name */
function getInstallmentName(
  headerValue
) {
  const override =
    clean(
      document.getElementById(
        "installmentName"
      ).value
    );

  if (override) {
    return override;
  }

  return (
    clean(headerValue) ||
    "Installment"
  );
}


/* Stable installment ID */
function makeInstallmentId(
  installmentName
) {
  return (
    "installment_" +
    makeIdForText(
      installmentName
    )
  );
}


/* Stable payment document ID */
function makeInstallmentPaymentId(
  installmentId,
  enumeratorId
) {
  return (
    `${installmentId}__${makeIdForText(
      enumeratorId
    )}`
  );
}


/* Import one installment Excel */
async function importInstallmentFromExcel() {
  const fileInput =
    document.getElementById(
      "paymentExcelFile"
    );

  const resultBox =
    document.getElementById(
      "paymentImportResult"
    );

  const file =
    fileInput.files[0];

  if (!file) {
    alert(
      "Please select an installment Excel file."
    );
    return;
  }

  try {
    const rows =
      await readExcelRows(
        file,
        false
      );

    if (!rows.length) {
      alert(
        "The payment Excel file is empty."
      );
      return;
    }

    const header =
      rows[0] || [];

    if (header.length < 5) {
      alert(
        "The Excel file needs at least five columns:\n" +
        "Date, Name, Mobile No, Booth Name and Installment Amount."
      );
      return;
    }

    const installmentName =
      getInstallmentName(
        header[4]
      );

    const installmentId =
      makeInstallmentId(
        installmentName
      );

    const paymentRows =
      rows.slice(1);

    let added = 0;
    let updated = 0;
    let skipped = 0;
    let failed = 0;

    const errors = [];

    for (
      const [index, row]
      of paymentRows.entries()
    ) {
      const excelRow =
        index + 2;

      const date =
        getPaymentColumn(row, 0);

      const name =
        getPaymentColumn(row, 1);

      const mobile =
        getPaymentColumn(row, 2);

      const boothName =
        getPaymentColumn(row, 3);

      const amount =
        parseNumber(
          getPaymentColumn(row, 4)
        );

      if (
        !date ||
        !name ||
        !mobile ||
        !boothName
      ) {
        skipped++;

        errors.push(
          `Row ${excelRow}: Date, Name, Mobile No and Booth Name are required.`
        );

        continue;
      }

      if (
        !Number.isFinite(amount) ||
        amount <= 0
      ) {
        skipped++;

        errors.push(
          `Row ${excelRow}: Installment amount is invalid.`
        );

        continue;
      }

      const enumerator =
        findEnumerator(
          name,
          mobile,
          boothName
        );

      if (!enumerator) {
        skipped++;

        errors.push(
          `Row ${excelRow}: Matching enumerator not found for ${name}.`
        );

        continue;
      }

      const paymentId =
        makeInstallmentPaymentId(
          installmentId,
          enumerator.id
        );

      const item = {
        id: paymentId,
        date,
        enumeratorId:
          enumerator.id,
        name:
          enumerator.name ||
          name,
        mobile:
          enumerator.mobile ||
          mobile,
        boothName:
          enumerator.boothName ||
          boothName,
        amount,
        installmentId,
        installmentName,
        source: "installment_excel"
      };

      try {
        const oldIndex =
          payments.findIndex(
            (payment) =>
              clean(payment.id) ===
              clean(paymentId)
          );

        await saveToFirebase(
          COLLECTIONS.payments,
          item
        );

        if (oldIndex >= 0) {
          payments[oldIndex] =
            item;

          updated++;
        } else {
          payments.push(item);
          added++;
        }
      } catch (error) {
        failed++;

        errors.push(
          `Row ${excelRow}: Firebase save failed.`
        );
      }
    }

    refreshAll();

    resultBox.innerHTML = `
      <div class="import-success">
        Installment import completed.<br>
        Installment: ${escapeHtml(
          installmentName
        )}<br>
        Added: ${added}<br>
        Updated: ${updated}<br>
        Skipped: ${skipped}<br>
        Failed: ${failed}
      </div>

      ${
        errors.length
          ? `
            <div class="import-warning">
              ${errors
                .slice(0, 10)
                .map(
                  (error) =>
                    escapeHtml(error)
                )
                .join("<br>")}
            </div>
          `
          : ""
      }
    `;

    fileInput.value = "";

    document.getElementById(
      "installmentName"
    ).value = "";
  } catch (error) {
    console.error(
      "Installment import error:",
      error
    );

    resultBox.innerHTML = `
      <div class="import-warning">
        Could not read the installment Excel file.
      </div>
    `;
  }
}


/* Payment Excel template */
function downloadPaymentTemplate() {
  const data = [
    [
      "Date",
      "Name",
      "Mobile No",
      "Booth Name",
      "1st Installment"
    ],
    [
      "2026-08-27",
      "Example Name",
      "9876543210",
      "Example Booth",
      1000
    ]
  ];

  const worksheet =
    XLSX.utils.aoa_to_sheet(data);

  const workbook =
    XLSX.utils.book_new();

  XLSX.utils.book_append_sheet(
    workbook,
    worksheet,
    "Payments"
  );

  XLSX.writeFile(
    workbook,
    "payment-installment-template.xlsx"
  );
}


/* Add manual payment */
async function addPayment() {
  const date =
    clean(
      document.getElementById(
        "paymentDate"
      ).value
    );

  const enumeratorId =
    clean(
      document.getElementById(
        "paymentEnumerator"
      ).value
    );

  const amount =
    Number(
      document.getElementById(
        "paymentAmount"
      ).value
    );

  const reference =
    clean(
      document.getElementById(
        "paymentReference"
      ).value
    );

  const remarks =
    clean(
      document.getElementById(
        "paymentRemarks"
      ).value
    );

  if (
    !date ||
    !enumeratorId ||
    amount <= 0
  ) {
    alert(
      "Date, Enumerator and Amount are required."
    );
    return;
  }

  const data =
    getEnumeratorData(
      enumeratorId
    );

  if (
    data &&
    amount > data.balance
  ) {
    const proceed =
      confirm(
        "Payment is greater than pending balance. Continue?"
      );

    if (!proceed) {
      return;
    }
  }

  const item = {
    id: makeId("manual_payment"),
    date,
    enumeratorId,
    amount,
    reference,
    remarks,
    source: "manual"
  };

  try {
    await saveToFirebase(
      COLLECTIONS.payments,
      item
    );

    payments.push(item);

    document.getElementById(
      "paymentAmount"
    ).value = "";

    document.getElementById(
      "paymentReference"
    ).value = "";

    document.getElementById(
      "paymentRemarks"
    ).value = "";

    refreshAll();

    alert(
      "Manual payment saved successfully."
    );
  } catch (error) {
    console.error(
      "Manual payment error:",
      error
    );

    alert(
      "Manual payment could not be saved."
    );
  }
}


/* Get all installment names */
function getInstallmentNames() {
  return [
    ...new Map(
      payments
        .filter(
          (item) =>
            item.source ===
            "installment_excel" &&
            item.installmentId &&
            item.installmentName
        )
        .map(
          (item) => [
            item.installmentId,
            item.installmentName
          ]
        )
    )
  ];
}


/* Delete one installment */
async function deleteInstallment(
  installmentId,
  installmentName
) {
  const confirmed =
    confirm(
      `Delete ${installmentName}?\n\n` +
      "All payment entries in this installment " +
      "will be deleted."
    );

  if (!confirmed) {
    return;
  }

  try {
    const relatedPayments =
      payments.filter(
        (item) =>
          item.installmentId ===
          installmentId
      );

    for (const payment of relatedPayments) {
      await deleteFromFirebase(
        COLLECTIONS.payments,
        payment.id
      );
    }

    payments =
      payments.filter(
        (item) =>
          item.installmentId !==
          installmentId
      );

    refreshAll();

    alert(
      "Installment and all its entries deleted."
    );
  } catch (error) {
    console.error(
      "Installment delete error:",
      error
    );

    alert(
      "Installment could not be deleted."
    );
  }
}


/* Render installment payment table */
function renderInstallmentTable() {
  const header =
    document.getElementById(
      "installmentHeaderRow"
    );

  const body =
    document.getElementById(
      "installmentTableBody"
    );

  const footer =
    document.getElementById(
      "installmentFooter"
    );

  if (
    !header ||
    !body ||
    !footer
  ) {
    return;
  }

  const installmentList =
    getInstallmentNames();

  header.innerHTML = `
    <th>Date</th>
    <th>Name</th>
    <th>Mobile No</th>
    <th>Booth Name</th>
  `;

  installmentList.forEach(
    (
      [
        installmentId,
        installmentName
      ]
    ) => {
      header.innerHTML += `
        <th class="installment-header">
          <div class="installment-title">
            ${escapeHtml(
              installmentName
            )}
          </div>

         <button
  type="button"
  class="btn btn-red small-btn"
  onclick="deleteInstallment(
    '${installmentId}',
    '${escapeHtml(installmentName)}'
  )">
  Delete This Installment
</button>
        </th>
      `;
    }
  );

  header.innerHTML += `
    <th>Total Paid</th>
  `;

  /*
    Important change:
    Group only by enumeratorId.

    Do not group by date + enumeratorId.
    This keeps every installment for one
    enumerator in the same row.
  */
  const grouped =
    new Map();

  payments
    .filter(
      (item) =>
        item.source ===
        "installment_excel"
    )
    .forEach(
      (item) => {
        const key =
          clean(
            item.enumeratorId
          );

        if (!key) {
          return;
        }

        if (!grouped.has(key)) {
          grouped.set(
            key,
            {
              date:
                item.date || "",
              name:
                item.name || "",
              mobile:
                item.mobile || "",
              boothName:
                item.boothName || "",
              values: {},
              total: 0
            }
          );
        }

        const row =
          grouped.get(key);

        /*
          Keep the first available date.
          The date is displayed once only.
        */
        if (
          !row.date &&
          item.date
        ) {
          row.date =
            item.date;
        }

        row.values[
          item.installmentId
        ] =
          (
            row.values[
              item.installmentId
            ] || 0
          ) +
          Number(
            item.amount || 0
          );

        row.total +=
          Number(
            item.amount || 0
          );
      }
    );

  body.innerHTML = "";

  if (!grouped.size) {
    body.innerHTML = `
      <tr>
        <td
          colspan="${
            5 +
            installmentList.length
          }"
          class="empty">
          No installment records.
        </td>
      </tr>
    `;

    footer.innerHTML = "";

    return;
  }

  grouped.forEach(
    (row) => {
      let html = `
        <tr>
          <td>
            ${escapeHtml(
              row.date || "-"
            )}
          </td>

          <td>
            ${escapeHtml(
              row.name || "-"
            )}
          </td>

          <td>
            ${escapeHtml(
              row.mobile || "-"
            )}
          </td>

          <td>
            ${escapeHtml(
              row.boothName || "-"
            )}
          </td>
      `;

      installmentList.forEach(
        (
          [installmentId]
        ) => {
          const amount =
            row.values[
              installmentId
            ] || 0;

          html += `
            <td class="installment-amount">
              ${
                amount
                  ? `₹${money(amount)}`
                  : "-"
              }
            </td>
          `;
        }
      );

      html += `
          <td class="positive">
            ₹${money(row.total)}
          </td>
        </tr>
      `;

      body.innerHTML += html;
    }
  );

  const totals = {};
  let grandTotal = 0;

  payments
    .filter(
      (item) =>
        item.source ===
        "installment_excel"
    )
    .forEach(
      (item) => {
        const amount =
          Number(
            item.amount || 0
          );

        totals[
          item.installmentId
        ] =
          (
            totals[
              item.installmentId
            ] || 0
          ) + amount;

        grandTotal += amount;
      }
    );

  let footerHtml = `
    <tr class="report-total">
      <td colspan="4">
        Total
      </td>
  `;

  installmentList.forEach(
    (
      [installmentId]
    ) => {
      footerHtml += `
        <td class="installment-amount">
          ₹${money(
            totals[
              installmentId
            ] || 0
          )}
        </td>
      `;
    }
  );

  footerHtml += `
      <td class="positive">
        ₹${money(grandTotal)}
      </td>
    </tr>
  `;

  footer.innerHTML =
    footerHtml;
}


/* Render manual payment history */
function renderPaymentHistory() {
  const table =
    document.getElementById(
      "paymentHistoryBody"
    );

  if (!table) {
    return;
  }

  table.innerHTML = "";

  const manualPayments =
    payments.filter(
      (item) =>
        item.source !==
        "installment_excel"
    );

  if (!manualPayments.length) {
    table.innerHTML = `
      <tr>
        <td
          colspan="6"
          class="empty">
          No manual payment records.
        </td>
      </tr>
    `;

    return;
  }

  manualPayments.forEach(
    (item) => {
      const enumerator =
        enumerators.find(
          (entry) =>
            clean(entry.id) ===
            clean(item.enumeratorId)
        );

      table.innerHTML += `
        <tr>
          <td>${escapeHtml(item.date)}</td>
          <td>
            ${
              enumerator
                ? `${escapeHtml(enumerator.id)}
                   -
                   ${escapeHtml(enumerator.name)}`
                : escapeHtml(
                    item.enumeratorId
                  )
            }
          </td>
          <td>₹${money(item.amount)}</td>
          <td>${escapeHtml(item.reference || "-")}</td>
          <td>${escapeHtml(item.remarks || "-")}</td>
         <td>
  <button
    type="button"
    class="btn btn-red small-btn"
    onclick="deletePayment('${item.id}')">
    Delete
  </button>
</td>
        </tr>
      `;
    }
  );
}


/* Delete manual payment */
async function deletePayment(id) {
  const confirmDelete =
    confirm(
      "Delete this manual payment?"
    );

  if (!confirmDelete) {
    return;
  }

  try {
    await deleteFromFirebase(
      COLLECTIONS.payments,
      id
    );

    payments =
      payments.filter(
        (item) =>
          item.id !== id
      );

    refreshAll();

    alert(
      "Payment deleted."
    );
  } catch (error) {
    console.error(
      "Payment delete error:",
      error
    );

    alert(
      "Payment could not be deleted."
    );
  }
}


/* Update Enumerator dropdowns */
function updateEnumeratorSelects() {
  const surveySelect =
    document.getElementById(
      "surveyEnumerator"
    );

  const paymentSelect =
    document.getElementById(
      "paymentEnumerator"
    );

  if (
    !surveySelect ||
    !paymentSelect
  ) {
    return;
  }

  surveySelect.innerHTML =
    "<option value=''>Select Enumerator</option>";

  paymentSelect.innerHTML =
    "<option value=''>Select Enumerator</option>";

  [...enumerators]
    .sort(sortByEnumeratorId)
    .forEach(
      (item) => {
        const text =
          `${item.id} - ${item.name}`;

        surveySelect.innerHTML += `
          <option value="${escapeHtml(item.id)}">
            ${escapeHtml(text)}
          </option>
        `;

        paymentSelect.innerHTML += `
          <option value="${escapeHtml(item.id)}">
            ${escapeHtml(text)}
          </option>
        `;
      }
    );
}


/* Update booth filter */
function updateBoothFilter() {
  const select =
    document.getElementById(
      "boothFilter"
    );

  if (!select) {
    return;
  }

  const booths = [
    ...new Set(
      enumerators
        .map(
          (item) =>
            clean(item.boothNo)
        )
        .filter(Boolean)
    )
  ];

  booths.sort(
    (a, b) =>
      a.localeCompare(
        b,
        undefined,
        {
          numeric: true
        }
      )
  );

  select.innerHTML =
    "<option value=''>All Booths</option>";

  booths.forEach(
    (booth) => {
      select.innerHTML += `
        <option value="${escapeHtml(booth)}">
          ${escapeHtml(booth)}
        </option>
      `;
    }
  );
}


/* Render enumerators */
function renderEnumerators() {
  const table =
    document.getElementById(
      "enumeratorTableBody"
    );

  if (!table) {
    return;
  }

  const search =
    clean(
      document.getElementById(
        "searchEnumerator"
      ).value
    ).toLowerCase();

  const boothSearch =
    clean(
      document.getElementById(
        "searchBoothNo"
      ).value
    ).toLowerCase();

  const boothFilter =
    clean(
      document.getElementById(
        "boothFilter"
      ).value
    );

  const statusFilter =
    document.getElementById(
      "statusFilter"
    ).value;

  const data =
    [...enumerators]
      .sort(sortByBoothNo)
      .map(
        (item) =>
          getEnumeratorData(
            item.id
          )
      )
      .filter(
        (item) => {
          const searchable =
            [
              item.id,
              item.name,
              item.mobile,
              item.boothNo,
              item.boothName
            ]
              .map(clean)
              .join(" ")
              .toLowerCase();

          const matchesSearch =
            searchable.includes(
              search
            );

          const matchesBoothSearch =
            clean(item.boothNo)
              .toLowerCase()
              .includes(
                boothSearch
              );

          const matchesBoothFilter =
            !boothFilter ||
            clean(item.boothNo) ===
              boothFilter;

          let matchesStatus = true;

          if (
            statusFilter ===
            "Active"
          ) {
            matchesStatus =
              item.status ===
              "Active";
          }

          if (
            statusFilter ===
            "Inactive"
          ) {
            matchesStatus =
              item.status ===
              "Inactive";
          }

          if (
            statusFilter ===
            "Pending"
          ) {
            matchesStatus =
              item.balance > 0;
          }

          if (
            statusFilter ===
            "Paid"
          ) {
            matchesStatus =
              item.balance <= 0 &&
              item.earned > 0;
          }

          return (
            matchesSearch &&
            matchesBoothSearch &&
            matchesBoothFilter &&
            matchesStatus
          );
        }
      );

  table.innerHTML = "";

  if (!data.length) {
    table.innerHTML = `
      <tr>
        <td
          colspan="14"
          class="empty">
          No enumerators found.
        </td>
      </tr>
    `;

    return;
  }

  data.forEach(
    (item) => {
      table.innerHTML += `
        <tr>
          <td>${escapeHtml(item.id)}</td>
          <td>${escapeHtml(item.name)}</td>
          <td>${escapeHtml(item.mobile || "-")}</td>
          <td>${escapeHtml(item.boothNo || "-")}</td>
          <td>${escapeHtml(item.boothName || "-")}</td>
          <td>${escapeHtml(item.bankAccountNo || "-")}</td>
          <td>${escapeHtml(item.ifscCode || "-")}</td>

          <td>
            <span class="badge ${
              item.status === "Active"
                ? "active-badge"
                : "inactive-badge"
            }">
              ${escapeHtml(item.status)}
            </span>
          </td>

          <td>${item.totalSurveys}</td>
          <td>₹${money(item.rate)}</td>
          <td>₹${money(item.earned)}</td>
          <td>₹${money(item.totalPaid)}</td>

          <td class="${
            item.balance > 0
              ? "pending"
              : "positive"
          }">
            ₹${money(item.balance)}
          </td>

          <td>
            <button
              class="btn small-btn"
              onclick="viewEnumerator('${escapeHtml(item.id)}')">
              View
            </button>

            <button
              class="btn btn-orange small-btn"
              onclick="toggleStatus('${escapeHtml(item.id)}')">
              ${
                item.status === "Active"
                  ? "Inactive"
                  : "Active"
              }
            </button>

            <button
              class="btn btn-red small-btn"
              onclick="deleteEnumerator('${escapeHtml(item.id)}')">
              Delete
            </button>
          </td>
        </tr>
      `;
    }
  );
}


/* View Enumerator */
function viewEnumerator(id) {
  const item =
    getEnumeratorData(id);

  if (!item) {
    return;
  }

  alert(
    `ID: ${item.id}\n` +
    `Name: ${item.name}\n` +
    `Mobile: ${item.mobile || "-"}\n` +
    `Booth No: ${item.boothNo || "-"}\n` +
    `Booth Name: ${item.boothName || "-"}\n` +
    `Status: ${item.status}\n` +
    `Total Surveys: ${item.totalSurveys}\n` +
    `Total Earned: ₹${money(item.earned)}\n` +
    `Total Paid: ₹${money(item.totalPaid)}\n` +
    `Pending: ₹${money(item.balance)}`
  );
}


/* Render survey history */
function renderSurveyHistory() {
  const table =
    document.getElementById(
      "surveyHistoryBody"
    );

  if (!table) {
    return;
  }

  table.innerHTML = "";

  if (!surveys.length) {
    table.innerHTML = `
      <tr>
        <td
          colspan="9"
          class="empty">
          No survey records.
        </td>
      </tr>
    `;

    return;
  }

  [...surveys]
    .sort(
      (a, b) =>
        new Date(b.date) -
        new Date(a.date)
    )
    .forEach(
      (item) => {
        const enumerator =
          enumerators.find(
            (entry) =>
              clean(entry.id) ===
              clean(item.enumeratorId)
          );

        const rate =
          Number(item.rate) ||
          SURVEY_RATE;

        const amount =
          Number(item.totalAmount) ||
          Number(item.count || 0) *
          rate;

        table.innerHTML += `
          <tr>
            <td>${escapeHtml(item.date)}</td>
            <td>
              ${
                enumerator
                  ? `${escapeHtml(enumerator.id)}
                     -
                     ${escapeHtml(enumerator.name)}`
                  : escapeHtml(
                      item.enumeratorId
                    )
              }
            </td>
            <td>${escapeHtml(item.name || "-")}</td>
            <td>${escapeHtml(item.mobile || "-")}</td>
            <td>${escapeHtml(item.boothName || "-")}</td>
            <td>${Number(item.count || 0)}</td>
            <td>₹${money(rate)}</td>
            <td>₹${money(amount)}</td>
            <td>${escapeHtml(item.remarks || "-")}</td>
          </tr>
        `;
      }
    );
}


/* Dashboard */
function renderDashboard() {
  const data =
    enumerators.map(
      (item) =>
        getEnumeratorData(
          item.id
        )
    );

  const totalSurveys =
    data.reduce(
      (sum, item) =>
        sum + item.totalSurveys,
      0
    );

  const totalEarned =
    data.reduce(
      (sum, item) =>
        sum + item.earned,
      0
    );

  const totalPaid =
    data.reduce(
      (sum, item) =>
        sum + item.totalPaid,
      0
    );

 const totalPending =
  data.reduce(
    (sum, item) =>
      sum + Number(item.balance || 0),
    0
  );

const inactivePendingAmount =
  data
    .filter(
      (item) =>
        item.status ===
          "Inactive" &&
        item.balance > 0
    )
    .reduce(
      (sum, item) =>
        sum + item.balance,
      0
    );
  
  document.getElementById(
    "totalEnumerators"
  ).innerText =
    enumerators.length;

  document.getElementById(
    "totalSurveys"
  ).innerText =
    totalSurveys;

  document.getElementById(
    "totalEarned"
  ).innerText =
    `₹${money(totalEarned)}`;

  document.getElementById(
    "totalPaid"
  ).innerText =
    `₹${money(totalPaid)}`;

  document.getElementById(
    "totalPending"
  ).innerText =
    `₹${money(totalPending)}`;
  
  document.getElementById(
  "inactivePendingAmount"
).innerText =
  `₹${money(inactivePendingAmount)}`;

/*start last card of dashboard */

	
	
/*end of last card of Dashboard code */

  const table =
    document.getElementById(
      "inactivePendingTable"
    );

  const pending =
    data.filter(
      (item) =>
        item.status ===
          "Inactive" &&
        item.balance > 0
    );

  table.innerHTML = "";

  if (!pending.length) {
    table.innerHTML = `
      <tr>
        <td
          colspan="6"
          class="empty">
          No inactive enumerator has pending payment.
        </td>
      </tr>
    `;

    return;
  }

  pending.forEach(
    (item) => {
      table.innerHTML += `
        <tr>
          <td>${escapeHtml(item.id)}</td>
          <td>${escapeHtml(item.name)}</td>
          <td>${item.totalSurveys}</td>
          <td>₹${money(item.earned)}</td>
          <td>₹${money(item.totalPaid)}</td>
          <td class="pending">
            ₹${money(item.balance)}
          </td>
        </tr>
      `;
    }
  );
}



/* Reports */
function getFilteredReportData() {
  const filter =
    document.getElementById(
      "reportFilter"
    ).value;

  let data = enumerators.map((enumerator) => {
    const reportItem =
      getEnumeratorData(enumerator.id);

    const fundReleased =
      getDDPOReleasedForEnumerator(
        enumerator
      );

    const paid =
      Number(reportItem.totalPaid || 0);

    const pendingFromEarned =
      Number(reportItem.earned || 0) - paid;

    const pendingFromFundReleased =
      fundReleased - paid;

    return {
      ...reportItem,
      fundReleased,
      pendingFromEarned,
      pendingFromFundReleased
    };
  });

  if (filter === "Active") {
    data = data.filter(
      (item) => item.status === "Active"
    );
  }

  if (filter === "Inactive") {
    data = data.filter(
      (item) => item.status === "Inactive"
    );
  }

  if (filter === "Pending") {
    data = data.filter(
      (item) =>
        item.pendingFromEarned !== 0 ||
        item.pendingFromFundReleased !== 0
    );
  }

  if (filter === "InactivePending") {
    data = data.filter(
      (item) =>
        item.status === "Inactive" &&
        (
          item.pendingFromEarned !== 0 ||
          item.pendingFromFundReleased !== 0
        )
    );
  }

  return data.sort(sortByBoothNo);
}


function getDdpoDate(item) {
  return String(
    item.date || ""
  ).trim();
}

function populateDdpoDateFilter() {
  const select =
    document.getElementById(
      "ddpoDateFilter"
    );

  if (!select) {
    return;
  }

  const dates =
    [
      ...new Set(
        ddpoAllocations
          .map(getDdpoDate)
          .filter(Boolean)
      )
    ].sort();

  select.innerHTML = `
    <option value="All">
      All Dates
    </option>
  `;

  dates.forEach((date) => {
    select.innerHTML += `
      <option value="${escapeHtml(date)}">
        ${escapeHtml(date)}
      </option>
    `;
  });
}

function renderDdpoReport() {
  const head =
    document.getElementById(
      "reportTableHead"
    );

  const body =
    document.getElementById(
      "reportTableBody"
    );

  const footer =
    document.getElementById(
      "reportFooter"
    );

  if (!head || !body) {
    return;
  }

  const selectedDate =
    document.getElementById(
      "ddpoDateFilter"
    )?.value || "All";

  const rows =
    ddpoAllocations.filter(
      (item) =>
        selectedDate === "All" ||
        getDdpoDate(item) === selectedDate
    );

  head.innerHTML = `
    <tr>
      <th>Name</th>
      <th>Mobile No</th>
      <th>Booth Name</th>
      <th>Bank Account No</th>
      <th>IFSC Code</th>
      <th>Amount of Fund Release</th>
    </tr>
  `;

  body.innerHTML = rows
    .map((item) => {
      const enumerator =
        enumerators.find(
          (person) =>
            normalizeMatchValue(
              person.mobile
            ) ===
              normalizeMatchValue(
                item.mobile
              ) &&
            normalizeMatchValue(
              person.boothName
            ) ===
              normalizeMatchValue(
                item.boothName
              )
        );

      return `
        <tr>
          <td>
            ${escapeHtml(
              item.name || "-"
            )}
          </td>

          <td>
            ${escapeHtml(
              item.mobile || "-"
            )}
          </td>

          <td>
            ${escapeHtml(
              item.boothName || "-"
            )}
          </td>

          <td>
            ${escapeHtml(
              enumerator?.bankAccountNo ||
              "-"
            )}
          </td>

          <td>
            ${escapeHtml(
              enumerator?.ifscCode ||
              "-"
            )}
          </td>

          <td>
            ₹${money(
              Number(
                item.releasedAmount
              ) || 0
            )}
          </td>
        </tr>
      `;
    })
    .join("");

  if (!rows.length) {
    body.innerHTML = `
      <tr>
        <td colspan="6">
          No DDPO records found.
        </td>
      </tr>
    `;
  }

  if (footer) {
    const total =
      rows.reduce(
        (sum, item) =>
          sum +
          (Number(
            item.releasedAmount
          ) || 0),
        0
      );

    footer.innerHTML = `
      <tr>
        <th colspan="5">
          Total Fund Released
        </th>
        <th>
          ₹${money(total)}
        </th>
      </tr>
    `;
  }
}

function handleReportFilterChange() {
  const filter =
    document.getElementById(
      "reportFilter"
    )?.value;

  const dateFilter =
    document.getElementById(
      "ddpoDateFilter"
    );

  if (filter === "DDPO") {
    if (dateFilter) {
      dateFilter.style.display =
        "inline-block";
    }

    populateDdpoDateFilter();
    renderDdpoReport();

    return;
  }

  if (dateFilter) {
    dateFilter.style.display =
      "none";
  }

  renderReports();
}

function renderEnumeratorReportHeader() {
  const head =
    document.getElementById(
      "reportTableHead"
    );

  if (!head) {
    return;
  }

  head.innerHTML = `
    <tr>
      <th>ID</th>
      <th>Name</th>
      <th>Status</th>
      <th>Total Surveys</th>
      <th>Rate</th>
      <th>Total Earned</th>
      <th>Fund Released</th>
      <th>Total Paid</th>
      <th>Pending from Earned</th>
      <th>Pending from Fund Released</th>
    </tr>
  `;
}

function renderReports() {
  const filter =
    document.getElementById(
      "reportFilter"
    )?.value || "All";

  if (filter === "DDPO") {
    renderDdpoReport();
    return;
  }

  renderEnumeratorReportHeader();

  const table =
    document.getElementById(
      "reportTableBody"
    );

  const footer =
    document.getElementById(
      "reportFooter"
    );

  if (!table) {
    return;
  }

  const data =
    getFilteredReportData();

  // Keep the rest of your existing
  // normal report code below this point.


  table.innerHTML =
    data
      .map(
        (item) => `
          <tr>
            <td>
              ${escapeHtml(
                item.id || "-"
              )}
            </td>

            <td>
              ${escapeHtml(
                item.name || "-"
              )}
            </td>

            <td>
              ${escapeHtml(
                item.status || "-"
              )}
            </td>

            <td>
              ${Number(
                item.totalSurveys || 0
              )}
            </td>

            <td>
              ₹${money(
                item.rate || 0
              )}
            </td>

            <td>
              ₹${money(
                item.earned || 0
              )}
            </td>

            <td>
              ₹${money(
                item.fundReleased || 0
              )}
            </td>

            <td>
              ₹${money(
                item.totalPaid || 0
              )}
            </td>

           <td class="${
  item.pendingFromEarned < 0
    ? "pending"
    : "positive"
}">
  ${money(item.pendingFromEarned)}
</td>

<td class="${
  item.pendingFromFundReleased < 0
    ? "pending"
    : "positive"
}">
  ${money(item.pendingFromFundReleased)}
</td>
          </tr>
        `
      )
      .join("");

  const totalSurveys =
    data.reduce(
      (sum, item) =>
        sum +
        Number(
          item.totalSurveys || 0
        ),
      0
    );

  const totalEarned =
    data.reduce(
      (sum, item) =>
        sum +
        Number(
          item.earned || 0
        ),
      0
    );

  const totalFundReleased =
    data.reduce(
      (sum, item) =>
        sum +
        Number(
          item.fundReleased || 0
        ),
      0
    );

  const totalPaid =
    data.reduce(
      (sum, item) =>
        sum +
        Number(
          item.totalPaid || 0
        ),
      0
    );

  const totalPendingFromEarned =
    data.reduce(
      (sum, item) =>
        sum +
        Number(
          item.pendingFromEarned || 0
        ),
      0
    );

  const totalPendingFromFundReleased =
    data.reduce(
      (sum, item) =>
        sum +
        Number(
          item.pendingFromFundReleased || 0
        ),
      0
    );

  if (footer) {
    footer.innerHTML = `
      <tr class="report-total">
        <td colspan="3">
          TOTAL
        </td>

        <td>
          ${totalSurveys}
        </td>

        <td>
          -
        </td>

        <td>
          ₹${money(
            totalEarned
          )}
        </td>

        <td>
          ₹${money(
            totalFundReleased
          )}
        </td>

        <td>
          ₹${money(
            totalPaid
          )}
        </td>

        <td>
          ₹${money(
            totalPendingFromEarned
          )}
        </td>

        <td>
          ₹${money(
            totalPendingFromFundReleased
          )}
        </td>
      </tr>
    `;
  }
}

/* CSV / Excel export */

function exportEnumeratorReport() {
  const data =
    getFilteredReportData();

  const rows = [
    [
      "Enumerator ID",
      "Name",
      "Mobile",
      "Bank Account No",
      "IFSC Code",
      "Booth No",
      "Booth Name",
      "Status",
      "Total Surveys",
      "Rate",
      "Total Earned",
      "DDPO Released",
      "Total Paid",
      "Pending from Earned",
      "Pending from DDPO"
    ]
  ];

  data.forEach((item) => {
    rows.push([
      item.id || "",
      item.name || "",
      item.mobile || "",
      item.bankAccountNo || "",
      item.ifscCode || "",
      item.boothNo || "",
      item.boothName || "",
      item.status || "",
      item.totalSurveys || 0,
      item.rate || 0,
      item.earned || 0,
      item.fundReleased || 0,
      item.totalPaid || 0,
      item.pendingFromEarned || 0,
      item.pendingFromFundReleased || 0
    ]);
  });

  const worksheet =
    XLSX.utils.aoa_to_sheet(rows);

  const workbook =
    XLSX.utils.book_new();

  XLSX.utils.book_append_sheet(
    workbook,
    worksheet,
    "Enumerator Report"
  );

  const reportFilter =
  document.getElementById(
    "reportFilter"
  );

const selectedReportName =
  reportFilter?.options[
    reportFilter.selectedIndex
  ]?.text || "Enumerator Report";

const safeFileName =
  selectedReportName
    .trim()
    .replace(
      /[\\/:*?"<>|]/g,
      ""
    )
    .replace(
      /\s+/g,
      "-"
    );

XLSX.writeFile(
  workbook,
  `${safeFileName}.xlsx`
);
}
function exportCSV() {
  const filter =
    document.getElementById(
      "reportFilter"
    )?.value;

  if (filter === "DDPO") {
    exportDdpoReport();
    return;
  }

  exportEnumeratorReport();
}

function exportDdpoReport() {
  const selectedDate =
    document.getElementById(
      "ddpoDateFilter"
    )?.value || "All";

  const allDates =
    [
      ...new Set(
        ddpoAllocations
          .map(
            (item) =>
              clean(item.date || "")
          )
          .filter(Boolean)
      )
    ].sort();

  const dates =
    selectedDate === "All"
      ? allDates
      : allDates.filter(
          (date) =>
            date === selectedDate
        );

  const grouped = new Map();

  ddpoAllocations.forEach(
    (allocation) => {
      const releaseDate =
        clean(
          allocation.date || ""
        );

      if (
        selectedDate !== "All" &&
        releaseDate !== selectedDate
      ) {
        return;
      }

      const mobile =
        clean(
          allocation.mobile || ""
        );

      const boothName =
        clean(
          allocation.boothName || ""
        );

      const enumerator =
        enumerators.find(
          (item) =>
            normalizeMatchValue(
              item.mobile
            ) ===
              normalizeMatchValue(
                mobile
              ) &&
            normalizeMatchValue(
              item.boothName
            ) ===
              normalizeMatchValue(
                boothName
              )
        );

      const key =
        clean(
          allocation.enumeratorId ||
          `${mobile}_${boothName}`
        );

      if (!grouped.has(key)) {
        grouped.set(key, {
          id:
            enumerator?.id ||
            allocation.enumeratorId ||
            "",

          name:
            allocation.name ||
            enumerator?.name ||
            "",

          mobile,

          boothName,

          bankAccountNo:
            enumerator?.bankAccountNo ||
            "",

          ifscCode:
            enumerator?.ifscCode ||
            "",

          amounts: {}
        });
      }

      const row =
        grouped.get(key);

      row.amounts[releaseDate] =
        (row.amounts[releaseDate] || 0) +
        (Number(
          allocation.releasedAmount
        ) || 0);
    }
  );

  const rows = [
    [
      "ID",
      "Name",
      "Mobile No",
      "Booth Name",
      "Bank Account No",
      "IFSC Code",
      ...dates.map(
        (date, index) =>
          `Release Amount ${index + 1}`
      )
    ]
  ];

  grouped.forEach((row) => {
    rows.push([
      row.id,
      row.name,
      row.mobile,
      row.boothName,
      row.bankAccountNo,
      row.ifscCode,
      ...dates.map(
        (date) =>
          row.amounts[date] || 0
      )
    ]);
  });

  const worksheet =
    XLSX.utils.aoa_to_sheet(rows);

  const workbook =
    XLSX.utils.book_new();

  XLSX.utils.book_append_sheet(
    workbook,
    worksheet,
    "DDPO Fund Release"
  );

  const fileDate =
    selectedDate === "All"
      ? "all-dates"
      : selectedDate.replaceAll(
          "/",
          "-"
        );

  XLSX.writeFile(
    workbook,
    `ddpo-fund-release-${fileDate}.xlsx`
  );
}

/* Change status */
async function toggleStatus(id) {
  const enumerator =
    enumerators.find(
      (item) =>
        clean(item.id) ===
        clean(id)
    );

  if (!enumerator) {
    return;
  }

  const updated = {
    ...enumerator,
    status:
      enumerator.status ===
      "Active"
        ? "Inactive"
        : "Active"
  };

  try {
    await saveToFirebase(
      COLLECTIONS.enumerators,
      updated
    );

    enumerators =
      enumerators.map(
        (item) =>
          clean(item.id) ===
          clean(id)
            ? updated
            : item
      );

    refreshAll();
  } catch (error) {
    console.error(
      "Status update error:",
      error
    );

    alert(
      "Status could not be updated."
    );
  }
}


/* Delete one Enumerator */
async function deleteEnumerator(id) {
  const enumerator =
    enumerators.find(
      (item) =>
        clean(item.id) ===
        clean(id)
    );

  if (!enumerator) {
    alert(
      "Enumerator not found."
    );
    return;
  }

  const confirmDelete =
    confirm(
      `Delete ${enumerator.id} - ${enumerator.name}?\n\n` +
      "This also deletes related surveys and payments."
    );

  if (!confirmDelete) {
    return;
  }

  try {
    await deleteFromFirebase(
      COLLECTIONS.enumerators,
      id
    );

    const surveySnapshot =
      await getDocs(
        collection(
          db,
          COLLECTIONS.surveys
        )
      );

    const paymentSnapshot =
      await getDocs(
        collection(
          db,
          COLLECTIONS.payments
        )
      );

    const batch =
      writeBatch(db);

    surveySnapshot.forEach(
      (document) => {
        const data =
          document.data();

        if (
          clean(
            data.enumeratorId
          ) === clean(id)
        ) {
          batch.delete(
            document.ref
          );
        }
      }
    );

    paymentSnapshot.forEach(
      (document) => {
        const data =
          document.data();

        if (
          clean(
            data.enumeratorId
          ) === clean(id)
        ) {
          batch.delete(
            document.ref
          );
        }
      }
    );

    await batch.commit();

    enumerators =
      enumerators.filter(
        (item) =>
          clean(item.id) !==
          clean(id)
      );

    surveys =
      surveys.filter(
        (item) =>
          clean(
            item.enumeratorId
          ) !== clean(id)
      );

    payments =
      payments.filter(
        (item) =>
          clean(
            item.enumeratorId
          ) !== clean(id)
      );

    refreshAll();

    alert(
      "Enumerator and related records deleted."
    );
  } catch (error) {
    console.error(
      "Enumerator delete error:",
      error
    );

    alert(
      "Enumerator could not be deleted."
    );
  }
}


/* Delete all collection documents */
async function deleteEntireCollection(
  collectionName
) {
  const snapshot =
    await getDocs(
      collection(db, collectionName)
    );

  const documents =
    snapshot.docs;

  for (
    let index = 0;
    index < documents.length;
    index += 500
  ) {
    const batch =
      writeBatch(db);

    documents
      .slice(
        index,
        index + 500
      )
      .forEach(
        (document) => {
          batch.delete(
            document.ref
          );
        }
      );

    await batch.commit();
  }
}




/* Set today's date */
function setToday() {
  const today =
    new Date()
      .toISOString()
      .split("T")[0];

  const surveyDate =
    document.getElementById(
      "surveyDate"
    );

  const paymentDate =
    document.getElementById(
      "paymentDate"
    );

  if (surveyDate) {
    surveyDate.value =
      today;
  }

  if (paymentDate) {
    paymentDate.value =
      today;
  }
}


/* Refresh all visible data */
function refreshAll() {
  updateBoothFilter();
  updateEnumeratorSelects();
  renderEnumerators();
  renderSurveyHistory();
  renderInstallmentTable();
  renderPaymentHistory();
  renderDashboard();
  renderDDPOAllocationTable();
  renderDDPODashboardCards();
  renderReports();
}


/* Expose HTML onclick functions */
window.showPage =
  showPage;

window.addEnumerator =
  addEnumerator;

window.importEnumeratorsFromExcel =
  importEnumeratorsFromExcel;

window.downloadEnumeratorTemplate =
  downloadEnumeratorTemplate;

window.importSurveysFromExcel =
  importSurveysFromExcel;

window.downloadSurveyTemplate =
  downloadSurveyTemplate;

window.addSurvey =
  addSurvey;

window.importInstallmentFromExcel =
  importInstallmentFromExcel;

window.downloadPaymentTemplate =
  downloadPaymentTemplate;

window.addPayment =
  addPayment;

window.deletePayment =
  deletePayment;

window.deleteInstallment =
  deleteInstallment;

window.toggleStatus =
  toggleStatus;

window.viewEnumerator =
  viewEnumerator;

window.deleteEnumerator =
  deleteEnumerator;


window.renderEnumerators =
  renderEnumerators;

window.renderReports =
  renderReports;

window.exportCSV =
  exportCSV;
window.importDDPOAllocations =
  importDDPOAllocations;
window.handleReportFilterChange =
  handleReportFilterChange;

window.renderDdpoReport =
  renderDdpoReport;

/* Start */
setToday();
loadDataFromFirebase();
