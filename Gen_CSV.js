fetch("https://solvnet.lightning.force.com/aura?r=77&ui-analytics-reporting-runpage.ReportPage.runReport=1", {
  method: "POST",
  headers: {
    "content-type": "application/x-www-form-urlencoded; charset=UTF-8"
  },
  body: "message=%7B%22actions%22%3A%5B%7B%22id%22%3A%223264%3Ba%22%2C%22descriptor%22%3A%22serviceComponent%3A%2F%2Fui.analytics.reporting.runpage.ReportPageController%2FACTION%24runReport%22%2C%22callingDescriptor%22%3A%22UNKNOWN%22%2C%22params%22%3A%7B%22reportId%22%3A%2200OWQ000005mZ5R2AU%22%2C%22reportMetadata%22%3A%22%7B%5C%22reportMetadata%5C%22%3A%7B%5C%22aggregates%5C%22%3A%5B%5C%22RowCount%5C%22%5D%2C%5C%22chart%5C%22%3A%7B%7D%2C%5C%22crossFilters%5C%22%3A%5B%5D%2C%5C%22currency%5C%22%3A%5C%22USD%5C%22%2C%5C%22dashboardSetting%5C%22%3Anull%2C%5C%22description%5C%22%3A%5C%22report_for_Mytime%5C%22%2C%5C%22detailColumns%5C%22%3A%5B%5C%22Case.Product_L1__c%5C%22%2C%5C%22Case.Logo_Name__c%5C%22%2C%5C%22Case.Date_Time_Assigned_to_User__c%5C%22%2C%5C%22STATUS%5C%22%2C%5C%22CLOSED_DATEONLY%5C%22%5D%2C%5C%22developerName%5C%22%3A%5C%22Report_for_Mytime%5C%22%2C%5C%22division%5C%22%3Anull%2C%5C%22folderId%5C%22%3A%5C%220054w00000CNhVAAA1%5C%22%2C%5C%22groupingsAcross%5C%22%3A%5B%5D%2C%5C%22groupingsDown%5C%22%3A%5B%5D%2C%5C%22hasDetailRows%5C%22%3Atrue%2C%5C%22hasRecordCount%5C%22%3Atrue%2C%5C%22historicalSnapshotDates%5C%22%3A%5B%5D%2C%5C%22id%5C%22%3A%5C%2200OWQ000005mZ5R2AU%5C%22%2C%5C%22name%5C%22%3A%5C%22Report_for_Mytime%5C%22%2C%5C%22paletteColors%5C%22%3A%5B%5C%22%23945bbe%5C%22%2C%5C%22%237942a4%5C%22%2C%5C%22%235b317c%5C%22%2C%5C%22%233f2056%5C%22%2C%5C%22%23220f31%5C%22%2C%5C%22%23c3a1d9%5C%22%2C%5C%22%23ae81cd%5C%22%5D%2C%5C%22presentationOptions%5C%22%3A%7B%5C%22hasStackedSummaries%5C%22%3Atrue%7D%2C%5C%22reportBooleanFilter%5C%22%3Anull%2C%5C%22reportFilters%5C%22%3A%5B%5D%2C%5C%22reportFormat%5C%22%3A%5C%22TABULAR%5C%22%2C%5C%22reportType%5C%22%3A%7B%5C%22label%5C%22%3A%5C%22Cases%5C%22%2C%5C%22type%5C%22%3A%5C%22CaseList%5C%22%7D%2C%5C%22scope%5C%22%3A%5C%22user%5C%22%2C%5C%22showGrandTotal%5C%22%3Atrue%2C%5C%22showSubtotals%5C%22%3Atrue%2C%5C%22sortBy%5C%22%3A%5B%7B%5C%22sortColumn%5C%22%3A%5C%22Case.Date_Time_Assigned_to_User__c%5C%22%2C%5C%22sortOrder%5C%22%3A%5C%22Desc%5C%22%7D%5D%2C%5C%22standardDateFilter%5C%22%3A%7B%5C%22column%5C%22%3A%5C%22CREATED_DATEONLY%5C%22%2C%5C%22durationValue%5C%22%3A%5C%22CUSTOM%5C%22%2C%5C%22endDate%5C%22%3Anull%2C%5C%22startDate%5C%22%3Anull%7D%2C%5C%22standardFilters%5C%22%3A%5B%7B%5C%22name%5C%22%3A%5C%22units%5C%22%2C%5C%22value%5C%22%3A%5C%22d%5C%22%7D%5D%2C%5C%22supportsRoleHierarchy%5C%22%3Afalse%2C%5C%22userOrHierarchyFilterId%5C%22%3Anull%2C%5C%22customSummaryFormula%5C%22%3A%7B%7D%2C%5C%22customDetailFormula%5C%22%3A%7B%7D%2C%5C%22buckets%5C%22%3A%5B%5D%2C%5C%22userOrHierarchyFilterName%5C%22%3Anull%2C%5C%22dataCategoryFilters%5C%22%3A%5B%5D%2C%5C%22aggregateFilters%5C%22%3A%5B%5D%7D%7D%22%2C%22isPreview%22%3Afalse%2C%22createReportInstance%22%3Afalse%2C%22fastCsv%22%3Afalse%2C%22requestOrigin%22%3A%22rpgd%22%2C%22includeChartData%22%3Afalse%2C%22skipReportResult%22%3Afalse%2C%22skipRocs%22%3Afalse%7D%2C%22storable%22%3Atrue%7D%5D%7D&aura.context=%7B%22mode%22%3A%22PROD%22%2C%22fwuid%22%3A%22cmpKNldRZXRSMkdjemxQdjBkbl9uQWtVMjdnTGFERUU2S3FfSVdrcU92bkExNC4xOTIuODM4ODYwOA%22%2C%22app%22%3A%22one%3Aone%22%2C%22loaded%22%3A%7B%22APPLICATION%40markup%3A%2F%2Fone%3Aone%22%3A%224146_iERZh3UXxQMRsITHi4MOkg%22%7D%2C%22dn%22%3A%5B%5D%2C%22globals%22%3A%7B%22appContextId%22%3A%2206m1U000000TFECQA4%22%7D%2C%22uad%22%3Atrue%7D&aura.pageURI=%2Flightning%2Fr%2FReport%2F00OWQ000005mZ5R2AU%2Fview%3FqueryScope%3DuserFolders&aura.token=eyJub25jZSI6Im5nbGhiLVJHbXNSM2g4MVFIa21jdW40eU1JZVlBUU96dnB0ak9BVW9KaWNcdTAwM2QiLCJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiIsImtpZCI6IntcInRcIjpcIjAwRFdRMDAwMDAwMDAwMVwiLFwidlwiOlwiMDJHV1EwMDAwMDAwMDFkXCIsXCJhXCI6XCJjYWltYW5zaWduZXJcIn0iLCJjcml0IjpbImlhdCJdLCJpYXQiOjE3ODIzNzk2NjA4ODAsImV4cCI6MH0%3D..P1KdGKPaZuQG8F12r--I5S3-wbBVAPQwuLPJq1jOHPQ%3D", // (keep your original full body unchanged)
  credentials: "include"
})
.then(res => res.text())

.then(text => {

  // =========================================================
  // ✅ STEP 1: Clean & parse Salesforce response
  // =========================================================
  const clean = text.replace(/^while\\(1\\);/, "");
  const json = JSON.parse(clean);
  const result = json.actions[0].returnValue;

  // Extract raw rows
  const rows = result.factMap["T!T"].rows;

  // Convert to simple array format
  // [Product L1, Logo, Open Date, Status, Close Date]
  const data = rows.map(r =>
    r.dataCells.map(c => c.label)
  );

  console.log("TABLE DATA:", data);


  // =========================================================
  // ✅ STEP 2: Define current week (Monday → Friday)
  // =========================================================
  const today = new Date();
  const day = today.getDay();

  const monday = new Date(today);
  monday.setDate(today.getDate() - (day === 0 ? 6 : day - 1));
  monday.setHours(0,0,0,0);

  const friday = new Date(monday);
  friday.setDate(monday.getDate() + 4);
  friday.setHours(23,59,59,999);


  // =========================================================
  // ✅ STEP 3: Helper functions
  // =========================================================
  function toDateSafe(d) {
    if (!d || d === "-") return null;
    return new Date(d);
  }


  // =========================================================
  // ✅ STEP 4: Count number of ACTIVE cases per weekday
  // =========================================================
  const weekdayCount = {
    Monday: 0,
    Tuesday: 0,
    Wednesday: 0,
    Thursday: 0,
    Friday: 0
  };

  data.forEach(row => {
    const openDate  = toDateSafe(row[2]);
    const status    = (row[3] || "").toLowerCase();
    const closeDate = toDateSafe(row[4]);

    if (
      openDate &&
      openDate <= friday &&
      (status === "open" || !closeDate || closeDate >= monday)
    ) {
      for (let i = 0; i < 5; i++) {
        const d = new Date(monday);
        d.setDate(monday.getDate() + i);

        const isActive =
          openDate <= d &&
          (status === "open" || !closeDate || closeDate >= d);

        if (isActive) {
          const key = ["Monday","Tuesday","Wednesday","Thursday","Friday"][i];
          weekdayCount[key]++;
        }
      }
    }
  });

  console.log("WEEKDAY COUNTS:", weekdayCount);


  // =========================================================
  // ✅ STEP 5: Distribute EXACT 100 per day across cases
  // =========================================================
  const weekdayWeightsPerCase = {
    Monday: [],
    Tuesday: [],
    Wednesday: [],
    Thursday: [],
    Friday: []
  };

  Object.keys(weekdayCount).forEach(day => {
    const count = weekdayCount[day];
    if (count === 0) return;

    const base = Math.floor(100 / count);
    const remainder = 100 % count;

    for (let i = 0; i < count; i++) {
      weekdayWeightsPerCase[day].push(
        i < remainder ? base + 1 : base
      );
    }
  });

  console.log("WEEKDAY DISTRIBUTION:", weekdayWeightsPerCase);


  // =========================================================
  // ✅ STEP 6: Assign weekday % to each CASE (row level)
  // Only for cases active in this week
  // =========================================================
  const dayIndex = {
    Monday: 0,
    Tuesday: 0,
    Wednesday: 0,
    Thursday: 0,
    Friday: 0
  };

  const enrichedRows = data
    .filter(row => {
      const openDate  = toDateSafe(row[2]);
      const status    = (row[3] || "").toLowerCase();
      const closeDate = toDateSafe(row[4]);

      if (!openDate) return false;

      // ✅ keep only rows active this week
      return (
        openDate <= friday &&
        (status === "open" || !closeDate || closeDate >= monday)
      );
    })

    .map(row => {
      const openDate  = toDateSafe(row[2]);
      const status    = (row[3] || "").toLowerCase();
      const closeDate = toDateSafe(row[4]);

      const values = {
        Monday: "",
        Tuesday: "",
        Wednesday: "",
        Thursday: "",
        Friday: ""
      };

      for (let i = 0; i < 5; i++) {
        const d = new Date(monday);
        d.setDate(monday.getDate() + i);
        d.setHours(23,59,59,999);

        const isActive =
          openDate <= d &&
          (status === "open" ? true : (!closeDate || closeDate >= d));

        if (isActive) {
          const dayName = ["Monday","Tuesday","Wednesday","Thursday","Friday"][i];

          const idx = dayIndex[dayName];
          values[dayName] = weekdayWeightsPerCase[dayName][idx];

          dayIndex[dayName]++;
        }
      }

      return [
        "Post-Sales",
        "Reactive/Tape-out support",
        row[1], // Logo
        row[0], // Product
        values.Monday,
        values.Tuesday,
        values.Wednesday,
        values.Thursday,
        values.Friday
      ];
    });


  // =========================================================
  // ✅ STEP 7: Merge rows with same (Logo + Product L1)
  // =========================================================
  const mergedMap = new Map();

  enrichedRows.forEach(row => {
    const key = row[2] + "||" + row[3];

    if (!mergedMap.has(key)) {
      mergedMap.set(key, [...row]);
    } else {
      const existing = mergedMap.get(key);

      for (let i = 4; i <= 8; i++) {
        existing[i] =
          Number(existing[i] || 0) + Number(row[i] || 0);
      }
    }
  });

  const finalRows = Array.from(mergedMap.values());


  // =========================================================
  // ✅ STEP 8: Normalize again → each weekday = EXACT 100
  // =========================================================
  const dayCols = {
    Monday: 4,
    Tuesday: 5,
    Wednesday: 6,
    Thursday: 7,
    Friday: 8
  };

  Object.entries(dayCols).forEach(([day, colIndex]) => {
    const total = finalRows.reduce((sum, row) =>
      sum + Number(row[colIndex] || 0), 0);

    if (total === 0) return;

    let runningTotal = 0;

    finalRows.forEach(row => {
      const scaled = Math.floor((Number(row[colIndex] || 0) / total) * 100);
      row[colIndex] = scaled;
      runningTotal += scaled;
    });

    // Fix rounding difference
    let diff = 100 - runningTotal;
    let i = 0;

    while (diff > 0) {
      finalRows[i % finalRows.length][colIndex]++;
      diff--;
      i++;
    }
  });

  console.log("FINAL NORMALIZED ROWS:", finalRows);


  // =========================================================
  // ✅ STEP 9: Generate CSV (Excel file)
  // =========================================================
  const headers = [
    "Category","activity","Logo","Product L1",
    "Monday","Tuesday","Wednesday","Thursday","Friday"
  ];

  const csvContent =
    [headers, ...finalRows]
      .map(r => r.join(","))
      .join("\n");

  const blob = new Blob([csvContent], { type: "text/csv" });
  const url = URL.createObjectURL(blob);

  const a = document.createElement("a");
  a.href = url;
  a.download = "report_output.csv";
  a.click();

  URL.revokeObjectURL(url);
});
