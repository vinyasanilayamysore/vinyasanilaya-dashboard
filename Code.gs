var ID_GUESTS_LIST = "1Puw0OezY18OWFt8wtwzv5BFxcJw314Hfov5GZMUXCbk";

var FIREBASE_PROJECT_ID = "vinyasanilaya-website"; // Replace with your Project ID

/**
 * Whitelist of allowed email addresses.
 * Add or remove emails here to control dashboard access.
 */
const ALLOWED_EMAILS = [
  "vinyasanilayamysore@gmail.com",
  "yashaswini.hj.vinay@gmail.com"
];

function doGet(e) {
  const userEmail = Session.getActiveUser().getEmail();
  
  // 1. Check if identity is hidden (Google Privacy Restriction)
  if (!userEmail) {
    return HtmlService.createHtmlOutput(
      "<div style='font-family: sans-serif; text-align: center; margin-top: 100px; padding: 20px;'>" +
        "<h2 style='color: #d93025;'>🔒 Identity Required</h2>" +
        "<p>To verify your access, you must grant this app permission to see your email address.</p>" +
        "<div style='margin: 20px 0;'>" +
          "<button onclick='authorize()' style='padding: 12px 24px; background-color: #1a73e8; color: white; border: none; border-radius: 4px; cursor: pointer; font-weight: bold;'>Grant Access & Sign In</button>" +
        "</div>" +
        "<p style='font-size: 12px; color: #666;'>Tip: If you are logged into multiple Google accounts, please use an <b>Incognito window</b>.</p>" +
        "<script>" +
          "function authorize() {" +
            "google.script.run.withSuccessHandler(function() {" +
              "window.location.reload();" +
            "}).loginTrigger();" +
          "}" +
        "</script>" +
      "</div>"
    ).setTitle('Identity Required');
  }

  // 2. Authorization Whitelist Check
  if (!ALLOWED_EMAILS.includes(userEmail)) {
    return HtmlService.createHtmlOutput(
      "<div style='font-family: sans-serif; text-align: center; margin-top: 50px;'>" +
      "<h2>🚫 Access Denied</h2>" +
      "<p>You are not authorized to access this dashboard.</p>" +
      "<p style='color: #666;'>Signed in as: <b>" + userEmail + "</b></p>" +
      "</div>"
    ).setTitle('Access Denied');
  }

  // Serve the dashboard if authorized
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Vinyasa Nilaya | Guest Dashboard')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setFaviconUrl('https://ik.imagekit.io/h87o83ayxm/Icons/Icon03_png');
}

/**
 * Dummy function used to trigger the Google OAuth Consent popup 
 * if the user has not yet authorized the web app.
 */
function loginTrigger() {
  console.log("Authorization trigger called by: " + Session.getActiveUser().getEmail());
  return true;
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Core Data Lifecycle Engine.
 * Fetches all valid year sheets and guarantees the active year's tracking sheet exists.
 */
function getAvailableYearsAndInitialize() {
  try {
    const ss = SpreadsheetApp.openById(ID_GUESTS_LIST);
    const currentYear = new Date().getFullYear().toString(); // e.g., "2026"

    // 1. DYNAMIC SYSTEM INITIALIZATION: Check/Create New Year Sheet on Jan 1st
    let targetSheet = ss.getSheetByName(currentYear);
    if (!targetSheet) {
      console.log(`>>> [SYSTEM] Jan 1st Lifecyle Trigger: Creating new tracking tab for [${currentYear}]`);

      // Attempt to find your most recent historical sheet to use as a structural template
      const allSheets = ss.getSheets();
      let templateSheet = allSheets[0]; // Fallback to first sheet

      // Try to find a sheet name that looks like a 4-digit number to copy layout from
      for (let s of allSheets) {
        if (/^\d{4}$/.test(s.getName())) {
          templateSheet = s;
          break;
        }
      }

      // Duplicate the template sheet structure to preserve column names, widths, and formatting models
      if (templateSheet) {
        let newSheet = templateSheet.copyTo(ss);
        newSheet.setName(currentYear);

        // Clean out old historical row cell values while maintaining structural formatting headers
        const lastRow = newSheet.getLastRow();
        if (lastRow > 1) {
          // Assuming row 1 has your "Month", "Name", "Amount" header values
          newSheet.getRange(2, 1, lastRow - 1, newSheet.getLastColumn()).clearContent();
        }
        // Move to front position for visibility
        ss.setActiveSheet(newSheet);
        ss.moveActiveSheet(1);
      }
    }

    // 2. RETRIEVAL SYSTEM: Gather all 4-digit numeric sheet names for the UI dropdown
    const sheets = ss.getSheets();
    let availableYears = [];

    sheets.forEach(sheet => {
      const name = sheet.getName().trim();
      if (/^\d{4}$/.test(name)) { // Matches exact 4-digit years like "2025", "2026"
        availableYears.push(name);
      }
    });

    // Sort years in descending order so the newest year is listed first (e.g., 2026, 2025)
    availableYears.sort((a, b) => Number(b) - Number(a));

    return {
      years: availableYears,
      activeYear: currentYear
    };

  } catch (err) {
    console.error("Initialization / Year discovery failure: " + err.message);
    return { years: [new Date().getFullYear().toString()], activeYear: new Date().getFullYear().toString() };
  }
}

/*
 * Fetches guest data and summaries based on the specific column headers:
 * Year, Month, Name, NoOfGuests, Amount, Check-in Date, Days, AirBnb\Personal, Floor, Mobile, Customer Ratings, Comments
 */
function getDashboardData(filterYear, filterMonth) {
  try {
    const ss = SpreadsheetApp.openById(ID_GUESTS_LIST);
    const sheets = ss.getSheets();

    // =========================================================================
    // REQUIRED SPEED CHANGE 1: PRE-FETCH ALL VERIFIED NUMBERS IN ONE ROUND-TRIP
    // =========================================================================
    // Now fetching from Firestore 'guests' collection instead of Spreadsheet
    let verifiedMobilesSet = new Set();
    try {
      const url = `https://firestore.googleapis.com/v1/projects/${FIREBASE_PROJECT_ID}/databases/(default)/documents/guests`;
      const response = UrlFetchApp.fetch(url, {
        headers: { "Authorization": "Bearer " + ScriptApp.getOAuthToken() }
      });
      const firestoreData = JSON.parse(response.getContentText());

      if (firestoreData.documents) {
        firestoreData.documents.forEach(doc => {
          const fields = doc.fields;
          // Accessing nested guestDetails.phone from Firestore REST structure
          let phone = "";
          if (fields.guestDetails && fields.guestDetails.mapValue.fields.phone) {
            phone = fields.guestDetails.mapValue.fields.phone.stringValue;
          } else if (fields.phone) { phone = fields.phone.stringValue; }

          if (phone) {
            let clean = phone.replace(/\D/g, "");
            if (clean.length > 10) clean = clean.slice(-10);
            if (clean.length === 10) verifiedMobilesSet.add(clean);
          }
        });
      }
    } catch (qidErr) {
      console.error(">>> [SPEED ENGINE WARNING] Cache pre-fetch failed: " + qidErr.message);
    }
    // =========================================================================

    // -----------------------------------------------------------------
    // PHASE 1: COMPUTE CONSOLIDATED LIFETIME REVENUE ACROSS ALL YEARS
    // -----------------------------------------------------------------
    let lifetimeTotalRevenue = 0;
    let lifetimeTotalCheckIns = 0;

    sheets.forEach(sheet => {
      const sheetName = sheet.getName().trim();

      if (/^\d{4}$/.test(sheetName)) {
        const sheetData = sheet.getDataRange().getValues();
        const headerRowIdx = sheetData.findIndex(row => row.includes("Name") || row.includes("Amount"));

        if (headerRowIdx !== -1) {
          const sheetHeaders = sheetData[headerRowIdx];
          const amountIdx = sheetHeaders.indexOf("Amount");
          const nameIdx = sheetHeaders.indexOf("Name");

          if (amountIdx !== -1) {
            const sheetRows = sheetData.slice(headerRowIdx + 1);

            sheetRows.forEach(row => {
              const nameVal = row[nameIdx] ? row[nameIdx].toString().trim() : "";
              if (!nameVal || nameVal === "Total" || nameVal === "No Guests" || nameVal === "") return;

              let amtStr = (row[amountIdx] || "0").toString().replace(/[₹,]/g, "");
              let amtNum = Number(amtStr) || 0;
              lifetimeTotalRevenue += amtNum;
              lifetimeTotalCheckIns++;
            });
          }
        }
      }
    });

    // -----------------------------------------------------------------
    // PHASE 2: PROCESSING SELECTED FOCUS TARGET SHEET DATA
    // -----------------------------------------------------------------
    const targetTab = filterYear;
    let sheet = ss.getSheetByName(targetTab) || ss.getSheets()[0];

    const data = sheet.getDataRange().getValues();
    const headerRowIndex = data.findIndex(row => row.includes("Name") || row.includes("Month"));

    if (headerRowIndex === -1) {
      return {
        guests: [],
        summary: {
          totalRevenue: "₹0",
          count: 0,
          period: "No Headers Found",
          lifetimeRevenue: lifetimeTotalRevenue.toLocaleString('en-IN')
        }
      };
    }

    const headers = data[headerRowIndex];
    const rows = data.slice(headerRowIndex + 1);

    let totalRevenue = 0;
    let guestCount = 0;

    const filteredData = rows.filter(row => {
      const name = row[headers.indexOf("Name")];
      const month = row[headers.indexOf("Month")];
      if (!name || name === "Total" || name === "No Guests" || name === "") return false;

      const rowMonth = month ? month.toString().trim() : "";
      return !filterMonth || rowMonth === filterMonth;
    }).map(row => {
      let obj = {};
      headers.forEach((header, i) => {
        let key = header.toString();
        if (key === "AirBnb\\Personal" || key === "Source") {
          key = "Source";
        }
        else if (key === "Floor") {
          key = "Floor";
        }
        else {
          key = key.replace(/\\|\s|-/g, "_");
        }

        let value = row[i];
        if (value instanceof Date) {
          value = Utilities.formatDate(value, Session.getScriptTimeZone(), "yyyy-MM-dd");
        }
        obj[key] = value;
      });

      // 1. Revenue & Stats Calculation
      let amtStr = (obj.Amount || "0").toString().replace(/[₹,]/g, "");
      let amtNum = Number(amtStr) || 0;
      totalRevenue += amtNum;
      guestCount++;

      // 2. Mobile Sanitization for WhatsApp & Verification
      let mobileRaw = (obj.Mobile || "").toString().replace(/\D/g, "");
      if (mobileRaw.length > 10) {
        mobileRaw = mobileRaw.slice(-10);
      }
      obj['WhatsApp_Num'] = mobileRaw;

      // =========================================================================
      // REQUIRED SPEED CHANGE 2: USE INSTANT MEMORY SET INSTEAD OF THE SLOW LOOP FUNCTION
      // =========================================================================
      obj['isVerified'] = (mobileRaw.length === 10 && verifiedMobilesSet.has(mobileRaw));
      // =========================================================================

      return obj;
    });

    return {
      guests: filteredData,
      summary: {
        totalRevenue: totalRevenue.toLocaleString('en-IN', { maximumFractionDigits: 0 }).replace("INR", "").trim(),
        count: guestCount,
        period: filterMonth ? `${filterMonth} ${targetTab}` : targetTab,
        lifetimeRevenue: lifetimeTotalRevenue.toLocaleString('en-IN', { maximumFractionDigits: 0 }),
        lifetimeCount: lifetimeTotalCheckIns
      }
    };
  } catch (err) {
    console.error("Dashboard Sync Error: " + err.message);
    return { guests: [], summary: { totalRevenue: "Error", count: 0, period: "Sheet Error", lifetimeRevenue: "Error" } };
  }
}

/**
 * Verifies if a specific WhatsApp/Mobile number exists in the Firestore 'guests' collection.
 */
function findGuestQIDVerified(whatsappNo) {
  if (!whatsappNo) return false;
  try {
    let targetClean = whatsappNo.toString().replace(/\D/g, "");
    if (targetClean.length > 10) targetClean = targetClean.slice(-10);
    if (targetClean.length !== 10) return false;

    const url = `https://firestore.googleapis.com/v1/projects/${FIREBASE_PROJECT_ID}/databases/(default)/documents/guests`;
    const response = UrlFetchApp.fetch(url, {
      headers: { "Authorization": "Bearer " + ScriptApp.getOAuthToken() }
    });
    const firestoreData = JSON.parse(response.getContentText());

    if (firestoreData.documents) {
      for (let doc of firestoreData.documents) {
        let phone = "";
        if (doc.fields.guestDetails && doc.fields.guestDetails.mapValue.fields.phone) {
          phone = doc.fields.guestDetails.mapValue.fields.phone.stringValue;
        } else if (doc.fields.phone) { phone = doc.fields.phone.stringValue; }

        let currentClean = phone.replace(/\D/g, "");
        if (currentClean.length > 10) currentClean = currentClean.slice(-10);
        if (currentClean === targetClean) {
          console.log(`>>> 🎯 [FIRESTORE MATCH FOUND] ${targetClean}`);
          return true;
        }
      }
    }
    return false; // No matches found across the iteration feed

  } catch (err) {
    console.error(">>> ❌ [findGuestQIDVerified Exception] Check failed: " + err.message);
    return false;
  }
}

/**** Sync with Airbnb bookings feature - Fresh Confirmations with Daily Timeline Sorting *****/
function syncAirbnbEmails(selectedYear) {
  try {
    const ss = SpreadsheetApp.openById(ID_GUESTS_LIST);

    // 1. DYNAMIC YEAR FALLBACK: Use passed year, or calculate current year if empty
    const targetYear = selectedYear || new Date().getFullYear().toString();

    let sheet = ss.getSheetByName(targetYear) || ss.getSheets()[0];

    const data = sheet.getDataRange().getValues();
    const headers = data.find(row => row.includes("Name"));
    if (!headers) return "Error: Could not find header row in sheet.";

    // Setup the Gmail Sync tracking label
    const labelName = "Vinyasa-Synced";
    let syncLabel = GmailApp.getUserLabelByName(labelName) || GmailApp.createLabel(labelName);

    // Dynamic Column Matrix Mapping
    const nameColIdx = headers.indexOf("Name");
    const checkInColIdx = headers.indexOf("Check-in Date");
    const amountColIdx = headers.indexOf("Amount");
    const commentsColIdx = headers.indexOf("Comments");
    const monthColIdx = headers.indexOf("Month");
    const daysColIdx = headers.indexOf("Days");
    const ratingsColIdx = headers.indexOf("Customer Ratings") !== -1 ? headers.indexOf("Customer Ratings") : headers.indexOf("Ratings");

    // 🔑 INDEX MAPPING FOR RESERVATION ID
    const resIdColIdx = headers.indexOf("Reservation ID") !== -1 ? headers.indexOf("Reservation ID") : headers.indexOf("ReservationID");

    // --- FULL MONTH NAME DICTIONARY ---
    const monthMap = {
      "jan": "January", "feb": "February", "mar": "March", "apr": "April",
      "may": "May", "jun": "June", "jul": "July", "aug": "August",
      "sep": "September", "oct": "October", "nov": "November", "dec": "December"
    };

    // 🔑 UPDATED QUERY: Included "wants to change" and "Reservation updated"
    const query = `from:automated@airbnb.com ("Reservation confirmed" OR "review" OR "Canceled:" OR "wants to change" OR "Reservation updated") -label:${labelName}`;
    const threads = GmailApp.search(query, 0, 15);

    let newBookingsCount = 0;
    let reviewsCount = 0;
    let cancellationCount = 0;
    let updateCount = 0; // Fixed: Initialized update counter

    threads.forEach(thread => {
      const messages = thread.getMessages();
      let processedThread = false;

      messages.forEach(message => {
        const subject = message.getSubject();
        const body = message.getPlainBody();
        const combinedTextToAnalyze = (subject + " " + body);

        // =========================================================================
        // SCENARIO D1: PENDING CHANGE REQUEST ("wants to change")
        // =========================================================================
        if (combinedTextToAnalyze.includes("wants to change their reservation")) {
          console.log(`>>> [ALTERATION REQUEST] Pending request received. Marking email synced until host approves.`);
          processedThread = true;
          return;
        }

        // =========================================================================
        // SCENARIO D2: CONFIRMED RESERVATION UPDATE ("Reservation updated")
        // =========================================================================
        if (combinedTextToAnalyze.includes("reservation updated") || combinedTextToAnalyze.includes("reservation with")) {
          console.log(`>>> [ALTERATION APPROVED] Processing reservation update email...`);

          // 1. Extract Guest Name
          const nameMatch = subject.match(/reservation\s+with\s+(.*?)\s+has\s+been\s+updated/i) ||
            body.match(/reservation\s+with\s+(.*?)\s+has\s+been\s+updated/i);
          const guestName = nameMatch ? nameMatch[1].trim() : "";

          if (!guestName) {
            console.warn(">>> [ALTERATION ERROR] Could not parse guest name from update email.");
            return;
          }

          // 2. Extract New Dates
          const dateRangeMatch = body.match(/([A-Z][a-z]{2}\s+\d+)(?:,\s*\d{4})?\s*-\s*([A-Z][a-z]{2}\s+\d+),\s*(\d{4})/i);

          let newCheckInStr = "";
          let newNights = 0;
          let newCheckInDateObj = null;

          if (dateRangeMatch) {
            const startStr = dateRangeMatch[1].trim(); 
            const endStr = dateRangeMatch[2].trim();   
            const yearStr = dateRangeMatch[3].trim();  

            newCheckInStr = `${startStr}, ${yearStr}`;
            newCheckInDateObj = new Date(newCheckInStr);

            const endDateObj = new Date(`${endStr}, ${yearStr}`);

            if (!isNaN(newCheckInDateObj.getTime()) && !isNaN(endDateObj.getTime())) {
              const diffTime = Math.abs(endDateObj - newCheckInDateObj);
              newNights = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
            }
          }

          // 3. Extract New Amount
          const amountMatch = body.match(/You earn[\s\S]*?₹?\s*([\d,]+\.?\d*)/i) || body.match(/Total payout[\s\S]*?₹?\s*([\d,]+\.?\d*)/i);
          let updatedAmount = null;
          if (amountMatch) {
            let rawAmt = amountMatch[1].replace(/,/g, "");
            updatedAmount = Math.round(parseFloat(rawAmt)).toString();
          }

          // 4. Locate Existing Booking in Sheet by Guest Name
          let currentSheetData = sheet.getDataRange().getValues();
          let matchedRowIdx = -1;

          for (let row = 1; row < currentSheetData.length; row++) {
            let sheetGuestName = currentSheetData[row][nameColIdx] ? currentSheetData[row][nameColIdx].toString().trim().toLowerCase() : "";

            if (sheetGuestName.includes(guestName.toLowerCase()) || guestName.toLowerCase().includes(sheetGuestName)) {
              matchedRowIdx = row + 1; // 1-based index
              break;
            }
          }

          if (matchedRowIdx !== -1) {
            console.log(`>>> [ALTERATION SUCCESS] Found existing row ${matchedRowIdx} for ${guestName}. Updating itinerary...`);

            // Update Check-in Date
            if (newCheckInStr && checkInColIdx !== -1) {
              sheet.getRange(matchedRowIdx, checkInColIdx + 1).setValue(newCheckInStr);

              // Update Month Column
              let rawMonthAbbr = newCheckInStr.split(" ")[0].toLowerCase().replace(/[^a-z]/g, "");
              let fullMonthName = monthMap[rawMonthAbbr] || newCheckInStr.split(" ")[0];
              if (monthColIdx !== -1) sheet.getRange(matchedRowIdx, monthColIdx + 1).setValue(fullMonthName);
            }

            // Update Nights/Days
            if (newNights > 0 && daysColIdx !== -1) {
              sheet.getRange(matchedRowIdx, daysColIdx + 1).setValue(newNights);
            }

            // Update Payout Amount
            if (updatedAmount !== null && amountColIdx !== -1) {
              sheet.getRange(matchedRowIdx, amountColIdx + 1).setValue(updatedAmount);
            }

            // Append comment
            if (commentsColIdx !== -1) {
              let existingComment = currentSheetData[matchedRowIdx - 1][commentsColIdx] || "";
              sheet.getRange(matchedRowIdx, commentsColIdx + 1).setValue(`${existingComment} | Updated via Sync on ${new Date().toLocaleDateString()}`.trim());
            }

            // 5. Update Calendar Entry (Fixed Parameter Signature: Passed reservationId first)
            if (newCheckInDateObj && newNights > 0) {
              let guestCount = currentSheetData[matchedRowIdx - 1][headers.indexOf("Guests")] || 1;
              let floorVal = currentSheetData[matchedRowIdx - 1][headers.indexOf("Floor")] || "Ground";
              let existingResId = (resIdColIdx !== -1) ? currentSheetData[matchedRowIdx - 1][resIdColIdx] : `HM-${Date.now()}`;

              syncBookingToVinyasaCalendar(
                existingResId,
                guestName,
                newCheckInDateObj,
                guestCount,
                "AirBnb",
                floorVal,
                newNights,
                "Updated Reservation Itinerary"
              );
            }

            updateCount++;
            processedThread = true;
          } else {
            console.warn(`>>> [ALTERATION WARNING] Received update email for "${guestName}", but no existing row was found in sheet.`);
          }

          return;
        }

        // =========================================================================
        // SCENARIO A: GUEST LEFT A REVIEW
        // =========================================================================
        if (combinedTextToAnalyze.toLowerCase().includes("left a") && combinedTextToAnalyze.toLowerCase().includes("review")) {
          console.log(`>>> [REVIEW TRACE] Processing Review Email`);

          const reviewMatch = combinedTextToAnalyze.match(/(.+?)\s+left\s+a\s+(\d+)-star\s+review/i);

          if (reviewMatch) {
            let rawName = reviewMatch[1].replace(/canceled:|cancelled:|reservation confirmed\s*-\s*/i, "").trim();
            const reviewerShortName = rawName.toLowerCase();

            const numericRating = parseInt(reviewMatch[2].trim());
            const starRating = "⭐".repeat(numericRating);

            let currentSheetData = sheet.getDataRange().getValues();
            const platformColIdx = headers.indexOf("AirBnb\\Personal");

            for (let row = 1; row < currentSheetData.length; row++) {
              let rawSheetName = currentSheetData[row][nameColIdx] ? currentSheetData[row][nameColIdx].toString() : "";
              let existingFullName = rawSheetName.trim().toLowerCase();

              let rawPlatform = platformColIdx !== -1 && currentSheetData[row][platformColIdx] ? currentSheetData[row][platformColIdx].toString() : "";
              let platformType = rawPlatform.trim().toLowerCase();

              let nameMatches = existingFullName.includes(reviewerShortName);

              if (nameMatches && platformType === "airbnb" && ratingsColIdx !== -1) {
                sheet.getRange(row + 1, ratingsColIdx + 1).setValue(starRating);
                reviewsCount++;
                processedThread = true;
                break;
              }
            }
          }
          return;
        }

        // =========================================================================
        // SCENARIO B: CANCELLATION NOTICE RECEIVED
        // =========================================================================
        if (combinedTextToAnalyze.toLowerCase().includes("canceled:") || combinedTextToAnalyze.toLowerCase().includes("cancelled:")) {
          const cancelMatch = combinedTextToAnalyze.match(/Reservation\s+([A-Z0-9]{10})/i) || combinedTextToAnalyze.match(/code\s+([A-Z0-9]{10})/i);
          const dateMatch = subject.match(/for\s+([A-Z][a-z]{2}\s+\d+)/i);

          let targetYearNum = parseInt(targetYear);
          let targetMonthNum = -1;
          let targetDayNum = -1;

          if (dateMatch) {
            let emailDateObj = new Date(dateMatch[1].trim() + `, ${targetYear}`);
            if (!isNaN(emailDateObj.getTime())) {
              targetMonthNum = emailDateObj.getMonth() + 1;
              targetDayNum = emailDateObj.getDate();
            }
          }

          const targetConfirmationCode = cancelMatch ? cancelMatch[1].trim() : "NOT_FOUND";
          let currentSheetData = sheet.getDataRange().getValues();
          let matchRowIdx = -1;

          for (let row = 1; row < currentSheetData.length; row++) {
            let commentsValue = commentsColIdx !== -1 && currentSheetData[row][commentsColIdx] ? currentSheetData[row][commentsColIdx].toString() : "";
            let resIdValue = resIdColIdx !== -1 && currentSheetData[row][resIdColIdx] ? currentSheetData[row][resIdColIdx].toString() : "";
            let rowCheckInValue = checkInColIdx !== -1 ? currentSheetData[row][checkInColIdx] : null;

            // Strategy 1: Unique Confirmation Code Matching
            if (targetConfirmationCode !== "NOT_FOUND" && (commentsValue.includes(targetConfirmationCode) || resIdValue.includes(targetConfirmationCode))) {
              matchRowIdx = row + 1;
              break;
            }

            // Strategy 2: Date Matching Fallback
            if (targetMonthNum !== -1 && rowCheckInValue) {
              let rowMonth, rowDay, rowYear;

              if (rowCheckInValue instanceof Date) {
                rowMonth = rowCheckInValue.getMonth() + 1;
                rowDay = rowCheckInValue.getDate();
                rowYear = rowCheckInValue.getFullYear();
              } else {
                let dateParts = rowCheckInValue.toString().trim().split("/");
                if (dateParts.length === 3) {
                  rowMonth = parseInt(dateParts[0]);
                  rowDay = parseInt(dateParts[1]);
                  rowYear = parseInt(dateParts[2]);
                }
              }

              if (rowMonth === targetMonthNum && rowDay === targetDayNum && rowYear === targetYearNum) {
                matchRowIdx = row + 1;
                break;
              }
            }
          }

          if (matchRowIdx !== -1) {
            if (commentsColIdx !== -1) {
              sheet.getRange(matchRowIdx, commentsColIdx + 1).setValue(`CANCELLED CODE: ${targetConfirmationCode}. Flagged via sync.`);
            }

            if (amountColIdx !== -1) {
              sheet.getRange(matchRowIdx, amountColIdx + 1).setValue("");
            }

            sheet.getRange(matchRowIdx, 1, 1, headers.length).setFontColor("#e53e3e");
            cancellationCount++;
            processedThread = true;
          }
          return;
        }

        // =========================================================================
        // SCENARIO C: RESERVATION CONFIRMED
        // =========================================================================
        const subjectMatch = subject.match(/Reservation confirmed\s*-\s*(.*?)\s+arrives\s+(.*)/i);

        let guestName = "";
        let checkInStr = "";

        if (subjectMatch) {
          guestName = subjectMatch[1].trim();
          checkInStr = subjectMatch[2].trim() + `, ${targetYear}`;
        } else {
          return;
        }

        if (!guestName) return;

        const nightsMatch = body.match(/(\d+)\s*nights\s*room\s*fee/i) || body.match(/(\d+)\s*night/i);
        const nights = nightsMatch ? Number(nightsMatch[1]) : 1;

        const amountMatch = body.match(/You earn[\s\S]*?₹?\s*([\d,]+\.?\d*)/i);
        let finalAmount = "0";
        if (amountMatch) {
          let rawAmt = amountMatch[1].replace(/,/g, "");
          finalAmount = Math.round(parseFloat(rawAmt)).toString();
        }

        const guestsMatch = body.match(/(\d+)\s*adult/i) || body.match(/(\d+)\s*guest(?!\s+will)/i);
        const totalGuests = guestsMatch ? Number(guestsMatch[1]) : 1;

        let rawMonthAbbreviation = checkInStr.split(" ")[0].toLowerCase().replace(/[^a-z]/g, "");
        let fullMonthName = monthMap[rawMonthAbbreviation] || checkInStr.split(" ")[0];

        const incomingCheckInDate = new Date(checkInStr);
        const incomingCheckInTime = !isNaN(incomingCheckInDate.getTime()) ? incomingCheckInDate.getTime() : 0;

        // 🔑 EXTRACT OR GENERATE RESERVATION ID
        const codeMatch = body.match(/Reservation\s*code\s*([A-Z0-9]{10})/i) || body.match(/Confirmation\s*code\s*([A-Z0-9]{10})/i);
        const confirmationCode = codeMatch ? codeMatch[1].trim() : "NOT_FOUND";

        const generatedResId = (confirmationCode !== "NOT_FOUND") ? confirmationCode : `HM-${Date.now()}`;
        const savedComment = confirmationCode !== "NOT_FOUND" ? `Code: ${confirmationCode}. Automated Gmail Sync Engine.` : "Automated Gmail Sync Engine.";

        let newRowData = new Array(headers.length).fill("");

        // POPULATE ROW ARRAY
        if (resIdColIdx !== -1) newRowData[resIdColIdx] = generatedResId;
        if (headers.indexOf("Month") !== -1) newRowData[headers.indexOf("Month")] = fullMonthName;
        if (nameColIdx !== -1) newRowData[nameColIdx] = guestName;
        if (headers.indexOf("Guests") !== -1) newRowData[headers.indexOf("Guests")] = totalGuests;
        if (amountColIdx !== -1) newRowData[amountColIdx] = finalAmount;
        if (checkInColIdx !== -1) newRowData[checkInColIdx] = checkInStr;
        if (daysColIdx !== -1) newRowData[daysColIdx] = nights;
        if (headers.indexOf("AirBnb\\Personal") !== -1) newRowData[headers.indexOf("AirBnb\\Personal")] = "AirBnb";
        if (headers.indexOf("Floor") !== -1) newRowData[headers.indexOf("Floor")] = "Ground";
        if (commentsColIdx !== -1) newRowData[commentsColIdx] = savedComment;

        let currentSheetData = sheet.getDataRange().getValues();
        let lastRowWithContent = sheet.getLastRow();

        let insertionRowIndex = lastRowWithContent;
        let foundInsertionSpot = false;

        for (let i = currentSheetData.length - 1; i >= 1; i--) {
          let rowCheckInVal = checkInColIdx !== -1 && currentSheetData[i][checkInColIdx] ? currentSheetData[i][checkInColIdx].toString().trim() : "";

          if (rowCheckInVal) {
            let rowDate = new Date(rowCheckInVal);
            let rowDateTime = rowDate.getTime();

            if (!isNaN(rowDateTime)) {
              if (rowDateTime <= incomingCheckInTime) {
                insertionRowIndex = i + 1;
                foundInsertionSpot = true;
                break;
              }
            }
          }
        }

        if (!foundInsertionSpot && lastRowWithContent > 1) {
          insertionRowIndex = 1;
        }

        console.log(`>>> [TIMELINE INSERT] Placing ${guestName} (${checkInStr}) directly after row: [${insertionRowIndex}]`);

        sheet.insertRowsAfter(insertionRowIndex, 1);

        // Safe row formatting template resolution
        let templateRow = (insertionRowIndex === 1) ? (sheet.getLastRow() > 2 ? 3 : 2) : insertionRowIndex;
        let templateRange = sheet.getRange(templateRow, 1, 1, headers.length);
        let targetRange = sheet.getRange(insertionRowIndex + 1, 1, 1, headers.length);

        templateRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
        targetRange.setValues([newRowData]);

        // 🔑 CALENDAR SYNC (Fixed Parameter Signature: Passed reservationId first)
        syncBookingToVinyasaCalendar(
          generatedResId,
          guestName,
          incomingCheckInDate,
          totalGuests,
          "AirBnb",
          "Ground",
          nights,
          savedComment
        );

        newBookingsCount++;
        processedThread = true;

      });

      if (processedThread) {
        thread.addLabel(syncLabel);
        thread.markRead();
      }
    });

    return `Sync Successfully Completed!\n\n` +
      `📥 New Bookings:\u2003\u2003\u2003\u2003${newBookingsCount} Added\n` +
      `🔄 Updates Processed:\u2003\u2003${updateCount} Updated\n` +
      `⭐ Ratings/Reviews:\u2003\u2003${reviewsCount} Updated\n` +
      `❌ Cancellations:\u2003\u2003\u2003${cancellationCount} Processed`;

  } catch (err) {
    console.error("Parser tracking fail: " + err.message);
    throw new Error("Sync processing aborted: " + err.message);
  }
}

/*****************************Modal to ADD\EDIT Guest Details ***********************************/
function writeGuestDataRow(mode, payload) {
  // --- FIRST-LINE GATEKEEPER VALIDATION ---
  if (!payload.checkIn || payload.checkIn.toString().trim() === "") {
    throw new Error("Transaction Denied: Check-in Date is a mandatory field and cannot be left blank.");
  }

  const ss = SpreadsheetApp.openById(ID_GUESTS_LIST);
  let sheet = ss.getSheetByName(payload.year.toString().trim());

  if (!sheet) {
    sheet = ss.insertSheet(payload.year.toString().trim());
    sheet.appendRow([
      "Month",
      "Name",
      "Guests",
      "Amount",
      "Check-in Date",
      "Days",
      "AirBnb\\Personal",
      "Floor",
      "Mobile",
      "Customer Ratings",
      "Advance",
      "Balance",
      "Base Price",
      "Extra Pax",
      "ReservationID",
      "Comments"
    ]);
    SpreadsheetApp.flush();
  }

  let dataRange = sheet.getDataRange().getValues();

  // Normalize headers (removes spaces & lowercases for robust matching)
  const headers = dataRange[0].map(h => h.toString().trim());
  const normalizedHeaders = headers.map(h => h.toLowerCase().replace(/\s+/g, ''));

  const mapping = {
    ReservationID: normalizedHeaders.indexOf("reservationid"),
    Name: normalizedHeaders.indexOf("name"),
    Amount: normalizedHeaders.indexOf("amount"),
    Source: normalizedHeaders.indexOf("airbnb\\personal"),
    Floor: normalizedHeaders.indexOf("floor"),
    Mobile: normalizedHeaders.indexOf("mobile"),
    CheckIn: normalizedHeaders.indexOf("check-indate"),
    Days: normalizedHeaders.indexOf("days"),
    Guests: normalizedHeaders.indexOf("guests"),
    Ratings: normalizedHeaders.indexOf("customerratings"),
    Comments: normalizedHeaders.indexOf("comments"),
    Advance: normalizedHeaders.indexOf("advance"),
    Balance: normalizedHeaders.indexOf("balance"),
    BasePrice: normalizedHeaders.indexOf("baseprice"),
    ExtraPax: normalizedHeaders.indexOf("extrapax")
  };

  const monthColIdx = normalizedHeaders.indexOf("month");

  if (mapping.ReservationID === -1) {
    throw new Error("Configuration Error: 'ReservationID' column header could not be found in the sheet.");
  }

  // --- 1. NORMALIZE MOBILE NUMBER FORMAT ---
  let cleanMobile = payload.mobile ? payload.mobile.toString().trim() : "";
  if (cleanMobile.startsWith("+91")) {
    cleanMobile = cleanMobile.substring(3);
  } else if (cleanMobile.startsWith("91") && cleanMobile.length > 10) {
    cleanMobile = cleanMobile.substring(2);
  }

  // --- 2. SECURE TIMESTAMP PARSING & UNIFIED STRINGS FORMATION ---
  let generatedMonthLabel = "January";
  let standardizedCheckInString = payload.checkIn;

  const monthsArray = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
  const shortMonthsArray = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];

  const dParsed = parseDateSecurely(payload.checkIn);

  if (!dParsed) {
    throw new Error("Transaction Denied: Invalid Date format provided for Check-in Date.");
  }

  generatedMonthLabel = monthsArray[dParsed.getMonth()];
  standardizedCheckInString = `${shortMonthsArray[dParsed.getMonth()]} ${dParsed.getDate()}, ${dParsed.getFullYear()}`;

  // =========================================================================
  // --- ✨ ANTI-DUPLICATE GUARDRAIL (TEXT-NORMALIZED COMPONENT MATCHING) ---
  // =========================================================================
  if (mode !== "EDIT" && mapping.Name !== -1 && mapping.CheckIn !== -1) {
    const inboundNameClean = payload.name.toString().trim().toLowerCase();

    let inboundMatchString = "";
    if (payload.checkIn.includes("-")) {
      inboundMatchString = payload.checkIn;
    } else {
      const targetYear = dParsed.getFullYear();
      const targetMonth = String(dParsed.getMonth() + 1).padStart(2, '0');
      const targetDay = String(dParsed.getDate()).padStart(2, '0');
      inboundMatchString = `${targetYear}-${targetMonth}-${targetDay}`;
    }

    const displayValues = sheet.getDataRange().getDisplayValues();

    for (let i = 1; i < dataRange.length; i++) {
      const existingName = dataRange[i][mapping.Name] ? dataRange[i][mapping.Name].toString().trim().toLowerCase() : "";
      const existingCheckInText = displayValues[i][mapping.CheckIn] ? displayValues[i][mapping.CheckIn].toString().trim() : "";

      if (existingName === inboundNameClean && existingCheckInText !== "") {
        let existingDateParsed = parseDateSecurely(existingCheckInText);

        if (existingDateParsed) {
          const rowYear = existingDateParsed.getFullYear();
          const rowMonth = String(existingDateParsed.getMonth() + 1).padStart(2, '0');
          const rowDay = String(existingDateParsed.getDate()).padStart(2, '0');
          const rowMatchString = `${rowYear}-${rowMonth}-${rowDay}`;

          if (inboundMatchString === rowMatchString) {
            console.warn(`>>> [DUPLICATE BLOCKED] Match Found on Row ${i + 1}: ${inboundNameClean} on ${inboundMatchString}`);
            throw new Error(`Duplicate Entry Blocked: "${payload.name}" is already registered for ${standardizedCheckInString}.`);
          }
        }
      }
    }
  }

  // =========================================================================
  // --- 3. EXECUTE ROUTING & VECTOR PREPARATION (STRICT RESERVATION ID) ---
  // =========================================================================
  let targetRowNumber = NaN;
  let previousCheckInDate = null;
  let newRow = new Array(headers.length).fill("");

  if (mode === "EDIT") {
    const searchResId = payload.reservationId ? payload.reservationId.toString().trim().toLowerCase() : "";

    if (!searchResId) {
      throw new Error("Transaction Aborted: Reservation ID is required for editing.");
    }

    // Strict Search: Find row matching ReservationID ONLY
    for (let i = 1; i < dataRange.length; i++) {
      const sheetResId = dataRange[i][mapping.ReservationID] ? dataRange[i][mapping.ReservationID].toString().trim().toLowerCase() : "";
      if (sheetResId === searchResId) {
        targetRowNumber = i + 1;
        break;
      }
    }

    // Reject edit if Reservation ID was not found in the sheet
    if (isNaN(targetRowNumber) || targetRowNumber < 2) {
      throw new Error(`Transaction Aborted: No existing record found with Reservation ID "${payload.reservationId}".`);
    }

    console.log(`>>> [EDIT MATCH SUCCESS] Reservation ID [${payload.reservationId}] matched at Row [${targetRowNumber}]`);

    newRow = sheet.getRange(targetRowNumber, 1, 1, headers.length).getValues()[0];
    const oldDateVal = newRow[mapping.CheckIn];
    if (oldDateVal) {
      previousCheckInDate = parseDateSecurely(oldDateVal);
    }
  }

  // =========================================================================
  // --- 4. RESERVATION ID GENERATION / PRESERVATION ---
  // =========================================================================
  let currentReservationId = "";
  if (mode === "EDIT") {
    currentReservationId = newRow[mapping.ReservationID] ? newRow[mapping.ReservationID].toString().trim() : payload.reservationId;
  } else if (payload.reservationId && payload.reservationId.toString().trim() !== "") {
    currentReservationId = payload.reservationId.toString().trim();
  } else {
    currentReservationId = generateReservationIdHelper(payload.source, dParsed);
  }
  newRow[mapping.ReservationID] = currentReservationId;

  // --- 5. DATA MAP LAYER INJECTION ---
  if (mapping.Name !== -1) newRow[mapping.Name] = payload.name;
  if (mapping.Amount !== -1) newRow[mapping.Amount] = payload.totalBill;
  if (monthColIdx !== -1) newRow[monthColIdx] = generatedMonthLabel;

  if (mapping.Source !== -1) {
    let cleanSource = payload.source.toString().trim().toLowerCase();
    newRow[mapping.Source] = cleanSource.includes("airbnb") ? "AirBnb" : "Personal";
  }

  if (mapping.Floor !== -1) {
    let cleanFloor = payload.floor.toString().trim().toLowerCase();
    newRow[mapping.Floor] = cleanFloor.includes("second") ? "Second" : "Ground";
  }

  if (mapping.Mobile !== -1) newRow[mapping.Mobile] = cleanMobile;
  if (mapping.CheckIn !== -1) newRow[mapping.CheckIn] = standardizedCheckInString;
  if (mapping.Days !== -1) newRow[mapping.Days] = payload.days;
  if (mapping.Guests !== -1) newRow[mapping.Guests] = payload.guests;
  if (mapping.Ratings !== -1) newRow[mapping.Ratings] = payload.ratings;
  if (mapping.Comments !== -1) newRow[mapping.Comments] = payload.comments;

  if (mapping.Advance !== -1) newRow[mapping.Advance] = payload.advancePaid;
  if (mapping.Balance !== -1) newRow[mapping.Balance] = payload.balanceDue;

  if (mapping.BasePrice !== -1) newRow[mapping.BasePrice] = payload.basePrice;
  if (mapping.ExtraPax !== -1) newRow[mapping.ExtraPax] = payload.extraPaxFee;

  // =========================================================================
  // --- 6. COMMIT DATA CHANGES TO SPREADSHEETS ---
  // =========================================================================
  if (mode === "EDIT") {
    console.log(`>>> [EDIT SUCCESS] Modifying row [${targetRowNumber}] for Res ID [${currentReservationId}] safely in-place.`);
    sheet.getRange(targetRowNumber, 1, 1, headers.length).setValues([newRow]);
  } else {
    let currentSheetData = sheet.getDataRange().getValues();
    let lastRowWithContent = sheet.getLastRow();
    let incomingCheckInTime = dParsed.getTime();

    let insertionRowIndex = lastRowWithContent;
    let foundInsertionSpot = false;

    for (let i = currentSheetData.length - 1; i >= 1; i--) {
      let rowCheckInVal = currentSheetData[i][mapping.CheckIn];
      if (rowCheckInVal) {
        let rowDate = parseDateSecurely(rowCheckInVal);
        if (rowDate) {
          if (rowDate.getTime() <= incomingCheckInTime) {
            insertionRowIndex = i + 1;
            foundInsertionSpot = true;
            break;
          }
        }
      }
    }

    if (!foundInsertionSpot && lastRowWithContent > 1) {
      insertionRowIndex = 1;
    }

    console.log(`>>> [ADD SUCCESS] Inserting ${payload.name} safely after row: [${insertionRowIndex}]`);
    sheet.insertRowsAfter(insertionRowIndex, 1);

    let templateRow = (insertionRowIndex === 1) ? 2 : insertionRowIndex;
    let templateRange = sheet.getRange(templateRow, 1, 1, headers.length);
    let targetRange = sheet.getRange(insertionRowIndex + 1, 1, 1, headers.length);

    templateRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    targetRange.setValues([newRow]);
  }

  // =========================================================================
  // --- 7. AUTOMATED SHEET RE-SORTATION ENGINE ---
  // =========================================================================
  SpreadsheetApp.flush();
  const activeLastRow = sheet.getLastRow();
  if (activeLastRow > 1) {
    const sortTargetRange = sheet.getRange(2, 1, activeLastRow - 1, headers.length);
    sortTargetRange.sort({ column: mapping.CheckIn + 1, ascending: true });
  }

  if (previousCheckInDate && previousCheckInDate.getTime() !== dParsed.getTime()) {
    if (typeof deleteOldBookingFromCalendar === "function") {
      deleteOldBookingFromCalendar(payload.name, previousCheckInDate);
    }
  }

  syncBookingToVinyasaCalendar(
    currentReservationId, // 🔑 UNIQUE KEY PASSING
    payload.name,
    dParsed,
    payload.guests,
    payload.source,
    payload.floor,
    payload.days,
    payload.comments
  );

  SpreadsheetApp.flush();
  return "SUCCESS";
}

/**
 * Helper function to generate Reservation ID: CHANNEL-YYYYMMDD-HEX
 */
function generateReservationIdHelper(sourceStr, checkInDate) {
  const cleanSource = (sourceStr || '').toString().toLowerCase().trim();
  const prefix = cleanSource.includes('airbnb') ? 'ABN' : 'PER';

  let dateFormatted = '00000000';
  if (checkInDate && !isNaN(checkInDate.getTime())) {
    const yyyy = checkInDate.getFullYear();
    const mm = String(checkInDate.getMonth() + 1).padStart(2, '0');
    const dd = String(checkInDate.getDate()).padStart(2, '0');
    dateFormatted = `${yyyy}${mm}${dd}`;
  }

  const randomHex = Math.floor(Math.random() * 65536)
    .toString(16)
    .toUpperCase()
    .padStart(4, '0');

  return `${prefix}-${dateFormatted}-${randomHex}`;
}


/**
 * Cleanly purges an old calendar event across a wide search window if the guest dates change.
 */
function deleteOldBookingFromCalendar(guestName, oldCheckInDateInput) {
  try {
    if (!oldCheckInDateInput) return;

    const calendarName = "Vinyasa Nilaya";
    const calendars = CalendarApp.getCalendarsByName(calendarName);
    const targetCalendar = (calendars.length > 0) ? calendars[0] : CalendarApp.getDefaultCalendar();

    let oldDate = (oldCheckInDateInput instanceof Date) ? oldCheckInDateInput : new Date(oldCheckInDateInput);
    if (isNaN(oldDate.getTime())) return;

    // Open a safe search boundary (from 2 days before the old check-in to 15 days out)
    // to make sure we catch the multi-day visual block completely.
    const searchStart = new Date(oldDate.getTime());
    searchStart.setDate(searchStart.getDate() - 2);
    searchStart.setHours(0, 0, 0, 0);

    const searchEnd = new Date(oldDate.getTime());
    searchEnd.setDate(searchEnd.getDate() + 15);
    searchEnd.setHours(23, 59, 59, 999);

    const rangeEvents = targetCalendar.getEvents(searchStart, searchEnd);

    rangeEvents.forEach(event => {
      const currentTitle = event.getTitle();
      if (currentTitle.toLowerCase().includes(guestName.trim().toLowerCase())) {
        console.log(`>>> [CALENDAR CLEANER] Purging shifted residual block: "${currentTitle}"`);
        event.deleteEvent();
      }
    });
  } catch (err) {
    console.warn(`>>> [CALENDAR CLEANER WARNING] Failed to clear old event: ${err.toString()}`);
  }
}

/**
 * Core Database & Calendar Sync Writer: Purges an entire reservation entry row
 * by ReservationID (with Name/Date fallback) AND automatically wipes the matching Google Calendar event.
 */
function deleteGuestRowBackend(reservationId, name, checkInStr) {
  try {
    if (!checkInStr || checkInStr === "undefined") {
      throw new Error("Missing mandatory parameter: checkInStr evaluation failed.");
    }

    // 1. Resolve target spreadsheet year natively on the server side
    const checkInDateObj = new Date(checkInStr);
    if (isNaN(checkInDateObj.getTime())) {
      throw new Error(`Invalid date string received by server: '${checkInStr}'`);
    }

    const targetSheetName = checkInDateObj.getFullYear().toString(); // e.g. "2026"

    const ss = SpreadsheetApp.openById(ID_GUESTS_LIST);
    const sheet = ss.getSheetByName(targetSheetName);

    if (!sheet) {
      throw new Error(`Target year data registry tab '${targetSheetName}' could not be discovered.`);
    }

    const dataRangeValues = sheet.getDataRange().getValues();
    const headers = dataRangeValues[0].map(h => h ? h.toString().trim() : "");

    const resIdColIdx = headers.indexOf("ReservationID");
    const nameColIdx = headers.indexOf("Name");
    const checkInColIdx = headers.indexOf("Check-in Date");

    const cleanResId = (reservationId || "").toString().trim().toLowerCase();
    const targetNameClean = (name || "").toString().trim().toLowerCase();

    let rowToDelete = -1;

    // 2. Scan rows to find the exact match
    for (let row = dataRangeValues.length - 1; row >= 1; row--) {
      const rowResId = (resIdColIdx !== -1 && dataRangeValues[row][resIdColIdx])
        ? dataRangeValues[row][resIdColIdx].toString().trim().toLowerCase()
        : "";

      // PRIMARY MATCH: By ReservationID
      if (cleanResId !== "" && rowResId === cleanResId) {
        rowToDelete = row + 1; // Convert 0-index to 1-indexed sheet row
        break;
      }

      // SECONDARY FALLBACK MATCH: By Name + Check-in Date (for legacy un-backfilled rows)
      if (cleanResId === "" && nameColIdx !== -1 && checkInColIdx !== -1) {
        const rowName = dataRangeValues[row][nameColIdx] ? dataRangeValues[row][nameColIdx].toString().trim().toLowerCase() : "";
        const rowCheckInRaw = dataRangeValues[row][checkInColIdx];

        if (rowName === targetNameClean && rowCheckInRaw) {
          const rowDateObj = new Date(rowCheckInRaw);
          if (!isNaN(rowDateObj.getTime()) && rowDateObj.toDateString() === checkInDateObj.toDateString()) {
            rowToDelete = row + 1;
            break;
          }
        }
      }
    }

    if (rowToDelete === -1) {
      throw new Error(`The requested reservation '${reservationId || name}' could not be located in sheet '${targetSheetName}'.`);
    }

    // 3. Automated Calendar Purge Module
    try {
      const calendarName = "Vinyasa Nilaya";
      const calendars = CalendarApp.getCalendarsByName(calendarName);
      const targetCalendar = (calendars.length > 0) ? calendars[0] : CalendarApp.getDefaultCalendar();

      const startWindow = new Date(checkInDateObj.getFullYear(), checkInDateObj.getMonth(), checkInDateObj.getDate(), 0, 0, 0);
      const endWindow = new Date(checkInDateObj.getFullYear(), checkInDateObj.getMonth(), checkInDateObj.getDate(), 23, 59, 59);

      const events = targetCalendar.getEvents(startWindow, endWindow);
      let calendarDeletedCount = 0;

      events.forEach(ev => {
        const title = ev.getTitle();
        if (title.toLowerCase().includes(targetNameClean)) {
          ev.deleteEvent();
          calendarDeletedCount++;
        }
      });

      if (calendarDeletedCount > 0) {
        console.log(`✨ [CALENDAR SYNC] Removed ${calendarDeletedCount} timeline blocks for ${name}.`);
      }
    } catch (calErr) {
      console.error("❌ [CALENDAR PURGE ERROR] Sync skipped:", calErr.toString());
    }

    // 4. Delete row from Sheet
    sheet.deleteRow(rowToDelete);
    SpreadsheetApp.flush();

    console.log(`>>> [BACKEND REMOVAL SUCCESS] Deleted row [${rowToDelete}] for Reservation ID: ${reservationId || 'N/A'} (${name})`);
    return `SUCCESS: Row ${rowToDelete} deleted cleanly.`;

  } catch (err) {
    console.error(">>> [DELETE GUEST BACKEND FATAL ERROR]", err);
    throw new Error(err.message);
  }
}

/**
 * HELPER FUNCTION: Safely parses dates from strings, objects, or numbers
 */
function parseDateSecurely(dateVal) {
  if (!dateVal) return null;
  if (dateVal instanceof Date) {
    return !isNaN(dateVal.getTime()) ? dateVal : null;
  }

  let dateStr = dateVal.toString().trim();
  let parsed = new Date(dateStr);
  if (!isNaN(parsed.getTime())) return parsed;

  const monthsMap = {
    jan: 0, feb: 1, mar: 2, apr: 3, may: 4, jun: 5, jul: 6, aug: 7, sep: 8, oct: 9, nov: 10, dec: 11,
    january: 0, february: 1, march: 2, april: 3, may: 4, june: 5, july: 6, august: 7, september: 8, october: 9, november: 10, december: 11
  };

  let tokens = dateStr.replace(/,/g, "").split(/\s+/);
  if (tokens.length >= 2) {
    let day = parseInt(tokens[1]);
    let monthStr = tokens[0].toLowerCase();
    let year = parseInt(tokens[2]) || new Date().getFullYear();

    if (isNaN(day)) {
      day = parseInt(tokens[0]);
      monthStr = tokens[1].toLowerCase();
    }

    if (!isNaN(day) && monthsMap[monthStr] !== undefined) {
      return new Date(year, monthsMap[monthStr], day);
    }
  }

  return null;
}

/***************************************************************************
 * FIRESTORE REGISTRY CORE MODULE
 ***************************************************************************/

/**
 * Fetch verified guests from Firestore collection 'guests'
 */
function fetchFirestoreGuestsRegistry() {
  try {
    const url = `https://firestore.googleapis.com/v1/projects/${FIREBASE_PROJECT_ID}/databases/(default)/documents/guests`;
    const response = UrlFetchApp.fetch(url, {
      headers: { "Authorization": "Bearer " + ScriptApp.getOAuthToken() }
    });
    const result = JSON.parse(response.getContentText());
    
    if (!result.documents) return [];

    return result.documents.map(doc => {
      const f = doc.fields;
      const docId = doc.name.split('/').pop();
      
      // Mapping nested fields to match the finalSubmit payload structure
      return {
        id: docId,
        timestamp: f.createdAt ? f.createdAt.timestampValue : (f.timestamp ? f.timestamp.timestampValue : ""),
        name: (f.guestDetails && f.guestDetails.mapValue.fields.name) ? f.guestDetails.mapValue.fields.name.stringValue : "Unknown",
        phone: (f.guestDetails && f.guestDetails.mapValue.fields.phone) ? f.guestDetails.mapValue.fields.phone.stringValue : "-",
        idType: (f.verification && f.verification.mapValue.fields.idType) ? f.verification.mapValue.fields.idType.stringValue : "Govt ID",
        idNo: (f.verification && f.verification.mapValue.fields.idNo) ? f.verification.mapValue.fields.idNo.stringValue : "-",
        frontUrl: (f.verification && f.verification.mapValue.fields.idFrontUrl) ? f.verification.mapValue.fields.idFrontUrl.stringValue : "",
        backUrl: (f.verification && f.verification.mapValue.fields.idBackUrl) ? f.verification.mapValue.fields.idBackUrl.stringValue : "",
        arrivingCity: (f.travelDetails && f.travelDetails.mapValue.fields.arrivingCity) ? f.travelDetails.mapValue.fields.arrivingCity.stringValue : "-",
        purpose: (f.travelDetails && f.travelDetails.mapValue.fields.purpose) ? f.travelDetails.mapValue.fields.purpose.stringValue : "-",
        emergencyName: (f.emergencyContact && f.emergencyContact.mapValue.fields.name) ? f.emergencyContact.mapValue.fields.name.stringValue : "-",
        emergencyPhone: (f.emergencyContact && f.emergencyContact.mapValue.fields.phone) ? f.emergencyContact.mapValue.fields.phone.stringValue : "-",
        selfieUrl: f.selfieUrl ? f.selfieUrl.stringValue : "",
        checkinStatus: f.verifiedStatus ? f.verifiedStatus.stringValue : "Verified",
        address: (f.travelDetails && f.travelDetails.mapValue.fields.arrivingCity) ? f.travelDetails.mapValue.fields.arrivingCity.stringValue : "-"
      };
    }).sort((a, b) => {
      // Client-side sort to mirror orderBy("createdAt", "desc")
      const dateA = new Date(a.timestamp || 0);
      const dateB = new Date(b.timestamp || 0);
      return dateB - dateA;
    });
  } catch (err) {
    console.error("Firestore fetch error: ", err);
    throw new Error(err.message);
  }
}

/**
 * Delete Firestore guest record by Document ID
 */
function deleteFirestoreGuestRecord(docId) {
  try {
    const docUrl = `https://firestore.googleapis.com/v1/projects/${FIREBASE_PROJECT_ID}/databases/(default)/documents/guests/${docId}`;
    
    // 1. Fetch the document to get image references BEFORE deletion
    const response = UrlFetchApp.fetch(docUrl, {
      headers: { "Authorization": "Bearer " + ScriptApp.getOAuthToken() }
    });
    const doc = JSON.parse(response.getContentText());
    const f = doc.fields;

    // 2. Identify potential storage assets based on your schema
    const assets = [];
    if (f.verification && f.verification.mapValue && f.verification.mapValue.fields.idFrontUrl) assets.push(f.verification.mapValue.fields.idFrontUrl.stringValue);
    if (f.verification && f.verification.mapValue && f.verification.mapValue.fields.idBackUrl) assets.push(f.verification.mapValue.fields.idBackUrl.stringValue);
    if (f.selfieUrl) assets.push(f.selfieUrl.stringValue);

    // 3. Attempt to delete files from Firebase Storage
    assets.forEach(url => {
      if (url && url.startsWith('http')) {
        try {
          deleteFileFromFirebaseStorage(url);
        } catch (storageErr) {
          console.warn(`[STORAGE CLEANUP] Could not delete asset ${url}: ${storageErr.message}`);
        }
      }
    });

    // 4. Finally, delete the Firestore metadata record
    UrlFetchApp.fetch(docUrl, {
      method: "delete",
      headers: { "Authorization": "Bearer " + ScriptApp.getOAuthToken() }
    });

    return "SUCCESS";
  } catch (err) {
    console.error("Firestore delete error: ", err);
    throw new Error(err.message);
  }
}

/**
 * Fetch Firebase Storage asset and proxy to Base64
 */
function getFirebaseStorageImage(rawUrl) {
  try {
    if (!rawUrl) return "";
    const response = UrlFetchApp.fetch(rawUrl);
    const blob = response.getBlob();
    const contentType = blob.getContentType();
    const base64 = Utilities.base64Encode(blob.getBytes());
    return `data:${contentType};base64,${base64}`;
  } catch (err) {
    console.error("Firebase Storage Proxy Error: ", err);
    return "ERROR";
  }
}

/**
 * Update Firestore guest record
 */
function updateFirestoreGuestRecord(docId, updates) {
  try {
    // Define all nested field paths that need to be updated.
    // Dots represent nesting within Firestore maps.
    const mask = [
      'updateMask.fieldPaths=guestDetails.name',
      'updateMask.fieldPaths=guestDetails.phone',
      'updateMask.fieldPaths=verification.idNo',
      'updateMask.fieldPaths=travelDetails.arrivingCity',
      'updateMask.fieldPaths=travelDetails.address',
      'updateMask.fieldPaths=emergencyContact.name',
      'updateMask.fieldPaths=emergencyContact.phone'
    ].join('&');

    const url = `https://firestore.googleapis.com/v1/projects/${FIREBASE_PROJECT_ID}/databases/(default)/documents/guests/${docId}?${mask}`;

    const payload = {
      fields: {
        guestDetails: {
          mapValue: {
            fields: {
              name: { stringValue: updates.name },
              phone: { stringValue: updates.phone }
            }
          }
        },
        verification: {
          mapValue: { fields: { idNo: { stringValue: updates.idNo } } }
        },
        travelDetails: {
          mapValue: { fields: { 
            arrivingCity: { stringValue: updates.address },
            address: { stringValue: updates.address }
          } }
        },
        emergencyContact: {
          mapValue: {
            fields: {
              name: { stringValue: updates.emergencyName },
              phone: { stringValue: updates.emergencyPhone }
            }
          }
        }
      }
    };
    UrlFetchApp.fetch(url, {
      method: "patch",
      contentType: "application/json",
      headers: { "Authorization": "Bearer " + ScriptApp.getOAuthToken() },
      payload: JSON.stringify(payload)
    });
    return "SUCCESS";
  } catch (err) {
    console.error("Firestore update error: ", err);
    throw new Error(err.message);
  }
}

/**
 * Batch delete Firestore documents
 */
function deleteFirestoreGuestsBatch(ids) {
  ids.forEach(id => deleteFirestoreGuestRecord(id));
  return "SUCCESS";
}

/**
 * Inventory available years from Firestore timestamps
 */
function getFirestoreRegistryConfig() {
  const currentYear = new Date().getFullYear().toString();
  try {
    const records = fetchFirestoreGuestsRegistry();
    const years = [...new Set(records.map(r => {
      if (!r.timestamp) return currentYear;
      const d = new Date(r.timestamp);
      return isNaN(d.getTime()) ? currentYear : d.getFullYear().toString();
    }))];
    years.sort((a,b) => b - a);
    return {
      years: years.length ? years : [currentYear],
      activeYear: currentYear
    };
  } catch (e) {
    return { years: [currentYear], activeYear: currentYear };
  }
}

/*** Google calender sync */
/**
 * Global Sync Engine: Creates or updates an all-day event on the Vinyasa Nilaya calendar
 * Attaches a professional metadata overview card and locks an alert reminder exactly 1 day prior.
 */
/*** Google calendar sync */
/**
 * Global Sync Engine: Creates or updates an all-day event on the Vinyasa Nilaya calendar
 * Uses Reservation ID as the unique immutable key to prevent accidental name collisions.
 */
function syncBookingToVinyasaCalendar(reservationId, guestName, checkInDateInput, totalGuests, platformType, floorName, nights, notes) {
  try {
    const calendarName = "Vinyasa Nilaya";
    const calendars = CalendarApp.getCalendarsByName(calendarName);
    const targetCalendar = (calendars.length > 0) ? calendars[0] : CalendarApp.getDefaultCalendar();

    // 1. Resolve date object coordinates safely across input variants
    let checkInDate = (checkInDateInput instanceof Date) ? checkInDateInput : new Date(checkInDateInput);
    if (isNaN(checkInDate.getTime())) {
      checkInDate = parseDateSecurely(checkInDateInput) || new Date();
    }

    // Calculate the precise check-out date based on duration
    const totalNights = parseInt(nights, 10) || 1;
    let checkOutDate = new Date(checkInDate.getTime());
    checkOutDate.setDate(checkOutDate.getDate() + totalNights);

    // Ensure valid Reservation ID fallback
    const resId = (reservationId && reservationId.toString().trim() !== "") 
      ? reservationId.toString().trim() 
      : `RES-${Date.now()}`;

    // 2. Blueprint Title with explicit Reservation ID tag
    const eventTitle = `Guest Check-In: ${guestName} (${platformType || 'Booking'}) [ID: ${resId}]`;

    // 3. Render professional text metadata card
    const eventDescription = [
      `=========================================`,
      `         VINYASA NILAYA RESERVATION       `,
      `=========================================`,
      `Reservation ID : ${resId}`,
      `Guest Name     : ${guestName}`,
      `Booking Channel: ${platformType || 'Direct Personal'}`,
      `Allocated Floor: ${floorName || 'Ground'} Floor`,
      `Total Guests   : ${totalGuests || 1} Pax`,
      `Duration       : ${totalNights} Night(s)`,
      `Check-In Date  : ${checkInDate.toDateString()}`,
      `Check-Out Date : ${checkOutDate.toDateString()}`,
      `-----------------------------------------`,
      `Reference Information / Notes:`,
      `${notes || 'No operational comments attached.'}`,
      `=========================================`,
      `Synced automatically via Vinyasa Workspace Integration Hub.`
    ].join('\n');

    // 4. Search search window across wider boundary to find matching Reservation ID
    const searchStart = new Date(checkInDate.getTime());
    searchStart.setDate(searchStart.getDate() - 30); // Look 30 days back in case check-in shifted
    const searchEnd = new Date(checkOutDate.getTime());
    searchEnd.setDate(searchEnd.getDate() + 30);   // Look 30 days forward

    const rangeEvents = targetCalendar.getEvents(searchStart, searchEnd);

    // 🔑 ACCURATE KEY MATCHING: Purge existing event matching Reservation ID specifically
    rangeEvents.forEach(event => {
      const title = event.getTitle();
      const description = event.getDescription() || "";

      // Check if title or description contains the unique Reservation ID
      if (title.includes(`[ID: ${resId}]`) || description.includes(`Reservation ID : ${resId}`)) {
        console.log(`>>> [CALENDAR ENGINE] Dropping existing block for Reservation ID [${resId}]: "${title}"`);
        event.deleteEvent();
      }
    });

    // 5. Create fresh multi-day event
    const targetEvent = targetCalendar.createAllDayEvent(eventTitle, checkInDate, checkOutDate, {
      description: eventDescription
    });
    console.log(`>>> [CALENDAR ENGINE] Created event for: ${guestName} | Res ID: ${resId}`);

    // 6. ENFORCE 24-HOUR ADVANCE REMINDERS
    targetEvent.removeAllReminders();
    targetEvent.addPopupReminder(1440);
    targetEvent.addEmailReminder(1440);

    return true;
  } catch (err) {
    console.warn(`>>> [CALENDAR INTERCEPT SYSTEM WARNING] Sync bypassed: ${err.toString()}`);
    return false;
  }
}

/**
 * Scans a date range day-by-day to find the exact window of availability for each floor.
 * 
 * @param {string} fromDateInput - Check-in date (YYYY-MM-DD)
 * @param {string} toDateInput - Check-out date (YYYY-MM-DD)
 * @return {Object} Dynamic availability strings for each floor
 */
function checkFloorWiseAvailability(fromDateInput, toDateInput) {
  try {
    const calendarName = "Vinyasa Nilaya";
    const calendars = CalendarApp.getCalendarsByName(calendarName);
    const targetCalendar = (calendars.length > 0) ? calendars[0] : CalendarApp.getDefaultCalendar();

    let startCheck = (fromDateInput instanceof Date) ? fromDateInput : new Date(fromDateInput);
    if (isNaN(startCheck.getTime()) && typeof parseDateSecurely === 'function') startCheck = parseDateSecurely(fromDateInput);

    let endCheck = (toDateInput instanceof Date) ? toDateInput : new Date(toDateInput);
    if (isNaN(endCheck.getTime()) && typeof parseDateSecurely === 'function') endCheck = parseDateSecurely(toDateInput);

    if (isNaN(startCheck.getTime()) || isNaN(endCheck.getTime())) {
      throw new Error("Invalid date inputs.");
    }

    // Calculate total nights to scan
    const totalNights = Math.round((endCheck.getTime() - startCheck.getTime()) / (1000 * 60 * 60 * 24));

    // Tracks status night-by-night (true = free, false = booked)
    let groundNights = [];
    let secondNights = [];
    let dateObjects = [];

    // 1. Scan night-by-night
    for (let i = 0; i < totalNights; i++) {
      let currentNightStart = new Date(startCheck.getTime());
      currentNightStart.setDate(currentNightStart.getDate() + i);
      dateObjects.push(new Date(currentNightStart.getTime())); // Store for reference

      // Set standard check-in window boundary (12:00 PM) to check-out boundary (11:00 AM next day)
      let nightStart = new Date(currentNightStart.getTime());
      nightStart.setHours(12, 0, 0, 0);

      let nightEnd = new Date(currentNightStart.getTime());
      nightEnd.setDate(nightEnd.getDate() + 1);
      nightEnd.setHours(11, 0, 0, 0);

      let rangeEvents = targetCalendar.getEvents(nightStart, nightEnd);

      let isGroundFree = true;
      let isSecondFree = true;

      rangeEvents.forEach(event => {
        const combinedContext = ((event.getTitle() || "") + " " + (event.getDescription() || "")).toLowerCase();
        if (combinedContext.includes("ground")) isGroundFree = false;
        if (combinedContext.includes("second")) isSecondFree = false;
      });

      groundNights.push(isGroundFree);
      secondNights.push(isSecondFree);
    }

    // 2. Helper function to generate smart availability phrasing
    const processFloorStatus = (nightsArray) => {
      const total = nightsArray.length;
      const freeCount = nightsArray.filter(Boolean).length;

      if (freeCount === total) return { status: "AVAILABLE", text: "Available" };
      if (freeCount === 0) return { status: "BOOKED", text: "Booked" };

      // Find the last consecutive run of available nights that reaches the end date
      let availableFromIndex = -1;
      for (let i = total - 1; i >= 0; i--) {
        if (nightsArray[i]) {
          availableFromIndex = i;
        } else {
          break; // Hit a booking block moving backward
        }
      }

      if (availableFromIndex > 0) {
        const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
        let freeDate = dateObjects[availableFromIndex];
        let formattedDate = `${freeDate.getDate()}-${months[freeDate.getMonth()]}`;
        return { status: "PARTIAL", text: `Available from ${formattedDate}` };
      }

      return { status: "BOOKED", text: "Partially Booked" };
    };

    return {
      ground: processFloorStatus(groundNights),
      second: processFloorStatus(secondNights)
    };

  } catch (err) {
    console.error(">>> [SMART SCANNED ERROR] " + err.toString());
    return {
      ground: { status: "BOOKED", text: "Error" },
      second: { status: "BOOKED", text: "Error" }
    };
  }
}

/**** QID fileter and bulk delte feature */
/**
 * Feature 2 Backend: Iterates over selected row vectors, wipes Drive file allocations,
 * and handles consecutive matrix indexing contractions from highest index down.
 *
 * @param {Array<string|number>} slNoArray - Unified collection of Serial mapping records.
 * @return {string} Confirmation token feed.
 */
function deleteQidRowsBatchBackend(slNoArray) {
  if (!slNoArray || !Array.isArray(slNoArray) || slNoArray.length === 0) {
    throw new Error("Invalid selection payload collection provided.");
  }

  try {
    const ss = SpreadsheetApp.openById(ID_QID_VERIFIED_LIST);
    const sheet = ss.getSheetByName(SHEET_NAME_QID);
    if (!sheet) throw new Error(`Target tab config "${SHEET_NAME_QID}" missing.`);

    const numericalSlNos = slNoArray.map(id => parseInt(id));
    const values = sheet.getDataRange().getValues();
    let deletedRowsRecordList = [];

    // 1. CLEAR ASSOCIATED DRIVE STORAGE ASSETS
    for (let i = 1; i < values.length; i++) {
      const currentSlNo = parseInt(values[i][0]);

      if (numericalSlNos.includes(currentSlNo)) {
        const targetRowData = values[i];
        const frontUrl = targetRowData[10] || "";
        const backUrl = targetRowData[11] || "";
        const selfieUrl = targetRowData[12] || "";
        const filesToPurge = [frontUrl, backUrl, selfieUrl];

        filesToPurge.forEach(url => {
          if (url && url.toString().trim() !== "") {
            const fileId = extractDriveIdSafely(url.toString().trim());
            if (fileId) {
              try {
                DriveApp.getFileById(fileId).setTrashed(true);
              } catch (fErr) {
                console.warn(`[BATCH PURGE SKIP] ID ${fileId} inaccessible: ${fErr.message}`);
              }
            }
          }
        });

        // Store the original sheet row index coordinate (1-indexed mapping adjustment)
        deletedRowsRecordList.push(i + 1);
      }
    }

    // 2. CRITICAL STEP: Sort row indices in DESCENDING order before deleting.
    // If you delete row 5 first, row 10 shifts up to row 9, causing data misalignment.
    // Deleting from the bottom up completely bypasses this indexing bug.
    deletedRowsRecordList.sort((a, b) => b - a);

    deletedRowsRecordList.forEach(rowIndex => {
      sheet.deleteRow(rowIndex);
    });

    SpreadsheetApp.flush();
    console.log(`>>> [BATCH DELETION SUCCESS] Successfully purged ${deletedRowsRecordList.length} records from ledger.`);
    return "SUCCESS";

  } catch (err) {
    console.error(">>> [BATCH DELETION EXCEPTION BLOCK] Action failed: ", err);
    throw new Error(err.message);
  }
}

/**
 * Backend API: Scans a parent Google Drive directory for folders named "QID-YYYY",
 * extracts unique year tokens, and returns them ordered with the current active year.
 * * @return {Object} An inventory of available years and the active calendar default year tag.
 */
function getDynamicQidYearsConfig() {
  // CONSTANT PARAMETER: Replace with the exact Folder ID where your QID folders live
  const PARENT_FOLDER_ID = "YOUR_MASTER_ROOT_DRIVE_FOLDER_ID";

  let detectedYears = [];
  const currentCalendarYear = new Date().getFullYear().toString(); // Default fallback "2026"

  try {
    const parentFolder = DriveApp.getFolderById(PARENT_FOLDER_ID);
    const subFolders = parentFolder.getFolders();

    // Regular Expression targeting patterns like "QID-2026" or "QID-2027"
    const pattern = /^QID-(\d{4})$/;

    while (subFolders.hasNext()) {
      const folder = subFolders.next();
      const match = folder.getName().trim().match(pattern);

      if (match && match[1]) {
        const yearValue = match[1];
        if (!detectedYears.includes(yearValue)) {
          detectedYears.push(yearValue);
        }
      }
    }

    // Sort years chronologically in descending order (newest years first)
    detectedYears.sort((a, b) => parseInt(b) - parseInt(a));

    // Failsafe condition: If no folders are matched, append current year to keep UI operational
    if (detectedYears.length === 0) {
      detectedYears.push(currentCalendarYear);
    }

    console.log(`>>> [SERVER DRIVE ARMED] Detected dynamic QID year sets: [${detectedYears.join(', ')}]`);

    return {
      years: detectedYears,
      activeYear: detectedYears.includes(currentCalendarYear) ? currentCalendarYear : detectedYears[0]
    };

  } catch (err) {
    console.error(">>> [SERVER EXCEPTION CRASH] Dynamic folder parsing dropped unexpected exception: ", err);
    // Hard fallback layout parameter return
    return {
      years: [currentCalendarYear],
      activeYear: currentCalendarYear
    };
  }
}

/**
 * Helper to delete a file from Firebase Storage using its Download URL via REST API
 */
function deleteFileFromFirebaseStorage(downloadUrl) {
  // Extract bucket and encoded path from the URL
  // Format: https://firebasestorage.googleapis.com/v0/b/[bucket]/o/[path]?alt=media
  const parts = downloadUrl.split('/o/');
  if (parts.length < 2) return;
  
  const bucket = parts[0].split('/b/')[1];
  const encodedPath = parts[1].split('?')[0];
  
  const deleteUrl = `https://firebasestorage.googleapis.com/v0/b/${bucket}/o/${encodedPath}`;
  
  return UrlFetchApp.fetch(deleteUrl, {
    method: "delete",
    headers: { "Authorization": "Bearer " + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true // Prevents whole batch from failing if one image is already gone
  });
}

/**
 * ONE-TIME BACKFILL (2026 ONWARDS):
 * Run directly from Apps Script Editor (Code.gs).
 * Generates unique Reservation IDs ONLY for rows in the "2026" sheet (and future years),
 * skipping all older legacy sheets.
 */
function backfillReservationIds2026Onwards() {
  const ss = SpreadsheetApp.openById(ID_GUESTS_LIST);
  const sheets = ss.getSheets();

  let totalUpdated = 0;

  sheets.forEach(sheet => {
    const sheetName = sheet.getName().trim();

    // Parse sheet name as year (e.g., "2026", "2027")
    const sheetYear = parseInt(sheetName, 10);

    // Filter: Ignore sheets that are not numeric years OR are strictly older than 2026
    if (isNaN(sheetYear) || sheetYear < 2026) {
      Logger.log(`>>> [IGNORE] Skipping sheet "${sheetName}" (Older than 2026 or non-year sheet).`);
      return;
    }

    let lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      Logger.log(`>>> [SKIP] Sheet "${sheetName}" has no data rows.`);
      return;
    }

    Logger.log(`>>> [PROCESSING] Evaluating sheet "${sheetName}"...`);

    // Read full grid for the target sheet
    let fullRange = sheet.getDataRange().getValues();
    let headers = fullRange[0].map(h => h ? h.toString().trim() : "");

    let resIdColIdx = headers.indexOf("ReservationID");

    // 1. If missing, insert "ReservationID" into Column 15 (Column O, before Comments)
    if (resIdColIdx === -1) {
      Logger.log(`>>> [SETUP] Inserting missing "ReservationID" column at position 15 in Sheet: "${sheetName}"`);

      if (sheet.getLastColumn() >= 15) {
        sheet.insertColumnBefore(15);
      } else {
        sheet.insertColumnsAfter(sheet.getLastColumn(), 15 - sheet.getLastColumn());
      }

      sheet.getRange(1, 15).setValue("ReservationID");
      SpreadsheetApp.flush();

      // Re-fetch range to keep in-memory positions aligned
      fullRange = sheet.getDataRange().getValues();
      headers = fullRange[0].map(h => h ? h.toString().trim() : "");
      resIdColIdx = headers.indexOf("ReservationID");
    }

    // 2. Locate helper columns dynamically (case-insensitive)
    const checkInColIdx = headers.findIndex(h => {
      const lower = h.toLowerCase();
      return lower.includes("check-in") || lower.includes("checkin");
    });

    const sourceColIdx = headers.findIndex(h => {
      const lower = h.toLowerCase();
      return lower.includes("airbnb") || lower.includes("source") || lower.includes("personal");
    });

    // Prepare batch update array
    const numDataRows = fullRange.length - 1;
    const updateColumnVector = [];
    let sheetUpdateCount = 0;

    for (let i = 1; i < fullRange.length; i++) {
      let existingResId = fullRange[i][resIdColIdx];

      if (!existingResId || existingResId.toString().trim() === "") {
        const checkInRaw = checkInColIdx !== -1 ? fullRange[i][checkInColIdx] : null;
        const sourceRaw = sourceColIdx !== -1 ? fullRange[i][sourceColIdx] : "Personal";

        // Parse date securely
        let parsedDate = new Date();
        if (checkInRaw) {
          if (typeof parseDateSecurely === "function") {
            parsedDate = parseDateSecurely(checkInRaw);
          } else {
            const tempDate = new Date(checkInRaw);
            if (!isNaN(tempDate.getTime())) parsedDate = tempDate;
          }
        }

        // Generate ID
        const generatedId = generateReservationIdHelper(sourceRaw, parsedDate);
        updateColumnVector.push([generatedId]);
        sheetUpdateCount++;
        totalUpdated++;
      } else {
        // Retain existing ID
        updateColumnVector.push([existingResId]);
      }
    }

    // 3. Write updates back to Google Sheet in a single batch
    if (sheetUpdateCount > 0) {
      sheet.getRange(2, resIdColIdx + 1, numDataRows, 1).setValues(updateColumnVector);
      Logger.log(`>>> [SHEET UPDATED] "${sheetName}": Backfilled ${sheetUpdateCount} missing IDs.`);
    } else {
      Logger.log(`>>> [NO ACTION] "${sheetName}": All rows already have valid Reservation IDs.`);
    }
  });

  SpreadsheetApp.flush();
  Logger.log(`=================================================================`);
  Logger.log(`>>> [BACKFILL COMPLETE] Assigned Reservation IDs to ${totalUpdated} rows in 2026+ sheets.`);
  Logger.log(`=================================================================`);
}
/********************* Test functions *************************/
/**
 * Run this function to see exactly what getDashboardData is producing
 * and where it might be failing.
 */
function debugDashboard() {
  const testYear = "2026"; // Change to a year that exists in your tabs
  const testMonth = "Febrauary"; // Change to a month that has data

  console.log(`--- DEBUG START: Year[${testYear}] Month[${testMonth}] ---`);

  try {
    const result = getDashboardData(testYear, testMonth);

    // 1. Check the Summary object
    console.log("Summary Result:", JSON.stringify(result.summary, null, 2));

    // 2. Check the Guest Count
    console.log("Number of guests found:", result.guests.length);

    // 3. Inspect the first guest's data structure
    if (result.guests.length > 0) {
      console.log("First Guest Data Mapping (Sample):");
      console.log(JSON.stringify(result.guests[0], null, 2));

      // Check specific keys that usually cause issues
      const sample = result.guests[0];
      console.log("Key Check - Name:", sample.Name);
      console.log("Key Check - Source (AirBnb_Personal):", sample.AirBnb_Personal);
      console.log("Key Check - Amount:", sample.Amount);
    } else {
      console.warn("⚠️ No guests found. Possible causes: Tab name mismatch, Month spelling mismatch, or Header row not found.");
    }

  } catch (e) {
    console.error("❌ CRITICAL ERROR during execution:");
    console.error("Message: " + e.message);
    console.error("Stack: " + e.stack);
  }

  console.log("--- DEBUG END ---");
}

/**
 * Helper to check all tab names in your spreadsheet
 * Run this if you get 'null' errors to verify your sheet naming.
 */
function listAllTabNames() {
  const ss = SpreadsheetApp.openById(ID_GUESTS_LIST);
  const sheets = ss.getSheets();
  console.log("Available Tabs in this Spreadsheet:");
  sheets.forEach(s => console.log("- " + s.getName()));
}

/**
 * Test function to verify the Airbnb email parser logic.
 * This runs locally in the Apps Script editor and logs the extracted data.
 */
function test_syncAirbnbEmails() {
  try {
    console.log(">>> [TEST] Starting Airbnb Sync Parser Test...");

    // 1. Target Sri Harsha's specific thread using the exact subject line from your PDF
    const searchQuery = 'from:automated@airbnb.com subject:"Reservation confirmed - Sri Harsha Kuchimanchi arrives May 30"';
    const threads = GmailApp.search(searchQuery, 0, 1);

    if (threads.length === 0) {
      console.warn(">>> [TEST] Could not find Sri Harsha's email thread. Ensure the email is in your inbox/trash and hasn't been permanently deleted.");
      return;
    }

    syncAirbnbEmails("2026");

    const message = threads[0].getMessages()[0];
    const subject = message.getSubject();
    const body = message.getPlainBody();

    console.log(">>> [TEST] Found Email Subject: " + subject);

    // 2. Execute the exact Regular Expressions used in your main function
    const currentYear = "2026";
    const subjectMatch = subject.match(/Reservation confirmed\s*-\s*(.*?)\s+arrives\s+(.*)/i);

    let guestName = "";
    let checkInStr = "";

    if (subjectMatch) {
      guestName = subjectMatch[1].trim();
      checkInStr = subjectMatch[2].trim() + `, ${currentYear}`;
    }

    const nightsMatch = body.match(/(\d+)\s*nights\s*room\s*fee/i) || body.match(/(\d+)\s*night/i);
    const nights = nightsMatch ? Number(nightsMatch[1]) : 1;

    const amountMatch = body.match(/You earn[\s\S]*?₹?\s*([\d,]+\.?\d*)/i);
    let finalAmount = "0";
    if (amountMatch) {
      let rawAmt = amountMatch[1].replace(/,/g, "");
      finalAmount = Math.round(parseFloat(rawAmt)).toString();
    }

    const guestsMatch = body.match(/(\d+)\s*adults/i) || body.match(/(\d+)\s*guest/i);
    const totalGuests = guestsMatch ? Number(guestsMatch[1]) : 2;

    // 3. Output the parsed results directly to the execution log
    console.log("----------------------------------------");
    console.log(">>> [TEST] PARSE RESULTS:");
    console.log("Parsed Name: ", guestName, (guestName === "Sri Harsha Kuchimanchi" ? "✅ MATCH" : "❌ MISMATCH"));
    console.log("Parsed Check-in: ", checkInStr, (checkInStr === "May 30, 2026" ? "✅ MATCH" : "❌ MISMATCH"));
    console.log("Parsed Nights: ", nights, (nights === 3 ? "✅ MATCH" : "❌ MISMATCH"));
    console.log("Parsed Payout Amount: ", finalAmount, (finalAmount === "4586" ? "✅ MATCH" : "❌ MISMATCH"));
    console.log("Parsed Guests: ", totalGuests, (totalGuests === 2 ? "✅ MATCH" : "❌ MISMATCH"));
    console.log("----------------------------------------");

  } catch (err) {
    console.error(">>> [TEST] Test function execution crashed: " + err.message);
  }
}

/**
 * Utility function to clear the 'Vinyasa-Synced' label from Sri Harsha's email 
 * if you want to test the real button multiple times.
 */
function debug_removeSyncLabel() {
  const label = GmailApp.getUserLabelByName("Vinyasa-Synced");
  if (!label) {
    console.log("Label 'Vinyasa-Synced' does not exist yet.");
    return;
  }
  const threads = GmailApp.search('subject:"Sri Harsha Kuchimanchi arrives May 30" label:Vinyasa-Synced');
  threads.forEach(t => {
    t.removeLabel(label);
    console.log("Removed sync label from thread: " + t.getFirstMessageSubject());
  });
}

/**
 * TEST HARNESS: Run this function directly inside the Apps Script Editor 
 * to debug and inspect writeGuestDataRow behavior.
 */
function debug_writeGuestDataRow_Suite() {
  console.log("=== 🧪 STARTING writeGuestDataRow DEBUG SUITE 🧪 ===");

  // Choose a real or test year sheet tab present in your spreadsheet
  const testYear = "2026";

  // -----------------------------------------------------------------
  // TEST CASE 1: INSERT NEW RECORD (ADD MODE)
  // -----------------------------------------------------------------
  const addPayload = {
    name: "Test Guest Debugger",
    mobile: "9999988888",
    amount: 2500,
    guests: 3,
    checkIn: "2026-05-25", // Will convert to "May"
    days: 2,
    source: "Personal",
    floor: "Second Floor",
    ratings: "⭐⭐⭐⭐⭐",
    comments: "Created via automated GAS test runner suite execution.",
    year: testYear,
    rowIndex: "" // Blank for fresh additions
  };

  console.log("\n▶️ [TEST 1] Dispatching ADD payload for:", addPayload.name);
  try {
    const addResult = writeGuestDataRow("ADD", addPayload);
    console.log("✅ [TEST 1 SUCCESS] Backend returned response:", addResult);
  } catch (error) {
    console.error("❌ [TEST 1 FAILED] Execution crashed with error:", error.message);
  }

  // -----------------------------------------------------------------
  // TEST CASE 2: MODIFY EXISTING RECORD (EDIT MODE)
  // -----------------------------------------------------------------
  // We will pass the same name/mobile to update the entry we just made
  const editPayload = {
    name: "Test Guest Debugger",
    mobile: "9999988888",
    amount: 3200, // Modifying amount from 2500 to 3200
    guests: 3,
    checkIn: "2026-05-25",
    days: 3,      // Modifying nights from 2 to 3
    source: "Airbnb", // Modifying source from Personal to Airbnb
    floor: "Second Floor",
    ratings: "⭐⭐⭐⭐", // Modifying ratings
    comments: "Updated successfully via test runner execution script.",
    year: testYear,
    rowIndex: "" // Leaving blank to test our robust Name/Mobile fallback scanner
  };

  console.log("\n▶️ [TEST 2] Dispatching EDIT payload (Fallback Scan) for:", editPayload.name);
  try {
    const editResult = writeGuestDataRow("EDIT", editPayload);
    console.log("✅ [TEST 2 SUCCESS] Backend returned response:", editResult);
  } catch (error) {
    console.error("❌ [TEST 2 FAILED] Execution crashed with error:", error.message);
  }

  console.log("\n=== 🧪 DEBUG SUITE COMPLETION LOGS END ===");
}


/**
 * Test Harness to safely debug the Airbnb Sync Logic 
 * without modifying real Gmail threads or inbox state.
 */
function runDebugTests() {
  console.log("=== STARTING AIRBNB SYNC ENGINE DEBUG SUITE ===");

  // 1. Setup Mock Headers matching your actual sheet layout
  const mockHeaders = ["Month", "Name", "Guests", "Amount", "Check-in Date", "Days", "AirBnb\\Personal", "Floor", "Customer Ratings", "Comments"];

  // 2. TEST CASE 1: A Raw Cancellation Email (The one causing issues)
  const sampleCancelSubject = "Canceled: Reservation HMNXX8RCKX for Jun 15 – 17, 2026";
  const sampleCancelBody = "Hi Host, Reservation HMNXX8RCKX has been canceled by the guest. These dates are now open.";

  console.log("\n--- Testing Scenario B: Cancellation Parsing ---");
  debugIndividualPayload(sampleCancelSubject, sampleCancelBody, mockHeaders);

  // 3. TEST CASE 2: A Raw Review Email
  const sampleReviewSubject = "Sri Harsha left a 5-star review!";
  const sampleReviewBody = "Read on for a snapshot of what Sri Harsha loved about their stay.";

  console.log("\n--- Testing Scenario A: Review Parsing ---");
  debugIndividualPayload(sampleReviewSubject, sampleReviewBody, mockHeaders);

  console.log("\n=== DEBUG SUITE COMPLETE ===");
}

/**
 * Isolated logic tester to print exactly what your Regex matches
 */
function debugIndividualPayload(subject, body, headers) {
  const combinedTextToAnalyze = (subject + " " + body);
  const targetYear = "2026";

  // --- ISOLATED CANCELLATION TEST ---
  if (combinedTextToAnalyze.toLowerCase().includes("canceled:") || combinedTextToAnalyze.toLowerCase().includes("cancelled:")) {
    console.log("[CHECK] Detected Cancellation Trigger keyword.");

    const cancelMatch = combinedTextToAnalyze.match(/Reservation\s+([A-Z0-9]{10})/i) || combinedTextToAnalyze.match(/code\s+([A-Z0-9]{10})/i);
    const dateMatch = subject.match(/for\s+([A-Z][a-z]{2}\s+\d+)/i);

    let backupCheckInStr = "";
    if (dateMatch) {
      backupCheckInStr = dateMatch[1].trim() + `, ${targetYear}`;
    }

    console.log(`-> Extracted Code: ${cancelMatch ? cancelMatch[1] : "FAILED TO PARSE CODE"}`);
    console.log(`-> Extracted Backup Date: ${backupCheckInStr || "FAILED TO PARSE DATE"}`);
    return;
  }

  // --- ISOLATED REVIEW TEST ---
  if (combinedTextToAnalyze.toLowerCase().includes("left a") && combinedTextToAnalyze.toLowerCase().includes("review")) {
    console.log("[CHECK] Detected Review Trigger keywords.");

    const reviewMatch = combinedTextToAnalyze.match(/([A-Z][a-z]+(?:\s+[A-Z][a-z]+)*)\s+left\s+a\s+(\d+)-star\s+review/i);

    if (reviewMatch) {
      console.log(`-> Extracted Reviewer Name: ${reviewMatch[1]}`);
      console.log(`-> Extracted Rating: ${reviewMatch[2]} Stars`);
    } else {
      console.log("-> FAILED TO PARSE REVIEW REGEX");
    }
    return;
  }
}

/**
 * Test Harness: Executed manually in the editor to isolate and debug
 * calendar event creation, duplicate search rules, and 24-hour reminder triggers.
 */
function debugCalendarSyncWorkflow() {
  console.log("🚀 [DEBUG START] Initializing Calendar Synchronization Test...");

  // 1. Simulate a realistic booking payload bundle
  const mockPayload = {
    guestName: "Test Guest Vinay",
    checkInDate: "2026-06-15", // Simulates a future date execution string
    totalGuests: 3,
    platformType: "AirBnb",
    floorName: "Ground",
    nights: 2,
    notes: "Code: ABC123XYZ9. Automated debug check runner."
  };

  console.log("📋 [MOCK DATA] Payload configuration compiled:", JSON.stringify(mockPayload));

  // 2. Perform validation pre-checks inside the logs
  const parsedDateCheck = new Date(mockPayload.checkInDate);
  console.log(`📅 [DATE PARSING] Raw string '${mockPayload.checkInDate}' translated to Object: ${parsedDateCheck.toString()}`);

  if (isNaN(parsedDateCheck.getTime())) {
    console.error("❌ [DATE ERROR] System failed to resolve time coordinates for the incoming check-in date string.");
    return;
  }

  // 3. Verify target calendar accessibility
  const calendarName = "Vinyasa Nilaya";
  const calendars = CalendarApp.getCalendarsByName(calendarName);
  console.log(`🔍 [CALENDAR ACCESSIBILITY] Searching for calendar named: '${calendarName}'`);

  if (calendars.length === 0) {
    console.warn(`⚠️ [CALENDAR WARNING] No calendar found named '${calendarName}'. The engine will use your primary default Google Account calendar instead.`);
  } else {
    console.log(`✅ [CALENDAR FOUND] Target calendar successfully bound. ID: ${calendars[0].getId()}`);
  }

  // 4. Fire the actual live function execution path
  console.log("⚙️ [EXECUTION] Dispatching payload variables straight to syncBookingToVinyasaCalendar...");

  try {
    const isSuccess = syncBookingToVinyasaCalendar(
      mockPayload.guestName,
      mockPayload.checkInDate,
      mockPayload.totalGuests,
      mockPayload.platformType,
      mockPayload.floorName,
      mockPayload.nights,
      mockPayload.notes
    );

    if (isSuccess) {
      console.log("🎉 [DEBUG SUCCESS] The sync engine executed flawlessly. Check your Google Calendar grid for June 15, 2026!");
    } else {
      console.error("❌ [DEBUG FAILED] Sync returned false. Read the execution logs above to trace structural bottlenecks.");
    }

  } catch (err) {
    console.error("💥 [CRITICAL CRASH] The calendar sync workflow thrown an unhandled exception:", err.toString());
  }

  console.log("🏁 [DEBUG END] Test sequence completed.");
}