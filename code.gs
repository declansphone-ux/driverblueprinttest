

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('DPD Driver Earnings')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function testBQConnection() {
  const projects = BigQuery.Projects.list();
  Logger.log(projects);
}

function getDriverRoutesData(fdNumber) {
  const projectId = 'dpduk-d-process-tbls-anlst-l1';

  const query = `
    SELECT
      route_date, day_of_week, route_no, franchise_code, driver_code,
      number_of_stops, weekdisplay, business_start_date,
      blueprint_stops_raw, blueprint_stops, blueprint_weekly_avg_stops
    FROM \`dpduk-d-process-tbls-anlst-l1.odf_health.driver_routes_and_blueprint_stops_flat\`
    WHERE franchise_code = '${fdNumber}'
  `;

  const job = BigQuery.Jobs.insert({
    configuration: {
      query: {
        query: query,
        useLegacySql: false,
        location: 'europe-west2'
      }
    }
  }, projectId);

  const jobId = job.jobReference.jobId;

  // Poll until done
  let status;
  do {
    Utilities.sleep(1000);
    status = BigQuery.Jobs.get(projectId, jobId, { location: 'europe-west2' });
  } while (status.status.state !== 'DONE');

  if (status.status.errorResult) {
    throw new Error(status.status.errorResult.message);
  }

  const result = BigQuery.Jobs.getQueryResults(projectId, jobId, { location: 'europe-west2' });

  if (!result.rows) return [];

  const headers = result.schema.fields.map(f => f.name);
  return result.rows.map(row =>
    Object.fromEntries(headers.map((h, i) => [h, row.f[i].v]))
  );
}

function testGetDriverRoutes() {
  const results = getDriverRoutesData('FD0844');
  Logger.log(results);
}

function ODFEarningsBQRefresh() {
  try {
    console.log('Starting BigQuery data refresh...');
    
    const ss = SpreadsheetApp.openById('1dgmin_xjL2WqyAoYmqsddJryNtABkNcAl4Yy8y3cAm8');
    const runSheet = ss.getSheetByName('Run');
    const dataSheet = ss.getSheetByName('Data');
    
    const projectId = 'dpduk-d-process-tbls-anlst-l1';
    const query = `SELECT * FROM \`dpduk-d-process-tbls-anlst-l1.odf_health.earnings_view\``;
    
    const request = {
      query: query,
      useLegacySql: false,
      useQueryCache: false,
      timeoutMs: 30000
    };
    
    const queryJob = BigQuery.Jobs.query(request, projectId);
    const jobId = queryJob.jobReference.jobId;
    const location = queryJob.jobReference.location || 'US';
    
    console.log('Query job initiated, total rows: ' + queryJob.totalRows);
    
    // Fetch all pages
    let allRows = queryJob.rows || [];
    let pageToken = queryJob.pageToken;
    
    while (pageToken) {
      const results = BigQuery.Jobs.getQueryResults(projectId, jobId, {
        pageToken: pageToken,
        maxResults: 10000,
        location: location
      });
      
      if (results.rows) {
        allRows = allRows.concat(results.rows);
      }
      pageToken = results.pageToken;
    }
    
    console.log('Retrieved ' + allRows.length + ' rows');
    
    // Clear Data sheet
    dataSheet.clear();
    
    // Write headers
    const headers = queryJob.schema.fields.map(field => field.name);
    dataSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // Write data in batches
    const batchSize = 1000;
    for (let i = 0; i < allRows.length; i += batchSize) {
      const batch = allRows.slice(i, i + batchSize);
      const values = batch.map(row => row.f.map(cell => cell.v));
      dataSheet.getRange(i + 2, 1, values.length, values[0].length).setValues(values);
      console.log('Wrote batch ' + Math.floor(i / batchSize + 1));
    }
    
    // Update last run timestamp in Run sheet
    runSheet.getRange('A1').setValue(new Date());
    
    console.log('Data refresh complete at ' + new Date());
    
    return {
      success: true,
      rowsLoaded: allRows.length,
      timestamp: new Date()
    };
    
  } catch (error) {
    console.error('Error in ODFEarningsBQRefresh: ' + error.toString());
    return {
      success: false,
      error: error.toString()
    };
  }
}

function setupDailyTrigger() {
  // Check if trigger already exists
  const triggers = ScriptApp.getProjectTriggers();
  let triggerExists = false;
  
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'ODFEarningsBQRefresh') {
      triggerExists = true;
      console.log('Trigger already exists');
      break;
    }
  }
  
  if (!triggerExists) {
    ScriptApp.newTrigger('ODFEarningsBQRefresh')
      .timeBased()
      .atHour(1)
      .everyDays(1)
      .create();
    
    console.log('Daily trigger created for 1-2am');
  }
}

function getDriverData() {
  try {
    const userEmail = Session.getActiveUser().getEmail().toLowerCase();
    console.log('User email: ' + userEmail);
    
    const ss = SpreadsheetApp.openById('1dgmin_xjL2WqyAoYmqsddJryNtABkNcAl4Yy8y3cAm8');
    const dataSheet = ss.getSheetByName('Data');
    
    const allData = dataSheet.getDataRange().getValues();
    const headers = allData[0];
    const rlsIndex = headers.indexOf('rls');
    
    if (rlsIndex === -1) {
      throw new Error('RLS column not found');
    }
    
    // Filter rows by RLS
    const filteredRows = [];
    for (let i = 1; i < allData.length; i++) {
      const rlsValue = allData[i][rlsIndex];
      if (rlsValue && rlsValue.toString().toLowerCase().includes(userEmail)) {
        // Remove RLS column from the row
        const row = allData[i].slice();
        row.splice(rlsIndex, 1);
        filteredRows.push(row);
      }
    }
    
    // Remove RLS from headers
    const filteredHeaders = headers.slice();
    filteredHeaders.splice(rlsIndex, 1);
    
    console.log('Returning ' + filteredRows.length + ' rows');
    
    return {
      success: true,
      schema: filteredHeaders.map(h => ({ name: h })),
      rows: filteredRows.map(row => ({ f: row.map(v => ({ v: v })) })),
      totalRows: filteredRows.length
    };
    
  } catch (error) {
    console.error('Error: ' + error.toString());
    return {
      success: false,
      error: error.toString()
    };
  }
}

function getNewStarterData() {
  try {
    const userEmail = Session.getActiveUser().getEmail().toLowerCase();
    console.log('User email: ' + userEmail);
    
    const ss = SpreadsheetApp.openById('1J7mkfIHRaMBMq19DIgOOlNjmhid6mYRGBuynIhN-7Gw');
    const dataSheet = ss.getSheetByName('data');
    
    const allData = dataSheet.getDataRange().getValues();
    const headers = allData[0];
    const rlsIndex = headers.indexOf('rls');
    
    if (rlsIndex === -1) {
      throw new Error('RLS column not found');
    }
    
    const filteredRows = [];
    for (let i = 1; i < allData.length; i++) {
      const rlsValue = allData[i][rlsIndex];
      if (rlsValue) {
        const emails = rlsValue.toString().toLowerCase().split(',').map(e => e.trim());
        if (emails.includes(userEmail)) {
          const row = allData[i].slice();
          row.splice(rlsIndex, 1);
          filteredRows.push(row);
        }
      }
    }
    
    const filteredHeaders = headers.slice();
    filteredHeaders.splice(rlsIndex, 1);
    
    console.log('Returning ' + filteredRows.length + ' rows');
    
    return {
      success: true,
      schema: filteredHeaders.map(h => ({ name: h })),
      rows: filteredRows.map(row => ({
  f: row.map(v => ({
    v: v instanceof Date ? Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd') : v
  }))
})),
      totalRows: filteredRows.length
    };
    
  } catch (error) {
    console.error('Error: ' + error.toString());
    return {
      success: false,
      error: error.toString()
    };
  }
}


// ===== MONDAY MORNING NEW STARTERS EMAIL =====

function sendMondayNewStarterEmail() {
  try {
    var recipient = 'declan.auston@dpdgroup.co.uk';
    var datastudioUrl = 'https://script.google.com/a/macros/dpdgroup.co.uk/s/AKfycbxPEE08BSS4wD0q5hCJGBz0E5upTLRp3DF7VHMkWaw/dev';

    var data = getAllNewStarterDataUnfiltered();
    if (!data || data.length === 0) {
      console.log('No NS data found, aborting email send');
      return;
    }

    data.forEach(function(row) { row._bpAvg = computeNsBlueprintGS(row); });

    var filtered = filterLast10Weeks(data);
    console.log('Filtered to ' + filtered.length + ' drivers in last 10 completed weeks');

    var netPcts = computeGroupPctsGS(filtered);

    var regionMap = {};
    filtered.forEach(function(d) {
      var r = d.depot_region || 'Unknown';
      if (!regionMap[r]) regionMap[r] = [];
      regionMap[r].push(d);
    });
    var regionEntries = Object.keys(regionMap).map(function(region) {
      return { region: region, drivers: regionMap[region], pcts: computeGroupPctsGS(regionMap[region]) };
    }).sort(function(a, b) { return a.region.localeCompare(b.region); });

    // Build depot rows grouped by region, sorted worst->best on Week 5 within each region
    var sortedDepotRows = [];
    regionEntries.forEach(function(entry) {
      var depotMap = {};
      entry.drivers.forEach(function(d) {
        var dep = d.depot_name || 'Unknown';
        if (!depotMap[dep]) depotMap[dep] = [];
        depotMap[dep].push(d);
      });
      var regionDepots = Object.keys(depotMap).map(function(depot) {
        return { depot: depot, region: entry.region, drivers: depotMap[depot], pcts: computeGroupPctsGS(depotMap[depot]) };
      });
      regionDepots.sort(function(a, b) {
        var av = a.pcts[4] !== null ? a.pcts[4] : -1;
        var bv = b.pcts[4] !== null ? b.pcts[4] : -1;
        return av - bv; // worst first
      });
      regionDepots.forEach(function(d) { sortedDepotRows.push(d); });
    });

    var today = new Date();
    var dateStr = formatNsEmailDate(today);
    var subject = 'New Starters vs Blueprint - Monday ' + dateStr;

    var emailHtml = buildNsEmailHtml(filtered.length, netPcts, regionEntries, datastudioUrl, dateStr);
    var pdfBlob = buildDepotPdf(sortedDepotRows, dateStr);

    MailApp.sendEmail({
      to: recipient,
      subject: subject,
      htmlBody: emailHtml,
      attachments: [pdfBlob]
    });

    console.log('Monday NS email sent successfully to ' + recipient);
  } catch (e) {
    console.error('Error in sendMondayNewStarterEmail: ' + e.toString());
    throw e;
  }
}

function setupMondayEmailTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'sendMondayNewStarterEmail') {
      console.log('Monday email trigger already exists');
      return;
    }
  }
  ScriptApp.newTrigger('sendMondayNewStarterEmail')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(7)
    .create();
  console.log('Monday 7-8am email trigger created');
}

function getAllNewStarterDataUnfiltered() {
  var ss = SpreadsheetApp.openById('1J7mkfIHRaMBMq19DIgOOlNjmhid6mYRGBuynIhN-7Gw');
  var dataSheet = ss.getSheetByName('data');
  var allData = dataSheet.getDataRange().getValues();
  var headers = allData[0];
  var result = [];
  for (var i = 1; i < allData.length; i++) {
    var row = allData[i];
    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      if (headers[j] === 'rls') continue;
      var v = row[j];
      obj[headers[j]] = v instanceof Date ? Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd') : v;
    }
    result.push(obj);
  }
  return result;
}

function filterLast10Weeks(data) {
  var today = new Date();
  var dayOfWeek = today.getDay(); // 0=Sun, 1=Mon
  var daysSinceMonday = (dayOfWeek + 6) % 7;
  var thisMonday = new Date(today);
  thisMonday.setHours(0, 0, 0, 0);
  thisMonday.setDate(today.getDate() - daysSinceMonday);
  var cutoff = new Date(thisMonday);
  cutoff.setDate(cutoff.getDate() - 70); // 10 weeks back
  return data.filter(function(d) {
    var s = d.earliest_date || d.business_start_date || '';
    if (!s) return false;
    var dt = new Date(s);
    return !isNaN(dt.getTime()) && dt >= cutoff;
  });
}

function computeNsBlueprintGS(row) {
  var dayKeys = ['mon', 'tue', 'wed', 'thu', 'fri', 'sat', 'sun'];
  var contracted = dayKeys.filter(function(d) { return String(row['contracted_' + d]).toLowerCase() === 'true'; });
  if (contracted.length > 0) {
    var cStops = contracted.map(function(d) { return parseFloat(row['number_of_stops_' + d]); }).filter(function(v) { return !isNaN(v) && v > 0; });
    if (cStops.length > 0) return cStops.reduce(function(s, v) { return s + v; }, 0) / cStops.length;
  }
  var allStops = dayKeys.map(function(d) { return parseFloat(row['number_of_stops_' + d]); }).filter(function(v) { return !isNaN(v) && v > 0; });
  return allStops.length > 0 ? allStops.reduce(function(s, v) { return s + v; }, 0) / allStops.length : 0;
}

function computeDriverWeekPctGS(row, week) {
  var bp = row._bpAvg;
  if (!bp || bp <= 0) return null;
  var daysWorked = parseInt(row['week_' + week + '_days_worked']) || 0;
  if (daysWorked === 0) return null;
  var val = parseFloat(row['week_' + week + '_avg_stops']) || 0;
  if (val <= 0) return null;
  return (val / bp) * 100;
}

function computeGroupPctsGS(drivers) {
  var weeks = [];
  for (var w = 1; w <= 5; w++) {
    var pcts = drivers.map(function(d) { return computeDriverWeekPctGS(d, w); }).filter(function(p) { return p !== null; });
    weeks.push(pcts.length > 0 ? pcts.reduce(function(s, v) { return s + v; }, 0) / pcts.length : null);
  }
  return weeks;
}

function formatNsEmailDate(date) {
  var months = ['January','February','March','April','May','June','July','August','September','October','November','December'];
  var d = date.getDate();
  var suffix = (d === 1 || d === 21 || d === 31) ? 'st' : (d === 2 || d === 22) ? 'nd' : (d === 3 || d === 23) ? 'rd' : 'th';
  return d + suffix + ' of ' + months[date.getMonth()] + ' ' + date.getFullYear();
}

function getNsColorHexGS(pct, week) {
  if (pct === null || pct === undefined) return { bg: '#f5f5f5', fg: '#999999' };
  var targets = [40, 60, 80, 100, 100];
  var target = targets[week - 1] || 100;
  if (pct >= target)           return { bg: '#d4edda', fg: '#155724' };
  if (pct >= target - 15)      return { bg: '#fff3cd', fg: '#856404' };
  return                                { bg: '#f8d7da', fg: '#721c24' };
}

function buildNsEmailHtml(totalDrivers, netPcts, regionEntries, datastudioUrl, dateStr) {
  var weekTargets = [40, 60, 80, 100, 100];

  function pctTd(pct, week) {
    var c = getNsColorHexGS(pct, week);
    var text = pct !== null ? Math.round(pct) + '%' : '&mdash;';
    var fw = pct !== null ? '700' : '400';
    return '<td style="padding:9px 12px;text-align:center;background:' + c.bg + ';color:' + c.fg + ';font-weight:' + fw + ';border:1px solid #e0e0e0">' + text + '</td>';
  }

  var headerRow = '<th style="padding:10px 12px;text-align:left;background:#DC0032;color:white;border:1px solid #b0001f;min-width:180px">Region</th>';
  headerRow += '<th style="padding:10px 12px;text-align:center;background:#DC0032;color:white;border:1px solid #b0001f;min-width:70px">Drivers</th>';
  for (var w = 1; w <= 5; w++) {
    headerRow += '<th style="padding:10px 12px;text-align:center;background:#DC0032;color:white;border:1px solid #b0001f;min-width:95px">Week ' + w + '<br><span style="font-size:10px;font-weight:400;opacity:0.85">Target: ' + weekTargets[w - 1] + '%</span></th>';
  }

  var networkRow = '<tr style="background:#f0f0f0"><td style="padding:9px 12px;font-weight:700;color:#DC0032;border:1px solid #e0e0e0">&#127760; Network</td>';
  networkRow += '<td style="padding:9px 12px;text-align:center;font-weight:700;border:1px solid #e0e0e0">' + totalDrivers + '</td>';
  netPcts.forEach(function(pct, i) { networkRow += pctTd(pct, i + 1); });
  networkRow += '</tr>';

  var regionRows = '';
  regionEntries.forEach(function(entry) {
    regionRows += '<tr style="background:#fafafa"><td style="padding:9px 12px 9px 24px;font-weight:600;border:1px solid #e0e0e0">' + entry.region + '</td>';
    regionRows += '<td style="padding:9px 12px;text-align:center;border:1px solid #e0e0e0">' + entry.drivers.length + '</td>';
    entry.pcts.forEach(function(pct, i) { regionRows += pctTd(pct, i + 1); });
    regionRows += '</tr>';
  });

  return '<!DOCTYPE html><html><head><meta charset="UTF-8"></head>'
    + '<body style="font-family:Arial,sans-serif;color:#333333;background:#f4f4f4;margin:0;padding:20px">'
    + '<div style="max-width:900px;margin:0 auto;background:white;border-radius:8px;overflow:hidden;box-shadow:0 2px 8px rgba(0,0,0,0.12)">'
    + '<div style="background:#DC0032;padding:24px 28px">'
    + '<div style="font-size:22px;font-weight:800;color:white;letter-spacing:-0.5px">DPD | New Starters vs Blueprint</div>'
    + '<div style="margin-top:6px;color:rgba(255,255,255,0.85);font-size:14px">Monday ' + dateStr + '</div>'
    + '</div>'
    + '<div style="background:#fff8e1;border-left:4px solid #ffc107;padding:11px 20px">'
    + '<span style="font-size:12px;color:#555">Data shown for drivers whose first date falls within the <strong>last 10 completed weeks</strong> &mdash; ' + totalDrivers + ' drivers included.</span>'
    + '</div>'
    + '<div style="padding:24px 28px">'
    + '<table style="width:100%;border-collapse:collapse;font-size:13px">'
    + '<thead><tr>' + headerRow + '</tr></thead>'
    + '<tbody>' + networkRow + regionRows + '</tbody>'
    + '</table>'
    + '</div>'
    + '<div style="padding:0 28px 20px 28px">'
    + '<p style="margin:0;font-size:12px;color:#888;font-style:italic">&#128206; See attached PDF for a full depot breakdown by region, ordered worst to best on Week 5.</p>'
    + '</div>'
    + '<div style="background:#f8f8f8;border-top:1px solid #e0e0e0;padding:20px 28px;text-align:center">'
    + '<a href="' + datastudioUrl + '" style="display:inline-block;background:#DC0032;color:white;text-decoration:none;padding:12px 32px;border-radius:6px;font-size:14px;font-weight:700">View DataStudio</a>'
    + '<p style="margin:14px 0 0 0;font-size:11px;color:#aaa">Automated Monday morning report &mdash; DPD Driver Blueprint System</p>'
    + '</div>'
    + '</div></body></html>';
}

function buildDepotPdf(depotRows, dateStr) {
  var ss = SpreadsheetApp.create('NS_Blueprint_Depot_Temp_' + Date.now());
  var sheet = ss.getActiveSheet();
  sheet.setName('Depot Summary');

  var weekTargets = [40, 60, 80, 100, 100];
  var COL_COUNT = 8;

  // Row 1: Title
  sheet.getRange(1, 1, 1, COL_COUNT).merge()
    .setValue('New Starters vs Blueprint — Depot Breakdown — Monday ' + dateStr)
    .setFontWeight('bold').setFontSize(12)
    .setBackground('#DC0032').setFontColor('white')
    .setHorizontalAlignment('center');
  sheet.setRowHeight(1, 30);

  // Row 2: Data note
  sheet.getRange(2, 1, 1, COL_COUNT).merge()
    .setValue('Data shown for drivers whose first date falls within the last 10 completed weeks. Depots ordered by region, then worst to best on Week 5.')
    .setFontStyle('italic').setFontSize(9)
    .setBackground('#fff8e1').setFontColor('#555555')
    .setHorizontalAlignment('left').setWrap(true);
  sheet.setRowHeight(2, 28);

  // Row 3: Column headers
  var headers = ['Region', 'Depot', 'Drivers', 'Week 1 (Target: 40%)', 'Week 2 (Target: 60%)', 'Week 3 (Target: 80%)', 'Week 4 (Target: 100%)', 'Week 5 (Target: 100%)'];
  sheet.getRange(3, 1, 1, COL_COUNT).setValues([headers])
    .setBackground('#333333').setFontColor('white')
    .setFontWeight('bold').setFontSize(10);
  sheet.setRowHeight(3, 26);

  if (depotRows.length > 0) {
    var dataValues = depotRows.map(function(r) {
      return [
        r.region, r.depot, r.drivers.length,
        r.pcts[0] !== null ? Math.round(r.pcts[0]) + '%' : '—',
        r.pcts[1] !== null ? Math.round(r.pcts[1]) + '%' : '—',
        r.pcts[2] !== null ? Math.round(r.pcts[2]) + '%' : '—',
        r.pcts[3] !== null ? Math.round(r.pcts[3]) + '%' : '—',
        r.pcts[4] !== null ? Math.round(r.pcts[4]) + '%' : '—'
      ];
    });
    sheet.getRange(4, 1, dataValues.length, COL_COUNT).setValues(dataValues);

    depotRows.forEach(function(row, rowIdx) {
      var sheetRow = 4 + rowIdx;
      if (rowIdx % 2 === 1) {
        sheet.getRange(sheetRow, 1, 1, 3).setBackground('#f5f5f5');
      }
      row.pcts.forEach(function(pct, weekIdx) {
        var cell = sheet.getRange(sheetRow, 4 + weekIdx);
        if (pct === null) { cell.setBackground('#f5f5f5').setFontColor('#999999'); return; }
        var target = weekTargets[weekIdx];
        if (pct >= target)           { cell.setBackground('#d4edda').setFontColor('#155724').setFontWeight('bold'); }
        else if (pct >= target - 15) { cell.setBackground('#fff3cd').setFontColor('#856404').setFontWeight('bold'); }
        else                         { cell.setBackground('#f8d7da').setFontColor('#721c24').setFontWeight('bold'); }
      });
    });

    sheet.getRange(3, 1, 1 + dataValues.length, COL_COUNT)
      .setBorder(true, true, true, true, true, true, '#cccccc', SpreadsheetApp.BorderStyle.SOLID);
  }

  sheet.setColumnWidth(1, 130);
  sheet.setColumnWidth(2, 160);
  sheet.setColumnWidth(3, 60);
  for (var c = 4; c <= 8; c++) sheet.setColumnWidth(c, 135);
  sheet.setFrozenRows(3);

  Utilities.sleep(1500);

  var fileId = ss.getId();
  var exportUrl = 'https://docs.google.com/spreadsheets/d/' + fileId + '/export'
    + '?format=pdf&size=A4&landscape=true&fitw=true'
    + '&sheetnames=false&printtitle=false&pagenumbers=false&gridlines=false';
  var token = ScriptApp.getOAuthToken();
  var response = UrlFetchApp.fetch(exportUrl, {
    headers: { Authorization: 'Bearer ' + token },
    muteHttpExceptions: true
  });
  var safeDateStr = dateStr.replace(/[^a-zA-Z0-9]/g, '_');
  var pdfBlob = response.getBlob().setName('NS_Blueprint_Depots_' + safeDateStr + '.pdf');

  DriveApp.getFileById(fileId).setTrashed(true);

  return pdfBlob;
}
