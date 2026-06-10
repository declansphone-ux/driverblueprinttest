

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
