// ============================================================
// AUSTRIAN GRAND PRIX — Google Apps Script Backend
// Deploy as Web App: Execute as Me, Anyone can access
// ============================================================

const SHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();

// ── ROUTING ─────────────────────────────────────────────────
function doGet(e) {
  // Support payload-based POST-via-GET for CORS compatibility
  if (e.parameter.payload) {
    try {
      const payload = JSON.parse(decodeURIComponent(e.parameter.payload));
      const action = payload.action || '';
      switch (action) {
        case 'logEntry':    return jsonResponse(logEntry(payload.entry));
        case 'voidEntry':   return jsonResponse(voidEntry(payload.entryId));
        case 'approveFlag': return jsonResponse(approveFlag(payload.entryId));
        case 'setDay':      return jsonResponse(setDay(payload.day, payload.open));
        case 'updatePR':    return jsonResponse(updatePR(payload.driverId, payload.categoryId, payload.value));
        default:            return jsonResponse({ error: 'Unknown payload action: ' + action });
      }
    } catch(err) {
      return jsonResponse({ error: 'Payload parse error: ' + err.toString() });
    }
  }

  const action = e.parameter.action || '';
  try {
    switch (action) {
      case 'getStatus':      return jsonResponse(getStatus());
      case 'getScoreboard':  return jsonResponse(getScoreboard());
      case 'getDriverStats': return jsonResponse(getDriverStats(e.parameter.driverId));
      case 'getPRRecords':   return jsonResponse(getPRRecords(e.parameter.driverId));
      default:               return jsonResponse({ error: 'Unknown GET action: ' + action });
    }
  } catch(err) {
    return jsonResponse({ error: err.toString() });
  }
}

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const action = payload.action || '';
    switch (action) {
      case 'logEntry':     return jsonResponse(logEntry(payload.entry));
      case 'voidEntry':    return jsonResponse(voidEntry(payload.entryId));
      case 'approveFlag':  return jsonResponse(approveFlag(payload.entryId));
      case 'setDay':       return jsonResponse(setDay(payload.day, payload.open));
      case 'updatePR':     return jsonResponse(updatePR(payload.driverId, payload.categoryId, payload.value));
      default:             return jsonResponse({ error: 'Unknown POST action: ' + action });
    }
  } catch(err) {
    return jsonResponse({ error: err.toString() });
  }
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── SHEET HELPERS ────────────────────────────────────────────
function getSheet(name) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
}

function sheetToObjects(sheetName) {
  const sheet = getSheet(sheetName);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  });
}

// ── STATUS ───────────────────────────────────────────────────
function getStatus() {
  const settings = sheetToObjects('Settings');
  const map = {};
  settings.forEach(r => map[r.Key] = r.Value);
  return {
    dayOpen: map['DayOpen'] === 'TRUE' || map['DayOpen'] === true,
    currentDay: parseInt(map['CurrentDay']) || 1,
    competitionName: map['CompetitionName'] || 'Austrian Grand Prix'
  };
}

function setDay(day, open) {
  const sheet = getSheet('Settings');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === 'DayOpen') sheet.getRange(i+1, 2).setValue(open ? 'TRUE' : 'FALSE');
    if (data[i][0] === 'CurrentDay') sheet.getRange(i+1, 2).setValue(day);
  }
  return { success: true, day, open };
}

// ── LOG ENTRY ────────────────────────────────────────────────
function logEntry(entry) {
  const status = getStatus();
  if (!status.dayOpen) return { success: false, error: 'Day is closed' };

  // Get category multiplier
  const cats = sheetToObjects('Categories');
  const cat = cats.find(c => c.ID === entry.categoryId);
  if (!cat) return { success: false, error: 'Unknown category: ' + entry.categoryId };

  const points = parseFloat(entry.qty) * parseFloat(cat.Multiplier);

  // Flag check
  const flagged = checkFlag(entry);

  // Append to Entries sheet
  const sheet = getSheet('Entries');
  const id = 'E' + Date.now() + Math.random().toString(36).slice(2,6).toUpperCase();
  sheet.appendRow([
    id,
    entry.driverId,
    entry.driverName,
    entry.constructor,
    entry.homeRole,
    entry.categoryId,
    entry.categoryName,
    parseInt(entry.qty),
    entry.loggedBracket,
    parseFloat(cat.Multiplier),
    points,
    entry.timestamp,
    status.currentDay,
    flagged ? 'TRUE' : 'FALSE',
    'FALSE', // voided
    '' // notes
  ]);

  return { success: true, entryId: id, points, flagged };
}

function checkFlag(entry) {
  const entries = sheetToObjects('Entries');
  const today = new Date().toDateString();

  // Point spike: single entry > driver's entire prior daily total
  const priorPoints = entries
    .filter(e => e.DriverID === entry.driverId &&
                 new Date(e.Timestamp).toDateString() === today &&
                 e.Voided !== 'TRUE')
    .reduce((sum, e) => sum + parseFloat(e.Points || 0), 0);

  const entryPoints = parseFloat(entry.qty) * parseFloat(entry.mult || 1);
  if (priorPoints > 0 && entryPoints > priorPoints) return true;

  // Duplicate: same driver + category within 60 seconds
  const recent = entries.find(e =>
    e.DriverID === entry.driverId &&
    e.CategoryID === entry.categoryId &&
    Math.abs(new Date(entry.timestamp) - new Date(e.Timestamp)) < 60000 &&
    e.Voided !== 'TRUE'
  );
  if (recent) return true;

  return false;
}

// ── VOID / APPROVE ───────────────────────────────────────────
function voidEntry(entryId) {
  const sheet = getSheet('Entries');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === entryId) {
      sheet.getRange(i+1, 15).setValue('TRUE'); // Voided column
      return { success: true };
    }
  }
  return { success: false, error: 'Entry not found' };
}

function approveFlag(entryId) {
  const sheet = getSheet('Entries');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === entryId) {
      sheet.getRange(i+1, 14).setValue('FALSE'); // Flagged column
      return { success: true };
    }
  }
  return { success: false, error: 'Entry not found' };
}

// ── SCOREBOARD ───────────────────────────────────────────────
function getScoreboard() {
  const entries = sheetToObjects('Entries').filter(e => e.Voided !== 'TRUE');
  const drivers = sheetToObjects('Drivers');
  const constructors = sheetToObjects('Constructors');
  const status = getStatus();

  // Build driver season totals
  const driverTotals = {};
  entries.forEach(e => {
    if (!driverTotals[e.DriverID]) {
      const d = drivers.find(d => d.ID === e.DriverID);
      driverTotals[e.DriverID] = {
        id: e.DriverID, name: e.DriverName, constructor: e.Constructor,
        bracket: d ? d.Bracket : e.HomeRole, pts: 0
      };
    }
    driverTotals[e.DriverID].pts += parseFloat(e.Points || 0);
  });

  // Build constructor totals
  const constructorTotals = {};
  constructors.forEach(c => { constructorTotals[c.Name] = { name: c.Name, pts: 0 }; });
  Object.values(driverTotals).forEach(d => {
    if (constructorTotals[d.constructor]) {
      constructorTotals[d.constructor].pts += d.pts;
    }
  });

  // Sort and group by bracket
  const allDrivers = Object.values(driverTotals);
  const byBracket = { PZ: [], GB: [], Forum: [], Ops: [] };
  allDrivers.forEach(d => {
    if (byBracket[d.bracket]) byBracket[d.bracket].push(d);
  });
  Object.keys(byBracket).forEach(b => byBracket[b].sort((a,b) => b.pts - a.pts));

  // Latest submission
  const sorted = entries.sort((a,b) => new Date(b.Timestamp) - new Date(a.Timestamp));
  const latest = sorted[0] ? {
    driverName: sorted[0].DriverName,
    constructor: sorted[0].Constructor,
    loggedBracket: sorted[0].LoggedBracket,
    categoryName: sorted[0].CategoryName,
    qty: sorted[0].Qty,
    timestamp: sorted[0].Timestamp,
  } : null;

  return {
    constructors: Object.values(constructorTotals).sort((a,b) => b.pts - a.pts),
    drivers: byBracket,
    day: status.currentDay,
    spotlight: latest
  };
}

// ── DRIVER STATS ─────────────────────────────────────────────
function getDriverStats(driverId) {
  const entries = sheetToObjects('Entries')
    .filter(e => e.DriverID === driverId && e.Voided !== 'TRUE');
  const status = getStatus();

  const todayEntries = entries.filter(e => parseInt(e.Day) === status.currentDay);
  const todayPts = todayEntries.reduce((s,e) => s + parseFloat(e.Points||0), 0);
  const seasonPts = entries.reduce((s,e) => s + parseFloat(e.Points||0), 0);

  // Constructor total
  const allEntries = sheetToObjects('Entries').filter(e => e.Voided !== 'TRUE');
  const driver = sheetToObjects('Drivers').find(d => d.ID === driverId);
  const constructorPts = driver ? allEntries
    .filter(e => e.Constructor === driver.Constructor)
    .reduce((s,e) => s + parseFloat(e.Points||0), 0) : 0;

  return { todayPts, seasonPts, constructorPts };
}

// ── PR RECORDS ───────────────────────────────────────────────
function getPRRecords(driverId) {
  const records = sheetToObjects('PRRecords').filter(r => r.DriverID === driverId);
  const map = {};
  records.forEach(r => { map[r.CategoryID] = parseInt(r.AllTimeBest) || 0; });
  return map;
}

function updatePR(driverId, categoryId, value) {
  const sheet = getSheet('PRRecords');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === driverId && data[i][1] === categoryId) {
      if (value > data[i][2]) {
        sheet.getRange(i+1, 3).setValue(value);
        sheet.getRange(i+1, 4).setValue(new Date().toISOString());
      }
      return { success: true };
    }
  }
  // New record
  sheet.appendRow([driverId, categoryId, value, new Date().toISOString(), 'Austrian Grand Prix']);
  return { success: true, newRecord: true };
}

// ── SEED FUNCTION (run once to set up sheet) ─────────────────
function seedAll() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // ── Drivers ──
  let sheet = ss.getSheetByName('Drivers') || ss.insertSheet('Drivers');
  sheet.clearContents();
  const driverHeaders = ['ID','Name','Bracket','Constructor','Active'];
  sheet.getRange(1, 1, 1, driverHeaders.length).setValues([driverHeaders]).setFontWeight('bold');
  const driverData = [
    ['D001','Ronald Joseph','PZ','Ferrari','TRUE'],
    ['D002','Miqdad Khakoo','PZ','Ferrari','TRUE'],
    ['D003','Eddie Blanck','PZ','Ferrari','TRUE'],
    ['D004','Janet Ayala','PZ','Ferrari','TRUE'],
    ['D005','Chad Gidaya Jr','GB','Ferrari','TRUE'],
    ['D006','Valery Perez Vega','GB','Ferrari','TRUE'],
    ['D007','Regina Talain','GB','Ferrari','TRUE'],
    ['D008','Kazuhiro Itoh','Forum','Ferrari','TRUE'],
    ['D009','Rafa Osoria Jr.','Ops','Ferrari','TRUE'],
    ['D010','Crystal Garcia','Ops','Ferrari','TRUE'],
    ['D011','Brian Nera','PZ','McLaren','TRUE'],
    ['D012','Nick Sturz','PZ','McLaren','TRUE'],
    ['D013','Sergio Brito','PZ','McLaren','TRUE'],
    ['D014','Devon Rodriguez','PZ','McLaren','TRUE'],
    ['D015','Ryon Martinez','GB','McLaren','TRUE'],
    ['D016','Syan Haghiri','GB','McLaren','TRUE'],
    ['D017','Josias Ezquivel Escobar','GB','McLaren','TRUE'],
    ['D018','Cybel Gomez','GB','McLaren','TRUE'],
    ['D019','Crystal Salinas','Forum','McLaren','TRUE'],
    ['D020','Felipe Zuniga','Ops','McLaren','TRUE'],
    ['D021','Ryan Leynes','PZ','Red Bull','TRUE'],
    ['D022','Sydnee Rosner','PZ','Red Bull','TRUE'],
    ['D023','Rachel Byrnes','PZ','Red Bull','TRUE'],
    ['D024','David Gonzalez','PZ','Red Bull','TRUE'],
    ['D025','Christian Badilla','GB','Red Bull','TRUE'],
    ['D026','Ivan Jackson','GB','Red Bull','TRUE'],
    ['D027','Gil Sanchez','GB','Red Bull','TRUE'],
    ['D028','Devin Flores','GB','Red Bull','TRUE'],
    ['D029','Mike Ladia','Forum','Red Bull','TRUE'],
    ['D030','America Del Rivero','Ops','Red Bull','TRUE'],
    ['D031','Brayden Bennett','PZ','Mercedes','TRUE'],
    ['D032','Mechan Williams','PZ','Mercedes','TRUE'],
    ['D033','Mathew Navarrete','PZ','Mercedes','TRUE'],
    ['D034','Matthew Bahena','PZ','Mercedes','TRUE'],
    ['D035','Anthony Lopez','GB','Mercedes','TRUE'],
    ['D036','Jesus Barboza','GB','Mercedes','TRUE'],
    ['D037','Andrew Catalan','GB','Mercedes','TRUE'],
    ['D038','Nick Ramos','GB','Mercedes','TRUE'],
    ['D039','Austin Artinger','Forum','Mercedes','TRUE'],
    ['D040','Lesley Vazquez','Ops','Mercedes','TRUE'],
    ['D041','Krystell Barreto','PZ','Aston Martin','TRUE'],
    ['D042','Malik Smith','PZ','Aston Martin','TRUE'],
    ['D043','Jessie Effendi','PZ','Aston Martin','TRUE'],
    ['D044','Jesus Espana Jr.','PZ','Aston Martin','TRUE'],
    ['D045','Leslie Rincon-Lares','GB','Aston Martin','TRUE'],
    ['D046','Dom Sanchez','GB','Aston Martin','TRUE'],
    ['D047','Robby Sherman','GB','Aston Martin','TRUE'],
    ['D048','Saul Reyes','GB','Aston Martin','TRUE'],
    ['D049','Christian Gutierrez','Forum','Aston Martin','TRUE'],
    ['D050','Ryan Cullen','Ops','Aston Martin','TRUE'],
    ['D051','Phillip Cendana','PZ','Alpine','TRUE'],
    ['D052','Lizz Sarinana','PZ','Alpine','TRUE'],
    ['D053','Johnny Vedolla','PZ','Alpine','TRUE'],
    ['D054','Izzy Arias','PZ','Alpine','TRUE'],
    ['D055','Mayra Gonzalez De La Torres','GB','Alpine','TRUE'],
    ['D056','SunDo Kim','GB','Alpine','TRUE'],
    ['D057','Karamoko Kane','GB','Alpine','TRUE'],
    ['D058','Gavin Tilles','GB','Alpine','TRUE'],
    ['D059','Darnell Burns','Forum','Alpine','TRUE'],
    ['D060','Asher Reeves','Ops','Alpine','TRUE'],
    ['D061','Karissa DeVary','PZ','Williams','TRUE'],
    ['D062','Pedro Negron','PZ','Williams','TRUE'],
    ['D063','Pedro Lozano','PZ','Williams','TRUE'],
    ['D064','Michael Monge','PZ','Williams','TRUE'],
    ['D065','David Mendiola','PZ','Williams','TRUE'],
    ['D066','Ammar Khakoo','GB','Williams','TRUE'],
    ['D067','Justin Drummond','GB','Williams','TRUE'],
    ['D068','Shane Howard','GB','Williams','TRUE'],
    ['D069','Andres Rodriguez','GB','Williams','TRUE'],
    ['D070','Manny Partida','GB','Williams','TRUE'],
    ['D071','Josh Umali','GB','Williams','TRUE'],
    ['D072','Marlene Hernandez','Ops','Williams','TRUE'],
    ['D073','Danny Dunne','PZ','Racing Bulls','TRUE'],
    ['D074','Perla Hernandez','PZ','Racing Bulls','TRUE'],
    ['D075','Nofo Keil','PZ','Racing Bulls','TRUE'],
    ['D076','Ernesto East','PZ','Racing Bulls','TRUE'],
    ['D077','Raquel Nevarez','GB','Racing Bulls','TRUE'],
    ['D078','Omar Ruano','GB','Racing Bulls','TRUE'],
    ['D079','Justin Valero','GB','Racing Bulls','TRUE'],
    ['D080','Cedric Corner','GB','Racing Bulls','TRUE'],
    ['D081','Oziel Aldape','GB','Racing Bulls','TRUE'],
    ['D082','Julio Castaneda','GB','Racing Bulls','TRUE'],
    ['D083','Avery Rodriguez','Ops','Racing Bulls','TRUE'],
    ['D084','Karina Castanon','PZ','Haas','TRUE'],
    ['D085','Bryan Rodriguez','PZ','Haas','TRUE'],
    ['D086','Martin Palacios','PZ','Haas','TRUE'],
    ['D087','Arnulfo Romero','PZ','Haas','TRUE'],
    ['D088','Serdar Kacar','GB','Haas','TRUE'],
    ['D089','Andrew Paredes Caceres','GB','Haas','TRUE'],
    ['D090','Richie Rodriguez','GB','Haas','TRUE'],
    ['D091','Vanessa Velazquez','GB','Haas','TRUE'],
    ['D092','Evan Corey','GB','Haas','TRUE'],
    ['D093','Cailin Wimberly','GB','Haas','TRUE'],
    ['D094','Daniel Camacho','GB','Haas','TRUE'],
    ['D095','Elaine Rodriguez','Ops','Haas','TRUE'],
    ['D096','Esmeralda Sonrey','Ops','Haas','TRUE'],
    ['D097','Nataly Sanchez','PZ','Cadillac','TRUE'],
    ['D098','Shane Sullivan','PZ','Cadillac','TRUE'],
    ['D099','Amy Rincon','PZ','Cadillac','TRUE'],
    ['D100','Priscilla Gonzalez','PZ','Cadillac','TRUE'],
    ['D101','Rachelle Scarpin','PZ','Cadillac','TRUE'],
    ['D102','Martin Eduardo Macias','GB','Cadillac','TRUE'],
    ['D103','Kyle Portman','GB','Cadillac','TRUE'],
    ['D104','Julio Corpeno','GB','Cadillac','TRUE'],
    ['D105','Martin Salgado','GB','Cadillac','TRUE'],
    ['D106','Marco Martinez','GB','Cadillac','TRUE'],
    ['D107','Marcus Grande','Ops','Cadillac','TRUE'],
    ['D108','Chelsea Barnum','Ops','Cadillac','TRUE'],
  ];
  sheet.getRange(2, 1, driverData.length, 5).setValues(driverData);
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, 5);

  // ── Constructors ──
  sheet = ss.getSheetByName('Constructors') || ss.insertSheet('Constructors');
  sheet.clearContents();
  sheet.getRange(1,1,1,3).setValues([['Name','PrimaryColor','AccentColor']]).setFontWeight('bold');
  sheet.getRange(2,1,10,3).setValues([
    ['Ferrari','#DC0000','#F7D117'],
    ['McLaren','#FF8000','#47C7FC'],
    ['Red Bull','#1E5BC6','#F7C300'],
    ['Mercedes','#00D2BE','#C0C0C0'],
    ['Aston Martin','#229971','#CEDC00'],
    ['Alpine','#0078FF','#000000'],
    ['Williams','#1868DB','#E40000'],
    ['Racing Bulls','#2647D8','#E10600'],
    ['Haas','#B0B3B8','#DC0000'],
    ['Cadillac','#1A1A1A','#CC0000'],
  ]);
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, 3);

  // ── Categories ──
  sheet = ss.getSheetByName('Categories') || ss.insertSheet('Categories');
  sheet.clearContents();
  sheet.getRange(1,1,1,5).setValues([['ID','Name','Bracket','Multiplier','PRThreshold']]).setFontWeight('bold');
  sheet.getRange(2,1,23,5).setValues([
    ['CAT01','iPhone AppleCare','PZ',3,4],
    ['CAT02','iPad/Watch/Mac AppleCare','PZ',1,2],
    ['CAT03','Biz Intro','PZ',2,3],
    ['CAT04','Service Sign-up Customer','PZ',2,3],
    ['CAT05','Today at Apple Sign-up','PZ',1,3],
    ['CAT06','Completed iPhone or Mac Repair','GB',3,4],
    ['CAT07','GB Appt','GB',0.5,5],
    ['CAT08','AppleCare Conversion','GB',3,3],
    ['CAT09','Conversion','GB',2,2],
    ['CAT10','Biz Intro','GB',2,2],
    ['CAT11','Today at Apple Sign-up','GB',1,3],
    ['CAT12','Today at Apple Session Led','Forum',5,1],
    ['CAT13','Today at Apple Sign-up','Forum',1,3],
    ['CAT14','Today at Apple Conversion','Forum',2,2],
    ['CAT15','Tip Session Delivered','Forum',2,2],
    ['CAT16','Biz Intro','Forum',2,2],
    ['CAT17','1 Hour of Setup Worked','Forum',2,2],
    ['CAT18','APU/IDL Fulfilled','Ops',0.5,5],
    ['CAT19','Run','Ops',0.125,8],
    ['CAT20','SFS Fulfilled','Ops',0.5,4],
    ['CAT21','Pallet Broken Down','Ops',2,2],
    ['CAT22','Biz Intro','Ops',2,2],
    ['CAT23','Today at Apple Sign-up','Ops',1,3],
  ]);
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, 5);

  // ── Entries ──
  sheet = ss.getSheetByName('Entries') || ss.insertSheet('Entries');
  sheet.clearContents();
  const entryHeaders = ['EntryID','DriverID','DriverName','Constructor','HomeRole','CategoryID','CategoryName','Qty','LoggedBracket','Multiplier','Points','Timestamp','Day','Flagged','Voided','Notes'];
  sheet.getRange(1,1,1,entryHeaders.length).setValues([entryHeaders]).setFontWeight('bold');
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, entryHeaders.length);

  // ── Settings ──
  sheet = ss.getSheetByName('Settings') || ss.insertSheet('Settings');
  sheet.clearContents();
  sheet.getRange(1,1,1,2).setValues([['Key','Value']]).setFontWeight('bold');
  sheet.getRange(2,1,4,2).setValues([
    ['DayOpen','FALSE'],
    ['CurrentDay','1'],
    ['CompetitionName','Austrian Grand Prix'],
    ['PINHash','0825'],
  ]);
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, 2);

  // ── PRRecords ── (don't clear — records persist across competitions)
  sheet = ss.getSheetByName('PRRecords');
  if (!sheet) {
    sheet = ss.insertSheet('PRRecords');
    sheet.getRange(1,1,1,5).setValues([['DriverID','CategoryID','AllTimeBest','SetOnDay','CompetitionName']]).setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, 5);
  }

  SpreadsheetApp.getUi().alert('✅ Austrian Grand Prix seeded successfully!\n\n108 drivers · 10 constructors · 23 categories\n\nRemember: Open Day 1 from the Manager Dashboard before the competition starts.');
}
