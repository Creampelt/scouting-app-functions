const functions = require('firebase-functions');
const {google} = require('googleapis');
const sheets = google.sheets('v4');
const drive = google.drive("v3");
const admin = require('firebase-admin');

const SPREADSHEET_ID = '1XL8KvmFeSvJ8VcQ8Gqf6NGLw6W3aaGT3wAFXo6hBTxc';

const credentials = require('./credentials.json');
const SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive'];
const COL_ORDER = [
  'matchNumber', 'startingLevel', 'crossedHabLine', 'cargoShipCargo', 'cargoShipHatch', 'rocketCargoLevel1',
  'rocketCargoLevel2', 'rocketCargoLevel3', 'rocketHatchLevel1', 'rocketHatchLevel2', 'rocketHatchLevel3',
  'playedDefense', 'defenseEffectiveness', 'habClimb', 'climbDuration', 'buddyClimb', 'robotBrokeDown', 'comments'
];

const config = {
  apiKey: "AIzaSyDlSxCXqPgjwBZTJOsstAuTdORwAf6De0E",
  authDomain: "scouting-app-5ec8f.firebaseapp.com",
  databaseURL: "https://scouting-app-5ec8f.firebaseio.com",
  storageBucket: "scouting-app-5ec8f.appspot.com"
};

admin.initializeApp(config);

exports.createDataSheet = functions.https.onRequest((req, res) => {
  let teamNumber = req.query.teamNumber;
  let event = req.query.event;
  let jwtClient = authenticateJwtClient();
  let request = {
    resource: {
      properties: {
        title: `${teamNumber} Scouting Data ${event}`
      }
    },
    auth: jwtClient,
  };
  sheets.spreadsheets.create(request, (err, response) => {
    if (err) {
      console.error(err);
      return;
    }
    let spreadsheetId = response.data.spreadsheetId;
    let spreadsheetUrl = response.data.spreadsheetUrl;
    admin.database().ref(teamNumber + '/' + event + '/spreadsheetId').set(spreadsheetId);
    let permissionsReq = {
      fileId: spreadsheetId,
      resource: {
        role: "commenter",
        type: "anyone"
      },
      auth: jwtClient
    };
    drive.permissions.create(permissionsReq, (err, response) => {
      if (err) {
        console.error(err);
        return;
      }
      res.status(200).send(spreadsheetUrl);
    });
  });
});

exports.copyDataToSheet = functions.database.ref().onWrite((snapshot, context) => {
  let afterData = snapshot.after.val();
  let beforeData = snapshot.before.val();
  let jwtClient = authenticateJwtClient();
  for (let teamNumber in afterData) {
    // Checks each team's scouting data
    if (afterData.hasOwnProperty(teamNumber) && afterData[teamNumber] !== beforeData[teamNumber]) {

      for (let regional in afterData[teamNumber]) {
        // Checks scouting data for each regional and makes sure there is an existing sheet to update
        if (afterData[teamNumber].hasOwnProperty(regional)
          && afterData[teamNumber][regional] !== beforeData[teamNumber][regional]
          && afterData[teamNumber][regional].spreadsheetId !== "") {
          let matchData = afterData[teamNumber][regional].matchData;
          let oldMatchData = beforeData[teamNumber][regional].matchData;
          // Checks scouting data on each team OR creates whole sheet if new sheet was created
          for (let team in matchData) {
            if (matchData.hasOwnProperty(team) && (matchData[team] !== oldMatchData[team] || afterData[teamNumber][regional].spreadsheetId !== beforeData[teamNumber][regional].spreadsheetId)) {
              updateTeamSheet(matchData[team], jwtClient, afterData[teamNumber][regional].spreadsheetId, team);
            }
          }
        }
      }
    }
  }
  return "Function executed!";
});

function authenticateJwtClient() {
  const jwtClient = new google.auth.JWT(
    credentials.client_email,
    null,
    credentials.private_key,
    SCOPES
  );
  jwtClient.authorize((err, tokens) => {
    if (err) {
      console.error(err);
    } else {
      console.log("Successfully connected!");
    }
  });
  return jwtClient;
}

function updateTeamSheet(data, jwtClient, spreadsheetId, teamNumber) {
  getSheetData(jwtClient, spreadsheetId, (sheets) => {
    let existingSheet = false;
    for (let i = 0; i < sheets.length; i++) {
      if (parseInt(sheets[i].properties.title) === parseInt(teamNumber) && !isNaN(parseInt(teamNumber))) {
        let sheetId = sheets[i].properties.sheetId;
        fillSheet(sheetId, jwtClient, spreadsheetId, teamNumber, data);
        existingSheet = true;
        break;
      }
    }
    if (!existingSheet) {
      console.log(`No existing sheet for team ${teamNumber} found. Creating new sheet.`);
      createSheet(jwtClient, spreadsheetId, teamNumber, (sheetId) => {
        console.log(data);
        fillSheet(sheetId, jwtClient, spreadsheetId, teamNumber, data);
      })
    }
  })
}

function fillSheet(sheetId, jwtClient, spreadsheetId, teamNumber, data) {
  pushValues(jwtClient, data, parseInt(teamNumber), spreadsheetId, () => pushFormat(data, jwtClient, sheetId, spreadsheetId));
}

function getSheetData(jwtClient, spreadsheetId, callback) {
  let request = {
    spreadsheetId: spreadsheetId,
    fields: "sheets",
    auth: jwtClient,
  };
  sheets.spreadsheets.get(request, (err, response) => {
    if (err) {
      console.error(err);
      return;
    }

    console.log('Sheets:', response.data.sheets);
    callback(response.data.sheets);
  });
}

function createSheet(jwtClient, spreadsheetId, teamNumber, callback) {
  let request = {
    spreadsheetId: spreadsheetId,
    resource: {
      requests: [{
        addSheet: {
          properties: {
            title: String(teamNumber),
          }
        }
      }]
    },
    auth: jwtClient
  };
  sheets.spreadsheets.batchUpdate(request, (err, response) => {
    if (err) {
      console.error(err);
      return;
    }
    callback(response.data.replies[0].addSheet.properties.sheetId);
  })
}

function pushValues(jwtClient, data, teamNumber, spreadsheetId, callback) {
  let request = {
    spreadsheetId: spreadsheetId,
    resource: {
      data: [
        copyTeamData(data, teamNumber)
      ],
      valueInputOption: "USER_ENTERED"
    },
    auth: jwtClient
  };
  console.log('Values request:', request.resource);
  sheets.spreadsheets.values.batchUpdate(request, (err, response) => {
    if (err) {
      console.error(err);
      return;
    }
    console.log(JSON.stringify(response, null, 2));
    callback();
  });
}

function copyTeamData(data, teamNumber) {
  let tableData = getTableData(data);
  return {
    majorDimension: "ROWS",
    range: `${teamNumber}!A:Z`,
    values: tableData
  };
}

function getTableData(data) {
  let tableData = [COL_ORDER.map(header => formatHeader(header))];
  // Sorts matches in ascending order
  let indexes = Object.keys(data);
  indexes.sort();
  // Adds rows to table
  for (let i = 0; i < indexes.length; i++) {
    let matchNumber = indexes[i];
    let match = flatten(data[matchNumber]);
    match.matchNumber = matchNumber;
    // Adds data to row in correct order
    let row = [];
    for (let i = 0; i < COL_ORDER.length; i++) {
      let header = COL_ORDER[i];
      let cellData = (Object.keys(match).includes(header) ? match[header] : "");
      row.push(cellData);
    }
    tableData.push(row);
  }
  return tableData;
}

function flatten(obj) {
  let flattenedObj = {};
  Object.assign(
    flattenedObj,
    ...function _flatten(o) {
      return [].concat(...Object.keys(o)
        .map(k =>
          typeof o[k] === 'object' ?
            _flatten(o[k]) :
            ({[k]: o[k]})
        )
      );
    } (obj)
  );
  return flattenedObj;
}

function formatHeader(header) {
  return header
    .replace(/([A-Z])/g, ' $1')
    .replace(/([1-3])/g, ' $1')
    .replace(/^./, (str) => str.toUpperCase());
}

function pushFormat(data, jwtClient, sheetId, spreadsheetId) {
  let requests = [
    {
      repeatCell: {
        range: {
          sheetId: sheetId,
          startRowIndex: 0,
          endRowIndex: 1
        },
        cell: {
          userEnteredFormat: {
            horizontalAlignment : "CENTER",
            textFormat: {
              bold: true
            }
          }
        },
        fields: "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)"
      }
    },
    {
      updateSheetProperties: {
        properties: {
          sheetId: sheetId,
          gridProperties: {
            frozenRowCount: 1,
            frozenColumnCount: 1
          }
        },
        fields: "gridProperties(frozenRowCount,frozenColumnCount)"
      }
    },
    {
      autoResizeDimensions: {
        dimensions: {
          sheetId: sheetId,
          dimension: "COLUMNS",
          startIndex: 0,
          endIndex: COL_ORDER.length + 1
        }
      }
    }
  ];
  requests.push(...getCheckboxReq(data, sheetId));

  let request = {
    spreadsheetId: spreadsheetId,
    resource: {
      requests: requests,
    },
    auth: jwtClient
  };
  console.log('Formatting request:', request.resource);
  sheets.spreadsheets.batchUpdate(request, (err, response) => {
    if (err) {
      console.error(err);
      return;
    }
    console.log(JSON.stringify(response, null, 2));
  });
}

function getCheckboxReq(data, sheetId) {
  let requests = [];
  let checkBoxCols = [COL_ORDER.indexOf('crossedHabLine'), COL_ORDER.indexOf('playedDefense'), COL_ORDER.indexOf('robotBrokeDown')];
  for (let i = 0; i < checkBoxCols.length; i++) {
    let colIndex = checkBoxCols[i];
    requests.push({
      repeatCell: {
        cell: {
          dataValidation: {
            condition: {
              type: "BOOLEAN"
            }
          }
        },
        range: {
          sheetId: sheetId,
            startColumnIndex: colIndex,
            endColumnIndex: colIndex + 1,
            startRowIndex: 1,
            endRowIndex: Object.keys(data).length + 1
        },
        fields: "dataValidation"
      }
    });
  }
  return requests;
}