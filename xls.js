const e = require('express');
const xlsx = require('xlsx');

// Read the Excel file
const workbook = xlsx.readFile('./demo.xlsx');

// Get the first worksheet
const worksheet = workbook.Sheets[workbook.SheetNames[0]];

// Convert the worksheet data to a JSON object
const data = xlsx.utils.sheet_to_json(worksheet, {
  header: 13,
  raw: false,
  dateNF: 'yyyy-mm-dd',
});

const template = {
  labelCode: 'xxxx',
  networkStatus: false,
  articleList: [
    {
      articleId: 'xxxx',
      articleName: 'xxxx',
      data: null,
    },
  ],
  updateStatus: 'xxxx',
  battery: 'xxxx',
  signal: 'xxxx',
  type: 'xxxx',
  gateway: {
    id: 1111,
    code: 'xxxx',
    macAddress: 'xxxx3',
    ipAddress: 'xxxx',
    name: 'xxxx',
    version: 'xxxx',
    status: 'xxxx',
    lastConnectionDate: '0000-00-00T00:00:00.000+0900',
    port: 1111,
  },
  firmwareVersion: 1111,
  templateName: ['xxxx'],
  templateType: ['xxxx'],
  lastResponseTime: '0000-00-00T00:00:00.000+0900',
  lastConnectionTime: '0000-00-00T00:00:00.000+0900',
  arrow: null,
  addInfo2: null,
  addInfo3: null,
  addInfo4: null,
  addInfo5: null,
  temperature: 1111,
  requestDate: '0000-00-00T00:00:00.000+0900',
  completedDate: '0000-00-00T00:00:00.000+0900',
  batteryLevel: 'xxxx',
  templateManual: false,
};

// Print the data
const renew = data.map((d) => updateValue(d));
console.log(renew);

function updateValue(t) {
  let s = JSON.parse(JSON.stringify(template));

  s.labelCode = t['LABEL ID'];
  if (t['PRODUCT ID']) {
    s.articleList[0].articleId = t['PRODUCT ID'];
    s.articleList[0].articleName = t['PRODUCT DESCRIPTION'];
  } else {
    s.articleList = [];
  }

  s.gateway.name = t['LINKED GATEWAY'];
  s.battery = t['BATTERY'];
  s.signal = t['SIGNAL STRENGTH'];
  if (t['TEMPLATE']) {
    s.templateType[0] = t['TEMPLATE'];
  } else {
    s.templateType = [];
  }
  if (t['NETWORK'] == 'ONLINE') {
    s.networkStatus = true;
  } else {
    s.networkStatus = false;
  }
  s.updateStatus = t['STATUS'];
  s.lastResponseTime =
    '20' + t['LATEST RESPONSE TIME'].replace(' ', 'T') + '.000+0900';

  return s;
}
