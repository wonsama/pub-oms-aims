var express = require('express');
var router = express.Router();

const xlsx = require('xlsx');

// Read the Excel file
const workbook = xlsx.readFile('demo.xlsx');

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

// Print the data
const json = data.map((d) => updateValue(d));

const result = {
  responseMessage: 'The request has been completed',
  responseCode: '200',
  customBatchId: null,
  labelList: json,
};

/* GET users listing. */
router.get('/', function (req, res, next) {
  res.send(result);
});

module.exports = router;

/*
const json = {
  responseMessage: 'The request has been completed',
  responseCode: '200',
  customBatchId: null,
  labelList: [
    {
      labelCode: '07D6E1D4B199',
      networkStatus: true,
      articleList: [
        {
          articleId: 'C129060',
          articleName: '클리오 샤프쏘심플브로우펜슬 1/2/3',
          data: null,
        },
      ],
      updateStatus: 'SUCCESS',
      battery: 'GOOD',
      signal: 'EXCELLENT',
      type: 'NEWTON_GRAPHIC_2_2_RED_NFC',
      gateway: {
        id: 3500126,
        code: 'D02544FFFE20D763',
        macAddress: 'D0-25-44-20-D7-63',
        ipAddress: '10.176.247.110',
        name: 'GW_D02544FFFE20D763',
        version: 'N1.3.7.0',
        status: 'CONNECTED',
        lastConnectionDate: '2024-05-07T14:38:52.328+0900',
        port: 80,
      },
      firmwareVersion: 23,
      templateName: ['OY_2_2_교차_240116.xsl'],
      templateType: ['OY_2_2_교차'],
      lastResponseTime: '2024-05-07T10:01:59.000+0900',
      lastConnectionTime: '2024-05-01T04:04:16.203+0900',
      arrow: null,
      addInfo2: null,
      addInfo3: null,
      addInfo4: null,
      addInfo5: null,
      temperature: 26,
      requestDate: '2024-05-01T04:04:16.025+0900',
      completedDate: '2024-05-01T04:20:59.000+0900',
      batteryLevel: 'FF',
      templateManual: true,
    },
  ],
};
*/
