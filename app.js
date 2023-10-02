const express = require('express');
const moment = require('moment');
const app = express();
const multer = require("multer");
const XLSX = require('xlsx')
const ChartJsImage = require('chartjs-to-image');

const initialUserList = [
  {
    name: 'Dev.',
    email: 'rd@riconnect.tech',
    count: 0,
    lastCount: 0,
  },
  {
    name: 'YOKE',
    email: 'riconnect.marketing@gmail.com',
    count: 0,
    lastCount: 0,
  },
  {
    name: 'Ana',
    email: 'ana_lay@riconnect.tech',
    count: 0,
    lastCount: 0,
  },
  {
    name: 'Hugo',
    email: 'hugo_lee@riconnect.tech',
    count: 0,
    lastCount: 0,
  },
  {
    name: 'Deloitte',
    email: 'angeyang@deloitte.com.ri',
    count: 0,
    lastCount: 0,
  },
  {
    name: 'Aoyama',
    email: 'tatsuya_aoyama@riconnect.tech',
    count: 0,
    lastCount: 0,
  },
  {
    name: 'Okada',
    email: 'daisaku_okada@riconnect.tech',
    count: 0,
    lastCount: 0,
  },
  {
    name: 'Helen',
    email: 'helen_tsai+ri@riconnect.tech',
    count: 0,
    lastCount: 0,
  },
  {
    name: 'Jim',
    email: 'jim_lin@riconnect.tech',
    count: 0,
    lastCount: 0,
  },
  {
    name: 'Dennis',
    email: 'dennis_wu@riconnect.tech',
    count: 0,
    lastCount: 0,
  },
]

const uploadXLSX = async (req, res, next) => {
  try {
    let workbook = XLSX.readFile(`./uploads/${req.file.filename}`);
    let sheet_name_list = workbook.SheetNames;
    let xlData = XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[0]]);
    
    let start = req.body.start;
    let end = req.body.end;
    let lastStart = req.body.last_start;
    let lastEnd = req.body.last_end;
    let tempUserList = JSON.parse(JSON.stringify(initialUserList));
    xlData.forEach((item) => {
      if ((moment(item['Login Time (??????)'], 'YYYY/MM/DD A hh:mm').format('YYYYMMDD') >= moment(start).format('YYYYMMDD') && moment(item['Login Time (??????)'], 'YYYY/MM/DD A hh:mm').format('YYYYMMDD') <= moment(end).format('YYYYMMDD')) || (moment(item['Login Time (??????)'], 'YYYY/MM/DD A hh:mm').format('YYYYMMDD') >= moment(lastStart).format('YYYYMMDD') && moment(item['Login Time (??????)'], 'YYYY/MM/DD A hh:mm').format('YYYYMMDD') <= moment(lastEnd).format('YYYYMMDD'))) {
        tempUserList.forEach((child, index) => {
          if (child.email === item.Username) {
            if (moment(item['Login Time (??????)'], 'YYYY/MM/DD A hh:mm').format('YYYYMMDD') >= moment(start).format('YYYYMMDD') && moment(item['Login Time (??????)'], 'YYYY/MM/DD A hh:mm').format('YYYYMMDD') <= moment(end).format('YYYYMMDD')) {
              tempUserList[index].count += 1;
            } else {
              tempUserList[index].lastCount += 1;
            }
          }
        })
      }
    });
    const myChart = new ChartJsImage();
    let nameList = [], valueList = [], lastValueList = [];
    tempUserList.forEach((item) => {
      nameList.push(item.name);
      valueList.push(item.count);
      lastValueList.push(item.lastCount);
    })
  
    myChart.setConfig({
      type: 'bar',
      data: {
        labels: nameList, 
        datasets: [
          {
            label: `${moment(lastStart).format('M')}月`,
            data: lastValueList,
            datalabels: {
              anchor: 'end',
              font: {
                size: 50
              }
            }
          },
          {
            label: `${moment(start).format('M')}月`,
            data: valueList,
            datalabels: {
              anchor: 'end',
              font: {
                size: 50
              }
            }
          },
        ]
      },

      options: {
        legend: {
          position: 'right',
          labels: {
            fontSize: 50,
            padding: 60
          },
        },
        layout: {
          padding: {
            top: 100,
            left: 30,
            right: 80,
            bottom: 60
          }
        },
        scales: {
          xAxes: [{
            ticks: {
              fontSize: 50,
              color: 'black'
            }
          }],
          yAxes: [{
            ticks: {
              fontSize: 50,
              color: 'black'
            }
          }]
        },
        plugins: {
        datalabels: {
            align: 'end',
            borderRadius: 4,
            color: 'black',
          }
        },
      }
    });
    myChart.setHeight(1200);
    myChart.setWidth(3000);
    myChart.toFile('./uploads/myChart.png');

    return res.status(201).json({
      success: true,
      data: tempUserList,
    });
  } catch (err) {
    return res.status(500).json({ success: false, message: err.message });
  }
};

var storage = multer.diskStorage({
  destination: function (req, file, cb) {
    cb(null, "uploads/");
  },
  filename: function (req, file, cb) {
    cb(null, Date.now() + "-" + file.originalname);
  },
});

const upload = multer({ storage: storage });

app.post("/upload", upload.single("file"), uploadXLSX);
  
  // 監聽 port
app.listen(3000, function() {
  console.log("App listening at 3000 Port",) 
})