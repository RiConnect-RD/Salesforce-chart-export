const express = require('express');
const moment = require('moment');
const app = express();
const multer = require("multer");
const XLSX = require('xlsx')
const ChartJsImage = require('chartjs-to-image');

const initialUserList = [
  {
    name: 'Software \rDev.  Dept.',
    email: 'rd@riconnect.tech',
    count: 0,
  },
  {
    name: 'YOKE',
    email: 'riconnect.marketing@riconnect.tech',
    count: 0,
  },
  {
    name: 'Lay Ana',
    email: 'ana_lay@riconnect.tech',
    count: 0,
  },
  {
    name: 'Lin  Jimmy',
    email: 'jimmy_lin@riconnect.tech',
    count: 0,
  },
  {
    name: 'Lee Hugo',
    email: 'hugo_lee@riconnect.tech',
    count: 0,
  },
  {
    name: 'Deloitte',
    email: 'angeyang@riconnect.tech',
    count: 0,
  },
  {
    name: '青山 龍也',
    email: 'tatsuya_aoyama@riconnect.tech',
    count: 0,
  },
  {
    name: 'Okada Daisaku',
    email: 'daisaku_okada@riconnect.tech',
    count: 0,
  },
  {
    name: 'Tsai Helen',
    email: 'helen_tsai+ri@riconnect.tech',
    count: 0,
  },
  {
    name: 'Lin Jim',
    email: 'jim_lin@riconnect.tech',
    count: 0,
  },
  {
    name: 'Wu Dennis',
    email: 'dennis_wu@riconnect.tech',
    count: 0,
  },
  {
    name: 'Kung Neo',
    email: 'neo_kung@riconnect.tech',
    count: 0,
  },
]

const uploadXLSX = async (req, res, next) => {
  try {
    let workbook = XLSX.readFile(`./uploads/${req.file.filename}`);
    let sheet_name_list = workbook.SheetNames;
    let xlData = XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[0]]);
    
    let start = req.body.start;
    let end = req.body.end;
    let tempUserList = JSON.parse(JSON.stringify(initialUserList));
    xlData.forEach((item) => {
      if (moment(item['Login Time (??????)'], 'YYYY/MM/DD A hh:mm').format('YYYYMMDD') >= moment(start).format('YYYYMMDD') && moment(item['Login Time (??????)'], 'YYYY/MM/DD A hh:mm').format('YYYYMMDD') <= moment(end).format('YYYYMMDD')) {
        tempUserList.forEach((child, index) => {
          if (child.email === item.Username) {
            tempUserList[index].count += 1;
          }
        })
      }
    });
    const myChart = new ChartJsImage();
    let nameList = [], valueList = [];
    tempUserList.forEach((item) => {
      nameList.push(item.name);
      valueList.push(item.count);
    })
  
    myChart.setConfig({
      type: 'bar',
      data: {
        labels: nameList, datasets: [{
          label: '登入次數',
          data: valueList,
          datalabels: {
            anchor: 'end',
            font: {
              size: 32
            }
          }
        }]
      },

      options: {
        legend: {
          position: 'right',
          labels: {
            fontSize: 32
          }
        },
        layout: {
          padding: {
            top: 30,
            left: 30,
            right: 30,
            bottom: 60
          }
        },
        scales: {
          xAxes: [{
            ticks: {
              fontSize: 26,
              color: 'black'
            }
          }],
          yAxes: [{
            ticks: {
              fontSize: 32,
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
    myChart.setHeight(700);
    myChart.setWidth(2600);
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