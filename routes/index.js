var express = require('express');
var XLSX = require('xlsx')
var path = require('path');
const Moment = require('moment');
const MomentRange = require('moment-range');
const moment = MomentRange.extendMoment(Moment);
var router = express.Router();

/* GET home page. */
router.get('/', function (req, res, next) {
  res.render('index', { title: 'Express' });
});

router.post('/submit', validateData, function (req, res) {
  try {
    let jsonData = createJSONData({ start: req.body.start, end: req.body.end, period: req.body.period })
    createFile(jsonData)
    var filePath = path.resolve('./test.xlsx');
    res.download(filePath)
    // res.status(200).json(jsonData)
  } catch (err) {
    console.log(err)
    res.status(400).json(err)
  }
})

var column1 = '1'
var column2 = '2'

function createJSONData(data) {
  let tableData = []
  const start = new Date(data.start);
  const end = new Date(data.end);
  const range = moment.range(start, end);
  const years = Array.from(range.by('month'));
  const annualTableColumn = createAnnualTableColumns(Array.from(range.by('year')))
  const quaterlyTableColumn = createQuaterlyTableColumn(Array.from(range.by('quarter')))
  const monthlyTableColumn = createMonthlyTableColumn(Array.from(range.by('month')))
  const period = data.period

  for (let i = 0; i < years.length; i++) {
    let randomNumber = getRandomNumber()
    let rowData = {}
    rowData[column1] = years[i].format('MM/YYYY')
    rowData[column2] = randomNumber
    let perMonth = randomNumber / period
    let currentDateYear = years[i].format('YYYY')
    let remainingPeriod = period
    let remainingPeriod2 = period
    let remainingPeriod3 = period
    let currentMonth = years[i].format('MM')
    let currentMonth2 = years[i].format('MM')
    let currentQuater = getQuater(years[i])

    for (var j = 0; j < annualTableColumn.length; j++) {
      let columnHeaderYear = parseInt(annualTableColumn[j])
      let remainingMonthOfCurrentColumnHeaderYear = (12 - currentMonth) + 1
      if (columnHeaderYear >= currentDateYear) {
        if (remainingPeriod > 0) {
          if (remainingPeriod < remainingMonthOfCurrentColumnHeaderYear) {
            remainingMonthOfCurrentColumnHeaderYear = remainingPeriod
            remainingPeriod = 0
          } else {
            remainingPeriod -= remainingMonthOfCurrentColumnHeaderYear
          }
          rowData[annualTableColumn[j]] = remainingMonthOfCurrentColumnHeaderYear * perMonth
        } else {
          rowData[annualTableColumn[j]] = 0
        }
      } else {
        rowData[annualTableColumn[j]] = 0
      }
      currentMonth = 1
    }

    let currentDate = moment(years[i])
    for (var k = 0; k < quaterlyTableColumn.length; k++) {
      let columnHeaderYear = parseInt(moment(quaterlyTableColumn[k].date).format('YYYY'))
      let headerQuater = getQuater(quaterlyTableColumn[k].date)
      if (columnHeaderYear >= currentDateYear) {
        if ((columnHeaderYear == currentDateYear) && headerQuater < currentQuater) {
          rowData[quaterlyTableColumn[k].headerString] = 0
          currentDate = moment(quaterlyTableColumn[k].date)
        } else {
          if (remainingPeriod2 > 0) {
            let firstDate = moment(quaterlyTableColumn[k].date)
            if (k == 0) {
              firstDate.endOf('quarter')
            }
            let multiplier = Math.ceil(firstDate.diff(currentDate, 'month', true)) || 1
            if (remainingPeriod2 < multiplier) {
              multiplier = remainingPeriod2
            }
            // rowData[quaterlyTableColumn[k].headerString] = perMonth + ' * ' + multiplier + ' -- ' + moment(quaterlyTableColumn[k].date).toISOString() + ' / ' + currentDate.toISOString()
            rowData[quaterlyTableColumn[k].headerString] =  perMonth * multiplier
            currentDate = moment(quaterlyTableColumn[k].date)
            remainingPeriod2 -= multiplier
          } else {
            rowData[quaterlyTableColumn[k].headerString] = 0
          }
        }
      } else {
        rowData[quaterlyTableColumn[k].headerString] = 0
        currentDate = moment(quaterlyTableColumn[k].date)
      }
    }

    for (l = 0; l < monthlyTableColumn.length; l++) {
      let columnHeaderYear = parseInt(moment(monthlyTableColumn[l].date).format('YYYY'))
      let headerMonth = parseInt(moment(monthlyTableColumn[l].date).format('MM'))
      if (columnHeaderYear >= currentDateYear) {
        if ((columnHeaderYear == currentDateYear) && headerMonth < currentMonth2) {
          rowData[monthlyTableColumn[l].headerString] = 0
        } else {
          if (remainingPeriod3 > 0) {
            rowData[monthlyTableColumn[l].headerString] = perMonth
            remainingPeriod3--;
          } else {
            rowData[monthlyTableColumn[l].headerString] = 0
          }
        }
      } else {
        rowData[monthlyTableColumn[l].headerString] = 0
      }
    }

    tableData.push(rowData)
  }
  return tableData
}

function createMonthlyTableColumn(months) {
  let temp = []
  months.forEach(month => {
    temp.push({ date: month, headerString: `${month.format('MM/YYYY')}` })
  })
  return temp
}

function getQuater(momentDate) {
  return Math.floor((momentDate.toDate().getMonth() + 3) / 3)
}

function createQuaterlyTableColumn(quarters) {
  let temp = []
  quarters.forEach(month => {
    temp.push({ date: month, headerString: `Q${Math.floor((month.toDate().getMonth() + 3) / 3)} ${month.format('YYYY')}` })
  })
  return temp
}

function createAnnualTableColumns(years) {
  let temp = []
  years.forEach(month => {
    temp.push(month.format('YYYY'))
  })
  return temp
}

function getRandomNumber() {
  let min = 1, max = 100000
  return Math.floor(Math.random() * (max - min + 1)) + min
}

function createFile(json) {
  const fileName = 'test.xlsx';
  const ws = XLSX.utils.json_to_sheet(json);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'test');
  XLSX.writeFile(wb, fileName);
}

function readFile() {
  var filePath = path.resolve('./public/images/dates.xlsx');
  var workbook = XLSX.readFile(filePath);
  var sheet_name_list = workbook.SheetNames;
  var xlData = XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[0]]);
  return xlData
}

function validateData(req, res, next) {
  if (req.body.start && req.body.end && req.body.period) {
    try {
      let startDate = new Date(req.body.start)
      let endDate = new Date(req.body.end)
      if (endDate < startDate) {
        res.status(400).json({ error: 'end date should be after start date' })
      } else {
        next()
      }
    } catch (error) {
      res.status(400).json({ error: 'invalid data' })
    }
  } else {
    res.status(400).json({ error: 'invalid data' })
  }
}

module.exports = router;
