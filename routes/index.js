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
  const start = new Date(data.start)
  const end = new Date(data.end)
  const range = moment.range(start, end)
  const years = Array.from(range.by('month'))
  const annualTableColumn = createAnnualTableColumns(range)
  const quaterlyTableColumn = createQuaterlyTableColumn(Array.from(range.by('quarter')), range)
  const monthlyTableColumn = createMonthlyTableColumn(Array.from(range.by('month')))
  const period = data.period
  const mul2 = [2, 5, 8, 11]
  const mul1 = [3, 6, 9, 12]

  for (let i = 0; i < years.length; i++) {
    let randomNumber = getRandomNumber()
    let rowData = {}
    rowData[column1] = years[i].format('MMM-YYYY')
    rowData[column2] = randomNumber
    let perMonth = randomNumber / period
    let currentDateYear = years[i].format('YYYY')
    let remainingPeriod = period
    let remainingPeriod2 = period
    let remainingPeriod3 = period
    let currentMonth = years[i].format('MM')
    let currentMonth2 = years[i].format('MM')
    let currentQuater = getQuater(years[i])

    // ************* annual table row loop **************************
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
          if (i > (years.length - period)) {
            remainingMonthOfCurrentColumnHeaderYear -= (i - (years.length - period))
          }
          rowData[annualTableColumn[j]] = Math.floor(remainingMonthOfCurrentColumnHeaderYear * perMonth)
        } else {
          rowData[annualTableColumn[j]] = 0
        }
      } else {
        rowData[annualTableColumn[j]] = 0
      }
      currentMonth = 1
    }

    // ************* quaterly table row loop **************************
    let currentDate = moment(years[i])
    let count = 1
    for (var k = 0; k < quaterlyTableColumn.length; k++) {
      let columnHeaderYear = parseInt(moment(quaterlyTableColumn[k].date).format('YYYY'))
      let headerQuater = getQuater(quaterlyTableColumn[k].date)
      if (columnHeaderYear >= currentDateYear) {
        if ((columnHeaderYear == currentDateYear) && headerQuater < currentQuater) {
          rowData[quaterlyTableColumn[k].headerString] = 0
          currentDate = moment(quaterlyTableColumn[k].date)
        } else {
          if (remainingPeriod2 > 0) {
            let multiplier = 3
            if (k == 0 || count == 1) {
              if (mul2.indexOf(parseInt(years[i].format('M'))) != -1) {
                multiplier = 2
              } else if (mul1.indexOf(parseInt(years[i].format('M'))) != -1) {
                multiplier = 1
              }
            }
            if (remainingPeriod2 < multiplier) {
              multiplier = remainingPeriod2
            }
            if (i > (years.length - period)) {
              if (k == (quaterlyTableColumn.length - 1)) {
                if (mul2.indexOf(parseInt(quaterlyTableColumn[k].date.format('M'))) != -1) {
                  multiplier = 2
                } else if (mul1.indexOf(parseInt(quaterlyTableColumn[k].date.format('M'))) != -1) {
                  multiplier = 1
                } else {
                  multiplier = 1
                }
              }
            }
            rowData[quaterlyTableColumn[k].headerString] = Math.floor(perMonth * multiplier)
            remainingPeriod2 -= multiplier
            count++
          } else {
            rowData[quaterlyTableColumn[k].headerString] = 0
          }
        }
      } else {
        rowData[quaterlyTableColumn[k].headerString] = 0
        currentDate = moment(quaterlyTableColumn[k].date)
      }
    }

    // ************* monthly table row loop **************************
    for (l = 0; l < monthlyTableColumn.length; l++) {
      let columnHeaderYear = parseInt(moment(monthlyTableColumn[l].date).format('YYYY'))
      let headerMonth = parseInt(moment(monthlyTableColumn[l].date).format('MM'))
      if (columnHeaderYear >= currentDateYear) {
        if ((columnHeaderYear == currentDateYear) && headerMonth < currentMonth2) {
          rowData[monthlyTableColumn[l].headerString] = 0
        } else {
          if (remainingPeriod3 > 0) {
            rowData[monthlyTableColumn[l].headerString] = Math.floor(perMonth)
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

function createQuaterlyTableColumn(quarters, range) {
  let temp = []
  quarters.forEach(month => {
    temp.push({ date: month, headerString: `Q${Math.floor((month.toDate().getMonth() + 3) / 3)} ${month.format('YYYY')}` })
  })
  if (moment(quarters[quarters.length - 1]).isBefore(moment(range.end))) {
    temp.push({ date: range.end, headerString: `Q${Math.floor(((range.end).toDate().getMonth() + 3) / 3)} ${(range.end).format('YYYY')}` })
  }
  return temp
}

function createAnnualTableColumns(range) {
  let startYear = moment(range.start).format('YYYY')
  let endYear = moment(range.end).format('YYYY')
  let temp = []
  for (let i = startYear; i <= endYear; i++) {
    temp.push(i)
  }
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
