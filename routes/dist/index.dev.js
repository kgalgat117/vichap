"use strict";

var express = require('express');

var XLSX = require('xlsx');

var path = require('path');

var Moment = require('moment');

var MomentRange = require('moment-range');

var moment = MomentRange.extendMoment(Moment);
var router = express.Router();
/* GET home page. */

router.get('/', function (req, res, next) {
  res.render('index', {
    title: 'Express'
  });
});
router.post('/submit', validateData, function (req, res) {
  try {
    var jsonData = createJSONData({
      start: req.body.start,
      end: req.body.end,
      period: req.body.period
    });
    createFile(jsonData);
    var filePath = path.resolve('./test.xlsx');
    res.download(filePath);
  } catch (err) {
    console.log(err);
    res.status(400).json(err);
  }
});
var column1 = '1';
var column2 = '2';

function createJSONData(data) {
  var tableData = [];
  var start = new Date(data.start);
  var end = new Date(data.end);
  var range = moment.range(start, end);
  var years = Array.from(range.by('month'));
  var annualTableColumn = createAnnualTableColumns(Array.from(range.by('year')));
  var quaterlyTableColumn = createQuaterlyTableColumn(Array.from(range.by('quarter')));
  var monthlyTableColumn = createMonthlyTableColumn(Array.from(range.by('month')));
  var period = data.period;

  for (var i = 0; i < years.length; i++) {
    var randomNumber = getRandomNumber();
    var rowData = {};
    rowData[column1] = years[i].format('MM/YYYY');
    rowData[column2] = randomNumber;
    var perMonth = randomNumber / period;
    var currentDateYear = years[i].format('YYYY');
    var remainingPeriod = period;
    var remainingPeriod2 = period;
    var remainingPeriod3 = period;
    var currentMonth = years[i].format('MM');
    var currentMonth2 = years[i].format('MM');
    var currentQuater = getQuater(years[i]);

    for (var j = 0; j < annualTableColumn.length; j++) {
      var columnHeaderYear = parseInt(annualTableColumn[j]);
      var remainingMonthOfCurrentColumnHeaderYear = 12 - currentMonth + 1;

      if (columnHeaderYear >= currentDateYear) {
        if (remainingPeriod > 0) {
          if (remainingPeriod < remainingMonthOfCurrentColumnHeaderYear) {
            remainingMonthOfCurrentColumnHeaderYear = remainingPeriod;
            remainingPeriod = 0;
          } else {
            remainingPeriod -= remainingMonthOfCurrentColumnHeaderYear;
          }

          rowData[annualTableColumn[j]] = remainingMonthOfCurrentColumnHeaderYear * perMonth;
        } else {
          rowData[annualTableColumn[j]] = 0;
        }
      } else {
        rowData[annualTableColumn[j]] = 0;
      }

      currentMonth = 1;
    }

    var currentDate = moment(years[i]);

    for (var k = 0; k < quaterlyTableColumn.length; k++) {
      var _columnHeaderYear = parseInt(moment(quaterlyTableColumn[k].date).format('YYYY'));

      var headerQuater = getQuater(quaterlyTableColumn[k].date);

      if (_columnHeaderYear >= currentDateYear) {
        if (_columnHeaderYear == currentDateYear && headerQuater < currentQuater) {
          rowData[quaterlyTableColumn[k].headerString] = 0;
          currentDate = moment(quaterlyTableColumn[k].date);
        } else {
          if (remainingPeriod2 > 0) {
            var multiplier = Math.ceil(moment(quaterlyTableColumn[k].date).diff(currentDate, 'month', true)) || 1;

            if (remainingPeriod2 < multiplier) {
              multiplier = remainingPeriod2;
            }

            rowData[quaterlyTableColumn[k].headerString] = perMonth * multiplier;
            currentDate = moment(quaterlyTableColumn[k].date);
            remainingPeriod2 -= multiplier;
          } else {
            rowData[quaterlyTableColumn[k].headerString] = 0;
          }
        }
      } else {
        rowData[quaterlyTableColumn[k].headerString] = 0;
        currentDate = moment(quaterlyTableColumn[k].date);
      }
    }

    for (l = 0; l < monthlyTableColumn.length; l++) {
      var _columnHeaderYear2 = parseInt(moment(monthlyTableColumn[l].date).format('YYYY'));

      var headerMonth = parseInt(moment(monthlyTableColumn[l].date).format('MM'));

      if (_columnHeaderYear2 >= currentDateYear) {
        if (_columnHeaderYear2 == currentDateYear && headerMonth < currentMonth2) {
          rowData[monthlyTableColumn[l].headerString] = 0;
        } else {
          if (remainingPeriod3 > 0) {
            rowData[monthlyTableColumn[l].headerString] = perMonth;
            remainingPeriod3--;
          } else {
            rowData[monthlyTableColumn[l].headerString] = 0;
          }
        }
      } else {
        rowData[monthlyTableColumn[l].headerString] = 0;
      }
    }

    tableData.push(rowData);
  }

  return tableData;
}

function createMonthlyTableColumn(months) {
  var temp = [];
  months.forEach(function (month) {
    temp.push({
      date: month,
      headerString: "".concat(month.format('MM/YYYY'))
    });
  });
  return temp;
}

function getQuater(momentDate) {
  return Math.floor((momentDate.toDate().getMonth() + 3) / 3);
}

function createQuaterlyTableColumn(quarters) {
  var temp = [];
  quarters.forEach(function (month) {
    temp.push({
      date: month,
      headerString: "Q".concat(Math.floor((month.toDate().getMonth() + 3) / 3), " ").concat(month.format('YYYY'))
    });
  });
  return temp;
}

function createAnnualTableColumns(years) {
  var temp = [];
  years.forEach(function (month) {
    temp.push(month.format('YYYY'));
  });
  return temp;
}

function getRandomNumber() {
  var min = 1,
      max = 100000;
  return Math.floor(Math.random() * (max - min + 1)) + min;
}

function createFile(json) {
  var fileName = 'test.xlsx';
  var ws = XLSX.utils.json_to_sheet(json);
  var wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'test');
  XLSX.writeFile(wb, fileName);
}

function readFile() {
  var filePath = path.resolve('./public/images/dates.xlsx');
  var workbook = XLSX.readFile(filePath);
  var sheet_name_list = workbook.SheetNames;
  var xlData = XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[0]]);
  return xlData;
}

function validateData(req, res, next) {
  if (req.body.start && req.body.end && req.body.period) {
    try {
      var startDate = new Date(req.body.start);
      var endDate = new Date(req.body.end);

      if (endDate < startDate) {
        res.status(400).json({
          error: 'end date should be after start date'
        });
      } else {
        next();
      }
    } catch (error) {
      res.status(400).json({
        error: 'invalid data'
      });
    }
  } else {
    res.status(400).json({
      error: 'invalid data'
    });
  }
}

module.exports = router;