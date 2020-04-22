"use strict";var express=require("express"),XLSX=require("xlsx"),path=require("path"),Moment=require("moment"),MomentRange=require("moment-range"),moment=MomentRange.extendMoment(Moment),router=express.Router();router.get("/",function(e,t,r){t.render("index",{title:"Express"})}),router.post("/submit",validateData,function(e,t){try{createFile(createJSONData({start:e.body.start,end:e.body.end,period:e.body.period}));var r=path.resolve("./test.xlsx");t.download(r)}catch(e){console.log(e),t.status(400).json(e)}});var column1="1",column2="2";function createJSONData(e){for(var t=[],r=new Date(e.start),a=new Date(e.end),n=moment.range(r,a),o=Array.from(n.by("month")),s=createAnnualTableColumns(Array.from(n.by("year"))),u=createQuaterlyTableColumn(Array.from(n.by("quarter"))),m=createMonthlyTableColumn(Array.from(n.by("month"))),d=e.period,i=0;i<o.length;i++){var f=getRandomNumber(),c={};c[column1]=o[i].format("MM/YYYY");for(var h=(c[column2]=f)/d,g=o[i].format("YYYY"),Y=d,b=d,p=d,v=o[i].format("MM"),M=o[i].format("MM"),y=getQuater(o[i]),S=0;S<s.length;S++){var x=12-v+1;g<=parseInt(s[S])&&0<Y?(Y<x?(x=Y,Y=0):Y-=x,c[s[S]]=x*h):c[s[S]]=0,v=1}for(var X=moment(o[i]),D=0;D<u.length;D++){var q=parseInt(moment(u[D].date).format("YYYY")),w=getQuater(u[D].date);if(g<=q)if(q==g&&w<y)c[u[D].headerString]=0,X=moment(u[D].date);else if(0<b){var L=moment(u[D].date);0==D&&L.endOf("quarter");var _=Math.ceil(L.diff(X,"month",!0))||1;b<_&&(_=b),c[u[D].headerString]=h*_,X=moment(u[D].date),b-=_}else c[u[D].headerString]=0;else c[u[D].headerString]=0,X=moment(u[D].date)}for(l=0;l<m.length;l++){var j=parseInt(moment(m[l].date).format("YYYY")),A=parseInt(moment(m[l].date).format("MM"));g<=j?j==g&&A<M?c[m[l].headerString]=0:0<p?(c[m[l].headerString]=h,p--):c[m[l].headerString]=0:c[m[l].headerString]=0}t.push(c)}return t}function createMonthlyTableColumn(e){var t=[];return e.forEach(function(e){t.push({date:e,headerString:"".concat(e.format("MM/YYYY"))})}),t}function getQuater(e){return Math.floor((e.toDate().getMonth()+3)/3)}function createQuaterlyTableColumn(e){var t=[];return e.forEach(function(e){t.push({date:e,headerString:"Q".concat(Math.floor((e.toDate().getMonth()+3)/3)," ").concat(e.format("YYYY"))})}),t}function createAnnualTableColumns(e){var t=[];return e.forEach(function(e){t.push(e.format("YYYY"))}),t}function getRandomNumber(){return Math.floor(1e5*Math.random())+1}function createFile(e){var t=XLSX.utils.json_to_sheet(e),r=XLSX.utils.book_new();XLSX.utils.book_append_sheet(r,t,"test"),XLSX.writeFile(r,"test.xlsx")}function readFile(){var e=path.resolve("./public/images/dates.xlsx"),t=XLSX.readFile(e),r=t.SheetNames;return XLSX.utils.sheet_to_json(t.Sheets[r[0]])}function validateData(e,t,r){if(e.body.start&&e.body.end&&e.body.period)try{var a=new Date(e.body.start);new Date(e.body.end)<a?t.status(400).json({error:"end date should be after start date"}):r()}catch(e){t.status(400).json({error:"invalid data"})}else t.status(400).json({error:"invalid data"})}module.exports=router;