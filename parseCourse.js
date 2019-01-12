const XLSX = require("xlsx")
const fs = require('fs')

const semester = "2018_Fall"
let workbook = XLSX.readFile(__dirname + "/" + semester + ".xlsx")
let worksheet = workbook.Sheets["Courses Offered"]

var indexNames = new Array()
var columnChar = new Array()

//  column characters from A to AZ
for(var i=0;i<52;i++) {
    var numIndex = i
    var charIndex = ""
    if (i >= 26) {
      charIndex += "A"
      numIndex -= 26
    } 
    columnChar.push(charIndex + String.fromCharCode("A".charCodeAt(0)+numIndex))
}

// parse indexName of table
for(var i=0;;i++) {
    var addressCell = columnChar[i] + 2
    var desiredCell = worksheet[addressCell]
    var indexName = (desiredCell ? desiredCell.v : undefined)
    if (indexName === undefined)
      break
    indexNames.push(indexName)
}

var indexLength = indexNames.length
var courses = new Array()

//parse data and make JSON file
for(var i=3;;i++) {
    var course = new Object()
    for(var j=0;j<indexLength;j++) {
        var addressCell = columnChar[j] + i
        var desiredCell = worksheet[addressCell]
        var courseInfo = (desiredCell ? desiredCell.v.toString().trim() : "")
        
        //end of data
        if (j===0 && courseInfo === "") {
            course[indexNames[0]] = ""
            break
        }
        
        const courseFunc = (key, info) => {
            if (info === "") return false

            switch(key) {
                case "강의계획서": return true
                case "AU": return Number(info)
                case "강:실:학":
                  var ksh = info.split(":")
                  return {"강의":Number(ksh[0]), "실험":Number(ksh[1]), "학점":Number(ksh[2])}
                case "영어": return true
                case "Edu 4.0": return true
                case "정원" : return Number(info)
                case "수강인원" : return Number(info)
                case "강의시간" :
                  var dts = info.split("\r\r\n")
                  var dtArray = new Array()
                  for (var i in dts) {
                      var x = dts[i].split(" ")
                      dtArray.push({"요일":x[0], "시간":x[1]})
                  }
                  return dtArray
                case "강의실" :
                  return info.split("\r\r\n")
                case "시험시간" :
                  var examTime = info.split(" ")
                  return {"요일":examTime[0], "시간":examTime[1]}
                case "학과코드" : return Number(info)
                default : return info
            }
        }
        course[indexNames[j]] = courseFunc(indexNames[j], courseInfo)
    }

    if(course[indexNames[0]] === "")
        break

    courses.push(course)
}

var json = new Object()
json[semester] = courses

fs.writeFile(semester+".json", JSON.stringify(json), 'utf8', (err) => {
    if (err) {
        console.log(err)
    }
})