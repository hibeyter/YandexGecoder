'use strict';

const request = require('request');
const xlsx = require("xlsx");


const { apikey } = {
    apikey: 'YOUR APİ KEY',
};


var wb = xlsx.readFile("adres.xlsx")
var ws = wb.Sheets["Sayfa1"]
var data = xlsx.utils.sheet_to_json(ws)
const dataLen = data.length
var findGeo = 0;
var opt = function(adres) {
    return {
        method: 'GET',
        url: 'https://geocode-maps.yandex.ru/1.x/',
        qs: {
            apikey,
            geocode: adres,
            format: 'json',
            lang: 'tr-TR'
        },
    };
}

readXlsx()

function readXlsx() {
    data.forEach((element, index) => {
        request(opt(element.adres), (error, response, body) => {
            if (error) throw new Error(error);
            const json = JSON.parse(body)
            var point = json.response.GeoObjectCollection.featureMember[0]
                .GeoObject.Point.pos
            var space = point.indexOf(' ')
            var pointX = point.substring(0, space)
            var pointY = point.substring(space, point.lenght)
            data[index].X = parseFloat(pointX)
            data[index].Y = parseFloat(pointY)
            findGeo++;
            if (findGeo == dataLen) printXlsx(data)
        })
    })
}

function printXlsx(data) {
    var newWB = xlsx.utils.book_new()
    var newWS = xlsx.utils.json_to_sheet(data)
    xlsx.utils.book_append_sheet(newWB, newWS, "AllCoordinate")
    xlsx.writeFile(newWB, "NewFile.xlsx")
}