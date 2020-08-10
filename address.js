const request = require('request');
const xlsx = require("xlsx");


const { apikey } = {
    apikey: 'YOUR API KEY',
};

var wb = xlsx.readFile("kordinat.xlsx") // okunacak olan xlsx dosyası 
var ws = wb.Sheets["Sayfa1"] // okunacak olan sheet adı


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
        request(opt(element.Y + "," + element.X), (error, response, body) => {
            if (error) {
                printXlsx(data)
                console.log("dosya hatalı kaydedildi thrown" + findGeo)
                throw new Error(error);
            }
            const json = JSON.parse(body)
            try {
                if (json.response.GeoObjectCollection.metaDataProperty.GeocoderResponseMetaData.found > 0) {
                    var adress = json.response.GeoObjectCollection.featureMember[0]
                        .GeoObject.metaDataProperty.GeocoderMetaData.text
                    data[index].Adres = adress

                } else {
                    data[index].Adres = ""
                }
                findGeo++;
                console.log("Tamamlandı " + findGeo)
                if (findGeo == dataLen) printXlsx(data)
            } catch (er) {
                printXlsx(data)
                console.log("dosya hatalı kaydedildi " + findGeo)
                console.log(er)
                return
            }
        })
    })
}

function printXlsx(data) {
    var newWB = xlsx.utils.book_new()
    var newWS = xlsx.utils.json_to_sheet(data)
    xlsx.utils.book_append_sheet(newWB, newWS, "Sayfa1")
    xlsx.writeFile(newWB, "Adresler.xlsx")
    console.log("Yeni xlsx dosyası oluşturuldu ")
}