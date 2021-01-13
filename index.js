const request = require('request-promise-native');
const cheerio = require('cheerio');
const fs = require('fs');
const _ = require("lodash");


const Init = async () => {
    return new Promise((resolve, reject) => {
        request({ url: "https://danhmuchanhchinh.gso.gov.vn/", resolveWithFullResponse: true }).then(response => {
            let $ = cheerio.load(response.body);
            var state = $('#__VIEWSTATE').first().attr("value");

            resolve({
                state,
                cookie: response.headers["set-cookie"][0].split(";")[0]
            })
        }).catch(err => {
            console.log(err);
            reject(err);
        });
    })
}
const Download = async (x) => {
    console.log("start downloading");
    return new Promise((resolve, reject) => {
        var options = {
            'method': 'POST',
            'url': 'https://danhmuchanhchinh.gso.gov.vn/default.aspx',

            'headers': {
                'Connection': 'keep-alive',
                'Pragma': 'no-cache',
                'Cache-Control': 'no-cache',
                'Upgrade-Insecure-Requests': '1',
                'Origin': 'https://danhmuchanhchinh.gso.gov.vn',
                'Content-Type': 'application/x-www-form-urlencoded',
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/87.0.4280.88 Safari/537.36 Edg/87.0.664.66',
                'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
                'Sec-Fetch-Site': 'same-origin',
                'Sec-Fetch-Mode': 'navigate',
                'Sec-Fetch-User': '?1',
                'Sec-Fetch-Dest': 'document',
                'Referer': 'https://danhmuchanhchinh.gso.gov.vn/',
                'Accept-Language': 'vi,en-US;q=0.9,en;q=0.8',
                'Cookie': x.cookie,
                'gzip': true,
            },
            form: {
                '__EVENTTARGET': 'ctl00$PlaceHolderMain$btnExcel',
                '__EVENTARGUMENT': 'Click',
                '__VIEWSTATE': x.state,
                'ctl00_PlaceHolderMain_cmbCap_VI': '1',
                'ctl00$PlaceHolderMain$cmbCap': 'Tỉnh',
                'ctl00$PlaceHolderMain$check': 'C',
            }
        };

        let file = fs.createWriteStream('data.xls');
        request(options, function (error, response) {
            if (error) throw new Error(error);
        }).pipe(file)
            .on("finish", () => resolve())
            .on("error", (err) => reject(err));
    })
};


const parseJson = () => {
    var XLSX = require('xlsx');
    var workbook = XLSX.readFile('./data.xls');
    var sheet_name_list = workbook.SheetNames;
    sheet_name_list.forEach(function (y) {
        var worksheet = workbook.Sheets[y];
        var headers = {};
        var data = [];
        for (z in worksheet) {
            if (z[0] === '!') continue;
            //parse out the column, row, and value
            var col = z.substring(0, 1);
            var row = parseInt(z.substring(1));
            var value = worksheet[z].v;

            //store header names
            if (row == 1) {
                headers[col] = value;
                continue;
            }

            if (!data[row]) data[row] = {};
            data[row][headers[col]] = value;
        }
        //drop those first two rows which are empty
        data.shift();
        data.shift();

        const all = [];

        data.map(item => {
            const provinceId = item["Mã TP"];
            const provinceName = item["Tỉnh Thành Phố"];
            const districtId = item["Mã QH"];
            const districtName = item["Quận Huyện"];
            const wardId = item["Mã PX"];
            const wardName = item["Phường Xã"];
            const level = item["Cấp"];

            if (all.findIndex(x => x.Id == provinceId) < 0) {
                all.push({
                    Id: provinceId,
                    Name: provinceName,
                    Districts: []
                })
            }
            const index = all.findIndex(x => x.Id == provinceId);

            if (all[index].Districts.findIndex(x => x.Id == districtId) < 0) {
                all[index].Districts.push({
                    Id: districtId,
                    Name: districtName,
                    Wards: []
                })
            }
            dIndex = all[index].Districts.findIndex(x => x.Id == districtId);
            all[index].Districts[dIndex].Wards.push({
                Id: wardId,
                Name: wardName,
                Level: level,
            })
        })

        fs.writeFile("data.json", JSON.stringify(all), null, () => {
            console.log("parse json success")
        })
    });
}


(async () => {
    var x = await Init();
    await Download(x);
    parseJson();
})();