const puppeteer = require("puppeteer");
const { Parser } = require('json2csv');
const fs = require('fs');
const XLSX = require('xlsx');

const reviewCrawling = async (pageUrl, fileName) => {
    try {
        const workbook = XLSX.readFile('Zarina_url.xlsx');
        const sheet_name_list = workbook.SheetNames;
        const data = XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[0]]);
        const browser = await puppeteer.launch(); // headless 브라우저 실행
        const page = await browser.newPage(); // 새로운 페이지 열기
        // 해당 URL에 접속
        // await page.goto("https://www.wildberries.ru/catalog/5966627/otzyvy?page={}&field=Date&order=Desc&link=4773432");
        await page.goto(pageUrl);
        await page.click('a.show-more');
        // evaluate() 함수를 이용해 History table을 선택하고 반복문으로 내용을 빈배열에 추가한다
        const reactHistory = await page.evaluate(async() => {
            let commentsDOM;
            let foo = function() {
                return new Promise (function (resolve, reject) {
                    setTimeout(function() {
                        if (document.querySelector('a.show-more') !== null) {
                            document.querySelector('a.show-more').click();
                        }
                        commentsDOM = document.getElementsByClassName('comment j-b-comment');
                        console.log(commentsDOM.length);
                        resolve();
                    }, 1000);
                })
            };
            let working = true;
            while (working) {
                await foo();
                if (document.querySelector('a.show-more') === null) {
                    working = false;
                }
            }
            
            
            // 반복문으로 <tbody> 내용 객체 형식으로 빈 배열에 추가
            let scrappedData = []; // 스크래핑 내용 담을 빈 배열
            for (let i = 0; i < commentsDOM.length; i++) {
                scrappedData.push({
                    customer_id: commentsDOM[i].getElementsByTagName('label')[0].getAttribute('data-user-id'),
                    date: commentsDOM[i].getElementsByClassName('time')[0].innerText,
                    rating: commentsDOM[i].getElementsByClassName('comment-rating')[0].innerText,
                    reviewText: commentsDOM[i].getElementsByClassName('body')[0].innerText,
                    imageCount: commentsDOM[i].getElementsByClassName('j-b-imgview')[0] === undefined 
                                ? 0 
                                : commentsDOM[i].getElementsByClassName('j-b-imgview')[0].getElementsByTagName('span').length,
                    helpfulCount: commentsDOM[i].getElementsByClassName('yescnt')[0].innerText,
                    nonHelpfulCount: commentsDOM[i].getElementsByClassName('nocnt')[0].innerText
                })
            }
            
            return scrappedData;
        });
        const parser = new Parser();
        const csv = parser.parse(reactHistory);
        
        fs.writeFile(`./data/${fileName}.csv`, csv, 'utf8', (err) => {
            if (err) {
                console.log('에러 : ', err);
            }
        })
        await browser.close();
    } catch (err) {
        console.log(err);
    }
};

const workbook = XLSX.readFile('Zarina_url.xlsx');
const sheet_name_list = workbook.SheetNames;
const data = XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[0]]);

(async () => {
    for (let i = 0; i < data.length; i++) {
        console.log((i + 1), '번째 url 크롤링 중');
        await reviewCrawling(data[i].Url, data[i].Item_Id);
    }
})();
