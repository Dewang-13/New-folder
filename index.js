const puppeteer = require('puppeteer');
const excelToJson = require('convert-excel-to-json');
const fs = require('fs');
const robot = require('robotjs');








(async () => {
    const browser = await puppeteer.launch({
        headless: false,
        defaultViewport: null
    }
    );
    const page = await browser.newPage();
    await page.setDefaultNavigationTimeout(0);
    await page.goto('https://cuh.campuspro.in/StudentFeePayment.aspx');
    const n = await page.$("#txtRollNo");
    await page.waitForSelector("#txtRollNo");
    await n.type(object.A + "");
    await page.waitForTimeout(1000);
    //dob
    const m = await page.$("#txtDOB");
    await page.waitForSelector("#txtDOB");
    await m.type(ndob);
    await page.waitForTimeout(1000);

    const btn = await page.$("#btnSearch");
    await page.waitForSelector("#btnSearch");
    await btn.click();
    await page.waitForTimeout(1000);

    const Sbtn = await page.$("#grdSearchResult_btnSelect_0");
    await page.waitForSelector("#grdSearchResult_btnSelect_0");
    await Sbtn.click();
    await page.waitForTimeout(500);

    const Cbtn = await page.$(".swal-button--confirm");
    await page.waitForSelector(".swal-button--confirm");
    await Cbtn.click();
    await page.waitForTimeout(500);

    const Pbtn = await page.$("#btnPrint");
    await page.waitForSelector("#btnPrint");
    await Pbtn.click();
    await page.waitForTimeout(10000);

    robot.keyTap("s", "control");
    robot.typeString(` ${object.A}`);
    robot.setKeyboardDelay(1);
    robot.keyTap("enter");



    await page.waitForTimeout(10000);
    await browser.close();

})();

let rowData = fs.readFileSync('Result.json');
let student = JSON.parse(rowData);
//console.log(student.Sheet1[1]);





const result1 = excelToJson({
    source: fs.readFileSync('CSE.xlsx') // fs.readFileSync return a Buffer
});

fs.writeFileSync("Result.json", JSON.stringify(result1));


const add = "C:\Users\Dewang\Desktop\New folder";
let dob = object.B;
let month = dob.substring(5, 7);
const ndate = new Date(dob);
let Monthname = ndate.toLocaleString('default', { month: 'long' });
let ndob = dob.replace(dob.substring(5, 7), Monthname);
console.log(ndob);

