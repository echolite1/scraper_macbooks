const puppeteer = require('puppeteer');
const excel = require('excel4node');
const prompt = require('prompt-sync')();

// create workbook 
var workbook = new excel.Workbook();

// creating date
var today = new Date();
var dd = String(today.getDate()).padStart(2, '0');
var mm = String(today.getMonth() + 1).padStart(2, '0'); //January is 0!
var yyyy = today.getFullYear();
today = dd + '-' + mm + '-' + yyyy;

// input startYear
var year = prompt('Start scan from which Year? -> ');
var inputYear = year;
if (inputYear < 2015 || year < 2015){
    inputYear = 2015;
    year = 2015;
}

// main variables
const constantLink = 'https://www.rebuy.de/verkaufen/apple/notebooks/macbook';          // https://www.rebuy.de/verkaufen/apple/notebooks/macbook-pro/15?f_prop_season=2018
var modelSwitcher = 1;                     // default  == 1 // make 0 for 12"
var pageNumber = 1;
var maxYear = yyyy+1;               // < yyyy == < 2021 == 2020 the last one
var defaultTime = 850;              // 550 w/o doubler // until 14.09 == 750

// variables for excel
const qwerty = 'QWERTY';
const color = [
    'silber', 
    'space grau', 
    'gold', 
    'roségold'
];
const titles = [
    'Processor', 
    'RAM', 
    'SSD', 
    'Color', 
    'Tastatur', 
    'Max $$',
    'Perfect', 
    'Good',
    'Worst',
    'Profit'
];

// giving styles to sheet
const title = workbook.createStyle({
    alignment: {
        horizontal: 'center',
    },
    font: {
        bold: true,
        size: 15,
    },
});
const simple = workbook.createStyle({
    alignment: {
        horizontal: 'center',
    },
    font: {
        color: '#1b3e75',
    },
    fill: {
        type: 'pattern',
        patternType: 'solid',
        bgColor: '#e6edf7',
        fgColor: '#e6edf7',
    }
});
const red = workbook.createStyle({
    alignment: {
        horizontal: 'center',
    },
    font: {
      color: '#D64A42',
    },
});
const yellow = workbook.createStyle({
    alignment: {
        horizontal: 'center',
    },
    font: {
      color: '#c46518',
      bold: true,
    },
});
const bgGreen = workbook.createStyle({
    alignment: {
        horizontal: 'center',
    },
    font: {
        color: '#007a3b',
        bold: true,
        size: 13,
    },
    fill: {
        type: 'pattern',
        patternType: 'solid',
        bgColor: '#edf0eb',
        fgColor: '#edf0eb',
    }
});
//        ˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙
function delay(time) {
    return new Promise(function(resolve) { setTimeout(resolve, time) });
}
function excelStyles(anzeigeCounter){
    worksheet.cell(anzeigeCounter + 1, 7).style(bgGreen);
    worksheet.cell(anzeigeCounter + 1, 9).style(red);
    worksheet.cell(anzeigeCounter + 1, 8).style(yellow);
    worksheet.cell(anzeigeCounter + 1, 6).style(simple);
    worksheet.cell(anzeigeCounter + 1, 10).style(simple);
}
function setupPrices(priceWN, anzeigeCounter){
    if (isNaN(priceWN) ? worksheet.cell(anzeigeCounter + 1, 6).number(0) : worksheet.cell(anzeigeCounter + 1, 6).number(parseInt(priceWN * 0.863)));
    if (isNaN(priceWN) ? worksheet.cell(anzeigeCounter + 1, 7).number(0) : worksheet.cell(anzeigeCounter + 1, 7).number(priceWN));
    if (isNaN(priceWN) ? worksheet.cell(anzeigeCounter + 1, 8).number(0) : worksheet.cell(anzeigeCounter + 1, 8).number(parseInt(priceWN * 0.908)));
    if (isNaN(priceWN) ? worksheet.cell(anzeigeCounter + 1, 9).number(0) : worksheet.cell(anzeigeCounter + 1, 9).number(parseInt(priceWN * 0.818)));
    if (isNaN(priceWN) ? worksheet.cell(anzeigeCounter + 1, 10).number(0) : worksheet.cell(anzeigeCounter + 1, 10).number(parseInt(priceWN - ((priceWN * 0.908 + priceWN * 0.818)/2))));
}

//        ˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙
async function pro13(){
        // launching pseudo-browser
        const browser = await puppeteer.launch();//{headless: false, slowMo: 100}); // [_][_][_][_][_][_][_][_] 
        const page = await browser.newPage();
        console.clear();
    try{
        var anzeigeCounter = 1;
        var macModel = '-pro/13?f_prop_season=2020&page=' + pageNumber;                                      // -pro/14?f_prop_season= 2021 &page=1
        var link = constantLink + macModel;

        await page.goto(link);

        worksheet = workbook.addWorksheet(year + '2020 13 Pro');
        worksheet.column(1).setWidth(12);
        worksheet.column(2).setWidth(8);
        worksheet.column(7).setWidth(9);
        var titleanzeigeCounter = 1;
        while(titleanzeigeCounter <= titles.length){
            worksheet.cell(1, titleanzeigeCounter).string(titles[titleanzeigeCounter - 1]).style(title);
            titleanzeigeCounter++;
        }

        // проверять что вообще оно собирается выполнять if (все исключения)

        // check the number of the results
        [anzeigeNummer] = await page.$x('//*[@id="ry"]/body/main/div[1]/div[2]/div/div/div/div/h2/span[2]');
        nummer = await anzeigeNummer.getProperty('textContent');
        anzeigen = await nummer.jsonValue();
        anzeigen = parseInt(anzeigen.substr(1, 2));             // here we understand how much available
        console.log(year, 'Total:', anzeigen); // need it here

        // if (anzeigen > 24){
        //      link = link + '&page' + pageNumber          // could be implemented in the future
        // }

        if (isNaN(anzeigen) ? worksheet.cell(1, 12).string('unknown amount') : worksheet.cell(1, 12).string('Total: ' + anzeigen));
        if (anzeigen > 24 ? anzeigen = 24 : console.log(anzeigen));
        worksheetCounter = 1;
        
        for( ; anzeigeCounter < anzeigen + 1; anzeigeCounter++){ // here you can limit number of anzeigen
            // check if it is 'Kein Ankauf'
            [mainPageButton] = await page.$x('//*[@id="ry"]/body/main/div[1]/div[2]/div/div/div/div/div/div[' + anzeigeCounter + ']/a/div/div[4]/button/ng-switch/span');
            btnTxt = await mainPageButton.getProperty('textContent');
            btnTxtKA = await btnTxt.jsonValue();
            try{
                if(btnTxtKA != 'Kein Ankauf'){
                    // getting model text (DO NOT PRINT in the console)
                    [modelText] = await page.$x('//*[@id="ry"]/body/main/div[1]/div[2]/div/div/div/div/div/div[' + anzeigeCounter + ']/a/div/div[3]/text()');
                    label = await modelText.getProperty('textContent');
                    model = await label.jsonValue();
                    model = model.substr(model.search('G') - 4, 150);

                    // making Processor text look nice
                    processor = model.substr(model.search('G') - 4, 22);
                    if(processor.search('Chip') == -1){
                        processor = processor.replace('Intel Core ', '');
                        // if(processor.search(')') != -1){
                        //     processor = processor.replace(')', '');    // not now
                        // }
                    }
                    // making RAM text look nice
                    RAM = parseInt(model.substr(model.search('RAM') - 6, 2));
                    // making SSD text look nice
                    if(model.search('PCIe') != -1){
                        if(model.search('TB') != -1){
                            SSD   = model.substr(model.search('TB') - 2, 4);
                        }
                        else{
                            SSD = model.substr(model.search('SSD') - 12, 6);
                        } 
                    }
                    else{
                        if(model.search('TB') != -1){
                            SSD   = model.substr(model.search('TB') - 2, 4);
                        }
                        else{
                            SSD = model.substr(model.search('SSD') - 7, 6);
                        }
                    }

                    // excel styles
                    excelStyles(worksheetCounter);

                    // pressing verkaufen button
                    [mainPageButton] = await page.$x('//*[@id="ry"]/body/main/div[1]/div[2]/div/div/div/div/div/div[' + anzeigeCounter + ']/a/div/div[4]/button/ng-switch/span');
                    await mainPageButton.evaluate( mainPageButton => mainPageButton.click() );
                    await delay(defaultTime); // it is necessary

                    // doing Zustand WieNeu survey
                    var questionDiv = 1;
                    for( ; questionDiv < 5; questionDiv++){
                        [survey] = await page.$x('//*[@id="grading-form"]/div[1]/ry-grading-questions/div[' + questionDiv + ']/div/ry-grading-radio/div[1]/div[1]/label');
                        await survey.evaluate( survey => survey.click() ); //await delay(defaultTime); // not necessary
                    }
    
                    // getting the price value
                    await delay(defaultTime * 1.5); // it is necessary       --------------
                    [bestPrice] = await page.$x('//*[@id="grading-form"]/div[2]/ry-grading-info/div/div[2]/div[1]/div[2]/div[1]/p/span');
                    value = await bestPrice.getProperty('textContent');
                    priceWN = await value.jsonValue();
                    priceWN = parseInt(priceWN);

                    // return to the initial page with all results
                    await page.goto(link);
                    
                    //write to xlsx
                    worksheet.cell(worksheetCounter + 1, 1).string(processor);
                    worksheet.cell(worksheetCounter + 1, 2).number(RAM);
                    worksheet.cell(worksheetCounter + 1, 3).string(SSD);

                    if (model.search(color[0]) != -1){
                        worksheet.cell(worksheetCounter + 1, 4).string(color[0]);
                    } 
                    else if (model.search(color[1]) != -1){
                        worksheet.cell(worksheetCounter + 1, 4).string(color[1]);
                    }
                    else if (model.search(color[2]) != -1){
                        worksheet.cell(worksheetCounter + 1, 4).string(color[2]);
                    }
                    else if (model.search(color[3]) != -1){
                        worksheet.cell(worksheetCounter + 1, 4).string(color[3]);
                    }
                    else {
                        worksheet.cell(worksheetCounter + 1, 4).string(color[1]);
                    }

                    if (model.search(qwerty) != -1){ worksheet.cell(worksheetCounter + 1, 5).string(qwerty); }

                    setupPrices(priceWN, worksheetCounter);
                    
                    console.log(worksheetCounter, 'done');
                }
                else {
                    console.log(worksheetCounter, 'KA');
                }
                worksheetCounter++;
            }
            catch{
                console.log('specific mac failed');
            }
            if (anzeigeCounter == 24 && link == 'https://www.rebuy.de/verkaufen/apple/notebooks/macbook-pro/13?f_prop_season=2020&page=1'){
                link = 'https://www.rebuy.de/verkaufen/apple/notebooks/macbook-pro/13?f_prop_season=2020&page=2';
                console.log(link);
                await page.goto(link);
                anzeigeCounter = 0;
            }
        }
    } 
    catch{
        console.log();//model + ' is not available or other issue in loop');
    }
    workbook.write(today + '.xlsx');        // create output folder
    console.log('\n--- file created ---\n');
    browser.close();
}
//        ˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙
pro13();



//TODO:
// 12 inch что-то про модель после свитча (как если сделать 2020 на 15 дюймов)
// Починить ошибку 500 (catch) и специфик модел (красная надпись на сайте)