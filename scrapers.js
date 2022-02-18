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
if (inputYear < 2018 || year < 2018){
    inputYear = 2018;
    year = 2018;
}

// main variables
const constantLink = 'https://www.rebuy.de/verkaufen/apple/notebooks/macbook';          // https://www.rebuy.de/verkaufen/apple/notebooks/macbook-pro/15?f_prop_season=2018
var modelSwitcher = 1;                     // default  == 1 // make 0 for 12"
//var pageNumber = 1;
var maxYear = yyyy;               // < yyyy == < 2021 == 2020 the last one
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
function formatTables() {
    worksheet.column(1).setWidth(12);
    worksheet.column(2).setWidth(8);
    worksheet.column(7).setWidth(9);
}
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
async function scrapeMacs(){
    
    // launching pseudo-browser
    const browser = await puppeteer.launch();//{headless: false, slowMo: 450}); // [_][_][_][_][_][_][_][_]
    const page = await browser.newPage();
    console.clear();
    console.log('- - - - - NEW SCAN ' + today + ' - - - - -')

    // even if unblock 13 Pro - will not work, new 16 inch in beta
    var { anzeigeCounter, link, titleanzeigeCounter, questionDiv } = await oldMacs();

    var answer = prompt('Proceed to old 16 inch? -> ');
    if(answer == 1 || answer == 'y' || answer == 'Y'){
        var { anzeigeCounter, link, titleanzeigeCounter, questionDiv } = await old16();
    }
    else {
        console.log('old 16 inch ignored')
    }

    answer = prompt('Proceed to new 16 inch? -> ');
    if(answer == 1 || answer == 'y' || answer == 'Y'){
        var { anzeigeCounter, link, titleanzeigeCounter, questionDiv } = await new16();
    }
    else {
        console.log('new 16 inch ignored')
    }
    
    //var { anzeigeCounter, link, titleanzeigeCounter, questionDiv } = await new14();

    async function oldMacs() {
        // 31.01.2022 var pageNumber = 1;
        while (modelSwitcher < 2) { //default from 1 < 3       // пока что 12 (0) и Эир (4) не нужны // вместо цифры вставить ответ от что сканить
            while (year < maxYear) {
                try {
                    var anzeigeCounter = 1; // 100% можно переместить свитч
                    var pageNumber = 1;

                    // ignore creating of non-existing models and creating worksheets
                    switch (modelSwitcher) {
                        case 0: // 12 inch
                            if (link.search('2018') != -1 || link.search('2019') != -1 || link.search('2020') != -1 || link.search('2021') != -1) {
                                console.log('2018/-19/-20 12 inch -'); //, link);
                            }
                            else {
                                worksheet = workbook.addWorksheet(year + ' 12 inch');
                            }
                            break;
                        case 1: // 15 inch
                            if (year < 2018 || year > 2019) {
                                console.log('-- wrong 15 inch --');
                                year++;
                            }
                            else {
                                worksheet = workbook.addWorksheet(year + ' 15 Pro');
                            }
                            break;
                        case 2: // 13 inch
                            if (year < 2020 || year > 2020) {
                                console.log('-- wrong 13 inch --');
                                year++;
                                worksheet = workbook.addWorksheet(year + ' 13 Pro');
                            }
                            else {
                                worksheet = workbook.addWorksheet(year + ' 13 Pro');
                            }
                            break;
                        case 3: // Air
                            if (year < 2020) {
                                console.log('-- wrong Air --');
                                year = 2020;
                            }
                            else {
                                worksheet = workbook.addWorksheet(year + ' Air');
                            }
                            break;
                        default:
                            console.log('modelSwitcher error');
                    }

                    var macModel = [
                        '?f_prop_season=' + year + '&page=' + pageNumber,
                        '-pro/15?f_prop_season=' + year + '&page=' + pageNumber,
                        '-pro/13?f_prop_season=' + year + '&page=' + pageNumber,
                        '-air/13?f_prop_season=' + year + '&page=' + pageNumber
                    ];
                    var link = constantLink + macModel[modelSwitcher];
                    await page.goto(link);

                    // make tables look better
                    formatTables();

                    var titleanzeigeCounter = 1;
                    while (titleanzeigeCounter <= titles.length) {
                        worksheet.cell(1, titleanzeigeCounter).string(titles[titleanzeigeCounter - 1]).style(title);
                        titleanzeigeCounter++;
                    }

                    // check the number of the results
                    [anzeigeNummer] = await page.$x('//*[@id="ry"]/body/main/div[1]/div[2]/div/div/div/div/h2/span[2]');
                    nummer = await anzeigeNummer.getProperty('textContent');
                    anzeigen = await nummer.jsonValue();
                    anzeigen = parseInt(anzeigen.substr(1, 2)); // here we understand how much available
                    console.log(macModel[modelSwitcher], year, 'Total:', anzeigen); // need it here

                    if (isNaN(anzeigen) ? worksheet.cell(1, 12).string('unknown amount') : worksheet.cell(1, 12).string('Total: ' + anzeigen));
                    if (anzeigen > 24 ? anzeigen = 24 : console.log(anzeigen));
                    worksheetCounter = 1;

                    for (; anzeigeCounter < anzeigen + 1; anzeigeCounter++) { // here you can limit number of anzeigen 'for( ; anzeigeCounter < anzeigen + 1; anzeigeCounter++){
                        // check if it is 'Kein Ankauf'
                        [mainPageButton] = await page.$x('//*[@id="ry"]/body/main/div[1]/div[2]/div/div/div/div/div/div[' + anzeigeCounter + ']/a/div/div[4]/button/ng-switch/span');
                        btnTxt = await mainPageButton.getProperty('textContent');
                        btnTxtKA = await btnTxt.jsonValue();
                        try {
                            if (btnTxtKA != 'Kein Ankauf') {
                                // getting model text (DO NOT PRINT in the console)
                                [modelText] = await page.$x('//*[@id="ry"]/body/main/div[1]/div[2]/div/div/div/div/div/div[' + anzeigeCounter + ']/a/div/div[3]/text()');
                                label = await modelText.getProperty('textContent');
                                model = await label.jsonValue();
                                model = model.substr(model.search('G') - 4, 150);

                                // making Processor text look nice
                                processor = model.substr(model.search('G') - 4, 22);
                                if (processor.search('Chip') == -1) {
                                    processor = processor.replace('Intel Core ', '');
                                    // if(processor.search(')') != -1){
                                    //     processor = processor.replace(')', '');    // not now
                                    // }
                                }
                                // making RAM text look nice
                                RAM = parseInt(model.substr(model.search('RAM') - 6, 2));
                                // making SSD text look nice
                                if (model.search('PCIe') != -1) {
                                    if (model.search('TB') != -1) {
                                        SSD = model.substr(model.search('TB') - 2, 4);
                                    }
                                    else {
                                        SSD = model.substr(model.search('SSD') - 12, 6);
                                    }
                                }
                                else {
                                    if (model.search('TB') != -1) {
                                        SSD = model.substr(model.search('TB') - 2, 4);
                                    }
                                    else {
                                        SSD = model.substr(model.search('SSD') - 7, 6);
                                    }
                                }

                                // excel styles
                                excelStyles(worksheetCounter);

                                // pressing verkaufen button
                                [mainPageButton] = await page.$x('//*[@id="ry"]/body/main/div[1]/div[2]/div/div/div/div/div/div[' + anzeigeCounter + ']/a/div/div[4]/button/ng-switch/span');
                                await mainPageButton.evaluate(mainPageButton => mainPageButton.click());
                                await delay(defaultTime); // it is necessary

                                // doing Zustand WieNeu survey
                                var questionDiv = 1;
                                for (; questionDiv < 5; questionDiv++) {
                                    [survey] = await page.$x('//*[@id="grading-form"]/div[1]/ry-grading-questions/div[' + questionDiv + ']/div/ry-grading-radio/div[1]/div[1]/label');
                                    await survey.evaluate(survey => survey.click()); //await delay(defaultTime); // not necessary
                                }

                                // getting the price value
                                await delay(defaultTime * 1.5); // it is necessary       --------------
                                [bestPrice] = await page.$x('//*[@id="grading-form"]/div[2]/ry-grading-info/div/div[2]/div[1]/div[2]/div[1]/p/span');
                                value = await bestPrice.getProperty('textContent');
                                priceWN = await value.jsonValue();
                                priceWN = parseInt(priceWN);

                                // return to the initial page with all results
                                await page.goto(link); // check if it is needed

                                //write to xlsx
                                worksheet.cell(worksheetCounter + 1, 1).string(processor);
                                worksheet.cell(worksheetCounter + 1, 2).number(RAM);
                                worksheet.cell(worksheetCounter + 1, 3).string(SSD);

                                if (model.search(color[0]) != -1) {
                                    worksheet.cell(worksheetCounter + 1, 4).string(color[0]);
                                }
                                else if (model.search(color[1]) != -1) {
                                    worksheet.cell(worksheetCounter + 1, 4).string(color[1]);
                                }
                                else if (model.search(color[2]) != -1) {
                                    worksheet.cell(worksheetCounter + 1, 4).string(color[2]);
                                }
                                else if (model.search(color[3]) != -1) {
                                    worksheet.cell(worksheetCounter + 1, 4).string(color[3]);
                                }
                                else {
                                    worksheet.cell(worksheetCounter + 1, 4).string(color[1]);
                                }

                                if (model.search(qwerty) != -1) { worksheet.cell(worksheetCounter + 1, 5).string(qwerty); }

                                setupPrices(priceWN, worksheetCounter);

                                console.log(worksheetCounter, 'done');
                            }
                            else {
                                console.log(worksheetCounter, 'KA');
                            }
                            worksheetCounter++;
                        }
                        catch {
                            console.log('specific mac failed');
                        }
                        if (anzeigeCounter == 24) {
                            link = link.replace('e=' + pageNumber, 'e=' + ++pageNumber);

                            await page.goto(link);

                            anzeigeCounter = 0;
                        }
                    }
                }
                catch {
                    console.log('this model is not available');
                }
                year++;
            }
            modelSwitcher++;
            year = inputYear;
        }
        return { anzeigeCounter, link, titleanzeigeCounter, questionDiv };
    }

    async function old16() {
        var anzeigeCounter = 1;
        var link = constantLink + '-pro/16?f_prop_season=2019&page=1';
        try {
            await page.goto(link);
            await delay(defaultTime); //test OCT

            // check the number of the results
            [anzeigeNummer] = await page.$x('//*[@id="ry"]/body/main/div[1]/div[2]/div/div/div/div/h2/span[2]');
            nummer = await anzeigeNummer.getProperty('textContent');
            anzeigen = await nummer.jsonValue();
            anzeigen = parseInt(anzeigen.substr(1, 3));
            console.log('Total:', anzeigen); // need it here
            worksheet = workbook.addWorksheet('2019 16 Pro');
            if (isNaN(anzeigen) ? worksheet.cell(1, 12).string('unknown amount') : worksheet.cell(1, 12).string('Total: ' + anzeigen));

            // giving styles to sheet
            worksheet.column(1).setWidth(12);
            worksheet.column(2).setWidth(8);
            var titleanzeigeCounter = 1;
            while (titleanzeigeCounter <= titles.length) {
                worksheet.cell(1, titleanzeigeCounter).string(titles[titleanzeigeCounter - 1]).style(title);
                titleanzeigeCounter++;
            }

            for (; anzeigeCounter < anzeigen + 1; anzeigeCounter++) { //  for( ; anzeigeCounter < anzeigen + 1; anzeigeCounter++){
                // check if it is 'Kein Ankauf'
                [mainPageButton] = await page.$x('//*[@id="ry"]/body/main/div[1]/div[2]/div/div/div/div/div/div[' + anzeigeCounter + ']/a/div/div[4]/button/ng-switch/span');
                btnTxt = await mainPageButton.getProperty('textContent');
                btnTxtKA = await btnTxt.jsonValue();

                if (btnTxtKA != 'Kein Ankauf') {
                    // getting model text (DO NOT PRINT in the console)
                    [modelText] = await page.$x('//*[@id="ry"]/body/main/div[1]/div[2]/div/div/div/div/div/div[' + anzeigeCounter + ']/a/div/div[3]/text()');
                    label = await modelText.getProperty('textContent');
                    model = await label.jsonValue();
                    model = model.substr(model.search('G') - 4, 150);
                    console.log('16 check 3');

                    // making Processor text look nice
                    processor = model.substr(model.search('G') - 4, 22);
                    if (processor.search('Chip') == -1) {
                        processor = processor.replace('Intel Core ', '');
                    }
                    // making RAM text look nice
                    RAM = parseInt(model.substr(model.search('RAM') - 6, 2));
                    // making SSD text look nice
                    if (model.search('PCIe') != -1) {
                        if (model.search('TB') != -1) {
                            SSD = model.substr(model.search('TB') - 2, 4);
                        }
                        else {
                            SSD = model.substr(model.search('SSD') - 12, 6);
                        }
                    }
                    else {
                        if (model.search('TB') != -1) {
                            SSD = model.substr(model.search('TB') - 2, 4);
                        }
                        else {
                            SSD = model.substr(model.search('SSD') - 7, 6);
                        }
                    }

                    // excel styles
                    excelStyles(anzeigeCounter);

                    // pressing verkaufen button
                    [mainPageButton] = await page.$x('//*[@id="ry"]/body/main/div[1]/div[2]/div/div/div/div/div/div[' + anzeigeCounter + ']/a/div/div[4]/button/ng-switch/span');
                    await mainPageButton.evaluate(mainPageButton => mainPageButton.click());
                    await delay(defaultTime); // it is necessary
                    console.log('16 check 4');
                    // doing Zustand WieNeu survey
                    var questionDiv = 1;
                    for (; questionDiv < 5; questionDiv++) {
                        [survey] = await page.$x('//*[@id="grading-form"]/div[1]/ry-grading-questions/div[' + questionDiv + ']/div/ry-grading-radio/div[1]/div[1]/label');
                        await survey.evaluate(survey => survey.click()); //await delay(defaultTime); // not necessary
                    }

                    // getting the price value
                    await delay(defaultTime * 1.5); // it is necessary       --------------
                    [bestPrice] = await page.$x('//*[@id="grading-form"]/div[2]/ry-grading-info/div/div[2]/div[1]/div[2]/div[1]/p/span');
                    value = await bestPrice.getProperty('textContent');
                    priceWN = await value.jsonValue();
                    priceWN = parseInt(priceWN);

                    await page.goto(link);
                    await delay(defaultTime); //test OCT

                    //write to xlsx
                    worksheet.cell(anzeigeCounter + 1, 1).string(processor);
                    worksheet.cell(anzeigeCounter + 1, 2).number(RAM);
                    worksheet.cell(anzeigeCounter + 1, 3).string(SSD);

                    if (model.search(color[0]) != -1) {
                        worksheet.cell(anzeigeCounter + 1, 4).string(color[0]);
                    }
                    if (model.search(color[1]) != -1) {
                        worksheet.cell(anzeigeCounter + 1, 4).string(color[1]);
                    }
                    if (model.search(color[2]) != -1) {
                        worksheet.cell(anzeigeCounter + 1, 4).string(color[2]);
                    }
                    if (model.search(color[3]) != -1) {
                        worksheet.cell(anzeigeCounter + 1, 4).string(color[3]);
                    }

                    if (model.search(qwerty) != -1) {
                        worksheet.cell(anzeigeCounter + 1, 5).string(qwerty);
                    }

                    setupPrices(priceWN, anzeigeCounter);

                    console.log(anzeigeCounter, 'done');
                }
                else {
                    console.log(anzeigeCounter, 'KA');
                }
            }
        }
        catch {
            console.log('16 Pro - issue');
        }
        return { anzeigeCounter, link, titleanzeigeCounter, questionDiv };
    }

    async function new16() {            // page counter does not work, implemented new title parser
        
        var anzeigeCounter = 1;
        var link = constantLink + '-pro/16?f_prop_season=2021&page=' + pageNumber;
        try {
            await page.goto(link);
            await delay(defaultTime); //test OCT
            var pageNumber = 1;

            // check the number of the results
            [anzeigeNummer] = await page.$x('//*[@id="ry"]/body/main/div[1]/div[2]/div/div/div/div/h2/span[2]');
            nummer = await anzeigeNummer.getProperty('textContent');
            anzeigen = await nummer.jsonValue();
            anzeigen = parseInt(anzeigen.substr(1, 3));
            console.log('Total:', anzeigen); // need it here
            worksheet = workbook.addWorksheet('2021 16 Pro');
            if (isNaN(anzeigen) ? worksheet.cell(1, 12).string('unknown amount') : worksheet.cell(1, 12).string('Total: ' + anzeigen));

            // giving styles to sheet
            worksheet.column(1).setWidth(12);
            worksheet.column(2).setWidth(8);
            var titleanzeigeCounter = 1;
            while (titleanzeigeCounter <= titles.length) {
                worksheet.cell(1, titleanzeigeCounter).string(titles[titleanzeigeCounter - 1]).style(title);
                titleanzeigeCounter++;
            }

            for (; anzeigeCounter < anzeigen + 1; anzeigeCounter++) { //  for( ; anzeigeCounter < anzeigen + 1; anzeigeCounter++){
                // check if it is 'Kein Ankauf'
                [mainPageButton] = await page.$x('//*[@id="ry"]/body/main/div[1]/div[2]/div/div/div/div/div/div[' + anzeigeCounter + ']/a/div/div[4]/button/ng-switch/span');
                btnTxt = await mainPageButton.getProperty('textContent');
                btnTxtKA = await btnTxt.jsonValue();

                if (btnTxtKA != 'Kein Ankauf') {
                    // getting model text (DO NOT PRINT in the console)
                    [modelText] = await page.$x('//*[@id="ry"]/body/main/div[1]/div[2]/div/div/div/div/div/div[' + anzeigeCounter + ']/a/div/div[3]/text()');
                    label = await modelText.getProperty('textContent');
                    model = await label.jsonValue();
                    model = model.substr(model.search('G') - 4, 150);
                    console.log('16 check 3');

                    // making Processor text look nice// Apple MacBook Pro CTO mit Touch ID 16.2" (Liquid Retina XDR Display) 3.2 GHz M1 Max Chip (24-Core GPU) 32 GB RAM 4 TB SSD [Late 2021, englisches Tastaturlayout, QWERTY] space grau
                    if (model.includes('Chip')) {
                        //processor = processor.replace('Intel Core ', '');// wft kakoi intel core
                        processor = model.substr(model.search('G') + 3, 26);
                    }
                    else{
                        processor = model.substr(model.search('G') - 4, 22);
                    }
                    // making RAM text look nice
                    RAM = parseInt(model.substr(model.search('RAM') - 6, 2));
                    // making SSD text look nice
                    if (model.search('PCIe') != -1) {
                        if (model.search('TB') != -1) {
                            SSD = model.substr(model.search('TB') - 2, 4);
                        }
                        else {
                            SSD = model.substr(model.search('SSD') - 12, 6);
                        }
                    }
                    else {
                        if (model.search('TB') != -1) {
                            SSD = model.substr(model.search('TB') - 2, 4);
                        }
                        else {
                            SSD = model.substr(model.search('SSD') - 7, 6);
                        }
                    }

                    // excel styles
                    excelStyles(anzeigeCounter);

                    // pressing verkaufen button
                    [mainPageButton] = await page.$x('//*[@id="ry"]/body/main/div[1]/div[2]/div/div/div/div/div/div[' + anzeigeCounter + ']/a/div/div[4]/button/ng-switch/span');
                    await mainPageButton.evaluate(mainPageButton => mainPageButton.click());
                    await delay(defaultTime); // it is necessary
                    console.log('16 check 4');
                    // doing Zustand WieNeu survey
                    var questionDiv = 1;
                    for (; questionDiv < 5; questionDiv++) {
                        [survey] = await page.$x('//*[@id="grading-form"]/div[1]/ry-grading-questions/div[' + questionDiv + ']/div/ry-grading-radio/div[1]/div[1]/label');
                        await survey.evaluate(survey => survey.click()); //await delay(defaultTime); // not necessary
                    }

                    // getting the price value
                    await delay(defaultTime * 1.5); // it is necessary       --------------
                    [bestPrice] = await page.$x('//*[@id="grading-form"]/div[2]/ry-grading-info/div/div[2]/div[1]/div[2]/div[1]/p/span');
                    value = await bestPrice.getProperty('textContent');
                    priceWN = await value.jsonValue();
                    priceWN = parseInt(priceWN);

                    await page.goto(link);
                    await delay(defaultTime); //test OCT

                    //write to xlsx
                    worksheet.cell(anzeigeCounter + 1, 1).string(processor);
                    worksheet.cell(anzeigeCounter + 1, 2).number(RAM);
                    worksheet.cell(anzeigeCounter + 1, 3).string(SSD);

                    if (model.search(color[0]) != -1) {
                        worksheet.cell(anzeigeCounter + 1, 4).string(color[0]);
                    }
                    if (model.search(color[1]) != -1) {
                        worksheet.cell(anzeigeCounter + 1, 4).string(color[1]);
                    }
                    if (model.search(color[2]) != -1) {
                        worksheet.cell(anzeigeCounter + 1, 4).string(color[2]);
                    }
                    if (model.search(color[3]) != -1) {
                        worksheet.cell(anzeigeCounter + 1, 4).string(color[3]);
                    }

                    if (model.search(qwerty) != -1) {
                        worksheet.cell(anzeigeCounter + 1, 5).string(qwerty);
                    }

                    setupPrices(priceWN, anzeigeCounter);

                    console.log(anzeigeCounter, 'done');
                }
                else {
                    console.log(anzeigeCounter, 'KA');
                }
            }
        }
        catch {
            console.log('new 16 Pro - issue');
        }
        if (anzeigeCounter == 24) {
            link = link.replace('e=' + pageNumber, 'e=' + ++pageNumber);

            await page.goto(link);

            anzeigeCounter = 0;
        }
        return { anzeigeCounter, link, titleanzeigeCounter, questionDiv };
    }

    async function new14() {            // need to consider cores and page counter does not work
        var pageNumber = 1;
        var anzeigeCounter = 1;
        var link = constantLink + '-pro/14?page=1';
        try {
            await page.goto(link); //https://www.rebuy.de/verkaufen/notebooks/apple/macbook-pro/14?page=1
            await delay(defaultTime); //test OCT

            // проверять что вообще оно собирается выполнять if (все исключения)
            // check the number of the results
            [anzeigeNummer] = await page.$x('//*[@id="ry"]/body/main/div[1]/div[2]/div/div/div/div/h2/span[2]');
            nummer = await anzeigeNummer.getProperty('textContent');
            anzeigen = await nummer.jsonValue();
            anzeigen = parseInt(anzeigen.substr(1, 3));
            console.log('Total:', anzeigen); // need it here
            worksheet = workbook.addWorksheet('2021 14 Pro');
            if (isNaN(anzeigen) ? worksheet.cell(1, 12).string('unknown amount') : worksheet.cell(1, 12).string('Total: ' + anzeigen))
                ;

            // giving styles to sheet
            worksheet.column(1).setWidth(12);
            worksheet.column(2).setWidth(8);
            var titleanzeigeCounter = 1;
            while (titleanzeigeCounter <= titles.length) {
                worksheet.cell(1, titleanzeigeCounter).string(titles[titleanzeigeCounter - 1]).style(title);
                titleanzeigeCounter++;
            }

            for (; anzeigeCounter < anzeigen + 1; anzeigeCounter++) {
                // check if it is 'Kein Ankauf'
                [mainPageButton] = await page.$x('//*[@id="ry"]/body/main/div[1]/div[2]/div/div/div/div/div/div[' + anzeigeCounter + ']/a/div/div[4]/button/ng-switch/span');
                btnTxt = await mainPageButton.getProperty('textContent');
                btnTxtKA = await btnTxt.jsonValue();

                if (btnTxtKA != 'Kein Ankauf') {
                    // getting model text (DO NOT PRINT in the console)
                    [modelText] = await page.$x('//*[@id="ry"]/body/main/div[1]/div[2]/div/div/div/div/div/div[' + anzeigeCounter + ']/a/div/div[3]/text()');
                    label = await modelText.getProperty('textContent');
                    model = await label.jsonValue();
                    model = model.substr(model.search('G') - 4, 150);

                    // making Processor text look nice
                    processor = model.substr(model.search('G') - 4, 22);
                    if (processor.search('Chip') == -1) {
                        processor = processor.replace('Intel Core ', '');
                        // if(processor.search(')') != -1){
                        //     processor = processor.replace(')', '');    // not now
                        // }
                    }
                    // making RAM text look nice
                    RAM = parseInt(model.substr(model.search('RAM') - 6, 2));
                    // making SSD text look nice
                    if (model.search('PCIe') != -1) {
                        if (model.search('TB') != -1) {
                            SSD = model.substr(model.search('TB') - 2, 4);
                        }
                        else {
                            SSD = model.substr(model.search('SSD') - 12, 6);
                        }
                    }
                    else {
                        if (model.search('TB') != -1) {
                            SSD = model.substr(model.search('TB') - 2, 4);
                        }
                        else {
                            SSD = model.substr(model.search('SSD') - 7, 6);
                        }
                    }

                    // excel styles
                    excelStyles(anzeigeCounter);

                    // pressing verkaufen button
                    [mainPageButton] = await page.$x('//*[@id="ry"]/body/main/div[1]/div[2]/div/div/div/div/div/div[' + anzeigeCounter + ']/a/div/div[4]/button/ng-switch/span');
                    await mainPageButton.evaluate(mainPageButton => mainPageButton.click());
                    await delay(defaultTime); // it is necessary
                    // doing Zustand WieNeu survey
                    var questionDiv = 1;
                    for (; questionDiv < 5; questionDiv++) {
                        [survey] = await page.$x('//*[@id="grading-form"]/div[1]/ry-grading-questions/div[' + questionDiv + ']/div/ry-grading-radio/div[1]/div[1]/label');
                        await survey.evaluate(survey => survey.click()); //await delay(defaultTime); // not necessary
                    }

                    // getting the price value
                    await delay(defaultTime * 1.5); // it is necessary       --------------
                    [bestPrice] = await page.$x('//*[@id="grading-form"]/div[2]/ry-grading-info/div/div[2]/div[1]/div[2]/div[1]/p/span');
                    value = await bestPrice.getProperty('textContent');
                    priceWN = await value.jsonValue();
                    priceWN = parseInt(priceWN);

                    await page.goto(link);
                    await delay(defaultTime); //test OCT


                    //write to xlsx
                    worksheet.cell(anzeigeCounter + 1, 1).string(processor);
                    worksheet.cell(anzeigeCounter + 1, 2).number(RAM);
                    worksheet.cell(anzeigeCounter + 1, 3).string(SSD);

                    if (model.search(color[0]) != -1) {
                        worksheet.cell(anzeigeCounter + 1, 4).string(color[0]);
                    }
                    if (model.search(color[1]) != -1) {
                        worksheet.cell(anzeigeCounter + 1, 4).string(color[1]);
                    }
                    if (model.search(color[2]) != -1) {
                        worksheet.cell(anzeigeCounter + 1, 4).string(color[2]);
                    }
                    if (model.search(color[3]) != -1) {
                        worksheet.cell(anzeigeCounter + 1, 4).string(color[3]);
                    }

                    if (model.search(qwerty) != -1) {
                        worksheet.cell(anzeigeCounter + 1, 5).string(qwerty);
                    }

                    setupPrices(priceWN, anzeigeCounter);

                    console.log(anzeigeCounter, 'done');
                }
                else {
                    console.log(anzeigeCounter, 'KA');
                }
            }
        }
        catch {
            console.log('14 Pro - issue in loop');
        }
        if (anzeigeCounter == 24) {
            link = link.replace('e=' + pageNumber, 'e=' + ++pageNumber);

            await page.goto(link);

            anzeigeCounter = 0;
        }
        return { anzeigeCounter, link, titleanzeigeCounter, questionDiv };
    }

    function finish() {
        workbook.write(today + '.xlsx'); // create output folder
        console.log('\n--- file created ---\n');
        browser.close();
    }
    finish();
}
//        ˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙
scrapeMacs();

//TODO:
// Починить ошибку 500 (catch) и специфик модел (красная надпись на сайте)