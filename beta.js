const puppeteer = require('puppeteer');
const excel = require('excel4node');
const prompt = require('prompt-sync')();
// creating date
var today = new Date();
var dd = String(today.getDate()).padStart(2, '0');
var mm = String(today.getMonth() + 1).padStart(2, '0'); //January is 0!
var yyyy = today.getFullYear();
today = dd + '-' + mm + '-' + yyyy;
// main variables
const constantLink = 'https://www.rebuy.de/verkaufen/apple/notebooks/macbook'; // cmd+K cmd+0
var workbook = new excel.Workbook();
var defaultTime = 3000;              // 550 w/o doubler // until 14.09 == 750
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
    'Buy',
    'Sell', 
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
function excelStyles(worksheetCounter){
    worksheet.cell(worksheetCounter + 1, 7).style(bgGreen);
    worksheet.cell(worksheetCounter + 1, 9).style(red);
    worksheet.cell(worksheetCounter + 1, 8).style(yellow);
    worksheet.cell(worksheetCounter + 1, 6).style(simple);
    worksheet.cell(worksheetCounter + 1, 10).style(simple);
}
// other functions
function delay(time) {
    return new Promise(function(resolve) { setTimeout(resolve, time) });
}
function setupPrices(priceWN, worksheetCounter){
    if (isNaN(priceWN) ? worksheet.cell(worksheetCounter + 1, 6).number(0) : worksheet.cell(worksheetCounter + 1, 6).number(parseInt(priceWN * 0.863)));
    if (isNaN(priceWN) ? worksheet.cell(worksheetCounter + 1, 7).number(0) : worksheet.cell(worksheetCounter + 1, 7).number(priceWN));
    if (isNaN(priceWN) ? worksheet.cell(worksheetCounter + 1, 8).number(0) : worksheet.cell(worksheetCounter + 1, 8).number(parseInt(priceWN * 0.908)));
    if (isNaN(priceWN) ? worksheet.cell(worksheetCounter + 1, 9).number(0) : worksheet.cell(worksheetCounter + 1, 9).number(parseInt(priceWN * 0.818)));
    if (isNaN(priceWN) ? worksheet.cell(worksheetCounter + 1, 10).number(0) : worksheet.cell(worksheetCounter + 1, 10).number(parseInt(priceWN - ((priceWN * 0.908 + priceWN * 0.818)/2))));
}
var scanCounter = 0;
async function launcher(call) { // 30.07.22
    var scanModel = ["15 Pro", "13 Pro", "16 Pro"];
    
    switch(scanCounter){
        case 0:
            var answer = prompt('Collect ' + scanModel[0] + '?(Y/n)-> ');
            break;
            
        case 1:
            var answer = prompt('Collect ' + scanModel[1] + '?(Y/n)-> ');
            break;
        
        case 2:
            var answer = prompt('Collect ' + scanModel[2] + '?(Y/n)-> ');
            break;
    }
    //var answer = prompt('Scan? (Y/n)-> ');
    if (answer == 1 || answer == 'y' || answer == 'Y') {
        var { anzeigeCounter, link, titleanzeigeCounter, questionDiv } = await call();
    }
    else {
        console.log('this model ignored');
    }
    scanCounter++;
}
async function parseModelName(modelInfo, worksheetCounter) {

    label = await modelInfo.getProperty('textContent');
    model = await label.jsonValue();
    model = model.substr(model.search('G') - 4, 150);

    // making Processor text look nice
    processor = model.substr(model.search('G') - 4, 22);
    if (processor.search('Chip') == -1) {
        processor = processor.replace('Intel Core ', '');
    }
    else{ // 30.07.22
        processor = processor.substr(7, 15);
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

    // write to xlsx
    worksheet.cell(worksheetCounter + 1, 1).string(processor);
    worksheet.cell(worksheetCounter + 1, 2).number(RAM);
    worksheet.cell(worksheetCounter + 1, 3).string(SSD);

    // this is colors
    if (model.search(color[0]) != -1) {
        worksheet.cell(worksheetCounter + 1, 4).string(color[0]);
    }
    if (model.search(color[1]) != -1) {
        worksheet.cell(worksheetCounter + 1, 4).string(color[1]);
    }
    if (model.search(color[2]) != -1) {
        worksheet.cell(worksheetCounter + 1, 4).string(color[2]);
    }
    if (model.search(color[3]) != -1) {
        worksheet.cell(worksheetCounter + 1, 4).string(color[3]);
    }

    // this is keyboard
    if (model.search(qwerty) != -1) {
        worksheet.cell(worksheetCounter + 1, 5).string(qwerty);
    }
}

// ************************************** MAIN ******************************************************
async function scrapeMacs(){ // next page does not work // 16 old cringe // проблема былоа что оно переписывало результаты заново после 24 был стремный каунтер типо анцайгекаунтер

    /*
    +   What is this doing? create RBY.nl NL!
    +   2018-2019 Pro 15
    +   2020      Pro 13
    +   2019      Pro 16
    +   //2021      Pro 16
    +   //2021      Pro 14
    +  (2020      Air)
    */

    const browser = await puppeteer.launch({headless: false, slowMo: 500}); // [_][_][_][_][_][_][_][_] );//
    const page = await browser.newPage();
    console.clear();
    console.log('- - - - - NEW SCAN ' + today + ' - - - - -')

    await launcher(old15);
    await launcher(old13);
    await launcher(old16);

    async function old15() {
        // input startYear
        var year = prompt('Start scan from which Year? -> ');
        if (year < 2018){
            year = 2018;
        }
        while (year < 2020) {
            try {
                var anzeigeCounter = 1;
                var pageNumber = 1;
                var link = constantLink + '-pro?f_prop_season=' + year + '&f_prop_display_size=15,4%20Zoll&page=' + pageNumber;
                await page.goto(link);
                await delay(defaultTime);
                worksheet = workbook.addWorksheet(year + ' 15 Pro');

                // make tables look better
                formatTables();

                var titleanzeigeCounter = 1;
                while (titleanzeigeCounter <= titles.length) {
                    worksheet.cell(1, titleanzeigeCounter).string(titles[titleanzeigeCounter - 1]).style(title);
                    titleanzeigeCounter++;
                }
                console.log('check1');

                // check the number of the results
                [anzeigeNummer] = await page.$x('//*[@id="ry"]/body/main/div[1]/div[2]/div/div/div/div/h2/span[2]');
                console.log('check1.5');
                nummer = await anzeigeNummer.getProperty('textContent');
                anzeigen = await nummer.jsonValue();
                console.log('check2');
                anzeigen = parseInt(anzeigen.substr(1, 2)); // here we understand how much available
                console.log(year, 'Total:', anzeigen); // need it here

                if (isNaN(anzeigen) ? worksheet.cell(1, 12).string('unknown amount') : worksheet.cell(1, 12).string('Total: ' + anzeigen));
                if (anzeigen > 24 ? anzeigen = 24 : console.log('one page'));
                worksheetCounter = 1;

                for (; anzeigeCounter < anzeigen + 1; anzeigeCounter++) { // here you can limit number of anzeigen 'for( ; anzeigeCounter < anzeigen + 1; anzeigeCounter++){
                    // check if it is 'Kein Ankauf'
                    [mainPageButton] = await page.$x('//*[@id="ry"]/body/main/div[1]/div[2]/div/div/div/div/div[1]/div[' + anzeigeCounter + ']/a/div/div[5]/button/ng-switch/span');
                    btnTxt = await mainPageButton.getProperty('textContent');
                    btnTxtKA = await btnTxt.jsonValue();
                    try {
                        if (btnTxtKA != 'Kein Ankauf') {
                            // getting model text (DO NOT PRINT in the console)
                            [modelText] = await page.$x('//*[@id="ry"]/body/main/div[1]/div[2]/div/div/div/div/div[1]/div[' + anzeigeCounter + ']/a/div/div[3]/text()');

                            // parse model name: Processor, RAM, SSD, Color, QWERTY
                            await parseModelName(modelText, worksheetCounter);

                            // pressing verkaufen button
                            [mainPageButton] = await page.$x('//*[@id="ry"]/body/main/div[1]/div[2]/div/div/div/div/div[1]/div[' + anzeigeCounter + ']/a/div/div[5]/button/ng-switch/span');
                            await mainPageButton.evaluate(mainPageButton => mainPageButton.click());
                            await delay(defaultTime * 1.5); // it is necessary

                            // doing Zustand WieNeu survey
                            var questionDiv = 1;
                            for (; questionDiv < 8; questionDiv++) { // 29.03.22 // 27.05.22
                                if (questionDiv < 6){
                                    [survey] = await page.$x('//*[@id="grading-form"]/div[1]/ry-grading-questions/div[' + questionDiv + ']/div/ry-grading-radio/div[1]/div[1]/label');
                                }
                                else{
                                    [survey] = await page.$x('//*[@id="grading-form"]/div[1]/ry-grading-questions/div[' + questionDiv + ']/div/ry-grading-radio/div[1]/div/label');
                                }
                                await survey.evaluate(survey => survey.click()); //await delay(defaultTime); // not necessary
                            }

                            // getting the price value
                            await delay(defaultTime * 1.5); // it is necessary       --------------
                            [bestPrice] = await page.$x('//*[@id="grading-form"]/div[2]/ry-grading-info/div/div[2]/div[1]/div[2]/div[1]/p/span'); //try get price if 0 wait time
                            // if(bestPrice.getProperty('textContent') != 0){
                            //     // proceed
                            // }
                            value = await bestPrice.getProperty('textContent');
                            priceWN = await value.jsonValue();
                            priceWN = parseInt(priceWN);

                            // return to the initial page with all results
                            await page.goto(link);

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
                console.log('this model is not available or website is down');
            }
            year++;
        }
        return { anzeigeCounter, link, titleanzeigeCounter, questionDiv };
    }
    
    async function old13() {
        try {
            var anzeigeCounter = 1; // 100% можно переместить свитч idk what I meant
            var pageNumber = 1;
            var code = "-pro?f_prop_season=2020&page=";

            worksheet = workbook.addWorksheet('2020 13 Pro');

            var link = constantLink + '-pro?f_prop_season=2020&page=' + pageNumber;
            await page.goto(link);

            // make tables look better
            formatTables();

            var titleanzeigeCounter = 1;
            while (titleanzeigeCounter <= titles.length) {
                worksheet.cell(1, titleanzeigeCounter).string(titles[titleanzeigeCounter - 1]).style(title);
                titleanzeigeCounter++;
            }

            // check the number of the results
            [anzeigeNummer] = await page.$x('//*[@id="ry"]/body/main/div[1]/div[2]/div/div/div/div/h2/span[2]');//*[@id="ry"]/body/rebuy-app/div/div/ry-product-list/ry-product-list-electronic/div[1]/div/div/h2
            nummer = await anzeigeNummer.getProperty('textContent');
            anzeigen = await nummer.jsonValue();
            anzeigen = parseInt(anzeigen.substr(1, 2)); // here we understand how much available
            console.log('Pro 13 Total:', anzeigen); // need it here

            if (isNaN(anzeigen) ? worksheet.cell(1, 12).string('unknown amount') : worksheet.cell(1, 12).string('Total: ' + anzeigen));
            if (anzeigen > 24 ? anzeigen = 24 : console.log('one page'));
            worksheetCounter = 1;

            for (; anzeigeCounter < anzeigen + 1; anzeigeCounter++) { // here you can limit number of anzeigen 'for( ; anzeigeCounter < anzeigen + 1; anzeigeCounter++){
                // check if it is 'Kein Ankauf'
                [mainPageButton] = await page.$x('//*[@id="ry"]/body/main/div[1]/div[2]/div/div/div/div/div[1]/div[' + anzeigeCounter + ']/a/div/div[5]/button/ng-switch/span');
                btnTxt = await mainPageButton.getProperty('textContent');
                btnTxtKA = await btnTxt.jsonValue();
                try {
                    if (btnTxtKA != 'Kein Ankauf') {
                        // getting model text (DO NOT PRINT in the console)
                        [modelText] = await page.$x('//*[@id="ry"]/body/main/div[1]/div[2]/div/div/div/div/div/div[' + anzeigeCounter + ']/a/div/div[3]/text()');

                        // parse model name: Processor, RAM, SSD, Color, QWERTY
                        await parseModelName(modelText, worksheetCounter);

                        // excel styles
                        excelStyles(worksheetCounter);

                        // pressing verkaufen button
                        [mainPageButton] = await page.$x('//*[@id="ry"]/body/main/div[1]/div[2]/div/div/div/div/div[1]/div[' + anzeigeCounter + ']/a/div/div[5]/button/ng-switch/span');
                        await mainPageButton.evaluate(mainPageButton => mainPageButton.click());
                        await delay(defaultTime * 1.5); // it is necessary

                        // doing Zustand WieNeu survey
                        var questionDiv = 1;
                        for (; questionDiv < 8; questionDiv++) { // 29.03.22
                            if (questionDiv < 6){
                                [survey] = await page.$x('//*[@id="grading-form"]/div[1]/ry-grading-questions/div[' + questionDiv + ']/div/ry-grading-radio/div[1]/div[1]/label');
                            }
                            else{
                                [survey] = await page.$x('//*[@id="grading-form"]/div[1]/ry-grading-questions/div[' + questionDiv + ']/div/ry-grading-radio/div[1]/div/label');
                            }
                            await survey.evaluate(survey => survey.click()); //await delay(defaultTime); // not necessary
                        }

                        // getting the price value
                        await delay(defaultTime * 1.5); // it is necessary       --------------
                        [bestPrice] = await page.$x('//*[@id="grading-form"]/div[2]/ry-grading-info/div/div[2]/div[1]/div[2]/div[1]/p/span'); //try get price if 0 wait time
                        value = await bestPrice.getProperty('textContent');
                        priceWN = await value.jsonValue();
                        priceWN = parseInt(priceWN);

                        // return to the initial page with all results
                        await page.goto(link); // check if it is needed

                        setupPrices(priceWN, worksheetCounter);

                        console.log(worksheetCounter, 'done');
                    }
                    else {
                        console.log(worksheetCounter, 'KA');
                    }
                    worksheetCounter++;
                }
                catch {
                    console.log('specific product card failed');
                }
                if (anzeigeCounter == 24) {
                    link = link.replace('e=' + pageNumber, 'e=' + ++pageNumber);

                    await page.goto(link);

                    anzeigeCounter = 0;
                }
            }
        }
        catch {
            //console.log(link);
            console.log('this model is not available or website is down');
        }
        return { anzeigeCounter, link, titleanzeigeCounter, questionDiv };
    }
    
    async function old16() {
        var link = constantLink + '-pro?f_prop_display_size=16%20Zoll';
        try {
            var anzeigeCounter = 1;

            await page.goto(link);
            await delay(defaultTime);

            // check the number of the results
            [anzeigeNummer] = await page.$x('//*[@id="ry"]/body/main/div[1]/div[2]/div/div/div/div/h2/span[2]');
            nummer = await anzeigeNummer.getProperty('textContent');
            anzeigen = await nummer.jsonValue();
            anzeigen = parseInt(anzeigen.substr(1, 3));
            console.log('16 Pro Total:', anzeigen); // need it here
            worksheet = workbook.addWorksheet('2019 16 Pro');
            if (isNaN(anzeigen) ? worksheet.cell(1, 12).string('unknown amount') : worksheet.cell(1, 12).string('Total: ' + anzeigen));

            // make tables look better
            formatTables();

            var titleanzeigeCounter = 1;
            while (titleanzeigeCounter <= titles.length) {
                worksheet.cell(1, titleanzeigeCounter).string(titles[titleanzeigeCounter - 1]).style(title);
                titleanzeigeCounter++;
            }

            for (; anzeigeCounter < anzeigen + 1; anzeigeCounter++) { //  for( ; anzeigeCounter < anzeigen + 1; anzeigeCounter++){
                // check if it is 'Kein Ankauf'
                console.log('1 test');
                [mainPageButton] = await page.$x('//*[@id="ry"]/body/main/div[1]/div[2]/div/div/div/div/div[1]/div[' + anzeigeCounter + ']/a/div/div[5]/button/ng-switch/span');
                btnTxt = await mainPageButton.getProperty('textContent');
                btnTxtKA = await btnTxt.jsonValue();

                if (btnTxtKA != 'Kein Ankauf') {
                    // getting model text (DO NOT PRINT in the console)
                    [modelText] = await page.$x('//*[@id="ry"]/body/main/div[1]/div[2]/div/div/div/div/div[1]/div[' + anzeigeCounter + ']/a/div/div[3]/text()');
                    
                    // parse model name: Processor, RAM, SSD, Color, QWERTY
                    await parseModelName(modelText, worksheetCounter);
                    console.log('2 test');
                    // pressing verkaufen button
                    [mainPageButton] = await page.$x('//*[@id="ry"]/body/main/div[1]/div[2]/div/div/div/div/div[1]/div[' + anzeigeCounter + ']/a/div/div[5]/button/ng-switch/span');
                    await mainPageButton.evaluate(mainPageButton => mainPageButton.click());
                    await delay(defaultTime * 1.5); // it is necessary
                    
                    // doing Zustand WieNeu survey
                    var questionDiv = 1;
                    for (; questionDiv < 8; questionDiv++) { // 29.03.22
                        if (questionDiv < 6){
                            [survey] = await page.$x('//*[@id="grading-form"]/div[1]/ry-grading-questions/div[' + questionDiv + ']/div/ry-grading-radio/div[1]/div[1]/label');
                        }
                        else{
                            [survey] = await page.$x('//*[@id="grading-form"]/div[1]/ry-grading-questions/div[' + questionDiv + ']/div/ry-grading-radio/div[1]/div/label');
                        }
                        await survey.evaluate(survey => survey.click()); //await delay(defaultTime); // not necessary
                    }

                    // getting the price value
                    await delay(defaultTime * 1.5); // it is necessary       --------------
                    [bestPrice] = await page.$x('//*[@id="grading-form"]/div[2]/ry-grading-info/div/div[2]/div[1]/div[2]/div[1]/p/span');
                    value = await bestPrice.getProperty('textContent');
                    priceWN = await value.jsonValue();
                    priceWN = parseInt(priceWN);

                    await page.goto(link);
                    //wait delay(defaultTime);

                    setupPrices(priceWN, worksheetCounter);

                    console.log(anzeigeCounter, 'done');
                }
                else {
                    console.log(anzeigeCounter, 'KA');
                }
            }
        }
        catch {
            console.log('16 Pro - issue/specific product card failed');
        }
        return { anzeigeCounter, link, titleanzeigeCounter, questionDiv };
    }

    async function new16() {  // outdated link // outdated multiple // need to consider cores and page counter does not work, >>> implemented new title parser
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
                    for (; questionDiv < 7; questionDiv++) { // 29.03.22
                        if (questionDiv < 6){
                            [survey] = await page.$x('//*[@id="grading-form"]/div[1]/ry-grading-questions/div[' + questionDiv + ']/div/ry-grading-radio/div[1]/div[1]/label');
                        }
                        else{
                            [survey] = await page.$x('//*[@id="grading-form"]/div[1]/ry-grading-questions/div[' + questionDiv + ']/div/ry-grading-radio/div[1]/div/label');
                        }
                        await survey.evaluate(survey => survey.click());
                    }

                    // getting the price value
                    await delay(defaultTime * 1.5); // it is necessary       --------------
                    [bestPrice] = await page.$x('//*[@id="grading-form"]/div[2]/ry-grading-info/div/div[2]/div[1]/div[2]/div[1]/p/span');
                    value = await bestPrice.getProperty('textContent');
                    priceWN = await value.jsonValue();
                    priceWN = parseInt(priceWN);

                    await page.goto(link);
                    //await delay(defaultTime); //test OCT

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

                    setupPrices(priceWN, worksheetCounter);

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

    function finish() {
        workbook.write(today + '.xlsx'); // create output folder
        console.log('\n--- file created ---\n');
        browser.close();
    }

    finish();
}
// ************************************** MAIN ******************************************************
scrapeMacs();



//TODO:
// сравнить сколько общего у всех типов, 15 13 16 16
// new 16
// 14 is just duplicate of 16 once its finished
// я сейчас разделю на функции и усовершенстую так, что потом вернуть к одной функции в которую буду передавать ссылки и определять сценарий исходя из этого