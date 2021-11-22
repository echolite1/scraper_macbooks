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
const constantLink = 'https://www.rebuy.de/verkaufen/apple/notebooks/macbook';
var modelSwitcher = 1;                     // make 0 for 12"
var maxYear = yyyy;
var defaultTime = 750;                     // 550 w/o doubler

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
//        ˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙
async function scrapeMacs(){
    // launching pseudo-browser
    const browser = await puppeteer.launch();//{headless: false, slowMo: 100});
    const page = await browser.newPage();
    console.clear();

    while (modelSwitcher < 3){                                                      // пока что 12 и Эир не нужны
        while(year < maxYear){
            try{
                var counter = 1;
                var macModel = [
                    '?f_prop_season=' + year,
                    '-pro/15-4?f_prop_season=' + year, 
                    '-pro/13-3?f_prop_season=' + year,
                    '-air/13-3?f_prop_season=' + year
                ];
                var link = constantLink + macModel[modelSwitcher];

                await page.goto(link);

                // ignore creating of non-existing models and creating worksheets
                switch(modelSwitcher){
                    case 0:
                        if(link.search('2018') != -1 || link.search('2019') != -1 || link.search('2020')){
                            console.log('2018/-19/-20 12 inch -');//, link);
                        }
                        else{
                            worksheet = workbook.addWorksheet(year + ' 12 inch');
                        }
                        break;
                    case 1:
                        if(link.search('2020') != -1 || link.search('2015') != -1){
                            console.log('2015/-20 15 inch -');//, link);
                            // year++;
                            // link = constantLink + macModel[modelSwitcher];
                            // await page.goto(link);
                        }
                        else{
                            worksheet = workbook.addWorksheet(year + ' 15 Pro');
                        }
                        break;
                    case 2:
                        if(link.search('2015') != -1){
                            console.log('2015 13 inch -', link);
                        }
                        else{
                            worksheet = workbook.addWorksheet(year + ' 13 Pro');
                        }
                        break;
                    case 3:
                        if(link.search('2016') != -1 || link.search('2017') != -1){
                            console.log('2016/17 Air -');//, link);
                        }
                        else{
                            worksheet = workbook.addWorksheet(year + ' Air');
                        }
                        break;
                    default:
                        console.log('switch error');
                }
                worksheet.column(1).setWidth(12);
                worksheet.column(2).setWidth(8);
                worksheet.column(7).setWidth(9);
                var titleCounter = 1;
                while(titleCounter <= titles.length){
                    worksheet.cell(1, titleCounter).string(titles[titleCounter - 1]).style(title);
                    titleCounter++;
                }

                // проверять что вообще оно собирается выполнять if (все исключения)

                // check the number of the results
                [anzeigeNummer] = await page.$x('//*[@id="ry"]/body/main/div[1]/div[2]/div/div/div/div/h2/span[2]');
                nummer = await anzeigeNummer.getProperty('textContent');
                anzeigen = await nummer.jsonValue();
                anzeigen = parseInt(anzeigen.substr(1, 2)); 
                console.log(year, 'Total:', anzeigen); // need it here

                if (isNaN(anzeigen) ? worksheet.cell(1, 12).string('unknown amount') : worksheet.cell(1, 12).string('Total: ' + anzeigen));
                if (anzeigen > 24 ? anzeigen = 24 : console.log(anzeigen));
                
                for( ; counter < anzeigen + 1; counter++){
                    // check if it is 'Kein Ankauf'
                    [mainPageButton] = await page.$x('//*[@id="ry"]/body/main/div[1]/div[2]/div/div/div/div/div/div[' + counter + ']/a/div/div[4]/button/ng-switch/span');
                    btnTxt = await mainPageButton.getProperty('textContent');
                    btnTxtKA = await btnTxt.jsonValue();
                    try{
                        if(btnTxtKA != 'Kein Ankauf')
                    {
                        // getting model text (DO NOT PRINT in the console)
                        [modelText] = await page.$x('//*[@id="ry"]/body/main/div[1]/div[2]/div/div/div/div/div/div[' + counter + ']/a/div/div[3]/text()');
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
                        worksheet.cell(counter + 1, 7).style(bgGreen);
                        worksheet.cell(counter + 1, 9).style(red);
                        worksheet.cell(counter + 1, 8).style(yellow);
                        worksheet.cell(counter + 1, 6).style(simple);
                        worksheet.cell(counter + 1, 10).style(simple);

                        // pressing verkaufen button
                        [mainPageButton] = await page.$x('//*[@id="ry"]/body/main/div[1]/div[2]/div/div/div/div/div/div[' + counter + ']/a/div/div[4]/button/ng-switch/span');
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
                        
                        // perform change WN -> SG
                        [changeZustand] = await page.$x('//*[@id="grading-form"]/div[1]/ry-grading-overview/div/span[2]');
                        await changeZustand.evaluate( changeZustand => changeZustand.click() );

                        // perform click SG
                        [surveySG] = await page.$x('//*[@id="grading-form"]/div[1]/ry-grading-questions/div[2]/div/ry-grading-radio/div[1]/div[2]/label');
                        await surveySG.evaluate( surveySG => surveySG.click() );

                        // get the updated price (SG)
                        await delay(defaultTime * 1.5); // it is necessary         ----------------
                        [bestPrice] = await page.$x('//*[@id="grading-form"]/div[2]/ry-grading-info/div/div[2]/div[1]/div[2]/div[1]/p/span');
                        value = await bestPrice.getProperty('textContent');
                        priceSG = await value.jsonValue();
                        priceSG = parseInt(priceSG);

                        // perform change SG -> Gut
                        [changeZustand] = await page.$x('//*[@id="grading-form"]/div[1]/ry-grading-overview/div/span[2]');
                        await changeZustand.evaluate( changeZustand => changeZustand.click() );

                        // perform click Gut
                        [surveyBP] = await page.$x('//*[@id="grading-form"]/div[1]/ry-grading-questions/div[2]/div/ry-grading-radio/div[1]/div[3]/label');
                        await surveyBP.evaluate( surveyBP => surveyBP.click() );

                        // get the updated price (Gut)
                        await delay(defaultTime * 1.5); // it is necessary         ----------------
                        [bestPrice] = await page.$x('//*[@id="grading-form"]/div[2]/ry-grading-info/div/div[2]/div[1]/div[2]/div[1]/p/span');
                        value = await bestPrice.getProperty('textContent');
                        priceBP = await value.jsonValue();
                        priceBP = parseInt(priceBP);

                        // return to the initial page with all results
                        await page.goto(link); // check if it is needed
                        
                        //write to xlsx
                        worksheet.cell(counter + 1, 1).string(processor);
                        worksheet.cell(counter + 1, 2).number(RAM);
                        worksheet.cell(counter + 1, 3).string(SSD);

                        if (model.search(color[0]) != -1){
                            worksheet.cell(counter + 1, 4).string(color[0]);
                        } 
                        else if (model.search(color[1]) != -1){
                            worksheet.cell(counter + 1, 4).string(color[1]);
                        }
                        else if (model.search(color[2]) != -1){
                            worksheet.cell(counter + 1, 4).string(color[2]);
                        }
                        else if (model.search(color[3]) != -1){
                            worksheet.cell(counter + 1, 4).string(color[3]);
                        }
                        else {
                            worksheet.cell(counter + 1, 4).string(color[1]);
                        }

                        if (model.search(qwerty) != -1){ worksheet.cell(counter + 1, 5).string(qwerty); }

                        if (isNaN(priceBP) ? worksheet.cell(counter + 1, 9).number(0) : worksheet.cell(counter + 1, 9).number(priceBP));
                        if (isNaN(priceWN) ? worksheet.cell(counter + 1, 7).number(0) : worksheet.cell(counter + 1, 7).number(priceWN));
                        if (isNaN(priceSG) ? worksheet.cell(counter + 1, 8).number(0) : worksheet.cell(counter + 1, 8).number(priceSG));
                        if (isNaN(priceSG) ? worksheet.cell(counter + 1, 6).number(0) : worksheet.cell(counter + 1, 6).number(parseInt((priceSG + priceBP)/2)));
                        if (isNaN(priceWN) ? worksheet.cell(counter + 1, 10).number(0) : worksheet.cell(counter + 1, 10).number(parseInt(priceWN - ((priceSG + priceBP)/2))));
                        
                        console.log(counter, 'done');
                    }
                    else {
                        //console.log(counter, 'KA');
                    }
                    }
                    catch{
                        console.log('specific mac failed');
                    }
                    
                }
            } 
            catch{
                console.log(model + ' is not available or other issue in loop');
            }
            year++;
        }
        modelSwitcher++;
        year = inputYear;
    }
    //        ˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙
    var counter = 1;
    var link = constantLink + '-pro/16';
console.log('16 check 0');
    try{
        await page.goto(link);
        
console.log('16 check 1');

        // проверять что вообще оно собирается выполнять if (все исключения)

        // check the number of the results
        [anzeigeNummer] = await page.$x('//*[@id="ry"]/body/main/div[1]/div[2]/div/div/div/div/h2/span[2]');
        nummer = await anzeigeNummer.getProperty('textContent');
        anzeigen = await nummer.jsonValue();
        anzeigen = parseInt(anzeigen.substr(1, 2)); 
        console.log(year, 'Total:', anzeigen); // need it here
        worksheet = workbook.addWorksheet('2019 16 Pro');
        if (isNaN(anzeigen) ? worksheet.cell(1, 8).string('unknown amount') : worksheet.cell(1, 10).string('Total: ' + anzeigen));

        // giving styles to sheet
        worksheet.column(1).setWidth(12);
        worksheet.column(2).setWidth(8);
        var titleCounter = 1;
        while(titleCounter <= titles.length){
            worksheet.cell(1, titleCounter).string(titles[titleCounter - 1]).style(title);
            titleCounter++;
        }
console.log('16 check 2');

        for( ; counter < anzeigen + 1; counter++){
            // check if it is 'Kein Ankauf'
            [mainPageButton] = await page.$x('//*[@id="ry"]/body/main/div[1]/div[2]/div/div/div/div/div/div[' + counter + ']/a/div/div[4]/button/ng-switch/span');
            btnTxt = await mainPageButton.getProperty('textContent');
            btnTxtKA = await btnTxt.jsonValue();

            if(btnTxtKA != 'Kein Ankauf')
            {
                // getting model text (DO NOT PRINT in the console)
                [modelText] = await page.$x('//*[@id="ry"]/body/main/div[1]/div[2]/div/div/div/div/div/div[' + counter + ']/a/div/div[3]/text()');
                label = await modelText.getProperty('textContent');
                model = await label.jsonValue();
                model = model.substr(model.search('G') - 4, 150);
console.log('16 check 3');

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
                worksheet.cell(counter + 1, 7).style(bgGreen);
                worksheet.cell(counter + 1, 9).style(red);
                worksheet.cell(counter + 1, 8).style(yellow);
                worksheet.cell(counter + 1, 6).style(simple);
                worksheet.cell(counter + 1, 10).style(simple);

                // pressing verkaufen button
                [mainPageButton] = await page.$x('//*[@id="ry"]/body/main/div[1]/div[2]/div/div/div/div/div/div[' + counter + ']/a/div/div[4]/button/ng-switch/span');
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
                
                // perform change WN -> SG
                [changeZustand] = await page.$x('//*[@id="grading-form"]/div[1]/ry-grading-overview/div/span[2]');
                await changeZustand.evaluate( changeZustand => changeZustand.click() );

                // perform click SG
                [surveySG] = await page.$x('//*[@id="grading-form"]/div[1]/ry-grading-questions/div[2]/div/ry-grading-radio/div[1]/div[2]/label');
                await surveySG.evaluate( surveySG => surveySG.click() );

                // get the updated price (SG)
                await delay(defaultTime * 1.5); // it is necessary         ----------------
                [bestPrice] = await page.$x('//*[@id="grading-form"]/div[2]/ry-grading-info/div/div[2]/div[1]/div[2]/div[1]/p/span');
                value = await bestPrice.getProperty('textContent');
                priceSG = await value.jsonValue();
                priceSG = parseInt(priceSG);

                // perform change SG -> Gut
                [changeZustand] = await page.$x('//*[@id="grading-form"]/div[1]/ry-grading-overview/div/span[2]');
                await changeZustand.evaluate( changeZustand => changeZustand.click() );

                // perform click Gut
                [surveyBP] = await page.$x('//*[@id="grading-form"]/div[1]/ry-grading-questions/div[2]/div/ry-grading-radio/div[1]/div[3]/label');
                await surveyBP.evaluate( surveyBP => surveyBP.click() );

                // get the updated price (Gut)
                await delay(defaultTime * 1.5); // it is necessary         ----------------
                [bestPrice] = await page.$x('//*[@id="grading-form"]/div[2]/ry-grading-info/div/div[2]/div[1]/div[2]/div[1]/p/span');
                value = await bestPrice.getProperty('textContent');
                priceBP = await value.jsonValue();
                priceBP = parseInt(priceBP);
                
                await page.goto(link);

                //write to xlsx
                worksheet.cell(counter + 1, 1).string(processor);
                worksheet.cell(counter + 1, 2).number(RAM);
                worksheet.cell(counter + 1, 3).string(SSD);

                if (model.search(color[0]) != -1){
                    worksheet.cell(counter + 1, 4).string(color[0]);
                } 
                if (model.search(color[1]) != -1){
                    worksheet.cell(counter + 1, 4).string(color[1]);
                }
                if (model.search(color[2]) != -1){
                    worksheet.cell(counter + 1, 4).string(color[2]);
                }
                if (model.search(color[3]) != -1){
                    worksheet.cell(counter + 1, 4).string(color[3]);
                }

                if (model.search(qwerty) != -1){
                    worksheet.cell(counter + 1, 5).string(qwerty);
                }

                if (isNaN(priceBP) ? worksheet.cell(counter + 1, 9).number(0) : worksheet.cell(counter + 1, 9).number(priceBP));
                if (isNaN(priceWN) ? worksheet.cell(counter + 1, 7).number(0) : worksheet.cell(counter + 1, 7).number(priceWN));
                if (isNaN(priceSG) ? worksheet.cell(counter + 1, 8).number(0) : worksheet.cell(counter + 1, 8).number(priceSG));
                if (isNaN(priceSG) ? worksheet.cell(counter + 1, 6).number(0) : worksheet.cell(counter + 1, 6).number(parseInt((priceSG + priceBP)/2)));
                if (isNaN(priceWN) ? worksheet.cell(counter + 1, 10).number(0) : worksheet.cell(counter + 1, 10).number(parseInt(priceWN - ((priceSG + priceBP)/2))));
                
                console.log(counter, 'done');
            }
            else {
                console.log(counter, 'KA');
            }
        }
    } 
    catch{
        console.log('16 Pro - issue in loop');
    }
    //        ˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙
    workbook.write(today + '.xlsx');
    console.log('\n--- file created ---\n');
    browser.close();
}
//        ˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙
scrapeMacs();

//TODO:
// 16 does not work
// Air does not work
// 12 inch что-то про модель после свитча (как если сделать 2020 на 15 дюймов)
// Починить ошибку 500 (catch)


// const sortByColumn = columnNum => {
//     if (worksheet) {
//       columnNum--;
//       const sortFunction = (a, b) => {
//         if (a[columnNum] === b[columnNum]) {
//           return 0;
//         }
//         else {
//           return (a[columnNum] < b[columnNum]) ? -1 : 1;
//         }
//       }
//       let rows = [];
//       for (let i = 1; i <= worksheet.actualRowCount; i++) {
//         let row = [];
//         for (let j = 1; j <= worksheet.columnCount; j++) {
//           row.push(worksheet.getRow(i).getCell(j).value);
//         }
//         rows.push(row);
//       }
//       rows.sort(sortFunction);
//       // Remove all rows from worksheet then add all back in sorted order
//       worksheet.spliceRows(1, worksheet.actualRowCount);
//       // Note worksheet.addRows() may add them to the end of empty rows so loop through and add to beginnning
//       for (let i = rows.length; i >= 0; i--) {
//         worksheet.spliceRows(1, 0, rows[i]);
//       }
//     }
// }

// for(i = 0; i < color.length; i++){
//     if (model.search(color[i]) != -1){
//         worksheet.cell(counter + 1, 4).string(color[i]);                // does not work
//     }
//     else{
//         worksheet.cell(counter + 1, 4).string('manual check');
//     }
// }