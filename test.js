const puppeteer = require('puppeteer');
const request = require('request-promise');
var fs = require('fs');
var XLSX = require('xlsx');
var sleep = require('system-sleep');
(async () => {

    try{
        // Viewport && Window size
        const width = 1080
        const height = 1080

        const browser = await puppeteer.launch({
            headless: true,
            args:[
               
                '--ignore-certificate-errors',
                '--ignore-certificate-errors-spki-list '
             ]
        } );
        var sheetToProcess = 0 // sheet number you want to process
        var startIndex= 0 // start row index
        var endIndex = 69999 // end row index
        var page = await browser.newPage();
        var workbook = XLSX.readFile('./astar.xlsx'); //Excel Sheet to read in
        var sheet_name_list = workbook.SheetNames;
        var temp = XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[sheetToProcess]])
        await page.setViewport({ width, height })

        var loopIndex = startIndex
        await page.goto('https://www.wikidata.org/w/index.php?search', {
           
            waitUntil: 'networkidle2',
            timeout: 3000000
        });
        for(let x = startIndex; x < endIndex ; x++){
            console.log("processing "+ x +" out of "+temp.length)
            loopIndex = x
            if(x==0 ){
                await page.goto('https://www.wikidata.org/w/index.php?search',{
           
                    waitUntil: 'networkidle2',
                    timeout: 3000000
                });
            }
            if(x%50 ==0){
                //Closeing tab and opeing tab to prevent RAM buffer overflow
                await page.close();
                page = await browser.newPage();
                await page.goto('https://www.wikidata.org/w/index.php?search',{
           
                    waitUntil: 'networkidle2',
                    timeout: 3000000
                });
            }
            if(  temp[x].token != "null"){

                
                await page.type('input[name="search"]', temp[x].token.toString()) 
               
                await page.click('button[type="submit"]')
                await page.waitForNavigation()
                var listofitems = await page.$$eval('#mw-content-text > div.searchresults > ul > li > div.mw-search-result-heading > a'
                 ,as => as.map(a =>({href:a.href , title : a.title})))
                if(listofitems.length > 0){
                    for (let y = 0 ; y < 4 ; y++){
                        if(listofitems[y].title.match(/[a-zA-Z ]+/).toString().replace(/^[ ]+|[ ]+$/g,'')== temp[x].token){
                            if(listofitems[y].href.match("https://www.wikidata.org/wiki/(.*)")[1].toString().substring(0, 1)!='P')
                            temp[x].Wikidata = "B-"+listofitems[y].href.match("https://www.wikidata.org/wiki/(.*)")[1]
                        }
                        
                    }
                }
    
                await page.click('input[name="search"]', {clickCount: 3})
                if(x%5 ==0){
                    //Used to slow the crawler to prevent wikidata firewall from blocking us :D
                    sleep(2*1000); 
                }
                else if(x%2 ==0){
                    sleep(2*1000); 

                }
            }

        }
        var ws = XLSX.utils.json_to_sheet(temp);
        var wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "report");
        //Output excel to root folder
        XLSX.writeFile(wb, 'out index from '+startIndex +"-"+ endIndex +" "+ (Date.now() % 171761)+'.xlsx');
        
     

        await browser.close();

    }catch(e){

        console.log(e)
        var ws = XLSX.utils.json_to_sheet(temp);
        var wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "report");
        //If error occur save processed rows , LoopIndex will indicate loop end at which row index
        XLSX.writeFile(wb, 'out index from '+startIndex +"-"+ endIndex +" end at "+loopIndex+" "+ (Date.now() % 171761)+'.xlsx');
    }
})();


