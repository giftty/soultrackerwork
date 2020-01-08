'use strict';
const prompts= require('prompts')
const {Builder,Key,until,By,error}= require ('selenium-webdriver');
const chrom =require('selenium-webdriver/chrome')
//const Excel = require ('exceljs');
const excelToJson = require('convert-excel-to-json');

const result = excelToJson({
    sourceFile: 'RealNOBDATA.xlsx',
    header:{
        rows: 1
    },
    columnToKey: {
        A: 'FIRSTNAME',
        B: 'LASTNAME',
        C: 'EMAIL',
        D: 'GENDER',
        E: 'PHONE',
        F: 'STATE',
        G: 'CITY'
    }
});
//console.log(result.Sheet2[1])
//let workbook = new Excel.Workbook();
  (async function(succ,err){
    let username = await prompts({
      type:'text',
    name:'usern',
    message:"Please enter username",
    validate:usern=>usern.length<=3?"No username entered program stopping":true
  });
   
    let pass=  await prompts({type:'text',
    name:'password',
    message:"Please enter your password",
    validate:password=>password.length<=3?"No password entered program stopping":true});

     //await console.log(`${username.usern}  ${pass.password}`)
   let  datastart=await prompts({type:'number',name:'start', message:"Enter the number to start",  validate:start=>isNaN(start)?"Enter the number to start":true})
   let dataend= await  prompts({type:'number',name:'end', message:"Enter where to end",  validate:end=>isNaN(end)?"End is not a number ":true})
     
    if(username.usern.length>=4  && pass.password.length>=3  ){
      try{
       
        let driver =  await  new Builder().setChromeOptions(new chrom.Options().addArguments('--incognito'))
        .forBrowser('chrome')
        .build();
        await driver.get('http://christembassysoultracker.org/soul');
        await  driver.wait(until.elementLocated(By.id('inputEmail')),5000).sendKeys(username.usern);
        await  driver.wait(until.elementLocated(By.id('inputPassword')),5000).sendKeys(pass.password);
        await  driver.wait(until.elementLocated(By.css('.btn-login ')),5000).sendKeys(Key.ENTER);
        await console.log(datastart.start+":"+dataend.end);
       // await workbook.xlsx.readFile('RealNOBDATA.xlsx')
       // var worksheet = workbook.getWorksheet(2);
        for (var i=datastart.start;i<=dataend.end;i++) 
        {
         try{

         var  firstname= result.Sheet2[i].FIRSTNAME             //worksheet.getRow(i).getCell(1).value;
         var lastname= result.Sheet2[i].LASTNAME                // worksheet.getRow(i).getCell(2).value;
          var phone= result.Sheet2[i].PHONE                    // worksheet.getRow(i).getCell(5).value;
          var email =result.Sheet2[i].EMAIL
          var gen=result.Sheet2[i].GENDER                    // worksheet.getRow(i).getCell(3).value
         console.log(`${firstname}  ${lastname}`)
         
         await driver.wait(until.elementLocated(By.name('soul_firstname')),5000).sendKeys(firstname);
         await driver.wait(until.elementLocated(By.name('soul_lastname')),5000).sendKeys(lastname);
         await driver.wait(until.elementLocated(By.name('soul_email')),5000).sendKeys(email);
          await /*worksheet.getRow(i).getCell(4)*/gen==="male"? await driver.wait(until.elementLocated(By.name('soul_gender')),5000).sendKeys(Key.BACK_SPACE,Key.ARROW_DOWN,Key.ENTER):await driver.wait(until.elementLocated(By.name('soul_gender')),5000).sendKeys(Key.BACK_SPACE,Key.ARROW_DOWN,Key.ARROW_DOWN,Key.ENTER);
         await driver.wait(until.elementLocated(By.name('soul_city')),5000).sendKeys('PORTARCOURT');
         await driver.wait(until.elementLocated(By.name('soul_country_code')),5000).sendKeys('NNNNNNNNNN');
         await driver.wait(until.elementLocated(By.name('soul_phoneno')),5000).sendKeys(phone);
          
         await driver.sleep(4000)
         await driver.wait(until.elementLocated(By.css('button.btn-primary')),5000).click();
         
        
         await driver.get('http://christembassysoultracker.org/track'); 
          await driver.sleep(5000)
         await driver.wait(until.elementLocated(By.css('input.input-sm')),20000).sendKeys(lastname,Key.ENTER); 
         await driver.sleep(4000)
          var trackValue= await driver.wait(until.elementLocated(By.id("mytable_processing")),20000).getCssValue()
           while(trackValue=="block"){
            await driver.sleep(4000)
           }
         //await driver.wait(until.elementLocated(By.css('a.edit_record')),5000).click()
          await driver.wait(until.elementLocated(By.css("a[href='javascript:void(0);'][data-format_phone_no='+234"+phone+"']")),25000).click()

         // await driver.wait(until.elementLocated(By.xpath("//a[ contains(@href, 'javascript:void(0);')]")), 10000);
         await driver.sleep(5000)
         await driver.wait(until.elementLocated(By.name('foundation_message')),5000).sendKeys(Key.ENTER,Key.ARROW_DOWN,Key.ARROW_DOWN,Key.ARROW_DOWN,Key.ARROW_DOWN,Key.ARROW_DOWN,Key.ARROW_DOWN,Key.ARROW_DOWN,Key.ENTER);
          await driver.sleep(3000)
           await driver.wait(until.elementLocated(By.css("button[type='submit']")),5000).click()
           console.log(`Number ${i} sucessfull`)
          worksheet.getRow(i).getCell(8).value =`Sucessfully Registered for ${username.usern}` ;
          
         
        }catch(error){
          //worksheet.getRow(i).getCell(8).value ='Error ';
          console.log(` error in ${i} ${error}`)
          await driver.get('http://christembassysoultracker.org/soul');
         }
          
        }
       // workbook.xlsx.writeFile('RealNOBDATAProcessed.xlsx')
        //  console.log('file written sucessfully')
       // var usernameCol = worksheet.getRow(i).getCell(3).value;
       // var passCol = worksheet.getRow(i).getCell(4).value; 
       
        
       
      }catch(error){
        console.log(error.message)
      }
    }else{
      driver.quit();
    }
    
    
  })()                                                                                                                         


