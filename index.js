
const { Builder, By, Key, until, WebDriverWait } = require("selenium-webdriver");
const Excel = require('xlsx');




//fonction de parsing depuis un excel
exports.parsing= async function parsing(link) {

    let row;
    const workbook = Excel.readFile(link);
    const firstSheet = workbook.SheetNames[0];
    const sheetDataJSON = Excel.utils.sheet_to_json(workbook.Sheets[firstSheet]);
    const nbLines = sheetDataJSON.length;
    //console.log(sheetDataJSON)
    const sheetDataArray = Object.values(sheetDataJSON);
    const headerRow = Object.values(sheetDataJSON[0]);
    //console.log("header row is "  +headerRow)
    const nbColumns = headerRow.length
    const result = []

    for (let i = 0; i < sheetDataArray.length; i++) {
        row = sheetDataJSON[i]
        result.push(Object.values(row))
    }
    console.log(result)
    finalresults = []
    console.log("Columns" + nbColumns)
    console.log("Lines" + nbLines)

    for (let c = 0; c < nbColumns; c++) {
        column = []
        for (let l = 1; l < nbLines; l++) {
            column.push(result[l][c])
        }
        finalresults.push(column)
    }
    console.log("final results are " + finalresults[1])
    return [finalresults,headerRow]
}

exports.pathprocess= async function pathprocess([Vendor_In, Salesman_In, sirenClient_In, sirenSupplier_In, FinanceType_In, IsLeaseBack_In, EquipementType_In, ProductType_In, VRRate_In, VRAmount_In, IsFirstRent_In, FirstRentRate_In,  Retake_In, ProdState_In,SecHandYr_In, Duration_In, Periodicity_In, Term_In, PayMean_In, FinanceAmount_In, Refin_In, Maint_In, Equip1Name, Equip1Brand, Equip1Amount]) {
    
    //fonctions pour remplir champs et dropdown

    //_remplir les champs classiques
    async function clickButtonByxpath(xpath) {
        try {
            let element = await driver.wait(until.elementLocated(By.xpath(xpath)), 30000);
            console.log("located")
            await driver.wait(until.elementIsVisible(element), 30000);
            console.log("visible")
            await driver.wait(until.elementIsEnabled(element), 30000).then(element => element.click())
        }
        catch (error) {
            if (error.message == 'TimeoutError') {
                console.error(xpath + " took to much time")
            }   
            stopAt = xpath
            console.log(stopAt)
            throw error 
        }
    }

    async function findbyClassname(className, fieldInput) {
        let element = await driver.wait(until.elementLocated(By.className(className)), 30000);
        await driver.wait(until.elementIsVisible(element), 10000);
    }

    async function findbyxpath(xpath) {
        try {
            let element = await driver.wait(until.elementLocated(By.xpath(xpath)), 30000);
            return driver.wait(until.elementIsVisible(element), 30000)
        }
        catch (error) {
            if (error.message == 'TimeoutError') {
                console.error(xpath + " took to much time")
            }  
            stopAt = xpath
            console.log(stopAt)
            throw error  
        }
    }

    async function fieldfill(fieldID, fieldInput) {
        try {
            let element = await driver.wait(until.elementLocated(By.id(fieldID)), 30000);
            await driver.wait(until.elementIsVisible(element), 30000);
            await driver.sleep(500)
            await element.sendKeys(fieldInput);
            console.log(fieldID + " Filled")
        }
        catch (error) {
            if (error.message == 'TimeoutError') {
                console.error(fieldInput + " took to much time")
            }
            stopAt = fieldInput
            console.log(fieldID)
            throw error
        }
    }

    //_remplir les dropdowns
    async function dropdownfill(dropID, dropField) {
        try {
            let element = await driver.wait(until.elementLocated(By.id(dropID)), 30000);
            await driver.wait(until.elementIsVisible(element), 30000).then(button => button.click());
            await driver.sleep(200);
            await driver.findElement(By.xpath("//*[text()='" + dropField + "']")).then(option => option.click())
            console.log(dropID + " Filled");
        }
        catch (error) {
            if (error.message == 'TimeoutError') {
                console.error(xpath + " took to much time")}
            stopAt = dropField
            console.log(stopAt)
            throw error
        }
    }

    async function autodropdownfill(dropID) {
        let element = await driver.wait(until.elementLocated(By.id(dropID)), 30000);
        await driver.wait(until.elementIsVisible(element), 30000);
        await element.click();
        await driver.sleep(200);
        let searchbar = await driver.wait(until.elementLocated(By.className("select2-search__field")),30000);
        await searchbar.sendKeys("", Key.RETURN);
        console.log(dropID + " Filled");
    }

    //caractérisation de tous les champs à remplir
    //page d'accueil 
    let Quot_Button = '//*[@id="appVue"]/header/div[2]/div/a'

    //page de devis
    //_by ID
    let Vendor_DD = "select2-platformId-container";
    let Salesman_DD = "select2-platformSalesmanformsimulation-container";
    let Supplier_Field = "sirenSupplier";
    let Client_Field = "sirenClient";
    let FinanceType_DD = "select2-financecontractTypeformsimulation-container";
    let LeaseBack_Check = "checkLeaseback";
    let FirstRent_Check = "checkinitialRent";
    let FirstRentAmount_Field = "finance.initialRent"
    let SecHandYear_DD= "select2-equipmentyearformsimulation-container"
    let EquipementType_DD = "select2-equipmentequipmentTypeformsimulation-container";
    let ProductType_DD = "select2-equipmentproductTypeformsimulation-container";
    let VRRate_Field = "finance.residualPercent";
    let VRRateVendor_DD="select2-financeresidualPercentformsimulation-container"
    let VRAmount_Field = "finance.residualValue";
    let Retake_Field = "finance.reversal";
    let ProdState_DD = "select2-equipmentconditionformsimulation-container";
    let Duration_DD = "select2-financedurationformsimulation-container";
    let Periodicity_DD = "select2-financeperiodformsimulation-container";
    let Term_DD = "select2-financepaymentTermformsimulation-container";
    let PayMean_DD = "select2-financepaymentMeanformsimulation-container";
    let FinanceAmount_Field = "finance.amount";

    //_by xpath
    let QuotCalc_Button = '//*[@id="appVue"]/main/div[1]/div/div[2]/div/div[23]/div/div[1]/div[1]/button'

    //page Ma simulation
    //_byxpath
    let Simulation_Tab = '//*[@id="main-content"]/div/div/div[1]/div[3]/div[1]/a'
    let SimulPop_Button = "//*[@id='modal-info']/div/div/div[2]/button"
    let Contract_Button = '//*[@id="d_proposition"]/div/div/div/div[3]/div[4]/div[2]/button[2]'

    //page d'équipements
    //_byID
    let EquipBrand_Field = "brand";
    let EquipModel_Field = "model";
    let EquipModelSeries_Field = "serialNumber";
    let EquipUnitAmount_Field = "unitAmount";
    let EquipUniAmounSoft_Field = "unitAmountSoft";
    let UnitAmountPrestation_Field = "unitAmountPrestation";
    let EquipQuantity_Field = "quantity"

    //_byXpath
    let Equip_Tab = '//*[@id="main-content"]/div/div/div[1]/div[3]/div[2]/a'
    let ModifyEquip_Button = '//*[@id="main-content"]/div/div/div[2]/div[2]/div/div/div[2]/div[1]/div/form/table/tbody/tr/td[8]/div/a'
    let ModifyEquipPop_Button = '/html/body/div[2]/div[3]/div/div/div[2]/div[12]/button[1]'
    let AddEquip_Button = '//*[@id="main-content"]/div/div/div[2]/div[2]/div/div/div[2]/div[1]/div/form/table/div/a';
    let AddEquipPop_Button = '//*[@id="main-content"]/div/div/div[2]/div[2]/div/div/div[2]/div[1]/div/form/table/tbody/tr/td[8]/div/a'

    //page de décision
    //_byID
    
    let MailType_In = " Mail d échéance d accord ";
    

    //_byXpath
    let Decision_Tab = "//*[@id='main-content']/div/div/div[1]/div[3]/div[5]/a"
    let AddRecipient_Button = '//*[@id="message-form"]/div/div[2]/div[1]/button'
    let MailType_DD='/html/body/div[2]/main/div/div/div/div[2]/div[3]/div/div/div[2]/div/div[2]/div[2]/div[1]/div/div[1]/div[1]/div/div/div/div/div[2]/form/div/div[1]/div/div/div[2]/div[1]/div/div/span/span[1]/span'
    let RecipentName_Field = '//*[@id="message-form"]/div/div[2]/div[1]/table/tbody/tr[4]/td[1]/input'
    let RecipientEmail_Field = '//*[@id="message-form"]/div/div[2]/div[1]/table/tbody/tr[4]/td[2]/input'
    let SendEmail_Button = '/html/body/div[2]/main/div/div/div/div[2]/div[3]/div/div/div[2]/div/div[2]/div[2]/div[1]/div/div[1]/div[1]/div/div/div/div/div[3]/button[2]'
    //_byClassName
    //let MailType_DD = "select2-selection__rendered" div/

    //page de validation 
    //_ByID div/div[1]/div/div/div[2]/div[1]/div/div/span/span[1]/span
    let Establishment_DD = "select2-ownerformcontract-container"
    let Director_DD = "select2-directorformcontract-container"
    let AddSigner_Button = "add-signer-add"
    let SignerFirstname_Field = "userEntityForm-user-firstname"
    let SignerLastname_Field = "userEntityForm-user-lastname"
    let SignerEmail_Field = "userEntityForm-user-email"
    let SignerPhone_Field = "userEntityForm-user-phone"
    let SignerPosition_DD = "select2-userEntityForm-position-container"

    //_ByXpath
    let SignerPop_Button = '//*[@id="modal-add-signer"]/div/div/div[3]/button[2]'
    let Validate_Button = '//*[@id="contract"]/div[6]/button'

    let stopAt;
    
    //démarrage du programme
    let driver = await new Builder().forBrowser("chrome").build();
    try {
        stopAt = "Start"

        //connect to novalend website
        await driver.get("https://test.novalend.com");
        await driver.findElement(By.name("email")).sendKeys("nolan@novalend.com");
        await driver.findElement(By.name("password")).sendKeys("nolanPC*38", Key.RETURN);

        //navigate to creer un devis
        await clickButtonByxpath(Quot_Button);

        //fill in creer un devis
        //console.log("helloworld1");
        stopAt = "devis"
        await dropdownfill(Vendor_DD, Vendor_In);
        await driver.sleep(500)
        await dropdownfill(Salesman_DD, Salesman_In);
        await driver.sleep(500)
        await fieldfill(Client_Field, sirenClient_In);
        let checkSupplier = await driver.findElements(By.id(Supplier_Field)).then(element => {return element.length})
        console.log(checkSupplier)
        if (checkSupplier != 0) {
            await fieldfill(Supplier_Field, sirenSupplier_In);
        }
        if (IsLeaseBack_In == "Oui") {
            await driver.findElement(By.id(LeaseBack_Check)).then(button => {return button.click()})
        }       
        await dropdownfill(FinanceType_DD, FinanceType_In);
        await dropdownfill(EquipementType_DD, EquipementType_In);
        await dropdownfill(ProductType_DD, ProductType_In);
        if (checkSupplier!=0){
            await fieldfill(VRRate_Field, VRRate_In)
        } else{await dropdownfill(VRRateVendor_DD, "2%");}
        if (IsFirstRent_In == "Oui") {
            await driver.findElement(By.id(FirstRent_Check)).then(button => {return button.click()})
            await fieldfill(FirstRentAmount_Field, FirstRentRate_In)
        }
        console.log ( "prodstate "+ ProdState_In )
        console.log("Second hand"+ SecHandYr_In)
        await fieldfill(Retake_Field, Retake_In);
        await dropdownfill(ProdState_DD, ProdState_In);
        if (ProdState_In=="Occasion"){

            console.log("please wait")

            await driver.sleep(2000)
            console.log("finish waiting")
            await dropdownfill(SecHandYear_DD, SecHandYr_In)
        }
        await dropdownfill(Duration_DD, Duration_In);
        await dropdownfill(Periodicity_DD, Periodicity_In);
        await dropdownfill(Term_DD, Term_In);
        if (checkSupplier!=0){
            await dropdownfill(PayMean_DD, PayMean_In);
        }
        
        await fieldfill(FinanceAmount_Field, FinanceAmount_In);
        await clickButtonByxpath(QuotCalc_Button);

        //_close popup
        await driver.wait(until.urlContains("devis.php"),30000)
        await driver.sleep(200);
        stopAt = "Ma simulation"
        await clickButtonByxpath(SimulPop_Button);
        await driver.sleep(5000)
        
        //console.log("pop up closed");

        //Ajouter des équipements
        stopAt = "Equipement"

        //_Aller sur la page des équipements
        // await clickButtonByxpath(Equip_Tab)
        // await driver.wait(until.urlContains("service"))

        //ajouter des équipements
        // await clickButtonByxpath(ModifyEquip_Button)
        // await fieldfill(EquipBrand_Field, "brand");
        // await fieldfill(EquipModel_Field, "Model")
        // await fieldfill(EquipModelSeries_Field, "123")
        // await fieldfill(EquipUnitAmount_Field, "23")
        // await fieldfill(EquipUniAmounSoft_Field, "0");
        // await fieldfill(UnitAmountPrestation_Field, "15")
        // await fieldfill(EquipQuantity_Field, "15");
        // await clickButtonByxpath(ModifyEquipPop_Button)

        //navigate to decision
//*[@id="collapse1"]/div[1]/div/div//*[@id="collapse1"]/div[1]/div/div
//*[@id="collapse1"]/div[1]/div/div/button/div
        //_go to decision tab
        stopAt = "Decision"
        await clickButtonByxpath(Decision_Tab)
        await driver.sleep(5000);
        //console.log("moved to decision");

        //recherche du bailleur
        let refinancerNumber = await driver.findElements(By.className('panel-group p-2')).then(list => { return list.length })
        //console.log(refinancerNumber)
        let n;
        console.log("go for loop");
        for (n = 0; n <= refinancerNumber; n++) {
            let RefinancerName = await findbyxpath('//*[@id="heading' + n + '"]/div[1]/div[1]/h4').then(element => { return element.getText() });
            //console.log(RefinancerName)
            if (RefinancerName == Refin_In) {
                break
            }
        };//*[@id="accordion-1"]/div/div
        //console.log("refinancer_found");//*[@id="accordion-1"]/div/div

        //selection du bailleur
        await clickButtonByxpath('//*[@id="heading'+n+'"]/div[3]/div[2]/button')
        await driver.sleep(10000)
        console.log("we are waiting" )
        // let SelectedTile=  findbyxpath('//*[@id="accordion-'+n+'"]/div/div');
        // driver.wait(until.attr(SelectedTile, 'class', 'bloc border__dec--selected'), 20000);
        await clickButtonByxpath('//*[@id="heading' + n + '"]/div[3]/div[2]/a')
        await driver.sleep(1000)
        //envoyer la demande
        await driver.wait(until.elementLocated(By.xpath('//*[@id="collapse' + n + '"]/div[1]/div/div/button')),30000);
        let demandSentDetect = await driver.findElements(By.xpath('//*[@id="collapse' + n + '"]/div[1]/div[2]/div'));
        console.log("is it displayed ?  " + demandSentDetect.length);

        //_distinction des cas accords/pas d'accords 
        if (demandSentDetect.length == 0) {

            //_Ouvrir popup

            await clickButtonByxpath('//*[@id="collapse' + n + '"]/div[1]/div/div/button')
            //console.log("demand popup opened");

            //_remplissage du dropdown type de mail
            console.log("start countdown")
            await driver.sleep(5000)
            console.log("finished countdown")
            await findbyxpath(MailType_DD).then(button => button.click())
            //await findbyClassname(MailType_DD).then(button => button.click())
        
            await driver.sleep(300)
            await driver.findElement(By.xpath("//*[text()='" + MailType_In + "']")).then(option => option.click());
            //console.log("template selected");

            //_ajout d'un destinataire
            await findbyxpath(AddRecipient_Button).then(button => button.click());
            //console.log("new destinary added");

            //_coordonnées destinataire et envoi
            await findbyxpath(RecipentName_Field).then(field => field.sendKeys("Commercial"));
            await findbyxpath(RecipientEmail_Field).then(field => field.sendKeys("commercial@novalend.com"))
            //console.log("Details filled in");
            await clickButtonByxpath(SendEmail_Button)
            //console.log("demand sent");

            //Réponse
            //_Remplissage du dropdown statut
            await dropdownfill("select2-statusformdemand"+n+"-container", "Accord")
            await driver.sleep(500)

            //_Transmettre ma décision
            await clickButtonByxpath('//*[@id="collapse' + n + '"]/div[1]/div[3]/div/button[2]')
            //console.log("decision sent");
        }

        //Retour sur simulation 
        //console.log("on simulation");
        stopAt = "Ma simulation bis";
        await driver.sleep(3000)
        await clickButtonByxpath(Simulation_Tab)
        await driver.wait(until.urlContains("devis.php"),30000)
        await driver.sleep(3000);

        //_Fermer la popup
        await clickButtonByxpath(SimulPop_Button)
        console.log("pop-up closed");

        await driver.sleep(2000);

        //_Créer le contrat
        let message = await findbyxpath(Contract_Button).then(button =>button.getText())
        console.log(message)
        if (message=="Simuler cet accord"){
            console.log("On recharge");
            clickButtonByxpath(Contract_Button)
            await driver.sleep(1000)
            clickButtonByxpath(SimulPop_Button)
            await driver.sleep(2000)
            console.log("pop-up closed");

        }
        await clickButtonByxpath(Contract_Button)

        //Valider les informations

        //_établissement
        await autodropdownfill(Establishment_DD);

        //_directeur 

        await autodropdownfill(Director_DD);
        //_signataire
        await driver.findElement(By.id(AddSigner_Button)).then(button => button.click());
        await fieldfill(SignerFirstname_Field, "Mathieu");
        await fieldfill(SignerLastname_Field, "Belle");
        await fieldfill(SignerEmail_Field, "raja@novalend.com");
        await fieldfill(SignerPhone_Field, "0770350647");
        await autodropdownfill(SignerPosition_DD)

        //_fermeture poppup
        await driver.findElement(By.xpath(SignerPop_Button)).click()

        //_validation de la page
        await clickButtonByxpath(Validate_Button)
        stopAt = "Success"
        driver.close()
        return stopAt
    }
    catch (error) {
            console.error("error is " + error.message)
        console.log(error.stack.split("\n"))
        driver.close()
        return stopAt
    }
}




/*
    
    //upload des documents
    await driver.sleep(5000)
    let documentList = await findbyxpath('//*[@id="main-content"]/div/div/div[2]/div[2]/div/div[2]/div/div/form/table/tbody[1]')
    console.log("Length is "+ documentList.length)
    console.log("Size is " + documentList.getSize)
    for (let i=0 ; i<=documentList.length;i++ ){
        documentTitle=await findbyxpath('//*[@id="main-content"]/div/div/div[2]/div[2]/div/div[2]/div/div/form/table/tbody[1]/tr['+n+']/td[1]').getText();
        console.log("Title " + documentTitle);
        documentStatusTile=await findbyxpath('//*[@id="main-content"]/div/div/div[2]/div[2]/div/div[2]/div/div/form/table/tbody[1]/tr['+n+']/td[2]')
        documentStatus=documentStatusTile.getText();
        docInput= await findbyxpath('//*[@id="main-content"]/div/div/div[2]/div[2]/div/div[2]/div/div/form/table/tbody[1]/tr['+n+']/td[3]/div/button');
        console.log("Status " + documentStatus)

        if (documentTitle.includes("Pièce d'identité") && documentStatus.includes("Manquant")){
            await docInput.sendKeys('C:/Users/NolanRiboulet/NOVALEND/GENERAL - Documents/3 - Plateforme SaaS/0. Divers/Pieces jointes Tests/Documents Tests/Passeport M Belle.pdf')
        }
        if (documentTitle.includes("RIB") && documentStatus.includes("Manquant") ){
            await docInput.sendKeys('C:/Users/NolanRiboulet/NOVALEND/GENERAL - Documents/3 - Plateforme SaaS/0. Divers/Pieces jointes Tests/Documents Tests/RIB Fournisseur.pdf')
        }
        if (documentTitle.includes("Délégation") && documentStatus.includes("Manquant") ){
            await docInput.sendKeys('C:/Users/NolanRiboulet/NOVALEND/GENERAL - Documents/3 - Plateforme SaaS/0. Divers/Pieces jointes Tests/Documents Tests/delegation.pdf')
        }
        await driver.wait(until.elementTextContains(documentStatusTile,"Chargé"))
    }

    await driver.findElement(By.id("_btn_sign_contract")).click()
    /*let directorInput= await driver.wait(until.elementLocated(By.id("48891")))
    await directorInput.sendKeys('C:/Users/NolanRiboulet/NOVALEND/GENERAL - Documents/3 - Plateforme SaaS/0. Divers/Pieces jointes Tests/Documents Tests/Passeport M Belle.pdf')
    let signerInput = await driver.wait(until.elementLocated(By.id("48892")))
    await signerInput.sendKeys('C:/Users/NolanRiboulet/NOVALEND/GENERAL - Documents/3 - Plateforme SaaS/0. Divers/Pieces jointes Tests/Documents Tests/Passeport M Belle.pdf')
    let delegInput= await driver.wait(until.elementLocated(By.id("48893"),30000));
    await delegInput.sendKeys( "C:/Users/NolanRiboulet/NOVALEND/GENERAL - Documents/3 - Plateforme SaaS/0. Divers/Pieces jointes Tests/Documents Tests/delegation.pdf")*/

