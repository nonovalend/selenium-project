launch.js
const index= require('./index.js');

let blockingpoint;
let success;
let oldStopAt;
let newStopAt;
let scenariolist;
let currentscenario;
let testsresults;
let testnames;
let temporary;
let trialNumber=0;

async function testrun() {
    testsresults=[]
    
    console.log("let's start !")
    temporary= await index.parsing("C:/Users/NolanRiboulet/OneDrive - NOVALEND/Documents/Différence de cas novalend.xlsx").then( parseresponse => { return parseresponse })
    scenariolist= temporary[0]
    testnames= temporary[1]
    console.log("test names are " + testnames)

    for (let scenarioNum = 1; scenarioNum < scenariolist.length; scenarioNum++) {
        currentscenario = scenariolist[scenarioNum]
        console.log("new scenarionum is " + scenarioNum)
        console.log(testnames[scenarioNum])

        blockingpoint=0
        success=0

        while (blockingpoint < 3 && success < 3) {
            trialNumber+=1
            newStopAt = await index.pathprocess(currentscenario).then(message => { return message })
            console.log(newStopAt)
            if (newStopAt == 'Success') { success += 1 ; console.log("Number of success"+ success)}
            else {
                if (oldStopAt == newStopAt|| blockingpoint==0) {
                    blockingpoint += 1
                    console.log("blocking point is " + blockingpoint)
                }
                else { blockingpoint = 0 }
            }
            oldStopAt = newStopAt
        }
    if (blockingpoint == 3) { console.log("failure at" + newStopAt) }
        else { console.log("Success in the path")
    }
    
    testsresults.push([testnames[scenarioNum],newStopAt])
    }
    console.log("all test conducted!")
    console.log(testsresults)
    console.log("rate of sucess of "+ 3*testnames.length/trialNumber)
}
//     console.log(results.length)
//     results.map(async result=>{
//         await pathprocess(result)
//     })

// parsing("C:/Users/NolanRiboulet/OneDrive - NOVALEND/Documents/Différence de cas novalend.xlsx").then(  results => {
//     console.log(results.length)
//     results.map(async result=>{
//         await pathprocess(result)
//     })
// for (let scenario=1; scenario <results.length; scenario++){
// console.log("scenario = "+scenario)
// awaitpathprocess(results[scenario])}
// })
//pathprocess("NOVALEND","Quemard Thomas","383000692","888979622","Crédit-bail","Bureautique, informatique","Lecteurs magnétique","2","10000","0","Neuf","36","Trimestrielle","A terme échu","Virement","500000","CAPITOLE FINANCE-TOFINSO" )

testrun()