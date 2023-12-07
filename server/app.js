// app.js
const express = require('express')
const app = express()
const cors = require('cors')
const { API_KEY, logPassword, logUser, firebaseConfig } = require('./config');
const { Z_OK } = require('zlib')

// Firebase
const admin = require('firebase-admin');
const serviceAccount = require('./firebase_config.json'); // Replace with your own service account key
admin.initializeApp({
    credential: admin.credential.cert(serviceAccount),
    databaseURL: `${firebaseConfig.databaseURL}` // Replace with your Firebase project URL
});

// Puppeteer
const puppeteer = require('puppeteer-extra')
const StealthPlugin = require('puppeteer-extra-plugin-stealth')
puppeteer.use(StealthPlugin())

// 2Captcha
const { Solver } = require('2captcha');
const { readFileSync } = require('fs');
const solver = new Solver(API_KEY);

// ExcelJS
const ExcelJS = require('exceljs');


app.use(cors())
app.use(express.json())
app.get('/', (req, res) => {
    res.json({ success: true, message: 'welcome to the api' })
})

app.get('/startProcess', (req, res) => {
    const collectionRef = admin.firestore().collection('processes');
    const startDate = new Date().toISOString()
    const newProcess = {
        numbNIFs: parseInt(req.query.numNifs),
        startingDate: startDate
    }
    collectionRef.add(newProcess)
        .then(docRef => {
            console.log('Document added with ID:', docRef.id);
            res.json({ success: true, message: 'Document added with ID: ' + docRef.id, docID: docRef.id })
        })
        .catch(err => {
            console.error('Error adding document:', err);
            res.json({ success: false, message: 'Document adding failed' })
        });
})

app.post('/validate', (req, res) => {
    const nif = req.body.nif
    const process = req.body.process
    let context = ""
    let validationStatus = false
    try {
        puppeteer.launch({ headless: true }).then(async browser => {
            const page = await browser.newPage()
            await page.setViewport({ width: 1280, height: 1024 });
            await page.goto('https://www.portaldasfinancas.gov.pt/at/html/index.html', { waitUntil: ['load', 'domcontentloaded'] })
            try {
                //Login
                const element = await page.waitForSelector('.login > li:nth-child(2) > a')
                await element.click()
                await element.dispose()
                const element2 = await page.waitForSelector('#content-area > div > div > label:nth-child(4)')
                await element2.click()
                await element2.dispose()
                await page.type("#username", logUser);
                await page.type("#password-nif", logPassword);
                const element3 = await page.waitForSelector('#sbmtLogin')
                await element3.click()
                await element3.dispose()

                while (!validationStatus) {
                    // Open "Identificação de Clientes/Fornecedores" page
                    await page.goto('https://sitfiscal.portaldasfinancas.gov.pt/terceiros/clientefornecedores', { waitUntil: ['load', 'domcontentloaded'] })
                    const element5 = await page.waitForSelector('#main-content > div > section > div > div > div > div:nth-child(2) > a')
                    await element5.click()
                    await element5.dispose()

                    //Process NIF
                    console.log('---------------------------');
                    console.log('NIF to be processed: ' + nif);
                    // Get text printscreen to use in validation
                    try {
                        await page.waitForSelector('#main-content > form > div:nth-child(3) > div > div', { waitUntil: ['load', 'domcontentloaded'] });
                        const screlement = await page.$('#main-content > form > div:nth-child(3) > div > div');
                        await screlement.screenshot({ path: 'captchanif.png' });
                        console.log('Captcha print OK');
                    } catch (error) {
                        console.log('Error in printscreen captcha');
                        validationStatus = false
                        //await browser.close()
                    }

                    // Decoding captcha
                    try {
                        const imageBase64File = await readFileSync('./captchanif.png', 'base64');
                        const response = await solver.imageCaptcha(imageBase64File);
                        await page.type("#nifCliente", nif.toString());
                        await page.type("#respostaCaptcha", response.data);
                        const element5 = await page.waitForSelector('#main-content > form > div:nth-child(2) > div > div.col-sm-4.col-xs-12 > div > button')
                        await element5.click()
                        await element5.dispose()
                        console.log('Decoding API call completed');
                    } catch (error) {
                        console.log('Error solving captcha');
                        validationStatus = false
                        //await browser.close()
                    }

                    // Validating if decoding worked
                    await page.waitForTimeout(2000)
                    console.log('Waited before printscreen');
                    if (await page.$("#main-content > form > div:nth-child(3) > div > div > label.help-block.server-error-message") !== null) {
                        await page.evaluate(() => document.getElementById("nifCliente").value = "")
                        console.log('Decoding failed');
                        validationStatus = false
                    } else {
                        console.log('Entered in decoding OK');
                        await page.waitForTimeout(1000)
                        console.log('Waited');
                        await page.screenshot({ path: "./prints/" + nif.toString() + ".png", fullPage: true })
                        if (await page.$("#main-content > div:nth-child(4) > div > div > div.panel-body > div > div:nth-child(1) > dl > dt")) {
                            const nifHeadData1 = await page.$eval('#main-content > div:nth-child(4) > div > div > div.panel-body > div > div:nth-child(1) > dl > dt', el => el.textContent);
                            const nifBodyData1 = await page.$eval('#main-content > div:nth-child(4) > div > div > div.panel-body > div > div:nth-child(1) > dl > dd', el => el.textContent);
                            const nifHeadData2 = await page.$eval('#main-content > div:nth-child(4) > div > div > div.panel-body > div > div:nth-child(2) > dl > dt', el => el.textContent);
                            const nifBodyData2 = await page.$eval('#main-content > div:nth-child(4) > div > div > div.panel-body > div > div:nth-child(2) > dl > dd', el => el.textContent);
                            context = nifHeadData1 + ': ' + nifBodyData1 + ' | ' + nifHeadData2 + ': ' + nifBodyData2
                            console.log('Context Value: ' + nifHeadData1 + ': ' + nifBodyData1 + ' | ' + nifHeadData2 + ': ' + nifBodyData2);
                        } else {
                            const nifBodyData = await page.$eval('#main-content > div:nth-child(4) > div > div > div.panel-body > div > div > dd', el => el.textContent);
                            context = nifBodyData
                            console.log("Context Value: " + nifBodyData);
                        }
                        console.log('Decoding worked');
                        validationStatus = true
                    }
                    console.log('Validation Status: ' + validationStatus);
                }
            } catch (error) {
                console.log('Validation NOK - 2');
                //await browser.close()
                //res.json({ message: 'Process failed' })
            }

            try {
                const baseURL = "http://localhost:3000";
                const response = await fetch(`${baseURL}/insertNIF`, {
                    method: "POST",
                    headers: {
                        "Content-Type": "application/json"
                    },
                    body: JSON.stringify({ nif: nif, process: process, context: context })
                });
                const data = await response.json();
                console.log(data);
            } catch (error) {
                console.error(error);
            }

            await browser.close()
            res.json({ success: true, message: nif + ' was validated with success.' })
            console.log('Process closed');
        })
    } catch {
        res.json({ success: false, message: 'There was an erro validating ' + nif })
    }
})

app.get('/getStatistics', async (req, res) => {
    const collectionRef = admin.firestore().collection('process');
    const snapshot = await collectionRef.count().get();
    const sumAggregateQuery = collectionRef.aggregate({
        totalNIFs: AggregateField.sum('numbNifs'),
    });
    const snapshotAggregate = await sumAggregateQuery.get();

    res.json({ runs: snapshot.data().count, totalNIFs: snapshotAggregate.data().totalNIFs })
})

app.get('/getLatestRuns', async (req, res) => {
    const docArray = [];
    const collectionRef = admin.firestore().collection('process');
    const snapshot = await collectionRef.get();
    snapshot.forEach(element => {
        docArray.push({ docBody: element.data(), docID: element.id })
    })

    res.json({ numRuns: docArray.length, runsArray: docArray })
})
app.get('/getLatestNIFs', async (req, res) => {
    const docArray = [];
    const collectionRef = admin.firestore().collection('nifs');
    const snapshot = await collectionRef.get();
    snapshot.forEach(element => {
        docArray.push({ docBody: element.data(), docID: element.id })
    })

    res.json({ numRuns: docArray.length, runsArray: docArray })
})

app.post('/insertNIF', async (req, res) => {

    const runID = req.body.process
    const nif = req.body.nif
    const context = req.body.context
    const collectionRef = admin.firestore().collection('nifs');
    const newData = {
        context: context,
        number: nif,
        process: runID,
        status: true,
    };
    collectionRef.add(newData)
        .then(docRef => {
            console.log('Document added with ID:', docRef.id);
        })
        .catch(err => {
            console.error('Error adding document:', err);
        });
    res.json({ message: "Done" })
})

app.post('/test', async (req, res) => {
    console.log('Response: ' + req.body.data1);
    res.send(req.body)
})

app.get('/export/:process', async (req, res) => {
    const process = req.params.process
    const docArray = [];
    const collectionRef = admin.firestore().collection('nifs');
    const snapshot = await collectionRef.where('process', '==', process).get();
    snapshot.forEach(element => {
        docArray.push({ nif: element.data().number, context: element.data().context })
    })

    const workbook = new ExcelJS.Workbook()
    const worksheet = workbook.addWorksheet('Process');
    const sheetColumns = [
        { key: 'nif', header: 'NIF' },
        { key: 'context', header: 'Contexto Fiscal' },
    ];
    worksheet.columns = sheetColumns
    console.log('docArray: ' + docArray.length);
    docArray.forEach(doc => {
        console.log('nif: ' + doc.nif + ' | context: ' + doc.context);
        worksheet.addRow({ nif: doc.nif, context: doc.context })
    })

    res.setHeader(
        "Content-Type",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );
    res.setHeader(
        "Content-Disposition",
        "attachment; filename=" + "exportTest.xlsx"
    );

    return workbook.xlsx.write(res).then(function () {
        res.status(200).end();
    });

    /* await workbook.xlsx.writeFile('./excel/exportTest.xlsx');
    res.send("Documento criado.") */
})

app.listen(3000, () => { console.log('API Running on http://localhost:3000') })