const puppeteer = require('puppeteer');
const fs = require('fs');
const path = require('path');
const handlebars = require('handlebars');
const Excel = require('exceljs');

// Lettura del file template Handlebars
const templateSource = fs.readFileSync('C:/Users/User/Desktop/cribis/index.hbs', 'utf8');

// Compilazione del template Handlebars
const template = handlebars.compile(templateSource);

const options = { format: 'A4' };

const filePath = path.resolve(__dirname, 'C:/Users/User/Desktop/cribis/dataExcel/dati.xlsx');


fs.access(filePath, fs.constants.R_OK, async (err) => {
    if (err) {
        console.error(`Il file ${filePath} non può essere letto:`, err);
        return;
    }

    const workbook = new Excel.Workbook();
    let worksheet;

    await workbook.xlsx.readFile(filePath)
        .then(async () => {
            worksheet = workbook.getWorksheet('uno');
            const dataPromises = [];

            worksheet.eachRow(async (row, rowNumber) => {
                if (rowNumber > 2) {
                    const bottoneCell = row.getCell(6);

                    if (bottoneCell && bottoneCell.value) {
                        const bottone = bottoneCell.value;
                        var veicolo = row.getCell(59).value;

                        var testo = veicolo;
                        var parti = testo.split(";");
                        var veicoli = [];
                        for (var i = 0; i < parti.length; i+=4) {
                            let pt = [];
                            if(parti[i] != '' ){
                                pt.push(parti[i]);
                                pt.push(parti[i+1]);
                                pt.push(parti[i+2]);
                                pt.push(parti[i+3]);
                                veicoli.push(pt);
                            } else {
                                console.log(' vuoto');
                            }
                        }

                        const rowData = {
                            informazioneNr: row.getCell(1).value,
                            dataCliente: row.getCell(2).value,
                            dataEvasione: row.getCell(3).value,
                            cliente: row.getCell(4).value,
                            riferimentoCliente: row.getCell(5).value,
                            bottone: bottone,
                            nominativo: row.getCell(6).value,
                            cf: row.getCell(7).value,
                            indirizzo: row.getCell(8).value,
                            cap: row.getCell(9).value,
                            comune: row.getCell(10).value,
                            provincia: row.getCell(11).value,
                            cognome: row.getCell(12).value,
                            nome: row.getCell(13).value,
                            dataNascita: row.getCell(14).value,
                            luogoNascita: row.getCell(15).value,
                            cognome: row.getCell(12).value,
                            provinciaNascita: row.getCell(16).value,
                            note: row.getCell(22).value,
                            categoria: row.getCell(23).value,
                            contratto: row.getCell(28).value,
                            orario: row.getCell(29).value,
                            dataInizio: row.getCell(30).value,
                            dataScadenza: row.getCell(31).value,
                            mensileLordo: row.getCell(32).value,
                            lordoAnnuale: row.getCell(33).value,
                            datoreLavoro: row.getCell(34).value,
                            partitaIva: row.getCell(35).value,
                            titolareIva: row.getCell(24).value,
                            dataDecorrenza: row.getCell(25).value,
                            statoIva: row.getCell(26).value,
                            descrizioneIva: row.getCell(27).value,
                            protesti: row.getCell(55).value,
                            pregiudizievoli: row.getCell(56).value,
                            procedureConcorsuali: row.getCell(57).value,
                            carica: row.getCell(46).value,
                            denominazione: row.getCell(47).value,
                            partitaIva: row.getCell(48).value,
                            sedeLegale: row.getCell(49).value,
                            nRea: row.getCell(50).value,
                            oggettoSociale: row.getCell(51).value,
                            stato: row.getCell(52).value,
                            operativita: row.getCell(53).value,
                            veicoli: veicoli,
                            giudizio: row.getCell(60).value
                        };

                        const html = template(rowData);

                        const browser = await puppeteer.launch({headless: "new"});
                        const page = await browser.newPage();
                        await page.setContent(html);
                        await page.pdf({path: 'esitiGo.pdf', format: 'A4', printBackground:true});
                        console.log('PDF salvato come esitiGo.pdf');
                        await browser.close();
                    }
                }
            });
        });
});
