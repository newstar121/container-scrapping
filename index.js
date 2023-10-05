const { format, isBefore, subDays } = require("date-fns");
const axios = require('axios');

const XLSX = require('xlsx');
// const workbook = XLSX.readFile('./result.xlsx');
const workbook = XLSX.utils.book_new();
// const data = require('./test');

const API_URL = 'https://inspectorulpadurii.ro/';

async function setData() {

    try {

        const ids = [];

        const currentTime = new Date().getTime();
        const monthAgo = subDays(currentTime, 1);

        const dataResult = await axios.get(API_URL + 'api/aviz/locations/');
        const data = dataResult.data.codAviz || [];

        for (let i = 0; i < data.length; i++) {
            const id = data[i];
            if (id.substring(0, 2) == "DC") {
                ids.push(id);
            }
        }
        // const result = {
        //   "poze": [],
        //   "hasFinishedTransport": false,
        //   "codAviz": "DC23008153004301900809201017",
        //   "nrIdentificare": "TLLU6866637",
        //   "nrApv": null,
        //   "provenienta": "Depozit Ferco Capital punct de lucru",
        //   "tipAviz": 2,
        //   "numeTipAviz": "Depozit",
        //   "volum": {
        //     "idAviz": 22029776,
        //     "volumSpecie": [
        //       {
        //         "numeSpecie": "RĂȘINOASE",
        //         "volumSortiment": [
        //           {
        //             "numeSortiment": "Cherestele",
        //             "volum": 43.328316
        //           }
        //         ],
        //         "volum": 43.328316
        //       }
        //     ],
        //     "total": 43.328316
        //   },
        //   "emitent": {
        //     "denumire": "Ferco Capital S.A.",
        //     "cui": "31471091"
        //   },
        //   "transportator": {
        //     "identificator": null,
        //     "denumire": null,
        //     "cui": null,
        //     "tip": "FARA",
        //     "nonSumal": true
        //   },
        //   "beneficiar": {
        //     "tip": "JURIDICA"
        //   },
        //   "destinatar": {
        //     "tip": "JURIDICA"
        //   },
        //   "valabilitate": {
        //     "emitere": 1695194222329,
        //     "finalizare": 1697786183964
        //   },
        //   "tipMarfa": null,
        //   "angajat": null,
        //   "remorca": null,
        //   "identificatorContainer": "",
        //   "marfa": {
        //     "grupeSpecii": "RĂȘINOASE",
        //     "specii": "MOLID ",
        //     "sortimente": "Cherestele",
        //     "total": 43.328316
        //   }
        // }


        const content = {};
        const header = [
            'codAviz',
            'emitent_denumire',
            'emitent_cui',
            'provenienta',
            'marfa_grupeSpecii',
            'marfa_specii',
            'marfa_sortimente',
            'marfa_total',
            'nrIdentificare',
            'transportator_denumire',
            'transportator_cui',
            'destination',
            'valabilitate_emitere',
            'valabilitate_finalizare',
            'volum_volumSpecie'
        ]

        for (let i = 0; i < ids.length; i++) {

            // console.log('count:', i);
            // if (i == 10) break;

            try {
                const response = await axios.get(API_URL + 'api/aviz/' + ids[i]);
                const result = response.data;
                const valabilitate = result.valabilitate;
                const emitere = valabilitate?.emitere;
                const finalizare = valabilitate?.finalizare;

                const before = isBefore(emitere, monthAgo);

                if (before) continue;

                const dateStr = format(emitere, "dd-MM-yyyy", { timeZone: 'UTC+1' });
                if (!content[dateStr]) {
                    content[dateStr] = [];
                }

                let destination = ''
                try {
                    destination = result.nrIdentificare ? (await axios.get(`https://api.maersk.com/synergy/tracking/${result.nrIdentificare}?operator=MAEU`)).data.destination.city : '';
                } catch (error) {

                }

                const item = [
                    result?.codAviz || '',
                    result?.emitent?.denumire || '',
                    result?.emitent?.cui || '',
                    result?.provenienta || '',
                    result?.marfa?.grupeSpecii || '',
                    result?.marfa?.specii || '',
                    result?.marfa?.sortimente || '',
                    result?.marfa?.total || '',
                    result?.nrIdentificare || '',
                    result?.transportator?.denumire || '',
                    result?.transportator?.cui || '',
                    destination,
                    format(result?.valabilitate?.emitere, "dd/MM/yyyy") || '',
                    format(result?.valabilitate?.finalizare, "dd/MM/yyyy") || '',
                    JSON.stringify(result?.volum?.volumSpecie) || ''
                ]

                content[dateStr].push(item)

            } catch (error) {
                console.log('Axios error:', error);
            }

        }

        for (i = 30; i >= 0; i--) {

            const sheetName = format(subDays(currentTime, i).getTime(), "dd-MM-yyyy");
            const sorted = content[sheetName] ? content[sheetName].sort((a, b) => a[13] > a[14] ? 1 : -1) : [];

            const worksheet = XLSX.utils.aoa_to_sheet([header].concat(sorted));
            XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);
        }

        XLSX.writeFile(workbook, 'result.xlsx');

    } catch (error) {
        console.log('error:', error);
    }

}

setData();
