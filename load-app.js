import show_error from './show-error.mjs';
import './dist/app.js';
import './src/sheet_element.js';

async function init_app() {
    const compilation_error = await show_error;
    if(compilation_error) {
        return;
    }
    const main = document.body.querySelector('main');
    main.innerHTML = '';
    const spreadsheet = document.createElement('spread-sheet');
    spreadsheet.setAttribute('disabled',false);
    main.appendChild(spreadsheet);

    spreadsheet.addEventListener('sheetchange', e => {
        const {sheet} = e.detail;
        console.log(sheet.A2);
    });

    const d = await (await fetch('examples/books-per-child.xlsx')).arrayBuffer();
    const wb = XLSX.read(d, {sheetStubs:true, cellNF: true});
    wb.Sheets.Sheet1.A2.style = {
        border: {
            left: {
                style: 'thick',
                color: {css: 'red'}
            }
        },
        font: {
            name: 'sans-serif',
            sz: 22,
            color: 'blue',
            bold: true,
            italic: true,
            underline: true
        },
        fill: {
            fgColor: {css: 'blue'}
        },
        alignment: {
            horizontal: 'left',
            vertical: 'top'
        }
    };
    wb.Sheets.Sheet1.A2.disabled = true;
    console.log(wb);
    const sheet = Object.values(wb['Sheets'])[0];
    spreadsheet.load_worksheet(sheet)

}

init_app();
