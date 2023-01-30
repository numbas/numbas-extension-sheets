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
    spreadsheet.setAttribute('disabled',true);
    main.appendChild(spreadsheet);

    const d = await (await fetch('examples/level-pegs-on-water-main.xlsx')).arrayBuffer();
    const wb = XLSX.read(d, {sheetStubs:true});
    const sheet = Object.values(wb['Sheets'])[0];
    spreadsheet.load_worksheet(sheet)

}

init_app();
