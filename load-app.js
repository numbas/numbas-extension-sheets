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
    main.appendChild(spreadsheet);

    const d = await (await fetch('examples/level-pegs-on-water-main.xlsx')).arrayBuffer();
    const wb = XLSX.read(d, {sheetStubs:true});
    const sheet = Object.values(wb['Sheets'])[0];
    console.log(wb);
    ['A1','B1','C1','D1','E1','F1','G1','G2','G15','A2','A16','A17','A18','A19'].forEach(c=>sheet[c].disabled = true);
    spreadsheet.load_worksheet(sheet)

}

init_app();
