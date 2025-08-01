Numbas.addExtension('sheets', ['display', 'util', 'jme','sheet-element', 'xlsx'], function(sheets) {
    const jme = Numbas.jme;

    /** Encode a UInt8Array as a base64 string.
     *
     * @param {UInt8Array} data
     * @returns {string}
     */
    function encode_array(data) {
        return btoa(String.fromCharCode(...new Uint8Array(data)));
    }

    /** Decode a base64-encoded string to an UInt8Array.
     *
     * @param {string} base64
     * @returns {UInt8Array}
     */
    function decode_array(base64) {
        return Uint8Array.from(atob(base64), c => c.charCodeAt(0))
    }

    /** A spreadsheet editor widget for a part.
     *
     * @param {Element} element - The parent element of the widget.
     * @param {Numbas.parts.Part} part - The part whose answer the widget represents.
     * @param {string} title - The `title` attribute for the widget: a text description of what the widget represents.
     * @param {Object.<Function>} events - Callback functions for events triggered by the widget.
     * @param {Numbas.answer_widgets.answer_changed} answer_changed - A function to call when the entered answer changes.
     * @param {Object} options - Any options for the widget.
     */
    class SpreadsheetWidget {
        constructor(element, part, title, events, answer_changed, options) {
            var w = this;
            this.part = part;
            this.title = title;
            this.events = events;
            this.answer_changed = answer_changed;
            this.options = options;

            var sheet = this.sheet = document.createElement('spread-sheet');
            element.appendChild(sheet);

            sheet.addEventListener('sheetchange', e => {
                if(!this.first_change) {
                    this.first_change = true;
                    return;
                }
                if(this.setting) {
                    this.setting = false;
                    return;
                }
                this.changing = true;
                const wb = this.options.initial_sheet.copy();
                wb.replace_worksheet(e.detail.sheet);
                w.answer_changed({valid: true, value: wb});
            });
          
          	if(this.options.initial_sheet) {
              sheet.load_worksheet(this.options.initial_sheet.get_worksheet());
            }

            if(w.events) {
                for(var x in w.events) {
                    sheet.addEventListener(x, e => w.events[x]({},e));
                }
            }
        }

        setAnswerJSON(answerJSON) {
            if(answerJSON.value === undefined) {
                return;
            }
            if(this.changing) {
                this.changing = false;
                return;
            }
            this.setting = true;
            this.sheet.load_worksheet(answerJSON.value.get_worksheet());
        }

        disable() {
            this.sheet.setAttribute('disabled',true);
        }

        enable() {
            this.sheet.removeAttribute('disabled');
        }
    }

    Numbas.answer_widgets.register_custom_widget({
        name: 'spread-sheet',
        niceName: 'Spreadsheet',
        widget: SpreadsheetWidget,
        signature: 'spreadsheet',
        answer_to_jme: function(answer) {
            return new TSpreadsheet(answer);
        },
        options_definition: [
          {
            name: 'initial_sheet',
            label: 'Initial sheet',
            input_type: 'anything',
            default_value: ''
          }
        ],
        scorm_storage: {
            interaction_type: function(part) { return 'fill-in'; },
            correct_answer: function(part) { return part.input_options().correctAnswer; },
            student_answer: function(part) { return part.studentAnswer ? JSON.stringify(part.studentAnswer.wb) : ''; },
            load: function(part, data) { 
                if(data.answer) {
                    const wb = JSON.parse(data.answer);
                    return new Workbook(wb);
                } else {
                    return undefined;
                }
            }
        }
    });

    function cells_equal(a,b) {
        return a !== undefined && b !== undefined && a.r==b.r && a.c==b.c;
    }

    class Workbook {
        constructor(wb) {
            this.wb = wb;
        }

        default_sheetname() {
            if(this.wb && this.wb.SheetNames) {
                return this.wb.SheetNames[0];
            }
        }

        get_worksheet(name) {
            if(!this.wb.Sheets) {
                return;
            }
            return this.wb.Sheets[name || this.default_sheetname()];
        }

        replace_worksheet(sheet, name) {
            name = name || this.default_sheetname()
            this.wb.Sheets[name] = sheet;
        }

        worksheet_names() {
            return this.wb?.SheetNames || [];
        }

        copy() {
            return new Workbook(JSON.parse(JSON.stringify(this.wb)));
        }

        find_corner(ws,ref) {
            if(ref in ws) {
                return ref;
            }
            const cell = XLSX.utils.decode_cell(ref);
            const row = cell.r;
            const col = cell.c;

            if('!merges' in ws) {
                const corner = ws['!merges'].find(range => {
                    const r = typeof range == 'string' ? XLSX.utils.decode_range(range) : range;
                    if(r.s.c<=col && r.s.r<=row && r.e.c>=col && r.e.r>=row) {
                        return XLSX.utils.encode_cell(r.s);
                    }
                });

                if(corner) {
                    return corner;
                }
            }
            return ref;
        }

        update_range(range_string,changes) {
            let wb = this.copy();
            const ws = wb.get_worksheet();
            const range = XLSX.utils.decode_range(range_string);
            for(let row=range.s.r; row<=range.e.r; row++) {
                for(let col=range.s.c; col<=range.e.c; col++) {
                    const ref = XLSX.utils.encode_cell({r:row, c:col});
                    let corner;
                    if(ref in ws) {
                        corner = ref;
                    } else if('!merges' in ws) {
                        corner = this.find_corner(ws,ref);
                    } else {
                        corner = ref;
                    }
                    if(!ws[corner]) {
                        ws[corner] = {t:'z',v:''};
                    }

                    Numbas.util.deep_extend_object(ws[corner],changes);
                }
            }
            return wb;
        }

        fill_range(range_string,values) {
            let wb = this.copy();
            const ws = wb.get_worksheet();
            const range = XLSX.utils.decode_range(range_string);
            for(let row=range.s.r;row<=range.e.r;row++) {
                const row_values = values[row-range.s.r];
                if(row_values === undefined) {
                    continue;
                }
                for(let col=range.s.c;col<=range.e.c;col++) {
                    const ref = this.find_corner(ws,XLSX.utils.encode_cell({r:row,c:col}));
                    let value = row_values[col-range.s.c];
                    if(value === undefined) {
                        continue;
                    }
                    if(ws[ref]===undefined) {
                        ws[ref] = {t:'z',v:''};
                    }
                    let t = 's';
                    if(typeof value == 'number') {
                        value = Numbas.math.niceNumber(value);
                    }
                    ws[ref].t = t;
                    ws[ref].v = value;
                }
            }
            const sheet_range = XLSX.utils.decode_range(ws['!ref']);
            ws['!ref'] = XLSX.utils.encode_range({s: {c:Math.min(range.s.c, sheet_range.s.c), r: Math.min(range.s.r, sheet_range.s.r)}, e: {c:Math.max(range.e.c, sheet_range.e.c), r: Math.max(range.e.r, sheet_range.e.r)}});
            return wb;
        }

        /** Return a new workbook containing only the given range.
         */
        slice(range_string) {
            const wb = this.copy();
            const ws = wb.get_worksheet();
            const range = XLSX.utils.decode_range(range_string);
            const os = {'!ref': range_string, '!merges': ws['!merges']};
            Object.entries(ws).forEach(([k,v]) => {
                const p = XLSX.utils.decode_cell(k);
                if(p.c<range.s.c || p.c>range.e.c || p.r<range.s.r || p.r>range.e.r) {
                    return;
                }
                os[k] = v;
            });
            wb.wb['Sheets'][wb.wb.SheetNames[0]] = os;
            return wb;
        }

        get_named_ref(name) {
            const names = this.wb.Workbook?.Names;
            if(!names) {
                return;
            }
            const ref = names.find(({Name}) => Name==name)?.Ref;
            return ref;
        }

        get_cell_value(sheet, ref) {
            return sheet[ref]?.v;
        }

        to_base64() {
            const buffer = XLSX.write(this.wb, {type: 'array', bookType: 'xlsx'});
            return encode_array(buffer);
        }
    }

    const {TString, TList, TDict, THTML} = Numbas.jme.types;

    class TSpreadsheet {
        constructor(wb) {
            this.value = wb;
        }

        /** Resolve a named reference or a string representation of a range to a 2D array of values or the value in a single cell.
         *
         * @param {string} ref
         * @returns {object} with either a property `range: Array.<Array.<string>>>` or `cell: string`
         */
        resolve_ref(ref) {
            const wb = this.value;
            let range;

            const named_ref = wb.get_named_ref(ref);
            if(named_ref) {
                range = XLSX.utils.dereference_range(named_ref);
                if(cells_equal(range.e,range.s)) {
                    delete range.e;
                }
                if(range.e === undefined) {
                    ref = XLSX.utils.encode_cell(range.s);
                }
            } else if(ref.match(/:/)) {
                range = XLSX.utils.decode_range(ref);
            } else {
                range = {s:XLSX.utils.decode_cell(ref)};
            }

            const sheet = wb.get_worksheet(range.sheet);

            if(range.e !== undefined) {
                const out = [];
                for(let row=range.s.r; row<=range.e.r; row++) {
                    const orow = [];
                    out.push(orow);
                    for(let col=range.s.c; col<=range.e.c; col++) {
                        const v = wb.get_cell_value(sheet, XLSX.utils.encode_cell({r:row,c:col}));
                        orow.push((v || '')+'');
                    }
                }
                return {range: out};
            } else {
                return {cell: (wb.get_cell_value(sheet,ref) || '')+''};
            }
        }
    };
    
    jme.registerType(
        TSpreadsheet,
        'spreadsheet',
        {
            'html': function(s) {
                const e = document.createElement('spread-sheet');
                e.setAttribute('disabled',true);
                e.load_worksheet(s.value.get_worksheet());
                return new jme.types.THTML(e);
            },
            'dict': function(s) {
                return jme.wrapValue(s.value);
            }
        }
    );
    Numbas.util.equalityTests['spreadsheet'] = function(a,b) {
        return false;
    };

    jme.display.registerType(TSpreadsheet,{
        tex: (thing,tok,texArgs) => {
            // TODO
            return '\\text{Spreadsheet}';
        },
        jme: (tree,tok,bits) => {
            return 'spreadsheet_from_base64_file("sheet.xlsx",safe("' + Numbas.jme.escape(tok.value.to_base64())+'"))';
        }
    });

    sheets.scope.addFunction(new jme.funcObj('spreadsheet',[],TSpreadsheet, 
        (content) => {
            const sheet = XLSX.utils.aoa_to_sheet([['']]);
            const workbook = {
                Sheets: { 'Sheet1': sheet },
                SheetNames: ['Sheet1']
            };
            return new TSpreadsheet(new Workbook(workbook));
        },
        { unwrapValues: true, random: false }
    ));

    sheets.scope.addFunction(new jme.funcObj('spreadsheet',['list of list'],TSpreadsheet, 
        (content) => {
            const sheet = XLSX.utils.aoa_to_sheet(content);
            const workbook = {
                Sheets: { 'Sheet1': sheet },
                SheetNames: ['Sheet1']
            };
            return new TSpreadsheet(new Workbook(workbook));
        },
        { unwrapValues: true, random: false }
    ));

    sheets.scope.addFunction(new jme.funcObj('spreadsheet_from_base64_file', ['string', 'string'], TSpreadsheet,
        (filename, base64) => {
            const data = decode_array(base64);
            return new TSpreadsheet(new Workbook(XLSX.read(data, {sheetStubs: true})));
        },
        { unwrapValues: true, random: false }
    ));

    sheets.scope.addFunction(new jme.funcObj('update_range',[TSpreadsheet,TString,TDict], TSpreadsheet,
        (spreadsheet, range_string, changes) => {
            return new TSpreadsheet(spreadsheet.update_range(range_string, changes));
        },
        { unwrapValues: true, random: false }
    ));

    sheets.scope.addFunction(new jme.funcObj('update_range',[TSpreadsheet,'list of string',TDict], TSpreadsheet,
        (spreadsheet, range_strings, changes) => {
            let wb = spreadsheet;
            for(let range_string of range_strings) {
                wb = wb.update_range(range_string, changes);
            }
            return new TSpreadsheet(wb);
        },
        { unwrapValues: true, random: false }
    ));

    sheets.scope.addFunction(new jme.funcObj('disable_cells',[TSpreadsheet,'list of string'], TSpreadsheet,
        (spreadsheet, range_strings) => {
            let wb = spreadsheet;
            for(let range_string of range_strings) {
                wb = wb.update_range(range_string, {'disabled':true});
            }
            return new TSpreadsheet(wb);
        },
        { unwrapValues: true, random: false }
    ));

    sheets.scope.addFunction(new jme.funcObj('fill_range',[TSpreadsheet, TString, 'list of (string or number)'], TSpreadsheet,
        (spreadsheet, range, values) => {
            const drange = XLSX.utils.decode_range(range);
            if(drange.e.c == drange.s.c) {
                values = values.map(x => [x]);
            } else if(drange.e.r == drange.s.r) {
                values = [values];
            } else {
                throw(new Numbas.Error("The values for a 2D range must be a list of lists."));
            }
            return new TSpreadsheet(spreadsheet.fill_range(range,values));
        },
        { unwrapValues: true, random: false }
    ));
    sheets.scope.addFunction(new jme.funcObj('fill_range',[TSpreadsheet,TString,'list of list of (string or number)'], TSpreadsheet,
        (spreadsheet, range, values) => {
            return new TSpreadsheet(spreadsheet.fill_range(range,values));
        },
        { unwrapValues: true, random: false }
    ));

    /** Interpret a range string and return a list of the cells it contains.
     */
    sheets.scope.addFunction(new jme.funcObj('parse_range',[TString],TList,
        (range_string) => {
            const range = XLSX.utils.decode_range(range_string);
            const out = [];
            for(let r=range.s.r; r<=range.e.r; r++) {
                for(let c=range.s.c; c<=range.e.c; c++) {
                    out.push(new TString(XLSX.utils.encode_cell({r,c})));
                }
            }
            return out;
        },
        { random: false }
    ));

    sheets.scope.addFunction(new jme.funcObj('slice', [TSpreadsheet, TString], TSpreadsheet, (wb,range) => {
        return wb.slice(range);
    }, {random: false}));

    sheets.scope.addFunction(new jme.funcObj('listval', [TSpreadsheet, TString], '?', null, {
        evaluate: function(args, scope) {
            const sheet = args[0];
            const ref = args[1].value;
            const result = sheet.resolve_ref(ref);
            if(result.range) {
                return jme.wrapValue(result.range);
            } else if(result.cell !== undefined) {
                return new TString(result.cell);
            }
        },
        random: false
    }));

    sheets.scope.addFunction(new jme.funcObj('range_as_numbers', [TSpreadsheet, TString, '[string or list of string]'], '?', null, {
        evaluate: function(args, scope) {
            const sheet = args[0];
            const ref = args[1].value;
            const notationStyle = jme.unwrapValue(args[2]);

            const result = sheet.resolve_ref(ref);

            function parsenumber(s) {
                return Numbas.util.parseNumber(s, false, notationStyle, true);
            }
            if(result.range) {
                return jme.wrapValue(result.range.map(row => row.map(parsenumber)));
            } else if(result.cell) {
                return new TNum(parsenumber(result.cell));
            }
        },
        random: false
    }));

    sheets.scope.addFunction(new jme.funcObj('encode_range',['integer','integer','integer','integer'], TString, (cs,rs,ce,re) => {
        return XLSX.utils.encode_range({s:{c:cs,r:rs}, e:{c:ce,r:re}});
    }, {random: false}));

    /** 
     * Create an ArrayBuffer containing the given string.
     *
     * @param {string} s
     * @returns ArrayBuffer
     */
    function string_to_buffer(str) {
        const buf = new ArrayBuffer(str.length);
        const view = new Uint8Array(buf);
        for (let i = 0; i < str.length; i++) {
            view[i] = str.charCodeAt(i) & 0xFF;
        }
        return buf;
    }
    
    sheets.scope.addFunction(new jme.funcObj('download_sheet', [TSpreadsheet, TString], THTML, (sheet, filename) => {
        const {wb} = sheet;

        var wbout = XLSX.write(wb, { type: 'array', bookType: 'xlsx' });

        const blob = new Blob([wbout], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        const url = URL.createObjectURL(blob);

        const link = document.createElement('a');
        link.href = url;
        link.target = '_blank';
        link.download = filename;
        link.innerHTML = `Download <code>${filename}</code>`;

        return link;
    }, {random: false}));

    function add_style_function(name,args,fn) {
        args = args.slice();
        sheets.scope.addFunction(new jme.funcObj(name, args, TDict, function(...args) {
            var style = fn(...args);
            return jme.wrapValue({style: style});
        },{unwrapValues: true, random: false}));

        var dargs = args.slice();
        dargs.splice(0,0,'dict');
        sheets.scope.addFunction(new jme.funcObj(name, dargs, TDict, function(pd,...args) {
            var style = fn(...args);
            return jme.wrapValue(Numbas.util.deep_extend_object(pd,{style: style}));
        },{unwrapValues: true, random: false}));
    }

    function css_color_to_rgba(color) {
        const d = document.createElement('div');
        document.body.appendChild(d);
        d.style.fill = color;
        const rgba = window.getComputedStyle(d).fill;
        document.body.removeChild(d);
        let m;
        if(m = rgba.match(/rgba\((.*),(.*),(.*),(.*)\)/)) {
            const [r,g,b] = m.slice(1).map(x => (Math.round(parseFloat(x)) % 256).toString(16).padStart(2,'0'));
            const a = Math.round(Math.min(1,parseFloat(m[4])) * 255).toString(16).padStart(2,'0');
            return a+r+g+b;
        } else if(m = rgba.match(/rgb\((.*),(.*),(.*)\)/)) {
            const [r,g,b] = m.slice(1).map(x => (Math.round(parseFloat(x)) % 256).toString(16).padStart(2,'0'));
            return r+g+b;
        }
    }
    sheets.css_color_to_rgba = css_color_to_rgba;

    add_style_function('border',['string','string'], (sides, style) => {
        sides = sides.split(/\s+/);

        const s = {};
        const style_bits = style.toLowerCase().split(/\s+/);
        for(let bit of style_bits) {
            if(bit.match(/thin|medium|thick/i)) {
                s['style'] = bit;
            } else {
                s['color'] = {css: bit};
            }
        }
        return {border: Object.fromEntries(sides.map(side => [side,s]))};
    });
    add_style_function('font_style',['string'], (style) => {
        const bits = style.toLowerCase().split(/\s+/);
        const o = {};
        for(let style of ['bold','italic','underline']) {
            if(bits.contains(style)) {
                o[style] = true;
            }
        }
        return {font: o};
    });
    add_style_function('font_family', ['string'], (font) => { return {font: {name: font}} });
    add_style_function('font_size', ['number'], (size) => { return {font: {sz: size*11}} });
    add_style_function('font_color', ['string'], (color) => { return {font: {color: {css: color}}} });
    add_style_function('bg_color', ['string'], (color) => { return {fill: {fgColor: {css: color} } } });
    add_style_function('horizontal_alignment',['string'], (alignment) => { return {alignment: {horizontal: alignment}} });
    add_style_function('vertical_alignment',['string'], (alignment) => { return {alignment: {vertical: alignment}} });

    sheets.scope.addFunction(new jme.funcObj('cell_type', ['string'], TDict, function(type) {
        var types_map = {
            'text': 's',
            'number': 'n'
        };
        return jme.wrapValue({t: types_map[type] || type});
    },{unwrapValues: true, random: false}));

    if(Numbas.editor?.register_variable_template_type !== undefined) {
        class SpreadsheetVariableTemplateWidget extends HTMLElement {
            static get observedAttributes() { return ['value']; }

            constructor() {
                super();
                this.attachShadow({mode:'open'});

                const template = `
                <style>
                    .container:not(.got-value) .when-got-value, .container.got-value .when-no-value {
                        display: none;
                    }
                </style>
                <div class="container">
                    <p id="got-value-message" class="when-got-value"><code id="filename">Spreadsheet</code> <a id="download">Download</a></p>
                    <label for="file-upload"><span class="when-no-value">Upload a file</span><span class="when-got-value">Replace with a different file</span>:</label> <input type="file" id="file-upload">
                </div>
                `;
                this.shadowRoot.innerHTML = template;

                const input = this.shadowRoot.querySelector('#file-upload');

                input.addEventListener('change', async e => {
                    const [file] = input.files;
                    if(!file) {
                        return;
                    }

                    const data = await file.arrayBuffer();
                    const file_info = {
                        filename: file.name,
                        base64: encode_array(data)
                    };
                    this.set_value(file_info);

                    this.dispatchEvent(new CustomEvent('change', {detail: {value: file_info}}));
                });
            }

            set_value(file_info) {
                const filename_display = this.shadowRoot.querySelector('#got-value-message #filename');
                const download_link = this.shadowRoot.querySelector('#got-value-message #download');
                this.value = file_info;
                if(file_info) {
                    filename_display.textContent = file_info.filename;
                    const array = decode_array(file_info.base64);
                    download_link.setAttribute('href', URL.createObjectURL(new Blob([array])));
                    download_link.setAttribute('download', file_info.filename);
                } else {
                    filename_display.textContent = `Nothing`;
                    download_link.removeAttribute('href');
                    download_link.removeAttribute('download');
                }
                this.shadowRoot.querySelector('.container').classList.toggle('got-value', !!file_info);
            }
        }
        window.customElements.define('variable-template-spreadsheet', SpreadsheetVariableTemplateWidget);

        Numbas.editor.register_variable_template_type(function(value) {
            return {
                id: 'spreadsheet', 
                name: 'Spreadsheet file',
                value: value,
                load_definition(definition) {
                    var tree = Numbas.jme.compile(definition);
                    if(!tree.args || tree.tok.type == 'nothing') {
                        this.value(null);
                    } else {
                        var filename = Numbas.jme.builtinScope.evaluate(tree.args[0]).value;
                        var base64 = Numbas.jme.builtinScope.evaluate(tree.args[1]).value;
                        this.value({filename, base64});
                    }
                },
                jme_definition() {
                    const v = this.value();
                    if(v === null) {
                        return 'nothing';
                    }
                    return 'spreadsheet_from_base64_file(safe("'+Numbas.jme.escape(v.filename)+'"), safe("'+Numbas.jme.escape(v.base64)+'"))'
                },
                widget: SpreadsheetVariableTemplateWidget
            }
        });

    }
});
