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
            student_answer: function(part) { return part.studentAnswer; },
            load: function(part, data) { return data.answer; }
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
                    const value = row_values[col-range.s.c];
                    if(value === undefined) {
                        continue;
                    }
                    if(ws[ref]===undefined) {
                        ws[ref] = {t:'z',v:''};
                    }
                    let t = 's';
                    if(typeof value == 'number') {
                        t = 'n';
                    }
                    ws[ref].t = t;
                    ws[ref].v = value;
                }
            }
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

    }

    class TSpreadsheet {
        constructor(wb) {
            this.value = wb;
        }
    }
    
    jme.registerType(
        TSpreadsheet,
        'spreadsheet',
        {
            'list': function(s) {
                // TODO
            },
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
            // TODO
            return 'spreadsheet()';
        }
    });

    const {TString, TList, TDict} = Numbas.jme.types;

    sheets.scope.addFunction(new jme.funcObj('spreadsheet',['list of list'],TSpreadsheet, 
        (content) => {
            return new TSpreadsheet(new Workbook(XLSX.utils.aoa_to_sheet(content)));
        },
        { unwrapValues: true}
    ));

    sheets.scope.addFunction(new jme.funcObj('spreadsheet_from_base64_file', ['string', 'string'], TSpreadsheet,
        (filename, base64) => {
            const data = decode_array(base64);
            return new TSpreadsheet(new Workbook(XLSX.read(data, {sheetStubs: true})));
        },
        { unwrapValues: true }
    ))

    sheets.scope.addFunction(new jme.funcObj('spreadsheet_from_workbook',['dict'], TSpreadsheet, 
        (workbook) => {
            return new TSpreadsheet(new Workbook(workbook));
        },
        {unwrapValues: true}
    ));

    sheets.scope.addFunction(new jme.funcObj('update_range',[TSpreadsheet,TString,TDict], TSpreadsheet,
        (spreadsheet, range_string, changes) => {
            return new TSpreadsheet(spreadsheet.update_range(range_string, changes));
        },
        {unwrapValues: true}
    ));

    sheets.scope.addFunction(new jme.funcObj('update_range',[TSpreadsheet,'list of string',TDict], TSpreadsheet,
        (spreadsheet, range_strings, changes) => {
            let wb = spreadsheet;
            for(let range_string of range_strings) {
                wb = wb.update_range(range_string, changes);
            }
            return new TSpreadsheet(wb);
        },
        {unwrapValues: true}
    ));

    sheets.scope.addFunction(new jme.funcObj('disable_cells',[TSpreadsheet,'list of string'], TSpreadsheet,
        (spreadsheet, range_strings) => {
            let wb = spreadsheet;
            for(let range_string of range_strings) {
                wb = wb.update_range(range_string, {'disabled':true});
            }
            return new TSpreadsheet(wb);
        },
        {unwrapValues: true}
    ));

    sheets.scope.addFunction(new jme.funcObj('fill_range',[TSpreadsheet,TString,'list of (string or number)'], TSpreadsheet,
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
        {unwrapValues: true}
    ));
    sheets.scope.addFunction(new jme.funcObj('fill_range',[TSpreadsheet,TString,'list of list of (string or number)'], TSpreadsheet,
        (spreadsheet, range, values) => {
            return new TSpreadsheet(spreadsheet.fill_range(range,values));
        },
        {unwrapValues: true}
    ));
    sheets.scope.addFunction(new jme.funcObj('fill_ranges',[TSpreadsheet,'list'], TSpreadsheet,
        (spreadsheet, values) => {
            let s = spreadsheet;
            for(let v of values) {
                const range = v[0];
                const values = v[1];
                s = s.fill_range(range, values);
            }
            return new TSpreadsheet(s);
        },
        {unwrapValues: true}
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
        }
    ));

    sheets.scope.addFunction(new jme.funcObj('slice', [TSpreadsheet, TString], TSpreadsheet, (wb,range) => {
        return wb.slice(range);
    }));

    sheets.scope.addFunction(new jme.funcObj('listval', [TSpreadsheet, TString], '?', null, {
        evaluate: function(args, scope) {
            const wb = args[0].value;
            let ref = args[1].value;
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
                    out.push(new TList(orow));
                    for(let col=range.s.c; col<=range.e.c; col++) {
                        const v = wb.get_cell_value(sheet, XLSX.utils.encode_cell({r:row,c:col}));
                        orow.push(new TString((v || '')+''));
                    }
                }
                return new TList(out);
            } else {
                return new TString((wb.get_cell_value(sheet,ref) || '')+'');
            }
        }
    }));

    sheets.scope.addFunction(new jme.funcObj('encode_range',['integer','integer','integer','integer'], TString, (cs,rs,ce,re) => {
        return XLSX.utils.encode_range({s:{c:cs,r:rs}, e:{c:ce,r:re}});
    }));

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
