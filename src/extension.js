Numbas.addExtension('sheets', ['display', 'util', 'jme','sheet-element', 'xlsx'], function(sheets) {
    const jme = Numbas.jme;

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
            sheet.setAttribute('editable',true);

            sheet.addEventListener('change', e => {
                this.changing = true;
                w.answer_changed({valid: true, value: e.detail.sheet});
            });
          
          	if(this.options.initial_sheet) {
              sheet.load_worksheet(this.options.initial_sheet);
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
            this.sheet.load_worksheet(answerJSON.value);
        }

        disable() {
            this.sheet.removeAttribute('editable');
        }

        enable() {
            this.sheet.setAttribute('editable',true);
        }
    }

    Numbas.answer_widgets.register_custom_widget({
        name: 'spread-sheet',
        niceName: 'Spreadsheet',
        widget: SpreadsheetWidget,
        signature: 'dict',
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

    class Workbook {
        constructor(wb) {
            this.wb = wb;
        }

        get_worksheet(name) {
            return this.wb.Sheets[name || this.wb.SheetNames[0]];
        }

        copy() {
            return new Workbook(JSON.parse(JSON.stringify(this.wb)));
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
                        corner = ws['!merges'].find(range => {
                            const r = XLSX.utils.decode_range(range);
                            if(r.s.c<=col && r.s.r<=row && r.e.c>=col && r.e.r>=row) {
                                return XLSX.utils.encode_cell(r.s);
                            }
                        });
                    }
                    if(!corner) {
                        corner = ref;
                        ws[ref] = {t:'z',v:''};
                    }

                    Object.assign(ws[ref],changes);
                }
            }
            return new Workbook(wb);
        }


        get_cell_value(ref) {
            const sheet = this.get_worksheet();
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
                e.load_worksheet(s.wb.get_worksheet());
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

    sheets.scope.addFunction(new jme.funcObj('listval', [TSpreadsheet, TString], '?', null, {
        evaluate: function(args, scope) {
            const wb = args[0].value;
            const ref = args[1].value;
            
            const sheet = wb.get_worksheet();

            if(ref.match(/:/)) {
                const range = XLSX.utils.decode_range(ref);
                const out = [];
                for(let row=range.s.r; row<=range.e.r; row++) {
                    const orow = [];
                    out.push(new TList(orow));
                    for(let col=range.s.c; col<=range.e.c; col++) {
                        const v = get_cell_value(sheet, XLSX.utils.encode_cell({r:row,c:col}));
                        orow.push(new TString(v));
                    }
                }
                return new TList(out);
            } else {
                return new TString(wb.get_cell_value(ref));
            }
        }
    }));
});
