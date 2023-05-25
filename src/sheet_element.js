function boolish_attribute(v) {
    return v !== null && ['false','no'].indexOf(v.toLowerCase().trim()) == -1;
}

class SheetElement extends HTMLElement {
    static get observedAttributes() { return ['disabled']; }

    constructor() {
        super();
        //this.attachShadow({mode:'open'});

        this.app_created = new Promise((resolve,reject) => {
            this.resolve_app_created = resolve;
        })
    }

    connectedCallback() {
        if(!this.app) {
            this.try_create();
        }
    }

    attributeChangedCallback(name, oldValue, newValue) {
        if(name == 'disabled') {
            this.disabled_changed(newValue);
        }
    }

    async disabled_changed(disabled) {
        await this.app_created;
        this.app.ports.receive_attributes.send({disabled: boolish_attribute(disabled)});
    }

    try_create() {
        const root = this;

        /*
        const link = document.createElement('link');
        link.setAttribute('rel','stylesheet');
        link.setAttribute('href',Numbas.getStandaloneFileURL('sheets','style.css'));
        shadowRoot.appendChild(link);
        */

        const container = document.createElement('div');
        root.appendChild(container)
        root.addEventListener('mousemove', e => {
            let el = e.target;
            while(el && !el.hasAttribute('data-annotated-mousemove')) {
                el = el.parentElement;
            }
            if(!el) {
                return;
            }
            const box = el.getBoundingClientRect();
            el.dispatchEvent(new CustomEvent('annotatedmousemove', {detail: 
                {x: e.x, y: e.y, box: box}
            }));
        });
        root.addEventListener('keydown', e => {
            const {target} = e;
            const {selectionStart, selectionEnd} = target;
            if(!target.classList.contains('input-value')) {
                return;
            }
            const collapsed = selectionStart == selectionEnd;
            const disabled = target.classList.contains('disabled');
            const atStart = disabled || (collapsed && selectionStart == 0);
            const atEnd = disabled || (collapsed && selectionEnd == target.value.length);
            switch(e.key) {
                case 'ArrowUp':
                    if(atStart) {
                        e.preventDefault();
                        target.blur();
                        target.dispatchEvent(new CustomEvent('cellup'));
                    }
                    break;
                case 'ArrowDown':
                    if(atEnd) {
                        e.preventDefault();
                        target.blur();
                        target.dispatchEvent(new CustomEvent('celldown'));
                    }
                    break;
                case 'ArrowLeft':
                    if(atStart) {
                        e.preventDefault();
                        target.blur();
                        target.dispatchEvent(new CustomEvent('cellleft'));
                    }
                    break;
                case 'ArrowRight':
                    if(atEnd) {
                        e.preventDefault();
                        target.blur();
                        target.dispatchEvent(new CustomEvent('cellright'));
                    }
                    break;
            }
        });

        const mutation_observer = new MutationObserver(mutations => {
            mutations.forEach(mutation => {
                if(mutation.target.getAttribute('data-input-active')) {
                    mutation.target.querySelector('.input-value').focus();
                }
            });
        });
        mutation_observer.observe(root, {childList: true, subtree: true, attributes: true, attributeFilter: ['data-input-active']});

        this.app = Elm.Spreadsheet.init({node: container, flags: {}});

        this.app.ports.send_spreadsheet.subscribe(data => {
            this.dispatchEvent(new CustomEvent('sheetchange', {detail: {sheet: data}}));
        });

        this.resolve_app_created(this.app);
    }

    async load_worksheet(data) {
        await this.app_created;
        this.app.ports.receive_spreadsheet.send(data);
    }
}

customElements.define('spread-sheet', SheetElement);
