DIRNAME=$(notdir $(CURDIR))
EDITOR_EXTENSION_PATH = ~/numbas/editor/media/user-extensions/extracted/47/sheets

ELMS=$(wildcard src/*.elm)

extension: $(EDITOR_EXTENSION_PATH)/sheets.js $(EDITOR_EXTENSION_PATH)/style.css

$(EDITOR_EXTENSION_PATH)/sheets.js: dist/sheets.js
	cp $^ $@

dist/sheets.js: src/extension.js lib/sheet_element.js lib/xlsx.js
	cat $^ > $@

$(EDITOR_EXTENSION_PATH)/style.css: style.css
	cp $^ $@

lib/xlsx.js: src/xlsx.js
	echo "Numbas.queueScript('xlsx',[],function() {" > $@
	cat $^ >> $@
	echo "})" >> $@

lib/sheet_element.js: lib/app.js src/sheet_element.js
	echo "Numbas.queueScript('sheet-element',['xlsx'], function(exports) {" > $@
	cat $^ >> $@
	echo "})" >> $@

lib/app.js: src/Spreadsheet.elm $(ELMS)
	-elm make --optimize $< --output=$@ 2> error.txt
	@sed -i "s/;}(this));$$/;}(globalThis));\n\n/" $@ # Fix the compiled JS so it uses globalThis, allowing it to be used as a module file
	@cat error.txt
