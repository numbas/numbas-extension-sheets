DIRNAME=$(notdir $(CURDIR))
EDITOR_EXTENSION_PATH = ~/numbas/editor/media/user-extensions/extracted/47/sheets

ELMS=$(wildcard src/*.elm)

extension: sheets.js $(EDITOR_EXTENSION_PATH)/style.css

sheets.js: src/extension.js dist/sheet_element.js dist/xlsx.js
	cat $^ > $@
	cp $@ $(EDITOR_EXTENSION_PATH)

$(EDITOR_EXTENSION_PATH)/style.css: style.css
	cp $^ $@

dist/xlsx.js: src/xlsx.js
	echo "Numbas.queueScript('xlsx',[],function() {" > $@
	cat $^ >> $@
	echo "})" >> $@

dist/sheet_element.js: dist/app.js src/sheet_element.js
	echo "Numbas.queueScript('sheet-element',['xlsx'], function(exports) {" > $@
	cat $^ >> $@
	echo "})" >> $@

dist/app.js: src/Spreadsheet.elm $(ELMS)
	-elm make $< --output=$@ 2> error.txt
	@sed -i "s/;}(this));$$/;}(globalThis));\n\n/" $@ # Fix the compiled JS so it uses globalThis, allowing it to be used as a module file
	@cat error.txt
