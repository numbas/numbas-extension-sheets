DIRNAME=$(notdir $(CURDIR))
EDITOR_EXTENSION_PATH = ~/numbas/editor/media/user-extensions/extracted/47/sheets

ELMS=$(wildcard src/*.elm)

sheets.js: src/extension.js dist/sheet_element.js dist/xlsx.js 
	cat $^ > $@
	cp $@ $(EDITOR_EXTENSION_PATH)
	cp standalone_scripts/style.css $(EDITOR_EXTENSION_PATH)/standalone_scripts/style.css

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

upload: app.js index.html style.css
	rsync -avz . clpland:~/domains/somethingorotherwhatever.com/html/$(DIRNAME)
	@echo "Uploaded to https://somethingorotherwhatever.com/$(DIRNAME)"
