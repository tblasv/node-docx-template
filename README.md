## node-docx-template

Generate docx documents from docx template

## API

```javascript

/* change brackets (optionaly) */
docxTemplate.openBracket = '@';
docxTemplate.closeBracket = '@';


var template = docxTemplate.compile(docx [,opt]); /* compile templates */
template(data); /* render */

/* or */
docxTemplate.compileFile(pathToFile [,opt], cb);
docxTemplate.compileFileSync(pathToFile [,opt]);

/* render */
docxTemplate.render(inputFile, data, outputFile, cb);
docxTemplate.rednerSync(inputFile, data, outputFile);
docxTemplate.renderToFile(file, template, data, cb)

```