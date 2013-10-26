## node-docx-template

Generate docx documents from docx template

## API

```javascript
var opt = {
    openTag: '(',
    closeTag: ')'
}
docxTemplate.settings(opt);

var template = docxTemplate.compile(docx [,opt]); /* compile templates */
template(data); /* render */

/* or */
docxTemplate.compileFile(pathToFile [,opt], cb);
docxTemplate.compileFileSync(pathToFile [,opt]);

docxTemplate.renderToFile(file, template, data, cb)

```