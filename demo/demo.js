var docxTemplate = require('../docxTemplate');
var fs = require('fs');

var templateData1 = {
    Title: 'Sync Test',
    header: 'https://github.com/tblasv/node-docx-template',
    name: 'tblasv',
    year: '2013',
    tag: 'simple text',
    inner: 'INNER TEXT'
};
var templateData2 = {
    Title: 'node-docx-template',
    header: 'Simple template',
    name: 'https://github.com/tblasv/node-docx-template',
    year: '2013'
};

var template = docxTemplate.compileFileSync('demo/template.docx');

/* async */
docxTemplate.compileFile('demo/template.docx', {openTag: '{', closeTag: '}'}, function(err, aTemplate){
    if(err) return console.log(err);
    fs.writeFile('demo/file2_gen.docx', aTemplate(templateData2), { encoding: 'binary' }, function(err){
        if(err) console.log(err);
    });
});

/* async */
docxTemplate.renderToFile('demo/file1_gen.docx', template, templateData1);


/* or manual

var mTemplate = docxTemplate.compile(/* data [,opt] *!/)
var docxFile = mTemplate(templateData1); // render
...
*/