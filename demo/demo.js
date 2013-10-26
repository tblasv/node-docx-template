var docxTemplate = require('../docxTemplate');
    fs = require('fs');

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

/* async */
docxTemplate.render('demo/template.docx', templateData1, 'demo/file1_gen.docx', function(err) {
    console.log('on complete');
});

/* sync */
docxTemplate.renderSync('demo/template.docx', templateData2, 'demo/file2_gen.docx');


var template = docxTemplate.compileFileSync('demo/template.docx'), // docxTemplate.compile(/* data [,opt] *!/)
    docxFile = template(templateData1); // render


/* async */
docxTemplate.compileFile('demo/template.docx', function(err, aTemplate){
    if(err) return console.log(err);
    fs.writeFile('demo/file3_gen.docx', aTemplate(templateData2), { encoding: 'binary' }, function(err){
        if(err) console.log(err);
    });
});