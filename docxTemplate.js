var fs = require('fs');
require('node-zip');

var reSpace = /\s*/,
    reFiles = /word\/(document|footer\d+|header\d+).xml/;

function compile(str, opt) {
    var openBracket = opt && opt.openBracket || module.exports.openBracket,
        closeBracket = opt && opt.closeBracket ||  module.exports.closeBracket,
        reKey = new RegExp('\\' + openBracket + '([^\\' + closeBracket + ']+)\\' + closeBracket);

    var keys = {},
        strParts = [],
        from = str.indexOf(openBracket),
        to = str.indexOf(closeBracket, from + 1),
        start = 0,
        key;

    while (from != -1 && to != -1) {
        var subStr = str.substring(start, from),
            keyPart = str.substring(from, to + closeBracket.length),
            xmlCloseTag = str.indexOf('</', to + closeBracket.length);

        strParts.push(subStr);

        if (str.substr(xmlCloseTag, 6) !== '</w:t>') {
            strParts.push(keyPart);
        } else {
            var reg = reKey.exec(keyPart);

            if (reg && reg[1].indexOf('<w:t>') === -1) {
                key = reg[1].replace(reSpace, '');
                addKey(key);
            } else {
                var hStr,
                    keyStart = 0,
                    keyEnd = keyPart.indexOf('</w:t>'),
                    keyIndex = -1;

                key = keyPart.substring(keyStart, keyEnd).replace(openBracket, '');

                if (key.length > 0) {
                    keyIndex = strParts.length;
                    addKey(key);
                }

                while (keyEnd !== -1) {
                    keyStart = keyEnd;
                    keyEnd = keyPart.indexOf('<w:t>', keyEnd + 6) + 5;
                    hStr =  keyPart.substring(keyStart, keyEnd);

                    strParts.push(hStr);

                    keyStart = keyEnd;
                    keyEnd = keyPart.indexOf('</w:t>', keyStart);
                    key += keyPart.substring(keyStart, keyEnd === -1 ? keyPart.length : keyEnd);

                    if (keyEnd === -1) key = key.replace(closeBracket, '');
                    if (keyIndex === -1 && key.length > 0) {
                        keyIndex = strParts.length;
                        addKey(key);
                    }
                }

                if (strParts[keyIndex] !== key) {
                    keys[strParts[keyIndex]].pop();
                    addKey(key, keyIndex);
                }
            }
        }

        start = to + closeBracket.length;
        from = str.indexOf(openBracket, start);
        to = str.indexOf(closeBracket, from + 1);
    }

    strParts.push(str.substring(start));

    // render
    return function(data) {
        for (var key in data) {
            if (keys[key]) {
                for (var i = 0; i < keys[key].length; i++) strParts[keys[key][i]] = data[key];
            }
        }

        return strParts.join('');
    };

    function addKey(key, index) {
        if (!keys[key]) keys[key] = [];
        if (index) {
            keys[key].push(index);
            strParts[index] = key;
        } else {
            keys[key].push(strParts.length);
            strParts.push(key);
        }
    }
}


function docxTemplate(file, opt) {
    var zip = new JSZip(file, { base64: false, checkCRC32: true}),
        files = zip.file(reFiles),
        subTemplates = {};

    for (var i = 0; i < files.length; i++) {
        subTemplates[files[i].name] = compile(files[i].asText(), opt);
    }

    return function(data) {
        var docx = new JSZip(file);

        for (var key in subTemplates) {
            var temp = subTemplates[key](data);
            docx.file(key, temp);
        }

        return docx.generate({base64: false, compression:'DEFLATE', type: 'nodebuffer'});
    }
}

module.exports = {
    openBracket: '@',
    closeBracket: '@',

    compileFile: function(file, opt, callback) {
        if (!callback && typeof opt === 'function') callback = opt;

        fs.stat(file, function(err, stat) {
            if (err) return callback(err);
            if (stat && stat.isFile()) {
                fs.readFile(file, { encoding: 'binary' }, function(err, data) {
                    if (err) return callback(err);

                    var temp = docxTemplate(data, opt);
                    callback(null, temp);
                });
            } else {
                callback('docx file not find');
            }
        });
    },

    compileFileSync: function(file, opt) {
        var stat = fs.statSync(file);
        if (stat && stat.isFile()) {
            var temp = fs.readFileSync(file, { encoding: 'binary' });
            return docxTemplate(temp, opt);
        } else {
            return null;
        }
    },

    renderToFile: function(file, template, data, cb) {
        fs.writeFile(file, template(data) , { encoding: 'binary' }, cb);
    },

    render: function(inputFile, data, outputFile, cb) {
        this.compileFile(inputFile, function(err, template) {
            module.exports.renderToFile(outputFile, template, data, cb);
        });
    },

    renderSync: function(inputFile, data, outputFile) {
        var template = this.compileFileSync(inputFile);
        fs.writeFileSync(outputFile, template(data), { encoding: 'binary' });
    },

    compile: function(data, opt) {
        return docxTemplate(data, opt);
    }
};