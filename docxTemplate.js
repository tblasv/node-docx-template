var fs = require('fs');
require('node-zip');

var defOpenTag = '{';
var defCloseTag = '}';

var defReKey = buildReKey(defOpenTag, defCloseTag);
var reSpace = /\s*/;
var reFiles = /word\/(document|footer\d+|header\d+).xml/;

function buildReKey(openTag, closeTag) {
    return new RegExp('\\' + openTag + '([^\\' + closeTag + ']+)\\' + closeTag);
}

function compile(str, opt) {
    var reKey = defReKey;
    var openTag = opt && opt.openTag || defOpenTag;
    var closeTag = opt && opt.closeTag || defCloseTag;
    if(opt && ( opt.openTag || opt.closeTag )) {
        reKey = buildReKey(openTag, closeTag);
    }

    var keys = {};
    var strParts = [];
    var from = str.indexOf(openTag);
    var to = str.indexOf(closeTag);
    var start = 0;
    var key;

    while(to != -1) {
        var subStr = str.substring(start, from);
        strParts.push(subStr);

        var keyPart = str.substring(from, to + closeTag.length);
        var xmlCloseTag = str.indexOf('</', to + closeTag.length);

        if(str.substr(xmlCloseTag, 6) !== '</w:t>') {
            strParts.push(keyPart);
        } else {
            var reg = reKey.exec(keyPart);

            if(reg && reg[1].indexOf('<w:t>') === -1) {
                key = reg[1].replace(reSpace, '');
                addKey(key);
            } else {
                var hStr;
                var keyStart = 0;
                var keyEnd = keyPart.indexOf('</w:t>');
                var keyIndex = -1;
                key = keyPart.substring(keyStart, keyEnd).replace(openTag, '');


                if(key.length > 0) {
                    keyIndex = strParts.length;
                    addKey(key);
                }

                while(keyEnd !== -1) {

                    keyStart = keyEnd;
                    keyEnd = keyPart.indexOf('<w:t>', keyEnd + 6) + 5;
                    hStr =  keyPart.substring(keyStart, keyEnd);

                    strParts.push(hStr);

                    keyStart = keyEnd;
                    keyEnd = keyPart.indexOf('</w:t>', keyStart);
                    key += keyPart.substring(keyStart, keyEnd === -1 ? keyPart.length : keyEnd);
                    if(keyEnd === -1) key = key.replace(closeTag, '');


                    if(keyIndex === -1 && key.length > 0) {
                        keyIndex = strParts.length;
                        addKey(key);
                    }
                }

                if(strParts[keyIndex] !== key) {
                    keys[strParts[keyIndex]].pop();
                    addKey(key, keyIndex);
                }
            }
        }

        start = to + closeTag.length;
        from = str.indexOf(openTag, start);
        to = str.indexOf(closeTag, start);
    }

    strParts.push(str.substring(start));

    // render
    return function(data) {
        for(var key in data) {
            if(keys[key]) {
                for(var i = 0; i < keys[key].length; i++) strParts[keys[key][i]] = data[key];
            }
        }

        var res = strParts.join('');

        return res;
    };

    function addKey(key, index) {
        if(!keys[key]) keys[key] = [];
        if(index) {
            keys[key].push(index);
            strParts[index] = key;
        } else {
            keys[key].push(strParts.length);
            strParts.push(key);
        }
    }
}


function docxTemplate(file, opt) {
    var zip = new JSZip(file, { base64: false, checkCRC32: true});
    var files = zip.file(reFiles);
    var subTemplates = {};

    for(var i = 0; i < files.length; i++) {
        subTemplates[files[i].name] = compile(files[i].asText(), opt);
    }

    return function(data) {
        var docx = new JSZip(file);

        for(var key in subTemplates) {
            var temp = subTemplates[key](data);
            docx.file(key, temp);
        }

        return docx.generate({base64: false, compression:'DEFLATE', type: 'nodebuffer'});
    }
}

function escape(str){
    if(str) {
        return str.replace(/&/g, '&amp;')
                  .replace(/</g, '&lt;')
                  .replace(/>/g, '&gt;')
                  .replace(/"/g, '&quot;');
    }
}

module.exports = {
    settings: function(p) {
        defOpenTag = (p && escape(p.openTag)) || defOpenTag;
        defCloseTag = (p && escape(p.closeTag)) || defCloseTag;

        defReKey = buildReKey(defOpenTag, defCloseTag);
    },

    compileFile: function(file, opt, cb) {
        if(!cb && typeof opt === 'function') cb = opt;

        fs.stat(file, function(err, stat) {
            if(err) return cb(err);
            if(stat && stat.isFile()) {
                fs.readFile(file, { encoding: 'binary' }, function(err, data) {
                    if(err) return cb(err);

                    var temp = docxTemplate(data, opt);
                    cb(null, temp);
                });
            } else {
                cb('docx file not find');
            }
        });
    },

    compileFileSync: function(file, opt) {
        try {
            var stat = fs.statSync(file);
            if(stat && stat.isFile()) {
                var temp = fs.readFileSync(file, { encoding: 'binary' });
                return docxTemplate(temp, opt);
            } else {
                return null;
            }
        } catch(e) {
            return null;
        }
    },

    renderToFile: function(file, template, data, cb) {
        fs.writeFile(file, template(data) , { encoding: 'binary' }, cb);
    },

    compile: function(data, opt) {
        return docxTemplate(data, opt);
    }
};
