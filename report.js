var PizZip = require('pizzip');
var Docxtemplater = require('docxtemplater');

var fs = require('fs');
var path = require('path');

var expressions = require('angular-expressions');
var merge = require("lodash.merge");

function replaceErrors(key, value) {
    if (value instanceof Error) {
        return Object.getOwnPropertyNames(value).reduce(function(error, key) {
            error[key] = value[key];
            return error;
        }, {});
    }
    return value;
}

function errorHandler(error) {
    console.log(JSON.stringify({error: error}, replaceErrors));

    if (error.properties && error.properties.errors instanceof Array) {
        const errorMessages = error.properties.errors.map(function (error) {
            return error.properties.explanation;
        }).join("\n");
        console.log('errorMessages', errorMessages);
    }
    throw error;
}

function angularParser(tag) {
    if (tag === '.') {
        return {
            get: function(s){ return s;}
        };
    }
    const expr = expressions.compile(
        tag.replace(/(’|‘)/g, "'").replace(/(“|”)/g, '"')
    );
    return {
        get: function(scope, context) {
            let obj = {};
            const scopeList = context.scopeList;
            const num = context.num;
            for (let i = 0, len = num + 1; i < len; i++) {
                obj = merge(obj, scopeList[i]);
            }
            return expr(scope, obj);
        }
    };
}

var content = fs.readFileSync(
    path.resolve(__dirname, 'input.docx'), 'binary');
var input_json = require('./input.json');
var zip = new PizZip(content);
var doc;

try {
    doc = new Docxtemplater().loadZip(zip).setOptions({parser:angularParser});
    //doc = new Docxtemplater(zip);
} catch(error) {
    errorHandler(error);
}

doc.setData(input_json);

try {
    doc.render()
}
catch (error) {
    errorHandler(error);
}

var buf = doc.getZip().generate(
    {type: 'nodebuffer'});

fs.writeFileSync(path.resolve(__dirname, 'output.docx'), buf);
