var PizZip = require('pizzip');
var Docxtemplater = require('docxtemplater');

var fs = require('fs');
var path = require('path');

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

var content = fs.readFileSync(
    path.resolve(__dirname, 'input.docx'), 'binary');
var input_json = require('./input.json');
var zip = new PizZip(content);
var doc;

try {
    doc = new Docxtemplater(zip);
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
