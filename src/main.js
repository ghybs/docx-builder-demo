var Docxtemplater = require('docxtemplater');
var JSZip = require('jszip');
var JSZipUtils = require('jszip-utils');
var saveAs = require('file-saver').saveAs;

var insertedDocsFormatted = [];
var listOfInsertedDocs = _getElementId('listOfInsertedDocs');
var fileInputLi = _getElementId('fileInputLi');

// Add the content of a local file system DocX document.
_onElementEvent('inputFile', 'change', function (event) {
  var inputEl = event.currentTarget,
      inputValue = inputEl.value,
      files = inputEl.files,
      fr = new FileReader(),
      file;

  if (files.length === 0) {
    alert('Please select a file');
    return;
  }

  file = files[0];
  inputEl.value = ''; // Reset the input.
  fr.onload = function () {
    var zip = new JSZip(fr.result),
        doc = new Docxtemplater().loadZip(zip);

    // Inspired from https://github.com/raulbojalil/docx-builder
    var xml = zip.files[doc.fileTypeConfig.textPath].asText();
    xml = xml.substring(xml.indexOf("<w:body>") + 8);
    xml = xml.substring(0, xml.indexOf("</w:body>"));
    xml = xml.substring(0, xml.indexOf("<w:sectPr"));
    insertedDocsFormatted.push(xml);

    // Append the file name in the display.
    var li = document.createElement('li');
    li.innerHTML = inputValue;

    listOfInsertedDocs.insertBefore(li, fileInputLi);
  };
  fr.readAsArrayBuffer(file);
});


// Build the final DocX document.
// Adapted from http://docxtemplater.readthedocs.io/en/latest/generate.html
_onElementClick('build', function (event) {

  _loadFile('template.docx', function(error, content){
    if (error) {
      throw error;
    }
    var zip = new JSZip(content);
    var doc = new Docxtemplater().loadZip(zip);
    doc.setData({
      first_name: 'John',
      last_name: 'Doe',
      phone: '0652455478',
      description: 'New Website',
      inserted_docs_formatted: insertedDocsFormatted.join('<w:br/><w:br/>') // Insert blank line between contents.
    });

    try {
      // render the document (replace all occurrences of {first_name} by John, {last_name} by Doe, ...)
      doc.render()
    }
    catch (error) {
      var e = {
        message: error.message,
        name: error.name,
        stack: error.stack,
        properties: error.properties
      };
      console.log(JSON.stringify({error: e}));
      // The error thrown here contains additional information when logged with JSON.stringify (it contains a property object).
      throw error;
    }

    var out = doc.getZip().generate({
      type: 'blob',
      mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    }); //Output the document using Data-URI
    saveAs(out, 'output.docx');
  });

});



function _loadFile(url,callback){
  JSZipUtils.getBinaryContent(url,callback);
}

function _onElementClick(elementId, clickListener) {
  _onElementEvent(elementId, 'click', clickListener);
}

function _onElementEvent(elementId, eventName, eventListener) {
  _getElementId(elementId).addEventListener(eventName, function (event) {
    event.preventDefault();

    eventListener(event);
  });
}

function _getElementId(elementId) {
  return document.getElementById(elementId);
}
