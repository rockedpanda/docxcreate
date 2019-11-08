var PizZip = require('pizzip');
var Docxtemplater = require('docxtemplater');

var fs = require('fs');
var path = require('path');

var templateMap = {
  cache:{},
  load: function(){
    this.cache = JSON.parse(fs.readFileSync(path.resolve(__dirname,'../public/tmpl/all.json'),'utf8'));
  },
  save: function(){
    fs.writeFileSync(path.resolve(__dirname,'../public/tmpl/all.json'),JSON.stringify(this.cache),'utf8');
  },
  set: function(name, p){
    this.cache[name] = p;
    this.save();
  },
  get: function(name){
    return this.cache[name] || null;
  }
};

templateMap.load();

//Load the docx file as a binary
function createDoc(templatePath, data) {
  var content = fs
    .readFileSync(templatePath, 'binary');

  var zip = new PizZip(content);

  var doc = new Docxtemplater();
  doc.loadZip(zip);

  //set the templateVariables
  doc.setData(data); /* data = {
  "clients": [
    {
      "first_name": "John",
      "last_name": "Doe",
      "phone": "+44546546454"
    },
    {
      "first_name": "Jane",
      "last_name": "Doe",
      "phone": "+445476454"
    },
	{
      "first_name": "王",
      "last_name": "海滨",
      "phone": "18612345678"
    }
  ],
  "name":"老王"
}*/

  try {
    // render the document (replace all occurences of {first_name} by John, {last_name} by Doe, ...)
    doc.render()
  }
  catch (error) {
    var e = {
      message: error.message,
      name: error.name,
      stack: error.stack,
      properties: error.properties,
    }
    console.log(JSON.stringify({ error: e }));
    // The error thrown here contains additional information when logged with JSON.stringify (it contains a property object).
    throw error;
  }

  var buf = doc.getZip()
    .generate({ type: 'nodebuffer' });

  // buf is a nodejs buffer, you can either write it to a file or do anything else with it.
  let relativePath = 'docx/output_' + Date.now() + '_' + (Math.random() * 0xFFFFFF>>0).toString(16) + '.docx';
  let docxPath = path.resolve(__dirname, '../public/', relativePath);
  fs.writeFileSync(docxPath, buf);
  return relativePath;
};

function createDocById(templateId, data) {
  let templatePath = templateMap.get(templateId);
  if (templatePath) {
    return createDoc(templatePath, data);
  } else {
    return null;
  }
}


function saveTemplate(name, p){
  templateMap.set(name, p);
}

exports.createDoc = createDoc;
exports.createDocById = createDocById;
exports.saveTemplate = saveTemplate;
