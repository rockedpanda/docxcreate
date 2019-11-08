var express = require('express');
var router = express.Router();
var fs = require('fs');
var path = require('path');
var tmpDir = global.configInfo.tmpDir || '.';
var uuid = require('uuid/v4');
var docxTemplate = require('../lib/docxTemplate');

function getTmpPath(name){
    return path.resolve(__dirname,'../public/tmpl/', name);
}

/* put file and save as template */
router.put('/tmpl', function(req, res, next) {
    let name = uuid().replace(/-/g,'');
    let p = getTmpPath(name+'.docx');
    docxTemplate.saveTemplate(name, p);

    let outs = fs.createWriteStream(p);
    req.pipe(outs).on('close',function(){
            res.send({error_code:0,msg:'模板保存成功', data:name});
            // res.send({error_code:1,msg:'保存失败',data:null});
    });
});

router.post('/create/:name', function(req, res, next){
    let name = req.params.name;
    let body = req.body; 
    let p = docxTemplate.createDocById(name, body);
    if(p){
        res.send({error_code:0, msg:'生成成功', data:'/'+p});
    }else{
        res.send({error_code:1, msg:'生成失败', data:null});
    }
});

module.exports = router;
