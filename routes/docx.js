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

/* 上传一个docx作为模板,并返回模板id */
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

/** 根据一个模板id删除服务端的一个模板 */
router.delete('/tmpl/:name', function(req, res, next){
    let name = req.params.name;
    if(docxTemplate.getTemplate(name)){
        docxTemplate.delTemplate(name).then(function(){
            res.send({error_code:0,msg:'模板移除成功',data: name});
        }).catch(function(err){
            console.log(err);
            res.send({error_code:1,msg:'模板移除失败',data: name});
        });
    }else{
        res.send({error_code:1,msg:'不存在该模板',data: name});
    }
})

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
