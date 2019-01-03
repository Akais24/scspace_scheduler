module.exports = function(app)
{
    var path = require('path');
    var formidable = require('formidable');

    const exec = require("child_process").exec;
    function execute(command, callback){
        exec(command, function(error, stdout, stderr){ callback(stdout); });
    };

     app.get('/',function(req,res){
        res.render('index.html')
     });

     app.post('/process',function(req,res){
        var form = new formidable.IncomingForm();
        form.parse(req);

        form.on('fileBegin', function (name, file){
            file.path = path.join(__dirname, '../input', file.name)
        });

        form.on('file', function (name, file){
            console.log('Uploaded ' + file.name);
            console.log("Execute : python vacation_scheduler_web.py " + file.name)
            execute("python vacation_scheduler_web.py " + file.name, function(result){
                console.log(result)
                var result_filename = file.name.split('.')[0] + "_result.xlsx"
                res.download(path.join(__dirname, '../result', result_filename), result_filename);
            })
        });
    });
}