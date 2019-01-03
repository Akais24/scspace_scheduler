const express = require('express');
const bodyParser = require('body-parser');

const app = express();
const PORT = 3000;

var router = require('./routes/index')(app);

app.set('views', __dirname + '/views');
app.set('view engine', 'ejs');
app.engine('html', require('ejs').renderFile);

const server = app.listen(PORT, () => {
  console.log("server is running at port 3000")
})
