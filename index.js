let moment = require ("moment");
let Excel = require ("exceljs");
let nodemailer = require ("nodemailer");
let EmailData = require ("./config");
var workbook = new Excel.Workbook();
console.log(EmailData);
workbook.xlsx.readFile("./check.xlsx").then(function(data) {
  const colData = data.getWorksheet(1);
  colData.getColumn(2).values.map((data, index) => {
    if (index > 1) {
      if (moment(data).format("DD/MM") === moment().format("DD/MM")) {
        console.log(
          "SENDING MAIL TO ",
          colData.getColumn(3).values[index].text
        );
        let transporter = nodemailer.createTransport({
          service: "gmail",
          auth: {
            user: EmailData.username,
            pass: EmailData.password
          }
        });

        let mailOptions = {
          from: EmailData.username,
          to: colData.getColumn(3).values[index].text,
          subject: `${colData.getColumn(4).values[index]} ${colData.getColumn(1).values[index]}`,
          text: `${colData.getColumn(5).values[index]} Wish you many more happy return of the day`
        };
        transporter.sendMail(mailOptions, function(err, res) {
          if (err) {
            console.log("Error");
          } else {
            console.log("Email Sent");
          }
        });
      }
    }
  });
});
