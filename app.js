// Your code goes here

const excelToJson = require("convert-excel-to-json");

const result = excelToJson({
  sourceFile: "names.xlsx",
});

const mail = require("@sendgrid/mail");
const APIkey = "XD";

mail.setApiKey(APIkey);

let date = new Date();
let year = date.getFullYear();
let month = date.getMonth() + 1;
let day = date.getDate();

for (let row = 2; row < result.Sheet1.length; row++) {
  let name = result.Sheet1[row].A;
  let email = result.Sheet1[row].B;
  let course = result.Sheet1[row].C;
  let grade = result.Sheet1[row].D;

  const msg = {
    to: email,
    from: "3bdullahshabib@gmail.com",
    subject: "congrat",
    html: `<div style="background-image: url(border.png); background-size: 100% 100%">
    <div
      class="container"
      style="
        width: 100%;
        height: 100;
        padding: 600px;
        display: inline;
        justify-content: center;
        text-align: center;
      "
    >
      <h1
        style="
          padding-top: 150px;
          font-size: 60nopx;
          font-family: 'Times New Roman', Times, serif;
          color: midnightblue;
        "
      >
        <strong> Certificate of Completion</strong>
      </h1>
      <p
        style="
          font-size: 60px;
          font-family: 'Times New Roman', Times, serif;
          color: midnightblue;
        "
      >
        <strong> certificate is awarded to </strong>
      </p>
      <h2
        style="
          margin: 40px;
          font-size: 40px;
          font-family: 'Times New Roman', Times, serif;
          color: black;
        "
      >
        ${name}
      </h2>
      <h4
        style="
          font-size: 60px;
          font-family: 'Times New Roman', Times, serif;
          color: midnightblue;
        "
      >
        for the secssful completion of
      </h4>
      <h3
        style="
          margin: 20px;
          font-size: 40px;
          font-family: 'Times New Roman', Times, serif;
          color: black;
        "
      >
        ${course}
      </h3>
      <h4
        style="
          margin-bottom: 80px;
          font-size: 30px;
          font-family: 'Times New Roman', Times, serif;
          color: black;
        "
      >
        with a grade of ${grade} on ${day}/${month}/${year}
      </h4>
      <h5
        style="
          font-size: 30px;
          font-family: 'Times New Roman', Times, serif;
          color: midnightblue;
        "
      >
        signature
      </h5>
    </div>
  </div>`,
  };

  mail
    .send(msg)
    .then((response) => console.log("email sent"))
    .catch((error) => console.log(error.message));
}
