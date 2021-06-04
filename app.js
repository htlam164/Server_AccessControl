const express = require('express');
const bodyParser = require('body-parser');
const exphbs = require('express-handlebars');
const path = require('path');
const nodemailer = require('nodemailer');
const { google } = require("googleapis");
var { DateTime } = require('luxon');
var uuid = require('uuid');

const OAuth2 = google.auth.OAuth2;
const app = express();


const oauth2Client = new OAuth2(
  "491941003687-2bjjg48s15h8kchmpjm8m02j31k23ce6.apps.googleusercontent.com", // ClientID
  "Xyin8MTLokDzGJYVj4iKeLYZ", // Client Secret
  "https://developers.google.com/oauthplayground" // Redirect URL
);

oauth2Client.setCredentials({
  refresh_token: "1//04t0hpofoAgbPCgYIARAAGAQSNwF-L9Iru36QHV28u6XUPsBE5nlKf7TR8qF14nr2pNR6gh3jQtpeX5vOXtyjrnCFmqC5Xa2WxKs"
});
const accessToken = oauth2Client.getAccessToken()




// View engine setup
app.engine('.hbs', exphbs({
  extname: '.hbs',
  defaultLayout: 'contact',
  layoutsDir: path.join(__dirname, 'views')
}));
app.set('view engine', '.hbs');
app.set('views',path.join(__dirname,'views'))
// Static folder
app.use('/public', express.static(path.join(__dirname, 'public')));

// Body Parser Middleware
app.use(bodyParser.urlencoded({ extended: false }));
app.use(bodyParser.json());


app.get('/', (req, res) => {
  res.render('contact');
});

app.post('/send', async (req, res) => {


  let id = uuid.v4();
  //const qrcode = await QRCode.toDataURL(id);
  const qrcode = '=Image("'+ 'https://chart.googleapis.com/chart?chs=250x250&cht=qr&chl=' + req.body.phone.slice(1) +'")';
  const output = `
    <p>Successfully submitted! Thank you!</p>
    <h3>Contact Details</h3>
    <ul>  
      <li>Name: ${req.body.name}</li>
      <li>Company: ${req.body.company}</li>
      <li>Email: ${req.body.email}</li>
      <li>Phone: ${req.body.phone}</li>
      <li>QrCode: <img src="${'https://chart.googleapis.com/chart?chs=250x250&cht=qr&chl=' + req.body.phone.slice(1)}"></li>
    </ul>
    <h3>Message</h3>
    <p>${req.body.message}</p>
  `;


  const auth = new google.auth.GoogleAuth({
    keyFile: "credentials.json",
    scopes: "https://www.googleapis.com/auth/spreadsheets",
  });

  // Create client instance for auth
  const client = await auth.getClient();

  // Instance of Google Sheets API
  const googleSheets = google.sheets({ version: "v4", auth: client });

  const spreadsheetId = "1UwI_mntoJqBcsKqBulmrDsm8SjIIpTyM4DMdQE-5phc";

  // Get metadata about spreadsheet
  const metaData = await googleSheets.spreadsheets.get({
    auth,
    spreadsheetId,
  });

 
  let d = DateTime.TIME_SIMPLE();
  let date_ob = d;
  let date_ob_set = "12/30/2030"
  //console.log(d); //Asia/Saigon
    // Clear 
  
  
  // Write row(s) to spreadsheet
  await googleSheets.spreadsheets.values.append({
    auth,
    spreadsheetId,
    range: "Users!A:G",
    valueInputOption: "USER_ENTERED",
    resource: {
      values: [[date_ob,date_ob_set, req.body.name, req.body.email,req.body.phone, id, qrcode]],
    },
  });


  // create reusable transporter object using the default SMTP transport
  const smtpTransport = nodemailer.createTransport({
    service: "gmail",
    auth: {
         type: "OAuth2",
         user: "htlam164@gmail.com", 
         clientId: "491941003687-2bjjg48s15h8kchmpjm8m02j31k23ce6.apps.googleusercontent.com",
         clientSecret: "Xyin8MTLokDzGJYVj4iKeLYZ",
         refreshToken: "1//04t0hpofoAgbPCgYIARAAGAQSNwF-L9Iru36QHV28u6XUPsBE5nlKf7TR8qF14nr2pNR6gh3jQtpeX5vOXtyjrnCFmqC5Xa2WxKs",
         accessToken: accessToken
    }
  });
  console.log(req.body.email);
  // setup email data with unicode symbols
  let mailOptions = {
      from: '"Nodemailer Contact" <htlam164@gmail.com>', // sender address
      to: req.body.email, // list of receivers
      subject: 'Node Contact Request', // Subject line
      text: 'Hello world?', // plain text body
      html: output // html body
  };

  // send mail with defined transport object
  smtpTransport.sendMail(mailOptions, (error, info) => {
      if (error) {
          return console.log(error);
      }
      console.log('Message sent: %s', info.messageId);   
      console.log('Preview URL: %s', nodemailer.getTestMessageUrl(info));
      
        res.render('contact', {msg:'Email has been sent'});

      
  });
  });
const PORT = 3000;
app.listen(process.env.PORT || PORT, () => console.log('Server started...'));