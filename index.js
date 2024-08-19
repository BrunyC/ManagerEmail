var Imap = require('imap')
var MailParser = require("mailparser").MailParser;
const fs = require("fs");
const { Base64Decode } = require('base64-stream')
var nodemailer = require('nodemailer');
const {simpleParser} = require('mailparser');

const email = 'brunycardoso1992@outlook.com'
const emailPass = 'aekpgfrfmmggdljk'

var imap = new Imap({
  user: email,
  password: emailPass,
  host: 'imap-mail.outlook.com',
  port: 993,
  tls: true,
  tlsOptions: { rejectUnauthorized: false }
});

const transporter = nodemailer.createTransport({
  service: 'outlook',
  auth: {
    user: email,
    pass: emailPass
  }
});

function buildAttachments(attachment) {
  const filename = attachment.params.name;
  const encoding = attachment.encoding;

  return function (msg, seqno) {
    var prefix = '(#' + seqno + ') ';
    msg.on('body', function(stream, info) {

      console.log(prefix + 'Streaming this attachment to file', filename, info);
      var writeStream = fs.createWriteStream(filename);
      writeStream.on('finish', function() {
        console.log(prefix + 'Done writing to file %s', filename);
      });

      if (encoding.toLowerCase() === 'base64') {
        stream.pipe(new Base64Decode()).pipe(writeStream)
      } else  {
        stream.pipe(writeStream);
      }
    });
  };
}

function findAttachmentParts(struct, attachments) {
  attachments = attachments ||  [];
  for (var i = 0, len = struct.length, r; i < len; ++i) {

    if (Array.isArray(struct[i])) {
      findAttachmentParts(struct[i], attachments);
    } else {
      if (struct[i].disposition !== null && struct[i].disposition.type === 'attachment') {
        attachments.push(struct[i]);
      }
    }
  }
  return attachments;
}

function emailSender(senderEmail, clientEmail, clientSuject, text) {
  transporter.sendMail({
    from: senderEmail,
    to: clientEmail,
    subject: clientSuject,
    text: text,
    }, function(error, info){
      if (error) {
        console.log(error);
      } else {
        console.log('Sending email to ' + clientEmail);
    }
  })
}

function processMessage(msg, seqno) {
    console.log("Processing msg #" + seqno);
    var parser = new MailParser();
    var emailFrom = '';

    msg.once('attributes', function(attrs) {
      var attachments = findAttachmentParts(attrs.struct);
      if(attachments.length > 0) {
        console.log('Downloading attachments...');
        for (var i = 0, len=attachments.length ; i < len; ++i) {
          var attachment = attachments[i];
          var f = imap.fetch(attrs.uid , { 
            bodies: [attachment.partID],
            struct: true
          });
          f.on('message', buildAttachments(attachment));
        }
      }
    });

    msg.on("body", function(stream) {
      var buffer = '';
      stream.on("data", function(chunk) {
        parser.write(chunk.toString("utf8"));
        buffer += chunk.toString('utf8');
      });
      stream.once('end', async function() {
        const parsedHeader = Imap.parseHeader(buffer);
        console.log((await simpleParser(buffer)).text)
        emailFrom = parsedHeader.from[0];
        const msg = 'Order received!'
        const clientSuject = 'Teste'
        emailSender(email, emailFrom, clientSuject, msg);
      });
    });

    msg.once("end", function() {
        parser.end();
    });
}

function openInbox(cb) {
  imap.openBox('INBOX', false, cb);
}
 
imap.once('ready', function() {
  openInbox(function(err, box) {
    imap.search(['UNSEEN'], (err, results) => {
        
      if(!results || !results.length) {
        console.log("No unread mails");
        imap.end();
        return
      }

      const f = imap.fetch(results, {bodies: '', struct: true, markSeen: true})
      f.on("message", processMessage);
      f.once("error", function(err) {
          return Promise.reject(err);
      });

      f.once("end", function() {
          console.log("Done fetching all unseen messages.");
          imap.end();
      });

    })
  });
});
 
imap.once('error', function(err) {
  console.log(err);
});
 
imap.once('end', function() {
  console.log('Connection ended');
});
 
imap.connect();