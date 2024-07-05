const express = require("express");
const app = express();
const path = require("path");
const { authenticate } = require("@google-cloud/local-auth");
const fs = require("fs").promises;
const { google } = require("googleapis");
const { Client } = require('@microsoft/microsoft-graph-client');
const { ConfidentialClientApplication } = require('@azure/msal-node');

const port = 8000;
// these are the scope that we want to access 
const SCOPES = [
  "https://www.googleapis.com/auth/gmail.readonly",
  "https://www.googleapis.com/auth/gmail.send",
  "https://www.googleapis.com/auth/gmail.labels",
  "https://mail.google.com/",
];

// i kept the label name
const labelName = "Vacation Auto-Reply";


app.get("/", async (req, res) => {

  // here i am taking google GMAIL  authentication 
  const auth = await authenticate({
    keyfilePath: path.join(__dirname, "credentials.json"),
    scopes: SCOPES,
  });

  console.log("this is auth", auth)

  // here i getting authorize gmail id
  const gmail = google.gmail({ version: "v1", auth });


  //  here i am finding all the labels availeble on current gmail
  const response = await gmail.users.labels.list({
    userId: "me",
  });


  //  this function is finding all email that have unreplied or unseen
  async function getUnrepliesMessages(auth) {
    const gmail = google.gmail({ version: "v1", auth });
    const response = await gmail.users.messages.list({
      userId: "me",
      labelIds: ["INBOX"],
      q: "is:unread",
    });

    return response.data.messages || [];
  }

  //  this function generating the label ID
  async function createLabel(auth) {
    const gmail = google.gmail({ version: "v1", auth });
    try {
      const response = await gmail.users.labels.create({
        userId: "me",
        requestBody: {
          name: labelName,
          labelListVisibility: "labelShow",
          messageListVisibility: "show",
        },
      });
      return response.data.id;
    } catch (error) {
      if (error.code === 409) {
        const response = await gmail.users.labels.list({
          userId: "me",
        });
        const label = response.data.labels.find(
          (label) => label.name === labelName
        );
        return label.id;
      } else {
        throw error;
      }
    }
  }

  async function main() {
    // Create a label for theApp
    const labelId = await createLabel(auth);
    // console.log(`Label  ${labelId}`);
    // Repeat  in Random intervals
    setInterval(async () => {
      //Get messages that have no prior reply
      const messages = await getUnrepliesMessages(auth);
      // console.log("Unreply messages", messages);

      //  Here i am checking is there any gmail that did not get reply
      if (messages && messages.length > 0) {
        for (const message of messages) {
          const messageData = await gmail.users.messages.get({
            auth,
            userId: "me",
            id: message.id,
          });

          const email = messageData.data;
          const hasReplied = email.payload.headers.some(
            (header) => header.name === "In-Reply-To"
          );

          if (!hasReplied) {
            // Craft the reply message
            const replyMessage = {
              userId: "me",
              resource: {
                raw: Buffer.from(
                  `To: ${email.payload.headers.find(
                    (header) => header.name === "From"
                  ).value
                  }\r\n` +
                  `Subject: Re: ${email.payload.headers.find(
                    (header) => header.name === "Subject"
                  ).value
                  }\r\n` +
                  `Content-Type: text/plain; charset="UTF-8"\r\n` +
                  `Content-Transfer-Encoding: 7bit\r\n\r\n` +
                  `Thank you for your email. I'm currently on vacation and will reply to you when I return.\r\n`
                ).toString("base64"),
              },
            };

            await gmail.users.messages.send(replyMessage);

            // Add label and move the email
            await gmail.users.messages.modify({
              auth,
              userId: "me",
              id: message.id,
              resource: {
                addLabelIds: [labelId],
                removeLabelIds: ["INBOX"],
              },
            });
          }
        }
      }
    }, Math.floor(Math.random() * (120 - 45 + 1) + 45) * 1000);
  }



  main();
  // const labels = response.data.labels;
  res.json({ "this is Auth": auth });
});


// please ignore it
// app.get('/outlook', async (req, res) => {
//   try {
//     const config = {
//       "auth": {
//         "clientId": "c8ddccc2-9c16-4cdf-a98b-30b5d816d74b",
//         "authority": "https://login.microsoftonline.com/38f83620-4d13-4dce-8449-483b60f48c63",
//         "clientSecret": "f8cdef31-a31e-4b4a-93e4-5f571e91255a"
//       }
//     };

//     const cca = new ConfidentialClientApplication(config);
//     const authResult = await cca.acquireTokenByClientCredential({
//       scopes: ['https://graph.microsoft.com/.default']
//     });

//     const token = authResult.accessToken;

//     const client = Client.init({
//       authProvider: (done) => {
//         done(null, token);
//       }
//     });

//     const messages = await client.api('/me/messages').get();
//     console.log(messages);

//   } catch (error) {
//     console.log(error);
//     return res.status(500).json(error);
//   }
// })

app.listen(port, () => {
  console.log(`server is running ${port}`);
});
