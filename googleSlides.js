const fs = require('fs').promises;
const path = require('path');
const process = require('process');
const { authenticate } = require('@google-cloud/local-auth');
const { google } = require('googleapis');
const request = require('request');

// If modifying these scopes, delete token.json.
const SCOPES = [
    'https://www.googleapis.com/auth/gmail.compose',
    'https://www.googleapis.com/auth/drive',
    'https://www.googleapis.com/auth/drive.file',
    'https://www.googleapis.com/auth/drive.readonly',
    'https://www.googleapis.com/auth/presentations',
    'https://www.googleapis.com/auth/presentations.readonly',
];
// The file token.json stores the user's access and refresh tokens, and is
// created automatically when the authorization flow completes for the first
// time.
const TOKEN_PATH = path.join(process.cwd(), 'token.json');
const CREDENTIALS_PATH = path.join(process.cwd(), 'credentials.json');

/**
 * Reads previously authorized credentials from the save file.
 *
 * @return {Promise<OAuth2Client|null>}
 */
async function loadSavedCredentialsIfExist() {
    try {
        const content = await fs.readFile(TOKEN_PATH);
        const credentials = JSON.parse(content);
        return google.auth.fromJSON(credentials);
    } catch (err) {
        return null;
    }
}

/**
 * Serializes credentials to a file comptible with GoogleAUth.fromJSON.
 *
 * @param {OAuth2Client} client
 * @return {Promise<void>}
 */
async function saveCredentials(client) {
    const content = await fs.readFile(CREDENTIALS_PATH);
    const keys = JSON.parse(content);
    const key = keys.installed || keys.web;
    const payload = JSON.stringify({
        type: 'authorized_user',
        client_id: key.client_id,
        client_secret: key.client_secret,
        refresh_token: client.credentials.refresh_token,
    });
    await fs.writeFile(TOKEN_PATH, payload);
}

/**
 * Load or request or authorization to call APIs.
 *
 */
async function authorize() {
    let client = await loadSavedCredentialsIfExist();
    if (client) {
        return client;
    }
    client = await authenticate({
        scopes: SCOPES,
        keyfilePath: CREDENTIALS_PATH,
    });
    if (client.credentials) {
        await saveCredentials(client);
    }
    return client;
}

/**
 * Prints the number of slides and elements in a sample presentation:
 * https://docs.google.com/presentation/d/1EAYk18WDjIG-zp_0vLm3CsfQh_i8eXc67Jo2O9C6Vuc/edit
 * @param {google.auth.OAuth2} auth The authenticated Google OAuth client.
 */
async function listSlides(auth) {
    const slidesApi = google.slides({ version: 'v1', auth });
    const res = await slidesApi.presentations.get({
        presentationId: '1M6_fdHKDRgrMlcRFmuzZQP9-WUBQryBzjG1QUF9BQEI',
    });
    const slides = res.data.slides;
    if (!slides || slides.length === 0) {
        console.log('No slides found.');
        return;
    }
    console.log('The presentation contains %s slides:', slides.length);
    //console.log(JSON.stringify(res.data.slides[0]))
    return res.data.slides[0].objectId;
}

async function updateSlideText(presentationId, slideId, text, auth, temp) {
    const slides = google.slides({ version: 'v1', auth });
    const requests = [
        {
            replaceAllText: {
                replaceText: text,
                containsText: {
                    text: temp || "<<NAME>>",
                    matchCase: false
                }
            }
        }
    ];
    await slides.presentations.batchUpdate({
        presentationId: presentationId,
        requestBody: {
            requests: requests,
            writeControl: {
                requiredRevisionId: (await slides.presentations.get({ presentationId })).data.revisionId
            }
        }
    });
    console.log(`Updated slide ${slideId} with text: ${text}`);
}



async function generatePdfAndSendEmail(presentationId, slideId, recipientEmail, name, auth) {
    const slides = google.slides({ version: 'v1', auth });

    // Generate the PDF file from the slide

    const pdfFile = await slides.presentations.pages.getThumbnail({
        presentationId: presentationId,
        pageObjectId: slideId,
    }, { responseType: 'json' });
    //console.log(pdfFile)
    updateSlideText(presentationId, slideId, "<<NAME>>", auth, name)

    sendEmailWithImage(recipientEmail, 'Test Email', 'Hello, World!', pdfFile.data.contentUrl, auth);
}



async function sendEmailWithImage(to, subject, body, imageUrl, auth) {
    const gmail = google.gmail({ version: 'v1', auth });
    const attachment = await new Promise((resolve, reject) => {
        request.get({ url: imageUrl, encoding: null }, (err, response, body) => {
            if (err) {
                reject(err);
            } else {
                const encodedImage = Buffer.from(body).toString('base64');
                const attachment = {
                    filename: path.basename(imageUrl),
                    content: encodedImage,
                    encoding: 'base64'
                };
                resolve(attachment);
            }
        });
    });

    const message = [
        `To: ${to}`,
        'Content-Type: multipart/mixed; boundary="boundary"',
        'Reply-To: <noreply@gmail.com>',
        '',
        '--boundary',
        'Content-Type: text/plain; charset="UTF-8"',
        'Content-Transfer-Encoding: 8bit',
        '',
        body,
        '--boundary',
        `Content-Type: image/png; name="${attachment.filename}"`,
        'Content-Transfer-Encoding: base64',
        'Content-Disposition: attachment',
        '',
        attachment.content,
        '--boundary--'
    ].join('\n');

    await gmail.users.messages.send({
        userId: 'me',
        requestBody: {
            raw: Buffer.from(message).toString('base64')
        }
    });
    console.log(`Sent email with attached image to ${to}`);
}



const startJob = async (name, email) => {
    await authorize().then(async (auth) => {
        const slideId = await listSlides(auth);
        await updateSlideText("1M6_fdHKDRgrMlcRFmuzZQP9-WUBQryBzjG1QUF9BQEI", slideId, name, auth)
        await generatePdfAndSendEmail("1M6_fdHKDRgrMlcRFmuzZQP9-WUBQryBzjG1QUF9BQEI", slideId, email, name, auth)
    });
    return
}

module.exports = { startJob };