//
// sheets2slides
// created by Matt Barker (Class IX)
//


// Data Variables

// The identifier of the data spreadsheet. This is the long string of text
// between (or after) "spreadsheets/d/" and "/edit" in the spreadsheet URL.
var spreadsheetID = "1e-jWpAv2PxwuPq-FVrsvRUKsTqqDZH912HCENwpBGDc"

// The name of the specific sheet with the data.
var sheetName = "Form Responses"

// The identifier of the template presentation. This is the long string of text
// between (or after) "presentation/d/" and "/edit" in the spreadsheet URL.
var presentationID = "1WecOJ0-4SowO9R2hyPH8DHO8kf6vzm8X7hH_pN4IEPM"

// A mapping from a spreadsheet column to a token on a template slide.
// Data in Column <key> will replace instances of <value> on a slide.
var sheets2SlidesDictionary = {

    "B": "{{SENDER_NAME}}",

    "D": "{{RECIPIENT_NAME}}",

    "F": "{{MESSAGE}}"
}

// The first column with valid data.
var firstColumnLetter = "A"

// The first row number with valid data.
var firstRowNumber = 2

// The last column with valid data. This can be arbitrarily high, as long
// as it's greater than or equal to the last column letter with valid data.
// Note: Does not support columns 
var lastColumnLetter = "Z"

// The last row number with valid data. This can be arbitrarily high, as long
// as it's greater than or equal to the last row number with valid data.
var lastRowNumber = 1000


//
// DO NOT MODIFY BELOW
//

const fs = require('fs');
const readline = require('readline');
const { google } = require('googleapis');
const openURL = require('opn');

// If modifying these scopes, delete token.json.
const SCOPES = [
    'https://www.googleapis.com/auth/presentations',
    'https://www.googleapis.com/auth/spreadsheets'
];
// The file token.json stores the user's access and refresh tokens, and is
// created automatically when the authorization flow completes for the first
// time.
const TOKEN_PATH = 'token.json';

// Load client secrets from a local file.
fs.readFile('credentials.json', (err, content) => {
    if (err) return console.log('Error loading client secret file:', err);
    // Authorize a client with credentials, then call the Google Slides API.
    authorize(JSON.parse(content), main);
});

/**
 * Create an OAuth2 client with the given credentials, and then execute the
 * given callback function.
 * @param {Object} credentials The authorization client credentials.
 * @param {function} callback The callback to call with the authorized client.
 */
function authorize(credentials, callback) {
    const { client_secret, client_id, redirect_uris } = credentials.installed;
    const oAuth2Client = new google.auth.OAuth2(
        client_id, client_secret, redirect_uris[0]
    );

    // Check if we have previously stored a token.
    fs.readFile(TOKEN_PATH, (err, token) => {
        if (err) return getNewToken(oAuth2Client, callback);
        oAuth2Client.setCredentials(JSON.parse(token));
        callback(oAuth2Client);
    });
}

/**
 * Get and store new token after prompting for user authorization, and then
 * execute the given callback with the authorized OAuth2 client.
 * @param {google.auth.OAuth2} oAuth2Client The OAuth2 client to get token for.
 * @param {getEventsCallback} callback The callback for the authorized client.
 */
function getNewToken(oAuth2Client, callback) {
    const authUrl = oAuth2Client.generateAuthUrl({
        access_type: 'online',
        scope: SCOPES,
    });
    console.log('Authorize this app by visiting this url:', authUrl);
    openURL(authUrl);
    const rl = readline.createInterface({
        input: process.stdin,
        output: process.stdout,
    });
    rl.question('Enter the code from that page here: ', (code) => {
        rl.close();
        oAuth2Client.getToken(code, (err, token) => {
            if (err) return console.error('Error retrieving access token', err);
            oAuth2Client.setCredentials(token);
            // Store the token to disk for later program executions
            fs.writeFile(TOKEN_PATH, JSON.stringify(token), (err) => {
                if (err) return console.error(err);
                console.log('Token stored to', TOKEN_PATH);
            });
            callback(oAuth2Client);
        });
    });
}

// Helper Functions

/** This function appends a duplicateObject request to an existing array of requests. */
function duplicateObject(requests, objectId, newObjectId) {
    return requests.concat([{
        "duplicateObject": {
            "objectId": objectId,
            "objectIds": {
                [objectId]: newObjectId
            }
        }
    }])
}

/** This function appends a replaceAllText request to an existing array of requests. */
function replaceText(requests, pageId, replaceText, matchText) {
    return requests.concat([{
        "replaceAllText": {
            "replaceText": replaceText,
            "pageObjectIds": [pageId],
            "containsText": {
                "text": matchText,
                "matchCase": true
            }
        }
    }])
}

/** Converts the first letter of the passed in string to a number. 
 * Example: A -> 0, B -> 1, ..., Z -> 25, AA -> 26, AB -> 27, etc.
 * */
function letterToNumber(letter) {
    var number = 0;
    for (var i = 0; i < letter.length(); i++) {
        number = number * 26 + (letter.charAt(i) - ('A' - 1));
    }
    return number;
}

/** Converts the first letter of the passed in string to a number. 
 * Example: A -> 0, B -> 1, etc. 
 * */
function letterToNumber(letter) {
    var asciiCodeCapitalA = 65
    return letter.charCodeAt(0) - asciiCodeCapitalA
}

// Main Function

async function main(auth) {

    const sheets = google.sheets({ version: 'v4', auth });
    const slides = google.slides({ version: 'v1', auth });

    console.log("Fetching template slide...");
    var slideID = await getSlideIdentifier(slides);

    console.log("Fetching spreadsheet data...");
    var values = await fetchSpreadsheetData(sheets);

    console.log(`Generating content for ${values.length} slides...`)
    var requests = generateSlidesContent(values, slideID)

    console.log("Updating template presentation with content. This may take a bit...")
    var replies = await batchUpdateSlides(slides, requests)

    console.log("Checking response...")
    var numberOfRequests = requests.length / values.length;
    var numberOfSlides = replies.length / numberOfRequests;

    console.log(`\n${numberOfSlides} slides have been created!`);

    var presentationUrl = `https://docs.google.com/presentation/d/${presentationID}`;
    console.log(`Check them out at ${presentationUrl}`);

}

/** Get the objectId of the first slide in the presentation.  */
function getSlideIdentifier(slides) {
    var slideID = "";
    return new Promise(function (resolve, reject) {
        slides.presentations.get({
            presentationId: presentationID
        }).then((slidesResponse) => {
            slideID = slidesResponse.data.slides[0].objectId;
            resolve(slideID);
        }).catch((errorResponse) => {
            console.log("ERROR: getSlideIdentifier failed!")
            console.log(errorResponse);
            resolve(slideID);
        })
    })
}

/** Fetch data from the spreadsheet */
function fetchSpreadsheetData(sheets) {
    var values = [];
    var rangeOpen = `${firstColumnLetter}${firstRowNumber}`;
    var rangeClose = `${lastColumnLetter}${lastRowNumber}`;
    var dataRangeNotation = `${sheetName}!${rangeOpen}:${rangeClose}`;
    return new Promise(function (resolve, reject) {
        sheets.spreadsheets.values.get({
            spreadsheetId: spreadsheetID,
            range: dataRangeNotation
        }).then((sheetsResponse) => {
            values = sheetsResponse.data.values;
            resolve(values);
        }).catch((errorResponse) => {
            console.log("ERROR: fetchSpreadsheetData failed!")
            console.log(errorResponse.errors);
            resolve(values)
        })
    })
}

/** Create Google Slides API requests to create a series of slides. */
function generateSlidesContent(data, slideID) {
    var requests = []
    for (var i = 0; i < data.length; ++i) {
        var row = data[i];
        var pageId = `slide_${i}`;
        // Basic check to confirm there is data in the row
        if (row[0] != "") {
            // Duplicate slide request
            requests = duplicateObject(requests, slideID, pageId);
            // Replace text requests
            for (var key in sheets2SlidesDictionary) {
                var value = row[letterToNumber(key)];
                var token = sheets2SlidesDictionary[key];
                requests = replaceText(requests, pageId, value, token);
            }
        }
    }
    return requests
}

/** Update the Google Slides presentation with requests. Returns API replies from requests. */
function batchUpdateSlides(slides, requests) {
    var replies = [];
    return new Promise(function (resolve, reject) {
        slides.presentations.batchUpdate({
            presentationId: presentationID,
            resource: { requests: requests }
        }).then((batchUpdateResponse) => {
            replies = batchUpdateResponse.data.replies;
            resolve(replies);
        }).catch((errorResponse) => {
            console.log("ERROR: batchUpdateSlides failed!")
            console.log(errorResponse.errors);
            resolve(replies)
        });
    })
}

