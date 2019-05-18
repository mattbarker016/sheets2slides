# sheets2slides
A tool to take Google Sheets data and create Google Slides based on an existing template slide.

# Prerequisites

## Download & Setup

1. Make sure you have **npm** and **Node.js** installed! npm is installed with Node.js, so you can nab both by downloading the LTS (Long Term Support) Node.js version [**here**](https://nodejs.org/en/).
2. Under Step 1 on Google's Node.js [Quick Start Guide](https://developers.google.com/slides/quickstart/nodejs), click **Enable the Google Slides API** for your Google account. Make sure this account has access to the files you want to use! When you are finished, make sure you **Download Client Configuration**. This will download a file called `credentials.json` that you should include in the root directory.

**Important:** You need to create your own `credentials.json` file to use this script! The provided file is included for template purposes only, and should be replaced with your own file.

## Required Google Documents

### Presentation

This script requires an existing Google Slides presentation with a pre-made template slide. The first slide in the presentation will be used as a **template**, and generated slides will be added to the end of the presentation. Your template slide should have tokens on them, or unique strings that will be replaced by actual data. See **Files** under **Customization** for more information.

**Example Template** • [Example Google Slides Document](https://docs.google.com/presentation/d/1WecOJ0-4SowO9R2hyPH8DHO8kf6vzm8X7hH_pN4IEPM/edit?usp=sharing)

![Template Slide Example](template-example.png)

In this example, "To:" and "From:" would not be changed and remain in the generated slides.

### Spreadsheet

This script requires an existing Google Sheets spreadsheet with existing data.

# Installation

1. [Download](https://github.com/mattbarker016/sheets2slides/archive/master.zip) or clone this repository. Navigate to the root directory. If you need help, see **Tutorial #1** below.
2. Run `npm install` to install the dependencies used.
3. If you have everything configured properly, run `node sheets2slides.js`. See **Customizations** for more information.

## Token Generation & Permissions

If this is your first time running the script, or your token has expired, you will need to generate an access token to use Google's APIs. On executing the script, your browser will open to a Google sign-in page, where you will grant access for the script to use APIs modifying data in Google Sheets and Google Slides. Then, copy the generated code in your browser in the Terminal / Command Prompt window and hit enter. The script will continue as normal.

# Customization

Open `sheets2slides.js` in your favorite text editor to add your documents and customize options at the top of the file.

## Files (Required)

- **`spreadsheetID`**: The identifier of the spreadsheet used for data. This is the long string of text after `spreadsheets/d/` in the spreadsheet URL.
- **`sheetName`**: The name of the specific sheet with the data, which appears on the bottom of the spreadsheet. This is different from the spreadsheet file name.
- **`presentationID`**: The identifier of the template presentation. This is the long string of text after `presentation/d/` in the spreadsheet URL.

## Data Mapping (Required)

**`sheets2SlidesDictionary`** is the bridge between a spreadsheet and a presentation. The ***key*** of the dictionary (left side of the colon) should be a string corresponding to the column of the data in the spreadsheet. The ***value*** of the dictionary (right side of the colon) is the token, or unique string, that will be replaced by the spreadsheet data. 

For example, let's say you have a spreadsheet with names and want to add someone's first name to a slide. On Google Slides, you would add **"{{FIRST_NAME}}"** in a text box on your template slide wherever you want someone's first name to appear. Then, you would make the dictionary *key* be **"A"**, where **"A"** is the letter of the spreadsheet column corresponding to first names, and make the dictionary *value* be the token **"{{FIRST_NAME}}"**.

**Example**
```
var sheets2SlidesDictionary = {

    "A": "{{FIRST_NAME}}",

    "B": "{{LAST_NAME}}",

    "C": "{{HOMETOWN}}"
    
}
```

Note: Make sure you add a comma after every line, *except* for the last one!

## Other

You can modify the range of data you query from the spreadsheet. By default, the data spans from columns A-Z and rows 1-1000. If your data set is smaller, you can still leave these alone because the script ignores empty rows.

# Tutorials

This section is to provide a step-by-step guide for certain instructions that might not be immediately clear for first-time users.

## 1. Navigate the File Hierarchy

1. Open a new shell window. This is called Terminal on macOS and Command Prompt on Windows.
2. Type the command `cd`, which stands for **change directory**.
3. Press the space bar once to put a space between the command and the argument, or the location we want to navigate to.
4. In your file browser, drag the folder of the desired folder (in this case, the `sheets2slides` folder just added) into the shell window to copy the file path.
5. Hit Enter. As a sanity check, you can type `ls`, followed by enter, and you should see the contents of the folder you navigated to. 

# Troubleshooting

- Try reading the console output for an indication of the error. Sometimes you can figure out if you made a typo!
- If you run into network request issues that don't make sense, try deleting `token.json` and re-creating a new token.
- Don't hesitate to message me on Slack or email mjb485@cornell.edu!

