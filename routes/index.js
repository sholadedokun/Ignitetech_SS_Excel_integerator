/*
 * Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/**
 * This sample shows how to:
 *    - Get the current user's metadata
 *    - Get the current user's profile photo
 *    - Attach the photo as a file attachment to an email message
 *    - Upload the photo to the user's root drive
 *    - Get a sharing link for the file and add it to the message
 *    - Send the email
 */
const express = require('express');
const router = express.Router();
const graphHelper = require('../utils/graphHelper.js');
const emailer = require('../utils/emailer.js');
const passport = require('passport');
const smartsheetHelper = require('../utils/smartsheetHelper');
const config = require('../utils/config');
const SheetsToEdit = [
	// { graphName: "P&L", graphRange:"A2:K120"},
	// { graphName: "HC Legacy", graphRange:"A1:U20"},
	// { graphName: "HC Ongoing", graphRange:"A1:BH5"},
	// { graphName: "Non HC", graphRange: "A1:N30" },
	// { graphName: "Units Budget", graphRange: "A1: I15" },
	// { graphName: "Custome Dashboard", graphRange: "A1:Q40"},
  {
    graphRef: 'MR Budget',
    graphRange: 'A1: I15',
    smartSheetRef: 8291048758765444,
    nonFormularRef: {
      smartSheet: ['Description|0:Month 12|65'],
      graph: ['A|1:P|74']
    }
  }
	// { graphName: "Control", graphRange: "A1:Z69" }
];
arrayAlpha = 'abcdefghijklmnopqrstuvwxyz'.toUpperCase().split('');

function parseNumberToExcelColumnName(valueToPass) {
  columnName = '';
  if (valueToPass / 27 >= 1) {
    columnName = arrayAlpha[Math.floor(valueToPass / 27 - 1)];
  }
  var value = (valueToPass % 27) - 1;
  value = value < 0 ? 0 : valueToPass >= 26 ? ++value : value;
  columnName += arrayAlpha[value];
  return columnName;
}

// Get the home page.
router.get('/', (req, res) => {
	// check if user is authenticated
  if (!req.isAuthenticated()) {
    res.render('login');
  } else {
    renderUserFiles(req, res);
  }
});
// Authentication request.
router.get(
	'/login',
	passport.authenticate('azuread-openidconnect', { failureRedirect: '/' }),
	(req, res) => {
  res.redirect('/');
}
);
router.get('/allSmartSheets/:id', function (req, res, next) {
  getSmartsheet(req.params.id, data => res.json(data));
});
router.get('/updateSmartToExcel', function (req, res, next) {
  updateSmartsheetToExcel(req, res);
});
router.get('/updateRangeSStoEX', function (req, res, next) {
  updateRangeSStoEX(req, res);
});

// Authentication callback.
// After we have an access token, get user data and load the sendMail page.
router.get(
	'/token',
	passport.authenticate('azuread-openidconnect', { failureRedirect: '/' }),
	(req, res) => {
		// updateSmartsheetToExcel(req, res, 'first');
  updateRangeSStoEX(req, res, 'first');
		// res.redirect('/updateSmartToExcel');
		// graphHelper.getUserData(req.user.accessToken, (err, user) => {
		// if (!err) {
		//   req.user.profile.displayName = user.body.displayName;
		//   req.user.profile.emails = [{ address: user.body.mail || user.body.userPrincipalName }];
		//   renderSendMail(req, res);

		// }
}
);
function updateSmartsheetToExcel(req, res, first) {
  console.log(new Date());
  getSmartsheet(config.smartSheet.defaultSheet, data => {
    let editRow = data.body.rows[0].cells;
    let value = { values: [[editRow[0].value, editRow[1].value]] };
    graphHelper.updateFile(req.user.accessToken, value, (err, userFiles) => {
      if (!err) {
        if (res) {
          if (first) autoUpdate(req, res);
        }
      } else {
        renderError(err, res);
      }
    });
		// res.render('ssIntegration', data);
  });
}
function updateRangeSStoEX(req, res, first) {
  getSmartsheet(SheetsToEdit[0].smartSheetRef, data => {
    let excelLastRange = parseNumberToExcelColumnName(data.body.columns.length);
    let lastRow = data.body.rows.length;
    excelLastRange += lastRow;
    let values = data.body.rows.map(item => item.cells.map(cellItems => cellItems.value || ''));
    values = { values };
    graphHelper.updateFile(
			req.user.accessToken,
			values,
			excelLastRange,
			'MR Budget',
			(err, userFiles) => {
  if (!err) {
    if (res) {
						// pausing autoUpdater for now
      if (first) autoUpdate(req, res);
    }
  } else {
    renderError(err, res);
  }
}
		);

		// let columnNumFrom = data.body.columns.filter(
		// 	item => item.title == SheetsToEdit[0].nonFormularRef.smartSheet[0].split(':')[0].split('|')[0]
		// )[0].index;
		// let columnNumTo = data.body.columns.filter(
		// 	item => item.title == SheetsToEdit[0].nonFormularRef.smartSheet[0].split(':')[1].split('|')[0]
		// )[0].index;
		// let rowStart = SheetsToEdit[0].nonFormularRef.smartSheet[0].split(':')[0].split('|')[1];
		// let rowEnd = SheetsToEdit[0].nonFormularRef.smartSheet[0].split(':')[1].split('|')[1];
		// // console.log(columnNumFrom, columnNumTo, rowStart, rowEnd);
		// let valueToEdit = [];
		// for (let rS = rowStart; rS <= rowEnd; rS++) {
		//   for (let cS = columnNumFrom; cS <= columnNumTo; cS++) {
		//     let value = data.body.rows[rS].cells[cS].value || '';
		//     valueToEdit.push(value);
		//   }
		// }
		// console.log(valueToEdit);
		// res.json(data.body);
  });
}
function autoUpdate(req, res) {
  setInterval(() => updateSmartsheetToExcel(req, res), 10000);
}

// // Load the sendMail page.
// function renderSendMail(req, res) {
//   res.render('sendMail', {
//     display_name: req.user.profile.displayName,
//     email_address: req.user.profile.emails[0].address
//   });
// }
function renderUserFiles(userFiles, res) {
  console.log(userFiles);
  res.render('userFiles', { userFiles: userFiles });
}
function getSmartsheet(id, callback) {
  smartsheetHelper.sendGetRequest('', 'https://api.smartsheet.com/2.0/sheets/' + id, function (
    err,
    data
	) {
    if (err) res.send(err);
    callback(data);
  });
}
// Do prep before building the email message.
// The message contains a file attachment and embeds a sharing link to the file in the message body.
function prepForEmailMessage(req, callback) {
  const accessToken = req.user.accessToken;
  const displayName = req.user.profile.displayName;
  const destinationEmailAddress = req.body.default_email;
	// Get the current user's profile photo.
  graphHelper.getProfilePhoto(accessToken, (errPhoto, profilePhoto) => {
		// //// TODO: MSA flow with local file (using fs and path?)
    if (!errPhoto) {
			// Upload profile photo as file to OneDrive.
      graphHelper.uploadFile(accessToken, profilePhoto, (errFile, file) => {
				// Get sharingLink for file.
        graphHelper.getSharingLink(accessToken, file.id, (errLink, link) => {
          const mailBody = emailer.generateMailBody(
						displayName,
						destinationEmailAddress,
						link.webUrl,
						profilePhoto
					);
          callback(null, mailBody);
        });
      });
    } else {
      var fs = require('fs');
      var readableStream = fs.createReadStream('public/img/test.jpg');
      var picFile;
      var chunk;
      readableStream.on('readable', function () {
        while ((chunk = readableStream.read()) != null) {
          picFile = chunk;
        }
      });

      readableStream.on('end', function () {
        graphHelper.uploadFile(accessToken, picFile, (errFile, file) => {
					// Get sharingLink for file.
          graphHelper.getSharingLink(accessToken, file.id, (errLink, link) => {
            const mailBody = emailer.generateMailBody(
							displayName,
							destinationEmailAddress,
							link.webUrl,
							picFile
						);
            callback(null, mailBody);
          });
        });
      });
    }
  });
}

// get all data

router.get('/getFiles', (req, res) => {
  graphHelper.getFileDetails(req.user.accessToken, (errFiles, allFiles) => {
    console.log(allFiles);
  });
});

// Send an email.
router.post('/sendMail', (req, res) => {
  const response = res;
  const templateData = {
    display_name: req.user.profile.displayName,
    email_address: req.user.profile.emails[0].address,
    actual_recipient: req.body.default_email
  };
  prepForEmailMessage(req, (errMailBody, mailBody) => {
    if (errMailBody) renderError(errMailBody);
    graphHelper.postSendMail(req.user.accessToken, JSON.stringify(mailBody), errSendMail => {
      if (!errSendMail) {
        response.render('sendMail', templateData);
      } else {
        if (hasAccessTokenExpired(errSendMail)) {
          errSendMail.message += ' Expired token. Please sign out and sign in again.';
        }
        renderError(errSendMail, response);
      }
    });
  });
});

router.get('/disconnect', (req, res) => {
  req.session.destroy(() => {
    req.logOut();
    res.clearCookie('graphNodeCookie');
    res.status(200);
    res.redirect('/');
  });
});

// helpers
function hasAccessTokenExpired(e) {
  let expired;
  if (!e.innerError) {
    expired = false;
  } else {
    expired =
			e.forbidden &&
			e.message === 'InvalidAuthenticationToken' &&
			e.response.error.message === 'Access token has expired.';
  }
  return expired;
}

function renderError(e, res) {
  e.innerError = e.response ? e.response.text : '';
  res.render('error', {
    error: e
  });
}

module.exports = router;
