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
  {
    graphRef: 'Joe Chart',
    graphRange: 'A1: I15',
    includeSSColumn: 0,
    smartSheetRef: 8150311270410116,
		/*
         --smartsheet reference--
         -[0]first column
         -[1]first row
         -[2]last column
         -[3]last row

         --graphReference--
         [0]startRange
         [1]lastRage
        */
    rangeValue: [
      {
        smartSheet: [3, 16, 22, 16],
        graph: ['C71', 'V71']
      },
      {
        smartSheet: [7, 19, 11, 19],
        graph: ['G74', 'K74']
      },
      {
        smartSheet: [19, 38, 19, 41],
        graph: ['S92', 'S95']
      },
      {
        smartSheet: [18, 54, 18, 57],
        graph: ['R108', 'R111']
      },
      {
        smartSheet: [15, 5, 16, 5],
        graph: ['V60', 'W60']
      },
      {
        smartSheet: [4, 2, 5, 2],
        graph: ['Y33', 'Z33']
      }
    ]
  },
	{
	  graphRef: 'P&L',
	  graphRange: 'A1: I15',
	  includeSSColumn: 0,
	  smartSheetRef: 6602198898501508
	},
	{
	  graphRef: 'HC Legacy',
	  includeSSColumn: 1,
	  rangeEdit: 0,
	  graphRange: 'A1: I15',
	  smartSheetRef: 1535649317709700
	},
	{
	  graphRef: 'HC Ongoing',
	  graphRange: 'A1: I15',
	  includeSSColumn: 1,
	  rangeEdit: 0,
	  smartSheetRef: 8088111688247172
	},
	{
	  graphRef: 'Non HC',
	  graphRange: 'A1: I15',
	  includeSSColumn: 0,
	  rangeEdit: 0,
	  smartSheetRef: 5248700218926980
	},
	{
	  graphRef: 'Units Budget',
	  graphRange: 'A1: I15',
	  includeSSColumn: 1,
	  rangeEdit: 0,
	  smartSheetRef: 6039248945080196
	},
	{
	  graphRef: 'Customer Dashboard',
	  graphRange: 'A1: I15',
	  includeSSColumn: 1,
	  rangeEdit: 0,
	  smartSheetRef: 3787449131394948
	},
	{
	  graphRef: 'MR Budget',
	  graphRange: 'A1: I15',
	  includeSSColumn: 1,
	  smartSheetRef: 8291048758765444
	}
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
function sendValueToLiveEdit(req, values, lastRange, firstRange, graphRef, first) {
  setTimeout(function () {
  graphHelper.updateFile(
		req.user.accessToken,
		values,
		firstRange,
		lastRange,
		graphRef,
		(err) => {
  if (!err) {
    if (first) autoUpdate(req);
    console.log('successfully update at ', new Date());
  } else {
    console.log(err);
				// renderError(err, res);
  }
}
	);},3000);
}
function updateRangeSStoEX(req, res, first) {
  console.log(new Date());
  for (let a = 0; a < SheetsToEdit.length; a++) {
    getSmartsheet(SheetsToEdit[a].smartSheetRef, data => {
			// copy all the values
      let values = data.body.rows.map(item => item.cells.map(cellItems => cellItems.value || ''));
      let mGraphStartRange = 'A1';
      let mGraphLastRange;
			// check if rangeEdit
      if (SheetsToEdit[a].rangeValue) {
        for (let x = 0; x < SheetsToEdit[a].rangeValue.length; x++) {
          let smartSheetRange = SheetsToEdit[a].rangeValue[x].smartSheet;
          let mGraphRange = SheetsToEdit[a].rangeValue[x].graph;

					// get the values needed by Row
          let RangeVal = values.filter(
						(item, index) => index >= smartSheetRange[1] && index <= smartSheetRange[3]
					);
          RangeVal = RangeVal.map(item =>
						item.filter(
							(tofilter, index) => index >= smartSheetRange[0] && index <= smartSheetRange[2]
						)
					);
          RangeVal = { values: RangeVal };
          sendValueToLiveEdit(
						req,
						RangeVal,
						mGraphRange[1],
						mGraphRange[0],
						SheetsToEdit[a].graphRef,
						first
					);
        }
      } else {
        if (SheetsToEdit[a].includeSSColumn) {
					// now copy the column name and put them at the top of the sheet;
          let columnName = data.body.columns.map(item => item.title);
          values.unshift(columnName);
        }

				// now shift all value to the next column for easy formatting
        values = values.map(item => {
          item.unshift(' ');
          return item;
        });
				// res.json(values);
				// get the right excel last Range to use for example AB55
        mGraphLastRange = parseNumberToExcelColumnName(values[0].length);
        mGraphLastRange += values.length;
        values = { values };
        sendValueToLiveEdit(
					req,
					values,
					mGraphLastRange,
					mGraphStartRange,
					SheetsToEdit[a].graphRef,
					first
				);
      }
    });
  }
}
function autoUpdate(req, res) {
  setInterval(() => updateRangeSStoEX(req, res), 60000);
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
    if (err) console.log(err);
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
