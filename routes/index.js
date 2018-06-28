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
const arrayAlpha = 'abcdefghijklmnopqrstuvwxyz'.toUpperCase().split('');
const SheetsToEdit = require('../utils/sheetList').sheetsToEdit;
let fs = require('fs');
let path = require('path');
let fileDB = require('../database/lastModified');
//so we could clear interval later
let interval;
let continueUpdate = true;

// Get the home page.
router.get('/', (req, res) => {
	// check if user is authenticated
	if (!req.isAuthenticated()) {
		res.render('login');
	} else {
		updateRangeSStoEX(req, res, 'first');
	}
});
router.get('/stop', (req, res) => {
	clearInterval(interval);
	res.send({ message: 'server stopped' });
});
// Authentication request.
router.get(
	'/login',
	passport.authenticate('azuread-openidconnect', { failureRedirect: '/' }),
	(req, res) => {
		res.redirect('/');
	}
);
//to get all the data of a particular smartsheet that has the ID
router.get('/allSmartSheets/:id', function(req, res, next) {
	getSmartsheet(req.params.id, data => res.send(data.text));
});
//this route can be called after authentication
router.get('/updateRangeSStoEX', function(req, res, next) {
	updateRangeSStoEX(req, res);
});

// Authentication callback.
// After we have an access token, we start the update process
router.get(
	'/token',
	passport.authenticate('azuread-openidconnect', { failureRedirect: '/' }),
	(req, res) => {
		//we sending 'first' so as to startup the interval timer.
		updateRangeSStoEX(req, 'first', res);
	}
);

//use for converting to excel column Name references like C78 or AB54
function parseNumberToExcelColumnName(valueToPass) {
	//algorithm needs review ... it breaks at some point
	columnName = '';
	if (valueToPass / 27 >= 1) {
		columnName = arrayAlpha[Math.floor(valueToPass / 27 - 1)];
	}
	var value = (valueToPass % 27) - 1;
	value = value < 0 ? 0 : valueToPass >= 26 ? ++value : value;
	columnName += arrayAlpha[value];
	return columnName;
}
// function updateSmartsheetToExcel(req, res, first) {
// 	getSmartsheet(config.smartSheet.defaultSheet, data => {
// 		let editRow = data.body.rows[0].cells;
// 		let value = { values: [[editRow[0].value, editRow[1].value]] };
// 		graphHelper.updateFile(req.user.accessToken, value, (err, userFiles) => {
// 			if (!err) {
// 				if (res) {
// 					if (first) autoUpdate(req, res);
// 				}
// 			} else {
// 				renderError(err, res);
// 			}
// 		});
// 		// res.render('ssIntegration',data)
// 	});
// }
//to call the microsoft graph API...
function sendValueToLiveEdit(req, values, lastRange, firstRange, graphRef, sheetID, lastModified) {
	return new Promise(function(resolve, reject) {
		//calls the microsoft API
		graphHelper.updateFile(req.user.accessToken, values, firstRange, lastRange, graphRef, err => {
			if (!err) {
				//update the json database with the edit timestamp
				updateDatabase(sheetID, lastModified, () =>
					resolve(`successfully update ${graphRef} at ${new Date()}`)
				);
			} else {
				console.log(err);
				// renderError(err, res);
			}
		});
	});
}

function checkIfThereAreNewChanges(sheetID, lastModified) {
	if (fileDB) {
		//if database Exists
		if (fileDB[sheetID]) {
			return fileDB[sheetID] == lastModified ? false : true;
		} else {
			return true; //the sheet is new so update.
		}
	} else {
		// if database doesn't exist
		return true;
	}
}
function updateDatabase(sheetID, lastModified, next) {
	//check it file doesn't exist or lastmodified is not equal
	if (!fileDB[sheetID] || fileDB[sheetID] != lastModified) {
		fileDB[sheetID] = lastModified;
		//call the Node file system (fs) to write the new database
		fs.writeFile(
			path.join(__dirname, '../database/lastModified.json'),
			JSON.stringify(fileDB),
			function(err) {
				if (err) console.log(err);
				console.log('Saved!');
				next();
			}
		);
	}
}

function updateRangeSStoEX(req, first, res) {
	console.log('Checking for new updates ... ' + new Date());
	//if this is the first time call the setInterval
	if (first == 'first') autoUpdate(req);
	for (let a = 0; a < SheetsToEdit.length; a++) {
		//retrieve smartsheet data
		getSmartsheet(SheetsToEdit[a].smartSheetRef, data => {
			if (data) {
				if (checkIfThereAreNewChanges(data.body.id, data.body.modifiedAt)) {
					console.log('Pass check new update', data.body.id);
					let values = data.body.rows.map(item =>
						item.cells.map(cellItems => cellItems.value || '')
					);
					let mGraphStartRange = 'A1'; //initally set to paste to the begining of excel document
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
							sending();
							//using asyn and wait because of non-blocking nature of Node... this will help pause the code till the last edit was done
							async function sending() {
								let result = await sendValueToLiveEdit(
									req,
									RangeVal,
									mGraphRange[1],
									mGraphRange[0],
									SheetsToEdit[a].graphRef,
									data.body.id,
									data.body.modifiedAt
								);
								console.log(result);
							}
						}
					} else {
						//to include the SS column name in the copying
						if (SheetsToEdit[a].includeSSColumn) {
							// now copy the column name and put them at the top of the sheet;
							let columnName = data.body.columns.map(item => item.title);
							values.unshift(columnName);
						}

						// now shift all value to the next column for easy formatting... just to make it look nicer with spacing
						values = values.map(item => {
							item.unshift(' ');
							return item;
						});
						// res.json(values);
						// get the right excel last Range to use for example AB55
						mGraphLastRange = parseNumberToExcelColumnName(values[0].length);
						mGraphLastRange += values.length;
						values = { values };
						sending();
						//using asyn and wait because of non-blocking nature of Node... this will help puase the code till the last edit was done
						async function sending() {
							let result = await sendValueToLiveEdit(
								req,
								values,
								mGraphLastRange,
								mGraphStartRange,
								SheetsToEdit[a].graphRef,
								data.body.id,
								data.body.modifiedAt
							);
							console.log(result);
						}
					}
				} else {
					console.log('there is no recent update');
				}
			}
		});
	}
}
// to set interval for the server to keep checking for updates
function autoUpdate(req) {
	// current time of 60000 ms for 1 minute
	interval = setInterval(() => updateRangeSStoEX(req), 60000);
}
async function getSmartsheet(id, callback) {
	await smartsheetHelper.sendGetRequest('', 'https://api.smartsheet.com/2.0/sheets/' + id, function(
		err,
		data
	) {
		if (err) console.log(err);
		callback(data);
	});
}

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
