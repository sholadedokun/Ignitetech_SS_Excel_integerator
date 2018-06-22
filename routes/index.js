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
		includeSSColumn: 0,
		smartSheetRef: 6602198898501508
	},
	{
		graphRef: 'HC Legacy',
		includeSSColumn: 1,
		rangeEdit: 0,
		smartSheetRef: 1535649317709700
	},
	{
		graphRef: 'HC Ongoing',
		includeSSColumn: 1,
		smartSheetRef: 8088111688247172
	},
	{
		graphRef: 'Non HC',
		includeSSColumn: 0,
		smartSheetRef: 5248700218926980
	},
	{
		graphRef: 'Units Budget',
		includeSSColumn: 1,
		smartSheetRef: 6039248945080196
	},
	{
		graphRef: 'Customer Dashboard',
		includeSSColumn: 1,
		smartSheetRef: 3787449131394948
	},
	{
		graphRef: 'MR Budget',
		includeSSColumn: 1,
		smartSheetRef: 8291048758765444
	}
];

// Get the home page.
router.get('/', (req, res) => {
	// check if user is authenticated
	if (!req.isAuthenticated()) {
		res.render('login');
	} else {
		updateRangeSStoEX(req, res, 'first');
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
router.get('/allSmartSheets/:id', function(req, res, next) {
	getSmartsheet(req.params.id, data => res.send(data.text));
});

router.get('/updateRangeSStoEX', function(req, res, next) {
	updateRangeSStoEX(req, res);
});

// Authentication callback.
// After we have an access token, get user data and load the sendMail page.
router.get(
	'/token',
	passport.authenticate('azuread-openidconnect', { failureRedirect: '/' }),
	(req, res) => {
		updateRangeSStoEX(req, 'first');
	}
);

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
		// res.render('ssIntegration',data)
	});
}
function sendValueToLiveEdit(req, values, lastRange, firstRange, graphRef) {
	return new Promise(function(resolve) {
		graphHelper.updateFile(req.user.accessToken, values, firstRange, lastRange, graphRef, err => {
			if (!err) {
				resolve('successfully update at ' + new Date());
			} else {
				console.log(err);
				// renderError(err, res);
			}
		});
	});
}

function updateRangeSStoEX(req, first) {
	console.log(new Date());
	if (first == 'first') autoUpdate(req);
	for (let a = 0; a < SheetsToEdit.length; a++) {
		getSmartsheet(SheetsToEdit[a].smartSheetRef, data => {
			// copy all the values
			if (data) {
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
						sending();
						async function sending() {
							let result = await sendValueToLiveEdit(
								req,
								RangeVal,
								mGraphRange[1],
								mGraphRange[0],
								SheetsToEdit[a].graphRef
							);
							console.log(result);
						}
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
					sending();
					async function sending() {
						let result = await sendValueToLiveEdit(
							req,
							values,
							mGraphLastRange,
							mGraphStartRange,
							SheetsToEdit[a].graphRef
						);
						console.log(result);
					}
				}
			}
		});
	}
}
function autoUpdate(req) {
	setInterval(() => updateRangeSStoEX(req), 60000);
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
