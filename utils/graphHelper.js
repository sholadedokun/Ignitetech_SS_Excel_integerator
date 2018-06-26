/*
 * Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

const request = require('superagent');

function getFileDetails(accessToken, callback) {
	// oluhsola's sheet 01LT4ZEQOEI5X4RS4ERNF2GNU244SO2DSA
	// other sheet 01LT4ZEQLUOIMUADP5YJBIT434WJRZDRD4
  request
		.get(
			"https://graph.microsoft.com/beta/me/drive/items/01LT4ZEQOEI5X4RS4ERNF2GNU244SO2DSA/workbook/worksheets('Sheet1')/range(address='A1:E30')?$select=values"
		)
		// .send({ type: 'view' })
		.set('Authorization', 'Bearer ' + accessToken)
		.end((err, res) => {
			// Returns 200 OK and the permission with the link in the body.
  callback(err, res.body);
});
}
function updateFile(accessToken, value, firstRange, lastRange, sheet, callback) {
	// test joe chart 01LT4ZEQNG2XUHMKCQ6FEZGWOPKRCWBRDZ
	// real joe chart 01LT4ZEQOEI5X4RS4ERNF2GNU244SO2DSA
  request
		.patch(
			`https://graph.microsoft.com/beta/me/drive/items/01LT4ZEQNG2XUHMKCQ6FEZGWOPKRCWBRDZ/workbook/worksheets('${sheet}')/range(address='${firstRange}:${lastRange}')`
		)
		.send(value)
		.set('Authorization', 'Bearer ' + accessToken)

		.end((err, res) => {
			// Returns 200 OK and the file metadata in the body.
  callback(err);
});
}

exports.getFileDetails = getFileDetails;
exports.updateFile = updateFile;
