/*
 * Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

const request = require("superagent");

function getFileDetails(accessToken, callback) {
	request
		.get("https://graph.microsoft.com/beta/me/drive/items/01LT4ZEQOEI5X4RS4ERNF2GNU244SO2DSA/workbook/worksheets('Sheet1')/range(address='A1:E30')?$select=values")
		// .send({ type: 'view' })
		.set("Authorization", "Bearer " + accessToken)
		.end((err, res) => {
			// Returns 200 OK and the permission with the link in the body.
			callback(err, res.body);
		});
}
function updateFile(accessToken, value, firstRange, lastRange, sheet, callback) {
	request
		.patch(`https://graph.microsoft.com/beta/me/drive/items/01LT4ZEQNG2XUHMKCQ6FEZGWOPKRCWBRDZ/workbook/worksheets('${sheet}')/range(address='${firstRange}:${lastRange}')`)
		.send(value)
		.set("Authorization", "Bearer " + accessToken)

		.end((err, res) => {
			// Returns 200 OK and the file metadata in the body.
			callback(err);
		});
}
function postSendMail(accessToken, message, callback) {
	request
		.post("https://graph.microsoft.com/beta/me/sendMail")
		.send(message)
		.set("Authorization", "Bearer " + accessToken)
		.set("Content-Type", "application/json")
		.set("Content-Length", message.length)
		.end((err, res) => {
			// Returns 202 if successful.
			// Note: If you receive a 500 - Internal Server Error
			// while using a Microsoft account (outlook.com, hotmail.com or live.com),
			// it's possible that your account has not been migrated to support this flow.
			// Check the inner error object for code 'ErrorInternalServerTransientError'.
			// You can try using a newly created Microsoft account or contact support.
			callback(err, res);
		});
}
exports.getFileDetails = getFileDetails;
exports.updateFile = updateFile;
exports.postSendMail = postSendMail;
