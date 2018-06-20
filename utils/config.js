/*
 * Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

exports.graph = {
	creds: {
		redirectUrl: "http://localhost:3000/token",
		clientID: "266809ae-388c-4341-8243-9f98f189e2f1",
		clientSecret: "rynaXZHY803}++crwQTX52{",
		identityMetadata: "https://login.microsoftonline.com/common/v2.0/.well-known/openid-configuration",
		allowHttpForRedirectUrl: true, // For development only
		responseType: "code",
		validateIssuer: false, // For development only
		responseMode: "query",
		scope: ["User.Read", "Mail.Send", "Files.ReadWrite"]
	}
};
exports.smartSheet = {
	creds: {
		clientSecret: "oo0x7b6c9w46eod0xforn5c6fy"
	},
	defaultSheet: "3520439604537220"
};
