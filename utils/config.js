/*
 * Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

exports.graph = {
  creds: {
    redirectUrl: '##########',
    clientID: '#######',
    clientSecret: '#########',
    identityMetadata:
			'https://login.microsoftonline.com/common/v2.0/.well-known/openid-configuration',
    allowHttpForRedirectUrl: true, // For development only
    responseType: 'code',
    validateIssuer: false, // For development only
    responseMode: 'query',
    scope: ['User.Read', 'Mail.Send', 'Files.ReadWrite']
  }
};
exports.smartSheet = {
  creds: {
    clientSecret: '#########'
  },
  defaultSheet: '##############'
};
