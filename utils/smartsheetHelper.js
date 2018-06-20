const request = require("superagent");
const config = require("../utils/config");
function sendGetRequest(query, url, callback) {
	request
		.get(url)
		.set("Authorization", "Bearer " + config.smartSheet.creds.clientSecret)
		.set("Content-type", "application/json")
		.send(query)
		.end((err, res) => {
			callback(err, res);
		});
}
exports.sendGetRequest = sendGetRequest;
