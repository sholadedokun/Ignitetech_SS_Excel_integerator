import axios from 'axios';
import { SERVER_STARTED, SERVER_ERROR } from './actionTypes';

const ROOT_URL = 'http://example.herokuapp.com/api';
export function startServer(payload) {
	payload.token = localStorage.getItem('IgniteTechToken');
	return function(dispatch) {
		return new Promise(resolve => {
			axios
				.post(`${ROOT_URL}/start`, payload)
				.then(response => {
					// If request is good...
					// - Update state to indicate server is started
					dispatch({ type: SERVER_STARTED, payload: response });
					resolve(response);
				})
				.catch(() => {
					// If request is bad...
					// - Show an error to the user
					dispatch(serverError('Error starting Server, Please try again.'));
				});
		});
	};
}

export function serverError(error) {
	return {
		type: SERVER_ERROR,
		payload: error
	};
}
