import React from 'react';
import { Grid, Row, Col } from 'react-bootstrap';
import socketIOClient from 'socket.io-client';
import { connect } from 'react-redux';
export default class Home extends React.Component {
	constructor() {
		super();
		this.state = {
			response: false,
			endpoint: 'http://127.0.0.1:3000/',
			stopServer: false
		};
	}
	componentDidMount() {
		const { endpoint } = this.state;
		const socket = socketIOClient(endpoint);
		socket.on('FromAPI', data => this.setState({ response: data }));
	}
	stopServer() {
		this.setState({ ...this.state, stopServer: !this.state.stopServer });
	}
	render() {
		return (
			<Col>
				<div>{this.state.response}</div>
				{this.state.stopServer ? (
					<button onClick={this.stopServer.bind(this)}>Resume Server</button>
				) : (
					<button onClick={this.stopServer.bind(this)}>Pause Server</button>
				)}
			</Col>
		);
	}
}
