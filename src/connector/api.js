/*
	***** BEGIN LICENSE BLOCK *****
	
	Copyright © 2017 Center for History and New Media
					George Mason University, Fairfax, Virginia, USA
					http://zotero.org
	
	This file is part of Zotero.
	
	Zotero is free software: you can redistribute it and/or modify
	it under the terms of the GNU Affero General Public License as published by
	the Free Software Foundation, either version 3 of the License, or
	(at your option) any later version.
	
	Zotero is distributed in the hope that it will be useful,
	but WITHOUT ANY WARRANTY; without even the implied warranty of
	MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
	GNU Affero General Public License for more details.

	You should have received a copy of the GNU Affero General Public License
	along with Zotero.  If not, see <http://www.gnu.org/licenses/>.
	
	***** END LICENSE BLOCK *****
*/

const TABS_CRITERIA_METHODS = new Set(['replaceNamedRangeContent', 'replaceAllText', 'deleteNamedRange'])
const TAB_ID_METHODS = new Set(['deletePositionedObject', 'replaceImage', 'updateDocumentStyle', 'deleteHeader', 'deleteFooter', 'location', 'range'])
const TAB_ID_PARAMS = new Set(['location', 'range'])

Zotero.WPS365 = Zotero.WPS365 || {};

Zotero.WPS365.API = {
	authDeferred: null,
	authCredentials: {},
	apiVersion: 6,
	
	init: async function() {
		this.authCredentials = await Zotero.Utilities.Connector.createMV3PersistentObject('WPS365AuthCredentials')
	},
	
	resetAuth: function() {
		delete this.authCredentials.headers;
		delete this.authCredentials.lastEmail;
	},

	getAuthHeaders: async function() {
		// Delete headers if expired which will cause a refetch
		if (Zotero.WPS365.API.authCredentials.expiresAt && Date.now() > Zotero.WPS365.API.authCredentials.expiresAt) {
			delete Zotero.WPS365.API.authCredentials.headers;
		}
		if (Zotero.WPS365.API.authCredentials.headers) {
			return Zotero.WPS365.API.authCredentials.headers;
		}
		
		// For macOS, since popping up an auth window or calling Connector_Browser.bringToFront()
		// doesn't move the progress window to the back
		Zotero.Connector.callMethod('sendToBack');
		
		// Request OAuth2 access token
		let params = {
			client_id: ZOTERO_CONFIG.OAUTH.GOOGLE_DOCS.CLIENT_KEY,
			redirect_uri: ZOTERO_CONFIG.OAUTH.GOOGLE_DOCS.CALLBACK_URL,
			response_type: 'token',
			scope: 'https://www.googleapis.com/auth/documents email',
			// Will be enabled by Google on June 17, 2024. Uncomment for testing
			// enable_granular_consent: "true",
			state: 'google-docs-auth-callback'
		};
		if (Zotero.WPS365.API.authCredentials.lastEmail) {
			params.login_hint = Zotero.WPS365.API.authCredentials.lastEmail;
		}
		let url = ZOTERO_CONFIG.OAUTH.GOOGLE_DOCS.AUTHORIZE_URL + "?";
		for (let key in params) {
			url += `${key}=${encodeURIComponent(params[key])}&`;
		}
		Zotero.Connector_Browser.openWindow(url, {type: 'normal', onClose: Zotero.WPS365.API.onAuthCancel});
		this.authDeferred = Zotero.Promise.defer();
		return this.authDeferred.promise;
	},
	
	onAuthComplete: async function(url, tab) {
		// close auth window
		// ensure that tab close listeners don't have a promise they can reject
		let deferred = this.authDeferred;
		this.authDeferred = null;
		if (Zotero.isBrowserExt) {
			browser.tabs.remove(tab.id);
		} else if (Zotero.isSafari) {
			Zotero.Connector_Browser.closeTab(tab);
		}
		try {
			var uri = new URL(url);
			var params = {};
			for (let keyvalue of uri.hash.split('&')) {
				let [key, value] = keyvalue.split('=');
				params[key] = decodeURIComponent(value);
			}
			let error = params.error || params['#error'];
			if (error) {
				if (error === 'access_denied') {
					throw new Error(`Google Auth permission to access WPS 365 not granted`);
				}
				else {
					throw new Error(error);
				}
			}
			
			if (!params.scope.includes("https://www.googleapis.com/auth/documents")) {
				throw new Error(`Google Auth permission to access WPS 365 not granted`);
			}
			
			url = ZOTERO_CONFIG.OAUTH.GOOGLE_DOCS.ACCESS_URL
				+ `?access_token=${params.access_token}`;
			let xhr = await Zotero.HTTP.request('GET', url);
			let response = JSON.parse(xhr.responseText);
			if (response.aud != ZOTERO_CONFIG.OAUTH.GOOGLE_DOCS.CLIENT_KEY) {
				throw new Error(`WPS 365 Access Token invalid ${xhr.responseText}`);
			}
			
			this.authCredentials.lastEmail = response.email;
			this.authCredentials.headers = {'Authorization': `Bearer ${params.access_token}`};
			this.authCredentials.expiresAt = Date.now() + (parseInt(params.expires_in)-60)*1000;
			response = await this.getAuthHeaders();
			deferred.resolve(response);
			return response;
		} catch (e) {
			return deferred.reject(e);
		}
	},
	
	onAuthCancel: function() {
		let error = new Error('WPS 365 authorization was cancelled');
		error.type = "Alert";
		Zotero.WPS365.API.authDeferred
			&& Zotero.WPS365.API.authDeferred.reject(error);
	},
	
	run: async function(documentSpecifier, method, args, tab) {
		// If not an array, discard or the docs script spews errors.
		if (! Array.isArray(args)) {
			args = [];
		}
		let headers;
		try {
			headers = await this.getAuthHeaders();
		}
		catch (e) {
			if (e.message.includes('not granted')) {
				this.displayPermissionsNotGrantedPrompt(tab)
				throw new Error('Handled Error');
			}
			else {
				throw e;
			}
		}
		headers["Content-Type"] = "application/json";
		var body = {
			function: 'callMethod',
			parameters: [documentSpecifier, method, args, Zotero.WPS365.API.apiVersion],
			devMode: ZOTERO_CONFIG.GOOGLE_DOCS_DEV_MODE
		};
		try {
			var xhr = await Zotero.HTTP.request('POST', ZOTERO_CONFIG.GOOGLE_DOCS_API_URL,
				{headers, body, timeout: null});
		} catch (e) {
			if (e.status >= 400 && e.status < 404) {
				this.resetAuth();
				this.displayWrongAccountPrompt();
				throw new Error('Handled Error');
			} else {
				throw new Error(`${e.status}: WPS 365 request failed.\n\n${e.responseText}`);
			}
		}
		var responseJSON = JSON.parse(xhr.responseText);
		
		if (responseJSON.error) {
			// For some reason, sometimes the still valid auth token starts being rejected
			if (responseJSON.error.details[0].errorMessage == "Authorization is required to perform that action.") {
				delete this.authCredentials.headers;
				return this.run(documentSpecifier, method, args);
			}
			var err = new Error(responseJSON.error.details[0].errorMessage);
			err.stack = responseJSON.error.details[0].scriptStackTraceElements;
			err.type = `WPS 365 ${responseJSON.error.message}`;
			throw err;
		}
		
		let resp = await this.handleResponseErrors(responseJSON, arguments, tab);
		if (resp) {
			return resp;
		}
		var response = responseJSON.response.result && responseJSON.response.result.response;
		if (responseJSON.response.result.debug) {
			Zotero.debug(`WPS 365 debug:\n\n${responseJSON.response.result.debug.join('\n\n')}`);
		}
		return response;
	},
	
	handleResponseErrors: async function(responseJSON, args, tab) {
		var lockError = responseJSON.response.result.lockError;
		if (lockError) {
			if (await this.displayLockErrorPrompt(lockError, tab)) {
				await this.run(args[0], "unlockTheDoc", [], args[3]);
				return this.run.apply(this, args);
			} else {
				throw new Error('Handled Error');
			}
		}
		var docAccessError = responseJSON.response.result.docAccessError;
		if (docAccessError) {
			this.resetAuth();
			this.displayWrongAccountPrompt();
			throw new Error('Handled Error');
		}
		var genericError = responseJSON.response.result.error;
		if (genericError) {
			Zotero.logError(new Error(`Non-fatal WPS 365 Error: ${genericError}`));
		}
	},

	displayLockErrorPrompt: async function(error, tab) {
		var message = Zotero.getString('integration_WPS365_documentLocked', ZOTERO_CONFIG.CLIENT_NAME);
		var result = await Zotero.Messaging.sendMessage('confirm', {
			title: ZOTERO_CONFIG.CLIENT_NAME,
			button2Text: "",
			button3Text: Zotero.getString('general_needHelp'),
			message
		}, tab);
		if (result.button != 3) return;

		message = Zotero.getString('integration_WPS365_documentLocked_moreInfo', ZOTERO_CONFIG.CLIENT_NAME);
		
		var result = await Zotero.Messaging.sendMessage('confirm', {
			title: ZOTERO_CONFIG.CLIENT_NAME,
			button1Text: Zotero.getString('general_yes'),
			button2Text: Zotero.getString('general_no'),
			message
		}, tab);
		return result.button == 1;
	},
	
	displayPermissionsNotGrantedPrompt: async function(tab) {
		var message = Zotero.getString('integration_WPS365_authScopeError', ZOTERO_CONFIG.CLIENT_NAME);
		var result = await Zotero.Messaging.sendMessage('confirm', {
			title: ZOTERO_CONFIG.CLIENT_NAME,
			button2Text: "",
			button3Text: Zotero.getString('general_moreInfo'),
			message
		}, tab);
		if (result.button != 3) return;
		Zotero.Connector_Browser.openTab('https://www.zotero.org/support/google_docs#authorization');
	},
	
	displayWrongAccountPrompt: async function(tab) {
		var message = Zotero.getString('integration_WPS365_documentPermissionError', ZOTERO_CONFIG.CLIENT_NAME);
		var result = await Zotero.Messaging.sendMessage('confirm', {
			title: ZOTERO_CONFIG.CLIENT_NAME,
			button2Text: "",
			button3Text: Zotero.getString('general_moreInfo'),
			message
		}, tab);
		if (result.button != 3) return;
		Zotero.Connector_Browser.openTab('https://www.zotero.org/support/google_docs#authorization');
	},

	getDocument: async function (docID, tabID=null) {
		var headers = await this.getAuthHeaders();
		headers["Content-Type"] = "application/json";
		try {
			var xhr = await Zotero.HTTP.request('GET', `https://docs.googleapis.com/v1/documents/${docID}?includeTabsContent=true`,
				{headers, timeout: 60000});
		} catch (e) {
			if (e.status == 403) {
					this.resetAuth();
					throw new Error(`${e.status}: WPS 365 Authorization failed. Try again.\n${e.responseText}`);
				} else {
					throw new Error(`${e.status}: WPS 365 request failed.\n\n${e.responseText}`);
				}
		}
		
		let document = JSON.parse(xhr.responseText);
		if (!document.tabs) return document;
		for (let tab of document.tabs) {
			// Return first tab if not specified
			if (tabID === null || tab.tabProperties.tabId == tabID) {
				let documentTab = tab.documentTab;
				documentTab.documentId = docID;
				documentTab.tabId = tabID;
				return documentTab;
			}
		}
		return document;
	},
	
	_addTabDataToObject(object, tabId) {
		const key = Object.keys(object)[0];
		if (TABS_CRITERIA_METHODS.has(key)) {
			object.tabsCriteria = { tabIds: [tabId] }
			return;
		}
		else if (TAB_ID_METHODS.has(key)) {
			object.tabId = tabId;
		}
		for (let k in object[key]) {
			if (TAB_ID_PARAMS.has(k)) {
				object[key][k].tabId = tabId;
			}
		}
	},

	batchUpdateDocument: async function (docId, tabId=null, body) {
		var headers = await this.getAuthHeaders();
		if (tabId) {
			for (let request of body.requests) {
				request = this._addTabDataToObject(request, tabId);
			}
		}
		try {
			var xhr = await Zotero.HTTP.request('POST', `https://docs.googleapis.com/v1/documents/${docId}:batchUpdate`,
				{headers, body, timeout: 60000});
		} catch (e) {
			if (e.status == 403) {
				this.resetAuth();
				throw new Error(`${e.status}: WPS 365 Authorization failed. Try again.\n${e.responseText}`);
			} else {
				throw new Error(`${e.status}: WPS 365 request failed.\n\n${e.responseText}`);
			}
		}

		return JSON.parse(xhr.responseText);
	}
};

Zotero.WPS365_API = Zotero.WPS365.API;
