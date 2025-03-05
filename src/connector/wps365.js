/*
	***** BEGIN LICENSE BLOCK *****
	
	Copyright © 2022 Corporation for Digital Scholarship
                     Vienna, Virginia, USA
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
(function() {

var isTopWindow = false;
if(window.top) {
	try {
		isTopWindow = window.top == window;
	} catch(e) {};
}
if (!isTopWindow) return;

Zotero.WPS365 = {
	config: {
		noteInsertionPlaceholderURL: 'https://www.zotero.org/?',
		fieldURL: 'https://www.zotero.org/google-docs/?',
		brokenFieldURL: 'https://www.zotero.org/google-docs/?broken=',
		fieldKeyLength: 6,
		citationPlaceholder: "{Updating}",
		fieldPrefix: "Z_F",
		dataPrefix: "Z_D",
		biblStylePrefix: "Z_B",
		twipsToPoints: 0.05,
	},
	clients: {},

	// Set to true if there are links in the doc with zotero field url. Causes the
	// download intercept warning to show telling users to unlink citations first before download
	hasZoteroCitations: false,
	// Prevent the download interception warning
	downloadInterceptBlocked: false,
	// Prevents the download intercept dialog from showing once if the user confirms they
	// want to download the document anyway
	downloadIntercepted: false,

	name: "Zotero WPS 365 Plugin",
	updateBatchSize: 32,

	init: async function() {
		if (!await Zotero.Prefs.getAsync('integration.WPS365.enabled')) return;
		if (!await Zotero.Prefs.getAsync('integration.WPS365.useWPS365API')) {
			Zotero.WPS365.Client = Zotero.WPS365.ClientAppsScript;
		}
		await Zotero.Inject.loadReactComponents();
		if (Zotero.isBrowserExt) {
			await Zotero.Connector_Browser.injectScripts(['zotero-google-docs-integration/ui.js']);
		}
		Zotero.WPS365.UI.init();
		window.addEventListener(`${Zotero.WPS365.name}.call`, async function(e) {
			var client = Zotero.WPS365.clients[e.data.client.id];
			if (!client) {
				client = new Zotero.WPS365.Client();
				await client.init();
			}
			client.call.apply(client, e.data.args);
		});
	},

	execCommand: async (command, client, showOrphanedCitationAlert=true) => {
		if (Zotero.WPS365.UI.isDocx) {
			return Zotero.WPS365.UI.displayDocxAlert();
		}
		if (!client) {
			client = new Zotero.WPS365.Client();
			await client.init();
		}

		if (command == 'addEditCitation') {
			// Check if we're in a broken field and cancel operation if user
			// wants to click More Info
			try {
				await client.cursorInField(showOrphanedCitationAlert);
			} catch (e) {
				if (e.message != "Handled Error") {
					Zotero.logError(e);
				}
				return;
			}
		}

		window.dispatchEvent(new MessageEvent('Zotero.Integration.execCommand', {
			data: {client: {documentID: client.documentID, name: Zotero.WPS365.name, id: client.id}, command}
		}));
		this.lastClient = client;
	},

	respond: function(client, response) {
		window.dispatchEvent(new MessageEvent('Zotero.Integration.respond', {
			data: {client: {documentID: client.documentID, name: Zotero.WPS365.name, id: client.id}, response}
		}));
	},

	editField: async function() {
		// Use the last client with a cached field list to speed up the cursorInField() lookup
		var client = this.lastClient || new Zotero.WPS365.Client();
		await client.init();
		try {
			var field = await client.cursorInField(true);
		} catch (e) {
			if (e.message == "Handled Error") {
				Zotero.debug('Handled Error in editField()');
				return;
			}
			Zotero.debug(`Exception in editField()`);
			Zotero.logError(e);
			return client.displayAlert(e.message, 0, 0);
		}
		// Remove lastClient fields to ensure execCommand calls receive fresh fields
		if (this.lastClient) {
			if (this.lastClient.resetGoogleDocument) {
				this.lastClient.resetGoogleDocument();
			}
			else {
				delete this.lastClient.fields;
			}
		}
		
		if (field && field.code.indexOf("BIBL") == 0) {
			return Zotero.WPS365.execCommand("addEditBibliography", client);
		} else {
			return Zotero.WPS365.execCommand("addEditCitation", client, false);
		}
	},
};
	
if (document.readyState !== "complete") {
	window.addEventListener("load", function(e) {
		if (e.target !== document) return;
		Zotero.WPS365.init();
	}, false);
} else {
	Zotero.WPS365.init();
}

})();
