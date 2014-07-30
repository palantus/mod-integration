var request = require('request');
var Spreadsheet = require('edit-google-spreadsheet');

var IntegrationModule = function () {
};

IntegrationModule.prototype.init = function(fw, onFinished) {
    this.fw = fw;
	onFinished.call(this);
}

IntegrationModule.prototype.onMessage = function (req, callback) {
	switch(req.body.type){
		case "sheetsquery" :

			if(typeof(req.body.sheet) !== "string" || typeof(req.body.query) !== "string"){
				callback({error: "Invalid sheet request"});
				return;
			}

			request.get(
			    'https://docs.google.com/spreadsheets/d/' + req.body.sheet + '/gviz/tq?tqx=out:json&tq=' + req.body.query,
			    function (error, response, body) {
			        if (!error && response.statusCode == 200) {
			        	var jsonBody = body.substring("google.visualization.Query.setResponse(".length, body.length - 2);
			            try{
					        callback(JSON.parse(jsonBody).table);
					    }catch(e){
					    	callback({error: "Unknown sheet parse error"});
					    }
			        	
			        } else {
			        	callback({error: "Unknown sheet error"});
			        }
			    }
			);

			break;
		case "sheetsupdate" :
			if(typeof(req.body.sheet) !== "string" || isNaN(req.body.row) || isNaN(req.body.col) || req.body.value === undefined){
				callback({error: "Invalid sheet update request"});
				return;
			}

			Spreadsheet.load({
					debug:true,
					spreadsheetId: req.body.sheet,
    				worksheetId: 'od6',
				    username: req.body.username,
				    password: req.body.password
		    	},
				function sheetReady(err, spreadsheet) {
					if(err){
						console.log("Spreadsheet error:");
						console.log(err);
						callback({error: err});
						return;
					}

					var toAdd = {};
					toAdd[req.body.row] = {};
					toAdd[req.body.row][req.body.col] = req.body.value;

				    spreadsheet.add(toAdd);

				    spreadsheet.send(function(err) {
						if(err) {
							console.log("Spreadsheet error:");
							console.log(err);
							callback({error: err});
							return;
						}

						callback({success:true});
				    });
				}
			);

			break;
	}
};
 
module.exports = IntegrationModule;