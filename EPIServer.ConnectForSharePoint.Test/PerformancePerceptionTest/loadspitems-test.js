var webpage=require('webpage'), 
    system = require('system'),
	page = webpage.create(),
    loadInProgress = false, 
	t, stepNo = 0, 
	entryPageUrl, entryId, username, password,
	strSPIdPostfix = '__sharepointcontent',
	strSPLibIds = '',
	countCheck = 25,	// for checking number of loaded SP item
	waitForTimeout = 20000,
	logAll = false, // for controlling console.log
	callExit = false,	// call phantom.exit() - this clear cookie data
	timeoutNum = 5	// will try number of timeout - for controlling/using with callExit
	;

    page.onLoadStarted = function() {
        loadInProgress = true;
		logAll && console.log('Load started');
    };
	
	page.onConsoleMessage = function(msg) {
		console.log(msg);
	};	

	page.onUrlChanged = function(targetUrl) {
		console.log('URL changed to: ' + targetUrl);
		logAll && console.log('Title: ' + page.title);
		loadInProgress = false;
	};	
 
    // page.onLoadFinished = function() {
        // loadInProgress = false;
		// console.log('Load finished');
    // };
	
	function loadedSPEntryPage(status) {		
		// On the SharePoint Entry page
		if (page.url.indexOf(entryId) > 0) {
			console.log('ON THE SP Entry PAGE');
		
			waitFor(
			function() {
				return page.evaluate(function(){
					var elmRows = document.querySelectorAll('div[id$="__sharepointcontent"]');
					return elmRows.length > 0;
					// var elmsHeader = document.querySelectorAll('h1[data-dojo-attach-point="header"]');
					// for (var i = 0; i < elmsHeader.length; i++) {
						// if (elmsHeader[i].textContent == "SharePoint content")
						// {
							// return true;
						// }
					// }
					// return false;
				});
			}, function() {
				// get all the Lib's Id
				var libIds = page.evaluate(function() {
					var elmRows = document.querySelectorAll('div[id$="__sharepointcontent"]');
					var stringIds = '', sId, __idx;
					for (var i = 0; i < elmRows.length; i++) {
						sId = elmRows[i].id;
						__idx = sId.indexOf('-row-');
						stringIds += ';' + sId.substr(__idx + 5);
					}
					if (stringIds.length > 1) {
						stringIds = stringIds.slice(1);
					}
					return stringIds;
				});
				strSPLibIds = libIds;
				console.log('SPLibs Id: ' + libIds);				
				// page.render('page'.concat(entryId,'.png'));
			}, waitForTimeout);
		} else if (page.url.indexOf(strSPIdPostfix) > 0) {
			console.log('LOAD SPLib Items');			
		}		
	};		
		
	function loadedSPLibItems(status) {
		if (page.url.indexOf(strSPIdPostfix) > 0) {
			console.log('LOAD SPLib Items');
			waitFor(
			function() {
				var nowCount = page.evaluate(function(){
					// $('div.dgrid-scroller:visible').height(1500);
					var $elmRows = $('div[id$="__sharepointcontent"]');
					return $elmRows.length;					
				});
				console.log("Now count = " + nowCount);
				return nowCount >= countCheck;
			}, function() {
				if (logAll === true) {
					// get all the Item's Id
					var itemIds = page.evaluate(function() {
						var elmRows = document.querySelectorAll('div[id$="__sharepointcontent"]');
						var stringIds = '', sId, __idx;
						for (var i = 0; i < elmRows.length; i++) {
							sId = elmRows[i].id;
							__idx = sId.indexOf('-row-');
							stringIds += ';' + sId.substr(__idx + 5);
						}
						if (stringIds.length > 1) {
							stringIds = stringIds.slice(1);
						}
						return stringIds;
					});
					console.log('SPItems ID: ' + itemIds);
				}
				page.render('SPLibItems.png');
			}, waitForTimeout * 3);
		}	
	}	

	// close the current then open new page
	function OpenPage(newUrl, loadedFx) {
		page.close();
		page = webpage.create();
		page.open(newUrl, loadedFx);
		console.log('BROWSE PAGE: ' + newUrl);
	};
	
	function printArgs() {
		var i, ilen;
		for (i = 0, ilen = arguments.length; i < ilen; ++i) {
			console.log("    arguments[" + i + "] = " + JSON.stringify(arguments[i]));
		}
		console.log("");
	};
	
	// extract arguments value, returns false if any invalid
	function evaluateArgs() {
		entryPageUrl = system.args[1];
		username = system.args[2];
		password = system.args[3];
		
		var idx = entryPageUrl.indexOf('///');
		if (idx < 12) {
			console.log('Invalid URL.');
			return false;
		}
		
		entryId = entryPageUrl.substr(idx + 3);
		if (entryId.length < 1)
		{
			console.log('Invalid EntryId.');
			return false;
		}
	
		if (system.args.length > 4)
		{	
			callExit = system.args[4] == 'true';
			logAll = system.args.length > 5 && system.args[5] == 'true';
		}
		return true;
	};
	
if (system.args.length < 4) {
    console.log('Usage: loadspitems-test.js <http://[siteURL]/[uiPath]/Cms#context=epi.cms.contentdata:///[EntryId]> <username> <password>');
	console.log('Or: loadspitems-test.js <http://[siteURL]/[uiPath]/Cms#context=epi.cms.contentdata:///[EntryId]> <username> <password> <bool value 1> <bool value 2>');
    phantom.exit(1);
} else {	
	var isvalid = evaluateArgs();
	if (isvalid != true)
	{
		phantom.exit(1);
	}
    t = Date.now();
		
	var steps = [
	function() {
		//Load the Page
		// page.open(entryPageUrl);
		page.open(entryPageUrl, loadedSPEntryPage);
		
		logAll && phantom.cookies.forEach(function(cookie, i) {
			for (var key in cookie) {
				console.log('[cookie:' + i + '] ' + key + ' = ' + cookie[key]);
			}
		});		
	},
	function() {	
		// has been redirected to Login page
		if (page.url.indexOf('login.aspx') > 0) {
			page.evaluate(function (uname, pass) {				
				var form = document.querySelector('#aspnetForm');
				var txtUsername = document.querySelector('#LoginControl_UserName');
				var txtPassword = document.querySelector('#LoginControl_Password');
				txtUsername.value = uname;
				txtPassword.value = pass;
				document.querySelector('#LoginControl_Button1').click();
				console.log("Login info submitted!");
			}, username, password);			
		}
	}, 
	function() {
		// jump to THE ENTRY PAGE after login		
		if (page.url.indexOf(entryId) < 0) {
			OpenPage(entryPageUrl, loadedSPEntryPage);
			// page.close();
			// page = webpage.create();
			// page.open(entryPageUrl, loadedSPEntryPage);
			// console.log('BROWSE THE ENTRY PAGE.');
		}
	},
	function() {
		// jump to a SP Lib page
		if (strSPLibIds.length > 0) {
			var ids = strSPLibIds.split(';');
			var randomIdx = Math.floor(Math.random() * ids.length)
			var newUrl = entryPageUrl.substring(0, entryPageUrl.indexOf('///')) + '///'.concat(ids[randomIdx]);
			OpenPage(newUrl, loadedSPLibItems);
		}
	},
	function() {
	}];
	
	interval = setInterval(function() {
		// if (!loadInProgress && typeof steps[stepNo] == 'function') {
		if (typeof steps[stepNo] == 'function') {
			console.log("");
			logAll && console.log('### step' + stepNo + ' FROM '+ page.url);
			steps[stepNo]();
			// page.render('step' + stepNo + '.png');
			stepNo++;
		}
		// if (typeof steps[stepNo] != 'function') {
		if (callExit && typeof steps[stepNo] != 'function') {
			console.log('');
			console.log('test complete!');
			phantom.exit();
		}
	}, 5000);

	function waitFor(testFx, onReady, timeOutMillis) {
		var maxtimeOutMillis = timeOutMillis ? timeOutMillis : 3000, //< Default Max Timout is 3s
		start = new Date().getTime(), condition = false,
		intervalWait = setInterval(function() {
			if ( (new Date().getTime() - start < maxtimeOutMillis) && !condition ) {
				// If not time-out yet and condition not yet fulfilled
				condition = (typeof(testFx) === "string" ? eval(testFx) : testFx()); //< defensive code
			} else {
				if(!condition) {
					// If condition still not fulfilled (timeout but condition is 'false')
					console.log("'waitFor()' timeout");
					page.render('page'.concat(entryId,'timeout.png'));
					
					// start = new Date().getTime();
					if (callExit === true)
					{
						callExit && phantom.exit(1);
					} else {
						clearInterval(intervalWait);
						waitFor(testFx, onReady, timeOutMillis);
					}
				} else {
					// Condition fulfilled (timeout and/or condition is 'true')
					console.log("'waitFor()' finished in " + (new Date().getTime() - start) + "ms.");
					typeof(onReady) === "string" ? eval(onReady) : onReady(); //< Do what it's supposed to do once the condition is fulfilled
					clearInterval(intervalWait); //< Stop this interval
				}
			}
		}, 500); //< repeat check every 500ms
	};	
}
