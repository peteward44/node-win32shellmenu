# win32shellmenu
Allows a node.js application to register a windows explorer context menu shell extension.

Intended to be used as part of the npm install / uninstall script.

Basic usage
```
var shellmenu = require( 'win32shellmenu' );

// An MSIL DLL is created as part of the registration process. The location to create this must be specified
// The same dll path will be required to unregister the shell extension
var dllname = 'mydll.dll';
var menu = [
	{
		// simple item with an icon image
		name: "Dynamic menu test - item 1",
		// "image" can specify a valid bmp / png to use as an icon image
		image: "icon.bmp"
	},
	{
		// specify an action to occur when clicked - will execute myjs.js with the argument arg1
		name: "Dynamic menu test - item 2",
		action: "myjs.js",
		args: [ "arg1" ]
	},
	{
		name: "Dynamic menu test - item 3",
		// "children" property can mean to specify a submenu
		children: [
			{
				name: "Sub menu item 1",
				image: "icon.bmp"
			},
			{
				name: "Sub menu item 2"
			}
		]
	}
];

var options = {
	actionpath: __dirname, // Root folder which the script specified by "action" will be executed from - most applications will want this set to __dirname if they are being installed globally
	association: [ 'files', 'folders' ], // types of shell item to bind to - this can be either 'all', 'files', 'folders', 'imagefiles', 'videofiles', 'desktopbackground', 'drive' or 'printers'
	fileExtensionFilter: [ '.md' ] // filter by only these file extensions
};

shellmenu.register( dllname, menu, options, function( err ) {
	if ( err ) {
		console.error( err );
	}
});

```

Unregistering
```

var shellmenu = require( 'win32shellmenu' );
var dllname = 'mydll.dll';
// For "restartExplorer" to work correctly, process.exit() must be called, terminating the process.
// Set to false if you want to restart explorer yourself
// Callback will only be executed if restartExplorer is set to false
shellmenu.unregister( dllname, { restartExplorer: true }, function( err ) {} );

```


