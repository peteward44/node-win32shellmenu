'use strict';

var path = require( 'path' );
var edge = require( 'edge' );


function start() {
	var dllPath = path.join( __dirname, 'dll', 'windowsexplorermenu-clr.dll' );
	var clrMethod = edge.func({
		assemblyFile: dllPath,
	//	typeName: 'windowsexplorermenu_clr.Class1',
	//	methodName: 'TestMethod' // This must be Func<object,Task<object>> 
		typeName: 'windowsexplorermenu_clr.TestClass',
		methodName: 'Create'
	});
	
	clrMethod( function() {
		console.log( "method executed" );
	} );
}


start();
