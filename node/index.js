'use strict';

var fs = require( 'fs-extra' );
var path = require( 'path' );
var edge = require( 'edge' );
var exec = require( 'child_process' ).exec;

var ourDllPath = path.join( __dirname, 'dll', 'windowsexplorermenu-clr.dll' );
var ourSharpDllPath = path.join( __dirname, 'dll', 'SharpShell.dll' );
var ourSrm = "srm.exe"; // path.join( __dirname, 'dll', 'srm.exe' );


function create( dllname, callback ) {
	var clrMethod = edge.func({
		assemblyFile: ourDllPath,
		typeName: 'windowsexplorermenu_clr.CreateComAssembly',
		methodName: 'Create'
	});
	
	var dll = path.normalize( path.resolve( dllname ) );
	var params = {
		dllpath: dll
	};
	
	clrMethod( params, function( err ) {
		fs.copySync( ourDllPath, path.join( path.dirname( dll ), path.basename( ourDllPath ) ) );
		fs.copySync( ourSharpDllPath, path.join( path.dirname( dll ), path.basename( ourSharpDllPath ) ) );
		callback( err );
	} );
}


exports.create = create;


function register( dllname, callback ) {
	var dll = path.normalize( path.resolve( dllname ) );
	return exec( ourSrm + ' install ' + path.basename( dll ) + ' -codebase', { cwd: path.dirname( dll ) }, callback );
	// var clrMethod = edge.func({
		// assemblyFile: ourDllPath,
		// typeName: 'windowsexplorermenu_clr.CreateComAssembly',
		// methodName: 'Register'
	// });
	
	// var params = {
		// dllpath: path.normalize( path.resolve( ourDllPath ) )
	// };
	
	// clrMethod( params, function( err ) {
		// callback( err );
	// } );
}

exports.register = register;


function unregister( dllname, callback ) {
	var dll = path.normalize( path.resolve( dllname ) );
	return exec( ourSrm + ' uninstall "' + dll + '"', { cwd: path.dirname( dll ) }, callback );
	
	// var clrMethod = edge.func({
		// assemblyFile: ourDllPath,
		// typeName: 'windowsexplorermenu_clr.CreateComAssembly',
		// methodName: 'Unregister'
	// });
	
	// var params = {
		// dllpath: path.normalize( path.resolve( dllname ) )
	// };
	
	// clrMethod( params, function( err ) {
		// callback( err );
	// } );
}

exports.unregister = unregister;

