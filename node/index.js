'use strict';

var fs = require( 'fs-extra' );
var path = require( 'path' );
var edge = require( 'edge' );
var appRoot = require('app-root-path');
var exec = require( 'child_process' ).exec;

var ourDllPath = path.join( __dirname, 'dll', 'windowsexplorermenu-clr.dll' );
var ourSharpDllPath = path.join( __dirname, 'dll', 'SharpShell.dll' );
var ourJsonFxDllPath = path.join( __dirname, 'dll', 'JsonFx.dll' );


function getResourceNameForImage( actionPath, imageFilename ) {
	var fullPath = path.resolve( actionPath, imageFilename );
	var unformatted = path.relative( actionPath, fullPath );
	return unformatted.replace( /\\|\//g, '.' ); // replace path separators with dots
}


function parseMenuForImagesRecurse( options, children ) {
	// parse menu structure and replace all images with references to a resource name, which will then be embedded into the generated dll
	for ( var i=0; i<children.length; ++i ) {
		var child = children[i];
		if ( child.image ) {
			child.image = child.image.replace( /\\|\//g, path.sep );
			var resourceName = getResourceNameForImage( options.actionpath, child.image );
			if ( !options.resources.hasOwnProperty( resourceName ) ) {
				options.resources[ resourceName ] = path.resolve( options.actionpath, child.image );
			}
			child.imageResource = resourceName;
		}
		if ( child.children ) {
			parseMenuForImagesRecurse( options, child.children );
		}
	}
}


function register( dllname, menu, options, callback ) {
	var clrMethod = edge.func({
		assemblyFile: ourDllPath,
		typeName: 'windowsexplorermenu_clr.ExplorerMenuInterface',
		methodName: 'Register'
	});
	
	var dll = path.normalize( path.resolve( dllname ) );
	// TODO: copy options object before modifying it
	// TODO: sanity check parameters
	options.dllpath = dll;
	if ( !Array.isArray( menu ) ) {
		menu = [ menu ];
	}
	options.actionpath = ( options.actionpath || appRoot.toString() ).toString();
	options.menu = { children: menu };
	options.resources = options.resources || {};
	options.association = options.association || 'all';
	options.associations = options.associations || [];
	
	parseMenuForImagesRecurse( options, options.menu.children );
	
	fs.ensureDirSync( path.dirname( dll ) );
	fs.copySync( ourDllPath, path.join( path.dirname( dll ), path.basename( ourDllPath ) ) );
	fs.copySync( ourSharpDllPath, path.join( path.dirname( dll ), path.basename( ourSharpDllPath ) ) );
	fs.copySync( ourJsonFxDllPath, path.join( path.dirname( dll ), path.basename( ourJsonFxDllPath ) ) );

	clrMethod( options, function( err ) {
		callback( err );
	} );
}

exports.register = register;


function unregister( dllname, callback ) {
	// var dll = path.normalize( path.resolve( dllname ) );
	// return exec( ourSrm + ' uninstall "' + dll + '"', { cwd: path.dirname( dll ) }, callback );
	
	var clrMethod = edge.func({
		assemblyFile: ourDllPath,
		typeName: 'windowsexplorermenu_clr.ExplorerMenuInterface',
		methodName: 'Unregister'
	});
	
	var params = {
		dllpath: path.normalize( path.resolve( dllname ) )
	};
	
	clrMethod( params, function( err ) {
		callback( err );
	} );
}

exports.unregister = unregister;

