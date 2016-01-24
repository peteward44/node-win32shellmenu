'use strict';

var fs = require( 'fs-extra' );
var path = require( 'path' );
var edge = require( 'edge' );
var appRoot = require('app-root-path');
var wincmd = require('node-windows');
var spawn = require( 'child_process' ).spawn;

var ourDllPath = path.join( __dirname, 'dll', 'windowsexplorermenu-clr.dll' );
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
		if ( child.action ) {
			child.action = path.resolve( options.actionpath, child.action );
		}
		if ( child.children ) {
			parseMenuForImagesRecurse( options, child.children );
		}
	}
}


function check( callback ) {
	wincmd.isAdminUser(function(isAdmin){
		if ( !isAdmin ) {
			console.error( "This command must be ran with admin priviledges" );
			return callback( false );
		}
		
		return callback( true );
	});
}


function register( dllname, menu, options, callback ) {

	check( function( valid ) {
		if ( !valid ) {
			return callback();
		}
		
		var clrMethod = edge.func({
			assemblyFile: ourDllPath,
			typeName: 'windowsexplorermenu_clr.ExplorerMenuInterface',
			methodName: 'Register'
		});
		
		var dll = path.normalize( path.resolve( __dirname, dllname ) );
		// TODO: copy options object before modifying it
		// TODO: sanity check parameters
		options.dllpath = dll;
		if ( !Array.isArray( menu ) ) {
			menu = [ menu ];
		}
		options.name = options.name || ( path.basename( dllname, path.extname( dllname ) ) );
		options.actionpath = ( options.actionpath || appRoot.toString() ).toString();
		options.menu = { children: menu };
		options.resources = options.resources || {};
		options.association = options.association || [ 'all' ];
		options.fileExtensionFilter = options.fileExtensionFilter || [];
		options.guid = options.guid || '';
		options.expandFileNames = options.expandFileNames === undefined ? true : options.expandFileNames;
		
		if ( !Array.isArray( options.association ) ) {
			options.association = [ options.association ];
		}
		if ( !Array.isArray( options.associations ) ) {
			options.associations = [ options.associations ];
		}
		if ( !Array.isArray( options.fileExtensionFilter ) ) {
			options.fileExtensionFilter = [ options.fileExtensionFilter ];
		}
		parseMenuForImagesRecurse( options, options.menu.children );
		
		fs.ensureDirSync( path.dirname( dll ) );
		fs.copySync( ourDllPath, path.join( path.dirname( dll ), path.basename( ourDllPath ) ) );
		fs.copySync( ourJsonFxDllPath, path.join( path.dirname( dll ), path.basename( ourJsonFxDllPath ) ) );

		clrMethod( options, function( err ) {
			callback( err );
		} );
	} );
}

exports.register = register;


function unregister( dllname, options, callback ) {

	check( function( valid ) {
		if ( !valid ) {
			return callback();
		}

		var clrMethod = edge.func({
			assemblyFile: ourDllPath,
			typeName: 'windowsexplorermenu_clr.ExplorerMenuInterface',
			methodName: 'Unregister'
		});
		
		var params = {
			dllpath: path.normalize( path.resolve( dllname ) )
		};
		
		options = options || {};
		options.restartExplorer = options.restartExplorer === undefined ? true : options.restartExplorer;
		
		clrMethod( params, function( err ) {
			if ( options.restartExplorer ) {
				var proc = spawn( "cmd.exe", [ "/C", path.join( __dirname, "restart_explorer.cmd" ) ], { detached: true, stdio: 'ignore' } );
				//proc.stdout.pipe( process.stdout );
				//proc.stderr.pipe( process.stderr );
				proc.on( 'exit', function( err2 ) {
					// Can't find a way of restarting explorer without requiring to call process.exit.
					setTimeout( function() { process.exit(); }, 10000 );
				} );
			} else {
				if ( callback ) {
					callback( err );
				}
			}
		} );
	} );
}

exports.unregister = unregister;



