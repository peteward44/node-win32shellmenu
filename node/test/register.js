'use strict';

var assert = require( 'assert' );
var fs = require( 'fs' );
var path = require( 'path' );
var explorerMenu = require( '../' );


describe('register', function () {
	
	this.timeout( 5 * 60 * 1000 );
	
	it('create', function ( done ) {
		var dllname = path.join( __dirname, 'mydll.dll' );
		explorerMenu.create( dllname, function( err ) {
			if ( err ) {
				console.error( err );
			}
			assert.equal( !err, true, "No error occurred" );
			assert.equal( fs.existsSync( dllname ), true, "DLL successfully created" );
			
			done();
		} );
	});
	
	it('register', function ( done ) {
		var dllname = path.join( __dirname, 'mydll.dll' );

		explorerMenu.register( dllname, function( err, stdo, stde ) {
			if ( err ) {
				console.error( err );
			}
			assert.equal( !err, true, "No error occurred" );
			done();
		} );
	});
	
	// it('unregister', function ( done ) {
		// var dllname = path.join( __dirname, 'mydll.dll' );

		// explorerMenu.unregister( dllname, function( err, stdo, stde ) {
			// if ( err ) {
				// console.error( err );
			// }
			// assert.equal( !err, true, "No error occurred" );
			// done();
		// } );
	// });
});

