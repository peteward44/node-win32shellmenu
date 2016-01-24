'use strict';

var assert = require( 'assert' );
var fs = require( 'fs' );
var path = require( 'path' );
var explorerMenu = require( '../' );


describe('unregister', function () {
	
	this.timeout( 5 * 60 * 1000 );
	
	it('unregister', function ( done ) {
		var dllname = path.join( __dirname, 'mydll.dll' );

		explorerMenu.unregister( dllname, {}, function( err ) {
			if ( err ) {
				console.error( err );
			}
			assert.equal( !err, true, "No error occurred" );
			done();
		} );
	});
});

