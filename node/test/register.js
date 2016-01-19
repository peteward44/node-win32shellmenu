'use strict';

var assert = require( 'assert' );
var fs = require( 'fs' );
var path = require( 'path' );
var explorerMenu = require( '../' );


describe('register', function () {
	
	this.timeout( 5 * 60 * 1000 );
	
	it('register', function ( done ) {
		var dllname = path.join( __dirname, 'mydll.dll' );
		var menu = [
			{
				name: "Dynamic menu test - item 1"
			},
			{
				name: "Dynamic menu test - item 2",
				action: "js/myjs.js",
				args: [ "arg1" ]
			},
			{
				name: "Dynamic menu test - item 3",
				children: [
					{
						name: "Sub menu item 1",
						image: "test/icon.bmp"
					},
					{
						name: "Sub menu item 2"
					}
				]
			}
		];
		var options = { association: [ 'all' /*'files', 'folders'*/ ] };
		explorerMenu.register( dllname, menu, options, function( err ) {
			if ( err ) {
				console.error( err );
			}
			assert.equal( !err, true, "No error occurred" );
			done();
		} );
	});
});

