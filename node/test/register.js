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
				name: "Dynamic menu test - item 1",
				position: 1
			},
			{
				name: "Dynamic menu test - item 2",
				action: "js/myjs.js",
				args: [ "arg1" ],
				position: 3
			},
			{
				name: "Dynamic menu test - item 3",
				position: 2,
				children: [
					{
						name: "Sub menu item 1",
						image: "test/icon.bmp"
					},
					{
						name: "Sub menu item 2"
					}
				]
			},
			{
				name: "Dynamic menu test - item 3 - cmd",
				cmd: "echo It worked!"
			}
		];
		var options = {
			association: [ 'all' /*'files', 'folders'*/ ],
			fileExtensionFilter: '.md',
			guid: "8e2b5ec9-c073-4750-a811-b218ca58c3ae"
		};
		explorerMenu.register( dllname, menu, options, function( err ) {
			if ( err ) {
				console.error( err );
			}
			assert.equal( !err, true, "No error occurred" );
			done();
		} );
	});
});

