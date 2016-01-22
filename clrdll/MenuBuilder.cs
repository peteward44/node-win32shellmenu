using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using System.Threading;
using System.Reflection;
using System.Reflection.Emit;

using System.Runtime.InteropServices;

using System.Security.Cryptography;
using System.Configuration.Assemblies;

using Microsoft.Win32;

namespace windowsexplorermenu_clr
{
	class MenuBuilder
	{
		static bool IsValidInputFile( List<string> fileExtensionFilter, string file )
		{
			if ( fileExtensionFilter.Count > 0 )
			{
				string ext = System.IO.Path.GetExtension( file ).ToLower();
				foreach ( string extc in fileExtensionFilter )
				{
					if ( extc.ToLower() == ext )
						return true;
				}
				return false;
			}
			return true;
		}


		static List<string> BuildFileList( List<string> fileExtensionFilter, List<string> inputFoldersAndFiles )
		{
			List<string> output = new List<string>();
			foreach ( string inputFile in inputFoldersAndFiles )
			{
				if ( System.IO.Directory.Exists( inputFile ) )
				{
					foreach ( string subFile in System.IO.Directory.GetFiles( inputFile ) )
					{
						if ( IsValidInputFile( fileExtensionFilter, subFile ) )
							output.Add( subFile );
					}
				}
				else if ( System.IO.File.Exists( inputFile ) )
				{
					if ( IsValidInputFile( fileExtensionFilter, inputFile ) )
						output.Add( inputFile );
				}
			}
			return output;
		}


		static string JoinWithQuotes( string[] str )
		{
			return @"""" + string.Join( @""" """, str ) + @"""";
		}


		static void AddChildren( Assembly assem, Platform.ShellRightClickContextMenuClass menuClass, InfoStorageAttribute storage, IntPtr menu, dynamic children, List<string> selectedFiles, bool isRoot, MenuClick onClick )
		{
			for ( var i=0; i<children.Length; ++i )
			{
				dynamic item = children[ i ];
				string name = Util.IsJsonProperty( item, "name" ) ? (string)item.name : "New Menu Item";
				
				// embed image into assembly and reference that
				System.Drawing.Bitmap bmp = null;
				IntPtr subMenu = IntPtr.Zero;

				if ( Util.IsJsonProperty( item, "imageResource" ) )
				{
					var resources = assem.GetManifestResourceStream( (string)item.imageResource + ".resources" );
					var rr = new System.Resources.ResourceReader( resources );
					string resourceType;
					byte[] resourceData;
					rr.GetResourceData( "image.bmp", out resourceType, out resourceData );
					// For some reason the resource compiler adds 4 bytes to the start of our data.
					bmp = new System.Drawing.Bitmap( new System.IO.MemoryStream( resourceData, 4, resourceData.Length-4 ) );
				}
				if ( Util.IsJsonProperty( item, "children" ) )
				{
					subMenu = menuClass.CreateSubMenu();
					AddChildren( assem, menuClass, storage, subMenu, item.children, selectedFiles, false, onClick );
				}

				int position = i+1;
				if ( isRoot )
				{
					position = Util.IsJsonProperty( item, "position" ) ? (int)item.position : position;
				}
				if ( menu == IntPtr.Zero )
				{
					// root element
					if ( subMenu == IntPtr.Zero )
					{
						uint id = menuClass.InsertMenuItem( name, position, ( List<string> s ) => { onClick( item, selectedFiles ); } );
						if ( bmp != null )
						{
							menuClass.SetMenuItemBitmap( id, bmp );
						}
					}
					else
					{
						menuClass.InsertSubMenu( subMenu, name, position, bmp );
					}
				}
				else
				{
					// sub menu
					if ( subMenu == IntPtr.Zero )
					{
						menuClass.InsertMenuItemIntoSubMenu( menu, name, position, bmp, ( List<string> s ) => { onClick( item, selectedFiles ); } );
					}
					else
					{
						uint id = menuClass.InsertSubMenuIntoSubMenu( menu, subMenu, name, position );
						if ( bmp != null )
						{
							menuClass.SetMenuItemBitmap( id, bmp );
						}
					}
				}
			}
		}


		public delegate void MenuClick( dynamic item, List<string> s );

		public static void BuildMenu( Assembly assem, Platform.ShellRightClickContextMenuClass menuClass, InfoStorageAttribute storage, List<string> list, MenuClick onClick )
		{
			try
			{
				// filter file extensions if required
				var filter = new List<string>( storage.FileExtensionFilter );
				var filesList = storage.ExpandFileNames ? BuildFileList( filter, list ) : list;

				if ( filesList.Count > 0 )
				{
					string menuFormatJson = storage.MenuFormat;

					var jsonReader = new JsonFx.Json.JsonReader();
					dynamic menuFormatObject = jsonReader.Read<dynamic>( menuFormatJson );
					dynamic menuFormatObjectChildren = menuFormatObject.children;

					AddChildren( assem, menuClass, storage, IntPtr.Zero, menuFormatObjectChildren, filesList, true, onClick );
				}
			}
			catch ( Exception e )
			{
				MessageBox.Show( e.Message + "\r\n\r\n" + e.StackTrace, "Error whilst creating menu" );
			}
		}
	}
}

