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
	class Util
	{
		public static bool IsJsonProperty( dynamic expandoObject, string name )
		{
			var dic = (IDictionary<String, Object>)expandoObject;
			return dic.ContainsKey( name );
		}

		public static bool IsProperty( dynamic expandoObject, string name )
		{
			return expandoObject.GetType().GetProperty( name ) != null;
		}


		public static string[] ObjectToStringArray( object[] array )
		{
			List<string> list = new List<string>();
			foreach ( object o in array )
			{
				list.Add( (string)o );
			}
			return list.ToArray();
		}
	}

	//[Guid( "9AA8DDCB-0540-4cd3-BD31-D91DADD81EE5" ), ComVisible( true )]
	public class MenuExtension : Platform.ShellRightClickContextMenuClass
	{
		AssociationStorageAttribute storage;


		bool IsValidInputFile( List<string> fileExtensionFilter, string file )
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

		List<string> BuildFileList( List<string> fileExtensionFilter, List<string> inputFoldersAndFiles )
		{
			if ( !storage.ExpandFileNames )
				return inputFoldersAndFiles;

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


		string JoinWithQuotes( string[] str )
		{
			return @"""" + string.Join( @""" """, str ) + @"""";
		}


		private void OnClick( dynamic item, List<string> filesList )
		{
			string action = Util.IsJsonProperty( item, "action" ) ? (string)item.action : "";

			if ( action.Length > 0 )
			{
				if ( filesList.Count > 0 )
				{
					string[] args = Util.IsJsonProperty( item, "args" ) ? Util.ObjectToStringArray( (object[])item.args ) : new string[ 0 ];
					var style = Util.IsJsonProperty( item, "style" ) ? (string)item.remain : "remain";

					string type = style == "remain" ? "/K" : "/C";
					string argsString = args.Length > 0 ? JoinWithQuotes( args ) : "";
					string filesString = filesList.Count > 0 ? JoinWithQuotes( filesList.ToArray() ) : "";
					string fullArgs = type + @" node """ + action + @""" " + argsString + " " + filesString;

					System.Diagnostics.Process process = new System.Diagnostics.Process();
					System.Diagnostics.ProcessStartInfo startInfo = new System.Diagnostics.ProcessStartInfo();
					startInfo.WindowStyle = style == "hidden" ? System.Diagnostics.ProcessWindowStyle.Hidden : System.Diagnostics.ProcessWindowStyle.Normal;
					startInfo.FileName = "cmd.exe";
					startInfo.Arguments = fullArgs;
					startInfo.WorkingDirectory = storage.ActionPath;
					process.StartInfo = startInfo;
					process.Start();
				}
			}
		}

		//IntPtr subMenu = CreateSubMenu();
		//InsertMenuItemIntoSubMenu( subMenu, @"Create audio sprite", 1, delegate( List<string> selectedFiles ) { Action.DoAction( selectedFiles, false ); } );
		//InsertMenuItemIntoSubMenu( subMenu, @"Convert file(s) to mp3/ogg", 2, delegate( List<string> selectedFiles ) { Action.DoAction( selectedFiles, true ); } );
		//InsertSeperator( 3 );
		//InsertSubMenu( subMenu, @"Inspired Audio Tool", 4, LoadBitmap() );
		//InsertSeperator( 5 );

		private void AddChildren( IntPtr menu, dynamic children, List<string> selectedFiles, bool isRoot )
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
					var resources = this.GetType().Assembly.GetManifestResourceStream( (string)item.imageResource + ".resources" );
					var rr = new System.Resources.ResourceReader( resources );
					string resourceType;
					byte[] resourceData;
					rr.GetResourceData( "image.bmp", out resourceType, out resourceData );
					// For some reason the resource compiler adds 4 bytes to the start of our data.
					bmp = new System.Drawing.Bitmap( new System.IO.MemoryStream( resourceData, 4, resourceData.Length-4 ) );
				}
				if ( Util.IsJsonProperty( item, "children" ) )
				{
					subMenu = CreateSubMenu();
					AddChildren( subMenu, item.children, selectedFiles, false );
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
						uint id = InsertMenuItem( name, position, ( List<string> s ) => { OnClick( item, selectedFiles ); } );
						if ( bmp != null )
						{
							SetMenuItemBitmap( id, bmp );
						}
					}
					else
					{
						InsertSubMenu( subMenu, name, position, bmp );
					}
				}
				else
				{
					// sub menu
					if ( subMenu == IntPtr.Zero )
					{
						InsertMenuItemIntoSubMenu( menu, name, position, bmp, ( List<string> s ) => { OnClick( item, selectedFiles ); } );
					}
					else
					{
						uint id = InsertSubMenuIntoSubMenu( menu, subMenu, name, position );
						if ( bmp != null )
						{
							SetMenuItemBitmap( id, bmp );
						}
					}
				}
			}
		}


		protected override string GetVerbString()
		{
			return "Verb String";
		}

		protected override string GetHelpString()
		{
			return "Help String";
		}

		protected override void OnBuildMenu( List<string> list )
		{
			try
			{
				Assembly assem = this.GetType().Assembly;
				this.storage = (AssociationStorageAttribute)assem.GetCustomAttributes( typeof( AssociationStorageAttribute ), true )[ 0 ];

				// filter file extensions if required
				var filter = new List<string>( storage.FileExtensionFilter );
				var filesList = BuildFileList( filter, list );

				if ( filesList.Count > 0 )
				{
					string menuFormatJson = storage.MenuFormat;

					var jsonReader = new JsonFx.Json.JsonReader();
					dynamic menuFormatObject = jsonReader.Read<dynamic>( menuFormatJson );
					dynamic menuFormatObjectChildren = menuFormatObject.children;

					AddChildren( IntPtr.Zero, menuFormatObjectChildren, filesList, true );
				}
			}
			catch ( Exception e )
			{
				MessageBox.Show( e.Message + "\r\n\r\n" + e.StackTrace, "Error whilst creating menu" );
			}
		}
	}


	// Custom attribute to store the file association data at the assembly level.
	public class AssociationStorageAttribute : Attribute
	{
		public Platform.ComRegisterClass.RightClickContextMenuOptions[] Association
		{
			private set;
			get;
		}
		public string MenuFormat
		{
			private set;
			get;
		}
		public string ActionPath
		{
			private set;
			get;
		}
		public string Name
		{
			private set;
			get;
		}
		public string[] FileExtensionFilter
		{
			private set;
			get;
		}
		public bool ExpandFileNames
		{
			private set;
			get;
		}
		public AssociationStorageAttribute( string name, string actionPath, string menuFormat, Platform.ComRegisterClass.RightClickContextMenuOptions[] _at, string[] fileExtensionFilter, bool expandFileNames )
		{
			this.Name = name;
			this.ActionPath = actionPath;
			this.MenuFormat = menuFormat;
			this.Association = _at;
			this.FileExtensionFilter = fileExtensionFilter;
			this.ExpandFileNames = expandFileNames;
		}

	}


	public class ExplorerMenuInterface
	{

		public static byte[] CreateKeyPair( string containerName, int keySize )
		{
			if ( ( keySize % 8 ) != 0 )
			{
				throw new CryptographicException( "Invalid key size. Valid size is 384 to 16384 mod 8.  Default 1024." );
			}

			CspParameters parms = new CspParameters();
			parms.KeyContainerName = containerName;
			parms.KeyNumber = 2;
			RSACryptoServiceProvider provider = new RSACryptoServiceProvider( keySize, parms );
			byte[] array = provider.ExportCspBlob( !provider.PublicOnly );
			return array;
		}


		private static void AddAttribute( dynamic targetType, Type attributeType, params object[] constructorParams )
		{
			Type[] ctorParams = new Type[ constructorParams.Length ];
			int index = 0;
			foreach ( object o in constructorParams )
			{
				ctorParams[ index++ ] = o.GetType();
			}
			ConstructorInfo classCtorInfo = attributeType.GetConstructor( ctorParams );
			CustomAttributeBuilder myCABuilder = new CustomAttributeBuilder( classCtorInfo, constructorParams );
			targetType.SetCustomAttribute( myCABuilder );
		}


		private Platform.ComRegisterClass.RightClickContextMenuOptions[] StringNameToAssociationType( string[] names )
		{
			List<Platform.ComRegisterClass.RightClickContextMenuOptions> list = new List<Platform.ComRegisterClass.RightClickContextMenuOptions>();

			foreach ( object n in names )
			{
				string name = (string)n;
				switch ( name )
				{
					case "all":
						list.Add( Platform.ComRegisterClass.RightClickContextMenuOptions.AllFileSystemObjects );
						break;
					case "files":
						list.Add( Platform.ComRegisterClass.RightClickContextMenuOptions.Files );
						break;
					case "folders":
						list.Add( Platform.ComRegisterClass.RightClickContextMenuOptions.Folders );
						break;
					case "imagefiles":
						list.Add( Platform.ComRegisterClass.RightClickContextMenuOptions.ImageFiles );
						break;
					case "videofiles":
						list.Add( Platform.ComRegisterClass.RightClickContextMenuOptions.VideoFiles );
						break;
					case "desktopbackground":
						list.Add( Platform.ComRegisterClass.RightClickContextMenuOptions.DesktopBackground );
						break;
					case "drive":
						list.Add( Platform.ComRegisterClass.RightClickContextMenuOptions.Drive );
						break;
					case "printers":
						list.Add( Platform.ComRegisterClass.RightClickContextMenuOptions.Printers );
						break;
				}
			}
			if ( list.Count == 0 )
			{
				list.Add( Platform.ComRegisterClass.RightClickContextMenuOptions.AllFileSystemObjects );
			}
			return list.ToArray();
		}


		public void Create( string name, Guid guid, string dllPath, IDictionary<string, object> resourceList, string actionPath,
			dynamic menuFormat, Platform.ComRegisterClass.RightClickContextMenuOptions[] association, string[] fileExtensionFilter, bool expandFileNames )
		{
			AssemblyName myAsmName = new AssemblyName( System.IO.Path.GetFileNameWithoutExtension( dllPath ) );
			myAsmName.CodeBase = String.Concat( "file:///", System.IO.Path.GetDirectoryName( dllPath ) );
			myAsmName.CultureInfo = new System.Globalization.CultureInfo( "en-US" );
			myAsmName.KeyPair = new StrongNameKeyPair( CreateKeyPair( System.IO.Path.GetFileNameWithoutExtension( dllPath ), 1024 ) );
			myAsmName.Flags = AssemblyNameFlags.PublicKey;
			myAsmName.VersionCompatibility = AssemblyVersionCompatibility.SameProcess;
			myAsmName.HashAlgorithm = AssemblyHashAlgorithm.SHA1;
			myAsmName.Version = new Version( "1.0.0.0" );

			AssemblyBuilder myAsmBuilder = AppDomain.CurrentDomain.DefineDynamicAssembly( myAsmName, AssemblyBuilderAccess.Save, System.IO.Path.GetDirectoryName( dllPath ) );
			ModuleBuilder myModBuilder = myAsmBuilder.DefineDynamicModule( "MyModule", System.IO.Path.GetFileName( dllPath ) );
			TypeBuilder myTypeBuilder = myModBuilder.DefineType( "MyType", TypeAttributes.Public, typeof( MenuExtension ) );

			var menuFormatWriter = new JsonFx.Json.JsonWriter();
			var menuFormatJson = menuFormatWriter.Write( menuFormat );

			AddAttribute( myTypeBuilder, typeof( ComVisibleAttribute ), true );
			AddAttribute( myTypeBuilder, typeof( GuidAttribute ), guid.ToString() );
			AddAttribute( myAsmBuilder, typeof( GuidAttribute ), guid.ToString() );
			AddAttribute( myAsmBuilder, typeof( AssociationStorageAttribute ), name, actionPath, menuFormatJson, association, fileExtensionFilter, expandFileNames );

			// embed all images found in the menu into the assembly
			foreach ( string keyname in resourceList.Keys )
			{
				string filename = (string)resourceList[ keyname ];
				System.Drawing.Image image = System.Drawing.Image.FromFile( filename );
				System.IO.MemoryStream memStream = new System.IO.MemoryStream();
				image.Save( memStream, System.Drawing.Imaging.ImageFormat.Bmp );
				byte[] rawdata = memStream.ToArray();
				System.Resources.IResourceWriter rw = myModBuilder.DefineResource( keyname + ".resources", "description", ResourceAttributes.Public );
				rw.AddResource( "image.bmp", rawdata );
			}

			myTypeBuilder.CreateType();
			myModBuilder.CreateGlobalFunctions();

			myAsmBuilder.Save( System.IO.Path.GetFileName( dllPath ) );
		}


		Assembly OnResolve( object sender, ResolveEventArgs args )
		{
			var requestedAssembly = new AssemblyName( args.Name );
			var name = requestedAssembly.Name;
			// look in the same directory as this executing dll
			var p = System.IO.Path.Combine( System.IO.Path.GetDirectoryName( System.Reflection.Assembly.GetExecutingAssembly().Location ), name + ".dll" );
			if ( System.IO.File.Exists( p ) ) {
				return Assembly.LoadFile( p );
			}
			return null;
		}


		public async Task<object> Register( dynamic input )
		{
			string name = (string)input.name;
			string dllPath = (string)input.dllpath;
			IDictionary<string, object> resourceList = (IDictionary<string, object>)input.resources;
			string actionPath = (string)input.actionpath;
			dynamic menuFormat = (dynamic)input.menu;
			Platform.ComRegisterClass.RightClickContextMenuOptions[] association = StringNameToAssociationType( Util.ObjectToStringArray( (object[])input.association ) );
			string[] fileExtensionFilter = Util.ObjectToStringArray( (object[])input.fileExtensionFilter );
			string guidString = (string)input.guid;
			bool expandFileNames = (bool)input.expandFileNames;

			Guid guid = guidString.Length > 0 ? new System.Guid( guidString ) : System.Guid.NewGuid();
			Create( name, guid, dllPath, resourceList, actionPath, menuFormat, association, fileExtensionFilter, expandFileNames );

			AppDomain.CurrentDomain.AssemblyResolve += ( sender, args ) => OnResolve( sender, args );
			Platform.ComRegisterClass.RegisterShellRightClickContextMenu( name, "{" + guid.ToString() + "}", association );
			Platform.ComRegisterClass.RegisterServer( Assembly.LoadFile( dllPath ) );

			return "";
		}


		public async Task<object> Unregister( dynamic input )
		{
			string dllPath = (string)input.dllpath;
			AppDomain.CurrentDomain.AssemblyResolve += ( sender, args ) => OnResolve( sender, args );
			Assembly assembly = Assembly.LoadFile( dllPath );

			var associationAttrib = (AssociationStorageAttribute)assembly.GetCustomAttributes( typeof( AssociationStorageAttribute ), true )[ 0 ];
			Platform.ComRegisterClass.RightClickContextMenuOptions[] associationArray = associationAttrib.Association;
			var guidAttrib = (GuidAttribute)assembly.GetCustomAttributes(typeof(GuidAttribute),true)[0];

			Platform.ComRegisterClass.RegisterShellRightClickContextMenu( associationAttrib.Name, "{" + guidAttrib.Value + "}", associationArray );
			Platform.ComRegisterClass.UnregisterServer( assembly );

			return "";
		}
	}
}

