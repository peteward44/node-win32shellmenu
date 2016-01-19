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
	public class Action
	{
		static string GetCurrentExecutingDirectory()
		{
			string filePath = new Uri( Assembly.GetExecutingAssembly().CodeBase ).LocalPath;
			return System.IO.Path.GetDirectoryName( filePath );
		}

		static bool IsValidInputFile( string file )
		{
			string ext = System.IO.Path.GetExtension( file ).ToLower();
			bool validExtension = ext == ".wav" || ext == ".ogg" || ext == ".mp3";
			return validExtension;
		}

		static List<string> BuildFileList( List<string> inputFoldersAndFiles )
		{
			List<string> output = new List<string>();
			foreach ( string inputFile in inputFoldersAndFiles )
			{
				if ( System.IO.Directory.Exists( inputFile ) )
				{
					foreach ( string subFile in System.IO.Directory.GetFiles( inputFile ) )
					{
						if ( IsValidInputFile( subFile ) )
							output.Add( subFile );
					}
				}
				else if ( System.IO.File.Exists( inputFile ) )
				{
					if ( IsValidInputFile( inputFile ) )
						output.Add( inputFile );
				}
			}
			return output;
		}

		public static void DoAction( string action, string[] args, List<string> inputFoldersAndFiles )
		{
			//try
			//{
			//	List<string> files = BuildFileList( inputFoldersAndFiles );
			//	if ( files.Count == 0 )
			//	{
			//		MessageBox.Show( "No valid image files found in selection", "Inspired Texture Tool", MessageBoxButtons.OK, MessageBoxIcon.Warning );
			//		return;
			//	}

			//	string toolPath = System.IO.Path.Combine( GetCurrentExecutingDirectory(), "js", "index.js" );
			//	string filesList = "\"" + string.Join( "\" \"", files ) + "\"";

			//	string arguments = "/S /K call node \"" + toolPath + "\" " + ( convertOnly ? "--convertOnly " : "" );
			//	arguments += filesList;

			//	System.Diagnostics.ProcessStartInfo psi = new System.Diagnostics.ProcessStartInfo();
			//	psi.CreateNoWindow = false;
			//	psi.WindowStyle = System.Diagnostics.ProcessWindowStyle.Normal;
			//	psi.Arguments = arguments;
			//	psi.FileName = "cmd.exe";
			//	psi.UseShellExecute = false;
			//	System.Diagnostics.Process.Start( psi );
			//}
			//catch ( Exception e )
			//{
			//	string err = e.Message + "\r\n\r\n" + e.StackTrace;
			//	Console.WriteLine( err );
			//	MessageBox.Show( err );
			//}
		}
	}


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


		private void OnClick( dynamic item )
		{
			var action = Util.IsJsonProperty( item, "action" ) ? (string)item.action : "";
			var args = Util.IsJsonProperty( item, "args" ) ? (string[])item.args : new string[]{};
			var style = Util.IsJsonProperty( item, "style" ) ? (string)item.remain : "remain";

			string type = style == "remain" ? "/K" : "/C";
			string argsString = "";
			if ( args.Length > 0 )
			{
				argsString = @"""" + string.Join( @""" """, args ) + @"""";
			}
			string fullArgs = type + @" node """ + action + @""" " + argsString;
		
			System.Diagnostics.Process process = new System.Diagnostics.Process();
			System.Diagnostics.ProcessStartInfo startInfo = new System.Diagnostics.ProcessStartInfo();
			startInfo.WindowStyle = style == "hidden" ? System.Diagnostics.ProcessWindowStyle.Hidden : System.Diagnostics.ProcessWindowStyle.Normal;
			startInfo.FileName = "cmd.exe";
			startInfo.Arguments = fullArgs;
			startInfo.WorkingDirectory = storage.ActionPath;
			process.StartInfo = startInfo;
			process.Start();
		}

		//IntPtr subMenu = CreateSubMenu();
		//InsertMenuItemIntoSubMenu( subMenu, @"Create audio sprite", 1, delegate( List<string> selectedFiles ) { Action.DoAction( selectedFiles, false ); } );
		//InsertMenuItemIntoSubMenu( subMenu, @"Convert file(s) to mp3/ogg", 2, delegate( List<string> selectedFiles ) { Action.DoAction( selectedFiles, true ); } );
		//InsertSeperator( 3 );
		//InsertSubMenu( subMenu, @"Inspired Audio Tool", 4, LoadBitmap() );
		//InsertSeperator( 5 );

		private void AddChildren( IntPtr menu, dynamic children )
		{
			for ( var i=0; i<children.Length; ++i )
			{
				dynamic item = children[ i ];
				JsonFx.Json.JsonWriter jw = new JsonFx.Json.JsonWriter();
				string s = jw.Write( item );
				string name = Util.IsJsonProperty( item, "name" ) ? (string)item.name : "New Menu Item";
				string action = Util.IsJsonProperty( item, "action" ) ? (string)item.action : "";
				string[] args = Util.IsJsonProperty( item, "args" ) ? Util.ObjectToStringArray( (object[])item.args ) : new string[0];

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
					AddChildren( subMenu, item.children );
				}

				int position = i+1;
				if ( menu == IntPtr.Zero )
				{
					// root element
					if ( subMenu == IntPtr.Zero )
					{
						uint id = InsertMenuItem( name, position, ( List<string> selectedFiles ) => { Action.DoAction( action, args, selectedFiles ); } );
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
						InsertMenuItemIntoSubMenu( menu, name, position, bmp, ( List<string> selectedFiles ) => { Action.DoAction( action, args, selectedFiles ); } );
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
				string menuFormatJson = storage.MenuFormat;

				var jsonReader = new JsonFx.Json.JsonReader();
				dynamic menuFormatObject = jsonReader.Read<dynamic>( menuFormatJson );
				dynamic menuFormatObjectChildren = menuFormatObject.children;

				AddChildren( IntPtr.Zero, menuFormatObjectChildren );
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
		private string name;
		private string actionPath;
		private string menuFormat;
		private Platform.ComRegisterClass.RightClickContextMenuOptions[] at;
		public Platform.ComRegisterClass.RightClickContextMenuOptions[] Association
		{
			get
			{
				return at;
			}
		}
		public string MenuFormat
		{
			get
			{
				return menuFormat;
			}
		}
		public string ActionPath
		{
			get
			{
				return actionPath;
			}
		}
		public string Name
		{
			get { return name; }
		}

		public AssociationStorageAttribute( string name, string actionPath, string menuFormat, Platform.ComRegisterClass.RightClickContextMenuOptions[] _at )
		{
			this.name = name;
			this.actionPath = actionPath;
			this.menuFormat = menuFormat;
			this.at = _at;
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


		public void Create( string name, Guid guid, string dllPath, IDictionary<string, object> resourceList, string actionPath, dynamic menuFormat, Platform.ComRegisterClass.RightClickContextMenuOptions[] association )
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
			AddAttribute( myAsmBuilder, typeof( AssociationStorageAttribute ), name, actionPath, menuFormatJson, association );

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

			Guid guid = System.Guid.NewGuid();
			Create( name, guid, dllPath, resourceList, actionPath, menuFormat, association );

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

