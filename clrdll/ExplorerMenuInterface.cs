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
using SharpShell.Attributes;
using SharpShell.SharpContextMenu;

using System.Security.Cryptography;
using System.Configuration.Assemblies;

using Microsoft.Win32;

namespace windowsexplorermenu_clr
{

	class Util
	{
		public static bool IsProperty( dynamic settings, string name )
		{
			return settings.GetType().GetProperty( name ) != null;
		}
	}

	//[ComVisible( true )]
	//[COMServerAssociation( AssociationType.FileExtension, ".txt" )]
	public class MenuExtension : SharpContextMenu
	{
		protected override bool CanShowMenu()
		{
			return true;
		}

		protected override ContextMenuStrip CreateMenu()
		{
			//(GuidAttribute)assembly.GetCustomAttributes( typeof( GuidAttribute ), true )[ 0 ];
			AssociationStorageAttribute storage = (AssociationStorageAttribute)Attribute.GetCustomAttribute( this.GetType(), typeof( AssociationStorageAttribute ) );
			string menuFormatJson = storage.MenuFormat;

			MessageBox.Show( "menuFormat=" + menuFormatJson );
			var jsonReader = new JsonFx.Json.JsonReader();
			dynamic[] menuFormatObject = jsonReader.Read<dynamic[]>( menuFormatJson );

			MessageBox.Show( "menuFormatObject=" + menuFormatObject.Length );

			//  Create the menu strip.
			var menu = new ContextMenuStrip();

	//		//  Create a 'count lines' item.
	//		var itemCountLines = new ToolStripMenuItem
	//		{
	//			Text = "Count Lines..."
	////			Image = Properties.Resources.CountLines
	//		};

	//		//  When we click, we'll count the lines.
	//		itemCountLines.Click += ( sender, args ) => OnClick();

	//		//  Add the item to the context menu.
	//		menu.Items.Add( itemCountLines );

			//  Return the menu.
			return menu;
		}

		private void OnClick()
		{
			MessageBox.Show( "Test worked" );
		}
	}


	// Custom attribute to store the file association data at the assembly level.
	public class AssociationStorageAttribute : Attribute
	{
		private string menuFormat;
		private AssociationType at;
		public AssociationType Association
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

		public AssociationStorageAttribute( string menuFormat, AssociationType _at )
		{
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


		private AssociationType StringNameToAssociationType( string name )
		{
			switch ( name )
			{
				default:
				case "all":
					return AssociationType.AllFiles;
				case "class":
					return AssociationType.Class;
				case "classofextension":
					return AssociationType.ClassOfExtension;
				case "directory":
					return AssociationType.Directory;
				case "drive":
					return AssociationType.Drive;
				case "fileextension":
					return AssociationType.FileExtension;
				case "none":
					return AssociationType.None;
				case "unknown":
					return AssociationType.UnknownFiles;
			}
		}


		public void Create( string dllPath, dynamic[] menuFormat, AssociationType association, string[] associations )
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

			Guid guid = System.Guid.NewGuid();
			AddAttribute( myTypeBuilder, typeof( ComVisibleAttribute ), true );
			AddAttribute( myTypeBuilder, typeof( COMServerAssociationAttribute ), association, associations );
			AddAttribute( myAsmBuilder, typeof( GuidAttribute ), guid.ToString() );
			AddAttribute( myAsmBuilder, typeof( AssociationStorageAttribute ), menuFormatJson, association );

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


		static string GetContextMenuOptionsBaseKey( AssociationType association )
		{
			// See what this is about: @"HKEY_CLASSES_ROOT\AllFilesystemObjects\shellex\ContextMenuHandlers" 
			// Compressed folders: HKEY_CLASSES_ROOT\CompressedFolder\ShellEx\ContextMenuHandlers
			// Desktop background: HKEY_CLASSES_ROOT\DesktopBackground\shellex\ContextMenuHandlers
			// Same as folder? HKEY_CLASSES_ROOT\Directory\Background\shellex\ContextMenuHandlers
			// Drive: HKEY_CLASSES_ROOT\Drive\shellex\ContextMenuHandlers
			// Windows 7 library folder: HKEY_CLASSES_ROOT\LibraryFolder\background\shellex\ContextMenuHandlers
			// HKEY_CLASSES_ROOT\LibraryFolder\shellex\ContextMenuHandlers
			// HKEY_CLASSES_ROOT\LibraryLocation\ShellEx\ContextMenuHandlers
			// Printers: HKEY_CLASSES_ROOT\Printers\shellex\ContextMenuHandlers
			// Results? HKEY_CLASSES_ROOT\Results\ShellEx\ContextMenuHandlers

			// System file associations: HKEY_CLASSES_ROOT\SystemFileAssociations\.bmp\ShellEx
			// Images: HKEY_CLASSES_ROOT\SystemFileAssociations\image\ShellEx\ContextMenuHandlers
			// Videos: HKEY_CLASSES_ROOT\SystemFileAssociations\video\ShellEx\ContextMenuHandlers

			switch ( association )
			{
				default:
				case AssociationType.AllFiles:
					return @"HKEY_CLASSES_ROOT\AllFilesystemObjects\shellex\ContextMenuHandler";
				case AssociationType.FileExtension:
				case AssociationType.Class:
				case AssociationType.ClassOfExtension:
				case AssociationType.UnknownFiles:
					return @"*\shellex\ContextMenuHandlers\";
				case AssociationType.Directory:
					return @"Folder\shellex\ContextMenuHandlers\";
				//case RightClickContextMenuOptions.ImageFiles:
				//	return @"SystemFileAssociations\image\ShellEx\ContextMenuHandlers\";
				//case RightClickContextMenuOptions.VideoFiles:
				//	return @"SystemFileAssociations\video\ShellEx\ContextMenuHandlers\";
				//case RightClickContextMenuOptions.DesktopBackground:
				//	return @"DesktopBackground\shellex\ContextMenuHandlers\";
				case AssociationType.Drive:
					return @"Drive\shellex\ContextMenuHandlers\";
				//case RightClickContextMenuOptions.Printers:
				//	return @"Printers\shellex\ContextMenuHandlers\";
			}
		}


		public async Task<object> Register( dynamic input )
		{
			string dllPath = (string)input.dllpath;
			dynamic[] menuFormat = (dynamic[])input.menu;
			AssociationType association = Util.IsProperty( input, "association" ) ? StringNameToAssociationType( (string)input.association ) : AssociationType.AllFiles;
			string[] associations = Util.IsProperty( input, "associations" ) ? (string[])input.associations : new string[ 0 ];

			Create( dllPath, menuFormat, association, associations );

			AppDomain.CurrentDomain.AssemblyResolve += ( sender, args ) => OnResolve( sender, args );
			RegistrationServices reg = new RegistrationServices();
			Assembly assembly = Assembly.LoadFile( dllPath );
			reg.RegisterAssembly( assembly, AssemblyRegistrationFlags.SetCodeBase );

			var attribute = (GuidAttribute)assembly.GetCustomAttributes( typeof( GuidAttribute ), true )[ 0 ];
			var clsid = "{" + attribute.Value + "}";
			var programName = assembly.GetName().Name;

			RegistryKey rk = Registry.CurrentUser.OpenSubKey( @"Software\Microsoft\Windows\CurrentVersion\Explorer", true );
			rk.SetValue( @"DesktopProcess", 1 );
			rk.Close();

			// For Winnt set me as an approved shellex
			rk = Registry.LocalMachine.OpenSubKey( @"Software\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved", true );
			rk.SetValue( clsid, programName + @" Shell Extension" );
			rk.Close();

			if ( association != AssociationType.None )
			{
				rk = Registry.ClassesRoot.CreateSubKey( GetContextMenuOptionsBaseKey( association ) + programName );
				rk.SetValue( "", clsid );
				rk.Close();
			}
			return "";
		}


		public async Task<object> Unregister( dynamic input )
		{
			string dllPath = (string)input.dllpath;
			RegistrationServices reg = new RegistrationServices();
			AppDomain.CurrentDomain.AssemblyResolve += ( sender, args ) => OnResolve( sender, args );
			Assembly assembly = Assembly.LoadFile( dllPath );
			reg.UnregisterAssembly( assembly );
			var programName = assembly.GetName().Name;
			
			var associationAttrib = (AssociationStorageAttribute)assembly.GetCustomAttributes( typeof( AssociationStorageAttribute ), true )[ 0 ];
			AssociationType association = associationAttrib.Association;
			var guidAttrib = (GuidAttribute)assembly.GetCustomAttributes(typeof(GuidAttribute),true)[0];
			var clsid = "{" + guidAttrib.Value + "}";

			RegistryKey rk = Registry.LocalMachine.OpenSubKey( @"Software\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved", true );
			try
			{
				rk.DeleteValue( clsid );
			}
			catch ( Exception ) { }
			rk.Close();

			try
			{
				if ( association != AssociationType.None )
				{
					Registry.ClassesRoot.DeleteSubKey( GetContextMenuOptionsBaseKey( association ) + programName );
				}
			}
			catch ( Exception ) {}

			return "";
		}
	}
}

