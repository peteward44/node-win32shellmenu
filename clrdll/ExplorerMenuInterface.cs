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
			//  Create the menu strip.
			var menu = new ContextMenuStrip();

			//  Create a 'count lines' item.
			var itemCountLines = new ToolStripMenuItem
			{
				Text = "Count Lines..."
	//			Image = Properties.Resources.CountLines
			};

			//  When we click, we'll count the lines.
			itemCountLines.Click += ( sender, args ) => OnClick();

			//  Add the item to the context menu.
			menu.Items.Add( itemCountLines );

			//  Return the menu.
			return menu;
		}

		private void OnClick()
		{
			MessageBox.Show( "Test worked" );
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


		public async Task<object> Create( dynamic input )
		{
			string dllPath = (string)input.dllpath;

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

			AddAttribute( myTypeBuilder, typeof( ComVisibleAttribute ), true );
			AddAttribute( myTypeBuilder, typeof( COMServerAssociationAttribute ), AssociationType.FileExtension, new string[] { ".txt" } );
			AddAttribute( myAsmBuilder, typeof( GuidAttribute ), "a64f5783-4e6d-4fd1-8ca7-aa50cdd144e6" );

			//Type[] ctorParams = new Type[] { typeof( bool ) };
			//ConstructorInfo classCtorInfo = typeof( ComVisibleAttribute ).GetConstructor( ctorParams );
			//CustomAttributeBuilder myCABuilder = new CustomAttributeBuilder( classCtorInfo, new object[] { true } );
			//myTypeBuilder.SetCustomAttribute( myCABuilder );

			//Type[] ctorParams2 = new Type[] { AssociationType.FileExtension.GetType(), typeof( string[] ) };
			//ConstructorInfo classCtorInfo2 = typeof( COMServerAssociationAttribute ).GetConstructor( ctorParams2 );
			//CustomAttributeBuilder myCABuilder2 = new CustomAttributeBuilder( classCtorInfo2, new object[] { AssociationType.FileExtension, new string[] { ".txt" } } );
			//myTypeBuilder.SetCustomAttribute( myCABuilder2 );

			//Type[] ctorParams3 = new Type[] { typeof( string ) };
			//ConstructorInfo classCtorInfo3 = typeof( GuidAttribute ).GetConstructor( ctorParams3 );
			//CustomAttributeBuilder myCABuilder3 = new CustomAttributeBuilder( classCtorInfo3, new object[] { "a64f5783-4e6d-4fd1-8ca7-aa50cdd144e6" } );
			//myAsmBuilder.SetCustomAttribute( myCABuilder3 );


			myTypeBuilder.CreateType();
			myModBuilder.CreateGlobalFunctions();

			myAsmBuilder.Save( System.IO.Path.GetFileName( dllPath ) );

			return "";
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


		public enum RightClickContextMenuOptions
		{
			AllFileSystemObjects,
			Files,
			Folders,
			ImageFiles,
			VideoFiles,
			DesktopBackground,
			Drive,
			Printers,
		}


		static string GetContextMenuOptionsBaseKey( RightClickContextMenuOptions contextMenu )
		{
			switch ( contextMenu )
			{
				case RightClickContextMenuOptions.AllFileSystemObjects:
					return @"HKEY_CLASSES_ROOT\AllFilesystemObjects\shellex\ContextMenuHandler";
				case RightClickContextMenuOptions.Files:
					return @"*\shellex\ContextMenuHandlers\";
				case RightClickContextMenuOptions.Folders:
					return @"Folder\shellex\ContextMenuHandlers\";
				case RightClickContextMenuOptions.ImageFiles:
					return @"SystemFileAssociations\image\ShellEx\ContextMenuHandlers\";
				case RightClickContextMenuOptions.VideoFiles:
					return @"SystemFileAssociations\video\ShellEx\ContextMenuHandlers\";
				case RightClickContextMenuOptions.DesktopBackground:
					return @"DesktopBackground\shellex\ContextMenuHandlers\";
				case RightClickContextMenuOptions.Drive:
					return @"Drive\shellex\ContextMenuHandlers\";
				case RightClickContextMenuOptions.Printers:
					return @"Printers\shellex\ContextMenuHandlers\";
				default:
					System.Diagnostics.Debug.Assert( false );
					return "";
			}
		}


		public async Task<object> Register( dynamic input )
		{
			string dllPath = (string)input.dllpath;
			AppDomain.CurrentDomain.AssemblyResolve += ( sender, args ) => OnResolve( sender, args );
			RegistrationServices reg = new RegistrationServices();
			Assembly assembly = Assembly.LoadFile( dllPath );
			reg.RegisterAssembly( assembly, AssemblyRegistrationFlags.SetCodeBase );

			var attribute = (GuidAttribute)assembly.GetCustomAttributes( typeof( GuidAttribute ), true )[ 0 ];
			var clsid = "{" + attribute.Value + "}";
			var programName = assembly.GetName().Name;

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

			RegistryKey rk = Registry.CurrentUser.OpenSubKey( @"Software\Microsoft\Windows\CurrentVersion\Explorer", true );
			rk.SetValue( @"DesktopProcess", 1 );
			rk.Close();

			// For Winnt set me as an approved shellex
			rk = Registry.LocalMachine.OpenSubKey( @"Software\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved", true );
			rk.SetValue( clsid, programName + @" Shell Extension" );
			rk.Close();

			RightClickContextMenuOptions[] contextMenuOptions = { RightClickContextMenuOptions.Files };

			foreach ( RightClickContextMenuOptions contextMenu in contextMenuOptions )
			{
				rk = Registry.ClassesRoot.CreateSubKey( GetContextMenuOptionsBaseKey( contextMenu ) + programName );
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

			var attribute = (GuidAttribute)assembly.GetCustomAttributes(typeof(GuidAttribute),true)[0];
			var clsid = "{" + attribute.Value + "}";

			RegistryKey rk = Registry.LocalMachine.OpenSubKey( @"Software\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved", true );
			rk.DeleteValue( clsid );
			rk.Close();

			RightClickContextMenuOptions[] contextMenuOptions = { RightClickContextMenuOptions.Files };

			foreach ( RightClickContextMenuOptions contextMenu in contextMenuOptions )
			{
				Registry.ClassesRoot.DeleteSubKey( GetContextMenuOptionsBaseKey( contextMenu ) + programName );
			}

			return "";
		}
	}
}

