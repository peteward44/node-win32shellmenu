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


namespace windowsexplorermenu_clr
{
	public class DateLastUpdated : Attribute
	{
		private string dateUpdated;
		public string DateUpdated
		{
			get
			{
				return dateUpdated;
			}
		}

		public DateLastUpdated( string theDate )
		{
			this.dateUpdated = theDate;
		}

	}



	[ComVisible( true )]
	[COMServerAssociation( AssociationType.FileExtension, ".txt" )]
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
			////  Builder for the output.
			//var builder = new StringBuilder();

			////  Go through each file.
			//foreach ( var filePath in SelectedItemPaths )
			//{
			//	//  Count the lines.
			//	builder.AppendLine( string.Format( "{0} - {1} Lines",
			//	  System.IO.Path.GetFileName( filePath ), System.IO.File.ReadAllLines( filePath ).Length ) );
			//}

			////  Show the ouput.
			//MessageBox.Show( builder.ToString() );
		}
	}


	////[ DateLastUpdated( "33-11-3333" ) ]
	//public class Class1
	//{
	//	public Class1()
	//	{
	//	}

	//	public Class1( string s )
	//	{
	//	}

	//	public async Task<object> TestMethod( dynamic input )
	//	{
	//		Test();
	//		return "";
	//	}

	//	public void Test()
	//	{
	//		DateLastUpdated MyAttribute = (DateLastUpdated)Attribute.GetCustomAttribute( this.GetType(), typeof( DateLastUpdated ) );

	//		MessageBox.Show( "Test works: " + ( MyAttribute != null ? MyAttribute.DateUpdated : "" ) );
	//	}
	//}


	public class CreateComAssembly
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
			//parms.Flags = CspProviderFlags.UseUserProtectedKey;
			//var password = new System.Security.SecureString();
			//password.AppendChar( 'p' );
			//parms.KeyPassword = password;
			RSACryptoServiceProvider provider = new RSACryptoServiceProvider( keySize, parms );
			byte[] array = provider.ExportCspBlob( !provider.PublicOnly );
			return array;
		}


		public async Task<object> Create( dynamic input )
		{
			string dllPath = (string)input.dllpath;

			AssemblyName myAsmName = new AssemblyName( System.IO.Path.GetFileNameWithoutExtension( dllPath ) );
			myAsmName.CodeBase = String.Concat( "file:///", System.IO.Path.GetDirectoryName( dllPath ) );
			myAsmName.CultureInfo = new System.Globalization.CultureInfo( "en-US" );
			myAsmName.KeyPair = new StrongNameKeyPair( CreateKeyPair( System.IO.Path.GetFileNameWithoutExtension( dllPath ), 1024 ) );
		//	myAsmName.KeyPair = new StrongNameKeyPair( System.IO.File.Open( System.IO.Path.Combine( System.IO.Path.GetDirectoryName( System.Reflection.Assembly.GetExecutingAssembly().Location ), "mysnk.snk" ), System.IO.FileMode.Open ) );
			myAsmName.Flags = AssemblyNameFlags.PublicKey;
			myAsmName.VersionCompatibility = AssemblyVersionCompatibility.SameProcess;
			myAsmName.HashAlgorithm = AssemblyHashAlgorithm.SHA1;
			myAsmName.Version = new Version( "1.0.0.0" );

			AssemblyBuilder myAsmBuilder = AppDomain.CurrentDomain.DefineDynamicAssembly( myAsmName, AssemblyBuilderAccess.Save, System.IO.Path.GetDirectoryName( dllPath ) );
			ModuleBuilder myModBuilder = myAsmBuilder.DefineDynamicModule( "MyModule", System.IO.Path.GetFileName( dllPath ) );

			TypeBuilder myTypeBuilder = myModBuilder.DefineType( "MyType", TypeAttributes.Public, typeof( MenuExtension ) );
			 //TypeAttributes.Public
			 //   | TypeAttributes.Class
			 //   | TypeAttributes.AutoClass
			 //   | TypeAttributes.AnsiClass
			 //   | TypeAttributes.ExplicitLayout,
			 //   typeof(SomeOtherNamespace.MyBase));

			//Type[] ctorParams = new Type[] { typeof( string ) };
			//ConstructorInfo classCtorInfo = typeof( DateLastUpdated ).GetConstructor( ctorParams );
			//CustomAttributeBuilder myCABuilder = new CustomAttributeBuilder( classCtorInfo, new object[] { "Custom attribute worked!" } );
			//myTypeBuilder.SetCustomAttribute( myCABuilder );
			Type[] ctorParams = new Type[] { typeof( bool ) };
			ConstructorInfo classCtorInfo = typeof( ComVisibleAttribute ).GetConstructor( ctorParams );
			CustomAttributeBuilder myCABuilder = new CustomAttributeBuilder( classCtorInfo, new object[] { true } );
			myTypeBuilder.SetCustomAttribute( myCABuilder );

			Type[] ctorParams2 = new Type[] { AssociationType.FileExtension.GetType(), typeof( string[] ) };
			ConstructorInfo classCtorInfo2 = typeof( COMServerAssociationAttribute ).GetConstructor( ctorParams2 );
			CustomAttributeBuilder myCABuilder2 = new CustomAttributeBuilder( classCtorInfo2, new object[] { AssociationType.FileExtension, new string[] { ".txt" } } );
			myTypeBuilder.SetCustomAttribute( myCABuilder2 );

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


		//public async Task<object> Register( dynamic input )
		//{
		//	string dllPath = (string)input.dllpath;
		//	AppDomain.CurrentDomain.AssemblyResolve += ( sender, args ) => OnResolve( sender, args );
		//	RegistrationServices reg = new RegistrationServices();
		//	reg.RegisterAssembly( Assembly.LoadFile( dllPath ), AssemblyRegistrationFlags.SetCodeBase );
		//	return "";
		//}


		//public async Task<object> Unregister( dynamic input )
		//{
		//	string dllPath = (string)input.dllpath;
		//	RegistrationServices reg = new RegistrationServices();
		//	reg.UnregisterAssembly( Assembly.LoadFile( dllPath ) );
		//	return "";
		//}
	}
}

