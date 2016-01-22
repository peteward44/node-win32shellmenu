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
	public class CreateAssembly
	{

		static byte[] CreateKeyPair( string containerName, int keySize )
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


		static void AddAttribute( dynamic targetType, Type attributeType, params object[] constructorParams )
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


		public static void Create( string name, Guid guid, string dllPath, IDictionary<string, object> resourceList, string actionPath,
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
			AddAttribute( myAsmBuilder, typeof( InfoStorageAttribute ), name, actionPath, menuFormatJson, association, fileExtensionFilter, expandFileNames );

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
	}
}

