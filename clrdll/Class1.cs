using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using System.Threading;
using System.Reflection;
using System.Reflection.Emit;



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

	//[ DateLastUpdated( "33-11-3333" ) ]
    public class Class1
    {
		public Class1()
		{
		}

		public Class1( string s )
		{
		}

		public async Task<object> TestMethod( dynamic input )
		{
			Test();
			return "";
		}

		public void Test()
		{
			DateLastUpdated MyAttribute = (DateLastUpdated)Attribute.GetCustomAttribute( this.GetType(), typeof( DateLastUpdated ) );

			MessageBox.Show( "Test works: " + ( MyAttribute != null ? MyAttribute.DateUpdated : "" ) );
		}
    }


	public class CreateComAssembly
	{
		public async Task<object> Create( dynamic input )
		{
			AppDomain currentDomain = Thread.GetDomain();

			AssemblyName myAsmName = new AssemblyName();
			myAsmName.Name = "MyAssembly";

			AssemblyBuilder myAsmBuilder = currentDomain.DefineDynamicAssembly(
							   myAsmName, AssemblyBuilderAccess.Run );

			ModuleBuilder myModBuilder = myAsmBuilder.DefineDynamicModule( "MyModule" );

			// First, we'll build a type with a custom attribute attached.

			TypeBuilder myTypeBuilder = myModBuilder.DefineType( "MyType",
								TypeAttributes.Public, typeof( Class1 ) );
			 //TypeAttributes.Public
			 //   | TypeAttributes.Class
			 //   | TypeAttributes.AutoClass
			 //   | TypeAttributes.AnsiClass
			 //   | TypeAttributes.ExplicitLayout,
			 //   typeof(SomeOtherNamespace.MyBase));

			Type[] ctorParams = new Type[] { typeof( string ) };
			ConstructorInfo classCtorInfo = typeof( DateLastUpdated ).GetConstructor( ctorParams );

			CustomAttributeBuilder myCABuilder = new CustomAttributeBuilder(
								classCtorInfo,
								new object[] { "Custom attribute worked!" } );

			myTypeBuilder.SetCustomAttribute( myCABuilder );

			var mynewtype = myTypeBuilder.CreateType();
			Class1 mynewinstance = (Class1)Activator.CreateInstance( mynewtype );
			mynewinstance.Test();
			return "";
		}
	}
}

