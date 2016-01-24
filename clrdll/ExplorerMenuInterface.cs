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
		string JoinWithQuotes( string[] str )
		{
			return @"""" + string.Join( @""" """, str ) + @"""";
		}


		private void OnClick( InfoStorageAttribute storage, dynamic item, List<string> filesList )
		{
			string action = Util.IsJsonProperty( item, "action" ) ? (string)item.action : "";

			if ( action.Length > 0 )
			{
				if ( filesList.Count > 0 )
				{
					string[] args = Util.IsJsonProperty( item, "args" ) ? Util.ObjectToStringArray( (object[])item.args ) : new string[ 0 ];
					var style = Util.IsJsonProperty( item, "style" ) ? (string)item.style : "remain";

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
			Assembly assem = this.GetType().Assembly;
			var storage = (InfoStorageAttribute)assem.GetCustomAttributes( typeof( InfoStorageAttribute ), true )[ 0 ];

			MenuBuilder.BuildMenu( assem, this, storage, list, ( dynamic item, List<string> s ) => { OnClick( storage, item, s ); } );
		}
	}


	public class ExplorerMenuInterface
	{

		static Platform.ComRegisterClass.RightClickContextMenuOptions[] StringNameToAssociationType( string[] names )
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
			CreateAssembly.Create( name, guid, dllPath, resourceList, actionPath, menuFormat, association, fileExtensionFilter, expandFileNames );

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

			var associationAttrib = (InfoStorageAttribute)assembly.GetCustomAttributes( typeof( InfoStorageAttribute ), true )[ 0 ];
			Platform.ComRegisterClass.RightClickContextMenuOptions[] associationArray = associationAttrib.Association;
			var guidAttrib = (GuidAttribute)assembly.GetCustomAttributes(typeof(GuidAttribute),true)[0];

			Platform.ComRegisterClass.RegisterShellRightClickContextMenu( associationAttrib.Name, "{" + guidAttrib.Value + "}", associationArray );
			Platform.ComRegisterClass.UnregisterServer( assembly );

			return "";
		}
	}
}

