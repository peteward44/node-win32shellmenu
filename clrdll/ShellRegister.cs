using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text;
using Microsoft.Win32;
using WinForms = System.Windows.Forms;
using Platform.ShellExt;


namespace Platform
{
	/// <summary>
	/// How to use: 
	/// Use the helper methods to hook onto common shell extension with your own delegates (RegisterShellRightClickContextMenu).
	/// Call RegisterServer() and UnregisterServer() from your application. This only needs needs to be called on installation / uninstall
	/// Hook the ComRegister & ComUnregister events to define your own registration methods.
	/// </summary>
	public class ComRegisterClass
	{
		public delegate void ComRegisterDelegate();
		public static event ComRegisterDelegate ComRegister, ComUnregister;


		private ComRegisterClass()
		{}


		[System.Runtime.InteropServices.ComRegisterFunctionAttribute()]
		static void ComRegisterServerPrivate( Type t )
		{
			try
			{
				RegisterServer( Assembly.GetEntryAssembly() );
			}
			catch ( System.Exception )
			{
			}
		}

		[System.Runtime.InteropServices.ComUnregisterFunctionAttribute()]
		static void ComUnregisterServerPrivate( Type t )
		{
			try
			{
				UnregisterServer( Assembly.GetEntryAssembly() );
			}
			catch (System.Exception)
			{
			}
		}


		public static void RegisterServer( Assembly assemblyToRegister )
		{
			if ( ComRegister != null )
			{
				RegistrationServices reg = new RegistrationServices();
				reg.RegisterAssembly( assemblyToRegister, AssemblyRegistrationFlags.SetCodeBase );
				ComRegister();
			}
		}

		public static void UnregisterServer( Assembly assemblyToRegister )
		{
			if ( ComUnregister != null )
			{
				RegistrationServices reg = new RegistrationServices();
				reg.UnregisterAssembly( assemblyToRegister );
				ComUnregister();
			}
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
					return @"AllFilesystemObjects\shellex\ContextMenuHandlers\";
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


		public static void RegisterShellRightClickContextMenu( string programName, string clsid, params RightClickContextMenuOptions[] contextMenuOptions )
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

			ComRegister += delegate()
			{
				RegistryKey rk = Registry.CurrentUser.OpenSubKey( @"Software\Microsoft\Windows\CurrentVersion\Explorer", true );
				rk.SetValue( @"DesktopProcess", 1 );
				rk.Close();

				// For Winnt set me as an approved shellex
				rk = Registry.LocalMachine.OpenSubKey( @"Software\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved", true );
				rk.SetValue( clsid, programName + @" Shell Extension" );
				rk.Close();

				foreach ( RightClickContextMenuOptions contextMenu in contextMenuOptions )
				{
					rk = Registry.ClassesRoot.CreateSubKey( GetContextMenuOptionsBaseKey( contextMenu ) + programName );
					rk.SetValue( "", clsid );
					rk.Close();
				}
			};

			ComUnregister += delegate()
			{
				RegistryKey rk = Registry.LocalMachine.OpenSubKey( @"Software\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved", true );
				rk.DeleteValue( clsid );
				rk.Close();

				foreach ( RightClickContextMenuOptions contextMenu in contextMenuOptions )
				{
					Registry.ClassesRoot.DeleteSubKey( GetContextMenuOptionsBaseKey( contextMenu ) + programName );
				}
			};
		}
	}


	 //<summary>
	 //If the ComRegister.RegisterShellRightClickContextMenu() has been called, you need to derive from this class and be sure to specify 
	 //the Guid and ComVisible attributes on it. This is where the names & types of the context menu options you are adding are defined.
	 //</summary>
	 //[Guid("2AA8DDCB-0540-4cd3-BD31-D91DADD81ED4"), ComVisible(true)]
	public abstract class ShellRightClickContextMenuClass : IContextMenu, IShellExtInit
	{
	    protected IDataObject mDataObject = null;
		protected IntPtr mDropHandle = IntPtr.Zero;
	    StringBuilder mStringBuilder = new StringBuilder( 65355 );

		int mMenuItemsAddedCount = 0;
		IntPtr mMenuHandle = IntPtr.Zero;
		int mMenuCmdFirst = 0;
		int mMenuMaximumItemsWeCanAdd = 0;

		public delegate void RightClickActionDelegate( List<string> selectedFiles );

		Dictionary<int, RightClickActionDelegate> mMenuIdsMap = new Dictionary<int, RightClickActionDelegate>(); // Windows ID vs. Our ID

	    protected abstract string GetVerbString();
	    protected abstract string GetHelpString();

		/// <summary>
		/// called when the menu is about to be displayed, add context menu options here using InsertMenuItem()
		/// </summary>
		/// <param name="list"></param>
		/// <returns></returns>
		protected abstract void OnBuildMenu( List<string> list );


		/// <summary>
		/// Call this within OnBuildMenu() to build menu
		/// </summary>
		protected uint InsertMenuItem( string textName, int position, RightClickActionDelegate actionDelegate )
		{
			return InsertMenuItemPrivate( mMenuHandle, textName, -1, actionDelegate );
		}

		
		protected IntPtr CreateSubMenu()
		{
			return DllImports.CreatePopupMenu();
		}


		protected void DestroySubMenu( IntPtr menu )
		{
			DllImports.DestroyMenu( menu );
		}


		protected uint InsertMenuItemIntoSubMenu( IntPtr subMenu, string textName, int position, RightClickActionDelegate actionDelegate )
		{
			return InsertMenuItemPrivate( subMenu, textName, position, actionDelegate );
		}

		protected uint InsertMenuItemIntoSubMenu( IntPtr subMenu, string textName, int position, Bitmap bitmap, RightClickActionDelegate actionDelegate )
		{
			return InsertMenuItemIntoSubMenu( subMenu, textName, position, bitmap, bitmap, actionDelegate );
		}

		protected uint InsertMenuItemIntoSubMenu( IntPtr subMenu, string textName, int position, Bitmap uncheckedBitmap, Bitmap checkedBitmap, RightClickActionDelegate actionDelegate )
		{
			uint id = InsertMenuItemPrivate( subMenu, textName, position, actionDelegate );
			if ( id != 0 )
			{
				SetMenuItemBitmapPrivate( subMenu, id, uncheckedBitmap, checkedBitmap );
			}
			return id;
		}


		protected uint InsertSeperator( int position )
		{
			return InsertSeperatorPrivate( mMenuHandle, position );
		}

		protected uint InsertSeperatorIntoSubMenu( IntPtr subMenu, int position )
		{
			return InsertSeperatorPrivate( subMenu, position );
		}


		protected uint InsertSubMenu( IntPtr subMenu, string textName, int position, Bitmap bitmap )
		{
			return InsertSubMenu( subMenu, textName, position, bitmap, bitmap );
		}


		protected uint InsertSubMenu( IntPtr subMenu, string textName, int position, Bitmap uncheckedBitmap, Bitmap checkedBitmap )
		{
			uint id = InsertSubMenu( subMenu, textName, position );
			if ( id != 0 )
			{
				SetMenuItemBitmapPrivate( mMenuHandle, id, uncheckedBitmap, checkedBitmap );
			}
			return id;
		}


		protected uint InsertSubMenu( IntPtr subMenu, string textName, int position )
		{
			return InsertSubMenuIntoSubMenu( mMenuHandle, subMenu, textName, position );
		}



		protected uint InsertSubMenuIntoSubMenu( IntPtr menu, IntPtr subMenu, string textName, int position )
		{
			MENUITEMINFO mii = new MENUITEMINFO();
			mii.cbSize = (uint)Marshal.SizeOf( typeof( MENUITEMINFO ) );
			mii.fMask = (uint)MIIM.ID | (uint)MIIM.STRING | (uint)MIIM.SUBMENU;
			mii.wID = (uint)( mMenuCmdFirst + mMenuItemsAddedCount );
			mii.fType = (uint)MF.STRING;
			mii.dwTypeData = textName;
			mii.cch = 0;
			mii.fState = (uint)MF.ENABLED;
			mii.hSubMenu = subMenu;

			// Add it to the item
			if ( DllImports.InsertMenuItem( menu, (uint)position, true, ref mii ) == 0 )
			{
				// failed
				return 0;
			}
			else
			{
				mMenuItemsAddedCount++;
				return mii.wID;
			}
		}


		protected uint InsertSeperatorPrivate( IntPtr subMenu, int position )
		{
			MENUITEMINFO mii = new MENUITEMINFO();
			mii.cbSize = ( uint )Marshal.SizeOf( typeof( MENUITEMINFO ) );
			mii.fMask = ( uint )MIIM.ID;
			mii.wID = ( uint )( mMenuCmdFirst + mMenuItemsAddedCount );
			mii.fType = ( uint )MF.SEPARATOR;
			mii.cch = 0;
			mii.fState = ( uint )MF.ENABLED;

			// Add it to the item
			if ( DllImports.InsertMenuItem( subMenu, (uint)position, true, ref mii ) == 0 )
			{
				// failed
				return 0;
			}
			else
			{
				mMenuItemsAddedCount++;
				return mii.wID;
			}
		}


		protected bool SetSubMenuItemBitmap( IntPtr menuHandle, uint id, Bitmap bitmap )
		{
			return SetMenuItemBitmapPrivate( menuHandle, id, bitmap, bitmap );
		}

		protected bool SetSubMenuItemBitmap( IntPtr menuHandle, uint id, Bitmap uncheckBitmap, Bitmap checkedBitmap )
		{
			return SetMenuItemBitmapPrivate( menuHandle, id, uncheckBitmap, checkedBitmap );
		}

		protected bool SetMenuItemBitmap( uint id, Bitmap uncheckBitmap, Bitmap checkedBitmap )
		{
			return SetMenuItemBitmapPrivate( mMenuHandle, id, uncheckBitmap, checkedBitmap );
		}

		protected bool SetMenuItemBitmap( uint id, Bitmap bitmap )
		{
			return SetMenuItemBitmapPrivate( mMenuHandle, id, bitmap, bitmap );
		}


	    #region Functional Implementation

		int mBitmapCorrectWidth = 0, mBitmapCorrectHeight = 0;


		public ShellRightClickContextMenuClass()
		{
			const int SM_CXMENUCHECK = 71;
			const int SM_CYMENUCHECK = 72;

			mBitmapCorrectWidth = DllImports.GetSystemMetrics( SM_CXMENUCHECK );
			mBitmapCorrectHeight = DllImports.GetSystemMetrics( SM_CYMENUCHECK );
		}

	    /// <summary>
	    /// Returns list of files & folders that are currently selected by the mouse
	    /// </summary>
	    /// <returns></returns>
	    protected List<string> GetSelectedFiles()
	    {
	        List<string> list = new List<string>();
			if ( mDropHandle != IntPtr.Zero )
			{
				uint nselected = DllImports.DragQueryFile( mDropHandle, 0xFFFFFFFF, null, 0 );
				for ( uint i = 0; i < nselected; ++i )
				{
					if ( DllImports.DragQueryFile( mDropHandle, i, mStringBuilder, ( uint )mStringBuilder.Capacity ) > 0 )
						list.Add( mStringBuilder.ToString().Trim() );
				}
			}

	        return list;
	    }


		private uint InsertMenuItemPrivate( IntPtr menuHandle, string textName, int position, RightClickActionDelegate actionDelegate )
		{
			if ( mMenuItemsAddedCount >= mMenuMaximumItemsWeCanAdd )
				return 0;

			// Create a new Menu Item to add to the popup menu
			MENUITEMINFO mii = new MENUITEMINFO();
			mii.cbSize = ( uint )Marshal.SizeOf( typeof( MENUITEMINFO ) );
			mii.fMask = ( uint )MIIM.ID | ( uint )MIIM.TYPE | ( uint )MIIM.STATE;
			mii.wID = ( uint )( mMenuCmdFirst + mMenuItemsAddedCount );
			mii.fType = ( uint )MF.STRING;
			mii.dwTypeData = textName;
			mii.cch = 0;
			mii.fState = ( uint )MF.ENABLED;

			// Add it to the item
			if ( DllImports.InsertMenuItem( menuHandle, (uint)position, true, ref mii ) == 0 )
			{
				// failed
				//int error = Marshal.GetLastWin32Error();
				//System.Windows.Forms.MessageBox.Show( "InsertMenuItem failed + 0x" + error.ToString( "X" ) + " : " + ( new System.ComponentModel.Win32Exception( error ) ).Message );
				return 0;
			}
			else
			{
				mMenuIdsMap.Add( mMenuItemsAddedCount, actionDelegate );
				mMenuItemsAddedCount++;
				return mii.wID;
			}
		}


		private bool SetMenuItemBitmapPrivate( IntPtr menuHandle, uint id, Bitmap uncheckBitmap, Bitmap checkedBitmap )
		{
		//	WinForms.MessageBox.Show( "Size: " + mBitmapCorrectWidth + "x" + mBitmapCorrectHeight );

			if ( !DllImports.SetMenuItemBitmaps( menuHandle, id, ( uint )MF.BITMAP | (uint)MF.BYCOMMAND, uncheckBitmap != null ? uncheckBitmap.GetHbitmap() : IntPtr.Zero, checkedBitmap != null ? checkedBitmap.GetHbitmap() : IntPtr.Zero ) )
			{
				WinForms.MessageBox.Show( "Message: " + ( new System.ComponentModel.Win32Exception( Marshal.GetLastWin32Error() ) ).Message );
				return false;
			}
			else
				return true;
		}


	    #endregion


	    #region IContextMenu Members


		uint MakeHResult( bool success, uint fac, uint code )
		{
			return ( ( ( uint )( success ? 0 : 1 ) << 31 ) | ( ( uint )( fac ) << 16 ) | ( ( uint )( code ) ) );
		}


		uint IContextMenu.QueryContextMenu( IntPtr hmenu, uint iMenu, int idCmdFirst, int idCmdLast, uint uFlags )
	    {
	        try
	        {
				if ( mDropHandle == IntPtr.Zero || ( uFlags & (uint)CMF.CMF_DEFAULTONLY ) != 0 )
					return MakeHResult( true, 0, 1 );

	            List<string> list = GetSelectedFiles();

				mMenuHandle = hmenu;
				mMenuCmdFirst = idCmdFirst;
				mMenuItemsAddedCount = 0;
				mMenuMaximumItemsWeCanAdd = idCmdLast - idCmdFirst;
				mMenuIdsMap.Clear();
				OnBuildMenu( list );

				return MakeHResult( true, 0, ((uint)mMenuItemsAddedCount + 1) );
	        }
	        catch ( System.Exception )
	        {
				return MakeHResult( true, 0, 1 );
	        }
	    }


		void IContextMenu.InvokeCommand( IntPtr cc )  
		{
			CMINVOKECOMMANDINFO invokeCommandEx = ( CMINVOKECOMMANDINFO )Marshal.PtrToStructure( cc, typeof( CMINVOKECOMMANDINFO ) );

			try
			{
				uint highWord = ( invokeCommandEx.lpVerb >> 16 );
				if ( highWord == 0 )
				{
					int lowWord = (int)( invokeCommandEx.lpVerb & 0xFFFF );
					if ( mMenuIdsMap.ContainsKey( lowWord ) )
						mMenuIdsMap[ lowWord ]( GetSelectedFiles() );
				}
				else
				{
					// else highWord is the pointer to a string that describes a non-language
					// specific command
				}
			}
			catch ( Exception )
			{
			}
	    }


		void IContextMenu.GetCommandString( IntPtr idcmd, uint uflags, uint reserved, StringBuilder commandstring, uint cch )
	    {
	        try
	        {
	            switch(uflags)
	            {
	            case (uint)GCS.VERB:
	                commandstring = new StringBuilder( GetVerbString().Substring(1, (int)cch-1));
	                break;
	            case (uint)GCS.HELPTEXT:
					commandstring = new StringBuilder( GetHelpString().Substring(1, (int)cch-1));
	                break;
	            case (uint)GCS.VALIDATE:
	                break;
	            }
	        }
	        catch ( Exception )
	        {
	        }
	    }

	    #endregion



	    #region IShellExtInit Members


	    int IShellExtInit.Initialize( IntPtr pidlFolder, IntPtr lpdobj, uint hKeyProgID )
	    {
	        try
	        {
				mDataObject = null;
	            if ( lpdobj != IntPtr.Zero )
	            {
	                // Get info about the directory
					mDataObject = ( ShellExt.IDataObject )Marshal.GetObjectForIUnknown( lpdobj );
	                FORMATETC fmt = new FORMATETC();
	                fmt.cfFormat = (UInt16)CLIPFORMAT.CF_HDROP;
	                fmt.ptd = IntPtr.Zero;
	                fmt.dwAspect = (UInt32)DVASPECT.DVASPECT_CONTENT;
	                fmt.lindex = -1;
					fmt.tymed = ( UInt32 )TYMED.TYMED_HGLOBAL;
	                STGMEDIUM medium = new STGMEDIUM();
					if ( mDataObject.GetData( ref fmt, ref medium ) == 0 )
						mDropHandle = medium.hGlobal;
					else
						mDropHandle = IntPtr.Zero;
	            }
	        }
	        catch ( Exception )
	        {
	        }

	        return 0;
	    }


	    #endregion
	}
}
