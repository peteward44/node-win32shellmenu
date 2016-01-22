using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;


namespace windowsexplorermenu_clr
{
	// Custom attribute to store the file association data at the assembly level.
	public class InfoStorageAttribute : Attribute
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
		public InfoStorageAttribute( string name, string actionPath, string menuFormat, Platform.ComRegisterClass.RightClickContextMenuOptions[] _at, string[] fileExtensionFilter, bool expandFileNames )
		{
			this.Name = name;
			this.ActionPath = actionPath;
			this.MenuFormat = menuFormat;
			this.Association = _at;
			this.FileExtensionFilter = fileExtensionFilter;
			this.ExpandFileNames = expandFileNames;
		}

	}
}

