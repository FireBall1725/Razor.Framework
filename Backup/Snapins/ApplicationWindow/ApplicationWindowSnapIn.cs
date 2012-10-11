using System;
using System.Diagnostics;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.IO;
using System.Reflection;
using System.Windows.Forms;

using Razor;
using Razor.Attributes;
using Razor.Configuration;
using Razor.Features;
using Razor.SnapIns;
using Razor.SnapIns.WindowPositioningEngine;

namespace Razor.SnapIns.ApplicationWindow
{
	/// <summary>
	/// Provides a generic window plugin that other plugins can modify as their main window
	/// </summary>
	[SnapInTitle("Application Window")]	
	[SnapInDescription("Provides a generic window to be used as the main window for an application.")]
	[SnapInCompany("CodeReflection")]
	[SnapInDevelopers("Mark Belles")]
	[SnapInVersion("1.0.0.0")]	
	[SnapInImage(typeof(ApplicationWindowSnapIn))]
	[SnapInDependency(typeof(WindowPositioningEngine.WindowPositioningEngineSnapIn))]
	public class ApplicationWindowSnapIn : SnapInWindow, IWindowPositioningEngineFeaturable	
	{
		protected static ApplicationWindowSnapIn _theInstance;			
		protected bool _exitApplicationThreadOnClose = true;
		protected bool _noPromptsOnClose;	
		public const string WindowPositioningEngineKey = "Main Window";
		private System.Windows.Forms.MainMenu _mainMenu;
		private System.Windows.Forms.MenuItem _menuItemFile;
		private System.Windows.Forms.MenuItem _menuItemTools;
		private System.Windows.Forms.MenuItem _menuItemHelp;
		private System.Windows.Forms.MenuItem menuItem1;
		private System.Windows.Forms.MenuItem menuItem2;
		private System.Windows.Forms.MenuItem menuItem3;
		private System.Windows.Forms.StatusBar _statusBar;

		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;
		
		/// <summary>
		/// Returns the one and only instance of the ApplicationWindowSnapIn
		/// </summary>
		public static ApplicationWindowSnapIn Instance
		{
			get
			{
				return _theInstance;
			}
		}

		/// <summary>
		/// Initializes a new instance of the ApplicationWindowSnapIn class
		/// </summary>
		public ApplicationWindowSnapIn() : base()
		{
			_theInstance = this;
			this.InitializeComponent();

			if (SnapInHostingEngine.Instance.SplashWindow != null)
				SnapInHostingEngine.Instance.SplashWindow.Closed += new EventHandler(OnSplashWindowClosed);

			this.Start += new EventHandler(OnSnapInStart);
			this.Stop += new EventHandler(OnSnapInStop);

			this.FileMenuItem.Popup += new EventHandler(OnFileMenuItemPopup);
			this.ToolsMenuItem.Popup += new EventHandler(OnToolsMenuItemPopup);
			this.HelpMenuItem.Popup += new EventHandler(OnHelpMenuItemPopup);
		}

		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		protected override void Dispose( bool disposing )
		{
			if( disposing )
			{
				if(components != null)
				{
					components.Dispose();
				}
			}
			base.Dispose( disposing );
		}

		#region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			this._mainMenu = new System.Windows.Forms.MainMenu();
			this._menuItemFile = new System.Windows.Forms.MenuItem();
			this.menuItem2 = new System.Windows.Forms.MenuItem();
			this._menuItemTools = new System.Windows.Forms.MenuItem();
			this.menuItem1 = new System.Windows.Forms.MenuItem();
			this._menuItemHelp = new System.Windows.Forms.MenuItem();
			this.menuItem3 = new System.Windows.Forms.MenuItem();
			this._statusBar = new System.Windows.Forms.StatusBar();
			this.SuspendLayout();
			// 
			// _mainMenu
			// 
			this._mainMenu.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
																					  this._menuItemFile,
																					  this._menuItemTools,
																					  this._menuItemHelp});
			// 
			// _menuItemFile
			// 
			this._menuItemFile.Index = 0;
			this._menuItemFile.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
																						  this.menuItem2});
			this._menuItemFile.Text = "File";
			// 
			// menuItem2
			// 
			this.menuItem2.Index = 0;
			this.menuItem2.Text = "Null";
			// 
			// _menuItemTools
			// 
			this._menuItemTools.Index = 1;
			this._menuItemTools.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
																						   this.menuItem1});
			this._menuItemTools.Text = "Tools";
			// 
			// menuItem1
			// 
			this.menuItem1.Index = 0;
			this.menuItem1.Text = "Null";
			// 
			// _menuItemHelp
			// 
			this._menuItemHelp.Index = 2;
			this._menuItemHelp.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
																						  this.menuItem3});
			this._menuItemHelp.Text = "Help";
			// 
			// menuItem3
			// 
			this.menuItem3.Index = 0;
			this.menuItem3.Text = "Null";
			// 
			// _statusBar
			// 
			this._statusBar.Location = new System.Drawing.Point(0, 244);
			this._statusBar.Name = "_statusBar";
			this._statusBar.Size = new System.Drawing.Size(292, 22);
			this._statusBar.TabIndex = 0;
			// 
			// ApplicationWindowSnapIn
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(292, 266);
			this.Controls.Add(this._statusBar);
			this.Menu = this._mainMenu;
			this.Name = "ApplicationWindowSnapIn";
			this.Text = "ApplicationWindowSnapIn";
			this.ResumeLayout(false);

		}
		#endregion

		#region My SnapIn Events

		/// <summary>
		/// Occurs when the snapin starts
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void OnSnapInStart(object sender, EventArgs e)
		{
			this.StartMyServices();
		}

		/// <summary>
		/// Occurs when the snapin stops
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void OnSnapInStop(object sender, EventArgs e)
		{
			this.StopMyServices();
		}

		#endregion

		#region My Instance Manager Events

		/// <summary>
		/// Handles the event when a command line is received from another instance of ourself
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void OnCommandLineReceivedFromAnotherInstance(object sender, ApplicationInstanceManagerEventArgs e)
		{
			try
			{																
				// flash the window
				WindowFlasher.FlashWindow(this.Handle);
			}
			catch(System.Exception systemException)
			{
				System.Diagnostics.Trace.WriteLine(systemException);
			}
		}

		#endregion

		#region My Overrides

		/// <summary>
		/// Provides startup services
		/// </summary>
		protected override void StartMyServices()
		{
			base.StartMyServices ();

			try
			{
				// assimilate the proprties of the starting executable
				this.MorphAttributesToReflectStartingExecutable();
				
				// use the window positioning engine to manage our state
				WindowPositioningEngine.WindowPositioningEngineSnapIn.Instance.Manage(this, WindowPositioningEngineKey);
				
				// wire up to previous instance events
				SnapInHostingEngine.GetExecutingInstance().InstanceManager.CommandLineReceivedFromAnotherInstance += new ApplicationInstanceManagerEventHandler(OnCommandLineReceivedFromAnotherInstance);

				// add ourself as a top level window to the hosting engine's application context
				SnapInHostingEngine.Instance.ApplicationContext.AddTopLevelWindow(this);

				// show ourself
				this.Show();
			}
			catch(Exception ex)
			{
				Trace.WriteLine(ex);
			}
		}

		/// <summary>
		/// Provides shutdown services
		/// </summary>
		protected override void StopMyServices()
		{
			base.StopMyServices ();

			// wire up to previous instance events
			SnapInHostingEngine.GetExecutingInstance().InstanceManager.CommandLineReceivedFromAnotherInstance -= new ApplicationInstanceManagerEventHandler(OnCommandLineReceivedFromAnotherInstance);

			// close ourself
			this.Close();
		}

		#endregion

		#region My Public Properties

		/// <summary>
		/// Gets or sets whether this Window will cause the Application thread to exit when it is closed
		/// </summary>
		public bool ExitApplicationThreadOnClose
		{
			get
			{
				return _exitApplicationThreadOnClose;
			}
			set
			{
				_exitApplicationThreadOnClose = value;
			}
		}	

		/// <summary>
		/// Gets or sets a flag that determines if security prompts will be displayed when the app closes
		/// </summary>
		public bool NoPromptsOnClose
		{
			get
			{
				return _noPromptsOnClose;
			}
			set
			{
				_noPromptsOnClose = value;
			}
		}
		
		/// <summary>
		/// Returns the statusbar attached to the window
		/// </summary>
		public StatusBar StatusBar
		{
			get
			{				
				return _statusBar;
			}
		}
		
		/// <summary>
		/// Returns the main menu attached to the window
		/// </summary>
		public MainMenu MainMenu
		{
			get
			{
				return _mainMenu;
			}
		}

		/// <summary>
		/// Returns the File menu item attached to the main menu
		/// </summary>
		public MenuItem FileMenuItem
		{
			get
			{
				return _menuItemFile;
			}
		}
		
		/// <summary>
		/// Returns the Tools menu item attached to the main menu
		/// </summary>
		public MenuItem ToolsMenuItem
		{
			get
			{
				return _menuItemTools;
			}
		}
		
		/// <summary>
		/// Returns the Help menu item attached to the main menu
		/// </summary>
		public MenuItem HelpMenuItem
		{
			get
			{
				return _menuItemHelp;
			}
		}

		#endregion

		#region My Protected Methods

		/// <summary>
		/// Reflect upon the starting executable and snag it's properties
		/// </summary>
		protected void MorphAttributesToReflectStartingExecutable()
		{
			try
			{
				Assembly assembly = SnapInHostingEngine.GetExecutingInstance().StartingExecutable;
				if (assembly != null)
				{
					// snag the name of the file minus path and extention and set it as the heading
					string filename = System.IO.Path.GetFileName(assembly.Location);
					filename = filename.Replace(System.IO.Path.GetExtension(assembly.Location), null);							
					
					AssemblyAttributeReader reader = new AssemblyAttributeReader(assembly);
									
					// snag the company that made the assembly, and set it in the title
					System.Reflection.AssemblyCompanyAttribute[] companyAttributes = reader.GetAssemblyCompanyAttributes();
					if (companyAttributes != null)
						if (companyAttributes.Length > 0)
							if (companyAttributes[0].Company != null && companyAttributes[0].Company != string.Empty)
								this.Text = filename; // companyAttributes[0].Company + " " + filename;

					// snag the image from the assembly, it should be an executable so...
					Icon icon = ShellInformation.GetIconFromPath(assembly.Location,	IconSizes.SmallIconSize, IconStyles.NormalIconStyle, FileAttributes.Normal);	
					if (icon != null)
						this.Icon = icon;
				}
			}
			catch(System.Exception systemException)
			{
				System.Diagnostics.Trace.WriteLine(systemException);
			}
		}

		#endregion

		#region IWindowPositioningEngineFeaturable Members

		public Point GetDefaultLocation()
		{
			Rectangle rc = System.Windows.Forms.Screen.PrimaryScreen.WorkingArea;
			Size defSize = this.GetDefaultSize();
			int x = (rc.Width / 2) - (defSize.Width / 2);
			int y = (rc.Height / 2) - (defSize.Height / 2);
			return new Point(x, y);
		}

		public Size GetDefaultSize()
		{			
			return new Size (300, 300);
		}

		public System.Windows.Forms.FormWindowState GetDefaultWindowState()
		{			
			return FormWindowState.Normal;
		}

		#endregion

		#region My Menu Item Events

		/// <summary>
		/// Occurs when the File menu pops up
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void OnFileMenuItemPopup(object sender, EventArgs e)
		{
			MenuItem mi = (MenuItem)sender;
			mi.MenuItems.Clear();
			mi.MenuItems.Add("-");
			mi.MenuItems.Add("Exit", new EventHandler(this.OnExitClicked));
		}

		/// <summary>
		/// Occurs when the Tools menu pops up
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void OnToolsMenuItemPopup(object sender, EventArgs e)
		{
			MenuItem mi = (MenuItem)sender;
			mi.MenuItems.Clear();

			MenuItem exploreMenuItem = new MenuItem("Explore");
			exploreMenuItem.MenuItems.Add("All Users Path", new EventHandler(this.OnExploreAllUsersPathClicked));
			exploreMenuItem.MenuItems.Add("Current User Path", new EventHandler(this.OnExploreCurrentUsersPathClicked));
			
			mi.MenuItems.Add(exploreMenuItem);
			mi.MenuItems.Add("-");
			mi.MenuItems.Add("Features...", new EventHandler(this.OnFeaturesClicked));
			mi.MenuItems.Add("SnapIns...", new EventHandler(this.OnSnapInsClicked));
			mi.MenuItems.Add("-");
			mi.MenuItems.Add("Options", new EventHandler(this.OnOptionsClicked));
		}

		/// <summary>
		/// Occurs when the Help menu pops up
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void OnHelpMenuItemPopup(object sender, EventArgs e)
		{
			MenuItem mi = (MenuItem)sender;
			mi.MenuItems.Clear();

			mi.MenuItems.Add("About", new EventHandler(this.OnAboutClicked));
		}

		/// <summary>
		/// Occurs when the Exit menu item is clicked
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void OnExitClicked(object sender, EventArgs e)
		{
			this.Close();
		}

		/// <summary>
		/// Occurs when the Explorer All Users Path menu item is clicked
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void OnExploreAllUsersPathClicked(object sender, EventArgs e)
		{
			if (Directory.Exists(SnapInHostingEngine.CommonDataPath))
			{
				System.Diagnostics.Process p = new System.Diagnostics.Process();							
				p.StartInfo.UseShellExecute = true;
				p.StartInfo.Verb = "EXPLORE";
				p.StartInfo.FileName = SnapInHostingEngine.CommonDataPath;
				p.Start();
			}
			else
			{
				ExceptionEngine.DisplayException(this, "Directory does not exist", MessageBoxIcon.Information, MessageBoxButtons.OK, null, "The path '" + SnapInHostingEngine.CommonDataPath + "' could not be found.");
			}
		}

		/// <summary>
		/// Occurs when the Explore Current Users Path menu item is clicked
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void OnExploreCurrentUsersPathClicked(object sender, EventArgs e)
		{
			if (Directory.Exists(SnapInHostingEngine.LocalUserDataPath))
			{
				System.Diagnostics.Process p = new System.Diagnostics.Process();
				p.StartInfo.UseShellExecute = true;
				p.StartInfo.Verb = "EXPLORE";
				p.StartInfo.FileName = SnapInHostingEngine.LocalUserDataPath;
				p.Start();
			}
			else
			{
				ExceptionEngine.DisplayException(this, "Directory does not exist", MessageBoxIcon.Information, MessageBoxButtons.OK, null, "The path '" + SnapInHostingEngine.LocalUserDataPath + "' could not be found.");
			}
		}

		/// <summary>
		/// Occurs when the Features menu item is clicked
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void OnFeaturesClicked(object sender, EventArgs e)
		{			
			FeatureEngine.ShowFeatureWindow(this, SnapInHostingEngine.GetExecutingInstance());
		}

		/// <summary>
		/// Occurs when the SnapIns menu item is clicked
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void OnSnapInsClicked(object sender, EventArgs e)
		{
			SnapInHostingEngine.GetExecutingInstance().ShowSnapInsWindow(this);
		}

		/// <summary>
		/// Occurs when the Options menu item is clicked
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void OnOptionsClicked(object sender, EventArgs e)
		{
			SnapInHostingEngine.GetExecutingInstance().ShowConfigurationWindow(this);
		}

		/// <summary>
		/// Occurs when the About menu item is clicked
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void OnAboutClicked(object sender, EventArgs e)
		{
//			SplashWindow window = new SplashWindow(SnapInHostingEngine.Instance.StartingExecutable, true);
//			window.ShowDialog(this);
		}

		#endregion

		/// <summary>
		/// Intercepts the event generated when the splash window closes
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void OnSplashWindowClosed(object sender, EventArgs e)
		{
			// we gotta get focus back to this window
			this.BringToFront();
		}
	}
}
