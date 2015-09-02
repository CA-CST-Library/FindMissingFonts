using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Drawing.Text;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace FindMissingFonts
{
	public partial class frmMain : Form
	{
		private AutoResetEvent _resetEvent = new AutoResetEvent( false );

		public frmMain()
		{
			InitializeComponent();
		}

		private void frmMain_Load( object sender, EventArgs e )
		{
			// set the initial search directory root to wherever the applicatino is running from
			tbSearch.Text = Environment.CurrentDirectory;
		}

		private void btnBrowse_Click( object sender, EventArgs e )
		{
			dlgFolderBrowser.SelectedPath = tbSearch.Text;
			if( dlgFolderBrowser.ShowDialog() == System.Windows.Forms.DialogResult.OK )
			{
				tbSearch.Text = dlgFolderBrowser.SelectedPath;
			}
		}

		private void btnSearch_Click( object sender, EventArgs e )
		{
			if( backgroundWorker.IsBusy != true )
			{
				btnSearch.Enabled = false;
				_resetEvent.Reset();
				backgroundWorker.RunWorkerAsync( tbSearch.Text );
			}
		}

		/// <summary>
		/// Return a HashSet containing all the fonts referenced in the supplied files
		/// </summary>
		/// <param name="files">collection of files to gather fonts from</param>
		/// <returns></returns>
		private HashSet<string> BuildFontSet( IEnumerable<string> files, BackgroundWorker worker )
		{
			HashSet<string> fonts = new HashSet<string>();

			foreach( var file in files )
			{
				if( worker.CancellationPending )
				{
					return null;
				}
				ShowMessage( file );
				// create a Document object to contain the document files information
				using( Document doc = new Document() )
				{
					doc.OnError = ShowError;
					doc.OnMessage = ShowMessage;
					// load the document file into the object
					if( doc.Load( file ) )
					{
						try
						{
							// get collection of fonts used in the document
							List<string> fontsInDoc = doc.Fonts;
							// add them to the HashSet
							fontsInDoc.ForEach( fontName => fonts.Add( fontName ) );
						}
						catch( FileFormatException ffex )
						{
							ShowError( String.Format( "File: {0}  Error: {1}", file, ffex.Message ) );
						}
					}
				}
			}
			return fonts;
		}

		/// <summary>
		/// Show a message box displaying the error
		/// </summary>
		/// <param name="message">error message to display</param>
		public void ShowError( string message )
		{
			//MessageBox.Show( message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error );
		}

		/// <summary>
		/// Show a message in the toolstrip status label
		/// </summary>
		/// <param name="message">message to display</param>
		public void ShowMessage( string message )
		{
			if( statusStrip.InvokeRequired )
			{
				Action<string> del = ShowMessage;
				this.Invoke( del, new object[] { message } );
			}
			else
			{
				toolStripStatusLabel.Text = message;
				statusStrip.Refresh();
			}
		}

		private void btnCancel_Click( object sender, EventArgs e )
		{
			if( backgroundWorker.WorkerSupportsCancellation && backgroundWorker.IsBusy )
			{
				backgroundWorker.CancelAsync();
			}
		}

		/// <summary>
		/// Return a HashSet containing the installed fonts in Windows
		/// </summary>
		/// <returns>HashSet of the installed Windows fonts</returns>
		private HashSet<string> GetInstalledFonts()
		{
			HashSet<string> fonts = new HashSet<string>();
			InstalledFontCollection installedFontCollection = new InstalledFontCollection();
			foreach( FontFamily fontFamily in installedFontCollection.Families )
			{
				fonts.Add( fontFamily.Name );
			}

			return fonts;
		}

		/// <summary>
		/// Copy the contents of the ListBox to the Windows Clipboard as text
		/// </summary>
		/// <param name="sender">ignored</param>
		/// <param name="e">ignored</param>
		private void btnSendToClipboard_Click( object sender, EventArgs e )
		{
			StringBuilder sb = new StringBuilder();
			foreach( string font in lbFonts.Items )
			{
				sb.AppendLine( font );
			}
			Clipboard.SetText( sb.ToString(), TextDataFormat.Text );
		}

		private void ScanFiles( object sender, DoWorkEventArgs e )
		{
			string baseDir = (string)e.Argument;
			BackgroundWorker worker = sender as BackgroundWorker;
			HashSet<string> installedFonts = null;
			HashSet<string> fonts = null;
			// get all the WinWord document files
			var files = from file in
							Directory.EnumerateFiles( baseDir, "*.doc?", SearchOption.AllDirectories )
						where file.Contains( "\\~" ) == false
						select ( file );

			// get all the fonts referenced in the documents
			fonts = BuildFontSet( files, worker );
			if( !worker.CancellationPending )
			{
				// get the fonts installed in the Windows fonts folder
				ShowMessage( "Getting installed fonts..." );
				installedFonts = GetInstalledFonts();
			}
			if( worker.CancellationPending )
			{
				e.Cancel = true;
				return;
			}
			// get only those referenced in the documents that are not in the fonts folder
			ShowMessage( "Filtering..." );
			System.Collections.Generic.IEnumerable<string> missing = fonts.Except( installedFonts );
			e.Result = missing;
			_resetEvent.Set();
		}

		private void frmMain_FormClosing( object sender, FormClosingEventArgs e )
		{
			if( backgroundWorker.WorkerSupportsCancellation && backgroundWorker.IsBusy )
			{
				backgroundWorker.CancelAsync();
				_resetEvent.WaitOne();
			}
		}

		private void ScanFilesCompleted( object sender, RunWorkerCompletedEventArgs e )
		{
			if( e.Cancelled )
			{
				ShowMessage( "Cancelled!" );
				return;
			}
			if( e.Error != null )
			{
				ShowMessage( "Fatal error!" );
				MessageBox.Show( e.Error.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error );
				return;
			}
			// show them in the listbox
			ShowMessage( "Filling list..." );
			lbFonts.Items.Clear();
			System.Collections.Generic.IEnumerable<string> missing = (System.Collections.Generic.IEnumerable<string>)e.Result;
			foreach( String fontName in missing )
			{
				lbFonts.Items.Add( fontName );
			}
			ShowMessage( String.Empty );

			btnSearch.Enabled = true;
		}
	}
}