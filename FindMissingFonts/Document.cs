using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace FindMissingFonts
{
	public class Document : IDisposable
	{
		public Action<string> OnError = null;
		public Action<string> OnMessage = null;

		private WordprocessingDocument wordprocessingDocument = null;
		private bool disposed = false;

		/// <summary>
		/// Load the document file into the Open XML WordprocessingDocument object
		/// </summary>
		/// <param name="path">Path to the file to load</param>
		/// <returns></returns>
		public bool Load( string path )
		{
			Exception exHolder = null;
			try
			{
				wordprocessingDocument = WordprocessingDocument.Open( path, true );
			}
			catch( Exception ex )
			{
				wordprocessingDocument = null;
				exHolder = ex;
			}
			if( wordprocessingDocument == null )
			{
				Word.Application wordApp = null;
				try
				{
					wordApp = new Word.Application();
					wordApp.Visible = true;
					Word.Document wordDoc = wordApp.Documents.Open( path );
					object filename = "c:\\temp\\tempfile.docx";
					object saveFormat = Word.WdSaveFormat.wdFormatDocumentDefault;
					wordDoc.SaveAs2( FileName: ref filename, FileFormat: ref saveFormat );
					object doNotSaveChanges = Word.WdSaveOptions.wdDoNotSaveChanges;
					wordDoc.Close( SaveChanges: ref doNotSaveChanges );
					wordDoc = null;
					wordApp.Quit( SaveChanges: Word.WdSaveOptions.wdDoNotSaveChanges );
					wordApp = null;

					wordprocessingDocument = WordprocessingDocument.Open( "c:\\temp\\tempfile.docx", true );
				}
				catch( Exception ex )
				{
					if( wordApp != null )
					{
						wordApp.Quit( SaveChanges: Word.WdSaveOptions.wdDoNotSaveChanges );
						wordApp = null;
					}
					if( exHolder != null )
					{
						ShowError( String.Format( "File: {0}  Error: {1}", path, exHolder.Message ) );
					}
					else
					{
						ShowError( String.Format( "File: {0}  Error: {1}", path, ex.Message ) );
					}
					return false;
				}
				finally
				{
					if( exHolder != null )
					{
						try
						{
							File.Delete( "c:\\temp\\tempfile.docx" );
						}
						catch( Exception )
						{
						}
					}
				}
			}
			return true;
		}

		/// <summary>
		/// Return a list of fonts referenced in the document
		/// </summary>
		public List<string> Fonts
		{
			get
			{
				// get all fonts of the word document 
				List<string> fonts = wordprocessingDocument
					.MainDocumentPart
					.Document
					.Descendants<RunFonts>()
					.Select( c => ( c.Ascii != null && c.Ascii.HasValue ) ? c.Ascii.InnerText : string.Empty )
					.Distinct().ToList();
				fonts.Sort();
				if( fonts.Count > 0 && fonts[0].Length == 0 )
				{
					fonts.RemoveAt( 0 );
				}
				return fonts;
			}
		}

		/// <summary>
		/// Callback to handle an error message
		/// </summary>
		/// <param name="message"></param>
		private void ShowError( string message )
		{
			if( OnError != null )
			{
				OnError( message );
			}
		}

		/// <summary>
		/// Callback to handle an informational message
		/// </summary>
		/// <param name="message"></param>
		private void ShowMessage( string message )
		{
			if( OnMessage != null )
			{
				OnMessage( message );
			}
		}

		public void Dispose()
		{
			Dispose( true );
			GC.SuppressFinalize( this );
		}

		protected virtual void Dispose( bool disposing )
		{
			if( disposed )
				return;

			if( disposing )
			{
				if( wordprocessingDocument != null )
				{
					wordprocessingDocument.Dispose();
				}
			}
			disposed = true;
		}
	}
}
