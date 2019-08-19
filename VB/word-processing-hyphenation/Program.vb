Imports DevExpress.XtraRichEdit
Imports DevExpress.XtraRichEdit.API.Native
Imports System
Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.Linq
Imports System.Reflection
Imports System.Text
Imports System.Threading.Tasks

Namespace word_processing_hyphenation
	Friend Class Program
		Shared Sub Main(ByVal args() As String)
			Using wordProcessor As New RichEditDocumentServer()
				'Register the created service implementation
				wordProcessor.LoadDocument("Multimodal.docx")

				'Load embedded dictionaries
				Dim openOfficePatternStream = System.Reflection.Assembly.GetExecutingAssembly().GetManifestResourceStream("hyphen.dic")
				Dim customDictionaryStream = System.Reflection.Assembly.GetExecutingAssembly().GetManifestResourceStream("hyphen_exc.dic")

				'Create dictionary objects
				Dim hyphenationDictionary As New OpenOfficeHyphenationDictionary(openOfficePatternStream, New System.Globalization.CultureInfo("EN-US"))
				Dim exceptionsDictionary As New CustomHyphenationDictionary(customDictionaryStream, New System.Globalization.CultureInfo("EN-US"))

				'Add them to the word processor's collection
				wordProcessor.HyphenationDictionaries.Add(hyphenationDictionary)
				wordProcessor.HyphenationDictionaries.Add(exceptionsDictionary)

				'Specify hyphenation settings
				wordProcessor.Document.Hyphenation = True
				wordProcessor.Document.HyphenateCaps = True

				'Export the result to the PDF format
				wordProcessor.ExportToPdf("Result.pdf")

			End Using
			'Open the result
			Process.Start("Result.pdf")

		End Sub
	End Class
End Namespace
