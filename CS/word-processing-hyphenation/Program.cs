using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace word_processing_hyphenation
{
    class Program
    {
        static void Main(string[] args)
        {
            using (RichEditDocumentServer wordProcessor = new RichEditDocumentServer())
            {
                //Register the created service implementation
                wordProcessor.LoadDocument("Multimodal.docx");

                //Load embedded dictionaries
                var openOfficePatternStream = Assembly.GetExecutingAssembly().GetManifestResourceStream("word_processing_hyphenation.hyphen.dic");
                var customDictionaryStream = Assembly.GetExecutingAssembly().GetManifestResourceStream("word_processing_hyphenation.hyphen_exc.dic");

                //Create dictionary objects
                OpenOfficeHyphenationDictionary hyphenationDictionary = new OpenOfficeHyphenationDictionary(openOfficePatternStream, new System.Globalization.CultureInfo("EN-US"));
                CustomHyphenationDictionary exceptionsDictionary = new CustomHyphenationDictionary(customDictionaryStream, new System.Globalization.CultureInfo("EN-US"));

                //Add them to the word processor's collection
                wordProcessor.HyphenationDictionaries.Add(hyphenationDictionary);
                wordProcessor.HyphenationDictionaries.Add(exceptionsDictionary);

                //Specify hyphenation settings
                wordProcessor.Document.Hyphenation = true;
                wordProcessor.Document.HyphenateCaps = true;

                //Export the result to the PDF format
                wordProcessor.ExportToPdf("Result.pdf");

            }
            //Open the result
            Process.Start("Result.pdf");

        }
    }
}
