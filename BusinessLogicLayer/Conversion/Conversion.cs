using Spire.Doc;
using Spire.Pdf;
using System;
using System.Collections.Generic;
using System.Text;

namespace BusinessLogicLayer.Conversion
{
    /// <summary>
    /// 
    /// </summary>
    public class Conversion : IConversionWord, IConversionPdf, IConversionHtml
    {
        private const string formatWord = "formating";
        public string ConvertHtmlToWord(string filepath)
        {
            throw new NotImplementedException();
        }

        public string ConvertPdfToHtml(string filepath)
        {
            throw new NotImplementedException();
        }

        public string ConvertPdfToOdt(string filepath)
        {
            throw new NotImplementedException();
        }

        public string ConvertPdfToRtf(string filepath)
        {
            throw new NotImplementedException();
        }

        public string ConvertPdfToTxt(string filepath)
        {
            throw new NotImplementedException();
        }

        public string ConvertPdfToWord(string filepath)
        {
            PdfDocument pdfDoc = new PdfDocument();
            pdfDoc.LoadFromFile(filepath);
            filepath = $"{filepath}_{formatWord}.doc";
            pdfDoc.SaveToFile(filepath, Spire.Pdf.FileFormat.DOC);
            return filepath;
        }

        public string ConvertWordToHtml(string filepath)
        {
            throw new NotImplementedException();
        }

        public string ConvertWordToOdt(string filepath)
        {
            throw new NotImplementedException();
        }

        public string ConvertWordToPdf(string filepath)
        {
            throw new NotImplementedException();
        }

        public string ConvertWordToRtf(string filepath)
        {
            throw new NotImplementedException();
        }

        public string ConvertWordToTxt(string filepath)
        {
            Document document = new Document();
            document.LoadText(filepath);
            filepath = $"{filepath}_{formatWord}.txt";
            document.SaveToFile(filepath, Spire.Doc.FileFormat.Txt);
            return filepath;
        }
    }
}
