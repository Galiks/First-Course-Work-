using NLog;
using Spire.Doc;
using Spire.Pdf;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace BusinessLogicLayer.Conversion
{
    /// <summary>
    /// 
    /// </summary>
    public static class Conversion
    {
        private static readonly Logger loggerException = LogManager.GetLogger("exception");
        private static readonly Logger loggerUser = LogManager.GetLogger("user");
        private const string formatWord = "formating";


        public static string ConvertToWordDoc(string filepath)
        {
            string extension = Path.GetExtension(filepath);
            if (extension.Equals(".doc"))
            {
                return filepath;
            }
            else if (extension.Equals(".pdf"))
            {
                return GetFilepathAfterConvertSpirePdf(filepath, Spire.Pdf.FileFormat.DOC);
            }
            else
            {
                return GetFilepathAfterConvertSpireDoc(filepath, Spire.Doc.FileFormat.Doc);
            }
        }

        public static string ConvertToWordDocx(string filepath)
        {
            string extension = Path.GetExtension(filepath);
            if (extension.Equals(".docx"))
            {
                return filepath;
            }
            else if (extension.Equals(".pdf"))
            {
                return GetFilepathAfterConvertSpirePdf(filepath, Spire.Pdf.FileFormat.DOCX);
            }
            else
            {
                return GetFilepathAfterConvertSpireDoc(filepath, Spire.Doc.FileFormat.Docx);
            }
        }

        public static string ConvertToHtml(string filepath)
        {
            string extension = Path.GetExtension(filepath);
            if (extension.Equals(".html"))
            {
                return filepath;
            }
            else if (extension.Equals(".pdf"))
            {
                return GetFilepathAfterConvertSpirePdf(filepath, Spire.Pdf.FileFormat.HTML);
            }
            else
            {
                return GetFilepathAfterConvertSpireDoc(filepath, Spire.Doc.FileFormat.Html);
            }
        }

        public static string ConvertToOdt(string filepath)
        {
            string extension = Path.GetExtension(filepath);
            if (extension.Equals(".odt"))
            {
                return filepath;
            }
            else if (extension.Equals(".pdf"))
            {
                return null;
            }         
            else
            {
                return GetFilepathAfterConvertSpireDoc(filepath, Spire.Doc.FileFormat.Odt);
            }
        }

        public static string ConvertToPdf(string filepath)
        {
            string extension = Path.GetExtension(filepath);
            if (extension.Equals(".pdf"))
            {
                return filepath;
            }
            else
            {
                return GetFilepathAfterConvertSpireDoc(filepath, Spire.Doc.FileFormat.PDF);
            }
        }

        public static string ConvertToRtf(string filepath)
        {
            string extension = Path.GetExtension(filepath);
            if (extension.Equals(".rtf"))
            {
                return filepath;
            }
            else if (extension.Equals(".pdf"))
            {
                return null;
            }
            else
            {
                return GetFilepathAfterConvertSpireDoc(filepath, Spire.Doc.FileFormat.Rtf);
            }
        }

        public static string ConvertToTxt(string filepath)
        {
            string extension = Path.GetExtension(filepath);
            if (extension.Equals(".txt"))
            {
                return filepath;
            }
            else if (extension.Equals(".pdf"))
            {
                return null;
            }
            else
            {
                return GetFilepathAfterConvertSpireDoc(filepath, Spire.Doc.FileFormat.Txt);
            }
        }
        //Подумать над названием
        private static string GetFilepathAfterConvertSpireDoc(string filepath, Spire.Doc.FileFormat fileFormat)
        {
            Document document = new Document();
            document.LoadFromFile(filepath);
            filepath = $"{filepath}_{formatWord}.{fileFormat.ToString().ToLower()}";
            document.SaveToFile(filepath, fileFormat);
            return filepath;
        }

        private static string GetFilepathAfterConvertSpirePdf(string filepath, Spire.Pdf.FileFormat fileFormat)
        {
            PdfDocument pdfDocument = new PdfDocument();
            pdfDocument.LoadFromFile(filepath);
            filepath = $"{filepath}_{formatWord}.{fileFormat.ToString().ToLower()}";
            pdfDocument.SaveToFile(filepath, fileFormat);
            return filepath;
        }
    }
}
