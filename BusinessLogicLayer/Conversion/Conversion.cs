using NLog;
using Spire.Doc;
using Spire.Pdf;
using System.IO;

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

        /// <summary>
        /// 
        /// </summary>
        /// <param name="filepath"></param>
        /// <returns></returns>
        public static string ConvertToWordDoc(string filepath)
        {
            string extension = Path.GetExtension(filepath);
            if (extension.Equals(".doc"))
            {
                loggerUser.Info("Конвертация не нужна.");
                return filepath;
            }
            else if (extension.Equals(".pdf"))
            {
                loggerUser.Info("Переконвертация в формат DOC.");
                return GetFilepathAfterConvertSpirePdf(filepath, Spire.Pdf.FileFormat.DOC);
            }
            else
            {
                loggerUser.Info("Переконвертация в формат DOC.");
                return GetFilepathAfterConvertSpireDoc(filepath, Spire.Doc.FileFormat.Doc);
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="filepath"></param>
        /// <returns></returns>
        public static string ConvertToWordDocx(string filepath)
        {
            string extension = Path.GetExtension(filepath);
            if (extension.Equals(".docx"))
            {
                loggerUser.Info("Конвертация не нужна.");
                return filepath;
            }
            else if (extension.Equals(".pdf"))
            {
                loggerUser.Info("Переконвертация в формат DOCX.");
                return GetFilepathAfterConvertSpirePdf(filepath, Spire.Pdf.FileFormat.DOCX);
            }
            else
            {
                loggerUser.Info("Переконвертация в формат DOCX.");
                return GetFilepathAfterConvertSpireDoc(filepath, Spire.Doc.FileFormat.Docx);
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="filepath"></param>
        /// <returns></returns>
        public static string ConvertToHtml(string filepath)
        {
            string extension = Path.GetExtension(filepath);
            if (extension.Equals(".html"))
            {
                loggerUser.Info("Конвертация не нужна.");
                return filepath;
            }
            else if (extension.Equals(".pdf"))
            {
                loggerUser.Info("Переконвертация в формат HTML.");
                return GetFilepathAfterConvertSpirePdf(filepath, Spire.Pdf.FileFormat.HTML);
            }
            else
            {
                loggerUser.Info("Переконвертация в формат HTML.");
                return GetFilepathAfterConvertSpireDoc(filepath, Spire.Doc.FileFormat.Html);
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="filepath"></param>
        /// <returns></returns>
        public static string ConvertToOdt(string filepath)
        {
            string extension = Path.GetExtension(filepath);
            if (extension.Equals(".odt"))
            {
                loggerUser.Info("Конвертация не нужна.");
                return filepath;
            }
            else if (extension.Equals(".pdf"))
            {
                loggerUser.Info("Конвертация невозможна.");
                return null;
            }         
            else
            {
                loggerUser.Info("Переконвертация в формат ODT.");
                return GetFilepathAfterConvertSpireDoc(filepath, Spire.Doc.FileFormat.Odt);
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="filepath"></param>
        /// <returns></returns>
        public static string ConvertToPdf(string filepath)
        {
            string extension = Path.GetExtension(filepath);
            if (extension.Equals(".pdf"))
            {
                loggerUser.Info("Конвертация не нужна.");
                return filepath;
            }
            else
            {
                loggerUser.Info("Переконвертация в формат PDF.");
                return GetFilepathAfterConvertSpireDoc(filepath, Spire.Doc.FileFormat.PDF);
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="filepath"></param>
        /// <returns></returns>
        public static string ConvertToRtf(string filepath)
        {
            string extension = Path.GetExtension(filepath);
            if (extension.Equals(".rtf"))
            {
                loggerUser.Info("Конвертация не нужна.");
                return filepath;
            }
            else if (extension.Equals(".pdf"))
            {
                loggerUser.Info("Конвертация невозможна.");
                return null;
            }
            else
            {
                loggerUser.Info("Переконвертация в формат RTF.");
                return GetFilepathAfterConvertSpireDoc(filepath, Spire.Doc.FileFormat.Rtf);
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="filepath"></param>
        /// <returns></returns>
        public static string ConvertToTxt(string filepath)
        {
            string extension = Path.GetExtension(filepath);
            if (extension.Equals(".txt"))
            {
                loggerUser.Info("Конвертация не нужна.");
                return filepath;
            }
            else if (extension.Equals(".pdf"))
            {
                loggerUser.Info("Конвертация невозможна.");
                return null;
            }
            else
            {
                loggerUser.Info("Переконвертация в формат TXT.");
                return GetFilepathAfterConvertSpireDoc(filepath, Spire.Doc.FileFormat.Txt);
            }
        }
        //Подумать над названием
        private static string GetFilepathAfterConvertSpireDoc(string filepath, Spire.Doc.FileFormat fileFormat)
        {
            Document document = new Document();
            document.LoadFromFile(filepath);

            if (fileFormat == Spire.Doc.FileFormat.Html)
            {
                filepath = FolderWork.CreateForlderForHTML(filepath);
                filepath = $"{filepath}\\{formatWord}.{fileFormat.ToString().ToLower()}"; 
            }
            else
            {
                filepath = $"{filepath}_{formatWord}.{fileFormat.ToString().ToLower()}";
            }
            document.SaveToFile(filepath, fileFormat);
            return filepath;
        }

        private static string GetFilepathAfterConvertSpirePdf(string filepath, Spire.Pdf.FileFormat fileFormat)
        {
            PdfDocument pdfDocument = new PdfDocument();
            pdfDocument.LoadFromFile(filepath);
            if (fileFormat == Spire.Pdf.FileFormat.HTML)
            {
                filepath = FolderWork.CreateForlderForHTML(filepath);
                filepath = $"{filepath}\\html\\{formatWord}.{fileFormat.ToString().ToLower()}";
            }
            else
            {
                filepath = $"{filepath}_{formatWord}.{fileFormat.ToString().ToLower()}";
            }
            pdfDocument.SaveToFile(filepath, fileFormat);
            return filepath;
        }
    }
}
