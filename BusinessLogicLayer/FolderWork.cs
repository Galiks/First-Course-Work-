using Microsoft.AspNetCore.Http;
using NLog;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace BusinessLogicLayer
{
    /// <summary>
    /// 
    /// </summary>
    public class FolderWork
    {
        private static readonly Logger loggerException = LogManager.GetLogger("exception");
        private static readonly Logger loggerUser = LogManager.GetLogger("user");


        /// <summary>
        /// 
        /// </summary>
        /// <param name="firstname"></param>
        /// <param name="lastname"></param>
        /// <param name="patronymic"></param>
        /// <param name="path"></param>
        public static string CreateFolder(string firstname, string lastname, string patronymic, string path)
        {
            Directory.SetCurrentDirectory(path);

            return Directory.CreateDirectory(GetFolderName(firstname, lastname, patronymic)).FullName;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="firstname"></param>
        /// <param name="lastname"></param>
        /// <param name="patronymic"></param>
        /// <returns></returns>
        public static string GetFolderName(string firstname, string lastname, string patronymic)
        {
            return $"{firstname[0]}{patronymic[0]}{lastname}";
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="file"></param>
        /// <param name="filename"></param>
        /// <param name="filepath"></param>
        /// <param name="userFolder"></param>
        public static void SaveFile(IFormFile file, string filename, ref string filepath, ref string userFolder)
        {
            if (System.IO.File.Exists(filepath))
            {
                filename = "NEW_" + filename;
                filepath = userFolder + filename;
                SaveFile(file, filename, ref filepath, ref userFolder);
            }
            else
            {
                using var fileStream = new FileStream(filepath, FileMode.CreateNew);
                file.CopyTo(fileStream);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="image"></param>
        /// <param name="path"></param>
        /// <returns></returns>
        public static async Task SaveImage(IFormFile image, string path)
        {
            var fileStream = new FileStream(path, FileMode.Create);
            await image.CopyToAsync(fileStream);
            loggerUser.Info($"Загружено изображение {image.FileName}");
            fileStream.Close();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="fileFormat"></param>
        /// <param name="userFolder"></param>
        /// <param name="filepath"></param>
        /// <param name="fileBytes"></param>
        /// <param name="fileExtension"></param>
        /// <param name="filename"></param>
        public static void SetParametresForDownloadFile(FileFormat fileFormat, string userFolder, string filepath, out byte[] fileBytes, out string fileExtension, out string filename)
        {
            string newFilepath = new DirectoryInfo(userFolder).GetFiles().Where(f => f.FullName == filepath).FirstOrDefault().FullName;
            switch (fileFormat)
            {
                case FileFormat.TXT:
                    newFilepath = Conversion.Conversion.ConvertToTxt(newFilepath); //Conversion.ConvertToTxt(newFilepath);
                    break;
                case FileFormat.RTF:
                    newFilepath = Conversion.Conversion.ConvertToRtf(newFilepath);
                    break;
                case FileFormat.DOC:
                    newFilepath = Conversion.Conversion.ConvertToWordDoc(newFilepath);
                    break;
                case FileFormat.DOCX:
                    newFilepath = Conversion.Conversion.ConvertToWordDocx(newFilepath);
                    break;
                case FileFormat.HTML:
                    newFilepath = Conversion.Conversion.ConvertToHtml(newFilepath);
                    break;
                case FileFormat.PDF:
                    newFilepath = Conversion.Conversion.ConvertToPdf(newFilepath);
                    break;
                case FileFormat.ODT:
                    newFilepath = Conversion.Conversion.ConvertToOdt(newFilepath);
                    break;
                default:
                    break;
            }


            FileInfo file = new DirectoryInfo(userFolder).GetFiles().Where(f => f.FullName == newFilepath).FirstOrDefault();
            fileBytes = System.IO.File.ReadAllBytes(file.FullName);
            fileExtension = file.Extension;
            filename = file.Name;
        }
    }
}
