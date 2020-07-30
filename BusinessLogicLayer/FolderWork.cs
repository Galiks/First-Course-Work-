using System.IO;

namespace BusinessLogicLayer
{
    /// <summary>
    /// 
    /// </summary>
    public class FolderWork
    {
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
    }
}
