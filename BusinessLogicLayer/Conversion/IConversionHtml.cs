using System;
using System.Collections.Generic;
using System.Text;

namespace BusinessLogicLayer.Conversion
{
    /// <summary>
    /// 
    /// </summary>
    public interface IConversionHtml
    {
        /// <summary>
        /// Конвертирует файл HTML в Word
        /// </summary>
        /// <param name="filepath">путь к файлу HTML</param>
        /// <returns>путь к файлу Word</returns>
        string ConvertHtmlToWord(string filepath);
    }
}
