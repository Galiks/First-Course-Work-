using System;
using System.Collections.Generic;
using System.Text;

namespace BusinessLogicLayer.Conversion
{
    /// <summary>
    /// 
    /// </summary>
    public interface IConversionWord
    {
        /// <summary>
        /// Конвертирует файл Word в TXT
        /// </summary>
        /// <param name="filepath">путь к файлу Word</param>
        /// <returns>путь к файлу TXT</returns>
        string ConvertWordToTxt(string filepath);
        /// <summary>
        /// Конвертирует файл Word в RTF
        /// </summary>
        /// <param name="filepath">путь к файлу Word</param>
        /// <returns>путь к файлу RTF</returns>
        string ConvertWordToRtf(string filepath);
        /// <summary>
        /// Конвертирует файл Word в HTML
        /// </summary>
        /// <param name="filepath">путь к файлу Word</param>
        /// <returns>путь к файлу HTML</returns>
        string ConvertWordToHtml(string filepath);
        /// <summary>
        /// Конвертирует файл Word в PDF
        /// </summary>
        /// <param name="filepath">путь к файлу Word</param>
        /// <returns>путь к файлу PDF</returns>
        string ConvertWordToPdf(string filepath);
        /// <summary>
        /// Конвертирует файл Word в ODT
        /// </summary>
        /// <param name="filepath">путь к файлу Word</param>
        /// <returns>путь к файлу ODT</returns>
        string ConvertWordToOdt(string filepath);
    }
}
