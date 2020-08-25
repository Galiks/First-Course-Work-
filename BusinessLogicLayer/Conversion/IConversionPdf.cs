using System;
using System.Collections.Generic;
using System.Text;

namespace BusinessLogicLayer.Conversion
{
    /// <summary>
    /// 
    /// </summary>
    public interface IConversionPdf
    {
        /// <summary>
        /// Конвертирует файл PDF в TXT
        /// </summary>
        /// <param name="filepath">путь к файлу PDF</param>
        /// <returns>путь к файлу TXT</returns>
        string ConvertPdfToTxt(string filepath);
        /// <summary>
        /// Конвертирует файл PDF в RTF
        /// </summary>
        /// <param name="filepath">путь к файлу PDF</param>
        /// <returns>путь к файлу RTF</returns>
        string ConvertPdfToRtf(string filepath);
        /// <summary>
        /// Конвертирует файл PDF в HTML
        /// </summary>
        /// <param name="filepath">путь к файлу PDF</param>
        /// <returns>путь к файлу HTML</returns>
        string ConvertPdfToHtml(string filepath);
        /// <summary>
        /// Конвертирует файл PDF в Word
        /// </summary>
        /// <param name="filepath">путь к файлу PDF</param>
        /// <returns>путь к файлу Word</returns>
        string ConvertPdfToWord(string filepath);
        /// <summary>
        /// Конвертирует файл PDF в ODT
        /// </summary>
        /// <param name="filepath">путь к файлу PDF</param>
        /// <returns>путь к файлу ODT</returns>
        string ConvertPdfToOdt(string filepath);
    }
}
