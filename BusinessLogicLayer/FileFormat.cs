using System;
using System.Collections.Generic;
using System.Text;

namespace BusinessLogicLayer
{
    /// <summary>
    /// 
    /// </summary>
    public enum FileFormat: int
    {
        /// <summary>
        /// Формат файла TXT
        /// </summary>
        TXT = 0,
        /// <summary>
        /// Формат файла RTF
        /// </summary>
        RTF = 1,
        /// <summary>
        /// Формат файла DOC
        /// </summary>
        DOC = 2,
        /// <summary>
        /// Формат файла DOCX
        /// </summary>
        DOCX = 3,
        /// <summary>
        /// Формат файла HTML
        /// </summary>
        HTML = 4,
        /// <summary>
        /// Формат файла PDF
        /// </summary>
        PDF = 5,
        /// <summary>
        /// Формат файла ODT
        /// </summary>
        ODT = 6
    }
}
