using Spire.Doc;
using Spire.Doc.Documents;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;

namespace BusinessLogicLayer
{
    public class WordDocument
    {
        private const string text = " – это вид текста, который содержит в себе информацию или ссылку на место в произведении, на другую литературу или на какое-то событие в мире.";
        private const string longText = "Гипертекст в тексте может выглядеть как гипертекст, а может как гипертекст на гипертексте, где гипертекст сам является гипертекстом.";
        private readonly List<string> wordCase;
        private readonly Document document;
        private readonly Section section;
        private readonly Paragraph mainParagraph;

        public WordDocument()
        {
            wordCase = new List<string> {"", "ы", "и", "а", "я", "у", "е", "ю", "о", "ой", "ою", "ей", "ею", "ом", "ем", "ью" };
            document = new Document();
            section = document.AddSection();
            mainParagraph = section.AddParagraph();
        }
        public void CreateDocument()
        {
            Paragraph paragraph = section.AddParagraph();

            //Set Paragraph Styles
            ParagraphStyle txtStyle = new ParagraphStyle(document);
            txtStyle.Name = "Style";
            txtStyle.CharacterFormat.FontName = "Impact";
            txtStyle.CharacterFormat.FontSize = 16;
            txtStyle.CharacterFormat.TextColor = Color.RosyBrown;
            document.Styles.Add(txtStyle);
            //Set Hyperlink Styles
            ParagraphStyle hyperlinkstyle = new ParagraphStyle(document);
            hyperlinkstyle.Name = "linkStyle";
            hyperlinkstyle.CharacterFormat.FontName = "Calibri";
            hyperlinkstyle.CharacterFormat.FontSize = 15;
            document.Styles.Add(hyperlinkstyle);

            ///главный параграф
            mainParagraph.AppendText(longText);
            ReplaceWords("гипертекст", "!!!");








            ////////заготовка на гипертекст
            //paragraph1.AppendHyperlink("Гипертекст", "Гипертекст", HyperlinkType.Bookmark);
            //paragraph1.ApplyStyle(hyperlinkstyle.Name);

            //for (int i = 0; i < 10; i++)
            //{
            //    paragraph = section.AddParagraph();
            //    paragraph.AppendHyperlink(paragraph1.Text, "Гипертекст", HyperlinkType.Bookmark);
            //    paragraph.AppendText(" поможет людям лучше понять текст.");
            //    Paragraph subParagraph = section.AddParagraph();
            //    subParagraph.AppendText("Слово употребляется в следующих параграфах: ");
            //    subParagraph.AppendHyperlink(paragraph1.ToString(), paragraph1.GetIndex(document).ToString(), HyperlinkType.Bookmark);
            //    paragraph.ApplyStyle(hyperlinkstyle.Name); 
            //}


            SaveDocument(document);

        }

        /// <summary>
        /// Заменяет одни слова на другие
        /// </summary>
        /// <param name="replacedWord">заменяемое слово</param>
        /// <param name="wordToReplace">слово-замена</param>
        private void ReplaceWords(string replacedWord, string wordToReplace)
        {
            foreach (var item in wordCase)
            {
                
                document.Replace($"{replacedWord}{item}", "", false, true);
            }
        }


        public static string FirstUpper(string str)
        {
            string[] s = str.Split(' ');

            for (int i = 0; i < s.Length; i++)
            {
                if (s[i].Length > 1)
                    s[i] = s[i].Substring(0, 1).ToUpper() + s[i].Substring(1, s[i].Length - 1).ToLower();
                else s[i] = s[i].ToUpper();
            }
            return string.Join(" ", s);
        }
        private static void SaveDocument(Document document)
        {
            document.SaveToFile("Hyperlink.docx", FileFormat.Docx);
        }
    }
}
