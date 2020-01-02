using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;

namespace BusinessLogicLayer
{
    public class WordDocument
    {
        //private const string text = " – это вид текста, который содержит в себе информацию или ссылку на место в произведении, на другую литературу или на какое-то событие в мире.";
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
            // ReplaceWords("гипертекст", "!!!");








            ////////заготовка на гипертекст
            //mainParagraph.AppendHyperlink("Гипертекст", "Гипертекст", HyperlinkType.Bookmark);
            mainParagraph.ApplyStyle(hyperlinkstyle.Name);

            for (int i = 0; i < 10; i++)
            {
                paragraph = section.AddParagraph();
                paragraph.AppendHyperlink(mainParagraph.Text, "Гипертекст", HyperlinkType.Bookmark);
                paragraph.AppendText(" поможет людям лучше понять текст.");
                //Paragraph subParagraph = section.AddParagraph();
                //subParagraph.AppendText("Слово употребляется в следующих параграфах: ");
                //subParagraph.AppendHyperlink(mainParagraph.ToString(), mainParagraph.GetIndex(document).ToString(), HyperlinkType.Bookmark);
                paragraph.ApplyStyle(hyperlinkstyle.Name);
            }


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


        public string FirstUpper(string str)
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
        private void SaveDocument(Document document)
        {
            document.SaveToFile("Hyperlink.docx", FileFormat.Docx);
        }

        public void Result()
        {
            Document doc = new Document();

            Section section = doc.AddSection();

            Paragraph mainPara = section.AddParagraph();

            mainPara.AppendText(longText);

            Paragraph para = section.AddParagraph();

            para.AppendText("Hypertext is also text. Hypertext is also text. Hypertext is also text.");

            //Find the string "Hypertext"

            TextSelection[] text = doc.FindAllString("Hypertext", false, true);

            foreach (TextSelection seletion in text)
            {

                //Get the text range

                TextRange tr = seletion.GetAsOneRange();

                int index = tr.OwnerParagraph.ChildObjects.IndexOf(tr);

                //Add hyperlink

                Field field = new Field(doc);

                field.Code = "HYPERLINK \"" + "#Гипертекст" + "\"";

                //field.Code = "HYPERLINK \"" + "http://www.e-iceblue.com" + "\"";

                field.Type = FieldType.FieldHyperlink;

                tr.OwnerParagraph.ChildObjects.Insert(index, field);

                FieldMark fm = new FieldMark(doc, FieldMarkType.FieldSeparator);

                tr.OwnerParagraph.ChildObjects.Insert(index + 1, fm);

                //Set character format

                tr.CharacterFormat.TextColor = Color.Blue;

                tr.CharacterFormat.UnderlineStyle = UnderlineStyle.Single;

                tr.CharacterFormat.Bold = tr.CharacterFormat.Bold;

                FieldMark fmend = new FieldMark(doc, FieldMarkType.FieldEnd);

                tr.OwnerParagraph.ChildObjects.Insert(index + 3, fmend);

                field.End = fmend;

            }

            doc.SaveToFile("result.docx", FileFormat.Docx);
        }
    }
}
