﻿using Spire.Doc;
using Spire.Doc.Collections;
using Spire.Doc.Documents;
using Spire.Doc.Fields;
using System.Collections.Generic;
using System.Drawing;
using System.Text;
using System.Text.RegularExpressions;

namespace BusinessLogicLayer
{
    public class WordDocument
    {
        //private const string text = " – это вид текста, который содержит в себе информацию или ссылку на место в произведении, на другую литературу или на какое-то событие в мире.";
        private const string longText = "Гипертекст в тексте может выглядеть как гипертекст, а может как гипертекст на гипертексте, где гипертекст сам является гипертекстом.";
        private readonly Document document;
        //private readonly Section section;
        private Paragraph referParagraph;
        private readonly string filename;
        private string referencesWord;

        public Document Document => document;

        public string ReferencesWord { get => referencesWord; private set => referencesWord = value; }

        public WordDocument(string filename)
        {
            this.filename = filename;
            document = new Document();
            document.LoadFromFile(filename + ".docx", FileFormat.Docx);

            //not use
            //section = document.AddSection();
            //mainParagraph = section.AddParagraph();
        }

        public void SetReferencesWord(string word)
        {
            ReferencesWord = word;
        }

        public void CreateBookmarks(string word, string sentence)
        {
            //Create bookmark objects
            BookmarkStart start = new BookmarkStart(document, word);
            BookmarkEnd end = new BookmarkEnd(document, word);
            //BookmarkStart start = new BookmarkStart(document, "MyBk");
            //BookmarkEnd end = new BookmarkEnd(document, "MyBk");



            //Get the last paragraph----"Hypertext – bla bla bla bla text bla bla blab la hyperlink and save the world"
            Paragraph lastPara = referParagraph;



            int startIndex = 0;

            int endIndex = lastPara.ChildObjects.Count;



            //Insert the bookmark for the last paragraph

            lastPara.ChildObjects.Insert(startIndex, start);

            lastPara.ChildObjects.Insert(endIndex, end);



            //Find the keyword "Hypertext"

            TextSelection[] text = document.FindAllString(word, true, true);

            if (text == null)
            {
                return;
            }

            if (text.Length == 0)
            {
                return;
            }

            //Get the first keyword

            for (int i = 0; i < text.Length; i++)
            {
                TextSelection keywordOne = text[i];

                //Get the textrange its locates

                TextRange tr = keywordOne.GetAsOneRange();

                //Set the formatting

                tr.CharacterFormat.UnderlineStyle = UnderlineStyle.Single;

                tr.CharacterFormat.TextColor = Color.Blue;

                //Get the paragraph it locates

                Paragraph paragraph = tr.OwnerParagraph;

                //Get the index of the keyword in its paragraph

                int index = paragraph.ChildObjects.IndexOf(tr);



                //Create a cross-reference field, and link it to bookmark                   

                Field field = new Field(document);

                field.Type = FieldType.FieldRef;

                field.Code = $@"REF {word} \p \h";



                //Insert field

                paragraph.ChildObjects.Insert(index, field);



                //Insert FieldSeparator object

                FieldMark fieldSeparator = new FieldMark(document, FieldMarkType.FieldSeparator);

                paragraph.ChildObjects.Insert(index + 1, fieldSeparator);



                //Insert FieldEnd object to mark the end of the field

                FieldMark fieldEnd = new FieldMark(document, FieldMarkType.FieldEnd);

                paragraph.ChildObjects.Insert(index + 3, fieldEnd); 
            }

            SaveCurrentDicument();
        }

        public void CreateHyperlinks(string word)
        {
            //пока что заготовленный файл
            //string filename = "hyperlink";
            //document.LoadFromFile(filename + ".docx", FileFormat.Docx);

            //слово тоже заготовлено заранее
            //string word = "Гипертекст";

            CreateHyperlinkByWord(word);
            SaveCurrentDicument();
        }

        private void CreateHyperlinkByWord(string word)
        {
            TextSelection[] text = document.FindAllPattern(new Regex(word));

            if (text == null)
            {
                return;
            }

            if (text.Length == 0)
            {
                return;
            }

            foreach (TextSelection seletion in text)
            {

                //Get the text range

                TextRange tr = seletion.GetAsOneRange();

                int index = tr.OwnerParagraph.ChildObjects.IndexOf(tr);

                //Add hyperlink

                Field field = new Field(document);

                field.Code = "HYPERLINK \"" + "#" + referencesWord + "\"";

                field.Type = FieldType.FieldHyperlink;

                tr.OwnerParagraph.ChildObjects.Insert(index, field);

                FieldMark fm = new FieldMark(document, FieldMarkType.FieldSeparator);

                tr.OwnerParagraph.ChildObjects.Insert(index + 1, fm);

                //Set character format

                tr.CharacterFormat.TextColor = Color.Blue;

                tr.CharacterFormat.UnderlineStyle = UnderlineStyle.Single;

                tr.CharacterFormat.Bold = tr.CharacterFormat.Bold;

                FieldMark fmend = new FieldMark(document, FieldMarkType.FieldEnd);

                tr.OwnerParagraph.ChildObjects.Insert(index + 3, fmend);

                field.End = fmend;
            }
        }

        public Paragraph GetParagraphByWord(string word)
        {
            var sections = Document.Sections;
            for (int i = sections.Count - 1; i > -1; i--)
            {
                Section section = sections[i];
                int count = section.Paragraphs.Count;
                ParagraphCollection paragraphCollection = section.Paragraphs;
                foreach (Paragraph paragraph in paragraphCollection)
                {
                    if (paragraph.Text.Contains(word))
                    {
                        referParagraph = paragraph;
                        return paragraph;
                    }
                }
            }

            return null;
        }

        #region Create Document
        //public void CreateDocument()
        //{
        //    Paragraph paragraph = section.AddParagraph();

        //    //Set Paragraph Styles
        //    ParagraphStyle txtStyle = new ParagraphStyle(document);
        //    txtStyle.Name = "Style";
        //    txtStyle.CharacterFormat.FontName = "Impact";
        //    txtStyle.CharacterFormat.FontSize = 16;
        //    txtStyle.CharacterFormat.TextColor = Color.RosyBrown;
        //    document.Styles.Add(txtStyle);
        //    //Set Hyperlink Styles
        //    ParagraphStyle hyperlinkstyle = new ParagraphStyle(document);
        //    hyperlinkstyle.Name = "linkStyle";
        //    hyperlinkstyle.CharacterFormat.FontName = "Calibri";
        //    hyperlinkstyle.CharacterFormat.FontSize = 15;
        //    document.Styles.Add(hyperlinkstyle);

        //    ///главный параграф
        //    mainParagraph.AppendText(longText);
        //    // ReplaceWords("гипертекст", "!!!");


        //    ////////заготовка на гипертекст
        //    //mainParagraph.AppendHyperlink("Гипертекст", "Гипертекст", HyperlinkType.Bookmark);
        //    mainParagraph.ApplyStyle(hyperlinkstyle.Name);

        //    for (int i = 0; i < 10; i++)
        //    {
        //        paragraph = section.AddParagraph();
        //        paragraph.AppendHyperlink(mainParagraph.Text, "Гипертекст", HyperlinkType.Bookmark);
        //        paragraph.AppendText(" поможет людям лучше понять текст.");
        //        //Paragraph subParagraph = section.AddParagraph();
        //        //subParagraph.AppendText("Слово употребляется в следующих параграфах: ");
        //        //subParagraph.AppendHyperlink(mainParagraph.ToString(), mainParagraph.GetIndex(document).ToString(), HyperlinkType.Bookmark);
        //        paragraph.ApplyStyle(hyperlinkstyle.Name);
        //    }


        //    SaveDocument(document, "Hyperlink");

        //} 
        #endregion

        public static string GetWordWithFirstLetterUpper(string str)
        {
            return str.Replace(str[0].ToString(), str[0].ToString().ToUpper());
        }

        private void SaveDocument(Document document, string filename)
        {
            document.SaveToFile(filename + ".docx", FileFormat.Docx);
        }

        public void SaveCurrentDicument()
        {
            Document.SaveToFile(filename + ".docx", FileFormat.Docx);
        }

        public void Result()
        {
            Document doc = new Document();

            doc.LoadFromFile("result.docx", FileFormat.Docx);
            int indexOfSection = doc.Sections.Count - 1;
            Paragraph mainPara = doc.Sections[indexOfSection].AddParagraph();

            mainPara.AppendText(longText);

            mainPara = doc.Sections[0].AddParagraph();

            mainPara.AppendText("Ссылка на текст это полезная штука");

            Paragraph para = doc.Sections[0].AddParagraph();

            para.AppendText("Hypertext is also text. Hypertext is also text. Hypertext is also text.");
            para.AppendText("Делайте ссылки на предложения");

            //Find the string "Hypertext"

            string word = "гипертекст";
            CreateHyperlinkByWord(word);

            //Find the string "ссылки"

            //text = doc.FindAllString("ссылки", false, true);

            //foreach (TextSelection seletion in text)
            //{

            //    //Get the text range

            //    TextRange tr = seletion.GetAsOneRange();

            //    int index = tr.OwnerParagraph.ChildObjects.IndexOf(tr);

            //    //Add hyperlink

            //    Field field = new Field(doc);

            //    field.Code = "HYPERLINK \"" + "#Ссылка" + "\"";

            //    //field.Code = "HYPERLINK \"" + "http://www.e-iceblue.com" + "\"";

            //    field.Type = FieldType.FieldHyperlink;

            //    tr.OwnerParagraph.ChildObjects.Insert(index, field);

            //    FieldMark fm = new FieldMark(doc, FieldMarkType.FieldSeparator);

            //    tr.OwnerParagraph.ChildObjects.Insert(index + 1, fm);

            //    //Set character format

            //    tr.CharacterFormat.TextColor = Color.Blue;

            //    tr.CharacterFormat.UnderlineStyle = UnderlineStyle.Single;

            //    tr.CharacterFormat.Bold = tr.CharacterFormat.Bold;

            //    FieldMark fmend = new FieldMark(doc, FieldMarkType.FieldEnd);

            //    tr.OwnerParagraph.ChildObjects.Insert(index + 3, fmend);

            //    field.End = fmend;

            //}


            //doc.SaveToFile("result.docx", FileFormat.Docx);
        }

        public void RemoveHyperlinks()
        {
            Document document = new Document();
            document.LoadFromFile("test.docx");

            #region Find hyperlink
            List<Field> hyperLink = FindAllHyperlinks(document);
            #endregion

            RemoveHyperlinksFromText(hyperLink);

            document.SaveToFile("test.docx", FileFormat.Docx);
        }

        private void RemoveHyperlinksFromText(List<Field> hyperLink)
        {
            for (int i = hyperLink.Count - 1; i >= 0; i--)
            {
                Field field = hyperLink[i];
                int ownerParagraphIndex = field.OwnerParagraph.OwnerTextBody.ChildObjects.IndexOf(field.OwnerParagraph);
                int fieldIndex = field.OwnerParagraph.ChildObjects.IndexOf(field);
                Paragraph sepOwnerParagraph = field.Separator.OwnerParagraph;
                int sepOwnerParagraphIndex = field.Separator.OwnerParagraph.OwnerTextBody.ChildObjects.IndexOf(field.Separator.OwnerParagraph);
                int sepIndex = field.Separator.OwnerParagraph.ChildObjects.IndexOf(field.Separator);
                int endIndex = field.End.OwnerParagraph.ChildObjects.IndexOf(field.End);
                int endOwnerParagraphIndex = field.End.OwnerParagraph.OwnerTextBody.ChildObjects.IndexOf(field.End.OwnerParagraph);

                #region Remove font color and etc
                FormatFieldResultText(field.Separator.OwnerParagraph.OwnerTextBody, sepOwnerParagraphIndex, endOwnerParagraphIndex, sepIndex, endIndex);
                #endregion

                field.End.OwnerParagraph.ChildObjects.RemoveAt(endIndex);

                for (int j = sepOwnerParagraphIndex; j >= ownerParagraphIndex; j--)
                {
                    if (j.Equals(sepOwnerParagraphIndex) && j.Equals(ownerParagraphIndex))
                    {
                        for (int k = sepIndex; k >= fieldIndex; k--)
                        {
                            field.OwnerParagraph.ChildObjects.RemoveAt(k);
                        }
                    }
                    else if (j.Equals(sepOwnerParagraphIndex))
                    {
                        for (int k = sepIndex; k >= 0; k--)
                        {
                            sepOwnerParagraph.ChildObjects.RemoveAt(k);
                        }
                    }
                    else if (j.Equals(ownerParagraphIndex))
                    {
                        for (int k = field.OwnerParagraph.ChildObjects.Count - 1; k >= fieldIndex; k--)
                        {
                            field.OwnerParagraph.ChildObjects.RemoveAt(k);
                        }
                    }
                    else
                    {
                        field.OwnerParagraph.ChildObjects.RemoveAt(j);
                    }
                }
            }
        }

        private List<Field> FindAllHyperlinks(Document document)
        {
            var hyperLink = new List<Field>();
            foreach (Section section in document.Sections)
            {
                foreach (DocumentObject sec in section.Body.ChildObjects)
                {
                    if (sec.DocumentObjectType == DocumentObjectType.Paragraph)
                    {
                        foreach (DocumentObject para in (sec as Paragraph).ChildObjects)
                        {
                            if (para.DocumentObjectType == DocumentObjectType.Field)
                            {
                                Field field = para as Field;

                                if (field.Type == FieldType.FieldHyperlink)
                                {
                                    hyperLink.Add(field);
                                }
                            }
                        }
                    }
                }
            }

            return hyperLink;
        }

        private void FormatFieldResultText(Body ownerBody, int sepOwnerParaIndex, int endOwnerParaIndex, int sepIndex, int endIndex)
        {
            for (int i = sepOwnerParaIndex; i <= endOwnerParaIndex; i++)
            {
                Paragraph para = ownerBody.ChildObjects[i] as Paragraph;
                if (i == sepOwnerParaIndex && i == endOwnerParaIndex)
                {
                    for (int j = sepIndex + 1; j < endIndex; j++)
                    {
                        FormatText(para.ChildObjects[j] as TextRange);
                    }

                }
                else if (i == sepOwnerParaIndex)
                {
                    for (int j = sepIndex + 1; j < para.ChildObjects.Count; j++)
                    {
                        FormatText(para.ChildObjects[j] as TextRange);
                    }
                }
                else if (i == endOwnerParaIndex)
                {
                    for (int j = 0; j < endIndex; j++)
                    {
                        FormatText(para.ChildObjects[j] as TextRange);
                    }
                }
                else
                {
                    for (int j = 0; j < para.ChildObjects.Count; j++)
                    {
                        FormatText(para.ChildObjects[j] as TextRange);
                    }
                }
            }
        }

        private void FormatText(TextRange tr)
        {
            if (tr != null)
            {
                tr.CharacterFormat.TextColor = Color.Black;
                tr.CharacterFormat.UnderlineStyle = UnderlineStyle.None;
            }
        }

        public string GetTextFromDocument()
        {
            StringBuilder longText = new StringBuilder();

            Document document = new Document();
            document.LoadFromFile(filename + ".docx", FileFormat.Docx);

            foreach (Section section in document.Sections)
            {
                longText.AppendLine("<div>");
                foreach (Paragraph paragraph in section.Paragraphs)
                {
                    StringBuilder paragraphText = new StringBuilder(paragraph.Text);

                    foreach (DocumentObject child in paragraph.ChildObjects)
                    {
                        if (child.DocumentObjectType == DocumentObjectType.Field)
                        {
                            Field field = child as Field;
                            if (field.Type == FieldType.FieldHyperlink & !string.IsNullOrWhiteSpace(field.FieldText))
                            {
                                paragraphText.Replace(field.FieldText, $"<strong>{field.FieldText}</strong>");
                            }
                        }
                        else if (child.DocumentObjectType == DocumentObjectType.Break)
                        {
                            Break @break = child as Break;
                            if (@break.BreakType == BreakType.LineBreak)
                            {
                                paragraphText.Replace("\v", $"<br>");
                            }
                        }
                    }

                    longText.AppendLine($"{paragraphText}<br>");
                }
                for (int i = 0; i < 5; i++)
                {
                    longText.AppendLine("<br>");
                }
                longText.AppendLine("</div>");
            }

            return longText.ToString();
        }
    }
}
