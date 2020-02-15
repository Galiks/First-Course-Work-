using Cyriller;
using Cyriller.Model;
using Spire.Doc;
using Spire.Doc.Collections;
using Spire.Doc.Documents;
using Spire.Doc.Fields;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;

namespace BusinessLogicLayer
{
    public class WordDocument
    {
        private readonly Document document;
        private readonly string filename;
        private string referencesWord;
        private readonly List<string> messages;
        private static int indexNextField = 0;
        private Paragraph referencesParagraph;
        private Section referencesSection;

        public Document Document => document;

        public List<string> Messages => messages;

        public int IndexNextField { get => indexNextField; private set => indexNextField = value; }

        public WordDocument(string filename)
        {
            this.filename = filename;
            document = new Document();
            document.LoadFromFile(filename + ".docx", FileFormat.Docx);
            messages = new List<string>();
        }

        public void IncreaseOfTwoindexNextField()
        {
            indexNextField += 2;
        }

        public void SetReferencesWord(string word)
        {
            //­­­U+00AD
            //005F
            referencesWord = word.Replace(' ', '\u005F');
        }
        public void CreateBookmarks(string word, string text)
        {
            SetReferencesWord(text);

            try
            {
                CyrNounCollection cyrNounCollection = new CyrNounCollection();
                CyrNoun cyrNoun = cyrNounCollection.Get(word, out CasesEnum @case, out NumbersEnum numbers);
                var nounsSet = new HashSet<string>(cyrNoun.Decline().ToList());
                foreach (var noun in cyrNoun.DeclinePlural().ToList())
                {
                    nounsSet.Add(noun);
                }
                var nouns = nounsSet.ToList();
                if (cyrNoun.WordType != WordTypesEnum.Surname)
                {
                    int nounLength = nouns.Count;
                    for (int i = 0; i < nounLength; i++)
                    {
                        nouns.Add(GetWordWithFirstLetterUpper(nouns[i]));
                    }
                }

                for (int i = 0; i < nouns.Count; i++)
                {
                    string noun = nouns[i];
                    if (i == 0)
                    {
                        CreateBookmarkByWord(noun, text);
                    }
                    else
                    {
                        CreateBookmarkByWord(noun);
                    }
                }
            }
            catch (CyrWordNotFoundException error)
            {
                CreateBookmarks(word, text);
            }
            finally
            {
                if (IndexNextField.Equals(default))
                {
                    IncreaseOfTwoindexNextField();
                }
            }
        }
        private void CreateBookmarkByWord(string word, string sentence = null)
        {
            //Create bookmark objects
            BookmarkStart start = new BookmarkStart(document, referencesWord);
            BookmarkEnd end = new BookmarkEnd(document, referencesWord);

            if (!string.IsNullOrWhiteSpace(sentence))
            {
                referencesParagraph = referencesSection.AddParagraph();
                referencesParagraph.AppendText(sentence);
            }

            int startIndex = 0;
            int paraIndex = referencesParagraph.ChildObjects.Count;

            //referParagraph.ChildObjects.Insert(startIndex, start);
            //referParagraph.ChildObjects.Insert(paraIndex, end);

            //int endIndex = referParagraph.ChildObjects.Count;

            //Insert the bookmark for the last paragraph
            referencesParagraph.ChildObjects.Insert(startIndex, start);
            referencesParagraph.ChildObjects.Insert(paraIndex, end);

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

            //Get the keywords
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

                if (paragraph.Equals(referencesParagraph))
                {
                    continue;
                }

                //Get the index of the keyword in its paragraph
                int index = paragraph.ChildObjects.IndexOf(tr);

                DocumentObject child = paragraph.ChildObjects[index];

                if (child.DocumentObjectType == DocumentObjectType.Field)
                {
                    Field textField = child as Field;

                    if (textField.Type == FieldType.FieldRef)
                    {
                        Messages.Add($"Поле {textField.FieldText} не было добавлено, так как оно уже имеет тип {FieldType.FieldRef.ToString()}");
                        continue;
                    }
                    else if (textField.Type == FieldType.FieldHyperlink)
                    {
                        Messages.Add($"Поле {textField.FieldText} не было добавлено, так как оно уже имеет тип {FieldType.FieldHyperlink.ToString()}");
                        continue;
                    }
                }

                //Create a cross-reference field, and link it to bookmark                   
                Field field = new Field(document)
                {
                    Type = FieldType.FieldRef
                };
                string code = $@"REF {referencesWord} \p \h";
                field.Code = code;

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
        public void CreateHyperlinksForText(string word, string text)
        {
            try
            {
                CyrNounCollection cyrNounCollection = new CyrNounCollection();
                CyrNoun cyrNoun = cyrNounCollection.Get(word, out CasesEnum @case, out NumbersEnum numbers);
                var nounsSet = new HashSet<string>(cyrNoun.Decline().ToList());
                foreach (var noun in cyrNoun.DeclinePlural().ToList())
                {
                    nounsSet.Add(noun);
                }
                var nouns = nounsSet.ToList();
                if (cyrNoun.WordType != WordTypesEnum.Surname)
                {
                    int nounLength = nouns.Count;
                    for (int i = 0; i < nounLength; i++)
                    {
                        nouns.Add(GetWordWithFirstLetterUpper(nouns[i]));
                    }
                }

                for (int i = 0; i < nouns.Count; i++)
                {
                    string noun = nouns[i];
                    CreateHyperlinkByWord(noun, text);
                }
            }
            catch (CyrWordNotFoundException error)
            {
                CreateHyperlinkByWord(word, text);
            }
            finally
            {
                if (IndexNextField.Equals(default))
                {
                    IncreaseOfTwoindexNextField();
                }
            }
        }
        private void CreateHyperlinkByWord(string word, string hyperlink)
        {
            TextSelection[] text = document.FindAllString(word, true, true);

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

                int index = tr.OwnerParagraph.ChildObjects.IndexOf(tr) - IndexNextField;

                Paragraph paragraph = tr.OwnerParagraph;

                DocumentObject child = paragraph.ChildObjects[index];

                if (child.DocumentObjectType == DocumentObjectType.Field)
                {
                    Field textField = child as Field;

                    if (textField.Type == FieldType.FieldRef)
                    {
                        Messages.Add($"Поле {textField.FieldText} не было добавлено, так как оно уже имеет тип {FieldType.FieldRef.ToString()}");
                        continue;
                    }
                    else if (textField.Type == FieldType.FieldHyperlink)
                    {
                        Messages.Add($"Поле {textField.FieldText} не было добавлено, так как оно уже имеет тип {FieldType.FieldHyperlink.ToString()}");
                        continue;
                    }
                }

                //Add hyperlink


                Field field = new Field(document)
                {
                    Code = "HYPERLINK \"" + hyperlink + "\"",

                    Type = FieldType.FieldHyperlink
                };

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



            SaveCurrentDicument();
        }
        public void CreatHyperlinkForImage()
        {

        }
        private void CreateHyperlinkByImage(DocPicture picture, string hyperlink)
        {

            int index = picture.OwnerParagraph.ChildObjects.IndexOf(picture) - IndexNextField;

            Paragraph paragraph = picture.OwnerParagraph;

            DocumentObject child = paragraph.ChildObjects[index];

            if (child.DocumentObjectType == DocumentObjectType.Field)
            {
                Field textField = child as Field;

                if (textField.Type == FieldType.FieldRef)
                {
                    Messages.Add($"Поле {textField.FieldText} не было добавлено, так как оно уже имеет тип {FieldType.FieldRef.ToString()}");
                    return;
                }
                else if (textField.Type == FieldType.FieldHyperlink)
                {
                    Messages.Add($"Поле {textField.FieldText} не было добавлено, так как оно уже имеет тип {FieldType.FieldHyperlink.ToString()}");
                    return;
                }
            }

            //Add hyperlink


            Field field = new Field(document)
            {
                Code = "HYPERLINK \"" + hyperlink + "\"",

                Type = FieldType.FieldHyperlink
            };

            picture.OwnerParagraph.ChildObjects.Insert(index, field);

            FieldMark fm = new FieldMark(document, FieldMarkType.FieldSeparator);

            picture.OwnerParagraph.ChildObjects.Insert(index + 1, fm);

            //Set character format

            picture.CharacterFormat.TextColor = Color.Blue;

            picture.CharacterFormat.UnderlineStyle = UnderlineStyle.Single;

            picture.CharacterFormat.Bold = picture.CharacterFormat.Bold;

            FieldMark fmend = new FieldMark(document, FieldMarkType.FieldEnd);

            picture.OwnerParagraph.ChildObjects.Insert(index + 3, fmend);

            field.End = fmend;




            SaveCurrentDicument();
        }

        public Tuple<Section, Paragraph> GetSectionAndParagraphByWord(string word)
        {
            var sections = Document.Sections;
            for (int i = sections.Count - 1; i > -1; i--)
            {
                Section section = sections[i];
                ParagraphCollection paragraphCollection = section.Paragraphs;
                foreach (Paragraph paragraph in paragraphCollection)
                {
                    if (paragraph.Text.Contains(word))
                    {
                        //referParagraph = paragraph;
                        return new Tuple<Section, Paragraph>(section, paragraph);
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
            return str[0].ToString().ToUpper() + str.Substring(1);
        }

        public void SaveCurrentDicument()
        {
            Document.SaveToFile(filename + ".docx", FileFormat.Docx);
        }
        #region Unused
        //private void SaveDocument(Document document, string filename)
        //{
        //    document.SaveToFile(filename + ".docx", FileFormat.Docx);
        //}
        //public void Result()
        //{
        //    Document doc = new Document();

        //    doc.LoadFromFile("result.docx", FileFormat.Docx);
        //    int indexOfSection = doc.Sections.Count - 1;
        //    Paragraph mainPara = doc.Sections[indexOfSection].AddParagraph();

        //    mainPara.AppendText(longText);

        //    mainPara = doc.Sections[0].AddParagraph();

        //    mainPara.AppendText("Ссылка на текст это полезная штука");

        //    Paragraph para = doc.Sections[0].AddParagraph();

        //    para.AppendText("Hypertext is also text. Hypertext is also text. Hypertext is also text.");
        //    para.AppendText("Делайте ссылки на предложения");

        //    //Find the string "Hypertext"

        //    string word = "гипертекст";
        //    CreateHyperlinkByWord(word);

        //    //Find the string "ссылки"

        //    //text = doc.FindAllString("ссылки", false, true);

        //    //foreach (TextSelection seletion in text)
        //    //{

        //    //    //Get the text range

        //    //    TextRange tr = seletion.GetAsOneRange();

        //    //    int index = tr.OwnerParagraph.ChildObjects.IndexOf(tr);

        //    //    //Add hyperlink

        //    //    Field field = new Field(doc);

        //    //    field.Code = "HYPERLINK \"" + "#Ссылка" + "\"";

        //    //    //field.Code = "HYPERLINK \"" + "http://www.e-iceblue.com" + "\"";

        //    //    field.Type = FieldType.FieldHyperlink;

        //    //    tr.OwnerParagraph.ChildObjects.Insert(index, field);

        //    //    FieldMark fm = new FieldMark(doc, FieldMarkType.FieldSeparator);

        //    //    tr.OwnerParagraph.ChildObjects.Insert(index + 1, fm);

        //    //    //Set character format

        //    //    tr.CharacterFormat.TextColor = Color.Blue;

        //    //    tr.CharacterFormat.UnderlineStyle = UnderlineStyle.Single;

        //    //    tr.CharacterFormat.Bold = tr.CharacterFormat.Bold;

        //    //    FieldMark fmend = new FieldMark(doc, FieldMarkType.FieldEnd);

        //    //    tr.OwnerParagraph.ChildObjects.Insert(index + 3, fmend);

        //    //    field.End = fmend;

        //    //}


        //    //doc.SaveToFile("result.docx", FileFormat.Docx);
        //}

        //public void RemoveHyperlinks()
        //{
        //    Document document = new Document();
        //    document.LoadFromFile("test.docx");

        //    #region Find hyperlink
        //    List<Field> hyperLink = FindAllHyperlinks(document);
        //    #endregion

        //    RemoveHyperlinksFromText(hyperLink);

        //    document.SaveToFile("test.docx", FileFormat.Docx);
        //}

        //private void RemoveHyperlinksFromText(List<Field> hyperLink)
        //{
        //    for (int i = hyperLink.Count - 1; i >= 0; i--)
        //    {
        //        Field field = hyperLink[i];
        //        int ownerParagraphIndex = field.OwnerParagraph.OwnerTextBody.ChildObjects.IndexOf(field.OwnerParagraph);
        //        int fieldIndex = field.OwnerParagraph.ChildObjects.IndexOf(field);
        //        Paragraph sepOwnerParagraph = field.Separator.OwnerParagraph;
        //        int sepOwnerParagraphIndex = field.Separator.OwnerParagraph.OwnerTextBody.ChildObjects.IndexOf(field.Separator.OwnerParagraph);
        //        int sepIndex = field.Separator.OwnerParagraph.ChildObjects.IndexOf(field.Separator);
        //        int endIndex = field.End.OwnerParagraph.ChildObjects.IndexOf(field.End);
        //        int endOwnerParagraphIndex = field.End.OwnerParagraph.OwnerTextBody.ChildObjects.IndexOf(field.End.OwnerParagraph);

        //        #region Remove font color and etc
        //        FormatFieldResultText(field.Separator.OwnerParagraph.OwnerTextBody, sepOwnerParagraphIndex, endOwnerParagraphIndex, sepIndex, endIndex);
        //        #endregion

        //        field.End.OwnerParagraph.ChildObjects.RemoveAt(endIndex);

        //        for (int j = sepOwnerParagraphIndex; j >= ownerParagraphIndex; j--)
        //        {
        //            if (j.Equals(sepOwnerParagraphIndex) && j.Equals(ownerParagraphIndex))
        //            {
        //                for (int k = sepIndex; k >= fieldIndex; k--)
        //                {
        //                    field.OwnerParagraph.ChildObjects.RemoveAt(k);
        //                }
        //            }
        //            else if (j.Equals(sepOwnerParagraphIndex))
        //            {
        //                for (int k = sepIndex; k >= 0; k--)
        //                {
        //                    sepOwnerParagraph.ChildObjects.RemoveAt(k);
        //                }
        //            }
        //            else if (j.Equals(ownerParagraphIndex))
        //            {
        //                for (int k = field.OwnerParagraph.ChildObjects.Count - 1; k >= fieldIndex; k--)
        //                {
        //                    field.OwnerParagraph.ChildObjects.RemoveAt(k);
        //                }
        //            }
        //            else
        //            {
        //                field.OwnerParagraph.ChildObjects.RemoveAt(j);
        //            }
        //        }
        //    }
        //}

        //private List<Field> FindAllHyperlinks(Document document)
        //{
        //    var hyperLink = new List<Field>();
        //    foreach (Section section in document.Sections)
        //    {
        //        foreach (DocumentObject sec in section.Body.ChildObjects)
        //        {
        //            if (sec.DocumentObjectType == DocumentObjectType.Paragraph)
        //            {
        //                foreach (DocumentObject para in (sec as Paragraph).ChildObjects)
        //                {
        //                    if (para.DocumentObjectType == DocumentObjectType.Field)
        //                    {
        //                        Field field = para as Field;

        //                        if (field.Type == FieldType.FieldHyperlink)
        //                        {
        //                            hyperLink.Add(field);
        //                        }
        //                    }
        //                }
        //            }
        //        }
        //    }

        //    return hyperLink;
        //}

        //private void FormatFieldResultText(Body ownerBody, int sepOwnerParaIndex, int endOwnerParaIndex, int sepIndex, int endIndex)
        //{
        //    for (int i = sepOwnerParaIndex; i <= endOwnerParaIndex; i++)
        //    {
        //        Paragraph para = ownerBody.ChildObjects[i] as Paragraph;
        //        if (i == sepOwnerParaIndex && i == endOwnerParaIndex)
        //        {
        //            for (int j = sepIndex + 1; j < endIndex; j++)
        //            {
        //                FormatText(para.ChildObjects[j] as TextRange);
        //            }

        //        }
        //        else if (i == sepOwnerParaIndex)
        //        {
        //            for (int j = sepIndex + 1; j < para.ChildObjects.Count; j++)
        //            {
        //                FormatText(para.ChildObjects[j] as TextRange);
        //            }
        //        }
        //        else if (i == endOwnerParaIndex)
        //        {
        //            for (int j = 0; j < endIndex; j++)
        //            {
        //                FormatText(para.ChildObjects[j] as TextRange);
        //            }
        //        }
        //        else
        //        {
        //            for (int j = 0; j < para.ChildObjects.Count; j++)
        //            {
        //                FormatText(para.ChildObjects[j] as TextRange);
        //            }
        //        }
        //    }
        //}

        //private void FormatText(TextRange tr)
        //{
        //    if (tr != null)
        //    {
        //        tr.CharacterFormat.TextColor = Color.Black;
        //        tr.CharacterFormat.UnderlineStyle = UnderlineStyle.None;
        //    }
        //} 
        #endregion

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
                                paragraphText.Replace(field.FieldText, $"<a href='{field.Code}'>{field.FieldText}</a>");
                            }
                            else if (field.Type == FieldType.FieldRef & !string.IsNullOrWhiteSpace(field.FieldText))
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
                        else if (child.DocumentObjectType == DocumentObjectType.Picture)
                        {
                            DocPicture picture = child as DocPicture;
                            paragraphText.Append("<img class='user-img' src=\"data:image/jpeg;base64," + Convert.ToBase64String(picture.ImageBytes) + "\" />");
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

            //foreach (var item in GetAllFootnotes())
            //{
            //    longText.Append(item);
            //}

            return longText.ToString();
        }

        public List<Field> FindAllLinksBySection(Section section)
        {
            var links = new List<Field>();

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
                                links.Add(field);
                            }
                            else if (field.Type == FieldType.FieldRef)
                            {
                                links.Add(field);
                            }
                        }
                    }
                }
            }
            return links;
        }
        public List<Field> FindAllLinks()
        {
            var links = new List<Field>();

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
                                    links.Add(field);
                                }
                            }
                        }
                    }
                }
            }
            return links;
        }

        public List<string> FindAllBookmarkBySection(Section section)
        {
            var bookmarks = new List<string>();

            BookmarksNavigator bookmarksNavigator = new BookmarksNavigator(document);
            var bookmarks1 = bookmarksNavigator.Document.Bookmarks;
            for (int i = 1; i < bookmarks1.Count; i++)
            {

                //bookmarks1[i].BookmarkEnd.OwnerParagraph.Text = "Hypertext is awesome";

                //var lastItem = bookmarks1[i].BookmarkStart.OwnerParagraph.Items.LastItem;
                //bookmarks1[i].BookmarkStart.OwnerParagraph.Items.Remove(lastItem);


                Section tempSection = document.AddSection();
                string text = "Hypertext is awesome";
                tempSection.AddParagraph().AppendText(text);

                ParagraphBase paragraphBaseFirstItem = tempSection.Paragraphs[0].Items.FirstItem as ParagraphBase;
                ParagraphBase paragraphBaseLastItem = tempSection.Paragraphs[0].Items.LastItem as ParagraphBase;
                TextBodySelection textBodySelection = new TextBodySelection(paragraphBaseFirstItem, paragraphBaseLastItem);
                TextBodyPart textBodyPart = new TextBodyPart(textBodySelection);

                bookmarksNavigator.MoveToBookmark(bookmarks1[i].Name);
                bookmarksNavigator.ReplaceBookmarkContent(textBodyPart, true, true);

                document.Sections.Remove(tempSection);

                SaveCurrentDicument();
            }

            #region find bookmark by section
            //foreach (DocumentObject sec in section.Body.ChildObjects)
            //{
            //    if (sec.DocumentObjectType == DocumentObjectType.Paragraph)
            //    {
            //        foreach (DocumentObject para in (sec as Paragraph).ChildObjects)
            //        {
            //            if (para.DocumentObjectType == DocumentObjectType.BookmarkStart)
            //            {
            //                BookmarkStart bookmarkStart = para as BookmarkStart;

            //                if (!string.IsNullOrWhiteSpace(bookmarkStart.Name))
            //                {
            //                    bookmarks.Add(bookmarkStart.Name);
            //                    BookmarksNavigator bookmarksNavigator = new BookmarksNavigator(document);
            //                    var bookmarks1 = bookmarksNavigator.Document.Bookmarks;
            //                    for (int i = 1; i < bookmarks1.Count; i++)
            //                    {
            //                        Section tempSection = document.AddSection();
            //                        string text = "Hypertext is awesome";
            //                        tempSection.AddParagraph().AppendText(text);

            //                        ParagraphBase paragraphBaseFirstItem = tempSection.Paragraphs[0].Items.FirstItem as ParagraphBase;
            //                        ParagraphBase paragraphBaseLastItem = tempSection.Paragraphs[0].Items.LastItem as ParagraphBase;
            //                        TextBodySelection textBodySelection = new TextBodySelection(paragraphBaseFirstItem, paragraphBaseLastItem);
            //                        TextBodyPart textBodyPart = new TextBodyPart(textBodySelection);

            //                        bookmarksNavigator.MoveToBookmark(bookmarks1[i].Name);

            //                        bookmarksNavigator.ReplaceBookmarkContent(textBodyPart);

            //                        document.Sections.Remove(tempSection);

            //                        SaveCurrentDicument();
            //                    }
            //                }
            //            }
            //        }
            //    }
            //} 
            #endregion
            return bookmarks;
        }

        public List<string> GetAllFootnotes()
        {
            var result = new HashSet<string>();

            foreach (Section section in document.Sections)
            {
                foreach (Paragraph paragraph in section.Paragraphs)
                {
                    int index = -1;
                    for (int i = 0, cnt = paragraph.ChildObjects.Count; i < cnt; i++)
                    {
                        ParagraphBase paragraphBase = paragraph.ChildObjects[i] as ParagraphBase;
                        if (paragraphBase is Footnote)
                        {
                            index = i;
                            break;
                        }
                    }

                    if (index > -1)
                    {
                        var footnotes = paragraph.ChildObjects[index].Document.Footnotes;
                        foreach (Footnote footnote in footnotes)
                        {
                            foreach (Paragraph item in footnote.TextBody.ChildObjects)
                            {
                                result.Add(item.Text);
                            }
                        }
                    }
                }
            }

            return result.ToList();
        }

        public void CreateReferencesSection()
        {
            var sectionAndParagraph = GetSectionAndParagraphByWord("Сноски");
            if (sectionAndParagraph == null)
            {
                Section sectionForReferences = document.Document.AddSection();
                var chapterForReferences = sectionForReferences.AddParagraph();
                chapterForReferences.AppendText("Сноски");
                chapterForReferences.AppendBreak(BreakType.LineBreak);
                //chapterForReferences.ApplyStyle(BuiltinStyle.Heading1);
                //chapterForReferences.ApplyStyle(BuiltinStyle.Title);

                //referencesParagraph = sectionForReferences.AddParagraph();
                referencesParagraph = chapterForReferences;
                referencesSection = sectionForReferences;
            }
            else
            {
                referencesParagraph = sectionAndParagraph.Item2;
                referencesSection = sectionAndParagraph.Item1;
            }

            SaveCurrentDicument();
        }
    }
}
