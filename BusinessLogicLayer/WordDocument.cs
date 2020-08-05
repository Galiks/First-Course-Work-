using Cyriller;
using Cyriller.Model;
using NLog;
using Spire.Doc;
using Spire.Doc.Collections;
using Spire.Doc.Documents;
using Spire.Doc.Fields;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BusinessLogicLayer
{
    /// <summary>
    /// 
    /// </summary>
    public class WordDocument
    {
        private static readonly Logger loggerException = LogManager.GetLogger("exception");
        private static readonly Logger loggerUser = LogManager.GetLogger("user");

        private const string badBookmark = "_GoBack";
        private const int width = 470;
        private const int height = 340;
        private readonly Document document;
        private readonly string filename;
        private string referencesWord;
        private readonly HashSet<string> messages;
        private static int indexNextField = 0;
        private Paragraph referencesParagraph;
        private static Section referencesSection;
        private const string styleName = "MyReferences";
        /// <summary>
        /// 
        /// </summary>
        public Document Document => document;
        /// <summary>
        /// 
        /// </summary>
        public HashSet<string> Messages => messages;
        /// <summary>
        /// 
        /// </summary>
        public int IndexNextField { get => indexNextField; private set => indexNextField = value; }
        /// <summary>
        /// 
        /// </summary>
        public static Section ReferencesSection { get => referencesSection; private set => referencesSection = value; }
        /// <summary>
        /// 
        /// </summary>
        public static int IndexReferencesSection { get => indexReferencesSection; set => indexReferencesSection = value; }

        private readonly string filepath;

        private static int indexReferencesSection;
        /// <summary>
        /// 
        /// </summary>
        /// <param name="filename"></param>
        public WordDocument(string filename)
        {
            loggerUser.Info($"Начата работа с файлом: {filename}");
            this.filename = filename;
            document = new Document();
            //проблема с дублями файла
            filepath = filename;
            document.LoadFromFile(filename);
            messages = new HashSet<string>();
        }
        /// <summary>
        /// 
        /// </summary>
        public void IncreaseOfTwoindexNextField()
        {
            indexNextField += 2;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="word"></param>
        public void SetReferencesWord(string word)
        {
            //­­­U+00AD
            //005F       
            if (!string.IsNullOrWhiteSpace(word))
            {
                referencesWord = TransformWordWithUnderline(word);
            }
        }
        private string TransformWordWithUnderline(string word)
        {
            if (!string.IsNullOrWhiteSpace(word))
            {
                var splitWord = word.Split("\\");
                return splitWord[^1].Replace(' ', '\u005F').Replace('/', '\u005F');
            }
            else
            {
                return null;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="word"></param>
        /// <param name="text"></param>
        /// <param name="count"></param>
        public void CreateBookmarksForText(string word, string text, int count = default)
        {
            SetReferencesWord(text);

            try
            {
                List<string> nouns = GetWordsByCases(word);

                Parallel.For(0, nouns.Count, i =>
                {
                    string noun = nouns[i];
                    if (i == 0)
                    {
                        CreateBookmarkByWord(noun, text, count: count);
                    }
                    else
                    {
                        CreateBookmarkByWord(noun, count: count);
                    }
                });

                //for (int i = 0; i < nouns.Count; i++)
                //{
                //string noun = nouns[i];
                //if (i == 0)
                //{
                //    CreateBookmarkByWord(noun, text, count: count);
                //}
                //else
                //{
                //    CreateBookmarkByWord(noun, count: count);
                //}
                //}
            }
            catch (CyrWordNotFoundException error)
            {
                loggerException.Error(error.Message);
                CreateBookmarkByWord(word, text, count);
            }
            finally
            {
                if (IndexNextField.Equals(default))
                {
                    IncreaseOfTwoindexNextField();
                }
            }
        }
        /// <summary>
        /// Возвращает список слов по падежам
        /// </summary>
        /// <param name="word">начальное слово</param>
        /// <returns></returns>
        private List<string> GetWordsByCases(string word)
        {
            CyrNounCollection cyrNounCollection = new CyrNounCollection();
            CyrNoun cyrNoun = cyrNounCollection.Get(word, out CasesEnum @case, out NumbersEnum numbers);
            var nounsSet = new HashSet<string>(cyrNoun.Decline().ToList());
            Parallel.ForEach(cyrNoun.DeclinePlural().ToList(), noun =>
            {
                nounsSet.Add(noun);
            });
            //foreach (var noun in cyrNoun.DeclinePlural().ToList())
            //{
            //    nounsSet.Add(noun);
            //}
            var nouns = nounsSet.ToList();
            if (cyrNoun.WordType != WordTypesEnum.Surname)
            {
                int nounLength = nouns.Count;
                Parallel.For(0, nounLength, i =>
                {
                    nouns.Add(GetWordWithFirstLetterUpper(nouns[i]));
                });
                //for (int i = 0; i < nounLength; i++)
                //{
                //    nouns.Add(GetWordWithFirstLetterUpper(nouns[i]));
                //}
            }

            return nouns;
        }
        private void CreateBookmarkByWord(string word, string sentence = null, int count = default)
        {
            //Create bookmark objects
            //BookmarkStart start = new BookmarkStart(document, referencesWord);
            //BookmarkEnd end = new BookmarkEnd(document, referencesWord);

            if (!string.IsNullOrWhiteSpace(sentence))
            {
                referencesParagraph = ReferencesSection.AddParagraph();
                referencesParagraph.AppendBookmarkStart(referencesWord);
                referencesParagraph.AppendText(sentence);
                referencesParagraph.AppendBookmarkEnd(referencesWord);
            }

            TextSelection[] text = document.FindAllString(word, true, true);

            if (text == null)
            {
                return;
            }

            if (text.Length == 0)
            {
                return;
            }

            int findLength = count == default ? text.Length : count;

            //Get the keywords
            //for (int i = 0; i < findLength; i++)
            Parallel.For(0, findLength, i =>
            {
                CreateBookmarkByWordHandler(i, text);
            });


            SaveCurrentDocument();
        }

        private void CreateBookmarkByWordHandler(int i, TextSelection[] text)
        {
            TextSelection keywordOne = text[i];
            TextRange tr = null;

            //Get the textrange its locates
            try
            {
                tr = keywordOne.GetAsOneRange();
            }
            catch (Exception)
            {
                CreateBookmarkByWordHandler(i, text);
            }

            //Set the formatting
            tr.CharacterFormat.UnderlineStyle = UnderlineStyle.Single;
            tr.CharacterFormat.TextColor = Color.Blue;

            //Get the paragraph it locates
            Paragraph paragraph = tr.OwnerParagraph;

            if (paragraph.Equals(referencesParagraph))
            {
                return;
            }

            //Get the index of the keyword in its paragraph
            int index = paragraph.ChildObjects.IndexOf(tr);

            DocumentObject child = paragraph.ChildObjects[index];

            if (child.DocumentObjectType == DocumentObjectType.Field)
            {
                Field textField = child as Field;

                if (textField.Type == FieldType.FieldRef)
                {
                    Messages.Add($"Поле {textField.FieldText} не было добавлено, так как оно уже имеет тип {FieldType.FieldRef}");
                    return;
                }
                else if (textField.Type == FieldType.FieldHyperlink)
                {
                    Messages.Add($"Поле {textField.FieldText} не было добавлено, так как оно уже имеет тип {FieldType.FieldHyperlink}");
                    return;
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

        /// <summary>
        /// 
        /// </summary>
        /// <param name="path"></param>
        /// <param name="word"></param>
        /// <param name="count"></param>
        public void CreateBookmarksForImage(string path, string word, int count = default)
        {
            SetReferencesWord(path);
            try
            {
                List<string> nouns = GetWordsByCases(word);

                for (int i = 0; i < nouns.Count; i++)
                {
                    string noun = nouns[i];

                    CreateBookmarkByImage(path, noun, count);
                    //SetBookmarkForImage(path);
                }
            }
            catch (CyrWordNotFoundException)
            {
                CreateBookmarkByImage(path, word, count);
                //SetBookmarkForImage(path);
            }
            finally
            {
                if (IndexNextField.Equals(default))
                {
                    IncreaseOfTwoindexNextField();
                }

            }
        }
        private void CreateBookmarkByImage(string path, string word, int count)
        {
            SetReferencesWord(path);
            BookmarksNavigator bn = new BookmarksNavigator(document);
            bn.MoveToBookmark(referencesWord, true, true);

            if (bn.CurrentBookmark == null)
            {
                var para = referencesSection.AddParagraph();
                para.AppendBookmarkStart(referencesWord);
                para.AppendBookmarkEnd(referencesWord);
                bn.MoveToBookmark(referencesWord, true, true);
                Section section0 = document.AddSection();
                Paragraph paragraph = section0.AddParagraph();
                Image image = Image.FromFile(path);
                DocPicture picture = paragraph.AppendPicture(image);
                picture.Width = width;
                picture.Height = height;
                bn.InsertParagraph(paragraph);
                document.Sections.Remove(section0);
            }

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

            //Create bookmark objects

            BookmarksNavigator bookmarksNavigator = new BookmarksNavigator(document);
            bookmarksNavigator.MoveToBookmark(referencesWord);

            if (bookmarksNavigator.CurrentBookmark == null)
            {
                Image image = Image.FromFile(path);

                SetReferencesWord(path);

                var referParagraph = referencesSection.AddParagraph();
                referParagraph.AppendBookmarkStart(referencesWord);
                referParagraph.AppendPicture(image);
                referParagraph.AppendBookmarkEnd(referencesWord);
                SaveCurrentDocument();

                bookmarksNavigator.MoveToBookmark(referencesWord);
                bookmarksNavigator.InsertParagraph(referParagraph);

                SaveCurrentDocument();
            }

            int findLength = count == default ? text.Length : count;

            //Get the keywords
            //for (int i = 0; i < findLength; i++)
            Parallel.For(0, findLength, i =>
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
                    return;
                }

                //Get the index of the keyword in its paragraph
                int index = paragraph.ChildObjects.IndexOf(tr);

                DocumentObject child = paragraph.ChildObjects[index];

                if (child.DocumentObjectType == DocumentObjectType.Field)
                {
                    Field textField = child as Field;

                    if (textField.Type == FieldType.FieldRef)
                    {
                        Messages.Add($"Поле {textField.FieldText} не было добавлено, так как оно уже имеет тип {FieldType.FieldRef}");
                        return;
                    }
                    else if (textField.Type == FieldType.FieldHyperlink)
                    {
                        Messages.Add($"Поле {textField.FieldText} не было добавлено, так как оно уже имеет тип {FieldType.FieldHyperlink}");
                        return;
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
            });

            SaveCurrentDocument();
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="word"></param>
        /// <param name="text"></param>
        /// <param name="count"></param>
        public void CreateHyperlinksForText(string word, string text, int count = default)
        {
            try
            {
                CyrNounCollection cyrNounCollection = new CyrNounCollection();
                CyrNoun cyrNoun = cyrNounCollection.Get(word, out CasesEnum @case, out NumbersEnum numbers);
                var nounsSet = new HashSet<string>(cyrNoun.Decline().ToList());
                //foreach (var noun in cyrNoun.DeclinePlural().ToList())
                Parallel.ForEach(cyrNoun.DeclinePlural().ToList(), noun =>
                {
                    nounsSet.Add(noun);
                });
                var nouns = nounsSet.ToList();
                if (cyrNoun.WordType != WordTypesEnum.Surname)
                {
                    int nounLength = nouns.Count;
                    //for (int i = 0; i < nounLength; i++)
                    Parallel.For(0, nounLength, i =>
                    {
                        nouns.Add(GetWordWithFirstLetterUpper(nouns[i]));
                    });
                }

                for (int i = 0; i < nouns.Count; i++)
                {
                    string noun = nouns[i];
                    CreateHyperlinkByWord(noun, text, count);
                }
            }
            catch (CyrWordNotFoundException)
            {
                CreateHyperlinkByWord(word, text, count);
            }
            finally
            {
                if (IndexNextField.Equals(default))
                {
                    IncreaseOfTwoindexNextField();
                }
            }
        }
        private void CreateHyperlinkByWord(string word, string hyperlink, int count)
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

            int findLength = count == default ? text.Length : count;

            //for (int i = 0; i < findLength; i++)
            Parallel.For(0, findLength, i =>
            {
                TextSelection seletion = text[i];

                //Get the text range

                TextRange tr = seletion.GetAsOneRange();

                int index = tr.OwnerParagraph.ChildObjects.IndexOf(tr) - IndexNextField;
                Paragraph paragraph = tr.OwnerParagraph;

                DocumentObject child;
                try
                {
                    child = paragraph.ChildObjects[index];
                }
                catch (IndexOutOfRangeException)
                {
                    index = tr.OwnerParagraph.ChildObjects.IndexOf(tr);
                    child = paragraph.ChildObjects[index];
                }

                if (child.DocumentObjectType == DocumentObjectType.Field)
                {
                    Field textField = child as Field;

                    if (textField.Type == FieldType.FieldRef)
                    {
                        Messages.Add($"Поле {textField.FieldText} не было добавлено, так как оно уже имеет тип {FieldType.FieldRef}");
                        return;
                    }
                    else if (textField.Type == FieldType.FieldHyperlink)
                    {
                        Messages.Add($"Поле {textField.FieldText} не было добавлено, так как оно уже имеет тип {FieldType.FieldHyperlink}");
                        return;
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
            });



            SaveCurrentDocument();
        }
        /// <summary>
        /// Добавляет гиперссылку на изображение, добавляемое в документ
        /// </summary>
        /// <param name="path">путь к изображению</param>
        /// <param name="hyperlink">гиперссылка</param>
        public void CreatHyperlinkForImage(string path, string hyperlink)
        {
            DocPicture picture = new DocPicture(Document);
            picture.LoadImage(path);
            picture.Width = 470;
            picture.Height = 340;

            ReferencesSection.AddParagraph().AppendHyperlink(hyperlink, picture, HyperlinkType.WebLink);
            SaveCurrentDocument();
        }

        /// <summary>
        /// Добавляет гиперссылку на имеющееся в документе изображение
        /// </summary>
        /// <param name="picture"></param>
        /// <param name="hyperlink"></param>
        public void CreateHyperlinkForImage(DocPicture picture, string hyperlink)
        {
            picture.Width = 470;
            picture.Height = 340;

            ReferencesSection.AddParagraph().AppendHyperlink(hyperlink, picture, HyperlinkType.WebLink);
            SaveCurrentDocument();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="style"></param>
        /// <returns></returns>
        public Tuple<Section, Paragraph> GetSectionAndParagraphByStyleName(string style)
        {
            var sections = Document.Sections;
            for (int i = sections.Count - 1; i > -1; i--)
            {
                Section section = sections[i];
                ParagraphCollection paragraphCollection = section.Paragraphs;
                foreach (Paragraph paragraph in paragraphCollection)
                {
                    if (paragraph.StyleName == style)
                    {
                        //referParagraph = paragraph;
                        return new Tuple<Section, Paragraph>(section, paragraph);
                    }
                }
            }

            return null;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        public static string GetWordWithFirstLetterUpper(string str)
        {
            return str[0].ToString().ToUpper() + str.Substring(1);
        }
        /// <summary>
        /// 
        /// </summary>
        public void SaveCurrentDocument()
        {
            Document.SaveToFile(filename);
        }
        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public string GetTextFromDocument()
        {
            string filename = filepath + ".html";
            document.SaveToFile(filename, FileFormat.Html);

            string[] splitFilename = filename.Split("\\");
            string rightFilename = null;

            //for (int i = 0; i < splitFilename.Length; i++)
            Parallel.For(0, splitFilename.Length, i =>
            {
                if (splitFilename[i].Equals("wwwroot"))
                {
                    rightFilename = @"\" + string.Join(@"\", splitFilename, i + 1, splitFilename.Length - i - 1);
                    //return splitFilename.Join("\\", splitFilename, i, splitFilename.Length - 1);
                }
            });

            if (string.IsNullOrWhiteSpace(rightFilename))
            {
                return "error";
            }
            else
            {
                return rightFilename;
            }

            #region old version
            //StringBuilder longText = new StringBuilder();
            //GetAllFootnotes();
            //foreach (Section section in document.Sections)
            //{
            //    bool flagForNumbered = false;
            //    longText.AppendLine(" < div>");
            //    foreach (Paragraph paragraph in section.Paragraphs)
            //    {
            //        string aligment = GetAligment(paragraph);
            //        string fontName = "";
            //        float? fontSize = default;
            //        StringBuilder paragraphText = new StringBuilder();

            //        #region нерабочий список
            //        if (paragraph.NextSibling != null)
            //        {
            //            Paragraph nextParagraph = paragraph.NextSibling as Paragraph;
            //            if (nextParagraph?.ListFormat.ListType == ListType.Numbered & flagForNumbered == false)
            //            {
            //                paragraphText.Append("<ol>");
            //                flagForNumbered = true;
            //            }
            //            else if (nextParagraph?.ListFormat.ListType != ListType.Numbered & flagForNumbered)
            //            {
            //                paragraphText.Append("</ol>");
            //                flagForNumbered = false;
            //            }
            //        }
            //        else
            //        {
            //            paragraphText.Append("</ol>");
            //        }

            //        if (paragraph.ListFormat.ListType == ListType.Numbered)
            //        {
            //            paragraphText.Append($"<li>");
            //        }
            //        #endregion

            //        var children = paragraph.ChildObjects;

            //        for (int i = 0; i < children.Count; i++)
            //        {
            //            DocumentObject child = children[i];
            //            if (child.DocumentObjectType == DocumentObjectType.TextRange)
            //            {
            //                TextRange textRange = child as TextRange;
            //                fontName = textRange?.CharacterFormat.FontName;
            //                fontSize = textRange?.CharacterFormat.FontSize;
            //                paragraphText.Append(textRange.Text);
            //            }
            //            else if (child.DocumentObjectType == DocumentObjectType.Field)
            //            {
            //                Field field = child as Field;
            //                if (field.Type == FieldType.FieldHyperlink & !string.IsNullOrWhiteSpace(field.FieldText))
            //                {
            //                    if (marking)
            //                    {
            //                        paragraphText.Append($"<a href='{field.Code}'>{field.FieldText}</a>");
            //                        i += 2;
            //                        continue;
            //                    }
            //                    else
            //                    {
            //                        paragraphText.Append($"{field.FieldText}");
            //                        i += 2;
            //                        continue;
            //                    }
            //                }
            //                else if (field.Type == FieldType.FieldRef & !string.IsNullOrWhiteSpace(field.FieldText))
            //                {
            //                    if (marking)
            //                    {
            //                        paragraphText.Append($"<strong>{field.FieldText}</strong>");
            //                        i += 2;
            //                        continue;
            //                    }
            //                    else
            //                    {
            //                        paragraphText.Append($"{field.FieldText}");
            //                        i += 2;
            //                        continue;
            //                    }

            //                }
            //            }
            //            else if (child.DocumentObjectType == DocumentObjectType.Break)
            //            {
            //                Break @break = child as Break;
            //                if (@break.BreakType == BreakType.LineBreak)
            //                {
            //                    paragraphText.Replace("\v", $"<br>");
            //                }
            //            }
            //            else if (child.DocumentObjectType == DocumentObjectType.Picture)
            //            {
            //                DocPicture picture = child as DocPicture;
            //                paragraphText.Append($"<img width='{picture.Width}px' height='{picture.Height}px' src=\"data:image/jpeg;base64," + Convert.ToBase64String(picture.ImageBytes) + "\" />");
            //            }
            //        }

            //        if (paragraph.ListFormat.ListType == ListType.Numbered)
            //        {
            //            paragraphText.Append($"</li>");
            //        }

            //        fontName = string.IsNullOrWhiteSpace(fontName) ? "Time New Roman" : fontName;
            //        fontSize = fontSize == default ? 12 : fontSize;

            //        longText.AppendLine($"<p align='{aligment}'><font size='{fontSize}' face='{fontName}'>{paragraphText}</font></p>");
            //    }
            //    for (int i = 0; i < 5; i++)
            //    {
            //        longText.AppendLine("<br>");
            //    }
            //    longText.AppendLine("</div>");
            //}

            //return longText.ToString(); 
            #endregion
        }
        #region old method
        //private string GetAligment(Paragraph paragraph)
        //{
        //    if (paragraph.Format.HorizontalAlignment == HorizontalAlignment.Left)
        //    {
        //        return "left";
        //    }
        //    else if (paragraph.Format.HorizontalAlignment == HorizontalAlignment.Right)
        //    {
        //        return "right";
        //    }
        //    else if (paragraph.Format.HorizontalAlignment == HorizontalAlignment.Center)
        //    {
        //        return "center";
        //    }
        //    else if (paragraph.Format.HorizontalAlignment == HorizontalAlignment.Justify)
        //    {
        //        return "justify";
        //    }
        //    else
        //    {
        //        return "";
        //    }
        //} 
        #endregion
        /// <summary>
        /// 
        /// </summary>
        /// <param name="section"></param>
        /// <returns></returns>
        public List<Field> FindAllLinksBySection(Section section)
        {
            var links = new List<Field>();

            foreach (DocumentObject sec in section.Body.ChildObjects)
            //Parallel.ForEach<DocumentObjectCollection>(section.Body.ChildObjects, sec =>
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
            //);
            return links;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public IEnumerable<Field> GetAllHyperlinks()
        {
            var links = new HashSet<Field>();

            //foreach (Section section in document.Sections)
            Parallel.For(0, document.Sections.Count, i =>
            {
                Section section = document.Sections[i];
                var childObjects = section.Body.ChildObjects;
                //foreach (DocumentObject sec in section.Body.ChildObjects)
                Parallel.For(0, childObjects.Count, i =>
                {
                    DocumentObject documentObject = childObjects[i];
                    if (documentObject.DocumentObjectType == DocumentObjectType.Paragraph)
                    {
                        var paragraphs = (documentObject as Paragraph).ChildObjects;
                        //foreach (DocumentObject para in (sec as Paragraph).ChildObjects)
                        Parallel.For(0, paragraphs.Count, i =>
                        {
                            DocumentObject para = paragraphs[i];
                            if (para.DocumentObjectType == DocumentObjectType.Field)
                            {
                                Field field = para as Field;

                                if (field.Type == FieldType.FieldHyperlink)
                                {
                                    links.Add(field);
                                }
                            }
                        });
                    }
                });
            });

            return links.ToList();
        }
        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public IEnumerable<Bookmark> GetAllBookmarks()
        {
            HashSet<Bookmark> hashSetBookmarks = new HashSet<Bookmark>();
            BookmarksNavigator bookmarksNavigator = new BookmarksNavigator(document);
            var bookmarks = bookmarksNavigator.Document.Bookmarks;

            //for (int i = 0; i < bookmarks.Count; i++)
            Parallel.For(0, bookmarks.Count, i =>
            {
                Bookmark bookmark = bookmarks[i];
                if (bookmark.Name != badBookmark)
                {
                    try
                    {
                        hashSetBookmarks.Add(bookmark);
                    }
                    catch (IndexOutOfRangeException e)
                    {
                        loggerException.Error($"При поиске закладки была ошибка: {e.Message}");
                    }
                }
                else
                {
                    return;
                }
            });

            return hashSetBookmarks.Where(b => b != null).ToList();
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="bookmarkText"></param>
        /// <param name="text"></param>
        public void EditTextInBookmark(string bookmarkText, string text)
        {
            BookmarksNavigator bookmarksNavigator = new BookmarksNavigator(document);


            Section tempSection = document.AddSection();
            tempSection.AddParagraph().AppendText(text);

            ParagraphBase paragraphBaseFirstItem = tempSection.Paragraphs[0].Items.FirstItem as ParagraphBase;
            ParagraphBase paragraphBaseLastItem = tempSection.Paragraphs[tempSection.Paragraphs.Count - 1].Items.LastItem as ParagraphBase;
            TextBodySelection textBodySelection = new TextBodySelection(paragraphBaseFirstItem, paragraphBaseLastItem);
            TextBodyPart textBodyPart = new TextBodyPart(textBodySelection);

            bookmarksNavigator.MoveToBookmark(TransformWordWithUnderline(bookmarkText));
            bookmarksNavigator.ReplaceBookmarkContent(textBodyPart);

            document.Sections.Remove(tempSection);

            SaveCurrentDocument();

        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="field"></param>
        /// <param name="hyperlink"></param>
        public void EditLinkInHypertext(Field field, string hyperlink)
        {
            field.Code = "HYPERLINK \"" + hyperlink + "\"";

            SaveCurrentDocument();
        }
        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public List<Tuple<string, string>> GetAllFootnotes()
        {
            var result = new List<Tuple<string, string>>();

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
                        //int countChildObjects = footnote.OwnerParagraph.ChildObjects.Count;
                        //while (countChildObjects % 2 != 0)
                        //{
                        //    countChildObjects--;
                        //}
                        //for (int i = 0; i < countChildObjects; i++)
                        //{
                        //    DocumentObject child = footnote.OwnerParagraph.ChildObjects[i];
                        //    if (child.DocumentObjectType is DocumentObjectType.TextRange)
                        //    {
                        //        TextRange textRange = child as TextRange;
                        //        //string[] splitTextRange = textRange.
                        //    }
                        //    else
                        //    {
                        //        continue;
                        //    }
                        //    StringBuilder innerText = new StringBuilder();
                        //    foreach (Paragraph item in footnote.TextBody.ChildObjects)
                        //    {
                        //        innerText.Append(item.Text);
                        //    }

                        //    result.Add(new Tuple<string, string>(footnote.OwnerParagraph.Text, innerText.ToString()));
                        //}
                        foreach (Footnote footnote in footnotes)
                        {
                            StringBuilder innerText = new StringBuilder();
                            foreach (Paragraph item in footnote.TextBody.ChildObjects)
                            {
                                innerText.Append(item.Text);
                            }

                            result.Add(new Tuple<string, string>(footnote.OwnerParagraph.Text, innerText.ToString()));
                        }
                    }
                }
            }

            return result;
        }
        //private string GetWordFromFootnoteForHyperlink(Paragraph footnote)
        //{
        //    int countChildObjects = footnote.ChildObjects.Count;
        //    while (countChildObjects % 2 != 0)
        //    {
        //        countChildObjects--;
        //    }

        //    return null;
        //}
        /// <summary>
        /// 
        /// </summary>
        public void CreateReferencesSection()
        {
            var sectionAndParagraph = GetSectionAndParagraphByStyleName(styleName);
            if (sectionAndParagraph == null)
            {
                Section sectionForReferences = document.Document.AddSection();
                var paragraphForReferences = sectionForReferences.AddParagraph();

                AddParagraphStyle(paragraphForReferences);

                paragraphForReferences.AppendText("Сноски");
                paragraphForReferences.AppendBreak(BreakType.LineBreak);
                //chapterForReferences.ApplyStyle(BuiltinStyle.Heading1);
                //chapterForReferences.ApplyStyle(BuiltinStyle.Title);

                //referencesParagraph = sectionForReferences.AddParagraph();
                referencesParagraph = paragraphForReferences;
                ReferencesSection = sectionForReferences;
                IndexReferencesSection = Document.GetIndex(sectionForReferences);
            }
            else
            {
                referencesParagraph = sectionAndParagraph.Item2;
                ReferencesSection = sectionAndParagraph.Item1;
                IndexReferencesSection = Document.GetIndex(sectionAndParagraph.Item1);
            }

            SaveCurrentDocument();
        }

        private void AddParagraphStyle(Paragraph paragraphForReferences)
        {
            ParagraphStyle referenceParagraphStyle = new ParagraphStyle(document)
            {
                Name = styleName
            };

            referenceParagraphStyle.CharacterFormat.Bold = true;
            referenceParagraphStyle.CharacterFormat.FontSize = 20;
            referenceParagraphStyle.CharacterFormat.FontName = "Calibri";

            document.Styles.Add(referenceParagraphStyle);

            paragraphForReferences.ApplyStyle(referenceParagraphStyle.Name);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="field"></param>
        public void DeleteHyperlink(Field field)
        {
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

            SaveCurrentDocument();
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
        /// <summary>
        /// 
        /// </summary>
        /// <param name="bookmarkText"></param>
        public void DeleteBookmark(string bookmarkText)
        {
            BookmarksNavigator bookmarksNavigator = new BookmarksNavigator(document);
            bookmarksNavigator.MoveToBookmark(TransformWordWithUnderline(bookmarkText));

            if (bookmarksNavigator.CurrentBookmark != null)
            {
                var paragraphs = Document.Sections[IndexReferencesSection].Paragraphs;
                foreach (var paragraph in from Paragraph paragraph in paragraphs
                                          where paragraph.FirstChild is BookmarkStart & paragraph.LastChild is BookmarkEnd
                                          let paragraphFirstChild = paragraph.FirstChild as BookmarkStart
                                          let paragraphLastChild = paragraph.LastChild as BookmarkEnd
                                          where paragraphFirstChild.Name == bookmarkText & paragraphLastChild.Name == bookmarkText
                                          select paragraph)
                {
                    paragraphs.Remove(paragraph);
                }
                document.Bookmarks.Remove(bookmarksNavigator.CurrentBookmark);
            }

            foreach (Section section in Document.Sections)
            {
                foreach (Paragraph paragraph in section.Paragraphs)
                {
                    foreach (DocumentObject child in paragraph.ChildObjects)
                    {
                        if (child.DocumentObjectType is DocumentObjectType.Field)
                        {
                            Field field = child as Field;

                            if (field.Code == $"REF {bookmarkText} \\p \\h")
                            {
                                var textRange = field.NextSibling.NextSibling as TextRange;
                                var nextFieldMark = textRange.NextSibling as FieldMark;
                                var previousFieldMark = textRange.PreviousSibling as FieldMark;
                                textRange.CharacterFormat.TextColor = Color.Black;
                                textRange.CharacterFormat.UnderlineStyle = UnderlineStyle.None;

                                paragraph.ChildObjects.Remove(nextFieldMark);
                                paragraph.ChildObjects.Remove(previousFieldMark);
                                paragraph.ChildObjects.Remove(child);
                                break;
                            }
                        }
                    }
                }
            }

            SaveCurrentDocument();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public IEnumerable<DocPicture> GetImages()
        {          
            foreach (Section section in document.Sections)
            {
                //Get Each Paragraph of Section
                foreach (Paragraph paragraph in section.Paragraphs)
                {
                    //Get Each Document Object of Paragraph Items
                    foreach (DocumentObject docObject in paragraph.ChildObjects)
                    {
                        //If Type of Document Object is Picture, Extract.
                        if (docObject.DocumentObjectType == DocumentObjectType.Picture)
                        {
                            yield return docObject as DocPicture;                          
                        }
                    }
                }
            }
        }
    } 
}
