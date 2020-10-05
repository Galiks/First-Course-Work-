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
using System.IO;
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

        //private const string badBookmark = "_GoBack";
        private const int width = 470;
        private const int height = 340;
        private readonly string filename;
        //private string referencesWord;
        private static int indexNextField = 0;
        private Paragraph referencesParagraph;
        private static Section referencesSection;
        private const string styleName = "MyReferences";
        private int countOfAddedLink;
        private int usersCountOfAddedWord;
        private bool isStopCreateLink = false;
        /// <summary>
        /// 
        /// </summary>
        public Document Document { get; }
        /// <summary>
        /// 
        /// </summary>
        public HashSet<string> Messages { get; }
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
            try
            {
                //loggerUser.Info($"Начата работа с файлом: {filename}");
                this.filename = filename;
                Document = new Document();
                //проблема с дублями файла
                filepath = filename;
                Document.LoadFromFile(filename);
                Messages = new HashSet<string>();
                CreateReferencesSection();
            }
            catch (Exception e)
            {
                loggerException.Error($"Message {e.Message}");
                WriteExceptionInLog(e);
                throw e;
            }
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
        private string GetReferencesWord(string word)
        {
            //­­­U+00AD
            //005F       
            if (!string.IsNullOrWhiteSpace(word))
            {
                return TransformWord(word);
            }
            else
            {
                loggerException.Error($"Передаваемый параметр является null");
                return null;
            }
        }
        private string TransformWord(string word)
        {
            if (!string.IsNullOrWhiteSpace(word))
            {
                var splitWord = word.Split("\\");
                string result = splitWord[^1].Replace(' ', '\u005F').Replace('/', '\u005F');
                if (result.Length > 20)
                {
                    result = result.Substring(0, 17) + Path.GetExtension(result);
                    return result;
                }
                else
                {
                    return result;
                }
            }
            else
            {
                loggerException.Error($"Передаваемый параметр является null");
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
            try
            {
                SetCountOfAddedWord(0);

                SetUsersCountOfAddedWord(count);

                List<string> nouns = GetWordsByCases(word);

                for (int i = 0; i < nouns.Count; i++)
                {
                    string noun = nouns[i];
                    if (i == 0)
                    {
                        try
                        {
                            if (!isStopCreateLink)
                            {
                                CreateBookmarkByWord(noun, text);
                            }
                            else
                            {
                                break;
                            }
                        }
                        catch (Exception e)
                        {
                            loggerException.Error($"Message {e.Message}");
                            WriteExceptionInLog(e);
                            throw e;
                        }
                    }
                    else
                    {
                        try
                        {
                            if (!isStopCreateLink)
                            {
                                CreateBookmarkByWord(noun, text);
                            }
                            else
                            {
                                break;
                            }
                        }
                        catch (Exception e)
                        {
                            loggerException.Error($"Message {e.Message}");
                            WriteExceptionInLog(e);
                            throw e;
                        }
                    }
                }
            }
            catch (CyrWordNotFoundException e)
            {
                loggerException.Error($"Message {e.Message}");
                WriteExceptionInLog(e);
                try
                {
                    if (!isStopCreateLink)
                    {
                        CreateBookmarkByWord(word, text);
                    }
                    else
                    {
                        return;
                    }
                }
                catch (Exception secondError)
                {
                    loggerException.Error($"Message {secondError.Message}");
                    WriteExceptionInLog(secondError);
                    throw secondError;
                }
            }
            finally
            {
                if (IndexNextField.Equals(default))
                {
                    IncreaseOfTwoindexNextField();
                }
            }
        }

        private void SetCountOfAddedWord(int count)
        {
            countOfAddedLink = count;
        }

        private void SetUsersCountOfAddedWord(int count)
        {
            if (count == default)
            {
                usersCountOfAddedWord = default;
            }
            else
            {
                usersCountOfAddedWord = count;
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
                    string wordWithFirstLetterUpper = GetWordWithFirstLetterUpper(nouns[i]);
                    if (!string.IsNullOrWhiteSpace(wordWithFirstLetterUpper))
                    {
                        nouns.Add(wordWithFirstLetterUpper);
                    }
                }
            }

            return nouns;
        }
        private void CreateBookmarkByWord(string word, string text = null)
        {
            try
            {
                string referencesWord = GetReferencesWord(text);
                if (string.IsNullOrWhiteSpace(referencesWord))
                {
                    loggerException.Error("Слово для закладки является NULL.");
                    return;
                }
                //Create bookmark objects
                //BookmarkStart start = new BookmarkStart(document, referencesWord);
                //BookmarkEnd end = new BookmarkEnd(document, referencesWord);

                AppendTextBookmark(text, referencesWord);

                TextSelection[] textSelection = Document.FindAllString(word, true, true);



                if (textSelection == null || textSelection.Length == 0)
                {
                    loggerException.Error($"Не было найдено слов");
                    return;
                }

                //Get the keywords
                for (int i = 0; i < textSelection.Length; i++)
                {
                    TextSelection keywordOne = textSelection[i];
                    try
                    {
                        if (usersCountOfAddedWord != default)
                        {
                            if (usersCountOfAddedWord > countOfAddedLink)
                            {
                                CreateBookmark(keywordOne, referencesWord);
                                countOfAddedLink++;
                            }
                            else
                            {
                                isStopCreateLink = true;
                                break;
                            }
                        }
                        else
                        {
                            CreateBookmark(keywordOne, referencesWord);
                        }

                    }
                    catch (Exception e)
                    {
                        loggerException.Error($"Message {e.Message}");
                        WriteExceptionInLog(e);
                        throw e;
                    }
                }

                SaveCurrentDocument();
            }
            catch (Exception e)
            {
                loggerException.Error($"Message {e.Message}");
                WriteExceptionInLog(e);
                throw e;
            }


            
        }

        private void AppendTextBookmark(string text, string referencesWord)
        {
            try
            {
                BookmarksNavigator bn = new BookmarksNavigator(Document);
                bn.MoveToBookmark(referencesWord, true, true);
                if (bn.CurrentBookmark == null)
                {
                    referencesParagraph = ReferencesSection.AddParagraph();
                    referencesParagraph.AppendBookmarkStart(referencesWord);
                    referencesParagraph.AppendText(text);
                    referencesParagraph.AppendBookmarkEnd(referencesWord);
                }
            }
            catch (Exception e)
            {
                loggerException.Error($"Message {e.Message}");
                WriteExceptionInLog(e);
                throw e;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="word"></param>
        /// <param name="text"></param>
        public void CreateBookmarkForOneWord(string word, string text)
        {
            string referencesWord = GetReferencesWord(text);
            try
            {
                AppendTextBookmark(text, referencesWord);
            }
            catch (Exception e)
            {
                loggerException.Error($"Message {e.Message}");
                WriteExceptionInLog(e);
                throw e;
            }
            TextSelection keyword = Document.FindString(word, true, true);
            try
            {
                CreateBookmark(keyword, referencesWord);
            }
            catch (Exception e)
            {
                loggerException.Error($"Message {e.Message}");
                WriteExceptionInLog(e);
                throw e;
            }

            Paragraph backlinkParagraph = keyword.GetAsOneRange().OwnerParagraph;

            BookmarksNavigator bn = new BookmarksNavigator(Document);
            bn.MoveToBookmark(referencesWord, true, true);
            if (bn.CurrentBookmark == null)
            {
                backlinkParagraph.AppendBookmarkStart(word);
                backlinkParagraph.AppendBookmarkEnd(word);

                SaveCurrentDocument();
            }

            keyword = Document.FindAllString(text, true, true).Last();
            referencesWord = GetReferencesWord(word);
            try
            {
                CreateBookmark(keyword, referencesWord, false);
            }
            catch (Exception e)
            {
                loggerException.Error($"Message {e.Message}");
                WriteExceptionInLog(e);
                throw e;
            }

        }

        private void CreateBookmark(TextSelection keywordOne, string referencesWord, bool isCheckField = true)
        {
            try
            {
                TextRange tr = null;

                //Get the textrange its locates
                try
                {
                    tr = keywordOne.GetAsOneRange();
                }
                catch (Exception e)
                {
                    loggerException.Error($"Message {e.Message}");
                    WriteExceptionInLog(e);
                    try
                    {
                        CreateBookmark(keywordOne, referencesWord);
                    }
                    catch (Exception secondError)
                    {
                        loggerException.Error($"Message {secondError.Message}");
                        WriteExceptionInLog(secondError);
                        throw secondError;
                    }
                    throw e;
                }

                //Set the formatting
                tr.CharacterFormat.UnderlineStyle = UnderlineStyle.Single;
                tr.CharacterFormat.TextColor = Color.Blue;

                //Get the paragraph it locates
                Paragraph paragraph = tr.OwnerParagraph;

                if (isCheckField)
                {
                    if (paragraph.Equals(referencesParagraph))
                    {
                        loggerException.Error("Параграф равен параграфу для сносок. Остановка поиска слов");
                        return;
                    }
                }

                //Get the index of the keyword in its paragraph
                int index = paragraph.ChildObjects.IndexOf(tr);

                DocumentObject child = paragraph.ChildObjects[index];

                if (isCheckField)
                {
                    if (child.DocumentObjectType == DocumentObjectType.Field)
                    {
                        Field textField = child as Field;

                        if (textField.Type == FieldType.FieldRef)
                        {
                            loggerException.Error($"Поле {textField.FieldText} не было добавлено, так как оно уже имеет тип {FieldType.FieldRef}");
                            Messages.Add($"Поле {textField.FieldText} не было добавлено, так как оно уже имеет тип {FieldType.FieldRef}");
                            return;
                        }
                        else if (textField.Type == FieldType.FieldHyperlink)
                        {
                            loggerException.Error($"Поле {textField.FieldText} не было добавлено, так как оно уже имеет тип {FieldType.FieldHyperlink}");
                            Messages.Add($"Поле {textField.FieldText} не было добавлено, так как оно уже имеет тип {FieldType.FieldHyperlink}");
                            return;
                        }
                    }
                }

                //Create a cross-reference field, and link it to bookmark                   
                Field field = new Field(Document)
                {
                    Type = FieldType.FieldRef
                };
                string code = $@"REF {referencesWord} \p \h";
                field.Code = code;

                //Insert field
                paragraph.ChildObjects.Insert(index, field);

                //Insert FieldSeparator object
                FieldMark fieldSeparator = new FieldMark(Document, FieldMarkType.FieldSeparator);
                paragraph.ChildObjects.Insert(index + 1, fieldSeparator);

                //Insert FieldEnd object to mark the end of the field
                FieldMark fieldEnd = new FieldMark(Document, FieldMarkType.FieldEnd);
                paragraph.ChildObjects.Insert(index + 3, fieldEnd);

                SaveCurrentDocument();
            }
            catch (Exception e)
            {
                loggerException.Error($"Message {e.Message}");
                WriteExceptionInLog(e);
                throw e;
            }
        }

        private void CreateBookmark(Paragraph paragraph, string referencesWord, bool isCheckField = true)
        {


            try
            {
                if (isCheckField)
                {
                    if (paragraph.Equals(referencesParagraph))
                    {
                        loggerException.Error("Параграф равен параграфу для сносок. Остановка поиска слов");
                        return;
                    }
                }

                //Get the index of the keyword in its paragraph
                int index = -1;
                DocumentObject child = null;

                for (int i = 0; i < paragraph.ChildObjects.Count; i++)
                {
                    DocumentObject documentObject = paragraph.ChildObjects[i];
                    if (documentObject.DocumentObjectType == DocumentObjectType.Picture)
                    {
                        index = i;
                        child = documentObject;
                        break;
                    }

                }

                if (index == -1 || child is null)
                {
                    return;
                }

                //DocumentObject child = paragraph.ChildObjects[index];

                if (isCheckField)
                {
                    if (child.DocumentObjectType == DocumentObjectType.Field)
                    {
                        Field textField = child as Field;

                        if (textField.Type == FieldType.FieldRef)
                        {
                            loggerException.Error($"Поле {textField.FieldText} не было добавлено, так как оно уже имеет тип {FieldType.FieldRef}");
                            Messages.Add($"Поле {textField.FieldText} не было добавлено, так как оно уже имеет тип {FieldType.FieldRef}");
                            return;
                        }
                        else if (textField.Type == FieldType.FieldHyperlink)
                        {
                            loggerException.Error($"Поле {textField.FieldText} не было добавлено, так как оно уже имеет тип {FieldType.FieldHyperlink}");
                            Messages.Add($"Поле {textField.FieldText} не было добавлено, так как оно уже имеет тип {FieldType.FieldHyperlink}");
                            return;
                        }
                    }
                }

                //Create a cross-reference field, and link it to bookmark                   
                Field field = new Field(Document)
                {
                    Type = FieldType.FieldRef
                };
                string code = $@"REF {referencesWord} \p \h";
                field.Code = code;

                //Insert field
                paragraph.ChildObjects.Insert(index, field);

                //Insert FieldSeparator object
                FieldMark fieldSeparator = new FieldMark(Document, FieldMarkType.FieldSeparator);
                paragraph.ChildObjects.Insert(index + 1, fieldSeparator);

                //Insert FieldEnd object to mark the end of the field
                FieldMark fieldEnd = new FieldMark(Document, FieldMarkType.FieldEnd);
                paragraph.ChildObjects.Insert(index + 3, fieldEnd);

                SaveCurrentDocument();
            }
            catch (Exception e)
            {
                loggerException.Error($"Message {e.Message}");
                WriteExceptionInLog(e);
                throw e;
            }
           
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="path"></param>
        /// <param name="word"></param>
        /// <param name="count"></param>
        public void CreateBookmarksForImage(string path, string word, int count = default)
        {
            try
            {
                SetCountOfAddedWord(0);

                SetUsersCountOfAddedWord(count);

                List<string> nouns = GetWordsByCases(word);

                for (int i = 0; i < nouns.Count; i++)
                {
                    string noun = nouns[i];

                    try
                    {
                        if (!isStopCreateLink)
                        {
                            CreateBookmarkByImage(path, noun);
                        }
                    }
                    catch (Exception e)
                    {
                        loggerException.Error($"Message {e.Message}");
                        WriteExceptionInLog(e);
                        throw e;
                    }
                    //SetBookmarkForImage(path);
                }
            }
            catch (Exception e)
            {
                loggerException.Error($"Message {e.Message}");
                WriteExceptionInLog(e);
                try
                {
                    if (!isStopCreateLink)
                    {
                        CreateBookmarkByImage(path, word);
                    }
                }
                catch (Exception secondError)
                {
                    loggerException.Error($"Message {secondError.Message}");
                    WriteExceptionInLog(secondError);
                    throw secondError;
                }
                throw e;
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

        private void CreateBookmarkByImage(string path, string word)
        {
            try
            {
                string referencesWord = GetReferencesWord(path);
                AppendImageBookmark(path, referencesWord);

                //Find the keyword "Hypertext"
                TextSelection[] text = Document.FindAllString(word, true, true);

                if (text == null || text.Length == 0)
                {
                    loggerException.Error($"Не было найдено слов");
                    return;
                }

                //Get the keywords
                for (int i = 0; i < text.Length; i++)
                {
                    TextSelection keywordOne = text[i];
                    try
                    {
                        if (usersCountOfAddedWord != default)
                        {
                            if (usersCountOfAddedWord > countOfAddedLink)
                            {
                                CreateBookmark(keywordOne, referencesWord);
                                countOfAddedLink++;
                            }
                            else
                            {
                                isStopCreateLink = true;
                                break;
                            }
                        }
                        else
                        {
                            CreateBookmark(keywordOne, referencesWord);
                        }
                    }
                    catch (Exception e)
                    {
                        WriteExceptionInLog(e);
                        throw e;
                    }
                }

                SaveCurrentDocument();
            }
            catch (Exception e)
            {
                loggerException.Error($"Message {e.Message}");
                WriteExceptionInLog(e);
                throw e;
            }
        }

        private Paragraph AppendImageBookmark(string path, string referencesWord)
        {
            try
            {
                var para = referencesSection.AddParagraph();
                BookmarksNavigator bn = new BookmarksNavigator(Document);
                bn.MoveToBookmark(referencesWord, true, true);

                if (bn.CurrentBookmark == null)
                {

                    para.AppendBookmarkStart(referencesWord);
                    DocPicture picture = para.AppendPicture(path);
                    picture.Width = width;
                    picture.Height = height;
                    para.AppendBookmarkEnd(referencesWord);


                    SaveCurrentDocument();


                }
                else
                {
                    throw new Exception("Подобная закладка уже есть!");
                }

                return para;
            }
            catch (Exception e)
            {
                loggerException.Error($"Message {e.Message}");
                WriteExceptionInLog(e);
                throw e;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="path"></param>
        /// <param name="word"></param>
        public void CreateBookmarkByImgeForOneWord(string path, string word)
        {
            try
            {
                string referencesWord = GetReferencesWord(path);
                Paragraph imageParagraph = AppendImageBookmark(path, referencesWord);

                TextSelection keyword = Document.FindString(word, true, true);

                CreateBookmark(keyword, referencesWord);

                Paragraph backlinkParagraph = keyword.GetAsOneRange().OwnerParagraph;

                referencesWord = GetReferencesWord(word);

                BookmarksNavigator bn = new BookmarksNavigator(Document);
                bn.MoveToBookmark(referencesWord, true, true);
                if (bn.CurrentBookmark == null)
                {
                    backlinkParagraph.AppendBookmarkStart(referencesWord);
                    backlinkParagraph.AppendBookmarkEnd(referencesWord);

                    SaveCurrentDocument();
                }

                //keyword = Document.FindAllString(referencesWord, true, true).Last();
                CreateBookmark(imageParagraph, referencesWord, false);
            }
            catch (Exception e)
            {
                loggerException.Error($"Message {e.Message}");
                WriteExceptionInLog(e);
                throw e;
            }
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
                SetCountOfAddedWord(0);

                SetUsersCountOfAddedWord(count);

                List<string> nouns = GetWordsByCases(word);

                //int lengthOfFindedWords = count == default ? nouns.Count : count;

                for (int i = 0; i < nouns.Count; i++)
                {
                    string noun = nouns[i];
                    try
                    {
                        if (!isStopCreateLink)
                        {
                            CreateHyperlinkByWord(noun, text);
                        }
                    }
                    catch (Exception e)
                    {
                        loggerException.Error($"Message {e.Message}");
                        WriteExceptionInLog(e);
                        throw e;
                    }
                }
            }
            catch (Exception e)
            {
                loggerException.Error($"Message {e.Message}");
                WriteExceptionInLog(e);
                try
                {
                    if (!isStopCreateLink)
                    {
                        CreateHyperlinkByWord(word, text);
                    }
                }
                catch (Exception secondError)
                {
                    loggerException.Error($"Message {secondError.Message}");
                    WriteExceptionInLog(secondError);
                    throw secondError;
                }
                throw e;
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
            try
            {
                TextSelection[] text = Document.FindAllString(word, true, true);

                if (text == null || text.Length == 0)
                {
                    loggerException.Error($"Не было найдено слов");
                    return;
                }

                for (int i = 0; i < text.Length; i++)
                {
                    TextSelection seletion = text[i];

                    //Get the text range

                    if (usersCountOfAddedWord != default)
                    {
                        if (usersCountOfAddedWord > countOfAddedLink)
                        {
                            CreateHyperlink(hyperlink, seletion);
                            countOfAddedLink++;
                        }
                        else
                        {
                            isStopCreateLink = true;
                            break;
                        }
                    }
                    else
                    {
                        CreateHyperlink(hyperlink, seletion);
                    }
                }
                SaveCurrentDocument();
            }
            catch (Exception e)
            {
                loggerException.Error($"Message {e.Message}");
                WriteExceptionInLog(e);
                throw e;
            }
            #region Parallel version
            //ParallelOptions po = new ParallelOptions();

            //try
            //{
            //    Parallel.For(0, text.Length, i =>
            //        {
            //            TextSelection seletion = text[i];

            //    //Get the text range

            //    if (usersCountOfAddedWord != default)
            //            {
            //                if (usersCountOfAddedWord > countOfAddedLink)
            //                {
            //                    CreateHyperlink(hyperlink, seletion);
            //                    countOfAddedLink++;
            //                }
            //                else
            //                {
            //                    isStopCreateLink = true;
            //                    po.CancellationToken.ThrowIfCancellationRequested();
            //                }
            //            }
            //            else
            //            {
            //                CreateHyperlink(hyperlink, seletion);
            //            }

            //        });
            //}
            //catch (OperationCanceledException)
            //{
            //    return;
            //} 
            #endregion  
        }
        /// <summary>
        /// 
        /// </summary>
        public void CreateHyperlinkForOneWord(string word, string hyperlink)
        {
            try
            {
                TextSelection keyword = Document.FindString(word, true, true);
                CreateHyperlink(hyperlink, keyword);
            }
            catch (Exception e)
            {
                loggerException.Error($"Message {e.Message}");
                WriteExceptionInLog(e);
                throw e;
            }
        }
        private void CreateHyperlink(string hyperlink, TextSelection seletion)
        {
            TextRange tr = seletion.GetAsOneRange();

            int index = tr.OwnerParagraph.ChildObjects.IndexOf(tr) - IndexNextField;
            Paragraph paragraph = tr.OwnerParagraph;

            DocumentObject child;
            try
            {
                child = paragraph.ChildObjects[index];
            }
            catch (IndexOutOfRangeException e)
            {
                WriteExceptionInLog(e);
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


            Field field = new Field(Document)
            {
                Code = "HYPERLINK \"" + hyperlink + "\"",

                Type = FieldType.FieldHyperlink
            };

            tr.OwnerParagraph.ChildObjects.Insert(index, field);

            FieldMark fm = new FieldMark(Document, FieldMarkType.FieldSeparator);

            tr.OwnerParagraph.ChildObjects.Insert(index + 1, fm);

            //Set character format

            tr.CharacterFormat.TextColor = Color.Blue;

            tr.CharacterFormat.UnderlineStyle = UnderlineStyle.Single;

            tr.CharacterFormat.Bold = tr.CharacterFormat.Bold;

            FieldMark fmend = new FieldMark(Document, FieldMarkType.FieldEnd);

            tr.OwnerParagraph.ChildObjects.Insert(index + 3, fmend);

            field.End = fmend;
        }

        /// <summary>
        /// Добавляет гиперссылку на изображение, добавляемое в документ
        /// </summary>
        /// <param name="path">путь к изображению</param>
        /// <param name="hyperlink">гиперссылка</param>
        public void CreatHyperlinkForImage(string path, string hyperlink)
        {
            try
            {
                DocPicture picture = new DocPicture(Document);
                picture.LoadImage(path);
                picture.Width = 470;
                picture.Height = 340;

                ReferencesSection.AddParagraph().AppendHyperlink(hyperlink, picture, HyperlinkType.WebLink);
                SaveCurrentDocument();
            }
            catch (Exception e)
            {
                loggerException.Error($"Message {e.Message}");
                WriteExceptionInLog(e);
                throw e;
            }
        }

        private Paragraph GetParagraphByWord(string word)
        {
            foreach (Section section in Document.Sections)
            {
                foreach (Paragraph paragraph in section.Paragraphs)
                {
                    if (paragraph.Text.Contains(word))
                    {
                        return paragraph;
                    }
                }
            }
            return null;
        }

        /// <summary>
        /// Добавляет гиперссылку на имеющееся в документе изображение
        /// </summary>
        /// <param name="picture"></param>
        /// <param name="hyperlink"></param>
        public void CreateHyperlinkForImage(DocPicture picture, string hyperlink)
        {
            try
            {
                picture.OwnerParagraph.AppendHyperlink(hyperlink, picture, HyperlinkType.WebLink);

                SaveCurrentDocument();
            }
            catch (Exception e)
            {
                loggerException.Error($"Message {e.Message}");
                WriteExceptionInLog(e);
                throw e;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="style"></param>
        /// <returns></returns>
        public Tuple<Section, Paragraph> GetSectionAndParagraphByStyleName(string style)
        {
            try
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
            catch (Exception e)
            {
                loggerException.Error($"Message {e.Message}");
                WriteExceptionInLog(e);
                throw e;
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        public static string GetWordWithFirstLetterUpper(string str)
        {
            if (!string.IsNullOrWhiteSpace(str))
            {
                return str[0].ToString().ToUpper() + str.Substring(1);
            }
            else
            {
                return null;
            }
        }
        /// <summary>
        /// 
        /// </summary>
        public void SaveCurrentDocument()
        {
            try
            {
                Document.SaveToFile(filename);
            }
            catch (Exception e)
            {
                loggerException.Error($"Message {e.Message}");
                WriteExceptionInLog(e);
                throw e;
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public string GetTextFromDocument()
        {
            try
            {
                string filename = Conversion.Conversion.ConvertToHtml(filepath);
                string[] splitFilename = filename.Split("\\");
                string pathToHtmlFile = null;

                //for (int i = 0; i < splitFilename.Length; i++)
                Parallel.For(0, splitFilename.Length, i =>
                {
                    if (splitFilename[i].Equals("wwwroot"))
                    {
                        pathToHtmlFile = @"\" + string.Join(@"\", splitFilename, i + 1, splitFilename.Length - i - 1);
                    //return splitFilename.Join("\\", splitFilename, i, splitFilename.Length - 1);
                }
                });

                if (string.IsNullOrWhiteSpace(pathToHtmlFile))
                {
                    loggerException.Error("Путь к HTML файлу не создан.");
                    return null;
                }
                else
                {
                    return pathToHtmlFile;
                }

            }
            catch (Exception e)
            {
                loggerException.Error($"Message {e.Message}");
                WriteExceptionInLog(e);
                throw e;
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
            try
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
            catch (Exception e)
            {
                loggerException.Error($"Message {e.Message}");
                WriteExceptionInLog(e);
                throw e;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public IEnumerable<Field> GetHyperlinks()
        {
            try
            {
                var links = new HashSet<Field>();

                foreach (Section section in Document.Sections)
                //Parallel.For(0, Document.Sections.Count, i =>
                {

                    var childObjects = section.Body.ChildObjects;
                    foreach (DocumentObject sec in section.Body.ChildObjects)
                    //Parallel.For(0, childObjects.Count, i =>
                    {
                        if (sec.DocumentObjectType == DocumentObjectType.Paragraph)
                        {
                            foreach (DocumentObject para in (sec as Paragraph).ChildObjects)
                            //Parallel.For(0, paragraphs.Count, i =>
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

                return links.ToList();
            }
            catch (Exception e)
            {
                loggerException.Error($"Message {e.Message}");
                WriteExceptionInLog(e);
                throw e;
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public IEnumerable<Bookmark> GetBookmarks()
        {
            //HashSet<Bookmark> hashSetBookmarks = new HashSet<Bookmark>();
            BookmarksNavigator bookmarksNavigator = new BookmarksNavigator(Document);
            var bookmarks = bookmarksNavigator.Document.Bookmarks;

            foreach (Bookmark bookmark in bookmarks)
            {
                yield return bookmark;
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="bookmarkText"></param>
        /// <param name="text"></param>
        public void EditTextInBookmark(string bookmarkText, string text)
        {
            try
            {
                BookmarksNavigator bookmarksNavigator = new BookmarksNavigator(Document);


                Section tempSection = Document.AddSection();
                tempSection.AddParagraph().AppendText(text);

                ParagraphBase paragraphBaseFirstItem = tempSection.Paragraphs[0].Items.FirstItem as ParagraphBase;
                ParagraphBase paragraphBaseLastItem = tempSection.Paragraphs[tempSection.Paragraphs.Count - 1].Items.LastItem as ParagraphBase;
                TextBodySelection textBodySelection = new TextBodySelection(paragraphBaseFirstItem, paragraphBaseLastItem);
                TextBodyPart textBodyPart = new TextBodyPart(textBodySelection);

                bookmarksNavigator.MoveToBookmark(TransformWord(bookmarkText));
                bookmarksNavigator.ReplaceBookmarkContent(textBodyPart);

                Document.Sections.Remove(tempSection);

                SaveCurrentDocument();
            }
            catch (Exception e)
            {
                loggerException.Error($"Message {e.Message}");
                WriteExceptionInLog(e);
                throw e;
            }

        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="field"></param>
        /// <param name="hyperlink"></param>
        public void EditLinkInHypertext(Field field, string hyperlink)
        {
            try
            {
                field.Code = "HYPERLINK \"" + hyperlink + "\"";

                SaveCurrentDocument();
            }
            catch (Exception e)
            {
                loggerException.Error($"Message {e.Message}");
                WriteExceptionInLog(e);
                throw e;
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public List<Tuple<string, string>> GetFootnotes()
        {
            var result = new List<Tuple<string, string>>();

            foreach (Section section in Document.Sections)
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
            try
            {
                var sectionAndParagraph = GetSectionAndParagraphByStyleName(styleName);
                if (sectionAndParagraph == null)
                {
                    Section sectionForReferences = Document.Document.AddSection();
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
            catch (Exception e)
            {
                loggerException.Error($"Message {e.Message}");
                WriteExceptionInLog(e);
                throw e;
            }
        }

        private void AddParagraphStyle(Paragraph paragraphForReferences)
        {
            try
            {
                ParagraphStyle referenceParagraphStyle = new ParagraphStyle(Document)
                {
                    Name = styleName
                };

                referenceParagraphStyle.CharacterFormat.Bold = true;
                referenceParagraphStyle.CharacterFormat.FontSize = 20;
                referenceParagraphStyle.CharacterFormat.FontName = "Calibri";

                Document.Styles.Add(referenceParagraphStyle);

                paragraphForReferences.ApplyStyle(referenceParagraphStyle.Name);
            }
            catch (Exception e)
            {
                loggerException.Error($"Message {e.Message}");
                WriteExceptionInLog(e);
                throw e;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="field"></param>
        public void DeleteHyperlink(Field field)
        {
            try
            {
                int ownerParagraphIndex = field.OwnerParagraph.OwnerTextBody.ChildObjects.IndexOf(field.OwnerParagraph);
                int fieldIndex = field.OwnerParagraph.ChildObjects.IndexOf(field);
                Paragraph sepOwnerParagraph = field.Separator.OwnerParagraph;
                int sepOwnerParagraphIndex = field.Separator.OwnerParagraph.OwnerTextBody.ChildObjects.IndexOf(field.Separator.OwnerParagraph);
                int sepIndex = field.Separator.OwnerParagraph.ChildObjects.IndexOf(field.Separator);
                int endIndex = field.End.OwnerParagraph.ChildObjects.IndexOf(field.End);
                int endOwnerParagraphIndex = field.End.OwnerParagraph.OwnerTextBody.ChildObjects.IndexOf(field.End.OwnerParagraph);

                #region Remove font color and etc
                try
                {
                    FormatFieldResultText(field.Separator.OwnerParagraph.OwnerTextBody, sepOwnerParagraphIndex, endOwnerParagraphIndex, sepIndex, endIndex);
                }
                catch (Exception e)
                {
                    loggerException.Error($"Message {e.Message}");
                    WriteExceptionInLog(e);
                    throw e;
                }
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
            catch (Exception e)
            {
                loggerException.Error($"Message {e.Message}");
                WriteExceptionInLog(e);
                throw e;
            }
        }
        private void FormatFieldResultText(Body ownerBody, int sepOwnerParaIndex, int endOwnerParaIndex, int sepIndex, int endIndex)
        {
            try
            {
                for (int i = sepOwnerParaIndex; i <= endOwnerParaIndex; i++)
                {
                    Paragraph para = ownerBody.ChildObjects[i] as Paragraph;
                    if (i == sepOwnerParaIndex && i == endOwnerParaIndex)
                    {
                        for (int j = sepIndex + 1; j < endIndex; j++)
                        {
                            try
                            {
                                FormatText(para.ChildObjects[j] as TextRange);
                            }
                            catch (Exception e)
                            {
                                loggerException.Error($"Message {e.Message}");
                                WriteExceptionInLog(e);
                                throw e;
                            }
                        }

                    }
                    else if (i == sepOwnerParaIndex)
                    {
                        for (int j = sepIndex + 1; j < para.ChildObjects.Count; j++)
                        {

                            try
                            {
                                FormatText(para.ChildObjects[j] as TextRange);
                            }
                            catch (Exception e)
                            {
                                loggerException.Error($"Message {e.Message}");
                                WriteExceptionInLog(e);
                                throw e;
                            }
                        }
                    }
                    else if (i == endOwnerParaIndex)
                    {
                        for (int j = 0; j < endIndex; j++)
                        {

                            try
                            {
                                FormatText(para.ChildObjects[j] as TextRange);
                            }
                            catch (Exception e)
                            {
                                loggerException.Error($"Message {e.Message}");
                                WriteExceptionInLog(e);
                                throw e;
                            }
                        }
                    }
                    else
                    {
                        for (int j = 0; j < para.ChildObjects.Count; j++)
                        {

                            try
                            {
                                FormatText(para.ChildObjects[j] as TextRange);
                            }
                            catch (Exception e)
                            {
                                loggerException.Error($"Message {e.Message}");
                                WriteExceptionInLog(e);
                                throw e;
                            }
                        }
                    }
                }
            }
            catch (Exception e)
            {
                loggerException.Error($"Message {e.Message}");
                WriteExceptionInLog(e);
                throw e;
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
            try
            {
                BookmarksNavigator bookmarksNavigator = new BookmarksNavigator(Document);
                bookmarksNavigator.MoveToBookmark(TransformWord(bookmarkText));

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
                    Document.Bookmarks.Remove(bookmarksNavigator.CurrentBookmark);
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
            catch (Exception e)
            {
                loggerException.Error($"Message {e.Message}");
                WriteExceptionInLog(e);
                throw e;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public IEnumerable<DocPicture> GetImages()
        {
            foreach (Section section in Document.Sections)
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

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public static List<string> GetFileFormats()
        {
            return Enum.GetValues(typeof(FileFormat)).Cast<FileFormat>().Select(f => $".{f.ToString().ToLower()}").ToList();
        }

        private void WriteExceptionInLog(Exception e)
        {
            string innerException = null;
            if (e?.InnerException != null)
            {
                innerException = GetInnerException(e.InnerException);
            }
            loggerException.Error($"Message {e.Message} | {innerException}");
        }

        private string GetInnerException(Exception innerException)
        {
            return $"{Environment.NewLine} InnerException: {innerException.Message}";
        }
    }
}
