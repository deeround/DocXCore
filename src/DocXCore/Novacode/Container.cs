using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace Novacode
{
    public abstract class Container : DocXElement
    {
        public ContainerType ParentContainer;

        public virtual ReadOnlyCollection<Content> Contents
        {
            get
            {
                List<Content> contents = GetContents();
                return contents.AsReadOnly();
            }
        }

        public virtual ReadOnlyCollection<Paragraph> Paragraphs
        {
            get
            {
                List<Paragraph> paragraphs = GetParagraphs();
                foreach (Paragraph item in paragraphs)
                {
                    if (item.Xml.ElementsAfterSelf().FirstOrDefault() != null && item.Xml.ElementsAfterSelf().First().Name.Equals(DocX.w + "tbl"))
                    {
                        item.FollowingTable = new Table(base.Document, item.Xml.ElementsAfterSelf().First());
                    }
                    item.ParentContainer = GetParentFromXmlName(item.Xml.Ancestors().First().Name.LocalName);
                    if (item.IsListItem)
                    {
                        GetListItemType(item);
                    }
                }
                return paragraphs.AsReadOnly();
            }
        }

        public virtual ReadOnlyCollection<Paragraph> ParagraphsDeepSearch
        {
            get
            {
                List<Paragraph> paragraphs = GetParagraphs(deepSearch: true);
                foreach (Paragraph item in paragraphs)
                {
                    if (item.Xml.ElementsAfterSelf().FirstOrDefault() != null && item.Xml.ElementsAfterSelf().First().Name.Equals(DocX.w + "tbl"))
                    {
                        item.FollowingTable = new Table(base.Document, item.Xml.ElementsAfterSelf().First());
                    }
                    item.ParentContainer = GetParentFromXmlName(item.Xml.Ancestors().First().Name.LocalName);
                    if (item.IsListItem)
                    {
                        GetListItemType(item);
                    }
                }
                return paragraphs.AsReadOnly();
            }
        }

        public virtual List<Section> Sections
        {
            get
            {
                ReadOnlyCollection<Paragraph> paragraphs = Paragraphs;
                List<Paragraph> list = new List<Paragraph>();
                List<Section> list2 = new List<Section>();
                foreach (Paragraph item3 in paragraphs)
                {
                    XElement xElement = item3.Xml.Descendants().FirstOrDefault((XElement s) => s.Name.LocalName == "sectPr");
                    if (xElement == null)
                    {
                        list.Add(item3);
                    }
                    else
                    {
                        list.Add(item3);
                        Section item = new Section(base.Document, xElement)
                        {
                            SectionParagraphs = list
                        };
                        list2.Add(item);
                        list = new List<Paragraph>();
                    }
                }
                XElement xElement2 = base.Xml.Element(XName.Get("body", DocX.w.NamespaceName));
                if (xElement2 != null)
                {
                    XElement xml = xElement2.Element(XName.Get("sectPr", DocX.w.NamespaceName));
                    Section item2 = new Section(base.Document, xml)
                    {
                        SectionParagraphs = list
                    };
                    list2.Add(item2);
                }
                return list2;
            }
        }

        public virtual List<Table> Tables => (from t in base.Xml.Descendants(DocX.w + "tbl")
                                              select new Table(base.Document, t)).ToList();

        public virtual List<List> Lists
        {
            get
            {
                List<List> list = new List<List>();
                List list2 = new List(base.Document, base.Xml);
                foreach (Paragraph paragraph in Paragraphs)
                {
                    if (paragraph.IsListItem)
                    {
                        if (list2.CanAddListItem(paragraph))
                        {
                            list2.AddItem(paragraph);
                        }
                        else
                        {
                            list.Add(list2);
                            list2 = new List(base.Document, base.Xml);
                            list2.AddItem(paragraph);
                        }
                    }
                }
                list.Add(list2);
                return list;
            }
        }

        public virtual List<Hyperlink> Hyperlinks
        {
            get
            {
                List<Hyperlink> list = new List<Hyperlink>();
                foreach (Paragraph paragraph in Paragraphs)
                {
                    list.AddRange(paragraph.Hyperlinks);
                }
                return list;
            }
        }

        public virtual List<Picture> Pictures
        {
            get
            {
                List<Picture> list = new List<Picture>();
                foreach (Paragraph paragraph in Paragraphs)
                {
                    list.AddRange(paragraph.Pictures);
                }
                return list;
            }
        }

        public bool RemoveParagraphAt(int index)
        {
            int num = 0;
            foreach (XElement item in base.Xml.Descendants(DocX.w + "p"))
            {
                if (num == index)
                {
                    item.Remove();
                    return true;
                }
                num++;
            }
            return false;
        }

        public bool RemoveParagraph(Paragraph p)
        {
            foreach (XElement item in base.Xml.Descendants(DocX.w + "p"))
            {
                if (item.Equals(p.Xml))
                {
                    item.Remove();
                    return true;
                }
            }
            return false;
        }

        private void GetListItemType(Paragraph p)
        {
            XElement xElement = p.ParagraphNumberProperties.Descendants().FirstOrDefault((XElement el) => el.Name.LocalName == "ilvl");
            string value = xElement.Attribute(DocX.w + "val").Value;
            XElement xElement2 = p.ParagraphNumberProperties.Descendants().FirstOrDefault((XElement el) => el.Name.LocalName == "numId");
            string numIdValue = xElement2.Attribute(DocX.w + "val").Value;
            IEnumerable<XElement> source = from n in base.Document.numbering.Descendants()
                                           where n.Name.LocalName == "num"
                                           select n;
            XElement xElement3 = source.FirstOrDefault((XElement node) => node.Attribute(DocX.w + "numId").Value.Equals(numIdValue));
            if (xElement3 != null)
            {
                XElement xElement4 = xElement3.Descendants().First((XElement n) => n.Name.LocalName == "abstractNumId");
                string abstractNumNodeValue = xElement4.Attribute(DocX.w + "val").Value;
                IEnumerable<XElement> source2 = from n in base.Document.numbering.Descendants()
                                                where n.Name.LocalName == "abstractNum"
                                                select n;
                XElement xElement5 = source2.FirstOrDefault((XElement node) => node.Attribute(DocX.w + "abstractNumId").Value.Equals(abstractNumNodeValue));
                IEnumerable<XElement> enumerable = from n in xElement5.Descendants()
                                                   where n.Name.LocalName == "lvl"
                                                   select n;
                XElement xElement6 = null;
                foreach (XElement item in enumerable)
                {
                    if (item.Attribute(DocX.w + "ilvl").Value.Equals(value))
                    {
                        xElement6 = item;
                        break;
                    }
                }
                XElement xElement7 = xElement6.Descendants().First((XElement n) => n.Name.LocalName == "numFmt");
                p.ListItemType = GetListItemType(xElement7.Attribute(DocX.w + "val").Value);
            }
        }

        internal List<Content> GetContents(bool deepSearch = false)
        {
            List<Content> list = new List<Content>();
            foreach (XElement item in base.Xml.Descendants(XName.Get("sdt", DocX.w.NamespaceName)))
            {
                Content content = new Content(base.Document, item, 0);
                XElement e = item.Elements(XName.Get("sdtPr", DocX.w.NamespaceName)).First();
                content.Name = GetAttribute(e, "alias", "val");
                content.Tag = GetAttribute(e, "tag", "val");
                list.Add(content);
            }
            return list;
        }

        private string GetAttribute(XElement e, string localName, string attributeName)
        {
            string empty = string.Empty;
            try
            {
                return e.Elements(XName.Get(localName, DocX.w.NamespaceName)).Attributes(XName.Get(attributeName, DocX.w.NamespaceName)).FirstOrDefault()
                    .Value;
            }
            catch (Exception)
            {
                return "Missing";
            }
        }

        internal List<Paragraph> GetParagraphs(bool deepSearch = false)
        {
            int num = 0;
            List<Paragraph> list = new List<Paragraph>();
            foreach (XElement item2 in base.Xml.Descendants(XName.Get("p", DocX.w.NamespaceName)))
            {
                Paragraph item = new Paragraph(base.Document, item2, num);
                list.Add(item);
                num += HelperFunctions.GetText(item2).Length;
            }
            return list;
        }

        internal void GetParagraphsRecursive(XElement Xml, ref int index, ref List<Paragraph> paragraphs, bool deepSearch = false)
        {
            bool flag = true;
            if (Xml.Name.LocalName == "p")
            {
                paragraphs.Add(new Paragraph(base.Document, Xml, index));
                index += HelperFunctions.GetText(Xml).Length;
                if (!deepSearch)
                {
                    flag = false;
                }
            }
            if (flag && Xml.HasElements)
            {
                foreach (XElement item in Xml.Elements())
                {
                    GetParagraphsRecursive(item, ref index, ref paragraphs, deepSearch);
                }
            }
        }

        public virtual void SetDirection(Direction direction)
        {
            foreach (Paragraph paragraph in Paragraphs)
            {
                paragraph.Direction = direction;
            }
        }

        public virtual List<int> FindAll(string str)
        {
            return FindAll(str, RegexOptions.None);
        }

        public virtual List<int> FindAll(string str, RegexOptions options)
        {
            List<int> list = new List<int>();
            foreach (Paragraph paragraph in Paragraphs)
            {
                List<int> list2 = paragraph.FindAll(str, options);
                for (int i = 0; i < list2.Count(); i++)
                {
                    List<int> list3 = list2;
                    int index = i;
                    list3[index] += paragraph.startIndex;
                }
                list.AddRange(list2);
            }
            return list;
        }

        public virtual List<string> FindUniqueByPattern(string pattern, RegexOptions options)
        {
            List<string> list = new List<string>();
            foreach (Paragraph paragraph in Paragraphs)
            {
                List<string> collection = paragraph.FindAllByPattern(pattern, options);
                list.AddRange(collection);
            }
            Dictionary<string, int> dictionary = new Dictionary<string, int>();
            foreach (string item in list)
            {
                if (!dictionary.ContainsKey(item))
                {
                    dictionary.Add(item, 0);
                }
            }
            return dictionary.Keys.ToList();
        }

        public virtual void ReplaceText(string searchValue, string newValue, bool trackChanges = false, RegexOptions options = RegexOptions.None, Formatting newFormatting = null, Formatting matchFormatting = null, MatchFormattingOptions formattingOptions = MatchFormattingOptions.SubsetMatch, bool escapeRegEx = true, bool useRegExSubstitutions = false)
        {
            if (string.IsNullOrEmpty(searchValue))
            {
                throw new ArgumentException("oldValue cannot be null or empty", "searchValue");
            }
            if (newValue == null)
            {
                throw new ArgumentException("newValue cannot be null or empty", "newValue");
            }
            List<Header> list = new List<Header>
                {
                    base.Document.Headers.first,
                    base.Document.Headers.even,
                    base.Document.Headers.odd
                };
            foreach (Header item in list)
            {
                if (item != null)
                {
                    foreach (Paragraph paragraph in item.Paragraphs)
                    {
                        paragraph.ReplaceText(searchValue, newValue, trackChanges, options, newFormatting, matchFormatting, formattingOptions, escapeRegEx, useRegExSubstitutions);
                    }
                }
            }
            foreach (Paragraph paragraph2 in Paragraphs)
            {
                paragraph2.ReplaceText(searchValue, newValue, trackChanges, options, newFormatting, matchFormatting, formattingOptions, escapeRegEx, useRegExSubstitutions);
            }
            List<Footer> list2 = new List<Footer>
                {
                    base.Document.Footers.first,
                    base.Document.Footers.even,
                    base.Document.Footers.odd
                };
            foreach (Footer item2 in list2)
            {
                if (item2 != null)
                {
                    foreach (Paragraph paragraph3 in item2.Paragraphs)
                    {
                        paragraph3.ReplaceText(searchValue, newValue, trackChanges, options, newFormatting, matchFormatting, formattingOptions, escapeRegEx, useRegExSubstitutions);
                    }
                }
            }
        }

        public virtual void ReplaceText(string searchValue, Func<string, string> regexMatchHandler, bool trackChanges = false, RegexOptions options = RegexOptions.None, Formatting newFormatting = null, Formatting matchFormatting = null, MatchFormattingOptions formattingOptions = MatchFormattingOptions.SubsetMatch)
        {
            if (string.IsNullOrEmpty(searchValue))
            {
                throw new ArgumentException("oldValue cannot be null or empty", "searchValue");
            }
            if (regexMatchHandler == null)
            {
                throw new ArgumentException("regexMatchHandler cannot be null", "regexMatchHandler");
            }
            List<IParagraphContainer> list = new List<IParagraphContainer>
                {
                    base.Document.Headers.first,
                    base.Document.Headers.even,
                    base.Document.Headers.odd,
                    base.Document.Footers.first,
                    base.Document.Footers.even,
                    base.Document.Footers.odd
                };
            foreach (IParagraphContainer item in list)
            {
                if (item != null)
                {
                    foreach (Paragraph paragraph in item.Paragraphs)
                    {
                        paragraph.ReplaceText(searchValue, regexMatchHandler, trackChanges, options, newFormatting, matchFormatting, formattingOptions);
                    }
                }
            }
            foreach (Paragraph paragraph2 in Paragraphs)
            {
                paragraph2.ReplaceText(searchValue, regexMatchHandler, trackChanges, options, newFormatting, matchFormatting, formattingOptions);
            }
        }

        public int RemoveTextInGivenFormat(Formatting matchFormatting, MatchFormattingOptions fo = MatchFormattingOptions.SubsetMatch)
        {
            int num = 0;
            foreach (XElement item in base.Xml.Elements())
            {
                num += RemoveTextWithFormatRecursive(item, matchFormatting, fo);
            }
            return num;
        }

        internal int RemoveTextWithFormatRecursive(XElement element, Formatting matchFormatting, MatchFormattingOptions fo)
        {
            int num = 0;
            foreach (XElement item in element.Elements())
            {
                if ("rPr".Equals(item.Name.LocalName) && HelperFunctions.ContainsEveryChildOf(matchFormatting.Xml, item, fo))
                {
                    item.Parent.Remove();
                    num++;
                }
                num += RemoveTextWithFormatRecursive(item, matchFormatting, fo);
            }
            return num;
        }

        public virtual void InsertAtBookmark(string toInsert, string bookmarkName)
        {
            if (bookmarkName.IsNullOrWhiteSpace())
            {
                throw new ArgumentException("bookmark cannot be null or empty", "bookmarkName");
            }
            Headers headers = base.Document.Headers;
            List<Header> source = new List<Header>
                {
                    headers.first,
                    headers.even,
                    headers.odd
                };
            foreach (Header item in from x in source
                                    where x != null
                                    select x)
            {
                foreach (Paragraph paragraph in item.Paragraphs)
                {
                    paragraph.InsertAtBookmark(toInsert, bookmarkName);
                }
            }
            foreach (Paragraph paragraph2 in Paragraphs)
            {
                paragraph2.InsertAtBookmark(toInsert, bookmarkName);
            }
            Footers footers = base.Document.Footers;
            List<Footer> source2 = new List<Footer>
                {
                    footers.first,
                    footers.even,
                    footers.odd
                };
            foreach (Footer item2 in from x in source2
                                     where x != null
                                     select x)
            {
                foreach (Paragraph paragraph3 in item2.Paragraphs)
                {
                    paragraph3.InsertAtBookmark(toInsert, bookmarkName);
                }
            }
        }

        public string[] ValidateBookmarks(params string[] bookmarkNames)
        {
            List<Header> source = (from h in new Header[3]
            {
                    base.Document.Headers.first,
                    base.Document.Headers.even,
                    base.Document.Headers.odd
            }
                                   where h != null
                                   select h).ToList();
            List<Footer> source2 = (from f in new Footer[3]
            {
                    base.Document.Footers.first,
                    base.Document.Footers.even,
                    base.Document.Footers.odd
            }
                                    where f != null
                                    select f).ToList();
            List<string> list = new List<string>();
            foreach (string bookmarkName in bookmarkNames)
            {
                if (source.SelectMany((Header h) => h.Paragraphs).Any((Paragraph p) => p.ValidateBookmark(bookmarkName)))
                {
                    return new string[0];
                }
                if (source2.SelectMany((Footer h) => h.Paragraphs).Any((Paragraph p) => p.ValidateBookmark(bookmarkName)))
                {
                    return new string[0];
                }
                if (Paragraphs.Any((Paragraph p) => p.ValidateBookmark(bookmarkName)))
                {
                    return new string[0];
                }
                list.Add(bookmarkName);
            }
            return list.ToArray();
        }

        public virtual Paragraph InsertParagraph(int index, string text, bool trackChanges)
        {
            return InsertParagraph(index, text, trackChanges, null);
        }

        public virtual Paragraph InsertParagraph()
        {
            return InsertParagraph(string.Empty, trackChanges: false);
        }

        public virtual Paragraph InsertParagraph(int index, Paragraph p)
        {
            XElement xElement2 = p.Xml = new XElement(p.Xml);
            Paragraph firstParagraphEffectedByInsert = HelperFunctions.GetFirstParagraphEffectedByInsert(base.Document, index);
            if (firstParagraphEffectedByInsert == null)
            {
                base.Xml.Add(p.Xml);
            }
            else
            {
                XElement[] array = HelperFunctions.SplitParagraph(firstParagraphEffectedByInsert, index - firstParagraphEffectedByInsert.startIndex);
                firstParagraphEffectedByInsert.Xml.ReplaceWith(array[0], xElement2, array[1]);
            }
            GetParent(p);
            return p;
        }

        public virtual Paragraph InsertParagraph(Paragraph p)
        {
            if (p.styles.Count() > 0)
            {
                Uri uri = new Uri("/word/styles.xml", UriKind.Relative);
                XDocument xDocument;
                if (!base.Document.package.PartExists(uri))
                {
                    PackagePart val = base.Document.package.CreatePart(uri, "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml",  CompressionOption.Maximum);
                    using (TextWriter textWriter = new StreamWriter(new PackagePartStream(val.GetStream())))
                    {
                        xDocument = new XDocument(new XDeclaration("1.0", "UTF-8", "yes"), new XElement(XName.Get("styles", DocX.w.NamespaceName)));
                        xDocument.Save(textWriter);
                    }
                }
                PackagePart part = base.Document.package.GetPart(uri);
                using (TextReader textReader = new StreamReader(part.GetStream()))
                {
                    xDocument = XDocument.Load(textReader);
                    XElement xElement = xDocument.Element(XName.Get("styles", DocX.w.NamespaceName));
                    IEnumerable<string> source = from d in xElement.Descendants(XName.Get("style", DocX.w.NamespaceName))
                                                 let a = d.Attribute(XName.Get("styleId", DocX.w.NamespaceName))
                                                 where a != null
                                                 select a.Value;
                    foreach (XElement style in p.styles)
                    {
                        if (!source.Contains(style.Attribute(XName.Get("styleId", DocX.w.NamespaceName)).Value))
                        {
                            xElement.Add(style);
                        }
                    }
                }
                using (TextWriter textWriter2 = new StreamWriter(new PackagePartStream(part.GetStream())))
                {
                    xDocument.Save(textWriter2);
                }
            }
            XElement xElement2 = new XElement(p.Xml);
            base.Xml.Add(xElement2);
            int num = 0;
            if (base.Document.paragraphLookup.Keys.Count() > 0)
            {
                num = base.Document.paragraphLookup.Last().Key;
                num = ((base.Document.paragraphLookup.Last().Value.Text.Length != 0) ? (num + base.Document.paragraphLookup.Last().Value.Text.Length) : (num + 1));
            }
            Paragraph paragraph = new Paragraph(base.Document, xElement2, num);
            base.Document.paragraphLookup.Add(num, paragraph);
            GetParent(paragraph);
            return paragraph;
        }

        public virtual Paragraph InsertParagraph(int index, string text, bool trackChanges, Formatting formatting)
        {
            Paragraph paragraph = new Paragraph(base.Document, new XElement(DocX.w + "p"), index);
            paragraph.InsertText(0, text, trackChanges, formatting);
            Paragraph firstParagraphEffectedByInsert = HelperFunctions.GetFirstParagraphEffectedByInsert(base.Document, index);
            if (firstParagraphEffectedByInsert != null)
            {
                int num = index - firstParagraphEffectedByInsert.startIndex;
                if (num <= 0)
                {
                    firstParagraphEffectedByInsert.Xml.ReplaceWith(paragraph.Xml, firstParagraphEffectedByInsert.Xml);
                }
                else
                {
                    XElement[] array = HelperFunctions.SplitParagraph(firstParagraphEffectedByInsert, num);
                    firstParagraphEffectedByInsert.Xml.ReplaceWith(array[0], paragraph.Xml, array[1]);
                }
            }
            else
            {
                base.Xml.Add(paragraph);
            }
            GetParent(paragraph);
            return paragraph;
        }

        private ContainerType GetParentFromXmlName(string xmlName)
        {
            switch (xmlName)
            {
                case "body":
                    return ContainerType.Body;
                case "p":
                    return ContainerType.Paragraph;
                case "tbl":
                    return ContainerType.Table;
                case "sectPr":
                    return ContainerType.Section;
                case "tc":
                    return ContainerType.Cell;
                default:
                    return ContainerType.None;
            }
        }

        private void GetParent(Paragraph newParagraph)
        {
            Type type = GetType();
            switch (type.Name)
            {
                case "Body":
                    newParagraph.ParentContainer = ContainerType.Body;
                    break;
                case "Table":
                    newParagraph.ParentContainer = ContainerType.Table;
                    break;
                case "TOC":
                    newParagraph.ParentContainer = ContainerType.TOC;
                    break;
                case "Section":
                    newParagraph.ParentContainer = ContainerType.Section;
                    break;
                case "Cell":
                    newParagraph.ParentContainer = ContainerType.Cell;
                    break;
                case "Header":
                    newParagraph.ParentContainer = ContainerType.Header;
                    break;
                case "Footer":
                    newParagraph.ParentContainer = ContainerType.Footer;
                    break;
                case "Paragraph":
                    newParagraph.ParentContainer = ContainerType.Paragraph;
                    break;
            }
        }

        private ListItemType GetListItemType(string styleName)
        {
            if (styleName == "bullet")
            {
                return ListItemType.Bulleted;
            }
            return ListItemType.Numbered;
        }

        public virtual void InsertSection()
        {
            InsertSection(trackChanges: false);
        }

        public virtual void InsertSection(bool trackChanges)
        {
            XElement content = new XElement(XName.Get("p", DocX.w.NamespaceName), new XElement(XName.Get("pPr", DocX.w.NamespaceName), new XElement(XName.Get("sectPr", DocX.w.NamespaceName), new XElement(XName.Get("type", DocX.w.NamespaceName), new XAttribute(DocX.w + "val", "continuous")))));
            if (trackChanges)
            {
                content = HelperFunctions.CreateEdit(EditType.ins, DateTime.Now, content);
            }
            base.Xml.Add(content);
        }

        public virtual void InsertSectionPageBreak(bool trackChanges = false)
        {
            XElement content = new XElement(XName.Get("p", DocX.w.NamespaceName), new XElement(XName.Get("pPr", DocX.w.NamespaceName), new XElement(XName.Get("sectPr", DocX.w.NamespaceName))));
            if (trackChanges)
            {
                content = HelperFunctions.CreateEdit(EditType.ins, DateTime.Now, content);
            }
            base.Xml.Add(content);
        }

        public virtual Paragraph InsertParagraph(string text)
        {
            return InsertParagraph(text, false, new Formatting());
        }

        public virtual Paragraph InsertParagraph(string text, bool trackChanges)
        {
            return InsertParagraph(text, trackChanges, new Formatting());
        }

        public virtual Paragraph InsertParagraph(string text, bool trackChanges, Formatting formatting)
        {
            XElement xElement = new XElement(XName.Get("p", DocX.w.NamespaceName), new XElement(XName.Get("pPr", DocX.w.NamespaceName)), HelperFunctions.FormatInput(text, formatting.Xml));
            if (trackChanges)
            {
                xElement = HelperFunctions.CreateEdit(EditType.ins, DateTime.Now, xElement);
            }
            base.Xml.Add(xElement);
            Paragraph paragraph = new Paragraph(base.Document, xElement, 0);
            if (this is Cell)
            {
                Cell cell = this as Cell;
                paragraph.PackagePart = cell.mainPart;
            }
            else if (this is DocX)
            {
                paragraph.PackagePart = base.Document.mainPart;
            }
            else if (this is Footer)
            {
                Footer footer = this as Footer;
                paragraph.mainPart = footer.mainPart;
            }
            else if (this is Header)
            {
                Header header = this as Header;
                paragraph.mainPart = header.mainPart;
            }
            else
            {
                Console.WriteLine("No idea what we are {0}", this);
                paragraph.PackagePart = base.Document.mainPart;
            }
            GetParent(paragraph);
            return paragraph;
        }

        public virtual Paragraph InsertEquation(string equation)
        {
            Paragraph paragraph = InsertParagraph();
            paragraph.AppendEquation(equation);
            return paragraph;
        }

        public virtual Paragraph InsertBookmark(string bookmarkName)
        {
            Paragraph paragraph = InsertParagraph();
            paragraph.AppendBookmark(bookmarkName);
            return paragraph;
        }

        public virtual Table InsertTable(int rowCount, int columnCount)
        {
            XElement xElement = HelperFunctions.CreateTable(rowCount, columnCount);
            base.Xml.Add(xElement);
            return new Table(base.Document, xElement)
            {
                mainPart = mainPart
            };
        }

        public Table InsertTable(int index, int rowCount, int columnCount)
        {
            XElement xElement = HelperFunctions.CreateTable(rowCount, columnCount);
            Paragraph firstParagraphEffectedByInsert = HelperFunctions.GetFirstParagraphEffectedByInsert(base.Document, index);
            if (firstParagraphEffectedByInsert == null)
            {
                base.Xml.Elements().First().AddFirst(xElement);
            }
            else
            {
                XElement[] array = HelperFunctions.SplitParagraph(firstParagraphEffectedByInsert, index - firstParagraphEffectedByInsert.startIndex);
                firstParagraphEffectedByInsert.Xml.ReplaceWith(array[0], xElement, array[1]);
            }
            return new Table(base.Document, xElement)
            {
                mainPart = mainPart
            };
        }

        public Table InsertTable(Table t)
        {
            XElement xElement = new XElement(t.Xml);
            base.Xml.Add(xElement);
            return new Table(base.Document, xElement)
            {
                mainPart = mainPart,
                Design = t.Design
            };
        }

        public Table InsertTable(int index, Table t)
        {
            Paragraph firstParagraphEffectedByInsert = HelperFunctions.GetFirstParagraphEffectedByInsert(base.Document, index);
            XElement[] array = HelperFunctions.SplitParagraph(firstParagraphEffectedByInsert, index - firstParagraphEffectedByInsert.startIndex);
            XElement xElement = new XElement(t.Xml);
            firstParagraphEffectedByInsert.Xml.ReplaceWith(array[0], xElement, array[1]);
            return new Table(base.Document, xElement)
            {
                mainPart = mainPart,
                Design = t.Design
            };
        }

        internal Container(DocX document, XElement xml)
            : base(document, xml)
        {
        }

        public List InsertList(List list)
        {
            foreach (Paragraph item in list.Items)
            {
                base.Xml.Add(item.Xml);
            }
            return list;
        }

        public List InsertList(List list, double fontSize)
        {
            foreach (Paragraph item in list.Items)
            {
                item.FontSize(fontSize);
                base.Xml.Add(item.Xml);
            }
            return list;
        }

        public List InsertList(List list, Font fontFamily, double fontSize)
        {
            foreach (Paragraph item in list.Items)
            {
                item.Font(fontFamily);
                item.FontSize(fontSize);
                base.Xml.Add(item.Xml);
            }
            return list;
        }

        public List InsertList(int index, List list)
        {
            Paragraph firstParagraphEffectedByInsert = HelperFunctions.GetFirstParagraphEffectedByInsert(base.Document, index);
            XElement[] array = HelperFunctions.SplitParagraph(firstParagraphEffectedByInsert, index - firstParagraphEffectedByInsert.startIndex);
            List<XElement> list2 = new List<XElement>
                {
                    array[0]
                };
            list2.AddRange(from i in list.Items
                           select new XElement(i.Xml));
            list2.Add(array[1]);
            XElement xml = firstParagraphEffectedByInsert.Xml;
            object[] content = list2.ToArray();
            xml.ReplaceWith(content);
            return list;
        }
    }
}
