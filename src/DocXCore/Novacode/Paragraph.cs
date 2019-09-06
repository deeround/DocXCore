using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Security.Principal;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace Novacode
{
    public class Paragraph : InsertBeforeOrAfter
    {
        internal List<XElement> runs;

        private Alignment alignment;

        public ContainerType ParentContainer;

        public ListItemType ListItemType;

        internal int startIndex;

        internal int endIndex;

        private List<DocProperty> docProperties;

        internal List<XElement> styles = new List<XElement>();

        private Direction direction;

        private float indentationFirstLine;

        private float indentationHanging;

        private float indentationBefore;

        private float indentationAfter;

        private Table followingTable;

        private XElement ParagraphNumberPropertiesBacker
        {
            get;
            set;
        }

        public XElement ParagraphNumberProperties
        {
            get
            {
                object xElement = ParagraphNumberPropertiesBacker;
                if (xElement == null)
                {
                    XElement xElement2 = ParagraphNumberPropertiesBacker = GetParagraphNumberProperties();
                    xElement = xElement2;
                }
                return (XElement)xElement;
            }
        }

        private bool? IsListItemBacker
        {
            get;
            set;
        }

        public bool IsListItem
        {
            get
            {
                IsListItemBacker = (IsListItemBacker ?? (ParagraphNumberProperties != null));
                return IsListItemBacker.Value;
            }
        }

        private int? IndentLevelBacker
        {
            get;
            set;
        }

        public int? IndentLevel
        {
            get
            {
                if (!IsListItem)
                {
                    return null;
                }
                int? indentLevelBacker = IndentLevelBacker;
                int? result;
                if (!indentLevelBacker.HasValue)
                {
                    int? num2 = IndentLevelBacker = int.Parse(ParagraphNumberProperties.Descendants().First((XElement el) => el.Name.LocalName == "ilvl").GetAttribute(DocX.w + "val"));
                    result = num2;
                }
                else
                {
                    result = indentLevelBacker;
                }
                return result;
            }
        }

        public List<Picture> Pictures
        {
            get
            {
                List<Picture> list = (from p in base.Xml.Descendants()
                                      where p.Name.LocalName == "drawing"
                                      let id = (from e in p.Descendants()
                                                where e.Name.LocalName.Equals("blip")
                                                select e.Attribute(XName.Get("embed", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")).Value).SingleOrDefault()
                                      where id != null
                                      let img = new Image(base.Document, mainPart.GetRelationship(id))
                                      select new Picture(base.Document, p, img)).ToList();
                List<Picture> list2 = (from p in base.Xml.Descendants()
                                       where p.Name.LocalName == "pict"
                                       let id = (from e in p.Descendants()
                                                 where e.Name.LocalName.Equals("imagedata")
                                                 select e.Attribute(XName.Get("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")).Value).SingleOrDefault()
                                       where id != null
                                       let img = new Image(base.Document, mainPart.GetRelationship(id))
                                       select new Picture(base.Document, p, img)).ToList();
                foreach (Picture item in list2)
                {
                    list.Add(item);
                }
                return list;
            }
        }

        public List<Hyperlink> Hyperlinks
        {
            get
            {
                List<Hyperlink> list = new List<Hyperlink>();
                List<XElement> list2 = (from h in base.Xml.Descendants()
                                        where h.Name.LocalName == "hyperlink" || h.Name.LocalName == "instrText"
                                        select h).ToList();
                foreach (XElement item in list2)
                {
                    if (item.Name.LocalName == "hyperlink")
                    {
                        try
                        {
                            Hyperlink hyperlink = new Hyperlink(base.Document, mainPart, item);
                            hyperlink.mainPart = mainPart;
                            list.Add(hyperlink);
                        }
                        catch (Exception)
                        {
                        }
                    }
                    else
                    {
                        XElement xElement = item;
                        while (xElement.Name.LocalName != "r")
                        {
                            xElement = xElement.Parent;
                        }
                        List<XElement> list3 = new List<XElement>();
                        foreach (XElement item2 in xElement.ElementsAfterSelf(XName.Get("r", DocX.w.NamespaceName)))
                        {
                            list3.Add(item2);
                            XElement xElement2 = item2.Descendants(XName.Get("fldChar", DocX.w.NamespaceName)).SingleOrDefault();
                            if (xElement2 != null && (xElement2.Attribute(XName.Get("fldCharType", DocX.w.NamespaceName))?.Value.Equals("end", StringComparison.CurrentCultureIgnoreCase) ?? false))
                            {
                                try
                                {
                                    Hyperlink hyperlink2 = new Hyperlink(base.Document, item, list3);
                                    hyperlink2.mainPart = mainPart;
                                    list.Add(hyperlink2);
                                }
                                catch (Exception)
                                {
                                }
                                break;
                            }
                        }
                    }
                }
                return list;
            }
        }

        public string StyleName
        {
            get
            {
                XElement orCreate_pPr = GetOrCreate_pPr();
                XAttribute xAttribute = orCreate_pPr.Element(XName.Get("pStyle", DocX.w.NamespaceName))?.Attribute(XName.Get("val", DocX.w.NamespaceName));
                if (!string.IsNullOrEmpty(xAttribute?.Value))
                {
                    return xAttribute.Value;
                }
                return "Normal";
            }
            set
            {
                if (string.IsNullOrEmpty(value))
                {
                    value = "Normal";
                }
                XElement orCreate_pPr = GetOrCreate_pPr();
                XElement xElement = orCreate_pPr.Element(XName.Get("pStyle", DocX.w.NamespaceName));
                if (xElement == null)
                {
                    orCreate_pPr.Add(new XElement(XName.Get("pStyle", DocX.w.NamespaceName)));
                    xElement = orCreate_pPr.Element(XName.Get("pStyle", DocX.w.NamespaceName));
                }
                xElement.SetAttributeValue(XName.Get("val", DocX.w.NamespaceName), value);
            }
        }

        public List<DocProperty> DocumentProperties => docProperties;

        public Direction Direction
        {
            get
            {
                XElement orCreate_pPr = GetOrCreate_pPr();
                XElement xElement = orCreate_pPr.Element(XName.Get("bidi", DocX.w.NamespaceName));
                return (xElement != null) ? Direction.RightToLeft : Direction.LeftToRight;
            }
            set
            {
                direction = value;
                XElement orCreate_pPr = GetOrCreate_pPr();
                XElement xElement = orCreate_pPr.Element(XName.Get("bidi", DocX.w.NamespaceName));
                if (direction == Direction.RightToLeft)
                {
                    if (xElement == null)
                    {
                        orCreate_pPr.Add(new XElement(XName.Get("bidi", DocX.w.NamespaceName)));
                    }
                }
                else
                {
                    xElement?.Remove();
                }
            }
        }

        public bool IsKeepWithNext
        {
            get
            {
                XElement orCreate_pPr = GetOrCreate_pPr();
                XElement xElement = orCreate_pPr.Element(XName.Get("keepNext", DocX.w.NamespaceName));
                if (xElement == null)
                {
                    return false;
                }
                return true;
            }
        }

        public float IndentationFirstLine
        {
            get
            {
                GetOrCreate_pPr();
                XElement orCreate_pPr_ind = GetOrCreate_pPr_ind();
                XAttribute xAttribute = orCreate_pPr_ind.Attribute(XName.Get("firstLine", DocX.w.NamespaceName));
                if (xAttribute != null)
                {
                    return float.Parse(xAttribute.Value);
                }
                return 0f;
            }
            set
            {
                if (IndentationFirstLine != value)
                {
                    indentationFirstLine = value;
                    XElement orCreate_pPr = GetOrCreate_pPr();
                    XElement orCreate_pPr_ind = GetOrCreate_pPr_ind();
                    orCreate_pPr_ind.Attribute(XName.Get("hanging", DocX.w.NamespaceName))?.Remove();
                    string value2 = ((double)indentationFirstLine / 0.1 * 57.0).ToString();
                    XAttribute xAttribute = orCreate_pPr_ind.Attribute(XName.Get("firstLine", DocX.w.NamespaceName));
                    if (xAttribute != null)
                    {
                        xAttribute.Value = value2;
                    }
                    else
                    {
                        orCreate_pPr_ind.Add(new XAttribute(XName.Get("firstLine", DocX.w.NamespaceName), value2));
                    }
                }
            }
        }

        public float IndentationHanging
        {
            get
            {
                GetOrCreate_pPr();
                XElement orCreate_pPr_ind = GetOrCreate_pPr_ind();
                XAttribute xAttribute = orCreate_pPr_ind.Attribute(XName.Get("hanging", DocX.w.NamespaceName));
                if (xAttribute != null)
                {
                    return float.Parse(xAttribute.Value) / 570f;
                }
                return 0f;
            }
            set
            {
                if (IndentationHanging != value)
                {
                    indentationHanging = value;
                    XElement orCreate_pPr = GetOrCreate_pPr();
                    XElement orCreate_pPr_ind = GetOrCreate_pPr_ind();
                    orCreate_pPr_ind.Attribute(XName.Get("firstLine", DocX.w.NamespaceName))?.Remove();
                    string value2 = ((double)indentationHanging / 0.1 * 57.0).ToString();
                    XAttribute xAttribute = orCreate_pPr_ind.Attribute(XName.Get("hanging", DocX.w.NamespaceName));
                    if (xAttribute != null)
                    {
                        xAttribute.Value = value2;
                    }
                    else
                    {
                        orCreate_pPr_ind.Add(new XAttribute(XName.Get("hanging", DocX.w.NamespaceName), value2));
                    }
                }
            }
        }

        public float IndentationBefore
        {
            get
            {
                XElement orCreate_pPr = GetOrCreate_pPr();
                XElement orCreate_pPr_ind = GetOrCreate_pPr_ind();
                XAttribute xAttribute = orCreate_pPr_ind.Attribute(XName.Get("left", DocX.w.NamespaceName));
                if (xAttribute != null)
                {
                    return float.Parse(xAttribute.Value) / 570f;
                }
                return 0f;
            }
            set
            {
                if (IndentationBefore != value)
                {
                    indentationBefore = value;
                    XElement orCreate_pPr = GetOrCreate_pPr();
                    XElement orCreate_pPr_ind = GetOrCreate_pPr_ind();
                    string value2 = ((double)indentationBefore / 0.1 * 57.0).ToString(CultureInfo.CurrentCulture);
                    XAttribute xAttribute = orCreate_pPr_ind.Attribute(XName.Get("left", DocX.w.NamespaceName));
                    if (xAttribute != null)
                    {
                        xAttribute.Value = value2;
                    }
                    else
                    {
                        orCreate_pPr_ind.Add(new XAttribute(XName.Get("left", DocX.w.NamespaceName), value2));
                    }
                }
            }
        }

        public float IndentationAfter
        {
            get
            {
                GetOrCreate_pPr();
                XElement orCreate_pPr_ind = GetOrCreate_pPr_ind();
                XAttribute xAttribute = orCreate_pPr_ind.Attribute(XName.Get("right", DocX.w.NamespaceName));
                if (xAttribute != null)
                {
                    return float.Parse(xAttribute.Value);
                }
                return 0f;
            }
            set
            {
                if (IndentationAfter != value)
                {
                    indentationAfter = value;
                    XElement orCreate_pPr = GetOrCreate_pPr();
                    XElement orCreate_pPr_ind = GetOrCreate_pPr_ind();
                    string value2 = ((double)indentationAfter / 0.1 * 57.0).ToString();
                    XAttribute xAttribute = orCreate_pPr_ind.Attribute(XName.Get("right", DocX.w.NamespaceName));
                    if (xAttribute != null)
                    {
                        xAttribute.Value = value2;
                    }
                    else
                    {
                        orCreate_pPr_ind.Add(new XAttribute(XName.Get("right", DocX.w.NamespaceName), value2));
                    }
                }
            }
        }

        public Alignment Alignment
        {
            get
            {
                XElement orCreate_pPr = GetOrCreate_pPr();
                XElement xElement = orCreate_pPr.Element(XName.Get("jc", DocX.w.NamespaceName));
                if (xElement != null)
                {
                    XAttribute xAttribute = xElement.Attribute(XName.Get("val", DocX.w.NamespaceName));
                    switch (xAttribute.Value.ToLower())
                    {
                        case "left":
                            return Alignment.left;
                        case "right":
                            return Alignment.right;
                        case "center":
                            return Alignment.center;
                        case "both":
                            return Alignment.both;
                    }
                }
                return Alignment.left;
            }
            set
            {
                alignment = value;
                XElement orCreate_pPr = GetOrCreate_pPr();
                XElement xElement = orCreate_pPr.Element(XName.Get("jc", DocX.w.NamespaceName));
                if (alignment != 0)
                {
                    if (xElement == null)
                    {
                        orCreate_pPr.Add(new XElement(XName.Get("jc", DocX.w.NamespaceName), new XAttribute(XName.Get("val", DocX.w.NamespaceName), alignment.ToString())));
                    }
                    else
                    {
                        xElement.Attribute(XName.Get("val", DocX.w.NamespaceName)).Value = alignment.ToString();
                    }
                }
                else
                {
                    xElement?.Remove();
                }
            }
        }

        public string Text => HelperFunctions.GetText(base.Xml);

        public List<FormattedText> MagicText
        {
            get
            {
                try
                {
                    return HelperFunctions.GetFormattedText(base.Xml);
                }
                catch (Exception)
                {
                    return null;
                }
            }
        }

        public Table FollowingTable
        {
            get
            {
                return followingTable;
            }
            internal set
            {
                followingTable = value;
            }
        }

        public float LineSpacing
        {
            get
            {
                XElement orCreate_pPr = GetOrCreate_pPr();
                XElement xElement = orCreate_pPr.Element(XName.Get("spacing", DocX.w.NamespaceName));
                if (xElement != null)
                {
                    XAttribute xAttribute = xElement.Attribute(XName.Get("line", DocX.w.NamespaceName));
                    if (xAttribute != null && float.TryParse(xAttribute.Value, out float result))
                    {
                        return result / 20f;
                    }
                }
                return 22f;
            }
            set
            {
                Spacing((double)value);
            }
        }

        public float LineSpacingBefore
        {
            get
            {
                XElement orCreate_pPr = GetOrCreate_pPr();
                XElement xElement = orCreate_pPr.Element(XName.Get("spacing", DocX.w.NamespaceName));
                if (xElement != null)
                {
                    XAttribute xAttribute = xElement.Attribute(XName.Get("before", DocX.w.NamespaceName));
                    if (xAttribute != null && float.TryParse(xAttribute.Value, out float result))
                    {
                        return result / 20f;
                    }
                }
                return 0f;
            }
            set
            {
                SpacingBefore((double)value);
            }
        }

        public float LineSpacingAfter
        {
            get
            {
                XElement orCreate_pPr = GetOrCreate_pPr();
                XElement xElement = orCreate_pPr.Element(XName.Get("spacing", DocX.w.NamespaceName));
                if (xElement != null)
                {
                    XAttribute xAttribute = xElement.Attribute(XName.Get("after", DocX.w.NamespaceName));
                    if (xAttribute != null && float.TryParse(xAttribute.Value, out float result))
                    {
                        return result / 20f;
                    }
                }
                return 10f;
            }
            set
            {
                SpacingAfter((double)value);
            }
        }

        private XElement GetParagraphNumberProperties()
        {
            XElement xElement = base.Xml.Descendants().FirstOrDefault((XElement el) => el.Name.LocalName == "numPr");
            if (xElement != null)
            {
                XElement xElement2 = xElement.Descendants().First((XElement numId) => numId.Name.LocalName == "numId");
                if (xElement2.Attribute(DocX.w + "val")?.Value.Equals("0") ?? false)
                {
                    return null;
                }
            }
            return xElement;
        }

        internal Paragraph(DocX document, XElement xml, int startIndex, ContainerType parent = ContainerType.None)
            : base(document, xml)
        {
            ParentContainer = parent;
            this.startIndex = startIndex;
            endIndex = startIndex + GetElementTextLength(xml);
            RebuildDocProperties();
            runs = base.Xml.Elements(XName.Get("r", DocX.w.NamespaceName)).ToList();
        }

        public override Table InsertTableBeforeSelf(Table t)
        {
            t = base.InsertTableBeforeSelf(t);
            t.mainPart = mainPart;
            return t;
        }

        public Paragraph KeepWithNext(bool keepWithNext = true)
        {
            XElement orCreate_pPr = GetOrCreate_pPr();
            XElement xElement = orCreate_pPr.Element(XName.Get("keepNext", DocX.w.NamespaceName));
            if (xElement == null && keepWithNext)
            {
                orCreate_pPr.Add(new XElement(XName.Get("keepNext", DocX.w.NamespaceName)));
            }
            if (!keepWithNext && xElement != null)
            {
                xElement.Remove();
            }
            return this;
        }

        public Paragraph KeepLinesTogether(bool keepTogether = true)
        {
            XElement orCreate_pPr = GetOrCreate_pPr();
            XElement xElement = orCreate_pPr.Element(XName.Get("keepLines", DocX.w.NamespaceName));
            if (xElement == null && keepTogether)
            {
                orCreate_pPr.Add(new XElement(XName.Get("keepLines", DocX.w.NamespaceName)));
            }
            if (!keepTogether)
            {
                xElement?.Remove();
            }
            return this;
        }

        internal XElement GetOrCreate_pPr()
        {
            XElement xElement = base.Xml.Element(XName.Get("pPr", DocX.w.NamespaceName));
            if (xElement == null)
            {
                base.Xml.AddFirst(new XElement(XName.Get("pPr", DocX.w.NamespaceName)));
                xElement = base.Xml.Element(XName.Get("pPr", DocX.w.NamespaceName));
            }
            return xElement;
        }

        internal XElement GetOrCreate_pPr_ind()
        {
            XElement orCreate_pPr = GetOrCreate_pPr();
            XElement xElement = orCreate_pPr.Element(XName.Get("ind", DocX.w.NamespaceName));
            if (xElement == null)
            {
                orCreate_pPr.Add(new XElement(XName.Get("ind", DocX.w.NamespaceName)));
                xElement = orCreate_pPr.Element(XName.Get("ind", DocX.w.NamespaceName));
            }
            return xElement;
        }

        public override Table InsertTableBeforeSelf(int rowCount, int columnCount)
        {
            return base.InsertTableBeforeSelf(rowCount, columnCount);
        }

        public override Table InsertTableAfterSelf(Table t)
        {
            t = base.InsertTableAfterSelf(t);
            t.mainPart = mainPart;
            return t;
        }

        public override Table InsertTableAfterSelf(int rowCount, int columnCount)
        {
            return base.InsertTableAfterSelf(rowCount, columnCount);
        }

        public override Paragraph InsertParagraphBeforeSelf(Paragraph p)
        {
            Paragraph paragraph = base.InsertParagraphBeforeSelf(p);
            paragraph.PackagePart = mainPart;
            return paragraph;
        }

        public override Paragraph InsertParagraphBeforeSelf(string text)
        {
            Paragraph paragraph = base.InsertParagraphBeforeSelf(text);
            paragraph.PackagePart = mainPart;
            return paragraph;
        }

        public override Paragraph InsertParagraphBeforeSelf(string text, bool trackChanges)
        {
            Paragraph paragraph = base.InsertParagraphBeforeSelf(text, trackChanges);
            paragraph.PackagePart = mainPart;
            return paragraph;
        }

        public override Paragraph InsertParagraphBeforeSelf(string text, bool trackChanges, Formatting formatting)
        {
            Paragraph paragraph = base.InsertParagraphBeforeSelf(text, trackChanges, formatting);
            paragraph.PackagePart = mainPart;
            return paragraph;
        }

        public override void InsertPageBreakBeforeSelf()
        {
            base.InsertPageBreakBeforeSelf();
        }

        public override void InsertPageBreakAfterSelf()
        {
            base.InsertPageBreakAfterSelf();
        }

        [Obsolete("Instead use: InsertHyperlink(Hyperlink h, int index)")]
        public Paragraph InsertHyperlink(int index, Hyperlink h)
        {
            return InsertHyperlink(h, index);
        }

        public Paragraph InsertHyperlink(Hyperlink h, int index = 0)
        {
            string arg = mainPart.Uri.OriginalString.Replace("/word/", "");
            Uri uri = new Uri($"/word/_rels/{arg}.rels", UriKind.Relative);
            if (!base.Document.package.PartExists(uri))
            {
                HelperFunctions.CreateRelsPackagePart(base.Document, uri);
            }
            string orGenerateRel = GetOrGenerateRel(h);
            if (index == 0)
            {
                base.Xml.AddFirst(h.Xml);
                XElement xElement = (XElement)base.Xml.FirstNode;
            }
            else
            {
                Run firstRunEffectedByEdit = GetFirstRunEffectedByEdit(index);
                XElement xElement;
                if (firstRunEffectedByEdit == null)
                {
                    base.Xml.Add(h.Xml);
                    xElement = (XElement)base.Xml.LastNode;
                }
                else
                {
                    XElement[] array = Run.SplitRun(firstRunEffectedByEdit, index);
                    firstRunEffectedByEdit.Xml.ReplaceWith(array[0], h.Xml, array[1]);
                    firstRunEffectedByEdit = GetFirstRunEffectedByEdit(index);
                    xElement = (XElement)firstRunEffectedByEdit.Xml.NextNode;
                }
                xElement.SetAttributeValue(DocX.r + "id", orGenerateRel);
            }
            return this;
        }

        public void RemoveHyperlink(int index)
        {
            if (index < 0)
            {
                throw new ArgumentOutOfRangeException();
            }
            int count = 0;
            bool found = false;
            RemoveHyperlinkRecursive(base.Xml, index, ref count, ref found);
            if (!found)
            {
                throw new ArgumentOutOfRangeException();
            }
        }

        internal void RemoveHyperlinkRecursive(XElement xml, int index, ref int count, ref bool found)
        {
            if (xml.Name.LocalName.Equals("hyperlink", StringComparison.CurrentCultureIgnoreCase))
            {
                if (count == index)
                {
                    found = true;
                    xml.Remove();
                }
                else
                {
                    count++;
                }
            }
            if (xml.HasElements)
            {
                foreach (XElement item in xml.Elements())
                {
                    if (!found)
                    {
                        RemoveHyperlinkRecursive(item, index, ref count, ref found);
                    }
                }
            }
        }

        public override Paragraph InsertParagraphAfterSelf(Paragraph p)
        {
            Paragraph paragraph = base.InsertParagraphAfterSelf(p);
            paragraph.PackagePart = mainPart;
            return paragraph;
        }

        public override Paragraph InsertParagraphAfterSelf(string text, bool trackChanges, Formatting formatting)
        {
            Paragraph paragraph = base.InsertParagraphAfterSelf(text, trackChanges, formatting);
            paragraph.PackagePart = mainPart;
            return paragraph;
        }

        public override Paragraph InsertParagraphAfterSelf(string text, bool trackChanges)
        {
            Paragraph paragraph = base.InsertParagraphAfterSelf(text, trackChanges);
            paragraph.PackagePart = mainPart;
            return paragraph;
        }

        public override Paragraph InsertParagraphAfterSelf(string text)
        {
            Paragraph paragraph = base.InsertParagraphAfterSelf(text);
            paragraph.PackagePart = mainPart;
            return paragraph;
        }

        private void RebuildDocProperties()
        {
            docProperties = (from xml in base.Xml.Descendants(XName.Get("fldSimple", DocX.w.NamespaceName))
                             select new DocProperty(base.Document, xml)).ToList();
        }

        public void Remove(bool trackChanges)
        {
            if (trackChanges)
            {
                DateTime edit_time = DateTime.Now.ToUniversalTime();
                List<XElement> list = base.Xml.Elements().ToList();
                List<XElement> list2 = new List<XElement>();
                for (int i = 0; i < list.Count(); i++)
                {
                    XElement xElement = list[i];
                    if (xElement.Name.LocalName != "del")
                    {
                        list2.Add(xElement);
                        xElement.Remove();
                    }
                    else if (list2.Count() > 0)
                    {
                        xElement.AddBeforeSelf(CreateEdit(EditType.del, edit_time, list2.Elements()));
                        list2.Clear();
                    }
                }
                if (list2.Count() > 0)
                {
                    base.Xml.Add(CreateEdit(EditType.del, edit_time, list2));
                }
            }
            else if (base.Xml.Parent.Name.LocalName == "tc" && base.Xml.Parent.Elements(XName.Get("p", DocX.w.NamespaceName)).Count() == 1)
            {
                base.Xml.Value = string.Empty;
            }
            else
            {
                base.Xml.Remove();
                base.Xml = null;
            }
        }

        internal static Picture CreatePicture(DocX document, string id, string name, string descr)
        {
            PackagePart part = document.package.GetPart(document.mainPart.GetRelationship(id).TargetUri);
            long nextFreeDocPrId = document.GetNextFreeDocPrId();
            int num = 0;
            int num2 = 0;
            using (PackagePartStream packagePartStream = new PackagePartStream(part.GetStream()))
            {
                using (System.Drawing.Image img = System.Drawing.Image.FromStream(packagePartStream, useEmbeddedColorManagement: false, validateImageData: false))
                {
                    num = img.Width * 9526;
                    num2 = img.Height * 9526;
                }
            }
            XElement i = XElement.Parse(string.Format("<w:r xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">\r\n                    <w:drawing xmlns = \"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">\r\n                        <wp:inline distT=\"0\" distB=\"0\" distL=\"0\" distR=\"0\" xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\">\r\n                            <wp:extent cx=\"{0}\" cy=\"{1}\" />\r\n                            <wp:effectExtent l=\"0\" t=\"0\" r=\"0\" b=\"0\" />\r\n                            <wp:docPr id=\"{5}\" name=\"{3}\" descr=\"{4}\" />\r\n                            <wp:cNvGraphicFramePr>\r\n                                <a:graphicFrameLocks xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" noChangeAspect=\"1\" />\r\n                            </wp:cNvGraphicFramePr>\r\n                            <a:graphic xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">\r\n                                <a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">\r\n                                    <pic:pic xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">\r\n                                        <pic:nvPicPr>\r\n                                        <pic:cNvPr id=\"0\" name=\"{3}\" />\r\n                                            <pic:cNvPicPr />\r\n                                        </pic:nvPicPr>\r\n                                        <pic:blipFill>\r\n                                            <a:blip r:embed=\"{2}\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"/>\r\n                                            <a:stretch>\r\n                                                <a:fillRect />\r\n                                            </a:stretch>\r\n                                        </pic:blipFill>\r\n                                        <pic:spPr>\r\n                                            <a:xfrm>\r\n                                                <a:off x=\"0\" y=\"0\" />\r\n                                                <a:ext cx=\"{0}\" cy=\"{1}\" />\r\n                                            </a:xfrm>\r\n                                            <a:prstGeom prst=\"rect\">\r\n                                                <a:avLst />\r\n                                            </a:prstGeom>\r\n                                        </pic:spPr>\r\n                                    </pic:pic>\r\n                                </a:graphicData>\r\n                            </a:graphic>\r\n                        </wp:inline>\r\n                    </w:drawing></w:r>\r\n                    ", num, num2, id, name, descr, nextFreeDocPrId.ToString()));
            return new Picture(document, i, new Image(document, document.mainPart.GetRelationship(id)));
        }

        internal static XElement CreateEdit(EditType t, DateTime edit_time, object content)
        {
            if (t == EditType.del)
            {
                foreach (XElement item in (IEnumerable<XElement>)content)
                {
                    if (item is XElement)
                    {
                        XElement xElement = item as XElement;
                        IEnumerable<XElement> source = xElement.DescendantsAndSelf(XName.Get("t", DocX.w.NamespaceName));
                        for (int i = 0; i < source.Count(); i++)
                        {
                            XElement xElement2 = source.ElementAt(i);
                            xElement2.ReplaceWith(new XElement(DocX.w + "delText", xElement2.Attributes(), xElement2.Value));
                        }
                    }
                }
            }
            return new XElement(DocX.w + t.ToString(), new XAttribute(DocX.w + "id", 0), new XAttribute(DocX.w + "author", "elane"), new XAttribute(DocX.w + "date", edit_time), content);
        }

        internal Run GetFirstRunEffectedByEdit(int index, EditType type = EditType.ins)
        {
            int length = HelperFunctions.GetText(base.Xml).Length;
            if (index < 0 || (type == EditType.ins && index > length) || (type == EditType.del && index >= length))
            {
                throw new ArgumentOutOfRangeException();
            }
            int count = 0;
            Run theOne = null;
            GetFirstRunEffectedByEditRecursive(base.Xml, index, ref count, ref theOne, type);
            return theOne;
        }

        internal void GetFirstRunEffectedByEditRecursive(XElement Xml, int index, ref int count, ref Run theOne, EditType type)
        {
            count += HelperFunctions.GetSize(Xml);
            if (count > 0 && ((type == EditType.del && count > index) || (type == EditType.ins && count >= index)))
            {
                foreach (XElement item in Xml.ElementsBeforeSelf())
                {
                    count -= HelperFunctions.GetSize(item);
                }
                count -= HelperFunctions.GetSize(Xml);
                while (Xml.Name.LocalName != "r" && Xml.Name.LocalName != "pPr")
                {
                    Xml = Xml.Parent;
                }
                theOne = new Run(base.Document, Xml, count);
            }
            else if (Xml.HasElements)
            {
                foreach (XElement item2 in Xml.Elements())
                {
                    if (theOne == null)
                    {
                        GetFirstRunEffectedByEditRecursive(item2, index, ref count, ref theOne, type);
                    }
                }
            }
        }

        internal static int GetElementTextLength(XElement run)
        {
            int num = 0;
            if (run == null)
            {
                return num;
            }
            foreach (XElement item in run.Descendants())
            {
                switch (item.Name.LocalName)
                {
                    case "tab":
                        if (!(item.Parent.Name.LocalName != "tabs"))
                        {
                            break;
                        }
                        goto case "br";
                    case "br":
                        num++;
                        break;
                    case "t":
                    case "delText":
                        num += item.Value.Length;
                        break;
                }
            }
            return num;
        }

        internal XElement[] SplitEdit(XElement edit, int index, EditType type)
        {
            Run firstRunEffectedByEdit = GetFirstRunEffectedByEdit(index, type);
            XElement[] array = Run.SplitRun(firstRunEffectedByEdit, index, type);
            XElement xElement = new XElement(edit.Name, edit.Attributes(), firstRunEffectedByEdit.Xml.ElementsBeforeSelf(), array[0]);
            if (GetElementTextLength(xElement) == 0)
            {
                xElement = null;
            }
            XElement xElement2 = new XElement(edit.Name, edit.Attributes(), array[1], firstRunEffectedByEdit.Xml.ElementsAfterSelf());
            if (GetElementTextLength(xElement2) == 0)
            {
                xElement2 = null;
            }
            return new XElement[2]
            {
                xElement,
                xElement2
            };
        }

        public void InsertText(string value, bool trackChanges = false, Formatting formatting = null)
        {
            if (formatting == null)
            {
                formatting = new Formatting();
            }
            List<XElement> content = HelperFunctions.FormatInput(value, formatting.Xml);
            base.Xml.Add(content);
            HelperFunctions.RenumberIDs(base.Document);
        }

        public void InsertText(int index, string value, bool trackChanges = false, Formatting formatting = null)
        {
            DateTime now = DateTime.Now;
            DateTime dateTime = new DateTime(now.Year, now.Month, now.Day, now.Hour, now.Minute, 0, DateTimeKind.Utc);
            Run firstRunEffectedByEdit = GetFirstRunEffectedByEdit(index);
            object obj;
            XElement parent;
            if (firstRunEffectedByEdit == null)
            {
                object content = (formatting == null) ? HelperFunctions.FormatInput(value, null) : HelperFunctions.FormatInput(value, formatting.Xml);
                if (trackChanges)
                {
                    content = CreateEdit(EditType.ins, dateTime, content);
                }
                base.Xml.Add(content);
            }
            else
            {
                XElement xElement = firstRunEffectedByEdit.Xml.Element(XName.Get("rPr", DocX.w.NamespaceName));
                if (formatting != null)
                {
                    Formatting formatting2 = null;
                    Formatting formatting3 = null;
                    if (xElement != null)
                    {
                        formatting3 = Formatting.Parse(xElement);
                    }
                    if (formatting3 != null)
                    {
                        formatting2 = formatting3.Clone();
                        if (formatting.Bold.HasValue)
                        {
                            formatting2.Bold = formatting.Bold;
                        }
                        if (formatting.CapsStyle.HasValue)
                        {
                            formatting2.CapsStyle = formatting.CapsStyle;
                        }
                        if (formatting.FontColor.HasValue)
                        {
                            formatting2.FontColor = formatting.FontColor;
                        }
                        formatting2.FontFamily = formatting.FontFamily;
                        if (formatting.Hidden.HasValue)
                        {
                            formatting2.Hidden = formatting.Hidden;
                        }
                        if (formatting.Highlight.HasValue)
                        {
                            formatting2.Highlight = formatting.Highlight;
                        }
                        if (formatting.Italic.HasValue)
                        {
                            formatting2.Italic = formatting.Italic;
                        }
                        if (formatting.Kerning.HasValue)
                        {
                            formatting2.Kerning = formatting.Kerning;
                        }
                        formatting2.Language = formatting.Language;
                        if (formatting.Misc.HasValue)
                        {
                            formatting2.Misc = formatting.Misc;
                        }
                        if (formatting.PercentageScale.HasValue)
                        {
                            formatting2.PercentageScale = formatting.PercentageScale;
                        }
                        if (formatting.Position.HasValue)
                        {
                            formatting2.Position = formatting.Position;
                        }
                        if (formatting.Script.HasValue)
                        {
                            formatting2.Script = formatting.Script;
                        }
                        if (formatting.Size.HasValue)
                        {
                            formatting2.Size = formatting.Size;
                        }
                        if (formatting.Spacing.HasValue)
                        {
                            formatting2.Spacing = formatting.Spacing;
                        }
                        if (formatting.StrikeThrough.HasValue)
                        {
                            formatting2.StrikeThrough = formatting.StrikeThrough;
                        }
                        if (formatting.UnderlineColor.HasValue)
                        {
                            formatting2.UnderlineColor = formatting.UnderlineColor;
                        }
                        if (formatting.UnderlineStyle.HasValue)
                        {
                            formatting2.UnderlineStyle = formatting.UnderlineStyle;
                        }
                    }
                    else
                    {
                        formatting2 = formatting;
                    }
                    obj = HelperFunctions.FormatInput(value, formatting2.Xml);
                }
                else
                {
                    obj = HelperFunctions.FormatInput(value, xElement);
                }
                parent = firstRunEffectedByEdit.Xml.Parent;
                string localName = parent.Name.LocalName;
                if (!(localName == "ins"))
                {
                    if (localName == "del")
                    {
                        goto IL_042a;
                    }
                }
                else
                {
                    DateTime dateTime2 = DateTime.Parse(parent.Attribute(XName.Get("date", DocX.w.NamespaceName)).Value);
                    if (!trackChanges || dateTime2.CompareTo(dateTime) != 0)
                    {
                        goto IL_042a;
                    }
                }
                object obj2 = obj;
                if (trackChanges && !parent.Name.LocalName.Equals("ins"))
                {
                    obj2 = CreateEdit(EditType.ins, dateTime, obj);
                }
                else
                {
                    XElement[] array = Run.SplitRun(firstRunEffectedByEdit, index);
                    firstRunEffectedByEdit.Xml.ReplaceWith(array[0], obj2, array[1]);
                }
            }
            goto IL_04db;
        IL_042a:
            object obj3 = obj;
            if (trackChanges)
            {
                obj3 = CreateEdit(EditType.ins, dateTime, obj);
            }
            XElement[] array2 = SplitEdit(parent, index, EditType.ins);
            parent.ReplaceWith(array2[0], obj3, array2[1]);
            goto IL_04db;
        IL_04db:
            HelperFunctions.RenumberIDs(base.Document);
        }

        public Paragraph CurentCulture()
        {
            ApplyTextFormattingProperty(XName.Get("lang", DocX.w.NamespaceName), string.Empty, new XAttribute(XName.Get("val", DocX.w.NamespaceName), CultureInfo.CurrentCulture.Name));
            return this;
        }

        public Paragraph Culture(CultureInfo culture)
        {
            ApplyTextFormattingProperty(XName.Get("lang", DocX.w.NamespaceName), string.Empty, new XAttribute(XName.Get("val", DocX.w.NamespaceName), culture.Name));
            return this;
        }

        public Paragraph Append(string text)
        {
            List<XElement> list = HelperFunctions.FormatInput(text, null);
            base.Xml.Add(list);
            runs = base.Xml.Elements(XName.Get("r", DocX.w.NamespaceName)).Reverse().Take(list.Count())
                .ToList();
            return this;
        }

        public void InsertHorizontalLine(string lineType = "single", int size = 6, int space = 1, string color = "auto")
        {
            XElement orCreate_pPr = GetOrCreate_pPr();
            XElement xElement = orCreate_pPr.Element(XName.Get("pBdr", DocX.w.NamespaceName));
            if (xElement == null)
            {
                orCreate_pPr.Add(new XElement(XName.Get("pBdr", DocX.w.NamespaceName)));
                xElement = orCreate_pPr.Element(XName.Get("pBdr", DocX.w.NamespaceName));
                xElement.Add(new XElement(XName.Get("bottom", DocX.w.NamespaceName)));
                XElement xElement2 = xElement.Element(XName.Get("bottom", DocX.w.NamespaceName));
                xElement2.SetAttributeValue(XName.Get("val", DocX.w.NamespaceName), lineType);
                xElement2.SetAttributeValue(XName.Get("sz", DocX.w.NamespaceName), size.ToString());
                xElement2.SetAttributeValue(XName.Get("space", DocX.w.NamespaceName), space.ToString());
                xElement2.SetAttributeValue(XName.Get("color", DocX.w.NamespaceName), color);
            }
        }

        public Paragraph AppendHyperlink(Hyperlink h)
        {
            string str = mainPart.Uri.OriginalString.Replace("/word/", "");
            Uri uri = new Uri("/word/_rels/" + str + ".rels", UriKind.Relative);
            if (!base.Document.package.PartExists(uri))
            {
                HelperFunctions.CreateRelsPackagePart(base.Document, uri);
            }
            string orGenerateRel = GetOrGenerateRel(h);
            base.Xml.Add(h.Xml);
            base.Xml.Elements().Last().SetAttributeValue(DocX.r + "id", orGenerateRel);
            runs = base.Xml.Elements().Last().Elements(XName.Get("r", DocX.w.NamespaceName))
                .ToList();
            return this;
        }

        public Paragraph AppendPicture(Picture p)
        {
            string str = mainPart.Uri.OriginalString.Replace("/word/", "");
            Uri uri = new Uri("/word/_rels/" + str + ".rels", UriKind.Relative);
            if (!base.Document.package.PartExists(uri))
            {
                HelperFunctions.CreateRelsPackagePart(base.Document, uri);
            }
            string orGenerateRel = GetOrGenerateRel(p);
            base.Xml.Add(p.Xml);
            XAttribute xAttribute = (from e in base.Xml.Elements().Last().Descendants()
                                     where e.Name.LocalName.Equals("blip")
                                     select e.Attribute(XName.Get("embed", "http://schemas.openxmlformats.org/officeDocument/2006/relationships"))).Single();
            xAttribute.SetValue(orGenerateRel);
            runs = base.Xml.Elements(XName.Get("r", DocX.w.NamespaceName)).Reverse().Take(p.Xml.Elements(XName.Get("r", DocX.w.NamespaceName)).Count())
                .ToList();
            return this;
        }

        public Paragraph AppendEquation(string equation)
        {
            XElement content = new XElement(XName.Get("oMathPara", DocX.m.NamespaceName), new XElement(XName.Get("oMath", DocX.m.NamespaceName), new XElement(XName.Get("r", DocX.w.NamespaceName), new Formatting
            {
                FontFamily = new Font("Cambria Math")
            }.Xml, new XElement(XName.Get("t", DocX.m.NamespaceName), equation))));
            base.Xml.Add(content);
            runs = base.Xml.Elements(XName.Get("oMathPara", DocX.m.NamespaceName)).ToList();
            return this;
        }

        public bool ValidateBookmark(string bookmarkName)
        {
            return GetBookmarks().Any((Bookmark b) => b.Name.Equals(bookmarkName));
        }

        public Paragraph AppendBookmark(string bookmarkName)
        {
            XElement content = new XElement(XName.Get("bookmarkStart", DocX.w.NamespaceName), new XAttribute(XName.Get("id", DocX.w.NamespaceName), 0), new XAttribute(XName.Get("name", DocX.w.NamespaceName), bookmarkName));
            base.Xml.Add(content);
            XElement content2 = new XElement(XName.Get("bookmarkEnd", DocX.w.NamespaceName), new XAttribute(XName.Get("id", DocX.w.NamespaceName), 0), new XAttribute(XName.Get("name", DocX.w.NamespaceName), bookmarkName));
            base.Xml.Add(content2);
            return this;
        }

        public IEnumerable<Bookmark> GetBookmarks()
        {
            return from x in base.Xml.Descendants(XName.Get("bookmarkStart", DocX.w.NamespaceName))
                   select x.Attribute(XName.Get("name", DocX.w.NamespaceName)) into x
                   select new Bookmark
                   {
                       Name = x.Value,
                       Paragraph = this
                   };
        }

        public void InsertAtBookmark(string toInsert, string bookmarkName)
        {
            XElement xElement = (from x in base.Xml.Descendants(XName.Get("bookmarkStart", DocX.w.NamespaceName))
                                 where x.Attribute(XName.Get("name", DocX.w.NamespaceName)).Value == bookmarkName
                                 select x).SingleOrDefault();
            if (xElement != null)
            {
                List<XElement> content = HelperFunctions.FormatInput(toInsert, null);
                xElement.AddBeforeSelf(content);
                runs = base.Xml.Elements(XName.Get("r", DocX.w.NamespaceName)).ToList();
                HelperFunctions.RenumberIDs(base.Document);
            }
        }

        public void ReplaceAtBookmark(string toInsert, string bookmarkName)
        {
            XElement xElement = (from x in base.Xml.Descendants(XName.Get("bookmarkStart", DocX.w.NamespaceName))
                                 where x.Attribute(XName.Get("name", DocX.w.NamespaceName)).Value == bookmarkName
                                 select x).SingleOrDefault();
            if (xElement != null)
            {
                XNode nextNode = xElement.NextNode;
                XElement xElement2 = nextNode as XElement;
                while (xElement2 == null || xElement2.Name.NamespaceName != DocX.w.NamespaceName || (xElement2.Name.LocalName != "r" && xElement2.Name.LocalName != "bookmarkEnd"))
                {
                    nextNode = nextNode.NextNode;
                    xElement2 = (nextNode as XElement);
                }
                if (xElement2.Name.LocalName == "bookmarkEnd")
                {
                    ReplaceAtBookmark_Add(toInsert, xElement);
                }
                else
                {
                    XElement xElement3 = xElement2.Elements(XName.Get("t", DocX.w.NamespaceName)).FirstOrDefault();
                    if (xElement3 == null)
                    {
                        ReplaceAtBookmark_Add(toInsert, xElement);
                    }
                    else
                    {
                        xElement3.Value = toInsert;
                    }
                }
            }
        }

        private void ReplaceAtBookmark_Add(string toInsert, XElement bookmark)
        {
            List<XElement> content = HelperFunctions.FormatInput(toInsert, null);
            bookmark.AddAfterSelf(content);
            runs = base.Xml.Elements(XName.Get("r", DocX.w.NamespaceName)).ToList();
            HelperFunctions.RenumberIDs(base.Document);
        }

        internal string GetOrGenerateRel(Picture p)
        {
            string originalString = p.img.pr.TargetUri.OriginalString;
            string text = null;
            foreach (PackageRelationship item in mainPart.GetRelationshipsByType("http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"))
            {
                if (string.Equals(item.TargetUri.OriginalString, originalString, StringComparison.Ordinal))
                {
                    text = item.Id;
                    break;
                }
            }
            if (text == null)
            {
                PackageRelationship val = mainPart.CreateRelationship(p.img.pr.TargetUri, 0, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image");
                text = val.Id;
            }
            return text;
        }

        internal string GetOrGenerateRel(Hyperlink h)
        {
            string image_uri_string = h.Uri.OriginalString;
            string text = (from r in (IEnumerable<PackageRelationship>)mainPart.GetRelationshipsByType("http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink")
                           where r.TargetUri.OriginalString == image_uri_string
                           select r.Id).SingleOrDefault();
            if (text == null)
            {
                PackageRelationship val = mainPart.CreateRelationship(h.Uri, TargetMode.External, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink");
                text = val.Id;
            }
            return text;
        }

        public Paragraph InsertPicture(Picture p, int index = 0)
        {
            string str = mainPart.Uri.OriginalString.Replace("/word/", "");
            Uri uri = new Uri("/word/_rels/" + str + ".rels", UriKind.Relative);
            if (!base.Document.package.PartExists(uri))
            {
                HelperFunctions.CreateRelsPackagePart(base.Document, uri);
            }
            string orGenerateRel = GetOrGenerateRel(p);
            XElement xElement;
            if (index == 0)
            {
                base.Xml.AddFirst(p.Xml);
                xElement = (XElement)base.Xml.FirstNode;
            }
            else
            {
                Run firstRunEffectedByEdit = GetFirstRunEffectedByEdit(index);
                if (firstRunEffectedByEdit == null)
                {
                    base.Xml.Add(p.Xml);
                    xElement = (XElement)base.Xml.LastNode;
                }
                else
                {
                    XElement[] array = Run.SplitRun(firstRunEffectedByEdit, index);
                    firstRunEffectedByEdit.Xml.ReplaceWith(array[0], p.Xml, array[1]);
                    firstRunEffectedByEdit = GetFirstRunEffectedByEdit(index);
                    xElement = (XElement)firstRunEffectedByEdit.Xml.NextNode;
                }
            }
            XAttribute xAttribute = (from e in xElement.Descendants()
                                     where e.Name.LocalName.Equals("blip")
                                     select e.Attribute(XName.Get("embed", "http://schemas.openxmlformats.org/officeDocument/2006/relationships"))).Single();
            xAttribute.SetValue(orGenerateRel);
            return this;
        }

        public Paragraph AppendLine(string text)
        {
            return Append("\n" + text);
        }

        public Paragraph AppendLine()
        {
            return Append("\n");
        }

        internal void ApplyTextFormattingProperty(XName textFormatPropName, string value, object content)
        {
            XElement xElement = null;
            if (runs.Count == 0)
            {
                XElement xElement2 = base.Xml.Element(XName.Get("pPr", DocX.w.NamespaceName));
                if (xElement2 == null)
                {
                    base.Xml.AddFirst(new XElement(XName.Get("pPr", DocX.w.NamespaceName)));
                    xElement2 = base.Xml.Element(XName.Get("pPr", DocX.w.NamespaceName));
                }
                xElement = xElement2.Element(XName.Get("rPr", DocX.w.NamespaceName));
                if (xElement == null)
                {
                    xElement2.AddFirst(new XElement(XName.Get("rPr", DocX.w.NamespaceName)));
                    xElement = xElement2.Element(XName.Get("rPr", DocX.w.NamespaceName));
                }
                xElement.SetElementValue(textFormatPropName, value);
                XElement xElement3 = xElement.Elements(textFormatPropName).Last();
                if (content is XAttribute)
                {
                    if (xElement3.Attribute(((XAttribute)content).Name) == null)
                    {
                        xElement3.Add(content);
                    }
                    else
                    {
                        xElement3.Attribute(((XAttribute)content).Name).Value = ((XAttribute)content).Value;
                    }
                }
            }
            else
            {
                bool flag = false;
                IEnumerable enumerable = content as IEnumerable;
                if (enumerable != null)
                {
                    foreach (object item in enumerable)
                    {
                        flag = (item is XAttribute);
                    }
                }
                foreach (XElement run in runs)
                {
                    xElement = run.Element(XName.Get("rPr", DocX.w.NamespaceName));
                    if (xElement == null)
                    {
                        run.AddFirst(new XElement(XName.Get("rPr", DocX.w.NamespaceName)));
                        xElement = run.Element(XName.Get("rPr", DocX.w.NamespaceName));
                    }
                    xElement.SetElementValue(textFormatPropName, value);
                    XElement xElement4 = xElement.Elements(textFormatPropName).Last();
                    if (flag)
                    {
                        foreach (object item2 in enumerable)
                        {
                            if (xElement4.Attribute(((XAttribute)item2).Name) == null)
                            {
                                xElement4.Add(item2);
                            }
                            else
                            {
                                xElement4.Attribute(((XAttribute)item2).Name).Value = ((XAttribute)item2).Value;
                            }
                        }
                    }
                    if (content is XAttribute)
                    {
                        if (xElement4.Attribute(((XAttribute)content).Name) == null)
                        {
                            xElement4.Add(content);
                        }
                        else
                        {
                            xElement4.Attribute(((XAttribute)content).Name).Value = ((XAttribute)content).Value;
                        }
                    }
                }
            }
        }

        public Paragraph Bold()
        {
            ApplyTextFormattingProperty(XName.Get("b", DocX.w.NamespaceName), string.Empty, null);
            return this;
        }

        public Paragraph Italic()
        {
            ApplyTextFormattingProperty(XName.Get("i", DocX.w.NamespaceName), string.Empty, null);
            return this;
        }

        public Paragraph Color(Color c)
        {
            ApplyTextFormattingProperty(XName.Get("color", DocX.w.NamespaceName), string.Empty, new XAttribute(XName.Get("val", DocX.w.NamespaceName), c.ToHex()));
            return this;
        }

        public Paragraph UnderlineStyle(UnderlineStyle underlineStyle)
        {
            string value;
            switch (underlineStyle)
            {
                case Novacode.UnderlineStyle.none:
                    value = string.Empty;
                    break;
                case Novacode.UnderlineStyle.singleLine:
                    value = "single";
                    break;
                case Novacode.UnderlineStyle.doubleLine:
                    value = "double";
                    break;
                default:
                    value = underlineStyle.ToString();
                    break;
            }
            ApplyTextFormattingProperty(XName.Get("u", DocX.w.NamespaceName), string.Empty, new XAttribute(XName.Get("val", DocX.w.NamespaceName), value));
            return this;
        }

        public Paragraph FontSize(double fontSize)
        {
            double num = fontSize * 2.0;
            if (num - (double)(int)num != 0.0)
            {
                throw new ArgumentException("Size", "Value must be either a whole or half number, examples: 32, 32.5");
            }
            if (!(fontSize > 0.0) || !(fontSize < 1639.0))
            {
                throw new ArgumentException("Size", "Value must be in the range 0 - 1638");
            }
            ApplyTextFormattingProperty(XName.Get("sz", DocX.w.NamespaceName), string.Empty, new XAttribute(XName.Get("val", DocX.w.NamespaceName), fontSize * 2.0));
            ApplyTextFormattingProperty(XName.Get("szCs", DocX.w.NamespaceName), string.Empty, new XAttribute(XName.Get("val", DocX.w.NamespaceName), fontSize * 2.0));
            return this;
        }

        public Paragraph Font(string fontName)
        {
            return Font(new Font(fontName));
        }

        public Paragraph Font(Font fontFamily)
        {
            ApplyTextFormattingProperty(XName.Get("rFonts", DocX.w.NamespaceName), string.Empty, new XAttribute[4]
            {
                new XAttribute(XName.Get("ascii", DocX.w.NamespaceName), fontFamily.Name),
                new XAttribute(XName.Get("hAnsi", DocX.w.NamespaceName), fontFamily.Name),
                new XAttribute(XName.Get("cs", DocX.w.NamespaceName), fontFamily.Name),
                new XAttribute(XName.Get("eastAsia", DocX.w.NamespaceName), fontFamily.Name)
            });
            return this;
        }

        public Paragraph CapsStyle(CapsStyle capsStyle)
        {
            if (capsStyle != 0)
            {
                ApplyTextFormattingProperty(XName.Get(capsStyle.ToString(), DocX.w.NamespaceName), string.Empty, null);
            }
            return this;
        }

        public Paragraph Script(Script script)
        {
            Script script2 = script;
            if (script2 != Novacode.Script.none)
            {
                ApplyTextFormattingProperty(XName.Get("vertAlign", DocX.w.NamespaceName), string.Empty, new XAttribute(XName.Get("val", DocX.w.NamespaceName), script.ToString()));
            }
            return this;
        }

        public Paragraph Highlight(Highlight highlight)
        {
            Highlight highlight2 = highlight;
            if (highlight2 != Novacode.Highlight.none)
            {
                ApplyTextFormattingProperty(XName.Get("highlight", DocX.w.NamespaceName), string.Empty, new XAttribute(XName.Get("val", DocX.w.NamespaceName), highlight.ToString()));
            }
            return this;
        }

        public Paragraph Misc(Misc misc)
        {
            switch (misc)
            {
                case Novacode.Misc.outlineShadow:
                    ApplyTextFormattingProperty(XName.Get("outline", DocX.w.NamespaceName), string.Empty, null);
                    ApplyTextFormattingProperty(XName.Get("shadow", DocX.w.NamespaceName), string.Empty, null);
                    break;
                case Novacode.Misc.engrave:
                    ApplyTextFormattingProperty(XName.Get("imprint", DocX.w.NamespaceName), string.Empty, null);
                    break;
                default:
                    ApplyTextFormattingProperty(XName.Get(misc.ToString(), DocX.w.NamespaceName), string.Empty, null);
                    break;
                case Novacode.Misc.none:
                    break;
            }
            return this;
        }

        public Paragraph StrikeThrough(StrikeThrough strikeThrough)
        {
            string localName;
            switch (strikeThrough)
            {
                case Novacode.StrikeThrough.strike:
                    localName = "strike";
                    break;
                case Novacode.StrikeThrough.doubleStrike:
                    localName = "dstrike";
                    break;
                default:
                    return this;
            }
            ApplyTextFormattingProperty(XName.Get(localName, DocX.w.NamespaceName), string.Empty, null);
            return this;
        }

        public Paragraph UnderlineColor(Color underlineColor)
        {
            foreach (XElement run in runs)
            {
                XElement xElement = run.Element(XName.Get("rPr", DocX.w.NamespaceName));
                if (xElement == null)
                {
                    run.AddFirst(new XElement(XName.Get("rPr", DocX.w.NamespaceName)));
                    xElement = run.Element(XName.Get("rPr", DocX.w.NamespaceName));
                }
                XElement xElement2 = xElement.Element(XName.Get("u", DocX.w.NamespaceName));
                if (xElement2 == null)
                {
                    xElement.SetElementValue(XName.Get("u", DocX.w.NamespaceName), string.Empty);
                    xElement2 = xElement.Element(XName.Get("u", DocX.w.NamespaceName));
                    xElement2.SetAttributeValue(XName.Get("val", DocX.w.NamespaceName), "single");
                }
                xElement2.SetAttributeValue(XName.Get("color", DocX.w.NamespaceName), underlineColor.ToHex());
            }
            return this;
        }

        public Paragraph Hide()
        {
            ApplyTextFormattingProperty(XName.Get("vanish", DocX.w.NamespaceName), string.Empty, null);
            return this;
        }

        public void SetLineSpacing(LineSpacingType spacingType, float spacingFloat)
        {
            spacingFloat *= 240f;
            int num = (int)spacingFloat;
            XElement orCreate_pPr = GetOrCreate_pPr();
            XElement xElement = orCreate_pPr.Element(XName.Get("spacing", DocX.w.NamespaceName));
            if (xElement == null)
            {
                orCreate_pPr.Add(new XElement(XName.Get("spacing", DocX.w.NamespaceName)));
                xElement = orCreate_pPr.Element(XName.Get("spacing", DocX.w.NamespaceName));
            }
            string localName = "";
            switch (spacingType)
            {
                case LineSpacingType.Line:
                    localName = "line";
                    break;
                case LineSpacingType.Before:
                    localName = "before";
                    break;
                case LineSpacingType.After:
                    localName = "after";
                    break;
            }
            xElement.SetAttributeValue(XName.Get(localName, DocX.w.NamespaceName), num);
        }

        public void SetLineSpacing(LineSpacingTypeAuto spacingType)
        {
            int num = 100;
            XElement orCreate_pPr = GetOrCreate_pPr();
            XElement xElement = orCreate_pPr.Element(XName.Get("spacing", DocX.w.NamespaceName));
            if (spacingType.Equals(LineSpacingTypeAuto.None))
            {
                xElement?.Remove();
            }
            else
            {
                if (xElement == null)
                {
                    orCreate_pPr.Add(new XElement(XName.Get("spacing", DocX.w.NamespaceName)));
                    xElement = orCreate_pPr.Element(XName.Get("spacing", DocX.w.NamespaceName));
                }
                string localName = "";
                string localName2 = "";
                switch (spacingType)
                {
                    case LineSpacingTypeAuto.AutoBefore:
                        localName = "before";
                        localName2 = "beforeAutospacing";
                        break;
                    case LineSpacingTypeAuto.AutoAfter:
                        localName = "after";
                        localName2 = "afterAutospacing";
                        break;
                    case LineSpacingTypeAuto.Auto:
                        localName = "before";
                        localName2 = "beforeAutospacing";
                        xElement.SetAttributeValue(XName.Get("after", DocX.w.NamespaceName), num);
                        xElement.SetAttributeValue(XName.Get("afterAutospacing", DocX.w.NamespaceName), 1);
                        break;
                }
                xElement.SetAttributeValue(XName.Get(localName2, DocX.w.NamespaceName), 1);
                xElement.SetAttributeValue(XName.Get(localName, DocX.w.NamespaceName), num);
            }
        }

        public Paragraph Spacing(double spacing)
        {
            spacing *= 20.0;
            if (spacing - (double)(int)spacing != 0.0)
            {
                throw new ArgumentException("Spacing", "Value must be either a whole or acurate to one decimal, examples: 32, 32.1, 32.2, 32.9");
            }
            if (!(spacing > -1585.0) || !(spacing < 1585.0))
            {
                throw new ArgumentException("Spacing", "Value must be in the range: -1584 - 1584");
            }
            ApplyTextFormattingProperty(XName.Get("spacing", DocX.w.NamespaceName), string.Empty, new XAttribute(XName.Get("val", DocX.w.NamespaceName), spacing));
            return this;
        }

        public Paragraph SpacingBefore(double spacingBefore)
        {
            spacingBefore *= 20.0;
            XElement orCreate_pPr = GetOrCreate_pPr();
            XElement xElement = orCreate_pPr.Element(XName.Get("spacing", DocX.w.NamespaceName));
            if (spacingBefore > 0.0)
            {
                if (xElement == null)
                {
                    xElement = new XElement(XName.Get("spacing", DocX.w.NamespaceName));
                    orCreate_pPr.Add(xElement);
                }
                XAttribute xAttribute = xElement.Attribute(XName.Get("before", DocX.w.NamespaceName));
                if (xAttribute == null)
                {
                    xElement.SetAttributeValue(XName.Get("before", DocX.w.NamespaceName), spacingBefore);
                }
                else
                {
                    xAttribute.SetValue(spacingBefore);
                }
            }
            if (Math.Abs(spacingBefore) < 0.10000000149011612 && xElement != null)
            {
                XAttribute xAttribute2 = xElement.Attribute(XName.Get("before", DocX.w.NamespaceName));
                xAttribute2.Remove();
                if (!xElement.HasAttributes)
                {
                    xElement.Remove();
                }
            }
            return this;
        }

        public Paragraph SpacingAfter(double spacingAfter)
        {
            spacingAfter *= 20.0;
            XElement orCreate_pPr = GetOrCreate_pPr();
            XElement xElement = orCreate_pPr.Element(XName.Get("spacing", DocX.w.NamespaceName));
            if (spacingAfter > 0.0)
            {
                if (xElement == null)
                {
                    xElement = new XElement(XName.Get("spacing", DocX.w.NamespaceName));
                    orCreate_pPr.Add(xElement);
                }
                XAttribute xAttribute = xElement.Attribute(XName.Get("after", DocX.w.NamespaceName));
                if (xAttribute == null)
                {
                    xElement.SetAttributeValue(XName.Get("after", DocX.w.NamespaceName), spacingAfter);
                }
                else
                {
                    xAttribute.SetValue(spacingAfter);
                }
            }
            if (Math.Abs(spacingAfter) < 0.10000000149011612 && xElement != null)
            {
                XAttribute xAttribute2 = xElement.Attribute(XName.Get("after", DocX.w.NamespaceName));
                xAttribute2.Remove();
                if (!xElement.HasAttributes)
                {
                    xElement.Remove();
                }
            }
            return this;
        }

        public Paragraph Kerning(int kerning)
        {
            if (!new int?[16]
            {
                8,
                9,
                10,
                11,
                12,
                14,
                16,
                18,
                20,
                22,
                24,
                26,
                28,
                36,
                48,
                72
            }.Contains(kerning))
            {
                throw new ArgumentOutOfRangeException("Kerning", "Value must be one of the following: 8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48 or 72");
            }
            ApplyTextFormattingProperty(XName.Get("kern", DocX.w.NamespaceName), string.Empty, new XAttribute(XName.Get("val", DocX.w.NamespaceName), kerning * 2));
            return this;
        }

        public Paragraph Position(double position)
        {
            if (!(position > -1585.0) || !(position < 1585.0))
            {
                throw new ArgumentOutOfRangeException("Position", "Value must be in the range -1585 - 1585");
            }
            ApplyTextFormattingProperty(XName.Get("position", DocX.w.NamespaceName), string.Empty, new XAttribute(XName.Get("val", DocX.w.NamespaceName), position * 2.0));
            return this;
        }

        public Paragraph PercentageScale(int percentageScale)
        {
            if (!new int?[8]
            {
                200,
                150,
                100,
                90,
                80,
                66,
                50,
                33
            }.Contains(percentageScale))
            {
                throw new ArgumentOutOfRangeException("PercentageScale", "Value must be one of the following: 200, 150, 100, 90, 80, 66, 50 or 33");
            }
            ApplyTextFormattingProperty(XName.Get("w", DocX.w.NamespaceName), string.Empty, new XAttribute(XName.Get("val", DocX.w.NamespaceName), percentageScale));
            return this;
        }

        public Paragraph AppendDocProperty(CustomProperty cp, bool trackChanges = false, Formatting f = null)
        {
            InsertDocProperty(cp, trackChanges, f);
            return this;
        }

        public DocProperty InsertDocProperty(CustomProperty cp, bool trackChanges = false, Formatting f = null)
        {
            XElement xElement = null;
            if (f != null)
            {
                xElement = f.Xml;
            }
            XElement xElement2 = new XElement(XName.Get("fldSimple", DocX.w.NamespaceName), new XAttribute(XName.Get("instr", DocX.w.NamespaceName), $"DOCPROPERTY {cp.Name} \\* MERGEFORMAT"), new XElement(XName.Get("r", DocX.w.NamespaceName), new XElement(XName.Get("t", DocX.w.NamespaceName), xElement, cp.Value)));
            XElement xml = xElement2;
            if (trackChanges)
            {
                DateTime now = DateTime.Now;
                DateTime edit_time = new DateTime(now.Year, now.Month, now.Day, now.Hour, now.Minute, 0, DateTimeKind.Utc);
                xElement2 = CreateEdit(EditType.ins, edit_time, xElement2);
            }
            base.Xml.Add(xElement2);
            return new DocProperty(base.Document, xml);
        }

        public void RemoveText(int index, int count, bool trackChanges = false, bool removeEmptyParagraph = true)
        {
            DateTime now = DateTime.Now;
            DateTime edit_time = new DateTime(now.Year, now.Month, now.Day, now.Hour, now.Minute, 0, DateTimeKind.Utc);
            int num = 0;
            do
            {
                Run firstRunEffectedByEdit = GetFirstRunEffectedByEdit(index, EditType.del);
                XElement parent = firstRunEffectedByEdit.Xml.Parent;
                string localName = parent.Name.LocalName;
                if (localName == "ins")
                {
                    goto IL_007e;
                }
                if (localName == "del")
                {
                    if (!trackChanges)
                    {
                        goto IL_007e;
                    }
                    num += GetElementTextLength(parent);
                }
                else
                {
                    XElement[] array = Run.SplitRun(firstRunEffectedByEdit, index, EditType.del);
                    int index2 = Math.Min(index + (count - num), firstRunEffectedByEdit.EndIndex);
                    XElement[] array2 = Run.SplitRun(firstRunEffectedByEdit, index2, EditType.del);
                    object obj = CreateEdit(EditType.del, edit_time, new List<XElement>
                    {
                        Run.SplitRun(new Run(base.Document, array[1], firstRunEffectedByEdit.StartIndex + GetElementTextLength(array[0])), index2, EditType.del)[0]
                    });
                    num += GetElementTextLength(obj as XElement);
                    if (!trackChanges)
                    {
                        obj = null;
                    }
                    firstRunEffectedByEdit.Xml.ReplaceWith(array[0], obj, array2[1]);
                }
                goto IL_0210;
            IL_007e:
                XElement[] array3 = SplitEdit(parent, index, EditType.del);
                int num2 = Math.Min(count - num, firstRunEffectedByEdit.Xml.ElementsAfterSelf().Sum((XElement e) => GetElementTextLength(e)));
                XElement[] array4 = SplitEdit(parent, index + num2, EditType.del);
                XElement xElement = SplitEdit(array3[1], index + num2, EditType.del)[0];
                object obj2 = CreateEdit(EditType.del, edit_time, xElement.Elements());
                num += GetElementTextLength(obj2 as XElement);
                if (!trackChanges)
                {
                    obj2 = null;
                }
                parent.ReplaceWith(array3[0], obj2, array4[1]);
                num += GetElementTextLength(obj2 as XElement);
                goto IL_0210;
            IL_0210:
                if (removeEmptyParagraph && GetElementTextLength(parent) == 0 && parent.Parent != null && parent.Parent.Name.LocalName != "tc" && parent.Descendants(XName.Get("drawing", DocX.w.NamespaceName)).Count() == 0)
                {
                    parent.Remove();
                }
            }
            while (num < count);
            HelperFunctions.RenumberIDs(base.Document);
        }

        public void RemoveText(int index, bool trackChanges = false)
        {
            RemoveText(index, Text.Length - index, trackChanges);
        }

        public void ReplaceText(string oldValue, string newValue, bool trackChanges = false, RegexOptions options = RegexOptions.None, Formatting newFormatting = null, Formatting matchFormatting = null, MatchFormattingOptions fo = MatchFormattingOptions.SubsetMatch, bool escapeRegEx = true, bool useRegExSubstitutions = false, bool removeEmptyParagraph = true)
        {
            string text = Text;
            MatchCollection source = Regex.Matches(text, escapeRegEx ? Regex.Escape(oldValue) : oldValue, options);
            foreach (Match item in source.Cast<Match>().Reverse())
            {
                bool flag = true;
                if (matchFormatting != null)
                {
                    int num = 0;
                    do
                    {
                        Run firstRunEffectedByEdit = GetFirstRunEffectedByEdit(item.Index + num);
                        XElement xElement = firstRunEffectedByEdit.Xml.Element(XName.Get("rPr", DocX.w.NamespaceName));
                        if (xElement == null)
                        {
                            xElement = new Formatting().Xml;
                        }
                        if (!HelperFunctions.ContainsEveryChildOf(matchFormatting.Xml, xElement, fo))
                        {
                            flag = false;
                            break;
                        }
                        num += firstRunEffectedByEdit.Value.Length;
                    }
                    while (num < item.Length);
                }
                if (flag)
                {
                    string text2 = newValue;
                    if (useRegExSubstitutions && !string.IsNullOrEmpty(text2))
                    {
                        text2 = text2.Replace("$&", item.Value);
                        if (item.Groups.Count > 0)
                        {
                            int groupnum = 0;
                            for (int i = 0; i < item.Groups.Count; i++)
                            {
                                Group group = item.Groups[i];
                                if (group != null && !(group.Value == ""))
                                {
                                    text2 = text2.Replace("$" + i.ToString(), group.Value);
                                    groupnum = i;
                                }
                            }
                            text2 = text2.Replace("$+", item.Groups[groupnum].Value);
                        }
                        if (item.Index > 0)
                        {
                            text2 = text2.Replace("$`", text.Substring(0, item.Index));
                        }
                        if (item.Index + item.Length < text.Length)
                        {
                            text2 = text2.Replace("$'", text.Substring(item.Index + item.Length));
                        }
                        text2 = text2.Replace("$_", text);
                        text2 = text2.Replace("$$", "$");
                    }
                    if (!string.IsNullOrEmpty(text2))
                    {
                        InsertText(item.Index + item.Length, text2, trackChanges, newFormatting);
                    }
                    if (item.Length > 0)
                    {
                        RemoveText(item.Index, item.Length, trackChanges, removeEmptyParagraph);
                    }
                }
            }
        }

        public void ReplaceText(string findPattern, Func<string, string> regexMatchHandler, bool trackChanges = false, RegexOptions options = RegexOptions.None, Formatting newFormatting = null, Formatting matchFormatting = null, MatchFormattingOptions fo = MatchFormattingOptions.SubsetMatch, bool removeEmptyParagraph = true)
        {
            MatchCollection source = Regex.Matches(Text, findPattern, options);
            foreach (Match item in source.Cast<Match>().Reverse())
            {
                bool flag = true;
                if (matchFormatting != null)
                {
                    int num = 0;
                    do
                    {
                        Run firstRunEffectedByEdit = GetFirstRunEffectedByEdit(item.Index + num);
                        XElement xElement = firstRunEffectedByEdit.Xml.Element(XName.Get("rPr", DocX.w.NamespaceName));
                        if (xElement == null)
                        {
                            xElement = new Formatting().Xml;
                        }
                        if (!HelperFunctions.ContainsEveryChildOf(matchFormatting.Xml, xElement, fo))
                        {
                            flag = false;
                            break;
                        }
                        num += firstRunEffectedByEdit.Value.Length;
                    }
                    while (num < item.Length);
                }
                if (flag)
                {
                    string value = regexMatchHandler(item.Groups[1].Value);
                    InsertText(item.Index + item.Value.Length, value, trackChanges, newFormatting);
                    RemoveText(item.Index, item.Value.Length, trackChanges, removeEmptyParagraph);
                }
            }
        }

        public List<int> FindAll(string str)
        {
            return FindAll(str, RegexOptions.None);
        }

        public List<int> FindAll(string str, RegexOptions options)
        {
            MatchCollection source = Regex.Matches(Text, Regex.Escape(str), options);
            return (from Match m in source
                    select m.Index).ToList();
        }

        public List<string> FindAllByPattern(string str, RegexOptions options)
        {
            MatchCollection source = Regex.Matches(Text, str, options);
            return (from Match m in source
                    select m.Value).ToList();
        }

        public void InsertPageNumber(PageNumberFormat pnf, int index = 0)
        {
            XElement xElement = new XElement(XName.Get("fldSimple", DocX.w.NamespaceName));
            if (pnf == PageNumberFormat.normal)
            {
                xElement.Add(new XAttribute(XName.Get("instr", DocX.w.NamespaceName), " PAGE   \\* MERGEFORMAT "));
            }
            else
            {
                xElement.Add(new XAttribute(XName.Get("instr", DocX.w.NamespaceName), " PAGE  \\* ROMAN  \\* MERGEFORMAT "));
            }
            XElement content = XElement.Parse("<w:r w:rsidR='001D0226' xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">\r\n                   <w:rPr>\r\n                       <w:noProof /> \r\n                   </w:rPr>\r\n                   <w:t>1</w:t> \r\n               </w:r>");
            xElement.Add(content);
            if (index == 0)
            {
                base.Xml.AddFirst(xElement);
            }
            else
            {
                Run firstRunEffectedByEdit = GetFirstRunEffectedByEdit(index);
                XElement[] array = SplitEdit(firstRunEffectedByEdit.Xml, index, EditType.ins);
                firstRunEffectedByEdit.Xml.ReplaceWith(array[0], xElement, array[1]);
            }
        }

        public void AppendPageNumber(PageNumberFormat pnf)
        {
            XElement xElement = new XElement(XName.Get("fldSimple", DocX.w.NamespaceName));
            if (pnf == PageNumberFormat.normal)
            {
                xElement.Add(new XAttribute(XName.Get("instr", DocX.w.NamespaceName), " PAGE   \\* MERGEFORMAT "));
            }
            else
            {
                xElement.Add(new XAttribute(XName.Get("instr", DocX.w.NamespaceName), " PAGE  \\* ROMAN  \\* MERGEFORMAT "));
            }
            XElement content = XElement.Parse("<w:r w:rsidR='001D0226' xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">\r\n                   <w:rPr>\r\n                       <w:noProof /> \r\n                   </w:rPr>\r\n                   <w:t>1</w:t> \r\n               </w:r>");
            xElement.Add(content);
            base.Xml.Add(xElement);
        }

        public void InsertPageCount(PageNumberFormat pnf, int index = 0)
        {
            XElement xElement = new XElement(XName.Get("fldSimple", DocX.w.NamespaceName));
            if (pnf == PageNumberFormat.normal)
            {
                xElement.Add(new XAttribute(XName.Get("instr", DocX.w.NamespaceName), " NUMPAGES   \\* MERGEFORMAT "));
            }
            else
            {
                xElement.Add(new XAttribute(XName.Get("instr", DocX.w.NamespaceName), " NUMPAGES  \\* ROMAN  \\* MERGEFORMAT "));
            }
            XElement content = XElement.Parse("<w:r w:rsidR='001D0226' xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">\r\n                   <w:rPr>\r\n                       <w:noProof /> \r\n                   </w:rPr>\r\n                   <w:t>1</w:t> \r\n               </w:r>");
            xElement.Add(content);
            if (index == 0)
            {
                base.Xml.AddFirst(xElement);
            }
            else
            {
                Run firstRunEffectedByEdit = GetFirstRunEffectedByEdit(index);
                XElement[] array = SplitEdit(firstRunEffectedByEdit.Xml, index, EditType.ins);
                firstRunEffectedByEdit.Xml.ReplaceWith(array[0], xElement, array[1]);
            }
        }

        public void AppendPageCount(PageNumberFormat pnf)
        {
            XElement xElement = new XElement(XName.Get("fldSimple", DocX.w.NamespaceName));
            if (pnf == PageNumberFormat.normal)
            {
                xElement.Add(new XAttribute(XName.Get("instr", DocX.w.NamespaceName), " NUMPAGES   \\* MERGEFORMAT "));
            }
            else
            {
                xElement.Add(new XAttribute(XName.Get("instr", DocX.w.NamespaceName), " NUMPAGES  \\* ROMAN  \\* MERGEFORMAT "));
            }
            XElement content = XElement.Parse("<w:r w:rsidR='001D0226' xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">\r\n                   <w:rPr>\r\n                       <w:noProof /> \r\n                   </w:rPr>\r\n                   <w:t>1</w:t> \r\n               </w:r>");
            xElement.Add(content);
            base.Xml.Add(xElement);
        }
    }
}
