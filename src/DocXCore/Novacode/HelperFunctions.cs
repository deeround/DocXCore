using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.IO.Packaging;
using System.Linq;
using System.Reflection;
using System.Security.Principal;
using System.Text;
using System.Xml.Linq;

namespace Novacode
{
	internal static class HelperFunctions
	{
		public const string DOCUMENT_DOCUMENTTYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";

		public const string TEMPLATE_DOCUMENTTYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml";

		public static bool IsNullOrWhiteSpace(this string value)
		{
			if (value == null)
			{
				return true;
			}
			return string.IsNullOrEmpty(value.Trim());
		}

		internal static bool ContainsEveryChildOf(XElement desired, XElement toCheck, MatchFormattingOptions fo)
		{
			foreach (XElement item in desired.Elements())
			{
				if (!(from bElement in toCheck.Elements(item.Name)
				where bElement.GetAttribute(XName.Get("val", DocX.w.NamespaceName)) == item.GetAttribute(XName.Get("val", DocX.w.NamespaceName))
				select bElement).Any())
				{
					return false;
				}
			}
			if (fo == MatchFormattingOptions.ExactMatch)
			{
				return desired.Elements().Count() == toCheck.Elements().Count();
			}
			return true;
		}

		internal static void CreateRelsPackagePart(DocX Document, Uri uri)
		{
			PackagePart val = Document.package.CreatePart(uri, "application/vnd.openxmlformats-package.relationships+xml",  CompressionOption.Maximum);
			using (TextWriter textWriter = new StreamWriter(new PackagePartStream(val.GetStream())))
			{
				XDocument xDocument = new XDocument(new XDeclaration("1.0", "UTF-8", "yes"), new XElement(XName.Get("Relationships", DocX.rel.NamespaceName)));
				XElement root = xDocument.Root;
				xDocument.Save(textWriter);
			}
		}

		internal static int GetSize(XElement Xml)
		{
			switch (Xml.Name.LocalName)
			{
			case "tab":
				return 1;
			case "br":
			case "tr":
			case "tc":
				return 1;
			case "t":
			case "delText":
				return Xml.Value.Length;
			default:
				return 0;
			}
		}

		internal static string GetText(XElement e)
		{
			StringBuilder sb = new StringBuilder();
			GetTextRecursive(e, ref sb);
			return sb.ToString();
		}

		internal static void GetTextRecursive(XElement Xml, ref StringBuilder sb)
		{
			sb.Append(ToText(Xml));
			if (Xml.HasElements)
			{
				foreach (XElement item in Xml.Elements())
				{
					GetTextRecursive(item, ref sb);
				}
			}
		}

		internal static List<FormattedText> GetFormattedText(XElement e)
		{
			List<FormattedText> alist = new List<FormattedText>();
			GetFormattedTextRecursive(e, ref alist);
			return alist;
		}

		internal static void GetFormattedTextRecursive(XElement Xml, ref List<FormattedText> alist)
		{
			FormattedText formattedText = ToFormattedText(Xml);
			FormattedText formattedText2 = null;
			if (formattedText != null)
			{
				if (alist.Count() > 0)
				{
					formattedText2 = alist.Last();
				}
				if (formattedText2 != null && formattedText2.CompareTo(formattedText) == 0)
				{
					formattedText2.text += formattedText.text;
				}
				else
				{
					if (formattedText2 != null)
					{
						formattedText.index = formattedText2.index + formattedText2.text.Length;
					}
					alist.Add(formattedText);
				}
			}
			if (Xml.HasElements)
			{
				foreach (XElement item in Xml.Elements())
				{
					GetFormattedTextRecursive(item, ref alist);
				}
			}
		}

		internal static FormattedText ToFormattedText(XElement e)
		{
			string text = ToText(e);
			if (text == string.Empty)
			{
				return null;
			}
			while (!e.Name.Equals(XName.Get("r", DocX.w.NamespaceName)) && !e.Name.Equals(XName.Get("tabs", DocX.w.NamespaceName)))
			{
				e = e.Parent;
			}
			XElement xElement = e.Element(XName.Get("rPr", DocX.w.NamespaceName));
			FormattedText formattedText = new FormattedText();
			formattedText.text = text;
			formattedText.index = 0;
			formattedText.formatting = null;
			if (xElement != null)
			{
				formattedText.formatting = Formatting.Parse(xElement);
			}
			return formattedText;
		}

		internal static string ToText(XElement e)
		{
			switch (e.Name.LocalName)
			{
			case "tab":
			case "tc":
				return "\t";
			case "br":
			case "tr":
				return "\n";
			case "t":
			case "delText":
				if (e.Parent != null && e.Parent.Name.LocalName == "r")
				{
					XElement parent = e.Parent;
					XElement xElement = parent.Elements().FirstOrDefault((XElement a) => a.Name.LocalName == "rPr");
					if (xElement != null)
					{
						XElement xElement2 = xElement.Elements().FirstOrDefault((XElement a) => a.Name.LocalName == "caps");
						if (xElement2 != null)
						{
							return e.Value.ToUpper();
						}
					}
				}
				return e.Value;
			default:
				return "";
			}
		}

		internal static XElement CloneElement(XElement element)
		{
			return new XElement(element.Name, element.Attributes(), element.Nodes().Select(delegate(XNode n)
			{
				XElement xElement = n as XElement;
				if (xElement != null)
				{
					return CloneElement(xElement);
				}
				return n;
			}));
		}

		internal static PackagePart CreateOrGetSettingsPart(Package package)
		{
			Uri uri = new Uri("/word/settings.xml", UriKind.Relative);
			PackagePart val;
			if (!package.PartExists(uri))
			{
				val = package.CreatePart(uri, "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml",  CompressionOption.Maximum);
				PackagePart val2 = ((IEnumerable<PackagePart>)package.GetParts()).Single((PackagePart p) => p.ContentType.Equals("application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml", StringComparison.CurrentCultureIgnoreCase) || p.ContentType.Equals("application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml", StringComparison.CurrentCultureIgnoreCase));
				val2.CreateRelationship(uri, 0, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings");
				XDocument xDocument = XDocument.Parse("<?xml version='1.0' encoding='utf-8' standalone='yes'?>\r\n                <w:settings xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships' xmlns:m='http://schemas.openxmlformats.org/officeDocument/2006/math' xmlns:v='urn:schemas-microsoft-com:vml' xmlns:w10='urn:schemas-microsoft-com:office:word' xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main' xmlns:sl='http://schemas.openxmlformats.org/schemaLibrary/2006/main'>\r\n                  <w:zoom w:percent='100' />\r\n                  <w:defaultTabStop w:val='720' />\r\n                  <w:characterSpacingControl w:val='doNotCompress' />\r\n                  <w:compat />\r\n                  <w:rsids>\r\n                    <w:rsidRoot w:val='00217F62' />\r\n                    <w:rsid w:val='001915A3' />\r\n                    <w:rsid w:val='00217F62' />\r\n                    <w:rsid w:val='00A906D8' />\r\n                    <w:rsid w:val='00AB5A74' />\r\n                    <w:rsid w:val='00F071AE' />\r\n                  </w:rsids>\r\n                  <m:mathPr>\r\n                    <m:mathFont m:val='Cambria Math' />\r\n                    <m:brkBin m:val='before' />\r\n                    <m:brkBinSub m:val='--' />\r\n                    <m:smallFrac m:val='off' />\r\n                    <m:dispDef />\r\n                    <m:lMargin m:val='0' />\r\n                    <m:rMargin m:val='0' />\r\n                    <m:defJc m:val='centerGroup' />\r\n                    <m:wrapIndent m:val='1440' />\r\n                    <m:intLim m:val='subSup' />\r\n                    <m:naryLim m:val='undOvr' />\r\n                  </m:mathPr>\r\n                  <w:themeFontLang w:val='en-IE' w:bidi='ar-SA' />\r\n                  <w:clrSchemeMapping w:bg1='light1' w:t1='dark1' w:bg2='light2' w:t2='dark2' w:accent1='accent1' w:accent2='accent2' w:accent3='accent3' w:accent4='accent4' w:accent5='accent5' w:accent6='accent6' w:hyperlink='hyperlink' w:followedHyperlink='followedHyperlink' />\r\n                  <w:shapeDefaults>\r\n                    <o:shapedefaults v:ext='edit' spidmax='2050' />\r\n                    <o:shapelayout v:ext='edit'>\r\n                      <o:idmap v:ext='edit' data='1' />\r\n                    </o:shapelayout>\r\n                  </w:shapeDefaults>\r\n                  <w:decimalSymbol w:val='.' />\r\n                  <w:listSeparator w:val=',' />\r\n                </w:settings>");
				XElement xElement = xDocument.Root.Element(XName.Get("themeFontLang", DocX.w.NamespaceName));
				xElement.SetAttributeValue(XName.Get("val", DocX.w.NamespaceName), CultureInfo.CurrentCulture);
				using (TextWriter textWriter = new StreamWriter(new PackagePartStream(val.GetStream())))
				{
					xDocument.Save(textWriter);
				}
			}
			else
			{
				val = package.GetPart(uri);
			}
			return val;
		}

		internal static void CreateCustomPropertiesPart(DocX document)
		{
			PackagePart val = document.package.CreatePart(new Uri("/docProps/custom.xml", UriKind.Relative), "application/vnd.openxmlformats-officedocument.custom-properties+xml",  CompressionOption.Maximum);
			XDocument xDocument = new XDocument(new XDeclaration("1.0", "UTF-8", "yes"), new XElement(XName.Get("Properties", DocX.customPropertiesSchema.NamespaceName), new XAttribute(XNamespace.Xmlns + "vt", DocX.customVTypesSchema)));
			using (TextWriter textWriter = new StreamWriter(new PackagePartStream(val.GetStream(FileMode.Create, FileAccess.Write))))
			{
				xDocument.Save(textWriter, SaveOptions.None);
			}
			document.package.CreateRelationship(val.Uri, 0, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties");
		}

		internal static XDocument DecompressXMLResource(string manifest_resource_name)
		{
			Assembly assembly = typeof(HelperFunctions).GetTypeInfo().Assembly;
			Stream manifestResourceStream = assembly.GetManifestResourceStream(manifest_resource_name.Replace("Novacode", "DocXCore"));
			int num = 0;
			string[] manifestResourceNames = assembly.GetManifestResourceNames();
			foreach (string text in manifestResourceNames)
			{
				num++;
			}
			using (GZipStream stream = new GZipStream(manifestResourceStream, CompressionMode.Decompress))
			{
				using (TextReader textReader = new StreamReader(stream))
				{
					return XDocument.Load(textReader);
				}
			}
		}

		internal static XDocument AddDefaultNumberingXml(Package package)
		{
			PackagePart val = package.CreatePart(new Uri("/word/numbering.xml", UriKind.Relative), "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml", CompressionOption.Maximum);
			XDocument xDocument = DecompressXMLResource("Novacode.Resources.numbering.xml.gz");
			using (TextWriter textWriter = new StreamWriter(new PackagePartStream(val.GetStream(FileMode.Create, FileAccess.Write))))
			{
				xDocument.Save(textWriter, SaveOptions.None);
			}
			PackagePart val2 = ((IEnumerable<PackagePart>)package.GetParts()).Single((PackagePart p) => p.ContentType.Equals("application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml", StringComparison.CurrentCultureIgnoreCase) || p.ContentType.Equals("application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml", StringComparison.CurrentCultureIgnoreCase));
			val2.CreateRelationship(val.Uri, 0, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering");
			return xDocument;
		}

		internal static XDocument AddDefaultStylesXml(Package package)
		{
			PackagePart val = package.CreatePart(new Uri("/word/styles.xml", UriKind.Relative), "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml", CompressionOption.Maximum);
			XDocument xDocument = DecompressXMLResource("Novacode.Resources.default_styles.xml.gz");
			XElement xElement = xDocument.Root.Element(XName.Get("docDefaults", DocX.w.NamespaceName)).Element(XName.Get("rPrDefault", DocX.w.NamespaceName)).Element(XName.Get("rPr", DocX.w.NamespaceName))
				.Element(XName.Get("lang", DocX.w.NamespaceName));
			xElement.SetAttributeValue(XName.Get("val", DocX.w.NamespaceName), CultureInfo.CurrentCulture);
			using (TextWriter textWriter = new StreamWriter(new PackagePartStream(val.GetStream(FileMode.Create, FileAccess.Write))))
			{
				xDocument.Save(textWriter, SaveOptions.None);
			}
			PackagePart val2 = (from p in (IEnumerable<PackagePart>)package.GetParts()
			where p.ContentType.Equals("application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml", StringComparison.CurrentCultureIgnoreCase) || p.ContentType.Equals("application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml", StringComparison.CurrentCultureIgnoreCase)
			select p).Single();
			val2.CreateRelationship(val.Uri, 0, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles");
			return xDocument;
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

		internal static XElement CreateTable(int rowCount, int columnCount)
		{
			int[] array = new int[columnCount];
			for (int i = 0; i < columnCount; i++)
			{
				array[i] = 2310;
			}
			return CreateTable(rowCount, array);
		}

		internal static XElement CreateTable(int rowCount, int[] columnWidths)
		{
			XElement xElement = new XElement(XName.Get("tbl", DocX.w.NamespaceName), new XElement(XName.Get("tblPr", DocX.w.NamespaceName), new XElement(XName.Get("tblStyle", DocX.w.NamespaceName), new XAttribute(XName.Get("val", DocX.w.NamespaceName), "TableGrid")), new XElement(XName.Get("tblW", DocX.w.NamespaceName), new XAttribute(XName.Get("w", DocX.w.NamespaceName), "5000"), new XAttribute(XName.Get("type", DocX.w.NamespaceName), "auto")), new XElement(XName.Get("tblLook", DocX.w.NamespaceName), new XAttribute(XName.Get("val", DocX.w.NamespaceName), "04A0"))));
			for (int i = 0; i < rowCount; i++)
			{
				XElement xElement2 = new XElement(XName.Get("tr", DocX.w.NamespaceName));
				for (int j = 0; j < columnWidths.Length; j++)
				{
					XElement content = CreateTableCell();
					xElement2.Add(content);
				}
				xElement.Add(xElement2);
			}
			return xElement;
		}

		internal static XElement CreateTableCell(double w = 2310.0)
		{
			return new XElement(XName.Get("tc", DocX.w.NamespaceName), new XElement(XName.Get("tcPr", DocX.w.NamespaceName), new XElement(XName.Get("tcW", DocX.w.NamespaceName), new XAttribute(XName.Get("w", DocX.w.NamespaceName), w), new XAttribute(XName.Get("type", DocX.w.NamespaceName), "dxa"))), new XElement(XName.Get("p", DocX.w.NamespaceName), new XElement(XName.Get("pPr", DocX.w.NamespaceName))));
		}

		internal static List CreateItemInList(List list, string listText, int level = 0, ListItemType listType = ListItemType.Numbered, int? startNumber = default(int?), bool trackChanges = false, bool continueNumbering = false)
		{
			if (list.NumId == 0)
			{
				list.CreateNewNumberingNumId(level, listType, startNumber, continueNumbering);
			}
			if (listText != null)
			{
				XElement xElement = new XElement(XName.Get("p", DocX.w.NamespaceName), new XElement(XName.Get("pPr", DocX.w.NamespaceName), new XElement(XName.Get("numPr", DocX.w.NamespaceName), new XElement(XName.Get("ilvl", DocX.w.NamespaceName), new XAttribute(DocX.w + "val", level)), new XElement(XName.Get("numId", DocX.w.NamespaceName), new XAttribute(DocX.w + "val", list.NumId)))), new XElement(XName.Get("r", DocX.w.NamespaceName), new XElement(XName.Get("t", DocX.w.NamespaceName), listText)));
				if (trackChanges)
				{
					xElement = CreateEdit(EditType.ins, DateTime.Now, xElement);
				}
				if (!startNumber.HasValue)
				{
					list.AddItem(new Paragraph(list.Document, xElement, 0, ContainerType.Paragraph));
				}
				else
				{
					list.AddItemWithStartValue(new Paragraph(list.Document, xElement, 0, ContainerType.Paragraph), startNumber.Value);
				}
			}
			return list;
		}

		internal static void RenumberIDs(DocX document)
		{
			IEnumerable<XAttribute> source = from d in document.mainDoc.Descendants()
			where d.Name.LocalName == "ins" || d.Name.LocalName == "del"
			select d.Attribute(XName.Get("id", "http://schemas.openxmlformats.org/wordprocessingml/2006/main"));
			for (int i = 0; i < source.Count(); i++)
			{
				source.ElementAt(i).Value = i.ToString();
			}
		}

		internal static Paragraph GetFirstParagraphEffectedByInsert(DocX document, int index)
		{
			if (document.Paragraphs.Count() != 0 || index != 0)
			{
				foreach (Paragraph paragraph in document.Paragraphs)
				{
					if (paragraph.endIndex >= index)
					{
						return paragraph;
					}
				}
				throw new ArgumentOutOfRangeException();
			}
			return null;
		}

		internal static List<XElement> FormatInput(string text, XElement rPr)
		{
			List<XElement> list = new List<XElement>();
			XElement xElement = new XElement(DocX.w + "tab");
			XElement xElement2 = new XElement(DocX.w + "br");
			StringBuilder stringBuilder = new StringBuilder();
			if (string.IsNullOrEmpty(text))
			{
				return list;
			}
			char c = '\0';
			foreach (char c2 in text)
			{
				switch (c2)
				{
				case '\t':
					if (stringBuilder.Length > 0)
					{
						XElement xElement5 = new XElement(DocX.w + "t", stringBuilder.ToString());
						Text.PreserveSpace(xElement5);
						list.Add(new XElement(DocX.w + "r", rPr, xElement5));
						stringBuilder = new StringBuilder();
					}
					list.Add(new XElement(DocX.w + "r", rPr, xElement));
					break;
				case '\r':
					if (stringBuilder.Length > 0)
					{
						XElement xElement4 = new XElement(DocX.w + "t", stringBuilder.ToString());
						Text.PreserveSpace(xElement4);
						list.Add(new XElement(DocX.w + "r", rPr, xElement4));
						stringBuilder = new StringBuilder();
					}
					list.Add(new XElement(DocX.w + "r", rPr, xElement2));
					break;
				case '\n':
					if (c != '\r')
					{
						if (stringBuilder.Length > 0)
						{
							XElement xElement3 = new XElement(DocX.w + "t", stringBuilder.ToString());
							Text.PreserveSpace(xElement3);
							list.Add(new XElement(DocX.w + "r", rPr, xElement3));
							stringBuilder = new StringBuilder();
						}
						list.Add(new XElement(DocX.w + "r", rPr, xElement2));
					}
					break;
				default:
					stringBuilder.Append(c2);
					break;
				}
				c = c2;
			}
			if (stringBuilder.Length > 0)
			{
				XElement xElement6 = new XElement(DocX.w + "t", stringBuilder.ToString());
				Text.PreserveSpace(xElement6);
				list.Add(new XElement(DocX.w + "r", rPr, xElement6));
			}
			return list;
		}

		internal static XElement[] SplitParagraph(Paragraph p, int index)
		{
			Run firstRunEffectedByEdit = p.GetFirstRunEffectedByEdit(index);
			XElement xElement;
			XElement xElement2;
			if (firstRunEffectedByEdit.Xml.Parent.Name.LocalName == "ins")
			{
				XElement[] array = p.SplitEdit(firstRunEffectedByEdit.Xml.Parent, index, EditType.ins);
				xElement = new XElement(p.Xml.Name, p.Xml.Attributes(), firstRunEffectedByEdit.Xml.Parent.ElementsBeforeSelf(), array[0]);
				xElement2 = new XElement(p.Xml.Name, p.Xml.Attributes(), firstRunEffectedByEdit.Xml.Parent.ElementsAfterSelf(), array[1]);
			}
			else if (firstRunEffectedByEdit.Xml.Parent.Name.LocalName == "del")
			{
				XElement[] array = p.SplitEdit(firstRunEffectedByEdit.Xml.Parent, index, EditType.del);
				xElement = new XElement(p.Xml.Name, p.Xml.Attributes(), firstRunEffectedByEdit.Xml.Parent.ElementsBeforeSelf(), array[0]);
				xElement2 = new XElement(p.Xml.Name, p.Xml.Attributes(), firstRunEffectedByEdit.Xml.Parent.ElementsAfterSelf(), array[1]);
			}
			else
			{
				XElement[] array = Run.SplitRun(firstRunEffectedByEdit, index);
				xElement = new XElement(p.Xml.Name, p.Xml.Attributes(), firstRunEffectedByEdit.Xml.ElementsBeforeSelf(), array[0]);
				xElement2 = new XElement(p.Xml.Name, p.Xml.Attributes(), array[1], firstRunEffectedByEdit.Xml.ElementsAfterSelf());
			}
			if (xElement.Elements().Count() == 0)
			{
				xElement = null;
			}
			if (xElement2.Elements().Count() == 0)
			{
				xElement2 = null;
			}
			return new XElement[2]
			{
				xElement,
				xElement2
			};
		}

		internal static bool IsSameFile(Stream streamOne, Stream streamTwo)
		{
			if (streamOne.Length != streamTwo.Length)
			{
				return false;
			}
			int num;
			int num2;
			do
			{
				num = streamOne.ReadByte();
				num2 = streamTwo.ReadByte();
			}
			while (num == num2 && num != -1);
			streamOne.Position = 0L;
			streamTwo.Position = 0L;
			return num - num2 == 0;
		}

		internal static UnderlineStyle GetUnderlineStyle(string underlineStyle)
		{
			switch (underlineStyle)
			{
			case "single":
				return UnderlineStyle.singleLine;
			case "double":
				return UnderlineStyle.doubleLine;
			case "thick":
				return UnderlineStyle.thick;
			case "dotted":
				return UnderlineStyle.dotted;
			case "dottedHeavy":
				return UnderlineStyle.dottedHeavy;
			case "dash":
				return UnderlineStyle.dash;
			case "dashedHeavy":
				return UnderlineStyle.dashedHeavy;
			case "dashLong":
				return UnderlineStyle.dashLong;
			case "dashLongHeavy":
				return UnderlineStyle.dashLongHeavy;
			case "dotDash":
				return UnderlineStyle.dotDash;
			case "dashDotHeavy":
				return UnderlineStyle.dashDotHeavy;
			case "dotDotDash":
				return UnderlineStyle.dotDotDash;
			case "dashDotDotHeavy":
				return UnderlineStyle.dashDotDotHeavy;
			case "wave":
				return UnderlineStyle.wave;
			case "wavyHeavy":
				return UnderlineStyle.wavyHeavy;
			case "wavyDouble":
				return UnderlineStyle.wavyDouble;
			case "words":
				return UnderlineStyle.words;
			default:
				return UnderlineStyle.none;
			}
		}
	}
}
