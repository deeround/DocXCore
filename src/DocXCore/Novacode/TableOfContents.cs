using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml;
using System.Xml.Linq;

namespace Novacode
{
	public class TableOfContents : DocXElement
	{
		private const string HeaderStyle = "TOCHeading";

		private const int RightTabPos = 9350;

		private TableOfContents(DocX document, XElement xml, string headerStyle)
			: base(document, xml)
		{
			AssureUpdateField(document);
			AssureStyles(document, headerStyle);
		}

		internal static TableOfContents CreateTableOfContents(DocX document, string title, TableOfContentsSwitches switches, string headerStyle = null, int lastIncludeLevel = 3, int? rightTabPos = default(int?))
		{
			XmlReader reader = XmlReader.Create(new StringReader(string.Format("<w:sdt xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>\r\n                  <w:sdtPr>\r\n                    <w:docPartObj>\r\n                      <w:docPartGallery w:val='Table of Contents'/>\r\n                      <w:docPartUnique/>\r\n                    </w:docPartObj>\\\r\n                  </w:sdtPr>\r\n                  <w:sdtEndPr>\r\n                    <w:rPr>\r\n                      <w:rFonts w:asciiTheme='minorHAnsi' w:cstheme='minorBidi' w:eastAsiaTheme='minorHAnsi' w:hAnsiTheme='minorHAnsi'/>\r\n                      <w:color w:val='auto'/>\r\n                      <w:sz w:val='22'/>\r\n                      <w:szCs w:val='22'/>\r\n                      <w:lang w:eastAsia='en-US'/>\r\n                    </w:rPr>\r\n                  </w:sdtEndPr>\r\n                  <w:sdtContent>\r\n                    <w:p>\r\n                      <w:pPr>\r\n                        <w:pStyle w:val='{0}'/>\r\n                      </w:pPr>\r\n                      <w:r>\r\n                        <w:t>{1}</w:t>\r\n                      </w:r>\r\n                    </w:p>\r\n                    <w:p>\r\n                      <w:pPr>\r\n                        <w:pStyle w:val='TOC1'/>\r\n                        <w:tabs>\r\n                          <w:tab w:val='right' w:leader='dot' w:pos='{2}'/>\r\n                        </w:tabs>\r\n                        <w:rPr>\r\n                          <w:noProof/>\r\n                        </w:rPr>\r\n                      </w:pPr>\r\n                      <w:r>\r\n                        <w:fldChar w:fldCharType='begin' w:dirty='true'/>\r\n                      </w:r>\r\n                      <w:r>\r\n                        <w:instrText xml:space='preserve'> {3} </w:instrText>\r\n                      </w:r>\r\n                      <w:r>\r\n                        <w:fldChar w:fldCharType='separate'/>\r\n                      </w:r>\r\n                    </w:p>\r\n                    <w:p>\r\n                      <w:r>\r\n                        <w:rPr>\r\n                          <w:b/>\r\n                          <w:bCs/>\r\n                          <w:noProof/>\r\n                        </w:rPr>\r\n                        <w:fldChar w:fldCharType='end'/>\r\n                      </w:r>\r\n                    </w:p>\r\n                  </w:sdtContent>\r\n                </w:sdt>\r\n            ", headerStyle ?? "TOCHeading", title, rightTabPos ?? 9350, BuildSwitchString(switches, lastIncludeLevel))));
			XElement xml = XElement.Load(reader);
			return new TableOfContents(document, xml, headerStyle);
		}

		private void AssureUpdateField(DocX document)
		{
			if (!document.settings.Descendants().Any((XElement x) => x.Name.Equals(DocX.w + "updateFields")))
			{
				XElement content = new XElement(XName.Get("updateFields", DocX.w.NamespaceName), new XAttribute(DocX.w + "val", true));
				document.settings.Root.Add(content);
			}
		}

		private void AssureStyles(DocX document, string headerStyle)
		{
			if (!HasStyle(document, headerStyle, "paragraph"))
			{
				XmlReader reader = XmlReader.Create(new StringReader(string.Format("<w:style w:type='paragraph' w:styleId='{0}' xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>\r\n            <w:name w:val='TOC Heading'/>\r\n            <w:basedOn w:val='Heading1'/>\r\n            <w:next w:val='Normal'/>\r\n            <w:uiPriority w:val='39'/>\r\n            <w:semiHidden/>\r\n            <w:unhideWhenUsed/>\r\n            <w:qFormat/>\r\n            <w:rsid w:val='00E67AA6'/>\r\n            <w:pPr>\r\n              <w:outlineLvl w:val='9'/>\r\n            </w:pPr>\r\n            <w:rPr>\r\n              <w:lang w:eastAsia='nb-NO'/>\r\n            </w:rPr>\r\n          </w:style>\r\n        ", headerStyle ?? "TOCHeading")));
				XElement content = XElement.Load(reader);
				document.styles.Root.Add(content);
			}
			if (!HasStyle(document, "TOC1", "paragraph"))
			{
				XmlReader reader2 = XmlReader.Create(new StringReader(string.Format("  <w:style w:type='paragraph' w:styleId='{0}' xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>\r\n            <w:name w:val='{1}' />\r\n            <w:basedOn w:val='Normal' />\r\n            <w:next w:val='Normal' />\r\n            <w:autoRedefine />\r\n            <w:uiPriority w:val='39' />\r\n            <w:unhideWhenUsed />\r\n            <w:pPr>\r\n              <w:spacing w:after='100' />\r\n              <w:ind w:left='440' />\r\n            </w:pPr>\r\n          </w:style>\r\n        ", "TOC1", "toc 1")));
				XElement content2 = XElement.Load(reader2);
				document.styles.Root.Add(content2);
			}
			if (!HasStyle(document, "TOC2", "paragraph"))
			{
				XmlReader reader3 = XmlReader.Create(new StringReader(string.Format("  <w:style w:type='paragraph' w:styleId='{0}' xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>\r\n            <w:name w:val='{1}' />\r\n            <w:basedOn w:val='Normal' />\r\n            <w:next w:val='Normal' />\r\n            <w:autoRedefine />\r\n            <w:uiPriority w:val='39' />\r\n            <w:unhideWhenUsed />\r\n            <w:pPr>\r\n              <w:spacing w:after='100' />\r\n              <w:ind w:left='440' />\r\n            </w:pPr>\r\n          </w:style>\r\n        ", "TOC2", "toc 2")));
				XElement content3 = XElement.Load(reader3);
				document.styles.Root.Add(content3);
			}
			if (!HasStyle(document, "TOC3", "paragraph"))
			{
				XmlReader reader4 = XmlReader.Create(new StringReader(string.Format("  <w:style w:type='paragraph' w:styleId='{0}' xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>\r\n            <w:name w:val='{1}' />\r\n            <w:basedOn w:val='Normal' />\r\n            <w:next w:val='Normal' />\r\n            <w:autoRedefine />\r\n            <w:uiPriority w:val='39' />\r\n            <w:unhideWhenUsed />\r\n            <w:pPr>\r\n              <w:spacing w:after='100' />\r\n              <w:ind w:left='440' />\r\n            </w:pPr>\r\n          </w:style>\r\n        ", "TOC3", "toc 3")));
				XElement content4 = XElement.Load(reader4);
				document.styles.Root.Add(content4);
			}
			if (!HasStyle(document, "TOC4", "paragraph"))
			{
				XmlReader reader5 = XmlReader.Create(new StringReader(string.Format("  <w:style w:type='paragraph' w:styleId='{0}' xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>\r\n            <w:name w:val='{1}' />\r\n            <w:basedOn w:val='Normal' />\r\n            <w:next w:val='Normal' />\r\n            <w:autoRedefine />\r\n            <w:uiPriority w:val='39' />\r\n            <w:unhideWhenUsed />\r\n            <w:pPr>\r\n              <w:spacing w:after='100' />\r\n              <w:ind w:left='440' />\r\n            </w:pPr>\r\n          </w:style>\r\n        ", "TOC4", "toc 4")));
				XElement content5 = XElement.Load(reader5);
				document.styles.Root.Add(content5);
			}
			if (!HasStyle(document, "Hyperlink", "character"))
			{
				XmlReader reader6 = XmlReader.Create(new StringReader($"  <w:style w:type='character' w:styleId='Hyperlink' xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>\r\n            <w:name w:val='Hyperlink' />\r\n            <w:basedOn w:val='Normal' />\r\n            <w:uiPriority w:val='99' />\r\n            <w:unhideWhenUsed />\r\n            <w:rPr>\r\n              <w:color w:val='0000FF' w:themeColor='hyperlink' />\r\n              <w:u w:val='single' />\r\n            </w:rPr>\r\n          </w:style>\r\n        "));
				XElement content6 = XElement.Load(reader6);
				document.styles.Root.Add(content6);
			}
		}

		private bool HasStyle(DocX document, string value, string type)
		{
			return document.styles.Descendants().Any((XElement x) => x.Name.Equals(DocX.w + "style") && (x.Attribute(DocX.w + "type") == null || x.Attribute(DocX.w + "type").Value.Equals(type)) && x.Attribute(DocX.w + "styleId") != null && x.Attribute(DocX.w + "styleId").Value.Equals(value));
		}

		private static string BuildSwitchString(TableOfContentsSwitches switches, int lastIncludeLevel)
		{
			IEnumerable<TableOfContentsSwitches> source = Enum.GetValues(typeof(TableOfContentsSwitches)).Cast<TableOfContentsSwitches>();
			string text = "TOC";
			foreach (TableOfContentsSwitches item in from s in source
			where s != 0 && switches.HasFlag(s)
			select s)
			{
				text = text + " " + item.EnumDescription();
				if (item == TableOfContentsSwitches.O)
				{
					text += $" '{1}-{lastIncludeLevel}'";
				}
			}
			return text;
		}
	}
}
