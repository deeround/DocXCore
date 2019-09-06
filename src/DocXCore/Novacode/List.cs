using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace Novacode
{
	public class List : InsertBeforeOrAfter
	{
		public List<Paragraph> Items
		{
			get;
			private set;
		}

		public int NumId
		{
			get;
			private set;
		}

		public ListItemType? ListType
		{
			get;
			private set;
		}

		internal List(DocX document, XElement xml)
			: base(document, xml)
		{
			Items = new List<Paragraph>();
			ListType = null;
		}

		public void AddItem(Paragraph paragraph)
		{
			if (paragraph.IsListItem)
			{
				XElement xElement = paragraph.Xml.Descendants().First((XElement s) => s.Name.LocalName == "numId");
				int numId = int.Parse(xElement.Attribute(DocX.w + "val").Value);
				if (!CanAddListItem(paragraph))
				{
					throw new InvalidOperationException("New list items can only be added to this list if they are have the same numId.");
				}
				NumId = numId;
				Items.Add(paragraph);
			}
		}

		public void AddItemWithStartValue(Paragraph paragraph, int start)
		{
			UpdateNumberingForLevelStartNumber(int.Parse(paragraph.IndentLevel.ToString()), start);
			if (ContainsLevel(start))
			{
				throw new InvalidOperationException("Cannot add a paragraph with a start value if another element already exists in this list with that level.");
			}
			AddItem(paragraph);
		}

		private void UpdateNumberingForLevelStartNumber(int iLevel, int start)
		{
			XElement abstractNum = GetAbstractNum(NumId);
			XElement xElement = abstractNum.Descendants().First((XElement el) => el.Name.LocalName == "lvl" && el.GetAttribute(DocX.w + "ilvl") == iLevel.ToString());
			xElement.Descendants().First((XElement el) => el.Name.LocalName == "start").SetAttributeValue(DocX.w + "val", start);
		}

		public bool CanAddListItem(Paragraph paragraph)
		{
			if (paragraph.IsListItem)
			{
				XElement xElement = paragraph.Xml.Descendants().First((XElement s) => s.Name.LocalName == "numId");
				int num = int.Parse(xElement.Attribute(DocX.w + "val").Value);
				if (NumId == 0 || (num == NumId && num > 0))
				{
					return true;
				}
			}
			return false;
		}

		public bool ContainsLevel(int ilvl)
		{
			return Items.Any((Paragraph i) => i.ParagraphNumberProperties.Descendants().First((XElement el) => el.Name.LocalName == "ilvl").Value == ilvl.ToString());
		}

		internal void CreateNewNumberingNumId(int level = 0, ListItemType listType = ListItemType.Numbered, int? startNumber = default(int?), bool continueNumbering = false)
		{
			ValidateDocXNumberingPartExists();
			if (base.Document.numbering.Root == null)
			{
				throw new InvalidOperationException("Numbering section did not instantiate properly.");
			}
			ListType = listType;
			int numId = GetMaxNumId() + 1;
			int num = GetMaxAbstractNumId() + 1;
			XDocument xDocument;
			switch (listType)
			{
			case ListItemType.Bulleted:
				xDocument = HelperFunctions.DecompressXMLResource("Novacode.Resources.numbering.default_bullet_abstract.xml.gz");
				break;
			case ListItemType.Numbered:
				xDocument = HelperFunctions.DecompressXMLResource("Novacode.Resources.numbering.default_decimal_abstract.xml.gz");
				break;
			default:
				throw new InvalidOperationException($"Unable to deal with ListItemType: {listType.ToString()}.");
			}
			XElement xElement2 = xDocument.Descendants().Single((XElement d) => d.Name.LocalName == "abstractNum");
			xElement2.SetAttributeValue(DocX.w + "abstractNumId", num);
			XElement abstractNumXml = GetAbstractNumXml(num, numId, startNumber, continueNumbering);
			XElement xElement3 = base.Document.numbering.Root.Descendants().LastOrDefault((XElement xElement) => xElement.Name.LocalName == "abstractNum");
			XElement xElement4 = base.Document.numbering.Root.Descendants().LastOrDefault((XElement xElement) => xElement.Name.LocalName == "num");
			if (xElement3 == null || xElement4 == null)
			{
				base.Document.numbering.Root.Add(xElement2);
				base.Document.numbering.Root.Add(abstractNumXml);
			}
			else
			{
				xElement3.AddAfterSelf(xElement2);
				xElement4.AddAfterSelf(abstractNumXml);
			}
			NumId = numId;
		}

		private XElement GetAbstractNumXml(int abstractNumId, int numId, int? startNumber, bool continueNumbering)
		{
			XElement xElement = new XElement(XName.Get("startOverride", DocX.w.NamespaceName), new XAttribute(DocX.w + "val", startNumber ?? 1));
			XElement xElement2 = new XElement(XName.Get("lvlOverride", DocX.w.NamespaceName), new XAttribute(DocX.w + "ilvl", 0), xElement);
			XElement xElement3 = new XElement(XName.Get("abstractNumId", DocX.w.NamespaceName), new XAttribute(DocX.w + "val", abstractNumId));
			return continueNumbering ? new XElement(XName.Get("num", DocX.w.NamespaceName), new XAttribute(DocX.w + "numId", numId), xElement3) : new XElement(XName.Get("num", DocX.w.NamespaceName), new XAttribute(DocX.w + "numId", numId), xElement3, xElement2);
		}

		private int GetMaxNumId()
		{
			if (base.Document.numbering == null)
			{
				return 0;
			}
			List<XElement> source = (from d in base.Document.numbering.Descendants()
			where d.Name.LocalName == "num"
			select d).ToList();
			if (source.Any())
			{
				return source.Attributes(DocX.w + "numId").Max((XAttribute e) => int.Parse(e.Value));
			}
			return 0;
		}

		private int GetMaxAbstractNumId()
		{
			if (base.Document.numbering == null)
			{
				return -1;
			}
			List<XElement> source = (from d in base.Document.numbering.Descendants()
			where d.Name.LocalName == "abstractNum"
			select d).ToList();
			if (source.Any())
			{
				return source.Attributes(DocX.w + "abstractNumId").Max((XAttribute e) => int.Parse(e.Value));
			}
			return -1;
		}

		internal XElement GetAbstractNum(int numId)
		{
			XElement xElement = base.Document.numbering.Descendants().First((XElement d) => d.Name.LocalName == "num" && d.GetAttribute(DocX.w + "numId").Equals(numId.ToString()));
			XElement abstractNumId = xElement.Descendants().First((XElement d) => d.Name.LocalName == "abstractNumId");
			return base.Document.numbering.Descendants().First((XElement d) => d.Name.LocalName == "abstractNum" && d.GetAttribute("abstractNumId").Equals(abstractNumId.Value));
		}

		private void ValidateDocXNumberingPartExists()
		{
			Uri uri = new Uri("/word/numbering.xml", UriKind.Relative);
			if (!base.Document.package.PartExists(uri))
			{
				base.Document.numbering = HelperFunctions.AddDefaultNumberingXml(base.Document.package);
			}
		}
	}
}
