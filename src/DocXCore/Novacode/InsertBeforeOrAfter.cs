using System;
using System.Linq;
using System.Xml.Linq;

namespace Novacode
{
	public abstract class InsertBeforeOrAfter : DocXElement
	{
		public InsertBeforeOrAfter(DocX document, XElement xml)
			: base(document, xml)
		{
		}

		public virtual void InsertPageBreakBeforeSelf()
		{
			XElement content = new XElement(XName.Get("p", DocX.w.NamespaceName), new XElement(XName.Get("r", DocX.w.NamespaceName), new XElement(XName.Get("br", DocX.w.NamespaceName), new XAttribute(XName.Get("type", DocX.w.NamespaceName), "page"))));
			base.Xml.AddBeforeSelf(content);
		}

		public virtual void InsertPageBreakAfterSelf()
		{
			XElement content = new XElement(XName.Get("p", DocX.w.NamespaceName), new XElement(XName.Get("r", DocX.w.NamespaceName), new XElement(XName.Get("br", DocX.w.NamespaceName), new XAttribute(XName.Get("type", DocX.w.NamespaceName), "page"))));
			base.Xml.AddAfterSelf(content);
		}

		public virtual Paragraph InsertParagraphBeforeSelf(Paragraph p)
		{
			base.Xml.AddBeforeSelf(p.Xml);
			XElement xElement2 = p.Xml = base.Xml.ElementsBeforeSelf().First();
			return p;
		}

		public virtual Paragraph InsertParagraphAfterSelf(Paragraph p)
		{
			base.Xml.AddAfterSelf(p.Xml);
			XElement xml = base.Xml.ElementsAfterSelf().First();
			if (this is Paragraph)
			{
				return new Paragraph(base.Document, xml, (this as Paragraph).endIndex);
			}
			p.Xml = xml;
			return p;
		}

		public virtual Paragraph InsertParagraphBeforeSelf(string text)
		{
			return InsertParagraphBeforeSelf(text,false, new Formatting());
		}

		public virtual Paragraph InsertParagraphAfterSelf(string text)
		{
			return InsertParagraphAfterSelf(text, false, new Formatting());
		}

		public virtual Paragraph InsertParagraphBeforeSelf(string text, bool trackChanges)
		{
			return InsertParagraphBeforeSelf(text, trackChanges, new Formatting());
		}

		public virtual Paragraph InsertParagraphAfterSelf(string text, bool trackChanges)
		{
			return InsertParagraphAfterSelf(text, trackChanges, new Formatting());
		}

		public virtual Paragraph InsertParagraphBeforeSelf(string text, bool trackChanges, Formatting formatting)
		{
			XElement content = new XElement(XName.Get("p", DocX.w.NamespaceName), new XElement(XName.Get("pPr", DocX.w.NamespaceName)), HelperFunctions.FormatInput(text, formatting.Xml));
			if (trackChanges)
			{
				content = Paragraph.CreateEdit(EditType.ins, DateTime.Now, content);
			}
			base.Xml.AddBeforeSelf(content);
			XElement xml = base.Xml.ElementsBeforeSelf().Last();
			return new Paragraph(base.Document, xml, -1);
		}

		public virtual Paragraph InsertParagraphAfterSelf(string text, bool trackChanges, Formatting formatting)
		{
			XElement content = new XElement(XName.Get("p", DocX.w.NamespaceName), new XElement(XName.Get("pPr", DocX.w.NamespaceName)), HelperFunctions.FormatInput(text, formatting.Xml));
			if (trackChanges)
			{
				content = Paragraph.CreateEdit(EditType.ins, DateTime.Now, content);
			}
			base.Xml.AddAfterSelf(content);
			XElement xml = base.Xml.ElementsAfterSelf().First();
			return new Paragraph(base.Document, xml, -1);
		}

		public virtual Table InsertTableAfterSelf(int rowCount, int columnCount)
		{
			XElement content = HelperFunctions.CreateTable(rowCount, columnCount);
			base.Xml.AddAfterSelf(content);
			XElement xml = base.Xml.ElementsAfterSelf().First();
			return new Table(base.Document, xml)
			{
				mainPart = mainPart
			};
		}

		public virtual Table InsertTableAfterSelf(Table t)
		{
			base.Xml.AddAfterSelf(t.Xml);
			XElement xml = base.Xml.ElementsAfterSelf().First();
			return new Table(base.Document, xml)
			{
				mainPart = mainPart
			};
		}

		public virtual Table InsertTableBeforeSelf(int rowCount, int columnCount)
		{
			XElement content = HelperFunctions.CreateTable(rowCount, columnCount);
			base.Xml.AddBeforeSelf(content);
			XElement xml = base.Xml.ElementsBeforeSelf().Last();
			return new Table(base.Document, xml)
			{
				mainPart = mainPart
			};
		}

		public virtual Table InsertTableBeforeSelf(Table t)
		{
			base.Xml.AddBeforeSelf(t.Xml);
			XElement xml = base.Xml.ElementsBeforeSelf().Last();
			return new Table(base.Document, xml)
			{
				mainPart = mainPart
			};
		}
	}
}
