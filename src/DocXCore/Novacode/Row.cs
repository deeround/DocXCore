using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Xml.Linq;

namespace Novacode
{
	public class Row : Container
	{
		internal Table table;

		private const string _hRule_Exact = "exact";

		private const string _hRule_AtLeast = "atLeast";

		public int ColumnCount
		{
			get
			{
				int num = 0;
				num += gridAfter;
				foreach (Cell cell in Cells)
				{
					if (cell.GridSpan != 0)
					{
						num += cell.GridSpan - 1;
					}
				}
				return Cells.Count + num;
			}
		}

		public int gridAfter
		{
			get
			{
				int num = 0;
				XElement xElement = base.Xml.Element(XName.Get("trPr", DocX.w.NamespaceName));
				if (xElement != null)
				{
					XAttribute xAttribute = xElement.Element(XName.Get("gridAfter", DocX.w.NamespaceName))?.Attribute(XName.Get("val", DocX.w.NamespaceName));
					if (xAttribute != null)
					{
						num += int.Parse(xAttribute.Value);
					}
				}
				return num;
			}
		}

		public List<Cell> Cells => (from c in base.Xml.Elements(XName.Get("tc", DocX.w.NamespaceName))
		select new Cell(this, base.Document, c)).ToList();

		public override ReadOnlyCollection<Paragraph> Paragraphs
		{
			get
			{
				List<Paragraph> list = (from p in base.Xml.Descendants(DocX.w + "p")
				select new Paragraph(base.Document, p, 0)).ToList();
				foreach (Paragraph item in list)
				{
					item.PackagePart = table.mainPart;
				}
				return list.AsReadOnly();
			}
		}

		public double Height
		{
			get
			{
				XAttribute xAttribute = (base.Xml.Element(XName.Get("trPr", DocX.w.NamespaceName))?.Element(XName.Get("trHeight", DocX.w.NamespaceName)))?.Attribute(XName.Get("val", DocX.w.NamespaceName));
				if (xAttribute == null)
				{
					return double.NaN;
				}
				if (!double.TryParse(xAttribute.Value, out double result))
				{
					xAttribute.Remove();
					return double.NaN;
				}
				return result / 15.0;
			}
			set
			{
				SetHeight(value, exact: true);
			}
		}

		public double MinHeight
		{
			get
			{
				return Height;
			}
			set
			{
				SetHeight(value, exact: false);
			}
		}

		public bool TableHeader
		{
			get
			{
				XElement xElement = base.Xml.Element(XName.Get("trPr", DocX.w.NamespaceName));
				if (xElement == null)
				{
					return false;
				}
				XElement xElement2 = xElement.Element(XName.Get("tblHeader", DocX.w.NamespaceName));
				return xElement2 != null;
			}
			set
			{
				XElement xElement = base.Xml.Element(XName.Get("trPr", DocX.w.NamespaceName));
				if (xElement == null)
				{
					base.Xml.SetElementValue(XName.Get("trPr", DocX.w.NamespaceName), string.Empty);
					xElement = base.Xml.Element(XName.Get("trPr", DocX.w.NamespaceName));
				}
				XElement xElement2 = xElement.Element(XName.Get("tblHeader", DocX.w.NamespaceName));
				if (xElement2 == null && value)
				{
					xElement.SetElementValue(XName.Get("tblHeader", DocX.w.NamespaceName), string.Empty);
				}
				if (xElement2 != null && !value)
				{
					xElement2.Remove();
				}
			}
		}

		public bool BreakAcrossPages
		{
			get
			{
				XElement xElement = base.Xml.Element(XName.Get("trPr", DocX.w.NamespaceName))?.Element(XName.Get("cantSplit", DocX.w.NamespaceName));
				return xElement == null;
			}
			set
			{
				if (!value)
				{
					XElement xElement = base.Xml.Element(XName.Get("trPr", DocX.w.NamespaceName));
					if (xElement == null)
					{
						base.Xml.SetElementValue(XName.Get("trPr", DocX.w.NamespaceName), string.Empty);
						xElement = base.Xml.Element(XName.Get("trPr", DocX.w.NamespaceName));
					}
					XElement xElement2 = xElement.Element(XName.Get("cantSplit", DocX.w.NamespaceName));
					if (xElement2 == null)
					{
						xElement.SetElementValue(XName.Get("cantSplit", DocX.w.NamespaceName), string.Empty);
					}
				}
				else
				{
					(base.Xml.Element(XName.Get("trPr", DocX.w.NamespaceName))?.Element(XName.Get("cantSplit", DocX.w.NamespaceName)))?.Remove();
				}
			}
		}

		public void Remove()
		{
			XElement parent = base.Xml.Parent;
			base.Xml.Remove();
			if (!parent.Elements(XName.Get("tr", DocX.w.NamespaceName)).Any())
			{
				parent.Remove();
			}
		}

		internal Row(Table table, DocX document, XElement xml)
			: base(document, xml)
		{
			this.table = table;
			mainPart = table.mainPart;
		}

		private void SetHeight(double height, bool exact)
		{
			XElement xElement = base.Xml.Element(XName.Get("trPr", DocX.w.NamespaceName));
			if (xElement == null)
			{
				base.Xml.SetElementValue(XName.Get("trPr", DocX.w.NamespaceName), string.Empty);
				xElement = base.Xml.Element(XName.Get("trPr", DocX.w.NamespaceName));
			}
			XElement xElement2 = xElement.Element(XName.Get("trHeight", DocX.w.NamespaceName));
			if (xElement2 == null)
			{
				xElement.SetElementValue(XName.Get("trHeight", DocX.w.NamespaceName), string.Empty);
				xElement2 = xElement.Element(XName.Get("trHeight", DocX.w.NamespaceName));
			}
			xElement2.SetAttributeValue(XName.Get("hRule", DocX.w.NamespaceName), exact ? "exact" : "atLeast");
			xElement2.SetAttributeValue(XName.Get("val", DocX.w.NamespaceName), (height * 15.0).ToString());
		}

		public void MergeCells(int startIndex, int endIndex)
		{
			if (startIndex < 0 || endIndex <= startIndex || endIndex > Cells.Count + 1)
			{
				throw new IndexOutOfRangeException();
			}
			int num = 0;
			foreach (Cell item in Cells.Where((Cell z, int i) => i > startIndex && i <= endIndex))
			{
				XElement xElement = item.Xml.Element(XName.Get("tcPr", DocX.w.NamespaceName))?.Element(XName.Get("gridSpan", DocX.w.NamespaceName));
				if (xElement != null)
				{
					XAttribute xAttribute = xElement.Attribute(XName.Get("val", DocX.w.NamespaceName));
					int result = default(int);
					if (xAttribute != null && int.TryParse(xAttribute.Value, out result))
					{
						num += result - 1;
					}
				}
				Cells[startIndex].Xml.Add(item.Xml.Elements(XName.Get("p", DocX.w.NamespaceName)));
				item.Xml.Remove();
			}
			XElement xElement2 = Cells[startIndex].Xml.Element(XName.Get("tcPr", DocX.w.NamespaceName));
			if (xElement2 == null)
			{
				Cells[startIndex].Xml.SetElementValue(XName.Get("tcPr", DocX.w.NamespaceName), string.Empty);
				xElement2 = Cells[startIndex].Xml.Element(XName.Get("tcPr", DocX.w.NamespaceName));
			}
			XElement xElement3 = xElement2.Element(XName.Get("gridSpan", DocX.w.NamespaceName));
			if (xElement3 == null)
			{
				xElement2.SetElementValue(XName.Get("gridSpan", DocX.w.NamespaceName), string.Empty);
				xElement3 = xElement2.Element(XName.Get("gridSpan", DocX.w.NamespaceName));
			}
			XAttribute xAttribute2 = xElement3.Attribute(XName.Get("val", DocX.w.NamespaceName));
			int result2 = 0;
			if (xAttribute2 != null && int.TryParse(xAttribute2.Value, out result2))
			{
				num += result2 - 1;
			}
			xElement3.SetAttributeValue(XName.Get("val", DocX.w.NamespaceName), (num + (endIndex - startIndex + 1)).ToString());
		}
	}
}
