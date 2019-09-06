using System;
using System.Collections.ObjectModel;
using System.Drawing;
using System.Globalization;
using System.Xml.Linq;

namespace Novacode
{
	public class Cell : Container
	{
		internal Row row;

		public override ReadOnlyCollection<Paragraph> Paragraphs
		{
			get
			{
				ReadOnlyCollection<Paragraph> paragraphs = base.Paragraphs;
				foreach (Paragraph item in paragraphs)
				{
					item.PackagePart = row.table.mainPart;
				}
				return paragraphs;
			}
		}

		public int GridSpan
		{
			get
			{
				int result = 0;
				XElement xElement = base.Xml.Element(XName.Get("tcPr", DocX.w.NamespaceName))?.Element(XName.Get("gridSpan", DocX.w.NamespaceName));
				if (xElement != null)
				{
					XAttribute xAttribute = xElement.Attribute(XName.Get("val", DocX.w.NamespaceName));
					int result2 = default(int);
					if (xAttribute != null && int.TryParse(xAttribute.Value, out result2))
					{
						result = result2;
					}
				}
				return result;
			}
		}

		public VerticalAlignment VerticalAlignment
		{
			get
			{
				XAttribute xAttribute = (base.Xml.Element(XName.Get("tcPr", DocX.w.NamespaceName))?.Element(XName.Get("vAlign", DocX.w.NamespaceName)))?.Attribute(XName.Get("val", DocX.w.NamespaceName));
				if (xAttribute != null)
				{
					try
					{
						return (VerticalAlignment)Enum.Parse(typeof(VerticalAlignment), xAttribute.Value, ignoreCase: true);
					}
					catch
					{
						xAttribute.Remove();
						return VerticalAlignment.Center;
					}
				}
				return VerticalAlignment.Center;
			}
			set
			{
				XElement xElement = base.Xml.Element(XName.Get("tcPr", DocX.w.NamespaceName));
				if (xElement == null)
				{
					base.Xml.SetElementValue(XName.Get("tcPr", DocX.w.NamespaceName), string.Empty);
					xElement = base.Xml.Element(XName.Get("tcPr", DocX.w.NamespaceName));
				}
				XElement xElement2 = xElement.Element(XName.Get("vAlign", DocX.w.NamespaceName));
				if (xElement2 == null)
				{
					xElement.SetElementValue(XName.Get("vAlign", DocX.w.NamespaceName), string.Empty);
					xElement2 = xElement.Element(XName.Get("vAlign", DocX.w.NamespaceName));
				}
				xElement2.SetAttributeValue(XName.Get("val", DocX.w.NamespaceName), value.ToString().ToLower());
			}
		}

		public Color Shading
		{
			get
			{
				XAttribute xAttribute = (base.Xml.Element(XName.Get("tcPr", DocX.w.NamespaceName))?.Element(XName.Get("shd", DocX.w.NamespaceName)))?.Attribute(XName.Get("fill", DocX.w.NamespaceName));
				if (xAttribute == null)
				{
					return Color.White;
				}
				return ColorTranslator.FromHtml($"#{xAttribute.Value}");
			}
			set
			{
				XElement xElement = base.Xml.Element(XName.Get("tcPr", DocX.w.NamespaceName));
				if (xElement == null)
				{
					base.Xml.SetElementValue(XName.Get("tcPr", DocX.w.NamespaceName), string.Empty);
					xElement = base.Xml.Element(XName.Get("tcPr", DocX.w.NamespaceName));
				}
				XElement xElement2 = xElement.Element(XName.Get("shd", DocX.w.NamespaceName));
				if (xElement2 == null)
				{
					xElement.SetElementValue(XName.Get("shd", DocX.w.NamespaceName), string.Empty);
					xElement2 = xElement.Element(XName.Get("shd", DocX.w.NamespaceName));
				}
				xElement2.SetAttributeValue(XName.Get("val", DocX.w.NamespaceName), "clear");
				xElement2.SetAttributeValue(XName.Get("color", DocX.w.NamespaceName), "auto");
				xElement2.SetAttributeValue(XName.Get("fill", DocX.w.NamespaceName), value.ToHex());
			}
		}

		public double Width
		{
			get
			{
				XAttribute xAttribute = (base.Xml.Element(XName.Get("tcPr", DocX.w.NamespaceName))?.Element(XName.Get("tcW", DocX.w.NamespaceName)))?.Attribute(XName.Get("w", DocX.w.NamespaceName));
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
				XElement xElement = base.Xml.Element(XName.Get("tcPr", DocX.w.NamespaceName));
				if (xElement == null)
				{
					base.Xml.SetElementValue(XName.Get("tcPr", DocX.w.NamespaceName), string.Empty);
					xElement = base.Xml.Element(XName.Get("tcPr", DocX.w.NamespaceName));
				}
				XElement xElement2 = xElement.Element(XName.Get("tcW", DocX.w.NamespaceName));
				if (xElement2 == null)
				{
					xElement.SetElementValue(XName.Get("tcW", DocX.w.NamespaceName), string.Empty);
					xElement2 = xElement.Element(XName.Get("tcW", DocX.w.NamespaceName));
				}
				if (value == -1.0)
				{
					xElement2.Remove();
				}
				else
				{
					xElement2.SetAttributeValue(XName.Get("type", DocX.w.NamespaceName), "dxa");
					xElement2.SetAttributeValue(XName.Get("w", DocX.w.NamespaceName), (value * 15.0).ToString());
				}
			}
		}

		public double MarginLeft
		{
			get
			{
				XElement xElement = base.Xml.Element(XName.Get("tcPr", DocX.w.NamespaceName));
				if (xElement == null)
				{
					return double.NaN;
				}
				XAttribute xAttribute = (xElement.Element(XName.Get("tcMar", DocX.w.NamespaceName))?.Element(XName.Get("left", DocX.w.NamespaceName)))?.Attribute(XName.Get("w", DocX.w.NamespaceName));
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
				XElement xElement = base.Xml.Element(XName.Get("tcPr", DocX.w.NamespaceName));
				if (xElement == null)
				{
					base.Xml.SetElementValue(XName.Get("tcPr", DocX.w.NamespaceName), string.Empty);
					xElement = base.Xml.Element(XName.Get("tcPr", DocX.w.NamespaceName));
				}
				XElement xElement2 = xElement.Element(XName.Get("tcMar", DocX.w.NamespaceName));
				if (xElement2 == null)
				{
					xElement.SetElementValue(XName.Get("tcMar", DocX.w.NamespaceName), string.Empty);
					xElement2 = xElement.Element(XName.Get("tcMar", DocX.w.NamespaceName));
				}
				XElement xElement3 = xElement2.Element(XName.Get("left", DocX.w.NamespaceName));
				if (xElement3 == null)
				{
					xElement2.SetElementValue(XName.Get("left", DocX.w.NamespaceName), string.Empty);
					xElement3 = xElement2.Element(XName.Get("left", DocX.w.NamespaceName));
				}
				xElement3.SetAttributeValue(XName.Get("type", DocX.w.NamespaceName), "dxa");
				xElement3.SetAttributeValue(XName.Get("w", DocX.w.NamespaceName), (value * 15.0).ToString());
			}
		}

		public double MarginRight
		{
			get
			{
				XElement xElement = base.Xml.Element(XName.Get("tcPr", DocX.w.NamespaceName));
				if (xElement == null)
				{
					return double.NaN;
				}
				XAttribute xAttribute = (xElement.Element(XName.Get("tcMar", DocX.w.NamespaceName))?.Element(XName.Get("right", DocX.w.NamespaceName)))?.Attribute(XName.Get("w", DocX.w.NamespaceName));
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
				XElement xElement = base.Xml.Element(XName.Get("tcPr", DocX.w.NamespaceName));
				if (xElement == null)
				{
					base.Xml.SetElementValue(XName.Get("tcPr", DocX.w.NamespaceName), string.Empty);
					xElement = base.Xml.Element(XName.Get("tcPr", DocX.w.NamespaceName));
				}
				XElement xElement2 = xElement.Element(XName.Get("tcMar", DocX.w.NamespaceName));
				if (xElement2 == null)
				{
					xElement.SetElementValue(XName.Get("tcMar", DocX.w.NamespaceName), string.Empty);
					xElement2 = xElement.Element(XName.Get("tcMar", DocX.w.NamespaceName));
				}
				XElement xElement3 = xElement2.Element(XName.Get("right", DocX.w.NamespaceName));
				if (xElement3 == null)
				{
					xElement2.SetElementValue(XName.Get("right", DocX.w.NamespaceName), string.Empty);
					xElement3 = xElement2.Element(XName.Get("right", DocX.w.NamespaceName));
				}
				xElement3.SetAttributeValue(XName.Get("type", DocX.w.NamespaceName), "dxa");
				xElement3.SetAttributeValue(XName.Get("w", DocX.w.NamespaceName), (value * 15.0).ToString());
			}
		}

		public double MarginTop
		{
			get
			{
				XElement xElement = base.Xml.Element(XName.Get("tcPr", DocX.w.NamespaceName));
				if (xElement == null)
				{
					return double.NaN;
				}
				XAttribute xAttribute = (xElement.Element(XName.Get("tcMar", DocX.w.NamespaceName))?.Element(XName.Get("top", DocX.w.NamespaceName)))?.Attribute(XName.Get("w", DocX.w.NamespaceName));
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
				XElement xElement = base.Xml.Element(XName.Get("tcPr", DocX.w.NamespaceName));
				if (xElement == null)
				{
					base.Xml.SetElementValue(XName.Get("tcPr", DocX.w.NamespaceName), string.Empty);
					xElement = base.Xml.Element(XName.Get("tcPr", DocX.w.NamespaceName));
				}
				XElement xElement2 = xElement.Element(XName.Get("tcMar", DocX.w.NamespaceName));
				if (xElement2 == null)
				{
					xElement.SetElementValue(XName.Get("tcMar", DocX.w.NamespaceName), string.Empty);
					xElement2 = xElement.Element(XName.Get("tcMar", DocX.w.NamespaceName));
				}
				XElement xElement3 = xElement2.Element(XName.Get("top", DocX.w.NamespaceName));
				if (xElement3 == null)
				{
					xElement2.SetElementValue(XName.Get("top", DocX.w.NamespaceName), string.Empty);
					xElement3 = xElement2.Element(XName.Get("top", DocX.w.NamespaceName));
				}
				xElement3.SetAttributeValue(XName.Get("type", DocX.w.NamespaceName), "dxa");
				xElement3.SetAttributeValue(XName.Get("w", DocX.w.NamespaceName), (value * 15.0).ToString());
			}
		}

		public double MarginBottom
		{
			get
			{
				XElement xElement = base.Xml.Element(XName.Get("tcPr", DocX.w.NamespaceName))?.Element(XName.Get("tcMar", DocX.w.NamespaceName));
				if (xElement == null)
				{
					return double.NaN;
				}
				XAttribute xAttribute = xElement.Element(XName.Get("bottom", DocX.w.NamespaceName))?.Attribute(XName.Get("w", DocX.w.NamespaceName));
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
				XElement xElement = base.Xml.Element(XName.Get("tcPr", DocX.w.NamespaceName));
				if (xElement == null)
				{
					base.Xml.SetElementValue(XName.Get("tcPr", DocX.w.NamespaceName), string.Empty);
					xElement = base.Xml.Element(XName.Get("tcPr", DocX.w.NamespaceName));
				}
				XElement xElement2 = xElement.Element(XName.Get("tcMar", DocX.w.NamespaceName));
				if (xElement2 == null)
				{
					xElement.SetElementValue(XName.Get("tcMar", DocX.w.NamespaceName), string.Empty);
					xElement2 = xElement.Element(XName.Get("tcMar", DocX.w.NamespaceName));
				}
				XElement xElement3 = xElement2.Element(XName.Get("bottom", DocX.w.NamespaceName));
				if (xElement3 == null)
				{
					xElement2.SetElementValue(XName.Get("bottom", DocX.w.NamespaceName), string.Empty);
					xElement3 = xElement2.Element(XName.Get("bottom", DocX.w.NamespaceName));
				}
				xElement3.SetAttributeValue(XName.Get("type", DocX.w.NamespaceName), "dxa");
				xElement3.SetAttributeValue(XName.Get("w", DocX.w.NamespaceName), (value * 15.0).ToString());
			}
		}

		public Color FillColor
		{
			get
			{
				XAttribute xAttribute = (base.Xml.Element(XName.Get("tcPr", DocX.w.NamespaceName))?.Element(XName.Get("shd", DocX.w.NamespaceName)))?.Attribute(XName.Get("fill", DocX.w.NamespaceName));
				if (xAttribute == null)
				{
					return Color.Empty;
				}
				int argb = int.Parse(xAttribute.Value.Replace("#", ""), NumberStyles.HexNumber);
				return Color.FromArgb(argb);
			}
			set
			{
				XElement xElement = base.Xml.Element(XName.Get("tcPr", DocX.w.NamespaceName));
				if (xElement == null)
				{
					base.Xml.SetElementValue(XName.Get("tcPr", DocX.w.NamespaceName), string.Empty);
					xElement = base.Xml.Element(XName.Get("tcPr", DocX.w.NamespaceName));
				}
				XElement xElement2 = xElement.Element(XName.Get("shd", DocX.w.NamespaceName));
				if (xElement2 == null)
				{
					xElement.SetElementValue(XName.Get("shd", DocX.w.NamespaceName), string.Empty);
					xElement2 = xElement.Element(XName.Get("shd", DocX.w.NamespaceName));
				}
				xElement2.SetAttributeValue(XName.Get("val", DocX.w.NamespaceName), "clear");
				xElement2.SetAttributeValue(XName.Get("color", DocX.w.NamespaceName), "auto");
				xElement2.SetAttributeValue(XName.Get("fill", DocX.w.NamespaceName), value.ToHex());
			}
		}

		public TextDirection TextDirection
		{
			get
			{
				XAttribute xAttribute = (base.Xml.Element(XName.Get("tcPr", DocX.w.NamespaceName))?.Element(XName.Get("textDirection", DocX.w.NamespaceName)))?.Attribute(XName.Get("val", DocX.w.NamespaceName));
				if (xAttribute != null)
				{
					try
					{
						return (TextDirection)Enum.Parse(typeof(TextDirection), xAttribute.Value, ignoreCase: true);
					}
					catch
					{
						xAttribute.Remove();
						return TextDirection.right;
					}
				}
				return TextDirection.right;
			}
			set
			{
				XElement xElement = base.Xml.Element(XName.Get("tcPr", DocX.w.NamespaceName));
				if (xElement == null)
				{
					base.Xml.SetElementValue(XName.Get("tcPr", DocX.w.NamespaceName), string.Empty);
					xElement = base.Xml.Element(XName.Get("tcPr", DocX.w.NamespaceName));
				}
				XElement xElement2 = xElement.Element(XName.Get("textDirection", DocX.w.NamespaceName));
				if (xElement2 == null)
				{
					xElement.SetElementValue(XName.Get("textDirection", DocX.w.NamespaceName), string.Empty);
					xElement2 = xElement.Element(XName.Get("textDirection", DocX.w.NamespaceName));
				}
				xElement2.SetAttributeValue(XName.Get("val", DocX.w.NamespaceName), value.ToString());
			}
		}

		internal Cell(Row row, DocX document, XElement xml)
			: base(document, xml)
		{
			this.row = row;
			mainPart = row.mainPart;
		}

		public void SetBorder(TableCellBorderType borderType, Border border)
		{
			XElement xElement = base.Xml.Element(XName.Get("tcPr", DocX.w.NamespaceName));
			if (xElement == null)
			{
				base.Xml.SetElementValue(XName.Get("tcPr", DocX.w.NamespaceName), string.Empty);
				xElement = base.Xml.Element(XName.Get("tcPr", DocX.w.NamespaceName));
			}
			XElement xElement2 = xElement.Element(XName.Get("tcBorders", DocX.w.NamespaceName));
			if (xElement2 == null)
			{
				xElement.SetElementValue(XName.Get("tcBorders", DocX.w.NamespaceName), string.Empty);
				xElement2 = xElement.Element(XName.Get("tcBorders", DocX.w.NamespaceName));
			}
			string text;
			switch (borderType)
			{
			case TableCellBorderType.TopLeftToBottomRight:
				text = "tl2br";
				break;
			case TableCellBorderType.TopRightToBottomLeft:
				text = "tr2bl";
				break;
			default:
				text = borderType.ToString();
				text = text.Substring(0, 1).ToLower() + text.Substring(1);
				break;
			}
			XElement xElement3 = xElement2.Element(XName.Get(borderType.ToString(), DocX.w.NamespaceName));
			if (xElement3 == null)
			{
				xElement2.SetElementValue(XName.Get(text, DocX.w.NamespaceName), string.Empty);
				xElement3 = xElement2.Element(XName.Get(text, DocX.w.NamespaceName));
			}
			string text2 = border.Tcbs.ToString().Substring(5);
			text2 = text2.Substring(0, 1).ToLower() + text2.Substring(1);
			xElement3.SetAttributeValue(XName.Get("val", DocX.w.NamespaceName), text2);
			int num;
			switch (border.Size)
			{
			case BorderSize.one:
				num = 2;
				break;
			case BorderSize.two:
				num = 4;
				break;
			case BorderSize.three:
				num = 6;
				break;
			case BorderSize.four:
				num = 8;
				break;
			case BorderSize.five:
				num = 12;
				break;
			case BorderSize.six:
				num = 18;
				break;
			case BorderSize.seven:
				num = 24;
				break;
			case BorderSize.eight:
				num = 36;
				break;
			case BorderSize.nine:
				num = 48;
				break;
			default:
				num = 2;
				break;
			}
			xElement3.SetAttributeValue(XName.Get("sz", DocX.w.NamespaceName), num.ToString());
			xElement3.SetAttributeValue(XName.Get("space", DocX.w.NamespaceName), border.Space.ToString());
			xElement3.SetAttributeValue(XName.Get("color", DocX.w.NamespaceName), border.Color.ToHex());
		}

		public Border GetBorder(TableCellBorderType borderType)
		{
			Border border = new Border();
			XElement xElement = base.Xml.Element(XName.Get("tcPr", DocX.w.NamespaceName));
			if (xElement == null)
			{
			}
			XElement xElement2 = xElement.Element(XName.Get("tcBorders", DocX.w.NamespaceName));
			if (xElement2 == null)
			{
			}
			string text = borderType.ToString();
			string a = text;
			text = ((a == "TopLeftToBottomRight") ? "tl2br" : ((!(a == "TopRightToBottomLeft")) ? (text.Substring(0, 1).ToLower() + text.Substring(1)) : "tr2bl"));
			XElement xElement3 = xElement2.Element(XName.Get(text, DocX.w.NamespaceName));
			if (xElement3 == null)
			{
			}
			XAttribute xAttribute = xElement3.Attribute(XName.Get("val", DocX.w.NamespaceName));
			if (xAttribute != null)
			{
				try
				{
					string value = "Tcbs_" + xAttribute.Value;
					border.Tcbs = (BorderStyle)Enum.Parse(typeof(BorderStyle), value);
				}
				catch
				{
					xAttribute.Remove();
				}
			}
			XAttribute xAttribute2 = xElement3.Attribute(XName.Get("sz", DocX.w.NamespaceName));
			if (xAttribute2 != null)
			{
				if (int.TryParse(xAttribute2.Value, out int result))
				{
					switch (result)
					{
					case 2:
						border.Size = BorderSize.one;
						break;
					case 4:
						border.Size = BorderSize.two;
						break;
					case 6:
						border.Size = BorderSize.three;
						break;
					case 8:
						border.Size = BorderSize.four;
						break;
					case 12:
						border.Size = BorderSize.five;
						break;
					case 18:
						border.Size = BorderSize.six;
						break;
					case 24:
						border.Size = BorderSize.seven;
						break;
					case 36:
						border.Size = BorderSize.eight;
						break;
					case 48:
						border.Size = BorderSize.nine;
						break;
					default:
						border.Size = BorderSize.one;
						break;
					}
				}
				else
				{
					xAttribute2.Remove();
				}
			}
			XAttribute xAttribute3 = xElement3.Attribute(XName.Get("space", DocX.w.NamespaceName));
			if (xAttribute3 != null)
			{
				if (!int.TryParse(xAttribute3.Value, out int result2))
				{
					xAttribute3.Remove();
				}
				else
				{
					border.Space = result2;
				}
			}
			XAttribute xAttribute4 = xElement3.Attribute(XName.Get("color", DocX.w.NamespaceName));
			if (xAttribute4 != null)
			{
				try
				{
					border.Color = ColorTranslator.FromHtml($"#{xAttribute4.Value}");
				}
				catch
				{
					xAttribute4.Remove();
				}
			}
			return border;
		}

		public override Table InsertTable(int rowCount, int columnCount)
		{
			Table table = base.InsertTable(rowCount, columnCount);
			table.mainPart = mainPart;
			InsertParagraph();
			return table;
		}
	}
}
