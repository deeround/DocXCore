using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Xml.Linq;

namespace Novacode
{
	public class Table : InsertBeforeOrAfter
	{
		private Alignment alignment;

		private AutoFit autofit;

		private float[] ColumnWidthsValue;

		private int _cachedColCount = -1;

		private TableDesign design;

		private string _customTableDesignName;

		private string _tableCaption;

		private string _tableDescription;

		public virtual List<Paragraph> Paragraphs
		{
			get
			{
				List<Paragraph> list = new List<Paragraph>();
				foreach (Row row in Rows)
				{
					list.AddRange(row.Paragraphs);
				}
				return list;
			}
		}

		public List<Picture> Pictures
		{
			get
			{
				List<Picture> list = new List<Picture>();
				foreach (Row row in Rows)
				{
					list.AddRange(row.Pictures);
				}
				return list;
			}
		}

		public List<Hyperlink> Hyperlinks
		{
			get
			{
				List<Hyperlink> list = new List<Hyperlink>();
				foreach (Row row in Rows)
				{
					list.AddRange(row.Hyperlinks);
				}
				return list;
			}
		}

		public List<double> ColumnWidths
		{
			get
			{
				List<double> list = new List<double>();
				IEnumerable<XElement> enumerable = base.Xml.Element(XName.Get("tblGrid", DocX.w.NamespaceName))?.Elements(XName.Get("gridCol", DocX.w.NamespaceName));
				if (enumerable == null)
				{
					return null;
				}
				foreach (XElement item in enumerable)
				{
					string attribute = item.GetAttribute(XName.Get("w", DocX.w.NamespaceName));
					list.Add(Convert.ToDouble(attribute));
				}
				return list;
			}
		}

		public int RowCount => base.Xml.Elements(XName.Get("tr", DocX.w.NamespaceName)).Count();

		public int ColumnCount
		{
			get
			{
				if (RowCount == 0)
				{
					return 0;
				}
				if (_cachedColCount == -1)
				{
					_cachedColCount = Rows.First().ColumnCount;
				}
				return _cachedColCount;
			}
		}

		public List<Row> Rows => (from r in base.Xml.Elements(XName.Get("tr", DocX.w.NamespaceName))
		select new Row(this, base.Document, r)).ToList();

		public string CustomTableDesignName
		{
			get
			{
				return _customTableDesignName;
			}
			set
			{
				_customTableDesignName = value;
				Design = TableDesign.Custom;
			}
		}

		public string TableCaption
		{
			get
			{
				XElement xElement = base.Xml.Element(XName.Get("tblPr", DocX.w.NamespaceName))?.Element(XName.Get("tblCaption", DocX.w.NamespaceName));
				if (xElement != null)
				{
					_tableCaption = xElement.GetAttribute(XName.Get("val", DocX.w.NamespaceName));
				}
				return _tableCaption;
			}
			set
			{
				XElement xElement = base.Xml.Element(XName.Get("tblPr", DocX.w.NamespaceName));
				if (xElement != null)
				{
					xElement.Descendants(XName.Get("tblCaption", DocX.w.NamespaceName)).FirstOrDefault()?.Remove();
					XElement content = new XElement(XName.Get("tblCaption", DocX.w.NamespaceName), new XAttribute(XName.Get("val", DocX.w.NamespaceName), value));
					xElement.Add(content);
				}
			}
		}

		public string TableDescription
		{
			get
			{
				XElement xElement = base.Xml.Element(XName.Get("tblPr", DocX.w.NamespaceName))?.Element(XName.Get("tblDescription", DocX.w.NamespaceName));
				if (xElement != null)
				{
					_tableDescription = xElement.GetAttribute(XName.Get("val", DocX.w.NamespaceName));
				}
				return _tableDescription;
			}
			set
			{
				XElement xElement = base.Xml.Element(XName.Get("tblPr", DocX.w.NamespaceName));
				if (xElement != null)
				{
					xElement.Descendants(XName.Get("tblDescription", DocX.w.NamespaceName)).FirstOrDefault()?.Remove();
					XElement content = new XElement(XName.Get("tblDescription", DocX.w.NamespaceName), new XAttribute(XName.Get("val", DocX.w.NamespaceName), value));
					xElement.Add(content);
				}
			}
		}

		public TableLook TableLook
		{
			get;
			set;
		}

		public Alignment Alignment
		{
			get
			{
				return alignment;
			}
			set
			{
				string value2 = string.Empty;
				switch (value)
				{
				case Alignment.left:
					value2 = "left";
					break;
				case Alignment.both:
					value2 = "both";
					break;
				case Alignment.right:
					value2 = "right";
					break;
				case Alignment.center:
					value2 = "center";
					break;
				}
				XElement xElement = base.Xml.Descendants(XName.Get("tblPr", DocX.w.NamespaceName)).First();
				xElement.Descendants(XName.Get("jc", DocX.w.NamespaceName)).FirstOrDefault()?.Remove();
				XElement content = new XElement(XName.Get("jc", DocX.w.NamespaceName), new XAttribute(XName.Get("val", DocX.w.NamespaceName), value2));
				xElement.Add(content);
				alignment = value;
			}
		}

		public AutoFit AutoFit
		{
			get
			{
				return autofit;
			}
			set
			{
				string value2 = string.Empty;
				string value3 = string.Empty;
				switch (value)
				{
				case AutoFit.ColumnWidth:
				{
					value2 = "auto";
					value3 = "dxa";
					XElement xElement5 = base.Xml.Element(XName.Get("tblPr", DocX.w.NamespaceName));
					if (xElement5 != null)
					{
						XElement xElement6 = xElement5.Element(XName.Get("tblLayout", DocX.w.NamespaceName));
						if (xElement6 == null)
						{
							xElement5.Add(new XElement(XName.Get("tblLayout", DocX.w.NamespaceName)));
							xElement6 = xElement5.Element(XName.Get("tblLayout", DocX.w.NamespaceName));
						}
						XAttribute xAttribute = xElement6.Attribute(XName.Get("type", DocX.w.NamespaceName));
						if (xAttribute == null)
						{
							xElement6.Add(new XAttribute(XName.Get("type", DocX.w.NamespaceName), string.Empty));
							xAttribute = xElement6.Attribute(XName.Get("type", DocX.w.NamespaceName));
						}
						xAttribute.Value = "fixed";
					}
					break;
				}
				case AutoFit.Contents:
					value2 = (value3 = "auto");
					break;
				case AutoFit.Window:
					value2 = (value3 = "pct");
					break;
				case AutoFit.Fixed:
				{
					value2 = (value3 = "dxa");
					XElement xElement = base.Xml.Element(XName.Get("tblPr", DocX.w.NamespaceName));
					XElement xElement2 = xElement.Element(XName.Get("tblLayout", DocX.w.NamespaceName));
					if (xElement2 == null)
					{
						XElement xElement3 = xElement.Element(XName.Get("tblInd", DocX.w.NamespaceName)) ?? xElement.Element(XName.Get("tblW", DocX.w.NamespaceName));
						xElement3.AddAfterSelf(new XElement(XName.Get("tblLayout", DocX.w.NamespaceName)));
						xElement3 = xElement.Element(XName.Get("tblLayout", DocX.w.NamespaceName));
						xElement3.SetAttributeValue(XName.Get("type", DocX.w.NamespaceName), "fixed");
						xElement3 = xElement.Element(XName.Get("tblW", DocX.w.NamespaceName));
						double num = 0.0;
						foreach (double columnWidth in ColumnWidths)
						{
							num += columnWidth;
						}
						xElement3.SetAttributeValue(XName.Get("w", DocX.w.NamespaceName), num.ToString());
					}
					else
					{
						IEnumerable<XAttribute> enumerable = from d in base.Xml.Descendants()
						let type = d.Attribute(XName.Get("type", DocX.w.NamespaceName))
						where d.Name.LocalName == "tblLayout" && type != null
						select type;
						foreach (XAttribute item in enumerable)
						{
							item.Value = "fixed";
						}
						XElement xElement4 = xElement.Element(XName.Get("tblW", DocX.w.NamespaceName));
						double num2 = 0.0;
						foreach (double columnWidth2 in ColumnWidths)
						{
							num2 += columnWidth2;
						}
						xElement4.SetAttributeValue(XName.Get("w", DocX.w.NamespaceName), num2.ToString());
					}
					break;
				}
				}
				IEnumerable<XAttribute> enumerable2 = from d in base.Xml.Descendants()
				let type = d.Attribute(XName.Get("type", DocX.w.NamespaceName))
				where d.Name.LocalName == "tblW" && type != null
				select type;
				foreach (XAttribute item2 in enumerable2)
				{
					item2.Value = value2;
				}
				enumerable2 = from d in base.Xml.Descendants()
				let type = d.Attribute(XName.Get("type", DocX.w.NamespaceName))
				where d.Name.LocalName == "tcW" && type != null
				select type;
				foreach (XAttribute item3 in enumerable2)
				{
					item3.Value = value3;
				}
				autofit = value;
			}
		}

		public TableDesign Design
		{
			get
			{
				return design;
			}
			set
			{
				XElement xElement = base.Xml.Element(XName.Get("tblPr", DocX.w.NamespaceName));
				XElement xElement2 = xElement.Element(XName.Get("tblStyle", DocX.w.NamespaceName));
				if (xElement2 == null)
				{
					xElement.Add(new XElement(XName.Get("tblStyle", DocX.w.NamespaceName)));
					xElement2 = xElement.Element(XName.Get("tblStyle", DocX.w.NamespaceName));
				}
				XAttribute val = xElement2.Attribute(XName.Get("val", DocX.w.NamespaceName));
				if (val == null)
				{
					xElement2.Add(new XAttribute(XName.Get("val", DocX.w.NamespaceName), ""));
					val = xElement2.Attribute(XName.Get("val", DocX.w.NamespaceName));
				}
				design = value;
				if (design == TableDesign.None)
				{
					xElement2.Remove();
				}
				if (design != 0)
				{
					switch (design)
					{
					case TableDesign.TableNormal:
						val.Value = "TableNormal";
						break;
					case TableDesign.TableGrid:
						val.Value = "TableGrid";
						break;
					case TableDesign.LightShading:
						val.Value = "LightShading";
						break;
					case TableDesign.LightShadingAccent1:
						val.Value = "LightShading-Accent1";
						break;
					case TableDesign.LightShadingAccent2:
						val.Value = "LightShading-Accent2";
						break;
					case TableDesign.LightShadingAccent3:
						val.Value = "LightShading-Accent3";
						break;
					case TableDesign.LightShadingAccent4:
						val.Value = "LightShading-Accent4";
						break;
					case TableDesign.LightShadingAccent5:
						val.Value = "LightShading-Accent5";
						break;
					case TableDesign.LightShadingAccent6:
						val.Value = "LightShading-Accent6";
						break;
					case TableDesign.LightList:
						val.Value = "LightList";
						break;
					case TableDesign.LightListAccent1:
						val.Value = "LightList-Accent1";
						break;
					case TableDesign.LightListAccent2:
						val.Value = "LightList-Accent2";
						break;
					case TableDesign.LightListAccent3:
						val.Value = "LightList-Accent3";
						break;
					case TableDesign.LightListAccent4:
						val.Value = "LightList-Accent4";
						break;
					case TableDesign.LightListAccent5:
						val.Value = "LightList-Accent5";
						break;
					case TableDesign.LightListAccent6:
						val.Value = "LightList-Accent6";
						break;
					case TableDesign.LightGrid:
						val.Value = "LightGrid";
						break;
					case TableDesign.LightGridAccent1:
						val.Value = "LightGrid-Accent1";
						break;
					case TableDesign.LightGridAccent2:
						val.Value = "LightGrid-Accent2";
						break;
					case TableDesign.LightGridAccent3:
						val.Value = "LightGrid-Accent3";
						break;
					case TableDesign.LightGridAccent4:
						val.Value = "LightGrid-Accent4";
						break;
					case TableDesign.LightGridAccent5:
						val.Value = "LightGrid-Accent5";
						break;
					case TableDesign.LightGridAccent6:
						val.Value = "LightGrid-Accent6";
						break;
					case TableDesign.MediumShading1:
						val.Value = "MediumShading1";
						break;
					case TableDesign.MediumShading1Accent1:
						val.Value = "MediumShading1-Accent1";
						break;
					case TableDesign.MediumShading1Accent2:
						val.Value = "MediumShading1-Accent2";
						break;
					case TableDesign.MediumShading1Accent3:
						val.Value = "MediumShading1-Accent3";
						break;
					case TableDesign.MediumShading1Accent4:
						val.Value = "MediumShading1-Accent4";
						break;
					case TableDesign.MediumShading1Accent5:
						val.Value = "MediumShading1-Accent5";
						break;
					case TableDesign.MediumShading1Accent6:
						val.Value = "MediumShading1-Accent6";
						break;
					case TableDesign.MediumShading2:
						val.Value = "MediumShading2";
						break;
					case TableDesign.MediumShading2Accent1:
						val.Value = "MediumShading2-Accent1";
						break;
					case TableDesign.MediumShading2Accent2:
						val.Value = "MediumShading2-Accent2";
						break;
					case TableDesign.MediumShading2Accent3:
						val.Value = "MediumShading2-Accent3";
						break;
					case TableDesign.MediumShading2Accent4:
						val.Value = "MediumShading2-Accent4";
						break;
					case TableDesign.MediumShading2Accent5:
						val.Value = "MediumShading2-Accent5";
						break;
					case TableDesign.MediumShading2Accent6:
						val.Value = "MediumShading2-Accent6";
						break;
					case TableDesign.MediumList1:
						val.Value = "MediumList1";
						break;
					case TableDesign.MediumList1Accent1:
						val.Value = "MediumList1-Accent1";
						break;
					case TableDesign.MediumList1Accent2:
						val.Value = "MediumList1-Accent2";
						break;
					case TableDesign.MediumList1Accent3:
						val.Value = "MediumList1-Accent3";
						break;
					case TableDesign.MediumList1Accent4:
						val.Value = "MediumList1-Accent4";
						break;
					case TableDesign.MediumList1Accent5:
						val.Value = "MediumList1-Accent5";
						break;
					case TableDesign.MediumList1Accent6:
						val.Value = "MediumList1-Accent6";
						break;
					case TableDesign.MediumList2:
						val.Value = "MediumList2";
						break;
					case TableDesign.MediumList2Accent1:
						val.Value = "MediumList2-Accent1";
						break;
					case TableDesign.MediumList2Accent2:
						val.Value = "MediumList2-Accent2";
						break;
					case TableDesign.MediumList2Accent3:
						val.Value = "MediumList2-Accent3";
						break;
					case TableDesign.MediumList2Accent4:
						val.Value = "MediumList2-Accent4";
						break;
					case TableDesign.MediumList2Accent5:
						val.Value = "MediumList2-Accent5";
						break;
					case TableDesign.MediumList2Accent6:
						val.Value = "MediumList2-Accent6";
						break;
					case TableDesign.MediumGrid1:
						val.Value = "MediumGrid1";
						break;
					case TableDesign.MediumGrid1Accent1:
						val.Value = "MediumGrid1-Accent1";
						break;
					case TableDesign.MediumGrid1Accent2:
						val.Value = "MediumGrid1-Accent2";
						break;
					case TableDesign.MediumGrid1Accent3:
						val.Value = "MediumGrid1-Accent3";
						break;
					case TableDesign.MediumGrid1Accent4:
						val.Value = "MediumGrid1-Accent4";
						break;
					case TableDesign.MediumGrid1Accent5:
						val.Value = "MediumGrid1-Accent5";
						break;
					case TableDesign.MediumGrid1Accent6:
						val.Value = "MediumGrid1-Accent6";
						break;
					case TableDesign.MediumGrid2:
						val.Value = "MediumGrid2";
						break;
					case TableDesign.MediumGrid2Accent1:
						val.Value = "MediumGrid2-Accent1";
						break;
					case TableDesign.MediumGrid2Accent2:
						val.Value = "MediumGrid2-Accent2";
						break;
					case TableDesign.MediumGrid2Accent3:
						val.Value = "MediumGrid2-Accent3";
						break;
					case TableDesign.MediumGrid2Accent4:
						val.Value = "MediumGrid2-Accent4";
						break;
					case TableDesign.MediumGrid2Accent5:
						val.Value = "MediumGrid2-Accent5";
						break;
					case TableDesign.MediumGrid2Accent6:
						val.Value = "MediumGrid2-Accent6";
						break;
					case TableDesign.MediumGrid3:
						val.Value = "MediumGrid3";
						break;
					case TableDesign.MediumGrid3Accent1:
						val.Value = "MediumGrid3-Accent1";
						break;
					case TableDesign.MediumGrid3Accent2:
						val.Value = "MediumGrid3-Accent2";
						break;
					case TableDesign.MediumGrid3Accent3:
						val.Value = "MediumGrid3-Accent3";
						break;
					case TableDesign.MediumGrid3Accent4:
						val.Value = "MediumGrid3-Accent4";
						break;
					case TableDesign.MediumGrid3Accent5:
						val.Value = "MediumGrid3-Accent5";
						break;
					case TableDesign.MediumGrid3Accent6:
						val.Value = "MediumGrid3-Accent6";
						break;
					case TableDesign.DarkList:
						val.Value = "DarkList";
						break;
					case TableDesign.DarkListAccent1:
						val.Value = "DarkList-Accent1";
						break;
					case TableDesign.DarkListAccent2:
						val.Value = "DarkList-Accent2";
						break;
					case TableDesign.DarkListAccent3:
						val.Value = "DarkList-Accent3";
						break;
					case TableDesign.DarkListAccent4:
						val.Value = "DarkList-Accent4";
						break;
					case TableDesign.DarkListAccent5:
						val.Value = "DarkList-Accent5";
						break;
					case TableDesign.DarkListAccent6:
						val.Value = "DarkList-Accent6";
						break;
					case TableDesign.ColorfulShading:
						val.Value = "ColorfulShading";
						break;
					case TableDesign.ColorfulShadingAccent1:
						val.Value = "ColorfulShading-Accent1";
						break;
					case TableDesign.ColorfulShadingAccent2:
						val.Value = "ColorfulShading-Accent2";
						break;
					case TableDesign.ColorfulShadingAccent3:
						val.Value = "ColorfulShading-Accent3";
						break;
					case TableDesign.ColorfulShadingAccent4:
						val.Value = "ColorfulShading-Accent4";
						break;
					case TableDesign.ColorfulShadingAccent5:
						val.Value = "ColorfulShading-Accent5";
						break;
					case TableDesign.ColorfulShadingAccent6:
						val.Value = "ColorfulShading-Accent6";
						break;
					case TableDesign.ColorfulList:
						val.Value = "ColorfulList";
						break;
					case TableDesign.ColorfulListAccent1:
						val.Value = "ColorfulList-Accent1";
						break;
					case TableDesign.ColorfulListAccent2:
						val.Value = "ColorfulList-Accent2";
						break;
					case TableDesign.ColorfulListAccent3:
						val.Value = "ColorfulList-Accent3";
						break;
					case TableDesign.ColorfulListAccent4:
						val.Value = "ColorfulList-Accent4";
						break;
					case TableDesign.ColorfulListAccent5:
						val.Value = "ColorfulList-Accent5";
						break;
					case TableDesign.ColorfulListAccent6:
						val.Value = "ColorfulList-Accent6";
						break;
					case TableDesign.ColorfulGrid:
						val.Value = "ColorfulGrid";
						break;
					case TableDesign.ColorfulGridAccent1:
						val.Value = "ColorfulGrid-Accent1";
						break;
					case TableDesign.ColorfulGridAccent2:
						val.Value = "ColorfulGrid-Accent2";
						break;
					case TableDesign.ColorfulGridAccent3:
						val.Value = "ColorfulGrid-Accent3";
						break;
					case TableDesign.ColorfulGridAccent4:
						val.Value = "ColorfulGrid-Accent4";
						break;
					case TableDesign.ColorfulGridAccent5:
						val.Value = "ColorfulGrid-Accent5";
						break;
					case TableDesign.ColorfulGridAccent6:
						val.Value = "ColorfulGrid-Accent6";
						break;
					}
				}
				else if (string.IsNullOrEmpty(_customTableDesignName))
				{
					design = TableDesign.None;
					xElement2?.Remove();
				}
				else
				{
					val.Value = _customTableDesignName;
				}
				if (base.Document.styles == null)
				{
					PackagePart part = base.Document.package.GetPart(new Uri("/word/styles.xml", UriKind.Relative));
					using (TextReader textReader = new StreamReader(part.GetStream()))
					{
						base.Document.styles = XDocument.Load(textReader);
					}
				}
				XElement xElement3 = (from e in base.Document.styles.Descendants()
				let styleId = e.Attribute(XName.Get("styleId", DocX.w.NamespaceName))
				where styleId != null && styleId.Value == val.Value
				select e).FirstOrDefault();
				if (xElement3 == null)
				{
					XDocument xDocument = HelperFunctions.DecompressXMLResource("Novacode.Resources.styles.xml.gz");
					XElement content = (from e in xDocument.Descendants()
					let styleId = e.Attribute(XName.Get("styleId", DocX.w.NamespaceName))
					where styleId != null && styleId.Value == val.Value
					select e).FirstOrDefault();
					base.Document.styles.Element(XName.Get("styles", DocX.w.NamespaceName)).Add(content);
				}
			}
		}

		public int Index
		{
			get
			{
				int num = 0;
				IEnumerable<XElement> enumerable = base.Xml.ElementsBeforeSelf();
				foreach (XElement item in enumerable)
				{
					num += Paragraph.GetElementTextLength(item);
				}
				return num;
			}
		}

		public void MergeCellsInColumn(int columnIndex, int startRow, int endRow)
		{
			if (columnIndex < 0 || columnIndex >= ColumnCount)
			{
				throw new IndexOutOfRangeException();
			}
			if (startRow < 0 || endRow <= startRow || endRow >= Rows.Count)
			{
				throw new IndexOutOfRangeException();
			}
			foreach (Row item in Rows.Where((Row z, int i) => i > startRow && i <= endRow))
			{
				Cell cell = item.Cells[columnIndex];
				XElement xElement = cell.Xml.Element(XName.Get("tcPr", DocX.w.NamespaceName));
				if (xElement == null)
				{
					cell.Xml.SetElementValue(XName.Get("tcPr", DocX.w.NamespaceName), string.Empty);
					xElement = cell.Xml.Element(XName.Get("tcPr", DocX.w.NamespaceName));
				}
				XElement xElement2 = xElement.Element(XName.Get("vMerge", DocX.w.NamespaceName));
				if (xElement2 == null)
				{
					xElement.SetElementValue(XName.Get("vMerge", DocX.w.NamespaceName), string.Empty);
					xElement2 = xElement.Element(XName.Get("vMerge", DocX.w.NamespaceName));
				}
			}
			XElement xElement3 = (columnIndex <= Rows[startRow].Cells.Count) ? Rows[startRow].Cells[columnIndex].Xml.Element(XName.Get("tcPr", DocX.w.NamespaceName)) : Rows[startRow].Cells[Rows[startRow].Cells.Count - 1].Xml.Element(XName.Get("tcPr", DocX.w.NamespaceName));
			if (xElement3 == null)
			{
				Rows[startRow].Cells[columnIndex].Xml.SetElementValue(XName.Get("tcPr", DocX.w.NamespaceName), string.Empty);
				xElement3 = Rows[startRow].Cells[columnIndex].Xml.Element(XName.Get("tcPr", DocX.w.NamespaceName));
			}
			XElement xElement4 = xElement3.Element(XName.Get("vMerge", DocX.w.NamespaceName));
			if (xElement4 == null)
			{
				xElement3.SetElementValue(XName.Get("vMerge", DocX.w.NamespaceName), string.Empty);
				xElement4 = xElement3.Element(XName.Get("vMerge", DocX.w.NamespaceName));
			}
			xElement4.SetAttributeValue(XName.Get("val", DocX.w.NamespaceName), "restart");
		}

		public void SetDirection(Direction direction)
		{
			XElement orCreate_tblPr = GetOrCreate_tblPr();
			orCreate_tblPr.Add(new XElement(DocX.w + "bidiVisual"));
			foreach (Row row in Rows)
			{
				row.SetDirection(direction);
			}
		}

		public void SetWidths(float[] widths)
		{
			ColumnWidthsValue = widths;
			foreach (Row row in Rows)
			{
				for (int i = 0; i < widths.Length; i++)
				{
					if (row.Cells.Count > i)
					{
						row.Cells[i].Width = (double)widths[i];
					}
				}
			}
		}

		public void SetWidthsPercentage(float[] widthsPercentage, float? totalWidth)
		{
			if (!totalWidth.HasValue)
			{
				totalWidth = base.Document.PageWidth - base.Document.MarginLeft - base.Document.MarginRight;
			}
			List<float> widths = new List<float>(widthsPercentage.Length);
			widthsPercentage.ToList().ForEach(delegate(float pWidth)
			{
				widths.Add(pWidth * totalWidth.Value / 100f * 1f);
			});
			SetWidths(widths.ToArray());
		}

		internal XElement GetOrCreate_tblPr()
		{
			XElement xElement = base.Xml.Element(XName.Get("tblPr", DocX.w.NamespaceName));
			if (xElement == null)
			{
				base.Xml.AddFirst(new XElement(XName.Get("tblPr", DocX.w.NamespaceName)));
				xElement = base.Xml.Element(XName.Get("tblPr", DocX.w.NamespaceName));
			}
			return xElement;
		}

		public void SetTableCellMargin(TableCellMarginType type, double margin)
		{
			XElement orCreate_tblPr = GetOrCreate_tblPr();
			XElement xElement = orCreate_tblPr.Element(XName.Get("tblCellMar", DocX.w.NamespaceName));
			if (xElement == null)
			{
				orCreate_tblPr.AddFirst(new XElement(XName.Get("tblCellMar", DocX.w.NamespaceName)));
				xElement = orCreate_tblPr.Element(XName.Get("tblCellMar", DocX.w.NamespaceName));
			}
			XElement xElement2 = xElement.Element(XName.Get(type.ToString(), DocX.w.NamespaceName));
			if (xElement2 == null)
			{
				xElement.AddFirst(new XElement(XName.Get(type.ToString(), DocX.w.NamespaceName)));
				xElement2 = xElement.Element(XName.Get(type.ToString(), DocX.w.NamespaceName));
			}
			xElement2.RemoveAttributes();
			xElement2.Add(new XAttribute(XName.Get("w", DocX.w.NamespaceName), margin));
			xElement2.Add(new XAttribute(XName.Get("type", DocX.w.NamespaceName), "dxa"));
		}

		public double GetColumnWidth(int index)
		{
			List<double> columnWidths = ColumnWidths;
			if (columnWidths == null || index > columnWidths.Count - 1)
			{
				return double.NaN;
			}
			return columnWidths[index];
		}

		public void SetColumnWidth(int index, double width)
		{
			List<double> list = ColumnWidths;
			if (list == null || index > list.Count - 1)
			{
				if (Rows.Count == 0)
				{
					throw new Exception("There is at least one row required to detect the existing columns.");
				}
				list = new List<double>();
				foreach (Cell cell in Rows[Rows.Count - 1].Cells)
				{
					list.Add(cell.Width);
				}
			}
			if (index > list.Count - 1)
			{
				throw new Exception("The index is greather than the available table columns.");
			}
			XElement xElement = base.Xml.Element(XName.Get("tblGrid", DocX.w.NamespaceName));
			if (xElement == null)
			{
				XElement orCreate_tblPr = GetOrCreate_tblPr();
				orCreate_tblPr.AddAfterSelf(new XElement(XName.Get("tblGrid", DocX.w.NamespaceName)));
				xElement = base.Xml.Element(XName.Get("tblGrid", DocX.w.NamespaceName));
			}
			xElement?.RemoveAll();
			int num = 0;
			foreach (double item in list)
			{
				double num2 = item;
				if (num == index)
				{
					num2 = width;
				}
				XElement content = new XElement(XName.Get("gridCol", DocX.w.NamespaceName), new XAttribute(XName.Get("w", DocX.w.NamespaceName), num2));
				xElement?.Add(content);
				num++;
			}
			foreach (Row row in Rows)
			{
				foreach (Cell cell2 in row.Cells)
				{
					cell2.Width = -1.0;
				}
			}
			AutoFit = AutoFit.Fixed;
		}

		internal Table(DocX document, XElement xml)
			: base(document, xml)
		{
			autofit = AutoFit.ColumnWidth;
			base.Xml = xml;
			mainPart = document.mainPart;
			XElement xElement = xml.Element(XName.Get("tblPr", DocX.w.NamespaceName));
			XElement xElement2 = xElement?.Element(XName.Get("tblStyle", DocX.w.NamespaceName));
			if (xElement2 != null)
			{
				XAttribute xAttribute = xElement2.Attribute(XName.Get("val", DocX.w.NamespaceName));
				if (xAttribute != null)
				{
					string value = xAttribute.Value.Replace("-", string.Empty);
					if (Enum.IsDefined(typeof(TableDesign), value))
					{
						design = (TableDesign)Enum.Parse(typeof(TableDesign), value);
					}
					else
					{
						design = TableDesign.Custom;
					}
				}
				else
				{
					design = TableDesign.None;
				}
			}
			else
			{
				design = TableDesign.None;
			}
			XElement xElement3 = xElement?.Element(XName.Get("tblLook", DocX.w.NamespaceName));
			if (xElement3 != null)
			{
				TableLook = new TableLook
				{
					FirstRow = (xElement3.GetAttribute(XName.Get("firstRow", DocX.w.NamespaceName)) == "1"),
					LastRow = (xElement3.GetAttribute(XName.Get("lastRow", DocX.w.NamespaceName)) == "1"),
					FirstColumn = (xElement3.GetAttribute(XName.Get("firstColumn", DocX.w.NamespaceName)) == "1"),
					LastColumn = (xElement3.GetAttribute(XName.Get("lastColumn", DocX.w.NamespaceName)) == "1"),
					NoHorizontalBanding = (xElement3.GetAttribute(XName.Get("noHBand", DocX.w.NamespaceName)) == "1"),
					NoVerticalBanding = (xElement3.GetAttribute(XName.Get("noVBand", DocX.w.NamespaceName)) == "1")
				};
			}
		}

		public void Remove()
		{
			base.Xml.Remove();
		}

		public Row InsertRow()
		{
			return InsertRow(RowCount);
		}

		public Row InsertRow(Row row)
		{
			return InsertRow(row, RowCount);
		}

		public void InsertColumn()
		{
			InsertColumn(ColumnCount, direction: true);
		}

		public void RemoveRow()
		{
			RemoveRow(RowCount - 1);
		}

		public void RemoveRow(int index)
		{
			if (index < 0 || index > RowCount - 1)
			{
				throw new IndexOutOfRangeException();
			}
			Rows[index].Xml.Remove();
			if (Rows.Count == 0)
			{
				Remove();
			}
		}

		public void RemoveColumn()
		{
			RemoveColumn(ColumnCount - 1);
		}

		public void RemoveColumn(int index)
		{
			if (index < 0 || index > ColumnCount - 1)
			{
				throw new IndexOutOfRangeException();
			}
			foreach (Row row in Rows)
			{
				if (row.Cells.Count < ColumnCount)
				{
					int num = 0;
					int num2 = 0;
					int num3 = 0;
					num3 = row.gridAfter;
					foreach (Cell cell in row.Cells)
					{
						int num4 = 0;
						if (cell.GridSpan != 0)
						{
							num4 = cell.GridSpan - 1;
						}
						if (index - num3 >= num2 && index - num3 <= num2 + num4)
						{
							row.Cells[num].Xml.Remove();
							break;
						}
						num++;
						num2 += num4 + 1;
					}
				}
				else
				{
					row.Cells[index].Xml.Remove();
				}
			}
			_cachedColCount = -1;
		}

		public Row InsertRow(int index)
		{
			if (index < 0 || index > RowCount)
			{
				throw new IndexOutOfRangeException();
			}
			List<XElement> list = new List<XElement>();
			for (int i = 0; i < ColumnCount; i++)
			{
				double w = 2310.0;
				if (ColumnWidthsValue != null && ColumnWidthsValue.Length > i)
				{
					w = (double)(ColumnWidthsValue[i] * 15f);
				}
				XElement item = HelperFunctions.CreateTableCell(w);
				list.Add(item);
			}
			return InsertRow(list, index);
		}

		public Row InsertRow(Row row, int index)
		{
			if (row == null)
			{
				throw new ArgumentNullException("row");
			}
			if (index < 0 || index > RowCount)
			{
				throw new IndexOutOfRangeException();
			}
			List<XElement> content = (from element in row.Xml.Elements(XName.Get("tc", DocX.w.NamespaceName))
			select HelperFunctions.CloneElement(element)).ToList();
			return InsertRow(content, index);
		}

		private Row InsertRow(List<XElement> content, int index)
		{
			Row row = new Row(this, base.Document, new XElement(XName.Get("tr", DocX.w.NamespaceName), content));
			if (index == Rows.Count)
			{
				XElement xml = Rows.Last().Xml;
				xml.AddAfterSelf(row.Xml);
			}
			else
			{
				XElement xml = Rows[index].Xml;
				xml.AddBeforeSelf(row.Xml);
			}
			return row;
		}

		public void InsertColumn(int index, bool direction)
		{
			int columnCount = ColumnCount;
			if (RowCount > 0)
			{
				if (index <= 0 || index > columnCount)
				{
					throw new IndexOutOfRangeException("Out of index bounds, column count is " + columnCount + " you input " + index);
				}
				_cachedColCount = -1;
				foreach (Row row in Rows)
				{
					XElement cell = HelperFunctions.CreateTableCell();
					if (row.Cells.Count < columnCount)
					{
						if (index >= columnCount)
						{
							AddCellToRow(row, cell, row.Cells.Count, direction);
						}
						else
						{
							int num = 1;
							int num2 = 1;
							int num3 = 0;
							num3 = row.gridAfter;
							foreach (Cell cell2 in row.Cells)
							{
								int num4 = 0;
								if (cell2.GridSpan != 0)
								{
									num4 = cell2.GridSpan - 1;
								}
								if (index - num3 >= num2 && index - num3 <= num2 + num4)
								{
									bool direction2 = (index == num2 + num4 && direction) ? true : false;
									AddCellToRow(row, cell, num, direction2);
									break;
								}
								num++;
								num2 += num4 + 1;
							}
						}
					}
					else if (row.Cells.Count == index)
					{
						AddCellToRow(row, cell, index, direction);
					}
					else
					{
						AddCellToRow(row, cell, index, direction);
					}
				}
			}
		}

		private void AddCellToRow(Row row, XElement cell, int index, bool direction)
		{
			index--;
			if (direction)
			{
				row.Cells[index].Xml.AddAfterSelf(cell);
			}
			else
			{
				row.Cells[index].Xml.AddBeforeSelf(cell);
			}
		}

		public void DeleteAndShiftCellsLeft(int rowIndex, int celIndex)
		{
			XElement xElement = Rows[rowIndex].Xml.Element(XName.Get("trPr", DocX.w.NamespaceName));
			if (xElement != null)
			{
				XElement xElement2 = xElement.Element(XName.Get("gridAfter", DocX.w.NamespaceName));
				if (xElement2 != null)
				{
					XAttribute xAttribute = xElement2.Attribute(XName.Get("val", DocX.w.NamespaceName));
					xAttribute.Value = (int.Parse(xAttribute.Value) + 1).ToString();
				}
				else
				{
					xElement2.SetAttributeValue("val", 1);
				}
			}
			else
			{
				XElement xElement3 = new XElement(XName.Get("trPr", DocX.w.NamespaceName));
				XElement xElement4 = new XElement(XName.Get("gridAfter", DocX.w.NamespaceName));
				XAttribute content = new XAttribute(XName.Get("val", DocX.w.NamespaceName), 1);
				xElement4.Add(content);
				xElement3.Add(xElement4);
				Rows[rowIndex].Xml.AddFirst(xElement3);
			}
			int columnCount = ColumnCount;
			if (celIndex <= ColumnCount && Rows[rowIndex].ColumnCount <= ColumnCount)
			{
				Rows[rowIndex].Cells[celIndex].Xml.Remove();
			}
		}

		public override void InsertPageBreakBeforeSelf()
		{
			base.InsertPageBreakBeforeSelf();
		}

		public override void InsertPageBreakAfterSelf()
		{
			base.InsertPageBreakAfterSelf();
		}

		public override Table InsertTableBeforeSelf(Table t)
		{
			return base.InsertTableBeforeSelf(t);
		}

		public override Table InsertTableBeforeSelf(int rowCount, int columnCount)
		{
			return base.InsertTableBeforeSelf(rowCount, columnCount);
		}

		public override Table InsertTableAfterSelf(Table t)
		{
			return base.InsertTableAfterSelf(t);
		}

		public override Table InsertTableAfterSelf(int rowCount, int columnCount)
		{
			return base.InsertTableAfterSelf(rowCount, columnCount);
		}

		public override Paragraph InsertParagraphBeforeSelf(Paragraph p)
		{
			return base.InsertParagraphBeforeSelf(p);
		}

		public override Paragraph InsertParagraphBeforeSelf(string text)
		{
			return base.InsertParagraphBeforeSelf(text);
		}

		public override Paragraph InsertParagraphBeforeSelf(string text, bool trackChanges)
		{
			return base.InsertParagraphBeforeSelf(text, trackChanges);
		}

		public override Paragraph InsertParagraphBeforeSelf(string text, bool trackChanges, Formatting formatting)
		{
			return base.InsertParagraphBeforeSelf(text, trackChanges, formatting);
		}

		public override Paragraph InsertParagraphAfterSelf(Paragraph p)
		{
			return base.InsertParagraphAfterSelf(p);
		}

		public override Paragraph InsertParagraphAfterSelf(string text, bool trackChanges, Formatting formatting)
		{
			return base.InsertParagraphAfterSelf(text, trackChanges, formatting);
		}

		public override Paragraph InsertParagraphAfterSelf(string text, bool trackChanges)
		{
			return base.InsertParagraphAfterSelf(text, trackChanges);
		}

		public override Paragraph InsertParagraphAfterSelf(string text)
		{
			return base.InsertParagraphAfterSelf(text);
		}

		public void SetBorder(TableBorderType borderType, Border border)
		{
			XElement xElement = base.Xml.Element(XName.Get("tblPr", DocX.w.NamespaceName));
			if (xElement == null)
			{
				base.Xml.SetElementValue(XName.Get("tblPr", DocX.w.NamespaceName), string.Empty);
				xElement = base.Xml.Element(XName.Get("tblPr", DocX.w.NamespaceName));
			}
			XElement xElement2 = xElement.Element(XName.Get("tblBorders", DocX.w.NamespaceName));
			if (xElement2 == null)
			{
				xElement.SetElementValue(XName.Get("tblBorders", DocX.w.NamespaceName), string.Empty);
				xElement2 = xElement.Element(XName.Get("tblBorders", DocX.w.NamespaceName));
			}
			string text = borderType.ToString();
			text = text.Substring(0, 1).ToLower() + text.Substring(1);
			XElement xElement3 = xElement2.Element(XName.Get(borderType.ToString(), DocX.w.NamespaceName));
			if (xElement3 == null)
			{
				xElement2.SetElementValue(XName.Get(text, DocX.w.NamespaceName), string.Empty);
				xElement3 = xElement2.Element(XName.Get(text, DocX.w.NamespaceName));
			}
			string text2 = border.Tcbs.ToString().Substring(5);
			text2 = text2.Substring(0, 1).ToLower() + text2.Substring(1);
			xElement3.SetAttributeValue(XName.Get("val", DocX.w.NamespaceName), text2);
			if (border.Tcbs != BorderStyle.Tcbs_nil)
			{
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
		}

		public Border GetBorder(TableBorderType borderType)
		{
			Border border = new Border();
			XElement xElement = base.Xml.Element(XName.Get("tblPr", DocX.w.NamespaceName));
			if (xElement == null)
			{
			}
			XElement xElement2 = xElement.Element(XName.Get("tblBorders", DocX.w.NamespaceName));
			if (xElement2 == null)
			{
			}
			string text = borderType.ToString();
			text = text.Substring(0, 1).ToLower() + text.Substring(1);
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
	}
}
