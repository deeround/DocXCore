using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace Novacode
{
	public abstract class Chart
	{
		protected XElement ChartXml
		{
			get;
			private set;
		}

		protected XElement ChartRootXml
		{
			get;
			private set;
		}

		public XDocument Xml
		{
			get;
			private set;
		}

		public List<Series> Series
		{
			get
			{
				List<Series> list = new List<Series>();
				int num = 1;
				foreach (XElement item in ChartXml.Elements(XName.Get("ser", DocX.c.NamespaceName)))
				{
					XElement xElement = item;
					object[] obj = new object[2]
					{
						new XElement(XName.Get("idx", DocX.c.NamespaceName)),
						null
					};
					int num2 = num;
					num = num2 + 1;
					obj[1] = num2.ToString();
					xElement.Add(obj);
					list.Add(new Series(item));
				}
				return list;
			}
		}

		public virtual short MaxSeriesCount => short.MaxValue;

		public ChartLegend Legend
		{
			get;
			private set;
		}

		public CategoryAxis CategoryAxis
		{
			get;
			private set;
		}

		public ValueAxis ValueAxis
		{
			get;
			private set;
		}

		public virtual bool IsAxisExist => true;

		public bool View3D
		{
			get
			{
				return ChartXml.Name.LocalName.Contains("3D");
			}
			set
			{
				if (value)
				{
					if (!View3D)
					{
						string localName = ChartXml.Name.LocalName;
						ChartXml.Name = XName.Get(localName.Replace("Chart", "3DChart"), DocX.c.NamespaceName);
					}
				}
				else if (View3D)
				{
					string localName2 = ChartXml.Name.LocalName;
					ChartXml.Name = XName.Get(localName2.Replace("3DChart", "Chart"), DocX.c.NamespaceName);
				}
			}
		}

		public DisplayBlanksAs DisplayBlanksAs
		{
			get
			{
				return XElementHelpers.GetValueToEnum<DisplayBlanksAs>(ChartRootXml.Element(XName.Get("dispBlanksAs", DocX.c.NamespaceName)));
			}
			set
			{
				XElementHelpers.SetValueFromEnum(ChartRootXml.Element(XName.Get("dispBlanksAs", DocX.c.NamespaceName)), value);
			}
		}

		public void AddSeries(Series series)
		{
			int num = ChartXml.Elements(XName.Get("ser", DocX.c.NamespaceName)).Count();
			if (num == MaxSeriesCount)
			{
				throw new InvalidOperationException("Maximum series for this chart is" + MaxSeriesCount.ToString() + "and have exceeded!");
			}
			series.Xml.AddFirst(new XElement(XName.Get("order", DocX.c.NamespaceName), new XAttribute(XName.Get("val"), (num + 1).ToString())));
			series.Xml.AddFirst(new XElement(XName.Get("idx", DocX.c.NamespaceName), new XAttribute(XName.Get("val"), (num + 1).ToString())));
			ChartXml.Add(series.Xml);
		}

		public void AddLegend()
		{
			AddLegend(ChartLegendPosition.Right, overlay: false);
		}

		public void AddLegend(ChartLegendPosition position, bool overlay)
		{
			if (Legend != null)
			{
				RemoveLegend();
			}
			Legend = new ChartLegend(position, overlay);
			ChartRootXml.Element(XName.Get("plotArea", DocX.c.NamespaceName)).AddAfterSelf(Legend.Xml);
		}

		public void RemoveLegend()
		{
			Legend.Xml.Remove();
			Legend = null;
		}

		public Chart()
		{
			Xml = XDocument.Parse("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n                   <c:chartSpace xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">  \r\n                       <c:roundedCorners val=\"0\"/>\r\n                       <c:chart>\r\n                           <c:autoTitleDeleted val=\"0\"/>\r\n                           <c:plotVisOnly val=\"1\"/>\r\n                           <c:dispBlanksAs val=\"gap\"/>\r\n                           <c:showDLblsOverMax val=\"0\"/>\r\n                       </c:chart>\r\n                   </c:chartSpace>");
			ChartXml = CreateChartXml();
			XElement xElement = new XElement(XName.Get("plotArea", DocX.c.NamespaceName), new XElement(XName.Get("layout", DocX.c.NamespaceName)), ChartXml);
			XElement content = XElement.Parse("<c:dLbls xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\">\r\n                    <c:showLegendKey val=\"0\"/>\r\n                    <c:showVal val=\"0\"/>\r\n                    <c:showCatName val=\"0\"/>\r\n                    <c:showSerName val=\"0\"/>\r\n                    <c:showPercent val=\"0\"/>\r\n                    <c:showBubbleSize val=\"0\"/>\r\n                    <c:showLeaderLines val=\"1\"/>\r\n                </c:dLbls>");
			ChartXml.Add(content);
			if (IsAxisExist)
			{
				CategoryAxis = new CategoryAxis("148921728");
				ValueAxis = new ValueAxis("154227840");
				XElement content2 = XElement.Parse($"<c:axId val=\"{CategoryAxis.Id}\" xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\"/>");
				XElement content3 = XElement.Parse($"<c:axId val=\"{ValueAxis.Id}\" xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\"/>");
				XElement xElement2 = ChartXml.Element(XName.Get("gapWidth", DocX.c.NamespaceName));
				if (xElement2 != null)
				{
					xElement2.AddAfterSelf(content3);
					xElement2.AddAfterSelf(content2);
				}
				else
				{
					ChartXml.Add(content2);
					ChartXml.Add(content3);
				}
				xElement.Add(CategoryAxis.Xml);
				xElement.Add(ValueAxis.Xml);
			}
			ChartRootXml = Xml.Root.Element(XName.Get("chart", DocX.c.NamespaceName));
			ChartRootXml.Element(XName.Get("autoTitleDeleted", DocX.c.NamespaceName)).AddAfterSelf(xElement);
		}

		protected abstract XElement CreateChartXml();
	}
}
