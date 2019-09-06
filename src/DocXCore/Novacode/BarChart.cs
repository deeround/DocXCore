using System;
using System.Xml.Linq;

namespace Novacode
{
	public class BarChart : Chart
	{
		public BarDirection BarDirection
		{
			get
			{
				return XElementHelpers.GetValueToEnum<BarDirection>(base.ChartXml.Element(XName.Get("barDir", DocX.c.NamespaceName)));
			}
			set
			{
				XElementHelpers.SetValueFromEnum(base.ChartXml.Element(XName.Get("barDir", DocX.c.NamespaceName)), value);
			}
		}

		public BarGrouping BarGrouping
		{
			get
			{
				return XElementHelpers.GetValueToEnum<BarGrouping>(base.ChartXml.Element(XName.Get("grouping", DocX.c.NamespaceName)));
			}
			set
			{
				XElementHelpers.SetValueFromEnum(base.ChartXml.Element(XName.Get("grouping", DocX.c.NamespaceName)), value);
			}
		}

		public int GapWidth
		{
			get
			{
				return Convert.ToInt32(base.ChartXml.Element(XName.Get("gapWidth", DocX.c.NamespaceName)).Attribute(XName.Get("val")).Value);
			}
			set
			{
				if (value < 1 || value > 500)
				{
					throw new ArgumentException("GapWidth lay between 0% and 500%!");
				}
				base.ChartXml.Element(XName.Get("gapWidth", DocX.c.NamespaceName)).Attribute(XName.Get("val")).Value = value.ToString();
			}
		}

		protected override XElement CreateChartXml()
		{
			return XElement.Parse("<c:barChart xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\">\r\n                    <c:barDir val=\"col\"/>\r\n                    <c:grouping val=\"clustered\"/>                    \r\n                    <c:gapWidth val=\"150\"/>\r\n                  </c:barChart>");
		}
	}
}
