using System.Xml.Linq;

namespace Novacode
{
	public class LineChart : Chart
	{
		public Grouping Grouping
		{
			get
			{
				return XElementHelpers.GetValueToEnum<Grouping>(base.ChartXml.Element(XName.Get("grouping", DocX.c.NamespaceName)));
			}
			set
			{
				XElementHelpers.SetValueFromEnum(base.ChartXml.Element(XName.Get("grouping", DocX.c.NamespaceName)), value);
			}
		}

		protected override XElement CreateChartXml()
		{
			return XElement.Parse("<c:lineChart xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\">\r\n                    <c:grouping val=\"standard\"/>                    \r\n                  </c:lineChart>");
		}
	}
}
