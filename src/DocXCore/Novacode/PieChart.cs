using System.Xml.Linq;

namespace Novacode
{
	public class PieChart : Chart
	{
		public override bool IsAxisExist => false;

		public override short MaxSeriesCount => 1;

		protected override XElement CreateChartXml()
		{
			return XElement.Parse("<c:pieChart xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\">\r\n                  </c:pieChart>");
		}
	}
}
