using System.Xml.Linq;

namespace Novacode
{
	public class CategoryAxis : Axis
	{
		internal CategoryAxis(XElement xml)
			: base(xml)
		{
		}

		public CategoryAxis(string id)
			: base(id)
		{
			base.Xml = XElement.Parse($"<c:catAx xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\"> \r\n                <c:axId val=\"{id}\"/>\r\n                <c:scaling>\r\n                  <c:orientation val=\"minMax\"/>\r\n                </c:scaling>\r\n                <c:delete val=\"0\"/>\r\n                <c:axPos val=\"b\"/>\r\n                <c:majorTickMark val=\"out\"/>\r\n                <c:minorTickMark val=\"none\"/>\r\n                <c:tickLblPos val=\"nextTo\"/>\r\n                <c:crossAx val=\"154227840\"/>\r\n                <c:crosses val=\"autoZero\"/>\r\n                <c:auto val=\"1\"/>\r\n                <c:lblAlgn val=\"ctr\"/>\r\n                <c:lblOffset val=\"100\"/>\r\n                <c:noMultiLvlLbl val=\"0\"/>\r\n              </c:catAx>");
		}
	}
}
