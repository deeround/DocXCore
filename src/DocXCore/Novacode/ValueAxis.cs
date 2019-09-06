using System.Xml.Linq;

namespace Novacode
{
	public class ValueAxis : Axis
	{
		internal ValueAxis(XElement xml)
			: base(xml)
		{
		}

		public ValueAxis(string id)
			: base(id)
		{
			base.Xml = XElement.Parse($"<c:valAx xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\">\r\n                <c:axId val=\"{id}\"/>\r\n                <c:scaling>\r\n                  <c:orientation val=\"minMax\"/>\r\n                </c:scaling>\r\n                <c:delete val=\"0\"/>\r\n                <c:axPos val=\"l\"/>\r\n                <c:numFmt sourceLinked=\"0\" formatCode=\"General\"/>\r\n                <c:majorGridlines/>\r\n                <c:majorTickMark val=\"out\"/>\r\n                <c:minorTickMark val=\"none\"/>\r\n                <c:tickLblPos val=\"nextTo\"/>\r\n                <c:crossAx val=\"148921728\"/>\r\n                <c:crosses val=\"autoZero\"/>\r\n                <c:crossBetween val=\"between\"/>\r\n              </c:valAx>");
		}
	}
}
