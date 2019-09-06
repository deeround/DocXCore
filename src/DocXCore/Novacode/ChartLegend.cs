using System.Xml.Linq;

namespace Novacode
{
	public class ChartLegend
	{
		internal XElement Xml
		{
			get;
			private set;
		}

		public bool Overlay
		{
			get
			{
				return Xml.Element(XName.Get("overlay", DocX.c.NamespaceName)).Attribute("val").Value == "1";
			}
			set
			{
				Xml.Element(XName.Get("overlay", DocX.c.NamespaceName)).Attribute("val").Value = GetOverlayValue(value);
			}
		}

		public ChartLegendPosition Position
		{
			get
			{
				return XElementHelpers.GetValueToEnum<ChartLegendPosition>(Xml.Element(XName.Get("legendPos", DocX.c.NamespaceName)));
			}
			set
			{
				XElementHelpers.SetValueFromEnum(Xml.Element(XName.Get("legendPos", DocX.c.NamespaceName)), value);
			}
		}

		internal ChartLegend(ChartLegendPosition position, bool overlay)
		{
			Xml = new XElement(XName.Get("legend", DocX.c.NamespaceName), new XElement(XName.Get("legendPos", DocX.c.NamespaceName), new XAttribute("val", XElementHelpers.GetXmlNameFromEnum(position))), new XElement(XName.Get("overlay", DocX.c.NamespaceName), new XAttribute("val", GetOverlayValue(overlay))));
		}

		private string GetOverlayValue(bool overlay)
		{
			if (overlay)
			{
				return "1";
			}
			return "0";
		}
	}
}
