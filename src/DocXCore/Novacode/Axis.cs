using System.Xml.Linq;

namespace Novacode
{
	public abstract class Axis
	{
		public string Id => Xml.Element(XName.Get("axId", DocX.c.NamespaceName)).Attribute(XName.Get("val")).Value;

		public bool IsVisible
		{
			get
			{
				return Xml.Element(XName.Get("delete", DocX.c.NamespaceName)).Attribute(XName.Get("val")).Value == "0";
			}
			set
			{
				if (value)
				{
					Xml.Element(XName.Get("delete", DocX.c.NamespaceName)).Attribute(XName.Get("val")).Value = "0";
				}
				else
				{
					Xml.Element(XName.Get("delete", DocX.c.NamespaceName)).Attribute(XName.Get("val")).Value = "1";
				}
			}
		}

		internal XElement Xml
		{
			get;
			set;
		}

		internal Axis(XElement xml)
		{
			Xml = xml;
		}

		public Axis(string id)
		{
		}
	}
}
