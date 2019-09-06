using System;
using System.Xml.Linq;

namespace Novacode
{
	public class PageLayout : DocXElement
	{
		public Orientation Orientation
		{
			get
			{
				XElement xElement = base.Xml.Element(XName.Get("pgSz", DocX.w.NamespaceName));
				if (xElement == null)
				{
					return Orientation.Portrait;
				}
				XAttribute xAttribute = xElement.Attribute(XName.Get("orient", DocX.w.NamespaceName));
				if (xAttribute == null)
				{
					return Orientation.Portrait;
				}
				if (xAttribute.Value.Equals("Landscape", StringComparison.CurrentCultureIgnoreCase))
				{
					return Orientation.Landscape;
				}
				return Orientation.Portrait;
			}
			set
			{
				if (Orientation != value)
				{
					XElement xElement = base.Xml.Element(XName.Get("pgSz", DocX.w.NamespaceName));
					if (xElement == null)
					{
						base.Xml.SetElementValue(XName.Get("pgSz", DocX.w.NamespaceName), string.Empty);
						xElement = base.Xml.Element(XName.Get("pgSz", DocX.w.NamespaceName));
					}
					xElement.SetAttributeValue(XName.Get("orient", DocX.w.NamespaceName), value.ToString().ToLower());
					switch (value)
					{
					case Orientation.Landscape:
						xElement.SetAttributeValue(XName.Get("w", DocX.w.NamespaceName), "16838");
						xElement.SetAttributeValue(XName.Get("h", DocX.w.NamespaceName), "11906");
						break;
					case Orientation.Portrait:
						xElement.SetAttributeValue(XName.Get("w", DocX.w.NamespaceName), "11906");
						xElement.SetAttributeValue(XName.Get("h", DocX.w.NamespaceName), "16838");
						break;
					}
				}
			}
		}

		internal PageLayout(DocX document, XElement xml)
			: base(document, xml)
		{
		}
	}
}
