using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Xml.Linq;

namespace Novacode
{
	internal static class Extensions
	{
		internal static string ToHex(this Color source)
		{
			byte r = source.R;
			byte g = source.G;
			byte b = source.B;
			string text = r.ToString("X");
			if (text.Length < 2)
			{
				text = "0" + text;
			}
			string text2 = b.ToString("X");
			if (text2.Length < 2)
			{
				text2 = "0" + text2;
			}
			string text3 = g.ToString("X");
			if (text3.Length < 2)
			{
				text3 = "0" + text3;
			}
			return $"{text}{text3}{text2}";
		}

		public static void Flatten(this XElement e, XName name, List<XElement> flat)
		{
			XElement xElement = CloneElement(e);
			xElement.Elements().Remove();
			if (xElement.Name == name)
			{
				flat.Add(xElement);
			}
			if (e.HasElements)
			{
				foreach (XElement item in e.Elements(name))
				{
					item.Flatten(name, flat);
				}
			}
		}

		private static XElement CloneElement(XElement element)
		{
			return new XElement(element.Name, element.Attributes(), element.Nodes().Select(delegate(XNode n)
			{
				XElement xElement = n as XElement;
				if (xElement != null)
				{
					return CloneElement(xElement);
				}
				return n;
			}));
		}

		public static string GetAttribute(this XElement el, XName name, string defaultValue = "")
		{
			XAttribute xAttribute = el.Attribute(name);
			if (xAttribute != null)
			{
				return xAttribute.Value;
			}
			return defaultValue;
		}

		public static void SetMargin(this DocX document, float top, float bottom, float right, float left)
		{
			XNamespace ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
			IEnumerable<XElement> enumerable = document.PageLayout.Xml.Descendants(ns + "pgMar");
			foreach (XElement item in enumerable)
			{
				if (left != -1f)
				{
					item.SetAttributeValue(ns + "left", 1440f * left / 1f);
				}
				if (right != -1f)
				{
					item.SetAttributeValue(ns + "right", 1440f * right / 1f);
				}
				if (top != -1f)
				{
					item.SetAttributeValue(ns + "top", 1440f * top / 1f);
				}
				if (bottom != -1f)
				{
					item.SetAttributeValue(ns + "bottom", 1440f * bottom / 1f);
				}
			}
		}
	}
}
