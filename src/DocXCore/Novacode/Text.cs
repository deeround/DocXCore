using System;
using System.Linq;
using System.Xml.Linq;

namespace Novacode
{
	internal class Text : DocXElement
	{
		private int startIndex;

		private int endIndex;

		private string text;

		public int StartIndex => startIndex;

		public int EndIndex => endIndex;

		public string Value => text;

		internal Text(DocX document, XElement xml, int startIndex)
			: base(document, xml)
		{
			this.startIndex = startIndex;
			string localName = base.Xml.Name.LocalName;
			switch (localName)
			{
			default:
				if (localName == "tab")
				{
					text = "\t";
					endIndex = startIndex + 1;
				}
				break;
			case "t":
			case "delText":
				endIndex = startIndex + xml.Value.Length;
				text = xml.Value;
				break;
			case "br":
				text = "\n";
				endIndex = startIndex + 1;
				break;
			}
		}

		internal static XElement[] SplitText(Text t, int index)
		{
			if (index < t.startIndex || index > t.EndIndex)
			{
				throw new ArgumentOutOfRangeException("index");
			}
			XElement xElement = null;
			XElement xElement2 = null;
			if (t.Xml.Name.LocalName == "t" || t.Xml.Name.LocalName == "delText")
			{
				xElement = new XElement(t.Xml.Name, t.Xml.Attributes(), t.Xml.Value.Substring(0, index - t.startIndex));
				if (xElement.Value.Length == 0)
				{
					xElement = null;
				}
				else
				{
					PreserveSpace(xElement);
				}
				xElement2 = new XElement(t.Xml.Name, t.Xml.Attributes(), t.Xml.Value.Substring(index - t.startIndex, t.Xml.Value.Length - (index - t.startIndex)));
				if (xElement2.Value.Length == 0)
				{
					xElement2 = null;
				}
				else
				{
					PreserveSpace(xElement2);
				}
			}
			else if (index == t.EndIndex)
			{
				xElement = t.Xml;
			}
			else
			{
				xElement2 = t.Xml;
			}
			return new XElement[2]
			{
				xElement,
				xElement2
			};
		}

		public static void PreserveSpace(XElement e)
		{
			if (!e.Name.Equals(DocX.w + "t") && !e.Name.Equals(DocX.w + "delText"))
			{
				throw new ArgumentException("SplitText can only split elements of type t or delText", "e");
			}
			XAttribute xAttribute = (from a in e.Attributes()
			where a.Name.Equals(XNamespace.Xml + "space")
			select a).SingleOrDefault();
			if (e.Value.StartsWith(" ") || e.Value.EndsWith(" "))
			{
				if (xAttribute == null)
				{
					e.Add(new XAttribute(XNamespace.Xml + "space", "preserve"));
				}
			}
			else
			{
				xAttribute?.Remove();
			}
		}
	}
}
