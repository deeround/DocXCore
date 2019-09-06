using System;
using System.Collections;
using System.Drawing;
using System.Globalization;
using System.Xml.Linq;

namespace Novacode
{
	public class Series
	{
		private XElement strCache;

		private XElement numCache;

		internal XElement Xml
		{
			get;
			private set;
		}

		public Color Color
		{
			get
			{
				XElement xElement = Xml.Element(XName.Get("spPr", DocX.c.NamespaceName));
				if (xElement == null)
				{
					return Color.Transparent;
				}
				return Color.FromArgb(int.Parse(xElement.Element(XName.Get("solidFill", DocX.a.NamespaceName)).Element(XName.Get("srgbClr", DocX.a.NamespaceName)).Attribute(XName.Get("val"))
					.Value, NumberStyles.HexNumber));
				}
				set
				{
					Xml.Element(XName.Get("spPr", DocX.c.NamespaceName))?.Remove();
					XElement content = new XElement(XName.Get("spPr", DocX.c.NamespaceName), new XElement(XName.Get("solidFill", DocX.a.NamespaceName), new XElement(XName.Get("srgbClr", DocX.a.NamespaceName), new XAttribute(XName.Get("val"), value.ToHex()))));
					Xml.Element(XName.Get("tx", DocX.c.NamespaceName)).AddAfterSelf(content);
				}
			}

			internal Series(XElement xml)
			{
				Xml = xml;
				strCache = xml.Element(XName.Get("cat", DocX.c.NamespaceName)).Element(XName.Get("strRef", DocX.c.NamespaceName)).Element(XName.Get("strCache", DocX.c.NamespaceName));
				numCache = xml.Element(XName.Get("val", DocX.c.NamespaceName)).Element(XName.Get("numRef", DocX.c.NamespaceName)).Element(XName.Get("numCache", DocX.c.NamespaceName));
			}

			public Series(string name)
			{
				strCache = new XElement(XName.Get("strCache", DocX.c.NamespaceName));
				numCache = new XElement(XName.Get("numCache", DocX.c.NamespaceName));
				Xml = new XElement(XName.Get("ser", DocX.c.NamespaceName), new XElement(XName.Get("tx", DocX.c.NamespaceName), new XElement(XName.Get("strRef", DocX.c.NamespaceName), new XElement(XName.Get("f", DocX.c.NamespaceName), ""), new XElement(XName.Get("strCache", DocX.c.NamespaceName), new XElement(XName.Get("pt", DocX.c.NamespaceName), new XAttribute(XName.Get("idx"), "0"), new XElement(XName.Get("v", DocX.c.NamespaceName), name))))), new XElement(XName.Get("invertIfNegative", DocX.c.NamespaceName), "0"), new XElement(XName.Get("cat", DocX.c.NamespaceName), new XElement(XName.Get("strRef", DocX.c.NamespaceName), new XElement(XName.Get("f", DocX.c.NamespaceName), ""), strCache)), new XElement(XName.Get("val", DocX.c.NamespaceName), new XElement(XName.Get("numRef", DocX.c.NamespaceName), new XElement(XName.Get("f", DocX.c.NamespaceName), ""), numCache)));
			}

			public void Bind(ICollection list, string categoryPropertyName, string valuePropertyName)
			{
				XElement content = new XElement(XName.Get("ptCount", DocX.c.NamespaceName), new XAttribute(XName.Get("val"), list.Count));
				XElement content2 = new XElement(XName.Get("formatCode", DocX.c.NamespaceName), "General");
				strCache.RemoveAll();
				numCache.RemoveAll();
				strCache.Add(content);
				numCache.Add(content2);
				numCache.Add(content);
				int num = 0;
				foreach (object item in list)
				{
					XElement content3 = new XElement(XName.Get("pt", DocX.c.NamespaceName), new XAttribute(XName.Get("idx"), num), new XElement(XName.Get("v", DocX.c.NamespaceName), item.GetType().GetProperty(categoryPropertyName).GetValue(item, null)));
					strCache.Add(content3);
					content3 = new XElement(XName.Get("pt", DocX.c.NamespaceName), new XAttribute(XName.Get("idx"), num), new XElement(XName.Get("v", DocX.c.NamespaceName), item.GetType().GetProperty(valuePropertyName).GetValue(item, null)));
					numCache.Add(content3);
					num++;
				}
			}

			public void Bind(IList categories, IList values)
			{
				if (categories.Count != values.Count)
				{
					throw new ArgumentException("Categories count must equal to Values count");
				}
				XElement content = new XElement(XName.Get("ptCount", DocX.c.NamespaceName), new XAttribute(XName.Get("val"), categories.Count));
				XElement content2 = new XElement(XName.Get("formatCode", DocX.c.NamespaceName), "General");
				strCache.RemoveAll();
				numCache.RemoveAll();
				strCache.Add(content);
				numCache.Add(content2);
				numCache.Add(content);
				for (int i = 0; i < categories.Count; i++)
				{
					XElement content3 = new XElement(XName.Get("pt", DocX.c.NamespaceName), new XAttribute(XName.Get("idx"), i), new XElement(XName.Get("v", DocX.c.NamespaceName), categories[i].ToString()));
					strCache.Add(content3);
					content3 = new XElement(XName.Get("pt", DocX.c.NamespaceName), new XAttribute(XName.Get("idx"), i), new XElement(XName.Get("v", DocX.c.NamespaceName), values[i].ToString()));
					numCache.Add(content3);
				}
			}
		}
	}
