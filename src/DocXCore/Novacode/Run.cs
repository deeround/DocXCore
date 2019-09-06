using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace Novacode
{
	public class Run : DocXElement
	{
		private Dictionary<int, Text> textLookup = new Dictionary<int, Text>();

		private int startIndex;

		private int endIndex;

		private string text;

		public int StartIndex => startIndex;

		public int EndIndex => endIndex;

		internal string Value
		{
			get
			{
				return text;
			}
			set
			{
				text = value;
			}
		}

		internal Run(DocX document, XElement xml, int startIndex)
			: base(document, xml)
		{
			this.startIndex = startIndex;
			IEnumerable<XElement> enumerable = xml.Descendants();
			int num = startIndex;
			foreach (XElement item in enumerable)
			{
				switch (item.Name.LocalName)
				{
				case "tab":
					textLookup.Add(num + 1, new Text(base.Document, item, num));
					text += "\t";
					num++;
					break;
				case "br":
					textLookup.Add(num + 1, new Text(base.Document, item, num));
					text += "\n";
					num++;
					break;
				case "t":
				case "delText":
					if (item.Value.Length > 0)
					{
						textLookup.Add(num + item.Value.Length, new Text(base.Document, item, num));
						text += item.Value;
						num += item.Value.Length;
					}
					break;
				}
			}
			endIndex = num;
		}

		internal static XElement[] SplitRun(Run r, int index, EditType type = EditType.ins)
		{
			index -= r.StartIndex;
			Text firstTextEffectedByEdit = r.GetFirstTextEffectedByEdit(index, type);
			XElement[] array = Text.SplitText(firstTextEffectedByEdit, index);
			XElement xElement = new XElement(r.Xml.Name, r.Xml.Attributes(), r.Xml.Element(XName.Get("rPr", DocX.w.NamespaceName)), from n in firstTextEffectedByEdit.Xml.ElementsBeforeSelf()
			where n.Name.LocalName != "rPr"
			select n, array[0]);
			if (Paragraph.GetElementTextLength(xElement) == 0)
			{
				xElement = null;
			}
			XElement xElement2 = new XElement(r.Xml.Name, r.Xml.Attributes(), r.Xml.Element(XName.Get("rPr", DocX.w.NamespaceName)), array[1], from n in firstTextEffectedByEdit.Xml.ElementsAfterSelf()
			where n.Name.LocalName != "rPr"
			select n);
			if (Paragraph.GetElementTextLength(xElement2) == 0)
			{
				xElement2 = null;
			}
			return new XElement[2]
			{
				xElement,
				xElement2
			};
		}

		internal Text GetFirstTextEffectedByEdit(int index, EditType type = EditType.ins)
		{
			if (index < 0 || index > HelperFunctions.GetText(base.Xml).Length)
			{
				throw new ArgumentOutOfRangeException();
			}
			int count = 0;
			Text theOne = null;
			GetFirstTextEffectedByEditRecursive(base.Xml, index, ref count, ref theOne, type);
			return theOne;
		}

		internal void GetFirstTextEffectedByEditRecursive(XElement Xml, int index, ref int count, ref Text theOne, EditType type = EditType.ins)
		{
			count += HelperFunctions.GetSize(Xml);
			if (count > 0 && ((type == EditType.del && count > index) || (type == EditType.ins && count >= index)))
			{
				theOne = new Text(base.Document, Xml, count - HelperFunctions.GetSize(Xml));
			}
			else if (Xml.HasElements)
			{
				foreach (XElement item in Xml.Elements())
				{
					if (theOne == null)
					{
						GetFirstTextEffectedByEditRecursive(item, index, ref count, ref theOne);
					}
				}
			}
		}
	}
}
