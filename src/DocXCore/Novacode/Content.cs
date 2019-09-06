using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace Novacode
{
	public class Content : InsertBeforeOrAfter
	{
		public ContainerType ParentContainer;

		public string Name
		{
			get;
			set;
		}

		public string Tag
		{
			get;
			set;
		}

		public string Text
		{
			get;
			set;
		}

		public List<Content> Sections
		{
			get;
			set;
		}

		internal Content(DocX document, XElement xml, int startIndex)
			: base(document, xml)
		{
		}

		public void SetText(string newText)
		{
			base.Xml.Descendants(XName.Get("t", DocX.w.NamespaceName)).First().Value = newText;
		}
	}
}
