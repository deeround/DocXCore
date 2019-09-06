using System.IO.Packaging;
using System.Xml.Linq;

namespace Novacode
{
	public abstract class DocXElement
	{
		internal PackagePart mainPart;

		public PackagePart PackagePart
		{
			get
			{
				return mainPart;
			}
			set
			{
				mainPart = value;
			}
		}

		public XElement Xml
		{
			get;
			set;
		}

		internal DocX Document
		{
			get;
			set;
		}

		public DocXElement(DocX document, XElement xml)
		{
			Document = document;
			Xml = xml;
		}
	}
}
