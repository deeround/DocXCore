using System.Collections.Generic;
using System.Xml.Linq;

namespace Novacode
{
	public class Section : Container
	{
		public SectionBreakType SectionBreakType;

		public List<Paragraph> SectionParagraphs
		{
			get;
			set;
		}

		internal Section(DocX document, XElement xml)
			: base(document, xml)
		{
		}
	}
}
