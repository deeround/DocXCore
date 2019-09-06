using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace Novacode
{
	public class DocProperty : DocXElement
	{
		internal Regex extractName = new Regex("DOCPROPERTY  (?<name>.*)  ");

		public string Name
		{
			get;
		}

		internal DocProperty(DocX document, XElement xml)
			: base(document, xml)
		{
			string value = base.Xml.Attribute(XName.Get("instr", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")).Value;
			Name = extractName.Match(value.Trim()).Groups["name"].Value;
		}
	}
}
