using System;
using System.Collections.Generic;
using System.IO.Packaging;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace Novacode
{
	public class Hyperlink : DocXElement
	{
		internal Uri uri;

		internal string text;

		internal Dictionary<PackagePart, PackageRelationship> hyperlink_rels;

		internal int type;

		internal string id;

		internal XElement instrText;

		internal List<XElement> runs;

		public string Text
		{
			get
			{
				return text;
			}
			set
			{
				XElement rPr = new XElement(DocX.w + "rPr", new XElement(DocX.w + "rStyle", new XAttribute(DocX.w + "val", "Hyperlink")));
				List<XElement> list = HelperFunctions.FormatInput(value, rPr);
				if (type == 0)
				{
					IEnumerable<XElement> source = from r in base.Xml.Elements()
					where r.Name.LocalName == "r"
					select r;
					for (int i = 0; i < source.Count(); i++)
					{
						source.Remove();
					}
					base.Xml.Add(list);
				}
				else
				{
					XElement xElement = XElement.Parse("\r\n                    <w:r xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>\r\n                        <w:fldChar w:fldCharType='separate'/> \r\n                    </w:r>");
					XElement xElement2 = XElement.Parse("\r\n                    <w:r xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>\r\n                        <w:fldChar w:fldCharType='end' /> \r\n                    </w:r>");
					runs.Last().AddAfterSelf(xElement, list, xElement2);
					runs.ForEach(delegate(XElement r)
					{
						r.Remove();
					});
				}
				text = value;
			}
		}

		public Uri Uri
		{
			get
			{
				if (type == 0 && id != string.Empty)
				{
					PackageRelationship relationship = mainPart.GetRelationship(id);
					return relationship.TargetUri;
				}
				return uri;
			}
			set
			{
				//IL_0022: Unknown result type (might be due to invalid IL or missing references)
				//IL_0027: Unknown result type (might be due to invalid IL or missing references)
				//IL_004c: Unknown result type (might be due to invalid IL or missing references)
				if (type == 0)
				{
					PackageRelationship relationship = mainPart.GetRelationship(id);
					TargetMode targetMode = relationship.TargetMode;
					string relationshipType = relationship.RelationshipType;
					string text = relationship.Id;
					mainPart.DeleteRelationship(text);
					mainPart.CreateRelationship(value, targetMode, relationshipType, text);
				}
				else
				{
					instrText.Value = "HYPERLINK \"" + value + "\"";
				}
				uri = value;
			}
		}

		public void Remove()
		{
			base.Xml.Remove();
		}

		internal Hyperlink(DocX document, PackagePart mainPart, XElement i)
			: base(document, i)
		{
			type = 0;
			id = i.Attribute(XName.Get("id", DocX.r.NamespaceName)).Value;
			StringBuilder sb = new StringBuilder();
			HelperFunctions.GetTextRecursive(i, ref sb);
			text = sb.ToString();
		}

		internal Hyperlink(DocX document, XElement instrText, List<XElement> runs)
			: base(document, null)
		{
			type = 1;
			this.instrText = instrText;
			this.runs = runs;
			try
			{
				int num = instrText.Value.IndexOf("HYPERLINK \"") + "HYPERLINK \"".Length;
				int num2 = instrText.Value.IndexOf("\"", num);
				if (num != -1 && num2 != -1)
				{
					uri = new Uri(instrText.Value.Substring(num, num2 - num), UriKind.Absolute);
					StringBuilder sb = new StringBuilder();
					HelperFunctions.GetTextRecursive(new XElement(XName.Get("temp", DocX.w.NamespaceName), runs), ref sb);
					text = sb.ToString();
				}
			}
			catch (Exception ex)
			{
				throw ex;
			}
		}
	}
}
