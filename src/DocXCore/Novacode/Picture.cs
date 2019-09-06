using System;
using System.Collections.Generic;
using System.IO.Packaging;
using System.Linq;
using System.Xml.Linq;

namespace Novacode
{
	public class Picture : DocXElement
	{
		private const int EmusInPixel = 9525;

		internal Dictionary<PackagePart, PackageRelationship> picture_rels;

		internal Image img;

		private string id;

		private string name;

		private string descr;

		private int cx;

		private int cy;

		private uint rotation;

		private bool hFlip;

		private bool vFlip;

		private object pictureShape;

		private XElement xfrm;

		private XElement prstGeom;

		public string Id => id;

		public bool FlipHorizontal
		{
			get
			{
				return hFlip;
			}
			set
			{
				hFlip = value;
				XAttribute xAttribute = xfrm.Attribute(XName.Get("flipH"));
				if (xAttribute == null)
				{
					xfrm.Add(new XAttribute(XName.Get("flipH"), "0"));
				}
				xfrm.Attribute(XName.Get("flipH")).Value = (hFlip ? "1" : "0");
			}
		}

		public bool FlipVertical
		{
			get
			{
				return vFlip;
			}
			set
			{
				vFlip = value;
				XAttribute xAttribute = xfrm.Attribute(XName.Get("flipV"));
				if (xAttribute == null)
				{
					xfrm.Add(new XAttribute(XName.Get("flipV"), "0"));
				}
				xfrm.Attribute(XName.Get("flipV")).Value = (vFlip ? "1" : "0");
			}
		}

		public uint Rotation
		{
			get
			{
				return rotation / 60000u;
			}
			set
			{
				rotation = value % 360u * 60000;
				XElement xElement = (from d in base.Xml.Descendants()
				where d.Name.LocalName.Equals("xfrm")
				select d).Single();
				XAttribute xAttribute = xElement.Attribute(XName.Get("rot"));
				if (xAttribute == null)
				{
					xElement.Add(new XAttribute(XName.Get("rot"), 0));
				}
				xElement.Attribute(XName.Get("rot")).Value = rotation.ToString();
			}
		}

		public string Name
		{
			get
			{
				return name;
			}
			set
			{
				name = value;
				foreach (XAttribute item in base.Xml.Descendants().Attributes(XName.Get("name")))
				{
					item.Value = name;
				}
			}
		}

		public string Description
		{
			get
			{
				return descr;
			}
			set
			{
				descr = value;
				foreach (XAttribute item in base.Xml.Descendants().Attributes(XName.Get("descr")))
				{
					item.Value = descr;
				}
			}
		}

		public string FileName => img.FileName;

		public int Width
		{
			get
			{
				return cx / 9525;
			}
			set
			{
				cx = value * 9525;
				foreach (XAttribute item in base.Xml.Descendants().Attributes(XName.Get("cx")))
				{
					item.Value = cx.ToString();
				}
			}
		}

		public int Height
		{
			get
			{
				return cy / 9525;
			}
			set
			{
				cy = value * 9525;
				foreach (XAttribute item in base.Xml.Descendants().Attributes(XName.Get("cy")))
				{
					item.Value = cy.ToString();
				}
			}
		}

		public void Remove()
		{
			base.Xml.Remove();
		}

		internal Picture(DocX document, XElement i, Image img)
			: base(document, i)
		{
			picture_rels = new Dictionary<PackagePart, PackageRelationship>();
			this.img = img;
			id = (from e in base.Xml.Descendants()
			where e.Name.LocalName.Equals("blip")
			select e.Attribute(XName.Get("embed", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")).Value).SingleOrDefault();
			if (id == null)
			{
				id = (from e in base.Xml.Descendants()
				where e.Name.LocalName.Equals("imagedata")
				select e.Attribute(XName.Get("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")).Value).SingleOrDefault();
			}
			name = (from e in base.Xml.Descendants()
			let a = e.Attribute(XName.Get("name"))
			where a != null
			select a.Value).FirstOrDefault();
			if (name == null)
			{
				name = (from e in base.Xml.Descendants()
				let a = e.Attribute(XName.Get("title"))
				where a != null
				select a.Value).FirstOrDefault();
			}
			descr = (from e in base.Xml.Descendants()
			let a = e.Attribute(XName.Get("descr"))
			where a != null
			select a.Value).FirstOrDefault();
			cx = (from e in base.Xml.Descendants()
			let a = e.Attribute(XName.Get("cx"))
			where a != null
			select int.Parse(a.Value)).FirstOrDefault();
			if (cx == 0)
			{
				XAttribute xAttribute = (from e in base.Xml.Descendants()
				let a = e.Attribute(XName.Get("style"))
				where a != null
				select a).FirstOrDefault();
				string text = xAttribute.Value.Substring(xAttribute.Value.IndexOf("width:") + 6);
				double value = double.Parse(text.Substring(0, text.IndexOf("pt")).Replace(".", ",")) / 72.0 * 914400.0;
				cx = Convert.ToInt32(value);
			}
			cy = (from e in base.Xml.Descendants()
			let a = e.Attribute(XName.Get("cy"))
			where a != null
			select int.Parse(a.Value)).FirstOrDefault();
			if (cy == 0)
			{
				XAttribute xAttribute2 = (from e in base.Xml.Descendants()
				let a = e.Attribute(XName.Get("style"))
				where a != null
				select a).FirstOrDefault();
				string text2 = xAttribute2.Value.Substring(xAttribute2.Value.IndexOf("height:") + 7);
				double value2 = double.Parse(text2.Substring(0, text2.IndexOf("pt")).Replace(".", ",")) / 72.0 * 914400.0;
				cy = Convert.ToInt32(value2);
			}
			xfrm = (from d in base.Xml.Descendants()
			where d.Name.LocalName.Equals("xfrm")
			select d).SingleOrDefault();
			prstGeom = (from d in base.Xml.Descendants()
			where d.Name.LocalName.Equals("prstGeom")
			select d).SingleOrDefault();
			if (xfrm != null)
			{
				rotation = ((xfrm.Attribute(XName.Get("rot")) != null) ? uint.Parse(xfrm.Attribute(XName.Get("rot")).Value) : 0);
			}
		}

		private void SetPictureShape(object shape)
		{
			pictureShape = shape;
			XAttribute xAttribute = prstGeom.Attribute(XName.Get("prst"));
			if (xAttribute == null)
			{
				prstGeom.Add(new XAttribute(XName.Get("prst"), "rectangle"));
			}
			prstGeom.Attribute(XName.Get("prst")).Value = shape.ToString();
		}

		public void SetPictureShape(BasicShapes shape)
		{
			SetPictureShape((object)shape);
		}

		public void SetPictureShape(RectangleShapes shape)
		{
			SetPictureShape((object)shape);
		}

		public void SetPictureShape(BlockArrowShapes shape)
		{
			SetPictureShape((object)shape);
		}

		public void SetPictureShape(EquationShapes shape)
		{
			SetPictureShape((object)shape);
		}

		public void SetPictureShape(FlowchartShapes shape)
		{
			SetPictureShape((object)shape);
		}

		public void SetPictureShape(StarAndBannerShapes shape)
		{
			SetPictureShape((object)shape);
		}

		public void SetPictureShape(CalloutShapes shape)
		{
			SetPictureShape((object)shape);
		}
	}
}
