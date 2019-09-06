using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Linq;

namespace Novacode
{
	public class DocX : Container, IDisposable
	{
		internal static XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

		internal static XNamespace rel = "http://schemas.openxmlformats.org/package/2006/relationships";

		internal static XNamespace r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

		internal static XNamespace m = "http://schemas.openxmlformats.org/officeDocument/2006/math";

		internal static XNamespace customPropertiesSchema = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties";

		internal static XNamespace customVTypesSchema = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes";

		internal static XNamespace wp = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";

		internal static XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";

		internal static XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";

		internal static XNamespace v = "urn:schemas-microsoft-com:vml";

		internal static XNamespace n = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering";

		internal const string relationshipImage = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

		internal const string contentTypeApplicationRelationShipXml = "application/vnd.openxmlformats-package.relationships+xml";

		private Headers headers;

		private Footers footers;

		internal PackagePart settingsPart;

		internal PackagePart endnotesPart;

		internal PackagePart footnotesPart;

		internal PackagePart stylesPart;

		internal PackagePart stylesWithEffectsPart;

		internal PackagePart numberingPart;

		internal PackagePart fontTablePart;

		internal Package package;

		internal XDocument mainDoc;

		internal XDocument settings;

		internal XDocument endnotes;

		internal XDocument footnotes;

		internal XDocument styles;

		internal XDocument stylesWithEffects;

		internal XDocument numbering;

		internal XDocument fontTable;

		internal XDocument header1;

		internal XDocument header2;

		internal XDocument header3;

		internal Dictionary<int, Paragraph> paragraphLookup = new Dictionary<int, Paragraph>();

		internal MemoryStream memoryStream;

		internal string filename;

		internal Stream stream;

		private readonly object nextFreeDocPrIdLock = new object();

		private long? nextFreeDocPrId;

		public BookmarkCollection Bookmarks
		{
			get
			{
				BookmarkCollection bookmarkCollection = new BookmarkCollection();
				foreach (Paragraph paragraph in Paragraphs)
				{
					bookmarkCollection.AddRange(paragraph.GetBookmarks());
				}
				return bookmarkCollection;
			}
		}

		public float MarginTop
		{
			get
			{
				return getMarginAttribute(XName.Get("top", w.NamespaceName));
			}
			set
			{
				setMarginAttribute(XName.Get("top", w.NamespaceName), value);
			}
		}

		public float MarginBottom
		{
			get
			{
				return getMarginAttribute(XName.Get("bottom", w.NamespaceName));
			}
			set
			{
				setMarginAttribute(XName.Get("bottom", w.NamespaceName), value);
			}
		}

		public float MarginLeft
		{
			get
			{
				return getMarginAttribute(XName.Get("left", w.NamespaceName));
			}
			set
			{
				setMarginAttribute(XName.Get("left", w.NamespaceName), value);
			}
		}

		public float MarginRight
		{
			get
			{
				return getMarginAttribute(XName.Get("right", w.NamespaceName));
			}
			set
			{
				setMarginAttribute(XName.Get("right", w.NamespaceName), value);
			}
		}

		public float MarginHeader
		{
			get
			{
				return getMarginAttribute(XName.Get("header", w.NamespaceName));
			}
			set
			{
				setMarginAttribute(XName.Get("header", w.NamespaceName), value);
			}
		}

		public float MarginFooter
		{
			get
			{
				return getMarginAttribute(XName.Get("footer", w.NamespaceName));
			}
			set
			{
				setMarginAttribute(XName.Get("footer", w.NamespaceName), value);
			}
		}

		public bool MirrorMargins
		{
			get
			{
				return getMirrorMargins(XName.Get("mirrorMargins", w.NamespaceName));
			}
			set
			{
				setMirrorMargins(XName.Get("mirrorMargins", w.NamespaceName), value);
			}
		}

		public float PageWidth
		{
			get
			{
				XElement xElement = mainDoc.Root.Element(XName.Get("body", w.NamespaceName));
				XElement xElement2 = xElement.Element(XName.Get("sectPr", w.NamespaceName))?.Element(XName.Get("pgSz", w.NamespaceName));
				if (xElement2 != null)
				{
					XAttribute xAttribute = xElement2.Attribute(XName.Get("w", w.NamespaceName));
					if (xAttribute != null && float.TryParse(xAttribute.Value, out float result))
					{
						return (float)(int)(result / 20f);
					}
				}
				return 612f;
			}
			set
			{
				((mainDoc.Root.Element(XName.Get("body", w.NamespaceName))?.Element(XName.Get("sectPr", w.NamespaceName)))?.Element(XName.Get("pgSz", w.NamespaceName)))?.SetAttributeValue(XName.Get("w", w.NamespaceName), value * 20f);
			}
		}

		public float PageHeight
		{
			get
			{
				XElement xElement = mainDoc.Root.Element(XName.Get("body", w.NamespaceName));
				XElement xElement2 = xElement.Element(XName.Get("sectPr", w.NamespaceName));
				if (xElement2 != null)
				{
					XElement xElement3 = xElement2.Element(XName.Get("pgSz", w.NamespaceName));
					if (xElement3 != null)
					{
						XAttribute xAttribute = xElement3.Attribute(XName.Get("h", w.NamespaceName));
						if (xAttribute != null && float.TryParse(xAttribute.Value, out float result))
						{
							return (float)(int)(result / 20f);
						}
					}
				}
				return 792f;
			}
			set
			{
				mainDoc.Root.Element(XName.Get("body", w.NamespaceName))?.Element(XName.Get("sectPr", w.NamespaceName))?.Element(XName.Get("pgSz", w.NamespaceName))?.SetAttributeValue(XName.Get("h", w.NamespaceName), value * 20f);
			}
		}

		public bool isProtected => settings.Descendants(XName.Get("documentProtection", w.NamespaceName)).Count() > 0;

		public PageLayout PageLayout
		{
			get
			{
				XElement xElement = base.Xml.Element(XName.Get("sectPr", w.NamespaceName));
				if (xElement == null)
				{
					base.Xml.SetElementValue(XName.Get("sectPr", w.NamespaceName), string.Empty);
					xElement = base.Xml.Element(XName.Get("sectPr", w.NamespaceName));
				}
				return new PageLayout(this, xElement);
			}
		}

		public Headers Headers => headers;

		public Footers Footers => footers;

		public bool DifferentOddAndEvenPages
		{
			get
			{
				XDocument xDocument;
				using (TextReader textReader = new StreamReader(settingsPart.GetStream()))
				{
					xDocument = XDocument.Load(textReader);
				}
				XElement xElement = xDocument.Root.Element(w + "evenAndOddHeaders");
				return xElement != null;
			}
			set
			{
				XDocument xDocument;
				using (TextReader textReader = new StreamReader(settingsPart.GetStream()))
				{
					xDocument = XDocument.Load(textReader);
				}
				XElement xElement = xDocument.Root.Element(w + "evenAndOddHeaders");
				if (xElement == null)
				{
					if (value)
					{
						xDocument.Root.AddFirst(new XElement(w + "evenAndOddHeaders"));
					}
				}
				else if (!value)
				{
					xElement.Remove();
				}
				using (TextWriter textWriter = new StreamWriter(new PackagePartStream(settingsPart.GetStream())))
				{
					xDocument.Save(textWriter);
				}
			}
		}

		public bool DifferentFirstPage
		{
			get
			{
				XElement xElement = mainDoc.Root.Element(w + "body");
				XElement xElement2 = xElement.Element(w + "sectPr")?.Element(w + "titlePg");
				return xElement2 != null;
			}
			set
			{
				XElement xElement = mainDoc.Root.Element(w + "body");
				xElement.Add(new XElement(w + "sectPr", string.Empty));
				XElement xElement2 = xElement.Element(w + "sectPr");
				XElement xElement3 = xElement2.Element(w + "titlePg");
				if (xElement3 == null)
				{
					if (value)
					{
						xElement2.Add(new XElement(w + "titlePg", string.Empty));
					}
				}
				else if (!value)
				{
					xElement3.Remove();
				}
			}
		}

		public List<Image> Images
		{
			get
			{
				PackageRelationshipCollection relationshipsByType = mainPart.GetRelationshipsByType("http://schemas.openxmlformats.org/officeDocument/2006/relationships/image");
				if (((IEnumerable<PackageRelationship>)relationshipsByType).Any())
				{
					return (from i in (IEnumerable<PackageRelationship>)relationshipsByType
					select new Image(this, i)).ToList();
				}
				return new List<Image>();
			}
		}

		public Dictionary<string, CustomProperty> CustomProperties
		{
			get
			{
				if (package.PartExists(new Uri("/docProps/custom.xml", UriKind.Relative)))
				{
					PackagePart part = package.GetPart(new Uri("/docProps/custom.xml", UriKind.Relative));
					XDocument xDocument;
					using (TextReader textReader = new StreamReader(part.GetStream(FileMode.Open, FileAccess.Read)))
					{
						xDocument = XDocument.Load(textReader, LoadOptions.PreserveWhitespace);
					}
					return (from p in xDocument.Descendants(XName.Get("property", customPropertiesSchema.NamespaceName))
					let Name = p.Attribute(XName.Get("name")).Value
					let Type = p.Descendants().Single().Name.LocalName
					let Value = p.Descendants().Single().Value
					select new CustomProperty(Name, Type, Value)).ToDictionary((CustomProperty p) => p.Name, StringComparer.CurrentCultureIgnoreCase);
				}
				return new Dictionary<string, CustomProperty>();
			}
		}

		public Dictionary<string, string> CoreProperties
		{
			get
			{
				if (package.PartExists(new Uri("/docProps/core.xml", UriKind.Relative)))
				{
					PackagePart part = package.GetPart(new Uri("/docProps/core.xml", UriKind.Relative));
					XDocument corePropDoc;
					using (TextReader textReader = new StreamReader(part.GetStream(FileMode.Open, FileAccess.Read)))
					{
						corePropDoc = XDocument.Load(textReader, LoadOptions.PreserveWhitespace);
					}
					return (from docProperty in corePropDoc.Root.Elements()
					select new KeyValuePair<string, string>($"{corePropDoc.Root.GetPrefixOfNamespace(docProperty.Name.Namespace)}:{docProperty.Name.LocalName}", docProperty.Value)).ToDictionary((KeyValuePair<string, string> p) => p.Key, (KeyValuePair<string, string> v) => v.Value);
				}
				return new Dictionary<string, string>();
			}
		}

		public string Text => HelperFunctions.GetText(base.Xml);

		public IEnumerable<string> FootnotesText
		{
			get
			{
				foreach (XElement item in footnotes.Root.Elements(w + "footnote"))
				{
					yield return HelperFunctions.GetText(item);
				}
			}
		}

		public IEnumerable<string> EndnotesText
		{
			get
			{
				foreach (XElement item in endnotes.Root.Elements(w + "endnote"))
				{
					yield return HelperFunctions.GetText(item);
				}
			}
		}

		public override ReadOnlyCollection<Content> Contents
		{
			get
			{
				ReadOnlyCollection<Content> contents = base.Contents;
				foreach (Content item in contents)
				{
					item.PackagePart = mainPart;
				}
				return contents;
			}
		}

		public override ReadOnlyCollection<Paragraph> Paragraphs
		{
			get
			{
				ReadOnlyCollection<Paragraph> paragraphs = base.Paragraphs;
				foreach (Paragraph item in paragraphs)
				{
					item.PackagePart = mainPart;
				}
				return paragraphs;
			}
		}

		public override List<List> Lists
		{
			get
			{
				List<List> lists = base.Lists;
				lists.ForEach(delegate(List x)
				{
					x.Items.ForEach(delegate(Paragraph i)
					{
						i.PackagePart = mainPart;
					});
				});
				return lists;
			}
		}

		public override List<Table> Tables
		{
			get
			{
				List<Table> tables = base.Tables;
				tables.ForEach(delegate(Table x)
				{
					x.mainPart = mainPart;
				});
				return tables;
			}
		}

		internal float getMarginAttribute(XName name)
		{
			XElement xElement = mainDoc.Root.Element(XName.Get("body", w.NamespaceName));
			XAttribute xAttribute = (xElement.Element(XName.Get("sectPr", w.NamespaceName))?.Element(XName.Get("pgMar", w.NamespaceName)))?.Attribute(name);
			if (xAttribute != null && float.TryParse(xAttribute.Value, out float result))
			{
				return (float)(int)(result / 20f);
			}
			return 0f;
		}

		internal void setMarginAttribute(XName xName, float value)
		{
			XElement xElement = mainDoc.Root.Element(XName.Get("body", w.NamespaceName));
			((xElement.Element(XName.Get("sectPr", w.NamespaceName))?.Element(XName.Get("pgMar", w.NamespaceName)))?.Attribute(xName))?.SetValue(value * 20f);
		}

		internal bool getMirrorMargins(XName name)
		{
			XElement xElement = mainDoc.Root.Element(XName.Get("body", w.NamespaceName));
			XElement xElement2 = xElement.Element(XName.Get("sectPr", w.NamespaceName));
			if (xElement2 != null)
			{
				XElement xElement3 = xElement2.Element(XName.Get("mirrorMargins", w.NamespaceName));
				if (xElement3 != null)
				{
					return true;
				}
			}
			return false;
		}

		internal void setMirrorMargins(XName name, bool value)
		{
			XElement xElement = mainDoc.Root.Element(XName.Get("body", w.NamespaceName));
			XElement xElement2 = xElement.Element(XName.Get("sectPr", w.NamespaceName));
			if (xElement2 != null)
			{
				XElement xElement3 = xElement2.Element(XName.Get("mirrorMargins", w.NamespaceName));
				if (xElement3 != null)
				{
					if (!value)
					{
						xElement3.Remove();
					}
				}
				else
				{
					xElement2.Add(new XElement(w + "mirrorMargins", string.Empty));
				}
			}
		}

		public EditRestrictions GetProtectionType()
		{
			if (isProtected)
			{
				XElement xElement = settings.Descendants(XName.Get("documentProtection", w.NamespaceName)).FirstOrDefault();
				string value = xElement.Attribute(XName.Get("edit", w.NamespaceName)).Value;
				return (EditRestrictions)Enum.Parse(typeof(EditRestrictions), value);
			}
			return EditRestrictions.none;
		}

		public void AddProtection(EditRestrictions er)
		{
			RemoveProtection();
			if (er != 0)
			{
				XElement xElement = new XElement(XName.Get("documentProtection", w.NamespaceName));
				xElement.Add(new XAttribute(XName.Get("edit", w.NamespaceName), er.ToString()));
				xElement.Add(new XAttribute(XName.Get("enforcement", w.NamespaceName), "1"));
				settings.Root.AddFirst(xElement);
			}
		}

		public void AddProtection(EditRestrictions er, string strPassword)
		{
			RemoveProtection();
			if (er != 0)
			{
				XElement xElement = new XElement(XName.Get("documentProtection", w.NamespaceName));
				xElement.Add(new XAttribute(XName.Get("edit", w.NamespaceName), er.ToString()));
				xElement.Add(new XAttribute(XName.Get("enforcement", w.NamespaceName), "1"));
				int[] array = new int[15]
				{
					57840,
					7439,
					52380,
					33984,
					4364,
					3600,
					61902,
					12606,
					6258,
					57657,
					54287,
					34041,
					10252,
					43370,
					20163
				};
				int[,] array2 = new int[15, 7]
				{
					{
						44796,
						19929,
						39858,
						10053,
						20106,
						40212,
						10761
					},
					{
						31585,
						63170,
						64933,
						60267,
						50935,
						40399,
						11199
					},
					{
						17763,
						35526,
						1453,
						2906,
						5812,
						11624,
						23248
					},
					{
						885,
						1770,
						3540,
						7080,
						14160,
						28320,
						56640
					},
					{
						55369,
						41139,
						20807,
						41614,
						21821,
						43642,
						17621
					},
					{
						28485,
						56970,
						44341,
						19019,
						38038,
						14605,
						29210
					},
					{
						60195,
						50791,
						40175,
						10751,
						21502,
						43004,
						24537
					},
					{
						18387,
						36774,
						3949,
						7898,
						15796,
						31592,
						63184
					},
					{
						47201,
						24803,
						49606,
						37805,
						14203,
						28406,
						56812
					},
					{
						17824,
						35648,
						1697,
						3394,
						6788,
						13576,
						27152
					},
					{
						43601,
						17539,
						35078,
						557,
						1114,
						2228,
						4456
					},
					{
						30388,
						60776,
						51953,
						34243,
						7079,
						14158,
						28316
					},
					{
						14128,
						28256,
						56512,
						43425,
						17251,
						34502,
						7597
					},
					{
						13105,
						26210,
						52420,
						35241,
						883,
						1766,
						3532
					},
					{
						4129,
						8258,
						16516,
						33032,
						4657,
						9314,
						18628
					}
				};
				byte[] array3 = new byte[16];
				RandomNumberGenerator randomNumberGenerator = RandomNumberGenerator.Create();
				randomNumberGenerator.GetBytes(array3);
				byte[] array4 = new byte[4];
				int num = 15;
				if (!string.IsNullOrEmpty(strPassword))
				{
					strPassword = strPassword.Substring(0, Math.Min(strPassword.Length, num));
					byte[] array5 = new byte[strPassword.Length];
					for (int i = 0; i < strPassword.Length; i++)
					{
						int num2 = Convert.ToInt32(strPassword[i]);
						array5[i] = Convert.ToByte(num2 & 0xFF);
						if (array5[i] == 0)
						{
							array5[i] = Convert.ToByte((num2 & 0xFF00) >> 8);
						}
					}
					int num3 = array[array5.Length - 1];
					for (int j = 0; j < array5.Length; j++)
					{
						int num4 = num - array5.Length + j;
						for (int k = 0; k < 7; k++)
						{
							if ((array5[j] & (1 << k)) != 0)
							{
								num3 ^= array2[num4, k];
							}
						}
					}
					int num5 = 0;
					for (int num6 = array5.Length - 1; num6 >= 0; num6--)
					{
						num5 = ((((num5 >> 14) & 1) | ((num5 << 1) & 0x7FFF)) ^ array5[num6]);
					}
					num5 = ((((num5 >> 14) & 1) | ((num5 << 1) & 0x7FFF)) ^ array5.Length ^ 0xCE4B);
					int num7 = (num3 << 16) + num5;
					for (int l = 0; l < 4; l++)
					{
						array4[l] = Convert.ToByte((uint)(num7 & (255 << l * 8)) >> l * 8);
					}
				}
				StringBuilder stringBuilder = new StringBuilder();
				for (int m = 0; m < 4; m++)
				{
					stringBuilder.Append(Convert.ToString(array4[m], 16));
				}
				array4 = Encoding.Unicode.GetBytes(stringBuilder.ToString().ToUpper());
				byte[] array6 = array4;
				byte[] array7 = array3;
				byte[] array8 = new byte[array6.Length + array7.Length];
				Buffer.BlockCopy(array7, 0, array8, 0, array7.Length);
				Buffer.BlockCopy(array6, 0, array8, array7.Length, array6.Length);
				array4 = array8;
				int num8 = 100000;
				HashAlgorithm hashAlgorithm = SHA1.Create();
				array4 = hashAlgorithm.ComputeHash(array4);
				byte[] array9 = new byte[4];
				for (int n = 0; n < num8; n++)
				{
					array9[0] = Convert.ToByte(n & 0xFF);
					array9[1] = Convert.ToByte((n & 0xFF00) >> 8);
					array9[2] = Convert.ToByte((n & 0xFF0000) >> 16);
					array9[3] = Convert.ToByte((n & 4278190080u) >> 24);
					array4 = concatByteArrays(array9, array4);
					array4 = hashAlgorithm.ComputeHash(array4);
				}
				xElement.Add(new XAttribute(XName.Get("cryptProviderType", w.NamespaceName), "rsaFull"));
				xElement.Add(new XAttribute(XName.Get("cryptAlgorithmClass", w.NamespaceName), "hash"));
				xElement.Add(new XAttribute(XName.Get("cryptAlgorithmType", w.NamespaceName), "typeAny"));
				xElement.Add(new XAttribute(XName.Get("cryptAlgorithmSid", w.NamespaceName), "4"));
				xElement.Add(new XAttribute(XName.Get("cryptSpinCount", w.NamespaceName), num8.ToString()));
				xElement.Add(new XAttribute(XName.Get("hash", w.NamespaceName), Convert.ToBase64String(array4)));
				xElement.Add(new XAttribute(XName.Get("salt", w.NamespaceName), Convert.ToBase64String(array3)));
				settings.Root.AddFirst(xElement);
			}
		}

		private byte[] concatByteArrays(byte[] array1, byte[] array2)
		{
			byte[] array3 = new byte[array1.Length + array2.Length];
			Buffer.BlockCopy(array2, 0, array3, 0, array2.Length);
			Buffer.BlockCopy(array1, 0, array3, array2.Length, array1.Length);
			return array3;
		}

		public void RemoveProtection()
		{
			settings.Descendants(XName.Get("documentProtection", w.NamespaceName)).Remove();
		}

		private Header GetHeaderByType(string type)
		{
			return (Header)GetHeaderOrFooterByType(type, isHeader: true);
		}

		private Footer GetFooterByType(string type)
		{
			return (Footer)GetHeaderOrFooterByType(type, isHeader: false);
		}

		private object GetHeaderOrFooterByType(string type, bool isHeader)
		{
			string reference = "footerReference";
			if (isHeader)
			{
				reference = "headerReference";
			}
			string text = (from e in mainDoc.Descendants(XName.Get("body", w.NamespaceName)).Descendants()
			where e.Name.LocalName == reference && e.Attribute(w + "type").Value == type
			select e.Attribute(r + "id").Value).LastOrDefault();
			if (text != null)
			{
				Uri uri = mainPart.GetRelationship(text).TargetUri;
				if (!uri.OriginalString.StartsWith("/word/"))
				{
					uri = new Uri("/word/" + uri.OriginalString, UriKind.Relative);
				}
				PackagePart part = package.GetPart(uri);
				using (TextReader textReader = new StreamReader(part.GetStream()))
				{
					XDocument xDocument = XDocument.Load(textReader);
					if (isHeader)
					{
						return new Header(this, xDocument.Element(w + "hdr"), part);
					}
					return new Footer(this, xDocument.Element(w + "ftr"), part);
				}
			}
			return null;
		}

		public List<Section> GetSections()
		{
			ReadOnlyCollection<Paragraph> paragraphs = Paragraphs;
			List<Paragraph> list = new List<Paragraph>();
			List<Section> list2 = new List<Section>();
			foreach (Paragraph item3 in paragraphs)
			{
				XElement xElement = item3.Xml.Descendants().FirstOrDefault((XElement s) => s.Name.LocalName == "sectPr");
				if (xElement == null)
				{
					list.Add(item3);
				}
				else
				{
					list.Add(item3);
					Section item = new Section(base.Document, xElement)
					{
						SectionParagraphs = list
					};
					list2.Add(item);
					list = new List<Paragraph>();
				}
			}
			XElement xElement2 = mainDoc.Root.Element(XName.Get("body", w.NamespaceName));
			XElement xml = xElement2.Element(XName.Get("sectPr", w.NamespaceName));
			Section item2 = new Section(base.Document, xml)
			{
				SectionParagraphs = list
			};
			list2.Add(item2);
			return list2;
		}

		internal DocX(DocX document, XElement xml)
			: base(document, xml)
		{
		}

		internal string GetCollectiveText(List<PackagePart> list)
		{
			string text = string.Empty;
			foreach (PackagePart item in list)
			{
				using (TextReader textReader = new StreamReader(item.GetStream()))
				{
					XDocument xDocument = XDocument.Load(textReader);
					StringBuilder stringBuilder = new StringBuilder();
					foreach (XElement item2 in xDocument.Descendants())
					{
						switch (item2.Name.LocalName)
						{
						case "tab":
							stringBuilder.Append("\t");
							break;
						case "br":
							stringBuilder.Append("\n");
							break;
						case "t":
						case "delText":
							stringBuilder.Append(item2.Value);
							break;
						}
					}
					text = text + "\n" + stringBuilder;
				}
			}
			return text;
		}

		public void InsertDocument(DocX remote_document, bool append = true)
		{
			//IL_061d: Unknown result type (might be due to invalid IL or missing references)
			//IL_071f: Unknown result type (might be due to invalid IL or missing references)
			XDocument xDocument = new XDocument(remote_document.mainDoc);
			XDocument remote_footnotes = null;
			if (remote_document.footnotes != null)
			{
				remote_footnotes = new XDocument(remote_document.footnotes);
			}
			XDocument remote_endnotes = null;
			if (remote_document.endnotes != null)
			{
				remote_endnotes = new XDocument(remote_document.endnotes);
			}
			xDocument.Descendants(XName.Get("headerReference", w.NamespaceName)).Remove();
			xDocument.Descendants(XName.Get("footerReference", w.NamespaceName)).Remove();
			XElement xElement = xDocument.Root.Element(XName.Get("body", w.NamespaceName));
			PackagePartCollection parts = remote_document.package.GetParts();
			List<string> list = new List<string>
			{
				"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
				"application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml",
				"application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml",
				"application/vnd.openxmlformats-package.core-properties+xml",
				"application/vnd.openxmlformats-officedocument.extended-properties+xml",
				"application/vnd.openxmlformats-package.relationships+xml"
			};
			List<string> list2 = new List<string>
			{
				"image/jpeg",
				"image/jpg",
				"image/png",
				"image/bmp",
				"image/gif",
				"image/tiff",
				"image/icon",
				"image/pcx",
				"image/emf",
				"image/wmf"
			};
			foreach (PackagePart item in parts)
			{
				if (!list.Contains(item.ContentType) && !list2.Contains(item.ContentType))
				{
					if (package.PartExists(item.Uri))
					{
						PackagePart part = package.GetPart(item.Uri);
						switch (item.ContentType)
						{
						case "application/vnd.openxmlformats-officedocument.custom-properties+xml":
							merge_customs(item, part, xDocument);
							break;
						case "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml":
							merge_footnotes(item, part, xDocument, remote_document, remote_footnotes);
							remote_footnotes = footnotes;
							break;
						case "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml":
							merge_endnotes(item, part, xDocument, remote_document, remote_endnotes);
							remote_endnotes = endnotes;
							break;
						case "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml":
							merge_styles(item, part, xDocument, remote_document, remote_footnotes, remote_endnotes);
							break;
						case "application/vnd.ms-word.stylesWithEffects+xml":
							merge_styles(item, part, xDocument, remote_document, remote_footnotes, remote_endnotes);
							break;
						case "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml":
							merge_fonts(item, part, xDocument, remote_document);
							break;
						case "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml":
							merge_numbering(item, part, xDocument, remote_document);
							break;
						}
					}
					else
					{
						PackagePart val = clonePackagePart(item);
						switch (item.ContentType)
						{
						case "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml":
							endnotesPart = val;
							endnotes = remote_endnotes;
							break;
						case "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml":
							footnotesPart = val;
							footnotes = remote_footnotes;
							break;
						case "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml":
							stylesPart = val;
							using (TextReader textReader4 = new StreamReader(stylesPart.GetStream()))
							{
								styles = XDocument.Load(textReader4);
							}
							break;
						case "application/vnd.ms-word.stylesWithEffects+xml":
							stylesWithEffectsPart = val;
							using (TextReader textReader3 = new StreamReader(stylesWithEffectsPart.GetStream()))
							{
								stylesWithEffects = XDocument.Load(textReader3);
							}
							break;
						case "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml":
							fontTablePart = val;
							using (TextReader textReader2 = new StreamReader(fontTablePart.GetStream()))
							{
								fontTable = XDocument.Load(textReader2);
							}
							break;
						case "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml":
							numberingPart = val;
							using (TextReader textReader = new StreamReader(numberingPart.GetStream()))
							{
								numbering = XDocument.Load(textReader);
							}
							break;
						}
						clonePackageRelationship(remote_document, item, xDocument);
					}
				}
			}
			foreach (PackageRelationship item2 in remote_document.mainPart.GetRelationshipsByType("http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"))
			{
				string id = item2.Id;
				string id2 = mainPart.CreateRelationship(item2.TargetUri, item2.TargetMode, item2.RelationshipType).Id;
				IEnumerable<XElement> enumerable = xDocument.Descendants(XName.Get("hyperlink", w.NamespaceName));
				foreach (XElement item3 in enumerable)
				{
					XAttribute xAttribute = item3.Attribute(XName.Get("id", r.NamespaceName));
					if (xAttribute != null && xAttribute.Value == id)
					{
						xAttribute.SetValue(id2);
					}
				}
			}
			foreach (PackageRelationship item4 in remote_document.mainPart.GetRelationshipsByType("http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject"))
			{
				string id3 = item4.Id;
				string id4 = mainPart.CreateRelationship(item4.TargetUri, item4.TargetMode, item4.RelationshipType).Id;
				IEnumerable<XElement> enumerable2 = xDocument.Descendants(XName.Get("OLEObject", "urn:schemas-microsoft-com:office:office"));
				foreach (XElement item5 in enumerable2)
				{
					XAttribute xAttribute2 = item5.Attribute(XName.Get("id", r.NamespaceName));
					if (xAttribute2 != null && xAttribute2.Value == id3)
					{
						xAttribute2.SetValue(id4);
					}
				}
			}
			foreach (PackagePart item6 in parts)
			{
				if (list2.Contains(item6.ContentType))
				{
					merge_images(item6, remote_document, xDocument, item6.ContentType);
				}
			}
			int num = 0;
			IEnumerable<XElement> enumerable3 = mainDoc.Root.Descendants(XName.Get("docPr", wp.NamespaceName));
			foreach (XElement item7 in enumerable3)
			{
				XAttribute xAttribute3 = item7.Attribute(XName.Get("id"));
				int result = default(int);
				if (xAttribute3 != null && int.TryParse(xAttribute3.Value, out result) && result > num)
				{
					num = result;
				}
			}
			num++;
			IEnumerable<XElement> enumerable4 = xElement.Descendants(XName.Get("docPr", wp.NamespaceName));
			foreach (XElement item8 in enumerable4)
			{
				item8.SetAttributeValue(XName.Get("id"), num);
				num++;
			}
			XElement xElement2 = mainDoc.Root.Element(XName.Get("body", w.NamespaceName));
			if (append)
			{
				xElement2.Add(xElement.Elements());
			}
			else
			{
				xElement2.AddFirst(xElement.Elements());
			}
			foreach (XAttribute item9 in xDocument.Root.Attributes())
			{
				if (mainDoc.Root.Attribute(item9.Name) == null)
				{
					mainDoc.Root.SetAttributeValue(item9.Name, item9.Value);
				}
			}
		}

		private void merge_images(PackagePart remote_pp, DocX remote_document, XDocument remote_mainDoc, string contentType)
		{
			PackageRelationship val = (from r in (IEnumerable<PackageRelationship>)remote_document.mainPart.GetRelationships()
			where r.TargetUri.OriginalString.Equals(remote_pp.Uri.OriginalString.Replace("/word/", ""))
			select r).FirstOrDefault();
			if (val == null)
			{
				val = (from r in (IEnumerable<PackageRelationship>)remote_document.mainPart.GetRelationships()
				where r.TargetUri.OriginalString.Equals(remote_pp.Uri.OriginalString)
				select r).FirstOrDefault();
				if (val == null)
				{
					return;
				}
			}
			string id = val.Id;
			string value = ComputeMD5HashString(remote_pp.GetStream());
			IEnumerable<PackagePart> enumerable = from pp in (IEnumerable<PackagePart>)package.GetParts()
			where pp.ContentType.Equals(contentType)
			select pp;
			bool flag = false;
			foreach (PackagePart item in enumerable)
			{
				string text = ComputeMD5HashString(item.GetStream());
				if (text.Equals(value))
				{
					flag = true;
					PackageRelationship val2 = (from r in (IEnumerable<PackageRelationship>)mainPart.GetRelationships()
					where r.TargetUri.OriginalString.Equals(item.Uri.OriginalString.Replace("/word/", ""))
					select r).FirstOrDefault();
					if (val2 == null)
					{
						val2 = (from r in (IEnumerable<PackageRelationship>)mainPart.GetRelationships()
						where r.TargetUri.OriginalString.Equals(item.Uri.OriginalString)
						select r).FirstOrDefault();
					}
					if (val2 != null)
					{
						string id2 = val2.Id;
						IEnumerable<XElement> enumerable2 = remote_mainDoc.Descendants(XName.Get("blip", a.NamespaceName));
						foreach (XElement item2 in enumerable2)
						{
							XAttribute xAttribute = item2.Attribute(XName.Get("embed", r.NamespaceName));
							if (xAttribute != null && xAttribute.Value == id)
							{
								xAttribute.SetValue(id2);
							}
						}
						IEnumerable<XElement> enumerable3 = remote_mainDoc.Descendants(XName.Get("imagedata", v.NamespaceName));
						foreach (XElement item3 in enumerable3)
						{
							XAttribute xAttribute2 = item3.Attribute(XName.Get("id", r.NamespaceName));
							if (xAttribute2 != null && xAttribute2.Value == id)
							{
								xAttribute2.SetValue(id2);
							}
						}
					}
					break;
				}
			}
			if (!flag)
			{
				string originalString = remote_pp.Uri.OriginalString;
				originalString = originalString.Remove(originalString.LastIndexOf("/"));
				originalString = originalString + "/" + Guid.NewGuid() + contentType.Replace("image/", ".");
				if (!originalString.StartsWith("/"))
				{
					originalString = "/" + originalString;
				}
				PackagePart val3 = package.CreatePart(new Uri(originalString, UriKind.Relative), remote_pp.ContentType, 0);
				using (Stream input = remote_pp.GetStream())
				{
					using (Stream output = new PackagePartStream(val3.GetStream(FileMode.Create)))
					{
						CopyStream(input, output);
					}
				}
				PackageRelationship val4 = mainPart.CreateRelationship(new Uri(originalString, UriKind.Relative), 0, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image");
				string id3 = val4.Id;
				Match match = Regex.Match(id, "rId\\d+", RegexOptions.IgnoreCase);
				IEnumerable<XElement> enumerable4 = remote_mainDoc.Descendants(XName.Get("blip", a.NamespaceName));
				foreach (XElement item4 in enumerable4)
				{
					XAttribute xAttribute3 = item4.Attribute(XName.Get("embed", r.NamespaceName));
					if (xAttribute3 != null && xAttribute3.Value == id)
					{
						xAttribute3.SetValue(id3);
					}
				}
				if (!match.Success)
				{
					IEnumerable<XElement> enumerable5 = mainDoc.Descendants(XName.Get("blip", a.NamespaceName));
					foreach (XElement item5 in enumerable5)
					{
						XAttribute xAttribute4 = item5.Attribute(XName.Get("embed", r.NamespaceName));
						if (xAttribute4 != null && xAttribute4.Value == id)
						{
							xAttribute4.SetValue(id3);
						}
					}
					IEnumerable<XElement> enumerable6 = mainDoc.Descendants(XName.Get("imagedata", v.NamespaceName));
					foreach (XElement item6 in enumerable6)
					{
						XAttribute xAttribute5 = item6.Attribute(XName.Get("id", r.NamespaceName));
						if (xAttribute5 != null && xAttribute5.Value == id)
						{
							xAttribute5.SetValue(id3);
						}
					}
				}
				IEnumerable<XElement> enumerable7 = remote_mainDoc.Descendants(XName.Get("imagedata", v.NamespaceName));
				foreach (XElement item7 in enumerable7)
				{
					XAttribute xAttribute6 = item7.Attribute(XName.Get("id", r.NamespaceName));
					if (xAttribute6 != null && xAttribute6.Value == id)
					{
						xAttribute6.SetValue(id3);
					}
				}
			}
		}

		private string ComputeMD5HashString(Stream stream)
		{
			MD5 mD = MD5.Create();
			byte[] array = mD.ComputeHash(stream);
			StringBuilder stringBuilder = new StringBuilder();
			byte[] array2 = array;
			for (int i = 0; i < array2.Length; i++)
			{
				byte b = array2[i];
				stringBuilder.Append(b.ToString("X2"));
			}
			return stringBuilder.ToString();
		}

		private void merge_endnotes(PackagePart remote_pp, PackagePart local_pp, XDocument remote_mainDoc, DocX remote, XDocument remote_endnotes)
		{
			IEnumerable<int> source = from d in endnotes.Root.Descendants()
			where d.Name.LocalName == "endnote"
			select int.Parse(d.Attribute(XName.Get("id", w.NamespaceName)).Value);
			int num = source.Max() + 1;
			IEnumerable<XElement> enumerable = remote_mainDoc.Descendants(XName.Get("endnoteReference", w.NamespaceName));
			foreach (XElement item in (from fr in remote_endnotes.Root.Elements()
			orderby fr.Attribute(XName.Get("id", r.NamespaceName))
			select fr).Reverse())
			{
				XAttribute xAttribute = item.Attribute(XName.Get("id", w.NamespaceName));
				int result = default(int);
				if (xAttribute != null && int.TryParse(xAttribute.Value, out result) && result > 0)
				{
					foreach (XElement item2 in enumerable)
					{
						XAttribute xAttribute2 = item2.Attribute(XName.Get("id", w.NamespaceName));
						if (xAttribute2 != null && int.Parse(xAttribute2.Value).Equals(result))
						{
							xAttribute2.SetValue(num);
						}
					}
					item.SetAttributeValue(XName.Get("id", w.NamespaceName), num);
					endnotes.Root.Add(item);
					num++;
				}
			}
		}

		private void merge_footnotes(PackagePart remote_pp, PackagePart local_pp, XDocument remote_mainDoc, DocX remote, XDocument remote_footnotes)
		{
			IEnumerable<int> source = from d in footnotes.Root.Descendants()
			where d.Name.LocalName == "footnote"
			select int.Parse(d.Attribute(XName.Get("id", w.NamespaceName)).Value);
			int num = source.Max() + 1;
			IEnumerable<XElement> enumerable = remote_mainDoc.Descendants(XName.Get("footnoteReference", w.NamespaceName));
			foreach (XElement item in (from fr in remote_footnotes.Root.Elements()
			orderby fr.Attribute(XName.Get("id", r.NamespaceName))
			select fr).Reverse())
			{
				XAttribute xAttribute = item.Attribute(XName.Get("id", w.NamespaceName));
				int result = default(int);
				if (xAttribute != null && int.TryParse(xAttribute.Value, out result) && result > 0)
				{
					foreach (XElement item2 in enumerable)
					{
						XAttribute xAttribute2 = item2.Attribute(XName.Get("id", w.NamespaceName));
						if (xAttribute2 != null && int.Parse(xAttribute2.Value).Equals(result))
						{
							xAttribute2.SetValue(num);
						}
					}
					item.SetAttributeValue(XName.Get("id", w.NamespaceName), num);
					footnotes.Root.Add(item);
					num++;
				}
			}
		}

		private void merge_customs(PackagePart remote_pp, PackagePart local_pp, XDocument remote_mainDoc)
		{
			XDocument xDocument;
			using (TextReader textReader = new StreamReader(remote_pp.GetStream()))
			{
				xDocument = XDocument.Load(textReader);
			}
			XDocument xDocument2;
			using (TextReader textReader2 = new StreamReader(local_pp.GetStream()))
			{
				xDocument2 = XDocument.Load(textReader2);
			}
			IEnumerable<int> source = from d in xDocument.Root.Descendants()
			where d.Name.LocalName == "property"
			select int.Parse(d.Attribute(XName.Get("pid")).Value);
			int num = source.Max() + 1;
			foreach (XElement item in xDocument.Root.Elements())
			{
				bool flag = false;
				foreach (XElement item2 in xDocument2.Root.Elements())
				{
					XAttribute xAttribute = item.Attribute(XName.Get("name"));
					XAttribute xAttribute2 = item2.Attribute(XName.Get("name"));
					if (item != null && xAttribute2 != null && xAttribute.Value.Equals(xAttribute2.Value))
					{
						flag = true;
					}
				}
				if (!flag)
				{
					item.SetAttributeValue(XName.Get("pid"), num);
					xDocument2.Root.Add(item);
					num++;
				}
			}
			using (TextWriter textWriter = new StreamWriter(new PackagePartStream(local_pp.GetStream(FileMode.Create, FileAccess.Write))))
			{
				xDocument2.Save(textWriter, SaveOptions.None);
			}
		}

		private void merge_numbering(PackagePart remote_pp, PackagePart local_pp, XDocument remote_mainDoc, DocX remote)
		{
			IEnumerable<XElement> enumerable = remote.numbering.Root.Elements(XName.Get("abstractNum", w.NamespaceName));
			int num = 0;
			foreach (XElement item in enumerable)
			{
				XAttribute xAttribute = item.Attribute(XName.Get("abstractNumId", w.NamespaceName));
				int result;
				if (xAttribute != null && int.TryParse(xAttribute.Value, out result) && result > num)
				{
					num = result;
				}
			}
			num++;
			IEnumerable<XElement> enumerable2 = remote.numbering.Root.Elements(XName.Get("num", w.NamespaceName));
			int num2 = 0;
			foreach (XElement item2 in enumerable2)
			{
				XAttribute xAttribute2 = item2.Attribute(XName.Get("numId", w.NamespaceName));
				int result2;
				if (xAttribute2 != null && int.TryParse(xAttribute2.Value, out result2) && result2 > num2)
				{
					num2 = result2;
				}
			}
			num2++;
			foreach (XElement item3 in enumerable)
			{
				XAttribute xAttribute3 = item3.Attribute(XName.Get("abstractNumId", w.NamespaceName));
				if (xAttribute3 != null)
				{
					string value = xAttribute3.Value;
					xAttribute3.SetValue(num);
					foreach (XElement item4 in enumerable2)
					{
						IEnumerable<XElement> enumerable3 = remote_mainDoc.Descendants(XName.Get("numId", w.NamespaceName));
						foreach (XElement item5 in enumerable3)
						{
							XAttribute xAttribute4 = item5.Attribute(XName.Get("val", w.NamespaceName));
							if (xAttribute4?.Value.Equals(item4.Attribute(XName.Get("numId", w.NamespaceName)).Value) ?? false)
							{
								xAttribute4.SetValue(num2);
							}
						}
						item4.SetAttributeValue(XName.Get("numId", w.NamespaceName), num2);
						XAttribute xAttribute5 = item4.Element(XName.Get("abstractNumId", w.NamespaceName))?.Attribute(XName.Get("val", w.NamespaceName));
						if (xAttribute5?.Value.Equals(value) ?? false)
						{
							xAttribute5.SetValue(num);
						}
						num2++;
					}
				}
				num++;
			}
			if (numbering.Root.Elements(XName.Get("abstractNum", w.NamespaceName)).Count() > 0)
			{
				numbering.Root.Elements(XName.Get("abstractNum", w.NamespaceName)).Last().AddAfterSelf(enumerable);
			}
			if (numbering.Root.Elements(XName.Get("num", w.NamespaceName)).Count() > 0)
			{
				numbering.Root.Elements(XName.Get("num", w.NamespaceName)).Last().AddAfterSelf(enumerable2);
			}
		}

		private void merge_fonts(PackagePart remote_pp, PackagePart local_pp, XDocument remote_mainDoc, DocX remote)
		{
			IEnumerable<XElement> enumerable = remote.fontTable.Root.Elements(XName.Get("font", w.NamespaceName));
			IEnumerable<XElement> enumerable2 = fontTable.Root.Elements(XName.Get("font", w.NamespaceName));
			foreach (XElement item in enumerable)
			{
				bool flag = true;
				foreach (XElement item2 in enumerable2)
				{
					if (item2.Attribute(XName.Get("name", w.NamespaceName)).Value == item.Attribute(XName.Get("name", w.NamespaceName)).Value)
					{
						flag = false;
						break;
					}
				}
				if (flag)
				{
					fontTable.Root.Add(item);
				}
			}
		}

		private void merge_styles(PackagePart remote_pp, PackagePart local_pp, XDocument remote_mainDoc, DocX remote, XDocument remote_footnotes, XDocument remote_endnotes)
		{
			Dictionary<string, string> dictionary = new Dictionary<string, string>();
			foreach (XElement item in styles.Root.Elements(XName.Get("style", w.NamespaceName)))
			{
				XElement xElement = new XElement(item);
				XAttribute xAttribute = xElement.Attribute(XName.Get("styleId", w.NamespaceName));
				string value = xAttribute.Value;
				xAttribute.Remove();
				string key = Regex.Replace(xElement.ToString(), "\\s+", "");
				if (!dictionary.ContainsKey(key))
				{
					dictionary.Add(key, value);
				}
			}
			IEnumerable<XElement> enumerable = remote.styles.Root.Elements(XName.Get("style", w.NamespaceName));
			foreach (XElement item2 in enumerable)
			{
				XElement xElement2 = new XElement(item2);
				XAttribute xAttribute2 = xElement2.Attribute(XName.Get("styleId", w.NamespaceName));
				string value2 = xAttribute2.Value;
				xAttribute2.Remove();
				string key2 = Regex.Replace(xElement2.ToString(), "\\s+", "");
				string value4;
				if (dictionary.ContainsKey(key2))
				{
					dictionary.TryGetValue(key2, out string value3);
					if (value3 == value2)
					{
						continue;
					}
					value4 = value3;
				}
				else
				{
					value4 = Guid.NewGuid().ToString();
					item2.SetAttributeValue(XName.Get("styleId", w.NamespaceName), value4);
				}
				foreach (XElement item3 in remote_mainDoc.Root.Descendants(XName.Get("pStyle", w.NamespaceName)))
				{
					XAttribute xAttribute3 = item3.Attribute(XName.Get("val", w.NamespaceName));
					if (xAttribute3?.Value.Equals(xAttribute2.Value) ?? false)
					{
						xAttribute3.SetValue(value4);
					}
				}
				foreach (XElement item4 in remote_mainDoc.Root.Descendants(XName.Get("rStyle", w.NamespaceName)))
				{
					XAttribute xAttribute4 = item4.Attribute(XName.Get("val", w.NamespaceName));
					if (xAttribute4?.Value.Equals(xAttribute2.Value) ?? false)
					{
						xAttribute4.SetValue(value4);
					}
				}
				foreach (XElement item5 in remote_mainDoc.Root.Descendants(XName.Get("tblStyle", w.NamespaceName)))
				{
					XAttribute xAttribute5 = item5.Attribute(XName.Get("val", w.NamespaceName));
					if (xAttribute5?.Value.Equals(xAttribute2.Value) ?? false)
					{
						xAttribute5.SetValue(value4);
					}
				}
				if (remote_endnotes != null)
				{
					foreach (XElement item6 in remote_endnotes.Root.Descendants(XName.Get("rStyle", w.NamespaceName)))
					{
						XAttribute xAttribute6 = item6.Attribute(XName.Get("val", w.NamespaceName));
						if (xAttribute6?.Value.Equals(xAttribute2.Value) ?? false)
						{
							xAttribute6.SetValue(value4);
						}
					}
					foreach (XElement item7 in remote_endnotes.Root.Descendants(XName.Get("pStyle", w.NamespaceName)))
					{
						XAttribute xAttribute7 = item7.Attribute(XName.Get("val", w.NamespaceName));
						if (xAttribute7?.Value.Equals(xAttribute2.Value) ?? false)
						{
							xAttribute7.SetValue(value4);
						}
					}
				}
				if (remote_footnotes != null)
				{
					foreach (XElement item8 in remote_footnotes.Root.Descendants(XName.Get("rStyle", w.NamespaceName)))
					{
						XAttribute xAttribute8 = item8.Attribute(XName.Get("val", w.NamespaceName));
						if (xAttribute8?.Value.Equals(xAttribute2.Value) ?? false)
						{
							xAttribute8.SetValue(value4);
						}
					}
					foreach (XElement item9 in remote_footnotes.Root.Descendants(XName.Get("pStyle", w.NamespaceName)))
					{
						XAttribute xAttribute9 = item9.Attribute(XName.Get("val", w.NamespaceName));
						if (xAttribute9?.Value.Equals(xAttribute2.Value) ?? false)
						{
							xAttribute9.SetValue(value4);
						}
					}
				}
				xAttribute2.SetValue(value4);
				styles.Root.Add(item2);
			}
		}

		protected void clonePackageRelationship(DocX remote_document, PackagePart pp, XDocument remote_mainDoc)
		{
			//IL_0086: Unknown result type (might be due to invalid IL or missing references)
			string text = pp.Uri.OriginalString.Replace("/", "");
			PackageRelationshipCollection relationships = remote_document.mainPart.GetRelationships();
			foreach (PackageRelationship item in relationships)
			{
				if (text.Equals("word" + item.TargetUri.OriginalString.Replace("/", "")))
				{
					string id = item.Id;
					string id2 = mainPart.CreateRelationship(item.TargetUri, item.TargetMode, item.RelationshipType).Id;
					IEnumerable<XElement> enumerable = remote_mainDoc.Descendants(XName.Get("blip", a.NamespaceName));
					foreach (XElement item2 in enumerable)
					{
						XAttribute xAttribute = item2.Attribute(XName.Get("embed", r.NamespaceName));
						if (xAttribute != null && xAttribute.Value == id)
						{
							xAttribute.SetValue(id2);
						}
					}
					IEnumerable<XElement> enumerable2 = remote_mainDoc.Descendants(XName.Get("imagedata", v.NamespaceName));
					foreach (XElement item3 in enumerable2)
					{
						XAttribute xAttribute2 = item3.Attribute(XName.Get("id", r.NamespaceName));
						if (xAttribute2 != null && xAttribute2.Value == id)
						{
							xAttribute2.SetValue(id2);
						}
					}
					break;
				}
			}
		}

		protected PackagePart clonePackagePart(PackagePart pp)
		{
			PackagePart val = package.CreatePart(pp.Uri, pp.ContentType, 0);
			using (Stream input = pp.GetStream())
			{
				using (Stream output = new PackagePartStream(val.GetStream(FileMode.Create)))
				{
					CopyStream(input, output);
				}
			}
			return val;
		}

		protected string GetMD5HashFromStream(Stream stream)
		{
			MD5 mD = MD5.Create();
			byte[] array = mD.ComputeHash(stream);
			StringBuilder stringBuilder = new StringBuilder();
			for (int i = 0; i < array.Length; i++)
			{
				stringBuilder.Append(array[i].ToString("x2"));
			}
			return stringBuilder.ToString();
		}

		public new Table InsertTable(int rowCount, int columnCount)
		{
			if (rowCount < 1 || columnCount < 1)
			{
				throw new ArgumentOutOfRangeException("Row and Column count must be greater than zero.");
			}
			Table table = base.InsertTable(rowCount, columnCount);
			table.mainPart = mainPart;
			return table;
		}

		public Table AddTable(int rowCount, int columnCount)
		{
			if (rowCount < 1 || columnCount < 1)
			{
				throw new ArgumentOutOfRangeException("Row and Column count must be greater than zero.");
			}
			Table table = new Table(this, HelperFunctions.CreateTable(rowCount, columnCount));
			table.mainPart = mainPart;
			return table;
		}

		public List AddList(string listText = null, int level = 0, ListItemType listType = ListItemType.Numbered, int? startNumber = default(int?), bool trackChanges = false, bool continueNumbering = false)
		{
			return AddListItem(new List(this, null), listText, level, listType, startNumber, trackChanges, continueNumbering);
		}

		public List AddListItem(List list, string listText, int level = 0, ListItemType listType = ListItemType.Numbered, int? startNumber = default(int?), bool trackChanges = false, bool continueNumbering = false)
		{
			if (startNumber.HasValue && continueNumbering)
			{
				throw new InvalidOperationException("Cannot specify a start number and at the same time continue numbering from another list");
			}
			List list2 = HelperFunctions.CreateItemInList(list, listText, level, listType, startNumber, trackChanges, continueNumbering);
			Paragraph paragraph = list2.Items.LastOrDefault();
			if (paragraph != null)
			{
				paragraph.PackagePart = mainPart;
			}
			return list2;
		}

		public new List InsertList(List list)
		{
			base.InsertList(list);
			return list;
		}

		public new List InsertList(List list, Font fontFamily, double fontSize)
		{
			base.InsertList(list, fontFamily, fontSize);
			return list;
		}

		public new List InsertList(List list, double fontSize)
		{
			base.InsertList(list, fontSize);
			return list;
		}

		public new List InsertList(int index, List list)
		{
			base.InsertList(index, list);
			return list;
		}

		internal XDocument AddStylesForList()
		{
			Uri uri = new Uri("/word/styles.xml", UriKind.Relative);
			if (!package.PartExists(uri))
			{
				HelperFunctions.AddDefaultStylesXml(package);
			}
			XDocument xDocument;
			using (TextReader textReader = new StreamReader(package.GetPart(uri).GetStream()))
			{
				xDocument = XDocument.Load(textReader);
			}
			if (!(from s in xDocument.Element(w + "styles").Elements()
			let styleId = s.Attribute(XName.Get("styleId", w.NamespaceName))
			where styleId != null && styleId.Value == "ListParagraph"
			select s).Any())
			{
				XElement content = new XElement(w + "style", new XAttribute(w + "type", "paragraph"), new XAttribute(w + "styleId", "ListParagraph"), new XElement(w + "name", new XAttribute(w + "val", "List Paragraph")), new XElement(w + "basedOn", new XAttribute(w + "val", "Normal")), new XElement(w + "uiPriority", new XAttribute(w + "val", "34")), new XElement(w + "qformat"), new XElement(w + "rsid", new XAttribute(w + "val", "00832EE1")), new XElement(w + "rPr", new XElement(w + "ind", new XAttribute(w + "left", "720")), new XElement(w + "contextualSpacing")));
				xDocument.Element(w + "styles").Add(content);
				using (TextWriter textWriter = new StreamWriter(new PackagePartStream(package.GetPart(uri).GetStream())))
				{
					xDocument.Save(textWriter);
				}
			}
			return xDocument;
		}

		public new Table InsertTable(int index, Table t)
		{
			Table table = base.InsertTable(index, t);
			table.mainPart = mainPart;
			return table;
		}

		public new Table InsertTable(Table t)
		{
			t = base.InsertTable(t);
			t.mainPart = mainPart;
			return t;
		}

		public new Table InsertTable(int index, int rowCount, int columnCount)
		{
			if (rowCount < 1 || columnCount < 1)
			{
				throw new ArgumentOutOfRangeException("Row and Column count must be greater than zero.");
			}
			Table table = base.InsertTable(index, rowCount, columnCount);
			table.mainPart = mainPart;
			return table;
		}

		public static DocX Create(Stream stream, DocumentTypes documentType = DocumentTypes.Document)
		{
			MemoryStream memoryStream = new MemoryStream();
			Package val = Package.Open((Stream)memoryStream, FileMode.Create, FileAccess.ReadWrite);
			PostCreation(val, documentType);
			DocX docX = Load(memoryStream);
			docX.stream = stream;
			return docX;
		}

		public static DocX Create(string filename, DocumentTypes documentType = DocumentTypes.Document)
		{
			MemoryStream memoryStream = new MemoryStream();
			Package val = Package.Open((Stream)memoryStream, FileMode.Create, FileAccess.ReadWrite);
			PostCreation(val, documentType);
			DocX docX = Load(memoryStream);
			docX.filename = filename;
			return docX;
		}

		internal static void PostCreation(Package package, DocumentTypes documentType = DocumentTypes.Document)
		{
			PackagePart val = (documentType != 0) ? package.CreatePart(new Uri("/word/document.xml", UriKind.Relative), "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml", 0) : package.CreatePart(new Uri("/word/document.xml", UriKind.Relative), "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml", 0);
			package.CreateRelationship(val.Uri, 0, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument");
			XDocument xDocument;
			using (new StreamReader(val.GetStream(FileMode.Create, FileAccess.ReadWrite)))
			{
				xDocument = XDocument.Parse("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n                   <w:document xmlns:ve=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\" xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" xmlns:w10=\"urn:schemas-microsoft-com:office:word\" xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:wne=\"http://schemas.microsoft.com/office/word/2006/wordml\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\">\r\n                   <w:body>\r\n                    <w:sectPr w:rsidR=\"003E25F4\" w:rsidSect=\"00FC3028\">\r\n                        <w:pgSz w:w=\"11906\" w:h=\"16838\"/>\r\n                        <w:pgMar w:top=\"1440\" w:right=\"1440\" w:bottom=\"1440\" w:left=\"1440\" w:header=\"708\" w:footer=\"708\" w:gutter=\"0\"/>\r\n                        <w:cols w:space=\"708\"/>\r\n                        <w:docGrid w:linePitch=\"360\"/>\r\n                    </w:sectPr>\r\n                   </w:body>\r\n                   </w:document>");
			}
			using (TextWriter textWriter = new StreamWriter(new PackagePartStream(val.GetStream(FileMode.Create, FileAccess.Write))))
			{
				xDocument.Save(textWriter, SaveOptions.None);
			}
			XDocument xDocument2 = HelperFunctions.AddDefaultStylesXml(package);
			XDocument xDocument3 = HelperFunctions.AddDefaultNumberingXml(package);
			package.Close();
		}

		internal static DocX PostLoad(ref Package package)
		{
			DocX docX = new DocX(null, null);
			docX.package = package;
			docX.Document = docX;
			docX.mainPart = (from p in (IEnumerable<PackagePart>)package.GetParts()
			where p.ContentType.Equals("application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml", StringComparison.CurrentCultureIgnoreCase) || p.ContentType.Equals("application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml", StringComparison.CurrentCultureIgnoreCase)
			select p).Single();
			using (TextReader textReader = new StreamReader(docX.mainPart.GetStream(FileMode.Open, FileAccess.Read)))
			{
				docX.mainDoc = XDocument.Load(textReader, LoadOptions.PreserveWhitespace);
			}
			PopulateDocument(docX, package);
			using (TextReader textReader2 = new StreamReader(docX.settingsPart.GetStream()))
			{
				docX.settings = XDocument.Load(textReader2);
			}
			docX.paragraphLookup.Clear();
			foreach (Paragraph paragraph in docX.Paragraphs)
			{
				if (!docX.paragraphLookup.ContainsKey(paragraph.endIndex))
				{
					docX.paragraphLookup.Add(paragraph.endIndex, paragraph);
				}
			}
			return docX;
		}

		private static void PopulateDocument(DocX document, Package package)
		{
			Headers headers = new Headers();
			headers.odd = document.GetHeaderByType("default");
			headers.even = document.GetHeaderByType("even");
			headers.first = document.GetHeaderByType("first");
			Footers footers = new Footers();
			footers.odd = document.GetFooterByType("default");
			footers.even = document.GetFooterByType("even");
			footers.first = document.GetFooterByType("first");
			document.Xml = document.mainDoc.Root.Element(w + "body");
			document.headers = headers;
			document.footers = footers;
			document.settingsPart = HelperFunctions.CreateOrGetSettingsPart(package);
			PackagePartCollection parts = package.GetParts();
			foreach (PackageRelationship relationship in document.mainPart.GetRelationships())
			{
				string uriString = "/word/" + relationship.TargetUri.OriginalString.Replace("/word/", "").Replace("file://", "");
				switch (relationship.RelationshipType)
				{
				case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes":
					document.endnotesPart = package.GetPart(new Uri(uriString, UriKind.Relative));
					using (TextReader textReader6 = new StreamReader(document.endnotesPart.GetStream()))
					{
						document.endnotes = XDocument.Load(textReader6);
					}
					break;
				case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes":
					document.footnotesPart = package.GetPart(new Uri(uriString, UriKind.Relative));
					using (TextReader textReader5 = new StreamReader(document.footnotesPart.GetStream()))
					{
						document.footnotes = XDocument.Load(textReader5);
					}
					break;
				case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles":
					document.stylesPart = package.GetPart(new Uri(uriString, UriKind.Relative));
					using (TextReader textReader4 = new StreamReader(document.stylesPart.GetStream()))
					{
						document.styles = XDocument.Load(textReader4);
					}
					break;
				case "http://schemas.microsoft.com/office/2007/relationships/stylesWithEffects":
					document.stylesWithEffectsPart = package.GetPart(new Uri(uriString, UriKind.Relative));
					using (TextReader textReader3 = new StreamReader(document.stylesWithEffectsPart.GetStream()))
					{
						document.stylesWithEffects = XDocument.Load(textReader3);
					}
					break;
				case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable":
					document.fontTablePart = package.GetPart(new Uri(uriString, UriKind.Relative));
					using (TextReader textReader2 = new StreamReader(document.fontTablePart.GetStream()))
					{
						document.fontTable = XDocument.Load(textReader2);
					}
					break;
				case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering":
					document.numberingPart = package.GetPart(new Uri(uriString, UriKind.Relative));
					using (TextReader textReader = new StreamReader(document.numberingPart.GetStream()))
					{
						document.numbering = XDocument.Load(textReader);
					}
					break;
				}
			}
		}

		public DocX Copy()
		{
			MemoryStream memoryStream = new MemoryStream();
			SaveAs(memoryStream);
			memoryStream.Seek(0L, SeekOrigin.Begin);
			return Load(memoryStream);
		}

		public static DocX Load(Stream stream)
		{
			MemoryStream memoryStream = new MemoryStream();
			stream.Position = 0L;
			byte[] buffer = new byte[stream.Length];
			stream.Read(buffer, 0, (int)stream.Length);
			memoryStream.Write(buffer, 0, (int)stream.Length);
			Package val = Package.Open((Stream)memoryStream, FileMode.Open, FileAccess.ReadWrite);
			DocX docX = PostLoad(ref val);
			docX.package = val;
			docX.memoryStream = memoryStream;
			docX.stream = stream;
			return docX;
		}

		public static DocX Load(string filename)
		{
			if (!File.Exists(filename))
			{
				throw new FileNotFoundException($"File could not be found {filename}");
			}
			MemoryStream memoryStream = new MemoryStream();
			using (FileStream input = new FileStream(filename, FileMode.Open, FileAccess.Read, FileShare.Read))
			{
				CopyStream(input, memoryStream);
			}
			Package val = Package.Open((Stream)memoryStream, FileMode.Open, FileAccess.ReadWrite);
			DocX docX = PostLoad(ref val);
			docX.package = val;
			docX.filename = filename;
			docX.memoryStream = memoryStream;
			return docX;
		}

		public void ApplyTemplate(string templateFilePath)
		{
			ApplyTemplate(templateFilePath, includeContent: true);
		}

		public void ApplyTemplate(string templateFilePath, bool includeContent)
		{
			if (!File.Exists(templateFilePath))
			{
				throw new FileNotFoundException($"File could not be found {templateFilePath}");
			}
			using (FileStream templateStream = new FileStream(templateFilePath, FileMode.Open, FileAccess.Read))
			{
				ApplyTemplate(templateStream, includeContent);
			}
		}

		public void ApplyTemplate(Stream templateStream)
		{
			ApplyTemplate(templateStream, includeContent: true);
		}

		public void ApplyTemplate(Stream templateStream, bool includeContent)
		{
			//IL_00d5: Unknown result type (might be due to invalid IL or missing references)
			//IL_01a1: Unknown result type (might be due to invalid IL or missing references)
			//IL_02f1: Unknown result type (might be due to invalid IL or missing references)
			//IL_0322: Unknown result type (might be due to invalid IL or missing references)
			Package val = Package.Open(templateStream, FileMode.Open, FileAccess.Read);
			try
			{
				PackagePart val2 = null;
				XDocument xDocument = null;
				foreach (PackagePart part4 in val.GetParts())
				{
					switch (part4.Uri.ToString())
					{
					case "/word/document.xml":
						val2 = part4;
						using (XmlReader xmlReader = XmlReader.Create(part4.GetStream(FileMode.Open, FileAccess.Read)))
						{
							xDocument = XDocument.Load(xmlReader);
							xmlReader.Dispose();
						}
						break;
					case "/_rels/.rels":
					{
						if (!package.PartExists(part4.Uri))
						{
							package.CreatePart(part4.Uri, part4.ContentType, part4.CompressionOption);
						}
						PackagePart part2 = package.GetPart(part4.Uri);
						using (StreamReader streamReader2 = new StreamReader(part4.GetStream(FileMode.Open, FileAccess.Read), Encoding.UTF8))
						{
							using (StreamWriter streamWriter2 = new StreamWriter(new PackagePartStream(part2.GetStream(FileMode.Create, FileAccess.Write)), Encoding.UTF8))
							{
								streamWriter2.Write(streamReader2.ReadToEnd());
								streamReader2.Dispose();
								streamWriter2.Dispose();
							}
						}
						break;
					}
					default:
					{
						if (!package.PartExists(part4.Uri))
						{
							package.CreatePart(part4.Uri, part4.ContentType, part4.CompressionOption);
						}
						Encoding uTF = Encoding.UTF8;
						if (part4.Uri.ToString().EndsWith(".xml") || part4.Uri.ToString().EndsWith(".rels"))
						{
							uTF = Encoding.UTF8;
						}
						PackagePart part = package.GetPart(part4.Uri);
						using (StreamReader streamReader = new StreamReader(part4.GetStream(FileMode.Open, FileAccess.Read), uTF))
						{
							using (StreamWriter streamWriter = new StreamWriter(new PackagePartStream(part.GetStream(FileMode.Create, FileAccess.Write)), streamReader.CurrentEncoding))
							{
								streamWriter.Write(streamReader.ReadToEnd());
								streamReader.Dispose();
								streamWriter.Dispose();
							}
						}
						break;
					}
					case "/word/_rels/document.xml.rels":
						break;
					}
				}
				if (val2 != null)
				{
					string text = val2.ContentType.Replace("template.main", "document.main");
					if (package.PartExists(val2.Uri))
					{
						package.DeletePart(val2.Uri);
					}
					PackagePart val3 = package.CreatePart(val2.Uri, text, val2.CompressionOption);
					foreach (PackageRelationship relationship in val2.GetRelationships())
					{
						val3.CreateRelationship(relationship.TargetUri, relationship.TargetMode, relationship.RelationshipType, relationship.Id);
					}
					mainPart = val3;
					mainDoc = xDocument;
					PopulateDocument(this, val);
					settingsPart = HelperFunctions.CreateOrGetSettingsPart(package);
				}
				if (!includeContent)
				{
					foreach (Paragraph paragraph in Paragraphs)
					{
						paragraph.Remove(trackChanges: false);
					}
				}
			}
			finally
			{
				package.Flush();
				PackagePart part3 = package.GetPart(new Uri("/word/_rels/document.xml.rels", UriKind.Relative));
				val.Close();
				PopulateDocument(base.Document, package);
			}
		}

		public Image AddImage(string filename, string contentType = null)
		{
			if (string.IsNullOrEmpty(contentType))
			{
				switch (Path.GetExtension(filename))
				{
				case ".tiff":
					contentType = "image/tif";
					break;
				case ".tif":
					contentType = "image/tif";
					break;
				case ".png":
					contentType = "image/png";
					break;
				case ".bmp":
					contentType = "image/png";
					break;
				case ".gif":
					contentType = "image/gif";
					break;
				case ".jpg":
					contentType = "image/jpg";
					break;
				case ".jpeg":
					contentType = "image/jpeg";
					break;
				default:
					contentType = "image/jpg";
					break;
				}
			}
			return AddImage(File.Open(filename, FileMode.Open), contentType);
		}

		public Image AddImage(Stream stream, string contentType = "image/jpeg")
		{
			return AddImage((object)stream, contentType);
		}

		public Hyperlink AddHyperlink(string text, Uri uri)
		{
			XElement i = new XElement(XName.Get("hyperlink", w.NamespaceName), new XAttribute(r + "id", string.Empty), new XAttribute(w + "history", "1"), new XElement(XName.Get("r", w.NamespaceName), new XElement(XName.Get("rPr", w.NamespaceName), new XElement(XName.Get("rStyle", w.NamespaceName), new XAttribute(w + "val", "Hyperlink"))), new XElement(XName.Get("t", w.NamespaceName), text)));
			Hyperlink hyperlink = new Hyperlink(this, mainPart, i);
			hyperlink.text = text;
			hyperlink.uri = uri;
			AddHyperlinkStyleIfNotPresent();
			return hyperlink;
		}

		internal void AddHyperlinkStyleIfNotPresent()
		{
			Uri uri = new Uri("/word/styles.xml", UriKind.Relative);
			if (!package.PartExists(uri))
			{
				HelperFunctions.AddDefaultStylesXml(package);
			}
			XDocument xDocument;
			using (TextReader textReader = new StreamReader(package.GetPart(uri).GetStream()))
			{
				xDocument = XDocument.Load(textReader);
			}
			if ((from s in xDocument.Element(w + "styles").Elements()
			let styleId = s.Attribute(XName.Get("styleId", w.NamespaceName))
			where styleId != null && styleId.Value == "Hyperlink"
			select s).Count() <= 0)
			{
				XElement content = new XElement(w + "style", new XAttribute(w + "type", "character"), new XAttribute(w + "styleId", "Hyperlink"), new XElement(w + "name", new XAttribute(w + "val", "Hyperlink")), new XElement(w + "basedOn", new XAttribute(w + "val", "DefaultParagraphFont")), new XElement(w + "uiPriority", new XAttribute(w + "val", "99")), new XElement(w + "unhideWhenUsed"), new XElement(w + "rsid", new XAttribute(w + "val", "0005416C")), new XElement(w + "rPr", new XElement(w + "color", new XAttribute(w + "val", "0000FF"), new XAttribute(w + "themeColor", "hyperlink")), new XElement(w + "u", new XAttribute(w + "val", "single"))));
				xDocument.Element(w + "styles").Add(content);
				using (TextWriter textWriter = new StreamWriter(new PackagePartStream(package.GetPart(uri).GetStream())))
				{
					xDocument.Save(textWriter);
				}
			}
		}

		private string GetNextFreeRelationshipID()
		{
			string s = (from r in (IEnumerable<PackageRelationship>)mainPart.GetRelationships()
			where r.Id.Substring(0, 3).Equals("rId")
			select int.Parse(r.Id.Substring(3))).DefaultIfEmpty().Max().ToString();
			if (int.TryParse(s, out int result))
			{
				return "rId" + (result + 1);
			}
			string empty = string.Empty;
			do
			{
				empty = Guid.NewGuid().ToString();
			}
			while (char.IsDigit(empty[0]));
			return empty;
		}

		public void AddHeaders()
		{
			AddHeadersOrFooters(b: true);
			headers.odd = base.Document.GetHeaderByType("default");
			headers.even = base.Document.GetHeaderByType("even");
			headers.first = base.Document.GetHeaderByType("first");
		}

		public void AddFooters()
		{
			AddHeadersOrFooters(b: false);
			footers.odd = base.Document.GetFooterByType("default");
			footers.even = base.Document.GetFooterByType("even");
			footers.first = base.Document.GetFooterByType("first");
		}

		internal void AddHeadersOrFooters(bool b)
		{
			string arg = "ftr";
			string text = "footer";
			if (b)
			{
				arg = "hdr";
				text = "header";
			}
			DeleteHeadersOrFooters(b);
			XElement xElement = mainDoc.Root.Element(w + "body").Element(w + "sectPr");
			for (int i = 1; i < 4; i++)
			{
				string uriString = $"/word/{text}{i}.xml";
				PackagePart val = package.CreatePart(new Uri(uriString, UriKind.Relative), $"application/vnd.openxmlformats-officedocument.wordprocessingml.{text}+xml", 0);
				PackageRelationship val2 = mainPart.CreateRelationship(val.Uri, 0, $"http://schemas.openxmlformats.org/officeDocument/2006/relationships/{text}");
				XDocument xDocument;
				using (new StreamReader(val.GetStream(FileMode.Create, FileAccess.ReadWrite)))
				{
					xDocument = XDocument.Parse(string.Format("<?xml version=\"1.0\" encoding=\"utf-16\" standalone=\"yes\"?>\r\n                       <w:{0} xmlns:ve=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\" xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" xmlns:w10=\"urn:schemas-microsoft-com:office:word\" xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:wne=\"http://schemas.microsoft.com/office/word/2006/wordml\">\r\n                         <w:p w:rsidR=\"009D472B\" w:rsidRDefault=\"009D472B\">\r\n                           <w:pPr>\r\n                             <w:pStyle w:val=\"{1}\" />\r\n                           </w:pPr>\r\n                         </w:p>\r\n                       </w:{0}>", arg, text));
				}
				using (TextWriter textWriter = new StreamWriter(new PackagePartStream(val.GetStream(FileMode.Create, FileAccess.Write))))
				{
					xDocument.Save(textWriter, SaveOptions.None);
				}
				string value;
				switch (i)
				{
				case 1:
					value = "default";
					break;
				case 2:
					value = "even";
					break;
				case 3:
					value = "first";
					break;
				default:
					throw new ArgumentOutOfRangeException();
				}
				xElement.Add(new XElement(w + $"{text}Reference", new XAttribute(w + "type", value), new XAttribute(r + "id", val2.Id)));
			}
		}

		internal void DeleteHeadersOrFooters(bool b)
		{
			string reference = "footer";
			if (b)
			{
				reference = "header";
			}
			PackageRelationshipCollection relationshipsByType = mainPart.GetRelationshipsByType($"http://schemas.openxmlformats.org/officeDocument/2006/relationships/{reference}");
			foreach (PackageRelationship item in relationshipsByType)
			{
				Uri uri = item.TargetUri;
				if (!uri.OriginalString.StartsWith("/word/"))
				{
					uri = new Uri("/word/" + uri.OriginalString, UriKind.Relative);
				}
				if (package.PartExists(uri))
				{
					package.DeletePart(uri);
					IEnumerable<XElement> source = from e in mainDoc.Descendants(XName.Get("body", w.NamespaceName)).Descendants()
					where e.Name.LocalName == $"{reference}Reference" && e.Attribute(r + "id").Value == item.Id
					select e;
					for (int i = 0; i < source.Count(); i++)
					{
						source.ElementAt(i).Remove();
					}
					package.DeleteRelationship(item.Id);
				}
			}
		}

		internal Image AddImage(object o, string contentType = "image/jpeg")
		{
			Stream stream = (!(o is string)) ? (o as Stream) : new FileStream(o as string, FileMode.Open, FileAccess.Read);
			PackagePartCollection parts = package.GetParts();
			var source = (from x in (IEnumerable<PackagePart>)parts
			select new
			{
				UriString = x.Uri.ToString(),
				Part = x
			}).ToList();
			Dictionary<string, PackagePart> dictionary = source.ToDictionary(x => x.UriString, x => x.Part, StringComparer.Ordinal);
			List<PackagePart> list = new List<PackagePart>();
			foreach (PackageRelationship item in mainPart.GetRelationshipsByType("http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"))
			{
				string key = item.TargetUri.ToString();
				if (dictionary.TryGetValue(key, out PackagePart value))
				{
					list.Add(value);
				}
			}
			IEnumerable<PackagePart> enumerable = from part in source
			where part.Part.ContentType.Equals("application/vnd.openxmlformats-package.relationships+xml", StringComparison.Ordinal) && part.UriString.IndexOf("/word/", StringComparison.Ordinal) > -1
			select part.Part;
			XName name = XName.Get("Target");
			XName name2 = XName.Get("TargetMode");
			foreach (PackagePart item2 in enumerable)
			{
				XDocument xDocument;
				using (TextReader textReader = new StreamReader(item2.GetStream(FileMode.Open, FileAccess.Read)))
				{
					xDocument = XDocument.Load(textReader);
				}
				IEnumerable<XElement> enumerable2 = from imageRel in xDocument.Root.Elements()
				where imageRel.Attribute(XName.Get("Type")).Value.Equals("http://schemas.openxmlformats.org/officeDocument/2006/relationships/image")
				select imageRel;
				foreach (XElement item3 in enumerable2)
				{
					XAttribute xAttribute = item3.Attribute(name);
					if (xAttribute != null)
					{
						string text = string.Empty;
						XAttribute xAttribute2 = item3.Attribute(name2);
						if (xAttribute2 != null)
						{
							text = xAttribute2.Value;
						}
						if (!text.Equals("External"))
						{
							string text2 = Path.Combine(Path.GetDirectoryName(item2.Uri.ToString()), xAttribute.Value);
							text2 = Path.GetFullPath(text2.Replace("\\_rels", string.Empty));
							text2 = text2.Replace(Path.GetFullPath("\\"), string.Empty).Replace("\\", "/");
							if (!text2.StartsWith("/"))
							{
								text2 = "/" + text2;
							}
							PackagePart part2 = package.GetPart(new Uri(text2, UriKind.Relative));
							list.Add(part2);
						}
					}
				}
			}
			foreach (PackagePart item4 in list)
			{
				using (Stream streamOne = item4.GetStream(FileMode.Open, FileAccess.Read))
				{
					if (HelperFunctions.IsSameFile(streamOne, stream))
					{
						PackageRelationship pr = ((IEnumerable<PackageRelationship>)mainPart.GetRelationshipsByType("http://schemas.openxmlformats.org/officeDocument/2006/relationships/image")).First((PackageRelationship x) => x.TargetUri == item4.Uri);
						return new Image(this, pr);
					}
				}
			}
			string empty = string.Empty;
			string arg = contentType.Substring(contentType.LastIndexOf("/") + 1);
			do
			{
				empty = $"/word/media/{Guid.NewGuid()}.{arg}";
			}
			while (package.PartExists(new Uri(empty, UriKind.Relative)));
			PackagePart val = package.CreatePart(new Uri(empty, UriKind.Relative), contentType, 0);
			PackageRelationship pr2 = mainPart.CreateRelationship(val.Uri, 0, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image");
			using (Stream output = new PackagePartStream(val.GetStream(FileMode.Create, FileAccess.Write)))
			{
				using (stream)
				{
					CopyStream(stream, output, 4096);
				}
			}
			return new Image(this, pr2);
		}

		public void Save()
		{
			Headers headers = Headers;
			using (TextWriter textWriter = new StreamWriter(new PackagePartStream(mainPart.GetStream(FileMode.Create, FileAccess.Write))))
			{
				mainDoc.Save(textWriter, SaveOptions.None);
			}
			if (settings == null)
			{
				using (TextReader textReader = new StreamReader(settingsPart.GetStream()))
				{
					settings = XDocument.Load(textReader);
				}
			}
			XElement xElement = mainDoc.Root.Element(w + "body");
			XElement xElement2 = xElement.Descendants(w + "sectPr").FirstOrDefault();
			if (xElement2 != null)
			{
				string text = (from e in mainDoc.Descendants(w + "headerReference")
				let type = e.Attribute(w + "type")
				where type != null && type.Value.Equals("even", StringComparison.CurrentCultureIgnoreCase)
				select e.Attribute(r + "id").Value).LastOrDefault();
				if (text != null)
				{
					XElement xml = headers.even.Xml;
					Uri uri = PackUriHelper.ResolvePartUri(mainPart.Uri, mainPart.GetRelationship(text).TargetUri);
					using (TextWriter textWriter2 = new StreamWriter(new PackagePartStream(package.GetPart(uri).GetStream(FileMode.Create, FileAccess.Write))))
					{
						new XDocument(new XDeclaration("1.0", "UTF-8", "yes"), xml).Save(textWriter2, SaveOptions.None);
					}
				}
				string text2 = (from e in mainDoc.Descendants(w + "headerReference")
				let type = e.Attribute(w + "type")
				where type != null && type.Value.Equals("default", StringComparison.CurrentCultureIgnoreCase)
				select e.Attribute(r + "id").Value).LastOrDefault();
				if (text2 != null)
				{
					XElement xml2 = headers.odd.Xml;
					Uri uri2 = PackUriHelper.ResolvePartUri(mainPart.Uri, mainPart.GetRelationship(text2).TargetUri);
					using (TextWriter textWriter3 = new StreamWriter(new PackagePartStream(package.GetPart(uri2).GetStream(FileMode.Create, FileAccess.Write))))
					{
						new XDocument(new XDeclaration("1.0", "UTF-8", "yes"), xml2).Save(textWriter3, SaveOptions.None);
					}
				}
				string text3 = (from e in mainDoc.Descendants(w + "headerReference")
				let type = e.Attribute(w + "type")
				where type != null && type.Value.Equals("first", StringComparison.CurrentCultureIgnoreCase)
				select e.Attribute(r + "id").Value).LastOrDefault();
				if (text3 != null)
				{
					XElement xml3 = headers.first.Xml;
					Uri uri3 = PackUriHelper.ResolvePartUri(mainPart.Uri, mainPart.GetRelationship(text3).TargetUri);
					using (TextWriter textWriter4 = new StreamWriter(new PackagePartStream(package.GetPart(uri3).GetStream(FileMode.Create, FileAccess.Write))))
					{
						new XDocument(new XDeclaration("1.0", "UTF-8", "yes"), xml3).Save(textWriter4, SaveOptions.None);
					}
				}
				string text4 = (from e in mainDoc.Descendants(w + "footerReference")
				let type = e.Attribute(w + "type")
				where type != null && type.Value.Equals("default", StringComparison.CurrentCultureIgnoreCase)
				select e.Attribute(r + "id").Value).LastOrDefault();
				if (text4 != null)
				{
					XElement xml4 = footers.odd.Xml;
					Uri uri4 = PackUriHelper.ResolvePartUri(mainPart.Uri, mainPart.GetRelationship(text4).TargetUri);
					using (TextWriter textWriter5 = new StreamWriter(new PackagePartStream(package.GetPart(uri4).GetStream(FileMode.Create, FileAccess.Write))))
					{
						new XDocument(new XDeclaration("1.0", "UTF-8", "yes"), xml4).Save(textWriter5, SaveOptions.None);
					}
				}
				string text5 = (from e in mainDoc.Descendants(w + "footerReference")
				let type = e.Attribute(w + "type")
				where type != null && type.Value.Equals("even", StringComparison.CurrentCultureIgnoreCase)
				select e.Attribute(r + "id").Value).LastOrDefault();
				if (text5 != null)
				{
					XElement xml5 = footers.even.Xml;
					Uri uri5 = PackUriHelper.ResolvePartUri(mainPart.Uri, mainPart.GetRelationship(text5).TargetUri);
					using (TextWriter textWriter6 = new StreamWriter(new PackagePartStream(package.GetPart(uri5).GetStream(FileMode.Create, FileAccess.Write))))
					{
						new XDocument(new XDeclaration("1.0", "UTF-8", "yes"), xml5).Save(textWriter6, SaveOptions.None);
					}
				}
				string text6 = (from e in mainDoc.Descendants(w + "footerReference")
				let type = e.Attribute(w + "type")
				where type != null && type.Value.Equals("first", StringComparison.CurrentCultureIgnoreCase)
				select e.Attribute(r + "id").Value).LastOrDefault();
				if (text6 != null)
				{
					XElement xml6 = footers.first.Xml;
					Uri uri6 = PackUriHelper.ResolvePartUri(mainPart.Uri, mainPart.GetRelationship(text6).TargetUri);
					using (TextWriter textWriter7 = new StreamWriter(new PackagePartStream(package.GetPart(uri6).GetStream(FileMode.Create, FileAccess.Write))))
					{
						new XDocument(new XDeclaration("1.0", "UTF-8", "yes"), xml6).Save(textWriter7, SaveOptions.None);
					}
				}
				using (TextWriter textWriter8 = new StreamWriter(new PackagePartStream(settingsPart.GetStream(FileMode.Create, FileAccess.Write))))
				{
					settings.Save(textWriter8, SaveOptions.None);
				}
				if (endnotesPart != null)
				{
					using (TextWriter textWriter9 = new StreamWriter(new PackagePartStream(endnotesPart.GetStream(FileMode.Create, FileAccess.Write))))
					{
						endnotes.Save(textWriter9, SaveOptions.None);
					}
				}
				if (footnotesPart != null)
				{
					using (TextWriter textWriter10 = new StreamWriter(new PackagePartStream(footnotesPart.GetStream(FileMode.Create, FileAccess.Write))))
					{
						footnotes.Save(textWriter10, SaveOptions.None);
					}
				}
				if (stylesPart != null)
				{
					using (TextWriter textWriter11 = new StreamWriter(new PackagePartStream(stylesPart.GetStream(FileMode.Create, FileAccess.Write))))
					{
						styles.Save(textWriter11, SaveOptions.None);
					}
				}
				if (stylesWithEffectsPart != null)
				{
					using (TextWriter textWriter12 = new StreamWriter(new PackagePartStream(stylesWithEffectsPart.GetStream(FileMode.Create, FileAccess.Write))))
					{
						stylesWithEffects.Save(textWriter12, SaveOptions.None);
					}
				}
				if (numberingPart != null)
				{
					using (TextWriter textWriter13 = new StreamWriter(new PackagePartStream(numberingPart.GetStream(FileMode.Create, FileAccess.Write))))
					{
						numbering.Save(textWriter13, SaveOptions.None);
					}
				}
				if (fontTablePart != null)
				{
					using (TextWriter textWriter14 = new StreamWriter(new PackagePartStream(fontTablePart.GetStream(FileMode.Create, FileAccess.Write))))
					{
						fontTable.Save(textWriter14, SaveOptions.None);
					}
				}
			}
			package.Flush();
			package.Close();
			if (filename != null)
			{
				using (FileStream fileStream = new FileStream(filename, FileMode.Create))
				{
					if (memoryStream.CanSeek)
					{
						memoryStream.Position = 0L;
						CopyStream(memoryStream, fileStream);
					}
					else
					{
						fileStream.Write(memoryStream.ToArray(), 0, (int)memoryStream.Length);
					}
				}
			}
			else
			{
				if (stream.CanSeek)
				{
					stream.SetLength(0L);
					stream.Position = 0L;
				}
				memoryStream.WriteTo(stream);
				memoryStream.Flush();
			}
		}

		public void SaveAs(string filename)
		{
			this.filename = filename;
			stream = null;
			Save();
		}

		public void SaveAs(Stream stream)
		{
			filename = null;
			this.stream = stream;
			Save();
		}

		public void AddCoreProperty(string propertyName, string propertyValue)
		{
			string prefix = propertyName.Contains(":") ? propertyName.Split(new char[1]
			{
				':'
			})[0] : "cp";
			string propertyLocalName = propertyName.Contains(":") ? propertyName.Split(new char[1]
			{
				':'
			})[1] : propertyName;
			if (!package.PartExists(new Uri("/docProps/core.xml", UriKind.Relative)))
			{
				throw new Exception("Core properties part doesn't exist.");
			}
			PackagePart part = package.GetPart(new Uri("/docProps/core.xml", UriKind.Relative));
			XDocument xDocument;
			using (TextReader textReader = new StreamReader(part.GetStream(FileMode.Open, FileAccess.Read)))
			{
				xDocument = XDocument.Load(textReader);
			}
			XElement xElement = (from propElement in xDocument.Root.Elements()
			where propElement.Name.LocalName.Equals(propertyLocalName)
			select propElement).SingleOrDefault();
			if (xElement != null)
			{
				xElement.SetValue(propertyValue);
			}
			else
			{
				XNamespace namespaceOfPrefix = xDocument.Root.GetNamespaceOfPrefix(prefix);
				xDocument.Root.Add(new XElement(XName.Get(propertyLocalName, namespaceOfPrefix.NamespaceName), propertyValue));
			}
			using (TextWriter textWriter = new StreamWriter(new PackagePartStream(part.GetStream(FileMode.Create, FileAccess.Write))))
			{
				xDocument.Save(textWriter);
			}
			UpdateCorePropertyValue(this, propertyLocalName, propertyValue);
		}

		internal static void UpdateCorePropertyValue(DocX document, string corePropertyName, string corePropertyValue)
		{
			string pattern = $"(DOCPROPERTY)?{corePropertyName}\\\\\\*MERGEFORMAT".ToLower();
			foreach (XElement item in document.mainDoc.Descendants(XName.Get("fldSimple", w.NamespaceName)))
			{
				string input = item.Attribute(XName.Get("instr", w.NamespaceName)).Value.Replace(" ", string.Empty).Trim().ToLower();
				if (Regex.IsMatch(input, pattern))
				{
					XElement xElement = item.Element(w + "r");
					XElement xElement2 = xElement.Element(w + "t");
					XElement xElement3 = xElement2.Element(w + "rPr");
					item.RemoveNodes();
					XElement xElement4 = new XElement(w + "t", xElement3, corePropertyValue);
					Novacode.Text.PreserveSpace(xElement4);
					item.Add(new XElement(xElement.Name, xElement.Attributes(), xElement.Element(XName.Get("rPr", w.NamespaceName)), xElement4));
				}
			}
			IEnumerable<PackagePart> enumerable = from headerPart in (IEnumerable<PackagePart>)document.package.GetParts()
			where Regex.IsMatch(headerPart.Uri.ToString(), "/word/header\\d?.xml")
			select headerPart;
			foreach (PackagePart item2 in enumerable)
			{
				XDocument xDocument = XDocument.Load(new StreamReader(item2.GetStream()));
				foreach (XElement item3 in xDocument.Descendants(XName.Get("fldSimple", w.NamespaceName)))
				{
					string input2 = item3.Attribute(XName.Get("instr", w.NamespaceName)).Value.Replace(" ", string.Empty).Trim().ToLower();
					if (Regex.IsMatch(input2, pattern))
					{
						XElement xElement5 = item3.Element(w + "r");
						item3.RemoveNodes();
						XElement xElement6 = new XElement(w + "t", corePropertyValue);
						Novacode.Text.PreserveSpace(xElement6);
						item3.Add(new XElement(xElement5.Name, xElement5.Attributes(), xElement5.Element(XName.Get("rPr", w.NamespaceName)), xElement6));
					}
				}
				using (TextWriter textWriter = new StreamWriter(new PackagePartStream(item2.GetStream(FileMode.Create, FileAccess.Write))))
				{
					xDocument.Save(textWriter);
				}
			}
			IEnumerable<PackagePart> enumerable2 = from footerPart in (IEnumerable<PackagePart>)document.package.GetParts()
			where Regex.IsMatch(footerPart.Uri.ToString(), "/word/footer\\d?.xml")
			select footerPart;
			foreach (PackagePart item4 in enumerable2)
			{
				XDocument xDocument2 = XDocument.Load(new StreamReader(item4.GetStream()));
				foreach (XElement item5 in xDocument2.Descendants(XName.Get("fldSimple", w.NamespaceName)))
				{
					string input3 = item5.Attribute(XName.Get("instr", w.NamespaceName)).Value.Replace(" ", string.Empty).Trim().ToLower();
					if (Regex.IsMatch(input3, pattern))
					{
						XElement xElement7 = item5.Element(w + "r");
						item5.RemoveNodes();
						XElement xElement8 = new XElement(w + "t", corePropertyValue);
						Novacode.Text.PreserveSpace(xElement8);
						item5.Add(new XElement(xElement7.Name, xElement7.Attributes(), xElement7.Element(XName.Get("rPr", w.NamespaceName)), xElement8));
					}
				}
				using (TextWriter textWriter2 = new StreamWriter(new PackagePartStream(item4.GetStream(FileMode.Create, FileAccess.Write))))
				{
					xDocument2.Save(textWriter2);
				}
			}
			PopulateDocument(document, document.package);
		}

		public void AddCustomProperty(CustomProperty cp)
		{
			if (!package.PartExists(new Uri("/docProps/custom.xml", UriKind.Relative)))
			{
				HelperFunctions.CreateCustomPropertiesPart(this);
			}
			PackagePart part = package.GetPart(new Uri("/docProps/custom.xml", UriKind.Relative));
			XDocument xDocument;
			using (TextReader textReader = new StreamReader(part.GetStream(FileMode.Open, FileAccess.Read)))
			{
				xDocument = XDocument.Load(textReader, LoadOptions.PreserveWhitespace);
			}
			IEnumerable<int> source = from d in xDocument.Descendants()
			where d.Name.LocalName == "property"
			select int.Parse(d.Attribute(XName.Get("pid")).Value);
			int num = 1;
			if (source.Count() > 0)
			{
				num = source.Max();
			}
			(from d in xDocument.Descendants()
			where d.Name.LocalName == "property" && d.Attribute(XName.Get("name")).Value.Equals(cp.Name, StringComparison.CurrentCultureIgnoreCase)
			select d).SingleOrDefault()?.Remove();
			XElement xElement = xDocument.Element(XName.Get("Properties", customPropertiesSchema.NamespaceName));
			xElement.Add(new XElement(XName.Get("property", customPropertiesSchema.NamespaceName), new XAttribute("fmtid", "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"), new XAttribute("pid", num + 1), new XAttribute("name", cp.Name), new XElement(customVTypesSchema + cp.Type, cp.Value ?? "")));
			using (TextWriter textWriter = new StreamWriter(new PackagePartStream(part.GetStream(FileMode.Create, FileAccess.Write))))
			{
				xDocument.Save(textWriter, SaveOptions.None);
			}
			UpdateCustomPropertyValue(this, cp.Name, (cp.Value ?? "").ToString());
		}

		internal static void UpdateCustomPropertyValue(DocX document, string customPropertyName, string customPropertyValue)
		{
			List<XElement> list = new List<XElement>
			{
				document.mainDoc.Root
			};
			Headers headers = document.Headers;
			if (headers.first != null)
			{
				list.Add(headers.first.Xml);
			}
			if (headers.odd != null)
			{
				list.Add(headers.odd.Xml);
			}
			if (headers.even != null)
			{
				list.Add(headers.even.Xml);
			}
			Footers footers = document.Footers;
			if (footers.first != null)
			{
				list.Add(footers.first.Xml);
			}
			if (footers.odd != null)
			{
				list.Add(footers.odd.Xml);
			}
			if (footers.even != null)
			{
				list.Add(footers.even.Xml);
			}
			string arg = customPropertyName;
			if (customPropertyName.Contains(" "))
			{
				arg = "\"" + customPropertyName + "\"";
			}
			string value = $"DOCPROPERTY  {arg}  \\* MERGEFORMAT".Replace(" ", string.Empty);
			foreach (XElement item in list)
			{
				foreach (XElement item2 in item.Descendants(XName.Get("instrText", w.NamespaceName)))
				{
					string text = item2.Value.Replace(" ", string.Empty).Trim();
					if (text.Equals(value, StringComparison.CurrentCultureIgnoreCase))
					{
						XNode nextNode = item2.Parent.NextNode;
						bool flag = false;
						while (true)
						{
							if (nextNode.NodeType == XmlNodeType.Element)
							{
								XElement xElement = nextNode as XElement;
								IEnumerable<XElement> source = xElement.Descendants(XName.Get("t", w.NamespaceName));
								if (source.Any())
								{
									if (!flag)
									{
										source.First().Value = customPropertyValue;
										flag = true;
									}
									else
									{
										xElement.RemoveNodes();
									}
								}
								else
								{
									source = xElement.Descendants(XName.Get("fldChar", w.NamespaceName));
									if (source.Any())
									{
										XAttribute xAttribute = source.First().Attribute(XName.Get("fldCharType", w.NamespaceName));
										if (xAttribute != null && xAttribute.Value == "end")
										{
											break;
										}
									}
								}
							}
							nextNode = nextNode.NextNode;
						}
					}
				}
				foreach (XElement item3 in item.Descendants(XName.Get("fldSimple", w.NamespaceName)))
				{
					string text2 = item3.Attribute(XName.Get("instr", w.NamespaceName)).Value.Replace(" ", string.Empty).Trim();
					if (text2.Equals(value, StringComparison.CurrentCultureIgnoreCase))
					{
						XElement xElement2 = item3.Element(w + "r");
						XElement xElement3 = xElement2.Element(w + "t");
						XElement xElement4 = xElement3.Element(w + "rPr");
						item3.RemoveNodes();
						XElement xElement5 = new XElement(w + "t", xElement4, customPropertyValue);
						Novacode.Text.PreserveSpace(xElement5);
						item3.Add(new XElement(xElement2.Name, xElement2.Attributes(), xElement2.Element(XName.Get("rPr", w.NamespaceName)), xElement5));
					}
				}
			}
		}

		public override Paragraph InsertParagraph()
		{
			Paragraph paragraph = base.InsertParagraph();
			paragraph.PackagePart = mainPart;
			return paragraph;
		}

		public override Paragraph InsertParagraph(int index, string text, bool trackChanges)
		{
			Paragraph paragraph = base.InsertParagraph(index, text, trackChanges);
			paragraph.PackagePart = mainPart;
			return paragraph;
		}

		public override Paragraph InsertParagraph(Paragraph p)
		{
			p.PackagePart = mainPart;
			return base.InsertParagraph(p);
		}

		public override Paragraph InsertParagraph(int index, Paragraph p)
		{
			p.PackagePart = mainPart;
			return base.InsertParagraph(index, p);
		}

		public override Paragraph InsertParagraph(int index, string text, bool trackChanges, Formatting formatting)
		{
			Paragraph paragraph = base.InsertParagraph(index, text, trackChanges, formatting);
			paragraph.PackagePart = mainPart;
			return paragraph;
		}

		public override Paragraph InsertParagraph(string text)
		{
			Paragraph paragraph = base.InsertParagraph(text);
			paragraph.PackagePart = mainPart;
			return paragraph;
		}

		public override Paragraph InsertParagraph(string text, bool trackChanges)
		{
			Paragraph paragraph = base.InsertParagraph(text, trackChanges);
			paragraph.PackagePart = mainPart;
			return paragraph;
		}

		public override Paragraph InsertParagraph(string text, bool trackChanges, Formatting formatting)
		{
			Paragraph paragraph = base.InsertParagraph(text, trackChanges, formatting);
			paragraph.PackagePart = mainPart;
			return paragraph;
		}

		public Paragraph[] InsertParagraphs(string text)
		{
			string[] array = text.Split(new char[1]
			{
				'\n'
			});
			List<Paragraph> list = new List<Paragraph>();
			string[] array2 = array;
			foreach (string text2 in array2)
			{
				Paragraph paragraph = base.InsertParagraph(text);
				paragraph.PackagePart = mainPart;
				list.Add(paragraph);
			}
			return list.ToArray();
		}

		public void SetContent(XElement el)
		{
			foreach (XElement item in el.Elements())
			{
				(from d in base.Document.Contents
				where d.Name == item.Name
				select d).First().SetText(item.Value);
			}
		}

		public void SetContent(Dictionary<string, string> dict)
		{
			foreach (KeyValuePair<string, string> item in dict)
			{
				(from d in base.Document.Contents
				where d.Name == item.Key
				select d).First().SetText(item.Value);
			}
		}

		public void SetContent(string path)
		{
			XDocument content = XDocument.Load(path);
			SetContent(content);
		}

		public void SetContent(XDocument xmlDoc)
		{
			foreach (XElement item in xmlDoc.ElementsAfterSelf())
			{
				(from d in base.Document.Contents
				where d.Name == item.Name
				select d).First().SetText(item.Value);
			}
		}

		public override Paragraph InsertEquation(string equation)
		{
			Paragraph paragraph = base.InsertEquation(equation);
			paragraph.PackagePart = mainPart;
			return paragraph;
		}

		public void InsertChart(Chart chart)
		{
			string empty = string.Empty;
			int num = 1;
			do
			{
				empty = $"/word/charts/chart{num}.xml";
				num++;
			}
			while (package.PartExists(new Uri(empty, UriKind.Relative)));
			PackagePart val = package.CreatePart(new Uri(empty, UriKind.Relative), "application/vnd.openxmlformats-officedocument.drawingml.chart+xml", 0);
			string nextFreeRelationshipID = GetNextFreeRelationshipID();
			PackageRelationship val2 = mainPart.CreateRelationship(val.Uri, 0, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart", nextFreeRelationshipID);
			using (TextWriter textWriter = new StreamWriter(new PackagePartStream(val.GetStream(FileMode.Create, FileAccess.Write))))
			{
				chart.Xml.Save(textWriter);
			}
			Paragraph paragraph = InsertParagraph();
			XElement content = new XElement(XName.Get("r", w.NamespaceName), new XElement(XName.Get("drawing", w.NamespaceName), new XElement(XName.Get("inline", wp.NamespaceName), new XElement(XName.Get("extent", wp.NamespaceName), new XAttribute("cx", "5486400"), new XAttribute("cy", "3200400")), new XElement(XName.Get("effectExtent", wp.NamespaceName), new XAttribute("l", "0"), new XAttribute("t", "0"), new XAttribute("r", "19050"), new XAttribute("b", "19050")), new XElement(XName.Get("docPr", wp.NamespaceName), new XAttribute("id", "1"), new XAttribute("name", "chart")), new XElement(XName.Get("graphic", a.NamespaceName), new XElement(XName.Get("graphicData", a.NamespaceName), new XAttribute("uri", c.NamespaceName), new XElement(XName.Get("chart", c.NamespaceName), new XAttribute(XName.Get("id", r.NamespaceName), nextFreeRelationshipID)))))));
			paragraph.Xml.Add(content);
		}

		public void InsertChartAfterParagraph(Chart chart, Paragraph paragraph)
		{
			string empty = string.Empty;
			int num = 1;
			do
			{
				empty = $"/word/charts/chart{num}.xml";
				num++;
			}
			while (package.PartExists(new Uri(empty, UriKind.Relative)));
			PackagePart val = package.CreatePart(new Uri(empty, UriKind.Relative), "application/vnd.openxmlformats-officedocument.drawingml.chart+xml", 0);
			string nextFreeRelationshipID = GetNextFreeRelationshipID();
			PackageRelationship val2 = mainPart.CreateRelationship(val.Uri, 0, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart", nextFreeRelationshipID);
			using (TextWriter textWriter = new StreamWriter(val.GetStream(FileMode.Create, FileAccess.Write)))
			{
				chart.Xml.Save(textWriter);
			}
			XElement content = new XElement(XName.Get("r", w.NamespaceName), new XElement(XName.Get("drawing", w.NamespaceName), new XElement(XName.Get("inline", wp.NamespaceName), new XElement(XName.Get("extent", wp.NamespaceName), new XAttribute("cx", "5486400"), new XAttribute("cy", "3200400")), new XElement(XName.Get("effectExtent", wp.NamespaceName), new XAttribute("l", "0"), new XAttribute("t", "0"), new XAttribute("r", "19050"), new XAttribute("b", "19050")), new XElement(XName.Get("docPr", wp.NamespaceName), new XAttribute("id", "1"), new XAttribute("name", "chart")), new XElement(XName.Get("graphic", a.NamespaceName), new XElement(XName.Get("graphicData", a.NamespaceName), new XAttribute("uri", c.NamespaceName), new XElement(XName.Get("chart", c.NamespaceName), new XAttribute(XName.Get("id", r.NamespaceName), nextFreeRelationshipID)))))));
			paragraph.Xml.Add(content);
		}

		public TableOfContents InsertDefaultTableOfContents()
		{
			return InsertTableOfContents("Table of contents", TableOfContentsSwitches.H | TableOfContentsSwitches.O | TableOfContentsSwitches.U | TableOfContentsSwitches.Z);
		}

		public TableOfContents InsertTableOfContents(string title, TableOfContentsSwitches switches, string headerStyle = null, int maxIncludeLevel = 3, int? rightTabPos = default(int?))
		{
			TableOfContents tableOfContents = TableOfContents.CreateTableOfContents(this, title, switches, headerStyle, maxIncludeLevel, rightTabPos);
			base.Xml.Add(tableOfContents.Xml);
			return tableOfContents;
		}

		public TableOfContents InsertTableOfContents(Paragraph reference, string title, TableOfContentsSwitches switches, string headerStyle = null, int maxIncludeLevel = 3, int? rightTabPos = default(int?))
		{
			TableOfContents tableOfContents = TableOfContents.CreateTableOfContents(this, title, switches, headerStyle, maxIncludeLevel, rightTabPos);
			reference.Xml.AddBeforeSelf(tableOfContents.Xml);
			return tableOfContents;
		}

		private static void CopyStream(Stream input, Stream output, int bufferSize = 32768)
		{
			byte[] array = new byte[bufferSize];
			int count;
			while ((count = input.Read(array, 0, array.Length)) > 0)
			{
				output.Write(array, 0, count);
			}
		}

		internal long GetNextFreeDocPrId()
		{
			lock (nextFreeDocPrIdLock)
			{
				if (nextFreeDocPrId.HasValue)
				{
					nextFreeDocPrId++;
					return nextFreeDocPrId.Value;
				}
				XName right = XName.Get("bookmarkStart", w.NamespaceName);
				XName right2 = XName.Get("docPr", wp.NamespaceName);
				long num = 1L;
				HashSet<string> hashSet = new HashSet<string>();
				foreach (XElement item in base.Xml.Descendants())
				{
					if (!(item.Name != right) || !(item.Name != right2))
					{
						XAttribute xAttribute = item.Attributes().FirstOrDefault((XAttribute x) => x.Name.LocalName == "id");
						if (xAttribute != null)
						{
							hashSet.Add(xAttribute.Value);
						}
					}
				}
				while (hashSet.Contains(num.ToString()))
				{
					num++;
				}
				nextFreeDocPrId = num;
				return nextFreeDocPrId.Value;
			}
		}

		public void Dispose()
		{
			package.Close();
		}
	}
}
