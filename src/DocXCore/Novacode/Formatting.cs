using System;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Xml.Linq;

namespace Novacode
{
	public class Formatting : IComparable
	{
		private XElement rPr;

		private bool? hidden;

		private bool? bold;

		private bool? italic;

		private StrikeThrough? strikethrough;

		private Script? script;

		private Highlight? highlight;

		private double? size;

		private Color? fontColor;

		private Color? underlineColor;

		private UnderlineStyle? underlineStyle;

		private Misc? misc;

		private CapsStyle? capsStyle;

		private Font fontFamily;

		private int? percentageScale;

		private int? kerning;

		private int? position;

		private double? spacing;

		private CultureInfo language;

		public CultureInfo Language
		{
			get
			{
				return language;
			}
			set
			{
				language = value;
			}
		}

		internal XElement Xml
		{
			get
			{
				rPr = new XElement(XName.Get("rPr", DocX.w.NamespaceName));
				if (language != null)
				{
					rPr.Add(new XElement(XName.Get("lang", DocX.w.NamespaceName), new XAttribute(XName.Get("val", DocX.w.NamespaceName), language.Name)));
				}
				if (spacing.HasValue)
				{
					rPr.Add(new XElement(XName.Get("spacing", DocX.w.NamespaceName), new XAttribute(XName.Get("val", DocX.w.NamespaceName), spacing.Value * 20.0)));
				}
				if (position.HasValue)
				{
					rPr.Add(new XElement(XName.Get("position", DocX.w.NamespaceName), new XAttribute(XName.Get("val", DocX.w.NamespaceName), position.Value * 2)));
				}
				if (kerning.HasValue)
				{
					rPr.Add(new XElement(XName.Get("kern", DocX.w.NamespaceName), new XAttribute(XName.Get("val", DocX.w.NamespaceName), kerning.Value * 2)));
				}
				if (percentageScale.HasValue)
				{
					rPr.Add(new XElement(XName.Get("w", DocX.w.NamespaceName), new XAttribute(XName.Get("val", DocX.w.NamespaceName), percentageScale)));
				}
				if (fontFamily != null)
				{
					rPr.Add(new XElement(XName.Get("rFonts", DocX.w.NamespaceName), new XAttribute(XName.Get("ascii", DocX.w.NamespaceName), fontFamily.Name), new XAttribute(XName.Get("hAnsi", DocX.w.NamespaceName), fontFamily.Name), new XAttribute(XName.Get("cs", DocX.w.NamespaceName), fontFamily.Name)));
				}
				if (hidden.HasValue && hidden.Value)
				{
					rPr.Add(new XElement(XName.Get("vanish", DocX.w.NamespaceName)));
				}
				if (bold.HasValue && bold.Value)
				{
					rPr.Add(new XElement(XName.Get("b", DocX.w.NamespaceName)));
				}
				if (italic.HasValue && italic.Value)
				{
					rPr.Add(new XElement(XName.Get("i", DocX.w.NamespaceName)));
				}
				if (underlineStyle.HasValue)
				{
					switch (underlineStyle)
					{
					case Novacode.UnderlineStyle.singleLine:
						rPr.Add(new XElement(XName.Get("u", DocX.w.NamespaceName), new XAttribute(XName.Get("val", DocX.w.NamespaceName), "single")));
						break;
					case Novacode.UnderlineStyle.doubleLine:
						rPr.Add(new XElement(XName.Get("u", DocX.w.NamespaceName), new XAttribute(XName.Get("val", DocX.w.NamespaceName), "double")));
						break;
					default:
						rPr.Add(new XElement(XName.Get("u", DocX.w.NamespaceName), new XAttribute(XName.Get("val", DocX.w.NamespaceName), underlineStyle.ToString())));
						break;
					case Novacode.UnderlineStyle.none:
						break;
					}
				}
				if (underlineColor.HasValue)
				{
					if (underlineStyle == Novacode.UnderlineStyle.none)
					{
						underlineStyle = Novacode.UnderlineStyle.singleLine;
						rPr.Add(new XElement(XName.Get("u", DocX.w.NamespaceName), new XAttribute(XName.Get("val", DocX.w.NamespaceName), "single")));
					}
					rPr.Element(XName.Get("u", DocX.w.NamespaceName)).Add(new XAttribute(XName.Get("color", DocX.w.NamespaceName), underlineColor.Value.ToHex()));
				}
				if (strikethrough.HasValue)
				{
					switch (strikethrough)
					{
					case Novacode.StrikeThrough.strike:
						rPr.Add(new XElement(XName.Get("strike", DocX.w.NamespaceName)));
						break;
					case Novacode.StrikeThrough.doubleStrike:
						rPr.Add(new XElement(XName.Get("dstrike", DocX.w.NamespaceName)));
						break;
					}
				}
				if (this.script.HasValue)
				{
					Script? script = this.script;
					Script? script2 = script;
					if (script2.HasValue)
					{
						Script valueOrDefault = script2.GetValueOrDefault();
						if (valueOrDefault == Novacode.Script.none)
						{
							goto IL_06b2;
						}
					}
					rPr.Add(new XElement(XName.Get("vertAlign", DocX.w.NamespaceName), new XAttribute(XName.Get("val", DocX.w.NamespaceName), this.script.ToString())));
				}
				goto IL_06b2;
				IL_06b2:
				if (size.HasValue)
				{
					rPr.Add(new XElement(XName.Get("sz", DocX.w.NamespaceName), new XAttribute(XName.Get("val", DocX.w.NamespaceName), (size * 2.0).ToString())));
					rPr.Add(new XElement(XName.Get("szCs", DocX.w.NamespaceName), new XAttribute(XName.Get("val", DocX.w.NamespaceName), (size * 2.0).ToString())));
				}
				if (fontColor.HasValue)
				{
					rPr.Add(new XElement(XName.Get("color", DocX.w.NamespaceName), new XAttribute(XName.Get("val", DocX.w.NamespaceName), fontColor.Value.ToHex())));
				}
				if (this.highlight.HasValue)
				{
					Highlight? highlight = this.highlight;
					Highlight? highlight2 = highlight;
					if (highlight2.HasValue)
					{
						Highlight valueOrDefault2 = highlight2.GetValueOrDefault();
						if (valueOrDefault2 == Novacode.Highlight.none)
						{
							goto IL_08ad;
						}
					}
					rPr.Add(new XElement(XName.Get("highlight", DocX.w.NamespaceName), new XAttribute(XName.Get("val", DocX.w.NamespaceName), this.highlight.ToString())));
				}
				goto IL_08ad;
				IL_08ad:
				if (capsStyle.HasValue && (capsStyle ?? Novacode.CapsStyle.caps) != 0)
				{
					rPr.Add(new XElement(XName.Get(capsStyle.ToString(), DocX.w.NamespaceName)));
				}
				if (misc.HasValue)
				{
					switch (misc)
					{
					case Novacode.Misc.outlineShadow:
						rPr.Add(new XElement(XName.Get("outline", DocX.w.NamespaceName)));
						rPr.Add(new XElement(XName.Get("shadow", DocX.w.NamespaceName)));
						break;
					case Novacode.Misc.engrave:
						rPr.Add(new XElement(XName.Get("imprint", DocX.w.NamespaceName)));
						break;
					default:
						rPr.Add(new XElement(XName.Get(misc.ToString(), DocX.w.NamespaceName)));
						break;
					case Novacode.Misc.none:
						break;
					}
				}
				return rPr;
			}
		}

		public bool? Bold
		{
			get
			{
				return bold;
			}
			set
			{
				bold = value;
			}
		}

		public bool? Italic
		{
			get
			{
				return italic;
			}
			set
			{
				italic = value;
			}
		}

		public StrikeThrough? StrikeThrough
		{
			get
			{
				return strikethrough;
			}
			set
			{
				strikethrough = value;
			}
		}

		public Script? Script
		{
			get
			{
				return script;
			}
			set
			{
				script = value;
			}
		}

		public double? Size
		{
			get
			{
				return size;
			}
			set
			{
				double? num = value * 2.0;
				if (num - (double)(int)num.Value != 0.0)
				{
					throw new ArgumentException("Size", "Value must be either a whole or half number, examples: 32, 32.5");
				}
				if (!(value > 0.0) || !(value < 1639.0))
				{
					throw new ArgumentException("Size", "Value must be in the range 0 - 1638");
				}
				size = value;
			}
		}

		public int? PercentageScale
		{
			get
			{
				return percentageScale;
			}
			set
			{
				if (!new int?[8]
				{
					200,
					150,
					100,
					90,
					80,
					66,
					50,
					33
				}.Contains(value))
				{
					throw new ArgumentOutOfRangeException("PercentageScale", "Value must be one of the following: 200, 150, 100, 90, 80, 66, 50 or 33");
				}
				percentageScale = value;
			}
		}

		public int? Kerning
		{
			get
			{
				return kerning;
			}
			set
			{
				if (!new int?[16]
				{
					8,
					9,
					10,
					11,
					12,
					14,
					16,
					18,
					20,
					22,
					24,
					26,
					28,
					36,
					48,
					72
				}.Contains(value))
				{
					throw new ArgumentOutOfRangeException("Kerning", "Value must be one of the following: 8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48 or 72");
				}
				kerning = value;
			}
		}

		public int? Position
		{
			get
			{
				return position;
			}
			set
			{
				if (!(value > -1585) || !(value < 1585))
				{
					throw new ArgumentOutOfRangeException("Position", "Value must be in the range -1585 - 1585");
				}
				position = value;
			}
		}

		public double? Spacing
		{
			get
			{
				return spacing;
			}
			set
			{
				double? num = value * 20.0;
				if (num - (double)(int)num.Value != 0.0)
				{
					throw new ArgumentException("Spacing", "Value must be either a whole or acurate to one decimal, examples: 32, 32.1, 32.2, 32.9");
				}
				if (!(value > -1585.0) || !(value < 1585.0))
				{
					throw new ArgumentException("Spacing", "Value must be in the range: -1584 - 1584");
				}
				spacing = value;
			}
		}

		public Color? FontColor
		{
			get
			{
				return fontColor;
			}
			set
			{
				fontColor = value;
			}
		}

		public Highlight? Highlight
		{
			get
			{
				return highlight;
			}
			set
			{
				highlight = value;
			}
		}

		public UnderlineStyle? UnderlineStyle
		{
			get
			{
				return underlineStyle;
			}
			set
			{
				underlineStyle = value;
			}
		}

		public Color? UnderlineColor
		{
			get
			{
				return underlineColor;
			}
			set
			{
				underlineColor = value;
			}
		}

		public Misc? Misc
		{
			get
			{
				return misc;
			}
			set
			{
				misc = value;
			}
		}

		public bool? Hidden
		{
			get
			{
				return hidden;
			}
			set
			{
				hidden = value;
			}
		}

		public CapsStyle? CapsStyle
		{
			get
			{
				return capsStyle;
			}
			set
			{
				capsStyle = value;
			}
		}

		public Font FontFamily
		{
			get
			{
				return fontFamily;
			}
			set
			{
				fontFamily = value;
			}
		}

		public Formatting()
		{
			capsStyle = Novacode.CapsStyle.none;
			strikethrough = Novacode.StrikeThrough.none;
			script = Novacode.Script.none;
			highlight = Novacode.Highlight.none;
			underlineStyle = Novacode.UnderlineStyle.none;
			misc = Novacode.Misc.none;
			language = CultureInfo.CurrentCulture;
			rPr = new XElement(XName.Get("rPr", DocX.w.NamespaceName));
		}

		public Formatting Clone()
		{
			Formatting formatting = new Formatting();
			formatting.Bold = bold;
			formatting.CapsStyle = capsStyle;
			formatting.FontColor = fontColor;
			formatting.FontFamily = fontFamily;
			formatting.Hidden = hidden;
			formatting.Highlight = highlight;
			formatting.Italic = italic;
			if (kerning.HasValue)
			{
				formatting.Kerning = kerning;
			}
			formatting.Language = language;
			formatting.Misc = misc;
			if (percentageScale.HasValue)
			{
				formatting.PercentageScale = percentageScale;
			}
			if (position.HasValue)
			{
				formatting.Position = position;
			}
			formatting.Script = script;
			if (size.HasValue)
			{
				formatting.Size = size;
			}
			if (spacing.HasValue)
			{
				formatting.Spacing = spacing;
			}
			formatting.StrikeThrough = strikethrough;
			formatting.UnderlineColor = underlineColor;
			formatting.UnderlineStyle = underlineStyle;
			return formatting;
		}

		public static Formatting Parse(XElement rPr)
		{
			Formatting formatting = new Formatting();
			foreach (XElement item in rPr.Elements())
			{
				switch (item.Name.LocalName)
				{
				case "lang":
					formatting.Language = new CultureInfo(item.GetAttribute(XName.Get("val", DocX.w.NamespaceName), null) ?? item.GetAttribute(XName.Get("eastAsia", DocX.w.NamespaceName), null) ?? item.GetAttribute(XName.Get("bidi", DocX.w.NamespaceName)));
					break;
				case "spacing":
					formatting.Spacing = double.Parse(item.GetAttribute(XName.Get("val", DocX.w.NamespaceName))) / 20.0;
					break;
				case "position":
					formatting.Position = int.Parse(item.GetAttribute(XName.Get("val", DocX.w.NamespaceName))) / 2;
					break;
				case "kern":
					formatting.Position = int.Parse(item.GetAttribute(XName.Get("val", DocX.w.NamespaceName))) / 2;
					break;
				case "w":
					formatting.PercentageScale = int.Parse(item.GetAttribute(XName.Get("val", DocX.w.NamespaceName)));
					break;
				case "sz":
					formatting.Size = (double)(int.Parse(item.GetAttribute(XName.Get("val", DocX.w.NamespaceName))) / 2);
					break;
				case "rFonts":
					formatting.FontFamily = new Font(item.GetAttribute(XName.Get("cs", DocX.w.NamespaceName), null) ?? item.GetAttribute(XName.Get("ascii", DocX.w.NamespaceName), null) ?? item.GetAttribute(XName.Get("hAnsi", DocX.w.NamespaceName), null) ?? item.GetAttribute(XName.Get("eastAsia", DocX.w.NamespaceName)));
					break;
				case "color":
					try
					{
						string attribute2 = item.GetAttribute(XName.Get("val", DocX.w.NamespaceName));
						formatting.FontColor = ColorTranslator.FromHtml($"#{attribute2}");
					}
					catch
					{
					}
					break;
				case "vanish":
					formatting.hidden = true;
					break;
				case "b":
					formatting.Bold = true;
					break;
				case "i":
					formatting.Italic = true;
					break;
				case "u":
					formatting.UnderlineStyle = HelperFunctions.GetUnderlineStyle(item.GetAttribute(XName.Get("val", DocX.w.NamespaceName)));
					break;
				case "vertAlign":
				{
					string attribute = item.GetAttribute(XName.Get("val", DocX.w.NamespaceName), null);
					formatting.Script = (Script)Enum.Parse(typeof(Script), attribute);
					break;
				}
				}
			}
			return formatting;
		}

		public int CompareTo(object obj)
		{
			Formatting formatting = (Formatting)obj;
			if (formatting.hidden != hidden)
			{
				return -1;
			}
			if (formatting.bold != bold)
			{
				return -1;
			}
			if (formatting.italic != italic)
			{
				return -1;
			}
			if (formatting.strikethrough != strikethrough)
			{
				return -1;
			}
			if (formatting.script != script)
			{
				return -1;
			}
			if (formatting.highlight != highlight)
			{
				return -1;
			}
			if (formatting.size != size)
			{
				return -1;
			}
			if (formatting.fontColor != fontColor)
			{
				return -1;
			}
			if (formatting.underlineColor != underlineColor)
			{
				return -1;
			}
			if (formatting.underlineStyle != underlineStyle)
			{
				return -1;
			}
			if (formatting.misc != misc)
			{
				return -1;
			}
			if (formatting.capsStyle != capsStyle)
			{
				return -1;
			}
			if (formatting.fontFamily != fontFamily)
			{
				return -1;
			}
			if (formatting.percentageScale != percentageScale)
			{
				return -1;
			}
			if (formatting.kerning != kerning)
			{
				return -1;
			}
			if (formatting.position != position)
			{
				return -1;
			}
			if (formatting.spacing != spacing)
			{
				return -1;
			}
			if (!formatting.language.Equals(language))
			{
				return -1;
			}
			return 0;
		}
	}
}
