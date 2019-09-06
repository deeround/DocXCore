using System.Drawing;

namespace Novacode
{
	public class Border
	{
		public BorderStyle Tcbs
		{
			get;
			set;
		}

		public BorderSize Size
		{
			get;
			set;
		}

		public int Space
		{
			get;
			set;
		}

		public Color Color
		{
			get;
			set;
		}

		public Border()
		{
			Tcbs = BorderStyle.Tcbs_single;
			Size = BorderSize.one;
			Space = 0;
			Color = Color.Black;
		}

		public Border(BorderStyle tcbs, BorderSize size, int space, Color color)
		{
			Tcbs = tcbs;
			Size = size;
			Space = space;
			Color = color;
		}
	}
}
