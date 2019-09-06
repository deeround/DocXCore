using System;

namespace Novacode
{
	public class FormattedText : IComparable
	{
		public int index;

		public string text;

		public Formatting formatting;

		public int CompareTo(object obj)
		{
			FormattedText formattedText = (FormattedText)obj;
			if (formattedText.formatting == null || formatting == null)
			{
				return -1;
			}
			return formatting.CompareTo(formattedText.formatting);
		}
	}
}
