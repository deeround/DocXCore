using System;
using System.ComponentModel;
using System.Reflection;

namespace Novacode
{
	public static class ExtensionsHeadings
	{
		public static Paragraph Heading(this Paragraph paragraph, HeadingType headingType)
		{
			string text2 = paragraph.StyleName = headingType.EnumDescription();
			return paragraph;
		}

		public static string EnumDescription(this Enum enumValue)
		{
			if (enumValue == null || enumValue.ToString() == "0")
			{
				return string.Empty;
			}
			FieldInfo field = enumValue.GetType().GetField(enumValue.ToString());
			DescriptionAttribute[] array = (DescriptionAttribute[])field.GetCustomAttributes(typeof(DescriptionAttribute), inherit: false);
			if (array.Length != 0)
			{
				return array[0].Description;
			}
			return enumValue.ToString();
		}

		public static bool HasFlag(this Enum variable, Enum value)
		{
			if (variable == null)
			{
				return false;
			}
			if (value == null)
			{
				throw new ArgumentNullException("value");
			}
			if (!Enum.IsDefined(variable.GetType(), value))
			{
				throw new ArgumentException($"Enumeration type mismatch.  The flag is of type '{value.GetType()}', was expecting '{variable.GetType()}'.");
			}
			ulong num = Convert.ToUInt64(value);
			return (Convert.ToUInt64(variable) & num) == num;
		}
	}
}
