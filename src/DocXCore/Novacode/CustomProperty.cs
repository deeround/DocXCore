using System;
using System.Globalization;

namespace Novacode
{
	public class CustomProperty
	{
		public string Name
		{
			get;
		}

		public object Value
		{
			get;
		}

		internal string Type
		{
			get;
		}

		internal CustomProperty(string name, string type, string value)
		{
			object obj;
			switch (type)
			{
			case "lpwstr":
				obj = value;
				break;
			case "i4":
				obj = int.Parse(value, CultureInfo.InvariantCulture);
				break;
			case "r8":
				obj = double.Parse(value, CultureInfo.InvariantCulture);
				break;
			case "filetime":
				obj = DateTime.Parse(value, CultureInfo.InvariantCulture);
				break;
			case "bool":
				obj = bool.Parse(value);
				break;
			default:
				throw new Exception();
			}
			Name = name;
			Type = type;
			Value = obj;
		}

		private CustomProperty(string name, string type, object value)
		{
			Name = name;
			Type = type;
			Value = value;
		}

		public CustomProperty(string name, string value)
			: this(name, "lpwstr", (object)value)
		{
		}

		public CustomProperty(string name, int value)
			: this(name, "i4", value)
		{
		}

		public CustomProperty(string name, double value)
			: this(name, "r8", value)
		{
		}

		public CustomProperty(string name, DateTime value)
			: this(name, "filetime", value.ToUniversalTime())
		{
		}

		public CustomProperty(string name, bool value)
			: this(name, "bool", value)
		{
		}
	}
}
