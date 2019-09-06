using System;
using System.Linq;
using System.Reflection;
using System.Xml.Linq;

namespace Novacode
{
	internal static class XElementHelpers
	{
		internal static T GetValueToEnum<T>(XElement element)
		{
			if (element == null)
			{
				throw new ArgumentNullException("element");
			}
			string value = element.Attribute(XName.Get("val")).Value;
			foreach (T value2 in Enum.GetValues(typeof(T)))
			{
				FieldInfo field = typeof(T).GetField(value2.ToString());
				if (field.GetCustomAttributes(typeof(XmlNameAttribute), inherit: false).Count() == 0)
				{
					throw new Exception($"Attribute 'XmlNameAttribute' is not assigned to {typeof(T).Name} fields!");
				}
				XmlNameAttribute xmlNameAttribute = (XmlNameAttribute)field.GetCustomAttributes(typeof(XmlNameAttribute), inherit: false).First();
				if (xmlNameAttribute.XmlName == value)
				{
					return value2;
				}
			}
			throw new ArgumentException("Invalid element value!");
		}

		internal static void SetValueFromEnum<T>(XElement element, T value)
		{
			if (element == null)
			{
				throw new ArgumentNullException("element");
			}
			element.Attribute(XName.Get("val")).Value = GetXmlNameFromEnum(value);
		}

		internal static string GetXmlNameFromEnum<T>(T value)
		{
			if (value == null)
			{
				throw new ArgumentNullException("value");
			}
			FieldInfo field = typeof(T).GetField(value.ToString());
			if (field.GetCustomAttributes(typeof(XmlNameAttribute), inherit: false).Count() == 0)
			{
				throw new Exception($"Attribute 'XmlNameAttribute' is not assigned to {typeof(T).Name} fields!");
			}
			XmlNameAttribute xmlNameAttribute = (XmlNameAttribute)field.GetCustomAttributes(typeof(XmlNameAttribute), inherit: false).First();
			return xmlNameAttribute.XmlName;
		}
	}
}
