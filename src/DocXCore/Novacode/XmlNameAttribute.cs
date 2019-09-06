using System;

namespace Novacode
{
	[AttributeUsage(AttributeTargets.Field, Inherited = false, AllowMultiple = false)]
	internal sealed class XmlNameAttribute : Attribute
	{
		public string XmlName
		{
			get;
			private set;
		}

		public XmlNameAttribute(string xmlName)
		{
			XmlName = xmlName;
		}
	}
}
