using System;

namespace Novacode
{
	public sealed class Font
	{
		public string Name
		{
			get;
			private set;
		}

		public Font(string name)
		{
			if (string.IsNullOrEmpty(name))
			{
				throw new ArgumentNullException("name");
			}
			Name = name;
		}

		public override string ToString()
		{
			return Name;
		}
	}
}
