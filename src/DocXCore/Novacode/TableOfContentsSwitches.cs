using System;
using System.ComponentModel;

namespace Novacode
{
	[Flags]
	public enum TableOfContentsSwitches
	{
		None = 0x0,
		[Description("\\a")]
		A = 0x1,
		[Description("\\b")]
		B = 0x2,
		[Description("\\c")]
		C = 0x4,
		[Description("\\d")]
		D = 0x8,
		[Description("\\f")]
		F = 0x10,
		[Description("\\h")]
		H = 0x20,
		[Description("\\l")]
		L = 0x40,
		[Description("\\n")]
		N = 0x80,
		[Description("\\o")]
		O = 0x100,
		[Description("\\p")]
		P = 0x200,
		[Description("\\s")]
		S = 0x400,
		[Description("\\t")]
		T = 0x800,
		[Description("\\u")]
		U = 0x1000,
		[Description("\\w")]
		W = 0x2000,
		[Description("\\x")]
		X = 0x4000,
		[Description("\\z")]
		Z = 0x8000
	}
}
