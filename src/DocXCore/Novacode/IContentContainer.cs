using System.Collections.ObjectModel;

namespace Novacode
{
	internal interface IContentContainer
	{
		ReadOnlyCollection<Content> Paragraphs
		{
			get;
		}
	}
}
