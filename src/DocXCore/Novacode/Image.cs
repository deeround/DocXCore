using System;
using System.IO;
using System.IO.Packaging;

namespace Novacode
{
	public class Image
	{
		private string id;

		private DocX document;

		internal PackageRelationship pr;

		public string Id => id;

		public string FileName => Path.GetFileName(pr.TargetUri.ToString());

		public Stream GetStream(FileMode mode, FileAccess access)
		{
			string originalString = pr.SourceUri.OriginalString;
			string text = originalString.Remove(originalString.LastIndexOf('/'));
			string originalString2 = pr.TargetUri.OriginalString;
			string uriString = originalString2.Contains(text) ? originalString2 : (text + "/" + originalString2);
			return new PackagePartStream(document.package.GetPart(new Uri(uriString, UriKind.Relative)).GetStream(mode, access));
		}

		internal Image(DocX document, PackageRelationship pr)
		{
			this.document = document;
			this.pr = pr;
			id = pr.Id;
		}

		public Picture CreatePicture()
		{
			return Paragraph.CreatePicture(document, id, string.Empty, string.Empty);
		}

		public Picture CreatePicture(int height, int width)
		{
			Picture picture = Paragraph.CreatePicture(document, id, string.Empty, string.Empty);
			picture.Height = height;
			picture.Width = width;
			return picture;
		}
	}
}
