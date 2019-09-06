using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO.Packaging;
using System.Xml.Linq;

namespace Novacode
{
	public class Footer : Container, IParagraphContainer
	{
		public bool PageNumbers
		{
			get
			{
				return false;
			}
			set
			{
				XElement content = XElement.Parse("<w:sdt xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>\r\n                    <w:sdtPr>\r\n                      <w:id w:val='157571950' />\r\n                      <w:docPartObj>\r\n                        <w:docPartGallery w:val='Page Numbers (Top of Page)' />\r\n                        <w:docPartUnique />\r\n                      </w:docPartObj>\r\n                    </w:sdtPr>\r\n                    <w:sdtContent>\r\n                      <w:p w:rsidR='008D2BFB' w:rsidRDefault='008D2BFB'>\r\n                        <w:pPr>\r\n                          <w:pStyle w:val='Header' />\r\n                          <w:jc w:val='center' />\r\n                        </w:pPr>\r\n                        <w:fldSimple w:instr=' PAGE \\* MERGEFORMAT'>\r\n                          <w:r>\r\n                            <w:rPr>\r\n                              <w:noProof />\r\n                            </w:rPr>\r\n                            <w:t>1</w:t>\r\n                          </w:r>\r\n                        </w:fldSimple>\r\n                      </w:p>\r\n                    </w:sdtContent>\r\n                  </w:sdt>");
				base.Xml.AddFirst(content);
			}
		}

		public override ReadOnlyCollection<Paragraph> Paragraphs
		{
			get
			{
				ReadOnlyCollection<Paragraph> paragraphs = base.Paragraphs;
				foreach (Paragraph item in paragraphs)
				{
					item.mainPart = mainPart;
				}
				return paragraphs;
			}
		}

		public override List<Table> Tables
		{
			get
			{
				List<Table> tables = base.Tables;
				tables.ForEach(delegate(Table x)
				{
					x.mainPart = mainPart;
				});
				return tables;
			}
		}

		internal Footer(DocX document, XElement xml, PackagePart mainPart)
			: base(document, xml)
		{
			base.mainPart = mainPart;
		}

		public override Paragraph InsertParagraph()
		{
			Paragraph paragraph = base.InsertParagraph();
			paragraph.PackagePart = mainPart;
			return paragraph;
		}

		public override Paragraph InsertParagraph(int index, string text, bool trackChanges)
		{
			Paragraph paragraph = base.InsertParagraph(index, text, trackChanges);
			paragraph.PackagePart = mainPart;
			return paragraph;
		}

		public override Paragraph InsertParagraph(Paragraph p)
		{
			p.PackagePart = mainPart;
			return base.InsertParagraph(p);
		}

		public override Paragraph InsertParagraph(int index, Paragraph p)
		{
			p.PackagePart = mainPart;
			return base.InsertParagraph(index, p);
		}

		public override Paragraph InsertParagraph(int index, string text, bool trackChanges, Formatting formatting)
		{
			Paragraph paragraph = base.InsertParagraph(index, text, trackChanges, formatting);
			paragraph.PackagePart = mainPart;
			return paragraph;
		}

		public override Paragraph InsertParagraph(string text)
		{
			Paragraph paragraph = base.InsertParagraph(text);
			paragraph.PackagePart = mainPart;
			return paragraph;
		}

		public override Paragraph InsertParagraph(string text, bool trackChanges)
		{
			Paragraph paragraph = base.InsertParagraph(text, trackChanges);
			paragraph.PackagePart = mainPart;
			return paragraph;
		}

		public override Paragraph InsertParagraph(string text, bool trackChanges, Formatting formatting)
		{
			Paragraph paragraph = base.InsertParagraph(text, trackChanges, formatting);
			paragraph.PackagePart = mainPart;
			return paragraph;
		}

		public override Paragraph InsertEquation(string equation)
		{
			Paragraph paragraph = base.InsertEquation(equation);
			paragraph.PackagePart = mainPart;
			return paragraph;
		}

		public new Table InsertTable(int rowCount, int columnCount)
		{
			if (rowCount < 1 || columnCount < 1)
			{
				throw new ArgumentOutOfRangeException("Row and Column count must be greater than zero.");
			}
			Table table = base.InsertTable(rowCount, columnCount);
			table.mainPart = mainPart;
			return table;
		}

		public new Table InsertTable(int index, Table t)
		{
			Table table = base.InsertTable(index, t);
			table.mainPart = mainPart;
			return table;
		}

		public new Table InsertTable(Table t)
		{
			t = base.InsertTable(t);
			t.mainPart = mainPart;
			return t;
		}

		public new Table InsertTable(int index, int rowCount, int columnCount)
		{
			if (rowCount < 1 || columnCount < 1)
			{
				throw new ArgumentOutOfRangeException("Row and Column count must be greater than zero.");
			}
			Table table = base.InsertTable(index, rowCount, columnCount);
			table.mainPart = mainPart;
			return table;
		}
	}
}
