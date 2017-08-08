using System.IO;
using Stimulsoft.Report;

namespace StimulsoftTest
{
	public class Program
	{
		static void Main(string[] args)
		{
			Stimulsoft.Base.StiLicense.Key = "right license here";

			if (File.Exists("PageBreaks.docx"))
				File.Delete("PageBreaks.docx");

			if (File.Exists("AllowHtmlTags.docx"))
				File.Delete("AllowHtmlTags.docx");

			StiReport report = new StiReport();

			report.Load("PageBreaks.mrt");

			report.Render(false);

			using (var stream = new MemoryStream())
			{
				report.ExportDocument(StiExportFormat.Word2007, stream);

				using (var fileStream = File.Create("PageBreaks.docx"))
				{
					stream.Seek(0, SeekOrigin.Begin);
					stream.CopyTo(fileStream);
				}
			}

			StiReport report2 = new StiReport();

			report2.Load("AllowHtmlTags.mrt");
			report2.Render(false);

			using (var stream = new MemoryStream())
			{
				report2.ExportDocument(StiExportFormat.Word2007, stream);

				using (var fileStream = File.Create("AllowHtmlTags.docx"))
				{
					stream.Seek(0, SeekOrigin.Begin);
					stream.CopyTo(fileStream);
				}
			}
		}
	}
}
