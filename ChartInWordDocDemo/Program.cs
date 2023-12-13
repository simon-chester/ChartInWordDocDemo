using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;

namespace ChartInWordDocDemo;

public class Program
{
    public static void Main()
    {
        using var memoryStream = new MemoryStream();
        CreateDocument(memoryStream);
        memoryStream.Position = 0;
        memoryStream.CopyTo(File.Create("Output.docx"));
    }

    private static void CreateDocument(Stream outputStream)
    {
        using var document = WordprocessingDocument.CreateFromTemplate("Template.dotx");
        CreateChart(document.MainDocumentPart); // Commenting this line will make the code work
        document.Clone(outputStream);
    }

    private static void CreateChart(MainDocumentPart mainPart)
    {
        var chartPart = mainPart.AddNewPart<ChartPart>("rId110");
        var chartSpace = new ChartSpace();
        var chart = new DocumentFormat.OpenXml.Drawing.Chart();
        chartSpace.Append(chart);
        chartPart.ChartSpace = chartSpace;
    }
}