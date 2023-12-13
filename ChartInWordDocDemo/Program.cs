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

    public static void CreateDocument(Stream outputStream)
    {
        using var document = WordprocessingDocument.CreateFromTemplate("Template.dotx");
        //ChartCreator.CreateChart(document.MainDocumentPart);
        document.Clone(outputStream);
    }

}