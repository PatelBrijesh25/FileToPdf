using PdfSharp.Fonts;
using System.IO;

public class CustomFontResolver : IFontResolver
{
    public static readonly CustomFontResolver Instance = new CustomFontResolver();

    public string DefaultFontName => "Arial";

    public byte[] GetFont(string faceName)
    {
        var fontPath = Path.Combine(Directory.GetCurrentDirectory(), "Fonts", "arial.ttf");
        return File.ReadAllBytes(fontPath);
    }

    public FontResolverInfo ResolveTypeface(string familyName, bool isBold, bool isItalic)
    {
        return new FontResolverInfo("Arial");
    }
}
