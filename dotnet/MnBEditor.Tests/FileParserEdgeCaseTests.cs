using MnBEditor.Core;

namespace MnBEditor.Tests;

public class FileParserEdgeCaseTests
{
    [Fact]
    public void GetWord_EmptyFile_ReturnsEmptyString()
    {
        var parser = new FileParser(Array.Empty<byte>());
        Assert.Equal("", parser.GetWord());
        Assert.True(parser.IsEndOfFile);
    }

    [Fact]
    public void GetWord_OnlyWhitespace_ReturnsEmpty()
    {
        var parser = new FileParser("   \r\n   "u8.ToArray());
        Assert.Equal("", parser.GetWord());
    }

    [Fact]
    public void GetWord_LeadingWhitespace_SkipsAndReturnsToken()
    {
        var parser = new FileParser("   hello"u8.ToArray());
        Assert.Equal("hello", parser.GetWord());
    }

    [Fact]
    public void GetWord_MultipleLines_ReadsAcrossCrLf()
    {
        var parser = new FileParser("token1\r\ntoken2\r\ntoken3"u8.ToArray());
        Assert.Equal("token1", parser.GetWord());
        Assert.Equal("token2", parser.GetWord());
        Assert.Equal("token3", parser.GetWord());
    }

    [Fact]
    public void GetWord_NumericStrings_PreservedAsStrings()
    {
        var parser = new FileParser("41876193280 9223388529554358287"u8.ToArray());
        Assert.Equal("41876193280", parser.GetWord());
        Assert.Equal("9223388529554358287", parser.GetWord());
    }

    [Fact]
    public void GetLong_ValidNumber_ReturnsParsedValue()
    {
        var parser = new FileParser("42 0 -1 806"u8.ToArray());
        Assert.Equal(42, parser.GetLong());
        Assert.Equal(0, parser.GetLong());
        Assert.Equal(-1, parser.GetLong());
        Assert.Equal(806, parser.GetLong());
    }

    [Fact]
    public void GetLong_InvalidInput_ReturnsZero()
    {
        var parser = new FileParser("not_a_number itm_sword"u8.ToArray());
        Assert.Equal(0, parser.GetLong());
        Assert.Equal(0, parser.GetLong()); // "itm_sword" also -> 0
    }

    [Fact]
    public void GetDouble_DecimalString_ReturnsParsedValue()
    {
        var parser = new FileParser("1.500000 3.0 0.25"u8.ToArray());
        Assert.Equal(1.5, parser.GetDouble(), 6);
        Assert.Equal(3.0, parser.GetDouble(), 6);
        Assert.Equal(0.25, parser.GetDouble(), 6);
    }

    [Fact]
    public void GetLine_ReadsUntilCrLf_ConsumesCrLf()
    {
        var parser = new FileParser("line1\r\nline2\r\n"u8.ToArray());
        Assert.Equal("line1", parser.GetLine());
        Assert.Equal("line2", parser.GetLine());
        Assert.Equal("", parser.GetLine()); // EOF
    }

    [Fact]
    public void GetWord_UnderscorePreserved_ForMBSpaceEncoding()
    {
        // M&B uses underscores in place of spaces in strings
        var parser = new FileParser("Test_Sword INVALID_ITEM"u8.ToArray());
        Assert.Equal("Test_Sword", parser.GetWord());
        Assert.Equal("INVALID_ITEM", parser.GetWord());
    }

    [Fact]
    public void SkipWords_SkipsCorrectCount()
    {
        var parser = new FileParser("a b c d e f"u8.ToArray());
        parser.SkipWords(3);
        Assert.Equal("d", parser.GetWord());
        Assert.Equal("e", parser.GetWord());
    }

    [Fact]
    public void Position_TracksCorrectly()
    {
        var parser = new FileParser("hello world 123"u8.ToArray());
        Assert.Equal(0, parser.Position);
        parser.GetWord(); // "hello"
        Assert.True(parser.Position > 0);
        parser.GetWord(); // "world"
        parser.GetWord(); // "123"
        Assert.Equal(parser.Length, parser.Position);
    }
}
