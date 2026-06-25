using System.Text;

namespace MnBEditor.Core;

/// <summary>
/// Low-level file parser for Mount & Blade TXT data files.
/// Ported from FileParser.bas. Uses Span&lt;byte&gt; for zero-allocation parsing.
///
/// The game's TXT format uses space-separated tokens. Strings containing spaces
/// are encoded with underscores (_) instead of spaces.
/// </summary>
public class FileParser
{
    private readonly byte[] _buffer;
    private int _position;
    private readonly int _length;

    public FileParser(string filePath)
    {
        _buffer = File.ReadAllBytes(filePath);
        _position = 0;
        _length = _buffer.Length;
    }

    public FileParser(byte[] buffer)
    {
        _buffer = buffer;
        _position = 0;
        _length = buffer.Length;
    }

    public int Position => _position;
    public int Length => _length;
    public bool IsEndOfFile => _position >= _length;

    /// <summary>
    /// Reads the next whitespace-delimited word from the buffer.
    /// Equivalent to VB6 GetWord().
    /// </summary>
    public string GetWord()
    {
        var sb = new StringBuilder();

        while (_position < _length)
        {
            byte b = _buffer[_position];
            _position++;

            if (b is 10 or 13 or 32) // LF, CR, space
            {
                if (sb.Length > 0)
                    break;
            }
            else
            {
                sb.Append((char)b);
            }
        }

        return sb.ToString();
    }

    /// <summary>
    /// Reads the next word as a long integer. Returns 0 if not a valid number.
    /// </summary>
    public long GetLong()
    {
        var word = GetWord();
        return long.TryParse(word, out var val) ? val : 0;
    }

    /// <summary>
    /// Reads the next word as a double. Returns 0.0 if not a valid number.
    /// </summary>
    public double GetDouble()
    {
        var word = GetWord();
        return double.TryParse(word,
            System.Globalization.NumberStyles.Float,
            System.Globalization.CultureInfo.InvariantCulture,
            out var val) ? val : 0.0;
    }

    /// <summary>
    /// Reads until the next CR/LF, returning the accumulated text.
    /// The CR+LF terminator is consumed.
    /// Equivalent to VB6 GetRealLine().
    /// </summary>
    public string GetLine()
    {
        var sb = new StringBuilder();

        while (_position < _length)
        {
            byte b = _buffer[_position];
            _position++;

            if (b == 13) // CR
            {
                // Consume LF if present
                if (_position < _length && _buffer[_position] == 10)
                    _position++;
                break;
            }
            else if (b == 10) // LF (bare, no CR)
            {
                break;
            }
            else
            {
                sb.Append((char)b);
            }
        }

        return sb.ToString();
    }

    /// <summary>
    /// Reads a complete line as a string, trimming trailing whitespace.
    /// </summary>
    public string GetLineTrimmed()
    {
        return GetLine().Trim();
    }

    /// <summary>
    /// Converts a large numeric string (64-bit game ID truncated format)
    /// to a signed 32-bit integer.
    /// Equivalent to VB6 MinusIT().
    /// </summary>
    public static long MinusIt(string num)
    {
        var n2 = num.Length > 9 ? num[^9..] : num;
        long num2 = long.Parse(n2);
        // "2000000000" - 2147483647 - 1 + num2
        return 2000000000L - 2147483647L - 1 + num2;
    }

    /// <summary>
    /// Skips a specified number of words (tokens).
    /// </summary>
    public void SkipWords(int count)
    {
        for (int i = 0; i < count; i++)
            GetWord();
    }

    /// <summary>
    /// Peeks at the next byte without advancing position.
    /// </summary>
    public byte PeekByte()
    {
        if (_position >= _length) return 0;
        return _buffer[_position];
    }
}
