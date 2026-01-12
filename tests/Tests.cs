using System;
using System.Collections.Generic;
using excelchop;
using NUnit.Framework;

namespace tests;

[TestFixture]
public class Tests
{

    [Test]
    public void TestJoin()
    {
        List<string> TestString = new List<string>()
        {
            "Hello", string.Empty, "Third Item"
        };

        var joined = string.Join('\t', TestString);

        Assert.AreEqual(3, joined.Split('\t', StringSplitOptions.None).Length);

    }

    [Test]
    public void TestEndingWithSingleNewLine()
    {
        var testString = "Hello\n";
        var expected = "Hello\n";

        Assert.AreEqual(expected, testString.EndWithSingleNewline());

        testString = "Hello\n\n";
        expected = "Hello\n";

        Assert.AreEqual(expected, testString.EndWithSingleNewline());

        testString = null;
        expected = "";

        Assert.AreEqual(expected, testString.EndWithSingleNewline());

        testString = "Hello";
        expected = "Hello\n";

        Assert.AreEqual(expected, testString.EndWithSingleNewline());
    }

    [Test]
    public void TestSigFigs()
    {
        double testDouble = 10.23030234;
        int testSigFigs = 3;
        string expected = "10.2";

        Assert.AreEqual(expected, testDouble.ToSigFigs(testSigFigs));

        testDouble = 2.000000001;
        testSigFigs = 3;
        expected = "2.00";
        Assert.AreEqual(expected, testDouble.ToSigFigs(testSigFigs));

        testDouble = 123456789;
        testSigFigs = 3;
        expected = "123456789";
        Assert.AreEqual(expected, testDouble.ToSigFigs(testSigFigs));

        testDouble = 100000000;
        testSigFigs = 3;
        expected = "100000000";
        Assert.AreEqual(expected, testDouble.ToSigFigs(testSigFigs));
    }

    [Test]
    public void TestCsvEscape()
    {
        // No escaping needed for simple strings
        Assert.AreEqual("Hello", "Hello".CsvEscape());
        Assert.AreEqual("", "".CsvEscape());

        // Fields with commas must be quoted
        Assert.AreEqual("\"Hello, World\"", "Hello, World".CsvEscape());

        // Fields with double quotes must be quoted and quotes doubled
        Assert.AreEqual("\"Say \"\"Hello\"\"\"", "Say \"Hello\"".CsvEscape());

        // Fields with newlines must be quoted
        Assert.AreEqual("\"Line1\nLine2\"", "Line1\nLine2".CsvEscape());

        // Fields with carriage returns must be quoted
        Assert.AreEqual("\"Line1\r\nLine2\"", "Line1\r\nLine2".CsvEscape());

        // Combined: comma and quote
        Assert.AreEqual("\"Hello, \"\"World\"\"\"", "Hello, \"World\"".CsvEscape());
    }
}
