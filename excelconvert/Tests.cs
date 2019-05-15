using System;
using System.Collections.Generic;
using System.Text;
using NUnit.Framework;

namespace excelconvert
{
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
    }
}
