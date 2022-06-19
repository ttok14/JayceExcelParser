using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.CodeAnalysis.CSharp;

namespace JayceExcelParser.Common
{
    class LexicalHelper
    {
        public static bool IsValidIdentifierName(string name)
        {
            return SyntaxFacts.IsValidIdentifier(name);
        }
    }
}
