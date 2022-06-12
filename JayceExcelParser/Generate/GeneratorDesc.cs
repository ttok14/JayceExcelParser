using System;
using System.Collections.Generic;
using System.Text;

namespace JayceExcelParser.Generate
{
    class GeneratorDesc
    {
        public string ExcelDirectory { get; private set; }
        public string OutputDirectory { get; private set; }

        public bool GenJson { get; private set; } = false;
        public bool GenMessagePack { get; private set; } = false;

        public bool IsOutputSpecified => GenJson || GenMessagePack;

        public GeneratorDesc(string excelDirectory, string outputDirectory)
        {
            this.ExcelDirectory = excelDirectory;
            this.OutputDirectory = outputDirectory;
        }

        public GeneratorDesc AddJsonOutput()
        {
            GenJson = true;
            return this;
        }

        public GeneratorDesc AddMessagePack()
        {
            GenMessagePack = true;
            return this;
        }
    }
}
