using System;
using RequirementsDocGenerator;
namespace RequirementsDocGenerator
{
    class Program
    {
        static void Main(string[] args)
        {
            // TODO use command lines
            var gen = new DocGenerator();
            string projectPath = @"C:\Users\jbadillo\Documents\workspace\RequirementsDocGenerator\RequirementsDocGenerator";
            gen.Generate($"{projectPath}/requirements.xlsx", $"{projectPath}/requirements.docx");
        }
    }
}
