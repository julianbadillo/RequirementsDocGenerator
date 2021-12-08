using Microsoft.VisualStudio.TestTools.UnitTesting;
using RequirementsDocGenerator;

[TestClass]
public class DocGeneratorTest
{
    [TestMethod]
    public void GenerateTest(){
        var gen = new DocGenerator();
        var input = "../../../../data/requirements_sample.xlsx";
        var output = "../../../../data/output.docx";
        gen.Generate(input, output);
    }

    [TestMethod]
    public void GenerateFromTemplateTest(){
        var gen = new DocGenerator();
        var input = "../../../../data/requirements_sample.xlsx";
        var template = "../../../../data/template.dotx";
        var output = "../../../../data/output2.docx";
        gen.Generate(input, output, template);
    }
}