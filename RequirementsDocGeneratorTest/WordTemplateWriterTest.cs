using Microsoft.VisualStudio.TestTools.UnitTesting;
using RequirementsDocGenerator;

namespace RequirementsDocGenerator
{
    [TestClass]
    public class WordTemplateWriterTest 
    {
        [TestMethod]
        public void StartDocumentFromTemplateTest(){
            using(var reader = new WordTemplateWriter()){
                reader.StartDocumentFromTemplate("../../../../data/template.dotx", "../../../../data/output3.docx");
                reader.WriteTitle("Requirements");
            }
        }
    }
}