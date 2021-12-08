using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace RequirementsDocGenerator
{
    public class WordTemplateWriter: WordWriter 
    {
        public void StartDocumentFromTemplate(string templateFileName, string outputFile)
        {
            // create from template
            using (var temp = WordprocessingDocument.CreateFromTemplate(templateFileName))
            {
                WordprocessingDocument = temp.Clone(outputFile) as WordprocessingDocument;
            }
            MainDocumentPart = WordprocessingDocument.MainDocumentPart;
            // TODO if null - add it.
            //MainDocumentPart = WordprocessingDocument.AddMainDocumentPart();
            //MainDocumentPart.Document = new Document();

            Body = MainDocumentPart.Document.Body;
            //Body = MainDocumentPart.Document.AppendChild(new Body());
            //AddStyles();
            
            // list styles
            FindStyles();
            RemoveTitles();
        }

        public new void WriteTitle(string title){
            // find the title and replace it
            Paragraph titlePar = 
                Body.ChildElements
                .OfType<Paragraph>()
                .FirstOrDefault(par => (par?.ParagraphProperties
                                            .OfType<ParagraphStyleId>()
                                            .Any(id => id.Val == TitleStyle.StyleId)
                                ?? false)
                            && (par?.ChildElements.OfType<Run>().Any()
                                ?? false));
            var run = titlePar.ChildElements
                                .OfType<Run>()
                                .FirstOrDefault();
            titlePar.ReplaceChild(new Run(new Text(title)), run);

        }


        private void FindStyles(){
            // Title
            var stylesPart = MainDocumentPart.StyleDefinitionsPart;
            // TODO if null - create them
            foreach(Style st in stylesPart.Styles.Elements<Style>()){
                var id = st.StyleId;
                if (id.Value == "Normal") {
                    NormalStyle = st;
                }
                else if (id.Value == "Hyperlink") {
                    HyperLinkStyle = st;
                }
                else if (id.Value == "TableGrid") {
                    TableBodyStyle = st;
                    TableHeaderStyle = st; // TODO different style
                }
                else if (id.Value == "Title") {
                    TitleStyle = st;
                }
                else if (id.Value == "Heading1") {
                    Heading1Style = st;
                }
                else if (id.Value == "Heading2") {
                    Heading2Style = st;
                    // TODO Heading 3 and 4
                }
                else if (id.Value == "Header") {
                    PageHeaderStyle = st;
                }
                else if (id.Value == "Footer") {
                    PageFooterStyle = st;
                }

            }
            // find numberings
            var numPart = MainDocumentPart.NumberingDefinitionsPart;

            // numeration id from heading
            var numId = Heading1Style.StyleParagraphProperties
                                    .ChildElements
                                    .OfType<NumberingProperties>()
                                    .FirstOrDefault()?.NumberingId?.Val;
            if(numId != null) {
                this.CurrentNumInst = numPart.Numbering.ChildElements.OfType<NumberingInstance>()
                                                .FirstOrDefault(n => n.NumberID == numId);
                                            
            }
        }

        private void RemoveTitles(){
            foreach(Paragraph par in Body.ChildElements.OfType<Paragraph>()){
                bool isTitle = par.ParagraphProperties.ChildElements
                                    .OfType<ParagraphStyleId>()
                                    .Any(id => id.Val == Heading1Style.StyleId);
                if(isTitle){
                    par.Remove();
                }
            }
        }

    }
}