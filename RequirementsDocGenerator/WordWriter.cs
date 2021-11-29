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
    /// <summary>
    /// To write a Word Document using Open XML
    /// </summary>
    public class WordWriter : IDisposable
    {
        private WordprocessingDocument WordprocessingDocument;
        private MainDocumentPart MainDocumentPart;
        private Body Body;

        /// <summary>
        /// Inits the documents - creates styles parts and common parts
        /// </summary>
        /// <param name="fileName"></param>
        public void StartDocument(string fileName)
        {
            WordprocessingDocument = WordprocessingDocument.Create(fileName, WordprocessingDocumentType.Document);

            MainDocumentPart = WordprocessingDocument.AddMainDocumentPart();
            MainDocumentPart.Document = new Document();

            Body = MainDocumentPart.Document.AppendChild(new Body());
            AddStyles();
        }

        /// <summary>
        /// Inits the documents - creates styles parts and common parts
        /// Using the stream
        /// </summary>
        /// <param name="fileName"></param>
        public void StartDocument(Stream stream)
        {
            WordprocessingDocument = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);

            MainDocumentPart = WordprocessingDocument.AddMainDocumentPart();
            MainDocumentPart.Document = new Document();

            Body = MainDocumentPart.Document.AppendChild(new Body());
            AddStyles();
        }

        /// <summary>
        /// Styles
        /// </summary>
        private Style NormalStyle, HyperLinkStyle, TableBodyStyle, TableHeaderStyle, TitleStyle, Heading1Style, Heading2Style, PageFooterStyle, PageHeaderStyle;

        /// <summary>
        /// Creates and populates styles parts
        /// </summary>
        private void AddStyles()
        {
            // get or create styles part
            var stylesPart = MainDocumentPart.StyleDefinitionsPart;
            if (stylesPart == null)
            {
                stylesPart = MainDocumentPart.AddNewPart<StyleDefinitionsPart>();
                stylesPart.Styles = new Styles();

                // body style
                NormalStyle = stylesPart.Styles.AppendChild(
                    new Style(new StyleName() { Val = "Normal" },
                        new NextParagraphStyle() { Val = "normal" },
                        new StyleParagraphProperties(
                            new SpacingBetweenLines() { Line = "300", LineRule = LineSpacingRuleValues.Auto }
                        ),
                        new StyleRunProperties(new RunFonts() { Ascii = "Helvetica", HighAnsi = "Helvetica" },
                            new FontSize() { Val = "22" }
                        )
                    )
                    { Type = StyleValues.Paragraph, StyleId = "normal", CustomStyle = true });

                // hyperlink
                HyperLinkStyle = stylesPart.Styles.AppendChild(
                    new Style(
                        new StyleName() { Val = "Hyperlink" },
                        new BasedOn() { Val = NormalStyle.StyleId },
                        new NextParagraphStyle() { Val = NormalStyle.StyleId },
                        new StyleParagraphProperties(),
                        new StyleRunProperties(
                            new Underline() { Val = UnderlineValues.Single },
                            new Color() { Val = "777777" }
                        )
                    )
                    { Type = StyleValues.Paragraph, StyleId = "hyperlink", CustomStyle = true });

                // table body
                TableBodyStyle = stylesPart.Styles.AppendChild(
                    new Style(
                        new StyleName() { Val = "Table Body" },
                        new BasedOn() { Val = NormalStyle.StyleId },
                        new NextParagraphStyle() { Val = "tablebody" },
                        new StyleParagraphProperties(
                            new SpacingBetweenLines() { Line = "200", After = "80", LineRule = LineSpacingRuleValues.Exact }
                        ),
                        new StyleRunProperties(
                            new FontSize() { Val = "18" }
                        )
                    )
                    { Type = StyleValues.Paragraph, StyleId = "tablebody", CustomStyle = true });

                // table header
                TableHeaderStyle = stylesPart.Styles.AppendChild(
                    new Style(
                        new StyleName() { Val = "Table Header" },
                        new BasedOn() { Val = TableBodyStyle.StyleId },
                        new NextParagraphStyle() { Val = "tableheader" },
                        new StyleParagraphProperties(
                            new Justification() { Val = JustificationValues.Center }
                        ),
                        new StyleRunProperties(
                            new Bold(),
                            new Color() { Val = "FFFFFF" }
                        )
                    )
                    { Type = StyleValues.Paragraph, StyleId = "tableheader", CustomStyle = true });


                // first page title
                TitleStyle = stylesPart.Styles.AppendChild(
                    new Style(
                        new StyleName() { Val = "Title" },
                        new BasedOn() { Val = NormalStyle.StyleId },
                        new NextParagraphStyle() { Val = NormalStyle.StyleId },
                        new StyleParagraphProperties(
                            new BottomBorder() { Color = "004C97", Val = BorderValues.Single, Size = 4, Space = 1 },
                            new Justification() { Val = JustificationValues.Center },
                            new SpacingBetweenLines { After = "700" }
                        ),
                        new StyleRunProperties(
                            new Bold(),
                            new Color() { Val = "004C97" },
                            new FontSize() { Val = "48" }
                        )
                    )
                    { Type = StyleValues.Paragraph, StyleId = "title", CustomStyle = true });


                // Heading1 style
                Heading1Style = stylesPart.Styles.AppendChild(
                    new Style(
                        new StyleName() { Val = "Heading 1" },
                        new BasedOn() { Val = NormalStyle.StyleId },
                        new NextParagraphStyle() { Val = NormalStyle.StyleId },
                        new StyleParagraphProperties(
                            new BottomBorder() { Color = "004C97", Val = BorderValues.Single, Size = 4, Space = 2 },
                            new SpacingBetweenLines { After = "400", Before = "600" },
                            new OutlineLevel { Val = 1 }
                        ),
                        new StyleRunProperties(
                            //new Bold(),
                            new Color() { Val = "004C97" },
                            new FontSize() { Val = "32" }
                        )
                    )
                    { Type = StyleValues.Paragraph, StyleId = "heading1", CustomStyle = true });


                // Heading2 style
                Heading2Style = stylesPart.Styles.AppendChild(
                    new Style(
                        new StyleName() { Val = "Heading 2" },
                        new BasedOn() { Val = Heading1Style.StyleId },
                        new NextParagraphStyle() { Val = NormalStyle.StyleId },
                        new StyleParagraphProperties(
                            new SpacingBetweenLines { After = "300", Before = "400" },
                            new OutlineLevel { Val = 2 }
                        ),
                        new StyleRunProperties(
                            new FontSize() { Val = "22" }
                        )
                    )
                    { Type = StyleValues.Paragraph, StyleId = "heading2", CustomStyle = true });

                // Page header
                PageHeaderStyle = stylesPart.Styles.AppendChild(
                    new Style(
                        new StyleName() { Val = "Header" },
                        new BasedOn() { Val = NormalStyle.StyleId },
                        new NextParagraphStyle() { Val = NormalStyle.StyleId },
                        new StyleParagraphProperties(
                            new TopBorder() { Color = "004C97", Val = BorderValues.Single, Size = 4, Space = 2 }
                        ),
                        new StyleRunProperties(
                            new Color() { Val = "004C97" },
                            new FontSize() { Val = "15" }
                        )
                    )
                    { Type = StyleValues.Paragraph, StyleId = "header", CustomStyle = true });

                // Page footer
                PageFooterStyle = stylesPart.Styles.AppendChild(
                    new Style(
                        new StyleName() { Val = "Footer" },
                        new BasedOn() { Val = PageHeaderStyle.StyleId },
                        new NextParagraphStyle() { Val = NormalStyle.StyleId },
                        new StyleParagraphProperties(
                        ),
                        new StyleRunProperties(
                        )
                    )
                    { Type = StyleValues.Paragraph, StyleId = "footer", CustomStyle = true });
            }

            var numPart = MainDocumentPart.NumberingDefinitionsPart;
            if (numPart == null)
            {
                numPart = MainDocumentPart.AddNewPart<NumberingDefinitionsPart>();
                numPart.Numbering = new Numbering();

                var currentAbstractNum = numPart.Numbering.AppendChild(
                    new AbstractNum(
                        new Nsid() { Val = "099A081C" },
                        new MultiLevelType() { Val = MultiLevelValues.HybridMultilevel },
                        new Level(
                            new StartNumberingValue() { Val = 1 },
                            new LevelText() { Val = "%1." },
                            new LevelJustification() { Val = LevelJustificationValues.Left }
                        )
                        { LevelIndex = 0 },
                        new Level(
                            new StartNumberingValue() { Val = 1 },
                            new LevelText() { Val = "%1.%2." },
                            new LevelJustification() { Val = LevelJustificationValues.Left }
                        )
                        { LevelIndex = 1 }
                    )
                    { AbstractNumberId = 0 });

                CurrentNumInst = numPart.Numbering.AppendChild(
                    // same id of abstract numbering
                    new NumberingInstance(new AbstractNumId() { Val = currentAbstractNum.AbstractNumberId })
                    { NumberID = 1 }); // id used in the style
            }
        }

        /// <summary>
        /// Reference to current title numbering instance
        /// </summary>
        private NumberingInstance CurrentNumInst;

        /// <summary>
        /// Starts the Heading 1 and Heading 2 numberings.
        /// </summary>
        public void ResetHeadingNumbering()
        {
            var lastAbstractNum = MainDocumentPart.NumberingDefinitionsPart.
                Numbering.ChildElements.OfType<AbstractNum>().Last();
            var lastNumInst = MainDocumentPart.NumberingDefinitionsPart.
                Numbering.ChildElements.OfType<NumberingInstance>().Last();
            // new abstract numb
            var currentAbstractNum = MainDocumentPart.NumberingDefinitionsPart.Numbering.InsertAfter(
                    new AbstractNum(
                        new Nsid() { Val = Convert.ToString(new Random().Next(), 16) },
                        new MultiLevelType() { Val = MultiLevelValues.HybridMultilevel },
                        new Level(
                            new StartNumberingValue() { Val = 1 },
                            new LevelText() { Val = "%1." },
                            new LevelJustification() { Val = LevelJustificationValues.Left }
                        )
                        { LevelIndex = 0 },
                        new Level(
                            new StartNumberingValue() { Val = 1 },
                            new LevelText() { Val = "%1.%2." },
                            new LevelJustification() { Val = LevelJustificationValues.Left }
                        )
                        { LevelIndex = 1 }
                    )
                    { AbstractNumberId = lastAbstractNum.AbstractNumberId + 1 }, lastAbstractNum);
            // new instance
            CurrentNumInst = MainDocumentPart.NumberingDefinitionsPart.Numbering.InsertAfter(
                new NumberingInstance(new AbstractNumId() { Val = currentAbstractNum.AbstractNumberId })
                { NumberID = lastNumInst.NumberID + 1 },
                lastNumInst
            );
        }

        /// <summary>
        /// Adds an image
        /// </summary>
        /// <param name="imageFile"></param>
        /// <param name="type"></param>
        public void AddImage(string imageFile, string type)
        {
            var imagePart = MainDocumentPart.AddImagePart(ImagePartType.Png);
            using (var stream = new FileStream(imageFile, FileMode.Open, FileAccess.Read))
                imagePart.FeedData(stream);
            // get id
            string imageId = MainDocumentPart.GetIdOfPart(imagePart);

            // Define the reference of the image.
            var element = new Drawing(
                            new Inline(
                                new Extent() { Cx = 4306824L, Cy = 813816L },
                                new EffectExtent() { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L },
                                new DocProperties() { Id = 1U, Name = "Picture 1" },
                                new NonVisualGraphicFrameDrawingProperties(
                                    new A.GraphicFrameLocks() { NoChangeAspect = true }),
                                new A.Graphic(
                                    new A.GraphicData(
                                        new PIC.Picture(
                                            new PIC.NonVisualPictureProperties(
                                                new PIC.NonVisualDrawingProperties() { Id = 0, Name = "My Image" },
                                                new PIC.NonVisualPictureDrawingProperties()),
                                            new PIC.BlipFill(
                                                new A.Blip(
                                                    new A.BlipExtensionList(
                                                        new A.BlipExtension() { Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}" }
                                                    )
                                                )
                                                { Embed = imageId, CompressionState = A.BlipCompressionValues.Print },
                                                new A.Stretch(new A.FillRectangle())
                                        ),
                                            new PIC.ShapeProperties(
                                                new A.Transform2D(
                                                    new A.Offset() { X = 0L, Y = 0L },
                                                    new A.Extents() { Cx = 4306824L, Cy = 813816L }),
                                                new A.PresetGeometry(new A.AdjustValueList())
                                                { Preset = A.ShapeTypeValues.Rectangle })
                                        )
                                    )
                                    { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
                            )
                            { DistanceFromTop = 0, DistanceFromBottom = 0, DistanceFromLeft = 0, DistanceFromRight = 0, EditId = "50D07946" });
            // Append the reference to body, the element should be in a Run.
            Body.AppendChild(new Paragraph(new Run(element)));
        }

        /// <summary>
        /// Reference to current table
        /// </summary>
        private Table CurrentTable;

        /// <summary>
        /// Inits a table
        /// </summary>
        public void StartTable()
        {
            CurrentTable = new Table(
                new TableProperties(
                    new TableBorders(
                        new TopBorder() { Val = BorderValues.Single, Size = 6, Color = "CCCCCC" },
                        new InsideHorizontalBorder() { Val = BorderValues.Single, Size = 6, Color = "CCCCCC" },
                        new BottomBorder() { Val = BorderValues.Single, Size = 6, Color = "CCCCCC" },
                        new InsideVerticalBorder() { Val = BorderValues.Single, Size = 6, Color = "CCCCCC" },
                        new LeftBorder() { Val = BorderValues.Single, Size = 6, Color = "CCCCCC" },
                        new RightBorder() { Val = BorderValues.Single, Size = 6, Color = "CCCCCC" }
                    ),
                    new TableJustification() { Val = TableRowAlignmentValues.Center }
                )
            );
            Body.Append(CurrentTable);
        }


        /// <summary>
        /// Adds a row with header format to the table
        /// </summary>
        /// <param name="row"></param>
        public void AddTableHeader(params string[] row)
        {
            AddTableHeader(row.AsEnumerable());
        }

        /// <summary>
        /// Adds a row with header format to the table
        /// </summary>
        /// <param name="row"></param>
        public void AddTableHeader(IEnumerable<string> row)
        {
            var tRow = CurrentTable.AppendChild(new TableRow());
            foreach (string t in row)
                tRow.Append(
                    new TableCell(
                        new TableCellProperties(
                            new TableCellMargin(
                                new TopMargin() { Width = "120", Type = TableWidthUnitValues.Dxa },
                                new LeftMargin() { Width = "100", Type = TableWidthUnitValues.Dxa },
                                new RightMargin() { Width = "100", Type = TableWidthUnitValues.Dxa }
                            ),
                            new Shading() { Fill = "004C97" }
                        ),
                        new Paragraph(
                            new ParagraphProperties(new ParagraphStyleId() { Val = TableHeaderStyle.StyleId }),
                            new Run(new Text(t))
                        )
                    )
                );
        }
        /// <summary>
        /// Adds a table row
        /// </summary>
        /// <param name="row"></param>
        public void AddTableRow(params string[] row)
        {
            AddTableRow(row.AsEnumerable());
        }

        /// <summary>
        /// Adds a table row
        /// </summary>
        /// <param name="row"></param>
        public void AddTableRow(IEnumerable<string> row)
        {
            var tRow = CurrentTable.AppendChild(new TableRow());
            foreach (string t in row)
                tRow.Append(
                    new TableCell(
                        new TableCellProperties(
                            new TableCellMargin(
                                new LeftMargin() { Width = "120", Type = TableWidthUnitValues.Dxa },
                                new RightMargin() { Width = "80", Type = TableWidthUnitValues.Dxa },
                                new TopMargin() { Width = "60", Type = TableWidthUnitValues.Dxa }
                            )
                        ),
                        new Paragraph(
                            new ParagraphProperties(new ParagraphStyleId() { Val = TableBodyStyle.StyleId }),
                            new Run(new Text(t))
                        )
                    )
                );
        }

        /// <summary>
        /// First page title
        /// </summary>
        /// <param name="title"></param>
        public void WriteTitle(string title)
        {
            Body.AppendChild(
                new Paragraph(
                    new ParagraphProperties(new ParagraphStyleId() { Val = TitleStyle.StyleId }),
                    new Run(new Text(title))));
        }

        /// <summary>
        /// Big heading for major document parts
        /// </summary>
        /// <param name="title"></param>
        public void WriteHeading0(string title)
        {
            Body.AppendChild(
                new Paragraph(
                    new ParagraphProperties(new ParagraphStyleId() { Val = Heading1Style.StyleId }),
                    new Run(new Text(title))));
        }

        /// <summary>
        /// Numbered headings
        /// </summary>
        /// <param name="title"></param>
        public void WriteHeading1(string title)
        {
            Body.AppendChild(
                new Paragraph(
                    new ParagraphProperties(new ParagraphStyleId() { Val = Heading1Style.StyleId }),
                    new NumberingProperties(
                        new NumberingId() { Val = CurrentNumInst.NumberID }, // points to the numbering id instance
                        new NumberingLevelReference() { Val = 0 }// points to the level
                    ),
                    new Run(new Text(title))));
        }

        /// <summary>
        /// Numbered secondary headings
        /// </summary>
        /// <param name="title"></param>
        public void WriteHeading2(string title)
        {
            Body.AppendChild(
                new Paragraph(
                    new ParagraphProperties(new ParagraphStyleId() { Val = Heading2Style.StyleId }),
                    new NumberingProperties(
                        new NumberingId() { Val = CurrentNumInst.NumberID }, // points to the numbering id instance
                        new NumberingLevelReference() { Val = 1 }// points to the level
                    ),
                    new Run(new Text(title))));
        }

        /// <summary>
        /// Text on the page header
        /// </summary>
        /// <param name="header"></param>
        public void SetHeader(string header)
        {
            // Delete and create new one
            MainDocumentPart.DeleteParts(MainDocumentPart.HeaderParts);
            var headerPart = MainDocumentPart.AddNewPart<HeaderPart>();

            // id to redirect
            string rId = MainDocumentPart.GetIdOfPart(headerPart);

            headerPart.Header = new Header();
            var par = headerPart.Header.AppendChild(
                new Paragraph(
                    new ParagraphProperties(new ParagraphStyleId() { Val = PageHeaderStyle.StyleId }),
                    new Run(new Text(header))));

            // create / update section properties
            var sectPrs = Body.Elements<SectionProperties>();
            if (sectPrs.Count() == 0)
                Body.AppendChild(new SectionProperties())
                    .AppendChild(new HeaderReference() { Id = rId });
            // if no footer references, add
            else if (sectPrs.SelectMany(pt => pt.Elements<HeaderReference>()).Count() == 0)
                sectPrs.ToList().ForEach(pt =>
                    pt.AppendChild(new HeaderReference() { Type = HeaderFooterValues.Default, Id = rId }));
            else
                // remap all references
                sectPrs.SelectMany(pt => pt.Elements<HeaderReference>())
                    .ToList()
                    .ForEach(hr => hr.Id = rId);
        }

        /// <summary>
        /// Text on the page footer
        /// </summary>
        /// <param name="footer"></param>
        public void SetFooter(string footer)
        {
            // Delete and create new one
            MainDocumentPart.DeleteParts(MainDocumentPart.FooterParts);
            var footerPart = MainDocumentPart.AddNewPart<FooterPart>();

            // id to redirect
            string rId = MainDocumentPart.GetIdOfPart(footerPart);

            footerPart.Footer = new Footer();
            footerPart.Footer.AppendChild(
                new Paragraph(
                    new ParagraphProperties(new ParagraphStyleId() { Val = PageFooterStyle.StyleId }),
                    new Run(new Text(footer)),
                    new Run(new TabChar(), new TabChar(), new TabChar(), new TabChar(), new TabChar(), new TabChar(), new TabChar(), new TabChar(), new TabChar()),
                    new Run(new Text("               ") { Space = SpaceProcessingModeValues.Preserve }),
                    new Run(new SimpleField() { Instruction = "PAGE" })
                ));

            // create / update section properties
            var sectPrs = Body.Elements<SectionProperties>();
            // if no section properties - add
            if (sectPrs.Count() == 0)
                Body.AppendChild(new SectionProperties(
                                    new FooterReference() { Type = HeaderFooterValues.Default, Id = rId }));
            // if no footer references, add
            else if (sectPrs.SelectMany(pt => pt.Elements<FooterReference>()).Count() == 0)
                sectPrs.ToList().ForEach(pt => pt.AppendChild(
                    new FooterReference() { Type = HeaderFooterValues.Default, Id = rId }));
            else
                // remap all references
                sectPrs.SelectMany(pt => pt.Elements<FooterReference>())
                    .ToList()
                    .ForEach(hr => hr.Id = rId);
        }

        /// <summary>
        /// A paragraph. Line-breaks will be ignored (should be appended as extra paragraphs).
        /// </summary>
        /// <param name="text"></param>
        public Paragraph WriteParagraph(string text = "")
        {
            return Body.AppendChild(
                new Paragraph(
                    new ParagraphProperties(new ParagraphStyleId() { Val = NormalStyle.StyleId }),
                    new Run(new Text(text))));
        }


        /// <summary>
        /// Jumps to next page
        /// </summary>
        public void WritePageBreak()
        {
            Body.AppendChild(
                new Paragraph(
                    new Run(new Break() { Type = BreakValues.Page })));
        }

        private NumberingInstance BulletListNumInst;

        /// <summary>
        /// Starts a bullet list
        /// </summary>
        public void StartBulletList()
        {
            var lastAbstractNum = MainDocumentPart.NumberingDefinitionsPart.
                Numbering.ChildElements.OfType<AbstractNum>().Last();
            var lastNumInst = MainDocumentPart.NumberingDefinitionsPart.
                Numbering.ChildElements.OfType<NumberingInstance>().Last();
            // add abstract num and num instance to numbering parts
            var newAbstractNum = MainDocumentPart.NumberingDefinitionsPart.Numbering.InsertAfter(
                    new AbstractNum(
                        new Nsid() { Val = Convert.ToString(new Random().Next(), 16) },
                        new Level(
                            new NumberingFormat() { Val = NumberFormatValues.Bullet },
                            new LevelText() { Val = "·" },
                            new LevelJustification() { Val = LevelJustificationValues.Left },
                            new ParagraphProperties(
                                new Indentation() { Left = "720", Hanging = "360" }
                            ),
                            new RunProperties(
                                new RunFonts() { Ascii = "Symbol", HighAnsi = "Symbol", Hint = FontTypeHintValues.Default }
                            )
                        )
                        { LevelIndex = 0 }
                    )
                    { AbstractNumberId = lastAbstractNum.AbstractNumberId + 1 }, lastAbstractNum);

            BulletListNumInst = MainDocumentPart.NumberingDefinitionsPart.Numbering.InsertAfter(
                new NumberingInstance(new AbstractNumId() { Val = newAbstractNum.AbstractNumberId })
                { NumberID = lastNumInst.NumberID + 1 },
                lastNumInst
            );
        }


        /// <summary>
        /// Adds a new item to the bullet list.
        /// </summary>
        /// <param name="item"></param>
        public void AddBulletListItem(string item)
        {
            Body.Append(
                new Paragraph(
                    new ParagraphProperties(new ParagraphStyleId() { Val = NormalStyle.StyleId }),
                    new NumberingProperties(
                        new NumberingId() { Val = BulletListNumInst.NumberID }, // points to the numbering id instance
                        new NumberingLevelReference() { Val = 0 }// points to the level
                    ),
                    new Run(new Text(item))
                )
            );
        }


        /// <summary>
        /// Adds a new link to the bullet list
        /// </summary>
        /// <param name="item"></param>
        /// <param name="url"></param>
        public void AddBulletListLinkItem(string item, string url)
        {
            var link = MainDocumentPart.AddHyperlinkRelationship(new Uri(url), true);

            Body.AppendChild(
                   new Paragraph(
                       new ParagraphProperties(new ParagraphStyleId() { Val = HyperLinkStyle.StyleId }),
                       new NumberingProperties(
                           new NumberingId() { Val = BulletListNumInst.NumberID }, // points to the numbering id instance
                           new NumberingLevelReference() { Val = 0 }// points to the level
                       ),
                       // link
                       new Hyperlink(new Run(new Text(item))) { Id = link.Id }
                    )
                );
        }

        /// <summary>
        /// Starts a bullet list item and adds all the items
        /// </summary>
        /// <param name="list"></param>
        public void AddBulletList(IEnumerable<string> list)
        {
            StartBulletList();
            foreach (var line in list)
                AddBulletListItem(line);
        }

        private NumberingInstance NumberListInstance;

        /// <summary>
        /// Starts a number list
        /// </summary>
        public void StartNumberList()
        {
            var lastAbstractNum = MainDocumentPart.NumberingDefinitionsPart.
                Numbering.ChildElements.OfType<AbstractNum>().Last();
            var lastNumInst = MainDocumentPart.NumberingDefinitionsPart.
                Numbering.ChildElements.OfType<NumberingInstance>().Last();
            // add abstract num and num instance to numbering parts
            var newAbstractNum = MainDocumentPart.NumberingDefinitionsPart.Numbering.InsertAfter(
                    new AbstractNum(
                        new Nsid() { Val = Convert.ToString(new Random().Next(), 16) },
                        new MultiLevelType() { Val = MultiLevelValues.HybridMultilevel },
                        new Level(
                            new StartNumberingValue() { Val = 1 },
                            new LevelText() { Val = "%1." },
                            new LevelJustification() { Val = LevelJustificationValues.Left }
                        )
                        { LevelIndex = 0 }
                    )
                    { AbstractNumberId = lastAbstractNum.AbstractNumberId + 1 }, lastAbstractNum);

            NumberListInstance = MainDocumentPart.NumberingDefinitionsPart.Numbering.InsertAfter(
                new NumberingInstance(new AbstractNumId() { Val = newAbstractNum.AbstractNumberId })
                { NumberID = lastNumInst.NumberID + 1 },
                lastNumInst
            );
        }

        /// <summary>
        /// Adds a new item to the bullet list.
        /// </summary>
        /// <param name="item"></param>
        public void AddNumberListItem(string item)
        {
            Body.Append(
                new Paragraph(
                    new ParagraphProperties(new ParagraphStyleId() { Val = NormalStyle.StyleId }),
                    new NumberingProperties(
                        new NumberingId() { Val = NumberListInstance.NumberID }, // points to the numbering id instance
                        new NumberingLevelReference() { Val = 0 }// points to the level
                    ),
                    new Run(new Text(item))
                )
            );
        }

        private SdtBlock TableOfContentsBlock;

        /// <summary>
        /// Writes the table of contents - empty so far.
        /// </summary>
        public void WriteTOC()
        {
            var sdt = new SdtBlock(
                new SdtProperties(
                    new RunProperties(
                        new RunFonts() { Ascii = "Helvetica", HighAnsi = "Helvetica" },
                        new Color() { Val = "auto" },
                        new FontSize { Val = "22" },
                        new FontSizeComplexScript { Val = "22" }
                    ),
                    new SdtContentDocPartObject(
                        new DocPartGallery() { Val = "Table of Contents" },
                        new DocPartUnique()
                    )
                ),
                new SdtEndCharProperties(
                    new RunProperties(new Bold(), new BoldComplexScript(), new NoProof())
                ),
                new SdtContentBlock(
                    new Paragraph(
                        new ParagraphProperties(new ParagraphStyleId() { Val = Heading2Style.StyleId }),
                        new Run(new Text("TABLE OF CONTENTS"))
                    ),
                    new Paragraph(
                        new ParagraphProperties(
                            new ParagraphStyleId() { Val = NormalStyle.StyleId },
                            new Tabs(new TabStop() { Val = TabStopValues.Right, Leader = TabStopLeaderCharValues.Dot, Position = 9350 }),
                            new RunProperties(new NoProof())
                        ),
                        new Run(new FieldChar() { FieldCharType = FieldCharValues.Begin }),
                        new Run(new FieldCode(@" TOC \o '1-3' \h \z \u ") { Space = SpaceProcessingModeValues.Preserve }),
                        new Run(new FieldChar() { FieldCharType = FieldCharValues.Separate }),
                        new Run(new FieldChar() { FieldCharType = FieldCharValues.End })
                    )
                )
            );
            //Body.AppendChild(x);
            TableOfContentsBlock = Body.AppendChild(sdt);

            // settings
            var settingsPart = MainDocumentPart.DocumentSettingsPart;
            if (settingsPart == null)
            {
                settingsPart = MainDocumentPart.AddNewPart<DocumentSettingsPart>();
                settingsPart.Settings = new Settings(
                    new UpdateFieldsOnOpen() { Val = true }
                );
            }
        }

        public void UpdateTOC()
        {
            // TODO implement
        }

        public void Dispose()
        {
            if (WordprocessingDocument != null)
            {
                WordprocessingDocument.Close();
                WordprocessingDocument.Dispose();
                WordprocessingDocument = null;
                MainDocumentPart = null;
                Body = null;
            }
        }
    }


    public static class OpenXMLExtensions
    {
        /// <summary>
        /// A shortcut for adding a run inside a paragraph
        /// </summary>
        /// <param name="par"></param>
        /// <param name="text"></param>
        /// <param name="bold"></param>
        /// <param name="italic"></param>
        /// <returns></returns>
        public static Paragraph AppendRun(this Paragraph par, string text = "", bool bold = false, bool italic = false, bool preserveSpaces = false)
        {
            par.AppendChild(new Run(
                                    new RunProperties(new Bold() { Val = bold }, new Italic() { Val = italic }),
                                    new Text(text) { Space = preserveSpaces ? SpaceProcessingModeValues.Preserve : SpaceProcessingModeValues.Default }
                                ));
            return par;
        }
    }

}
