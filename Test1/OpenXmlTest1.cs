using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.IO;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using DocumentFormat.OpenXml;

namespace Test1
{
    class OpenXmlTest1
    {
        /// <summary>
        /// 表格中插入数据
        /// </summary>
        /// <param name="fileName"></param>
        /// <param name="data"></param>
        public static void AddTableByBookMark(string fileName, string[,] data, string bm)
        {
            using (var document = WordprocessingDocument.Open(fileName, true))
            {

                var doc = document.MainDocumentPart.Document;

                Table table = new Table();

                TableProperties props = new TableProperties(

                    new TableBorders(
                    new TopBorder
                    {
                        Val = new EnumValue<BorderValues>(BorderValues.Single),
                        Size = 12,
                        Color = "red"
                    },
                    new BottomBorder
                    {
                        Val = new EnumValue<BorderValues>(BorderValues.Single),
                        Size = 12
                    },
                    new LeftBorder
                    {
                        Val = new EnumValue<BorderValues>(BorderValues.Single),
                        Size = 12
                    },
                    new RightBorder
                    {
                        Val = new EnumValue<BorderValues>(BorderValues.Single),
                        Size = 12
                    },
                    new InsideHorizontalBorder
                    {
                        Val = new EnumValue<BorderValues>(BorderValues.Single),
                        Size = 12
                    },
                    new InsideVerticalBorder
                    {
                        Val = new EnumValue<BorderValues>(BorderValues.Single),
                        Size = 12
                    }));

                table.AppendChild<TableProperties>(props);

                for (var i = 0; i <= data.GetUpperBound(0); i++)
                {
                    var tr = new TableRow();
                    for (var j = 0; j <= data.GetUpperBound(1); j++)
                    {
                        var tc = new TableCell();
                        tc.Append(new Paragraph(new Run(new Text(data[i, j]))));

                        tc.Append(new TableCellProperties(
                            new TableCellWidth { Type = TableWidthUnitValues.Auto }));

                        tr.Append(tc);
                    }
                    table.Append(tr);
                }

                //Run r = new Run(table);

                //Run r= doc.Body.Append(table);


                BookmarkStart bmStart = findBookMarkStart(document, bm);
                if (bmStart == null)
                {
                    return;
                }
                bmStart.Parent.InsertAfterSelf(table);

                doc.Save();
            }

        }


        /// <summary>
        /// 创建文档操作
        /// </summary>
        /// <param name="filepath"></param>
        /// <param name="msg"></param>
        public static void CreateWordDoc(string filepath, string msg)
        {
            using (WordprocessingDocument doc = WordprocessingDocument.Create(filepath, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
            {
                MainDocumentPart mainPart = doc.AddMainDocumentPart();

                mainPart.Document = new Document();
                Body body = mainPart.Document.AppendChild(new Body());
                Paragraph para = body.AppendChild(new Paragraph());
                Run run = para.AppendChild(new Run());

                run.AppendChild(new Text(msg));
            }
        }

        /// <summary>
        /// 写文档操作
        /// </summary>
        /// <param name="filepath"></param>
        /// <param name="txt"></param>
        public static void WriteToWordDoc(string filepath, string txt)
        {
            // Open a WordprocessingDocument for editing using the filepath.
            using (WordprocessingDocument wordprocessingDocument = WordprocessingDocument.Open(filepath, true))
            {
                Body body = wordprocessingDocument.MainDocumentPart.Document.Body;

                Paragraph para = body.AppendChild(new Paragraph());
                Run run = para.AppendChild(new Run());

                RunProperties runProperties = run.AppendChild(new RunProperties(new Bold()));
                run.AppendChild(new Text(txt));
            }
        }

        /// <summary>
        /// 插入表格
        /// </summary>
        /// <param name="filepath"></param>
        public static void InsertTableInDoc(string filepath)
        {
            // Open a WordprocessingDocument for editing using the filepath.
            using (WordprocessingDocument wordprocessingDocument = WordprocessingDocument.Open(filepath, true))
            {
                Body body = wordprocessingDocument.MainDocumentPart.Document.Body;

                // Create a table.
                Table tbl = new Table();

                // Set the style and width for the table.
                TableProperties tableProp = new TableProperties();
                TableStyle tableStyle = new TableStyle() { Val = "TableGrid" };

                // Make the table width 100% of the page width.
                TableWidth tableWidth = new TableWidth() { Width = "5000", Type = TableWidthUnitValues.Pct };

                // Apply
                tableProp.Append(tableStyle, tableWidth);
                tbl.AppendChild(tableProp);

                // Add 3 columns to the table.
                TableGrid tg = new TableGrid(new GridColumn(), new GridColumn(), new GridColumn());
                tbl.AppendChild(tg);

                // Create 1 row to the table.
                TableRow tr1 = new TableRow();

                // Add a cell to each column in the row.
                TableCell tc1 = new TableCell(new Paragraph(new Run(new Text("1"))));
                TableCell tc2 = new TableCell(new Paragraph(new Run(new Text("2"))));
                TableCell tc3 = new TableCell(new Paragraph(new Run(new Text("3"))));
                tr1.Append(tc1, tc2, tc3);

                // Add row to the table.
                tbl.AppendChild(tr1);

                // Add the table to the document
                body.AppendChild(tbl);
            }
        }

        /// <summary>
        /// 表格中插入数据
        /// </summary>
        /// <param name="fileName"></param>
        /// <param name="data"></param>
        public static void AddTable(string fileName, string[] data)
        {
            using (var document = WordprocessingDocument.Open(fileName, true))
            {
                var doc = document.MainDocumentPart.Document;

                Table table = new Table();


                TableProperties tableProps = GetTableProperties(); //table properties
                table.AppendChild<TableProperties>(tableProps);

                for (var i = 0; i < data.Length; i++)
                {
                    TableRow tr = new TableRow(new TableRowProperties(new TableRowHeight { Val = Convert.ToUInt32("227") }));

                    TableCell tc = new TableCell();
                    Paragraph p = new Paragraph();
                    Run run = new Run();
                    if (i == 0)
                    {
                        Drawing r = InsertAPicture(document, @"D:\mapicon.jpg");
                        run.Append(r);
                    }

                    RunProperties rPr = GetFontProperites(i + 1); //font properties
                    run.AppendChild<RunProperties>(rPr);

                    Text text = new Text(data[i]);

                    run.Append(text);
                    p.Append(run);
                    tc.Append(p);

                    tc.Append(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Auto }));
                    tr.Append(tc);

                    table.Append(tr);
                }
                doc.Body.Append(table);
                doc.Save();
            }
        }

        public static Drawing InsertAPicture(WordprocessingDocument doc, string fileName)
        {

            MainDocumentPart mainPart = doc.MainDocumentPart;

            ImagePart imagePart = mainPart.AddImagePart(ImagePartType.Jpeg);

            using (FileStream stream = new FileStream(fileName, FileMode.Open))
            {
                imagePart.FeedData(stream);
            }

            Drawing dw = AddImageToBody(mainPart.GetIdOfPart(imagePart));
            return dw;

        }

        private static Drawing AddImageToBody(string relationshipId)
        {
            // Define the reference of the image.
            var element =
                    new Drawing(
                        new DW.Inline(
                            new DW.Extent() { Cx = 86400L, Cy = 140400L },
                            new DW.EffectExtent()
                            {
                                LeftEdge = 0L,
                                TopEdge = 0L,
                                RightEdge = 0L,
                                BottomEdge = 0L
                            },
                            new DW.DocProperties()
                            {
                                Id = (UInt32Value)1U,
                                Name = "Picture 1"
                            },
                            new DW.NonVisualGraphicFrameDrawingProperties(
                                new A.GraphicFrameLocks() { NoChangeAspect = true }),
                            new A.Graphic(
                                new A.GraphicData(
                                    new PIC.Picture(
                                        new PIC.NonVisualPictureProperties(
                                            new PIC.NonVisualDrawingProperties()
                                            {
                                                Id = (UInt32Value)0U,
                                                Name = "New Bitmap Image.jpg"
                                            },
                                            new PIC.NonVisualPictureDrawingProperties()),
                                        new PIC.BlipFill(
                                            new A.Blip(
                                                new A.BlipExtensionList(
                                                    new A.BlipExtension()
                                                    {
                                                        Uri =
                                                           "{28A0092B-C50C-407E-A947-70E740481C1C}"
                                                    })
                                            )
                                            {
                                                Embed = relationshipId,
                                                CompressionState =
                                                A.BlipCompressionValues.Print
                                            },
                                            new A.Stretch(
                                                new A.FillRectangle())),
                                        new PIC.ShapeProperties(
                                            new A.Transform2D(
                                                new A.Offset() { X = 0L, Y = 0L },
                                                new A.Extents() { Cx = 86400L, Cy = 140400L }),
                                            new A.PresetGeometry(
                                                new A.AdjustValueList()
                                            )
                                            { Preset = A.ShapeTypeValues.Rectangle }))
                                )
                                { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
                        )
                        {
                            DistanceFromTop = (UInt32Value)0U,
                            DistanceFromBottom = (UInt32Value)0U,
                            DistanceFromLeft = (UInt32Value)0U,
                            DistanceFromRight = (UInt32Value)0U,
                            EditId = "50D07946"
                        });
            return element;
            // Append the reference to body, the element should be in a Run.
            //wordDoc.MainDocumentPart.Document.Body.AppendChild(new Paragraph(new Run(element)));
        }

        /// <summary>
        /// 字体样式
        /// </summary>
        /// <param name="rowNum"></param>
        /// <returns></returns>
        private static RunProperties GetFontProperites(int rowNum)
        {
            RunProperties rPr = new RunProperties();

            Color color = new Color(); //颜色
            RunFonts runFont = new RunFonts(); //字体属性
            FontSize fs = new FontSize(); //大小
            Bold b = null;  //加粗

            runFont.Ascii = "Times New Roman";
            runFont.EastAsia = "宋体";

            if (rowNum == 1)
            {
                color.Val = "FF0000";
                fs.Val = "14";
                b = new Bold();
            }
            else if (rowNum == 2)
            {
                color.Val = "404040";
                color.ThemeTint = "BF";
                runFont.Hint = FontTypeHintValues.EastAsia;
                b = new Bold();
                fs.Val = "14";
            }
            else
            {
                color.Val = "404040";
                color.ThemeTint = "BF";
                fs.Val = "12";
            }

            rPr.Append(color);
            rPr.Append(runFont);
            rPr.Append(fs);
            rPr.Append(b);

            return rPr;
        }

        /// <summary>
        /// 表格样式
        /// </summary>
        /// <returns></returns>
        private static TableProperties GetTableProperties()
        {
            TableProperties props = new TableProperties(
                new TableWidth { Width = "4500" },
                new TableBorders(
                    new TopBorder
                    {
                        Val = new EnumValue<BorderValues>(BorderValues.Single),
                        Size = 4
                    },
                    new BottomBorder
                    {
                        Val = new EnumValue<BorderValues>(BorderValues.Single),
                        Size = 4
                    },
                    new LeftBorder
                    {
                        Val = new EnumValue<BorderValues>(BorderValues.Single),
                        Size = 4
                    },
                    new RightBorder
                    {
                        Val = new EnumValue<BorderValues>(BorderValues.Single),
                        Size = 4
                    },
                    new InsideHorizontalBorder
                    {
                        Val = new EnumValue<BorderValues>(BorderValues.Single),
                        Size = 4
                    },
                    new InsideVerticalBorder
                    {
                        Val = new EnumValue<BorderValues>(BorderValues.Single),
                        Size = 4
                    })
                );
            return props;
        }

        /// <summary>
        /// 获取文档内容
        /// </summary>
        /// <param name="document"></param>
        /// <returns></returns>
        public static string GetCommentsFromDocument(string document)
        {
            string comments = null;

            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(document, false))
            {

                MainDocumentPart mainPart = wordDoc.MainDocumentPart;
                WordprocessingCommentsPart WordprocessingCommentsPart = mainPart.WordprocessingCommentsPart;
                Stream stream = WordprocessingCommentsPart.GetStream();

                using (StreamReader streamReader = new StreamReader(stream))
                {
                    comments = streamReader.ReadToEnd();
                }
            }
            return comments;
        }

        /// <summary>
        /// 查找书签
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="bmName"></param>
        /// <returns></returns>
        private static BookmarkStart findBookMarkStart(WordprocessingDocument doc, string bmName)
        {
            foreach (var footer in doc.MainDocumentPart.FooterParts)
            {
                foreach (var inst in footer.Footer.Descendants<BookmarkStart>())
                {
                    if (inst.Name == bmName)
                    {
                        return inst;
                    }
                }
            }

            foreach (var header in doc.MainDocumentPart.HeaderParts)
            {
                foreach (var inst in header.Header.Descendants<BookmarkStart>())
                {
                    if (inst.Name == bmName)
                    {
                        return inst;
                    }
                }
            }

            foreach (var inst in doc.MainDocumentPart.RootElement.Descendants<BookmarkStart>())
            {
                if (inst is BookmarkStart)
                {
                    if (inst.Name == bmName)
                    {
                        return inst;
                    }
                }
            }

            return null;
        }

        /// <summary>
        /// 修改书签
        /// </summary>
        /// <param name="filePath">word文档</param>
        /// <param name="bmName">书签名字</param>
        /// <param name="text">替换的文本</param>
        public static void ModifyBM(string filePath, string bmName, string text)
        {
            using (WordprocessingDocument doc = WordprocessingDocument.Open(filePath, true))
            {
                BookmarkStart bmStart = findBookMarkStart(doc, bmName);

                Run bookmarkText = bmStart.NextSibling<Run>();
                if (bookmarkText != null)
                {
                    Text t = bookmarkText.GetFirstChild<Text>();
                    if (t != null)
                    {
                        t.Text = text;
                    }
                }
            }
        }
    }
}
