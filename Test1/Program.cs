using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

namespace Test1
{
    class Program
    {
        static void Main(string[] args)
        {
            //创建文档
            //CreateWordDoc(@"D:\Test.docx", "Hello World");

            //写文档
            //WriteToWordDoc(@"D:\Test.docx", "测试1");

            //插入Table
            //InsertTableInDoc(@"D:\Test.docx");

            //读取文档
            //string contentDoc = GetCommentsFromDocument(@"D:\Test.docx");

            //替换书签
            //ModifyBM(@"D:\Test.docx", "test","这是测试");


            //查找书签
            //WordprocessingDocument doc = WordprocessingDocument.Open(@"D:\LegalServiceTemplate.docx", true);
            // BookmarkStart bs = findBookMarkStart(doc, "table_test");

            //在表格中插入数据
            //const string fileName = @"D:\Test.docx";
            //AddTable(fileName, new string[,]
            //    { { "Texas", "TX","dc" },
            //    { "California", "CA","dd" },
            //    { "New York", "NY","" },
            //    { "Massachusetts", "MA","cy" } }
            //);

            //查找书签
            //AddTableByBookMark(@"D:\Test.docx", data, "table_Address");
            //AddTable(@"D:\test.docx", data);

            //string document = @"D:\Test.docx";
            //string fileName = @"D:\mapicon.jpg";
            //InsertAPicture(document, fileName);

            //测试模板页面
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(@"D:\test1.docx", true))
            {
                TOffice office = new TOffice();
                var lst = office.GetListTOffice();

                //2.根据书签获取word位置信息
                BookmarkStart bmAddress1 = findBookMarkStart(wordDoc, "table_Address1");
                BookmarkStart bmAddress2 = findBookMarkStart(wordDoc, "table_Address2");

                //3.创建表格填充数据
                int i = 0;
                foreach (var item in lst)
                {
                    i++;
                    if (i % 2 != 0)
                    {
                        Table tb = AddTable(wordDoc, item);
                        bmAddress1.Parent.InsertAfterSelf(tb); //add table at bookmark
                    }
                    else
                    {
                        Table tb = AddTable(wordDoc, item);
                        bmAddress2.Parent.InsertAfterSelf(tb); //add table at bookmark
                    }
                }

            }

        }


        /// <summary>
        /// 表格中插入数据
        /// </summary>
        /// <param name="fileName"></param>
        /// <param name="data"></param>
        public static Table AddTable(WordprocessingDocument wordDoc, TOffice office)
        {
            string[] data = new string[] { office.Name + "｜" + office.NameEn, office.Address, office.AddressEn, "Tel：" + office.Telephone, "Fax：" + office.Fax };

            var doc = wordDoc.MainDocumentPart.Document;

            Table table = new Table();
            TableProperties tableProps = GetTableProperties(); //table properties
            table.AppendChild<TableProperties>(tableProps);

            for (var i = 0; i < data.Length; i++)
            {
                TableRow tr = new TableRow() { RsidTableRowAddition = "00BF0C52", RsidTableRowProperties = "001640FD" };
                TableCell tc = new TableCell();

                Paragraph p = new Paragraph(new ParagraphProperties(new SpacingBetweenLines() { Line = "140", LineRule = LineSpacingRuleValues.Exact }));
                Run run = new Run();
                if (i == 0)
                {
                    Drawing r = InsertAPicture(wordDoc, @"D:\mapicon.jpg");
                    run.Append(r);
                }

                RunProperties rPr = GetFontProperites(i + 1); //font properties
                run.AppendChild<RunProperties>(rPr);

                Text text = new Text(data[i]);

                run.Append(text);
                p.Append(run);
                tc.Append(p);

                //tc.Append(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Auto }));
                tr.Append(tc);

                table.Append(tr);
            }
            return table;
            // doc.Body.Append(table);
            // doc.Save();
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
        /// 设置表格字体
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
            SpacingBetweenLines spacingBetweenLines1 = new SpacingBetweenLines() { Line = "140", LineRule = LineSpacingRuleValues.Exact };
            Indentation indentation1 = new Indentation() { FirstLine = "281", FirstLineChars = 200 };

            runFont.Ascii = "Times New Roman";
            runFont.EastAsia = "宋体";

            if (rowNum == 1)
            {
                color.Val = "FF6600";
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
            rPr.Append(spacingBetweenLines1);
            rPr.Append(indentation1);

            return rPr;
        }

        /// <summary>
        /// 设置表格样式
        /// </summary>
        /// <returns></returns>
        private static TableProperties GetTableProperties()
        {
            TableProperties props = new TableProperties(
                new TableStyle { Val = "a7" },
                new TableWidth { Width = "4000" },
                new TableRowProperties(new TableRowHeight { Val = (UInt32Value)227U }),
                new TableCellWidth() { Type = TableWidthUnitValues.Dxa },
                new TableLook() { Val = "04A0", FirstRow = true, LastRow = false, FirstColumn = true, LastColumn = false, NoHorizontalBand = false, NoVerticalBand = true },
                new TableCellMarginDefault(new TableCellLeftMargin { Width = 0, Type = TableWidthValues.Dxa }),

                new TableBorders(
                    new TopBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U },
                    new BottomBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U },
                    new LeftBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U },
                    new RightBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U },
                    new InsideHorizontalBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U },
                    new InsideVerticalBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U })
                );

            return props;
        }




        #region 图片设置

        /// <summary>
        /// 填充图片
        /// </summary>
        /// <param name="relationshipId"></param>
        /// <returns></returns>
        private static Drawing AddImageToBody(string relationshipId)
        {
            // Define the reference of the image.
            var element =
                    new Drawing(
                        new DW.Anchor() { SimplePos = false, BehindDoc = false, Locked = false, LayoutInCell = true, AllowOverlap = true},
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
        /// 画图片
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="fileName"></param>
        /// <returns></returns>
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

        #endregion



    }
}
