using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using Aspose.Words;
using Aspose.Words.Tables;
using TemplateSpeAsposeCLI.Utils;

namespace AnalysisReportCLI
{
    public static class AnalysisReport
    {
        public static void ProcessFile(string pathSourceFile, string output = "")
        {
            Document doc = new Document(pathSourceFile + @"\" + "exemple RA-25-034 (avant post traitement).docm");

            Bookmark bookmark = doc.Range.Bookmarks["tableau_analyse"];
            if (bookmark != null)
            {
                Table bookmarkTable = (Table)bookmark.BookmarkStart.GetAncestor(NodeType.Table);
                if (bookmarkTable != null)
                {
                    Console.WriteLine("Table Found at the Bookmark Function");
                    List<Row> rows = bookmarkTable.Rows.Skip(1).Cast<Row>().ToList();
                    Dictionary<string, List<Row>> groupedRows = new Dictionary<string, List<Row>>();
                    foreach (Row row in rows.Cast<Row>())

                    {
                        string family = row.Cells[0].GetText().Trim(); // Extract Family column
                        if (!groupedRows.ContainsKey(family))
                        {
                            groupedRows[family] = new List<Row>();
                        }
                        groupedRows[family].Add(row);
                    }

                    //Clear existing rows except header
                    for (int i = bookmarkTable.Rows.Count - 1; i > 0; i--)
                    {
                        bookmarkTable.Rows.RemoveAt(i);
                    }
                    bookmarkTable.FirstRow.FirstCell.Remove();
                    foreach (var group in groupedRows.OrderBy(g => g.Key))
                    {
                        Row mergedRow = new Row(doc);
                        Cell mergedCell = new Cell(doc);
                        mergedCell.CellFormat.HorizontalMerge = CellMerge.First;
                        mergedCell.CellFormat.Shading.BackgroundPatternColor = System.Drawing.ColorTranslator.FromHtml("#D9EDF2");
                        mergedCell.CellFormat.Borders.Color = System.Drawing.Color.White;
                        mergedCell.Paragraphs.Add(new Paragraph(doc));
                        mergedCell.FirstParagraph.AppendChild(new Run(doc, group.Key)
                        {
                            Font = { Color = System.Drawing.ColorTranslator.FromHtml("#17365C"), Bold = true, Size = 9 }
                        });

                        mergedRow.Cells.Add(mergedCell);
                        for (int i = 1; i < bookmarkTable.FirstRow.Cells.Count; i++)
                        {
                            Cell emptyCell = new Cell(doc);
                            emptyCell.CellFormat.HorizontalMerge = CellMerge.Previous;
                            mergedRow.Cells.Add(emptyCell);
                        }

                        bookmarkTable.Rows.Add(mergedRow);
                        foreach (var row in group.Value.OrderBy(r => r.Cells[1].GetText().Trim())) // Sort by Parameter
                        {
                            Row newRow = new Row(doc);
                            for (int i = 1; i < row.Cells.Count; i++) // Exclude Family column
                            {
                                Cell newCell = (Cell)row.Cells[i].Clone(true);
                                newRow.Cells.Add(newCell);
                            }
                            bookmarkTable.Rows.Add(newRow);
                        }

                    }
                    doc.Save(pathSourceFile + @"\Processed_Document.docm");
                    Console.WriteLine("Processing completed. Document saved.");
                }
                else
                {
                    Console.WriteLine("No table found at the bookmark.");
                }
            }



        }

        public static void MySecondFunction(List<string> sourceFiles, string output = "")
        {
            for (int i = 0; i < sourceFiles.Count; i++)
            {
                // Load word file
                Document doc = new Document(sourceFiles[i]);

                // Update title labels (4., 4.1., 4.1.1., ...)
                doc.UpdateListLabels();
                // Init a builder that'll be used to write information in the document
                DocumentBuilder builder = new DocumentBuilder(doc);

                // Write after a named paragraph
                Paragraph par = UtilsWord.GetParagraph(doc, "4.1.4.");
                if (par != null)
                {
                    // Move to the paragraph
                    builder.MoveTo(par);
                    // Go to a new line
                    builder.InsertBreak(BreakType.ParagraphBreak);
                    // Set the style back to normal
                    builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
                    builder.Writeln("ICI");
                }

                // Manage text insertion after/before/replace bookmark
                UtilsWord.WriteBeforeBookmark(builder, "branche", "BEFORE BOOKMARK ");
                UtilsWord.WriteAfterBookmark(builder, "branche", " AFTER BOOKMARK");
                UtilsWord.ReplaceBookmarkWithText(doc, builder, "branche", "REPLACE BOOKMARK");

                // Concat documents
                Document doc2 = new Document(@"D:\Exemple - Copie.docx");
                UtilsWord.ConcatDoc(doc, doc2, false, true);

                // TOC Management
                builder.MoveToDocumentStart();
                builder.InsertTableOfContents("\\o \"1-5\" \\h \\z \\u");
                UtilsWord.UpdateTableOfContents(doc);
                //UtilsWord.RemoveTableOfContents(doc);

                // Save doc
                doc.Save(output);
            }
        }

        /// <summary>
        /// Function necessary to work with TxExtraction
        /// </summary>
        /// <param name="outputPath">The path of the output file</param>
        /// <param name="pathSourceFile">Use to know where to save the xml output</param>
        private static void CreateOutputXml(string outputPath, string pathSourceFile)
        {
            // Creating output xml file for TxExtraction
            var xOutput = new XDocument();
            xOutput.Add(new XElement("Output"));
            xOutput.Root.Add(new XElement("Status", "OK"));
            xOutput.Root.Add(new XElement("OutputFilePath", outputPath));
            xOutput.Save(string.Concat(Path.GetDirectoryName(pathSourceFile) + @"\", System.Diagnostics.Process.GetCurrentProcess().ProcessName, ".xml"));
        }
    }
}
