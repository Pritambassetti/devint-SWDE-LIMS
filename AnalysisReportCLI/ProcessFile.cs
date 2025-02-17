using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using AnalysisReportCLI.Models;
using Aspose.Cells;
using TemplateSpeAsposeCLI.Utils;

namespace AnalysisReportCLI
{
    public static class AnalysisReport
    {
        public static void ProcessFile(string pathSourceFile, string output = "")
        {
            if (!File.Exists(pathSourceFile))
            {
                throw new FileNotFoundException("File Not Found!");
            }

            Workbook sourceWorkbook = new Workbook(pathSourceFile);
            Worksheet RésultatsSheet = sourceWorkbook.Worksheets["Résultats"];
            Worksheet rapportSheet = sourceWorkbook.Worksheets["Rapport"];

            if (RésultatsSheet == null)
            {
                throw new Exception("Sheet 'Résultats' not found!");
            }

            if (rapportSheet == null)
            {
                throw new Exception("Sheet 'Rapport' not found!");
            }

            ProcessData(RésultatsSheet, rapportSheet);
            string outputFile = string.IsNullOrEmpty(output) ?
                             Path.GetDirectoryName(pathSourceFile) + @"\output_" + Path.GetFileName(pathSourceFile)
                             : output;

            sourceWorkbook.Save(outputFile);

        }


        public static void ProcessData(Worksheet resultantsSheet, Worksheet rapportSheet)
        {
            int resultantsRowCount = resultantsSheet.Cells.MaxDataRow;
            int resultantsColCount = resultantsSheet.Cells.MaxDataColumn;
            Dictionary<string, List<ResultData>> resultList = new Dictionary<string, List<ResultData>>();
            for (int row = 1; row <= resultantsRowCount; row++)
            {
                string description = resultantsSheet.Cells[row, 0].StringValue;
                string resultatConsolide = resultantsSheet.Cells[row, 1].StringValue;
                string unite = resultantsSheet.Cells[row, 2].StringValue;
                string norme = resultantsSheet.Cells[row, 5].StringValue;
                string dateDebutAnalyse = resultantsSheet.Cells[row, 6].StringValue;
                string cellule = resultantsSheet.Cells[row, 7].StringValue;
                string accreditation = resultantsSheet.Cells[row, 8].StringValue;
                ResultData resultData = new ResultData
                {
                    Description = description,
                    ResultatConsolide = resultatConsolide,
                    Unite = unite,
                    Norme = norme,
                    DateDebutAnalyse = dateDebutAnalyse,
                    Accreditation = accreditation
                };
                if (resultList.ContainsKey(cellule)) resultList[cellule].Add(resultData);
                else
                    resultList.Add(cellule, new List<ResultData>() { resultData });
            }

            foreach (var result in resultList)
            {
                if (result.Key == "Microbiologie")
                {
                    Cell mklml = UtilsCell.FindCellByValue(rapportSheet, "Bac - Bacto classique");
                    if (mklml != null && !string.IsNullOrEmpty(UtilsCell.GetCellStringValue(rapportSheet, mklml.Row - 2, 0)) && UtilsCell.GetCellStringValue(rapportSheet, mklml.Row - 2, 0) == "bactériologie")
                    {
                        int rapportRow = mklml.Row + 1;
                        rapportSheet.Cells.InsertRows(rapportRow, result.Value.Count);
                        foreach (var data in result.Value)
                        {
                            rapportSheet.Cells[rapportRow, 1].Value = (data.Accreditation == "Oui") ? "X" : "";
                            MergeCellsAndSetValue(rapportSheet, rapportRow, 3, 1, 4, data.Description);
                            rapportSheet.Cells[rapportRow, 7].Value = data.ResultatConsolide;
                            rapportSheet.Cells[rapportRow, 9].Value = data.Unite;
                            MergeCellsAndSetValue(rapportSheet, rapportRow, 11, 1, 2, data.Norme);
                            MergeCellsAndSetValue(rapportSheet, rapportRow, 13, 1, 5, data.DateDebutAnalyse);
                            rapportRow++;
                        }
                    }
                }
            }
        }




        // Function to merge a range of cells and set the value for the first cell in the range
        public static void MergeCellsAndSetValue(Worksheet worksheet, int row, int colStart, int rowCount, int colCount, string value)
        {
            var range = worksheet.Cells.CreateRange(row, colStart, rowCount, colCount);
            range.Merge();
            worksheet.Cells[row, colStart].Value = value;
        }
        // Assuming "Microbiologie" is in column H (Index 7)
        //string celluleValue = resultantsSheet.Cells[row, 7].StringValue;

        //if (celluleValue == "Microbiologie")
        //{
        // Find "Bac - Bacto classique" in the "Rapport" sheet
        //int rapportRowCount = rapportSheet.Cells.MaxDataRow;
        //for (int rapportRow = 0; rapportRow <= rapportRowCount; rapportRow++)
        //{
        //    string rapportCellValue = rapportSheet.Cells[rapportRow, 1].StringValue;

        //    if (rapportCellValue == "Bac - Bacto classique")
        //    {
        //        rapportSheet.Cells[rapportRow, 1].Value = (accreditation == "Oui") ? "X" : "";

        //        MergeCellsAndSetValue(rapportSheet, rapportRow, 3, 1, 4, description);
        //        rapportSheet.Cells[rapportRow, 7].Value = resultatConsolide;
        //        rapportSheet.Cells[rapportRow, 9].Value = unite;
        //        MergeCellsAndSetValue(rapportSheet, rapportRow, 11, 1, 2, norme);
        //        MergeCellsAndSetValue(rapportSheet, rapportRow, 13, 1, 5, dateDebutAnalyse);
        //    }
        //}
        //break;  

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
