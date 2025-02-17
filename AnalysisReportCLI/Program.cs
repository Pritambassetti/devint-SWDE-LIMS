using System;
using System.IO;
using System.Xml.Linq;
using Aspose.Words;
using Microsoft.Extensions.CommandLineUtils;

namespace AnalysisReportCLI
{
    /// <summary>
    /// THIS IS A PROJECT TEMPLATE. THINK ABOUT UPDATING YOUR EXE NAME.
    /// THIS PROJECT USE ASPOSE.CELLS AND ASPOSE.WORDS, REMOVE WHAT YOU DON'T USE. IF YOU'RE USING ASPOSE.CELLS, REMOVE UtilsWord.cs VICE VERSA.
    /// </summary>
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                LoadLicense();

                #region 'Case for TxExtraction
                var app = new CommandLineApplication()
                {
                    Name = System.Diagnostics.Process.GetCurrentProcess().ProcessName + ".exe",
                    FullName = System.Diagnostics.Process.GetCurrentProcess().ProcessName + ".exe"
                };
                app.HelpOption("-?|-h|--help");
                app.VersionOption("--version", typeof(Program).Assembly.GetName().Version.ToString);
                app.Description = "An executable compatible with TxExtraction";

                CommandArgument arg = app.Argument("[PATH]", "Mandatory. Full path of the input document");

                app.OnExecute(() =>
                {
                    string inPath = arg.Value;
                    if (string.IsNullOrEmpty(inPath))
                        throw new Exception("Input file path missing or doesn't exist");
                    AnalysisReport.ProcessFile(inPath);
                    return 0;
                });
                #endregion

                app.Execute(args);
            }
            catch (Exception e)
            {
                # region KEEP THIS REGION
                var xOutput = new XDocument();
                xOutput.Add(new XElement("Output"));
                xOutput.Root.Add(new XElement("Status", "KO"));
                xOutput.Root.Add(new XElement("Error", e.Message));

                if (args.Length == 1 && File.Exists(args[0]))
                {
                    xOutput.Save(string.Concat(Path.GetDirectoryName(args[0]), "\\", System.Diagnostics.Process.GetCurrentProcess().ProcessName, ".xml"));
                }
                else
                {
                    string executableDirectory = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + @"\";
                    xOutput.Save(string.Concat(executableDirectory, System.Diagnostics.Process.GetCurrentProcess().ProcessName, ".xml"));
                }
                #endregion
            }
        }

        private static void LoadLicense()
        {
            // Instantiate the License object
            License lic = new License();
            // Set the License stream
            lic.SetLicense("Aspose.Total.NET.lic");
        }
    }
}
