using ConsoleTestResultsParser.Models;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;

namespace ConsoleTestResultsParser
{
    class Progra
    {
        static void Main(string[] args)
        {
            string xmlPath = "";
            if (args.Length == 0) xmlPath = ConfigurationManager.AppSettings["xlsResultFileFolder"];
            else xmlPath = args[0];

            string[] xmls = new DirectoryInfo(xmlPath).
            GetFiles("*.xml").Select(x => Path.Combine(xmlPath, x.Name)).Where(x => !x.Contains("TestResult")).ToArray();

            List<TestRun> runs = new List<TestRun>();
            foreach (var file in xmls)
            {
                try
                {
                    TestRun run = XmlParser.ReadXml(file);
                    runs.Add(run);
                }
                catch (Exception e)
                {
                    Console.Error.WriteLine(file + " file wasn't read. Exc: " + e.Message);
                }
            }
            FileWriter.WriteXls(runs, args);
        }
    }
}
