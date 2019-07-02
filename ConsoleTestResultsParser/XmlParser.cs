using System;
using System.Collections.Generic;
using System.Configuration;
using System.Globalization;
using System.IO;
using System.Text.RegularExpressions;
using System.Xml;
using ConsoleTestResultsParser.Models;

namespace ConsoleTestResultsParser
{
    public static class XmlParser
    {
        public static TestRun ReadXml(string filename)
        {
            TestRun run = new TestRun();
            var testResultsXmlPath = Path.Combine(ConfigurationManager.AppSettings["RunTestsBatFolderAndXmlFolder"], filename);
            var testResultsXml = new XmlDocument();
            testResultsXml.Load(testResultsXmlPath);
            GetRunDataFromXml(testResultsXml, filename, ref run);
            return run;
        }

        private static void GetRunDataFromXml(XmlDocument testResultsXml, string filename, ref TestRun run)
        {
            const char splitter1 = '\\';

            run.Environment = testResultsXml.DocumentElement.FirstChild.InnerText.Split(new string[] { "bin\\" }, StringSplitOptions.None)[1].Split(splitter1)[0];

            var testSuiteNode = testResultsXml.DocumentElement.SelectSingleNode("/test-run/test-suite");
            if (!filename.EndsWith("_API.xml"))
            {
                string prefix = "Venue ver. ";
                run.VenueVersion = prefix + Regex.Match(testSuiteNode.InnerText, "Current Venue version is equal to '(.*?)'.\\r").Groups[1].Value;
                if (run.VenueVersion == prefix) run.VenueVersion = prefix + "not found";
            }
            else
            {
                string prefix = "Api ver. ";
                run.VenueVersion = prefix + Regex.Match(testSuiteNode.InnerText, " Api version from response: (.*?)\\r").Groups[1].Value;
                if (run.VenueVersion == prefix) run.VenueVersion = prefix + "not found";
            }

            run.NumberOfTests = int.Parse(testSuiteNode.Attributes["total"]?.InnerText);
            run.NumberOfPassedTests = int.Parse(testSuiteNode.Attributes["passed"]?.InnerText);
            run.NumberOfSkippedTests = int.Parse(testSuiteNode.Attributes["skipped"]?.InnerText);
            run.NumberOfFailedTests = int.Parse(testSuiteNode.Attributes["failed"]?.InnerText);
            run.RunResult = run.NumberOfFailedTests == 0 ? "Passed" : "Failed";
            run.StartTime = DateTime.Parse(testSuiteNode.Attributes["start-time"]?.InnerText);
            run.EndTime = DateTime.Parse(testSuiteNode.Attributes["end-time"]?.InnerText);
            run.Duration = TimeSpan.FromSeconds(double.Parse(testSuiteNode.Attributes["duration"]?.InnerText, CultureInfo.InvariantCulture));
            run.TestsInRun = new List<TestInRun>();
            run.FileName = filename;

            GetTestsData(testResultsXml, ref run);
        }

        private static void GetTestsData(XmlDocument testResultsXml, ref TestRun run)
        {
            var testNodes = testResultsXml.DocumentElement.SelectNodes("//test-case");
            for (int i = 0; i < run.NumberOfTests; i++)
            {
                var test = new TestInRun
                {
                    Number = i + 1,
                    Name = testNodes[i].Attributes["name"]?.InnerText,
                    Result = testNodes[i].Attributes["result"]?.InnerText,
                    Duration = TimeSpan.FromSeconds(double.Parse(testNodes[i].Attributes["duration"]?.InnerText, CultureInfo.InvariantCulture)),
                    UAE = "Not UAE fail.",
                    ServerId = "N/A"
                };

                if (!run.FileName.EndsWith("_API.xml"))
                {
                    test.ServerId = Regex.Match(testNodes[i].InnerText, "INFO Current Server Id is equal to '(.*?)'.").Groups[1].Value;
                    if (string.IsNullOrEmpty(test.ServerId)) test.ServerId = "N/A";
                    test.VenueVersion = "Venue ver. " + Regex.Match(testNodes[i].InnerText, "INFO Current Venue version is equal to '(.*?)'.").Groups[1].Value;
                    if (test.VenueVersion == "Venue ver. ") test.VenueVersion += "not found.";
                }
                else
                {
                    test.ServerId = Regex.Match(testNodes[i].InnerText, "INFO Machine number from response: (.*?)\\r").Groups[1].Value;
                    if (string.IsNullOrEmpty(test.ServerId)) test.ServerId = "N/A";
                    test.VenueVersion = "Api ver. " + Regex.Match(testNodes[i].InnerText, "INFO Api version from response: (.*?)\\r").Groups[1].Value;
                    if (test.VenueVersion == "Api ver. ") test.VenueVersion += "not found.";
                }

                if (test.Result == "Failed")
                {
                    var msg = testNodes[i].ChildNodes[1].ChildNodes[0].InnerText.Replace('\n', ' ').Replace('\r', ' ');
                    test.ErrorMessage = testNodes[i].ChildNodes[1].ChildNodes[0].InnerText.Replace('\n', ' ').Replace('\r', ' ');
                    test.StackTrace = testNodes[i]?.ChildNodes[1]?.ChildNodes[1]?.InnerText?.Replace('\n', ' ').Replace('\r', ' '); ;
                    var id = Regex.Match(testNodes[i].InnerText, "ID occurred on '(.*?)' server.").Groups[1].Value;
                    var uae = Regex.Match(testNodes[i].InnerText, "Unexpected application error with '(.*?)'").Groups[1].Value;
                    if (id.Length > 0) test.ServerId = id;
                    if (uae.Length > 0) test.UAE = uae;
                }
                run.TestsInRun.Add(test);
            }
        }
    }
}
