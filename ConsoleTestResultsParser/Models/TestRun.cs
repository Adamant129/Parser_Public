using System;
using System.Collections.Generic;

namespace ConsoleTestResultsParser.Models
{
    public class TestRun
    {
        private string _environment;
        private string _fileName;

        //GetFromCommandline value, 
        //Replaces on get "VM26" to devInt(Required by Maksim)
        public string Environment
        {
            get => _environment.Contains("VM26") ? _environment.Replace("VM26", "devInt") : _environment;
            set => _environment = value;
        }

        //Replaces on get "VM26" to devInt(Required by Maksim)
        public string FileName
        {
            get => _fileName.Contains("VM26") ? _fileName.Replace("VM26", "devInt") : _fileName;
            set => _fileName = value;
        }
        public string VenueVersion { get; set; }
        public string ApiVersion { get; set; }
        public string RunResult { get; set; }
        public int NumberOfTests { get; set; }
        public int NumberOfPassedTests { get; set; }
        public int NumberOfFailedTests { get; set; }
        public int NumberOfSkippedTests { get; set; }
        public DateTime StartTime { get; set; }
        public DateTime EndTime { get; set; }
        public TimeSpan Duration { get; set; }
        public List<TestInRun> TestsInRun { get; set; }

    }
}
