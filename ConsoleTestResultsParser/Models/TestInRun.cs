using System;

namespace ConsoleTestResultsParser.Models
{
    public class TestInRun
    {
        public int Number { get; set; }
        public string Name { get; set; }
        public string Result { get; set; }
        public TimeSpan Duration { get; set; }
        public string ErrorMessage { get; set; }
        public string StackTrace { get; set; }
        public string VenueVersion { get; set; }
        public string ServerId { get; set; }
        public string UAE { get; set; }
        public TestInRun()
        {
            ErrorMessage = "N/A";
            StackTrace = "N/A";
        }
    }
}
