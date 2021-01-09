using System;

namespace RusProfileApplication.Models
{
    public class ParseError
    {
        public Guid ID { get; set; }
        public string Place { get; set; }
        public string StackTrace { get; set; }
        public string Message { get; set; }
    }
}
