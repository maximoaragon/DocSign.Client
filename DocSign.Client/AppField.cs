using System;
using System.Collections.Generic;
using System.Text;

namespace DocSign.Client
{
    public class AppField
    {
        public string FieldID { get; set; }
        public string Value { get; set; }
        public bool IsSignature { get; set; }
        public string SignatureTag { get; set; }
    }
}
