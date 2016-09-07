using System.Collections.Generic;

namespace OtoDom
{
    public class Ogloszenie
    {
        public string Url { get; set; }
        public Dictionary<string, string> Properties { get; set; }
        public Ogloszenie()
        {
            Properties = new Dictionary<string, string>();
        }
    }
}