using System;
using System.Collections.Generic;

namespace EngineeringPlaybooksAddIn.Models
{
    public class Outcome
    {
        public string title { get; set; }
        public string contentUrl { get; set; }
        public List<Outcome> childOutcomes { get; set; }
    }
}
