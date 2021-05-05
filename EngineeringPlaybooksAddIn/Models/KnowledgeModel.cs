using System;
using System.Collections.Generic;

namespace EngineeringPlaybooksAddIn.Models
{
    public class KnowledgeModel
    {
        public string title { get; set; }
        public string description { get; set; }
        public List<Outcome> outcomes { get; set; }
    }
}
