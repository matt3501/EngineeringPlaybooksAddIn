using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EngineeringPlaybooksAddIn.Models
{
    public class Color
    {
        public int R { get; set; }
        public int G { get; set; }
        public int B { get; set; }

        public Color(int R, int G, int B)
        {
            this.R = R;
            this.G = G;
            this.B = B;
        }
    }
}
