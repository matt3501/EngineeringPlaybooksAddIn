using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EngineeringPlaybooksAddIn.Models
{
    public class VertexColorPair
    {
        public double X { get; set; }
        public double Y { get; set; }
        public Color Color { get; set; }

        public VertexColorPair(double x, double y, Color color)
        {
            X = x;
            Y = y;
            Color = color;
        }
    }
}
