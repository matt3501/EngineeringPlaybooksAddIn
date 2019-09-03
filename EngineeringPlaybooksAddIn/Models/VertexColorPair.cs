using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EngineeringPlaybooksAddIn.Models
{
    public class VertexColorPair
    {
        public double X { get; }
        public double Y { get; }
        public Color Color { get; }

        public VertexColorPair(double x, double y, Color color)
        {
            X = x;
            Y = y;
            Color = color;
        }
    }
}
