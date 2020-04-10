namespace EngineeringPlaybooksAddIn.Models
{
    public class VertexColorPair
    {
        public double X => Point.X;
        public double Y => Point.Y;
        public Color Color { get; set; }

        public Point Point { get; }

        public VertexColorPair(Point point, Color color)
        {
            Point = point;
            Color = color;
        }

        /// <summary>
        /// Extra constructor to provide
        /// </summary>
        /// <param name="point"></param>
        /// <param name="color"></param>
        /// <param name="transformAxis"></param>
        public VertexColorPair(Point point, Color color, bool transformAxis)
        {
            var transformScalar = transformAxis ? -1 : 1;
            Point = new Point(point.X * transformScalar, point.Y * transformScalar);
            Color = color;
        }
    }
}
