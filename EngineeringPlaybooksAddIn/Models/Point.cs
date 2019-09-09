namespace EngineeringPlaybooksAddIn.Models
{
    public class Point
    {
        public double X { get; set; }
        public double Y { get; set; }
        public Point(double x, double y)
        {
            X = x;
            Y = y;
        }

        public override bool Equals(object obj)
        {
            if (obj == null || obj.GetType() != typeof(Point)) return false;

            Point tempPoint = (Point)obj;

            return tempPoint.X.Equals(X) && tempPoint.Y.Equals(Y);
        }

        public override int GetHashCode()
        {
            return (int)(55667 * X + 129 * Y);
        }
    }
}