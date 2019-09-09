using System;
using System.Collections.Generic;
using System.Linq;
using System.Numerics;
using System.Text;
using System.Threading.Tasks;
using EngineeringPlaybooksAddIn.Models;

namespace EngineeringPlaybooksAddIn.Controllers
{
    public class GeometryController
    {
        public enum RotationStarts
        {
            RotationStartsAtAxisX,
            RotationStartsAtAxisY
        };

        private static double _rectangleLeft;
        private static double _rectangleRight;
        private static double _rectangleTop;
        private static double _rectangleBottom;
        private static double _majorRadius;
        private static double _minorRadius;
        private static double ARC_ACCURACY = 0.01;//10%

        private static Point ellipseCenter = new Point(0, 0);
        public static List<Point> GetPointsForEllipse(double ellipseMajorRadius, double ellipseMinorRadius, int equilateralSides, RotationStarts atWhichAxis, double xOffsetAngleRadians)
        {
            if (atWhichAxis == RotationStarts.RotationStartsAtAxisY)
            {
                _majorRadius = ellipseMajorRadius;
                _minorRadius = ellipseMinorRadius;
            }
            else
            {
                _majorRadius = ellipseMinorRadius;
                _minorRadius = ellipseMajorRadius;
            }

            InitializePointBackToCenter();
            List<Point> points = new List<Point>();

            // Distance in radians between angles measured on the ellipse
            double deltaAngle = 0.001;
            double circumference = GetLengthOfEllipse(deltaAngle);

            double arcLength = circumference / equilateralSides;
            //double arcLength = 0.1;

            double angle = xOffsetAngleRadians;

            // Loop until we get all the points out of the ellipse
            for (int numPoints = 0; numPoints < circumference / arcLength; numPoints++)
            {
                angle = GetAngleForArcLengthRecursively(0, arcLength, angle, deltaAngle);

                double xCandidate = _majorRadius * Math.Cos(angle);
                double yCandidate = _minorRadius * Math.Sin(angle);

                if (atWhichAxis == RotationStarts.RotationStartsAtAxisY)
                {
                    points.Add(new Point(xCandidate, yCandidate));
                } else {
                    points.Add(new Point(yCandidate, xCandidate));
                }
            }

            return points;
        }

        private static void InitializePointBackToCenter()
        {
            _rectangleBottom = ellipseCenter.Y - _minorRadius;
            _rectangleTop = ellipseCenter.Y + _minorRadius;
            _rectangleLeft = ellipseCenter.X - _majorRadius;
            _rectangleRight = ellipseCenter.X + _majorRadius;
        }

        private static double GetLengthOfEllipse(double deltaAngle)
        {
            double length = 0;

            // Distance in radians between angles
            double numIntegrals = Math.Round(Math.PI * 2.0 / deltaAngle);

            double radiusX = (_rectangleRight - _rectangleLeft) / 2;
            double radiusY = (_rectangleBottom - _rectangleTop) / 2;

            // integrate over the elipse to get the circumference
            for (int i = 0; i < numIntegrals; i++)
            {
                length += ComputeArcOverAngle(radiusX, radiusY, i * deltaAngle, deltaAngle);
            }

            return length;
        }

        private static double GetAngleForArcLengthRecursively(double currentArcPos, double goalArcPos, double angle, double angleSeg)
        {

            // Calculate arc length at new angle
            double nextSegLength = ComputeArcOverAngle(_majorRadius, _minorRadius, angle + angleSeg, angleSeg);

            // If we've overshot, reduce the delta angle and try again
            if (currentArcPos + nextSegLength > goalArcPos)
            {
                return GetAngleForArcLengthRecursively(currentArcPos, goalArcPos, angle, angleSeg / 2);

                // We're below the our goal value but not in range (
            }

            if (currentArcPos + nextSegLength < goalArcPos - ((goalArcPos - currentArcPos) * ARC_ACCURACY))
            {
                return GetAngleForArcLengthRecursively(currentArcPos + nextSegLength, goalArcPos, angle + angleSeg, angleSeg);

                // current arc length is in range (within error), so return the angle
            }

            return angle;
        }

        private static double ComputeArcOverAngle(double r1, double r2, double angle, double angleSeg)
        {
            double distance = 0.0;

            double dpt_sin = Math.Pow(r1 * Math.Sin(angle), 2.0);
            double dpt_cos = Math.Pow(r2 * Math.Cos(angle), 2.0);
            distance = Math.Sqrt(dpt_sin + dpt_cos);

            // Scale the value of distance
            return distance * angleSeg;
        }

        public static double GetDegreesBetweenVector(Point vector1, Point vector2)
        {
            var X1 = vector1.X;
            var Y1 = vector1.Y;
            var X2 = vector2.X;
            var Y2 = vector2.Y;

            var t1 = (X1 * X2) + (Y1 * Y2);
            var t2 = (X1 * X1) + (Y1 * Y1);
            var s1 = Math.Sqrt(t2);
            var t3 = (X2 * X2) + (Y2 * Y2);
            var s2 = Math.Sqrt(t3);
            var af = t1 / (s1 * s2);
            var inv = Math.Acos(af) * 360 / (Math.PI * 2);

            var angleBetween = (Math.Round(100 * inv) / 100);

            return angleBetween;
        }

        //public Vector2 GetNormalAt(Point intersection)
        //{
        //    var vectorX = (float)(intersection.X - ellipseCenter.X);
        //    var vectorY = (float)((intersection.Y - ellipseCenter.Y) * (_majorRadius / _minorRadius));
        //    var perpendicularVector = new Vector2(vectorX, vectorY);

        //    var normalVector = Vector2.Normalize(perpendicularVector);
        //    return normalVector;
        //}
    }
}
