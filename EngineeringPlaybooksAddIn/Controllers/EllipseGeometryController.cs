using System;
using System.Collections.Generic;
using EngineeringPlaybooksAddIn.Models;

namespace EngineeringPlaybooksAddIn.Controllers
{
    /// <summary>
    /// Recursive Implementation based on https://stackoverflow.com/a/30853013/902833
    /// </summary>
    public class EllipseGeometryController
    {
        private const double DeltaAngle = 0.0005;
        private const double ArcAccuracy = 0.01;//10%

        public enum RotationStarts
        {
            RotationStartsAtAxisX,
            RotationStartsAtAxisY
        };

        private static readonly Point EllipseCenter = new Point(0, 0);
        private static double _rectangleLeft;
        private static double _rectangleRight;
        private static double _rectangleTop;
        private static double _rectangleBottom;
        private static double _majorRadius;
        private static double _minorRadius;

        /// <summary>
        /// Based on the number of equilateralSides passed in, each point will be (recursively) targeted on
        /// the ellipse within the tolerance for arc length.
        /// </summary>
        /// <param name="ellipseMajorRadius"></param>
        /// <param name="ellipseMinorRadius"></param>
        /// <param name="equilateralSides"></param>
        /// <param name="atWhichAxis"></param>
        /// <param name="xOffsetAngleRadians"></param>
        /// <returns></returns>
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
            var points = new List<Point>();

            // Distance in radians between angles measured on the ellipse
            double circumference = GetLengthOfEllipse(DeltaAngle);

            double arcLength = circumference / equilateralSides;

            double angle = xOffsetAngleRadians - DeltaAngle;

            // Loop until we get all the points out of the ellipse
            for (var numPoints = 0; numPoints < equilateralSides; numPoints++)
            {
                angle = equilateralSides == 1 ? xOffsetAngleRadians : GetAngleForArcLengthRecursively(0, arcLength, angle, DeltaAngle);

                var xCandidate = _majorRadius * Math.Cos(angle);
                var yCandidate = _minorRadius * Math.Sin(angle);
                
                var point = atWhichAxis == RotationStarts.RotationStartsAtAxisY ? new Point(xCandidate, yCandidate) : new Point(yCandidate, xCandidate);

                points.Add(point);
            }

            return points;
        }

        /// <summary>
        /// Helper used to create a bounding box the size of the ellipse
        /// </summary>
        private static void InitializePointBackToCenter()
        {
            _rectangleBottom = EllipseCenter.Y - _minorRadius;
            _rectangleTop = EllipseCenter.Y + _minorRadius;
            _rectangleLeft = EllipseCenter.X - _majorRadius;
            _rectangleRight = EllipseCenter.X + _majorRadius;
        }

        private static double GetLengthOfEllipse(double deltaAngle)
        {
            double length = 0.0;

            // Distance in radians between angles
            double numIntegrals = Math.Round(Math.PI * 2.0 / deltaAngle);

            double radiusX = (_rectangleRight - _rectangleLeft) / 2.0;
            double radiusY = (_rectangleBottom - _rectangleTop) / 2.0;

            // integrate over the ellipse to get the circumference
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

                // We're below the our goal value, but not in range
            }

            if (currentArcPos + nextSegLength < goalArcPos - ((goalArcPos - currentArcPos) * ArcAccuracy))
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

        /// <summary>
        /// AngleBetweenInDegrees - the angle between 2 vectors
        /// </summary>
        /// <returns>
        /// Returns the the angle in degrees between vector1 and vector2
        /// </returns>
        /// <param name="vector1"> The first Vector </param>
        /// <param name="vector2"> The second Vector </param>
        public static double AngleBetweenInDegrees(Point vector1, Point vector2)
        {
            double sin = vector1.X * vector2.Y - vector2.X * vector1.Y;
            double cos = vector1.X * vector2.X + vector1.Y * vector2.Y;

            return Math.Atan2(sin, cos);
        }
    }
}
