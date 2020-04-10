using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Drawing;
using System.Linq;
using System.Text.RegularExpressions;
using EngineeringPlaybooksAddIn.Models;
using Microsoft.Office.Interop.Visio;
using Newtonsoft.Json;
using Color = EngineeringPlaybooksAddIn.Models.Color;
using Point = EngineeringPlaybooksAddIn.Models.Point;

namespace EngineeringPlaybooksAddIn.Controllers
{
    public class JciPlaybooksDrawingController
    {
        private const double XCenter = 5.2875;
        private const double YCenter = 3.72;
        private const double KeyNodeMajorRadius = 1.5375;
        private const double KeyNodeMinorRadius = 1.2775;
        private const double ChildNodeMajorRadius = 0.532;
        private const double ChildNodeMinorRadius = 0.442;
        private const string DefaultEllipseColor = "RGB(248, 248, 248)";

        /// <summary>
        /// These colors have been sourced from https://my.jci.com/brand-center/Resources/Johnson%20Controls%20brand%20guidelines_January%202018_v2.2.pdf
        /// </summary>
        private static readonly IList<Color> JciColors = new ReadOnlyCollection<Color>(new List<Color>()
        {
            new Color(0, 128, 182), //Secondary JCI Blue, HEX = 0080B6
            new Color(0, 183, 168), //Secondary JCI Aqua, HEX = 00B7A8
            new Color(106, 193, 123), //Secondary JCI Green, HEX = 6AC07A
            new Color(214, 213, 37), //Secondary JCI Yellow, HEX = D5D537
            new Color(254, 189, 56), //Secondary JCI Orange, HEX = FDBC38
            new Color(244, 119, 33), //Secondary JCI Burnt Orange, HEX = F37720
            new Color(203, 36, 57), //Secondary JCI Red, HEX = CA2339
        });

        /// <summary>
        /// Allows the calling into ActivePage similar to the syntax used by traditional VB6 macros
        /// </summary>
        // ReSharper disable once InconsistentNaming
        private Page ActivePage;

        private readonly Point _yAxisUnitVector;
        private double _doubleTolerance;

        public JciPlaybooksDrawingController()
        {
            _yAxisUnitVector = new Point(0, 1);
        }

        /// <summary>
        /// public facing method for drawing an Engineering Playbook, use samples/knowledge_WorkflowsBusinessAnalyst.json as a schema example
        /// </summary>
        /// <param name="jsonText"></param>
        public void DrawPlaybook(string jsonText)
        {
            ActivePage = Globals.ThisAddIn.Application.ActivePage;

            SetOrientationToLandscape();

            KnowledgeModel model = ValidateAndTrimModel(jsonText);

            DrawHeader(model);

            DrawMap(model);
        }

        protected static KnowledgeModel ValidateAndTrimModel(string jsonText)
        {
            var validateAndTrimModel = JsonConvert.DeserializeObject<KnowledgeModel>(jsonText);

            return validateAndTrimModel;
        }

        /// <summary>
        /// Draws the box header based on model.title and model.description which is parsed from jsonText
        /// </summary>
        /// <param name="model"></param>
        private void DrawHeader(KnowledgeModel model)
        {
            ActivePage.DrawRectangle(0.25, 8.25, 10.75, 7.133);
            var title = ActivePage.DrawRectangle(1.25, 8.25, 10.75, 7.9);
            title.LineStyle = "Guide";
            title.Text = model.title;
            title.Characters.CharProps[(short) VisCellIndices.visCharacterStyle] = (short) tagVisCellVals.visBold;

            var description = ActivePage.DrawRectangle(1.25, 7.9, 10.75, 7.133);
            description.LineStyle = "Guide";
            description.Text = model.description;

            title.CellsSRC[(short) VisSectionIndices.visSectionParagraph, 0,
                (short) VisCellIndices.visHorzAlign].FormulaU = "0";


            var headRadius = 0.171875;

            var userCenterX = .75;
            var userHead = ActivePage.DrawOval(userCenterX - headRadius, 7.75 + headRadius, userCenterX + headRadius,
                7.75 - headRadius);
            userHead.CellsU["Fillforegnd"].FormulaU = "RGB(255, 255, 255)";
            userHead.CellsSRC[(short) VisSectionIndices.visSectionObject, (short) VisRowIndices.visRowLine,
                (short) VisCellIndices.visLineWeight].FormulaForce = "3pt";

            var userBody = ActivePage.DrawCircularArc(userCenterX, 7.0, .5, 0.78541, 2.356);
            userBody.CellsSRC[(short) VisSectionIndices.visSectionObject, (short) VisRowIndices.visRowLine,
                (short) VisCellIndices.visLineWeight].FormulaForce = "3pt";

            description.CellsSRC[(short) VisSectionIndices.visSectionParagraph, 0,
                (short) VisCellIndices.visHorzAlign].FormulaU = "0";

            description.CellsSRC[(short) VisSectionIndices.visSectionObject,
                (short) VisRowIndices.visRowText,
                (short) VisCellIndices.visTxtBlkVerticalAlign].FormulaU = "0";
            description.Cells["Char.Size"].FormulaU = "10 pt";
        }

        /// <summary>
        /// Sets the currently open document to landscape (if not already set)
        /// </summary>
        private void SetOrientationToLandscape()
        {
            var orientation = ActivePage.PageSheet.CellsSRC[(short) VisSectionIndices.visSectionObject,
                (short) VisRowIndices.visRowPrintProperties,
                (short) VisCellIndices.visPrintPropertiesPageOrientation].FormulaU;

            //If orientation flag isn't set to 'landscape'
            if (orientation != ((int) tagVisCellVals.visPPOLandscape).ToString())
            {
                //Landscape
                ActivePage.PageSheet.CellsSRC[(short) VisSectionIndices.visSectionObject,
                        (short) VisRowIndices.visRowPrintProperties,
                        (short) VisCellIndices.visPrintPropertiesPageOrientation].FormulaU =
                    ((int) tagVisCellVals.visPPOLandscape).ToString();

                var currentWidth = ActivePage.PageSheet.CellsSRC[(short) VisSectionIndices.visSectionObject,
                    (short) VisRowIndices.visRowPage, (short) VisCellIndices.visPageWidth].FormulaU;
                var currentHeight = ActivePage.PageSheet.CellsSRC[(short) VisSectionIndices.visSectionObject,
                    (short) VisRowIndices.visRowPage, (short) VisCellIndices.visPageHeight].FormulaU;


                ActivePage.PageSheet.CellsSRC[(short) VisSectionIndices.visSectionObject,
                    (short) VisRowIndices.visRowPage, (short) VisCellIndices.visPageWidth].FormulaU = currentHeight;
                ActivePage.PageSheet.CellsSRC[(short) VisSectionIndices.visSectionObject,
                    (short) VisRowIndices.visRowPage, (short) VisCellIndices.visPageHeight].FormulaU = currentWidth;
            }
        }

        private void DrawMap(KnowledgeModel model)
        {
            var ellipseVectors = GetEllipseVertices(model);

            var coreOval = ActivePage.DrawOval(XCenter - KeyNodeMajorRadius, YCenter + KeyNodeMinorRadius,
                XCenter + KeyNodeMajorRadius, YCenter - KeyNodeMinorRadius);
            coreOval.CellsU["Fillforegnd"].FormulaU = DefaultEllipseColor;
            coreOval.Text = "Key Outcomes";
            coreOval.Cells["Char.Size"].FormulaU = "16 pt";
            coreOval.Characters.CharProps[(short) VisCellIndices.visCharacterStyle] = (short) tagVisCellVals.visBold;

            for (var index = 0; index < model.outcomes.Count; index++)
            {
                var keyOutcome = model.outcomes[index];
                var vertexColorPair = ellipseVectors[index];
                DrawKeyOutcomeChildNode(keyOutcome, vertexColorPair, XCenter, YCenter);
            }
        }

        /// <summary>
        /// Queries for the geometric vectors and node colors for each json 'outcome'
        /// </summary>
        /// <param name="model"></param>
        /// <returns></returns>
        protected List<VertexColorPair> GetEllipseVertices(KnowledgeModel model)
        {
            var keyOutcomesCount = model.outcomes.Count;

            double aestheticXOffsetAngleRadians;
            switch (keyOutcomesCount)
            {
                case 2:
                    aestheticXOffsetAngleRadians = Math.PI / 2;
                    break;
                case 4:
                    aestheticXOffsetAngleRadians = Math.PI / 4;
                    break;
                default:
                    aestheticXOffsetAngleRadians = 0.0;
                    break;
            }

            var geometry = EllipseGeometryController.GetPointsForEllipse(
                KeyNodeMajorRadius,
                KeyNodeMinorRadius,
                keyOutcomesCount,
                EllipseGeometryController.RotationStarts.RotationStartsAtAxisX,
                aestheticXOffsetAngleRadians);

            var aestheticCountFlipSet = new[] {5,7};

            var transformAxis = aestheticCountFlipSet.Contains(keyOutcomesCount);

            return geometry.Select((point, index) =>
                new VertexColorPair(point, JciColors[index], transformAxis)).ToList();
        }

        protected List<VertexColorPair> CalculateEllipseBranchVertices(Outcome outcome, double xOffsetAngleRadians)
        {
            var count = outcome.childOutcomes.Count;

            double radiusScalar;
            if (count < 4)
                radiusScalar = 0.05;//tighter coupling
            else if (count == 4)
                radiusScalar = 0.07;//slightly looser coupling
            else
                radiusScalar = 0.085;//farthest coupling

            var geometry = EllipseGeometryController.GetPointsForEllipse(
                ChildNodeMajorRadius + (radiusScalar * count),
                ChildNodeMinorRadius + (radiusScalar * count),
                count,
                EllipseGeometryController.RotationStarts.RotationStartsAtAxisY,
                xOffsetAngleRadians);

            return geometry.Select((point, index) =>
                new VertexColorPair(point, JciColors[index])).ToList();
        }

        /// <summary>
        /// Pulls text and URL from keyOutcome, relative vector drawn off childEllipseCenter
        ///
        /// The size of the oval drawn is set by KeyChildNodeMajor/Minor Radius
        /// </summary>
        /// <param name="keyOutcome"></param>
        /// <param name="vertexColorPair"></param>
        /// <param name="childEllipseCenterX"></param>
        /// <param name="childEllipseCenterY"></param>
        public void DrawKeyOutcomeChildNode(Outcome keyOutcome, VertexColorPair vertexColorPair,
            double childEllipseCenterX, double childEllipseCenterY)
        {
            var ellipseNodeX1 = childEllipseCenterX + vertexColorPair.X - ChildNodeMajorRadius;
            var ellipseNodeY1 = childEllipseCenterY + vertexColorPair.Y + ChildNodeMinorRadius;
            var ellipseNodeX2 = childEllipseCenterX + vertexColorPair.X + ChildNodeMajorRadius;
            var ellipseNodeY2 = childEllipseCenterY + vertexColorPair.Y - ChildNodeMinorRadius;
            var ellipseColor = "RGB(" + vertexColorPair.Color.R + ", " + vertexColorPair.Color.G + ", " +
                               vertexColorPair.Color.B + ")";

            DrawEllipse(keyOutcome, ellipseNodeX1, ellipseNodeY1, ellipseNodeX2, ellipseNodeY2, ellipseColor);

            if (!keyOutcome.childOutcomes.Any()) return;

            DrawBranch(keyOutcome, vertexColorPair);
        }

        /// <summary>
        /// Assumes there is at least one child outcome
        /// </summary>
        /// <param name="keyOutcome"></param>
        /// <param name="vertexColorPair"></param>
        private void DrawBranch(Outcome keyOutcome, VertexColorPair vertexColorPair)
        {
            if (!keyOutcome.childOutcomes.Any()) return;

            double vectorAddendX = 0.0;
            double vectorSubtrahendY = 0.0;

            var vertexX = Math.Round(vertexColorPair.X, 2);
            var vertexY = Math.Round(vertexColorPair.Y, 2);
            _doubleTolerance = 0.001;
            if (keyOutcome.childOutcomes.Count > 3)
            {
                var signVector = (vertexX > 0) ? ChildNodeMajorRadius : -ChildNodeMajorRadius;
                var makeSureDoublesAreZeroScalar = (Math.Abs(vertexX) < _doubleTolerance) ? 0.0 : 1.0;
                vectorAddendX = ((keyOutcome.childOutcomes.Count - 2) * .5) * signVector * makeSureDoublesAreZeroScalar;
            }
            else if (Math.Abs(vertexX) < _doubleTolerance)
            {
                var signVector = (vertexY > 0) ? ChildNodeMinorRadius : -ChildNodeMinorRadius;
                var makeSureDoublesAreZeroScalar = (Math.Abs(vertexY) < _doubleTolerance) ? 0.0 : 1.0;
                vectorSubtrahendY = (keyOutcome.childOutcomes.Count * .15) * signVector * makeSureDoublesAreZeroScalar;
            }

            var vectorOffsetX = (2 * vertexX) + vectorAddendX;
            var vectorOffsetY = (2 * vertexY) - vectorSubtrahendY;
            var bulbCenterX = XCenter + vectorOffsetX;
            var bulbCenterY = YCenter + vectorOffsetY;
            var oversizedBulbNode = keyOutcome.childOutcomes.Count > 5;

            DrawKeyOutcomeConnectorAndBulb(vertexColorPair, bulbCenterX, bulbCenterY, oversizedBulbNode);

            var xOffsetAngleRadians = GetOffsetAngleRadians(keyOutcome, vertexColorPair);

            var ellipseVectors = CalculateEllipseBranchVertices(keyOutcome, xOffsetAngleRadians);
            for (var index = 0; index < keyOutcome.childOutcomes.Count; index++)
            {
                var keyOutcomeChildOutcome = keyOutcome.childOutcomes[index];
                var keyOutcomeChildVector = ellipseVectors[index];

                var insideEllipseNodeX1 = bulbCenterX + keyOutcomeChildVector.X - ChildNodeMajorRadius;
                var insideEllipseNodeY1 = bulbCenterY + keyOutcomeChildVector.Y + ChildNodeMinorRadius;
                var insideEllipseNodeX2 = bulbCenterX + keyOutcomeChildVector.X + ChildNodeMajorRadius;
                var insideEllipseNodeY2 = bulbCenterY + keyOutcomeChildVector.Y - ChildNodeMinorRadius;
                var insideEllipseColor = DefaultEllipseColor;

                DrawEllipse(keyOutcomeChildOutcome, insideEllipseNodeX1, insideEllipseNodeY1, insideEllipseNodeX2,
                    insideEllipseNodeY2, insideEllipseColor);
            }
        }

        /// <summary>
        /// First calculates the tangent vector of the point on the ellipse.
        ///
        /// Then depending on the count of the child outcomes, will pad the result so the outcomes better line up with the node.
        /// </summary>
        /// <param name="keyOutcome"></param>
        /// <param name="vertexColorPair"></param>
        /// <returns></returns>
        protected double GetOffsetAngleRadians(Outcome keyOutcome, VertexColorPair vertexColorPair)
        {
            var p1 = new PointF(0,0);
            var p2 = vertexColorPair.Point;

            var deltaY = (p2.Y - p1.Y);
            var deltaX = (p2.X - p1.X);
            var xOffsetAngleRadians = Math.Atan2(deltaY, deltaX);

            if (keyOutcome.childOutcomes.Count == 2)
            {
                xOffsetAngleRadians += Math.PI  /2.0;
            }
            else if (keyOutcome.childOutcomes.Count > 4)
            {
                xOffsetAngleRadians += Math.PI + Math.PI * (1.0 / (2.0 * keyOutcome.childOutcomes.Count));
            }

            return xOffsetAngleRadians;
        }

        /// <summary>
        /// Creates a connection line and destination bulb (with matching color as key outcome node)
        ///
        /// Bulb can be larger if count of nodes extends past the outline of the standard node (usually count > 5 nodes)
        /// </summary>
        /// <param name="vertexColorPair"></param>
        /// <param name="bulbCenterX"></param>
        /// <param name="bulbCenterY"></param>
        /// <param name="oversizedBulbNode"></param>
        private void DrawKeyOutcomeConnectorAndBulb(VertexColorPair vertexColorPair, double bulbCenterX,
            double bulbCenterY, bool oversizedBulbNode)
        {
            var connectorLine = ActivePage.DrawLine(XCenter + vertexColorPair.X, YCenter + vertexColorPair.Y,
                bulbCenterX, bulbCenterY);
            connectorLine.SendToBack();

            var offset = oversizedBulbNode ? .2 : 0.0;

            var ellipseNodeX1 = bulbCenterX - (ChildNodeMajorRadius + offset);
            var ellipseNodeY1 = bulbCenterY + (ChildNodeMinorRadius + offset);
            var ellipseNodeX2 = bulbCenterX + (ChildNodeMajorRadius + offset);
            var ellipseNodeY2 = bulbCenterY - (ChildNodeMinorRadius + offset);
            var ellipseColor = "RGB(" + vertexColorPair.Color.R + ", " + vertexColorPair.Color.G + ", " +
                               vertexColorPair.Color.B + ")";

            DrawEllipse(null, ellipseNodeX1, ellipseNodeY1, ellipseNodeX2, ellipseNodeY2, ellipseColor);
        }

        /// <summary>
        /// Generic Implementation of child node drawing
        /// </summary>
        /// <param name="keyOutcome"></param>
        /// <param name="ellipseNodeX1"></param>
        /// <param name="ellipseNodeY1"></param>
        /// <param name="ellipseNodeX2"></param>
        /// <param name="ellipseNodeY2"></param>
        /// <param name="ellipseColor"></param>
        private void DrawEllipse(Outcome keyOutcome, double ellipseNodeX1, double ellipseNodeY1, double ellipseNodeX2,
            double ellipseNodeY2, string ellipseColor)
        {
            var node = ActivePage.DrawOval(ellipseNodeX1, ellipseNodeY1, ellipseNodeX2, ellipseNodeY2);
            node.CellsU["Fillforegnd"].FormulaU = ellipseColor;

            var shouldInvertTextColors = BackgroundColorIsToDark(ellipseColor);

            if (shouldInvertTextColors)
            {
                node.Characters.set_CharProps((short) VisCellIndices.visCharacterColor,
                    (short)VisDefaultColors.visWhite);
                //node.CellsU["Chars.Color"].FormulaU = "RGB(255, 255, 255)";
            } 

            if (keyOutcome != null && !string.IsNullOrEmpty(keyOutcome.title))
            {
                var keyOutcomeTitle = keyOutcome.title;
                node.Text = keyOutcomeTitle;
                var cellFontSizePt = CalculateSizeFontForChildNodeFromText(keyOutcomeTitle);

                node.Cells["Char.Size"].FormulaU = $"{cellFontSizePt} pt";
            }

            if (keyOutcome != null && !string.IsNullOrEmpty(keyOutcome.contentUrl))
            {
                var nodeHyperlink = node.AddHyperlink();
                nodeHyperlink.Address = keyOutcome.contentUrl;
            }
        }

        /// <summary>
        /// Will return true if the RGB value averages less than half of the color spectrum
        /// </summary>
        /// <param name="ellipseColor"></param>
        /// <returns></returns>
        private static bool BackgroundColorIsToDark(string ellipseColor)
        {
            ellipseColor = ellipseColor.Replace(" ", "");
            const string rgbSearchString = @"\d{1,3},\d{1,3},\d{1,3}";

            var match = Regex.Match(ellipseColor, rgbSearchString);

            if (!match.Success) return false;

            var rgbParsed = match.Groups[0].Value;
            var rgbAverage = rgbParsed.Split(',').Select(int.Parse).Average();

            return rgbAverage < 127;
        }

        /// <summary>
        /// Created by guessing and checking a few test cases for drawing.
        /// </summary>
        /// <param name="keyOutcomeTitle"></param>
        /// <returns></returns>
        private static int CalculateSizeFontForChildNodeFromText(string keyOutcomeTitle)
        {
            var textLength = keyOutcomeTitle.Length;
            var firstWordLength = keyOutcomeTitle.Contains(" ") 
                ? keyOutcomeTitle.Substring(0, keyOutcomeTitle.IndexOf(" ", StringComparison.Ordinal)).Length
                : keyOutcomeTitle.Length;
            var lastWordLength = keyOutcomeTitle.Contains(" ") 
                ? keyOutcomeTitle.Substring(keyOutcomeTitle.LastIndexOf(" ", StringComparison.Ordinal), keyOutcomeTitle.Length - keyOutcomeTitle.LastIndexOf(" ", StringComparison.Ordinal)).Length - 1
                : keyOutcomeTitle.Length;

            var edgeLength = (firstWordLength > lastWordLength) ? firstWordLength : lastWordLength;

            var cellFontSizePt = 12;

            if (edgeLength > 9)
            {
                cellFontSizePt -= 1;
            }
            
            if (textLength > 54)
            {
                cellFontSizePt = 9;
                if (edgeLength > 8)
                {
                    cellFontSizePt -= 1;
                }
                else if (edgeLength > 10)
                {
                    cellFontSizePt -= 2;
                }
            } 
            else if (textLength > 43)
            {
                cellFontSizePt = 10;
                if (edgeLength > 8)
                {
                    cellFontSizePt -= 1;
                }
                else if (edgeLength > 10)
                {
                    cellFontSizePt -= 2;
                }
            } 
            else if (textLength > 33)
            {
                cellFontSizePt = 11;
                if (edgeLength > 9)
                {
                    cellFontSizePt -= 1;
                }
            }

            return cellFontSizePt;
        }
    }
}