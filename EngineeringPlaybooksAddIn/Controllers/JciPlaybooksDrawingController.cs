using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using EngineeringPlaybooksAddIn.Models;
using Microsoft.Office.Interop.Visio;
using Newtonsoft.Json;
using Color = EngineeringPlaybooksAddIn.Models.Color;

namespace EngineeringPlaybooksAddIn.Controllers
{
    public class JciPlaybooksDrawingController
    {
        private const double MajorRadius = 1.5375;
        private const double MinorRadius = 1.2775;
        private static readonly IList<Color> JciColors = new ReadOnlyCollection<Color>(new List<Color>()
        {
            new Color(0, 128, 182),
            new Color(0, 183, 168),
            new Color(106, 193, 123),
            new Color(214, 213, 37),
            new Color(254, 189, 56),
            new Color(244, 119, 33),
            new Color(203, 36, 57),
        });

        /// <summary>
        /// Allows the calling into ActivePage similar to the syntax used by traditional VB6 macros
        /// </summary>
        // ReSharper disable once InconsistentNaming
        private Page ActivePage;

        public void DrawPlaybookJson(string textResult)
        {
            ActivePage = Globals.ThisAddIn.Application.ActivePage;

            var model = JsonConvert.DeserializeObject<KnowledgeModel>(textResult);

            SetOrientationToLandscape();

            DrawHeader(model);

            DrawMap(model);
        }

        public void DrawHeader(KnowledgeModel model)
        {
            ActivePage.DrawRectangle(0.25, 8.25, 10.75, 7.133);
            var title = ActivePage.DrawRectangle(1.25, 8.25, 10.75, 7.9);
            title.LineStyle = "Guide";
            title.Text = model.title;

            var description = ActivePage.DrawRectangle(1.25, 7.9, 10.75, 7.133);
            description.LineStyle = "Guide";
            description.Text = model.description;

            title.CellsSRC[(short) VisSectionIndices.visSectionParagraph, 0,
                (short) VisCellIndices.visHorzAlign].FormulaU = "0";

            description.CellsSRC[(short) VisSectionIndices.visSectionParagraph, 0,
                (short) VisCellIndices.visHorzAlign].FormulaU = "0";

            description.CellsSRC[(short) VisSectionIndices.visSectionObject,
                (short) VisRowIndices.visRowText,
                (short) VisCellIndices.visTxtBlkVerticalAlign].FormulaU = "0";
            description.Cells["Char.Size"].FormulaU = "10 pt";
        }

        /// <summary>
        ///     Sets the currently open document to landscape (if not already set)
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
            var ellipseVertices = GetEllipseVertices(model);

            var xCenter = 5.2875;
            var yCenter = 3.9725;


            var coreOval = ActivePage.DrawOval(xCenter - MajorRadius, yCenter + MinorRadius, xCenter + MajorRadius, yCenter - MinorRadius);
            coreOval.CellsU["Fillforegnd"].FormulaU = "RGB(248, 248, 248)";
            coreOval.Text = "Key Outcomes";

            for (var index = 0; index < model.outcomes.Count; index++)
            {
                var outcome = model.outcomes[index];
                var xOffset = ellipseVertices[index].X;
                var yOffset = ellipseVertices[index].Y;
                var color = ellipseVertices[index].Color;
                DrawChildNode(outcome.title, xCenter + xOffset, yCenter + yOffset, color, outcome.contentUrl);
            }
        }
        
        private List<VertexColorPair> GetEllipseVertices(KnowledgeModel model)
        {
            var count = model.outcomes.Count;

            var geometry = GeometryController.GetPointsForEllipse(MajorRadius, MinorRadius, count, GeometryController.RotationStarts.RotationStartsAtAxisX);

            return geometry.Select((point, index) => new VertexColorPair(Math.Round(point.X, 2), Math.Round(point.Y, 2), JciColors[index])).ToList();
        }

        public void DrawChildNode(string nodeText, double xCenter, double yCenter, Color color, string nodeUrl)
        {
            var childNodeMajorRadius = 0.532;
            var childNodeMinorRadius = 0.442;

            var node = ActivePage.DrawOval(xCenter - childNodeMajorRadius, yCenter + childNodeMinorRadius, xCenter + childNodeMajorRadius, yCenter - childNodeMinorRadius);
            node.CellsU["Fillforegnd"].FormulaU = "RGB(" + color.R + ", " + color.G + ", " + color.B + ")";
            node.Text = nodeText;

            var nodeHyperlink = node.AddHyperlink();
            nodeHyperlink.Address = nodeUrl;
        }
    }

}






