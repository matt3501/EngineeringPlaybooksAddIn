using System.Collections.Generic;
using System.Drawing;
using EngineeringPlaybooksAddIn.Models;
using Microsoft.Office.Interop.Visio;
using Newtonsoft.Json;
using Color = EngineeringPlaybooksAddIn.Models.Color;

namespace EngineeringPlaybooksAddIn.Engines
{
    public class JciPlaybooksDrawingEngine
    {
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

            var coreOval = ActivePage.DrawOval(xCenter - 1.5375, yCenter + 1.2775, xCenter + 1.5375, yCenter - 1.2775);
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

            var vertexColorPairs = new List<VertexColorPair>
            {
                new VertexColorPair(0.7125, 1.2275 ,new Color {R = 255, G = 176, B = 201} ),
                new VertexColorPair(-0.5875,1.6275 , new Color {R = 72,  G = 172, B = 198} ),
                new VertexColorPair(1.0125, 0.0275 ,new Color {R = 204, G = 196, B = 100} ),
                new VertexColorPair(-0.5875,-0.7725,  new Color {R = 179, G = 172, B = 96 } ),
                new VertexColorPair(-2.0375,-0.1025,  new Color {R = 150, G = 232, B = 255} ),
                new VertexColorPair(-2.0375,1.0275 , new Color {R = 255, G = 244, B = 112} )
            };

            return vertexColorPairs;
        }

        public void DrawChildNode(string nodeText, double xpos, double ypos, Color color, string nodeUrl)
        {
            var node = ActivePage.DrawOval(xpos, ypos, xpos + 1.064, ypos - 0.884);
            node.CellsU["Fillforegnd"].FormulaU = "RGB(" + color.R + ", " + color.G + ", " + color.B + ")";
            node.Text = nodeText;

            var nodeHyperlink = node.AddHyperlink();
            nodeHyperlink.Address = nodeUrl;
        }
    }

}






