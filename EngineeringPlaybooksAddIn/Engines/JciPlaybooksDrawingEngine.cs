using EngineeringPlaybooksAddIn.Models;
using Microsoft.Office.Interop.Visio;
using Newtonsoft.Json;

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
            var xpos = 3.75;
            var ypos = 5.25;

            var coreOval = ActivePage.DrawOval(xpos, ypos, xpos + 3.075, ypos - 2.555);
            coreOval.CellsU["Fillforegnd"].FormulaU = "RGB(248, 248, 248)";
            coreOval.Text = "Key Outcomes";

            DrawChildNode("Pursue Design", 6, 5.2, true, "255", "176", "201", "https://www.google.com");
            DrawChildNode("Gather Requirements", 4.7, 5.6, true, "72", "172", "198", "https://www.google.com");
            DrawChildNode("Develop Solution", 6.3, 4, true, "204", "196", "100", "https://www.google.com");
            DrawChildNode("Deploy Product", 4.7, 3.2, true, "179", "172", "96", "https://www.google.com");
            DrawChildNode("Engineer Titles", 3.25, 3.87, true, "150", "232", "255", "https://www.google.com");
            DrawChildNode("Product Design Life cycle", 3.25, 5, true, "255", "244", "112", "https://www.google.com");
        }

        public void DrawChildNode(string nodeText, double xpos, double ypos, bool isBig, string colorR, string colorG,
            string colorB, string nodeUrl)
        {
            var node = ActivePage.DrawOval(xpos, ypos, xpos + 1.064, ypos - 0.884);
            node.CellsU["Fillforegnd"].FormulaU = "RGB(" + colorR + ", " + colorG + ", " + colorB + ")";
            node.Text = nodeText;

            var nodeHyperlink = node.AddHyperlink();
            nodeHyperlink.Address = nodeUrl;
        }
    }
}