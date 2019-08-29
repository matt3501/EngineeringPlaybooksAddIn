using System;
using EngineeringPlaybooksAddIn.Models;
using Newtonsoft.Json;
using Visio = Microsoft.Office.Interop.Visio;

namespace EngineeringPlaybooksAddIn.Engines
{
    public class JciPlaybooksDrawingEngine
    {
        private Visio.Page ActivePage;

        public void DrawPlaybookJson(string textResult)
        {
            KnowledgeModel model = JsonConvert.DeserializeObject<KnowledgeModel>(textResult);

            DrawHeader(model);

            DrawMap(textResult);
        }

        public void DrawHeader(KnowledgeModel model)
        {
            Visio.Shape containingBox;
            Visio.Shape title;
            Visio.Shape description;

            //Visio.VisPageSizingBehaviors

            ActivePage = Globals.ThisAddIn.Application.ActivePage;

            //var ActiveDocument = Globals.ThisAddIn.Application.ActiveDocument;
            //var ActiveDocument = Globals.ThisAddIn.Application.ActiveDocument;

            //ActiveDocument.PrintLandscape = true;

            SetOrientationToLandscape();

            //Globals.ThisAddIn.Application.

                ;

            containingBox = ActivePage.DrawRectangle(0.25, 8.25, 10.75, 7.133);
            title = ActivePage.DrawRectangle(1.25, 8.25, 10.75, 7.9);
            title.LineStyle = "Guide";

            description = ActivePage.DrawRectangle(1.25, 7.9, 10.75, 7.133);

            description.LineStyle = "Guide";

            title.Text = model.title;
            description.Text = model.description;

            title.CellsSRC[(short) Visio.VisSectionIndices.visSectionParagraph, 0, 
                (short) Visio.VisCellIndices.visHorzAlign].FormulaU = "0";

            description.CellsSRC[(short) Visio.VisSectionIndices.visSectionParagraph, 0,
                (short) Visio.VisCellIndices.visHorzAlign].FormulaU = "0";

            description.CellsSRC[(short)Visio.VisSectionIndices.visSectionObject, 
                (short)Visio.VisRowIndices.visRowText, 
                (short)Visio.VisCellIndices.visTxtBlkVerticalAlign].FormulaU = "0";
            description.Cells["Char.Size"].FormulaU = "10 pt";
        }

        private void SetOrientationToLandscape()
        {
            var orientation = ActivePage.PageSheet.CellsSRC[(short) Visio.VisSectionIndices.visSectionObject,
                (short) Visio.VisRowIndices.visRowPrintProperties,
                (short) Visio.VisCellIndices.visPrintPropertiesPageOrientation].FormulaU;
            
            //If orientation flag isn't set to 'landscape'
            if (orientation != ((int)Visio.tagVisCellVals.visPPOLandscape).ToString())
            {
                //Landscape
                ActivePage.PageSheet.CellsSRC[(short) Visio.VisSectionIndices.visSectionObject,
                    (short) Visio.VisRowIndices.visRowPrintProperties,
                    (short) Visio.VisCellIndices.visPrintPropertiesPageOrientation].FormulaU = ((int)Visio.tagVisCellVals.visPPOLandscape).ToString();

                var currentWidth = ActivePage.PageSheet.CellsSRC[(short) Visio.VisSectionIndices.visSectionObject,
                    (short) Visio.VisRowIndices.visRowPage, (short) Visio.VisCellIndices.visPageWidth].FormulaU;
                var currentHeight = ActivePage.PageSheet.CellsSRC[(short) Visio.VisSectionIndices.visSectionObject,
                    (short) Visio.VisRowIndices.visRowPage, (short) Visio.VisCellIndices.visPageHeight].FormulaU;


                ActivePage.PageSheet.CellsSRC[(short) Visio.VisSectionIndices.visSectionObject,
                    (short) Visio.VisRowIndices.visRowPage, (short) Visio.VisCellIndices.visPageWidth].FormulaU = currentHeight;
                ActivePage.PageSheet.CellsSRC[(short) Visio.VisSectionIndices.visSectionObject,
                    (short) Visio.VisRowIndices.visRowPage, (short) Visio.VisCellIndices.visPageHeight].FormulaU = currentWidth;
            }
        }

        private void DrawMap(string textResult)
        {
            //throw new NotImplementedException();
        }

    }
}
