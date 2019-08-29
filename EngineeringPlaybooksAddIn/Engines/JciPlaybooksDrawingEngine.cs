using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Visio = Microsoft.Office.Interop.Visio;

namespace EngineeringPlaybooksAddIn.Engines
{
    public class JciPlaybooksDrawingEngine
    {
        public void DrawPlaybookJson(string textResult)
        {
            DrawHeader();

            DrawMap(textResult);
        }

        public void DrawHeader()
        {
            Visio.Shape containingBox;
            Visio.Shape title;
            Visio.Shape description;


            var ActivePage = Globals.ThisAddIn.Application.ActivePage;

            containingBox = ActivePage.DrawRectangle(0.25, 8.25, 10.75, 7.133);
            title = ActivePage.DrawRectangle(1.25, 8.25, 10.75, 7.9);
            title.LineStyle = "Guide";

            description = ActivePage.DrawRectangle(1.25, 7.9, 10.75, 7.133);

            description.LineStyle = "Guide";

            title.Text = "Johnson Controls VnA Team Engineering Role";
            description.Text = "Basic set of Outcomes based specifically on what a All Engineers would want to know on the Valves and Actuators Team";

            title.CellsSRC[(short) Visio.VisSectionIndices.visSectionParagraph, 0, (short) Visio.VisHorizontalAlignTypes.visHorzAlignLeft].FormulaU = "0";

            description.CellsSRC[(short) Visio.VisSectionIndices.visSectionParagraph, 0,
                (short) Visio.VisHorizontalAlignTypes.visHorzAlignLeft].FormulaU = "0";

            description.CellsSRC[(short)Visio.VisSectionIndices.visSectionObject, (short)Visio.VisRowIndices.visRowText, 
                (short)Visio.VisCellIndices.visTxtBlkVerticalAlign].FormulaU = "0";
            description.Cells["Char.Size"].FormulaU = "10 pt";
        }

        private void DrawMap(string textResult)
        {
            throw new NotImplementedException();
        }

    }
}
