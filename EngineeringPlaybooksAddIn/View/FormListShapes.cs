using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Visio = Microsoft.Office.Interop.Visio;

namespace EngineeringPlaybooksAddIn
{
    public partial class FormListShapes : Form
    {
        public FormListShapes()
        {
            InitializeComponent();

            foreach (Visio.Shape shape in Globals.ThisAddIn.Application.ActivePage.Shapes)
            {
                lstShapes.Items.Add(shape.Text + " (" + shape.Name + ")");
            }
        }
    }
}
