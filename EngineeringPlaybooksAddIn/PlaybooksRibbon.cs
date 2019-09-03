﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security;
using System.Text;
using System.Windows.Forms;
using EngineeringPlaybooksAddIn.Engines;
using Microsoft.Office.Tools.Ribbon;
using Newtonsoft.Json;

namespace EngineeringPlaybooksAddIn
{
    public partial class PlaybooksRibbon
    {

        private OpenFileDialog _openFileDialog;

        private void PlaybooksRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void BtnListShapes_Click(object sender, RibbonControlEventArgs e)
        {
            FormListShapes formListShapes = new FormListShapes();

            formListShapes.Show();
        }

        private void drayPlaybooksButton_Click(object sender, RibbonControlEventArgs e)
        {
            _openFileDialog = new OpenFileDialog
            {
                Title = FormsResource.OpenFileDialogTitle,
                AddExtension = true,
                Filter = FormsResource.playbooksFileTypeFilter
            };
            if (_openFileDialog.ShowDialog() != DialogResult.OK) return;
            string textResult = ParseFileForPlaybookJson();

            ProcessAndDrawPlaybookJson(textResult);
        }

        private string ParseFileForPlaybookJson()
        {
            string textResult = null;
            try
            {
                var sr = new StreamReader(_openFileDialog.FileName);
                textResult = sr.ReadToEnd();
            }
            catch (SecurityException ex)
            {
                MessageBox.Show(string.Format(FormsResource.OpenFileParseError, ex.Message, ex.StackTrace));
            }

            return textResult;
        }

        private void ProcessAndDrawPlaybookJson(string textResult)
        {
            object fun = JsonConvert.DeserializeObject(textResult);

            JciPlaybooksDrawingEngine eng = new JciPlaybooksDrawingEngine();
            eng.DrawPlaybookJson(textResult);
        }
    }
}