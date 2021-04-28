using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace VSML
{
    public partial class VSMLProgram : Form
    {
        private string fileName = "";


        public VSMLProgram()
        {
            InitializeComponent();
            this.Text = "VSML - New File";
            ZoomLabel.Text = "Zoom: " + MainText.ZoomFactor.ToString();
        }

        private void OpenFileBtn_Click(object sender, EventArgs e)
        {
            OpenFileDialog openfile = new OpenFileDialog();
            openfile.Title = "Open file:";
            openfile.Filter = "Text files (*.vsml)| *.vsml";
            openfile.RestoreDirectory = true;
            if (openfile.ShowDialog() == DialogResult.OK)
            {
                MainText.Clear();
                using (StreamReader sr = new StreamReader(openfile.FileName))
                {
                    MainText.Text = sr.ReadToEnd();
                    var t = MainText.Text;
                    MainText.Text = t;
                    sr.Close();
                }
            }
            this.fileName = openfile.FileName;
            this.Text = "VSML - " + Path.GetFileName(this.fileName);
        }

        private void SaveBtn_Click(object sender, EventArgs e)
        {
            // If empty filename
            if (this.fileName == "")
            {
                // Use save as functionality
                this.SaveAsBtn_Click(sender, e);
            }
            else
            {
                // Try and save the file
                try
                {
                    StreamWriter txtoutput = new StreamWriter(this.fileName);
                    txtoutput.Write(MainText.Text);
                    txtoutput.Close();
                }
                catch (System.ArgumentException)
                {
                    return;
                }
            }
        }

        private void SaveAsBtn_Click(object sender, EventArgs e)
        {
            SaveFileDialog savefile = new SaveFileDialog();
            savefile.Title = "Save file as..";
            savefile.Filter = "Text files (*.vsml)| *.vsml";
            if (savefile.ShowDialog() == DialogResult.OK)
            {
                StreamWriter txtoutput = new StreamWriter(savefile.FileName);
                txtoutput.Write(MainText.Text);
                txtoutput.Close();
            }
            this.fileName = savefile.FileName;
            this.Text = "VSML - " + Path.GetFileName(this.fileName);
        }

        private void ExportXmlBtn_Click(object sender, EventArgs e)
        {
            VsmlReader reader = new VsmlReader();
            var toXml = reader.VsmlToXml(MainText.Text);
            if (toXml.Item1)    //if successful and no errors
            {
                SaveFileDialog savefile = new SaveFileDialog
                {
                    Title = "Save file as..",
                    Filter = "XML Files|*.xml"
                };

                if (savefile.ShowDialog() == DialogResult.OK)
                {
                    StreamWriter txtoutput = new StreamWriter(savefile.FileName);
                    txtoutput.Write(reader.Document.OuterXml);
                    txtoutput.Close();
                }
                AuxText.Text = "Exported as XML to " + savefile.FileName + "\nExported content:\n" + reader.Document.OuterXml;
            }
            else
                AuxText.Text = "Error on line " + toXml.Item2.Item1 + " of the XML " + toXml.Item2.Item2 + reader.Document.OuterXml;
        }

        private void ExportHtmlBtn_Click(object sender, EventArgs e)
        {
            VsmlReader reader = new VsmlReader();
            VsmlWriter writer = new VsmlWriter(reader.Document);
            var toXml = reader.VsmlToXml(MainText.Text);
            if (toXml.Item1)    //if successful and no errors
            {
                SaveFileDialog savefile = new SaveFileDialog
                {
                    Title = "Save file as..",
                    Filter = "HTML Files|*.html"
                };

                if (savefile.ShowDialog() == DialogResult.OK)
                {
                    StreamWriter txtoutput = new StreamWriter(savefile.FileName);
                    txtoutput.Write(writer.ToHtml());
                    txtoutput.Close();
                }
                AuxText.Text = "Exported as HTML to " + savefile.FileName + "\nExported content:\n" + writer.ToHtml();
            }
            else
                AuxText.Text = "Error on line " + toXml.Item2.Item1 + " of the XML " + toXml.Item2.Item2 + reader.Document.OuterXml;
        }


        public void ExportDocxBtn_Click(object sender, EventArgs e)
        {
            VsmlReader reader = new VsmlReader();
            VsmlWriter writer = new VsmlWriter(reader.Document);
            var toXml = reader.VsmlToXml(MainText.Text);
            if (toXml.Item1)    //if successful and no errors
            {
                SaveFileDialog savefile = new SaveFileDialog
                {
                    Title = "Save file as..",
                    Filter = "Docx Files|*.Docx"
                };

                if (savefile.ShowDialog() == DialogResult.OK)
                {
                    var success = writer.ToDocx(savefile.FileName);
                    if (success.Item1)
                        AuxText.Text = "Exported .Docx version of XML to " + savefile.FileName + "\nExported content:\n" + reader.Document.OuterXml;
                    else
                        AuxText.Text = "Error: " + success.Item2;
                }
            }
            else
                AuxText.Text = "Error on line " + toXml.Item2.Item1 + " of the XML " + toXml.Item2.Item2 + reader.Document.OuterXml;
        }

        // Key shortcuts
        private void Form1_KeyDown(object sender, KeyEventArgs e)
        {
           
            if (e.Control & e.KeyCode == Keys.O)
            {
                this.OpenFileBtn_Click(sender, e);
                e.Handled = true;
            }
            if (e.Control & e.KeyCode == Keys.S)
            {
                this.SaveBtn_Click(sender, e);
                e.Handled = true;
            }
        }

        private void ZoomIn_Click(object sender, EventArgs e)
        {
            float zoom = MainText.ZoomFactor;
            if (zoom < 63)
            {
                MainText.ZoomFactor += (float)0.25;
            }
            ZoomLabel.Text = "Zoom: " + MainText.ZoomFactor.ToString();
        }

        private void ZoomOut_Click(object sender, EventArgs e)
        {
            float zoom = MainText.ZoomFactor;
            if (zoom > 1)
            {
                MainText.ZoomFactor -= (float).25;
            }
            ZoomLabel.Text = "Zoom: " + MainText.ZoomFactor.ToString();
        }
    }
}