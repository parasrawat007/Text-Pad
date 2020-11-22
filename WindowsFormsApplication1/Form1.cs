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
using System.Speech.Synthesis;
using System.Speech.Recognition;


namespace WindowsFormsApplication1
{


    public partial class Form1 : Form
    {
        public int n = 1;
        float zoom;
        public String text;
        RichTextBox r;
        Panel p;
        Panel p1;
        Form2 f = new Form2();
       

        public Form1()
        {
            InitializeComponent();
            
        }
        
        //For new tab
        private void newToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (tabControl1.TabCount == 0)
                n = 1;

            String s = "new " + n.ToString();
            //objects
            //
            //
            r = new RichTextBox();
            r.Name = "r1";
             p = new Panel();
            p1 = new Panel();
            TabPage tb = new TabPage(s);            
            

            // 
            // tabPage1
            // 
            tb.Controls.Add(p1);
            tb.Controls.Add(p);
            tb.Location = new System.Drawing.Point(4, 22);
            tb.Padding = new System.Windows.Forms.Padding(3);
            tb.Size = new System.Drawing.Size(1047, 409);
            tb.TabIndex = 0;
            tb.UseVisualStyleBackColor = true;
            // 
            // panel1
            // 
            p.Dock = System.Windows.Forms.DockStyle.Left;
            p.Location = new System.Drawing.Point(3, 3);
            p.Size = new System.Drawing.Size(38, 403);
            p.TabIndex = 0;
            // 
            // panel2
            // 
            p1.Controls.Add(r);
            p1.Dock = System.Windows.Forms.DockStyle.Fill;
            p1.Location = new System.Drawing.Point(47, 3);
            p1.Size = new System.Drawing.Size(997, 403);
            p1.TabIndex = 1;
            // 
            // richTextBox1
            // 
            r.Dock = System.Windows.Forms.DockStyle.Fill;
            r.Location = new System.Drawing.Point(0, 0);
            r.Size = new System.Drawing.Size(997, 403);
            r.TabIndex = 0;
            r.Name = "Page";

            


            //other
            r.Font = new Font("Microsoft Sans Serif", 18);
            tabControl1.TabPages.Add(tb);
            tabControl1.SelectTab(tb);
            tabControl1.SelectedTab.Focus();
            ++n;
        }


        // to remove tab

        private void closToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if(tabControl1.TabCount!=0)
            tabControl1.TabPages.Remove(tabControl1.SelectedTab);           
        }

        //to close all tabes
        private void closeAllToolStripMenuItem_Click(object sender, EventArgs e)
        {
            tabControl1.TabPages.Clear();
        }

        //To save
        private void saveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            saveFileDialog1.FileName = tabControl1.SelectedTab.Text;
            
            if (saveFileDialog1.CheckPathExists)
            {
                RichTextBox o = null;
                label1.Text = openFileDialog1.FileName;
                label2.Text = tabControl1.SelectedTab.Text;
                if (p1.Controls.Contains(r))
                    o = p1.Controls["Page"] as RichTextBox;
                o.SaveFile(saveFileDialog1.FileName);

            }
            else {
                label1.Text = openFileDialog1.FileName;
                label2.Text = tabControl1.SelectedTab.Text;
                saveFileDialog1.ShowDialog();
            }
        }
        //to save
        private void saveFileDialog1_FileOk(object sender, CancelEventArgs e)
        {
           
            string name = saveFileDialog1.FileName;
            
            if (tabControl1.SelectedTab.Contains(r) == true)
            {
                try
                {
                    RichTextBox o = null;

                    if (p1.Controls.Contains(r))
                        o = p1.Controls["Page"] as RichTextBox;
                    o.SaveFile(name);
                }

                catch (Exception ex)
                {

                    MessageBox.Show(ex.Message);
                }
            }
        }

      
        //find
        private void findToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.AddOwnedForm(f);
            f.Show();
        }

        private void styleToolStripMenuItem_Click(object sender, EventArgs e)
        {
        }

        private void fontDialog1_Apply(object sender, EventArgs e)
        {
            
        }
        //to exit
        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            RichTextBox o = null;

            if (p1.Controls.Contains(r))
                o = p1.Controls["Page"] as RichTextBox;
            if (o.Text.Length != 0)
            {

                MessageBox.Show("Ae you Sure You Want to Exit");
                
               
            }
            else
            {
                System.Windows.Forms.Application.Exit();
            }
        }

        private void wordWrapToolStripMenuItem_Click(object sender, EventArgs e)
        {
            RichTextBox o = null;

            if (p1.Controls.Contains(r))
                o = p1.Controls["Page"] as RichTextBox;

            if (wordWrapToolStripMenuItem.Checked)
                o.WordWrap = true;
            else
                o.WordWrap = false;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            newToolStripMenuItem.PerformClick();
            zoom = r.ZoomFactor;
            f.MdiParent=this;
         }

        private void statusStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }

        private void fontColorToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (colorDialog1.ShowDialog() != DialogResult.Cancel) {
                RichTextBox o = null;

                if (p1.Controls.Contains(r))
                    o = p1.Controls["Page"] as RichTextBox;
                o.ForeColor = colorDialog1.Color;
            }
            
        }

        private void backgroundToolStripMenuItem_Click(object sender, EventArgs e)
        {
            
        }

        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {
            if (fontDialog1.ShowDialog() != DialogResult.Cancel)
            {
                RichTextBox o = null;

                if (p1.Controls.Contains(r))
                    o = p1.Controls["Page"] as RichTextBox;

                o.Font = fontDialog1.Font;
                o.ForeColor = fontDialog1.Color;
            }
        }

        private void backgroundToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            if (colorDialog1.ShowDialog() != DialogResult.Cancel)
            {
                RichTextBox o = null;

                if (p1.Controls.Contains(r))
                    o = p1.Controls["Page"] as RichTextBox;
                o.BackColor = colorDialog1.Color;
            }
        }

        private void undoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            RichTextBox o = null;

            if (p1.Controls.Contains(r))
                o = p1.Controls["Page"] as RichTextBox;
            toolStripMenuItem1.Text = o.GetType().ToString();          
            o.Undo();
        }

        private void redoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            RichTextBox o = null;

            if (p1.Controls.Contains(r))
                o = p1.Controls["Page"] as RichTextBox;
            toolStripMenuItem1.Text = o.GetType().ToString();
            o.Redo();
        }

        private void cutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            RichTextBox o = null;

            if (p1.Controls.Contains(r))
                o = p1.Controls["Page"] as RichTextBox;

            toolStripMenuItem1.Text = o.GetType().ToString();
            o.Cut();
        }

        private void copyToolStripMenuItem_Click(object sender, EventArgs e)
        {
            RichTextBox o = null;

            if (p1.Controls.Contains(r))
                o = p1.Controls["Page"] as RichTextBox;

            toolStripMenuItem1.Text = o.GetType().ToString();
            o.Copy();
        }

        private void pasteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            RichTextBox o = (RichTextBox)this.tabControl1.SelectedTab.GetNextControl(p1, true);

            toolStripMenuItem1.Text = o.GetType().ToString();
            o.Paste();
        }

        private void selectAllToolStripMenuItem_Click(object sender, EventArgs e)
        {
            RichTextBox o = (RichTextBox)this.tabControl1.SelectedTab.GetNextControl(p1, true);

            toolStripMenuItem1.Text = o.GetType().ToString();
            o.SelectAll();
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {

        }

        private void newToolStripButton_Click(object sender, EventArgs e)
        {
            newToolStripMenuItem.PerformClick();
        }

        private void openToolStripButton_Click(object sender, EventArgs e)
        {
            openFileDialog1.ShowDialog();
        }

        private void saveToolStripButton_Click(object sender, EventArgs e)
        {
            saveToolStripMenuItem.PerformClick();
        }

        private void printToolStripButton_Click(object sender, EventArgs e)
        {
            printToolStripMenuItem.PerformClick();
        }

        private void cutToolStripButton_Click(object sender, EventArgs e)
        {
            cutToolStripMenuItem.PerformClick();
        }

        private void copyToolStripButton_Click(object sender, EventArgs e)
        {
            copyToolStripMenuItem.PerformClick();
        }

        private void pasteToolStripButton_Click(object sender, EventArgs e)
        {
            pasteToolStripMenuItem.PerformClick();
        }

        private void helpToolStripButton_Click(object sender, EventArgs e)
        {
            hElpToolStripMenuItem.PerformClick();
        }

        private void toolStripButton1_Click_1(object sender, EventArgs e)
        {
            RichTextBox o = (RichTextBox)this.tabControl1.SelectedTab.GetNextControl(p1, true);
            //String t=o.SelectedText.
        }

        private void rightAlignToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (leftAlignToolStripMenuItem.Checked || justifyToolStripMenuItem.Checked) {
                leftAlignToolStripMenuItem.Checked = false;
                justifyToolStripMenuItem.Checked = false;
                RichTextBox o = null;

                if (p1.Controls.Contains(r))
                    o = p1.Controls["Page"] as RichTextBox;
                o.SelectionAlignment = HorizontalAlignment.Right;
            }
        }

        private void leftAlignToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (rightAlignToolStripMenuItem.Checked || justifyToolStripMenuItem.Checked)
            {
                rightAlignToolStripMenuItem.Checked = false;
                justifyToolStripMenuItem.Checked = false;
                RichTextBox o = (RichTextBox)this.tabControl1.SelectedTab.GetNextControl(p1, true);
                o.SelectionAlignment = HorizontalAlignment.Left;
            }
        }

        private void justifyToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (leftAlignToolStripMenuItem.Checked || rightAlignToolStripMenuItem.Checked)
            {
                leftAlignToolStripMenuItem.Checked = false;
                rightAlignToolStripMenuItem.Checked = false;
                RichTextBox o = null;

                if (p1.Controls.Contains(r))
                    o = p1.Controls["Page"] as RichTextBox;
                o.SelectionAlignment = HorizontalAlignment.Center;
            }
        }

        private void upperCaseToolStripMenuItem_Click(object sender, EventArgs e)
        {
            RichTextBox o = null;

            if (p1.Controls.Contains(r))
                o = p1.Controls["Page"] as RichTextBox;
            o.Text = o.Text.ToUpper();
        }

        private void lowerCAseToolStripMenuItem_Click(object sender, EventArgs e)
        {
            RichTextBox o = null;

            if (p1.Controls.Contains(r))
                o = p1.Controls["Page"] as RichTextBox;
            o.Text = o.Text.ToLower();
            }

        private void openToolStripMenuItem_Click(object sender, EventArgs e)
        {
            openFileDialog1.ShowDialog();          
        }

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {
            
            String file =openFileDialog1.FileName;
       
            newToolStripButton.PerformClick();
            tabControl1.SelectedTab.Text= openFileDialog1.SafeFileName;
            RichTextBox o = null;

            if (p1.Controls.Contains(r))
                o = p1.Controls["Page"] as RichTextBox;
            o.LoadFile(file);
            
            tabControl1.DeselectTab(tabControl1.SelectedTab);
        }

        private void printToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (printDialog1.ShowDialog() == DialogResult.OK) {
                
            
            }    
        }

        private void aboutUsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AboutBox1 a = new AboutBox1();
            a.Show();
        }

        private void speechToTextToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SpeechRecognitionEngine recognizer = new SpeechRecognitionEngine();
            //rec.
            Grammar dictationGrammar = new DictationGrammar();
            recognizer.LoadGrammar(dictationGrammar);
            try
            {
                
                recognizer.SetInputToDefaultAudioDevice();
                RecognitionResult result = recognizer.Recognize();
                label1.Text = result.Text;
            }
            catch (InvalidOperationException exception)
            {
                label1.Text = String.Format("Could not recognize input from default aduio device. Is a microphone or sound card available?\r\n{0} - {1}.", exception.Source, exception.Message);
            }
            finally
            {
                recognizer.UnloadAllGrammars();
            }                         
           
        }

        private void toolStripStatusLabel1_Click(object sender, EventArgs e)
        {
             RichTextBox o = null;
            
            if (p1.Controls.Contains(r))
                o = p1.Controls["Page"] as RichTextBox;
            toolStrip1.Text = o.Lines.ToString() ;
        }

        private void FontDecoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            RichTextBox o = null;

            if (p1.Controls.Contains(r))
                o = p1.Controls["Page"] as RichTextBox;
           
        }

        private void UnderlineToolStripMenuItem_Click(object sender, EventArgs e)
        {
            RichTextBox o = null;

            if (p1.Controls.Contains(r))
                o = p1.Controls["Page"] as RichTextBox;
            label1.Text = o.SelectedText;

            if (o.SelectionFont != new Font(o.Font, FontStyle.Underline))
            {
                underlineToolStripMenuItem.Checked = true;
                o.SelectionFont = new Font(o.Font, FontStyle.Underline);

            }
            else {

                underlineToolStripMenuItem.Checked = false;
               // o.SelectionFont = new Font(o.Font);
            }
            o.DeselectAll();

        }

        private void SaveasToolStripMenuItem_Click(object sender, EventArgs e)
        {
            saveFileDialog1.ShowDialog();
        }

        private void ZoomInToolStripMenuItem_Click(object sender, EventArgs e)
        {
            RichTextBox o = null;

            if (p1.Controls.Contains(r))
                o = p1.Controls["Page"] as RichTextBox;
          
                if (o.ZoomFactor < 10)
                    o.ZoomFactor += 0.10f;
            
        }

        private void ZoomOutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            RichTextBox o = null;

            if (p1.Controls.Contains(r))
                o = p1.Controls["Page"] as RichTextBox;
            
            if(o.ZoomFactor>0.50)
                o.ZoomFactor -= 0.10f;
            
            
        }

        private void HElpToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void ItalicToolStripMenuItem_Click(object sender, EventArgs e)
        {
            RichTextBox o = null;

            if (p1.Controls.Contains(r))
                o = p1.Controls["Page"] as RichTextBox;
            label1.Text = o.SelectedText;

            if (o.SelectionFont != new Font(o.Font, FontStyle.Italic))
            {
                italicToolStripMenuItem.Checked = true;
                o.SelectionFont = new Font(o.Font, FontStyle.Italic);

            }
            else
            {

                 italicToolStripMenuItem.Checked = false;
                // o.SelectionFont = new Font(o.Font);
            }
            o.DeselectAll();
        }

        private void NormalToolStripMenuItem_Click(object sender, EventArgs e)
        {
            RichTextBox o = null;
            if (p1.Controls.Contains(r))
                o = p1.Controls["Page"] as RichTextBox;
            o.ZoomFactor = zoom;
        }

        private void textToSpeechToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // Initialize a new instance of the SpeechSynthesizer.  
            SpeechSynthesizer synth = new SpeechSynthesizer();

            // Configure the audio output.   
            synth.SetOutputToDefaultAudioDevice();
           
            
            RichTextBox o = null;
            if (p1.Controls.Contains(r))
                o = p1.Controls["Page"] as RichTextBox;
            label1.Text = o.Text;
            
            // Speak a string.  
            synth.Speak(o.Text);

            
        }

        private void MenuHideToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (menuHideToolStripMenuItem.CheckState==CheckState.Checked)
            {
                toolStrip1.Visible = false;
                panel1.Visible = false;
                menuHideToolStripMenuItem.Checked = false;
            }
            else {
                panel1.Visible = true;
                toolStrip1.Visible = true;
                menuHideToolStripMenuItem.Checked = true;
            }
        }

        private void BoldToolStripMenuItem_Click(object sender, EventArgs e)
        {
            RichTextBox o = null;

            if (p1.Controls.Contains(r))
                o = p1.Controls["Page"] as RichTextBox;
            label1.Text = o.SelectedText;

            if (o.SelectionFont != new Font(o.Font, FontStyle.Bold))
            {
               boldToolStripMenuItem.Checked= true;
                o.SelectionFont = new Font(o.Font, FontStyle.Bold);

            }
            else
            {

                boldToolStripMenuItem.Checked = false;
                // o.SelectionFont = new Font(o.Font);
            }
            o.DeselectAll();
        }

        private void StikeOutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            RichTextBox o = null;

            if (p1.Controls.Contains(r))
                o = p1.Controls["Page"] as RichTextBox;
            label1.Text = o.SelectedText;

            if (o.SelectionFont != new Font(o.Font, FontStyle.Strikeout))
            {
                stikeOutToolStripMenuItem.Checked = true;
                o.SelectionFont = new Font(o.Font, FontStyle.Strikeout);

            }
            else
            {

                stikeOutToolStripMenuItem.Checked = false;
                // o.SelectionFont = new Font(o.Font);
            }
            o.DeselectAll();
        }
    }
}
