using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace WindowsApplication1
{
    public partial class Form1 : Form
    {
        private uint clrHashForeColor;
        private uint clrHashBackColor;

        public Form1()
        {
            InitializeComponent();
        }

        public uint Color2Uint32(Color clr)
        {

            int t;
            byte[] a;

            t = ColorTranslator.ToOle(clr);

            a = BitConverter.GetBytes(t);

            return BitConverter.ToUInt32(a, 0);

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            chkautosaveallpage.Text = "Automate save to PDF" + "\r\n" + "when last page scanned";

            this.tiffCompressionComboBox.Items.AddRange(new string[] { "LZW", "CITT3", "CITT4", "RLE", "None" });
            this.tiffCompressionComboBox.SelectedIndex = 0;

            this.outputTypeComboBox.Items.AddRange(new string[] { "BMP", "JPG", "GIF", "PNG", "TIF", "PDF" });
            this.outputTypeComboBox.SelectedIndex = 0;

            this.dPIComboBox.Items.AddRange(new string[] { "Onscreen Viewing 96dpi", "Fax 200dpi", "OCR Text 300dpi", "Laser Print Fine 600dpi" });
            this.dPIComboBox.SelectedIndex = 0;

            this.pixelTypeComboBox.Items.AddRange(new string[] { "Default", "Gray Color", "Black & White Color", "True Color" });
            this.pixelTypeComboBox.SelectedIndex = 0;

            this.clrHashForeColor =this.Color2Uint32( Color.Black);
            this.clrHashBackColor = this.Color2Uint32(Color.White);

            short iCount = this.axScanner1.GetNumImageSources();

            for (short i = 0; i < iCount; i++)
            {
                this.imageSourceComboBox.Items.Add(this.axScanner1.GetImageSourceName(i));
            }

          

            this.ApplyDefaultCap();

        }

      

        private void textStyleComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
         }

        

        private void greyScaleCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            this.axScanner1.OutputGrayScale = this.greyScaleCheckBox.Checked;
        }

        private void showTWAIN_CheckedChanged(object sender, EventArgs e)
        {
            this.axScanner1.ShowTwainUI = this.showTWAIN.Checked;
        }

        private void clearImage_CheckedChanged(object sender, EventArgs e)
        {
            this.axScanner1.ClearImageBuffer = this.clearImage.Checked;
        }

        private void captureAreaCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            this.ApplyDefaultCap();
        }

        private void ApplyDefaultCap()
        {
            bool check = this.captureAreaCheckBox.Checked;

            this.captureLeft.Enabled = !check;
            this.captureTop.Enabled = !check;
            this.captureWidth.Enabled = !check;
            this.captureHeight.Enabled = !check;
        }

   



        private void rotate90Button_Click(object sender, EventArgs e)
        {
            this.axScanner1.Rotate90();
        }

        private void zoom200Button_Click(object sender, EventArgs e)
        {
            this.axScanner1.View = 7;
            this.axScanner1.Focus();
        }

        private void zoom33Button_Click(object sender, EventArgs e)
        {
            this.axScanner1.View = 2;
            this.axScanner1.Focus();
        }

        private void zoom25Button_Click(object sender, EventArgs e)
        {
            this.axScanner1.View = 1;
            this.axScanner1.Focus();
        }

        private void backColorButton_Click(object sender, EventArgs e)
        {
            using (ColorDialog cd = new ColorDialog())
            {
                if (cd.ShowDialog(this) == DialogResult.OK)
                {
                    this.axScanner1.SetBackgroundColor(this.Color2Uint32(cd.Color));
                }
            }
        }

        private void printButton_Click(object sender, EventArgs e)
        {
            this.axScanner1.PrintImage(true);
        }

        private void saveMultiPageTiffButton_Click(object sender, EventArgs e)
        {
            this.saveFileDialog1.Filter = "TIFF File (*.tif)|*.tif";
            this.saveFileDialog1.DefaultExt = "tif";

            this.axScanner1.View = 5;
            this.axScanner1.TIFCompression = (short)this.tiffCompressionComboBox.SelectedIndex;

            if (this.saveFileDialog1.ShowDialog(this) == DialogResult.OK)
            {
                bool a = this.axScanner1.SaveAllPage2TIF(saveFileDialog1.FileName, this.singleFileCheckBox.Checked, 1);
                if (a)
                {
                    MessageBox.Show("save to " + saveFileDialog1.FileName + " complete");
                }
            }
        }

        private void SaveToPDF()
        {
            this.saveFileDialog1.Filter = "PDF File (*.pdf)|*.pdf";
            this.saveFileDialog1.DefaultExt = "pdf";

            this.axScanner1.View = 5;
            this.axScanner1.TIFCompression = (short)this.tiffCompressionComboBox.SelectedIndex;

            if (this.saveFileDialog1.ShowDialog(this) == DialogResult.OK)
            {
                bool a = this.axScanner1.SaveAllPage2PDF(saveFileDialog1.FileName, this.singleFileCheckBox.Checked, 1);
                if (a)
                {
                    MessageBox.Show("save to " + saveFileDialog1.FileName + " complete");
                }
            }
        }
        private void saveMultiPagePDFButton_Click(object sender, EventArgs e)
        {
            SaveToPDF();
        }

        private void deleteButton_Click(object sender, EventArgs e)
        {
            this.axScanner1.DeletePage(this.axScanner1.GetActivePageNo());
            this.axScanner1.SetActivePageNo(1);
            RefreshPageNo();

        }

        private void RefreshPageNo()
        {
            this.totalPageLabel.Text = string.Format((string)this.totalPageLabel.Tag, this.axScanner1.TotalPage.ToString(), this.axScanner1.GetActivePageNo().ToString());
        }

        private void saveBySizeButton_Click(object sender, EventArgs e)
        {
            string strFile = "c:\\test";
            string strType="jpg";
            short a = this.axScanner1.SaveBySize(strFile,strType, (short)this.outputWidth.Value, (short)this.outputHeight.Value);

            if (a==1)
            {
                MessageBox.Show("Save" + strFile + "." + strType + " complete");
            }
            else
            {
                MessageBox.Show("Save fail");
            }
        }

        private void rotate180Button_Click(object sender, EventArgs e)
        {
            this.axScanner1.Rotate180();
        }

        private void ratioButton_Click(object sender, EventArgs e)
        {
            this.axScanner1.View = 10;
            this.axScanner1.Focus();
        }

        private void saveToClipboardButton_Click(object sender, EventArgs e)
        {
            this.axScanner1.Copy2Clipboard();
        }

        private void saveToPictureboxButton_Click(object sender, EventArgs e)
        {
            this.pictureBox1.Image = this.axScanner1.Copy2PictureBox();
        }

        private void saveToHBButton_Click(object sender, EventArgs e)
        {
            int h = this.axScanner1.Copy2HBITMAP();
            if (h != 0)
            {
                this.pictureBox1.Image = Image.FromHbitmap(new IntPtr(h));
            }
        }

        private void rotate270Button_Click(object sender, EventArgs e)
        {
            this.axScanner1.Rotate270();
        }

        private void goToPageButton_Click(object sender, EventArgs e)
        {
            this.axScanner1.SetActivePageNo((short)this.pageNumericUpDown.Value);
            this.RefreshPageNo();
        }

        private void updateButton_Click(object sender, EventArgs e)
        {
            this.axScanner1.ApplyChange();
        }

    



   
        private void scanButton_Click(object sender, EventArgs e)
        {
            if (this.manulRadioButton.Checked)
            {
                this.axScanner1.SelectImageSourceByIndex((short)this.imageSourceComboBox.SelectedIndex);
            }
            else
            {
                this.axScanner1.SelectImageSource();
            }

            this.axScanner1.DuplexEnabled = this.enableDuplex.Checked;
            this.axScanner1.FeederEnabled = this.enableFeeder.Checked;
            switch (this.dPIComboBox.SelectedIndex)
            {
                case 0:
                    this.axScanner1.DPI = 96;
                    break;
                case 1:
                    this.axScanner1.DPI = 200;
                    break;
                case 2:
                    this.axScanner1.DPI = 300;
                    break;
                case 3:
                    this.axScanner1.DPI = 600;
                    break;
            }


            this.axScanner1.PixelType = (short)(this.pixelTypeComboBox.SelectedIndex - 1);

            double left = 0;
            double top = 0;
            double width = 0;
            double height = 0;

            if (!this.captureAreaCheckBox.Checked)
            {
                left = (double)this.captureLeft.Value;
                top = (double)this.captureTop.Value;
                width = (double)this.captureWidth.Value;
                height = (double)this.captureHeight.Value;
            }

            this.axScanner1.SetCaptureArea(left, top, width, height);

            this.axScanner1.Scan();
        }

        private void saveToFileButton_Click(object sender, EventArgs e)
        {
            string strFile = "c:\\test";
            string strType = (string)this.outputTypeComboBox.SelectedItem;

            short result = this.axScanner1.Save(strFile, strType);
            if (result == 1)
            {
                MessageBox.Show(this, "Save " + strFile + "." + strType + " complete");
            }
            else
            {
                MessageBox.Show(this, "Save fail");
            }

            this.axScanner1.Focus();
   
        }

        private void fitButton_Click(object sender, EventArgs e)
        {
            this.axScanner1.View = 9;
            this.axScanner1.Focus();
        }

        private void zoom150Button_Click(object sender, EventArgs e)
        {
            this.axScanner1.View = 6;
            this.axScanner1.Focus();
        }

        private void zoom100Button_Click(object sender, EventArgs e)
        {
            this.axScanner1.View = 5;
            this.axScanner1.Focus();
        }

        private void zoom50Button_Click(object sender, EventArgs e)
        {
            this.axScanner1.View = 3;
            this.axScanner1.Focus();
        }

        private void zoom75Button_Click(object sender, EventArgs e)
        {
            this.axScanner1.View = 4;
            this.axScanner1.Focus();
        }

        private void SetBorder()
        {
            this.axScanner1.Border = this.borderOnRadioButton.Checked;
        }

        private void borderOnRadioButton_CheckedChanged(object sender, EventArgs e)
        {
            this.SetBorder();
        }

        private void hqYesRadioButton_CheckedChanged(object sender, EventArgs e)
        {
            this.SetHighQ();
        }

        private void SetHighQ()
        {
            this.axScanner1.HighQuality = this.hqYesRadioButton.Checked;
        }

    

        private void userDefineNumericUpDown_ValueChanged(object sender, EventArgs e)
        {
            this.axScanner1.View = 8;
            this.axScanner1.ViewSize = (double)(this.userDefineNumericUpDown.Value / 100);
        }

        private void leftLocation_ValueChanged(object sender, EventArgs e)
        {
            RefreshFileLocation();
        }

        private void RefreshFileLocation()
        {
            this.axScanner1.FileLeft = (int)this.leftLocation.Value;
            this.axScanner1.FileTop = (int)this.topLocation.Value;
            this.axScanner1.Redraw();
        }

        private void topLocation_ValueChanged(object sender, EventArgs e)
        {
            this.RefreshFileLocation();
        }

   

        private void axScanner1_EndScan_1(object sender, EventArgs e)
        {
            this.outputWidth.Value = (decimal)this.axScanner1.FileWidth;
            this.outputHeight.Value = (decimal)this.axScanner1.FileHeight;

           

            this.RefreshPageNo();
        }
        private void axScanner1_ScanningError(object sender, EventArgs e)
        {
            listBox1.Items.Add("Error occur, may be paper jam");
        }
        private void axScanner1_EndAllScan(object sender, EventArgs e)
        {
            if (chkautosaveallpage.Checked)
            {
                SaveToPDF();
            }
        }
        private void button1_Click(object sender, EventArgs e)
        {

            if (this.axScanner1.BlankPageIsBlank((double)numericconf.Value))
                MessageBox.Show("This page is blank page");
            else
                MessageBox.Show("This page is not a blank page");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            numericconf.Value = (decimal)axScanner1.BlankPageGetConfidence();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.saveFileDialog1.Filter = "Ms Word File (*.docx)|*.docx";
            this.saveFileDialog1.DefaultExt = "docx";

            this.axScanner1.View = 5;
           
            if (this.saveFileDialog1.ShowDialog(this) == DialogResult.OK)
            {
                bool a = this.axScanner1.SaveAllPage2Docx(saveFileDialog1.FileName, this.singleFileCheckBox.Checked, 1);
                if (a)
                {
                    MessageBox.Show("save to " + saveFileDialog1.FileName + " complete");
                }
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "PDF (*.pdf)|*.pdf|Open XML Paper Specification (*.xps *.oxps)|*.xps;*.oxps|TIFF (*.tiff)|*.tif|JPEG (*.jpg)|*.jpg|BMP (*.bmp)|*.bmp|PNG (*.png)|*.png|GIF (*.gif)|*.gif";

            if (chkpanning.Checked)
                axScanner1.EnablePanning = true;
            else
                axScanner1.EnablePanning = false;
        

            if (this.openFileDialog1.ShowDialog(this) == DialogResult.OK)
            {
                  textBox1.Text = openFileDialog1.FileName;
                  axScanner1.LoadImage(openFileDialog1.FileName);

                  this.totalPageLabel.Text = string.Format((string)this.totalPageLabel.Tag, this.axScanner1.TotalPage.ToString(), this.axScanner1.GetActivePageNo().ToString());
                  pageNumericUpDown.Value = 1;
     
            }
        }

      
     


    }
}