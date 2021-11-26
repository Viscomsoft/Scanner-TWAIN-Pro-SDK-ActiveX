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
            this.tiffCompressionComboBox.Items.AddRange(new string[] { "LZW", "CITT3", "CITT4", "RLE", "None","JPEG" });
            this.tiffCompressionComboBox.SelectedIndex = 0;

            this.cbocurrtifcomp.Items.AddRange(new string[] { "LZW", "CITT3", "CITT4", "RLE", "None", "JPEG" });
            this.cbocurrtifcomp.SelectedIndex = 0;

            this.outputTypeComboBox.Items.AddRange(new string[] { "BMP", "JPG", "GIF", "PNG", "TIF", "PDF", "PCX", "TGA", "JP2", "JPC","RAS","PGX","PNM" });
            this.outputTypeComboBox.SelectedIndex = 0;

            this.dPIComboBox.Items.AddRange(new string[] { "Onscreen Viewing 96dpi", "Fax 200dpi", "OCR Text 300dpi", "Laser Print Fine 600dpi" });
            this.dPIComboBox.SelectedIndex = 0;

            this.pixelTypeComboBox.Items.AddRange(new string[] { "Default", "Gray Color", "Black & White Color", "True Color" });
            this.pixelTypeComboBox.SelectedIndex = 0;

            this.cbobpp.Items.AddRange(new string[] { "To 1bpp", "To 4bpp", "To 8bpp", "To 8bpp Grayscale", "To 16bpp RGB555", "To 16bpp RGB565", "To 24bpp", "To 32bpp", "To 32bpp ARGB" });
            this.cbobpp.SelectedIndex = 0;

            this.cbobufferbpp.Items.AddRange(new string[] { "To 1bpp", "To 4bpp", "To 8bpp", "To 8bpp Grayscale", "To 16bpp RGB555", "To 16bpp RGB565", "To 24bpp", "To 32bpp", "To 32bpp ARGB" });
            this.cbobufferbpp.SelectedIndex = 0;

            


            this.clrHashForeColor =this.Color2Uint32( Color.Black);
            this.clrHashBackColor = this.Color2Uint32(Color.White);

            short iCount = this.axScanner1.GetNumImageSources();

            for (short i = 0; i < iCount; i++)
            {
                this.imageSourceComboBox.Items.Add(this.axScanner1.GetImageSourceName(i));
            }

            for (short j = 8; j <= 60; j += 2)
            {
                this.fontSizeComboBox.Items.Add(j);
            }
            this.fontSizeComboBox.SelectedIndex = 2;

            this.fontComboBox.Items.AddRange(new string[] { "Arial", "Arial Black", "Comic Sans MS", "Times New Roman" });
            this.fontComboBox.SelectedIndex = 0;

            this.fontStyleComboBox.Items.AddRange(new string[] {"Regular","Bold","Italic","BoldItalic","Underline"  });
            this.fontStyleComboBox.SelectedIndex = 0;

            this.textStyleComboBox.Items.AddRange(new string[] { "Normal", "Outline", "Filled Outline", "Hash Brush", "Texture Brush" });
            this.textStyleComboBox.SelectedIndex = 0;

            this.hashBrushComboBox.Items.AddRange(new string[]{"HatchStyleHorizontal ","HatchStyleVertical ","HatchStyleForwardDiagonal ", "HatchStyleBackwardDiagonal "
, "HatchStyleCross ", "HatchStyleDiagonalCross ","HatchStyle05Percent ", "HatchStyle10Percent ","HatchStyle20Percent ", "HatchStyle25Percent ", "HatchStyle30Percent "
, "HatchStyle40Percent ", "HatchStyle50Percent ", "HatchStyle60Percent ","HatchStyle70Percent ","HatchStyle75Percent ","HatchStyle80Percent ", "HatchStyle90Percent "
, "HatchStyleLightDownwardDiagonal ", "HatchStyleLightUpwardDiagonal ","HatchStyleDarkDownwardDiagonal ", "HatchStyleDarkUpwardDiagonal ","HatchStyleWideDownwardDiagonal "
,"HatchStyleWideUpwardDiagonal ", "HatchStyleLightVertical ", "HatchStyleLightHorizontal ", "HatchStyleNarrowVertical ","HatchStyleNarrowHorizontal "
,"HatchStyleDarkVertical ","HatchStyleDarkHorizontal ","HatchStyleDashedDownwardDiagonal ", "HatchStyleDashedUpwardDiagonal ", "HatchStyleDashedHorizontal "
, "HatchStyleDashedVertical ", "HatchStyleSmallConfetti ","HatchStyleLargeConfetti ", "HatchStyleZigZag ", "HatchStyleWave ", "HatchStyleDiagonalBrick "
,"HatchStyleHorizontalBrick ", "HatchStyleWeave ","HatchStylePlaid ", "HatchStyleDivot ", "HatchStyleDottedGrid ","HatchStyleDottedDiamond ", "HatchStyleShingle "
, "HatchStyleTrellis ", "HatchStyleSphere ","HatchStyleSmallGrid ","HatchStyleSmallCheckerBoard ", "HatchStyleLargeCheckerBoard ","HatchStyleOutlinedDiamond "
, "HatchStyleSolidDiamond "});
            this.hashBrushComboBox.SelectedIndex = 0;


            this.ApplyDefaultCap();

        }

        private void fontComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            this.axScanner1.TextFontName = this.fontComboBox.SelectedItem.ToString();
        }

        private void fontSizeComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            this.axScanner1.TextFontSize = (short)this.fontSizeComboBox.SelectedItem;
        }

        private void fontStyleComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            this.axScanner1.TextFontStyle = (short)this.fontStyleComboBox.SelectedIndex;
        }

        private void hashBrushComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            SetHashBrushValue();

        }

        private void SetHashBrushValue()
        {
            this.axScanner1.SetHashBrushValue((short)this.hashBrushComboBox.SelectedIndex, this.clrHashForeColor, this.clrHashBackColor);
        }

        private void textStyleComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            this.ApplyTextStyle();
        }

        private void ApplyTextStyle()
        {
            switch (this.textStyleComboBox.SelectedIndex)
            {
                case 0:
                    this.textColorButton.Enabled = true;
                    this.outlineBorderButton.Enabled = false;
                    this.outlineBackButton.Enabled = false;
                    this.hashBrushComboBox.Enabled = false;
                    this.hashForeColorButton.Enabled = false;
                    this.hashBackColorButton.Enabled = false;
                    this.textureImageButton.Enabled = false;
                    break;
                case 1:
                    this.textColorButton.Enabled = false;
                    this.outlineBorderButton.Enabled = true;
                    this.outlineBackButton.Enabled = false;
                    this.hashBrushComboBox.Enabled = false;
                    this.hashForeColorButton.Enabled = false;
                    this.hashBackColorButton.Enabled = false;
                    this.textureImageButton.Enabled = false;
                    break;
                case 2:
                    this.textColorButton.Enabled = false;
                    this.outlineBorderButton.Enabled = true;
                    this.outlineBackButton.Enabled = true;
                    this.hashBrushComboBox.Enabled = false;
                    this.hashForeColorButton.Enabled = false;
                    this.hashBackColorButton.Enabled = false;
                    this.textureImageButton.Enabled = false;
                    break;
                case 3:
                    this.textColorButton.Enabled = false;
                    this.outlineBorderButton.Enabled = false;
                    this.outlineBackButton.Enabled = false;
                    this.hashBrushComboBox.Enabled = true;
                    this.hashForeColorButton.Enabled = true;
                    this.hashBackColorButton.Enabled = true;
                    this.textureImageButton.Enabled = false;
                    break;
                case 4:
                    this.textColorButton.Enabled = false;
                    this.outlineBorderButton.Enabled = false;
                    this.outlineBackButton.Enabled = false;
                    this.hashBrushComboBox.Enabled = false;
                    this.hashForeColorButton.Enabled = false;
                    this.hashBackColorButton.Enabled = false;
                    this.textureImageButton.Enabled = true;
                    if (this.textureImageTextBox.Text.Trim() == string.Empty)
                    {
                        MessageBox.Show("Please select texture brush");
                        return;
                    }

                    this.axScanner1.SetTextureBrushImage(this.textureImageTextBox.Text.Trim());

                    break;

            }

            this.axScanner1.TextStyle = (short)this.textStyleComboBox.SelectedIndex;

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

        private void showTextCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            this.axScanner1.ShowText = this.showTextCheckBox.Checked;
        }

        private void hashBackColorButton_Click(object sender, EventArgs e)
        {
            using (ColorDialog cd = new ColorDialog())
            {
                if (cd.ShowDialog(this) == DialogResult.OK)
                {
                    this.clrHashBackColor = this.Color2Uint32(cd.Color);
                    this.SetHashBrushValue();
                }
            }
        }

        private void hashForeColorButton_Click(object sender, EventArgs e)
        {
            using (ColorDialog cd = new ColorDialog())
            {
                if (cd.ShowDialog(this) == DialogResult.OK)
                {
                    this.clrHashForeColor = this.Color2Uint32(cd.Color);
                    this.SetHashBrushValue();
                }
            }
        }

        private void outlineBackButton_Click(object sender, EventArgs e)
        {
            using (ColorDialog cd = new ColorDialog())
            {
                if (cd.ShowDialog(this) == DialogResult.OK)
                {
                    this.axScanner1.SetOutlineTextBackColor(this.Color2Uint32(cd.Color));
                }
            }
        }

        private void outlineBorderButton_Click(object sender, EventArgs e)
        {
            using (ColorDialog cd = new ColorDialog())
            {
                if (cd.ShowDialog(this) == DialogResult.OK)
                {
                    this.axScanner1.SetOutlineTextBorderColor(this.Color2Uint32(cd.Color));
                }
            }
        }

        private void textColorButton_Click(object sender, EventArgs e)
        {
            using (ColorDialog cd = new ColorDialog())
            {
                if (cd.ShowDialog(this) == DialogResult.OK)
                {
                    this.axScanner1.TextColor=cd.Color;
                }
            }
        }

        private void textureImageButton_Click(object sender, EventArgs e)
        {
            if (this.openFileDialog1.ShowDialog(this) == DialogResult.OK)
            {
                this.textureImageTextBox.Text = this.openFileDialog1.FileName;
                this.axScanner1.SetTextureBrushImage(this.textureImageTextBox.Text.Trim());
            }
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
            this.axScanner1.View = 5;
            this.axScanner1.TIFCompression = (short)this.tiffCompressionComboBox.SelectedIndex;

            this.saveFileDialog1.Filter = "TIF File (*.tif)|*.tif";
            if (this.saveFileDialog1.ShowDialog(this) == DialogResult.OK)
            {

                bool a = this.axScanner1.SaveAllPage2TIF(saveFileDialog1.FileName, this.singleFileCheckBox.Checked, 1);
                if (a)
                    MessageBox.Show("save "+saveFileDialog1.FileName+ " complete");
                
            }
        }
        private void AddPDFAnnotation()
        {
            axScanner1.PDFInitAnnotation();

            axScanner1.PDFDrawFillRectangle(0, 10000, 50000, 200000, 20000, 1000, 255, 0, 0);


            long FontID = axScanner1.PDFAddFont("Arial", true, true);
            axScanner1.PDFEmbedFont((short)FontID);

            axScanner1.PDFSetTextColor(255, 255, 255);
            axScanner1.PDFDrawText(0, 10000, 40000, "Scanner Pro's PDF annotation features", (int)FontID, 20, 0);

                  // draw line from 2cm x 0.9 cm to 20cm x 0.9 cm
            axScanner1.PDFDrawLine(0, 20000, 9000, 200000, 9000, 1000, 0, 0, 0);

            long FontID2 = axScanner1.PDFAddFont("Times New Roman", true, true);
            axScanner1.PDFSetTextColor(0, 255, 0);

                // draw text at  1cm x 6cm and center alignment
            axScanner1.PDFDrawTextAlign(0, 10000, 60000, "www.viscomsoft.com", (int)FontID2, 18, 1);
            axScanner1.PDFSetTextColor(255, 0, 0);

             //draw text at specificed centimeters 3cm x 20cm
           axScanner1.PDFDrawText(0, 30000, 200000, "Rotated Text 1234567890", (int)FontID, 40, 45);


           axScanner1.PDFDrawRectangle(0, 30000, 150000, 50000, 20000, 1000, 255, 255, 0);

          axScanner1.PDFDrawRectangle(0, 20000, 140000, 30000, 20000, 1000, 0, 255, 0);

          axScanner1.PDFDrawRectangle(0, 50000, 120000, 50000, 20000, 1000, 255, 100, 0);

        }

        private void SaveToMultiPagePDF()
        {
            this.axScanner1.View = 5;
            this.axScanner1.TIFCompression = (short)this.tiffCompressionComboBox.SelectedIndex;

            PdfSecureSetting();

            if (chkaddPDFAnnotation.Checked)
                AddPDFAnnotation();

            if (chkpdfusejpegcomp2.Checked)
            {

                axScanner1.PDFUseJPEGCompression = true;
                axScanner1.PDFJPEGQuality = (short)this.numericjpegquality2.Value;
            }
            else
            {
                axScanner1.PDFUseJPEGCompression = false;
            }

            this.saveFileDialog1.Filter = "PDF File (*.pdf)|*.pdf";



            if (this.saveFileDialog1.ShowDialog(this) == DialogResult.OK)
            {

                bool a = this.axScanner1.SaveAllPage2PDF(saveFileDialog1.FileName, this.singleFileCheckBox.Checked, 1);
                if (a)
                    MessageBox.Show("save " + saveFileDialog1.FileName + " complete");

            }
        }

        private void saveMultiPagePDFButton_Click(object sender, EventArgs e)
        {
            SaveToMultiPagePDF();
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
          
            string strType = (string)this.outputTypeComboBox.SelectedItem;

            PdfSecureSetting();

            if (chkaddPDFAnnotation.Checked)
                AddPDFAnnotation();

            if (chkpdfusejpegcomp.Checked)
            {
                axScanner1.PDFUseJPEGCompression = true;
                axScanner1.PDFJPEGQuality = (short)this.numericjpegquality.Value;
            }
            else
            {
                axScanner1.PDFUseJPEGCompression = false;
            }

            this.saveFileDialog1.Filter = "";
            
            if (this.saveFileDialog1.ShowDialog(this) == DialogResult.OK)
            {
                string strTmpFile = saveFileDialog1.FileName;

                short a = this.axScanner1.SaveBySize(strTmpFile, strType, (short)this.outputWidth.Value, (short)this.outputHeight.Value);

                if (a == 1)
                    MessageBox.Show("Save" + strTmpFile + "." + strType + " complete");
                else
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

        private void gammaButton_Click(object sender, EventArgs e)
        {
            this.axScanner1.Gamma = (double)this.gamma.Value;
        }

        private void hueButton_Click(object sender, EventArgs e)
        {
            this.axScanner1.Hue = (double)this.hue.Value;
        }

        private void contrastButton_Click(object sender, EventArgs e)
        {
            this.axScanner1.Contrast = (double)this.contrast.Value;
        }

        private void brightnessButton_Click(object sender, EventArgs e)
        {
            this.axScanner1.Brightness = (double)this.brightness.Value;
        }

        private void saturationButton_Click(object sender, EventArgs e)
        {
            this.axScanner1.Saturation = (double)this.saturation.Value;
        }

        private void drawTextButton_Click(object sender, EventArgs e)
        {
            this.axScanner1.DrawText((short)this.textLeft.Value, (short)this.textTop.Value, this.textBox1.Text + Environment.NewLine + this.textBox2.Text);
        }

        private void scanButton_Click(object sender, EventArgs e)
        {
            if (this.chkpanning.Checked)
                this.axScanner1.EnablePanning = true;
            else
                this.axScanner1.EnablePanning = false;
         
            if (this.chkbufferresize.Checked)
            {
                this.axScanner1.BufferResizeMode = true;
                this.axScanner1.BufferResizeScale = Convert.ToDouble(this.txtbufferresizescale.Text);
                this.axScanner1.BufferResizeMaxWidth = Convert.ToInt16(this.txtbufferresizemaxwidth.Text);
            }
            else
                this.axScanner1.BufferResizeMode = false;

            
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
            string strType = (string)this.outputTypeComboBox.SelectedItem;

            axScanner1.TIFCompression = (short)cbocurrtifcomp.SelectedIndex;

            PdfSecureSetting();
            if (chkaddPDFAnnotation.Checked)
                AddPDFAnnotation();

            if (chkpdfusejpegcomp.Checked)
            {
                axScanner1.PDFUseJPEGCompression = true;
                axScanner1.PDFJPEGQuality = (short)this.numericjpegquality.Value;
            }
            else
            {
                axScanner1.PDFUseJPEGCompression = false;
            }

            this.saveFileDialog1.Filter = "";
            
         
              if (this.saveFileDialog1.ShowDialog(this) == DialogResult.OK)
              {
                  string strTmpFile = saveFileDialog1.FileName;


                  short result = this.axScanner1.Save(strTmpFile, strType);
                  if (result == 1)
                  {
                      MessageBox.Show(this, "Save " + strTmpFile + "." + strType + " complete");
                  }
                  else
                  {
                      MessageBox.Show(this, "Save fail");
                  }

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

        private void trackBar1_Scroll(object sender, EventArgs e)
        {
            axScanner1.RotateAt(trackBar1.Value);
        }

        private void axScanner1_EndScan_1(object sender, EventArgs e)
        {
            this.outputWidth.Value = (decimal)this.axScanner1.FileWidth;
            this.outputHeight.Value = (decimal)this.axScanner1.FileHeight;

           

            this.RefreshPageNo();
        }

        private void PdfSecureSetting()
        {
            this.axScanner1.PDFOutputPDFA = true;
            //'no password
            if (optpdfopt1.Checked)
                this.axScanner1.PDFSetPassword("", "");

            //Secure PDF with 40 bit RC4 and owner, user password, allow all permissions
            if (optpdfopt2.Checked)
            {
                this.axScanner1.PDFSetPassword("123", "123");
                this.axScanner1.PDFSetEncryption40bit(true, true, true, true);
            }

            //Secure PDF with 128 bit RC4 and owner, user password, allow all permissions
            if (optpdfopt3.Checked)
            {
                this.axScanner1.PDFSetPassword("123", "123");
                this.axScanner1.PDFSetEncryption128bit(true, true, true, true, true, true, true, true);
            }

            //Secure PDF with 40 bit RC4 ,allow all permissions
            if (optpdfopt4.Checked)
            {
                this.axScanner1.PDFSetPassword("", "");
                this.axScanner1.PDFSetEncryption40bit(true, true, true, true);
            }

            //Secure PDF with 128 bit RC4, allow all permissions
            if (optpdfopt5.Checked)
            {
                this.axScanner1.PDFSetPassword("", "");
                this.axScanner1.PDFSetEncryption128bit(true, true, true, true, true, true, true, true);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            int iIndex= this.cbobpp.SelectedIndex;

            if (iIndex == 0)
                axScanner1.ConvertTo1bpp();
            else if (iIndex == 1)
                axScanner1.ConvertTo4bpp();
            else if (iIndex == 2)
                axScanner1.ConvertTo8bpp();
            else if (iIndex == 3)
                axScanner1.ConvertTo8bppGrayScale();
            else if (iIndex == 4)
                axScanner1.ConvertTo16bppRGB555();
            else if (iIndex == 5)
                axScanner1.ConvertTo16bppRGB565();
            else if (iIndex == 6)
                axScanner1.ConvertTo24bpp();
            else if (iIndex == 7)
                axScanner1.ConvertTo32bpp();
            else if (iIndex == 8)
                axScanner1.ConvertTo32bppARGB();

        }

        private void outputTypeComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            string strType = (string)this.outputTypeComboBox.SelectedItem;

            if (strType == "TIF")
                cbocurrtifcomp.Enabled = true;
            else
                cbocurrtifcomp.Enabled = false;
         
        }

        private void button2_Click(object sender, EventArgs e)
        {
            int iIndex = this.cbobufferbpp.SelectedIndex;

            if (iIndex == 0)
                axScanner1.ConvertTo1bpp();
            else if (iIndex == 1)
                axScanner1.ConvertTo4bpp();
            else if (iIndex == 2)
                axScanner1.ConvertTo8bpp();
            else if (iIndex == 3)
                axScanner1.ConvertTo8bppGrayScale();
            else if (iIndex == 4)
                axScanner1.ConvertTo16bppRGB555();
            else if (iIndex == 5)
                axScanner1.ConvertTo16bppRGB565();
            else if (iIndex == 6)
                axScanner1.ConvertTo24bpp();
            else if (iIndex == 7)
                axScanner1.ConvertTo32bpp();
            else if (iIndex == 8)
                axScanner1.ConvertTo32bppARGB();

        }

        private void button3_Click(object sender, EventArgs e)
        {

            if (this.axScanner1.BlankPageIsBlank((double)numericconf.Value))
                MessageBox.Show("This page is blank page");
            else
                MessageBox.Show("This page is not a blank page");

        }

        private void button4_Click(object sender, EventArgs e)
        {
            numericconf.Value =(decimal) axScanner1.BlankPageGetConfidence();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            this.axScanner1.View = 5;

            if(singleFileCheckBox.Checked)
            {
            axScanner1.DocxAddChar(0, "T", 0, 0, "Arial", 10, 255, 255, 255, 255, 0, 0, false, false, false);
            axScanner1.DocxAddChar(0, "h", 1, 0, "Arial", 10, 255, 255, 255, 255, 0, 0, false, false, false);
            axScanner1.DocxAddChar(0, "i", 2, 0, "Arial", 10, 255, 255, 255, 255, 0, 0, false, false, false);
            axScanner1.DocxAddChar(0, "s", 3, 0, "Arial", 10, 255, 255, 255, 255, 0, 0, false, false, false);
            axScanner1.DocxAddChar(0, " ", 4, 0, "Arial", 10, 255, 255, 255, 255, 0, 0, false, false, false);
            axScanner1.DocxAddChar(0, "i", 5, 0, "Arial", 10, 255, 255, 255, 255, 0, 0, false, false, false);
            axScanner1.DocxAddChar(0, "s", 6, 0, "Arial", 10, 255, 255, 255, 255, 0, 0, false, false, false);
            axScanner1.DocxAddChar(0, " ", 7, 0, "Arial", 10, 255, 255, 255, 255, 0, 0, false, false, false);
            axScanner1.DocxAddChar(0, "t", 8, 0, "Arial", 10, 255, 255, 255, 255, 0, 0, false, false, false);
            axScanner1.DocxAddChar(0, "e", 9, 0, "Arial", 10, 255, 255, 255, 255, 0, 0, false, false, false);
            axScanner1.DocxAddChar(0, "x", 10, 0, "Arial", 10, 255, 255, 255, 255, 0, 0, false, false, false);
            axScanner1.DocxAddChar(0, "t", 11, 0, "Arial", 10, 255, 255, 255, 255, 0, 0, false, false, false);
            axScanner1.DocxAddChar(0, "1", 12, 0, "Arial", 10, 255, 255, 255, 255, 0, 0, false, false, false);
            }

            this.saveFileDialog1.Filter = "Ms Word Docx File (*.docx)|*.docx";
            if (this.saveFileDialog1.ShowDialog(this) == DialogResult.OK)
            {

                bool a = this.axScanner1.SaveAllPage2Docx(saveFileDialog1.FileName, this.singleFileCheckBox.Checked, 1);
                if (a)
                    MessageBox.Show("save " + saveFileDialog1.FileName + " complete");

            }

        }

        private void axScanner1_EndAllScan(object sender, EventArgs e)
        {
            if (chkautosaveallpages.Checked)
            {
                SaveToMultiPagePDF();
            }
        }

        private void axScanner1_ScanningError(object sender, EventArgs e)
        {
            this.listBox1.Items.Add("Error occur, may be Paper Jam");
        }
     


    }
}