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

        void Enable_SelectFont()
        {
        }

        void Disable_SelectFont()
        {
       
        }

        private void Form1_Load(object sender, EventArgs e)
        {
             this.dPIComboBox.Items.AddRange(new string[] { "Onscreen Viewing 96dpi", "Fax 200dpi", "OCR Text 300dpi", "Laser Print Fine 600dpi" });
            this.dPIComboBox.SelectedIndex = 0;

            this.pixelTypeComboBox.Items.AddRange(new string[] { "Default", "Gray Color", "Black & White Color", "True Color" });
            this.pixelTypeComboBox.SelectedIndex = 0;

        
            short iCount = this.axScanner1.GetNumImageSources();

            for (short i = 0; i < iCount; i++)
            {
                this.imageSourceComboBox.Items.Add(this.axScanner1.GetImageSourceName(i));
            }

         

            this.ApplyDefaultCap();

        }

        private void fontComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            
        }

        private void fontSizeComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            
        }

        private void fontStyleComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            
        }

        private void hashBrushComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
           

        }

        private void SetHashBrushValue()
        {
         }

        private void textStyleComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            this.ApplyTextStyle();
        }

        private void ApplyTextStyle()
        {
           
        }

        private void greyScaleCheckBox_CheckedChanged(object sender, EventArgs e)
        {
           
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
           
        }

        private void hashBackColorButton_Click(object sender, EventArgs e)
        {
           
        }

        private void hashForeColorButton_Click(object sender, EventArgs e)
        {
           
        }

        private void outlineBackButton_Click(object sender, EventArgs e)
        {
         
        }

        private void outlineBorderButton_Click(object sender, EventArgs e)
        {
           
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
           
        }

        private void saveMultiPagePDFButton_Click(object sender, EventArgs e)
        {
            
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
            
        }

        private void saveToHBButton_Click(object sender, EventArgs e)
        {
          
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

        private void axScanner1_EndScan(object sender, EventArgs e)
        {
            
            this.RefreshPageNo();
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

        private void button1_Click(object sender, EventArgs e)
        {
            short ilangindex = 0;

            short iMode = 0;


            if (optlang1.Checked)
                ilangindex = 0;
            else if (optlang2.Checked)
                ilangindex = 1;
            else if (optlang3.Checked)
                ilangindex = 2;
            else if (optlang4.Checked)
                ilangindex = 3;
            else if (optlang5.Checked)
                ilangindex = 4;
            else if (optlang6.Checked)
                ilangindex = 5;
            else if (optlang7.Checked)
                ilangindex = 6;
            else if (optlang8.Checked)
                ilangindex = 7;
            else if (optlang9.Checked)
                ilangindex = 8;
            else if (optlang10.Checked)
                ilangindex = 9;
            else if (optlang11.Checked)
                ilangindex = 10;
            else if (optlang12.Checked)
                ilangindex = 11;
            else if (optlang13.Checked)
                ilangindex = 12;
            else if (optlang14.Checked)
                ilangindex = 13;
            else if (optlang15.Checked)
                ilangindex = 14;
            else if (optlang16.Checked)
                ilangindex = 15;
            else if (optlang17.Checked)
                ilangindex = 16;
            else if (optlang18.Checked)
                ilangindex = 17;
            else if (optlang19.Checked)
                ilangindex = 18;
            else if (optlang20.Checked)
                ilangindex = 19;
            else if (optlang21.Checked)
                ilangindex = 20;
            else if (optlang22.Checked)
                ilangindex = 21;
            else if (optlang23.Checked)
                ilangindex = 22;
            else if (optlang24.Checked)
                ilangindex = 23;
            else if (optlang25.Checked)
                ilangindex = 24;
            else if (optlang26.Checked)
                ilangindex = 25;
            else if (optlang27.Checked)
                ilangindex = 26;
            else if (optlang28.Checked)
                ilangindex = 27;
            else if (optlang29.Checked)
                ilangindex = 28;
            else if (optlang30.Checked)
                ilangindex = 29;
            else if (optlang31.Checked)
                ilangindex = 30;
            else if (optlang32.Checked)
                ilangindex = 31;
            else if (optlang33.Checked)
                ilangindex = 32;
            else if (optlang34.Checked)
                ilangindex = 33;
            else if (optlang35.Checked)
                ilangindex = 34;
            else if (optlang36.Checked)
                ilangindex = 35;
            else if (optlang37.Checked)
                ilangindex = 36;
            else if (optlang38.Checked)
                ilangindex = 37;
            else if (optlang39.Checked)
                ilangindex = 38;
            else if (optlang40.Checked)
                ilangindex = 39;
            else if (optlang41.Checked)
                ilangindex = 40;
            else if (optlang42.Checked)
                ilangindex = 41;


         
            
            if (optmode1.Checked)
                iMode = 0;

            if (optmode2.Checked)
                iMode = 1;

            if (optmode3.Checked)
                iMode = 2;

            if (optfilter1.Checked)
                axScanner1.OCRCharFilter = "";

            if (optfilter2.Checked)
                axScanner1.OCRCharFilter = "0123456789";

            if (optfilter3.Checked)
                axScanner1.OCRCharFilter = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

            if (optfilter4.Checked)
                axScanner1.OCRCharFilter = "abcdefghijklmnopqrstuvwxyz";

            axScanner1.OCRRecognizeMode =(short)iMode;

            if (chkfullpage.Checked)
                axScanner1.OCRSetRect(0, 0, 0, 0);
            else
            {
               
                axScanner1.OCRSetRect(Convert.ToInt16(txtocrleft.Text), Convert.ToInt16(txtocrtop.Text), Convert.ToInt16(txtocrwidth.Text), Convert.ToInt16(txtocrheight.Text));

            }       
        
          
            short iresult;
            

                iresult= axScanner1.OCRStartScan(ilangindex, "dictfiles");

                if (iresult == -3)
                {
                    MessageBox.Show("Please scan the image first");
                    return;
                }
                else
                {
                    if (iresult != 1)
                    {
                        MessageBox.Show("Start OCR Failed");
                        return;
                    }


                }


                button1.Enabled = false;
                button3.Enabled = true;


            

        }

        private void chkfullpage_CheckedChanged(object sender, EventArgs e)
        {
            if (chkfullpage.Checked)
            {
                this.txtocrleft.Enabled = false;
                this.txtocrtop.Enabled = false;
                this.txtocrwidth.Enabled = false;
                this.txtocrheight.Enabled = false;
            }
            else
            {
                this.txtocrleft.Enabled = true;
                this.txtocrtop.Enabled = true;
                this.txtocrwidth.Enabled = true;
                this.txtocrheight.Enabled = true;


            }
        }

        private void axScanner1_EndScan_1(object sender, EventArgs e)
        {
            this.RefreshPageNo();

        }

        private void optlang1_Click(object sender, EventArgs e)
        {
            Disable_SelectFont();
        }

        private void optlang2_Click(object sender, EventArgs e)
        {
            Enable_SelectFont();
        }

        private void optlang3_Click(object sender, EventArgs e)
        {
            Enable_SelectFont();
        }

        private void optlang4_Click(object sender, EventArgs e)
        {
            Enable_SelectFont();
        }

        private void optlang5_Click(object sender, EventArgs e)
        {
            Enable_SelectFont();
        }

        private void optlang6_Click(object sender, EventArgs e)
        {
            Enable_SelectFont();
        }

        private void optlang7_Click(object sender, EventArgs e)
        {
            Enable_SelectFont();
        }

        private void optlang8_Click(object sender, EventArgs e)
        {
            Enable_SelectFont();
        }

        private void optlang9_Click(object sender, EventArgs e)
        {
            Enable_SelectFont();
        }

        private void optlang10_Click(object sender, EventArgs e)
        {
            Enable_SelectFont();
        }

        private void optlang11_Click(object sender, EventArgs e)
        {
            Enable_SelectFont();
        }

        private void optlang12_Click(object sender, EventArgs e)
        {
            Enable_SelectFont();
        }

        private void optlang13_Click(object sender, EventArgs e)
        {
            Enable_SelectFont();
        }

        private void optlang14_Click(object sender, EventArgs e)
        {
            Enable_SelectFont();
        }

        private void optlang15_Click(object sender, EventArgs e)
        {
            Enable_SelectFont();
        }

        private void optlang16_Click(object sender, EventArgs e)
        {
            Enable_SelectFont();
        }

        private void optlang17_Click(object sender, EventArgs e)
        {
            Enable_SelectFont();
        }

        private void optlang18_Click(object sender, EventArgs e)
        {
            Enable_SelectFont();
        }

        private void optlang19_Click(object sender, EventArgs e)
        {
            Enable_SelectFont();
        }

        private void optlang20_Click(object sender, EventArgs e)
        {
            Enable_SelectFont();
        }

        private void optlang21_Click(object sender, EventArgs e)
        {
            Enable_SelectFont();
        }

        private void optlang22_Click(object sender, EventArgs e)
        {
            Enable_SelectFont();
        }

        private void optlang23_Click(object sender, EventArgs e)
        {
            Enable_SelectFont();
        }

        private void optlang24_Click(object sender, EventArgs e)
        {
            Enable_SelectFont();
        }

        private void optlang25_Click(object sender, EventArgs e)
        {
            Enable_SelectFont();
        }

        private void optlang26_Click(object sender, EventArgs e)
        {
            Enable_SelectFont();
        }

        private void optlang27_Click(object sender, EventArgs e)
        {
            Enable_SelectFont();
        }

        private void optlang28_Click(object sender, EventArgs e)
        {
            Enable_SelectFont();
        }

        private void optlang29_Click(object sender, EventArgs e)
        {
            Enable_SelectFont();
        }

        private void optlang30_Click(object sender, EventArgs e)
        {
            Enable_SelectFont();
        }

        private void optlang31_Click(object sender, EventArgs e)
        {
            Enable_SelectFont();
        }

        private void optlang32_Click(object sender, EventArgs e)
        {
            Enable_SelectFont();
        }

        private void optlang33_Click(object sender, EventArgs e)
        {
            Enable_SelectFont();
        }

        private void optlang34_Click(object sender, EventArgs e)
        {
            Enable_SelectFont();
        }

        private void optlang35_Click(object sender, EventArgs e)
        {
            Enable_SelectFont();
        }

        private void optlang36_Click(object sender, EventArgs e)
        {
            Enable_SelectFont();
        }

        private void optlang37_Click(object sender, EventArgs e)
        {
            Enable_SelectFont();
        }

        private void optlang38_Click(object sender, EventArgs e)
        {
            Enable_SelectFont();
        }

        private void optlang39_Click(object sender, EventArgs e)
        {
            Enable_SelectFont();
        }

        private void optlang40_Click(object sender, EventArgs e)
        {
            Enable_SelectFont();
        }

        private void Button3_Click(object sender, EventArgs e)
        {
           
        }

        private void Button4_Click(object sender, EventArgs e)
        {
          
        }

        private void axScanner1_OCRRecognized(object sender, AxSCANNERLib._DScannerEvents_OCRRecognizedEvent e)
        {
            if(e.bSuccess)
            {
                txttotalchar.Text= axScanner1.OCRGetRecognizedCharCount().ToString();

                if (optoutput1.Checked)
                    MessageBox.Show(axScanner1.OCRGetRecognizedText());
                else
                {
                            bool bresult;
                          this.saveFileDialog1.Filter = "Text File (*.txt)|*.txt";

                        if (this.saveFileDialog1.ShowDialog(this) == DialogResult.OK)
                        {
                            bresult = axScanner1.OCRSaveRecognizedText(saveFileDialog1.FileName);

                            if (bresult)
                                MessageBox.Show("Save Completed");
                            else
                                MessageBox.Show("Save Failed");
                          

                        }

                }

                button1.Enabled = true;
                button3.Enabled = false;



            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
          }

        private void button3_Click_1(object sender, EventArgs e)
        {
            axScanner1.OCRStopScan();

            button1.Enabled = true;
            button3.Enabled = false;


        }

     


    }
}