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
            
            //do not resize the image, it may affect the barcode quality
            this.axScanner1.BufferResizeMode = false;

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
            long ibarcodeCount;
            string str1;
            string strTmp;

          axScanner1.BarCodeReadScanMultiple = chkScanMultiple.Checked;
         axScanner1.BarCodeReadScanWithoutRotation = chkscannorotation.Checked;
         axScanner1.BarCodeReadScan45Rotation = chkScan45Rotation.Checked;
         axScanner1.BarCodeReadScan45CouterRotation = chkScan45CounterRotation.Checked;
         axScanner1.BarCodeReadScan90Rotation = chkScan90Rotation.Checked;
         axScanner1.BarCodeReadCustomRotation =Convert.ToDouble(txtcustomrotation.Text);
         axScanner1.BarCodeReadScanAccuracy = chkscanmore.Checked;

            if(chkfullpage.Checked)
                 ibarcodeCount = axScanner1.BarCodeReadFullPage();
            else
               ibarcodeCount = axScanner1.BarCodeReadByZone(Convert.ToInt16(txtbarcodeleft.Text), Convert.ToInt16(txtbarcodetop.Text), Convert.ToInt16(txtbarcodewidth.Text), Convert.ToInt16(txtbarcodeheight.Text));

           if (ibarcodeCount < 1)
           {
               MessageBox.Show("No barcode found, try scan to black and white color and set the specific area for detect the barcode");
               return;
           }

            str1 = "Total " +ibarcodeCount.ToString()+ " BarCode detected" + "\r\n";

            for (int i = 0; i < ibarcodeCount - 1; i++)
                strTmp = axScanner1.BarCodeGetType((short)i).ToString() + " score:" + axScanner1.BarCodeGetScore((short)i).ToString() + " value:" + axScanner1.BarCodeGetValue((short)i).ToString();

            str1 = str1 + "\r\n";

               MessageBox.Show(str1);


        }

        private void chkfullpage_CheckedChanged(object sender, EventArgs e)
        {
           
        }

        private void axScanner1_EndScan_1(object sender, EventArgs e)
        {
            this.RefreshPageNo();

        }

        private void PdfSecureSetting()
        {
            
        }


    }
}