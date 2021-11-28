using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {
        string strApp;

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
             int iFindIndex;
            strApp = Application.ExecutablePath;
            iFindIndex = strApp.IndexOf("Examples");
            if (iFindIndex != -1)
            {
                strApp = strApp.Substring(0, iFindIndex);

                strApp = strApp + "barcodeimage";
            }

        }

        private void BarcodeEngineSetting()
        {
         axScanner1.BarCodeReadScanMultiple = chkScanMultiple.Checked;
        axScanner1.BarCodeReadScanWithoutRotation = chkscannorotation.Checked;
        axScanner1.BarCodeReadScan45Rotation = chkScan45Rotation.Checked;
        axScanner1.BarCodeReadScan45CouterRotation = chkScan45CounterRotation.Checked;
        axScanner1.BarCodeReadScan90Rotation = chkScan90Rotation.Checked;
        axScanner1.BarCodeReadCustomRotation = Double.Parse(txtcustomrotation.Text);
        axScanner1.BarCodeReadScanAccuracy = chkscanmore.Checked;
        }
        private void Button3_Click(object sender, EventArgs e)
        {
            int ibarcodeCount;

            chkscanmore.Checked = true;
            chkScanMultiple.Checked = false;
            chkScan90Rotation.Checked = false;
            chkScan45Rotation.Checked = false;
            chkScan45CounterRotation.Checked = false;
            chkscannorotation.Checked = true;
            

           axScanner1.LoadImage(strApp + "\\barcodetest2.jpg");

            BarcodeEngineSetting();
            ibarcodeCount = axScanner1.BarCodeReadFullPage();

          DisplayBarCode(ibarcodeCount);

        }

        private void DisplayBarCode(int ibarcodeCount)
        {
              string str1;
              string strTmp;
                int i;

           if( ibarcodeCount < 1)
           {

                MessageBox.Show("No barcode found, make sure barcode setting selected correctly");
                return;
           }
          
        str1 = "Total " + ibarcodeCount.ToString()  + " BarCode detected" + "\r\n";

        for (i = 0; i < ibarcodeCount; i++)
        {
               strTmp = axScanner1.BarCodeGetType((short)i) + " score:" + axScanner1.BarCodeGetScore((short)i).ToString() + " value:" + axScanner1.BarCodeGetValue((short)i);

               str1 = str1 + "\r\n" + strTmp;

    
        }

        MessageBox.Show(str1);

            /*
      For i = 0 To ibarcodeCount - 1

          strTmp = Me.Scanner1.BarCodeGetType(i) + " score:" + Str(Me.Scanner1.BarCodeGetScore(i)) + " value:" + Me.Scanner1.BarCodeGetValue(i)

          str1 = str1 + Chr(13) + Chr(10) + strTmp
      Next

      MessageBox.Show(str1)
           **/
        }

        private void Button4_Click(object sender, EventArgs e)
        {
         int ibarcodeCount;
         chkscanmore.Checked = true;
         chkScanMultiple.Checked = true;
         chkScan90Rotation.Checked = false;
         chkScan45Rotation.Checked = false;
         chkScan45CounterRotation.Checked = false;
         chkscannorotation.Checked = true;
            

            axScanner1.LoadImage(strApp + "\\barcodetest3.jpg");
            BarcodeEngineSetting();
            ibarcodeCount = axScanner1.BarCodeReadFullPage();

            DisplayBarCode(ibarcodeCount);
        }

        private void button8_Click(object sender, EventArgs e)
        {
            int ibarcodeCount;

            chkscanmore.Checked = true;
            chkScanMultiple.Checked=false;
            chkScan90Rotation.Checked = false;
            chkScan45Rotation.Checked = false;
            chkScan45CounterRotation.Checked = false;
            chkscannorotation.Checked = true;
            
            axScanner1.LoadImage(strApp + "\\barcodetest1.jpg");

            BarcodeEngineSetting();
            ibarcodeCount = axScanner1.BarCodeReadFullPage();

            DisplayBarCode(ibarcodeCount);
        }

        private void Button5_Click(object sender, EventArgs e)
        {
            int ibarcodeCount;

            chkscanmore.Checked = true;
            chkScanMultiple.Checked = false;
            chkScan90Rotation.Checked = false;
            chkScan45Rotation.Checked = false;
            chkScan45CounterRotation.Checked = false;
            chkscannorotation.Checked = true;

            axScanner1.LoadImage(strApp + "\\barcodetest4.png");

            BarcodeEngineSetting();
            ibarcodeCount = axScanner1.BarCodeReadFullPage();

            DisplayBarCode(ibarcodeCount);
        }

        private void Button6_Click(object sender, EventArgs e)
        {
            int ibarcodeCount;

            chkscanmore.Checked = true;
            chkScanMultiple.Checked = false;
            chkScan90Rotation.Checked =true;
            chkScan45Rotation.Checked = true;
            chkScan45CounterRotation.Checked = false;
            chkscannorotation.Checked = true;

            axScanner1.LoadImage(strApp + "\\barcodetest5.jpg");

            BarcodeEngineSetting();
            ibarcodeCount = axScanner1.BarCodeReadFullPage();

            DisplayBarCode(ibarcodeCount);
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            int ibarcodeCount;

            
            axScanner1.LoadImage(strApp + "\\barcodetest6.pdf");

            AdvBarcodeReader_OpenPDF(strApp + "\\barcodetest6.pdf");
      
        }

        private void AdvBarcodeReader_OpenPDF(string strFileName)
        {
              int iPDFPageCount;
              int i;

          cbopdfpage.Items.Clear();
          cbopdfimagecount.Items.Clear();

        iPDFPageCount = axScanner1.BarCodeReadPDFOpen(strFileName);

        cbopdfpage.Items.Clear();

        for(i = 1 ; i <= iPDFPageCount;i++)
        {
            cbopdfpage.Items.Add("Page " + i.ToString());
        }

        if(cbopdfpage.Items.Count > 0)
            cbopdfpage.SelectedIndex = 0;

        chkScanMultiple.Checked = false;
        chkScan45CounterRotation.Checked =false;
        chkScan90Rotation.Checked = false;
        chkScan45Rotation.Checked = false;
        }

        private void cbopdfpage_SelectedIndexChanged(object sender, EventArgs e)
        {
           short iImageCountPerPage;
             int i;
        cbopdfimagecount.Items.Clear();

        //first page is 1
        iImageCountPerPage = axScanner1.BarCodeReadPDFImageCount((short)(cbopdfpage.SelectedIndex + 1));
     
            for( i = 1; i<=iImageCountPerPage;i++)
        {
            cbopdfimagecount.Items.Add("Image " + i.ToString());
        }

        if (cbopdfimagecount.Items.Count > 0)
            cbopdfimagecount.SelectedIndex = 0;

       

        }

        private void Button7_Click(object sender, EventArgs e)
        {
            short ibarcodeCount;
            BarcodeEngineSetting();
             ibarcodeCount = axScanner1.BarCodeReadPDFScan((short)(cbopdfpage.SelectedIndex + 1), (short)(cbopdfimagecount.SelectedIndex + 1));
             DisplayBarCode(ibarcodeCount);
        }

        private void Button1_Click_1(object sender, EventArgs e)
        {
            short ibarcodeCount;
        
         axScanner1.LoadImage(strApp + "\\barcodetest6.pdf");
         AdvBarcodeReader_OpenPDF(strApp + "\\barcodetest6.pdf");

        }
    }
}
