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
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Style;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        private ExcelPackage excel = new ExcelPackage();
     // Create new workbook
        private ExcelWorksheet excelAnnuity;
     //Add cover page image
        private static string path = Path.Combine(Environment.CurrentDirectory, "cover.png");
        private Image logo = Image.FromFile(path);


        public Form1()
        {
            InitializeComponent();
         //Set file properties
            excel.Workbook.Properties.Author = "David Ivashenko";
            excel.Workbook.Properties.Title = "Annuity Projection Esimation";
            excel.Workbook.Properties.Subject = "Annuties";
            excel.Workbook.Properties.Created = DateTime.Now;

        //Build cover
            excelAnnuity = excel.Workbook.Worksheets.Add("Cover Page");
            //var picture = excelAnnuity.Drawings.AddPicture("Logo", logo);
            //picture.SetPosition(0, 0, 0, 0);

        }

        private void Submit_Click(object sender, EventArgs e)
        {
            int excelWorksheetcounter = excelAnnuity.Workbook.Worksheets.Count;

                if (excelWorksheetcounter > 1)
                    {
                        button1_Click(sender,e);
                    }

            excelAnnuity = excel.Workbook.Worksheets["Cover Page"];
            excelAnnuity.Cells["A2"].Value = firstname.Text;
            excelAnnuity.Cells["A3"].Value = lastname.Text;
            excelAnnuity.Cells["A1"].Value = "Spreadsheet Annuity Book For";
            excelAnnuity.Cells["A4"].Value = "Generated on : " + DateTime.Now.ToString("MM/dd/yyyy");
            excelAnnuity.Cells["A1:A4"].Style.Font.Bold = true;
            excelAnnuity.Cells["A1:A4"].Style.Font.Size = 25;
            excelAnnuity.Cells["A1:A4"].Style.Font.Name = "Arial";
            FileInfo excelFile = new FileInfo(@"C:\Users\Unreal1\Desktop\Annuity.xlsx");
            excelAnnuity.PrinterSettings.Orientation = eOrientation.Landscape;
            excelAnnuity.PrinterSettings.FitToPage = true;
            excelAnnuity.PrinterSettings.VerticalCentered = true;
            excelAnnuity.PrinterSettings.HorizontalCentered = true;
            //excelAnnuity.PrinterSettings.TopMargin = (decimal).5 / 2.54M; // narrow border
            //excelAnnuity.PrinterSettings.RightMargin = (decimal).5 / 2.54M; //narrow border
            excel.SaveAs(excelFile);
            Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Calculate Age
                var today = DateTime.Today;
                int day = int.Parse(daybirth.Text);
                int month = int.Parse(monthbirth.Text);
                int year = int.Parse(yearbirth.Text);
                var a = (today.Year * 100 + today.Month) * 100 + today.Day;
                var b = (year * 100 + month) * 100 + day;
                var age = (a - b) / 10000;
                excelAnnuity.Cells["A4"].Value = "Age: " + age;
                
            //Set Workbook Settings
                int excelWorksheetcounter = excelAnnuity.Workbook.Worksheets.Count;
                excelAnnuity.Workbook.Worksheets.Add("Worksheet" + excelWorksheetcounter);
                excelAnnuity = excel.Workbook.Worksheets["Worksheet" + excelWorksheetcounter];
                excelAnnuity.Cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
                excelAnnuity.Cells.Style.Fill.BackgroundColor.SetColor(Color.White);
                //excelAnnuity.Cells.Style.Fill.PatternType = ExcelFillStyle.None;

            //Set Printer Settings
                excelAnnuity.PrinterSettings.Orientation = eOrientation.Landscape;
                excelAnnuity.PrinterSettings.FitToPage = true;
                excelAnnuity.PrinterSettings.VerticalCentered = true;
                excelAnnuity.PrinterSettings.HorizontalCentered = true;
                excelAnnuity.PrinterSettings.TopMargin = (decimal).25; // narrow border
                excelAnnuity.PrinterSettings.RightMargin = (decimal).25; //narrow border
                excelAnnuity.PrinterSettings.LeftMargin = (decimal).25; //narrow border
                excelAnnuity.PrinterSettings.BottomMargin = (decimal).25; //narrow border


            //Add Account information
                excelAnnuity.Cells["A1"].Value = firstname.Text;
                excelAnnuity.Cells["A2"].Value = lastname.Text;
                excelAnnuity.Cells["A3"].Value = monthbirth.Text + "-" + daybirth.Text + "-" + yearbirth.Text;
                excelAnnuity.Cells["C1"].Value = "Account " + accountype.SelectedItem + " : 478-" + accountnumber.Text;
                
            excelAnnuity.Cells["A5"].Value = "Age";
            excelAnnuity.Cells["A5"].Style.Font.Bold = true;
            excelAnnuity.Cells["A5"].Style.Font.UnderLine = true;
            //excelAnnuity.Cells["A5"].Style.Fill.BackgroundColor.SetColor(Color.DarkBlue);
            for (int i = 6; i <= 36; i++)
            {

                excelAnnuity.Cells[i, 1].Value = age;
                age++;

            }


            excelAnnuity.Cells["B5"].Value = "Year";
            excelAnnuity.Cells["B5"].Style.Font.Bold = true;
            excelAnnuity.Cells["B5"].Style.Font.UnderLine = true;
            //excelAnnuity.Cells["B5"].Style.Fill.BackgroundColor.SetColor(Color.DarkBlue);
            int todaysyear = DateTime.Today.Year;
            for (int i = 6; i <= 36; i++)
            {

                excelAnnuity.Cells[i, 2].Value = todaysyear;
                todaysyear++;

            }

            excelAnnuity.Cells["C5"].Value = "Rider";
            excelAnnuity.Cells["C5"].Style.Font.Bold = true;
            excelAnnuity.Cells["C5"].Style.Font.UnderLine = true;
            //excelAnnuity.Cells["C5"].Style.Fill.BackgroundColor.SetColor(Color.DarkBlue);

            excelAnnuity.Cells["D5"].Value = "Income Benefit";
            excelAnnuity.Cells["D5"].Style.Font.Bold = true;
            excelAnnuity.Cells["D5"].Style.Font.UnderLine = true;


            if (radioButton1.Checked)
                {
                excelAnnuity.Cells["C2"].Value = "Type of Rider: GMIB";
                excelAnnuity.Cells["C3"].Value = "Contract Value: $"+contractvalue.Text;
                excelAnnuity.Cells["C4:C36"].Style.Numberformat.Format = "$###,###,###,##0.00";
                excelAnnuity.Cells["D4:D36"].Style.Numberformat.Format = "$###,###,###,##0.00";
                //excelAnnuity.Cells["C4:C36"].Style.Fill.BackgroundColor.SetColor(Color.CadetBlue);
                excelAnnuity.Column(3).Width = 17.2;
                excelAnnuity.Column(4).Width = 15.2;
                excelAnnuity.Column(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                excelAnnuity.Column(2).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                excelAnnuity.Column(3).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                excelAnnuity.Column(4).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                decimal rideramount = decimal.Parse(rider.Text);

                decimal dd = (decimal.Parse(rate.Text))/100;
                dd = dd * (decimal.Parse(premium.Text));

                //decimal incomebenefit;

                excelAnnuity.Cells["C4"].Value = "Initial Premium: $ " + premium.Text;
                excelAnnuity.Cells["C4"].Style.Numberformat.Format = "$###,###,###,##0.00";

                for (int i = 6; i <= 36; i++)
                {

                    excelAnnuity.Cells[i, 3].Value = rideramount;
                    excelAnnuity.Cells[i, 4].Value = dd;

                    rideramount = rideramount + dd;

                }




            }

                if (radioButton2.Checked)
                {
                excelAnnuity.Cells["D2"].Value = "Contract # " + contractnumber.Text;

                excelAnnuity.Cells["C2"].Value = "Type of Rider: GWLB";
                excelAnnuity.Cells["C3"].Value = "Contract Value: $" + contractvalue.Text;
                excelAnnuity.Cells["C4:C36"].Style.Numberformat.Format = "$###,###,###,##0.00";
                excelAnnuity.Cells["D4:D36"].Style.Numberformat.Format = "$###,###,###,##0.00";
                //excelAnnuity.Cells["C4:C36"].Style.Fill.BackgroundColor.SetColor(Color.CadetBlue);
                excelAnnuity.Column(3).Width = 17.2;
                excelAnnuity.Column(4).Width = 17.2;
                excelAnnuity.Column(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                excelAnnuity.Column(2).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                excelAnnuity.Column(3).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                excelAnnuity.Column(4).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                decimal rideramount = decimal.Parse(rider.Text);

                decimal dd = (decimal.Parse(rate.Text)) / 100;

                //decimal incomebenefit;

                excelAnnuity.Cells["C4"].Value = "Initial Premium: $ " + premium.Text;
                excelAnnuity.Cells["C4"].Style.Numberformat.Format = "$###,###,###,##0.00";

                for (int i = 6; i <= 36; i++)
                {

                    excelAnnuity.Cells[i, 3].Value = rideramount;
                    excelAnnuity.Cells[i, 4].Value = rideramount * dd;

                    rideramount = rideramount + (rideramount * dd);

                }
            }

            ExcelChart chart = excelAnnuity.Drawings.AddChart("FindingsChart", eChartType.ColumnClustered);
            chart.Title.Text = "Annuity Rider Projection with Income Benefit";
            chart.SetPosition(4, 0, 4, 0);
            chart.SetSize(800, 300);
            var ser1 = chart.Series.Add(excelAnnuity.Cells["C6:C36"], excelAnnuity.Cells["A6:A36"]);
            var ser2 = chart.Series.Add(excelAnnuity.Cells["D6:D36"], excelAnnuity.Cells["A6:A36"]);
            ser1.Header = "Rider Value";
            ser2.Header = "Income Benefit";


            ExcelChart chart2 = excelAnnuity.Drawings.AddChart("Income Chart", eChartType.ColumnClustered);
            chart2.Title.Text = "Income Benefit Projection";
            chart2.SetPosition(21, 0, 4, 0);
            chart2.SetSize(800, 300);
            var ser3 = chart2.Series.Add(excelAnnuity.Cells["D6:D36"], excelAnnuity.Cells["A6:A36"]);
            ser3.Header = "Income Benfit Value";
            excelAnnuity.Cells["A37:Q40"].Merge = true;
            excelAnnuity.Cells["A37"].Value = "Your contract with " + companylist.SelectedItem + " compounds annually at a rate of " + rate.Text +
            "% you may begin withdrawing your annuity for the value of the income benefit at anytime."    ;
            excelAnnuity.Cells["A37"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            excelAnnuity.Cells["A37"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            excelAnnuity.Cells["A37:Q40"].Style.WrapText = true;

            
            excelAnnuity.Cells["A5:D5"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            excelAnnuity.Cells["A5:D5"].Style.Fill.BackgroundColor.SetColor(Color.LightSteelBlue);

            excelAnnuity.Cells["A6:A36"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            excelAnnuity.Cells["A6:D36"].Style.Fill.BackgroundColor.SetColor(Color.LightGray);

        }

        private void rate_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
