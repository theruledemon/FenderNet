using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Net;
using System.IO;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;
using System.Globalization;
using OfficeOpenXml;
using System.Reflection;

namespace FenderNet
{
    public partial class Form1 : Form
    {

        String fileName = "";
        String[] priceResults = new String[10];
        String[] resultsName = new String[10];
        String[] excelOutputRow = new String[16];
        List<String[]> excelRows = new List<String[]>();
        

        public Form1()
        {
            InitializeComponent();
        }
        /*
         Look for the intended file        
        */
        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            if (openFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                fileName = openFileDialog.FileName;
                FileTextBox.Text = fileName;
            }


        }
        /*
         Calculate avg prices and save them to new excel. 
        */
        private void button2_Click(object sender, EventArgs e)
        {
            OutputTXT.Text = "";
            int from = Convert.ToInt32(FromTextBox.Text);
            int to = Convert.ToInt32(ToTextBox.Text);
            float avgPrice = 0;
            int resultPlace = 0;
            double priceToCompare = 0;
            string lookFor = "";
            int sheet = 1;

            _Application excel = new _Excel.Application();
            Workbook wb = excel.Workbooks.Open(fileName);
            Worksheet ws = wb.Worksheets[sheet];
            int row = from;
            int count = 0;
            int delete = 0;
            
            while (ws.Cells[row, 1].Value2 != null && delete < to)
            {
                // look from source file for prices to look for
                delete++;
                lookFor = ws.Cells[row, 5].Value2;
                priceToCompare = ws.Cells[row, 14].Value2;
                if (ws.Cells[row, 5].Value2 == null)
                {
                    count = 0;
                    row++;
                    continue;
                }
                // end


                string url = "https://www.ebay.com/sch/i.html?_ftrt=901&_sop=12&_sadis=15&_dmd=1&_osacat=0&_ipg=25&_ftrv=1&_from=R40&_trksid=m570.l1313&_nkw=" + lookFor + "&_sacat=0&LH_TitleDesc=1";
                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);
                HttpWebResponse response = (HttpWebResponse)request.GetResponse();
                StreamReader streamreaderEbay = new StreamReader(response.GetResponseStream());
                String sourceCodeEbay = streamreaderEbay.ReadToEnd();
                streamreaderEbay.Close();
                string patternForPrice = "s-item__price";
                string sentence = sourceCodeEbay;

                // look into the web html and find the prices and save them to array
                foreach (Match match in Regex.Matches(sentence, patternForPrice))
                {
                    string foundOnePrice = sourceCodeEbay.Substring(match.Index + 16, 10);
                    string price = foundOnePrice.Substring(0, foundOnePrice.IndexOf('<')); // result
                    count++;
                    priceResults[resultPlace] = price;
                    if (count >= 10)
                    {
                        break;
                    }

                    resultPlace++;
                }
                // calculate avg price
                foreach (String result in priceResults)
                {
                    avgPrice += float.Parse(result, CultureInfo.InvariantCulture.NumberFormat);
                }
                avgPrice = (float)Math.Round(avgPrice / 10); // round up the avg price
                // create result excel row
                excelOutputRow = new String[]{ (ws.Cells[row, 1].Value2).ToString("G") ,  lookFor, (priceToCompare).ToString("G"), avgPrice.ToString("G"),
                    Math.Round((avgPrice - priceToCompare), 2).ToString("G"), priceResults[0], priceResults[1], priceResults[2], priceResults[3], priceResults[4], priceResults[5], priceResults[6],
                    priceResults[7], priceResults[8], priceResults[9], url};
                excelRows.Add(excelOutputRow);
                OutputTXT.Text += "Tehtud " + row + " rida\n";
                // every ten rows done refresh outup log
                if (row % 10 == 0) 
                {
                    OutputTXT.Text = "";
                }
                count = 0;
                row++;
                resultPlace = 0;
                avgPrice = 0;


            }
            wb.Close();
            // create new excel file for results
            using (ExcelPackage excelCreate = new ExcelPackage())
            {
                excelCreate.Workbook.Worksheets.Add("Tulemused");

                var headerRow = new List<string[]>()
                    {
                    new string[] { "ID", "Tootja kood", "Hind", "Keskmine hind", "Hindade vahe",
                        "Leitud hind 1", "Leitud hind 2", "Leitud hind 3", "Leitud hind 4", "Leitud hind 5", "Leitud hind 6", "Leitud hind 7", "Leitud hind 8", "Leitud hind 9", "Leitud hind 10", "URL" }
                    };


                // Determine the header range (e.g. A1:D1)
                string headerRange = "A1:" + Char.ConvertFromUtf32(headerRow[0].Length + 64) + "1";


                // Target a worksheet
                var worksheet = excelCreate.Workbook.Worksheets["Tulemused"];

                // Popular header row data
                worksheet.Cells[headerRange].LoadFromArrays(headerRow);
                worksheet.Cells[headerRange].AutoFitColumns();

                int rowToWrite = 2;
                int columnToWrite = 1;
                foreach (String[] write in excelRows)
                {
                    foreach (String rowValue in write)
                    {
                        worksheet.Cells[rowToWrite, columnToWrite].Value = rowValue;
                        columnToWrite++;
                    }
                    columnToWrite = 1;
                    rowToWrite++;

                }


                FileInfo excelFile = new FileInfo(fileName.Substring(0, fileName.LastIndexOf(@"\")) + @"\Result.xls");
                excelCreate.SaveAs(excelFile);
                System.Diagnostics.Process.Start(fileName.Substring(0, fileName.LastIndexOf(@"\")) + @"\Result.xls");
            }
            OutputTXT.Text = "Keskmised hinnad leitud";
        }
        /*
         Save modified prices to a new file.             
        */
        private void button3_Click(object sender, EventArgs e)
        {
            String resultFile = (FileTextBox.Text).Substring(0, fileName.LastIndexOf(@"\")) + @"\Result.xls";
            int sheet = 1;
            int row = 2;
            List<string> priceToReplace = new List<string>();
            List<string> ids = new List<string>();
            _Excel.Application excel = new _Excel.Application();
            Workbook wb = excel.Workbooks.Open(resultFile);
            Worksheet ws = wb.Worksheets[sheet];
            while (ws.Cells[row, 1].Value2 != null)
            {
                priceToReplace.Add("" + ws.Cells[row, 4].Value2);
                ids.Add(ws.Cells[row, 1].Value2);
                row++;
            }
            
            wb.Close();
            wb = excel.Workbooks.Open(fileName);
            ws = wb.Worksheets[sheet];
            int replaceCount = ids.Count();
            row = 2;
            
            while(replaceCount > 0)
            {
                OutputTXT.Text = "";
                for (int i = 0; i < ids.Count; i++)
                {
                    if (Convert.ToInt32(ids[i]) == Convert.ToInt32(ws.Cells[row, 1].Value2))
                    {
                        ws.Cells[row, 14].Value2 = priceToReplace[i];
                        //ws.Cells[row, 13].Value2 = (Convert.ToDouble(priceToReplace[i]) * 0.8);
                        ids.RemoveAt(i);
                        priceToReplace.RemoveAt(i);
                        OutputTXT.Text += "Teha veel ridu: " + replaceCount;
                        replaceCount--;
                        break;
                    }
                }
                row++;
                
            }
            wb.SaveAs((FileTextBox.Text).Substring(0, fileName.LastIndexOf(@"\")) + @"\Replaced.xls", _Excel.XlFileFormat.xlOpenXMLWorkbook, Missing.Value,
    Missing.Value, false, false, _Excel.XlSaveAsAccessMode.xlNoChange,
    _Excel.XlSaveConflictResolution.xlUserResolution, true,
    Missing.Value, Missing.Value, Missing.Value);
            wb.Close();
            OutputTXT.Text = "Salvestatud";
            
        }
    }
}
