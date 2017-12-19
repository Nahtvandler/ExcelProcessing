using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {
        public static string XPath, FPath,fcount;
        //public static int fcount;
        
        public Form1()
        {
            InitializeComponent();
            XPath = null;
            FPath = null;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            
            //string[] list = new string[LastCell.Row];
            if (XPath == null || FPath == null)
            {
                MessageBox.Show("Не выбран путь для xlsl файла или директории с файлами");
                return;
            }

            //открытие excel файла и формирования массива занчений первой ячейки
            Excel.Application ObjWorkExcel = new Excel.Application();
            Excel.Workbook ObjWorkbook = ObjWorkExcel.Workbooks.Open(XPath, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkbook.Sheets[1];
            var LastCell = ObjWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);
            //
            int LastColumn = (int)LastCell.Column;
            int LastRow = (int)LastCell.Row;
            //
            string[] list = new string[LastCell.Row];
            string efname;
            label5.Text = "0 из " + (int)LastCell.Row; //кол-во строк
            //progressBar1.Maximum = (int)LastCell.Row; //
            //progressBar1.Minimum = 1; //
            for (int i = 0; i < (int)LastCell.Row; i++)
            {
                //собираем имя файла из ячеек
                efname = ObjWorkSheet.Cells[i + 1, 1].text.ToString()+"_"+ ObjWorkSheet.Cells[i + 1, 2].text.ToString()+"_"+
                    ObjWorkSheet.Cells[i + 1, 3].text.ToString()+"_"+ObjWorkSheet.Cells[i + 1, 4].text.ToString().Replace("/", "_I_");
                    //efname = efname.Replace("_I_", "/");
                list[i] = efname;
                //richTextBox1.Text = "Собранное имя: " + richTextBox1.Text + Environment.NewLine + list[i];
                //-собираем имя файла из ячеек

                //list[i] = ObjWorkSheet.Cells[i + 1, 4].text.ToString();
                //progressBar1.Value = i+1;
                //label1.Text = i.ToString(); //
            }
            
            //ObjWorkbook.Close(false, Type.Missing, Type.Missing);
            //ObjWorkExcel.Quit();
            //GC.Collect();
           
            //получение имен файлов
            string nstr,fname;
            int count=0,acount=0,flcount=0;
            Boolean chek=false;
            DirectoryInfo dirInfo = new DirectoryInfo(FPath);
            FileInfo[] files = dirInfo.GetFiles();
            for (int j = 0; j < LastRow; j++)
            //foreach (FileInfo file in files)
            {
                count = count + 1;
                //textBox3.Text = textBox3.Text + Environment.NewLine + "проверяем файл: "+file.Name; //
                foreach (FileInfo file in files)
                //for (int j = 0; j < LastRow; j++)
                {
                    nstr = file.Name;
                    //richTextBox1.Text = richTextBox1.Text + Environment.NewLine + "имя файла: " + nstr;
                    fname = nstr;
                    //nstr = nstr.Substring(nstr.IndexOf("_N_") + 3, nstr.IndexOf(".") - (nstr.IndexOf("_N_") + 3));
                    nstr = nstr.Substring(0, nstr.LastIndexOf("."));
                    //richTextBox1.Text = richTextBox1.Text + Environment.NewLine + "Проверяемое имя: " + nstr;
                    label1.Text = nstr;
                    if (list[j] == nstr)
                    {

                        //добавление ссылки
                        Excel.Range rangeToHoldHyperlink = (Excel.Range)ObjWorkSheet.Cells[j+1, 15];

                        string hyperlinkTargetAddress = file.FullName;
                        ObjWorkSheet.Hyperlinks.Add(rangeToHoldHyperlink,hyperlinkTargetAddress,"", "Screen Tip Text", hyperlinkTargetAddress);
                        //

                        //счетчик файлов которые успешно прошли проверку
                        acount = acount + 1;
                        label7.Text = acount.ToString();
                        //textBox3.Text = textBox3.Text + Environment.NewLine + "прошел проверку"+fname; //
                        chek = true;
                        break;
                    }
                    else
                    {
                        //счетчик + к файлам которых нет, добавить в текстбокс
                        
                        //label9.Text =flcount.ToString();
                        //textBox3.Text = textBox3.Text + Environment.NewLine + "не прошел проверку" + fname; //
                        //textBox3.Text = textBox3.Text + Environment.NewLine + fname;
                        chek = false;
                    }
                }
                if (chek == false)
                {
                    //textBox3.Text = textBox3.Text + Environment.NewLine + file.Name;
                    //textBox3.Text = textBox3.Text + Environment.NewLine + list[j];
                    richTextBox1.Text = richTextBox1.Text + Environment.NewLine + list[j];
                    flcount = flcount + 1;                    
                }
                label5.Text = count + " из " + LastRow;
            }
            label9.Text = flcount.ToString();

            //сохраняем результат
            //ObjWorkbook = ObjWorkExcel.Workbooks;
            //excelappworkbook = excelappworkbooks[1];

            //---
            //Excel.Range excelcells;
            //excelcells = (Excel.Range)ObjWorkSheet.Cells[1, 1];
            //excelcells.Value2 = "test";

            //excelcells = ObjWorkSheet.get_Range("A1", Type.Missing);
            //Выводим значение текстовую строку
            //excelcells.Value2 = "Лист 2";
            //excelcells.Font.Size = 20;
            //excelcells.Font.Italic = true;
            //excelcells.Font.Bold = true;

            ObjWorkExcel.DisplayAlerts = true;
            /*
            ObjWorkbook.SaveAs(Type.Missing,  //object Filename
               Type.Missing,                       //object FileFormat
               Type.Missing,                       //object Password 
               Type.Missing,                       //object WriteResPassword  
               Type.Missing,                       //object ReadOnlyRecommended
               Type.Missing,                       //object CreateBackup
               Excel.XlSaveAsAccessMode.xlNoChange,//XlSaveAsAccessMode AccessMode
               Type.Missing,                       //object ConflictResolution
               Type.Missing,                       //object AddToMru 
               Type.Missing,                       //object TextCodepage
               Type.Missing,                       //object TextVisualLayout
               Type.Missing);                      //object Local
               */
            ObjWorkbook.Save(); //делаем обычный сейв вместо "сохранить как"

            //ObjWorkbook.Close(false, Type.Missing, Type.Missing);
            ObjWorkExcel.Quit();
            GC.Collect();
            //
            MessageBox.Show("Документ сохранен");
        }

        private void button3_Click(object sender, EventArgs e)
        {
            var dialog = new FolderBrowserDialog();
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                textBox2.Text = dialog.SelectedPath;
                FPath= dialog.SelectedPath;

                DirectoryInfo dirInfo = new DirectoryInfo(FPath);
                fcount = dirInfo.GetFiles().Length.ToString();
                //label5.Text = "0 из " + dirInfo.GetFiles().Length.ToString();
                //label5.Text = "0 из " + fcount;
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            //string str = "abc_N_123_I_456_I_789.abc.text",str1="123/456/789";
            string str = "акт_ооо ромашка_15.06.2017_85, сч.ф 8653_I_15.txt", str1 = "123/456/789";
            
            label1.Text = str;
            int len = str.IndexOf(".") - str.IndexOf("_№_");
            //str = str.Substring(str.IndexOf("_№_")+3, str.IndexOf(".")-(str.IndexOf("_№_")+3));
            //str = str.Substring(str.IndexOf("_№_") + 3, str.LastIndexOf(".") - (str.IndexOf("_№_") + 3));
            str = str.Replace("_I_", "/");
            str= str.Substring(0, str.LastIndexOf(".")); ;
            label1.Text = label1.Text+" ; " + str+" ; "+(str==str1).ToString();
            /*
            for (int i = 0; i < 10; i++)
            {
                //label1.Text = (int)label1.Text + 1;
                textBox3.Text=textBox3.Text + Environment.NewLine + "dfgdfg";
            }
            
            //label1.Text = FPath + " ; " + XPath;
            //создание дохренища файлов-
            Excel.Application ObjWorkExcel = new Excel.Application();
            Excel.Workbook ObjWorkbook = ObjWorkExcel.Workbooks.Open(XPath, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkbook.Sheets[1];
            var LastCell = ObjWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);
            //
            int LastColumn = (int)LastCell.Column;
            int LastRow = (int)LastCell.Row;
            //
            string[] list = new string[LastCell.Row];
            label5.Text = "0 из " + (int)LastCell.Row; //кол-во строк
            Excel.Range excelcells;
            //excelcells = (Excel.Range)ObjWorkSheet.Cells[1, 1];
            //excelcells.Value2 = "test";

           
            //Выводим значение текстовую строку
            //excelcells.Value2 = "Лист 2";
            //excelcells.Font.Size = 20;
            //excelcells.Font.Italic = true;
            //excelcells.Font.Bold = true;


            /*string path,cell;
            int ind;
            Random rnd = new Random();
            for (int i = 0; i < 10; i++)
            {
                ind = i + 1;
                cell = "A" + ind;
                textBox3.Text = textBox3.Text + "A"+ind + Environment.NewLine;
                excelcells = ObjWorkSheet.get_Range(cell, Type.Missing);
                path = @"C:\Users\Михаил\Desktop\тестовые файлы\2\abc_N_"+rnd.Next(1000)+"_I_"+ rnd.Next(1000)+"_I_"+rnd.Next(1000)+".txt";
                System.IO.File.Create(path);

                path = path.Substring(path.IndexOf("_N_") + 3, path.IndexOf(".") - (path.IndexOf("_N_") + 3));
                path = path.Replace("_I_", "/");
                excelcells.Value2 = path;

            }
            ObjWorkExcel.DisplayAlerts = true;
            ObjWorkbook.Save(); //делаем обычный сейв вместо "сохранить как"

            ObjWorkExcel.Quit();
            GC.Collect();
            MessageBox.Show("Конец");*/

            //-
        }

        private void label8_Click(object sender, EventArgs e)
        {

        }

        private void label5_Click(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog OPF = new OpenFileDialog();
            if (OPF.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = OPF.FileName;
                XPath = OPF.FileName;
            }
        }
    }
}
