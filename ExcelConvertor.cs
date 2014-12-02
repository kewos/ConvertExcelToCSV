using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Excel;
using System.Windows;


namespace ConvertExcelToCSV
{
    public class ExcelConvertor
    {
        Dictionary<string, System.Data.DataSet> NeedConvertExcelFiles = new Dictionary<string, System.Data.DataSet>();
        private string dicPath;
        private string DicPath
        {
            get
            {
                return dicPath;
            }
        }
        private string OutPutPath
        {
            get
            {
                string newDicName = "ConverToCSV";
                return string.Format("{0}\\{1}{2}", DicPath, newDicName, DateTime.Now.ToString("_yyyy-MM-dd_HH-mm-ss"));
            }
        }

        private static ExcelConvertor instance;
        public static ExcelConvertor Instance
        {
            get
            {
                if (instance == null)
                { 
                    instance = new ExcelConvertor(); 
                }
                return instance;
            }
        }

        public  void Excute(string dicPath)
        {
            this.dicPath = dicPath;
            SetNeedConvertExcels();
            converToCSV();
        }

        private void SetNeedConvertExcels()
        {
            NeedConvertExcelFiles.Clear();
            foreach (string file in System.IO.Directory.GetFileSystemEntries(DicPath, "*.xlsx"))
            {
                var filename = System.IO.Path.GetFileNameWithoutExtension(file);
                FileStream stream = File.Open(file, FileMode.Open, FileAccess.Read);
                using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream))
                {
                    var result = excelReader.AsDataSet();
                    NeedConvertExcelFiles.Add(filename, result);
                }
            }
        }

        private void converToCSV()
        {
            if (NeedConvertExcelFiles.Count == 0) return;
            if (!Directory.Exists(OutPutPath)) Directory.CreateDirectory(OutPutPath);

            foreach (var File in NeedConvertExcelFiles)
            {
                var content = GetExcelFile(File.Value);
                string output = string.Format("{0}\\{1}.csv", OutPutPath, File.Key);
                using (var sw = new StreamWriter(@output, false, System.Text.Encoding.UTF8))
                {
                    sw.Write(content);
                }
            }
            MessageBox.Show("Finish!!");
        }

        private static StringBuilder GetExcelFile(System.Data.DataSet File, int ind = 0)
        {
            var content = new StringBuilder();
            var rowNumber = 0;

            while (rowNumber++ < File.Tables[ind].Rows.Count)
            {
                for (int i = 0; i < File.Tables[ind].Columns.Count; i++)
                {
                    content.Append(File.Tables[ind].Rows[rowNumber][i].ToString() + ",");
                }
                content.Append(Environment.NewLine);
            }
            return content;
        }
    }
}
