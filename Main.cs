using Newtonsoft.Json;
using Newtonsoft.Json.Converters;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelToJson
{
    public partial class Main : Form
    {
        public Main()
        {
            InitializeComponent();
        }

        private void btn_broswerFile_Click(object sender, EventArgs e)
        {
            OpenFileDialog fileDialog = new OpenFileDialog();
            fileDialog.Title = "选择Excel文件";
            fileDialog.Filter = "All Files(*.*)|*.*";
            if (fileDialog.ShowDialog() == DialogResult.OK)
            {
                string file = fileDialog.FileName;
                this.txt_excelFilePath.Text = file.Trim();
            }
        }

        public delegate void ShowResult(string msg);

        public void showResultToTextBox(string msg)
        {
            this.txt_result.Text = msg;
        }

        private void btn_toJson_Click(object sender, EventArgs e)
        {
            var excelFilePath = this.txt_excelFilePath.Text.Trim();
            if (string.IsNullOrEmpty(excelFilePath))
            {
                MessageBox.Show("请选择Excel文件路径");
                return;
            }          

            Task.Factory.StartNew(() =>
            {
                var excelDataTable = NPOIUtility.GetDataFromExcel(excelFilePath);
                Newtonsoft.Json.JsonSerializerSettings serializerSettings = new Newtonsoft.Json.JsonSerializerSettings();
                serializerSettings.Converters.Add(new DataTableConverter());
                var jsonString = JsonConvert.SerializeObject(excelDataTable, Formatting.None, serializerSettings);
                this.txt_result.Invoke(new ShowResult(showResultToTextBox),jsonString);
            });

            this.txt_result.Text = "正在转换中...请稍后...";
        }

        
        private void btn_saveToFile_Click(object sender, EventArgs e)
        {

        }
    }
}
