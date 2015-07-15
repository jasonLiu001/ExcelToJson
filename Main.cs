using Newtonsoft.Json;
using Newtonsoft.Json.Converters;
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

namespace ExcelToJson
{
    public partial class Main : Form
    {
        private static string jsonResult = string.Empty;

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
            this.txt_result.AppendText(msg + "\r\n");
        }

        private void btn_toJson_Click(object sender, EventArgs e)
        {
            this.DisableButtonControls();
            TaskScheduler m_syncContextTaskScheduler = TaskScheduler.FromCurrentSynchronizationContext();
            var excelFilePath = this.txt_excelFilePath.Text.Trim();
            if (string.IsNullOrEmpty(excelFilePath))
            {
                MessageBox.Show("请选择Excel文件路径");
                return;
            }

            Task<string> task = new Task<string>(() =>
            {
                var excelDataTable = NPOIUtility.GetDataFromExcel(excelFilePath);
                Newtonsoft.Json.JsonSerializerSettings serializerSettings = new Newtonsoft.Json.JsonSerializerSettings();
                serializerSettings.Converters.Add(new DataTableConverter());
                jsonResult = JsonConvert.SerializeObject(excelDataTable, Formatting.None, serializerSettings);
                return jsonResult;
            });

            task.ContinueWith(t =>
            {
                this.showLogMessage("转换完成");
                this.showLogMessage("点击保存按钮，可保存为单一的Json文件");
                this.EnableButtonControls();
            }, m_syncContextTaskScheduler);

            task.Start();

            //Task.Factory.StartNew(() =>
            // {
            //     var excelDataTable = NPOIUtility.GetDataFromExcel(excelFilePath);
            //     Newtonsoft.Json.JsonSerializerSettings serializerSettings = new Newtonsoft.Json.JsonSerializerSettings();
            //     serializerSettings.Converters.Add(new DataTableConverter());
            //     var jsonString = JsonConvert.SerializeObject(excelDataTable, Formatting.None, serializerSettings);
            //     this.txt_result.Invoke(new ShowResult(showResultToTextBox), jsonString);
            // });

            showLogMessage("正在转换中...请稍后...");
        }

        private void btn_saveToFile_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(jsonResult))
            {
                MessageBox.Show("未进行转换操作，或者正在转换中...");
                return;
            }

            SaveFileDialog saveDialog = new SaveFileDialog();
            saveDialog.CreatePrompt = true;
            saveDialog.OverwritePrompt = true;

            saveDialog.DefaultExt = "json";
            saveDialog.Filter = "Text Files(*.json)|*.json";

            if (DialogResult.OK == saveDialog.ShowDialog())
            {
                if (File.Exists(saveDialog.FileName))
                {
                    DialogResult result = MessageBox.Show("是否覆盖？", "确定", MessageBoxButtons.YesNo);
                    if (result == DialogResult.Yes)
                    {
                        StreamWriter swLog = new StreamWriter(saveDialog.FileName);
                        swLog.WriteLine(jsonResult);
                        swLog.WriteLine();
                        swLog.Close();
                    }
                    if (result == DialogResult.No)
                    {
                        return;
                    }
                }
                else
                {
                    StreamWriter swLog = new StreamWriter(saveDialog.FileName);
                    swLog.WriteLine(jsonResult);
                    swLog.WriteLine();
                    swLog.Close();
                }
            }
        }

        private void btn_updateWaterData_Click(object sender, EventArgs e)
        {
            showLogMessage("正在更新水体Excel数据..");
            this.DisableButtonControls();
            TaskScheduler m_syncContextTaskScheduler = TaskScheduler.FromCurrentSynchronizationContext();
            var excelFilePath = this.txt_excelFilePath.Text.Trim();
            if (string.IsNullOrEmpty(excelFilePath))
            {
                MessageBox.Show("请选择Excel文件路径");
                return;
            }

            Task<bool> task = new Task<bool>(() =>
            {
                bool message = UpdateExcel.UpdateExcelData(excelFilePath, Constant.Water);
                return message;
            });

            task.ContinueWith(t =>
            {
                if (task.Result)
                {
                    this.showLogMessage("水体Excel数据生成完成！");
                    this.showLogMessage("生成文件路径：d:\\" + Constant.Water + ".xlsx");
                }
                else
                {
                    this.showLogMessage("水体Excel数据生成失败！");
                }
                this.EnableButtonControls();
            }, m_syncContextTaskScheduler);

            task.Start();
        }

        private void btn_updateAirData_Click(object sender, EventArgs e)
        {
            showLogMessage("正在更新空气Excel数据..");
            this.DisableButtonControls();
            this.btn_saveToFile.Enabled = false;
            this.btn_broswerFile.Enabled = false;
            this.btn_toJson.Enabled = false;
            this.btn_updateWaterData.Enabled = false;
            TaskScheduler m_syncContextTaskScheduler = TaskScheduler.FromCurrentSynchronizationContext();
            var excelFilePath = this.txt_excelFilePath.Text.Trim();
            if (string.IsNullOrEmpty(excelFilePath))
            {
                MessageBox.Show("请选择Excel文件路径");
                return;
            }

            Task<bool> task = new Task<bool>(() =>
            {
                bool message = UpdateExcel.UpdateExcelData(excelFilePath, Constant.Air);
                return message;
            });

            task.ContinueWith(t =>
            {
                if (task.Result)
                {
                    this.showLogMessage("空气Excel数据生成完成！");
                    this.showLogMessage("生成文件路径：d:\\" + Constant.Air + ".xlsx");                    
                }
                else
                {
                    this.showLogMessage("空气Excel数据生成失败！");
                }
                this.EnableButtonControls();
            }, m_syncContextTaskScheduler);

            task.Start();
        }

        private void showLogMessage(string message)
        {
            this.txt_result.AppendText(message + "\r\n");
        }

        private void EnableButtonControls()
        {
            this.btn_saveToFile.Enabled = true;
            this.btn_broswerFile.Enabled = true;
            this.btn_toJson.Enabled = true;
            this.btn_updateWaterData.Enabled = true;
            this.btn_updateAirData.Enabled = true;
        }

        private void DisableButtonControls()
        {
            this.btn_saveToFile.Enabled = false;
            this.btn_broswerFile.Enabled = false;
            this.btn_toJson.Enabled = false;
            this.btn_updateWaterData.Enabled = false;
            this.btn_updateAirData.Enabled = false;
        }
    }
}
