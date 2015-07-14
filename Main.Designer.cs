namespace ExcelToJson
{
    partial class Main
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.label1 = new System.Windows.Forms.Label();
            this.txt_excelFilePath = new System.Windows.Forms.TextBox();
            this.btn_broswerFile = new System.Windows.Forms.Button();
            this.btn_toJson = new System.Windows.Forms.Button();
            this.txt_result = new System.Windows.Forms.TextBox();
            this.btn_saveToFile = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(19, 28);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(81, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Excel文件路径";
            // 
            // txt_excelFilePath
            // 
            this.txt_excelFilePath.BackColor = System.Drawing.SystemColors.ButtonHighlight;
            this.txt_excelFilePath.Location = new System.Drawing.Point(108, 25);
            this.txt_excelFilePath.Name = "txt_excelFilePath";
            this.txt_excelFilePath.ReadOnly = true;
            this.txt_excelFilePath.Size = new System.Drawing.Size(281, 20);
            this.txt_excelFilePath.TabIndex = 1;
            // 
            // btn_broswerFile
            // 
            this.btn_broswerFile.Location = new System.Drawing.Point(401, 23);
            this.btn_broswerFile.Name = "btn_broswerFile";
            this.btn_broswerFile.Size = new System.Drawing.Size(75, 23);
            this.btn_broswerFile.TabIndex = 2;
            this.btn_broswerFile.Text = "浏览";
            this.btn_broswerFile.UseVisualStyleBackColor = true;
            this.btn_broswerFile.Click += new System.EventHandler(this.btn_broswerFile_Click);
            // 
            // btn_toJson
            // 
            this.btn_toJson.Location = new System.Drawing.Point(308, 160);
            this.btn_toJson.Name = "btn_toJson";
            this.btn_toJson.Size = new System.Drawing.Size(75, 23);
            this.btn_toJson.TabIndex = 3;
            this.btn_toJson.Text = "转换";
            this.btn_toJson.UseVisualStyleBackColor = true;
            this.btn_toJson.Click += new System.EventHandler(this.btn_toJson_Click);
            // 
            // txt_result
            // 
            this.txt_result.Location = new System.Drawing.Point(21, 56);
            this.txt_result.Multiline = true;
            this.txt_result.Name = "txt_result";
            this.txt_result.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.txt_result.Size = new System.Drawing.Size(455, 95);
            this.txt_result.TabIndex = 4;
            // 
            // btn_saveToFile
            // 
            this.btn_saveToFile.Location = new System.Drawing.Point(401, 160);
            this.btn_saveToFile.Name = "btn_saveToFile";
            this.btn_saveToFile.Size = new System.Drawing.Size(75, 23);
            this.btn_saveToFile.TabIndex = 5;
            this.btn_saveToFile.Text = "保存";
            this.btn_saveToFile.UseVisualStyleBackColor = true;
            this.btn_saveToFile.Click += new System.EventHandler(this.btn_saveToFile_Click);
            // 
            // Main
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(503, 198);
            this.Controls.Add(this.btn_saveToFile);
            this.Controls.Add(this.txt_result);
            this.Controls.Add(this.btn_toJson);
            this.Controls.Add(this.btn_broswerFile);
            this.Controls.Add(this.txt_excelFilePath);
            this.Controls.Add(this.label1);
            this.MaximizeBox = false;
            this.Name = "Main";
            this.Text = "Excel To Json";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txt_excelFilePath;
        private System.Windows.Forms.Button btn_broswerFile;
        private System.Windows.Forms.Button btn_toJson;
        private System.Windows.Forms.TextBox txt_result;
        private System.Windows.Forms.Button btn_saveToFile;
    }
}

