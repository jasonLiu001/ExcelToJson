﻿namespace ExcelToJson
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
            this.button2 = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(30, 28);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(81, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Excel文件路径";
            // 
            // txt_excelFilePath
            // 
            this.txt_excelFilePath.Location = new System.Drawing.Point(117, 25);
            this.txt_excelFilePath.Name = "txt_excelFilePath";
            this.txt_excelFilePath.Size = new System.Drawing.Size(281, 20);
            this.txt_excelFilePath.TabIndex = 1;
            // 
            // btn_broswerFile
            // 
            this.btn_broswerFile.Location = new System.Drawing.Point(404, 23);
            this.btn_broswerFile.Name = "btn_broswerFile";
            this.btn_broswerFile.Size = new System.Drawing.Size(75, 23);
            this.btn_broswerFile.TabIndex = 2;
            this.btn_broswerFile.Text = "浏览";
            this.btn_broswerFile.UseVisualStyleBackColor = true;
            this.btn_broswerFile.Click += new System.EventHandler(this.btn_broswerFile_Click);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(71, 80);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(75, 23);
            this.button2.TabIndex = 3;
            this.button2.Text = "To Json";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // Main
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(515, 139);
            this.Controls.Add(this.button2);
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
        private System.Windows.Forms.Button button2;
    }
}
