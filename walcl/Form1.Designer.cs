namespace walcl
{
    partial class Form1
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.btExec = new System.Windows.Forms.Button();
            this.tbInfo = new System.Windows.Forms.TextBox();
            this.tbFolder = new System.Windows.Forms.TextBox();
            this.btFolder = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // btExec
            // 
            this.btExec.Location = new System.Drawing.Point(448, 324);
            this.btExec.Name = "btExec";
            this.btExec.Size = new System.Drawing.Size(75, 23);
            this.btExec.TabIndex = 0;
            this.btExec.Text = "执行";
            this.btExec.UseVisualStyleBackColor = true;
            this.btExec.Click += new System.EventHandler(this.btExec_Click);
            // 
            // tbInfo
            // 
            this.tbInfo.AllowDrop = true;
            this.tbInfo.Location = new System.Drawing.Point(12, 30);
            this.tbInfo.Multiline = true;
            this.tbInfo.Name = "tbInfo";
            this.tbInfo.ReadOnly = true;
            this.tbInfo.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.tbInfo.ShortcutsEnabled = false;
            this.tbInfo.Size = new System.Drawing.Size(512, 272);
            this.tbInfo.TabIndex = 1;
            this.tbInfo.DragDrop += new System.Windows.Forms.DragEventHandler(this.tbFolder_DragDrop);
            this.tbInfo.DragEnter += new System.Windows.Forms.DragEventHandler(this.tbFolder_DragEnter);
            // 
            // tbFolder
            // 
            this.tbFolder.AllowDrop = true;
            this.tbFolder.Location = new System.Drawing.Point(12, 325);
            this.tbFolder.Name = "tbFolder";
            this.tbFolder.Size = new System.Drawing.Size(330, 21);
            this.tbFolder.TabIndex = 2;
            this.tbFolder.DragDrop += new System.Windows.Forms.DragEventHandler(this.tbFolder_DragDrop);
            this.tbFolder.DragEnter += new System.Windows.Forms.DragEventHandler(this.tbFolder_DragEnter);
            // 
            // btFolder
            // 
            this.btFolder.Location = new System.Drawing.Point(358, 324);
            this.btFolder.Name = "btFolder";
            this.btFolder.Size = new System.Drawing.Size(75, 23);
            this.btFolder.TabIndex = 3;
            this.btFolder.Text = "文件夹";
            this.btFolder.UseVisualStyleBackColor = true;
            this.btFolder.Click += new System.EventHandler(this.btFolder_Click);
            // 
            // Form1
            // 
            this.AllowDrop = true;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(541, 361);
            this.Controls.Add(this.btFolder);
            this.Controls.Add(this.tbFolder);
            this.Controls.Add(this.tbInfo);
            this.Controls.Add(this.btExec);
            this.Name = "Form1";
            this.Text = "Office";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btExec;
        private System.Windows.Forms.TextBox tbInfo;
        private System.Windows.Forms.TextBox tbFolder;
        private System.Windows.Forms.Button btFolder;
    }
}

