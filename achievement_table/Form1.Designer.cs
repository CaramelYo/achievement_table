namespace achievement_table
{
    partial class Form1
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
            this.log_tbx = new System.Windows.Forms.TextBox();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.檔案ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.執行ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.achievement_table_generator_tsmi = new System.Windows.Forms.ToolStripMenuItem();
            this.monthly_achievement_table_generator_tsmi = new System.Windows.Forms.ToolStripMenuItem();
            this.menuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // log_tbx
            // 
            this.log_tbx.Location = new System.Drawing.Point(303, 23);
            this.log_tbx.Multiline = true;
            this.log_tbx.Name = "log_tbx";
            this.log_tbx.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.log_tbx.Size = new System.Drawing.Size(484, 371);
            this.log_tbx.TabIndex = 0;
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.檔案ToolStripMenuItem,
            this.執行ToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(825, 24);
            this.menuStrip1.TabIndex = 1;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // 檔案ToolStripMenuItem
            // 
            this.檔案ToolStripMenuItem.Name = "檔案ToolStripMenuItem";
            this.檔案ToolStripMenuItem.Size = new System.Drawing.Size(43, 20);
            this.檔案ToolStripMenuItem.Text = "檔案";
            // 
            // 執行ToolStripMenuItem
            // 
            this.執行ToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.achievement_table_generator_tsmi,
            this.monthly_achievement_table_generator_tsmi});
            this.執行ToolStripMenuItem.Name = "執行ToolStripMenuItem";
            this.執行ToolStripMenuItem.Size = new System.Drawing.Size(43, 20);
            this.執行ToolStripMenuItem.Text = "執行";
            // 
            // achievement_table_generator_tsmi
            // 
            this.achievement_table_generator_tsmi.Name = "achievement_table_generator_tsmi";
            this.achievement_table_generator_tsmi.Size = new System.Drawing.Size(182, 22);
            this.achievement_table_generator_tsmi.Text = "產生分隊部績效表";
            this.achievement_table_generator_tsmi.Click += new System.EventHandler(this.achievement_table_generator_tsmi_Click);
            // 
            // monthly_achievement_table_generator_tsmi
            // 
            this.monthly_achievement_table_generator_tsmi.Name = "monthly_achievement_table_generator_tsmi";
            this.monthly_achievement_table_generator_tsmi.Size = new System.Drawing.Size(182, 22);
            this.monthly_achievement_table_generator_tsmi.Text = "產生分隊部月績效表";
            this.monthly_achievement_table_generator_tsmi.Click += new System.EventHandler(this.monthly_achievement_table_generator_tsmi_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(825, 416);
            this.Controls.Add(this.log_tbx);
            this.Controls.Add(this.menuStrip1);
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "Form1";
            this.Text = "Form1";
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox log_tbx;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem 檔案ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 執行ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem achievement_table_generator_tsmi;
        private System.Windows.Forms.ToolStripMenuItem monthly_achievement_table_generator_tsmi;
    }
}

