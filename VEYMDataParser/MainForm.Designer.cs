namespace VEYMDataParser
{
    partial class MainForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainForm));
            this.BtnScrape = new System.Windows.Forms.Button();
            this.BtnMyself = new System.Windows.Forms.Button();
            this.BtnLogOff = new System.Windows.Forms.Button();
            this.buttonBETA = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // BtnScrape
            // 
            this.BtnScrape.Location = new System.Drawing.Point(168, 18);
            this.BtnScrape.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.BtnScrape.Name = "BtnScrape";
            this.BtnScrape.Size = new System.Drawing.Size(206, 42);
            this.BtnScrape.TabIndex = 0;
            this.BtnScrape.Text = "Scrape EVERYTHING!";
            this.BtnScrape.UseVisualStyleBackColor = true;
            this.BtnScrape.Click += new System.EventHandler(this.BtnScrape_Click);
            // 
            // BtnMyself
            // 
            this.BtnMyself.Location = new System.Drawing.Point(20, 18);
            this.BtnMyself.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.BtnMyself.Name = "BtnMyself";
            this.BtnMyself.Size = new System.Drawing.Size(123, 42);
            this.BtnMyself.TabIndex = 1;
            this.BtnMyself.Text = "Get Myself";
            this.BtnMyself.UseVisualStyleBackColor = true;
            this.BtnMyself.Click += new System.EventHandler(this.BtnMyself_Click);
            // 
            // BtnLogOff
            // 
            this.BtnLogOff.Font = new System.Drawing.Font("Microsoft Sans Serif", 16.2F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.BtnLogOff.Location = new System.Drawing.Point(425, 395);
            this.BtnLogOff.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.BtnLogOff.Name = "BtnLogOff";
            this.BtnLogOff.Size = new System.Drawing.Size(206, 42);
            this.BtnLogOff.TabIndex = 2;
            this.BtnLogOff.Text = "Log off";
            this.BtnLogOff.UseVisualStyleBackColor = true;
            this.BtnLogOff.Click += new System.EventHandler(this.BtnLogOff_Click);
            // 
            // buttonBETA
            // 
            this.buttonBETA.Location = new System.Drawing.Point(168, 83);
            this.buttonBETA.Margin = new System.Windows.Forms.Padding(2);
            this.buttonBETA.Name = "buttonBETA";
            this.buttonBETA.Size = new System.Drawing.Size(206, 42);
            this.buttonBETA.TabIndex = 3;
            this.buttonBETA.Text = "Scrape EVERYTHING (BETA)!";
            this.buttonBETA.UseVisualStyleBackColor = true;
            this.buttonBETA.Click += new System.EventHandler(this.buttonBETA_Click);
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(660, 470);
            this.Controls.Add(this.buttonBETA);
            this.Controls.Add(this.BtnLogOff);
            this.Controls.Add(this.BtnMyself);
            this.Controls.Add(this.BtnScrape);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.Name = "MainForm";
            this.Text = "VEYM Parser";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button BtnScrape;
        private System.Windows.Forms.Button BtnMyself;
        private System.Windows.Forms.Button BtnLogOff;
        private System.Windows.Forms.Button buttonBETA;
    }
}

