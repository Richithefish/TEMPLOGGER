namespace AXIS_LOGGER
{
    partial class Device_2_Dialog
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
            this.webBrowser_Device2 = new System.Windows.Forms.WebBrowser();
            this.SuspendLayout();
            // 
            // webBrowser_Device2
            // 
            this.webBrowser_Device2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.webBrowser_Device2.Location = new System.Drawing.Point(0, 0);
            this.webBrowser_Device2.MinimumSize = new System.Drawing.Size(20, 20);
            this.webBrowser_Device2.Name = "webBrowser_Device2";
            this.webBrowser_Device2.Size = new System.Drawing.Size(1061, 698);
            this.webBrowser_Device2.TabIndex = 0;
            // 
            // Device_2_Dialog
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1061, 698);
            this.Controls.Add(this.webBrowser_Device2);
            this.Name = "Device_2_Dialog";
            this.Text = "Device_2_Dialog";
            this.Load += new System.EventHandler(this.Device_2_Dialog_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.WebBrowser webBrowser_Device2;
    }
}