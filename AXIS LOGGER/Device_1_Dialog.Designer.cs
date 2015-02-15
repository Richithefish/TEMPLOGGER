namespace AXIS_LOGGER
{
    partial class Device_1_Dialog
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
            this.webBrowser_Device1 = new System.Windows.Forms.WebBrowser();
            this.SuspendLayout();
            // 
            // webBrowser_Device1
            // 
            this.webBrowser_Device1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.webBrowser_Device1.Location = new System.Drawing.Point(0, 0);
            this.webBrowser_Device1.MinimumSize = new System.Drawing.Size(20, 20);
            this.webBrowser_Device1.Name = "webBrowser_Device1";
            this.webBrowser_Device1.Size = new System.Drawing.Size(1061, 698);
            this.webBrowser_Device1.TabIndex = 0;
            // 
            // Device_1_Dialog
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1061, 698);
            this.Controls.Add(this.webBrowser_Device1);
            this.Name = "Device_1_Dialog";
            this.Text = "Device_1_Dialog";
            this.Load += new System.EventHandler(this.Device_1_Dialog_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.WebBrowser webBrowser_Device1;
    }
}