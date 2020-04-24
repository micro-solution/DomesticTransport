namespace DomesticTransport.Forms
{
    partial class Settings
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
            System.Windows.Forms.Button btnOFD;
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Settings));
            System.Windows.Forms.Button button1;
            this.label3 = new System.Windows.Forms.Label();
            this.tbTransortTable = new System.Windows.Forms.TextBox();
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnAcept = new System.Windows.Forms.Button();
            this.tbHelper = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            btnOFD = new System.Windows.Forms.Button();
            button1 = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // btnOFD
            // 
            btnOFD.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            btnOFD.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("btnOFD.BackgroundImage")));
            btnOFD.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom;
            btnOFD.FlatAppearance.BorderSize = 0;
            btnOFD.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            btnOFD.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            btnOFD.ImageAlign = System.Drawing.ContentAlignment.TopLeft;
            btnOFD.Location = new System.Drawing.Point(376, 24);
            btnOFD.Name = "btnOFD";
            btnOFD.Size = new System.Drawing.Size(20, 20);
            btnOFD.TabIndex = 9;
            btnOFD.UseVisualStyleBackColor = true;
            btnOFD.Click += new System.EventHandler(this.btnOFD_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(12, 7);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(79, 13);
            this.label3.TabIndex = 8;
            this.label3.Text = "TransportTable";
            // 
            // tbTransortTable
            // 
            this.tbTransortTable.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tbTransortTable.Location = new System.Drawing.Point(12, 25);
            this.tbTransortTable.Name = "tbTransortTable";
            this.tbTransortTable.Size = new System.Drawing.Size(358, 20);
            this.tbTransortTable.TabIndex = 7;
            // 
            // btnCancel
            // 
            this.btnCancel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.Location = new System.Drawing.Point(287, 117);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(104, 29);
            this.btnCancel.TabIndex = 5;
            this.btnCancel.Text = "Отменить";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // btnAcept
            // 
            this.btnAcept.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnAcept.Location = new System.Drawing.Point(179, 117);
            this.btnAcept.Name = "btnAcept";
            this.btnAcept.Size = new System.Drawing.Size(104, 29);
            this.btnAcept.TabIndex = 6;
            this.btnAcept.Text = "Принять";
            this.btnAcept.UseVisualStyleBackColor = true;
            this.btnAcept.Click += new System.EventHandler(this.btnAcept_Click);
            // 
            // tbHelper
            // 
            this.tbHelper.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tbHelper.Location = new System.Drawing.Point(12, 68);
            this.tbHelper.Name = "tbHelper";
            this.tbHelper.Size = new System.Drawing.Size(358, 20);
            this.tbHelper.TabIndex = 7;
            // 
            // button1
            // 
            button1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            button1.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("button1.BackgroundImage")));
            button1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom;
            button1.FlatAppearance.BorderSize = 0;
            button1.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            button1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            button1.ImageAlign = System.Drawing.ContentAlignment.TopLeft;
            button1.Location = new System.Drawing.Point(376, 67);
            button1.Name = "button1";
            button1.Size = new System.Drawing.Size(20, 20);
            button1.TabIndex = 9;
            button1.UseVisualStyleBackColor = true;
            button1.Click += new System.EventHandler(this.btnOFD_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(9, 51);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(56, 13);
            this.label1.TabIndex = 8;
            this.label1.Text = "Спркавка";
            // 
            // Settings
            // 
            this.AcceptButton = this.btnAcept;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.btnCancel;
            this.ClientSize = new System.Drawing.Size(402, 149);
            this.Controls.Add(button1);
            this.Controls.Add(btnOFD);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.tbHelper);
            this.Controls.Add(this.tbTransortTable);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnAcept);
            this.Name = "Settings";
            this.Text = "Настройки";
            this.Load += new System.EventHandler(this.Settings_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox tbTransortTable;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Button btnAcept;
        private System.Windows.Forms.TextBox tbHelper;
        private System.Windows.Forms.Label label1;
    }
}