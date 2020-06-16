namespace DomesticTransport.Forms
{
    partial class SelectProvider
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(SelectProvider));
            this.ComboboxProvider = new System.Windows.Forms.ComboBox();
            this.label3 = new System.Windows.Forms.Label();
            this.ButtonСancel = new System.Windows.Forms.Button();
            this.ButtonAccept = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // ComboboxProvider
            // 
            this.ComboboxProvider.FormattingEnabled = true;
            this.ComboboxProvider.Location = new System.Drawing.Point(12, 26);
            this.ComboboxProvider.Name = "ComboboxProvider";
            this.ComboboxProvider.Size = new System.Drawing.Size(353, 21);
            this.ComboboxProvider.TabIndex = 12;
            // 
            // label3
            // 
            this.label3.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(9, 10);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(63, 13);
            this.label3.TabIndex = 11;
            this.label3.Text = "Провайдер";
            // 
            // ButtonСancel
            // 
            this.ButtonСancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.ButtonСancel.Location = new System.Drawing.Point(262, 53);
            this.ButtonСancel.Name = "ButtonСancel";
            this.ButtonСancel.Size = new System.Drawing.Size(104, 22);
            this.ButtonСancel.TabIndex = 9;
            this.ButtonСancel.Text = "Отменить";
            this.ButtonСancel.UseVisualStyleBackColor = true;
            this.ButtonСancel.Click += new System.EventHandler(this.ButtonСancel_Click);
            // 
            // ButtonAccept
            // 
            this.ButtonAccept.Location = new System.Drawing.Point(149, 53);
            this.ButtonAccept.Name = "ButtonAccept";
            this.ButtonAccept.Size = new System.Drawing.Size(104, 22);
            this.ButtonAccept.TabIndex = 10;
            this.ButtonAccept.Text = "Отправить";
            this.ButtonAccept.UseVisualStyleBackColor = true;
            this.ButtonAccept.Click += new System.EventHandler(this.ButtonAccept_Click);
            // 
            // SelectProvider
            // 
            this.AcceptButton = this.ButtonAccept;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.ButtonСancel;
            this.ClientSize = new System.Drawing.Size(371, 82);
            this.Controls.Add(this.ComboboxProvider);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.ButtonСancel);
            this.Controls.Add(this.ButtonAccept);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "SelectProvider";
            this.Text = "Выберите провайдера";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ComboBox ComboboxProvider;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button ButtonСancel;
        private System.Windows.Forms.Button ButtonAccept;
    }
}