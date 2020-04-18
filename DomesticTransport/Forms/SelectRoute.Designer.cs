namespace DomesticTransport.Forms
{
    partial class RouteSelector
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
            this.btnAcept = new System.Windows.Forms.Button();
            this.btncancel = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // btnAcept
            // 
            this.btnAcept.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnAcept.Location = new System.Drawing.Point(67, 291);
            this.btnAcept.Name = "btnAcept";
            this.btnAcept.Size = new System.Drawing.Size(91, 29);
            this.btnAcept.TabIndex = 0;
            this.btnAcept.Text = "Принять";
            this.btnAcept.UseVisualStyleBackColor = true;
            // 
            // btncancel
            // 
            this.btncancel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btncancel.Location = new System.Drawing.Point(175, 291);
            this.btncancel.Name = "btncancel";
            this.btncancel.Size = new System.Drawing.Size(91, 29);
            this.btncancel.TabIndex = 0;
            this.btncancel.Text = "Отменить";
            this.btncancel.UseVisualStyleBackColor = true;
            // 
            // Calendar
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(269, 322);
            this.Controls.Add(this.btncancel);
            this.Controls.Add(this.btnAcept);
            this.Name = "Calendar";
            this.Text = "Calendar";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btnAcept;
        private System.Windows.Forms.Button btncancel;
    }
}