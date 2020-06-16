namespace DomesticTransport
{
    partial class SapFiles
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
            System.Windows.Forms.Button button2;
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(SapFiles));
            System.Windows.Forms.Button button1;
            this.tbOrders = new System.Windows.Forms.TextBox();
            this.tbExport = new System.Windows.Forms.TextBox();
            this.Cancel = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.Accept = new System.Windows.Forms.Button();
            this.calendarControl = new System.Windows.Forms.MonthCalendar();
            this.label3 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            button2 = new System.Windows.Forms.Button();
            button1 = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // button2
            // 
            button2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            button2.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("button2.BackgroundImage")));
            button2.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom;
            button2.FlatAppearance.BorderSize = 0;
            button2.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            button2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            button2.ImageAlign = System.Drawing.ContentAlignment.TopLeft;
            button2.Location = new System.Drawing.Point(718, 47);
            button2.Name = "button2";
            button2.Size = new System.Drawing.Size(20, 20);
            button2.TabIndex = 3;
            button2.UseVisualStyleBackColor = true;
            button2.Click += new System.EventHandler(this.SelectExport_Click);
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
            button1.Location = new System.Drawing.Point(718, 96);
            button1.Name = "button1";
            button1.Size = new System.Drawing.Size(20, 20);
            button1.TabIndex = 3;
            button1.UseVisualStyleBackColor = true;
            button1.Click += new System.EventHandler(this.SelectAllOrders_Click);
            // 
            // tbOrders
            // 
            this.tbOrders.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tbOrders.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.tbOrders.Location = new System.Drawing.Point(194, 96);
            this.tbOrders.Name = "tbOrders";
            this.tbOrders.Size = new System.Drawing.Size(518, 20);
            this.tbOrders.TabIndex = 0;
            // 
            // tbExport
            // 
            this.tbExport.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tbExport.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.tbExport.Location = new System.Drawing.Point(194, 47);
            this.tbExport.Name = "tbExport";
            this.tbExport.Size = new System.Drawing.Size(518, 20);
            this.tbExport.TabIndex = 0;
            // 
            // Cancel
            // 
            this.Cancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.Cancel.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.Cancel.Location = new System.Drawing.Point(652, 162);
            this.Cancel.Name = "Cancel";
            this.Cancel.Size = new System.Drawing.Size(86, 23);
            this.Cancel.TabIndex = 1;
            this.Cancel.Text = "Отменить";
            this.Cancel.UseVisualStyleBackColor = true;
            this.Cancel.Click += new System.EventHandler(this.Cancel_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.label2.Location = new System.Drawing.Point(191, 28);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(312, 13);
            this.label2.TabIndex = 2;
            this.label2.Text = "Выберите скачанный из SAP файл c данными по отгрузкам";
            // 
            // Accept
            // 
            this.Accept.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.Accept.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.Accept.Location = new System.Drawing.Point(560, 162);
            this.Accept.Name = "Accept";
            this.Accept.Size = new System.Drawing.Size(86, 23);
            this.Accept.TabIndex = 1;
            this.Accept.Text = "Выбрать";
            this.Accept.UseVisualStyleBackColor = true;
            this.Accept.Click += new System.EventHandler(this.Accept_Click);
            // 
            // calendarControl
            // 
            this.calendarControl.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.calendarControl.Cursor = System.Windows.Forms.Cursors.Hand;
            this.calendarControl.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.calendarControl.Location = new System.Drawing.Point(15, 28);
            this.calendarControl.MaxDate = new System.DateTime(2200, 12, 31, 0, 0, 0, 0);
            this.calendarControl.MaxSelectionCount = 1;
            this.calendarControl.MinDate = new System.DateTime(2010, 1, 1, 0, 0, 0, 0);
            this.calendarControl.Name = "calendarControl";
            this.calendarControl.ShowTodayCircle = false;
            this.calendarControl.TabIndex = 6;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.label3.Location = new System.Drawing.Point(191, 78);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(271, 13);
            this.label3.TabIndex = 7;
            this.label3.Text = "Выберите файл All Orders с заявками (опционально)";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.label1.Location = new System.Drawing.Point(12, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(81, 13);
            this.label1.TabIndex = 8;
            this.label1.Text = "Дата отгрузки";
            // 
            // SapFiles
            // 
            this.AcceptButton = this.Accept;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.Cancel;
            this.ClientSize = new System.Drawing.Size(747, 197);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.label3);
            this.Controls.Add(button1);
            this.Controls.Add(this.tbOrders);
            this.Controls.Add(button2);
            this.Controls.Add(this.tbExport);
            this.Controls.Add(this.calendarControl);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.Accept);
            this.Controls.Add(this.Cancel);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Name = "SapFiles";
            this.Text = "Выбор файлов";
            this.Load += new System.EventHandler(this.SapFiles_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox tbOrders;
        private System.Windows.Forms.TextBox tbExport;
        private System.Windows.Forms.Button Cancel;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button Accept;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label1;
        public System.Windows.Forms.MonthCalendar calendarControl;
    }
}