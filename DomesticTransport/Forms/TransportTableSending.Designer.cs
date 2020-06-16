namespace DomesticTransport.Forms
{
    partial class TransportTableSending
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
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(TransportTableSending));
            this.ButtonСancel = new System.Windows.Forms.Button();
            this.ButtonAccept = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.label3 = new System.Windows.Forms.Label();
            this.ComboboxProvider = new System.Windows.Forms.ComboBox();
            this.monthCalendarStart = new System.Windows.Forms.MonthCalendar();
            this.monthCalendarEnd = new System.Windows.Forms.MonthCalendar();
            this.SuspendLayout();
            // 
            // ButtonСancel
            // 
            this.ButtonСancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.ButtonСancel.Location = new System.Drawing.Point(258, 245);
            this.ButtonСancel.Name = "ButtonСancel";
            this.ButtonСancel.Size = new System.Drawing.Size(104, 22);
            this.ButtonСancel.TabIndex = 1;
            this.ButtonСancel.Text = "Отменить";
            this.ButtonСancel.UseVisualStyleBackColor = true;
            this.ButtonСancel.Click += new System.EventHandler(this.ButtonСancel_Click);
            // 
            // ButtonAccept
            // 
            this.ButtonAccept.Location = new System.Drawing.Point(145, 245);
            this.ButtonAccept.Name = "ButtonAccept";
            this.ButtonAccept.Size = new System.Drawing.Size(104, 22);
            this.ButtonAccept.TabIndex = 2;
            this.ButtonAccept.Text = "Отправить";
            this.ButtonAccept.UseVisualStyleBackColor = true;
            this.ButtonAccept.Click += new System.EventHandler(this.ButtonAccept_Click);
            // 
            // label2
            // 
            this.label2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(5, 9);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(116, 13);
            this.label2.TabIndex = 6;
            this.label2.Text = "Дата начала периода";
            // 
            // label1
            // 
            this.label1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(195, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(111, 13);
            this.label1.TabIndex = 6;
            this.label1.Text = "Дата конца периода";
            // 
            // label3
            // 
            this.label3.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(5, 202);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(63, 13);
            this.label3.TabIndex = 7;
            this.label3.Text = "Провайдер";
            // 
            // ComboboxProvider
            // 
            this.ComboboxProvider.FormattingEnabled = true;
            this.ComboboxProvider.Location = new System.Drawing.Point(8, 218);
            this.ComboboxProvider.Name = "ComboboxProvider";
            this.ComboboxProvider.Size = new System.Drawing.Size(353, 21);
            this.ComboboxProvider.TabIndex = 8;
            // 
            // monthCalendarStart
            // 
            this.monthCalendarStart.Location = new System.Drawing.Point(8, 31);
            this.monthCalendarStart.MaxSelectionCount = 1;
            this.monthCalendarStart.Name = "monthCalendarStart";
            this.monthCalendarStart.ShowToday = false;
            this.monthCalendarStart.ShowTodayCircle = false;
            this.monthCalendarStart.TabIndex = 9;
            // 
            // monthCalendarEnd
            // 
            this.monthCalendarEnd.Location = new System.Drawing.Point(198, 31);
            this.monthCalendarEnd.MaxSelectionCount = 1;
            this.monthCalendarEnd.Name = "monthCalendarEnd";
            this.monthCalendarEnd.ShowToday = false;
            this.monthCalendarEnd.ShowTodayCircle = false;
            this.monthCalendarEnd.TabIndex = 10;
            // 
            // UnloadOnDate
            // 
            this.AcceptButton = this.ButtonAccept;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.ButtonСancel;
            this.ClientSize = new System.Drawing.Size(371, 273);
            this.Controls.Add(this.monthCalendarEnd);
            this.Controls.Add(this.monthCalendarStart);
            this.Controls.Add(this.ComboboxProvider);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.ButtonСancel);
            this.Controls.Add(this.ButtonAccept);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "UnloadOnDate";
            this.Text = "Отправить провайдерам";
            this.Load += new System.EventHandler(this.UnloadOnDate_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button ButtonСancel;
        private System.Windows.Forms.Button ButtonAccept;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ToolTip toolTip1;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ComboBox ComboboxProvider;
        private System.Windows.Forms.MonthCalendar monthCalendarStart;
        private System.Windows.Forms.MonthCalendar monthCalendarEnd;
    }
}