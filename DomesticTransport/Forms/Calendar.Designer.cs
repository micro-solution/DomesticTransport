namespace DomesticTransport.Forms
{
    partial class Calendar

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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Calendar));
            this.btnAcept = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.calendarControl = new System.Windows.Forms.MonthCalendar();
            this.tbDate = new System.Windows.Forms.TextBox();
            this.btnFormula = new System.Windows.Forms.Button();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.SuspendLayout();
            // 
            // btnAcept
            // 
            this.btnAcept.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnAcept.Location = new System.Drawing.Point(3, 235);
            this.btnAcept.Name = "btnAcept";
            this.btnAcept.Size = new System.Drawing.Size(122, 27);
            this.btnAcept.TabIndex = 0;
            this.btnAcept.Text = "Принять";
            this.toolTip1.SetToolTip(this.btnAcept, "Установить дату отгрузки");
            this.btnAcept.UseVisualStyleBackColor = true;
            this.btnAcept.Click += new System.EventHandler(this.BtnAccept_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.Location = new System.Drawing.Point(134, 235);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(122, 27);
            this.btnCancel.TabIndex = 0;
            this.btnCancel.Text = "Отменить";
            this.btnCancel.UseVisualStyleBackColor = true;
            // 
            // calendarControl
            // 
            this.calendarControl.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.calendarControl.Cursor = System.Windows.Forms.Cursors.Hand;
            this.calendarControl.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.calendarControl.Location = new System.Drawing.Point(47, 37);
            this.calendarControl.MaxDate = new System.DateTime(2200, 12, 31, 0, 0, 0, 0);
            this.calendarControl.MaxSelectionCount = 1;
            this.calendarControl.MinDate = new System.DateTime(2010, 1, 1, 0, 0, 0, 0);
            this.calendarControl.Name = "calendarControl";
            this.calendarControl.TabIndex = 1;
            this.calendarControl.DateChanged += new System.Windows.Forms.DateRangeEventHandler(this.CalendarControl_DateChanged);
            // 
            // tbDate
            // 
            this.tbDate.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tbDate.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.tbDate.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.tbDate.Location = new System.Drawing.Point(3, 11);
            this.tbDate.Name = "tbDate";
            this.tbDate.Size = new System.Drawing.Size(253, 15);
            this.tbDate.TabIndex = 2;
            this.tbDate.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.tbDate.TextChanged += new System.EventHandler(this.TbDate_TextChanged);
            // 
            // btnFormula
            // 
            this.btnFormula.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.btnFormula.Location = new System.Drawing.Point(3, 206);
            this.btnFormula.Name = "btnFormula";
            this.btnFormula.Size = new System.Drawing.Size(253, 25);
            this.btnFormula.TabIndex = 3;
            this.btnFormula.Text = "Вставить формулу";
            this.toolTip1.SetToolTip(this.btnFormula, "Определить дату с помощью формулы (Следующий рабочий день)");
            this.btnFormula.UseVisualStyleBackColor = true;
            this.btnFormula.Click += new System.EventHandler(this.BtnFormula_Click);
            // 
            // Calendar
            // 
            this.AcceptButton = this.btnAcept;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.Window;
            this.CancelButton = this.btnCancel;
            this.ClientSize = new System.Drawing.Size(259, 265);
            this.Controls.Add(this.btnFormula);
            this.Controls.Add(this.tbDate);
            this.Controls.Add(this.calendarControl);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnAcept);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MinimumSize = new System.Drawing.Size(235, 301);
            this.Name = "Calendar";
            this.Text = "Дата отгрузки";
            this.Load += new System.EventHandler(this.Calendar_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnAcept;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.MonthCalendar calendarControl;
        private System.Windows.Forms.TextBox tbDate;
        private System.Windows.Forms.Button btnFormula;
        private System.Windows.Forms.ToolTip toolTip1;
    }
}