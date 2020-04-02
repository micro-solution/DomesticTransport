namespace DomesticTransport.Forms
{
    partial class ProviderEditor
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
            this.lvProvider = new System.Windows.Forms.ListView();
            this.btnAccept = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.lvMap = new System.Windows.Forms.ListView();
            this.columnProvider = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnCost = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.tbWeight = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // lvProvider
            // 
            this.lvProvider.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.lvProvider.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.columnProvider,
            this.columnCost});
            this.lvProvider.HideSelection = false;
            this.lvProvider.Location = new System.Drawing.Point(2, 22);
            this.lvProvider.Name = "lvProvider";
            this.lvProvider.Size = new System.Drawing.Size(277, 121);
            this.lvProvider.TabIndex = 0;
            this.lvProvider.UseCompatibleStateImageBehavior = false;
            this.lvProvider.View = System.Windows.Forms.View.Details;
            // 
            // btnAccept
            // 
            this.btnAccept.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnAccept.Location = new System.Drawing.Point(211, 149);
            this.btnAccept.Name = "btnAccept";
            this.btnAccept.Size = new System.Drawing.Size(93, 23);
            this.btnAccept.TabIndex = 1;
            this.btnAccept.Text = "Принять";
            this.btnAccept.UseVisualStyleBackColor = true;
            // 
            // btnCancel
            // 
            this.btnCancel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.Location = new System.Drawing.Point(318, 149);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(93, 23);
            this.btnCancel.TabIndex = 1;
            this.btnCancel.Text = "Отменить";
            this.btnCancel.UseVisualStyleBackColor = true;
            // 
            // lvMap
            // 
            this.lvMap.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.lvMap.HideSelection = false;
            this.lvMap.Location = new System.Drawing.Point(285, 22);
            this.lvMap.Name = "lvMap";
            this.lvMap.Size = new System.Drawing.Size(146, 121);
            this.lvMap.TabIndex = 0;
            this.lvMap.UseCompatibleStateImageBehavior = false;
            // 
            // columnProvider
            // 
            this.columnProvider.Text = "Провайдер";
            this.columnProvider.Width = 136;
            // 
            // columnCost
            // 
            this.columnCost.Text = "Стоимость";
            this.columnCost.Width = 80;
            // 
            // tbWeight
            // 
            this.tbWeight.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.tbWeight.Location = new System.Drawing.Point(330, 1);
            this.tbWeight.Name = "tbWeight";
            this.tbWeight.Size = new System.Drawing.Size(100, 20);
            this.tbWeight.TabIndex = 2;
            // 
            // label1
            // 
            this.label1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(250, 4);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(74, 13);
            this.label1.TabIndex = 3;
            this.label1.Text = "Вес груза, кг";
            // 
            // Provider
            // 
            this.AcceptButton = this.btnAccept;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.btnCancel;
            this.ClientSize = new System.Drawing.Size(434, 174);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.tbWeight);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnAccept);
            this.Controls.Add(this.lvMap);
            this.Controls.Add(this.lvProvider);
            this.Name = "Provider";
            this.Text = "Смена провайдера";
            this.Load += new System.EventHandler(this.Provider_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ListView lvProvider;
        private System.Windows.Forms.ColumnHeader columnProvider;
        private System.Windows.Forms.ColumnHeader columnCost;
        private System.Windows.Forms.Button btnAccept;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.ListView lvMap;
        private System.Windows.Forms.TextBox tbWeight;
        private System.Windows.Forms.Label label1;
    }
}