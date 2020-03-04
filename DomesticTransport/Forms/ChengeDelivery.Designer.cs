namespace DomesticTransport.Forms
{
    partial class ChengeDelivery
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
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.lbTruck = new System.Windows.Forms.Label();
            this.btnDelete = new System.Windows.Forms.Button();
            this.Accept = new System.Windows.Forms.Button();
            this.Cancel = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // dataGridView1
            // 
            this.dataGridView1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(2, 49);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.Size = new System.Drawing.Size(776, 91);
            this.dataGridView1.TabIndex = 0;
            // 
            // lbTruck
            // 
            this.lbTruck.AutoSize = true;
            this.lbTruck.Location = new System.Drawing.Point(13, 30);
            this.lbTruck.Name = "lbTruck";
            this.lbTruck.Size = new System.Drawing.Size(35, 13);
            this.lbTruck.TabIndex = 1;
            this.lbTruck.Text = "label1";
            // 
            // btnDelete
            // 
            this.btnDelete.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnDelete.Location = new System.Drawing.Point(699, 20);
            this.btnDelete.Name = "btnDelete";
            this.btnDelete.Size = new System.Drawing.Size(75, 23);
            this.btnDelete.TabIndex = 2;
            this.btnDelete.Text = "Удалить";
            this.btnDelete.UseVisualStyleBackColor = true;
            // 
            // Accept
            // 
            this.Accept.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.Accept.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.Accept.Location = new System.Drawing.Point(586, 149);
            this.Accept.Name = "Accept";
            this.Accept.Size = new System.Drawing.Size(86, 23);
            this.Accept.TabIndex = 3;
            this.Accept.Text = "Сохранить";
            this.Accept.UseVisualStyleBackColor = true;
            // 
            // Cancel
            // 
            this.Cancel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.Cancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.Cancel.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.Cancel.Location = new System.Drawing.Point(683, 149);
            this.Cancel.Name = "Cancel";
            this.Cancel.Size = new System.Drawing.Size(86, 23);
            this.Cancel.TabIndex = 4;
            this.Cancel.Text = "Отменить";
            this.Cancel.UseVisualStyleBackColor = true;
            // 
            // ChengeDelivery
            // 
            this.AcceptButton = this.Cancel;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.Cancel;
            this.ClientSize = new System.Drawing.Size(781, 174);
            this.Controls.Add(this.Accept);
            this.Controls.Add(this.Cancel);
            this.Controls.Add(this.btnDelete);
            this.Controls.Add(this.lbTruck);
            this.Controls.Add(this.dataGridView1);
            this.Name = "ChengeDelivery";
            this.Text = "ChengeDelivery";
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Label lbTruck;
        private System.Windows.Forms.Button btnDelete;
        private System.Windows.Forms.Button Accept;
        private System.Windows.Forms.Button Cancel;
    }
}