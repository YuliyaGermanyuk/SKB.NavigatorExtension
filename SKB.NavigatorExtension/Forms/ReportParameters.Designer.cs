namespace SKB.NavigatorExtension.Forms
{
    partial class ReportParameters
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ReportParameters));
            this.StartDate = new DevExpress.XtraEditors.DateEdit();
            this.StartDateLabel = new DevExpress.XtraEditors.LabelControl();
            this.EndDateLabel = new DevExpress.XtraEditors.LabelControl();
            this.EndDate = new DevExpress.XtraEditors.DateEdit();
            this.MyOKButton = new DevExpress.XtraEditors.SimpleButton();
            this.MyCancelButton = new DevExpress.XtraEditors.SimpleButton();
            ((System.ComponentModel.ISupportInitialize)(this.StartDate.Properties.VistaTimeProperties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.StartDate.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.EndDate.Properties.VistaTimeProperties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.EndDate.Properties)).BeginInit();
            this.SuspendLayout();
            // 
            // StartDate
            // 
            this.StartDate.EditValue = null;
            this.StartDate.Location = new System.Drawing.Point(16, 31);
            this.StartDate.Name = "StartDate";
            this.StartDate.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
            this.StartDate.Properties.VistaTimeProperties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton()});
            this.StartDate.Size = new System.Drawing.Size(170, 20);
            this.StartDate.TabIndex = 0;
            // 
            // StartDateLabel
            // 
            this.StartDateLabel.Location = new System.Drawing.Point(20, 12);
            this.StartDateLabel.Name = "StartDateLabel";
            this.StartDateLabel.Size = new System.Drawing.Size(87, 13);
            this.StartDateLabel.TabIndex = 1;
            this.StartDateLabel.Text = "Начало периода:";
            // 
            // EndDateLabel
            // 
            this.EndDateLabel.Location = new System.Drawing.Point(204, 12);
            this.EndDateLabel.Name = "EndDateLabel";
            this.EndDateLabel.Size = new System.Drawing.Size(81, 13);
            this.EndDateLabel.TabIndex = 2;
            this.EndDateLabel.Text = "Конец периода:";
            // 
            // EndDate
            // 
            this.EndDate.EditValue = null;
            this.EndDate.Location = new System.Drawing.Point(204, 31);
            this.EndDate.Name = "EndDate";
            this.EndDate.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
            this.EndDate.Properties.VistaTimeProperties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton()});
            this.EndDate.Size = new System.Drawing.Size(170, 20);
            this.EndDate.TabIndex = 3;
            // 
            // MyOKButton
            // 
            this.MyOKButton.Location = new System.Drawing.Point(204, 68);
            this.MyOKButton.Name = "MyOKButton";
            this.MyOKButton.Size = new System.Drawing.Size(80, 23);
            this.MyOKButton.TabIndex = 4;
            this.MyOKButton.Text = "ОК";
            this.MyOKButton.Click += new System.EventHandler(this.OKButton_Click);
            // 
            // MyCancelButton
            // 
            this.MyCancelButton.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.MyCancelButton.Location = new System.Drawing.Point(294, 68);
            this.MyCancelButton.Name = "MyCancelButton";
            this.MyCancelButton.Size = new System.Drawing.Size(80, 23);
            this.MyCancelButton.TabIndex = 5;
            this.MyCancelButton.Text = "Отмена";
            this.MyCancelButton.Click += new System.EventHandler(this.CancelButton_Click);
            // 
            // ReportParameters
            // 
            this.AcceptButton = this.MyOKButton;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(391, 106);
            this.Controls.Add(this.MyCancelButton);
            this.Controls.Add(this.MyOKButton);
            this.Controls.Add(this.EndDate);
            this.Controls.Add(this.EndDateLabel);
            this.Controls.Add(this.StartDateLabel);
            this.Controls.Add(this.StartDate);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MaximumSize = new System.Drawing.Size(407, 144);
            this.MinimizeBox = false;
            this.MinimumSize = new System.Drawing.Size(407, 144);
            this.Name = "ReportParameters";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Параметры отчета";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.ReportParameters_FormClosed);
            ((System.ComponentModel.ISupportInitialize)(this.StartDate.Properties.VistaTimeProperties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.StartDate.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.EndDate.Properties.VistaTimeProperties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.EndDate.Properties)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private DevExpress.XtraEditors.DateEdit StartDate;
        private DevExpress.XtraEditors.LabelControl StartDateLabel;
        private DevExpress.XtraEditors.LabelControl EndDateLabel;
        private DevExpress.XtraEditors.DateEdit EndDate;
        private DevExpress.XtraEditors.SimpleButton MyOKButton;
        private DevExpress.XtraEditors.SimpleButton MyCancelButton;
    }
}