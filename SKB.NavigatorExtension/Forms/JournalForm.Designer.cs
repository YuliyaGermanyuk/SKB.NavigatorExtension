namespace SKB.NavigatorExtension.Forms
{
    partial class JournalForm
    {
        /// <summary>
        /// Требуется переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором форм Windows

        /// <summary>
        /// Обязательный метод для поддержки конструктора - не изменяйте
        /// содержимое данного метода при помощи редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(JournalForm));
            this.Date = new DevExpress.XtraEditors.DateEdit();
            this.Temperature = new DevExpress.XtraEditors.SpinEdit();
            this.Humidity = new DevExpress.XtraEditors.SpinEdit();
            this.Pressure = new DevExpress.XtraEditors.SpinEdit();
            this.Employee = new DevExpress.XtraEditors.ButtonEdit();
            this.DateLable = new DevExpress.XtraEditors.LabelControl();
            this.EmployeeLable = new DevExpress.XtraEditors.LabelControl();
            this.TemperatureLable = new DevExpress.XtraEditors.LabelControl();
            this.HumidityLable = new DevExpress.XtraEditors.LabelControl();
            this.PressureLable = new DevExpress.XtraEditors.LabelControl();
            this.groupControl1 = new DevExpress.XtraEditors.GroupControl();
            this.OKButton = new DevExpress.XtraEditors.SimpleButton();
            this.toolTipController = new DevExpress.Utils.ToolTipController(this.components);
            ((System.ComponentModel.ISupportInitialize)(this.Date.Properties.VistaTimeProperties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.Date.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.Temperature.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.Humidity.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.Pressure.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.Employee.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.groupControl1)).BeginInit();
            this.groupControl1.SuspendLayout();
            this.SuspendLayout();
            // 
            // Date
            // 
            this.Date.EditValue = null;
            this.Date.Enabled = false;
            this.Date.Location = new System.Drawing.Point(24, 34);
            this.Date.Name = "Date";
            this.Date.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
            this.Date.Properties.VistaTimeProperties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton()});
            this.Date.Size = new System.Drawing.Size(117, 20);
            this.Date.TabIndex = 0;
            this.Date.EditValueChanged += new System.EventHandler(this.dateEdit1_EditValueChanged);
            // 
            // Temperature
            // 
            this.Temperature.EditValue = new decimal(new int[] {
            0,
            0,
            0,
            0});
            this.Temperature.Location = new System.Drawing.Point(244, 38);
            this.Temperature.Name = "Temperature";
            this.Temperature.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton()});
            this.Temperature.Size = new System.Drawing.Size(132, 20);
            this.Temperature.TabIndex = 1;
            // 
            // Humidity
            // 
            this.Humidity.EditValue = new decimal(new int[] {
            0,
            0,
            0,
            0});
            this.Humidity.Location = new System.Drawing.Point(244, 71);
            this.Humidity.Name = "Humidity";
            this.Humidity.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton()});
            this.Humidity.Size = new System.Drawing.Size(132, 20);
            this.Humidity.TabIndex = 2;
            // 
            // Pressure
            // 
            this.Pressure.EditValue = new decimal(new int[] {
            0,
            0,
            0,
            0});
            this.Pressure.Location = new System.Drawing.Point(244, 104);
            this.Pressure.Name = "Pressure";
            this.Pressure.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton()});
            this.Pressure.Size = new System.Drawing.Size(132, 20);
            this.Pressure.TabIndex = 3;
            // 
            // Employee
            // 
            this.Employee.Enabled = false;
            this.Employee.Location = new System.Drawing.Point(167, 34);
            this.Employee.Name = "Employee";
            this.Employee.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton()});
            this.Employee.Size = new System.Drawing.Size(244, 20);
            this.Employee.TabIndex = 4;
            // 
            // DateLable
            // 
            this.DateLable.Appearance.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.DateLable.Location = new System.Drawing.Point(24, 14);
            this.DateLable.Name = "DateLable";
            this.DateLable.Size = new System.Drawing.Size(101, 16);
            this.DateLable.TabIndex = 5;
            this.DateLable.Text = "Дата измерения:";
            // 
            // EmployeeLable
            // 
            this.EmployeeLable.Appearance.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.EmployeeLable.Location = new System.Drawing.Point(167, 14);
            this.EmployeeLable.Name = "EmployeeLable";
            this.EmployeeLable.Size = new System.Drawing.Size(221, 16);
            this.EmployeeLable.TabIndex = 6;
            this.EmployeeLable.Text = "Сотрудник, проводивший измерения:";
            // 
            // TemperatureLable
            // 
            this.TemperatureLable.Appearance.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.TemperatureLable.Location = new System.Drawing.Point(11, 39);
            this.TemperatureLable.Name = "TemperatureLable";
            this.TemperatureLable.Size = new System.Drawing.Size(124, 16);
            this.TemperatureLable.TabIndex = 7;
            this.TemperatureLable.Text = "Температура (гр. С):";
            // 
            // HumidityLable
            // 
            this.HumidityLable.Appearance.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.HumidityLable.Location = new System.Drawing.Point(11, 72);
            this.HumidityLable.Name = "HumidityLable";
            this.HumidityLable.Size = new System.Drawing.Size(189, 16);
            this.HumidityLable.TabIndex = 8;
            this.HumidityLable.Text = "Относительная влажность (%):";
            // 
            // PressureLable
            // 
            this.PressureLable.Appearance.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.PressureLable.Location = new System.Drawing.Point(11, 105);
            this.PressureLable.Name = "PressureLable";
            this.PressureLable.Size = new System.Drawing.Size(220, 16);
            this.PressureLable.TabIndex = 9;
            this.PressureLable.Text = "Атмосферное давление (мм. рт. ст.):";
            // 
            // groupControl1
            // 
            this.groupControl1.AppearanceCaption.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.groupControl1.AppearanceCaption.Options.UseFont = true;
            this.groupControl1.Controls.Add(this.Pressure);
            this.groupControl1.Controls.Add(this.PressureLable);
            this.groupControl1.Controls.Add(this.Humidity);
            this.groupControl1.Controls.Add(this.TemperatureLable);
            this.groupControl1.Controls.Add(this.HumidityLable);
            this.groupControl1.Controls.Add(this.Temperature);
            this.groupControl1.Location = new System.Drawing.Point(24, 72);
            this.groupControl1.Name = "groupControl1";
            this.groupControl1.Size = new System.Drawing.Size(386, 139);
            this.groupControl1.TabIndex = 11;
            this.groupControl1.Text = "Заполните данные о текущих условиях калибровки:";
            // 
            // OKButton
            // 
            this.OKButton.Location = new System.Drawing.Point(24, 229);
            this.OKButton.Name = "OKButton";
            this.OKButton.Size = new System.Drawing.Size(387, 23);
            this.OKButton.TabIndex = 12;
            this.OKButton.Text = "ОК";
            this.OKButton.Click += new System.EventHandler(this.OKButton_Click);
            // 
            // JournalForm
            // 
            this.AcceptButton = this.OKButton;
            this.Appearance.Font = new System.Drawing.Font("Verdana", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.Appearance.Options.UseFont = true;
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSize = true;
            this.ClientSize = new System.Drawing.Size(437, 264);
            this.ControlBox = false;
            this.Controls.Add(this.OKButton);
            this.Controls.Add(this.groupControl1);
            this.Controls.Add(this.EmployeeLable);
            this.Controls.Add(this.DateLable);
            this.Controls.Add(this.Employee);
            this.Controls.Add(this.Date);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.LookAndFeel.SkinName = "Office 2010 Blue";
            this.LookAndFeel.UseDefaultLookAndFeel = false;
            this.MaximizeBox = false;
            this.MaximumSize = new System.Drawing.Size(453, 302);
            this.MinimizeBox = false;
            this.MinimumSize = new System.Drawing.Size(453, 302);
            this.Name = "JournalForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Журнал условий калибровки";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.JournalForm_FormClosing);
            this.Load += new System.EventHandler(this.JournalForm_Load);
            ((System.ComponentModel.ISupportInitialize)(this.Date.Properties.VistaTimeProperties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.Date.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.Temperature.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.Humidity.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.Pressure.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.Employee.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.groupControl1)).EndInit();
            this.groupControl1.ResumeLayout(false);
            this.groupControl1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private DevExpress.XtraEditors.DateEdit Date;
        private DevExpress.XtraEditors.SpinEdit Temperature;
        private DevExpress.XtraEditors.SpinEdit Humidity;
        private DevExpress.XtraEditors.SpinEdit Pressure;
        private DevExpress.XtraEditors.ButtonEdit Employee;
        private DevExpress.XtraEditors.LabelControl DateLable;
        private DevExpress.XtraEditors.LabelControl EmployeeLable;
        private DevExpress.XtraEditors.LabelControl TemperatureLable;
        private DevExpress.XtraEditors.LabelControl HumidityLable;
        private DevExpress.XtraEditors.LabelControl PressureLable;
        private DevExpress.XtraEditors.GroupControl groupControl1;
        private DevExpress.XtraEditors.SimpleButton OKButton;
        private DevExpress.Utils.ToolTipController toolTipController;
    }
}