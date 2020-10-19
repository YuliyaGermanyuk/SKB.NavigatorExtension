namespace SKB.NavigatorExtension.Forms
{
    partial class CompleteReportParameters
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose (bool disposing)
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
        private void InitializeComponent ()
        {
            DevExpress.Utils.SerializableAppearanceObject serializableAppearanceObject3 = new DevExpress.Utils.SerializableAppearanceObject();
            DevExpress.Utils.SerializableAppearanceObject serializableAppearanceObject4 = new DevExpress.Utils.SerializableAppearanceObject();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(CompleteReportParameters));
            DevExpress.Utils.SerializableAppearanceObject serializableAppearanceObject1 = new DevExpress.Utils.SerializableAppearanceObject();
            DevExpress.Utils.SerializableAppearanceObject serializableAppearanceObject2 = new DevExpress.Utils.SerializableAppearanceObject();
            this.Control_Layout = new DevExpress.XtraLayout.LayoutControl();
            this.Button_Cancel = new DevExpress.XtraEditors.SimpleButton();
            this.Button_OK = new DevExpress.XtraEditors.SimpleButton();
            this.Edit_Devices = new RKIT.MyCollectionControl.Design.Control.CollectionControlView();
            this.Group_MainLayout = new DevExpress.XtraLayout.LayoutControlGroup();
            this.Item_Edit_Devices = new DevExpress.XtraLayout.LayoutControlItem();
            this.Item_Button_OK = new DevExpress.XtraLayout.LayoutControlItem();
            this.Item_Button_Cancel = new DevExpress.XtraLayout.LayoutControlItem();
            this.emptySpaceItem1 = new DevExpress.XtraLayout.EmptySpaceItem();
            this.emptySpaceItem2 = new DevExpress.XtraLayout.EmptySpaceItem();
            this.emptySpaceItem3 = new DevExpress.XtraLayout.EmptySpaceItem();
            this.AllDevices = new System.Windows.Forms.CheckBox();
            this.layoutControlItem1 = new DevExpress.XtraLayout.LayoutControlItem();
            this.ChooseDevices = new System.Windows.Forms.CheckBox();
            this.layoutControlItem2 = new DevExpress.XtraLayout.LayoutControlItem();
            this.ChooseCompletes = new System.Windows.Forms.CheckBox();
            this.layoutControlItem3 = new DevExpress.XtraLayout.LayoutControlItem();
            this.Edit_Completes = new RKIT.MyCollectionControl.Design.Control.CollectionControlView();
            this.Item_Edit_Completes = new DevExpress.XtraLayout.LayoutControlItem();
            ((System.ComponentModel.ISupportInitialize)(this.Control_Layout)).BeginInit();
            this.Control_Layout.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.Edit_Devices.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.Group_MainLayout)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.Item_Edit_Devices)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.Item_Button_OK)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.Item_Button_Cancel)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.emptySpaceItem1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.emptySpaceItem2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.emptySpaceItem3)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem3)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.Edit_Completes.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.Item_Edit_Completes)).BeginInit();
            this.SuspendLayout();
            // 
            // Control_Layout
            // 
            this.Control_Layout.AllowCustomizationMenu = false;
            this.Control_Layout.Controls.Add(this.Edit_Completes);
            this.Control_Layout.Controls.Add(this.ChooseCompletes);
            this.Control_Layout.Controls.Add(this.ChooseDevices);
            this.Control_Layout.Controls.Add(this.AllDevices);
            this.Control_Layout.Controls.Add(this.Button_Cancel);
            this.Control_Layout.Controls.Add(this.Button_OK);
            this.Control_Layout.Controls.Add(this.Edit_Devices);
            this.Control_Layout.Dock = System.Windows.Forms.DockStyle.Fill;
            this.Control_Layout.Location = new System.Drawing.Point(0, 0);
            this.Control_Layout.Name = "Control_Layout";
            this.Control_Layout.OptionsCustomizationForm.DesignTimeCustomizationFormPositionAndSize = new System.Drawing.Rectangle(900, 283, 250, 350);
            this.Control_Layout.OptionsView.UseDefaultDragAndDropRendering = false;
            this.Control_Layout.Root = this.Group_MainLayout;
            this.Control_Layout.Size = new System.Drawing.Size(549, 208);
            this.Control_Layout.TabIndex = 0;
            this.Control_Layout.Text = "layoutControl1";
            // 
            // Button_Cancel
            // 
            this.Button_Cancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.Button_Cancel.Location = new System.Drawing.Point(425, 172);
            this.Button_Cancel.Name = "Button_Cancel";
            this.Button_Cancel.Size = new System.Drawing.Size(114, 22);
            this.Button_Cancel.StyleController = this.Control_Layout;
            this.Button_Cancel.TabIndex = 6;
            this.Button_Cancel.Text = "Отмена";
            // 
            // Button_OK
            // 
            this.Button_OK.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.Button_OK.Location = new System.Drawing.Point(297, 172);
            this.Button_OK.Name = "Button_OK";
            this.Button_OK.Size = new System.Drawing.Size(114, 22);
            this.Button_OK.StyleController = this.Control_Layout;
            this.Button_OK.TabIndex = 5;
            this.Button_OK.Text = "ОК";
            this.Button_OK.Click += new System.EventHandler(this.Button_Click);
            // 
            // Edit_Devices
            // 
            this.Edit_Devices.AllowEdit = true;
            this.Edit_Devices.ColumnsCollection = ((System.Collections.Generic.List<RKIT.MyCollectionControl.CustomTableColumn>)(resources.GetObject("Edit_Devices.ColumnsCollection")));
            this.Edit_Devices.ControlValue = new System.Guid[0];
            this.Edit_Devices.Location = new System.Drawing.Point(10, 74);
            this.Edit_Devices.Name = "Edit_Devices";
            this.Edit_Devices.Properties.Appearance.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Underline);
            this.Edit_Devices.Properties.Appearance.Options.UseFont = true;
            this.Edit_Devices.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Ellipsis, "", -1, true, false, false, DevExpress.XtraEditors.ImageLocation.MiddleCenter, null, new DevExpress.Utils.KeyShortcut(System.Windows.Forms.Keys.None), serializableAppearanceObject3, "", null, null, true),
            new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Down),
            new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Delete, "", -1, false, true, false, DevExpress.XtraEditors.ImageLocation.MiddleCenter, null, new DevExpress.Utils.KeyShortcut(System.Windows.Forms.Keys.None), serializableAppearanceObject4, "", null, null, true)});
            this.Edit_Devices.ShowBorder = true;
            this.Edit_Devices.Signed = false;
            this.Edit_Devices.SingleResult = false;
            this.Edit_Devices.Size = new System.Drawing.Size(529, 20);
            this.Edit_Devices.StyleController = this.Control_Layout;
            this.Edit_Devices.TabIndex = 4;
            this.Edit_Devices.ToolTipSettings = ((System.Collections.Generic.List<RKIT.MyCollectionControl.CustomTableColumn>)(resources.GetObject("Edit_Devices.ToolTipSettings")));
            this.Edit_Devices.TypeIds = null;
            // 
            // Group_MainLayout
            // 
            this.Group_MainLayout.CustomizationFormText = "Group_MainLayout";
            this.Group_MainLayout.EnableIndentsWithoutBorders = DevExpress.Utils.DefaultBoolean.True;
            this.Group_MainLayout.GroupBordersVisible = false;
            this.Group_MainLayout.Items.AddRange(new DevExpress.XtraLayout.BaseLayoutItem[] {
            this.Item_Button_OK,
            this.Item_Button_Cancel,
            this.emptySpaceItem1,
            this.emptySpaceItem2,
            this.emptySpaceItem3,
            this.layoutControlItem1,
            this.layoutControlItem2,
            this.layoutControlItem3,
            this.Item_Edit_Devices,
            this.Item_Edit_Completes});
            this.Group_MainLayout.Location = new System.Drawing.Point(0, 0);
            this.Group_MainLayout.Name = "Group_MainLayout";
            this.Group_MainLayout.Padding = new DevExpress.XtraLayout.Utils.Padding(8, 8, 8, 8);
            this.Group_MainLayout.Size = new System.Drawing.Size(549, 208);
            this.Group_MainLayout.Text = "Group_MainLayout";
            this.Group_MainLayout.TextVisible = false;
            // 
            // Item_Edit_Devices
            // 
            this.Item_Edit_Devices.Control = this.Edit_Devices;
            this.Item_Edit_Devices.CustomizationFormText = "Приборы:";
            this.Item_Edit_Devices.Location = new System.Drawing.Point(0, 48);
            this.Item_Edit_Devices.Name = "Item_Edit_Devices";
            this.Item_Edit_Devices.Size = new System.Drawing.Size(533, 40);
            this.Item_Edit_Devices.Text = "Приборы:";
            this.Item_Edit_Devices.TextLocation = DevExpress.Utils.Locations.Top;
            this.Item_Edit_Devices.TextSize = new System.Drawing.Size(89, 13);
            // 
            // Item_Button_OK
            // 
            this.Item_Button_OK.Control = this.Button_OK;
            this.Item_Button_OK.CustomizationFormText = "Item_Button_Start";
            this.Item_Button_OK.Location = new System.Drawing.Point(287, 162);
            this.Item_Button_OK.Name = "Item_Button_Start";
            this.Item_Button_OK.Size = new System.Drawing.Size(118, 30);
            this.Item_Button_OK.Text = "Item_Button_Start";
            this.Item_Button_OK.TextSize = new System.Drawing.Size(0, 0);
            this.Item_Button_OK.TextToControlDistance = 0;
            this.Item_Button_OK.TextVisible = false;
            // 
            // Item_Button_Cancel
            // 
            this.Item_Button_Cancel.Control = this.Button_Cancel;
            this.Item_Button_Cancel.CustomizationFormText = "Item_Button_Cancel";
            this.Item_Button_Cancel.Location = new System.Drawing.Point(415, 162);
            this.Item_Button_Cancel.Name = "Item_Button_Cancel";
            this.Item_Button_Cancel.Size = new System.Drawing.Size(118, 30);
            this.Item_Button_Cancel.Text = "Item_Button_Cancel";
            this.Item_Button_Cancel.TextSize = new System.Drawing.Size(0, 0);
            this.Item_Button_Cancel.TextToControlDistance = 0;
            this.Item_Button_Cancel.TextVisible = false;
            // 
            // emptySpaceItem1
            // 
            this.emptySpaceItem1.AllowHotTrack = false;
            this.emptySpaceItem1.CustomizationFormText = "emptySpaceItem1";
            this.emptySpaceItem1.Location = new System.Drawing.Point(0, 152);
            this.emptySpaceItem1.MaxSize = new System.Drawing.Size(0, 10);
            this.emptySpaceItem1.MinSize = new System.Drawing.Size(10, 10);
            this.emptySpaceItem1.Name = "emptySpaceItem1";
            this.emptySpaceItem1.Size = new System.Drawing.Size(533, 10);
            this.emptySpaceItem1.SizeConstraintsType = DevExpress.XtraLayout.SizeConstraintsType.Custom;
            this.emptySpaceItem1.Text = "emptySpaceItem1";
            this.emptySpaceItem1.TextSize = new System.Drawing.Size(0, 0);
            // 
            // emptySpaceItem2
            // 
            this.emptySpaceItem2.AllowHotTrack = false;
            this.emptySpaceItem2.CustomizationFormText = "emptySpaceItem2";
            this.emptySpaceItem2.Location = new System.Drawing.Point(405, 162);
            this.emptySpaceItem2.MaxSize = new System.Drawing.Size(10, 0);
            this.emptySpaceItem2.MinSize = new System.Drawing.Size(10, 10);
            this.emptySpaceItem2.Name = "emptySpaceItem2";
            this.emptySpaceItem2.Size = new System.Drawing.Size(10, 30);
            this.emptySpaceItem2.SizeConstraintsType = DevExpress.XtraLayout.SizeConstraintsType.Custom;
            this.emptySpaceItem2.Text = "emptySpaceItem2";
            this.emptySpaceItem2.TextSize = new System.Drawing.Size(0, 0);
            // 
            // emptySpaceItem3
            // 
            this.emptySpaceItem3.AllowHotTrack = false;
            this.emptySpaceItem3.CustomizationFormText = "emptySpaceItem3";
            this.emptySpaceItem3.Location = new System.Drawing.Point(0, 162);
            this.emptySpaceItem3.Name = "emptySpaceItem3";
            this.emptySpaceItem3.Size = new System.Drawing.Size(287, 30);
            this.emptySpaceItem3.Text = "emptySpaceItem3";
            this.emptySpaceItem3.TextSize = new System.Drawing.Size(0, 0);
            // 
            // AllDevices
            // 
            this.AllDevices.Location = new System.Drawing.Point(10, 10);
            this.AllDevices.Name = "AllDevices";
            this.AllDevices.Size = new System.Drawing.Size(529, 20);
            this.AllDevices.TabIndex = 7;
            this.AllDevices.Text = "Все приборы";
            this.AllDevices.UseVisualStyleBackColor = true;
            this.AllDevices.CheckedChanged += new System.EventHandler(this.AllDevices_CheckedChanged);
            // 
            // layoutControlItem1
            // 
            this.layoutControlItem1.Control = this.AllDevices;
            this.layoutControlItem1.CustomizationFormText = "layoutControlItem1";
            this.layoutControlItem1.Location = new System.Drawing.Point(0, 0);
            this.layoutControlItem1.Name = "layoutControlItem1";
            this.layoutControlItem1.Size = new System.Drawing.Size(533, 24);
            this.layoutControlItem1.Text = "layoutControlItem1";
            this.layoutControlItem1.TextSize = new System.Drawing.Size(0, 0);
            this.layoutControlItem1.TextToControlDistance = 0;
            this.layoutControlItem1.TextVisible = false;
            // 
            // ChooseDevices
            // 
            this.ChooseDevices.Location = new System.Drawing.Point(10, 34);
            this.ChooseDevices.Name = "ChooseDevices";
            this.ChooseDevices.Size = new System.Drawing.Size(529, 20);
            this.ChooseDevices.TabIndex = 8;
            this.ChooseDevices.Text = "Выбрать приборы:";
            this.ChooseDevices.UseVisualStyleBackColor = true;
            this.ChooseDevices.CheckedChanged += new System.EventHandler(this.ChooseDevices_CheckedChanged);
            // 
            // layoutControlItem2
            // 
            this.layoutControlItem2.Control = this.ChooseDevices;
            this.layoutControlItem2.CustomizationFormText = "layoutControlItem2";
            this.layoutControlItem2.Location = new System.Drawing.Point(0, 24);
            this.layoutControlItem2.Name = "layoutControlItem2";
            this.layoutControlItem2.Size = new System.Drawing.Size(533, 24);
            this.layoutControlItem2.Text = "layoutControlItem2";
            this.layoutControlItem2.TextSize = new System.Drawing.Size(0, 0);
            this.layoutControlItem2.TextToControlDistance = 0;
            this.layoutControlItem2.TextVisible = false;
            // 
            // ChooseCompletes
            // 
            this.ChooseCompletes.Location = new System.Drawing.Point(10, 98);
            this.ChooseCompletes.Name = "ChooseCompletes";
            this.ChooseCompletes.Size = new System.Drawing.Size(529, 20);
            this.ChooseCompletes.TabIndex = 9;
            this.ChooseCompletes.Text = "Выбрать комплектующие:";
            this.ChooseCompletes.UseVisualStyleBackColor = true;
            this.ChooseCompletes.CheckedChanged += new System.EventHandler(this.ChooseCompletes_CheckedChanged);
            // 
            // layoutControlItem3
            // 
            this.layoutControlItem3.Control = this.ChooseCompletes;
            this.layoutControlItem3.CustomizationFormText = "layoutControlItem3";
            this.layoutControlItem3.Location = new System.Drawing.Point(0, 88);
            this.layoutControlItem3.Name = "layoutControlItem3";
            this.layoutControlItem3.Size = new System.Drawing.Size(533, 24);
            this.layoutControlItem3.Text = "layoutControlItem3";
            this.layoutControlItem3.TextSize = new System.Drawing.Size(0, 0);
            this.layoutControlItem3.TextToControlDistance = 0;
            this.layoutControlItem3.TextVisible = false;
            // 
            // Edit_Completes
            // 
            this.Edit_Completes.AllowEdit = true;
            this.Edit_Completes.ColumnsCollection = ((System.Collections.Generic.List<RKIT.MyCollectionControl.CustomTableColumn>)(resources.GetObject("Edit_Completes.ColumnsCollection")));
            this.Edit_Completes.ControlValue = new System.Guid[0];
            this.Edit_Completes.Location = new System.Drawing.Point(10, 138);
            this.Edit_Completes.Name = "Edit_Completes";
            this.Edit_Completes.Properties.Appearance.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Underline);
            this.Edit_Completes.Properties.Appearance.Options.UseFont = true;
            this.Edit_Completes.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Ellipsis, "", -1, true, false, false, DevExpress.XtraEditors.ImageLocation.MiddleCenter, null, new DevExpress.Utils.KeyShortcut(System.Windows.Forms.Keys.None), serializableAppearanceObject1, "", null, null, true),
            new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Down),
            new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Delete, "", -1, false, true, false, DevExpress.XtraEditors.ImageLocation.MiddleCenter, null, new DevExpress.Utils.KeyShortcut(System.Windows.Forms.Keys.None), serializableAppearanceObject2, "", null, null, true)});
            this.Edit_Completes.ShowBorder = true;
            this.Edit_Completes.Signed = false;
            this.Edit_Completes.SingleResult = false;
            this.Edit_Completes.Size = new System.Drawing.Size(529, 20);
            this.Edit_Completes.StyleController = this.Control_Layout;
            this.Edit_Completes.TabIndex = 5;
            this.Edit_Completes.ToolTipSettings = ((System.Collections.Generic.List<RKIT.MyCollectionControl.CustomTableColumn>)(resources.GetObject("Edit_Completes.ToolTipSettings")));
            this.Edit_Completes.TypeIds = null;
            // 
            // Item_Edit_Completes
            // 
            this.Item_Edit_Completes.Control = this.Edit_Completes;
            this.Item_Edit_Completes.CustomizationFormText = "Комплектующие:";
            this.Item_Edit_Completes.Location = new System.Drawing.Point(0, 112);
            this.Item_Edit_Completes.Name = "Item_Edit_Completes";
            this.Item_Edit_Completes.Size = new System.Drawing.Size(533, 40);
            this.Item_Edit_Completes.Text = "Комплектующие:";
            this.Item_Edit_Completes.TextLocation = DevExpress.Utils.Locations.Top;
            this.Item_Edit_Completes.TextSize = new System.Drawing.Size(89, 13);
            // 
            // CompleteReportParameters
            // 
            this.AcceptButton = this.Button_OK;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.Button_Cancel;
            this.ClientSize = new System.Drawing.Size(549, 208);
            this.Controls.Add(this.Control_Layout);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "CompleteReportParameters";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Выберите параметры отчета";
            ((System.ComponentModel.ISupportInitialize)(this.Control_Layout)).EndInit();
            this.Control_Layout.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.Edit_Devices.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.Group_MainLayout)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.Item_Edit_Devices)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.Item_Button_OK)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.Item_Button_Cancel)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.emptySpaceItem1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.emptySpaceItem2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.emptySpaceItem3)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem3)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.Edit_Completes.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.Item_Edit_Completes)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private DevExpress.XtraLayout.LayoutControl Control_Layout;
        private DevExpress.XtraLayout.LayoutControlGroup Group_MainLayout;
        private DevExpress.XtraEditors.SimpleButton Button_Cancel;
        private DevExpress.XtraEditors.SimpleButton Button_OK;
        private RKIT.MyCollectionControl.Design.Control.CollectionControlView Edit_Devices;
        private DevExpress.XtraLayout.LayoutControlItem Item_Edit_Devices;
        private DevExpress.XtraLayout.LayoutControlItem Item_Button_OK;
        private DevExpress.XtraLayout.LayoutControlItem Item_Button_Cancel;
        private DevExpress.XtraLayout.EmptySpaceItem emptySpaceItem1;
        private DevExpress.XtraLayout.EmptySpaceItem emptySpaceItem2;
        private DevExpress.XtraLayout.EmptySpaceItem emptySpaceItem3;
        private System.Windows.Forms.CheckBox ChooseCompletes;
        private System.Windows.Forms.CheckBox ChooseDevices;
        private System.Windows.Forms.CheckBox AllDevices;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem1;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem2;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem3;
        private RKIT.MyCollectionControl.Design.Control.CollectionControlView Edit_Completes;
        private DevExpress.XtraLayout.LayoutControlItem Item_Edit_Completes;
    }
}