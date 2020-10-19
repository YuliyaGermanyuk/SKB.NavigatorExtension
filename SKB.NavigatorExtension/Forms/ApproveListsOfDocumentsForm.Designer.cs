namespace SKB.NavigatorExtension.Forms
{
    partial class ApproveListsOfDocumentsForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ApproveListsOfDocumentsForm));
            this.Control_Layout = new DevExpress.XtraLayout.LayoutControl();
            this.Button_Cancel = new DevExpress.XtraEditors.SimpleButton();
            this.Button_Start = new DevExpress.XtraEditors.SimpleButton();
            this.Edit_Devices = new RKIT.MyCollectionControl.Design.Control.CollectionControlView();
            this.Group_MainLayout = new DevExpress.XtraLayout.LayoutControlGroup();
            this.Item_Edit_Devices = new DevExpress.XtraLayout.LayoutControlItem();
            this.Item_Button_Start = new DevExpress.XtraLayout.LayoutControlItem();
            this.Item_Button_Cancel = new DevExpress.XtraLayout.LayoutControlItem();
            this.emptySpaceItem1 = new DevExpress.XtraLayout.EmptySpaceItem();
            this.emptySpaceItem2 = new DevExpress.XtraLayout.EmptySpaceItem();
            this.emptySpaceItem3 = new DevExpress.XtraLayout.EmptySpaceItem();
            ((System.ComponentModel.ISupportInitialize)(this.Control_Layout)).BeginInit();
            this.Control_Layout.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.Edit_Devices.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.Group_MainLayout)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.Item_Edit_Devices)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.Item_Button_Start)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.Item_Button_Cancel)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.emptySpaceItem1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.emptySpaceItem2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.emptySpaceItem3)).BeginInit();
            this.SuspendLayout();
            // 
            // Control_Layout
            // 
            this.Control_Layout.AllowCustomizationMenu = false;
            this.Control_Layout.Controls.Add(this.Button_Cancel);
            this.Control_Layout.Controls.Add(this.Button_Start);
            this.Control_Layout.Controls.Add(this.Edit_Devices);
            this.Control_Layout.Dock = System.Windows.Forms.DockStyle.Fill;
            this.Control_Layout.Location = new System.Drawing.Point(0, 0);
            this.Control_Layout.Name = "Control_Layout";
            this.Control_Layout.OptionsCustomizationForm.DesignTimeCustomizationFormPositionAndSize = new System.Drawing.Rectangle(900, 283, 250, 350);
            this.Control_Layout.OptionsView.UseDefaultDragAndDropRendering = false;
            this.Control_Layout.Root = this.Group_MainLayout;
            this.Control_Layout.Size = new System.Drawing.Size(370, 92);
            this.Control_Layout.TabIndex = 0;
            this.Control_Layout.Text = "Control_Layout";
            // 
            // Button_Cancel
            // 
            this.Button_Cancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.Button_Cancel.Location = new System.Drawing.Point(269, 60);
            this.Button_Cancel.Name = "Button_Cancel";
            this.Button_Cancel.Size = new System.Drawing.Size(91, 22);
            this.Button_Cancel.StyleController = this.Control_Layout;
            this.Button_Cancel.TabIndex = 6;
            this.Button_Cancel.Text = "Отмена";
            // 
            // Button_Start
            // 
            this.Button_Start.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.Button_Start.Location = new System.Drawing.Point(163, 60);
            this.Button_Start.Name = "Button_Start";
            this.Button_Start.Size = new System.Drawing.Size(92, 22);
            this.Button_Start.StyleController = this.Control_Layout;
            this.Button_Start.TabIndex = 5;
            this.Button_Start.Text = "Запуск";
            this.Button_Start.Click += new System.EventHandler(this.Button_Click);
            // 
            // Edit_Devices
            // 
            this.Edit_Devices.AllowEdit = true;
            this.Edit_Devices.ControlValue = new System.Guid[0];
            this.Edit_Devices.Location = new System.Drawing.Point(10, 26);
            this.Edit_Devices.Name = "Edit_Devices";
            this.Edit_Devices.Properties.Appearance.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Underline);
            this.Edit_Devices.Properties.Appearance.Options.UseFont = true;
            this.Edit_Devices.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Down, "", -1, true, false, false, DevExpress.XtraEditors.ImageLocation.MiddleCenter, null, new DevExpress.Utils.KeyShortcut(System.Windows.Forms.Keys.None), serializableAppearanceObject3, "", null, null, true),
            new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Delete, "", -1, false, true, false, DevExpress.XtraEditors.ImageLocation.MiddleCenter, null, new DevExpress.Utils.KeyShortcut(System.Windows.Forms.Keys.None), serializableAppearanceObject4, "", null, null, true)});
            this.Edit_Devices.ShowBorder = true;
            this.Edit_Devices.Signed = false;
            this.Edit_Devices.SingleResult = true;
            this.Edit_Devices.Size = new System.Drawing.Size(350, 20);
            this.Edit_Devices.StyleController = this.Control_Layout;
            this.Edit_Devices.TabIndex = 4;
            // 
            // Group_MainLayout
            // 
            this.Group_MainLayout.CustomizationFormText = "Group_MainLayout";
            this.Group_MainLayout.EnableIndentsWithoutBorders = DevExpress.Utils.DefaultBoolean.True;
            this.Group_MainLayout.GroupBordersVisible = false;
            this.Group_MainLayout.Items.AddRange(new DevExpress.XtraLayout.BaseLayoutItem[] {
            this.Item_Edit_Devices,
            this.Item_Button_Start,
            this.Item_Button_Cancel,
            this.emptySpaceItem1,
            this.emptySpaceItem2,
            this.emptySpaceItem3});
            this.Group_MainLayout.Location = new System.Drawing.Point(0, 0);
            this.Group_MainLayout.Name = "Group_MainLayout";
            this.Group_MainLayout.Padding = new DevExpress.XtraLayout.Utils.Padding(8, 8, 8, 8);
            this.Group_MainLayout.Size = new System.Drawing.Size(370, 92);
            this.Group_MainLayout.Text = "Group_MainLayout";
            this.Group_MainLayout.TextVisible = false;
            // 
            // Item_Edit_Devices
            // 
            this.Item_Edit_Devices.Control = this.Edit_Devices;
            this.Item_Edit_Devices.CustomizationFormText = "Прибор:";
            this.Item_Edit_Devices.Location = new System.Drawing.Point(0, 0);
            this.Item_Edit_Devices.Name = "Item_Edit_Parties";
            this.Item_Edit_Devices.Size = new System.Drawing.Size(354, 40);
            this.Item_Edit_Devices.Text = "Прибор:";
            this.Item_Edit_Devices.TextLocation = DevExpress.Utils.Locations.Top;
            this.Item_Edit_Devices.TextSize = new System.Drawing.Size(41, 13);
            // 
            // Item_Button_Start
            // 
            this.Item_Button_Start.Control = this.Button_Start;
            this.Item_Button_Start.CustomizationFormText = "Item_Button_Start";
            this.Item_Button_Start.Location = new System.Drawing.Point(153, 50);
            this.Item_Button_Start.Name = "Item_Button_Start";
            this.Item_Button_Start.Size = new System.Drawing.Size(96, 26);
            this.Item_Button_Start.Text = "Item_Button_Start";
            this.Item_Button_Start.TextSize = new System.Drawing.Size(0, 0);
            this.Item_Button_Start.TextToControlDistance = 0;
            this.Item_Button_Start.TextVisible = false;
            // 
            // Item_Button_Cancel
            // 
            this.Item_Button_Cancel.Control = this.Button_Cancel;
            this.Item_Button_Cancel.CustomizationFormText = "Item_Button_Cancel";
            this.Item_Button_Cancel.Location = new System.Drawing.Point(259, 50);
            this.Item_Button_Cancel.Name = "Item_Button_Cancel";
            this.Item_Button_Cancel.Size = new System.Drawing.Size(95, 26);
            this.Item_Button_Cancel.Text = "Item_Button_Cancel";
            this.Item_Button_Cancel.TextSize = new System.Drawing.Size(0, 0);
            this.Item_Button_Cancel.TextToControlDistance = 0;
            this.Item_Button_Cancel.TextVisible = false;
            // 
            // emptySpaceItem1
            // 
            this.emptySpaceItem1.AllowHotTrack = false;
            this.emptySpaceItem1.CustomizationFormText = "emptySpaceItem1";
            this.emptySpaceItem1.Location = new System.Drawing.Point(0, 40);
            this.emptySpaceItem1.MaxSize = new System.Drawing.Size(0, 10);
            this.emptySpaceItem1.MinSize = new System.Drawing.Size(10, 10);
            this.emptySpaceItem1.Name = "emptySpaceItem1";
            this.emptySpaceItem1.Size = new System.Drawing.Size(354, 10);
            this.emptySpaceItem1.SizeConstraintsType = DevExpress.XtraLayout.SizeConstraintsType.Custom;
            this.emptySpaceItem1.Text = "emptySpaceItem1";
            this.emptySpaceItem1.TextSize = new System.Drawing.Size(0, 0);
            // 
            // emptySpaceItem2
            // 
            this.emptySpaceItem2.AllowHotTrack = false;
            this.emptySpaceItem2.CustomizationFormText = "emptySpaceItem2";
            this.emptySpaceItem2.Location = new System.Drawing.Point(249, 50);
            this.emptySpaceItem2.MaxSize = new System.Drawing.Size(10, 0);
            this.emptySpaceItem2.MinSize = new System.Drawing.Size(10, 10);
            this.emptySpaceItem2.Name = "emptySpaceItem2";
            this.emptySpaceItem2.Size = new System.Drawing.Size(10, 26);
            this.emptySpaceItem2.SizeConstraintsType = DevExpress.XtraLayout.SizeConstraintsType.Custom;
            this.emptySpaceItem2.Text = "emptySpaceItem2";
            this.emptySpaceItem2.TextSize = new System.Drawing.Size(0, 0);
            // 
            // emptySpaceItem3
            // 
            this.emptySpaceItem3.AllowHotTrack = false;
            this.emptySpaceItem3.CustomizationFormText = "emptySpaceItem3";
            this.emptySpaceItem3.Location = new System.Drawing.Point(0, 50);
            this.emptySpaceItem3.Name = "emptySpaceItem3";
            this.emptySpaceItem3.Size = new System.Drawing.Size(153, 26);
            this.emptySpaceItem3.Text = "emptySpaceItem3";
            this.emptySpaceItem3.TextSize = new System.Drawing.Size(0, 0);
            // 
            // ApproveListsOfDocumentsForm
            // 
            this.AcceptButton = this.Button_Start;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.Button_Cancel;
            this.ClientSize = new System.Drawing.Size(370, 92);
            this.Controls.Add(this.Control_Layout);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "ApproveListsOfDocumentsForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Утверждение документации";
            ((System.ComponentModel.ISupportInitialize)(this.Control_Layout)).EndInit();
            this.Control_Layout.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.Edit_Devices.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.Group_MainLayout)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.Item_Edit_Devices)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.Item_Button_Start)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.Item_Button_Cancel)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.emptySpaceItem1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.emptySpaceItem2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.emptySpaceItem3)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private DevExpress.XtraLayout.LayoutControl Control_Layout;
        private DevExpress.XtraLayout.LayoutControlGroup Group_MainLayout;
        private DevExpress.XtraEditors.SimpleButton Button_Cancel;
        private DevExpress.XtraEditors.SimpleButton Button_Start;
        private RKIT.MyCollectionControl.Design.Control.CollectionControlView Edit_Devices;
        private DevExpress.XtraLayout.LayoutControlItem Item_Edit_Devices;
        private DevExpress.XtraLayout.LayoutControlItem Item_Button_Start;
        private DevExpress.XtraLayout.LayoutControlItem Item_Button_Cancel;
        private DevExpress.XtraLayout.EmptySpaceItem emptySpaceItem1;
        private DevExpress.XtraLayout.EmptySpaceItem emptySpaceItem2;
        private DevExpress.XtraLayout.EmptySpaceItem emptySpaceItem3;
    }
}