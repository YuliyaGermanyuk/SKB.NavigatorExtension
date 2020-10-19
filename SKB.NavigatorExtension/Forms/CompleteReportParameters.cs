using DevExpress.XtraEditors;
using DocsVision.Platform.ObjectManager;
using DocsVision.Platform.ObjectManager.Metadata;
using DocsVision.Platform.ObjectManager.SearchModel;
using DocsVision.Platform.ObjectModel;
using RKIT.MyCollectionControl.Design.Layout;
using RKIT.MyMessageBox;
using SKB.Base;
using SKB.Base.Dictionary;
using SKB.Base.Ref.Properties;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using CardOrd = DocsVision.TakeOffice.Cards.Constants.CardOrd;
using RefUniversal = DocsVision.TakeOffice.Cards.Constants.RefUniversal;

namespace SKB.NavigatorExtension.Forms
{
    /// <summary>
    /// Форма выбора параметров отчета по остаткам комплектующих.
    /// </summary>
    public partial class CompleteReportParameters : XtraForm
    {
        ObjectContext Context;
        UserSession Session;
        /// <summary>
        /// Выбраны все приборы.
        /// </summary>
        public bool allDevices
        {
            get
            {
                return AllDevices.Checked;
            }
        }
        /// <summary>
        /// Выбран перечень приборов.
        /// </summary>
        public bool chooseDevices
        {
            get
            {
                return ChooseDevices.Checked;
            }
        }
        /// <summary>
        /// Выбран перечень комплектующих.
        /// </summary>
        public bool chooseCompletes
        {
            get
            {
                return ChooseCompletes.Checked;
            }
        }
        /// <summary>
        /// Выбранные приборы.
        /// </summary>
        public List<SelectionItem> Devices
        {
            get
            {
                return Edit_Devices.SelectedItems.Select(item => new SelectionItem(item.ObjectId, item.DisplayValue)).ToList();
            }
        }
        /// <summary>
        /// Выбранные комплектующие.
        /// </summary>
        public List<SelectionItem> Completes
        {
            get
            {
                return Edit_Completes.SelectedItems.Select(item => new SelectionItem(item.ObjectId, item.DisplayValue)).ToList();
            }
        }
        /// <summary>
        /// Инициализирует форму выбора параметров отчета по остаткам комплектующих.
        /// </summary>
        /// <param name="Session">Пользовательская сессия.</param>
        /// <param name="Context">Объектный контекст.</param>
        public CompleteReportParameters(UserSession Session, ObjectContext Context)
        {
            InitializeComponent();

            this.SetNavigatorSkin();

            this.Context = Context;
            this.Session = Session;

            this.Edit_Devices.Session = Session;
            this.Edit_Devices.ObjectContext = Context;
            this.Edit_Devices.TypeIds = new List<MyCollectionControlType>()
                {
                    new MyCollectionControlType() 
                    { 
                        CardTypeId = RefUniversal.ID,
                        NodeId = MyHelper.RefItem_Devices,
                        SectionId = RefUniversal.Item.ID 
                    }
                };

            this.Edit_Completes.Session = Session;
            this.Edit_Completes.ObjectContext = Context;
            this.Edit_Completes.TypeIds = new List<MyCollectionControlType>()
                {
                    new MyCollectionControlType()
                    {
                        CardTypeId = RefUniversal.ID,
                        NodeId = MyHelper.RefItem_SKBCode,
                        SectionId = RefUniversal.Item.ID
                    }
                };

            AllDevices.Checked = true;
            ChooseDevices.Checked = false;
            ChooseCompletes.Checked = false;
            Edit_Devices.Enabled = false;
            Edit_Completes.Enabled = false;

        }

        private void Button_Click (Object sender, EventArgs e)
        {
            if (sender == Button_OK)
            {
                if (allDevices)
                {
                    DialogResult = DialogResult.OK;
                }
                else
                {
                    if (ChooseDevices.Checked)
                    {
                        if (!Devices.Any())
                        {
                            MyMessageBox.Show("Не выбрано ни одного прибора!", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            DialogResult = DialogResult.None;
                        }
                        else
                            DialogResult = DialogResult.OK;
                    }
                    else
                    {
                        if (ChooseCompletes.Checked)
                        {
                            if (!Completes.Any())
                            {
                                MyMessageBox.Show("Не выбрано ни одного комплектующего!", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                DialogResult = DialogResult.None;
                            }
                            else
                                DialogResult = DialogResult.OK;
                        }
                        else
                        {
                            MyMessageBox.Show("Выберите параметры отчета!", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            DialogResult = DialogResult.None;
                        }
                    }
                }
            }
            else
                DialogResult = DialogResult.None;
        }

        private void AllDevices_CheckedChanged(object sender, EventArgs e)
        {
            if (AllDevices.Checked)
            {
                ChooseDevices.Checked = false;
                Edit_Devices.Enabled = false;
                ChooseCompletes.Checked = false;
                Edit_Completes.Enabled = false;
            }
        }

        private void ChooseDevices_CheckedChanged(object sender, EventArgs e)
        {
            if (ChooseDevices.Checked)
            {
                AllDevices.Checked = false;
                ChooseCompletes.Checked = false;
                Edit_Completes.Enabled = false;

                Edit_Devices.Enabled = true;
            }
            else
            {
                Edit_Devices.Enabled = false;
            }
        }

        private void ChooseCompletes_CheckedChanged(object sender, EventArgs e)
        {
            if (ChooseCompletes.Checked)
            {
                AllDevices.Checked = false;
                ChooseDevices.Checked = false;
                Edit_Devices.Enabled = false;

                Edit_Completes.Enabled = true;
            }
            else
            {
                Edit_Completes.Enabled = false;
            }
        }
    }
}