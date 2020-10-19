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
    /// Форма выбора прибора для утверждения документации.
    /// </summary>
    public partial class ApproveListsOfDocumentsForm : XtraForm
    {
        ObjectContext Context;
        UserSession Session;

        /// <summary>
        /// Выбранный прибор для утверждения документации
        /// </summary>
        public SelectionItem Device
        {
            get
            {
                return Edit_Devices.SelectedItems.Any() ? Edit_Devices.SelectedItems.Select(item => new SelectionItem(item.ObjectId, item.DisplayValue)).FirstOrDefault() : new SelectionItem();
            }
        }

        /// <summary>
        /// Инициализирует форму выбора прибора для утверждения документации.
        /// </summary>
        /// <param name="Session">Пользовательская сессия.</param>
        /// <param name="Context">Объектный контекст.</param>
        public ApproveListsOfDocumentsForm (UserSession Session, ObjectContext Context)
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
                        SectionId = RefUniversal.Item.ID,
                        DisableChildSearch = false
                    }
                };
        }

        private void Button_Click (Object sender, EventArgs e)
        {
            if (sender == Button_Start)
            {
                if (Device.Id.IsEmpty())
                {
                    MyMessageBox.Show("Не выбран прибор!", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    DialogResult = DialogResult.None;
                }
            }
            else
                DialogResult = DialogResult.None;
        }
    }
}