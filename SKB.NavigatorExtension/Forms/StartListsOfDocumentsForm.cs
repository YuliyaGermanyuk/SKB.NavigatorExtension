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
    /// Форма выбора партий для запуска в производство.
    /// </summary>
    public partial class StartListsOfDocumentsForm : XtraForm
    {
        ObjectContext Context;
        UserSession Session;

        /// <summary>
        /// Выбранные партии для запуска в производство.
        /// </summary>
        public List<SelectionItem> Parties
        {
            get
            {
                return Edit_Parties.SelectedItems.Select(item => new SelectionItem(item.ObjectId, item.DisplayValue)).ToList();
            }
        }

        /// <summary>
        /// Инициализирует форму выбора партий для запуска в производство.
        /// </summary>
        /// <param name="Session">Пользовательская сессия.</param>
        /// <param name="Context">Объектный контекст.</param>
        public StartListsOfDocumentsForm (UserSession Session, ObjectContext Context)
        {
            InitializeComponent();

            this.SetNavigatorSkin();

            this.Context = Context;
            this.Session = Session;

            this.Edit_Parties.Session = Session;
            this.Edit_Parties.ObjectContext = Context;
            this.Edit_Parties.TypeIds = new List<MyCollectionControlType>()
                {
                    new MyCollectionControlType() 
                    { 
                        CardTypeId = RefUniversal.ID,
                        NodeId = MyHelper.RefItem_Party,
                        SectionId = RefUniversal.Item.ID 
                    }
                };
        }

        private void Button_Click (Object sender, EventArgs e)
        {
            if (sender == Button_Start)
            {
                if (Parties.Any())
                {
                    SearchQuery Query_Search = Session.CreateSearchQuery();
                    Query_Search.Limit = 0;
                    Query_Search.CombineResults = ConditionGroupOperation.Or;

                    CardTypeQuery Query_CardType = Query_Search.AttributiveSearch.CardTypeQueries.AddNew(CardOrd.ID);

                    SectionQuery Query_Section = Query_CardType.SectionQueries.AddNew(CardOrd.MainInfo.ID);
                    Query_Section.Operation = SectionQueryOperation.And;
                    Query_Section.ConditionGroup.Operation = ConditionGroupOperation.And;
                    Query_Section.ConditionGroup.Conditions.AddNew(CardOrd.MainInfo.Type, FieldType.RefId, ConditionOperation.Equals, MyHelper.RefType_ListofDocs);

                    Query_Section = Query_CardType.SectionQueries.AddNew(CardOrd.Properties.ID);
                    Query_Section.Operation = SectionQueryOperation.And;
                    Query_Section.ConditionGroup.Operation = ConditionGroupOperation.And;
                    Query_Section.ConditionGroup.Conditions.AddNew(CardOrd.Properties.Name, FieldType.Unistring, ConditionOperation.Equals, RefPropertiesListOfDocs.Requisities.Party);
                    ConditionGroup Query_ConditionGroup = Query_Section.ConditionGroup.ConditionGroups.AddNew();
                    Query_ConditionGroup.Operation = ConditionGroupOperation.Or;
                    foreach (SelectionItem Party in Parties)
                    {
                        Condition Query_Condition = Query_ConditionGroup.Conditions.AddNew(CardOrd.Properties.Value, FieldType.RefId, ConditionOperation.Equals, Party.Id);
                        Query_Condition.FieldSubtype = FieldSubtype.Universal;
                        Query_Condition.FieldSubtypeId = MyHelper.RefItem_Party;
                    }

                    CardDataCollection ListOfDocsDatas = Session.CardManager.FindCards(Query_Search.GetXml());
                    if (ListOfDocsDatas.Count > 0)
                    {
                        List<Guid> WrongPartyIds = ListOfDocsDatas.Select(card => card.Sections[CardOrd.Properties.ID].GetPropertyValue(RefPropertiesListOfDocs.Requisities.Party).ToGuid())
                           .Where(g => !g.IsEmpty()).Distinct().ToList();
                        List<String> WrongPartyNames = WrongPartyIds.Select(WrongPartyId => Edit_Parties.SelectedItems.FirstOrDefault(item => item.ObjectId == WrongPartyId))
                            .Where(item => !item.IsNull()).Select(item => "«" + item.DisplayValue + "»").ToList();
                        if (WrongPartyNames.Any())
                        {
                            if (WrongPartyNames.Count == 1)
                                MyMessageBox.Show("Партия " + WrongPartyNames[0] + " уже запущена в производство."
                                    + Environment.NewLine + "Удалите её из списка!", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            else
                                MyMessageBox.Show("Партии " + WrongPartyNames.Aggregate((a, b) => a + ", " + b) + " уже запущены в производство."
                                        + Environment.NewLine + "Удалите их из списка!", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            DialogResult = DialogResult.None;
                        }
                    }
                }
                else
                {
                    MyMessageBox.Show("Не выбрано ни одной партии!", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    DialogResult = DialogResult.None;
                }
            }
            else
                DialogResult = DialogResult.None;
        }
    }
}