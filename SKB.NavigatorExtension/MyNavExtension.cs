using DevExpress.XtraSplashScreen;
using DocsVision.BackOffice.ObjectModel;
using DocsVision.BackOffice.CardLib;
using DocsVision.BackOffice.CardLib.CardDefs;
using DocsVision.BackOffice.ObjectModel.Mapping;
using DocsVision.BackOffice.ObjectModel.Services;
using DocsVision.Platform.Cards.Constants;
using DocsVision.Platform.Data.Metadata;
using DocsVision.Platform.Extensibility;
using DocsVision.Platform.ObjectManager;
using DocsVision.Platform.ObjectManager.SystemCards;
using DocsVision.Platform.ObjectManager.SearchModel;
using DocsVision.Platform.ObjectManager.Metadata;
using DocsVision.Platform.ObjectModel;
using DocsVision.Platform.ObjectModel.Mapping;
using DocsVision.Platform.ObjectModel.Persistence;
using DocsVision.Platform.SystemCards.ObjectModel.Mapping;
using DocsVision.Platform.SystemCards.ObjectModel.Services;
using DocsVision.Platform.WinForms;
using RKIT.MyMessageBox;
using SKB.Base;
using SKB.Base.Dictionary;
using SKB.Base.Forms;
using SKB.Base.Ref;
using SKB.ListOfDocuments;
using SKB.NavigatorExtension.Forms;
using SKB.NavigatorExtension.Properties;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using RefProcessCard = DocsVision.Workflow.Constants.Process;
using RefUniversal = DocsVision.TakeOffice.Cards.Constants.RefUniversal;
using DocsVision.Platform.CardHost;
using System.Collections;
using DocsVision.TakeOffice.Cards.Constants;
using SKB.Base.Enums;
using DocsVision.Platform.Security.AccessControl;
using SKB.Base.AssignRights;

namespace SKB.NavigatorExtension
{
    /// <summary>
    /// Расширение Навигатора. 
    /// </summary>
    [ComVisible(true)]
    [Guid("725E48FD-FCFF-4166-A871-8DCD963D64F1")]
    [ClassInterface(ClassInterfaceType.None)]
    public class MyNavExtension : NavExtension
    {
        #region Properties
        /// <summary>
        /// Текст названия команды «StartListsOfDocuments».
        /// </summary>
        public const String Command_Name_StartListsOfDocuments = "Запуск партий в производство";
        /// <summary>
        /// Текст названия команды «StartListsOfDocuments».
        /// </summary>
        public const String Command_Name_StartListsOfDocuments_Folder = "Запуск партий в производство!";
        /// <summary>
        /// Текст описания команды «StartListsOfDocuments».
        /// </summary>
        public const String Command_Description_StartListsOfDocuments = "Формирование актуальных перечней технической документации для партий приборов";
        /// <summary>
        /// Текст названия команды «ApproveListsOfDocuments».
        /// </summary>
        public const String Command_Name_ApproveListsOfDocuments = "Утверждение документации";
        /// <summary>
        /// Текст названия команды «ApproveListsOfDocuments».
        /// </summary>
        public const String Command_Name_ApproveListsOfDocuments_Folder = "Утверждение документации ";
        /// <summary>
        /// Текст описания команды «ApproveListsOfDocuments».
        /// </summary>
        public const String Command_Description_ApproveListsOfDocuments = "Формирование перечня технической документации для утверждения на тех. совете";
        /// <summary>
        /// Текст названия команды «DeleteCardAndFiles».
        /// </summary>
        public const String Command_Name_DeleteCardAndFiles = "Удалить карточку и файлы";
        /// <summary>
        /// Текст описания команды «DeleteCardAndFiles».
        /// </summary>
        public const String Command_Description_DeleteCardAndFiles = "Удалить карточку и файлы в архиве.";
        /// <summary>
        /// Текст названия команды «SendToAgreement».
        /// </summary>
        public const String Command_Name_SendToAgreement = "Отправить на согласование";
        /// <summary>
        /// Текст описания команды «DeleteCardAndFiles».
        /// </summary>
        public const String Command_Description_SendToAgreement = "Отправить на согласование.";
        /// <summary>
        /// Текст названия команды «ReportCalibrationLaboratory».
        /// </summary>
        public const String Command_Name_ReportCalibrationLaboratory = "Отчет по калибровочной лаборатории";
        /// <summary>
        /// Текст описания команды «ReportCalibrationLaboratory».
        /// </summary>
        public const String Command_Description_ReportCalibrationLaboratory = "Отчет \"План-факт\" по работе калибровочной лаборатории.";
        /// <summary>
        /// Текст названия команды «ReportWarehouse».
        /// </summary>
        public const String Command_Name_ReportWarehouse = "Отчет по складу готовой продукции";
        /// <summary>
        /// Текст описания команды «ReportWarehouse».
        /// </summary>
        public const String Command_Description_ReportWarehouse = "Отчет по движению приборов на складе готовой продукции.";
        /// <summary>
        /// Текст названия команды «LoadCalibrationDocuments».
        /// </summary>
        public const String Command_Name_LoadCalibrationDocuments = "Загрузить протоколы калибровки";
        /// <summary>
        /// Текст описания команды «LoadCalibrationDocuments».
        /// </summary>
        public const String Command_Description_LoadCalibrationDocuments = "Загрузить протоколы калибровки в паспорта приборов.";
        /// <summary>
        /// Текст названия команды «FillingVerifyConditionsJournal238».
        /// </summary>
        public const String Command_Name_FillingVerifyConditionsJournal238 = "Регистрация условий поверки: каб. №228";
        /// <summary>
        /// Текст описания команды «FillingVerifyConditionsJournal238».
        /// </summary>
        public const String Command_Description_FillingVerifyConditionsJournal238 = "Заполнить журнал условий поверки для кабинета №228.";
        /// <summary>
        /// Текст названия команды «FillingVerifyConditionsJournal237».
        /// </summary>
        public const String Command_Name_FillingVerifyConditionsJournal237 = "Регистрация условий поверки: каб. №226";
        /// <summary>
        /// Текст описания команды «FillingVerifyConditionsJournal237».
        /// </summary>
        public const String Command_Description_FillingVerifyConditionsJournal237 = "Заполнить журнал условий поверки для кабинета №226.";
        /// <summary>
		/// Текст названия команды «CreateTaskManager».
		/// </summary>
		public const String Command_Name_CreateTaskManager = "Создать задание по прибору";
        /// <summary>
        /// Текст описания команды «CreateTaskManager».
        /// </summary>
        public const String Command_Description_CreateTaskManager = "Создание карточки задания по прибору";
        /// <summary>
        /// Текст названия команды «CreateTaskOnSub».
        /// </summary>
        public const String Command_Name_CreateTaskOnSub = "Создать подчиненное задание";
        /// <summary>
        /// Текст описания команды «CreateTaskOnSub».
        /// </summary>
        public const String Command_Description_CreateTaskOnSub = "Создание карточки задания для подэтапа";
        /// <summary>
        /// Текст названия команды «BalanceComplete».
        /// </summary>
        public const String Command_Name_BalanceComplete = "Остаток комплектующих";
        /// <summary>
        /// Текст описания команды «BalanceComplete».
        /// </summary>
        public const String Command_Description_BalanceComplete = "Остаток комплектующих на текущий момент";

        private ObjectContext objectContext;
        private IStaffService staffService;
        private IStateService stateService;
        private IAccessCheckingService accessService;
        private ILockService lockService;
        

        private ObjectContext Context
        {
            get
            {
                if (objectContext == null)
                {
                    objectContext = new ObjectContext(this);

                    var mapperFactoryRegistry = objectContext.GetService<IObjectMapperFactoryRegistry>();
                    mapperFactoryRegistry.RegisterFactory(typeof(SystemCardsMapperFactory));
                    mapperFactoryRegistry.RegisterFactory(typeof(BackOfficeMapperFactory));

                    var serviceFactoryRegistry = objectContext.GetService<IServiceFactoryRegistry>();
                    serviceFactoryRegistry.RegisterFactory(typeof(SystemCardsServiceFactory));
                    serviceFactoryRegistry.RegisterFactory(typeof(BackOfficeServiceFactory));

                    IMetadataProvider metadataProvider = DocsVisionObjectFactory.CreateMetadataProvider(new SessionProvider(this.Session));
                    objectContext.AddService<IMetadataManager>(DocsVisionObjectFactory.CreateMetadataManager(metadataProvider, this.Session));
                    objectContext.AddService<IMetadataProvider>(metadataProvider);
                    objectContext.AddService<IPersistentStore>(DocsVisionObjectFactory.CreatePersistentStore(this.Session));
                }

                return objectContext;
            }
        }

        private IStaffService StaffService
        {
            get
            {
                if (staffService.IsNull())
                    staffService = Context.GetService<IStaffService>();
                return staffService;
            }
        }
        private IStateService StateService
        {
            get
            {
                if (stateService.IsNull())
                    stateService = Context.GetService<IStateService>();
                return stateService;
            }
        }
        private IAccessCheckingService AccessService
        {
            get
            {
                if (accessService.IsNull())
                    accessService = Context.GetService<IAccessCheckingService>();
                return accessService;
            }
        }
        private ILockService LockService
        {
            get
            {
                if (lockService.IsNull())
                    lockService = Context.GetService<ILockService>();
                return lockService;
            }
        }

        #endregion

        /// <summary>
        /// Инициализирует расширение Навигатора.
        /// </summary>
        public MyNavExtension () { }

        #region Command Methods
        /// <summary>
        /// Выполняет команду «Запуск партий в производство».
        /// </summary>
        public void StartListsOfDocuments ()
        {
            /*  */
            StartListsOfDocumentsForm Form = new StartListsOfDocumentsForm(Session, Context);
            switch (Form.ShowDialog())
            {
                case System.Windows.Forms.DialogResult.OK:
                    SplashScreenManager.CloseForm(false);
                    SplashScreenManager.ShowForm(typeof(MyWaitForm), true, true);
                    SplashScreenManager.Default.SetWaitFormDescription("Получение данных...");
                    FolderCard FolderCard = (FolderCard)Session.CardManager.GetCard(FoldersCard.ID, false);
                    foreach (SelectionItem Party in Form.Parties)
                    {
                        SplashScreenManager.Default.SetWaitFormDescription("Обработка «" + Party.Name + "»...");
                        /* Cоздание бизнес-процесса из шаблона */
                        CardData ListOfDocProcessData = Session.CardManager.GetCardData(ListProcesses.RefProcessTemplate_DocumnetsCompilation).Copy();

                        ListOfDocProcessData.BeginUpdate();
                        ListOfDocProcessData.IsTemplate = false;
                        SectionData VariablesSection = ListOfDocProcessData.Sections[RefProcessCard.Variables.ID];
                        VariablesSection.Rows.Find(RefProcessCard.Variables.Name, "Составитель").SetGuid(RefProcessCard.Variables.Value, Context.GetCurrentUser());
                        VariablesSection.Rows.Find(RefProcessCard.Variables.Name, "Партия").SetString(RefProcessCard.Variables.Value, RefUniversal.ID.ToString("B")
                            + RefUniversal.Item.ID.ToString("B")
                            + Party.Id.ToString("B"));
                        ListOfDocProcessData.EndUpdate();

                        if (!FolderCard.GetShortcuts(ListOfDocProcessData.Id).Any(sc => sc.IsHardLink))
                            FolderCard.CreateShortcut(ListProcesses.RefListofDocsFolder, ListOfDocProcessData.Id, true);

                        Session.StartProcess(ListOfDocProcessData.Id);
                    }
                    SplashScreenManager.CloseForm(false);
                    if (Form.Parties.Count > 1)
                        MyMessageBox.Show("«Для указанных партий сформированы «Перечни технической документации» и отправлены" + Environment.NewLine
                            + "на согласование в конструкторско-технологический отдел." + Environment.NewLine
                            + "Перед началом производства дождитесь утверждения «Перечней технической документации».", "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    else
                        MyMessageBox.Show("«Для указанной партии сформирован «Перечень технической документации» и отправлен" + Environment.NewLine
                            + "на согласование в конструкторско-технологический отдел." + Environment.NewLine
                            + "Перед началом производства дождитесь утверждения «Перечня технической документации».", "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    break;
            }
        }
        /// <summary>
        /// Выполняет команду «Утверждение документации».
        /// </summary>
        public void ApproveListsOfDocuments ()
        {
            ApproveListsOfDocumentsForm Form = new ApproveListsOfDocumentsForm(Session, Context);
            switch (Form.ShowDialog())
            {
                case System.Windows.Forms.DialogResult.OK:
                    SplashScreenManager.ShowForm(typeof(MyWaitForm), true, true);
                    SplashScreenManager.Default.SetWaitFormDescription("Получение данных...");
                    FolderCard FolderCard = (FolderCard)Session.CardManager.GetCard(FoldersCard.ID, false);

                    SplashScreenManager.Default.SetWaitFormDescription("Обработка прибора «" + Form.Device.Name + "»...");
                    /* Cоздание бизнес-процесса из шаблона */
                    CardData ListOfDocProcessData = Session.CardManager.GetCardData(ListProcesses.RefProcessTemplate_ApproveDocumentation).Copy();

                    ListOfDocProcessData.BeginUpdate();
                    ListOfDocProcessData.IsTemplate = false;
                    SectionData VariablesSection = ListOfDocProcessData.Sections[RefProcessCard.Variables.ID];
                    VariablesSection.Rows.Find(RefProcessCard.Variables.Name, "Запустивший").SetGuid(RefProcessCard.Variables.Value, Context.GetCurrentUser());
                    VariablesSection.Rows.Find(RefProcessCard.Variables.Name, "Прибор").SetString(RefProcessCard.Variables.Value, RefUniversal.ID.ToString("B")
                        + RefUniversal.Item.ID.ToString("B")
                        + Form.Device.Id.ToString("B"));
                    ListOfDocProcessData.EndUpdate();

                    if (!FolderCard.GetShortcuts(ListOfDocProcessData.Id).Any(sc => sc.IsHardLink))
                        FolderCard.CreateShortcut(ListProcesses.RefListofDocsFolder, ListOfDocProcessData.Id, true);

                    Session.StartProcess(ListOfDocProcessData.Id);
                    SplashScreenManager.CloseForm(false);
                    MyMessageBox.Show("«Для указанного прибора сформирован «Перечень технической документации» и отправлен" + Environment.NewLine
                        + "на согласование в конструкторско-технологический отдел." + Environment.NewLine
                        + "Перед началом производства дождитесь утверждения «Перечня технической документации».", "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    break;
            }
        }
        /// <summary>
        /// Определяет доступность и/или выполняет команду «Удалить карточку и файлы».
        /// </summary>
        /// <param name="OnlyCheck">Только проверить доступность.</param>
        /// <param name="Ids">Проверяемые объекты.</param>
        /// <returns></returns>
        public NavCommandStatus DeleteCardAndFiles (Boolean OnlyCheck, params Guid[] Ids)
        {
            Dictionary<Guid, BaseCard> Cards = Ids.Where(cardId => Session.CardManager.GetCardState(cardId) == DocsVision.Platform.ObjectManager.ObjectState.Existing
                 && Session.CardManager.GetCardData(cardId).Type.Id == RefMarketingFilesCard.ID)
                 .ToDictionary(cardId => cardId, cardId => Context.GetObject<BaseCard>(cardId))
                 .Where(card => !card.Value.IsNull()).ToDictionary(card => card.Key, card => card.Value);

            if (Cards.Any())
            {
                Cards = Cards.ToDictionary(card => card, card => StateService.GetOperations(card.Value.SystemInfo.CardKind))
                                .Where(card => !card.Value.IsNull())
                                .ToDictionary(card => card.Key, card => card.Value.FirstOrDefault(item => item.DefaultName == "Modify"))
                                .Where(card => !card.Value.IsNull() && AccessService.IsOperationAllowed(card.Key.Value, card.Value))
                                .ToDictionary(card => card.Key.Key, card => card.Key.Value);
                if (Cards.Any())
                {
                    if (!OnlyCheck)
                    {
                        switch (ShowMessage(Cards.Count == 1 ? "Вы уверены, что хотите удалить выбранную карточку и связанные файлы?"
                            : "Вы уверены, что хотите удалить выбранные карточки и связанные с ними файлы?", "Docsvision Navigator", MessageType.Question, MessageButtons.YesNo))
                        {
                            case MessageResult.Yes:
                                List<String> GoodCards = new List<String>();
                                List<String> WarnCards = new List<String>();
                                List<String> ErrorCards = new List<String>();
                                foreach (KeyValuePair<Guid, BaseCard> Card in Cards)
                                {
                                    Boolean ByMe;
                                    String OwnerName;
                                    if (!LockService.IsObjectLocked<BaseCard>(Card.Value, out ByMe, out OwnerName))
                                    {
                                        String CardDescription = Card.Value.Description;
                                        if (Session.DeleteCard(Card.Key))
                                            GoodCards.Add(CardDescription);
                                        else
                                            ErrorCards.Add(CardDescription);
                                    }
                                    else
                                        WarnCards.Add("Невозможно удалить карточку " + Card.Value.Description + "." + Environment.NewLine
                                                + "Карточка заблокирована " + (ByMe ? "вами" : "пользователем " + OwnerName) + "!");
                                }

                                if (GoodCards.Any())
                                {
                                    if (GoodCards.Count == 1)
                                        ShowMessage("Карточка и файлы удалены!", "Docsvision Navigator", MessageType.Information, MessageButtons.Ok);
                                    else
                                    {
                                        ShowMessage("Карточки и файлы удалены!", "Docsvision Navigator",
                                            "Удаленные карточки:" + Environment.NewLine + GoodCards.Aggregate((a, b) => a + ";" + Environment.NewLine + b), MessageType.Information, MessageButtons.Ok);
                                    }
                                }
                                if (WarnCards.Any())
                                {
                                    if (WarnCards.Count == 1)
                                        ShowMessage(WarnCards[0], "Docsvision Navigator", MessageType.Warning, MessageButtons.Ok);
                                    else
                                    {
                                        ShowMessage("Невозможно удалить карточки!", "Docsvision Navigator",
                                            WarnCards.Aggregate((a, b) => a + ";" + Environment.NewLine + Environment.NewLine + b), MessageType.Information, MessageButtons.Ok);
                                    }
                                }
                                if (ErrorCards.Any())
                                {
                                    if (ErrorCards.Count == 1)
                                        ShowMessage("Не удалось удалить карточку!" + Environment.NewLine
                                            + "Обратитесь к системному администратору!", "Docsvision Navigator", MessageType.Error, MessageButtons.Ok);
                                    else
                                    {
                                        ShowMessage("Не удалось удалить карточки!", "Docsvision Navigator",
                                            "Неудаленные карточки: " + Environment.NewLine + ErrorCards.Aggregate((a, b) => a + ";" + Environment.NewLine + b), MessageType.Information, MessageButtons.Ok);
                                    }
                                }
                                break;
                        }
                    }
                    return NavCommandStatus.Supported | NavCommandStatus.Enabled;
                }
                return NavCommandStatus.Supported;
            }
            return NavCommandStatus.None;
        }
        /// <summary>
        /// Определяет доступность и/или выполняет команду «Отправить на согласование».
        /// </summary>
        /// <param name="OnlyCheck">Только проверить доступность.</param>
        /// <param name="Ids">Проверяемые объекты.</param>
        /// <returns></returns>
        public NavCommandStatus SendToAgreement(Boolean OnlyCheck, params Guid[] Ids)
        {
                Guid AgreementOfDocumentsFolderID = new Guid("{35F2FF8E-0EC7-44F9-920B-922C90919C57}");

                Guid[] AgreementTypes = new Guid[] { MyHelper.RefType_CD, MyHelper.RefType_TD };
                IEnumerable<CardData> Cards = Ids.Where(cardId => Session.CardManager.GetCardState(cardId) == DocsVision.Platform.ObjectManager.ObjectState.Existing)
                    .Select(cardId => Session.CardManager.GetCardData(cardId))
                    .Where(Card => Card.Type.Id == CardOrd.ID && AgreementTypes.Any(TypeGuid => TypeGuid == Card.Sections[CardOrd.MainInfo.ID].FirstRow.GetGuid(CardOrd.MainInfo.Type)));
                List<CardData> VerifyCards = new List<CardData>();

                if (Cards.Any())
                {
                    if (!OnlyCheck)
                    {
                        List<MyException> ErrorsList = new List<MyException>();
                        foreach (CardData Card in Cards)
                        {
                            if (!Session.AccessManager.AccessCheck(SecureObjectType.Card, Card.Id, Guid.Empty, (Int32)(CardDataRights.Read | CardDataRights.Modify | CardDataRights.Copy)))
                                ErrorsList.Add(new MyException(0, Card.Description));
                            else
                            {
                                if ((Int32?)Card.Sections[CardOrd.Properties.ID].GetPropertyValue("Статус") == (Int32)DocumentState.Draft)
                                    VerifyCards.Add(Card);
                                else
                                    ErrorsList.Add(new MyException(2, Card.Description));   // ErrorsList.Add(DocData.Description + ": Выбраный документ не является черновиком!"); // throw new MyException("Выбраный документ не является черновиком!");
                            }
                        }
                        if (ErrorsList.Count() > 0)
                        {
                            if (Cards.Count() == 1)
                                MyMessageBox.Show(ExtensionHelper.GetErrorText(ErrorsList.First().ErrorCode, 1) + ".", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            else
                                MyMessageBox.Show(ErrorsList.GroupBy(r => r.ErrorCode)
                                    .Select(r => ExtensionHelper.GetErrorText(r.First().ErrorCode, r.Count()) + ":\n" + r.Select(s => " - " + s.Message).Aggregate(";\n") + ".")
                                    .Aggregate("\n") + "\n" + ErrorsList.Count().GetCaseString("Он не будет добавлен", "Все вышеперечисленные документы не будут добавлены", "Все вышеперечисленные документы не будут добавлены") + " в карточку согласования.", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }
                        if (VerifyCards.Any())
                        {
                            CardData NewCard = Context.CreateCard(RefAgreementOfDocumentsCard.ID);
                            FolderCard FolderCard = (FolderCard)Session.CardManager.GetCard(FoldersCard.ID, false);
                            FolderCard.CreateShortcut(AgreementOfDocumentsFolderID, NewCard.Id, true);

                            RowDataCollection DocumentRows = NewCard.Sections[RefAgreementOfDocumentsCard.Documents.ID].Rows;
                            RowDataCollection DevicesCollection = NewCard.Sections[RefAgreementOfDocumentsCard.Devices.ID].Rows;

                            foreach (CardData DocumentCard in VerifyCards)
                            {
                                ExtraCard DocExtraCard = ExtraCardCD.GetExtraCard(DocumentCard);
                                if (DocExtraCard.IsNull())
                                    DocExtraCard = ExtraCardTD.GetExtraCard(DocumentCard);
                                RowData NewRow = DocumentRows.AddNew();

                                NewRow[RefAgreementOfDocumentsCard.Documents.CodeID] = DocExtraCard.CodeId;
                                NewRow[RefAgreementOfDocumentsCard.Documents.DocumentsName] = DocExtraCard.Name;
                                NewRow[RefAgreementOfDocumentsCard.Documents.DocumentsCategory] = DocExtraCard.CategoryId;
                                NewRow[RefAgreementOfDocumentsCard.Documents.DocumentsVersion] = DocExtraCard.Version;
                                NewRow[RefAgreementOfDocumentsCard.Documents.DocumentsComment] = "";
                                NewRow[RefAgreementOfDocumentsCard.Documents.DocumentsAuthor] = DocExtraCard.DeveloperID;
                                NewRow[RefAgreementOfDocumentsCard.Documents.DocumentsCard] = DocExtraCard.Card.Id;
                                NewRow[RefAgreementOfDocumentsCard.Documents.IsApproved] = false;
                                NewRow[RefAgreementOfDocumentsCard.Documents.ApprovalDate] = null;
                                if (NewRow[RefAgreementOfDocumentsCard.Documents.Id] == null || NewRow[RefAgreementOfDocumentsCard.Documents.Id].ToGuid() == Guid.Empty)
                                    NewRow[RefAgreementOfDocumentsCard.Documents.Id] = Guid.NewGuid();

                                IEnumerable<String> NewDevices = DocExtraCard.Devices.Except(DevicesCollection.Select(r => r.GetGuid(RefAgreementOfDocumentsCard.Devices.Id).ToString().ToUpper()));
                                if (!NewDevices.IsNull())
                                {
                                    foreach (String DeviceId in NewDevices)
                                        DevicesCollection.AddNew().SetGuid(RefAgreementOfDocumentsCard.Devices.Id, DeviceId.ToGuid());
                                }

                                // Выдача номера карточке согласования
                                //BaseCardNumber CurrentNumerator = CardScript.GetNumber(RefAgreementOfDocumentsCard.NumberRuleName);
                                //CurrentNumerator.Number = Convert.ToInt32(CurrentNumerator.Number).ToString("00000");
                                //SetControlValue(RefAgreementOfDocumentsCard.MainInfo.Number, Context.GetObjectRef<BaseCardNumber>(CurrentNumerator).Id);

                                RowData MainInfoRow = NewCard.Sections[RefAgreementOfDocumentsCard.MainInfo.ID].FirstRow;
                                MainInfoRow.SetGuid(RefAgreementOfDocumentsCard.MainInfo.Registrar, Context.GetObjectRef(Context.GetCurrentEmployee()).Id);
                                MainInfoRow.SetDateTime(RefAgreementOfDocumentsCard.MainInfo.CreationDate, DateTime.Now);
                                MainInfoRow.SetInt32(RefAgreementOfDocumentsCard.MainInfo.State, (int)RefAgreementOfDocumentsCard.MainInfo.DisplayCardState.NotStarted);

                            }
                            ICardHost host = DocsVision.Platform.CardHost.CardHost.CreateInstance(Session);
                            if (!host.ShowCardModal(NewCard.Id, DocsVision.Platform.CardHost.ActivateMode.Edit, DocsVision.Platform.CardHost.ActivateFlags.New))
                            {
                                Session.ReleaseNumber(Context.GetObject<BaseCardNumber>(NewCard.Sections[RefAgreementOfDocumentsCard.MainInfo.ID].FirstRow.GetGuid(RefAgreementOfDocumentsCard.MainInfo.Number)).NumericPart);
                                FolderCard.DeleteCard(NewCard.Id, true);
                            }                               
                        }
                    }
                    return NavCommandStatus.Supported | NavCommandStatus.Enabled;
                }
                return NavCommandStatus.None;
        }
        /// <summary>
        /// Выполняет команду «Отчет по калибровочной лаборатории».
        /// </summary>
        public void GetReportCalibrationLaboratory()
        {
            
            ReportParameters Form = new ReportParameters(Session, Context);
            switch (Form.ShowDialog())
            {
                case System.Windows.Forms.DialogResult.OK:
                    ReportCalibrationLaboratory NewReport = new ReportCalibrationLaboratory(Session, Form.StartDateValue, Form.EndDateValue);
                    SplashScreenManager.ShowForm(typeof(MyWaitForm), true, true);
                    SplashScreenManager.Default.SetWaitFormDescription("Идет формирование отчета...");
                    NewReport.Create();
                    SplashScreenManager.CloseForm(false);
                    NewReport.OpenReport();
                    break;
            }
        }
        /// <summary>
        /// Выполняет команду «Отчет по складу готовой продукции».
        /// </summary>
        public void GetReportWarehouse()
        {

            ReportParameters Form = new ReportParameters(Session, Context);
            switch (Form.ShowDialog())
            {
                case System.Windows.Forms.DialogResult.OK:
                    SplashScreenManager.ShowForm(typeof(MyWaitForm), true, true);
                    SplashScreenManager.Default.SetWaitFormDescription("Начато формирование отчета...");

                    CardData UniversalDictionary = Session.CardManager.GetDictionaryData(RefUniversal.ID);

                    // Поиск передач приборов по актам
                    SplashScreenManager.Default.SetWaitFormDescription("Идет поиск актов передачи приборов...");
                    IEnumerable<TransferRow> Acts = ReportHelper.FindAct(Session, Form.StartDateValue, Form.EndDateValue, UniversalDictionary);
                    SplashScreenManager.Default.SetWaitFormDescription("Идет поиск заданий на комплектацию...");
                    IEnumerable<TransferRow> Tasks = ReportHelper.FindCompleteTasks(Session, Form.StartDateValue, Form.EndDateValue, UniversalDictionary);
                    SplashScreenManager.Default.SetWaitFormDescription("Идет объединение результатов поиска...");
                    IEnumerable<TransferRow> AllTransfers = Acts.Union(Tasks);

                    // ОСТАТОК НА НАЧАЛО ПЕРИОДА //
                    // На складе
                    SplashScreenManager.Default.SetWaitFormDescription("Определяется остаток приборов на складе ГП на начало периода...");
                    IEnumerable<TransferCountByParty> InWarehouseAtBeginningByParty = AllTransfers.StatisticsOnDate(UniversalDictionary, Form.StartDateValue, TransferRow.TransferTypes.ToWarehouse);
                    // На выставках
                    SplashScreenManager.Default.SetWaitFormDescription("Определяется остаток приборов на выставках на начало периода...");
                    IEnumerable<TransferCountByParty> InExpositionAtBeginningByParty = AllTransfers.StatisticsOnDate(UniversalDictionary, Form.StartDateValue, TransferRow.TransferTypes.FromWarehouse, TransferRow.Action.DeliveryToExposition);
                    // На сертификации
                    SplashScreenManager.Default.SetWaitFormDescription("Определяется остаток приборов на сертификации на начало периода...");
                    IEnumerable<TransferCountByParty> InCertificationAtBeginningByParty = AllTransfers.StatisticsOnDate(UniversalDictionary, Form.StartDateValue, TransferRow.TransferTypes.FromWarehouse, TransferRow.Action.DeliveryToCertification);
                    // На испытаниях
                    SplashScreenManager.Default.SetWaitFormDescription("Определяется остаток приборов на испытаниях на начало периода...");
                    IEnumerable<TransferCountByParty> InTestingAtBeginningByParty = AllTransfers.StatisticsOnDate(UniversalDictionary, Form.StartDateValue, TransferRow.TransferTypes.FromWarehouse, TransferRow.Action.DeliveryToTesting);
                    // На тест-драйве
                    SplashScreenManager.Default.SetWaitFormDescription("Определяется остаток приборов на тест-драйве на начало периода...");
                    IEnumerable<TransferCountByParty> InTestDriveAtBeginningByParty = AllTransfers.StatisticsOnDate(UniversalDictionary, Form.StartDateValue, TransferRow.TransferTypes.FromWarehouse, TransferRow.Action.DeliveryToTestDrive);

                    // ПРИХОД ЗА ПЕРИОД //
                    // Приход из производства (новые приборы)
                    SplashScreenManager.Default.SetWaitFormDescription("Определяется перечень новых приборов, переданных из производства...");
                    IEnumerable<TransferCountByParty> NewReceiptFromProductionByParty = AllTransfers.StatisticsForPeriod(Form.StartDateValue, Form.EndDateValue, TransferRow.Action.NewReceiptFromProduction, UniversalDictionary);
                    // Приход из производства (повторная передача)
                    SplashScreenManager.Default.SetWaitFormDescription("Определяется перечень приборов, повторно переданных из производства...");
                    IEnumerable<TransferCountByParty> RepeatReceiptFromProductionByParty = AllTransfers.StatisticsForPeriod(Form.StartDateValue, Form.EndDateValue, TransferRow.Action.RepeatReceiptFromProduction, UniversalDictionary);
                    // Возврат на склад с выставок
                    SplashScreenManager.Default.SetWaitFormDescription("Определяется перечень приборов, возвращенных на склад с выставок...");
                    IEnumerable<TransferCountByParty> ReturnFromExpositionByParty = AllTransfers.StatisticsForPeriod(Form.StartDateValue, Form.EndDateValue, TransferRow.Action.ReturnFromExposition, UniversalDictionary);
                    // Возврат на склад с сертификации
                    SplashScreenManager.Default.SetWaitFormDescription("Определяется перечень приборов, возвращенных на склад с сертификации...");
                    IEnumerable<TransferCountByParty> ReturnFromCertificationByParty = AllTransfers.StatisticsForPeriod(Form.StartDateValue, Form.EndDateValue, TransferRow.Action.ReturnFromCertification, UniversalDictionary);
                    // Возврат на склад с испытаний
                    SplashScreenManager.Default.SetWaitFormDescription("Определяется перечень приборов, возвращенных на склад с испытаний...");
                    IEnumerable<TransferCountByParty> ReturnFromTestingByParty = AllTransfers.StatisticsForPeriod(Form.StartDateValue, Form.EndDateValue, TransferRow.Action.ReturnFromTesting, UniversalDictionary);
                    // Возврат на склад с тест-драйва
                    SplashScreenManager.Default.SetWaitFormDescription("Определяется перечень приборов, возвращенных на склад с тест-драйва...");
                    IEnumerable<TransferCountByParty> ReturnFromTestDriveByParty = AllTransfers.StatisticsForPeriod(Form.StartDateValue, Form.EndDateValue, TransferRow.Action.ReturnFromTestDrive, UniversalDictionary);

                    // РАСХОД ЗА ПЕРИОД //
                    // Возврат со склада в производство
                    SplashScreenManager.Default.SetWaitFormDescription("Определяется перечень приборов, возвращенных со склада в производство...");
                    IEnumerable<TransferCountByParty> ReturnFromWarehouseToProductionByParty = AllTransfers.StatisticsForPeriod(Form.StartDateValue, Form.EndDateValue, TransferRow.Action.ReturnFromWarehouseToProduction, UniversalDictionary);
                    // Выдача со склада на выставки
                    SplashScreenManager.Default.SetWaitFormDescription("Определяется перечень приборов, выданных со склада на выставки...");
                    IEnumerable<TransferCountByParty> DeliveryToExpositionByParty = AllTransfers.StatisticsForPeriod(Form.StartDateValue, Form.EndDateValue, TransferRow.Action.DeliveryToExposition, UniversalDictionary);
                    // Выдача со склада на сертификацию
                    SplashScreenManager.Default.SetWaitFormDescription("Определяется перечень приборов, выданных со склада на сертификацию...");
                    IEnumerable<TransferCountByParty> DeliveryToCertificationByParty = AllTransfers.StatisticsForPeriod(Form.StartDateValue, Form.EndDateValue, TransferRow.Action.DeliveryToCertification, UniversalDictionary);
                    // Выдача со склада на испытания
                    SplashScreenManager.Default.SetWaitFormDescription("Определяется перечень приборов, выданных со склада на испытания...");
                    IEnumerable<TransferCountByParty> DeliveryToTestingByParty = AllTransfers.StatisticsForPeriod(Form.StartDateValue, Form.EndDateValue, TransferRow.Action.DeliveryToTesting, UniversalDictionary);
                    // Передача на тест-драйв
                    SplashScreenManager.Default.SetWaitFormDescription("Определяется перечень приборов, переданных со склада на тест-драйв...");
                    IEnumerable<TransferCountByParty> DeliveryToTestDriveByParty = AllTransfers.StatisticsForPeriod(Form.StartDateValue, Form.EndDateValue, TransferRow.Action.DeliveryToTestDrive, UniversalDictionary);
                    // Отгрузка новых приборов
                    SplashScreenManager.Default.SetWaitFormDescription("Определяется перечень отгруженных приборов...");
                    IEnumerable<TransferCountByParty> DeliveryNewDevicesByParty = AllTransfers.StatisticsForPeriod(Form.StartDateValue, Form.EndDateValue, TransferRow.Action.DeliveryNewDevices, UniversalDictionary);

                    // ОСТАТОК НА КОНЕЦ ПЕРИОДА //
                    // На складе готовой продукции
                    SplashScreenManager.Default.SetWaitFormDescription("Определяется остаток приборов на складе ГП на конец периода...");
                    IEnumerable<TransferCountByParty> InWarehouseAtEndingByParty = AllTransfers.StatisticsOnDate(UniversalDictionary, Form.EndDateValue.AddDays(1), TransferRow.TransferTypes.ToWarehouse);
                    // На выставках
                    SplashScreenManager.Default.SetWaitFormDescription("Определяется остаток приборов на выставках на конец периода...");
                    IEnumerable<TransferCountByParty> InExpositionAtEndingByParty = AllTransfers.StatisticsOnDate(UniversalDictionary, Form.EndDateValue.AddDays(1), TransferRow.TransferTypes.FromWarehouse, TransferRow.Action.DeliveryToExposition);
                    // На сертификации
                    SplashScreenManager.Default.SetWaitFormDescription("Определяется остаток приборов на сертификации на конец периода...");
                    IEnumerable<TransferCountByParty> InCertificationAtEndingByParty = AllTransfers.StatisticsOnDate(UniversalDictionary, Form.EndDateValue.AddDays(1), TransferRow.TransferTypes.FromWarehouse, TransferRow.Action.DeliveryToCertification);
                    // На испытаниях
                    SplashScreenManager.Default.SetWaitFormDescription("Определяется остаток приборов на испытаниях на конец периода...");
                    IEnumerable<TransferCountByParty> InTestingAtEndingByParty = AllTransfers.StatisticsOnDate(UniversalDictionary, Form.EndDateValue.AddDays(1), TransferRow.TransferTypes.FromWarehouse, TransferRow.Action.DeliveryToTesting);
                    // На тест-драйве
                    SplashScreenManager.Default.SetWaitFormDescription("Определяется остаток приборов на тест-драйве на конец периода...");
                    IEnumerable<TransferCountByParty> InTestDriveAtEndingByParty = AllTransfers.StatisticsOnDate(UniversalDictionary, Form.EndDateValue.AddDays(1), TransferRow.TransferTypes.FromWarehouse, TransferRow.Action.DeliveryToTestDrive);

                    // СОЗДАНИЕ ОТЧЕТА //
                    SplashScreenManager.Default.SetWaitFormDescription("Идет занесение данных в отчет...");
                    ReportWarehouse ReportDocument = new ReportWarehouse(Session, Form.StartDateValue, Form.EndDateValue);

                    // ОСТАТОК НА НАЧАЛО ПЕРИОДА //
                    // Занесение данных в отчет: На складе на начало периода
                    SplashScreenManager.Default.SetWaitFormDescription("Идет занесение данных об остатке приборов на складе на начало периода...");
                    foreach (TransferCountByParty Row in InWarehouseAtBeginningByParty)
                        ReportDocument.AddValue(Row.DeviceType, Row.DeviceParty, Columns.InWarehouseAtBeginningPeriod, Row.DeviceCount, Row.DeviceNumbersCollection);
                    // Занесение данных в отчет: На выставке на начало периода
                    SplashScreenManager.Default.SetWaitFormDescription("Идет занесение данных об остатке приборов на выставках на начало периода...");
                    foreach (TransferCountByParty Row in InExpositionAtBeginningByParty)
                        ReportDocument.AddValue(Row.DeviceType, Row.DeviceParty, Columns.InExpositionAtBeginningPeriod, Row.DeviceCount, Row.DeviceNumbersCollection);
                    // Занесение данных в отчет: На сертификации на начало периода
                    SplashScreenManager.Default.SetWaitFormDescription("Идет занесение данных об остатке приборов на сертификации на начало периода...");
                    foreach (TransferCountByParty Row in InCertificationAtBeginningByParty)
                        ReportDocument.AddValue(Row.DeviceType, Row.DeviceParty, Columns.InCertificationAtBeginningPeriod, Row.DeviceCount, Row.DeviceNumbersCollection);
                    // Занесение данных в отчет: На испытаниях на начало периода
                    SplashScreenManager.Default.SetWaitFormDescription("Идет занесение данных об остатке приборов на испытаниях на начало периода...");
                    foreach (TransferCountByParty Row in InTestingAtBeginningByParty)
                        ReportDocument.AddValue(Row.DeviceType, Row.DeviceParty, Columns.InTestingAtBeginningPeriod, Row.DeviceCount, Row.DeviceNumbersCollection);
                    // Занесение данных в отчет: На тест-драйве на начало периода
                    SplashScreenManager.Default.SetWaitFormDescription("Идет занесение данных об остатке приборов на тест-драйве на начало периода...");
                    foreach (TransferCountByParty Row in InTestDriveAtBeginningByParty)
                        ReportDocument.AddValue(Row.DeviceType, Row.DeviceParty, Columns.InTestDriveAtBeginningPeriod, Row.DeviceCount, Row.DeviceNumbersCollection);

                    // ПРИХОД ЗА ПЕРИОД //
                    // Занесение данных в отчет: Приход из производства (новые приборы)
                    SplashScreenManager.Default.SetWaitFormDescription("Идет занесение данных о поступлении новых приборов из производства за период...");
                    foreach (TransferCountByParty Row in NewReceiptFromProductionByParty)
                        ReportDocument.AddValue(Row.DeviceType, Row.DeviceParty, Columns.NewReceiptFromProduction, Row.DeviceCount, Row.DeviceNumbersCollection);
                    // Занесение данных в отчет: Приход из производства (повторная передача)
                    SplashScreenManager.Default.SetWaitFormDescription("Идет занесение данных о повторном поступлении приборов из производства за период...");
                    foreach (TransferCountByParty Row in RepeatReceiptFromProductionByParty)
                        ReportDocument.AddValue(Row.DeviceType, Row.DeviceParty, Columns.RepeatReceiptFromProduction, Row.DeviceCount, Row.DeviceNumbersCollection);
                    // Занесение данных в отчет: Возврат на склад с выставок
                    SplashScreenManager.Default.SetWaitFormDescription("Идет занесение данных о возврате приборов на склад с выставок за период...");
                    foreach (TransferCountByParty Row in ReturnFromExpositionByParty)
                        ReportDocument.AddValue(Row.DeviceType, Row.DeviceParty, Columns.ReturnFromExposition, Row.DeviceCount, Row.DeviceNumbersCollection);
                    // Занесение данных в отчет: Возврат на склад с сертификации
                    SplashScreenManager.Default.SetWaitFormDescription("Идет занесение данных о возврате приборов на склад с сертификации за период...");
                    foreach (TransferCountByParty Row in ReturnFromCertificationByParty)
                        ReportDocument.AddValue(Row.DeviceType, Row.DeviceParty, Columns.ReturnFromCertification, Row.DeviceCount, Row.DeviceNumbersCollection);
                    // Занесение данных в отчет: Возврат на склад с испытаний
                    SplashScreenManager.Default.SetWaitFormDescription("Идет занесение данных о возврате приборов на склад с испытаний за период...");
                    foreach (TransferCountByParty Row in ReturnFromTestingByParty)
                        ReportDocument.AddValue(Row.DeviceType, Row.DeviceParty, Columns.ReturnFromTesting, Row.DeviceCount, Row.DeviceNumbersCollection);
                    // Занесение данных в отчет: Возврат на склад с тест-драйва
                    SplashScreenManager.Default.SetWaitFormDescription("Идет занесение данных о возврате приборов на склад с тест-драйва за период...");
                    foreach (TransferCountByParty Row in ReturnFromTestDriveByParty)
                        ReportDocument.AddValue(Row.DeviceType, Row.DeviceParty, Columns.ReturnFromTestDrive, Row.DeviceCount, Row.DeviceNumbersCollection);

                    // РАСХОД ЗА ПЕРИОД //
                    // Занесение данных в отчет: Возврат со склада в производство
                    SplashScreenManager.Default.SetWaitFormDescription("Идет занесение данных о возврате приборов со склада в производство за период...");
                    foreach (TransferCountByParty Row in ReturnFromWarehouseToProductionByParty)
                        ReportDocument.AddValue(Row.DeviceType, Row.DeviceParty, Columns.ReturnFromWarehouseToProduction, Row.DeviceCount, Row.DeviceNumbersCollection);
                    // Занесение данных в отчет: Выдача со склада на выставки
                    SplashScreenManager.Default.SetWaitFormDescription("Идет занесение данных о выдаче приборов со склада на выставки за период...");
                    foreach (TransferCountByParty Row in DeliveryToExpositionByParty)
                        ReportDocument.AddValue(Row.DeviceType, Row.DeviceParty, Columns.DeliveryToExposition, Row.DeviceCount, Row.DeviceNumbersCollection);
                    // Занесение данных в отчет: Выдача со склада на сертификацию
                    SplashScreenManager.Default.SetWaitFormDescription("Идет занесение данных о выдаче приборов со склада на сертификацию за период...");
                    foreach (TransferCountByParty Row in DeliveryToCertificationByParty)
                        ReportDocument.AddValue(Row.DeviceType, Row.DeviceParty, Columns.DeliveryToCertification, Row.DeviceCount, Row.DeviceNumbersCollection);
                    // Занесение данных в отчет: Выдача со склада на испытания
                    SplashScreenManager.Default.SetWaitFormDescription("Идет занесение данных о выдаче приборов со склада на испытания за период...");
                    foreach (TransferCountByParty Row in DeliveryToTestingByParty)
                        ReportDocument.AddValue(Row.DeviceType, Row.DeviceParty, Columns.DeliveryToTesting, Row.DeviceCount, Row.DeviceNumbersCollection);
                    // Занесение данных в отчет: Передача на тест-драйв
                    SplashScreenManager.Default.SetWaitFormDescription("Идет занесение данных о передаче приборов со склада на тест-драйв за период...");
                    foreach (TransferCountByParty Row in DeliveryToTestDriveByParty)
                        ReportDocument.AddValue(Row.DeviceType, Row.DeviceParty, Columns.DeliveryToTestDrive, Row.DeviceCount, Row.DeviceNumbersCollection);
                    // Занесение данных в отчет: Отгрузка новых приборов
                    SplashScreenManager.Default.SetWaitFormDescription("Идет занесение данных об отгрузке приборов за период...");
                    foreach (TransferCountByParty Row in DeliveryNewDevicesByParty)
                        ReportDocument.AddValue(Row.DeviceType, Row.DeviceParty, Columns.DeliveryNewDevices, Row.DeviceCount, Row.DeviceNumbersCollection);

                    // ОСТАТОК НА КОНЕЦ ПЕРИОДА //
                    // Занесение данных в отчет: На складе готовой продукции
                    SplashScreenManager.Default.SetWaitFormDescription("Идет занесение данных об остатке приборов на складе на конец периода...");
                    foreach (TransferCountByParty Row in InWarehouseAtEndingByParty)
                        ReportDocument.AddValue(Row.DeviceType, Row.DeviceParty, Columns.InWarehouseAtEndingPeriod, Row.DeviceCount, Row.DeviceNumbersCollection);
                    // Занесение данных в отчет: На выставках
                    SplashScreenManager.Default.SetWaitFormDescription("Идет занесение данных об остатке приборов на выставках на конец периода...");
                    foreach (TransferCountByParty Row in InExpositionAtEndingByParty)
                        ReportDocument.AddValue(Row.DeviceType, Row.DeviceParty, Columns.InExpositionAtEndingPeriod, Row.DeviceCount, Row.DeviceNumbersCollection);
                    // Занесение данных в отчет: На сертификации
                    SplashScreenManager.Default.SetWaitFormDescription("Идет занесение данных об остатке приборов на сертификации на конец периода...");
                    foreach (TransferCountByParty Row in InCertificationAtEndingByParty)
                        ReportDocument.AddValue(Row.DeviceType, Row.DeviceParty, Columns.InCertificationAtEndingPeriod, Row.DeviceCount, Row.DeviceNumbersCollection);
                    // Занесение данных в отчет: На испытаниях
                    SplashScreenManager.Default.SetWaitFormDescription("Идет занесение данных об остатке приборов на испытаниях на конец периода...");
                    foreach (TransferCountByParty Row in InTestingAtEndingByParty)
                        ReportDocument.AddValue(Row.DeviceType, Row.DeviceParty, Columns.InTestingAtEndingPeriod, Row.DeviceCount, Row.DeviceNumbersCollection);
                    // Занесение данных в отчет: На тест-драйве
                    SplashScreenManager.Default.SetWaitFormDescription("Идет занесение данных об остатке приборов на тест-драйве на конец периода...");
                    foreach (TransferCountByParty Row in InTestDriveAtEndingByParty)
                        ReportDocument.AddValue(Row.DeviceType, Row.DeviceParty, Columns.InTestDriveAtEndingPeriod, Row.DeviceCount, Row.DeviceNumbersCollection);

                    // Формирование отчета по заданным данным
                    SplashScreenManager.Default.SetWaitFormDescription("Форматирование отчета...");
                    ReportDocument.WriteData();
                    SplashScreenManager.CloseForm(false);

                    ReportDocument.ReportDocument.SaveAs(ReportDocument.TempFolder + ReportDocument.FileName);
                    ReportHelper.OpenReport(ReportDocument.TempFolder + ReportDocument.FileName);
                    break;
            }
        }
        /// <summary>
        /// Выполняет команду «Остаток комплектующих на складе ГП».
        /// </summary>
        public void GetBalanceComplete()
        {
            // Фильтры для отчета по остаткам компектующих:
            List<Guid> Devices = new List<Guid>();  // Приборы
            List<Guid> Completes = new List<Guid>();  // Сборочные узлы

            CompleteReportParameters NewForm = new CompleteReportParameters(Session, Context);
            DialogResult Result = NewForm.ShowDialog();
            if (Result == DialogResult.Cancel)
                return;
            else
            {
                if (NewForm.chooseDevices)
                {
                    Devices = NewForm.Devices.Select(r => r.Id).ToList();
                    MyMessageBox.Show("Получен перечень приборов: " + Devices.Select(r=>r.ToString()).Aggregate("; "));
                }

                if (NewForm.chooseCompletes)
                {
                    Completes = NewForm.Completes.Select(r => r.Id).ToList();
                    MyMessageBox.Show("Получен перечень комплектующих: " + Completes.Select(r => r.ToString()).Aggregate("; "));
                }
            }

            SplashScreenManager.CloseForm(false);
            SplashScreenManager.ShowForm(typeof(SKB.Base.Forms.MyWaitForm), true, true);
            SplashScreenManager.Default.SetWaitFormDescription("Начат подсчет остатков комплектующих...");

            // Узел справочника "Приборы и комплектующие"
            string DevicesAndCompleteID = "{DC3EE278-B3A2-493A-BE7A-74F08B6D57CB}";

            // Получение справочника "Остатки комплектующих"
            IBaseUniversalService baseUniversalService = Context.GetService<IBaseUniversalService>();
            BaseUniversal baseUniversal = Context.GetObject<BaseUniversal>(RefBaseUniversal.ID);
            if (!baseUniversal.ItemTypes.Any(r => r.GetObjectId() == new Guid(BalanceOfCompleteCard.BalanceOfCompleteDictionaryID)))
            {
                MyMessageBox.Show("Ошибка! Не найден справочник остатков комплектующих'.");
                return;
            }

            // Поиск записей справочника
            SearchQuery searchQuery = Session.CreateSearchQuery();
            searchQuery.CombineResults = ConditionGroupOperation.And;

            CardTypeQuery typeQuery = searchQuery.AttributiveSearch.CardTypeQueries.AddNew(DocsVision.BackOffice.CardLib.CardDefs.CardBaseUniversalItem.ID);
            SectionQuery sectionQuery = typeQuery.SectionQueries.AddNew(new Guid(BalanceOfCompleteCard.MainInfo.ID));
            sectionQuery.Operation = SectionQueryOperation.Or;

            sectionQuery.ConditionGroup.Operation = ConditionGroupOperation.Or;
            ConditionGroup ConditionGroup2 = sectionQuery.ConditionGroup.ConditionGroups.AddNew();
            ConditionGroup2.Operation = ConditionGroupOperation.And;
            ConditionGroup2.Conditions.AddNew(BalanceOfCompleteCard.MainInfo.EndDate, FieldType.DateTime, ConditionOperation.IsNotNull);

            searchQuery.Limit = 0;
            string query = searchQuery.GetXml();
            CardDataCollection CardBaseUniversalItems = Session.CardManager.FindCards(query);
            if (CardBaseUniversalItems.Count() == 0)
            {
                MyMessageBox.Show("Остатки комплектующих не подсчитаны.");
                return;
            }
            // Поиск записи с наибольшей датой
            DateTime MaxDate = CardBaseUniversalItems.Max(r => (DateTime)r.Sections[new Guid(BalanceOfCompleteCard.MainInfo.ID)].FirstRow.GetDateTime(BalanceOfCompleteCard.MainInfo.EndDate));
            CardData CardBaseUniversalItem = CardBaseUniversalItems.First(r => (DateTime)r.Sections[new Guid(BalanceOfCompleteCard.MainInfo.ID)].FirstRow.GetDateTime(BalanceOfCompleteCard.MainInfo.EndDate) == MaxDate);
            RowDataCollection BaseUniversalItemCollection = CardBaseUniversalItem.Sections[new Guid(BalanceOfCompleteCard.BalanceOfComplete.ID)].Rows;

            //Получение универсального справочника 4.5
            CardData UniversalDictionary = Session.CardManager.GetDictionaryData(RefUniversal.ID);
            RowData DevicesAndCompleteRow = UniversalDictionary.GetItemTypeRow(new Guid(DevicesAndCompleteID));
            RowDataCollection DevicesCollection = DevicesAndCompleteRow.ChildRows;

            List<CurrentBalanceComplete> CurrentBalanceCompleteCollection = new List<CurrentBalanceComplete>();

            // Построение перечня комплектующих
            if (!Completes.IsNull() && Completes.Count > 0)
            {
                foreach (RowData CurrentDevice in DevicesCollection)
                {
                    RowData DeviceRow = DevicesAndCompleteRow.ChildSections[RefUniversal.Item.ID].Rows.First(r => r.GetString("Name") == CurrentDevice.GetString("Name"));
                    if (DeviceRow.IsNull())
                    {
                        MyMessageBox.Show("Не найден прибор в справочнике Приборов и комплектующих.");
                        return;
                    }
                    foreach (RowData CurrentComplete in CurrentDevice.ChildSections[RefUniversal.Item.ID].Rows)
                    {
                        if (!UniversalDictionary.GetItemPropertyValue(CurrentComplete.Id, "Показывать в отчете").IsNull() && (bool)UniversalDictionary.GetItemPropertyValue(CurrentComplete.Id, "Показывать в отчете") == true &&
                            !UniversalDictionary.GetItemPropertyValue(CurrentComplete.Id, "Код СКБ").IsNull() && Completes.Any(g => g.Equals(UniversalDictionary.GetItemPropertyValue(CurrentComplete.Id, "Код СКБ").ToGuid())))
                        {
                            RowData BalanceReportRow = BaseUniversalItemCollection.FirstOrDefault(r => r.GetGuid(BalanceOfCompleteCard.BalanceOfComplete.CompleteID).Equals(CurrentComplete.Id)
                            && r.GetGuid(BalanceOfCompleteCard.BalanceOfComplete.AllocationID) == Allocation.InWarehouse);
                            CurrentBalanceCompleteCollection.Add(new CurrentBalanceComplete(DeviceRow.Id, CurrentComplete.Id,
                                UniversalDictionary.GetItemPropertyValue(CurrentComplete.Id, "Код СКБ").ToGuid(),
                                BalanceReportRow.IsNull() ? 0 : (int)BalanceReportRow.GetInt32(BalanceOfCompleteCard.BalanceOfComplete.EndCount),
                                0, 0, BalanceReportRow.IsNull() ? 0 : (int)BalanceReportRow.GetInt32(BalanceOfCompleteCard.BalanceOfComplete.EndCount)));
                        }
                    }
                }
            }
            else
            {
                if (!Devices.IsNull() && Devices.Count > 0)
                {
                    IEnumerable<RowData> devicesCollection = DevicesCollection.Where(r => Devices.Any(g => UniversalDictionary.GetItemName(g) == r.GetString("Name")));
                    if (!devicesCollection.IsNull() && devicesCollection.Count() > 0)
                    {
                        foreach (RowData CurrentDevice in devicesCollection)
                        {
                            RowData DeviceRow = DevicesAndCompleteRow.ChildSections[RefUniversal.Item.ID].Rows.First(r => r.GetString("Name") == CurrentDevice.GetString("Name"));
                            if (DeviceRow.IsNull())
                            {
                                MyMessageBox.Show("Не найден прибор в справочнике Приборов и комплектующих.");
                                return;
                            }
                            foreach (RowData CurrentComplete in CurrentDevice.ChildSections[RefUniversal.Item.ID].Rows)
                            {
                                if (!UniversalDictionary.GetItemPropertyValue(CurrentComplete.Id, "Показывать в отчете").IsNull() && (bool)UniversalDictionary.GetItemPropertyValue(CurrentComplete.Id, "Показывать в отчете") == true)
                                {
                                    RowData BalanceReportRow = BaseUniversalItemCollection.FirstOrDefault(r => r.GetGuid(BalanceOfCompleteCard.BalanceOfComplete.CompleteID).Equals(CurrentComplete.Id)
                                    && r.GetGuid(BalanceOfCompleteCard.BalanceOfComplete.AllocationID) == Allocation.InWarehouse);
                                    CurrentBalanceCompleteCollection.Add(new CurrentBalanceComplete(DeviceRow.Id, CurrentComplete.Id,
                                        UniversalDictionary.GetItemPropertyValue(CurrentComplete.Id, "Код СКБ").ToGuid(),
                                        BalanceReportRow.IsNull() ? 0 : (int)BalanceReportRow.GetInt32(BalanceOfCompleteCard.BalanceOfComplete.EndCount),
                                        0, 0, BalanceReportRow.IsNull() ? 0 : (int)BalanceReportRow.GetInt32(BalanceOfCompleteCard.BalanceOfComplete.EndCount)));
                                }
                            }
                        }
                    }
                }
                else
                {
                    foreach (RowData CurrentDevice in DevicesCollection)
                    {
                        RowData DeviceRow = DevicesAndCompleteRow.ChildSections[RefUniversal.Item.ID].Rows.First(r => r.GetString("Name") == CurrentDevice.GetString("Name"));
                        if (DeviceRow.IsNull())
                        {
                            MyMessageBox.Show("Не найден прибор в справочнике Приборов и комплектующих.");
                            return;
                        }
                        foreach (RowData CurrentComplete in CurrentDevice.ChildSections[RefUniversal.Item.ID].Rows)
                        {
                            if (!UniversalDictionary.GetItemPropertyValue(CurrentComplete.Id, "Показывать в отчете").IsNull() && (bool)UniversalDictionary.GetItemPropertyValue(CurrentComplete.Id, "Показывать в отчете") == true)
                            {
                                RowData BalanceReportRow = BaseUniversalItemCollection.FirstOrDefault(r => r.GetGuid(BalanceOfCompleteCard.BalanceOfComplete.CompleteID).Equals(CurrentComplete.Id)
                                && r.GetGuid(BalanceOfCompleteCard.BalanceOfComplete.AllocationID) == Allocation.InWarehouse);
                                CurrentBalanceCompleteCollection.Add(new CurrentBalanceComplete(DeviceRow.Id, CurrentComplete.Id,
                                    UniversalDictionary.GetItemPropertyValue(CurrentComplete.Id, "Код СКБ").ToGuid(),
                                    BalanceReportRow.IsNull() ? 0 : (int)BalanceReportRow.GetInt32(BalanceOfCompleteCard.BalanceOfComplete.EndCount),
                                    0, 0, BalanceReportRow.IsNull() ? 0 : (int)BalanceReportRow.GetInt32(BalanceOfCompleteCard.BalanceOfComplete.EndCount)));
                            }
                        }
                    }
                }
            }
            SplashScreenManager.Default.SetWaitFormDescription("Идет поиск актов передачи комплектующих...");
            IEnumerable<CompleteTransferRow> Acts = ReportCompleteHelper.FindAct(Session, (DateTime)MaxDate, DateTime.Today, UniversalDictionary);
            SplashScreenManager.Default.SetWaitFormDescription("Всего найдено передач по актам: " + Acts.Count());

            SplashScreenManager.Default.SetWaitFormDescription("Идет поиск заданий на комплектацию...");
            IEnumerable<CompleteTransferRow> Tasks = ReportCompleteHelper.FindCompleteTasks(Session, (DateTime)MaxDate, DateTime.Today, UniversalDictionary);
            SplashScreenManager.Default.SetWaitFormDescription("Всего найдено передач по заданиям на комплектацию: " + Tasks.Count());

            SplashScreenManager.Default.SetWaitFormDescription("Идет объединение результатов поиска...");
            IEnumerable<CompleteTransferRow> AllTransfers = Acts.Union(Tasks);
            SplashScreenManager.Default.SetWaitFormDescription("Найдено передач: " + AllTransfers.Count());

            if (!Devices.IsNull() && Devices.Count() > 0)
                AllTransfers = AllTransfers.Where(r => Devices.Any(g => g.Equals(r.ParentDeviceID)));
            if (!Completes.IsNull() && Completes.Count() > 0)
                AllTransfers = AllTransfers.Where(r => Completes.Any(g => g.Equals(r.CompleteCodeID)));

            // ПРИХОД ЗА ПЕРИОД //
            // Приход из производства (новые приборы)
            SplashScreenManager.Default.SetWaitFormDescription("Определяется перечень новых комплектующих, переданных из производства...");
            IEnumerable<TransferCountByCompleteType> NewReceiptFromProductionByCompleteType = AllTransfers.StatisticsForPeriod((DateTime)MaxDate, DateTime.Today, CompleteTransferRow.Action.NewReceiptFromProduction, UniversalDictionary);
            // Приход из производства (повторная передача)
            SplashScreenManager.Default.SetWaitFormDescription("Определяется перечень комплектующих, повторно переданных из производства...");
            IEnumerable<TransferCountByCompleteType> RepeatReceiptFromProductionByCompleteType = AllTransfers.StatisticsForPeriod((DateTime)MaxDate, DateTime.Today, CompleteTransferRow.Action.RepeatReceiptFromProduction, UniversalDictionary);
            // Возврат на склад с выставок
            SplashScreenManager.Default.SetWaitFormDescription("Определяется перечень комплектующих, возвращенных на склад с выставок...");
            IEnumerable<TransferCountByCompleteType> ReturnFromExpositionByCompleteType = AllTransfers.StatisticsForPeriod((DateTime)MaxDate, DateTime.Today, CompleteTransferRow.Action.ReturnFromExposition, UniversalDictionary);
            // Возврат на склад с сертификации
            SplashScreenManager.Default.SetWaitFormDescription("Определяется перечень комплектующих, возвращенных на склад с сертификации...");
            IEnumerable<TransferCountByCompleteType> ReturnFromCertificationByCompleteType = AllTransfers.StatisticsForPeriod((DateTime)MaxDate, DateTime.Today, CompleteTransferRow.Action.ReturnFromCertification, UniversalDictionary);
            // Возврат на склад с испытаний
            SplashScreenManager.Default.SetWaitFormDescription("Определяется перечень комплектующих, возвращенных на склад с испытаний...");
            IEnumerable<TransferCountByCompleteType> ReturnFromTestingByCompleteType = AllTransfers.StatisticsForPeriod((DateTime)MaxDate, DateTime.Today, CompleteTransferRow.Action.ReturnFromTesting, UniversalDictionary);
            // Возврат на склад с тест-драйва
            SplashScreenManager.Default.SetWaitFormDescription("Определяется перечень комплектующих, возвращенных на склад с тест-драйва...");
            IEnumerable<TransferCountByCompleteType> ReturnFromTestDriveByCompleteType = AllTransfers.StatisticsForPeriod((DateTime)MaxDate, DateTime.Today, CompleteTransferRow.Action.ReturnFromTestDrive, UniversalDictionary);
            // Возврат проданных комплектующих
            SplashScreenManager.Default.SetWaitFormDescription("Определяется перечень комплектующих, возвращенных клиентом...");
            IEnumerable<TransferCountByCompleteType> ReturnFromPayByCompleteType = AllTransfers.StatisticsForPeriod((DateTime)MaxDate, DateTime.Today, CompleteTransferRow.Action.Return, UniversalDictionary);


            // РАСХОД ЗА ПЕРИОД //
            // Возврат со склада в производство
            SplashScreenManager.Default.SetWaitFormDescription("Определяется перечень комплектующих, возвращенных со склада в производство...");
            IEnumerable<TransferCountByCompleteType> ReturnFromWarehouseToProductionByCompleteType = AllTransfers.StatisticsForPeriod((DateTime)MaxDate, DateTime.Today, CompleteTransferRow.Action.ReturnFromWarehouseToProduction, UniversalDictionary);
            // Выдача со склада на выставки
            SplashScreenManager.Default.SetWaitFormDescription("Определяется перечень комплектующих, выданных со склада на выставки...");
            IEnumerable<TransferCountByCompleteType> DeliveryToExpositionByCompleteType = AllTransfers.StatisticsForPeriod((DateTime)MaxDate, DateTime.Today, CompleteTransferRow.Action.DeliveryToExposition, UniversalDictionary);
            // Выдача со склада на сертификацию
            SplashScreenManager.Default.SetWaitFormDescription("Определяется перечень комплектующих, выданных со склада на сертификацию...");
            IEnumerable<TransferCountByCompleteType> DeliveryToCertificationByCompleteType = AllTransfers.StatisticsForPeriod((DateTime)MaxDate, DateTime.Today, CompleteTransferRow.Action.DeliveryToCertification, UniversalDictionary);
            // Выдача со склада на испытания
            SplashScreenManager.Default.SetWaitFormDescription("Определяется перечень комплектующих, выданных со склада на испытания...");
            IEnumerable<TransferCountByCompleteType> DeliveryToTestingByCompleteType = AllTransfers.StatisticsForPeriod((DateTime)MaxDate, DateTime.Today, CompleteTransferRow.Action.DeliveryToTesting, UniversalDictionary);
            // Передача на тест-драйв
            SplashScreenManager.Default.SetWaitFormDescription("Определяется перечень комплектующих, переданных со склада на тест-драйв...");
            IEnumerable<TransferCountByCompleteType> DeliveryToTestDriveByCompleteType = AllTransfers.StatisticsForPeriod((DateTime)MaxDate, DateTime.Today, CompleteTransferRow.Action.DeliveryToTestDrive, UniversalDictionary);
            // Отгрузка новых приборов
            SplashScreenManager.Default.SetWaitFormDescription("Определяется перечень отгруженных комплектующих...");
            IEnumerable<TransferCountByCompleteType> DeliveryNewDevicesByCompleteType = AllTransfers.StatisticsForPeriod((DateTime)MaxDate, DateTime.Today, CompleteTransferRow.Action.DeliveryNewDevices, UniversalDictionary);


            // ОПРЕДЕЛЕНИЕ ЗАРЕЗЕРВИРОВАННЫХ КОМПЛЕКТУЮЩИХ //
            SplashScreenManager.Default.SetWaitFormDescription("Идет поиск Договоров с зарезервированными комплектующими...");
            IEnumerable<ReservedCompleteRow> ReservedCompletes = ReportCompleteHelper.FindAccounts(Session, UniversalDictionary);
            SplashScreenManager.Default.SetWaitFormDescription("Всего найдено резервов по договорам: " + ReservedCompletes.Count());

            SplashScreenManager.Default.SetWaitFormDescription("Идет поиск Заданий на комплектацию с зарезервированными комплектующими...");
            ReservedCompletes.Union(ReportCompleteHelper.FindTaskCompleteWithoutShipment(Session, UniversalDictionary));
            SplashScreenManager.Default.SetWaitFormDescription("Всего найдено резервов по заданиям на комплектацию: " + ReservedCompletes.Count());

            if (!Devices.IsNull() && Devices.Count() > 0)
                ReservedCompletes = ReservedCompletes.Where(r => Devices.Any(g => g.Equals(r.DeviceID)));
            if (!Completes.IsNull() && Completes.Count() > 0)
                ReservedCompletes = ReservedCompletes.Where(r => Completes.Any(g => g.Equals(r.CodeSKB)));

            SplashScreenManager.Default.SetWaitFormDescription("Выполняется подсчет остатков...");

            CurrentBalanceCompleteCollection.ReceivedCalculation(NewReceiptFromProductionByCompleteType); // Новые комплектующие, переданные из производства
            CurrentBalanceCompleteCollection.ReceivedCalculation(RepeatReceiptFromProductionByCompleteType); // Новые комплектующие, повторно переданные из производства
            CurrentBalanceCompleteCollection.ReceivedCalculation(ReturnFromExpositionByCompleteType); // Комплектующие, возвращенные на склад с выставок
            CurrentBalanceCompleteCollection.ReceivedCalculation(ReturnFromCertificationByCompleteType); // Комплектующие, возвращенные на склад с сертификации
            CurrentBalanceCompleteCollection.ReceivedCalculation(ReturnFromTestingByCompleteType); // Комплектующие, возвращенные на склад с тест-драйва
            CurrentBalanceCompleteCollection.ReceivedCalculation(ReturnFromPayByCompleteType); // Комплектующие, возвращенные клиентом
            CurrentBalanceCompleteCollection.DescendedCalculation(ReturnFromWarehouseToProductionByCompleteType); // Комплектующие, возвращенные со склада в производство
            CurrentBalanceCompleteCollection.DescendedCalculation(DeliveryToExpositionByCompleteType); // Комплектующие, отправленные на выставку
            CurrentBalanceCompleteCollection.DescendedCalculation(DeliveryToCertificationByCompleteType); // Комплектующие, отправленные на сертификацию
            CurrentBalanceCompleteCollection.DescendedCalculation(DeliveryToTestingByCompleteType); // Комплектующие, отправленные на испытания
            CurrentBalanceCompleteCollection.DescendedCalculation(DeliveryToTestDriveByCompleteType); // Комплектующие, отправленные на тест-драйв
            CurrentBalanceCompleteCollection.DescendedCalculation(DeliveryNewDevicesByCompleteType); // Комплектующие отгруженные
            CurrentBalanceCompleteCollection.ReservedCalculation(ReservedCompletes); // Комплектующие зарезервированные

            // Формируем текст записи в лог
            string MyReport = "";
            foreach (CurrentBalanceComplete CurrentBalanceCompleteRow in CurrentBalanceCompleteCollection)
            {
                MyReport = MyReport + "\n" + UniversalDictionary.GetItemName(CurrentBalanceCompleteRow.DeviceID) + "\t" +
                    UniversalDictionary.GetItemName(CurrentBalanceCompleteRow.CompleteID) + "\t" + CurrentBalanceCompleteRow.StartBalance + "\t" +
                    CurrentBalanceCompleteRow.Received + "\t" + CurrentBalanceCompleteRow.Descended + "\t" + CurrentBalanceCompleteRow.Reserved + "\t" +
                    CurrentBalanceCompleteRow.EndBalance + "\t" +
                    "Приходные документы: " + CurrentBalanceCompleteRow.ReceivedDocuments.Distinct().Aggregate("; ") + "\t" +
                    "Расходные документы: " + CurrentBalanceCompleteRow.DescendedDocuments.Distinct().Aggregate("; ") + "\t" +
                    "Документы для резерва: " + CurrentBalanceCompleteRow.ReservedDocuments.Distinct().Aggregate("; ");
            }

            System.IO.File.Create("C:\\Tmp\\TestFile.txt").Close();
            System.IO.File.WriteAllText("C:\\Tmp\\TestFile.txt", MyReport);


            // СОЗДАНИЕ ОТЧЕТА //
            SplashScreenManager.Default.SetWaitFormDescription("Идет занесение данных в отчет...");
            ReportBalanceComplete ReportDocument = new ReportBalanceComplete(Session, MaxDate);
            ReportDocument.WriteData(CurrentBalanceCompleteCollection, UniversalDictionary);

            SplashScreenManager.CloseForm(false);

            ReportDocument.ReportDocument.SaveAs(ReportDocument.TempFolder + ReportDocument.FileName);
            ReportHelper.OpenReport(ReportDocument.TempFolder + ReportDocument.FileName);

        }
        /// <summary>
        /// Выполняет команду «Загрузить протоколы калибровки».
        /// </summary>
        public void LoadCalibrationDocuments()
        {
            string LoadFolderPath = @"\\folder\TMP\Калибровочная лаборатория\Протоколы калибровки";
            string ArchiveFolderPath = @"\\dv5\ARCDEL\Протоколы калибровки";
            string PassportFolderID = "{dde96a16-438b-46db-ac70-b477769e2124}";
            string TemplateCardID = "{c16d04ee-9f0d-4388-9579-cb82cd66c05e}";
            //try
            //{
                SplashScreenManager.ShowForm(typeof(MyWaitForm), true, true);
                SplashScreenManager.Default.SetWaitFormDescription("Идет загрузка...");

                ExtensionMethod method = Session.ExtensionManager.GetExtensionMethod("UploadExtension", "LoadCalibrationDocuments");
                method.Parameters.AddNew("LoadFolderPath", 0).Value = LoadFolderPath;
                method.Parameters.AddNew("ArchiveFolderPath", 0).Value = ArchiveFolderPath;
                method.Parameters.AddNew("PassportFolderID", 0).Value = PassportFolderID;
                method.Parameters.AddNew("TemplateCardID", 0).Value = TemplateCardID;

                SplashScreenManager.CloseForm(false);

                string Result = method.Execute().ToString();
                if (Result == "")
                { MyMessageBox.Show("Не обнаружено ни одного документа для загрузки.", "Результат исполнения:"); }
                else
                { MyMessageBox.Show(Result, "Результат исполнения:"); }
                return;
           // }
           // catch (Exception Ex)
           // {
           //     return;
           // }
        }
        /// <summary>
        /// Выполняет команду «Заполнить журнал условий калибровки».
        /// </summary>
        public void FillingCalibrationConditionsJournal(Int32 CabinetNumber)
        {
            string JournalName = "Журнал условий поверки";

                IBaseUniversalService baseUniversalService = Context.GetService<IBaseUniversalService>();
                BaseUniversal baseUniversal = Context.GetObject<BaseUniversal>(RefBaseUniversal.ID);
                if (!baseUniversal.ItemTypes.Any(r => r.Name == JournalName))
                {
                    MyMessageBox.Show("Ошибка! Не найден '" + JournalName + "'.");
                    return;
                }
                BaseUniversalItemType JournalItemType = baseUniversal.ItemTypes.First(r => r.Name == JournalName);
                if (!JournalItemType.Items.Any(r => r.Name == "Каб. №"+ CabinetNumber + ". Условия на " + DateTime.Today.ToShortDateString()))
                {
                    JournalForm NewJournalForm = new JournalForm(Session, Context, JournalItemType, CabinetNumber);
                    NewJournalForm.ShowDialog();
                }
                else
                {
                    BaseUniversalItem CurrentConditionsItem = JournalItemType.Items.First(r => r.Name == "Каб. №" + CabinetNumber + ". Условия на " + DateTime.Today.ToShortDateString());
                    JournalForm NewJournalForm = new JournalForm(Session, Context, JournalItemType, CurrentConditionsItem, CabinetNumber);
                    NewJournalForm.ShowDialog();
                }
        }
        #endregion

        #region Methods
        /// <summary>
        /// Возвращает имя расширения Навигатора.
        /// </summary>
        /// <param name="ExtensionType">Тип расширения Навигатора.</param>
        /// <returns></returns>
        protected override String GetExtensionName (NavExtensionTypes ExtensionType)
        {
            return "Расширения навигатора";
        }
        /// <summary>
        /// Возвращает поддерживаемые типы расширений Навигатора.
        /// </summary>
        protected override NavExtensionTypes SupportedTypes
        {
            get
            {
                return NavExtensionTypes.Command;
            }
        }
        /// <summary>
        /// Создаёт команды этого расширения Навигатора.
        /// </summary>
        /// <returns></returns>
        protected override IEnumerable<NavCommand> CreateCommands ()
        {
            List<NavCommand> NavCommands = new List<NavCommand>();
            NavCommands.Add(new NavCommand()
            {
                CommandType = NavCommandTypes.ToolBar | NavCommandTypes.Picture,
                Name = Command_Name_StartListsOfDocuments,
                Description = Command_Description_StartListsOfDocuments,
                Icon = Resources.StartListsOfDocuments
            });
            NavCommands.Add(new NavCommand()
            {
                CommandType = NavCommandTypes.ContextMenuFolder | NavCommandTypes.Picture,
                Name = Command_Name_StartListsOfDocuments_Folder,
                Description = Command_Description_StartListsOfDocuments,
                Icon = Resources.StartListsOfDocuments
            });
            NavCommands.Add(new NavCommand()
            {
                CommandType = NavCommandTypes.ToolBar | NavCommandTypes.Picture,
                Name = Command_Name_ApproveListsOfDocuments,
                Description = Command_Description_ApproveListsOfDocuments,
                Icon = Resources.ApproveListsOfDocuments
            });
            NavCommands.Add(new NavCommand()
            {
                CommandType = NavCommandTypes.ContextMenuFolder | NavCommandTypes.Picture,
                Name = Command_Name_ApproveListsOfDocuments_Folder,
                Description = Command_Description_ApproveListsOfDocuments,
                Icon = Resources.ApproveListsOfDocuments
            });
            NavCommands.Add(new NavCommand()
            {
                CommandType = NavCommandTypes.ContextMenuCard | NavCommandTypes.Picture,
                Name = Command_Name_DeleteCardAndFiles,
                Description = Command_Description_DeleteCardAndFiles,
                Icon = Resources.DeleteCardAndFiles
            });
            NavCommands.Add(new NavCommand()
            {
                CommandType = NavCommandTypes.ContextMenuCard | NavCommandTypes.Picture,
                Name = Command_Name_SendToAgreement,
                Description = Command_Description_SendToAgreement,
                Icon = Resources.SendToAgreement
            });
            NavCommands.Add(new NavCommand()
            {
                CommandType = NavCommandTypes.ToolBar | NavCommandTypes.Picture,
                Name = Command_Name_ReportCalibrationLaboratory,
                Description = Command_Description_ReportCalibrationLaboratory,
                Icon = Resources.ReportCalibrationLaboratory
            });
            NavCommands.Add(new NavCommand()
            {
                CommandType = NavCommandTypes.ToolBar | NavCommandTypes.Picture,
                Name = Command_Name_ReportWarehouse,
                Description = Command_Description_ReportWarehouse,
                Icon = Resources.ReportWarehouse
            });
            NavCommands.Add(new NavCommand()
            {
                CommandType = NavCommandTypes.ToolBar | NavCommandTypes.Picture,
                Name = Command_Name_LoadCalibrationDocuments,
                Description = Command_Description_LoadCalibrationDocuments,
                Icon = Resources.LoadCalibrationDocuments
            });
            NavCommands.Add(new NavCommand()
            {
                CommandType = NavCommandTypes.ToolBar | NavCommandTypes.Picture,
                Name = Command_Name_FillingVerifyConditionsJournal238,
                Description = Command_Description_FillingVerifyConditionsJournal238,
                Icon = Resources.FillingCalibrationConditionsJournal
            });
            NavCommands.Add(new NavCommand()
            {
                CommandType = NavCommandTypes.ToolBar | NavCommandTypes.Picture,
                Name = Command_Name_FillingVerifyConditionsJournal237,
                Description = Command_Description_FillingVerifyConditionsJournal237,
                Icon = Resources.FillingCalibrationConditionsJournal
            });
            NavCommands.Add(new NavCommand()
            {
                CommandType = NavCommandTypes.ToolBar | NavCommandTypes.Picture,
                Name = Command_Name_CreateTaskManager,
                Description = Command_Description_CreateTaskManager,
                Icon = Session.CardManager.CardTypes[DocsVision.BackOffice.CardLib.CardDefs.CardTask.ID].Icon
            });
            NavCommands.Add(new NavCommand()
            {
                CommandType = NavCommandTypes.ContextMenuCard | NavCommandTypes.Picture,
                Name = Command_Name_CreateTaskOnSub,
                Description = Command_Description_CreateTaskOnSub,
                Icon = Session.CardManager.CardTypes[DocsVision.BackOffice.CardLib.CardDefs.CardTask.ID].Icon
            });
            NavCommands.Add(new NavCommand()
            {
                CommandType = NavCommandTypes.ToolBar | NavCommandTypes.Picture,
                Name = Command_Name_BalanceComplete,
                Description = Command_Description_BalanceComplete,
                Icon = Resources.ReportWarehouse
            });

            return NavCommands;
        }
        /// <summary>
        /// Возвращает состояние команды расширения Навигатора.
        /// </summary>
        /// <param name="Command">Команда расширения Навигатора.</param>
        /// <param name="Context">Среда команды расширения Навигатора.</param>
        /// <returns></returns>
        protected override NavCommandStatus QueryCommandStatus (NavCommand Command, NavCommandContext Context)
        {
            NavCommandStatus Status;

            switch (Command.Name)
            {
                case Command_Name_StartListsOfDocuments:
                    Status = NavCommandStatus.Supported | NavCommandStatus.Enabled;
                    break;
                case Command_Name_StartListsOfDocuments_Folder:
                    Status = Context.FolderId == ListProcesses.RefListofDocsFolder ?  NavCommandStatus.Supported | NavCommandStatus.Enabled : NavCommandStatus.None;
                    break;
                case Command_Name_ApproveListsOfDocuments:
                    Status = NavCommandStatus.Supported | NavCommandStatus.Enabled;
                    break;
                case Command_Name_ApproveListsOfDocuments_Folder:
                    Status = Context.FolderId == ListProcesses.RefListofDocsFolder_Archived ? NavCommandStatus.Supported | NavCommandStatus.Enabled : NavCommandStatus.None;
                    break;
                case Command_Name_DeleteCardAndFiles:
                    Status = DeleteCardAndFiles(true, Context.Selection);
                    break;
                case Command_Name_SendToAgreement:
                    Status = SendToAgreement(true, Context.Selection);
                    break;
                case Command_Name_ReportCalibrationLaboratory:
                    Status = NavCommandStatus.Supported | NavCommandStatus.Enabled;
                    break;
                case Command_Name_ReportWarehouse:
                    Status = NavCommandStatus.Supported | NavCommandStatus.Enabled;
                    break;
                case Command_Name_LoadCalibrationDocuments:
                    Status = NavCommandStatus.Supported | NavCommandStatus.Enabled;
                    break;
                case Command_Name_FillingVerifyConditionsJournal238:
                    Status = NavCommandStatus.Supported | NavCommandStatus.Enabled;
                    break;
                case Command_Name_FillingVerifyConditionsJournal237:
                    Status = NavCommandStatus.Supported | NavCommandStatus.Enabled;
                    break;
                case Command_Name_BalanceComplete:
                    Status = NavCommandStatus.Supported | NavCommandStatus.Enabled;
                    break;
                case Command_Name_CreateTaskManager:
                    var group0 = this.Context.FindObject<StaffGroup>(new DocsVision.Platform.ObjectModel.Search.QueryObject("Name", "DocsVision Administrators"));
                    var group1 = this.Context.FindObject<StaffGroup>(new DocsVision.Platform.ObjectModel.Search.QueryObject("Name", "DV Managers"));
                    var user = StaffService.GetCurrentEmployee();
                    if (group0 != null && StaffService.IsEmployeeInGroup(user, group0, true) || group1 != null && StaffService.IsEmployeeInGroup(user, group1, true))
                        Status = NavCommandStatus.Supported | NavCommandStatus.Enabled;
                    else
                        Status = NavCommandStatus.None;
                    break;
                case Command_Name_CreateTaskOnSub:
                    if (Context.Selection.Length != 1)
                    {
                        Status = NavCommandStatus.None;
                        break;
                    }
                    else
                    {
                        Task selected = this.Context.GetObject<Task>(Context.Selection[0]);

                        if (selected == null)
                            Status = NavCommandStatus.None;
                        else if (selected.SystemInfo.CardKind == null || (selected.SystemInfo.CardKind.Name != "Подэтап" && selected.SystemInfo.CardKind.Name != "Задание-задание"))
                            Status = NavCommandStatus.None;
                        else
                        {
                            StatesOperation operation = StateService.GetOperations(selected.SystemInfo.CardKind).FirstOrDefault(o => o.DefaultName == "Create child task");
                            if (operation != null && StateService.IsOperationAllowedFull(operation, selected))
                                Status = NavCommandStatus.Supported | NavCommandStatus.Enabled;
                            else
                                Status = NavCommandStatus.Supported;
                        }
                        break;
                    }
                default: Status = NavCommandStatus.None; break;
            }
            return Status;
        }
        /// <summary>
        /// Вызывает команду расширения Навигатора.
        /// </summary>
        /// <param name="Command">Команда расширения Навигатора.</param>
        /// <param name="Context">Среда команды расширения Навигатора.</param>
        protected override void InvokeCommand (NavCommand Command, NavCommandContext Context)
        {
            switch (Command.Name)
            {
                case Command_Name_StartListsOfDocuments:
                    StartListsOfDocuments();
                    break;
                case Command_Name_StartListsOfDocuments_Folder:
                    StartListsOfDocuments();
                    break;
                case Command_Name_ApproveListsOfDocuments:
                    ApproveListsOfDocuments();
                    break;
                case Command_Name_ApproveListsOfDocuments_Folder:
                    ApproveListsOfDocuments();
                    break;
                case Command_Name_DeleteCardAndFiles:
                    DeleteCardAndFiles(false, Context.Selection);
                    break;
                case Command_Name_SendToAgreement:
                    SendToAgreement(false, Context.Selection);
                    break;
                case Command_Name_ReportCalibrationLaboratory:
                    GetReportCalibrationLaboratory();
                    break;
                case Command_Name_ReportWarehouse:
                    GetReportWarehouse();
                    break;
                case Command_Name_LoadCalibrationDocuments:
                    LoadCalibrationDocuments();
                    break;
                case Command_Name_FillingVerifyConditionsJournal238:
                    FillingCalibrationConditionsJournal(238);
                    break;
                case Command_Name_FillingVerifyConditionsJournal237:
                    FillingCalibrationConditionsJournal(237);
                    break;
                case Command_Name_CreateTaskManager:
                    new SKB.TaskExtension.ObjectModel.Service(this.Context).CreateTaskFromManagerForm(CardFrame.CardHost);
                    break;
                case Command_Name_BalanceComplete:
                    GetBalanceComplete();
                    break;
                case Command_Name_CreateTaskOnSub:
                    if (Context.Selection.Length != 1)
                        break;

                    Task selected = this.Context.GetObject<Task>(Context.Selection[0]);
                    if (selected == null)
                        break;

                    var stateService = this.Context.GetService<IStateService>();
                    StatesOperation operation = stateService.GetOperations(selected.SystemInfo.CardKind).FirstOrDefault(o => o.DefaultName == "Create child task");
                    if (operation != null && stateService.IsOperationAllowedFull(operation, selected))
                        new SKB.TaskExtension.ObjectModel.Service(this.Context).CreateTaskFromManagerForm(CardFrame.CardHost, Context.Selection[0]);

                    break;
                default:
                    break;
            }
        }
        #endregion
    }
}
