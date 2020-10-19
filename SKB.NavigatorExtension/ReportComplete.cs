using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.ComponentModel;
using ClosedXML.Excel;
using DocsVision.Platform.ObjectManager;
using DocsVision.Platform.ObjectModel;
using SKB.Base.Ref;
using SKB.Base;
using SKB.PaymentAndShipment.Forms.AccountCard;
using DocsVision.BackOffice.ObjectModel.Services;
using DocsVision.BackOffice.ObjectModel;
using System.Diagnostics;
using DocsVision.Platform.ObjectManager.SearchModel;
using DocsVision.TakeOffice.Cards.Constants;
using DocsVision.Platform.ObjectManager.Metadata;
using System.IO;
using RKIT.MyMessageBox;



namespace SKB.NavigatorExtension
{
    /// <summary>
    /// Карточка строки справочника "Остаток комплектующих"
    /// </summary>
    public class BalanceOfCompleteCard
    {
        /// <summary>
        /// Идентификатор типа в конструкторе справочников "Остаток комплектующих"
        /// </summary>
        public static string BalanceOfCompleteDictionaryID = "{3D7ACC98-0FA2-442A-9B99-DFB29C017D53}";
        public static class BalanceOfComplete
        {
            public static string Name = "BalanceOfComplete";
            public static string ID = "{E6DB53B7-7677-4978-8562-6B17917516A6}";
            public static string CompleteID = "CompleteID";
            public static string StartCount = "StartCount";
            public static string EndCount = "EndCount";
            public static string Arrival = "Arrival";
            public static string Consumption = "Consumption";
            public static string CodeID = "CodeID";
            public static string AllocationID = "AllocationID";
            public static string DeviceID = "DeviceID";
        }
        public static class MainInfo
        {
            public static string Name = "MainInfo";
            public static string ID = "{F5641A7E-83AF-4C20-9C60-EA2973C4F135}";
            public static string StartDate = "StartDate";
            public static string EndDate = "EndDate";
        }
    }
    /// <summary>
    /// Размещение приборов/комплектующих.
    /// </summary>
    public class Allocation
    {
        /// <summary>
        /// На складе.
        /// </summary>
        public static Guid InWarehouse = new Guid("{DE1C813B-D3A7-40B1-90AF-A6BFEFDB5E64}");
        /// <summary>
        /// На выставке.
        /// </summary>
        public static Guid InExposition = new Guid("{CEB0BCE1-0AEF-4E58-81F1-A4C34B65FE8B}");
        /// <summary>
        /// На сертификации.
        /// </summary>
        public static Guid InCertification = new Guid("{222615A0-685D-44E2-AF13-5B6056D43AE0}");
        /// <summary>
        /// На испытаниях.
        /// </summary>
        public static Guid InTesting = new Guid("{7F2B3AF6-B312-44C7-9280-A2340C010C9E}");
        /// <summary>
        /// На тест-драйве.
        /// </summary>
        public static Guid InTestDrive = new Guid("{0C75617F-E4D1-4CF5-941C-826CFE16A591}");
    };

    /// <summary>
    /// Экземпляр карточки строки справочника "Остаток комплектующих"
    /// </summary>
    public class BalanceOfCompleteItem
    {
        public DateTime startDate;
        public DateTime endDate;
        public List<BalanceOfCompleteRowItem> balanceOfCompleteTable;
        public CardData balanceOfCompleteCardData;

        public BalanceOfCompleteItem(CardData ReportCard)
        {
            balanceOfCompleteCardData = ReportCard;
            SectionData ValiditySection = ReportCard.Sections[new Guid(BalanceOfCompleteCard.MainInfo.ID)];
            SectionData BalanceOfCompleteSection = ReportCard.Sections[new Guid(BalanceOfCompleteCard.BalanceOfComplete.ID)];

            startDate = (DateTime)ValiditySection.FirstRow.GetDateTime(BalanceOfCompleteCard.MainInfo.StartDate);
            endDate = (DateTime)ValiditySection.FirstRow.GetDateTime(BalanceOfCompleteCard.MainInfo.EndDate);

            balanceOfCompleteTable = new List<BalanceOfCompleteRowItem>();
            foreach (RowData Row in BalanceOfCompleteSection.Rows)
            {
                balanceOfCompleteTable.Add(new BalanceOfCompleteRowItem(
                    Row.GetGuid(BalanceOfCompleteCard.BalanceOfComplete.CompleteID).ToGuid(),
                    (int)Row.GetInt32(BalanceOfCompleteCard.BalanceOfComplete.StartCount),
                    (int)Row.GetInt32(BalanceOfCompleteCard.BalanceOfComplete.EndCount),
                    (int)Row.GetInt32(BalanceOfCompleteCard.BalanceOfComplete.Arrival),
                    (int)Row.GetInt32(BalanceOfCompleteCard.BalanceOfComplete.Consumption),
                    Row.GetGuid(BalanceOfCompleteCard.BalanceOfComplete.CodeID).ToGuid(),
                    Row.GetGuid(BalanceOfCompleteCard.BalanceOfComplete.AllocationID).ToGuid(),
                    Row.GetGuid(BalanceOfCompleteCard.BalanceOfComplete.DeviceID).ToGuid()
                    ));
            }
        }
    }
    /// <summary>
    /// Строка справочника "Остаток комплектующих"
    /// </summary>
    public class BalanceOfCompleteRowItem
    {
        public Guid completeID;
        public int startCount;
        public int endCount;
        public int arrival;
        public int consumption;
        public Guid codeID;
        public Guid allocationID;
        public Guid deviceID;

        public BalanceOfCompleteRowItem(Guid CompleteID, int StartCount, int EndCount, int Arrival, int Consumption, Guid CodeID, Guid AllocationID, Guid DeviceID)
        {
            completeID = CompleteID;
            startCount = StartCount;
            endCount = EndCount;
            arrival = Arrival;
            consumption = Consumption;
            codeID = CodeID;
            allocationID = AllocationID;
            deviceID = DeviceID;
        }
    }

    /// <summary>
    /// Столбцы отчета.
    /// </summary>
    public enum CompleteColumns
    {
        /// <summary>
        /// Прибор/Партия.
        /// </summary>
        [Description("Прибор/Партия")]
        DeviceOrParty = 1,
        /// <summary>
        /// На складе (на начало периода).
        /// </summary>
        [Description("На складе (на начало месяца)")]
        InWarehouseAtBeginningPeriod = 2,
        /// <summary>
        /// На выставке (на начало периода).
        /// </summary>
        [Description("На выставке (на начало месяца)")]
        InExpositionAtBeginningPeriod = 3,
        /// <summary>
        /// На сертификации (на начало периода).
        /// </summary>
        [Description("На сертификации (на начало месяца)")]
        InCertificationAtBeginningPeriod = 4,
        /// <summary>
        /// На испытаниях (на начало периода).
        /// </summary>
        [Description("На испытаниях (на начало месяца)")]
        InTestingAtBeginningPeriod = 5,
        /// <summary>
        /// На тест-драйве (на начало периода).
        /// </summary>
        [Description("На тест-драйве (на начало месяца)")]
        InTestDriveAtBeginningPeriod = 6,
        /// <summary>
        /// Приход из производства (новые приборы).
        /// </summary>
        [Description("Приход из производства (новые приборы)")]
        NewReceiptFromProduction = 7,
        /// <summary>
        /// Приход из производства (повторная передача).
        /// </summary>
        [Description("Приход из производства (повторная передача)")]
        RepeatReceiptFromProduction = 8,
        /// <summary>
        /// Возврат на склад с выставок.
        /// </summary>
        [Description("Возврат на склад с выставок")]
        ReturnFromExposition = 9,
        /// <summary>
        /// Возврат на склад с сертификации.
        /// </summary>
        [Description("Возврат на склад с сертификации")]
        ReturnFromCertification = 10,
        /// <summary>
        /// Возврат на склад с испытаний.
        /// </summary>
        [Description("Возврат на склад с испытаний")]
        ReturnFromTesting = 11,
        /// <summary>
        /// Возврат на склад с тест-драйва.
        /// </summary>
        [Description("Возврат на склад с тест-драйва")]
        ReturnFromTestDrive = 12,
        /// <summary>
        /// Возврат со склада в производство.
        /// </summary>
        [Description("Возврат со склада в производство")]
        ReturnFromWarehouseToProduction = 13,
        /// <summary>
        /// Выдача со склада на выставки.
        /// </summary>
        [Description("Выдача со склада на выставки")]
        DeliveryToExposition = 14,
        /// <summary>
        /// Выдача со склада на сертификацию.
        /// </summary>
        [Description("Выдача со склада на сертификацию")]
        DeliveryToCertification = 15,
        /// <summary>
        /// Выдача со склада на испытания.
        /// </summary>
        [Description("Выдача со склада на испытания")]
        DeliveryToTesting = 16,
        /// <summary>
        /// Передача на тест-драйв.
        /// </summary>
        [Description("Передача на тест-драйв")]
        DeliveryToTestDrive = 17,
        /// <summary>
        /// Отгружено новых приборов.
        /// </summary>
        [Description("Отгружено новых приборов")]
        DeliveryNewDevices = 18,
        /// <summary>
        /// На складе (на конец периода).
        /// </summary>
        [Description("На складе (на начало месяца)")]
        InWarehouseAtEndingPeriod = 19,
        /// <summary>
        /// На выставке (на конец периода).
        /// </summary>
        [Description("На выставке (на начало месяца)")]
        InExpositionAtEndingPeriod = 20,
        /// <summary>
        /// На сертификации (на конец периода).
        /// </summary>
        [Description("На сертификации (на начало месяца)")]
        InCertificationAtEndingPeriod = 21,
        /// <summary>
        /// На испытаниях (на конец периода).
        /// </summary>
        [Description("На испытаниях (на начало месяца)")]
        InTestingAtEndingPeriod = 22,
        /// <summary>
        /// На тест-драйве (на конец периода).
        /// </summary>
        [Description("На тест-драйве (на начало месяца)")]
        InTestDriveAtEndingPeriod = 23,
    };

    /// <summary>
    /// Столбцы отчета по остатку комплектующих.
    /// </summary>
    public enum BalanceCompleteColumns
    {
        /// <summary>
        /// Название комплектующего.
        /// </summary>
        [Description("Название комплектующего")]
        CompleteName = 1,
        /// <summary>
        /// Код СКБ.
        /// </summary>
        [Description("Код СКБ")]
        CodeSKB = 2,
        /// <summary>
        /// Остаток на...
        /// </summary>
        [Description("Остаток на ")]
        BalanceOn = 3,
        /// <summary>
        /// Приход.
        /// </summary>
        [Description("Приход")]
        Received = 4,
        /// <summary>
        /// Расход.
        /// </summary>
        [Description("Расход")]
        Descended = 5,
        /// <summary>
        /// Остаток на текущий момент.
        /// </summary>
        [Description("Остаток на текущий момент")]
        CurrentBalance = 6,
        /// <summary>
        /// Из них зарезервировано.
        /// </summary>
        [Description("Из них зарезервировано")]
        Reserved = 7,
        /// <summary>
        /// ИТОГО.
        /// </summary>
        [Description("ИТОГО:")]
        Total = 8
    };

    class ReportComplete
    {
        #region Properties
        /// <summary>
        /// Название файла отчета.
        /// </summary>
        public string FileName
        { get { return "Отчет по комплектующим с " + StartDate.ToShortDateString() + " по " + EndDate.ToShortDateString() + ".xlsx"; } }
        /// <summary>
        /// Документ.
        /// </summary>
        public XLWorkbook ReportDocument;
        /// <summary>
        /// Рабочий лист.
        /// </summary>
        public IXLWorksheet ReportWorksheet;
        /// <summary>
        /// Сессия DV.
        /// </summary>
        UserSession session;
        /// <summary>
        /// Объектный контекст.
        /// </summary>
        ObjectContext Context;
        /// <summary>
        /// Перечень позиций отчета.
        /// </summary>
        List<ReportWareHouseItem> ReportItems;
        //public static Guid ActTypeID = SKB.Base.MyHelper.RefType_ActOfTransfer;
        //public static Guid ServiceCardTypeID = RefServiceCard.ID;
        /// <summary>
        /// Шаблон отчета.
        /// </summary>
        private Guid TemplateID = new Guid("{5437DEB7-74B4-E611-BE1D-00155D11531B}");
        /// <summary>
        /// Временная папка.
        /// </summary>
        public string TempFolder = System.IO.Path.GetTempPath();
        /// <summary>
        /// Индексы столбцов.
        /// </summary>
        private int[] columnsIndex;
        /// <summary>
        /// Текущее количество строк в отчете.
        /// </summary>
        public int RowsCount = 6;
        /// <summary>
        /// Текст заголовка.
        /// </summary>
        const string HeaderText = "Отчёт по складу готовой продукции";
        /// <summary>
        /// Название рабочего листа.
        /// </summary>
        const string WorksheetsName = "Отчёт";
        /// <summary>
        /// Индексы столбцов.
        /// </summary>
        private int[] ColumnsIndex
        {
            get
            {
                if (columnsIndex.IsNull())
                {
                    Array EnumValues = Enum.GetValues(typeof(Columns));
                    columnsIndex = new int[EnumValues.Length];

                    for (int i = 0; i < EnumValues.Length; i++)
                        columnsIndex[i] = (int)EnumValues.GetValue(i);
                }
                return columnsIndex;
            }
        }
        /// <summary>
        /// Дата начала периода.
        /// </summary>
        DateTime StartDate;
        /// <summary>
        /// Дата окончания периода.
        /// </summary>
        DateTime EndDate;
        #endregion
        /// <summary>
        /// Конструктор отчета по комплектующим.
        /// </summary>
        /// <param name="session">Пользовательская сессия DV.</param>
        /// <param name="StartDate">Дата начала периода.</param>
        /// <param name="EndDate">Дата окончания периода.</param>
        public ReportComplete(UserSession session, DateTime StartDate, DateTime EndDate)
        {
            this.session = session;
            this.StartDate = StartDate;
            this.EndDate = EndDate;
            Context = session.CreateContext();
            ReportItems = new List<ReportWareHouseItem>();

            IDocumentService DocService = Context.GetService<IDocumentService>();
            DocsVision.BackOffice.ObjectModel.Document Template = Context.GetObject<DocsVision.BackOffice.ObjectModel.Document>(TemplateID.ToGuid());
            DocumentFile File = Template.Files.First();

            string FilePath = DocService.DownloadFile(File.FileId);

            ReportDocument = new XLWorkbook(FilePath);
            ReportWorksheet = ReportDocument.Worksheets.First(r => r.Name == WorksheetsName);

            //ReportDocument.SaveAs(TempFolder + FileName);
        }
        /// <summary>
        /// Добавить новую запись.
        /// </summary>
        /// <param name="DeviceType">Тип прибора.</param>
        /// <param name="Party">Название партии.</param>
        /// <param name="ColumnName">Название столбца.</param>
        /// <param name="Value">Значение.</param>
        /// <param name="Description">Примечание.</param>
        public void AddValue(string DeviceType, string Party, Columns ColumnName, int Value, string Description)
        {
            // Поиск записи отчета для искомой партии приборов (если запись отчета для искомой партии не найдена, то она создается)
            ReportWareHouseItem FindItem;
            if (ReportItems.Any(r => r.DeviceParty == Party))
            {
                FindItem = ReportItems.First(r => r.DeviceParty == Party);
            }
            else
            {
                FindItem = new ReportWareHouseItem(ColumnsIndex, DeviceType, Party);
                ReportItems.Add(FindItem);
            }
            // Заполнение данных
            ReportWarehouseValue FindValue = FindItem.Values.First(r => r.ColumnIndex == (int)ColumnName);
            FindValue.Value = Value;
            FindValue.Description = Description;
        }
        /// <summary>
        /// Запись данных в отчет.
        /// </summary>
        public void WriteData()
        {
            IEnumerable<ReportWareHouseItem> DeviceTypesItemCollection = ReportItems.GroupBy(r => r.DeviseType, (deviceType, items) => new ReportWareHouseItem(ColumnsIndex, deviceType, "")).OrderBy(r => r.DeviseType);
            foreach (ReportWareHouseItem Item in DeviceTypesItemCollection)
            {
                IEnumerable<ReportWareHouseItem> PartyItemCollection = ReportItems.Where(r => r.DeviseType == Item.DeviseType).OrderBy(r => r.DeviceParty, new PartyComparer());
                foreach (ReportWarehouseValue Value in Item.Values)
                {
                    Value.Value = PartyItemCollection.Select(r => r.Values.First(s => s.ColumnIndex == Value.ColumnIndex).Value).Aggregate(0, (start, next) => start + next);
                    Value.Description = PartyItemCollection.Select(r => r.Values.First(s => s.ColumnIndex == Value.ColumnIndex).Description).Aggregate(String.Empty, (start, next) => JoinDevices(", ", start, next));
                }

                AddDeviceRow(Item);
                foreach (ReportWareHouseItem PartyItem in PartyItemCollection)
                    AddPartyRow(PartyItem);
            }
            AddFinalRow();
            SetBorders(ReportWorksheet.Range(7, 1, RowsCount, 1), XLBorderStyleValues.Medium, XLColor.Black);
            SetBorders(ReportWorksheet.Range(7, 2, RowsCount, 6), XLBorderStyleValues.Medium, XLColor.Black);
            SetBorders(ReportWorksheet.Range(7, 7, RowsCount, 12), XLBorderStyleValues.Medium, XLColor.Black);
            SetBorders(ReportWorksheet.Range(7, 13, RowsCount, 18), XLBorderStyleValues.Medium, XLColor.Black);
            SetBorders(ReportWorksheet.Range(7, 19, RowsCount, 23), XLBorderStyleValues.Medium, XLColor.Black);

            SetBorders(ReportWorksheet.Range(7, 1, RowsCount, 23), XLBorderStyleValues.Medium, XLColor.Black);

            // Подписи ответственных сотрудников
            String ProductionManager = MyHelper.PositionName_PM;
            String SpecialistInSale = MyHelper.PositionName_SpecialistInSale;
            ReportWorksheet.Cell(RowsCount + 2, 2).SetValue(ProductionManager);
            ApplySubscribeStyle(ReportWorksheet.Cell(RowsCount + 2, 2));
            ReportWorksheet.Cell(RowsCount + 2, 10).SetValue(SKB.Base.MyHelper.GetEmployeeByPosition(Context, ProductionManager).GetDisplayString());
            ApplySubscribeStyle(ReportWorksheet.Cell(RowsCount + 2, 10));
            ReportWorksheet.Cell(RowsCount + 3, 2).SetValue(SpecialistInSale);
            ApplySubscribeStyle(ReportWorksheet.Cell(RowsCount + 3, 2));
            ReportWorksheet.Cell(RowsCount + 3, 10).SetValue(SKB.Base.MyHelper.GetEmployeeByPosition(Context, SpecialistInSale).GetDisplayString());
            ApplySubscribeStyle(ReportWorksheet.Cell(RowsCount + 3, 10));

            // Дата формирования
            ReportWorksheet.Cell(3, 1).SetValue("Дата формирования: " + DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToShortTimeString());
            ApplySubscribeStyle(ReportWorksheet.Cell(3, 1));

        }
        /// <summary>
        /// Добавление итоговой строки
        /// </summary>
        private void AddFinalRow()
        {
            ReportWareHouseItem FinalItem = new ReportWareHouseItem(ColumnsIndex, "Итого:", "");
            foreach (ReportWarehouseValue Value in FinalItem.Values)
                Value.Value = ReportItems.Select(r => r.Values.First(s => s.ColumnIndex == Value.ColumnIndex).Value).Aggregate(0, (start, next) => start + next);
            AddDeviceRow(FinalItem);
        }
        /// <summary>
        /// Добавление строки нвого типа прибора
        /// </summary>
        /// <param name="Item">Запись отчета.</param>
        private void AddDeviceRow(ReportWareHouseItem Item)
        {
            RowsCount++;
            // Прибор/партия

            ReportWorksheet.Cell(RowsCount, (int)Columns.DeviceOrParty).SetValue(Item.DeviseType);
            ApplySubHeaderRowStyle(ReportWorksheet.Cell(RowsCount, (int)Columns.DeviceOrParty), XLAlignmentHorizontalValues.Left);
            foreach (ReportWarehouseValue Value in Item.Values)
            {
                if (Value.Value != 0)
                {
                    ReportWorksheet.Cell(RowsCount, Value.ColumnIndex).SetValue(Value.Value);
                    ReportWorksheet.Cell(RowsCount, Value.ColumnIndex).Comment.AddText(Value.Description);

                }
                if (Value.ColumnIndex != 1)
                    ApplySubHeaderRowStyle(ReportWorksheet.Cell(RowsCount, Value.ColumnIndex), XLAlignmentHorizontalValues.Center);
            }
        }
        /// <summary>
        /// Добавление строки новой партии прибора
        /// </summary>
        /// <param name="Item">Запись отчета.</param>
        private void AddPartyRow(ReportWareHouseItem Item)
        {
            RowsCount++;

            ReportWorksheet.Cell(RowsCount, (int)Columns.DeviceOrParty).SetValue(Item.DeviceParty);
            ApplyEmptyPartyStyle(ReportWorksheet.Cell(RowsCount, (int)Columns.DeviceOrParty));

            foreach (ReportWarehouseValue Value in Item.Values)
            {
                if (Value.Value != 0)
                {
                    ReportWorksheet.Cell(RowsCount, Value.ColumnIndex).SetValue(Value.Value);
                    ReportWorksheet.Cell(RowsCount, Value.ColumnIndex).Comment.AddText(Value.Description);
                }
                if (Value.ColumnIndex != 1)
                    ApplyEmptyValueStyle(ReportWorksheet.Cell(RowsCount, Value.ColumnIndex));
            }
        }
        /// <summary>
        /// Применить форматирование для обычной ячейки.
        /// </summary>
        /// <param name="Cell">Ячейка.</param>
        private static void ApplyEmptyValueStyle(IXLCell Cell)
        {
            // Шрифт
            Cell.Style.Font.FontName = "Calibri";
            Cell.Style.Font.FontSize = 13;
            Cell.Style.Font.FontColor = XLColor.Black;
            // Выравнивание
            Cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            // Границы
            SetBorders(Cell, XLBorderStyleValues.Thin, XLColor.Black);
        }
        /// <summary>
        /// Применить форматирование для названия партии.
        /// </summary>
        /// <param name="Cell">Ячейка.</param>
        private static void ApplyEmptyPartyStyle(IXLCell Cell)
        {
            // Шрифт
            Cell.Style.Font.FontName = "Calibri";
            Cell.Style.Font.FontSize = 13;
            Cell.Style.Font.FontColor = XLColor.Black;
            // Выравнивание
            Cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
            // Границы
            SetBorders(Cell, XLBorderStyleValues.Thin, XLColor.Black);
        }
        /// <summary>
        /// Применить форматирование для подписи.
        /// </summary>
        /// <param name="Cell">Ячейка.</param>
        private static void ApplySubscribeStyle(IXLCell Cell)
        {
            // Шрифт
            Cell.Style.Font.FontName = "Calibri";
            Cell.Style.Font.FontSize = 12;
            Cell.Style.Font.FontColor = XLColor.Black;
            // Выравнивание
            Cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
        }
        /// <summary>
        /// Применить форматирование для записи типа прибора.
        /// </summary>
        /// <param name="Cell">Ячейка.</param>
        private static void ApplySubHeaderRowStyle(IXLCell Cell, XLAlignmentHorizontalValues Aligment)
        {
            // Шрифт
            Cell.Style.Font.FontName = "Calibri";
            Cell.Style.Font.FontSize = 14;
            Cell.Style.Font.FontColor = XLColor.Black;
            Cell.Style.Font.Bold = true;
            // Выравнивание
            Cell.Style.Alignment.Horizontal = Aligment;
            // Границы
            SetBorders(Cell, XLBorderStyleValues.Thin, XLColor.Black);
            // Заливка
            Cell.Style.Fill.BackgroundColor = XLColor.Gray;
        }
        /// <summary>
        /// Установить границы ячейки.
        /// </summary>
        /// <param name="Cell">Ячейка.</param>
        /// <param name="BorderStyleValue">Стиль границ.</param>
        /// <param name="BorderColor">Цвет границ.</param>
        private static void SetBorders(IXLCell Cell, XLBorderStyleValues BorderStyleValue, XLColor BorderColor)
        {
            Cell.Style.Border.BottomBorder = BorderStyleValue;
            Cell.Style.Border.BottomBorderColor = BorderColor;

            Cell.Style.Border.TopBorder = BorderStyleValue;
            Cell.Style.Border.TopBorderColor = BorderColor;

            Cell.Style.Border.RightBorder = BorderStyleValue;
            Cell.Style.Border.RightBorderColor = BorderColor;

            Cell.Style.Border.LeftBorder = BorderStyleValue;
            Cell.Style.Border.LeftBorderColor = BorderColor;
        }
        /// <summary>
        /// Установить наружные границы диапазона ячеек.
        /// </summary>
        /// <param name="Range">Диапазон ячеек.</param>
        /// <param name="BorderStyleValue">Стиль границ.</param>
        /// <param name="BorderColor">Цвет границ.</param>
        private static void SetBorders(IXLRange Range, XLBorderStyleValues BorderStyleValue, XLColor BorderColor)
        {
            Range.Style.Border.OutsideBorder = BorderStyleValue;
            Range.Style.Border.OutsideBorderColor = BorderColor;
        }
        /// <summary>
        /// Агрегатор приборов
        /// </summary>
        /// <param name="separator">Сепаратор.</param>
        /// <param name="value1">Значение 1</param>
        /// <param name="value2">Значение 2</param>
        /// <returns></returns>
        public static string JoinDevices(string separator, string value1, string value2)
        {
            if (String.IsNullOrWhiteSpace(value1))
            {
                if (String.IsNullOrWhiteSpace(value2))
                    return String.Empty;
                else
                    return value2;
            }
            else
            {
                if (String.IsNullOrWhiteSpace(value2))
                    return value1;
                else
                    return value1 + separator + value2;
            }
        }
        /// <summary>
        /// Освобождает занимаемую процессом память.
        /// </summary>
        private static void Clear()
        {
            GC.Collect(GC.MaxGeneration);
            GC.WaitForPendingFinalizers();
            Process p = Process.GetCurrentProcess();
            p.MinWorkingSet = p.MinWorkingSet;
        }
    }

    class ReportBalanceComplete
    {
        #region Properties
        /// <summary>
        /// Название файла отчета.
        /// </summary>
        public string FileName
        { get { return "Отчет по остатку комплектующих.xlsx"; } }
        /// <summary>
        /// Документ.
        /// </summary>
        public XLWorkbook ReportDocument;
        /// <summary>
        /// Рабочий лист.
        /// </summary>
        public IXLWorksheet ReportWorksheet;
        /// <summary>
        /// Сессия DV.
        /// </summary>
        UserSession session;
        /// <summary>
        /// Объектный контекст.
        /// </summary>
        ObjectContext Context;
        /// <summary>
        /// Перечень позиций отчета.
        /// </summary>
        List<ReportWareHouseItem> ReportItems;
        /// <summary>
        /// Шаблон отчета.
        /// </summary>
        private Guid TemplateID = new Guid("{7BD6776E-62DE-E911-8D08-00155D11531B}");
        /// <summary>
        /// Временная папка.
        /// </summary>
        public string TempFolder = System.IO.Path.GetTempPath();
        /// <summary>
        /// Индексы столбцов.
        /// </summary>
        private int[] columnsIndex;
        /// <summary>
        /// Текущее количество строк в отчете.
        /// </summary>
        public int RowsCount = 5;
        /// <summary>
        /// Текст заголовка.
        /// </summary>
        const string HeaderText = "Отчёт по остатку комплектующих на складе готовой продукции";
        /// <summary>
        /// Название рабочего листа.
        /// </summary>
        const string WorksheetsName = "Отчёт";
        /// <summary>
        /// Индексы столбцов.
        /// </summary>
        private int[] ColumnsIndex
        {
            get
            {
                if (columnsIndex.IsNull())
                {
                    Array EnumValues = Enum.GetValues(typeof(BalanceCompleteColumns));
                    columnsIndex = new int[EnumValues.Length];

                    for (int i = 0; i < EnumValues.Length; i++)
                        columnsIndex[i] = (int)EnumValues.GetValue(i);
                }
                return columnsIndex;
            }
        }
        /// <summary>
        /// Дата последнего подсчета остатков.
        /// </summary>
        DateTime StartDate;
        #endregion
        /// <summary>
        /// Конструктор отчета по комплектующим.
        /// </summary>
        /// <param name="session">Пользовательская сессия DV.</param>
        /// <param name="StartDate">Дата начала периода.</param>
        public ReportBalanceComplete(UserSession session, DateTime StartDate)
        {
            this.session = session;
            this.StartDate = StartDate;
            Context = session.CreateContext();

            IDocumentService DocService = Context.GetService<IDocumentService>();
            DocsVision.BackOffice.ObjectModel.Document Template = Context.GetObject<DocsVision.BackOffice.ObjectModel.Document>(TemplateID.ToGuid());
            DocumentFile File = Template.Files.First();

            string FilePath = DocService.DownloadFile(File.FileId);

            ReportDocument = new XLWorkbook(FilePath);
            ReportWorksheet = ReportDocument.Worksheets.First(r => r.Name == WorksheetsName);
        }
        /// <summary>
        /// Запись данных в отчет.
        /// </summary>
        public void WriteData(List<CurrentBalanceComplete> CurrentBalanceCompleteCollection, CardData UniversalDictionary)
        {
            IEnumerable<Guid> DeviceTypes = CurrentBalanceCompleteCollection.GroupBy(r => r.DeviceID, (deviceType, items) => deviceType);
            foreach (Guid CurrentDeviceType in DeviceTypes)
            {
                AddDeviceRow(UniversalDictionary.GetItemName(CurrentDeviceType));
                foreach (CurrentBalanceComplete Complete in CurrentBalanceCompleteCollection.Where(r => r.DeviceID == CurrentDeviceType))
                {
                    AddCompleteRow(Complete, UniversalDictionary);
                }
            }

            // Дата и время формирования
            ReportWorksheet.Cell(3, 2).SetValue(DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToShortTimeString());
        }
        /// <summary>
        /// Добавление строки нвого типа прибора
        /// </summary>
        /// <param name="DeviceType">Тип прибора.</param>
        private void AddDeviceRow(string DeviceType)
        {
            RowsCount++;
            // Прибор/партия

            ReportWorksheet.Cell(RowsCount, (int)BalanceCompleteColumns.CompleteName).SetValue(DeviceType);

            ApplySubHeaderRowStyle(ReportWorksheet.Cell(RowsCount, (int)BalanceCompleteColumns.CompleteName), XLAlignmentHorizontalValues.Left);
            ApplySubHeaderRowStyle(ReportWorksheet.Cell(RowsCount, (int)BalanceCompleteColumns.CodeSKB), XLAlignmentHorizontalValues.Left);
            ApplySubHeaderRowStyle(ReportWorksheet.Cell(RowsCount, (int)BalanceCompleteColumns.BalanceOn), XLAlignmentHorizontalValues.Left);
            ApplySubHeaderRowStyle(ReportWorksheet.Cell(RowsCount, (int)BalanceCompleteColumns.Received), XLAlignmentHorizontalValues.Left);
            ApplySubHeaderRowStyle(ReportWorksheet.Cell(RowsCount, (int)BalanceCompleteColumns.Descended), XLAlignmentHorizontalValues.Left);
            ApplySubHeaderRowStyle(ReportWorksheet.Cell(RowsCount, (int)BalanceCompleteColumns.CurrentBalance), XLAlignmentHorizontalValues.Left);
            ApplySubHeaderRowStyle(ReportWorksheet.Cell(RowsCount, (int)BalanceCompleteColumns.Reserved), XLAlignmentHorizontalValues.Left);
            ApplySubHeaderRowStyle(ReportWorksheet.Cell(RowsCount, (int)BalanceCompleteColumns.Total), XLAlignmentHorizontalValues.Left);
        }
        /// <summary>
        /// Добавление строки новой партии прибора
        /// </summary>
        /// <param name="Item">Запись отчета.</param>
        /// <param name="UniversalDictionary">Универсальных справочник.</param>
        private void AddCompleteRow(CurrentBalanceComplete Item, CardData UniversalDictionary)
        {
            RowsCount++;

            ReportWorksheet.Cell(RowsCount, (int)BalanceCompleteColumns.CompleteName).SetValue(UniversalDictionary.GetItemName(Item.CompleteID));
            ApplyEmptyPartyStyle(ReportWorksheet.Cell(RowsCount, (int)BalanceCompleteColumns.CompleteName));

            ReportWorksheet.Cell(RowsCount, (int)BalanceCompleteColumns.CodeSKB).SetValue(UniversalDictionary.GetItemName(Item.CodeSKB));
            ApplyEmptyValueStyle(ReportWorksheet.Cell(RowsCount, (int)BalanceCompleteColumns.CodeSKB), XLColor.NoColor, XLAlignmentHorizontalValues.Left);

            ReportWorksheet.Cell(RowsCount, (int)BalanceCompleteColumns.BalanceOn).SetValue(Item.StartBalance);
            ApplyEmptyValueStyle(ReportWorksheet.Cell(RowsCount, (int)BalanceCompleteColumns.BalanceOn), XLColor.NoColor);

            ReportWorksheet.Cell(RowsCount, (int)BalanceCompleteColumns.Received).SetValue(Item.Received);
            ReportWorksheet.Cell(RowsCount, (int)BalanceCompleteColumns.Received).Comment.AddText(Item.ReceivedDocuments.Aggregate(";\n"));
            ApplyEmptyValueStyle(ReportWorksheet.Cell(RowsCount, (int)BalanceCompleteColumns.Received), XLColor.NoColor);

            ReportWorksheet.Cell(RowsCount, (int)BalanceCompleteColumns.Descended).SetValue(Item.Descended);
            ReportWorksheet.Cell(RowsCount, (int)BalanceCompleteColumns.Descended).Comment.AddText(Item.DescendedDocuments.Aggregate(";\n"));
            ApplyEmptyValueStyle(ReportWorksheet.Cell(RowsCount, (int)BalanceCompleteColumns.Descended), XLColor.NoColor);

            ReportWorksheet.Cell(RowsCount, (int)BalanceCompleteColumns.CurrentBalance).SetValue(Item.EndBalance);
            ApplyEmptyValueStyle(ReportWorksheet.Cell(RowsCount, (int)BalanceCompleteColumns.CurrentBalance), XLColor.YellowMunsell);

            ReportWorksheet.Cell(RowsCount, (int)BalanceCompleteColumns.Reserved).SetValue(Item.Reserved);
            ReportWorksheet.Cell(RowsCount, (int)BalanceCompleteColumns.Reserved).Comment.AddText(Item.ReservedDocuments.Aggregate(";\n"));
            ApplyEmptyValueStyle(ReportWorksheet.Cell(RowsCount, (int)BalanceCompleteColumns.Reserved), XLColor.YellowMunsell);

            ReportWorksheet.Cell(RowsCount, (int)BalanceCompleteColumns.Total).SetValue(Item.EndBalance - Item.Reserved);
            ApplyEmptyValueStyle(ReportWorksheet.Cell(RowsCount, (int)BalanceCompleteColumns.Total), XLColor.RedMunsell);
        }
        /// <summary>
        /// Применить форматирование для обычной ячейки.
        /// </summary>
        /// <param name="Cell">Ячейка.</param>
        /// <param name="Aligment">Выравнивание.</param>
        private static void ApplyEmptyValueStyle(IXLCell Cell, XLColor MyColor, XLAlignmentHorizontalValues Aligment = XLAlignmentHorizontalValues.Center)
        {
            // Шрифт
            Cell.Style.Font.FontName = "Calibri";
            Cell.Style.Font.FontSize = 11;
            Cell.Style.Font.FontColor = XLColor.Black;
            // Выравнивание
            Cell.Style.Alignment.Horizontal = Aligment;
            // Границы
            SetBorders(Cell, XLBorderStyleValues.Thin, XLColor.Black);
            //Заливка
            Cell.Style.Fill.BackgroundColor = MyColor;
        }
        /// <summary>
        /// Применить форматирование для названия партии.
        /// </summary>
        /// <param name="Cell">Ячейка.</param>
        private static void ApplyEmptyPartyStyle(IXLCell Cell)
        {
            // Шрифт
            Cell.Style.Font.FontName = "Calibri";
            Cell.Style.Font.FontSize = 11;
            Cell.Style.Font.FontColor = XLColor.Black;
            // Выравнивание
            Cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
            // Границы
            SetBorders(Cell, XLBorderStyleValues.Thin, XLColor.Black);
        }
        /// <summary>
        /// Применить форматирование для записи типа прибора.
        /// </summary>
        /// <param name="Cell">Ячейка.</param>
        private static void ApplySubHeaderRowStyle(IXLCell Cell, XLAlignmentHorizontalValues Aligment)
        {
            // Шрифт
            Cell.Style.Font.FontName = "Calibri";
            Cell.Style.Font.FontSize = 12;
            Cell.Style.Font.FontColor = XLColor.Black;
            Cell.Style.Font.Bold = true;
            // Выравнивание
            Cell.Style.Alignment.Horizontal = Aligment;
            // Границы
            SetBorders(Cell, XLBorderStyleValues.Thin, XLColor.Black);
            // Заливка
            Cell.Style.Fill.BackgroundColor = XLColor.Gray;
        }
        /// <summary>
        /// Установить границы ячейки.
        /// </summary>
        /// <param name="Cell">Ячейка.</param>
        /// <param name="BorderStyleValue">Стиль границ.</param>
        /// <param name="BorderColor">Цвет границ.</param>
        private static void SetBorders(IXLCell Cell, XLBorderStyleValues BorderStyleValue, XLColor BorderColor)
        {
            Cell.Style.Border.BottomBorder = BorderStyleValue;
            Cell.Style.Border.BottomBorderColor = BorderColor;

            Cell.Style.Border.TopBorder = BorderStyleValue;
            Cell.Style.Border.TopBorderColor = BorderColor;

            Cell.Style.Border.RightBorder = BorderStyleValue;
            Cell.Style.Border.RightBorderColor = BorderColor;

            Cell.Style.Border.LeftBorder = BorderStyleValue;
            Cell.Style.Border.LeftBorderColor = BorderColor;
        }
        /// <summary>
        /// Установить наружные границы диапазона ячеек.
        /// </summary>
        /// <param name="Range">Диапазон ячеек.</param>
        /// <param name="BorderStyleValue">Стиль границ.</param>
        /// <param name="BorderColor">Цвет границ.</param>
        private static void SetBorders(IXLRange Range, XLBorderStyleValues BorderStyleValue, XLColor BorderColor)
        {
            Range.Style.Border.OutsideBorder = BorderStyleValue;
            Range.Style.Border.OutsideBorderColor = BorderColor;
        }
        /// <summary>
        /// Освобождает занимаемую процессом память.
        /// </summary>
        private static void Clear()
        {
            GC.Collect(GC.MaxGeneration);
            GC.WaitForPendingFinalizers();
            Process p = Process.GetCurrentProcess();
            p.MinWorkingSet = p.MinWorkingSet;
        }
    }


    /// <summary>
    /// Запись отчета.
    /// </summary>
    public class ReportCompleteItem
    {
        #region Properties
        /// <summary>
        /// Тип прибора.
        /// </summary>
        private string deviseType;
        /// <summary>
        /// Партия прибора.
        /// </summary>
        private string deviceParty;
        /// <summary>
        /// Значения записи.
        /// </summary>
        public ReportCompleteValue[] Values;
        #endregion
        #region Fields
        /// <summary>
        /// Тип прибора.
        /// </summary>
        public string DeviseType
        { get { return deviseType; } }
        /// <summary>
        /// Партия прибора.
        /// </summary>
        public string DeviceParty
        { get { return deviceParty; } }
        #endregion

        /// <summary>
        /// Конструктор записи отчета.
        /// </summary>
        /// <param name="ColumnsIndex">Индекс столбца.</param>
        /// <param name="DeviseType">Тип прибора.</param>
        /// <param name="DeviceParty">Название партии.</param>
        public ReportCompleteItem(int[] ColumnsIndex, string DeviseType, string DeviceParty)
        {

            this.deviseType = DeviseType;
            this.deviceParty = DeviceParty;
            this.Values = new ReportCompleteValue[ColumnsIndex.Length];
            for (int i = 0; i < ColumnsIndex.Length; i++)
                this.Values[i] = new ReportCompleteValue(ColumnsIndex[i], 0, "");
        }
    }
    /// <summary>
    /// Значение записи отчета.
    /// </summary>
    public class ReportCompleteValue
    {
        #region Properties
        /// <summary>
        /// Индекс столбца.
        /// </summary>
        private int columnIndex;
        /// <summary>
        /// Значение.
        /// </summary>
        public int Value;
        /// <summary>
        /// Примечание.
        /// </summary>
        public string Description;
        #endregion
        #region Fields
        /// <summary>
        /// Индекс столбца.
        /// </summary>
        public int ColumnIndex
        { get { return columnIndex; } }
        #endregion
        /// <summary>
        /// Конструктор значения записи отчета.
        /// </summary>
        /// <param name="ColumnIndex">Индекс столбца.</param>
        /// <param name="Value">Значение.</param>
        /// <param name="Description">Примечание.</param>
        public ReportCompleteValue(int ColumnIndex, int Value, string Description)
        {
            this.columnIndex = ColumnIndex;
            this.Value = Value;
            this.Description = Description;
        }
    }
    /// <summary>
    /// Операция передачи комплектующего.
    /// </summary>
    public class CompleteTransferRow : IComparable<CompleteTransferRow>
    {
        #region Properties
        /// <summary>
        /// Тип передачи комплектующего.
        /// </summary>
        public enum TransferTypes
        {
            /// <summary>
            /// Передача на склад ГП.
            /// </summary>
            [Description("На склад")]
            ToWarehouse = 0,
            /// <summary>
            /// Передача со склада ГП.
            /// </summary>
            [Description("Со склада")]
            FromWarehouse = 1,
            /// <summary>
            /// Внутренняя передача (не затрагивает склад ГП).
            /// </summary>
            [Description("Внутренняя")]
            Internal = 3,
        };
        /// <summary>
        /// Действие с комплектующим.
        /// </summary>
        public enum Action
        {
            /// <summary>
            /// Приход из производства (новые комплектующие).
            /// </summary>
            [Description("Приход из производства (новые комплектующие)")]
            NewReceiptFromProduction = 0,
            /// <summary>
            /// Приход из производства (повторная передача).
            /// </summary>
            [Description("Приход из производства (повторная передача)")]
            RepeatReceiptFromProduction = 1,
            /// <summary>
            /// Возврат на склад с выставок.
            /// </summary>
            [Description("Возврат на склад с выставок")]
            ReturnFromExposition = 2,
            /// <summary>
            /// Возврат на склад с сертификации.
            /// </summary>
            [Description("Возврат на склад с сертификации")]
            ReturnFromCertification = 3,
            /// <summary>
            /// Возврат на склад с испытаний.
            /// </summary>
            [Description("Возврат на склад с испытаний")]
            ReturnFromTesting = 4,
            /// <summary>
            /// Возврат на склад с тест-драйва.
            /// </summary>
            [Description("Возврат на склад с тест-драйва")]
            ReturnFromTestDrive = 5,
            /// <summary>
            /// Возврат со склада в производство.
            /// </summary>
            [Description("Возврат со склада в производство")]
            ReturnFromWarehouseToProduction = 6,
            /// <summary>
            /// Выдача со склада на выставки.
            /// </summary>
            [Description("Выдача со склада на выставки")]
            DeliveryToExposition = 7,
            /// <summary>
            /// Выдача со склада на сертификацию.
            /// </summary>
            [Description("Выдача со склада на сертификацию")]
            DeliveryToCertification = 8,
            /// <summary>
            /// Выдача со склада на испытания.
            /// </summary>
            [Description("Выдача со склада на испытания")]
            DeliveryToTesting = 9,
            /// <summary>
            /// Передача на тест-драйв.
            /// </summary>
            [Description("Передача на тест-драйв")]
            DeliveryToTestDrive = 10,
            /// <summary>
            /// Отгрузка новых комплектующих.
            /// </summary>
            [Description("Отгрузка новых комплектующих")]
            DeliveryNewDevices = 11,
            /// <summary>
            /// Возврат на склад (после продажи/акции).
            /// </summary>
            [Description("Возврат на склад (после продажи/акции)")]
            Return = 12,
            /// <summary>
            /// Нет действия.
            /// </summary>
            [Description("Нет действия")]
            None = 13,
        };
        /// <summary>
        /// Количество комплектующих.
        /// </summary>
        private int completeCount;
        /// <summary>
        /// Запись справочника комплектующих.
        /// </summary>
        private Guid completeID;
        /// <summary>
        /// Запись справочника кодов СКБ.
        /// </summary>
        private Guid completeCodeID;
        /// <summary>
        /// Запись справочника приборов.
        /// </summary>
        private Guid parentDeviceID;
        /// <summary>
        /// Тип передачи.
        /// </summary>
        private TransferTypes transferType;
        /// <summary>
        /// Действие с комплектующим.
        /// </summary>
        private Action transferAction;
        /// <summary>
        /// Дата передачи.
        /// </summary>
        private DateTime transferDate;
        /// <summary>
        /// Название комплектующего.
        /// </summary>
        private string completeType = "";
        /// <summary>
        /// Код СКБ комплектующего.
        /// </summary>
        private string completeCode = "";
        /// <summary>
        /// Родительский прибор для комплектующего.
        /// </summary>
        private string parentDevice = "";
        /// <summary>
        /// Универсальный справочник.
        /// </summary>
        private CardData universalDictionary;
        /// <summary>
        /// Название документа-основания.
        /// </summary>
        private string documentName;
        #endregion

        #region Fields
        /// <summary>
        /// Количество приборов.
        /// </summary>
        public int CompleteCount { get { return completeCount; } }
        /// <summary>
        /// Количество приборов для калькуляции.
        /// </summary>
        public int CompleteCalc
        {
            get
            {
                switch (this.TransferAction)
                {
                    case Action.DeliveryNewDevices:
                        return completeCount * (-1);
                    case Action.DeliveryToCertification:
                        return completeCount * (-1);
                    case Action.DeliveryToExposition:
                        return completeCount * (-1);
                    case Action.DeliveryToTestDrive:
                        return completeCount * (-1);
                    case Action.DeliveryToTesting:
                        return completeCount * (-1);
                    case Action.NewReceiptFromProduction:
                        return completeCount;
                    case Action.RepeatReceiptFromProduction:
                        return completeCount;
                    case Action.Return:
                        return completeCount;
                    case Action.ReturnFromCertification:
                        return completeCount;
                    case Action.ReturnFromExposition:
                        return completeCount;
                    case Action.ReturnFromTestDrive:
                        return completeCount;
                    case Action.ReturnFromTesting:
                        return completeCount;
                    case Action.ReturnFromWarehouseToProduction:
                        return completeCount * (-1);
                    default:
                        return 0;
                }
            }
        }
        /// <summary>
        /// Идентификатор заводского номера прибора.
        /// </summary>
        public Guid CompleteID { get { return completeID; } }
        /// <summary>
        /// Идентификатор кода СКБ.
        /// </summary>
        public Guid CompleteCodeID
        {
            get
            {
                if (completeCodeID == Guid.Empty)
                {
                    if (this.universalDictionary.GetItemPropertyValue(this.completeID, "Код СКБ") != null)
                    {
                        this.completeCodeID = this.universalDictionary.GetItemPropertyValue(this.completeID, "Код СКБ").ToGuid();
                    }
                }
                return completeCodeID;
            }
        }
        /// <summary>
        /// Тип передачи.
        /// </summary>
        public TransferTypes TransferType { get { return transferType; } }
        /// <summary>
        /// Действие.
        /// </summary>
        public Action TransferAction { get { return transferAction; } }
        /// <summary>
        /// Дата передачи.
        /// </summary>
        public DateTime TransferDate { get { return transferDate; } }
        /// <summary>
        /// Код комплектующего.
        /// </summary>
        public string CompleteCode
        {
            get
            {
                if (!CompleteCodeID.IsNull())
                {
                    this.completeCode = this.universalDictionary.GetItemName(CompleteCodeID);
                }
                else
                {
                    this.completeCode = null;
                }
                return this.completeCode;
            }
        }
        /// <summary>
        /// Название комплектующего.
        /// </summary>
        public string СompleteType
        {
            get
            {
                if (string.IsNullOrEmpty(this.completeType))
                {
                    this.completeType = this.universalDictionary.GetItemName(this.completeID);
                }
                return this.completeType;
            }
        }
        /// <summary>
        /// Родительский прибор комплектующего.
        /// </summary>
        public Guid ParentDeviceID
        {
            get
            {
                if (!string.IsNullOrEmpty(this.ParentDevice))
                {
                    RowData DeviceRowData = universalDictionary.GetItemTypeRow(new Guid("{DC3EE278-B3A2-493A-BE7A-74F08B6D57CB}")).ChildSections[new Guid("{DD20BF9B-90F8-4D9A-9553-5B5F17AD724E}")].Rows.FirstOrDefault(r => r.GetString("Name") == this.ParentDevice);
                    if (!DeviceRowData.IsNull())
                        this.parentDeviceID = DeviceRowData.Id;
                }
                return this.parentDeviceID;
            }
        }
        /// <summary>
        /// Родительский прибор комплектующего.
        /// </summary>
        public string ParentDevice
        {
            get
            {
                if (string.IsNullOrEmpty(this.parentDevice))
                    this.parentDevice = universalDictionary.GetItemTypeName(new Guid(universalDictionary.GetItemRow(this.completeID).GetGuid("ParentRowID").ToString()));
                return this.parentDevice;
            }
        }
        /// <summary>
        /// Название документа-основания
        /// </summary>
        public string DocumentName
        {
            get
            {
                return this.documentName;
            }
        }
        #endregion
        /// <summary>
        /// Конструктор операции передачи комплектующего.
        /// </summary>
        /// <param name="UniversalDictionary">Карточка универсального справочника.</param>
        /// <param name="СompleteID">Идентификатор комплектующего.</param>
        /// <param name="TransferType">Тип передачи.</param>
        /// <param name="TransferAction">Действие с комплектующим.</param>
        /// <param name="TransferDate">Дата передачи.</param>
        /// <param name="Count">Количество комплектующих.</param>
        /// <param name="DocumentName">Название документа-основания.</param>
        public CompleteTransferRow(CardData UniversalDictionary, Guid СompleteID, TransferTypes TransferType, Action TransferAction, DateTime TransferDate, int Count = 1, string DocumentName = "")
        {
            if (TransferAction == Action.ReturnFromTestDrive)
                this.transferAction = TransferAction;
            this.universalDictionary = UniversalDictionary;
            this.completeID = СompleteID;
            this.transferType = TransferType;
            this.transferAction = TransferAction;
            this.transferDate = TransferDate;
            this.completeCount = Count;
            this.documentName = DocumentName;
        }
        /// <summary>
        /// Сравнение операций передачи по дате.
        /// </summary>
        /// <param name="other">Операция для сравнения.</param>
        /// <returns></returns>
        int IComparable<CompleteTransferRow>.CompareTo(CompleteTransferRow other)
        {
            DateTime DateOther = other.TransferDate.Date;
            DateTime DateThis = this.TransferDate.Date;

            if (DateOther > DateThis)
                return -1;
            else if (DateOther == DateThis)
            {
                List<Action> DeliveryList = new List<Action> { Action.DeliveryNewDevices, Action.DeliveryToCertification, Action.DeliveryToExposition, Action.DeliveryToTestDrive, Action.DeliveryToTesting };
                List<Action> ReceiptList = new List<Action> { Action.NewReceiptFromProduction, Action.RepeatReceiptFromProduction };
                List<Action> ReturnToWHList = new List<Action> { Action.ReturnFromCertification, Action.ReturnFromExposition, Action.ReturnFromTestDrive, Action.ReturnFromTesting };
                List<Action> ReturnToProductionList = new List<Action> { Action.ReturnFromWarehouseToProduction };

                if (DeliveryList.Any(r => r == other.TransferAction) && ReceiptList.Any(r => r == this.TransferAction))
                    return -1;
                if (DeliveryList.Any(r => r == this.TransferAction) && ReceiptList.Any(r => r == other.TransferAction))
                    return 1;
                if (ReturnToProductionList.Any(r => r == other.TransferAction) && ReturnToWHList.Any(r => r == this.TransferAction))
                    return -1;
                if (ReturnToProductionList.Any(r => r == this.TransferAction) && ReturnToWHList.Any(r => r == other.TransferAction))
                    return 1;
                if (ReceiptList.Any(r => r == other.TransferAction) && ReturnToProductionList.Any(r => r == this.TransferAction))
                    return -1;
                if (ReceiptList.Any(r => r == this.TransferAction) && ReturnToProductionList.Any(r => r == other.TransferAction))
                    return 1;
                if (DeliveryList.Any(r => r == other.TransferAction) && ReturnToWHList.Any(r => r == this.TransferAction))
                    return -1;
                if (DeliveryList.Any(r => r == this.TransferAction) && ReturnToWHList.Any(r => r == other.TransferAction))
                    return 1;
                return 0;
            }
            else
                return 1;
        }
    }
    /// <summary>
    /// Операция передачи комплектующих по типу.
    /// </summary>
    public class TransferCountByCompleteType
    {
        #region Properties
        /// <summary>
        /// Количество комплектующих.
        /// </summary>
        private int completeCount;
        /// <summary>
        /// ID комплектующего.
        /// </summary>
        private Guid completeID;
        /// <summary>
        /// ID кода СКБ комплектующего.
        /// </summary>
        private Guid completeCodeID;
        /// <summary>
        /// ID типа прибора.
        /// </summary>
        private Guid parentDeviceID;
        /// <summary>
        /// Тип комплектующего.
        /// </summary>
        private string completeType;
        /// <summary>
        /// Код СКБ комплектующего.
        /// </summary>
        private string completeCode;
        /// <summary>
        /// Код СКБ комплектующего.
        /// </summary>
        private string parentDevice;
        /// <summary>
        /// Название документа-основания.
        /// </summary>
        private string documentName;
        #endregion
        #region Fields
        /// <summary>
        /// Количество комплектующих.
        /// </summary>
        public int CompleteCount
        {
            get { return completeCount; }
            set { completeCount = value; }
        }
        /// <summary>
        /// Идентификатор заводского номера прибора.
        /// </summary>
        public Guid CompleteID { get { return completeID; } }
        /// <summary>
        /// Код комплектующего.
        /// </summary>
        public Guid CompleteCodeID
        {
            get { return completeCodeID; }
        }
        /// <summary>
        /// Код типа прибора.
        /// </summary>
        public Guid ParentDeviceID
        {
            get { return parentDeviceID; }
        }
        /// <summary>
        /// Название типа комплектующего.
        /// </summary>
        public string CompleteType { get { return completeType; } }
        /// <summary>
        /// Название Кода СКБ комплектующего.
        /// </summary>
        public string CompleteCode { get { return completeCode; } }
        /// <summary>
        /// Название типа прибора.
        /// </summary>
        public string ParentDevice { get { return parentDevice; } }
        /// <summary>
        /// Название документа-основания.
        /// </summary>
        public string DocumentName { get { return documentName; } }
        #endregion

        /// <summary>
        /// Конструктор операции передачи комплектующих по типу.
        /// </summary>
        /// <param name="CompleteID">ID комплектующего.</param>
        /// <param name="CompleteCodeID">ID кода СКБ комплектующего.</param>
        /// <param name="ParentDeviceID">ID типа прибора.</param>
        /// <param name="CompleteType">Тип комплектующего.</param>
        /// <param name="CompleteCode">Код СКБ комплектующего.</param>
        /// <param name="ParentDevice">Тип прибора.</param>
        /// <param name="CompleteCount">Количество комплектующих.</param>
        /// <param name="DocumentName">Название документа-основания.</param>
        public TransferCountByCompleteType(Guid CompleteID, Guid CompleteCodeID, Guid ParentDeviceID, string CompleteType, string CompleteCode, string ParentDevice, int CompleteCount, string DocumentName)
        {
            this.completeID = CompleteID;
            this.completeCodeID = CompleteCodeID;
            this.parentDeviceID = ParentDeviceID;
            this.completeType = CompleteType;
            this.completeCode = CompleteCode;
            this.parentDevice = ParentDevice;
            this.completeCount = CompleteCount;
            this.documentName = DocumentName;
        }
    }

    /// <summary>
    /// Работа с отчетом по комплектующим
    /// </summary>
    public static class ReportCompleteHelper
    {
        // Дополнительные изделия:
        /// <summary>
        /// Перечень доп. изделий, исключенных из отчета
        /// </summary>
        //public static string[] AdditionalWaresList = new string[] { "ДП12", "ДП21", "ДП22", "ТК-021", "ДП32.1", "ДП32.2", "ТК-026", "ПКП-08-00", "ПКП-08-01", "ПКП-08-02", "ПКП-25-00", "ПКП-25-01", "ПКП-25-02" };
        // Предметы договора/счета:
        /// <summary>
        /// Поставка готовой продукции
        /// </summary>
        const string DeliveryID = SKB.PaymentAndShipment.Cards.AccountCard.Item_Subject_Delivery;
        /// <summary>
        /// Отправка на выставку
        /// </summary>
        const string ExpositionID = SKB.PaymentAndShipment.Cards.AccountCard.Item_Subject_ShippingExhibition;
        /// <summary>
        /// Тест-драйв
        /// </summary>
        const string TestDriveID = SKB.PaymentAndShipment.Cards.AccountCard.Item_Subject_TestDrive;
        /// <summary>
        /// Отправка на сертификацию
        /// </summary>
        const string CertificationID = SKB.PaymentAndShipment.Cards.AccountCard.Item_Subject_ShippingСertification;
        /// <summary>
        /// Тендер
        /// </summary>
        const string TenderID = SKB.PaymentAndShipment.Cards.AccountCard.Item_Subject_Tender;
        /// <summary>
        /// Акция
        /// </summary>
        const string ActionID = SKB.PaymentAndShipment.Cards.AccountCard.Item_Subject_Action;
        /// <summary>
        /// Сервисное обслуживание
        /// </summary>
        const string ServiceID = SKB.PaymentAndShipment.Cards.AccountCard.Item_Subject_Service;
        /// <summary>
        /// Сервисное обслуживание
        /// </summary>
        const string ServiceDeliveryID = SKB.PaymentAndShipment.Cards.AccountCard.Item_Subject_ServiceDelivery;
        /// <summary>
        /// Семинар
        /// </summary>
        const string SeminarID = SKB.PaymentAndShipment.Cards.AccountCard.Item_Subject_Seminar;
        /// <summary>
        /// Освобождает занимаемую процессом память.
        /// </summary>
        private static void Clear()
        {
            GC.Collect(GC.MaxGeneration);
            GC.WaitForPendingFinalizers();
            Process p = Process.GetCurrentProcess();
            p.MinWorkingSet = p.MinWorkingSet;
        }
        /// <summary>
        /// Поиск Актов передачи комплектующих на склад.
        /// </summary>
        /// <param name="session">Пользовательская сессия DV.</param>
        /// <param name="StartDate">Начальная дата.</param>
        /// <param name="EndDate">Конечная дата.</param>
        /// <param name="UniversalDictionary">Карточка универсального справочника.</param>
        /// <returns></returns>
        public static IEnumerable<CompleteTransferRow> FindAct(UserSession session, DateTime StartDate, DateTime EndDate, CardData UniversalDictionary)
        {
            /* Поиск Актов передачи приборов и комплектующих */
            SearchQuery searchQuery = session.CreateSearchQuery();
            searchQuery.CombineResults = ConditionGroupOperation.And;

            CardTypeQuery typeQuery = searchQuery.AttributiveSearch.CardTypeQueries.AddNew(CardOrd.ID);
            SectionQuery sectionQuery = typeQuery.SectionQueries.AddNew(CardOrd.MainInfo.ID);
            sectionQuery.Operation = SectionQueryOperation.And;
            sectionQuery.ConditionGroup.Operation = ConditionGroupOperation.And;
            sectionQuery.ConditionGroup.Conditions.AddNew(CardOrd.MainInfo.Type, FieldType.RefId, ConditionOperation.Equals, MyHelper.RefType_ActOfTransfer);

            // Режим передачи - НЕ "Новые комплектующие".
            /*sectionQuery = typeQuery.SectionQueries.AddNew(CardOrd.Properties.ID);
            sectionQuery.Operation = SectionQueryOperation.And;
            sectionQuery.ConditionGroup.Operation = ConditionGroupOperation.And;
            sectionQuery.ConditionGroup.Conditions.AddNew(CardOrd.Properties.Name, FieldType.Unistring, ConditionOperation.Equals, "Режим передачи");
            sectionQuery.ConditionGroup.Conditions.AddNew(CardOrd.Properties.Value, FieldType.Int, ConditionOperation.NotEquals, 2);*/

            // Состояние акта - "Принят".
            sectionQuery = typeQuery.SectionQueries.AddNew(CardOrd.Properties.ID);
            sectionQuery.Operation = SectionQueryOperation.And;
            sectionQuery.ConditionGroup.Operation = ConditionGroupOperation.And;
            sectionQuery.ConditionGroup.Conditions.AddNew(CardOrd.Properties.Name, FieldType.Unistring, ConditionOperation.Equals, "Состояние акта");
            sectionQuery.ConditionGroup.Conditions.AddNew(CardOrd.Properties.Value, FieldType.Int, ConditionOperation.Equals, 3);

            // Дата принятия акта меньше стартовой даты отчета
            sectionQuery = typeQuery.SectionQueries.AddNew(CardOrd.Properties.ID);
            sectionQuery.Operation = SectionQueryOperation.And;
            sectionQuery.ConditionGroup.Operation = ConditionGroupOperation.And;
            sectionQuery.ConditionGroup.Conditions.AddNew(CardOrd.Properties.Name, FieldType.Unistring, ConditionOperation.Equals, "Дата принятия");
            sectionQuery.ConditionGroup.Conditions.AddNew(CardOrd.Properties.Value, FieldType.Date, ConditionOperation.IsNotNull);
            sectionQuery.ConditionGroup.Conditions.AddNew(CardOrd.Properties.Value, FieldType.Date, ConditionOperation.LessThan, EndDate.AddDays(1));
            sectionQuery.ConditionGroup.Conditions.AddNew(CardOrd.Properties.Value, FieldType.Date, ConditionOperation.GreaterThan, StartDate.AddDays(-1)); //new DateTime(2013, 1, 1));

            // Получение текста запроса
            searchQuery.Limit = 0;
            string query = searchQuery.GetXml();
            CardDataCollection ActCollection = session.CardManager.FindCards(query);
            //MyMessageBox.Show("Нашли акты..." + ActCollection.Count());
            int i = 0;
            IEnumerable<CompleteTransferRow> DeviceNumberIDs = ActCollection.SelectMany(r => r.ActToTransferRowCollection(UniversalDictionary, ref i));
            //MyMessageBox.Show("Обработали акты");
            Clear();

            return DeviceNumberIDs;
        }
        /// <summary>
        /// Преобразование Акта передачи в коллекцию операций передачи комплектующих.
        /// </summary>
        /// <param name="Card">Карточка Акта.</param>
        /// <param name="UniversalDictionary">Карточка универсального справочника.</param>
        /// <param name="i">Количество операций.</param>
        /// <returns></returns>
        private static IEnumerable<CompleteTransferRow> ActToTransferRowCollection(this CardData Card, CardData UniversalDictionary, ref int i)
        {
            try
            {
                CompleteTransferRow.TransferTypes TransferType;
                CompleteTransferRow.Action TransferAction;

                switch (Card.Sections[CardOrd.Properties.ID].FindRow("@Name = 'Вид передачи'").GetString(CardOrd.Properties.Value))
                {
                    case "Калибровка -> Сбыт":
                        TransferType = CompleteTransferRow.TransferTypes.ToWarehouse;
                        if ((int)Card.Sections[CardOrd.Properties.ID].FindRow("@Name = 'Режим передачи'").GetInt32(CardOrd.Properties.Value) == 1)
                            TransferAction = CompleteTransferRow.Action.NewReceiptFromProduction;
                        else
                            TransferAction = CompleteTransferRow.Action.RepeatReceiptFromProduction;
                        break;
                    case "Сбыт -> Калибровка":
                        TransferType = CompleteTransferRow.TransferTypes.FromWarehouse;
                        switch ((int)Card.Sections[CardOrd.Properties.ID].FindRow("@Name = 'Цель передачи'").GetInt32(CardOrd.Properties.Value))
                        {
                            case 2:
                                TransferAction = CompleteTransferRow.Action.DeliveryToTesting;
                                break;
                            case 3:
                                TransferAction = CompleteTransferRow.Action.ReturnFromWarehouseToProduction;
                                break;
                            default:
                                TransferAction = CompleteTransferRow.Action.ReturnFromWarehouseToProduction;
                                break;
                        }
                        break;
                    default:
                        TransferType = CompleteTransferRow.TransferTypes.Internal;
                        TransferAction = CompleteTransferRow.Action.None;
                        break;
                }

                IEnumerable<RowData> VerifyRows = Card.Sections[CardOrd.Properties.ID].FindRow("@Name = 'Наименование компл.'").ChildSections[CardOrd.SelectedValues.ID].Rows;
                IEnumerable<CompleteTransferRow> Result = VerifyRows.Select(r =>
                        new CompleteTransferRow(
                            UniversalDictionary,
                            (Guid)r.GetGuid(CardOrd.SelectedValues.SelectedValue),
                            TransferType, TransferAction,
                            (DateTime)Card.Sections[CardOrd.Properties.ID].FindRow("@Name = 'Дата принятия'").GetDateTime(CardOrd.Properties.Value),
                            (int)Card.Sections[CardOrd.Properties.ID].FindRow("@Name = 'Кол-во компл.'").ChildSections[CardOrd.SelectedValues.ID].FindRow("@Order = '" + r.GetString(CardOrd.SelectedValues.Order) + "'").GetInt32(CardOrd.SelectedValues.SelectedValue),
                            Card.Description));
                return Result;
            }
            catch
            {
                //MyMessageBox.Show("Произошла ошибка при преобразовании Акта в операции передачи комплектующих. Акт: " + Card.Description);
                return null;
            }
        }
        /// <summary>
        /// Поиск Заданий на комплектацию.
        /// </summary>
        /// <param name="session">Пользовательская сессия DV.</param>
        /// <param name="StartDate">Начальная дата.</param>
        /// <param name="EndDate">Конечная дата.</param>
        /// <param name="UniversalDictionary">Карточка универсального справочника.</param>
        /// <returns></returns>
        public static IEnumerable<CompleteTransferRow> FindCompleteTasks(UserSession session, DateTime StartDate, DateTime EndDate, CardData UniversalDictionary)
        {
            List<CompleteTransferRow> Result = new List<CompleteTransferRow>();

            /* Поиск Заданий на комплектацию, у которых дата отгрузки попадает в заданный интервал */
            SearchQuery searchQuery = session.CreateSearchQuery();
            searchQuery.CombineResults = ConditionGroupOperation.And;

            CardTypeQuery typeQuery = searchQuery.AttributiveSearch.CardTypeQueries.AddNew(RefCompleteCard.ID);

            // Фактическая дата отгрузки попадает в заданный интервал
            SectionQuery sectionQuery = typeQuery.SectionQueries.AddNew(RefCompleteCard.Devices.ID);
            sectionQuery.Operation = SectionQueryOperation.And;
            sectionQuery.ConditionGroup.Operation = ConditionGroupOperation.And;
            sectionQuery.ConditionGroup.Conditions.AddNew(RefCompleteCard.Devices.FactShipDate, FieldType.DateTime, ConditionOperation.LessEqual, EndDate.AddDays(1));
            sectionQuery.ConditionGroup.Conditions.AddNew(RefCompleteCard.Devices.FactShipDate, FieldType.DateTime, ConditionOperation.GreaterThan, StartDate.AddDays(-1));

            // Получение текста запроса
            searchQuery.Limit = 0;
            string query = searchQuery.GetXml();

            CardDataCollection CompleteTaskCollection = session.CardManager.FindCards(query);
            //MyMessageBox.Show("Нашли Задания на комплектацию, у которых дата отгрузки попадает в заданный интервал..." + CompleteTaskCollection.Count());

            // Получение конвертированных строк Заявок на комплектацию (у которых дата отгрузки попадает в заданный интервал)



            foreach (CardData CompleteTask in CompleteTaskCollection)
            {
                IEnumerable<RowData> FindRows = CompleteTask.Sections[RefCompleteCard.Devices.ID].Rows;
                //MyMessageBox.Show("Проверили задание, всего строк: " + FindRows.Count());
                IEnumerable<RowData> CheckRows = FindRows.Where(r => !r.GetDateTime(RefCompleteCard.Devices.FactShipDate).IsNull() &&
            r.GetDateTime(RefCompleteCard.Devices.FactShipDate) > StartDate.AddDays(-1) &&
            r.GetDateTime(RefCompleteCard.Devices.FactShipDate) < EndDate.AddDays(1));
                //MyMessageBox.Show("Подходящих строк: " + CheckRows.Count());
                IEnumerable<CompleteTransferRow> ConvertRows = CheckRows.SelectMany(r => ConvertCompleteToTransferRowsCollection(r, session, UniversalDictionary, StartDate, EndDate, CompleteTask.Description));
                //MyMessageBox.Show("Сконвертированных строк: " + ConvertRows.Count());
                List<CompleteTransferRow> R = ConvertRows.ToList();
                Result.AddRange(R);
                //MyMessageBox.Show("Итого в результате: " + Result.Count());
            }
            //MyMessageBox.Show("Получили коллекцию строк заданий на комплектацию, у которых дата ОТГРУЗКИ попадает в заданный интервал. Всего строк: " + Result.Count());
            int ResultCount = Result.Count();

            /* Поиск Заданий на комплектацию, у которых дата возврата попадает в указанный интервал */
            SearchQuery searchQuery2 = session.CreateSearchQuery();
            searchQuery2.CombineResults = ConditionGroupOperation.And;

            CardTypeQuery typeQuery2 = searchQuery2.AttributiveSearch.CardTypeQueries.AddNew(RefCompleteCard.ID);

            // Фактическая дата возврата попадает в заданный интервал
            SectionQuery sectionQuery2 = typeQuery2.SectionQueries.AddNew(RefCompleteCard.Devices.ID);
            sectionQuery2.Operation = SectionQueryOperation.And;
            sectionQuery2.ConditionGroup.Operation = ConditionGroupOperation.And;
            sectionQuery2.ConditionGroup.Conditions.AddNew(RefCompleteCard.Devices.FactReturnDate, FieldType.DateTime, ConditionOperation.LessEqual, EndDate.AddDays(1));
            sectionQuery2.ConditionGroup.Conditions.AddNew(RefCompleteCard.Devices.FactReturnDate, FieldType.DateTime, ConditionOperation.GreaterThan, StartDate.AddDays(-1));

            // Получение текста запроса
            searchQuery2.Limit = 0;
            string query2 = searchQuery2.GetXml();

            CardDataCollection CompleteTaskCollection2 = session.CardManager.FindCards(query2);
            //MyMessageBox.Show("Нашли Задания на комплектацию, у которых дата возврата попадает в указанный интервал..." + CompleteTaskCollection2.Count());

            foreach (CardData CompleteTask in CompleteTaskCollection)
            {
                Result.Union(CompleteTask.Sections[RefCompleteCard.Devices.ID].Rows.Where(r => !r.GetDateTime(RefCompleteCard.Devices.FactShipDate).IsNull() &&
            (DateTime)r.GetDateTime(RefCompleteCard.Devices.FactReturnDate) > StartDate.AddDays(-1) &&
            (DateTime)r.GetDateTime(RefCompleteCard.Devices.FactReturnDate) < EndDate.AddDays(1)).SelectMany(r =>
            ConvertCompleteToTransferRowsCollection(r, session, UniversalDictionary, StartDate, EndDate, CompleteTask.Description)));
            }
            //MyMessageBox.Show("Получили коллекцию строк заданий на комплектацию, у которых дата ВОЗВРАТА попадает в заданный интервал. Всего строк: " + (Result.Count() - ResultCount));

            Clear();
            //MyMessageBox.Show("Всего строк по отгрузке/возврату: " + Result.Count());
            return Result;
        }
        /// <summary>
        /// Преобразование строки таблицы "Информация по приборам" в коллекцию операций передачи приборов.
        /// </summary>
        /// <param name="CompleteRow">Строка таблицы "Информация по приборам."</param>
        /// <param name="session">Сессия."</param>
        /// <param name="UniversalDictionary">Карточка универсального справочника.</param>
        /// <param name="StartDate">Дата начала.</param>
        /// <param name="EndDate">Дата окончания.</param>
        /// <param name="DocumentName">Название документа основания.</param>
        /// <returns></returns>
        private static IEnumerable<CompleteTransferRow> ConvertCompleteToTransferRowsCollection(RowData CompleteRow, UserSession session, CardData UniversalDictionary, DateTime StartDate, DateTime EndDate, string DocumentName)//, CardDataCollection ShipmentTaskCollection, CardDataCollection AccountCardCollection)
        {
            int s = 0;
            string MyCompleteName = "";
            try
            {
                //MyMessageBox.Show("Трансформируем задание на комплектацию: " + CompleteRow.GetObject(RefCompleteCard.Devices.FactShipDate).ToString());

                List<CompleteTransferRow> TransferRowsCollection = new List<CompleteTransferRow>();
                /*if ((bool)CompleteRow.GetBoolean(RefCompleteCard.Devices.AC) || AdditionalWaresList.Any(r => r == UniversalDictionary.GetItemName((Guid)CompleteRow.GetGuid(RefCompleteCard.Devices.DeviceId))))
                    return TransferRowsCollection.AsEnumerable();*/

                string ContractSubjectId = CompleteRow.GetString(RefCompleteCard.Devices.ContractSubjectId);
                //MyMessageBox.Show("1");
                s = 1;
                //object DeviceNumberId = CompleteRow.GetGuid(RefCompleteCard.Devices.DeviceNumberId);
                CompleteRow.GetGuid(RefCompleteCard.Devices.ShipmentTaskId);
                //MyMessageBox.Show("2");
                s = 2;
                object FactShipDate = CompleteRow.GetObject(RefCompleteCard.Devices.FactShipDate);
                //MyMessageBox.Show("3");
                s = 3;
                object FactReturnDate = CompleteRow.GetObject(RefCompleteCard.Devices.FactReturnDate);
                //MyMessageBox.Show("4");
                s = 4;
                object PaymentDate = CompleteRow.GetObject(RefCompleteCard.Devices.PaymentDate);
                //MyMessageBox.Show("5");
                s = 5;
                object DeviceID = CompleteRow.GetObject(RefCompleteCard.Devices.DeviceId);
                //MyMessageBox.Show("6");
                s = 6;
                Guid ParentItemTypeID = UniversalDictionary.GetItemRow(DeviceID.ToGuid()).GetString("ParentRowID").ToGuid();
                //MyMessageBox.Show("7");
                s = 7;
                string DeviceName = UniversalDictionary.GetItemName(DeviceID.ToGuid());
                //MyMessageBox.Show("8");
                s = 8;
                RowData DeviceDictionary = UniversalDictionary.GetItemTypeRow(ParentItemTypeID).ChildRows.FirstOrDefault(r => r.GetString("Name") == DeviceName);
                if (DeviceDictionary.IsNull())
                    return TransferRowsCollection;
                //MyMessageBox.Show("9");
                s = 9;

                // Если дата фактической отгрузки задана и предмет договора счета - НЕ Сервисное обслуживание/Семинар
                if (!FactShipDate.IsNull() && ContractSubjectId != ServiceID && ContractSubjectId != SeminarID)
                {
                    //MyMessageBox.Show("10");
                    s = 10;
                    // Производится поиск связанного Задания на отгрузку для определения комплектации
                    /*CardData ShipmentTask = ShipmentTaskCollection.FirstOrDefault(r => r.Id.Equals(CompleteRow.GetString(RefCompleteCard.Devices.ShipmentTaskId).ToGuid()));
                    if (ShipmentTask.IsNull())
                        return TransferRowsCollection;*/
                    //MyMessageBox.Show("ShipmentTaskId = " + CompleteRow.GetString(RefCompleteCard.Devices.ShipmentTaskId));
                    /* Поиск Задания на отгрузку */
                    SearchQuery searchQuery2 = session.CreateSearchQuery();
                    searchQuery2.CombineResults = ConditionGroupOperation.And;

                    CardTypeQuery typeQuery2 = searchQuery2.AttributiveSearch.CardTypeQueries.AddNew(RefShipmentCard.ID);

                    // Фактическая дата отгрузки меньше конечной даты
                    SectionQuery sectionQuery2 = typeQuery2.SectionQueries.AddNew(RefShipmentCard.MainInfo.ID);
                    sectionQuery2.Operation = SectionQueryOperation.And;
                    sectionQuery2.ConditionGroup.Operation = ConditionGroupOperation.And;
                    sectionQuery2.ConditionGroup.Conditions.AddNew("InstanceID", FieldType.UniqueId, ConditionOperation.Equals, CompleteRow.GetString(RefCompleteCard.Devices.ShipmentTaskId).ToGuid());
                    // Получение текста запроса
                    searchQuery2.Limit = 0;
                    string query2 = searchQuery2.GetXml();

                    CardDataCollection ShipmentTaskCollection = session.CardManager.FindCards(query2);
                    //MyMessageBox.Show("ShipmentTaskCollection = " + ShipmentTaskCollection.Count());

                    if (ShipmentTaskCollection.IsNull())
                        return TransferRowsCollection;
                    //MyMessageBox.Show("ShipmentTask = !");
                    CardData ShipmentTask = ShipmentTaskCollection.FirstOrDefault();
                    if (ShipmentTask.IsNull())
                        return TransferRowsCollection;

                    //MyMessageBox.Show("11");
                    s = 11;

                    //MyMessageBox.Show("12");
                    s = 12;
                    RowData ShipmentTaskRow = ShipmentTask.Sections[RefShipmentCard.Devices.ID].Rows.FirstOrDefault(r => r.GetString(RefShipmentCard.Devices.Id).ToGuid().Equals(CompleteRow.GetString(RefCompleteCard.Devices.ShipmentTaskRowId).ToGuid()));
                    //MyMessageBox.Show("13");
                    s = 13;

                    if (ShipmentTaskRow.IsNull())
                        return TransferRowsCollection;

                    //MyMessageBox.Show("14");
                    s = 14;

                    /* Поиск Договоров/счетов */
                    SearchQuery searchQuery = session.CreateSearchQuery();
                    searchQuery.CombineResults = ConditionGroupOperation.And;

                    CardTypeQuery typeQuery = searchQuery.AttributiveSearch.CardTypeQueries.AddNew(RefAccountCard.ID);

                    //MyMessageBox.Show("AccountNumber = " + ShipmentTask.Sections[RefShipmentCard.MainInfo.ID].FirstRow.GetString(RefShipmentCard.MainInfo.AccountNumber));
                    //MyMessageBox.Show("AccountDate = " + (DateTime)ShipmentTask.Sections[RefShipmentCard.MainInfo.ID].FirstRow.GetDateTime(RefShipmentCard.MainInfo.AccountDate));

                    // Фактическая дата отгрузки меньше конечной даты
                    SectionQuery sectionQuery = typeQuery.SectionQueries.AddNew(RefAccountCard.MainInfo.ID);
                    sectionQuery.Operation = SectionQueryOperation.And;
                    sectionQuery.ConditionGroup.Operation = ConditionGroupOperation.And;
                    sectionQuery.ConditionGroup.Conditions.AddNew(RefAccountCard.MainInfo.Number, FieldType.String, ConditionOperation.Equals, ShipmentTask.Sections[RefShipmentCard.MainInfo.ID].FirstRow.GetString(RefShipmentCard.MainInfo.AccountNumber).ToString());
                    sectionQuery.ConditionGroup.Conditions.AddNew(RefAccountCard.MainInfo.ContractDate, FieldType.Date, ConditionOperation.Equals, (DateTime)ShipmentTask.Sections[RefShipmentCard.MainInfo.ID].FirstRow.GetDateTime(RefShipmentCard.MainInfo.AccountDate));
                    // Получение текста запроса
                    searchQuery.Limit = 0;
                    string query = searchQuery.GetXml();

                    CardDataCollection FindAccountCardCollection = session.CardManager.FindCards(query);
                    //MyMessageBox.Show("FindAccountCardCollection = " + FindAccountCardCollection.Count());
                    //MyMessageBox.Show("Нашли связанные договоры..." + AccountCardCollection.Count());

                    //CardData AccountCard = AccountCardCollection.FirstOrDefault(r =>
                    //    (r.Sections[RefAccountCard.MainInfo.ID].FirstRow.GetString(RefAccountCard.MainInfo.Number) == ShipmentTask.Sections[RefShipmentCard.MainInfo.ID].FirstRow.GetString(RefShipmentCard.MainInfo.AccountNumber)) &&
                    //    (r.Sections[RefAccountCard.MainInfo.ID].FirstRow.GetDateTime(RefAccountCard.MainInfo.ContractDate) == ShipmentTask.Sections[RefShipmentCard.MainInfo.ID].FirstRow.GetDateTime(RefShipmentCard.MainInfo.AccountDate)));

                    CardData AccountCard = FindAccountCardCollection.FirstOrDefault();

                    //MyMessageBox.Show("15");
                    s = 15;
                    if (AccountCard.IsNull())
                        return TransferRowsCollection;

                    //MyMessageBox.Show("16");
                    s = 16;
                    RowData AccountCardRow = AccountCard.Sections[RefAccountCard.Devices.ID].Rows.FirstOrDefault(r => r.GetString(RefAccountCard.Devices.Id).ToGuid().Equals(ShipmentTaskRow.GetString(RefShipmentCard.Devices.AccountCardRowId).ToGuid()));
                    //MyMessageBox.Show("17");
                    s = 17;
                    if (AccountCardRow.IsNull())
                        return TransferRowsCollection;
                    //MyMessageBox.Show("18");
                    s = 18;
                    string PackedListData = AccountCardRow.GetString(RefAccountCard.Devices.PackedListData);
                    //MyMessageBox.Show("19");
                    s = 19;
                    string[] Completes = PackedListData.Split('\n');
                    //MyMessageBox.Show("20");
                    s = 20;
                    foreach (string Complete in Completes)
                    {
                        //MyMessageBox.Show("20.1");
                        s = 21;
                        DeviceCompleteRow DevCompleteRow = (DeviceCompleteRow)Complete;
                        string CompleteName = DevCompleteRow.Name;
                        //MyMessageBox.Show(CompleteName);
                        //string CompleteCode = Complete.Split('\t')[1] == "-" ? "" : Complete.Split('\t')[1].Substring(4);
                        //MyMessageBox.Show("20.2");
                        s = 22;
                        int CompleteType = DevCompleteRow.TableType;
                        //MyMessageBox.Show("20.3");
                        s = 23;
                        int CompleteCount = DevCompleteRow.Count.IsNull() ? 0 : Convert.ToInt32(DevCompleteRow.Count);
                        //int CompleteCount = Complete.Split('\t')[3] == null || Complete.Split('\t')[3] == String.Empty || Complete.Split('\t')[3] == "" ? 0 : (int)Complete.Split('\t')[3].ToInt32();
                        //MyMessageBox.Show("20.4");
                        s = 24;
                        MyCompleteName = CompleteName;
                        if (CompleteCount > 0)
                        {
                            s = 26;
                            RowData CompleteRowData = DeviceDictionary.ChildSections[new Guid("{DD20BF9B-90F8-4D9A-9553-5B5F17AD724E}")].Rows.FirstOrDefault(r => r.GetString("Name") == CompleteName);

                            if (!CompleteRowData.IsNull())
                            {
                                Guid CompleteID = CompleteRowData.Id;
                                //MyMessageBox.Show("20.5");
                                s = 27;
                                switch (ContractSubjectId)
                                {
                                    case DeliveryID:
                                        //MyMessageBox.Show("20.6");
                                        s = 28;
                                        TransferRowsCollection.Add(new CompleteTransferRow(UniversalDictionary, CompleteID, CompleteTransferRow.TransferTypes.FromWarehouse, CompleteTransferRow.Action.DeliveryNewDevices, (DateTime)FactShipDate, CompleteCount, DocumentName));
                                        //MyMessageBox.Show("20.7");
                                        s = 29;
                                        if (!FactReturnDate.IsNull()) TransferRowsCollection.Add(new CompleteTransferRow(UniversalDictionary, CompleteID, CompleteTransferRow.TransferTypes.ToWarehouse, CompleteTransferRow.Action.Return, ((DateTime)FactReturnDate).AddHours(1), CompleteCount, DocumentName));
                                        //MyMessageBox.Show("20.8");
                                        s = 30;
                                        break;
                                    case ExpositionID:
                                        //MyMessageBox.Show("20.9");
                                        s = 31;
                                        TransferRowsCollection.Add(new CompleteTransferRow(UniversalDictionary, CompleteID, CompleteTransferRow.TransferTypes.FromWarehouse, CompleteTransferRow.Action.DeliveryToExposition, (DateTime)FactShipDate, CompleteCount, DocumentName));
                                        //MyMessageBox.Show("20.10");
                                        s = 32;
                                        if (!FactReturnDate.IsNull()) TransferRowsCollection.Add(new CompleteTransferRow(UniversalDictionary, CompleteID, CompleteTransferRow.TransferTypes.ToWarehouse, CompleteTransferRow.Action.ReturnFromExposition, ((DateTime)FactReturnDate).AddHours(1), CompleteCount, DocumentName));
                                        //MyMessageBox.Show("20.11");
                                        s = 33;
                                        break;
                                    case TestDriveID:
                                        // MyMessageBox.Show("20.12");
                                        s = 34;
                                        TransferRowsCollection.Add(new CompleteTransferRow(UniversalDictionary, CompleteID, CompleteTransferRow.TransferTypes.FromWarehouse, CompleteTransferRow.Action.DeliveryToTestDrive, (DateTime)FactShipDate, CompleteCount, DocumentName));
                                        //MyMessageBox.Show("20.13");
                                        s = 35;
                                        if (!FactReturnDate.IsNull())
                                            TransferRowsCollection.Add(new CompleteTransferRow(UniversalDictionary, CompleteID, CompleteTransferRow.TransferTypes.ToWarehouse, CompleteTransferRow.Action.ReturnFromTestDrive, ((DateTime)FactReturnDate).AddHours(1), CompleteCount, DocumentName));
                                        //MyMessageBox.Show("20.14");
                                        s = 36;
                                        if (!PaymentDate.IsNull())
                                        {
                                            //MyMessageBox.Show("20.15");
                                            s = 37;
                                            TransferRowsCollection.Add(new CompleteTransferRow(UniversalDictionary, CompleteID, CompleteTransferRow.TransferTypes.ToWarehouse, CompleteTransferRow.Action.ReturnFromTestDrive, (DateTime)PaymentDate, CompleteCount, DocumentName));
                                            //MyMessageBox.Show("20.16");
                                            s = 38;
                                            TransferRowsCollection.Add(new CompleteTransferRow(UniversalDictionary, CompleteID, CompleteTransferRow.TransferTypes.FromWarehouse, CompleteTransferRow.Action.DeliveryNewDevices, ((DateTime)PaymentDate).AddHours(1), CompleteCount, DocumentName));
                                            //MyMessageBox.Show("20.17");
                                            s = 39;
                                        }
                                        break;
                                    case CertificationID:
                                        //MyMessageBox.Show("20.18");
                                        s = 40;
                                        TransferRowsCollection.Add(new CompleteTransferRow(UniversalDictionary, CompleteID, CompleteTransferRow.TransferTypes.FromWarehouse, CompleteTransferRow.Action.DeliveryToCertification, (DateTime)FactShipDate, CompleteCount, DocumentName));
                                        //MyMessageBox.Show("20.19");
                                        s = 41;
                                        if (!FactReturnDate.IsNull()) TransferRowsCollection.Add(new CompleteTransferRow(UniversalDictionary, CompleteID, CompleteTransferRow.TransferTypes.ToWarehouse, CompleteTransferRow.Action.ReturnFromCertification, ((DateTime)FactReturnDate).AddHours(1), CompleteCount, DocumentName));
                                        //MyMessageBox.Show("20.20");
                                        s = 42;
                                        break;
                                    case TenderID:
                                        //MyMessageBox.Show("20.21");
                                        s = 43;
                                        TransferRowsCollection.Add(new CompleteTransferRow(UniversalDictionary, CompleteID, CompleteTransferRow.TransferTypes.FromWarehouse, CompleteTransferRow.Action.DeliveryNewDevices, (DateTime)FactShipDate, CompleteCount, DocumentName));
                                        //MyMessageBox.Show("20.22");
                                        s = 44;
                                        break;
                                    case ActionID:
                                        //MyMessageBox.Show("20.23");
                                        s = 45;
                                        TransferRowsCollection.Add(new CompleteTransferRow(UniversalDictionary, CompleteID, CompleteTransferRow.TransferTypes.FromWarehouse, CompleteTransferRow.Action.DeliveryNewDevices, (DateTime)FactShipDate, CompleteCount, DocumentName));
                                        //MyMessageBox.Show("20.24");
                                        s = 46;
                                        if (!FactReturnDate.IsNull()) TransferRowsCollection.Add(new CompleteTransferRow(UniversalDictionary, CompleteID, CompleteTransferRow.TransferTypes.ToWarehouse, CompleteTransferRow.Action.Return, ((DateTime)FactReturnDate).AddHours(1), CompleteCount, DocumentName));
                                        //MyMessageBox.Show("20.25");
                                        s = 47;
                                        break;
                                        //MyMessageBox.Show("21");
                                }
                                //MyMessageBox.Show("22");
                            }
                            //MyMessageBox.Show("23");
                        }
                        //MyMessageBox.Show("24");
                    }
                    //MyMessageBox.Show("25");
                }
                //MyMessageBox.Show("26");
                return TransferRowsCollection.AsEnumerable();
            }
            catch
            {
                MyMessageBox.Show("Произошла ошибка при преобразовании Задания на комплектацию в операции передачи комплектующих. Задание на комплектацию: " + CompleteRow.GetObject(RefCompleteCard.Devices.FactShipDate).ToString() + ", " + CompleteRow.GetObject(RefCompleteCard.Devices.DeviceNumber) + ", s=" + s + ", Complete=" + MyCompleteName);
                return null;
            }
        }

        /// <summary>
        /// Поиск Договоров с зарезервированными комплектующими.
        /// </summary>
        /// <param name="session">Пользовательская сессия DV.</param>
        /// <param name="UniversalDictionary">Карточка универсального справочника.</param>
        /// <returns></returns>
        public static IEnumerable<ReservedCompleteRow> FindAccounts(UserSession session, CardData UniversalDictionary)
        {
            List<ReservedCompleteRow> Result = new List<ReservedCompleteRow>();

            // Поиск Договоров/счетов, в которых зарезервированы комплектующие
            SearchQuery searchQuery = session.CreateSearchQuery();
            searchQuery.CombineResults = ConditionGroupOperation.And;

            CardTypeQuery typeQuery = searchQuery.AttributiveSearch.CardTypeQueries.AddNew(RefAccountCard.ID);

            // Задание на отгрузку еще не создано
            SectionQuery sectionQuery = typeQuery.SectionQueries.AddNew(RefAccountCard.Devices.ID);
            sectionQuery.Operation = SectionQueryOperation.And;
            sectionQuery.ConditionGroup.Operation = ConditionGroupOperation.And;
            sectionQuery.ConditionGroup.Conditions.AddNew(RefAccountCard.Devices.ToShip, FieldType.Int, ConditionOperation.GreaterThan, 0);

            // Получение текста запроса
            searchQuery.Limit = 0;
            string query = searchQuery.GetXml();

            CardDataCollection AccountCollection = session.CardManager.FindCards(query);
            //MyMessageBox.Show("Нашли Договоры, у которых есть неотгруженные позиции..." + AccountCollection.Count());

            // Получение конвертированных строк Договора
            foreach (CardData Account in AccountCollection)
            {
                IEnumerable<RowData> FindRows = Account.Sections[RefAccountCard.Devices.ID].Rows;
                IEnumerable<RowData> CheckRows = FindRows.Where(r => r.GetInt32(RefAccountCard.Devices.Shipped) < r.GetInt32(RefAccountCard.Devices.Count));
                IEnumerable<ReservedCompleteRow> ConvertRows = CheckRows.SelectMany(r => ConvertCompleteToReserverdRowsCollection(r, UniversalDictionary,
                    Account.Sections[RefAccountCard.MainInfo.ID].FirstRow.GetString(RefAccountCard.MainInfo.ContractSubjectId).ToString(), "Договор/счет " +
                    Account.Sections[RefAccountCard.MainInfo.ID].FirstRow.GetString(RefAccountCard.MainInfo.Number)));
                if (!ConvertRows.IsNull() && ConvertRows.Count() > 0)
                {
                    List<ReservedCompleteRow> R = ConvertRows.ToList();
                    Result.AddRange(R);
                }
            }
            int ResultCount = Result.Count();

            Clear();
            return Result;
        }

        /// <summary>
        /// Преобразование строки таблицы "Приборы и комплектующие" Договора в коллекцию зарезервированных комплектующих.
        /// </summary>
        /// <param name="AccountCardRow">Строка таблицы "Приборы и комплектующие" карточки Договора/счета </param>
        /// <param name="UniversalDictionary">Карточка универсального справочника.</param>
        /// <param name="ContractSubjectId">Идентификатор предмета договора/счета.</param>
        /// <param name="DocumentName">Название документа основания.</param>
        /// <returns></returns>
        private static IEnumerable<ReservedCompleteRow> ConvertCompleteToReserverdRowsCollection(RowData AccountCardRow, CardData UniversalDictionary, string ContractSubjectId, string DocumentName) //, CardDataCollection ShipmentTaskCollection, CardDataCollection AccountCardCollection)
        {
            List<ReservedCompleteRow> TransferRowsCollection = new List<ReservedCompleteRow>();

            // Получаем справочник приборов и комплектующих для прибора текущей строки
            object DeviceID = AccountCardRow.GetObject(RefAccountCard.Devices.DeviceId);
            Guid ParentItemTypeID = UniversalDictionary.GetItemRow(DeviceID.ToGuid()).GetString("ParentRowID").ToGuid();
            string DeviceName = UniversalDictionary.GetItemName(DeviceID.ToGuid());
            RowData DeviceDictionary = UniversalDictionary.GetItemTypeRow(ParentItemTypeID).ChildRows.FirstOrDefault(r => r.GetString("Name") == DeviceName);
            if (DeviceDictionary.IsNull())
                return TransferRowsCollection;

            string MyCompleteName = "";
            try
            {
                string PackedListData = AccountCardRow.GetString(RefAccountCard.Devices.PackedListData);
                string[] Completes = PackedListData.Split('\n');

                foreach (string Complete in Completes)
                {
                    DeviceCompleteRow DevCompleteRow = (DeviceCompleteRow)Complete;
                    string CompleteName = DevCompleteRow.Name;
                    int CompleteType = DevCompleteRow.TableType;
                    int CompleteCount = DevCompleteRow.Count.IsNull() ? 0 : Convert.ToInt32(DevCompleteRow.Count);
                    MyCompleteName = CompleteName;
                    if (CompleteCount > 0)
                    {
                        RowData CompleteRowData = DeviceDictionary.ChildSections[new Guid("{DD20BF9B-90F8-4D9A-9553-5B5F17AD724E}")].Rows.FirstOrDefault(r => r.GetString("Name") == CompleteName);
                        if (!CompleteRowData.IsNull())
                        {
                            Guid CompleteID = CompleteRowData.Id;
                            string ContractSubject = "";
                            switch (ContractSubjectId)
                            {
                                case DeliveryID:
                                    {
                                        ContractSubject = " (поставка готовой продукции)";
                                        break;
                                    }
                                case ExpositionID:
                                    {
                                        ContractSubject = " (отправка на выставку)";
                                        break;
                                    }
                                case TestDriveID:
                                    {
                                        ContractSubject = " (отправка на тест-драйв)";
                                        break;
                                    }
                                case CertificationID:
                                    {
                                        ContractSubject = " (отправка на сертификацию)";
                                        break;
                                    }
                                case TenderID:
                                    {
                                        ContractSubject = " (тендер)";
                                        break;
                                    }
                                case ActionID:
                                    {
                                        ContractSubject = " (акция)";
                                        break;
                                    }
                                case ServiceDeliveryID:
                                    {
                                        ContractSubject = " (сервисное обслуживание + поставка готовой продукции)";
                                        break;
                                    }
                                case SeminarID:
                                    {
                                        ContractSubject = " (семинар)";
                                        break;
                                    }
                            }

                            TransferRowsCollection.Add(new ReservedCompleteRow(DeviceID.ToGuid(), CompleteID, UniversalDictionary.GetItemPropertyValue(CompleteID, "Код СКБ").ToGuid(),
                                CompleteCount * ((int)AccountCardRow.GetInt32(RefAccountCard.Devices.Count) - (int)AccountCardRow.GetInt32(RefAccountCard.Devices.Shipped)), DocumentName + ContractSubject));
                        }
                    }
                }
                return TransferRowsCollection.AsEnumerable();
            }
            catch
            {
                MyMessageBox.Show("Произошла ошибка при преобразовании Договора в операции резерва комплектующих. Договор: " + DocumentName + ", Complete=" + MyCompleteName);
                return null;
            }
        }

        /// <summary>
        /// Поиск Заданий на комплектацию, в которых не указана фактическая дата отгрузки.
        /// </summary>
        /// <param name="session">Пользовательская сессия DV.</param>
        /// <param name="UniversalDictionary">Карточка универсального справочника.</param>
        /// <returns></returns>
        public static IEnumerable<ReservedCompleteRow> FindTaskCompleteWithoutShipment(UserSession session, CardData UniversalDictionary)
        {
            List<ReservedCompleteRow> Result = new List<ReservedCompleteRow>();

            // Поиск Заданий на комплектацию, у которых дата отгрузки не указана
            SearchQuery searchQuery = session.CreateSearchQuery();
            searchQuery.CombineResults = ConditionGroupOperation.And;

            CardTypeQuery typeQuery = searchQuery.AttributiveSearch.CardTypeQueries.AddNew(RefCompleteCard.ID);

            // Фактическая дата отгрузки попадает в заданный интервал
            SectionQuery sectionQuery = typeQuery.SectionQueries.AddNew(RefCompleteCard.Devices.ID);
            sectionQuery.Operation = SectionQueryOperation.And;
            sectionQuery.ConditionGroup.Operation = ConditionGroupOperation.And;
            sectionQuery.ConditionGroup.Conditions.AddNew(RefCompleteCard.Devices.FactShipDate, FieldType.DateTime, ConditionOperation.IsNull);

            // Получение текста запроса
            searchQuery.Limit = 0;
            string query = searchQuery.GetXml();

            CardDataCollection CompleteTaskCollection = session.CardManager.FindCards(query);
            //MyMessageBox.Show("Нашли Задания на комплектацию, у которых дата отгрузки не указана:" + CompleteTaskCollection.Count());

            // Получение конвертированных строк Заявок на комплектацию (у которых дата отгрузки попадает в заданный интервал)



            foreach (CardData CompleteTask in CompleteTaskCollection)
            {
                IEnumerable<RowData> FindRows = CompleteTask.Sections[RefCompleteCard.Devices.ID].Rows;
                IEnumerable<RowData> CheckRows = FindRows.Where(r => r.GetDateTime(RefCompleteCard.Devices.FactShipDate).IsNull());
                IEnumerable<ReservedCompleteRow> ConvertRows = CheckRows.SelectMany(r => ConvertTaskCompleteToReserverdRowsCollection(r, session, UniversalDictionary, CompleteTask.Description));
                List<ReservedCompleteRow> R = ConvertRows.ToList();
                Result.AddRange(R);
            }
            int ResultCount = Result.Count();

            Clear();
            //MyMessageBox.Show("Всего строк по незавершенной отгрузке: " + Result.Count());
            return Result;
        }

        /// <summary>
        /// Преобразование строки таблицы "Информация по приборам" в коллекцию операций передачи приборов.
        /// </summary>
        /// <param name="CompleteRow">Строка таблицы "Информация по приборам."</param>
        /// <param name="session">Сессия."</param>
        /// <param name="UniversalDictionary">Карточка универсального справочника.</param>
        /// <param name="DocumentName">Название документа основания.</param>
        /// <returns></returns>
        private static IEnumerable<ReservedCompleteRow> ConvertTaskCompleteToReserverdRowsCollection(RowData CompleteRow, UserSession session, CardData UniversalDictionary, string DocumentName)//, CardDataCollection ShipmentTaskCollection, CardDataCollection AccountCardCollection)
        {
            string MyCompleteName = "";
            try
            {
                List<ReservedCompleteRow> TransferRowsCollection = new List<ReservedCompleteRow>();

                string ContractSubjectId = CompleteRow.GetString(RefCompleteCard.Devices.ContractSubjectId);
                CompleteRow.GetGuid(RefCompleteCard.Devices.ShipmentTaskId);
                object DeviceID = CompleteRow.GetObject(RefCompleteCard.Devices.DeviceId);
                Guid ParentItemTypeID = UniversalDictionary.GetItemRow(DeviceID.ToGuid()).GetString("ParentRowID").ToGuid();
                string DeviceName = UniversalDictionary.GetItemName(DeviceID.ToGuid());
                RowData DeviceDictionary = UniversalDictionary.GetItemTypeRow(ParentItemTypeID).ChildRows.FirstOrDefault(r => r.GetString("Name") == DeviceName);
                if (DeviceDictionary.IsNull())
                    return TransferRowsCollection;

                // Если предмет договора счета - НЕ Сервисное обслуживание/Семинар
                if (ContractSubjectId != ServiceID && ContractSubjectId != SeminarID)
                {
                    // Поиск Задания на отгрузку
                    SearchQuery searchQuery2 = session.CreateSearchQuery();
                    searchQuery2.CombineResults = ConditionGroupOperation.And;

                    CardTypeQuery typeQuery2 = searchQuery2.AttributiveSearch.CardTypeQueries.AddNew(RefShipmentCard.ID);

                    // Фактическая дата отгрузки меньше конечной даты
                    SectionQuery sectionQuery2 = typeQuery2.SectionQueries.AddNew(RefShipmentCard.MainInfo.ID);
                    sectionQuery2.Operation = SectionQueryOperation.And;
                    sectionQuery2.ConditionGroup.Operation = ConditionGroupOperation.And;
                    sectionQuery2.ConditionGroup.Conditions.AddNew("InstanceID", FieldType.UniqueId, ConditionOperation.Equals, CompleteRow.GetString(RefCompleteCard.Devices.ShipmentTaskId).ToGuid());
                    // Получение текста запроса
                    searchQuery2.Limit = 0;
                    string query2 = searchQuery2.GetXml();

                    CardDataCollection ShipmentTaskCollection = session.CardManager.FindCards(query2);

                    if (ShipmentTaskCollection.IsNull())
                        return TransferRowsCollection;
                    CardData ShipmentTask = ShipmentTaskCollection.FirstOrDefault();
                    if (ShipmentTask.IsNull())
                        return TransferRowsCollection;

                    RowData ShipmentTaskRow = ShipmentTask.Sections[RefShipmentCard.Devices.ID].Rows.FirstOrDefault(r => r.GetString(RefShipmentCard.Devices.Id).ToGuid().Equals(CompleteRow.GetString(RefCompleteCard.Devices.ShipmentTaskRowId).ToGuid()));

                    if (ShipmentTaskRow.IsNull())
                        return TransferRowsCollection;

                    // Поиск Договоров/счетов
                    SearchQuery searchQuery = session.CreateSearchQuery();
                    searchQuery.CombineResults = ConditionGroupOperation.And;

                    CardTypeQuery typeQuery = searchQuery.AttributiveSearch.CardTypeQueries.AddNew(RefAccountCard.ID);

                    // Фактическая дата отгрузки меньше конечной даты
                    SectionQuery sectionQuery = typeQuery.SectionQueries.AddNew(RefAccountCard.MainInfo.ID);
                    sectionQuery.Operation = SectionQueryOperation.And;
                    sectionQuery.ConditionGroup.Operation = ConditionGroupOperation.And;
                    sectionQuery.ConditionGroup.Conditions.AddNew(RefAccountCard.MainInfo.Number, FieldType.String, ConditionOperation.Equals, ShipmentTask.Sections[RefShipmentCard.MainInfo.ID].FirstRow.GetString(RefShipmentCard.MainInfo.AccountNumber).ToString());
                    sectionQuery.ConditionGroup.Conditions.AddNew(RefAccountCard.MainInfo.ContractDate, FieldType.Date, ConditionOperation.Equals, (DateTime)ShipmentTask.Sections[RefShipmentCard.MainInfo.ID].FirstRow.GetDateTime(RefShipmentCard.MainInfo.AccountDate));
                    // Получение текста запроса
                    searchQuery.Limit = 0;
                    string query = searchQuery.GetXml();

                    CardDataCollection FindAccountCardCollection = session.CardManager.FindCards(query);

                    CardData AccountCard = FindAccountCardCollection.FirstOrDefault();

                    if (AccountCard.IsNull())
                        return TransferRowsCollection;

                    RowData AccountCardRow = AccountCard.Sections[RefAccountCard.Devices.ID].Rows.FirstOrDefault(r => r.GetString(RefAccountCard.Devices.Id).ToGuid().Equals(ShipmentTaskRow.GetString(RefShipmentCard.Devices.AccountCardRowId).ToGuid()));

                    if (AccountCardRow.IsNull())
                        return TransferRowsCollection;

                    string PackedListData = AccountCardRow.GetString(RefAccountCard.Devices.PackedListData);

                    string[] Completes = PackedListData.Split('\n');

                    foreach (string Complete in Completes)
                    {
                        DeviceCompleteRow DevCompleteRow = (DeviceCompleteRow)Complete;
                        string CompleteName = DevCompleteRow.Name;
                        int CompleteType = DevCompleteRow.TableType;
                        int CompleteCount = DevCompleteRow.Count.IsNull() ? 0 : Convert.ToInt32(DevCompleteRow.Count);
                        MyCompleteName = CompleteName;
                        if (CompleteCount > 0)
                        {
                            RowData CompleteRowData = DeviceDictionary.ChildSections[new Guid("{DD20BF9B-90F8-4D9A-9553-5B5F17AD724E}")].Rows.FirstOrDefault(r => r.GetString("Name") == CompleteName);

                            if (!CompleteRowData.IsNull())
                            {
                                Guid CompleteID = CompleteRowData.Id;
                                string ContractSubject = "";
                                switch (ContractSubjectId)
                                {
                                    case DeliveryID:
                                        {
                                            ContractSubject = " (поставка готовой продукции)";
                                            break;
                                        }
                                    case ExpositionID:
                                        {
                                            ContractSubject = " (отправка на выставку)";
                                            break;
                                        }
                                    case TestDriveID:
                                        {
                                            ContractSubject = " (отправка на тест-драйв)";
                                            break;
                                        }
                                    case CertificationID:
                                        {
                                            ContractSubject = " (отправка на сертификацию)";
                                            break;
                                        }
                                    case TenderID:
                                        {
                                            ContractSubject = " (тендер)";
                                            break;
                                        }
                                    case ActionID:
                                        {
                                            ContractSubject = " (акция)";
                                            break;
                                        }
                                    case ServiceDeliveryID:
                                        {
                                            ContractSubject = " (сервисное обслуживание + поставка готовой продукции)";
                                            break;
                                        }
                                    case SeminarID:
                                        {
                                            ContractSubject = " (семинар)";
                                            break;
                                        }
                                }
                                TransferRowsCollection.Add(new ReservedCompleteRow(DeviceID.ToGuid(), CompleteID, UniversalDictionary.GetItemPropertyValue(CompleteID, "Код СКБ").ToGuid(),
                                CompleteCount * ((int)AccountCardRow.GetInt32(RefAccountCard.Devices.Count) - (int)AccountCardRow.GetInt32(RefAccountCard.Devices.Shipped)), DocumentName + ContractSubject));
                            }
                        }
                    }
                }
                return TransferRowsCollection.AsEnumerable();
            }
            catch
            {
                MyMessageBox.Show("Произошла ошибка при преобразовании Задания на комплектацию в операции передачи комплектующих. Задание на комплектацию: " + DocumentName + ", Complete=" + MyCompleteName);
                return null;
            }
        }

        /// <summary>
        /// Получение статистики на конкретную дату
        /// </summary>
        /// <param name="AllTransfers">Перечень всех операций передач приборов.</param>
        /// <param name="KeyDate">Дата.</param>
        /// <param name="TransferType">Тип передач.</param>
        /// <param name="Action">Действие с прибором.</param>
        /// <returns></returns>
        public static IEnumerable<TransferCountByCompleteType> StatisticsOnDate(this IEnumerable<CompleteTransferRow> AllTransfers, DateTime KeyDate,
            CompleteTransferRow.TransferTypes TransferType, CompleteTransferRow.Action Action = CompleteTransferRow.Action.None)
        {
            // Перечень операций с комплектующими на указанную дату
            IEnumerable<CompleteTransferRow> AllCompleteActionsOnDate = AllTransfers.Where(r => r.TransferDate < KeyDate);
            // Группировка операций по типу комплектующего

            IEnumerable<TransferCountByCompleteType> AllCompleteOnDate = null;
            switch (Action)
            {
                case CompleteTransferRow.Action.None:
                    AllCompleteOnDate = AllCompleteActionsOnDate.GroupBy(r => r.CompleteID, (completeID, transfers) =>
                    new TransferCountByCompleteType(
                        transfers.First().CompleteID,
                        transfers.First().CompleteCodeID,
                        transfers.First().ParentDeviceID,
                        transfers.First().СompleteType,
                        transfers.First().CompleteCode,
                        transfers.First().ParentDevice,
                        transfers.Select<CompleteTransferRow, Int32>(s => s.CompleteCalc).Aggregate(0, (start, next) => start + next),
                        transfers.First().DocumentName));
                    break;
                case CompleteTransferRow.Action.DeliveryToCertification:
                    AllCompleteOnDate = AllCompleteActionsOnDate.Where(r => (r.TransferAction == CompleteTransferRow.Action.DeliveryToCertification) ||
                    (r.TransferAction == CompleteTransferRow.Action.ReturnFromCertification)).GroupBy(r => r.CompleteID, (completeID, transfers) =>
                    new TransferCountByCompleteType(
                        transfers.First().CompleteID,
                        transfers.First().CompleteCodeID,
                        transfers.First().ParentDeviceID,
                        transfers.First().СompleteType,
                        transfers.First().CompleteCode,
                        transfers.First().ParentDevice,
                        transfers.Select<CompleteTransferRow, Int32>(s => s.CompleteCalc).Aggregate(0, (start, next) => start + next) * (-1),
                        transfers.First().DocumentName));
                    break;
                case CompleteTransferRow.Action.DeliveryToExposition:
                    AllCompleteOnDate = AllCompleteActionsOnDate.Where(r => (r.TransferAction == CompleteTransferRow.Action.DeliveryToExposition) ||
                    (r.TransferAction == CompleteTransferRow.Action.ReturnFromExposition)).GroupBy(r => r.CompleteID, (completeID, transfers) =>
                    new TransferCountByCompleteType(
                        transfers.First().CompleteID,
                        transfers.First().CompleteCodeID,
                        transfers.First().ParentDeviceID,
                        transfers.First().СompleteType,
                        transfers.First().CompleteCode,
                        transfers.First().ParentDevice,
                        transfers.Select<CompleteTransferRow, Int32>(s => s.CompleteCalc).Aggregate(0, (start, next) => start + next) * (-1),
                        transfers.First().DocumentName));
                    return AllCompleteOnDate;
                case CompleteTransferRow.Action.DeliveryToTestDrive:
                    AllCompleteOnDate = AllCompleteActionsOnDate.Where(r => (r.TransferAction == CompleteTransferRow.Action.DeliveryToTestDrive) ||
                    (r.TransferAction == CompleteTransferRow.Action.ReturnFromTestDrive)).GroupBy(r => r.CompleteID, (completeID, transfers) =>
                    new TransferCountByCompleteType(
                        transfers.First().CompleteID,
                        transfers.First().CompleteCodeID,
                        transfers.First().ParentDeviceID,
                        transfers.First().СompleteType,
                        transfers.First().CompleteCode,
                        transfers.First().ParentDevice,
                        transfers.Select<CompleteTransferRow, Int32>(s => s.CompleteCalc).Aggregate(0, (start, next) => start + next) * (-1),
                        transfers.First().DocumentName));
                    return AllCompleteOnDate;
                case CompleteTransferRow.Action.DeliveryToTesting:
                    AllCompleteOnDate = AllCompleteActionsOnDate.Where(r => (r.TransferAction == CompleteTransferRow.Action.DeliveryToTesting) ||
                    (r.TransferAction == CompleteTransferRow.Action.ReturnFromTesting)).GroupBy(r => r.CompleteID, (completeID, transfers) =>
                    new TransferCountByCompleteType(
                        transfers.First().CompleteID,
                        transfers.First().CompleteCodeID,
                        transfers.First().ParentDeviceID,
                        transfers.First().СompleteType,
                        transfers.First().CompleteCode,
                        transfers.First().ParentDevice,
                        transfers.Select<CompleteTransferRow, Int32>(s => s.CompleteCalc).Aggregate(0, (start, next) => start + next) * (-1),
                        transfers.First().DocumentName));
                    return AllCompleteOnDate;
            }
            return AllCompleteOnDate;
        }
        /// <summary>
        /// Получение статистики на конкретную дату
        /// </summary>
        /// <param name="AllTransfers">Перечень всех операций передач приборов.</param>
        /// <param name="KeyDate">Дата.</param>
        /// <param name="TransferType">Тип передач.</param>
        /// <param name="Action">Действие с прибором.</param>
        /// <param name="PreviousReport">Предыдущий отчет по остаткам.</param>
        /// <returns></returns>
        public static IEnumerable<TransferCountByCompleteType> StatisticsOnDate(this IEnumerable<CompleteTransferRow> AllTransfers, DateTime KeyDate,
            CompleteTransferRow.TransferTypes TransferType, BalanceOfCompleteItem PreviousReport, CompleteTransferRow.Action Action = CompleteTransferRow.Action.None)
        {
            // Перечень операций с комплектующими на указанную дату
            IEnumerable<CompleteTransferRow> AllCompleteActionsOnDate = AllTransfers.Where(r => r.TransferDate < KeyDate);
            // Группировка операций по типу комплектующего

            IEnumerable<TransferCountByCompleteType> AllCompleteOnDate = null;
            switch (Action)
            {
                case CompleteTransferRow.Action.None:
                    AllCompleteOnDate = AllCompleteActionsOnDate.GroupBy(r => r.CompleteID, (completeID, transfers) =>
                    new TransferCountByCompleteType(
                        transfers.First().CompleteID,
                        transfers.First().CompleteCodeID,
                        transfers.First().ParentDeviceID,
                        transfers.First().СompleteType,
                        transfers.First().CompleteCode,
                        transfers.First().ParentDevice,
                        transfers.Select<CompleteTransferRow, Int32>(s => s.CompleteCalc).Aggregate(0, (start, next) => start + next),
                        transfers.First().DocumentName));
                    // Добавляем остаток на конец прошлого периода
                    foreach (TransferCountByCompleteType CompleteRow in AllCompleteOnDate)
                    {
                        BalanceOfCompleteRowItem FindRowItem = PreviousReport.balanceOfCompleteTable.FirstOrDefault(r =>
                        (r.completeID == CompleteRow.CompleteID) && (r.allocationID == Allocation.InWarehouse));
                        if (!FindRowItem.IsNull())
                        { CompleteRow.CompleteCount = CompleteRow.CompleteCount + FindRowItem.endCount; }
                    }
                    break;
                case CompleteTransferRow.Action.DeliveryToCertification:
                    AllCompleteOnDate = AllCompleteActionsOnDate.Where(r => (r.TransferAction == CompleteTransferRow.Action.DeliveryToCertification) ||
                    (r.TransferAction == CompleteTransferRow.Action.ReturnFromCertification)).GroupBy(r => r.CompleteID, (completeID, transfers) =>
                    new TransferCountByCompleteType(
                        transfers.First().CompleteID,
                        transfers.First().CompleteCodeID,
                        transfers.First().ParentDeviceID,
                        transfers.First().СompleteType,
                        transfers.First().CompleteCode,
                        transfers.First().ParentDevice,
                        transfers.Select<CompleteTransferRow, Int32>(s => s.CompleteCalc).Aggregate(0, (start, next) => start + next) * (-1),
                        transfers.First().DocumentName));
                    // Добавляем остаток на конец прошлого периода
                    foreach (TransferCountByCompleteType CompleteRow in AllCompleteOnDate)
                    {
                        BalanceOfCompleteRowItem FindRowItem = PreviousReport.balanceOfCompleteTable.FirstOrDefault(r =>
                        (r.completeID == CompleteRow.CompleteID) && (r.allocationID == Allocation.InCertification));
                        if (!FindRowItem.IsNull())
                        { CompleteRow.CompleteCount = CompleteRow.CompleteCount + FindRowItem.endCount; }
                    }
                    break;
                case CompleteTransferRow.Action.DeliveryToExposition:
                    AllCompleteOnDate = AllCompleteActionsOnDate.Where(r => (r.TransferAction == CompleteTransferRow.Action.DeliveryToExposition) ||
                    (r.TransferAction == CompleteTransferRow.Action.ReturnFromExposition)).GroupBy(r => r.CompleteID, (completeID, transfers) =>
                    new TransferCountByCompleteType(
                        transfers.First().CompleteID,
                        transfers.First().CompleteCodeID,
                        transfers.First().ParentDeviceID,
                        transfers.First().СompleteType,
                        transfers.First().CompleteCode,
                        transfers.First().ParentDevice,
                        transfers.Select<CompleteTransferRow, Int32>(s => s.CompleteCalc).Aggregate(0, (start, next) => start + next) * (-1),
                        transfers.First().DocumentName));

                    // Добавляем остаток на конец прошлого периода
                    foreach (TransferCountByCompleteType CompleteRow in AllCompleteOnDate)
                    {
                        BalanceOfCompleteRowItem FindRowItem = PreviousReport.balanceOfCompleteTable.FirstOrDefault(r =>
                        (r.completeID == CompleteRow.CompleteID) && (r.allocationID == Allocation.InExposition));
                        if (!FindRowItem.IsNull())
                        { CompleteRow.CompleteCount = CompleteRow.CompleteCount + FindRowItem.endCount; }
                    }
                    return AllCompleteOnDate;
                case CompleteTransferRow.Action.DeliveryToTestDrive:
                    AllCompleteOnDate = AllCompleteActionsOnDate.Where(r => (r.TransferAction == CompleteTransferRow.Action.DeliveryToTestDrive) ||
                    (r.TransferAction == CompleteTransferRow.Action.ReturnFromTestDrive)).GroupBy(r => r.CompleteID, (completeID, transfers) =>
                    new TransferCountByCompleteType(
                        transfers.First().CompleteID,
                        transfers.First().CompleteCodeID,
                        transfers.First().ParentDeviceID,
                        transfers.First().СompleteType,
                        transfers.First().CompleteCode,
                        transfers.First().ParentDevice,
                        transfers.Select<CompleteTransferRow, Int32>(s => s.CompleteCalc).Aggregate(0, (start, next) => start + next) * (-1),
                        transfers.First().DocumentName));
                    // Добавляем остаток на конец прошлого периода
                    foreach (TransferCountByCompleteType CompleteRow in AllCompleteOnDate)
                    {
                        BalanceOfCompleteRowItem FindRowItem = PreviousReport.balanceOfCompleteTable.FirstOrDefault(r =>
                        (r.completeID == CompleteRow.CompleteID) && (r.allocationID == Allocation.InTestDrive));
                        if (!FindRowItem.IsNull())
                        { CompleteRow.CompleteCount = CompleteRow.CompleteCount + FindRowItem.endCount; }
                    }
                    return AllCompleteOnDate;
                case CompleteTransferRow.Action.DeliveryToTesting:
                    AllCompleteOnDate = AllCompleteActionsOnDate.Where(r => (r.TransferAction == CompleteTransferRow.Action.DeliveryToTesting) ||
                    (r.TransferAction == CompleteTransferRow.Action.ReturnFromTesting)).GroupBy(r => r.CompleteID, (completeID, transfers) =>
                    new TransferCountByCompleteType(
                        transfers.First().CompleteID,
                        transfers.First().CompleteCodeID,
                        transfers.First().ParentDeviceID,
                        transfers.First().СompleteType,
                        transfers.First().CompleteCode,
                        transfers.First().ParentDevice,
                        transfers.Select<CompleteTransferRow, Int32>(s => s.CompleteCalc).Aggregate(0, (start, next) => start + next) * (-1),
                        transfers.First().DocumentName));

                    // Добавляем остаток на конец прошлого периода
                    foreach (TransferCountByCompleteType CompleteRow in AllCompleteOnDate)
                    {
                        BalanceOfCompleteRowItem FindRowItem = PreviousReport.balanceOfCompleteTable.FirstOrDefault(r =>
                        (r.completeID == CompleteRow.CompleteID) && (r.allocationID == Allocation.InTesting));
                        if (!FindRowItem.IsNull())
                        { CompleteRow.CompleteCount = CompleteRow.CompleteCount + FindRowItem.endCount; }
                    }
                    return AllCompleteOnDate;
            }
            return AllCompleteOnDate;
        }
        /// <summary>
        /// Получение статистики за период
        /// </summary>
        /// <param name="AllTransfers">Перечень всех операций передач приборов.</param>
        /// <param name="StartDate">Начальная дата.</param>
        /// <param name="EndDate">Конечная дата.</param>
        /// <param name="Action">Действие с прибором</param>
        /// <param name="UniversalDictionary">Карточка универсального справочника.</param>
        /// <returns></returns>
        public static IEnumerable<TransferCountByCompleteType> StatisticsForPeriod(this IEnumerable<CompleteTransferRow> AllTransfers, DateTime StartDate, DateTime EndDate, CompleteTransferRow.Action Action, CardData UniversalDictionary)
        {
            // Перечень всех операций с заданным действием за указанный период
            IEnumerable<CompleteTransferRow> AllActionsForPeriod = AllTransfers.Where(r => r.TransferAction == Action && r.TransferDate >= StartDate && r.TransferDate < EndDate.AddDays(1));

            // Группировка операций по типу комплектующего
            IEnumerable<TransferCountByCompleteType> AllCompleteForPeriod = AllActionsForPeriod.GroupBy(r => r.CompleteID, (completeID, transfers) =>
               new TransferCountByCompleteType(
                   transfers.First().CompleteID,
                   transfers.First().CompleteCodeID,
                   transfers.First().ParentDeviceID,
                   transfers.First().СompleteType,
                   transfers.First().CompleteCode,
                   transfers.First().ParentDevice,
                   transfers.Select<CompleteTransferRow, Int32>(s => s.CompleteCount).Aggregate(0, (start, next) => start + next),
                   transfers.First().DocumentName));

            return AllCompleteForPeriod;
        }

        /// <summary>
        /// Подсчет прихода комплектующих за период
        /// </summary>
        /// <param name="CurrentBalanceCompleteCollection">Коллекция остатков комплектующих.</param>
        /// <param name="TransferCollection">Перечень приходных операций.</param>
        public static void ReceivedCalculation(this List<CurrentBalanceComplete> CurrentBalanceCompleteCollection, IEnumerable<TransferCountByCompleteType> TransferCollection)
        {
            foreach (TransferCountByCompleteType CurrentRow in TransferCollection)
            {
                CurrentBalanceComplete CurrentBalanceCompleteRow = null;
                if (CurrentBalanceCompleteCollection.Any(s => s.CompleteID.Equals(CurrentRow.CompleteID)))
                {
                    CurrentBalanceCompleteRow = CurrentBalanceCompleteCollection.First(s => s.CompleteID.Equals(CurrentRow.CompleteID));
                }
                else
                {
                    //CurrentBalanceCompleteRow = new CurrentBalanceComplete(CurrentRow.ParentDeviceID, CurrentRow.CompleteID, CurrentRow.CompleteCodeID, 0, 0, 0, 0);
                    //CurrentBalanceCompleteCollection.Add(CurrentBalanceCompleteRow);
                }
                CurrentBalanceCompleteRow.Received = CurrentBalanceCompleteRow.Received + CurrentRow.CompleteCount;
                CurrentBalanceCompleteRow.EndBalance = CurrentBalanceCompleteRow.EndBalance + CurrentRow.CompleteCount;
                CurrentBalanceCompleteRow.ReceivedDocuments.Add(CurrentRow.DocumentName);
            }
        }
        /// <summary>
        /// Подсчет расхода комплектующих за период
        /// </summary>
        /// <param name="CurrentBalanceCompleteCollection"> Коллекция остатков комплектующих.</param>
        /// <param name="TransferCollection"> Перечень расходных операций. </param>
        public static void DescendedCalculation(this List<CurrentBalanceComplete> CurrentBalanceCompleteCollection, IEnumerable<TransferCountByCompleteType> TransferCollection)
        {
            foreach (TransferCountByCompleteType CurrentRow in TransferCollection)
            {
                CurrentBalanceComplete CurrentBalanceCompleteRow = null;
                if (CurrentBalanceCompleteCollection.Any(s => s.CompleteID.Equals(CurrentRow.CompleteID)))
                {
                    CurrentBalanceCompleteRow = CurrentBalanceCompleteCollection.First(s => s.CompleteID.Equals(CurrentRow.CompleteID));
                }
                else
                {
                    //CurrentBalanceCompleteRow = new CurrentBalanceComplete(CurrentRow.ParentDeviceID, CurrentRow.CompleteID, CurrentRow.CompleteCodeID, 0, 0, 0, 0);
                    //CurrentBalanceCompleteCollection.Add(CurrentBalanceCompleteRow);
                }
                CurrentBalanceCompleteRow.Descended = CurrentBalanceCompleteRow.Descended + CurrentRow.CompleteCount;
                CurrentBalanceCompleteRow.EndBalance = CurrentBalanceCompleteRow.EndBalance - CurrentRow.CompleteCount;
                CurrentBalanceCompleteRow.DescendedDocuments.Add(CurrentRow.DocumentName);
            }
        }
        /// <summary>
        /// Подсчет зарезервированных комплектующих на текущий момент
        /// </summary>
        /// <param name="CurrentBalanceCompleteCollection"> Коллекция остатков комплектующих.</param>
        /// <param name="ReservedCollection"> Перечень операций резервирования. </param>
        public static void ReservedCalculation(this List<CurrentBalanceComplete> CurrentBalanceCompleteCollection, IEnumerable<ReservedCompleteRow> ReservedCollection)
        {
            foreach (ReservedCompleteRow CurrentRow in ReservedCollection)
            {
                CurrentBalanceComplete CurrentBalanceCompleteRow = null;
                if (CurrentBalanceCompleteCollection.Any(s => s.CompleteID.Equals(CurrentRow.CompleteID)))
                {
                    CurrentBalanceCompleteRow = CurrentBalanceCompleteCollection.First(s => s.CompleteID.Equals(CurrentRow.CompleteID));
                }
                else
                {
                    //CurrentBalanceCompleteRow = new CurrentBalanceComplete(CurrentRow.DeviceID, CurrentRow.CompleteID, CurrentRow.CodeSKB, 0, 0, 0, 0);
                    //CurrentBalanceCompleteCollection.Add(CurrentBalanceCompleteRow);
                }
                CurrentBalanceCompleteRow.Reserved = CurrentBalanceCompleteRow.Reserved + CurrentRow.Count;
                //CurrentBalanceCompleteRow.EndBalance = CurrentBalanceCompleteRow.EndBalance - CurrentRow.Count;
                CurrentBalanceCompleteRow.ReservedDocuments.Add(CurrentRow.DocumentName);
            }
        }
    }
    /// <summary>
    /// Зарезервированные комплектующие.
    /// </summary>
    public class ReservedCompleteRow
    {
        /// <summary>
        /// Идентификатор прибора.
        /// </summary>
        public Guid DeviceID;
        /// <summary>
        /// Идентификатор комплектующего.
        /// </summary>
        public Guid CompleteID;
        /// <summary>
        /// Идентификатор кода СКБ.
        /// </summary>
        public Guid CodeSKB;
        /// <summary>
        /// Количество.
        /// </summary>
        public int Count;
        /// <summary>
        /// Название документа-основания.
        /// </summary>
        public string DocumentName;
        /// <summary>
        /// Конструктор зарезервированного комплектующего.
        /// </summary>
        /// <param name="deviceID">Идентификатор прибора.</param>
        /// <param name="completeID">Идентификатор комплектующего.</param>
        /// <param name="codeSKB">Идентификатор кода СКБ.</param>
        /// <param name="count">Количество.</param>
        /// <param name="documentName">Название документа-основания.</param>
        public ReservedCompleteRow(Guid deviceID, Guid completeID, Guid codeSKB, int count, string documentName)
        {
            DeviceID = deviceID;
            CompleteID = completeID;
            CodeSKB = codeSKB;
            Count = count;
            DocumentName = documentName;
        }
    }
    /// <summary>
    /// Информация о текущем остатке комплектующего.
    /// </summary>
    public class CurrentBalanceComplete
    {
        /// <summary>
        /// Идентификатор прибора.
        /// </summary>
        public Guid DeviceID;
        /// <summary>
        /// Идентификатор комплектующего.
        /// </summary>
        public Guid CompleteID;
        /// <summary>
        /// Идентификатор кода СКБ.
        /// </summary>
        public Guid CodeSKB;
        /// <summary>
        /// Остаток на дату последнего подсчета остатков.
        /// </summary>
        public int StartBalance;
        /// <summary>
        /// Остаток на текущий момент.
        /// </summary>
        public int EndBalance;
        /// <summary>
        /// Поступление с момента последнего подсчета остатков.
        /// </summary>
        public int Received;
        /// <summary>
        /// Расход с момента последнего подсчета остатков.
        /// </summary>
        public int Descended;
        /// <summary>
        /// Зарезервировано комплектующего на текущий момент.
        /// </summary>
        public int Reserved;
        /// <summary>
        /// Документы основания для поступления комплектующих с момента последнего подсчета остатков.
        /// </summary>
        public List<string> ReceivedDocuments;
        /// <summary>
        /// Документы основания для расхода комплектующих с момента последнего подсчета остатков.
        /// </summary>
        public List<string> DescendedDocuments;
        /// <summary>
        /// Документы основания для зарезервированных комплектующих.
        /// </summary>
        public List<string> ReservedDocuments;
        /// <summary>
        /// Конструктор информации о текущем остатке комплектующего
        /// </summary>
        /// <param name="deviceID">Идентификатор прибора.</param>
        /// <param name="completeID">Идентификатор комплектующего.</param>
        /// <param name="codeSKB">Идентификатор кода СКБ.</param>
        /// <param name="startBalance">Остаток на дату последнего подсчета остатков.</param>
        /// <param name="received">Расход с момента последнего подсчета остатков.</param>
        /// <param name="descended">Расход с момента последнего подсчета остатков.</param>
        /// <param name="endBalance"> Остаток на текущий момент.</param>
        public CurrentBalanceComplete(Guid deviceID, Guid completeID, Guid codeSKB, int startBalance, int received, int descended, int endBalance)
        {
            DeviceID = deviceID;
            CompleteID = completeID;
            CodeSKB = codeSKB;
            StartBalance = startBalance;
            EndBalance = endBalance;
            Received = received;
            Descended = descended;
            Reserved = 0;
            ReceivedDocuments = new List<string>();
            DescendedDocuments = new List<string>();
            ReservedDocuments = new List<string>();
        }
    }
}
