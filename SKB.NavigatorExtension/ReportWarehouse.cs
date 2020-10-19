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
using DocsVision.BackOffice.ObjectModel.Services;
using DocsVision.BackOffice.ObjectModel;
using System.Diagnostics;
using DocsVision.Platform.ObjectManager.SearchModel;
using DocsVision.TakeOffice.Cards.Constants;
using DocsVision.Platform.ObjectManager.Metadata;
using System.IO;



namespace SKB.NavigatorExtension
{
    /// <summary>
    /// Столбцы отчета.
    /// </summary>
    public enum Columns
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

    class ReportWarehouse
    {
        #region Properties
        /// <summary>
        /// Название файла отчета.
        /// </summary>
        public string FileName
        { get { return "Отчет по складу готовой продукции с " + StartDate.ToShortDateString() + " по " + EndDate.ToShortDateString() + ".xlsx"; } }
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
        /// Конструкток отчета по складу готовой продукции.
        /// </summary>
        /// <param name="session">Пользовательская сессия DV.</param>
        /// <param name="StartDate">Дата начала периода.</param>
        /// <param name="EndDate">Дата окончания периода.</param>
        public ReportWarehouse(UserSession session, DateTime StartDate, DateTime EndDate)
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
    /// <summary>
    /// Запись отчета.
    /// </summary>
    public class ReportWareHouseItem
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
        public ReportWarehouseValue[] Values;
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
        public ReportWareHouseItem(int[] ColumnsIndex, string DeviseType, string DeviceParty)
        {

            this.deviseType = DeviseType;
            this.deviceParty = DeviceParty;
            this.Values = new ReportWarehouseValue[ColumnsIndex.Length];
            for (int i = 0; i < ColumnsIndex.Length; i++)
                this.Values[i] = new ReportWarehouseValue(ColumnsIndex[i], 0, "");
        }
    }
    /// <summary>
    /// Значение записи отчета.
    /// </summary>
    public class ReportWarehouseValue
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
        public ReportWarehouseValue(int ColumnIndex, int Value, string Description)
        {
            this.columnIndex = ColumnIndex;
            this.Value = Value;
            this.Description = Description;
        }
    }
    /// <summary>
    /// Операция передачи прибора.
    /// </summary>
    public class TransferRow : IComparable<TransferRow>
    {
        #region Properties
        /// <summary>
        /// Тип передачи прибора.
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
        /// Действие с прибором.
        /// </summary>
        public enum Action
        {
            /// <summary>
            /// Приход из производства (новые приборы).
            /// </summary>
            [Description("Приход из производства (новые приборы)")]
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
            /// Отгрузка новых приборов.
            /// </summary>
            [Description("Отгрузка новых приборов")]
            DeliveryNewDevices = 11,
            /// <summary>
            /// Нет действия.
            /// </summary>
            [Description("Нет действия")]
            None = 12,
        };
        /// <summary>
        /// Количество приборов.
        /// </summary>
        private int deviceCount;
        /// <summary>
        /// Идентификатор заводского номера приборов.
        /// </summary>
        private Guid deviceNumberID;
        /// <summary>
        /// Тип передачи.
        /// </summary>
        private TransferTypes transferType;
        /// <summary>
        /// Действие с прибором.
        /// </summary>
        private Action transferAction;
        /// <summary>
        /// Дата передачи.
        /// </summary>
        private DateTime transferDate;
        /// <summary>
        /// Тип прибора.
        /// </summary>
        private string deviceType = "";
        /// <summary>
        /// Заводской номер прибора.
        /// </summary>
        private string deviceNumber = "";
        /// <summary>
        /// Партия прибора.
        /// </summary>
        private string deviceParty = "";
        /// <summary>
        /// Универсальный спрвочник.
        /// </summary>
        private CardData universalDictionary;
        #endregion

        #region Fields
        /// <summary>
        /// Количество приборов.
        /// </summary>
        public int DeviceCount { get { return deviceCount; } }
        /// <summary>
        /// Идентификатор заводского номера прибора.
        /// </summary>
        public Guid DeviceNumberID { get { return deviceNumberID; } }
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
        /// Тип прибора.
        /// </summary>
        public string DeviceType
        {
            get
            {
                if (string.IsNullOrEmpty(this.deviceType))
                    this.deviceType = this.universalDictionary.GetItemName(new Guid(this.universalDictionary.GetItemPropertyValue(this.deviceNumberID, "Наименование прибора").ToString()));
                return this.deviceType;
            }
        }
        /// <summary>
        /// Заводской номер прибора
        /// </summary>
        public string DeviceNumber
        {
            get
            {
                if (string.IsNullOrEmpty(this.deviceNumber))
                {
                    string DeviceNumber = this.universalDictionary.GetItemPropertyDisplayValue(this.deviceNumberID, "Номер прибора");
                    string DeviceYear = this.universalDictionary.GetItemPropertyDisplayValue(this.deviceNumberID, "Год прибора");
                    this.deviceNumber = DeviceNumber.Length >= 4 ? DeviceNumber : DeviceNumber + "/" + DeviceYear;
                }
                return this.deviceNumber;
            }
        }
        /// <summary>
        /// Партия пибора.
        /// </summary>
        public string DeviceParty
        {
            get
            {
                if (string.IsNullOrEmpty(this.deviceParty))
                    this.deviceParty = universalDictionary.GetItemName(new Guid(universalDictionary.GetItemPropertyValue(this.deviceNumberID, "Партия").ToString()));
                return this.deviceParty;
            }
        }
        #endregion

        /// <summary>
        /// Конструктор операции передачи прибора.
        /// </summary>
        /// <param name="UniversalDictionary">Карточка универсального справочника.</param>
        /// <param name="DeviceNumberID">Идентификатор заводского номера прибора.</param>
        /// <param name="TransferType">Тип передачи.</param>
        /// <param name="TransferAction">Действие с прибором.</param>
        /// <param name="TransferDate">Дата передачи.</param>
        /// <param name="Count">Количество приборов.</param>
        public TransferRow(CardData UniversalDictionary, Guid DeviceNumberID, TransferTypes TransferType, Action TransferAction, DateTime TransferDate, int Count = 1)
        {
            if (TransferAction == Action.ReturnFromTestDrive)
                this.transferAction = TransferAction;
            this.universalDictionary = UniversalDictionary;
            this.deviceNumberID = DeviceNumberID;
            this.transferType = TransferType;
            this.transferAction = TransferAction;
            this.transferDate = TransferDate;
            this.deviceCount = Count;
        }
        /// <summary>
        /// Сравнение операций передачи по дате.
        /// </summary>
        /// <param name="other">Операция для сравнения.</param>
        /// <returns></returns>
        int IComparable<TransferRow>.CompareTo(TransferRow other)
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
    /// Операция передачи приборов в партии.
    /// </summary>
    public class TransferCountByParty : IComparable<TransferCountByParty>
    {
        #region Properties
        /// <summary>
        /// Количество приборов.
        /// </summary>
        private int deviceCount;
        /// <summary>
        /// Перечень заводских номеров приборов.
        /// </summary>
        private string deviceNumbersCollection;
        /// <summary>
        /// Тип приборов.
        /// </summary>
        private string deviceType;
        /// <summary>
        /// Название партии.
        /// </summary>
        private string deviceParty;
        /// <summary>
        /// Год партии.
        /// </summary>
        private int partyYear;
        /// <summary>
        /// Месяц партии.
        /// </summary>
        private int partyMonth;
        #endregion
        #region Fields
        /// <summary>
        /// Количество приборов.
        /// </summary>
        public int DeviceCount { get { return deviceCount; } }
        /// <summary>
        /// Перечень заводских номеров приборов.
        /// </summary>
        public string DeviceNumbersCollection { get { return deviceNumbersCollection; } }
        /// <summary>
        /// Тип приборов.
        /// </summary>
        public string DeviceType { get { return deviceType; } }
        /// <summary>
        /// Название партии
        /// </summary>
        public string DeviceParty { get { return deviceParty; } }
        /// <summary>
        /// Год партии.
        /// </summary>
        public int PartyYear { get { return partyYear; } }
        /// <summary>
        /// Месяц партии.
        /// </summary>
        public int PartyMonth { get { return partyMonth; } }
        #endregion

        /// <summary>
        /// Конструктор операции передачи приборов в партии.
        /// </summary>
        /// <param name="DeviceParty">Название партии.</param>
        /// <param name="DeviceType">Тип приборов.</param>
        /// <param name="DeviceNumbersCollection">Перечнь заводских номеров приборов.</param>
        /// <param name="DeviceCount">Количество приборов.</param>
        public TransferCountByParty(string DeviceParty, string DeviceType, string DeviceNumbersCollection, int DeviceCount)
        {
            this.deviceParty = DeviceParty;
            this.deviceType = DeviceType;
            this.deviceNumbersCollection = DeviceNumbersCollection;
            this.deviceCount = DeviceCount;

            string[] PartyOptions = DeviceParty.Split(new string[] { " - " }, StringSplitOptions.RemoveEmptyEntries);
            this.partyYear = Convert.ToInt32(PartyOptions[2].Split('/')[1]);
            this.partyMonth = Convert.ToInt32(PartyOptions[2].Split('/')[0]);
        }
        /// <summary>
        /// Сравнение операций передачи приборов в партии по месяцу и году партии
        /// </summary>
        /// <param name="other"></param>
        /// <returns></returns>
        int IComparable<TransferCountByParty>.CompareTo(TransferCountByParty other)
        {
            int PartyYearOther = other.PartyYear;
            int PartyMonthOther = other.PartyMonth;

            int PartyYearThis = this.PartyYear;
            int PartyMonthThis = this.PartyMonth;

            if (PartyYearOther > PartyYearThis)
                return -1;
            else
            {
                if (PartyYearOther == PartyYearThis)
                {
                    if (PartyMonthOther > PartyMonthThis)
                        return -1;
                    else if (PartyMonthOther == PartyMonthThis)
                        return 0;
                    else
                        return 1;
                }
                else
                    return 1;
            }
        }
    }

    /// <summary>
    /// Сравнение партий по месяцу и году выпуска
    /// </summary>
    public class PartyComparer : IComparer<string>
    {
        /// <summary>
        /// Компаратор
        /// </summary>
        /// <param name="party1">Партия 1</param>
        /// <param name="party2">Партия 2</param>
        /// <returns></returns>
        public int Compare(string party1, string party2)
        {
            string[] PartyOptions1 = party1.Split(new string[] { " - " }, StringSplitOptions.RemoveEmptyEntries);
            string[] PartyOptions2 = party2.Split(new string[] { " - " }, StringSplitOptions.RemoveEmptyEntries);

            int PartyYear1 = Convert.ToInt32(PartyOptions1[2].Split('/')[1]); ;
            int PartyMonth1 = Convert.ToInt32(PartyOptions1[2].Split('/')[0]); ;

            int PartyYear2 = Convert.ToInt32(PartyOptions2[2].Split('/')[1]); ;
            int PartyMonth2 = Convert.ToInt32(PartyOptions2[2].Split('/')[0]); ;

            if (PartyYear1 > PartyYear2)
                return -1;
            else
            {
                if (PartyYear1 == PartyYear2)
                {
                    if (PartyMonth1 > PartyMonth2)
                        return -1;
                    else if (PartyMonth1 == PartyMonth2)
                        return 0;
                    else
                        return 1;
                }
                else
                    return 1;
            }
        }
    }

    /// <summary>
    /// Сравнение партий по месяцу и году выпуска
    /// </summary>
    public static class ReportHelper
    {
        // Дополнительные изделия:
        /// <summary>
        /// Перечень доп. изделий, исключенных из отчета
        /// </summary>
        public static string[] AdditionalWaresList = new string[] { "ДП12", "ДП21", "ДП22", "ТК-021", "ДП32.1", "ДП32.2", "ТК-026", "ПКП-08-00", "ПКП-08-01", "ПКП-08-02", "ПКП-25-00", "ПКП-25-01", "ПКП-25-02" };
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
        /// Поиск Актов передачи приборов на склад.
        /// </summary>
        /// <param name="session">Пользовательская сессия DV.</param>
        /// <param name="StartDate">Начальная дата.</param>
        /// <param name="EndDate">Конечная дата.</param>
        /// <param name="UniversalDictionary">Карточка универсального справочника.</param>
        /// <returns></returns>
        public static IEnumerable<TransferRow> FindAct(UserSession session, DateTime StartDate, DateTime EndDate, CardData UniversalDictionary)
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
            sectionQuery = typeQuery.SectionQueries.AddNew(CardOrd.Properties.ID);
            sectionQuery.Operation = SectionQueryOperation.And;
            sectionQuery.ConditionGroup.Operation = ConditionGroupOperation.And;
            sectionQuery.ConditionGroup.Conditions.AddNew(CardOrd.Properties.Name, FieldType.Unistring, ConditionOperation.Equals, "Режим передачи");
            sectionQuery.ConditionGroup.Conditions.AddNew(CardOrd.Properties.Value, FieldType.Int, ConditionOperation.NotEquals, 2);

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
            sectionQuery.ConditionGroup.Conditions.AddNew(CardOrd.Properties.Value, FieldType.Date, ConditionOperation.GreaterThan, new DateTime(2013, 1, 1));

            // Получение текста запроса
            searchQuery.Limit = 0;
            string query = searchQuery.GetXml();
            CardDataCollection ActCollection = session.CardManager.FindCards(query);

            int i = 0;
            IEnumerable<TransferRow> DeviceNumberIDs = ActCollection.SelectMany(r => r.ActToTransferRowCollection(UniversalDictionary, ref i));
            Clear();
            return DeviceNumberIDs;
        }
        /// <summary>
        /// Преобразование Акта передачи в коллекцию операций передачи приборов.
        /// </summary>
        /// <param name="Card">Карточка Акта.</param>
        /// <param name="UniversalDictionary">Карточка универсального справочника.</param>
        /// <param name="i">Количество операций.</param>
        /// <returns></returns>
        private static IEnumerable<TransferRow> ActToTransferRowCollection(this CardData Card, CardData UniversalDictionary, ref int i)
        {
            TransferRow.TransferTypes TransferType;
            TransferRow.Action TransferAction;

            switch (Card.Sections[CardOrd.Properties.ID].FindRow("@Name = 'Вид передачи'").GetString(CardOrd.Properties.Value))
            {
                case "Калибровка -> Сбыт":
                    TransferType = TransferRow.TransferTypes.ToWarehouse;
                    if ((int)Card.Sections[CardOrd.Properties.ID].FindRow("@Name = 'Режим передачи'").GetInt32(CardOrd.Properties.Value) == 1)
                        TransferAction = TransferRow.Action.NewReceiptFromProduction;
                    else
                        TransferAction = TransferRow.Action.RepeatReceiptFromProduction;
                    break;
                case "Сбыт -> Калибровка":
                    TransferType = TransferRow.TransferTypes.FromWarehouse;
                    switch ((int)Card.Sections[CardOrd.Properties.ID].FindRow("@Name = 'Цель передачи'").GetInt32(CardOrd.Properties.Value))
                    {
                        case 2:
                            TransferAction = TransferRow.Action.DeliveryToTesting;
                            break;
                        case 3:
                            TransferAction = TransferRow.Action.ReturnFromWarehouseToProduction;
                            break;
                        default:
                            TransferAction = TransferRow.Action.ReturnFromWarehouseToProduction;
                            break;
                    }
                    break;
                default:
                    TransferType = TransferRow.TransferTypes.Internal;
                    TransferAction = TransferRow.Action.None;
                    break;
            }

            IEnumerable<RowData> VerifyRows = Card.Sections[CardOrd.Properties.ID].FindRow("@Name = 'Заводской номер'").ChildSections[CardOrd.SelectedValues.ID].Rows.Where(r =>
                !AdditionalWaresList.Any(rr => UniversalDictionary.GetItemName((Guid)r.GetGuid(CardOrd.SelectedValues.SelectedValue)).StartsWith(rr)));
            IEnumerable<TransferRow> Result = VerifyRows.Select(r =>
                    new TransferRow(
                        UniversalDictionary,
                        (Guid)r.GetGuid(CardOrd.SelectedValues.SelectedValue),
                        TransferType, TransferAction,
                        (DateTime)Card.Sections[CardOrd.Properties.ID].FindRow("@Name = 'Дата принятия'").GetDateTime(CardOrd.Properties.Value)));
            return Result;
        }
        /// <summary>
        /// Поиск Заданий на комплектацию.
        /// </summary>
        /// <param name="session">Пользовательская сессия DV.</param>
        /// <param name="StartDate">Начальная дата.</param>
        /// <param name="EndDate">Конечная дата.</param>
        /// <param name="UniversalDictionary">Карточка универсального справочника.</param>
        /// <returns></returns>
        public static IEnumerable<TransferRow> FindCompleteTasks(UserSession session, DateTime StartDate, DateTime EndDate, CardData UniversalDictionary)
        {
            /* Поиск Заявок на комплектацию */
            SearchQuery searchQuery = session.CreateSearchQuery();
            searchQuery.CombineResults = ConditionGroupOperation.And;

            CardTypeQuery typeQuery = searchQuery.AttributiveSearch.CardTypeQueries.AddNew(RefCompleteCard.ID);

            // Фактическая дата отгрузки меньше конечной даты
            SectionQuery sectionQuery = typeQuery.SectionQueries.AddNew(RefCompleteCard.Devices.ID);
            sectionQuery.Operation = SectionQueryOperation.And;
            sectionQuery.ConditionGroup.Operation = ConditionGroupOperation.And;
            sectionQuery.ConditionGroup.Conditions.AddNew(RefCompleteCard.Devices.FactShipDate, FieldType.DateTime, ConditionOperation.LessEqual, EndDate);
            sectionQuery.ConditionGroup.Conditions.AddNew(RefCompleteCard.Devices.AC, FieldType.Bool, ConditionOperation.Equals, false);

            // Получение текста запроса
            searchQuery.Limit = 0;
            string query = searchQuery.GetXml();

            CardDataCollection CompleteTaskCollection = session.CardManager.FindCards(query);
            IEnumerable<TransferRow> DeviceNumberIDs = CompleteTaskCollection.SelectMany(r => r.CompleteToTransferRowCollection(UniversalDictionary));

            return DeviceNumberIDs;
        }
        /// <summary>
        /// Преобразование Задания на комплектацию в коллекцию операций передачи приборов.
        /// </summary>
        /// <param name="Card">Карточка Задания на комплектацию.</param>
        /// <param name="UniversalDictionary">Карточка универсального справочника.</param>
        /// <returns></returns>
        private static IEnumerable<TransferRow> CompleteToTransferRowCollection(this CardData Card, CardData UniversalDictionary)
        {
            IEnumerable<TransferRow> Result = Card.Sections[new Guid("{27DB67DE-5FA9-4BF1-BC7B-9F5FF2B972E2}")].Rows.SelectMany(r => ConvertCompleteToTransferRowsCollection(r, UniversalDictionary));
            return Result;
        }
        /// <summary>
        /// Преобразование строки таблицы "Информация по приборам" в коллекцию операций передачи приборов.
        /// </summary>
        /// <param name="CompleteRow">Строка таблицы "Информация по приборам."</param>
        /// <param name="UniversalDictionary">Карточка универсального справочника.</param>
        /// <returns></returns>
        private static IEnumerable<TransferRow> ConvertCompleteToTransferRowsCollection(RowData CompleteRow, CardData UniversalDictionary)
        {
            List<TransferRow> TransferRowsCollection = new List<TransferRow>();
            if ((bool)CompleteRow.GetBoolean(RefCompleteCard.Devices.AC) || AdditionalWaresList.Any(r => r == UniversalDictionary.GetItemName((Guid)CompleteRow.GetGuid(RefCompleteCard.Devices.DeviceId))))
                return TransferRowsCollection.AsEnumerable();

            string ContractSubjectId = CompleteRow.GetString(RefCompleteCard.Devices.ContractSubjectId);
            object DeviceNumberId = CompleteRow.GetGuid(RefCompleteCard.Devices.DeviceNumberId);
            object FactShipDate = CompleteRow.GetObject(RefCompleteCard.Devices.FactShipDate);
            object FactReturnDate = CompleteRow.GetObject(RefCompleteCard.Devices.FactReturnDate);
            object PaymentDate = CompleteRow.GetObject(RefCompleteCard.Devices.PaymentDate);

            switch (ContractSubjectId)
            {
                case DeliveryID:
                    if (!FactShipDate.IsNull()) TransferRowsCollection.Add(new TransferRow(UniversalDictionary, (Guid)DeviceNumberId, TransferRow.TransferTypes.FromWarehouse, TransferRow.Action.DeliveryNewDevices, (DateTime)FactShipDate));
                    break;
                case ExpositionID:
                    if (!FactShipDate.IsNull() && !DeviceNumberId.IsNull() && !((Guid)DeviceNumberId).IsEmpty()) TransferRowsCollection.Add(new TransferRow(UniversalDictionary, (Guid)DeviceNumberId, TransferRow.TransferTypes.FromWarehouse, TransferRow.Action.DeliveryToExposition, (DateTime)FactShipDate));
                    if (!FactReturnDate.IsNull() && !DeviceNumberId.IsNull() && !((Guid)DeviceNumberId).IsEmpty()) TransferRowsCollection.Add(new TransferRow(UniversalDictionary, (Guid)DeviceNumberId, TransferRow.TransferTypes.ToWarehouse, TransferRow.Action.ReturnFromExposition, ((DateTime)FactReturnDate).AddHours(1)));
                    break;
                case TestDriveID:
                    if (!FactShipDate.IsNull())
                        TransferRowsCollection.Add(new TransferRow(UniversalDictionary, (Guid)DeviceNumberId, TransferRow.TransferTypes.FromWarehouse, TransferRow.Action.DeliveryToTestDrive, (DateTime)FactShipDate));
                    if (!FactReturnDate.IsNull())
                        TransferRowsCollection.Add(new TransferRow(UniversalDictionary, (Guid)DeviceNumberId, TransferRow.TransferTypes.ToWarehouse, TransferRow.Action.ReturnFromTestDrive, ((DateTime)FactReturnDate).AddHours(1)));
                    if (!PaymentDate.IsNull())
                    {
                        TransferRowsCollection.Add(new TransferRow(UniversalDictionary, (Guid)DeviceNumberId, TransferRow.TransferTypes.ToWarehouse, TransferRow.Action.ReturnFromTestDrive, (DateTime)PaymentDate));
                        TransferRowsCollection.Add(new TransferRow(UniversalDictionary, (Guid)DeviceNumberId, TransferRow.TransferTypes.FromWarehouse, TransferRow.Action.DeliveryNewDevices, ((DateTime)PaymentDate).AddHours(1)));
                    }
                    
                    break;
                case CertificationID:
                    if (!FactShipDate.IsNull() && !DeviceNumberId.IsNull() && !((Guid)DeviceNumberId).IsEmpty()) TransferRowsCollection.Add(new TransferRow(UniversalDictionary, (Guid)DeviceNumberId, TransferRow.TransferTypes.FromWarehouse, TransferRow.Action.DeliveryToCertification, (DateTime)FactShipDate));
                    if (!FactReturnDate.IsNull() && !DeviceNumberId.IsNull() && !((Guid)DeviceNumberId).IsEmpty()) TransferRowsCollection.Add(new TransferRow(UniversalDictionary, (Guid)DeviceNumberId, TransferRow.TransferTypes.ToWarehouse, TransferRow.Action.ReturnFromCertification, ((DateTime)FactReturnDate).AddHours(1)));
                    break;
                case TenderID:
                    if (!FactShipDate.IsNull()) TransferRowsCollection.Add(new TransferRow(UniversalDictionary, (Guid)DeviceNumberId, TransferRow.TransferTypes.FromWarehouse, TransferRow.Action.DeliveryNewDevices, (DateTime)FactShipDate));
                    break;
                case ActionID:
                    if (!FactShipDate.IsNull()) TransferRowsCollection.Add(new TransferRow(UniversalDictionary, (Guid)DeviceNumberId, TransferRow.TransferTypes.FromWarehouse, TransferRow.Action.DeliveryNewDevices, (DateTime)FactShipDate));
                    break;
            }
            return TransferRowsCollection.AsEnumerable();
        }
        /// <summary>
        /// Получение статистики на конкретную дату
        /// </summary>
        /// <param name="AllTransfers">Перечень всех операций передач приборов.</param>
        /// <param name="UniversalDictionary">Карточка универсального справочника.</param>
        /// <param name="KeyDate">Дата.</param>
        /// <param name="TransferType">Тип передач.</param>
        /// <param name="Action">Действие с прибором.</param>
        /// <returns></returns>
        public static IEnumerable<TransferCountByParty> StatisticsOnDate(this IEnumerable<TransferRow> AllTransfers, CardData UniversalDictionary, DateTime KeyDate,
            TransferRow.TransferTypes TransferType, TransferRow.Action Action = TransferRow.Action.None)
        {
            // Перечень последних операций с приборами на указанную дату
            IEnumerable<TransferRow> AllActionsOnDate = AllTransfers.Where(r => r.TransferDate < KeyDate).GroupBy(r => r.DeviceNumberID, (number, rows) =>
                        new TransferRow(UniversalDictionary, number, rows.Max().TransferType, rows.Max().TransferAction, rows.Max().TransferDate)).Where(r =>
                            r.TransferType == TransferType && (Action == TransferRow.Action.None || r.TransferAction == Action));
            // Группировка операций по партиям
            IEnumerable<TransferCountByParty> AllActionsOnDateByParty = AllActionsOnDate.GroupBy(r => r.DeviceParty, (deviceParty, transfers) =>
                new TransferCountByParty(
                    transfers.First().DeviceParty,
                    transfers.First().DeviceType,
                    transfers.Select<TransferRow, string>(r => r.DeviceNumber).Aggregate(String.Empty, (start, next) => ReportWarehouse.JoinDevices(", ", start, next), result => result),
                    transfers.Count()));
            return AllActionsOnDateByParty;
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
        public static IEnumerable<TransferCountByParty> StatisticsForPeriod(this IEnumerable<TransferRow> AllTransfers, DateTime StartDate, DateTime EndDate, TransferRow.Action Action, CardData UniversalDictionary)
        {
            // Перечень всех операций с заданным действием за указанный период
            IEnumerable<TransferRow> AllActionsForPeriod = AllTransfers.Where(r => r.TransferAction == Action && r.TransferDate >= StartDate && r.TransferDate < EndDate.AddDays(1));

            // Группировка операций по партиям
            IEnumerable<TransferCountByParty> AllActionsForPeriodByParty = AllActionsForPeriod.GroupBy(r => r.DeviceParty, (deviceParty, transfers) =>
                new TransferCountByParty(
                    transfers.First().DeviceParty,
                    transfers.First().DeviceType,
                    transfers.Select<TransferRow, string>(r => r.DeviceNumber).Aggregate(String.Empty, (start, next) => ReportWarehouse.JoinDevices(", ", start, next), result => result),
                    transfers.Count()));

            return AllActionsForPeriodByParty;
        }
        /// <summary>
        /// Открыть файл отчета.
        /// </summary>
        public static void OpenReport(String Path)
        {
            if (File.Exists(Path))
            { System.Diagnostics.Process.Start(Path); }
        }
    }
}
