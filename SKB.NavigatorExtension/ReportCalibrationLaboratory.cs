using System;
using System.Collections;
using System.Diagnostics;
using System.Windows.Forms;
using System.IO;

using DocsVision.Platform.ObjectManager;
using DocsVision.Platform.ObjectManager.SearchModel;
using DocsVision.Platform.ObjectManager.Metadata;
using DocsVision.TakeOffice.Cards.Constants;
using SKB.Base;
using SKB.Base.Ref;
using Excel = Microsoft.Office.Interop.Excel;
using RKIT.MyMessageBox;

namespace SKB.NavigatorExtension
{
    /// <summary>
    /// Отчет по работе калибровочной лаборатории.
    /// </summary>
    public class ReportCalibrationLaboratory
    {
        /// <summary>
        /// Идентификатор типа каточки "Акт передачи приборов и комплектующих".
        /// </summary>
        Guid ActTypeID = new Guid("{1284381C-B70D-4D9C-9258-718414F1BDB3}");   // Тип - Акт передачи
        /// <summary>
        /// Идентификатор типа каточки "Cервисное обслуживание прибора".
        /// </summary>
        Guid ServiceCardTypeID = new Guid("{C3548925-A7ED-4BD7-BCBB-ED7066D7C81A}");   // Тип - Сервисное обслуживание прибора

        // Постоянные параметры отчета:
        /// <summary>
        /// Пароль для разблокировки отчета.
        /// </summary>
        const string Password = "Cuckoo#3";
        /// <summary>
        /// Стартовая позиция в отчете.
        /// </summary>
        const int StartPosition = 8;
        /// <summary>
        /// Количество столбцов в отчете.
        /// </summary>
        const int ColumnsCount = 15;
        /// <summary>
        /// Позиция заголовка в отчете.
        /// </summary>
        const int HeaderPosition = 5;
        /// <summary>
        /// Текст заголовка отчета.
        /// </summary>
        string HeaderText
        { get { return "План-факт работ калибровочной лаборатории за период" + " c " + StartDate.ToString("d.MM.yy") + " по " + EndDate.ToString("d.MM.yy") + ".xlsx"; } }
        /// <summary>
        /// Путь к временной папке пользователя.
        /// </summary>
        string TempPath = Path.GetTempPath();

        // Переменные параметры отчета:
        /// <summary>
        /// Начальная дата отчетного периода
        /// </summary>
        DateTime StartDate;
        /// <summary>
        /// Конечная дата отчетного периода.
        /// </summary>
        DateTime EndDate;
        /// <summary>
        /// Сессия DV.
        /// </summary>
        UserSession session;
        /// <summary>
        /// Количество строк в отчете
        /// </summary>
        int RowsCount = 6;

        // Объекты для работы с Excel:
        /// <summary>
        /// Приложение Excel
        /// </summary>
        Excel.Application objExcel = null;
        /// <summary>
        /// Рабочая книга Excel
        /// </summary>
        Excel.Workbook objWorkbook = null;
        /// <summary>
        /// Рабочий лист Excel
        /// </summary>
        Excel.Worksheet objWorksheet = null;

        /// <summary>
        /// Отчет по работе калибровочной лаборатории.
        /// </summary>
        public ReportCalibrationLaboratory(UserSession session, DateTime StartDate, DateTime EndDate)
        {
            this.session = session;
            this.StartDate = StartDate;
            this.EndDate = EndDate;

            // Выбор файла, содержащего план работ калибровочной лаборатории на неделю
            OpenFileDialog NewFileDialog = new OpenFileDialog();
            NewFileDialog.Multiselect = false;
            NewFileDialog.Title = "Выберите план работ калибровочной лаборатории на текущий период:";
            DialogResult Result = NewFileDialog.ShowDialog();
            if (Result == DialogResult.OK)
            {
                if (!NewFileDialog.FileName.EndsWith(".xlsx"))
                {
                    MyMessageBox.Show("Ошибка! Выбран неверный файл. Для формирования отчета \"План-факт работ калибровочной лаборатории\" требуется выбрать файл Excel, содержащий план работ калибровочной лаборатории на соответствующий период.");
                    return;
                }
                // Копирование выбранного файла во временную папку
                File.Copy(NewFileDialog.FileName, TempPath + NewFileDialog.SafeFileName, true);
                // Удаление старого отчета, если он существует
                if (File.Exists(TempPath + HeaderText)) File.Delete(TempPath + HeaderText);
                // Создание объектов для работы с Excel
                objExcel = new Excel.Application();
                objWorkbook = objExcel.Workbooks.Add(TempPath + NewFileDialog.SafeFileName);
                objWorksheet = objWorkbook.Worksheets[1];
            }
        }
        /// <summary>
        /// Формирование отчета "План-факт по работе калибровочной лаборатории за период"
        /// </summary>
        public void Create()
        {
            // Отмена формирования отчета, если не найден рабочий лист
            if (objWorksheet == null) return;
            // Добавление в исходный отчет столбцов "Факт"
            AddFactColumns(objWorksheet);
            RowsCount--;
            // Получение данных о фактической работе калибровочной лаборатории за указанный период
            ArrayList InitiatingDocuments = GetFactCalibrationData();
            // Заполнение отчета на основе данных о фактической работе калибровочной лаборатории
            foreach (ReportItem Item in InitiatingDocuments)
            {
                // Заполнение информации о новых приборах, переданных на склад готовой продукции
                if (Item.StatusOfTransfer == StatusOfTransferValues.NewDevices)
                { int NewPartyPosition = FindParty(objWorksheet, Item.Party, Item.DeviceType, Item.PlanCalibrationTime); }
                // Заполнение информации о новых комплектующих, переданных на склад готовой продукции
                if (Item.StatusOfTransfer == StatusOfTransferValues.NewComplete)
                { int NewCompletePosition = FindComplete(objWorksheet, Item.CompleteName, Item.CompleteDocument, Item.FactCalibrationTime, Item.Count); }
                // Заполнение информации о приборах и комплектующих, переданных на склад готовой продукции после повторной калибровки
                if (Item.StatusOfTransfer == StatusOfTransferValues.AfterRecalibration)
                {
                    string Name = Item.CompleteName == "" ? Item.DeviceName : Item.CompleteName;
                    int NewCompletePosition = FindDeviceOrCompleteAfterRecalibration(objWorksheet, Name, Item.CompleteDocument, Item.FactCalibrationTime, Item.Count, Item.PlanCalibrationTime);
                }
                // Заполнение информации о приборах/комплектующих, принадлежащих заказчикам, переданных в ремонт или на склад готовой продукции
                if (Item.StatusOfTransfer == StatusOfTransferValues.ClientsDevices)
                { int NewCompletePosition = FindDevice(objWorksheet, Item.DeviceName, Item.ClientName, Item.PlanCalibrationTime, Item.FactCalibrationTime, Item.DeviceState, Item.ServiceType, Item.PlanEndServiceDate, Item.Warranty, Item.TypeOfTransfer, Item.ServiceNumber, Item.ServiceComment); }
            }
            // Обновление итоговых формул отчета
            TotalCount(objWorksheet);
            // Проставление цветовых меток в отчете
            SetColorMarks(objWorksheet);
            // Создание заголовка
            SetHeader(objWorksheet);
            // Создание легенды
            SetLegend(objWorksheet);
            // Блокировка отчета
            SetAlignment(objWorksheet.Columns[ColumnsCount], Excel.XlHAlign.xlHAlignGeneral, true);
            objWorksheet.Protection.AllowEditRanges.Add("Примечания", objWorksheet.Columns[ColumnsCount]);
            //objWorksheet.EnableSelection = Excel.XlEnableSelection.xlUnlockedCells;
            foreach (Excel.Worksheet MyWorksheet in objWorkbook.Worksheets)
            {MyWorksheet.Protect(Password, true, true, true);}
            objWorkbook.Protect(Password, true);
            // Сохранение изменений
            Object wMissing = System.Reflection.Missing.Value;
            objWorksheet.SaveAs(TempPath + HeaderText, wMissing, wMissing, wMissing, wMissing, wMissing, wMissing, wMissing, wMissing, wMissing);
            objWorkbook.Save();
            objExcel.Quit();
        }
        /// <summary>
        /// Открытие файла отчета "План-факт по работе калибровочной лаборатории за период"
        /// </summary>
        public void OpenReport()
        {
            if (objWorksheet == null) return;

            if (File.Exists(TempPath + HeaderText))
            {System.Diagnostics.Process.Start(TempPath + HeaderText);}
            else
            {MyMessageBox.Show("Ошибка! Не удалось сформировать отчет.");}
        }
        /// <summary>
        /// Получение данных о фактической работе калибровочной лаборатории за указанный период
        /// </summary>
        private ArrayList GetFactCalibrationData()
        {
            ArrayList InitiatingDocuments = new ArrayList();
            CardData UniversalDictionary = session.CardManager.GetDictionaryData(RefUniversal.ID);
            // Поиск Актов передачи приборов и комплектующих
            CardDataCollection Acts = FindAct();
            Clear();
            // Обработка найденных Актов передачи приборов и комплектующих
            CardData CurrentAct;
            SectionData ActPropertiesSection;
            for (int i = 0; i < Acts.Count; i++)
            {
                CurrentAct = Acts[i];
                ActPropertiesSection = CurrentAct.Sections[CardOrd.Properties.ID];
                RowData ModeOfTransfer = ActPropertiesSection.FindRow("@Name = 'Режим передачи'");
                SubSectionData DeviceItem, DeviceCalibrationTime, CompleteItem, CompleteCount, CompleteCalibrationTime;
                switch (ModeOfTransfer.GetInt32(CardOrd.Properties.Value))
                {
                    case 1:     // НОВЫЕ приборы
                        DeviceItem = ActPropertiesSection.FindRow("@Name = 'Заводской номер'").ChildSections[CardOrd.SelectedValues.ID];
                        DeviceCalibrationTime = ActPropertiesSection.FindRow("@Name = 'Время калибровки'").ChildSections[CardOrd.SelectedValues.ID];

                        for (int j = 0; j < DeviceItem.Rows.Count; j++)
                        {
                            Guid DeviceID = UniversalDictionary.GetItemPropertyValue(new Guid(DeviceItem.Rows[j].GetString(CardOrd.SelectedValues.SelectedValue)), "Паспорт прибора").ToGuid();
                            CardData DeviceCard = session.CardManager.GetCardData(DeviceID);
                            double DeviceCalibrationTimeValue = (double)DeviceCalibrationTime.Rows[j].GetDecimal(CardOrd.SelectedValues.SelectedValue);
                            InitiatingDocuments.Add(new ReportItem(session, CurrentAct, DeviceCard, StatusOfTransferValues.NewDevices, DeviceCalibrationTimeValue));
                        }
                        break;
                    case 2:     // НОВЫЕ комплектующие
                        CompleteItem = ActPropertiesSection.FindRow("@Name = 'Наименование компл.'").ChildSections[CardOrd.SelectedValues.ID];
                        CompleteCount = ActPropertiesSection.FindRow("@Name = 'Кол-во компл.'").ChildSections[CardOrd.SelectedValues.ID];
                        CompleteCalibrationTime = ActPropertiesSection.FindRow("@Name = 'Время калибровки компл.'").ChildSections[CardOrd.SelectedValues.ID];

                        for (int j = 0; j < CompleteItem.Rows.Count; j++)
                        {
                            RowData NewCompleteItem = UniversalDictionary.GetItemRow(new Guid(CompleteItem.Rows[j].GetString(CardOrd.SelectedValues.SelectedValue)));
                            int CompleteCountValue = (int)CompleteCount.Rows[j].GetInt32(CardOrd.SelectedValues.SelectedValue);
                            double CompleteCalibrationTimeValue = (double)CompleteCalibrationTime.Rows[j].GetDecimal(CardOrd.SelectedValues.SelectedValue);
                            InitiatingDocuments.Add(new ReportItem(session, CurrentAct, NewCompleteItem, StatusOfTransferValues.NewComplete, CompleteCountValue, CompleteCalibrationTimeValue));
                        }
                        break;
                    case 3:     // Приборы и комплектующие СО СКЛАДА
                        DeviceItem = ActPropertiesSection.FindRow("@Name = 'Заводской номер'").ChildSections[CardOrd.SelectedValues.ID];
                        DeviceCalibrationTime = ActPropertiesSection.FindRow("@Name = 'Время калибровки'").ChildSections[CardOrd.SelectedValues.ID];
                        CompleteItem = ActPropertiesSection.FindRow("@Name = 'Наименование компл.'").ChildSections[CardOrd.SelectedValues.ID];
                        CompleteCount = ActPropertiesSection.FindRow("@Name = 'Кол-во компл.'").ChildSections[CardOrd.SelectedValues.ID];
                        CompleteCalibrationTime = ActPropertiesSection.FindRow("@Name = 'Время калибровки компл.'").ChildSections[CardOrd.SelectedValues.ID];
                        
                        for (int j = 0; j < DeviceItem.Rows.Count; j++)
                        {
                            Guid DeviceID = UniversalDictionary.GetItemPropertyValue(new Guid(DeviceItem.Rows[j].GetString(CardOrd.SelectedValues.SelectedValue)), "Паспорт прибора").ToGuid();
                            CardData DeviceCard = session.CardManager.GetCardData(DeviceID);
                            double DeviceCalibrationTimeValue = (double)DeviceCalibrationTime.Rows[j].GetDecimal(CardOrd.SelectedValues.SelectedValue);
                            InitiatingDocuments.Add(new ReportItem(session, CurrentAct, DeviceCard, StatusOfTransferValues.AfterRecalibration, DeviceCalibrationTimeValue));
                        }
                        
                        for (int j = 0; j < CompleteItem.Rows.Count; j++)
                        {
                            RowData NewCompleteItem = UniversalDictionary.GetItemRow(new Guid(CompleteItem.Rows[j].GetString(CardOrd.SelectedValues.SelectedValue)));
                            int CompleteCountValue = (int)CompleteCount.Rows[j].GetInt32(CardOrd.SelectedValues.SelectedValue);
                            double CompleteCalibrationTimeValue = (double)CompleteCalibrationTime.Rows[j].GetDecimal(CardOrd.SelectedValues.SelectedValue);
                            InitiatingDocuments.Add(new ReportItem(session, CurrentAct, NewCompleteItem, StatusOfTransferValues.AfterRecalibration, CompleteCountValue, 
                                CompleteCalibrationTimeValue));
                        }
                        break;
                }
            }

            // Поиск нарядов на сервисное обслуживание
            CardDataCollection ServiceCards1 = FindCompleteCalibration();
            CardData CurrentServiceCard;

            // Приборы заказчиков, переданные в ремонт
            for (int i = 0; i < ServiceCards1.Count; i++)
            {
                CurrentServiceCard = ServiceCards1[i];
                string DeviceCardID = CurrentServiceCard.Sections[RefServiceCard.MainInfo.ID].FirstRow.GetString(RefServiceCard.MainInfo.DeviceCardID) == null ? "" :
                    CurrentServiceCard.Sections[RefServiceCard.MainInfo.ID].FirstRow.GetString(RefServiceCard.MainInfo.DeviceCardID).ToString();
                CardData DeviceCard = DeviceCardID == "" ? null : session.CardManager.GetCardData(new Guid(DeviceCardID));
                double DeviceCalibrationTimeValue = CurrentServiceCard.Sections[RefServiceCard.Calibration.ID].FirstRow.GetDecimal(RefServiceCard.Calibration.DiagnosticsTime) == null ? 0 :
                    (double)CurrentServiceCard.Sections[RefServiceCard.Calibration.ID].FirstRow.GetDecimal(RefServiceCard.Calibration.DiagnosticsTime);
                InitiatingDocuments.Add(new ReportItem(session, CurrentServiceCard, DeviceCard, StatusOfTransferValues.ClientsDevices, DeviceCalibrationTimeValue, TypeOfTransferValues.ToRepair));

                SectionData AdditionalWaresList = CurrentServiceCard.Sections[RefServiceCard.AdditionalWaresList.ID];
                for (int j = 0; j < AdditionalWaresList.Rows.Count; j++)
                {
                    Guid AdditionalWareItemID = new Guid(AdditionalWaresList.Rows[j].GetString(RefServiceCard.AdditionalWaresList.WaresNumberID));
                    Guid AdditionalWareID = UniversalDictionary.GetItemPropertyValue(AdditionalWareItemID, "Паспорт прибора").ToGuid();
                    CardData AdditionalWareCard = session.CardManager.GetCardData(AdditionalWareID);
                    double AdditionalWareCalibrationTimeValue = AdditionalWaresList.Rows[j].GetDecimal(RefServiceCard.AdditionalWaresList.DiagnosticsTime) == null ? 0 :
                        (double)AdditionalWaresList.Rows[j].GetDecimal(RefServiceCard.AdditionalWaresList.DiagnosticsTime);
                    InitiatingDocuments.Add(new ReportItem(session, CurrentServiceCard, AdditionalWareCard, StatusOfTransferValues.ClientsDevices, AdditionalWareCalibrationTimeValue, TypeOfTransferValues.ToRepair));
                }
            }

            CardDataCollection ServiceCards2 = FindCompleteService();
            Clear();

            // Приборы заказчиков, переданные на склад
            for (int i = 0; i < ServiceCards2.Count; i++)
            {
                CurrentServiceCard = ServiceCards2[i];
                string DeviceCardID = CurrentServiceCard.Sections[RefServiceCard.MainInfo.ID].FirstRow.GetString(RefServiceCard.MainInfo.DeviceCardID) == null ? "" :
                    CurrentServiceCard.Sections[RefServiceCard.MainInfo.ID].FirstRow.GetString(RefServiceCard.MainInfo.DeviceCardID).ToString();
                CardData DeviceCard = DeviceCardID == "" ? null : session.CardManager.GetCardData(new Guid(DeviceCardID));

                double DeviceCalibrationTime = CurrentServiceCard.Sections[RefServiceCard.Calibration.ID].FirstRow.GetDecimal(RefServiceCard.Calibration.CalibrationTime) == null ? 0 :
                    (double)CurrentServiceCard.Sections[RefServiceCard.Calibration.ID].FirstRow.GetDecimal(RefServiceCard.Calibration.CalibrationTime);
                double DeviceDiagnosticsTime = CurrentServiceCard.Sections[RefServiceCard.Calibration.ID].FirstRow.GetDecimal(RefServiceCard.Calibration.DiagnosticsTime) == null ? 0 :
                    (double)CurrentServiceCard.Sections[RefServiceCard.Calibration.ID].FirstRow.GetDecimal(RefServiceCard.Calibration.DiagnosticsTime);
                double DeviceCalibrationTimeValue = CurrentServiceCard.Sections[RefServiceCard.Calibration.ID].FirstRow.GetBoolean(RefServiceCard.Calibration.DeviceRepair) == true ||
                    CurrentServiceCard.Sections[RefServiceCard.Calibration.ID].FirstRow.GetBoolean(RefServiceCard.Calibration.AccessoriesRepair) == true ? DeviceCalibrationTime : 
                    DeviceDiagnosticsTime + DeviceCalibrationTime;
                InitiatingDocuments.Add(new ReportItem(session, CurrentServiceCard, DeviceCard, StatusOfTransferValues.ClientsDevices, DeviceCalibrationTimeValue, TypeOfTransferValues.ToWarehouse));

                SectionData AdditionalWaresList = CurrentServiceCard.Sections[RefServiceCard.AdditionalWaresList.ID];
                for (int j = 0; j < AdditionalWaresList.Rows.Count; j++)
                {
                    Guid AdditionalWareItemID = new Guid(AdditionalWaresList.Rows[j].GetString(RefServiceCard.AdditionalWaresList.WaresNumberID));
                    Guid AdditionalWareID = UniversalDictionary.GetItemPropertyValue(AdditionalWareItemID, "Паспорт прибора").ToGuid();
                    CardData AdditionalWareCard = session.CardManager.GetCardData(AdditionalWareID);

                    double AdditionalWareCalibrationTime = AdditionalWaresList.Rows[j].GetDecimal(RefServiceCard.AdditionalWaresList.CalibrationTime) == null ? 0 :
                        (double)AdditionalWaresList.Rows[j].GetDecimal(RefServiceCard.AdditionalWaresList.CalibrationTime);
                    double AdditionalWareDiagnosticsTime = AdditionalWaresList.Rows[j].GetDecimal(RefServiceCard.AdditionalWaresList.DiagnosticsTime) == null ? 0 :
                        (double)AdditionalWaresList.Rows[j].GetDecimal(RefServiceCard.AdditionalWaresList.DiagnosticsTime);

                    double AdditionalWareCalibrationTimeValue = CurrentServiceCard.Sections[RefServiceCard.Calibration.ID].FirstRow.GetBoolean(RefServiceCard.Calibration.DeviceRepair) == true ||
                        CurrentServiceCard.Sections[RefServiceCard.Calibration.ID].FirstRow.GetBoolean(RefServiceCard.Calibration.AccessoriesRepair) == true ? AdditionalWareCalibrationTime :
                        AdditionalWareCalibrationTime + AdditionalWareDiagnosticsTime;
                    InitiatingDocuments.Add(new ReportItem(session, CurrentServiceCard, AdditionalWareCard, StatusOfTransferValues.ClientsDevices, AdditionalWareCalibrationTimeValue, TypeOfTransferValues.ToWarehouse));
                }
            }
            return InitiatingDocuments;
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
        /// <summary>
        /// Поиск Актов передачи приборов и комплектующих, принятых в отчетный период
        /// </summary>
        private CardDataCollection FindAct()
        {
            /* Поиск Актов передачи приборов и комплектующих */
            SearchQuery searchQuery = session.CreateSearchQuery();
            searchQuery.CombineResults = ConditionGroupOperation.And;
            CardTypeQuery typeQuery = searchQuery.AttributiveSearch.CardTypeQueries.AddNew(CardOrd.ID);
            SectionQuery sectionQuery = typeQuery.SectionQueries.AddNew(CardOrd.MainInfo.ID);
            sectionQuery.Operation = SectionQueryOperation.And;
            sectionQuery.ConditionGroup.Operation = ConditionGroupOperation.And;
            sectionQuery.ConditionGroup.Conditions.AddNew(CardOrd.MainInfo.Type, FieldType.RefId, ConditionOperation.Equals, ActTypeID);
            // Вид передачи - "Калибровка -> Сбыт"
            sectionQuery = typeQuery.SectionQueries.AddNew(CardOrd.Properties.ID);
            sectionQuery.Operation = SectionQueryOperation.And;
            sectionQuery.ConditionGroup.Operation = ConditionGroupOperation.And;
            sectionQuery.ConditionGroup.Conditions.AddNew(CardOrd.Properties.Name, FieldType.Unistring, ConditionOperation.Equals, "Вид передачи");
            sectionQuery.ConditionGroup.Conditions.AddNew(CardOrd.Properties.Value, FieldType.Unistring, ConditionOperation.Equals, "Калибровка -> Сбыт");
            // Состояние акта - "Принят"
            sectionQuery = typeQuery.SectionQueries.AddNew(CardOrd.Properties.ID);
            sectionQuery.Operation = SectionQueryOperation.And;
            sectionQuery.ConditionGroup.Operation = ConditionGroupOperation.And;
            sectionQuery.ConditionGroup.Conditions.AddNew(CardOrd.Properties.Name, FieldType.Unistring, ConditionOperation.Equals, "Состояние акта");
            sectionQuery.ConditionGroup.Conditions.AddNew(CardOrd.Properties.Value, FieldType.Int, ConditionOperation.Equals, 3);
            // Дата принятия акта находится в пределах параметров отчета
            sectionQuery = typeQuery.SectionQueries.AddNew(CardOrd.Properties.ID);
            sectionQuery.Operation = SectionQueryOperation.And;
            sectionQuery.ConditionGroup.Operation = ConditionGroupOperation.And;
            sectionQuery.ConditionGroup.Conditions.AddNew(CardOrd.Properties.Name, FieldType.Unistring, ConditionOperation.Equals, "Дата принятия");
            sectionQuery.ConditionGroup.Conditions.AddNew(CardOrd.Properties.Value, FieldType.Date, ConditionOperation.GreaterEqual, StartDate);
            sectionQuery.ConditionGroup.Conditions.AddNew(CardOrd.Properties.Value, FieldType.Date, ConditionOperation.LessThan, EndDate.AddDays(1));
            // Получение текста запроса
            searchQuery.Limit = 0;
            string query = searchQuery.GetXml();
            return session.CardManager.FindCards(query);
        }
        /// <summary>
        /// Поиск нарядов на сервисное обслуживание, переданных на ремонт в отчетный период
        /// </summary>
        private CardDataCollection FindCompleteCalibration()
        {
            /* Поиск Актов передачи приборов и комплектующих */
            SearchQuery searchQuery = session.CreateSearchQuery();
            searchQuery.CombineResults = ConditionGroupOperation.And;
            CardTypeQuery typeQuery = searchQuery.AttributiveSearch.CardTypeQueries.AddNew(ServiceCardTypeID);
            // Дата окончания калибровки находится в пределах параметров отчета
            SectionQuery sectionQuery = typeQuery.SectionQueries.AddNew(RefServiceCard.Calibration.ID);
            sectionQuery.Operation = SectionQueryOperation.And;
            sectionQuery.ConditionGroup.Operation = ConditionGroupOperation.And;
            sectionQuery.ConditionGroup.Conditions.AddNew(RefServiceCard.Calibration.CalibDateEnd, FieldType.DateTime, ConditionOperation.GreaterEqual, StartDate);
            sectionQuery.ConditionGroup.Conditions.AddNew(RefServiceCard.Calibration.CalibDateEnd, FieldType.DateTime, ConditionOperation.LessThan, EndDate.AddDays(1));
            // Получение текста запроса
            searchQuery.Limit = 0;
            string query = searchQuery.GetXml();
            return session.CardManager.FindCards(query);
        }
        /// <summary>
        /// Поиск нарядов на сервисное обслуживание, сданных на склад в отчетный период
        /// </summary>
        private CardDataCollection FindCompleteService()
        {
            /* Поиск Актов передачи приборов и комплектующих */
            SearchQuery searchQuery = session.CreateSearchQuery();
            searchQuery.CombineResults = ConditionGroupOperation.And;
            CardTypeQuery typeQuery = searchQuery.AttributiveSearch.CardTypeQueries.AddNew(ServiceCardTypeID);
            // Дата окончания сервисного обслуживания находится в пределах параметров отчета
            SectionQuery sectionQuery = typeQuery.SectionQueries.AddNew(RefServiceCard.MainInfo.ID);
            sectionQuery.Operation = SectionQueryOperation.And;
            sectionQuery.ConditionGroup.Operation = ConditionGroupOperation.And;
            sectionQuery.ConditionGroup.Conditions.AddNew(RefServiceCard.MainInfo.DateEndFact, FieldType.DateTime, ConditionOperation.GreaterEqual, StartDate);
            sectionQuery.ConditionGroup.Conditions.AddNew(RefServiceCard.MainInfo.DateEndFact, FieldType.DateTime, ConditionOperation.LessEqual, EndDate.AddDays(1));
            // Получение текста запроса
            searchQuery.Limit = 0;
            string query = searchQuery.GetXml();
            return session.CardManager.FindCards(query);
        }
        /// <summary>
        /// Добавление в отчет столбцов "Факт"
        /// </summary>
        /// <param name="Worksheet">Рабочий лист.</param>
        private void AddFactColumns(Excel.Worksheet Worksheet)
        {
            Worksheet.Columns[9].Insert(Excel.XlInsertShiftDirection.xlShiftToRight, System.Type.Missing);
            Worksheet.Columns[9].Insert(Excel.XlInsertShiftDirection.xlShiftToRight, System.Type.Missing);
            Worksheet.Columns[9].Insert(Excel.XlInsertShiftDirection.xlShiftToRight, System.Type.Missing);
            Worksheet.Columns["F:H"].Copy();
            Worksheet.Paste(Worksheet.Cells[1, 9]);
            bool finish = false;
            while (finish == false)
            {
                if ((Worksheet.Cells[RowsCount, 1].MergeCells == true) && (Worksheet.Cells[RowsCount, 1].Value != null) && (Worksheet.Cells[RowsCount, 1].Value != "Наименование прибора"))
                { Worksheet.Range[Worksheet.Cells[RowsCount, 1], Worksheet.Cells[RowsCount, ColumnsCount]].Merge(); }

                if (Worksheet.Cells[RowsCount, 11].FormulaR1C1 == "=RC[-6]*(RC[-1]+RC[-2])")
                {
                    Worksheet.Cells[RowsCount, 11].FormulaR1C1 = "=RC[-9]*(RC[-1]+RC[-2])";
                    SetGrayInterior(Worksheet.Range[Worksheet.Cells[RowsCount, 6], Worksheet.Cells[RowsCount, 7]]);
                    SetGrayInterior(Worksheet.Range[Worksheet.Cells[RowsCount, 9], Worksheet.Cells[RowsCount, 10]]);
                }

                if (Worksheet.Cells[RowsCount, 11].FormulaR1C1 == "=RC[-6]*RC[-2]")
                {
                    Worksheet.Cells[RowsCount, 11].FormulaR1C1 = "=RC[-9]*RC[-2]";
                    Worksheet.Cells[RowsCount, 11].FormulaR1C1 = "=RC[-9]*(RC[-1]+RC[-2])";
                    SetGrayInterior(Worksheet.Range[Worksheet.Cells[RowsCount, 3], Worksheet.Cells[RowsCount, 3]]);
                    SetGrayInterior(Worksheet.Range[Worksheet.Cells[RowsCount, 6], Worksheet.Cells[RowsCount, 7]]);
                    SetGrayInterior(Worksheet.Range[Worksheet.Cells[RowsCount, 9], Worksheet.Cells[RowsCount, 10]]);
                }

                if (GetString(Worksheet.Cells[RowsCount, 9].Value) == "План сдачи на склад, шт")
                { Worksheet.Cells[RowsCount, 9].Value = "Сдано на склад, шт"; }
                else
                {
                    if (GetString(Worksheet.Cells[RowsCount, 9].Value) == "План сдачи, шт")
                    { Worksheet.Cells[RowsCount, 9].Value = "Сдано, шт"; }
                    else
                    {
                        if ((GetString(Worksheet.Cells[RowsCount, 9].Value) != "В ремонт") && (GetString(Worksheet.Cells[RowsCount, 9].FormulaR1C1).IndexOf("=") < 0))
                        { Worksheet.Cells[RowsCount, 9].Value = ""; }
                        if ((GetString(Worksheet.Cells[RowsCount, 9].Value) != "На склад") && (GetString(Worksheet.Cells[RowsCount, 10].FormulaR1C1).IndexOf("=") < 0))
                        { Worksheet.Cells[RowsCount, 10].Value = ""; }
                    }
                }
                if (GetString(Worksheet.Cells[RowsCount, 1].Value) == "ИТОГО:") finish = true;
                RowsCount++;
            }
        }
        /// <summary>
        /// Поиск партии в отчете
        /// </summary>
        /// <param name="Worksheet">Рабочий лист.</param>
        /// <param name="PartyName">Название партии.</param>
        /// <param name="DeviceType">Тип прибора.</param>
        /// <param name="planNorm">Плановая норма трудоемкости на калибровку одного прибора.</param>
        private int FindParty(Excel.Worksheet Worksheet, string PartyName, string DeviceType, double planNorm)
        {
            int PartyPosition = 0;
            int DeviceTypePosition = 0;
            int NextSectionPosition = 0;
            int SummaryItemPosition = 0;
            // Поиск партии в отчете
            for (int i = StartPosition; i < RowsCount; i++)
            {
                if (Worksheet.Cells[i, 1].Value == PartyName)
                {
                    PartyPosition = i;
                    Worksheet.Cells[PartyPosition, 9].Value = GetInt32(Worksheet.Cells[PartyPosition, 9].Value) + 1;
                    break;
                }
            }
            // Если партия в отчете не найдена
            if (PartyPosition == 0)
            {
                // Поиск типа прибора
                for (int i = StartPosition; i < RowsCount; i++)
                { if (GetString(Worksheet.Cells[i, 1].Value) == DeviceType) { DeviceTypePosition = i; break; } }

                // Если тип прибора не найден
                if (DeviceTypePosition == 0)
                {
                    if ((DeviceType == "ДП12") || (DeviceType == "ДП21") || (DeviceType == "ДП22"))  // Определение местоположения датчиков в отчете
                    {
                        DeviceTypePosition = StartPosition;
                        // Поиск итоговой строки по всем датчикам
                        for (int i = StartPosition; i < RowsCount; i++)
                        { if (GetString(Worksheet.Cells[i, 1].Value) == "Всего датчиков:") { SummaryItemPosition = i; break; } }
                    }
                    else                                                                             // Определение местоположения приборов в отчете
                    {
                        // Поиск раздела "Приборы заказчиков" в отчете
                        for (int i = StartPosition; i < RowsCount; i++)
                        {
                            if ((GetString(Worksheet.Cells[i, 1].Value) == "Приборы заказчиков") || (GetString(Worksheet.Cells[i, 1].Value) == "Новые комплектующие") ||
                              (GetString(Worksheet.Cells[i, 1].Value) == "Приборы и комплектующие после повторной калибровки")) { NextSectionPosition = i; break; }
                        }
                        DeviceTypePosition = NextSectionPosition - 2;
                        // Поиск итоговой строки по всем приборам
                        for (int i = StartPosition; i < RowsCount; i++)
                        { if (GetString(Worksheet.Cells[i, 1].Value) == "Всего приборов:") { SummaryItemPosition = i; break; } }
                    }
                    // Добавление типа прибора/датчика в отчет
                    AddDeviceType(Worksheet, DeviceTypePosition, DeviceType, planNorm);
                    RowsCount++;
                    SummaryItemPosition++;
                    // Обновление формулы суммарной трудоемкости по всем приборам/датчикам
                    Worksheet.Cells[SummaryItemPosition, 11].FormulaLocal = GetString(Worksheet.Cells[SummaryItemPosition, 11].FormulaLocal).IndexOf("=СУММ") >= 0 ?
                        GetString(Worksheet.Cells[SummaryItemPosition, 11].FormulaLocal).Replace(")", ";K" + DeviceTypePosition + ")") : "=СУММ(K" + DeviceTypePosition + ")";
                }

                // Создание новой партии
                PartyPosition = DeviceTypePosition + 1;
                AddParty(Worksheet, PartyPosition, PartyName);
                RowsCount++;
                // Обновление формулы суммарной трудоемкости по типу прибора
                int ReportPosition = PartyPosition;
                while (Worksheet.Cells[ReportPosition, 1].Font.Bold == false) ReportPosition++;
                Worksheet.Cells[DeviceTypePosition, 9].FormulaR1C1 = "=SUM(R[1]C:R[" + (ReportPosition - PartyPosition) + "]C)";
                // Форматирование границ
                SetEdgeBorders(Worksheet.Range[Worksheet.Cells[DeviceTypePosition, 1], Worksheet.Cells[ReportPosition - 1, ColumnsCount]], Excel.XlLineStyle.xlContinuous, 0, 0, Excel.XlBorderWeight.xlMedium);
            }
            return PartyPosition;
        }
        /// <summary>
        /// Поиск в отчете прибора, принадлежащего заказчику
        /// </summary>
        /// <param name="Worksheet">Рабочий лист.</param>
        /// <param name="DeviceName">Название прибора (в одном из следующих форматов: МИКО-1 № 1/2014; МИКО-1 № 001C/2012; МИКО-1 (Только комплектующие)).</param>
        /// <param name="ClientName">Название клиента.</param>
        /// <param name="Norm">Фактическая трудоемкость калибровки прибора.</param>
        /// <param name="FactTime">Фактическое время.</param>
        /// <param name="State">Текущее состояние прибора.</param>
        /// <param name="Service">Вид сервиса.</param>
        /// <param name="EndDate">Плановая дата окончания сервисного обслуживания.</param>
        /// <param name="WarrantyType">Гарантийные обязательства.</param>
        /// <param name="TypeOfTransfer">Тип передачи.</param>
        /// <param name="ServiceNumber">Номер заявки на сервисное обслуживание.</param>
        /// <param name="ServiceComment">Комментарий сервисного специалиста к заявке на сервисное обслуживание.</param>
        private int FindDevice(Excel.Worksheet Worksheet, string DeviceName, string ClientName, double Norm, double FactTime, string State, string Service, DateTime EndDate, string WarrantyType, string TypeOfTransfer, string ServiceNumber, string ServiceComment)
        {
            int DevicePosition = 0;
            int WarrantyPosition = 0;
            int NonWarrantyPosition = 0;
            int WarrantyTypePosition = 0;
            int ClientsPosition = 0;

            // Поиск ремонтного прибора в отчете
            for (int i = StartPosition; i < RowsCount; i++)
            {
                if (GetString(Worksheet.Cells[i, 1].Value) == DeviceName)
                {
                    // Обновление данных для найденного прибора
                    DevicePosition = i;
                    Worksheet.Cells[i, 3].Value = State;
                    Worksheet.Cells[i, 11].Value = GetDouble(Worksheet.Cells[i, 11].Value) + FactTime;
                    if (TypeOfTransfer == TypeOfTransferValues.ToRepair) Worksheet.Cells[i, 9].Value = 1;
                    if (TypeOfTransfer == TypeOfTransferValues.ToWarehouse) Worksheet.Cells[i, 10].Value = 1;
                    break;
                }
            }
            // Если прибор не найден, поиск заказчика в отчете
            if (DevicePosition == 0)
            {
                for (int i = StartPosition; i < RowsCount; i++)
                {
                    if (GetString(Worksheet.Cells[i, 1].Value) == ClientName)
                    {
                        // Создание записи для прибора
                        AddClientsDevice(Worksheet, i + 1, DeviceName, Norm, FactTime, State, Service, EndDate, TypeOfTransfer, ClientName, ServiceNumber, ServiceComment);
                        RowsCount++;
                        DevicePosition = i + 1;
                        break;
                    }
                }
            }
            // Если заказчик не найден, поиск раздела с нужным типом гарантии
            if (DevicePosition == 0)
            {
                // Поиск раздела "Гарантийные"
                for (int i = StartPosition; i < RowsCount; i++)
                { if (GetString(Worksheet.Cells[i, 1].Value) == TypeOfWarranty.Warranty) { WarrantyPosition = i; break; } }
                // Поиск раздела "Негарантийные"
                for (int i = StartPosition; i < RowsCount; i++)
                { if (GetString(Worksheet.Cells[i, 1].Value) == TypeOfWarranty.NonWarranty) { NonWarrantyPosition = i; break; } }
                // Определение раздела с нужным типом гарантии
                WarrantyTypePosition = WarrantyType == TypeOfWarranty.Warranty ? WarrantyPosition : NonWarrantyPosition;

                if (WarrantyTypePosition == 0)        // Если раздел с нужным типом гарантии не найден
                {
                    // Создание нового раздела для гарантийного типа
                    WarrantyTypePosition = RowsCount;
                    AddWarrantyType(Worksheet, WarrantyTypePosition, WarrantyType);
                    if (WarrantyType == TypeOfWarranty.Warranty) WarrantyPosition = WarrantyTypePosition;
                    if (WarrantyType == TypeOfWarranty.NonWarranty) NonWarrantyPosition = WarrantyTypePosition;
                    RowsCount++;
                    // Создание итоговой строки для нового раздела гарантийного типа
                    AddSummaryItem(Worksheet, WarrantyTypePosition + 1);
                    RowsCount++;
                    // Создание нового раздела для заказчика
                    ClientsPosition = WarrantyTypePosition + 1;
                    AddClient(Worksheet, ClientsPosition, ClientName);
                    RowsCount++;
                    // Создание новой записи для прибора
                    DevicePosition = ClientsPosition + 1;
                    AddClientsDevice(Worksheet, DevicePosition, DeviceName, Norm, FactTime, State, Service, EndDate, TypeOfTransfer, ClientName, ServiceNumber, ServiceComment);
                    RowsCount++;
                    // Формула суммарной трудоемкости по новому разделу
                    Worksheet.Cells[DevicePosition + 1, 11].FormulaLocal = "=СУММ(K" + WarrantyTypePosition + ":K" + DevicePosition + ")";
                }
                else                                 // Если раздел с нужным типом гарантии найден
                {
                    // Создание нового раздела для заказчика
                    ClientsPosition = WarrantyTypePosition + 1;
                    AddClient(Worksheet, ClientsPosition, ClientName);
                    RowsCount++;
                    // Создание записи для прибора
                    DevicePosition = ClientsPosition + 1;
                    AddClientsDevice(Worksheet, DevicePosition, DeviceName, Norm, FactTime, State, Service, EndDate, TypeOfTransfer, ClientName, ServiceNumber, ServiceComment);
                    RowsCount++;
                    // Обновление формулы суммарной трудоемкости по разделу
                    int FinalyItemPosition = 0;
                    if (WarrantyType == TypeOfWarranty.Warranty) { FinalyItemPosition = WarrantyPosition < NonWarrantyPosition ? NonWarrantyPosition + 1 : RowsCount - 1; }
                    if (WarrantyType == TypeOfWarranty.NonWarranty) { FinalyItemPosition = NonWarrantyPosition < WarrantyPosition ? WarrantyPosition + 1 : RowsCount - 1; }

                    Worksheet.Cells[FinalyItemPosition, 11].FormulaLocal = "=СУММ(K" + WarrantyTypePosition + ":K" + (FinalyItemPosition - 1) + ")";
                }
            }
            return DevicePosition;
        }
        /// <summary>
        /// Поиск в отчете новых комплектующих
        /// </summary>
        /// <param name="Worksheet">Рабочий лист.</param>
        /// <param name="CompleteName">Название комплектующего в формате "'Название' 'Код СКБ'".</param>
        /// <param name="WorkOrderName">Наряд-заказ на изготовление комплектующих/партия.</param>
        /// <param name="Norm">Фактическая трудоемкость проверки комплектующего.</param>
        /// <param name="Count">Количество комплектующего.</param>
        private int FindComplete(Excel.Worksheet Worksheet, string CompleteName, string WorkOrderName, double Norm, int Count)
        {
            int NewCompletesPosition = 0;
            int ClientsDevicesPosition = 0;
            int AfterRecalibrationPosition = 0;
            int FindCompletePosition = 0;
            int NextSectionPosition = 0;

            // Поиск раздела "Новые комплекрующие" в отчете
            for (int i = StartPosition; i < RowsCount; i++)
            { if (GetString(Worksheet.Cells[i, 1].Value) == "Новые комплектующие") { NewCompletesPosition = i; break; } }
            // Поиск раздела "Приборы заказчиков" в отчете
            for (int i = StartPosition; i < RowsCount; i++)
            { if (GetString(Worksheet.Cells[i, 1].Value) == "Приборы заказчиков") { ClientsDevicesPosition = i; break; } }
            // Поиск раздела "Приборы заказчиков" в отчете
            for (int i = StartPosition; i < RowsCount; i++)
            { if (GetString(Worksheet.Cells[i, 1].Value) == "Приборы и комплектующие после повторной калибровки") { AfterRecalibrationPosition = i; break; } }
            NextSectionPosition = Math.Min(AfterRecalibrationPosition, ClientsDevicesPosition) != 0 ? Math.Min(AfterRecalibrationPosition, ClientsDevicesPosition) : Math.Max(AfterRecalibrationPosition, ClientsDevicesPosition);

            // Если раздел "Новые комплектующие" не найден
            if (NewCompletesPosition == 0)
            {
                // Создание раздела "Новые комплектующие"
                NewCompletesPosition = NextSectionPosition;
                AddCompleteSection(Worksheet, NewCompletesPosition);
                NextSectionPosition++;
                RowsCount++;
            }
            // Если итоговая строка для раздела "Новые комплектующие" не найдена
            if (GetString(Worksheet.Cells[NextSectionPosition - 1, 1].Value).IndexOf("Всего:") < 0)
            {
                // Создание итоговой строки для раздел "Новые комплектующие"
                AddSummaryItem(Worksheet, NextSectionPosition);
                NextSectionPosition++;
                RowsCount++;
            }
            // Поиск комплектующего в отчете
            for (int i = StartPosition; i < RowsCount; i++)
            {
                if ((GetString(Worksheet.Cells[i, 1].Value).IndexOf(CompleteName) >= 0) && (GetString(Worksheet.Cells[i, 1].Value).IndexOf(CompleteName + "-") < 0))
                {
                    // Обновление данных для найденного комплектующего
                    FindCompletePosition = i;
                    Worksheet.Cells[i, 9].Value = GetInt32(Worksheet.Cells[i, 9].Value) + Count;
                    Worksheet.Cells[i, 11].Value = GetDouble(Worksheet.Cells[i, 11].Value) + Norm;
                    break;
                }
            }
            // Если комплектующее в отчете не найдено
            if (FindCompletePosition == 0)
            {
                FindCompletePosition = NextSectionPosition - 1;
                // Добавление нового комплектующего в отчет
                AddComplete(Worksheet, FindCompletePosition, CompleteName, Norm, Count, 0D);
                NextSectionPosition++;
                RowsCount++;
            }
            // Обновление формул суммарного количества и трудоемкости по разделу "Новые комплектующие"
            Worksheet.Cells[NextSectionPosition - 1, 11].FormulaLocal = "=СУММ(K" + (NewCompletesPosition + 1) + ":K" + (NextSectionPosition - 2) + ")";
            Worksheet.Cells[NextSectionPosition - 1, 9].FormulaLocal = "=СУММ(I" + (NewCompletesPosition + 1) + ":I" + (NextSectionPosition - 2) + ")";
            return FindCompletePosition;
        }
        /// <summary>
        /// Поиск в отчете приборов/комплектующих после повторной калибровки
        /// </summary>
        /// <param name="Worksheet">Рабочий лист.</param>
        /// <param name="Name">Название прибора/комплектующего.</param>
        /// <param name="WorkOrderName">Родительский акт передачи.</param>
        /// <param name="Norm">Фактическая трудоемкость калибровки прибора/комплектующего.</param>
        /// <param name="Count">Количество приборов/комплектующих.</param>
        /// <param name="planNorm">Количество комплектующего.</param>
        private int FindDeviceOrCompleteAfterRecalibration(Excel.Worksheet Worksheet, string Name, string WorkOrderName, double Norm, int Count, double planNorm)
        {
            int AfterRecalibrationPosition = 0;
            int ClientsDevicesPosition = 0;
            int CompletePosition = 0;

            // Поиск раздела "Новые комплекрующие" в отчете
            for (int i = StartPosition; i < RowsCount; i++)
            { if (Worksheet.Cells[i, 1].Value == "Приборы и комплектующие после повторной калибровки") { AfterRecalibrationPosition = i; break; } }
            // Поиск раздела "Приборы заказчиков" в отчете
            for (int i = StartPosition; i < RowsCount; i++)
            { if (Worksheet.Cells[i, 1].Value == "Приборы заказчиков") { ClientsDevicesPosition = i; break; } }

            // Если раздел "Приборы и комплектующие после повторной калибровки" не найден
            if (AfterRecalibrationPosition == 0)
            {
                // Создание раздела "Новые комплектующие"
                AfterRecalibrationPosition = ClientsDevicesPosition;
                AddSectionAfterRecalibration(Worksheet, AfterRecalibrationPosition);
                ClientsDevicesPosition++;
                RowsCount++;
            }
            // Если итоговая строка для раздела "Приборы и комплектующие после повторной калибровки" не найдена
            if (Worksheet.Cells[ClientsDevicesPosition - 1, 1].Value.ToString().IndexOf("Всего:") < 0)
            {
                // Создание итоговой строки для раздела "Приборы и комплектующие после повторной калибровки"
                AddSummaryItem(Worksheet, ClientsDevicesPosition);
                ClientsDevicesPosition++;
                RowsCount++;
            }
            // Поиск прибора/комплектующего в отчете
            for (int i = StartPosition; i < RowsCount; i++)
            {
                if ((GetString(Worksheet.Cells[i, 1].Value).IndexOf(Name) >= 0) && (GetString(Worksheet.Cells[i, 1].Value).IndexOf(Name + "-") < 0))
                {
                    // Обновление данных для найденного прбора/комплектующего
                    CompletePosition = i;
                    Worksheet.Cells[i, 9].Value = GetInt32(Worksheet.Cells[i, 9].Value) + Count;
                    Worksheet.Cells[i, 11].Value = GetDouble(Worksheet.Cells[i, 11].Value) + Norm;
                    break;
                }
            }
            // Если прибор/комплектующее в отчете не найдено
            if (CompletePosition == 0)
            {
                // Добавление прибора/комплектующего в отчет
                CompletePosition = ClientsDevicesPosition - 1;
                AddComplete(Worksheet, CompletePosition, Name, Norm, Count, planNorm);
                ClientsDevicesPosition++;
                RowsCount++;
            }
            // Обновление формул суммарного количества и трудоемкости по разделу "Приборы и комплектующие после повторной калибровки"
            Worksheet.Cells[ClientsDevicesPosition - 1, 11].FormulaLocal = "=СУММ(K" + (AfterRecalibrationPosition) + ":K" + (ClientsDevicesPosition - 2) + ")";
            Worksheet.Cells[ClientsDevicesPosition - 1, 9].FormulaLocal = "=СУММ(I" + (AfterRecalibrationPosition) + ":I" + (ClientsDevicesPosition - 2) + ")";
            return CompletePosition;
        }
        /// <summary>
        /// Добавление в отчет прибора, принадлежащего заказчику
        /// </summary>
        /// <param name="Worksheet">Рабочий лист.</param>
        /// <param name="RowIndex">Индекс строки.</param>
        /// <param name="DeviceName">Название прибора.</param>
        /// <param name="Norm">Норма трудоемкости калибровки прибора.</param>
        /// <param name="FactTime">Фактическая трудоемкость калибровки прибора.</param>
        /// <param name="State">Текущее состояние прибора.</param>
        /// <param name="Service">Вид сервиса.</param>
        /// <param name="EndDate">Плановая дата окончания сервисного обслуживания.</param>
        /// <param name="TypeOfTransfer">Тип передачи.</param>
        /// <param name="Client">Название заказчика.</param>
        /// <param name="ServiceNumber">Номер заявки на сервисное обслуживание.</param>
        /// <param name="ServiceComment">Комментарий сервисного специалиста.</param>
        private static void AddClientsDevice(Excel.Worksheet Worksheet, int RowIndex, string DeviceName, double Norm, double FactTime, string State, string Service, DateTime EndDate, string TypeOfTransfer, string Client, string ServiceNumber, string ServiceComment)
        {
            // Добавление новой строки в отчет
            Worksheet.Rows[RowIndex].Insert(Excel.XlInsertShiftDirection.xlShiftDown, System.Type.Missing);
            // Форматирование
            ApplyStyle2(Worksheet.Range[Worksheet.Cells[RowIndex, 1], Worksheet.Cells[RowIndex, ColumnsCount]]);
            // Заполнение данных
            Worksheet.Cells[RowIndex, 1].Value = DeviceName;
            Worksheet.Cells[RowIndex, 1].WrapText = true;
            Worksheet.Cells[RowIndex, 2].Value = Norm;
            Worksheet.Cells[RowIndex, 3].Value = State;
            Worksheet.Cells[RowIndex, 3].WrapText = true;
            Worksheet.Cells[RowIndex, 4].Value = Service;
            Worksheet.Cells[RowIndex, 4].WrapText = true;
            Worksheet.Cells[RowIndex, 5].Value = EndDate.ToString("d.MM.yy");
            Worksheet.Cells[RowIndex, 13].Value = Client;
            Worksheet.Cells[RowIndex, 14].Value = ServiceNumber;
            Worksheet.Cells[RowIndex, 15].Value = ServiceComment;
            if (TypeOfTransfer == TypeOfTransferValues.ToRepair) Worksheet.Cells[RowIndex, 9].Value = 1;
            if (TypeOfTransfer == TypeOfTransferValues.ToWarehouse) Worksheet.Cells[RowIndex, 10].Value = 1;
            SetGrayInterior(Worksheet.Range[Worksheet.Cells[RowIndex, 6], Worksheet.Cells[RowIndex, 7]]);
            SetGrayInterior(Worksheet.Range[Worksheet.Cells[RowIndex, 9], Worksheet.Cells[RowIndex, 10]]);
            Worksheet.Cells[RowIndex, 11].Value = GetDouble(Worksheet.Cells[RowIndex, 11].Value) + FactTime;
        }
        /// <summary>
        /// Добавление в отчет клиента
        /// </summary>
        /// <param name="Worksheet">Рабочий лист.</param>
        /// <param name="RowIndex">Индекс строки.</param>
        /// <param name="ClientName">Название клиента.</param>
        private static void AddClient(Excel.Worksheet Worksheet, int RowIndex, string ClientName)
        {
            Worksheet.Rows[RowIndex].Insert(Excel.XlInsertShiftDirection.xlShiftDown, System.Type.Missing);
            ApplyStyle4(Worksheet.Range[Worksheet.Cells[RowIndex, 1], Worksheet.Cells[RowIndex, ColumnsCount]]);
            Worksheet.Range[Worksheet.Cells[RowIndex, 1], Worksheet.Cells[RowIndex, ColumnsCount]].Merge();
            Worksheet.Cells[RowIndex, 1].Value = ClientName;
        }
        /// <summary>
        /// Добавление в отчет гарантийного типа
        /// </summary>
        /// <param name="Worksheet">Рабочий лист.</param>
        /// <param name="RowIndex">Индекс строки.</param>
        /// <param name="Warranty">Тип гарантии.</param>
        private static void AddWarrantyType(Excel.Worksheet Worksheet, int RowIndex, string Warranty)
        {
            Worksheet.Rows[RowIndex].Insert(Excel.XlInsertShiftDirection.xlShiftDown, System.Type.Missing);
            Worksheet.Cells[RowIndex, 1].Value = Warranty;
            ApplyStyle1(Worksheet.Range[Worksheet.Cells[RowIndex, 1], Worksheet.Cells[RowIndex, ColumnsCount]]);
            Worksheet.Range[Worksheet.Cells[RowIndex, 1], Worksheet.Cells[RowIndex, ColumnsCount]].Merge();
        }
        /// <summary>
        /// Добавление в отчет итоговой строки к разделу
        /// </summary>
        /// <param name="Worksheet">Рабочий лист.</param>
        /// <param name="RowIndex">Индекс строки.</param>
        private static void AddSummaryItem(Excel.Worksheet Worksheet, int RowIndex)
        {
            // Создание итоговой строки для раздела "Новые комплектующие"
            Worksheet.Rows[RowIndex].Insert(Excel.XlInsertShiftDirection.xlShiftDown, System.Type.Missing);
            ApplyStyle3(Worksheet.Range[Worksheet.Cells[RowIndex, 1], Worksheet.Cells[RowIndex, ColumnsCount]]);
            Worksheet.Cells[RowIndex, 1].Value = "Всего:";
            Worksheet.Range[Worksheet.Cells[RowIndex, 6], Worksheet.Cells[RowIndex, 7]].Merge();
            Worksheet.Range[Worksheet.Cells[RowIndex, 9], Worksheet.Cells[RowIndex, 10]].Merge();
        }
        /// <summary>
        /// Добавление в отчет раздела "Новые комплектующие"
        /// </summary>
        /// <param name="Worksheet">Рабочий лист.</param>
        /// <param name="RowIndex">Индекс строки.</param>
        private static void AddCompleteSection(Excel.Worksheet Worksheet, int RowIndex)
        {
            Worksheet.Rows[RowIndex].Insert(Excel.XlInsertShiftDirection.xlShiftDown, System.Type.Missing);
            Worksheet.Cells[RowIndex, 1].Value = "Новые комплектующие";
            Worksheet.Range[Worksheet.Cells[RowIndex, 1], Worksheet.Cells[RowIndex, ColumnsCount]].Merge();
            ApplyStyle5(Worksheet.Range[Worksheet.Cells[RowIndex, 1], Worksheet.Cells[RowIndex, ColumnsCount]]);
        }
        /// <summary>
        /// Добавление в отчет раздела "Приборы и комплектующие после повторной калибровки"
        /// </summary>
        /// <param name="Worksheet">Рабочий лист.</param>
        /// <param name="RowIndex">Индекс строки.</param>
        private static void AddSectionAfterRecalibration(Excel.Worksheet Worksheet, int RowIndex)
        {
            Worksheet.Rows[RowIndex].Insert(Excel.XlInsertShiftDirection.xlShiftDown, System.Type.Missing);
            Worksheet.Cells[RowIndex, 1].Value = "Приборы и комплектующие после повторной калибровки";
            Worksheet.Range[Worksheet.Cells[RowIndex, 1], Worksheet.Cells[RowIndex, ColumnsCount]].Merge();
            ApplyStyle5(Worksheet.Range[Worksheet.Cells[RowIndex, 1], Worksheet.Cells[RowIndex, ColumnsCount]]);
        }
        /// <summary>
        /// Добавление в отчет комплектующего/ прибора после повторной калибровки.
        /// </summary>
        /// <param name="Worksheet">Рабочий лист.</param>
        /// <param name="RowIndex">Индекс строки.</param>
        /// <param name="CompleteName">Название комплектующего/ прибора после повторной калибровки.</param>
        /// <param name="Norm">Фактическая трудоемкость калибровки комплектующего/ прибора после повторной калибровки.</param>
        /// <param name="Count">Количество комплектующих/ приборов после повторной калибровки.</param>
        /// <param name="planNorm">Плановая норма калибровки прибора после повторной калибровки.</param>
        private static void AddComplete(Excel.Worksheet Worksheet, int RowIndex, string CompleteName, double Norm, int Count, double planNorm)
        {
            Worksheet.Rows[RowIndex].Insert(Excel.XlInsertShiftDirection.xlShiftDown, System.Type.Missing);
            ApplyStyle2(Worksheet.Range[Worksheet.Cells[RowIndex, 1], Worksheet.Cells[RowIndex, ColumnsCount]]);
            Worksheet.Range[Worksheet.Cells[RowIndex, 6], Worksheet.Cells[RowIndex, 7]].Merge();
            SetGrayInterior(Worksheet.Range[Worksheet.Cells[RowIndex, 6], Worksheet.Cells[RowIndex, 7]]);
            Worksheet.Range[Worksheet.Cells[RowIndex, 9], Worksheet.Cells[RowIndex, 10]].Merge();
            SetGrayInterior(Worksheet.Range[Worksheet.Cells[RowIndex, 9], Worksheet.Cells[RowIndex, 10]]);

            Worksheet.Cells[RowIndex, 1].Value = CompleteName;
            Worksheet.Cells[RowIndex, 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
            Worksheet.Cells[RowIndex, 11].Value = Norm;
            Worksheet.Cells[RowIndex, 9].Value = Count;
            if (planNorm != 0D) Worksheet.Cells[RowIndex, 2].Value = planNorm;
        }
        /// <summary>
        /// Добавление в отчет партии.
        /// </summary>
        /// <param name="Worksheet">Рабочий лист.</param>
        /// <param name="RowIndex">Индекс строки.</param>
        /// <param name="PartyName">Название партии.</param>
        private static void AddParty(Excel.Worksheet Worksheet, int RowIndex, string PartyName)
        {
            Worksheet.Rows[RowIndex].Insert(Excel.XlInsertShiftDirection.xlShiftDown, System.Type.Missing);
            Worksheet.Range[Worksheet.Cells[RowIndex, 6], Worksheet.Cells[RowIndex, 7]].Merge();
            Worksheet.Range[Worksheet.Cells[RowIndex, 9], Worksheet.Cells[RowIndex, 10]].Merge();

            ApplyStyle2(Worksheet.Range[Worksheet.Cells[RowIndex, 1], Worksheet.Cells[RowIndex, ColumnsCount]]);
            SetGrayInterior(Worksheet.Range[Worksheet.Cells[RowIndex, 3], Worksheet.Cells[RowIndex, 3]]);
            SetGrayInterior(Worksheet.Range[Worksheet.Cells[RowIndex, 6], Worksheet.Cells[RowIndex, 7]]);
            SetGrayInterior(Worksheet.Range[Worksheet.Cells[RowIndex, 9], Worksheet.Cells[RowIndex, 10]]);

            Worksheet.Cells[RowIndex, 1].Value = PartyName;
            Worksheet.Cells[RowIndex, 2].Value = Worksheet.Cells[RowIndex - 1, 2].Value;
            Worksheet.Cells[RowIndex, 11].FormulaR1C1 = "=RC[-9]*RC[-2]";
            Worksheet.Cells[RowIndex, 9].Value = GetInt32(Worksheet.Cells[RowIndex, 9].Value) + 1;
        }
        /// <summary>
        /// Добавление в отчет типа прибора.
        /// </summary>
        /// <param name="Worksheet">Рабочий лист.</param>
        /// <param name="RowIndex">Индекс строки.</param>
        /// <param name="DeviceType">Название типа прибора.</param>
        /// <param name="planNorm">Плановая норма калибровки прибора.</param>
        private static void AddDeviceType(Excel.Worksheet Worksheet, int RowIndex, string DeviceType, double planNorm)
        {
            Worksheet.Rows[RowIndex].Insert(Excel.XlInsertShiftDirection.xlShiftDown, System.Type.Missing);
            Worksheet.Range[Worksheet.Cells[RowIndex, 6], Worksheet.Cells[RowIndex, 7]].Merge();
            Worksheet.Range[Worksheet.Cells[RowIndex, 9], Worksheet.Cells[RowIndex, 10]].Merge();

            ApplyStyle3(Worksheet.Range[Worksheet.Cells[RowIndex, 1], Worksheet.Cells[RowIndex, ColumnsCount]]);

            Worksheet.Cells[RowIndex, 1].Value = DeviceType;
            Worksheet.Cells[RowIndex, 2].Value = planNorm;
            Worksheet.Cells[RowIndex, 11].FormulaR1C1 = "=RC[-9]*RC[-2]";
        }
        /// <summary>
        /// Подсчет итоговой трудоемкости.
        /// </summary>
        /// <param name="Worksheet">Рабочий лист.</param>
        private void TotalCount(Excel.Worksheet Worksheet)
        {
            string SummaryFormula = "=СУММ(";
            // Происк промежуточных итогов
            for (int i = StartPosition; i < RowsCount; i++)
            { if (GetString(Worksheet.Cells[i, 1].Value).IndexOf("Всего") >= 0) SummaryFormula = SummaryFormula != "=СУММ(" ? SummaryFormula + ";K" + i : SummaryFormula + "K" + i; }
            SummaryFormula = SummaryFormula + ")";
            // Обновление формулы итоговой трудоемкости в отчете
            Worksheet.Cells[RowsCount, 11].FormulaLocal = SummaryFormula;
        }
        /// <summary>
        /// Цветовая разметка отчета на основании результатов выполнения плана
        /// </summary>
        /// <param name="Worksheet">Рабочий лист.</param>
        private void SetColorMarks(Excel.Worksheet Worksheet)
        {
            string CurrentDeviceType = "Новые приборы";
            for (int i = StartPosition; i < RowsCount; i++)
            {
                if (GetString(Worksheet.Cells[i, 1].Value) == "Приборы заказчиков")
                { CurrentDeviceType = "Приборы заказчиков"; }

                if ((GetString(Worksheet.Cells[i, 1].Value) != "Приборы заказчиков") && (GetString(Worksheet.Cells[i, 1].Value) != "Новые приборы") && (CurrentDeviceType != ""))
                {
                    Excel.Range CurrentRow;
                    switch (CurrentDeviceType)
                    {
                        case "Новые приборы":
                            CurrentRow = Worksheet.Range[Worksheet.Cells[i, 1], Worksheet.Cells[i, ColumnsCount]];
                            if ((Worksheet.Cells[i, 1].Font.Bold == false) && (Worksheet.Cells[i, 1].Value != "Наименование прибора"))
                            {
                                if (GetInt32(Worksheet.Cells[i, 6].Value) < GetInt32(Worksheet.Cells[i, 9].Value)) SetYellowInterior(CurrentRow);
                                if (GetInt32(Worksheet.Cells[i, 6].Value) > GetInt32(Worksheet.Cells[i, 9].Value)) SetRedInterior(CurrentRow);
                                if ((GetInt32(Worksheet.Cells[i, 6].Value) == GetInt32(Worksheet.Cells[i, 9].Value)) && (GetInt32(Worksheet.Cells[i, 6].Value) != 0)) SetGreenInterior(CurrentRow);
                            }
                            break;
                        case "Приборы заказчиков":
                            CurrentRow = Worksheet.Range[Worksheet.Cells[i, 1], Worksheet.Cells[i, ColumnsCount]];
                            if (Worksheet.Cells[i, 1].MergeCells == false)
                            {
                                int PlanRepair = GetInt32(Worksheet.Cells[i, 6].Value);
                                int PlanWarehouse = GetInt32(Worksheet.Cells[i, 7].Value);
                                int FactRepair = GetInt32(Worksheet.Cells[i, 9].Value);
                                int FactWarehouse = GetInt32(Worksheet.Cells[i, 10].Value);
                                // если все запланированные работы выполнены
                                if ((PlanRepair == FactRepair) && (PlanWarehouse == FactWarehouse) && ((PlanRepair != 0) || (PlanWarehouse != 0)))
                                { SetGreenInterior(CurrentRow); }
                                // если хотя бы одна запланированная работа не выполнена
                                if (((FactRepair < PlanRepair) && (FactWarehouse == 0)) || (FactWarehouse < PlanWarehouse) && (FactRepair == 0))
                                { SetRedInterior(CurrentRow); }
                                // если выполнены незапланированные работы
                                if ((PlanRepair < FactRepair) || (PlanWarehouse < FactWarehouse))
                                { SetYellowInterior(CurrentRow); }
                            }
                            break;
                    }
                }
            }
        }
        /// <summary>
        /// Создание заголовка отчета
        /// </summary>
        /// <param name="Worksheet">Рабочий лист.</param>
        private void SetHeader(Excel.Worksheet Worksheet)
        {
            Worksheet.Cells[HeaderPosition, 1].Value = HeaderText;
            Worksheet.Range[Worksheet.Cells[HeaderPosition, 1], Worksheet.Cells[HeaderPosition, ColumnsCount]].Merge();
        }
        /// <summary>
        /// Создание легенды.
        /// </summary>
        /// <param name="Worksheet">Рабочий лист.</param>
        private void SetLegend(Excel.Worksheet Worksheet)
        {
            Worksheet.Range[Worksheet.Cells[1, 1], Worksheet.Cells[3, ColumnsCount]].Clear();
            Worksheet.Range[Worksheet.Cells[RowsCount + 1, 1], Worksheet.Cells[RowsCount + 5, ColumnsCount]].Clear();
            SetGreenInterior(Worksheet.Cells[1, 9]);
            Worksheet.Cells[1, 10].Value = " - план выполнен";
            SetRedInterior(Worksheet.Cells[2, 9]);
            Worksheet.Cells[2, 10].Value = " - план не выполнен";
            SetYellowInterior(Worksheet.Cells[3, 9]);
            Worksheet.Cells[3, 10].Value = " - не запланировано";
        }
        /// <summary>
        /// Стиль 1 (крайние границы: средние; внутренние границы: тонкие; заливка: серый (15%), выравнивание: по центру; шрифт: полужирный курсив).
        /// </summary>
        /// <param name="Range">Рабочий диапазон ячеек.</param>
        private static void ApplyStyle1(Excel.Range Range)
        {
            SetDiagonalBorders(Range, Excel.XlLineStyle.xlLineStyleNone);
            SetEdgeBorders(Range, Excel.XlLineStyle.xlContinuous, 0, 0, Excel.XlBorderWeight.xlMedium);
            SetInsideBorders(Range, Excel.XlLineStyle.xlContinuous, 0, 0, Excel.XlBorderWeight.xlThin);
            SetInterior(Range, Excel.XlPattern.xlPatternSolid, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlThemeColor.xlThemeColorDark1, -0.149998474074526, 0);
            SetFont(Range, "Verdana", 10, true, true, false, false, false, false, false, Excel.XlUnderlineStyle.xlUnderlineStyleNone, 0, Excel.XlThemeFont.xlThemeFontNone);
            SetAlignment(Range, Excel.XlHAlign.xlHAlignCenter, Excel.XlHAlign.xlHAlignCenter, 0, false, 0, false);
        }
        /// <summary>
        /// Стиль 2 (крайние границы: тонкие; внутренние границы: тонкие; заливка: нет; выравнивание: по центру; шрифт: обычный).
        /// </summary>
        /// <param name="Range">Рабочий диапазон ячеек.</param>
        private static void ApplyStyle2(Excel.Range Range)
        {
            SetDiagonalBorders(Range, Excel.XlLineStyle.xlLineStyleNone);
            SetEdgeBorders(Range, Excel.XlLineStyle.xlContinuous, 0, 0, Excel.XlBorderWeight.xlThin);
            SetInsideBorders(Range, Excel.XlLineStyle.xlContinuous, 0, 0, Excel.XlBorderWeight.xlThin);
            SetInterior(Range, Excel.XlPattern.xlPatternSolid, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlThemeColor.xlThemeColorDark1, 0, 0);
            SetFont(Range, "Verdana", 10, false, false, false, false, false, false, false, Excel.XlUnderlineStyle.xlUnderlineStyleNone, 0, Excel.XlThemeFont.xlThemeFontNone);
            SetAlignment(Range, Excel.XlHAlign.xlHAlignCenter, Excel.XlHAlign.xlHAlignCenter, 0, false, 0, false);
        }
        /// <summary>
        /// Стиль 3 (внешние границы: тонкие; внутренние границы: тонкие; заливка: серый (15%); выравнивание: по центру; шрифт: полужирный.
        /// </summary>
        /// <param name="Range">Рабочий диапазон ячеек.</param>
        private static void ApplyStyle3(Excel.Range Range)
        {
            SetDiagonalBorders(Range, Excel.XlLineStyle.xlLineStyleNone);
            SetEdgeBorders(Range, Excel.XlLineStyle.xlContinuous, 0, 0, Excel.XlBorderWeight.xlThin);
            SetInsideBorders(Range, Excel.XlLineStyle.xlContinuous, 0, 0, Excel.XlBorderWeight.xlThin);
            SetInterior(Range, Excel.XlPattern.xlPatternSolid, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlThemeColor.xlThemeColorDark1, -0.149998474074526, 0);
            SetFont(Range, "Verdana", 10, false, true, false, false, false, false, false, Excel.XlUnderlineStyle.xlUnderlineStyleNone, 0, Excel.XlThemeFont.xlThemeFontNone);
            SetAlignment(Range, Excel.XlHAlign.xlHAlignCenter, Excel.XlHAlign.xlHAlignCenter, 0, false, 0, false);
        }
        /// <summary>
        /// Стиль 4 (внешние границы: тонкие; внутренние границы: тонкие; заливка: серый (5%), выравнивание: по центру; шрифт: обычный.
        /// </summary>
        /// <param name="Range">Рабочий диапазон ячеек.</param>
        private static void ApplyStyle4(Excel.Range Range)
        {
            SetDiagonalBorders(Range, Excel.XlLineStyle.xlLineStyleNone);
            SetEdgeBorders(Range, Excel.XlLineStyle.xlContinuous, 0, 0, Excel.XlBorderWeight.xlThin);
            SetInsideBorders(Range, Excel.XlLineStyle.xlContinuous, 0, 0, Excel.XlBorderWeight.xlThin);
            SetInterior(Range, Excel.XlPattern.xlPatternSolid, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlThemeColor.xlThemeColorDark1, -4.99893185216834E-02, 0);
            SetFont(Range, "Verdana", 10, false, false, false, false, false, false, false, Excel.XlUnderlineStyle.xlUnderlineStyleNone, 0, Excel.XlThemeFont.xlThemeFontNone);
            SetAlignment(Range, Excel.XlHAlign.xlHAlignCenter, Excel.XlHAlign.xlHAlignCenter, 0, false, 0, false);
        }
        /// <summary>
        /// Стиль 5 (внешние границы: жирные; внутренние границы: тонкие; заливка: серый (25%), выравнивание: по центру; шрифт: полужирный, 11.
        /// </summary>
        /// <param name="Range">Рабочий диапазон ячеек.</param>
        private static void ApplyStyle5(Excel.Range Range)
        {
            SetDiagonalBorders(Range, Excel.XlLineStyle.xlLineStyleNone);
            SetEdgeBorders(Range, Excel.XlLineStyle.xlContinuous, 0, 0, Excel.XlBorderWeight.xlThick);
            SetInsideBorders(Range, Excel.XlLineStyle.xlContinuous, 0, 0, Excel.XlBorderWeight.xlThin);
            SetInterior(Range, Excel.XlPattern.xlPatternSolid, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlThemeColor.xlThemeColorDark1, -0.249977111117893, 0);
            SetFont(Range, "Verdana", 11, false, true, false, false, false, false, false, Excel.XlUnderlineStyle.xlUnderlineStyleNone, 0, Excel.XlThemeFont.xlThemeFontNone);
            SetAlignment(Range, Excel.XlHAlign.xlHAlignCenter, Excel.XlHAlign.xlHAlignCenter, 0, false, 0, false);
        }
        /// <summary>
        /// Установка диагональных границ.
        /// </summary>
        /// <param name="Range">Рабочий диапазон ячеек.</param>
        /// <param name="LineStyle">Стиль границ.</param>
        private static void SetDiagonalBorders(Excel.Range Range, Excel.XlLineStyle LineStyle)
        {
            Range.Borders[Excel.XlBordersIndex.xlDiagonalDown].LineStyle = LineStyle;
            Range.Borders[Excel.XlBordersIndex.xlDiagonalUp].LineStyle = LineStyle;
        }
        /// <summary>
        /// Установка наружных границ.
        /// </summary>
        /// <param name="Range">Рабочий диапазон ячеек.</param>
        /// <param name="LineStyle">Стиль границ.</param>
        /// <param name="ColorIndex">Цвет.</param>
        /// <param name="TintAndShade">Оттенок.</param>
        /// <param name="Weight">Толщина границ.</param>
        private static void SetEdgeBorders(Excel.Range Range, Excel.XlLineStyle LineStyle, int ColorIndex, int TintAndShade, Excel.XlBorderWeight Weight)
        {
            Range.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = LineStyle;
            Range.Borders[Excel.XlBordersIndex.xlEdgeLeft].ColorIndex = ColorIndex;
            Range.Borders[Excel.XlBordersIndex.xlEdgeLeft].TintAndShade = TintAndShade;
            Range.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Weight;

            Range.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = LineStyle;
            Range.Borders[Excel.XlBordersIndex.xlEdgeTop].ColorIndex = ColorIndex;
            Range.Borders[Excel.XlBordersIndex.xlEdgeTop].TintAndShade = TintAndShade;
            Range.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Weight;

            Range.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = LineStyle;
            Range.Borders[Excel.XlBordersIndex.xlEdgeBottom].ColorIndex = ColorIndex;
            Range.Borders[Excel.XlBordersIndex.xlEdgeBottom].TintAndShade = TintAndShade;
            Range.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Weight;

            Range.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = LineStyle;
            Range.Borders[Excel.XlBordersIndex.xlEdgeRight].ColorIndex = ColorIndex;
            Range.Borders[Excel.XlBordersIndex.xlEdgeRight].TintAndShade = TintAndShade;
            Range.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Weight;
        }
        /// <summary>
        /// Установка внутренних границ.
        /// </summary>
        /// <param name="Range">Рабочий диапазон ячеек.</param>
        /// <param name="LineStyle">Стиль границ.</param>
        /// <param name="ColorIndex">Цвет.</param>
        /// <param name="TintAndShade">Оттенок.</param>
        /// <param name="Weight">Толщина границ.</param>
        private static void SetInsideBorders(Excel.Range Range, Excel.XlLineStyle LineStyle, int ColorIndex, int TintAndShade, Excel.XlBorderWeight Weight)
        {
            Range.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = LineStyle;
            Range.Borders[Excel.XlBordersIndex.xlInsideVertical].ColorIndex = ColorIndex;
            Range.Borders[Excel.XlBordersIndex.xlInsideVertical].TintAndShade = TintAndShade;
            Range.Borders[Excel.XlBordersIndex.xlInsideVertical].Weight = Weight;

            Range.Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = LineStyle;
            Range.Borders[Excel.XlBordersIndex.xlInsideHorizontal].ColorIndex = ColorIndex;
            Range.Borders[Excel.XlBordersIndex.xlInsideHorizontal].TintAndShade = TintAndShade;
            Range.Borders[Excel.XlBordersIndex.xlInsideHorizontal].Weight = Weight;
        }
        /// <summary>
        /// Установка внешнего вида.
        /// </summary>
        /// <param name="Range">Рабочий диапазон ячеек.</param>
        /// <param name="Pattern">Pattern.</param>
        /// <param name="PatternColorIndex">PatternColorIndex.</param>
        /// <param name="ThemeColor">ThemeColor.</param>
        /// <param name="TintAndShade">TintAndShade.</param>
        /// <param name="PatternTintAndShade">PatternTintAndShade.</param>
        private static void SetInterior(Excel.Range Range, Excel.XlPattern Pattern, Excel.XlColorIndex PatternColorIndex, Excel.XlThemeColor ThemeColor, double TintAndShade, int PatternTintAndShade)
        {
            Range.Interior.Pattern = Pattern;
            Range.Interior.PatternColorIndex = PatternColorIndex;
            Range.Interior.ThemeColor = ThemeColor;
            Range.Interior.TintAndShade = TintAndShade;
            Range.Interior.PatternTintAndShade = PatternTintAndShade;
        }
        /// <summary>
        /// Установка параметров текста.
        /// </summary>
        /// <param name="Range">Рабочий диапазон ячеек.</param>
        /// <param name="Name">Название шрифта.</param>
        /// <param name="Size">Размер.</param>
        /// <param name="Italic">Курсив.</param>
        /// <param name="Bold">Полужирный.</param>
        /// <param name="Strikethrough">Зачеркнутый.</param>
        /// <param name="Superscript">Верхний индекс.</param>
        /// <param name="Subscript">Нижний индекс.</param>
        /// <param name="OutlineFont">Контур.</param>
        /// <param name="Shadow">Тень.</param>
        /// <param name="Underline">Подчеркивание.</param>
        /// <param name="TintAndShade">Оттенки.</param>
        /// <param name="ThemeFont">Тема.</param>
        private static void SetFont(Excel.Range Range, string Name, int Size, bool Italic, bool Bold, bool Strikethrough, bool Superscript, bool Subscript, bool OutlineFont, bool Shadow, Excel.XlUnderlineStyle Underline,
            int TintAndShade, Excel.XlThemeFont ThemeFont)
        {
            Range.Font.Name = Name;
            Range.Font.Size = Size;
            Range.Font.Strikethrough = Strikethrough;
            Range.Font.Superscript = Superscript;
            Range.Font.Subscript = Subscript;
            Range.Font.OutlineFont = OutlineFont;
            Range.Font.Shadow = Shadow;
            Range.Font.Underline = Underline;
            Range.Font.TintAndShade = TintAndShade;
            Range.Font.ThemeFont = ThemeFont;
            Range.Font.Italic = Italic;
            Range.Font.Bold = Bold;
        }
        /// <summary>
        /// Установка параметров выравнивания для защищенных ячеек.
        /// </summary>
        /// <param name="Range">Рабочий диапазон ячеек.</param>
        /// <param name="HorizontalAlignment">Горизонтальное выравнивание.</param>
        /// <param name="VerticalAlignment">Вертикальное выравнивание.</param>
        /// <param name="Orientation">Ориентация текста.</param>
        /// <param name="AddIndent">Добавить отступ.</param>
        /// <param name="IndentLevel">Уровень отступа.</param>
        /// <param name="ShrinkToFit">Выровнять по размеру.</param>
        /// <param name="WrapText">Перенос текста по словам.</param>
        private static void SetAlignment(Excel.Range Range, Excel.XlHAlign HorizontalAlignment, Excel.XlHAlign VerticalAlignment, int Orientation, bool AddIndent, int IndentLevel, bool ShrinkToFit, bool WrapText = true)
        {
            Range.HorizontalAlignment = HorizontalAlignment;
            Range.VerticalAlignment = VerticalAlignment;
            Range.Orientation = Orientation;
            Range.AddIndent = AddIndent;
            Range.IndentLevel = IndentLevel;
            Range.ShrinkToFit = ShrinkToFit;
            Range.EntireRow.AutoFit();
            Range.WrapText = WrapText;
        }
        /// <summary>
        /// Установка параметров выравнивания для открытых ячеек.
        /// </summary>
        /// <param name="Range">Рабочий диапазон ячеек.</param>
        /// <param name="HorizontalAlignment">Горизонтальное выравнивание.</param>
        /// <param name="WrapText">Перенос текста по словам.</param>
        private static void SetAlignment(Excel.Range Range, Excel.XlHAlign HorizontalAlignment, bool WrapText)
        {
            Range.HorizontalAlignment = HorizontalAlignment;
            Range.WrapText = WrapText;
        }
        /// <summary>
        /// Установка серой заливки.
        /// </summary>
        /// <param name="Range">Рабочий диапазон ячеек.</param>
        private static void SetGrayInterior(Excel.Range Range)
        { SetInterior(Range, Excel.XlPattern.xlPatternSolid, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlThemeColor.xlThemeColorDark1, -0.149998474074526, 0); }
        /// <summary>
        /// Установка зеленой заливки.
        /// </summary>
        /// <param name="Range">Рабочий диапазон ячеек.</param>
        private static void SetGreenInterior(Excel.Range Range)
        { Range.Interior.Color = 5296274; }
        /// <summary>
        /// Установка красной заливки.
        /// </summary>
        /// <param name="Range">Рабочий диапазон ячеек.</param>
        private static void SetRedInterior(Excel.Range Range)
        { Range.Interior.Color = 255; }
        /// <summary>
        /// Установка желтой заливки.
        /// </summary>
        /// <param name="Range">Рабочий диапазон ячеек.</param>
        private static void SetYellowInterior(Excel.Range Range)
        { Range.Interior.Color = 65535; }
        /// <summary>
        /// Преобразование к типу string.
        /// </summary>
        /// <param name="Value">Значение.</param>
        private static string GetString(object Value)
        { return Value == null ? "" : Value.ToString(); }
        /// <summary>
        /// Преобразование к типу int.
        /// </summary>
        /// <param name="Value">Значение.</param>
        private static int GetInt32(object Value)
        { return Value == null ? 0 : Convert.ToInt32(Value); }
        /// <summary>
        /// Преобразование к типу double.
        /// </summary>
        /// <param name="Value">Значение.</param>
        private static double GetDouble(object Value)
        { return Value == null ? 0D : (double)Value; }
    }
    /// <summary>
    /// Строка отчета
    /// </summary>
    public class ReportItem
    {
        // Общие
        /// <summary>
        /// Документ основание
        /// </summary>
        CardData InitiatingDocument = null;
        /// <summary>
        /// Тип передачи.
        /// </summary>
        public string StatusOfTransfer = "";
        /// <summary>
        /// Дата окончания калибровки.
        /// </summary>
        DateTime OperationDate = new DateTime();
        /// <summary>
        /// Фактическое время калибровки.
        /// </summary>
        double factCalibrationTime = 0D;

        // Для приборов (принадлежащих СКБ ЭП и принадлежащих заказчикам)
        /// <summary>
        /// Карточка "Паспорт прибора".
        /// </summary>
        CardData DeviceCard = null;
        /// <summary>
        /// Заводской номер прибора.
        /// </summary>
        string deviceNumber = "";
        /// <summary>
        /// Год прибора.
        /// </summary>
        string deviceYear = "";
        /// <summary>
        /// Тип прибора.
        /// </summary>
        string deviceType = "";
        /// <summary>
        /// Партия.
        /// </summary>
        string party = "";
        /// <summary>
        /// Плановая норма калибровки прибора.
        /// </summary>
        double planCalibrationTime = 0D;

        // Только для комплектующих
        /// <summary>
        /// Название комплектующего.
        /// </summary>
        RowData Complete = null;
        /// <summary>
        /// Код СКБ комплектующего.
        /// </summary>
        string CodeSKB = "";
        /// <summary>
        /// Количество комплектующего.
        /// </summary>
        int count = 1;
        /// <summary>
        /// Наряд-заказ
        /// </summary>
        string WorkOrderDescription;
        /// <summary>
        /// Партия.
        /// </summary>
        string ParentPartyDescription;

        // Только для приборов, принадлежащих заказчикам
        /// <summary>
        /// Название клиента.
        /// </summary>
        string clientName = "";
        /// <summary>
        /// Тип гарантии.
        /// </summary>
        string warranty = "";
        /// <summary>
        /// Тип передачи.
        /// </summary>
        string typeOfTransfer = "";
        /// <summary>
        /// Состояние прибора.
        /// </summary>
        string deviceState = "";
        /// <summary>
        /// Тип сервисного обслуживания.
        /// </summary>
        string serviceType = "";
        /// <summary>
        /// Текущий этап сервисного обслуживания.
        /// </summary>
        string serviceStage;
        /// <summary>
        /// Номер заявки на сервисное обслуживание.
        /// </summary>
        string serviceNumber;
        /// <summary>
        /// Комментарий сервисного специалиста к заявке на сервисное обслуживание.
        /// </summary>
        string serviceComment;
        /// <summary>
        /// Плановая дата окончания сервисного обслуживания
        /// </summary>
        DateTime planEndServiceDate = new DateTime();
        /// <summary>
        /// Партия.
        /// </summary>
        public string Party
        { get { return party; } }
        /// <summary>
        /// Тип прибора.
        /// </summary>
        public string DeviceType
        { get { return deviceType; } }
        /// <summary>
        /// Плановая норма калибровки прибора.
        /// </summary>
        public double PlanCalibrationTime
        { get { return planCalibrationTime; } }
        /// <summary>
        /// Наименование комплектующих.
        /// </summary>
        public string CompleteName
        { get { return Complete == null ? "" : Complete.GetString("Name") + " " + CodeSKB; } }
        /// <summary>
        /// Код СКБ комплектующих.
        /// </summary>
        public string CompleteCode
        { get { return CodeSKB; } }
        /// <summary>
        /// Количество комплектующих.
        /// </summary>
        public int Count
        { get { return count; } }
        /// <summary>
        /// Документ-основание для изготовления комплектующих.
        /// </summary>
        public string CompleteDocument
        { get { return WorkOrderDescription == "" ? ParentPartyDescription : WorkOrderDescription; } }
        /// <summary>
        /// Фактическое время калибровки.
        /// </summary>
        public double FactCalibrationTime
        { get { return factCalibrationTime; } }
        /// <summary>
        /// Название прибора.
        /// </summary>
        public string DeviceName
        {
            get
            {
                if (deviceNumber == "Только комплектующие")
                { return deviceType + " (Только комплектующие)"; }
                else
                { return deviceType + " № " + deviceNumber + "/" + deviceYear; }
            }
        }
        /// <summary>
        /// Название клиента.
        /// </summary>
        public string ClientName
        { get { return clientName; } }
        /// <summary>
        /// Состояние прибора.
        /// </summary>
        public string DeviceState
        {
            get
            { return deviceState == "На калибровке (СО)" ? ServiceStage : deviceState; }
        }
        /// <summary>
        /// Вид сервиса.
        /// </summary>
        public string ServiceType
        { get { return serviceType; } }
        /// <summary>
        /// Плановая дата окончания сервисного обслуживания.
        /// </summary>
        public DateTime PlanEndServiceDate
        { get { return planEndServiceDate; } }
        /// <summary>
        /// Тип гарантии.
        /// </summary>
        public string Warranty
        { get { return warranty; } }
        /// <summary>
        /// Тип передачи.
        /// </summary>
        public string TypeOfTransfer
        { get { return typeOfTransfer; } }
        /// <summary>
        /// Текущий этап серивсного обслуживания.
        /// </summary>
        string ServiceStage
        {
            get
            {
                switch (serviceStage)
                {
                    case "0":
                        return "На первичной калибровке";
                    case "1":
                        return "На согласовании в сбыте";
                    case "2":
                        return "В ремонте";
                    case "3":
                        return "На повторной калибровке";
                    case "4":
                        return "Отказ от ремонта";
                    case "5":
                        return "Завершено";
                    case "6":
                        return "На диагностике";
                    default:
                        return serviceStage;
                }
            }
        }
        /// <summary>
        /// Номер заявки на сервисное обслуживание.
        /// </summary>
        public string ServiceNumber
        { get { return serviceNumber; } }
        /// <summary>
        /// Комментарий сервисного специалиста к заявке на сервисное обслуживание.
        /// </summary>
        public string ServiceComment
        { get { return serviceComment; } }
        /// <summary>
        /// Инициализация строки отчета, соответствующей прибору СКБ ЭП.
        /// </summary>
        /// <param name="Session">Пользовательская сессия.</param>
        /// <param name="Document">Документ-основание.</param>
        /// <param name="DeviceCard">Карточка прибора.</param>
        /// <param name="StatusOfTransfer">Статус передачи.</param>
        /// <param name="FactCalibrationTime">Фактическая трудоемкость калибровки.</param>
        public ReportItem(UserSession Session, CardData Document, CardData DeviceCard, string StatusOfTransfer, double FactCalibrationTime)
        {
            CardData UniversalDictionary = Session.CardManager.GetDictionaryData(RefUniversal.ID);
            SectionData ActPropertiesSection = Document.Sections[CardOrd.Properties.ID];
            SectionData DevicePropertiesSection = DeviceCard.Sections[CardOrd.Properties.ID];

            RowData AcceptDate = ActPropertiesSection.FindRow("@Name = 'Дата принятия'");
            RowData DeviceNumberRow = DevicePropertiesSection.FindRow("@Name = 'Заводской номер прибора'");
            RowData DeviceYearRow = DevicePropertiesSection.FindRow("@Name = '/Год прибора'");
            RowData DeviceTypeRow = DevicePropertiesSection.FindRow("@Name = 'Прибор'");
            RowData DevicePartyRow = DevicePropertiesSection.FindRow("@Name = '№ партии'");

            // Заполняем общие переменные
            this.InitiatingDocument = Document;
            this.StatusOfTransfer = StatusOfTransfer;
            this.OperationDate = (DateTime)AcceptDate.GetDateTime(CardOrd.Properties.Value);
            this.factCalibrationTime = FactCalibrationTime;

            // Заполняем переменные для приборов, принадлежащих СКБ ЭП
            this.DeviceCard = DeviceCard;
            this.deviceNumber = DeviceNumberRow.GetString(CardOrd.Properties.Value);
            this.deviceYear = DeviceYearRow.GetString(CardOrd.Properties.Value);
            this.deviceType = UniversalDictionary.GetItemName(DeviceTypeRow.GetString(CardOrd.Properties.Value));
            this.party = UniversalDictionary.GetItemName(DevicePartyRow.GetString(CardOrd.Properties.Value));
            object PlanCalibrationTimeValue = UniversalDictionary.GetItemPropertyValue(new Guid(DeviceTypeRow.GetString(CardOrd.Properties.Value)), "Время калибровки (ч.)");
            this.planCalibrationTime = PlanCalibrationTimeValue == null ? 0D : (double)PlanCalibrationTimeValue;
        }
        /// <summary>
        /// Инициализация строки отчета, соответствующей комплектующему СКБ ЭП.
        /// </summary>
        /// <param name="Session">Пользовательская сессия.</param>
        /// <param name="Document">Документ-основание.</param>
        /// <param name="Complete">Запись комплектующего в справочнике.</param>
        /// <param name="StatusOfTransfer">Статус передачи.</param>
        /// <param name="Count">Статус передачи.</param>
        /// <param name="FactCalibrationTime">Фактическая трудоемкость калибровки.</param>
        public ReportItem(UserSession Session, CardData Document, RowData Complete, string StatusOfTransfer, int Count, double FactCalibrationTime)
        {
            CardData UniversalDictionary = Session.CardManager.GetDictionaryData(RefUniversal.ID);
            SectionData ActPropertiesSection = Document.Sections[CardOrd.Properties.ID];
            RowData AcceptDate = ActPropertiesSection.FindRow("@Name = 'Дата принятия'");
            RowData WorkOrder = ActPropertiesSection.FindRow("@Name = 'Наряд-заказ'");
            RowData ParentParty = ActPropertiesSection.FindRow("@Name = 'Партия'");

            // Заполняем общие переменные
            this.InitiatingDocument = Document;
            this.StatusOfTransfer = StatusOfTransfer;
            this.OperationDate = (DateTime)AcceptDate.GetDateTime(CardOrd.Properties.Value);
            this.factCalibrationTime = FactCalibrationTime;

            // Заполняем переменные только для комплектующих
            this.Complete = Complete;
            string CodeID = UniversalDictionary.GetItemPropertyValue(Complete.Id, "Код СКБ") == null ? "" : UniversalDictionary.GetItemPropertyValue(Complete.Id, "Код СКБ").ToString();
            string CodeName = CodeID == "" ? "" : UniversalDictionary.GetItemName(CodeID);
            this.CodeSKB = CodeName;
            this.count = Count;
            WorkOrderDescription = WorkOrder.GetString(CardOrd.Properties.Value) == null ? "" : WorkOrder.GetString(CardOrd.Properties.DisplayValue).ToString();
            ParentPartyDescription = ParentParty.GetString(CardOrd.Properties.Value) == null ? "" : ParentParty.GetString(CardOrd.Properties.DisplayValue).ToString();

        }
        /// <summary>
        /// Инициализация строки отчета, соответствующей прибору заказчика.
        /// </summary>
        /// <param name="Session">Пользовательская сессия.</param>
        /// <param name="Document">Документ-основание.</param>
        /// <param name="DeviceCard">Карточка прибора.</param>
        /// <param name="StatusOfTransfer">Статус передачи.</param>
        /// <param name="FactCalibrationTime">Фактическая трудоемкость калибровки.</param>
        /// <param name="TypeOfTransfer">Тип передачи.</param>
        public ReportItem(UserSession Session, CardData Document, CardData DeviceCard, string StatusOfTransfer, double FactCalibrationTime, string TypeOfTransfer)
        {
            CardData UniversalDictionary = Session.CardManager.GetDictionaryData(RefUniversal.ID);
            CardData PartnersDictionary = Session.CardManager.GetDictionaryData(new Guid("{65FF9382-17DC-4E9F-8E93-84D6D3D8FE8C}"));
            RowData ServiceCardMainInfoSection = Document.Sections[RefServiceCard.MainInfo.ID].FirstRow;
            RowData ServiceCardCalibrationSection = Document.Sections[RefServiceCard.Calibration.ID].FirstRow;

            // Заполняем общие переменные
            this.InitiatingDocument = Document;
            this.StatusOfTransfer = StatusOfTransfer;

            this.OperationDate = TypeOfTransfer == TypeOfTransferValues.ToWarehouse ? (DateTime)ServiceCardMainInfoSection.GetDateTime(RefServiceCard.MainInfo.DateEndFact) :
                (DateTime)ServiceCardCalibrationSection.GetDateTime(RefServiceCard.Calibration.CalibDateEnd);
            this.factCalibrationTime = FactCalibrationTime;

            // Для приборов (принадлежащих СКБ ЭП и принадлежащих заказчикам)
            if (DeviceCard != null)
            {
                SectionData DevicePropertiesSection = DeviceCard.Sections[CardOrd.Properties.ID];

                RowData DeviceNumberRow = DevicePropertiesSection.FindRow("@Name = 'Заводской номер прибора'");
                RowData DeviceYearRow = DevicePropertiesSection.FindRow("@Name = '/Год прибора'");
                RowData DeviceTypeRow = DevicePropertiesSection.FindRow("@Name = 'Прибор'");
                RowData DevicePartyRow = DevicePropertiesSection.FindRow("@Name = '№ партии'");
                RowData DeviceStateRow = DevicePropertiesSection.FindRow("@Name = 'Состояние прибора'");

                this.DeviceCard = DeviceCard;
                this.deviceNumber = DeviceNumberRow.GetString(CardOrd.Properties.Value);
                this.deviceYear = DeviceYearRow.GetString(CardOrd.Properties.Value);
                this.deviceType = UniversalDictionary.GetItemName(DeviceTypeRow.GetString(CardOrd.Properties.Value));
                this.party = UniversalDictionary.GetItemName(DevicePartyRow.GetString(CardOrd.Properties.Value));
                this.planCalibrationTime = (double)UniversalDictionary.GetItemPropertyValue(new Guid(DeviceTypeRow.GetString(CardOrd.Properties.Value)), "Время калибровки (ч.)");
                this.deviceState = DeviceStateRow.GetString(CardOrd.Properties.DisplayValue); // Только для приборов, принадлежащих заказчикам
            }
            else
            {
                this.deviceNumber = "Только комплектующие";
                this.deviceYear = "";
                this.deviceType = UniversalDictionary.GetItemName(ServiceCardMainInfoSection.GetGuid(RefServiceCard.MainInfo.DeviceType));
                this.party = "";
                this.planCalibrationTime = 0D;
            }

            // Только для приборов, принадлежащих заказчикам
            this.clientName = PartnersDictionary.Sections[new Guid("{C78ABDED-DB1C-4217-AE0D-51A400546923}")].GetRow(new Guid(ServiceCardMainInfoSection.GetString(RefServiceCard.MainInfo.Client))).GetString("Name");
            this.warranty = ServiceCardCalibrationSection.GetBoolean(RefServiceCard.Calibration.WarrantyService) == true ? TypeOfWarranty.Warranty : TypeOfWarranty.NonWarranty;
            this.typeOfTransfer = TypeOfTransfer;
            this.serviceType = ServiceCardCalibrationSection.GetString(RefServiceCard.Calibration.ReqTypeService);
            this.serviceStage = ServiceCardMainInfoSection.GetInt32(RefServiceCard.MainInfo.Status).ToString();
            this.planEndServiceDate = (DateTime)ServiceCardMainInfoSection.GetDateTime(RefServiceCard.MainInfo.DateEndPlan);

            CardData ServiceLinks = Session.CardManager.GetCardData(ServiceCardMainInfoSection.GetGuid(RefServiceCard.MainInfo.Links).ToGuid());
            CardData ApplicationCard = null;
            foreach (RowData LinksRow in ServiceLinks.Sections[new Guid("{568CE0A6-7096-43CC-9800-E0B268B14CC4}")].Rows)
            {
                if (LinksRow.GetGuid("CardType") == RefApplicationCard.ID)
                    ApplicationCard = Session.CardManager.GetCardData(LinksRow.GetGuid("Card").ToGuid());
            }
            serviceNumber = ApplicationCard == null ? "" : ApplicationCard.Sections[new Guid("{250D2733-B164-453D-A440-057352DD5D74}")].FirstRow.GetString("Number");
            serviceComment = ApplicationCard == null ? "" : ApplicationCard.Sections[RefApplicationCard.MainInfo.ID].FirstRow.GetString(RefApplicationCard.MainInfo.SalesComment);
        }
    }
    /// <summary>
    /// Статусы передачи из калибровочной лаборатории.
    /// </summary>
    public static class StatusOfTransferValues
    {
        /// <summary>
        /// Новые приборы.
        /// </summary>
        public const String NewDevices = "Новые приборы";
        /// <summary>
        /// Новые комплектующие.
        /// </summary>
        public const String NewComplete = "Новые комплектующие";
        /// <summary>
        /// Приборы и комплектующие после повторной калибровки.
        /// </summary>
        public const String AfterRecalibration = "Приборы и комплектующие после повторной калибровки";
        /// <summary>
        /// Приборы заказчиков.
        /// </summary>
        public const String ClientsDevices = "Приборы заказчиков";
    }
    /// <summary>
    /// Типы передачи из калибровочной лаборатории.
    /// </summary>
    public static class TypeOfTransferValues
    {
        /// <summary>
        /// На склад.
        /// </summary>
        public const String ToWarehouse = "На склад";
        /// <summary>
        /// В ремонт.
        /// </summary>
        public const String ToRepair = "В ремонт";
    }
    /// <summary>
    /// Типы гарантии.
    /// </summary>
    public static class TypeOfWarranty
    {
        /// <summary>
        /// На склад.
        /// </summary>
        public const String Warranty = "Гарантийные";
        /// <summary>
        /// В ремонт.
        /// </summary>
        public const String NonWarranty = "Негарантийные";
    }
}
