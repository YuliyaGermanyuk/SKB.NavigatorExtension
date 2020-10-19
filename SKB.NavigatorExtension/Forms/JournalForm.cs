using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using DocsVision.Platform.ObjectModel;
using DocsVision.Platform.ObjectManager;
using DocsVision.Platform.ObjectManager.SearchModel;
using DocsVision.Platform.ObjectManager.Metadata;
using DocsVision.TakeOffice.Cards.Constants;
using DocsVision.BackOffice.CardLib;
using DocsVision.BackOffice.ObjectModel;
using DocsVision.BackOffice.ObjectModel.Services;

using SKB.Base;
using RKIT.MyMessageBox;

using DevExpress.XtraEditors;
using DevExpress.Utils;

namespace SKB.NavigatorExtension.Forms
{
    /// <summary>
    /// Форма заполнения журнала условий калибровки
    /// </summary>
    public partial class JournalForm : DevExpress.XtraEditors.XtraForm
    {
        /// <summary>
        /// Проверено.
        /// </summary>
        bool verify = false;
        /// <summary>
        /// Минимально допустимая температура.
        /// </summary>
        const int MinTemperature = 15;
        /// <summary>
        /// Максимально допустимая температура.
        /// </summary>
        const int MaxTemperature = 25;
        /// <summary>
        /// Минимально допустимая влажность.
        /// </summary>
        const int MinHumidity = 30;
        /// <summary>
        /// Максимально допустимая влажность.
        /// </summary>
        const int MaxHumidity = 80;
        /// <summary>
        /// Минимально допустимое атмосферное давление.
        /// </summary>
        const int MinPressure = 1;
        /// <summary>
        /// Максимально допустимое атмосферное давление.
        /// </summary>
        const int MaxPressure = 1500;
        /// <summary>
        /// Пользовательская сессия DV.
        /// </summary>
        UserSession Session;
        /// <summary>
        /// Объектный контекст.
        /// </summary>
        ObjectContext Context;
        /// <summary>
        /// Номер кабинета.
        /// </summary>
        Int32 CabinetNumber;
        /// <summary>
        /// Карточка строки справочника.
        /// </summary>
        BaseUniversalItemCard itemCard;
        /// <summary>
        /// Сотрудник.
        /// </summary>
        StaffEmployee staffEmployee;
        /// <summary>
        /// Конструктор формы заполнения журнала условий калибровки.
        /// </summary>
        /// <param name="Session">Пользовательская сессия DV.</param>
        /// <param name="Context">Объектный контекст.</param>
        /// <param name="JournalItemType">Тип справочника.</param>
        /// <param name="CabinetNumber">Номер кабинета.</param>
        public JournalForm(UserSession Session, ObjectContext Context, BaseUniversalItemType JournalItemType, Int32 CabinetNumber)
        {
            InitializeComponent();
            this.Session = Session;
            this.Context = Context;
            IBaseUniversalService baseUniversalService = Context.GetService<IBaseUniversalService>();
            this.CabinetNumber = CabinetNumber;

            staffEmployee = Context.GetCurrentEmployee();
            BaseUniversalItem NewItem = baseUniversalService.AddNewItem(JournalItemType);
            NewItem.Name = "Каб. №" + CabinetNumber + ". Условия на " + DateTime.Today.ToShortDateString();
            this.Text = "Каб. №" + (CabinetNumber == 237 ? 226 : 228) + ". Условия на " + DateTime.Today.ToShortDateString();
            itemCard = baseUniversalService.OpenOrCreateItemCard(NewItem);
            NewItem.ItemCard = itemCard;
            Context.AcceptChanges();
            this.Date.DateTime = DateTime.Today;
            this.Employee.Text = staffEmployee.DisplayString;
        }
        /// <summary>
        /// Конструктор формы заполнения журнала условий калибровки.
        /// </summary>
        /// <param name="Session">Пользовательская сессия DV.</param>
        /// <param name="Context">Объектный контекст.</param>
        /// <param name="JournalItemType">Тип справочника.</param>
        /// <param name="CurrentItem">Текущая строка справочника.</param>
        /// <param name="CabinetNumber">Номер кабинета.</param>
        public JournalForm(UserSession Session, ObjectContext Context, BaseUniversalItemType JournalItemType, BaseUniversalItem CurrentItem, Int32 CabinetNumber)
        {
            InitializeComponent();
            this.Session = Session;
            this.Context = Context;
            IBaseUniversalService baseUniversalService = Context.GetService<IBaseUniversalService>();

            staffEmployee = Context.GetCurrentEmployee();
            itemCard = baseUniversalService.OpenOrCreateItemCard(CurrentItem);
            CardData itemCardData = Session.CardManager.GetCardData(Context.GetObjectRef<BaseUniversalItemCard>(itemCard).Id);
            SectionData CalibrationConditionsSection = itemCardData.Sections[itemCardData.Type.Sections["CalibrationConditions"].Id];
            RowData CalibrationConditionsRow = CalibrationConditionsSection.FirstRow;

            this.Text = "Каб. №" + (CabinetNumber == 237 ? 226 : 228) + ". Условия на " + DateTime.Today.ToShortDateString();
            this.Date.DateTime = (DateTime?)CalibrationConditionsRow.GetDateTime("Date") ?? DateTime.Today;
            this.Employee.Text = CalibrationConditionsRow.GetString("Employee") != null ? Context.GetEmployeeDisplay(new Guid(CalibrationConditionsRow.GetString("Employee"))) : staffEmployee.DisplayString;
            this.Temperature.Value = (decimal?)CalibrationConditionsRow.GetDecimal("Temperature") ?? 0;
            this.Humidity.Value = (decimal?)CalibrationConditionsRow.GetDecimal("Humidity") ?? 0;
            this.Pressure.Value = (decimal?)CalibrationConditionsRow.GetDecimal("Pressure") ?? 0;
        }

        private void dateEdit1_EditValueChanged(object sender, EventArgs e)
        {

        }

        private void OKButton_Click(object sender, EventArgs e)
        {
            decimal TemperatureValue = this.Temperature.Value;
            decimal HumidityValue = this.Humidity.Value;
            decimal PressureValue = this.Pressure.Value;
            bool Error = false;
            if ((TemperatureValue < MinTemperature) || (TemperatureValue > MaxTemperature))
            {
                ShowError(this.Temperature, "Введите значение от " + MinTemperature + " до " + MaxTemperature + ".");
                Error = true;
            }
            if ((HumidityValue < MinHumidity) || (HumidityValue > MaxHumidity))
            {
                ShowError(this.Humidity, "Введите значение от " + MinHumidity + " до " + MaxHumidity + ".");
                Error = true;
            }
            if ((PressureValue < MinPressure) || (PressureValue > MaxPressure))
            {
                ShowError(this.Pressure, "Введите значение от " + MinPressure + " до " + MaxPressure + ".");
                Error = true;
            }
            if (!Error)
            {
                CardData itemCardData = Session.CardManager.GetCardData(Context.GetObjectRef<BaseUniversalItemCard>(itemCard).Id);
                SectionData CalibrationConditionsSection = itemCardData.Sections[itemCardData.Type.Sections["CalibrationConditions"].Id];

                RowData NewCalibrationConditions = CalibrationConditionsSection.FirstRow == null ? CalibrationConditionsSection.Rows.AddNew() : CalibrationConditionsSection.FirstRow;
                NewCalibrationConditions.SetDateTime("Date", (DateTime)this.Date.EditValue);
                NewCalibrationConditions.SetGuid("Employee", Context.GetObjectRef<StaffEmployee>(staffEmployee).Id);
                NewCalibrationConditions.SetDecimal("Temperature", this.Temperature.Value);
                NewCalibrationConditions.SetDecimal("Humidity", this.Humidity.Value);
                NewCalibrationConditions.SetDecimal("Pressure", this.Pressure.Value);
                NewCalibrationConditions.SetInt32("CabinetNumber", this.CabinetNumber);
                verify = true;
                this.Close();
            }
        }

        private void JournalForm_Load(object sender, EventArgs e)
        {

        }

        private void ShowError(SpinEdit control, string errorText)
        {
            control.ErrorText = errorText;
            control.ErrorIconAlignment = ErrorIconAlignment.MiddleRight;
        }

        private void JournalForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            decimal TemperatureValue = this.Temperature.Value;
            decimal HumidityValue = this.Humidity.Value;
            decimal PressureValue = this.Pressure.Value;

            if ((TemperatureValue < MinTemperature) || (TemperatureValue > MaxTemperature))
                ShowError(this.Temperature, "Введите значение от " + MinTemperature + " до " + MaxTemperature + ".");

            if ((HumidityValue < MinHumidity) || (HumidityValue > MaxHumidity))
                ShowError(this.Humidity, "Введите значение от " + MinHumidity + " до " + MaxHumidity + ".");

            if ((PressureValue < MinPressure) || (PressureValue > MaxPressure))
                ShowError(this.Pressure, "Введите значение от " + MinPressure + " до " + MaxPressure + ".");

            if (!verify)
                e.Cancel = true;
        }
    }
}