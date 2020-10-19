using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using DocsVision.Platform.ObjectManager;
using DocsVision.Platform.ObjectModel;
using RKIT.MyMessageBox;

namespace SKB.NavigatorExtension.Forms
{
    /// <summary>
    /// Форма ввода параметров отчета
    /// </summary>
    public partial class ReportParameters : DevExpress.XtraEditors.XtraForm
    {
        /// <summary>
        /// Объектный контекст
        /// </summary>
        ObjectContext Context;
        /// <summary>
        /// Пользовательская сессия
        /// </summary>
        UserSession Session;
        /// <summary>
        /// Дата начала периода
        /// </summary>
        public DateTime StartDateValue
        { get {return ((DateTime)this.StartDate.EditValue).Date;}}
        /// <summary>
        /// Дата окончания периода
        /// </summary>
        public DateTime EndDateValue
        { get { return ((DateTime)this.EndDate.EditValue).Date; } }

        /// <summary>
        /// Инициализирует форму выбора параметров отчета.
        /// </summary>
        /// <param name="Session">Пользовательская сессия.</param>
        /// <param name="Context">Объектный контекст.</param>
        public ReportParameters(UserSession Session, ObjectContext Context)
        {
            InitializeComponent();
            this.SetNavigatorSkin();

            this.Context = Context;
            this.Session = Session;
        }
        /// <summary>
        /// Инициализирует форму выбора параметров отчета (параметры заданы по умолчанию).
        /// </summary>
        /// <param name="Session">Пользовательская сессия.</param>
        /// <param name="Context">Объектный контекст.</param>
        /// <param name="DefaultStartDate">Начальная дата по умолчанию.</param>
        /// <param name="DefaultEndDate">Конечная дата по умолчанию.</param>
        public ReportParameters(UserSession Session, ObjectContext Context, DateTime DefaultStartDate, DateTime DefaultEndDate)
        {
            InitializeComponent();
            this.SetNavigatorSkin();

            this.Context = Context;
            this.Session = Session;

            this.StartDate.EditValue = DefaultStartDate;
            this.EndDate.EditValue = DefaultEndDate;
        }
        
        private void OKButton_Click(object sender, EventArgs e)
        {
            if (this.StartDate.EditValue == null)
            {
                MyMessageBox.Show("Укажите дату, соответствующую началу периода.");
                return;
            }
            if (this.EndDate.EditValue == null)
            {
                MyMessageBox.Show("Укажите дату, соответствующую концу периода.");
                return;
            }
            if ((DateTime)this.StartDate.EditValue > (DateTime)this.EndDate.EditValue)
            {
                MyMessageBox.Show("Дата начала периода должна быть меньше даты окончания периода.");
                return;
            }
            DialogResult = DialogResult.OK;
            this.Close();
        }

        private void CancelButton_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
            this.Close();
        }

        private void ReportParameters_FormClosed(object sender, FormClosedEventArgs e)
        {
        }

    }
}