using DevExpress.Skins;
using DevExpress.XtraEditors;
using DocsVision.Platform.Wpf.Navigator;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using SKB.Base;

namespace SKB.NavigatorExtension
{
    /// <summary>
    /// Класс-помощник.
    /// </summary>
    public static class ExtensionHelper
    {
        /// <summary>
        /// Устанавливает тему Навигатора.
        /// </summary>
        /// <param name="Form"></param>
        public static void SetNavigatorSkin (this XtraForm Form)
        {
            SkinManager.EnableFormSkins();
            Type NavigatorControlType = Type.GetType("DocsVision.Platform.Wpf.Navigator.NavigatorControl, DocsVision.Platform.Wpf.Navigator, Version=5.0.0.0, Culture=neutral, PublicKeyToken=7148afe997f90519");
            UserSettings Settings = (UserSettings)NavigatorControlType.GetProperty("UserSettings").GetValue(NavigatorControlType.GetProperty("Current").GetValue(null, null), null);

            String SkinName;
            switch (Settings.ColorScheme)
            {
                case "Blue": SkinName = "Office 2010 Blue"; break;
                case "Silver": SkinName = "Office 2010 Silver"; break;
                case "Future": SkinName = "Seven"; break; 
                case "Classic": SkinName = "Seven Classic"; break;
                default: SkinName = String.Empty; break;
            }

            Form.LookAndFeel.SetSkinStyle(SkinName);
            DevExpress.LookAndFeel.UserLookAndFeel.Default.SetSkinStyle(SkinName);
        }

        /// <summary>
        /// Получение текста ошибки по коду.
        /// </summary>
        /// <param name="ErrorCode"> Код ошибки.</param>
        /// <param name="ErrorCount"> Количество ошибок.</param>
        /// <returns></returns>
        public static string GetErrorText(Int32 ErrorCode, Int32 ErrorCount)
        {
            switch (ErrorCode)
            {
                case 0:
                    return "У вас нет прав на редактирование " + ErrorCount.GetCaseString("данного документа", "данных документов", "данных документов");
                case 1:
                    return ErrorCount.GetCaseString("Данный документ уже добавлен", "Данные документы уже добавлены", "Данные документы уже добавлены") + " в текущую карточку согласования";
                case 2:
                    return ErrorCount.GetCaseString("Данный документ не является черновиком", "Данные документы не являются черновиками", "Данные документы не являются черновиками");
                case 3:
                case 4:
                    return ErrorCount.GetCaseString("Данный документ имеет", "Данные документы имеют", "Данные документы имеют") + " неверный тип";
                default:
                    return String.Empty;
            }
        }
    }
}