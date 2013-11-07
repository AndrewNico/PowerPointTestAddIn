using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using System.Windows.Forms;

namespace PowerPointTestAddIn
{
    public partial class ThisAddIn
    {
        private const string customPropertyName = "TestAddInCreator";

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            //Обработчики кнопок
            Globals.Ribbons.AddInRibbon.OnButtonCreatePresentationClick += new Action(createPresentation);
            Globals.Ribbons.AddInRibbon.OnButtonSlideAddClick += new Action(addSlide);

            //Обработчик активации окна
            Application.WindowActivate += Application_WindowActivate;
        }

        private void Application_WindowActivate(Presentation Pres, DocumentWindow Wn)
        {

            Globals.Ribbons.AddInRibbon.AddSlideEnable = isCreateByThisAddIn();
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e) { }

        /// <summary>
        /// Создание новой презентации с тремя слайдами и CustomDocumentProperty "TestAddInCreator"
        /// </summary>
        private void createPresentation()
        {
            try
            {
                // Добавляем презентацию.
                Presentation presentation = Application.Presentations.Add(MsoTriState.msoTrue);

                // Добавляем флаг и устанавливаем, что презентация была создана надстройкой.
                DocumentProperties properties = presentation.CustomDocumentProperties;
                properties.Add(customPropertyName, false, MsoDocProperties.msoPropertyTypeBoolean,
                    true, missing);

                // Добавляем слайды.
                presentation.Slides.Add(1, PpSlideLayout.ppLayoutTitleOnly);		// Только с заголовком
                presentation.Slides.Add(2, PpSlideLayout.ppLayoutText);			    // С текстовым полем
                presentation.Slides.Add(3, PpSlideLayout.ppLayoutTwoColumnText);	// С двумя текстовыми полями
            }
            catch (Exception e)
            {
                showErrStack(e);        
            }
        }

        /// <summary>
        /// Добавление пустого слайда в конец активной презентации
        /// </summary>
        private void addSlide()
        {
            try
            {
                Presentation presentation = Application.ActivePresentation;

                presentation.Slides.Add(presentation.Slides.Count + 1, PpSlideLayout.ppLayoutBlank);
            }
            catch (Exception e)
            {
                showErrStack(e);
            }
        }

        /// <summary>
        /// Создана ли активная презентация данной надстройкой
        /// </summary>
        /// <returns>True - создана, false - не создана</returns>
        private bool isCreateByThisAddIn()
        {
            Presentation presentation = Application.ActivePresentation;

            if (presentation == null)
                return false;

            DocumentProperties properties = presentation.CustomDocumentProperties;

            if (isPropertyExist(customPropertyName))
                return (bool)properties[customPropertyName].Value;	
            else
                return false;

        }

        /// <summary>
        /// Поиск необходимого свойства в активном документе
        /// </summary>
        /// <param name="propertyName">Имя свойства</param>
        /// <returns>True - свойство существует, false - свойство отсутствует</returns>
        private bool isPropertyExist(string propertyName)
        {
            DocumentProperties properties = Application.ActivePresentation.CustomDocumentProperties;
            foreach (DocumentProperty prop in properties)
            {
                if (prop.Name == propertyName)
                    return true;
            }
            return false;
        }

        /// <summary>
        /// Отображает сообщения об ошибке по всему стеку 
        /// </summary>
        /// <param name="e">Объект ошибки</param>
        private void showErrStack(Exception e)
        {
            string err = e.Message + Environment.NewLine;
            Exception tmp = e;

            while (tmp.InnerException != null)
            {
                tmp = tmp.InnerException;
                err = err + tmp.Message + Environment.NewLine;
            }

            MessageBox.Show(err, "Ошибка!",
                MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
        }

        #region Код, автоматически созданный VSTO

        /// <summary>
        /// Обязательный метод для поддержки конструктора - не изменяйте
        /// содержимое данного метода при помощи редактора кода.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
