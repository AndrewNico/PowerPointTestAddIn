using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

namespace PowerPointTestAddIn
{
    public partial class AddInRibbon
    {
        /// <summary>
        /// Событие, возникающее при клике на кнопку "Создать презентацию"
        /// </summary>
        public event Action OnButtonCreatePresentationClick;

        /// <summary>
        /// Событие, возникающее при клике на кнопку "Добавить слайд"
        /// </summary>
        public event Action OnButtonSlideAddClick;

        /// <summary>
        /// Свойство доступности кнопки "Добавить слайд"
        /// </summary>
        public bool AddSlideEnable
        {
            set { buttonSlideAdd.Enabled = value; } 
        }

        private void buttonCreatePresentation_Click(object sender, RibbonControlEventArgs e)
        {
            if (OnButtonCreatePresentationClick != null)
                OnButtonCreatePresentationClick();
        }

        private void buttonSlideAdd_Click(object sender, RibbonControlEventArgs e)
        {
            if (OnButtonSlideAddClick != null)
                OnButtonSlideAddClick();
        }
    }
}
