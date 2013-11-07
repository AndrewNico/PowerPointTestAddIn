namespace PowerPointTestAddIn
{
    partial class AddInRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Требуется переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public AddInRibbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором компонентов

        /// <summary>
        /// Обязательный метод для поддержки конструктора - не изменяйте
        /// содержимое данного метода при помощи редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            this.tabTest = this.Factory.CreateRibbonTab();
            this.groupTest = this.Factory.CreateRibbonGroup();
            this.buttonCreatePresentation = this.Factory.CreateRibbonButton();
            this.buttonSlideAdd = this.Factory.CreateRibbonButton();
            this.tabTest.SuspendLayout();
            this.groupTest.SuspendLayout();
            // 
            // tabTest
            // 
            this.tabTest.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tabTest.Groups.Add(this.groupTest);
            this.tabTest.Label = "Тестовое задание";
            this.tabTest.Name = "tabTest";
            // 
            // groupTest
            // 
            this.groupTest.Items.Add(this.buttonCreatePresentation);
            this.groupTest.Items.Add(this.buttonSlideAdd);
            this.groupTest.Label = "Тест";
            this.groupTest.Name = "groupTest";
            // 
            // buttonCreatePresentation
            // 
            this.buttonCreatePresentation.Label = "Создать презентацию";
            this.buttonCreatePresentation.Name = "buttonCreatePresentation";
            this.buttonCreatePresentation.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonCreatePresentation_Click);
            // 
            // buttonSlideAdd
            // 
            this.buttonSlideAdd.Enabled = false;
            this.buttonSlideAdd.Label = "Добавить слайд";
            this.buttonSlideAdd.Name = "buttonSlideAdd";
            this.buttonSlideAdd.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonSlideAdd_Click);
            // 
            // AddInRibbon
            // 
            this.Name = "AddInRibbon";
            this.RibbonType = "Microsoft.PowerPoint.Presentation";
            this.Tabs.Add(this.tabTest);
            this.tabTest.ResumeLayout(false);
            this.tabTest.PerformLayout();
            this.groupTest.ResumeLayout(false);
            this.groupTest.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabTest;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupTest;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonCreatePresentation;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonSlideAdd;
    }

    partial class ThisRibbonCollection
    {
        internal AddInRibbon AddInRibbon
        {
            get { return this.GetRibbon<AddInRibbon>(); }
        }
    }
}
