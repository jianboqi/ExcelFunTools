namespace ExcelFun
{
    partial class Ribbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region 组件设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.MyFirstAddin = this.Factory.CreateRibbonTab();
            this.groupCommon = this.Factory.CreateRibbonGroup();
            this.buttonGroup1 = this.Factory.CreateRibbonButtonGroup();
            this.btnAdressConv = this.Factory.CreateRibbonButton();
            this.btnFor2Num = this.Factory.CreateRibbonButton();
            this.buttonGroup2 = this.Factory.CreateRibbonButtonGroup();
            this.btnAddStr = this.Factory.CreateRibbonButton();
            this.groupOption = this.Factory.CreateRibbonGroup();
            this.toggleFormula = this.Factory.CreateRibbonToggleButton();
            this.groupCal = this.Factory.CreateRibbonGroup();
            this.buttonGroup3 = this.Factory.CreateRibbonButtonGroup();
            this.splitBtnCal = this.Factory.CreateRibbonSplitButton();
            this.btnCalAdd = this.Factory.CreateRibbonButton();
            this.btnCalMin = this.Factory.CreateRibbonButton();
            this.btnCalMui = this.Factory.CreateRibbonButton();
            this.btnCalDiv = this.Factory.CreateRibbonButton();
            this.groupView = this.Factory.CreateRibbonGroup();
            this.box1 = this.Factory.CreateRibbonBox();
            this.toggleBtnCol = this.Factory.CreateRibbonToggleButton();
            this.toggleBtnRow = this.Factory.CreateRibbonToggleButton();
            this.NumberTrans = this.Factory.CreateRibbonButton();
            this.box2 = this.Factory.CreateRibbonBox();
            this.MyFirstAddin.SuspendLayout();
            this.groupCommon.SuspendLayout();
            this.buttonGroup1.SuspendLayout();
            this.buttonGroup2.SuspendLayout();
            this.groupOption.SuspendLayout();
            this.groupCal.SuspendLayout();
            this.buttonGroup3.SuspendLayout();
            this.groupView.SuspendLayout();
            this.box1.SuspendLayout();
            this.box2.SuspendLayout();
            // 
            // MyFirstAddin
            // 
            this.MyFirstAddin.Groups.Add(this.groupCommon);
            this.MyFirstAddin.Groups.Add(this.groupOption);
            this.MyFirstAddin.Groups.Add(this.groupCal);
            this.MyFirstAddin.Groups.Add(this.groupView);
            this.MyFirstAddin.Label = "ExcelFun Tools";
            this.MyFirstAddin.Name = "MyFirstAddin";
            // 
            // groupCommon
            // 
            this.groupCommon.Items.Add(this.buttonGroup1);
            this.groupCommon.Items.Add(this.buttonGroup2);
            this.groupCommon.Label = "Common Tools";
            this.groupCommon.Name = "groupCommon";
            // 
            // buttonGroup1
            // 
            this.buttonGroup1.Items.Add(this.btnAdressConv);
            this.buttonGroup1.Items.Add(this.btnFor2Num);
            this.buttonGroup1.Name = "buttonGroup1";
            // 
            // btnAdressConv
            // 
            this.btnAdressConv.Label = "地址转换";
            this.btnAdressConv.Name = "btnAdressConv";
            this.btnAdressConv.OfficeImageId = "ScheduledProjectFinishDate";
            this.btnAdressConv.ShowImage = true;
            this.btnAdressConv.SuperTip = "绝对地址和相对地址转换。";
            this.btnAdressConv.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAdressConv_Click);
            // 
            // btnFor2Num
            // 
            this.btnFor2Num.Label = "数值转换";
            this.btnFor2Num.Name = "btnFor2Num";
            this.btnFor2Num.OfficeImageId = "_1";
            this.btnFor2Num.ShowImage = true;
            this.btnFor2Num.SuperTip = "将公式转换成数值。";
            this.btnFor2Num.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnFor2Num_Click);
            // 
            // buttonGroup2
            // 
            this.buttonGroup2.Items.Add(this.btnAddStr);
            this.buttonGroup2.Name = "buttonGroup2";
            // 
            // btnAddStr
            // 
            this.btnAddStr.Label = "追加字符串";
            this.btnAddStr.Name = "btnAddStr";
            this.btnAddStr.OfficeImageId = "MenuAddACalendar";
            this.btnAddStr.ShowImage = true;
            this.btnAddStr.SuperTip = "在单元格值后追加字符串。";
            this.btnAddStr.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAddStr_Click);
            // 
            // groupOption
            // 
            this.groupOption.Items.Add(this.toggleFormula);
            this.groupOption.Label = "Options";
            this.groupOption.Name = "groupOption";
            // 
            // toggleFormula
            // 
            this.toggleFormula.Label = "显示公式";
            this.toggleFormula.Name = "toggleFormula";
            this.toggleFormula.OfficeImageId = "TableFormulaDialog";
            this.toggleFormula.ShowImage = true;
            this.toggleFormula.SuperTip = "在单元格显示公式,而不显示值。";
            this.toggleFormula.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.toggleFormula_Click);
            // 
            // groupCal
            // 
            this.groupCal.Items.Add(this.buttonGroup3);
            this.groupCal.Items.Add(this.box2);
            this.groupCal.Label = "Calculation";
            this.groupCal.Name = "groupCal";
            // 
            // buttonGroup3
            // 
            this.buttonGroup3.Items.Add(this.splitBtnCal);
            this.buttonGroup3.Name = "buttonGroup3";
            // 
            // splitBtnCal
            // 
            this.splitBtnCal.Items.Add(this.btnCalAdd);
            this.splitBtnCal.Items.Add(this.btnCalMin);
            this.splitBtnCal.Items.Add(this.btnCalMui);
            this.splitBtnCal.Items.Add(this.btnCalDiv);
            this.splitBtnCal.Label = "四则运算";
            this.splitBtnCal.Name = "splitBtnCal";
            this.splitBtnCal.OfficeImageId = "MenuExpandCollapse";
            this.splitBtnCal.SuperTip = "对区域进行四则运算。";
            // 
            // btnCalAdd
            // 
            this.btnCalAdd.Label = "加(+)";
            this.btnCalAdd.Name = "btnCalAdd";
            this.btnCalAdd.ShowImage = true;
            this.btnCalAdd.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCalAdd_Click);
            // 
            // btnCalMin
            // 
            this.btnCalMin.Label = "减(-)";
            this.btnCalMin.Name = "btnCalMin";
            this.btnCalMin.ShowImage = true;
            this.btnCalMin.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCalMin_Click);
            // 
            // btnCalMui
            // 
            this.btnCalMui.Label = "乘(×)";
            this.btnCalMui.Name = "btnCalMui";
            this.btnCalMui.ShowImage = true;
            this.btnCalMui.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCalMui_Click);
            // 
            // btnCalDiv
            // 
            this.btnCalDiv.Label = "除(÷)";
            this.btnCalDiv.Name = "btnCalDiv";
            this.btnCalDiv.ShowImage = true;
            this.btnCalDiv.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCalDiv_Click);
            // 
            // groupView
            // 
            this.groupView.Items.Add(this.box1);
            this.groupView.Label = "Views";
            this.groupView.Name = "groupView";
            // 
            // box1
            // 
            this.box1.Items.Add(this.toggleBtnCol);
            this.box1.Items.Add(this.toggleBtnRow);
            this.box1.Name = "box1";
            // 
            // toggleBtnCol
            // 
            this.toggleBtnCol.Label = "列对比";
            this.toggleBtnCol.Name = "toggleBtnCol";
            this.toggleBtnCol.OfficeImageId = "SplitVertically";
            this.toggleBtnCol.ShowImage = true;
            this.toggleBtnCol.SuperTip = "将选中列集中在一起便于查看。";
            this.toggleBtnCol.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.toggleBtnCol_Click);
            // 
            // toggleBtnRow
            // 
            this.toggleBtnRow.Label = "行对比";
            this.toggleBtnRow.Name = "toggleBtnRow";
            this.toggleBtnRow.OfficeImageId = "SplitHorizontally";
            this.toggleBtnRow.ShowImage = true;
            this.toggleBtnRow.SuperTip = "将选中行集中在一起便于查看。";
            this.toggleBtnRow.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.toggleBtnRow_Click);
            // 
            // NumberTrans
            // 
            this.NumberTrans.Label = "数值转置";
            this.NumberTrans.Name = "NumberTrans";
            this.NumberTrans.OfficeImageId = "AccessFormPivotTable";
            this.NumberTrans.ShowImage = true;
            this.NumberTrans.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.NumberTrans_Click);
            // 
            // box2
            // 
            this.box2.Items.Add(this.NumberTrans);
            this.box2.Name = "box2";
            // 
            // Ribbon
            // 
            this.Name = "Ribbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.MyFirstAddin);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon_Load);
            this.MyFirstAddin.ResumeLayout(false);
            this.MyFirstAddin.PerformLayout();
            this.groupCommon.ResumeLayout(false);
            this.groupCommon.PerformLayout();
            this.buttonGroup1.ResumeLayout(false);
            this.buttonGroup1.PerformLayout();
            this.buttonGroup2.ResumeLayout(false);
            this.buttonGroup2.PerformLayout();
            this.groupOption.ResumeLayout(false);
            this.groupOption.PerformLayout();
            this.groupCal.ResumeLayout(false);
            this.groupCal.PerformLayout();
            this.buttonGroup3.ResumeLayout(false);
            this.buttonGroup3.PerformLayout();
            this.groupView.ResumeLayout(false);
            this.groupView.PerformLayout();
            this.box1.ResumeLayout(false);
            this.box1.PerformLayout();
            this.box2.ResumeLayout(false);
            this.box2.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab MyFirstAddin;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupCommon;
        internal Microsoft.Office.Tools.Ribbon.RibbonButtonGroup buttonGroup1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAdressConv;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupOption;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton toggleFormula;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupCal;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnFor2Num;
        internal Microsoft.Office.Tools.Ribbon.RibbonButtonGroup buttonGroup2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAddStr;
        internal Microsoft.Office.Tools.Ribbon.RibbonButtonGroup buttonGroup3;
        internal Microsoft.Office.Tools.Ribbon.RibbonSplitButton splitBtnCal;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCalAdd;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCalMin;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCalMui;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCalDiv;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupView;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton toggleBtnCol;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton toggleBtnRow;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box1;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton NumberTrans;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon Ribbon
        {
            get { return this.GetRibbon<Ribbon>(); }
        }
    }
}
