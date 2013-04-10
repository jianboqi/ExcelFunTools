using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Tools.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;

namespace ExcelFun
{
    public partial class Ribbon
    {
        Excel.Range __Range;
        Excel.Worksheet __worksheet;//列视图时保存的worksheet
        Excel.Worksheet __worksheet1;//行视图
        string addrTempstr;
        string addrTempstr1;//行视图时保存隐藏的行
        System.Collections.ArrayList arrCol;
        
        private void Ribbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        //公式转换  绝对引用改为相对引用  相对应用改为绝对引用
        private void btnAdressConv_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Range selectRng = Globals.ThisAddIn.Application.Selection;
            selectRng = selectRng.SpecialCells(Excel.XlCellType.xlCellTypeFormulas);
            string formulaStr;
            if (selectRng != null)
            {
                foreach (Excel.Range rng in selectRng)
                {
                    formulaStr = rng.Formula;
                    if (formulaStr.Contains("$"))
                    {
                        formulaStr = Globals.ThisAddIn.Application.ConvertFormula(rng.Formula, Excel.XlReferenceStyle.xlA1, Excel.XlReferenceStyle.xlA1, Excel.XlReferenceType.xlRelative);
                        if (rng.HasArray)
                        {
                            rng.FormulaArray = formulaStr;
                        }
                        else
                        {
                            rng.Formula = formulaStr;
                        }
                    } 
                    else
                    {
                        formulaStr = Globals.ThisAddIn.Application.ConvertFormula(rng.Formula, Excel.XlReferenceStyle.xlA1, Excel.XlReferenceStyle.xlA1, Excel.XlReferenceType.xlAbsolute);
                        if (rng.HasArray)
                        {
                            rng.FormulaArray = formulaStr;
                        }
                        else
                        {
                            rng.Formula = formulaStr;
                        }
                    }
                }
            }
        }

        //是否显示公式
        private void toggleFormula_Click(object sender, RibbonControlEventArgs e)
        {
            if (toggleFormula.Checked)
            {
                Globals.ThisAddIn.Application.ActiveWindow.DisplayFormulas = true;
            }
            else
            {
                Globals.ThisAddIn.Application.ActiveWindow.DisplayFormulas = false;
            }
            
            
        }
        //公式转换成数值
        private void btnFor2Num_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Range selectRng = Globals.ThisAddIn.Application.Selection;
            if (selectRng != null)
            {
                selectRng.Value = selectRng.Value;
            }
        }
        //在单元格后追加字符串
        private void btnAddStr_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Range selectRng = Globals.ThisAddIn.Application.Selection;
            if (selectRng != null)
            {
                string inputstr =  Globals.ThisAddIn.Application.InputBox("请输入字符串:");
                if (inputstr != "")
                {
                    foreach (Excel.Range rng in selectRng)
                    {
                        rng.Value = rng.Value + inputstr;
                    }
                }
               
            }
        }
        //四则运算  
        private void btnCalAdd_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Range selectRng = Globals.ThisAddIn.Application.Selection;
            string inputNum = Globals.ThisAddIn.Application.InputBox("请输入一个数:");
            if (selectRng != null && inputNum != "")
            {
                double numtemp;
                bool isOK;
                isOK = double.TryParse(inputNum, out numtemp);
                if (!isOK)
                {
                    MessageBox.Show("请输入一个数字!");
                } 
                else
                {
                    foreach (Excel.Range rng in selectRng)
                    {
                        dynamic d = rng.Value;
                        if (d !=null)
                        {
                            string str = d.ToString();
                            double rngVal;
                            isOK = double.TryParse(str, out rngVal);
                            if (isOK)
                            {
                                rng.Value = rngVal + numtemp;
                            }
                        }
                        
                        
                    }
                }

            }
        }

        private void btnCalMin_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Range selectRng = Globals.ThisAddIn.Application.Selection;
            string inputNum = Globals.ThisAddIn.Application.InputBox("请输入一个数:");
            if (selectRng != null && inputNum != "")
            {
                double numtemp;
                bool isOK;
                isOK = double.TryParse(inputNum, out numtemp);
                if (!isOK)
                {
                    MessageBox.Show("请输入一个数字!");
                }
                else
                {
                    foreach (Excel.Range rng in selectRng)
                    {
                        dynamic d = rng.Value;
                        if (d != null)
                        {
                            string str = d.ToString();
                            double rngVal;
                            isOK = double.TryParse(str, out rngVal);
                            if (isOK)
                            {
                                rng.Value = rngVal - numtemp;
                            }
                        }


                    }
                }

            }
        }

        private void btnCalMui_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Range selectRng = Globals.ThisAddIn.Application.Selection;
            string inputNum = Globals.ThisAddIn.Application.InputBox("请输入一个数:");
            if (selectRng != null && inputNum != "")
            {
                double numtemp;
                bool isOK;
                isOK = double.TryParse(inputNum, out numtemp);
                if (!isOK)
                {
                    MessageBox.Show("请输入一个数字!");
                }
                else
                {
                    foreach (Excel.Range rng in selectRng)
                    {
                        dynamic d = rng.Value;
                        if (d != null)
                        {
                            string str = d.ToString();
                            double rngVal;
                            isOK = double.TryParse(str, out rngVal);
                            if (isOK)
                            {
                                rng.Value = rngVal * numtemp;
                            }
                        }


                    }
                }

            }
        }

        private void btnCalDiv_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Range selectRng = Globals.ThisAddIn.Application.Selection;
            string inputNum = Globals.ThisAddIn.Application.InputBox("请输入一个数:");
            if (selectRng != null && inputNum != "")
            {
                double numtemp;
                bool isOK;
                isOK = double.TryParse(inputNum, out numtemp);
                if (!isOK)
                {
                    MessageBox.Show("请输入一个数字!");
                }
                else
                {
                    foreach (Excel.Range rng in selectRng)
                    {
                        dynamic d = rng.Value;
                        if (d != null)
                        {
                            string str = d.ToString();
                            double rngVal;
                            isOK = double.TryParse(str, out rngVal);
                            if (isOK)
                            {
                                rng.Value = rngVal / numtemp;
                            }
                        }


                    }
                }

            }
        }

        //lie
        private void toggleBtnCol_Click(object sender, RibbonControlEventArgs e)
        {
            if (toggleBtnCol.Checked)
            {
                __worksheet = Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet;
             //   arlist = new System.Collections.ArrayList();
                Excel.Range selectRng = Globals.ThisAddIn.Application.Selection;
                String selectAddr = selectRng.Address;
                if (selectAddr == selectRng.EntireColumn.Address)
                {
                    if (selectAddr.Contains(","))//至少包含两列
                    {
                        int minCol = selectRng.Column;
                        int areaCount = selectRng.Areas.Count;
                        int maxCol = selectRng.Areas[areaCount].Columns[selectRng.Areas[areaCount].Columns.Count].Column;
                        string allAddressStr = "";
                        System.Collections.ArrayList selectArr = new System.Collections.ArrayList();
                        for (int k = 1; k <= areaCount; k++)
                        {
                            for (int j = 1; j <= selectRng.Areas[k].Columns.Count; j++)
                            {
                                allAddressStr = allAddressStr + selectRng.Areas[k].Columns[j].address + ",";
                               // selectArr.Add(selectRng.Areas[k].Columns[j].address);
                            }
                        }


                        //对两列之间进行循环
                        string addrStr = "";
                        bool isRestart = true;
                        System.Collections.ArrayList arr = new System.Collections.ArrayList();
                        
                        for (int i = minCol + 1; i <= maxCol; i++)
                        {
                            if (!allAddressStr.Contains(Globals.ThisAddIn.Application.Columns[i].address))
                            {
                                if (isRestart == true)//如果重新开始记录一个空白区域
                                {
                                    arr.Add(Globals.ThisAddIn.Application.Columns[i].address);
                                    isRestart = false;
                                }
                            }
                            else
                            {
                                if (isRestart == false)
                                {
                                    arr.Add(Globals.ThisAddIn.Application.Columns[i - 1].address);
                                    isRestart = true;
                                } 
                            }
                        }
                        System.Collections.ArrayList unarr = new System.Collections.ArrayList();
                        for (int i = 0; i < arr.Count; i = i + 2)
                        {
                            string temp = arr[i].ToString();
                            string temp1 = arr[i + 1].ToString();
                            string area = temp.Split(':')[0] +":"+ temp1.Split(':')[1];
                            unarr.Add(area);
                        }
                        string[] addrlist = (string[])unarr.ToArray(typeof(string));
                        addrStr = string.Join(",", addrlist);
                        addrTempstr = addrStr;
                       Globals.ThisAddIn.Application.Range[addrStr].EntireColumn.Hidden = true;
                    }
                }
                else
                {
                    MessageBox.Show("请选择整列!");
                    toggleBtnCol.Checked = false;
                }
            } 
            else
            {
                __worksheet.Range[addrTempstr].EntireColumn.Hidden = false;
            }
            
        }

        private void toggleBtnRow_Click(object sender, RibbonControlEventArgs e)
        {
            if (toggleBtnRow.Checked)
            {
                __worksheet1 = Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet;
                //arlist1 = new System.Collections.ArrayList();
                Excel.Range selectRng = Globals.ThisAddIn.Application.Selection;
                String selectAddr = selectRng.Address;
                if (selectAddr == selectRng.EntireRow.Address)
                {
                    if (selectAddr.Contains(","))//至少包含两行
                    {
                        int minRow = selectRng.Row;
                        int areaCount = selectRng.Areas.Count;
                        int maxRow = selectRng.Areas[areaCount].Rows[selectRng.Areas[areaCount].Rows.Count].Row;
                        string allAddressStr = "";
                        for (int k = 1; k <= areaCount; k++)
                        {
                            for (int j = 1; j <= selectRng.Areas[k].Rows.Count; j++)
                            {
                                allAddressStr = allAddressStr + selectRng.Areas[k].Rows[j].address + ",";
                            }
                        }
                        
                        //对两列之间进行循环tr
                        string addstr1 = "";
                        bool isRestart = true;
                        System.Collections.ArrayList arr = new System.Collections.ArrayList();
                        for (int i = minRow + 1; i <= maxRow; i++)
                        {
                            if (!allAddressStr.Contains(Globals.ThisAddIn.Application.Rows[i].address))
                            {
                                //arr.Add(Globals.ThisAddIn.Application.Rows[i].address);
                                if (isRestart == true)//如果重新开始记录一个空白区域
                                {
                                    arr.Add(Globals.ThisAddIn.Application.Rows[i].address);
                                    isRestart = false;
                                }
                            }
                            else
                            {
                                if (isRestart == false)
                                {
                                    arr.Add(Globals.ThisAddIn.Application.Rows[i - 1].address);
                                    isRestart = true;
                                }
                            }

                        }
                        System.Collections.ArrayList unarr = new System.Collections.ArrayList();
                        for (int i = 0; i < arr.Count; i = i + 2)
                        {
                            string temp = arr[i].ToString();
                            string temp1 = arr[i + 1].ToString();
                            string area = temp.Split(':')[0] + ":" + temp1.Split(':')[1];
                            unarr.Add(area);
                        }
                        string[] addrlist = (string[])unarr.ToArray(typeof(string));
                        addstr1=string.Join(",", addrlist);
                        addrTempstr1 = addstr1;
                        Globals.ThisAddIn.Application.Range[addstr1].EntireRow.Hidden = true;
                    }
                }
                else
                {
                    MessageBox.Show("请选择整行!");
                    toggleBtnRow.Checked = false;
                }
            }
            else
            {
                    __worksheet1.Range[addrTempstr1].EntireRow.Hidden = false;
            }
        }

        private void NumberTrans_Click(object sender, RibbonControlEventArgs e)
        {
           Excel.Range selectRng = Globals.ThisAddIn.Application.Selection;
           selectRng.PasteSpecial(Excel.XlPasteType.xlPasteValues, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone,false,true);
           Globals.ThisAddIn.Application.CutCopyMode = Excel.XlCutCopyMode.xlCopy;
        } 
    }
}
