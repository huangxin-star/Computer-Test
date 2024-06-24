using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.Streaming.Values;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using static NPOI.HSSF.Util.HSSFColor;
using static NPOI.SS.UserModel.IFont;


namespace Computer_Test
{
    public class Document
    {
        public MainForm mf;
        public Document(MainForm mf) {
            this.mf = mf;
        }
        public void createDocuments(string mappingPath,string machinePath,string startupPath,string savaPath)
        {
            XSSFWorkbook WB;
            WB = new XSSFWorkbook(File.Open(startupPath + "\\Test Plan.xlsx", FileMode.Open));//打开报表
            //XSSFSheet mappingsheet = (XSSFSheet)WB.CreateSheet("映射表");//在newexcel里新建映射表
            //WB.SetSheetOrder("映射表", 0);
            XSSFWorkbook workbookwrite = new XSSFWorkbook(File.Open(mappingPath, FileMode.Open));//打开映射表总表
            // {Name: /xl/worksheets/sheet3.xml - Content Type: application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml}
            XSSFSheet mappingTable = (XSSFSheet)workbookwrite.GetSheet("NB");
            WB.SetSheetOrder("NB", 0);

            XSSFSheet testPlan = (XSSFSheet)WB.GetSheet("TestPlan");
            int flag = 1; //保存文件名

            string modelName = "";
            string modelName2 = "";
            string modelNameNext2 = "";
            int typeA = 0;
            int typeC = 0;

            //for (int i = 0; i <= mappingTable.LastRowNum; i++) // 97
            //{
            //    IRow mappingRow = mappingTable.GetRow(i);
            //    IRow decRow = mappingsheet.CreateRow(i);
            //    if (mappingRow == null)
            //        continue;
            //    for (int j = 0; j < 5; j++)
            //    {
            //        if (mappingRow.GetCell(j) == null)
            //            continue;
            //        string s = mappingRow.GetCell(j).ToString();

            //        decRow.CreateCell(j, CellType.String).SetCellValue(s);
            //        // testRow.CreateCell(j, CellType.String).SetCellValue(s);

            //    }
            //}

            XSSFWorkbook machineMapping = new XSSFWorkbook(File.Open(machinePath, FileMode.Open));//打开機種映射表总表
            XSSFSheet p_machineSheet = (XSSFSheet)WB.CreateSheet("機種映射表");//在newexcel里新建映射表
            WB.SetSheetOrder("機種映射表", 1);



            string searchString1 = "Spec";
            string searchString2 = "SPEC";
            string searchString3 = "spec";
            List<XSSFSheet> sheetsContainingString = new List<XSSFSheet>();

            // int rowOffset = 0; // 行偏移量

            // 遍历所有工作表
            foreach (ISheet sheet in machineMapping)
            {
                // 检查工作表名称是否包含特定字符串
                if (sheet.SheetName.Contains(searchString1) || sheet.SheetName.Contains(searchString2) || sheet.SheetName.Contains(searchString3))
                {
                    // 如果包含，添加到列表中，并转换为XSSFSheet类型（如果需要）
                    sheetsContainingString.Add(sheet as XSSFSheet);
                    for (int i = 0; i < sheetsContainingString.Count; i++)
                    {
                        for (int rowNum = 0; rowNum <= sheet.LastRowNum; rowNum++)
                        {
                            // 當前行
                            XSSFRow row = (XSSFRow)sheet.GetRow(rowNum);
                            IRow machineRow = p_machineSheet.CreateRow(rowNum);
                            if (row == null)
                                continue;
                            for (int j = 0; j < 14; j++) // 可能需要改動
                            {
                                if (row.GetCell(j) == null)
                                    continue;
                                string s = row.GetCell(j).ToString();
                                machineRow.CreateCell(j, CellType.String).SetCellValue(s);

                            }
                        }
                    }

                }
            }



            // 移除空列
            RemoveEmptyColumns(p_machineSheet);

            // 在前3行插入空行
            InsertEmptyRows(testPlan, 3);
            // 原有的代码逻辑
            for (int i = 0; i <= testPlan.LastRowNum; i++)
            {
                IRow testPlanRow = testPlan.GetRow(i);
                if (testPlanRow == null) continue;


                for (int j = 1; j < 3; j++)
                {
                    ICell cell = testPlanRow.GetCell(j);
                    if (cell == null) continue;
                    
                    string cellValue = cell.ToString();
                    // 设置名称
                    if (cellValue.Contains("NPI"))
                    {
                        // 后面第一个单元格
                        modelName = GetFieldValue(p_machineSheet, "Model name");
                        modelName2 = GetFieldValue(p_machineSheet, "Model Name");
                        // 后面第二个单元格
                        string modelNameNext = GetFieldValue2(p_machineSheet, "Model name");
                        modelNameNext2 = GetFieldValue2(p_machineSheet, "Model Name");

                        if ((modelName != null && modelName!=""))
                        {

                            if (ContainsField(p_machineSheet, "Model name"))
                            {
                                SetCellValue(testPlanRow, j + 1, modelName);
                                flag = 1;

                            }
                        }

                        if ((modelNameNext != null && modelNameNext != "") && (modelName == null || modelName == "N/A" || modelName == "na" || modelName == "NA" || modelName == ""))
                        {
                            if (ContainsField(p_machineSheet, "Model name"))
                            {
                                SetCellValue(testPlanRow, j + 1, modelNameNext);
                                flag = 0;
                            }
                        }
                        if ((modelName2 != null && modelName2 != ""))
                        {

                            if (ContainsField(p_machineSheet, "Model Name"))
                            {
                                SetCellValue(testPlanRow, j + 1, modelName2);
                                flag = 1;
                            }
                        }

                        if ((modelNameNext2 != null && modelNameNext2 != "") && (modelName2 == null || modelName2 == "N/A" || modelName2 == "na" || modelName2 == "NA" || modelName2 == ""))
                        {
                            if (ContainsField(p_machineSheet, "Model Name"))
                            {
                                SetCellValue(testPlanRow, j + 1, modelNameNext2);
                                flag = 0;
                            }
                        }
                    }                  

                    else if (cellValue.Contains("Processor Type"))
                    {
                        // string processorValue = GetFieldValue(p_machineSheet, "Processor Type");
                        if (ContainsField(p_machineSheet, "Processor Type"))
                        {
                            SetCellValue(testPlanRow, j + 1, "Y");
                            SetProcessorPlatValue(testPlan);
                        }
                       
                    }
                    else if (cellValue.Equals("PCH型號"))
                    {
                        // string vramValue = GetFieldValue(p_machineSheet, "vRAM");
                        if (EqualsField(p_machineSheet, "PCH"))
                        {
                            SetCellValue(testPlanRow, j + 1, "Y");
                        }

                    }
                    else if (cellValue.Equals("GPU型號"))
                    {
                        string PCIEValue = GetFieldValue(p_machineSheet, "PCIe Interface");
                        string gpu = GetFieldValue(p_machineSheet, "GPU-1");
                        
                        string block = "TDP / VRAM Interface / Capacity";
                        string blockValue = GetFieldValue2(p_machineSheet, block);



                        if (ContainsField(p_machineSheet, "PCIe Interface"))
                        {
                            SetCellValue(testPlanRow, j + 1, "Y");
                            SetCellValue(testPlanRow, j + 3, PCIEValue);
                            SetRamValue(testPlan);
                        }else if(ContainsField(p_machineSheet, "GPU-1"))
                        {
                            SetCellValue(testPlanRow, j + 1, "Y");
                            SetCellValue(testPlanRow, j + 3, gpu);
                            SetRamValue(testPlan);
                        }
                        else if (ContainsField(p_machineSheet, block))
                        {
                            SetCellValue(testPlanRow, j + 1, "Y");
                            SetCellValue(testPlanRow, j + 3, blockValue);
                            SetRamValue(testPlan);
                        }

                    }
                    else if (cellValue.Equals("HDD數量"))
                    {
                        // string vramValue = GetFieldValue(p_machineSheet, "vRAM");
                        if (ContainsField(p_machineSheet, "SSD Slot") || ContainsField(p_machineSheet, "HDD") || ContainsField(p_machineSheet, "SSD Form Factor") || ContainsField(p_machineSheet, "SSD form factor"))
                        {
                            SetCellValue(testPlanRow, j + 1, "Y");
                        }
                        
                    }
                    else if (cellValue.Equals("Memory Slot"))
                    {
                        string slotValue = GetFieldValue(p_machineSheet, "Memory Slot");
                        if (ContainsField(p_machineSheet, "Memory Slot"))
                        {
                            SetCellValue(testPlanRow, j + 1, "Y");
                            SetCellValue(testPlanRow, j + 3, slotValue);
                        }
                        else
                        {
                            // False

                            SetCapacity(testPlan);
                        }

                    }

                    // 相机
                    else if (cellValue.Equals("Camera"))
                    {
                        string carmValue = GetFieldValue(p_machineSheet, "Cam Resolution");
                        if (ContainsField(p_machineSheet, "Cam Resolution"))
                        {
                            SetCellValue(testPlanRow, j + 1, "Y");
                            SetCellValue(testPlanRow, j + 3, carmValue);
                        }
                        SetCellValue(testPlanRow, j + 1, "Y");

                    }
                    else if (cellValue.Equals("Rear Cam"))
                    {
                        // string vramValue = GetFieldValue(p_machineSheet, "vRAM");
                        if (ContainsField(p_machineSheet, "Rear  Cam") || ContainsField(p_machineSheet, "Camera (Rear)"))
                        {
                            SetCellValue(testPlanRow, j + 1, "Y");
                        }
                     
                    }
                    else if (cellValue.Equals("Front Cam"))
                    {
                        // string vramValue = GetFieldValue(p_machineSheet, "vRAM");
                        if (ContainsField(p_machineSheet, "Front Cam") || ContainsField(p_machineSheet, "Camera (Front)"))
                        {
                            SetCellValue(testPlanRow, j + 1, "Y");
                        }
                      
                    }
                    else if (cellValue.Equals("IR Cam"))
                    {
                        // string vramValue = GetFieldValue(p_machineSheet, "vRAM");
                        if (ContainsField(p_machineSheet, "Type (IR/Normal)"))
                        {
                            SetCellValue(testPlanRow, j + 1, "Y");
                        }
                        else if (ContainsField(p_machineSheet, "IR Cam"))
                        {
                            SetCellValue(testPlanRow, j + 1, "Y");
                        }


                    }
                    else if (cellValue.Equals("Cam Switch"))
                    {
                        // string vramValue = GetFieldValue(p_machineSheet, "vRAM");
                        if (ContainsField(p_machineSheet, "Cam Switch")||ContainsField(p_machineSheet, "Camera Switch"))
                        {
                            SetCellValue(testPlanRow, j + 1, "Y");
                        }
                      
                    }
                    else if (cellValue.Contains("Mic")) 
                    {
                        // string vramValue = GetFieldValue(p_machineSheet, "vRAM");
                        if (ContainsField(p_machineSheet, "Mic Q’ty")|| ContainsField(p_machineSheet, "Mic Q'ty") || ContainsField(p_machineSheet, "Mic Type"))
                        {
                            SetCellValue(testPlanRow, j + 1, "Y");
                        }else if(ContainsField(p_machineSheet, "Other'"))
                        {
                            SetCellValue(testPlanRow, j + 1, "Y");
                        }
                       
                    }
                    // 风扇
                    else if (cellValue.Equals("FAN"))
                    {
                        // string vramValue = GetFieldValue(p_machineSheet, "vRAM");
                        if (ContainsField(p_machineSheet, "Fan"))
                        {
                            SetCellValue(testPlanRow, j + 1, "Y");
                        }

                    }
                    // BT & WLAN
                    else if (cellValue.Equals("BT"))
                    {
                      
                        // string vramValue = GetFieldValue(p_machineSheet, "vRAM");
                        if (EqualsField(p_machineSheet, "BT")||ContainsField(p_machineSheet, "Wi-Fi/BT"))
                        {
                            SetCellValue(testPlanRow, j + 1, "Y");
                           
                        }

                    }
                    else if (cellValue.Equals("Wlan"))
                    {
                        // string vramValue = GetFieldValue(p_machineSheet, "vRAM");
                        if (EqualsField(p_machineSheet, "WLAN") || ContainsField(p_machineSheet, "Wi-Fi/BT"))
                        {
                            SetCellValue(testPlanRow, j + 1, "Y");
                        }

                    }
                    // audio
                    else if (cellValue.Equals("Codec check"))
                    {
                        // string vramValue = GetFieldValue(p_machineSheet, "vRAM");
                        if (ContainsField(p_machineSheet, "Codec"))
                        {
                            SetCellValue(testPlanRow, j + 1, "Y");
                        }

                    }
                    else if (cellValue.Equals("speak test"))
                    {
                        {
                            SetCellValue(testPlanRow, j + 1, "Y");
                        }
                    }
                    else if (cellValue.Equals("Headphone"))
                    {
                        // string vramValue = GetFieldValue(p_machineSheet, "vRAM");
                        if (ContainsField(p_machineSheet, "Audio jack") || ContainsField(p_machineSheet, "Jack"))
                        {
                            SetCellValue(testPlanRow, j + 1, "Y");
                            SetAudioValue(testPlan);
                        }                     
                    }
                    else if (cellValue.Equals("Woofer"))
                    {
                        // string vramValue = GetFieldValue(p_machineSheet, "vRAM");
                        if (ContainsField(p_machineSheet, "Woofer"))
                        {
                            SetCellValue(testPlanRow, j + 1, "Y");
                        }

                    }
                    else if (cellValue.Equals("Tweeter"))
                    {
                        // string vramValue = GetFieldValue(p_machineSheet, "vRAM");
                        if (ContainsField(p_machineSheet, "Tweeter"))
                        {
                            SetCellValue(testPlanRow, j + 1, "Y");
                        }

                    }
                    else if (cellValue.Equals("DC-in"))
                    {
                        // string vramValue = GetFieldValue(p_machineSheet, "vRAM");
                        // if (ContainsField(p_machineSheet, "DC-in"))
                        // 默认都测
                        
                            SetCellValue(testPlanRow, j + 1, "Y");
                        

                    }
                    else if (cellValue.Contains("FN+F2"))
                    {
                        string fn2 = GetFieldValue(p_machineSheet, "Fn2");
                        string nextF2 = GetFieldValue2(p_machineSheet, "Fn2");

                        if (fn2 != null && fn2!="")
                        {
                            if (ContainsField(p_machineSheet, "Fn2"))
                            {
                                SetCellValue(testPlanRow, j + 1, "Y");
                                SetCellValue(testPlanRow, j + 3, fn2);

                            }
                        }
                        if ((nextF2 != null && nextF2!="") && (fn2 == null||fn2=="N/A"||fn2=="na"||fn2=="NA"||fn2==""))
                        {
                            if (ContainsField(p_machineSheet, "Fn2"))
                            {
                                SetCellValue(testPlanRow, j + 1, "Y");
                                SetCellValue(testPlanRow, j + 3, nextF2);

                            }
                        }


                    }
                    else if (cellValue.Contains("FN+F3"))
                    {
                        string fn3 = GetFieldValue(p_machineSheet, "Fn3");
                        string nextF3 = GetFieldValue2(p_machineSheet, "Fn3");

                        if (fn3 != null && fn3 != "")
                        {
                            if (ContainsField(p_machineSheet, "Fn2"))
                            {
                                SetCellValue(testPlanRow, j + 1, "Y");
                                SetCellValue(testPlanRow, j + 3, fn3);

                            }
                        }
                        if ((nextF3 != null && nextF3 != "") && (fn3 == null || fn3 == "N/A" || fn3 == "na" || fn3 == "NA" || fn3 == ""))
                        {
                            if (ContainsField(p_machineSheet, "Fn3"))
                            {
                                SetCellValue(testPlanRow, j + 1, "Y");
                                SetCellValue(testPlanRow, j + 3, nextF3);

                            }
                        }

                    }
                    // hot keys
                    else if (cellValue.Contains("FN+F4"))
                    {
                        string fn4 = GetFieldValue(p_machineSheet, "Fn4");
                        string nextF4 = GetFieldValue2(p_machineSheet, "Fn4");

                        if (fn4 != null && fn4 != "")
                        {
                            if (ContainsField(p_machineSheet, "Fn4"))
                            {
                                SetCellValue(testPlanRow, j + 1, "Y");
                                SetCellValue(testPlanRow, j + 3, fn4);

                            }
                        }
                        if ((nextF4 != null && nextF4 != "") && (fn4 == null || fn4 == "N/A" || fn4 == "na" || fn4 == "NA" || fn4 == ""))
                        {
                            if (ContainsField(p_machineSheet, "Fn4"))
                            {
                                SetCellValue(testPlanRow, j + 1, "Y");
                                SetCellValue(testPlanRow, j + 3, nextF4);

                            }
                        }

                    }
                    else if (cellValue.Contains("FN+F5"))
                    {
                        string fn5 = GetFieldValue(p_machineSheet, "Fn5");
                        string nextF5 = GetFieldValue2(p_machineSheet, "Fn5");

                        if (fn5 != null && fn5 != "")
                        {
                            if (ContainsField(p_machineSheet, "Fn5"))
                            {
                                SetCellValue(testPlanRow, j + 1, "Y");
                                SetCellValue(testPlanRow, j + 3, fn5);

                            }
                        }
                        if ((nextF5 != null && nextF5 != "") && (fn5 == null || fn5 == "N/A" || fn5 == "na" || fn5 == "NA" || fn5 == ""))
                        {
                            if (ContainsField(p_machineSheet, "Fn5"))
                            {
                                SetCellValue(testPlanRow, j + 1, "Y");
                                SetCellValue(testPlanRow, j + 3, nextF5);

                            }
                        }

                    }
                    else if (cellValue.Contains("FN+F6"))
                    {
                        string fn6 = GetFieldValue(p_machineSheet, "Fn6");
                        string nextF6 = GetFieldValue2(p_machineSheet, "Fn6");

                        if (fn6 != null && fn6 != "")
                        {
                            if (ContainsField(p_machineSheet, "Fn6"))
                            {
                                SetCellValue(testPlanRow, j + 1, "Y");
                                SetCellValue(testPlanRow, j + 3, fn6);

                            }
                        }
                        if ((nextF6 != null && nextF6 != "") && (fn6 == null || fn6 == "N/A" || fn6 == "na" || fn6 == "NA" || fn6 == ""))
                        {
                            if (ContainsField(p_machineSheet, "Fn6"))
                            {
                                SetCellValue(testPlanRow, j + 1, "Y");
                                SetCellValue(testPlanRow, j + 3, nextF6);

                            }
                        }

                    }
                    else if (cellValue.Contains("FN+F7"))
                    {
                        string fn7 = GetFieldValue(p_machineSheet, "Fn7");
                        string nextF7 = GetFieldValue2(p_machineSheet, "Fn7");

                        if (fn7 != null && fn7 != "")
                        {
                            if (ContainsField(p_machineSheet, "Fn7"))
                            {
                                SetCellValue(testPlanRow, j + 1, "Y");
                                SetCellValue(testPlanRow, j + 3, fn7);

                            }
                        }
                        if ((nextF7 != null && nextF7 != "") && (fn7 == null || fn7 == "N/A" || fn7 == "na" || fn7 == "NA" || fn7 == ""))
                        {
                            if (ContainsField(p_machineSheet, "Fn7"))
                            {
                                SetCellValue(testPlanRow, j + 1, "Y");
                                SetCellValue(testPlanRow, j + 3, nextF7);

                            }
                        }

                    }
                    else if (cellValue.Contains("FN+F8"))
                    {
                        string fn8 = GetFieldValue(p_machineSheet, "Fn8");
                        string nextF8 = GetFieldValue2(p_machineSheet, "Fn8");

                        if (fn8 != null && fn8 != "")
                        {
                            if (ContainsField(p_machineSheet, "Fn8"))
                            {
                                SetCellValue(testPlanRow, j + 1, "Y");
                                SetCellValue(testPlanRow, j + 3, fn8);

                            }
                        }
                        if ((nextF8 != null && nextF8 != "") && (fn8 == null || fn8 == "N/A" || fn8 == "na" || fn8 == "NA" || fn8 == ""))
                        {
                            if (ContainsField(p_machineSheet, "Fn8"))
                            {
                                SetCellValue(testPlanRow, j + 1, "Y");
                                SetCellValue(testPlanRow, j + 3, nextF8);

                            }
                        }

                    }
                    else if (cellValue.Contains("FN+F9"))
                    {
                        string fn9 = GetFieldValue(p_machineSheet, "Fn9");
                        string nextF9 = GetFieldValue2(p_machineSheet, "Fn9");

                        if (fn9 != null && fn9 != "")
                        {
                            if (ContainsField(p_machineSheet, "Fn9"))
                            {
                                SetCellValue(testPlanRow, j + 1, "Y");
                                SetCellValue(testPlanRow, j + 3, fn9);

                            }
                        }
                        if ((nextF9 != null && nextF9 != "") && (fn9 == null || fn9 == "N/A" || fn9 == "na" || fn9 == "NA" || fn9 == ""))
                        {
                            if (ContainsField(p_machineSheet, "Fn9"))
                            {
                                SetCellValue(testPlanRow, j + 1, "Y");
                                SetCellValue(testPlanRow, j + 3, nextF9);

                            }
                        }

                    }
                    else if (cellValue.Contains("FN+F10"))
                    {
                        string fn10 = GetFieldValue(p_machineSheet, "Fn10");
                        string nextF10 = GetFieldValue2(p_machineSheet, "Fn10");

                        if (fn10 != null && fn10 != "")
                        {
                            if (ContainsField(p_machineSheet, "Fn10"))
                            {
                                SetCellValue(testPlanRow, j + 1, "Y");
                                SetCellValue(testPlanRow, j + 3, fn10);

                            }
                        }
                        if ((nextF10 != null && nextF10 != "") && (fn10 == null || fn10 == "N/A" || fn10 == "na" || fn10 == "NA" || fn10 == ""))
                        {
                            if (ContainsField(p_machineSheet, "Fn10"))
                            {
                                SetCellValue(testPlanRow, j + 1, "Y");
                                SetCellValue(testPlanRow, j + 3, nextF10);

                            }
                        }

                    }
                    else if (cellValue.Contains("FN+F11"))
                    {
                        string fn11 = GetFieldValue(p_machineSheet, "Fn11");
                        string nextF11 = GetFieldValue2(p_machineSheet, "Fn11");

                        if (fn11 != null && fn11 != "")
                        {
                            if (ContainsField(p_machineSheet, "Fn11"))
                            {
                                SetCellValue(testPlanRow, j + 1, "Y");
                                SetCellValue(testPlanRow, j + 3, fn11);

                            }
                        }
                        if ((nextF11 != null && nextF11 != "") && (fn11 == null || fn11 == "N/A" || fn11 == "na" || fn11 == "NA" || fn11 == ""))
                        {
                            if (ContainsField(p_machineSheet, "Fn11"))
                            {
                                SetCellValue(testPlanRow, j + 1, "Y");
                                SetCellValue(testPlanRow, j + 3, nextF11);

                            }
                        }

                    }
                    else if (cellValue.Contains("FN+F12"))
                    {
                        string fn12 = GetFieldValue(p_machineSheet, "Fn12");
                        string nextF12 = GetFieldValue2(p_machineSheet, "Fn12");

                        if (fn12 != null && fn12 != "")
                        {
                            if (ContainsField(p_machineSheet, "Fn12"))
                            {
                                SetCellValue(testPlanRow, j + 1, "Y");
                                SetCellValue(testPlanRow, j + 3, fn12);

                            }
                        }
                        if ((nextF12 != null && nextF12 != "") && (fn12 == null || fn12 == "N/A" || fn12 == "na" || fn12 == "NA" || fn12 == ""))
                        {
                            if (ContainsField(p_machineSheet, "Fn12"))
                            {
                                SetCellValue(testPlanRow, j + 1, "Y");
                                SetCellValue(testPlanRow, j + 3, nextF12);

                            }
                        }

                    }
                    else if (cellValue.Equals("FN+UP"))
                    {
                        string fnUp = GetFieldValue(p_machineSheet, "Fn+up");
                        string nextFnUp = GetFieldValue2(p_machineSheet, "Fn+up");

                        if (fnUp != null && fnUp != "")
                        {
                            if (ContainsField(p_machineSheet, "Fn+up"))
                            {
                                SetCellValue(testPlanRow, j + 1, "Y");
                                SetCellValue(testPlanRow, j + 3, fnUp);

                            }
                        }
                        if ((nextFnUp != null && nextFnUp != "") && (nextFnUp == null || nextFnUp == "N/A" || nextFnUp == "na" || nextFnUp == "NA" || nextFnUp == ""))
                        {
                            if (ContainsField(p_machineSheet, "Fn+up"))
                            {
                                SetCellValue(testPlanRow, j + 1, "Y");
                                SetCellValue(testPlanRow, j + 3, nextFnUp);

                            }
                        }

                    }
                    else if (cellValue.Equals("FN+down"))
                    {
                        string fnDown = GetFieldValue(p_machineSheet, "Fn+down");
                        string nextFnDown = GetFieldValue2(p_machineSheet, "Fn+down");

                        if (fnDown != null && fnDown != "")
                        {
                            if (ContainsField(p_machineSheet, "Fn+down"))
                            {
                                SetCellValue(testPlanRow, j + 1, "Y");
                                SetCellValue(testPlanRow, j + 3, fnDown);

                            }
                        }
                        if ((nextFnDown != null && nextFnDown != "") && (nextFnDown == null || nextFnDown == "N/A" || nextFnDown == "na" || nextFnDown == "NA" || nextFnDown == ""))
                        {
                            if (ContainsField(p_machineSheet, "Fn+down"))
                            {
                                SetCellValue(testPlanRow, j + 1, "Y");
                                SetCellValue(testPlanRow, j + 3, nextFnDown);

                            }
                        }

                    }
                    else if (cellValue.Equals("FN+left"))
                    {
                        string fnLeft = GetFieldValue(p_machineSheet, "Fn+left");
                        string nextFnLeft = GetFieldValue2(p_machineSheet, "Fn+left");

                        if (fnLeft != null && fnLeft != "")
                        {
                            if (ContainsField(p_machineSheet, "Fn+left"))
                            {
                                SetCellValue(testPlanRow, j + 1, "Y");
                                SetCellValue(testPlanRow, j + 3, fnLeft);

                            }
                        }
                        if ((nextFnLeft != null && nextFnLeft != "") && (nextFnLeft == null || nextFnLeft == "N/A" || nextFnLeft == "na" || nextFnLeft == "NA" || nextFnLeft == ""))
                        {
                            if (ContainsField(p_machineSheet, "Fn+left"))
                            {
                                SetCellValue(testPlanRow, j + 1, "Y");
                                SetCellValue(testPlanRow, j + 3, nextFnLeft);

                            }
                        }

                    }
                    else if (cellValue.Equals("FN+right"))
                    {
                        string fnRight = GetFieldValue(p_machineSheet, "Fn+right");
                        string nextFnRight = GetFieldValue2(p_machineSheet, "Fn+right");

                        if (fnRight != null && fnRight != "")
                        {
                            if (ContainsField(p_machineSheet, "Fn+right"))
                            {
                                SetCellValue(testPlanRow, j + 1, "Y");
                                SetCellValue(testPlanRow, j + 3, fnRight);

                            }
                        }
                        if ((nextFnRight != null && nextFnRight != "") && (nextFnRight == null || nextFnRight == "N/A" || nextFnRight == "na" || nextFnRight == "NA" || nextFnRight == ""))
                        {
                            if (ContainsField(p_machineSheet, "Fn+left"))
                            {
                                SetCellValue(testPlanRow, j + 1, "Y");
                                SetCellValue(testPlanRow, j + 3, nextFnRight);

                            }
                        }

                    }
                    else if (cellValue.Equals("FN+end"))
                    {
                        string fnEnd = GetFieldValue(p_machineSheet, "Fn+end");
                        string nextFnEnd = GetFieldValue2(p_machineSheet, "Fn+end");

                        if (fnEnd != null && fnEnd != "")
                        {
                            if (ContainsField(p_machineSheet, "Fn+end"))
                            {
                                SetCellValue(testPlanRow, j + 1, "Y");
                                SetCellValue(testPlanRow, j + 3, fnEnd);

                            }
                        }
                        if ((nextFnEnd != null && nextFnEnd != "") && (nextFnEnd == null || nextFnEnd == "N/A" || nextFnEnd == "na" || nextFnEnd == "NA" || nextFnEnd == ""))
                        {
                            if (ContainsField(p_machineSheet, "Fn+end"))
                            {
                                SetCellValue(testPlanRow, j + 1, "Y");
                                SetCellValue(testPlanRow, j + 3, nextFnEnd);

                            }
                        }

                    }
                    // soft
                    else if (cellValue.Contains("Device check"))
                    {                     
                        SetCellValue(testPlanRow, j + 1, "Y");

                    }
                    else if (cellValue.Contains("BIOS/EC/VBIOS check"))
                    {
                        SetCellValue(testPlanRow, j + 1, "Y");

                    }


                    else if (cellValue.Contains("KB LED FW check"))
                    {
                        if (ContainsField(p_machineSheet, "KB Backlight"))
                        {
                            SetCellValue(testPlanRow, j + 1, "Y");
                        }
                       

                    }
                    else if (cellValue.Contains("ME check"))
                    {
                        SetCellValue(testPlanRow, j + 1, "Y");

                    }
                    else if (cellValue.Contains("Device check"))
                    {
                        SetCellValue(testPlanRow, j + 1, "Y");

                    }
                    else if (cellValue.Contains("PTT check"))
                    {
                        SetCellValue(testPlanRow, j + 1, "Y");

                    }

                    else if (cellValue.Contains("Vpro"))
                    {
                        // string vramValue = GetFieldValue(p_machineSheet, "vRAM");
                        if (ContainsField(p_machineSheet, "vPro"))
                        {
                            SetCellValue(testPlanRow, j + 1, "Y");
                        }

                    }
                    else if (cellValue.Contains("lan loop"))
                    {
                        if (EqualsField(p_machineSheet, "LAN"))
                        {
                            SetCellValue(testPlanRow, j + 1, "Y");
                            SetLANValue(testPlan);
                        }
                        // SetCellValue(testPlanRow, j + 1, "Y");

                    }
                    else if (cellValue.Contains("KB Backlight"))
                    {
                        // string vramValue = GetFieldValue(p_machineSheet, "vRAM");
                        if (ContainsField(p_machineSheet, "Backlight"))
                        {
                            SetKeyboardValue(testPlan);
                        }
                        SetCellValue(testPlanRow, j + 1, "Y");

                    }
                    else if (cellValue.Contains("touch pad"))
                    {
                        // string vramValue = GetFieldValue(p_machineSheet, "vRAM");
                        if (EqualsField(p_machineSheet, "Form Factor")||EqualsField(p_machineSheet, "Touch Type") || EqualsField(p_machineSheet, "Touch panel"))
                        {
                            SetCellValue(testPlanRow, j + 1, "Y");
                        }
                      

                    }
                    else if (cellValue.Equals("Finger print HW ID"))
                    {
                        // string vramValue = GetFieldValue(p_machineSheet, "vRAM");
                        if (EqualsField(p_machineSheet, "Form Factor")|| EqualsField(p_machineSheet, "Touch Type"))
                        {
                            SetCellValue(testPlanRow, j + 1, "Y");
                        }

                    }
                    else if (cellValue.Equals("Support Touch panel"))
                    {
                        // string vramValue = GetFieldValue(p_machineSheet, "vRAM");
                        if (EqualsField(p_machineSheet, "Touch panel"))
                        {
                            SetCellValue(testPlanRow, j + 1, "Y");
                        }

                    }
                    // power
                    else if (cellValue.Equals("S3睡眠"))
                    {
                        SetCellValue(testPlanRow, j + 1, "Y");
                    }
                    else if (cellValue.Equals("LID"))
                    {
                        SetCellValue(testPlanRow, j + 1, "Y");
                    }
                    // I/O ports
                    else if (cellValue.Equals("TypeA"))
                    {
                        // 统计 TypeA 或 TypeC 后面单元格有值的数量
                        for (int k = 0; k <= p_machineSheet.LastRowNum; k++)
                        {
                            IRow row = p_machineSheet.GetRow(k);
                            if (row == null) continue;

                            for (int d = 0; d < row.LastCellNum; d++)
                            {
                                ICell cell1 = row.GetCell(d);
                                if (cell1 == null) continue;

                                string cellValue1 = cell1.ToString();
                                if (cellValue1.StartsWith("TypeA"))
                                {
                                    ICell nextCell = row.GetCell(d + 1);
                                    if (nextCell != null && !string.IsNullOrEmpty(nextCell.ToString())&& nextCell.ToString() != "NA" && nextCell.ToString() != "N/A" && nextCell.ToString() != "na")
                                    {                                    
                                        typeA++;
                                       
                                    }
                                    SetCellValue(testPlanRow, j + 3, typeA.ToString());
                                    SetCellValue(testPlanRow, j + 1, "Y");
                                }
                               
                                if (cellValue1.StartsWith("Type-A"))
                                {
                                    ICell nextCell = row.GetCell(d + 2);
                                    int indexOfX = nextCell.ToString().IndexOf('x');
                                    if (indexOfX > 0)
                                    {
                                        string extractedValue = nextCell.ToString().Substring(0, indexOfX);
                                        // 在这里对提取出来的内容进行处理
                                        SetCellValue(testPlanRow, j + 1, "Y");

                                        SetCellValue(testPlanRow, j + 3, extractedValue.ToString());
                                    }
                                }
                               

                            }
                        }
                    }
                    else if (cellValue.Equals("TypeC"))
                    {
                        // 统计 TypeA 或 TypeC 后面单元格有值的数量
                        for (int k = 0; k <= p_machineSheet.LastRowNum; k++)
                        {
                            IRow row = p_machineSheet.GetRow(k);
                            if (row == null) continue;

                            for (int d = 0; d < row.LastCellNum; d++)
                            {
                                ICell cell1 = row.GetCell(d);
                                if (cell1 == null) continue;

                                string cellValue1 = cell1.ToString();
                                if (cellValue1.StartsWith("TypeC"))
                                {
                                    ICell nextCell = row.GetCell(d + 1);
                                    if (nextCell != null && !string.IsNullOrEmpty(nextCell.ToString()) && nextCell.ToString() != "NA" && nextCell.ToString() != "N/A" && nextCell.ToString() != "na")
                                    {
                                        typeC++;
                                       

                                    }
                                    SetCellValue(testPlanRow, j + 3, typeC.ToString());
                                    SetCellValue(testPlanRow, j + 1, "Y");
                                    SetTypecValue(testPlan);
                                }

                                if (cellValue1.StartsWith("Type-C"))
                                {
                                    ICell nextCell = row.GetCell(d + 2);
                                    int indexOfX = nextCell.ToString().IndexOf('x');
                                    if (indexOfX > 0)
                                    {
                                        string extractedValue = nextCell.ToString().Substring(0, indexOfX);
                                        // 在这里对提取出来的内容进行处理
                                        SetCellValue(testPlanRow, j + 1, "Y");

                                        SetCellValue(testPlanRow, j + 3, extractedValue.ToString());
                                        SetTypecValue(testPlan);
                                    }
                                }


                            }
                        }
                        /*if (ContainsField(p_machineSheet, "TypeC") || ContainsField(p_machineSheet, "Type-C"))
                        {                    
                            SetTypecValue(testPlan);
                            SetCellValue(testPlanRow, j + 3, typeC.ToString());
                        }
                        SetCellValue(testPlanRow, j + 1, "Y");
                         */
                    

                    }


                    else if (cellValue.Equals("HDMI"))
                    {
                        string hdmi = GetFieldValueEqual(p_machineSheet, "HDMI");
                        if (EqualsField(p_machineSheet, "HDMI"))
                        {
                            SetCellValue(testPlanRow, j + 1, "Y");
                            SetCellValue(testPlanRow, j + 3, hdmi);
                        }
                      
                    }
                    else if (cellValue.Equals("Type-c to HDMI")) //问题
                    {
                        if (ContainsField(p_machineSheet, "TypeC"))
                        {
                            SetCellValue(testPlanRow, j + 1, "Y");
                        }
                     
                    }
                    else if (cellValue.Equals("loop back"))
                    {
                        if (ContainsField(p_machineSheet, "TypeC-1")|| ContainsField(p_machineSheet, "TypeC-2")|| ContainsField(p_machineSheet, "TypeC-3")|| ContainsField(p_machineSheet, "TypeC-4"))
                        {
                            SetCellValue(testPlanRow, j + 1, "Y");
                        }
                       
                    }
                    else if (cellValue.Equals("CRT"))
                    {
                        if (EqualsField(p_machineSheet, "CRT"))
                        {
                            SetCellValue(testPlanRow, j + 1, "Y");
                        }
                       
                    }
                    else if (cellValue.Equals("Mini Display Port"))
                    {
                        if (EqualsField(p_machineSheet, "Mini Display Port"))
                        {
                            SetCellValue(testPlanRow, j + 1, "Y");
                        }
                       
                    }
                    else if (cellValue.Equals("COM PORT"))
                    {
                        if (EqualsField(p_machineSheet, "COM PORT"))
                        {
                            SetCellValue(testPlanRow, j + 1, "Y");
                        }

                    }
                    else if (cellValue.Equals("SIM card"))
                    {
                        if (EqualsField(p_machineSheet, "SIM"))
                        {
                            SetCellValue(testPlanRow, j + 1, "Y");
                        }

                    }
                    else if (cellValue.Equals("Card Reader"))
                    {
                        if (EqualsField(p_machineSheet, "Card Reader"))
                        {
                            SetCellValue(testPlanRow, j + 1, "Y");
                        }

                    }
                   
                    // Sensors 为NA的处理
                    else if (cellValue.Equals("Accelerometer"))
                    {
                        if (EqualsField(p_machineSheet, "Accelerometer"))
                        {
                            SetCellValue(testPlanRow, j + 1, "Y");
                        }

                    }
                    else if (cellValue.Equals("Magnetometer"))
                    {
                        if (EqualsField(p_machineSheet, "Magnetometer"))
                        {
                            SetCellValue(testPlanRow, j + 1, "Y");
                        }

                    }
                    else if (cellValue.Equals("Gyroscope"))
                    {
                        if (EqualsField(p_machineSheet, "Gyroscope"))
                        {
                            SetCellValue(testPlanRow, j + 1, "Y");
                        }

                    }
                    else if (cellValue.Equals("GPS"))
                    {
                        if (EqualsField(p_machineSheet, "GPS"))
                        {
                            SetCellValue(testPlanRow, j + 1, "Y");
                        }

                    }
                    else if (cellValue.Equals("Ambient Light Sensor"))
                    {
                        if (EqualsField(p_machineSheet, "Ambient Light Sensor"))
                        {
                            SetCellValue(testPlanRow, j + 1, "Y");
                        }

                    }



                }
            }


            // 移除空行
            EmptyRows();
            // 合并
            MergeCell();



            // 插入键盘新行
            // 查找并保存特殊格式的单元格
            // 查找并保存特殊格式的单元格及其后面的值
            //List<Tuple<string, string>> specialCellsAndValues = FindSpecialCellsAndValues(p_machineSheet);

            // 将这些值添加到 testPlan 表中
            // AddValuesToTestPlan(testPlan, specialCellsAndValues);

            /*void ProcessKeyValueEql(IRow testPlanRow, string cellValue, string key, string fieldName, int currentColumnIndex)
            {
                if (cellValue.Equals(key))
                {
                    if (EqualsField(p_machineSheet, fieldName))
                    {
                        SetCellValue(testPlanRow, currentColumnIndex + 1, "Y");
                    }
                }
            }
            void ProcessKeyValueCon(IRow testPlanRow, string cellValue, string key, string fieldName, int currentColumnIndex)
            {
                if (cellValue.Contains(key))
                {
                    if (EqualsField(p_machineSheet, fieldName))
                    {
                        SetCellValue(testPlanRow, currentColumnIndex + 1, "Y");
                    }
                }
            }
            */
            bool ContainsField(ISheet sheet, string fieldName)
            {
                // 遍历所有行
                for (int k = 0; k <= sheet.LastRowNum; k++)
                {
                    IRow row = sheet.GetRow(k);
                    if (row == null) continue; // 跳过空行

                    // 遍历当前行的所有单元格
                    for (int d = 0; d < row.LastCellNum; d++)
                    {
                        if (row.GetCell(d)?.ToString().Contains(fieldName) == true)
                        {
                            // 获取当前单元格的下一个单元格
                            ICell nextCell = row.GetCell(d + 1);
                            ICell nextCell2 = row.GetCell((d + 2));
                            // 统计 TypeA 或 TypeC 后面单元格有值的数量
                            /*if (fieldName.StartsWith("TypeA"))
                            {
                                if (nextCell != null && !string.IsNullOrEmpty(nextCell.ToString()))
                                {
                                    typeA++;
                                }
                            }
                            else if (fieldName.StartsWith("TypeC"))
                            {
                                if (nextCell != null && !string.IsNullOrEmpty(nextCell.ToString()))
                                {
                                    typeC++;
                                }
                            }
                            */
                            if (nextCell == null)
                            {
                            }
                            else if (nextCell.ToString().Equals("N/A") || nextCell.ToString().Equals("na") || nextCell.ToString().Equals("NA"))
                            {
                                return false;
                            }

                            if ((nextCell == null ||
                              nextCell.ToString().Equals("N/A") || nextCell.ToString().Equals("na") || nextCell.ToString().Equals("NA") || nextCell.ToString().Equals("")) && (nextCell2 == null ||
                              nextCell2.ToString().Equals("N/A") || nextCell2.ToString().Equals("na") || nextCell2.ToString().Equals("NA") || nextCell2.ToString().Equals("")))

                            {

                                return false;
                            }
                            //if ((nextCell == null || nextCell.ToString().Equals("N/A") || nextCell.ToString().Equals("na") || nextCell.ToString().Equals("NA")))
                            //{
                            //    Console.WriteLine("123");
                            //}
                            if ((nextCell == null ||
                               nextCell.ToString().Equals("N/A") || nextCell.ToString().Equals("na") || nextCell.ToString().Equals("NA")) &&
                               (nextCell2.ToString() != "" || nextCell2.ToString() != "N/A" || nextCell2.ToString() != "na" || nextCell2.ToString() != "NA"))
                            //&& nextCell2.ToString().Equals("N/A") || nextCell2.ToString().Equals("na") || nextCell2.ToString().Equals("NA")

                            {
                                return true;
                            }


                            // 检查下一个单元格是否为空、为"N/A"或为空字符串


                            // 检查下一个单元格是否包含"PD"或"DP"或"Per"
                            if (fieldName.ToString().Contains("KB Backlight") && !nextCell.ToString().Contains("Per"))
                            {
                                return false;
                            }
                            if (fieldName.ToString().Contains("Memory Slot"))
                            {
                                // Memory Slot 的逻辑
                                if (!nextCell.ToString().Contains("Onboard") || !nextCell.ToString().Contains("onboard"))
                                {
                                    // 不是 onboard，检查下一个单元格的值是否为非空整数
                                    if (int.TryParse(nextCell.ToString(), out int intValue))
                                    {
                                        return true;
                                    }
                                    else
                                    {
                                        return false;
                                    }
                                }
                            }
                            /*if (fieldName.ToString().Contains("Memory Slot") &&( !nextCell.ToString().Contains("Onboard") || !nextCell.ToString().Contains("onboard")))
                            {
                                return true;
                            }*/
                            if ((fieldName.ToString().Contains("TypeC-1") || fieldName.ToString().Contains("TypeC-2") || fieldName.ToString().Contains("TypeC-3") || fieldName.ToString().Contains("TypeC-4"))
                               && (!nextCell.ToString().Contains("TBT")))
                            {
                                return false;
                            }
                            if ((fieldName.ToString().Contains("TypeC-1") || fieldName.ToString().Contains("TypeC-2") || fieldName.ToString().Contains("TypeC-3") || fieldName.ToString().Contains("TypeC-4"))
                                && (!nextCell.ToString().Contains("PD")))
                            {
                                return false;
                            }
                            if (fieldName.ToString().Contains("GPU-1") && nextCell.ToString().Contains("UMA")|| nextCell.ToString().Contains("AMD"))
                            {
                                return false;
                            }



                            return true;
                        }
                    }
                }
                // 如果没有找到匹配的fieldName，返回false
                return false;
            }
            // 插入新行
            void InsertEmptyRows(ISheet sheet, int numberOfRows)
            {
                // 向下移动现有行
                sheet.ShiftRows(0, sheet.LastRowNum, numberOfRows);

                // 在顶部插入空行
                for (int i = 0; i < numberOfRows; i++)
                {
                    IRow newRow = sheet.CreateRow(i);
                }
                IRow row0 = testPlan.CreateRow(0);
                row0.CreateCell(0, CellType.String).SetCellValue("PCBA NPI機種測試test plan");

                // 设置单元格样式
                ICellStyle style = testPlan.Workbook.CreateCellStyle();
                style.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.LightGreen.Index;
                style.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
                style.VerticalAlignment = VerticalAlignment.Center;
                style.FillPattern = FillPattern.SolidForeground; // 设置填充模式为实心填充
                                                                 // 添加设置字体大小的代码
                IFont font = testPlan.Workbook.CreateFont();
                font.FontHeightInPoints = 12; // 设置字体大小为 12 磅
                style.SetFont(font);
                row0.GetCell(0).CellStyle = style;

                row0.HeightInPoints = 30;
                // 合并A1到A5单元格
                CellRangeAddress mergeRegion = new CellRangeAddress(0, 0, 0, 4);
                testPlan.AddMergedRegion(mergeRegion);

                //CellRangeAddress mergeRegion2 = new CellRangeAddress(1, 1, 0, 4);
                //testPlan.AddMergedRegion(mergeRegion2);

                IRow row1 = testPlan.CreateRow(1);
                row1.CreateCell(1, CellType.String).SetCellValue("NPI機種：");

            }
            // 合并删减后的单元格
            void MergeCell()
            {
                // 假设 testPlan 是你的工作表
                int lastMergeStart = -1; // 用于记录最后一次合并的起始行
                int lastMergeEnd = -1;   // 用于记录最后一次合并的结束行

                for (int rowIndex = 0; rowIndex <= testPlan.LastRowNum; rowIndex++)
                {
                    IRow row = testPlan.GetRow(rowIndex);
                    if (row == null) continue;

                    // 检查当前行是否在一个合并区域内
                    bool isMerged = false;
                    for (int i = 0; i < testPlan.NumMergedRegions; i++)
                    {
                        CellRangeAddress mergedRegion = testPlan.GetMergedRegion(i);
                        if (mergedRegion.IsInRange(rowIndex, 0))
                        {
                            isMerged = true;
                            break;
                        }
                    }

                    if (isMerged) continue; // 如果当前行在一个合并区域内，跳过

                    ICell currentCell = row.GetCell(0); // 获取第1列的单元格

                    if (currentCell != null && !string.IsNullOrWhiteSpace(currentCell.ToString()))
                    {
                        int mergeStart = rowIndex;
                        int mergeEnd = rowIndex;

                        // 查找下一个有值的单元格
                        for (int checkIndex = rowIndex + 1; checkIndex <= testPlan.LastRowNum; checkIndex++)
                        {
                            IRow nextRow = testPlan.GetRow(checkIndex);
                            if (nextRow == null) continue;

                            ICell nextCell = nextRow.GetCell(0);

                            // 检查下一个单元格是否属于已合并区域
                            bool nextCellIsMerged = false;
                            for (int i = 0; i < testPlan.NumMergedRegions; i++)
                            {
                                CellRangeAddress mergedRegion = testPlan.GetMergedRegion(i);
                                if (mergedRegion.IsInRange(checkIndex, 0))
                                {
                                    nextCellIsMerged = true;
                                    break;
                                }
                            }

                            if (nextCellIsMerged)
                            {
                                mergeEnd = checkIndex - 1;
                                break;
                            }

                            if (nextCell != null && !string.IsNullOrWhiteSpace(nextCell.ToString()))
                            {
                                mergeEnd = checkIndex - 1;
                                break;
                            }
                            if (checkIndex == testPlan.LastRowNum) // 如果检查到最后一行
                            {
                                mergeEnd = checkIndex;
                            }
                        }

                        // 如果找到的mergeEnd比当前行号大，则进行合并
                        if (mergeEnd > mergeStart)
                        {
                            testPlan.AddMergedRegion(new CellRangeAddress(mergeStart, mergeEnd, 0, 0));
                            rowIndex = mergeEnd; // 跳过合并过的行
                            lastMergeStart = -1; // 重置最后合并的起始行
                            lastMergeEnd = -1;   // 重置最后合并的结束行
                        }
                        else
                        {
                            lastMergeStart = mergeStart;
                            lastMergeEnd = mergeEnd;
                        }
                    }
                }

                // 在循环结束后，检查是否有未合并的单元格
                if (lastMergeStart != -1 && lastMergeEnd > lastMergeStart)
                {
                    testPlan.AddMergedRegion(new CellRangeAddress(lastMergeStart, lastMergeEnd, 0, 0));
                }
            }
            void EmptyRows()
            {
                for (int rowIndex = 4; rowIndex <= testPlan.LastRowNum; rowIndex++)
                {
                    IRow row = testPlan.GetRow(rowIndex);
                    if (row == null) continue;

                    bool hasEmptyCell = false;

                    for (int colIndex = 2; colIndex < row.LastCellNum; colIndex++) // 从第 3 列开始
                    {
                        ICell cell = row.GetCell(colIndex);
                        if (cell == null || string.IsNullOrWhiteSpace(cell.ToString()))
                        {
                            hasEmptyCell = true;
                            break;
                        }
                    }

                    if (hasEmptyCell)
                    {

                        testPlan.RemoveRow(row);

                        // 将后面的行上移
                        if (rowIndex < testPlan.LastRowNum)
                        {
                            testPlan.ShiftRows(rowIndex + 1, testPlan.LastRowNum, -1);
                        }
                        rowIndex--;
                    }
                }
            }
            bool EqualsField(ISheet sheet, string fieldName)
            {
                for (int k = 0; k <= sheet.LastRowNum; k++)
                {
                    IRow row = sheet.GetRow(k);
                    if (row == null) continue;

                    for (int d = 0; d < row.LastCellNum; d++)
                    {
                        if (row.GetCell(d)?.ToString().Equals(fieldName) == true)
                        {
                            ICell nextCell = row.GetCell(d + 1);
                            if (nextCell == null ||
                               string.IsNullOrWhiteSpace(nextCell.ToString()) ||
                               nextCell.ToString().Equals("N/A") || nextCell.ToString().Equals("na") || nextCell.ToString().Equals("NA"))
                            {
                                return false;
                            }
                            if ((fieldName.ToString().Equals("Touch Type") || fieldName.ToString().Equals("Touch Cell Type")) && !nextCell.ToString().Contains("Finger"))
                            {
                                return false;
                            }


                            return true;

                        }
                    }
                }
                return false;
            }
            void RemoveEmptyColumns(ISheet sheet)
            {
                int maxColIndex = 0;

                // 先找出最大列索引
                for (int rowIndex = 0; rowIndex <= sheet.LastRowNum; rowIndex++)
                {
                    IRow row = sheet.GetRow(rowIndex);
                    if (row != null && row.LastCellNum > maxColIndex)
                    {
                        maxColIndex = row.LastCellNum;
                    }
                }

                // 记录空列索引
                List<int> emptyColumnIndexes = new List<int>();

                // 先遍历一次，找到所有空列的索引
                for (int colIndex = 0; colIndex < maxColIndex; colIndex++)
                {
                    bool isColumnEmpty = true;
                    for (int rowIndex = 0; rowIndex <= sheet.LastRowNum; rowIndex++)
                    {
                        IRow row = sheet.GetRow(rowIndex);
                        if (row == null) continue;

                        ICell cell = row.GetCell(colIndex);
                        if (cell != null && !string.IsNullOrWhiteSpace(cell.ToString()))
                        {
                            isColumnEmpty = false;
                            break;
                        }
                    }

                    if (isColumnEmpty)
                    {
                        emptyColumnIndexes.Add(colIndex);
                    }
                }

                // 反向遍历空列索引并删除空列
                emptyColumnIndexes.Sort((a, b) => b.CompareTo(a)); // 降序排序
                foreach (int colIndex in emptyColumnIndexes)
                {
                    for (int rowIndex = 0; rowIndex <= sheet.LastRowNum; rowIndex++)
                    {
                        IRow row = sheet.GetRow(rowIndex);
                        if (row == null) continue;

                        ICell cell = row.GetCell(colIndex);
                        if (cell != null)
                        {
                            row.RemoveCell(cell);
                        }

                        // 左移后面的列
                        for (int colShift = colIndex; colShift < row.LastCellNum - 1; colShift++)
                        {
                            ICell leftCell = row.GetCell(colShift);
                            ICell rightCell = row.GetCell(colShift + 1);

                            if (rightCell != null)
                            {
                                if (leftCell == null)
                                {
                                    leftCell = row.CreateCell(colShift);
                                }
                                leftCell.SetCellValue(rightCell.ToString());
                                row.RemoveCell(rightCell);
                            }
                        }
                    }
                }
            }
            string GetFieldValue(ISheet sheet, string fieldName)
            {
                for (int k = 0; k <= sheet.LastRowNum; k++)
                {
                    IRow row = sheet.GetRow(k);
                    if (row == null) continue;

                    for (int d = 0; d < row.LastCellNum; d++)
                    {                     
                        if (row.GetCell(d)?.ToString().Contains(fieldName) == true)
                        {
                            ICell nextCell = row.GetCell(d + 1);
                            return nextCell?.ToString();
                        }
                    }
                }
                return null;
            }
            string GetFieldValueEqual(ISheet sheet, string fieldName)
            {
                for (int k = 0; k <= sheet.LastRowNum; k++)
                {
                    IRow row = sheet.GetRow(k);
                    if (row == null) continue;

                    for (int d = 0; d < row.LastCellNum; d++)
                    {
                        if (row.GetCell(d)?.ToString().Equals(fieldName) == true)
                        {
                            ICell nextCell = row.GetCell(d + 1);
                            return nextCell?.ToString();
                        }
                    }
                }
                return null;
            }
            string GetFieldValue2(ISheet sheet, string fieldName)
            {
                for (int k = 0; k <= sheet.LastRowNum; k++)
                {
                    IRow row = sheet.GetRow(k);
                    if (row == null) continue;

                    for (int d = 0; d < row.LastCellNum; d++)
                    {
                        if (row.GetCell(d)?.ToString().Contains(fieldName) == true)
                        {
                            ICell nextCell2 = row.GetCell(d + 2);
                            return nextCell2?.ToString();
                        }
                    }
                }
                return null;
            }

            void SetCellValue(IRow row, int cellIndex, string value)
            {
                ICell cell = row.GetCell(cellIndex) ?? row.CreateCell(cellIndex);
                cell.SetCellValue(value);
            }

            // 调用每个方法
            void SetProcessorPlatValue(ISheet ItestPlan)
            {
                for (int i = 0; i <= ItestPlan.LastRowNum; i++)
                {
                    IRow row = ItestPlan.GetRow(i);
                    if (row == null) continue;

                    for (int j = 0; j < row.LastCellNum; j++)
                    {
                        ICell cell = row.GetCell(j);
                        if (cell == null) continue;

                        string cellValue = cell.ToString();
                        if (cellValue.Equals("Processor Plat"))
                        {
                            SetCellValue(row, j + 1, "Y");
                        }
                       

                    }
                }
            }

            void SetRamValue(ISheet ItestPlan)
            {
                for (int i = 0; i <= ItestPlan.LastRowNum; i++)
                {
                    IRow row = ItestPlan.GetRow(i);
                    if (row == null) continue;

                    for (int j = 0; j < row.LastCellNum; j++)
                    {
                        ICell cell = row.GetCell(j);
                        if (cell == null) continue;

                        string cellValue = cell.ToString();
                        if (cellValue.Equals("VRAM大小"))
                        {
                            SetCellValue(row, j + 1, "Y");
                        }
                        if (cellValue.Contains("GPUZ"))
                        {
                            SetCellValue(row, j + 1, "Y");
                        }
                        if (cellValue.Contains("modern check"))
                        {
                            SetCellValue(row, j + 1, "Y");
                        }

                    }
                }
            }
            void SetCapacity(ISheet ItestPlan)
            {
                for (int i = 0; i <= ItestPlan.LastRowNum; i++)
                {
                    IRow row = ItestPlan.GetRow(i);
                    if (row == null) continue;

                    for (int j = 0; j < row.LastCellNum; j++)
                    {
                        ICell cell = row.GetCell(j);
                        if (cell == null) continue;

                        string cellValue = cell.ToString();
                        if (cellValue.Equals("Maximum Capacity"))
                        {
                            SetCellValue(row, j + 1, "Y");
                        } 

                    }
                }
            }
            void SetTypecValue(ISheet ItestPlan)
            {
                for (int i = 0; i <= ItestPlan.LastRowNum; i++)
                {
                    IRow row = ItestPlan.GetRow(i);
                    if (row == null) continue;

                    for (int j = 0; j < row.LastCellNum; j++)
                    {
                        ICell cell = row.GetCell(j);
                        if (cell == null) continue;

                        string cellValue = cell.ToString();
                        if (cellValue.Equals("Type_C FW check"))
                        {
                            SetCellValue(row, j + 1, "Y");
                        }                   

                    }
                }
            }
            void SetAudioValue(ISheet ItestPlan)
            {
                for (int i = 0; i <= ItestPlan.LastRowNum; i++)
                {
                    IRow row = ItestPlan.GetRow(i);
                    if (row == null) continue;

                    for (int j = 0; j < row.LastCellNum; j++)
                    {
                        ICell cell = row.GetCell(j);
                        if (cell == null) continue;

                        string cellValue = cell.ToString();
                        if (cellValue.Equals("Headphone mic"))
                        {
                            SetCellValue(row, j + 1, "Y");
                        }                       

                    }
                }
            }

            void SetLANValue(ISheet ItestPlan)
            {
                for (int i = 0; i <= ItestPlan.LastRowNum; i++)
                {
                    IRow row = ItestPlan.GetRow(i);
                    if (row == null) continue;

                    for (int j = 0; j < row.LastCellNum; j++)
                    {
                        ICell cell = row.GetCell(j);
                        if (cell == null) continue;

                        string cellValue = cell.ToString();
                        if (cellValue.Equals("mac address"))
                        {
                            SetCellValue(row, j + 1, "Y");
                        }
                        if (cellValue.Equals("SSID/SVID check"))
                        {
                            SetCellValue(row, j + 1, "Y");
                        }

                    }
                }
            }
            void SetKeyboardValue(ISheet ItestPlan)
            {
                for (int i = 0; i <= ItestPlan.LastRowNum; i++)
                {
                    IRow row = ItestPlan.GetRow(i);
                    if (row == null) continue;

                    for (int j = 0; j <2; j++)
                    {
                        ICell cell = row.GetCell(j);
                        if (cell == null) continue;

                        string cellValue = cell.ToString();
                        if (cellValue.Equals("key test")) 
                        {
                            SetCellValue(row, j + 1, "Y");
                        }

                    }
                }
            }
            //void SetFixedColumnWidth(ISheet sheet, int width)
            //{
            //    // 获取最大列数
            //    int maxColIndex = 0;
            //    for (int rowIndex = 0; rowIndex <= sheet.LastRowNum; rowIndex++)
            //    {
            //        IRow row = sheet.GetRow(rowIndex);
            //        if (row != null && row.LastCellNum > maxColIndex)
            //        {
            //            maxColIndex = row.LastCellNum;
            //        }
            //    }

            //    // 设置每列的宽度为固定值
            //    for (int colIndex = 0; colIndex < maxColIndex; colIndex++)
            //    {
            //        sheet.SetColumnWidth(colIndex, width * 256); // 设置列宽，单位是1/256个字符宽度
            //    }
            //}

            List<Tuple<string, string>> FindSpecialCellsAndValues(ISheet sheet)
            {
                List<Tuple<string, string>> foundCellsAndValues = new List<Tuple<string, string>>();

                // 正则表达式匹配 Fn+ 开头的格式
                Regex regex = new Regex(@"^Fn\+[A-Z]");
                for (int rowIndex = 0; rowIndex <= sheet.LastRowNum; rowIndex++)
                {
                    IRow row = sheet.GetRow(rowIndex);
                    if (row == null) continue;

                    for (int colIndex = 0; colIndex < row.LastCellNum; colIndex++)
                    {
                        ICell cell = row.GetCell(colIndex);
                        if (cell != null)
                        {
                            string cellValue = cell.ToString();
                            if (regex.IsMatch(cellValue))
                            {
                                // 获取后面的一个单元格的值
                                ICell nextCell = row.GetCell(colIndex + 1);
                                string nextCellValue = nextCell != null ? nextCell.ToString() : string.Empty;
                                foundCellsAndValues.Add(new Tuple<string, string>(cellValue, nextCellValue));
                                Console.WriteLine($"Found matching cell: {cellValue} with value: {nextCellValue}");
                            }
                        }
                    }
                }

                return foundCellsAndValues;
            }

            //void AddValuesToTestPlan(ISheet testPlanSheet, List<Tuple<string, string>> values)
            //{
            //    int startRow = testPlanSheet.LastRowNum + 1; // 在最后一行之后添加

            //    foreach (var valuePair in values)
            //    {
            //        IRow row = testPlanSheet.CreateRow(startRow++);
            //        ICell cell1 = row.CreateCell(0);
            //        ICell cell2 = row.CreateCell(1);
            //        cell1.SetCellValue(valuePair.Item1);
            //        cell2.SetCellValue(valuePair.Item2);
            //    }
            //}

            // 设置每列的宽度为20字符宽度
            // SetFixedColumnWidth(testPlan, 20);

            Console.WriteLine(machinePath);
            if(flag == 1)
            {
                Stream stream = new FileStream(savaPath + "\\" + modelName + " " + "TestPlan.xlsx", FileMode.OpenOrCreate);
                WB.Write(stream);
                stream.Close();
                WB.Close();
                workbookwrite.Close();
            }
            if (flag == 0)
            {
                Stream stream = new FileStream(savaPath + "\\" + modelNameNext2 + " " + "TestPlan.xlsx", FileMode.OpenOrCreate);
                WB.Write(stream);
                stream.Close();
                WB.Close();
                workbookwrite.Close();
            }
        }
           
           
    }
}
