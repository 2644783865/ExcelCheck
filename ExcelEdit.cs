using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop;
using Microsoft.Office.Core;
using System.IO;

namespace ExcelCheck
{
    public partial class ExcelEdit : Form
    {
        public string mFilename;

        //定义Excel操作类
        public Microsoft.Office.Interop.Excel.Application app;
        public Microsoft.Office.Interop.Excel.Workbooks wbs;
        public Microsoft.Office.Interop.Excel.Workbook wb;
        public Microsoft.Office.Interop.Excel.Worksheets wss;
        public Microsoft.Office.Interop.Excel.Worksheet ws;

        //单位
        public string[] unitArray;
        //MRP控制
        public string[] mrpArray;
        //物料组
        public string[] itemGroupArray;
        //物料组1
        public string[] itemGroup1Array;
        //物料组2
        public string[] itemGroup2Array;
        //物料组3
        public string[] itemGroup3Array;
        //物料组4
        public string[] itemGroup4Array;
        //物料组5
        public string[] itemGroup5Array;
        //库存地点
        public string[] stockArray;
        //评估类
        public string[] evaluateArray;
        //采购组
        public string[] purchaseArray;
        //物料类型
        public string[] itemTypeArray;
        //工厂
        public string[] organizationArray;
        //科目设置组
        public string[] subjectArray;
        //项目类别组
        public string[] projectArray;
        //MRP类型
        public string[] mrpTypeArray;
        //批量大小
        public string[] batchArray;
        //采购类型
        public string[] purchaseTypeArray;
        //特殊采购类
        public string[] spurchaseTypeArray;
        //反冲
        public string[] backflushArray;

        public ExcelEdit()
        {
            InitializeComponent();
        }


        //配置文件初始化
        public void ConfigInit()
        {
            app = new Microsoft.Office.Interop.Excel.Application();
            wbs = app.Workbooks;
            wb = wbs.Add(System.Windows.Forms.Application.StartupPath + "\\主数据收集模板配置.xlsx");
            if (wb.Worksheets.Count != 20)
            {
                Close();
                MessageBox.Show("主数据收集模板配置文件内容缺失！");
            }
            else
            {
                //加载单位配置
                ws = (Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets["单位"];
                unitArray = new string[ws.UsedRange.Rows.Count - 1];
                for (int i = 2; i <= ws.UsedRange.Rows.Count; i++)
                {
                    unitArray[i - 2] = ws.get_Range(ws.Cells[i, 3], ws.Cells[i, 3]).Value.ToString();
                }
                //加载MRP控制配置
                ws = (Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets["MRP控制者"];
                mrpArray = new string[ws.UsedRange.Rows.Count - 1];
                for (int i = 2; i <= ws.UsedRange.Rows.Count; i++)
                {
                    mrpArray[i - 2] = ws.get_Range(ws.Cells[i, 2], ws.Cells[i, 2]).Value.ToString();
                }
                //加载物料组配置
                ws = (Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets["物料组"];
                itemGroupArray = new string[ws.UsedRange.Rows.Count - 1];
                for (int i = 2; i <= ws.UsedRange.Rows.Count; i++)
                {
                    itemGroupArray[i - 2] = ws.get_Range(ws.Cells[i, 2], ws.Cells[i, 2]).Value.ToString();
                }
                //加载物料组1配置
                ws = (Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets["物料组1"];
                itemGroup1Array = new string[ws.UsedRange.Rows.Count - 1];
                for (int i = 2; i <= ws.UsedRange.Rows.Count; i++)
                {
                    itemGroup1Array[i - 2] = ws.get_Range(ws.Cells[i, 1], ws.Cells[i, 1]).Value.ToString();
                }
                //加载物料组2配置
                ws = (Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets["物料组2"];
                itemGroup2Array = new string[ws.UsedRange.Rows.Count - 1];
                for (int i = 2; i <= ws.UsedRange.Rows.Count; i++)
                {
                    itemGroup2Array[i - 2] = ws.get_Range(ws.Cells[i, 1], ws.Cells[i, 1]).Value.ToString();
                }
                //加载物料3配置
                ws = (Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets["物料组3"];
                itemGroup3Array = new string[ws.UsedRange.Rows.Count - 1];
                for (int i = 2; i <= ws.UsedRange.Rows.Count; i++)
                {
                    itemGroup3Array[i - 2] = ws.get_Range(ws.Cells[i, 1], ws.Cells[i, 1]).Value.ToString();
                }
                //加载物料组4配置
                ws = (Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets["物料组4"];
                itemGroup4Array = new string[ws.UsedRange.Rows.Count - 1];
                for (int i = 2; i <= ws.UsedRange.Rows.Count; i++)
                {
                    itemGroup4Array[i - 2] = ws.get_Range(ws.Cells[i, 1], ws.Cells[i, 1]).Value.ToString();
                }
                //加载物料组5配置
                ws = (Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets["物料组5"];
                itemGroup5Array = new string[ws.UsedRange.Rows.Count - 1];
                for (int i = 2; i <= ws.UsedRange.Rows.Count; i++)
                {
                    itemGroup5Array[i - 2] = ws.get_Range(ws.Cells[i, 1], ws.Cells[i, 1]).Value.ToString();
                }
                //加载库存地点配置
                ws = (Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets["库存地点"];
                stockArray = new string[ws.UsedRange.Rows.Count - 1];
                for (int i = 2; i <= ws.UsedRange.Rows.Count; i++)
                {
                    stockArray[i - 2] = ws.get_Range(ws.Cells[i, 3], ws.Cells[i, 3]).Value.ToString();
                }
                //加载评估类配置
                ws = (Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets["评估类"];
                evaluateArray = new string[ws.UsedRange.Rows.Count - 1];
                for (int i = 2; i <= ws.UsedRange.Rows.Count; i++)
                {
                    evaluateArray[i - 2] = ws.get_Range(ws.Cells[i, 1], ws.Cells[i, 1]).Value.ToString();
                }
                //加载采购组配置
                ws = (Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets["采购组"];
                purchaseArray = new string[ws.UsedRange.Rows.Count - 1];
                for (int i = 2; i <= ws.UsedRange.Rows.Count; i++)
                {
                    purchaseArray[i - 2] = ws.get_Range(ws.Cells[i, 1], ws.Cells[i, 1]).Value.ToString();
                }
                //加载物料类型配置
                ws = (Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets["物料类型"];
                itemTypeArray = new string[ws.UsedRange.Rows.Count - 1];
                for (int i = 2; i <= ws.UsedRange.Rows.Count; i++)
                {
                    itemTypeArray[i - 2] = ws.get_Range(ws.Cells[i, 1], ws.Cells[i, 1]).Value.ToString();
                }
                //加载工厂配置
                ws = (Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets["工厂"];
                organizationArray = new string[ws.UsedRange.Rows.Count - 1];
                for (int i = 2; i <= ws.UsedRange.Rows.Count; i++)
                {
                    organizationArray[i - 2] = ws.get_Range(ws.Cells[i, 1], ws.Cells[i, 1]).Value.ToString();
                }
                //加载科目设置组配置
                ws = (Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets["科目设置组"];
                subjectArray = new string[ws.UsedRange.Rows.Count - 1];
                for (int i = 2; i <= ws.UsedRange.Rows.Count; i++)
                {
                    subjectArray[i - 2] = ws.get_Range(ws.Cells[i, 1], ws.Cells[i, 1]).Value.ToString();
                }
                //加载项目类别组配置
                ws = (Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets["项目类别组"];
                projectArray = new string[ws.UsedRange.Rows.Count - 1];
                for (int i = 2; i <= ws.UsedRange.Rows.Count; i++)
                {
                    projectArray[i - 2] = ws.get_Range(ws.Cells[i, 1], ws.Cells[i, 1]).Value.ToString();
                }
                //加载MRP类型配置
                ws = (Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets["MRP类型"];
                mrpTypeArray = new string[ws.UsedRange.Rows.Count - 1];
                for (int i = 2; i <= ws.UsedRange.Rows.Count; i++)
                {
                    mrpTypeArray[i - 2] = ws.get_Range(ws.Cells[i, 1], ws.Cells[i, 1]).Value.ToString();
                }
                //加载批量大小配置
                ws = (Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets["批量大小"];
                batchArray = new string[ws.UsedRange.Rows.Count - 1];
                for (int i = 2; i <= ws.UsedRange.Rows.Count; i++)
                {
                    batchArray[i - 2] = ws.get_Range(ws.Cells[i, 1], ws.Cells[i, 1]).Value.ToString();
                }
                //加载采购类型配置
                ws = (Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets["采购类型"];
                purchaseTypeArray = new string[ws.UsedRange.Rows.Count - 1];
                for (int i = 2; i <= ws.UsedRange.Rows.Count; i++)
                {
                    purchaseTypeArray[i - 2] = ws.get_Range(ws.Cells[i, 1], ws.Cells[i, 1]).Value.ToString();
                }
                //加载特殊采购类配置
                ws = (Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets["特殊采购类"];
                spurchaseTypeArray = new string[ws.UsedRange.Rows.Count - 1];
                for (int i = 2; i <= ws.UsedRange.Rows.Count; i++)
                {
                    spurchaseTypeArray[i - 2] = ws.get_Range(ws.Cells[i, 1], ws.Cells[i, 1]).Value.ToString();
                }
                //加载反冲配置
                ws = (Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets["反冲"];
                backflushArray = new string[ws.UsedRange.Rows.Count - 1];
                for (int i = 2; i <= ws.UsedRange.Rows.Count; i++)
                {
                    backflushArray[i - 2] = ws.get_Range(ws.Cells[i, 1], ws.Cells[i, 1]).Value.ToString();
                }
                Close();
            }
        }

        //excel校验
        public void Check()
        {
            ConfigInit();

            Open(FileNameTxt.Text);

            ws = (Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets[1];

            if (ws.UsedRange.Columns.Count != 57 || ws.UsedRange.Rows.Count < 5)
            {
                Close();
                ConfigRelease();
                MessageBox.Show("请选择正确的主数据收集模板！");
            }
            else
            {
                //开始校验
                for (int i = 6; i <= ws.UsedRange.Rows.Count; i++)
                {
                    try
                    {
                        #region 校验配置文件
                        //校验工厂
                        if (!organizationArray.Contains(GetValue(i, 1)) && GetValue(i, 1) != "")
                        {
                            ws.get_Range(ws.Cells[i, 1], ws.Cells[i, 1]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.FromArgb(255, 0, 0));
                        }
                        //校验单位
                        if (!unitArray.Contains(GetValue(i, 4)) && GetValue(i, 4) != "")
                        {
                            ws.get_Range(ws.Cells[i, 4], ws.Cells[i, 4]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.FromArgb(255, 0, 0));
                        }
                        //校验物料类型
                        if (!itemTypeArray.Contains(GetValue(i, 5)) && GetValue(i, 5) != "")
                        {
                            ws.get_Range(ws.Cells[i, 5], ws.Cells[i, 5]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.FromArgb(255, 0, 0));
                        }
                        //校验物料组
                        if (!itemGroupArray.Contains(GetValue(i, 6)) && GetValue(i, 6) != "")
                        {
                            ws.get_Range(ws.Cells[i, 6], ws.Cells[i, 6]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.FromArgb(255, 0, 0));
                        }
                        //校验工厂
                        if (!organizationArray.Contains(GetValue(i, 11)) && GetValue(i, 11) != "")
                        {
                            ws.get_Range(ws.Cells[i, 11], ws.Cells[i, 11]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.FromArgb(255, 0, 0));
                        }
                        //校验交货工厂
                        if (!organizationArray.Contains(GetValue(i, 16)) && GetValue(i, 16) != "")
                        {
                            ws.get_Range(ws.Cells[i, 16], ws.Cells[i, 16]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.FromArgb(255, 0, 0));
                        }
                        //校验科目设置组
                        if (!subjectArray.Contains(GetValue(i, 18)) && GetValue(i, 18) != "")
                        {
                            ws.get_Range(ws.Cells[i, 18], ws.Cells[i, 18]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.FromArgb(255, 0, 0));
                        }
                        //校验项目类别组
                        if (!projectArray.Contains(GetValue(i, 19)) && GetValue(i, 19) != "")
                        {
                            ws.get_Range(ws.Cells[i, 19], ws.Cells[i, 19]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.FromArgb(255, 0, 0));
                        }
                        //校验物料组1
                        if (!itemGroup1Array.Contains(GetValue(i, 20)) && GetValue(i, 20) != "")
                        {
                            ws.get_Range(ws.Cells[i, 20], ws.Cells[i, 20]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.FromArgb(255, 0, 0));
                        }
                        //校验物料组2
                        if (!itemGroup2Array.Contains(GetValue(i, 21)) && GetValue(i, 21) != "")
                        {
                            ws.get_Range(ws.Cells[i, 21], ws.Cells[i, 21]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.FromArgb(255, 0, 0));
                        }
                        //校验物料组3
                        if (!itemGroup3Array.Contains(GetValue(i, 22)) && GetValue(i, 22) != "")
                        {
                            ws.get_Range(ws.Cells[i, 22], ws.Cells[i, 22]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.FromArgb(255, 0, 0));
                        }
                        //校验物料组4
                        if (!itemGroup4Array.Contains(GetValue(i, 23)) && GetValue(i, 23) != "")
                        {
                            ws.get_Range(ws.Cells[i, 23], ws.Cells[i, 23]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.FromArgb(255, 0, 0));
                        }
                        //校验物料组5
                        if (!itemGroup5Array.Contains(GetValue(i, 24)) && GetValue(i, 24) != "")
                        {
                            ws.get_Range(ws.Cells[i, 24], ws.Cells[i, 24]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.FromArgb(255, 0, 0));
                        }
                        //校验装载组
                        if (!organizationArray.Contains(GetValue(i, 25)) && GetValue(i, 25) != "")
                        {
                            ws.get_Range(ws.Cells[i, 25], ws.Cells[i, 25]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.FromArgb(255, 0, 0));
                        }
                        //校验采购组
                        if (!purchaseArray.Contains(GetValue(i, 26)) && GetValue(i, 26) != "")
                        {
                            ws.get_Range(ws.Cells[i, 26], ws.Cells[i, 26]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.FromArgb(255, 0, 0));
                        }
                        //校验MRP类型
                        if (!mrpTypeArray.Contains(GetValue(i, 30)) && GetValue(i, 30) != "")
                        {
                            ws.get_Range(ws.Cells[i, 30], ws.Cells[i, 30]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.FromArgb(255, 0, 0));
                        }
                        //校验MRP控制者
                        if (!mrpArray.Contains(GetValue(i, 31)) && GetValue(i, 31) != "")
                        {
                            ws.get_Range(ws.Cells[i, 31], ws.Cells[i, 31]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.FromArgb(255, 0, 0));
                        }
                        //校验批量大小
                        if (!batchArray.Contains(GetValue(i, 32)) && GetValue(i, 32) != "")
                        {
                            ws.get_Range(ws.Cells[i, 32], ws.Cells[i, 32]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.FromArgb(255, 0, 0));
                        }
                        //校验采购类型
                        if (!purchaseTypeArray.Contains(GetValue(i, 37)) && GetValue(i, 37) != "")
                        {
                            ws.get_Range(ws.Cells[i, 37], ws.Cells[i, 37]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.FromArgb(255, 0, 0));
                        }
                        //校验特殊采购类
                        if (!spurchaseTypeArray.Contains(GetValue(i, 38)) && GetValue(i, 38) != "")
                        {
                            ws.get_Range(ws.Cells[i, 38], ws.Cells[i, 38]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.FromArgb(255, 0, 0));
                        }
                        //校验反冲
                        if (!backflushArray.Contains(GetValue(i, 39)) && GetValue(i, 39) != "")
                        {
                            ws.get_Range(ws.Cells[i, 39], ws.Cells[i, 39]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.FromArgb(255, 0, 0));
                        }
                        //校验生产存储地点
                        if (!stockArray.Contains(GetValue(i, 40)) && GetValue(i, 40) != "")
                        {
                            ws.get_Range(ws.Cells[i, 40], ws.Cells[i, 40]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.FromArgb(255, 0, 0));
                        }
                        //校验采购储存地点
                        if (!stockArray.Contains(GetValue(i, 41)) && GetValue(i, 41) != "")
                        {
                            ws.get_Range(ws.Cells[i, 41], ws.Cells[i, 41]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.FromArgb(255, 0, 0));
                        }
                        //校验评估类
                        if (!evaluateArray.Contains(GetValue(i, 52)) && GetValue(i, 52) != "")
                        {
                            ws.get_Range(ws.Cells[i, 52], ws.Cells[i, 52]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.FromArgb(255, 0, 0));
                        }
                        #endregion

                        #region 校验逻辑
                        //描述不能大于40
                        if (GetValue(i, 3).Length > 80)
                        {
                            ws.get_Range(ws.Cells[i, 3], ws.Cells[i, 3]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.FromArgb(255, 255, 0));
                        }
                        //单位必填
                        if (GetValue(i, 4) == "")
                        {
                            ws.get_Range(ws.Cells[i, 4], ws.Cells[i, 4]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.FromArgb(255, 255, 0));
                        }
                        //物料类型必填
                        if (GetValue(i, 5) == "")
                        {
                            ws.get_Range(ws.Cells[i, 5], ws.Cells[i, 5]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.FromArgb(255, 255, 0));
                        }
                        //物料组必填
                        if (GetValue(i, 6) == "" && GetValue(i, 5) != "ZFIG")
                        {
                            ws.get_Range(ws.Cells[i, 6], ws.Cells[i, 6]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.FromArgb(255, 255, 0));
                        }

                        //采购类型、评估类、mrp类型校验
                        if (GetValue(i, 5) == "ZROH")
                        {
                            if (GetValue(i, 37) != "F")
                            {
                                ws.get_Range(ws.Cells[i, 37], ws.Cells[i, 37]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.FromArgb(255, 255, 0));
                            }
                            if (GetValue(i, 52) != "Z300")
                            {
                                ws.get_Range(ws.Cells[i, 52], ws.Cells[i, 52]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.FromArgb(255, 255, 0));
                            }
                            if (GetValue(i, 30) != "VB")
                            {
                                ws.get_Range(ws.Cells[i, 30], ws.Cells[i, 30]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.FromArgb(255, 255, 0));
                            }
                        }
                        else if (GetValue(i, 5) == "ZHAL")
                        {
                            if (GetValue(i, 37) != "E")
                            {
                                ws.get_Range(ws.Cells[i, 37], ws.Cells[i, 37]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.FromArgb(255, 255, 0));
                            }
                            if (GetValue(i, 52) != "Z790")
                            {
                                ws.get_Range(ws.Cells[i, 52], ws.Cells[i, 52]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.FromArgb(255, 255, 0));
                            }
                            if (GetValue(i, 30) != "PD")
                            {
                                ws.get_Range(ws.Cells[i, 30], ws.Cells[i, 30]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.FromArgb(255, 255, 0));
                            }
                        }
                        else if (GetValue(i, 5) == "ZFER")
                        {
                            if (GetValue(i, 37) != "E")
                            {
                                ws.get_Range(ws.Cells[i, 37], ws.Cells[i, 37]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.FromArgb(255, 255, 0));
                            }
                            if (GetValue(i, 52) != "Z792")
                            {
                                ws.get_Range(ws.Cells[i, 52], ws.Cells[i, 52]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.FromArgb(255, 255, 0));
                            }
                            if (GetValue(i, 30) != "PD")
                            {
                                ws.get_Range(ws.Cells[i, 30], ws.Cells[i, 30]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.FromArgb(255, 255, 0));
                            }
                        }
                        else if (GetValue(i, 5) == "ZACC")
                        {
                            if (GetValue(i, 37) != "F")
                            {
                                ws.get_Range(ws.Cells[i, 37], ws.Cells[i, 37]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.FromArgb(255, 255, 0));
                            }
                            if (GetValue(i, 52) != "Z200")
                            {
                                ws.get_Range(ws.Cells[i, 52], ws.Cells[i, 52]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.FromArgb(255, 255, 0));
                            }
                            if (GetValue(i, 30) != "PD")
                            {
                                ws.get_Range(ws.Cells[i, 30], ws.Cells[i, 30]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.FromArgb(255, 255, 0));
                            }
                        }
                        else if (GetValue(i, 5) == "ZFIG")
                        {
                            if (GetValue(i, 37) != "F" && GetValue(i, 37) != "")
                            {
                                ws.get_Range(ws.Cells[i, 37], ws.Cells[i, 37]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.FromArgb(255, 255, 0));
                            }
                            if (GetValue(i, 52) != "")
                            {
                                ws.get_Range(ws.Cells[i, 52], ws.Cells[i, 52]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.FromArgb(255, 255, 0));
                            }
                        }
                        else if (GetValue(i, 5) == "ZEQU")
                        {
                            if (GetValue(i, 37) != "")
                            {
                                ws.get_Range(ws.Cells[i, 37], ws.Cells[i, 37]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.FromArgb(255, 255, 0));
                            }
                            if (GetValue(i, 52) != "")
                            {
                                ws.get_Range(ws.Cells[i, 52], ws.Cells[i, 52]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.FromArgb(255, 255, 0));
                            }
                        }

                        //校验plm物料号
                        if (!isnumeric(GetValue(i, 2)) || GetValue(i, 2).Length != 8)
                        {

                            ws.get_Range(ws.Cells[i, 2], ws.Cells[i, 2]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.FromArgb(255, 255, 0));
                        }

                        //判断销售组织、交货工厂、装配组
                        if (GetValue(i, 1) == "1000")
                        {
                            if (GetValue(i, 11) == "1000")
                            {
                                if (GetValue(i, 16) != "1000")
                                {
                                    ws.get_Range(ws.Cells[i, 16], ws.Cells[i, 16]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.FromArgb(255, 255, 0));
                                }
                                if (GetValue(i, 25) != "1000")
                                {
                                    ws.get_Range(ws.Cells[i, 25], ws.Cells[i, 25]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.FromArgb(255, 255, 0));
                                }
                            }
                            else
                            {
                                if (GetValue(i, 11) != "")
                                {
                                    ws.get_Range(ws.Cells[i, 11], ws.Cells[i, 11]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.FromArgb(255, 255, 0));
                                }
                            }
                        }
                        else if (GetValue(i, 1) == "2000")
                        {
                            if (GetValue(i, 11) == "2000")
                            {
                                if (GetValue(i, 16) != "2000")
                                {
                                    ws.get_Range(ws.Cells[i, 16], ws.Cells[i, 16]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.FromArgb(255, 255, 0));
                                }
                                if (GetValue(i, 25) != "2000")
                                {
                                    ws.get_Range(ws.Cells[i, 25], ws.Cells[i, 25]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.FromArgb(255, 255, 0));
                                }
                            }
                            else
                            {
                                if (GetValue(i, 11) != "")
                                {
                                    ws.get_Range(ws.Cells[i, 11], ws.Cells[i, 11]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.FromArgb(255, 255, 0));
                                }
                            }
                        }

                        //分销渠道校验
                        if (GetValue(i, 11) != "" && GetValue(i, 12) == "")
                        {
                            ws.get_Range(ws.Cells[i, 12], ws.Cells[i, 12]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.FromArgb(255, 255, 0));
                        }
                        //产品组校验
                        if (GetValue(i, 11) != "" && GetValue(i, 13) == "")
                        {
                            ws.get_Range(ws.Cells[i, 13], ws.Cells[i, 13]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.FromArgb(255, 255, 0));
                        }
                        //税分类校验
                        if (GetValue(i, 11) != "" && GetValue(i, 17) == "")
                        {
                            ws.get_Range(ws.Cells[i, 17], ws.Cells[i, 17]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.FromArgb(255, 255, 0));
                        }
                        //科目设置组校验
                        if (GetValue(i, 11) != "" && GetValue(i, 18) == "")
                        {
                            ws.get_Range(ws.Cells[i, 18], ws.Cells[i, 18]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.FromArgb(255, 255, 0));
                        }
                        //项目类别组校验
                        if (GetValue(i, 11) != "" && GetValue(i, 19) == "")
                        {
                            ws.get_Range(ws.Cells[i, 19], ws.Cells[i, 19]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.FromArgb(255, 255, 0));
                        }

                        //采购组校验
                        //if (GetValue(i, 11) == "" && GetValue(i, 26) == "")
                        //{
                        //    ws.get_Range(ws.Cells[i, 26], ws.Cells[i, 26]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.FromArgb(255, 255, 0));
                        //}
                        //源清单校验
                        if (GetValue(i, 26) != "" && GetValue(i, 29) == "")
                        {
                            ws.get_Range(ws.Cells[i, 29], ws.Cells[i, 29]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.FromArgb(255, 255, 0));
                        }
                        //mrp类型校验
                        if (GetValue(i, 5) != "ZFIG")
                        {
                            if (GetValue(i, 30) == "")
                            {
                                ws.get_Range(ws.Cells[i, 30], ws.Cells[i, 30]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.FromArgb(255, 255, 0));
                            }
                            else
                            {
                                if (GetValue(i, 30) == "PD" && GetValue(i, 5) != "ZHAL" && GetValue(i, 5) != "ZFER" && GetValue(i, 5) != "ZACC")
                                {
                                    ws.get_Range(ws.Cells[i, 30], ws.Cells[i, 30]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.FromArgb(255, 255, 0));
                                }
                                else if (GetValue(i, 30) == "VB" && GetValue(i, 5) != "ZROH")
                                {
                                    ws.get_Range(ws.Cells[i, 30], ws.Cells[i, 30]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.FromArgb(255, 255, 0));
                                }
                            }
                        }
                        //mrp控制者校验
                        if (GetValue(i, 31) == "" && GetValue(i, 5) != "ZFIG")
                        {
                            ws.get_Range(ws.Cells[i, 31], ws.Cells[i, 31]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.FromArgb(255, 255, 0));
                        }
                        //批量大小校验
                        if (GetValue(i, 32) == "" && GetValue(i, 5) != "ZFIG")
                        {
                            ws.get_Range(ws.Cells[i, 32], ws.Cells[i, 32]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.FromArgb(255, 255, 0));
                        }
                        //再订货点校验
                        if ((GetValue(i, 5) == "ZROH" || GetValue(i, 30) == "VB") && GetValue(i, 33) == "")
                        {
                            ws.get_Range(ws.Cells[i, 33], ws.Cells[i, 33]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.FromArgb(255, 255, 0));
                        }
                        //固定批量大小校验
                        if (GetValue(i, 5) == "ZROH" && GetValue(i, 34) == "")
                        {
                            ws.get_Range(ws.Cells[i, 34], ws.Cells[i, 34]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.FromArgb(255, 255, 0));
                        }
                        //采购类型校验
                        if (GetValue(i, 37) == "" && GetValue(i, 5) != "ZFIG")
                        {
                            ws.get_Range(ws.Cells[i, 37], ws.Cells[i, 37]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.FromArgb(255, 255, 0));
                        }
                        //计划边际码校验
                        if (GetValue(i, 44) == "" && GetValue(i, 5) != "ZFIG")
                        {
                            ws.get_Range(ws.Cells[i, 44], ws.Cells[i, 44]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.FromArgb(255, 255, 0));
                        }
                        //可用性检查必填
                        if (GetValue(i, 47) == "" && GetValue(i, 5) != "ZFIG")
                        {
                            ws.get_Range(ws.Cells[i, 47], ws.Cells[i, 47]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.FromArgb(255, 255, 0));
                        }
                        //独立集中必填
                        if (GetValue(i, 48) == "" && GetValue(i, 5) != "ZFIG")
                        {
                            ws.get_Range(ws.Cells[i, 48], ws.Cells[i, 48]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.FromArgb(255, 255, 0));
                        }
                        //评估类校验
                        if (GetValue(i, 52) == "" && GetValue(i, 5) != "ZFIG")
                        {
                            ws.get_Range(ws.Cells[i, 52], ws.Cells[i, 52]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.FromArgb(255, 255, 0));
                        }
                        //价格控制校验
                        if (GetValue(i, 53) == "" && GetValue(i, 5) != "ZFIG")
                        {
                            ws.get_Range(ws.Cells[i, 53], ws.Cells[i, 53]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.FromArgb(255, 255, 0));
                        }
                        //价格单位校验
                        if (GetValue(i, 54) == "" && GetValue(i, 5) != "ZFIG")
                        {
                            ws.get_Range(ws.Cells[i, 54], ws.Cells[i, 54]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.FromArgb(255, 255, 0));
                        }
                        //物料来源校验
                        if (GetValue(i, 55) == "" && GetValue(i, 5) != "ZFIG")
                        {
                            ws.get_Range(ws.Cells[i, 55], ws.Cells[i, 55]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.FromArgb(255, 255, 0));
                        }
                        //用QS的成本核算校验
                        if (GetValue(i, 56) == "" && GetValue(i, 5) != "ZFIG")
                        {
                            ws.get_Range(ws.Cells[i, 56], ws.Cells[i, 56]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.FromArgb(255, 255, 0));
                        }
                        //成本核算批量校验
                        if (GetValue(i, 57) == "" && GetValue(i, 5) != "ZFIG")
                        {
                            ws.get_Range(ws.Cells[i, 57], ws.Cells[i, 57]).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.FromArgb(255, 255, 0));
                        }

                        #endregion
                    }
                    catch (Exception se)
                    {
                        MessageBox.Show(se.ToString());
                    }
                }

                SaveAs(mFilename.Substring(0, mFilename.LastIndexOf(".")) + "_" + DateTime.Now.Year + "-" + DateTime.Now.Month + "-" + DateTime.Now.Day + ".xlsx");

                Close();
                ConfigRelease();
                MessageBox.Show("校验完成！");
            }
        }


        //打开一个Microsoft.Office.Interop.Excel文件
        public void Open(string FileName)
        {
            app = new Microsoft.Office.Interop.Excel.Application();
            wbs = app.Workbooks;
            wb = wbs.Add(FileName);
            mFilename = FileName;
        }


        //文档另存为
        public bool SaveAs(object FileName)
        {

            app.AlertBeforeOverwriting = false;

            try
            {
                wb.SaveAs(FileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                return true;

            }
            catch (Exception ex)
            {
                return false;

            }
        }

        //关闭一个Microsoft.Office.Interop.Excel对象，销毁对象
        public void Close()
        {
            wb.Close(Type.Missing, Type.Missing, Type.Missing);
            wbs.Close();
            app.Quit();
            wb = null;
            wbs = null;
            app = null;
            GC.Collect();
        }

        //清除配置数据
        public void ConfigRelease()
        {
            unitArray = null;
            mrpArray = null;
            itemGroupArray = null;
            itemGroup1Array = null;
            itemGroup2Array = null;
            itemGroup3Array = null;
            itemGroup4Array = null;
            itemGroup5Array = null;
            stockArray = null;
            evaluateArray = null;
            purchaseArray = null;
            itemTypeArray = null;
            organizationArray = null;
            subjectArray = null;
            projectArray = null;
            mrpTypeArray = null;
            batchArray = null;
            purchaseTypeArray = null;
            spurchaseTypeArray = null;
            backflushArray = null;
        }

        //打开选择文件窗口
        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {
            FileNameTxt.Text = openFileDialog1.FileName;
        }

        //浏览按钮
        private void OpenBtn_Click(object sender, EventArgs e)
        {
            openFileDialog1.ShowDialog();
        }

        //校验按钮
        private void CheckBtn_Click(object sender, EventArgs e)
        {
            if (FileNameTxt.Text == "")
            {
                MessageBox.Show("请选择需要校验的模板文件！");
            }
            else
            {
                if (!File.Exists(System.Windows.Forms.Application.StartupPath + "\\主数据收集模板配置.xlsx"))
                {
                    MessageBox.Show("配置文件不存在！");
                }
                else
                {
                    if (!File.Exists(FileNameTxt.Text))
                    {
                        MessageBox.Show("需要校验的文件不存在！");
                    }
                    else
                    {
                        pictureBox1.Visible = true;
                        backgroundWorker1.RunWorkerAsync();
                    }
                }
            }
        }

        //判断是否为数字
        public bool isnumeric(string str)
        {
            char[] ch = new char[str.Length];
            ch = str.ToCharArray();
            for (int i = 0; i < str.Length; i++)
            {
                if (ch[i] < 48 || ch[i] > 57)
                {
                    return false;
                }
            }
            return true;
        }

        //下载模板文件
        private void DownLoadBtn_Click(object sender, EventArgs e)
        {
            try
            {
                saveFileDialog1.Filter = "xlsx|*.xlsx";
                saveFileDialog1.ShowDialog();
                if (File.Exists(@saveFileDialog1.FileName))
                {
                    File.Delete(@saveFileDialog1.FileName);
                }
                FileStream fs = new FileStream(saveFileDialog1.FileName, FileMode.CreateNew, FileAccess.Write);
                byte[] buffer = ExcelCheck.Properties.Resources.主数据收集模板;
                fs.Write(buffer, 0, buffer.Length);
                fs.Close();

            }
            catch (Exception se)
            {

            }
        }
        //获取单元格值
        private string GetValue(int x, int y)
        {
            string _Return = "";
            try
            {
                _Return = ws.get_Range(ws.Cells[x, y], ws.Cells[x, y]).Value.ToString();
            }
            catch
            { }
            return _Return;
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            Check();
        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            pictureBox1.Visible = false;
        }

    }
}
