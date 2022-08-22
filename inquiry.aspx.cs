using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Text;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using Newtonsoft.Json;
using System.IO;
using System.Net;
using System.Drawing;

namespace meteorCRMExport
{
    public partial class inquiry : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            var request = HttpContext.Current.Request;

            if (request.InputStream.Length == 0)
            {
                Response.ContentType = "text/html";
                Response.Write("数据错误!");
                Response.End();
            }

            byte[] requestData = new byte[request.InputStream.Length];

            request.InputStream.Read(requestData, 0, (int)request.InputStream.Length);

            var jsonData = Encoding.UTF8.GetString(requestData);

            dynamic m = JsonConvert.DeserializeObject<dynamic>(jsonData);

            var orderModel = m.orderModel;
            var productModel = m.productModel;

            OutputExcel(orderModel, productModel);
        }

        [DllImport("user32.dll", SetLastError = true)]
        static extern int GetWindowThreadProcessId(IntPtr hWnd, out int lpdwProcessId);

        public void OutputExcel(dynamic orderModel, dynamic productModel)
        {
            GC.Collect();
            Application excel = new Application();
            _Workbook xBk = excel.Workbooks.Add(true);
            _Worksheet xSt = (_Worksheet)xBk.ActiveSheet;

            excel.DisplayAlerts = false;

            excel.Cells[1, 1] = "_id";
            excel.Cells[1, 2] = "询价型号";
            excel.Cells[1, 3] = "询价名称";
            excel.Cells[1, 4] = "询价品牌";
            excel.Cells[1, 5] = "询价数量";
            excel.Cells[1, 6] = "询价单位";

            excel.Cells[1, 7] = "产品名称";
            excel.Cells[1, 8] = "分类名称";
            excel.Cells[1, 9] = "品牌（中文）";
            excel.Cells[1, 10] = "制造商型号";
            excel.Cells[1, 11] = "单位";


            excel.Cells[1, 12] = "报价建议";
            excel.Cells[1, 13] = "推荐供货商【1】";
            excel.Cells[1, 14] = "联系信息【1】";
            excel.Cells[1, 15] = "含税（y|n）【1】";
            excel.Cells[1, 16] = "单价【1】";
            excel.Cells[1, 17] = "单件运费【1】";
            excel.Cells[1, 18] = "最小起订量【1】";
            excel.Cells[1, 19] = "交货期【1】";
            excel.Cells[1, 20] = "备注（推荐原因）【1】";
            excel.Cells[1, 21] = "供货商【2】";
            excel.Cells[1, 22] = "联系信息【2】";
            excel.Cells[1, 23] = "含税（y|n）【2】";
            excel.Cells[1, 24] = "单价【2】";
            excel.Cells[1, 25] = "单件运费【2】";
            excel.Cells[1, 26] = "最小起订量【2】";
            excel.Cells[1, 27] = "交货期【2】";
            excel.Cells[1, 28] = "备注（推荐原因）【2】";
            excel.Cells[1, 29] = "供货商【3】";
            excel.Cells[1, 30] = "联系信息【3】";
            excel.Cells[1, 31] = "含税（y|n）【3】";
            excel.Cells[1, 32] = "单价【3】";
            excel.Cells[1, 33] = "单件运费【3】";
            excel.Cells[1, 34] = "最小起订量【3】";
            excel.Cells[1, 35] = "交货期【3】";
            excel.Cells[1, 36] = "备注（推荐原因）【3】";

            xSt.Range[excel.Cells[1, 1], excel.Cells[1, 36]].HorizontalAlignment = XlHAlign.xlHAlignCenter;//设置标题格式为居中对齐 
            xSt.Range[excel.Cells[1, 1], excel.Cells[1, 36]].Font.Bold = true;

            for (int i = 0; i < Convert.ToInt32(orderModel["number"]); i++)
            {
                excel.Cells[i + 2, 1] = productModel[i]["_id"].ToString().Trim();
                excel.Cells[i + 2, 2] = productModel[i]["productModel"].ToString().Trim();
                excel.Cells[i + 2, 3] = productModel[i]["productName"].ToString().Trim();
                excel.Cells[i + 2, 4] = productModel[i]["brandName"].ToString().Trim();
                excel.Cells[i + 2, 5] = productModel[i]["quantity"].ToString().Trim();
                excel.Cells[i + 2, 6] = productModel[i]["unit"].ToString().Trim();

                excel.Cells[i + 2, 15] = "y";
                excel.Cells[i + 2, 23] = "y";
                excel.Cells[i + 2, 31] = "y";

                xSt.Range[excel.Cells[i + 2, 1], excel.Cells[i + 2, 6]].HorizontalAlignment = XlHAlign.xlHAlignLeft;//设置标题格式为居中对齐 
                xSt.Range[excel.Cells[i + 2, 5], excel.Cells[i + 2, 5]].HorizontalAlignment = XlHAlign.xlHAlignCenter;//设置标题格式为居中对齐 
                xSt.Range[excel.Cells[i + 2, 6], excel.Cells[i + 2, 6]].HorizontalAlignment = XlHAlign.xlHAlignCenter;//设置标题格式为居中对齐 
            }

            // 
            //加载一个合计行 
            // 
            int rowSum = Convert.ToInt32(orderModel["number"].ToString()) + 1;
            //int colSum = 2;

            //设置背景色
            Color c1 = Color.FromArgb(216, 216, 216);
            Color c2 = Color.FromArgb(255, 255, 0);
            Color c3 = Color.FromArgb(252, 213, 180);
            Color c4 = Color.FromArgb(141, 180, 227);
            Color c5 = Color.FromArgb(194, 214, 154);
            xSt.Range[excel.Cells[1, 1], excel.Cells[rowSum, 6]].Interior.Color = System.Drawing.ColorTranslator.ToOle(c1);
            xSt.Range[excel.Cells[1, 12], excel.Cells[rowSum, 12]].Interior.Color = System.Drawing.ColorTranslator.ToOle(c2);
            xSt.Range[excel.Cells[1, 13], excel.Cells[rowSum, 20]].Interior.Color = System.Drawing.ColorTranslator.ToOle(c3);
            xSt.Range[excel.Cells[1, 21], excel.Cells[rowSum, 28]].Interior.Color = System.Drawing.ColorTranslator.ToOle(c4);
            xSt.Range[excel.Cells[1, 29], excel.Cells[rowSum, 36]].Interior.Color = System.Drawing.ColorTranslator.ToOle(c5);
            // 
            //绘制边框 
            // 
            xSt.Range[excel.Cells[1, 1], excel.Cells[rowSum, 36]].Borders.LineStyle = 1;
            xSt.Range[excel.Cells[1, 1], excel.Cells[rowSum, 1]].Borders[XlBordersIndex.xlEdgeLeft].Weight = XlBorderWeight.xlThin;//设置左边线加粗 
            xSt.Range[excel.Cells[1, 1], excel.Cells[1, 36]].Borders[XlBordersIndex.xlEdgeTop].Weight = XlBorderWeight.xlThin;//设置上边线加粗 
            xSt.Range[excel.Cells[1, 36], excel.Cells[rowSum, 36]].Borders[XlBordersIndex.xlEdgeRight].Weight = XlBorderWeight.xlThin;//设置右边线加粗 
            xSt.Range[excel.Cells[rowSum, 1], excel.Cells[rowSum, 36]].Borders[XlBordersIndex.xlEdgeBottom].Weight = XlBorderWeight.xlThin;//设置下边线加粗 
                                                                                                                                                     //
                                                                                                                                                     //设置高度为20
                                                                                                                                                     //
            xSt.Range[excel.Cells[1, 1], excel.Cells[rowSum, 36]].RowHeight = 23;
            xSt.Range[excel.Cells[1, 1], excel.Cells[rowSum, 36]].Font.Name = "微软雅黑";
            xSt.Range[excel.Cells[1, 1], excel.Cells[rowSum, 36]].Font.Size = 9;
            // 
            //设置报表表格为最适应宽度 
            // 
            xSt.Range[excel.Cells[1, 1], excel.Cells[rowSum, 36]].Select();
            xSt.Range[excel.Cells[1, 1], excel.Cells[rowSum, 36]].Columns.AutoFit();
            //
            //冻结窗格
            //
            xSt.Application.ActiveWindow.SplitRow = 1;
            xSt.Application.ActiveWindow.SplitColumn = 6;
            xSt.Application.ActiveWindow.FreezePanes = true;
            //
            //设置最后一列的列宽为30
            //
            //xSt.get_Range(excel.Cells[1, colIndex], excel.Cells[rowSum, colIndex]).ColumnWidth = 30;
            // 
            //显示效果 
            // 
            excel.Visible = true;

            string filename = orderModel["_id"] + ".xlsx";

            //xSt.Export(Server.MapPath(".")+"\\"+this.xlfile.Text+".xls",SheetExportActionEnum.ssExportActionNone,Microsoft.Office.Interop.OWC.SheetExportFormat.ssExportHTML ); 
            xBk.SaveCopyAs(Server.MapPath("~/") + "temp\\" + filename);

            //ds = null;
            xBk.Close(false, null, null);

            excel.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xSt);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xBk);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);

            try
            {
                if (excel != null)
                {
                    int lpdwProcessId;
                    GetWindowThreadProcessId(new IntPtr(excel.Hwnd), out lpdwProcessId);
                    System.Diagnostics.Process.GetProcessById(lpdwProcessId).Kill();
                }
            }
            catch (Exception ex)
            {
            }

            xBk = null;
            excel = null;
            xSt = null;
            GC.Collect();


            Response.Clear();
            Response.ContentType = "text/html";
            Response.Write(filename);
            Response.End();

        }
    }
}