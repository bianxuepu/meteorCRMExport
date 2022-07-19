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
    public partial class externalInquiry : System.Web.UI.Page
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

            xSt.PageSetup.Orientation = XlPageOrientation.xlLandscape; //页面横版

            xSt.PageSetup.LeftMargin = 0.9 / 0.035;
            xSt.PageSetup.RightMargin = 0.9 / 0.035;
            xSt.PageSetup.HeaderMargin = 2 / 0.035;
            xSt.PageSetup.FooterMargin = 1 / 0.035;
            xSt.PageSetup.TopMargin = 1 / 0.035;
            xSt.PageSetup.BottomMargin = 1 / 0.035;


            //设置列宽
            xSt.Range[excel.Cells[1, 1], excel.Cells[1, 1]].ColumnWidth = 4;
            xSt.Range[excel.Cells[1, 2], excel.Cells[1, 2]].ColumnWidth = 15;
            xSt.Range[excel.Cells[1, 3], excel.Cells[1, 3]].ColumnWidth = 24;
            xSt.Range[excel.Cells[1, 4], excel.Cells[1, 4]].ColumnWidth = 4;
            xSt.Range[excel.Cells[1, 5], excel.Cells[1, 5]].ColumnWidth = 4;
            xSt.Range[excel.Cells[1, 6], excel.Cells[1, 6]].ColumnWidth = 12;
            xSt.Range[excel.Cells[1, 7], excel.Cells[1, 7]].ColumnWidth = 12;
            xSt.Range[excel.Cells[1, 8], excel.Cells[1, 8]].ColumnWidth = 12;
            xSt.Range[excel.Cells[1, 9], excel.Cells[1, 9]].ColumnWidth = 8;
            xSt.Range[excel.Cells[1, 10], excel.Cells[1, 10]].ColumnWidth = 12;
            xSt.Range[excel.Cells[1, 11], excel.Cells[1, 11]].ColumnWidth = 12;
            xSt.Range[excel.Cells[1, 12], excel.Cells[1, 12]].ColumnWidth = 9;
            xSt.Range[excel.Cells[1, 13], excel.Cells[1, 13]].ColumnWidth = 12;

            xSt.Range[excel.Cells[1, 1], excel.Cells[1, 13]].RowHeight = 25;
            xSt.Shapes.AddPicture(Server.MapPath("~/").ToString().Trim() + "image\\00001.gif", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoTrue, 0, 4, 150, 13);

            xSt.Range[excel.Cells[2, 1], excel.Cells[2, 2]].Merge(false);
            xSt.Range[excel.Cells[2, 1], excel.Cells[2, 2]].Value2 = "询价单号：";
            excel.Cells[2, 3] =  orderModel["inquiryNo"];
            xSt.Range[excel.Cells[3, 1], excel.Cells[3, 2]].Merge(false);
            xSt.Range[excel.Cells[3, 1], excel.Cells[3, 2]].Value2 = "供货商：";
            xSt.Range[excel.Cells[4, 1], excel.Cells[4, 2]].Merge(false);
            xSt.Range[excel.Cells[4, 1], excel.Cells[4, 2]].Value2 = "联系方式：";
            xSt.Range[excel.Cells[5, 1], excel.Cells[5, 2]].Merge(false);
            xSt.Range[excel.Cells[5, 1], excel.Cells[5, 2]].Value2 = "运费负担：";
            xSt.Range[excel.Cells[6, 1], excel.Cells[6, 2]].Merge(false);
            xSt.Range[excel.Cells[6, 1], excel.Cells[6, 2]].Value2 = "价格有效期：";

            excel.Cells[7, 1] = "序号";
            excel.Cells[7, 2] = "产品型号";
            excel.Cells[7, 3] = "产品描述";
            excel.Cells[7, 4] = "数量";
            excel.Cells[7, 5] = "单位";
            excel.Cells[7, 6] = "报价品牌";
            excel.Cells[7, 7] = "报价型号";
            excel.Cells[7, 8] = "报价名称";
            excel.Cells[7, 9] = "含税单价";
            excel.Cells[7, 10] = "含税总价";
            excel.Cells[7, 11] = "交货期";
            excel.Cells[7, 12] = "最小起订量";
            excel.Cells[7, 13] = "备注";

            xSt.Range[excel.Cells[7, 1], excel.Cells[7, 13]].HorizontalAlignment = XlHAlign.xlHAlignCenter;//设置标题格式为居中对齐 
            xSt.Range[excel.Cells[7, 1], excel.Cells[7, 13]].Font.Bold = true;


            for (int i = 0; i < Convert.ToInt32(orderModel["number"]); i++)
            {
                string brandName = productModel[i]["brandName"].ToString().Trim();
                string strFuze = "";
                if (brandName.Contains('('))
                {
                    strFuze = "(" + brandName.Split('(')[1];
                }

                excel.Cells[i + 8, 1] = Convert.ToString(i + 1);
                excel.Cells[i + 8, 2] = productModel[i]["productModel"].ToString().Trim();
                excel.Cells[i + 8, 3] = brandName.Replace(strFuze, "") + ", " + productModel[i]["productName"].ToString().Trim();
                
                excel.Cells[i + 8, 4] = productModel[i]["quantity"].ToString().Trim();
                excel.Cells[i + 8, 5] = productModel[i]["unit"].ToString().Trim();

                xSt.Range[excel.Cells[i + 8, 1], excel.Cells[i + 8, 1]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                xSt.Range[excel.Cells[i + 8, 2], excel.Cells[i + 8, 3]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
                xSt.Range[excel.Cells[i + 8, 4], excel.Cells[i + 8, 5]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                xSt.Range[excel.Cells[i + 8, 6], excel.Cells[i + 8, 8]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
                xSt.Range[excel.Cells[i + 8, 9], excel.Cells[i + 8, 10]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
                xSt.Range[excel.Cells[i + 8, 11], excel.Cells[i + 8, 13]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;

                xSt.Range[excel.Cells[i + 8, 10], excel.Cells[i + 8, 10]].Formula = "=D" + (i + 8) + "*I" + (i + 8);
            }

            // 
            //加载一个合计行 
            // 
            int rowSum = Convert.ToInt32(orderModel["number"].ToString()) + 7;
            //int colSum = 2;


            xSt.Range[excel.Cells[rowSum + 1, 10], excel.Cells[rowSum + 1, 10]].Formula = "=SUM(J8:J" + rowSum + ")";
            xSt.Range[excel.Cells[rowSum+1, 10], excel.Cells[rowSum+1, 10]].Borders.LineStyle = 1;
            xSt.Range[excel.Cells[rowSum + 1, 10], excel.Cells[rowSum + 1, 10]].Borders.Weight = XlBorderWeight.xlThin;//设置左边线加粗 

            // 
            //绘制边框 
            // 
            xSt.Range[excel.Cells[7, 1], excel.Cells[rowSum, 13]].Borders.LineStyle = 1;
            xSt.Range[excel.Cells[7, 1], excel.Cells[rowSum, 1]].Borders[XlBordersIndex.xlEdgeLeft].Weight = XlBorderWeight.xlThin;//设置左边线加粗 
            xSt.Range[excel.Cells[7, 1], excel.Cells[7, 13]].Borders[XlBordersIndex.xlEdgeTop].Weight = XlBorderWeight.xlThin;//设置上边线加粗 
            xSt.Range[excel.Cells[7, 10], excel.Cells[rowSum, 13]].Borders[XlBordersIndex.xlEdgeRight].Weight = XlBorderWeight.xlThin;//设置右边线加粗 
            xSt.Range[excel.Cells[rowSum, 1], excel.Cells[rowSum, 13]].Borders[XlBordersIndex.xlEdgeBottom].Weight = XlBorderWeight.xlThin;//设置下边线加粗 

            xSt.Range[excel.Cells[2, 1], excel.Cells[rowSum + 1, 13]].RowHeight = 24;
            xSt.Range[excel.Cells[2, 1], excel.Cells[rowSum + 1, 13]].Font.Name = "微软雅黑";
            xSt.Range[excel.Cells[2, 1], excel.Cells[rowSum + 1, 13]].Font.Size = 9;

            //
            //冻结窗格
            //
            xSt.Application.ActiveWindow.SplitRow = 7;
            xSt.Application.ActiveWindow.SplitColumn = 0;
            xSt.Application.ActiveWindow.FreezePanes = true;
            //
            //设置最后一列的列宽为30
            //
            //xSt.Range[excel.Cells[1, colIndex], excel.Cells[rowSum, colIndex]].ColumnWidth = 30;
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