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
    public partial class salesInvoiceDetails : System.Web.UI.Page
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
            var number = Convert.ToInt32(m.number);

            OutputExcel(orderModel, number);
        }

        [DllImport("user32.dll", SetLastError = true)]
        static extern int GetWindowThreadProcessId(IntPtr hWnd, out int lpdwProcessId);

        public void OutputExcel(dynamic orderModel, int number)
        {
            GC.Collect();
            Application excel = new Application();
            _Workbook xBk = excel.Workbooks.Add(true);
            _Worksheet xSt = (_Worksheet)xBk.ActiveSheet;

            excel.DisplayAlerts = false;

            xSt.PageSetup.Orientation = XlPageOrientation.xlLandscape;
            xSt.PageSetup.LeftMargin = 0 / 0.035;
            xSt.PageSetup.RightMargin = 0 / 0.035;
            xSt.PageSetup.HeaderMargin = 0.8 / 0.035;
            xSt.PageSetup.FooterMargin = 0.8 / 0.035;
            xSt.PageSetup.TopMargin = 0.8 / 0.035;
            xSt.PageSetup.BottomMargin = 0.8 / 0.035;

            excel.Cells.Font.Name = "微软雅黑";
            excel.Cells.Font.Size = 8;

            xSt.Range[excel.Cells[1, 1], excel.Cells[1, 27]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            xSt.Range[excel.Cells[1, 1], excel.Cells[1, 27]].Font.Bold = true;

            excel.Cells[1, 1] = "编号";
            excel.Cells[1, 2] = "出票公司";
            excel.Cells[1, 3] = "单据编号";
            excel.Cells[1, 4] = "销售员";
            excel.Cells[1, 5] = "开票日期";
            excel.Cells[1, 6] = "票号";
            excel.Cells[1, 7] = "类型";
            excel.Cells[1, 8] = "价税合计";
            excel.Cells[1, 9] = "制单人";
            excel.Cells[1, 10] = "状态";
            excel.Cells[1, 11] = "订单型号";
            excel.Cells[1, 12] = "订单名称";
            excel.Cells[1, 13] = "订货号";
            excel.Cells[1, 14] = "订货名称";
            excel.Cells[1, 15] = "税收编码";
            excel.Cells[1, 16] = "进项型号";
            excel.Cells[1, 17] = "进项名称";
            excel.Cells[1, 18] = "订单数量";
            excel.Cells[1, 19] = "单位";
            excel.Cells[1, 20] = "销售单价";
            excel.Cells[1, 21] = "销售总价";
            excel.Cells[1, 22] = "已付金额";
            excel.Cells[1, 23] = "进货数量";
            excel.Cells[1, 24] = "进货单价";
            excel.Cells[1, 25] = "进货总价";
            excel.Cells[1, 26] = "进货含税";
            excel.Cells[1, 27] = "采购状态";

            int j = 1;
            for (int i = 0; i < number; i++)
            {
                j++;

                xSt.Range[excel.Cells[j, 2], excel.Cells[j, 4]].NumberFormat = "@";
                xSt.Range[excel.Cells[j, 6], excel.Cells[j, 7]].NumberFormat = "@";
                xSt.Range[excel.Cells[j, 9], excel.Cells[j, 17]].NumberFormat = "@";
                xSt.Range[excel.Cells[j, 26], excel.Cells[j, 27]].NumberFormat = "@";

                xSt.Range[excel.Cells[j, 1], excel.Cells[j, 26]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
                xSt.Range[excel.Cells[j, 1], excel.Cells[j, 2]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                xSt.Range[excel.Cells[j, 4], excel.Cells[j, 5]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                xSt.Range[excel.Cells[j, 7], excel.Cells[j, 7]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                xSt.Range[excel.Cells[j, 8], excel.Cells[j, 8]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
                xSt.Range[excel.Cells[j, 9], excel.Cells[j, 10]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                xSt.Range[excel.Cells[j, 18], excel.Cells[j, 19]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                xSt.Range[excel.Cells[j, 20], excel.Cells[j, 22]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
                xSt.Range[excel.Cells[j, 23], excel.Cells[j, 23]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                xSt.Range[excel.Cells[j, 24], excel.Cells[j, 25]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
                xSt.Range[excel.Cells[j, 26], excel.Cells[j, 27]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

                excel.Cells[j, 1] = orderModel[i]["no"];
                excel.Cells[j, 2] = orderModel[i]["companyName"];
                excel.Cells[j, 3] = orderModel[i]["orderNo"];
                excel.Cells[j, 4] = orderModel[i]["personInChargeName"];
                excel.Cells[j, 5] = orderModel[i]["invoiceDate"];
                excel.Cells[j, 6] = orderModel[i]["invoiceNo"];
                excel.Cells[j, 7] = orderModel[i]["invoiceType"];
                excel.Cells[j, 8] = orderModel[i]["total"].ToString();
                excel.Cells[j, 9] = orderModel[i]["createBy"];
                excel.Cells[j, 10] = orderModel[i]["state"];
                excel.Cells[j, 11] = orderModel[i]["quotationProductModel"];
                excel.Cells[j, 12] = orderModel[i]["quotationProductName"];
                excel.Cells[j, 13] = orderModel[i]["productNo"];
                excel.Cells[j, 14] = orderModel[i]["productName"];
                excel.Cells[j, 15] = orderModel[i]["taxEncoding"];
                excel.Cells[j, 16] = orderModel[i]["taxNo"];
                excel.Cells[j, 17] = orderModel[i]["taxName"];
                excel.Cells[j, 18] = orderModel[i]["quantity"].ToString();
                excel.Cells[j, 19] = orderModel[i]["unit"];
                excel.Cells[j, 20] = orderModel[i]["unitPrice"].ToString();
                excel.Cells[j, 21] = orderModel[i]["totalPrice"].ToString();
                excel.Cells[j, 22] = orderModel[i]["payTotal"].ToString();
                excel.Cells[j, 23] = orderModel[i]["purchaseQuantity"].ToString();
                excel.Cells[j, 24] = orderModel[i]["purchaseUnitPrice"].ToString();
                excel.Cells[j, 25] = orderModel[i]["purchaseTotalPrice"].ToString();
                excel.Cells[j, 26] = orderModel[i]["purchaseInvoiceType"];
                excel.Cells[j, 27] = orderModel[i]["purchaseState"];
            }

            xSt.Range[excel.Cells[1, 1], excel.Cells[j, 27]].Columns.AutoFit();//行高根据内容自动调整

            xSt.Range[excel.Cells[1, 1], excel.Cells[j, 27]].Borders.LineStyle = 1;
            xSt.Range[excel.Cells[1, 1], excel.Cells[j, 1]].Borders[XlBordersIndex.xlEdgeLeft].Weight = XlBorderWeight.xlThin;//设置左边线加粗 
            xSt.Range[excel.Cells[1, 1], excel.Cells[1, 27]].Borders[XlBordersIndex.xlEdgeTop].Weight = XlBorderWeight.xlThin;//设置上边线加粗 
            xSt.Range[excel.Cells[1, 27], excel.Cells[j, 27]].Borders[XlBordersIndex.xlEdgeRight].Weight = XlBorderWeight.xlThin;//设置右边线加粗 
            xSt.Range[excel.Cells[j, 1], excel.Cells[j, 27]].Borders[XlBordersIndex.xlEdgeBottom].Weight = XlBorderWeight.xlThin;//设置下边线加粗 

            xSt.Range[excel.Cells[1, 1], excel.Cells[j, 27]].RowHeight = 20;

            excel.Visible = true;

            string path = "";
            string filename = "";

            filename = "salesInvoiceDetails" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx";
            path = Server.MapPath("~/") + "temp\\" + filename;

            //保存excel
            xBk.SaveCopyAs(path);

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