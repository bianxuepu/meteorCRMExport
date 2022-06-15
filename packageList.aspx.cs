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
    public partial class packageList : System.Web.UI.Page
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

            var filetype = Request.QueryString["type"];
            if (filetype == null || filetype == "")
            {
                filetype = "excel";
            }

            var orderModel = m.orderModel;
            var companyModel = m.companyModel;
            var customerModel = m.customerModel;
            var productModel = m.productModel;

            OutputExcel(orderModel, companyModel, customerModel, productModel, filetype);
        }

        [DllImport("user32.dll", SetLastError = true)]
        private static extern int GetWindowThreadProcessId(IntPtr hWnd, out int lpdwProcessId);

        public void OutputExcel(dynamic orderModel, dynamic companyModel, dynamic customerModel, dynamic productModel, string strType)
        {
            GC.Collect();
            Application excel = new Application();
            _Workbook xBk = excel.Workbooks.Add(true);
            _Worksheet xSt = (_Worksheet)xBk.ActiveSheet;

            excel.DisplayAlerts = false;

            xSt.Name = "装箱单";
            xSt.PageSetup.LeftMargin = 300.0 / 7.0;
            xSt.PageSetup.RightMargin = 200.0 / 7.0;
            xSt.PageSetup.HeaderMargin = 0.0;
            xSt.PageSetup.FooterMargin = 0.0;
            xSt.PageSetup.TopMargin = 200.0 / 7.0;
            xSt.PageSetup.BottomMargin = 300.0 / 7.0;

            excel.Cells.Font.Name = "微软雅黑";
            excel.Cells.Font.Size = 9;
            excel.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;

            xSt.Range[excel.Cells[1, 1], excel.Cells[1, 1]].ColumnWidth = 4;
            xSt.Range[excel.Cells[1, 2], excel.Cells[1, 2]].ColumnWidth = 21;
            xSt.Range[excel.Cells[1, 3], excel.Cells[1, 3]].ColumnWidth = 34.5;
            xSt.Range[excel.Cells[1, 4], excel.Cells[1, 4]].ColumnWidth = 8;
            xSt.Range[excel.Cells[1, 5], excel.Cells[1, 5]].ColumnWidth = 7.5;
            xSt.Range[excel.Cells[1, 6], excel.Cells[1, 6]].ColumnWidth = 12;

            xSt.Range[excel.Cells[1, 1], excel.Cells[1, 6]].RowHeight = 24;
            xSt.Range[excel.Cells[1, 1], excel.Cells[1, 6]].Merge(false);
            xSt.Range[excel.Cells[1, 1], excel.Cells[1, 6]].Font.Size = 12;
            xSt.Range[excel.Cells[1, 1], excel.Cells[1, 6]].Value2 = "装箱 / 出库单";
            xSt.Range[excel.Cells[1, 1], excel.Cells[1, 6]].Borders[XlBordersIndex.xlEdgeBottom].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            xSt.Range[excel.Cells[1, 1], excel.Cells[1, 6]].Borders[XlBordersIndex.xlEdgeBottom].Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlHairline;

            xSt.Range[excel.Cells[2, 1], excel.Cells[2, 8]].RowHeight = 7.5;

            xSt.Range[excel.Cells[3, 1], excel.Cells[4, 6]].RowHeight = 30;
            xSt.Cells[3, 1] = orderModel["orderNo"];
            xSt.Cells[3, 1].Font.Size = 18;
            xSt.Cells[3, 1].Font.Bold = true;
            xSt.Cells[3, 6] = customerModel["customerKeyName"];
            xSt.Cells[3, 6].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
            xSt.Cells[3, 6].Font.Size = 18;

            xSt.Cells[4, 1] = "页码：1-1";
            xSt.Cells[4, 6] = orderModel["customerOrderNo"];
            xSt.Cells[4, 6].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
            xSt.Cells[4, 6].Font.Size = 18;
            xSt.Cells[4, 6].Font.Bold = true;

            xSt.Cells[5, 1] = "订单共" + orderModel["number"] + "项产品";
            xSt.Cells[5, 6] = orderModel["userName"];
            xSt.Cells[5, 6].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
            xSt.Cells[5, 6].Font.Size = 12;

            xSt.Range[excel.Cells[6, 1], excel.Cells[6, 6]].RowHeight = 24.75;

            xSt.Range[excel.Cells[7, 1], excel.Cells[7, 6]].RowHeight = 25.5;
            xSt.Range[excel.Cells[7, 1], excel.Cells[7, 6]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            xSt.Cells[7, 1] = "序号";
            xSt.Cells[7, 2] = "产品货号";
            xSt.Cells[7, 3] = "产品描述";
            xSt.Cells[7, 4] = "送货数量";
            xSt.Cells[7, 5] = "单位";
            xSt.Cells[7, 6] = "备注";
            xSt.Range[excel.Cells[7, 1], excel.Cells[7, 6]].Borders[XlBordersIndex.xlEdgeBottom].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            xSt.Range[excel.Cells[7, 1], excel.Cells[7, 6]].Borders[XlBordersIndex.xlEdgeBottom].Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlHairline;

            int j = 8;
            for (int i = 0; i < Convert.ToInt32(orderModel["number"]); i++)
            {
                xSt.Cells[j, 1].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                xSt.Cells[j, 4].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                xSt.Cells[j, 5].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

                xSt.Range[excel.Cells[j, 1], excel.Cells[j, 8]].RowHeight = 36;
                xSt.Cells[j, 1] = Convert.ToString(i + 1);

                excel.Cells[j, 2].WrapText = true;
                excel.Cells[j, 2].NumberFormat = "@";
                excel.Cells[j, 2] = (Convert.ToBoolean(productModel[i]["Sales_OrderProduct"]["isQuotation"].ToString()) ? productModel[i]["Sales_OrderProduct"]["quotationProductModel"].ToString().Trim() : productModel[i]["Product_Product"]["proNo"].ToString().Trim()) + " ";
                xSt.Cells[j, 2].Font.Size = 12;
                xSt.Cells[j, 2].Font.Bold = true;

                excel.Cells[j, 3].WrapText = true;
                excel.Cells[j, 3] = (Convert.ToBoolean(productModel[i]["Sales_OrderProduct"]["isQuotation"].ToString()) ? productModel[i]["Sales_OrderProduct"]["quotationProductName"].ToString().Trim() : productModel[i]["Product_Brand"]["chineseName"].ToString().Trim() + (string.IsNullOrEmpty(productModel[i]["Product_Brand"]["englishName"].ToString().Trim()) ? "" : "/" + productModel[i]["Product_Brand"]["englishName"].ToString().Trim()) + "　" + productModel[i]["Product_Product"]["name"].ToString().Trim() + (string.IsNullOrEmpty(productModel[i]["Product_Product"]["ordNo"].ToString().Trim()) ? "" : "　" + productModel[i]["Product_Product"]["ordNo"].ToString().Trim()) + (string.IsNullOrEmpty(productModel[i]["Product_Product"]["package"].ToString().Trim()) ? "" : "　" + productModel[i]["Product_Product"]["package"].ToString().Trim())) + " ";

                xSt.Cells[j, 4] = productModel[i]["Sales_OutboundOrderProduct"]["quantity"].ToString("N2");
                xSt.Cells[j, 5] = productModel[i]["Product_Product"]["unit"];
                xSt.Cells[j, 6] = productModel[i]["Sales_OutboundOrderProduct"]["remark"];

                xSt.Range[excel.Cells[j, 1], excel.Cells[j, 8]].Borders[XlBordersIndex.xlEdgeBottom].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlDot;
                xSt.Range[excel.Cells[j, 1], excel.Cells[j, 8]].Borders[XlBordersIndex.xlEdgeBottom].Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlHairline;

                j++;
            }

            xSt.Range[excel.Cells[j, 1], excel.Cells[j, 6]].RowHeight = 40.5;
            excel.Cells[j, 3].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignBottom;
            excel.Cells[j, 3].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
            excel.Cells[j, 3] = "出库签字：";

            excel.Cells[j, 5].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignBottom;
            excel.Cells[j, 5].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
            excel.Cells[j, 5] = "日期：";

            string path = "";
            string filename = "";

            if (strType.ToLower() == "excel")
            {
                filename = orderModel["_id"] + "_packageList.xlsx";
                path = Server.MapPath("~/") + "temp\\" + orderModel["_id"] + "_packageList.xlsx";

                //保存excel
                xBk.SaveCopyAs(path);
            }
            else
            {
                filename = orderModel["_id"] + "_packageList.pdf";
                path = Server.MapPath("~/") + "temp\\" + orderModel["_id"] + "_packageList.pdf";

                //保存pdf
                xBk.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, path.Replace(".xlsx", ".pdf"), XlFixedFormatQuality.xlQualityStandard, true, false, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                //加密pdf
                //Encry(path, path.Replace(".pdf", "-1.pdf"));
                //path = path.Replace(".pdf", "-1.pdf");
            }



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