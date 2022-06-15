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
    public partial class innerInvoice : System.Web.UI.Page
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

            xSt.PageSetup.Orientation = XlPageOrientation.xlLandscape;
            xSt.PageSetup.LeftMargin = 0.5 / 0.035;
            xSt.PageSetup.RightMargin = 0.5 / 0.035;
            xSt.PageSetup.HeaderMargin = 0.8 / 0.035;
            xSt.PageSetup.FooterMargin = 0.8 / 0.035;
            xSt.PageSetup.TopMargin = 0.8 / 0.035;
            xSt.PageSetup.BottomMargin = 0.8 / 0.035;

            excel.Cells.Font.Name = "微软雅黑";
            excel.Cells.Font.Size = 8;

            xSt.Range[excel.Cells[1, 1], excel.Cells[1, 1]].ColumnWidth = 12;
            xSt.Range[excel.Cells[1, 2], excel.Cells[1, 2]].ColumnWidth = 12;
            xSt.Range[excel.Cells[1, 3], excel.Cells[1, 3]].ColumnWidth = 14;
            xSt.Range[excel.Cells[1, 4], excel.Cells[1, 4]].ColumnWidth = 8.5;
            xSt.Range[excel.Cells[1, 5], excel.Cells[1, 5]].ColumnWidth = 8.5;
            xSt.Range[excel.Cells[1, 6], excel.Cells[1, 6]].ColumnWidth = 12;
            xSt.Range[excel.Cells[1, 7], excel.Cells[1, 7]].ColumnWidth = 14;
            xSt.Range[excel.Cells[1, 8], excel.Cells[1, 8]].ColumnWidth = 8.5;
            xSt.Range[excel.Cells[1, 9], excel.Cells[1, 9]].ColumnWidth = 8.5;
            xSt.Range[excel.Cells[1, 10], excel.Cells[1, 10]].ColumnWidth = 12;
            xSt.Range[excel.Cells[1, 11], excel.Cells[1, 11]].ColumnWidth = 14;
            xSt.Range[excel.Cells[1, 12], excel.Cells[1, 12]].ColumnWidth = 5;
            xSt.Range[excel.Cells[1, 13], excel.Cells[1, 13]].ColumnWidth = 6;
            xSt.Range[excel.Cells[1, 14], excel.Cells[1, 14]].ColumnWidth = 9;
            xSt.Range[excel.Cells[1, 15], excel.Cells[1, 14]].ColumnWidth = 9;
            xSt.Range[excel.Cells[1, 16], excel.Cells[1, 16]].ColumnWidth = 9;
            xSt.Range[excel.Cells[1, 17], excel.Cells[1, 17]].ColumnWidth = 9;
            xSt.Range[excel.Cells[1, 18], excel.Cells[1, 18]].ColumnWidth = 16;
            xSt.Range[excel.Cells[1, 19], excel.Cells[1, 19]].ColumnWidth = 5;

            xSt.Range[excel.Cells[1, 1], excel.Cells[1, 19]].RowHeight = 1;

            xSt.Range[excel.Cells[2, 1], excel.Cells[2, 19]].RowHeight = 30;
            xSt.Range[excel.Cells[2, 1], excel.Cells[2, 19]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            xSt.Range[excel.Cells[2, 1], excel.Cells[2, 19]].Merge(false);
            xSt.Range[excel.Cells[2, 1], excel.Cells[2, 19]].Font.Size = 13;
            xSt.Range[excel.Cells[2, 1], excel.Cells[2, 19]].Font.Bold = true;
            xSt.Range[excel.Cells[2, 1], excel.Cells[2, 19]].Value2 = orderModel["orderNo"] + " 内部发票";

            xSt.Range[excel.Cells[3, 1], excel.Cells[5, 19]].RowHeight = 24;

            xSt.Range[excel.Cells[3, 1], excel.Cells[3, 3]].Merge(false);
            xSt.Range[excel.Cells[3, 1], excel.Cells[3, 3]].Value2 = "开票公司：" + orderModel["companyName"].ToString().Trim();

            xSt.Range[excel.Cells[3, 4], excel.Cells[3, 9]].Merge(false);
            xSt.Range[excel.Cells[3, 4], excel.Cells[3, 9]].Value2 = "收票公司：" + orderModel["otherCompanyName"].ToString().Trim();

            xSt.Range[excel.Cells[3, 11], excel.Cells[3, 14]].Merge(false);
            //if (Convert.ToBoolean(orderModel["invoiceType"].ToString()))
            //{
            //    xSt.Range[excel.Cells[3, 11], excel.Cells[3, 14]].Value2 = "发票类型：增值税发票";
            //}
            //else
            //{
            //    xSt.Range[excel.Cells[3, 11], excel.Cells[3, 14]].Font.Color = System.Drawing.ColorTranslator.ToOle(Color.Red);
            //    xSt.Range[excel.Cells[3, 11], excel.Cells[3, 14]].Value2 = "发票类型：普通发票";
            //}
            xSt.Range[excel.Cells[3, 11], excel.Cells[3, 14]].Value2 = "发票类型：" + orderModel["invoiceType"].ToString();

            xSt.Range[excel.Cells[4, 1], excel.Cells[4, 3]].Merge(false);
            xSt.Range[excel.Cells[4, 1], excel.Cells[4, 3]].Value2 = "价税合计：" + Convert.ToDecimal(orderModel["total"].ToString()).ToString("N2");

            //xSt.PageSetup.PrintTitleRows = "$8:$8";  标题行 
            xSt.Range[excel.Cells[6, 1], excel.Cells[6, 19]].RowHeight = 22;
            xSt.Range[excel.Cells[6, 1], excel.Cells[6, 19]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            excel.Cells[6, 1] = "内部销售单号";
            excel.Cells[6, 2] = "订货号";
            excel.Cells[6, 3] = "描述";
            excel.Cells[6, 4] = "税收分类编码";
            excel.Cells[6, 5] = "税收分类名称";
            excel.Cells[6, 6] = "进项型号";
            excel.Cells[6, 7] = "进项名称";
            excel.Cells[6, 8] = "产品—税收分类编码";
            excel.Cells[6, 9] = "产品—税收分类名称";
            excel.Cells[6, 10] = "产品—开票型号";
            excel.Cells[6, 11] = "产品—开票名称";
            excel.Cells[6, 12] = "单位";
            excel.Cells[6, 13] = "订单数量";
            excel.Cells[6, 14] = "未税单价";
            excel.Cells[6, 15] = "未税总价";
            excel.Cells[6, 16] = "含税单价";
            excel.Cells[6, 17] = "含税总价";
            excel.Cells[6, 18] = "备注";
            excel.Cells[6, 19] = "采购人";

            int j = 6;
            decimal totalNoTax = 0;
            decimal total = 0;
            string innerOrderNo = "";
            for (int i = 0; i < Convert.ToInt32(orderModel["number"]); i++)
            {
                j++;

                xSt.Range[excel.Cells[j, 1], excel.Cells[j, 19]].WrapText = true;

                xSt.Range[excel.Cells[j, 1], excel.Cells[j, 12]].NumberFormat = "@";
                excel.Cells[j, 1] = productModel[i]["Inner_InvoiceProduct"]["innerOrderNo"].ToString().Trim();
                if (innerOrderNo.IndexOf(productModel[i]["Inner_InvoiceProduct"]["innerOrderNo"].ToString().Trim()) == -1)
                {
                    innerOrderNo += productModel[i]["Inner_InvoiceProduct"]["innerOrderNo"].ToString().Trim() + '；';
                }
                excel.Cells[j, 2] = productModel[i]["Product_Product"]["proNo"].ToString().Trim() + " ";
                excel.Cells[j, 3] = productModel[i]["Product_Brand"]["chineseName"].ToString().Trim() + (string.IsNullOrEmpty(productModel[i]["Product_Brand"]["englishName"].ToString().Trim()) ? "" : "/" + productModel[i]["Product_Brand"]["englishName"].ToString().Trim()) + "　" + productModel[i]["Product_Product"]["name"].ToString().Trim() + (string.IsNullOrEmpty(productModel[i]["Product_Product"]["ordNo"].ToString().Trim()) ? "" : "　" + productModel[i]["Product_Product"]["ordNo"].ToString().Trim()) + (string.IsNullOrEmpty(productModel[i]["Product_Product"]["package"].ToString().Trim()) ? "" : "　" + productModel[i]["Product_Product"]["package"].ToString().Trim());
                try
                {
                    excel.Cells[j, 4] = Convert.ToString(productModel[i]["Purchase_InvoiceProduct"]["taxEncodingNo"].ToString().Trim() + "00000000000000000000").Substring(0,19).Replace("0000000000000000000","");
                    excel.Cells[j, 5] = productModel[i]["Purchase_InvoiceProduct"]["taxEncoding"].ToString().Trim();
                    excel.Cells[j, 6] = productModel[i]["Purchase_InvoiceProduct"]["taxNo"].ToString().Trim();
                    excel.Cells[j, 7] = productModel[i]["Purchase_InvoiceProduct"]["taxName"].ToString().Trim();
                }
                catch { }
                excel.Cells[j, 8] = Convert.ToString(productModel[i]["Product_Product"]["taxEncodingNo"].ToString().Trim() + "00000000000000000000").Substring(0, 19).Replace("0000000000000000000", "");
                excel.Cells[j, 9] = productModel[i]["Product_Product"]["taxEncoding"].ToString().Trim();
                excel.Cells[j, 10] = productModel[i]["Product_Product"]["taxNo"].ToString().Trim();
                excel.Cells[j, 11] = productModel[i]["Product_Product"]["taxName"].ToString().Trim();

                xSt.Range[excel.Cells[j, 12], excel.Cells[j, 12]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                excel.Cells[j, 12] = productModel[i]["Product_Product"]["unit"].ToString().Trim();
                excel.Cells[j, 13] = (string.IsNullOrEmpty(productModel[i]["Inner_InvoiceProduct"]["quantity"].ToString()) ? 0 : Convert.ToDecimal(productModel[i]["Inner_InvoiceProduct"]["quantity"].ToString()).ToString("N2"));

                xSt.Range[excel.Cells[j, 14], excel.Cells[j, 17]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
                excel.Cells[j, 14] = (string.IsNullOrEmpty(productModel[i]["Inner_InvoiceProduct"]["unitPriceNoTax"].ToString()) ? 0 : productModel[i]["Inner_InvoiceProduct"]["unitPriceNoTax"].ToString());
                excel.Cells[j, 15] = (string.IsNullOrEmpty(productModel[i]["Inner_InvoiceProduct"]["totalPriceNoTax"].ToString()) ? 0 : productModel[i]["Inner_InvoiceProduct"]["totalPriceNoTax"].ToString());
                if (!string.IsNullOrEmpty(productModel[i]["Inner_InvoiceProduct"]["totalPriceNoTax"].ToString()))
                {
                    totalNoTax += Convert.ToDecimal(productModel[i]["Inner_InvoiceProduct"]["totalPriceNoTax"].ToString());
                }
                excel.Cells[j, 16] = (string.IsNullOrEmpty(productModel[i]["Inner_InvoiceProduct"]["unitPrice"].ToString()) ? 0 : productModel[i]["Inner_InvoiceProduct"]["unitPrice"].ToString());
                excel.Cells[j, 17] = (string.IsNullOrEmpty(productModel[i]["Inner_InvoiceProduct"]["totalPrice"].ToString()) ? 0 : productModel[i]["Inner_InvoiceProduct"]["totalPrice"].ToString());
                if (!string.IsNullOrEmpty(productModel[i]["Inner_InvoiceProduct"]["totalPrice"].ToString()))
                {
                    total += Convert.ToDecimal(productModel[i]["Inner_InvoiceProduct"]["totalPrice"].ToString());
                }

                excel.Cells[j, 18] = productModel[i]["Inner_InvoiceProduct"]["remark"].ToString().Trim();

                xSt.Range[excel.Cells[j, 19], excel.Cells[j, 19]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

                try
                {
                    excel.Cells[j, 19] = productModel[i]["Purchase_Order"]["personInChargeName"].ToString().Trim();
                }
                catch
                { }


                xSt.Range[excel.Cells[j, 1], excel.Cells[j, 19]].EntireRow.AutoFit();//行高根据内容自动调整

                //如果不换行，行高自适应后，行高小于23，则最低行高为23
                if (Convert.ToInt32(xSt.Range[excel.Cells[j, 1], excel.Cells[j, 19]].RowHeight) < 22)
                {
                    xSt.Range[excel.Cells[j, 1], excel.Cells[j, 19]].RowHeight = 22;
                }
            }

            xSt.Range[excel.Cells[j + 1, 15], excel.Cells[j + 1, 15]].HorizontalAlignment = XlHAlign.xlHAlignRight;//设置标题格式为居中对齐 
            xSt.Range[excel.Cells[j + 1, 15], excel.Cells[j + 1, 15]].Formula = "=SUM(O7:O" + j + ")";

            xSt.Range[excel.Cells[j + 1, 17], excel.Cells[j + 1, 17]].HorizontalAlignment = XlHAlign.xlHAlignRight;//设置标题格式为居中对齐 
            xSt.Range[excel.Cells[j + 1, 17], excel.Cells[j + 1, 17]].Formula = "=SUM(Q7:Q" + j + ")";

            xSt.Range[excel.Cells[6, 1], excel.Cells[j, 19]].Borders.LineStyle = 1;
            xSt.Range[excel.Cells[6, 1], excel.Cells[j, 1]].Borders[XlBordersIndex.xlEdgeLeft].Weight = XlBorderWeight.xlThin;//设置左边线加粗 
            xSt.Range[excel.Cells[6, 1], excel.Cells[6, 19]].Borders[XlBordersIndex.xlEdgeTop].Weight = XlBorderWeight.xlThin;//设置上边线加粗 
            xSt.Range[excel.Cells[6, 19], excel.Cells[j, 19]].Borders[XlBordersIndex.xlEdgeRight].Weight = XlBorderWeight.xlThin;//设置右边线加粗 
            xSt.Range[excel.Cells[j, 1], excel.Cells[j, 19]].Borders[XlBordersIndex.xlEdgeBottom].Weight = XlBorderWeight.xlThin;//设置下边线加粗 

            xSt.Range[excel.Cells[j + 1, 15], excel.Cells[j + 1, 15]].Borders.LineStyle = 1;
            xSt.Range[excel.Cells[j + 1, 15], excel.Cells[j + 1, 15]].Borders[XlBordersIndex.xlEdgeLeft].Weight = XlBorderWeight.xlThin;//设置左边线加粗 
            xSt.Range[excel.Cells[j + 1, 15], excel.Cells[j + 1, 15]].Borders[XlBordersIndex.xlEdgeTop].Weight = XlBorderWeight.xlThin;//设置上边线加粗 
            xSt.Range[excel.Cells[j + 1, 15], excel.Cells[j + 1, 15]].Borders[XlBordersIndex.xlEdgeRight].Weight = XlBorderWeight.xlThin;//设置右边线加粗 
            xSt.Range[excel.Cells[j + 1, 15], excel.Cells[j + 1, 15]].Borders[XlBordersIndex.xlEdgeBottom].Weight = XlBorderWeight.xlThin;//设置下边线加粗 

            xSt.Range[excel.Cells[j + 1, 17], excel.Cells[j + 1, 17]].Borders.LineStyle = 1;
            xSt.Range[excel.Cells[j + 1, 17], excel.Cells[j + 1, 17]].Borders[XlBordersIndex.xlEdgeLeft].Weight = XlBorderWeight.xlThin;//设置左边线加粗 
            xSt.Range[excel.Cells[j + 1, 17], excel.Cells[j + 1, 17]].Borders[XlBordersIndex.xlEdgeTop].Weight = XlBorderWeight.xlThin;//设置上边线加粗 
            xSt.Range[excel.Cells[j + 1, 17], excel.Cells[j + 1, 17]].Borders[XlBordersIndex.xlEdgeRight].Weight = XlBorderWeight.xlThin;//设置右边线加粗 
            xSt.Range[excel.Cells[j + 1, 17], excel.Cells[j + 1, 17]].Borders[XlBordersIndex.xlEdgeBottom].Weight = XlBorderWeight.xlThin;//设置下边线加粗 

            xSt.Range[excel.Cells[j + 1, 1], excel.Cells[j + 1, 19]].RowHeight = 22;

            xSt.Range[excel.Cells[4, 4], excel.Cells[4, 9]].Merge(false);
            xSt.Range[excel.Cells[4, 4], excel.Cells[4, 9]].Value2 = "税前金额：" + totalNoTax.ToString();

            xSt.Range[excel.Cells[4, 11], excel.Cells[4, 14]].Merge(false);
            xSt.Range[excel.Cells[4, 11], excel.Cells[4, 14]].Value2 = "税　　额：" + (total - totalNoTax).ToString();

            xSt.Range[excel.Cells[5, 1], excel.Cells[5, 14]].Merge(false);
            xSt.Range[excel.Cells[5, 1], excel.Cells[5, 14]].Value2 = "内部销售单号：" + innerOrderNo.TrimEnd('；');

            excel.Visible = true;

            string path = "";
            string filename = "";

            filename = orderModel["_id"] + ".xlsx";
            path = Server.MapPath("~/") + "temp\\" + orderModel["_id"] + ".xlsx";

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