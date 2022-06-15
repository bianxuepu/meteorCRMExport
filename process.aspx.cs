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
    public partial class process : System.Web.UI.Page
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
            var companyModel = m.companyModel;
            var customerModel = m.customerModel;
            var productModel = m.productModel;
            var userModel = m.userModel;

            OutputExcel(orderModel, companyModel, customerModel, productModel, userModel);
        }

        [DllImport("user32.dll", SetLastError = true)]
        static extern int GetWindowThreadProcessId(IntPtr hWnd, out int lpdwProcessId);

        public void OutputExcel(dynamic orderModel, dynamic companyModel, dynamic customerModel, dynamic productModel, dynamic userModel)
        {
            // 
            // TODO: 在此处添加构造函数逻辑 
            // 
            //dv为要输出到Excel的数据，str为标题名称 
            GC.Collect();
            Application excel;// = new Application(); 
            int rowIndex = 1;

            _Workbook xBk;
            _Worksheet xSt;
            _Worksheet xSt1;
            _Worksheet xSt2;

            excel = new Application();

            excel.DisplayAlerts = false;

            xBk = excel.Workbooks.Add(true);

            //#region 邮件发货清单

            //xSt2 = (_Worksheet)xBk.Worksheets.Add(Type.Missing, Type.Missing, 1, Type.Missing);

            //xSt2.Name = "邮件发货清单";

            //xSt2.PageSetup.LeftMargin = 1.9 / 0.035;
            //xSt2.PageSetup.RightMargin = 1.9 / 0.035;
            //xSt2.PageSetup.HeaderMargin = 0 / 0.035;
            //xSt2.PageSetup.FooterMargin = 0 / 0.035;
            //xSt2.PageSetup.TopMargin = 1.8 / 0.035;
            //xSt2.PageSetup.BottomMargin = 1.8 / 0.035;

            //excel.Cells.Font.Name = "微软雅黑";
            //excel.Cells.Font.Size = 9;

            //xSt2.Range[excel.Cells[1, 1], excel.Cells[1, 1]].ColumnWidth = 15;
            //xSt2.Range[excel.Cells[1, 2], excel.Cells[1, 2]].ColumnWidth = 44;
            //xSt2.Range[excel.Cells[1, 3], excel.Cells[1, 3]].ColumnWidth = 8;
            //xSt2.Range[excel.Cells[1, 4], excel.Cells[1, 4]].ColumnWidth = 8;

            //xSt2.Range[excel.Cells[1, 1], excel.Cells[1, 4]].RowHeight = 29;
            //xSt2.Shapes.AddPicture(Server.MapPath("~/").ToString().Trim() + "image\\" + companyModel["nid"].ToString().Trim() + ".gif", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoTrue, 4, 4, 150, 13);

            //xSt2.Range[excel.Cells[2, 1], excel.Cells[2, 4]].RowHeight = 44;
            //xSt2.Range[excel.Cells[2, 1], excel.Cells[2, 4]].Merge(false);
            //xSt2.Range[excel.Cells[2, 1], excel.Cells[2, 4]].Value2 = "发货清单";
            //xSt2.Range[excel.Cells[2, 1], excel.Cells[2, 4]].Font.Bold = true;
            //xSt2.Range[excel.Cells[2, 1], excel.Cells[2, 4]].Font.Size = 14;
            //xSt2.Range[excel.Cells[2, 1], excel.Cells[2, 4]].Borders[XlBordersIndex.xlEdgeBottom].Weight = XlBorderWeight.xlThin;
            //xSt2.Range[excel.Cells[2, 1], excel.Cells[2, 4]].Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
            //Color c1 = Color.FromArgb(242, 242, 242);
            //xSt2.Range[excel.Cells[2, 1], excel.Cells[2, 4]].Interior.Color = System.Drawing.ColorTranslator.ToOle(c1);


            //xSt2.Range[excel.Cells[3, 1], excel.Cells[4, 4]].RowHeight = 20;
            //xSt2.Range[excel.Cells[3, 1], excel.Cells[3, 4]].Merge(false);
            //xSt2.Range[excel.Cells[3, 1], excel.Cells[3, 4]].Value2 = "您订购的以下产品已从我司仓库发出。（交货日期可能会受到天气或运输等不可控因素影响）";

            //xSt2.Range[excel.Cells[5, 1], excel.Cells[8, 4]].RowHeight = 16;
            //excel.Cells[5, 1] = "您的订单号：";
            //xSt2.Range[excel.Cells[5, 2], excel.Cells[5, 4]].Merge(false);
            //xSt2.Range[excel.Cells[5, 2], excel.Cells[5, 4]].NumberFormat = "@";
            //xSt2.Range[excel.Cells[5, 2], excel.Cells[5, 4]].Value2 = orderModel["customerOrderNo"].ToString().Trim();

            //excel.Cells[6, 1] = "客户名：";
            //xSt2.Range[excel.Cells[6, 2], excel.Cells[6, 4]].Merge(false);
            //xSt2.Range[excel.Cells[6, 2], excel.Cells[6, 4]].NumberFormat = "@";
            //xSt2.Range[excel.Cells[6, 2], excel.Cells[6, 4]].Value2 = customerModel["customerName"].ToString().Trim();

            //excel.Cells[7, 1] = "采购人：";
            //xSt2.Range[excel.Cells[7, 2], excel.Cells[7, 4]].Merge(false);
            //xSt2.Range[excel.Cells[7, 2], excel.Cells[7, 4]].NumberFormat = "@";
            //xSt2.Range[excel.Cells[7, 2], excel.Cells[7, 4]].Value2 = orderModel["purchaserName"].ToString().Trim();

            //excel.Cells[8, 1] = "申请人：";
            //xSt2.Range[excel.Cells[8, 2], excel.Cells[8, 4]].Merge(false);
            //xSt2.Range[excel.Cells[8, 2], excel.Cells[8, 4]].NumberFormat = "@";
            //xSt2.Range[excel.Cells[8, 2], excel.Cells[8, 4]].Value2 = orderModel["userName"].ToString().Trim();

            //xSt2.Range[excel.Cells[9, 1], excel.Cells[9, 4]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            //xSt2.Range[excel.Cells[9, 1], excel.Cells[9, 4]].RowHeight = 20;
            //xSt2.Range[excel.Cells[9, 1], excel.Cells[9, 4]].Borders.Weight = XlBorderWeight.xlThin;
            //xSt2.Range[excel.Cells[9, 1], excel.Cells[9, 4]].Borders.LineStyle = XlLineStyle.xlContinuous;
            //xSt2.Range[excel.Cells[9, 1], excel.Cells[9, 4]].Interior.Color = System.Drawing.ColorTranslator.ToOle(c1);
            //excel.Cells[9, 1] = "产品货号";
            //excel.Cells[9, 2] = "产品描述";
            //excel.Cells[9, 3] = "数量";
            //excel.Cells[9, 4] = "单位";

            //rowIndex = 10;
            //for (int i = 0; i < Convert.ToInt32(orderModel["number"]); i++)
            //{
            //    xSt2.Range[excel.Cells[rowIndex, 1], excel.Cells[rowIndex, 4]].RowHeight = 20;
            //    xSt2.Range[excel.Cells[rowIndex, 1], excel.Cells[rowIndex, 4]].Borders.Weight = XlBorderWeight.xlThin;
            //    xSt2.Range[excel.Cells[rowIndex, 1], excel.Cells[rowIndex, 4]].Borders.LineStyle = XlLineStyle.xlContinuous;

            //    xSt2.Range[excel.Cells[rowIndex, 1], excel.Cells[rowIndex, 1]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
            //    xSt2.Range[excel.Cells[rowIndex, 2], excel.Cells[rowIndex, 2]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
            //    xSt2.Range[excel.Cells[rowIndex, 3], excel.Cells[rowIndex, 3]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            //    xSt2.Range[excel.Cells[rowIndex, 4], excel.Cells[rowIndex, 4]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

            //    xSt2.Range[excel.Cells[rowIndex, 1], excel.Cells[rowIndex, 4]].NumberFormat = "@";
            //    xSt2.Range[excel.Cells[rowIndex, 1], excel.Cells[rowIndex, 4]].WrapText = true;

            //    excel.Cells[rowIndex, 1] = (Convert.ToBoolean(productModel[i]["Sales_OrderProduct"]["isQuotation"].ToString()) ? productModel[i]["Sales_OrderProduct"]["quotationProductModel"].ToString().Trim() : productModel[i]["Product_Product"]["proNo"].ToString().Trim()) + " ";
            //    excel.Cells[rowIndex, 2] = (Convert.ToBoolean(productModel[i]["Sales_OrderProduct"]["isQuotation"].ToString()) ? productModel[i]["Sales_OrderProduct"]["quotationProductName"].ToString().Trim() : productModel[i]["Product_Brand"]["chineseName"].ToString().Trim() + (string.IsNullOrEmpty(productModel[i]["Product_Brand"]["englishName"].ToString().Trim()) ? "" : "/" + productModel[i]["Product_Brand"]["englishName"].ToString().Trim()) + "　" + productModel[i]["Product_Product"]["name"].ToString().Trim() + (string.IsNullOrEmpty(productModel[i]["Product_Product"]["ordNo"].ToString().Trim()) ? "" : "　" + productModel[i]["Product_Product"]["ordNo"].ToString().Trim()) + (string.IsNullOrEmpty(productModel[i]["Product_Product"]["package"].ToString().Trim()) ? "" : "　" + productModel[i]["Product_Product"]["package"].ToString().Trim())) + " ";
            //    excel.Cells[rowIndex, 3] = (string.IsNullOrEmpty(productModel[i]["Sales_OrderProduct"]["quantity"].ToString()) ? 0 : Convert.ToDecimal(productModel[i]["Sales_OrderProduct"]["quantity"]).ToString("N2")) + " ";
            //    excel.Cells[rowIndex, 4] = productModel[i]["Product_Product"]["unit"].ToString().Trim() + " ";
            //    rowIndex++;
            //}

            //for (int i = 0; i < 12 - Convert.ToInt32(orderModel["number"]); i++)
            //{
            //    xSt2.Range[excel.Cells[rowIndex, 1], excel.Cells[rowIndex, 4]].RowHeight = 20;
            //    xSt2.Range[excel.Cells[rowIndex, 1], excel.Cells[rowIndex, 4]].Borders.Weight = XlBorderWeight.xlThin;
            //    xSt2.Range[excel.Cells[rowIndex, 1], excel.Cells[rowIndex, 4]].Borders.LineStyle = XlLineStyle.xlContinuous;

            //    xSt2.Range[excel.Cells[rowIndex, 1], excel.Cells[rowIndex, 1]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
            //    xSt2.Range[excel.Cells[rowIndex, 2], excel.Cells[rowIndex, 2]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
            //    xSt2.Range[excel.Cells[rowIndex, 3], excel.Cells[rowIndex, 3]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            //    xSt2.Range[excel.Cells[rowIndex, 4], excel.Cells[rowIndex, 4]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

            //    xSt2.Range[excel.Cells[rowIndex, 1], excel.Cells[rowIndex, 4]].NumberFormat = "@";
            //    xSt2.Range[excel.Cells[rowIndex, 1], excel.Cells[rowIndex, 4]].WrapText = true;

            //    excel.Cells[rowIndex, 1] = " ";
            //    excel.Cells[rowIndex, 2] = " ";
            //    excel.Cells[rowIndex, 3] = " ";
            //    excel.Cells[rowIndex, 4] = " ";
            //    rowIndex++;
            //}

            //xSt2.Range[excel.Cells[rowIndex, 1], excel.Cells[rowIndex + 8, 4]].RowHeight = 20;

            //excel.Cells[rowIndex + 1, 1] = companyModel["keyname"] + "送货单号：";
            //xSt2.Range[excel.Cells[rowIndex + 1, 2], excel.Cells[rowIndex + 1, 4]].Merge(false);

            //xSt2.Range[excel.Cells[rowIndex + 2, 1], excel.Cells[rowIndex + 2, 4]].Merge(false);
            //xSt2.Range[excel.Cells[rowIndex + 2, 1], excel.Cells[rowIndex + 2, 4]].Value2 = "快递承运商（单号）： 自送";

            //xSt2.Range[excel.Cells[rowIndex + 4, 1], excel.Cells[rowIndex + 4, 4]].Merge(false);
            //xSt2.Range[excel.Cells[rowIndex + 4, 1], excel.Cells[rowIndex + 4, 4]].Value2 = "如果涉及其它未交付产品，我们会尽快安排。";

            //xSt2.Range[excel.Cells[rowIndex + 6, 1], excel.Cells[rowIndex + 6, 4]].Merge(false);
            //xSt2.Range[excel.Cells[rowIndex + 6, 1], excel.Cells[rowIndex + 6, 4]].Value2 = "非常感谢您的订购！";

            //xSt2.Range[excel.Cells[rowIndex + 7, 1], excel.Cells[rowIndex + 7, 4]].Merge(false);
            //xSt2.Range[excel.Cells[rowIndex + 7, 1], excel.Cells[rowIndex + 7, 4]].Value2 = companyModel["name"];

            //xSt2.Range[excel.Cells[rowIndex + 8, 1], excel.Cells[rowIndex + 8, 4]].Merge(false);
            //xSt2.Range[excel.Cells[rowIndex + 8, 1], excel.Cells[rowIndex + 8, 4]].Value2 = "http://www.mro9.com";
            //Range rn = xSt2.Range[excel.Cells[rowIndex + 8, 1], excel.Cells[rowIndex + 8, 4]];
            //rn.Hyperlinks.Add(rn, "http://www.mro9.com", "", Type.Missing, "http://www.mro9.com");
            //xSt2.Range[excel.Cells[rowIndex + 8, 1], excel.Cells[rowIndex + 8, 4]].Font.Name = "Arial Unicode MS";
            //xSt2.Range[excel.Cells[rowIndex + 8, 1], excel.Cells[rowIndex + 8, 4]].Font.Size = 9;
            //#endregion

            #region 送货单

            xSt1 = (_Worksheet)xBk.Worksheets.Add(Type.Missing, Type.Missing, 1, Type.Missing);

            xSt1.Name = "送货单";

            xSt1.PageSetup.LeftMargin = 1.8 / 0.035;
            xSt1.PageSetup.RightMargin = 1.8 / 0.035;
            xSt1.PageSetup.HeaderMargin = 0 / 0.035;
            xSt1.PageSetup.FooterMargin = 0 / 0.035;
            xSt1.PageSetup.TopMargin = 1.8 / 0.035;
            xSt1.PageSetup.BottomMargin = 2 / 0.035;

            excel.Cells.Font.Name = "微软雅黑";
            excel.Cells.Font.Size = 9;

            xSt1.Range[excel.Cells[1, 1], excel.Cells[1, 1]].ColumnWidth = 4;
            xSt1.Range[excel.Cells[1, 2], excel.Cells[1, 2]].ColumnWidth = 5;
            xSt1.Range[excel.Cells[1, 3], excel.Cells[1, 3]].ColumnWidth = 10;
            xSt1.Range[excel.Cells[1, 4], excel.Cells[1, 4]].ColumnWidth = 25;
            xSt1.Range[excel.Cells[1, 5], excel.Cells[1, 5]].ColumnWidth = 9;
            xSt1.Range[excel.Cells[1, 6], excel.Cells[1, 6]].ColumnWidth = 10;
            xSt1.Range[excel.Cells[1, 7], excel.Cells[1, 7]].ColumnWidth = 8;
            xSt1.Range[excel.Cells[1, 8], excel.Cells[1, 8]].ColumnWidth = 13;

            xSt1.Range[excel.Cells[1, 1], excel.Cells[1, 8]].RowHeight = 25;
            xSt1.Range[excel.Cells[1, 1], excel.Cells[1, 8]].Merge(false);
            xSt1.Range[excel.Cells[1, 1], excel.Cells[1, 8]].Value2 = "送货清单";
            xSt1.Shapes.AddPicture(Server.MapPath("~/").ToString().Trim() + "image\\" + companyModel["nid"].ToString().Trim() + ".gif", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoTrue, 380, 4, 150, 13);
            xSt1.Range[excel.Cells[1, 1], excel.Cells[1, 8]].Font.Bold = true;
            xSt1.Range[excel.Cells[1, 1], excel.Cells[1, 8]].Font.Size = 14;

            xSt1.Range[excel.Cells[1, 1], excel.Cells[1, 8]].Borders[XlBordersIndex.xlEdgeBottom].Weight = XlBorderWeight.xlThin;
            xSt1.Range[excel.Cells[1, 1], excel.Cells[1, 8]].Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;

            xSt1.Range[excel.Cells[2, 1], excel.Cells[2, 8]].RowHeight = 8;
            xSt1.Range[excel.Cells[3, 1], excel.Cells[9, 8]].RowHeight = 16;

            xSt1.Range[excel.Cells[3, 1], excel.Cells[3, 2]].Merge(false);
            xSt1.Range[excel.Cells[3, 1], excel.Cells[3, 2]].Value2 = "收货单位：";
            xSt1.Range[excel.Cells[3, 3], excel.Cells[3, 4]].Merge(false);
            xSt1.Range[excel.Cells[3, 3], excel.Cells[3, 4]].Value2 = customerModel["customerName"].ToString().Trim();
            excel.Cells[3, 5] = "供货单位：";
            xSt1.Range[excel.Cells[3, 6], excel.Cells[3, 8]].Merge(false);
            xSt1.Range[excel.Cells[3, 6], excel.Cells[3, 8]].Value2 = companyModel["name"].ToString().Trim();

            xSt1.Range[excel.Cells[4, 1], excel.Cells[4, 2]].Merge(false);
            xSt1.Range[excel.Cells[4, 1], excel.Cells[4, 2]].Value2 = "收货信息：";
            xSt1.Range[excel.Cells[4, 3], excel.Cells[4, 4]].Merge(false);
            xSt1.Range[excel.Cells[4, 3], excel.Cells[4, 4]].Value2 = orderModel["salesOrderFHXX"].ToString().Trim();
            excel.Cells[4, 5] = "联系人：";
            xSt1.Range[excel.Cells[4, 6], excel.Cells[4, 8]].Merge(false);
            xSt1.Range[excel.Cells[4, 6], excel.Cells[4, 8]].Value2 = userModel["name"].ToString().Trim();

            xSt1.Range[excel.Cells[5, 1], excel.Cells[5, 2]].Merge(false);
            xSt1.Range[excel.Cells[5, 3], excel.Cells[5, 4]].Merge(false);
            excel.Cells[5, 5] = "手机：";
            xSt1.Range[excel.Cells[5, 6], excel.Cells[5, 8]].Merge(false);
            xSt1.Range[excel.Cells[5, 6], excel.Cells[5, 8]].NumberFormat = "@";
            xSt1.Range[excel.Cells[5, 6], excel.Cells[5, 8]].Value2 = userModel["mobilePhone"].ToString().Trim() + " ";

            xSt1.Range[excel.Cells[6, 1], excel.Cells[6, 2]].Merge(false);
            xSt1.Range[excel.Cells[6, 1], excel.Cells[6, 2]].Value2 = "客户订单号：";
            xSt1.Range[excel.Cells[6, 3], excel.Cells[6, 4]].Merge(false);
            xSt1.Range[excel.Cells[6, 3], excel.Cells[6, 4]].NumberFormat = "@";
            xSt1.Range[excel.Cells[6, 3], excel.Cells[6, 4]].Value2 = orderModel["customerOrderNo"].ToString().Trim();
            excel.Cells[6, 5] = "服务专线：";
            xSt1.Range[excel.Cells[6, 6], excel.Cells[6, 8]].Merge(false);
            xSt1.Range[excel.Cells[6, 6], excel.Cells[6, 8]].Value2 = "400-816-1658 转 " + userModel["extensionNum"].ToString().Trim();

            xSt1.Range[excel.Cells[7, 1], excel.Cells[7, 2]].Merge(false);
            xSt1.Range[excel.Cells[7, 1], excel.Cells[7, 2]].Value2 = "申请人：";
            xSt1.Range[excel.Cells[7, 3], excel.Cells[7, 4]].Merge(false);
            xSt1.Range[excel.Cells[7, 3], excel.Cells[7, 4]].Value2 = orderModel["userName"].ToString().Trim();
            xSt1.Range[excel.Cells[7, 6], excel.Cells[7, 8]].Merge(false);

            xSt1.Range[excel.Cells[8, 1], excel.Cells[8, 2]].Merge(false);
            xSt1.Range[excel.Cells[8, 3], excel.Cells[8, 4]].Merge(false);
            excel.Cells[8, 5] = "系统单号：";
            xSt1.Range[excel.Cells[8, 6], excel.Cells[8, 8]].Merge(false);
            xSt1.Range[excel.Cells[8, 6], excel.Cells[8, 8]].Value2 = orderModel["orderNo"].ToString().Trim();

            xSt1.Range[excel.Cells[9, 1], excel.Cells[9, 2]].Merge(false);
            xSt1.Range[excel.Cells[9, 1], excel.Cells[9, 2]].Value2 = "送货日期：";
            xSt1.Range[excel.Cells[9, 3], excel.Cells[9, 4]].Merge(false);
            xSt1.Range[excel.Cells[9, 3], excel.Cells[9, 4]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
            xSt1.Range[excel.Cells[9, 3], excel.Cells[9, 4]].Value2 = "";
            excel.Cells[9, 5] = "页码：";
            xSt1.Range[excel.Cells[9, 6], excel.Cells[9, 8]].Merge(false);
            xSt1.Range[excel.Cells[9, 6], excel.Cells[9, 8]].NumberFormat = "@";
            xSt1.Range[excel.Cells[9, 6], excel.Cells[9, 8]].Value2 = "1 / 1";

            xSt1.Range[excel.Cells[10, 1], excel.Cells[10, 8]].RowHeight = 23;

            //excel.Ce
            xSt1.Range[excel.Cells[11, 1], excel.Cells[11, 8]].Borders[XlBordersIndex.xlEdgeBottom].Weight = XlBorderWeight.xlMedium;
            xSt1.Range[excel.Cells[11, 1], excel.Cells[11, 8]].Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
            excel.Cells[11, 1] = "序号";
            xSt1.Range[excel.Cells[11, 2], excel.Cells[11, 3]].Merge(false);
            xSt1.Range[excel.Cells[11, 2], excel.Cells[11, 3]].Value2 = "产品货号";
            xSt1.Range[excel.Cells[11, 4], excel.Cells[11, 5]].Merge(false);
            xSt1.Range[excel.Cells[11, 4], excel.Cells[11, 5]].Value2 = "产品描述";
            excel.Cells[11, 6] = "送货数量";
            excel.Cells[11, 7] = "单位";
            excel.Cells[11, 8] = "备注";

            rowIndex = 12;
            for (int i = 0; i < Convert.ToInt32(orderModel["number"]); i++)
            {
                xSt1.Range[excel.Cells[rowIndex, 1], excel.Cells[rowIndex, 8]].Borders[XlBordersIndex.xlEdgeBottom].Weight = XlBorderWeight.xlHairline;
                xSt1.Range[excel.Cells[rowIndex, 1], excel.Cells[rowIndex, 8]].Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;

                xSt1.Range[excel.Cells[rowIndex, 1], excel.Cells[rowIndex, 8]].RowHeight = 36;
                xSt1.Range[excel.Cells[rowIndex, 1], excel.Cells[rowIndex, 1]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                xSt1.Range[excel.Cells[rowIndex, 2], excel.Cells[rowIndex, 3]].Merge(false);
                xSt1.Range[excel.Cells[rowIndex, 2], excel.Cells[rowIndex, 3]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
                xSt1.Range[excel.Cells[rowIndex, 4], excel.Cells[rowIndex, 5]].Merge(false);
                xSt1.Range[excel.Cells[rowIndex, 4], excel.Cells[rowIndex, 5]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
                xSt1.Range[excel.Cells[rowIndex, 6], excel.Cells[rowIndex, 6]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                xSt1.Range[excel.Cells[rowIndex, 7], excel.Cells[rowIndex, 7]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                xSt1.Range[excel.Cells[rowIndex, 8], excel.Cells[rowIndex, 8]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;

                xSt1.Range[excel.Cells[rowIndex, 2], excel.Cells[rowIndex, 8]].NumberFormat = "@";
                xSt1.Range[excel.Cells[rowIndex, 2], excel.Cells[rowIndex, 8]].WrapText = true;

                excel.Cells[rowIndex, 1] = Convert.ToString(i + 1);
                xSt1.Range[excel.Cells[rowIndex, 2], excel.Cells[rowIndex, 3]].Value2 = (Convert.ToBoolean(productModel[i]["Sales_OrderProduct"]["isQuotation"].ToString()) ? productModel[i]["Sales_OrderProduct"]["quotationProductModel"].ToString().Trim() : productModel[i]["Product_Product"]["proNo"].ToString().Trim()) + " ";
                xSt1.Range[excel.Cells[rowIndex, 4], excel.Cells[rowIndex, 5]].Value2 = (Convert.ToBoolean(productModel[i]["Sales_OrderProduct"]["isQuotation"].ToString()) ? productModel[i]["Sales_OrderProduct"]["quotationProductName"].ToString().Trim() : productModel[i]["Product_Brand"]["chineseName"].ToString().Trim() + (string.IsNullOrEmpty(productModel[i]["Product_Brand"]["englishName"].ToString().Trim()) ? "" : "/" + productModel[i]["Product_Brand"]["englishName"].ToString().Trim()) + "　" + productModel[i]["Product_Product"]["name"].ToString().Trim() + (string.IsNullOrEmpty(productModel[i]["Product_Product"]["ordNo"].ToString().Trim()) ? "" : "　" + productModel[i]["Product_Product"]["ordNo"].ToString().Trim()) + (string.IsNullOrEmpty(productModel[i]["Product_Product"]["package"].ToString().Trim()) ? "" : "　" + productModel[i]["Product_Product"]["package"].ToString().Trim())) + " ";
                excel.Cells[rowIndex, 6] = (string.IsNullOrEmpty(productModel[i]["Sales_OrderProduct"]["quantity"].ToString()) ? 0 : Convert.ToDecimal(productModel[i]["Sales_OrderProduct"]["quantity"]).ToString("N2")) + " ";
                excel.Cells[rowIndex, 7] = productModel[i]["Product_Product"]["unit"].ToString().Trim() + " ";
                try
                {
                    excel.Cells[rowIndex, 8] = productModel[i]["Sales_OrderProduct"]["remark"].ToString().Trim() + " ";
                }
                catch { }
                
                rowIndex++;
            }

            for (int i = 0; i < 12 - Convert.ToInt32(orderModel["number"]); i++)
            {
                xSt1.Range[excel.Cells[rowIndex, 1], excel.Cells[rowIndex, 8]].Borders[XlBordersIndex.xlEdgeBottom].Weight = XlBorderWeight.xlHairline;
                xSt1.Range[excel.Cells[rowIndex, 1], excel.Cells[rowIndex, 8]].Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;

                xSt1.Range[excel.Cells[rowIndex, 1], excel.Cells[rowIndex, 8]].RowHeight = 36;
                xSt1.Range[excel.Cells[rowIndex, 1], excel.Cells[rowIndex, 1]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                xSt1.Range[excel.Cells[rowIndex, 2], excel.Cells[rowIndex, 3]].Merge(false);
                xSt1.Range[excel.Cells[rowIndex, 2], excel.Cells[rowIndex, 3]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
                xSt1.Range[excel.Cells[rowIndex, 4], excel.Cells[rowIndex, 5]].Merge(false);
                xSt1.Range[excel.Cells[rowIndex, 4], excel.Cells[rowIndex, 5]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
                xSt1.Range[excel.Cells[rowIndex, 6], excel.Cells[rowIndex, 6]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                xSt1.Range[excel.Cells[rowIndex, 7], excel.Cells[rowIndex, 7]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                xSt1.Range[excel.Cells[rowIndex, 8], excel.Cells[rowIndex, 8]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;

                xSt1.Range[excel.Cells[rowIndex, 2], excel.Cells[rowIndex, 8]].NumberFormat = "@";
                xSt1.Range[excel.Cells[rowIndex, 2], excel.Cells[rowIndex, 8]].WrapText = true;

                excel.Cells[rowIndex, 1] = Convert.ToString(i + 1 + Convert.ToInt32(orderModel["number"]));
                xSt1.Range[excel.Cells[rowIndex, 2], excel.Cells[rowIndex, 3]].Value2 = " ";
                xSt1.Range[excel.Cells[rowIndex, 4], excel.Cells[rowIndex, 5]].Value2 = " ";
                excel.Cells[rowIndex, 6] = " ";
                excel.Cells[rowIndex, 7] = " ";
                excel.Cells[rowIndex, 8] = " ";
                rowIndex++;
            }

            xSt1.Range[excel.Cells[rowIndex, 1], excel.Cells[rowIndex, 8]].RowHeight = 30;
            xSt1.Range[excel.Cells[rowIndex, 1], excel.Cells[rowIndex, 8]].Borders[XlBordersIndex.xlEdgeTop].Weight = XlBorderWeight.xlMedium;
            xSt1.Range[excel.Cells[rowIndex, 1], excel.Cells[rowIndex, 8]].Borders[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;

            xSt1.Range[excel.Cells[rowIndex + 1, 1], excel.Cells[rowIndex + 1, 8]].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignBottom;
            xSt1.Range[excel.Cells[rowIndex + 1, 1], excel.Cells[rowIndex + 1, 8]].RowHeight = 28;
            xSt1.Range[excel.Cells[rowIndex + 1, 1], excel.Cells[rowIndex + 1, 4]].Merge(false);
            xSt1.Range[excel.Cells[rowIndex + 1, 1], excel.Cells[rowIndex + 1, 4]].Value2 = "非常感谢您的订购、验货及签收！";
            xSt1.Range[excel.Cells[rowIndex + 1, 6], excel.Cells[rowIndex + 1, 6]].Font.Bold = true;
            xSt1.Range[excel.Cells[rowIndex + 1, 6], excel.Cells[rowIndex + 1, 6]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
            excel.Cells[rowIndex + 1, 6] = "收货方签收：";
            xSt1.Range[excel.Cells[rowIndex + 1, 7], excel.Cells[rowIndex + 1, 8]].Merge(false);
            xSt1.Range[excel.Cells[rowIndex + 1, 7], excel.Cells[rowIndex + 1, 8]].Borders[XlBordersIndex.xlEdgeBottom].Weight = XlBorderWeight.xlHairline;
            xSt1.Range[excel.Cells[rowIndex + 1, 7], excel.Cells[rowIndex + 1, 8]].Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;

            xSt1.Range[excel.Cells[rowIndex + 2, 1], excel.Cells[rowIndex + 2, 8]].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignBottom;
            xSt1.Range[excel.Cells[rowIndex + 2, 1], excel.Cells[rowIndex + 2, 8]].RowHeight = 28;
            xSt1.Range[excel.Cells[rowIndex + 2, 1], excel.Cells[rowIndex + 2, 4]].Merge(false);
            xSt1.Range[excel.Cells[rowIndex + 2, 1], excel.Cells[rowIndex + 2, 4]].Value2 = companyModel["name"] + "（www.mro9.com）";
            xSt1.Range[excel.Cells[rowIndex + 2, 6], excel.Cells[rowIndex + 2, 6]].Font.Bold = true;
            xSt1.Range[excel.Cells[rowIndex + 2, 6], excel.Cells[rowIndex + 2, 6]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
            excel.Cells[rowIndex + 2, 6] = "日期：";
            xSt1.Range[excel.Cells[rowIndex + 2, 7], excel.Cells[rowIndex + 2, 8]].Merge(false);
            xSt1.Range[excel.Cells[rowIndex + 2, 7], excel.Cells[rowIndex + 2, 8]].Borders[XlBordersIndex.xlEdgeBottom].Weight = XlBorderWeight.xlHairline;
            xSt1.Range[excel.Cells[rowIndex + 2, 7], excel.Cells[rowIndex + 2, 8]].Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;


            #endregion

            //#region 流程单

            //xSt = (_Worksheet)xBk.Worksheets.Add(xSt1, Type.Missing, 1, Type.Missing);

            //xSt.Name = "流程单";

            //xSt.PageSetup.Orientation = XlPageOrientation.xlLandscape;
            //xSt.PageSetup.LeftMargin = 1.5 / 0.035;
            //xSt.PageSetup.RightMargin = 1 / 0.035;
            //xSt.PageSetup.HeaderMargin = 0 / 0.035;
            //xSt.PageSetup.FooterMargin = 0.8 / 0.035;
            //xSt.PageSetup.TopMargin = 0.8 / 0.035;
            //xSt.PageSetup.BottomMargin = 0.8 / 0.035;

            //excel.Cells.Font.Name = "微软雅黑";
            //excel.Cells.Font.Size = 8;

            //xSt.Range[excel.Cells[1, 1], excel.Cells[1, 1]].ColumnWidth = 2;
            //xSt.Range[excel.Cells[1, 2], excel.Cells[1, 2]].ColumnWidth = 10;
            //xSt.Range[excel.Cells[1, 3], excel.Cells[1, 3]].ColumnWidth = 18;
            //xSt.Range[excel.Cells[1, 4], excel.Cells[1, 4]].ColumnWidth = 4;
            //xSt.Range[excel.Cells[1, 5], excel.Cells[1, 5]].ColumnWidth = 3;
            //xSt.Range[excel.Cells[1, 6], excel.Cells[1, 6]].ColumnWidth = 7;
            //xSt.Range[excel.Cells[1, 7], excel.Cells[1, 7]].ColumnWidth = 8;
            //xSt.Range[excel.Cells[1, 8], excel.Cells[1, 8]].ColumnWidth = 10;
            //xSt.Range[excel.Cells[1, 9], excel.Cells[1, 9]].ColumnWidth = 17;
            //xSt.Range[excel.Cells[1, 10], excel.Cells[1, 10]].ColumnWidth = 5;
            //xSt.Range[excel.Cells[1, 11], excel.Cells[1, 11]].ColumnWidth = 4;
            //xSt.Range[excel.Cells[1, 12], excel.Cells[1, 12]].ColumnWidth = 7;
            //xSt.Range[excel.Cells[1, 13], excel.Cells[1, 13]].ColumnWidth = 7;
            //xSt.Range[excel.Cells[1, 14], excel.Cells[1, 14]].ColumnWidth = 7;
            //xSt.Range[excel.Cells[1, 15], excel.Cells[1, 15]].ColumnWidth = 7;
            //xSt.Range[excel.Cells[1, 16], excel.Cells[1, 16]].ColumnWidth = 7;

            //xSt.Range[excel.Cells[1, 1], excel.Cells[1, 16]].RowHeight = 25;
            //xSt.Range[excel.Cells[1, 1], excel.Cells[1, 9]].Merge(false);
            //xSt.Range[excel.Cells[1, 1], excel.Cells[1, 9]].Value2 = "销售订单：" + orderModel["orderNo"].ToString().Trim();
            //xSt.Shapes.AddPicture(Server.MapPath("~/").ToString().Trim() + "image\\" + companyModel["nid"].ToString().Trim() + ".gif", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoTrue, 640, 4, 150, 13);
            //xSt.Range[excel.Cells[1, 1], excel.Cells[1, 3]].Font.Bold = true;
            //xSt.Range[excel.Cells[1, 1], excel.Cells[1, 3]].Font.Size = 14;

            //xSt.Range[excel.Cells[2, 1], excel.Cells[4, 16]].RowHeight = 24;
            //xSt.Range[excel.Cells[2, 1], excel.Cells[4, 16]].Font.Bold = true;
            //xSt.Range[excel.Cells[2, 1], excel.Cells[4, 16]].Font.Size = 10;

            //xSt.Range[excel.Cells[2, 1], excel.Cells[2, 2]].Merge(false);
            //xSt.Range[excel.Cells[2, 1], excel.Cells[2, 2]].Value2 = "客户名称：";
            //xSt.Range[excel.Cells[2, 3], excel.Cells[2, 4]].Merge(false);
            //xSt.Range[excel.Cells[2, 3], excel.Cells[2, 4]].Value2 = customerModel["customerKeyName"].ToString().Trim();
            //xSt.Range[excel.Cells[2, 5], excel.Cells[2, 6]].Merge(false);
            //xSt.Range[excel.Cells[2, 5], excel.Cells[2, 6]].Value2 = "客户订单号：";
            //xSt.Range[excel.Cells[2, 7], excel.Cells[2, 8]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
            //xSt.Range[excel.Cells[2, 7], excel.Cells[2, 8]].Merge(false);
            //xSt.Range[excel.Cells[2, 7], excel.Cells[2, 8]].NumberFormat = "@";
            //xSt.Range[excel.Cells[2, 7], excel.Cells[2, 8]].Value2 = orderModel["customerOrderNo"].ToString().Trim();
            //xSt.Range[excel.Cells[2, 9], excel.Cells[2, 11]].Merge(false);
            //xSt.Range[excel.Cells[2, 9], excel.Cells[2, 11]].Value2 = "出货仓库：" + orderModel["storehouseName"].ToString().Trim();
            //xSt.Range[excel.Cells[2, 12], excel.Cells[2, 16]].Merge(false);
            //xSt.Range[excel.Cells[2, 12], excel.Cells[2, 16]].Value2 = "订单日期：" + Convert.ToDateTime(orderModel["recordDate"].ToString()).ToString("yyyy年MM月dd日");

            //xSt.Range[excel.Cells[2, 1], excel.Cells[2, 16]].Borders.Weight = XlBorderWeight.xlHairline;
            //xSt.Range[excel.Cells[2, 1], excel.Cells[2, 16]].Borders.LineStyle = 1;

            //xSt.Range[excel.Cells[3, 1], excel.Cells[3, 2]].Merge(false);
            //xSt.Range[excel.Cells[3, 1], excel.Cells[3, 2]].Value2 = "Buyer：";
            //xSt.Range[excel.Cells[3, 3], excel.Cells[3, 4]].Merge(false);
            //xSt.Range[excel.Cells[3, 3], excel.Cells[3, 4]].Value2 = orderModel["purchaserName"].ToString().Trim();
            //xSt.Range[excel.Cells[3, 5], excel.Cells[3, 6]].Merge(false);
            //xSt.Range[excel.Cells[3, 5], excel.Cells[3, 6]].Value2 = "User：";
            //xSt.Range[excel.Cells[3, 7], excel.Cells[3, 8]].Merge(false);
            //xSt.Range[excel.Cells[3, 7], excel.Cells[3, 8]].Value2 = orderModel["userName"].ToString().Trim();
            //xSt.Range[excel.Cells[3, 9], excel.Cells[3, 11]].Merge(false);
            //xSt.Range[excel.Cells[3, 9], excel.Cells[3, 11]].Value2 = "销售：" + orderModel["personInChargeName"].ToString().Trim();
            //xSt.Range[excel.Cells[3, 12], excel.Cells[3, 16]].Merge(false);
            //xSt.Range[excel.Cells[3, 12], excel.Cells[3, 16]].Value2 = "交货时间：" + Convert.ToDateTime(orderModel["deliveryDate"]).ToString("yyyy年MM月dd日");

            //xSt.Range[excel.Cells[3, 1], excel.Cells[3, 11]].Borders.Weight = XlBorderWeight.xlHairline;
            //xSt.Range[excel.Cells[3, 12], excel.Cells[3, 16]].Borders.LineStyle = XlLineStyle.xlContinuous;
            //xSt.Range[excel.Cells[3, 12], excel.Cells[3, 16]].Borders.Weight = XlBorderWeight.xlThick;
            //xSt.Range[excel.Cells[3, 12], excel.Cells[3, 16]].Borders.LineStyle = XlLineStyle.xlSlantDashDot;

            //xSt.Range[excel.Cells[4, 1], excel.Cells[4, 2]].Merge(false);
            //xSt.Range[excel.Cells[4, 1], excel.Cells[4, 2]].Value2 = "收货信息：";
            //xSt.Range[excel.Cells[4, 3], excel.Cells[4, 8]].Merge(false);
            //xSt.Range[excel.Cells[4, 3], excel.Cells[4, 8]].Value2 = orderModel["salesOrderFHXX"].ToString().Trim();
            //xSt.Range[excel.Cells[4, 9], excel.Cells[4, 16]].Merge(false);
            //xSt.Range[excel.Cells[4, 9], excel.Cells[4, 16]].Value2 = "备注：" + orderModel["remark"].ToString().Trim(); ;

            //xSt.Range[excel.Cells[4, 1], excel.Cells[4, 16]].Borders.Weight = XlBorderWeight.xlHairline;
            //xSt.Range[excel.Cells[4, 1], excel.Cells[4, 16]].Borders[XlBordersIndex.xlEdgeBottom].Weight = XlBorderWeight.xlThin;
            //xSt.Range[excel.Cells[4, 1], excel.Cells[4, 16]].Borders.LineStyle = 1;

            //xSt.Range[excel.Cells[4, 12], excel.Cells[4, 16]].Borders[XlBordersIndex.xlEdgeTop].Weight = XlBorderWeight.xlThick;
            //xSt.Range[excel.Cells[4, 12], excel.Cells[4, 16]].Borders[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlSlantDashDot;

            //xSt.Range[excel.Cells[5, 1], excel.Cells[6, 16]].RowHeight = 16;
            //xSt.Range[excel.Cells[5, 1], excel.Cells[6, 16]].Font.Bold = true;
            //xSt.Range[excel.Cells[5, 1], excel.Cells[6, 16]].Font.Size = 8;
            //xSt.Range[excel.Cells[5, 1], excel.Cells[6, 16]].HorizontalAlignment = XlHAlign.xlHAlignCenter;//设置标题格式为居中对齐 
            //xSt.Range[excel.Cells[5, 2], excel.Cells[5, 7]].Merge(false);
            //xSt.Range[excel.Cells[5, 2], excel.Cells[5, 7]].Value2 = "【客户】订单项";
            //xSt.Range[excel.Cells[5, 8], excel.Cells[5, 13]].Merge(false);
            //xSt.Range[excel.Cells[5, 8], excel.Cells[5, 13]].Value2 = "【内部】订货项";
            //xSt.Range[excel.Cells[5, 1], excel.Cells[6, 1]].Merge(false);
            //xSt.Range[excel.Cells[5, 1], excel.Cells[6, 1]].Value2 = "NO";

            //xSt.Range[excel.Cells[5, 1], excel.Cells[5, 16]].Borders.Weight = XlBorderWeight.xlHairline;
            //xSt.Range[excel.Cells[5, 1], excel.Cells[5, 16]].Borders[XlBordersIndex.xlEdgeTop].Weight = XlBorderWeight.xlThin;
            //xSt.Range[excel.Cells[5, 1], excel.Cells[5, 16]].Borders.LineStyle = 1;

            //excel.Cells[6, 2] = "订单型号";
            //excel.Cells[6, 3] = "订单名称";
            //excel.Cells[6, 4] = "数量";
            //excel.Cells[6, 5] = "单位";
            //excel.Cells[6, 6] = "含税单价";
            //excel.Cells[6, 7] = "销售总价";
            //excel.Cells[6, 8] = "产品型号";
            //excel.Cells[6, 9] = "产品名称";
            //excel.Cells[6, 10] = "数量";
            //excel.Cells[6, 11] = "单位";
            //excel.Cells[6, 12] = "含税进价";
            //excel.Cells[6, 13] = "进货总价";
            //xSt.Range[excel.Cells[5, 14], excel.Cells[6, 14]].Merge(false);
            //xSt.Range[excel.Cells[5, 14], excel.Cells[6, 14]].Value2 = "订货确认";
            //xSt.Range[excel.Cells[5, 14], excel.Cells[6, 14]].Borders.Weight = XlBorderWeight.xlHairline;
            //xSt.Range[excel.Cells[5, 14], excel.Cells[6, 14]].Borders.LineStyle = 1;
            //xSt.Range[excel.Cells[5, 15], excel.Cells[6, 15]].Merge(false);
            //xSt.Range[excel.Cells[5, 15], excel.Cells[6, 15]].Value2 = "到货确认";
            //xSt.Range[excel.Cells[5, 15], excel.Cells[6, 15]].Borders.Weight = XlBorderWeight.xlHairline;
            //xSt.Range[excel.Cells[5, 15], excel.Cells[6, 15]].Borders.LineStyle = 1;
            //xSt.Range[excel.Cells[5, 16], excel.Cells[6, 16]].Merge(false);
            //xSt.Range[excel.Cells[5, 16], excel.Cells[6, 16]].Value2 = "备注";
            //xSt.Range[excel.Cells[5, 16], excel.Cells[6, 16]].Borders.Weight = XlBorderWeight.xlHairline;
            //xSt.Range[excel.Cells[5, 16], excel.Cells[6, 16]].Borders.LineStyle = 1;

            //xSt.Range[excel.Cells[6, 1], excel.Cells[6, 16]].Borders.Weight = XlBorderWeight.xlHairline;
            //xSt.Range[excel.Cells[6, 1], excel.Cells[6, 16]].Borders.LineStyle = 1;

            //rowIndex = 7;
            //for (int i = 0; i < Convert.ToInt32(orderModel["number"]); i++)
            //{
            //    xSt.Range[excel.Cells[rowIndex, 1], excel.Cells[rowIndex, 16]].Borders.Weight = XlBorderWeight.xlHairline;
            //    xSt.Range[excel.Cells[rowIndex, 1], excel.Cells[rowIndex, 16]].Borders.LineStyle = 1;

            //    xSt.Range[excel.Cells[rowIndex, 1], excel.Cells[rowIndex, 16]].RowHeight = 30;
            //    xSt.Range[excel.Cells[rowIndex, 1], excel.Cells[rowIndex, 16]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            //    xSt.Range[excel.Cells[rowIndex, 2], excel.Cells[rowIndex, 3]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
            //    xSt.Range[excel.Cells[rowIndex, 6], excel.Cells[rowIndex, 7]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
            //    xSt.Range[excel.Cells[rowIndex, 8], excel.Cells[rowIndex, 9]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
            //    xSt.Range[excel.Cells[rowIndex, 12], excel.Cells[rowIndex, 13]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
            //    xSt.Range[excel.Cells[rowIndex, 2], excel.Cells[rowIndex, 16]].NumberFormat = "@";
            //    xSt.Range[excel.Cells[rowIndex, 2], excel.Cells[rowIndex, 16]].WrapText = true;
            //    excel.Cells[rowIndex, 1] = Convert.ToString(i + 1);
            //    excel.Cells[rowIndex, 2] = (Convert.ToBoolean(productModel[i]["Sales_OrderProduct"]["isQuotation"].ToString()) ? productModel[i]["Sales_OrderProduct"]["quotationProductModel"].ToString().Trim() : productModel[i]["Product_Product"]["proNo"].ToString().Trim()) + " ";
            //    excel.Cells[rowIndex, 3] = (Convert.ToBoolean(productModel[i]["Sales_OrderProduct"]["isQuotation"].ToString()) ? productModel[i]["Sales_OrderProduct"]["quotationProductName"].ToString().Trim() : productModel[i]["Product_Brand"]["chineseName"].ToString().Trim() + (string.IsNullOrEmpty(productModel[i]["Product_Brand"]["englishName"].ToString().Trim()) ? "" : "/" + productModel[i]["Product_Brand"]["englishName"].ToString().Trim()) + "　" + productModel[i]["Product_Product"]["name"].ToString().Trim() + (string.IsNullOrEmpty(productModel[i]["Product_Product"]["ordNo"].ToString().Trim()) ? "" : "　" + productModel[i]["Product_Product"]["ordNo"].ToString().Trim()) + (string.IsNullOrEmpty(productModel[i]["Product_Product"]["package"].ToString().Trim()) ? "" : "　" + productModel[i]["Product_Product"]["package"].ToString().Trim())) + " ";
            //    excel.Cells[rowIndex, 4] = (string.IsNullOrEmpty(productModel[i]["Sales_OrderProduct"]["quantity"].ToString()) ? 0 : Convert.ToDecimal(productModel[i]["Sales_OrderProduct"]["quantity"]).ToString("N2")) + " ";
            //    excel.Cells[rowIndex, 5] = productModel[i]["Product_Product"]["unit"].ToString().Trim() + " ";
            //    excel.Cells[rowIndex, 6] = (string.IsNullOrEmpty(productModel[i]["Sales_OrderProduct"]["unitPrice"].ToString()) ? "0" : Convert.ToDecimal(productModel[i]["Sales_OrderProduct"]["unitPrice"].ToString()).ToString("N2"));
            //    excel.Cells[rowIndex, 7] = (string.IsNullOrEmpty(productModel[i]["Sales_OrderProduct"]["totalPrice"].ToString()) ? "0" : Convert.ToDecimal(productModel[i]["Sales_OrderProduct"]["totalPrice"].ToString()).ToString("N2"));
            //    excel.Cells[rowIndex, 8] = productModel[i]["Product_Product"]["proNo"].ToString().Trim() + " ";
            //    excel.Cells[rowIndex, 9] = productModel[i]["Product_Brand"]["chineseName"].ToString().Trim() + (string.IsNullOrEmpty(productModel[i]["Product_Brand"]["englishName"].ToString().Trim()) ? "" : "/" + productModel[i]["Product_Brand"]["englishName"].ToString().Trim()) + "　" + productModel[i]["Product_Product"]["name"].ToString().Trim() + (string.IsNullOrEmpty(productModel[i]["Product_Product"]["ordNo"].ToString().Trim()) ? "" : "　" + productModel[i]["Product_Product"]["ordNo"].ToString().Trim()) + (string.IsNullOrEmpty(productModel[i]["Product_Product"]["package"].ToString().Trim()) ? "" : "　" + productModel[i]["Product_Product"]["package"].ToString().Trim()) + " ";
            //    excel.Cells[rowIndex, 10] = " ";
            //    excel.Cells[rowIndex, 11] = productModel[i]["Product_Product"]["unit"].ToString().Trim() + " ";
            //    excel.Cells[rowIndex, 12] = " ";
            //    excel.Cells[rowIndex, 13] = " ";
            //    excel.Cells[rowIndex, 14] = " ";
            //    excel.Cells[rowIndex, 15] = " ";
            //    excel.Cells[rowIndex, 16] = " ";
            //    rowIndex++;
            //}

            //for (int i = 0; i < 12 - Convert.ToInt32(orderModel["number"]); i++)
            //{
            //    xSt.Range[excel.Cells[rowIndex, 1], excel.Cells[rowIndex, 16]].Borders.Weight = XlBorderWeight.xlHairline;
            //    xSt.Range[excel.Cells[rowIndex, 1], excel.Cells[rowIndex, 16]].Borders.LineStyle = 1;

            //    xSt.Range[excel.Cells[rowIndex, 1], excel.Cells[rowIndex, 16]].RowHeight = 30;
            //    xSt.Range[excel.Cells[rowIndex, 1], excel.Cells[rowIndex, 16]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            //    xSt.Range[excel.Cells[rowIndex, 2], excel.Cells[rowIndex, 3]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
            //    xSt.Range[excel.Cells[rowIndex, 6], excel.Cells[rowIndex, 7]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
            //    xSt.Range[excel.Cells[rowIndex, 8], excel.Cells[rowIndex, 9]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
            //    xSt.Range[excel.Cells[rowIndex, 12], excel.Cells[rowIndex, 13]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
            //    xSt.Range[excel.Cells[rowIndex, 2], excel.Cells[rowIndex, 16]].NumberFormat = "@";
            //    xSt.Range[excel.Cells[rowIndex, 2], excel.Cells[rowIndex, 16]].WrapText = true;
            //    excel.Cells[rowIndex, 1] = Convert.ToString(i + 1 + Convert.ToInt32(orderModel["number"]));
            //    excel.Cells[rowIndex, 2] = " ";
            //    excel.Cells[rowIndex, 3] = " ";
            //    excel.Cells[rowIndex, 4] = " ";
            //    excel.Cells[rowIndex, 5] = " ";
            //    excel.Cells[rowIndex, 6] = " ";
            //    excel.Cells[rowIndex, 7] = " ";
            //    excel.Cells[rowIndex, 8] = " ";
            //    excel.Cells[rowIndex, 9] = " ";
            //    excel.Cells[rowIndex, 10] = " ";
            //    excel.Cells[rowIndex, 11] = " ";
            //    excel.Cells[rowIndex, 12] = " ";
            //    excel.Cells[rowIndex, 13] = " ";
            //    excel.Cells[rowIndex, 14] = " ";
            //    excel.Cells[rowIndex, 15] = " ";
            //    excel.Cells[rowIndex, 16] = " ";
            //    rowIndex++;
            //}

            //xSt.Range[excel.Cells[rowIndex, 1], excel.Cells[rowIndex, 16]].Borders.Weight = XlBorderWeight.xlHairline;
            //xSt.Range[excel.Cells[rowIndex, 1], excel.Cells[rowIndex, 16]].Borders[XlBordersIndex.xlEdgeBottom].Weight = XlBorderWeight.xlThin;
            //xSt.Range[excel.Cells[rowIndex, 1], excel.Cells[rowIndex, 16]].Borders.LineStyle = 1;

            //xSt.Range[excel.Cells[rowIndex, 1], excel.Cells[rowIndex, 16]].RowHeight = 22;
            //xSt.Range[excel.Cells[rowIndex, 1], excel.Cells[rowIndex, 6]].Merge(false);
            //xSt.Range[excel.Cells[rowIndex, 1], excel.Cells[rowIndex, 6]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
            //xSt.Range[excel.Cells[rowIndex, 1], excel.Cells[rowIndex, 6]].Font.Size = 8;
            //xSt.Range[excel.Cells[rowIndex, 1], excel.Cells[rowIndex, 6]].Value2 = "销售总额";
            //excel.Cells[rowIndex, 7] = (string.IsNullOrEmpty(orderModel["totalPrice"].ToString()) ? "0" : Convert.ToDecimal(orderModel["totalPrice"].ToString()).ToString("N2"));

            //xSt.Range[excel.Cells[rowIndex, 8], excel.Cells[rowIndex, 12]].Merge(false);
            //xSt.Range[excel.Cells[rowIndex, 8], excel.Cells[rowIndex, 12]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
            //xSt.Range[excel.Cells[rowIndex, 8], excel.Cells[rowIndex, 12]].Font.Size = 8;
            //xSt.Range[excel.Cells[rowIndex, 8], excel.Cells[rowIndex, 12]].Value2 = "进货总额";
            ////if (this.IsDecimal(Request.QueryString["PurchaseTotalPrices"].ToString()))
            ////{
            ////    excel.Cells[rowIndex, 13] = Convert.ToDecimal(Request.QueryString["PurchaseTotalPrices"].ToString()).ToString("N2") + " ";
            ////}

            //xSt.Range[excel.Cells[rowIndex, 14], excel.Cells[rowIndex, 15]].Merge(false);
            //xSt.Range[excel.Cells[rowIndex, 14], excel.Cells[rowIndex, 15]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
            //xSt.Range[excel.Cells[rowIndex, 14], excel.Cells[rowIndex, 15]].Font.Size = 8;
            //xSt.Range[excel.Cells[rowIndex, 14], excel.Cells[rowIndex, 15]].Value2 = "利润率";
            ////excel.Cells[rowIndex, 16] = Request.QueryString["Ratio"].ToString().Trim() + " ";

            //xSt.Range[excel.Cells[rowIndex + 1, 1], excel.Cells[rowIndex + 1, 16]].Borders.Weight = XlBorderWeight.xlHairline;
            //xSt.Range[excel.Cells[rowIndex + 1, 1], excel.Cells[rowIndex + 1, 16]].Borders[XlBordersIndex.xlEdgeTop].Weight = XlBorderWeight.xlThin;
            //xSt.Range[excel.Cells[rowIndex + 1, 1], excel.Cells[rowIndex + 1, 16]].Borders.LineStyle = 1;

            //xSt.Range[excel.Cells[rowIndex + 1, 1], excel.Cells[rowIndex + 1, 16]].RowHeight = 27;
            //xSt.Range[excel.Cells[rowIndex + 1, 1], excel.Cells[rowIndex + 1, 16]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
            //xSt.Range[excel.Cells[rowIndex + 1, 1], excel.Cells[rowIndex + 1, 16]].Font.Size = 8;
            //xSt.Range[excel.Cells[rowIndex + 1, 1], excel.Cells[rowIndex + 1, 16]].Font.Bold = true;

            //xSt.Range[excel.Cells[rowIndex + 1, 1], excel.Cells[rowIndex + 1, 3]].Merge(false);
            //xSt.Range[excel.Cells[rowIndex + 1, 1], excel.Cells[rowIndex + 1, 3]].Value2 = "制单人：" + userModel["name"].ToString().Trim();

            //xSt.Range[excel.Cells[rowIndex + 1, 4], excel.Cells[rowIndex + 1, 7]].Merge(false);
            //xSt.Range[excel.Cells[rowIndex + 1, 4], excel.Cells[rowIndex + 1, 7]].WrapText = true;
            //xSt.Range[excel.Cells[rowIndex + 1, 4], excel.Cells[rowIndex + 1, 7]].Value2 = "采购人：" + orderModel["purchasersName"];

            //xSt.Range[excel.Cells[rowIndex + 1, 8], excel.Cells[rowIndex + 1, 9]].Merge(false);
            //xSt.Range[excel.Cells[rowIndex + 1, 8], excel.Cells[rowIndex + 1, 9]].Value2 = "批准人：";

            //xSt.Range[excel.Cells[rowIndex + 1, 10], excel.Cells[rowIndex + 1, 12]].Merge(false);
            //xSt.Range[excel.Cells[rowIndex + 1, 10], excel.Cells[rowIndex + 1, 12]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
            //xSt.Range[excel.Cells[rowIndex + 1, 10], excel.Cells[rowIndex + 1, 12]].Value2 = "";

            //xSt.Range[excel.Cells[rowIndex + 1, 13], excel.Cells[rowIndex + 1, 16]].Merge(false);
            //xSt.Range[excel.Cells[rowIndex + 1, 13], excel.Cells[rowIndex + 1, 16]].Value2 = "";

            //#endregion

            excel.Visible = true;

            string filename = orderModel["_id"].ToString().Trim() + "_process.xlsx";
            //xSt.Export(Server.MapPath(".")+"\\"+this.xlfile.Text+".xls",SheetExportActionEnum.ssExportActionNone,Microsoft.Office.Interop.OWC.SheetExportFormat.ssExportHTML ); 
            xBk.SaveCopyAs(Server.MapPath("~/") + "temp\\" + orderModel["_id"].ToString().Trim() + "_process.xlsx");

            //ds = null;
            xBk.Close(false, null, null);

            excel.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xSt1);
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