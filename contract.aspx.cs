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
    public partial class contract : System.Web.UI.Page
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
        static extern int GetWindowThreadProcessId(IntPtr hWnd, out int lpdwProcessId);

        public void OutputExcel(dynamic orderModel, dynamic companyModel, dynamic customerModel, dynamic productModel, string strType)
        {
            GC.Collect();
            Application excel = new Application();
            _Workbook xBk = excel.Workbooks.Add(true);
            _Worksheet xSt = (_Worksheet)xBk.ActiveSheet;

            excel.DisplayAlerts = false;

            xSt.PageSetup.LeftMargin = 0.9 / 0.035;
            xSt.PageSetup.RightMargin = 0.9 / 0.035;
            xSt.PageSetup.HeaderMargin = 2 / 0.035;
            xSt.PageSetup.FooterMargin = 1 / 0.035;
            xSt.PageSetup.TopMargin = 3.3 / 0.035;
            xSt.PageSetup.BottomMargin = 1.8 / 0.035;
            xSt.PageSetup.LeftHeaderPicture.Filename = Server.MapPath("~/").ToString().Trim() + "image\\" + companyModel["nid"].ToString().Trim() + "logo.jpg";
            xSt.PageSetup.LeftHeader = "&G";
            xSt.PageSetup.CenterHeader = @"&""微软雅黑,Bold""&14" + "合同";
            xSt.PageSetup.RightHeader = @"&""微软雅黑""&9" + "合同单号：" + orderModel["orderNo"].ToString().Trim() + "　";
            xSt.PageSetup.LeftFooterPicture.Filename = Server.MapPath("~/").ToString().Trim() + "image\\" + "footline.jpg";
            xSt.PageSetup.LeftFooter = "&G";
            xSt.PageSetup.CenterFooter = @"&""微软雅黑""&8" + "共&N页，第&P页\n" + companyModel["name"].ToString().Trim() + "　　地址：" + companyModel["address"].ToString().Trim() + "　　电话：" + companyModel["phone"].ToString().Trim();

            excel.Cells.Font.Name = "微软雅黑";
            excel.Cells.Font.Size = 8;

            xSt.Range[excel.Cells[1, 1], excel.Cells[1, 1]].ColumnWidth = 5;
            xSt.Range[excel.Cells[1, 2], excel.Cells[1, 2]].ColumnWidth = 8;
            xSt.Range[excel.Cells[1, 3], excel.Cells[1, 3]].ColumnWidth = 36;
            xSt.Range[excel.Cells[1, 4], excel.Cells[1, 4]].ColumnWidth = 10;
            xSt.Range[excel.Cells[1, 5], excel.Cells[1, 5]].ColumnWidth = 8;
            xSt.Range[excel.Cells[1, 6], excel.Cells[1, 6]].ColumnWidth = 10;
            xSt.Range[excel.Cells[1, 7], excel.Cells[1, 7]].ColumnWidth = 12;

            xSt.Range[excel.Cells[1, 1], excel.Cells[1, 7]].RowHeight = 1;
            xSt.Range[excel.Cells[2, 1], excel.Cells[4, 7]].RowHeight = 15;

            excel.Cells[2, 1] = "购货方：" + customerModel["customerName"].ToString().Trim();
            excel.Cells[2, 4] = "（以下简称甲方）";
            excel.Cells[3, 1] = "供货方：" + companyModel["name"].ToString().Trim();
            excel.Cells[3, 4] = "（以下简称乙方）";
            excel.Cells[4, 1] = "合同日期：" + Convert.ToDateTime(orderModel["recordDate"].ToString()).ToString("yyyy年MM月dd日");

            xSt.Range[excel.Cells[5, 1], excel.Cells[5, 7]].RowHeight = 12;
            xSt.Range[excel.Cells[6, 1], excel.Cells[7, 7]].RowHeight = 16;

            excel.Cells[6, 1] = "感谢您的订单，以下是销售及交货条款。";
            xSt.Range[excel.Cells[6, 1], excel.Cells[6, 7]].Font.Bold = true;
            excel.Cells[7, 1] = "一、甲方向乙方购买如下产品：";

            //xSt.PageSetup.PrintTitleRows = "$8:$8";  标题行 
            xSt.Range[excel.Cells[8, 1], excel.Cells[8, 7]].RowHeight = 20;
            excel.Cells[8, 1] = "行号";
            xSt.Range[excel.Cells[8, 2], excel.Cells[8, 3]].Merge(false);
            xSt.Range[excel.Cells[8, 2], excel.Cells[8, 3]].Value2 = "产品描述";
            excel.Cells[8, 4] = "数量";
            excel.Cells[8, 5] = "单位";
            excel.Cells[8, 6] = "单价";
            excel.Cells[8, 7] = "总价";
            xSt.Range[excel.Cells[8, 1], excel.Cells[8, 7]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            xSt.Range[excel.Cells[8, 1], excel.Cells[8, 7]].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignBottom;

            int j = 8;
            for (int i = 0; i < Convert.ToInt32(orderModel["number"]); i++)
            {
                j = j + 6;

                xSt.Range[excel.Cells[j - 5, 1], excel.Cells[j - 5, 1]].RowHeight = 16;
                xSt.Range[excel.Cells[j - 4, 1], excel.Cells[j - 2, 1]].RowHeight = 13;
                xSt.Range[excel.Cells[j - 1, 1], excel.Cells[j, 1]].RowHeight = 14;

                xSt.Range[excel.Cells[j - 5, 1], excel.Cells[j - 5, 7]].Borders[XlBordersIndex.xlEdgeTop].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                xSt.Range[excel.Cells[j - 5, 1], excel.Cells[j - 5, 7]].Borders[XlBordersIndex.xlEdgeTop].Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlHairline;

                excel.Cells[j - 5, 1] = i + 1;
                xSt.Range[excel.Cells[j - 5, 1], excel.Cells[j - 5, 1]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                xSt.Range[excel.Cells[j - 5, 2], excel.Cells[j - 5, 3]].Merge(false);
                xSt.Range[excel.Cells[j - 5, 2], excel.Cells[j - 5, 3]].NumberFormat = "@";
                xSt.Range[excel.Cells[j - 5, 2], excel.Cells[j - 5, 3]].Value2 = (Convert.ToBoolean(productModel[i]["Sales_OrderProduct"]["isQuotation"].ToString()) ? productModel[i]["Sales_OrderProduct"]["quotationProductModel"].ToString().Trim() : productModel[i]["Product_Product"]["proNo"].ToString().Trim()) + " ";
                xSt.Range[excel.Cells[j - 5, 2], excel.Cells[j - 5, 3]].Font.Bold = true;
                xSt.Range[excel.Cells[j - 5, 2], excel.Cells[j - 5, 3]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
                xSt.Range[excel.Cells[j - 5, 4], excel.Cells[j - 5, 4]].NumberFormat = "@";
                excel.Cells[j - 5, 4] = (string.IsNullOrEmpty(productModel[i]["Sales_OrderProduct"]["quantity"].ToString()) ? 0 : Convert.ToDecimal(productModel[i]["Sales_OrderProduct"]["quantity"]).ToString("N2"));
                xSt.Range[excel.Cells[j - 5, 4], excel.Cells[j - 5, 4]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                excel.Cells[j - 5, 5] = productModel[i]["Product_Product"]["unit"].ToString().Trim() + " ";
                xSt.Range[excel.Cells[j - 5, 5], excel.Cells[j - 5, 5]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                excel.Cells[j - 5, 6] = (string.IsNullOrEmpty(productModel[i]["Sales_OrderProduct"]["unitPrice"].ToString()) ? "0" : Convert.ToDecimal(productModel[i]["Sales_OrderProduct"]["unitPrice"].ToString()).ToString("N2") + " ￥");
                xSt.Range[excel.Cells[j - 5, 6], excel.Cells[j - 5, 6]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
                excel.Cells[j - 5, 7] = (string.IsNullOrEmpty(productModel[i]["Sales_OrderProduct"]["totalPrice"].ToString()) ? "0" : Convert.ToDecimal(productModel[i]["Sales_OrderProduct"]["totalPrice"].ToString()).ToString("N2") + " ￥");
                xSt.Range[excel.Cells[j - 5, 7], excel.Cells[j - 5, 7]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
                xSt.Range[excel.Cells[j - 4, 2], excel.Cells[j - 1, 3]].Merge(false);
                xSt.Range[excel.Cells[j - 4, 2], excel.Cells[j - 1, 3]].Value2 = (Convert.ToBoolean(productModel[i]["Sales_OrderProduct"]["isQuotation"].ToString()) ? productModel[i]["Sales_OrderProduct"]["quotationProductName"].ToString().Trim() : productModel[i]["Product_Brand"]["chineseName"].ToString().Trim() + (string.IsNullOrEmpty(productModel[i]["Product_Brand"]["englishName"].ToString().Trim()) ? "" : "/" + productModel[i]["Product_Brand"]["englishName"].ToString().Trim()) + "　" + productModel[i]["Product_Product"]["name"].ToString().Trim() + (string.IsNullOrEmpty(productModel[i]["Product_Product"]["ordNo"].ToString().Trim()) ? "" : "　" + productModel[i]["Product_Product"]["ordNo"].ToString().Trim()) + (string.IsNullOrEmpty(productModel[i]["Product_Product"]["package"].ToString().Trim()) ? "" : "　" + productModel[i]["Product_Product"]["package"].ToString().Trim())) + " ";
                xSt.Range[excel.Cells[j - 4, 2], excel.Cells[j - 1, 3]].WrapText = true;
                xSt.Range[excel.Cells[j - 4, 2], excel.Cells[j - 1, 3]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
                xSt.Range[excel.Cells[j - 4, 2], excel.Cells[j - 1, 3]].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignTop;
                xSt.Range[excel.Cells[j, 2], excel.Cells[j, 3]].Merge(false);
                xSt.Range[excel.Cells[j, 2], excel.Cells[j, 3]].Value2 = productModel[i]["Sales_OrderProduct"]["delivery"].ToString().Trim();
                xSt.Range[excel.Cells[j, 2], excel.Cells[j, 3]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
            }

            xSt.Range[excel.Cells[j + 1, 1], excel.Cells[j + 1, 7]].Borders[XlBordersIndex.xlEdgeTop].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            xSt.Range[excel.Cells[j + 1, 1], excel.Cells[j + 1, 7]].Borders[XlBordersIndex.xlEdgeTop].Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlHairline;
            xSt.Range[excel.Cells[j + 1, 1], excel.Cells[j + 1, 1]].RowHeight = 22;
            excel.Cells[j + 1, 6] = "总金额：";
            xSt.Range[excel.Cells[j + 1, 6], excel.Cells[j + 1, 6]].Font.Bold = true;
            xSt.Range[excel.Cells[j + 1, 6], excel.Cells[j + 1, 6]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
            excel.Cells[j + 1, 7] = (string.IsNullOrEmpty(orderModel["totalPrice"].ToString()) ? "0" : Convert.ToDecimal(orderModel["totalPrice"].ToString()).ToString("N2") + " ￥");
            xSt.Range[excel.Cells[j + 1, 7], excel.Cells[j + 1, 7]].Font.Bold = true;
            xSt.Range[excel.Cells[j + 1, 7], excel.Cells[j + 1, 7]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;

            xSt.Range[excel.Cells[j + 2, 1], excel.Cells[j + 2, 1]].RowHeight = 22;
            xSt.Range[excel.Cells[j + 2, 1], excel.Cells[j + 2, 1]].Font.Bold = true;
            if (Convert.ToBoolean(orderModel["hasTaxNoInvoice"].ToString()))
            {
                excel.Cells[j + 2, 1] = "注：以上价格以人民币计算，不含税。";
            }
            else
            {
                excel.Cells[j + 2, 1] = "注：以上价格以人民币计算，含" + orderModel["hasTax"];
            }

            xSt.Range[excel.Cells[j + 3, 1], excel.Cells[j + 3, 1]].RowHeight = 16;
            excel.Cells[j + 3, 1] = "二、附加说明：";
            xSt.Range[excel.Cells[j + 3, 1], excel.Cells[j + 3, 1]].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignTop;
            xSt.Range[excel.Cells[j + 3, 3], excel.Cells[j + 3, 7]].Merge(false);
            xSt.Range[excel.Cells[j + 3, 3], excel.Cells[j + 3, 7]].Value2 = string.IsNullOrEmpty(orderModel["salesOrderFJSM"].ToString().Trim()) ? "无" : orderModel["salesOrderFJSM"].ToString().Trim();
            xSt.Range[excel.Cells[j + 3, 3], excel.Cells[j + 3, 7]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
            xSt.Range[excel.Cells[j + 3, 3], excel.Cells[j + 3, 7]].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignTop;

            xSt.Range[excel.Cells[j + 4, 1], excel.Cells[j + 4, 1]].RowHeight = 5;

            xSt.Range[excel.Cells[j + 5, 1], excel.Cells[j + 5, 1]].RowHeight = 16;
            excel.Cells[j + 5, 1] = "三、运费负担：";
            xSt.Range[excel.Cells[j + 5, 1], excel.Cells[j + 5, 1]].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignTop;
            xSt.Range[excel.Cells[j + 5, 3], excel.Cells[j + 5, 7]].Merge(false);
            xSt.Range[excel.Cells[j + 5, 3], excel.Cells[j + 5, 7]].Value2 = orderModel["salesOrderYFFD"].ToString().Trim();
            xSt.Range[excel.Cells[j + 5, 3], excel.Cells[j + 5, 7]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
            xSt.Range[excel.Cells[j + 5, 3], excel.Cells[j + 5, 7]].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignTop;

            xSt.Range[excel.Cells[j + 6, 1], excel.Cells[j + 6, 1]].RowHeight = 5;

            xSt.Range[excel.Cells[j + 7, 1], excel.Cells[j + 7, 1]].RowHeight = 16;
            excel.Cells[j + 7, 1] = "四、付款方式：";
            xSt.Range[excel.Cells[j + 7, 1], excel.Cells[j + 7, 1]].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignTop;
            xSt.Range[excel.Cells[j + 7, 3], excel.Cells[j + 7, 7]].Merge(false);
            xSt.Range[excel.Cells[j + 7, 3], excel.Cells[j + 7, 7]].Value2 = orderModel["salesOrderFKFS"].ToString().Trim();
            xSt.Range[excel.Cells[j + 7, 3], excel.Cells[j + 7, 7]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
            xSt.Range[excel.Cells[j + 7, 3], excel.Cells[j + 7, 7]].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignTop;

            xSt.Range[excel.Cells[j + 8, 1], excel.Cells[j + 8, 1]].RowHeight = 5;

            xSt.Range[excel.Cells[j + 9, 1], excel.Cells[j + 9, 1]].RowHeight = 25;
            excel.Cells[j + 9, 1] = "五、接受货物：";
            xSt.Range[excel.Cells[j + 9, 1], excel.Cells[j + 9, 1]].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignTop;
            xSt.Range[excel.Cells[j + 9, 3], excel.Cells[j + 9, 7]].Merge(false);
            xSt.Range[excel.Cells[j + 9, 3], excel.Cells[j + 9, 7]].Value2 = "1、甲方在签收货物前，需检验包装的完整性和货物的完整性。2、如甲方接收货物时发现包装损坏或产品损坏，甲方应拒绝签收并立即通知乙方。 3、甲方一旦签收货物，即表明甲方认为包装完好无损和货物完好无损。";
            xSt.Range[excel.Cells[j + 9, 3], excel.Cells[j + 9, 7]].WrapText = true;
            xSt.Range[excel.Cells[j + 9, 3], excel.Cells[j + 9, 7]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
            xSt.Range[excel.Cells[j + 9, 3], excel.Cells[j + 9, 7]].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignTop;

            xSt.Range[excel.Cells[j + 10, 1], excel.Cells[j + 10, 1]].RowHeight = 5;

            xSt.Range[excel.Cells[j + 11, 1], excel.Cells[j + 11, 1]].RowHeight = 48;
            excel.Cells[j + 11, 1] = "六、质量保证：";
            xSt.Range[excel.Cells[j + 11, 1], excel.Cells[j + 11, 1]].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignTop;
            xSt.Range[excel.Cells[j + 11, 3], excel.Cells[j + 11, 7]].Merge(false);
            xSt.Range[excel.Cells[j + 11, 3], excel.Cells[j + 11, 7]].Value2 = "乙方之产品对由于原材料和生产工艺缺陷而造成的损坏，自出厂之日保用12个月。但因使用不当、超载、改装或安装错误造成的损坏责任除外。经乙方认可的质量问题将获得免费的维修更换。本保用条款仅对乙方之产品而言，不含有其他的承诺，即明示或暗示保证。这种保用补偿仅限于合约上免费维修更换，乙方不负担其他责任，包括因事故、相关损失或特殊损坏等产生的责任。";
            xSt.Range[excel.Cells[j + 11, 3], excel.Cells[j + 11, 7]].WrapText = true;
            xSt.Range[excel.Cells[j + 11, 3], excel.Cells[j + 11, 7]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
            xSt.Range[excel.Cells[j + 11, 3], excel.Cells[j + 11, 7]].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignTop;

            xSt.Range[excel.Cells[j + 12, 1], excel.Cells[j + 12, 1]].RowHeight = 5;

            xSt.Range[excel.Cells[j + 13, 1], excel.Cells[j + 13, 1]].RowHeight = 16;
            excel.Cells[j + 13, 1] = "七、争议解决：";
            xSt.Range[excel.Cells[j + 13, 1], excel.Cells[j + 13, 1]].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignTop;
            xSt.Range[excel.Cells[j + 13, 3], excel.Cells[j + 13, 7]].Merge(false);
            xSt.Range[excel.Cells[j + 13, 3], excel.Cells[j + 13, 7]].Value2 = "在本合同履行过程中，甲乙双方如有争议，应通过友好协商解决。若协商不成， 应提交至天津仲裁委员会解决。";
            xSt.Range[excel.Cells[j + 13, 3], excel.Cells[j + 13, 7]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
            xSt.Range[excel.Cells[j + 13, 3], excel.Cells[j + 13, 7]].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignTop;

            xSt.Range[excel.Cells[j + 14, 1], excel.Cells[j + 14, 1]].RowHeight = 5;

            xSt.Range[excel.Cells[j + 15, 1], excel.Cells[j + 15, 1]].RowHeight = 16;
            excel.Cells[j + 15, 1] = "八、合同签订：";
            xSt.Range[excel.Cells[j + 15, 1], excel.Cells[j + 15, 1]].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignTop;
            xSt.Range[excel.Cells[j + 15, 3], excel.Cells[j + 15, 7]].Merge(false);
            xSt.Range[excel.Cells[j + 15, 3], excel.Cells[j + 15, 7]].Value2 = "本合同一式两份，需甲乙双方代表签字并加盖甲乙双方公章后生效。合同传真件与原件具有相同的法律效力。";
            xSt.Range[excel.Cells[j + 15, 3], excel.Cells[j + 15, 7]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
            xSt.Range[excel.Cells[j + 15, 3], excel.Cells[j + 15, 7]].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignTop;

            xSt.Range[excel.Cells[j + 16, 1], excel.Cells[j + 16, 1]].RowHeight = 5;

            xSt.Range[excel.Cells[j + 17, 1], excel.Cells[j + 17, 1]].RowHeight = 25;
            excel.Cells[j + 17, 1] = "九、发货信息：";
            xSt.Range[excel.Cells[j + 17, 1], excel.Cells[j + 17, 1]].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignTop;
            xSt.Range[excel.Cells[j + 17, 3], excel.Cells[j + 17, 7]].Merge(false);
            xSt.Range[excel.Cells[j + 17, 3], excel.Cells[j + 17, 7]].Value2 = "乙方将把本合同中的货物发送到：" + orderModel["salesOrderFHXX"].ToString().Trim();
            xSt.Range[excel.Cells[j + 17, 3], excel.Cells[j + 17, 7]].Font.Bold = true;
            xSt.Range[excel.Cells[j + 17, 3], excel.Cells[j + 17, 7]].WrapText = true;
            xSt.Range[excel.Cells[j + 17, 3], excel.Cells[j + 17, 7]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
            xSt.Range[excel.Cells[j + 17, 3], excel.Cells[j + 17, 7]].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignTop;

            xSt.Range[excel.Cells[j + 18, 1], excel.Cells[j + 18, 1]].RowHeight = 5;

            xSt.Range[excel.Cells[j + 19, 1], excel.Cells[j + 19, 1]].RowHeight = 25;

            if (!Convert.ToBoolean(orderModel["hasTaxNoInvoice"].ToString()))
            {
                excel.Cells[j + 19, 1] = "十、发票信息：";
                xSt.Range[excel.Cells[j + 19, 1], excel.Cells[j + 19, 1]].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignTop;
                xSt.Range[excel.Cells[j + 19, 3], excel.Cells[j + 19, 7]].Merge(false);
                xSt.Range[excel.Cells[j + 19, 3], excel.Cells[j + 19, 7]].Value2 = "乙方开出发票信息为：" + orderModel["salesOrderFPXX"].ToString().Trim();
                xSt.Range[excel.Cells[j + 19, 3], excel.Cells[j + 19, 7]].Font.Bold = true;
                xSt.Range[excel.Cells[j + 19, 3], excel.Cells[j + 19, 7]].WrapText = true;
                xSt.Range[excel.Cells[j + 19, 3], excel.Cells[j + 19, 7]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
                xSt.Range[excel.Cells[j + 19, 3], excel.Cells[j + 19, 7]].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignTop;
            }
            else
            {
                j = j - 1;
            }

            xSt.Range[excel.Cells[j + 20, 1], excel.Cells[j + 20, 1]].RowHeight = 16;

            xSt.Range[excel.Cells[j + 21, 1], excel.Cells[j + 27, 7]].RowHeight = 18;

            excel.Cells[j + 21, 1] = "甲方：" + customerModel["customerName"].ToString().Trim();
            excel.Cells[j + 21, 4] = "乙方：" + companyModel["name"].ToString().Trim();
            excel.Cells[j + 22, 1] = "甲方代表：" + (string.IsNullOrEmpty(orderModel["purchaserName"].ToString().Trim()) ? orderModel["userName"].ToString().Trim() : orderModel["purchaserName"].ToString().Trim());
            excel.Cells[j + 22, 4] = "乙方代表：" + orderModel["personInChargeName"].ToString().Trim() + "　　签名：";
            excel.Cells[j + 23, 1] = "甲方公章：";
            excel.Cells[j + 23, 4] = "乙方公章：";
            excel.Cells[j + 24, 1] = "地址：" + (customerModel["customerAddress"]["provinceName"].ToString().Trim() == "请选择省" ? "" : customerModel["customerAddress"]["provinceName"].ToString().Trim()) + (customerModel["customerAddress"]["cityName"].ToString().Trim() == "请选择市" || customerModel["customerAddress"]["cityName"].ToString().Trim() == "市辖区" ? "" : customerModel["customerAddress"]["cityName"].ToString().Trim()) + (customerModel["customerAddress"]["districtName"].ToString().Trim() == "请选择县/区" ? "" : customerModel["customerAddress"]["districtName"].ToString().Trim()) + customerModel["customerAddress"]["address"].ToString().Trim();
            excel.Cells[j + 24, 4] = "地址：" + companyModel["address"].ToString().Trim();

            Range range = xSt.Range[excel.Cells[j + 23, 5], excel.Cells[j + 23, 5]];
            xSt.Shapes.AddPicture(Server.MapPath("~/").ToString().Trim() + "/image/" + companyModel["nid"].ToString().Trim() + "OC.png", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoTrue, Convert.ToSingle(range.Left) - 10, Convert.ToSingle(range.Top) - 35, 106, 106);

            if (customerModel["customerPhone"].ToString()!="[]")
            {
                excel.Cells[j + 25, 1] = "电话：" + customerModel["customerPhone"][0]["phone"].ToString().Trim();
            }
            else
            {
                excel.Cells[j + 25, 1] = "电话：";
            }
            excel.Cells[j + 25, 4] = "电话：" + companyModel["phone"].ToString().Trim();

            excel.Cells[j + 26, 4] = "手机：" + orderModel["createByMobile"].ToString().Trim();
            excel.Cells[j + 27, 4] = "邮箱：" + orderModel["createByEmail"].ToString().Trim();

            xSt.Range[excel.Cells[j + 28, 1], excel.Cells[j + 28, 1]].RowHeight = 10;

            xSt.Range[excel.Cells[j + 29, 1], excel.Cells[j + 29, 1]].RowHeight = 25;
            xSt.Range[excel.Cells[j + 29, 1], excel.Cells[j + 29, 7]].Merge(false);
            xSt.Range[excel.Cells[j + 29, 1], excel.Cells[j + 29, 7]].WrapText = true;
            xSt.Range[excel.Cells[j + 29, 1], excel.Cells[j + 29, 7]].Font.Bold = true;
            xSt.Range[excel.Cells[j + 29, 1], excel.Cells[j + 29, 7]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
            xSt.Range[excel.Cells[j + 29, 1], excel.Cells[j + 29, 7]].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignTop;
            if (!Convert.ToBoolean(orderModel["hasTaxNoInvoice"].ToString()))
            {
                xSt.Range[excel.Cells[j + 29, 1], excel.Cells[j + 29, 7]].Value2 = "甲方付款信息：乙方开户行：" + companyModel["license"]["bank"].ToString().Trim() + "　帐号：" + companyModel["license"]["account"].ToString().Trim();
            }
            else
            {
                xSt.Range[excel.Cells[j + 29, 1], excel.Cells[j + 29, 7]].Value2 = "甲方付款信息：" + companyModel["bankInfo"].ToString().Trim();
            }

            xSt.Range[excel.Cells[j + 30, 1], excel.Cells[j + 32, 7]].RowHeight = 15.75;
            excel.Cells[j + 30, 1] = "********************************************************";
            excel.Cells[j + 31, 1] = "为提升我们的服务质量，作为重点客户，您可以直接联系我们的销售总监关于合作、建议、意见和投诉等事宜。";
            excel.Cells[j + 32, 1] = "邮箱：csr01@mro9.com";

            excel.Visible = true;

            string path = "";
            string filename = "";

            if (strType.ToLower() == "excel")
            {
                filename = orderModel["_id"] + ".xlsx";
                path = Server.MapPath("~/") + "temp\\" + orderModel["_id"] + ".xlsx";

                //保存excel
                xBk.SaveCopyAs(path);
            }
            else
            {
                filename = orderModel["_id"] + ".pdf";
                path = Server.MapPath("~/") + "temp\\" + orderModel["_id"] + ".pdf";

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