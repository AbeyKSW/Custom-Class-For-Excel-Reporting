using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using System.Globalization;
using System.Data.SqlClient;
using System.Data.SqlTypes;
using System.Collections;
using Microsoft.Office.Interop;

namespace PerformanceViewer
{
    public partial class PerformanceViewer : Form
    {
        SqlConnection myConnection;

        #region Objects
        DataTable drill = new DataTable();
        DataTable distTBL = new DataTable();
        DataTable dataTBL = new DataTable();
        DataTable stkTBL = new DataTable();
        DataTable areaTBL = new DataTable();
        DataTable regTBL = new DataTable();
        DataTable invTBL = new DataTable();
        DataTable sndTBL = new DataTable();
        ArrayList productList = new ArrayList();
        ArrayList catIDList = new ArrayList();
        ArrayList monthList = new ArrayList();
        #endregion

        #region Declarations
        int divCode = 0;
        int row = 0;
        int col = 0;
        int repID = 0;
        int secCode = 0;
        int areaCode = 0;
        int stockiestID = 0;
        string sqlDataInv = "";
        string sqlDataSnd = "";
        string regCode = "";
        string regName = "";
        string areaName = "";
        string sqlStkName = "";
        string sqlSecName = "";
        string stockieName = "";
        string sectorName = "";
        string sqlStk = "";
        string sqlData = "";
        string sqlArea = "";
        string sqlReg = "";
        string fromDT = "";
        string toDT = "";

        double invValue = 0.00;
        double sndValue = 0.00;
        double totalValue = 0.00;
        double stkTotValue = 0.00;
        double areaTotValue = 0.00;
        double regTotValue = 0.00;
        #endregion

        public PerformanceViewer()
        {
            InitializeComponent();

            myConnection = new SqlConnection("Data Source=10.1.6.36,1433; Network Library=DBMSSOCN; Initial Catalog=SFAHeadofficeDBC; User ID=sa; Password=admin@Sfa99");
            myConnection.Open();
        }

        private void btnGenerate_Click(object sender, EventArgs e)
        {
            custoMeth cls = new custoMeth();

            //Declarations
            divCode = int.Parse(cmbDivCode.SelectedItem.ToString());
            fromDT = DateTime.Parse(dtFrom.Text.Trim()).ToString("yyyy/MM/dd", CultureInfo.InvariantCulture);
            toDT = DateTime.Parse(dtTo.Text.Trim()).ToString("yyyy/MM/dd", CultureInfo.InvariantCulture);

            //string month = cls.selectedmonth(this.dtFrom);

            // load excel, and create a new workbook
            Microsoft.Office.Interop.Excel.Application Excel = new Microsoft.Office.Interop.Excel.Application();
            Excel.Workbooks.Add();

            // single worksheet
            Microsoft.Office.Interop.Excel._Worksheet Worksheet = Excel.ActiveSheet;

            // Custom Header
            Worksheet.Range[Worksheet.Cells[1, 1], Worksheet.Cells[2, 8]].Merge();
            cls.excelHeader(Worksheet, 1, 1, "DARLEY BUTLER & CO. LIMITED", true, "18", "Center", "Center", 0, 0);

            Worksheet.Range[Worksheet.Cells[3, 1], Worksheet.Cells[3, 8]].Merge();
            string month = dtFrom.Value.ToString("MMMM", CultureInfo.InvariantCulture).ToUpper();
            cls.excelHeader(Worksheet, 3, 1, "RD SALES FOR THE MONTH OF " + month + " 2018", true, "15", "Center", "Center", 0, 0);

            Worksheet.Range[Worksheet.Cells[4, 1], Worksheet.Cells[4, 6]].Merge();
            cls.excelHeader(Worksheet, 4, 1, "Generated Date : " + System.DateTime.Today.ToShortDateString() + "", true, "none", "Left", "none", 0, 0);

            Worksheet.Range[Worksheet.Cells[5, 1], Worksheet.Cells[5, 6]].Merge();
            cls.excelHeader(Worksheet, 5, 1, "Performance Comparison (Division Code : " + divCode + ")", true, "none", "Left", "none", 0, 0);

            Worksheet.Range[Worksheet.Cells[6, 1], Worksheet.Cells[6, 6]].Merge();
            cls.excelHeader(Worksheet, 6, 1, ("Report Date: " + fromDT + " To " + toDT), true, "none", "Left", "none", 0, 0);

            cls.excelHeader(Worksheet, 8, 1, "REGIONAL MANAGER", true, "none", "Left", "none", 20, 0);
            cls.excelHeader(Worksheet, 8, 2, "AREA MANAGER", true, "none", "Left", "none", 22, 0);
            cls.excelHeader(Worksheet, 8, 3, "STOCKIEST", true, "none", "Left", "none", 47, 0);
            cls.excelHeader(Worksheet, 8, 4, "SECTOR", true, "none", "Left", "none", 24, 0);
            cls.excelHeader(Worksheet, 8, 5, "REPID", true, "none", "Left", "none", 6, 0);
            cls.excelHeader(Worksheet, 8, 6, "CALLSHEET VALUE", true, "none", "Left", "none", 17, 0);
            Worksheet.Columns[6].NumberFormat = "#,##0.00";
            cls.excelHeader(Worksheet, 8, 7, "SYSTEM VALUE", true, "none", "Left", "none", 17, 0);
            Worksheet.Columns[7].NumberFormat = "#,##0.00";
            cls.excelHeader(Worksheet, 8, 8, "DIFFERENCE", true, "none", "Left", "none", 17, 0);
            Worksheet.Columns[8].NumberFormat = "#,##0.00";
            Worksheet.Columns[9].NumberFormat = "#,##0.00";
            Worksheet.Columns[10].NumberFormat = "#,##0.00";
            Worksheet.Columns[11].NumberFormat = "#,##0.00";

            row = 9;

            if (cmbDivCode.Text == "16")
            { divCode = 15; }

            sqlReg = "SELECT RegCode, Name  FROM RegionMGR WHERE DivCode = '" + divCode + "'";
            regTBL = cls.data(sqlReg, myConnection, regTBL);

            prBar.Maximum = regTBL.Rows.Count;

            #region Data

            if (cmbDivCode.Text == "16")
            { divCode = 16; }

            //Region
            for (int i = 0; i < regTBL.Rows.Count; i++)
            {
                regTotValue = 0.00;

                regCode = regTBL.Rows[i]["RegCode"].ToString();
                regName = regTBL.Rows[i]["Name"].ToString();
                cls.excelValue(Worksheet, row, 1, regName, true);

                sqlArea = "SELECT AreaCode, AreaNme FROM AreaMGR WHERE RegMgr = '" + regCode + "' AND DivCode = '" + divCode + "'";
                areaTBL = cls.data(sqlArea, myConnection, areaTBL);

                prBar1.Maximum = areaTBL.Rows.Count;

                //Area
                for (int j = 0; j < areaTBL.Rows.Count; j++)
                {
                    areaTotValue = 0.00;

                    areaCode = int.Parse(areaTBL.Rows[j]["AreaCode"].ToString());
                    areaName = areaTBL.Rows[j]["AreaNme"].ToString();
                    cls.excelValue(Worksheet, row, 2, areaName, true);

                    //sqlStk = "SELECT StockiestID, SecCode, RepID FROM InvoiceHeader WHERE AreaCode = '" + areaCode + "' AND SalesDate BETWEEN '" + fromDT + "' AND '" + toDT + "' GROUP BY StockiestID, SecCode, RepID ORDER BY StockiestID, SecCode, RepID";
                    sqlStk = "SELECT StockiestID FROM InvoiceHeader WHERE AreaCode = '" + areaCode + "' AND SalesDate BETWEEN '" + fromDT + "' AND '" + toDT + "' GROUP BY StockiestID ORDER BY StockiestID";
                    stkTBL = cls.data(sqlStk, myConnection, stkTBL);

                    prBar2.Maximum = stkTBL.Rows.Count;

                    //Stockiest
                    for (int k = 0; k < stkTBL.Rows.Count; k++)
                    {
                        stkTotValue = 0.00;

                        stockiestID = int.Parse(stkTBL.Rows[k]["StockiestID"].ToString());

                        sqlStkName = "SELECT Name FROM Distributors WHERE Code = '" + stockiestID + "' AND DivCode = '" + divCode + "'";
                        stockieName = cls.GetValue<string>(sqlStkName, myConnection);
                        cls.excelValue(Worksheet, row, 3, stockieName, true);

                        sqlData = "SELECT StockiestID, SecCode, RepID FROM InvoiceHeader WHERE AreaCode = '" + areaCode + "' AND SalesDate BETWEEN '" + fromDT + "' AND '" + toDT + "' AND StockiestID = '" + stockiestID + "' GROUP BY StockiestID, SecCode, RepID ORDER BY StockiestID, SecCode, RepID";
                        dataTBL = cls.data(sqlData, myConnection, dataTBL);

                        //Sectorwise
                        for (int l = 0; l < dataTBL.Rows.Count; l++)
                        {
                            repID = int.Parse(dataTBL.Rows[l]["RepID"].ToString());
                            secCode = int.Parse(dataTBL.Rows[l]["SecCode"].ToString());

                            sqlSecName = "SELECT SectorName FROM Sector WHERE SectCode = '" + secCode + "' AND DivCode = '" + divCode + "'";
                            sectorName = cls.GetValue<string>(sqlSecName, myConnection);
                            cls.excelValue(Worksheet, row, 4, sectorName, false);
                            cls.excelValue(Worksheet, row, 5, repID, true);

                            sqlDataInv = "SELECT SUM(ItemValue) AS InvValue FROM InvoiceDetails INNER JOIN InvoiceHeader ON InvoiceDetails.InvID = InvoiceHeader.InvID ";
                            sqlDataInv = sqlDataInv + "AND InvoiceDetails.DailySalesID = InvoiceHeader.DailySalesID AND InvoiceDetails.RepID = InvoiceHeader.RepID AND InvoiceDetails.StockiestID = InvoiceHeader.StockiestID AND InvoiceDetails.DivCode = InvoiceHeader.DivCode ";
                            sqlDataInv = sqlDataInv + "WHERE InvoiceDetails.RepID = '" + repID + "' AND InvoiceDetails.StockiestID = '" + stockiestID + "' AND InvoiceDetails.InvoiceType IN ('Invoice') AND SalesDate BETWEEN '" + fromDT + "' AND '" + toDT + "' AND InvoiceDetails.DivCode = '" + divCode + "'";

                            invTBL = cls.data(sqlDataInv, myConnection, invTBL);
                            invValue = 0.00;
                            if (invTBL.Rows[0]["InvValue"].ToString() != null && invTBL.Rows[0]["InvValue"].ToString() != string.Empty)
                            { invValue = double.Parse(invTBL.Rows[0]["InvValue"].ToString()); }

                            sqlDataSnd = "SELECT SUM(ItemValue) AS SndValue FROM InvoiceDetails INNER JOIN InvoiceHeader ON InvoiceDetails.InvID = InvoiceHeader.InvID ";
                            sqlDataSnd = sqlDataSnd + "AND InvoiceDetails.DailySalesID = InvoiceHeader.DailySalesID AND InvoiceDetails.RepID = InvoiceHeader.RepID AND InvoiceDetails.StockiestID = InvoiceHeader.StockiestID AND InvoiceDetails.DivCode = InvoiceHeader.DivCode ";
                            sqlDataSnd = sqlDataSnd + "WHERE InvoiceDetails.RepID = '" + repID + "' AND InvoiceDetails.StockiestID = '" + stockiestID + "' AND InvoiceDetails.InvoiceType NOT IN ('Invoice') AND SalesDate BETWEEN '" + fromDT + "' AND '" + toDT + "' AND InvoiceDetails.DivCode = '" + divCode + "'";

                            sndTBL = cls.data(sqlDataSnd, myConnection, sndTBL);
                            sndValue = 0.00;
                            if (sndTBL.Rows[0]["SndValue"].ToString() != null && sndTBL.Rows[0]["SndValue"].ToString() != string.Empty)
                            { sndValue = double.Parse(sndTBL.Rows[0]["SndValue"].ToString()); }

                            totalValue = invValue - sndValue;
                            stkTotValue = stkTotValue + totalValue;

                            cls.excelValue(Worksheet, row, 6, 0.00, false);
                            cls.excelValue(Worksheet, row, 7, totalValue, false);

                            //Difference
                            Worksheet.Cells[row, 8].Formula = "=Sum(" + Worksheet.Cells[row, 6].Address + "-" + Worksheet.Cells[row, 7].Address + ")";

                            row = row + 1;
                        }

                        cls.excelValue(Worksheet, row, 9, stkTotValue, false);
                        areaTotValue = areaTotValue + stkTotValue;

                        row = row + 1;
                        k = k + 1;
                        prBar2.Value = k;
                        k = k - 1;
                    }

                    cls.excelValue(Worksheet, row, 10, areaTotValue, true);
                    regTotValue = regTotValue + areaTotValue;

                    row = row + 1;
                    j = j + 1;
                    prBar1.Value = j;
                    j = j - 1;
                }

                cls.excelValue(Worksheet, row, 11, regTotValue, true);

                row = row + 2;
                i = i + 1;
                prBar.Value = i;
                i = i - 1;
            }

            Worksheet.Range[Worksheet.Cells[8, 1], Worksheet.Cells[row, 8]].Cells.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

            #endregion

            Excel.Visible = true;
        }
    }
}
