using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Collections;
using System.Globalization;

using System.Data;
using System.Data.SqlClient;
using System.Data.SqlTypes;
using Microsoft.Office.Interop;

namespace PerformanceViewer
{
    class custoMeth
    {
        #region List Objects
        ArrayList listMonths = new ArrayList();
        #endregion

        #region Public Variables
        string fromDT = "";
        string toDT = "";
        #endregion

        #region selectedmonth
        public string selectedmonth(DateTimePicker dtFrom)
        {
            string DateF = "";
            string DTF = dtFrom.Text.ToString();
            string ddF = "";
            string mmF = "";
            string yyF = "";

            if (DTF != null)
            {
                string[] dateAr = DTF.Split('/');
                ddF = dateAr[0];
                mmF = dateAr[1];
                yyF = dateAr[2];

                DateF = yyF + mmF;
            }

            return DateF;
        }
        #endregion

        #region monthList
        public ArrayList monthsList(DateTimePicker dtFrom, DateTimePicker dtTo)
        {
            string sysFormat = CultureInfo.CurrentCulture.DateTimeFormat.ShortDatePattern;
            int months = 0;

            if (sysFormat != "dd/MM/yyyy")
            {
                fromDT = DateTime.Parse(dtFrom.Text.Trim()).ToString("dd/MM/yyyy", CultureInfo.InvariantCulture);
                toDT = DateTime.Parse(dtTo.Text.Trim()).ToString("dd/MM/yyyy", CultureInfo.InvariantCulture);
            }
            else
            {
                fromDT = dtFrom.Text.ToString();
                toDT = dtTo.Text.ToString();
            }

            string[] dateF = fromDT.Split('/');
            string ddF = dateF[0];
            string mmF = dateF[1];
            string yyF = dateF[2];

            string TRMMF = yyF + mmF;

            string[] dateT = toDT.Split('/');
            string ddT = dateT[0];
            string mmT = dateT[1];
            string yyT = dateT[2];

            string TRMMT = yyT + mmT;

            listMonths.Clear();

            if (int.Parse(yyT) == int.Parse(yyF))
            {
                if (int.Parse(mmT) == int.Parse(mmF))
                {
                    listMonths.Add(TRMMF);
                }
                else if (int.Parse(mmT) > int.Parse(mmF))
                {
                    months = int.Parse(mmT) - int.Parse(mmF);
                    listMonths.Add(TRMMF);
                    int ValF = int.Parse(TRMMF);
                    for (int q = 0; q <= months - 1; q++)
                    {
                        string NTRMMF = (ValF + 1).ToString();
                        listMonths.Add(NTRMMF);
                        ValF = ValF + 1;
                    }
                }
                else
                {
                    MessageBox.Show("Please select From Date less than To Date");
                }
            }
            else if (int.Parse(yyT) > int.Parse(yyF))
            {
                int yyFMnts = 12 - int.Parse(mmF);
                int midyrCnt = 0;
                int fYear = int.Parse(yyF);
                midyrCnt = int.Parse(yyT) - int.Parse(yyF);

                listMonths.Add(TRMMF);
                int datValF = int.Parse(TRMMF);
                for (int q = 0; q <= yyFMnts - 1; q++)
                {
                    string NTRMMF = (datValF + 1).ToString();
                    listMonths.Add(NTRMMF);
                    datValF = datValF + 1;
                }
                if (midyrCnt > 1)
                {
                    for (int s = 1; s < midyrCnt; s++)
                    {
                        fYear = fYear + 1;
                        fYear = fYear * 100;
                        for (int t = 1; t <= 12; t++)
                        {
                            int midMnth = fYear + t;
                            listMonths.Add(midMnth);
                        }
                        fYear = fYear / 100;
                    }
                }
                int yyTmnts = int.Parse(mmT) - int.Parse(mmT);
                string NTRMMT = yyT + "0" + yyTmnts;
                int datValT = int.Parse(NTRMMT);
                for (int r = 0; r <= int.Parse(mmT) - 1; r++)
                {
                    NTRMMT = (datValT + 1).ToString();
                    listMonths.Add(NTRMMT);
                    datValT = datValT + 1;
                }

                return listMonths;
            }
            else
            {
                MessageBox.Show("Please select From Date less than To Date");
                return listMonths;
            }
            return listMonths;
        }
        #endregion

        #region GetIntFromString
        public int GetIntFromString(string sqlQuery, SqlConnection conn)
        {
            SqlCommand queryCommand = new SqlCommand(sqlQuery, conn);

            int strVal = 0;
            try
            {
                SqlDataReader reader = queryCommand.ExecuteReader();

                if (reader.Read())
                {
                    if (!reader.IsDBNull(0))
                    {
                        strVal = reader.GetInt32(0);
                        reader.Close();
                        return strVal;
                    }
                    else { reader.Close(); return strVal = 0; }
                }
                else { reader.Close(); return strVal = 0; }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message.ToString());
                return strVal;
            }
        }
        #endregion

        #region GetStringFromString
        public string GetStringFromString(string sqlQuery, SqlConnection conn)
        {
            SqlCommand queryCommand = new SqlCommand(sqlQuery, conn);

            string strVal = "";
            try
            {
                SqlDataReader reader = queryCommand.ExecuteReader();

                if (reader.Read())
                {
                    if (!reader.IsDBNull(0))
                    {
                        strVal = reader.GetString(0);
                        reader.Close();
                        return strVal;
                    }
                    else { reader.Close(); return strVal = ""; }
                }
                else { reader.Close(); return strVal = ""; }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message.ToString());
                return strVal;
            }
        }
        #endregion

        #region excelHeader
        public void excelHeader(Microsoft.Office.Interop.Excel._Worksheet Worksheet, int row, int column, string value, Boolean bold, string size, string hAlign, string vAlign, int colWidth, int orientation)
        {
            Worksheet.Cells[row, column].Value = value;
            Worksheet.Cells[row, column].Font.Bold = bold;

            if (size != "none")
            { Worksheet.Cells[row, column].Font.Size = size; }

            if (colWidth != 0)
            { Worksheet.Columns[column].ColumnWidth = colWidth; }

            if (hAlign == "Left")
            { Worksheet.Cells[row, column].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft; }
            else if (hAlign == "Center")
            { Worksheet.Cells[row, column].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter; }
            else if (hAlign == "Right")
            { Worksheet.Cells[row, column].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight; }
            else
            { Worksheet.Cells[row, column].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignGeneral; }

            if (vAlign == "Left")
            { Worksheet.Cells[row, column].VerticalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft; }
            else if (vAlign == " Center")
            { Worksheet.Cells[row, column].VerticalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter; }
            else if (vAlign == "Right")
            { Worksheet.Cells[row, column].VerticalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight; }
            else
            { Worksheet.Cells[row, column].VerticalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignGeneral; }

            if (orientation != 0)
            { Worksheet.Cells[row, column].Orientation = orientation; }
        }
        #endregion

        #region excelValue
        public void excelValue(Microsoft.Office.Interop.Excel._Worksheet Worksheet, int row, int column, double value, Boolean bold)
        {
            Worksheet.Cells[row, column].Value = value;
            Worksheet.Cells[row, column].Font.Bold = bold;
        }
        #endregion

        #region excelValue
        public void excelValue(Microsoft.Office.Interop.Excel._Worksheet Worksheet, int row, int column, string value, Boolean bold)
        {
            Worksheet.Cells[row, column].Value = value;
            Worksheet.Cells[row, column].Font.Bold = bold;
        }
        #endregion

        #region releaseObject
        public void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Unable to release the Object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }
        #endregion

        #region IsOpened
        static bool IsOpened(string wbook)
        {
            bool isOpened = true;
            Microsoft.Office.Interop.Excel.Application exApp;
            exApp = (Microsoft.Office.Interop.Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
            try
            {
                exApp.Workbooks.get_Item(wbook);
            }
            catch (Exception)
            {
                isOpened = false;
            }
            return isOpened;
        }
        #endregion

        #region Data Table
        public DataTable data(string sql, SqlConnection conn, DataTable dt)
        {
            dt.Clear();
            SqlCommand queryCommand = new SqlCommand(sql, conn);
            queryCommand.CommandTimeout = 50;
            queryCommand.CommandType = CommandType.Text;
            SqlDataAdapter adapter = new SqlDataAdapter(queryCommand);
            adapter.Fill(dt);

            return dt;
        }
        #endregion

        #region Generics GetValue<T>
        public dynamic GetValue<T>(string sqlQuery, SqlConnection conn)
        {
            SqlCommand queryCommand = new SqlCommand(sqlQuery, conn);

            if (typeof(T) == typeof(int))
            {
                int strVal = 0;
                try
                {
                    SqlDataReader reader = queryCommand.ExecuteReader();

                    if (reader.Read())
                    {
                        if (!reader.IsDBNull(0))
                        {
                            strVal = reader.GetInt32(0);
                            reader.Close();
                            return strVal;
                        }
                        else { reader.Close(); return strVal = 0; }
                    }
                    else { reader.Close(); return strVal = 0; }
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message.ToString());
                    return strVal;
                }
            }
            else if (typeof(T) == typeof(string))
            {
                string strVal = "";
                try
                {
                    SqlDataReader reader = queryCommand.ExecuteReader();

                    if (reader.Read())
                    {
                        if (!reader.IsDBNull(0))
                        {
                            strVal = reader.GetString(0);
                            reader.Close();
                            return strVal;
                        }
                        else { reader.Close(); return strVal = ""; }
                    }
                    else { reader.Close(); return strVal = ""; }
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message.ToString());
                    return strVal;
                }
            }
            else if (typeof(T) == typeof(double))
            {
                double strVal = 0.00;
                try
                {
                    SqlDataReader reader = queryCommand.ExecuteReader();

                    if (reader.Read())
                    {
                        if (!reader.IsDBNull(0))
                        {
                            strVal = reader.GetDouble(0);
                            reader.Close();
                            return strVal;
                        }
                        else { reader.Close(); return strVal = 0.00; }
                    }
                    else { reader.Close(); return strVal = 0.00; }
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message.ToString());
                    return strVal;
                }
            }
            else if (typeof(T) == typeof(bool))
            {
                bool strVal;
                try
                {
                    SqlDataReader reader = queryCommand.ExecuteReader();

                    if (reader.Read())
                    {
                        if (!reader.IsDBNull(0))
                        {
                            strVal = true;
                            reader.Close();
                            return strVal;
                        }
                        else { reader.Close(); return strVal = false; }
                    }
                    else { reader.Close(); return strVal = false; }
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message.ToString());
                    return strVal = false;
                }
            }
            else if (typeof(T) == typeof(ArrayList))
            {
                ArrayList strVal = new ArrayList();
                try
                {
                    SqlDataReader reader = queryCommand.ExecuteReader();

                    while (reader.Read())
                    {
                        strVal.Add(reader.GetString(0));
                    }
                    reader.Close();
                    return strVal;
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message.ToString());
                    return strVal;
                }
            }
            else
            {
                return false;
            }
        }
        #endregion
    }
}
