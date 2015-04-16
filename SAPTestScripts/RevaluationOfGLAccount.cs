using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Young.Data;
using SAPAutoLogon;
using SAPGUIAutomationLib;
using SAPFEWSELib;
using Young.Excel;
using ex=Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.Windows.Forms;

namespace SAPTestScripts
{
    public class RevaluationOfGLAccount
    {
        private char split = '|';
        private TXTDataTableReader _reader;
        private DataTable _dt;
        private DataTable _curErrorTable;
        private Dictionary<string, string> _tcDic;
        private Dictionary<string, string> _lcDic;
        private List<CurrencyRate> _tcList;
        private List<CurrencyRate> _lcList;
        private DataTable _errorTable;

        public RevaluationOfGLAccount()
        {

        }


        public DataTable Read(string fileName)
        {
            _reader = new TXTDataTableReader(fileName);
            _dt = _reader.DefineTable(getColumns);
            _reader.Read(getRow);
            for (int i = 0; i < _dt.Columns.Count; i++)
            {
                DataColumn dcl = _dt.Columns[i];
                if (dcl.ColumnName.ToLower().Contains("column"))
                {
                    _dt.Columns.Remove(dcl);
                }
            }
            setCurrency();

            return _dt;
        }

        

        private void exportReport(string fileName)
        {
            
            ExcelHelper helper = new ExcelHelper();
            try
            {
                ex.Workbook wb = helper.NewWorkBook();
                ex.Worksheet ws = wb.Worksheets.Add(Missing.Value);
                ws.Name = "Error Data";
                ExcelHelper.Write(ws, 1, 1, _errorTable);
                ws = wb.Worksheets.Add(Missing.Value);
                ws.Name = "Raw Data";
                ExcelHelper.Write(ws, 1, 1, _dt);
                ws = wb.Worksheets.Add(Missing.Value);
                ws.Name = "Error Currency";
                ExcelHelper.Write(ws,1,1,_curErrorTable);
                wb.SaveAs(fileName);
                wb.Close();
            }
            
            catch
            {
                
            }
            finally
            {
                helper.MyExcel.Quit();
            }
            
            
        }

        private void getCurrency(Dictionary<string, string> curList, string date)
        {
            foreach (var k in curList.Keys.ToList())
            {
                SAPTestHelper.Current.GetElementById<GuiCTextField>("wnd[0]/usr/ctxtI1-LOW").Text = "M";
                SAPTestHelper.Current.GetElementById<GuiCTextField>("wnd[0]/usr/ctxtI2-LOW").Text = k;
                SAPTestHelper.Current.GetElementById<GuiCTextField>("wnd[0]/usr/ctxtI3-LOW").Text = "USD";
                SAPTestHelper.Current.GetElementById<GuiTextField>("wnd[0]/usr/txtI4-LOW").Text = date;
                SAPTestHelper.Current.GetElementById<GuiButton>("wnd[0]/tbar[1]/btn[8]").Press();
                curList[k] = SAPTestHelper.Current.GetElementById<GuiGridView>("wnd[0]/usr/cntlGRID1/shellcont/shell").GetCellValue(0, "UKURS");
                SAPTestHelper.Current.GetElementById<GuiButton>("wnd[0]/tbar[0]/btn[3]").Press();
            }

        }

        public void GetReport(string date,string reportFile)
        {
            _errorTable = _dt.Clone();
            

            var rows = from r in _dt.AsEnumerable()
                       where 
                       double.Parse(r["Delta LC/GC in USD"].ToString()) != 0 ||
                       double.Parse(r["Delta GC/LC in LOC"].ToString()) != 0 ||
                       double.Parse(r["Delta TC/LC/GC"].ToString()) != 0
                       select r;

            foreach (DataRow r in rows)
            {
                _errorTable.ImportRow(r);
            }

           
            Logon autoLogon = new Logon();
            autoLogon.StartLogon("LH4");
            SAPTestHelper.Current.SetSession();
            SAPTestHelper.Current.SAPGuiSession.StartTransaction("SE16");
            SAPTestHelper.Current.GetElementById<GuiCTextField>("wnd[0]/usr/ctxtDATABROWSE-TABLENAME").Text = "TCURR";
            SAPTestHelper.Current.GetElementById<GuiButton>("wnd[0]/tbar[0]/btn[0]").Press();

            getCurrency(_tcDic, date);
            getCurrency(_lcDic, date);


            SAPTestHelper.Current.SAPGuiConnection.CloseConnection();
            
            DataColumn dc1 = new DataColumn("Actual Rate LC/USD");
            _dt.Columns.Add(dc1);
            DataColumn dc2 = new DataColumn("Actual Rate TC/USD");
            _dt.Columns.Add(dc2);

            _curErrorTable = _dt.Clone();

            foreach (DataRow dr in _dt.Rows)
            {
                if (dr["CC Curr"].ToString() == "USD")
                {
                    dr["Actual Rate LC/USD"] = "1.000000000";
                }
                else
                {
                    dr["Actual Rate LC/USD"] = _lcDic[dr["CC Curr"].ToString()];
                }

                if (dr["TC"].ToString() == "USD")
                {
                    dr["Actual Rate TC/USD"] = "1.000000000";
                }
                else
                {
                    dr["Actual Rate TC/USD"] = _tcDic[dr["TC"].ToString()];
                }

                if (dr["Rate LC/USD"].ToString() != dr["Actual Rate LC/USD"].ToString() || dr["Rate TC/USD"].ToString() != dr["Actual Rate TC/USD"].ToString())
                {
                    _curErrorTable.ImportRow(dr);
                }
            }

            _dt.Columns.Remove(dc1);
            _dt.Columns.Remove(dc2);

            exportReport(reportFile);
        }

        private void setCurrency()
        {
            _tcDic = (from row in _dt.AsEnumerable()
                      where row["TC"].ToString().ToUpper() != "USD"
                      group row by row["TC"] into g
                      select g.Key.ToString()).ToDictionary(p => p);


            _lcDic = (from row in _dt.AsEnumerable()
                      where row["CC Curr"].ToString().ToUpper() != "USD"
                      group row by row["CC Curr"] into g
                      select g.Key.ToString()).ToDictionary(p => p);
        }

        private List<DataColumn> getColumns(string stringRow)
        {
            if (stringRow.Contains(split.ToString()))
            {
                var datas = stringRow.Split(split);
                if (datas.Count() > 7)
                {
                    List<DataColumn> columns = new List<DataColumn>();
                    for (int i = 0; i < datas.Count(); i++)
                    {
                        DataColumn col = new DataColumn(datas[i].Trim());
                        columns.Add(col);
                    }
                    return columns;
                }
            }
            return null;
        }

        private DataRow getRow(DataRow dr, string stringRow)
        {
            if (stringRow.Contains(split.ToString()) && stringRow.Contains("*") == false && stringRow != _reader.ColumnString)
            {
                var datas = stringRow.Split(split);
                var count = datas.Count();
                if (count > 7)
                {
                    for (int i = 0; i < _dt.Columns.Count; i++)
                    {
                        if (i < count)
                        {
                            datas[i] = datas[i].Replace('"', ' ').Replace(",", "").Trim();
                            if (datas[i].Contains('-') && datas[i].IndexOf('-') == datas[i].Length - 1)
                                datas[i] = '-' + datas[i].Substring(0, datas[i].Length - 1).Trim();
                            dr[i] = datas[i];
                        }
                    }
                    return dr;
                }
            }
            return null;
        }
    }
}
