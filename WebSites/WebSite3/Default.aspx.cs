﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using Ex = Microsoft.Office.Interop.Excel;
using System.Data;
using System.Data.OleDb;
using System.Data.Sql;
using System.Data.SqlClient;
using System.IO;
using System.Configuration;


public partial class _Default : Page
{

    public void Page_Load(object sender, EventArgs e)
    {
        //List<string> Years = new List<string>();
        //Years.Add("2013-2014");
        //Years.Add("2014-2015");
        //ListBox1.DataSource = Years;

        //List<string> Sems = new List<string>();
        //Sems.Add("1-й семестр");
        //Sems.Add("2-й семестр");
        //ListBox2.DataSource = Sems;
    }
    public class GEx : System.Exception
    {
        public GEx(string message, Exception inner): base(message, inner)
        { }
    }
    public void Button1_Click(object sender, EventArgs e)
    {
        try
        {
            string savePath = FileHandle();
            ExcelWorks(savePath);
        }
        catch (GEx t)
        {
            Label1.Text = t.Message;
            Exception inner = t.InnerException;
            while (inner != null)
            {
                Label1.Text = t.Message;
                inner = inner.InnerException;
            }       

        }        
              
    }
    public class ExApp
    {
        public static Ex.Application app = null;
    }
    public class ExBook
    {
        public static Ex.Workbook book = null;
    }
    public class Extens
    {
        public static string filex = "";
    }
    
    protected void ExcelWorks(string path)
    {       

        try
        {
            //сделать для разных типов файлов
            string strExcelConn = "";
            strExcelConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + path + ";Extended Properties='Excel 8.0;HDR=YES;'";
            
            ExApp.app = new Ex.Application();
            ExBook.book = ExApp.app.Workbooks.Open(path);
            Ex.Sheets excelsheets = ExBook.book.Worksheets;
            Ex.Worksheet excelwsheet;
            Ex.Range rowstodelete;

            for (int i = 1; i < ExBook.book.Worksheets.Count; i++)
            {
                //выбор диапазона ячеек
                excelwsheet = (Ex.Worksheet)excelsheets.get_Item(i);
                rowstodelete = (Ex.Range)excelwsheet.Rows["1:9", Type.Missing];

                //действия со строками        
                rowstodelete.Delete(Ex.XlDirection.xlDown);
            }
                        
            Ex.Range frow;

            //работа с листами
            for (int i = 1; i <= ExBook.book.Worksheets.Count; i++)
            {
                excelwsheet = (Ex.Worksheet)excelsheets.get_Item(i);
                frow = (Ex.Range)excelwsheet.Rows["1:1"];

                string[] names = new string[excelwsheet.Columns.Count];

                object cells = excelwsheet.Cells;
                string shname = excelwsheet.Name.ToString();

                string cn = "";

                string command3 = "";
                string command4 = "";

                for (int j = 1; j <= excelwsheet.UsedRange.Columns.Count; j++)
                {                    
                    cn = GetExcelColumnName(j);
                    cells = excelwsheet.get_Range(cn + "1").Value;                    
                    
                    if (cells.ToString().Trim() == "ауд." || cells.ToString().Trim() == "Дни" || cells.ToString().Trim() == "Часы" || cells.ToString().Trim() == "неделя")
                    {
                        cells += j.ToString();
                    }

                    command3 = "[" + cells.ToString().Trim() + "]" + " nvarchar(max), ";
                    command4 += command3;

                    names[j - 1] = cells.ToString().Trim();

                }

                ExBook.book.Save();
                ExApp.app.DisplayAlerts = false;
                ExBook.book.Close();

                string command1 = "create table " + "[" + shname + " " + ListBox2.SelectedValue.ToString() + " " + ListBox1.SelectedValue.ToString() + "] (";
                string command2 = command1 + command4.Trim(',', ' ') + ");";                

                SqlCreateTab(shname, command2);
                DataTable dtExcel = RetrieveData(strExcelConn, shname);
                SqlBulkCopyImport(dtExcel, shname, names);

            }          
            
        }
        catch(NullReferenceException e)
        {
            GEx ex = new GEx(e.Message + "Возможно, что один из листов книги пуст", e);
            throw ex;            
        }
        catch (ApplicationException e)
        {
            GEx ex = new GEx(e.Message, e);
            throw ex;           
        }
        catch(SystemException e)
        {
            GEx ex = new GEx(e.Message, e);
            throw ex;
        }
        finally
        {            
            if (ExApp.app != null)
            {
                System.Diagnostics.Process[] process = System.Diagnostics.Process.GetProcessesByName("Excel");
                foreach (System.Diagnostics.Process p in process)
                {
                    if (!string.IsNullOrEmpty(p.ProcessName))
                    {
                        try
                        {
                            p.Kill();
                        }
                        catch { }
                    }
                }
                //ExApp.app.DisplayAlerts = false;
                //ExApp.app.Quit();
                //GC.Collect();
                //ExApp.app.DisplayAlerts = false;
                //System.Runtime.InteropServices.Marshal.ReleaseComObject(ExApp.app);
                //ExApp.app = null;
            }
        }
               
    }

    private string GetExcelColumnName(int columnNumber)
    {
        int dividend = columnNumber;
        string columnName = String.Empty;
        int modulo;

        while (dividend > 0)
        {
            modulo = (dividend - 1) % 26;
            columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
            dividend = (int)((dividend - modulo) / 26);
        }

        return columnName;
    }

    protected string FileHandle()
    {
        try
        {
            //работа с файлом
            //string savePath = @"C:\Users\Denis\Documents\Visual Studio 2013\WebSites\WebSite3\Files\";
            //string fileName = FileUpload1.FileName;
            //savePath += fileName;
            //FileUpload1.SaveAs(savePath);
            //TextBox1.Text = savePath;
            string savePath = "";

            if(FileUpload1.HasFile)
            {
                string filen = Server.HtmlEncode(FileUpload1.FileName);

                Extens.filex = Path.GetExtension(filen);

                if(Extens.filex != ".xls" && Extens.filex != ".xlsx")
                {
                    Label1.Text = "Загружаемый файл должен являться таблицей Excel";
                }

                string filenup = "~/Files/" + DateTime.Now.ToString("yyyyMMddHHmmss") + Extens.filex;

                FileUpload1.SaveAs(Server.MapPath(filenup));

                savePath = Server.MapPath(filenup);

                TextBox1.Text = savePath;
            }            

            return savePath;
        }
        catch(FileNotFoundException e)
        {
            GEx ex = new GEx(e.Message, e);
            throw ex;
        }
        catch(SystemException e)
        {
            GEx ex = new GEx(e.Message, e);
            throw ex;
        }
    }

    protected void SqlCreateTab(string shn, string command)
    {
        using (SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["SQL2012"].ToString()))
        {
            try
            {
                conn.Open();
                SqlCommand crcom = new SqlCommand();
                crcom.Connection = conn;
                crcom.CommandText = command;

                crcom.ExecuteNonQuery();
            }
            catch (SqlException e)
            {
                GEx ex = new GEx(e.Message, e);
                throw ex;
            }
            catch(SystemException e)
            {
                GEx ex = new GEx(e.Message, e);
                throw ex;
            }
            finally
            {
                conn.Close();
            }
        }
    }

    protected DataTable RetrieveData(string strConn, string ShName)
    {
        DataTable dtExcel = new DataTable();

        using (OleDbConnection conn = new OleDbConnection(strConn))
        {
            try
            {
                //получить все листы
                OleDbDataAdapter da = new OleDbDataAdapter("select * from [" + ShName + "$]", conn);

                da.Fill(dtExcel);
            }
            catch(DataException e)
            {
                GEx ex = new GEx(e.Message, e);
                throw ex;                
            }
            catch(SystemException e)
            {
                GEx ex = new GEx(e.Message, e);
                throw ex;
            }
        }

        return dtExcel;
    }

    protected void SqlBulkCopyImport(DataTable dtExcel, string tabName, string[] colName)
    {
        using (SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["SQL2012"].ToString()))
        {
            try
            {
                conn.Open();

                using (SqlBulkCopy bulkCopy = new SqlBulkCopy(conn))
                {
                    bulkCopy.DestinationTableName = "[" + tabName + " " + ListBox2.SelectedValue.ToString() + " " + ListBox1.SelectedValue.ToString() + "]";

                    foreach (DataColumn dc in dtExcel.Columns)
                    {
                        bulkCopy.ColumnMappings.Add(dc.ColumnName, colName[dc.Ordinal]);
                    }

                    bulkCopy.WriteToServer(dtExcel);
                }
            }
            catch (SqlException e)
            {
                GEx ex = new GEx(e.Message, e);
                throw ex;
            }
            catch(SystemException e)
            {
                GEx ex = new GEx(e.Message, e);
                throw ex;
            }
            finally
            {
                conn.Close();
            }
        }
    }

    protected void Button2_Click(object sender, EventArgs e)
    {
        if (ExApp.app != null)
        {
            ExBook.book.Close(null, null, null);
            ExApp.app.Quit();
            ExApp.app = null;
        }
    }
}