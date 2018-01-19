using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using Aspose.Cells;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Office.Interop.Excel;
using MongoDB.Driver;
using Application = System.Windows.Forms.Application;
using Border = DocumentFormat.OpenXml.Spreadsheet.Border;
using Borders = DocumentFormat.OpenXml.Spreadsheet.Borders;
using Cell = DocumentFormat.OpenXml.Spreadsheet.Cell;
using CellFormat = DocumentFormat.OpenXml.Spreadsheet.CellFormat;
using Color = System.Drawing.Color;
using DataTable = System.Data.DataTable;
using Font = DocumentFormat.OpenXml.Spreadsheet.Font;
using Range = Microsoft.Office.Interop.Excel.Range;
using Row = DocumentFormat.OpenXml.Spreadsheet.Row;
using Sheets = DocumentFormat.OpenXml.Spreadsheet.Sheets;
using Workbook = Aspose.Cells.Workbook;
using Worksheet = Microsoft.Office.Interop.Excel.Worksheet;

// I fucking hate this.

/// <Todo>
///
/// <--- Todo --->
/// 
/// <!!! TOP PRIORITY !!!>
/// . Test Performance - Optimize Outputing to Excel. Interop is way too fucking slow. I gave up coz Memory Overflow.
/// . Update: Ultilizing OpenXMLWriter ( SAX Method ) whenever possible. Finally able to Custom Format on it.
/// . ( Done Jun 22, 17 ) YesNoKPI for ThuMua in Coord. Freaking Filtering in Looping in Preparing. 
/// . ( Done Jun 26, 17 ) Fixed rate. Fucking rate. 
/// 
/// <* High Priority *>
/// . ( Done May 01, 17 ) Actual Demand / Supply Function ( Remove UpperCap that's currently 100%. )
/// 
/// <. low priority .>
/// o Redesign Mastah UI, for dynamic forecast updating.
/// . ( Done Jul 25, 17 ) Enable Reading PO in pieces ( seperated files. )
/// . ( Done May 01, 17 ) Print DBSL into VE Farm or whatever 
/// . ( Done May 01, 17 ) Print ThuMua into VE ThuMua 
/// . ( Done May 07, 17 ) Fucking formula 
/// . ( Done May 01, 17 ) And uhm Region stuff I guess 
///  
/// </Todo>
namespace AllocatingStuff
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
            richTextBoxOutput.TextChanged += RichTextBoxOutput_TextChanged;
        }

        /// <summary>
        ///     Print Purchase Order, either horizontally ( true ) or vertically ( false )
        /// </summary>
        /// <param name="YesNoHorizontal"></param>
        private void PrintPO(string Choice, bool YesNoByUnit = false)
        {
            try
            {
                WriteToRichTextBoxOutput("Start!");

                var stopwatch = Stopwatch.StartNew();

                var mongoClient = new MongoClient();
                var db = mongoClient.GetDatabase("localtest");

                var PO = db.GetCollection<PurchaseOrderDate>("PurchaseOrder").Find(x =>
                        x.DateOrder >= DateFrom.Date &&
                        x.DateOrder <= DateTo.Date)
                    .ToList()
                    .OrderBy(x => x.DateOrder);

                var Product = db.GetCollection<Product>("Product").AsQueryable().ToList();
                var Customer = db.GetCollection<Customer>("Customer").AsQueryable().ToList();
                var dicProductUnit = db.GetCollection<ProductUnit>("ProductUnit").AsQueryable()
                    .ToDictionary(x => x.ProductCode);

                var dicProduct = new Dictionary<Guid, Product>();
                var dicCustomer = new Dictionary<Guid, Customer>();

                var dicUnit = new Dictionary<string, string>(StringComparer.Ordinal);

                //var dicClass = new Dictionary<string, string>
                //{
                //    { "A", "Rau ăn lá" },
                //    { "B", "Rau ăn thân hoa" },
                //    { "C", "Rau ăn quả" },
                //    { "D", "Rau ăn củ" },
                //    { "E", "Hạt" },
                //    { "F", "Rau gia vị" },
                //    { "G", "Thủy canh" },
                //    { "H", "Rau mầm" },
                //    { "I", "Nấm" },
                //    { "J", "Lá" },
                //    { "K", "Trái cây (Quả)" }
                //};

                foreach (var _Product in Product)
                    //if (dicClass.TryGetValue(_Product.ProductCode.Substring(0, 1), out string _ProductClass))
                    //{
                    //    _Product.ProductClassification = _ProductClass;
                    //}
                    dicProduct.Add(_Product.ProductId, _Product);

                foreach (var _Customer in Customer)
                {
                    _Customer.CustomerType = _Customer.CustomerType.Replace("Priority", "").Trim();
                    dicCustomer.Add(_Customer.CustomerId, _Customer);
                }

                var dtPO = new DataTable
                {
                    TableName = string.Format("PO {0} - {1}", DateFrom.Date.ToString("dd.MM"),
                        DateTo.Date.ToString("dd.MM"))
                };

                var DicColDate = new Dictionary<string, int>();

                if (Choice == "Horizontal")
                {
                    dtPO.TableName += " Horizontal";

                    dtPO.Columns.Add("VE Code", typeof(string));
                    dtPO.Columns.Add("VE Name", typeof(string));
                    dtPO.Columns.Add("Class", typeof(string));
                    dtPO.Columns.Add("StoreCode", typeof(string));
                    dtPO.Columns.Add("StoreName", typeof(string));
                    dtPO.Columns.Add("StoreType", typeof(string));
                    dtPO.Columns.Add("SubRegion", typeof(string));
                    dtPO.Columns.Add("Region", typeof(string));
                    dtPO.Columns.Add("P&L", typeof(string));
                    dtPO.Columns.Add("Unit", typeof(string)).DefaultValue = "Kg";
                    dtPO.Columns.Add("Note", typeof(string));

                    foreach (var PODate in PO)
                        dtPO.Columns.Add(PODate.DateOrder.Date.ToString("MM/dd/yyyy"), typeof(double)).DefaultValue = 0;

                    var dicRow = new Dictionary<string, int>();
                    foreach (var PODate in PO)
                        foreach (var _ProductOrder in PODate.ListProductOrder)
                            foreach (var _CustomerOrder in _ProductOrder.ListCustomerOrder)
                            {
                                if (!YesNoByUnit)
                                {
                                    var _OrderUnitType = ProperUnit(_CustomerOrder.Unit.ToLower(), dicUnit);
                                    _CustomerOrder.Unit = _OrderUnitType;
                                    if (dicCustomer[_CustomerOrder.CustomerId].CustomerType == "VM+" && _OrderUnitType != "Kg")
                                    {
                                        var _ProductCode = dicProduct[_ProductOrder.ProductId].ProductCode;
                                        if (dicProductUnit.TryGetValue(_ProductCode, out var _ProductUnit))
                                        {
                                            var _ProductUnitRegion =
                                                _ProductUnit.ListRegion.FirstOrDefault(x => x.OrderUnitType == _OrderUnitType);
                                            if (_ProductUnitRegion != null)
                                            {
                                                _CustomerOrder.Unit = _OrderUnitType;
                                                _CustomerOrder.QuantityOrderKg =
                                                    _CustomerOrder.QuantityOrder * _ProductUnitRegion.OrderUnitPer;
                                            }
                                        }
                                        else
                                        {
                                            _CustomerOrder.Unit = "Kg";
                                            _CustomerOrder.QuantityOrderKg = _CustomerOrder.QuantityOrder;
                                        }
                                    }
                                    else
                                    {
                                        _CustomerOrder.Unit = "Kg";
                                        _CustomerOrder.QuantityOrderKg = _CustomerOrder.QuantityOrder;
                                    }
                                }

                                var _Product = dicProduct[_ProductOrder.ProductId];
                                var _Customer = dicCustomer[_CustomerOrder.CustomerId];

                                DataRow dr = null;

                                var sKey = _Product.ProductCode + _Customer.CustomerCode;
                                if (!dicRow.TryGetValue(sKey, out var _rowIndex))
                                {
                                    dr = dtPO.NewRow();
                                    dicRow.Add(sKey, dtPO.Rows.Count);

                                    dr["VE Code"] = _Product.ProductCode;
                                    dr["VE Name"] = _Product.ProductName;
                                    dr["Class"] = _Product.ProductClassification;
                                    dr["StoreCode"] = _Customer.CustomerCode;
                                    dr["StoreName"] = _Customer.CustomerName;
                                    dr["StoreType"] = _Customer.CustomerType;
                                    dr["SubRegion"] = _Customer.CustomerRegion;
                                    dr["Region"] = _Customer.CustomerBigRegion;
                                    dr["P&L"] = _Customer.Company;
                                    dr["Note"] =
                                        _Product.ProductNote.Contains(_Customer.CustomerBigRegion == "Miền Nam"
                                            ? "South"
                                            : "North")
                                            ? "Ok"
                                            : "Out of List";
                                    if (YesNoByUnit)
                                    {
                                        dr["Unit"] = _CustomerOrder.Unit;
                                        dr[PODate.DateOrder.Date.ToString("MM/dd/yyyy")] = _CustomerOrder.QuantityOrder;
                                    }
                                    else
                                    {
                                        dr[PODate.DateOrder.Date.ToString("MM/dd/yyyy")] = _CustomerOrder.QuantityOrderKg;
                                    }
                                    dtPO.Rows.Add(dr);
                                }
                                else
                                {
                                    dr = dtPO.Rows[_rowIndex];
                                    if (YesNoByUnit)
                                        dr[PODate.DateOrder.Date.ToString("MM/dd/yyyy")] =
                                            (double)dr[PODate.DateOrder.Date.ToString("MM/dd/yyyy")] +
                                            _CustomerOrder.QuantityOrder;
                                    else
                                        dr[PODate.DateOrder.Date.ToString("MM/dd/yyyy")] =
                                            (double)dr[PODate.DateOrder.Date.ToString("MM/dd/yyyy")] +
                                            _CustomerOrder.QuantityOrderKg;
                                }
                            }
                }
                // Vertical PO - making it pivot-able ( No pun intended. )
                else if (Choice == "Vertical")
                {
                    dtPO.TableName += " Vertical";

                    dtPO.Columns.Add("PCODE", typeof(string)).DefaultValue = "";
                    dtPO.Columns.Add("PNAME", typeof(string)).DefaultValue = "";
                    dtPO.Columns.Add("PCLASS", typeof(string)).DefaultValue = "";
                    dtPO.Columns.Add("NOTE", typeof(string)).DefaultValue = "";
                    dtPO.Columns.Add("Climate", typeof(string)).DefaultValue = "";
                    dtPO.Columns.Add("CCODE", typeof(string)).DefaultValue = "";
                    dtPO.Columns.Add("CNAME", typeof(string)).DefaultValue = "";
                    dtPO.Columns.Add("CTYPE", typeof(string)).DefaultValue = "";
                    dtPO.Columns.Add("CREGION", typeof(string)).DefaultValue = "";
                    dtPO.Columns.Add("REGION", typeof(string)).DefaultValue = "";
                    dtPO.Columns.Add("P&L", typeof(string)).DefaultValue = "";
                    dtPO.Columns.Add("DateOrder", typeof(int)).DefaultValue = 0;
                    dtPO.Columns.Add("QuantityOrder", typeof(double)).DefaultValue = 0;
                    dtPO.Columns.Add("DateReceive", typeof(int)).DefaultValue = 0;

                    DicColDate.Add("DateOrder", dtPO.Columns.IndexOf("DateOrder"));
                    DicColDate.Add("DateReceive", dtPO.Columns.IndexOf("DateReceive"));

                    var dicRow = new Dictionary<string, int>();
                    foreach (var PODate in PO)
                        foreach (var _ProductOrder in PODate.ListProductOrder)
                            foreach (var _CustomerOrder in _ProductOrder.ListCustomerOrder)
                            {
                                var _Product = dicProduct[_ProductOrder.ProductId];
                                var _Customer = dicCustomer[_CustomerOrder.CustomerId];

                                DataRow dr = null;

                                var _rowIndex = 0;
                                var sKey = _Product.ProductCode + _Customer.CustomerCode + _Customer.Company +
                                           PODate.DateOrder.Date.ToString("yyyyMMdd");

                                if (!dicRow.TryGetValue(sKey, out _rowIndex))
                                {
                                    dr = dtPO.NewRow();
                                    dicRow.Add(sKey, dtPO.Rows.Count);

                                    dr["PCODE"] = _Product.ProductCode;
                                    dr["PNAME"] = _Product.ProductName;
                                    dr["PCLASS"] = _Product.ProductClassification;
                                    dr["CCODE"] = _Customer.CustomerCode;
                                    dr["CNAME"] = _Customer.CustomerName;
                                    dr["CTYPE"] = _Customer.CustomerType;
                                    dr["CREGION"] = _Customer.CustomerRegion;
                                    dr["REGION"] = _Customer.CustomerBigRegion;
                                    dr["P&L"] = _Customer.Company;
                                    dr["Note"] =
                                        _Product.ProductNote.Contains(_Customer.CustomerBigRegion == "Miền Nam"
                                            ? "South"
                                            : "North")
                                            ? "Ok"
                                            : "Out of List";
                                    dr["Climate"] = _Product.ProductClimate;
                                    dr["DateOrder"] = (int)(PODate.DateOrder.Date - new DateTime(1900, 1, 1)).TotalDays + 2;
                                    dr["QuantityOrder"] = _CustomerOrder.QuantityOrder;
                                    dr["DateReceive"] = (string)dr["REGION"] == "Miền Nam"
                                        ? (int)dr["DateOrder"] + 1
                                        : (int)dr["DateOrder"];

                                    dtPO.Rows.Add(dr);
                                }
                                else
                                {
                                    dr = dtPO.Rows[_rowIndex];
                                    dr["QuantityOrder"] = (double)dr["QuantityOrder"] + _CustomerOrder.QuantityOrder;
                                }
                            }
                }
                // Compact PO - Condensing PO to Customer Type level.
                else if (Choice == "Compact")
                {
                    dtPO.TableName += " Compact";

                    dtPO.Columns.Add("PCODE", typeof(string));
                    dtPO.Columns.Add("PNAME", typeof(string));
                    dtPO.Columns.Add("PCLASS", typeof(string));
                    dtPO.Columns.Add("ProductOrientation", typeof(string));
                    dtPO.Columns.Add("ProductClimate", typeof(string));
                    dtPO.Columns.Add("ProductionGroup", typeof(string));
                    dtPO.Columns.Add("Note", typeof(string));
                    dtPO.Columns.Add("CustomerCode", typeof(string));
                    dtPO.Columns.Add("CustomerName", typeof(string));
                    dtPO.Columns.Add("CTYPE", typeof(string));
                    dtPO.Columns.Add("CREGION", typeof(string));
                    dtPO.Columns.Add("P&L", typeof(string));
                    dtPO.Columns.Add("REGION", typeof(string));
                    dtPO.Columns.Add("DateOrder", typeof(DateTime));
                    dtPO.Columns.Add("QuantityOrder", typeof(double)).DefaultValue = 0;

                    DicColDate.Add("DateOrder", dtPO.Columns.IndexOf("DateOrder"));

                    var dicRow = new Dictionary<string, int>();
                    foreach (var PODate in PO)
                        foreach (var _ProductOrder in PODate.ListProductOrder)
                            foreach (var _CustomerOrder in _ProductOrder.ListCustomerOrder)
                            {
                                var _Product = dicProduct[_ProductOrder.ProductId];
                                var _Customer = dicCustomer[_CustomerOrder.CustomerId];

                                DataRow dr = null;

                                var _rowIndex = 0;
                                var sKey = _Product.ProductCode + _Customer.CustomerBigRegion +
                                           _Customer.CustomerRegion + _Customer.CustomerType + _Customer.Company +
                                           PODate.DateOrder.Date.ToString("yyyyMMdd");

                                if (!dicRow.TryGetValue(sKey, out _rowIndex))
                                {
                                    dr = dtPO.NewRow();
                                    dicRow.Add(sKey, dtPO.Rows.Count);

                                    dr["PCODE"] = _Product.ProductCode;
                                    dr["PNAME"] = _Product.ProductName;
                                    dr["PCLASS"] = _Product.ProductClassification;
                                    dr["ProductOrientation"] = _Product.ProductOrientation;
                                    dr["ProductClimate"] = _Product.ProductClimate;
                                    dr["ProductionGroup"] = _Product.ProductionGroup;
                                    //dr["CustomerCode"] = _Customer.CustomerCode;
                                    //dr["CustomerName"] = _Customer.CustomerName;
                                    dr["Note"] =
                                        _Product.ProductNote.Contains(_Customer.CustomerBigRegion == "Miền Nam"
                                            ? "South"
                                            : "North")
                                            ? "Ok"
                                            : "Out of List";
                                    dr["CTYPE"] = _Customer.CustomerType;
                                    dr["CREGION"] = _Customer.CustomerRegion;
                                    dr["P&L"] = _Customer.Company;
                                    dr["REGION"] = _Customer.CustomerBigRegion;
                                    //dr["DateOrder"] = (int)(PODate.DateOrder.Date - new DateTime(1900, 1, 1)).TotalDays + 2;
                                    dr["DateOrder"] = PODate.DateOrder.Date;
                                    dr["QuantityOrder"] = _CustomerOrder.QuantityOrder;

                                    dtPO.Rows.Add(dr);
                                }
                                else
                                {
                                    dr = dtPO.Rows[_rowIndex];
                                    dr["QuantityOrder"] = (double)dr["QuantityOrder"] + _CustomerOrder.QuantityOrder;
                                }
                            }
                }
                else if (Choice == "Report")
                {
                    dtPO.TableName += " Report";

                    dtPO.Columns.Add("REGION", typeof(string));
                    dtPO.Columns.Add("CTYPE", typeof(string));
                    dtPO.Columns.Add("PCODE", typeof(string));
                    dtPO.Columns.Add("SRegion", typeof(string));
                    dtPO.Columns.Add("Nguồn", typeof(string)).DefaultValue = "VCM";
                    dtPO.Columns.Add("DateReceive", typeof(DateTime));
                    dtPO.Columns.Add("DateProcess", typeof(string));
                    dtPO.Columns.Add("Supplier", typeof(string));
                    dtPO.Columns.Add("QuantityOrder", typeof(double)).DefaultValue = 0;
                    dtPO.Columns.Add("Source", typeof(string)).DefaultValue = "PO";
                    dtPO.Columns.Add("DateOrder", typeof(DateTime));
                    dtPO.Columns.Add("P&L", typeof(string));
                    dtPO.Columns.Add("Note", typeof(string));

                    DicColDate.Add("DateOrder", dtPO.Columns.IndexOf("DateOrder"));
                    DicColDate.Add("DateReceive", dtPO.Columns.IndexOf("DateReceive"));

                    var dicRow = new Dictionary<string, int>();
                    foreach (var PODate in PO)
                        foreach (var _ProductOrder in PODate.ListProductOrder)
                            foreach (var _CustomerOrder in _ProductOrder.ListCustomerOrder)
                            {
                                var _Product = dicProduct[_ProductOrder.ProductId];
                                var _Customer = dicCustomer[_CustomerOrder.CustomerId];

                                DataRow dr = null;

                                var _rowIndex = 0;
                                var sKey = _Product.ProductCode + _Customer.CustomerType +
                                           _Customer.CustomerBigRegion + _Customer.Company +
                                           PODate.DateOrder.Date.ToString("yyyyMMdd");

                                if (!dicRow.TryGetValue(sKey, out _rowIndex))
                                {
                                    dr = dtPO.NewRow();
                                    dicRow.Add(sKey, dtPO.Rows.Count);

                                    dr["REGION"] = string.Join(string.Empty,
                                        _Customer.CustomerBigRegion.Split(' ').Select(x => x.First())).ToUpper();
                                    dr["PCODE"] = _Product.ProductCode;
                                    dr["Note"] = _Product.ProductCode.Substring(0, 1) == "K"
                                        ? "Ok"
                                        : _Product.ProductNote.Contains(_Customer.CustomerBigRegion == "Miền Nam"
                                            ? "South"
                                            : "North")
                                            ? "Ok"
                                            : "Out of List";
                                    dr["CTYPE"] = _Customer.CustomerType;
                                    dr["P&L"] = _Customer.Company;
                                    dr["DateOrder"] = PODate.DateOrder.Date;
                                    dr["QuantityOrder"] = _CustomerOrder.QuantityOrder;
                                    dr["DateReceive"] = _Customer.CustomerBigRegion == "Miền Nam"
                                        ? PODate.DateOrder.Date.AddDays(1)
                                        : PODate.DateOrder.Date;

                                    dtPO.Rows.Add(dr);
                                }
                                else
                                {
                                    dr = dtPO.Rows[_rowIndex];
                                    dr["QuantityOrder"] = (double)dr["QuantityOrder"] + _CustomerOrder.QuantityOrder;
                                }
                            }
                }

                //Excel.Application xlApp = new Excel.Application();
                //Aspose.Cells.Workbook xlWb = new Aspose.Cells.Workbook();

                var fileName =
                    $"PO {DateFrom.ToString("dd.MM") + " - " + DateTo.ToString("dd.MM") + " (" + DateTime.Now.ToString("yyyyMMdd HH\\hmm") + ")"} - {Choice}.xlsx";

                var fileNameXlsb =
                    $"PO {DateFrom.ToString("dd.MM") + " - " + DateTo.ToString("dd.MM") + " (" + DateTime.Now.ToString("yyyyMMdd HH\\hmm") + ")"} - {Choice}.xlsb";

                var path = @"D:\Documents\Stuff\VinEco\Mastah Project\Test\";
                WriteToRichTextBoxOutput(path);
                //var missing = Type.Missing;
                //xlWb.SaveAs(path, Excel.XlFileFormat.xlOpenXMLWorkbook, missing, missing, false, false, Excel.XlSaveAsAccessMode.xlNoChange, missing, missing, missing);

                // Winner in term of Pure speed.
                //LargeExport(dtPO, path + fileName, DicColDate, true, true);

                //ConvertToXlsbInterop(path + fileName, ".xlsx", ".xlsb", true);

                #region Epplus Approach - Just, lol. Doesn't work with .xls ( including .xlsb )

                //FileInfo fileInfo = new FileInfo(path);
                //using (ExcelPackage pck = new ExcelPackage(fileInfo))
                //{
                //    var ws = pck.Workbook.Worksheets.Add("PO");
                //    ws.Cells["A1"].LoadFromDataTable(dtPO, true);
                //    pck.Save();
                //}

                #endregion

                #region ClosedXML Approach - Failed coz Out of Memory Exception

                //using (ClosedXML.Excel.XLWorkbook xlWb = new ClosedXML.Excel.XLWorkbook(ClosedXML.Excel.XLEventTracking.Disabled))
                //{
                //    xlWb.Worksheets.Add("PO");

                //    string[] dtHeaders = new string[dtPO.Columns.Count];

                //    for (int colIndex = 0; colIndex < dtPO.Columns.Count; colIndex++)
                //    {
                //        DateTime dateValue;
                //        string columnName = dtPO.Columns[colIndex].ColumnName;
                //        if (DateTime.TryParse(columnName, out dateValue))
                //        {
                //            dtHeaders[colIndex] = dateValue.Date.ToString();
                //        }
                //        else
                //        {
                //            dtHeaders[colIndex] = columnName;
                //        }
                //    }

                //    xlWb.Worksheet("PO").Cell(1, 1).InsertData(dtHeaders);
                //    xlWb.Worksheet("PO").Cell(2, 1).InsertData(dtPO.AsEnumerable());

                //    xlWb.SaveAs(path);
                //}

                #endregion

                #region Aspose.Cells Approach - Failed coz Out of Memory Exception

                using (var xlWb = new Workbook())
                {
                    // Optimize for Performance
                    xlWb.Settings.MemorySetting = MemorySetting.MemoryPreference;

                    //OutputExcel(POTable, "Sheet1", xlWb, true, 1);
                    OutputExcelAspose(dtPO, "Sheet1", xlWb, true, 1, "A1", DicColDate, "dd-MMM");

                    xlWb.CalculateFormula();
                    xlWb.Save(path + fileNameXlsb, SaveFormat.Xlsb);
                    //xlWb.Close(SaveChanges: true);

                    //GC.Collect();
                    //GC.WaitForPendingFinalizers();

                    //if (xlWb != null) { Marshal.ReleaseComObject(xlWb); }

                    //xlApp.Quit();
                    //if (xlApp != null) { Marshal.ReleaseComObject(xlApp); }
                    //xlApp = null;

                    ////GC.Collect();
                    ////GC.WaitForPendingFinalizers();
                }
                Delete_Evaluation_Sheet_Interop(path + fileNameXlsb);

                #endregion

                stopwatch.Stop();
                WriteToRichTextBoxOutput($"Done in {Math.Round(stopwatch.Elapsed.TotalSeconds, 2)}s!");
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        ///     The Main Character.
        ///     Never die
        ///     Ever shine.
        /// </summary>
        private void FiteMoi(DateTime DateFrom, DateTime DateTo, bool YesNoCompact = false, bool YesNoNoSup = false,
            bool YesNoLimit = false, bool YesNoGroupFarm = true, bool YesNoGroupThuMua = true,
            bool YesNoReportM1 = false, bool YesNoByUnit = false, bool YesNoOnlyFarm = false)
        {
            try
            {
                var stopWatch = Stopwatch.StartNew();

                #region Preparing!

                #region Initializing

                //string connectionString = System.Configuration.ConfigurationManager.ConnectionStrings["mongodb_vecrops.salesms"].ConnectionString;
                //MongoClient mongoClient = new MongoClient(connectionString);
                var mongoClient = new MongoClient();
                var db = mongoClient.GetDatabase("localtest");
                var coreStructure = new CoordStructure();

                // Need to find out how to query this shit.
                // Coz reading the entire fucking thing THEN query is not good. At all.
                // Solved using .Find
                // Also saved a lot of memory
                //var PO = db.GetCollection<PurchaseOrderDate>("PurchaseOrder").AsQueryable().ToList()
                //    .Where(x =>
                //       (x.DateOrder >= DateFrom.Date) &&
                //       (x.DateOrder <= DateTo.Date))
                //    .OrderBy(x => x.DateOrder);
                ////.ToList();

                var PO = db.GetCollection<PurchaseOrderDate>("PurchaseOrder").Find(x =>
                        x.DateOrder >= DateFrom.Date &&
                        x.DateOrder <= DateTo.Date)
                    .ToList();

                //var FC = db.GetCollection<ForecastDate>("Forecast").AsQueryable().ToList()
                //    .Where(x =>
                //        (x.DateForecast.Date >= DateFrom.Date) &&
                //        (x.DateForecast.Date <= DateTo.Date))
                //    .OrderBy(x => x.DateForecast);

                var FC = db.GetCollection<ForecastDate>("Forecast").Find(x =>
                        x.DateForecast >= DateFrom.Date &&
                        x.DateForecast <= DateTo.Date)
                    .ToList().ToArray();

                var Product = db.GetCollection<Product>("Product").AsQueryable().ToList();
                var Supplier = db.GetCollection<Supplier>("Supplier").AsQueryable().ToList();
                var Customer = db.GetCollection<Customer>("Customer").AsQueryable().ToList();

                coreStructure.dicProductUnit = db.GetCollection<ProductUnit>("ProductUnit").AsQueryable()
                    .ToDictionary(x => x.ProductCode);

                coreStructure.dicProductRate = db.GetCollection<ProductRate>("ProductRate").AsQueryable()
                    .ToDictionary(x => x.ProductCode);

                coreStructure.dicProduct = new Dictionary<Guid, Product>(Product.Count() * 4);
                coreStructure.dicSupplier = new Dictionary<Guid, Supplier>(Supplier.Count() * 4);
                coreStructure.dicCustomer = new Dictionary<Guid, Customer>(Customer.Count() * 4);

                //var dicClass = new Dictionary<string, string>
                //{
                //    { "A", "Rau ăn lá" },
                //    { "B", "Rau ăn thân hoa" },
                //    { "C", "Rau ăn quả " },
                //    { "D", "Rau ăn củ" },
                //    { "E", "Cây ăn hạt" },
                //    { "F", "Rau gia vị " },
                //    { "G", "Thủy canh" },
                //    { "H", "Rau mầm " },
                //    { "I", "Nấm" },
                //    { "J", "Lá " },
                //    { "K", "Trái cây (Quả)" },
                //    { "L", "Gạo" },
                //    { "M", "Cỏ và cây công trình" },
                //    { "N", "Hoa" },
                //    { "O", "Dược liệu" }
                //};

                #endregion

                #region Product

                foreach (var _Product in Product)
                {
                    //string _ProductClass = "";
                    //if (dicClass.TryGetValue(_Product.ProductCode.Substring(0, 1), out _ProductClass))
                    //{
                    //    _Product.ProductClassification = _ProductClass;
                    //}

                    // Remove all Special ( non-ASCII ) characters.
                    //_Product.ProductName = Regex.Replace(_Product.ProductName, @"[^\u0000-\u007F]+", string.Empty);
                    _Product.ProductName = ConvertToUnsigned(_Product.ProductName);
                    coreStructure.dicProduct.Add(_Product.ProductId, _Product);
                }

                coreStructure.dicProductCrossRegion = db.GetCollection<ProductCrossRegion>("ProductCrossRegion")
                    .AsQueryable().ToDictionary(x => x.ProductId);

                #endregion

                #region Supplier

                foreach (var _Supplier in Supplier)
                {
                    //_Supplier.SupplierName = Regex.Replace(_Supplier.SupplierName, @"[^\u0000-\u007F]+", string.Empty);
                    _Supplier.SupplierName = ConvertToUnsigned(_Supplier.SupplierName);
                    if (!coreStructure.dicSupplier.ContainsKey(_Supplier.SupplierId))
                        coreStructure.dicSupplier.Add(_Supplier.SupplierId, _Supplier);
                }

                #endregion Supplier

                #region Customer

                foreach (var customer in Customer)
                {
                    Customer _customer = null;
                    if (!coreStructure.dicCustomer.TryGetValue(customer.CustomerId, out _customer))
                        coreStructure.dicCustomer.Add(customer.CustomerId, customer);
                }

                #endregion

                #region PO

                if (NoFruit) FruitOnly = false;

                // Everything related to PO.
                //int maxCalculation = 0;
                coreStructure.dicPO =
                    new Dictionary<DateTime, Dictionary<Product, Dictionary<CustomerOrder, bool>>>();
                var dicUnit = new Dictionary<string, string>(StringComparer.Ordinal);
                foreach (var PODate in PO.OrderByDescending(x => x.DateOrder.Date).Reverse())
                {
                    // Extra function - Only allocate Fruits and uhm, a potato.
                    // I'm sirius. Rlly.
                    if (FruitOnly)
                        PODate.ListProductOrder.RemoveAll(x =>
                            coreStructure.dicProduct[x.ProductId].ProductCode.Substring(0, 1) != "K" &&
                            coreStructure.dicProduct[x.ProductId].ProductCode != "D01401");

                    if (NoFruit)
                        PODate.ListProductOrder.RemoveAll(x =>
                            coreStructure.dicProduct[x.ProductId].ProductCode.Substring(0, 1) == "K");

                    if (PODate.ListProductOrder.Count == 0)
                    {
                        PO.Remove(PODate);
                        continue;
                    }

                    // Core
                    coreStructure.dicPO.Add(PODate.DateOrder.Date,
                        new Dictionary<Product, Dictionary<CustomerOrder, bool>>(Product.Count() * 4));

                    foreach (var _ProductOrder in PODate.ListProductOrder.Reverse<ProductOrder>())
                    {
                        coreStructure.dicPO[PODate.DateOrder.Date].Add(
                            coreStructure.dicProduct[_ProductOrder.ProductId],
                            new Dictionary<CustomerOrder, bool>(_ProductOrder.ListCustomerOrder.Count() * 4));

                        foreach (var _CustomerOrder in _ProductOrder.ListCustomerOrder)
                        {
                            // Handling Unit.

                            // Round to the nearest 2nd decimal digit.
                            _CustomerOrder.QuantityOrder = Math.Round(_CustomerOrder.QuantityOrder, 2);

                            // Proper Type name.
                            var _OrderUnitType =
                                ProperUnit(_CustomerOrder.Unit.ToLower(), dicUnit); // Optimization Purposes.

                            // Converting to Kg.
                            // Only applicable to VM+. For now.
                            _CustomerOrder.Unit = _OrderUnitType;
                            if (coreStructure.dicCustomer[_CustomerOrder.CustomerId].CustomerType == "VM+" &&
                                _OrderUnitType != "Kg")
                            {
                                var _ProductCode = coreStructure.dicProduct[_ProductOrder.ProductId].ProductCode;
                                if (coreStructure.dicProductUnit.TryGetValue(_ProductCode, out var _ProductUnit))
                                {
                                    var _ProductUnitRegion =
                                        _ProductUnit.ListRegion.FirstOrDefault(x => x.OrderUnitType == _OrderUnitType);

                                    if (_ProductUnitRegion != null)
                                    {
                                        _CustomerOrder.Unit = _OrderUnitType;
                                        _CustomerOrder.QuantityOrderKg =
                                            _CustomerOrder.QuantityOrder * _ProductUnitRegion.OrderUnitPer;
                                    }
                                }
                                else
                                {
                                    _CustomerOrder.Unit = "Kg";
                                    _CustomerOrder.QuantityOrderKg = _CustomerOrder.QuantityOrder;
                                }
                            }
                            else
                            {
                                _CustomerOrder.Unit = "Kg";
                                _CustomerOrder.QuantityOrderKg = _CustomerOrder.QuantityOrder;
                            }

                            if (_CustomerOrder.QuantityOrderKg >= 0.1)
                                coreStructure.dicPO[PODate.DateOrder.Date][
                                    coreStructure.dicProduct[_ProductOrder.ProductId]].Add(_CustomerOrder, true);

                            //maxCalculation++;
                        }
                    }
                }

                #endregion

                #region FC

                coreStructure.dicFC =
                    new Dictionary<DateTime, Dictionary<Product, Dictionary<SupplierForecast, bool>>>(FC.Count() * 4);
                foreach (var FCDate in FC.OrderBy(x => x.DateForecast.Date))
                {
                    coreStructure.dicFC.Add(FCDate.DateForecast.Date,
                        new Dictionary<Product, Dictionary<SupplierForecast, bool>>(Product.Count() * 4));
                    foreach (var _ProductForecast in FCDate.ListProductForecast)
                    {
                        coreStructure.dicFC[FCDate.DateForecast.Date].Add(
                            coreStructure.dicProduct[_ProductForecast.ProductId],
                            new Dictionary<SupplierForecast, bool>(FCDate.ListProductForecast.Count() * 4));
                        // To allow user to store their plans on the Forecast
                        // Added a filter on 0 forecast.
                        foreach (var _SupplierForecast in _ProductForecast.ListSupplierForecast
                            .Where(x => x.QualityControlPass /*&& x.QuantityForecast > 0*/).OrderBy(x => x.Level))
                            coreStructure.dicFC[FCDate.DateForecast.Date][
                                coreStructure.dicProduct[_ProductForecast.ProductId]].Add(_SupplierForecast, true);
                    }
                }

                #endregion

                #region Best of both worlds.

                coreStructure.dicCoord =
                    new Dictionary<DateTime, Dictionary<Product,
                        Dictionary<CustomerOrder, Dictionary<SupplierForecast, DateTime>>>>();
                foreach (var PODate in PO)
                {
                    coreStructure.dicCoord.Add(PODate.DateOrder.Date,
                        new Dictionary<Product, Dictionary<CustomerOrder, Dictionary<SupplierForecast, DateTime>>>(
                            Product.Count() * 4));
                    foreach (var _ProductOrder in PODate.ListProductOrder)
                    {
                        coreStructure.dicCoord[PODate.DateOrder.Date].Add(
                            coreStructure.dicProduct[_ProductOrder.ProductId],
                            new Dictionary<CustomerOrder, Dictionary<SupplierForecast, DateTime>>(
                                PODate.ListProductOrder.Count() * 4));
                        foreach (var _CustomerOrder in _ProductOrder.ListCustomerOrder)
                        {
                            _CustomerOrder.QuantityOrder = Math.Round(_CustomerOrder.QuantityOrder, 1);
                            coreStructure.dicCoord[PODate.DateOrder.Date][
                                coreStructure.dicProduct[_ProductOrder.ProductId]].Add(_CustomerOrder, null);
                        }
                    }
                }

                #endregion

                #region VE Farm Table

                var dtVeFarm = new DataTable();

                dtVeFarm.Columns.Add("Region", typeof(string));
                dtVeFarm.Columns.Add("SCODE", typeof(string));
                dtVeFarm.Columns.Add("SNAME", typeof(string));
                dtVeFarm.Columns.Add("PCLASS", typeof(string));
                dtVeFarm.Columns.Add("VECrops Code", typeof(string));
                dtVeFarm.Columns.Add("PCODE", typeof(string));
                dtVeFarm.Columns.Add("PNAME", typeof(string));

                foreach (var DateFC in coreStructure.dicFC.Keys.OrderBy(x => x.Date))
                {
                    var _colName = DateFC.Date.ToString();
                    dtVeFarm.Columns.Add(_colName, typeof(double)).DefaultValue = 0;
                }

                var dicRow = new Dictionary<string, int>();
                foreach (var DateFC in coreStructure.dicFC.Keys.OrderBy(x => x.Date))
                {
                    var _colName = DateFC.Date.ToString();
                    foreach (var _Product in coreStructure.dicFC[DateFC].Keys.OrderBy(x => x.ProductCode))
                    {
                        var _listSupplierForecast = coreStructure.dicFC[DateFC][_Product].Keys.Where(x =>
                            coreStructure.dicSupplier[x.SupplierId].SupplierType == "VinEco");
                        if (_listSupplierForecast != null)
                            foreach (var _SupplierForecast in _listSupplierForecast.OrderBy(x =>
                                coreStructure.dicSupplier[x.SupplierId].SupplierName))
                            {
                                var _Supplier = coreStructure.dicSupplier[_SupplierForecast.SupplierId];
                                var sKey = string.Format("{0}{1}", _Product.ProductCode, _Supplier.SupplierCode);

                                DataRow dr = null;

                                var _rowIndex = 0;
                                if (!dicRow.TryGetValue(sKey, out _rowIndex))
                                {
                                    _rowIndex = dtVeFarm.Rows.Count;
                                    dicRow.Add(sKey, _rowIndex);
                                    dr = dtVeFarm.NewRow();
                                    dtVeFarm.Rows.Add(dr);
                                    dr = dtVeFarm.Rows[_rowIndex];
                                }
                                else
                                {
                                    dr = dtVeFarm.Rows[_rowIndex];
                                }

                                dr["Region"] = _Supplier.SupplierRegion;
                                dr["SCODE"] = _Supplier.SupplierCode;
                                dr["SNAME"] = _Supplier.SupplierName;
                                dr["PCLASS"] = _Product.ProductClassification;
                                dr["VECrops Code"] = _Product.ProductVECode;
                                dr["PCODE"] = _Product.ProductCode;
                                dr["PNAME"] = _Product.ProductName;

                                dr[_colName] = Convert.ToDouble(dr[_colName]) + _SupplierForecast.QuantityForecast;
                            }
                    }
                }

                #endregion

                #region ThuMua Table

                var dtThuMua = new DataTable();

                dtThuMua.Columns.Add("Region", typeof(string));
                dtThuMua.Columns.Add("SCODE", typeof(string));
                dtThuMua.Columns.Add("SNAME", typeof(string));
                dtThuMua.Columns.Add("PNAME", typeof(string));
                dtThuMua.Columns.Add("PCLASS", typeof(string));
                dtThuMua.Columns.Add("PCODE", typeof(string));
                dtThuMua.Columns.Add("Quantity", typeof(double)).DefaultValue = 0;
                dtThuMua.Columns.Add("QC", typeof(string));
                dtThuMua.Columns.Add("Label VE", typeof(string));
                dtThuMua.Columns.Add("100%", typeof(string));
                dtThuMua.Columns.Add("CrossRegion", typeof(string));
                dtThuMua.Columns.Add("Level", typeof(string));
                dtThuMua.Columns.Add("Availability", typeof(string));

                foreach (var DateFC in coreStructure.dicFC.Keys.OrderBy(x => x.Date))
                {
                    var _colName = DateFC.Date.ToString();
                    dtThuMua.Columns.Add(_colName, typeof(double)).DefaultValue = 0;
                }

                dicRow = new Dictionary<string, int>();
                foreach (var DateFC in coreStructure.dicFC.Keys.OrderBy(x => x.Date))
                {
                    var _colName = DateFC.Date.ToString();
                    foreach (var _Product in coreStructure.dicFC[DateFC].Keys.OrderBy(x => x.ProductCode))
                    {
                        var _listSupplierForecast = coreStructure.dicFC[DateFC][_Product].Keys.Where(x =>
                            coreStructure.dicSupplier[x.SupplierId].SupplierType != "VinEco");
                        if (_listSupplierForecast != null)
                            foreach (var _SupplierForecast in _listSupplierForecast.OrderBy(x =>
                                coreStructure.dicSupplier[x.SupplierId].SupplierName))
                            {
                                var _Supplier = coreStructure.dicSupplier[_SupplierForecast.SupplierId];
                                var sKey = string.Format("{0}{1}", _Product.ProductCode, _Supplier.SupplierCode);

                                DataRow dr = null;

                                var _rowIndex = 0;
                                if (!dicRow.TryGetValue(sKey, out _rowIndex))
                                {
                                    _rowIndex = dtThuMua.Rows.Count;
                                    dicRow.Add(sKey, _rowIndex);
                                    dr = dtThuMua.NewRow();
                                    dtThuMua.Rows.Add(dr);
                                    dr = dtThuMua.Rows[_rowIndex];
                                }
                                else
                                {
                                    dr = dtThuMua.Rows[_rowIndex];
                                }

                                dr["Region"] = _Supplier.SupplierRegion;
                                dr["SCODE"] = _Supplier.SupplierCode;
                                dr["SNAME"] = _Supplier.SupplierName;
                                dr["PNAME"] = _Product.ProductName;
                                dr["PCLASS"] = _Product.ProductClassification;
                                dr["PCODE"] = _Product.ProductCode;
                                dr["Quantity"] = Convert.ToDouble(dr["Quantity"]) + _SupplierForecast.QuantityForecast;
                                dr["QC"] = _SupplierForecast.QualityControlPass ? "Ok" : "";
                                dr["Label VE"] = _SupplierForecast.LabelVinEco ? "Yes" : "No";
                                dr["100%"] = _SupplierForecast.FullOrder ? "Yes" : "No";
                                dr["CrossRegion"] = _SupplierForecast.CrossRegion ? "Yes" : "No";
                                dr["Level"] = _SupplierForecast.Level;
                                dr["Availability"] = _SupplierForecast.Availability;

                                dr[_colName] = Convert.ToDouble(dr[_colName]) + _SupplierForecast.QuantityForecast;
                            }
                    }
                }

                #endregion

                #region Everyone has a bite!

                // To deal with uhm, "Everybody has a bite."
                // Make a Dictionary storing the quantity ordered from each Suppliers.

                // First, by Harvest date.
                coreStructure.dicDeli =
                    new Dictionary<DateTime, Dictionary<Product, Dictionary<SupplierForecast, double>>>();
                foreach (var DateFC in coreStructure.dicFC.Keys)
                {
                    // Then, by product.
                    coreStructure.dicDeli.Add(DateFC, new Dictionary<Product, Dictionary<SupplierForecast, double>>());
                    foreach (var _Product in coreStructure.dicFC[DateFC].Keys)
                    {
                        // And finally, by Suppliers. This would be the fairest.
                        coreStructure.dicDeli[DateFC].Add(_Product, new Dictionary<SupplierForecast, double>());
                        foreach (var _SupplierForecast in coreStructure.dicFC[DateFC][_Product].Keys)
                            // ... And of course, the initialized value is 0.
                            coreStructure.dicDeli[DateFC][_Product].Add(_SupplierForecast, 0);
                    }
                }

                #endregion

                // Minimum Order Quantity - MOQ for short.
                // To deal with uhm, OrderQuantity of like, 3 grams. 
                // Who the fuck order 3 grams, seriously.
                coreStructure.dicMinimum = new Dictionary<string, double>
                {
                    {"A", 0.5}, // Rau ăn lá
                    {"B", 0.5},
                    {"C", 0.5}, // Rau ăn quả
                    {"D", 0.5}, // Rau ăn củ
                    {"E", 0.5},
                    {"F", 0.1}, // Rau gia vị
                    {"G", 0.5}, // Thủy canh
                    {"H", 0.2},
                    {"I", 0.2},
                    {"J", 0.5},
                    {"K", 0.7}, // Trái cây
                    {"L", 1},
                    {"M", 1},
                    {"N", 69},
                    {"1", 0.01}, // wat the fuck
                    {"2", 0.01} // wat the duck
                };

                coreStructure.dicTransferDays = new Dictionary<string, byte>
                {
                    {"North-North", 1},
                    {"Highland-North", 3},
                    {"Highland-South", 0},
                    {"South-South", 0},
                    {"South-North", 3},
                    {"North-South", 3}
                };

                #endregion

                #region Main Body

                if (!YesNoLimit) UpperCap = -1;

                // In case of no uppercap, to prevent allocating EVERY FUCKING THING INTO ONE TYPE.
                if (UpperCap == -1)
                    foreach (var _Customer in coreStructure.dicCustomer.Values)
                        _Customer.CustomerType = _Customer.CustomerType.Replace("Priority", "").Trim();

                //byte LDtoMB = 3;
                //byte MBtoMB = 1;
                //byte MNtoMN = 0;
                //byte LDtoMN = 0;
                //byte MBtoMN = 3;
                //byte MNtoMB = 3;

                var ListRegion = new object[6, 5]
                {
                    {"Miền Bắc", "Miền Bắc", (byte) 1, (byte) 3, false},
                    {"Miền Nam", "Miền Nam", (byte) 0, (byte) 0, false},
                    {"Lâm Đồng", "Miền Bắc", (byte) 3, (byte) 3, false},
                    {"Lâm Đồng", "Miền Nam", (byte) 0, (byte) 0, false},
                    {"Miền Bắc", "Miền Nam", (byte) 3, (byte) 0, true},
                    {"Miền Nam", "Miền Bắc", (byte) 3, (byte) 3, true}
                };

                WriteToRichTextBoxOutput();
                WriteToRichTextBoxOutput("UpperCap = " + UpperCap);
                WriteToRichTextBoxOutput();

                // P&L Goes here - using KPI first
                if (!YesNoCompact & !YesNoOnlyFarm)
                {
                    CoordCaller(coreStructure, ListRegion, "VinEco", 1, "B2B", YesNoByUnit, false, YesNoKPI: true);
                    CoordCaller(coreStructure, ListRegion, "ThuMua", 1, "B2B", YesNoByUnit, false, true);
                }

                #region VM+ VinEco Priority

                if (!YesNoOnlyFarm)
                    CoordCaller(coreStructure, ListRegion, "VCM", 1, "VM+ VinEco Priority", YesNoByUnit);

                CoordCaller(coreStructure, ListRegion, "VinEco", UpperCap, "VM+ VinEco Priority", YesNoByUnit, false, true);

                if (!YesNoOnlyFarm)
                {
                    CoordCaller(coreStructure, ListRegion, "ThuMua", UpperCap, "VM+ VinEco Priority", YesNoByUnit, false, true);
                    CoordCaller(coreStructure, ListRegion, "ThuMua", UpperCap, "VM+ VinEco Priority", YesNoByUnit, YesNoContracted: true);
                }

                CoordCaller(coreStructure, ListRegion, "VinEco", UpperCap, "VM+ VinEco Priority", YesNoByUnit);
                if (!YesNoOnlyFarm)
                    CoordCaller(coreStructure, ListRegion, "ThuMua", UpperCap, "VM+ VinEco Priority", YesNoByUnit);

                #endregion

                if (!YesNoOnlyFarm)
                    CoordCaller(coreStructure, ListRegion, "VCM", 1, "VM+ VinEco", YesNoByUnit);

                CoordCaller(coreStructure, ListRegion, "VinEco", UpperCap, "VM+ VinEco", YesNoByUnit, false, true);

                if (!YesNoOnlyFarm)
                {
                    CoordCaller(coreStructure, ListRegion, "ThuMua", UpperCap, "VM+ VinEco", YesNoByUnit, false, true);
                    CoordCaller(coreStructure, ListRegion, "ThuMua", UpperCap, "VM+ VinEco", YesNoByUnit, YesNoContracted: true);
                }

                CoordCaller(coreStructure, ListRegion, "VinEco", UpperCap, "VM+ VinEco", YesNoByUnit);
                if (!YesNoOnlyFarm)
                    CoordCaller(coreStructure, ListRegion, "ThuMua", UpperCap, "VM+ VinEco", YesNoByUnit);
                
                // In any cases, VCM's Vendors go first.
                // Targetting their Ship-to, with their specific Vendors' target defined in FC.
                // VinMart Plus > VinMart.
                if (!YesNoOnlyFarm)
                {
                    CoordCaller(coreStructure, ListRegion, "VCM", UpperCap, "VM+ Priority", YesNoByUnit);
                    CoordCaller(coreStructure, ListRegion, "VCM", UpperCap, "VM+", YesNoByUnit);
                    CoordCaller(coreStructure, ListRegion, "VCM", UpperCap, "VM Priority", YesNoByUnit);
                    CoordCaller(coreStructure, ListRegion, "VCM", UpperCap, "VM", YesNoByUnit);
                }

                #region KPI - Confirmed PO

                // Allocate using confirmed amounts first.

                CoordCaller(coreStructure, ListRegion, "VinEco", UpperCap, "VM+ Priority", YesNoByUnit, false, true);
                CoordCaller(coreStructure, ListRegion, "VinEco", UpperCap, "VM+", YesNoByUnit, false, true);
                CoordCaller(coreStructure, ListRegion, "VinEco", UpperCap, "VM Priority", false, false, true);
                CoordCaller(coreStructure, ListRegion, "VinEco", UpperCap, "VM", false, false, true);

                //int UpperCap = 1;
                if (!YesNoOnlyFarm)
                {
                    // KPI first
                    CoordCaller(coreStructure, ListRegion, "ThuMua", UpperCap, "VM+ Priority", YesNoByUnit, false,
                        true);
                    CoordCaller(coreStructure, ListRegion, "ThuMua", UpperCap, "VM+", YesNoByUnit, false, true);
                    CoordCaller(coreStructure, ListRegion, "ThuMua", UpperCap, "VM Priority", false, false, true);
                    CoordCaller(coreStructure, ListRegion, "ThuMua", UpperCap, "VM", false, false, true);

                    // Contracted afterward
                    CoordCaller(coreStructure, ListRegion, "ThuMua", UpperCap, "VM+ Priority", YesNoByUnit, true);
                    CoordCaller(coreStructure, ListRegion, "ThuMua", UpperCap, "VM+", YesNoByUnit, true);
                    CoordCaller(coreStructure, ListRegion, "ThuMua", UpperCap, "VM Priority", false, true);
                    CoordCaller(coreStructure, ListRegion, "ThuMua", UpperCap, "VM", false, true);
                }

                #endregion

                // P&L goes here
                if (!YesNoCompact & !YesNoOnlyFarm)
                {
                    CoordCaller(coreStructure, ListRegion, "VinEco", 1, "B2B", YesNoByUnit);
                    CoordCaller(coreStructure, ListRegion, "ThuMua", 1, "B2B", YesNoByUnit);
                }


                CoordCaller(coreStructure, ListRegion, "VinEco", UpperCap, "VM+ Priority", YesNoByUnit);
                CoordCaller(coreStructure, ListRegion, "VinEco", UpperCap, "VM+", YesNoByUnit);
                CoordCaller(coreStructure, ListRegion, "VinEco", UpperCap, "VM Priority");
                CoordCaller(coreStructure, ListRegion, "VinEco", UpperCap, "VM", false);

                if (!YesNoOnlyFarm)
                {
                    CoordCaller(coreStructure, ListRegion, "ThuMua", UpperCap, "VM+ Priority", YesNoByUnit);
                    CoordCaller(coreStructure, ListRegion, "ThuMua", UpperCap, "VM+", YesNoByUnit);
                    CoordCaller(coreStructure, ListRegion, "ThuMua", UpperCap, "VM Priority");
                    CoordCaller(coreStructure, ListRegion, "ThuMua", UpperCap, "VM", false);
                }

                CoordCaller(coreStructure, ListRegion, "VCM", UpperCap, "");
                CoordCaller(coreStructure, ListRegion, "VinEco", UpperCap, "");

                if (!YesNoOnlyFarm)
                    CoordCaller(coreStructure, ListRegion, "ThuMua", UpperCap, "");

                #endregion

                #region Output to MongoDb

                //var CoordResultList = new List<CoordResult>();
                //foreach (DateTime DatePO in coreStructure.dicCoord.Keys)
                //{
                //    var _CoordResult = new CoordResult();
                //    _CoordResult._id = Guid.NewGuid();
                //    _CoordResult.CoordResultId = _CoordResult._id;

                //    _CoordResult.DateOrder = DatePO.Date;

                //    _CoordResult.ListCoordResultDate = new List<CoordResultDate>();

                //    foreach (Product _Product in coreStructure.dicCoord[DatePO].Keys)
                //    {
                //        var _CoordResultDate = new CoordResultDate();
                //        _CoordResultDate._id = Guid.NewGuid();
                //        _CoordResultDate.CoordResultDateId = _CoordResult._id;
                //        _CoordResultDate.ProductId = _Product.ProductId;

                //        var _ListCoordinateDate = new List<CoordinateDate>();

                //        foreach (CustomerOrder _CustomerOrder in coreStructure.dicCoord[DatePO][_Product].Keys)
                //        {
                //            if (coreStructure.dicCoord[DatePO][_Product][_CustomerOrder] == null)
                //            {
                //                var _CoordinateDate = new CoordinateDate();
                //                _CoordinateDate._id = Guid.NewGuid();
                //                _CoordinateDate.CoordinateDateId = _CoordinateDate._id;
                //                _CoordinateDate.CustomerOrderId = _CustomerOrder.CustomerOrderId;
                //                _CoordinateDate.SupplierOrderId = null;
                //                _CoordinateDate.DateDelier = null;

                //                _ListCoordinateDate.Add(_CoordinateDate);
                //            }
                //            else
                //            {
                //                foreach (SupplierForecast _SupplierForecast in coreStructure.dicCoord[DatePO][_Product][_CustomerOrder].Keys)
                //                {
                //                    var _CoordinateDate = new CoordinateDate();
                //                    _CoordinateDate._id = Guid.NewGuid();
                //                    _CoordinateDate.CoordinateDateId = _CoordinateDate._id;
                //                    _CoordinateDate.CustomerOrderId = _CustomerOrder.CustomerOrderId;
                //                    _CoordinateDate.SupplierOrderId = _SupplierForecast.SupplierForecastId;
                //                    _CoordinateDate.DateDelier = coreStructure.dicCoord[DatePO][_Product][_CustomerOrder][_SupplierForecast].Date;

                //                    _ListCoordinateDate.Add(_CoordinateDate);
                //                }
                //            }
                //        }

                //        _CoordResultDate.ListCoordinateDate = _ListCoordinateDate;

                //        _CoordResult.ListCoordResultDate.Add(_CoordResultDate);
                //    }

                //    CoordResultList.Add(_CoordResult);
                //}

                #endregion

                #region to DataTable

                // Dealing with Prioritized Customers.
                // Well, more like hiding what I have done >:D

                foreach (var VmpVinEco in coreStructure.dicCustomer.Values.Where(x => x.CustomerType == "VM+ VinEco"))
                    foreach (var SiteToChange in coreStructure.dicCustomer.Values.Where(x =>
                        x.CustomerCode == VmpVinEco.CustomerCode))
                        SiteToChange.CustomerType = VmpVinEco.CustomerType;

                foreach (var _Customer in coreStructure.dicCustomer.Values)
                    if (_Customer.CustomerType == "B2B")
                        _Customer.CustomerType = _Customer.Company;
                    else
                        _Customer.CustomerType = _Customer.CustomerType.Replace("Priority", "").Trim();

                // Dealing with stubborn Procuring Forcasts.
                foreach (var DateFC in coreStructure.dicFC.Keys.OrderBy(x => x.Date))
                    foreach (var _Product in coreStructure.dicFC[DateFC].Keys.OrderBy(x => x.ProductCode))
                    {
                        var _ListSupplier = coreStructure.dicFC[DateFC][_Product].Keys.Where(x =>
                            coreStructure.dicSupplier[x.SupplierId].SupplierType != "VinEco");
                        if (_ListSupplier != null)
                            foreach (var _SupplierForecast in _ListSupplier.Reverse())
                                if (_SupplierForecast.QuantityForecastPlanned == null)
                                    coreStructure.dicFC[DateFC][_Product].Remove(_SupplierForecast);
                    }

                var _dateBase = new DateTime(1900, 1, 1);

                if (YesNoReportM1)
                {
                    #region Report M1

                    #region Mastah Table - Report M+1

                    var dtMastah = new DataTable("ReportM1");

                    dtMastah.Columns.Add("Mã VinEco", typeof(string)).DefaultValue = "";
                    dtMastah.Columns.Add("Tên VinEco", typeof(string)).DefaultValue = "";
                    dtMastah.Columns.Add("Loại cửa hàng", typeof(string)).DefaultValue = "";
                    dtMastah.Columns.Add("Ngày tiêu thụ", typeof(int));
                    dtMastah.Columns.Add("Vùng tiêu thụ", typeof(string)).DefaultValue = "";
                    dtMastah.Columns.Add("Tỉnh tiêu thụ", typeof(string)).DefaultValue = "";
                    dtMastah.Columns.Add("NoSup", typeof(string)).DefaultValue = "No";
                    dtMastah.Columns.Add("Class", typeof(string)).DefaultValue = "";
                    dtMastah.Columns.Add("DS1", typeof(string)).DefaultValue = "";
                    dtMastah.Columns.Add("Region", typeof(string)).DefaultValue = "";
                    dtMastah.Columns.Add("Bắt buộc?", typeof(string)).DefaultValue = "";
                    dtMastah.Columns.Add("VCM", typeof(double)).DefaultValue = 0;
                    dtMastah.Columns.Add("VE", typeof(double)).DefaultValue = 0;
                    dtMastah.Columns.Add("TM", typeof(double)).DefaultValue = 0;
                    dtMastah.Columns.Add("Ngày sơ chế", typeof(double)).DefaultValue = 0;
                    dtMastah.Columns.Add("Vùng sản xuất", typeof(double)).DefaultValue = 0;

                    dicRow = new Dictionary<string, int>();

                    foreach (var DatePO in coreStructure.dicCoord.Keys)
                        foreach (var _Product in coreStructure.dicCoord[DatePO].Keys)
                            foreach (var _CustomerOrder in coreStructure.dicCoord[DatePO][_Product].Keys)
                                if (coreStructure.dicCoord[DatePO][_Product][_CustomerOrder] != null)
                                {
                                    foreach (var _SupplierForecast in coreStructure.dicCoord[DatePO][_Product][_CustomerOrder]
                                        .Keys.OrderBy(x => coreStructure.dicSupplier[x.SupplierId].SupplierType))
                                    {
                                        DataRow dr = null;

                                        var _Customer = coreStructure.dicCustomer[_CustomerOrder.CustomerId];
                                        var _Supplier = coreStructure.dicSupplier[_SupplierForecast.SupplierId];

                                        var sKey = DatePO.Date + _Product.ProductCode + _Customer.CustomerType +
                                                   _Customer.CustomerRegion;
                                        var newRow = false;

                                        var _rowPos = 0;
                                        if (dicRow.TryGetValue(sKey, out _rowPos))
                                        {
                                            dr = dtMastah.Rows[_rowPos];
                                        }
                                        else
                                        {
                                            dr = dtMastah.NewRow();
                                            _rowPos = dtMastah.Rows.Count;
                                            dicRow.Add(sKey, _rowPos);
                                            newRow = true;
                                        }

                                        dr["Mã VinEco"] = _Product.ProductCode;
                                        dr["Tên VinEco"] = _Product.ProductName;
                                        dr["Loại cửa hàng"] = _Customer.CustomerType;
                                        dr["Ngày tiêu thụ"] = (int)(DatePO.Date - _dateBase).TotalDays + 2;
                                        dr["Vùng tiêu thụ"] = _Customer.CustomerBigRegion;
                                        dr["Tỉnh tiêu thụ"] = _Customer.CustomerRegion;

                                        //dr["NoSup"] = "";

                                        //string productClass = "";
                                        //if (!dicClass.TryGetValue(_Product.ProductCode.Substring(0, 1), out productClass))
                                        //{
                                        //    productClass = "???";
                                        //}
                                        dr["Class"] = _Product.ProductClassification;

                                        //dr["DS1"] = "";
                                        //dr["Region"] = "";
                                        //dr["Bắt buộc?"] = "";

                                        dr["VCM"] = (double)dr["VCM"] + _CustomerOrder.QuantityOrderKg;
                                        if (_Supplier.SupplierType == "VinEco")
                                            dr["VE"] = (double)dr["VE"] + _SupplierForecast.QuantityForecast;
                                        else if (_Supplier.SupplierType == "ThuMua")
                                            dr["TM"] = (double)dr["TM"] + _SupplierForecast.QuantityForecast;

                                        if (newRow)
                                            dtMastah.Rows.Add(dr);
                                    }
                                }
                                else
                                {
                                    DataRow dr = null;

                                    var _Customer = coreStructure.dicCustomer[_CustomerOrder.CustomerId];

                                    var sKey = DatePO.Date + _Product.ProductCode + _Customer.CustomerType +
                                               _Customer.CustomerRegion;
                                    var newRow = false;

                                    var _rowPos = 0;
                                    if (dicRow.TryGetValue(sKey, out _rowPos))
                                    {
                                        dr = dtMastah.Rows[_rowPos];
                                    }
                                    else
                                    {
                                        dr = dtMastah.NewRow();
                                        _rowPos = dtMastah.Rows.Count;
                                        dicRow.Add(sKey, _rowPos);
                                        newRow = true;
                                    }

                                    dr["Mã VinEco"] = _Product.ProductCode;
                                    dr["Tên VinEco"] = _Product.ProductName;
                                    dr["Loại cửa hàng"] = _Customer.CustomerType;
                                    dr["Ngày tiêu thụ"] = (int)(DatePO.Date - _dateBase).TotalDays + 2;
                                    dr["Vùng tiêu thụ"] = _Customer.CustomerBigRegion;
                                    dr["Tỉnh tiêu thụ"] = _Customer.CustomerRegion;

                                    //dr["NoSup"] = "";

                                    //string productClass = "";
                                    //if (!dicClass.TryGetValue(_Product.ProductCode.Substring(0, 1), out productClass))
                                    //{
                                    //    productClass = "???";
                                    //}
                                    dr["Class"] = _Product.ProductClassification;

                                    //dr["DS1"] = "";
                                    //dr["Region"] = "";
                                    //dr["Bắt buộc?"] = "";

                                    dr["VCM"] = (double)dr["VCM"] + _CustomerOrder.QuantityOrderKg;


                                    if (newRow)
                                        dtMastah.Rows.Add(dr);
                                }

                    foreach (DataRow dr in dtMastah.Rows)
                        if ((double)dr["VCM"] > (double)dr["VE"] + (double)dr["TM"])
                            dr["NoSup"] = "Yes";

                    #endregion

                    #region LeftoverVinEco

                    var dtLeftoverVe = new DataTable();

                    dtLeftoverVe.TableName = "NoCusVinEco";

                    dtLeftoverVe.Columns.Add("Mã VinEco", typeof(string)).DefaultValue = "";
                    dtLeftoverVe.Columns.Add("Tên VinEco", typeof(string)).DefaultValue = "";
                    dtLeftoverVe.Columns.Add("Mã Farm", typeof(string)).DefaultValue = "";
                    dtLeftoverVe.Columns.Add("Tên Farm", typeof(string)).DefaultValue = "";
                    dtLeftoverVe.Columns.Add("Ngày thu hoạch", typeof(int));
                    dtLeftoverVe.Columns.Add("Vùng sản xuất", typeof(string)).DefaultValue = "";
                    dtLeftoverVe.Columns.Add("Sản lượng", typeof(double)).DefaultValue = 0;

                    foreach (var DateFC in coreStructure.dicFC.Keys)
                        foreach (var _Product in coreStructure.dicFC[DateFC].Keys.OrderBy(x => x.ProductCode))
                        {
                            var _ListSupplier = coreStructure.dicFC[DateFC][_Product].Keys.Where(x =>
                                coreStructure.dicSupplier[x.SupplierId].SupplierType == "VinEco");
                            if (_ListSupplier != null)
                                foreach (var _SupplierForecast in _ListSupplier.OrderBy(x =>
                                    coreStructure.dicSupplier[x.SupplierId].SupplierName))
                                    if (_SupplierForecast.QuantityForecast > 0)
                                    {
                                        var dr = dtLeftoverVe.NewRow();

                                        //Customer _Customer =coreStructure. dicCustomer[_CustomerOrder.CustomerId];
                                        var _Supplier = coreStructure.dicSupplier[_SupplierForecast.SupplierId];

                                        dr["Mã VinEco"] = _Product.ProductCode;
                                        dr["Tên VinEco"] = _Product.ProductName;
                                        dr["Mã Farm"] = _Supplier.SupplierCode;
                                        dr["Tên Farm"] = _Supplier.SupplierName;
                                        dr["Ngày thu hoạch"] = (int)(DateFC.Date - _dateBase).TotalDays + 2;
                                        dr["Vùng sản xuất"] = _Supplier.SupplierRegion;
                                        dr["Sản lượng"] = _SupplierForecast.QuantityForecast;

                                        dtLeftoverVe.Rows.Add(dr);
                                    }
                        }

                    #endregion

                    #region LeftoverThuMua

                    var dtLeftoverTm = new DataTable();

                    dtLeftoverTm.TableName = "NoCusThuMua";

                    dtLeftoverTm.Columns.Add("Mã VinEco", typeof(string)).DefaultValue = "";
                    dtLeftoverTm.Columns.Add("Tên VinEco", typeof(string)).DefaultValue = "";
                    dtLeftoverTm.Columns.Add("Mã Farm", typeof(string)).DefaultValue = "";
                    dtLeftoverTm.Columns.Add("Tên Farm", typeof(string)).DefaultValue = "";
                    dtLeftoverTm.Columns.Add("Ngày thu hoạch", typeof(int));
                    dtLeftoverTm.Columns.Add("Vùng sản xuất", typeof(string)).DefaultValue = "";
                    dtLeftoverTm.Columns.Add("Sản lượng", typeof(double)).DefaultValue = 0;

                    foreach (var DateFC in coreStructure.dicFC.Keys)
                        foreach (var _Product in coreStructure.dicFC[DateFC].Keys.OrderBy(x => x.ProductCode))
                        {
                            var _ListSupplier = coreStructure.dicFC[DateFC][_Product].Keys.Where(x =>
                                coreStructure.dicSupplier[x.SupplierId].SupplierType == "ThuMua");
                            if (_ListSupplier != null)
                                foreach (var _SupplierForecast in _ListSupplier.OrderBy(x =>
                                    coreStructure.dicSupplier[x.SupplierId].SupplierName))
                                    if (_SupplierForecast.QuantityForecast > 0)
                                    {
                                        var dr = dtLeftoverTm.NewRow();

                                        //Customer _Customer =coreStructure. dicCustomer[_CustomerOrder.CustomerId];
                                        var _Supplier = coreStructure.dicSupplier[_SupplierForecast.SupplierId];

                                        dr["Mã VinEco"] = _Product.ProductCode;
                                        dr["Tên VinEco"] = _Product.ProductName;
                                        dr["Mã Farm"] = _Supplier.SupplierCode;
                                        dr["Tên Farm"] = _Supplier.SupplierName;
                                        dr["Ngày thu hoạch"] = (int)(DateFC.Date - _dateBase).TotalDays + 2;
                                        dr["Vùng sản xuất"] = _Supplier.SupplierRegion;
                                        dr["Sản lượng"] = _SupplierForecast.QuantityForecast;

                                        dtLeftoverTm.Rows.Add(dr);
                                    }
                        }

                    #endregion

                    #region Output to Excel

                    var dicDateCol = new Dictionary<string, int>();

                    dicDateCol.Add("Ngày tiêu thụ", dtMastah.Columns.IndexOf("Ngày tiêu thụ"));

                    var fileName = string.Format("Report M plus 1 {0}.xlsx",
                        DateFrom.ToString("dd.MM") + " - " + DateTo.ToString("dd.MM") + " (" +
                        DateTime.Now.ToString("yyyyMMdd HH\\hmm") + ")");
                    var path = string.Format(
                        @"D:\Documents\Stuff\VinEco\Mastah Project\Test\" + fileName);

                    var listDt = new List<DataTable> {dtMastah, dtLeftoverVe, dtLeftoverTm};


                    LargeExportOneWorkbook(path, listDt, true, true);

                    ConvertToXlsbInterop(path, "xlsx", "xlsb", true);

                    #endregion

                    #endregion
                }
                else if (YesNoNoSup)
                {
                    #region NoSup

                    #region Mastah Table

                    var dtMastah = new DataTable("NoSup");

                    dtMastah.Columns.Add("Mã VinEco", typeof(string)).DefaultValue = "";
                    dtMastah.Columns.Add("Tên VinEco", typeof(string)).DefaultValue = "";
                    dtMastah.Columns.Add("Mã cửa hàng", typeof(string)).DefaultValue = "";
                    dtMastah.Columns.Add("Loại cửa hàng", typeof(string)).DefaultValue = "";
                    dtMastah.Columns.Add("Ngày tiêu thụ", typeof(int));
                    dtMastah.Columns.Add("Vùng tiêu thụ", typeof(string)).DefaultValue = "";
                    dtMastah.Columns.Add("Tỉnh tiêu thụ", typeof(string)).DefaultValue = "";
                    dtMastah.Columns.Add("NoSup", typeof(string)).DefaultValue = "No";
                    dtMastah.Columns.Add("Class", typeof(string)).DefaultValue = "";
                    dtMastah.Columns.Add("DS1", typeof(string)).DefaultValue = "";
                    dtMastah.Columns.Add("Region", typeof(string)).DefaultValue = "";
                    dtMastah.Columns.Add("Bắt buộc?", typeof(string)).DefaultValue = "";
                    dtMastah.Columns.Add("VCM", typeof(double)).DefaultValue = 0;

                    dicRow = new Dictionary<string, int>();

                    foreach (var DatePO in coreStructure.dicCoord.Keys)
                        foreach (var _Product in coreStructure.dicCoord[DatePO].Keys)
                            foreach (var _CustomerOrder in coreStructure.dicCoord[DatePO][_Product].Keys)
                                if (coreStructure.dicCoord[DatePO][_Product][_CustomerOrder] != null)
                                {
                                    foreach (var _SupplierForecast in coreStructure.dicCoord[DatePO][_Product][_CustomerOrder]
                                        .Keys.OrderBy(x => coreStructure.dicSupplier[x.SupplierId].SupplierType))
                                        if (_SupplierForecast.QuantityForecast < _CustomerOrder.QuantityOrderKg)
                                        {
                                            var dr = dtMastah.NewRow();

                                            var _Customer = coreStructure.dicCustomer[_CustomerOrder.CustomerId];
                                            var _Supplier = coreStructure.dicSupplier[_SupplierForecast.SupplierId];

                                            dr["Mã VinEco"] = _Product.ProductCode;
                                            dr["Tên VinEco"] = _Product.ProductName;
                                            dr["Mã cửa hàng"] = _Customer.CustomerCode;
                                            dr["Loại cửa hàng"] = _Customer.CustomerType;
                                            dr["Ngày tiêu thụ"] = (int)(DatePO.Date - _dateBase).TotalDays + 2;
                                            dr["Vùng tiêu thụ"] = _Customer.CustomerBigRegion;
                                            dr["Tỉnh tiêu thụ"] = _Customer.CustomerRegion;

                                            //dr["NoSup"] = "";

                                            //string productClass = "";
                                            //if (!dicClass.TryGetValue(_Product.ProductCode.Substring(0, 1), out productClass))
                                            //{
                                            //    productClass = "???";
                                            //}
                                            dr["Class"] = _Product.ProductClassification;

                                            //dr["DS1"] = "";
                                            //dr["Region"] = "";
                                            //dr["Bắt buộc?"] = "";

                                            dr["VCM"] = _CustomerOrder.QuantityOrderKg - _SupplierForecast.QuantityForecast;
                                            //switch (_Supplier.SupplierType)
                                            //{
                                            //    case "VinEco": dr["VE"] = _SupplierForecast.QuantityForecast; break;
                                            //    case "ThuMua": dr["TM"] = _SupplierForecast.QuantityForecast; break;
                                            //    default: break;
                                            //}

                                            dtMastah.Rows.Add(dr);
                                        }
                                }
                                else
                                {
                                    var dr = dtMastah.NewRow();

                                    var _Customer = coreStructure.dicCustomer[_CustomerOrder.CustomerId];

                                    dr["Mã VinEco"] = _Product.ProductCode;
                                    dr["Tên VinEco"] = _Product.ProductName;
                                    dr["Mã cửa hàng"] = _Customer.CustomerCode;
                                    dr["Loại cửa hàng"] = _Customer.CustomerType;
                                    dr["Ngày tiêu thụ"] = (int)(DatePO.Date - _dateBase).TotalDays + 2;
                                    dr["Vùng tiêu thụ"] = _Customer.CustomerBigRegion;
                                    dr["Tỉnh tiêu thụ"] = _Customer.CustomerRegion;

                                    //dr["NoSup"] = "";

                                    //string productClass = "";
                                    //if (!dicClass.TryGetValue(_Product.ProductCode.Substring(0, 1), out productClass))
                                    //{
                                    //    productClass = "???";
                                    //}
                                    dr["Class"] = _Product.ProductClassification;

                                    //dr["DS1"] = "";
                                    //dr["Region"] = "";
                                    //dr["Bắt buộc?"] = "";

                                    dr["VCM"] = _CustomerOrder.QuantityOrderKg;

                                    dtMastah.Rows.Add(dr);
                                }

                    var dtNoSup = new DataTable { TableName = "NoSup" };

                    dtNoSup.Columns.Add("Mã VinEco", typeof(string)).DefaultValue = "";
                    dtNoSup.Columns.Add("Tên VinEco", typeof(string)).DefaultValue = "";
                    dtNoSup.Columns.Add("Mã cửa hàng", typeof(string)).DefaultValue = "";
                    dtNoSup.Columns.Add("Loại cửa hàng", typeof(string)).DefaultValue = "";
                    dtNoSup.Columns.Add("Ngày tiêu thụ", typeof(int));
                    dtNoSup.Columns.Add("Vùng tiêu thụ", typeof(string)).DefaultValue = "";
                    dtNoSup.Columns.Add("Tỉnh tiêu thụ", typeof(string)).DefaultValue = "";
                    dtNoSup.Columns.Add("NoSup", typeof(string)).DefaultValue = "No";
                    dtNoSup.Columns.Add("Class", typeof(string)).DefaultValue = "";
                    dtNoSup.Columns.Add("DS1", typeof(string)).DefaultValue = "";
                    dtNoSup.Columns.Add("Region", typeof(string)).DefaultValue = "";
                    dtNoSup.Columns.Add("Bắt buộc?", typeof(string)).DefaultValue = "";
                    dtNoSup.Columns.Add("VCM", typeof(double)).DefaultValue = 0;

                    foreach (DataRow dr in dtMastah.Rows)
                    {
                        //if ((double)dr["VCM"] - (double)dr["VE"] - (double)dr["TM"] > 0)
                        //{
                        //    DataRow drNoSup = dtNoSup.NewRow();

                        //    int _colIndex = 0;
                        //    foreach (DataColumn dc in dtMastah.Columns)
                        //    {
                        //        drNoSup[_colIndex] = dr[_colIndex];
                        //        _colIndex++;
                        //    }

                        //dtNoSup.Rows.Add(drNoSup);
                        //}
                    }

                    #endregion

                    #region LeftoverVinEco

                    var dtLeftoverVe = new DataTable();

                    dtLeftoverVe.TableName = "NoCusVinEco";

                    dtLeftoverVe.Columns.Add("Mã VinEco", typeof(string)).DefaultValue = "";
                    dtLeftoverVe.Columns.Add("Tên VinEco", typeof(string)).DefaultValue = "";
                    dtLeftoverVe.Columns.Add("Mã Farm", typeof(string)).DefaultValue = "";
                    dtLeftoverVe.Columns.Add("Tên Farm", typeof(string)).DefaultValue = "";
                    dtLeftoverVe.Columns.Add("Ngày thu hoạch", typeof(int));
                    dtLeftoverVe.Columns.Add("Vùng sản xuất", typeof(string)).DefaultValue = "";
                    dtLeftoverVe.Columns.Add("Sản lượng", typeof(double)).DefaultValue = 0;

                    foreach (var DateFC in coreStructure.dicFC.Keys)
                        foreach (var _Product in coreStructure.dicFC[DateFC].Keys.OrderBy(x => x.ProductCode))
                        {
                            var _ListSupplier = coreStructure.dicFC[DateFC][_Product].Keys.Where(x =>
                                coreStructure.dicSupplier[x.SupplierId].SupplierType == "VinEco");
                            if (_ListSupplier != null)
                                foreach (var _SupplierForecast in _ListSupplier.OrderBy(x =>
                                    coreStructure.dicSupplier[x.SupplierId].SupplierName))
                                    if (_SupplierForecast.QuantityForecast > 3)
                                    {
                                        var dr = dtLeftoverVe.NewRow();

                                        //Customer _Customer =coreStructure. dicCustomer[_CustomerOrder.CustomerId];
                                        var _Supplier = coreStructure.dicSupplier[_SupplierForecast.SupplierId];

                                        dr["Mã VinEco"] = _Product.ProductCode;
                                        dr["Tên VinEco"] = _Product.ProductName;
                                        dr["Mã Farm"] = _Supplier.SupplierCode;
                                        dr["Tên Farm"] = _Supplier.SupplierName;
                                        dr["Ngày thu hoạch"] = (int)(DateFC.Date - _dateBase).TotalDays + 2;
                                        dr["Vùng sản xuất"] = _Supplier.SupplierRegion;
                                        dr["Sản lượng"] = _SupplierForecast.QuantityForecast;

                                        dtLeftoverVe.Rows.Add(dr);
                                    }
                        }

                    #endregion

                    #region LeftoverThuMua

                    var dtLeftoverTm = new DataTable();

                    dtLeftoverTm.TableName = "NoCusThuMua";

                    dtLeftoverTm.Columns.Add("Mã VinEco", typeof(string)).DefaultValue = "";
                    dtLeftoverTm.Columns.Add("Tên VinEco", typeof(string)).DefaultValue = "";
                    dtLeftoverTm.Columns.Add("Mã Farm", typeof(string)).DefaultValue = "";
                    dtLeftoverTm.Columns.Add("Tên Farm", typeof(string)).DefaultValue = "";
                    dtLeftoverTm.Columns.Add("Ngày thu hoạch", typeof(int));
                    dtLeftoverTm.Columns.Add("Vùng sản xuất", typeof(string)).DefaultValue = "";
                    dtLeftoverTm.Columns.Add("Sản lượng", typeof(double)).DefaultValue = 0;

                    foreach (var DateFC in coreStructure.dicFC.Keys)
                        foreach (var _Product in coreStructure.dicFC[DateFC].Keys.OrderBy(x => x.ProductCode))
                        {
                            var _ListSupplier = coreStructure.dicFC[DateFC][_Product].Keys.Where(x =>
                                coreStructure.dicSupplier[x.SupplierId].SupplierType == "ThuMua");
                            if (_ListSupplier != null)
                                foreach (var _SupplierForecast in _ListSupplier
                                    .Where(x => x.QuantityForecastPlanned != null).OrderBy(x =>
                                        coreStructure.dicSupplier[x.SupplierId].SupplierName))
                                    if (_SupplierForecast.QuantityForecast > 0)
                                    {
                                        var dr = dtLeftoverTm.NewRow();

                                        //Customer _Customer =coreStructure. dicCustomer[_CustomerOrder.CustomerId];
                                        var _Supplier = coreStructure.dicSupplier[_SupplierForecast.SupplierId];

                                        dr["Mã VinEco"] = _Product.ProductCode;
                                        dr["Tên VinEco"] = _Product.ProductName;
                                        dr["Mã Farm"] = _Supplier.SupplierCode;
                                        dr["Tên Farm"] = _Supplier.SupplierName;
                                        dr["Ngày thu hoạch"] = (int)(DateFC.Date - _dateBase).TotalDays + 2;
                                        dr["Vùng sản xuất"] = _Supplier.SupplierRegion;
                                        dr["Sản lượng"] = _SupplierForecast.QuantityForecast;

                                        dtLeftoverTm.Rows.Add(dr);
                                    }
                        }

                    #endregion

                    #region Output to Excel

                    var dicDateCol = new Dictionary<string, int>();

                    dicDateCol.Add("Ngày tiêu thụ", dtMastah.Columns.IndexOf("Ngày tiêu thụ"));

                    var fileName = string.Format("NoSup {0}.xlsx",
                        DateFrom.ToString("dd.MM") + " - " + DateTo.ToString("dd.MM") + " (" +
                        DateTime.Now.ToString("yyyyMMdd HH\\hmm") + ")");
                    var path = string.Format(
                        @"D:\Documents\Stuff\VinEco\Mastah Project\Test\" + fileName);

                    var listDt = new List<DataTable>();

                    listDt.Add(dtMastah);
                    listDt.Add(dtLeftoverVe);
                    listDt.Add(dtLeftoverTm);

                    LargeExportOneWorkbook(path, listDt, true, true);

                    ConvertToXlsbInterop(path, "xlsx", "xlsb", true);

                    #endregion

                    #region Output to Excel

                    ////Excel.Application xlApp = new Excel.Application();

                    ////xlApp.ScreenUpdating = false;
                    ////xlApp.EnableEvents = false;
                    ////xlApp.DisplayAlerts = false;
                    ////xlApp.DisplayStatusBar = false;
                    ////xlApp.AskToUpdateLinks = false;

                    //string filePath = Application.StartupPath.Replace("\\bin\\Debug", "") + "\\Template\\{0}";
                    //string fileFullPath = string.Format(filePath, "NoSup.xlsb");

                    ////Excel.Workbook xlWb = xlApp.Workbooks.Add();

                    ////xlApp.Calculation = Excel.XlCalculation.xlCalculationManual;

                    //var missing = Type.Missing;
                    ////string path = string.Format(@"D:\Documents\Stuff\VinEco\Mastah Project\Test\NoSup {0}.xlsb", DateFrom.ToString("dd.MM") + " - " + DateTo.ToString("dd.MM") + " (" + DateTime.Now.ToString("yyyyMMdd HH\\hmm") + ")");

                    //string fileName = string.Format("NoSup {0}.xlsx", DateFrom.ToString("dd.MM") + " - " + DateTo.ToString("dd.MM") + " (" + DateTime.Now.ToString("yyyyMMdd HH\\hmm") + ")");
                    //string path = string.Format(@"D:\Documents\Stuff\VinEco\Mastah Project\Test\" + fileName);

                    //LargeExport(dtMastah, path, dicDateCol, true, false);

                    //ConvertToXlsbInterop(path, "xlsx", "xlsb", true);

                    ////xlWb.SaveAs(path, Excel.XlFileFormat.xlExcel12, missing, missing, false, false, Excel.XlSaveAsAccessMode.xlNoChange, missing, missing, missing);

                    ////OutputExcel(dtMastah, "Sheet1", xlWb, true, 1);
                    ////dtMastah = null;

                    ////xlApp.Calculation = Excel.XlCalculation.xlCalculationAutomatic;

                    ////xlWb.Close(SaveChanges: true);
                    ////Marshal.ReleaseComObject(xlWb);
                    ////xlWb = null;

                    ////xlApp.ScreenUpdating = true;
                    ////xlApp.EnableEvents = true;
                    ////xlApp.DisplayAlerts = false;
                    ////xlApp.DisplayStatusBar = true;
                    ////xlApp.AskToUpdateLinks = true;

                    ////xlApp.Quit();
                    ////if (xlApp != null) { Marshal.ReleaseComObject(xlApp); }
                    ////xlApp = null;

                    #endregion

                    #endregion
                }
                else
                {
                    if (!YesNoCompact)
                    {
                        #region Normal file

                        #region Mastah Table

                        var dtMastah = new DataTable { TableName = "Mastah" };

                        dtMastah.Columns.Add("Mã 6 ký tự", typeof(string));
                        dtMastah.Columns.Add("Mã thành phẩm VinEco", typeof(string));
                        dtMastah.Columns.Add("P&L", typeof(string));
                        dtMastah.Columns.Add("Tên Sản phẩm", typeof(string));
                        dtMastah.Columns.Add("Mã Cửa hàng", typeof(string));
                        dtMastah.Columns.Add("Tên Cửa hàng", typeof(string));
                        dtMastah.Columns.Add("Loại Cửa hàng", typeof(string));
                        dtMastah.Columns.Add("Ngày Tiêu thụ", typeof(DateTime));
                        dtMastah.Columns.Add("Vùng Tiêu thụ", typeof(string));

                        dtMastah.Columns.Add("Nhu cầu Kg VinCommerce", typeof(double)).DefaultValue = 0;
                        //dtMastah.Columns.Add("Số lượng đặt", typeof(double)).DefaultValue = 0;
                        //dtMastah.Columns.Add("Đơn vị đặt", typeof(string)).DefaultValue = "Kg";
                        //dtMastah.Columns.Add("Đặt Kg/Unit", typeof(double)).DefaultValue = 1;

                        dtMastah.Columns.Add("Nhu cầu Kg Đã đáp ứng", typeof(string));
                        //dtMastah.Columns.Add("Số lượng bán", typeof(string));
                        //dtMastah.Columns.Add("Đơn vị bán", typeof(string)).DefaultValue = "Kg";
                        //dtMastah.Columns.Add("Bán Kg/Unit", typeof(double)).DefaultValue = 1;

                        dtMastah.Columns.Add("Tên VinEco MB", typeof(string));
                        dtMastah.Columns.Add("Đáp ứng từ VinEco MB", typeof(double));
                        dtMastah.Columns.Add("Ngày sơ chế VinEco MB", typeof(DateTime));

                        dtMastah.Columns.Add("Tên VinEco MN", typeof(string));
                        dtMastah.Columns.Add("Đáp ứng từ VinEco MN", typeof(double));
                        dtMastah.Columns.Add("Ngày sơ chế VinEco MN", typeof(DateTime));

                        dtMastah.Columns.Add("Tên VinEco LĐ", typeof(string));
                        dtMastah.Columns.Add("Đáp ứng từ VinEco LĐ", typeof(double));
                        dtMastah.Columns.Add("Ngày sơ chế VinEco LĐ", typeof(DateTime));

                        dtMastah.Columns.Add("Tên ThuMua MB", typeof(string));
                        dtMastah.Columns.Add("Đáp ứng từ ThuMua MB", typeof(double));
                        dtMastah.Columns.Add("Ngày sơ chế ThuMua MB", typeof(DateTime));
                        //dtMastah.Columns.Add("Giá mua ThuMua MB", typeof(double));

                        dtMastah.Columns.Add("Tên ThuMua MN", typeof(string));
                        dtMastah.Columns.Add("Đáp ứng từ ThuMua MN", typeof(double));
                        dtMastah.Columns.Add("Ngày sơ chế ThuMua MN", typeof(DateTime));
                        //dtMastah.Columns.Add("Giá mua ThuMua MN", typeof(double));

                        dtMastah.Columns.Add("Tên ThuMua LĐ", typeof(string));
                        dtMastah.Columns.Add("Đáp ứng từ ThuMua LĐ", typeof(double));
                        dtMastah.Columns.Add("Ngày sơ chế ThuMua LĐ", typeof(DateTime));
                        //dtMastah.Columns.Add("Giá mua ThuMua LĐ", typeof(double));

                        dtMastah.Columns.Add("Note", typeof(string));

                        foreach (var DatePO in coreStructure.dicCoord.Keys.OrderBy(x => x.Date)
                            .Where(x => x.Date >= DateFrom.AddDays(dayDistance).Date))
                            foreach (var _Product in coreStructure.dicCoord[DatePO].Keys.OrderBy(x => x.ProductCode))
                                foreach (var _CustomerOrder in coreStructure.dicCoord[DatePO][_Product].Keys
                                    .Where(x => x.QuantityOrderKg > 0)
                                    .OrderBy(x => coreStructure.dicCustomer[x.CustomerId].CustomerType)
                                    .ThenBy(x => coreStructure.dicCustomer[x.CustomerId].CustomerCode))
                                    if (coreStructure.dicCoord[DatePO][_Product][_CustomerOrder] != null)
                                    {
                                        foreach (var _SupplierForecast in coreStructure.dicCoord[DatePO][_Product]
                                            [_CustomerOrder].Keys.OrderBy(x =>
                                                coreStructure.dicSupplier[x.SupplierId].SupplierType))
                                        {
                                            var dr = dtMastah.NewRow();

                                            var _Customer = coreStructure.dicCustomer[_CustomerOrder.CustomerId];
                                            var _Supplier = coreStructure.dicSupplier[_SupplierForecast.SupplierId];

                                            ProductUnitRegion _ProductUnitRegion = null;
                                            if (coreStructure.dicProductUnit.TryGetValue(_Product.ProductCode,
                                                out var _ProductUnit))
                                            {
                                                _ProductUnitRegion = coreStructure.dicProductUnit[_Product.ProductCode]
                                                    .ListRegion.FirstOrDefault(x => x.OrderUnitType == _CustomerOrder.Unit);
                                                if (_ProductUnitRegion == null)
                                                    _ProductUnitRegion = new ProductUnitRegion
                                                    {
                                                        OrderUnitType = "Kg",
                                                        OrderUnitPer = 1,
                                                        SaleUnitType = "Kg",
                                                        SaleUnitPer = 1
                                                    };
                                            }
                                            else
                                            {
                                                _ProductUnitRegion = new ProductUnitRegion
                                                {
                                                    OrderUnitType = "Kg",
                                                    OrderUnitPer = 1,
                                                    SaleUnitType = "Kg",
                                                    SaleUnitPer = 1
                                                };
                                            }

                                            dr["Mã 6 ký tự"] = _Product.ProductCode;
                                            dr["Mã thành phẩm VinEco"] = _Product.ProductCode;
                                            //dr["Mã thành phẩm VinCommerce"] = "";
                                            dr["P&L"] = _Customer.Company;
                                            dr["Tên Sản phẩm"] = _Product.ProductName;
                                            dr["Mã Cửa hàng"] = _Customer.CustomerCode;
                                            dr["Tên Cửa hàng"] = _Customer.CustomerName;
                                            dr["Loại Cửa hàng"] = _Customer.CustomerType;
                                            dr["Ngày Tiêu thụ"] = DatePO.Date;
                                            dr["Vùng Tiêu thụ"] = _Customer.CustomerBigRegion == "Miền Bắc" ? "MB" : "MN";
                                            dr["Nhu cầu Kg VinCommerce"] = _CustomerOrder.QuantityOrderKg;

                                            //dr["Số lượng đặt"] = _CustomerOrder.QuantityOrder;
                                            //dr["Đơn vị đặt"] = _CustomerOrder.Unit;
                                            //dr["Đặt Kg/Unit"] = _ProductUnitRegion.OrderUnitPer;

                                            //dr["Số lượng bán"] = (double)_SupplierForecast.QuantityForecast / (double)_ProductUnitRegion.SaleUnitPer;
                                            //dr["Số lượng bán"] = String.Format("= N{0} * Q{0}", dtMastah.Rows.Count + 6);
                                            //dr["Đơn vị bán"] = _ProductUnitRegion.SaleUnitType;
                                            //dr["Bán Kg/Unit"] = _ProductUnitRegion.SaleUnitPer;

                                            dr["Nhu cầu Kg Đã đáp ứng"] = string.Format(
                                                "=SUM( M{0}, P{0}, S{0}, V{0}, Y{0}, AB{0} )", dtMastah.Rows.Count + 6);
                                            //dr["Nhu cầu Kg Đã đáp ứng"] = String.Format("=SUM( S{0}, V{0}, Y{0}, AB{0}, AF{0}, AJ{0} )", dtMastah.Rows.Count + 6);
                                            //dr["Nhu cầu Kg Đã đáp ứng"] = (double)dr["Nhu cầu Kg Đã đáp ứng"] + _CustomerOrder.QuantityOrderKg;

                                            var _Region = string.Join(string.Empty,
                                                _Supplier.SupplierRegion.Where((ch, index) =>
                                                    ch != ' ' && (index == 0 || _Supplier.SupplierRegion[index - 1] == ' ')));
                                            switch (_Supplier.SupplierType)
                                            {
                                                case "VinEco":
                                                    dr["Tên VinEco " + _Region] = _Supplier.SupplierName;
                                                    dr["Đáp ứng từ VinEco " + _Region] = _SupplierForecast.QuantityForecast;
                                                    dr["Ngày sơ chế VinEco " + _Region] =
                                                        coreStructure.dicCoord[DatePO][_Product][_CustomerOrder][
                                                            _SupplierForecast].Date;
                                                    break;
                                                case "ThuMua":
                                                    dr["Tên ThuMua " + _Region] = _Supplier.SupplierName;
                                                    dr["Đáp ứng từ ThuMua " + _Region] = _SupplierForecast.QuantityForecast;
                                                    dr["Ngày sơ chế ThuMua " + _Region] =
                                                        coreStructure.dicCoord[DatePO][_Product][_CustomerOrder][
                                                            _SupplierForecast].Date;
                                                    //dr["Giá mua ThuMua " + _Region] = 0;
                                                    break;
                                                case "VCM":
                                                    dr["Tên ThuMua " + _Region] = "VCM - " + _Supplier.SupplierName;
                                                    dr["Đáp ứng từ ThuMua " + _Region] = _SupplierForecast.QuantityForecast;
                                                    dr["Ngày sơ chế ThuMua " + _Region] =
                                                        coreStructure.dicCoord[DatePO][_Product][_CustomerOrder][
                                                            _SupplierForecast].Date;
                                                    break;
                                                default:
                                                    break;
                                            }

                                            dtMastah.Rows.Add(dr);
                                        }
                                    }
                                    else
                                    {
                                        if (_Product.ProductCode.Substring(0, 1) != "K" &&
                                            _Product.ProductCode.Substring(0, 1) != "D" &&
                                            (DatePO >= DateTo.AddDays(-1) || DatePO < DateFrom))
                                            continue;

                                        var dr = dtMastah.NewRow();

                                        var _Customer = coreStructure.dicCustomer[_CustomerOrder.CustomerId];
                                        //Supplier _Supplier =coreStructure. dicSupplier[_SupplierForecast.SupplierId];

                                        ProductUnitRegion _ProductUnitRegion = null;
                                        if (coreStructure.dicProductUnit.TryGetValue(_Product.ProductCode,
                                            out var _ProductUnit))
                                        {
                                            _ProductUnitRegion = coreStructure.dicProductUnit[_Product.ProductCode].ListRegion
                                                .FirstOrDefault(x => x.OrderUnitType == _CustomerOrder.Unit);
                                            if (_ProductUnitRegion == null)
                                                _ProductUnitRegion = new ProductUnitRegion
                                                {
                                                    OrderUnitType = "Kg",
                                                    OrderUnitPer = 1,
                                                    SaleUnitType = "Kg",
                                                    SaleUnitPer = 1
                                                };
                                        }
                                        else
                                        {
                                            _ProductUnitRegion = new ProductUnitRegion
                                            {
                                                OrderUnitType = "Kg",
                                                OrderUnitPer = 1,
                                                SaleUnitType = "Kg",
                                                SaleUnitPer = 1
                                            };
                                        }

                                        dr["Mã 6 ký tự"] = _Product.ProductCode;
                                        dr["Mã thành phẩm VinEco"] = _Product.ProductCode;
                                        //dr["Mã thành phẩm VinCommerce"] = "";
                                        dr["P&L"] = _Customer.Company;
                                        dr["Tên Sản phẩm"] = _Product.ProductName;
                                        dr["Mã Cửa hàng"] = _Customer.CustomerCode;
                                        dr["Tên Cửa hàng"] = _Customer.CustomerName;
                                        dr["Loại Cửa hàng"] = _Customer.CustomerType;
                                        dr["Ngày Tiêu thụ"] = DatePO.Date;
                                        dr["Vùng Tiêu thụ"] = _Customer.CustomerBigRegion == "Miền Bắc" ? "MB" : "MN";
                                        dr["Nhu cầu Kg VinCommerce"] = _CustomerOrder.QuantityOrderKg;

                                        //dr["Số lượng đặt"] = _CustomerOrder.QuantityOrder;
                                        //dr["Đơn vị đặt"] = _CustomerOrder.Unit;
                                        //dr["Đặt Kg/Unit"] = _ProductUnitRegion.OrderUnitPer;

                                        //dr["Số lượng bán"] = (double)_SupplierForecast.QuantityForecast / (double)_ProductUnitRegion.SaleUnitPer;
                                        //dr["Số lượng bán"] = String.Format("= N{0} * Q{0}", dtMastah.Rows.Count + 6);
                                        //dr["Đơn vị bán"] = _ProductUnitRegion.SaleUnitType;
                                        //dr["Bán Kg/Unit"] = _ProductUnitRegion.SaleUnitPer;

                                        dr["Nhu cầu Kg Đã đáp ứng"] =
                                            string.Format("=SUM( M{0}, P{0}, S{0}, V{0}, Y{0}, AB{0} )",
                                                dtMastah.Rows.Count + 6);
                                        //dr["Nhu cầu Kg Đã đáp ứng"] = String.Format("=SUM( S{0}, V{0}, Y{0}, AB{0}, AF{0}, AJ{0} )", dtMastah.Rows.Count + 6);
                                        //dr["Nhu cầu Kg Đã đáp ứng"] = (double)dr["Nhu cầu Kg Đã đáp ứng"] + _CustomerOrder.QuantityOrderKg;

                                        dtMastah.Rows.Add(dr);
                                    }

                        #endregion

                        #region LeftOverVinEco Table

                        var dtLeftOverVE = new DataTable { TableName = "DBSL dư" };

                        dtLeftOverVE.Columns.Add("Mã 6 ký tự", typeof(string));
                        dtLeftOverVE.Columns.Add("Mã thành phẩm VinEco", typeof(string));
                        dtLeftOverVE.Columns.Add("Mã thành phẩm VinCommerce", typeof(string));
                        dtLeftOverVE.Columns.Add("Tên Sản phẩm", typeof(string));
                        dtLeftOverVE.Columns.Add("Mã Cửa hàng", typeof(string));
                        dtLeftOverVE.Columns.Add("Tên Cửa hàng", typeof(string));
                        dtLeftOverVE.Columns.Add("Loại Cửa hàng", typeof(string));
                        dtLeftOverVE.Columns.Add("Ngày Tiêu thụ", typeof(DateTime));
                        dtLeftOverVE.Columns.Add("Vùng Tiêu thụ", typeof(string));
                        dtLeftOverVE.Columns.Add("Nhu cầu VinCommerce", typeof(double)).DefaultValue = 0;
                        dtLeftOverVE.Columns.Add("Nhu cầu Đã đáp ứng", typeof(string));
                        dtLeftOverVE.Columns.Add("Tên VinEco MB", typeof(string));
                        dtLeftOverVE.Columns.Add("Đáp ứng từ VinEco MB", typeof(double));
                        dtLeftOverVE.Columns.Add("Ngày sơ chế VinEco MB", typeof(DateTime));
                        dtLeftOverVE.Columns.Add("Tên VinEco MN", typeof(string));
                        dtLeftOverVE.Columns.Add("Đáp ứng từ VinEco MN", typeof(double));
                        dtLeftOverVE.Columns.Add("Ngày sơ chế VinEco MN", typeof(DateTime));
                        dtLeftOverVE.Columns.Add("Tên VinEco LĐ", typeof(string));
                        dtLeftOverVE.Columns.Add("Đáp ứng từ VinEco LĐ", typeof(double));
                        dtLeftOverVE.Columns.Add("Ngày sơ chế VinEco LĐ", typeof(DateTime));
                        dtLeftOverVE.Columns.Add("Tên ThuMua MB", typeof(string));
                        dtLeftOverVE.Columns.Add("Đáp ứng từ ThuMua MB", typeof(double));
                        dtLeftOverVE.Columns.Add("Ngày sơ chế ThuMua MB", typeof(DateTime));
                        //dtLeftOverVE.Columns.Add("Giá mua ThuMua MB", typeof(double));
                        dtLeftOverVE.Columns.Add("Tên ThuMua MN", typeof(string));
                        dtLeftOverVE.Columns.Add("Đáp ứng từ ThuMua MN", typeof(double));
                        dtLeftOverVE.Columns.Add("Ngày sơ chế ThuMua MN", typeof(DateTime));
                        //dtLeftOverVE.Columns.Add("Giá mua ThuMua MN", typeof(double));
                        dtLeftOverVE.Columns.Add("Tên ThuMua LĐ", typeof(string));
                        dtLeftOverVE.Columns.Add("Đáp ứng từ ThuMua LĐ", typeof(double));
                        dtLeftOverVE.Columns.Add("Ngày sơ chế ThuMua LĐ", typeof(DateTime));
                        //dtLeftOverVE.Columns.Add("Giá mua ThuMua LĐ", typeof(double));

                        foreach (var DateFC in coreStructure.dicFC.Keys)
                            foreach (var _Product in coreStructure.dicFC[DateFC].Keys.OrderBy(x => x.ProductCode))
                            {
                                var _ListSupplier = coreStructure.dicFC[DateFC][_Product].Keys.Where(x =>
                                    coreStructure.dicSupplier[x.SupplierId].SupplierType == "VinEco");
                                if (_ListSupplier != null)
                                    foreach (var _SupplierForecast in _ListSupplier.Where(x => x.QuantityForecast >= 3)
                                        .OrderBy(x => coreStructure.dicSupplier[x.SupplierId].SupplierName))
                                    {
                                        var dr = dtLeftOverVE.NewRow();

                                        //Customer _Customer =coreStructure. dicCustomer[_CustomerOrder.CustomerId];
                                        var _Supplier = coreStructure.dicSupplier[_SupplierForecast.SupplierId];

                                        dr["Mã 6 ký tự"] = _Product.ProductCode;
                                        dr["Mã thành phẩm VinEco"] = _Product.ProductCode;
                                        dr["Mã thành phẩm VinCommerce"] = "";
                                        dr["Tên Sản phẩm"] = _Product.ProductName;
                                        //dr["Mã Cửa hàng"] = _Customer.CustomerCode;
                                        //dr["Tên Cửa hàng"] = _Customer.CustomerName;
                                        //dr["Loại Cửa hàng"] = _Customer.CustomerType;
                                        //dr["Ngày Tiêu thụ"] = DatePO.Date;
                                        //dr["Vùng Tiêu thụ"] = _Customer.CustomerBigRegion;
                                        //dr["Nhu cầu VinCommerce"] = _CustomerOrder.QuantityOrder;
                                        //dr["Nhu cầu Đã đáp ứng"] = String.Format("=SUM(M{0}, P{0}, S{0}, V{0}, Z{0}, AD{0})", dtLeftOverVE.Rows.Count + 6);

                                        var _Region = string.Join(string.Empty,
                                            _Supplier.SupplierRegion.Where((ch, index) =>
                                                ch != ' ' && (index == 0 || _Supplier.SupplierRegion[index - 1] == ' ')));
                                        switch (_Supplier.SupplierType)
                                        {
                                            case "VinEco":
                                                dr["Tên VinEco " + _Region] = _Supplier.SupplierName;
                                                dr["Đáp ứng từ VinEco " + _Region] = _SupplierForecast.QuantityForecast;
                                                dr["Ngày sơ chế VinEco " + _Region] = DateFC;
                                                break;
                                            case "ThuMua":
                                                dr["Tên ThuMua " + _Region] = _Supplier.SupplierName;
                                                dr["Đáp ứng từ ThuMua " + _Region] = _SupplierForecast.QuantityForecast;
                                                dr["Ngày sơ chế ThuMua " + _Region] = DateFC;
                                                //dr["Giá mua ThuMua " + _Region] = 0;
                                                break;
                                            default:
                                                break;
                                        }

                                        //dr["VE Code"] = _Product.ProductCode;
                                        //dr["VE Name"] = _Product.ProductName;

                                        //Supplier _supplier =coreStructure. dicSupplier[_SupplierForecast.SupplierId];

                                        //dr["SupplierCode"] = _supplier.SupplierCode;
                                        //dr["SupplierName"] = _supplier.SupplierName;
                                        //dr["SupplierRegion"] = _supplier.SupplierRegion;
                                        //dr["SupplierType"] = _supplier.SupplierType;
                                        //dr["QuantityForecast"] = _SupplierForecast.QuantityForecast;
                                        //dr["DateForecast"] = DateFC;

                                        dtLeftOverVE.Rows.Add(dr);
                                    }
                            }

                        var _rowCount = 6;
                        foreach (DataRow dr in dtLeftOverVE.Rows)
                        {
                            dr["Nhu cầu đã đáp ứng"] = string.Format(
                                "=SUMIFS($M${2}:$M${1},$N${2}:$N${1},N{0},$A${2}:$A${1},A{0})+SUMIFS($P${2}:$P${1},$Q${2}:$Q${1},Q{0},$A${2}:$A${1},A{0})+SUMIFS($S${2}:$S${1},$T${2}:$T${1},T{0},$A${2}:$A${1},A{0})",
                                _rowCount, dtLeftOverVE.Rows.Count, 6);
                            _rowCount++;
                        }

                        #endregion

                        #region LeftOverThuMua Table

                        //DataTable dtLeftOverTM = new DataTable() { TableName = "Cam Kết dư" };

                        //dtLeftOverTM.Columns.Add("Mã 6 ký tự", typeof(string)).DefaultValue = "";
                        //dtLeftOverTM.Columns.Add("Mã thành phẩm VinEco", typeof(string)).DefaultValue = "";
                        //dtLeftOverTM.Columns.Add("Mã thành phẩm VinCommerce", typeof(string)).DefaultValue = "";
                        //dtLeftOverTM.Columns.Add("Tên Sản phẩm", typeof(string)).DefaultValue = "";
                        //dtLeftOverTM.Columns.Add("Mã Cửa hàng", typeof(string)).DefaultValue = "";
                        //dtLeftOverTM.Columns.Add("Tên Cửa hàng", typeof(string)).DefaultValue = "";
                        //dtLeftOverTM.Columns.Add("Loại Cửa hàng", typeof(string)).DefaultValue = "";
                        //dtLeftOverTM.Columns.Add("Ngày Tiêu thụ", typeof(DateTime));
                        //dtLeftOverTM.Columns.Add("Vùng Tiêu thụ", typeof(string)).DefaultValue = "";
                        //dtLeftOverTM.Columns.Add("Nhu cầu VinCommerce", typeof(double)).DefaultValue = 0;
                        //dtLeftOverTM.Columns.Add("Nhu cầu Đã đáp ứng", typeof(double)).DefaultValue = 0;
                        //dtLeftOverTM.Columns.Add("Tên VinEco MB", typeof(string)).DefaultValue = "";
                        //dtLeftOverTM.Columns.Add("Đáp ứng từ VinEco MB", typeof(double)).DefaultValue = 0;
                        //dtLeftOverTM.Columns.Add("Ngày sơ chế VinEco MB", typeof(DateTime));
                        //dtLeftOverTM.Columns.Add("Tên VinEco MN", typeof(string)).DefaultValue = "";
                        //dtLeftOverTM.Columns.Add("Đáp ứng từ VinEco MN", typeof(double)).DefaultValue = 0;
                        //dtLeftOverTM.Columns.Add("Ngày sơ chế VinEco MN", typeof(DateTime));
                        //dtLeftOverTM.Columns.Add("Tên VinEco LĐ", typeof(string)).DefaultValue = "";
                        //dtLeftOverTM.Columns.Add("Đáp ứng từ VinEco LĐ", typeof(double)).DefaultValue = 0;
                        //dtLeftOverTM.Columns.Add("Ngày sơ chế VinEco LĐ", typeof(DateTime));
                        //dtLeftOverTM.Columns.Add("Tên ThuMua MB", typeof(string)).DefaultValue = "";
                        //dtLeftOverTM.Columns.Add("Đáp ứng từ ThuMua MB", typeof(double)).DefaultValue = 0;
                        //dtLeftOverTM.Columns.Add("Ngày sơ chế ThuMua MB", typeof(DateTime));
                        //dtLeftOverTM.Columns.Add("Giá mua ThuMua MB", typeof(double)).DefaultValue = 0;
                        //dtLeftOverTM.Columns.Add("Tên ThuMua MN", typeof(string)).DefaultValue = "";
                        //dtLeftOverTM.Columns.Add("Đáp ứng từ ThuMua MN", typeof(double)).DefaultValue = 0;
                        //dtLeftOverTM.Columns.Add("Ngày sơ chế ThuMua MN", typeof(DateTime));
                        //dtLeftOverTM.Columns.Add("Giá mua ThuMua MN", typeof(double)).DefaultValue = 0;
                        //dtLeftOverTM.Columns.Add("Tên ThuMua LĐ", typeof(string)).DefaultValue = "";
                        //dtLeftOverTM.Columns.Add("Đáp ứng từ ThuMua LĐ", typeof(double)).DefaultValue = 0;
                        //dtLeftOverTM.Columns.Add("Ngày sơ chế ThuMua LĐ", typeof(DateTime));
                        //dtLeftOverTM.Columns.Add("Giá mua ThuMua LĐ", typeof(double)).DefaultValue = 0;

                        //foreach (DateTime DateFC in coreStructure.dicFC.Keys.OrderBy(x => x.Date))
                        //{
                        //    foreach (Product _Product in coreStructure.dicFC[DateFC].Keys.OrderBy(x => x.ProductCode))
                        //    {
                        //        var _ListSupplier = coreStructure.dicFC[DateFC][_Product].Keys.Where(x => coreStructure.dicSupplier[x.SupplierId].SupplierType == "ThuMua");
                        //        if (_ListSupplier != null)
                        //        {
                        //            foreach (SupplierForecast _SupplierForecast in _ListSupplier.Where(x => x.QuantityForecast >= 3 && x.QuantityForecastPlanned != null).OrderBy(x => coreStructure.dicSupplier[x.SupplierId].SupplierName))
                        //            {
                        //                DataRow dr = dtLeftOverTM.NewRow();

                        //                //Customer _Customer =coreStructure. dicCustomer[_CustomerOrder.CustomerId];
                        //                Supplier _Supplier = coreStructure.dicSupplier[_SupplierForecast.SupplierId];

                        //                dr["Mã 6 ký tự"] = _Product.ProductCode;
                        //                dr["Mã thành phẩm VinEco"] = _Product.ProductCode;
                        //                dr["Mã thành phẩm VinCommerce"] = "";
                        //                dr["Tên Sản phẩm"] = _Product.ProductName;
                        //                //dr["Mã Cửa hàng"] = _Customer.CustomerCode;
                        //                //dr["Tên Cửa hàng"] = _Customer.CustomerName;
                        //                //dr["Loại Cửa hàng"] = _Customer.CustomerType;
                        //                //dr["Ngày Tiêu thụ"] = DatePO.Date;
                        //                //dr["Vùng Tiêu thụ"] = _Customer.CustomerBigRegion;
                        //                //dr["Nhu cầu VinCommerce"] = _CustomerOrder.QuantityOrder;
                        //                dr["Nhu cầu Đã đáp ứng"] = _SupplierForecast.QuantityForecast;

                        //                string _Region = string.Join(String.Empty, _Supplier.SupplierRegion.Where((ch, index) => ch != ' ' && (index == 0 || _Supplier.SupplierRegion[index - 1] == ' ')));
                        //                switch (_Supplier.SupplierType)
                        //                {
                        //                    case "VinEco":
                        //                        dr["Tên VinEco " + _Region] = _Supplier.SupplierName;
                        //                        dr["Đáp ứng từ VinEco " + _Region] = _SupplierForecast.QuantityForecast;
                        //                        dr["Ngày sơ chế VinEco " + _Region] = DateFC;
                        //                        break;
                        //                    case "ThuMua":
                        //                        dr["Tên ThuMua " + _Region] = _Supplier.SupplierName;
                        //                        dr["Đáp ứng từ ThuMua " + _Region] = _SupplierForecast.QuantityForecast;
                        //                        dr["Ngày sơ chế ThuMua " + _Region] = DateFC;
                        //                        //dr["Giá mua ThuMua " + _Region] = 0;
                        //                        break;
                        //                    default:
                        //                        break;
                        //                }

                        //                //dr["VE Code"] = _Product.ProductCode;
                        //                //dr["VE Name"] = _Product.ProductName;

                        //                //Supplier _supplier =coreStructure. dicSupplier[_SupplierForecast.SupplierId];

                        //                //dr["SupplierCode"] = _supplier.SupplierCode;
                        //                //dr["SupplierName"] = _supplier.SupplierName;
                        //                //dr["SupplierRegion"] = _supplier.SupplierRegion;
                        //                //dr["SupplierType"] = _supplier.SupplierType;
                        //                //dr["QuantityForecast"] = _SupplierForecast.QuantityForecast;
                        //                //dr["DateForecast"] = DateFC;

                        //                dtLeftOverTM.Rows.Add(dr);
                        //            }
                        //        }
                        //    }
                        //}

                        #endregion

                        #region Customer Table

                        var dtCustomer = new DataTable { TableName = "Region I guess" };

                        dtCustomer.Columns.Add("Mã cửa hàng", typeof(string));
                        dtCustomer.Columns.Add("Vùng đặt hàng", typeof(string));
                        dtCustomer.Columns.Add("Loại cửa hàng", typeof(string));
                        dtCustomer.Columns.Add("Vùng tiêu thụ", typeof(string));
                        dtCustomer.Columns.Add("Tên cửa hàng", typeof(string));
                        dtCustomer.Columns.Add("P&L", typeof(string));

                        foreach (var _Customer in coreStructure.dicCustomer.Values)
                        {
                            var dr = dtCustomer.NewRow();

                            dr["Mã cửa hàng"] = _Customer.CustomerCode;
                            dr["Tên cửa hàng"] = _Customer.CustomerName;
                            dr["Loại cửa hàng"] = _Customer.CustomerType;
                            dr["Vùng tiêu thụ"] = _Customer.CustomerRegion;
                            dr["Vùng đặt hàng"] = _Customer.CustomerBigRegion;
                            dr["P&L"] = _Customer.Company;

                            dtCustomer.Rows.Add(dr);
                        }

                        #endregion

                        #region Output to Excel

                        //Excel.Application xlApp = new Excel.Application();

                        //xlApp.ScreenUpdating = false;
                        //xlApp.EnableEvents = false;
                        //xlApp.DisplayAlerts = false;
                        //xlApp.DisplayStatusBar = false;
                        //xlApp.AskToUpdateLinks = false;

                        //db.DropCollection("CoordResult");
                        //await db.GetCollection<CoordResult>("CoordResult").InsertManyAsync(CoordResultList);

                        //CoordResultList = null;

                        var filePath = Application.StartupPath.Replace("\\bin\\Debug", "") + "\\Template\\{0}";
                        var fileFullPath = string.Format(filePath, "ChiaHang Mastah.xlsb");
                        var fileFullPath2007 = string.Format(filePath, "ChiaHang Mastah.xlsm");
                        //WriteToRichTextBoxOutput(filePath);
                        //WriteToRichTextBoxOutput(fileFullPath);

                        //Excel.Workbook xlWb = xlApp.Workbooks.Open(
                        //    Filename: fileFullPath,
                        //    UpdateLinks: false,
                        //    ReadOnly: false,
                        //    Format: 5,
                        //    Password: "",
                        //    WriteResPassword: "",
                        //    IgnoreReadOnlyRecommended: true,
                        //    Origin: Excel.XlPlatform.xlWindows,
                        //    Delimiter: "",
                        //    Editable: true,
                        //    Notify: false,
                        //    Converter: 0,
                        //    AddToMru: true,
                        //    Local: false,
                        //    CorruptLoad: false);

                        //xlApp.Calculation = Excel.XlCalculation.xlCalculationManual;

                        var missing = Type.Missing;
                        var path = string.Format(
                            @"D:\Documents\Stuff\VinEco\Mastah Project\Test\ChiaHang Mastah {1}{0}.xlsb",
                            DateFrom.AddDays(dayDistance).ToString("dd.MM") + " - " +
                            DateTo.AddDays(-dayDistance).ToString("dd.MM") + " (" +
                            DateTime.Now.ToString("yyyyMMdd HH\\hmm") + ")", FruitOnly ? "Fruit " : "");
                        //string path2007 = string.Format(@"D:\Documents\Stuff\VinEco\Mastah Project\Test\ChiaHang Mastah {0}.xlsm", DateFrom.AddDays(3).ToString("dd.MM") + " - " + DateTo.AddDays(-3).ToString("dd.MM") + " (" + DateTime.Now.ToString("yyyyMMdd HH\\hmm") + ")");
                        WriteToRichTextBoxOutput(path);

                        //xlWb.SaveAs(path, Excel.XlFileFormat.xlExcel12, missing, missing, false, false, Excel.XlSaveAsAccessMode.xlNoChange, missing, missing, missing);

                        var xlWb = new Workbook(fileFullPath);

                        // Optimize for Performance
                        xlWb.Settings.MemorySetting = MemorySetting.MemoryPreference;

                        //OutputExcel(dtMastah, "Mastah", xlWb);
                        //OutputExcel(dtLeftOverVE, "DBSL dư", xlWb);
                        //OutputExcel(dtLeftOverTM, "Cam Kết dư", xlWb);
                        //OutputExcel(dtVeFarm, "VE Farm", xlWb, true, 1, false);
                        //OutputExcel(dtThuMua, "VE ThuMua", xlWb, true, 1, false);
                        //OutputExcel(dtCustomer, "Region I guess", xlWb, true, 1, false);

                        var DicColDate = new Dictionary<string, int>();

                        //foreach (DataColumn dc in dtMastah.Columns)
                        //{
                        //    if (dc.DataType == typeof(DateTime))
                        //    {
                        //        DicColDate.Add(dc.ColumnName, dtMastah.Columns.IndexOf(dc));
                        //    }
                        //}

                        OutputExcelAspose(dtMastah, "Mastah", xlWb, false, 6, "A1", DicColDate, "dd/MM/yyyy");
                        OutputExcelAspose(dtLeftOverVE, "DBSL dư", xlWb, false, 6, "A1", DicColDate, "dd/MM/yyyy");
                        //OutputExcelAspose(dtLeftOverTM, "Cam Kết dư", xlWb, false, 6);
                        OutputExcelAspose(dtVeFarm, "VE Farm", xlWb, true, 1);
                        OutputExcelAspose(dtThuMua, "VE ThuMua", xlWb, true, 1);
                        OutputExcelAspose(dtCustomer, "Region I guess", xlWb, true, 1);

                        // Handling Formulas from DataTable.
                        // ... 'cause they will not be handled in Aspose.Cells.
                        var xlWsMastah = xlWb.Worksheets["Mastah"];

                        var dicTable = new Dictionary<DataTable, int>();

                        dicTable.Add(dtMastah, 6);
                        dicTable.Add(dtLeftOverVE, 6);

                        foreach (var _dt in dicTable.Keys)
                        {
                            var _xlWs = xlWb.Worksheets[_dt.TableName];
                            var _colIndex = 0;
                            var _rowFirst = dicTable[_dt];

                            foreach (DataColumn dc in _dt.Columns)
                            {
                                if (dc.DataType == typeof(string))
                                {
                                    var _dr = _dt.Rows[1];
                                    if (_dr[dc].ToString().Length > 0 && _dr[dc].ToString().Substring(0, 1) == "=")
                                    {
                                        var _rowIndex = _rowFirst - 1;
                                        foreach (DataRow dr in _dt.Rows)
                                        {
                                            _xlWs.Cells[_rowIndex, _colIndex].Formula = dr[dc].ToString();
                                            _rowIndex++;
                                        }
                                    }
                                }
                                _colIndex++;
                            }
                        }

                        // Date stuff
                        //xlWb.Worksheets["Mastah"].Cells[2, 3].Value = DateTo.Date - DateFrom.Date;
                        xlWsMastah.Cells[1, 2].Value =
                            DateFrom.Date
                                .AddDays(dayDistance); // (int)(DateFrom.Date - _dateBase).TotalDays + 2 + dayDistance;
                        xlWsMastah.Cells[2, 2].Value = DateTo.Date < DateFrom.Date
                            ? DateFrom.Date.AddDays(-dayDistance)
                            : DateTo.Date
                                .AddDays(
                                    -dayDistance); // (int)((DateFrom > DateTo ? DateFrom : DateTo).Date - _dateBase).TotalDays + 2 - dayDistance;

                        // Formula Stuff
                        //xlWb.Worksheets["Mastah"].Cells[4, 1].Formula = String.Format("=SUBTOTAL(3,A6:A{0}", dtMastah.Rows.Count + 5); // A4

                        //xlWb.Worksheets["Mastah"].Cells[4, 5].Formula = String.Format("=SUBTOTAL(3,E6:E{0}", dtMastah.Rows.Count + 5); // E4

                        //xlWb.Worksheets["Mastah"].Cells[4, 9].Formula = String.Format("=SUBTOTAL(3,I6:I{0}", dtMastah.Rows.Count + 5); // I4
                        //xlWb.Worksheets["Mastah"].Cells[4, 10].Formula = String.Format("=SUBTOTAL(9,J6:J{0}", dtMastah.Rows.Count + 5); // J4
                        //xlWb.Worksheets["Mastah"].Cells[4, 11].Formula = String.Format("=SUBTOTAL(9,K6:K{0}", dtMastah.Rows.Count + 5); // K4
                        //xlWb.Worksheets["Mastah"].Cells[4, 12].Formula = String.Format("=SUBTOTAL(3,L6:L{0}", dtMastah.Rows.Count + 5); // L4
                        //xlWb.Worksheets["Mastah"].Cells[4, 13].Formula = String.Format("=SUBTOTAL(9,M6:M{0}", dtMastah.Rows.Count + 5); // M4

                        //xlWb.Worksheets["Mastah"].Cells[4, 15].Formula = String.Format("=SUBTOTAL(3,O6:O{0}", dtMastah.Rows.Count + 5); // O4
                        //xlWb.Worksheets["Mastah"].Cells[4, 16].Formula = String.Format("=SUBTOTAL(9,P6:P{0}", dtMastah.Rows.Count + 5); // P4

                        //xlWb.Worksheets["Mastah"].Cells[4, 18].Formula = String.Format("=SUBTOTAL(3,R6:R{0}", dtMastah.Rows.Count + 5); // R4
                        //xlWb.Worksheets["Mastah"].Cells[4, 19].Formula = String.Format("=SUBTOTAL(9,S6:S{0}", dtMastah.Rows.Count + 5); // S4

                        //xlWb.Worksheets["Mastah"].Cells[4, 21].Formula = String.Format("=SUBTOTAL(3,U6:U{0}", dtMastah.Rows.Count + 5); // U4
                        //xlWb.Worksheets["Mastah"].Cells[4, 22].Formula = String.Format("=SUBTOTAL(9,V6:V{0}", dtMastah.Rows.Count + 5); // V4

                        //xlWb.Worksheets["Mastah"].Cells[4, 25].Formula = String.Format("=SUBTOTAL(3,Y6:Y{0}", dtMastah.Rows.Count + 5); // Y4
                        //xlWb.Worksheets["Mastah"].Cells[4, 26].Formula = String.Format("=SUBTOTAL(9,Z6:Z{0}", dtMastah.Rows.Count + 5); // Z4

                        //xlWb.Worksheets["Mastah"].Cells[4, 29].Formula = String.Format("=SUBTOTAL(3,AC6:AC{0}", dtMastah.Rows.Count + 5); // AC4
                        //xlWb.Worksheets["Mastah"].Cells[4, 30].Formula = String.Format("=SUBTOTAL(9,AD6:AD{0}", dtMastah.Rows.Count + 5); // AD4

                        // For new format - with Unit
                        //xlWsMastah.Cells[3, 0].Formula = String.Format("=SUBTOTAL(3, A6:A{0} )", dtMastah.Rows.Count + 5); // A
                        //xlWsMastah.Cells[3, 4].Formula = String.Format("=SUBTOTAL(3,E6:E{0})", dtMastah.Rows.Count + 5); // E
                        //xlWsMastah.Cells[3, 9].Formula = String.Format("=SUBTOTAL(9,J6:J{0})", dtMastah.Rows.Count + 5); // J
                        //xlWsMastah.Cells[3, 10].Formula = String.Format("=SUBTOTAL(9,K6:K{0})", dtMastah.Rows.Count + 5); // K
                        //xlWsMastah.Cells[3, 11].Formula = String.Format("=SUBTOTAL(3,L6:L{0})", dtMastah.Rows.Count + 5); // L
                        //xlWsMastah.Cells[3, 13].Formula = String.Format("=SUBTOTAL(9,N6:N{0})", dtMastah.Rows.Count + 5); // N
                        //xlWsMastah.Cells[3, 14].Formula = String.Format("=SUBTOTAL(9,O6:O{0})", dtMastah.Rows.Count + 5); // O
                        //xlWsMastah.Cells[3, 15].Formula = String.Format("=SUBTOTAL(3,P6:P{0})", dtMastah.Rows.Count + 5); // P
                        //xlWsMastah.Cells[3, 17].Formula = String.Format("=SUBTOTAL(3,R6:R{0})", dtMastah.Rows.Count + 5); // R
                        //xlWsMastah.Cells[3, 18].Formula = String.Format("=SUBTOTAL(9,S6:S{0})", dtMastah.Rows.Count + 5); // S
                        //xlWsMastah.Cells[3, 20].Formula = String.Format("=SUBTOTAL(3,U6:U{0})", dtMastah.Rows.Count + 5); // U
                        //xlWsMastah.Cells[3, 21].Formula = String.Format("=SUBTOTAL(9,V6:V{0})", dtMastah.Rows.Count + 5); // V
                        //xlWsMastah.Cells[3, 23].Formula = String.Format("=SUBTOTAL(3,X6:X{0})", dtMastah.Rows.Count + 5); // X
                        //xlWsMastah.Cells[3, 24].Formula = String.Format("=SUBTOTAL(9,Y6:Y{0})", dtMastah.Rows.Count + 5); // Y
                        //xlWsMastah.Cells[3, 26].Formula = String.Format("=SUBTOTAL(3,AA6:AA{0})", dtMastah.Rows.Count + 5); // AA
                        //xlWsMastah.Cells[3, 27].Formula = String.Format("=SUBTOTAL(9,AB6:AB{0})", dtMastah.Rows.Count + 5); // AB
                        //xlWsMastah.Cells[3, 30].Formula = String.Format("=SUBTOTAL(3,AE6:AE{0})", dtMastah.Rows.Count + 5); // AE
                        //xlWsMastah.Cells[3, 31].Formula = String.Format("=SUBTOTAL(9,AF6:AF{0})", dtMastah.Rows.Count + 5); // AF
                        //xlWsMastah.Cells[3, 34].Formula = String.Format("=SUBTOTAL(3,AI6:AI{0})", dtMastah.Rows.Count + 5); // AI
                        //xlWsMastah.Cells[3, 35].Formula = String.Format("=SUBTOTAL(9,AJ6:AJ{0})", dtMastah.Rows.Count + 5); // AJ

                        // Old format - without Unit
                        xlWsMastah.Cells[3, 0].Formula =
                            string.Format("=SUBTOTAL(3,A6:A{0})", dtMastah.Rows.Count + 5); // A4
                        xlWsMastah.Cells[3, 4].Formula =
                            string.Format("=SUBTOTAL(3,E6:E{0})", dtMastah.Rows.Count + 5); // E4
                        xlWsMastah.Cells[3, 8].Formula =
                            string.Format("=SUBTOTAL(3,I6:I{0})", dtMastah.Rows.Count + 5); // I4
                        xlWsMastah.Cells[3, 9].Formula =
                            string.Format("=SUBTOTAL(9,J6:J{0})", dtMastah.Rows.Count + 5); // J4
                        xlWsMastah.Cells[3, 10].Formula =
                            string.Format("=SUBTOTAL(9,K6:K{0})", dtMastah.Rows.Count + 5); // K4
                        xlWsMastah.Cells[3, 11].Formula =
                            string.Format("=SUBTOTAL(3,L6:L{0})", dtMastah.Rows.Count + 5); // L4
                        xlWsMastah.Cells[3, 12].Formula =
                            string.Format("=SUBTOTAL(9,M6:M{0})", dtMastah.Rows.Count + 5); // M4
                        xlWsMastah.Cells[3, 14].Formula =
                            string.Format("=SUBTOTAL(3,O6:O{0})", dtMastah.Rows.Count + 5); // O4
                        xlWsMastah.Cells[3, 15].Formula =
                            string.Format("=SUBTOTAL(9,P6:P{0})", dtMastah.Rows.Count + 5); // P4
                        xlWsMastah.Cells[3, 17].Formula =
                            string.Format("=SUBTOTAL(3,R6:R{0})", dtMastah.Rows.Count + 5); // R4
                        xlWsMastah.Cells[3, 18].Formula =
                            string.Format("=SUBTOTAL(9,S6:S{0})", dtMastah.Rows.Count + 5); // S4
                        xlWsMastah.Cells[3, 20].Formula =
                            string.Format("=SUBTOTAL(3,U6:U{0})", dtMastah.Rows.Count + 5); // U4
                        xlWsMastah.Cells[3, 21].Formula =
                            string.Format("=SUBTOTAL(9,V6:V{0})", dtMastah.Rows.Count + 5); // V4
                        xlWsMastah.Cells[3, 23].Formula =
                            string.Format("=SUBTOTAL(3,X6:X{0})", dtMastah.Rows.Count + 5); // Y4
                        xlWsMastah.Cells[3, 24].Formula =
                            string.Format("=SUBTOTAL(9,Y6:Y{0})", dtMastah.Rows.Count + 5); // Z4
                        xlWsMastah.Cells[3, 26].Formula =
                            string.Format("=SUBTOTAL(3,AA6:AA{0})", dtMastah.Rows.Count + 5); // AC4
                        xlWsMastah.Cells[3, 27].Formula =
                            string.Format("=SUBTOTAL(9,AB6:AB{0})", dtMastah.Rows.Count + 5); // AD4

                        xlWsMastah = null;

                        // Formula Stuff for Leftover VE
                        xlWb.Worksheets["DBSL Dư"].Cells[3, 0].Formula =
                            string.Format("=SUBTOTAL(3,A6:A{0})", dtLeftOverVE.Rows.Count + 5); // A4
                        xlWb.Worksheets["DBSL Dư"].Cells[3, 4].Formula =
                            string.Format("=SUBTOTAL(3,E6:E{0})", dtLeftOverVE.Rows.Count + 5); // E4
                        xlWb.Worksheets["DBSL Dư"].Cells[3, 8].Formula =
                            string.Format("=SUBTOTAL(3,I6:I{0})", dtLeftOverVE.Rows.Count + 5); // I4
                        xlWb.Worksheets["DBSL Dư"].Cells[3, 9].Formula =
                            string.Format("=SUBTOTAL(9,J6:J{0})", dtLeftOverVE.Rows.Count + 5); // J4
                        xlWb.Worksheets["DBSL Dư"].Cells[3, 10].Formula =
                            string.Format("=SUBTOTAL(9,K6:K{0})", dtLeftOverVE.Rows.Count + 5); // K4
                        xlWb.Worksheets["DBSL Dư"].Cells[3, 11].Formula =
                            string.Format("=SUBTOTAL(3,L6:L{0})", dtLeftOverVE.Rows.Count + 5); // L4
                        xlWb.Worksheets["DBSL Dư"].Cells[3, 12].Formula =
                            string.Format("=SUBTOTAL(9,M6:M{0})", dtLeftOverVE.Rows.Count + 5); // M4
                        xlWb.Worksheets["DBSL Dư"].Cells[3, 14].Formula =
                            string.Format("=SUBTOTAL(3,O6:O{0})", dtLeftOverVE.Rows.Count + 5); // O4
                        xlWb.Worksheets["DBSL Dư"].Cells[3, 15].Formula =
                            string.Format("=SUBTOTAL(9,P6:P{0})", dtLeftOverVE.Rows.Count + 5); // P4
                        xlWb.Worksheets["DBSL Dư"].Cells[3, 17].Formula =
                            string.Format("=SUBTOTAL(3,R6:R{0})", dtLeftOverVE.Rows.Count + 5); // R4
                        xlWb.Worksheets["DBSL Dư"].Cells[3, 18].Formula =
                            string.Format("=SUBTOTAL(9,S6:S{0})", dtLeftOverVE.Rows.Count + 5); // S4
                        //xlWb.Worksheets["DBSL Dư"].Cells[3, 20].Formula = String.Format("=SUBTOTAL(3,U6:U{0})", dtLeftOverVE.Rows.Count + 5); // U4
                        //xlWb.Worksheets["DBSL Dư"].Cells[3, 21].Formula = String.Format("=SUBTOTAL(9,V6:V{0})", dtLeftOverVE.Rows.Count + 5); // V4
                        //xlWb.Worksheets["DBSL Dư"].Cells[3, 24].Formula = String.Format("=SUBTOTAL(3,Y6:Y{0})", dtLeftOverVE.Rows.Count + 5); // Y4
                        //xlWb.Worksheets["DBSL Dư"].Cells[3, 26].Formula = String.Format("=SUBTOTAL(9,Z6:Z{0})", dtLeftOverVE.Rows.Count + 5); // Z4
                        //xlWb.Worksheets["DBSL Dư"].Cells[3, 29].Formula = String.Format("=SUBTOTAL(3,AA6:AA{0})", dtLeftOverVE.Rows.Count + 5); // AC4
                        //xlWb.Worksheets["DBSL Dư"].Cells[3, 30].Formula = String.Format("=SUBTOTAL(9,AC6:AC{0})", dtLeftOverVE.Rows.Count + 5); // AD4

                        xlWb.CalculateFormula();
                        xlWb.Save(path, SaveFormat.Xlsb);

                        Delete_Evaluation_Sheet_Interop(path);

                        //using (ExcelPackage pck = new ExcelPackage(new FileInfo(fileFullPath2007)))
                        //{
                        //    OutputExcelEpplus(pck, dtMastah, "Mastah", false, 6, false);

                        //    pck.SaveAs(new FileInfo(path2007));
                        //}

                        //xlApp.Calculation = Excel.XlCalculation.xlCalculationAutomatic;

                        //xlWb.Close(SaveChanges: true);
                        //if (xlWb != null) { Marshal.ReleaseComObject(xlWb); }
                        xlWb = null;

                        //xlApp.ScreenUpdating = true;
                        //xlApp.EnableEvents = true;
                        //xlApp.DisplayAlerts = false;
                        //xlApp.DisplayStatusBar = true;
                        //xlApp.AskToUpdateLinks = true;

                        //xlApp.Quit();
                        //if (xlApp != null) { Marshal.ReleaseComObject(xlApp); }
                        //xlApp = null;

                        #endregion

                        #endregion
                    }
                    if (!YesNoGroupFarm)
                    {
                        var dtMastah = new DataTable { TableName = "Mastah" };

                        if (YesNoGroupThuMua)
                        {
                            #region Mastah Table

                            dtMastah.Columns.Add("Mã 6 ký tự", typeof(string));
                            dtMastah.Columns.Add("Tên sản phẩm", typeof(string));
                            dtMastah.Columns.Add("Loại cửa hàng", typeof(string));
                            dtMastah.Columns.Add("Ngày tiêu thụ", typeof(DateTime));
                            dtMastah.Columns.Add("Vùng tiêu thụ", typeof(string));
                            dtMastah.Columns.Add("Nhu cầu VinCommerce", typeof(double));
                            dtMastah.Columns.Add("Nhu cầu Đáp ứng", typeof(double));
                            dtMastah.Columns.Add("Nguồn", typeof(string));
                            dtMastah.Columns.Add("Vùng sản xuất", typeof(string));
                            dtMastah.Columns.Add("Tên NCC", typeof(string));
                            dtMastah.Columns.Add("Ngày sơ chế", typeof(DateTime));

                            dicRow = new Dictionary<string, int>();
                            foreach (var DatePO in coreStructure.dicCoord.Keys)
                                foreach (var _Product in coreStructure.dicCoord[DatePO].Keys.OrderBy(x => x.ProductCode))
                                    foreach (var _CustomerOrder in coreStructure.dicCoord[DatePO][_Product].Keys
                                        .OrderBy(x => coreStructure.dicCustomer[x.CustomerId].CustomerType))
                                    {
                                        var _Customer = coreStructure.dicCustomer[_CustomerOrder.CustomerId];
                                        var sKey = string.Format("{0}{1}{2}{3}", DatePO.Date, _Customer.CustomerType,
                                            _Customer.CustomerBigRegion, _Product.ProductCode);
                                        if (coreStructure.dicCoord[DatePO][_Product][_CustomerOrder] != null)
                                        {
                                            foreach (var _SupplierForecast in coreStructure.dicCoord[DatePO][_Product]
                                                [_CustomerOrder].Keys.OrderBy(x =>
                                                    coreStructure.dicSupplier[x.SupplierId].SupplierType))
                                            {
                                                var _Supplier = coreStructure.dicSupplier[_SupplierForecast.SupplierId];

                                                sKey += _Supplier.SupplierType;

                                                DataRow dr = null;
                                                var _rowIndex = 0;
                                                if (!dicRow.TryGetValue(sKey, out _rowIndex))
                                                {
                                                    dr = dtMastah.NewRow();
                                                    dicRow.Add(sKey, dtMastah.Rows.Count);
                                                    dtMastah.Rows.Add(dr);
                                                    dr = dtMastah.Rows[dtMastah.Rows.Count - 1];
                                                }
                                                else
                                                {
                                                    dr = dtMastah.Rows[_rowIndex];
                                                }

                                                var _Region = string.Join(string.Empty,
                                                    _Supplier.SupplierRegion.Where((ch, index) =>
                                                        ch != ' ' &&
                                                        (index == 0 || _Supplier.SupplierRegion[index - 1] == ' ')));
                                                var _colName = string.Format("{0} {1}", _Supplier.SupplierType, _Region);

                                                dr["Mã 6 ký tự"] = _Product.ProductCode;
                                                dr["Tên sản phẩm"] = _Product.ProductName;
                                                dr["Loại cửa hàng"] = _Customer.CustomerType;
                                                //dr["Ngày tiêu thụ"] = (int)(DatePO.Date - _dateBase).TotalDays + 2;
                                                dr["Ngày tiêu thụ"] = DatePO.Date;
                                                dr["Vùng tiêu thụ"] = _Customer.CustomerBigRegion;
                                                dr["Nhu cầu VinCommerce"] =
                                                    Convert.ToDouble(dr["Nhu cầu VinCommerce"]) + _CustomerOrder.QuantityOrder;
                                                dr["Nhu cầu Đáp ứng"] =
                                                    Convert.ToDouble(dr["Nhu cầu Đáp ứng"]) +
                                                    _SupplierForecast.QuantityForecast;
                                                ;
                                                dr["Nguồn"] = _Supplier.SupplierType;
                                                dr["Vùng sản xuất"] = _Supplier.SupplierRegion;
                                                dr["Tên NCC"] = _Supplier.SupplierType == "ThuMua"
                                                    ? "ThuMua"
                                                    : _Supplier.SupplierName;
                                                //dr["Ngày sơ chế"] = (int)(coreStructure.dicCoord[DatePO][_Product][_CustomerOrder][_SupplierForecast].Date - _dateBase).TotalDays + 2;
                                                dr["Ngày sơ chế"] =
                                                    coreStructure.dicCoord[DatePO][_Product][_CustomerOrder][_SupplierForecast]
                                                        .Date;
                                            }
                                        }
                                        else
                                        {
                                            DataRow dr = null;
                                            var _rowIndex = 0;
                                            if (!dicRow.TryGetValue(sKey, out _rowIndex))
                                            {
                                                dr = dtMastah.NewRow();
                                                dicRow.Add(sKey, dtMastah.Rows.Count);
                                                dtMastah.Rows.Add(dr);
                                                dr = dtMastah.Rows[dtMastah.Rows.Count - 1];
                                            }
                                            else
                                            {
                                                dr = dtMastah.Rows[_rowIndex];
                                            }

                                            dr["Mã 6 ký tự"] = _Product.ProductCode;
                                            dr["Tên sản phẩm"] = _Product.ProductName;
                                            dr["Loại cửa hàng"] = _Customer.CustomerType;
                                            //dr["Ngày tiêu thụ"] = (int)(DatePO.Date - _dateBase).TotalDays + 2;
                                            dr["Ngày tiêu thụ"] = DatePO.Date;
                                            dr["Vùng tiêu thụ"] = _Customer.CustomerBigRegion;
                                            dr["Nhu cầu VinCommerce"] =
                                                Convert.ToDouble(dr["Nhu cầu VinCommerce"]) + _CustomerOrder.QuantityOrder;
                                        }
                                    }

                            #endregion
                        }
                        else
                        {
                            #region Mastah Table

                            dtMastah.Columns.Add("Mã 6 ký tự", typeof(string));
                            dtMastah.Columns.Add("Tên sản phẩm", typeof(string));
                            dtMastah.Columns.Add("ProductOrientation", typeof(string));
                            dtMastah.Columns.Add("ProductClimate", typeof(string));
                            dtMastah.Columns.Add("ProductionGroup", typeof(string));
                            dtMastah.Columns.Add("Nhóm sản phẩm", typeof(string));
                            dtMastah.Columns.Add("Ghi chú", typeof(string));
                            dtMastah.Columns.Add("Loại cửa hàng", typeof(string));
                            dtMastah.Columns.Add("P&L", typeof(string));
                            dtMastah.Columns.Add("Ngày tiêu thụ", typeof(DateTime));
                            dtMastah.Columns.Add("Tỉnh tiêu thụ", typeof(string));
                            dtMastah.Columns.Add("Vùng tiêu thụ", typeof(string));
                            dtMastah.Columns.Add("Vùng SX yêu cầu", typeof(string));
                            dtMastah.Columns.Add("Nguồn yêu cầu", typeof(string));
                            dtMastah.Columns.Add("Nhu cầu", typeof(double)).DefaultValue = 0;
                            dtMastah.Columns.Add("Đáp ứng", typeof(double)).DefaultValue = 0;
                            dtMastah.Columns.Add("Nguồn", typeof(string));
                            dtMastah.Columns.Add("Vùng sản xuất", typeof(string));
                            dtMastah.Columns.Add("Mã NCC", typeof(string));
                            dtMastah.Columns.Add("Tên NCC", typeof(string));
                            dtMastah.Columns.Add("Ngày sơ chế", typeof(DateTime));
                            dtMastah.Columns.Add("NoSup", typeof(double)).DefaultValue = 0;
                            dtMastah.Columns.Add("KPI", typeof(double)).DefaultValue = 0;
                            dtMastah.Columns.Add("Label", typeof(string));
                            dtMastah.Columns.Add("CodeSFG", typeof(string));
                            dtMastah.Columns.Add("IsNoSup", typeof(bool)).DefaultValue = false;

                            foreach (var DatePO in coreStructure.dicCoord.Keys.OrderBy(x => x.Date)
                                .Where(x => x.Date >= DateFrom.AddDays(dayDistance).Date))
                                foreach (var _Product in coreStructure.dicCoord[DatePO].Keys.OrderBy(x => x.ProductCode))
                                    foreach (var _CustomerOrder in coreStructure.dicCoord[DatePO][_Product].Keys
                                        .Where(x => x.QuantityOrderKg > 0)
                                        .OrderBy(x => coreStructure.dicCustomer[x.CustomerId].CustomerType)
                                        .ThenBy(x => coreStructure.dicCustomer[x.CustomerId].CustomerCode))
                                    {
                                        var _Customer = coreStructure.dicCustomer[_CustomerOrder.CustomerId];
                                        var sKey = string.Format("{0}{1}{2}{3}{4}{5}", DatePO.Date.ToString("yyyyMMdd"),
                                            _Customer.CustomerType, _Customer.Company, _Customer.CustomerBigRegion,
                                            _Product.ProductCode, YesNoSubRegion ? _Customer.CustomerRegion : null);
                                        if (coreStructure.dicCoord[DatePO][_Product][_CustomerOrder] != null)
                                        {
                                            foreach (var _SupplierForecast in coreStructure.dicCoord[DatePO][_Product]
                                                [_CustomerOrder].Keys.OrderBy(x =>
                                                    coreStructure.dicSupplier[x.SupplierId].SupplierType))
                                            {
                                                var _Supplier = coreStructure.dicSupplier[_SupplierForecast.SupplierId];

                                                sKey += _Supplier.SupplierCode;

                                                DataRow dr = null;
                                                var _rowIndex = 0;
                                                if (!dicRow.TryGetValue(sKey, out _rowIndex))
                                                {
                                                    dr = dtMastah.NewRow();
                                                    dicRow.Add(sKey, dtMastah.Rows.Count);
                                                    dtMastah.Rows.Add(dr);
                                                    dr = dtMastah.Rows[dtMastah.Rows.Count - 1];
                                                }
                                                else
                                                {
                                                    dr = dtMastah.Rows[_rowIndex];
                                                }

                                                var _Region = string.Join(string.Empty,
                                                    _Supplier.SupplierRegion.Split(' ').Select(x => x.First()));

                                                //string _Region = string.Join(String.Empty, _Supplier.SupplierRegion.Where((ch, index) => ch != ' ' && (index == 0 || _Supplier.SupplierRegion[index - 1] == ' ')));
                                                var _colName = string.Format("{0} {1}", _Supplier.SupplierType, _Region);

                                                dr["Mã 6 ký tự"] = _Product.ProductCode;
                                                dr["Tên sản phẩm"] = _Product.ProductName;
                                                dr["Nhóm sản phẩm"] = _Product.ProductClassification;
                                                dr["ProductOrientation"] = _Product.ProductOrientation;
                                                dr["ProductClimate"] = _Product.ProductClimate;
                                                dr["ProductionGroup"] = _Product.ProductionGroup;
                                                dr["Ghi chú"] =
                                                    _Product.ProductNote.Contains(_Customer.CustomerBigRegion == "Miền Nam"
                                                        ? "South"
                                                        : "North")
                                                        ? "Ok"
                                                        : "Out of List";
                                                dr["Loại cửa hàng"] = _Customer.CustomerType;
                                                dr["P&L"] = _Customer.Company;
                                                //dr["Ngày tiêu thụ"] = (int)(DatePO.Date - _dateBase).TotalDays + 2;
                                                dr["Ngày tiêu thụ"] = DatePO.Date;
                                                dr["Vùng tiêu thụ"] = _Customer.CustomerBigRegion;
                                                dr["Tỉnh tiêu thụ"] = YesNoSubRegion ? _Customer.CustomerRegion : null;
                                                dr["Vùng SX yêu cầu"] = _CustomerOrder.DesiredRegion ?? "Any";
                                                dr["Nguồn yêu cầu"] = _CustomerOrder.DesiredSource ?? "Any";
                                                dr["Nhu cầu"] = (double)dr["Nhu cầu"] + _CustomerOrder.QuantityOrderKg;
                                                dr["Đáp ứng"] = (double)dr["Đáp ứng"] + _SupplierForecast.QuantityForecast;
                                                ;
                                                dr["Nguồn"] = _Supplier.SupplierType;
                                                dr["Vùng sản xuất"] = _Supplier.SupplierRegion;
                                                dr["Mã NCC"] = _Supplier.SupplierCode;
                                                dr["Tên NCC"] = _Supplier.SupplierName;
                                                //dr["Ngày sơ chế"] = (int)(coreStructure.dicCoord[DatePO][_Product][_CustomerOrder][_SupplierForecast].Date - _dateBase).TotalDays + 2;
                                                dr["Ngày sơ chế"] =
                                                    coreStructure.dicCoord[DatePO][_Product][_CustomerOrder][_SupplierForecast]
                                                        .Date;
                                                dr["Label"] = _SupplierForecast.LabelVinEco ? "Yes" : "No";
                                                dr["CodeSFG"] = string.Format("{0}{1}{2}", _Product.ProductCode, 1,
                                                    (_Supplier.SupplierRegion == "Lâm Đồng" ? 0 : 2) +
                                                    (_SupplierForecast.LabelVinEco ? 1 : 0));
                                            }
                                        }
                                        else
                                        {
                                            DataRow dr = null;
                                            var _rowIndex = 0;
                                            if (!dicRow.TryGetValue(sKey, out _rowIndex))
                                            {
                                                dr = dtMastah.NewRow();
                                                dicRow.Add(sKey, dtMastah.Rows.Count);
                                                dtMastah.Rows.Add(dr);
                                                dr = dtMastah.Rows[dtMastah.Rows.Count - 1];
                                            }
                                            else
                                            {
                                                dr = dtMastah.Rows[_rowIndex];
                                            }

                                            dr["Mã 6 ký tự"] = _Product.ProductCode;
                                            dr["Tên sản phẩm"] = _Product.ProductName;
                                            dr["Nhóm sản phẩm"] = _Product.ProductClassification;
                                            dr["ProductOrientation"] = _Product.ProductOrientation;
                                            dr["ProductClimate"] = _Product.ProductClimate;
                                            dr["ProductionGroup"] = _Product.ProductionGroup;
                                            dr["Ghi chú"] =
                                                _Product.ProductNote.Contains(_Customer.CustomerBigRegion == "Miền Nam"
                                                    ? "South"
                                                    : "North")
                                                    ? "Ok"
                                                    : "Out of List";
                                            dr["Loại cửa hàng"] = _Customer.CustomerType;
                                            dr["P&L"] = _Customer.Company;
                                            //dr["Ngày tiêu thụ"] = (int)(DatePO.Date - _dateBase).TotalDays + 2;
                                            dr["Ngày tiêu thụ"] = DatePO.Date;
                                            dr["Vùng tiêu thụ"] = _Customer.CustomerBigRegion;
                                            dr["Tỉnh tiêu thụ"] = YesNoSubRegion ? _Customer.CustomerRegion : null;
                                            dr["Vùng SX yêu cầu"] = _CustomerOrder.DesiredRegion ?? "Any";
                                            dr["Nguồn yêu cầu"] = _CustomerOrder.DesiredSource ?? "Any";
                                            dr["Nhu cầu"] = Convert.ToDouble(dr["Nhu cầu"] ?? 0) +
                                                            _CustomerOrder.QuantityOrderKg;
                                            dr["Nguồn"] = "Không đáp ứng";
                                        }
                                    }

                            foreach (DataRow dr in dtMastah.Rows)
                            {
                                dr["NoSup"] = Math.Max((double)dr["Nhu cầu"] - (double)dr["Đáp ứng"], 0);
                                if ((double)dr["NoSup"] > 1) dr["IsNoSup"] = true;
                            }

                            var _FC = db.GetCollection<ForecastDate>("Forecast").Find(x =>
                                    x.DateForecast >= DateFrom.Date &&
                                    x.DateForecast <= DateTo.Date)
                                .ToList()
                                .OrderByDescending(x => x.DateForecast);

                            foreach (var _ForecastDate in _FC)
                                foreach (var _ProductForecast in _ForecastDate.ListProductForecast)
                                    foreach (var _SupplierForecast in _ProductForecast.ListSupplierForecast.Where(x =>
                                        x.QualityControlPass && x.QuantityForecastPlanned > 0))
                                    {
                                        var dr = dtMastah.NewRow();

                                        var _Product = coreStructure.dicProduct[_ProductForecast.ProductId];
                                        var _Supplier = coreStructure.dicSupplier[_SupplierForecast.SupplierId];

                                        if (FruitOnly)
                                            if (_Product.ProductCode != "K" && _Product.ProductCode != "D01401")
                                                continue;

                                        dr["Mã 6 ký tự"] = _Product.ProductCode;
                                        dr["Tên sản phẩm"] = _Product.ProductName;
                                        dr["Nhóm sản phẩm"] = _Product.ProductClassification;
                                        dr["ProductOrientation"] = _Product.ProductOrientation;
                                        dr["ProductClimate"] = _Product.ProductClimate;
                                        dr["ProductionGroup"] = _Product.ProductionGroup;
                                        dr["Ghi chú"] = _Product.ProductNote.Count != 0 ? "Ok" : "Out of List";

                                        dr["Nguồn"] = _Supplier.SupplierType;
                                        dr["Vùng sản xuất"] = _Supplier.SupplierRegion;
                                        dr["Mã NCC"] = _Supplier.SupplierCode;
                                        dr["Tên NCC"] = _Supplier.SupplierName;

                                        //dr["Ngày sơ chế"] = (int)(_ForecastDate.DateForecast.Date - _dateBase).TotalDays + 2;
                                        dr["Ngày sơ chế"] = _ForecastDate.DateForecast.Date;
                                        dr["KPI"] = _SupplierForecast.QuantityForecastPlanned;

                                        dtMastah.Rows.Add(dr);
                                    }

                            #endregion
                        }

                        #region LeftoverVinEco

                        var dtLeftoverVe = new DataTable();

                        dtLeftoverVe.TableName = "NoCusVinEco";

                        dtLeftoverVe.Columns.Add("Mã VinEco", typeof(string));
                        dtLeftoverVe.Columns.Add("Tên VinEco", typeof(string));
                        dtLeftoverVe.Columns.Add("Nhóm sản phẩm", typeof(string));
                        dtLeftoverVe.Columns.Add("Mã Farm", typeof(string));
                        dtLeftoverVe.Columns.Add("Tên Farm", typeof(string));
                        dtLeftoverVe.Columns.Add("Ngày thu hoạch", typeof(DateTime));
                        dtLeftoverVe.Columns.Add("Vùng sản xuất", typeof(string));
                        dtLeftoverVe.Columns.Add("Sản lượng", typeof(double)).DefaultValue = 0;

                        foreach (var DateFC in coreStructure.dicFC.Keys)
                            foreach (var _Product in coreStructure.dicFC[DateFC].Keys.OrderBy(x => x.ProductCode))
                            {
                                var _ListSupplier = coreStructure.dicFC[DateFC][_Product].Keys.Where(x =>
                                    coreStructure.dicSupplier[x.SupplierId].SupplierType == "VinEco");
                                if (_ListSupplier != null)
                                    foreach (var _SupplierForecast in _ListSupplier.Where(x => x.QuantityForecast > 3)
                                        .OrderBy(x => coreStructure.dicSupplier[x.SupplierId].SupplierName))
                                        if (_SupplierForecast.QuantityForecast > 0)
                                        {
                                            var dr = dtLeftoverVe.NewRow();

                                            //Customer _Customer =coreStructure. dicCustomer[_CustomerOrder.CustomerId];
                                            var _Supplier = coreStructure.dicSupplier[_SupplierForecast.SupplierId];

                                            dr["Mã VinEco"] = _Product.ProductCode;
                                            dr["Tên VinEco"] = _Product.ProductName;
                                            dr["Nhóm sản phẩm"] = _Product.ProductClassification;
                                            dr["Mã Farm"] = _Supplier.SupplierCode;
                                            dr["Tên Farm"] = _Supplier.SupplierName;
                                            //dr["Ngày thu hoạch"] = (int)(DateFC.Date - _dateBase).TotalDays + 2;
                                            dr["Ngày thu hoạch"] = DateFC.Date;
                                            dr["Vùng sản xuất"] = _Supplier.SupplierRegion;
                                            dr["Sản lượng"] = _SupplierForecast.QuantityForecast;

                                            dtLeftoverVe.Rows.Add(dr);
                                        }
                            }

                        #endregion

                        #region LeftoverThuMua

                        var dtLeftoverTmKPI = new DataTable();

                        dtLeftoverTmKPI.TableName = "NoCusThuMuaKPI";

                        dtLeftoverTmKPI.Columns.Add("Mã VinEco", typeof(string));
                        dtLeftoverTmKPI.Columns.Add("Tên VinEco", typeof(string));
                        dtLeftoverTmKPI.Columns.Add("Nhóm sản phẩm", typeof(string));
                        dtLeftoverTmKPI.Columns.Add("Ghi chú", typeof(string));
                        dtLeftoverTmKPI.Columns.Add("Mã NCC", typeof(string));
                        dtLeftoverTmKPI.Columns.Add("Tên NCC", typeof(string));
                        dtLeftoverTmKPI.Columns.Add("Ngày thu hoạch", typeof(DateTime));
                        dtLeftoverTmKPI.Columns.Add("Vùng sản xuất", typeof(string));
                        dtLeftoverTmKPI.Columns.Add("Sản lượng", typeof(double)).DefaultValue = 0;

                        foreach (var DateFC in coreStructure.dicFC.Keys)
                            foreach (var _Product in coreStructure.dicFC[DateFC].Keys.OrderBy(x => x.ProductCode))
                            {
                                var _ListSupplier = coreStructure.dicFC[DateFC][_Product].Keys.Where(x =>
                                    coreStructure.dicSupplier[x.SupplierId].SupplierType != "VinEco");
                                if (_ListSupplier != null)
                                    foreach (var _SupplierForecast in _ListSupplier.Where(x => x.QuantityForecast >= 3)
                                        .OrderBy(x => coreStructure.dicSupplier[x.SupplierId].SupplierName))
                                        if (_SupplierForecast.QuantityForecast > 0)
                                        {
                                            var dr = dtLeftoverTmKPI.NewRow();

                                            //Customer _Customer =coreStructure. dicCustomer[_CustomerOrder.CustomerId];
                                            var _Supplier = coreStructure.dicSupplier[_SupplierForecast.SupplierId];

                                            dr["Mã VinEco"] = _Product.ProductCode;
                                            dr["Tên VinEco"] = _Product.ProductName;
                                            dr["Nhóm sản phẩm"] = _Product.ProductClassification;
                                            dr["Ghi chú"] = _Product.ProductNote.Count != 0 ? "Ok" : "Out of List";
                                            dr["Mã NCC"] = _Supplier.SupplierCode;
                                            dr["Tên NCC"] = _Supplier.SupplierName;
                                            //dr["Ngày thu hoạch"] = (int)(DateFC.Date - _dateBase).TotalDays + 2;
                                            dr["Ngày thu hoạch"] = DateFC.Date;
                                            dr["Vùng sản xuất"] = _Supplier.SupplierRegion;
                                            dr["Sản lượng"] = _SupplierForecast.QuantityForecast;

                                            dtLeftoverTmKPI.Rows.Add(dr);
                                        }
                            }

                        #endregion

                        #region LeftoverThuMua

                        var dtLeftoverTm = new DataTable();

                        dtLeftoverTm.TableName = "NoCusThuMua";

                        dtLeftoverTm.Columns.Add("Mã VinEco", typeof(string));
                        dtLeftoverTm.Columns.Add("Tên VinEco", typeof(string));
                        dtLeftoverTm.Columns.Add("Nhóm sản phẩm", typeof(string));
                        dtLeftoverTm.Columns.Add("Ghi chú", typeof(string));
                        dtLeftoverTm.Columns.Add("Mã NCC", typeof(string));
                        dtLeftoverTm.Columns.Add("Tên NCC", typeof(string));
                        dtLeftoverTm.Columns.Add("Ngày thu hoạch", typeof(DateTime));
                        dtLeftoverTm.Columns.Add("Vùng sản xuất", typeof(string));
                        dtLeftoverTm.Columns.Add("Sản lượng", typeof(double)).DefaultValue = 0;

                        foreach (var DateFC in coreStructure.dicFC.Keys)
                            foreach (var _Product in coreStructure.dicFC[DateFC].Keys.OrderBy(x => x.ProductCode))
                            {
                                var _ListSupplier = coreStructure.dicFC[DateFC][_Product].Keys.Where(x =>
                                    coreStructure.dicSupplier[x.SupplierId].SupplierType != "VinEco");
                                if (_ListSupplier != null)
                                    foreach (var _SupplierForecast in _ListSupplier
                                        .Where(x => x.QuantityForecastOriginal >= 3).OrderBy(x =>
                                            coreStructure.dicSupplier[x.SupplierId].SupplierName))
                                        if (_SupplierForecast.QuantityForecast > 0)
                                        {
                                            var dr = dtLeftoverTm.NewRow();

                                            //Customer _Customer =coreStructure. dicCustomer[_CustomerOrder.CustomerId];
                                            var _Supplier = coreStructure.dicSupplier[_SupplierForecast.SupplierId];

                                            dr["Mã VinEco"] = _Product.ProductCode;
                                            dr["Tên VinEco"] = _Product.ProductName;
                                            dr["Nhóm sản phẩm"] = _Product.ProductClassification;
                                            dr["Ghi chú"] = _Product.ProductNote.Count != 0 ? "Ok" : "Out of List";
                                            dr["Mã NCC"] = _Supplier.SupplierCode;
                                            dr["Tên NCC"] = _Supplier.SupplierName;
                                            //dr["Ngày thu hoạch"] = (int)(DateFC.Date - _dateBase).TotalDays + 2;
                                            dr["Ngày thu hoạch"] = DateFC.Date;
                                            dr["Vùng sản xuất"] = _Supplier.SupplierRegion;
                                            dr["Sản lượng"] = _SupplierForecast.QuantityForecastOriginal;

                                            dtLeftoverTm.Rows.Add(dr);
                                        }
                            }

                        #endregion

                        #region Output to Excel - OpenXMLWriter Style, super fast.

                        //string fileName = string.Format("Mastah Compact {0}.xlsx", DateFrom.AddDays(dayDistance).ToString("dd.MM") + " - " + DateTo.AddDays(-dayDistance).ToString("dd.MM") + " (" + DateTime.Now.ToString("yyyyMMdd HH\\hmm") + ")");
                        //string path = string.Format(@"D:\Documents\Stuff\VinEco\Mastah Project\Test\" + fileName);

                        //var listDt = new List<DataTable>();

                        //listDt.Add(dtMastah);
                        //listDt.Add(dtLeftoverVe);
                        //listDt.Add(dtLeftoverTm);

                        //LargeExportOneWorkbook(path, listDt, true, true);

                        //ConvertToXlsbInterop(path, "xlsx", "xlsb", true);

                        //LargeExport(dtMastah, path, true, true);

                        #endregion

                        #region Output to Excel - Aspose.Cell with Optimized Memory Settings.

                        var fileName =
                            $"Mastah Compact {UpperCap:P0} {(FruitOnly ? "Fruit " : "")}{DateFrom.AddDays(dayDistance).ToString("dd.MM") + " - " + DateTo.AddDays(-dayDistance).ToString("dd.MM") + " (" + DateTime.Now.ToString("yyyyMMdd HH\\hmm") + ")"}.xlsb";
                        var path = string.Format(
                            @"D:\Documents\Stuff\VinEco\Mastah Project\Test\" + fileName);

                        using (var xlWb = new Workbook())
                        {
                            var DicColDate = new Dictionary<string, int>();
                            xlWb.Settings.MemorySetting = MemorySetting.MemoryPreference;

                            // Mastah
                            xlWb.Worksheets.Add(dtMastah.TableName);
                            DicColDate.Clear();
                            foreach (DataColumn dc in dtMastah.Columns)
                                if (dc.DataType == typeof(DateTime))
                                    DicColDate.Add(dc.ColumnName, dtMastah.Columns.IndexOf(dc));
                            OutputExcelAspose(dtMastah, dtMastah.TableName, xlWb, true, 1, "A1", DicColDate);

                            // VinEco Leftover
                            xlWb.Worksheets.Add(dtLeftoverVe.TableName);
                            DicColDate.Clear();
                            foreach (DataColumn dc in dtLeftoverVe.Columns)
                                if (dc.DataType == typeof(DateTime))
                                    DicColDate.Add(dc.ColumnName, dtLeftoverVe.Columns.IndexOf(dc));
                            OutputExcelAspose(dtLeftoverVe, dtLeftoverVe.TableName, xlWb, true, 1, "A1", DicColDate);

                            // Contracted procurement Leftover.
                            xlWb.Worksheets.Add(dtLeftoverTmKPI.TableName);
                            DicColDate.Clear();
                            foreach (DataColumn dc in dtLeftoverTmKPI.Columns)
                                if (dc.DataType == typeof(DateTime))
                                    DicColDate.Add(dc.ColumnName, dtLeftoverTmKPI.Columns.IndexOf(dc));
                            OutputExcelAspose(dtLeftoverTmKPI, dtLeftoverTmKPI.TableName, xlWb, true, 1, "A1",
                                DicColDate);

                            // Normal procurement Leftover.
                            xlWb.Worksheets.Add(dtLeftoverTm.TableName);
                            DicColDate.Clear();
                            foreach (DataColumn dc in dtLeftoverTm.Columns)
                                if (dc.DataType == typeof(DateTime))
                                    DicColDate.Add(dc.ColumnName, dtLeftoverTm.Columns.IndexOf(dc));
                            OutputExcelAspose(dtLeftoverTm, dtLeftoverTm.TableName, xlWb, true, 1, "A1", DicColDate);

                            xlWb.Worksheets.RemoveAt("sheet1");

                            xlWb.CalculateFormula();
                            xlWb.Save(path, SaveFormat.Xlsb);
                        }
                        Delete_Evaluation_Sheet_Interop(path);

                        #endregion
                    }
                    else
                    {
                        #region Mastah Table

                        var dtMastah = new DataTable();

                        dtMastah.TableName = "Mastah";

                        dtMastah.Columns.Add("Mã 6 ký tự", typeof(string)).DefaultValue = "";
                        dtMastah.Columns.Add("Tên sản phẩm", typeof(string)).DefaultValue = "";
                        dtMastah.Columns.Add("Loại cửa hàng", typeof(string)).DefaultValue = "";
                        dtMastah.Columns.Add("Ngày tiêu thụ", typeof(int)).DefaultValue = 0;
                        dtMastah.Columns.Add("Vùng tiêu thụ", typeof(string)).DefaultValue = "";
                        dtMastah.Columns.Add("Nhu cầu VinCommerce", typeof(double)).DefaultValue = 0;
                        dtMastah.Columns.Add("Nhu cầu Đáp ứng", typeof(double)).DefaultValue = 0;
                        dtMastah.Columns.Add("Tổng VinEco", typeof(double)).DefaultValue = 0;
                        dtMastah.Columns.Add("Tổng ThuMua", typeof(double)).DefaultValue = 0;
                        dtMastah.Columns.Add("VinEco MB", typeof(double)).DefaultValue = 0;
                        dtMastah.Columns.Add("VinEco MN", typeof(double)).DefaultValue = 0;
                        dtMastah.Columns.Add("VinEco LĐ", typeof(double)).DefaultValue = 0;
                        dtMastah.Columns.Add("ThuMua MB", typeof(double)).DefaultValue = 0;
                        dtMastah.Columns.Add("ThuMua MN", typeof(double)).DefaultValue = 0;
                        dtMastah.Columns.Add("ThuMua LĐ", typeof(double)).DefaultValue = 0;

                        dicRow = new Dictionary<string, int>();
                        foreach (var DatePO in coreStructure.dicCoord.Keys)
                            foreach (var _Product in coreStructure.dicCoord[DatePO].Keys.OrderBy(x => x.ProductCode))
                                foreach (var _CustomerOrder in coreStructure.dicCoord[DatePO][_Product].Keys
                                    .OrderBy(x => coreStructure.dicCustomer[x.CustomerId].CustomerType))
                                {
                                    var _Customer = coreStructure.dicCustomer[_CustomerOrder.CustomerId];
                                    var sKey = string.Format("{0}{1}{2}{3}", DatePO.Date, _Customer.CustomerType,
                                        _Customer.CustomerBigRegion, _Product.ProductCode);
                                    if (coreStructure.dicCoord[DatePO][_Product][_CustomerOrder] != null)
                                    {
                                        foreach (var _SupplierForecast in coreStructure.dicCoord[DatePO][_Product]
                                            [_CustomerOrder].Keys.OrderBy(x =>
                                                coreStructure.dicSupplier[x.SupplierId].SupplierType))
                                        {
                                            var _Supplier = coreStructure.dicSupplier[_SupplierForecast.SupplierId];

                                            DataRow dr = null;
                                            var _rowIndex = 0;
                                            if (!dicRow.TryGetValue(sKey, out _rowIndex))
                                            {
                                                dr = dtMastah.NewRow();
                                                dicRow.Add(sKey, dtMastah.Rows.Count);
                                                dtMastah.Rows.Add(dr);
                                                dr = dtMastah.Rows[dtMastah.Rows.Count - 1];
                                            }
                                            else
                                            {
                                                dr = dtMastah.Rows[_rowIndex];
                                            }

                                            var _Region = string.Join(string.Empty,
                                                _Supplier.SupplierRegion.Where((ch, index) =>
                                                    ch != ' ' && (index == 0 || _Supplier.SupplierRegion[index - 1] == ' ')));
                                            var _colName = string.Format("{0} {1}", _Supplier.SupplierType, _Region);

                                            dr["Mã 6 ký tự"] = _Product.ProductCode;
                                            dr["Tên sản phẩm"] = _Product.ProductName;
                                            dr["Loại cửa hàng"] = _Customer.CustomerType;
                                            dr["Ngày tiêu thụ"] = (int)(DatePO.Date - _dateBase).TotalDays + 2;
                                            dr["Vùng tiêu thụ"] = _Customer.CustomerBigRegion;
                                            dr["Nhu cầu VinCommerce"] =
                                                Convert.ToDouble(dr["Nhu cầu VinCommerce"]) + _CustomerOrder.QuantityOrder;

                                            dr[_colName] = Convert.ToDouble(dr[_colName]) + _SupplierForecast.QuantityForecast;

                                            dr["Tổng " + _Supplier.SupplierType] =
                                                Convert.ToDouble(dr["Tổng " + _Supplier.SupplierType]) +
                                                _SupplierForecast.QuantityForecast;
                                            dr["Nhu cầu Đáp ứng"] =
                                                Convert.ToDouble(dr["Nhu cầu Đáp ứng"]) + _SupplierForecast.QuantityForecast;
                                            ;
                                        }
                                    }
                                    else
                                    {
                                        DataRow dr = null;
                                        var _rowIndex = 0;
                                        if (!dicRow.TryGetValue(sKey, out _rowIndex))
                                        {
                                            dr = dtMastah.NewRow();
                                            dicRow.Add(sKey, dtMastah.Rows.Count);
                                            dtMastah.Rows.Add(dr);
                                            dr = dtMastah.Rows[dtMastah.Rows.Count - 1];
                                        }
                                        else
                                        {
                                            dr = dtMastah.Rows[_rowIndex];
                                        }

                                        dr["Mã 6 ký tự"] = _Product.ProductCode;
                                        dr["Tên sản phẩm"] = _Product.ProductName;
                                        dr["Loại cửa hàng"] = _Customer.CustomerType;
                                        dr["Ngày tiêu thụ"] = (int)(DatePO.Date - _dateBase).TotalDays + 2;
                                        dr["Vùng tiêu thụ"] = _Customer.CustomerBigRegion;
                                        dr["Nhu cầu VinCommerce"] =
                                            Convert.ToDouble(dr["Nhu cầu VinCommerce"]) + _CustomerOrder.QuantityOrder;
                                    }
                                }

                        #endregion

                        #region LeftoverVinEco

                        var dtLeftoverVe = new DataTable();

                        dtLeftoverVe.TableName = "NoCusVinEco";

                        dtLeftoverVe.Columns.Add("Mã VinEco", typeof(string)).DefaultValue = "";
                        dtLeftoverVe.Columns.Add("Tên VinEco", typeof(string)).DefaultValue = "";
                        dtLeftoverVe.Columns.Add("Mã Farm", typeof(string)).DefaultValue = "";
                        dtLeftoverVe.Columns.Add("Tên Farm", typeof(string)).DefaultValue = "";
                        dtLeftoverVe.Columns.Add("Ngày thu hoạch", typeof(int));
                        dtLeftoverVe.Columns.Add("Vùng sản xuất", typeof(string)).DefaultValue = "";
                        dtLeftoverVe.Columns.Add("Sản lượng", typeof(double)).DefaultValue = 0;

                        foreach (var DateFC in coreStructure.dicFC.Keys)
                            foreach (var _Product in coreStructure.dicFC[DateFC].Keys.OrderBy(x => x.ProductCode))
                            {
                                var _ListSupplier = coreStructure.dicFC[DateFC][_Product].Keys.Where(x =>
                                    coreStructure.dicSupplier[x.SupplierId].SupplierType == "VinEco");
                                if (_ListSupplier != null)
                                    foreach (var _SupplierForecast in _ListSupplier.OrderBy(x =>
                                        coreStructure.dicSupplier[x.SupplierId].SupplierName))
                                        if (_SupplierForecast.QuantityForecast > 0)
                                        {
                                            var dr = dtLeftoverVe.NewRow();

                                            //Customer _Customer =coreStructure. dicCustomer[_CustomerOrder.CustomerId];
                                            var _Supplier = coreStructure.dicSupplier[_SupplierForecast.SupplierId];

                                            dr["Mã VinEco"] = _Product.ProductCode;
                                            dr["Tên VinEco"] = _Product.ProductName;
                                            dr["Mã Farm"] = _Supplier.SupplierCode;
                                            dr["Tên Farm"] = _Supplier.SupplierName;
                                            dr["Ngày thu hoạch"] = (int)(DateFC.Date - _dateBase).TotalDays + 2;
                                            dr["Vùng sản xuất"] = _Supplier.SupplierRegion;
                                            dr["Sản lượng"] = _SupplierForecast.QuantityForecast;

                                            dtLeftoverVe.Rows.Add(dr);
                                        }
                            }

                        #endregion

                        #region LeftoverThuMua

                        var dtLeftoverTm = new DataTable();

                        dtLeftoverTm.TableName = "NoCusThuMua";

                        dtLeftoverTm.Columns.Add("Mã VinEco", typeof(string)).DefaultValue = "";
                        dtLeftoverTm.Columns.Add("Tên VinEco", typeof(string)).DefaultValue = "";
                        dtLeftoverTm.Columns.Add("Mã Farm", typeof(string)).DefaultValue = "";
                        dtLeftoverTm.Columns.Add("Tên Farm", typeof(string)).DefaultValue = "";
                        dtLeftoverTm.Columns.Add("Ngày thu hoạch", typeof(int));
                        dtLeftoverTm.Columns.Add("Vùng sản xuất", typeof(string)).DefaultValue = "";
                        dtLeftoverTm.Columns.Add("Sản lượng", typeof(double)).DefaultValue = 0;

                        foreach (var DateFC in coreStructure.dicFC.Keys)
                            foreach (var _Product in coreStructure.dicFC[DateFC].Keys.OrderBy(x => x.ProductCode))
                            {
                                var _ListSupplier = coreStructure.dicFC[DateFC][_Product].Keys.Where(x =>
                                    coreStructure.dicSupplier[x.SupplierId].SupplierType == "ThuMua");
                                if (_ListSupplier != null)
                                    foreach (var _SupplierForecast in _ListSupplier.OrderBy(x =>
                                        coreStructure.dicSupplier[x.SupplierId].SupplierName))
                                        if (_SupplierForecast.QuantityForecast > 0)
                                        {
                                            var dr = dtLeftoverTm.NewRow();

                                            //Customer _Customer =coreStructure. dicCustomer[_CustomerOrder.CustomerId];
                                            var _Supplier = coreStructure.dicSupplier[_SupplierForecast.SupplierId];

                                            dr["Mã VinEco"] = _Product.ProductCode;
                                            dr["Tên VinEco"] = _Product.ProductName;
                                            dr["Mã Farm"] = _Supplier.SupplierCode;
                                            dr["Tên Farm"] = _Supplier.SupplierName;
                                            dr["Ngày thu hoạch"] = (int)(DateFC.Date - _dateBase).TotalDays + 2;
                                            dr["Vùng sản xuất"] = _Supplier.SupplierRegion;
                                            dr["Sản lượng"] = _SupplierForecast.QuantityForecast;

                                            dtLeftoverTm.Rows.Add(dr);
                                        }
                            }

                        #endregion

                        #region Output to Excel - OpenXMLWriter Style, super fast.

                        var fileName =
                            $"Mastah Compact {UpperCap:P0} {(FruitOnly ? "Fruit " : "")}{DateFrom.AddDays(dayDistance).ToString("dd.MM") + " - " + DateTo.AddDays(-dayDistance).ToString("dd.MM") + " (" + DateTime.Now.ToString("yyyyMMdd HH\\hmm") + ")"}.xlsb";
                        var path = string.Format(
                            @"D:\Documents\Stuff\VinEco\Mastah Project\Test\" + fileName);

                        var listDt = new List<DataTable>();

                        listDt.Add(dtMastah);
                        listDt.Add(dtLeftoverVe);
                        listDt.Add(dtLeftoverTm);

                        LargeExportOneWorkbook(path, listDt, true, true);

                        ConvertToXlsbInterop(path, "xlsx", "xlsb", true);

                        //LargeExport(dtMastah, path, true, true);

                        #endregion
                    }
                }

                #endregion

                #region Clean up

                // Cleanup
                db = null;

                PO = null;
                FC = null;

                coreStructure = null;

                //GC.Collect();
                //GC.WaitForPendingFinalizers();

                #endregion

                stopWatch.Stop();
                WriteToRichTextBoxOutput(string.Format("Done in {0}s!", Math.Round(stopWatch.Elapsed.TotalSeconds, 2)));
            }
            catch (Exception ex)
            {
                WriteToRichTextBoxOutput(ex.Message);
                throw;
            }
            finally
            {
            }
        }

        private void CoordNew(CoordStructure coreStructure)
        {
            // List of Suppliers' Regions
            var ListSupplierRegion = new[]
            {
                "North",
                "South",
                "Highland"
            };
            // List of Targets, in prioritized order.
            var ListPriorityTarget = new[]
            {
                "B2B",
                "VM+ VinEco",
                "VM+ Priority",
                "VM Priority",
                "VM+",
                "VM",
                ""
            };
            var ListPrioritySupplier = new[]
            {
                "VCM",
                "VinEco",
                "ThuMua"
            };

            // Highest layer. Date of Demand.
            foreach (var DemandDate in coreStructure.dicPO.Keys.OrderByDescending(x => x.Date).Reverse())
                // Second layer - Priority Target.
                foreach (var PriorityTarget in ListPriorityTarget)
                {
                    var TemporaryProductDictionary = new Dictionary<Product, bool>();

                    var PONorth = coreStructure.dicPO[DemandDate];
                    var POSouth = coreStructure.dicPO[DemandDate.AddDays(-coreStructure.dicTransferDays["Highland-North"])];

                    foreach (var CurrentProduct in PONorth.Keys)
                        if (!TemporaryProductDictionary.ContainsKey(CurrentProduct))
                            TemporaryProductDictionary.Add(CurrentProduct, true);
                    foreach (var CurrentProduct in POSouth.Keys)
                        if (!TemporaryProductDictionary.ContainsKey(CurrentProduct))
                            TemporaryProductDictionary.Add(CurrentProduct, true);

                    foreach (var CurrentProduct in TemporaryProductDictionary.Keys)
                    {
                        var _result = new Dictionary<SupplierForecast, bool>();
                        var SupplyNorth = new Dictionary<Guid, SupplierForecast>();
                        var SupplySouth = new Dictionary<Guid, SupplierForecast>();
                        var SupplyHighland = new Dictionary<Guid, SupplierForecast>();

                        if (coreStructure.dicFC[DemandDate.AddDays(-coreStructure.dicTransferDays["North-North"])]
                            .TryGetValue(CurrentProduct, out _result))
                            SupplyNorth = _result.Keys.Where(x =>
                                x.Availability.Contains(
                                    (DemandDate.AddDays(-coreStructure.dicTransferDays["North-North"]).DayOfWeek + 1)
                                    .ToString())).ToDictionary(x => x.SupplierForecastId);

                        if (coreStructure.dicFC[DemandDate.AddDays(-coreStructure.dicTransferDays["Highland-North"])]
                            .TryGetValue(CurrentProduct, out _result))
                            SupplySouth = _result.Keys.Where(x =>
                                x.Availability.Contains(
                                    (DemandDate.AddDays(-coreStructure.dicTransferDays["Highland-North"]).DayOfWeek + 1)
                                    .ToString())).ToDictionary(x => x.SupplierForecastId);
                        ;

                        if (coreStructure.dicFC[DemandDate.AddDays(-coreStructure.dicTransferDays["Highland-North"])]
                            .TryGetValue(CurrentProduct, out _result))
                            SupplyHighland = _result.Keys.Where(x =>
                                x.Availability.Contains(
                                    (DemandDate.AddDays(-coreStructure.dicTransferDays["Highland-North"]).DayOfWeek + 1)
                                    .ToString())).ToDictionary(x => x.SupplierForecastId);
                        ;

                        var ListRate = new double[3];

                        // Total Demand. Customers' Regions
                        var DemandNorth = !PONorth.ContainsKey(CurrentProduct)
                            ? 0
                            : PONorth[CurrentProduct].Keys
                                .Where(x =>
                                    coreStructure.dicCustomer[x.CustomerId].CustomerBigRegion == "Miền Bắc" &&
                                    PriorityTarget != ""
                                        ? coreStructure.dicCustomer[x.CustomerId].CustomerType == PriorityTarget
                                        : true)
                                .Sum(x => x.QuantityOrderKg);

                        // In case VM+, have to calculate rate twice coz fuck the police. Really.
                        var DemandNorthVM = !PriorityTarget.Contains("VM+")
                            ? 0
                            : PONorth[CurrentProduct].Keys
                                .Where(x =>
                                    coreStructure.dicCustomer[x.CustomerId].CustomerBigRegion == "Miền Bắc" &&
                                    coreStructure.dicCustomer[x.CustomerId].CustomerType == "VM")
                                .Sum(x => x.QuantityOrderKg);

                        var DemandSouth = !POSouth.ContainsKey(CurrentProduct)
                            ? 0
                            : POSouth[CurrentProduct].Keys
                                .Where(x =>
                                    coreStructure.dicCustomer[x.CustomerId].CustomerBigRegion == "Miền Nam" &&
                                    PriorityTarget != ""
                                        ? coreStructure.dicCustomer[x.CustomerId].CustomerType == PriorityTarget
                                        : true)
                                .Sum(x => x.QuantityOrderKg);

                        var DemandSouthVM = !PriorityTarget.Contains("VM+")
                            ? 0
                            : PONorth[CurrentProduct].Keys
                                .Where(x =>
                                    coreStructure.dicCustomer[x.CustomerId].CustomerBigRegion == "Miền Nam" &&
                                    coreStructure.dicCustomer[x.CustomerId].CustomerType == "VM")
                                .Sum(x => x.QuantityOrderKg);

                        // Total Missing. Customers' Regions
                        var MissingNorth = DemandNorth - SupplyNorth.Values.Sum(x => x.QuantityForecast);
                        var MissingSouth = DemandSouth - SupplySouth.Values.Sum(x => x.QuantityForecast);

                        var QtyNorthNoXRegion = SupplyNorth.Values.Where(x => !x.CrossRegion).Sum(x => x.QuantityForecast);
                        var QtyNorthXRegion = SupplyNorth.Values.Where(x => x.CrossRegion).Sum(x => x.QuantityForecast);

                        var QtySouthNoXRegion = SupplySouth.Values.Where(x => !x.CrossRegion).Sum(x => x.QuantityForecast);
                        var QtySouthXRegion = SupplySouth.Values.Where(x => x.CrossRegion).Sum(x => x.QuantityForecast);

                        // Credit goes to someone very special, for figuring out the entire logic, the simplest way.
                        // Made by her. Hah!
                        var QtySouthCanSpare = Math.Min(Math.Max(QtySouthNoXRegion + QtySouthXRegion - DemandSouth, 0),
                            QtySouthXRegion);
                        var QtyNorthCanSpare = Math.Min(Math.Max(QtySouthNoXRegion + QtySouthXRegion - DemandSouth, 0),
                            QtySouthXRegion);

                        var QtyHighland = SupplyHighland.Values.Sum(x => x.QuantityForecast);

                        var _ProductCrossRegion = new ProductCrossRegion();
                        var flagNoHighlandToNorth = true;
                        if (coreStructure.dicProductCrossRegion.TryGetValue(CurrentProduct.ProductId,
                            out _ProductCrossRegion))
                            if (!_ProductCrossRegion.ToNorth)
                                flagNoHighlandToNorth = false;
                        var QtyHighlandToNorth = flagNoHighlandToNorth ? QtyHighland : 0;

                        var RateNorth =
                            (QtyNorthNoXRegion + QtyNorthXRegion +
                             QtyHighlandToNorth * (MissingNorth / (MissingNorth + MissingSouth)) + QtySouthCanSpare)
                            / DemandNorth;

                        var RateNorthWithVM =
                            (QtyNorthNoXRegion + QtyNorthXRegion +
                             QtyHighlandToNorth * (MissingNorth / (MissingNorth + MissingSouth)) + QtySouthCanSpare)
                            / (DemandNorth + DemandNorthVM);

                        if (RateNorthWithVM < 1)
                            RateNorth = 1;

                        RateNorth = Math.Min(RateNorth, UpperCap);

                        var RateSouth =
                            (QtySouthNoXRegion + QtySouthXRegion +
                             QtyHighland * (MissingSouth / (MissingNorth + MissingSouth)) + QtyNorthCanSpare)
                            / DemandSouth;

                        var RateSouthWithVM =
                            (QtySouthNoXRegion + QtySouthXRegion +
                             QtyHighland * (MissingSouth / (MissingNorth + MissingSouth)) + QtyNorthCanSpare)
                            / (DemandSouth + DemandSouthVM);

                        if (RateSouthWithVM < 1)
                            RateSouth = 1;

                        RateSouth = Math.Min(RateSouth, UpperCap);
                    }
                }
        }

        private void CoordDoWhile(CoordStructure coreStructure, string SupplierRegion, string CustomerRegion,
            string SupplierType, byte dayBefore, byte dayLdBefore, double UpperLimit = 1, bool CrossRegion = false,
            string PriorityTarget = "", bool YesNoByUnit = false, bool YesNoContracted = false, bool YesNoKPI = false)
        {
            try
            {
                /// <* IMPORTANTO! *>
                // Nothing shall begin before this happens
                var stopwatch = Stopwatch.StartNew();

                #region Preparing.

                #endregion

                // PO Date Layer.
                //Console.Write("{0} => {1}, {2}{3}", String.Concat(SupplierRegion.Split(' ').Select(x => x.First())), String.Concat(CustomerRegion.Split(' ').Select(x => x.First().ToString().ToUpper())), SupplierType, (PriorityTarget != "" ? " " + PriorityTarget : ""));
                foreach (var DatePO in coreStructure.dicPO.Keys.OrderByDescending(x => x.Date).Reverse())
                    // Product Layer.
                    foreach (var _Product in coreStructure.dicPO[DatePO].Keys.OrderByDescending(x => x.ProductCode)
                        .Reverse())
                    {
                        double _MOQ = 0;
                        // In case they are ordering and checking performance through an unit that's NOT FUCKING KILOGRAM!
                        //if (YesNoByUnit)
                        //{
                        //    // Cheapest way to calculate Kg per Unit.
                        //    // Man I'm so smart.
                        //    _MOQ = _CustomerOrder.QuantityOrderKg / _CustomerOrder.QuantityOrder;
                        //}
                        // ... Otherwise, we're cool boys.
                        //else
                        //{

                        _MOQ = coreStructure.dicMinimum[_Product.ProductCode.Substring(0, 1)];
                        // Special cases for Lemon. Apparently it's not Fruit but Spices :\
                        if (_Product.ProductCode.Substring(0, 1) == "K" &&
                            (_Product.ProductCode == "K01901" || _Product.ProductCode == "K02201"))
                            _MOQ = 0.3;

                        //}

                        restartThis:

                        /// <! For Debuging Purposes Only !>
                        // Only uncomment in very specific debugging situation.
                        //if (_Product.ProductCode == "A04801" && DatePO.Day == 26 && CustomerRegion == "Miền Nam" && SupplierRegion == "Miền Nam" && SupplierType == "VCM")
                        //{
                        //    string WhatAmIEvenDoing = "I have no freaking idea.";
                        //}

                        // Skip if product is not in the List VinEco supplies.
                        if (SupplierType != "VinEco" && _Product.ProductCode.Substring(0, 1) != "K" &&
                            (PriorityTarget == "VM" || PriorityTarget == "VM+"))
                            if (!_Product.ProductNote.Contains(CustomerRegion == "Miền Bắc" ? "North" : "South"))
                                continue;

                        // Dealing with cases of some Products that will not go to either region, from Lâm Đồng
                        var _ProductCrossRegion = new ProductCrossRegion();
                        if (coreStructure.dicProductCrossRegion.TryGetValue(_Product.ProductId, out _ProductCrossRegion) &&
                            SupplierRegion == "Lâm Đồng")
                            switch (CustomerRegion)
                            {
                                case "Miền Bắc":
                                    if (!_ProductCrossRegion.ToNorth) continue;
                                    break;
                                case "Miền Nam":
                                    if (!_ProductCrossRegion.ToSouth) continue;
                                    break;
                                default: break;
                            }

                        #region Demand from Chosen Customers.

                        // Total Order.
                        var sumVCM = coreStructure.dicPO[DatePO][_Product]
                            .Where(x =>
                                coreStructure.dicCustomer[x.Key.CustomerId].CustomerBigRegion == CustomerRegion &&
                                x.Value &&
                                (PriorityTarget != ""
                                    ? coreStructure.dicCustomer[x.Key.CustomerId].CustomerType == PriorityTarget
                                    : true))
                            .Sum(x => x.Key.QuantityOrderKg); // Sum of Demand.

                        var sumVM = PriorityTarget.Contains("VM+")
                            ? coreStructure.dicPO[DatePO][_Product]
                                .Where(x =>
                                    coreStructure.dicCustomer[x.Key.CustomerId].CustomerBigRegion == CustomerRegion &&
                                    x.Value && coreStructure.dicCustomer[x.Key.CustomerId].CustomerType.Contains("VM") &&
                                    !coreStructure.dicCustomer[x.Key.CustomerId].CustomerType.Contains("VM+"))
                                .Sum(x => x.Key.QuantityOrderKg)
                            : 0; // Sum of Demand.

                        var sumVcmMN = sumVCM + sumVM;

                        if (SupplierRegion == "Lâm Đồng")
                        {
                            var _DatePO = CustomerRegion == "Miền Nam" ? DatePO.AddDays(2) : DatePO.AddDays(-2);
                            if (coreStructure.dicPO.ContainsKey(_DatePO) &&
                                coreStructure.dicPO[_DatePO].ContainsKey(_Product))
                            {
                                var _CustomerRegion = CustomerRegion == "Miền Nam" ? "Miền Bắc" : "Miền Nam";
                                sumVCM += coreStructure.dicPO[_DatePO][_Product]
                                    .Where(x =>
                                        coreStructure.dicCustomer[x.Key.CustomerId].CustomerBigRegion == _CustomerRegion &&
                                        x.Value &&
                                        (PriorityTarget != ""
                                            ? coreStructure.dicCustomer[x.Key.CustomerId].CustomerType == PriorityTarget
                                            : true))
                                    .Sum(x => x.Key.QuantityOrderKg);

                                sumVM += PriorityTarget.Contains("VM+")
                                    ? coreStructure.dicPO[_DatePO][_Product]
                                        .Where(x =>
                                            coreStructure.dicCustomer[x.Key.CustomerId].CustomerBigRegion ==
                                            _CustomerRegion &&
                                            x.Value &&
                                            coreStructure.dicCustomer[x.Key.CustomerId].CustomerType.Contains("VM") &&
                                            !coreStructure.dicCustomer[x.Key.CustomerId].CustomerType.Contains("VM+"))
                                        .Sum(x => x.Key.QuantityOrderKg)
                                    : 0; // Sum of Demand.
                            }
                        }

                        #endregion

                        // To deal with Minimum Order Quantity.
                        double wallet = 0;

                        // Grabbing Suppliers by Harvest days.
                        // One for all, one for Lâm Đồng coz Suppliers from there supply both regions.

                        var _dicProductFC = coreStructure.dicFC.Where(x => x.Key.Date == DatePO.AddDays(-dayBefore))
                            .FirstOrDefault();
                        var _dicProductFcLd = coreStructure.dicFC.Where(x => x.Key.Date == DatePO.AddDays(-dayLdBefore))
                            .FirstOrDefault();

                        if (sumVCM != 0 && _dicProductFC.Value != null)
                        {
                            double sumThuMuaLd = 0;
                            double sumFarmLd = 0;

                            #region Supply from Lâm Đồng

                            if (SupplierRegion != "Lâm Đồng" && _dicProductFcLd.Value != null)
                            {
                                // Check if Inventory has stock in other places.
                                // If no, equally distributed stuff.
                                // If yes, hah hah hah no.
                                var dicSupplierLdFC = _dicProductFcLd.Value
                                    .Where(x => x.Key.ProductCode == _Product.ProductCode).FirstOrDefault();
                                if (dicSupplierLdFC.Value != null)
                                {
                                    // Check Lâm Đồng
                                    // Please NEVER FullOrder == true.
                                    //var _SupplierThuMuaLd = 

                                    var _dicSupplierLdFC = dicSupplierLdFC.Value
                                        .Where(x =>
                                            coreStructure.dicSupplier[x.Key.SupplierId].SupplierRegion == "Lâm Đồng" &&
                                            (x.Key.Target == "All" || x.Key.Target == PriorityTarget) &&
                                            (YesNoKPI
                                                ? x.Key.QuantityForecastPlanned
                                                : YesNoContracted
                                                    ? x.Key.QuantityForecastContracted
                                                    : x.Key.QuantityForecast) > 0);

                                    // Normal case
                                    sumFarmLd = _dicSupplierLdFC
                                        .Where(x =>
                                            coreStructure.dicSupplier[x.Key.SupplierId].SupplierType == "VinEco")
                                        .Sum(x => x.Key.QuantityForecast);

                                    sumThuMuaLd = _dicSupplierLdFC
                                        .Where(x =>
                                            coreStructure.dicSupplier[x.Key.SupplierId].SupplierType != "VinEco" &&
                                            x.Key.Availability.Contains(
                                                Convert.ToString((int)DatePO.AddDays(-dayLdBefore).DayOfWeek + 1)))
                                        .Sum(x => x.Key.QuantityForecast);
                                }
                            }

                            #endregion

                            var dicSupplierFC = _dicProductFC.Value.Where(x => x.Key.ProductCode == _Product.ProductCode)
                                .FirstOrDefault();
                            if (dicSupplierFC.Value != null)
                            {
                                #region Total Supply.

                                var _resultSupplier = dicSupplierFC.Value
                                    .Where(x =>
                                        coreStructure.dicSupplier[x.Key.SupplierId].SupplierRegion == "VinEco" &&
                                        coreStructure.dicSupplier[x.Key.SupplierId].SupplierType == SupplierType &&
                                        (x.Key.Target == "All" || x.Key.Target == PriorityTarget) &&
                                        (SupplierType != "VinEco"
                                            ? x.Key.Availability.Contains(
                                                Convert.ToString((int)DatePO.AddDays(-dayBefore).DayOfWeek + 1))
                                            : true));

                                var _dicSupplierFC = dicSupplierFC.Value
                                    .Where(x =>
                                        coreStructure.dicSupplier[x.Key.SupplierId].SupplierRegion == SupplierRegion &&
                                        (x.Key.Target == "All" || x.Key.Target == PriorityTarget) &&
                                        (YesNoKPI
                                            ? x.Key.QuantityForecastPlanned
                                            : YesNoContracted
                                                ? x.Key.QuantityForecastContracted
                                                : x.Key.QuantityForecast) > 0);

                                var sumFarm = _dicSupplierFC
                                    .Where(x =>
                                        coreStructure.dicSupplier[x.Key.SupplierId].SupplierType == "VinEco")
                                    .Sum(x => x.Key.QuantityForecast);

                                var sumThuMua = _dicSupplierFC
                                    .Where(x =>
                                        coreStructure.dicSupplier[x.Key.SupplierId].SupplierType != "VinEco" &&
                                        x.Key.Availability.Contains(
                                            Convert.ToString((int)DatePO.AddDays(-dayBefore).DayOfWeek + 1)))
                                    .Sum(x => x.Key.QuantityForecast);

                                //_resultSupplier
                                //    .Sum(x => YesNoKPI ? x.Key.QuantityForecastPlanned : YesNoContracted ? x.Key.QuantityForecastContracted : x.Key.QuantityForecast);

                                var flagFullOrder = false;

                                var sumVE = sumFarm + sumThuMua;

                                var _DatePO = SupplierRegion == "Miền Bắc"
                                    ? DatePO.AddDays(-2).Date
                                    : DatePO.AddDays(2).Date;
                                if (CustomerRegion == "Miền Nam" && coreStructure.dicPO.ContainsKey(_DatePO) &&
                                    coreStructure.dicPO[_DatePO].ContainsKey(_Product))
                                    sumVE += Math.Max(sumFarmLd + sumThuMuaLd - coreStructure.dicPO[_DatePO][_Product]
                                                          .Where(x =>
                                                              coreStructure.dicCustomer[x.Key.CustomerId]
                                                                  .CustomerBigRegion ==
                                                              (CustomerRegion == "Miền Bắc" ? "Miền Nam" : "Miền Bắc") &&
                                                              x.Value)
                                                          .Sum(x => x.Key.QuantityOrderKg), 0);
                                else
                                    sumVE += sumFarmLd + sumThuMuaLd;

                                if (_resultSupplier.Where(x => YesNoKPI || YesNoContracted ? false : x.Key.FullOrder)
                                        .FirstOrDefault().Key != null)
                                    flagFullOrder = true;
                                //else
                                //{
                                //sumVE = _resultSupplier
                                //    .Sum(x => YesNoKPI ? x.Key.QuantityForecastPlanned : YesNoContracted ? x.Key.QuantityForecastContracted : x.Key.QuantityForecast);  // Sum of Supply
                                //sumVE = sumFarm + sumThuMua + sumFarmLd + sumThuMuaLd;
                                //}

                                #endregion

                                if (sumVE > 0)
                                {
                                    #region Rate.

                                    // Hack - Freaking need to dissect this part.
                                    // Todo - Further Optimization.

                                    // For fuck sake, this is the hardest to code part.
                                    // Also very important. Too important.

                                    // Rate = Supply / Demand --> Deli = Demand * Rate.
                                    var rate = sumVE / (sumVCM + sumVM);

                                    // If Screw-the-upper-limit flag is up.
                                    if (flagFullOrder)
                                        rate = UpperCap;
                                    // If it's VinCommerce's Supplier, always 1.
                                    else if (rate < 1 && SupplierType == "VCM" && sumVE > 0)
                                        rate = UpperCap;
                                    // Otherwise, in case of an UpperLimit, obey it
                                    else if (!flagFullOrder)
                                        if (rate < 1)
                                        {
                                            rate = Math.Max(sumVE / sumVCM, 1);
                                            rate = SupplierRegion != "Lâm Đồng" &&
                                                   (YesNoKPI || sumFarm > 0 || sumFarmLd > 0 || sumThuMua > 0 ||
                                                    sumThuMuaLd > 0)
                                                ? Math.Max(rate, 1)
                                                : rate;
                                            //if (SupplierRegion == "Lâm Đồng" && rate < 1) { rate = sumVE / sumVcmMN; }
                                        }
                                        else if (rate > 1)
                                        {
                                            //if (sumVcmMN > sumVCM)
                                            //{
                                            //    rate = 1;
                                            //}
                                            /*else */
                                            //if ((sumFarm + sumFarmLd + sumThuMua + sumThuMuaLd) / (sumVCM + sumVM) > 1)
                                            //{
                                            rate = (sumFarm + sumFarmLd + sumThuMua + sumThuMuaLd) / (sumVCM + sumVM);
                                            rate = SupplierRegion != "Lâm Đồng" &&
                                                   (YesNoKPI || sumFarm > 0 || sumFarmLd > 0 || sumThuMua > 0 ||
                                                    sumThuMuaLd > 0)
                                                ? Math.Max(rate, 1)
                                                : rate;
                                            if (rate < 1 && SupplierType == "VCM" && sumVE > 0)
                                                rate = UpperCap;
                                            //}
                                        }
                                    rate = UpperLimit > 0 ? Math.Min(rate, UpperLimit) : rate;

                                    #endregion

                                    // Only the bravest would tread deeper.
                                    // ... I was once young, brave and foolish ...

                                    // Optimization - Filtering Customer Orders that has been dealt with.
                                    //var ListCustomerOrder = coreStructure.dicPO[DatePO][_Product].Where(x => x.Value == true).ToDictionary(x => x.Key);
                                    var ValidCustomerList = coreStructure.dicPO[DatePO][_Product].Where(x => x.Value)
                                        .ToDictionary(x => x.Key).Keys
                                        .Where(x =>
                                            x.QuantityOrderKg >= 0.1 &&
                                            coreStructure.dicCustomer[x.CustomerId].CustomerBigRegion == CustomerRegion &&
                                            (PriorityTarget != ""
                                                ? coreStructure.dicCustomer[x.CustomerId].CustomerType == PriorityTarget
                                                : true) &&
                                            (x.DesiredRegion == null ? true : x.DesiredRegion == SupplierRegion) &&
                                            (x.DesiredSource == null ? true : x.DesiredSource == SupplierType))
                                        .OrderByDescending(x => coreStructure.dicCustomer[x.CustomerId].CustomerCode)
                                        //.Reverse()
                                        .ToList();

                                    do
                                    {
                                        #region Qualified Suppliers.

                                        SupplierForecast _SupplierForecast = null;

                                        var _dicSupplierFC_inner = dicSupplierFC.Value
                                            .Where(x => x.Key.QuantityForecast >= _MOQ)
                                            .Where(x => coreStructure.dicSupplier[x.Key.SupplierId].SupplierRegion ==
                                                        SupplierRegion &&
                                                        coreStructure.dicSupplier[x.Key.SupplierId].SupplierType ==
                                                        SupplierType &&
                                                        (SupplierType != "VinEco"
                                                            ? x.Key.Availability.Contains(
                                                                Convert.ToString(
                                                                    (int)DatePO.AddDays(-dayBefore).DayOfWeek + 1))
                                                            : true) &&
                                                        (x.Key.Target == "All" || x.Key.Target == PriorityTarget) &&
                                                        (CrossRegion ? x.Key.CrossRegion : true))
                                            .OrderBy(x => x.Key.Level)
                                            .ThenByDescending(x => x.Key.FullOrder)
                                            .ThenBy(x => coreStructure.dicDeli[DatePO.AddDays(-dayBefore)][_Product][x.Key])
                                            .ThenByDescending(x => x.Key.QuantityForecast)
                                            .ThenByDescending(x => x.Key.LabelVinEco).ToDictionary(x => x.Key);

                                        var result = _dicSupplierFC_inner.FirstOrDefault();
                                        if (result.Key == null)
                                            break;

                                        var _CustomerOrder = ValidCustomerList
                                            .Where(x => x.QuantityOrderKg * rate <= result.Key.QuantityForecast)
                                            .FirstOrDefault();

                                        if (_CustomerOrder == null)
                                            _CustomerOrder = ValidCustomerList.OrderBy(x => x.QuantityOrderKg)
                                                .FirstOrDefault();

                                        if (_CustomerOrder == null)
                                            break;

                                        // Coz for fuck sake, it can return null

                                        var totalSupplier = _dicSupplierFC_inner.Count();
                                        _SupplierForecast = result.Key;

                                        #endregion

                                        var _rate = rate;
                                        if (coreStructure.dicPO[DatePO][_Product].Count <= totalSupplier) _rate = 1;

                                        if (_SupplierForecast != null)
                                        {
                                            Dictionary<Product, Dictionary<CustomerOrder,
                                                Dictionary<SupplierForecast, DateTime>>> _dicCoordProduct = null;

                                            if (coreStructure.dicCoord.TryGetValue(DatePO, out _dicCoordProduct))
                                            {
                                                Dictionary<CustomerOrder, Dictionary<SupplierForecast, DateTime>>
                                                    _dicCoordCusSup = null;
                                                if (_dicCoordProduct.TryGetValue(_Product, out _dicCoordCusSup))
                                                {
                                                    Dictionary<SupplierForecast, DateTime> _SupplierForecastCoord = null;
                                                    if (_dicCoordCusSup.TryGetValue(_CustomerOrder,
                                                            out _SupplierForecastCoord) && _SupplierForecastCoord == null)
                                                    {
                                                        wallet +=
                                                        (YesNoKPI || YesNoContracted
                                                            ? false
                                                            : _SupplierForecast.FullOrder)
                                                            ? _CustomerOrder.QuantityOrderKg
                                                            : Math.Round(_CustomerOrder.QuantityOrderKg * _rate, 1);

                                                        #region MOQ.

                                                        if (wallet < _MOQ &&
                                                            (YesNoKPI
                                                                ? _SupplierForecast.QuantityForecastPlanned
                                                                : (YesNoContracted
                                                                    ? _SupplierForecast.QuantityForecastContracted
                                                                    : _SupplierForecast.QuantityForecast)) >= _MOQ)
                                                            wallet = _MOQ;

                                                        //if (_MOQ == 0.05)
                                                        //{
                                                        //    // Let's hope this will never be hit.
                                                        //    // I fucking do hope that.
                                                        //    string OhMyFuckingGodWhy = "Holy shit idk, why, oh god, why";
                                                        //}

                                                        #endregion

                                                        if (wallet < _MOQ && PriorityTarget != "") wallet = _MOQ;

                                                        if (wallet >= _MOQ && _SupplierForecast.QuantityForecast >= _MOQ)
                                                        {
                                                            //if (sumVE <= 0) { continue; }
                                                            // Honestly, this should never be hit
                                                            // Jk I changed stuff. This should ALWAYS be hit
                                                            _SupplierForecastCoord =
                                                                new Dictionary<SupplierForecast, DateTime>();

                                                            var _QuantityForecast = Math.Min(wallet,
                                                                _SupplierForecast.QuantityForecast);

                                                            if (YesPlanningFuckMe)
                                                                _QuantityForecast =
                                                                    Math.Min(Math.Max(wallet / totalSupplier, _MOQ),
                                                                        _SupplierForecast.QuantityForecast);

                                                            if (UpperCap > 0)
                                                                _QuantityForecast =
                                                                    Math.Min(_CustomerOrder.QuantityOrderKg * UpperLimit,
                                                                        _QuantityForecast);

                                                            _QuantityForecast = Math.Round(_QuantityForecast, 1);

                                                            #region Unit.

                                                            if (_CustomerOrder.Unit != "Kg")
                                                            {
                                                                var something = coreStructure
                                                                    .dicProductUnit[_Product.ProductCode].ListRegion
                                                                    .FirstOrDefault(x =>
                                                                        x.OrderUnitType == _CustomerOrder.Unit);
                                                                if (something != null)
                                                                {
                                                                    var _SaleUnitPer = something.SaleUnitPer;
                                                                    _QuantityForecast =
                                                                        _QuantityForecast / _MOQ * _SaleUnitPer;
                                                                }
                                                            }

                                                            #endregion

                                                            #region Defer extra days for Crossing Regions ( North --> South and vice versa. )

                                                            // To coup with merging PO ( Tue Thu Sat to Mon Wed Fri )
                                                            var _Date = DatePO.AddDays(-dayBefore).Date;
                                                            if (CrossRegion && _SupplierForecast.CrossRegion &&
                                                                CustomerRegion == "Miền Bắc" &&
                                                                SupplierRegion ==
                                                                "Miền Nam" /*&& _Product.ProductCode.Substring(0, 1) == "K"*/ &&
                                                                (_Date.DayOfWeek == DayOfWeek.Tuesday ||
                                                                 _Date.DayOfWeek == DayOfWeek.Thursday ||
                                                                 _Date.DayOfWeek == DayOfWeek.Saturday))
                                                                _Date = _Date.AddDays(-1).Date;

                                                            #endregion

                                                            // To coup with Supply has custom rates, depending on Region.
                                                            var _ProductRate = new ProductRate();
                                                            double _Rate = 1;
                                                            if (!YesNoKPI && SupplierRegion == "Miền Nam" &&
                                                                coreStructure.dicProductRate.TryGetValue(
                                                                    _Product.ProductCode, out _ProductRate))
                                                                switch (CustomerRegion)
                                                                {
                                                                    case "Miền Bắc":
                                                                        _Rate = _ProductRate.ToNorth;
                                                                        break;
                                                                    case "Miền Nam":
                                                                        _Rate = _ProductRate.ToSouth;
                                                                        break;
                                                                    default: break;
                                                                }

                                                            var newId = Guid.NewGuid();
                                                            _SupplierForecastCoord.Add(new SupplierForecast
                                                            {
                                                                _id = newId,
                                                                SupplierForecastId = newId,

                                                                SupplierId = _SupplierForecast.SupplierId,
                                                                LabelVinEco = _SupplierForecast.LabelVinEco,
                                                                FullOrder = _SupplierForecast.FullOrder,
                                                                QualityControlPass = _SupplierForecast.QualityControlPass,
                                                                CrossRegion = _SupplierForecast.CrossRegion,
                                                                Level = _SupplierForecast.Level,
                                                                Availability = _SupplierForecast.Availability,
                                                                Target = _SupplierForecast.Target,

                                                                QuantityForecast = _QuantityForecast
                                                            }, _Date);

                                                            // KPI cases
                                                            if (YesNoKPI)
                                                            {
                                                                _SupplierForecast.QuantityForecastPlanned -=
                                                                    _QuantityForecast;
                                                                _SupplierForecast.QuantityForecastContracted -=
                                                                    _QuantityForecast;
                                                            }
                                                            // Minimum cases
                                                            else if (YesNoContracted)
                                                            {
                                                                _SupplierForecast.QuantityForecastContracted -=
                                                                    _QuantityForecast;
                                                            }
                                                            // Default cases
                                                            _SupplierForecast.QuantityForecast -= _QuantityForecast;
                                                            _SupplierForecast.QuantityForecastOriginal -= _QuantityForecast;
                                                            if (!_SupplierForecast.FullOrder &&
                                                                _SupplierForecast.QuantityForecast <= 0)
                                                                _SupplierForecast.QuantityForecast = _MOQ * 7;
                                                            // To make sure Full Order Supplier will still go.

                                                            // Pretty sure I don't need to recalculate sumVCM here anymore.
                                                            // Only sumVE matters here, to trigger a break.
                                                            // Then again even that is not really needed.
                                                            //sumVCM -= _CustomerOrder.QuantityOrder;
                                                            //sumVE -= !YesNoContracted && !YesNoKPI && _SupplierForecast.FullOrder ? 0 : _QuantityForecast;

                                                            //// Recalculating _rate - Unneccesary here I think.
                                                            //_rate = sumVCM <= 0 ? 0 : Math.Min(sumVCM != 0 ? sumVE / sumVCM : 0, UpperLimit);
                                                            //_rate = sumVCM <= 0 ? 0 : ((SupplierType == "VinEco") && (sumVEThuMua > 0) ? Math.Max(_rate, 1) : _rate);

                                                            coreStructure.dicCoord[DatePO][_Product][_CustomerOrder] =
                                                                _SupplierForecastCoord;
                                                            coreStructure.dicDeli[DatePO.AddDays(-dayBefore)][_Product][
                                                                _SupplierForecast] += wallet;

                                                            //coreStructure.dicPO[DatePO][_Product][_CustomerOrder] = false;

                                                            // Roburst way, might optimize Procedures a little bit better.
                                                            // Remove Customers and Suppliers fulfilled their roles.

                                                            if (YesPlanningFuckMe && _CustomerOrder.QuantityOrder >=
                                                                _QuantityForecast)
                                                            {
                                                                var CustomerOrder = new CustomerOrder();

                                                                CustomerOrder.Company = _CustomerOrder.Company;
                                                                CustomerOrder.CustomerId = _CustomerOrder.CustomerId;
                                                                CustomerOrder.CustomerOrderId = Guid.NewGuid();
                                                                CustomerOrder.DesiredRegion = _CustomerOrder.DesiredRegion;
                                                                CustomerOrder.DesiredSource = _CustomerOrder.DesiredSource;
                                                                CustomerOrder.QuantityOrder =
                                                                    _CustomerOrder.QuantityOrder - _QuantityForecast;
                                                                CustomerOrder.QuantityOrderKg =
                                                                    _CustomerOrder.QuantityOrderKg - _QuantityForecast;
                                                                CustomerOrder.Unit = _CustomerOrder.Unit;
                                                                CustomerOrder._id = CustomerOrder.CustomerOrderId;

                                                                _CustomerOrder.QuantityOrder =
                                                                    Math.Min(_CustomerOrder.QuantityOrder,
                                                                        _QuantityForecast);
                                                                _CustomerOrder.QuantityOrderKg =
                                                                    Math.Min(_CustomerOrder.QuantityOrderKg,
                                                                        _QuantityForecast);

                                                                coreStructure.dicPO[DatePO][_Product]
                                                                    .Add(CustomerOrder, true);

                                                                coreStructure.dicCoord[DatePO][_Product]
                                                                    .Add(CustomerOrder, null);

                                                                goto restartThis;
                                                            }

                                                            if (_SupplierForecast.QuantityForecast < _MOQ)
                                                            {
                                                                coreStructure.dicFC[DatePO.AddDays(-dayBefore)][_Product]
                                                                    .Remove(_SupplierForecast);
                                                                //_dicSupplierFC_inner.Remove(_SupplierForecast);
                                                                dicSupplierFC.Value.Remove(_SupplierForecast);
                                                            }

                                                            wallet -= _QuantityForecast;
                                                        }
                                                        coreStructure.dicPO[DatePO][_Product].Remove(_CustomerOrder);
                                                        //ListCustomerOrder.Remove(_CustomerOrder);
                                                        ValidCustomerList.Remove(_CustomerOrder);

                                                        if (coreStructure.dicPO[DatePO][_Product].Count == 0)
                                                            coreStructure.dicPO[DatePO].Remove(_Product);

                                                        if (coreStructure.dicPO[DatePO].Keys.Count == 0)
                                                            coreStructure.dicPO.Remove(DatePO);
                                                    }
                                                }
                                            }
                                        }
                                    } while (ValidCustomerList.Count > 0);
                                }
                            }
                        }
                    }
                //}
                stopwatch.Stop();
                WriteToRichTextBoxOutput(string.Format(" - Done in {0}s!",
                    Math.Round(stopwatch.Elapsed.TotalSeconds, 2)));
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
            }
        }

        /// <summary>
        ///     Reading PO from VCM
        /// </summary>
        private async Task UpdatePO(string fileNameMB, string fileNameMN)
        {
            //Console.OutputEncoding = System.Text.Encoding.UTF8;

            //Process[] processBefore = Process.GetProcessesByName("excel");
            //string extension = Path.GetExtension(filePath);
            var header = "YES";
            //string conStr, sheetName;

            // These are openned here so they could be closed / released even in the case of Exceptions
            // Open First Workbook
            //Excel.Application xlApp = new Excel.Application();
            //Excel.Workbook xlWb = xlApp.Workbooks.Open(filePath, 0, true, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            //Excel.Worksheet xlWs = xlWb.Worksheets[1];
            //Excel.Range xlRng = xlWs.UsedRange;

            //Excel.Application xlApp = new Excel.Application()
            //{
            //    ScreenUpdating = false,
            //    EnableEvents = false,
            //    DisplayAlerts = false,
            //    DisplayStatusBar = false,
            //    AskToUpdateLinks = false
            //};
            //Excel.Workbook xlWb = null;
            //Excel.Worksheet xlWs = null;
            //Excel.Range xlRng = null;

            //Process[] processAfter = Process.GetProcessesByName("excel");

            //int processID = 0;

            //foreach (Process process in processAfter)
            //{
            //    if (!processBefore.Select(p => p.Id).Contains(process.Id))
            //    {
            //        processID = process.Id;
            //        break;
            //    }
            //}

            try
            {
                //using (OleDbCommand oleCmd = new OleDbCommand())
                //{
                //    using (OleDbDataAdapter oleAdapt = new OleDbDataAdapter())
                //    {
                var mongoClient = new MongoClient();
                var db = mongoClient.GetDatabase("localtest");

                var
                    PO = new List<PurchaseOrderDate>(); // mongoClient.GetDatabase("localtest").GetCollection<PurchaseOrderDate>("PurchaseOrder").AsQueryable().ToList();
                var Product = db.GetCollection<Product>("Product").AsQueryable().ToList();
                var Customer = new List<Customer>(); //  db.GetCollection<Customer>("Customer").AsQueryable().ToList();

                var dicPO = new Dictionary<DateTime, Dictionary<string, Dictionary<string, Guid>>>(1000);
                var dicProduct = new Dictionary<string, Product>(1000);
                var dicCustomer = new Dictionary<string, Customer>(10000);

                // Product Dictionary.
                foreach (var _Product in Product)
                    if (!dicProduct.ContainsKey(_Product.ProductCode))
                        dicProduct.Add(_Product.ProductCode, _Product);

                // Customer Dictionary.
                foreach (var _Customer in Customer)
                    if (!dicCustomer.ContainsKey(_Customer.CustomerCode + _Customer.CustomerType))
                        dicCustomer.Add(_Customer.CustomerCode + _Customer.CustomerType, _Customer);

                var filePath =
                    $"D:\\Documents\\Stuff\\VinEco\\Mastah Project\\{fileNameMB}";

                var conStr = string.Format(Constants.Excel07ConString, filePath, header);

                var directoryPath = "D:\\Documents\\Stuff\\VinEco\\Mastah Project\\PO";

                #region Reading PO files in folder.

                var dirInfo = new DirectoryInfo(directoryPath);
                var ListFile = dirInfo.GetFiles();

                foreach (var _FileInfo in ListFile)
                {
                    var opt = new LoadOptions {MemorySetting = MemorySetting.MemoryPreference};
                    var xlWbAspose = new Workbook(_FileInfo.FullName, opt);
                    var xlWsAspose = xlWbAspose.Worksheets.OrderByDescending(x => x.Cells.MaxDataRow).First();

                    //xlWb = xlApp.Workbooks.Open(_FileInfo.FullName, false, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "", false, false, 0, true, false, false);

                    //xlWs = xlWb.Worksheets[1];
                    //xlRng = xlWs.UsedRange;

                    var _Region = "";
                    switch (_FileInfo.Name.Substring(0, 2))
                    {
                        case "MB":
                            _Region = "Miền Bắc";
                            break;
                        case "MN":
                            _Region = "Miền Nam";
                            break;
                        default: break;
                    }

                    WriteToRichTextBoxOutput(_FileInfo.Name, false); // + " - Done!");

                    var stopwatch = new Stopwatch();

                    stopwatch.Start();

                    EatPOAspose(PO: PO, 
                        xlWs: xlWsAspose,
                        conStr: string.Format(Constants.Excel07ConString, _FileInfo.FullName, header),
                        PORegion: _Region, 
                        dicPO: dicPO, 
                        dicProduct: dicProduct,
                        dicCustomer: dicCustomer,
                        Product: Product,
                        Customer: Customer,
                        YesNoNew: false);

                    stopwatch.Stop();

                    WriteToRichTextBoxOutput(
                        Message: $"- Done in {Math.Round(stopwatch.Elapsed.TotalSeconds, 2)}s!",
                        NewLine: false);
                    WriteToRichTextBoxOutput();

                    //EatPO(PO, xlRng, xlWs, string.Format(Constants.Excel07ConString, _FileInfo.FullName, header), _Region, dicPO, dicProduct, dicCustomer, Product, Customer, false);

                    //Marshal.ReleaseComObject(xlRng); xlRng = null;
                    //Marshal.ReleaseComObject(xlWs); xlWs = null;

                    //xlWb.Close(SaveChanges: false);
                    //Marshal.ReleaseComObject(xlWb); xlWb = null;

                    xlWsAspose = null;
                    xlWbAspose = null;
                }

                #endregion

                #region Reading PO files specially especially specifically specified.

                //// North PO
                //xlWb = xlApp.Workbooks.Open(filePath, UpdateLinks: false, ReadOnly: true, Format: 5, Password: "", WriteResPassword: "", IgnoreReadOnlyRecommended: true, Origin: Excel.XlPlatform.xlWindows, Delimiter: "", Editable: false, Notify: false, Converter: 0, AddToMru: true, Local: false, CorruptLoad: false);

                //xlWs = xlWb.Worksheets[1];
                //xlRng = xlWs.UsedRange;

                //EatPO(PO, xlRng, xlWs, conStr, "Miền Bắc", dicPO, dicProduct, dicCustomer, Product, Customer, true);

                //// South PO
                //filePath = string.Format("D:\\Documents\\Stuff\\VinEco\\Mastah Project\\{0}", fileNameMN);
                //conStr = string.Format(Constants.Excel07ConString, filePath, header);

                //xlWb = xlApp.Workbooks.Open(filePath, UpdateLinks: false, ReadOnly: true, Format: 5, Password: "", WriteResPassword: "", IgnoreReadOnlyRecommended: true, Origin: Excel.XlPlatform.xlWindows, Delimiter: "", Editable: false, Notify: false, Converter: 0, AddToMru: true, Local: false, CorruptLoad: false);

                //xlWs = xlWb.Worksheets[1];
                //xlRng = xlWs.UsedRange;

                //EatPO(PO, xlRng, xlWs, conStr, "Miền Nam", dicPO, dicProduct, dicCustomer, Product, Customer, false);

                //// Priority PO
                //filePath = string.Format("D:\\Documents\\Stuff\\VinEco\\Mastah Project\\{0}", "Forecast MB Priority.xlsx");
                //conStr = string.Format(Constants.Excel07ConString, filePath, header);

                //xlWb = xlApp.Workbooks.Open(filePath, UpdateLinks: false, ReadOnly: true, Format: 5, Password: "", WriteResPassword: "", IgnoreReadOnlyRecommended: true, Origin: Excel.XlPlatform.xlWindows, Delimiter: "", Editable: false, Notify: false, Converter: 0, AddToMru: true, Local: false, CorruptLoad: false);

                //xlWs = xlWb.Worksheets[1];
                //xlRng = xlWs.UsedRange;

                //EatPO(PO, xlRng, xlWs, conStr, "Miền Bắc", dicPO, dicProduct, dicCustomer, Product, Customer, false);

                #endregion

                WriteToRichTextBoxOutput("Here goes pain");

                db.DropCollection("PurchaseOrder");
                await db.GetCollection<PurchaseOrderDate>("PurchaseOrder").InsertManyAsync(PO);

                db.DropCollection("Product");
                await db.GetCollection<Product>("Product").InsertManyAsync(Product);

                db.DropCollection("Customer");
                await db.GetCollection<Customer>("Customer").InsertManyAsync(Customer);

                db = null;
                //    }
                //}
            }
            catch (Exception ex)
            {
                //WriteToRichTextBoxOutput(ex);
                throw ex;
                //MessageBox.Show(ex.Message, "Exception Error");
            }
            finally
            {
                #region Clean up

                // Cleanup
                //GC.Collect();
                //GC.WaitForPendingFinalizers();

                // Rule of thumb for releasing com objects:
                //   never use two dots, all COM objects must be referenced and released individually
                //   ex: [somthing].[something].[something] is bad

                //xlApp.ScreenUpdating = true;
                //xlApp.EnableEvents = true;
                //xlApp.DisplayAlerts = false;
                //xlApp.DisplayStatusBar = true;
                //xlApp.AskToUpdateLinks = true;

                //// Quit and release
                //xlApp.Quit();
                //if (xlApp != null) { Marshal.ReleaseComObject(xlApp); xlApp = null; }

                //GC.Collect();
                //GC.WaitForPendingFinalizers();

                //if (processID != 0)
                //{
                //    Process process = Process.GetProcessById(processID);
                //    process.Kill();
                //}

                #endregion
            }
        }

        /// <summary>
        ///     Reading Forecast
        /// </summary>
        private async Task UpdateFC(string fileVE, string fileTM, bool YesNoPlanning = false)
        {
            //Process[] processBefore = Process.GetProcessesByName("excel");
            //string extension = Path.GetExtension(filePath);
            var header = "YES";
            //string conStr, sheetName;

            // These are openned here so they could be closed / released even in the case of Exceptions
            //Excel.Application xlApp = new Excel.Application();
            //Excel.Workbook xlWb = null;
            //Excel.Worksheet xlWs = null;
            //Excel.Range xlRng = null;

            //xlApp.ScreenUpdating = false;
            //xlApp.EnableEvents = false;
            //xlApp.DisplayAlerts = false;
            //xlApp.DisplayStatusBar = false;
            //xlApp.AskToUpdateLinks = false;

            try
            {
                //Console.OutputEncoding = System.Text.Encoding.UTF8;

                #region Part of Read from Chosen file

                //conStr = string.Empty;
                //switch (extension)
                //{

                //    case ".xls": //Excel 97-03
                //        conStr = string.Format(Excel03ConString, filePath, header);
                //        break;

                //    case ".xlsx": //Excel 07
                //        conStr = string.Format(Excel07ConString, filePath, header);
                //        break;

                //    case ".xlsb": //Excel 07
                //        conStr = string.Format(Excel07ConString, filePath, header);
                //        break;
                //}

                ////Get the name of the First Sheet.
                //using (OleDbConnection con = new OleDbConnection(conStr))
                //{
                //    using (OleDbCommand cmd = new OleDbCommand())
                //    {
                //        cmd.Connection = con;
                //        con.Open();
                //        DataTable dtExcelSchema = con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                //        sheetName = dtExcelSchema.Rows[0]["TABLE_NAME"].ToString();
                //        con.Close();
                //    }
                //}

                #endregion

                // Connection
                //string conStr = string.Format(Excel07ConString, filePath, header);

                //using (OleDbCommand oleCmd = new OleDbCommand())
                //{
                //    using (OleDbDataAdapter oleAdapt = new OleDbDataAdapter())
                //    {
                //PurchaseOrder PO = new PurchaseOrder();
                //PO.PurchaseOrderCode = DateTime.Today.ToString();
                //PO.ListPurchaseOrderDate = new List<PurchaseOrderDate>();

                var FC = new List<ForecastDate>();

                var mongoClient = new MongoClient();
                var db = mongoClient.GetDatabase("localtest");

                var Product = mongoClient.GetDatabase("localtest").GetCollection<Product>("Product").AsQueryable()
                    .ToList();
                var Supplier =
                    new List<Supplier>(); // mongoClient.GetDatabase("localtest").GetCollection<Supplier>("Supplier").AsQueryable().ToList();

                //var mongoClient = new MongoClient().GetDatabase("localtest");
                //var mongoCollection = mongoClient.GetCollection<PurchaseOrderDate>("PO");
                //var PO = mongoCollection.AsQueryable().ToList();

                #region Part of Read from Chosen file

                //oleCmd.CommandText = "SELECT * From [" + sheetName + "]";
                //oleCmd.Connection = oleCon;
                //oleCon.Open();
                //oleAdapt.SelectCommand = oleCmd;
                //oleAdapt.Fill(dt);
                //oleCon.Close();

                // Destination DataTable
                //DataTable database = new DataTable();

                //database.Columns.Add("VE Code", typeof(string));
                //database.Columns.Add("VE Code 6 số mới", typeof(string));
                //database.Columns.Add("VE Name", typeof(string));
                //database.Columns.Add("Unit", typeof(string));
                //database.Columns.Add("StoreCode", typeof(string));
                //database.Columns.Add("StoreName", typeof(string));
                //database.Columns.Add("StoreType", typeof(string));
                //database.Columns.Add("Region", typeof(string));
                //database.Columns.Add("SupplierRegion", typeof(string));
                //database.Columns.Add("PO Region", typeof(string));
                //database.Columns.Add("OrderDate", typeof(DateTime));
                //database.Columns.Add("OrderQuantity", typeof(double));

                #endregion

                #region Main body.

                var filePath = "";
                var conStr = "";

                var defaultPath = "D:\\Documents\\Stuff\\VinEco\\Mastah Project\\{0}";
                filePath = string.Format(defaultPath, fileVE);
                conStr = string.Format(Constants.Excel07ConString, filePath, header);

                var dicFC = new Dictionary<DateTime, Dictionary<string, Dictionary<string, Guid>>>();
                var dicProduct = new Dictionary<string, Product>();
                var dicSupplier = new Dictionary<string, Supplier>();

                foreach (var _Product in Product)
                {
                    Product _product = null;
                    if (!dicProduct.TryGetValue(_Product.ProductCode, out _product))
                        dicProduct.Add(_Product.ProductCode, _Product);
                }

                foreach (var _Supplier in Supplier)
                {
                    Supplier _supplier = null;
                    if (!dicSupplier.TryGetValue(_Supplier.SupplierCode, out _supplier))
                        dicSupplier.Add(_Supplier.SupplierCode, _Supplier);
                }

                var fileName = "";
                var listFcFileName = new Dictionary<string, Dictionary<string, bool>>();

                fileName = YesNoPlanning ? "DBSL Planning.xlsb" : fileVE;
                listFcFileName.Add(fileName, new Dictionary<string, bool>());
                listFcFileName[fileName].Add("VinEco", false);

                fileName = fileTM;
                listFcFileName.Add(fileName, new Dictionary<string, bool>());
                listFcFileName[fileName].Add("ThuMua", false);

                if (!YesNoPlanning)
                {
                    fileName = "ThuMua KPI.xlsb";
                    listFcFileName.Add(fileName, new Dictionary<string, bool>());
                    listFcFileName[fileName].Add("VinEco", true);
                }

                foreach (var _fileName in listFcFileName.Keys)
                {
                    var xlWbAspose = new Workbook(string.Format(defaultPath, _fileName));
                    var xlWsAspose = xlWbAspose.Worksheets[0];

                    WriteToRichTextBoxOutput(_fileName, false);
                    EatForecastAspose(FC, xlWsAspose, listFcFileName[_fileName].Keys.First(), dicFC, dicProduct,
                        dicSupplier, Product, Supplier, listFcFileName[_fileName].Values.First());
                    WriteToRichTextBoxOutput(" - Done!");
                }

                #region old stuff

                //#region Forecast Farm
                //xlWb = xlApp.Workbooks.Open(filePath,
                //    UpdateLinks: false,
                //    ReadOnly: true,
                //    Format: 5,
                //    Password: "",
                //    WriteResPassword: "",
                //    IgnoreReadOnlyRecommended: true,
                //    Origin: Excel.XlPlatform.xlWindows,
                //    Delimiter: "",
                //    Editable: false,
                //    Notify: false,
                //    Converter: 0,
                //    AddToMru: true,
                //    Local: false,
                //    CorruptLoad: false);

                //xlApp.Calculation = Excel.XlCalculation.xlCalculationManual;

                //xlWs = xlWb.Worksheets[1];
                //xlRng = xlWs.UsedRange;

                //EatForecast(
                //    FC: FC,
                //    xlRng: xlRng, xlWs: xlWs,
                //    conStr: conStr,
                //    SupplierType: "VinEco",
                //    dicFC: dicFC,
                //    dicProduct: dicProduct,
                //    dicSupplier: dicSupplier,
                //    Product: Product,
                //    Supplier: Supplier);

                //xlWb.Close(SaveChanges: false);
                //#endregion

                //#region Forecast ThuMua
                //filePath = string.Format("D:\\Documents\\Stuff\\VinEco\\Mastah Project\\{0}", fileTM);
                //conStr = string.Format(Constants.Excel07ConString, filePath, header);

                //xlWb = xlApp.Workbooks.Open(filePath,
                //    UpdateLinks: false,
                //    ReadOnly: true,
                //    Format: 5,
                //    Password: "",
                //    WriteResPassword: "",
                //    IgnoreReadOnlyRecommended: true,
                //    Origin: Excel.XlPlatform.xlWindows,
                //    Delimiter: "",
                //    Editable: false,
                //    Notify: false,
                //    Converter: 0,
                //    AddToMru: true,
                //    Local: false,
                //    CorruptLoad: false);

                //xlApp.Calculation = Excel.XlCalculation.xlCalculationManual;

                //xlWs = xlWb.Worksheets[1];
                //xlRng = xlWs.UsedRange;

                //EatForecast(
                //    FC: FC,
                //    xlRng: xlRng, xlWs: xlWs,
                //    conStr: conStr,
                //    SupplierType: "ThuMua",
                //    dicFC: dicFC,
                //    dicProduct: dicProduct,
                //    dicSupplier: dicSupplier,
                //    Product: Product,
                //    Supplier: Supplier);

                //xlApp.Calculation = Excel.XlCalculation.xlCalculationAutomatic;

                //xlWb.Close(SaveChanges: false);
                //#endregion

                //#region Forecast ThuMua - KPI
                //if (!YesNoPlanning)
                //{
                //    filePath = string.Format("D:\\Documents\\Stuff\\VinEco\\Mastah Project\\{0}", "ThuMua KPI.xlsb");
                //    conStr = string.Format(Constants.Excel07ConString, filePath, header);

                //    xlWb = xlApp.Workbooks.Open(filePath,
                //        UpdateLinks: false,
                //        ReadOnly: true,
                //        Format: 5,
                //        Password: "",
                //        WriteResPassword: "",
                //        IgnoreReadOnlyRecommended: true,
                //        Origin: Excel.XlPlatform.xlWindows,
                //        Delimiter: "",
                //        Editable: false,
                //        Notify: false,
                //        Converter: 0,
                //        AddToMru: true,
                //        Local: false,
                //        CorruptLoad: false);

                //    xlApp.Calculation = Excel.XlCalculation.xlCalculationManual;

                //    xlWs = xlWb.Worksheets[1];
                //    xlRng = xlWs.UsedRange;

                //    EatForecast(FC, xlRng, xlWs, conStr, "ThuMua", dicFC, dicProduct, dicSupplier, Product, Supplier, true);

                //    xlWb.Close(SaveChanges: false);
                //}

                //#endregion

                #endregion

                #endregion

                #region Afterward Services.

                // Compact Forecasts before importing into Database.
                // All afterward services will be here.
                // Current jobs:
                //   - Deal with Confirmed PO from Purchasing Department.

                // Date layer.
                foreach (var _ForecastDate in FC.OrderByDescending(x => x.DateForecast.Date).Reverse())
                {
                    // Product layer.
                    foreach (var _ProductForecast in _ForecastDate.ListProductForecast.Reverse<ProductForecast>())
                    {
                        // Supplier layer.
                        foreach (var _SupplierForecast in _ProductForecast.ListSupplierForecast
                            .Reverse<SupplierForecast>())
                        {
                            if (_SupplierForecast.FullOrder)
                                _SupplierForecast.QuantityForecast = Math.Max(_SupplierForecast.QuantityForecast, 7);

                            /// <! For debugging Purposes !>
                            //if (_ForecastDate.DateForecast.Day == 16 && Product.Where(x => x.ProductId == _ProductForecast.ProductId).FirstOrDefault().ProductCode == "A04201" && Supplier.Where(x => x.SupplierId == _SupplierForecast.SupplierId).FirstOrDefault().SupplierCode == "AG03030000")
                            //{
                            //    var AmIHandsome = true;
                            //}

                            // Excluding FullOrder cases - Special cases.
                            // Also excluding VinEco cases - Even more special.
                            var _Supplier =
                                dicSupplier.Values.FirstOrDefault(x => x.SupplierId == _SupplierForecast.SupplierId);

                            if (_Supplier == null)
                                continue;

                            if (_Supplier.SupplierType != "VinEco" && !_SupplierForecast.FullOrder)
                            {
                                // If Purchasing Department already ordered:
                                //  -   Obey it.
                                //  -   Delete Minimum.
                                // Reason behind: 
                                //  -   Purchasing Department has interacted and dealt with Suppliers - their numbers have higher priority over normal Forecasts.
                                if (!YesNoPlanning)
                                {
                                    _SupplierForecast.QuantityForecastOriginal = _SupplierForecast.QuantityForecast;
                                    _SupplierForecast.QuantityForecast =
                                        _SupplierForecast.QuantityForecastPlanned ?? _SupplierForecast.QuantityForecast;
                                    _SupplierForecast.QuantityForecastContracted =
                                        _SupplierForecast.QuantityForecastPlanned == null
                                            ? Math.Min(_SupplierForecast.QuantityForecastContracted,
                                                _SupplierForecast.QuantityForecast)
                                            : 0;
                                }
                                // Old logic.
                                // In case of Planning, by default FC is Planned.
                                else if (_Supplier.SupplierType == "ThuMua")
                                {
                                    //_SupplierForecast.QuantityForecastPlanned = _SupplierForecast.QuantityForecast;
                                }
                            }
                            else if (_Supplier.SupplierType == "VinEco")
                            {
                                if (_SupplierForecast.QuantityForecastPlanned != null)
                                    _SupplierForecast.QuantityForecastPlanned = Math.Min(
                                        _SupplierForecast.QuantityForecastPlanned ?? 0,
                                        _SupplierForecast.QuantityForecast);
                                if (_SupplierForecast.QuantityForecastPlanned == 0)
                                    _SupplierForecast.QuantityForecastPlanned = null;
                            }

                            // If the Supplier can supply 0 product, well, remove it from the list of Suppliers.
                            if (_SupplierForecast.QuantityForecastPlanned == 0 ||
                                _SupplierForecast.QuantityForecast == 0)
                                _ProductForecast.ListSupplierForecast.Remove(_SupplierForecast);
                        }
                        // End of Supplier layer.

                        // If the Product has no Supplier, well, remove it from the list of suppliable Products..
                        if (_ProductForecast.ListSupplierForecast.Count == 0)
                            _ForecastDate.ListProductForecast.Remove(_ProductForecast);
                    }
                    // End of Product layer.

                    // If the Harvest Date has no Product to supply, well, remove it from the list of Harvest Date.
                    if (_ForecastDate.ListProductForecast.Count == 0)
                        FC.Remove(_ForecastDate);
                }

                #endregion

                db.DropCollection("Forecast");
                await db.GetCollection<ForecastDate>("Forecast").InsertManyAsync(FC);

                db.DropCollection("Product");
                await db.GetCollection<Product>("Product").InsertManyAsync(Product);

                db.DropCollection("Supplier");
                await db.GetCollection<Supplier>("Supplier").InsertManyAsync(Supplier);

                db = null;

                FC = null;

                dicProduct = null;
                dicSupplier = null;
                dicFC = null;

                //GC.Collect();
                //GC.WaitForPendingFinalizers();

                #region Part of Read from Chosen file

                //var cbr = new OleDbCommandBuilder(oleAdapt);
                //var adapter = new OleDbDataAdapter();

                //var con = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\\Users\\Shirayuki\\Desktop\\PO.accdb;");

                //adapter.InsertCommand = new OleDbCommand("Insert Into [MB] VALUES", con);

                //try
                //{
                //    adapter.Update(database);
                //}
                //catch (OleDbException ex)
                //{
                //    MessageBox.Show(ex.Message, "OleDbException Error");
                //}
                //catch (Exception x)
                //{
                //    MessageBox.Show(x.Message, "Exception Error");
                //}

                //dataGridView1.DataSource = database;
                //dataGridView1.Dock = DockStyle.Fill;

                //dataGridView1.Refresh();

                //MessageBox.Show(Environment.CurrentDirectory);

                #endregion

                WriteToRichTextBoxOutput(MethodBase.GetCurrentMethod().Name + " - Done");
                //    }
                //}
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                #region Clean up

                // Cleanup
                //GC.Collect();
                //GC.WaitForPendingFinalizers();

                // Rule of thumb for releasing com objects:
                //   never use two dots, all COM objects must be referenced and released individually
                //   ex: [somthing].[something].[something] is bad

                // Release com objects to fully kill excel process from running in the background
                //if (xlRng != null) { Marshal.ReleaseComObject(xlRng); }
                //if (xlWs != null) { Marshal.ReleaseComObject(xlWs); }

                //// Close and release
                //if (xlWb != null) { Marshal.ReleaseComObject(xlWs); }

                //xlApp.ScreenUpdating = true;
                //xlApp.EnableEvents = true;
                //xlApp.DisplayAlerts = false;
                //xlApp.DisplayStatusBar = true;
                //xlApp.AskToUpdateLinks = true;

                //// Quit and release
                //xlApp.Quit();
                //if (xlApp != null) { Marshal.ReleaseComObject(xlApp); }

                //xlRng = null;
                //xlWs = null;
                //xlWb = null;
                //xlApp = null;

                //GC.Collect();
                //GC.WaitForPendingFinalizers();

                //if (processID != 0)
                //{
                //    Process process = Process.GetProcessById(processID);
                //    process.Kill();
                //}

                #endregion
            }
        }

        /// <summary>
        ///     Updating OpenConfig file & Do afterward updating.
        /// </summary>
        private async Task UpdateOpenConfig()
        {
            try
            {
                //Console.OutputEncoding = System.Text.Encoding.UTF8;

                WriteToRichTextBoxOutput("Start!");

                #region Initialization.

                var mongoClient = new MongoClient();
                var db = mongoClient.GetDatabase("localtest");

                var Product = db.GetCollection<Product>("Product").AsQueryable().ToList();
                var ProductUnitList = new List<ProductUnit>();

                var filePath =
                    string.Format("D:\\Documents\\Stuff\\VinEco\\Mastah Project\\{0}",
                        "ChiaHang OpenConfig.xlsb");
                var conStr = string.Format(Constants.Excel07ConString, filePath, "YES");

                //var xlWb = xlApp.Workbooks.Open(filePath,
                //    UpdateLinks: false,
                //    ReadOnly: true,
                //    Format: 5,
                //    Password: "",
                //    WriteResPassword: "",
                //    IgnoreReadOnlyRecommended: true,
                //    Origin: Excel.XlPlatform.xlWindows,
                //    Delimiter: "",
                //    Editable: false,
                //    Notify: false,
                //    Converter: 0,
                //    AddToMru: true,
                //    Local: false,
                //    CorruptLoad: false);

                //xlApp.Calculation = Excel.XlCalculation.xlCalculationManual;

                var loadOpts = new LoadOptions
                {
                    MemorySetting = MemorySetting.MemoryPreference
                };
                var xlWb = new Workbook(filePath, loadOpts);
                loadOpts = null;

                var opts = new ExportTableOptions
                {
                    CheckMixedValueType = true,
                    ExportAsString = false,
                    FormatStrategy = CellValueFormatStrategy.None,
                    ExportColumnName = true
                };

                var dicUnit = new Dictionary<string, string>(StringComparer.Ordinal);

                #endregion

                #region UnitConversion

                var xlWs = xlWb.Worksheets["UnitConversion"];

                var dt = new DataTable();
                dt = xlWs.Cells.ExportDataTable(0, 0, xlWs.Cells.MaxDataRow + 1, xlWs.Cells.MaxDataColumn + 1, opts);

                //OleDbConnection oleCon = new OleDbConnection(conStr);

                //string connectionString = "Select * From [" + xlWs.Name.ToString() + "$" + xlRng.Offset[0, 0].Address[RowAbsolute: false, ColumnAbsolute: false, ReferenceStyle: Excel.XlReferenceStyle.xlA1, External: xlRng] + "]";
                //OleDbDataAdapter _oleAdapt = new OleDbDataAdapter(connectionString, oleCon);
                //_oleAdapt.Fill(dt);

                //oleCon.Close();

                foreach (DataRow dr in dt.Rows)
                {
                    var _Product = Product.FirstOrDefault(x => x.ProductCode == dr["VECode"].ToString());
                    if (_Product == null)
                    {
                        // To be fucking honest, this should NEVER be hit.
                        // Unit Converstion definition for a product that's NOT EVEN EXIST.
                        // ... and of fucking course IT IS HIT.
                    }
                    else
                    {
                        var _ProductUnit = ProductUnitList
                            .FirstOrDefault(x =>
                                x.ProductCode == dr["VECode"].ToString());

                        var _Region = dr["Region"].ToString();
                        switch (_Region)
                        {
                            case "MB":
                                _Region = "Miền Bắc";
                                break;
                            case "MN":
                                _Region = "Miền Nam";
                                break;
                            case "All":
                                _Region = "All";
                                break;
                            default: break;
                        }

                        if (_ProductUnit != null)
                        {
                            if (_ProductUnit.ListRegion == null)
                            {
                                var _ProductUnitRegion = new ProductUnitRegion
                                {
                                    _id = Guid.NewGuid(),
                                    Region = _Region,
                                    OrderUnitType = ProperUnit(dr["OrderUnitType"].ToString(), dicUnit),
                                    OrderUnitPer =
                                        ProperUnit(dr["OrderUnitType"].ToString(), dicUnit) == "Kg"
                                            ? 1
                                            : (double)dr["OderUnitPer"],
                                    SaleUnitType = ProperUnit(dr["SaleUnitType"].ToString(), dicUnit),
                                    SaleUnitPer =
                                        ProperUnit(dr["SaleUnitType"].ToString(), dicUnit) == "Kg"
                                            ? 1
                                            : (double)dr["SaleUnitPer"]
                                };

                                var _ListRegion = new List<ProductUnitRegion>();
                                _ListRegion.Add(_ProductUnitRegion);

                                _ProductUnit.ListRegion = _ListRegion;
                            }
                            else
                            {
                                var _ProductUnitRegion =
                                    _ProductUnit.ListRegion.FirstOrDefault(x => x.Region == _Region);

                                if (_ProductUnitRegion == null)
                                {
                                    _ProductUnitRegion = new ProductUnitRegion
                                    {
                                        _id = Guid.NewGuid(),
                                        Region = _Region,
                                        OrderUnitType = ProperUnit(dr["OrderUnitType"].ToString(), dicUnit),
                                        OrderUnitPer =
                                            ProperUnit(dr["OrderUnitType"].ToString(), dicUnit) == "Kg"
                                                ? 1
                                                : (double)dr["OrderUnitPer"],
                                        SaleUnitType = ProperUnit(dr["SaleUnitType"].ToString(), dicUnit),
                                        SaleUnitPer =
                                            ProperUnit(dr["SaleUnitType"].ToString(), dicUnit) == "Kg"
                                                ? 1
                                                : (double)dr["SaleUnitPer"]
                                    };
                                    _ProductUnit.ListRegion.Add(_ProductUnitRegion);
                                }
                                else
                                {
                                    _ProductUnitRegion = new ProductUnitRegion
                                    {
                                        _id = Guid.NewGuid(),
                                        Region = _Region,
                                        OrderUnitType = ProperUnit(dr["OrderUnitType"].ToString(), dicUnit),
                                        OrderUnitPer =
                                            ProperUnit(dr["OrderUnitType"].ToString(), dicUnit) == "Kg"
                                                ? 1
                                                : (double)dr["OrderUnitPer"],
                                        SaleUnitType = ProperUnit(dr["SaleUnitType"].ToString(), dicUnit),
                                        SaleUnitPer =
                                            ProperUnit(dr["SaleUnitType"].ToString(), dicUnit) == "Kg"
                                                ? 1
                                                : (double)dr["SaleUnitPer"]
                                    };
                                }
                            }
                        }
                        else
                        {
                            _ProductUnit = new ProductUnit
                            {
                                ProductCode = dr["VECode"].ToString(),
                                ProductId = Product.FirstOrDefault(x => x.ProductCode == dr["VECode"].ToString())
                                    .ProductId,
                                ListRegion = new List<ProductUnitRegion>()
                            };

                            _ProductUnit.ListRegion.Add(new ProductUnitRegion
                            {
                                _id = Guid.NewGuid(),
                                Region = _Region,
                                OrderUnitType = ProperUnit(dr["OrderUnitType"].ToString(), dicUnit),
                                OrderUnitPer =
                                    ProperUnit(dr["OrderUnitType"].ToString(), dicUnit) == "Kg"
                                        ? 1
                                        : (double)dr["OrderUnitPer"],
                                SaleUnitType = ProperUnit(dr["SaleUnitType"].ToString(), dicUnit),
                                SaleUnitPer = ProperUnit(dr["SaleUnitType"].ToString(), dicUnit) == "Kg"
                                    ? 1
                                    : (double)dr["SaleUnitPer"]
                            });

                            ProductUnitList.Add(_ProductUnit);
                        }
                    }
                }

                db.DropCollection("ProductUnit");
                await db.GetCollection<ProductUnit>("ProductUnit").InsertManyAsync(ProductUnitList);

                WriteToRichTextBoxOutput(string.Format("{0} done!", dt.TableName));

                #endregion

                #region CrossRegion

                var ListProductRegion = new List<ProductCrossRegion>();

                xlWs = xlWb.Worksheets["CrossRegion"];

                dt = new DataTable { TableName = "CrossRegion" };

                dt = xlWs.Cells.ExportDataTable(0, 0, xlWs.Cells.MaxDataRow + 1, xlWs.Cells.MaxDataColumn + 1, opts);

                foreach (DataRow dr in dt.Rows)
                {
                    var _Product = Product.Where(x => x.ProductCode == dr["Code"].ToString()).FirstOrDefault();
                    if (_Product != null)
                    {
                        var _ProductCrossRegion = new ProductCrossRegion
                        {
                            _id = _Product._id,
                            ProductId = _Product.ProductId,
                            ToNorth = dr["ToNorth"].ToString() == "Yes" ? true : false,
                            ToSouth = dr["ToSouth"].ToString() == "Yes" ? true : false
                        };
                        ListProductRegion.Add(_ProductCrossRegion);
                    }
                }

                db.DropCollection("ProductCrossRegion");
                await db.GetCollection<ProductCrossRegion>("ProductCrossRegion").InsertManyAsync(ListProductRegion);

                WriteToRichTextBoxOutput(string.Format("{0} done!", dt.TableName));

                #endregion

                #region ProductRate

                var ListProductRate = new List<ProductRate>();

                xlWs = xlWb.Worksheets["ProductRate"];

                dt = new DataTable { TableName = "ProductRate Table" };
                dt = xlWs.Cells.ExportDataTable(0, 0, xlWs.Cells.MaxDataRow + 1, xlWs.Cells.MaxDataColumn + 1, opts);

                foreach (DataRow dr in dt.Rows)
                {
                    var _Product = Product.Where(x => x.ProductCode == dr["Code"].ToString()).FirstOrDefault();
                    if (_Product != null)
                    {
                        var _ProductRate = new ProductRate
                        {
                            _id = _Product._id,
                            ProductId = _Product.ProductId,
                            ProductCode = _Product.ProductCode,
                            ToNorth = Convert.ToDouble(dr["ToNorth"] ?? 1),
                            ToSouth = Convert.ToDouble(dr["ToSouth"] ?? 1)
                        };
                        ListProductRate.Add(_ProductRate);
                    }
                }

                db.DropCollection("ProductRate");
                await db.GetCollection<ProductRate>("ProductRate").InsertManyAsync(ListProductRate);

                WriteToRichTextBoxOutput(string.Format("{0} done!", dt.TableName));

                #endregion

                #region ProductClass

                var dicClass = new Dictionary<string, string>
                {
                    {"A", "Rau ăn lá"},
                    {"B", "Rau ăn thân hoa"},
                    {"C", "Rau ăn quả "},
                    {"D", "Rau ăn củ"},
                    {"E", "Cây ăn hạt"},
                    {"F", "Rau gia vị "},
                    {"G", "Thủy canh"},
                    {"H", "Rau mầm "},
                    {"I", "Nấm"},
                    {"J", "Lá "},
                    {"K", "Trái cây (Quả)"},
                    {"L", "Gạo"},
                    {"M", "Cỏ và cây công trình"},
                    {"N", "Hoa"},
                    {"O", "Dược liệu"}
                };

                foreach (var _Product in Product)
                {
                    // Freaking amazingly gloriously typo.
                    _Product.ProductName = ProperStr(_Product.ProductName);

                    if (dicClass.TryGetValue(_Product.ProductCode.Substring(0, 1), out var _ProductClassification))
                    {
                        if (_Product.ProductCode == "K01901" || _Product.ProductCode == "K02201")
                            _ProductClassification = dicClass["F"];

                        _Product.ProductClassification = _ProductClassification;
                    }
                }

                db.DropCollection("Product");
                await db.GetCollection<Product>("Product").InsertManyAsync(Product);

                WriteToRichTextBoxOutput("Update ProductClassification - Done!");

                #endregion

                #region ExtraProductionInformation

                xlWs = xlWb.Worksheets["ExtraProductInformation"];

                dt = new DataTable { TableName = "ExtraProductInformation Table" };
                dt = xlWs.Cells.ExportDataTable(0, 0, xlWs.Cells.MaxDataRow + 1, xlWs.Cells.MaxDataColumn + 1, opts);

                foreach (DataRow dr in dt.Rows)
                {
                    var _Product = Product.Where(x => x.ProductCode == dr["Code"].ToString()).FirstOrDefault();

                    if (_Product != null)
                    {
                        _Product.ProductOrientation = dr["ProductionOrientation"].ToString();
                        _Product.ProductClimate = dr["ProductClimate"].ToString();
                        _Product.ProductionGroup = dr["ProductionGroup"].ToString();
                    }
                }

                db.DropCollection("Product");
                await db.GetCollection<Product>("Product").InsertManyAsync(Product);

                WriteToRichTextBoxOutput(string.Format("Update {0} - Done!", xlWs.Name));

                #endregion

                #region Remove Products outside of Master List

                // Dealing with all Products outside of MasterList.
                // Like, really, why the heck would you order something we don't even produce.

                xlWs = xlWb.Worksheets["MasterList"];

                dt = new DataTable { TableName = "MasterList Table" };
                dt = xlWs.Cells.ExportDataTable(0, 0, xlWs.Cells.MaxDataRow + 1, xlWs.Cells.MaxDataColumn + 1, opts);

                var ProductMasterList = new Dictionary<Guid, string>();
                foreach (DataRow dr in dt.Rows)
                {
                    var _Product = Product.Where(x => x.ProductCode == dr["Code"].ToString()).FirstOrDefault();
                    if (_Product != null)
                    {
                        if ((string)dr["North"] == "Yes") _Product.ProductNote.Add("North");
                        if ((string)dr["South"] == "Yes") _Product.ProductNote.Add("South");
                    }
                }

                db.DropCollection("Product");
                await db.GetCollection<Product>("Product").InsertManyAsync(Product);

                WriteToRichTextBoxOutput(string.Format("Update {0} - Done!", xlWs.Name));

                #endregion

                #region Priority

                xlWs = xlWb.Worksheets["Priority"];

                dt = new DataTable { TableName = "Priority Table" };
                dt = xlWs.Cells.ExportDataTable(0, 0, xlWs.Cells.MaxDataRow + 1, xlWs.Cells.MaxDataColumn + 1, opts);

                var ListCustomer = db.GetCollection<Customer>("Customer").AsQueryable().ToList();

                var dicPriority = new Dictionary<string, bool>();

                foreach (DataRow dr in dt.Rows)
                    dicPriority.Add(dr["CCODE"].ToString(), true);

                foreach (var _Customer in ListCustomer.Where(Customer =>
                    Customer.CustomerCode == "VM" || Customer.CustomerCode == "VM+" || Customer.CustomerCode == "VM+ VinEco"))
                {
                    // Cleaning stuff
                    _Customer.CustomerType = _Customer.CustomerType.Replace("Priority", "").Trim();

                    // Ooooh, I heard that you didn't get enough vegetables.
                    if (dicPriority.ContainsKey(_Customer.CustomerCode))
                        _Customer.CustomerType += " Priority";

                    _Customer.CustomerRegion = ProperStr(_Customer.CustomerRegion);
                }

                db.DropCollection("Customer");
                await db.GetCollection<Customer>("Customer").InsertManyAsync(ListCustomer);

                WriteToRichTextBoxOutput(string.Format("Update {0} - Done!", "Customer"));

                #endregion

                #region Clean up.

                dt = null;

                xlWs = null;
                xlWb = null;

                //Marshal.ReleaseComObject(xlRng); xlRng = null;
                //Marshal.ReleaseComObject(xlWs); xlWs = null;

                //xlWb.Close(SaveChanges: false);
                //Marshal.ReleaseComObject(xlWb); xlWb = null;

                //xlApp.Quit();
                //Marshal.ReleaseComObject(xlApp); xlApp = null;

                //// Kill the instance of Interop Excel.Application used by this call.
                //if (processID != 0)
                //{
                //    Process process = Process.GetProcessById(processID);
                //    process.Kill();
                //}

                #endregion

                WriteToRichTextBoxOutput("Update OpenConfig Done!");
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        ///     Do naughty stuff with FC
        /// </summary>
        private void EatForecast(List<ForecastDate> FC, Range xlRng, Worksheet xlWs, string conStr,
            string SupplierType, Dictionary<DateTime, Dictionary<string, Dictionary<string, Guid>>> dicFC,
            Dictionary<string, Product> dicProduct, Dictionary<string, Supplier> dicSupplier, List<Product> Product,
            List<Supplier> Supplier, bool YesNoKPI = false)
        {
            try
            {
                var rowIndex = 0;
                if ((xlRng.Cells[1, 1].value != "Region") & (xlRng.Cells[1, 1].value != "Vùng"))
                    do
                    {
                        rowIndex++;
                        if (rowIndex >= xlRng.Rows.Count) return;
                    } while ((xlRng.Cells[rowIndex + 1, 1].Value != "Region") &
                             (xlRng.Cells[rowIndex + 1, 1].Value != "Vùng"));

                var dt = new DataTable();

                var oleCon = new OleDbConnection(conStr);

                var _oleAdapt = new OleDbDataAdapter(
                    "Select * From [" + xlWs.Name + "$" + xlRng.Offset[rowIndex, 0]
                        .Address[false, false, XlReferenceStyle.xlA1,
                            xlRng] + "]", oleCon);
                var _str = xlRng.Offset[rowIndex, 0].Address;
                WriteToRichTextBoxOutput(_str);
                _oleAdapt.Fill(dt);

                oleCon.Close();

                // To deal with the uhm, Templates having different Headers.
                // Please shoot me.
                foreach (DataColumn dc in dt.Columns)
                {
                    if (dt.Columns.Contains("Vùng")) dt.Columns["Vùng"].ColumnName = "Region";
                    if (dt.Columns.Contains("Mã Farm")) dt.Columns["Mã Farm"].ColumnName = "SCODE";
                    if (dt.Columns.Contains("Tên Farm")) dt.Columns["Tên Farm"].ColumnName = "SNAME";
                    if (dt.Columns.Contains("Nhóm")) dt.Columns["Nhóm"].ColumnName = "PCLASS";
                    if (dt.Columns.Contains("Mã VECrops")) dt.Columns["Mã VECrops"].ColumnName = "VECrops Code";
                    if (dt.Columns.Contains("Mã VinEco")) dt.Columns["Mã VinEco"].ColumnName = "PCODE";
                    if (dt.Columns.Contains("Tên VinEco")) dt.Columns["Tên VinEco"].ColumnName = "PNAME";
                }

                // Main Loop
                foreach (DataColumn dc in dt.Columns)
                {
                    DateTime dateValue;

                    // Loop for every column. Only stop at column with Date at the top ( Indicating it being PurchaseOrder for that Date )
                    if (DateTime.TryParse(dc.ColumnName, out dateValue))
                    {
                        ForecastDate _FC = null;
                        var isNewFC = false;

                        Dictionary<string, Dictionary<string, Guid>> _dicProduct = null;
                        // Find PurchaseOrder for that Date
                        if (dicFC.TryGetValue(dateValue.Date, out _dicProduct))
                        {
                            _FC = FC.Where(x => x.DateForecast.Date == dateValue.Date).FirstOrDefault();
                        }
                        // Create a blank one in case it doesn't exist
                        else
                        {
                            isNewFC = true;

                            _FC = new ForecastDate();
                            _FC._id = Guid.NewGuid();
                            _FC.ForecastDateId = _FC._id;

                            _FC.DateForecast = dateValue.Date;
                            _FC.ListProductForecast = new List<ProductForecast>();

                            dicFC.Add(dateValue.Date, new Dictionary<string, Dictionary<string, Guid>>());
                        }

                        // First layer
                        // Get the list of all Products being Ordered that day.
                        var _listProductForecast = _FC.ListProductForecast;
                        if (_listProductForecast == null) _listProductForecast = new List<ProductForecast>();

                        // Loop for every value
                        foreach (DataRow dr in dt.Rows)
                        {
                            // In case of empty SCODE. I really hate to deal with this case. Like, really.
                            if (dr["SCODE"] == null || string.IsNullOrEmpty(dr["SCODE"].ToString()))
                                dr["SCODE"] = dr["SNAME"]; // Oh for god's sake.

                            // If OrderQuantity is not 0 - Not Anymore?
                            //object _OrderQuantity = dr[dc.ColumnName];
                            if (
                                dr["PCODE"] !=
                                DBNull.Value /*&& dr[dc.ColumnName] != DBNull.Value*/ /*&& Convert.ToDouble(dr[dc.ColumnName]) > 0*/ &&
                                (SupplierType == "ThuMua" ? dr["SCODE"] != DBNull.Value : true))
                            {
                                // Olala
                                List<SupplierForecast> _ListSupplierForecast = null;
                                SupplierForecast _SupplierForecast = null;
                                ProductForecast _ProductForecast = null;
                                // Olala2
                                var isNewProductOrder = false;
                                var isNewCustomerOrder = false;
                                // Olala3
                                Dictionary<string, Guid> dicStore = null;
                                // #RandomGreenStuff
                                _dicProduct = dicFC[dateValue.Date];
                                if (_dicProduct.TryGetValue(dr["PCODE"].ToString(), out dicStore))
                                {
                                    Product _product = null;
                                    if (!dicProduct.TryGetValue(dr["PCODE"].ToString(), out _product))
                                    {
                                        _product = dicProduct.Values.Where(x => x.ProductCode == dr["PCODE"].ToString())
                                            .FirstOrDefault();
                                        if (_product == null)
                                        {
                                            _product = new Product();

                                            _product._id = Guid.NewGuid();
                                            _product.ProductId = _product._id;
                                            _product.ProductCode = dr["PCODE"].ToString();
                                            _product.ProductName = dr["PNAME"].ToString();
                                            //_product.ProductClassification = dr["PCLASS"].ToString();
                                            _product.ProductVECode = dt.Columns.Contains("VECrops Code")
                                                ? dr["VECrops Code"].ToString()
                                                : "";

                                            Product.Add(_product);

                                            dicProduct.Add(dr["PCODE"].ToString(), _product);
                                        }
                                    }
                                    _ProductForecast = _FC.ListProductForecast
                                        .Where(x => x.ProductId == _product.ProductId).FirstOrDefault();

                                    Guid _id;
                                    if (dicStore.TryGetValue(dr["SCODE"].ToString(), out _id))
                                    {
                                        _SupplierForecast = _ProductForecast.ListSupplierForecast
                                            .Where(x => x.SupplierId == _id).FirstOrDefault();
                                    }
                                    else
                                    {
                                        isNewCustomerOrder = true;

                                        Supplier _supplier = null;
                                        if (!dicSupplier.TryGetValue(dr["SCODE"].ToString(), out _supplier))
                                        {
                                            _supplier = dicSupplier.Values
                                                .Where(x => x.SupplierCode == dr["SCODE"].ToString()).FirstOrDefault();
                                            if (_supplier == null)
                                            {
                                                _supplier = new Supplier();

                                                _supplier._id = Guid.NewGuid();
                                                _supplier.SupplierId = _supplier._id;
                                                _supplier.SupplierCode =
                                                    dr["SCODE"]
                                                        .ToString(); //SupplierType == "ThuMua" ? (dr["SCODE"] == null ? dr["SCODE"].ToString() : dr["SNAME"].ToString()) : dr["SNAME"].ToString();
                                                _supplier.SupplierName = dr["SNAME"].ToString();
                                                _supplier.SupplierType =
                                                    SupplierType == "ThuMua" && dr["Tag"].ToString() == "VCM"
                                                        ? "VCM"
                                                        : SupplierType;

                                                var _region = dr["Region"].ToString();
                                                switch (_region)
                                                {
                                                    case "LD":
                                                        _region = "Lâm Đồng";
                                                        break;
                                                    case "MB":
                                                        _region = "Miền Bắc";
                                                        break;
                                                    case "MN":
                                                        _region = "Miền Nam";
                                                        break;
                                                    default: break;
                                                }
                                                _supplier.SupplierRegion = _region;
                                                _supplier.SupplierType =
                                                    SupplierType == "ThuMua" && dr["Tag"].ToString() == "VCM"
                                                        ? "VCM"
                                                        : SupplierType;

                                                Supplier.Add(_supplier);
                                                dicSupplier.Add(_supplier.SupplierCode, _supplier);
                                            }
                                        }

                                        _SupplierForecast = new SupplierForecast();
                                        _SupplierForecast._id = Guid.NewGuid();
                                        _SupplierForecast.SupplierForecastId = _SupplierForecast._id;
                                        _SupplierForecast.SupplierId = _supplier.SupplierId;

                                        dicFC[dateValue.Date][dr["PCODE"].ToString()]
                                            .Add(dr["SCODE"].ToString(), _supplier.SupplierId);
                                    }
                                }
                                else
                                {
                                    isNewProductOrder = true;
                                    isNewCustomerOrder = true;

                                    Product _product = null;
                                    if (!dicProduct.TryGetValue(dr["PCODE"].ToString(), out _product))
                                    {
                                        _product = dicProduct.Values.Where(x => x.ProductCode == dr["PCODE"].ToString())
                                            .FirstOrDefault();
                                        if (_product == null)
                                        {
                                            _product = new Product();

                                            _product._id = Guid.NewGuid();
                                            _product.ProductId = _product._id;
                                            _product.ProductCode = dr["PCODE"].ToString();
                                            _product.ProductName = dr["PNAME"].ToString();
                                            //_product.ProductClassification = dr["PCLASS"].ToString();
                                            _product.ProductVECode = dt.Columns.Contains("VECrops Code")
                                                ? dr["VECrops Code"].ToString()
                                                : "";

                                            Product.Add(_product);

                                            dicProduct.Add(dr["PCODE"].ToString(), _product);
                                        }
                                    }

                                    _ProductForecast = new ProductForecast();
                                    _ProductForecast._id = Guid.NewGuid();
                                    _ProductForecast.ProductForecastId = _ProductForecast._id;
                                    _ProductForecast.ProductId = _product.ProductId;

                                    _ProductForecast.ListSupplierForecast = new List<SupplierForecast>();

                                    Supplier _supplier = null;
                                    if (!dicSupplier.TryGetValue(dr["SCODE"].ToString(), out _supplier))
                                    {
                                        _supplier = dicSupplier.Values
                                            .Where(x => x.SupplierCode == dr["SCODE"].ToString()).FirstOrDefault();
                                        if (_supplier == null)
                                        {
                                            _supplier = new Supplier();

                                            _supplier._id = Guid.NewGuid();
                                            _supplier.SupplierId = _supplier._id;
                                            _supplier.SupplierCode =
                                                dr["SCODE"]
                                                    .ToString(); //SupplierType == "ThuMua" ? (dr["SCODE"] == null ? dr["SCODE"].ToString() : dr["SNAME"].ToString()) : dr["SNAME"].ToString();
                                            _supplier.SupplierName = dr["SNAME"].ToString();
                                            _supplier.SupplierType =
                                                SupplierType == "ThuMua" && dr["Tag"].ToString() == "VCM"
                                                    ? "VCM"
                                                    : SupplierType;

                                            var _region = dr["Region"].ToString();
                                            switch (_region)
                                            {
                                                case "LD":
                                                    _region = "Lâm Đồng";
                                                    break;
                                                case "MB":
                                                    _region = "Miền Bắc";
                                                    break;
                                                case "MN":
                                                    _region = "Miền Nam";
                                                    break;
                                                default: break;
                                            }
                                            _supplier.SupplierRegion = _region;
                                            _supplier.SupplierType =
                                                SupplierType == "ThuMua" && dr["Tag"].ToString() == "VCM"
                                                    ? "VCM"
                                                    : SupplierType;

                                            Supplier.Add(_supplier);
                                            dicSupplier.Add(_supplier.SupplierCode, _supplier);
                                        }
                                    }

                                    _SupplierForecast = new SupplierForecast();
                                    _SupplierForecast._id = Guid.NewGuid();
                                    _SupplierForecast.SupplierForecastId = _SupplierForecast._id;
                                    _SupplierForecast.SupplierId = _supplier.SupplierId;

                                    dicFC[dateValue.Date].Add(dr["PCODE"].ToString(), new Dictionary<string, Guid>());
                                    dicFC[dateValue.Date][dr["PCODE"].ToString()]
                                        .Add(dr["SCODE"].ToString(), _supplier.SupplierId);
                                }

                                // Filling in data
                                _ListSupplierForecast = _ProductForecast.ListSupplierForecast;

                                // Special part for ThuMua
                                var myTI = new CultureInfo("en-US", false).TextInfo;
                                if (SupplierType != "VinEco" && !YesNoKPI)
                                {
                                    _SupplierForecast.QualityControlPass = string.IsNullOrEmpty(dr["QC"].ToString())
                                        ? false
                                        : (myTI.ToTitleCase(dr["QC"].ToString()) == "Ok" ? true : false);
                                    _SupplierForecast.LabelVinEco = string.IsNullOrEmpty(dr["Label VE"].ToString())
                                        ? false
                                        : (myTI.ToTitleCase(dr["Label VE"].ToString()) == "Yes" ? true : false);
                                    _SupplierForecast.FullOrder = string.IsNullOrEmpty(dr["100%"].ToString())
                                        ? false
                                        : (myTI.ToTitleCase(dr["100%"].ToString()) == "Yes" ? true : false);
                                    _SupplierForecast.CrossRegion = string.IsNullOrEmpty(dr["CrossRegion"].ToString())
                                        ? false
                                        : (myTI.ToTitleCase(dr["CrossRegion"].ToString()) == "Yes" ? true : false);
                                    _SupplierForecast.Level = string.IsNullOrEmpty(dr["Level"].ToString())
                                        ? Convert.ToByte(0)
                                        : Convert.ToByte(dr["Level"]);
                                    _SupplierForecast.Availability = string.IsNullOrEmpty(dr["Availability"].ToString())
                                        ? ""
                                        : dr["Availability"].ToString();
                                }
                                else if (!YesNoKPI)
                                {
                                    _SupplierForecast.QualityControlPass = true;
                                    _SupplierForecast.LabelVinEco = true;
                                    _SupplierForecast.FullOrder = false;
                                    _SupplierForecast.CrossRegion = false;
                                    _SupplierForecast.Level = 1;
                                    _SupplierForecast.Availability = "1234567";

                                    // To deal with some Supplier only Supply for a targetted Customer Group.
                                    _SupplierForecast.Target =
                                        dt.Columns.Contains("Target") ? dr["Target"].ToString() : "All";
                                }
                                else if (YesNoKPI && dr["Source"].ToString() == "ThuMua")
                                {
                                    _SupplierForecast.QualityControlPass = string.IsNullOrEmpty(dr["QC"].ToString())
                                        ? _SupplierForecast.QualityControlPass
                                        : (myTI.ToTitleCase(dr["QC"].ToString()) == "Ok" ? true : false);
                                    _SupplierForecast.LabelVinEco = string.IsNullOrEmpty(dr["Label VE"].ToString())
                                        ? _SupplierForecast.LabelVinEco
                                        : (myTI.ToTitleCase(dr["Label VE"].ToString()) == "Yes" ? true : false);
                                    _SupplierForecast.FullOrder = string.IsNullOrEmpty(dr["100%"].ToString())
                                        ? _SupplierForecast.FullOrder
                                        : (myTI.ToTitleCase(dr["100%"].ToString()) == "Yes" ? true : false);
                                    _SupplierForecast.CrossRegion = string.IsNullOrEmpty(dr["CrossRegion"].ToString())
                                        ? _SupplierForecast.CrossRegion
                                        : (myTI.ToTitleCase(dr["CrossRegion"].ToString()) == "Yes" ? true : false);
                                    _SupplierForecast.Level = string.IsNullOrEmpty(dr["Level"].ToString())
                                        ? _SupplierForecast.Level
                                        : Convert.ToByte(dr["Level"]);
                                    _SupplierForecast.Availability = string.IsNullOrEmpty(dr["Availability"].ToString())
                                        ? _SupplierForecast.Availability
                                        : dr["Availability"].ToString();
                                }

                                if (SupplierType == "VinEco" && dr["PCODE"].ToString().Substring(0, 1) == "K" &&
                                    (dr["Region"].ToString() == "MN" || dr["Region"].ToString() == "Miền Nam")
                                ) //dicCrossRegionVinEco.ContainsKey(dr["PCODE"].ToString()))
                                {
                                    _SupplierForecast.CrossRegion = true;
                                    if (dr["PCODE"].ToString() == "K03501") _SupplierForecast.CrossRegion = false;
                                }

                                ///// < !For debugging purposes !>
                                //if (!YesNoKPI && dateValue.Day == 16 && (string)dr["PCODE"] == "C02801" && (string)dr["SCODE"] == "AG03030000")
                                //{
                                //    byte AmIHandsome = 0;
                                //}

                                // 3rd FC layer - Normal Forecast.
                                double _QuantityForecast = 0;
                                if (double.TryParse(
                                    (dr[dc.ColumnName] == DBNull.Value ? 0 : dr[dc.ColumnName]).ToString(),
                                    out _QuantityForecast))
                                {
                                    if (!YesNoKPI)
                                        _SupplierForecast.QuantityForecast += _QuantityForecast;

                                    // 2nd FC layer - Minimum / Contracted Forecast - 2nd Highest Priority. 
                                    if (dt.Columns.Contains("Min"))
                                    {
                                        double _QuantityForecastContracted = 0;
                                        if (double.TryParse((dr["Min"] == DBNull.Value ? 0 : dr["Min"]).ToString(),
                                            out _QuantityForecastContracted))
                                            _SupplierForecast.QuantityForecastContracted += _QuantityForecastContracted;
                                    }
                                }

                                if (YesNoKPI &&
                                    Convert.ToDateTime(dr["EffectiveFrom"]).Date <=
                                    DateTime.Parse(dc.ColumnName).Date && Convert.ToDateTime(dr["EffectiveTo"]).Date >=
                                    DateTime.Parse(dc.ColumnName).Date)
                                {
                                    _SupplierForecast.QualityControlPass = true;
                                    double _QuantityForecastPlanned = 0;
                                    if (double.TryParse(
                                        (dr[dc.ColumnName] == DBNull.Value ? 0 : dr[dc.ColumnName]).ToString(),
                                        out _QuantityForecastPlanned))
                                    {
                                        _SupplierForecast.QuantityForecastPlanned =
                                            _SupplierForecast.QuantityForecastPlanned ?? 0;
                                        _SupplierForecast.QuantityForecastPlanned += _QuantityForecastPlanned;

                                        // In case outside of Forecast, which, is an entirely new Supplier.
                                        // Yes this does happen.
                                        _SupplierForecast.QualityControlPass = true;

                                        //_SupplierForecast.QuantityForecastContracted = Math.Max(_SupplierForecast.QuantityForecastContracted - _SupplierForecast.QuantityForecastPlanned, 0);
                                        //_SupplierForecast.QuantityForecast = Math.Max(_SupplierForecast.QuantityForecast - _SupplierForecast.QuantityForecastPlanned - _SupplierForecast.QuantityForecastContracted, 0);
                                    }
                                }

                                if (isNewCustomerOrder) _ListSupplierForecast.Add(_SupplierForecast);

                                _ProductForecast.ListSupplierForecast = _ListSupplierForecast;
                                if (isNewProductOrder) _FC.ListProductForecast.Add(_ProductForecast);
                            }
                        }

                        _FC.ListProductForecast = _listProductForecast;

                        if (isNewFC) FC.Add(_FC);
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        ///     Do naughty stuff with FC
        /// </summary>
        private void EatForecastAspose(List<ForecastDate> FC, Aspose.Cells.Worksheet xlWs, string SupplierType,
            Dictionary<DateTime, Dictionary<string, Dictionary<string, Guid>>> dicFC,
            Dictionary<string, Product> dicProduct, Dictionary<string, Supplier> dicSupplier, List<Product> Product,
            List<Supplier> Supplier, bool YesNoKPI = false)
        {
            try
            {
                //int rowIndex = 0;
                //if (xlRng.Cells[1, 1].value != "Region" & xlRng.Cells[1, 1].value != "Vùng")
                //{
                //    do
                //    {
                //        rowIndex++;
                //        if (rowIndex >= xlRng.Rows.Count) { return; }
                //    } while (xlRng.Cells[rowIndex + 1, 1].Value != "Region" & xlRng.Cells[rowIndex + 1, 1].Value != "Vùng");
                //}

                //DataTable dt = new DataTable();

                //OleDbConnection oleCon = new OleDbConnection(conStr);

                //OleDbDataAdapter _oleAdapt = new OleDbDataAdapter("Select * From [" + xlWs.Name.ToString() + "$" + xlRng.Offset[rowIndex, 0].Address[RowAbsolute: false, ColumnAbsolute: false, ReferenceStyle: Excel.XlReferenceStyle.xlA1, External: xlRng] + "]", oleCon);
                //string _str = xlRng.Offset[rowIndex, 0].Address as string;
                //WriteToRichTextBoxOutput(_str);
                //_oleAdapt.Fill(dt);

                //oleCon.Close();

                // Find first row.
                var rowIndex = 0;
                do
                {
                    if (xlWs.Cells[rowIndex, 0].Value == null || xlWs.Cells[rowIndex, 0].Value.ToString() != "Vùng" &&
                        xlWs.Cells[rowIndex, 0].Value.ToString() != "Region") rowIndex++;
                } while (rowIndex <= xlWs.Cells.MaxDataRow + 1 && xlWs.Cells[rowIndex, 0].Value == null ||
                         xlWs.Cells[rowIndex, 0].Value.ToString() != "Vùng" &&
                         xlWs.Cells[rowIndex, 0].Value.ToString() != "Region");

                if (rowIndex > xlWs.Cells.MaxDataRow + 1)
                {
                    for (var i = 0; i < 7; i++)
                        WriteToRichTextBoxOutput("Wrong Format.");
                    return;
                }

                // ... ah well, option based 0.
                //rowIndex--;

                // Import into a DataTable.
                var opts = new ExportTableOptions
                {
                    CheckMixedValueType = true,
                    ExportAsString = false,
                    FormatStrategy = CellValueFormatStrategy.None,
                    ExportColumnName = true
                };

                var dt = new DataTable { TableName = xlWs.Name };
                dt = xlWs.Cells.ExportDataTable(rowIndex, 0, xlWs.Cells.MaxDataRow + 1, xlWs.Cells.MaxDataColumn + 1,
                    opts);

                opts = null;

                // To deal with the uhm, Templates having different Headers.
                // Please shoot me.
                if (dt.Columns.Contains("Vùng")) dt.Columns["Vùng"].ColumnName = "Region";
                if (dt.Columns.Contains("Mã Farm")) dt.Columns["Mã Farm"].ColumnName = "SCODE";
                if (dt.Columns.Contains("Tên Farm")) dt.Columns["Tên Farm"].ColumnName = "SNAME";
                if (dt.Columns.Contains("Nhóm")) dt.Columns["Nhóm"].ColumnName = "PCLASS";
                if (dt.Columns.Contains("Mã VECrops")) dt.Columns["Mã VECrops"].ColumnName = "VECrops Code";
                if (dt.Columns.Contains("Mã VinEco")) dt.Columns["Mã VinEco"].ColumnName = "PCODE";
                if (dt.Columns.Contains("Tên VinEco")) dt.Columns["Tên VinEco"].ColumnName = "PNAME";


                // Main Loop
                foreach (DataColumn dc in dt.Columns)
                {

                    // Loop for every column. Only stop at column with Date at the top ( Indicating it being PurchaseOrder for that Date )
                    //if (DateTime.TryParse(dc.ColumnName, out dateValue))
                    if (StringToDate(dc.ColumnName) != null)
                    {
                        DateTime dateValue = StringToDate(dc.ColumnName) ?? DateTime.MinValue;

                        ForecastDate _FC = null;
                        var isNewFC = false;

                        Dictionary<string, Dictionary<string, Guid>> _dicProduct = null;
                        // Find PurchaseOrder for that Date
                        if (dicFC.TryGetValue(dateValue.Date, out _dicProduct))
                        {
                            _FC = FC.Where(x => x.DateForecast.Date == dateValue.Date).FirstOrDefault();
                        }
                        // Create a blank one in case it doesn't exist
                        else
                        {
                            isNewFC = true;

                            _FC = new ForecastDate();
                            _FC._id = Guid.NewGuid();
                            _FC.ForecastDateId = _FC._id;

                            _FC.DateForecast = dateValue.Date;
                            _FC.ListProductForecast = new List<ProductForecast>();

                            dicFC.Add(dateValue.Date, new Dictionary<string, Dictionary<string, Guid>>());
                        }

                        // First layer
                        // Get the list of all Products being Ordered that day.
                        var _listProductForecast = _FC.ListProductForecast;
                        if (_listProductForecast == null) _listProductForecast = new List<ProductForecast>();

                        // Loop for every value
                        foreach (DataRow dr in dt.Rows)
                        {
                            // In case of empty SCODE. I really hate to deal with this case. Like, really.
                            if (dr["SCODE"] == null || string.IsNullOrEmpty(dr["SCODE"].ToString()))
                                dr["SCODE"] = dr["SNAME"]; // Oh for god's sake.

                            // If OrderQuantity is not 0 - Not Anymore?
                            //object _OrderQuantity = dr[dc.ColumnName];
                            if (
                                dr["PCODE"] !=
                                DBNull.Value /*&& dr[dc.ColumnName] != DBNull.Value*/ /*&& Convert.ToDouble(dr[dc.ColumnName]) > 0*/ &&
                                (SupplierType == "ThuMua" ? dr["SCODE"] != DBNull.Value : true))
                            {
                                // Olala
                                List<SupplierForecast> _ListSupplierForecast = null;
                                SupplierForecast _SupplierForecast = null;
                                ProductForecast _ProductForecast = null;
                                // Olala2
                                var isNewProductOrder = false;
                                var isNewCustomerOrder = false;
                                // Olala3
                                Dictionary<string, Guid> dicStore = null;
                                // #RandomGreenStuff
                                _dicProduct = dicFC[dateValue.Date];
                                if (_dicProduct.TryGetValue(dr["PCODE"].ToString(), out dicStore))
                                {
                                    Product _product = null;
                                    if (!dicProduct.TryGetValue(dr["PCODE"].ToString(), out _product))
                                    {
                                        _product = dicProduct.Values.Where(x => x.ProductCode == dr["PCODE"].ToString())
                                            .FirstOrDefault();
                                        if (_product == null)
                                        {
                                            _product = new Product();

                                            _product._id = Guid.NewGuid();
                                            _product.ProductId = _product._id;
                                            _product.ProductCode = dr["PCODE"].ToString();
                                            _product.ProductName = dr["PNAME"].ToString();
                                            //_product.ProductClassification = dr["PCLASS"].ToString();
                                            _product.ProductVECode = dt.Columns.Contains("VECrops Code")
                                                ? dr["VECrops Code"].ToString()
                                                : "";

                                            Product.Add(_product);

                                            dicProduct.Add(dr["PCODE"].ToString(), _product);
                                        }
                                    }
                                    _ProductForecast = _FC.ListProductForecast
                                        .Where(x => x.ProductId == _product.ProductId).FirstOrDefault();

                                    Guid _id;
                                    if (dicStore.TryGetValue(dr["SCODE"].ToString(), out _id))
                                    {
                                        _SupplierForecast = _ProductForecast.ListSupplierForecast
                                            .Where(x => x.SupplierId == _id).FirstOrDefault();
                                    }
                                    else
                                    {
                                        isNewCustomerOrder = true;

                                        Supplier _supplier = null;
                                        if (!dicSupplier.TryGetValue(dr["SCODE"].ToString(), out _supplier))
                                        {
                                            _supplier = dicSupplier.Values
                                                .Where(x => x.SupplierCode == dr["SCODE"].ToString()).FirstOrDefault();
                                            if (_supplier == null)
                                            {
                                                _supplier = new Supplier();

                                                _supplier._id = Guid.NewGuid();
                                                _supplier.SupplierId = _supplier._id;
                                                _supplier.SupplierCode =
                                                    dr["SCODE"]
                                                        .ToString(); //SupplierType == "ThuMua" ? (dr["SCODE"] == null ? dr["SCODE"].ToString() : dr["SNAME"].ToString()) : dr["SNAME"].ToString();
                                                _supplier.SupplierName = dr["SNAME"].ToString();
                                                _supplier.SupplierType =
                                                    SupplierType == "ThuMua" && dr["Tag"].ToString() == "VCM"
                                                        ? "VCM"
                                                        : SupplierType;

                                                var _region = dr["Region"].ToString();
                                                switch (_region)
                                                {
                                                    case "LD":
                                                        _region = "Lâm Đồng";
                                                        break;
                                                    case "MB":
                                                        _region = "Miền Bắc";
                                                        break;
                                                    case "MN":
                                                        _region = "Miền Nam";
                                                        break;
                                                    default: break;
                                                }
                                                _supplier.SupplierRegion = _region;
                                                _supplier.SupplierType =
                                                    SupplierType == "ThuMua" && dr["Tag"].ToString() == "VCM"
                                                        ? "VCM"
                                                        : SupplierType;

                                                Supplier.Add(_supplier);
                                                dicSupplier.Add(_supplier.SupplierCode, _supplier);
                                            }
                                        }

                                        _SupplierForecast = new SupplierForecast();
                                        _SupplierForecast._id = Guid.NewGuid();
                                        _SupplierForecast.SupplierForecastId = _SupplierForecast._id;
                                        _SupplierForecast.SupplierId = _supplier.SupplierId;

                                        dicFC[dateValue.Date][dr["PCODE"].ToString()]
                                            .Add(dr["SCODE"].ToString(), _supplier.SupplierId);
                                    }
                                }
                                else
                                {
                                    isNewProductOrder = true;
                                    isNewCustomerOrder = true;

                                    Product _product = null;
                                    if (!dicProduct.TryGetValue(dr["PCODE"].ToString(), out _product))
                                    {
                                        _product = dicProduct.Values.Where(x => x.ProductCode == dr["PCODE"].ToString())
                                            .FirstOrDefault();
                                        if (_product == null)
                                        {
                                            _product = new Product();

                                            _product._id = Guid.NewGuid();
                                            _product.ProductId = _product._id;
                                            _product.ProductCode = dr["PCODE"].ToString();
                                            _product.ProductName = dr["PNAME"].ToString();
                                            //_product.ProductClassification = dr["PCLASS"].ToString();
                                            _product.ProductVECode = dt.Columns.Contains("VECrops Code")
                                                ? dr["VECrops Code"].ToString()
                                                : "";

                                            Product.Add(_product);

                                            dicProduct.Add(dr["PCODE"].ToString(), _product);
                                        }
                                    }

                                    _ProductForecast = new ProductForecast();
                                    _ProductForecast._id = Guid.NewGuid();
                                    _ProductForecast.ProductForecastId = _ProductForecast._id;
                                    _ProductForecast.ProductId = _product.ProductId;

                                    _ProductForecast.ListSupplierForecast = new List<SupplierForecast>();

                                    Supplier _supplier = null;
                                    if (!dicSupplier.TryGetValue(dr["SCODE"].ToString(), out _supplier))
                                    {
                                        _supplier = dicSupplier.Values
                                            .Where(x => x.SupplierCode == dr["SCODE"].ToString()).FirstOrDefault();
                                        if (_supplier == null)
                                        {
                                            _supplier = new Supplier();

                                            _supplier._id = Guid.NewGuid();
                                            _supplier.SupplierId = _supplier._id;
                                            _supplier.SupplierCode =
                                                dr["SCODE"]
                                                    .ToString(); //SupplierType == "ThuMua" ? (dr["SCODE"] == null ? dr["SCODE"].ToString() : dr["SNAME"].ToString()) : dr["SNAME"].ToString();
                                            _supplier.SupplierName = dr["SNAME"].ToString();
                                            _supplier.SupplierType =
                                                SupplierType == "ThuMua" && dr["Tag"].ToString() == "VCM"
                                                    ? "VCM"
                                                    : SupplierType;

                                            var _region = dr["Region"].ToString();
                                            switch (_region)
                                            {
                                                case "LD":
                                                    _region = "Lâm Đồng";
                                                    break;
                                                case "MB":
                                                    _region = "Miền Bắc";
                                                    break;
                                                case "MN":
                                                    _region = "Miền Nam";
                                                    break;
                                                default: break;
                                            }
                                            _supplier.SupplierRegion = _region;
                                            _supplier.SupplierType =
                                                SupplierType == "ThuMua" && dr["Tag"].ToString() == "VCM"
                                                    ? "VCM"
                                                    : SupplierType;

                                            Supplier.Add(_supplier);
                                            dicSupplier.Add(_supplier.SupplierCode, _supplier);
                                        }
                                    }

                                    _SupplierForecast = new SupplierForecast();
                                    _SupplierForecast._id = Guid.NewGuid();
                                    _SupplierForecast.SupplierForecastId = _SupplierForecast._id;
                                    _SupplierForecast.SupplierId = _supplier.SupplierId;

                                    dicFC[dateValue.Date].Add(dr["PCODE"].ToString(), new Dictionary<string, Guid>());
                                    dicFC[dateValue.Date][dr["PCODE"].ToString()]
                                        .Add(dr["SCODE"].ToString(), _supplier.SupplierId);
                                }

                                // Filling in data
                                _ListSupplierForecast = _ProductForecast.ListSupplierForecast;

                                // Special part for ThuMua
                                var myTI = new CultureInfo("en-US", false).TextInfo;
                                if (SupplierType != "VinEco" && !YesNoKPI)
                                {
                                    _SupplierForecast.QualityControlPass = string.IsNullOrEmpty(dr["QC"].ToString())
                                        ? false
                                        : (myTI.ToTitleCase(dr["QC"].ToString()) == "Ok" ? true : false);
                                    _SupplierForecast.LabelVinEco = string.IsNullOrEmpty(dr["Label VE"].ToString())
                                        ? false
                                        : (myTI.ToTitleCase(dr["Label VE"].ToString()) == "Yes" ? true : false);
                                    _SupplierForecast.FullOrder = string.IsNullOrEmpty(dr["100%"].ToString())
                                        ? false
                                        : (myTI.ToTitleCase(dr["100%"].ToString()) == "Yes" ? true : false);
                                    _SupplierForecast.CrossRegion = string.IsNullOrEmpty(dr["CrossRegion"].ToString())
                                        ? false
                                        : (myTI.ToTitleCase(dr["CrossRegion"].ToString()) == "Yes" ? true : false);
                                    _SupplierForecast.Level = string.IsNullOrEmpty(dr["Level"].ToString())
                                        ? Convert.ToByte(0)
                                        : Convert.ToByte(dr["Level"]);
                                    _SupplierForecast.Availability = string.IsNullOrEmpty(dr["Availability"].ToString())
                                        ? ""
                                        : dr["Availability"].ToString();
                                }
                                else if (!YesNoKPI)
                                {
                                    _SupplierForecast.QualityControlPass = true;
                                    _SupplierForecast.LabelVinEco = true;
                                    _SupplierForecast.FullOrder = false;
                                    _SupplierForecast.CrossRegion = false;
                                    _SupplierForecast.Level = 1;
                                    _SupplierForecast.Availability = "1234567";

                                    // To deal with some Supplier only Supply for a targetted Customer Group.
                                    _SupplierForecast.Target =
                                        dt.Columns.Contains("Target") ? dr["Target"].ToString() : "All";
                                }
                                else if (YesNoKPI && dr["Source"].ToString() == "ThuMua")
                                {
                                    _SupplierForecast.QualityControlPass = string.IsNullOrEmpty(dr["QC"].ToString())
                                        ? _SupplierForecast.QualityControlPass
                                        : (myTI.ToTitleCase(dr["QC"].ToString()) == "Ok" ? true : false);
                                    _SupplierForecast.LabelVinEco = string.IsNullOrEmpty(dr["Label VE"].ToString())
                                        ? _SupplierForecast.LabelVinEco
                                        : (myTI.ToTitleCase(dr["Label VE"].ToString()) == "Yes" ? true : false);
                                    _SupplierForecast.FullOrder = string.IsNullOrEmpty(dr["100%"].ToString())
                                        ? _SupplierForecast.FullOrder
                                        : (myTI.ToTitleCase(dr["100%"].ToString()) == "Yes" ? true : false);
                                    _SupplierForecast.CrossRegion = string.IsNullOrEmpty(dr["CrossRegion"].ToString())
                                        ? _SupplierForecast.CrossRegion
                                        : (myTI.ToTitleCase(dr["CrossRegion"].ToString()) == "Yes" ? true : false);
                                    _SupplierForecast.Level = string.IsNullOrEmpty(dr["Level"].ToString())
                                        ? _SupplierForecast.Level
                                        : Convert.ToByte(dr["Level"]);
                                    _SupplierForecast.Availability = string.IsNullOrEmpty(dr["Availability"].ToString())
                                        ? _SupplierForecast.Availability
                                        : dr["Availability"].ToString();
                                }

                                if (SupplierType == "VinEco" && dr["PCODE"].ToString().Substring(0, 1) == "K" &&
                                    (dr["Region"].ToString() == "MN" || dr["Region"].ToString() == "Miền Nam")
                                ) //dicCrossRegionVinEco.ContainsKey(dr["PCODE"].ToString()))
                                {
                                    _SupplierForecast.CrossRegion = true;
                                    if (dr["PCODE"].ToString() == "K03501") _SupplierForecast.CrossRegion = false;
                                }

                                if (dr["PCODE"].ToString() == "K01901") _SupplierForecast.CrossRegion = false;
                                if (dr["PCODE"].ToString() == "K02201") _SupplierForecast.CrossRegion = false;

                                ///// <! For debugging purposes !>
                                //if (dateValue.Day == 16 && (string)dr["PCODE"] == "A04201" && (string)dr["SCODE"] == "AG03030000")
                                //{
                                //    byte AmIHandsome = 0;
                                //}

                                // 3rd FC layer - Normal Forecast.
                                if (double.TryParse(
                                    (dr[dc.ColumnName] == DBNull.Value ? 0 : dr[dc.ColumnName]).ToString(),
                                    out double _QuantityForecast))
                                {
                                    if (!YesNoKPI)
                                        _SupplierForecast.QuantityForecast += _QuantityForecast;

                                    // 2nd FC layer - Minimum / Contracted Forecast - 2nd Highest Priority. 
                                    if (dt.Columns.Contains("Min"))
                                    {
                                        if (double.TryParse((dr["Min"] == DBNull.Value ? 0 : dr["Min"]).ToString(),
                                            out double _QuantityForecastContracted))
                                            _SupplierForecast.QuantityForecastContracted += _QuantityForecastContracted;
                                    }
                                }

                                //if (YesNoKPI &&
                                //    Convert.ToDateTime(dr["EffectiveFrom"]).Date <=
                                //    DateTime.Parse(dc.ColumnName).Date && Convert.ToDateTime(dr["EffectiveTo"]).Date >=
                                //    DateTime.Parse(dc.ColumnName).Date)
                                if (YesNoKPI && 
                                    StringToDate(dr["EffectiveFrom"].ToString())?.Date <= dateValue.Date &&
                                    StringToDate(dr["EffectiveTo"].ToString())?.Date >= dateValue.Date)
                                {
                                    _SupplierForecast.QualityControlPass = true;
                                    if (double.TryParse(
                                        (dr[dc.ColumnName] == DBNull.Value ? 0 : dr[dc.ColumnName]).ToString(),
                                        out double _QuantityForecastPlanned))
                                    {
                                        _SupplierForecast.QuantityForecastPlanned =
                                            _SupplierForecast.QuantityForecastPlanned ?? 0;
                                        _SupplierForecast.QuantityForecastPlanned += _QuantityForecastPlanned;

                                        // In case outside of Forecast, which, is an entirely new Supplier.
                                        // Yes this does happen.
                                        _SupplierForecast.QualityControlPass = true;

                                        //_SupplierForecast.QuantityForecastContracted = Math.Max(_SupplierForecast.QuantityForecastContracted - _SupplierForecast.QuantityForecastPlanned, 0);
                                        //_SupplierForecast.QuantityForecast = Math.Max(_SupplierForecast.QuantityForecast - _SupplierForecast.QuantityForecastPlanned - _SupplierForecast.QuantityForecastContracted, 0);
                                    }
                                }

                                if (isNewCustomerOrder) _ListSupplierForecast.Add(_SupplierForecast);

                                _ProductForecast.ListSupplierForecast = _ListSupplierForecast;
                                if (isNewProductOrder) _FC.ListProductForecast.Add(_ProductForecast);
                            }
                        }

                        _FC.ListProductForecast = _listProductForecast;

                        if (isNewFC) FC.Add(_FC);
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        ///     Do naughty stuff with PO
        /// </summary>
        private void EatPO(List<PurchaseOrderDate> PO, Range xlRng, Worksheet xlWs, string conStr,
            string PORegion, Dictionary<DateTime, Dictionary<string, Dictionary<string, Guid>>> dicPO,
            Dictionary<string, Product> dicProduct, Dictionary<string, Customer> dicCustomer, List<Product> Product,
            List<Customer> Customer, bool YesNoNew = false)
        {
            try
            {
                //WriteToRichTextBoxOutput(PORegion);
                //WriteToRichTextBoxOutput(xlWs.Name.ToString());

                var dicUnit = new Dictionary<string, string>(StringComparer.Ordinal);

                var dt = new DataTable();
                // Find first row
                var rowIndex = 0;
                do
                {
                    rowIndex++;
                } while (Convert.ToString(xlRng.Cells[rowIndex + 1, 1].Value) != "VE Code");

                xlRng = xlRng.Offset[rowIndex, 0].Resize[xlRng.Rows.Count - rowIndex, xlRng.Columns.Count];

                var oleCon = new OleDbConnection(conStr);

                var _oleAdapt =
                    new OleDbDataAdapter(
                        "Select * From [" + xlWs.Name + "$" +
                        xlRng.Address[false, false, XlReferenceStyle.xlA1, xlRng] + "]", oleCon);
                var _str = xlRng.Offset[rowIndex, 0].Address;
                WriteToRichTextBoxOutput(_str);
                _oleAdapt.Fill(dt);

                oleCon.Close();

                _oleAdapt = null;
                oleCon = null;

                var mongoClient = new MongoClient();
                var db = mongoClient.GetDatabase("localtest").GetCollection<PurchaseOrderDate>("PurchaseOrder");

                // Main Loop
                foreach (DataColumn dc in dt.Columns)
                {
                    DateTime dateValue;

                    // Loop for every column. Only stop at column with Date at the top ( Indicating it being PurchaseOrder for that Date )
                    if (DateTime.TryParse(dc.ColumnName,
                        out dateValue) /* && (dateValue.Date >= DateTime.Today.AddDays(0).Date)*/)
                    {
                        PurchaseOrderDate _PODate = null;
                        if (YesNoNew)
                            PO.RemoveAll(x => x.DateOrder.Date == dateValue);
                        //else
                        //{
                        //    _PODate = db.Find(x => x.DateOrder.Date == dateValue).FirstOrDefault();
                        //}
                        var isNewPODate = false;

                        Dictionary<string, Dictionary<string, Guid>> _dicProduct = null;
                        // Find PurchaseOrder for that Date
                        if (dicPO.TryGetValue(dateValue.Date, out _dicProduct))
                        {
                            _PODate = PO.Where(x => x.DateOrder.Date == dateValue.Date).FirstOrDefault();
                        }
                        // Create a blank one in case it doesn't exist
                        else
                        {
                            isNewPODate = true;

                            _PODate = new PurchaseOrderDate();
                            _PODate._id = Guid.NewGuid();
                            _PODate.PurchaseOrderDateId = _PODate._id;

                            _PODate.DateOrder = dateValue.Date;
                            _PODate.ListProductOrder = new List<ProductOrder>();

                            dicPO.Add(dateValue.Date, new Dictionary<string, Dictionary<string, Guid>>());
                        }

                        // First layer
                        // Get the list of all Products being Ordered that day.
                        var _listProductOrder = _PODate.ListProductOrder;
                        if (_listProductOrder == null) _listProductOrder = new List<ProductOrder>();

                        // Loop for every value
                        foreach (DataRow dr in dt.Rows)
                        {
                            // If OrderQuantity is not 0 - Not Anymore?
                            //object _OrderQuantity = dr[dc.ColumnName];
                            double _value = 0;
                            if (dr["VE Code"] != DBNull.Value && dr[dt.Columns.IndexOf(dc)] != DBNull.Value &&
                                double.TryParse(dr[dt.Columns.IndexOf(dc)].ToString(), out _value)
                            ) //&& Convert.ToDouble(dr[dc.ColumnName]) > 0)
                                if (_value > 0)
                                {
                                    List<CustomerOrder> _listCustomerOrder = null;
                                    CustomerOrder _CustomerOrder = null;
                                    ProductOrder _productOrder = null;

                                    var isNewProductOrder = false;
                                    var isNewCustomerOrder = false;

                                    Dictionary<string, Guid> dicStore = null;

                                    _dicProduct = dicPO[dateValue.Date];
                                    if (_dicProduct.TryGetValue(dr["VE Code"].ToString(), out dicStore))
                                    {
                                        Product _product = null;
                                        if (!dicProduct.TryGetValue(dr["VE Code"].ToString(), out _product))
                                        {
                                            _product = new Product();

                                            _product._id = Guid.NewGuid();
                                            _product.ProductId = _product._id;
                                            _product.ProductCode = dr["VE Code"].ToString();
                                            _product.ProductName = dr["VE Name"].ToString();

                                            Product.Add(_product);

                                            dicProduct.Add(_product.ProductCode, _product);
                                        }
                                        _productOrder = _PODate.ListProductOrder
                                            .Where(x => x.ProductId == _product.ProductId).FirstOrDefault();

                                        Guid _id;
                                        if (dicStore.TryGetValue(
                                            dr["StoreCode"] + (dt.Columns.Contains("P&L")
                                                ? dr["P&L"].ToString()
                                                : dr["StoreType"].ToString()), out _id))
                                        {
                                            _CustomerOrder = _productOrder.ListCustomerOrder
                                                .Where(x => x.CustomerId == _id).FirstOrDefault();
                                        }
                                        else
                                        {
                                            isNewCustomerOrder = true;

                                            Customer _customer;
                                            var sKey = dr["StoreCode"] + (dt.Columns.Contains("P&L")
                                                           ? dr["P&L"].ToString()
                                                           : dr["StoreType"].ToString());
                                            if (!dicCustomer.TryGetValue(sKey, out _customer))
                                            {
                                                _customer = new Customer();

                                                _customer._id = Guid.NewGuid();
                                                _customer.CustomerId = _customer._id;
                                                _customer.CustomerCode = dr["StoreCode"].ToString();
                                                _customer.CustomerName = dr["StoreName"].ToString();
                                                _customer.CustomerRegion = dr["Region"].ToString();
                                                _customer.CustomerType = dr["StoreType"].ToString();
                                                _customer.Company = dt.Columns.Contains("P&L")
                                                    ? dr["P&L"].ToString()
                                                    : "VinCommerce";
                                                _customer.CustomerBigRegion = PORegion;

                                                Customer.Add(_customer);

                                                dicCustomer.Add(sKey, _customer);
                                            }

                                            var _NewGuid = Guid.NewGuid();
                                            _CustomerOrder = new CustomerOrder
                                            {
                                                _id = _NewGuid,
                                                CustomerOrderId = _NewGuid,
                                                CustomerId = _customer.CustomerId
                                            };

                                            dicPO[dateValue.Date][dr["VE Code"].ToString()]
                                                .Add(sKey, _customer.CustomerId);
                                        }
                                    }
                                    else
                                    {
                                        isNewProductOrder = true;
                                        isNewCustomerOrder = true;

                                        Product _product = null;
                                        if (!dicProduct.TryGetValue(dr["VE Code"].ToString(), out _product))
                                        {
                                            _product = new Product
                                            {
                                                _id = Guid.NewGuid(),
                                                ProductCode = dr["VE Code"].ToString(),
                                                ProductName = dr["VE Name"].ToString()
                                            };

                                            _product.ProductId = _product._id;


                                            Product.Add(_product);

                                            dicProduct.Add(_product.ProductCode, _product);
                                        }

                                        _productOrder = new ProductOrder();
                                        _productOrder._id = Guid.NewGuid();
                                        _productOrder.ProductOrderId = _productOrder._id;
                                        _productOrder.ProductId = _product.ProductId;

                                        _productOrder.ListCustomerOrder = new List<CustomerOrder>();

                                        Customer _customer;
                                        var sKey = dr["StoreCode"] + (dt.Columns.Contains("P&L")
                                                       ? dr["P&L"].ToString()
                                                       : dr["StoreType"].ToString());
                                        if (!dicCustomer.TryGetValue(sKey, out _customer))
                                        {
                                            _customer = new Customer();

                                            _customer._id = Guid.NewGuid();
                                            _customer.CustomerId = _customer._id;
                                            _customer.CustomerCode = dr["StoreCode"].ToString();
                                            _customer.CustomerName = dr["StoreName"].ToString();
                                            _customer.CustomerRegion = dr["Region"].ToString();
                                            _customer.CustomerType = dr["StoreType"].ToString();
                                            _customer.Company = dt.Columns.Contains("P&L")
                                                ? dr["P&L"].ToString()
                                                : "VinCommerce";
                                            _customer.CustomerBigRegion = PORegion;

                                            Customer.Add(_customer);

                                            dicCustomer.Add(sKey, _customer);
                                        }

                                        _CustomerOrder = new CustomerOrder();
                                        _CustomerOrder._id = Guid.NewGuid();
                                        _CustomerOrder.CustomerOrderId = _CustomerOrder._id;
                                        _CustomerOrder.CustomerId = _customer.CustomerId;

                                        dicPO[dateValue.Date].Add(dr["VE Code"].ToString(),
                                            new Dictionary<string, Guid>());
                                        dicPO[dateValue.Date][dr["VE Code"].ToString()].Add(sKey, _customer.CustomerId);
                                    }

                                    // Filling in data
                                    _listCustomerOrder = _productOrder.ListCustomerOrder;

                                    // Desired Region
                                    if (dt.Columns.Contains("Vùng sản xuất") && dr["Vùng sản xuất"] != null)
                                    {
                                        var _DesiredRegion = dr["Vùng sản xuất"].ToString();

                                        if (_DesiredRegion != "" &&
                                            (_DesiredRegion == "Lâm Đồng" || _DesiredRegion == "Miền Bắc" ||
                                             _DesiredRegion == "Miền Nam"))
                                            _CustomerOrder.DesiredRegion = _DesiredRegion;
                                    }

                                    // Desired Source
                                    if (dt.Columns.Contains("Nguồn") && dr["Nguồn"] != null)
                                    {
                                        var _DesiredSource = dr["Nguồn"].ToString();

                                        if (_DesiredSource != "" &&
                                            (_DesiredSource == "VinEco" || _DesiredSource == "ThuMua" ||
                                             _DesiredSource == "VCM"))
                                            _CustomerOrder.DesiredSource = _DesiredSource;
                                    }

                                    _CustomerOrder.Unit =
                                        ProperUnit(dr["Unit"].ToString() == "" ? "Kg" : dr["Unit"].ToString(), dicUnit);
                                    _CustomerOrder.QuantityOrder += _value;

                                    if (isNewCustomerOrder) _listCustomerOrder.Add(_CustomerOrder);

                                    _productOrder.ListCustomerOrder = _listCustomerOrder;
                                    if (isNewProductOrder) _PODate.ListProductOrder.Add(_productOrder);
                                }
                        }

                        _PODate.ListProductOrder = _listProductOrder;

                        if (isNewPODate) PO.Add(_PODate);

                        //WriteToRichTextBoxOutput(Region + " " + dc.ColumnName + ": " + sumColumn);
                        //WriteToRichTextBoxOutput(PO.Where(x => x == _PODate).FirstOrDefault().ListProductOrder.Sum(po => po.ListCustomerOrder.Sum(co => co.QuantityOrder)));
                        //WriteToRichTextBoxOutput("MB " + dc.ColumnName + ": " + PO.Where(x => x.DateOrder.Date.ToString() == dc.ColumnName).FirstOrDefault().ListProductOrder.Sum(x => x.ListCustomerOrder.Where(co => dicCustomer.Values.Where(_co => _co.CustomerId == co.CustomerId).FirstOrDefault().CustomerBigRegion == "Miền Bắc").Sum(o => o.QuantityOrder)));
                        //WriteToRichTextBoxOutput("MN " + dc.ColumnName + ": " + PO.Where(x => x.DateOrder.Date.ToString() == dc.ColumnName).FirstOrDefault().ListProductOrder.Sum(x => x.ListCustomerOrder.Where(co => dicCustomer.Values.Where(_co => _co.CustomerId == co.CustomerId).FirstOrDefault().CustomerBigRegion == "Miền Nam").Sum(o => o.QuantityOrder)));
                    }
                }

                #region Old method

                //var dicHeader = new Dictionary<string, int>();
                //
                //// Dictionary of Columns in Destination DataTable
                //foreach (DataColumn dc in database.Columns)
                //{
                //    try { dicHeader.Add(dc.ColumnName, 0); }
                //    catch (Exception) { }
                //}

                //// Dictionary of Columns in Targeted DataTable
                //foreach (DataColumn dc in dt.Columns)
                //{
                //    int _colIndex;
                //    if (dicHeader.TryGetValue(dc.ColumnName, out _colIndex))
                //    {
                //        dicHeader[dc.ColumnName] = dt.Columns.IndexOf(dc);
                //    }
                //}

                //// Main loop
                //foreach (DataRow dr in dt.Rows)
                //{
                //    for (int colIndex = 0; colIndex < dt.Columns.Count; colIndex++)
                //    {
                //        DateTime dateValue.Date;
                //        if (DateTime.TryParse(dt.Columns[colIndex].ColumnName, out dateValue.Date) &&
                //            dateValue.Date >= DateFrom &&
                //            dateValue.Date <= DateTo &&
                //            dr[colIndex] != null &&
                //            dr[colIndex].ToString().Length > 0 &&
                //            Convert.ToDouble(dr[colIndex]) > 0)
                //        {

                //            DataRow drNew = database.NewRow();

                //            //int _count = 0;
                //            for (int colPos = 0; colPos < dt.Columns.Count; colPos++)
                //            {
                //                int _colIndex;
                //                if (dicHeader.TryGetValue(dt.Columns[colPos].ColumnName, out _colIndex))
                //                {
                //                    if (dr[_colIndex].ToString().Length == 0)
                //                    {
                //                        drNew[colPos] = "";
                //                    }
                //                    else
                //                    {
                //                        drNew[colPos] = dr[_colIndex];
                //                    }
                //                    //_count++;
                //                    //if (_count > database.Columns.Count - 2) { break; }
                //                }
                //            }

                //            drNew["PO Region"] = PORegion;
                //            drNew["OrderDate"] = dateValue.Date;
                //            drNew["OrderQuantity"] = Convert.ToDouble(dr[colIndex]);

                //            database.Rows.Add(drNew);

                //        }
                //    }
                //}

                #endregion
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        ///     Do naughty stuff with PO
        /// </summary>
        private void EatPOAspose(List<PurchaseOrderDate> PO, Aspose.Cells.Worksheet xlWs, string conStr,
            string PORegion, Dictionary<DateTime, Dictionary<string, Dictionary<string, Guid>>> dicPO,
            Dictionary<string, Product> dicProduct, Dictionary<string, Customer> dicCustomer, List<Product> Product,
            List<Customer> Customer, bool YesNoNew = false)
        {
            try
            {
                var stopwatch = Stopwatch.StartNew();

                Debug.WriteLine($"File: {xlWs.Workbook.FileName}");

                var dicUnit = new Dictionary<string, string>(StringComparer.Ordinal);
                // Find first row.
                var rowIndex = 0;
                var colIndex = 0;
                var value = string.Empty;
                do
                {
                    value = xlWs.Cells[rowIndex, colIndex].Value?.ToString().Trim();

                    if (value == "VE Code" || value == "Mã Planning" || value == "Mã Planing")
                        break;

                    //if (value == null || value == string.Empty || (value != "VE Code" && value != "Mã Planning"))
                    //    rowIndex++;
                    //else
                    //    break;

                    rowIndex++;

                    if (rowIndex > 100)
                    {
                        colIndex++;
                        if (colIndex > 100) break;
                        rowIndex = 0;
                    }
                } while ((value == null || value == string.Empty || (value != "VE Code" && value != "Mã Planning")) &&
                         rowIndex <= 100 &&
                         colIndex <= 100);

                if (rowIndex > 100 || colIndex > 100)
                {
                    Debug.WriteLine($"Sai định dạng - {xlWs.Workbook.FileName}");
                    WriteToRichTextBoxOutput($"Sai định dạng - {xlWs.Workbook.FileName}");
                    return;
                }

                // ... ah well, option based 0.
                //rowIndex--;

                // Import into a DataTable.
                var opts = new ExportTableOptions
                {
                    CheckMixedValueType = true,
                    ExportAsString = false,
                    FormatStrategy = CellValueFormatStrategy.None,
                    ExportColumnName = true
                };

                var dt = new DataTable { TableName = xlWs.Name };
                dt = xlWs.Cells.ExportDataTable(rowIndex, colIndex, xlWs.Cells.MaxDataRow + 1, xlWs.Cells.MaxDataColumn + 1,
                    opts);

                //var mongoClient = new MongoClient();
                //var db = mongoClient.GetDatabase("localtest").GetCollection<PurchaseOrderDate>("PurchaseOrder");

                if (!dt.Columns.Contains("VE Code"))
                    dt.Columns[colIndex].ColumnName = "VE Code";

                if (dt.Columns.Contains("Tỉnh tiêu thụ"))
                    dt.Columns["Tỉnh tiêu thụ"].ColumnName = "Region";

                if (dt.Columns.Contains("Store Code")) dt.Columns["Store Code"].ColumnName = "StoreCode";
                if (dt.Columns.Contains("Store Name")) dt.Columns["Store Name"].ColumnName = "StoreName";
                if (dt.Columns.Contains("Store Type")) dt.Columns["Store Type"].ColumnName = "StoreType";

                // Main Loop
                foreach (DataColumn dc in dt.Columns)
                {
                    // Loop for every column. Only stop at column with Date at the top ( Indicating it being PurchaseOrder for that Date )
                    //if (DateTime.TryParse(dc.ColumnName,
                    //    out dateValue) /* && (dateValue.Date >= DateTime.Today.AddDays(0).Date)*/)
                    if (StringToDate(dc.ColumnName) != null)
                    {
                        DateTime dateValue = StringToDate(dc.ColumnName) ?? DateTime.MinValue;

                        PurchaseOrderDate _PODate = null;
                        if (YesNoNew)
                            PO.RemoveAll(x => x.DateOrder.Date == dateValue);
                        //else
                        //{
                        //    _PODate = db.Find(x => x.DateOrder.Date == dateValue).FirstOrDefault();
                        //}
                        var isNewPODate = false;

                        Dictionary<string, Dictionary<string, Guid>> _dicProduct = null;
                        // Find PurchaseOrder for that Date
                        if (dicPO.TryGetValue(dateValue.Date, out _dicProduct))
                        {
                            _PODate = PO.Where(x => x.DateOrder.Date == dateValue.Date).FirstOrDefault();
                        }
                        // Create a blank one in case it doesn't exist
                        else
                        {
                            isNewPODate = true;

                            _PODate = new PurchaseOrderDate();
                            _PODate._id = Guid.NewGuid();
                            _PODate.PurchaseOrderDateId = _PODate._id;

                            _PODate.DateOrder = dateValue.Date;
                            _PODate.ListProductOrder = new List<ProductOrder>();

                            dicPO.Add(dateValue.Date, new Dictionary<string, Dictionary<string, Guid>>());
                        }

                        // First layer
                        // Get the list of all Products being Ordered that day.
                        var _listProductOrder = _PODate.ListProductOrder;
                        if (_listProductOrder == null) _listProductOrder = new List<ProductOrder>();

                        // Loop for every value
                        foreach (DataRow dr in dt.Rows)
                        {
                            if (dr["VE Code"] != DBNull.Value && dr["VE Code"].ToString().Substring(0, 1) == "9")
                            {
                                byte whatInTheActualFuck = 0;
                                whatInTheActualFuck++;
                            }
                            // If OrderQuantity is not 0 - Not Anymore?
                            //object _OrderQuantity = dr[dc.ColumnName];
                            double _value = 0;
                            if (dr["VE Code"] != DBNull.Value && dr[dt.Columns.IndexOf(dc)] != DBNull.Value &&
                                double.TryParse(dr[dt.Columns.IndexOf(dc)].ToString(), out _value)
                            ) //&& Convert.ToDouble(dr[dc.ColumnName]) > 0)
                                if (_value > 0)
                                {
                                    List<CustomerOrder> _listCustomerOrder = null;
                                    CustomerOrder _CustomerOrder = null;
                                    ProductOrder _productOrder = null;

                                    var isNewProductOrder = false;
                                    var isNewCustomerOrder = false;

                                    _dicProduct = dicPO[dateValue.Date];
                                    if (_dicProduct.TryGetValue(dr["VE Code"].ToString(), out var dicStore))
                                    {
                                        if (!dicProduct.TryGetValue(dr["VE Code"].ToString(), out var _product))
                                        {
                                            _product = new Product
                                            {
                                                _id = Guid.NewGuid(),
                                                ProductCode = dr["VE Code"].ToString(),
                                                ProductName = dr["VE Name"].ToString()
                                            };

                                            _product.ProductId = _product._id;

                                            Product.Add(_product);

                                            dicProduct.Add(_product.ProductCode, _product);
                                        }
                                        _productOrder = _PODate.ListProductOrder
                                            .Where(x => x.ProductId == _product.ProductId).FirstOrDefault();

                                        Guid _id;
                                        if (dicStore.TryGetValue(
                                            dr["StoreCode"] + (dt.Columns.Contains("P&L")
                                                ? dr["P&L"].ToString()
                                                : dr["StoreType"].ToString()), out _id))
                                        {
                                            _CustomerOrder = _productOrder.ListCustomerOrder
                                                .Where(x => x.CustomerId == _id).FirstOrDefault();
                                        }
                                        else
                                        {
                                            isNewCustomerOrder = true;

                                            Customer _customer;
                                            var sKey = dr["StoreCode"] + (dt.Columns.Contains("P&L")
                                                           ? dr["P&L"].ToString()
                                                           : dr["StoreType"].ToString());
                                            if (!dicCustomer.TryGetValue(sKey, out _customer))
                                            {
                                                _customer = new Customer();

                                                _customer._id = Guid.NewGuid();
                                                _customer.CustomerId = _customer._id;
                                                _customer.CustomerCode = dr["StoreCode"].ToString();
                                                _customer.CustomerName = dr["StoreName"].ToString();
                                                _customer.CustomerRegion = dr["Region"].ToString();
                                                _customer.CustomerType = dr["StoreType"].ToString();
                                                _customer.Company = dt.Columns.Contains("P&L")
                                                    ? dr["P&L"].ToString()
                                                    : "VinCommerce";
                                                _customer.CustomerBigRegion = PORegion;

                                                Customer.Add(_customer);

                                                dicCustomer.Add(sKey, _customer);
                                            }

                                            var _NewGuid = Guid.NewGuid();
                                            _CustomerOrder = new CustomerOrder
                                            {
                                                _id = _NewGuid,
                                                CustomerOrderId = _NewGuid,
                                                CustomerId = _customer.CustomerId
                                            };

                                            dicPO[dateValue.Date][dr["VE Code"].ToString()]
                                                .Add(sKey, _customer.CustomerId);
                                        }
                                    }
                                    else
                                    {
                                        isNewProductOrder = true;
                                        isNewCustomerOrder = true;

                                        Product _product = null;
                                        if (!dicProduct.TryGetValue(dr["VE Code"].ToString(), out _product))
                                        {
                                            _product = new Product();

                                            _product._id = Guid.NewGuid();
                                            _product.ProductId = _product._id;
                                            _product.ProductCode = dr["VE Code"].ToString();
                                            _product.ProductName = dr["VE Name"].ToString();

                                            Product.Add(_product);

                                            dicProduct.Add(_product.ProductCode, _product);
                                        }

                                        _productOrder = new ProductOrder();
                                        _productOrder._id = Guid.NewGuid();
                                        _productOrder.ProductOrderId = _productOrder._id;
                                        _productOrder.ProductId = _product.ProductId;

                                        _productOrder.ListCustomerOrder = new List<CustomerOrder>();

                                        Customer _customer;
                                        var sKey = dr["StoreCode"] + (dt.Columns.Contains("P&L")
                                                       ? dr["P&L"].ToString()
                                                       : dr["StoreType"].ToString());
                                        if (!dicCustomer.TryGetValue(sKey, out _customer))
                                        {
                                            _customer = new Customer();

                                            _customer._id = Guid.NewGuid();
                                            _customer.CustomerId = _customer._id;
                                            _customer.CustomerCode = dr["StoreCode"].ToString();
                                            _customer.CustomerName = dr["StoreName"].ToString();
                                            _customer.CustomerRegion = dr["Region"].ToString();
                                            _customer.CustomerType = dr["StoreType"].ToString();
                                            _customer.Company = dt.Columns.Contains("P&L")
                                                ? dr["P&L"].ToString()
                                                : "VinCommerce";
                                            _customer.CustomerBigRegion = PORegion;

                                            Customer.Add(_customer);

                                            dicCustomer.Add(sKey, _customer);
                                        }

                                        _CustomerOrder = new CustomerOrder {_id = Guid.NewGuid()};
                                        _CustomerOrder.CustomerOrderId = _CustomerOrder._id;
                                        _CustomerOrder.CustomerId = _customer.CustomerId;

                                        dicPO[dateValue.Date].Add(dr["VE Code"].ToString(),
                                            new Dictionary<string, Guid>());
                                        dicPO[dateValue.Date][dr["VE Code"].ToString()].Add(sKey, _customer.CustomerId);
                                    }

                                    // Filling in data
                                    _listCustomerOrder = _productOrder.ListCustomerOrder;

                                    // Desired Region
                                    if (dt.Columns.Contains("Vùng sản xuất") && dr["Vùng sản xuất"] != null)
                                    {
                                        var _DesiredRegion = dr["Vùng sản xuất"].ToString();

                                        if (_DesiredRegion != "" &&
                                            (_DesiredRegion == "Lâm Đồng" || _DesiredRegion == "Miền Bắc" ||
                                             _DesiredRegion == "Miền Nam"))
                                            _CustomerOrder.DesiredRegion = _DesiredRegion;
                                    }

                                    // Desired Source
                                    if (dt.Columns.Contains("Nguồn") && dr["Nguồn"] != null)
                                    {
                                        var _DesiredSource = dr["Nguồn"].ToString();

                                        if (_DesiredSource != "" &&
                                            (_DesiredSource == "VinEco" || _DesiredSource == "ThuMua" ||
                                             _DesiredSource == "VCM"))
                                            _CustomerOrder.DesiredSource = _DesiredSource;
                                    }

                                    _CustomerOrder.Unit =
                                        ProperUnit(dr["Unit"].ToString() == "" ? "Kg" : dr["Unit"].ToString(), dicUnit);
                                    _CustomerOrder.QuantityOrder += _value;

                                    if (isNewCustomerOrder) _listCustomerOrder.Add(_CustomerOrder);

                                    _productOrder.ListCustomerOrder = _listCustomerOrder;
                                    if (isNewProductOrder) _PODate.ListProductOrder.Add(_productOrder);
                                }
                        }

                        _PODate.ListProductOrder = _listProductOrder;

                        if (isNewPODate) PO.Add(_PODate);
                    }
                }

                stopwatch.Stop();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        ///     Reading Deli into Database for Storing purposes.
        /// </summary>
        private async Task UpdateDeli()
        {
            try
            {
                // Grab the Database by the tail.
                var db = new MongoClient().GetDatabase("localtest");

                // Initialize stuff.
                var ListAllo = new List<AllocateDetail>();
                var Product = db.GetCollection<Product>("Product").AsQueryable().ToList();
                var Customer = new List<Customer>();

                // Core!
                var core = new CoordStructure();

                // ... and of course, core stuff.
                var dicPO = new Dictionary<DateTime, Dictionary<string, Dictionary<string, Guid>>>();
                core.dicProduct = new Dictionary<Guid, Product>();
                core.dicCustomer = new Dictionary<Guid, Customer>();

                // Product Dictionary.
                foreach (var _Product in Product)
                    if (!core.dicProduct.ContainsKey(_Product.ProductId))
                        core.dicProduct.Add(_Product.ProductId, _Product);

                // Customer Dictionary.
                foreach (var _Customer in Customer)
                    if (!core.dicCustomer.ContainsKey(_Customer.CustomerId))
                        core.dicCustomer.Add(_Customer.CustomerId, _Customer);

                // Directory.
                // Todo - Hardcoded, need to change.
                var directoryPath =
                    "D:\\Documents\\Stuff\\VinEco\\Mastah Project\\Deli";

                #region Reading PO files in folder.

                var dirInfo = new DirectoryInfo(directoryPath);
                var ListFile = dirInfo.GetFiles();

                foreach (var _FileInfo in ListFile)
                {
                    var xlWb = new Workbook(_FileInfo.FullName);

                    EatDeli(xlWb, core);
                }

                #endregion

                WriteToRichTextBoxOutput("Here goes pain");

                db.DropCollection("AllocateDetail");
                await db.GetCollection<AllocateDetail>("AllocateDetail").InsertManyAsync(ListAllo);

                db = null;
            }
            catch (Exception ex)
            {
                throw ex;
                //MessageBox.Show(ex.Message, "Exception Error");
            }
            finally
            {
            }
        }

        /// <summary>
        ///     Naughty stuff with Deli files.
        /// </summary>
        private void EatDeli(Workbook xlWb, CoordStructure core)
        {
            try
            {
                // Grab the worksheet with the most data rows.
                // Lazy way, I know, but will dodge random unneeded table from Pivot.
                var xlWs = xlWb.Worksheets.OrderByDescending(x => x.Cells.MaxDataRow + 1).FirstOrDefault();

                // Find first row.
                var rowIndex = 0;
                do
                {
                    rowIndex++;
                } while (xlWs.Cells[rowIndex, 0].Value == null && rowIndex <= xlWs.Cells.MaxDataRow + 1);

                // ... ah well, option based 0.
                rowIndex--;

                // Import into a DataTable.
                var dt = new DataTable { TableName = xlWs.Name };
                dt = xlWs.Cells.ExportDataTable(rowIndex, 0, xlWs.Cells.MaxDataRow + 1, xlWs.Cells.MaxDataColumn + 1,
                    true);

                // Here we go.
                // Dissecting DataTable into Database.
                foreach (DataRow dr in dt.Rows)
                {
                    // Harvest Date.
                    var DateProcess = Convert.ToDateTime(dr["DATE_PROCESS"]);

                    // Order Date.
                    var DateOrder = Convert.ToDateTime(dr["DATE_ORDER"]);

                    // Product.
                    var _ProductCode = dr["PCODE1"].ToString().Substring(0, 6);
                    var _Product = core.dicProduct.Values.Where(x => x.ProductCode == _ProductCode).FirstOrDefault();

                    // Customer.
                    var _CustomerCode = (string)dr["CCODE"];
                    var _Customer = core.dicCustomer.Values.Where(x => x.CustomerCode == _CustomerCode)
                        .FirstOrDefault();

                    // Supplier.
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        #region Proper string

        /// <summary>
        ///     Proper a string
        /// </summary>
        public static string ProperStr(string myString)
        {
            // Creates a TextInfo based on the "en-US" culture.
            var myTI = new CultureInfo("en-US", false).TextInfo;

            //// Changes a string to lowercase.
            //WriteToRichTextBoxOutput("\"{0}\" to lowercase: {1}", myString, myTI.ToLower(myString));

            //// Changes a string to uppercase.
            //WriteToRichTextBoxOutput("\"{0}\" to uppercase: {1}", myString, myTI.ToUpper(myString));

            //// Changes a string to titlecase.
            //WriteToRichTextBoxOutput("\"{0}\" to titlecase: {1}", myString, myTI.ToTitleCase(myString));

            return myTI.ToTitleCase(myString.ToLower());
        }

        #endregion

        #region Proper UnitType.

        /// <summary>
        ///     Proper UnitType.
        /// </summary>
        /// <param name="Unit"></param>
        /// <returns></returns>
        private static string ProperUnit(string Unit, Dictionary<string, string> dicUnit)
        {
            if (dicUnit.TryGetValue(Unit, out var unit))
                return unit;

            // Initialize empty result.
            var _Unit = Unit.Trim().ToLower();

            // Looping through every letter.
            for (var stringIndex = 0; stringIndex < _Unit.Length; stringIndex++)
                // If a forward dash is found.
                if (_Unit.Substring(stringIndex, 1) == "/")
                {
                    // Insert a space if the letter right before it isn't already a space.
                    if (stringIndex != 0 && _Unit.Substring(stringIndex - 1, 1) != " ")
                        _Unit = _Unit.Insert(stringIndex - 1, " ");

                    // Insert a space if the letter right after it isn't already a space.
                    if (stringIndex != _Unit.Length && _Unit.Substring(stringIndex + 1, 1) != " ")
                        _Unit = _Unit.Insert(stringIndex + 1, " ");
                }

            // Creates a TextInfo based on the "en-US" culture.
            var myTI = new CultureInfo("en-US", false).TextInfo;

            dicUnit.Add(Unit, myTI.ToTitleCase(_Unit));

            // Return the "Proper" Unit.
            return myTI.ToTitleCase(_Unit);
        }

        #endregion

        private void WriteToRichTextBoxOutput(object Message = null, bool NewLine = true)
        {
            try
            {
                if (Message == null) Message = "";
                richTextBoxOutput.AppendText(string.Format("{0},{1}", Message.ToString(), NewLine ? "\n" : " "));
                richTextBoxOutput.Refresh();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private void RichTextBoxOutput_TextChanged(object sender, EventArgs e)
        {
            // set the current caret position to the end
            richTextBoxOutput.SelectionStart = richTextBoxOutput.Text.Length;
            // scroll it automatically
            richTextBoxOutput.ScrollToCaret();
        }

        /// <summary>
        ///     All Constants should be declared here
        /// </summary>
        private static class Constants
        {
            //public const string Excel03ConString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties='Excel 8.0;HDR={1};IMEX=1;'";
            public const string Excel07ConString =
                "Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='Excel 8.0;HDR={1};IMEX=1;'";
        }

        #region Global Variables

        private DateTime DateFrom = DateTime.Today;
        private DateTime DateTo = DateTime.Today;
        private double UpperCap = 1.4;
        private readonly byte dayDistance = 4;
        private bool FruitOnly;
        private bool NoFruit;
        private readonly bool YesPlanningFuckMe = false;
        private bool YesNoSubRegion;

        #endregion

        #region Behaviour

        private void DateFromPicker_ValueChanged(object sender, EventArgs e)
        {
            DateFrom = DateFromPicker.Value;
        }

        private void DateToPicker_ValueChanged(object sender, EventArgs e)
        {
            DateTo = DateToPicker.Value;
        }

        private void upperCapBox_TextChanged(object sender, EventArgs e)
        {
            try
            {
                if (upperCapBox.Text != null)
                {
                    var _UpperCap = UpperCap;
                    if (double.TryParse(upperCapBox.Text, out _UpperCap))
                        UpperCap = _UpperCap;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private void checkBoxFruitOnly_CheckedChanged(object sender, EventArgs e)
        {
            FruitOnly = checkBoxFruitOnly.Checked;
        }

        private void checkBoxNoFruit_CheckedChanged(object sender, EventArgs e)
        {
            NoFruit = checkBoxNoFruit.Checked;
        }

        private void YesNoSubRegion_CheckedChanged(object sender, EventArgs e)
        {
            YesNoSubRegion = YesNoSubRegionchkBox.Checked;
        }

        private async void readPOToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //openFileDialog1.ShowDialog();
            await UpdatePO("Forecast MB.xlsb", "Forecast MN.xlsb");
        }

        private async void readFCToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //openFileDialog1.ShowDialog();
            //UpdatePO("Forecast MB.xlsb", "Forecast MN.xlsb");
            await UpdateFC("DBSL.xlsb", "ThuMua.xlsb");
        }

        private async void readFCPlanningToolStripMenuItem_Click(object sender, EventArgs e)
        {
            await UpdateFC("DBSL.xlsb", "ThuMua Planning.xlsb", true);
        }

        private void kgToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FiteMoi(DateFrom, DateTo > DateFrom ? DateTo : DateFrom, false, false, true, false, false, false, false);
        }

        private void unitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FiteMoi(DateFrom, DateTo > DateFrom ? DateTo : DateFrom, false, false, true, false, false, false, true);
        }

        private void noSupToolStripMenuItem_Click(object sender, EventArgs e)
        {
            #region Old Print Result

            //#region Preparing!
            ////string connectionString = System.Configuration.ConfigurationManager.ConnectionStrings["mongodb_vecrops.salesms"].ConnectionString;
            ////MongoClient mongoClient = new MongoClient(connectionString);
            //var mongoClient = new MongoClient();
            //var db = mongoClient.GetDatabase("localtest");

            //var PO = db.GetCollection<PurchaseOrderDate>("PurchaseOrder").AsQueryable().ToList()
            //    .OrderBy(x => x.DateOrder.Date)
            //    .Where(x => x.DateOrder.Date >= DateFrom.Date & x.DateOrder.Date <= DateTo.Date);

            //var FC = db.GetCollection<ForecastDate>("Forecast").AsQueryable().ToList()
            //    .OrderBy(x => x.DateForecast)
            //    .Where(x => x.DateForecast.Date >= DateFrom.Date && x.DateForecast.Date <= DateTo.Date);

            //var Mastah = db.GetCollection<CoordResult>("CoordResult").AsQueryable().ToList().
            //        OrderBy(x => x.DateOrder.Date)
            //        .Where(x => x.DateOrder.Date >= DateFrom.Date && x.DateOrder.Date <= DateTo.Date);

            //var Product = db.GetCollection<Product>("Product").AsQueryable().ToList();
            //var Supplier = db.GetCollection<Supplier>("Supplier").AsQueryable().ToList();
            //var Customer = db.GetCollection<Customer>("Customer").AsQueryable().ToList();

            //var dicProduct = new Dictionary<Guid, Product>();
            //var dicSupplier = new Dictionary<Guid, Supplier>();
            //var dicCustomer = new Dictionary<Guid, Customer>();

            //foreach (var product in Product)
            //{
            //    dicProduct.Add(product.ProductId, product);

            //}

            //foreach (var supplier in Supplier)
            //{
            //    Supplier _supplier = null;
            //    if (!dicSupplier.TryGetValue(supplier.SupplierId, out _supplier))
            //    {
            //        dicSupplier.Add(supplier.SupplierId, supplier);
            //    }
            //}

            //foreach (var customer in Customer)
            //{
            //    Customer _customer = null;
            //    if (!dicCustomer.TryGetValue(customer.CustomerId, out _customer))
            //    {
            //        dicCustomer.Add(customer.CustomerId, customer);
            //    }
            //}

            //var dicCustomerOrder = new Dictionary<Guid, CustomerOrder>();
            //foreach (PurchaseOrderDate PODate in PO)
            //{
            //    foreach (ProductOrder _ProductOrder in PODate.ListProductOrder)
            //    {
            //        foreach (CustomerOrder _CustomerOrder in _ProductOrder.ListCustomerOrder)
            //        {
            //            dicCustomerOrder.Add(_CustomerOrder.CustomerOrderId, _CustomerOrder);
            //        }
            //    }
            //}

            //var dicSupplierForecast = new Dictionary<Guid, SupplierForecast>();
            //foreach (ForecastDate FCDate in FC)
            //{
            //    foreach (ProductForecast _ProductForecast in FCDate.ListProductForecast)
            //    {
            //        foreach (SupplierForecast _SupplierForecast in _ProductForecast.ListSupplierForecast)
            //        {
            //            dicSupplierForecast.Add(_SupplierForecast.SupplierForecastId, _SupplierForecast);
            //        }
            //    }
            //}

            //var dicCoordResult = new Dictionary<DateTime, Dictionary<Product, Dictionary<CustomerOrder, Dictionary<SupplierForecast, DateTime>>>>();
            //foreach (CoordResult _CoordResult in Mastah)
            //{
            //    dicCoordResult.Add(_CoordResult.DateOrder.Date, new Dictionary<Product, Dictionary<CustomerOrder, Dictionary<SupplierForecast, DateTime>>>());
            //    foreach (CoordResultDate _CoordResultDate in _CoordResult.ListCoordResultDate)
            //    {
            //        dicCoordResult[_CoordResult.DateOrder.Date].Add(dicProduct[_CoordResultDate.ProductId], new Dictionary<CustomerOrder, Dictionary<SupplierForecast, DateTime>>());
            //        foreach (CoordinateDate _CoordinateDate in _CoordResultDate.ListCoordinateDate)
            //        {

            //            if (_CoordinateDate.SupplierOrderId != null)
            //            {
            //                dicCoordResult[_CoordResult.DateOrder.Date][dicProduct[_CoordResultDate.ProductId]].Add(dicCustomerOrder[_CoordinateDate.CustomerOrderId], new Dictionary<SupplierForecast, DateTime>());
            //                dicCoordResult[_CoordResult.DateOrder.Date][dicProduct[_CoordResultDate.ProductId]][dicCustomerOrder[_CoordinateDate.CustomerOrderId]].Add(dicSupplierForecast[_CoordinateDate.SupplierOrderId.Value], _CoordinateDate.DateDelier.Value.Date);
            //            }
            //            else
            //            {
            //                dicCoordResult[_CoordResult.DateOrder.Date][dicProduct[_CoordResultDate.ProductId]].Add(dicCustomerOrder[_CoordinateDate.CustomerOrderId], null);
            //            }

            //        }
            //    }
            //}

            //PO = null;
            //FC = null;
            //Mastah = null;
            //#endregion

            #endregion

            FiteMoi(DateFrom, DateTo > DateFrom ? DateTo : DateFrom, false, true, true, false, false, false, false);
        }

        #region PO

        private void kgToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            PrintPO("Horizontal", false);
        }

        private void unitToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            PrintPO("Horizontal", true);
        }

        private void kgToolStripMenuItem2_Click(object sender, EventArgs e)
        {
            PrintPO("Vertical", false);
        }

        private void unitToolStripMenuItem2_Click(object sender, EventArgs e)
        {
            PrintPO("Vertical", false);
        }

        private void compactPOkgToolStripMenuItem_Click(object sender, EventArgs e)
        {
            PrintPO("Compact", false);
        }

        private void compactPOunitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            PrintPO("Compact", true);
        }

        private void btnPrintPOReport_Click(object sender, EventArgs e)
        {
            PrintPO("Report", false);
        }

        #endregion

        #region Cap

        private void groupAllToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FiteMoi(DateFrom, DateFrom > DateTo ? DateFrom : DateTo, true, false, true, true, true);
        }

        private void seperateFarmsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FiteMoi(DateFrom, DateFrom > DateTo ? DateFrom : DateTo, true, false, true, false, true);
        }

        private void seperateThuMuaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FiteMoi(DateFrom, DateFrom > DateTo ? DateFrom : DateTo, true, false, true, true, false);
        }

        private void seperateAllToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FiteMoi(DateFrom, DateFrom > DateTo ? DateFrom : DateTo, true, false, true, false, false);
        }

        private void seperateAllOnlyFarmToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FiteMoi(DateFrom, DateFrom > DateTo ? DateFrom : DateTo, true, false, true, false, false, false, false,
                true);
        }

        #endregion

        #region NoCap

        private void groupAllNoCapToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FiteMoi(DateFrom, DateFrom > DateTo ? DateFrom : DateTo, true, false, false, true, true);
        }

        private void seperateFarmNoCapToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FiteMoi(DateFrom, DateFrom > DateTo ? DateFrom : DateTo, true, false, false, false, true);
        }

        private void seperateThuMuaNoCapToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FiteMoi(DateFrom, DateFrom > DateTo ? DateFrom : DateTo, true, false, false, true, false);
        }

        private void seperateAllNoCapToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FiteMoi(DateFrom, DateFrom > DateTo ? DateFrom : DateTo, true, false, false, false, false);
        }

        private void seperateAllOnlyFarmToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            FiteMoi(DateFrom, DateFrom > DateTo ? DateFrom : DateTo, true, false, false, false, false, false, false,
                true);
        }

        #endregion

        private void testExportLargeExcelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var stopwatch = Stopwatch.StartNew();

            var fileName = string.Format("PO {0}.xlsx",
                DateFrom.ToString("dd.MM") + " - " + DateTo.ToString("dd.MM") + " (" +
                DateTime.Now.ToString("yyyyMMdd HH\\hmm") + ")");
            var path = string.Format(@"D:\Documents\Stuff\VinEco\Mastah Project\Test\" +
                                     fileName);

            //LargeExport(path);

            stopwatch.Stop();
            WriteToRichTextBoxOutput(string.Format("Done in {0}s!", Math.Round(stopwatch.Elapsed.TotalSeconds, 2)));
        }

        private void printToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            FiteMoi(DateFrom, DateFrom > DateTo ? DateFrom : DateTo, false, false, true, false, false, true, false);
        }

        private async void readOpenConfigToolStripMenuItem_Click(object sender, EventArgs e)
        {
            await UpdateOpenConfig();
        }

        private void mongoDbToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var mongoClient = new MongoClient();
            var db = mongoClient.GetDatabase("localtest");
            try
            {
                db.DropCollection("Customer");
                db.DropCollection("Forecast");
                db.DropCollection("Product");
                db.DropCollection("ProductCrossRegion");
                db.DropCollection("ProductRate");
                db.DropCollection("ProductUnit");
                db.DropCollection("PurchaseOrder");
                db.DropCollection("Supplier");
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        #region Picking & Deli stuff

        private async void readDeliToolStripMenuItem_Click(object sender, EventArgs e)
        {
            await UpdateDeli();
        }

        #endregion

        #endregion

        #region Output stuff

        /// <summary>
        ///     Writing Output to Excel. Interop Style.
        ///     <para />
        ///     Old. Classic. Working. Slow.
        /// </summary>
        private void OutputExcel(DataTable dt, string sheetName, Microsoft.Office.Interop.Excel.Workbook xlWb,
            bool YesNoHeader = false, int RowFirst = 6, bool YesNoFirstSheet = false)
        {
            try
            {
                // Open Second Workbook
                //string filePath = string.Format("D:\\Documents\\Stuff\\VinEco\\Mastah Project\\{0}", fileName);
                //var xlWb2 = new Aspose.Cells.Workbook(filePath);
                //var xlWs2 = xlWb2.Worksheets[0];

                var rowTotal = dt.Rows.Count;
                var colTotal = dt.Columns.Count;

                if (rowTotal == 0 || colTotal == 0)
                    return;

                //xlWb2.Worksheets.RemoveAt("PO MB");

                //xlApp.DisplayAlerts = false;
                //foreach (Excel.Worksheet _xlWs in xlWb2.Worksheets)
                //{
                //    WriteToRichTextBoxOutput(_xlWs.Name);
                //    if (_xlWs.Name == sheetName)
                //    {
                //        _xlWs.Delete();
                //    }
                //}
                //xlApp.DisplayAlerts = false;

                //xlWb2.Worksheets.Add(After: xlWb2.Worksheets[xlWb2.Worksheets.Count]);

                //foreach (Excel.Worksheet _xlWs in xlWb.Worksheets)
                //{
                //    WriteToRichTextBoxOutput(_xlWs.Name);
                //}

                Worksheet xlWs = null;
                if (YesNoFirstSheet)
                {
                    xlWs = xlWb.Worksheets[0];
                    xlWs.Name = sheetName;
                }
                else
                {
                    xlWs = xlWb.Worksheets[sheetName]; //xlWb2.Worksheets.Count];
                }
                var rangeToDelete = xlWs.get_Range("A" + RowFirst, (Range)xlWs.Cells[rowTotal, colTotal]);
                rangeToDelete.EntireRow.Delete();

                //int _wsIndex = xlWb2.Worksheets.Add();

                //var xlWs2 = xlWb2.Worksheets[_wsIndex];
                //xlWs2.Name = sheetName;

                //var xlCell2 = xlWs2.Cells;

                #region HeaderStuff

                if (YesNoHeader)
                {
                    var Header = new object[colTotal];

                    // column headings               
                    for (var i = 0; i < colTotal; i++)
                        Header[i] = dt.Columns[i].ColumnName;

                    var HeaderRange = xlWs.get_Range((Range)xlWs.Cells[RowFirst, 1], (Range)xlWs.Cells[1, colTotal]);
                    HeaderRange.Value = Header;
                    HeaderRange.Interior.Color = ColorTranslator.ToOle(Color.LightGray);
                    HeaderRange.Font.Bold = true;
                }

                #endregion

                var _RowFirst = YesNoHeader ? RowFirst + 1 : RowFirst;

                // Limiting the size of object. If this is too large, expect Out of Memory Exception.
                // Interop pls.
                // Apparently larger yields worse performance. Idk why.
                // Too small also negatively impacts performance. Oh god why.

                //var _rowPerBlock = Math.Round(rowTotal / 17, 0);
                //int rowPerBlock = (int)Math.Round(rowTotal / 17d, 0); // 7777;
                var rowPerBlock = 7777;
                //int rowPerBlock = (int)Math.Max(Math.Round(rowTotal / 17d, 0), 7777); // 7777;
                //WriteToRichTextBoxOutput(rowPerBlock);

                var dbCells = new object[rowPerBlock, colTotal];
                var count = 0;
                var rowPos = 0;
                var rowIndex = 0;
                //byte[] cellCheck = new byte[] { 17, 20, 23, 26, 30, 34 };
                foreach (DataRow dr in dt.Rows)
                {
                    // Hardcoding for more efficiency.
                    // Currently this is too slow.
                    for (var colIndex = 0; colIndex < colTotal; colIndex++)
                    {
                        //if (dt.Rows[rowIndex][colIndex] == null) { continue; }

                        var _value = (dr[colIndex] ?? "").ToString();
                        var _type = dt.Columns[colIndex].DataType;
                        if (_value != "" && _value != "0")
                            if (_type == typeof(DateTime))
                                dbCells[rowIndex - rowPos, colIndex] = dr[colIndex];
                            else if (_type == typeof(double))
                                dbCells[rowIndex - rowPos, colIndex] = Convert.ToDouble(_value);
                            else
                                dbCells[rowIndex - rowPos, colIndex] = _value;
                        else
                            continue;
                    }
                    count++;
                    if (count >= rowPerBlock)
                    {
                        xlWs.get_Range((Range)xlWs.Cells[rowPos + _RowFirst, 1],
                            (Range)xlWs.Cells[rowPos + rowPerBlock + _RowFirst - 1, colTotal]).Formula = dbCells;
                        //xlWs2.Range[rowPos + _RowFirst, 1].Resize[rowPos + rowPerBlock + _RowFirst - 1, colTotal].Value = dbCells;
                        dbCells = new object[Math.Min(rowTotal - rowPos, rowPerBlock), colTotal];
                        count = 0;
                        rowPos = rowIndex + 1;
                    }
                    rowIndex++;
                }

                xlWs.get_Range((Range)xlWs.Cells[Math.Max(rowPos + _RowFirst, 2), 1],
                    (Range)xlWs.Cells[rowPos + rowPerBlock + _RowFirst - 1, colTotal]).Formula = dbCells;
                //xlWs2.Range["A" + RowFirst].get_Resize(dbCells.Length[0], dbCells.Length(1)).Value = dbCells;

                dbCells = null;

                if (xlWs != null) Marshal.ReleaseComObject(xlWs);
                xlWs = null;

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
            }
        }

        /// <summary>
        ///     Epplus Approach. Failed horribly.
        /// </summary>
        /// <summary>
        ///     Aspose.Cells Approach. Also failed horribly.
        /// </summary>
        private void OutputExcelAspose(DataTable dataTable, string sheetName, Workbook xlWb, bool YesNoHeader = false,
            int RowFirst = 6, string Position = "A1", Dictionary<string, int> DicColDate = null,
            string CustomDateFormat = "", bool AutoFilter = false)
        {
            try
            {
                var defaultStyle = xlWb.CreateStyle();

                defaultStyle.Font.Name = "Calibri";
                defaultStyle.Font.Size = 11;

                xlWb.DefaultStyle = defaultStyle;

                defaultStyle = null;

                var rowTotal = dataTable.Rows.Count;
                var colTotal = dataTable.Columns.Count;

                //foreach (Aspose.Cells.Worksheet _xlWs in xlWb.Worksheets)
                //{
                //    WriteToRichTextBoxOutput(_xlWs.Name);
                //}

                var xlWs = xlWb.Worksheets[sheetName];

                // Optimize for Performance?
                xlWs.Cells.MemorySetting = MemorySetting.MemoryPreference;

                //xlWs.Cells.DeleteRows(RowFirst - 1, rowTotal);
                //if (DicColDate != null)
                //{

                //    Aspose.Cells.Style style = new Aspose.Cells.Style()
                //    {
                //        Custom = CustomDateFormat == "" ? "dd-MMM" : CustomDateFormat
                //    };

                //    Aspose.Cells.StyleFlag styleFlag = new Aspose.Cells.StyleFlag()
                //    {
                //        NumberFormat = true
                //    };

                //    foreach (int colIndex in DicColDate.Values)
                //    {
                //        xlWs.Cells.Columns[colIndex].ApplyStyle(style, styleFlag);
                //    }
                //}

                xlWs.Cells.ImportDataTable(dataTable, YesNoHeader, RowFirst - 1, 0, rowTotal, colTotal, false,
                    CustomDateFormat == "" ? "dd-MMM" : CustomDateFormat, false);

                if (!xlWs.HasAutofilter)
                    xlWs.AutoFilter.Range =
                        "A1:" + xlWs.Cells[xlWs.Cells.MaxDataRow + 1, xlWs.Cells.MaxDataColumn + 1].Name;

                if (AutoFilter)
                {
                }


                //dataTable = null;

                xlWs = null;

                //GC.Collect();
                //GC.WaitForPendingFinalizers();
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
            }
        }

        /// <summary>
        ///     Exporting to Excel, using OpenXMLWriter.
        ///     <para />
        ///     Super uber fast. Still have no idea how to use this on an already existing Worksheet. Lel.
        /// </summary>
        public static void LargeExport(DataTable dt, string filename, Dictionary<string, int> DicDateCol,
            bool YesNoHeader = false, bool YesNoZero = false, bool YesNoDateColumn = false)
        {
            try
            {
                using (var document = SpreadsheetDocument.Create(filename, SpreadsheetDocumentType.Workbook))
                {
                    var dicType = new Dictionary<Type, CellValues>();

                    var dicColName = new Dictionary<int, string>();

                    for (var colIndex = 0; colIndex < dt.Columns.Count; colIndex++)
                        dicColName.Add(colIndex + 1, GetColumnName(colIndex + 1));

                    dicType.Add(typeof(DateTime), CellValues.Date);
                    dicType.Add(typeof(string), CellValues.InlineString);
                    dicType.Add(typeof(double), CellValues.Number);
                    dicType.Add(typeof(int), CellValues.Number);
                    dicType.Add(typeof(bool), CellValues.Boolean);

                    //this list of attributes will be used when writing a start element
                    List<OpenXmlAttribute> attributes;
                    OpenXmlWriter writer;

                    document.AddWorkbookPart();
                    var workSheetPart = document.WorkbookPart.AddNewPart<WorksheetPart>();

                    // Add Stylesheet.
                    var WorkbookStylesPart = document.WorkbookPart.AddNewPart<WorkbookStylesPart>();
                    WorkbookStylesPart.Stylesheet = AddStyleSheet();
                    WorkbookStylesPart.Stylesheet.Save();

                    writer = OpenXmlWriter.Create(workSheetPart);
                    writer.WriteStartElement(new DocumentFormat.OpenXml.Spreadsheet.Worksheet());
                    writer.WriteStartElement(new SheetData());

                    if (YesNoHeader)
                    {
                        //create a new list of attributes
                        attributes = new List<OpenXmlAttribute>();
                        // add the row index attribute to the list
                        attributes.Add(new OpenXmlAttribute("r", null, 1.ToString()));

                        //write the row start element with the row index attribute
                        writer.WriteStartElement(new Row(), attributes);

                        for (var columnNum = 1; columnNum <= dt.Columns.Count; ++columnNum)
                        {
                            var type = dt.Columns[columnNum - 1].DataType;
                            //reset the list of attributes
                            //attributes = new List<OpenXmlAttribute>();
                            // add data type attribute - in this case inline string (you might want to look at the shared strings table)
                            //attributes.Add(new OpenXmlAttribute("t", null, "str")); //(type == typeof(string) ? "str" : (YesNoDateColumn == false ? "str" : dicType[typeof(DateTime)].ToString()))));

                            //add the cell reference attribute
                            //attributes.Add(new OpenXmlAttribute("r", "", string.Format("{0}{1}", dicColName[columnNum], 1)));

                            //write the cell start element with the type and reference attributes
                            //writer.WriteStartElement(new DocumentFormat.OpenXml.Spreadsheet.Cell(), attributes);

                            DateTime _value;
                            var _dateValue = 0;
                            if (DateTime.TryParse(dt.Columns[columnNum - 1].ColumnName, out _value))
                                _dateValue = (int)(_value.Date - new DateTime(1900, 1, 1)).TotalDays + 2;

                            //write the cell value
                            var cell = new Cell
                            {
                                DataType = type == typeof(double) && _dateValue != 0
                                    ? CellValues.Number
                                    : CellValues.String,
                                CellReference = string.Format("{0}{1}", dicColName[columnNum], 1),
                                CellValue = new CellValue(type == typeof(double) && _dateValue != 0
                                    ? _dateValue.ToString()
                                    : dt.Columns[columnNum - 1].ColumnName),
                                StyleIndex = (uint)(type == typeof(double) && _dateValue != 0 ? 1 : 0)
                            };
                            writer.WriteElement(cell);

                            //writer.WriteElement(new CellValue((type == typeof(double) && YesNoDateColumn == true) ? _dateValue.ToString() : dt.Columns[columnNum - 1].ColumnName));

                            //writer.WriteElement(new CellValue(string.Format("This is Row {0}, Cell {1}", rowNum, columnNum)));

                            // write the end cell element
                            //writer.WriteEndElement();
                        }

                        // write the end row element
                        writer.WriteEndElement();
                    }

                    for (var rowNum = 1; rowNum <= dt.Rows.Count; rowNum++)
                    {
                        //create a new list of attributes
                        attributes = new List<OpenXmlAttribute>();
                        // add the row index attribute to the list
                        attributes.Add(new OpenXmlAttribute("r", null, (YesNoHeader ? rowNum + 1 : rowNum).ToString()));

                        //write the row start element with the row index attribute
                        writer.WriteStartElement(new Row(), attributes);

                        var dr = dt.Rows[rowNum - 1];
                        for (var columnNum = 1; columnNum <= dt.Columns.Count; columnNum++)
                        {
                            var colName = dt.Columns[columnNum - 1].ColumnName;
                            var type = dt.Columns[columnNum - 1].DataType;
                            //reset the list of attributes
                            //attributes = new List<OpenXmlAttribute>();
                            // add data type attribute - in this case inline string (you might want to look at the shared strings table)
                            //attributes.Add(new OpenXmlAttribute("t", null, type == typeof(string) ? "str" : dicType[type].ToString()));
                            //add the cell reference attribute
                            //attributes.Add(new OpenXmlAttribute("r", "", string.Format("{0}{1}", dicColName[columnNum], (YesNoHeader ? rowNum + 1 : rowNum))));

                            //write the cell start element with the type and reference attributes
                            //writer.WriteStartElement(new DocumentFormat.OpenXml.Spreadsheet.Cell(), attributes);

                            ////write the cell value
                            //if (YesNoZero | dr[columnNum - 1].ToString() != "0")
                            //{
                            //    writer.WriteElement(new CellValue(dr[columnNum - 1].ToString()));
                            //}
                            //{
                            //    // In case of 0. Can safely forsake this part.
                            //    //writer.WriteElement(new CellValue(""));
                            //}

                            writer.WriteElement(new Cell
                            {
                                DataType = type == typeof(string)
                                    ? CellValues.String
                                    : (YesNoDateColumn && type == typeof(DateTime) ? CellValues.Number : dicType[type]),
                                CellReference = string.Format("{0}{1}", dicColName[columnNum],
                                    YesNoHeader ? rowNum + 1 : rowNum),
                                CellValue = new CellValue(dr[columnNum - 1].ToString()),
                                StyleIndex = (uint)(DicDateCol.ContainsKey(colName) ? 1 : 0)
                            });

                            //writer.WriteElement(new CellValue(string.Format("This is Row {0}, Cell {1}", rowNum, columnNum)));

                            // write the end cell element
                            //writer.WriteEndElement();
                        }

                        // write the end row element
                        writer.WriteEndElement();
                    }

                    // write the end SheetData element
                    writer.WriteEndElement();
                    // write the end Worksheet element
                    writer.WriteEndElement();
                    writer.Close();

                    writer = OpenXmlWriter.Create(document.WorkbookPart);
                    writer.WriteStartElement(new DocumentFormat.OpenXml.Spreadsheet.Workbook());
                    writer.WriteStartElement(new Sheets());

                    writer.WriteElement(new Sheet
                    {
                        Name = dt.TableName == "" ? "Whatever" : dt.TableName,
                        SheetId = 1,
                        Id = document.WorkbookPart.GetIdOfPart(workSheetPart)
                    });

                    // End Sheets
                    writer.WriteEndElement();
                    // End Workbook
                    writer.WriteEndElement();

                    writer.Close();

                    document.Close();
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private static Stylesheet AddStyleSheet()
        {
            try
            {
                var workbookstylesheet = new Stylesheet();

                var font0 = new Font(); // Default font

                var font1 = new Font(); // Bold font
                var bold = new Bold();
                font1.Append(bold);

                var fonts = new Fonts(); // <APENDING Fonts>
                fonts.Append(font0);
                fonts.Append(font1);

                // <Fills>
                var fill0 = new Fill(); // Default fill

                var fills = new Fills(); // <APENDING Fills>
                fills.Append(fill0);

                // <Borders>
                var border0 = new Border(); // Defualt border

                var borders = new Borders(); // <APENDING Borders>
                borders.Append(border0);

                var nf2DateTime = new NumberingFormat
                {
                    NumberFormatId = UInt32Value.FromUInt32(7170),
                    FormatCode = StringValue.FromString("dd-MMM")
                };
                workbookstylesheet.NumberingFormats = new NumberingFormats();
                workbookstylesheet.NumberingFormats.Append(nf2DateTime);

                // <CellFormats>
                var cellformat0 = new CellFormat
                {
                    FontId = 0,
                    FillId = 0,
                    BorderId = 0
                }; // Default style : Mandatory | Style ID =0

                var cellformat1 = new CellFormat
                {
                    BorderId = 0,
                    FillId = 0,
                    FontId = 0,
                    NumberFormatId = 7170,
                    FormatId = 0,
                    ApplyNumberFormat = true
                };

                var cellformat2 = new CellFormat
                {
                    BorderId = 0,
                    FillId = 0,
                    FontId = 0,
                    NumberFormatId = 14,
                    FormatId = 0,
                    ApplyNumberFormat = true
                };

                // <APENDING CellFormats>
                var cellformats = new CellFormats();
                cellformats.Append(cellformat0);
                cellformats.Append(cellformat1);
                cellformats.Append(cellformat2);


                // Append FONTS, FILLS , BORDERS & CellFormats to stylesheet <Preserve the ORDER>
                workbookstylesheet.Append(fonts);
                workbookstylesheet.Append(fills);
                workbookstylesheet.Append(borders);
                workbookstylesheet.Append(cellformats);

                //// Finalize
                //stylesheet.Stylesheet = workbookstylesheet;
                //stylesheet.Stylesheet.Save();

                return workbookstylesheet;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        ///     OpenWriter Style, for Multiple DataTable into Multiple Worksheets in a single Workbook. A real fucking pain.
        /// </summary>
        public static void LargeExportOneWorkbook(string filePath, List<DataTable> listDt, bool YesNoHeader = false,
            bool YesNoZero = false)
        {
            try
            {
                using (var document = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook))
                {
                    document.AddWorkbookPart();

                    OpenXmlWriter writer;

                    OpenXmlWriter writerXb;

                    writerXb = OpenXmlWriter.Create(document.WorkbookPart);
                    writerXb.WriteStartElement(new DocumentFormat.OpenXml.Spreadsheet.Workbook());
                    writerXb.WriteStartElement(new Sheets());

                    var count = 0;

                    foreach (var dt in listDt)
                    {
                        var dicType = new Dictionary<Type, CellValues>();

                        var dicColName = new Dictionary<int, string>();

                        for (var colIndex = 0; colIndex < dt.Columns.Count; colIndex++)
                            dicColName.Add(colIndex + 1, GetColumnName(colIndex + 1));

                        dicType.Add(typeof(DateTime), CellValues.Date);
                        dicType.Add(typeof(string), CellValues.InlineString);
                        dicType.Add(typeof(double), CellValues.Number);
                        dicType.Add(typeof(int), CellValues.Number);

                        //this list of attributes will be used when writing a start element
                        List<OpenXmlAttribute> attributes;

                        var workSheetPart = document.WorkbookPart.AddNewPart<WorksheetPart>();

                        writer = OpenXmlWriter.Create(workSheetPart);
                        writer.WriteStartElement(new DocumentFormat.OpenXml.Spreadsheet.Worksheet());
                        writer.WriteStartElement(new SheetData());

                        if (YesNoHeader)
                        {
                            //create a new list of attributes
                            attributes = new List<OpenXmlAttribute>();
                            // add the row index attribute to the list
                            attributes.Add(new OpenXmlAttribute("r", null, 1.ToString()));

                            //write the row start element with the row index attribute
                            writer.WriteStartElement(new Row(), attributes);

                            for (var columnNum = 1; columnNum <= dt.Columns.Count; ++columnNum)
                            {
                                var type = dt.Columns[columnNum - 1].DataType;
                                //reset the list of attributes
                                attributes = new List<OpenXmlAttribute>();
                                // add data type attribute - in this case inline string (you might want to look at the shared strings table)
                                attributes.Add(new OpenXmlAttribute("t", null,
                                    "str")); // type == typeof(string) ? "str" : dicType[type].ToString()));
                                //add the cell reference attribute
                                attributes.Add(new OpenXmlAttribute("r", "",
                                    string.Format("{0}{1}", dicColName[columnNum], 1)));

                                //write the cell start element with the type and reference attributes
                                writer.WriteStartElement(new Cell(), attributes);

                                //write the cell value
                                writer.WriteElement(new CellValue(dt.Columns[columnNum - 1].ColumnName));

                                //writer.WriteElement(new CellValue(string.Format("This is Row {0}, Cell {1}", rowNum, columnNum)));

                                // write the end cell element
                                writer.WriteEndElement();
                            }

                            // write the end row element
                            writer.WriteEndElement();
                        }

                        for (var rowNum = 1; rowNum <= dt.Rows.Count; rowNum++)
                        {
                            //create a new list of attributes
                            attributes = new List<OpenXmlAttribute>();
                            // add the row index attribute to the list
                            attributes.Add(new OpenXmlAttribute("r", null,
                                (YesNoHeader ? rowNum + 1 : rowNum).ToString()));

                            //write the row start element with the row index attribute
                            writer.WriteStartElement(new Row(), attributes);

                            var dr = dt.Rows[rowNum - 1];
                            for (var columnNum = 1; columnNum <= dt.Columns.Count; columnNum++)
                            {
                                var type = dt.Columns[columnNum - 1].DataType;
                                //reset the list of attributes
                                attributes = new List<OpenXmlAttribute>();
                                // add data type attribute - in this case inline string (you might want to look at the shared strings table)
                                attributes.Add(new OpenXmlAttribute("t", null,
                                    type == typeof(string) ? "str" : dicType[type].ToString()));
                                //add the cell reference attribute
                                attributes.Add(new OpenXmlAttribute("r", "",
                                    string.Format("{0}{1}", dicColName[columnNum], YesNoHeader ? rowNum + 1 : rowNum)));

                                //write the cell start element with the type and reference attributes
                                writer.WriteStartElement(new Cell(), attributes);

                                //write the cell value
                                if (YesNoZero | (dr[columnNum - 1].ToString() != "0"))
                                    writer.WriteElement(new CellValue(dr[columnNum - 1].ToString()));
                                {
                                    // In case of 0. Can safely forsake this part.
                                    //writer.WriteElement(new CellValue(""));
                                }

                                //writer.WriteElement(new CellValue(string.Format("This is Row {0}, Cell {1}", rowNum, columnNum)));

                                // write the end cell element
                                writer.WriteEndElement();
                            }

                            // write the end row element
                            writer.WriteEndElement();
                        }

                        // write the end SheetData element
                        writer.WriteEndElement();
                        // write the end Worksheet element
                        writer.WriteEndElement();
                        writer.Close();

                        writerXb.WriteElement(new Sheet
                        {
                            Name = dt.TableName,
                            SheetId = Convert.ToUInt32(count + 1),
                            Id = document.WorkbookPart.GetIdOfPart(workSheetPart)
                        });

                        count++;
                    }
                    // End Sheets
                    writerXb.WriteEndElement();
                    // End Workbook
                    writerXb.WriteEndElement();

                    writerXb.Close();

                    document.Close();
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public static void LargeExportOriginal(string filename)
        {
            using (var document = SpreadsheetDocument.Create(filename, SpreadsheetDocumentType.Workbook))
            {
                //this list of attributes will be used when writing a start element
                List<OpenXmlAttribute> attributes;
                OpenXmlWriter writer;

                document.AddWorkbookPart();
                var workSheetPart = document.WorkbookPart.AddNewPart<WorksheetPart>();

                writer = OpenXmlWriter.Create(workSheetPart);
                writer.WriteStartElement(new DocumentFormat.OpenXml.Spreadsheet.Worksheet());
                writer.WriteStartElement(new SheetData());

                for (var rowNum = 1; rowNum <= 115000; ++rowNum)
                {
                    //create a new list of attributes
                    attributes = new List<OpenXmlAttribute>();
                    // add the row index attribute to the list
                    attributes.Add(new OpenXmlAttribute("r", null, rowNum.ToString()));

                    //write the row start element with the row index attribute
                    writer.WriteStartElement(new Row(), attributes);

                    for (var columnNum = 1; columnNum <= 30; ++columnNum)
                    {
                        //reset the list of attributes
                        attributes = new List<OpenXmlAttribute>();
                        // add data type attribute - in this case inline string (you might want to look at the shared strings table)
                        attributes.Add(new OpenXmlAttribute("t", null, "str"));
                        //add the cell reference attribute
                        attributes.Add(new OpenXmlAttribute("r", "",
                            string.Format("{0}{1}", GetColumnName(columnNum), rowNum)));

                        //write the cell start element with the type and reference attributes
                        writer.WriteStartElement(new Cell(), attributes);
                        //write the cell value
                        writer.WriteElement(
                            new CellValue(string.Format("This is Row {0}, Cell {1}", rowNum, columnNum)));

                        // write the end cell element
                        writer.WriteEndElement();
                    }

                    // write the end row element
                    writer.WriteEndElement();
                }

                // write the end SheetData element
                writer.WriteEndElement();
                // write the end Worksheet element
                writer.WriteEndElement();
                writer.Close();

                writer = OpenXmlWriter.Create(document.WorkbookPart);
                writer.WriteStartElement(new DocumentFormat.OpenXml.Spreadsheet.Workbook());
                writer.WriteStartElement(new Sheets());

                writer.WriteElement(new Sheet
                {
                    Name = "Large Sheet",
                    SheetId = 1,
                    Id = document.WorkbookPart.GetIdOfPart(workSheetPart)
                });

                // End Sheets
                writer.WriteEndElement();
                // End Workbook
                writer.WriteEndElement();

                writer.Close();

                document.Close();
            }
        }

        /// <summary>
        ///     A simple helper to get the column name from the column index. This is not well tested!
        ///     <para />
        ///     Worked anyway. For a Dictionary anyway.
        /// </summary>
        private static string GetColumnName(int columnIndex)
        {
            var dividend = columnIndex;
            var columnName = string.Empty;
            int modifier;

            while (dividend > 0)
            {
                modifier = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modifier) + columnName;
                dividend = (dividend - modifier) / 26;
            }

            return columnName;
        }

        /// <summary>
        ///     Convert from one file format to another, using Interop.
        ///     Because apparently OpenXML doesn't deal with .xls type ( Including, but not exclusive to .xlsb )
        /// </summary>
        private void ConvertToXlsbInterop(string filePath, string PreviousExtension = "",
            string AfterwardExtension = "", bool YesNoDeleteFile = false)
        {
            try
            {
                // Remember the list of running Excel.Application.
                // Before initialize xlApp.
                var processBefore = Process.GetProcessesByName("excel");

                // Initialize new instance of Interop Excel.Application.
                var xlApp = new Microsoft.Office.Interop.Excel.Application();

                // Remember the list of running Excel.Application.
                // After initialize xlApp.
                var processAfter = Process.GetProcessesByName("excel");

                var processID = 0;

                // Compare two lists, get the first and the only process that's not in the 'Before' List.
                foreach (var process in processAfter)
                    if (!processBefore.Select(p => p.Id).Contains(process.Id))
                    {
                        processID = process.Id;
                        break;
                    }

                xlApp.ScreenUpdating = false;
                xlApp.EnableEvents = false;
                xlApp.DisplayAlerts = false;
                xlApp.DisplayStatusBar = false;
                xlApp.AskToUpdateLinks = false;

                var xlWb = xlApp.Workbooks.Open(filePath);

                var missing = Type.Missing;
                xlWb.SaveAs(filePath.Replace(PreviousExtension, AfterwardExtension), XlFileFormat.xlExcel12, missing,
                    missing, false, false, XlSaveAsAccessMode.xlExclusive, missing, missing, missing);

                xlWb.Close(false);
                Marshal.ReleaseComObject(xlWb);
                xlWb = null;

                xlApp.Quit();
                Marshal.ReleaseComObject(xlApp);
                xlApp = null;

                if (YesNoDeleteFile)
                    File.Delete(filePath);

                // Kill the instance of Interop Excel.Application used by this call.
                if (processID != 0)
                {
                    var process = Process.GetProcessById(processID);
                    process.Kill();
                }

                WriteToRichTextBoxOutput(Name + " finished peacefully! ");
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        #region Delete the thing

        //public void Delete_Evaluation_Sheet(string fileName)
        //{
        //    try
        //    {
        //        string Sheetid = "";
        //        using (SpreadsheetDocument document = SpreadsheetDocument.Open(fileName, true))
        //        {
        //            WorkbookPart wbPart = document.WorkbookPart;
        //            Sheet theSheet = wbPart.Workbook.Descendants<Sheet>().Where(s => s.Name == "Evaluation Warning").FirstOrDefault();
        //            if (theSheet == null)
        //            {
        //                // The specified sheet doesn't exist.
        //            }
        //            //Store the SheetID for the reference
        //            Sheetid = theSheet.SheetId;

        //            // Remove the sheet reference from the workbook.
        //            WorksheetPart worksheetPart = (WorksheetPart)(wbPart.GetPartById(theSheet.Id));
        //            theSheet.Remove();

        //            // Delete the worksheet part.
        //            wbPart.DeletePart(worksheetPart);

        //            // Save the workbook.
        //            wbPart.Workbook.Save();
        //        }
        //    }
        //    catch (Exception ex)
        //    {

        //        throw ex;
        //    }
        //}

        public void Delete_Evaluation_Sheet_Interop(string filePath)
        {
            try
            {
                // Remember the list of running Excel.Application.
                // Before initialize xlApp.
                var processBefore = Process.GetProcessesByName("excel");

                // Initialize new instance of Interop Excel.Application.
                var xlApp = new Microsoft.Office.Interop.Excel.Application();

                // Remember the list of running Excel.Application.
                // After initialize xlApp.
                var processAfter = Process.GetProcessesByName("excel");

                var processID = 0;

                // Compare two lists, get the first and the only process that's not in the 'Before' List.
                foreach (var process in processAfter)
                    if (!processBefore.Select(p => p.Id).Contains(process.Id))
                    {
                        processID = process.Id;
                        break;
                    }

                var xlWb = xlApp.Workbooks.Open(
                    filePath,
                    false,
                    false,
                    5,
                    "",
                    "",
                    true,
                    XlPlatform.xlWindows,
                    "",
                    true,
                    false,
                    0,
                    true,
                    false,
                    false);

                xlApp.ScreenUpdating = false;
                xlApp.Calculation = XlCalculation.xlCalculationManual;
                xlApp.EnableEvents = false;
                xlApp.DisplayAlerts = false;
                xlApp.DisplayStatusBar = false;
                xlApp.AskToUpdateLinks = false;

                foreach (Worksheet _ws in xlWb.Worksheets)
                    if (_ws.Name == "Evaluation Warning") _ws.Delete();
                xlWb.Sheets[1].Activate();

                xlApp.ScreenUpdating = true;
                xlApp.Calculation = XlCalculation.xlCalculationAutomatic;
                xlApp.EnableEvents = true;
                xlApp.DisplayAlerts = false;
                xlApp.DisplayStatusBar = true;
                xlApp.AskToUpdateLinks = true;

                xlWb.Close(true);

                if (xlWb != null) Marshal.ReleaseComObject(xlWb);
                xlWb = null;

                xlApp.Quit();
                if (xlApp != null) Marshal.ReleaseComObject(xlApp);
                xlApp = null;

                //GC.Collect();
                //GC.WaitForPendingFinalizers();

                // Kill the instance of Interop Excel.Application used by this call.
                if (processID != 0)
                {
                    var process = Process.GetProcessById(processID);
                    process.Kill();
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        #endregion

        #endregion

        #region Customized Functions.

        // Convert non-ASCII characters in Vietnamese to unsigned, ASCII equivalents.
        public static string ConvertToUnsigned(string text)
        {
            var ExcludedChars = "(-)"; // lol.

            for (var i = 33; i < 48; i++)
                if (!ExcludedChars.Contains(((char)i).ToString()))
                    text = text.Replace(((char)i).ToString(), "");

            for (var i = 58; i < 65; i++)
                text = text.Replace(((char)i).ToString(), "");

            for (var i = 91; i < 97; i++)
                text = text.Replace(((char)i).ToString(), "");

            for (var i = 123; i < 127; i++)
                text = text.Replace(((char)i).ToString(), "");
            //text = text.Replace(" ", "-"); //Comment lại để không đưa khoảng trắng thành ký tự -
            var regex = new Regex(@"\p{IsCombiningDiacriticalMarks}+");

            var strFormD = text.Normalize(NormalizationForm.FormD);

            return regex.Replace(strFormD, string.Empty).Replace('\u0111', 'd').Replace('\u0110', 'D');
        }

        private static readonly string[] VietNamChar =
        {
            "aAeEoOuUiIdDyY",
            "áàạảãâấầậẩẫăắằặẳẵ",
            "ÁÀẠẢÃÂẤẦẬẨẪĂẮẰẶẲẴ",
            "éèẹẻẽêếềệểễ",
            "ÉÈẸẺẼÊẾỀỆỂỄ",
            "óòọỏõôốồộổỗơớờợởỡ",
            "ÓÒỌỎÕÔỐỒỘỔỖƠỚỜỢỞỠ",
            "úùụủũưứừựửữ",
            "ÚÙỤỦŨƯỨỪỰỬỮ",
            "íìịỉĩ",
            "ÍÌỊỈĨ",
            "đ",
            "Đ",
            "ýỳỵỷỹ",
            "ÝỲỴỶỸ"
        };

        #endregion
    }
}