using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
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
using Application = Microsoft.Office.Interop.Excel.Application;
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
using Style = Aspose.Cells.Style;
using Workbook = Aspose.Cells.Workbook;
using Worksheet = Microsoft.Office.Interop.Excel.Worksheet;
using MongoDB.Driver.GridFS;

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
        //const directoryPath = System.IO.Path.GetDirectoryName(Application.ExecutablePath);
        
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

                Stopwatch stopwatch = Stopwatch.StartNew();

                var            mongoClient = new MongoClient();
                IMongoDatabase db          = mongoClient.GetDatabase("localtest");

                IOrderedEnumerable<PurchaseOrderDate> PO = db.GetCollection<PurchaseOrderDate>("PurchaseOrder")
                                                             .Find(
                                                                  x =>
                                                                      x.DateOrder >= DateFrom.Date &&
                                                                      x.DateOrder <= DateTo.Date)
                                                             .ToList()
                                                             .OrderBy(x => x.DateOrder);

                List<Product>                   Product        = db.GetCollection<Product>("Product").AsQueryable().ToList();
                List<Customer>                  Customer       = db.GetCollection<Customer>("Customer").AsQueryable().ToList();
                Dictionary<string, ProductUnit> dicProductUnit = db.GetCollection<ProductUnit>("ProductUnit")
                                                                   .AsQueryable()
                                                                   .ToDictionary(x => x.ProductCode);

                var dicProduct  = new Dictionary<Guid, Product>();
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

                foreach (Product _Product in Product)
                    //if (dicClass.TryGetValue(_Product.ProductCode.Substring(0, 1), out string _ProductClass))
                    //{
                    //    _Product.ProductClassification = _ProductClass;
                    //}
                {
                    dicProduct.Add(_Product.ProductId, _Product);
                }

                foreach (Customer _Customer in Customer)
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

                    foreach (PurchaseOrderDate PODate in PO)
                    {
                        dtPO.Columns.Add(PODate.DateOrder.Date.ToString("MM/dd/yyyy"), typeof(double)).DefaultValue = 0;
                    }

                    var dicRow = new Dictionary<string, int>();
                    foreach (PurchaseOrderDate PODate in PO)
                    {
                        foreach (ProductOrder _ProductOrder in PODate.ListProductOrder)
                        {
                            foreach (CustomerOrder _CustomerOrder in _ProductOrder.ListCustomerOrder)
                            {
                                if (!YesNoByUnit)
                                {
                                    string _OrderUnitType = ProperUnit(_CustomerOrder.Unit.ToLower(), dicUnit);
                                    _CustomerOrder.Unit   = _OrderUnitType;
                                    if (dicCustomer[_CustomerOrder.CustomerId].CustomerType == "VM+" && _OrderUnitType != "Kg")
                                    {
                                        string _ProductCode = dicProduct[_ProductOrder.ProductId].ProductCode;
                                        if (dicProductUnit.TryGetValue(_ProductCode, out ProductUnit _ProductUnit))
                                        {
                                            ProductUnitRegion _ProductUnitRegion =
                                                _ProductUnit.ListRegion.FirstOrDefault(x => x.OrderUnitType == _OrderUnitType);
                                            if (_ProductUnitRegion                                          != null)
                                            {
                                                _CustomerOrder.Unit            = _OrderUnitType;
                                                _CustomerOrder.QuantityOrderKg =
                                                    _CustomerOrder.QuantityOrder * _ProductUnitRegion.OrderUnitPer;
                                            }
                                        }
                                        else
                                        {
                                            _CustomerOrder.Unit            = "Kg";
                                            _CustomerOrder.QuantityOrderKg = _CustomerOrder.QuantityOrder;
                                        }
                                    }
                                    else
                                    {
                                        _CustomerOrder.Unit            = "Kg";
                                        _CustomerOrder.QuantityOrderKg = _CustomerOrder.QuantityOrder;
                                    }
                                }

                                Product  _Product  = dicProduct[_ProductOrder.ProductId];
                                Customer _Customer = dicCustomer[_CustomerOrder.CustomerId];

                                DataRow dr = null;

                                string sKey = _Product.ProductCode + _Customer.CustomerCode;
                                if (!dicRow.TryGetValue(sKey, out int _rowIndex))
                                {
                                    dr = dtPO.NewRow();
                                    dicRow.Add(sKey, dtPO.Rows.Count);

                                    dr["VE Code"]   = _Product.ProductCode;
                                    dr["VE Name"]   = _Product.ProductName;
                                    dr["Class"]     = _Product.ProductClassification;
                                    dr["StoreCode"] = _Customer.CustomerCode;
                                    dr["StoreName"] = _Customer.CustomerName;
                                    dr["StoreType"] = _Customer.CustomerType;
                                    dr["SubRegion"] = _Customer.CustomerRegion;
                                    dr["Region"]    = _Customer.CustomerBigRegion;
                                    dr["P&L"]       = _Customer.Company;
                                    dr["Note"]      =
                                        _Product.ProductNote.Contains(_Customer.CustomerBigRegion == "Miền Nam"
                                                                          ? "South"
                                                                          : "North")
                                            ? "Ok"
                                            : "Out of List";
                                    if (YesNoByUnit)
                                    {
                                        dr["Unit"]                                       = _CustomerOrder.Unit;
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
                                    {
                                        dr[PODate.DateOrder.Date.ToString("MM/dd/yyyy")] =
                                            (double) dr[PODate.DateOrder.Date.ToString("MM/dd/yyyy")] +
                                            _CustomerOrder.QuantityOrder;
                                    }
                                    else
                                    {
                                        dr[PODate.DateOrder.Date.ToString("MM/dd/yyyy")] =
                                            (double) dr[PODate.DateOrder.Date.ToString("MM/dd/yyyy")] +
                                            _CustomerOrder.QuantityOrderKg;
                                    }
                                }
                            }
                        }
                    }
                }
                // Vertical PO - making it pivot-able ( No pun intended. )
                else if (Choice == "Vertical")
                {
                    dtPO.TableName += " Vertical";

                    dtPO.Columns.Add("PCODE", typeof(string)).DefaultValue         = "";
                    dtPO.Columns.Add("PNAME", typeof(string)).DefaultValue         = "";
                    dtPO.Columns.Add("PCLASS", typeof(string)).DefaultValue        = "";
                    dtPO.Columns.Add("NOTE", typeof(string)).DefaultValue          = "";
                    dtPO.Columns.Add("Climate", typeof(string)).DefaultValue       = "";
                    dtPO.Columns.Add("CCODE", typeof(string)).DefaultValue         = "";
                    dtPO.Columns.Add("CNAME", typeof(string)).DefaultValue         = "";
                    dtPO.Columns.Add("CTYPE", typeof(string)).DefaultValue         = "";
                    dtPO.Columns.Add("CREGION", typeof(string)).DefaultValue       = "";
                    dtPO.Columns.Add("REGION", typeof(string)).DefaultValue        = "";
                    dtPO.Columns.Add("P&L", typeof(string)).DefaultValue           = "";
                    dtPO.Columns.Add("DateOrder", typeof(int)).DefaultValue        = 0;
                    dtPO.Columns.Add("QuantityOrder", typeof(double)).DefaultValue = 0;
                    dtPO.Columns.Add("DateReceive", typeof(int)).DefaultValue      = 0;

                    DicColDate.Add("DateOrder", dtPO.Columns.IndexOf("DateOrder"));
                    DicColDate.Add("DateReceive", dtPO.Columns.IndexOf("DateReceive"));

                    var dicRow = new Dictionary<string, int>();
                    foreach (PurchaseOrderDate PODate in PO)
                    {
                        foreach (ProductOrder _ProductOrder in PODate.ListProductOrder)
                        {
                            foreach (CustomerOrder _CustomerOrder in _ProductOrder.ListCustomerOrder)
                            {
                                Product  _Product  = dicProduct[_ProductOrder.ProductId];
                                Customer _Customer = dicCustomer[_CustomerOrder.CustomerId];

                                DataRow dr = null;

                                string sKey =
                                    $"{_Product.ProductCode}{_Customer.CustomerCode}{_Customer.Company}{PODate.DateOrder.Date:yyyyMMdd}";

                                if (!dicRow.TryGetValue(sKey, out int _rowIndex))
                                {
                                    dr = dtPO.NewRow();
                                    dicRow.Add(sKey, dtPO.Rows.Count);

                                    dr["PCODE"]   = _Product.ProductCode;
                                    dr["PNAME"]   = _Product.ProductName;
                                    dr["PCLASS"]  = _Product.ProductClassification;
                                    dr["CCODE"]   = _Customer.CustomerCode;
                                    dr["CNAME"]   = _Customer.CustomerName;
                                    dr["CTYPE"]   = _Customer.CustomerType;
                                    dr["CREGION"] = _Customer.CustomerRegion;
                                    dr["REGION"]  = _Customer.CustomerBigRegion;
                                    dr["P&L"]     = _Customer.Company;
                                    dr["Note"]    =
                                        _Product.ProductNote.Contains(_Customer.CustomerBigRegion == "Miền Nam"
                                                                          ? "South"
                                                                          : "North")
                                            ? "Ok"
                                            : "Out of List";
                                    dr["Climate"]       = _Product.ProductClimate;
                                    dr["DateOrder"]     = (int) (PODate.DateOrder.Date - new DateTime(1900, 1, 1)).TotalDays + 2;
                                    dr["QuantityOrder"] = _CustomerOrder.QuantityOrder;
                                    dr["DateReceive"]   = (string) dr["REGION"] == "Miền Nam"
                                                              ? (int) dr["DateOrder"] + 1
                                                              : (int) dr["DateOrder"];

                                    dtPO.Rows.Add(dr);
                                }
                                else
                                {
                                    dr                  = dtPO.Rows[_rowIndex];
                                    dr["QuantityOrder"] = (double) dr["QuantityOrder"] + _CustomerOrder.QuantityOrder;
                                }
                            }
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
                    foreach (PurchaseOrderDate PODate in PO)
                    {
                        foreach (ProductOrder _ProductOrder in PODate.ListProductOrder)
                        {
                            foreach (CustomerOrder _CustomerOrder in _ProductOrder.ListCustomerOrder)
                            {
                                Product  _Product  = dicProduct[_ProductOrder.ProductId];
                                Customer _Customer = dicCustomer[_CustomerOrder.CustomerId];

                                DataRow dr = null;

                                string sKey =
                                    $"{_Product.ProductCode}{_Customer.CustomerBigRegion}{_Customer.CustomerRegion}{_Customer.CustomerType}{_Customer.Company}{PODate.DateOrder.Date:yyyyMMdd}";

                                if (!dicRow.TryGetValue(sKey, out int _rowIndex))
                                {
                                    dr = dtPO.NewRow();
                                    dicRow.Add(sKey, dtPO.Rows.Count);

                                    dr["PCODE"]              = _Product.ProductCode;
                                    dr["PNAME"]              = _Product.ProductName;
                                    dr["PCLASS"]             = _Product.ProductClassification;
                                    dr["ProductOrientation"] = _Product.ProductOrientation;
                                    dr["ProductClimate"]     = _Product.ProductClimate;
                                    dr["ProductionGroup"]    = _Product.ProductionGroup;
                                    //dr["CustomerCode"] = _Customer.CustomerCode;
                                    //dr["CustomerName"] = _Customer.CustomerName;
                                    dr["Note"] =
                                        _Product.ProductNote.Contains(_Customer.CustomerBigRegion == "Miền Nam"
                                                                          ? "South"
                                                                          : "North")
                                            ? "Ok"
                                            : "Out of List";
                                    dr["CTYPE"]   = _Customer.CustomerType;
                                    dr["CREGION"] = _Customer.CustomerRegion;
                                    dr["P&L"]     = _Customer.Company;
                                    dr["REGION"]  = _Customer.CustomerBigRegion;
                                    //dr["DateOrder"] = (int)(PODate.DateOrder.Date - new DateTime(1900, 1, 1)).TotalDays + 2;
                                    dr["DateOrder"]     = PODate.DateOrder.Date;
                                    dr["QuantityOrder"] = _CustomerOrder.QuantityOrder;

                                    dtPO.Rows.Add(dr);
                                }
                                else
                                {
                                    dr                  = dtPO.Rows[_rowIndex];
                                    dr["QuantityOrder"] = (double) dr["QuantityOrder"] + _CustomerOrder.QuantityOrder;
                                }
                            }
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
                    dtPO.Columns.Add("Source", typeof(string)).DefaultValue        = "PO";
                    dtPO.Columns.Add("DateOrder", typeof(DateTime));
                    dtPO.Columns.Add("P&L", typeof(string));
                    dtPO.Columns.Add("Note", typeof(string));

                    DicColDate.Add("DateOrder", dtPO.Columns.IndexOf("DateOrder"));
                    DicColDate.Add("DateReceive", dtPO.Columns.IndexOf("DateReceive"));

                    var dicRow = new Dictionary<string, int>();
                    foreach (PurchaseOrderDate PODate in PO)
                    {
                        foreach (ProductOrder _ProductOrder in PODate.ListProductOrder)
                        {
                            foreach (CustomerOrder _CustomerOrder in _ProductOrder.ListCustomerOrder)
                            {
                                Product  _Product  = dicProduct[_ProductOrder.ProductId];
                                Customer _Customer = dicCustomer[_CustomerOrder.CustomerId];

                                DataRow dr = null;

                                string sKey =
                                    $"{_Product.ProductCode}{_Customer.CustomerType}{_Customer.CustomerBigRegion}{_Customer.Company}{PODate.DateOrder.Date:yyyyMMdd}";

                                if (!dicRow.TryGetValue(sKey, out int _rowIndex))
                                {
                                    dr = dtPO.NewRow();
                                    dicRow.Add(sKey, dtPO.Rows.Count);

                                    dr["REGION"] = string.Join(string.Empty,
                                                               _Customer.CustomerBigRegion.Split(' ').Select(x => x.First()))
                                                         .ToUpper();
                                    dr["PCODE"] = _Product.ProductCode;
                                    dr["Note"]  = _Product.ProductCode.Substring(0, 1) == "K"
                                                      ? "Ok"
                                                      : _Product.ProductNote.Contains(_Customer.CustomerBigRegion == "Miền Nam"
                                                                                          ? "South"
                                                                                          : "North")
                                                          ? "Ok"
                                                          : "Out of List";
                                    dr["CTYPE"]         = _Customer.CustomerType;
                                    dr["P&L"]           = _Customer.Company;
                                    dr["DateOrder"]     = PODate.DateOrder.Date;
                                    dr["QuantityOrder"] = _CustomerOrder.QuantityOrder;
                                    dr["DateReceive"]   = _Customer.CustomerBigRegion == "Miền Nam"
                                                              ? PODate.DateOrder.Date.AddDays(1)
                                                              : PODate.DateOrder.Date;

                                    dtPO.Rows.Add(dr);
                                }
                                else
                                {
                                    dr                  = dtPO.Rows[_rowIndex];
                                    dr["QuantityOrder"] = (double) dr["QuantityOrder"] + _CustomerOrder.QuantityOrder;
                                }
                            }
                        }
                    }
                }

                //Excel.Application xlApp = new Excel.Application();
                //Aspose.Cells.Workbook xlWb = new Aspose.Cells.Workbook();

                string fileName =
                    $"PO {DateFrom.ToString("dd.MM") + " - " + DateTo.ToString("dd.MM") + " (" + DateTime.Now.ToString("yyyyMMdd HH\\hmm") + ")"} - {Choice}.xlsx";

                string fileNameXlsb =
                    $"PO {DateFrom.ToString("dd.MM") + " - " + DateTo.ToString("dd.MM") + " (" + DateTime.Now.ToString("yyyyMMdd HH\\hmm") + ")"} - {Choice}.xlsb";

                // var path = @"D:\Documents\Stuff\VinEco\Mastah Project\Test\";
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
            foreach (DateTime DemandDate in coreStructure.dicPO.Keys.OrderByDescending(x => x.Date).Reverse())
                // Second layer - Priority Target.
            {
                foreach (string PriorityTarget in ListPriorityTarget)
                {
                    var TemporaryProductDictionary = new Dictionary<Product, bool>();

                    Dictionary<Product, Dictionary<CustomerOrder, bool>> PONorth = coreStructure.dicPO[DemandDate];
                    Dictionary<Product, Dictionary<CustomerOrder, bool>> POSouth =
                        coreStructure.dicPO[DemandDate.AddDays(-coreStructure.dicTransferDays["Highland-North"])];

                    foreach (Product CurrentProduct in PONorth.Keys)
                    {
                        if (!TemporaryProductDictionary.ContainsKey(CurrentProduct))
                        {
                            TemporaryProductDictionary.Add(CurrentProduct, true);
                        }
                    }

                    foreach (Product CurrentProduct in POSouth.Keys)
                    {
                        if (!TemporaryProductDictionary.ContainsKey(CurrentProduct))
                        {
                            TemporaryProductDictionary.Add(CurrentProduct, true);
                        }
                    }

                    foreach (Product CurrentProduct in TemporaryProductDictionary.Keys)
                    {
                        var _result        = new Dictionary<SupplierForecast, bool>();
                        var SupplyNorth    = new Dictionary<Guid, SupplierForecast>();
                        var SupplySouth    = new Dictionary<Guid, SupplierForecast>();
                        var SupplyHighland = new Dictionary<Guid, SupplierForecast>();

                        if (coreStructure.dicFC[DemandDate.AddDays(-coreStructure.dicTransferDays["North-North"])]
                                         .TryGetValue(CurrentProduct, out _result))
                        {
                            SupplyNorth = _result.Keys.Where(x =>
                                                                 x.Availability.Contains(
                                                                     (DemandDate.AddDays(-coreStructure.dicTransferDays["North-North"]).DayOfWeek + 1)
                                                                    .ToString()))
                                                 .ToDictionary(x => x.SupplierForecastId);
                        }

                        if (coreStructure.dicFC[DemandDate.AddDays(-coreStructure.dicTransferDays["Highland-North"])]
                                         .TryGetValue(CurrentProduct, out _result))
                        {
                            SupplySouth = _result.Keys.Where(x =>
                                                                 x.Availability.Contains(
                                                                     (DemandDate.AddDays(-coreStructure.dicTransferDays["Highland-North"]).DayOfWeek + 1)
                                                                    .ToString()))
                                                 .ToDictionary(x => x.SupplierForecastId);
                        }

                        ;

                        if (coreStructure.dicFC[DemandDate.AddDays(-coreStructure.dicTransferDays["Highland-North"])]
                                         .TryGetValue(CurrentProduct, out _result))
                        {
                            SupplyHighland = _result.Keys.Where(x =>
                                                                    x.Availability.Contains(
                                                                        (DemandDate.AddDays(-coreStructure.dicTransferDays["Highland-North"]).DayOfWeek + 1)
                                                                       .ToString()))
                                                    .ToDictionary(x => x.SupplierForecastId);
                        }

                        ;

                        var ListRate = new double[3];

                        // Total Demand. Customers' Regions
                        double DemandNorth = !PONorth.ContainsKey(CurrentProduct)
                                                 ? 0
                                                 : PONorth[CurrentProduct]
                                                  .Keys
                                                  .Where(x =>
                                                             coreStructure.dicCustomer[x.CustomerId].CustomerBigRegion  == "Miền Bắc" &&
                                                             PriorityTarget                                             != ""
                                                                 ? coreStructure.dicCustomer[x.CustomerId].CustomerType == PriorityTarget
                                                                 : true)
                                                  .Sum(x => x.QuantityOrderKg);

                        // In case VM+, have to calculate rate twice coz fuck the police. Really.
                        double DemandNorthVM = !PriorityTarget.Contains("VM+")
                                                   ? 0
                                                   : PONorth[CurrentProduct]
                                                    .Keys
                                                    .Where(x =>
                                                               coreStructure.dicCustomer[x.CustomerId].CustomerBigRegion == "Miền Bắc" &&
                                                               coreStructure.dicCustomer[x.CustomerId].CustomerType      == "VM")
                                                    .Sum(x => x.QuantityOrderKg);

                        double DemandSouth = !POSouth.ContainsKey(CurrentProduct)
                                                 ? 0
                                                 : POSouth[CurrentProduct]
                                                  .Keys
                                                  .Where(x =>
                                                             coreStructure.dicCustomer[x.CustomerId].CustomerBigRegion  == "Miền Nam" &&
                                                             PriorityTarget                                             != ""
                                                                 ? coreStructure.dicCustomer[x.CustomerId].CustomerType == PriorityTarget
                                                                 : true)
                                                  .Sum(x => x.QuantityOrderKg);

                        double DemandSouthVM = !PriorityTarget.Contains("VM+")
                                                   ? 0
                                                   : PONorth[CurrentProduct]
                                                    .Keys
                                                    .Where(x =>
                                                               coreStructure.dicCustomer[x.CustomerId].CustomerBigRegion == "Miền Nam" &&
                                                               coreStructure.dicCustomer[x.CustomerId].CustomerType      == "VM")
                                                    .Sum(x => x.QuantityOrderKg);

                        // Total Missing. Customers' Regions
                        double MissingNorth = DemandNorth - SupplyNorth.Values.Sum(x => x.QuantityForecast);
                        double MissingSouth = DemandSouth - SupplySouth.Values.Sum(x => x.QuantityForecast);

                        double QtyNorthNoXRegion =
                            SupplyNorth.Values.Where(x => !x.CrossRegion).Sum(x => x.QuantityForecast);
                        double QtyNorthXRegion = SupplyNorth.Values.Where(x => x.CrossRegion).Sum(x => x.QuantityForecast);

                        double QtySouthNoXRegion =
                            SupplySouth.Values.Where(x => !x.CrossRegion).Sum(x => x.QuantityForecast);
                        double QtySouthXRegion = SupplySouth.Values.Where(x => x.CrossRegion).Sum(x => x.QuantityForecast);

                        // Credit goes to someone very special, for figuring out the entire logic, the simplest way.
                        // Made by her. Hah!
                        double QtySouthCanSpare = Math.Min(Math.Max(QtySouthNoXRegion + QtySouthXRegion - DemandSouth, 0),
                                                           QtySouthXRegion);
                        double QtyNorthCanSpare = Math.Min(Math.Max(QtySouthNoXRegion + QtySouthXRegion - DemandSouth, 0),
                                                           QtySouthXRegion);

                        double QtyHighland = SupplyHighland.Values.Sum(x => x.QuantityForecast);

                        var _ProductCrossRegion   = new ProductCrossRegion();
                        var flagNoHighlandToNorth = true;
                        if (coreStructure.dicProductCrossRegion.TryGetValue(CurrentProduct.ProductId,
                                                                            out _ProductCrossRegion))
                        {
                            if (!_ProductCrossRegion.ToNorth)
                            {
                                flagNoHighlandToNorth = false;
                            }
                        }

                        double QtyHighlandToNorth = flagNoHighlandToNorth ? QtyHighland : 0;

                        double RateNorth =
                            (QtyNorthNoXRegion                                  +
                             QtyNorthXRegion                                    +
                             QtyHighlandToNorth * (MissingNorth / (MissingNorth + MissingSouth)) +
                             QtySouthCanSpare)  /
                            DemandNorth;

                        double RateNorthWithVM =
                            (QtyNorthNoXRegion                                  +
                             QtyNorthXRegion                                    +
                             QtyHighlandToNorth * (MissingNorth / (MissingNorth + MissingSouth)) +
                             QtySouthCanSpare)  /
                            (DemandNorth + DemandNorthVM);

                        if (RateNorthWithVM < 1)
                        {
                            RateNorth = 1;
                        }

                        RateNorth = Math.Min(RateNorth, UpperCap);

                        double RateSouth =
                            (QtySouthNoXRegion                                 +
                             QtySouthXRegion                                   +
                             QtyHighland       * (MissingSouth / (MissingNorth + MissingSouth)) +
                             QtyNorthCanSpare) /
                            DemandSouth;

                        double RateSouthWithVM =
                            (QtySouthNoXRegion                                 +
                             QtySouthXRegion                                   +
                             QtyHighland       * (MissingSouth / (MissingNorth + MissingSouth)) +
                             QtyNorthCanSpare) /
                            (DemandSouth + DemandSouthVM);

                        if (RateSouthWithVM < 1)
                        {
                            RateSouth = 1;
                        }

                        RateSouth = Math.Min(RateSouth, UpperCap);
                    }
                }
            }
        }

        private void CoordDoWhile(CoordStructure coreStructure, string SupplierRegion, string CustomerRegion,
                                  string         SupplierType, byte    dayBefore, byte        dayLdBefore, double UpperLimit = 1, bool  CrossRegion = false,
                                  string         PriorityTarget                                                              = "", bool YesNoByUnit = false, bool YesNoContracted = false, bool YesNoKPI = false)
        {
            try
            {
                /// <* IMPORTANTO! *>
                // Nothing shall begin before this happens
                Stopwatch stopwatch = Stopwatch.StartNew();

                #region Preparing.

                #endregion

                // PO Date Layer.
                //Console.Write("{0} => {1}, {2}{3}", String.Concat(SupplierRegion.Split(' ').Select(x => x.First())), String.Concat(CustomerRegion.Split(' ').Select(x => x.First().ToString().ToUpper())), SupplierType, (PriorityTarget != "" ? " " + PriorityTarget : ""));
                foreach (DateTime DatePO in coreStructure.dicPO.Keys.OrderByDescending(x => x.Date).Reverse())
                    // Product Layer.
                {
                    foreach (Product _Product in coreStructure.dicPO[DatePO]
                                                              .Keys.OrderByDescending(x => x.ProductCode)
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
                            (_Product.ProductCode                == "K01901" || _Product.ProductCode == "K02201"))
                        {
                            _MOQ = 0.3;
                        }

                        //}

                        restartThis:

                        /// <! For Debuging Purposes Only !>
                        // Only uncomment in very specific debugging situation.
                        //if (_Product.ProductCode == "A04801" && DatePO.Day == 26 && CustomerRegion == "Miền Nam" && SupplierRegion == "Miền Nam" && SupplierType == "VCM")
                        //{
                        //    string WhatAmIEvenDoing = "I have no freaking idea.";
                        //}

                        // Skip if product is not in the List VinEco supplies.
                        if (SupplierType                         != "VinEco" &&
                            _Product.ProductCode.Substring(0, 1) != "K"      &&
                            (PriorityTarget                      == "VM" || PriorityTarget == "VM+"))
                        {
                            if (!_Product.ProductNote.Contains(CustomerRegion == "Miền Bắc" ? "North" : "South"))
                            {
                                continue;
                            }
                        }

                        // Dealing with cases of some Products that will not go to either region, from Lâm Đồng
                        var _ProductCrossRegion = new ProductCrossRegion();
                        if (coreStructure.dicProductCrossRegion.TryGetValue(_Product.ProductId, out _ProductCrossRegion) &&
                            SupplierRegion == "Lâm Đồng")
                        {
                            switch (CustomerRegion)
                            {
                                case "Miền Bắc":
                                    if (!_ProductCrossRegion.ToNorth)
                                    {
                                        continue;
                                    }

                                    break;
                                case "Miền Nam":
                                    if (!_ProductCrossRegion.ToSouth)
                                    {
                                        continue;
                                    }

                                    break;
                                default: break;
                            }
                        }

                        #region Demand from Chosen Customers.

                        // Total Order.
                        double sumVCM = coreStructure.dicPO[DatePO][_Product]
                                                     .Where(x =>
                                                                coreStructure.dicCustomer[x.Key.CustomerId].CustomerBigRegion == CustomerRegion &&
                                                                x.Value                                                                         &&
                                                                (PriorityTarget                                                 != ""
                                                                     ? coreStructure.dicCustomer[x.Key.CustomerId].CustomerType == PriorityTarget
                                                                     : true))
                                                     .Sum(x => x.Key.QuantityOrderKg); // Sum of Demand.

                        double sumVM = PriorityTarget.Contains("VM+")
                                           ? coreStructure.dicPO[DatePO][_Product]
                                                          .Where(x =>
                                                                     coreStructure.dicCustomer[x.Key.CustomerId].CustomerBigRegion == CustomerRegion &&
                                                                     x.Value                                                                         &&
                                                                     coreStructure.dicCustomer[x.Key.CustomerId].CustomerType.Contains("VM")         &&
                                                                     !coreStructure.dicCustomer[x.Key.CustomerId].CustomerType.Contains("VM+"))
                                                          .Sum(x => x.Key.QuantityOrderKg)
                                           : 0; // Sum of Demand.

                        double sumVcmMN = sumVCM + sumVM;

                        if (SupplierRegion == "Lâm Đồng")
                        {
                            DateTime _DatePO = CustomerRegion == "Miền Nam" ? DatePO.AddDays(2) : DatePO.AddDays(-2);
                            if (coreStructure.dicPO.ContainsKey(_DatePO) &&
                                coreStructure.dicPO[_DatePO].ContainsKey(_Product))
                            {
                                string _CustomerRegion = CustomerRegion == "Miền Nam" ? "Miền Bắc" : "Miền Nam";
                                sumVCM                 += coreStructure.dicPO[_DatePO][_Product]
                                                                       .Where(x =>
                                                                                  coreStructure.dicCustomer[x.Key.CustomerId].CustomerBigRegion == _CustomerRegion &&
                                                                                  x.Value                                                                          &&
                                                                                  (PriorityTarget                                                 != ""
                                                                                       ? coreStructure.dicCustomer[x.Key.CustomerId].CustomerType == PriorityTarget
                                                                                       : true))
                                                                       .Sum(x => x.Key.QuantityOrderKg);

                                sumVM += PriorityTarget.Contains("VM+")
                                             ? coreStructure.dicPO[_DatePO][_Product]
                                                            .Where(x =>
                                                                       coreStructure.dicCustomer[x.Key.CustomerId].CustomerBigRegion ==
                                                                       _CustomerRegion                                                         &&
                                                                       x.Value                                                                 &&
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

                        KeyValuePair<DateTime, Dictionary<Product, Dictionary<SupplierForecast, bool>>> _dicProductFC =
                            coreStructure.dicFC.Where(x => x.Key.Date == DatePO.AddDays(-dayBefore))
                                         .FirstOrDefault();
                        KeyValuePair<DateTime, Dictionary<Product, Dictionary<SupplierForecast, bool>>> _dicProductFcLd =
                            coreStructure.dicFC.Where(x => x.Key.Date == DatePO.AddDays(-dayLdBefore))
                                         .FirstOrDefault();

                        if (sumVCM != 0 && _dicProductFC.Value != null)
                        {
                            double sumThuMuaLd = 0;
                            double sumFarmLd   = 0;

                            #region Supply from Lâm Đồng

                            if (SupplierRegion != "Lâm Đồng" && _dicProductFcLd.Value != null)
                            {
                                // Check if Inventory has stock in other places.
                                // If no, equally distributed stuff.
                                // If yes, hah hah hah no.
                                KeyValuePair<Product, Dictionary<SupplierForecast, bool>> dicSupplierLdFC = _dicProductFcLd
                                                                                                           .Value
                                                                                                           .Where(x => x.Key.ProductCode == _Product.ProductCode)
                                                                                                           .FirstOrDefault();
                                if (dicSupplierLdFC.Value != null)
                                {
                                    // Check Lâm Đồng
                                    // Please NEVER FullOrder == true.
                                    //var _SupplierThuMuaLd = 

                                    IEnumerable<KeyValuePair<SupplierForecast, bool>> _dicSupplierLdFC = dicSupplierLdFC
                                                                                                        .Value
                                                                                                        .Where(x =>
                                                                                                                   coreStructure.dicSupplier[x.Key.SupplierId].SupplierRegion == "Lâm Đồng"                               &&
                                                                                                                   (x.Key.Target                                              == "All" || x.Key.Target == PriorityTarget) &&
                                                                                                                   (YesNoKPI
                                                                                                                        ? x.Key.QuantityForecastPlanned
                                                                                                                        : YesNoContracted
                                                                                                                            ? x.Key.QuantityForecastContracted
                                                                                                                            : x.Key.QuantityForecast) >
                                                                                                                   0);

                                    // Normal case
                                    sumFarmLd = _dicSupplierLdFC
                                               .Where(x =>
                                                          coreStructure.dicSupplier[x.Key.SupplierId].SupplierType == "VinEco")
                                               .Sum(x => x.Key.QuantityForecast);

                                    sumThuMuaLd = _dicSupplierLdFC
                                                 .Where(x =>
                                                            coreStructure.dicSupplier[x.Key.SupplierId].SupplierType != "VinEco" &&
                                                            x.Key.Availability.Contains(
                                                                Convert.ToString((int) DatePO.AddDays(-dayLdBefore).DayOfWeek + 1)))
                                                 .Sum(x => x.Key.QuantityForecast);
                                }
                            }

                            #endregion

                            KeyValuePair<Product, Dictionary<SupplierForecast, bool>> dicSupplierFC = _dicProductFC.Value
                                                                                                                   .Where(x => x.Key.ProductCode == _Product.ProductCode)
                                                                                                                   .FirstOrDefault();
                            if (dicSupplierFC.Value != null)
                            {
                                #region Total Supply.

                                IEnumerable<KeyValuePair<SupplierForecast, bool>> _resultSupplier = dicSupplierFC.Value
                                                                                                                 .Where(x =>
                                                                                                                            coreStructure.dicSupplier[x.Key.SupplierId].SupplierRegion == "VinEco"                                 &&
                                                                                                                            coreStructure.dicSupplier[x.Key.SupplierId].SupplierType   == SupplierType                             &&
                                                                                                                            (x.Key.Target                                              == "All" || x.Key.Target == PriorityTarget) &&
                                                                                                                            (SupplierType                                              != "VinEco"
                                                                                                                                 ? x.Key.Availability.Contains(
                                                                                                                                     Convert.ToString((int) DatePO.AddDays(-dayBefore).DayOfWeek + 1))
                                                                                                                                 : true));

                                IEnumerable<KeyValuePair<SupplierForecast, bool>> _dicSupplierFC = dicSupplierFC.Value
                                                                                                                .Where(x =>
                                                                                                                           coreStructure.dicSupplier[x.Key.SupplierId].SupplierRegion == SupplierRegion                           &&
                                                                                                                           (x.Key.Target                                              == "All" || x.Key.Target == PriorityTarget) &&
                                                                                                                           (YesNoKPI
                                                                                                                                ? x.Key.QuantityForecastPlanned
                                                                                                                                : YesNoContracted
                                                                                                                                    ? x.Key.QuantityForecastContracted
                                                                                                                                    : x.Key.QuantityForecast) >
                                                                                                                           0);

                                double sumFarm = _dicSupplierFC
                                                .Where(x =>
                                                           coreStructure.dicSupplier[x.Key.SupplierId].SupplierType == "VinEco")
                                                .Sum(x => x.Key.QuantityForecast);

                                double sumThuMua = _dicSupplierFC
                                                  .Where(x =>
                                                             coreStructure.dicSupplier[x.Key.SupplierId].SupplierType != "VinEco" &&
                                                             x.Key.Availability.Contains(
                                                                 Convert.ToString((int) DatePO.AddDays(-dayBefore).DayOfWeek + 1)))
                                                  .Sum(x => x.Key.QuantityForecast);

                                //_resultSupplier
                                //    .Sum(x => YesNoKPI ? x.Key.QuantityForecastPlanned : YesNoContracted ? x.Key.QuantityForecastContracted : x.Key.QuantityForecast);

                                var flagFullOrder = false;

                                double sumVE = sumFarm + sumThuMua;

                                DateTime _DatePO = SupplierRegion == "Miền Bắc"
                                                       ? DatePO.AddDays(-2).Date
                                                       : DatePO.AddDays(2).Date;
                                if (CustomerRegion == "Miền Nam"             &&
                                    coreStructure.dicPO.ContainsKey(_DatePO) &&
                                    coreStructure.dicPO[_DatePO].ContainsKey(_Product))
                                {
                                    sumVE += Math.Max(sumFarmLd   +
                                                      sumThuMuaLd -
                                                      coreStructure.dicPO[_DatePO][_Product]
                                                                   .Where(x =>
                                                                              coreStructure.dicCustomer[x.Key.CustomerId]
                                                                                           .CustomerBigRegion ==
                                                                              (CustomerRegion                 == "Miền Bắc" ? "Miền Nam" : "Miền Bắc") &&
                                                                              x.Value)
                                                                   .Sum(x => x.Key.QuantityOrderKg), 0);
                                }
                                else
                                {
                                    sumVE += sumFarmLd + sumThuMuaLd;
                                }

                                if (_resultSupplier.Where(x => YesNoKPI || YesNoContracted ? false : x.Key.FullOrder)
                                                   .FirstOrDefault()
                                                   .Key !=
                                    null)
                                {
                                    flagFullOrder = true;
                                }
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
                                    double rate = sumVE / (sumVCM + sumVM);

                                    // If Screw-the-upper-limit flag is up.
                                    if (flagFullOrder)
                                    {
                                        rate = UpperCap;
                                    }
                                    // If it's VinCommerce's Supplier, always 1.
                                    else if (rate < 1 && SupplierType == "VCM" && sumVE > 0)
                                    {
                                        rate = UpperCap;
                                    }
                                    // Otherwise, in case of an UpperLimit, obey it
                                    else if (!flagFullOrder)
                                    {
                                        if (rate < 1)
                                        {
                                            rate = Math.Max(sumVE / sumVCM, 1);
                                            rate = SupplierRegion != "Lâm Đồng" &&
                                                   (YesNoKPI        ||
                                                    sumFarm     > 0 ||
                                                    sumFarmLd   > 0 ||
                                                    sumThuMua   > 0 ||
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
                                                   (YesNoKPI        ||
                                                    sumFarm     > 0 ||
                                                    sumFarmLd   > 0 ||
                                                    sumThuMua   > 0 ||
                                                    sumThuMuaLd > 0)
                                                       ? Math.Max(rate, 1)
                                                       : rate;
                                            if (rate < 1 && SupplierType == "VCM" && sumVE > 0)
                                            {
                                                rate = UpperCap;
                                            }

                                            //}
                                        }
                                    }

                                    rate = UpperLimit > 0 ? Math.Min(rate, UpperLimit) : rate;

                                    #endregion

                                    // Only the bravest would tread deeper.
                                    // ... I was once young, brave and foolish ...

                                    // Optimization - Filtering Customer Orders that has been dealt with.
                                    //var ListCustomerOrder = coreStructure.dicPO[DatePO][_Product].Where(x => x.Value == true).ToDictionary(x => x.Key);
                                    List<CustomerOrder> ValidCustomerList = coreStructure.dicPO[DatePO][_Product]
                                                                                         .Where(x => x.Value)
                                                                                         .ToDictionary(x => x.Key)
                                                                                         .Keys
                                                                                         .Where(x =>
                                                                                                    x.QuantityOrderKg                                           >= 0.1            &&
                                                                                                    coreStructure.dicCustomer[x.CustomerId].CustomerBigRegion   == CustomerRegion &&
                                                                                                    (PriorityTarget                                             != ""
                                                                                                         ? coreStructure.dicCustomer[x.CustomerId].CustomerType == PriorityTarget
                                                                                                         : true)                                                         
                                                                                                    //     &&
                                                                                                    //(x.DesiredRegion == null ? true : x.DesiredRegion == SupplierRegion) &&
                                                                                                    //(x.DesiredSource == null ? true : x.DesiredSource == SupplierType)
                                                                                                    )
                                                                                         .OrderByDescending(x => coreStructure.dicCustomer[x.CustomerId].CustomerCode)
                                                                                          //.Reverse()
                                                                                         .ToList();

                                    do
                                    {
                                        #region Qualified Suppliers.

                                        SupplierForecast _SupplierForecast = null;

                                        Dictionary<SupplierForecast, KeyValuePair<SupplierForecast, bool>>
                                            _dicSupplierFC_inner = dicSupplierFC.Value
                                                                                .Where(x => x.Key.QuantityForecast                                     >= _MOQ)
                                                                                .Where(x => coreStructure.dicSupplier[x.Key.SupplierId].SupplierRegion ==
                                                                                            SupplierRegion &&
                                                                                            coreStructure.dicSupplier[x.Key.SupplierId].SupplierType ==
                                                                                            SupplierType &&
                                                                                            (SupplierType != "VinEco"
                                                                                                 ? x.Key.Availability.Contains(
                                                                                                     Convert.ToString(
                                                                                                         (int) DatePO.AddDays(-dayBefore).DayOfWeek + 1))
                                                                                                 : true)                                              &&
                                                                                            (x.Key.Target == "All" || x.Key.Target == PriorityTarget) &&
                                                                                            (CrossRegion ? x.Key.CrossRegion : true))
                                                                                .OrderBy(x => x.Key.Level)
                                                                                .ThenByDescending(x => x.Key.FullOrder)
                                                                                .ThenBy(x =>
                                                                                            coreStructure.dicDeli[DatePO.AddDays(-dayBefore)][_Product][x.Key])
                                                                                .ThenByDescending(x => x.Key.QuantityForecast)
                                                                                .ThenByDescending(x => x.Key.LabelVinEco)
                                                                                .ToDictionary(x => x.Key);

                                        KeyValuePair<SupplierForecast, KeyValuePair<SupplierForecast, bool>> result =
                                            _dicSupplierFC_inner.FirstOrDefault();
                                        if (result.Key == null)
                                        {
                                            break;
                                        }

                                        CustomerOrder _CustomerOrder = ValidCustomerList
                                                                      .Where(x => x.QuantityOrderKg * rate <= result.Key.QuantityForecast)
                                                                      .FirstOrDefault();

                                        if (_CustomerOrder == null)
                                        {
                                            _CustomerOrder = ValidCustomerList.OrderBy(x => x.QuantityOrderKg)
                                                                              .FirstOrDefault();
                                        }

                                        if (_CustomerOrder == null)
                                        {
                                            break;
                                        }

                                        // Coz for fuck sake, it can return null

                                        int totalSupplier = _dicSupplierFC_inner.Count();
                                        _SupplierForecast = result.Key;

                                        #endregion

                                        double _rate = rate;
                                        if (coreStructure.dicPO[DatePO][_Product].Count <= totalSupplier)
                                        {
                                            _rate = 1;
                                        }

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
                                                                                    out _SupplierForecastCoord) &&
                                                        _SupplierForecastCoord == null)
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
                                                                        : _SupplierForecast.QuantityForecast)) >=
                                                            _MOQ)
                                                        {
                                                            wallet = _MOQ;
                                                        }

                                                        //if (_MOQ == 0.05)
                                                        //{
                                                        //    // Let's hope this will never be hit.
                                                        //    // I fucking do hope that.
                                                        //    string OhMyFuckingGodWhy = "Holy shit idk, why, oh god, why";
                                                        //}

                                                        #endregion

                                                        if (wallet < _MOQ && PriorityTarget != "")
                                                        {
                                                            wallet = _MOQ;
                                                        }

                                                        if (wallet >= _MOQ && _SupplierForecast.QuantityForecast >= _MOQ)
                                                        {
                                                            //if (sumVE <= 0) { continue; }
                                                            // Honestly, this should never be hit
                                                            // Jk I changed stuff. This should ALWAYS be hit
                                                            _SupplierForecastCoord =
                                                                new Dictionary<SupplierForecast, DateTime>();

                                                            double _QuantityForecast = Math.Min(wallet,
                                                                                                _SupplierForecast.QuantityForecast);

                                                            if (YesPlanningFuckMe)
                                                            {
                                                                _QuantityForecast =
                                                                    Math.Min(Math.Max(wallet / totalSupplier, _MOQ),
                                                                             _SupplierForecast.QuantityForecast);
                                                            }

                                                            if (UpperCap > 0)
                                                            {
                                                                _QuantityForecast =
                                                                    Math.Min(_CustomerOrder.QuantityOrderKg * UpperLimit,
                                                                             _QuantityForecast);
                                                            }

                                                            _QuantityForecast = Math.Round(_QuantityForecast, 1);

                                                            #region Unit.

                                                            if (_CustomerOrder.Unit != "Kg")
                                                            {
                                                                ProductUnitRegion something = coreStructure
                                                                                             .dicProductUnit[_Product.ProductCode]
                                                                                             .ListRegion
                                                                                             .FirstOrDefault(x =>
                                                                                                                 x.OrderUnitType == _CustomerOrder.Unit);
                                                                if (something                                                    != null)
                                                                {
                                                                    double _SaleUnitPer = something.SaleUnitPer;
                                                                    _QuantityForecast   =
                                                                        _QuantityForecast / _MOQ * _SaleUnitPer;
                                                                }
                                                            }

                                                            #endregion

                                                            #region Defer extra days for Crossing Regions ( North --> South and vice versa. )

                                                            // To coup with merging PO ( Tue Thu Sat to Mon Wed Fri )
                                                            DateTime _Date = DatePO.AddDays(-dayBefore).Date;
                                                            if (CrossRegion                   &&
                                                                _SupplierForecast.CrossRegion &&
                                                                CustomerRegion == "Miền Bắc"  &&
                                                                SupplierRegion ==
                                                                "Miền Nam" /*&& _Product.ProductCode.Substring(0, 1) == "K"*/ &&
                                                                (_Date.DayOfWeek == DayOfWeek.Tuesday  ||
                                                                 _Date.DayOfWeek == DayOfWeek.Thursday ||
                                                                 _Date.DayOfWeek == DayOfWeek.Saturday))
                                                            {
                                                                _Date = _Date.AddDays(-1).Date;
                                                            }

                                                            #endregion

                                                            // To coup with Supply has custom rates, depending on Region.
                                                            var    _ProductRate = new ProductRate();
                                                            double _Rate        = 1;
                                                            if (!YesNoKPI                    &&
                                                                SupplierRegion == "Miền Nam" &&
                                                                coreStructure.dicProductRate.TryGetValue(
                                                                    _Product.ProductCode, out _ProductRate))
                                                            {
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
                                                            }

                                                            Guid newId = Guid.NewGuid();
                                                            _SupplierForecastCoord.Add(new SupplierForecast
                                                                                           {
                                                                                               _id                = newId,
                                                                                               SupplierForecastId = newId,

                                                                                               SupplierId         = _SupplierForecast.SupplierId,
                                                                                               LabelVinEco        = _SupplierForecast.LabelVinEco,
                                                                                               FullOrder          = _SupplierForecast.FullOrder,
                                                                                               QualityControlPass = _SupplierForecast.QualityControlPass,
                                                                                               CrossRegion        = _SupplierForecast.CrossRegion,
                                                                                               Level              = _SupplierForecast.Level,
                                                                                               Availability       = _SupplierForecast.Availability,
                                                                                               Target             = _SupplierForecast.Target,

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
                                                            _SupplierForecast.QuantityForecast         -= _QuantityForecast;
                                                            _SupplierForecast.QuantityForecastOriginal -= _QuantityForecast;
                                                            if (!_SupplierForecast.FullOrder &&
                                                                _SupplierForecast.QuantityForecast <= 0)
                                                            {
                                                                _SupplierForecast.QuantityForecast = _MOQ * 7;
                                                            }
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

                                                            if (YesPlanningFuckMe &&
                                                                _CustomerOrder.QuantityOrder >=
                                                                _QuantityForecast)
                                                            {
                                                                var CustomerOrder = new CustomerOrder();

                                                                //CustomerOrder.Company         = _CustomerOrder.Company;
                                                                CustomerOrder.CustomerId      = _CustomerOrder.CustomerId;
                                                                //CustomerOrder.CustomerOrderId = Guid.NewGuid();
                                                                //CustomerOrder.DesiredRegion   = _CustomerOrder.DesiredRegion;
                                                                //CustomerOrder.DesiredSource   = _CustomerOrder.DesiredSource;
                                                                CustomerOrder.QuantityOrder   =
                                                                    _CustomerOrder.QuantityOrder - _QuantityForecast;
                                                                CustomerOrder.QuantityOrderKg =
                                                                    _CustomerOrder.QuantityOrderKg - _QuantityForecast;
                                                                CustomerOrder.Unit = _CustomerOrder.Unit;
                                                                //CustomerOrder._id  = CustomerOrder.CustomerOrderId;
                                                                CustomerOrder._id = Guid.NewGuid();

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
                                                        {
                                                            coreStructure.dicPO[DatePO].Remove(_Product);
                                                        }

                                                        if (coreStructure.dicPO[DatePO].Keys.Count == 0)
                                                        {
                                                            coreStructure.dicPO.Remove(DatePO);
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    } while (ValidCustomerList.Count > 0);
                                }
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
                var            mongoClient = new MongoClient();
                IMongoDatabase db          = mongoClient.GetDatabase("localtest");

                var PO                 = new List<PurchaseOrderDate>(); // mongoClient.GetDatabase("localtest").GetCollection<PurchaseOrderDate>("PurchaseOrder").AsQueryable().ToList();
                List<Product> Product  = db.GetCollection<Product>("Product").AsQueryable().ToList();
                var           Customer = new List<Customer>(); //  db.GetCollection<Customer>("Customer").AsQueryable().ToList();

                var dicPO       = new Dictionary<DateTime, Dictionary<string, Dictionary<string, Guid>>>(1000);
                var dicProduct  = new Dictionary<string, Product>(1000);
                var dicCustomer = new Dictionary<string, Customer>(10000);

                // Product Dictionary.
                foreach (Product _Product in Product)
                {
                    if (!dicProduct.ContainsKey(_Product.ProductCode))
                    {
                        dicProduct.Add(_Product.ProductCode, _Product);
                    }
                }

                // Customer Dictionary.
                foreach (Customer _Customer in Customer)
                {
                    if (!dicCustomer.ContainsKey(_Customer.CustomerCode + _Customer.CustomerType))
                    {
                        dicCustomer.Add(_Customer.CustomerCode + _Customer.CustomerType, _Customer);
                    }
                }

                string filePath = $"D:\\Documents\\Stuff\\VinEco\\Mastah Project\\{fileNameMB}";

                string conStr = string.Format(Constants.Excel07ConString, filePath, header);

                var directoryPath = "D:\\Documents\\Stuff\\VinEco\\Mastah Project\\PO";

                #region Reading PO files in folder.

                var        dirInfo  = new DirectoryInfo(directoryPath);
                FileInfo[] ListFile = dirInfo.GetFiles();

                await db.DropCollectionAsync("PurchaseOrder");

                foreach (FileInfo _FileInfo in ListFile)
                {
                    var                    opt        = new LoadOptions { MemorySetting = MemorySetting.MemoryPreference };
                    var                    xlWbAspose = new Workbook(_FileInfo.FullName, opt);
                    Aspose.Cells.Worksheet xlWsAspose =
                        xlWbAspose.Worksheets.OrderByDescending(x => x.Cells.MaxDataRow).First();

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

                    EatPOAspose(PO,
                                xlWsAspose,
                                string.Format(Constants.Excel07ConString, _FileInfo.FullName, header),
                                _Region,
                                dicPO,
                                dicProduct,
                                dicCustomer,
                                Product,
                                Customer,
                                false);

                    //await db.GetCollection<PurchaseOrderDate>("PurchaseOrder").InsertManyAsync(PO);
                    //PO.Clear();

                    stopwatch.Stop();

                    WriteToRichTextBoxOutput(
                        $"- Done in {Math.Round(stopwatch.Elapsed.TotalSeconds, 2)}s!",
                        false);
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

                await db.DropCollectionAsync("PurchaseOrder");
                //await db.GetCollection<PurchaseOrderDate>("PurchaseOrder");

                var list_smaller = new List<PurchaseOrderDate>();

                //int threshold = PO.Count / 2;
                //for (int i = 0; i < PO.Count; i++)
                //{
                //    list_smaller.Add(PO[i]);
                //    if (list_smaller.Count > threshold)
                //    {
                //        await db.GetCollection<PurchaseOrderDate>("PurchaseOrder").InsertManyAsync(list_smaller);
                //        list_smaller = new List<PurchaseOrderDate>();
                //    }
                //}

                foreach (PurchaseOrderDate poItem in PO)
                {
                    list_smaller.Add(poItem);
                    await db.GetCollection<PurchaseOrderDate>("PurchaseOrder").InsertManyAsync(list_smaller);
                    WriteToRichTextBoxOutput($"Done {PO.IndexOf(poItem) + 1}/{PO.Count}");
                    list_smaller.Clear();
                }

                //if (list_smaller.Count > 0)
                //    await db.GetCollection<PurchaseOrderDate>("PurchaseOrder").InsertManyAsync(list_smaller);


                //await db.GetCollection<PurchaseOrderDate>("PurchaseOrder").InsertManyAsync(PO);

                await db.DropCollectionAsync("Product");
                await db.GetCollection<Product>("Product").InsertManyAsync(Product);

                await db.DropCollectionAsync("Customer");
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
        ///     Updating OpenConfig file & Do afterward updating.
        /// </summary>
        private async Task UpdateOpenConfig()
        {
            try
            {
                //Console.OutputEncoding = System.Text.Encoding.UTF8;

                WriteToRichTextBoxOutput("Start!");

                #region Initialization.

                var            mongoClient = new MongoClient();
                IMongoDatabase db          = mongoClient.GetDatabase("localtest");

                List<Product> Product         = db.GetCollection<Product>("Product").AsQueryable().ToList();
                var           ProductUnitList = new List<ProductUnit>();

                string filePath =
                    string.Format("D:\\Documents\\Stuff\\VinEco\\Mastah Project\\{0}",
                                  "ChiaHang OpenConfig.xlsb");
                string conStr = string.Format(Constants.Excel07ConString, filePath, "YES");

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
                                   ExportAsString      = false,
                                   FormatStrategy      = CellValueFormatStrategy.None,
                                   ExportColumnName    = true
                               };

                var dicUnit = new Dictionary<string, string>(StringComparer.Ordinal);

                #endregion

                #region UnitConversion

                Aspose.Cells.Worksheet xlWs = xlWb.Worksheets["UnitConversion"];

                var dt = new DataTable();
                dt     = xlWs.Cells.ExportDataTable(0, 0, xlWs.Cells.MaxDataRow + 1, xlWs.Cells.MaxDataColumn + 1, opts);

                //OleDbConnection oleCon = new OleDbConnection(conStr);

                //string connectionString = "Select * From [" + xlWs.Name.ToString() + "$" + xlRng.Offset[0, 0].Address[RowAbsolute: false, ColumnAbsolute: false, ReferenceStyle: Excel.XlReferenceStyle.xlA1, External: xlRng] + "]";
                //OleDbDataAdapter _oleAdapt = new OleDbDataAdapter(connectionString, oleCon);
                //_oleAdapt.Fill(dt);

                //oleCon.Close();

                foreach (DataRow dr in dt.Rows)
                {
                    Product _Product = Product.FirstOrDefault(x => x.ProductCode == dr["VECode"].ToString());
                    if (_Product                                                 == null)
                    {
                        // To be fucking honest, this should NEVER be hit.
                        // Unit Converstion definition for a product that's NOT EVEN EXIST.
                        // ... and of fucking course IT IS HIT.
                    }
                    else
                    {
                        ProductUnit _ProductUnit = ProductUnitList
                           .FirstOrDefault(x =>
                                               x.ProductCode == dr["VECode"].ToString());

                        string _Region = dr["Region"].ToString();
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
                                                                 _id           = Guid.NewGuid(),
                                                                 Region        = _Region,
                                                                 OrderUnitType = ProperUnit(dr["OrderUnitType"].ToString(), dicUnit),
                                                                 OrderUnitPer  =
                                                                     ProperUnit(dr["OrderUnitType"].ToString(), dicUnit) == "Kg"
                                                                         ? 1
                                                                         : (double) dr["OderUnitPer"],
                                                                 SaleUnitType = ProperUnit(dr["SaleUnitType"].ToString(), dicUnit),
                                                                 SaleUnitPer  =
                                                                     ProperUnit(dr["SaleUnitType"].ToString(), dicUnit) == "Kg"
                                                                         ? 1
                                                                         : (double) dr["SaleUnitPer"]
                                                             };

                                var _ListRegion = new List<ProductUnitRegion>();
                                _ListRegion.Add(_ProductUnitRegion);

                                _ProductUnit.ListRegion = _ListRegion;
                            }
                            else
                            {
                                ProductUnitRegion _ProductUnitRegion =
                                    _ProductUnit.ListRegion.FirstOrDefault(x => x.Region == _Region);

                                if (_ProductUnitRegion == null)
                                {
                                    _ProductUnitRegion = new ProductUnitRegion
                                                             {
                                                                 _id           = Guid.NewGuid(),
                                                                 Region        = _Region,
                                                                 OrderUnitType = ProperUnit(dr["OrderUnitType"].ToString(), dicUnit),
                                                                 OrderUnitPer  =
                                                                     ProperUnit(dr["OrderUnitType"].ToString(), dicUnit) == "Kg"
                                                                         ? 1
                                                                         : (double) dr["OrderUnitPer"],
                                                                 SaleUnitType = ProperUnit(dr["SaleUnitType"].ToString(), dicUnit),
                                                                 SaleUnitPer  =
                                                                     ProperUnit(dr["SaleUnitType"].ToString(), dicUnit) == "Kg"
                                                                         ? 1
                                                                         : (double) dr["SaleUnitPer"]
                                                             };
                                    _ProductUnit.ListRegion.Add(_ProductUnitRegion);
                                }
                                else
                                {
                                    _ProductUnitRegion = new ProductUnitRegion
                                                             {
                                                                 _id           = Guid.NewGuid(),
                                                                 Region        = _Region,
                                                                 OrderUnitType = ProperUnit(dr["OrderUnitType"].ToString(), dicUnit),
                                                                 OrderUnitPer  =
                                                                     ProperUnit(dr["OrderUnitType"].ToString(), dicUnit) == "Kg"
                                                                         ? 1
                                                                         : (double) dr["OrderUnitPer"],
                                                                 SaleUnitType = ProperUnit(dr["SaleUnitType"].ToString(), dicUnit),
                                                                 SaleUnitPer  =
                                                                     ProperUnit(dr["SaleUnitType"].ToString(), dicUnit) == "Kg"
                                                                         ? 1
                                                                         : (double) dr["SaleUnitPer"]
                                                             };
                                }
                            }
                        }
                        else
                        {
                            _ProductUnit = new ProductUnit
                                               {
                                                   ProductCode = dr["VECode"].ToString(),
                                                   ProductId   = Product.FirstOrDefault(x => x.ProductCode == dr["VECode"].ToString())
                                                                        .ProductId,
                                                   ListRegion = new List<ProductUnitRegion>()
                                               };

                            _ProductUnit.ListRegion.Add(new ProductUnitRegion
                                                            {
                                                                _id           = Guid.NewGuid(),
                                                                Region        = _Region,
                                                                OrderUnitType = ProperUnit(dr["OrderUnitType"].ToString(), dicUnit),
                                                                OrderUnitPer  =
                                                                    ProperUnit(dr["OrderUnitType"].ToString(), dicUnit) == "Kg"
                                                                        ? 1
                                                                        : (double) dr["OrderUnitPer"],
                                                                SaleUnitType = ProperUnit(dr["SaleUnitType"].ToString(), dicUnit),
                                                                SaleUnitPer  = ProperUnit(dr["SaleUnitType"].ToString(), dicUnit) == "Kg"
                                                                                   ? 1
                                                                                   : (double) dr["SaleUnitPer"]
                                                            });

                            ProductUnitList.Add(_ProductUnit);
                        }
                    }
                }

                await db.DropCollectionAsync("ProductUnit");
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
                    Product _Product = Product.Where(x => x.ProductCode == dr["Code"].ToString()).FirstOrDefault();
                    if (_Product                                        != null)
                    {
                        var _ProductCrossRegion = new ProductCrossRegion
                                                      {
                                                          _id       = _Product._id,
                                                          ProductId = _Product.ProductId,
                                                          ToNorth   = dr["ToNorth"].ToString() == "Yes" ? true : false,
                                                          ToSouth   = dr["ToSouth"].ToString() == "Yes" ? true : false
                                                      };
                        ListProductRegion.Add(_ProductCrossRegion);
                    }
                }

                await db.DropCollectionAsync("ProductCrossRegion");
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
                    Product _Product = Product.Where(x => x.ProductCode == dr["Code"].ToString()).FirstOrDefault();
                    if (_Product                                        != null)
                    {
                        var _ProductRate = new ProductRate
                                               {
                                                   _id         = _Product._id,
                                                   ProductId   = _Product.ProductId,
                                                   ProductCode = _Product.ProductCode,
                                                   ToNorth     = Convert.ToDouble(dr["ToNorth"] ?? 1),
                                                   ToSouth     = Convert.ToDouble(dr["ToSouth"] ?? 1)
                                               };
                        ListProductRate.Add(_ProductRate);
                    }
                }

                await db.DropCollectionAsync("ProductRate");
                await db.GetCollection<ProductRate>("ProductRate").InsertManyAsync(ListProductRate);

                WriteToRichTextBoxOutput(string.Format("{0} done!", dt.TableName));

                #endregion

                #region ProductClass

                var dicClass = new Dictionary<string, string>
                                   {
                                       { "A", "Rau ăn lá" },
                                       { "B", "Rau ăn thân hoa" },
                                       { "C", "Rau ăn quả " },
                                       { "D", "Rau ăn củ" },
                                       { "E", "Cây ăn hạt" },
                                       { "F", "Rau gia vị " },
                                       { "G", "Thủy canh" },
                                       { "H", "Rau mầm " },
                                       { "I", "Nấm" },
                                       { "J", "Lá " },
                                       { "K", "Trái cây (Quả)" },
                                       { "L", "Gạo" },
                                       { "M", "Cỏ và cây công trình" },
                                       { "N", "Hoa" },
                                       { "O", "Dược liệu" }
                                   };

                foreach (Product _Product in Product)
                {
                    // Freaking amazingly gloriously typo.
                    _Product.ProductName = ProperStr(_Product.ProductName);

                    if (dicClass.TryGetValue(_Product.ProductCode.Substring(0, 1), out string _ProductClassification))
                    {
                        if (_Product.ProductCode == "K01901" || _Product.ProductCode == "K02201")
                        {
                            _ProductClassification = dicClass["F"];
                        }

                        _Product.ProductClassification = _ProductClassification;
                    }
                }

                await db.DropCollectionAsync("Product");
                await db.GetCollection<Product>("Product").InsertManyAsync(Product);

                WriteToRichTextBoxOutput("Update ProductClassification - Done!");

                #endregion

                #region ExtraProductionInformation

                xlWs = xlWb.Worksheets["ExtraProductInformation"];

                dt = new DataTable { TableName = "ExtraProductInformation Table" };
                dt = xlWs.Cells.ExportDataTable(0, 0, xlWs.Cells.MaxDataRow + 1, xlWs.Cells.MaxDataColumn + 1, opts);

                foreach (DataRow dr in dt.Rows)
                {
                    Product _Product = Product.Where(x => x.ProductCode == dr["Code"].ToString()).FirstOrDefault();

                    if (_Product != null)
                    {
                        _Product.ProductOrientation = dr["ProductionOrientation"].ToString();
                        _Product.ProductClimate     = dr["ProductClimate"].ToString();
                        _Product.ProductionGroup    = dr["ProductionGroup"].ToString();
                    }
                }

                await db.DropCollectionAsync("Product");
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
                    Product _Product = Product.Where(x => x.ProductCode == dr["Code"].ToString()).FirstOrDefault();
                    if (_Product                                        != null)
                    {
                        if ((string) dr["North"] == "Yes")
                        {
                            _Product.ProductNote.Add("North");
                        }

                        if ((string) dr["South"] == "Yes")
                        {
                            _Product.ProductNote.Add("South");
                        }
                    }
                }

                await db.DropCollectionAsync("Product");
                await db.GetCollection<Product>("Product").InsertManyAsync(Product);

                WriteToRichTextBoxOutput(string.Format("Update {0} - Done!", xlWs.Name));

                #endregion

                #region Priority

                xlWs = xlWb.Worksheets["Priority"];

                dt = new DataTable { TableName = "Priority Table" };
                dt = xlWs.Cells.ExportDataTable(0, 0, xlWs.Cells.MaxDataRow + 1, xlWs.Cells.MaxDataColumn + 1, opts);

                List<Customer> ListCustomer = db.GetCollection
                    <Customer>("Customer").AsQueryable().ToList();

                var dicPriority = new Dictionary<string, bool>();

                foreach (DataRow dr in dt.Rows)
                {
                    dicPriority.Add(dr["CCODE"].ToString(), true);
                }

                foreach (Customer _Customer in ListCustomer.Where(Customer =>
                                                                      Customer.CustomerCode == "VM"  ||
                                                                      Customer.CustomerCode == "VM+" ||
                                                                      Customer.CustomerCode == "VM+ VinEco"))
                {
                    // Cleaning stuff
                    _Customer.CustomerType = _Customer.CustomerType.Replace("Priority", "").Trim();

                    // Ooooh, I heard that you didn't get enough vegetables.
                    if (dicPriority.ContainsKey(_Customer.CustomerCode))
                    {
                        _Customer.CustomerType += " Priority";
                    }

                    _Customer.CustomerRegion = ProperStr(_Customer.CustomerRegion);
                }

                await db.DropCollectionAsync("Customer");
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
        private void EatForecast(List<ForecastDate>          FC, Range                                                                        xlRng, Worksheet xlWs, string conStr,
                                 string                      SupplierType, Dictionary<DateTime, Dictionary<string, Dictionary<string, Guid>>> dicFC,
                                 Dictionary<string, Product> dicProduct, Dictionary<string, Supplier>                                         dicSupplier, List<Product> Product,
                                 List<Supplier>              Supplier, bool                                                                   YesNoKPI = false)
        {
            try
            {
                var rowIndex = 0;
                if ((xlRng.Cells[1, 1].value != "Region") & (xlRng.Cells[1, 1].value != "Vùng"))
                {
                    do
                    {
                        rowIndex++;
                        if (rowIndex >= xlRng.Rows.Count)
                        {
                            return;
                        }
                    } while ((xlRng.Cells[rowIndex + 1, 1].Value != "Region") &
                             (xlRng.Cells[rowIndex + 1, 1].Value != "Vùng"));
                }

                var dt = new DataTable();

                var oleCon = new OleDbConnection(conStr);

                var _oleAdapt = new OleDbDataAdapter(
                    "Select * From [" +
                    xlWs.Name         +
                    "$"               +
                    xlRng.Offset[rowIndex, 0]
                         .Address[false, false, XlReferenceStyle.xlA1,
                                  xlRng] +
                    "]", oleCon);
                string _str = xlRng.Offset[rowIndex, 0].Address;
                WriteToRichTextBoxOutput(_str);
                _oleAdapt.Fill(dt);

                oleCon.Close();

                // To deal with the uhm, Templates having different Headers.
                // Please shoot me.
                foreach (DataColumn dc in dt.Columns)
                {
                    if (dt.Columns.Contains("Vùng"))
                    {
                        dt.Columns["Vùng"].ColumnName = "Region";
                    }

                    if (dt.Columns.Contains("Mã Farm"))
                    {
                        dt.Columns["Mã Farm"].ColumnName = "SCODE";
                    }

                    if (dt.Columns.Contains("Tên Farm"))
                    {
                        dt.Columns["Tên Farm"].ColumnName = "SNAME";
                    }

                    if (dt.Columns.Contains("Nhóm"))
                    {
                        dt.Columns["Nhóm"].ColumnName = "PCLASS";
                    }

                    if (dt.Columns.Contains("Mã VECrops"))
                    {
                        dt.Columns["Mã VECrops"].ColumnName = "VECrops Code";
                    }

                    if (dt.Columns.Contains("Mã VinEco"))
                    {
                        dt.Columns["Mã VinEco"].ColumnName = "PCODE";
                    }

                    if (dt.Columns.Contains("Tên VinEco"))
                    {
                        dt.Columns["Tên VinEco"].ColumnName = "PNAME";
                    }
                }

                // Main Loop
                foreach (DataColumn dc in dt.Columns)
                {
                    DateTime dateValue;

                    // Loop for every column. Only stop at column with Date at the top ( Indicating it being PurchaseOrder for that Date )
                    if (DateTime.TryParse(dc.ColumnName, out dateValue))
                    {
                        ForecastDate _FC     = null;
                        var          isNewFC = false;

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

                            _FC                = new ForecastDate();
                            _FC._id            = Guid.NewGuid();
                            _FC.ForecastDateId = _FC._id;

                            _FC.DateForecast        = dateValue.Date;
                            _FC.ListProductForecast = new List<ProductForecast>();

                            dicFC.Add(dateValue.Date, new Dictionary<string, Dictionary<string, Guid>>());
                        }

                        // First layer
                        // Get the list of all Products being Ordered that day.
                        List<ProductForecast> _listProductForecast = _FC.ListProductForecast;
                        if (_listProductForecast == null)
                        {
                            _listProductForecast = new List<ProductForecast>();
                        }

                        // Loop for every value
                        foreach (DataRow dr in dt.Rows)
                        {
                            // In case of empty SCODE. I really hate to deal with this case. Like, really.
                            if (dr["SCODE"] == null || string.IsNullOrEmpty(dr["SCODE"].ToString()))
                            {
                                dr["SCODE"] = dr["SNAME"]; // Oh for god's sake.
                            }

                            // If OrderQuantity is not 0 - Not Anymore?
                            //object _OrderQuantity = dr[dc.ColumnName];
                            if (
                                dr["PCODE"] !=
                                DBNull.Value /*&& dr[dc.ColumnName] != DBNull.Value*/ /*&& Convert.ToDouble(dr[dc.ColumnName]) > 0*/ &&
                                (SupplierType == "ThuMua" ? dr["SCODE"] != DBNull.Value : true))
                            {
                                // Olala
                                List<SupplierForecast> _ListSupplierForecast = null;
                                SupplierForecast       _SupplierForecast     = null;
                                ProductForecast        _ProductForecast      = null;
                                // Olala2
                                var isNewProductOrder  = false;
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

                                            _product._id         = Guid.NewGuid();
                                            _product.ProductId   = _product._id;
                                            _product.ProductCode = dr["PCODE"].ToString();
                                            _product.ProductName = dr["PNAME"].ToString();
                                            //_product.ProductClassification = dr["PCLASS"].ToString();
                                            //_product.ProductVECode = dt.Columns.Contains("VECrops Code")
                                            //                             ? dr["VECrops Code"].ToString()
                                            //                             : "";

                                            Product.Add(_product);

                                            dicProduct.Add(dr["PCODE"].ToString(), _product);
                                        }
                                    }

                                    _ProductForecast = _FC.ListProductForecast
                                                          .Where(x => x.ProductId == _product.ProductId)
                                                          .FirstOrDefault();

                                    Guid _id;
                                    if (dicStore.TryGetValue(dr["SCODE"].ToString(), out _id))
                                    {
                                        _SupplierForecast = _ProductForecast.ListSupplierForecast
                                                                            .Where(x => x.SupplierId == _id)
                                                                            .FirstOrDefault();
                                    }
                                    else
                                    {
                                        isNewCustomerOrder = true;

                                        Supplier _supplier = null;
                                        if (!dicSupplier.TryGetValue(dr["SCODE"].ToString(), out _supplier))
                                        {
                                            _supplier = dicSupplier.Values
                                                                   .Where(x => x.SupplierCode == dr["SCODE"].ToString())
                                                                   .FirstOrDefault();
                                            if (_supplier == null)
                                            {
                                                _supplier = new Supplier();

                                                _supplier._id          = Guid.NewGuid();
                                                _supplier.SupplierId   = _supplier._id;
                                                _supplier.SupplierCode =
                                                    dr["SCODE"]
                                                       .ToString(); //SupplierType == "ThuMua" ? (dr["SCODE"] == null ? dr["SCODE"].ToString() : dr["SNAME"].ToString()) : dr["SNAME"].ToString();
                                                _supplier.SupplierName = dr["SNAME"].ToString();
                                                _supplier.SupplierType =
                                                    SupplierType == "ThuMua" && dr["Tag"].ToString() == "VCM"
                                                        ? "VCM"
                                                        : SupplierType;

                                                string _region = dr["Region"].ToString();
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
                                                _supplier.SupplierType   =
                                                    SupplierType == "ThuMua" && dr["Tag"].ToString() == "VCM"
                                                        ? "VCM"
                                                        : SupplierType;

                                                Supplier.Add(_supplier);
                                                dicSupplier.Add(_supplier.SupplierCode, _supplier);
                                            }
                                        }

                                        _SupplierForecast                    = new SupplierForecast();
                                        _SupplierForecast._id                = Guid.NewGuid();
                                        _SupplierForecast.SupplierForecastId = _SupplierForecast._id;
                                        _SupplierForecast.SupplierId         = _supplier.SupplierId;

                                        dicFC[dateValue.Date][dr["PCODE"].ToString()]
                                           .Add(dr["SCODE"].ToString(), _supplier.SupplierId);
                                    }
                                }
                                else
                                {
                                    isNewProductOrder  = true;
                                    isNewCustomerOrder = true;

                                    Product _product = null;
                                    if (!dicProduct.TryGetValue(dr["PCODE"].ToString(), out _product))
                                    {
                                        _product = dicProduct.Values.Where(x => x.ProductCode == dr["PCODE"].ToString())
                                                             .FirstOrDefault();
                                        if (_product == null)
                                        {
                                            _product = new Product();

                                            _product._id         = Guid.NewGuid();
                                            _product.ProductId   = _product._id;
                                            _product.ProductCode = dr["PCODE"].ToString();
                                            _product.ProductName = dr["PNAME"].ToString();
                                            //_product.ProductClassification = dr["PCLASS"].ToString();
                                            //_product.ProductVECode = dt.Columns.Contains("VECrops Code")
                                            //                             ? dr["VECrops Code"].ToString()
                                            //                             : "";

                                            Product.Add(_product);

                                            dicProduct.Add(dr["PCODE"].ToString(), _product);
                                        }
                                    }

                                    _ProductForecast                   = new ProductForecast();
                                    _ProductForecast._id               = Guid.NewGuid();
                                    _ProductForecast.ProductForecastId = _ProductForecast._id;
                                    _ProductForecast.ProductId         = _product.ProductId;

                                    _ProductForecast.ListSupplierForecast = new List<SupplierForecast>();

                                    Supplier _supplier = null;
                                    if (!dicSupplier.TryGetValue(dr["SCODE"].ToString(), out _supplier))
                                    {
                                        _supplier = dicSupplier.Values
                                                               .Where(x => x.SupplierCode == dr["SCODE"].ToString())
                                                               .FirstOrDefault();
                                        if (_supplier == null)
                                        {
                                            _supplier = new Supplier();

                                            _supplier._id          = Guid.NewGuid();
                                            _supplier.SupplierId   = _supplier._id;
                                            _supplier.SupplierCode =
                                                dr["SCODE"]
                                                   .ToString(); //SupplierType == "ThuMua" ? (dr["SCODE"] == null ? dr["SCODE"].ToString() : dr["SNAME"].ToString()) : dr["SNAME"].ToString();
                                            _supplier.SupplierName = dr["SNAME"].ToString();
                                            _supplier.SupplierType =
                                                SupplierType == "ThuMua" && dr["Tag"].ToString() == "VCM"
                                                    ? "VCM"
                                                    : SupplierType;

                                            string _region = dr["Region"].ToString();
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
                                            _supplier.SupplierType   =
                                                SupplierType == "ThuMua" && dr["Tag"].ToString() == "VCM"
                                                    ? "VCM"
                                                    : SupplierType;

                                            Supplier.Add(_supplier);
                                            dicSupplier.Add(_supplier.SupplierCode, _supplier);
                                        }
                                    }

                                    _SupplierForecast                    = new SupplierForecast();
                                    _SupplierForecast._id                = Guid.NewGuid();
                                    _SupplierForecast.SupplierForecastId = _SupplierForecast._id;
                                    _SupplierForecast.SupplierId         = _supplier.SupplierId;

                                    dicFC[dateValue.Date].Add(dr["PCODE"].ToString(), new Dictionary<string, Guid>());
                                    dicFC[dateValue.Date][dr["PCODE"].ToString()]
                                       .Add(dr["SCODE"].ToString(), _supplier.SupplierId);
                                }

                                // Filling in data
                                _ListSupplierForecast = _ProductForecast.ListSupplierForecast;

                                // Special part for ThuMua
                                TextInfo myTI = new CultureInfo("en-US", false).TextInfo;
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
                                    _SupplierForecast.LabelVinEco        = true;
                                    _SupplierForecast.FullOrder          = false;
                                    _SupplierForecast.CrossRegion        = false;
                                    _SupplierForecast.Level              = 1;
                                    _SupplierForecast.Availability       = "1234567";

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

                                if (SupplierType                           == "VinEco" &&
                                    dr["PCODE"].ToString().Substring(0, 1) == "K"      &&
                                    (dr["Region"].ToString()               == "MN" || dr["Region"].ToString() == "Miền Nam")
                                ) //dicCrossRegionVinEco.ContainsKey(dr["PCODE"].ToString()))
                                {
                                    _SupplierForecast.CrossRegion = true;
                                    if (dr["PCODE"].ToString() == "K03501")
                                    {
                                        _SupplierForecast.CrossRegion = false;
                                    }
                                }

                                ///// < !For debugging purposes !>
                                //if (!YesNoKPI && dateValue.Day == 16 && (string)dr["PCODE"] == "C02801" && (string)dr["SCODE"] == "AG03030000")
                                //{
                                //    byte AmIHandsome = 0;
                                //}

                                // 3rd FC layer - Normal Forecast.
                                if (double.TryParse(
                                    (dr[dc.ColumnName] == DBNull.Value ? 0 : dr[dc.ColumnName]).ToString(),
                                    out double _QuantityForecast))
                                {
                                    if (!YesNoKPI)
                                    {
                                        _SupplierForecast.QuantityForecast += _QuantityForecast;
                                    }

                                    // 2nd FC layer - Minimum / Contracted Forecast - 2nd Highest Priority. 
                                    if (dt.Columns.Contains("Min"))
                                    {
                                        if (double.TryParse((dr["Min"] == DBNull.Value ? 0 : dr["Min"]).ToString(),
                                                            out double _QuantityForecastContracted))
                                        {
                                            _SupplierForecast.QuantityForecastContracted += _QuantityForecastContracted;
                                        }
                                    }
                                }

                                if (YesNoKPI &&
                                    Convert.ToDateTime(dr["EffectiveFrom"]).Date <=
                                    DateTime.Parse(dc.ColumnName).Date &&
                                    Convert.ToDateTime(dr["EffectiveTo"]).Date >=
                                    DateTime.Parse(dc.ColumnName).Date)
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

                                if (isNewCustomerOrder)
                                {
                                    _ListSupplierForecast.Add(_SupplierForecast);
                                }

                                _ProductForecast.ListSupplierForecast = _ListSupplierForecast;
                                if (isNewProductOrder)
                                {
                                    _FC.ListProductForecast.Add(_ProductForecast);
                                }
                            }
                        }

                        _FC.ListProductForecast = _listProductForecast;

                        if (isNewFC)
                        {
                            FC.Add(_FC);
                        }
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
        private void EatPO(List<PurchaseOrderDate>     PO, Range                                                                    xlRng, Worksheet xlWs, string conStr,
                           string                      PORegion, Dictionary<DateTime, Dictionary<string, Dictionary<string, Guid>>> dicPO,
                           Dictionary<string, Product> dicProduct, Dictionary<string, Customer>                                     dicCustomer, List<Product> Product,
                           List<Customer>              Customer, bool                                                               YesNoNew = false)
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
                        "Select * From ["                                         +
                        xlWs.Name                                                 +
                        "$"                                                       +
                        xlRng.Address[false, false, XlReferenceStyle.xlA1, xlRng] +
                        "]", oleCon);
                string _str = xlRng.Offset[rowIndex, 0].Address;
                WriteToRichTextBoxOutput(_str);
                _oleAdapt.Fill(dt);

                oleCon.Close();

                _oleAdapt = null;
                oleCon    = null;

                var                                 mongoClient = new MongoClient();
                IMongoCollection<PurchaseOrderDate> db          = mongoClient.GetDatabase("localtest")
                                                                             .GetCollection<PurchaseOrderDate>("PurchaseOrder");

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
                        {
                            PO.RemoveAll(x => x.DateOrder.Date == dateValue);
                        }

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

                            _PODate                     = new PurchaseOrderDate();
                            _PODate._id                 = Guid.NewGuid();
                            _PODate.PurchaseOrderDateId = _PODate._id;

                            _PODate.DateOrder        = dateValue.Date;
                            _PODate.ListProductOrder = new List<ProductOrder>();

                            dicPO.Add(dateValue.Date, new Dictionary<string, Dictionary<string, Guid>>());
                        }

                        // First layer
                        // Get the list of all Products being Ordered that day.
                        List<ProductOrder> _listProductOrder = _PODate.ListProductOrder;
                        if (_listProductOrder == null)
                        {
                            _listProductOrder = new List<ProductOrder>();
                        }

                        // Loop for every value
                        foreach (DataRow dr in dt.Rows)
                        {
                            // If OrderQuantity is not 0 - Not Anymore?
                            //object _OrderQuantity = dr[dc.ColumnName];
                            double _value = 0;
                            if (dr["VE Code"]              != DBNull.Value &&
                                dr[dt.Columns.IndexOf(dc)] != DBNull.Value &&
                                double.TryParse(dr[dt.Columns.IndexOf(dc)].ToString(), out _value)
                            ) //&& Convert.ToDouble(dr[dc.ColumnName]) > 0)
                            {
                                if (_value > 0)
                                {
                                    List<CustomerOrder> _listCustomerOrder = null;
                                    CustomerOrder       _CustomerOrder     = null;
                                    ProductOrder        _productOrder      = null;

                                    var isNewProductOrder  = false;
                                    var isNewCustomerOrder = false;

                                    Dictionary<string, Guid> dicStore = null;

                                    _dicProduct = dicPO[dateValue.Date];
                                    if (_dicProduct.TryGetValue(dr["VE Code"].ToString(), out dicStore))
                                    {
                                        Product _product = null;
                                        if (!dicProduct.TryGetValue(dr["VE Code"].ToString(), out _product))
                                        {
                                            _product = new Product();

                                            _product._id         = Guid.NewGuid();
                                            _product.ProductId   = _product._id;
                                            _product.ProductCode = dr["VE Code"].ToString();
                                            _product.ProductName = dr["VE Name"].ToString();

                                            Product.Add(_product);

                                            dicProduct.Add(_product.ProductCode, _product);
                                        }

                                        _productOrder = _PODate.ListProductOrder
                                                               .Where(x => x.ProductId == _product.ProductId)
                                                               .FirstOrDefault();

                                        Guid _id;
                                        if (dicStore.TryGetValue(
                                            dr["StoreCode"] +
                                            (dt.Columns.Contains("P&L")
                                                 ? dr["P&L"].ToString()
                                                 : dr["StoreType"].ToString()), out _id))
                                        {
                                            _CustomerOrder = _productOrder.ListCustomerOrder
                                                                          .Where(x => x.CustomerId == _id)
                                                                          .FirstOrDefault();
                                        }
                                        else
                                        {
                                            isNewCustomerOrder = true;

                                            Customer _customer;
                                            string   sKey = dr["StoreCode"] +
                                                            (dt.Columns.Contains("P&L")
                                                                 ? dr["P&L"].ToString()
                                                                 : dr["StoreType"].ToString());
                                            if (!dicCustomer.TryGetValue(sKey, out _customer))
                                            {
                                                _customer = new Customer();

                                                _customer._id            = Guid.NewGuid();
                                                _customer.CustomerId     = _customer._id;
                                                _customer.CustomerCode   = dr["StoreCode"].ToString();
                                                _customer.CustomerName   = dr["StoreName"].ToString();
                                                _customer.CustomerRegion = dr["Region"].ToString();
                                                _customer.CustomerType   = dr["StoreType"].ToString();
                                                _customer.Company        = dt.Columns.Contains("P&L")
                                                                               ? dr["P&L"].ToString()
                                                                               : "VinCommerce";
                                                _customer.CustomerBigRegion = PORegion;

                                                Customer.Add(_customer);

                                                dicCustomer.Add(sKey, _customer);
                                            }

                                            Guid _NewGuid  = Guid.NewGuid();
                                            _CustomerOrder = new CustomerOrder
                                                                 {
                                                                     _id             = _NewGuid,
                                                                     //CustomerOrderId = _NewGuid,
                                                                     CustomerId      = _customer.CustomerId
                                                                 };

                                            dicPO[dateValue.Date][dr["VE Code"].ToString()]
                                               .Add(sKey, _customer.CustomerId);
                                        }
                                    }
                                    else
                                    {
                                        isNewProductOrder  = true;
                                        isNewCustomerOrder = true;

                                        Product _product = null;
                                        if (!dicProduct.TryGetValue(dr["VE Code"].ToString(), out _product))
                                        {
                                            _product = new Product
                                                           {
                                                               _id         = Guid.NewGuid(),
                                                               ProductCode = dr["VE Code"].ToString(),
                                                               ProductName = dr["VE Name"].ToString()
                                                           };

                                            _product.ProductId = _product._id;


                                            Product.Add(_product);

                                            dicProduct.Add(_product.ProductCode, _product);
                                        }

                                        _productOrder                = new ProductOrder();
                                        _productOrder._id            = Guid.NewGuid();
                                        //_productOrder.ProductOrderId = _productOrder._id;
                                        _productOrder.ProductId      = _product.ProductId;

                                        _productOrder.ListCustomerOrder = new List<CustomerOrder>();

                                        Customer _customer;
                                        string   sKey = dr["StoreCode"] +
                                                        (dt.Columns.Contains("P&L")
                                                             ? dr["P&L"].ToString()
                                                             : dr["StoreType"].ToString());
                                        if (!dicCustomer.TryGetValue(sKey, out _customer))
                                        {
                                            _customer = new Customer();

                                            _customer._id            = Guid.NewGuid();
                                            _customer.CustomerId     = _customer._id;
                                            _customer.CustomerCode   = dr["StoreCode"].ToString();
                                            _customer.CustomerName   = dr["StoreName"].ToString();
                                            _customer.CustomerRegion = dr["Region"].ToString();
                                            _customer.CustomerType   = dr["StoreType"].ToString();
                                            _customer.Company        = dt.Columns.Contains("P&L")
                                                                           ? dr["P&L"].ToString()
                                                                           : "VinCommerce";
                                            _customer.CustomerBigRegion = PORegion;

                                            Customer.Add(_customer);

                                            dicCustomer.Add(sKey, _customer);
                                        }

                                        _CustomerOrder                 = new CustomerOrder();
                                        _CustomerOrder._id             = Guid.NewGuid();
                                        //_CustomerOrder.CustomerOrderId = _CustomerOrder._id;
                                        _CustomerOrder.CustomerId      = _customer.CustomerId;

                                        dicPO[dateValue.Date]
                                           .Add(dr["VE Code"].ToString(),
                                                new Dictionary<string, Guid>());
                                        dicPO[dateValue.Date][dr["VE Code"].ToString()].Add(sKey, _customer.CustomerId);
                                    }

                                    // Filling in data
                                    _listCustomerOrder = _productOrder.ListCustomerOrder;

                                    // Desired Region
                                    if (dt.Columns.Contains("Vùng sản xuất") && dr["Vùng sản xuất"] != null)
                                    {
                                        string _DesiredRegion = dr["Vùng sản xuất"].ToString();

                                        if (_DesiredRegion  != "" &&
                                            (_DesiredRegion == "Lâm Đồng" ||
                                             _DesiredRegion == "Miền Bắc" ||
                                             _DesiredRegion == "Miền Nam"))
                                        {
                                            //_CustomerOrder.DesiredRegion = _DesiredRegion;
                                        }
                                    }

                                    // Desired Source
                                    if (dt.Columns.Contains("Nguồn") && dr["Nguồn"] != null)
                                    {
                                        string _DesiredSource = dr["Nguồn"].ToString();

                                        if (_DesiredSource  != "" &&
                                            (_DesiredSource == "VinEco" ||
                                             _DesiredSource == "ThuMua" ||
                                             _DesiredSource == "VCM"))
                                        {
                                            //_CustomerOrder.DesiredSource = _DesiredSource;
                                        }
                                    }

                                    _CustomerOrder.Unit =
                                        ProperUnit(dr["Unit"].ToString() == "" ? "Kg" : dr["Unit"].ToString(), dicUnit);
                                    _CustomerOrder.QuantityOrder += _value;

                                    if (isNewCustomerOrder)
                                    {
                                        _listCustomerOrder.Add(_CustomerOrder);
                                    }

                                    _productOrder.ListCustomerOrder = _listCustomerOrder;
                                    if (isNewProductOrder)
                                    {
                                        _PODate.ListProductOrder.Add(_productOrder);
                                    }
                                }
                            }
                        }

                        _PODate.ListProductOrder = _listProductOrder;

                        if (isNewPODate)
                        {
                            PO.Add(_PODate);
                        }

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
        private void EatPOAspose(List<PurchaseOrderDate>     PO, Aspose.Cells.Worksheet                                                   xlWs, string conStr,
                                 string                      PORegion, Dictionary<DateTime, Dictionary<string, Dictionary<string, Guid>>> dicPO,
                                 Dictionary<string, Product> dicProduct, Dictionary<string, Customer>                                     dicCustomer, List<Product> Product,
                                 List<Customer>              Customer, bool                                                               YesNoNew = false)
        {
            try
            {
                Stopwatch stopwatch = Stopwatch.StartNew();

                Debug.WriteLine($"File: {xlWs.Workbook.FileName}");

                var dicUnit = new Dictionary<string, string>(StringComparer.Ordinal);
                // Find first row.
                var    rowIndex = 0;
                var    colIndex = 0;
                string value    = string.Empty;
                do
                {
                    value = xlWs.Cells[rowIndex, colIndex].Value?.ToString().Trim();

                    if (value == "VE Code" || value == "Mã Planning" || value == "Mã Planing")
                    {
                        break;
                    }

                    //if (value == null || value == string.Empty || (value != "VE Code" && value != "Mã Planning"))
                    //    rowIndex++;
                    //else
                    //    break;

                    rowIndex++;

                    if (rowIndex > 100)
                    {
                        colIndex++;
                        if (colIndex > 100)
                        {
                            break;
                        }

                        rowIndex = 0;
                    }
                } while ((value   == null || value == string.Empty || value != "VE Code" && value != "Mã Planning") &&
                         rowIndex <= 100                                                 &&
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
                                   ExportAsString      = false,
                                   FormatStrategy      = CellValueFormatStrategy.None,
                                   ExportColumnName    = true
                               };

                var dt = new DataTable { TableName = xlWs.Name };
                dt     = xlWs.Cells.ExportDataTable(rowIndex, colIndex, xlWs.Cells.MaxDataRow + 1,
                                                    xlWs.Cells.MaxDataColumn                  + 1,
                                                    opts);

                //var mongoClient = new MongoClient();
                //var db = mongoClient.GetDatabase("localtest").GetCollection<PurchaseOrderDate>("PurchaseOrder");

                if (!dt.Columns.Contains("VE Code"))
                {
                    dt.Columns[colIndex].ColumnName = "VE Code";
                }

                if (dt.Columns.Contains("Tên mới"))
                {
                    dt.Columns["Tên mới"].ColumnName = "VE Name";
                }

                if (dt.Columns.Contains("Tỉnh tiêu thụ"))
                {
                    dt.Columns["Tỉnh tiêu thụ"].ColumnName = "Region";
                }

                if (dt.Columns.Contains("Store Code"))
                {
                    dt.Columns["Store Code"].ColumnName = "StoreCode";
                }

                if (dt.Columns.Contains("Store Name"))
                {
                    dt.Columns["Store Name"].ColumnName = "StoreName";
                }

                if (dt.Columns.Contains("Store Type"))
                {
                    dt.Columns["Store Type"].ColumnName = "StoreType";
                }

                // Main Loop
                foreach (DataColumn dc in dt.Columns)
                    // Loop for every column. Only stop at column with Date at the top ( Indicating it being PurchaseOrder for that Date )
                    //if (DateTime.TryParse(dc.ColumnName,
                    //    out dateValue) /* && (dateValue.Date >= DateTime.Today.AddDays(0).Date)*/)
                {
                    if (StringToDate(dc.ColumnName) != null)
                    {
                        DateTime dateValue = StringToDate(dc.ColumnName) ?? DateTime.MinValue;

                        PurchaseOrderDate _PODate = null;
                        if (YesNoNew)
                        {
                            PO.RemoveAll(x => x.DateOrder.Date == dateValue);
                        }

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

                            _PODate                     = new PurchaseOrderDate();
                            _PODate._id                 = Guid.NewGuid();
                            _PODate.PurchaseOrderDateId = _PODate._id;

                            _PODate.DateOrder        = dateValue.Date;
                            _PODate.ListProductOrder = new List<ProductOrder>();

                            dicPO.Add(dateValue.Date, new Dictionary<string, Dictionary<string, Guid>>());
                        }

                        // First layer
                        // Get the list of all Products being Ordered that day.
                        List<ProductOrder> _listProductOrder = _PODate.ListProductOrder;
                        if (_listProductOrder == null)
                        {
                            _listProductOrder = new List<ProductOrder>();
                        }

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
                            if (dr["VE Code"]              != DBNull.Value &&
                                dr[dt.Columns.IndexOf(dc)] != DBNull.Value &&
                                double.TryParse(dr[dt.Columns.IndexOf(dc)].ToString(), out _value)
                            ) //&& Convert.ToDouble(dr[dc.ColumnName]) > 0)
                            {
                                if (_value > 0)
                                {
                                    List<CustomerOrder> _listCustomerOrder = null;
                                    CustomerOrder       _CustomerOrder     = null;
                                    ProductOrder        _productOrder      = null;

                                    var isNewProductOrder  = false;
                                    var isNewCustomerOrder = false;

                                    _dicProduct = dicPO[dateValue.Date];
                                    if (_dicProduct.TryGetValue(dr["VE Code"].ToString(),
                                                                out Dictionary<string, Guid> dicStore))
                                    {
                                        if (!dicProduct.TryGetValue(dr["VE Code"].ToString(), out Product _product))
                                        {
                                            _product = new Product
                                                           {
                                                               _id         = Guid.NewGuid(),
                                                               ProductCode = dr["VE Code"].ToString(),
                                                               ProductName = dr["VE Name"].ToString()
                                                           };

                                            _product.ProductId = _product._id;

                                            Product.Add(_product);

                                            dicProduct.Add(_product.ProductCode, _product);
                                        }

                                        _productOrder = _PODate.ListProductOrder
                                                               .Where(x => x.ProductId == _product.ProductId)
                                                               .FirstOrDefault();

                                        Guid _id;
                                        if (dicStore.TryGetValue(
                                            dr["StoreCode"] +
                                            (dt.Columns.Contains("P&L")
                                                 ? dr["P&L"].ToString()
                                                 : dr["StoreType"].ToString()), out _id))
                                        {
                                            _CustomerOrder = _productOrder.ListCustomerOrder
                                                                          .Where(x => x.CustomerId == _id)
                                                                          .FirstOrDefault();
                                        }
                                        else
                                        {
                                            isNewCustomerOrder = true;

                                            Customer _customer;
                                            string   sKey = dr["StoreCode"] +
                                                            (dt.Columns.Contains("P&L")
                                                                 ? dr["P&L"].ToString()
                                                                 : dr["StoreType"].ToString());
                                            if (!dicCustomer.TryGetValue(sKey, out _customer))
                                            {
                                                _customer = new Customer();

                                                _customer._id            = Guid.NewGuid();
                                                _customer.CustomerId     = _customer._id;
                                                _customer.CustomerCode   = dr["StoreCode"].ToString();
                                                _customer.CustomerName   = dr["StoreName"].ToString();
                                                _customer.CustomerRegion = dr["Region"].ToString();
                                                _customer.CustomerType   = dr["StoreType"].ToString();
                                                _customer.Company        = dt.Columns.Contains("P&L")
                                                                               ? dr["P&L"].ToString()
                                                                               : "VinCommerce";
                                                _customer.CustomerBigRegion = PORegion;

                                                Customer.Add(_customer);

                                                dicCustomer.Add(sKey, _customer);
                                            }

                                            Guid _NewGuid  = Guid.NewGuid();
                                            _CustomerOrder = new CustomerOrder
                                                                 {
                                                                     _id             = _NewGuid,
                                                                     //CustomerOrderId = _NewGuid,
                                                                     CustomerId      = _customer.CustomerId
                                                                 };

                                            dicPO[dateValue.Date][dr["VE Code"].ToString()]
                                               .Add(sKey, _customer.CustomerId);
                                        }
                                    }
                                    else
                                    {
                                        isNewProductOrder  = true;
                                        isNewCustomerOrder = true;

                                        Product _product = null;
                                        if (!dicProduct.TryGetValue(dr["VE Code"].ToString(), out _product))
                                        {
                                            _product = new Product();

                                            _product._id         = Guid.NewGuid();
                                            _product.ProductId   = _product._id;
                                            _product.ProductCode = dr["VE Code"].ToString();
                                            _product.ProductName = dr["VE Name"].ToString();

                                            Product.Add(_product);

                                            dicProduct.Add(_product.ProductCode, _product);
                                        }

                                        _productOrder = new ProductOrder
                                        {
                                            _id = Guid.NewGuid(),
                                            ProductId = _product.ProductId,
                                            ListCustomerOrder = new List<CustomerOrder>()
                                        };

                                        //_productOrder.ProductOrderId = _productOrder._id;


                                        string sKey = dr["StoreCode"] +
                                                        (dt.Columns.Contains("P&L")
                                                             ? dr["P&L"].ToString()
                                                             : dr["StoreType"].ToString());
                                        if (!dicCustomer.TryGetValue(sKey, out Customer _customer))
                                        {
                                            _customer = new Customer
                                            {
                                                _id = Guid.NewGuid(),
                                                CustomerCode = dr["StoreCode"].ToString(),
                                                CustomerName = dr["StoreName"].ToString(),
                                                CustomerRegion = dr["Region"].ToString(),
                                                CustomerType = dr["StoreType"].ToString(),
                                                Company = dt.Columns.Contains("P&L")
                                                                           ? dr["P&L"].ToString()
                                                                           : "VinCommerce",
                                                CustomerBigRegion = PORegion
                                            };

                                            _customer.CustomerId     = _customer._id;
                                           
                                            Customer.Add(_customer);

                                            dicCustomer.Add(sKey, _customer);
                                        }

                                        _CustomerOrder                 = new CustomerOrder { _id = Guid.NewGuid() };
                                        //_CustomerOrder.CustomerOrderId = _CustomerOrder._id;
                                        _CustomerOrder.CustomerId      = _customer.CustomerId;

                                        dicPO[dateValue.Date]
                                           .Add(dr["VE Code"].ToString(),
                                                new Dictionary<string, Guid>());
                                        dicPO[dateValue.Date][dr["VE Code"].ToString()].Add(sKey, _customer.CustomerId);
                                    }

                                    // Filling in data
                                    _listCustomerOrder = _productOrder.ListCustomerOrder;

                                    // Desired Region
                                    if (dt.Columns.Contains("Vùng sản xuất") && dr["Vùng sản xuất"] != null)
                                    {
                                        string _DesiredRegion = dr["Vùng sản xuất"].ToString();

                                        if (_DesiredRegion  != "" &&
                                            (_DesiredRegion == "Lâm Đồng" ||
                                             _DesiredRegion == "Miền Bắc" ||
                                             _DesiredRegion == "Miền Nam"))
                                        {
                                            //_CustomerOrder.DesiredRegion = _DesiredRegion;
                                        }
                                    }

                                    // Desired Source
                                    if (dt.Columns.Contains("Nguồn") && dr["Nguồn"] != null)
                                    {
                                        string _DesiredSource = dr["Nguồn"].ToString();

                                        if (_DesiredSource  != "" &&
                                            (_DesiredSource == "VinEco" ||
                                             _DesiredSource == "ThuMua" ||
                                             _DesiredSource == "VCM"))
                                        {
                                            //_CustomerOrder.DesiredSource = _DesiredSource;
                                        }
                                    }

                                    _CustomerOrder.Unit =
                                        ProperUnit(dr["Unit"].ToString() == "" ? "Kg" : dr["Unit"].ToString(), dicUnit);
                                    _CustomerOrder.QuantityOrder += _value;

                                    if (isNewCustomerOrder)
                                    {
                                        _listCustomerOrder.Add(_CustomerOrder);
                                    }

                                    _productOrder.ListCustomerOrder = _listCustomerOrder;
                                    if (isNewProductOrder)
                                    {
                                        _PODate.ListProductOrder.Add(_productOrder);
                                    }
                                }
                            }
                        }

                        _PODate.ListProductOrder = _listProductOrder;

                        if (isNewPODate)
                        {
                            PO.Add(_PODate);
                        }
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
                IMongoDatabase db = new MongoClient().GetDatabase("localtest");

                // Initialize stuff.
                var           ListAllo = new List<AllocateDetail>();
                List<Product> Product  = db.GetCollection<Product>("Product").AsQueryable().ToList();
                var           Customer = new List<Customer>();

                // Core!
                var core = new CoordStructure();

                // ... and of course, core stuff.
                var dicPO        = new Dictionary<DateTime, Dictionary<string, Dictionary<string, Guid>>>();
                core.dicProduct  = new Dictionary<Guid, Product>();
                core.dicCustomer = new Dictionary<Guid, Customer>();

                // Product Dictionary.
                foreach (Product _Product in Product)
                {
                    if (!core.dicProduct.ContainsKey(_Product.ProductId))
                    {
                        core.dicProduct.Add(_Product.ProductId, _Product);
                    }
                }

                // Customer Dictionary.
                foreach (Customer _Customer in Customer)
                {
                    if (!core.dicCustomer.ContainsKey(_Customer.CustomerId))
                    {
                        core.dicCustomer.Add(_Customer.CustomerId, _Customer);
                    }
                }

                // Directory.
                // Todo - Hardcoded, need to change.
                var directoryPath =
                    "D:\\Documents\\Stuff\\VinEco\\Mastah Project\\Deli";

                #region Reading PO files in folder.

                var        dirInfo  = new DirectoryInfo(directoryPath);
                FileInfo[] ListFile = dirInfo.GetFiles();

                foreach (FileInfo _FileInfo in ListFile)
                {
                    var xlWb = new Workbook(_FileInfo.FullName);

                    EatDeli(xlWb, core);
                }

                #endregion

                WriteToRichTextBoxOutput("Here goes pain");

                await db.DropCollectionAsync("AllocateDetail");
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
                Aspose.Cells.Worksheet xlWs = xlWb.Worksheets.OrderByDescending(x => x.Cells.MaxDataRow + 1)
                                                  .FirstOrDefault();

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
                dt     = xlWs.Cells.ExportDataTable(rowIndex, 0, xlWs.Cells.MaxDataRow + 1, xlWs.Cells.MaxDataColumn + 1,
                                                    true);

                // Here we go.
                // Dissecting DataTable into Database.
                foreach (DataRow dr in dt.Rows)
                {
                    // Harvest Date.
                    DateTime DateProcess = Convert.ToDateTime(dr["DATE_PROCESS"]);

                    // Order Date.
                    DateTime DateOrder = Convert.ToDateTime(dr["DATE_ORDER"]);

                    // Product.
                    string  _ProductCode = dr["PCODE1"].ToString().Substring(0, 6);
                    Product _Product     = core.dicProduct.Values.Where(x => x.ProductCode == _ProductCode)
                                               .FirstOrDefault();

                    // Customer.
                    var      _CustomerCode = (string) dr["CCODE"];
                    Customer _Customer     = core.dicCustomer.Values.Where(x => x.CustomerCode == _CustomerCode)
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
            TextInfo myTI = new CultureInfo("en-US", false).TextInfo;

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
            if (dicUnit.TryGetValue(Unit, out string unit))
            {
                return unit;
            }

            // Initialize empty result.
            string _Unit = Unit.Trim().ToLower();

            // Looping through every letter.
            for (var stringIndex = 0; stringIndex < _Unit.Length; stringIndex++)
                // If a forward dash is found.
            {
                if (_Unit.Substring(stringIndex, 1) == "/")
                {
                    // Insert a space if the letter right before it isn't already a space.
                    if (stringIndex != 0 && _Unit.Substring(stringIndex - 1, 1) != " ")
                    {
                        _Unit = _Unit.Insert(stringIndex - 1, " ");
                    }

                    // Insert a space if the letter right after it isn't already a space.
                    if (stringIndex != _Unit.Length && _Unit.Substring(stringIndex + 1, 1) != " ")
                    {
                        _Unit = _Unit.Insert(stringIndex + 1, " ");
                    }
                }
            }

            // Creates a TextInfo based on the "en-US" culture.
            TextInfo myTI = new CultureInfo("en-US", false).TextInfo;

            dicUnit.Add(Unit, myTI.ToTitleCase(_Unit));

            // Return the "Proper" Unit.
            return myTI.ToTitleCase(_Unit);
        }

        #endregion

        private void WriteToRichTextBoxOutput(object Message = null, bool NewLine = true)
        {
            try
            {
                if (Message == null)
                {
                    Message = "";
                }

                richTextBoxOutput.AppendText($"{Message},{(NewLine ? "\n" : " ")}");
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

        private          DateTime DateFrom    = DateTime.Today;
        private          DateTime DateTo      = DateTime.Today;
        private          double   UpperCap    = 1.2;
        private readonly byte     dayDistance = 4;

        //private readonly byte dayCrossRegion = 4;
        private          bool FruitOnly;
        private          bool NoFruit;
        private readonly bool YesPlanningFuckMe = false;
        private          bool YesNoSubRegion;

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
                    double _UpperCap = UpperCap;
                    if (double.TryParse(upperCapBox.Text, out _UpperCap))
                    {
                        UpperCap = _UpperCap;
                    }
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
            await UpdateFcAsync("DBSL.xlsb", "ThuMua.xlsb");
        }

        private async void readFCPlanningToolStripMenuItem_Click(object sender, EventArgs e)
        {
            await UpdateFcAsync("DBSL.xlsb", "ThuMua Planning.xlsb", true);
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

            FiteMoi(
                DateFrom: DateFrom,
                DateTo: DateTo > DateFrom ? DateTo : DateFrom,
                YesNoCompact: false,
                YesNoNoSup: true,
                YesNoLimit: true,
                YesNoGroupFarm: false,
                YesNoGroupThuMua: false,
                YesNoReportM1: false,
                YesNoByUnit: false);
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
            Stopwatch stopwatch = Stopwatch.StartNew();

            string fileName = $"PO {$"{DateFrom:dd.MM} - {DateTo:dd.MM} ({DateTime.Now:yyyyMMdd HH\\hmm})"}.xlsx";
            string path     = $@"D:\Documents\Stuff\VinEco\Mastah Project\Test\{fileName}";

            //LargeExport(path);

            stopwatch.Stop();
            WriteToRichTextBoxOutput($"Done in {Math.Round(stopwatch.Elapsed.TotalSeconds, 2)}s!");
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
            var            mongoClient = new MongoClient();
            IMongoDatabase db          = mongoClient.GetDatabase("localtest");
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
        private void OutputExcel(DataTable dt, string               sheetName, Microsoft.Office.Interop.Excel.Workbook xlWb,
                                 bool      YesNoHeader = false, int RowFirst = 6, bool                                 YesNoFirstSheet = false)
        {
            try
            {
                // Open Second Workbook
                //string filePath = string.Format("D:\\Documents\\Stuff\\VinEco\\Mastah Project\\{0}", fileName);
                //var xlWb2 = new Aspose.Cells.Workbook(filePath);
                //var xlWs2 = xlWb2.Worksheets[0];

                int rowTotal = dt.Rows.Count;
                int colTotal = dt.Columns.Count;

                if (rowTotal == 0 || colTotal == 0)
                {
                    return;
                }

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
                    xlWs      = xlWb.Worksheets[0];
                    xlWs.Name = sheetName;
                }
                else
                {
                    xlWs = xlWb.Worksheets[sheetName]; //xlWb2.Worksheets.Count];
                }

                Range rangeToDelete = xlWs.get_Range("A" + RowFirst, (Range) xlWs.Cells[rowTotal, colTotal]);
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
                    {
                        Header[i] = dt.Columns[i].ColumnName;
                    }

                    Range HeaderRange =
                        xlWs.get_Range((Range) xlWs.Cells[RowFirst, 1], (Range) xlWs.Cells[1, colTotal]);
                    HeaderRange.Value          = Header;
                    HeaderRange.Interior.Color = ColorTranslator.ToOle(Color.LightGray);
                    HeaderRange.Font.Bold      = true;
                }

                #endregion

                int _RowFirst = YesNoHeader ? RowFirst + 1 : RowFirst;

                // Limiting the size of object. If this is too large, expect Out of Memory Exception.
                // Interop pls.
                // Apparently larger yields worse performance. Idk why.
                // Too small also negatively impacts performance. Oh god why.

                //var _rowPerBlock = Math.Round(rowTotal / 17, 0);
                //int rowPerBlock = (int)Math.Round(rowTotal / 17d, 0); // 7777;
                var rowPerBlock = 7777;
                //int rowPerBlock = (int)Math.Max(Math.Round(rowTotal / 17d, 0), 7777); // 7777;
                //WriteToRichTextBoxOutput(rowPerBlock);

                var dbCells  = new object[rowPerBlock, colTotal];
                var count    = 0;
                var rowPos   = 0;
                var rowIndex = 0;
                //byte[] cellCheck = new byte[] { 17, 20, 23, 26, 30, 34 };
                foreach (DataRow dr in dt.Rows)
                {
                    // Hardcoding for more efficiency.
                    // Currently this is too slow.
                    for (var colIndex = 0; colIndex < colTotal; colIndex++)
                    {
                        //if (dt.Rows[rowIndex][colIndex] == null) { continue; }

                        string _value = (dr[colIndex] ?? "").ToString();
                        Type   _type  = dt.Columns[colIndex].DataType;
                        if (_value != "" && _value != "0")
                        {
                            if (_type == typeof(DateTime))
                            {
                                dbCells[rowIndex - rowPos, colIndex] = dr[colIndex];
                            }
                            else if (_type == typeof(double))
                            {
                                dbCells[rowIndex - rowPos, colIndex] = Convert.ToDouble(_value);
                            }
                            else
                            {
                                dbCells[rowIndex - rowPos, colIndex] = _value;
                            }
                        }
                    }

                    count++;
                    if (count >= rowPerBlock)
                    {
                        xlWs.get_Range((Range) xlWs.Cells[rowPos + _RowFirst, 1],
                                       (Range) xlWs.Cells[rowPos + rowPerBlock + _RowFirst - 1, colTotal])
                            .Formula = dbCells;
                        //xlWs2.Range[rowPos + _RowFirst, 1].Resize[rowPos + rowPerBlock + _RowFirst - 1, colTotal].Value = dbCells;
                        dbCells = new object[Math.Min(rowTotal - rowPos, rowPerBlock), colTotal];
                        count   = 0;
                        rowPos  = rowIndex + 1;
                    }

                    rowIndex++;
                }

                xlWs.get_Range((Range) xlWs.Cells[Math.Max(rowPos + _RowFirst, 2), 1],
                               (Range) xlWs.Cells[rowPos          + rowPerBlock + _RowFirst - 1, colTotal])
                    .Formula = dbCells;
                //xlWs2.Range["A" + RowFirst].get_Resize(dbCells.Length[0], dbCells.Length(1)).Value = dbCells;

                dbCells = null;

                if (xlWs != null)
                {
                    Marshal.ReleaseComObject(xlWs);
                }

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
                                       int       RowFirst                                                     = 6, string Position   = "A1", Dictionary<string, int> DicColDate = null,
                                       string    CustomDateFormat                                             = "", bool  AutoFilter = false)
        {
            try
            {
                Style defaultStyle = xlWb.CreateStyle();

                defaultStyle.Font.Name = "Calibri";
                defaultStyle.Font.Size = 11;

                xlWb.DefaultStyle = defaultStyle;

                defaultStyle = null;

                int rowTotal = dataTable.Rows.Count;
                int colTotal = dataTable.Columns.Count;

                //foreach (Aspose.Cells.Worksheet _xlWs in xlWb.Worksheets)
                //{
                //    WriteToRichTextBoxOutput(_xlWs.Name);
                //}

                Aspose.Cells.Worksheet xlWs = xlWb.Worksheets[sheetName];

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
                {
                    xlWs.AutoFilter.Range =
                        $"A1:{xlWs.Cells[xlWs.Cells.MaxDataRow + 1, xlWs.Cells.MaxDataColumn + 1].Name}";
                }

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
        public static void LargeExport(DataTable dt, string                filename, Dictionary<string, int> DicDateCol,
                                       bool      YesNoHeader = false, bool YesNoZero = false, bool           YesNoDateColumn = false)
        {
            try
            {
                using (SpreadsheetDocument document =
                    SpreadsheetDocument.Create(filename, SpreadsheetDocumentType.Workbook))
                {
                    var dicType = new Dictionary<Type, CellValues>();

                    var dicColName = new Dictionary<int, string>();

                    for (var colIndex = 0; colIndex < dt.Columns.Count; colIndex++)
                    {
                        dicColName.Add(colIndex + 1, GetColumnName(colIndex + 1));
                    }

                    dicType.Add(typeof(DateTime), CellValues.Date);
                    dicType.Add(typeof(string), CellValues.InlineString);
                    dicType.Add(typeof(double), CellValues.Number);
                    dicType.Add(typeof(int), CellValues.Number);
                    dicType.Add(typeof(bool), CellValues.Boolean);

                    //this list of attributes will be used when writing a start element
                    List<OpenXmlAttribute> attributes;
                    OpenXmlWriter          writer;

                    document.AddWorkbookPart();
                    var workSheetPart = document.WorkbookPart.AddNewPart<WorksheetPart>();

                    // Add Stylesheet.
                    var WorkbookStylesPart        = document.WorkbookPart.AddNewPart<WorkbookStylesPart>();
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
                            Type type = dt.Columns[columnNum - 1].DataType;
                            //reset the list of attributes
                            //attributes = new List<OpenXmlAttribute>();
                            // add data type attribute - in this case inline string (you might want to look at the shared strings table)
                            //attributes.Add(new OpenXmlAttribute("t", null, "str")); //(type == typeof(string) ? "str" : (YesNoDateColumn == false ? "str" : dicType[typeof(DateTime)].ToString()))));

                            //add the cell reference attribute
                            //attributes.Add(new OpenXmlAttribute("r", "", string.Format("{0}{1}", dicColName[columnNum], 1)));

                            //write the cell start element with the type and reference attributes
                            //writer.WriteStartElement(new DocumentFormat.OpenXml.Spreadsheet.Cell(), attributes);

                            DateTime _value;
                            var      _dateValue = 0;
                            if (DateTime.TryParse(dt.Columns[columnNum - 1].ColumnName, out _value))
                            {
                                _dateValue = (int) (_value.Date - new DateTime(1900, 1, 1)).TotalDays + 2;
                            }

                            //write the cell value
                            var cell = new Cell
                                           {
                                               DataType = type == typeof(double) && _dateValue != 0
                                                              ? CellValues.Number
                                                              : CellValues.String,
                                               CellReference = $"{dicColName[columnNum]}1",
                                               CellValue     = new CellValue(type == typeof(double) && _dateValue != 0
                                                                                 ? _dateValue.ToString()
                                                                                 : dt.Columns[columnNum - 1].ColumnName),
                                               StyleIndex = (uint) (type == typeof(double) && _dateValue != 0 ? 1 : 0)
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

                        DataRow dr = dt.Rows[rowNum - 1];
                        for (var columnNum = 1; columnNum <= dt.Columns.Count; columnNum++)
                        {
                            string colName = dt.Columns[columnNum - 1].ColumnName;
                            Type   type    = dt.Columns[columnNum - 1].DataType;
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
                                                        CellReference = $"{dicColName[columnNum]}{(YesNoHeader ? rowNum + 1 : rowNum)}",
                                                        CellValue  = new CellValue(dr[columnNum            - 1].ToString()),
                                                        StyleIndex = (uint) (DicDateCol.ContainsKey(colName) ? 1 : 0)
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
                                                Name    = dt.TableName == "" ? "Whatever" : dt.TableName,
                                                SheetId = 1,
                                                Id      = document.WorkbookPart.GetIdOfPart(workSheetPart)
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
                var bold  = new Bold();
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
                                          FormatCode     = StringValue.FromString("dd-MMM")
                                      };
                workbookstylesheet.NumberingFormats = new NumberingFormats();
                workbookstylesheet.NumberingFormats.Append(nf2DateTime);

                // <CellFormats>
                var cellformat0 = new CellFormat
                                      {
                                          FontId   = 0,
                                          FillId   = 0,
                                          BorderId = 0
                                      }; // Default style : Mandatory | Style ID =0

                var cellformat1 = new CellFormat
                                      {
                                          BorderId          = 0,
                                          FillId            = 0,
                                          FontId            = 0,
                                          NumberFormatId    = 7170,
                                          FormatId          = 0,
                                          ApplyNumberFormat = true
                                      };

                var cellformat2 = new CellFormat
                                      {
                                          BorderId          = 0,
                                          FillId            = 0,
                                          FontId            = 0,
                                          NumberFormatId    = 14,
                                          FormatId          = 0,
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
        /// <param name="filePath">Where your file will be.</param>
        /// <param name="listDataTables">List of dataTables. Can contain just 1, doesn't matter.</param>
        /// <param name="yesHeader">You want headers?</param>
        /// <param name="yesZero">You want zero instead of null?</param>
        public void LargeExportOneWorkbook(
            string                 filePath,
            IEnumerable<DataTable> listDataTables,
            bool                   yesHeader = false,
            bool                   yesZero   = false)
        {
            try
            {
                using (SpreadsheetDocument document = SpreadsheetDocument.Create(
                    filePath,
                    SpreadsheetDocumentType.Workbook))
                {
                    document.AddWorkbookPart();

                    OpenXmlWriter writerXb = OpenXmlWriter.Create(document.WorkbookPart);
                    writerXb.WriteStartElement(new DocumentFormat.OpenXml.Spreadsheet.Workbook());
                    writerXb.WriteStartElement(new Sheets());

                    var count = 0;

                    foreach (DataTable dt in listDataTables)
                    {
                        var dicColName = new Dictionary<int, string>();

                        for (var colIndex = 0; colIndex < dt.Columns.Count; colIndex++)
                        {
                            int    dividend   = colIndex + 1;
                            string columnName = string.Empty;

                            while (dividend > 0)
                            {
                                int modifier = (dividend - 1) % 26;
                                columnName   =
                                    $"{Convert.ToChar(65 + modifier).ToString(CultureInfo.InvariantCulture)}{columnName}";
                                dividend = (dividend     - modifier) / 26;
                            }

                            dicColName.Add(colIndex + 1, columnName);
                        }

                        // var dicType = new Dictionary<Type, CellValues>(4)
                        //                  {
                        //                      { typeof(DateTime), CellValues.Date },
                        //                      { typeof(string), CellValues.InlineString },
                        //                      { typeof(double), CellValues.Number },
                        //                      { typeof(int), CellValues.Number },
                        //                      { typeof(bool), CellValues.Boolean }
                        //                  };
                        var dicType = new Dictionary<Type, string>(4)
                                          {
                                              { typeof(DateTime), "d" },
                                              { typeof(string), "s" },
                                              { typeof(double), "n" },
                                              { typeof(int), "n" },
                                              { typeof(bool), "b" }
                                          };

                        // this list of attributes will be used when writing a start element
                        List<OpenXmlAttribute> attributes;

                        var workSheetPart = document.WorkbookPart.AddNewPart<WorksheetPart>();

                        OpenXmlWriter writer = OpenXmlWriter.Create(workSheetPart);
                        writer.WriteStartElement(new DocumentFormat.OpenXml.Spreadsheet.Worksheet());
                        writer.WriteStartElement(new SheetData());

                        if (yesHeader)
                        {
                            // create a new list of attributes
                            attributes = new List<OpenXmlAttribute>
                                             {
                                                 // add the row index attribute to the list
                                                 new OpenXmlAttribute("r", null, 1.ToString())
                                             };

                            // write the row start element with the row index attribute
                            writer.WriteStartElement(new Row(), attributes);

                            for (var columnNum = 1; columnNum <= dt.Columns.Count; ++columnNum)
                            {
                                // reset the list of attributes
                                attributes = new List<OpenXmlAttribute>
                                                 {
                                                     new OpenXmlAttribute("t", null, "str"),
                                                     new OpenXmlAttribute("r", string.Empty, $"{dicColName[columnNum]}1")
                                                 };

                                // add data type attribute - in this case inline string (you might want to look at the shared strings table)
                                // add the cell reference attribute

                                // write the cell start element with the type and reference attributes
                                writer.WriteStartElement(new Cell(), attributes);

                                // write the cell value
                                writer.WriteElement(new CellValue(dt.Columns[columnNum - 1].ColumnName));

                                // writer.WriteElement(new CellValue(string.Format("This is Row {0}, Cell {1}", rowNum, columnNum)));

                                // write the end cell element
                                writer.WriteEndElement();
                            }

                            // write the end row element
                            writer.WriteEndElement();
                        }

                        for (var rowNum = 1; rowNum <= dt.Rows.Count; rowNum++)
                        {
                            // create a new list of attributes
                            attributes = new List<OpenXmlAttribute> { new OpenXmlAttribute("r", null, (yesHeader ? rowNum + 1 : rowNum).ToString()) };

                            // add the row index attribute to the list

                            // write the row start element with the row index attribute
                            writer.WriteStartElement(new Row(), attributes);

                            DataRow dr = dt.Rows[rowNum - 1];
                            for (var columnNum = 1; columnNum <= dt.Columns.Count; columnNum++)
                            {
                                Type type = dt.Columns[columnNum - 1].DataType;

                                // reset the list of attributes
                                attributes = new List<OpenXmlAttribute>
                                                 {
                                                     // Add data type attribute - in this case inline string (you might want to look at the shared strings table)
                                                     new OpenXmlAttribute("t", null, type == typeof(string) ? "str" : dicType[type]),

                                                     // Add the cell reference attribute
                                                     new OpenXmlAttribute("r", string.Empty, $"{dicColName[columnNum]}{(yesHeader ? rowNum + 1 : rowNum).ToString(CultureInfo.InvariantCulture)}")
                                                 };

                                // write the cell start element with the type and reference attributes
                                writer.WriteStartElement(new Cell(), attributes);

                                // write the cell value
                                if (yesZero | (dr[columnNum - 1].ToString() != "0"))
                                {
                                    writer.WriteElement(new CellValue(dr[columnNum - 1].ToString()));
                                }

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

                        writerXb.WriteElement(
                            new Sheet
                                {
                                    Name    = dt.TableName,
                                    SheetId = Convert.ToUInt32(count + 1),
                                    Id      = document.WorkbookPart.GetIdOfPart(workSheetPart)
                                });

                        count++;
                    }

                    // End Sheets
                    writerXb.WriteEndElement();

                    // End Workbook
                    writerXb.WriteEndElement();

                    writerXb.Close();

                    document.WorkbookPart.Workbook.Save();

                    document.Close();
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
                throw;
            }
        }

        public static void LargeExportOriginal(string filename)
        {
            using (SpreadsheetDocument document = SpreadsheetDocument.Create(filename, SpreadsheetDocumentType.Workbook)
            )
            {
                //this list of attributes will be used when writing a start element
                List<OpenXmlAttribute> attributes;
                OpenXmlWriter          writer;

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
                        attributes.Add(new OpenXmlAttribute("r", "", $"{GetColumnName(columnNum)}{rowNum}"));

                        //write the cell start element with the type and reference attributes
                        writer.WriteStartElement(new Cell(), attributes);
                        //write the cell value
                        writer.WriteElement(new CellValue($"This is Row {rowNum}, Cell {columnNum}"));

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
                                            Name    = "Large Sheet",
                                            SheetId = 1,
                                            Id      = document.WorkbookPart.GetIdOfPart(workSheetPart)
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
            int    dividend   = columnIndex;
            string columnName = string.Empty;
            int    modifier;

            while (dividend > 0)
            {
                modifier   = (dividend         - 1) % 26;
                columnName = Convert.ToChar(65 + modifier) + columnName;
                dividend   = (dividend         - modifier) / 26;
            }

            return columnName;
        }

        /// <summary>
        ///     Convert from one file format to another, using Interop.
        ///     Because apparently OpenXML doesn't deal with .xls type ( Including, but not exclusive to .xlsb )
        /// </summary>
        private void ConvertToXlsbInterop(string filePath,
                                          string PreviousExtension  = "",
                                          string AfterwardExtension = "",
                                          bool   YesNoDeleteFile    = false)
        {
            try
            {
                // Remember the list of running Excel.Application.
                // Before initialize xlApp.
                Process[] processBefore = Process.GetProcessesByName("excel");

                // Initialize new instance of Interop Excel.Application.
                var xlApp = new Application();

                // Remember the list of running Excel.Application.
                // After initialize xlApp.
                Process[] processAfter = Process.GetProcessesByName("excel");

                var processID = 0;

                // Compare two lists, get the first and the only process that's not in the 'Before' List.
                foreach (Process process in processAfter)
                {
                    if (!processBefore.Select(p => p.Id).Contains(process.Id))
                    {
                        processID = process.Id;
                        break;
                    }
                }

                xlApp.ScreenUpdating   = false;
                xlApp.EnableEvents     = false;
                xlApp.DisplayAlerts    = false;
                xlApp.DisplayStatusBar = false;
                xlApp.AskToUpdateLinks = false;

                Microsoft.Office.Interop.Excel.Workbook xlWb = xlApp.Workbooks.Open(filePath);

                object missing = Type.Missing;
                xlWb.SaveAs(filePath.Replace(PreviousExtension, AfterwardExtension), XlFileFormat.xlExcel12, missing,
                            missing, false, false, XlSaveAsAccessMode.xlExclusive, missing, missing, missing);

                xlWb.Close(false);
                Marshal.ReleaseComObject(xlWb);
                xlWb = null;

                xlApp.Quit();
                Marshal.ReleaseComObject(xlApp);
                xlApp = null;

                if (YesNoDeleteFile)
                {
                    File.Delete(filePath);
                }

                // Kill the instance of Interop Excel.Application used by this call.
                if (processID != 0)
                {
                    Process process = Process.GetProcessById(processID);
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
                Process[] processBefore = Process.GetProcessesByName("excel");

                // Initialize new instance of Interop Excel.Application.
                var xlApp = new Application();

                // Remember the list of running Excel.Application.
                // After initialize xlApp.
                Process[] processAfter = Process.GetProcessesByName("excel");

                var processID = 0;

                // Compare two lists, get the first and the only process that's not in the 'Before' List.
                foreach (Process process in processAfter)
                {
                    if (!processBefore.Select(p => p.Id).Contains(process.Id))
                    {
                        processID = process.Id;
                        break;
                    }
                }

                Microsoft.Office.Interop.Excel.Workbook xlWb = xlApp.Workbooks.Open(
                    Filename: filePath,
                    UpdateLinks: false,
                    ReadOnly: false,
                    Format: 5,
                    Password: "",
                    WriteResPassword: "",
                    IgnoreReadOnlyRecommended: true,
                    Origin: XlPlatform.xlWindows,
                    Delimiter: "",
                    Editable: true,
                    Notify: false,
                    Converter: 0,
                    AddToMru: true,
                    Local: false,
                    CorruptLoad: false);

                xlApp.ScreenUpdating   = false;
                xlApp.Calculation      = XlCalculation.xlCalculationManual;
                xlApp.EnableEvents     = false;
                xlApp.DisplayAlerts    = false;
                xlApp.DisplayStatusBar = false;
                xlApp.AskToUpdateLinks = false;

                foreach (Worksheet _ws in xlWb.Worksheets)
                {
                    if (_ws.Name == "Evaluation Warning")
                    {
                        _ws.Delete();
                    }
                }

                xlWb.Sheets[1].Activate();

                xlApp.ScreenUpdating   = true;
                xlApp.Calculation      = XlCalculation.xlCalculationAutomatic;
                xlApp.EnableEvents     = true;
                xlApp.DisplayAlerts    = false;
                xlApp.DisplayStatusBar = true;
                xlApp.AskToUpdateLinks = true;

                xlWb.Close(true);

                if (xlWb != null)
                {
                    Marshal.ReleaseComObject(xlWb);
                }

                xlWb = null;

                xlApp.Quit();
                if (xlApp != null)
                {
                    Marshal.ReleaseComObject(xlApp);
                }

                xlApp = null;

                //GC.Collect();
                //GC.WaitForPendingFinalizers();

                // Kill the instance of Interop Excel.Application used by this call.
                if (processID != 0)
                {
                    Process process = Process.GetProcessById(processID);
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
            {
                if (!ExcludedChars.Contains(((char) i).ToString()))
                {
                    text = text.Replace(((char) i).ToString(), "");
                }
            }

            for (var i = 58; i < 65; i++)
            {
                text = text.Replace(((char) i).ToString(), "");
            }

            for (var i = 91; i < 97; i++)
            {
                text = text.Replace(((char) i).ToString(), "");
            }

            for (var i = 123; i < 127; i++)
            {
                text = text.Replace(((char) i).ToString(), "");
            }

            //text = text.Replace(" ", "-"); //Comment lại để không đưa khoảng trắng thành ký tự -
            var regex = new Regex(@"\p{IsCombiningDiacriticalMarks}+");

            string strFormD = text.Normalize(NormalizationForm.FormD);

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