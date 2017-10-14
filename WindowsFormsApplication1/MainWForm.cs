using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using MongoDB.Bson.Serialization.Attributes;
using MongoDB.Driver;
using MongoDB.Driver.Linq;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

/// <summary>
/// Stuff
/// </summary>
/// 
/// Todo:
/// 
/// <!!! TOP PRIORITY !!!>
/// x Test Performance - Optimize Outputing to Excel. Interop is way too fucking slow.
/// x I gave up coz Memory Overflow
/// x Update: Ultilizing OpenXMLWriter ( SAX Method ) whenever possible. Finally able to Custom Format on it.
/// 
/// <* High Priority *>
/// x Actual Demand / Supply Function ( Remove UpperCap that's currently 100% ) ( Done May 1, 17 )
/// 
/// <* Low Priority *>
/// x Print DBSL into VE Farm or whatever ( Done May 1, 17 )
/// x Print ThuMua into VE ThuMua ( Done May 1, 17 )
/// x Fucking formula ( Done May 7, 17 )
/// x And uhm Region stuff I guess ( Done May 1, 17 )
/// 
namespace ChiaHang
{
    public partial class MainForm : Form
    {
        /// <summary>
        /// All Constants should be declared here
        /// </summary>
        private static class Constants
        {
            //public const string Excel03ConString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties='Excel 8.0;HDR={1};IMEX=1;'";
            public const string Excel07ConString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='Excel 8.0;HDR={1};IMEX=1;'";
        }

        private DateTime DateFrom = DateTime.Today;
        private DateTime DateTo = DateTime.Today;

        public MainForm()
        {
            InitializeComponent();
        }

        #region Behaviour
        private void DateFromPicker_ValueChanged(object sender, EventArgs e)
        {
            DateFrom = DateFromPicker.Value;
        }

        private void DateToPicker_ValueChanged(object sender, EventArgs e)
        {
            DateTo = DateToPicker.Value;
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

            FiteMoi(DateFrom, DateTo > DateFrom ? DateTo : DateFrom, false, true);
        }

        private void verticallyToolStripMenuItem_Click(object sender, EventArgs e)
        {
            PrintPO(false);
        }

        private void horizontallyToolStripMenuItem_Click(object sender, EventArgs e)
        {
            PrintPO(true);
        }

        private async void readFCPlanningToolStripMenuItem_Click(object sender, EventArgs e)
        {
            await UpdateFC("DBSL.xlsb", "ThuMua Planning.xlsb");
        }

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
        #endregion

        private void testExportLargeExcelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Stopwatch stopwatch = Stopwatch.StartNew();

            string fileName = string.Format("PO {0}.xlsx", DateFrom.ToString("dd.MM") + " - " + DateTo.ToString("dd.MM") + " (" + DateTime.Now.ToString("yyyyMMdd HH\\hmm") + ")");
            string path = string.Format(@"C:\Users\Shirayuki\Documents\VinEco\Project - Chia Hang\Mastah Project\Test\" + fileName);

            //LargeExport(path);

            stopwatch.Stop();
            Debug.WriteLine(String.Format("Done in {0}s!", Math.Round(stopwatch.Elapsed.TotalSeconds, 2)));
        }

        private void printToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            FiteMoi(DateFrom, DateFrom > DateTo ? DateFrom : DateTo, false, false, true, false, false, true);
        }

        private void readOpenConfigToolStripMenuItem_Click(object sender, EventArgs e)
        {
            UpdateOpenConfig();
        }
        #endregion

        /// <summary>
        /// Print Purchase Order, either horizontally ( true ) or vertically ( false )
        /// </summary>
        /// <param name="YesNoHoriziontal"></param>
        private void PrintPO(bool YesNoHoriziontal = true)
        {
            try
            {
                Debug.WriteLine("Start!");

                Stopwatch stopwatch = Stopwatch.StartNew();

                MongoClient mongoClient = new MongoClient();
                var db = mongoClient.GetDatabase("localtest");

                var PO = db.GetCollection<PurchaseOrderDate>("PurchaseOrder").Find(x =>
                        (x.DateOrder >= DateFrom.Date) &&
                        (x.DateOrder <= DateTo.Date))
                    .ToList()
                    .OrderBy(x => x.DateOrder);

                var Product = db.GetCollection<Product>("Product").AsQueryable().ToList();
                var Customer = db.GetCollection<Customer>("Customer").AsQueryable().ToList();

                var dicProduct = new Dictionary<Guid, Product>();
                var dicCustomer = new Dictionary<Guid, Customer>();

                foreach (var _Product in Product)
                {
                    dicProduct.Add(_Product.ProductId, _Product);
                }

                foreach (var _Customer in Customer)
                {
                    dicCustomer.Add(_Customer.CustomerId, _Customer);
                }

                DataTable dtPO = new DataTable();
                dtPO.TableName = String.Format("Purchase Order {0} - {1}", DateFrom.Date.ToString("dd.MM"), DateTo.Date.ToString("dd.MM"));

                var DicColDate = new Dictionary<string, int>();

                if (YesNoHoriziontal)
                {
                    dtPO.Columns.Add("PCODE", typeof(string)).DefaultValue = "";
                    dtPO.Columns.Add("PNAME", typeof(string)).DefaultValue = "";
                    dtPO.Columns.Add("PCLASS", typeof(string)).DefaultValue = "";
                    dtPO.Columns.Add("CCODE", typeof(string)).DefaultValue = "";
                    dtPO.Columns.Add("CNAME", typeof(string)).DefaultValue = "";
                    dtPO.Columns.Add("CTYPE", typeof(string)).DefaultValue = "";
                    dtPO.Columns.Add("CREGION", typeof(string)).DefaultValue = "";
                    dtPO.Columns.Add("REGION", typeof(string)).DefaultValue = "";

                    foreach (PurchaseOrderDate PODate in PO)
                    {
                        dtPO.Columns.Add(PODate.DateOrder.Date.ToString("MM/dd/yyyy"), typeof(double)).DefaultValue = 0;
                        //Debug.Write(PODate.DateOrder.Date.ToString());
                        //Debug.Write(" " + PODate.ListProductOrder.Sum(x => x.ListCustomerOrder.Sum(co => co.QuantityOrder)));
                        //Debug.Write("  " + PO.Where(wtf => wtf.DateOrder.Date.ToString() == PODate.DateOrder.Date.ToString()).FirstOrDefault().ListProductOrder.Sum(x => x.ListCustomerOrder.Sum(co => co.QuantityOrder)));
                        //Debug.WriteLine("");
                        //Debug.WriteLine("MB " + PODate.DateOrder.Date + ": " + PO.Where(x => x == PODate).FirstOrDefault().ListProductOrder.Sum(x => x.ListCustomerOrder.Where(co => dicCustomer[co.CustomerId].CustomerBigRegion == "Miền Bắc").Sum(o => o.QuantityOrder)));
                        //Debug.WriteLine("MN " + PODate.DateOrder.Date + ": " + PO.Where(x => x == PODate).FirstOrDefault().ListProductOrder.Sum(x => x.ListCustomerOrder.Where(co => dicCustomer[co.CustomerId].CustomerBigRegion == "Miền Nam").Sum(o => o.QuantityOrder)));
                    }

                    var dicRow = new Dictionary<string, int>();
                    foreach (PurchaseOrderDate PODate in PO)
                    {
                        foreach (ProductOrder _ProductOrder in PODate.ListProductOrder)
                        {
                            foreach (CustomerOrder _CustomerOrder in _ProductOrder.ListCustomerOrder)
                            {
                                var _Product = dicProduct[_ProductOrder.ProductId];
                                var _Customer = dicCustomer[_CustomerOrder.CustomerId];

                                DataRow dr = null;

                                int _rowIndex = 0;
                                string sKey = _Product.ProductCode + _Customer.CustomerCode;

                                if (!dicRow.TryGetValue(sKey, out _rowIndex))
                                {
                                    dr = dtPO.NewRow();
                                    dicRow.Add(sKey, dtPO.Rows.Count);

                                    dr["PCODE"] = _Product.ProductCode;
                                    dr["PNAME"] = _Product.ProductName;
                                    dr["CCODE"] = _Customer.CustomerCode;
                                    dr["CNAME"] = _Customer.CustomerName;
                                    dr["CTYPE"] = _Customer.CustomerType;
                                    dr["CREGION"] = _Customer.CustomerRegion;
                                    dr["REGION"] = _Customer.CustomerBigRegion;
                                    dr[PODate.DateOrder.Date.ToString("MM/dd/yyyy")] = _CustomerOrder.QuantityOrder;

                                    dtPO.Rows.Add(dr);
                                }
                                else
                                {
                                    dr = dtPO.Rows[_rowIndex];
                                    dr[PODate.DateOrder.Date.ToString("MM/dd/yyyy")] = Convert.ToDouble(dr[PODate.DateOrder.Date.ToString("MM/dd/yyyy")]) + _CustomerOrder.QuantityOrder;
                                }
                            }
                        }
                    }
                }
                else
                {
                    dtPO.Columns.Add("PCODE", typeof(string)).DefaultValue = "";
                    dtPO.Columns.Add("PNAME", typeof(string)).DefaultValue = "";
                    dtPO.Columns.Add("PCLASS", typeof(string)).DefaultValue = "";
                    dtPO.Columns.Add("CCODE", typeof(string)).DefaultValue = "";
                    dtPO.Columns.Add("CNAME", typeof(string)).DefaultValue = "";
                    dtPO.Columns.Add("CTYPE", typeof(string)).DefaultValue = "";
                    dtPO.Columns.Add("CREGION", typeof(string)).DefaultValue = "";
                    dtPO.Columns.Add("REGION", typeof(string)).DefaultValue = "";
                    dtPO.Columns.Add("DateOrder", typeof(int)).DefaultValue = 0;
                    dtPO.Columns.Add("QuantityOrder", typeof(double)).DefaultValue = 0;

                    DicColDate.Add("DateOrder", dtPO.Columns.IndexOf("DateOrder"));

                    var dicRow = new Dictionary<string, int>();
                    foreach (PurchaseOrderDate PODate in PO)
                    {
                        foreach (ProductOrder _ProductOrder in PODate.ListProductOrder)
                        {
                            foreach (CustomerOrder _CustomerOrder in _ProductOrder.ListCustomerOrder)
                            {
                                var _Product = dicProduct[_ProductOrder.ProductId];
                                var _Customer = dicCustomer[_CustomerOrder.CustomerId];

                                DataRow dr = null;

                                int _rowIndex = 0;
                                string sKey = _Product.ProductCode + _Customer.CustomerCode + PODate.DateOrder.Date.ToString("yyyyMMdd");

                                if (!dicRow.TryGetValue(sKey, out _rowIndex))
                                {
                                    dr = dtPO.NewRow();
                                    dicRow.Add(sKey, dtPO.Rows.Count);

                                    dr["PCODE"] = _Product.ProductCode;
                                    dr["PNAME"] = _Product.ProductName;
                                    dr["CCODE"] = _Customer.CustomerCode;
                                    dr["CNAME"] = _Customer.CustomerName;
                                    dr["CTYPE"] = _Customer.CustomerType;
                                    dr["CREGION"] = _Customer.CustomerRegion;
                                    dr["REGION"] = _Customer.CustomerBigRegion;
                                    dr["DateOrder"] = (int)(PODate.DateOrder.Date - new DateTime(1900, 1, 1)).TotalDays + 2;
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
                    }
                }
                //Excel.Application xlApp = new Excel.Application();
                //Aspose.Cells.Workbook xlWb = new Aspose.Cells.Workbook();

                string fileName = string.Format("PO {0}.xlsx", DateFrom.ToString("dd.MM") + " - " + DateTo.ToString("dd.MM") + " (" + DateTime.Now.ToString("yyyyMMdd HH\\hmm") + ")");
                string path = string.Format(@"C:\Users\Shirayuki\Documents\VinEco\Project - Chia Hang\Mastah Project\Test\" + fileName);
                Debug.WriteLine(path);
                //var missing = Type.Missing;
                //xlWb.SaveAs(path, Excel.XlFileFormat.xlOpenXMLWorkbook, missing, missing, false, false, Excel.XlSaveAsAccessMode.xlNoChange, missing, missing, missing);

                // Winner in term of Pure speed.
                LargeExport(dtPO, path, DicColDate, true, true);

                ConvertToXlsbInterop(path, ".xlsx", ".xlsb", true);

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
                //using (Aspose.Cells.Workbook xlWb = new Aspose.Cells.Workbook())
                //{
                //    //OutputExcel(POTable, "Sheet1", xlWb, true, 1);
                //    OutputExcelAspose(POTable, "Sheet1", xlWb, true, 1);

                //    xlWb.Save(path);
                //    //xlWb.Close(SaveChanges: true);

                //    //GC.Collect();
                //    //GC.WaitForPendingFinalizers();

                //    //if (xlWb != null) { Marshal.ReleaseComObject(xlWb); }

                //    //xlApp.Quit();
                //    //if (xlApp != null) { Marshal.ReleaseComObject(xlApp); }
                //    //xlApp = null;

                //    ////GC.Collect();
                //    ////GC.WaitForPendingFinalizers();
                //}
                //Delete_Evaluation_Sheet_Interop(path);
                #endregion

                stopwatch.Stop();
                Debug.WriteLine(String.Format("Done in {0}s!", Math.Round(stopwatch.Elapsed.TotalSeconds, 2)));
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        /// The Main Character. <para /> 
        /// Never die. <para />
        /// Ever shine.
        /// </summary>
        /// <param name="DateFrom"></param>
        /// <param name="DateTo"></param>
        /// <param name="YesNoCompact"></param>
        /// <param name="YesNoNoSup"></param>
        /// <param name="YesNoLimit"></param>
        private void FiteMoi(DateTime DateFrom, DateTime DateTo, bool YesNoCompact = false, bool YesNoNoSup = false, bool YesNoLimit = false, bool YesNoGroupFarm = true, bool YesNoGroupThuMua = true, bool YesNoReportM1 = false, bool YesNoByUnit = false)
        {
            try
            {
                Stopwatch stopWatch = Stopwatch.StartNew();

                #region Preparing!

                #region Initializing
                //string connectionString = System.Configuration.ConfigurationManager.ConnectionStrings["mongodb_vecrops.salesms"].ConnectionString;
                //MongoClient mongoClient = new MongoClient(connectionString);
                var mongoClient = new MongoClient();
                var db = mongoClient.GetDatabase("localtest");
                var coreStructure = new CoordStructure();

                // Need to find out how to query this shit.
                // Coz reading every fucking thing THEN query is not good. At all.
                // Solved using .Find
                // Also saved a lot of memory
                //var PO = db.GetCollection<PurchaseOrderDate>("PurchaseOrder").AsQueryable().ToList()
                //    .Where(x =>
                //       (x.DateOrder >= DateFrom.Date) &&
                //       (x.DateOrder <= DateTo.Date))
                //    .OrderBy(x => x.DateOrder);
                ////.ToList();

                var PO = db.GetCollection<PurchaseOrderDate>("PurchaseOrder").Find(x =>
                        (x.DateOrder >= DateFrom.Date) &&
                        (x.DateOrder <= DateTo.Date))
                    .ToList()
                    .OrderBy(x => x.DateOrder);

                //var FC = db.GetCollection<ForecastDate>("Forecast").AsQueryable().ToList()
                //    .Where(x =>
                //        (x.DateForecast.Date >= DateFrom.Date) &&
                //        (x.DateForecast.Date <= DateTo.Date))
                //    .OrderBy(x => x.DateForecast);

                var FC = db.GetCollection<ForecastDate>("Forecast").Find(x =>
                        (x.DateForecast >= DateFrom.Date) &&
                        (x.DateForecast <= DateTo.Date))
                    .ToList()
                    .OrderBy(x => x.DateForecast);

                var Product = db.GetCollection<Product>("Product").AsQueryable().ToList();
                coreStructure.dicProductUnit = db.GetCollection<ProductUnit>("ProductUnit").AsQueryable().ToDictionary(x => x.ProductCode);
                var Supplier = db.GetCollection<Supplier>("Supplier").AsQueryable().ToList();
                var Customer = db.GetCollection<Customer>("Customer").AsQueryable().ToList();

                coreStructure.dicProduct = new Dictionary<Guid, Product>();
                coreStructure.dicSupplier = new Dictionary<Guid, Supplier>();
                coreStructure.dicCustomer = new Dictionary<Guid, Customer>();

                var dicClass = new Dictionary<string, string>();

                dicClass.Add("A", "Rau ăn lá");
                dicClass.Add("B", "Rau ăn thân hoa");
                dicClass.Add("C", "Rau ăn quả");
                dicClass.Add("D", "Rau ăn củ");
                dicClass.Add("E", "Hạt");
                dicClass.Add("F", "Rau gia vị");
                dicClass.Add("G", "Thủy canh");
                dicClass.Add("H", "Rau mầm");
                dicClass.Add("I", "Nấm");
                dicClass.Add("J", "Lá");
                dicClass.Add("K", "Trái cây (Quả)");
                #endregion

                #region Product
                foreach (var product in Product)
                {
                    string _ProductClass = "";
                    if (dicClass.TryGetValue(product.ProductCode.Substring(0, 1), out _ProductClass))
                    {
                        product.ProductClassification = _ProductClass;
                    }
                    coreStructure.dicProduct.Add(product.ProductId, product);
                }
                #endregion

                #region Supplier
                foreach (var supplier in Supplier)
                {
                    Supplier _supplier = null;
                    if (!coreStructure.dicSupplier.TryGetValue(supplier.SupplierId, out _supplier))
                    {
                        coreStructure.dicSupplier.Add(supplier.SupplierId, supplier);
                    }
                }
                #endregion Supplier

                #region Customer
                foreach (var customer in Customer)
                {
                    Customer _customer = null;
                    if (!coreStructure.dicCustomer.TryGetValue(customer.CustomerId, out _customer))
                    {
                        coreStructure.dicCustomer.Add(customer.CustomerId, customer);
                    }
                }
                #endregion

                #region PO
                // Everything related to PO.
                int maxCalculation = 0;
                coreStructure.dicPO = new Dictionary<DateTime, Dictionary<Product, Dictionary<CustomerOrder, bool>>>();
                foreach (PurchaseOrderDate PODate in PO)
                {
                    coreStructure.dicPO.Add(PODate.DateOrder.Date, new Dictionary<Product, Dictionary<CustomerOrder, bool>>());
                    foreach (ProductOrder _ProductOrder in PODate.ListProductOrder)
                    {
                        coreStructure.dicPO[PODate.DateOrder.Date].Add(coreStructure.dicProduct[_ProductOrder.ProductId], new Dictionary<CustomerOrder, bool>());
                        foreach (CustomerOrder _CustomerOrder in _ProductOrder.ListCustomerOrder)
                        {
                            _CustomerOrder.QuantityOrder = Math.Round(_CustomerOrder.QuantityOrder, 2);
                            string _OrderUnitType = ProperUnit(_CustomerOrder.Unit.ToLower());
                            _CustomerOrder.Unit = _OrderUnitType;
                            if (coreStructure.dicCustomer[_CustomerOrder.CustomerId].CustomerType == "VM+" && _OrderUnitType != "Kg")
                            {
                                string _ProductCode = coreStructure.dicProduct[_ProductOrder.ProductId].ProductCode;
                                ProductUnit _ProductUnit = null;
                                if (coreStructure.dicProductUnit.TryGetValue(_ProductCode, out _ProductUnit))
                                {
                                    ProductUnitRegion _ProductUnitRegion = _ProductUnit.ListRegion.Where(x => x.OrderUnitType == _OrderUnitType).FirstOrDefault();
                                    if (_ProductUnitRegion != null)
                                    {
                                        _CustomerOrder.Unit = _OrderUnitType;
                                        _CustomerOrder.QuantityOrderKg = _CustomerOrder.QuantityOrder * _ProductUnitRegion.OrderUnitPer;
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
                            coreStructure.dicPO[PODate.DateOrder.Date][coreStructure.dicProduct[_ProductOrder.ProductId]].Add(_CustomerOrder, true);
                            maxCalculation++;
                        }
                    }
                }
                #endregion

                #region FC
                coreStructure.dicFC = new Dictionary<DateTime, Dictionary<Product, Dictionary<SupplierForecast, bool>>>();
                foreach (ForecastDate FCDate in FC)
                {
                    coreStructure.dicFC.Add(FCDate.DateForecast.Date, new Dictionary<Product, Dictionary<SupplierForecast, bool>>());
                    foreach (ProductForecast _ProductForecast in FCDate.ListProductForecast)
                    {
                        coreStructure.dicFC[FCDate.DateForecast.Date].Add(coreStructure.dicProduct[_ProductForecast.ProductId], new Dictionary<SupplierForecast, bool>());
                        // To allow user to store their plans on the Forecast, .Where() here serves as the filter - Bringing out only stuff that "can" be used.
                        foreach (SupplierForecast _SupplierForecast in _ProductForecast.ListSupplierForecast.Where(x => x.QualityControlPass == true))
                        {
                            //// To deal with unexpected cases of Cross Region for VinEco.
                            //// Have been dealt with in UpdateForecast.
                            //if (coreStructure.dicSupplier[_SupplierForecast.SupplierId].SupplierType == "VinEco" && dicCrossRegionVinEco.ContainsKey(coreStructure.dicProduct[_ProductForecast.ProductId].ProductCode))
                            //{
                            //    _SupplierForecast.CrossRegion = dicCrossRegionVinEco[coreStructure.dicProduct[_ProductForecast.ProductId].ProductCode];
                            //}
                            coreStructure.dicFC[FCDate.DateForecast.Date][coreStructure.dicProduct[_ProductForecast.ProductId]].Add(_SupplierForecast, true);
                        }
                    }
                }
                #endregion

                #region Best of both worlds.
                coreStructure.dicCoord = new Dictionary<DateTime, Dictionary<Product, Dictionary<CustomerOrder, Dictionary<SupplierForecast, DateTime>>>>();
                foreach (PurchaseOrderDate PODate in PO)
                {
                    coreStructure.dicCoord.Add(PODate.DateOrder.Date, new Dictionary<MainForm.Product, Dictionary<CustomerOrder, Dictionary<SupplierForecast, DateTime>>>());
                    foreach (ProductOrder _ProductOrder in PODate.ListProductOrder)
                    {
                        coreStructure.dicCoord[PODate.DateOrder.Date].Add(coreStructure.dicProduct[_ProductOrder.ProductId], new Dictionary<CustomerOrder, Dictionary<SupplierForecast, DateTime>>());
                        foreach (CustomerOrder _CustomerOrder in _ProductOrder.ListCustomerOrder)
                        {
                            _CustomerOrder.QuantityOrder = Math.Round(_CustomerOrder.QuantityOrder, 1);
                            coreStructure.dicCoord[PODate.DateOrder.Date][coreStructure.dicProduct[_ProductOrder.ProductId]].Add(_CustomerOrder, null);
                        }
                    }
                }
                #endregion

                #region VE Farm Table
                DataTable dtVeFarm = new DataTable();

                dtVeFarm.Columns.Add("Region", typeof(string));
                dtVeFarm.Columns.Add("SCODE", typeof(string));
                dtVeFarm.Columns.Add("SNAME", typeof(string));
                dtVeFarm.Columns.Add("PCLASS", typeof(string));
                dtVeFarm.Columns.Add("VECrops Code", typeof(string));
                dtVeFarm.Columns.Add("PCODE", typeof(string));
                dtVeFarm.Columns.Add("PNAME", typeof(string));

                foreach (DateTime DateFC in coreStructure.dicFC.Keys.OrderBy(x => x.Date))
                {
                    string _colName = DateFC.Date.ToString();
                    dtVeFarm.Columns.Add(_colName, typeof(double)).DefaultValue = 0;
                }

                var dicRow = new Dictionary<string, int>();
                foreach (DateTime DateFC in coreStructure.dicFC.Keys.OrderBy(x => x.Date))
                {
                    string _colName = DateFC.Date.ToString();
                    foreach (Product _Product in coreStructure.dicFC[DateFC].Keys.OrderBy(x => x.ProductCode))
                    {
                        var _listSupplierForecast = coreStructure.dicFC[DateFC][_Product].Keys.Where(x => coreStructure.dicSupplier[x.SupplierId].SupplierType == "VinEco");
                        if (_listSupplierForecast != null)
                        {
                            foreach (SupplierForecast _SupplierForecast in _listSupplierForecast.OrderBy(x => coreStructure.dicSupplier[x.SupplierId].SupplierName))
                            {
                                Supplier _Supplier = coreStructure.dicSupplier[_SupplierForecast.SupplierId];
                                string sKey = String.Format("{0}{1}", _Product.ProductCode, _Supplier.SupplierCode);

                                DataRow dr = null;

                                int _rowIndex = 0;
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
                }
                #endregion

                #region ThuMua Table
                DataTable dtThuMua = new DataTable();

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

                foreach (DateTime DateFC in coreStructure.dicFC.Keys.OrderBy(x => x.Date))
                {
                    string _colName = DateFC.Date.ToString();
                    dtThuMua.Columns.Add(_colName, typeof(double)).DefaultValue = 0;
                }

                dicRow = new Dictionary<string, int>();
                foreach (DateTime DateFC in coreStructure.dicFC.Keys.OrderBy(x => x.Date))
                {
                    string _colName = DateFC.Date.ToString();
                    foreach (Product _Product in coreStructure.dicFC[DateFC].Keys.OrderBy(x => x.ProductCode))
                    {
                        var _listSupplierForecast = coreStructure.dicFC[DateFC][_Product].Keys.Where(x => coreStructure.dicSupplier[x.SupplierId].SupplierType != "VinEco");
                        if (_listSupplierForecast != null)
                        {
                            foreach (SupplierForecast _SupplierForecast in _listSupplierForecast.OrderBy(x => coreStructure.dicSupplier[x.SupplierId].SupplierName))
                            {
                                Supplier _Supplier = coreStructure.dicSupplier[_SupplierForecast.SupplierId];
                                string sKey = String.Format("{0}{1}", _Product.ProductCode, _Supplier.SupplierCode);

                                DataRow dr = null;

                                int _rowIndex = 0;
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
                                dr["Level"] = _SupplierForecast.level;
                                dr["Availability"] = _SupplierForecast.Availability;

                                dr[_colName] = Convert.ToDouble(dr[_colName]) + _SupplierForecast.QuantityForecast;
                            }
                        }
                    }
                }
                #endregion
                #endregion

                #region Main Body
                if (!YesNoLimit)
                {
                    if (!YesNoNoSup)
                    {
                        Coord(coreStructure, "Miền Bắc", "Miền Bắc", "VCM", 1, 2, 1, false, "VM+");
                        Coord(coreStructure, "Miền Nam", "Miền Nam", "VCM", 0, 0, 1, false, "VM+");
                    }

                    Coord(coreStructure, "Miền Bắc", "Miền Bắc", "VinEco", 1, 2, -1, false, "VM+");
                    Coord(coreStructure, "Miền Nam", "Miền Nam", "VinEco", 0, 0, -1, false, "VM+");
                    Coord(coreStructure, "Lâm Đồng", "Miền Bắc", "VinEco", 2, 2, -1, false, "VM+");
                    Coord(coreStructure, "Lâm Đồng", "Miền Nam", "VinEco", 0, 0, -1, false, "VM+");
                    Coord(coreStructure, "Miền Nam", "Miền Bắc", "VinEco", 3, 2, -1, true, "VM+");
                    Coord(coreStructure, "Miền Bắc", "Miền Nam", "VinEco", 3, 0, -1, true, "VM+");

                    Coord(coreStructure, "Miền Bắc", "Miền Bắc", "ThuMua", 1, 2, 1, false, "VM+");
                    Coord(coreStructure, "Miền Nam", "Miền Nam", "ThuMua", 0, 0, 1, false, "VM+");
                    Coord(coreStructure, "Lâm Đồng", "Miền Bắc", "ThuMua", 2, 2, 1, false, "VM+");
                    Coord(coreStructure, "Lâm Đồng", "Miền Nam", "ThuMua", 0, 0, 1, false, "VM+");
                    Coord(coreStructure, "Miền Nam", "Miền Bắc", "ThuMua", 3, 2, 1, true, "VM+");
                    Coord(coreStructure, "Miền Bắc", "Miền Nam", "ThuMua", 3, 0, 1, true, "VM+");

                    Coord(coreStructure, "Miền Bắc", "Miền Bắc", "VinEco", 1, 2, -1, false);
                    Coord(coreStructure, "Miền Nam", "Miền Nam", "VinEco", 0, 0, -1, false);
                    Coord(coreStructure, "Lâm Đồng", "Miền Bắc", "VinEco", 2, 2, -1, false);
                    Coord(coreStructure, "Lâm Đồng", "Miền Nam", "VinEco", 0, 0, -1, false);
                    Coord(coreStructure, "Miền Nam", "Miền Bắc", "VinEco", 3, 2, -1, true);
                    Coord(coreStructure, "Miền Bắc", "Miền Nam", "VinEco", 3, 0, -1, true);

                    Coord(coreStructure, "Miền Bắc", "Miền Bắc", "ThuMua", 1, 2, 1, false);
                    Coord(coreStructure, "Miền Nam", "Miền Nam", "ThuMua", 0, 0, 1, false);
                    Coord(coreStructure, "Lâm Đồng", "Miền Bắc", "ThuMua", 2, 2, 1, false);
                    Coord(coreStructure, "Lâm Đồng", "Miền Nam", "ThuMua", 0, 0, 1, false);
                    Coord(coreStructure, "Miền Nam", "Miền Bắc", "ThuMua", 3, 2, 1, true);
                    Coord(coreStructure, "Miền Bắc", "Miền Nam", "ThuMua", 3, 0, 1, true);
                }
                else
                {
                    Coord(coreStructure, "Miền Bắc", "Miền Bắc", "VinEco", 1, 2, 1, false, "Adayroi", YesNoByUnit);
                    Coord(coreStructure, "Lâm Đồng", "Miền Bắc", "VinEco", 2, 2, 1, false, "Adayroi", YesNoByUnit);

                    if (!YesNoNoSup)
                    {
                        Coord(coreStructure, "Miền Bắc", "Miền Bắc", "VCM", 1, 2, 1, false, "VM+", YesNoByUnit);
                        Coord(coreStructure, "Miền Nam", "Miền Nam", "VCM", 0, 0, 1, false, "VM+", YesNoByUnit);
                    }

                    Coord(coreStructure, "Miền Bắc", "Miền Bắc", "VinEco", 1, 2, 1, false, "VM+", YesNoByUnit);
                    Coord(coreStructure, "Miền Nam", "Miền Nam", "VinEco", 0, 0, 1, false, "VM+", YesNoByUnit);
                    Coord(coreStructure, "Lâm Đồng", "Miền Bắc", "VinEco", 2, 2, 1, false, "VM+", YesNoByUnit);
                    Coord(coreStructure, "Lâm Đồng", "Miền Nam", "VinEco", 0, 0, 1, false, "VM+", YesNoByUnit);
                    Coord(coreStructure, "Miền Nam", "Miền Bắc", "VinEco", 3, 2, 1, true, "VM+", YesNoByUnit);
                    Coord(coreStructure, "Miền Bắc", "Miền Nam", "VinEco", 3, 0, 1, true, "VM+", YesNoByUnit);

                    Coord(coreStructure, "Miền Bắc", "Miền Bắc", "ThuMua", 1, 2, 1, false, "VM+", YesNoByUnit);
                    Coord(coreStructure, "Miền Nam", "Miền Nam", "ThuMua", 0, 0, 1, false, "VM+", YesNoByUnit);
                    Coord(coreStructure, "Lâm Đồng", "Miền Bắc", "ThuMua", 2, 2, 1, false, "VM+", YesNoByUnit);
                    Coord(coreStructure, "Lâm Đồng", "Miền Nam", "ThuMua", 0, 0, 1, false, "VM+", YesNoByUnit);
                    Coord(coreStructure, "Miền Nam", "Miền Bắc", "ThuMua", 3, 2, 1, true, "VM+", YesNoByUnit);
                    Coord(coreStructure, "Miền Bắc", "Miền Nam", "ThuMua", 3, 0, 1, true, "VM+", YesNoByUnit);

                    Coord(coreStructure, "Miền Bắc", "Miền Bắc", "VinEco", 1, 2, 1, false);
                    Coord(coreStructure, "Miền Nam", "Miền Nam", "VinEco", 0, 0, 1, false);
                    Coord(coreStructure, "Lâm Đồng", "Miền Bắc", "VinEco", 2, 2, 1, false);
                    Coord(coreStructure, "Lâm Đồng", "Miền Nam", "VinEco", 0, 0, 1, false);
                    Coord(coreStructure, "Miền Nam", "Miền Bắc", "VinEco", 3, 2, 1, true);
                    Coord(coreStructure, "Miền Bắc", "Miền Nam", "VinEco", 3, 0, 1, true);

                    Coord(coreStructure, "Miền Bắc", "Miền Bắc", "ThuMua", 1, 2, 1, false);
                    Coord(coreStructure, "Miền Nam", "Miền Nam", "ThuMua", 0, 0, 1, false);
                    Coord(coreStructure, "Lâm Đồng", "Miền Bắc", "ThuMua", 2, 2, 1, false);
                    Coord(coreStructure, "Lâm Đồng", "Miền Nam", "ThuMua", 0, 0, 1, false);
                    Coord(coreStructure, "Miền Nam", "Miền Bắc", "ThuMua", 3, 2, 1, true);
                    Coord(coreStructure, "Miền Bắc", "Miền Nam", "ThuMua", 3, 0, 1, true);
                }
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
                if (YesNoReportM1)
                {

                    #region Mastah Table - Report M+1
                    DataTable dtMastah = new DataTable(tableName: "ReportM1");

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

                    dicRow = new Dictionary<string, int>();

                    foreach (DateTime DatePO in coreStructure.dicCoord.Keys)
                    {
                        foreach (Product _Product in coreStructure.dicCoord[DatePO].Keys)
                        {
                            foreach (CustomerOrder _CustomerOrder in coreStructure.dicCoord[DatePO][_Product].Keys)
                            {
                                if (coreStructure.dicCoord[DatePO][_Product][_CustomerOrder] != null)
                                {
                                    foreach (SupplierForecast _SupplierForecast in coreStructure.dicCoord[DatePO][_Product][_CustomerOrder].Keys.OrderBy(x => coreStructure.dicSupplier[x.SupplierId].SupplierType))
                                    {
                                        DataRow dr = null;

                                        Customer _Customer = coreStructure.dicCustomer[_CustomerOrder.CustomerId];
                                        Supplier _Supplier = coreStructure.dicSupplier[_SupplierForecast.SupplierId];

                                        string sKey = DatePO.Date + _Product.ProductCode + _Customer.CustomerType + _Customer.CustomerRegion;
                                        bool newRow = false;

                                        int _rowPos = 0;
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
                                        dr["Ngày tiêu thụ"] = (int)(DatePO.Date - new DateTime(1900, 1, 1)).TotalDays + 2;
                                        dr["Vùng tiêu thụ"] = _Customer.CustomerBigRegion;
                                        dr["Tỉnh tiêu thụ"] = _Customer.CustomerRegion;

                                        //dr["NoSup"] = "";

                                        string productClass = "";
                                        if (!dicClass.TryGetValue(_Product.ProductCode.Substring(0, 1), out productClass))
                                        {
                                            productClass = "???";
                                        }
                                        dr["Class"] = productClass;

                                        //dr["DS1"] = "";
                                        //dr["Region"] = "";
                                        //dr["Bắt buộc?"] = "";

                                        dr["VCM"] = (double)(dr["VCM"]) + _CustomerOrder.QuantityOrderKg;
                                        if (_Supplier.SupplierType == "VinEco")
                                        {
                                            dr["VE"] = (double)(dr["VE"]) + _SupplierForecast.QuantityForecast;
                                        }
                                        else if (_Supplier.SupplierType == "ThuMua")
                                        {
                                            dr["TM"] = (double)(dr["TM"]) + _SupplierForecast.QuantityForecast;

                                        }

                                        if (newRow)
                                        {
                                            dtMastah.Rows.Add(dr);
                                        }

                                    }
                                }
                                else
                                {
                                    DataRow dr = null;

                                    Customer _Customer = coreStructure.dicCustomer[_CustomerOrder.CustomerId];

                                    string sKey = DatePO.Date + _Product.ProductCode + _Customer.CustomerType + _Customer.CustomerRegion;
                                    bool newRow = false;

                                    int _rowPos = 0;
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
                                    dr["Ngày tiêu thụ"] = (int)(DatePO.Date - new DateTime(1900, 1, 1)).TotalDays + 2;
                                    dr["Vùng tiêu thụ"] = _Customer.CustomerBigRegion;
                                    dr["Tỉnh tiêu thụ"] = _Customer.CustomerRegion;

                                    //dr["NoSup"] = "";

                                    string productClass = "";
                                    if (!dicClass.TryGetValue(_Product.ProductCode.Substring(0, 1), out productClass))
                                    {
                                        productClass = "???";
                                    }
                                    dr["Class"] = productClass;

                                    //dr["DS1"] = "";
                                    //dr["Region"] = "";
                                    //dr["Bắt buộc?"] = "";

                                    dr["VCM"] = (double)(dr["VCM"]) + _CustomerOrder.QuantityOrderKg;


                                    if (newRow)
                                    {
                                        dtMastah.Rows.Add(dr);
                                    }
                                }
                            }
                        }
                    }

                    foreach (DataRow dr in dtMastah.Rows)
                    {
                        if ((double)dr["VCM"] > (double)dr["VE"] + (double)dr["TM"])
                        {
                            dr["NoSup"] = "Yes";
                        }
                    }

                    #endregion

                    #region LeftoverVinEco
                    DataTable dtLeftoverVe = new DataTable();

                    dtLeftoverVe.TableName = "NoCusVinEco";

                    dtLeftoverVe.Columns.Add("Mã VinEco", typeof(string)).DefaultValue = "";
                    dtLeftoverVe.Columns.Add("Tên VinEco", typeof(string)).DefaultValue = "";
                    dtLeftoverVe.Columns.Add("Mã Farm", typeof(string)).DefaultValue = "";
                    dtLeftoverVe.Columns.Add("Tên Farm", typeof(string)).DefaultValue = "";
                    dtLeftoverVe.Columns.Add("Ngày thu hoạch", typeof(int));
                    dtLeftoverVe.Columns.Add("Vùng sản xuất", typeof(string)).DefaultValue = "";
                    dtLeftoverVe.Columns.Add("Sản lượng", typeof(double)).DefaultValue = 0;

                    foreach (DateTime DateFC in coreStructure.dicFC.Keys)
                    {
                        foreach (Product _Product in coreStructure.dicFC[DateFC].Keys.OrderBy(x => x.ProductCode))
                        {
                            var _ListSupplier = coreStructure.dicFC[DateFC][_Product].Keys.Where(x => coreStructure.dicSupplier[x.SupplierId].SupplierType == "VinEco");
                            if (_ListSupplier != null)
                            {
                                foreach (SupplierForecast _SupplierForecast in _ListSupplier.OrderBy(x => coreStructure.dicSupplier[x.SupplierId].SupplierName))
                                {
                                    if (_SupplierForecast.QuantityForecast > 0)
                                    {
                                        DataRow dr = dtLeftoverVe.NewRow();

                                        //Customer _Customer =coreStructure. dicCustomer[_CustomerOrder.CustomerId];
                                        Supplier _Supplier = coreStructure.dicSupplier[_SupplierForecast.SupplierId];

                                        dr["Mã VinEco"] = _Product.ProductCode;
                                        dr["Tên VinEco"] = _Product.ProductName;
                                        dr["Mã Farm"] = _Supplier.SupplierCode;
                                        dr["Tên Farm"] = _Supplier.SupplierName;
                                        dr["Ngày thu hoạch"] = (int)(DateFC.Date - new DateTime(1900, 1, 1)).TotalDays + 2;
                                        dr["Vùng sản xuất"] = _Supplier.SupplierRegion;
                                        dr["Sản lượng"] = _SupplierForecast.QuantityForecast;

                                        dtLeftoverVe.Rows.Add(dr);
                                    }
                                }
                            }
                        }
                    }
                    #endregion

                    #region LeftoverThuMua
                    DataTable dtLeftoverTm = new DataTable();

                    dtLeftoverTm.TableName = "NoCusThuMua";

                    dtLeftoverTm.Columns.Add("Mã VinEco", typeof(string)).DefaultValue = "";
                    dtLeftoverTm.Columns.Add("Tên VinEco", typeof(string)).DefaultValue = "";
                    dtLeftoverTm.Columns.Add("Mã Farm", typeof(string)).DefaultValue = "";
                    dtLeftoverTm.Columns.Add("Tên Farm", typeof(string)).DefaultValue = "";
                    dtLeftoverTm.Columns.Add("Ngày thu hoạch", typeof(int));
                    dtLeftoverTm.Columns.Add("Vùng sản xuất", typeof(string)).DefaultValue = "";
                    dtLeftoverTm.Columns.Add("Sản lượng", typeof(double)).DefaultValue = 0;

                    foreach (DateTime DateFC in coreStructure.dicFC.Keys)
                    {
                        foreach (Product _Product in coreStructure.dicFC[DateFC].Keys.OrderBy(x => x.ProductCode))
                        {
                            var _ListSupplier = coreStructure.dicFC[DateFC][_Product].Keys.Where(x => coreStructure.dicSupplier[x.SupplierId].SupplierType == "ThuMua");
                            if (_ListSupplier != null)
                            {
                                foreach (SupplierForecast _SupplierForecast in _ListSupplier.OrderBy(x => coreStructure.dicSupplier[x.SupplierId].SupplierName))
                                {
                                    if (_SupplierForecast.QuantityForecast > 0)
                                    {
                                        DataRow dr = dtLeftoverTm.NewRow();

                                        //Customer _Customer =coreStructure. dicCustomer[_CustomerOrder.CustomerId];
                                        Supplier _Supplier = coreStructure.dicSupplier[_SupplierForecast.SupplierId];

                                        dr["Mã VinEco"] = _Product.ProductCode;
                                        dr["Tên VinEco"] = _Product.ProductName;
                                        dr["Mã Farm"] = _Supplier.SupplierCode;
                                        dr["Tên Farm"] = _Supplier.SupplierName;
                                        dr["Ngày thu hoạch"] = (int)(DateFC.Date - new DateTime(1900, 1, 1)).TotalDays + 2;
                                        dr["Vùng sản xuất"] = _Supplier.SupplierRegion;
                                        dr["Sản lượng"] = _SupplierForecast.QuantityForecast;

                                        dtLeftoverTm.Rows.Add(dr);
                                    }
                                }
                            }
                        }
                    }
                    #endregion

                    #region Output to Excel
                    var dicDateCol = new Dictionary<string, int>();

                    dicDateCol.Add("Ngày tiêu thụ", dtMastah.Columns.IndexOf("Ngày tiêu thụ"));

                    string fileName = string.Format("Report M plus 1 {0}.xlsx", DateFrom.ToString("dd.MM") + " - " + DateTo.ToString("dd.MM") + " (" + DateTime.Now.ToString("yyyyMMdd HH\\hmm") + ")");
                    string path = string.Format(@"C:\Users\Shirayuki\Documents\VinEco\Project - Chia Hang\Mastah Project\Test\" + fileName);

                    var listDt = new List<DataTable>();

                    listDt.Add(dtMastah);
                    listDt.Add(dtLeftoverVe);
                    listDt.Add(dtLeftoverTm);

                    LargeExportOneWorkbook(path, listDt, true, true);

                    ConvertToXlsbInterop(path, "xlsx", "xlsb", true);
                    #endregion

                }
                else if (YesNoNoSup)
                {

                    #region Mastah Table - NoSup
                    DataTable dtMastah = new DataTable();

                    dtMastah.Columns.Add("Mã VinEco", typeof(string)).DefaultValue = "";
                    dtMastah.Columns.Add("Tên VinEco", typeof(string)).DefaultValue = "";
                    dtMastah.Columns.Add("Mã cửa hàng", typeof(string)).DefaultValue = "";
                    dtMastah.Columns.Add("Tên cửa hàng", typeof(string)).DefaultValue = "";
                    dtMastah.Columns.Add("Loại cửa hàng", typeof(string)).DefaultValue = "";
                    dtMastah.Columns.Add("Ngày tiêu thụ", typeof(double)).DefaultValue = 0;
                    dtMastah.Columns.Add("Vùng tiêu thụ", typeof(string)).DefaultValue = "";
                    dtMastah.Columns.Add("Nhu cầu", typeof(double)).DefaultValue = 0;

                    foreach (DateTime DatePO in coreStructure.dicCoord.Keys)
                    {
                        foreach (Product _Product in coreStructure.dicCoord[DatePO].Keys)
                        {
                            foreach (CustomerOrder _CustomerOrder in coreStructure.dicCoord[DatePO][_Product].Keys)
                            {
                                if (coreStructure.dicCoord[DatePO][_Product][_CustomerOrder] == null)
                                {
                                    DataRow dr = dtMastah.NewRow();

                                    Customer _Customer = coreStructure.dicCustomer[_CustomerOrder.CustomerId];
                                    //Supplier _Supplier =coreStructure. dicSupplier[_SupplierForecast.SupplierId];

                                    dr["Mã VinEco"] = _Product.ProductCode;
                                    dr["Tên VinEco"] = _Product.ProductName;
                                    dr["Mã cửa hàng"] = _Customer.CustomerCode;
                                    dr["Tên cửa hàng"] = _Customer.CustomerName;
                                    dr["Loại cửa hàng"] = _Customer.CustomerType;
                                    //dr["Ngày tiêu thụ"] = DatePO.Date;
                                    dr["Ngày tiêu thụ"] = (int)(DatePO.Date - new DateTime(1900, 1, 1)).TotalDays + 2;
                                    dr["Vùng tiêu thụ"] = _Customer.CustomerBigRegion;  // == "Miền Bắc" ? "MB" : "MN";
                                    dr["Nhu cầu"] = _CustomerOrder.QuantityOrderKg;

                                    dtMastah.Rows.Add(dr);
                                }
                            }
                        }
                    }

                    var dicDateCol = new Dictionary<string, int>();
                    dicDateCol.Add("Ngày tiêu thụ", dtMastah.Columns.IndexOf("Ngày tiêu thụ"));
                    #endregion

                    #region Output to Excel
                    //Excel.Application xlApp = new Excel.Application();

                    //xlApp.ScreenUpdating = false;
                    //xlApp.EnableEvents = false;
                    //xlApp.DisplayAlerts = false;
                    //xlApp.DisplayStatusBar = false;
                    //xlApp.AskToUpdateLinks = false;

                    string filePath = Application.StartupPath.Replace("\\bin\\Debug", "") + "\\Template\\{0}";
                    string fileFullPath = string.Format(filePath, "NoSup.xlsb");

                    //Excel.Workbook xlWb = xlApp.Workbooks.Add();

                    //xlApp.Calculation = Excel.XlCalculation.xlCalculationManual;

                    var missing = Type.Missing;
                    //string path = string.Format(@"C:\Users\Shirayuki\Documents\VinEco\Project - Chia Hang\Mastah Project\Test\NoSup {0}.xlsb", DateFrom.ToString("dd.MM") + " - " + DateTo.ToString("dd.MM") + " (" + DateTime.Now.ToString("yyyyMMdd HH\\hmm") + ")");

                    string fileName = string.Format("NoSup {0}.xlsx", DateFrom.ToString("dd.MM") + " - " + DateTo.ToString("dd.MM") + " (" + DateTime.Now.ToString("yyyyMMdd HH\\hmm") + ")");
                    string path = string.Format(@"C:\Users\Shirayuki\Documents\VinEco\Project - Chia Hang\Mastah Project\Test\" + fileName);

                    LargeExport(dtMastah, path, dicDateCol, true, false);

                    ConvertToXlsbInterop(path, "xlsx", "xlsb", true);

                    //xlWb.SaveAs(path, Excel.XlFileFormat.xlExcel12, missing, missing, false, false, Excel.XlSaveAsAccessMode.xlNoChange, missing, missing, missing);

                    //OutputExcel(dtMastah, "Sheet1", xlWb, true, 1);
                    //dtMastah = null;

                    //xlApp.Calculation = Excel.XlCalculation.xlCalculationAutomatic;

                    //xlWb.Close(SaveChanges: true);
                    //Marshal.ReleaseComObject(xlWb);
                    //xlWb = null;

                    //xlApp.ScreenUpdating = true;
                    //xlApp.EnableEvents = true;
                    //xlApp.DisplayAlerts = false;
                    //xlApp.DisplayStatusBar = true;
                    //xlApp.AskToUpdateLinks = true;

                    //xlApp.Quit();
                    //if (xlApp != null) { Marshal.ReleaseComObject(xlApp); }
                    //xlApp = null;
                    #endregion

                }
                else if (!YesNoCompact)
                {

                    #region Mastah Table
                    DataTable dtMastah = new DataTable();

                    dtMastah.Columns.Add("Mã 6 ký tự", typeof(string)).DefaultValue = "";
                    dtMastah.Columns.Add("Mã thành phẩm VinEco", typeof(string)).DefaultValue = "";
                    dtMastah.Columns.Add("Mã thành phẩm VinCommerce", typeof(string)).DefaultValue = "";
                    dtMastah.Columns.Add("Tên Sản phẩm", typeof(string)).DefaultValue = "";
                    dtMastah.Columns.Add("Mã Cửa hàng", typeof(string)).DefaultValue = "";
                    dtMastah.Columns.Add("Tên Cửa hàng", typeof(string)).DefaultValue = "";
                    dtMastah.Columns.Add("Loại Cửa hàng", typeof(string)).DefaultValue = "";
                    dtMastah.Columns.Add("Ngày Tiêu thụ", typeof(DateTime));
                    dtMastah.Columns.Add("Vùng Tiêu thụ", typeof(string)).DefaultValue = "";

                    dtMastah.Columns.Add("Nhu cầu Kg VinCommerce", typeof(double)).DefaultValue = 0;
                    dtMastah.Columns.Add("Số lượng đặt", typeof(double)).DefaultValue = 0;
                    dtMastah.Columns.Add("Đơn vị đặt", typeof(string)).DefaultValue = "";
                    dtMastah.Columns.Add("Đặt Kg/Unit", typeof(double)).DefaultValue = 0;

                    dtMastah.Columns.Add("Nhu cầu Kg Đã đáp ứng", typeof(string)).DefaultValue = "";
                    dtMastah.Columns.Add("Số lượng bán", typeof(string)).DefaultValue = "";
                    dtMastah.Columns.Add("Đơn vị bán", typeof(string)).DefaultValue = "";
                    dtMastah.Columns.Add("Bán Kg/Unit", typeof(double)).DefaultValue = 0;

                    dtMastah.Columns.Add("Tên VinEco MB", typeof(string)).DefaultValue = "";
                    dtMastah.Columns.Add("Đáp ứng từ VinEco MB", typeof(double)).DefaultValue = 0;
                    dtMastah.Columns.Add("Ngày sơ chế VinEco MB", typeof(DateTime));

                    dtMastah.Columns.Add("Tên VinEco MN", typeof(string)).DefaultValue = "";
                    dtMastah.Columns.Add("Đáp ứng từ VinEco MN", typeof(double)).DefaultValue = 0;
                    dtMastah.Columns.Add("Ngày sơ chế VinEco MN", typeof(DateTime));

                    dtMastah.Columns.Add("Tên VinEco LĐ", typeof(string)).DefaultValue = "";
                    dtMastah.Columns.Add("Đáp ứng từ VinEco LĐ", typeof(double)).DefaultValue = 0;
                    dtMastah.Columns.Add("Ngày sơ chế VinEco LĐ", typeof(DateTime));

                    dtMastah.Columns.Add("Tên ThuMua MB", typeof(string)).DefaultValue = "";
                    dtMastah.Columns.Add("Đáp ứng từ ThuMua MB", typeof(double)).DefaultValue = 0;
                    dtMastah.Columns.Add("Ngày sơ chế ThuMua MB", typeof(DateTime));
                    dtMastah.Columns.Add("Giá mua ThuMua MB", typeof(double)).DefaultValue = 0;

                    dtMastah.Columns.Add("Tên ThuMua MN", typeof(string)).DefaultValue = "";
                    dtMastah.Columns.Add("Đáp ứng từ ThuMua MN", typeof(double)).DefaultValue = 0;
                    dtMastah.Columns.Add("Ngày sơ chế ThuMua MN", typeof(DateTime));
                    dtMastah.Columns.Add("Giá mua ThuMua MN", typeof(double)).DefaultValue = 0;

                    dtMastah.Columns.Add("Tên ThuMua LĐ", typeof(string)).DefaultValue = "";
                    dtMastah.Columns.Add("Đáp ứng từ ThuMua LĐ", typeof(double)).DefaultValue = 0;
                    dtMastah.Columns.Add("Ngày sơ chế ThuMua LĐ", typeof(DateTime));
                    dtMastah.Columns.Add("Giá mua ThuMua LĐ", typeof(double)).DefaultValue = 0;

                    dtMastah.Columns.Add("Note", typeof(string)).DefaultValue = "";

                    foreach (DateTime DatePO in coreStructure.dicCoord.Keys.OrderBy(x => x.Date).Where(x => x.Date >= DateFrom.AddDays(3).Date))
                    {
                        foreach (Product _Product in coreStructure.dicCoord[DatePO].Keys.OrderBy(x => x.ProductCode))
                        {
                            foreach (CustomerOrder _CustomerOrder in coreStructure.dicCoord[DatePO][_Product].Keys.Where(x => x.QuantityOrder > 0).OrderBy(x => coreStructure.dicCustomer[x.CustomerId].CustomerType))
                            {
                                if (coreStructure.dicCoord[DatePO][_Product][_CustomerOrder] != null)
                                {
                                    foreach (SupplierForecast _SupplierForecast in coreStructure.dicCoord[DatePO][_Product][_CustomerOrder].Keys.OrderBy(x => coreStructure.dicSupplier[x.SupplierId].SupplierType))
                                    {
                                        DataRow dr = dtMastah.NewRow();

                                        Customer _Customer = coreStructure.dicCustomer[_CustomerOrder.CustomerId];
                                        Supplier _Supplier = coreStructure.dicSupplier[_SupplierForecast.SupplierId];
                                        ProductUnit _ProductUnit = null;
                                        ProductUnitRegion _ProductUnitRegion = null;
                                        if (coreStructure.dicProductUnit.TryGetValue(_Product.ProductCode, out _ProductUnit))
                                        {
                                            _ProductUnitRegion = coreStructure.dicProductUnit[_Product.ProductCode].ListRegion.Where(x => x.OrderUnitType == _CustomerOrder.Unit).FirstOrDefault();
                                            if (_ProductUnitRegion == null)
                                            {
                                                _ProductUnitRegion = new ProductUnitRegion()
                                                {
                                                    OrderUnitType = "Kg",
                                                    OrderUnitPer = 1,
                                                    SaleUnitType = "Kg",
                                                    SaleUnitPer = 1
                                                };
                                            }
                                        }
                                        else
                                        {
                                            _ProductUnitRegion = new ProductUnitRegion()
                                            {
                                                OrderUnitType = "Kg",
                                                OrderUnitPer = 1,
                                                SaleUnitType = "Kg",
                                                SaleUnitPer = 1
                                            };
                                        }

                                        dr["Mã 6 ký tự"] = _Product.ProductCode;
                                        dr["Mã thành phẩm VinEco"] = _Product.ProductCode;
                                        dr["Mã thành phẩm VinCommerce"] = "";
                                        dr["Tên Sản phẩm"] = _Product.ProductName;
                                        dr["Mã Cửa hàng"] = _Customer.CustomerCode;
                                        dr["Tên Cửa hàng"] = _Customer.CustomerName;
                                        dr["Loại Cửa hàng"] = _Customer.CustomerType;
                                        dr["Ngày Tiêu thụ"] = DatePO.Date;
                                        dr["Vùng Tiêu thụ"] = _Customer.CustomerBigRegion == "Miền Bắc" ? "MB" : "MN";
                                        dr["Nhu cầu Kg VinCommerce"] = _CustomerOrder.QuantityOrderKg;

                                        dr["Số lượng đặt"] = _CustomerOrder.QuantityOrder;
                                        dr["Đơn vị đặt"] = _CustomerOrder.Unit;
                                        dr["Đặt Kg/Unit"] = _ProductUnitRegion.OrderUnitPer;

                                        //dr["Số lượng bán"] = (double)_SupplierForecast.QuantityForecast / (double)_ProductUnitRegion.SaleUnitPer;
                                        dr["Số lượng bán"] = String.Format("= N{0} * Q{0}", dtMastah.Rows.Count + 6);
                                        dr["Đơn vị bán"] = _ProductUnitRegion.SaleUnitType;
                                        dr["Bán Kg/Unit"] = _ProductUnitRegion.SaleUnitPer;

                                        //dr["Nhu cầu Đã đáp ứng"] = String.Format("=SUM(M{0}, P{0}, S{0}, V{0}, Z{0}, AD{0})", dtMastah.Rows.Count + 6);
                                        dr["Nhu cầu Kg Đã đáp ứng"] = String.Format("=SUM( S{0}, V{0}, Y{0}, AB{0}, AF{0}, AJ{0} )", dtMastah.Rows.Count + 6);
                                        //dr["Nhu cầu Kg Đã đáp ứng"] = (double)dr["Nhu cầu Kg Đã đáp ứng"] + _CustomerOrder.QuantityOrderKg;

                                        string _Region = string.Join(String.Empty, _Supplier.SupplierRegion.Where((ch, index) => ch != ' ' && (index == 0 || _Supplier.SupplierRegion[index - 1] == ' ')));
                                        switch (_Supplier.SupplierType)
                                        {
                                            case "VinEco":
                                                dr["Tên VinEco " + _Region] = _Supplier.SupplierName;
                                                dr["Đáp ứng từ VinEco " + _Region] = _SupplierForecast.QuantityForecast;
                                                dr["Ngày sơ chế VinEco " + _Region] = coreStructure.dicCoord[DatePO][_Product][_CustomerOrder][_SupplierForecast].Date;
                                                break;
                                            case "ThuMua":
                                                dr["Tên ThuMua " + _Region] = _Supplier.SupplierName;
                                                dr["Đáp ứng từ ThuMua " + _Region] = _SupplierForecast.QuantityForecast;
                                                dr["Ngày sơ chế ThuMua " + _Region] = coreStructure.dicCoord[DatePO][_Product][_CustomerOrder][_SupplierForecast].Date;
                                                //dr["Giá mua ThuMua " + _Region] = 0;
                                                break;
                                            case "VCM":
                                                dr["Tên ThuMua " + _Region] = "VCM - " + _Supplier.SupplierName;
                                                dr["Đáp ứng từ ThuMua " + _Region] = _SupplierForecast.QuantityForecast;
                                                dr["Ngày sơ chế ThuMua " + _Region] = coreStructure.dicCoord[DatePO][_Product][_CustomerOrder][_SupplierForecast].Date;
                                                break;
                                            default:
                                                break;
                                        }

                                        #region OldStuff
                                        //dr["Tên VinEco MB"] =
                                        //dr["Đáp ứng từ VinEco MB"] =
                                        //dr["Ngày sơ chế VinEco MB"] =
                                        //dr["Tên VinEco MN"] =
                                        //dr["Đáp ứng từ VinEco MN"] =
                                        //dr["Ngày sơ chế VinEco MN"] =
                                        //dr["Tên VinEco LĐ"] =
                                        //dr["Đáp ứng từ VinEco LĐ"] =
                                        //dr["Ngày sơ chế VinEco LĐ"] =
                                        //dr["Tên ThuMua MB"] =
                                        //dr["Đáp ứng từ ThuMua MB"] =
                                        //dr["Ngày sơ chế ThuMua MB"] =
                                        //dr["Giá mua ThuMua MB"] =
                                        //dr["Tên ThuMua MN"] =
                                        //dr["Đáp ứng ThuMua MN"] =
                                        //dr["Ngày sơ chế ThuMua MN"] =
                                        //dr["Giá mua ThuMua MN"] =
                                        //dr["Tên ThuMua LĐ"] =
                                        //dr["Đáp ứng từ ThuMua LĐ"] =
                                        //dr["Ngày sơ chế ThuMua LĐ"] =
                                        //dr["Giá mua ThuMua LĐ"] = 0;

                                        //dr["VE Code"] = _Product.ProductCode;
                                        //dr["VE Name"] = _Product.ProductName;

                                        //dr["StoreCode"] = _Customer.CustomerCode;
                                        //dr["StoreName"] = _Customer.CustomerName;
                                        //dr["StoreType"] = _Customer.CustomerType;
                                        //dr["StoreRegion"] = _Customer.CustomerRegion;
                                        //dr["StoreBigRegion"] = _Customer.CustomerBigRegion;
                                        //dr["DateOrder"] = DatePO.Date;
                                        //dr["QuantityOrder"] = _CustomerOrder.QuantityOrder;
                                        //dr["SupplierCode"] = _Supplier.SupplierCode;
                                        //dr["SupplierName"] = _Supplier.SupplierName;
                                        //dr["SupplierRegion"] = _Supplier.SupplierRegion;
                                        //dr["SupplierType"] = _Supplier.SupplierType;
                                        //dr["QuantityForecast"] = _SupplierForecast.QuantityForecast;
                                        //dr["DateForecast"] = coreStructure.dicCoord[DatePO][_Product][_CustomerOrder][_SupplierForecast].Date;
                                        #endregion

                                        dtMastah.Rows.Add(dr);
                                    }
                                }
                                else
                                {
                                    DataRow dr = dtMastah.NewRow();

                                    Customer _Customer = coreStructure.dicCustomer[_CustomerOrder.CustomerId];
                                    //Supplier _Supplier =coreStructure. dicSupplier[_SupplierForecast.SupplierId];
                                    ProductUnit _ProductUnit = null;
                                    ProductUnitRegion _ProductUnitRegion = null;
                                    if (coreStructure.dicProductUnit.TryGetValue(_Product.ProductCode, out _ProductUnit))
                                    {
                                        _ProductUnitRegion = coreStructure.dicProductUnit[_Product.ProductCode].ListRegion.Where(x => x.OrderUnitType == _CustomerOrder.Unit).FirstOrDefault();
                                        if (_ProductUnitRegion == null)
                                        {
                                            _ProductUnitRegion = new ProductUnitRegion()
                                            {
                                                OrderUnitType = "Kg",
                                                OrderUnitPer = 1,
                                                SaleUnitType = "Kg",
                                                SaleUnitPer = 1
                                            };
                                        }
                                    }
                                    else
                                    {
                                        _ProductUnitRegion = new ProductUnitRegion()
                                        {
                                            OrderUnitType = "Kg",
                                            OrderUnitPer = 1,
                                            SaleUnitType = "Kg",
                                            SaleUnitPer = 1
                                        };
                                    }

                                    dr["Mã 6 ký tự"] = _Product.ProductCode;
                                    dr["Mã thành phẩm VinEco"] = _Product.ProductCode;
                                    dr["Mã thành phẩm VinCommerce"] = "";
                                    dr["Tên Sản phẩm"] = _Product.ProductName;
                                    dr["Mã Cửa hàng"] = _Customer.CustomerCode;
                                    dr["Tên Cửa hàng"] = _Customer.CustomerName;
                                    dr["Loại Cửa hàng"] = _Customer.CustomerType;
                                    dr["Ngày Tiêu thụ"] = DatePO.Date;
                                    dr["Vùng Tiêu thụ"] = _Customer.CustomerBigRegion == "Miền Bắc" ? "MB" : "MN";
                                    dr["Nhu cầu Kg VinCommerce"] = _CustomerOrder.QuantityOrderKg;

                                    dr["Số lượng đặt"] = _CustomerOrder.QuantityOrder;
                                    dr["Đơn vị đặt"] = _CustomerOrder.Unit;
                                    dr["Đặt Kg/Unit"] = _ProductUnitRegion.OrderUnitPer;

                                    //dr["Số lượng bán"] = (double)_SupplierForecast.QuantityForecast / (double)_ProductUnitRegion.SaleUnitPer;
                                    dr["Số lượng bán"] = String.Format("= N{0} * Q{0}", dtMastah.Rows.Count + 6);
                                    dr["Đơn vị bán"] = _ProductUnitRegion.SaleUnitType;
                                    dr["Bán Kg/Unit"] = _ProductUnitRegion.SaleUnitPer;

                                    //dr["Nhu cầu Đã đáp ứng"] = String.Format("=SUM(M{0}, P{0}, S{0}, V{0}, Z{0}, AD{0})", dtMastah.Rows.Count + 6);
                                    dr["Nhu cầu Kg Đã đáp ứng"] = String.Format("=SUM( S{0}, V{0}, Y{0}, AB{0}, AF{0}, AJ{0} )", dtMastah.Rows.Count + 6);
                                    //dr["Nhu cầu Kg Đã đáp ứng"] = (double)dr["Nhu cầu Kg Đã đáp ứng"] + _CustomerOrder.QuantityOrderKg;

                                    dtMastah.Rows.Add(dr);
                                }
                            }

                        }
                    }
                    #endregion

                    #region LeftOverVinEco Table
                    DataTable dtLeftOverVE = new DataTable();

                    dtLeftOverVE.Columns.Add("Mã 6 ký tự", typeof(string)).DefaultValue = "";
                    dtLeftOverVE.Columns.Add("Mã thành phẩm VinEco", typeof(string)).DefaultValue = "";
                    dtLeftOverVE.Columns.Add("Mã thành phẩm VinCommerce", typeof(string)).DefaultValue = "";
                    dtLeftOverVE.Columns.Add("Tên Sản phẩm", typeof(string)).DefaultValue = "";
                    dtLeftOverVE.Columns.Add("Mã Cửa hàng", typeof(string)).DefaultValue = "";
                    dtLeftOverVE.Columns.Add("Tên Cửa hàng", typeof(string)).DefaultValue = "";
                    dtLeftOverVE.Columns.Add("Loại Cửa hàng", typeof(string)).DefaultValue = "";
                    dtLeftOverVE.Columns.Add("Ngày Tiêu thụ", typeof(DateTime));
                    dtLeftOverVE.Columns.Add("Vùng Tiêu thụ", typeof(string)).DefaultValue = "";
                    dtLeftOverVE.Columns.Add("Nhu cầu VinCommerce", typeof(double)).DefaultValue = 0;
                    dtLeftOverVE.Columns.Add("Nhu cầu Đã đáp ứng", typeof(string)).DefaultValue = "";
                    dtLeftOverVE.Columns.Add("Tên VinEco MB", typeof(string)).DefaultValue = "";
                    dtLeftOverVE.Columns.Add("Đáp ứng từ VinEco MB", typeof(double)).DefaultValue = 0;
                    dtLeftOverVE.Columns.Add("Ngày sơ chế VinEco MB", typeof(DateTime));
                    dtLeftOverVE.Columns.Add("Tên VinEco MN", typeof(string)).DefaultValue = "";
                    dtLeftOverVE.Columns.Add("Đáp ứng từ VinEco MN", typeof(double)).DefaultValue = 0;
                    dtLeftOverVE.Columns.Add("Ngày sơ chế VinEco MN", typeof(DateTime));
                    dtLeftOverVE.Columns.Add("Tên VinEco LĐ", typeof(string)).DefaultValue = "";
                    dtLeftOverVE.Columns.Add("Đáp ứng từ VinEco LĐ", typeof(double)).DefaultValue = 0;
                    dtLeftOverVE.Columns.Add("Ngày sơ chế VinEco LĐ", typeof(DateTime));
                    dtLeftOverVE.Columns.Add("Tên ThuMua MB", typeof(string)).DefaultValue = "";
                    dtLeftOverVE.Columns.Add("Đáp ứng từ ThuMua MB", typeof(double)).DefaultValue = 0;
                    dtLeftOverVE.Columns.Add("Ngày sơ chế ThuMua MB", typeof(DateTime));
                    dtLeftOverVE.Columns.Add("Giá mua ThuMua MB", typeof(double)).DefaultValue = 0;
                    dtLeftOverVE.Columns.Add("Tên ThuMua MN", typeof(string)).DefaultValue = "";
                    dtLeftOverVE.Columns.Add("Đáp ứng từ ThuMua MN", typeof(double)).DefaultValue = 0;
                    dtLeftOverVE.Columns.Add("Ngày sơ chế ThuMua MN", typeof(DateTime));
                    dtLeftOverVE.Columns.Add("Giá mua ThuMua MN", typeof(double)).DefaultValue = 0;
                    dtLeftOverVE.Columns.Add("Tên ThuMua LĐ", typeof(string)).DefaultValue = "";
                    dtLeftOverVE.Columns.Add("Đáp ứng từ ThuMua LĐ", typeof(double)).DefaultValue = 0;
                    dtLeftOverVE.Columns.Add("Ngày sơ chế ThuMua LĐ", typeof(DateTime));
                    dtLeftOverVE.Columns.Add("Giá mua ThuMua LĐ", typeof(double)).DefaultValue = 0;

                    foreach (DateTime DateFC in coreStructure.dicFC.Keys)
                    {
                        foreach (Product _Product in coreStructure.dicFC[DateFC].Keys.OrderBy(x => x.ProductCode))
                        {
                            var _ListSupplier = coreStructure.dicFC[DateFC][_Product].Keys.Where(x => coreStructure.dicSupplier[x.SupplierId].SupplierType == "VinEco");
                            if (_ListSupplier != null)
                            {
                                foreach (SupplierForecast _SupplierForecast in _ListSupplier.Where(x => x.QuantityForecast >= 3).OrderBy(x => coreStructure.dicSupplier[x.SupplierId].SupplierName))
                                {
                                    DataRow dr = dtLeftOverVE.NewRow();

                                    //Customer _Customer =coreStructure. dicCustomer[_CustomerOrder.CustomerId];
                                    Supplier _Supplier = coreStructure.dicSupplier[_SupplierForecast.SupplierId];

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
                                    dr["Nhu cầu Đã đáp ứng"] = String.Format("=SUM(M{0}, P{0}, S{0}, V{0}, Z{0}, AD{0})", dtLeftOverVE.Rows.Count + 6);

                                    string _Region = string.Join(String.Empty, _Supplier.SupplierRegion.Where((ch, index) => ch != ' ' && (index == 0 || _Supplier.SupplierRegion[index - 1] == ' ')));
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
                        }
                    }
                    #endregion

                    #region LeftOverThuMua Table
                    DataTable dtLeftOverTM = new DataTable();

                    dtLeftOverTM.Columns.Add("Mã 6 ký tự", typeof(string)).DefaultValue = "";
                    dtLeftOverTM.Columns.Add("Mã thành phẩm VinEco", typeof(string)).DefaultValue = "";
                    dtLeftOverTM.Columns.Add("Mã thành phẩm VinCommerce", typeof(string)).DefaultValue = "";
                    dtLeftOverTM.Columns.Add("Tên Sản phẩm", typeof(string)).DefaultValue = "";
                    dtLeftOverTM.Columns.Add("Mã Cửa hàng", typeof(string)).DefaultValue = "";
                    dtLeftOverTM.Columns.Add("Tên Cửa hàng", typeof(string)).DefaultValue = "";
                    dtLeftOverTM.Columns.Add("Loại Cửa hàng", typeof(string)).DefaultValue = "";
                    dtLeftOverTM.Columns.Add("Ngày Tiêu thụ", typeof(DateTime));
                    dtLeftOverTM.Columns.Add("Vùng Tiêu thụ", typeof(string)).DefaultValue = "";
                    dtLeftOverTM.Columns.Add("Nhu cầu VinCommerce", typeof(double)).DefaultValue = 0;
                    dtLeftOverTM.Columns.Add("Nhu cầu Đã đáp ứng", typeof(double)).DefaultValue = 0;
                    dtLeftOverTM.Columns.Add("Tên VinEco MB", typeof(string)).DefaultValue = "";
                    dtLeftOverTM.Columns.Add("Đáp ứng từ VinEco MB", typeof(double)).DefaultValue = 0;
                    dtLeftOverTM.Columns.Add("Ngày sơ chế VinEco MB", typeof(DateTime));
                    dtLeftOverTM.Columns.Add("Tên VinEco MN", typeof(string)).DefaultValue = "";
                    dtLeftOverTM.Columns.Add("Đáp ứng từ VinEco MN", typeof(double)).DefaultValue = 0;
                    dtLeftOverTM.Columns.Add("Ngày sơ chế VinEco MN", typeof(DateTime));
                    dtLeftOverTM.Columns.Add("Tên VinEco LĐ", typeof(string)).DefaultValue = "";
                    dtLeftOverTM.Columns.Add("Đáp ứng từ VinEco LĐ", typeof(double)).DefaultValue = 0;
                    dtLeftOverTM.Columns.Add("Ngày sơ chế VinEco LĐ", typeof(DateTime));
                    dtLeftOverTM.Columns.Add("Tên ThuMua MB", typeof(string)).DefaultValue = "";
                    dtLeftOverTM.Columns.Add("Đáp ứng từ ThuMua MB", typeof(double)).DefaultValue = 0;
                    dtLeftOverTM.Columns.Add("Ngày sơ chế ThuMua MB", typeof(DateTime));
                    dtLeftOverTM.Columns.Add("Giá mua ThuMua MB", typeof(double)).DefaultValue = 0;
                    dtLeftOverTM.Columns.Add("Tên ThuMua MN", typeof(string)).DefaultValue = "";
                    dtLeftOverTM.Columns.Add("Đáp ứng từ ThuMua MN", typeof(double)).DefaultValue = 0;
                    dtLeftOverTM.Columns.Add("Ngày sơ chế ThuMua MN", typeof(DateTime));
                    dtLeftOverTM.Columns.Add("Giá mua ThuMua MN", typeof(double)).DefaultValue = 0;
                    dtLeftOverTM.Columns.Add("Tên ThuMua LĐ", typeof(string)).DefaultValue = "";
                    dtLeftOverTM.Columns.Add("Đáp ứng từ ThuMua LĐ", typeof(double)).DefaultValue = 0;
                    dtLeftOverTM.Columns.Add("Ngày sơ chế ThuMua LĐ", typeof(DateTime));
                    dtLeftOverTM.Columns.Add("Giá mua ThuMua LĐ", typeof(double)).DefaultValue = 0;

                    foreach (DateTime DateFC in coreStructure.dicFC.Keys.OrderBy(x => x.Date))
                    {
                        foreach (Product _Product in coreStructure.dicFC[DateFC].Keys.OrderBy(x => x.ProductCode))
                        {
                            var _ListSupplier = coreStructure.dicFC[DateFC][_Product].Keys.Where(x => coreStructure.dicSupplier[x.SupplierId].SupplierType == "ThuMua");
                            if (_ListSupplier != null)
                            {
                                foreach (SupplierForecast _SupplierForecast in _ListSupplier.Where(x => x.QuantityForecast >= 3).OrderBy(x => coreStructure.dicSupplier[x.SupplierId].SupplierName))
                                {
                                    DataRow dr = dtLeftOverTM.NewRow();

                                    //Customer _Customer =coreStructure. dicCustomer[_CustomerOrder.CustomerId];
                                    Supplier _Supplier = coreStructure.dicSupplier[_SupplierForecast.SupplierId];

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
                                    dr["Nhu cầu Đã đáp ứng"] = _SupplierForecast.QuantityForecast;

                                    string _Region = string.Join(String.Empty, _Supplier.SupplierRegion.Where((ch, index) => ch != ' ' && (index == 0 || _Supplier.SupplierRegion[index - 1] == ' ')));
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

                                    dtLeftOverTM.Rows.Add(dr);
                                }
                            }
                        }
                    }
                    #endregion

                    #region Customer Table
                    DataTable dtCustomer = new DataTable();

                    dtCustomer.Columns.Add("Mã cửa hàng", typeof(string));
                    dtCustomer.Columns.Add("Vùng đặt hàng", typeof(string));
                    dtCustomer.Columns.Add("Loại cửa hàng", typeof(string));
                    dtCustomer.Columns.Add("Vùng tiêu thụ", typeof(string));
                    dtCustomer.Columns.Add("Tên cửa hàng", typeof(string));

                    foreach (Customer _Customer in coreStructure.dicCustomer.Values)
                    {
                        DataRow dr = dtCustomer.NewRow();

                        dr["Mã cửa hàng"] = _Customer.CustomerCode.ToString();
                        dr["Tên cửa hàng"] = _Customer.CustomerName;
                        dr["Loại cửa hàng"] = _Customer.CustomerType.ToString();
                        dr["Vùng tiêu thụ"] = _Customer.CustomerRegion;
                        dr["Vùng đặt hàng"] = _Customer.CustomerBigRegion;

                        dtCustomer.Rows.Add(dr);

                    }
                    #endregion

                    #region Output to Excel
                    Excel.Application xlApp = new Excel.Application();

                    xlApp.ScreenUpdating = false;
                    xlApp.EnableEvents = false;
                    xlApp.DisplayAlerts = false;
                    xlApp.DisplayStatusBar = false;
                    xlApp.AskToUpdateLinks = false;

                    //db.DropCollection("CoordResult");
                    //await db.GetCollection<CoordResult>("CoordResult").InsertManyAsync(CoordResultList);

                    //CoordResultList = null;

                    string filePath = Application.StartupPath.Replace("\\bin\\Debug", "") + "\\Template\\{0}";
                    string fileFullPath = string.Format(filePath, "ChiaHang Mastah.xlsb");
                    string fileFullPath2007 = string.Format(filePath, "ChiaHang Mastah.xlsm");
                    //Debug.WriteLine(filePath);
                    //Debug.WriteLine(fileFullPath);

                    Excel.Workbook xlWb = xlApp.Workbooks.Open(
                        Filename: fileFullPath,
                        UpdateLinks: false,
                        ReadOnly: false,
                        Format: 5,
                        Password: "",
                        WriteResPassword: "",
                        IgnoreReadOnlyRecommended: true,
                        Origin: Excel.XlPlatform.xlWindows,
                        Delimiter: "",
                        Editable: true,
                        Notify: false,
                        Converter: 0,
                        AddToMru: true,
                        Local: false,
                        CorruptLoad: false);

                    xlApp.Calculation = Excel.XlCalculation.xlCalculationManual;

                    var missing = Type.Missing;
                    string path = string.Format(@"C:\Users\Shirayuki\Documents\VinEco\Project - Chia Hang\Mastah Project\Test\ChiaHang Mastah {0}.xlsb", DateFrom.AddDays(3).ToString("dd.MM") + " - " + DateTo.AddDays(-3).ToString("dd.MM") + " (" + DateTime.Now.ToString("yyyyMMdd HH\\hmm") + ")");
                    //string path2007 = string.Format(@"C:\Users\Shirayuki\Documents\VinEco\Project - Chia Hang\Mastah Project\Test\ChiaHang Mastah {0}.xlsm", DateFrom.AddDays(3).ToString("dd.MM") + " - " + DateTo.AddDays(-3).ToString("dd.MM") + " (" + DateTime.Now.ToString("yyyyMMdd HH\\hmm") + ")");
                    Debug.WriteLine(path);
                    xlWb.SaveAs(path, Excel.XlFileFormat.xlExcel12, missing, missing, false, false, Excel.XlSaveAsAccessMode.xlNoChange, missing, missing, missing);

                    OutputExcel(dtMastah, "Mastah", xlWb);
                    OutputExcel(dtLeftOverVE, "DBSL dư", xlWb);
                    OutputExcel(dtLeftOverTM, "Cam Kết dư", xlWb);
                    OutputExcel(dtVeFarm, "VE Farm", xlWb, true, 1, false);
                    OutputExcel(dtThuMua, "VE ThuMua", xlWb, true, 1, false);
                    OutputExcel(dtCustomer, "Region I guess", xlWb, true, 1, false);

                    // Date stuff
                    //xlWb.Worksheets["Mastah"].Cells[2, 3].Value = DateTo.Date - DateFrom.Date;
                    xlWb.Worksheets["Mastah"].Cells[2, 3].Value = (int)(DateFrom.Date - new DateTime(1900, 1, 1)).TotalDays + 2 + 3;
                    xlWb.Worksheets["Mastah"].Cells[3, 3].Value = (int)((DateFrom > DateTo ? DateFrom : DateTo).Date - new DateTime(1900, 1, 1)).TotalDays + 2 - 3;

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

                    xlWb.Worksheets["Mastah"].Cells[4, 1].Formula = String.Format("=SUBTOTAL(3,A6:A{0})", dtMastah.Rows.Count + 5); // A
                    xlWb.Worksheets["Mastah"].Cells[4, 5].Formula = String.Format("=SUBTOTAL(3,E6:E{0})", dtMastah.Rows.Count + 5); // E
                    xlWb.Worksheets["Mastah"].Cells[4, 10].Formula = String.Format("=SUBTOTAL(9,J6:J{0})", dtMastah.Rows.Count + 5); // J
                    xlWb.Worksheets["Mastah"].Cells[4, 11].Formula = String.Format("=SUBTOTAL(9,K6:K{0})", dtMastah.Rows.Count + 5); // K
                    xlWb.Worksheets["Mastah"].Cells[4, 12].Formula = String.Format("=SUBTOTAL(3,L6:L{0})", dtMastah.Rows.Count + 5); // L
                    xlWb.Worksheets["Mastah"].Cells[4, 14].Formula = String.Format("=SUBTOTAL(9,N6:N{0})", dtMastah.Rows.Count + 5); // N
                    xlWb.Worksheets["Mastah"].Cells[4, 15].Formula = String.Format("=SUBTOTAL(9,O6:O{0})", dtMastah.Rows.Count + 5); // O
                    xlWb.Worksheets["Mastah"].Cells[4, 16].Formula = String.Format("=SUBTOTAL(3,P6:P{0})", dtMastah.Rows.Count + 5); // P
                    xlWb.Worksheets["Mastah"].Cells[4, 18].Formula = String.Format("=SUBTOTAL(3,R6:R{0})", dtMastah.Rows.Count + 5); // R
                    xlWb.Worksheets["Mastah"].Cells[4, 19].Formula = String.Format("=SUBTOTAL(9,S6:S{0})", dtMastah.Rows.Count + 5); // S
                    xlWb.Worksheets["Mastah"].Cells[4, 21].Formula = String.Format("=SUBTOTAL(3,U6:U{0})", dtMastah.Rows.Count + 5); // U
                    xlWb.Worksheets["Mastah"].Cells[4, 22].Formula = String.Format("=SUBTOTAL(9,V6:V{0})", dtMastah.Rows.Count + 5); // V
                    xlWb.Worksheets["Mastah"].Cells[4, 24].Formula = String.Format("=SUBTOTAL(3,X6:X{0})", dtMastah.Rows.Count + 5); // X
                    xlWb.Worksheets["Mastah"].Cells[4, 25].Formula = String.Format("=SUBTOTAL(9,Y6:Y{0})", dtMastah.Rows.Count + 5); // Y
                    xlWb.Worksheets["Mastah"].Cells[4, 27].Formula = String.Format("=SUBTOTAL(3,AA6:AA{0})", dtMastah.Rows.Count + 5); // AA
                    xlWb.Worksheets["Mastah"].Cells[4, 28].Formula = String.Format("=SUBTOTAL(9,AB6:AB{0})", dtMastah.Rows.Count + 5); // AB
                    xlWb.Worksheets["Mastah"].Cells[4, 31].Formula = String.Format("=SUBTOTAL(3,AE6:AE{0})", dtMastah.Rows.Count + 5); // AE
                    xlWb.Worksheets["Mastah"].Cells[4, 32].Formula = String.Format("=SUBTOTAL(9,AF6:AF{0})", dtMastah.Rows.Count + 5); // AF
                    xlWb.Worksheets["Mastah"].Cells[4, 35].Formula = String.Format("=SUBTOTAL(3,AI6:AI{0})", dtMastah.Rows.Count + 5); // AI
                    xlWb.Worksheets["Mastah"].Cells[4, 36].Formula = String.Format("=SUBTOTAL(9,AJ6:AJ{0})", dtMastah.Rows.Count + 5); // AJ

                    // Formula Stuff for Leftover VE
                    xlWb.Worksheets["DBSL Dư"].Cells[4, 1].Formula = String.Format("=SUBTOTAL(3,A6:A{0}", dtLeftOverVE.Rows.Count + 5); // A4

                    xlWb.Worksheets["DBSL Dư"].Cells[4, 5].Formula = String.Format("=SUBTOTAL(3,E6:E{0}", dtLeftOverVE.Rows.Count + 5); // E4

                    xlWb.Worksheets["DBSL Dư"].Cells[4, 9].Formula = String.Format("=SUBTOTAL(3,I6:I{0}", dtLeftOverVE.Rows.Count + 5); // I4
                    xlWb.Worksheets["DBSL Dư"].Cells[4, 10].Formula = String.Format("=SUBTOTAL(9,J6:J{0}", dtLeftOverVE.Rows.Count + 5); // J4
                    xlWb.Worksheets["DBSL Dư"].Cells[4, 11].Formula = String.Format("=SUBTOTAL(9,K6:K{0}", dtLeftOverVE.Rows.Count + 5); // K4
                    xlWb.Worksheets["DBSL Dư"].Cells[4, 12].Formula = String.Format("=SUBTOTAL(3,L6:L{0}", dtLeftOverVE.Rows.Count + 5); // L4
                    xlWb.Worksheets["DBSL Dư"].Cells[4, 13].Formula = String.Format("=SUBTOTAL(9,M6:M{0}", dtLeftOverVE.Rows.Count + 5); // M4

                    xlWb.Worksheets["DBSL Dư"].Cells[4, 15].Formula = String.Format("=SUBTOTAL(3,O6:O{0}", dtLeftOverVE.Rows.Count + 5); // O4
                    xlWb.Worksheets["DBSL Dư"].Cells[4, 16].Formula = String.Format("=SUBTOTAL(9,P6:P{0}", dtLeftOverVE.Rows.Count + 5); // P4

                    xlWb.Worksheets["DBSL Dư"].Cells[4, 18].Formula = String.Format("=SUBTOTAL(3,R6:R{0}", dtLeftOverVE.Rows.Count + 5); // R4
                    xlWb.Worksheets["DBSL Dư"].Cells[4, 19].Formula = String.Format("=SUBTOTAL(9,S6:S{0}", dtLeftOverVE.Rows.Count + 5); // S4

                    xlWb.Worksheets["DBSL Dư"].Cells[4, 21].Formula = String.Format("=SUBTOTAL(3,U6:U{0}", dtLeftOverVE.Rows.Count + 5); // U4
                    xlWb.Worksheets["DBSL Dư"].Cells[4, 22].Formula = String.Format("=SUBTOTAL(9,V6:V{0}", dtLeftOverVE.Rows.Count + 5); // V4

                    xlWb.Worksheets["DBSL Dư"].Cells[4, 25].Formula = String.Format("=SUBTOTAL(3,Y6:Y{0}", dtLeftOverVE.Rows.Count + 5); // Y4
                    xlWb.Worksheets["DBSL Dư"].Cells[4, 26].Formula = String.Format("=SUBTOTAL(9,Z6:Z{0}", dtLeftOverVE.Rows.Count + 5); // Z4

                    xlWb.Worksheets["DBSL Dư"].Cells[4, 29].Formula = String.Format("=SUBTOTAL(3,AC6:AC{0}", dtLeftOverVE.Rows.Count + 5); // AC4
                    xlWb.Worksheets["DBSL Dư"].Cells[4, 30].Formula = String.Format("=SUBTOTAL(9,AD6:AD{0}", dtLeftOverVE.Rows.Count + 5); // AD4



                    //using (ExcelPackage pck = new ExcelPackage(new FileInfo(fileFullPath2007)))
                    //{
                    //    OutputExcelEpplus(pck, dtMastah, "Mastah", false, 6, false);

                    //    pck.SaveAs(new FileInfo(path2007));
                    //}

                    //xlApp.Calculation = Excel.XlCalculation.xlCalculationAutomatic;

                    xlWb.Close(SaveChanges: true);
                    if (xlWb != null) { Marshal.ReleaseComObject(xlWb); }
                    xlWb = null;

                    xlApp.ScreenUpdating = true;
                    xlApp.EnableEvents = true;
                    xlApp.DisplayAlerts = false;
                    xlApp.DisplayStatusBar = true;
                    xlApp.AskToUpdateLinks = true;

                    xlApp.Quit();
                    if (xlApp != null) { Marshal.ReleaseComObject(xlApp); }
                    xlApp = null;
                    #endregion

                }
                else if (!YesNoGroupFarm)
                {
                    DataTable dtMastah = new DataTable();
                    dtMastah.TableName = "Mastah";

                    if (YesNoGroupThuMua)
                    {
                        #region Mastah Table
                        dtMastah.Columns.Add("Mã 6 ký tự", typeof(string)).DefaultValue = "";
                        dtMastah.Columns.Add("Tên sản phẩm", typeof(string)).DefaultValue = "";
                        dtMastah.Columns.Add("Loại cửa hàng", typeof(string)).DefaultValue = "";
                        dtMastah.Columns.Add("Ngày tiêu thụ", typeof(int)).DefaultValue = 0;
                        dtMastah.Columns.Add("Vùng tiêu thụ", typeof(string)).DefaultValue = "";
                        dtMastah.Columns.Add("Nhu cầu VinCommerce", typeof(double)).DefaultValue = 0;
                        dtMastah.Columns.Add("Nhu cầu Đáp ứng", typeof(double)).DefaultValue = 0;
                        dtMastah.Columns.Add("Nguồn", typeof(string)).DefaultValue = "";
                        dtMastah.Columns.Add("Vùng sản xuất", typeof(string)).DefaultValue = "";
                        dtMastah.Columns.Add("Tên NCC", typeof(string)).DefaultValue = "";
                        dtMastah.Columns.Add("Ngày sơ chế", typeof(int));

                        dicRow = new Dictionary<string, int>();
                        foreach (DateTime DatePO in coreStructure.dicCoord.Keys)
                        {
                            foreach (Product _Product in coreStructure.dicCoord[DatePO].Keys.OrderBy(x => x.ProductCode))
                            {
                                foreach (CustomerOrder _CustomerOrder in coreStructure.dicCoord[DatePO][_Product].Keys.OrderBy(x => coreStructure.dicCustomer[x.CustomerId].CustomerType))
                                {
                                    Customer _Customer = coreStructure.dicCustomer[_CustomerOrder.CustomerId];
                                    string sKey = String.Format("{0}{1}{2}{3}", DatePO.Date, _Customer.CustomerType, _Customer.CustomerBigRegion, _Product.ProductCode);
                                    if (coreStructure.dicCoord[DatePO][_Product][_CustomerOrder] != null)
                                    {
                                        foreach (SupplierForecast _SupplierForecast in coreStructure.dicCoord[DatePO][_Product][_CustomerOrder].Keys.OrderBy(x => coreStructure.dicSupplier[x.SupplierId].SupplierType))
                                        {
                                            Supplier _Supplier = coreStructure.dicSupplier[_SupplierForecast.SupplierId];

                                            sKey += _Supplier.SupplierType == "ThuMua" ? "ThuMua" : _Supplier.SupplierType;

                                            DataRow dr = null;
                                            int _rowIndex = 0;
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

                                            string _Region = string.Join(String.Empty, _Supplier.SupplierRegion.Where((ch, index) => ch != ' ' && (index == 0 || _Supplier.SupplierRegion[index - 1] == ' ')));
                                            string _colName = string.Format("{0} {1}", _Supplier.SupplierType, _Region);

                                            dr["Mã 6 ký tự"] = _Product.ProductCode;
                                            dr["Tên sản phẩm"] = _Product.ProductName;
                                            dr["Loại cửa hàng"] = _Customer.CustomerType;
                                            dr["Ngày tiêu thụ"] = (int)(DatePO.Date - new DateTime(1900, 1, 1)).TotalDays + 2;
                                            dr["Vùng tiêu thụ"] = _Customer.CustomerBigRegion;
                                            dr["Nhu cầu VinCommerce"] = Convert.ToDouble(dr["Nhu cầu VinCommerce"]) + _CustomerOrder.QuantityOrder;
                                            dr["Nhu cầu Đáp ứng"] = Convert.ToDouble(dr["Nhu cầu Đáp ứng"]) + _SupplierForecast.QuantityForecast; ;
                                            dr["Nguồn"] = _Supplier.SupplierType;
                                            dr["Vùng sản xuất"] = _Supplier.SupplierRegion;
                                            dr["Tên NCC"] = _Supplier.SupplierType == "ThuMua" ? "ThuMua" : _Supplier.SupplierName;
                                            dr["Ngày sơ chế"] = (int)(coreStructure.dicCoord[DatePO][_Product][_CustomerOrder][_SupplierForecast].Date - new DateTime(1900, 1, 1)).TotalDays + 2;
                                        }
                                    }
                                    else
                                    {
                                        DataRow dr = null;
                                        int _rowIndex = 0;
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
                                        dr["Ngày tiêu thụ"] = (int)(DatePO.Date - new DateTime(1900, 1, 1)).TotalDays + 2;
                                        dr["Vùng tiêu thụ"] = _Customer.CustomerBigRegion;
                                        dr["Nhu cầu VinCommerce"] = Convert.ToDouble(dr["Nhu cầu VinCommerce"]) + _CustomerOrder.QuantityOrder;
                                    }
                                }
                            }
                        }
                        #endregion
                    }
                    else
                    {
                        #region Mastah Table
                        dtMastah.Columns.Add("Mã 6 ký tự", typeof(string)).DefaultValue = "";
                        dtMastah.Columns.Add("Tên sản phẩm", typeof(string)).DefaultValue = "";
                        dtMastah.Columns.Add("Loại cửa hàng", typeof(string)).DefaultValue = "";
                        dtMastah.Columns.Add("Ngày tiêu thụ", typeof(int)).DefaultValue = 0;
                        dtMastah.Columns.Add("Vùng tiêu thụ", typeof(string)).DefaultValue = "";
                        dtMastah.Columns.Add("Nhu cầu VinCommerce", typeof(double)).DefaultValue = 0;
                        dtMastah.Columns.Add("Nhu cầu Đáp ứng", typeof(double)).DefaultValue = 0;
                        dtMastah.Columns.Add("Nguồn", typeof(string)).DefaultValue = "";
                        dtMastah.Columns.Add("Vùng sản xuất", typeof(string)).DefaultValue = "";
                        dtMastah.Columns.Add("Tên NCC", typeof(string)).DefaultValue = "";
                        dtMastah.Columns.Add("Ngày sơ chế", typeof(int));

                        dicRow = new Dictionary<string, int>();
                        foreach (DateTime DatePO in coreStructure.dicCoord.Keys)
                        {
                            foreach (Product _Product in coreStructure.dicCoord[DatePO].Keys.OrderBy(x => x.ProductCode))
                            {
                                foreach (CustomerOrder _CustomerOrder in coreStructure.dicCoord[DatePO][_Product].Keys.OrderBy(x => coreStructure.dicCustomer[x.CustomerId].CustomerType))
                                {
                                    Customer _Customer = coreStructure.dicCustomer[_CustomerOrder.CustomerId];
                                    string sKey = String.Format("{0}{1}{2}{3}", DatePO.Date, _Customer.CustomerType, _Customer.CustomerBigRegion, _Product.ProductCode);
                                    if (coreStructure.dicCoord[DatePO][_Product][_CustomerOrder] != null)
                                    {
                                        foreach (SupplierForecast _SupplierForecast in coreStructure.dicCoord[DatePO][_Product][_CustomerOrder].Keys.OrderBy(x => coreStructure.dicSupplier[x.SupplierId].SupplierType))
                                        {
                                            Supplier _Supplier = coreStructure.dicSupplier[_SupplierForecast.SupplierId];

                                            sKey += _Supplier.SupplierCode;

                                            DataRow dr = null;
                                            int _rowIndex = 0;
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

                                            string _Region = string.Join(String.Empty, _Supplier.SupplierRegion.Where((ch, index) => ch != ' ' && (index == 0 || _Supplier.SupplierRegion[index - 1] == ' ')));
                                            string _colName = string.Format("{0} {1}", _Supplier.SupplierType, _Region);

                                            dr["Mã 6 ký tự"] = _Product.ProductCode;
                                            dr["Tên sản phẩm"] = _Product.ProductName;
                                            dr["Loại cửa hàng"] = _Customer.CustomerType;
                                            dr["Ngày tiêu thụ"] = (int)(DatePO.Date - new DateTime(1900, 1, 1)).TotalDays + 2;
                                            dr["Vùng tiêu thụ"] = _Customer.CustomerBigRegion;
                                            dr["Nhu cầu VinCommerce"] = (double)(dr["Nhu cầu VinCommerce"]) + _CustomerOrder.QuantityOrder;
                                            dr["Nhu cầu Đáp ứng"] = (double)(dr["Nhu cầu Đáp ứng"]) + _SupplierForecast.QuantityForecast; ;
                                            dr["Nguồn"] = _Supplier.SupplierType;
                                            dr["Vùng sản xuất"] = _Supplier.SupplierRegion;
                                            dr["Tên NCC"] = _Supplier.SupplierName;
                                            dr["Ngày sơ chế"] = (int)(coreStructure.dicCoord[DatePO][_Product][_CustomerOrder][_SupplierForecast].Date - new DateTime(1900, 1, 1)).TotalDays + 2;
                                        }
                                    }
                                    else
                                    {
                                        DataRow dr = null;
                                        int _rowIndex = 0;
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
                                        dr["Ngày tiêu thụ"] = (int)(DatePO.Date - new DateTime(1900, 1, 1)).TotalDays + 2;
                                        dr["Vùng tiêu thụ"] = _Customer.CustomerBigRegion;
                                        dr["Nhu cầu VinCommerce"] = Convert.ToDouble(dr["Nhu cầu VinCommerce"]) + _CustomerOrder.QuantityOrder;
                                    }
                                }
                            }
                        }
                        #endregion
                    }

                    #region LeftoverVinEco
                    DataTable dtLeftoverVe = new DataTable();

                    dtLeftoverVe.TableName = "NoCusVinEco";

                    dtLeftoverVe.Columns.Add("Mã VinEco", typeof(string)).DefaultValue = "";
                    dtLeftoverVe.Columns.Add("Tên VinEco", typeof(string)).DefaultValue = "";
                    dtLeftoverVe.Columns.Add("Mã Farm", typeof(string)).DefaultValue = "";
                    dtLeftoverVe.Columns.Add("Tên Farm", typeof(string)).DefaultValue = "";
                    dtLeftoverVe.Columns.Add("Ngày thu hoạch", typeof(int));
                    dtLeftoverVe.Columns.Add("Vùng sản xuất", typeof(string)).DefaultValue = "";
                    dtLeftoverVe.Columns.Add("Sản lượng", typeof(double)).DefaultValue = 0;

                    foreach (DateTime DateFC in coreStructure.dicFC.Keys)
                    {
                        foreach (Product _Product in coreStructure.dicFC[DateFC].Keys.OrderBy(x => x.ProductCode))
                        {
                            var _ListSupplier = coreStructure.dicFC[DateFC][_Product].Keys.Where(x => coreStructure.dicSupplier[x.SupplierId].SupplierType == "VinEco");
                            if (_ListSupplier != null)
                            {
                                foreach (SupplierForecast _SupplierForecast in _ListSupplier.OrderBy(x => coreStructure.dicSupplier[x.SupplierId].SupplierName))
                                {
                                    if (_SupplierForecast.QuantityForecast > 0)
                                    {
                                        DataRow dr = dtLeftoverVe.NewRow();

                                        //Customer _Customer =coreStructure. dicCustomer[_CustomerOrder.CustomerId];
                                        Supplier _Supplier = coreStructure.dicSupplier[_SupplierForecast.SupplierId];

                                        dr["Mã VinEco"] = _Product.ProductCode;
                                        dr["Tên VinEco"] = _Product.ProductName;
                                        dr["Mã Farm"] = _Supplier.SupplierCode;
                                        dr["Tên Farm"] = _Supplier.SupplierName;
                                        dr["Ngày thu hoạch"] = (int)(DateFC.Date - new DateTime(1900, 1, 1)).TotalDays + 2;
                                        dr["Vùng sản xuất"] = _Supplier.SupplierRegion;
                                        dr["Sản lượng"] = _SupplierForecast.QuantityForecast;

                                        dtLeftoverVe.Rows.Add(dr);
                                    }
                                }
                            }
                        }
                    }
                    #endregion

                    #region LeftoverThuMua
                    DataTable dtLeftoverTm = new DataTable();

                    dtLeftoverTm.TableName = "NoCusThuMua";

                    dtLeftoverTm.Columns.Add("Mã VinEco", typeof(string)).DefaultValue = "";
                    dtLeftoverTm.Columns.Add("Tên VinEco", typeof(string)).DefaultValue = "";
                    dtLeftoverTm.Columns.Add("Mã Farm", typeof(string)).DefaultValue = "";
                    dtLeftoverTm.Columns.Add("Tên Farm", typeof(string)).DefaultValue = "";
                    dtLeftoverTm.Columns.Add("Ngày thu hoạch", typeof(int));
                    dtLeftoverTm.Columns.Add("Vùng sản xuất", typeof(string)).DefaultValue = "";
                    dtLeftoverTm.Columns.Add("Sản lượng", typeof(double)).DefaultValue = 0;

                    foreach (DateTime DateFC in coreStructure.dicFC.Keys)
                    {
                        foreach (Product _Product in coreStructure.dicFC[DateFC].Keys.OrderBy(x => x.ProductCode))
                        {
                            var _ListSupplier = coreStructure.dicFC[DateFC][_Product].Keys.Where(x => coreStructure.dicSupplier[x.SupplierId].SupplierType == "ThuMua");
                            if (_ListSupplier != null)
                            {
                                foreach (SupplierForecast _SupplierForecast in _ListSupplier.OrderBy(x => coreStructure.dicSupplier[x.SupplierId].SupplierName))
                                {
                                    if (_SupplierForecast.QuantityForecast > 0)
                                    {
                                        DataRow dr = dtLeftoverTm.NewRow();

                                        //Customer _Customer =coreStructure. dicCustomer[_CustomerOrder.CustomerId];
                                        Supplier _Supplier = coreStructure.dicSupplier[_SupplierForecast.SupplierId];

                                        dr["Mã VinEco"] = _Product.ProductCode;
                                        dr["Tên VinEco"] = _Product.ProductName;
                                        dr["Mã Farm"] = _Supplier.SupplierCode;
                                        dr["Tên Farm"] = _Supplier.SupplierName;
                                        dr["Ngày thu hoạch"] = (int)(DateFC.Date - new DateTime(1900, 1, 1)).TotalDays + 2;
                                        dr["Vùng sản xuất"] = _Supplier.SupplierRegion;
                                        dr["Sản lượng"] = _SupplierForecast.QuantityForecast;

                                        dtLeftoverTm.Rows.Add(dr);
                                    }
                                }
                            }
                        }
                    }
                    #endregion

                    #region Output to Excel - OpenXMLWriter Style, super fast.
                    string fileName = string.Format("Mastah Compact {0}.xlsx", DateFrom.AddDays(3).ToString("dd.MM") + " - " + DateTo.AddDays(-3).ToString("dd.MM") + " (" + DateTime.Now.ToString("yyyyMMdd HH\\hmm") + ")");
                    string path = string.Format(@"C:\Users\Shirayuki\Documents\VinEco\Project - Chia Hang\Mastah Project\Test\" + fileName);

                    var listDt = new List<DataTable>();

                    listDt.Add(dtMastah);
                    listDt.Add(dtLeftoverVe);
                    listDt.Add(dtLeftoverTm);

                    LargeExportOneWorkbook(path, listDt, true, true);

                    ConvertToXlsbInterop(path, "xlsx", "xlsb", true);

                    //LargeExport(dtMastah, path, true, true);
                    #endregion
                }
                else
                {
                    #region Mastah Table
                    DataTable dtMastah = new DataTable();

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
                    foreach (DateTime DatePO in coreStructure.dicCoord.Keys)
                    {
                        foreach (Product _Product in coreStructure.dicCoord[DatePO].Keys.OrderBy(x => x.ProductCode))
                        {
                            foreach (CustomerOrder _CustomerOrder in coreStructure.dicCoord[DatePO][_Product].Keys.OrderBy(x => coreStructure.dicCustomer[x.CustomerId].CustomerType))
                            {
                                Customer _Customer = coreStructure.dicCustomer[_CustomerOrder.CustomerId];
                                string sKey = String.Format("{0}{1}{2}{3}", DatePO.Date, _Customer.CustomerType, _Customer.CustomerBigRegion, _Product.ProductCode);
                                if (coreStructure.dicCoord[DatePO][_Product][_CustomerOrder] != null)
                                {
                                    foreach (SupplierForecast _SupplierForecast in coreStructure.dicCoord[DatePO][_Product][_CustomerOrder].Keys.OrderBy(x => coreStructure.dicSupplier[x.SupplierId].SupplierType))
                                    {
                                        Supplier _Supplier = coreStructure.dicSupplier[_SupplierForecast.SupplierId];

                                        DataRow dr = null;
                                        int _rowIndex = 0;
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

                                        string _Region = string.Join(String.Empty, _Supplier.SupplierRegion.Where((ch, index) => ch != ' ' && (index == 0 || _Supplier.SupplierRegion[index - 1] == ' ')));
                                        string _colName = string.Format("{0} {1}", _Supplier.SupplierType, _Region);

                                        dr["Mã 6 ký tự"] = _Product.ProductCode;
                                        dr["Tên sản phẩm"] = _Product.ProductName;
                                        dr["Loại cửa hàng"] = _Customer.CustomerType;
                                        dr["Ngày tiêu thụ"] = (int)(DatePO.Date - new DateTime(1900, 1, 1)).TotalDays + 2;
                                        dr["Vùng tiêu thụ"] = _Customer.CustomerBigRegion;
                                        dr["Nhu cầu VinCommerce"] = Convert.ToDouble(dr["Nhu cầu VinCommerce"]) + _CustomerOrder.QuantityOrder;

                                        dr[_colName] = Convert.ToDouble(dr[_colName]) + _SupplierForecast.QuantityForecast;

                                        dr["Tổng " + _Supplier.SupplierType] = Convert.ToDouble(dr["Tổng " + _Supplier.SupplierType]) + _SupplierForecast.QuantityForecast;
                                        dr["Nhu cầu Đáp ứng"] = Convert.ToDouble(dr["Nhu cầu Đáp ứng"]) + _SupplierForecast.QuantityForecast; ;
                                    }
                                }
                                else
                                {
                                    DataRow dr = null;
                                    int _rowIndex = 0;
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
                                    dr["Ngày tiêu thụ"] = (int)(DatePO.Date - new DateTime(1900, 1, 1)).TotalDays + 2;
                                    dr["Vùng tiêu thụ"] = _Customer.CustomerBigRegion;
                                    dr["Nhu cầu VinCommerce"] = Convert.ToDouble(dr["Nhu cầu VinCommerce"]) + _CustomerOrder.QuantityOrder;
                                }
                            }
                        }
                    }
                    #endregion

                    #region LeftoverVinEco
                    DataTable dtLeftoverVe = new DataTable();

                    dtLeftoverVe.TableName = "NoCusVinEco";

                    dtLeftoverVe.Columns.Add("Mã VinEco", typeof(string)).DefaultValue = "";
                    dtLeftoverVe.Columns.Add("Tên VinEco", typeof(string)).DefaultValue = "";
                    dtLeftoverVe.Columns.Add("Mã Farm", typeof(string)).DefaultValue = "";
                    dtLeftoverVe.Columns.Add("Tên Farm", typeof(string)).DefaultValue = "";
                    dtLeftoverVe.Columns.Add("Ngày thu hoạch", typeof(int));
                    dtLeftoverVe.Columns.Add("Vùng sản xuất", typeof(string)).DefaultValue = "";
                    dtLeftoverVe.Columns.Add("Sản lượng", typeof(double)).DefaultValue = 0;

                    foreach (DateTime DateFC in coreStructure.dicFC.Keys)
                    {
                        foreach (Product _Product in coreStructure.dicFC[DateFC].Keys.OrderBy(x => x.ProductCode))
                        {
                            var _ListSupplier = coreStructure.dicFC[DateFC][_Product].Keys.Where(x => coreStructure.dicSupplier[x.SupplierId].SupplierType == "VinEco");
                            if (_ListSupplier != null)
                            {
                                foreach (SupplierForecast _SupplierForecast in _ListSupplier.OrderBy(x => coreStructure.dicSupplier[x.SupplierId].SupplierName))
                                {
                                    if (_SupplierForecast.QuantityForecast > 0)
                                    {
                                        DataRow dr = dtLeftoverVe.NewRow();

                                        //Customer _Customer =coreStructure. dicCustomer[_CustomerOrder.CustomerId];
                                        Supplier _Supplier = coreStructure.dicSupplier[_SupplierForecast.SupplierId];

                                        dr["Mã VinEco"] = _Product.ProductCode;
                                        dr["Tên VinEco"] = _Product.ProductName;
                                        dr["Mã Farm"] = _Supplier.SupplierCode;
                                        dr["Tên Farm"] = _Supplier.SupplierName;
                                        dr["Ngày thu hoạch"] = (int)(DateFC.Date - new DateTime(1900, 1, 1)).TotalDays + 2;
                                        dr["Vùng sản xuất"] = _Supplier.SupplierRegion;
                                        dr["Sản lượng"] = _SupplierForecast.QuantityForecast;

                                        dtLeftoverVe.Rows.Add(dr);
                                    }
                                }
                            }
                        }
                    }
                    #endregion

                    #region LeftoverThuMua
                    DataTable dtLeftoverTm = new DataTable();

                    dtLeftoverTm.TableName = "NoCusThuMua";

                    dtLeftoverTm.Columns.Add("Mã VinEco", typeof(string)).DefaultValue = "";
                    dtLeftoverTm.Columns.Add("Tên VinEco", typeof(string)).DefaultValue = "";
                    dtLeftoverTm.Columns.Add("Mã Farm", typeof(string)).DefaultValue = "";
                    dtLeftoverTm.Columns.Add("Tên Farm", typeof(string)).DefaultValue = "";
                    dtLeftoverTm.Columns.Add("Ngày thu hoạch", typeof(int));
                    dtLeftoverTm.Columns.Add("Vùng sản xuất", typeof(string)).DefaultValue = "";
                    dtLeftoverTm.Columns.Add("Sản lượng", typeof(double)).DefaultValue = 0;

                    foreach (DateTime DateFC in coreStructure.dicFC.Keys)
                    {
                        foreach (Product _Product in coreStructure.dicFC[DateFC].Keys.OrderBy(x => x.ProductCode))
                        {
                            var _ListSupplier = coreStructure.dicFC[DateFC][_Product].Keys.Where(x => coreStructure.dicSupplier[x.SupplierId].SupplierType == "ThuMua");
                            if (_ListSupplier != null)
                            {
                                foreach (SupplierForecast _SupplierForecast in _ListSupplier.OrderBy(x => coreStructure.dicSupplier[x.SupplierId].SupplierName))
                                {
                                    if (_SupplierForecast.QuantityForecast > 0)
                                    {
                                        DataRow dr = dtLeftoverTm.NewRow();

                                        //Customer _Customer =coreStructure. dicCustomer[_CustomerOrder.CustomerId];
                                        Supplier _Supplier = coreStructure.dicSupplier[_SupplierForecast.SupplierId];

                                        dr["Mã VinEco"] = _Product.ProductCode;
                                        dr["Tên VinEco"] = _Product.ProductName;
                                        dr["Mã Farm"] = _Supplier.SupplierCode;
                                        dr["Tên Farm"] = _Supplier.SupplierName;
                                        dr["Ngày thu hoạch"] = (int)(DateFC.Date - new DateTime(1900, 1, 1)).TotalDays + 2;
                                        dr["Vùng sản xuất"] = _Supplier.SupplierRegion;
                                        dr["Sản lượng"] = _SupplierForecast.QuantityForecast;

                                        dtLeftoverTm.Rows.Add(dr);
                                    }
                                }
                            }
                        }
                    }
                    #endregion

                    #region Output to Excel - OpenXMLWriter Style, super fast.
                    string fileName = string.Format("Mastah Compact {0}.xlsx", DateFrom.ToString("dd.MM") + " - " + DateTo.ToString("dd.MM") + " (" + DateTime.Now.ToString("yyyyMMdd HH\\hmm") + ")");
                    string path = string.Format(@"C:\Users\Shirayuki\Documents\VinEco\Project - Chia Hang\Mastah Project\Test\" + fileName);

                    var listDt = new List<DataTable>();

                    listDt.Add(dtMastah);
                    listDt.Add(dtLeftoverVe);
                    listDt.Add(dtLeftoverTm);

                    LargeExportOneWorkbook(path, listDt, true, true);

                    ConvertToXlsbInterop(path, "xlsx", "xlsb", true);

                    //LargeExport(dtMastah, path, true, true);
                    #endregion
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
                Debug.WriteLine(String.Format("Done in {0}s!", Math.Round(stopWatch.Elapsed.TotalSeconds, 2)));

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
        /// The Core of all Algorithm. Where everything begins and ends.
        /// </summary>
        /// <param name="coreStructure"></param>
        /// <param name="SupplierRegion"></param>
        /// <param name="CustomerRegion"></param>
        /// <param name="SupplierType"></param>
        /// <param name="dayBefore"></param>
        /// <param name="dayLdBefore"></param>
        /// <param name="UpperLimit"></param>
        /// <param name="CrossRegion"></param>
        /// <param name="PriorityTarget"></param>
        private void Coord(CoordStructure coreStructure, string SupplierRegion, string CustomerRegion, string SupplierType, byte dayBefore = 0, byte dayLdBefore = 0, float UpperLimit = 1, bool CrossRegion = false, string PriorityTarget = "", bool YesNoByUnit = false)
        {
            try
            {
                /// <* IMPORTANTO! *>
                // Nothing shall begin before this happens
                Stopwatch stopwatch = Stopwatch.StartNew();

                // To deal with uhm, "Everybody has a bite."
                // Make a Dictionary storing the quantity ordered from each Suppliers.
                var dicDeli = new Dictionary<DateTime, Dictionary<Product, Dictionary<SupplierForecast, double>>>();
                foreach (DateTime DateFC in coreStructure.dicFC.Keys)
                {
                    dicDeli.Add(DateFC, new Dictionary<Product, Dictionary<SupplierForecast, double>>());
                    foreach (Product _Product in coreStructure.dicFC[DateFC].Keys)
                    {
                        dicDeli[DateFC].Add(_Product, new Dictionary<SupplierForecast, double>());
                        foreach (SupplierForecast _SupplierForecast in coreStructure.dicFC[DateFC][_Product].Keys)
                        {
                            if (coreStructure.dicSupplier[_SupplierForecast.SupplierId].SupplierType == SupplierType)
                            {
                                // ... And of course, the initialized value is 0.
                                dicDeli[DateFC][_Product].Add(_SupplierForecast, 0);
                            }
                        }
                    }
                }

                // To deal with uhm, OrderQuantity of like, 3 grams. 
                // Who the fuck order 3 grams, seriously.
                var dicMinimum = new Dictionary<string, double>();
                dicMinimum.Add("A", 0.5);
                dicMinimum.Add("B", 0.5);
                dicMinimum.Add("C", 0.5);
                dicMinimum.Add("D", 0.5);
                dicMinimum.Add("E", 0.5);
                dicMinimum.Add("F", 0.2);
                dicMinimum.Add("G", 0.5);
                dicMinimum.Add("H", 0.2);
                dicMinimum.Add("I", 0.2);
                dicMinimum.Add("J", 0.5);
                dicMinimum.Add("K", 0.7);
                dicMinimum.Add("L", 1);
                dicMinimum.Add("M", 1);
                dicMinimum.Add("N", 69);
                dicMinimum.Add("1", 0.01);
                dicMinimum.Add("2", 0.01);

                // PO Date Layer.
                foreach (DateTime DatePO in coreStructure.dicPO.Keys.ToList())
                {
                    // Product Layer.
                    foreach (Product _Product in coreStructure.dicPO[DatePO].Keys.ToList())
                    {
                        /// <! Debuging Purposes !>
                        // Only uncomment in emergency situation.
                        //if (_Product.ProductCode == "H00101")
                        //{
                        //    Debug.WriteLine("lol");
                        //}

                        // LinQ sum. Way better than VBA method.
                        double sumVCM = coreStructure.dicPO[DatePO][_Product]
                            .Where(x =>
                                (coreStructure.dicCustomer[x.Key.CustomerId].CustomerBigRegion == CustomerRegion) &&
                                (x.Value == true) &&
                                (PriorityTarget != "" ? coreStructure.dicCustomer[x.Key.CustomerId].CustomerType == PriorityTarget : true))
                            .Sum(x => x.Key.QuantityOrderKg); // Sum of Demand.

                        // To deal with Minimum Order Quantity.
                        double wallet = 0;

                        var _dicProductFC = coreStructure.dicFC.Where(x => x.Key.Date == DatePO.AddDays(-dayBefore)).FirstOrDefault();
                        var _dicProductFcLd = coreStructure.dicFC.Where(x => x.Key.Date == DatePO.AddDays(-dayLdBefore)).FirstOrDefault();

                        if (sumVCM != 0 && _dicProductFC.Value != null)
                        {
                            var dicSupplierFC = _dicProductFC.Value.Where(x => x.Key.ProductCode == _Product.ProductCode).FirstOrDefault();

                            double sumLdThuMua = 0;
                            if (_dicProductFcLd.Value != null)
                            {
                                // Check if Inventory has stock in other places.
                                // If no, equally distributed stuff.
                                // If yes, hah hah hah no.
                                var dicSupplierLdFC = _dicProductFcLd.Value.Where(x => x.Key.ProductCode == _Product.ProductCode).FirstOrDefault();
                                if (dicSupplierLdFC.Value != null)
                                {
                                    // Please NEVER FullOrder == true.
                                    var _SupplierThuMua = dicSupplierLdFC.Value
                                        .Where(x =>
                                            (coreStructure.dicSupplier[x.Key.SupplierId].SupplierRegion == "Lâm Đồng") &&
                                            (coreStructure.dicSupplier[x.Key.SupplierId].SupplierType == "ThuMua" || (SupplierType == "VCM" ? coreStructure.dicSupplier[x.Key.SupplierId].SupplierType == "VinEco" : false)) &&
                                            (x.Key.Availability.Contains(Convert.ToString((int)DatePO.AddDays(-dayLdBefore).DayOfWeek + 1))));

                                    if (_SupplierThuMua.Count() != 0)
                                    {
                                        // Doesn't matter what value SumVEThuMua holds, as long as it's a positive number to trigger a flag for calculating Rate
                                        sumLdThuMua = 7;  // Lolololol
                                    }
                                    else
                                    {
                                        sumLdThuMua = _SupplierThuMua.Sum(x => x.Key.QuantityForecast);
                                        //sumVEThuMua = _coreStructure.dicSupplierLdFC.Value
                                        //   .Where(x =>
                                        //       (coreStructure.dicSupplier[x.Key.SupplierId].SupplierRegion == SupplierRegion | coreStructure.dicSupplier[x.Key.SupplierId].SupplierRegion == "Lâm Đồng") &&
                                        //       (coreStructure.dicSupplier[x.Key.SupplierId].SupplierType == "ThuMua") &&
                                        //       (x.Key.Availability.Contains(Convert.ToString((int)DatePO.AddDays(-dayLdBefore).DayOfWeek + 1))))
                                        //   .Sum(x => x.Key.QuantityForecast);
                                    }
                                }
                            }

                            if (dicSupplierFC.Value != null)
                            {
                                var _resultSupplier = dicSupplierFC.Value
                                    .Where(x =>
                                        (coreStructure.dicSupplier[x.Key.SupplierId].SupplierRegion == SupplierRegion) &&
                                        (coreStructure.dicSupplier[x.Key.SupplierId].SupplierType == SupplierType) &&
                                        (SupplierType != "VinEco" ? x.Key.Availability.Contains(Convert.ToString((int)DatePO.AddDays(-dayBefore).DayOfWeek + 1)) : true));

                                double sumVE = 0;
                                if (_resultSupplier.Where(x => x.Key.FullOrder == true).FirstOrDefault().Key != null)
                                {
                                    // As long as it is a positive number
                                    // ... to counter FullOrder Supplier having negative Inventory.
                                    sumVE = 777;
                                }
                                else
                                {
                                    sumVE = _resultSupplier.Sum(x => x.Key.QuantityForecast);  // Sum of Supply
                                }

                                // Rate = Supply / Demand --> Deli = Demand * Rate.
                                double rate = sumVCM <= 0 ? 0 : sumVE / sumVCM;
                                if (rate > 0)
                                {
                                    // In case of an UpperLimit, obey it
                                    if (UpperLimit > 0)
                                    {
                                        rate = Math.Min(sumVCM != 0 ? sumVE / sumVCM : 0, UpperLimit);
                                        // Determining either equally distributing, or full order per PO until out of inventory.
                                        // Whoever wrote this is a FOOL :D
                                        // ... fuck that's me.
                                        rate = sumLdThuMua > 0 ? Math.Max(rate, 1) : rate;
                                    }

                                    // Customer Layer
                                    foreach (CustomerOrder _CustomerOrder in coreStructure.dicPO[DatePO][_Product].Keys.ToList())
                                    {
                                        if (coreStructure.dicCustomer[_CustomerOrder.CustomerId].CustomerBigRegion == CustomerRegion && (PriorityTarget != "" ? coreStructure.dicCustomer[_CustomerOrder.CustomerId].CustomerType == PriorityTarget : true))
                                        {
                                            double _quantityOrder = _CustomerOrder.QuantityOrder;
                                            SupplierForecast _SupplierForecast = null;

                                            var result = dicSupplierFC.Value
                                                .OrderBy(x => x.Key.level)
                                                .ThenByDescending(x => x.Key.FullOrder)
                                                .Where(x => (coreStructure.dicSupplier[x.Key.SupplierId].SupplierRegion == SupplierRegion) &&
                                                    (coreStructure.dicSupplier[x.Key.SupplierId].SupplierType == SupplierType) &&
                                                    (SupplierType == "ThuMua" ? x.Key.Availability.Contains(Convert.ToString((int)DatePO.AddDays(-dayBefore).DayOfWeek + 1)) : true) &&  // +1 coz by defaut, Sunday is 0. By "my" default, Sunday is 1.
                                                    (x.Key.FullOrder == true ? true : x.Key.QuantityForecast >= _CustomerOrder.QuantityOrder) &&
                                                    (CrossRegion == true ? x.Key.CrossRegion == true : true));

                                            // Coz for fuck sake, it can return null
                                            if (result.Count() != 0)
                                            {
                                                var _result = result.Aggregate((l, r) => dicDeli[DatePO.AddDays(-dayBefore)][_Product][l.Key] < dicDeli[DatePO.AddDays(-dayBefore)][_Product][r.Key] ? l : r).Key;
                                                if (_result != null && SupplierType == "ThuMua")
                                                {
                                                    _SupplierForecast = _result;
                                                }
                                                else
                                                {
                                                    _SupplierForecast = result.FirstOrDefault().Key;
                                                }
                                            }
                                            else
                                            {
                                                // Counter situation where there is no FullOrder Supplier for that Product
                                                _SupplierForecast = dicSupplierFC.Value
                                                    .OrderBy(x => x.Key.level)
                                                    .ThenByDescending(x => x.Key.FullOrder)
                                                    .Where(x => (coreStructure.dicSupplier[x.Key.SupplierId].SupplierRegion == SupplierRegion) &&
                                                        (coreStructure.dicSupplier[x.Key.SupplierId].SupplierType == SupplierType) &&
                                                        (SupplierType == "ThuMua" ? x.Key.Availability.Contains(Convert.ToString((int)DatePO.AddDays(-dayBefore).DayOfWeek + 1)) : true) &&  // +1 coz by defaut, Sunday is 0. By "my" default, Sunday is 1.
                                                        (CrossRegion == true ? x.Key.CrossRegion == true : true))
                                                    .FirstOrDefault().Key;
                                            }

                                            if (_SupplierForecast != null)
                                            {
                                                Dictionary<Product, Dictionary<CustomerOrder, Dictionary<SupplierForecast, DateTime>>> _dicCoordProduct = null;

                                                if (coreStructure.dicCoord.TryGetValue(DatePO, out _dicCoordProduct))
                                                {
                                                    Dictionary<CustomerOrder, Dictionary<SupplierForecast, DateTime>> _dicCoordCusSup = null;
                                                    if (_dicCoordProduct.TryGetValue(_Product, out _dicCoordCusSup))
                                                    {
                                                        Dictionary<SupplierForecast, DateTime> _SupplierForecastCoord = null;
                                                        if (_dicCoordCusSup.TryGetValue(_CustomerOrder, out _SupplierForecastCoord) && _SupplierForecastCoord == null)
                                                        {
                                                            wallet += _SupplierForecast.FullOrder == true ? _CustomerOrder.QuantityOrderKg : Math.Round(_CustomerOrder.QuantityOrderKg * rate, 2);

                                                            double _MOQ = 0;
                                                            // In case they are ordering and checking performance through an unit that's NOT FUCKING KILOGRAM!
                                                            if (YesNoByUnit)
                                                            {
                                                                // Cheapest way to calculate Kg per Unit.
                                                                // Man I'm so smart.
                                                                _MOQ = _CustomerOrder.QuantityOrderKg / _CustomerOrder.QuantityOrder;
                                                            }
                                                            // ... Otherwise, we're cool boys.
                                                            else
                                                            {
                                                                _MOQ = dicMinimum[_Product.ProductCode.Substring(0, 1)];

                                                            }
                                                            if (wallet <= _MOQ && _SupplierForecast.QuantityForecast >= _MOQ)
                                                            {
                                                                wallet = _MOQ;
                                                            }

                                                            if (wallet >= _MOQ)
                                                            {
                                                                if (sumVE <= 0) { continue; }
                                                                // Honestly, this should never be hit
                                                                // Jk I changed stuff. This should ALWAYS be hit
                                                                _SupplierForecastCoord = new Dictionary<SupplierForecast, DateTime>();
                                                                Guid _newGuid = Guid.NewGuid();
                                                                double _QuantityForecast = _CustomerOrder.Unit != "Kg" ? (wallet / _MOQ) * _MOQ : wallet;
                                                                _SupplierForecastCoord.Add(new SupplierForecast()
                                                                {
                                                                    _id = _newGuid,
                                                                    SupplierForecastId = _newGuid,
                                                                    SupplierId = _SupplierForecast.SupplierId,

                                                                    QuantityForecast = _QuantityForecast
                                                                }
                                                                , DatePO.AddDays(-dayBefore).Date);
                                                                //_SupplierForecastCoord.Keys.First()._id = Guid.NewGuid();
                                                                //_SupplierForecastCoord.Keys.First().SupplierForecastId = _SupplierForecastCoord.Keys.First()._id;
                                                                //_SupplierForecastCoord.Keys.First().SupplierId = _SupplierForecast.SupplierId;

                                                                //_SupplierForecastCoord.Keys.First().QuantityForecast = wallet;
                                                                _SupplierForecast.QuantityForecast -= _QuantityForecast;

                                                                // Pretty sure I don't need to recalculate sumVCM here anymore.
                                                                // Only sumVE matters here, to trigger a break.
                                                                // Then again even that is not really needed.
                                                                //sumVCM -= _CustomerOrder.QuantityOrder;
                                                                sumVE -= _SupplierForecast.FullOrder == true ? 0 : wallet;

                                                                //// Recalculating Rate - Unneccesary here I think.
                                                                //rate = sumVCM <= 0 ? 0 : Math.Min(sumVCM != 0 ? sumVE / sumVCM : 0, UpperLimit);
                                                                //rate = sumVCM <= 0 ? 0 : ((SupplierType == "VinEco") && (sumVEThuMua > 0) ? Math.Max(rate, 1) : rate);

                                                                coreStructure.dicCoord[DatePO][_Product][_CustomerOrder] = _SupplierForecastCoord;
                                                                dicDeli[DatePO.AddDays(-dayBefore)][_Product][_SupplierForecast] += wallet;

                                                                coreStructure.dicPO[DatePO][_Product][_CustomerOrder] = false;

                                                                wallet -= _QuantityForecast;
                                                            }
                                                        }
                                                        //else
                                                        //{
                                                        //// I have no fucking idea how this is hit now
                                                        //// In case of choosing more than one Supplier, this should be hit I guess
                                                        //var _CustomerOrder = new CustomerOrder();
                                                        //_CustomerOrder.CustomerId = CustomerOrder.CustomerId;
                                                        //_CustomerOrder.QuantityOrder = CustomerOrder.QuantityOrder;

                                                        //var __SupplierForecast = new SupplierForecast();
                                                        //__SupplierForecast.SupplierId = _SupplierForecast.SupplierId;
                                                        //__SupplierForecast.QuantityForecast = _SupplierForecast.FullOrder == true ? CustomerOrder.QuantityOrder : Math.Round(CustomerOrder.QuantityOrder * rate, 2);

                                                        //coreStructure.dicCoord[DatePO][Product].Add(_CustomerOrder, new Dictionary<SupplierForecast, DateTime>());
                                                        //coreStructure.dicCoord[DatePO][Product][_CustomerOrder].Add(__SupplierForecast, DatePO.AddDays(-dayBefore));
                                                        //}
                                                    }
                                                    //    else
                                                    //    {
                                                    //        var _CustomerOrder = new CustomerOrder();
                                                    //        _CustomerOrder.CustomerId = CustomerOrder.CustomerId;
                                                    //        _CustomerOrder.QuantityOrder = CustomerOrder.QuantityOrder;

                                                    //        var __SupplierForecast = new SupplierForecast();
                                                    //        __SupplierForecast.SupplierId = _SupplierForecast.SupplierId;
                                                    //        __SupplierForecast.QuantityForecast = _SupplierForecast.FullOrder == true ? CustomerOrder.QuantityOrder : Math.Round(CustomerOrder.QuantityOrder * rate, 2);

                                                    //        coreStructure.dicCoord[DatePO].Add(Product, new Dictionary<CustomerOrder, Dictionary<SupplierForecast, DateTime>>());
                                                    //        coreStructure.dicCoord[DatePO][Product].Add(_CustomerOrder, new Dictionary<SupplierForecast, DateTime>());
                                                    //        coreStructure.dicCoord[DatePO][Product][_CustomerOrder].Add(__SupplierForecast, DatePO.AddDays(-dayBefore));
                                                    //    }
                                                    //}
                                                    //else
                                                    //{
                                                    //    var _CustomerOrder = new CustomerOrder();
                                                    //    _CustomerOrder.CustomerId = CustomerOrder.CustomerId;
                                                    //    _CustomerOrder.QuantityOrder = CustomerOrder.QuantityOrder;

                                                    //    var __SupplierForecast = new SupplierForecast();
                                                    //    __SupplierForecast.SupplierId = _SupplierForecast.SupplierId;
                                                    //    __SupplierForecast.QuantityForecast = _SupplierForecast.FullOrder == true ? CustomerOrder.QuantityOrder : Math.Round(CustomerOrder.QuantityOrder * rate, 2);

                                                    //    coreStructure.dicCoord.Add(DatePO.Date, new Dictionary<Product, Dictionary<CustomerOrder, Dictionary<SupplierForecast, DateTime>>>());
                                                    //    coreStructure.dicCoord[DatePO].Add(Product, new Dictionary<CustomerOrder, Dictionary<SupplierForecast, DateTime>>());
                                                    //    coreStructure.dicCoord[DatePO][Product].Add(_CustomerOrder, new Dictionary<SupplierForecast, DateTime>());
                                                    //    coreStructure.dicCoord[DatePO][Product][_CustomerOrder].Add(__SupplierForecast, DatePO.AddDays(-dayBefore));
                                                }

                                                //CustomerOrder.QuantityOrder = 0;    //-= Math.Round(CustomerOrder.QuantityOrder * rate, 2);

                                                //coreStructure.dicPO[DatePO][Product][CustomerOrder] = false;
                                            }
                                        }
                                        //progressBarLabel.Text = (progressBar1.Value / progressBar1.Maximum).ToString("#0%");
                                        //progressBar1.PerformStep();
                                    }
                                }
                            }
                        }
                    }
                }
                stopwatch.Stop();
                Debug.WriteLine(Math.Round(stopwatch.Elapsed.TotalSeconds, 2) + "s From {0} to {1} for {2}{3}: Done!", SupplierRegion, CustomerRegion, SupplierType, (PriorityTarget != "" ? " for " + PriorityTarget : ""));
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
        /// Reading PO from VCM
        /// </summary>
        /// <param name="fileNameMB"></param>
        /// <param name="fileNameMN"></param>
        /// <param name="DateFrom"></param>
        /// <param name="DateTo"></param>
        private async Task UpdatePO(string fileNameMB, string fileNameMN)
        {
            //Process[] processBefore = Process.GetProcessesByName("excel");
            //string extension = Path.GetExtension(filePath);
            string header = "YES";
            //string conStr, sheetName;

            // These are openned here so they could be closed / released even in the case of Exceptions
            // Open First Workbook
            //Excel.Application xlApp = new Excel.Application();
            //Excel.Workbook xlWb = xlApp.Workbooks.Open(filePath, 0, true, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            //Excel.Worksheet xlWs = xlWb.Worksheets[1];
            //Excel.Range xlRng = xlWs.UsedRange;

            Excel.Application xlApp = new Excel.Application()
            {
                ScreenUpdating = false,
                EnableEvents = false,
                DisplayAlerts = false,
                DisplayStatusBar = false,
                AskToUpdateLinks = false
            };
            Excel.Workbook xlWb = null;
            Excel.Worksheet xlWs = null;
            Excel.Range xlRng = null;

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

                using (OleDbCommand oleCmd = new OleDbCommand())
                {
                    using (OleDbDataAdapter oleAdapt = new OleDbDataAdapter())
                    {
                        //string userChoice = Microsoft.VisualBasic.Interaction.InputBox(Prompt: "Save old PO?", Title: "Ayyyyyyy", DefaultResponse: "Default", XPos: -1, YPos: -1);
                        //PurchaseOrder PO = new PurchaseOrder();
                        //PO.PurchaseOrderCode = DateTime.Today.ToString();
                        //PO.ListPurchaseOrderDate = new List<PurchaseOrderDate>();

                        //var PO = new List<PurchaseOrderDate>();

                        var mongoClient = new MongoClient();
                        var db = mongoClient.GetDatabase("localtest");

                        var PO = mongoClient.GetDatabase("localtest").GetCollection<PurchaseOrderDate>("PurchaseOrder").AsQueryable().ToList();
                        var Product = db.GetCollection<Product>("Product").AsQueryable().ToList();
                        var Customer = db.GetCollection<Customer>("Customer").AsQueryable().ToList();

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

                        var dicPO = new Dictionary<DateTime, Dictionary<string, Dictionary<string, Guid>>>();
                        var dicProduct = new Dictionary<string, Product>();
                        var dicCustomer = new Dictionary<string, Customer>();

                        foreach (var _Product in Product)
                        {
                            if (!dicProduct.ContainsKey(_Product.ProductCode))
                                dicProduct.Add(_Product.ProductCode, _Product);
                        }

                        foreach (var _Customer in Customer)
                        {
                            if (!dicCustomer.ContainsKey(_Customer.CustomerCode + _Customer.CustomerType))
                                dicCustomer.Add(_Customer.CustomerCode + _Customer.CustomerType, _Customer);
                        }

                        var filePath = string.Format("C:\\Users\\Shirayuki\\Documents\\VinEco\\Project - Chia Hang\\Mastah Project\\{0}", fileNameMB);
                        var conStr = string.Format(Constants.Excel07ConString, filePath, header);

                        // North PO
                        xlWb = xlApp.Workbooks.Open(filePath, UpdateLinks: false, ReadOnly: true, Format: 5, Password: "", WriteResPassword: "", IgnoreReadOnlyRecommended: true, Origin: Excel.XlPlatform.xlWindows, Delimiter: "", Editable: false, Notify: false, Converter: 0, AddToMru: true, Local: false, CorruptLoad: false);

                        xlWs = xlWb.Worksheets[1];
                        xlRng = xlWs.UsedRange;

                        EatPO(PO, xlRng, xlWs, conStr, "Miền Bắc", dicPO, dicProduct, dicCustomer, Product, Customer, true);

                        // South PO
                        filePath = string.Format("C:\\Users\\Shirayuki\\Documents\\VinEco\\Project - Chia Hang\\Mastah Project\\{0}", fileNameMN);
                        conStr = string.Format(Constants.Excel07ConString, filePath, header);

                        xlWb = xlApp.Workbooks.Open(filePath, UpdateLinks: false, ReadOnly: true, Format: 5, Password: "", WriteResPassword: "", IgnoreReadOnlyRecommended: true, Origin: Excel.XlPlatform.xlWindows, Delimiter: "", Editable: false, Notify: false, Converter: 0, AddToMru: true, Local: false, CorruptLoad: false);

                        xlWs = xlWb.Worksheets[1];
                        xlRng = xlWs.UsedRange;

                        EatPO(PO, xlRng, xlWs, conStr, "Miền Nam", dicPO, dicProduct, dicCustomer, Product, Customer, false);

                        // Priority PO
                        filePath = string.Format("C:\\Users\\Shirayuki\\Documents\\VinEco\\Project - Chia Hang\\Mastah Project\\{0}", "Forecast MB Priority.xlsx");
                        conStr = string.Format(Constants.Excel07ConString, filePath, header);

                        xlWb = xlApp.Workbooks.Open(filePath, UpdateLinks: false, ReadOnly: true, Format: 5, Password: "", WriteResPassword: "", IgnoreReadOnlyRecommended: true, Origin: Excel.XlPlatform.xlWindows, Delimiter: "", Editable: false, Notify: false, Converter: 0, AddToMru: true, Local: false, CorruptLoad: false);

                        xlWs = xlWb.Worksheets[1];
                        xlRng = xlWs.UsedRange;

                        EatPO(PO, xlRng, xlWs, conStr, "Miền Bắc", dicPO, dicProduct, dicCustomer, Product, Customer, false);

                        #region Export Old PO before deleting
                        //var OldPO = db.GetCollection<PurchaseOrderDate>("PurchaseOrderDate").AsQueryable().ToList();

                        //DataTable dtOldPO = new DataTable();

                        //dtOldPO.Columns.Add("PCODE", typeof(string)).DefaultValue = "";
                        //dtOldPO.Columns.Add("CCODE", typeof(string)).DefaultValue = "";

                        //var dicRow = new Dictionary<string, int>();
                        //foreach (PurchaseOrderDate _PurchaseOrderDate in OldPO)
                        //{
                        //    string dateTarget = _PurchaseOrderDate.DateOrder.Date.ToString();
                        //    dtOldPO.Columns.Add(dateTarget, typeof(double)).DefaultValue = 0;
                        //    foreach (ProductOrder _ProductOrder in _PurchaseOrderDate.ListProductOrder)
                        //    {
                        //        foreach (CustomerOrder _CustomerOrder in _ProductOrder.ListCustomerOrder)
                        //        {
                        //            string pcode = dicProduct.Values.Where(x => x.ProductId == _ProductOrder.ProductId).FirstOrDefault().ProductCode;
                        //            string ccode = dicCustomer.Values.Where(x => x.CustomerId == _CustomerOrder.CustomerId).FirstOrDefault().CustomerCode;

                        //            DataRow dr = null;

                        //            int _rowIndex = 0;
                        //            if (!dicRow.TryGetValue(pcode + ccode, out _rowIndex))
                        //            {
                        //                dicRow.Add(pcode + ccode, _rowIndex);
                        //                _rowIndex++;

                        //                dr = dtOldPO.NewRow();
                        //            }
                        //            else
                        //            {
                        //                dr = dtOldPO.Rows[_rowIndex];
                        //            }

                        //            dr["PCODE"] = pcode;
                        //            dr["CCODE"] = ccode;

                        //            dr[dateTarget] = Convert.ToDouble(dr[dateTarget]) + _CustomerOrder.QuantityOrder;

                        //            dtOldPO.Rows.Add(dr);
                        //        }

                        //    }
                        //}

                        //string _filePathLocal = Application.StartupPath.Replace("\\bin\\Debug", "") + "\\OldPO\\{0}";
                        //string fileName = string.Format("PO {0}.xlsx", DateTime.Now.ToString("yyyyMMdd HH.mm.ss") + ")");
                        //string path = string.Format(_filePathLocal, fileName);
                        //Debug.WriteLine(path);
                        //using (Aspose.Cells.Workbook xlWbNew = new Aspose.Cells.Workbook())
                        //{
                        //    OutputExcelAspose(dtOldPO, "Sheet1", xlWbNew, true, 1);
                        //    dtOldPO.Dispose();

                        //    xlWbNew.Save(path);
                        //}
                        //Delete_Evaluation_Sheet_Interop(path);
                        #endregion

                        db.DropCollection("PurchaseOrder");
                        await db.GetCollection<PurchaseOrderDate>("PurchaseOrder").InsertManyAsync(PO);

                        db.DropCollection("Product");
                        await db.GetCollection<Product>("Product").InsertManyAsync(Product);

                        db.DropCollection("Customer");
                        await db.GetCollection<Customer>("Customer").InsertManyAsync(Customer);

                        db = null;

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
                    }
                }
            }
            catch (Exception ex)
            {
                //Debug.WriteLine(ex);
                throw ex;
                //MessageBox.Show(ex.Message, "Exception Error");
            }
            finally
            {
                xlApp.ScreenUpdating = true;
                xlApp.Calculation = Excel.XlCalculation.xlCalculationAutomatic;
                xlApp.EnableEvents = true;
                xlApp.DisplayAlerts = false;
                xlApp.DisplayStatusBar = true;
                xlApp.AskToUpdateLinks = true;

                #region Clean up
                // Cleanup
                //GC.Collect();
                //GC.WaitForPendingFinalizers();

                // Rule of thumb for releasing com objects:
                //   never use two dots, all COM objects must be referenced and released individually
                //   ex: [somthing].[something].[something] is bad

                // Release com objects to fully kill excel process from running in the background
                if (xlRng != null) { Marshal.ReleaseComObject(xlRng); }
                if (xlWs != null) { Marshal.ReleaseComObject(xlWs); }

                // Close and release
                xlWb.Close();
                if (xlWb != null) { Marshal.ReleaseComObject(xlWs); }

                // Quit and release
                xlApp.Quit();
                if (xlApp != null) { Marshal.ReleaseComObject(xlApp); }

                xlRng = null;
                xlWs = null;
                xlWb = null;
                xlApp = null;

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
        /// Reading Forecast
        /// </summary>
        /// <param name="fileVE"></param>
        /// <param name="fileTM"></param>
        /// <returns></returns>
        private async Task UpdateFC(string fileVE, string fileTM)
        {
            //Process[] processBefore = Process.GetProcessesByName("excel");
            //string extension = Path.GetExtension(filePath);
            string header = "YES";
            //string conStr, sheetName;

            // These are openned here so they could be closed / released even in the case of Exceptions
            // Open First Workbook
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWb = null;
            Excel.Worksheet xlWs = null;
            Excel.Range xlRng = null;

            xlApp.ScreenUpdating = false;
            xlApp.EnableEvents = false;
            xlApp.DisplayAlerts = false;
            xlApp.DisplayStatusBar = false;
            xlApp.AskToUpdateLinks = false;

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

                using (OleDbCommand oleCmd = new OleDbCommand())
                {
                    using (OleDbDataAdapter oleAdapt = new OleDbDataAdapter())
                    {
                        //PurchaseOrder PO = new PurchaseOrder();
                        //PO.PurchaseOrderCode = DateTime.Today.ToString();
                        //PO.ListPurchaseOrderDate = new List<PurchaseOrderDate>();

                        var FC = new List<ForecastDate>();

                        var mongoClient = new MongoClient();
                        var db = mongoClient.GetDatabase("localtest");

                        var Product = mongoClient.GetDatabase("localtest").GetCollection<Product>("Product").AsQueryable().ToList();
                        var Supplier = mongoClient.GetDatabase("localtest").GetCollection<Supplier>("Supplier").AsQueryable().ToList();

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

                        string filePath = "";
                        string conStr = "";

                        filePath = string.Format("C:\\Users\\Shirayuki\\Documents\\VinEco\\Project - Chia Hang\\Mastah Project\\{0}", fileVE);
                        conStr = string.Format(Constants.Excel07ConString, filePath, header);

                        var dicFC = new Dictionary<DateTime, Dictionary<string, Dictionary<string, Guid>>>();
                        var dicProduct = new Dictionary<string, Product>();
                        var dicSupplier = new Dictionary<string, Supplier>();

                        foreach (var _Product in Product)
                        {
                            Product _product = null;
                            if (!dicProduct.TryGetValue(_Product.ProductCode, out _product))
                            {
                                dicProduct.Add(_Product.ProductCode, _Product);
                            }
                        }

                        foreach (var _Supplier in Supplier)
                        {
                            Supplier _supplier = null;
                            if (!dicSupplier.TryGetValue(_Supplier.SupplierCode, out _supplier))
                            {
                                dicSupplier.Add(_Supplier.SupplierCode, _Supplier);
                            }
                        }

                        xlWb = xlApp.Workbooks.Open(filePath,
                            UpdateLinks: false,
                            ReadOnly: true,
                            Format: 5,
                            Password: "",
                            WriteResPassword: "",
                            IgnoreReadOnlyRecommended: true,
                            Origin: Excel.XlPlatform.xlWindows,
                            Delimiter: "",
                            Editable: false,
                            Notify: false,
                            Converter: 0,
                            AddToMru: true,
                            Local: false,
                            CorruptLoad: false);

                        xlApp.Calculation = Excel.XlCalculation.xlCalculationManual;

                        xlWs = xlWb.Worksheets[1];
                        xlRng = xlWs.UsedRange;

                        EatForecast(
                            FC: FC,
                            xlRng: xlRng, xlWs: xlWs,
                            conStr: conStr,
                            SupplierType: "VinEco",
                            dicFC: dicFC,
                            dicProduct: dicProduct,
                            dicSupplier: dicSupplier,
                            Product: Product,
                            Supplier: Supplier);

                        xlWb.Close(SaveChanges: false);

                        filePath = string.Format("C:\\Users\\Shirayuki\\Documents\\VinEco\\Project - Chia Hang\\Mastah Project\\{0}", fileTM);
                        conStr = string.Format(Constants.Excel07ConString, filePath, header);

                        xlWb = xlApp.Workbooks.Open(filePath,
                            UpdateLinks: false,
                            ReadOnly: true,
                            Format: 5,
                            Password: "",
                            WriteResPassword: "",
                            IgnoreReadOnlyRecommended: true,
                            Origin: Excel.XlPlatform.xlWindows,
                            Delimiter: "",
                            Editable: false,
                            Notify: false,
                            Converter: 0,
                            AddToMru: true,
                            Local: false,
                            CorruptLoad: false);

                        xlApp.Calculation = Excel.XlCalculation.xlCalculationManual;

                        xlWs = xlWb.Worksheets[1];
                        xlRng = xlWs.UsedRange;

                        EatForecast(
                            FC: FC,
                            xlRng: xlRng, xlWs: xlWs,
                            conStr: conStr,
                            SupplierType: "ThuMua",
                            dicFC: dicFC,
                            dicProduct: dicProduct,
                            dicSupplier: dicSupplier,
                            Product: Product,
                            Supplier: Supplier);

                        xlApp.Calculation = Excel.XlCalculation.xlCalculationAutomatic;

                        //string connectionString = System.Configuration.ConfigurationManager.ConnectionStrings["mongodb_vecrops.salesms"].ConnectionString;
                        //MongoClient mongoClient = new MongoClient(connectionString);
                        //var db = mongoClient.GetDatabase("salesms_uat");
                        //var mongoClient = new MongoClient();
                        //var db = mongoClient.GetDatabase("localtest");
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
                    }
                }
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
                if (xlRng != null) { Marshal.ReleaseComObject(xlRng); }
                if (xlWs != null) { Marshal.ReleaseComObject(xlWs); }

                // Close and release
                xlWb.Close(SaveChanges: false);
                if (xlWb != null) { Marshal.ReleaseComObject(xlWs); }

                xlApp.ScreenUpdating = true;
                xlApp.EnableEvents = true;
                xlApp.DisplayAlerts = false;
                xlApp.DisplayStatusBar = true;
                xlApp.AskToUpdateLinks = true;

                // Quit and release
                xlApp.Quit();
                if (xlApp != null) { Marshal.ReleaseComObject(xlApp); }

                xlRng = null;
                xlWs = null;
                xlWb = null;
                xlApp = null;

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
        /// Updating OpenConfig file & Do afterward updating.
        /// </summary>
        /// <param name="xlWb"></param>
        /// <param name="conStr"></param>
        private async void UpdateOpenConfig()
        {
            try
            {
                var mongoClient = new MongoClient();
                var db = mongoClient.GetDatabase("localtest");

                var Product = mongoClient.GetDatabase("localtest").GetCollection<Product>("Product").AsQueryable().ToList();
                var ProductUnitList = new List<ProductUnit>();

                // Remember the list of running Excel.Application.
                // Before initialize xlApp.
                Process[] processBefore = Process.GetProcessesByName("excel");

                // Initialize new instance of Interop Excel.Application.
                Excel.Application xlApp = new Excel.Application()
                {
                    ScreenUpdating = false,
                    EnableEvents = false,
                    DisplayAlerts = false,
                    DisplayStatusBar = false,
                    AskToUpdateLinks = false
                };

                // Remember the list of running Excel.Application.
                // After initialize xlApp.
                Process[] processAfter = Process.GetProcessesByName("excel");

                int processID = 0;

                // Compare two lists, get the first and the only process that's not in the 'Before' List.
                foreach (Process process in processAfter)
                {
                    if (!processBefore.Select(p => p.Id).Contains(process.Id))
                    {
                        processID = process.Id;
                        break;
                    }
                }

                var filePath = string.Format("C:\\Users\\Shirayuki\\Documents\\VinEco\\Project - Chia Hang\\Mastah Project\\{0}", "ChiaHang OpenConfig.xlsb");
                var conStr = string.Format(Constants.Excel07ConString, filePath, "YES");

                var xlWb = xlApp.Workbooks.Open(filePath,
                    UpdateLinks: false,
                    ReadOnly: true,
                    Format: 5,
                    Password: "",
                    WriteResPassword: "",
                    IgnoreReadOnlyRecommended: true,
                    Origin: Excel.XlPlatform.xlWindows,
                    Delimiter: "",
                    Editable: false,
                    Notify: false,
                    Converter: 0,
                    AddToMru: true,
                    Local: false,
                    CorruptLoad: false);

                xlApp.Calculation = Excel.XlCalculation.xlCalculationManual;

                Excel.Worksheet xlWsUnitConversion = xlWb.Worksheets["UnitConversion"];
                Excel.Range xlRng = xlWsUnitConversion.UsedRange;

                DataTable dtCombo = new DataTable();

                OleDbConnection oleCon = new OleDbConnection(conStr);

                string connectionString = "Select * From [" + xlWsUnitConversion.Name.ToString() + "$" + xlRng.Offset[0, 0].Address[RowAbsolute: false, ColumnAbsolute: false, ReferenceStyle: Excel.XlReferenceStyle.xlA1, External: xlRng] + "]";
                OleDbDataAdapter _oleAdapt = new OleDbDataAdapter(connectionString, oleCon);
                _oleAdapt.Fill(dtCombo);

                oleCon.Close();

                foreach (DataRow dr in dtCombo.Rows)
                {
                    var _Product = Product.Where(x => x.ProductCode == dr["VECode"].ToString()).FirstOrDefault();
                    if (_Product == null)
                    {
                        // To be fucking honest, this should NEVER be hit.
                        // Unit Converstion definition for a product that's NOT EVEN EXIST.
                        // ... and of fucking course IT IS HIT.
                    }
                    else
                    {
                        ProductUnit _ProductUnit = ProductUnitList
                        .Where(x =>
                            x.ProductCode == dr["VECode"].ToString())
                        .FirstOrDefault();

                        string _Region = dr["Region"].ToString();
                        switch (_Region)
                        {
                            case "MB": _Region = "Miền Bắc"; break;
                            case "MN": _Region = "Miền Nam"; break;
                            case "All": _Region = "All"; break;
                            default: break;
                        }

                        if (_ProductUnit != null)
                        {

                            if (_ProductUnit.ListRegion == null)
                            {
                                ProductUnitRegion _ProductUnitRegion = new ProductUnitRegion()
                                {
                                    _id = Guid.NewGuid(),
                                    Region = _Region,
                                    OrderUnitType = ProperUnit(dr["OrderUnitType"].ToString()),
                                    OrderUnitPer = ProperUnit(dr["OrderUnitType"].ToString()) == "Kg" ? 1 : (double)dr["OderUnitPer"],
                                    SaleUnitType = ProperUnit(dr["SaleUnitType"].ToString()),
                                    SaleUnitPer = ProperUnit(dr["SaleUnitType"].ToString()) == "Kg" ? 1 : (double)dr["SaleUnitPer"]
                                };

                                var _ListRegion = new List<ProductUnitRegion>();
                                _ListRegion.Add(_ProductUnitRegion);

                                _ProductUnit.ListRegion = _ListRegion;

                            }
                            else
                            {
                                ProductUnitRegion _ProductUnitRegion = _ProductUnit.ListRegion.Where(x => x.Region == _Region).FirstOrDefault();

                                if (_ProductUnitRegion == null)
                                {
                                    _ProductUnitRegion = new ProductUnitRegion()
                                    {
                                        _id = Guid.NewGuid(),
                                        Region = _Region,
                                        OrderUnitType = ProperUnit(dr["OrderUnitType"].ToString()),
                                        OrderUnitPer = ProperUnit(dr["OrderUnitType"].ToString()) == "Kg" ? 1 : (double)dr["OrderUnitPer"],
                                        SaleUnitType = ProperUnit(dr["SaleUnitType"].ToString()),
                                        SaleUnitPer = ProperUnit(dr["SaleUnitType"].ToString()) == "Kg" ? 1 : (double)dr["SaleUnitPer"]
                                    };
                                    _ProductUnit.ListRegion.Add(_ProductUnitRegion);
                                }
                                else
                                {
                                    _ProductUnitRegion = new ProductUnitRegion()
                                    {
                                        _id = Guid.NewGuid(),
                                        Region = _Region,
                                        OrderUnitType = ProperUnit(dr["OrderUnitType"].ToString()),
                                        OrderUnitPer = ProperUnit(dr["OrderUnitType"].ToString()) == "Kg" ? 1 : (double)dr["OrderUnitPer"],
                                        SaleUnitType = ProperUnit(dr["SaleUnitType"].ToString()),
                                        SaleUnitPer = ProperUnit(dr["SaleUnitType"].ToString()) == "Kg" ? 1 : (double)dr["SaleUnitPer"]
                                    };
                                }
                            }
                        }
                        else
                        {

                            _ProductUnit = new ProductUnit()
                            {
                                ProductCode = dr["VECode"].ToString(),
                                ProductId = Product.Where(x => x.ProductCode == dr["VECode"].ToString()).FirstOrDefault().ProductId,
                                ListRegion = new List<ProductUnitRegion>()
                            };

                            _ProductUnit.ListRegion.Add(new ProductUnitRegion()
                            {
                                _id = Guid.NewGuid(),
                                Region = _Region,
                                OrderUnitType = ProperUnit(dr["OrderUnitType"].ToString()),
                                OrderUnitPer = ProperUnit(dr["OrderUnitType"].ToString()) == "Kg" ? 1 : (double)dr["OrderUnitPer"],
                                SaleUnitType = ProperUnit(dr["SaleUnitType"].ToString()),
                                SaleUnitPer = ProperUnit(dr["SaleUnitType"].ToString()) == "Kg" ? 1 : (double)dr["SaleUnitPer"]
                            });

                            ProductUnitList.Add(_ProductUnit);
                        }
                    }
                }

                db.DropCollection("ProductUnit");
                await db.GetCollection<ProductUnit>("ProductUnit").InsertManyAsync(ProductUnitList);

                Marshal.ReleaseComObject(xlWsUnitConversion); xlWsUnitConversion = null;

                xlWb.Close(SaveChanges: false);
                Marshal.ReleaseComObject(xlWb); xlWb = null;

                xlApp.Quit();
                Marshal.ReleaseComObject(xlApp); xlApp = null;

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

        /// <summary>
        /// Do naughty stuff with FC
        /// </summary>
        /// <param name="FC"></param>
        /// <param name="xlRng"></param>
        /// <param name="xlWs"></param>
        /// <param name="conStr"></param>
        /// <param name="SupplierType"></param>
        /// <param name="dicFC"></param>
        /// <param name="dicProduct"></param>
        /// <param name="dicSupplier"></param>
        /// <param name="Product"></param>
        /// <param name="Supplier"></param>
        private void EatForecast(List<ForecastDate> FC, Excel.Range xlRng, Excel.Worksheet xlWs, string conStr, string SupplierType, Dictionary<DateTime, Dictionary<string, Dictionary<string, Guid>>> dicFC, Dictionary<string, Product> dicProduct, Dictionary<string, Supplier> dicSupplier, List<Product> Product, List<Supplier> Supplier)
        {
            try
            {
                // To combat those which will be Cross-Region ( Mostly, if not only, Fruits )
                //var dicCrossRegionVinEco = new Dictionary<string, bool>();
                //dicCrossRegionVinEco.Add("K04401", true);   // Đu đủ ruột vàng
                //dicCrossRegionVinEco.Add("K05301", true);   // Dưa hấu đỏ có hạt
                //dicCrossRegionVinEco.Add("K05501", true);   // Dưa hấu vàng có hạt
                //dicCrossRegionVinEco.Add("K05701", true);   // Dưa lê
                //dicCrossRegionVinEco.Add("K15101", true);   // Dưa lê hoàng cẩm
                //dicCrossRegionVinEco.Add("K05901", true);   // Dưa lê hoàng kim
                //dicCrossRegionVinEco.Add("K06001", true);   // Dưa lê kim cô nương
                //dicCrossRegionVinEco.Add("K15401", true);   // Dưa lê kim đế vương
                //dicCrossRegionVinEco.Add("K06101", true);   // Dưa lê kim hoàng hậu
                //dicCrossRegionVinEco.Add("K16201", true);   // Dưa lê Kim Thúy Mật
                //dicCrossRegionVinEco.Add("K06301", true);   // Dưa lưới dài vỏ vàng
                //dicCrossRegionVinEco.Add("K06401", true);   // Dưa lưới dài vỏ xanh
                //dicCrossRegionVinEco.Add("K06501", true);   // Dưa lưới tròn vỏ vàng
                //dicCrossRegionVinEco.Add("K06601", true);   // Dưa lưới tròn vỏ xanh

                int rowIndex = 0;
                if (xlRng.Cells[1, 1].value != "Region" & xlRng.Cells[1, 1].value != "Vùng")
                {
                    do
                    {
                        rowIndex++;
                        if (rowIndex >= xlRng.Rows.Count) { return; }
                    } while (xlRng.Cells[rowIndex + 1, 1].Value != "Region" & xlRng.Cells[rowIndex + 1, 1].Value != "Vùng");
                }

                DataTable dt = new DataTable();

                OleDbConnection oleCon = new OleDbConnection(conStr);

                OleDbDataAdapter _oleAdapt = new OleDbDataAdapter("Select * From [" + xlWs.Name.ToString() + "$" + xlRng.Offset[rowIndex, 0].Address[RowAbsolute: false, ColumnAbsolute: false, ReferenceStyle: Excel.XlReferenceStyle.xlA1, External: xlRng] + "]", oleCon);
                string _str = xlRng.Offset[rowIndex, 0].Address as string;
                Debug.WriteLine(_str);
                _oleAdapt.Fill(dt);

                oleCon.Close();

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
                        bool isNewFC = false;

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
                        if (_listProductForecast == null) { _listProductForecast = new List<ProductForecast>(); }

                        // Loop for every value
                        foreach (DataRow dr in dt.Rows)
                        {
                            // In case of empty SCODE. I really hate to deal with this case. Like, really.
                            if (dr["SCODE"] == null || String.IsNullOrEmpty(dr["SCODE"].ToString()))
                            {
                                dr["SCODE"] = dr["SNAME"];  // Oh for god's sake.
                            }

                            // If OrderQuantity is not 0 - Not Anymore?
                            //object _OrderQuantity = dr[dc.ColumnName];
                            if (dr["PCODE"] != DBNull.Value /*&& dr[dc.ColumnName] != DBNull.Value*/ /*&& Convert.ToDouble(dr[dc.ColumnName]) > 0*/ && (SupplierType == "ThuMua" ? dr["SCODE"] != DBNull.Value : true))
                            {
                                // Olala
                                List<SupplierForecast> _ListSupplierForecast = null;
                                SupplierForecast _SupplierForcast = null;
                                ProductForecast _ProductForecast = null;
                                // Olala2
                                bool isNewProductOrder = false;
                                bool isNewCustomerOrder = false;
                                // Olala3
                                Dictionary<string, Guid> dicStore = null;
                                // #RandomGreenStuff
                                _dicProduct = dicFC[dateValue.Date];
                                if (_dicProduct.TryGetValue(dr["PCODE"].ToString(), out dicStore))
                                {
                                    Product _product = null;
                                    if (!dicProduct.TryGetValue(dr["PCODE"].ToString(), out _product))
                                    {
                                        _product = dicProduct.Values.Where(x => x.ProductCode == dr["PCODE"].ToString()).FirstOrDefault();
                                        if (_product == null)
                                        {
                                            _product = new Product();

                                            _product._id = Guid.NewGuid();
                                            _product.ProductId = _product._id;
                                            _product.ProductCode = dr["PCODE"].ToString();
                                            _product.ProductName = dr["PNAME"].ToString();
                                            _product.ProductClassification = dr["PCLASS"].ToString();
                                            _product.ProductVECode = dt.Columns.Contains("VECrops Code") ? dr["VECrops Code"].ToString() : "";

                                            Product.Add(_product);

                                            dicProduct.Add(dr["PCODE"].ToString(), _product);
                                        }
                                    }
                                    _ProductForecast = _FC.ListProductForecast.Where(x => x.ProductId == _product.ProductId).FirstOrDefault();

                                    Guid _id;
                                    if (dicStore.TryGetValue(dr["SCODE"].ToString(), out _id))
                                    {
                                        _SupplierForcast = _ProductForecast.ListSupplierForecast.Where(x => x.SupplierId == _id).FirstOrDefault();
                                    }
                                    else
                                    {
                                        isNewCustomerOrder = true;

                                        Supplier _supplier = null;
                                        if (!dicSupplier.TryGetValue(dr["SCODE"].ToString(), out _supplier))
                                        {
                                            _supplier = dicSupplier.Values.Where(x => x.SupplierCode == dr["SCODE"].ToString()).FirstOrDefault();
                                            if (_supplier == null)
                                            {
                                                _supplier = new Supplier();

                                                _supplier._id = Guid.NewGuid();
                                                _supplier.SupplierId = _supplier._id;
                                                _supplier.SupplierCode = dr["SCODE"].ToString(); //SupplierType == "ThuMua" ? (dr["SCODE"] == null ? dr["SCODE"].ToString() : dr["SNAME"].ToString()) : dr["SNAME"].ToString();
                                                _supplier.SupplierName = dr["SNAME"].ToString();
                                                _supplier.SupplierType = SupplierType == "ThuMua" && dr["Tag"].ToString() == "VCM" ? "VCM" : SupplierType;

                                                string _region = dr["Region"].ToString();
                                                switch (_region)
                                                {
                                                    case "LD": _region = "Lâm Đồng"; break;
                                                    case "MB": _region = "Miền Bắc"; break;
                                                    case "MN": _region = "Miền Nam"; break;
                                                    default: break;
                                                }
                                                _supplier.SupplierRegion = _region;
                                                _supplier.SupplierType = SupplierType == "ThuMua" && dr["Tag"].ToString() == "VCM" ? "VCM" : SupplierType;

                                                Supplier.Add(_supplier);
                                                dicSupplier.Add(_supplier.SupplierCode, _supplier);
                                            }
                                        }

                                        _SupplierForcast = new SupplierForecast();
                                        _SupplierForcast._id = Guid.NewGuid();
                                        _SupplierForcast.SupplierForecastId = _SupplierForcast._id;
                                        _SupplierForcast.SupplierId = _supplier.SupplierId;

                                        dicFC[dateValue.Date][dr["PCODE"].ToString()].Add(dr["SCODE"].ToString(), _supplier.SupplierId);
                                    }
                                }
                                else
                                {
                                    isNewProductOrder = true;
                                    isNewCustomerOrder = true;

                                    Product _product = null;
                                    if (!dicProduct.TryGetValue(dr["PCODE"].ToString(), out _product))
                                    {
                                        _product = dicProduct.Values.Where(x => x.ProductCode == dr["PCODE"].ToString()).FirstOrDefault();
                                        if (_product == null)
                                        {
                                            _product = new Product();

                                            _product._id = Guid.NewGuid();
                                            _product.ProductId = _product._id;
                                            _product.ProductCode = dr["PCODE"].ToString();
                                            _product.ProductName = dr["PNAME"].ToString();
                                            _product.ProductClassification = dr["PCLASS"].ToString();
                                            _product.ProductVECode = dt.Columns.Contains("VECrops Code") ? dr["VECrops Code"].ToString() : "";

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
                                        _supplier = dicSupplier.Values.Where(x => x.SupplierCode == dr["SCODE"].ToString()).FirstOrDefault();
                                        if (_supplier == null)
                                        {
                                            _supplier = new Supplier();

                                            _supplier._id = Guid.NewGuid();
                                            _supplier.SupplierId = _supplier._id;
                                            _supplier.SupplierCode = dr["SCODE"].ToString(); //SupplierType == "ThuMua" ? (dr["SCODE"] == null ? dr["SCODE"].ToString() : dr["SNAME"].ToString()) : dr["SNAME"].ToString();
                                            _supplier.SupplierName = dr["SNAME"].ToString();
                                            _supplier.SupplierType = SupplierType == "ThuMua" && dr["Tag"].ToString() == "VCM" ? "VCM" : SupplierType;

                                            string _region = dr["Region"].ToString();
                                            switch (_region)
                                            {
                                                case "LD": _region = "Lâm Đồng"; break;
                                                case "MB": _region = "Miền Bắc"; break;
                                                case "MN": _region = "Miền Nam"; break;
                                                default: break;
                                            }
                                            _supplier.SupplierRegion = _region;
                                            _supplier.SupplierType = SupplierType == "ThuMua" && dr["Tag"].ToString() == "VCM" ? "VCM" : SupplierType;

                                            Supplier.Add(_supplier);
                                            dicSupplier.Add(_supplier.SupplierCode, _supplier);
                                        }
                                    }

                                    _SupplierForcast = new SupplierForecast();
                                    _SupplierForcast._id = Guid.NewGuid();
                                    _SupplierForcast.SupplierForecastId = _SupplierForcast._id;
                                    _SupplierForcast.SupplierId = _supplier.SupplierId;

                                    dicFC[dateValue.Date].Add(dr["PCODE"].ToString(), new Dictionary<string, Guid>());
                                    dicFC[dateValue.Date][dr["PCODE"].ToString()].Add(dr["SCODE"].ToString(), _supplier.SupplierId);
                                }

                                // Filling in data
                                _ListSupplierForecast = _ProductForecast.ListSupplierForecast;

                                //_SupplierForcast.Unit = dr["Unit"].ToString();
                                _SupplierForcast.QuantityForecast = Convert.ToDouble((dr[dc.ColumnName] == DBNull.Value ? 0 : dr[dc.ColumnName]).ToString());

                                // Special part for ThuMua
                                TextInfo myTI = new CultureInfo("en-US", false).TextInfo;
                                if (SupplierType == "ThuMua")
                                {
                                    _SupplierForcast.QualityControlPass = String.IsNullOrEmpty(dr["QC"].ToString()) ? false : (myTI.ToTitleCase(dr["QC"].ToString()) == "Ok" ? true : false);
                                    _SupplierForcast.LabelVinEco = String.IsNullOrEmpty(dr["Label VE"].ToString()) ? false : (myTI.ToTitleCase(dr["Label VE"].ToString()) == "Yes" ? true : false);
                                    _SupplierForcast.FullOrder = String.IsNullOrEmpty(dr["100%"].ToString()) ? false : (myTI.ToTitleCase(dr["100%"].ToString()) == "Yes" ? true : false);
                                    _SupplierForcast.CrossRegion = String.IsNullOrEmpty(dr["CrossRegion"].ToString()) ? false : (myTI.ToTitleCase(dr["CrossRegion"].ToString()) == "Yes" ? true : false);
                                    _SupplierForcast.level = String.IsNullOrEmpty(dr["Level"].ToString()) ? Convert.ToByte(0) : Convert.ToByte(dr["Level"]);
                                    _SupplierForcast.Availability = String.IsNullOrEmpty(dr["Availability"].ToString()) ? "" : dr["Availability"].ToString();
                                }
                                else
                                {
                                    _SupplierForcast.QualityControlPass = true;
                                    _SupplierForcast.LabelVinEco = true;
                                    _SupplierForcast.FullOrder = false;
                                    _SupplierForcast.CrossRegion = false;
                                    _SupplierForcast.level = 1;
                                    _SupplierForcast.Availability = "1234567";
                                }

                                if (SupplierType == "VinEco" && dr["PCODE"].ToString().Substring(0, 1) == "K" && (dr["Region"].ToString() == "MN" || dr["Region"].ToString() == "Miền Nam")) //dicCrossRegionVinEco.ContainsKey(dr["PCODE"].ToString()))
                                {
                                    //_SupplierForcast.FullOrder = false;
                                    _SupplierForcast.CrossRegion = true;
                                }

                                if (isNewCustomerOrder) { _ListSupplierForecast.Add(_SupplierForcast); }

                                _ProductForecast.ListSupplierForecast = _ListSupplierForecast;
                                if (isNewProductOrder) { _FC.ListProductForecast.Add(_ProductForecast); }
                            }
                        }

                        _FC.ListProductForecast = _listProductForecast;

                        if (isNewFC) { FC.Add(_FC); }
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        /// Do naughty stuff with PO
        /// </summary>
        /// <param name="Purchase Order"></param>
        /// <param name="xlRng"></param>
        /// <param name="xlWs"></param>
        /// <param name="conStr"></param>
        /// <param name="PORegion"></param>
        /// <param name="dicPO"></param>
        /// <param name="dicProduct"></param>
        /// <param name="dicCustomer"></param>
        /// <param name="Product"></param>
        /// <param name="Customer"></param>
        private void EatPO(List<PurchaseOrderDate> PO, Excel.Range xlRng, Excel.Worksheet xlWs, string conStr, string PORegion, Dictionary<DateTime, Dictionary<string, Dictionary<string, Guid>>> dicPO, Dictionary<string, Product> dicProduct, Dictionary<string, Customer> dicCustomer, List<Product> Product, List<Customer> Customer, bool YesNoNew = false)
        {
            try
            {
                DataTable dt = new DataTable();
                // Find first row
                int rowIndex = 0;
                do { rowIndex++; } while (xlRng.Cells[rowIndex + 1, 1].Value != "VE Code");

                OleDbConnection oleCon = new OleDbConnection(conStr);

                OleDbDataAdapter _oleAdapt = new OleDbDataAdapter("Select * From [" + xlWs.Name.ToString() + "$" + xlRng.Offset[rowIndex, 0].Address[false, false, Excel.XlReferenceStyle.xlA1, xlRng] + "]", oleCon);
                string _str = xlRng.Offset[rowIndex, 0].Address as string;
                Debug.WriteLine(_str);
                _oleAdapt.Fill(dt);

                oleCon.Close();

                var mongoClient = new MongoClient();
                var db = mongoClient.GetDatabase("localtest").GetCollection<PurchaseOrderDate>("PurchaseOrder");

                // Main Loop
                foreach (DataColumn dc in dt.Columns)
                {
                    DateTime dateValue;

                    // Loop for every column. Only stop at column with Date at the top ( Indicating it being PurchaseOrder for that Date )
                    if (DateTime.TryParse(dc.ColumnName, out dateValue)/* && (dateValue.Date >= DateTime.Today.AddDays(0).Date)*/)
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
                        bool isNewPODate = false;

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
                        if (_listProductOrder == null) { _listProductOrder = new List<ProductOrder>(); }

                        // Loop for every value
                        foreach (DataRow dr in dt.Rows)
                        {
                            // If OrderQuantity is not 0 - Not Anymore?
                            //object _OrderQuantity = dr[dc.ColumnName];
                            double _value = 0;
                            if (dr["VE Code"] != DBNull.Value && dr[dt.Columns.IndexOf(dc)] != DBNull.Value && Double.TryParse(dr[dt.Columns.IndexOf(dc)].ToString(), out _value)) //&& Convert.ToDouble(dr[dc.ColumnName]) > 0)
                            {
                                //_value = Math.Round(_value, 1);
                                if (_value > 0)
                                {
                                    List<CustomerOrder> _listCustomerOrder = null;
                                    CustomerOrder _CustomerOrder = null;
                                    ProductOrder _productOrder = null;

                                    bool isNewProductOrder = false;
                                    bool isNewCustomerOrder = false;

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
                                        _productOrder = _PODate.ListProductOrder.Where(x => x.ProductId == _product.ProductId).FirstOrDefault();

                                        Guid _id;
                                        if (dicStore.TryGetValue(dr["StoreCode"].ToString() + (dt.Columns.Contains("P&L") ? dr["P&L"].ToString() : dr["StoreType"].ToString()), out _id))
                                        {
                                            _CustomerOrder = _productOrder.ListCustomerOrder.Where(x => x.CustomerId == _id).FirstOrDefault();
                                        }
                                        else
                                        {
                                            isNewCustomerOrder = true;

                                            Customer _customer;
                                            string sKey = dr["StoreCode"].ToString() + (dt.Columns.Contains("P&L") ? dr["P&L"].ToString() : dr["StoreType"].ToString());
                                            if (!dicCustomer.TryGetValue(sKey, out _customer))
                                            {
                                                _customer = new Customer();

                                                _customer._id = Guid.NewGuid();
                                                _customer.CustomerId = _customer._id;
                                                _customer.CustomerCode = dr["StoreCode"].ToString();
                                                _customer.CustomerName = dr["StoreName"].ToString();
                                                _customer.CustomerRegion = dr["Region"].ToString();
                                                _customer.CustomerType = dt.Columns.Contains("P&L") ? dr["P&L"].ToString() : dr["StoreType"].ToString();
                                                _customer.CustomerBigRegion = PORegion;

                                                Customer.Add(_customer);

                                                dicCustomer.Add(sKey, _customer);
                                            }

                                            Guid _NewGuid = Guid.NewGuid();
                                            _CustomerOrder = new CustomerOrder()
                                            {
                                                _id = _NewGuid,
                                                CustomerOrderId = _NewGuid,
                                                CustomerId = _customer.CustomerId
                                            };

                                            dicPO[dateValue.Date][dr["VE Code"].ToString()].Add(sKey, _customer.CustomerId);
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
                                        string sKey = dr["StoreCode"].ToString() + (dt.Columns.Contains("P&L") ? dr["P&L"].ToString() : dr["StoreType"].ToString());
                                        if (!dicCustomer.TryGetValue(sKey, out _customer))
                                        {
                                            _customer = new Customer();

                                            _customer._id = Guid.NewGuid();
                                            _customer.CustomerId = _customer._id;
                                            _customer.CustomerCode = dr["StoreCode"].ToString();
                                            _customer.CustomerName = dr["StoreName"].ToString();
                                            _customer.CustomerRegion = dr["Region"].ToString();
                                            _customer.CustomerType = dt.Columns.Contains("P&L") ? dr["P&L"].ToString() : dr["StoreType"].ToString();
                                            _customer.CustomerBigRegion = PORegion;

                                            Customer.Add(_customer);

                                            dicCustomer.Add(sKey, _customer);
                                        }

                                        _CustomerOrder = new CustomerOrder();
                                        _CustomerOrder._id = Guid.NewGuid();
                                        _CustomerOrder.CustomerOrderId = _CustomerOrder._id;
                                        _CustomerOrder.CustomerId = _customer.CustomerId;

                                        dicPO[dateValue.Date].Add(dr["VE Code"].ToString(), new Dictionary<string, Guid>());
                                        dicPO[dateValue.Date][dr["VE Code"].ToString()].Add(sKey, _customer.CustomerId);
                                    }

                                    // Filling in data
                                    _listCustomerOrder = _productOrder.ListCustomerOrder;

                                    _CustomerOrder.Unit = ProperUnit((dr.IsNull("Unit") ? "Kg" : dr["Unit"]).ToString());
                                    _CustomerOrder.QuantityOrder += _value;

                                    if (isNewCustomerOrder) { _listCustomerOrder.Add(_CustomerOrder); }

                                    _productOrder.ListCustomerOrder = _listCustomerOrder;
                                    if (isNewProductOrder) { _PODate.ListProductOrder.Add(_productOrder); }
                                }
                            }

                        }

                        _PODate.ListProductOrder = _listProductOrder;

                        if (isNewPODate) { PO.Add(_PODate); }

                        //Debug.WriteLine(Region + " " + dc.ColumnName + ": " + sumColumn);
                        //Debug.WriteLine(PO.Where(x => x == _PODate).FirstOrDefault().ListProductOrder.Sum(po => po.ListCustomerOrder.Sum(co => co.QuantityOrder)));
                        //Debug.WriteLine("MB " + dc.ColumnName + ": " + PO.Where(x => x.DateOrder.Date.ToString() == dc.ColumnName).FirstOrDefault().ListProductOrder.Sum(x => x.ListCustomerOrder.Where(co => dicCustomer.Values.Where(_co => _co.CustomerId == co.CustomerId).FirstOrDefault().CustomerBigRegion == "Miền Bắc").Sum(o => o.QuantityOrder)));
                        //Debug.WriteLine("MN " + dc.ColumnName + ": " + PO.Where(x => x.DateOrder.Date.ToString() == dc.ColumnName).FirstOrDefault().ListProductOrder.Sum(x => x.ListCustomerOrder.Where(co => dicCustomer.Values.Where(_co => _co.CustomerId == co.CustomerId).FirstOrDefault().CustomerBigRegion == "Miền Nam").Sum(o => o.QuantityOrder)));
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
        /// Writing Output to Excel. Interop Style. <para />
        /// Old. Classic. Working. Slow.
        /// </summary>
        /// <param name="dt ( DataTable to be written out )"></param>
        /// <param name="fileName ( Name of destination file )"></param>
        /// <param name="xlApp ( Current Excel.Application )"></param>
        /// <param name="YesNoHeader"></param>
        /// <param name="RowFirst"></param>
        /// <param name="YesNoFirstSheet"></param>
        private void OutputExcel(DataTable dt, string sheetName, Excel.Workbook xlWb, bool YesNoHeader = false, int RowFirst = 6, bool YesNoFirstSheet = false)
        {
            try
            {
                // Open Second Workbook
                //string filePath = string.Format("C:\\Users\\Shirayuki\\Documents\\VinEco\\Project - Chia Hang\\Mastah Project\\{0}", fileName);
                //var xlWb2 = new Aspose.Cells.Workbook(filePath);
                //var xlWs2 = xlWb2.Worksheets[0];

                int rowTotal = dt.Rows.Count;
                int colTotal = dt.Columns.Count;

                if (rowTotal == 0 || colTotal == 0)
                {
                    // For fucking god's sake
                    return;
                }

                //xlWb2.Worksheets.RemoveAt("PO MB");

                //xlApp.DisplayAlerts = false;
                //foreach (Excel.Worksheet _xlWs in xlWb2.Worksheets)
                //{
                //    Debug.WriteLine(_xlWs.Name);
                //    if (_xlWs.Name == sheetName)
                //    {
                //        _xlWs.Delete();
                //    }
                //}
                //xlApp.DisplayAlerts = false;

                //xlWb2.Worksheets.Add(After: xlWb2.Worksheets[xlWb2.Worksheets.Count]);

                //foreach (Excel.Worksheet _xlWs in xlWb.Worksheets)
                //{
                //    Debug.WriteLine(_xlWs.Name);
                //}

                Excel.Worksheet xlWs = null;
                if (YesNoFirstSheet)
                {
                    xlWs = xlWb.Worksheets[0];
                    xlWs.Name = sheetName;
                }
                else
                {
                    xlWs = xlWb.Worksheets[sheetName]; //xlWb2.Worksheets.Count];
                }
                Excel.Range rangeToDelete = (Excel.Range)(xlWs.get_Range("A" + RowFirst, (Excel.Range)(xlWs.Cells[rowTotal, colTotal])));
                rangeToDelete.EntireRow.Delete();

                //int _wsIndex = xlWb2.Worksheets.Add();

                //var xlWs2 = xlWb2.Worksheets[_wsIndex];
                //xlWs2.Name = sheetName;

                //var xlCell2 = xlWs2.Cells;

                #region HeaderStuff
                if (YesNoHeader)
                {
                    object[] Header = new object[colTotal];

                    // column headings               
                    for (int i = 0; i < colTotal; i++)
                        Header[i] = dt.Columns[i].ColumnName;

                    Excel.Range HeaderRange = xlWs.get_Range((Excel.Range)(xlWs.Cells[RowFirst, 1]), (Excel.Range)(xlWs.Cells[1, colTotal]));
                    HeaderRange.Value = Header;
                    HeaderRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
                    HeaderRange.Font.Bold = true;
                }
                #endregion

                int _RowFirst = YesNoHeader ? RowFirst + 1 : RowFirst;

                // Limiting the size of object. If this is too large, expect Out of Memory Exception.
                // Interop pls.
                // Apparently larger yields worse performance. Idk why.
                //var _rowPerBlock = Math.Round(rowTotal / 17, 0);
                //int rowPerBlock = (int)Math.Round(rowTotal / 17d, 0); // 7777;
                int rowPerBlock = 7777;
                //int rowPerBlock = (int)Math.Max(Math.Round(rowTotal / 17d, 0), 7777); // 7777;
                //Debug.WriteLine(rowPerBlock);

                object[,] dbCells = new object[rowPerBlock, colTotal];
                int count = 0;
                int rowPos = 0;
                int rowIndex = 0;
                foreach (DataRow dr in dt.Rows)
                {
                    for (int colIndex = 0; colIndex < colTotal; colIndex++)
                    {
                        //double randomDoubleValue;
                        string _value = (dr[colIndex] ?? String.Empty).ToString();
                        Type _type = dt.Columns[colIndex].DataType;
                        if (dt.Rows[rowIndex][colIndex] != null && _value != "" && _value != "0")
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
                        else
                        {
                            dbCells[rowIndex - rowPos, colIndex] = "";
                        }
                    }
                    count++;
                    if (count >= rowPerBlock)
                    {
                        xlWs.get_Range((Excel.Range)(xlWs.Cells[rowPos + _RowFirst, 1]), (Excel.Range)(xlWs.Cells[rowPos + rowPerBlock + _RowFirst - 1, colTotal])).Formula = dbCells;
                        //xlWs2.Range[rowPos + _RowFirst, 1].Resize[rowPos + rowPerBlock + _RowFirst - 1, colTotal].Value = dbCells;
                        dbCells = new object[Math.Min(rowTotal - rowPos, rowPerBlock), colTotal];
                        count = 0;
                        rowPos = rowIndex + 1;
                    }
                    rowIndex++;
                }

                xlWs.get_Range((Excel.Range)(xlWs.Cells[Math.Max(rowPos + _RowFirst, 2), 1]), (Excel.Range)(xlWs.Cells[rowPos + rowPerBlock + _RowFirst - 1, colTotal])).Formula = dbCells;
                //xlWs2.Range["A" + RowFirst].get_Resize(dbCells.Length[0], dbCells.Length(1)).Value = dbCells;

                dbCells = null;

                if (xlWs != null) { Marshal.ReleaseComObject(xlWs); }
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
        /// Epplus Approach. Failed horribly.
        /// </summary>
        /// <param name="pck"></param>
        /// <param name="dt"></param>
        /// <param name="sheetName"></param>
        /// <param name="YesNoHeader"></param>
        /// <param name="RowFirst"></param>
        /// <param name="YesNoFirstSheet"></param>
        private void OutputExcelEpplus(ExcelPackage pck, DataTable dt, string sheetName, bool YesNoHeader = false, int RowFirst = 6, bool YesNoFirstSheet = false)
        {
            int rowPerBlock = 117000;

            object[,] dataCell = new object[rowPerBlock, dt.Columns.Count];

            int count = 0;
            int rowPos = YesNoHeader ? RowFirst + 1 : RowFirst;

            DataTable dtTemp = new DataTable();

            foreach (DataColumn dc in dt.Columns)
            {
                dtTemp.Columns.Add(dc.ColumnName, dc.DataType);
            }

            foreach (DataRow dr in dt.Rows)
            {
                dtTemp.Rows.Add(dr.ItemArray);

                count++;
                if (count >= rowPerBlock)
                {
                    count = 0;
                    pck.Workbook.Worksheets[sheetName].Cells["A" + rowPos].LoadFromDataTable(dtTemp, false, OfficeOpenXml.Table.TableStyles.None);
                    rowPos += rowPerBlock;
                    dtTemp.Clear();
                }
            }

            pck.Workbook.Worksheets[sheetName].Cells["A" + rowPos].LoadFromDataTable(dtTemp, false, OfficeOpenXml.Table.TableStyles.None);
        }

        /// <summary>
        /// Aspose.Cells Approach. Also failed horribly.
        /// </summary>
        /// <param name="dataTable"></param>
        /// <param name="sheetName"></param>
        /// <param name="xlWb"></param>
        /// <param name="YesNoHeader"></param>
        /// <param name="RowFirst"></param>
        /// <param name="YesNoFirstSheet"></param>
        private void OutputExcelAspose(DataTable dataTable, string sheetName, Aspose.Cells.Workbook xlWb, bool YesNoHeader = false, int RowFirst = 6, bool YesNoFirstSheet = false)
        {
            try
            {
                int rowTotal = dataTable.Rows.Count;
                int colTotal = dataTable.Columns.Count;

                foreach (Aspose.Cells.Worksheet _xlWs in xlWb.Worksheets)
                {
                    Debug.WriteLine(_xlWs.Name);
                }

                Aspose.Cells.Worksheet xlWs = null;
                if (YesNoFirstSheet)
                {
                    xlWs = xlWb.Worksheets[0];
                    xlWs.Name = sheetName;
                }
                else
                {
                    xlWs = xlWb.Worksheets[sheetName]; //xlWb2.Worksheets.Count];
                }
                xlWs.Cells.DeleteRows(RowFirst - 1, rowTotal);

                xlWs.Cells.ImportDataTable(dataTable, true, "A" + RowFirst);

                dataTable = null;

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
        /// Exporting to Excel, using OpenXMLWriter. <para />
        /// Super uber fast. Still have no idea how to use this on an already existing Worksheet. Lel.
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="filename"></param>
        /// <param name="YesNoHeader"></param>
        /// <param name="YesNoZero"></param>
        public static void LargeExport(DataTable dt, string filename, Dictionary<string, int> DicDateCol, bool YesNoHeader = false, bool YesNoZero = false, bool YesNoDateColumn = false)
        {
            try
            {
                using (SpreadsheetDocument document = SpreadsheetDocument.Create(filename, SpreadsheetDocumentType.Workbook))
                {
                    var dicType = new Dictionary<Type, CellValues>();

                    var dicColName = new Dictionary<int, string>();

                    for (int colIndex = 0; colIndex < dt.Columns.Count; colIndex++)
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
                    OpenXmlWriter writer;

                    document.AddWorkbookPart();
                    WorksheetPart workSheetPart = document.WorkbookPart.AddNewPart<WorksheetPart>();

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
                        writer.WriteStartElement(new DocumentFormat.OpenXml.Spreadsheet.Row(), attributes);

                        for (int columnNum = 1; columnNum <= dt.Columns.Count; ++columnNum)
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
                            int _dateValue = 0;
                            if (DateTime.TryParse(dt.Columns[columnNum - 1].ColumnName, out _value))
                            {
                                _dateValue = (int)(_value.Date - new DateTime(1900, 1, 1)).TotalDays + 2;
                            }

                            //write the cell value
                            Cell cell = new Cell()
                            {
                                DataType = type == typeof(double) && _dateValue != 0 ? CellValues.Number : CellValues.String,
                                CellReference = string.Format("{0}{1}", dicColName[columnNum], 1),
                                CellValue = new CellValue(type == typeof(double) && _dateValue != 0 ? _dateValue.ToString() : dt.Columns[columnNum - 1].ColumnName),
                                StyleIndex = (UInt32)(type == typeof(double) && _dateValue != 0 ? 1 : 0)
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

                    for (int rowNum = 1; rowNum <= dt.Rows.Count; rowNum++)
                    {
                        //create a new list of attributes
                        attributes = new List<OpenXmlAttribute>();
                        // add the row index attribute to the list
                        attributes.Add(new OpenXmlAttribute("r", null, (YesNoHeader ? rowNum + 1 : rowNum).ToString()));

                        //write the row start element with the row index attribute
                        writer.WriteStartElement(new DocumentFormat.OpenXml.Spreadsheet.Row(), attributes);

                        DataRow dr = dt.Rows[rowNum - 1];
                        for (int columnNum = 1; columnNum <= dt.Columns.Count; columnNum++)
                        {
                            string colName = dt.Columns[columnNum - 1].ColumnName;
                            Type type = dt.Columns[columnNum - 1].DataType;
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

                            writer.WriteElement(new Cell()
                            {
                                DataType = (type == typeof(string) ? CellValues.String : (YesNoDateColumn == true && type == typeof(DateTime) ? CellValues.Number : dicType[type])),
                                CellReference = string.Format("{0}{1}", dicColName[columnNum], (YesNoHeader ? rowNum + 1 : rowNum)),
                                CellValue = new CellValue(dr[columnNum - 1].ToString()),
                                StyleIndex = (UInt32)(DicDateCol.ContainsKey(colName) ? 1 : 0)
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

                    writer.WriteElement(new Sheet()
                    {
                        Name = "Whatever!",
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


                //ID Format Code
                //0 General
                //1 0
                //2 0.00
                //3 #,##0
                //4 #,##0.00
                //9 0 %
                //10 0.00 %
                //11 0.00E+00
                //12 # ?/?
                //13 # ??/??
                //14 mm - dd - yy
                //15 d - mmm - yy
                //16 d - mmm
                //17 mmm - yy
                //18 h: mm AM/ PM
                //19 h: mm: ss AM/ PM
                //20 h: mm
                //21 h: mm: ss
                //22 m / d / yy h: mm
                //37 #,##0 ;(#,##0)
                //38 #,##0 ;Red
                //39 #,##0.00;(#,##0.00)
                //40 #,##0.00;Red
                //45 mm: ss
                //46[h]:mm: ss
                //47 mmss.0
                //48 ##0.0E+0
                //49 @

                Stylesheet workbookstylesheet = new Stylesheet();

                Font font0 = new Font();         // Default font

                Font font1 = new Font();         // Bold font
                Bold bold = new Bold();
                font1.Append(bold);

                Fonts fonts = new Fonts();      // <APENDING Fonts>
                fonts.Append(font0);
                fonts.Append(font1);

                // <Fills>
                Fill fill0 = new Fill();        // Default fill

                Fills fills = new Fills();      // <APENDING Fills>
                fills.Append(fill0);

                // <Borders>
                Border border0 = new Border();     // Defualt border

                Borders borders = new Borders();    // <APENDING Borders>
                borders.Append(border0);

                NumberingFormat nf2DateTime = new NumberingFormat()
                {
                    NumberFormatId = UInt32Value.FromUInt32(7170),
                    FormatCode = StringValue.FromString("dd-MMM")
                };
                workbookstylesheet.NumberingFormats = new NumberingFormats();
                workbookstylesheet.NumberingFormats.Append(nf2DateTime);

                // <CellFormats>
                CellFormat cellformat0 = new CellFormat()
                {
                    FontId = 0,
                    FillId = 0,
                    BorderId = 0
                }; // Default style : Mandatory | Style ID =0

                CellFormat cellformat1 = new CellFormat()
                {
                    BorderId = 0,
                    FillId = 0,
                    FontId = 0,
                    NumberFormatId = 7170,
                    FormatId = 0,
                    ApplyNumberFormat = true
                };

                CellFormat cellformat2 = new CellFormat()
                {
                    BorderId = 0,
                    FillId = 0,
                    FontId = 0,
                    NumberFormatId = 14,
                    FormatId = 0,
                    ApplyNumberFormat = true
                };

                // <APENDING CellFormats>
                CellFormats cellformats = new CellFormats();
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
        /// OpenWriter Style, for Multiple DataTable into Multiple Worksheets in a single Workbook. A real fucking pain.
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="listDt"></param>
        /// <param name="YesNoHeader"></param>
        /// <param name="YesNoZero"></param>
        public static void LargeExportOneWorkbook(string filePath, List<DataTable> listDt, bool YesNoHeader = false, bool YesNoZero = false)
        {
            try
            {
                using (SpreadsheetDocument document = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook))
                {
                    document.AddWorkbookPart();

                    OpenXmlWriter writer;

                    OpenXmlWriter writerXb;

                    writerXb = OpenXmlWriter.Create(document.WorkbookPart);
                    writerXb.WriteStartElement(new DocumentFormat.OpenXml.Spreadsheet.Workbook());
                    writerXb.WriteStartElement(new Sheets());

                    int count = 0;

                    foreach (DataTable dt in listDt)
                    {

                        var dicType = new Dictionary<Type, CellValues>();

                        var dicColName = new Dictionary<int, string>();

                        for (int colIndex = 0; colIndex < dt.Columns.Count; colIndex++)
                        {
                            dicColName.Add(colIndex + 1, GetColumnName(colIndex + 1));
                        }

                        dicType.Add(typeof(DateTime), CellValues.Date);
                        dicType.Add(typeof(string), CellValues.InlineString);
                        dicType.Add(typeof(double), CellValues.Number);
                        dicType.Add(typeof(int), CellValues.Number);

                        //this list of attributes will be used when writing a start element
                        List<OpenXmlAttribute> attributes;

                        WorksheetPart workSheetPart = document.WorkbookPart.AddNewPart<WorksheetPart>();

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
                            writer.WriteStartElement(new DocumentFormat.OpenXml.Spreadsheet.Row(), attributes);

                            for (int columnNum = 1; columnNum <= dt.Columns.Count; ++columnNum)
                            {
                                Type type = dt.Columns[columnNum - 1].DataType;
                                //reset the list of attributes
                                attributes = new List<OpenXmlAttribute>();
                                // add data type attribute - in this case inline string (you might want to look at the shared strings table)
                                attributes.Add(new OpenXmlAttribute("t", null, "str")); // type == typeof(string) ? "str" : dicType[type].ToString()));
                                                                                        //add the cell reference attribute
                                attributes.Add(new OpenXmlAttribute("r", "", string.Format("{0}{1}", dicColName[columnNum], 1)));

                                //write the cell start element with the type and reference attributes
                                writer.WriteStartElement(new DocumentFormat.OpenXml.Spreadsheet.Cell(), attributes);

                                //write the cell value
                                writer.WriteElement(new CellValue(dt.Columns[columnNum - 1].ColumnName));

                                //writer.WriteElement(new CellValue(string.Format("This is Row {0}, Cell {1}", rowNum, columnNum)));

                                // write the end cell element
                                writer.WriteEndElement();
                            }

                            // write the end row element
                            writer.WriteEndElement();
                        }

                        for (int rowNum = 1; rowNum <= dt.Rows.Count; rowNum++)
                        {
                            //create a new list of attributes
                            attributes = new List<OpenXmlAttribute>();
                            // add the row index attribute to the list
                            attributes.Add(new OpenXmlAttribute("r", null, (YesNoHeader ? rowNum + 1 : rowNum).ToString()));

                            //write the row start element with the row index attribute
                            writer.WriteStartElement(new DocumentFormat.OpenXml.Spreadsheet.Row(), attributes);

                            DataRow dr = dt.Rows[rowNum - 1];
                            for (int columnNum = 1; columnNum <= dt.Columns.Count; columnNum++)
                            {
                                Type type = dt.Columns[columnNum - 1].DataType;
                                //reset the list of attributes
                                attributes = new List<OpenXmlAttribute>();
                                // add data type attribute - in this case inline string (you might want to look at the shared strings table)
                                attributes.Add(new OpenXmlAttribute("t", null, type == typeof(string) ? "str" : dicType[type].ToString()));
                                //add the cell reference attribute
                                attributes.Add(new OpenXmlAttribute("r", "", string.Format("{0}{1}", dicColName[columnNum], (YesNoHeader ? rowNum + 1 : rowNum))));

                                //write the cell start element with the type and reference attributes
                                writer.WriteStartElement(new DocumentFormat.OpenXml.Spreadsheet.Cell(), attributes);

                                //write the cell value
                                if (YesNoZero | dr[columnNum - 1].ToString() != "0")
                                {
                                    writer.WriteElement(new CellValue(dr[columnNum - 1].ToString()));
                                }
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

                        writerXb.WriteElement(new Sheet()
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
            using (SpreadsheetDocument document = SpreadsheetDocument.Create(filename, SpreadsheetDocumentType.Workbook))
            {
                //this list of attributes will be used when writing a start element
                List<OpenXmlAttribute> attributes;
                OpenXmlWriter writer;

                document.AddWorkbookPart();
                WorksheetPart workSheetPart = document.WorkbookPart.AddNewPart<WorksheetPart>();

                writer = OpenXmlWriter.Create(workSheetPart);
                writer.WriteStartElement(new DocumentFormat.OpenXml.Spreadsheet.Worksheet());
                writer.WriteStartElement(new SheetData());

                for (int rowNum = 1; rowNum <= 115000; ++rowNum)
                {
                    //create a new list of attributes
                    attributes = new List<OpenXmlAttribute>();
                    // add the row index attribute to the list
                    attributes.Add(new OpenXmlAttribute("r", null, rowNum.ToString()));

                    //write the row start element with the row index attribute
                    writer.WriteStartElement(new DocumentFormat.OpenXml.Spreadsheet.Row(), attributes);

                    for (int columnNum = 1; columnNum <= 30; ++columnNum)
                    {
                        //reset the list of attributes
                        attributes = new List<OpenXmlAttribute>();
                        // add data type attribute - in this case inline string (you might want to look at the shared strings table)
                        attributes.Add(new OpenXmlAttribute("t", null, "str"));
                        //add the cell reference attribute
                        attributes.Add(new OpenXmlAttribute("r", "", string.Format("{0}{1}", GetColumnName(columnNum), rowNum)));

                        //write the cell start element with the type and reference attributes
                        writer.WriteStartElement(new DocumentFormat.OpenXml.Spreadsheet.Cell(), attributes);
                        //write the cell value
                        writer.WriteElement(new CellValue(string.Format("This is Row {0}, Cell {1}", rowNum, columnNum)));

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

                writer.WriteElement(new Sheet()
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
        /// A simple helper to get the column name from the column index. This is not well tested! <para />
        /// Worked anyway. For a Dictionary anyway.
        /// </summary>
        /// <param name="columnIndex"></param>
        /// <returns></returns>
        private static string GetColumnName(int columnIndex)
        {
            int dividend = columnIndex;
            string columnName = String.Empty;
            int modifier;

            while (dividend > 0)
            {
                modifier = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modifier).ToString() + columnName;
                dividend = (int)((dividend - modifier) / 26);
            }

            return columnName;
        }

        /// <summary>
        /// Convert from one file format to another, using Interop.
        /// Because apparently OpenXML doesn't deal with .xls type ( Including, but not exclusive to .xlsb )
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="PreviousExtension"></param>
        /// <param name="AfterwardExtension"></param>
        /// <param name="YesNoDeleteFile"></param>
        private void ConvertToXlsbInterop(string filePath, string PreviousExtension = "", string AfterwardExtension = "", bool YesNoDeleteFile = false)
        {
            try
            {
                // Remember the list of running Excel.Application.
                // Before initialize xlApp.
                Process[] processBefore = Process.GetProcessesByName("excel");

                // Initialize new instance of Interop Excel.Application.
                Excel.Application xlApp = new Excel.Application();

                // Remember the list of running Excel.Application.
                // After initialize xlApp.
                Process[] processAfter = Process.GetProcessesByName("excel");

                int processID = 0;

                // Compare two lists, get the first and the only process that's not in the 'Before' List.
                foreach (Process process in processAfter)
                {
                    if (!processBefore.Select(p => p.Id).Contains(process.Id))
                    {
                        processID = process.Id;
                        break;
                    }
                }

                xlApp.ScreenUpdating = false;
                xlApp.EnableEvents = false;
                xlApp.DisplayAlerts = false;
                xlApp.DisplayStatusBar = false;
                xlApp.AskToUpdateLinks = false;

                Excel.Workbook xlWb = xlApp.Workbooks.Open(filePath);

                var missing = Type.Missing;
                xlWb.SaveAs(filePath.Replace(PreviousExtension, AfterwardExtension), Excel.XlFileFormat.xlExcel12, missing, missing, false, false, Excel.XlSaveAsAccessMode.xlExclusive, missing, missing, missing);

                xlWb.Close(SaveChanges: false);
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

                Debug.WriteLine(this.Name + " finished peacefully! ");
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
                var xlApp = new Excel.Application();

                var xlWb = xlApp.Workbooks.Open(
                    Filename: filePath,
                    UpdateLinks: false,
                    ReadOnly: false,
                    Format: 5,
                    Password: "",
                    WriteResPassword: "",
                    IgnoreReadOnlyRecommended: true,
                    Origin: Excel.XlPlatform.xlWindows,
                    Delimiter: "",
                    Editable: true,
                    Notify: false,
                    Converter: 0,
                    AddToMru: true,
                    Local: false,
                    CorruptLoad: false);

                xlApp.ScreenUpdating = false;
                xlApp.Calculation = Excel.XlCalculation.xlCalculationManual;
                xlApp.EnableEvents = false;
                xlApp.DisplayAlerts = false;
                xlApp.DisplayStatusBar = false;
                xlApp.AskToUpdateLinks = false;

                foreach (Excel.Worksheet _ws in xlWb.Worksheets)
                {
                    if (_ws.Name == "Evaluation Warning") { _ws.Delete(); }
                }

                xlApp.ScreenUpdating = true;
                xlApp.Calculation = Excel.XlCalculation.xlCalculationAutomatic;
                xlApp.EnableEvents = true;
                xlApp.DisplayAlerts = false;
                xlApp.DisplayStatusBar = true;
                xlApp.AskToUpdateLinks = true;

                xlWb.Close(SaveChanges: true);

                if (xlWb != null) { Marshal.ReleaseComObject(xlWb); }
                xlWb = null;

                xlApp.Quit();
                if (xlApp != null) { Marshal.ReleaseComObject(xlApp); }
                xlApp = null;

                //GC.Collect();
                //GC.WaitForPendingFinalizers();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        #endregion

        #region Proper string
        /// <summary>
        /// Proper a string
        /// </summary>
        public static string ProperStr(string myString)
        {

            // Creates a TextInfo based on the "en-US" culture.
            TextInfo myTI = new CultureInfo("en-US", false).TextInfo;

            //// Changes a string to lowercase.
            //Console.WriteLine("\"{0}\" to lowercase: {1}", myString, myTI.ToLower(myString));

            //// Changes a string to uppercase.
            //Console.WriteLine("\"{0}\" to uppercase: {1}", myString, myTI.ToUpper(myString));

            //// Changes a string to titlecase.
            //Console.WriteLine("\"{0}\" to titlecase: {1}", myString, myTI.ToTitleCase(myString));

            return myTI.ToTitleCase(myString);

        }


        #endregion

        #region Proper UnitType.
        /// <summary>
        /// Proper UnitType.
        /// </summary>
        /// <param name="Unit"></param>
        /// <returns></returns>
        private static string ProperUnit(string Unit)
        {

            // Initialize empty result.
            string _Unit = (Unit).Trim().ToLower();

            // Looping through every letter.
            for (int stringIndex = 0; stringIndex < _Unit.Length; stringIndex++)
            {

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
            }

            // Creates a TextInfo based on the "en-US" culture.
            TextInfo myTI = new CultureInfo("en-US", false).TextInfo;

            // Return the "Proper" Unit.
            return myTI.ToTitleCase(_Unit);
        }
        #endregion

        #region Declaring Model
        //public class PurchaseOrder
        //{
        //    public PurchaseOrder()
        //    {
        //    }
        //    public Guid _id { get; set; }
        //    public Guid PurchaseOrderId { get; set; }
        //    public string PurchaseOrderCode { get; set; }
        //    public List<PurchaseOrderDate> ListPurchaseOrderDate { get; set; }
        //}
        public class CoordResult
        {
            public CoordResult()
            {
            }
            [BsonId]
            public Guid _id { get; set; }
            public Guid CoordResultId { get; set; }
            [BsonDateTimeOptions(Kind = DateTimeKind.Local)]
            public DateTime DateOrder { get; set; }
            public List<CoordResultDate> ListCoordResultDate { get; set; }

        }
        public class CoordResultDate
        {
            public CoordResultDate()
            {
            }
            [BsonId]
            public Guid _id { get; set; }
            public Guid CoordResultDateId { get; set; }
            public Guid ProductId { get; set; }
            public List<CoordinateDate> ListCoordinateDate { get; set; }

        }
        public class CoordinateDate
        {
            public CoordinateDate()
            {
            }
            [BsonId]
            public Guid _id { get; set; }
            public Guid CoordinateDateId { get; set; }
            public Guid CustomerOrderId { get; set; }
            public Guid? SupplierOrderId { get; set; }
            [BsonDateTimeOptions(Kind = DateTimeKind.Local)]
            public DateTime? DateDelier { get; set; }
        }
        public class PurchaseOrderDate
        {
            public PurchaseOrderDate()
            {
            }
            [BsonId]
            public Guid _id { get; set; }
            public Guid PurchaseOrderDateId { get; set; }
            [BsonDateTimeOptions(Kind = DateTimeKind.Local)]
            public DateTime DateOrder { get; set; }
            public List<ProductOrder> ListProductOrder { get; set; }
        }
        public class ForecastDate
        {
            public ForecastDate()
            {
            }
            [BsonId]
            public Guid _id { get; set; }
            public Guid ForecastDateId { get; set; }
            [BsonDateTimeOptions(Kind = DateTimeKind.Local)]
            public DateTime DateForecast { get; set; }
            public List<ProductForecast> ListProductForecast { get; set; }
        }
        public class Product
        {
            public Product()
            {
            }
            [BsonId]
            public Guid _id { get; set; }
            public Guid ProductId { get; set; }
            public string ProductCode { get; set; }
            public string ProductName { get; set; }
            public string ProductVECode { get; set; }
            public string ProductClassification { get; set; }
        }
        public class ProductUnit
        {
            public ProductUnit()
            {
            }
            [BsonId]
            public Guid _id { get; set; }
            public Guid ProductId { get; set; }
            public string ProductCode { get; set; }
            public List<ProductUnitRegion> ListRegion { get; set; }

        }
        public class ProductUnitRegion
        {
            [BsonId]
            public Guid _id { get; set; }
            public string Region { get; set; }
            public string OrderUnitType { get; set; }
            public double OrderUnitPer { get; set; }
            public string SaleUnitType { get; set; }
            public double SaleUnitPer { get; set; }
        }
        public class ProductOrder
        {
            public ProductOrder()
            {
            }
            [BsonId]
            public Guid _id { get; set; }
            public Guid ProductOrderId { get; set; }
            public Guid ProductId { get; set; }
            public List<CustomerOrder> ListCustomerOrder { get; set; }
        }
        public class ProductForecast
        {
            public ProductForecast()
            {
            }
            [BsonId]
            public Guid _id { get; set; }
            public Guid ProductForecastId { get; set; }
            public Guid ProductId { get; set; }
            public List<SupplierForecast> ListSupplierForecast { get; set; }
        }
        public class CustomerOrder
        {
            public CustomerOrder()
            {
            }
            [BsonId]
            public Guid _id { get; set; }
            public Guid CustomerOrderId { get; set; }
            public Guid CustomerId { get; set; }
            public string Company { get; set; }
            public string Unit { get; set; }
            public double QuantityOrder { get; set; }
            public double QuantityOrderKg { get; set; }
        }
        public class Customer
        {
            public Customer()
            {
            }
            [BsonId]
            public Guid _id { get; set; }
            public Guid CustomerId { get; set; }
            public string CustomerCode { get; set; }
            public string CustomerName { get; set; }
            public string CustomerType { get; set; }
            public string CustomerRegion { get; set; }
            public string CustomerBigRegion { get; set; }
        }
        public class Supplier
        {
            public Supplier()
            {
            }
            [BsonId]
            public Guid _id { get; set; }
            public Guid SupplierId { get; set; }
            public string SupplierCode { get; set; }
            public string SupplierName { get; set; }
            public string SupplierType { get; set; }
            public string SupplierRegion { get; set; }
        }
        public class SupplierForecast
        {
            public SupplierForecast()
            {
            }
            [BsonId]
            public Guid _id { get; set; }
            public Guid SupplierForecastId { get; set; }
            public Guid SupplierId { get; set; }
            public bool LabelVinEco { get; set; }
            public bool FullOrder { get; set; }
            public bool QualityControlPass { get; set; }
            public bool CrossRegion { get; set; }
            public byte level { get; set; }
            public string Availability { get; set; }
            public double QuantityForecast { get; set; }
        }

        public class CoordStructure
        {
            public Dictionary<DateTime, Dictionary<Product, Dictionary<CustomerOrder, bool>>> dicPO;
            public Dictionary<DateTime, Dictionary<Product, Dictionary<SupplierForecast, bool>>> dicFC;
            public Dictionary<DateTime, Dictionary<Product, Dictionary<CustomerOrder, Dictionary<SupplierForecast, DateTime>>>> dicCoord;
            public Dictionary<Guid, Product> dicProduct;
            public Dictionary<Guid, Supplier> dicSupplier;
            public Dictionary<Guid, Customer> dicCustomer;
            public Dictionary<string, ProductUnit> dicProductUnit;
        }
        #endregion

        #region GenProduct
        //public async void GenProduct()
        //{
        //    var dicProduct = new Dictionary<string, string>();
        //    var Product = new List<Product>();
        //    #region MasterList
        //    dicProduct.Add("A00101", "Bắp cải bao tử");
        //    dicProduct.Add("A00201", "Bắp cải tím");
        //    dicProduct.Add("A00203", "Bắp cải tím HC");
        //    dicProduct.Add("A00301", "Bắp cải trái tim");
        //    dicProduct.Add("A00401", "Bắp cải trắng");
        //    dicProduct.Add("A00403", "Bắp cải trắng HC");
        //    dicProduct.Add("A00501", "Cải bẹ dưa (cải sậy)");
        //    dicProduct.Add("A00601", "Cải bó xôi");
        //    dicProduct.Add("A00603", "Cải bó xôi HC");
        //    dicProduct.Add("A00701", "Cải bó xôi baby");
        //    dicProduct.Add("A00703", "Cải bó xôi baby HC");
        //    dicProduct.Add("A00801", "Cải cầu vồng");
        //    dicProduct.Add("A00901", "Cải chíp (cải thìa)");
        //    dicProduct.Add("A00903", "Cải chíp (cải thìa) HC");
        //    dicProduct.Add("A01001", "Cải chíp (cải thìa) baby");
        //    dicProduct.Add("A01101", "Cải củ ăn lá");
        //    dicProduct.Add("A01201", "Cải củ baby");
        //    dicProduct.Add("A01301", "Cải củ dưa");
        //    dicProduct.Add("A01401", "Cải cúc (tần ô)");
        //    dicProduct.Add("A01403", "Cải cúc (tần ô) HC");
        //    dicProduct.Add("A01501", "Cải cúc (tần ô) baby");
        //    dicProduct.Add("A01601", "Cải dún");
        //    dicProduct.Add("A01701", "Cải đuôi phụng đỏ");
        //    dicProduct.Add("A01703", "Cải đuôi phụng đỏ HC");
        //    dicProduct.Add("A01801", "Cải đuôi phụng xanh");
        //    dicProduct.Add("A01803", "Cải đuôi phụng xanh HC");
        //    dicProduct.Add("A01901", "Cải làn (cải rổ)");
        //    dicProduct.Add("A01903", "Cải làn (cải rổ) HC");
        //    dicProduct.Add("A02001", "Cải mèo");
        //    dicProduct.Add("A02003", "Cải mèo HC");
        //    dicProduct.Add("A02101", "Cải ngồng (cải ngọt bông)");
        //    dicProduct.Add("A02103", "Cải ngồng (cải ngọt bông) HC");
        //    dicProduct.Add("A02201", "Cải ngồng (cải ngọt bông) baby");
        //    dicProduct.Add("A02301", "Cải ngọt");
        //    dicProduct.Add("A02303", "Cải ngọt HC");
        //    dicProduct.Add("A02401", "Cải ngọt baby");
        //    dicProduct.Add("A02501", "Cải thảo");
        //    dicProduct.Add("A02601", "Cải thảo hỏa tiễn");
        //    dicProduct.Add("A02701", "Cải xanh (cải canh)");
        //    dicProduct.Add("A02703", "Cải xanh (cải canh) HC");
        //    dicProduct.Add("A02801", "Cải xanh (cải canh) baby");
        //    dicProduct.Add("A02803", "Cải xanh (cải canh) baby HC");
        //    dicProduct.Add("A02901", "Cải xoăn (cải kale)");
        //    dicProduct.Add("A03001", "Cải xoong (xà lách xoong)");
        //    dicProduct.Add("A03101", "Chùm ngây");
        //    dicProduct.Add("A03103", "Chùm ngây HC");
        //    dicProduct.Add("A03201", "Mồng tơi");
        //    dicProduct.Add("A03203", "Mồng tơi HC");
        //    dicProduct.Add("A03301", "Ngải cứu");
        //    dicProduct.Add("A03303", "Ngải cứu HC");
        //    dicProduct.Add("A03401", "Ngọn bí");
        //    dicProduct.Add("A03403", "Ngọn bí HC");
        //    dicProduct.Add("A03501", "Ngọn su su");
        //    dicProduct.Add("A03601", "Rau đắng");
        //    dicProduct.Add("A03701", "Rau đay đỏ");
        //    dicProduct.Add("A03801", "Rau đay xanh");
        //    dicProduct.Add("A03901", "Rau dền");
        //    dicProduct.Add("A03903", "Rau dền HC");
        //    dicProduct.Add("A04001", "Rau dền cơm");
        //    dicProduct.Add("A04003", "Rau dền cơm HC");
        //    dicProduct.Add("A04101", "Rau dền đỏ");
        //    dicProduct.Add("A04103", "Rau dền đỏ HC");
        //    dicProduct.Add("A04201", "Rau dền tía");
        //    dicProduct.Add("A04301", "Rau dền xanh");
        //    dicProduct.Add("A04401", "Rau diếp");
        //    dicProduct.Add("A04501", "Rau lang");
        //    dicProduct.Add("A04503", "Rau lang HC");
        //    dicProduct.Add("A04601", "Rau lủi rừng");
        //    dicProduct.Add("A04701", "Rau má");
        //    dicProduct.Add("A04801", "Rau muống");
        //    dicProduct.Add("A04803", "Rau muống HC");
        //    dicProduct.Add("A04901", "Rau muống baby");
        //    dicProduct.Add("A05001", "Rau muống cọng");
        //    dicProduct.Add("A05101", "Rau ngót");
        //    dicProduct.Add("A05103", "Rau ngót HC");
        //    dicProduct.Add("A05201", "Rau ngót nhật");
        //    dicProduct.Add("A05203", "Rau ngót nhật HC");
        //    dicProduct.Add("A05301", "Xà lách batavia tím");
        //    dicProduct.Add("A05401", "Xà lách burnet");
        //    dicProduct.Add("A05501", "Xà lách carol");
        //    dicProduct.Add("A05601", "Xà lách frisse");
        //    dicProduct.Add("A05701", "Xà lách iceberg");
        //    dicProduct.Add("A05801", "Xà lách lolo tím");
        //    dicProduct.Add("A05901", "Xà lách lolo xanh");
        //    dicProduct.Add("A05903", "Xà lách lolo xanh HC");
        //    dicProduct.Add("A06001", "Xà lách mỡ");
        //    dicProduct.Add("A06003", "Xà lách mỡ HC");
        //    dicProduct.Add("A06101", "Xà lách oakleaf đỏ");
        //    dicProduct.Add("A06201", "Xà lách oakleaf xanh");
        //    dicProduct.Add("A06301", "Xà lách radicchio (tím búp)");
        //    dicProduct.Add("A06401", "Xà lách rocket");
        //    dicProduct.Add("A06501", "Xà lách romaine");
        //    dicProduct.Add("A06601", "Xà lách romaine baby");
        //    dicProduct.Add("A06701", "Xà lách salanova đỏ");
        //    dicProduct.Add("A06801", "Xà lách salanova xanh");
        //    dicProduct.Add("A06901", "Cải nhật đỏ");
        //    dicProduct.Add("A07001", "Xà lách cuộn (xà lách ta)");
        //    dicProduct.Add("A07003", "Xà lách cuộn (xà lách ta) HC");
        //    dicProduct.Add("A07101", "Bắp cải chồi");
        //    dicProduct.Add("A07201", "Mồng tơi đỏ");
        //    dicProduct.Add("A07301", "Rau đay đỏ baby");
        //    dicProduct.Add("A07401", "Rau dền baby");
        //    dicProduct.Add("A07501", "Cải thảo xanh");
        //    dicProduct.Add("A07601", "Rau húng hỗn hợp");
        //    dicProduct.Add("A07701", "Rau nêm nấu canh chua (ngò gai, om, húng quế, hành)");
        //    dicProduct.Add("B00101", "Bông so đũa");
        //    dicProduct.Add("B00201", "Dọc mùng (bạc hà)");
        //    dicProduct.Add("B00203", "Dọc mùng (bạc hà) HC");
        //    dicProduct.Add("B00301", "Hành tây");
        //    dicProduct.Add("B00401", "Hành tây lột vỏ");
        //    dicProduct.Add("B00501", "Hành tây tím");
        //    dicProduct.Add("B00601", "Hoa atiso");
        //    dicProduct.Add("B00701", "Hoa bí");
        //    dicProduct.Add("B00801", "Hoa chuối");
        //    dicProduct.Add("B00901", "Hoa điên điển");
        //    dicProduct.Add("B01001", "Hoa kim châm");
        //    dicProduct.Add("B01101", "Hoa thiên lý");
        //    dicProduct.Add("B01201", "Măng tây tím");
        //    dicProduct.Add("B01301", "Măng tây trắng");
        //    dicProduct.Add("B01401", "Măng tây xanh");
        //    dicProduct.Add("B01501", "Ngồng tỏi");
        //    dicProduct.Add("B01601", "Nha đam");
        //    dicProduct.Add("B01701", "Su hào tím");
        //    dicProduct.Add("B01801", "Su hào xanh");
        //    dicProduct.Add("B01803", "Su hào xanh HC");
        //    dicProduct.Add("B01901", "Súp lơ (bông cải) ngọc bích");
        //    dicProduct.Add("B02001", "Súp lơ (bông cải) san hô");
        //    dicProduct.Add("B02101", "Súp lơ (bông cải) trắng");
        //    dicProduct.Add("B02103", "Súp lơ (bông cải) trắng HC");
        //    dicProduct.Add("B02201", "Súp lơ (bông cải) xanh");
        //    dicProduct.Add("B02203", "Súp lơ (bông cải) xanh HC");
        //    dicProduct.Add("B02301", "Súp lơ (bông cải) xanh baby");
        //    dicProduct.Add("B02401", "Ngó sen");
        //    dicProduct.Add("B02501", "Măng củ");
        //    dicProduct.Add("B02601", "Măng lá");
        //    dicProduct.Add("C00101", "Bầu sao");
        //    dicProduct.Add("C00201", "Bầu xanh");
        //    dicProduct.Add("C00203", "Bầu xanh HC");
        //    dicProduct.Add("C00301", "Bí đao chanh");
        //    dicProduct.Add("C00401", "Bí đỏ dài");
        //    dicProduct.Add("C00501", "Bí đỏ giống nhật");
        //    dicProduct.Add("C00503", "Bí đỏ giống nhật HC");
        //    dicProduct.Add("C00601", "Bí đỏ hồ lô");
        //    dicProduct.Add("C00701", "Bí đỏ tròn");
        //    dicProduct.Add("C00801", "Bí ngô non");
        //    dicProduct.Add("C00803", "Bí ngô non HC");
        //    dicProduct.Add("C00901", "Bí ngồi vàng");
        //    dicProduct.Add("C01001", "Bí ngồi xanh");
        //    dicProduct.Add("C01101", "Bí xanh (bí đao)");
        //    dicProduct.Add("C01103", "Bí xanh (bí đao) HC");
        //    dicProduct.Add("C01201", "Cà dĩa");
        //    dicProduct.Add("C01301", "Cà bát trắng");
        //    dicProduct.Add("C01401", "Cà bát xanh");
        //    dicProduct.Add("C01501", "Cà chua beef");
        //    dicProduct.Add("C01601", "Cà chua cherry chùm đỏ (dài)");
        //    dicProduct.Add("C01701", "Cà chua cherry chùm đỏ (tròn)");
        //    dicProduct.Add("C01801", "Cà chua cherry đỏ (dài)");
        //    dicProduct.Add("C01901", "Cà chua cherry đỏ (tròn)");
        //    dicProduct.Add("C02001", "Cà chua cherry socola (dài)");
        //    dicProduct.Add("C02101", "Cà chua cherry socola (tròn)");
        //    dicProduct.Add("C02201", "Cà chua cherry tím (dài)");
        //    dicProduct.Add("C02301", "Cà chua cherry tím (tròn)");
        //    dicProduct.Add("C02401", "Cà chua cherry vàng (dài)");
        //    dicProduct.Add("C02501", "Cà chua cherry vàng (tròn)");
        //    dicProduct.Add("C02601", "Cà chua cocktail");
        //    dicProduct.Add("C02701", "Cà chua đen");
        //    dicProduct.Add("C02801", "Cà chua đỏ");
        //    dicProduct.Add("C02803", "Cà chua đỏ HC");
        //    dicProduct.Add("C02901", "Cà chua farcies");
        //    dicProduct.Add("C03001", "Cà chua lê đỏ");
        //    dicProduct.Add("C03101", "Cà chua lê vàng");
        //    dicProduct.Add("C03201", "Cà chua roseberry đỏ");
        //    dicProduct.Add("C03301", "Cà chua roseberry vàng");
        //    dicProduct.Add("C03401", "Cà chua tím");
        //    dicProduct.Add("C03501", "Cà pháo tím");
        //    dicProduct.Add("C03601", "Cà pháo trắng");
        //    dicProduct.Add("C03701", "Cà pháo xanh");
        //    dicProduct.Add("C03801", "Cà tím dài");
        //    dicProduct.Add("C03803", "Cà tím dài HC");
        //    dicProduct.Add("C03901", "Cà tím đen dài");
        //    dicProduct.Add("C04001", "Cà tím giống nhật");
        //    dicProduct.Add("C04101", "Cà tím giống thái");
        //    dicProduct.Add("C04201", "Cà tím tròn");
        //    dicProduct.Add("C04203", "Cà tím tròn HC");
        //    dicProduct.Add("C04301", "Cà xanh dài");
        //    dicProduct.Add("C04401", "Đậu bắp xanh");
        //    dicProduct.Add("C04403", "Đậu bắp xanh HC");
        //    dicProduct.Add("C04501", "Đậu cove baby");
        //    dicProduct.Add("C04601", "Đậu cove nhật");
        //    dicProduct.Add("C04701", "Đậu cove vàng");
        //    dicProduct.Add("C04703", "Đậu cove vàng HC");
        //    dicProduct.Add("C04801", "Đậu cove xanh");
        //    dicProduct.Add("C04803", "Đậu cove xanh HC");
        //    dicProduct.Add("C04901", "Đậu đũa");
        //    dicProduct.Add("C04903", "Đậu đũa HC");
        //    dicProduct.Add("C05001", "Đậu hà lan");
        //    dicProduct.Add("C05101", "Đậu ngọt");
        //    dicProduct.Add("C05201", "Đậu phộng (lạc) tươi");
        //    dicProduct.Add("C05301", "Đậu rồng");
        //    dicProduct.Add("C05401", "Đậu tương nhật");
        //    dicProduct.Add("C05501", "Đậu ván");
        //    dicProduct.Add("C05601", "Dưa chuột (dưa leo)");
        //    dicProduct.Add("C05701", "Dưa chuột baby (dưa leo baby)");
        //    dicProduct.Add("C05801", "Dưa chuột dài (dưa leo dài)");
        //    dicProduct.Add("C05901", "Lặc lày (mướp nhật)");
        //    dicProduct.Add("C06001", "Mướp đài loan");
        //    dicProduct.Add("C06101", "Mướp đắng (khổ qua)");
        //    dicProduct.Add("C06201", "Mướp đắng gai");
        //    dicProduct.Add("C06301", "Mướp đắng rừng (khổ qua rừng)");
        //    dicProduct.Add("C06401", "Mướp hương");
        //    dicProduct.Add("C06501", "Mướp khía");
        //    dicProduct.Add("C06601", "Ngô bao tử");
        //    dicProduct.Add("C06701", "Ngô nếp");
        //    dicProduct.Add("C06801", "Ngô ngọt");
        //    dicProduct.Add("C06901", "Ớt ngọt (ớt chuông) baby 3 màu");
        //    dicProduct.Add("C07001", "Ớt ngọt (ớt chuông) cam");
        //    dicProduct.Add("C07101", "Ớt ngọt (ớt chuông) đỏ");
        //    dicProduct.Add("C07201", "Ớt ngọt (ớt chuông) vàng");
        //    dicProduct.Add("C07301", "Ớt ngọt (ớt chuông) xanh");
        //    dicProduct.Add("C07401", "Su su baby");
        //    dicProduct.Add("C07501", "Su su quả");
        //    dicProduct.Add("C07601", "Bí nhật non (bí nhật bao tử)");
        //    dicProduct.Add("C07701", "Ớt ngọt (ớt chuông) cam dài");
        //    dicProduct.Add("C07801", "Ớt ngọt (ớt chuông) đỏ dài");
        //    dicProduct.Add("C07901", "Ớt ngọt (ớt chuông) vàng dài");
        //    dicProduct.Add("C08001", "Ớt ngọt (ớt chuông) xanh dài");
        //    dicProduct.Add("C08101", "Bầu hồ lô");
        //    dicProduct.Add("C08201", "Cà tím dài baby");
        //    dicProduct.Add("C08301", "Đậu bắp đỏ");
        //    dicProduct.Add("C08401", "Đậu bắp tím");
        //    dicProduct.Add("C08501", "Đậu bắp trắng");
        //    dicProduct.Add("C08601", "Đậu đũa baby");
        //    dicProduct.Add("C08701", "Đậu hà lan baby");
        //    dicProduct.Add("C08801", "Dưa chuột (dưa leo) giống hà lan");
        //    dicProduct.Add("C08901", "Dưa chuột nếp (dưa leo nếp)");
        //    dicProduct.Add("C09001", "Cà chua cherry hỗn hợp");
        //    dicProduct.Add("C09101", "Ớt ngọt dài");
        //    dicProduct.Add("D00101", "Cà rốt");
        //    dicProduct.Add("D00201", "Cà rốt baby");
        //    dicProduct.Add("D00301", "Củ cải đỏ");
        //    dicProduct.Add("D00401", "Củ cải trắng");
        //    dicProduct.Add("D00403", "Củ cải trắng HC");
        //    dicProduct.Add("D00501", "Củ đậu (củ sắn nước)");
        //    dicProduct.Add("D00601", "Củ dền đỏ");
        //    dicProduct.Add("D00701", "Củ mỡ (khoai mỡ) tím");
        //    dicProduct.Add("D00801", "Củ mỡ (khoai mỡ) trắng");
        //    dicProduct.Add("D00901", "Củ năng (mã thầy)");
        //    dicProduct.Add("D01001", "Củ sen");
        //    dicProduct.Add("D01101", "Củ từ (khoai từ)");
        //    dicProduct.Add("D01201", "Khoai lang nhật");
        //    dicProduct.Add("D01301", "Khoai lang ruột vàng");
        //    dicProduct.Add("D01401", "Khoai lang tím");
        //    dicProduct.Add("D01501", "Khoai môn");
        //    dicProduct.Add("D01601", "Khoai môn sáp");
        //    dicProduct.Add("D01701", "Khoai sọ");
        //    dicProduct.Add("D01801", "Khoai tây bi");
        //    dicProduct.Add("D01901", "Khoai tây hồng");
        //    dicProduct.Add("D02001", "Khoai tây vàng");
        //    dicProduct.Add("D02101", "Sắn (khoai mì)");
        //    dicProduct.Add("D02201", "Cà rốt rainbow");
        //    dicProduct.Add("D02301", "Cà rốt tím");
        //    dicProduct.Add("D02401", "Cà rốt vàng");
        //    dicProduct.Add("D02501", "Củ cải trắng baby");
        //    dicProduct.Add("D02601", "Củ cải trắng ruột đỏ");
        //    dicProduct.Add("D02701", "Củ dền tím");
        //    dicProduct.Add("D02801", "Củ dền vàng");
        //    dicProduct.Add("D02901", "Khoai môn tím");
        //    dicProduct.Add("D03001", "Khoai môn trắng");
        //    dicProduct.Add("D03101", "Khoai sáp vàng");
        //    dicProduct.Add("D03201", "Khoai sọ ruột tím");
        //    dicProduct.Add("D03301", "Khoai sọ ruột vàng");
        //    dicProduct.Add("D03401", "Khoai tây tím");
        //    dicProduct.Add("D03501", "Khoai lang bi");
        //    dicProduct.Add("E00101", "Đậu bích ngọc hạt tươi");
        //    dicProduct.Add("E00201", "Đậu đỏ tươi");
        //    dicProduct.Add("E00301", "Đậu ngự hạt tươi");
        //    dicProduct.Add("E00401", "Đậu petit pois");
        //    dicProduct.Add("E00501", "Đậu trắng");
        //    dicProduct.Add("E00601", "Hạt sen");
        //    dicProduct.Add("F00101", "Bạc hà tây");
        //    dicProduct.Add("F00201", "Bông hẹ");
        //    dicProduct.Add("F00301", "Cần tây lớn");
        //    dicProduct.Add("F00401", "Cần tây nhỏ (cần tàu )");
        //    dicProduct.Add("F00403", "Cần tây nhỏ (cần tàu ) HC");
        //    dicProduct.Add("F00501", "Củ hồi");
        //    dicProduct.Add("F00601", "Diếp cá");
        //    dicProduct.Add("F00701", "Đinh lăng");
        //    dicProduct.Add("F00801", "Gừng tươi");
        //    dicProduct.Add("F00901", "Hành lá");
        //    dicProduct.Add("F01001", "Hành paro");
        //    dicProduct.Add("F01101", "Hẹ lá");
        //    dicProduct.Add("F01201", "Húng bạc hà (húng lủi)");
        //    dicProduct.Add("F01301", "Húng cay");
        //    dicProduct.Add("F01401", "Húng đỏ");
        //    dicProduct.Add("F01501", "Húng hương tây");
        //    dicProduct.Add("F01601", "Húng láng");
        //    dicProduct.Add("F01701", "Húng quế");
        //    dicProduct.Add("F01801", "Húng xanh");
        //    dicProduct.Add("F01901", "Hương thảo tây");
        //    dicProduct.Add("F02001", "Kiệu tươi");
        //    dicProduct.Add("F02101", "Kinh giới");
        //    dicProduct.Add("F02201", "Kinh giới tây");
        //    dicProduct.Add("F02301", "Lá chanh");
        //    dicProduct.Add("F02401", "Lá giang");
        //    dicProduct.Add("F02501", "Lá lốt");
        //    dicProduct.Add("F02601", "Lá mơ");
        //    dicProduct.Add("F02701", "Lá móc mật");
        //    dicProduct.Add("F02801", "Lá sung");
        //    dicProduct.Add("F02901", "Mùi chocolate mint");
        //    dicProduct.Add("F03001", "Mùi ta (ngò rí)");
        //    dicProduct.Add("F03101", "Mùi taragon");
        //    dicProduct.Add("F03201", "Mùi tàu (ngò gai)");
        //    dicProduct.Add("F03301", "Mùi tây (ngò rí tây)");
        //    dicProduct.Add("F03401", "Nghệ");
        //    dicProduct.Add("F03501", "Ngổ");
        //    dicProduct.Add("F03601", "Nhút chua");
        //    dicProduct.Add("F03701", "Oregano");
        //    dicProduct.Add("F03801", "Ớt cay chỉ thiên đỏ");
        //    dicProduct.Add("F03901", "Ớt cay chỉ thiên vàng");
        //    dicProduct.Add("F04001", "Ớt cay chỉ thiên xanh");
        //    dicProduct.Add("F04101", "Ớt hiểm đỏ");
        //    dicProduct.Add("F04201", "Ớt hiểm xanh");
        //    dicProduct.Add("F04301", "Ớt sừng đỏ");
        //    dicProduct.Add("F04401", "Ớt sừng xanh");
        //    dicProduct.Add("F04501", "Quế tây");
        //    dicProduct.Add("F04601", "Quế tây lá nhỏ");
        //    dicProduct.Add("F04701", "Quế tây tím");
        //    dicProduct.Add("F04801", "Răm");
        //    dicProduct.Add("F04901", "Rau thơm hỗn hợp");
        //    dicProduct.Add("F05001", "Riềng");
        //    dicProduct.Add("F05101", "Sả");
        //    dicProduct.Add("F05201", "Thì là");
        //    dicProduct.Add("F05301", "Tía tô");
        //    dicProduct.Add("F05401", "Tía tô nhật đỏ");
        //    dicProduct.Add("F05501", "Tía tô nhật xanh");
        //    dicProduct.Add("F05601", "Tiêu non");
        //    dicProduct.Add("F05701", "Tỏi tây");
        //    dicProduct.Add("F05703", "Tỏi tây HC");
        //    dicProduct.Add("F05801", "Xạ hương tây");
        //    dicProduct.Add("F05901", "Xô thơm tây");
        //    dicProduct.Add("F06001", "Lá dứa");
        //    dicProduct.Add("F06101", "Lá nguyệt quế (lá bay)");
        //    dicProduct.Add("F06201", "Lavender");
        //    dicProduct.Add("F06301", "Ớt sừng cam");
        //    dicProduct.Add("F06401", "Ớt sừng vàng");
        //    dicProduct.Add("F06501", "Hành hàn quốc");
        //    dicProduct.Add("F06601", "Củ hẹ");
        //    dicProduct.Add("F06701", "Thì là tây");
        //    dicProduct.Add("F06801", "Bị trùng nên bỏ Mùi tây (Ngò tây)");
        //    dicProduct.Add("F06901", "Hành chăm");
        //    dicProduct.Add("F07001", "Hành củ");
        //    dicProduct.Add("F07101", "Tỏi");
        //    dicProduct.Add("G00101", "Cải bó xôi baby TC");
        //    dicProduct.Add("G00201", "Cải chíp (cải thìa) TC");
        //    dicProduct.Add("G00301", "Cải đuôi phụng đỏ TC");
        //    dicProduct.Add("G00401", "Cải đuôi phụng xanh TC");
        //    dicProduct.Add("G00501", "Cải ngọt TC");
        //    dicProduct.Add("G00601", "Củ dền đỏ TC");
        //    dicProduct.Add("G00701", "Húng xanh TC");
        //    dicProduct.Add("G00801", "Húng xanh Thái TC");
        //    dicProduct.Add("G00901", "Mù tạt đỏ thủy canh TC");
        //    dicProduct.Add("G01001", "Mù tạt xanh thủy canh TC");
        //    dicProduct.Add("G01101", "Rau diếp TC");
        //    dicProduct.Add("G01201", "Rau muống TC");
        //    dicProduct.Add("G01301", "Xà lách batavia tím TC");
        //    dicProduct.Add("G01401", "Xà lách burnet TC");
        //    dicProduct.Add("G01501", "Xà lách frisse TC");
        //    dicProduct.Add("G01601", "Xà lách iceberg TC");
        //    dicProduct.Add("G01701", "Xà lách lolo tím TC");
        //    dicProduct.Add("G01801", "Xà lách lolo xanh TC");
        //    dicProduct.Add("G01901", "Xà lách mỡ TC");
        //    dicProduct.Add("G02001", "Xà lách oakleaf đỏ TC");
        //    dicProduct.Add("G02101", "Xà lách oakleaf xanh TC");
        //    dicProduct.Add("G02201", "Xà lách romaine TC");
        //    dicProduct.Add("G02301", "Xà lách salanova đỏ TC");
        //    dicProduct.Add("G02401", "Xà lách salanova xanh TC");
        //    dicProduct.Add("G02501", "Xà lách Romain Baby TC");
        //    dicProduct.Add("G02601", "Xà lách radicchio (tím búp) TC");
        //    dicProduct.Add("G02701", "xà lách Leonardo TC");
        //    dicProduct.Add("G02801", "Cải xanh (cải canh) TC");
        //    dicProduct.Add("G02901", "xà lách Indigo tím TC");
        //    dicProduct.Add("G03001", "Xà lách carol TC");
        //    dicProduct.Add("H00101", "Giá đỗ");
        //    dicProduct.Add("H00201", "Rau mầm cải chíp");
        //    dicProduct.Add("H00301", "Rau mầm cải củ trắng");
        //    dicProduct.Add("H00401", "Rau mầm cải đuôi phụng đỏ");
        //    dicProduct.Add("H00501", "Rau mầm cải đuôi phụng xanh");
        //    dicProduct.Add("H00601", "Rau mầm cải ngồng");
        //    dicProduct.Add("H00701", "Rau mầm cải ngọt");
        //    dicProduct.Add("H00801", "Rau mầm cải tím");
        //    dicProduct.Add("H00901", "Rau mầm cải xanh");
        //    dicProduct.Add("H01001", "Rau mầm cải xoong");
        //    dicProduct.Add("H01101", "Rau mầm cỏ thơm");
        //    dicProduct.Add("H01201", "Rau mầm củ dền");
        //    dicProduct.Add("H01301", "Rau mầm đậu hà lan");
        //    dicProduct.Add("H01401", "Rau mầm dền đỏ");
        //    dicProduct.Add("H01501", "Rau mầm dền tía");
        //    dicProduct.Add("H01601", "Rau mầm húng đỏ");
        //    dicProduct.Add("H01701", "Rau mầm húng xanh");
        //    dicProduct.Add("H01801", "Rau mầm húng xanh Thái");
        //    dicProduct.Add("H01901", "Rau mầm hướng dương");
        //    dicProduct.Add("H02001", "Rau mầm mù tạt đỏ");
        //    dicProduct.Add("H02101", "Rau mầm mù tạt xanh");
        //    dicProduct.Add("H02201", "Rau mầm mùi");
        //    dicProduct.Add("H02301", "Rau mầm mùi tây");
        //    dicProduct.Add("H02401", "Rau mầm muống");
        //    dicProduct.Add("H02501", "Rau mầm rau chua");
        //    dicProduct.Add("H02601", "Rau mầm rocket");
        //    dicProduct.Add("H02701", "Rau mầm súp lơ xanh");
        //    dicProduct.Add("H02801", "Rau mầm tía tô đỏ");
        //    dicProduct.Add("H02901", "Rau mầm tía tô xanh");
        //    dicProduct.Add("H03001", "Rau mầm cải củ đỏ");
        //    dicProduct.Add("H03101", "Rau mầm đậu tương");
        //    dicProduct.Add("H03201", "Rau mầm đậu xanh");
        //    dicProduct.Add("I00101", "Nấm bạch tuyết");
        //    dicProduct.Add("I00201", "Nấm bào ngư lớn");
        //    dicProduct.Add("I00301", "Nấm bào ngư nhỏ");
        //    dicProduct.Add("I00401", "Nấm cẩm thạch");
        //    dicProduct.Add("I00501", "Nấm chân dài");
        //    dicProduct.Add("I00601", "Nấm dạ dày");
        //    dicProduct.Add("I00701", "Nấm đậu xanh");
        //    dicProduct.Add("I00801", "Nấm đông cô tươi (nấm hương tươi)");
        //    dicProduct.Add("I00901", "Nấm đùi gà");
        //    dicProduct.Add("I01001", "Nấm hải sản");
        //    dicProduct.Add("I01101", "Nấm hoàng kim (nấm ngô)");
        //    dicProduct.Add("I01201", "Nấm sò sữa");
        //    dicProduct.Add("I01301", "Nấm khác");
        //    dicProduct.Add("I01401", "Nấm kim châm trắng");
        //    dicProduct.Add("I01501", "Nấm lim xanh");
        //    dicProduct.Add("I01601", "Nấm linh chi");
        //    dicProduct.Add("I01701", "Nấm mỡ");
        //    dicProduct.Add("I01801", "Nấm mộc nhĩ");
        //    dicProduct.Add("I01901", "Nấm nâu tây (sò nâu)");
        //    dicProduct.Add("I02001", "Nấm ngọc châm nâu");
        //    dicProduct.Add("I02101", "Nấm ngọc châm trắng");
        //    dicProduct.Add("I02201", "Nấm ngọc thạch");
        //    dicProduct.Add("I02301", "Nấm notaky");
        //    dicProduct.Add("I02401", "Nấm rơm");
        //    dicProduct.Add("I02501", "Nấm sò hương");
        //    dicProduct.Add("I02601", "Nấm sò nhật");
        //    dicProduct.Add("I02701", "Nấm bào ngư trắng");
        //    dicProduct.Add("I02801", "Nấm sò xám (nấm hào hương)");
        //    dicProduct.Add("I02901", "Nấm sò yến");
        //    dicProduct.Add("I03001", "Nấm kim châm vàng");
        //    dicProduct.Add("I03101", "Nấm trà tân");
        //    dicProduct.Add("I03201", "Nấm tú trân");
        //    dicProduct.Add("I03301", "Nấm tuyết");
        //    dicProduct.Add("I03401", "Nấm bào ngư xám");
        //    dicProduct.Add("I03501", "Nấm mỡ Yoshimoto trắng");
        //    dicProduct.Add("I03601", "Nấm mỡ Yoshimoto nâu");
        //    dicProduct.Add("I03701", "Nấm mỡ Jambo Yoshimoto trắng");
        //    dicProduct.Add("I03801", "Nấm mỡ Jambo Yoshimoto nâu");
        //    dicProduct.Add("J00101", "Lá chè xanh");
        //    dicProduct.Add("J00201", "Lá dong");
        //    dicProduct.Add("K00101", "Bơ sáp");
        //    dicProduct.Add("K00201", "Bưởi da xanh");
        //    dicProduct.Add("K00301", "Bưởi da xanh túi lưới");
        //    dicProduct.Add("K00401", "Bưởi diễn");
        //    dicProduct.Add("K00501", "Bưởi đoan hùng");
        //    dicProduct.Add("K00601", "Bưởi đường");
        //    dicProduct.Add("K00701", "Bưởi hồ lô (tết)");
        //    dicProduct.Add("K00801", "Bưởi năm roi");
        //    dicProduct.Add("K00901", "Bưởi năm roi túi lưới");
        //    dicProduct.Add("K01001", "Bưởi phúc trạch");
        //    dicProduct.Add("K01101", "Bưởi ruột đỏ");
        //    dicProduct.Add("K01201", "Cam canh");
        //    dicProduct.Add("K01301", "Cam hàm yên");
        //    dicProduct.Add("K01401", "Cam hòa bình");
        //    dicProduct.Add("K01501", "Cam mật");
        //    dicProduct.Add("K01601", "Cam sành");
        //    dicProduct.Add("K01701", "Cam vinh");
        //    dicProduct.Add("K01801", "Cam xoàn");
        //    dicProduct.Add("K01901", "Chanh có hạt");
        //    dicProduct.Add("K02001", "Chanh đào");
        //    dicProduct.Add("K02101", "Chanh dây");
        //    dicProduct.Add("K02201", "Chanh không hạt");
        //    dicProduct.Add("K02301", "Chanh tứ quý");
        //    dicProduct.Add("K02401", "Chanh vàng giống úc");
        //    dicProduct.Add("K02501", "Chôm chôm giống thái");
        //    dicProduct.Add("K02601", "Chôm chôm nhãn");
        //    dicProduct.Add("K02701", "Chôm chôm thường");
        //    dicProduct.Add("K02801", "Chuối cau");
        //    dicProduct.Add("K02901", "Chuối hột (chuối chát)");
        //    dicProduct.Add("K03001", "Chuối laba");
        //    dicProduct.Add("K03101", "Chuối ngự");
        //    dicProduct.Add("K03201", "Chuối sáp");
        //    dicProduct.Add("K03301", "Chuối tây (chuối sứ)");
        //    dicProduct.Add("K03401", "Bỏ vì gộp với Chuối tây (chuối sứ) - chuối tây");
        //    dicProduct.Add("K03501", "Chuối tiêu (chuối già)");
        //    dicProduct.Add("K03601", "Chuối tiêu xanh (tết)");
        //    dicProduct.Add("K03701", "Cóc xanh");
        //    dicProduct.Add("K03801", "Đào");
        //    dicProduct.Add("K03901", "Dâu tây đà lạt");
        //    dicProduct.Add("K04001", "Dâu tây giống nhật");
        //    dicProduct.Add("K04101", "Dâu tây new zealand");
        //    dicProduct.Add("K04201", "Dư (tết)");
        //    dicProduct.Add("K04301", "Đu đủ ruột đỏ");
        //    dicProduct.Add("K04401", "Đu đủ ruột vàng");
        //    dicProduct.Add("K04501", "Đu đủ vàng (tết)");
        //    dicProduct.Add("K04601", "Đu đủ xanh");
        //    dicProduct.Add("K04701", "Dừa (nguyên trái)");
        //    dicProduct.Add("K04801", "Dứa chín (thơm chín)");
        //    dicProduct.Add("K04901", "Dưa gang");
        //    dicProduct.Add("K05001", "Dừa gọt vỏ");
        //    dicProduct.Add("K05101", "Dưa hấu (tết)");
        //    dicProduct.Add("K05201", "Dưa hấu baby");
        //    dicProduct.Add("K05301", "Dưa hấu đỏ có hạt");
        //    dicProduct.Add("K05401", "Dưa hấu đỏ không hạt");
        //    dicProduct.Add("K05501", "Dưa hấu vàng có hạt");
        //    dicProduct.Add("K05601", "Dưa hấu vàng ít hạt");
        //    dicProduct.Add("K05701", "Dưa lê");
        //    dicProduct.Add("K05801", "Dưa lê geum sang");
        //    dicProduct.Add("K05901", "Dưa lê hoàng kim");
        //    dicProduct.Add("K06001", "Dưa lê kim cô nương");
        //    dicProduct.Add("K06101", "Dưa lê kim hoàng hậu");
        //    dicProduct.Add("K06201", "Dưa lê kim vương");
        //    dicProduct.Add("K06301", "Dưa lưới dài vỏ vàng");
        //    dicProduct.Add("K06401", "Dưa lưới dài vỏ xanh");
        //    dicProduct.Add("K06501", "Dưa lưới tròn vỏ vàng");
        //    dicProduct.Add("K06601", "Dưa lưới tròn vỏ xanh");
        //    dicProduct.Add("K06701", "Dưa thỏi vàng (tết)");
        //    dicProduct.Add("K06801", "Dứa xanh (thơm xanh)");
        //    dicProduct.Add("K06901", "Dừa xiêm (tết)");
        //    dicProduct.Add("K07001", "Gấc");
        //    dicProduct.Add("K07101", "Hồng giòn");
        //    dicProduct.Add("K07201", "Hồng xiêm (Sapochê)");
        //    dicProduct.Add("K07301", "Khế chua");
        //    dicProduct.Add("K07401", "Khế ngọt");
        //    dicProduct.Add("K07501", "Khóm đỏ (tết)");
        //    dicProduct.Add("K07601", "Khóm long phụng (tết)");
        //    dicProduct.Add("K07701", "Khóm xanh (tết)");
        //    dicProduct.Add("K07801", "Lê sapa");
        //    dicProduct.Add("K07901", "Mắc cọp");
        //    dicProduct.Add("K08001", "Mận cơm");
        //    dicProduct.Add("K08101", "Mận đỏ");
        //    dicProduct.Add("K08201", "Mận hậu");
        //    dicProduct.Add("K08301", "Mận thép lạng sơn");
        //    dicProduct.Add("K08401", "Mãng cầu (tết)");
        //    dicProduct.Add("K08501", "Mãng cầu xiêm");
        //    dicProduct.Add("K08601", "Măng cụt");
        //    dicProduct.Add("K08701", "Me chín giống Thái");
        //    dicProduct.Add("K08801", "Mít ta");
        //    dicProduct.Add("K08901", "Mít thái");
        //    dicProduct.Add("K09001", "Mít tố nữ");
        //    dicProduct.Add("K09101", "Na (mãng cầu ta)");
        //    dicProduct.Add("K09201", "Nhãn idor");
        //    dicProduct.Add("K09301", "Nhãn lồng");
        //    dicProduct.Add("K09401", "Nhãn quế");
        //    dicProduct.Add("K09501", "Nhãn xuồng");
        //    dicProduct.Add("K09601", "Nho đỏ Ninh Thuận");
        //    dicProduct.Add("K09701", "Nho xanh Ninh Thuận");
        //    dicProduct.Add("K09801", "Ổi đào");
        //    dicProduct.Add("K09901", "Ổi găng");
        //    dicProduct.Add("K10001", "Ổi lê");
        //    dicProduct.Add("K10101", "Ổi long khánh");
        //    dicProduct.Add("K10201", "Ổi nữ hoàng");
        //    dicProduct.Add("K10301", "Ổi thanh hà");
        //    dicProduct.Add("K10401", "Phật thủ (tết)");
        //    dicProduct.Add("K10501", "Phúc bồn tử");
        //    dicProduct.Add("K10601", "Quýt đường");
        //    dicProduct.Add("K10701", "Quýt hòa bình");
        //    dicProduct.Add("K10801", "Quýt tiều");
        //    dicProduct.Add("K10901", "Roi đỏ (mận) an phước");
        //    dicProduct.Add("K11001", "Roi trắng (mận sữa)");
        //    dicProduct.Add("K11101", "Roi xanh (mận xanh)");
        //    dicProduct.Add("K11201", "Sake");
        //    dicProduct.Add("K11301", "Sấu");
        //    dicProduct.Add("K11401", "Sầu riêng");
        //    dicProduct.Add("K11501", "Sơ ri");
        //    dicProduct.Add("K11601", "Sung (tết)");
        //    dicProduct.Add("K11701", "Táo xanh");
        //    dicProduct.Add("K11801", "Thanh long đỏ");
        //    dicProduct.Add("K11901", "Thanh long trắng");
        //    dicProduct.Add("K12001", "Vải thiều");
        //    dicProduct.Add("K12101", "Vú sữa tím");
        //    dicProduct.Add("K12201", "Vú sữa trắng");
        //    dicProduct.Add("K12301", "Xoài (tết)");
        //    dicProduct.Add("K12401", "Xoài cát bồ");
        //    dicProduct.Add("K12501", "Xoài cát chu");
        //    dicProduct.Add("K12601", "Xoài cát hòa lộc");
        //    dicProduct.Add("K12701", "Xoài giống đài loan");
        //    dicProduct.Add("K12801", "Xoài giống úc");
        //    dicProduct.Add("K12901", "Xoài keo");
        //    dicProduct.Add("K13001", "Xoài nha trang");
        //    dicProduct.Add("K13101", "Xoài tứ quý");
        //    dicProduct.Add("K13201", "Xoài xanh giòn");
        //    dicProduct.Add("K13301", "Xoài xanh ngọt");
        //    dicProduct.Add("K13401", "Xoài xiêm núm");
        //    dicProduct.Add("K13501", "Xoài yên châu");
        //    dicProduct.Add("K13601", "Ổi");
        //    dicProduct.Add("K13701", "Quýt Hồng");
        //    dicProduct.Add("K13801", "Sầu riêng Monthoong");
        //    dicProduct.Add("K13901", "Quất");
        //    dicProduct.Add("K14001", "Chuối dole");
        //    dicProduct.Add("K14101", "Chuối vàng");
        //    dicProduct.Add("K14201", "Dưa gang trái dài");
        //    dicProduct.Add("K14301", "Dưa gang trái tròn");
        //    dicProduct.Add("K14401", "Dưa hấu rằn");
        //    dicProduct.Add("K14501", "Ổi nữ hoàng baby");
        //    dicProduct.Add("K14601", "Ổi nữ hoàng ruột trắng");
        //    dicProduct.Add("K14701", "Ổi trân châu ruột đỏ");
        //    dicProduct.Add("K14801", "Ổi trân châu ruột trắng");
        //    dicProduct.Add("K14901", "Quýt đường túi lưới");
        //    dicProduct.Add("K15001", "Cóc non");
        //    dicProduct.Add("K15101", "Dưa lê hoàng cẩm");
        //    dicProduct.Add("K15201", "Cam lòng vàng");
        //    dicProduct.Add("K15301", "Nhãn tiêu da bò");
        //    dicProduct.Add("K15401", "Dưa lê kim đế vương");
        //    dicProduct.Add("K15501", "Dưa lê hoàng cẩm 800");
        //    dicProduct.Add("K15601", "Dưa lê hoàng cẩm 2000");
        //    dicProduct.Add("K15701", "Chôm chôm giống thái MB");
        //    dicProduct.Add("K15801", "Chôm chôm thường MB");
        //    dicProduct.Add("K15901", "Chôm chôm nhãn MB");
        //    dicProduct.Add("K16001", "Chuối đỏ");
        //    dicProduct.Add("K16101", "Mít ruột đỏ");
        //    dicProduct.Add("K16201", "Dưa lê Kim Thúy Mật");
        //    dicProduct.Add("M00101", "Cỏ voi");

        //    #endregion


        //    foreach (var ProductCode in dicProduct.Keys)
        //    {

        //        if (Product.Where(x => x.ProductCode == ProductCode).FirstOrDefault() == null)
        //        {
        //            Product _product = new Product();
        //            _product._id = Guid.NewGuid();
        //            _product.ProductId = _product._id;
        //            _product.ProductCode = ProductCode;
        //            _product.ProductName = dicProduct[ProductCode];

        //            Product.Add(_product);
        //        }
        //    }

        //    var mongoClient = new MongoClient();
        //    var db = mongoClient.GetDatabase("localtest");
        //    db.DropCollection("Product");

        //    var collection = db.GetCollection<Product>("Product");

        //    await collection.InsertManyAsync(Product);
        //}
        #endregion
    }

}