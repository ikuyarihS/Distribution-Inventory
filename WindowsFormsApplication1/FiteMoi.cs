#region

using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Windows.Forms;
using Aspose.Cells;
using MongoDB.Driver;

#endregion

namespace AllocatingStuff
{
    public partial class MainForm
    {
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
                Stopwatch stopWatch = Stopwatch.StartNew();

                #region Preparing!

                #region Initializing

                //string connectionString = System.Configuration.ConfigurationManager.ConnectionStrings["mongodb_vecrops.salesms"].ConnectionString;
                //MongoClient mongoClient = new MongoClient(connectionString);
                var mongoClient = new MongoClient();
                IMongoDatabase db = mongoClient.GetDatabase("localtest");
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

                List<PurchaseOrderDate> PO = db.GetCollection<PurchaseOrderDate>("PurchaseOrder")
                    .Find(x =>
                        x.DateOrder >= DateFrom.Date &&
                        x.DateOrder <= DateTo.Date)
                    .ToList();

                //var FC = db.GetCollection<ForecastDate>("Forecast").AsQueryable().ToList()
                //    .Where(x =>
                //        (x.DateForecast.Date >= DateFrom.Date) &&
                //        (x.DateForecast.Date <= DateTo.Date))
                //    .OrderBy(x => x.DateForecast);

                ForecastDate[] FC = db.GetCollection<ForecastDate>("Forecast").Find(x =>
                        x.DateForecast >= DateFrom.Date &&
                        x.DateForecast <= DateTo.Date)
                    .ToList()
                    .ToArray();

                List<Product> Product = db.GetCollection<Product>("Product").AsQueryable().ToList();
                List<Supplier> Supplier = db.GetCollection<Supplier>("Supplier").AsQueryable().ToList();
                List<Customer> Customer = db.GetCollection<Customer>("Customer").AsQueryable().ToList();

                coreStructure.dicProductUnit = db.GetCollection<ProductUnit>("ProductUnit")
                    .AsQueryable()
                    .ToDictionary(x => x.ProductCode);

                coreStructure.dicProductRate = db.GetCollection<ProductRate>("ProductRate")
                    .AsQueryable()
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

                foreach (Product _Product in Product)
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
                    .AsQueryable()
                    .ToDictionary(x => x.ProductId);

                #endregion

                #region Supplier

                foreach (Supplier _Supplier in Supplier)
                {
                    //_Supplier.SupplierName = Regex.Replace(_Supplier.SupplierName, @"[^\u0000-\u007F]+", string.Empty);
                    _Supplier.SupplierName = ConvertToUnsigned(_Supplier.SupplierName);
                    if (!coreStructure.dicSupplier.ContainsKey(_Supplier.SupplierId))
                        coreStructure.dicSupplier.Add(_Supplier.SupplierId, _Supplier);
                }

                #endregion Supplier

                #region Customer

                foreach (Customer customer in Customer)
                    if (!coreStructure.dicCustomer.TryGetValue(customer.CustomerId, out Customer _customer))
                        coreStructure.dicCustomer.Add(customer.CustomerId, customer);

                #endregion

                #region PO

                if (NoFruit) FruitOnly = false;

                // Everything related to PO.
                //int maxCalculation = 0;
                coreStructure.dicPO =
                    new Dictionary<DateTime, Dictionary<Product, Dictionary<CustomerOrder, bool>>>();
                var dicUnit = new Dictionary<string, string>(StringComparer.Ordinal);
                foreach (PurchaseOrderDate PODate in PO.OrderByDescending(x => x.DateOrder.Date).Reverse())
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

                    foreach (ProductOrder _ProductOrder in PODate.ListProductOrder.Reverse<ProductOrder>())
                    {
                        coreStructure.dicPO[PODate.DateOrder.Date]
                            .Add(
                                coreStructure.dicProduct[_ProductOrder.ProductId],
                                new Dictionary<CustomerOrder, bool>(_ProductOrder.ListCustomerOrder.Count() * 4));

                        foreach (CustomerOrder _CustomerOrder in _ProductOrder.ListCustomerOrder)
                        {
                            // Handling Unit.

                            // Round to the nearest 2nd decimal digit.
                            _CustomerOrder.QuantityOrder = Math.Round(_CustomerOrder.QuantityOrder, 2);

                            // Proper Type name.
                            string _OrderUnitType =
                                ProperUnit(_CustomerOrder.Unit.ToLower(), dicUnit); // Optimization Purposes.

                            // Converting to Kg.
                            // Only applicable to VM+. For now.
                            _CustomerOrder.Unit = _OrderUnitType;
                            if (coreStructure.dicCustomer[_CustomerOrder.CustomerId].CustomerType == "VM+" &&
                                _OrderUnitType != "Kg")
                            {
                                string _ProductCode = coreStructure.dicProduct[_ProductOrder.ProductId].ProductCode;
                                if (coreStructure.dicProductUnit.TryGetValue(_ProductCode,
                                    out ProductUnit _ProductUnit))
                                {
                                    ProductUnitRegion _ProductUnitRegion =
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
                                        coreStructure.dicProduct[_ProductOrder.ProductId]]
                                    .Add(_CustomerOrder, true);

                            //maxCalculation++;
                        }
                    }
                }

                #endregion

                #region FC

                coreStructure.dicFC =
                    new Dictionary<DateTime, Dictionary<Product, Dictionary<SupplierForecast, bool>>>(FC.Count() * 4);
                foreach (ForecastDate FCDate in FC.OrderBy(x => x.DateForecast.Date))
                {
                    coreStructure.dicFC.Add(FCDate.DateForecast.Date,
                        new Dictionary<Product, Dictionary<SupplierForecast, bool>>(Product.Count() * 4));
                    foreach (ProductForecast _ProductForecast in FCDate.ListProductForecast)
                    {
                        coreStructure.dicFC[FCDate.DateForecast.Date]
                            .Add(
                                coreStructure.dicProduct[_ProductForecast.ProductId],
                                new Dictionary<SupplierForecast, bool>(FCDate.ListProductForecast.Count() * 4));
                        // To allow user to store their plans on the Forecast
                        // Added a filter on 0 forecast.
                        foreach (SupplierForecast _SupplierForecast in _ProductForecast.ListSupplierForecast
                            .Where(x => x.QualityControlPass /*&& x.QuantityForecast > 0*/)
                            .OrderBy(x => x.Level))
                            coreStructure.dicFC[FCDate.DateForecast.Date][
                                    coreStructure.dicProduct[_ProductForecast.ProductId]]
                                .Add(_SupplierForecast, true);
                    }
                }

                #endregion

                #region Best of both worlds.

                coreStructure.dicCoord =
                    new Dictionary<DateTime, Dictionary<Product,
                        Dictionary<CustomerOrder, Dictionary<SupplierForecast, DateTime>>>>();
                foreach (PurchaseOrderDate PODate in PO)
                {
                    coreStructure.dicCoord.Add(PODate.DateOrder.Date,
                        new Dictionary<Product, Dictionary<CustomerOrder, Dictionary<SupplierForecast, DateTime>>>(
                            Product.Count() * 4));
                    foreach (ProductOrder _ProductOrder in PODate.ListProductOrder)
                    {
                        coreStructure.dicCoord[PODate.DateOrder.Date]
                            .Add(
                                coreStructure.dicProduct[_ProductOrder.ProductId],
                                new Dictionary<CustomerOrder, Dictionary<SupplierForecast, DateTime>>(
                                    PODate.ListProductOrder.Count() * 4));
                        foreach (CustomerOrder _CustomerOrder in _ProductOrder.ListCustomerOrder)
                        {
                            _CustomerOrder.QuantityOrder = Math.Round(_CustomerOrder.QuantityOrder, 1);
                            coreStructure.dicCoord[PODate.DateOrder.Date][
                                    coreStructure.dicProduct[_ProductOrder.ProductId]]
                                .Add(_CustomerOrder, null);
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
                        IEnumerable<SupplierForecast> _listSupplierForecast = coreStructure.dicFC[DateFC][_Product]
                            .Keys
                            .Where(x =>
                                coreStructure.dicSupplier[x.SupplierId].SupplierType == "VinEco");
                        if (_listSupplierForecast != null)
                            foreach (SupplierForecast _SupplierForecast in _listSupplierForecast.OrderBy(x =>
                                coreStructure.dicSupplier[x.SupplierId].SupplierName))
                            {
                                Supplier _Supplier = coreStructure.dicSupplier[_SupplierForecast.SupplierId];
                                string sKey = string.Format("{0}{1}", _Product.ProductCode, _Supplier.SupplierCode);

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
                        IEnumerable<SupplierForecast> _listSupplierForecast = coreStructure.dicFC[DateFC][_Product]
                            .Keys
                            .Where(x =>
                                coreStructure.dicSupplier[x.SupplierId].SupplierType != "VinEco");
                        if (_listSupplierForecast != null)
                            foreach (SupplierForecast _SupplierForecast in _listSupplierForecast.OrderBy(x =>
                                coreStructure.dicSupplier[x.SupplierId].SupplierName))
                            {
                                Supplier _Supplier = coreStructure.dicSupplier[_SupplierForecast.SupplierId];
                                string sKey = string.Format("{0}{1}", _Product.ProductCode, _Supplier.SupplierCode);

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
                foreach (DateTime DateFC in coreStructure.dicFC.Keys)
                {
                    // Then, by product.
                    coreStructure.dicDeli.Add(DateFC, new Dictionary<Product, Dictionary<SupplierForecast, double>>());
                    foreach (Product _Product in coreStructure.dicFC[DateFC].Keys)
                    {
                        // And finally, by Suppliers. This would be the fairest.
                        coreStructure.dicDeli[DateFC].Add(_Product, new Dictionary<SupplierForecast, double>());
                        foreach (SupplierForecast _SupplierForecast in coreStructure.dicFC[DateFC][_Product].Keys)
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
                    {"1", 0.1}, // wat the fuck
                    {"2", 0.1}, // wat the duck
                    {"9", 0.1} // I give up.
                };

                coreStructure.dicTransferDays = new Dictionary<string, byte>
                {
                    {"North-North", 1},
                    {"Highland-North", 3},
                    {"Highland-South", 0},
                    {"South-South", 0},
                    {"South-North", dayCrossRegion},
                    {"North-South", dayCrossRegion}
                };

                #endregion

                #region Main Body

                if (!YesNoLimit) UpperCap = -1;

                // In case of no uppercap, to prevent allocating EVERY FUCKING THING INTO ONE TYPE.
                if (UpperCap == -1)
                    foreach (Customer _Customer in coreStructure.dicCustomer.Values)
                        _Customer.CustomerType = _Customer.CustomerType.Replace("Priority", "").Trim();

                //byte LDtoMB = 3;
                //byte MBtoMB = 1;
                //byte MNtoMN = 0;
                //byte LDtoMN = 0;
                //byte MBtoMN = 3;
                //byte MNtoMB = 3;

                var ListRegion = new (string From, string To, byte DayForFrom, byte DayForLD, bool IsCrossRegion)[]
                {
                    ("Miền Bắc", "Miền Bắc", (byte) 1, (byte) 3, false),
                    ("Miền Nam", "Miền Nam", (byte) 0, (byte) 0, false),
                    ("Lâm Đồng", "Miền Bắc", (byte) 3, (byte) 3, false),
                    ("Lâm Đồng", "Miền Nam", (byte) 0, (byte) 0, false),
                    ("Miền Bắc", "Miền Nam", dayCrossRegion, (byte) 0, true),
                    ("Miền Nam", "Miền Bắc", dayCrossRegion, dayCrossRegion, true)
                };

                WriteToRichTextBoxOutput();
                WriteToRichTextBoxOutput("UpperCap = " + UpperCap);
                WriteToRichTextBoxOutput();

                // P&L Goes here - using KPI first
                if (!YesNoCompact & !YesNoOnlyFarm)
                {
                    CoordCaller(coreStructure, ListRegion, "VinEco", 1, "B2B", YesNoByUnit, false, true);
                    CoordCaller(coreStructure, ListRegion, "ThuMua", 1, "B2B", YesNoByUnit, false, true);
                }

                #region VM+ VinEco Priority

                if (!YesNoOnlyFarm)
                    CoordCaller(coreStructure, ListRegion, "VCM", 1, "VM+ VinEco Priority", YesNoByUnit);

                CoordCaller(coreStructure, ListRegion, "VinEco", UpperCap, "VM+ VinEco Priority", YesNoByUnit, false,
                    true);

                if (!YesNoOnlyFarm)
                {
                    CoordCaller(coreStructure, ListRegion, "ThuMua", UpperCap, "VM+ VinEco Priority", YesNoByUnit,
                        false, true);
                    CoordCaller(coreStructure, ListRegion, "ThuMua", UpperCap, "VM+ VinEco Priority", YesNoByUnit,
                        true);
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
                    CoordCaller(coreStructure, ListRegion, "ThuMua", UpperCap, "VM+ VinEco", YesNoByUnit, true);
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

                foreach (Customer VmpVinEco in coreStructure.dicCustomer.Values.Where(x =>
                    x.CustomerType == "VM+ VinEco"))
                foreach (Customer SiteToChange in coreStructure.dicCustomer.Values.Where(x =>
                    x.CustomerCode == VmpVinEco.CustomerCode))
                    SiteToChange.CustomerType = VmpVinEco.CustomerType;

                foreach (Customer _Customer in coreStructure.dicCustomer.Values)
                    if (_Customer.CustomerType == "B2B")
                        _Customer.CustomerType = _Customer.Company;
                    else
                        _Customer.CustomerType = _Customer.CustomerType.Replace("Priority", "").Trim();

                // Dealing with stubborn Procuring Forcasts.
                foreach (DateTime DateFC in coreStructure.dicFC.Keys.OrderBy(x => x.Date))
                foreach (Product _Product in coreStructure.dicFC[DateFC].Keys.OrderBy(x => x.ProductCode))
                {
                    IEnumerable<SupplierForecast> _ListSupplier = coreStructure.dicFC[DateFC][_Product]
                        .Keys.Where(x =>
                            coreStructure.dicSupplier[x.SupplierId].SupplierType != "VinEco");
                    if (_ListSupplier != null)
                        foreach (SupplierForecast _SupplierForecast in _ListSupplier.Reverse())
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

                    foreach (DateTime DatePO in coreStructure.dicCoord.Keys)
                    foreach (Product _Product in coreStructure.dicCoord[DatePO].Keys)
                    foreach (CustomerOrder _CustomerOrder in coreStructure.dicCoord[DatePO][_Product].Keys)
                        if (coreStructure.dicCoord[DatePO][_Product][_CustomerOrder] != null)
                        {
                            foreach (SupplierForecast _SupplierForecast in coreStructure.dicCoord[DatePO][_Product][
                                    _CustomerOrder]
                                .Keys.OrderBy(x => coreStructure.dicSupplier[x.SupplierId].SupplierType))
                            {
                                DataRow dr = null;

                                Customer _Customer = coreStructure.dicCustomer[_CustomerOrder.CustomerId];
                                Supplier _Supplier = coreStructure.dicSupplier[_SupplierForecast.SupplierId];

                                string sKey = DatePO.Date +
                                              _Product.ProductCode +
                                              _Customer.CustomerType +
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
                                dr["Ngày tiêu thụ"] = (int) (DatePO.Date - _dateBase).TotalDays + 2;
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

                                dr["VCM"] = (double) dr["VCM"] + _CustomerOrder.QuantityOrderKg;
                                if (_Supplier.SupplierType == "VinEco")
                                    dr["VE"] = (double) dr["VE"] + _SupplierForecast.QuantityForecast;
                                else if (_Supplier.SupplierType == "ThuMua")
                                    dr["TM"] = (double) dr["TM"] + _SupplierForecast.QuantityForecast;

                                if (newRow)
                                    dtMastah.Rows.Add(dr);
                            }
                        }
                        else
                        {
                            DataRow dr = null;

                            Customer _Customer = coreStructure.dicCustomer[_CustomerOrder.CustomerId];

                            string sKey = DatePO.Date +
                                          _Product.ProductCode +
                                          _Customer.CustomerType +
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
                            dr["Ngày tiêu thụ"] = (int) (DatePO.Date - _dateBase).TotalDays + 2;
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

                            dr["VCM"] = (double) dr["VCM"] + _CustomerOrder.QuantityOrderKg;


                            if (newRow)
                                dtMastah.Rows.Add(dr);
                        }

                    foreach (DataRow dr in dtMastah.Rows)
                        if ((double) dr["VCM"] > (double) dr["VE"] + (double) dr["TM"])
                            dr["NoSup"] = "Yes";

                    #endregion

                    #region LeftoverVinEco

                    var dtLeftoverVe = new DataTable {TableName = "NoCusVinEco"};

                    dtLeftoverVe.Columns.Add("Mã VinEco", typeof(string)).DefaultValue = "";
                    dtLeftoverVe.Columns.Add("Tên VinEco", typeof(string)).DefaultValue = "";
                    dtLeftoverVe.Columns.Add("Mã Farm", typeof(string)).DefaultValue = "";
                    dtLeftoverVe.Columns.Add("Tên Farm", typeof(string)).DefaultValue = "";
                    dtLeftoverVe.Columns.Add("Ngày thu hoạch", typeof(int));
                    dtLeftoverVe.Columns.Add("Vùng sản xuất", typeof(string)).DefaultValue = "";
                    dtLeftoverVe.Columns.Add("Sản lượng", typeof(double)).DefaultValue = 0;

                    foreach (DateTime DateFC in coreStructure.dicFC.Keys)
                    foreach (Product _Product in coreStructure.dicFC[DateFC].Keys.OrderBy(x => x.ProductCode))
                    {
                        IEnumerable<SupplierForecast> _ListSupplier = coreStructure.dicFC[DateFC][_Product]
                            .Keys.Where(
                                x =>
                                    coreStructure.dicSupplier[x.SupplierId].SupplierType == "VinEco");
                        if (_ListSupplier != null)
                            foreach (SupplierForecast _SupplierForecast in _ListSupplier.OrderBy(x =>
                                coreStructure.dicSupplier[x.SupplierId].SupplierName))
                                if (_SupplierForecast.QuantityForecast > 0)
                                {
                                    DataRow dr = dtLeftoverVe.NewRow();

                                    //Customer _Customer =coreStructure. dicCustomer[_CustomerOrder.CustomerId];
                                    Supplier _Supplier = coreStructure.dicSupplier[_SupplierForecast.SupplierId];

                                    dr["Mã VinEco"] = _Product.ProductCode;
                                    dr["Tên VinEco"] = _Product.ProductName;
                                    dr["Mã Farm"] = _Supplier.SupplierCode;
                                    dr["Tên Farm"] = _Supplier.SupplierName;
                                    dr["Ngày thu hoạch"] = (int) (DateFC.Date - _dateBase).TotalDays + 2;
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

                    foreach (DateTime DateFC in coreStructure.dicFC.Keys)
                    foreach (Product _Product in coreStructure.dicFC[DateFC].Keys.OrderBy(x => x.ProductCode))
                    {
                        IEnumerable<SupplierForecast> _ListSupplier = coreStructure.dicFC[DateFC][_Product]
                            .Keys.Where(
                                x =>
                                    coreStructure.dicSupplier[x.SupplierId].SupplierType == "ThuMua");
                        if (_ListSupplier != null)
                            foreach (SupplierForecast _SupplierForecast in _ListSupplier.OrderBy(x =>
                                coreStructure.dicSupplier[x.SupplierId].SupplierName))
                                if (_SupplierForecast.QuantityForecast > 0)
                                {
                                    DataRow dr = dtLeftoverTm.NewRow();

                                    //Customer _Customer =coreStructure. dicCustomer[_CustomerOrder.CustomerId];
                                    Supplier _Supplier = coreStructure.dicSupplier[_SupplierForecast.SupplierId];

                                    dr["Mã VinEco"] = _Product.ProductCode;
                                    dr["Tên VinEco"] = _Product.ProductName;
                                    dr["Mã Farm"] = _Supplier.SupplierCode;
                                    dr["Tên Farm"] = _Supplier.SupplierName;
                                    dr["Ngày thu hoạch"] = (int) (DateFC.Date - _dateBase).TotalDays + 2;
                                    dr["Vùng sản xuất"] = _Supplier.SupplierRegion;
                                    dr["Sản lượng"] = _SupplierForecast.QuantityForecast;

                                    dtLeftoverTm.Rows.Add(dr);
                                }
                    }

                    #endregion

                    #region Output to Excel

                    var dicDateCol = new Dictionary<string, int>();

                    dicDateCol.Add("Ngày tiêu thụ", dtMastah.Columns.IndexOf("Ngày tiêu thụ"));

                    string fileName = string.Format("Report M plus 1 {0}.xlsx",
                        DateFrom.ToString("dd.MM") +
                        " - " +
                        DateTo.ToString("dd.MM") +
                        " (" +
                        DateTime.Now.ToString("yyyyMMdd HH\\hmm") +
                        ")");
                    string path = string.Format(
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

                    foreach (DateTime DatePO in coreStructure.dicCoord.Keys)
                    foreach (Product _Product in coreStructure.dicCoord[DatePO].Keys)
                    foreach (CustomerOrder _CustomerOrder in coreStructure.dicCoord[DatePO][_Product].Keys)
                        if (coreStructure.dicCoord[DatePO][_Product][_CustomerOrder] != null)
                        {
                            foreach (SupplierForecast _SupplierForecast in coreStructure.dicCoord[DatePO][_Product][
                                    _CustomerOrder]
                                .Keys.OrderBy(x => coreStructure.dicSupplier[x.SupplierId].SupplierType))
                                if (_SupplierForecast.QuantityForecast < _CustomerOrder.QuantityOrderKg)
                                {
                                    DataRow dr = dtMastah.NewRow();

                                    Customer _Customer = coreStructure.dicCustomer[_CustomerOrder.CustomerId];
                                    Supplier _Supplier = coreStructure.dicSupplier[_SupplierForecast.SupplierId];

                                    dr["Mã VinEco"] = _Product.ProductCode;
                                    dr["Tên VinEco"] = _Product.ProductName;
                                    dr["Mã cửa hàng"] = _Customer.CustomerCode;
                                    dr["Loại cửa hàng"] = _Customer.CustomerType;
                                    dr["Ngày tiêu thụ"] = (int) (DatePO.Date - _dateBase).TotalDays + 2;
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
                            DataRow dr = dtMastah.NewRow();

                            Customer _Customer = coreStructure.dicCustomer[_CustomerOrder.CustomerId];

                            dr["Mã VinEco"] = _Product.ProductCode;
                            dr["Tên VinEco"] = _Product.ProductName;
                            dr["Mã cửa hàng"] = _Customer.CustomerCode;
                            dr["Loại cửa hàng"] = _Customer.CustomerType;
                            dr["Ngày tiêu thụ"] = (int) (DatePO.Date - _dateBase).TotalDays + 2;
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

                    var dtNoSup = new DataTable {TableName = "NoSup"};

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

                    foreach (DateTime DateFC in coreStructure.dicFC.Keys)
                    foreach (Product _Product in coreStructure.dicFC[DateFC].Keys.OrderBy(x => x.ProductCode))
                    {
                        IEnumerable<SupplierForecast> _ListSupplier = coreStructure.dicFC[DateFC][_Product]
                            .Keys.Where(
                                x =>
                                    coreStructure.dicSupplier[x.SupplierId].SupplierType == "VinEco");
                        if (_ListSupplier != null)
                            foreach (SupplierForecast _SupplierForecast in _ListSupplier.OrderBy(x =>
                                coreStructure.dicSupplier[x.SupplierId].SupplierName))
                                if (_SupplierForecast.QuantityForecast > 3)
                                {
                                    DataRow dr = dtLeftoverVe.NewRow();

                                    //Customer _Customer =coreStructure. dicCustomer[_CustomerOrder.CustomerId];
                                    Supplier _Supplier = coreStructure.dicSupplier[_SupplierForecast.SupplierId];

                                    dr["Mã VinEco"] = _Product.ProductCode;
                                    dr["Tên VinEco"] = _Product.ProductName;
                                    dr["Mã Farm"] = _Supplier.SupplierCode;
                                    dr["Tên Farm"] = _Supplier.SupplierName;
                                    dr["Ngày thu hoạch"] = (int) (DateFC.Date - _dateBase).TotalDays + 2;
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

                    foreach (DateTime DateFC in coreStructure.dicFC.Keys)
                    foreach (Product _Product in coreStructure.dicFC[DateFC].Keys.OrderBy(x => x.ProductCode))
                    {
                        IEnumerable<SupplierForecast> _ListSupplier = coreStructure.dicFC[DateFC][_Product]
                            .Keys.Where(
                                x =>
                                    coreStructure.dicSupplier[x.SupplierId].SupplierType == "ThuMua");
                        if (_ListSupplier != null)
                            foreach (SupplierForecast _SupplierForecast in _ListSupplier
                                .Where(x => x.QuantityForecastPlanned != null)
                                .OrderBy(x =>
                                    coreStructure.dicSupplier[x.SupplierId].SupplierName))
                                if (_SupplierForecast.QuantityForecast > 0)
                                {
                                    DataRow dr = dtLeftoverTm.NewRow();

                                    //Customer _Customer =coreStructure. dicCustomer[_CustomerOrder.CustomerId];
                                    Supplier _Supplier = coreStructure.dicSupplier[_SupplierForecast.SupplierId];

                                    dr["Mã VinEco"] = _Product.ProductCode;
                                    dr["Tên VinEco"] = _Product.ProductName;
                                    dr["Mã Farm"] = _Supplier.SupplierCode;
                                    dr["Tên Farm"] = _Supplier.SupplierName;
                                    dr["Ngày thu hoạch"] = (int) (DateFC.Date - _dateBase).TotalDays + 2;
                                    dr["Vùng sản xuất"] = _Supplier.SupplierRegion;
                                    dr["Sản lượng"] = _SupplierForecast.QuantityForecast;

                                    dtLeftoverTm.Rows.Add(dr);
                                }
                    }

                    #endregion

                    #region Output to Excel

                    var dicDateCol = new Dictionary<string, int> {{"Ngày tiêu thụ", dtMastah.Columns.IndexOf("Ngày tiêu thụ")}};

                    string fileName = $"NoSup {DateFrom:dd.MM} - {DateTo:dd.MM} ({DateTime.Now:yyyyMMdd HH\\hmm}).xlsx";
                    string path = $@"D:\Documents\Stuff\VinEco\Mastah Project\Test\{fileName}";

                    var listDt = new List<DataTable> {dtMastah, dtLeftoverVe, dtLeftoverTm};

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

                        var dtMastah = new DataTable {TableName = "Mastah"};

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

                        foreach (DateTime DatePO in coreStructure.dicCoord.Keys.OrderBy(x => x.Date)
                            .Where(x => x.Date >= DateFrom.AddDays(dayDistance).Date))
                        foreach (Product _Product in coreStructure.dicCoord[DatePO].Keys.OrderBy(x => x.ProductCode))
                        foreach (CustomerOrder _CustomerOrder in coreStructure.dicCoord[DatePO][_Product]
                            .Keys
                            .Where(x => x.QuantityOrderKg > 0)
                            .OrderBy(x => coreStructure.dicCustomer[x.CustomerId].CustomerType)
                            .ThenBy(x => coreStructure.dicCustomer[x.CustomerId].CustomerCode))
                            if (coreStructure.dicCoord[DatePO][_Product][_CustomerOrder] != null)
                            {
                                foreach (SupplierForecast _SupplierForecast in coreStructure.dicCoord[DatePO][_Product]
                                    [_CustomerOrder]
                                    .Keys.OrderBy(x =>
                                        coreStructure.dicSupplier[x.SupplierId].SupplierType))
                                {
                                    DataRow dr = dtMastah.NewRow();

                                    Customer _Customer = coreStructure.dicCustomer[_CustomerOrder.CustomerId];
                                    Supplier _Supplier = coreStructure.dicSupplier[_SupplierForecast.SupplierId];

                                    ProductUnitRegion _ProductUnitRegion = null;
                                    if (coreStructure.dicProductUnit.TryGetValue(_Product.ProductCode,
                                        out ProductUnit _ProductUnit))
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

                                    string _Region = string.Join(string.Empty,
                                        _Supplier.SupplierRegion.Where((ch, index) =>
                                            ch != ' ' && (index == 0 || _Supplier.SupplierRegion[index - 1] == ' ')));
                                    switch (_Supplier.SupplierType)
                                    {
                                        case "VinEco":
                                            dr["Tên VinEco " + _Region] = _Supplier.SupplierName;
                                            dr["Đáp ứng từ VinEco " + _Region] = _SupplierForecast.QuantityForecast;
                                            dr["Ngày sơ chế VinEco " + _Region] =
                                                coreStructure.dicCoord[DatePO][_Product][_CustomerOrder][
                                                        _SupplierForecast]
                                                    .Date;
                                            break;
                                        case "ThuMua":
                                            dr["Tên ThuMua " + _Region] = _Supplier.SupplierName;
                                            dr["Đáp ứng từ ThuMua " + _Region] = _SupplierForecast.QuantityForecast;
                                            dr["Ngày sơ chế ThuMua " + _Region] =
                                                coreStructure.dicCoord[DatePO][_Product][_CustomerOrder][
                                                        _SupplierForecast]
                                                    .Date;
                                            //dr["Giá mua ThuMua " + _Region] = 0;
                                            break;
                                        case "VCM":
                                            dr["Tên ThuMua " + _Region] = "VCM - " + _Supplier.SupplierName;
                                            dr["Đáp ứng từ ThuMua " + _Region] = _SupplierForecast.QuantityForecast;
                                            dr["Ngày sơ chế ThuMua " + _Region] =
                                                coreStructure.dicCoord[DatePO][_Product][_CustomerOrder][
                                                        _SupplierForecast]
                                                    .Date;
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

                                DataRow dr = dtMastah.NewRow();

                                Customer _Customer = coreStructure.dicCustomer[_CustomerOrder.CustomerId];
                                //Supplier _Supplier =coreStructure. dicSupplier[_SupplierForecast.SupplierId];

                                ProductUnitRegion _ProductUnitRegion = null;
                                if (coreStructure.dicProductUnit.TryGetValue(_Product.ProductCode,
                                    out ProductUnit _ProductUnit))
                                {
                                    _ProductUnitRegion = coreStructure.dicProductUnit[_Product.ProductCode]
                                        .ListRegion
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

                        var dtLeftOverVE = new DataTable {TableName = "DBSL dư"};

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

                        foreach (DateTime DateFC in coreStructure.dicFC.Keys)
                        foreach (Product _Product in coreStructure.dicFC[DateFC].Keys.OrderBy(x => x.ProductCode))
                        {
                            IEnumerable<SupplierForecast> _ListSupplier = coreStructure.dicFC[DateFC][_Product]
                                .Keys
                                .Where(x =>
                                    coreStructure.dicSupplier[x.SupplierId].SupplierType == "VinEco");
                            if (_ListSupplier != null)
                                foreach (SupplierForecast _SupplierForecast in _ListSupplier
                                    .Where(x => x.QuantityForecast >= 3)
                                    .OrderBy(x => coreStructure.dicSupplier[x.SupplierId].SupplierName))
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
                                    //dr["Nhu cầu Đã đáp ứng"] = String.Format("=SUM(M{0}, P{0}, S{0}, V{0}, Z{0}, AD{0})", dtLeftOverVE.Rows.Count + 6);

                                    string _Region = string.Join(string.Empty,
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

                        var dtCustomer = new DataTable {TableName = "Region I guess"};

                        dtCustomer.Columns.Add("Mã cửa hàng", typeof(string));
                        dtCustomer.Columns.Add("Vùng đặt hàng", typeof(string));
                        dtCustomer.Columns.Add("Loại cửa hàng", typeof(string));
                        dtCustomer.Columns.Add("Vùng tiêu thụ", typeof(string));
                        dtCustomer.Columns.Add("Tên cửa hàng", typeof(string));
                        dtCustomer.Columns.Add("P&L", typeof(string));

                        foreach (Customer _Customer in coreStructure.dicCustomer.Values)
                        {
                            DataRow dr = dtCustomer.NewRow();

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

                        string filePath = Application.StartupPath.Replace("\\bin\\Debug", "") + "\\Template\\{0}";
                        string fileFullPath = string.Format(filePath, "ChiaHang Mastah.xlsb");
                        string fileFullPath2007 = string.Format(filePath, "ChiaHang Mastah.xlsm");
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

                        object missing = Type.Missing;
                        string path = string.Format(
                            @"D:\Documents\Stuff\VinEco\Mastah Project\Test\ChiaHang Mastah {1}{0}.xlsb",
                            DateFrom.AddDays(dayDistance).ToString("dd.MM") +
                            " - " +
                            DateTo.AddDays(-dayDistance).ToString("dd.MM") +
                            " (" +
                            DateTime.Now.ToString("yyyyMMdd HH\\hmm") +
                            ")", FruitOnly ? "Fruit " : "");
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
                        Worksheet xlWsMastah = xlWb.Worksheets["Mastah"];

                        var dicTable = new Dictionary<DataTable, int>();

                        dicTable.Add(dtMastah, 6);
                        dicTable.Add(dtLeftOverVE, 6);

                        foreach (DataTable _dt in dicTable.Keys)
                        {
                            Worksheet _xlWs = xlWb.Worksheets[_dt.TableName];
                            var _colIndex = 0;
                            int _rowFirst = dicTable[_dt];

                            foreach (DataColumn dc in _dt.Columns)
                            {
                                if (dc.DataType == typeof(string))
                                {
                                    DataRow _dr = _dt.Rows[1];
                                    if (_dr[dc].ToString().Length > 0 && _dr[dc].ToString().Substring(0, 1) == "=")
                                    {
                                        int _rowIndex = _rowFirst - 1;
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
                        xlWsMastah.Cells[3, 0].Formula = $"=SUBTOTAL(3,A6:A{dtMastah.Rows.Count + 5})"; // A4
                        xlWsMastah.Cells[3, 4].Formula = $"=SUBTOTAL(3,E6:E{dtMastah.Rows.Count + 5})"; // E4
                        xlWsMastah.Cells[3, 8].Formula = $"=SUBTOTAL(3,I6:I{dtMastah.Rows.Count + 5})"; // I4
                        xlWsMastah.Cells[3, 9].Formula = $"=SUBTOTAL(9,J6:J{dtMastah.Rows.Count + 5})"; // J4
                        xlWsMastah.Cells[3, 10].Formula = $"=SUBTOTAL(9,K6:K{dtMastah.Rows.Count + 5})"; // K4
                        xlWsMastah.Cells[3, 11].Formula = $"=SUBTOTAL(3,L6:L{dtMastah.Rows.Count + 5})"; // L4
                        xlWsMastah.Cells[3, 12].Formula = $"=SUBTOTAL(9,M6:M{dtMastah.Rows.Count + 5})"; // M4
                        xlWsMastah.Cells[3, 14].Formula = $"=SUBTOTAL(3,O6:O{dtMastah.Rows.Count + 5})"; // O4
                        xlWsMastah.Cells[3, 15].Formula = $"=SUBTOTAL(9,P6:P{dtMastah.Rows.Count + 5})"; // P4
                        xlWsMastah.Cells[3, 17].Formula = $"=SUBTOTAL(3,R6:R{dtMastah.Rows.Count + 5})"; // R4
                        xlWsMastah.Cells[3, 18].Formula = $"=SUBTOTAL(9,S6:S{dtMastah.Rows.Count + 5})"; // S4
                        xlWsMastah.Cells[3, 20].Formula = $"=SUBTOTAL(3,U6:U{dtMastah.Rows.Count + 5})"; // U4
                        xlWsMastah.Cells[3, 21].Formula = $"=SUBTOTAL(9,V6:V{dtMastah.Rows.Count + 5})"; // V4
                        xlWsMastah.Cells[3, 23].Formula = $"=SUBTOTAL(3,X6:X{dtMastah.Rows.Count + 5})"; // Y4
                        xlWsMastah.Cells[3, 24].Formula = $"=SUBTOTAL(9,Y6:Y{dtMastah.Rows.Count + 5})"; // Z4
                        xlWsMastah.Cells[3, 26].Formula = $"=SUBTOTAL(3,AA6:AA{dtMastah.Rows.Count + 5})"; // AC4
                        xlWsMastah.Cells[3, 27].Formula = $"=SUBTOTAL(9,AB6:AB{dtMastah.Rows.Count + 5})"; // AD4

                        // Formula Stuff for Leftover VE
                        xlWb.Worksheets["DBSL Dư"].Cells[3, 0].Formula = $"=SUBTOTAL(3,A6:A{dtLeftOverVE.Rows.Count + 5})"; // A4
                        xlWb.Worksheets["DBSL Dư"].Cells[3, 4].Formula = $"=SUBTOTAL(3,E6:E{dtLeftOverVE.Rows.Count + 5})"; // E4
                        xlWb.Worksheets["DBSL Dư"].Cells[3, 8].Formula = $"=SUBTOTAL(3,I6:I{dtLeftOverVE.Rows.Count + 5})"; // I4
                        xlWb.Worksheets["DBSL Dư"].Cells[3, 9].Formula = $"=SUBTOTAL(9,J6:J{dtLeftOverVE.Rows.Count + 5})"; // J4
                        xlWb.Worksheets["DBSL Dư"].Cells[3, 10].Formula = $"=SUBTOTAL(9,K6:K{dtLeftOverVE.Rows.Count + 5})"; // K4
                        xlWb.Worksheets["DBSL Dư"].Cells[3, 11].Formula = $"=SUBTOTAL(3,L6:L{dtLeftOverVE.Rows.Count + 5})"; // L4
                        xlWb.Worksheets["DBSL Dư"].Cells[3, 12].Formula = $"=SUBTOTAL(9,M6:M{dtLeftOverVE.Rows.Count + 5})"; // M4
                        xlWb.Worksheets["DBSL Dư"].Cells[3, 14].Formula = $"=SUBTOTAL(3,O6:O{dtLeftOverVE.Rows.Count + 5})"; // O4
                        xlWb.Worksheets["DBSL Dư"].Cells[3, 15].Formula = $"=SUBTOTAL(9,P6:P{dtLeftOverVE.Rows.Count + 5})"; // P4
                        xlWb.Worksheets["DBSL Dư"].Cells[3, 17].Formula = $"=SUBTOTAL(3,R6:R{dtLeftOverVE.Rows.Count + 5})"; // R4
                        xlWb.Worksheets["DBSL Dư"].Cells[3, 18].Formula = $"=SUBTOTAL(9,S6:S{dtLeftOverVE.Rows.Count + 5})"; // S4
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
                        var dtMastah = new DataTable {TableName = "Mastah"};

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
                            foreach (DateTime DatePO in coreStructure.dicCoord.Keys)
                            foreach (Product _Product in coreStructure.dicCoord[DatePO]
                                .Keys
                                .OrderBy(x => x.ProductCode))
                            foreach (CustomerOrder _CustomerOrder in coreStructure.dicCoord[DatePO][_Product]
                                .Keys
                                .OrderBy(x => coreStructure.dicCustomer[x.CustomerId].CustomerType))
                            {
                                Customer _Customer = coreStructure.dicCustomer[_CustomerOrder.CustomerId];
                                string sKey = $"{DatePO.Date:yyyyMMdd}{_Customer.CustomerType}{_Customer.CustomerBigRegion}{_Product.ProductCode}";
                                if (coreStructure.dicCoord[DatePO][_Product][_CustomerOrder] != null)
                                {
                                    foreach (SupplierForecast _SupplierForecast in coreStructure.dicCoord[DatePO][
                                            _Product]
                                        [_CustomerOrder]
                                        .Keys.OrderBy(x =>
                                            coreStructure.dicSupplier[x.SupplierId].SupplierType))
                                    {
                                        Supplier _Supplier = coreStructure.dicSupplier[_SupplierForecast.SupplierId];

                                        sKey += _Supplier.SupplierType;

                                        DataRow dr = null;
                                        if (!dicRow.TryGetValue(sKey, out int _rowIndex))
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

                                        string _Region = string.Join(string.Empty,
                                            _Supplier.SupplierRegion.Where((ch, index) =>
                                                ch != ' ' &&
                                                (index == 0 || _Supplier.SupplierRegion[index - 1] == ' ')));

                                        dr["Mã 6 ký tự"] = _Product.ProductCode;
                                        dr["Tên sản phẩm"] = _Product.ProductName;
                                        dr["Loại cửa hàng"] = _Customer.CustomerType;
                                        //dr["Ngày tiêu thụ"] = (int)(DatePO.Date - _dateBase).TotalDays + 2;
                                        dr["Ngày tiêu thụ"] = DatePO.Date;
                                        dr["Vùng tiêu thụ"] = _Customer.CustomerBigRegion;
                                        dr["Nhu cầu VinCommerce"] =
                                            Convert.ToDouble(dr["Nhu cầu VinCommerce"]) + _CustomerOrder.QuantityOrder;
                                        dr["Nhu cầu Đáp ứng"] =
                                            Convert.ToDouble(dr["Nhu cầu Đáp ứng"]) + _SupplierForecast.QuantityForecast;
                                        ;
                                        dr["Nguồn"] = _Supplier.SupplierType;
                                        dr["Vùng sản xuất"] = _Supplier.SupplierRegion;
                                        dr["Tên NCC"] = _Supplier.SupplierType == "ThuMua"
                                            ? "ThuMua"
                                            : _Supplier.SupplierName;
                                        //dr["Ngày sơ chế"] = (int)(coreStructure.dicCoord[DatePO][_Product][_CustomerOrder][_SupplierForecast].Date - _dateBase).TotalDays + 2;
                                        dr["Ngày sơ chế"] = coreStructure.dicCoord[DatePO][_Product][_CustomerOrder][_SupplierForecast].Date;
                                    }
                                }
                                else
                                {
                                    DataRow dr = null;
                                    if (!dicRow.TryGetValue(sKey, out int _rowIndex))
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

                            foreach (DateTime DatePO in coreStructure.dicCoord.Keys.OrderBy(x => x.Date)
                                .Where(x => x.Date >= DateFrom.AddDays(dayDistance).Date))
                            foreach (Product _Product in coreStructure.dicCoord[DatePO]
                                .Keys
                                .OrderBy(x => x.ProductCode))
                            foreach (CustomerOrder _CustomerOrder in coreStructure.dicCoord[DatePO][_Product]
                                .Keys
                                .Where(x => x.QuantityOrderKg > 0)
                                .OrderBy(x => coreStructure.dicCustomer[x.CustomerId].CustomerType)
                                .ThenBy(x => coreStructure.dicCustomer[x.CustomerId].CustomerCode))
                            {
                                Customer _Customer = coreStructure.dicCustomer[_CustomerOrder.CustomerId];
                                string sKey = 
                                    $"{DatePO.Date:yyyyMMdd}{_Customer.CustomerType}{_Customer.Company}{_Customer.CustomerBigRegion}{_Product.ProductCode}{(YesNoSubRegion ? _Customer.CustomerRegion : null)}";
                                if (coreStructure.dicCoord[DatePO][_Product][_CustomerOrder] != null)
                                {
                                    foreach (SupplierForecast _SupplierForecast in coreStructure.dicCoord[DatePO][
                                            _Product]
                                        [_CustomerOrder]
                                        .Keys.OrderBy(x =>
                                            coreStructure.dicSupplier[x.SupplierId].SupplierType))
                                    {
                                        Supplier _Supplier = coreStructure.dicSupplier[_SupplierForecast.SupplierId];

                                        sKey += _Supplier.SupplierCode;

                                        DataRow dr = null;
                                        if (!dicRow.TryGetValue(sKey, out int _rowIndex))
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

                                        string _Region = string.Join(string.Empty,
                                            _Supplier.SupplierRegion.Split(' ').Select(x => x.First()));

                                        //string _Region = string.Join(String.Empty, _Supplier.SupplierRegion.Where((ch, index) => ch != ' ' && (index == 0 || _Supplier.SupplierRegion[index - 1] == ' ')));
                                        string _colName = $"{_Supplier.SupplierType} {_Region}";

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
                                        dr["Nhu cầu"] = (double) dr["Nhu cầu"] + _CustomerOrder.QuantityOrderKg;
                                        dr["Đáp ứng"] = (double) dr["Đáp ứng"] + _SupplierForecast.QuantityForecast;
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
                                        dr["CodeSFG"] = $"{_Product.ProductCode}{1}{(_Supplier.SupplierRegion == "Lâm Đồng" ? 0 : 2) + (_SupplierForecast.LabelVinEco ? 1 : 0)}";
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
                                dr["NoSup"] = Math.Max((double) dr["Nhu cầu"] - (double) dr["Đáp ứng"], 0);
                                if ((double) dr["NoSup"] > 1) dr["IsNoSup"] = true;
                            }

                            IOrderedEnumerable<ForecastDate> _FC = db.GetCollection<ForecastDate>("Forecast")
                                .Find(x =>
                                    x.DateForecast >= DateFrom.Date &&
                                    x.DateForecast <= DateTo.Date)
                                .ToList()
                                .OrderByDescending(x => x.DateForecast);

                            foreach (ForecastDate _ForecastDate in _FC)
                            foreach (ProductForecast _ProductForecast in _ForecastDate.ListProductForecast)
                            foreach (SupplierForecast _SupplierForecast in _ProductForecast.ListSupplierForecast.Where(
                                x =>
                                    x.QualityControlPass && x.QuantityForecastPlanned > 0))
                            {
                                DataRow dr = dtMastah.NewRow();

                                Product _Product = coreStructure.dicProduct[_ProductForecast.ProductId];
                                Supplier _Supplier = coreStructure.dicSupplier[_SupplierForecast.SupplierId];

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

                        foreach (DateTime DateFC in coreStructure.dicFC.Keys)
                        foreach (Product _Product in coreStructure.dicFC[DateFC].Keys.OrderBy(x => x.ProductCode))
                        {
                            IEnumerable<SupplierForecast> _ListSupplier = coreStructure.dicFC[DateFC][_Product]
                                .Keys
                                .Where(x =>
                                    coreStructure.dicSupplier[x.SupplierId].SupplierType == "VinEco");
                            if (_ListSupplier != null)
                                foreach (SupplierForecast _SupplierForecast in _ListSupplier
                                    .Where(x => x.QuantityForecast > 3)
                                    .OrderBy(x => coreStructure.dicSupplier[x.SupplierId].SupplierName))
                                    if (_SupplierForecast.QuantityForecast > 0)
                                    {
                                        DataRow dr = dtLeftoverVe.NewRow();

                                        //Customer _Customer =coreStructure. dicCustomer[_CustomerOrder.CustomerId];
                                        Supplier _Supplier = coreStructure.dicSupplier[_SupplierForecast.SupplierId];

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

                        foreach (DateTime DateFC in coreStructure.dicFC.Keys)
                        foreach (Product _Product in coreStructure.dicFC[DateFC].Keys.OrderBy(x => x.ProductCode))
                        {
                            IEnumerable<SupplierForecast> _ListSupplier = coreStructure.dicFC[DateFC][_Product]
                                .Keys
                                .Where(x =>
                                    coreStructure.dicSupplier[x.SupplierId].SupplierType != "VinEco");
                            if (_ListSupplier != null)
                                foreach (SupplierForecast _SupplierForecast in _ListSupplier
                                    .Where(x => x.QuantityForecast >= 3)
                                    .OrderBy(x => coreStructure.dicSupplier[x.SupplierId].SupplierName))
                                    if (_SupplierForecast.QuantityForecast > 0)
                                    {
                                        DataRow dr = dtLeftoverTmKPI.NewRow();

                                        //Customer _Customer =coreStructure. dicCustomer[_CustomerOrder.CustomerId];
                                        Supplier _Supplier = coreStructure.dicSupplier[_SupplierForecast.SupplierId];

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

                        foreach (DateTime DateFC in coreStructure.dicFC.Keys)
                        foreach (Product _Product in coreStructure.dicFC[DateFC].Keys.OrderBy(x => x.ProductCode))
                        {
                            IEnumerable<SupplierForecast> _ListSupplier = coreStructure.dicFC[DateFC][_Product]
                                .Keys
                                .Where(x =>
                                    coreStructure.dicSupplier[x.SupplierId].SupplierType != "VinEco");
                            if (_ListSupplier != null)
                                foreach (SupplierForecast _SupplierForecast in _ListSupplier
                                    .Where(x => x.QuantityForecastOriginal >= 3)
                                    .OrderBy(x =>
                                        coreStructure.dicSupplier[x.SupplierId].SupplierName))
                                    if (_SupplierForecast.QuantityForecast > 0)
                                    {
                                        DataRow dr = dtLeftoverTm.NewRow();

                                        //Customer _Customer =coreStructure. dicCustomer[_CustomerOrder.CustomerId];
                                        Supplier _Supplier = coreStructure.dicSupplier[_SupplierForecast.SupplierId];

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

                        string fileName =
                            $"Mastah Compact {UpperCap:P0} {(FruitOnly ? "Fruit " : "")}{DateFrom.AddDays(dayDistance):dd.MM} - {DateTo.AddDays(-dayDistance):dd.MM} ({DateTime.Now:yyyyMMdd HH\\hmm}).xlsb";
                        string path = string.Format(
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

                        var dtMastah = new DataTable {TableName = "Mastah"};


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
                        foreach (Product _Product in coreStructure.dicCoord[DatePO].Keys.OrderBy(x => x.ProductCode))
                        foreach (CustomerOrder _CustomerOrder in coreStructure.dicCoord[DatePO][_Product]
                            .Keys
                            .OrderBy(x => coreStructure.dicCustomer[x.CustomerId].CustomerType))
                        {
                            Customer _Customer = coreStructure.dicCustomer[_CustomerOrder.CustomerId];
                            string sKey = $"{DatePO.Date}{_Customer.CustomerType}{_Customer.CustomerBigRegion}{_Product.ProductCode}";
                            if (coreStructure.dicCoord[DatePO][_Product][_CustomerOrder] != null)
                            {
                                foreach (SupplierForecast _SupplierForecast in coreStructure.dicCoord[DatePO][_Product]
                                    [_CustomerOrder]
                                    .Keys.OrderBy(x =>
                                        coreStructure.dicSupplier[x.SupplierId].SupplierType))
                                {
                                    Supplier _Supplier = coreStructure.dicSupplier[_SupplierForecast.SupplierId];

                                    DataRow dr = null;
                                    if (!dicRow.TryGetValue(sKey, out int _rowIndex))
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

                                    string _Region = string.Join(string.Empty,
                                        _Supplier.SupplierRegion.Where((ch, index) =>
                                            ch != ' ' && (index == 0 || _Supplier.SupplierRegion[index - 1] == ' ')));
                                    string _colName = string.Format("{0} {1}", _Supplier.SupplierType, _Region);

                                    dr["Mã 6 ký tự"] = _Product.ProductCode;
                                    dr["Tên sản phẩm"] = _Product.ProductName;
                                    dr["Loại cửa hàng"] = _Customer.CustomerType;
                                    dr["Ngày tiêu thụ"] = (int) (DatePO.Date - _dateBase).TotalDays + 2;
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
                                dr["Ngày tiêu thụ"] = (int) (DatePO.Date - _dateBase).TotalDays + 2;
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

                        foreach (DateTime DateFC in coreStructure.dicFC.Keys)
                        foreach (Product _Product in coreStructure.dicFC[DateFC].Keys.OrderBy(x => x.ProductCode))
                        {
                            IEnumerable<SupplierForecast> _ListSupplier = coreStructure.dicFC[DateFC][_Product]
                                .Keys
                                .Where(x =>
                                    coreStructure.dicSupplier[x.SupplierId].SupplierType == "VinEco");
                            if (_ListSupplier != null)
                                foreach (SupplierForecast _SupplierForecast in _ListSupplier.OrderBy(x =>
                                    coreStructure.dicSupplier[x.SupplierId].SupplierName))
                                    if (_SupplierForecast.QuantityForecast > 0)
                                    {
                                        DataRow dr = dtLeftoverVe.NewRow();

                                        //Customer _Customer =coreStructure. dicCustomer[_CustomerOrder.CustomerId];
                                        Supplier _Supplier = coreStructure.dicSupplier[_SupplierForecast.SupplierId];

                                        dr["Mã VinEco"] = _Product.ProductCode;
                                        dr["Tên VinEco"] = _Product.ProductName;
                                        dr["Mã Farm"] = _Supplier.SupplierCode;
                                        dr["Tên Farm"] = _Supplier.SupplierName;
                                        dr["Ngày thu hoạch"] = (int) (DateFC.Date - _dateBase).TotalDays + 2;
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

                        foreach (DateTime DateFC in coreStructure.dicFC.Keys)
                        foreach (Product _Product in coreStructure.dicFC[DateFC].Keys.OrderBy(x => x.ProductCode))
                        {
                            IEnumerable<SupplierForecast> _ListSupplier = coreStructure.dicFC[DateFC][_Product]
                                .Keys
                                .Where(x =>
                                    coreStructure.dicSupplier[x.SupplierId].SupplierType == "ThuMua");
                            if (_ListSupplier != null)
                                foreach (SupplierForecast _SupplierForecast in _ListSupplier.OrderBy(x =>
                                    coreStructure.dicSupplier[x.SupplierId].SupplierName))
                                    if (_SupplierForecast.QuantityForecast > 0)
                                    {
                                        DataRow dr = dtLeftoverTm.NewRow();

                                        //Customer _Customer =coreStructure. dicCustomer[_CustomerOrder.CustomerId];
                                        Supplier _Supplier = coreStructure.dicSupplier[_SupplierForecast.SupplierId];

                                        dr["Mã VinEco"] = _Product.ProductCode;
                                        dr["Tên VinEco"] = _Product.ProductName;
                                        dr["Mã Farm"] = _Supplier.SupplierCode;
                                        dr["Tên Farm"] = _Supplier.SupplierName;
                                        dr["Ngày thu hoạch"] = (int) (DateFC.Date - _dateBase).TotalDays + 2;
                                        dr["Vùng sản xuất"] = _Supplier.SupplierRegion;
                                        dr["Sản lượng"] = _SupplierForecast.QuantityForecast;

                                        dtLeftoverTm.Rows.Add(dr);
                                    }
                        }

                        #endregion

                        #region Output to Excel - OpenXMLWriter Style, super fast.

                        string fileName =
                            $"Mastah Compact {UpperCap:P0} {(FruitOnly ? "Fruit " : "")}{DateFrom.AddDays(dayDistance).ToString("dd.MM") + " - " + DateTo.AddDays(-dayDistance).ToString("dd.MM") + " (" + DateTime.Now.ToString("yyyyMMdd HH\\hmm") + ")"}.xlsb";
                        string path = string.Format(
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
    }
}