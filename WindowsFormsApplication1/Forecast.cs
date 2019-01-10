// --------------------------------------------------------------------------------------------------------------------
// <copyright file="Forecast.cs" company="VinEco">
//   Shirayuki 2018.
// </copyright>
// <summary>
//   Defines the MainForm type.
// </summary>
// --------------------------------------------------------------------------------------------------------------------

using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using Aspose.Cells;
using MongoDB.Driver;

namespace AllocatingStuff
{
    #region

    #endregion

    /// <summary>
    ///     The main form.
    /// </summary>
    public partial class MainForm
    {
        /// <summary>
        ///     Do naughty stuff with FC
        /// </summary>
        /// <param name="forecasts"> The FC. </param>
        /// <param name="worksheet"> The xl Ws. </param>
        /// <param name="supplierType"> The Supplier Type. </param>
        /// <param name="dicFc"> The dic FC. </param>
        /// <param name="dicProducts"> The dic Product. </param>
        /// <param name="dicSuppliers"> The dic Supplier. </param>
        /// <param name="products"> The Product. </param>
        /// <param name="suppliers"> The Supplier. </param>
        /// <param name="yesNoKpi"> The Yes No KPI. </param>
        private void EatForecastAspose(
            ICollection<ForecastDate>                                           forecasts,
            Worksheet                                                           worksheet,
            string                                                              supplierType,
            IDictionary<DateTime, Dictionary<string, Dictionary<string, Guid>>> dicFc,
            Dictionary<string, Product>                                         dicProducts,
            Dictionary<string, Supplier>                                        dicSuppliers,
            ICollection<Product>                                                products,
            ICollection<Supplier>                                               suppliers,
            bool                                                                yesNoKpi = false)
        {
            try
            {
                // int rowIndex = 0;
                // if (xlRng.Cells[1, 1].value != "Region" & xlRng.Cells[1, 1].value != "Vùng")
                // {
                // do
                // {
                // rowIndex++;
                // if (rowIndex >= xlRng.Rows.Count) { return; }
                // } while (xlRng.Cells[rowIndex + 1, 1].Value != "Region" & xlRng.Cells[rowIndex + 1, 1].Value != "Vùng");
                // }

                // DataTable dt = new DataTable();

                // OleDbConnection oleCon = new OleDbConnection(conStr);

                // OleDbDataAdapter _oleAdapt = new OleDbDataAdapter("Select * From [" + xlWs.Name.ToString() + "$" + xlRng.Offset[rowIndex, 0].Address[RowAbsolute: false, ColumnAbsolute: false, ReferenceStyle: Excel.XlReferenceStyle.xlA1, External: xlRng] + "]", oleCon);
                // string _str = xlRng.Offset[rowIndex, 0].Address as string;
                // WriteToRichTextBoxOutput(_str);
                // _oleAdapt.Fill(dt);

                // oleCon.Close();

                // Find first row.
                var rowIndex = 0;
                do
                {
                    if (worksheet.Cells[rowIndex, 0].Value == null || worksheet.Cells[rowIndex, 0].Value.ToString() != "Vùng" && worksheet.Cells[rowIndex, 0].Value.ToString() != "Region")
                    {
                        rowIndex++;
                    }
                } while (rowIndex <= worksheet.Cells.MaxDataRow + 1 && worksheet.Cells[rowIndex, 0].Value == null || worksheet.Cells[rowIndex, 0].Value.ToString() != "Vùng" && worksheet.Cells[rowIndex, 0].Value.ToString() != "Region");

                if (rowIndex > worksheet.Cells.MaxDataRow + 1)
                {
                    for (var i = 0; i < 7; i++)
                    {
                        WriteToRichTextBoxOutput("Wrong Format.");
                    }

                    return;
                }

                // ... ah well, option based 0.
                // rowIndex--;

                // Import into a DataTable.
                var opts = new ExportTableOptions
                               {
                                   CheckMixedValueType = true,
                                   ExportAsString      = false,
                                   FormatStrategy      = CellValueFormatStrategy.None,
                                   ExportColumnName    = true
                               };

                var dt = new DataTable { TableName = worksheet.Name };
                dt     = worksheet.Cells.ExportDataTable(
                    rowIndex,
                    0,
                    worksheet.Cells.MaxDataRow    + 1,
                    worksheet.Cells.MaxDataColumn + 1,
                    opts);

                // To deal with the uhm, Templates having different Headers.
                // Please shoot me.
                // if (dt.Columns.Contains("Vùng")) dt.Columns["Vùng"].ColumnName = "Region";
                // if (dt.Columns.Contains("Mã Farm")) dt.Columns["Mã Farm"].ColumnName = "SCODE";
                // if (dt.Columns.Contains("Tên Farm")) dt.Columns["Tên Farm"].ColumnName = "SNAME";
                // if (dt.Columns.Contains("Nhóm")) dt.Columns["Nhóm"].ColumnName = "PCLASS";
                // if (dt.Columns.Contains("Mã VECrops")) dt.Columns["Mã VECrops"].ColumnName = "VECrops Code";
                // if (dt.Columns.Contains("Mã VinEco")) dt.Columns["Mã VinEco"].ColumnName = "PCODE";
                // if (dt.Columns.Contains("Tên VinEco")) dt.Columns["Tên VinEco"].ColumnName = "PNAME";

                // To deal with the uhm, Templates having different Headers.
                // Please shoot me.
                // ReSharper disable once SuggestVarOrType_SimpleTypes
                foreach (var key in new (string oldName, string newName)[]
                                        {
                                            ("Vùng", "Region"),
                                            ("Mã Farm", "SCODE"),
                                            ("Tên Farm", "SNAME"),
                                            ("Nhóm", "PCLASS"),
                                            ("Mã VECrops", "VECrops Code"),
                                            ("Mã VinEco", "PCODE"),
                                            ("Tên VinEco", "PNAME")
                                        })
                {
                    if (dt.Columns.Contains(key.oldName))
                    {
                        dt.Columns[key.oldName].ColumnName = key.newName;
                    }
                }

                // Main Loop
                foreach (DataColumn dc in dt.Columns)
                {
                    // Loop for every column. Only stop at column with Date at the top ( Indicating it being PurchaseOrder for that Date )
                    // if (DateTime.TryParse(dc.ColumnName, out dateValue))
                    if (StringToDate(dc.ColumnName) == null)
                    {
                        continue;
                    }

                    DateTime dateValue = StringToDate(dc.ColumnName) ?? DateTime.MinValue;

                    ForecastDate fc      = null;
                    var          isNewFC = false;

                    // Find PurchaseOrder for that Date
                    if (dicFc.TryGetValue(dateValue.Date, out Dictionary<string, Dictionary<string, Guid>> _dicProduct))
                    {
                        fc = forecasts.FirstOrDefault(x => x.DateForecast.Date == dateValue.Date);
                    }
                    else
                    {
                        // Create a blank one in case it doesn't exist
                        isNewFC = true;

                        fc = new ForecastDate
                                 {
                                     _id                 = Guid.NewGuid(),
                                     DateForecast        = dateValue.Date,
                                     ListProductForecast = new List<ProductForecast>()
                                 };
                        fc.ForecastDateId = fc._id;

                        dicFc.Add(dateValue.Date, new Dictionary<string, Dictionary<string, Guid>>());
                    }

                    // First layer
                    // Get the list of all Products being Ordered that day.
                    List<ProductForecast> listProductForecast = fc.ListProductForecast ?? new List<ProductForecast>();

                    // Loop for every value
                    foreach (DataRow dr in dt.Rows)
                    {
                        // In case of empty SCODE. I really hate to deal with this case. Like, really.
                        if (string.IsNullOrEmpty(dr["SCODE"]?.ToString()))
                        {
                            dr["SCODE"] = dr["SNAME"]; // Oh for god's sake.
                        }

                        // If OrderQuantity is not 0 - Not Anymore?
                        // object _OrderQuantity = dr[dc.ColumnName];
                        if (dr["PCODE"] == DBNull.Value || supplierType == "ThuMua" && dr["SCODE"] == DBNull.Value)
                        {
                            continue;
                        }

                        // Olala
                        List<SupplierForecast> listSupplierForecast = null;
                        SupplierForecast       supplierForecast     = null;
                        ProductForecast        productForecast      = null;

                        // Olala2
                        var isNewProductOrder  = false;
                        var isNewCustomerOrder = false;

                        // Olala3

                        // #RandomGreenStuff
                        _dicProduct = dicFc[dateValue.Date];
                        if (_dicProduct.TryGetValue(dr["PCODE"].ToString(), out Dictionary<string, Guid> dicStore))
                        {
                            if (!dicProducts.TryGetValue(dr["PCODE"].ToString(), out Product _product))
                            {
                                _product = dicProducts.Values
                                                      .FirstOrDefault(x => x.ProductCode == dr["PCODE"].ToString());
                                if (_product                                             == null)
                                {
                                    _product = new Product
                                                   {
                                                       _id           = Guid.NewGuid(),
                                                       ProductCode   = dr["PCODE"].ToString(),
                                                       ProductName   = dr["PNAME"].ToString(),
                                                       //ProductVECode = dt.Columns.Contains("VECrops Code")
                                                       //                    ? dr["VECrops Code"].ToString()
                                                       //                    : string.Empty
                                                   };


                                    // _product.ProductClassification = dr["PCLASS"].ToString();
                                    _product.ProductId = _product._id;

                                    products.Add(_product);

                                    dicProducts.Add(dr["PCODE"].ToString(), _product);
                                }
                            }

                            productForecast = fc.ListProductForecast
                                                .FirstOrDefault(x => x.ProductId == _product.ProductId);

                            if (dicStore.TryGetValue(dr["SCODE"].ToString(), out Guid id))
                            {
                                supplierForecast = productForecast.ListSupplierForecast
                                                                  .FirstOrDefault(x => x.SupplierId == id);
                            }
                            else
                            {
                                isNewCustomerOrder = true;

                                if (!dicSuppliers.TryGetValue(dr["SCODE"].ToString(), out Supplier supplier))
                                {
                                    supplier = dicSuppliers.Values
                                                           .FirstOrDefault(x => x.SupplierCode == dr["SCODE"].ToString());
                                    if (supplier                                               == null)
                                    {
                                        supplier = new Supplier
                                                       {
                                                           _id          = Guid.NewGuid(),
                                                           SupplierCode = dr["SCODE"]
                                                              .ToString(),
                                                           SupplierName = dr["SNAME"].ToString(),
                                                           SupplierType = supplierType == "ThuMua" && dr["Tag"].ToString() == "VCM"
                                                                              ? "VCM"
                                                                              : supplierType
                                                       };

                                        // SupplierType == "ThuMua" ? (dr["SCODE"] == null ? dr["SCODE"].ToString() : dr["SNAME"].ToString()) : dr["SNAME"].ToString();
                                        supplier.SupplierId = supplier._id;

                                        string region = dr["Region"].ToString();
                                        switch (region)
                                        {
                                            case "LD":
                                                region = "Lâm Đồng";
                                                break;
                                            case "MB":
                                                region = "Miền Bắc";
                                                break;
                                            case "MN":
                                                region = "Miền Nam";
                                                break;
                                            default: break;
                                        }

                                        supplier.SupplierRegion = region;
                                        supplier.SupplierType   =
                                            supplierType == "ThuMua" && dr["Tag"].ToString() == "VCM"
                                                ? "VCM"
                                                : supplierType;

                                        suppliers.Add(supplier);
                                        dicSuppliers.Add(supplier.SupplierCode, supplier);
                                    }
                                }

                                supplierForecast = new SupplierForecast
                                                       {
                                                           _id        = Guid.NewGuid(),
                                                           SupplierId = supplier.SupplierId
                                                       };
                                supplierForecast.SupplierForecastId = supplierForecast._id;

                                dicFc[dateValue.Date][dr["PCODE"].ToString()]
                                   .Add(dr["SCODE"].ToString(), supplier.SupplierId);
                            }
                        }
                        else
                        {
                            isNewProductOrder  = true;
                            isNewCustomerOrder = true;

                            if (!dicProducts.TryGetValue(dr["PCODE"].ToString(), out Product product))
                            {
                                product = dicProducts.Values
                                                     .FirstOrDefault(x => x.ProductCode == dr["PCODE"].ToString());
                                if (product                                             == null)
                                {
                                    product = new Product
                                                  {
                                                      _id           = Guid.NewGuid(),
                                                      ProductCode   = dr["PCODE"].ToString(),
                                                      ProductName   = dr["PNAME"].ToString(),
                                                      //ProductVECode = dt.Columns.Contains("VECrops Code")
                                                      //                    ? dr["VECrops Code"].ToString()
                                                      //                    : string.Empty
                                                  };
                                    product.ProductId = product._id;

                                    // _product.ProductClassification = dr["PCLASS"].ToString();
                                    products.Add(product);

                                    dicProducts.Add(dr["PCODE"].ToString(), product);
                                }
                            }

                            productForecast = new ProductForecast
                                                  {
                                                      _id       = Guid.NewGuid(),
                                                      ProductId = product.ProductId
                                                  };

                            productForecast.ProductForecastId = productForecast._id;

                            productForecast.ListSupplierForecast = new List<SupplierForecast>();

                            if (!dicSuppliers.TryGetValue(dr["SCODE"].ToString(), out Supplier _supplier))
                            {
                                _supplier = dicSuppliers.Values
                                                        .FirstOrDefault(x => x.SupplierCode == dr["SCODE"].ToString());
                                if (_supplier                                               == null)
                                {
                                    _supplier = new Supplier
                                                    {
                                                        _id          = Guid.NewGuid(),
                                                        SupplierCode = dr["SCODE"]
                                                           .ToString(),
                                                        SupplierName = dr["SNAME"].ToString(),
                                                        SupplierType = supplierType == "ThuMua" && dr["Tag"].ToString() == "VCM"
                                                                           ? "VCM"
                                                                           : supplierType
                                                    };

                                    // SupplierType == "ThuMua" ? (dr["SCODE"] == null ? dr["SCODE"].ToString() : dr["SNAME"].ToString()) : dr["SNAME"].ToString();
                                    _supplier.SupplierId = _supplier._id;

                                    string region = dr["Region"].ToString();
                                    switch (region)
                                    {
                                        case "LD":
                                            region = "Lâm Đồng";
                                            break;
                                        case "MB":
                                            region = "Miền Bắc";
                                            break;
                                        case "MN":
                                            region = "Miền Nam";
                                            break;
                                        default: break;
                                    }

                                    _supplier.SupplierRegion = region;
                                    _supplier.SupplierType   =
                                        supplierType == "ThuMua" && dr["Tag"].ToString() == "VCM"
                                            ? "VCM"
                                            : supplierType;

                                    suppliers.Add(_supplier);
                                    dicSuppliers.Add(_supplier.SupplierCode, _supplier);
                                }
                            }

                            supplierForecast = new SupplierForecast
                                                   {
                                                       _id        = Guid.NewGuid(),
                                                       SupplierId = _supplier.SupplierId
                                                   };
                            supplierForecast.SupplierForecastId = supplierForecast._id;

                            dicFc[dateValue.Date].Add(dr["PCODE"].ToString(), new Dictionary<string, Guid>());
                            dicFc[dateValue.Date][dr["PCODE"].ToString()]
                               .Add(dr["SCODE"].ToString(), _supplier.SupplierId);
                        }

                        // Filling in data
                        listSupplierForecast = productForecast.ListSupplierForecast;

                        // Special part for ThuMua
                        TextInfo myTi = new CultureInfo("en-US", false).TextInfo;
                        if (supplierType != "VinEco" && !yesNoKpi)
                        {
                            supplierForecast.QualityControlPass = !string.IsNullOrEmpty(dr["QC"].ToString()) && myTi.ToTitleCase(dr["QC"].ToString()) == "Ok";

                            supplierForecast.LabelVinEco = !string.IsNullOrEmpty(dr["Label VE"].ToString())    && myTi.ToTitleCase(dr["Label VE"].ToString())    == "Yes";
                            supplierForecast.FullOrder   = !string.IsNullOrEmpty(dr["100%"].ToString())        && myTi.ToTitleCase(dr["100%"].ToString())        == "Yes";
                            supplierForecast.CrossRegion = !string.IsNullOrEmpty(dr["CrossRegion"].ToString()) && myTi.ToTitleCase(dr["CrossRegion"].ToString()) == "Yes";
                            supplierForecast.Level       = string.IsNullOrEmpty(dr["Level"].ToString())
                                                               ? Convert.ToByte(0)
                                                               : Convert.ToByte(dr["Level"]);
                            supplierForecast.Availability = string.IsNullOrEmpty(dr["Availability"].ToString())
                                                                ? string.Empty
                                                                : dr["Availability"].ToString();
                        }
                        else if (!yesNoKpi)
                        {
                            supplierForecast.QualityControlPass = true;
                            supplierForecast.LabelVinEco        = true;
                            supplierForecast.FullOrder          = false;
                            supplierForecast.CrossRegion        = false;
                            supplierForecast.Level              = 1;
                            supplierForecast.Availability       = "1234567";

                            string code = dr["PCODE"].ToString();

                            if (supplierType         == "VinEco" &&
                                code.Substring(0, 1) == "K")
                            {
                                supplierForecast.CrossRegion = true;
                            }

                            // To deal with some Supplier only Supply for a targetted Customer Group.
                            supplierForecast.Target = dt.Columns.Contains("Target") ? dr["Target"].ToString() : "All";
                        }
                        else if (yesNoKpi && dr["Source"].ToString() == "ThuMua")
                        {
                            supplierForecast.QualityControlPass = string.IsNullOrEmpty(dr["QC"].ToString())
                                                                      ? supplierForecast.QualityControlPass
                                                                      : myTi.ToTitleCase(dr["QC"].ToString()) == "Ok";
                            supplierForecast.LabelVinEco = string.IsNullOrEmpty(dr["Label VE"].ToString())
                                                               ? supplierForecast.LabelVinEco
                                                               : myTi.ToTitleCase(dr["Label VE"].ToString()) == "Yes";
                            supplierForecast.FullOrder = string.IsNullOrEmpty(dr["100%"].ToString())
                                                             ? supplierForecast.FullOrder
                                                             : myTi.ToTitleCase(dr["100%"].ToString()) == "Yes";
                            supplierForecast.CrossRegion = string.IsNullOrEmpty(dr["CrossRegion"].ToString())
                                                               ? supplierForecast.CrossRegion
                                                               : myTi.ToTitleCase(dr["CrossRegion"].ToString()) == "Yes";
                            supplierForecast.Level = string.IsNullOrEmpty(dr["Level"].ToString())
                                                         ? supplierForecast.Level
                                                         : Convert.ToByte(dr["Level"]);
                            supplierForecast.Availability = string.IsNullOrEmpty(dr["Availability"].ToString())
                                                                ? supplierForecast.Availability
                                                                : dr["Availability"].ToString();
                        }

                        // Todo - Special treatment goes here.
                        if (supplierType == "VinEco")
                        {
                            string code   = dr["PCODE"].ToString();
                            string region = dr["Region"].ToString();

                            if (code.Substring(0, 1) == "K" &&
                                (region              == "MN" || region == "Miền Nam"))
                            {
                                // dicCrossRegionVinEco.ContainsKey(dr["PCODE"].ToString()))
                                if (code == "K03501" ||
                                    code == "K01901" ||
                                    code == "K02201")
                                {
                                    supplierForecast.CrossRegion = false;
                                }

                                // supplierForecast.CrossRegion = true;

                                // if (dr["PCODE"].ToString() == "K06501" 
                                //    || dr["PCode"].ToString() == "K06601")
                                // {
                                //    supplierForecast.CrossRegion = true;
                                //    supplierForecast.FullOrder = true;
                                // }
                            }

                            // February 20, 2018.
                            // Special cases for Cross region.
                            // Todo - Cross Region Special treatments go here.
                            if (code == "B00301" || // Hành tây
                                code == "D01901" || // Khoai tây hồng
                                code == "D02001")   // Khoai tây vàng
                            {
                                supplierForecast.CrossRegion = false;
                            }
                        }

                        ///// <! For debugging purposes !>
                        // if (dateValue.Day == 16 && (string)dr["PCODE"] == "A04201" && (string)dr["SCODE"] == "AG03030000")
                        // {
                        // byte AmIHandsome = 0;
                        // }

                        // 3rd FC layer - Normal Forecast.
                        if (double.TryParse((dr[dc.ColumnName] == DBNull.Value ? 0 : dr[dc.ColumnName]).ToString(), out double quantityForecast))
                        {
                            if (!yesNoKpi)
                            {
                                supplierForecast.QuantityForecast += quantityForecast;
                            }

                            // 2nd FC layer - Minimum / Contracted Forecast - 2nd Highest Priority. 
                            if (dt.Columns.Contains("Min"))
                            {
                                if (double.TryParse((dr["Min"] == DBNull.Value ? 0 : dr["Min"]).ToString(), out double quantityForecastContracted))
                                {
                                    supplierForecast.QuantityForecastContracted += quantityForecastContracted;
                                }
                            }
                        }

                        // if (YesNoKPI &&
                        // Convert.ToDateTime(dr["EffectiveFrom"]).Date <=
                        // DateTime.Parse(dc.ColumnName).Date && Convert.ToDateTime(dr["EffectiveTo"]).Date >=
                        // DateTime.Parse(dc.ColumnName).Date)
                        if (yesNoKpi && StringToDate(dr["EffectiveFrom"].ToString())?.Date <= dateValue.Date && StringToDate(dr["EffectiveTo"].ToString())?.Date >= dateValue.Date)
                        {
                            supplierForecast.QualityControlPass = true;
                            if (double.TryParse((dr[dc.ColumnName] == DBNull.Value ? 0 : dr[dc.ColumnName]).ToString(), out double quantityForecastPlanned))
                            {
                                supplierForecast.QuantityForecastPlanned =  supplierForecast.QuantityForecastPlanned ?? 0;
                                supplierForecast.QuantityForecastPlanned += quantityForecastPlanned;

                                // In case outside of Forecast, which, is an entirely new Supplier.
                                // Yes this does happen.
                                supplierForecast.QualityControlPass = true;

                                // _SupplierForecast.QuantityForecastContracted = Math.Max(_SupplierForecast.QuantityForecastContracted - _SupplierForecast.QuantityForecastPlanned, 0);
                                // _SupplierForecast.QuantityForecast = Math.Max(_SupplierForecast.QuantityForecast - _SupplierForecast.QuantityForecastPlanned - _SupplierForecast.QuantityForecastContracted, 0);
                            }
                        }

                        if (isNewCustomerOrder)
                        {
                            listSupplierForecast.Add(supplierForecast);
                        }

                        productForecast.ListSupplierForecast = listSupplierForecast;
                        if (isNewProductOrder)
                        {
                            fc.ListProductForecast.Add(productForecast);
                        }
                    }

                    fc.ListProductForecast = listProductForecast;

                    if (isNewFC)
                    {
                        forecasts.Add(fc);
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        ///     Reading Forecast
        /// </summary>
        /// <param name="fileVe">
        ///     The file Ve.
        /// </param>
        /// <param name="fileTm">
        ///     The file TM.
        /// </param>
        /// <param name="yesNoPlanning">
        ///     The Yes No Planning.
        /// </param>
        /// <returns>
        ///     The <see cref="Task" />.
        /// </returns>
        private async Task UpdateFcAsync(string fileVe, string fileTm, bool yesNoPlanning = false)
        {
            try
            {
                var fc = new List<ForecastDate>();

                var            mongoClient = new MongoClient();
                IMongoDatabase db          = mongoClient.GetDatabase("localtest");

                List<Product> products = mongoClient.GetDatabase("localtest")
                                                    .GetCollection<Product>("Product")
                                                    .AsQueryable()
                                                    .ToList();
                var suppliers = new List<Supplier>();

                var dicFc       = new Dictionary<DateTime, Dictionary<string, Dictionary<string, Guid>>>();
                var dicProduct  = new Dictionary<string, Product>();
                var dicSupplier = new Dictionary<string, Supplier>();

                foreach (Product product in products)
                {
                    if (!dicProduct.TryGetValue(product.ProductCode, out _))
                    {
                        dicProduct.Add(product.ProductCode, product);
                    }
                }

                foreach (Supplier supplier in suppliers)
                {
                    if (!dicSupplier.TryGetValue(supplier.SupplierCode, out _))
                    {
                        dicSupplier.Add(supplier.SupplierCode, supplier);
                    }
                }

                // string fileName = string.Empty;
                var listFcFileName = new Dictionary<string, Dictionary<string, bool>>
                                         {
                                             {
                                                 yesNoPlanning
                                                     ? "DBSL Planning.xlsb"
                                                     : fileVe,
                                                 new Dictionary<string, bool>
                                                     {
                                                         {
                                                             "VinEco",
                                                             false
                                                         }
                                                     }
                                             },
                                             {
                                                 fileTm,
                                                 new Dictionary<string, bool>
                                                     {
                                                         {
                                                             "ThuMua",
                                                             false
                                                         }
                                                     }
                                             }
                                         };

                if (!yesNoPlanning)
                {
                    listFcFileName.Add("ThuMua KPI.xlsb", new Dictionary<string, bool> { { "VinEco", true } });
                }

                foreach (string localFileName in listFcFileName.Keys)
                {
                    var       workbook  = new Workbook($"D:\\Documents\\Stuff\\VinEco\\Mastah Project\\{localFileName}");
                    Worksheet worksheet = workbook.Worksheets[0];

                    WriteToRichTextBoxOutput(localFileName, false);
                    EatForecastAspose(
                        fc,
                        worksheet,
                        listFcFileName[localFileName].Keys.First(),
                        dicFc,
                        dicProduct,
                        dicSupplier,
                        products,
                        suppliers,
                        listFcFileName[localFileName].Values.First());
                    WriteToRichTextBoxOutput(" - Done!");
                }

                // Compact Forecasts before importing into Database.
                // All afterward services will be here.
                // Current jobs:
                //   - Deal with Confirmed PO from Purchasing Department.

                // Date layer.
                foreach (ForecastDate forecastDate in fc.OrderByDescending(x => x.DateForecast.Date).Reverse())
                {
                    // Product layer.
                    foreach (ProductForecast productForecast in forecastDate.ListProductForecast
                                                                            .Reverse<ProductForecast>())
                    {
                        // Supplier layer.
                        foreach (SupplierForecast supplierForecast in productForecast.ListSupplierForecast
                                                                                     .Reverse<SupplierForecast>())
                        {
                            if (supplierForecast.FullOrder)
                            {
                                supplierForecast.QuantityForecast = Math.Max(supplierForecast.QuantityForecast, 7);
                            }

                            // <! For debugging Purposes !>
                            // if (_ForecastDate.DateForecast.Day == 16 && Product.Where(x => x.ProductId == _ProductForecast.ProductId).FirstOrDefault().ProductCode == "A04201" && Supplier.Where(x => x.SupplierId == _SupplierForecast.SupplierId).FirstOrDefault().SupplierCode == "AG03030000")
                            // {
                            //    var AmIHandsome = true;
                            // }

                            // Excluding FullOrder cases - Special cases.
                            // Also excluding VinEco cases - Even more special.
                            Supplier supplier = dicSupplier.Values.FirstOrDefault(x => x.SupplierId == supplierForecast.SupplierId);

                            if (supplier == null)
                            {
                                continue;
                            }

                            if (supplier.SupplierType != "VinEco" && !supplierForecast.FullOrder)
                            {
                                // If Purchasing Department already ordered:
                                //  -   Obey it.
                                //  -   Delete Minimum.
                                // Reason behind: 
                                //  -   Purchasing Department has interacted and dealt with Suppliers - their numbers have higher priority over normal Forecasts.
                                if (!yesNoPlanning)
                                {
                                    supplierForecast.QuantityForecastOriginal = supplierForecast.QuantityForecast;
                                    supplierForecast.QuantityForecast         =
                                        supplierForecast.QuantityForecastPlanned ?? supplierForecast.QuantityForecast;
                                    supplierForecast.QuantityForecastContracted =
                                        supplierForecast.QuantityForecastPlanned == null
                                            ? Math.Min(
                                                supplierForecast.QuantityForecastContracted,
                                                supplierForecast.QuantityForecast)
                                            : 0;
                                }

                                //// Old logic.
                                //// In case of Planning, by default FC is Planned.
                                // else if (_Supplier.SupplierType == "ThuMua")
                                // {
                                //     //_SupplierForecast.QuantityForecastPlanned = _SupplierForecast.QuantityForecast;
                                // }
                            }
                            else if (supplier.SupplierType == "VinEco")
                            {
                                if (supplierForecast.QuantityForecastPlanned != null)
                                {
                                    supplierForecast.QuantityForecastPlanned = Math.Min(
                                        supplierForecast.QuantityForecastPlanned ?? 0d,
                                        supplierForecast.QuantityForecast);
                                }

                                if (supplierForecast.QuantityForecastPlanned <= 0d)
                                {
                                    supplierForecast.QuantityForecastPlanned = null;
                                }
                            }

                            // If the Supplier can supply 0 product, well, remove it from the list of Suppliers.
                            if (supplierForecast.QuantityForecastPlanned <= 0d || supplierForecast.QuantityForecast <= 0d)
                            {
                                productForecast.ListSupplierForecast.Remove(supplierForecast);
                            }
                        }

                        // End of Supplier layer.
                        // If the Product has no Supplier, well, remove it from the list of suppliable Products..
                        if (productForecast.ListSupplierForecast.Count == 0)
                        {
                            forecastDate.ListProductForecast.Remove(productForecast);
                        }
                    }

                    // End of Product layer.
                    // If the Harvest Date has no Product to supply, well, remove it from the list of Harvest Date.
                    if (forecastDate.ListProductForecast.Count == 0)
                    {
                        fc.Remove(forecastDate);
                    }
                }

                await db.DropCollectionAsync("Forecast").ConfigureAwait(true);
                await db.GetCollection<ForecastDate>("Forecast").InsertManyAsync(fc).ConfigureAwait(true);

                await db.DropCollectionAsync("Product").ConfigureAwait(true);
                await db.GetCollection<Product>("Product").InsertManyAsync(products).ConfigureAwait(true);

                await db.DropCollectionAsync("Supplier").ConfigureAwait(true);
                await db.GetCollection<Supplier>("Supplier").InsertManyAsync(suppliers).ConfigureAwait(true);

                WriteToRichTextBoxOutput(MethodBase.GetCurrentMethod().Name + " - Done");
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }
    }
}