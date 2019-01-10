using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Diagnostics.Contracts;
using System.Linq;

namespace AllocatingStuff
{
    #region

    #endregion

    public partial class MainForm
    {
        private readonly Dictionary<string, DateTime> _dicDate =
            new Dictionary<string, DateTime>(StringComparer.OrdinalIgnoreCase);

        // Todo - Thoroughly comment on every line.
        // Todo - In dire need of overhaul / upgrading.
        // Todo - Thoroughly overhaul this. Every little things. Too time-consuming.
        /// <summary>
        ///     The Core of all Algorithm.
        ///     Where everything begins and ends.
        /// </summary>
        private void Coord(
            CoordStructure coreStructure,
            string supplierRegion,
            string customerRegion,
            string supplierType,
            byte dayBefore,
            byte dayLdBefore,
            double upperLimit,
            bool crossRegion = false,
            string priorityTarget = "",
            bool yesNoByUnit = false,
            bool yesNoContracted = false,
            bool yesNoKpi = false)
        {
            Contract.Requires(coreStructure != null);

            try
            {
#pragma warning disable 1587

                /// <* IMPORTANTO! *>
#pragma warning restore 1587

                // Nothing shall begin before this happens
                Stopwatch stopwatch = Stopwatch.StartNew();

                // PO Date Layer.
                // WriteToRichTextBoxOutput(String.Format("{0} => {1}, {2}{3}", String.Concat(SupplierRegion.Split(' ').Select(x => x.First())), String.Concat(CustomerRegion.Split(' ').Select(x => x.First().ToString().ToUpper())), SupplierType, (PriorityTarget != "" ? " " + PriorityTarget : "")), false);
                foreach (DateTime datePo in coreStructure.dicPO.Keys.OrderByDescending(x => x.Date).Reverse())

                    // Product Layer.
                {
                    foreach (Product product in coreStructure.dicPO[datePo]
                        .Keys.OrderByDescending(x => x.ProductCode)
                        .Reverse())
                    {
                        // double _MOQ = 0;
                        // In case they are ordering and checking performance through an unit that's NOT FUCKING KILOGRAM!
                        // if (YesNoByUnit)
                        // {
                        // // Cheapest way to calculate Kg per Unit.
                        // // Man I'm so smart.
                        // _MOQ = _CustomerOrder.QuantityOrderKg / _CustomerOrder.QuantityOrder;
                        // }
                        // ... Otherwise, we're cool boys.
                        // else
                        // {

                        //// Todo - Special treatment.
                        //if (product.ProductCode == "K06501" || product.ProductCode == "K06601")
                        //{
                        //    upperLimit = 1.4;
                        //}

                        var listMoq = new Dictionary<string, double>
                        {
                            {"A05501", 0.3}, // Xà lách carol
                            {"A05601", 0.3}, // Xà lách frisse
                            {"A05701", 0.3}, // Xà lách iceberg
                            {"A05801", 0.3}, // Xà lách lolo tím
                            {"A05901", 0.3}, // Xà lách lolo xanh
                            {"A06001", 0.3}, // Xà lách mỡ
                            {"A06101", 0.3}, // Xà lách oakleaf đỏ
                            {"A06201", 0.3}, // Xà lách oakleaf xanh
                            {"A06301", 0.3}, // Xà lách radicchio (tím búp)
                            {"A06401", 0.3}, // Xà lách rocket
                            {"A06501", 0.3}, // Xà lách romaine
                            {"A06701", 0.3}, // Xà lách salanova đỏ
                            {"A06801", 0.3}, // Xà lách salanova xanh
                            {"B00201", 0.3}, // Dọc mùng (bạc hà)
                            {"C01801", 0.3}, // Cà chua cherry đỏ
                            {"C02001", 0.3}, // Cà chua cherry socola
                            {"C02401", 0.3}, // Cà chua cherry vàng
                            {"C04401", 0.3}, // Đậu bắp xanh
                            {"C07101", 0.1}, // Ớt ngọt (ớt chuông) đỏ
                            {"C07201", 0.1}, // Ớt ngọt (ớt chuông) vàng
                            {"C07301", 0.1}, // Ớt ngọt (ớt chuông) xanh
                            {"C09001", 0.3}, // Cà chua cherry hỗn hợp
                            {"C09901", 0.3}, // Cà chua cherry
                            {"C10001", 0.1}, // Ớt ngọt (ớt chuông) màu
                            {"D00401", 0.15}, // Củ cải trắng
                            {"G01301", 0.3}, // Xà lách batavia tím TC
                            {"G01501", 0.3}, // Xà lách frisse TC
                            {"G01601", 0.3}, // Xà lách iceberg TC
                            {"G01701", 0.3}, // Xà lách lolo tím TC
                            {"G01801", 0.3}, // Xà lách lolo xanh TC
                            {"G01901", 0.3}, // Xà lách mỡ TC
                            {"G02001", 0.3}, // Xà lách oakleaf đỏ TC
                            {"G02101", 0.3}, // Xà lách oakleaf xanh TC
                            {"G02201", 0.3}, // Xà lách romaine TC
                            {"G02301", 0.3}, // Xà lách salanova đỏ TC
                            {"G02401", 0.3}, // Xà lách salanova xanh TC
                            {"G03001", 0.3}, // Xà lách carol TC
                            {"K01901", 0.3}, // Chanh có hạt
                            {"K02201", 0.3}, // Chanh không hạt
                        };

                        // _MOQ = coreStructure.dicMinimum[_Product.ProductCode.Substring(0, 1)];
                        //// Special cases for Lemon. Apparently it's not Fruit but Spices :\
                        // if (_Product.ProductCode.Substring(0, 1) == "K" &&
                        // (_Product.ProductCode == "K01901" || _Product.ProductCode == "K02201"))
                        // _MOQ = 0.3;
                        //if (!listMoq.TryGetValue(product.ProductCode, out double moq))
                        //{
                        //    moq = coreStructure.dicMinimum[product.ProductCode.Substring(0, 1)];
                        //}

                        double moq = 0.01;

                        // }

                        // restartThis:
#pragma warning disable 1587

                        /// <! For Debuging Purposes Only !>
#pragma warning restore 1587

                        // Only uncomment in very specific debugging situation.
                        // if (_Product.ProductCode == "A04801" && DatePO.Day == 26 && CustomerRegion == "Miền Nam" && SupplierRegion == "Miền Nam" && SupplierType == "VCM")
                        // {
                        // string WhatAmIEvenDoing = "I have no freaking idea.";
                        // }

                        // Skip if product is not in the List VinEco supplies.
                        if (supplierType != "VinEco" && product.ProductCode.Substring(0, 1) != "K" &&
                            (priorityTarget == "VM" || priorityTarget == "VM+"))
                        {
                            if (!product.ProductNote.Contains(customerRegion == "Miền Bắc" ? "North" : "South"))
                            {
                                continue;
                            }
                        }

                        // Dealing with cases of some Products that will not go to either region, from Lâm Đồng

                        // var _ProductCrossRegion = new ProductCrossRegion();
                        if (coreStructure.dicProductCrossRegion.TryGetValue(
                                product.ProductId,
                                out ProductCrossRegion _ProductCrossRegion) &&
                            supplierRegion == "Lâm Đồng")
                        {
                            switch (customerRegion)
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
                                default:
                                    break;
                            }
                        }

                        bool CheckDaNang(KeyValuePair<CustomerOrder, bool> x)
                        {
                            return true;
                            //return supplierRegion != "Miền Nam" ||
                            //       datePo.DayOfWeek != DayOfWeek.Tuesday && datePo.DayOfWeek != DayOfWeek.Friday ||
                            //       coreStructure.dicCustomer[x.Key.CustomerId]
                            //                    .CustomerRegion
                            //                    .IndexOf("Đà Nẵng", StringComparison.OrdinalIgnoreCase) <
                            //       0;
                        }

                        // Total Order.
                        double SumTarget = coreStructure.dicPO[datePo][product]
                            .Where(x =>
                                coreStructure.dicCustomer[x.Key.CustomerId].CustomerBigRegion == customerRegion &&
                                x.Value && (priorityTarget == string.Empty ||
                                            coreStructure.dicCustomer[x.Key.CustomerId].CustomerType == priorityTarget
                                ) && CheckDaNang(x))
                            .Sum(x => x.Key.QuantityOrderKg); // Sum of Demand.

                        double SumVM = priorityTarget.Contains("VM+")
                            ? coreStructure.dicPO[datePo][product]
                                .Where(
                                    x =>
                                        coreStructure.dicCustomer[x.Key.CustomerId].CustomerBigRegion ==
                                        customerRegion && x.Value &&
                                        coreStructure.dicCustomer[x.Key.CustomerId].CustomerType.Contains("VM") &&
                                        coreStructure.dicCustomer[x.Key.CustomerId].CustomerType != priorityTarget &&
                                        CheckDaNang(x))
                                .Sum(x => x.Key.QuantityOrderKg)
                            : 0; // Sum of Demand.

                        double SumSameRegion = SumTarget + SumVM;

                        if (supplierRegion == "Lâm Đồng")
                        {
                            DateTime _DatePO = customerRegion == "Miền Nam" ? datePo.AddDays(3) : datePo.AddDays(-3);
                            if (coreStructure.dicPO.ContainsKey(_DatePO) &&
                                coreStructure.dicPO[_DatePO].ContainsKey(product))
                            {
                                string _CustomerRegion = customerRegion == "Miền Nam" ? "Miền Bắc" : "Miền Nam";
                                SumTarget += coreStructure.dicPO[_DatePO][product]
                                    .Where(
                                        x =>
                                            coreStructure.dicCustomer[x.Key.CustomerId].CustomerBigRegion ==
                                            _CustomerRegion && x.Value &&
                                            (priorityTarget == string.Empty ||
                                             coreStructure.dicCustomer[x.Key.CustomerId].CustomerType == priorityTarget)
                                            && CheckDaNang(x))
                                    .Sum(x => x.Key.QuantityOrderKg);

                                SumVM += priorityTarget.Contains("VM+")
                                    ? coreStructure.dicPO[_DatePO][product]
                                        .Where(
                                            x =>
                                                coreStructure.dicCustomer[x.Key.CustomerId].CustomerBigRegion ==
                                                _CustomerRegion && x.Value &&
                                                coreStructure.dicCustomer[x.Key.CustomerId].CustomerType
                                                    .Contains("VM") &&
                                                coreStructure.dicCustomer[x.Key.CustomerId].CustomerType !=
                                                priorityTarget && CheckDaNang(x))
                                        .Sum(x => x.Key.QuantityOrderKg)
                                    : 0; // Sum of Demand.
                            }
                        }

                        // Optimization. Skip if Demand = 0.
                        // if (sumVCM == 0)
                        // continue;

                        // To deal with Minimum Order Quantity.
                        double wallet = 0;

                        // var wallet = new Dictionary<Guid, double>();

                        // foreach (var _SupplierId in coreStructure.dicSupplier.Keys)
                        // {
                        // if (!wallet.ContainsKey(_SupplierId))
                        // wallet.Add(_SupplierId, 0);
                        // }

                        // Grabbing Suppliers by Harvest days.
                        // One for all, one for Lâm Đồng coz Suppliers from there supply both regions.
                        KeyValuePair<DateTime, Dictionary<Product, Dictionary<SupplierForecast, bool>>> _dicProductFC =
                            coreStructure.dicFC.FirstOrDefault(x => x.Key.Date == datePo.AddDays(-dayBefore));

                        KeyValuePair<DateTime, Dictionary<Product, Dictionary<SupplierForecast, bool>>> _dicProductFcLd
                            =
                            coreStructure.dicFC.FirstOrDefault(x => x.Key.Date == datePo.AddDays(-dayLdBefore));

                        //// Optimization. Skip if No Supplier.
                        // if (_dicProductFC.Value != null && _dicProductFcLd.Value == null)
                        // continue;
                        if (SumTarget > 0 && _dicProductFC.Value != null)
                        {
                            double sumThuMuaLd = 0;
                            double sumFarmLd = 0;

                            var flagFullOrder = false;

                            #region Supply from Lâm Đồng

                            if (supplierRegion != "Lâm Đồng" && _dicProductFcLd.Value != null)
                            {
                                // Check if Inventory has stock in other places.
                                // If no, equally distributed stuff.
                                // If yes, hah hah hah no.
                                KeyValuePair<Product, Dictionary<SupplierForecast, bool>> dicSupplierLdFC =
                                    _dicProductFcLd.Value.FirstOrDefault(
                                        x =>
                                            x.Key.ProductCode == product.ProductCode);
                                if (dicSupplierLdFC.Value != null)
                                {
                                    // Check Lâm Đồng
                                    // Please NEVER FullOrder == true.
                                    // var _SupplierThuMuaLd = 
                                    IEnumerable<KeyValuePair<SupplierForecast, bool>> _dicSupplierLdFC = dicSupplierLdFC
                                        .Value
                                        .Where(
                                            x =>
                                                coreStructure.dicSupplier[x.Key.SupplierId].SupplierRegion ==
                                                "Lâm Đồng" &&
                                                (x.Key.Target == "All" || x.Key.Target == priorityTarget) &&
                                                (yesNoKpi
                                                    ? x.Key.QuantityForecastPlanned
                                                    : yesNoContracted
                                                        ? x.Key.QuantityForecastContracted
                                                        : x.Key.QuantityForecast) >
                                                0);

                                    // Normal case
                                    sumFarmLd = _dicSupplierLdFC
                                        .Where(
                                            x =>
                                                coreStructure.dicSupplier[x.Key.SupplierId].SupplierType == "VinEco")
                                        .Sum(x => x.Key.QuantityForecast);

                                    sumThuMuaLd = _dicSupplierLdFC
                                        .Where(
                                            x =>
                                                coreStructure.dicSupplier[x.Key.SupplierId].SupplierType != "VinEco" &&
                                                x.Key.Availability.Contains(
                                                    Convert.ToString((int) datePo.AddDays(-dayLdBefore).DayOfWeek + 1)))
                                        .Sum(x => x.Key.QuantityForecast);

                                    flagFullOrder = dicSupplierLdFC.Value
                                        .Any(
                                            x =>
                                                coreStructure.dicSupplier[x.Key.SupplierId].SupplierRegion ==
                                                "Lâm Đồng" &&
                                                (x.Key.Target == "All" || x.Key.Target == priorityTarget) &&
                                                x.Key.FullOrder);
                                }
                            }

                            #endregion

                            KeyValuePair<Product, Dictionary<SupplierForecast, bool>> dicSupplierFC =
                                _dicProductFC.Value.FirstOrDefault(x => x.Key.ProductCode == product.ProductCode);

                            if (dicSupplierFC.Value != null)
                            {
                                #region Total Supply.

                                IEnumerable<KeyValuePair<SupplierForecast, bool>> _resultSupplier = dicSupplierFC.Value
                                    .Where(
                                        x =>
                                            coreStructure.dicSupplier[x.Key.SupplierId].SupplierRegion == "VinEco" &&
                                            coreStructure.dicSupplier[x.Key.SupplierId].SupplierType == supplierType &&
                                            (x.Key.Target == "All" || x.Key.Target == priorityTarget) &&
                                            (supplierType == "VinEco" ||
                                             x.Key.Availability.Contains(
                                                 Convert.ToString((int) datePo.AddDays(-dayBefore).DayOfWeek + 1))));

                                IEnumerable<KeyValuePair<SupplierForecast, bool>> _dicSupplierFC = dicSupplierFC.Value
                                    .Where(
                                        x =>
                                            coreStructure.dicSupplier[x.Key.SupplierId].SupplierRegion ==
                                            supplierRegion &&
                                            (x.Key.Target == "All" || x.Key.Target == priorityTarget) &&
                                            (yesNoKpi
                                                ? x.Key.QuantityForecastPlanned
                                                : yesNoContracted
                                                    ? x.Key.QuantityForecastContracted
                                                    : x.Key.QuantityForecast) >
                                            0);

                                double sumFarm = _dicSupplierFC
                                    .Where(
                                        x =>
                                            coreStructure.dicSupplier[x.Key.SupplierId].SupplierType == "VinEco")
                                    .Sum(x => x.Key.QuantityForecast);

                                double sumThuMua = _dicSupplierFC
                                    .Where(
                                        x =>
                                            coreStructure.dicSupplier[x.Key.SupplierId].SupplierType != "VinEco" &&
                                            x.Key.Availability.Contains(
                                                Convert.ToString((int) datePo.AddDays(-dayBefore).DayOfWeek + 1)))
                                    .Sum(x => x.Key.QuantityForecast);

                                // _resultSupplier
                                // .Sum(x => YesNoKPI ? x.Key.QuantityForecastPlanned : YesNoContracted ? x.Key.QuantityForecastContracted : x.Key.QuantityForecast);
                                if (!flagFullOrder)
                                {
                                    flagFullOrder = dicSupplierFC.Value
                                        .Any(
                                            x =>
                                                coreStructure.dicSupplier[x.Key.SupplierId].SupplierRegion ==
                                                supplierRegion &&
                                                (x.Key.Target == "All" || x.Key.Target == priorityTarget) &&
                                                x.Key.FullOrder);
                                }

                                double sumVE = sumFarm + sumThuMua;

                                // What the fuck is this for?
                                DateTime _DatePO = supplierRegion == "Miền Bắc"
                                    ? datePo.AddDays(-2).Date
                                    : datePo.AddDays(2).Date;
                                if (customerRegion == "Miền Nam" && coreStructure.dicPO.ContainsKey(_DatePO) &&
                                    coreStructure.dicPO[_DatePO].ContainsKey(product))
                                {
                                    sumVE += Math.Max(
                                        sumFarmLd +
                                        sumThuMuaLd -
                                        coreStructure.dicPO[_DatePO][product]
                                            .Where(
                                                x =>
                                                    coreStructure.dicCustomer[x.Key.CustomerId]
                                                        .CustomerBigRegion ==
                                                    (customerRegion == "Miền Bắc"
                                                        ? "Miền Nam"
                                                        : "Miền Bắc") &&
                                                    x.Value)
                                            .Sum(x => x.Key.QuantityOrderKg),
                                        0);
                                }
                                else
                                {
                                    sumVE += sumFarmLd + sumThuMuaLd;
                                }

                                // if (_resultSupplier
                                // .FirstOrDefault(x => YesNoKPI || YesNoContracted ? false : x.Key.FullOrder)
                                // .Key != null)
                                // flagFullOrder = true;

                                // flagFullOrder = _resultSupplier.Any(x =>
                                // (YesNoKPI || YesNoContracted)
                                // ? false
                                // : x.Key.FullOrder);

                                #endregion

                                if (sumVE > 0)
                                {
                                    #region Rate.

                                    // Hack - Freaking need to dissect this part.
                                    // Todo - Further Optimization.

                                    // For fuck sake, this is the hardest to code part.
                                    // Also very important. Too important.

                                    // Rate = Supply / Demand --> Deli = Demand * Rate.

                                    // Changed.
                                    // First calculate if VinEco has enough to supply every order.
                                    double rate = (sumFarm + sumFarmLd) / (SumTarget + SumVM);

                                    bool flagNoPurchasing = rate >= 1;

                                    //// Todo - Delete this for normal usage.
                                    //if (product.ProductCode == "K06501" || product.ProductCode == "K06601")
                                    //{
                                    //    rate = 1.4;
                                    //}

                                    // If not, now add in Purchasing.
                                    if (rate < 1)
                                    {
                                        rate = sumVE / (SumTarget + SumVM);

                                        // But in this special case, cannot go beyond 1 or it will be wasted.
                                        rate = Math.Min(rate, 1);
                                    }

                                    // If Screw-the-upper-limit flag is up.
                                    if (flagFullOrder)
                                    {
                                        rate = priorityTarget == "VM+ VinEco"
                                            ? 1
                                            : (upperLimit > 0 ? Math.Min(rate, upperLimit) : rate);
                                    }
                                    else if (rate < 1 && supplierType == "VCM" && sumVE > 0)
                                    {
                                        // If it's VinCommerce's Supplier, always 1.
                                        rate = 1;
                                    }
                                    else if (!flagFullOrder)
                                    {
                                        // Otherwise, in case of an UpperLimit, obey it
                                        if (rate < 1)
                                        {
                                            if (rate < 1 && priorityTarget != string.Empty)
                                            {
                                                rate = Math.Min(sumVE / SumTarget, 1);
                                            }
                                            else if (rate < 1)
                                            {
                                                rate = Math.Min(sumVE / SumSameRegion, 1);
                                                if (rate < 1)
                                                {
                                                    rate = Math.Min(sumVE / SumTarget, 1);
                                                }
                                            }

                                            if (rate < 1)
                                            {
                                                rate = supplierRegion != "Lâm Đồng" &&
                                                       (yesNoKpi || sumFarm > 0 || sumFarmLd > 0 || sumThuMua > 0 ||
                                                        sumThuMuaLd > 0)
                                                    ? Math.Max(rate, 1)
                                                    : rate;
                                            }

                                            if (supplierRegion == "Lâm Đồng" && rate < 1 &&
                                                priorityTarget == string.Empty)
                                            {
                                                rate = sumVE / SumSameRegion;
                                            }
                                        }
                                        else if (rate > 1)
                                        {
                                            rate = (sumFarm + sumFarmLd) / (SumTarget + SumVM);

                                            if (rate < 1)
                                            {
                                                // This is the case where we do not have enough from our own Farm, but if
                                                // we take from Purchasing, we'll get a rate greater than 100%. Thus, we
                                                // will recalculate the rate, and then make sure it will not be greate than
                                                // 100%.
                                                // To avoid unneccessary purchasing. Too troublesome.
                                                rate = (sumFarm + sumFarmLd + sumThuMua + sumThuMuaLd) /
                                                       (SumTarget + SumVM);
                                                rate = Math.Min(rate, 1);
                                            }

                                            rate = supplierRegion != "Lâm Đồng" &&
                                                   (yesNoKpi || sumFarm > 0 || sumFarmLd > 0 || sumThuMua > 0 ||
                                                    sumThuMuaLd > 0)
                                                ? Math.Max(rate, 1)
                                                : rate;
                                        }
                                    }

                                    rate = upperLimit > 0 ? Math.Min(rate, upperLimit) : rate;
                                    if (product.ProductCode.Substring(0, 1) == "K")
                                    {
                                        rate = Math.Min(rate, 1);
                                    }

                                    //// Todo - Delete this for normal usage.
                                    //if (product.ProductCode == "K06501" || product.ProductCode == "K06601")
                                    //{
                                    //    rate = 1.4;
                                    //}

                                    // rate = Math.Max(rate, 1);

                                    #endregion

                                    // Only the bravest would tread deeper.
                                    // ... I was once young, brave and foolish ...

                                    // Customer Layer
                                    foreach (CustomerOrder customerOrder in coreStructure.dicPO[datePo][product]
                                            .Where(x => x.Value)
                                            .ToDictionary(x => x.Key)
                                            .Keys
                                            .Where(x => x.QuantityOrderKg >= moq)
                                            .Where(
                                                x =>
                                                    coreStructure.dicCustomer[x.CustomerId].CustomerBigRegion ==
                                                    customerRegion &&
                                                    (priorityTarget == string.Empty ||
                                                     coreStructure.dicCustomer[x.CustomerId].CustomerType ==
                                                     priorityTarget)
                                                // &&
                                                //(x.DesiredRegion                                      == null || x.DesiredRegion == supplierRegion) &&
                                                //(x.DesiredSource                                      == null || x.DesiredSource == supplierType)
                                            )

                                            // .OrderByDescending(x => coreStructure.dicCustomer[x.CustomerId].CustomerCode)
                                            .OrderBy(x => x.QuantityOrderKg)
                                            .Reverse())

                                        // Todo - Change this to false when doing Planning
                                    {
                                        if (supplierRegion == "Miền Nam" &&
                                            coreStructure.dicCustomer[customerOrder.CustomerId]
                                                .CustomerRegion
                                                .IndexOf("Đà Nẵng", StringComparison.CurrentCultureIgnoreCase) >=
                                            0 &&
                                            (datePo.DayOfWeek == DayOfWeek.Tuesday ||
                                             datePo.DayOfWeek == DayOfWeek.Friday))
                                        {
                                            continue;
                                        }

                                        #region Qualified Suppliers.

                                        SupplierForecast _SupplierForecast;

                                        IOrderedEnumerable<KeyValuePair<SupplierForecast, bool>> _dicSupplierFC_inner =
                                            dicSupplierFC.Value.Where(x => x.Key.QuantityForecast >= moq)
                                                .Where(
                                                    x =>
                                                        coreStructure.dicSupplier[x.Key.SupplierId].SupplierRegion ==
                                                        supplierRegion &&
                                                        coreStructure.dicSupplier[x.Key.SupplierId].SupplierType ==
                                                        supplierType &&
                                                        (supplierType == "VinEco" ||
                                                         x.Key.Availability.Contains(
                                                             Convert.ToString(
                                                                 (int) datePo.AddDays(-dayBefore).DayOfWeek + 1))) &&
                                                        (x.Key.Target == "All" || x.Key.Target == priorityTarget) &&
                                                        (!crossRegion || x.Key.CrossRegion))

                                                // Just in case, really
                                                .Where(
                                                    x => !flagNoPurchasing || supplierType != "ThuMua" ||
                                                         coreStructure.dicSupplier[x.Key.SupplierId].SupplierType ==
                                                         "VinEco")
                                                .OrderBy(x => x.Key.Level)
                                                .ThenByDescending(x => x.Key.FullOrder)
                                                .ThenBy(
                                                    x => coreStructure.dicDeli[datePo.AddDays(-dayBefore)][product][
                                                        x.Key])
                                                .ThenByDescending(x => x.Key.QuantityForecast)
                                                .ThenByDescending(x => x.Key.LabelVinEco);

                                        IEnumerable<KeyValuePair<SupplierForecast, bool>> result = _dicSupplierFC_inner
                                            .Where(
                                                x =>
                                                    yesNoKpi
                                                        ? x.Key.QuantityForecastPlanned >=
                                                          customerOrder.QuantityOrderKg * rate
                                                        : yesNoContracted
                                                            ? x.Key.QuantityForecastContracted >=
                                                              customerOrder.QuantityOrderKg * rate
                                                            : x.Key.FullOrder || x.Key.QuantityForecast >=
                                                              customerOrder.QuantityOrderKg * rate)
                                            .ToList();

                                        if (!result.Any())
                                        {
                                            result = _dicSupplierFC_inner
                                                .Where(x =>
                                                    yesNoKpi
                                                        ? x.Key.QuantityForecastPlanned >= moq
                                                        : yesNoContracted
                                                            ? x.Key.QuantityForecastContracted >= moq
                                                            : x.Key.FullOrder || x.Key.QuantityForecast >= moq);
                                        }

                                        if (!result.Any())
                                        {
                                            continue;
                                        }

                                        // Coz for fuck sake, it can return null
                                        int totalSupplier = result.Count();

                                        // _SupplierForecast = result.Key;
                                        if (totalSupplier != 0)
                                        {
                                            SupplierForecast _result = result.Aggregate(
                                                    (l, r) =>
                                                        coreStructure.dicDeli[datePo.AddDays(-dayBefore)][product]
                                                            [l.Key] <
                                                        coreStructure.dicDeli[datePo.AddDays(-dayBefore)][product][
                                                            r.Key]
                                                            ? l
                                                            : r)
                                                .Key;
                                            if (_result != null && supplierType == "ThuMua")
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
                                            // Counter situation where there is no Supplier with Forecast greater than PO
                                            _SupplierForecast = _dicSupplierFC_inner
                                                .FirstOrDefault(
                                                    x =>
                                                        yesNoKpi
                                                            ? x.Key.QuantityForecastPlanned >= moq
                                                            : yesNoContracted
                                                                ? x.Key.QuantityForecastContracted >= moq
                                                                : x.Key.QuantityForecast >= moq)
                                                .Key;

                                            totalSupplier = _dicSupplierFC_inner
                                                .Count(
                                                    x =>
                                                        yesNoKpi
                                                            ? x.Key.QuantityForecastPlanned >= moq
                                                            : yesNoContracted
                                                                ? x.Key.QuantityForecastContracted >= moq
                                                                : x.Key.QuantityForecast >= moq);
                                        }

                                        #endregion

                                        double _rate = rate;

                                        if ((sumFarm + sumThuMua) * (sumFarmLd + sumThuMuaLd) > 0)
                                        {
                                            _rate = Math.Min(_rate, upperLimit);
                                        }

                                        if (coreStructure.dicPO[datePo][product].Count <= totalSupplier && rate < 1)
                                        {
                                            _rate = upperLimit;
                                        }

                                        _rate = Math.Max(_rate, 1);

                                        // _rate = PriorityTarget == "VM+ VinEco"
                                        // ? Math.Min(_rate, 1)
                                        // : _rate;
                                        if (_SupplierForecast == null)
                                        {
                                            continue;
                                        }

                                        if (!coreStructure.dicCoord.TryGetValue(
                                            datePo,
                                            out Dictionary<Product, Dictionary<CustomerOrder,
                                                Dictionary<SupplierForecast, DateTime>>> _dicCoordProduct))
                                        {
                                            continue;
                                        }

                                        if (!_dicCoordProduct.TryGetValue(
                                            product,
                                            out Dictionary<CustomerOrder, Dictionary<SupplierForecast, DateTime>>
                                                _dicCoordCusSup))
                                        {
                                            continue;
                                        }

                                        if (!_dicCoordCusSup.TryGetValue(
                                                customerOrder,
                                                out Dictionary<SupplierForecast, DateTime> _SupplierForecastCoord) ||
                                            _SupplierForecastCoord != null)
                                        {
                                            continue;
                                        }

                                        wallet +=
                                            !yesNoKpi && !yesNoContracted && _SupplierForecast.FullOrder
                                                ? customerOrder.QuantityOrderKg
                                                : Math.Round(customerOrder.QuantityOrderKg * _rate, 1);

                                        #region MOQ.

                                        if (wallet < moq &&
                                            (yesNoKpi
                                                ? _SupplierForecast.QuantityForecastPlanned
                                                : (yesNoContracted
                                                    ? _SupplierForecast.QuantityForecastContracted
                                                    : _SupplierForecast.QuantityForecast)) >=
                                            moq)
                                        {
                                            wallet = moq;
                                        }

                                        // if (_MOQ == 0.05)
                                        // {
                                        // // Let's hope this will never be hit.
                                        // // I fucking do hope that.
                                        // string OhMyFuckingGodWhy = "Holy shit idk, why, oh god, why";
                                        // }

                                        #endregion

                                        if (wallet < moq && priorityTarget != string.Empty)
                                        {
                                            wallet = moq;
                                        }

                                        wallet = Math.Max(wallet, moq);
                                        if ( /*wallet >= _MOQ &&*/
                                            _SupplierForecast.QuantityForecast >= moq)
                                        {
                                            // if (sumVE <= 0) { continue; }
                                            // Honestly, this should never be hit
                                            // Jk I changed stuff. This should ALWAYS be hit

                                            // double _QuantityForecast = Math.Min(wallet, _SupplierForecast.QuantityForecast, _CustomerOrder.QuantityOrderKg * _rate);
                                            double _QuantityForecast = new[]
                                            {
                                                wallet,
                                                _SupplierForecast.QuantityForecast,
                                                customerOrder.QuantityOrderKg * _rate
                                            }.Min();

                                            // if (UpperCap > 0)
                                            // _QuantityForecast = Math.Min(Math.Max(_CustomerOrder.QuantityOrderKg * UpperLimit, _MOQ), _QuantityForecast);
                                            if (flagFullOrder)
                                            {
                                                _QuantityForecast =
                                                    customerOrder.QuantityOrderKg * _rate;
                                            }
                                            else
                                            {
                                                _QuantityForecast =
                                                    Math.Round(_QuantityForecast, 1);
                                                _QuantityForecast = Math.Max(
                                                    _QuantityForecast,
                                                    moq);
                                            }

                                            #region Unit.

                                            // if (_CustomerOrder.Unit != "Kg")
                                            // {
                                            // var something = coreStructure.dicProductUnit[_Product.ProductCode].ListRegion.Where(x => x.OrderUnitType == _CustomerOrder.Unit).FirstOrDefault();
                                            // if (something != null)
                                            // {
                                            // double _SaleUnitPer = something.SaleUnitPer;
                                            // _QuantityForecast = (_QuantityForecast / _MOQ) * _SaleUnitPer;
                                            // }
                                            // }

                                            #endregion

                                            // Defer extra days for Crossing Regions ( North --> South and vice versa. )
                                            // To coup with merging PO ( Tue Thu Sat to Mon Wed Fri )
                                            DateTime date = datePo.AddDays(-dayBefore).Date;
                                            if (crossRegion && _SupplierForecast.CrossRegion &&
                                                customerRegion == "Miền Bắc" &&
                                                supplierRegion ==
                                                "Miền Nam" /*&& _Product.ProductCode.Substring(0, 1) == "K"*/ &&
                                                (date.DayOfWeek == DayOfWeek.Tuesday ||
                                                 date.DayOfWeek == DayOfWeek.Thursday ||
                                                 date.DayOfWeek == DayOfWeek.Saturday))
                                            {
                                                date = date.AddDays(-1).Date;
                                            }

                                            //// To coup with Supply has custom rates, depending on Region.
                                            ////var _ProductRate = new ProductRate();
                                            // double CrossRegionRate = 1;
                                            // if (!YesNoKPI && SupplierRegion == "Miền Nam" && coreStructure.dicProductRate.TryGetValue(_Product.ProductCode, out var _ProductRate))
                                            // {
                                            // switch (CustomerRegion)
                                            // {
                                            // case "Miền Bắc": CrossRegionRate = _ProductRate.ToNorth; break;
                                            // case "Miền Nam": CrossRegionRate = _ProductRate.ToSouth; break;
                                            // default: break;
                                            // }
                                            // }

                                            // _QuantityForecast *= 1;

                                            // Another Nth attempt at dealing with idk why > 100% for VM+ VinEco
                                            // if (coreStructure.dicCustomer[_CustomerOrder.CustomerId].CustomerType == "VM+ VinEco")
                                            // _QuantityForecast = Math.Min(_QuantityForecast, _CustomerOrder.QuantityOrderKg);
                                            Guid newId = Guid.NewGuid();
                                            _SupplierForecastCoord =
                                                new Dictionary<SupplierForecast, DateTime>
                                                {
                                                    {
                                                        new SupplierForecast
                                                        {
                                                            _id = newId,
                                                            SupplierForecastId = newId,

                                                            SupplierId =
                                                                _SupplierForecast.SupplierId,
                                                            LabelVinEco =
                                                                _SupplierForecast.LabelVinEco,
                                                            FullOrder = _SupplierForecast.FullOrder,
                                                            QualityControlPass =
                                                                _SupplierForecast
                                                                    .QualityControlPass,
                                                            CrossRegion =
                                                                _SupplierForecast.CrossRegion,
                                                            Level = _SupplierForecast.Level,
                                                            Availability =
                                                                _SupplierForecast.Availability,
                                                            Target = _SupplierForecast.Target,

                                                            QuantityForecast = _QuantityForecast
                                                        },
                                                        date
                                                    }
                                                };

                                            // if (PriorityTarget == "VM+ VinEco" && _CustomerOrder.QuantityOrderKg >= _MOQ && _QuantityForecast > Math.Round(_CustomerOrder.QuantityOrderKg, 1))
                                            // {
                                            // byte ReallyDoodReally = 0;
                                            // }

                                            // KPI cases
                                            if (yesNoKpi)
                                            {
                                                _SupplierForecast.QuantityForecastPlanned -=
                                                    _QuantityForecast;
                                                _SupplierForecast.QuantityForecastContracted -=
                                                    _QuantityForecast;
                                            }
                                            else if (yesNoContracted)
                                            {
                                                // Minimum cases
                                                _SupplierForecast.QuantityForecastContracted -= _QuantityForecast;
                                            }

                                            // Default cases
                                            _SupplierForecast.QuantityForecast -= _QuantityForecast;
                                            _SupplierForecast.QuantityForecastOriginal -=
                                                _QuantityForecast;
                                            if (_SupplierForecast.FullOrder && _SupplierForecast.QuantityForecast < moq)
                                            {
                                                _SupplierForecast.QuantityForecast = moq * 7;
                                            }

                                            // To make sure Full Order Supplier will still go.
                                            coreStructure.dicCoord[datePo][product][customerOrder]
                                                =
                                                _SupplierForecastCoord;
                                            coreStructure.dicDeli[datePo.AddDays(-dayBefore)][
                                                product][
                                                _SupplierForecast] += wallet;

                                            // coreStructure.dicPO[DatePO][_Product][_CustomerOrder] = false;

                                            // Roburst way, might optimize Procedures a little bit better.
                                            // Remove Customers and Suppliers fulfilled their roles.
                                            if (_SupplierForecast.QuantityForecast < moq)
                                            {
                                                coreStructure.dicFC[datePo.AddDays(-dayBefore)][
                                                        product]
                                                    .Remove(_SupplierForecast);
                                                dicSupplierFC.Value.Remove(_SupplierForecast);
                                            }

                                            wallet -= _QuantityForecast;
                                        }

                                        coreStructure.dicPO[datePo][product]
                                            .Remove(customerOrder);

                                        if (coreStructure.dicPO[datePo][product].Count == 0)
                                        {
                                            coreStructure.dicPO[datePo].Remove(product);
                                        }

                                        if (coreStructure.dicPO[datePo].Keys.Count == 0)
                                        {
                                            coreStructure.dicPO.Remove(datePo);
                                        }
                                    }
                                }
                            }
                        }
                    }
                }

                // }
                stopwatch.Stop();

                // WriteToRichTextBoxOutput(String.Format(" UpperLimit = {1} - Done in {0}s!", Math.Round(stopwatch.Elapsed.TotalSeconds, 2), UpperLimit));
            }
            catch (Exception ex)
            {
                WriteToRichTextBoxOutput(ex.Message);
                throw;
            }
        }

        /// <summary>
        ///     The coord caller.
        /// </summary>
        /// <param name="coreStructure"> The core structure. </param>
        /// <param name="listRegion"> The list region. </param>
        /// <param name="supplierType"> The supplier type. </param>
        /// <param name="upperLimit"> The upper limit. </param>
        /// <param name="priorityTarget"> The priority target. </param>
        /// <param name="YesNoByUnit"> The yes no by unit. </param>
        /// <param name="YesNoContracted"> The yes no contracted. </param>
        /// <param name="YesNoKPI"> The yes no kpi. </param>
        private void CoordCaller(
            CoordStructure coreStructure,
            (string From, string To, byte DayForFrom, byte DayForLD, bool IsCrossRegion)[] listRegion,
            string supplierType,
            double upperLimit,
            string priorityTarget,
            bool YesNoByUnit = false,
            bool YesNoContracted = false,
            bool YesNoKPI = false)
        {
            for (var index = 0; index < listRegion.Length; index++)
            {
                var stopwatch = new Stopwatch();
                stopwatch.Start();

                if (supplierType == "VCM" && listRegion[index].From != listRegion[index].To)
                {
                    continue;
                }

                double upperLimitLocal = upperLimit;

                // if (priorityTarget == "VM+" || priorityTarget == "VM+ Priority" && UpperCap > 0)
                // upperLimitLocal = 1.1;
                if (priorityTarget.Contains("VM+") && UpperCap > 0)
                {
                    // upperLimitLocal = priorityTarget == "VM+" && (string) listRegion[index, 1] == "Miền Nam"
                    // ? 1.4
                    // : 1.1;
                    upperLimitLocal = 1.1;
                }

                // Obey upperLimit in B2B cases.
                if (priorityTarget != "B2B" && UpperCap <= -1)
                {
                    upperLimitLocal = -1;
                }

                // if (priorityTarget == "VM+ VinEco")
                // // ReSharper disable once SwitchStatementMissingSomeCases
                // switch ((string) listRegion[index, 1])
                // {
                // case "Miền Bắc":
                // upperLimitLocal = 1.1;
                // break;
                // case "Miền Nam":
                // upperLimitLocal = 1.1;
                // break;
                // }

                // Even in unlimited case, ThuMua cap at 100%.
                // ReSharper disable once SwitchStatementMissingSomeCases
                if (supplierType == "VCM" || supplierType == "ThuMua" && UpperCap <= -1)
                {
                    upperLimitLocal = 1;
                }

                // switch (supplierType)
                // {
                // case "ThuMua" when UpperCap <= -1:
                // upperLimitLocal = 1;
                // break;
                // case "VCM":
                // upperLimitLocal = 1;
                // break;
                // }
                if (UpperCap > 0)
                {
                    upperLimitLocal = Math.Min(upperLimitLocal, UpperCap);
                }

                // Thread newThread = new Thread(() => Coord(coreStructure, (string)ListRegion[index, 0], (string)ListRegion[index, 1], SupplierType, (byte)ListRegion[index, 2], (byte)ListRegion[index, 3], UpperLimit, (bool)ListRegion[index, 4], PriorityTarget, YesNoByUnit, YesNoContracted, YesNoKPI));
                // newThread.Start();
                // newThread.Join();

                // ReSharper disable ArgumentsStyleNamedExpression
                // ReSharper disable ArgumentsStyleOther
                Coord(
                    coreStructure: coreStructure,
                    supplierRegion: listRegion[index].From,
                    customerRegion: listRegion[index].To,
                    supplierType: supplierType,
                    dayBefore: listRegion[index].DayForFrom,
                    dayLdBefore: listRegion[index].DayForLD,
                    upperLimit: upperLimitLocal,
                    crossRegion: listRegion[index].IsCrossRegion,
                    priorityTarget: priorityTarget,
                    yesNoByUnit: YesNoByUnit,
                    yesNoContracted: YesNoContracted,
                    yesNoKpi: YesNoKPI);

                // ReSharper restore ArgumentsStyleNamedExpression
                // ReSharper restore ArgumentsStyleOther
                stopwatch.Stop();

                WriteToRichTextBoxOutput(
                    $"{string.Concat(listRegion[index].Item1.Split(' ').Select(x => x.First()))} => {string.Concat(listRegion[index].Item2.Split(' ').Select(x => x.First().ToString().ToUpper()))}, {supplierType}{(priorityTarget != string.Empty ? " " + priorityTarget : string.Empty)}",
                    false);

                WriteToRichTextBoxOutput(
                    $" UpperLimit = {upperLimitLocal} - Done in {Math.Round(stopwatch.Elapsed.TotalSeconds, 2)}s!");

                // Coord(coreStructure, (string)ListRegion[index, 0], (string)ListRegion[index, 1], SupplierType, (byte)ListRegion[index, 2], (byte)ListRegion[index, 3], UpperLimit, (bool)ListRegion[index, 4], PriorityTarget, YesNoByUnit, YesNoContracted, YesNoKPI);
            }
        }

        /// <summary>
        ///     Convert string to DateTime.
        ///     Optimization.
        /// </summary>
        /// <param name="suspect">String to convert to Date.</param>
        /// <returns>A DateTime value from a string, if convertible.</returns>
        private DateTime? StringToDate(string suspect)
        {
            // If string has been converted before.
            if (_dicDate.TryGetValue(suspect, out DateTime dateResult))
            {
                return dateResult == DateTime.MinValue ? (DateTime?) null : dateResult;
            }

            // Otherwise, check if it's even a date.
            if (!DateTime.TryParse(suspect, out DateTime date))
            {
                // Looks like it isn't.
                // Return null, and also record string used.
                _dicDate.Add(suspect, DateTime.MinValue);
                return null;
            }

            // Welp, it's actually a date.
            // Record the string anyway. Dis many importanto.
            _dicDate.Add(suspect, date);
            return date;
        }
    }
}