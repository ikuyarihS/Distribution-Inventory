using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

namespace AllocatingStuff
{
    public partial class MainForm
    {
        private void CoordCaller(CoordStructure coreStructure, object[,] listRegion, string supplierType,
            double upperLimit, string priorityTarget, bool YesNoByUnit = false, bool YesNoContracted = false,
            bool YesNoKPI = false)
        {
            for (var index = 0; index < listRegion.GetUpperBound(1); index++)
            {
                var stopwatch = new Stopwatch();
                stopwatch.Start();

                if (supplierType == "VCM" && (string)listRegion[index, 0] != (string)listRegion[index, 1])
                    continue;

                double upperLimitLocal = upperLimit;

                if (priorityTarget == "VM+" || priorityTarget == "VM+ Priority" && UpperCap > 0)
                    upperLimitLocal = 1.1;

                // Obey upperLimit in B2B cases.
                if (priorityTarget != "B2B" && UpperCap <= -1)
                    upperLimitLocal = -1;

                // Even in unlimited case, ThuMua cap at 100%.
                if (supplierType == "ThuMua" && UpperCap <= -1)
                    upperLimitLocal = 1;

                //Thread newThread = new Thread(() => Coord(coreStructure, (string)ListRegion[index, 0], (string)ListRegion[index, 1], SupplierType, (byte)ListRegion[index, 2], (byte)ListRegion[index, 3], UpperLimit, (bool)ListRegion[index, 4], PriorityTarget, YesNoByUnit, YesNoContracted, YesNoKPI));
                //newThread.Start();
                //newThread.Join();

                Coord(coreStructure, (string)listRegion[index, 0], (string)listRegion[index, 1], supplierType,
                    (byte)listRegion[index, 2], (byte)listRegion[index, 3], upperLimitLocal, (bool)listRegion[index, 4],
                    priorityTarget, YesNoByUnit, YesNoContracted, YesNoKPI);

                stopwatch.Stop();

                WriteToRichTextBoxOutput(
                    $"{string.Concat(listRegion[index, 0].ToString().Split(' ').Select(x => x.First()))} => {string.Concat(listRegion[index, 1].ToString().Split(' ').Select(x => x.First().ToString().ToUpper()))}, {supplierType}{(priorityTarget != "" ? " " + priorityTarget : "")}", false);

                WriteToRichTextBoxOutput(
                    $" UpperLimit = {upperLimitLocal} - Done in {Math.Round(stopwatch.Elapsed.TotalSeconds, 2)}s!");

                //Coord(coreStructure, (string)ListRegion[index, 0], (string)ListRegion[index, 1], SupplierType, (byte)ListRegion[index, 2], (byte)ListRegion[index, 3], UpperLimit, (bool)ListRegion[index, 4], PriorityTarget, YesNoByUnit, YesNoContracted, YesNoKPI);
            }
        }

        // Todo - Thoroughly comment on every line.
        // Todo - In dire need of overhaul / upgrading.
        // Todo - Thoroughly overhaul this. Every little things. Too time-consuming.
        /// <summary>
        ///     The Core of all Algorithm.
        ///     Where everything begins and ends.
        /// </summary>
        private void Coord(CoordStructure coreStructure, string SupplierRegion, string CustomerRegion,
            string SupplierType, byte dayBefore, byte dayLdBefore, double UpperLimit, bool CrossRegion = false,
            string PriorityTarget = "", bool YesNoByUnit = false, bool YesNoContracted = false, bool YesNoKPI = false)
        {
            try
            {
                /// <* IMPORTANTO! *>
                // Nothing shall begin before this happens
                var stopwatch = Stopwatch.StartNew();

                // PO Date Layer.
                //WriteToRichTextBoxOutput(String.Format("{0} => {1}, {2}{3}", String.Concat(SupplierRegion.Split(' ').Select(x => x.First())), String.Concat(CustomerRegion.Split(' ').Select(x => x.First().ToString().ToUpper())), SupplierType, (PriorityTarget != "" ? " " + PriorityTarget : "")), false);
                foreach (var DatePO in coreStructure.dicPO.Keys.OrderByDescending(x => x.Date).Reverse())
                {
                    // Product Layer.
                    foreach (var _Product in coreStructure.dicPO[DatePO].Keys.OrderByDescending(x => x.ProductCode)
                        .Reverse())
                    {
                        //double _MOQ = 0;
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

                        var listMoq = new Dictionary<string, double>
                        {
                            {"K01901", 0.3}, // Chanh có hạt
                            {"K02201", 0.3}, // Chanh không hạt
                            {"C07101", 0.1}, // Ớt ngọt ( chuông ) đỏ
                            {"C07201", 0.1}, // Ớt ngọt ( chuông ) vàng
                            {"C07301", 0.1}, // Ớt ngọt ( chuông ) xanh
                            {"B00201", 0.3}, // Dọc mùng ( bạc hà )
                            {"C01801", 0.3}, // Cà chua cherry đỏ
                            {"C04401", 0.3}, // Đậu bắp xanh
                        };

                        //_MOQ = coreStructure.dicMinimum[_Product.ProductCode.Substring(0, 1)];
                        //// Special cases for Lemon. Apparently it's not Fruit but Spices :\
                        //if (_Product.ProductCode.Substring(0, 1) == "K" &&
                        //    (_Product.ProductCode == "K01901" || _Product.ProductCode == "K02201"))
                        //    _MOQ = 0.3;

                        if (!listMoq.TryGetValue(_Product.ProductCode, out var moq))
                            moq = coreStructure.dicMinimum[_Product.ProductCode.Substring(0, 1)];

                        //}

                        //restartThis:

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

                        //var _ProductCrossRegion = new ProductCrossRegion();
                        if (coreStructure.dicProductCrossRegion.TryGetValue(_Product.ProductId,
                                out var _ProductCrossRegion) && SupplierRegion == "Lâm Đồng")
                            switch (CustomerRegion)
                            {
                                case "Miền Bắc":
                                    if (!_ProductCrossRegion.ToNorth) continue;
                                    break;
                                case "Miền Nam":
                                    if (!_ProductCrossRegion.ToSouth) continue;
                                    break;
                                default:
                                    break;
                            }

                        #region Demand from Chosen Customers.

                        bool CheckDaNang(KeyValuePair<CustomerOrder, bool> x)
                        {
                            return SupplierRegion != "Miền Nam" ||
                                   DatePO.DayOfWeek != DayOfWeek.Tuesday && DatePO.DayOfWeek != DayOfWeek.Friday ||
                                   coreStructure.dicCustomer[x.Key.CustomerId].CustomerRegion
                                       .IndexOf("Đà Nẵng", StringComparison.OrdinalIgnoreCase) < 0;
                        }

                        // Total Order.
                        var SumTarget = coreStructure.dicPO[DatePO][_Product]
                            .Where(x =>
                                coreStructure.dicCustomer[x.Key.CustomerId].CustomerBigRegion == CustomerRegion &&
                                x.Value &&
                                (PriorityTarget == "" || coreStructure.dicCustomer[x.Key.CustomerId].CustomerType ==
                                 PriorityTarget)
                                && CheckDaNang(x))
                            .Sum(x => x.Key.QuantityOrderKg); // Sum of Demand.

                        var SumVM = PriorityTarget.Contains("VM+")
                            ? coreStructure.dicPO[DatePO][_Product]
                                .Where(x =>
                                    coreStructure.dicCustomer[x.Key.CustomerId].CustomerBigRegion == CustomerRegion &&
                                    x.Value &&
                                    coreStructure.dicCustomer[x.Key.CustomerId].CustomerType.Contains("VM") &&
                                    coreStructure.dicCustomer[x.Key.CustomerId].CustomerType != PriorityTarget
                                    && CheckDaNang(x))
                                .Sum(x => x.Key.QuantityOrderKg)
                            : 0; // Sum of Demand.

                        var SumSameRegion = SumTarget + SumVM;

                        if (SupplierRegion == "Lâm Đồng")
                        {
                            var _DatePO = CustomerRegion == "Miền Nam" ? DatePO.AddDays(3) : DatePO.AddDays(-3);
                            if (coreStructure.dicPO.ContainsKey(_DatePO) &&
                                coreStructure.dicPO[_DatePO].ContainsKey(_Product))
                            {
                                var _CustomerRegion = CustomerRegion == "Miền Nam" ? "Miền Bắc" : "Miền Nam";
                                SumTarget += coreStructure.dicPO[_DatePO][_Product]
                                    .Where(x =>
                                        coreStructure.dicCustomer[x.Key.CustomerId].CustomerBigRegion ==
                                        _CustomerRegion && x.Value &&
                                        (PriorityTarget == "" ||
                                         coreStructure.dicCustomer[x.Key.CustomerId].CustomerType ==
                                         PriorityTarget)
                                        && CheckDaNang(x))
                                    .Sum(x => x.Key.QuantityOrderKg);

                                SumVM += PriorityTarget.Contains("VM+")
                                    ? coreStructure.dicPO[_DatePO][_Product]
                                        .Where(x =>
                                            coreStructure.dicCustomer[x.Key.CustomerId].CustomerBigRegion ==
                                            _CustomerRegion && x.Value &&
                                            coreStructure.dicCustomer[x.Key.CustomerId].CustomerType.Contains("VM") &&
                                            coreStructure.dicCustomer[x.Key.CustomerId].CustomerType !=
                                            PriorityTarget
                                            && CheckDaNang(x))
                                        .Sum(x => x.Key.QuantityOrderKg)
                                    : 0; // Sum of Demand.
                            }
                        }

                        #endregion

                        // Optimization. Skip if Demand = 0.
                        //if (sumVCM == 0)
                        //    continue;

                        // To deal with Minimum Order Quantity.
                        double wallet = 0;
                        //var wallet = new Dictionary<Guid, double>();

                        //foreach (var _SupplierId in coreStructure.dicSupplier.Keys)
                        //{
                        //    if (!wallet.ContainsKey(_SupplierId))
                        //        wallet.Add(_SupplierId, 0);
                        //}

                        // Grabbing Suppliers by Harvest days.
                        // One for all, one for Lâm Đồng coz Suppliers from there supply both regions.

                        var _dicProductFC =
                            coreStructure.dicFC.FirstOrDefault(x => x.Key.Date == DatePO.AddDays(-dayBefore));

                        var _dicProductFcLd =
                            coreStructure.dicFC.FirstOrDefault(x => x.Key.Date == DatePO.AddDays(-dayLdBefore));

                        //// Optimization. Skip if No Supplier.
                        //if (_dicProductFC.Value != null && _dicProductFcLd.Value == null)
                        //    continue;

                        if (SumTarget != 0 && _dicProductFC.Value != null)
                        {
                            double sumThuMuaLd = 0;
                            double sumFarmLd = 0;

                            var flagFullOrder = false;

                            #region Supply from Lâm Đồng

                            if (SupplierRegion != "Lâm Đồng" && _dicProductFcLd.Value != null)
                            {
                                // Check if Inventory has stock in other places.
                                // If no, equally distributed stuff.
                                // If yes, hah hah hah no.
                                var dicSupplierLdFC =
                                    _dicProductFcLd.Value.FirstOrDefault(x =>
                                        x.Key.ProductCode == _Product.ProductCode);
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
                                                Convert.ToString((int) DatePO.AddDays(-dayLdBefore).DayOfWeek + 1)))
                                        .Sum(x => x.Key.QuantityForecast);

                                    flagFullOrder = dicSupplierLdFC.Value
                                        .Any(x =>
                                            coreStructure.dicSupplier[x.Key.SupplierId].SupplierRegion == "Lâm Đồng" &&
                                            (x.Key.Target == "All" || x.Key.Target == PriorityTarget) &&
                                            x.Key.FullOrder);
                                }
                            }

                            #endregion

                            var dicSupplierFC =
                                _dicProductFC.Value.FirstOrDefault(x => x.Key.ProductCode == _Product.ProductCode);

                            if (dicSupplierFC.Value != null)
                            {
                                #region Total Supply.

                                var _resultSupplier = dicSupplierFC.Value
                                    .Where(x =>
                                        coreStructure.dicSupplier[x.Key.SupplierId].SupplierRegion == "VinEco" &&
                                        coreStructure.dicSupplier[x.Key.SupplierId].SupplierType == SupplierType &&
                                        (x.Key.Target == "All" || x.Key.Target == PriorityTarget) &&
                                        (SupplierType == "VinEco" || x.Key.Availability.Contains(
                                             Convert.ToString((int) DatePO.AddDays(-dayBefore).DayOfWeek + 1))));

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
                                            Convert.ToString((int) DatePO.AddDays(-dayBefore).DayOfWeek + 1)))
                                    .Sum(x => x.Key.QuantityForecast);

                                //_resultSupplier
                                //    .Sum(x => YesNoKPI ? x.Key.QuantityForecastPlanned : YesNoContracted ? x.Key.QuantityForecastContracted : x.Key.QuantityForecast);

                                if (!flagFullOrder)
                                    flagFullOrder = dicSupplierFC.Value
                                        .Any(x =>
                                            coreStructure.dicSupplier[x.Key.SupplierId].SupplierRegion ==
                                            SupplierRegion &&
                                            (x.Key.Target == "All" || x.Key.Target == PriorityTarget) &&
                                            x.Key.FullOrder);

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
                                                              (CustomerRegion == "Miền Bắc"
                                                                  ? "Miền Nam"
                                                                  : "Miền Bắc") &&
                                                              x.Value)
                                                          .Sum(x => x.Key.QuantityOrderKg), 0);
                                else
                                    sumVE += sumFarmLd + sumThuMuaLd;

                                //if (_resultSupplier
                                //        .FirstOrDefault(x => YesNoKPI || YesNoContracted ? false : x.Key.FullOrder)
                                //        .Key != null)
                                //    flagFullOrder = true;

                                //flagFullOrder = _resultSupplier.Any(x =>
                                //    (YesNoKPI || YesNoContracted)
                                //        ? false
                                //        : x.Key.FullOrder);

                                #endregion

                                if (sumVE > 0)
                                {

                                    #region Rate.

                                    //
                                    // Hack - Freaking need to dissect this part.
                                    // Todo - Further Optimization.

                                    // For fuck sake, this is the hardest to code part.
                                    // Also very important. Too important.

                                    // Rate = Supply / Demand --> Deli = Demand * Rate.
                                    var rate = sumVE / (SumTarget + SumVM);

                                    // If Screw-the-upper-limit flag is up.
                                    if (flagFullOrder)
                                        rate = PriorityTarget == "VM+ VinEco"
                                            ? 1
                                            : (UpperLimit > 0 ? Math.Min(rate, UpperLimit) : rate);
                                    // If it's VinCommerce's Supplier, always 1.
                                    else if (rate < 1 && SupplierType == "VCM" && sumVE > 0)
                                        rate = UpperLimit;
                                    // Otherwise, in case of an UpperLimit, obey it
                                    else if (!flagFullOrder)
                                        if (rate < 1)
                                        {
                                            if (rate < 1 && PriorityTarget != "")
                                            {
                                                rate = Math.Min(sumVE / SumTarget, 1);
                                            }
                                            else if (rate < 1)
                                            {
                                                rate = Math.Min(sumVE / SumSameRegion, 1);
                                                if (rate < 1)
                                                    rate = Math.Min(sumVE / SumTarget, 1);
                                            }
                                            if (rate < 1)
                                                rate = SupplierRegion != "Lâm Đồng" &&
                                                       (YesNoKPI || sumFarm > 0 || sumFarmLd > 0 || sumThuMua > 0 ||
                                                        sumThuMuaLd > 0)
                                                    ? Math.Max(rate, 1)
                                                    : rate;
                                            if (SupplierRegion == "Lâm Đồng" && rate < 1 && PriorityTarget == "")
                                                rate = sumVE / SumSameRegion;
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
                                            rate = (sumFarm + sumFarmLd + sumThuMua + sumThuMuaLd) /
                                                   (SumTarget + SumVM);
                                            rate = SupplierRegion != "Lâm Đồng" &&
                                                   (YesNoKPI || sumFarm > 0 || sumFarmLd > 0 || sumThuMua > 0 ||
                                                    sumThuMuaLd > 0)
                                                ? Math.Max(rate, 1)
                                                : rate;
                                            //}
                                        }

                                    rate = UpperLimit > 0 ? Math.Min(rate, UpperLimit) : rate;
                                    if (_Product.ProductCode.Substring(0, 1) == "K")
                                        rate = Math.Min(rate, 1);
                                    //rate = Math.Max(rate, 1);

                                    #endregion

                                    // Only the bravest would tread deeper.
                                    // ... I was once young, brave and foolish ...

                                    // Customer Layer
                                    foreach (var _CustomerOrder in coreStructure.dicPO[DatePO][_Product]
                                            .Where(x => x.Value).ToDictionary(x => x.Key).Keys
                                            .Where(x => x.QuantityOrderKg >= moq)
                                            .Where(x =>
                                                coreStructure.dicCustomer[x.CustomerId].CustomerBigRegion ==
                                                CustomerRegion &&
                                                (PriorityTarget == string.Empty ||
                                                 coreStructure.dicCustomer[x.CustomerId].CustomerType ==
                                                 PriorityTarget) &&
                                                (x.DesiredRegion == null || x.DesiredRegion == SupplierRegion) &&
                                                (x.DesiredSource == null || x.DesiredSource == SupplierType))
                                            //.OrderByDescending(x => coreStructure.dicCustomer[x.CustomerId].CustomerCode)
                                            .OrderBy(x => x.QuantityOrderKg)
                                            .Reverse())
                                        // Todo - Change this to false when doing Planning
                                        if (true)
                                        {
                                            if (SupplierRegion == "Miền Nam" &&
                                                coreStructure.dicCustomer[_CustomerOrder.CustomerId].CustomerRegion
                                                    .IndexOf("Đà Nẵng", StringComparison.CurrentCultureIgnoreCase) >=
                                                0 &&
                                                (DatePO.DayOfWeek == DayOfWeek.Tuesday ||
                                                 DatePO.DayOfWeek == DayOfWeek.Friday))
                                                continue;

                                            #region Qualified Suppliers.

                                            SupplierForecast _SupplierForecast = null;

                                            var _dicSupplierFC_inner = dicSupplierFC.Value
                                                .Where(x => x.Key.QuantityForecast >= moq)
                                                .Where(x =>
                                                    coreStructure.dicSupplier[x.Key.SupplierId].SupplierRegion ==
                                                    SupplierRegion &&
                                                    coreStructure.dicSupplier[x.Key.SupplierId].SupplierType ==
                                                    SupplierType &&
                                                    (SupplierType == "VinEco" || x.Key.Availability.Contains(
                                                         Convert.ToString(
                                                             (int) DatePO.AddDays(-dayBefore).DayOfWeek + 1))) &&
                                                    (x.Key.Target == "All" || x.Key.Target == PriorityTarget) &&
                                                    (!CrossRegion || x.Key.CrossRegion))
                                                .OrderBy(x => x.Key.Level)
                                                .ThenByDescending(x => x.Key.FullOrder)
                                                .ThenBy(x =>
                                                    coreStructure.dicDeli[DatePO.AddDays(-dayBefore)][_Product][x.Key])
                                                .ThenByDescending(x => x.Key.QuantityForecast)
                                                .ThenByDescending(x => x.Key.LabelVinEco);

                                            var result = _dicSupplierFC_inner
                                                .Where(x =>
                                                    YesNoKPI
                                                        ? x.Key.QuantityForecastPlanned >=
                                                          _CustomerOrder.QuantityOrderKg * rate
                                                        : YesNoContracted
                                                            ? x.Key.QuantityForecastContracted >=
                                                              _CustomerOrder.QuantityOrderKg * rate
                                                            : (x.Key.FullOrder || x.Key.QuantityForecast >=
                                                               _CustomerOrder.QuantityOrderKg * rate));

                                            if (!result.Any())
                                                result = _dicSupplierFC_inner
                                                    .Where(x =>
                                                        YesNoKPI
                                                            ? x.Key.QuantityForecastPlanned >= moq
                                                            : YesNoContracted
                                                                ? x.Key.QuantityForecastContracted >= moq
                                                                : (x.Key.FullOrder || x.Key.QuantityForecast >= moq));

                                            if (!result.Any())
                                                continue;

                                            // Coz for fuck sake, it can return null
                                            var totalSupplier = result.Count();
                                            //_SupplierForecast = result.Key;
                                            if (totalSupplier != 0)
                                            {
                                                var _result = result.Aggregate((l, r) =>
                                                    coreStructure.dicDeli[DatePO.AddDays(-dayBefore)][_Product][l.Key] <
                                                    coreStructure.dicDeli[DatePO.AddDays(-dayBefore)][_Product][r.Key]
                                                        ? l
                                                        : r).Key;
                                                if (_result != null && SupplierType == "ThuMua")
                                                    _SupplierForecast = _result;
                                                else
                                                    _SupplierForecast = result.FirstOrDefault().Key;
                                            }
                                            else
                                            {
                                                // Counter situation where there is no Supplier with Forecast greater than PO
                                                _SupplierForecast = _dicSupplierFC_inner
                                                    .FirstOrDefault(x =>
                                                        YesNoKPI
                                                            ? x.Key.QuantityForecastPlanned >= moq
                                                            : YesNoContracted
                                                                ? x.Key.QuantityForecastContracted >= moq
                                                                : x.Key.QuantityForecast >= moq).Key;

                                                totalSupplier = _dicSupplierFC_inner
                                                    .Count(x =>
                                                        YesNoKPI
                                                            ? x.Key.QuantityForecastPlanned >= moq
                                                            : YesNoContracted
                                                                ? x.Key.QuantityForecastContracted >= moq
                                                                : x.Key.QuantityForecast >= moq);
                                            }

                                            #endregion

                                            var _rate = rate;

                                            if ((sumFarm + sumThuMua) * (sumFarmLd + sumThuMuaLd) > 0)
                                                _rate = Math.Min(_rate, UpperLimit);
                                            if (coreStructure.dicPO[DatePO][_Product].Count <= totalSupplier &&
                                                rate < 1)
                                                _rate = UpperLimit;

                                            _rate = Math.Max(_rate, 1);
                                            //_rate = PriorityTarget == "VM+ VinEco"
                                            //    ? Math.Min(_rate, 1)
                                            //    : _rate;

                                            if (_SupplierForecast == null) continue;
                                            if (!coreStructure.dicCoord.TryGetValue(DatePO,
                                                out var _dicCoordProduct)) continue;
                                            if (!_dicCoordProduct.TryGetValue(_Product, out var _dicCoordCusSup))
                                                continue;
                                            if (!_dicCoordCusSup.TryGetValue(_CustomerOrder,
                                                    out var _SupplierForecastCoord) ||
                                                _SupplierForecastCoord != null) continue;

                                            wallet +=
                                            (!YesNoKPI && !YesNoContracted &&
                                             _SupplierForecast.FullOrder)
                                                ? _CustomerOrder.QuantityOrderKg
                                                : Math.Round(_CustomerOrder.QuantityOrderKg * _rate, 1);

                                            #region MOQ.

                                            if (wallet < moq &&
                                                (YesNoKPI
                                                    ? _SupplierForecast.QuantityForecastPlanned
                                                    : (YesNoContracted
                                                        ? _SupplierForecast.QuantityForecastContracted
                                                        : _SupplierForecast.QuantityForecast)) >= moq)
                                                wallet = moq;

                                            //if (_MOQ == 0.05)
                                            //{
                                            //    // Let's hope this will never be hit.
                                            //    // I fucking do hope that.
                                            //    string OhMyFuckingGodWhy = "Holy shit idk, why, oh god, why";
                                            //}

                                            #endregion

                                            if (wallet < moq && PriorityTarget != "") wallet = moq;

                                            wallet = Math.Max(wallet, moq);
                                            if ( /*wallet >= _MOQ &&*/
                                                _SupplierForecast.QuantityForecast >= moq)
                                            {
                                                //if (sumVE <= 0) { continue; }
                                                // Honestly, this should never be hit
                                                // Jk I changed stuff. This should ALWAYS be hit

                                                //double _QuantityForecast = Math.Min(wallet, _SupplierForecast.QuantityForecast, _CustomerOrder.QuantityOrderKg * _rate);
                                                var _QuantityForecast = new[]
                                                {
                                                    wallet, _SupplierForecast.QuantityForecast,
                                                    _CustomerOrder.QuantityOrderKg * _rate
                                                }.Min();

                                                //if (UpperCap > 0)
                                                //    _QuantityForecast = Math.Min(Math.Max(_CustomerOrder.QuantityOrderKg * UpperLimit, _MOQ), _QuantityForecast);

                                                if (flagFullOrder)
                                                    _QuantityForecast =
                                                        _CustomerOrder.QuantityOrderKg * _rate;
                                                else
                                                {
                                                    _QuantityForecast =
                                                        Math.Round(_QuantityForecast, 1);
                                                    _QuantityForecast = Math.Max(_QuantityForecast,
                                                        moq);
                                                }

                                                #region Unit.

                                                //if (_CustomerOrder.Unit != "Kg")
                                                //{
                                                //    var something = coreStructure.dicProductUnit[_Product.ProductCode].ListRegion.Where(x => x.OrderUnitType == _CustomerOrder.Unit).FirstOrDefault();
                                                //    if (something != null)
                                                //    {
                                                //        double _SaleUnitPer = something.SaleUnitPer;
                                                //        _QuantityForecast = (_QuantityForecast / _MOQ) * _SaleUnitPer;
                                                //    }
                                                //}

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

                                                //// To coup with Supply has custom rates, depending on Region.
                                                ////var _ProductRate = new ProductRate();
                                                //double CrossRegionRate = 1;
                                                //if (!YesNoKPI && SupplierRegion == "Miền Nam" && coreStructure.dicProductRate.TryGetValue(_Product.ProductCode, out var _ProductRate))
                                                //{
                                                //    switch (CustomerRegion)
                                                //    {
                                                //        case "Miền Bắc": CrossRegionRate = _ProductRate.ToNorth; break;
                                                //        case "Miền Nam": CrossRegionRate = _ProductRate.ToSouth; break;
                                                //        default: break;
                                                //    }
                                                //}

                                                //_QuantityForecast *= 1;

                                                // Another Nth attempt at dealing with idk why > 100% for VM+ VinEco
                                                //if (coreStructure.dicCustomer[_CustomerOrder.CustomerId].CustomerType == "VM+ VinEco")
                                                //    _QuantityForecast = Math.Min(_QuantityForecast, _CustomerOrder.QuantityOrderKg);

                                                var newId = Guid.NewGuid();
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
                                                            _Date
                                                        }
                                                    };

                                                //if (PriorityTarget == "VM+ VinEco" && _CustomerOrder.QuantityOrderKg >= _MOQ && _QuantityForecast > Math.Round(_CustomerOrder.QuantityOrderKg, 1))
                                                //{
                                                //    byte ReallyDoodReally = 0;
                                                //}

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
                                                _SupplierForecast.QuantityForecastOriginal -=
                                                    _QuantityForecast;
                                                if (_SupplierForecast.FullOrder &&
                                                    _SupplierForecast.QuantityForecast < moq)
                                                    _SupplierForecast.QuantityForecast = moq * 7;
                                                // To make sure Full Order Supplier will still go.

                                                coreStructure.dicCoord[DatePO][_Product][_CustomerOrder]
                                                    =
                                                    _SupplierForecastCoord;
                                                coreStructure.dicDeli[DatePO.AddDays(-dayBefore)][
                                                    _Product][
                                                    _SupplierForecast] += wallet;

                                                //coreStructure.dicPO[DatePO][_Product][_CustomerOrder] = false;

                                                // Roburst way, might optimize Procedures a little bit better.
                                                // Remove Customers and Suppliers fulfilled their roles.

                                                if (_SupplierForecast.QuantityForecast < moq)
                                                {
                                                    coreStructure.dicFC[DatePO.AddDays(-dayBefore)][
                                                        _Product].Remove(_SupplierForecast);
                                                    dicSupplierFC.Value.Remove(_SupplierForecast);
                                                }

                                                wallet -= _QuantityForecast;
                                            }
                                            coreStructure.dicPO[DatePO][_Product]
                                                .Remove(_CustomerOrder);

                                            if (coreStructure.dicPO[DatePO][_Product].Count == 0)
                                                coreStructure.dicPO[DatePO].Remove(_Product);

                                            if (coreStructure.dicPO[DatePO].Keys.Count == 0)
                                                coreStructure.dicPO.Remove(DatePO);
                                        }
                                        else
                                        {
                                            //    var ValidSupplier = dicSupplierFC.Value
                                            //        .Where(x => x.Key.QuantityForecast > 0)
                                            //        .Where(x =>
                                            //            coreStructure.dicSupplier[x.Key.SupplierId].SupplierRegion ==
                                            //            SupplierRegion &&
                                            //            coreStructure.dicSupplier[x.Key.SupplierId].SupplierType ==
                                            //            SupplierType &&
                                            //            (SupplierType != "VinEco"
                                            //                ? x.Key.Availability.Contains(
                                            //                    Convert.ToString(
                                            //                        (int)DatePO.AddDays(-dayBefore).DayOfWeek + 1))
                                            //                : true) &&
                                            //            (x.Key.Target == "All" || x.Key.Target == PriorityTarget) &&
                                            //            (CrossRegion ? x.Key.CrossRegion : true))
                                            //        .OrderBy(x => x.Key.QuantityForecast);

                                            //    var SupplierCount = ValidSupplier.Count();

                                            //    foreach (var key in ValidSupplier)
                                            //    {
                                            //        var _SupplierForecast = key.Key;

                                            //        var _QuantityForecast =
                                            //            Math.Min(_CustomerOrder.QuantityOrderKg / SupplierCount * rate,
                                            //                _SupplierForecast.QuantityForecast);

                                            //        _QuantityForecast = Math.Round(_QuantityForecast, 1);

                                            //        var newId = Guid.NewGuid();

                                            //        var _Date = DatePO.AddDays(-dayBefore).Date;
                                            //        if (CrossRegion && _SupplierForecast.CrossRegion &&
                                            //            CustomerRegion == "Miền Bắc" &&
                                            //            SupplierRegion ==
                                            //            "Miền Nam" /*&& _Product.ProductCode.Substring(0, 1) == "K"*/ &&
                                            //            (_Date.DayOfWeek == DayOfWeek.Tuesday ||
                                            //             _Date.DayOfWeek == DayOfWeek.Thursday ||
                                            //             _Date.DayOfWeek == DayOfWeek.Saturday))
                                            //            _Date = _Date.AddDays(-1).Date;

                                            //        var _SupplierForecastCoord = new Dictionary<SupplierForecast, DateTime>
                                            //    {
                                            //        {
                                            //            new SupplierForecast
                                            //            {
                                            //                _id = newId,
                                            //                SupplierForecastId = newId,

                                            //                SupplierId = _SupplierForecast.SupplierId,
                                            //                LabelVinEco = _SupplierForecast.LabelVinEco,
                                            //                FullOrder = _SupplierForecast.FullOrder,
                                            //                QualityControlPass =
                                            //                    _SupplierForecast.QualityControlPass,
                                            //                CrossRegion = _SupplierForecast.CrossRegion,
                                            //                Level = _SupplierForecast.Level,
                                            //                Availability = _SupplierForecast.Availability,
                                            //                Target = _SupplierForecast.Target,

                                            //                QuantityForecast = _QuantityForecast
                                            //            },
                                            //            _Date
                                            //        }
                                            //    };

                                            //        // KPI cases
                                            //        if (YesNoKPI)
                                            //        {
                                            //            _SupplierForecast.QuantityForecastPlanned -= _QuantityForecast;
                                            //            _SupplierForecast.QuantityForecastContracted -= _QuantityForecast;
                                            //        }
                                            //        // Minimum cases
                                            //        if (YesNoContracted)
                                            //            _SupplierForecast.QuantityForecastContracted -= _QuantityForecast;
                                            //        // Default cases
                                            //        _SupplierForecast.QuantityForecast -= _QuantityForecast;
                                            //        _SupplierForecast.QuantityForecastOriginal -= _QuantityForecast;
                                            //        if (_SupplierForecast.FullOrder && _SupplierForecast.QuantityForecast <= 0)
                                            //            _SupplierForecast.QuantityForecast = moq * 7;
                                            //        // To make sure Full Order Supplier will still go.

                                            //        var CustomerOrder = new CustomerOrder
                                            //        {
                                            //            Company = _CustomerOrder.Company,
                                            //            CustomerId = _CustomerOrder.CustomerId,
                                            //            _id = Guid.NewGuid(),
                                            //            CustomerOrderId = Guid.NewGuid(),
                                            //            DesiredRegion = _CustomerOrder.DesiredRegion,
                                            //            DesiredSource = _CustomerOrder.DesiredSource,
                                            //            QuantityOrder = _QuantityForecast,
                                            //            QuantityOrderKg = _QuantityForecast,
                                            //            Unit = _CustomerOrder.Unit
                                            //        };

                                            //        _CustomerOrder.QuantityOrderKg -= _QuantityForecast;

                                            //        coreStructure.dicPO[DatePO][_Product].Add(CustomerOrder, false);

                                            //        coreStructure.dicCoord[DatePO][_Product]
                                            //            .Add(CustomerOrder, _SupplierForecastCoord);

                                            //        coreStructure.dicDeli[DatePO.AddDays(-dayBefore)][_Product][
                                            //            _SupplierForecast] += _QuantityForecast;
                                            //    }
                                            //    //coreStructure.dicPO[DatePO][_Product].Remove(_CustomerOrder);

                                            //    if (coreStructure.dicPO[DatePO][_Product].Count == 0)
                                            //        coreStructure.dicPO[DatePO].Remove(_Product);

                                            //    if (coreStructure.dicPO[DatePO].Keys.Count == 0)
                                            //        coreStructure.dicPO.Remove(DatePO);
                                        }
                                }
                            }
                        }
                    }
                }
                //}
                stopwatch.Stop();
                //WriteToRichTextBoxOutput(String.Format(" UpperLimit = {1} - Done in {0}s!", Math.Round(stopwatch.Elapsed.TotalSeconds, 2), UpperLimit));
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