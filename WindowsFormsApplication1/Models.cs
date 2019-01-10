using System;
using System.Collections.Generic;
using MongoDB.Bson.Serialization.Attributes;

namespace AllocatingStuff
{
    #region

    #endregion

    #region Declaring Model

    /// <summary>
    ///     The coord result.
    /// </summary>
    public class CoordResult
    {
        [BsonId] public Guid _id { get; set; }

        /// <summary>
        ///     Gets or sets the coord result id.
        /// </summary>
        public Guid CoordResultId { get; set; }

        /// <summary>
        ///     Gets or sets the date order.
        /// </summary>
        [BsonDateTimeOptions(Kind = DateTimeKind.Local)]
        public DateTime DateOrder { get; set; }

        /// <summary>
        ///     Gets or sets the list coord result date.
        /// </summary>
        public List<CoordResultDate> ListCoordResultDate { get; set; }
    }

    public class CoordResultDate
    {
        [BsonId] public Guid _id { get; set; }

        public Guid CoordResultDateId { get; set; }

        public List<CoordinateDate> ListCoordinateDate { get; set; }

        public Guid ProductId { get; set; }
    }

    public class CoordinateDate
    {
        [BsonId] public Guid _id { get; set; }

        public Guid CoordinateDateId { get; set; }

        public Guid CustomerOrderId { get; set; }

        [BsonDateTimeOptions(Kind = DateTimeKind.Local)]
        public DateTime? DateDelier { get; set; }

        public Guid? SupplierOrderId { get; set; }
    }

    public class PurchaseOrderDate
    {
        [BsonId] public Guid _id { get; set; }

        [BsonDateTimeOptions(Kind = DateTimeKind.Local)]
        public DateTime DateOrder { get; set; }

        public List<ProductOrder> ListProductOrder { get; set; }

        public Guid PurchaseOrderDateId { get; set; }
    }

    public class ForecastDate
    {
        [BsonId] public Guid _id { get; set; }

        [BsonDateTimeOptions(Kind = DateTimeKind.Local)]
        public DateTime DateForecast { get; set; }

        public Guid ForecastDateId { get; set; }

        public List<ProductForecast> ListProductForecast { get; set; }
    }

    /// <summary>
    ///     The product.
    /// </summary>
    public class Product
    {
        public Product()
        {
            ProductClassification = "???";
            ProductOrientation = "???";
            ProductClimate = "???";
            ProductionGroup = "???";
            ProductNote = new List<string>();
        }

        [BsonId] public Guid _id { get; set; }

        public string ProductClassification { get; set; }

        public string ProductClimate { get; set; }

        public string ProductCode { get; set; }

        public Guid ProductId { get; set; }

        public string ProductionGroup { get; set; }

        public string ProductName { get; set; }

        public List<string> ProductNote { get; set; }

        public string ProductOrientation { get; set; }

        // public string ProductVECode { get; set; }
    }

    public class ProductCrossRegion
    {
        public ProductCrossRegion()
        {
            ToNorth = true;
            ToSouth = true;
        }

        [BsonId] public Guid _id { get; set; }

        public Guid ProductId { get; set; }

        public bool ToNorth { get; set; }

        public bool ToSouth { get; set; }
    }

    public class ProductUnit
    {
        [BsonId] public Guid _id { get; set; }

        public List<ProductUnitRegion> ListRegion { get; set; }

        public string ProductCode { get; set; }

        public Guid ProductId { get; set; }
    }

    public class ProductUnitRegion
    {
        [BsonId] public Guid _id { get; set; }

        public double OrderUnitPer { get; set; }

        public string OrderUnitType { get; set; }

        public string Region { get; set; }

        public double SaleUnitPer { get; set; }

        public string SaleUnitType { get; set; }
    }

    public class ProductOrder
    {
        [BsonId] public Guid _id { get; set; }

        public List<CustomerOrder> ListCustomerOrder { get; set; }

        public Guid ProductId { get; set; }

        // public Guid ProductOrderId { get; set; }
    }

    public class ProductForecast
    {
        [BsonId] public Guid _id { get; set; }

        public List<SupplierForecast> ListSupplierForecast { get; set; }

        public Guid ProductForecastId { get; set; }

        public Guid ProductId { get; set; }
    }

    public class CustomerOrder
    {
        [BsonId] public Guid _id { get; set; }

        // public string Company { get; set; }

        public Guid CustomerId { get; set; }

        // public Guid CustomerOrderId { get; set; }

        //public string DesiredRegion { get; set; }

        //public string DesiredSource { get; set; }

        public double QuantityOrder { get; set; }

        public double QuantityOrderKg { get; set; }

        public string Unit { get; set; }
    }

    public class Customer
    {
        [BsonId] public Guid _id { get; set; }

        public string Company { get; set; }

        public string CustomerBigRegion { get; set; }

        public string CustomerCode { get; set; }

        public Guid CustomerId { get; set; }

        public string CustomerName { get; set; }

        public string CustomerRegion { get; set; }

        public string CustomerType { get; set; }
    }

    public class Supplier
    {
        [BsonId] public Guid _id { get; set; }

        public string SupplierCode { get; set; }

        public Guid SupplierId { get; set; }

        public string SupplierName { get; set; }

        public string SupplierRegion { get; set; }

        public string SupplierType { get; set; }
    }

    public class SupplierForecast
    {
        public SupplierForecast()
        {
            _id = Guid.NewGuid();
            LabelVinEco = false;
            FullOrder = false;
            QualityControlPass = false;
            CrossRegion = false;
            Level = 1;
            Availability = "1234567";
            Target = "All";
            QuantityForecast = 0;
            QuantityForecastContracted = 0;
            QuantityForecastPlanned = null;
        }

        [BsonId] public Guid _id { get; set; }

        public string Availability { get; set; }

        public bool CrossRegion { get; set; }

        public bool FullOrder { get; set; }

        public bool LabelVinEco { get; set; }

        public byte Level { get; set; }

        public bool QualityControlPass { get; set; }

        public double QuantityForecast { get; set; }

        public double QuantityForecastContracted { get; set; }

        public double QuantityForecastOriginal { get; set; }

        public double? QuantityForecastPlanned { get; set; }

        public Guid SupplierForecastId { get; set; }

        public Guid SupplierId { get; set; }

        public string Target { get; set; }
    }

    public class ProductRate
    {
        [BsonId] public Guid _id { get; set; }

        public string ProductCode { get; set; }

        public Guid ProductId { get; set; }

        public double ToNorth { get; set; }

        public double ToSouth { get; set; }
    }

    public class CoordStructure
    {
        public Dictionary<DateTime,
            Dictionary<Product, Dictionary<CustomerOrder, Dictionary<SupplierForecast, DateTime>>>> dicCoord;

        public Dictionary<Guid, Customer> dicCustomer;

        public Dictionary<DateTime, Dictionary<Product, Dictionary<SupplierForecast, double>>> dicDeli;

        public Dictionary<DateTime, Dictionary<Product, Dictionary<SupplierForecast, bool>>> dicFC;

        public Dictionary<string, double> dicMinimum;

        public Dictionary<DateTime, Dictionary<Product, Dictionary<CustomerOrder, bool>>> dicPO;

        public Dictionary<Guid, Product> dicProduct;

        public Dictionary<Guid, ProductCrossRegion> dicProductCrossRegion;

        public Dictionary<string, ProductRate> dicProductRate;

        public Dictionary<string, ProductUnit> dicProductUnit;

        public Dictionary<Guid, Supplier> dicSupplier;

        public Dictionary<string, byte> dicTransferDays;
    }

    #endregion

    #region Declaring Model 2 ( Added stuff on top of old stuff )

    public class AllocateDetail
    {
        public AllocateDetail()
        {
            AllocateDetailId = _id;
        }

        [BsonId] public Guid _id { get; set; }

        public Guid AllocateDetailId { get; set; }

        public Guid CustomerId { get; set; }

        public Guid CustomerOrderId { get; set; }

        public double DeliQuantity { get; set; }

        public double PickingQuantity { get; set; }

        public Guid ProductId { get; set; }

        public Guid SupplierId { get; set; }
    }

    #endregion
}