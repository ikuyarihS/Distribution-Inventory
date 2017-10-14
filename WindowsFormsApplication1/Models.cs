using MongoDB.Bson.Serialization.Attributes;
using System;
using System.Collections.Generic;


namespace Models
{

    #region Declaring Model

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
            ProductClassification = "???";
            ProductOrientation = "???";
            ProductClimate = "???";
            ProductionGroup = "???";
            ProductNote = new List<string>();
        }
        [BsonId]
        public Guid _id { get; set; }
        public Guid ProductId { get; set; }
        public string ProductCode { get; set; }
        public string ProductName { get; set; }
        public string ProductVECode { get; set; }
        public string ProductClassification { get; set; }
        public string ProductOrientation { get; set; }
        public string ProductClimate { get; set; }
        public string ProductionGroup { get; set; }
        public List<string> ProductNote { get; set; }
    }
    public class ProductCrossRegion
    {
        public ProductCrossRegion()
        {
            ToNorth = true;
            ToSouth = true;
        }
        [BsonId]
        public Guid _id { get; set; }
        public Guid ProductId { get; set; }
        public bool ToNorth { get; set; }
        public bool ToSouth { get; set; }
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
        public string DesiredRegion { get; set; }
        public string DesiredSource { get; set; }
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
        public string Company { get; set; }
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
        [BsonId]
        public Guid _id { get; set; }
        public Guid SupplierForecastId { get; set; }
        public Guid SupplierId { get; set; }
        public bool LabelVinEco { get; set; }
        public bool FullOrder { get; set; }
        public bool QualityControlPass { get; set; }
        public bool CrossRegion { get; set; }
        public byte Level { get; set; }
        public string Availability { get; set; }
        public string Target { get; set; }
        public double QuantityForecast { get; set; }
        public double QuantityForecastContracted { get; set; }
        public double? QuantityForecastPlanned { get; set; }
        public double QuantityForecastOriginal { get; set; }
    }
    public class ProductRate
    {
        public ProductRate()
        {
        }
        [BsonId]
        public Guid _id { get; set; }
        public Guid ProductId { get; set; }
        public string ProductCode { get; set; }
        public double ToNorth { get; set; }
        public double ToSouth { get; set; }
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
        public Dictionary<Guid, ProductCrossRegion> dicProductCrossRegion;
        public Dictionary<string, ProductRate> dicProductRate;
        public Dictionary<DateTime, Dictionary<Product, Dictionary<SupplierForecast, double>>> dicDeli;
        public Dictionary<string, double> dicMinimum;
        public Dictionary<string, byte> dicTransferDays;
    }
    #endregion

    #region Declaring Model 2 ( Added stuff on top of old stuff )

    public class AllocateDetail
    {
        public AllocateDetail()
        {
            this.AllocateDetailId = _id;
        }
        [BsonId]
        public Guid _id { get; set; }
        public Guid AllocateDetailId { get; set; }
        public Guid ProductId { get; set; }
        public Guid SupplierId { get; set; }
        public Guid CustomerId { get; set; }
        public Guid CustomerOrderId { get; set; }
        public double PickingQuantity { get; set; }
        public double DeliQuantity { get; set; }
    }

    #endregion

}
