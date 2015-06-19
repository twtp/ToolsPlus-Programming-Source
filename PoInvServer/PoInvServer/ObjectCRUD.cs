using System;
using System.Collections.Generic;
using System.Linq;
using Nancy;

using TP.Object.BackEnd;

namespace PoInvServer
{
  // ProductLineCRUD handles:
  //   OBJECT                           READ  CREATE  UPDATE  DELETE
  //   DbProductLine                      Y     Y       Y       N
  //   DbProductLineContactInfo           Y     Y       Y       Y
  //   DbProductLineWholesalePriceLevel   Y     N       Y       N
  //   DbProductLineSalesRep              Y     Y       Y       Y
  //   DbProductLineSubstitution          Y     Y       Y       Y
  //   DbProductLineDropshipCharge        Y     Y       Y       Y
  //   DbProductLineDropshipMatrix        N     Y       Y       Y
  //   DbProductLineCostField             Y     Y       Y       N
  public class ProductLineCRUD : NancyModule
  {
    public ProductLineCRUD() : base("/productline")
    {
      #region DbProductLine - Read

      /// Returns a list of product line info, probably suitable for caching
      /// client-side (probaly reformat into Dict[Prefix] = ID), since the
      /// only values are ProductLineID and ProductLinePrefix.
      /// 
      /// URL Parameters:
      ///   none
      /// Query Parameters:
      ///   none
      /// Returns:
      ///   <ArrayOfBaseProductLine>
      ///     <BaseProductLine>
      ///       <ProductLineID>1</ProductLineID>
      ///       <ProductLinePrefix>AAA</ProductLineID>
      ///     </BaseProductLine>
      ///     ...
      ///   </ArrayOfBaseProductLine>
      Get["/list"] = _ =>
      {
        try {
          IEnumerable<DbBaseProductLine> list = DbProductLine.ListAll();
          return NancyExtensions.Respond(list);
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      /// Returns all info about a single product line, looked up by the autoid
      /// field.
      ///
      /// URL Parameters:
      ///   product line id
      /// Query Parameters:
      ///   preload:
      ///     wholesalepricing
      ///     vendor
      ///     contactinfo
      ///     dropshipcharges
      ///     purchasingsubs
      ///     salesreps
      ///     costfields
      /// Returns:
      ///   <ProductLine>
      ///     <ProductLineID>1</ProductLineID>
      ///     <ProductLinePrefix>AAA</ProductLineID>
      ///     <ProductLineName>TEST PRODUCT LINE</ProductLineName>
      ///     <DefaultVendorNo>TRUS01</DefaultVendorNo>
      ///     ...
      ///   </ProductLine>
      Get["/{id:int}"] = _ =>
      {
        try {
          DbProductLine item = DbProductLine.GetByID(_.id, NancyExtensions.GetPreloads(Request.Query));
          return NancyExtensions.Respond(item);
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      /// Returns all info about a single product line, looked up by the prefix
      /// of the product line. This should really only be used when we don't
      /// know the id to lookup for whatever reason.
      ///
      /// URL Parameters:
      ///   product line prefix (should be ^[-A-Z0-9]{3}$)
      /// Query Parameters:
      ///   preload:
      ///     wholesalepricing
      ///     vendor
      ///     contactinfo
      ///     dropshipcharges
      ///     purchasingsubs
      ///     salesreps
      /// Returns:
      ///   <ProductLine>
      ///     <ProductLineID>1</ProductLineID>
      ///     <ProductLinePrefix>AAA</ProductLineID>
      ///     <ProductLineName>TEST PRODUCT LINE</ProductLineName>
      ///     <DefaultVendorNo>TRUS01</DefaultVendorNo>
      ///     ...
      ///   </ProductLine>
      Get["/{prefix}"] = _ =>
      {
        try {
          DbProductLine item = DbProductLine.GetByPrefix(_.prefix, NancyExtensions.GetPreloads(Request.Query));
          return NancyExtensions.Respond(item);
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      #endregion


      #region DbProductLine - Create

      /// Creates a brand new product line, really just the stub only. Other
      /// fields will be filled in with an update. Only requires the new
      /// product line's prefix as a parameter.
      /// 
      /// URL Parameters:
      ///   none
      /// Query Parameters:
      ///   Prefix: the ^[-0-9A-Z]{3}$ product line prefix code
      /// Returns:
      ///   <ProductLine>
      ///     <ProductLineID>1</ProductLineID>
      ///     <ProductLinePrefix>AAA</ProductLineID>
      ///     <DefaultVendorNumber i:nil="true" />
      ///     ...
      ///   </ProductLine>
      Post["/create"] = _ =>
      {
        try
        {
          DbProductLine item = DbProductLine.CreateProductLine(
              prefix: NancyExtensions.GetBasicParameter<string>(Request.Form, "Prefix")
            );
          return NancyExtensions.Respond(item);
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      #endregion


      #region DbProductLine - Update

      /// Modifies an existing product line, given a new serialized version of
      /// the product line object. Returns the updated, current version of the
      /// object in the database. Note that, like most other updates, this
      /// will overwrite someone else's changes, if there is a race condition.
      /// 
      /// URL Parameters:
      ///   product line id
      /// Query Parameters:
      ///   TP.Object.ProductLine => serialized object with new data
      /// Returns:
      ///   <ProductLine>
      ///     <ProductLineID>1</ProductLineID>
      ///     <ProductLinePrefix>AAA</ProductLineID>
      ///     <DefaultVendorNumber i:nil="true" />
      ///     ...
      ///   </ProductLine>
      Post["/update/{id:int}"] = _ =>
      {
        try
        {
          DbProductLine item = NancyExtensions.GetDbObjectParameter<DbProductLine>(Request.Form, "TP.Object.ProductLine");
          if (item.ProductLineID != _.id)
          {
            throw new ArgumentException("url id does not match given object!");
          }
          return NancyExtensions.Respond(item.UpdateProductLine());
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      #endregion


      #region DbProductLineContactInfo - Read

      /// Returns a list of contact info objects, each one is identified by a
      /// unique InfoType enum.
      /// 
      /// URL Parameters:
      ///   product line id
      /// Query Parameters:
      ///   none
      /// Returns:
      ///   <ArrayOfProductLineContactInfo>
      ///     <ProductLineContactInfo>
      ///       <ProductLineID>191</ProductLineID>
      ///       <InfoType>GeneralContact</InfoType>
      ///       <TelephoneNo>800-235-2000</TelephoneNo>
      ///       ...
      ///     </ProductLineContactInfo>
      ///     ...
      ///   </ArrayOfProductLineContactInfo>
      Get["/{id:int}/contact/"] = _ =>
      {
        try
        {
          IEnumerable<DbProductLineContactInfo> list = DbProductLineContactInfo.GetByProductLineID(_.id, NancyExtensions.GetPreloads(Request.Query));
          return NancyExtensions.Respond(list);
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      #endregion


      #region DbProductLineContactInfo - Create

      /// Creates a new contact info type for the specified product line. Type
      /// must be specified with the base int of the ContactInfoType enum.
      /// Note that all query parameters are optional.
      /// 
      /// URL Parameters:
      ///   product line id
      ///   (int)ProductLineContactInfoType
      /// Query Parameters:
      ///   Phone
      ///   Fax
      ///   Url
      ///   Email
      /// Returns:
      ///     <ProductLineContactInfo>
      ///       <ProductLineID>191</ProductLineID>
      ///       <InfoType>GeneralContact</InfoType>
      ///       <Phone>203-573-0750</Phone>
      ///       ...
      ///     </ProductLineContactInfo>
      Post["/{id:int}/contact/create/{typeid:int}"] = _ =>
      {
        try
        {
          if (!Enum.IsDefined(typeof(TP.Object.ProductLineContactInfo), (int)_.typeid))
          {
            throw new ArgumentException("invalid contact type id");
          }
          DbProductLineContactInfo item = DbProductLineContactInfo.CreateContactInfo(
              productLineID: _.id,
              infoType: (TP.Object.ProductLineContactInfoType)_.typeid,
              telephoneNo: NancyExtensions.GetBasicParameter<string>(Request.Form, "Phone"),
              faxNo: NancyExtensions.GetBasicParameter<string>(Request.Form, "Fax"),
              url: NancyExtensions.GetBasicParameter<string>(Request.Form, "Url"),
              email: NancyExtensions.GetBasicParameter<string>(Request.Form, "Email")
            );
          return NancyExtensions.Respond(item);
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      #endregion


      #region DbProductLineContactInfo - Update

      /// Modifies an existing product line contact info tuple. Can't change
      /// the type of contact, just the actual info, that must be a delete/add
      /// operation.
      /// 
      /// URL Parameters:
      ///   product line id
      ///   (int)ProductLineContactInfoType
      /// Query Parameters:
      ///   TP.Object.ProductLineContactInfo => serialized object with new data
      /// Returns:
      ///   <ProductLineContactInfo>
      ///     <ProductLineID>80</ProductLineID>
      ///     <InfoType>Payables</InfoType>
      ///     <EmailAddress>foo@example.com</EmailAddress>
      ///     <FaxNo i:nil="true" /><HomepageURL i:nil="true" />
      ///     <TelephoneNo>800-222-6133</TelephoneNo>
      ///   </ProductLineContactInfo>
      Post["/{id:int}/contact/update/{typeid:int}"] = _ =>
      {
        try
        {
          if (!Enum.IsDefined(typeof(TP.Object.ProductLineContactInfo), (int)_.typeid))
          {
            throw new ArgumentException("invalid contact type id");
          }
          DbProductLineContactInfo item = NancyExtensions.GetDbObjectParameter<DbProductLineContactInfo>(Request.Form, "TP.Object.ProductLineContactInfo");
          if (item.ProductLineID != _.id)
          {
            throw new ArgumentException("url id does not match given object!");
          }
          if (item.InfoType != (TP.Object.ProductLineContactInfoType)_.typeid)
          {
            throw new ArgumentException("url id does not match given object!");
          }
          return NancyExtensions.Respond(item.UpdateContactInfo());
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      #endregion


      #region DbProductLineContactInfo - Delete

      /// Deletes the given product line contact info tuple.
      /// 
      /// URL Parameters:
      ///   product line id
      ///   (int)ProductLineContactInfoType
      /// Query Parameters:
      ///   TP.Object.ProductLineContactInfo => serialized object with current data
      /// Returns:
      ///   serialized SuccessResponse or ErrorResponse
      Post["/{id:int}/contact/delete/{typeid:int}"] = _ =>
      {
        try
        {
          if (!Enum.IsDefined(typeof(TP.Object.ProductLineContactInfo), (int)_.typeid))
          {
            throw new ArgumentException("invalid contact type id");
          }
          DbProductLineContactInfo item = NancyExtensions.GetDbObjectParameter<DbProductLineContactInfo>(Request.Form, "TP.Object.ProductLineContactInfo");
          if (item.ProductLineID != _.id)
          {
            throw new ArgumentException("url id does not match given object!");
          }
          if (item.InfoType != (TP.Object.ProductLineContactInfoType)_.typeid)
          {
            throw new ArgumentException("url id does not match given object!");
          }
          if (item.DeleteContactInfo())
          {
            return NancyExtensions.Respond(new TP.Base.SuccessResponse("contact info entry deleted"));
          }
          else
          {
            return NancyExtensions.Respond(new TP.Base.ErrorResponse("delete failed"));
          }
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      #endregion


      #region DbProductLineWholesalePriceLevel - Read

      /// Returns a list of wholesale price levels for the line, theoretically
      /// an undefined number, but in practice there are always six, numbered
      /// zero through five, corresponding with Mas price levels A-F.
      /// 
      /// URL Parameters:
      ///   product line id
      /// Query Parameters:
      ///   none
      /// Returns:
      ///   <ArrayOfProductLineWholesalePriceLevel>
      ///     <ProductLineWholesalePriceLevel>
      ///       <ProductLineID>191</ProductLineID>
      ///       <Level>0</Level>
      ///       <Percentage>20</Percentage>
      ///     </ProductLineWholesalePriceLevel>
      ///     ...
      ///   </ArrayOfProductLineContactInfo>
      Get["/{id:int}/wholesale"] = _ =>
      {
        try
        {
          IEnumerable<DbProductLineWholesalePriceLevel> list = DbProductLineWholesalePriceLevel.GetByProductLineID(_.id, NancyExtensions.GetPreloads(Request.Query));
          return NancyExtensions.Respond(list);
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      #endregion


      #region DbProductLineWholesalePriceLevel - Update

      /// Modify an existing product line wholesale pricing level.
      /// 
      /// URL Parameters:
      ///   product line id
      ///   price level (number, not letter)
      /// Query Parameters:
      ///   TP.Object.ProductLineWholesalePriceLevel => serialized object with new data
      /// Returns:
      ///     <ProductLineWholesalePriceLevel>
      ///       <ProductLineID>191</ProductLineID>
      ///       <Level>0</Level>
      ///       <Percentage>15</Percentage>
      ///     </ProductLineWholesalePriceLevel>
      Post["/{id:int}/wholesale/update/{level:int}"] = _ =>
      {
        try
        {
          DbProductLineWholesalePriceLevel item = NancyExtensions.GetDbObjectParameter<DbProductLineWholesalePriceLevel>(Request.Form, "TP.Object.ProductLineWholesalePriceLevel");
          if (item.ProductLineID != _.id)
          {
            throw new ArgumentException("url id does not match given object!");
          }
          if (item.Level != _.level)
          {
            throw new ArgumentException("url id does not match given object!");
          }
          return NancyExtensions.Respond(item.UpdateWholesalePriceLevel());
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      #endregion


      #region DbProductLineSalesRep - Read

      /// Returns a list of product line / sales rep many-to-many link objects.
      /// There should only be one "primary" sales rep per product line, but
      /// this is currently not enforced.
      /// 
      /// URL Parameters:
      ///   product line id
      /// Query Parameters:
      ///   preload:
      ///     reps
      /// Returns:
      ///   <ArrayOfProductLineSalesRep>
      ///     <ProductLineSalesRep>
      ///       <TupleID>141</TupleID>
      ///       <ProductLineID>191</ProductLineID>
      ///       <SalesRepID>40</SalesRepID>
      ///       <IsPrimary>false</IsPrimary>
      ///     </ProductLineSalesRep>
      ///     ...
      ///   </ProductLineSalesRep>
      Get["/{id:int}/salesreps"] = _ =>
      {
        try
        {
          IEnumerable<DbProductLineSalesRep> list = DbProductLineSalesRep.GetByProductLineID(_.id, NancyExtensions.GetPreloads(Request.Query));
          return NancyExtensions.Respond(list);
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      #endregion


      #region DbProductLineSalesRep - Create

      /// Attaches a sales rep to a product line. Always defaults to creating
      /// as a non-primary, update required to change this.
      /// 
      /// URL Parameters:
      ///   product line id
      ///   sales rep id
      /// Query Parameters:
      ///   none -- note: this might need to have an explicit content length zero set client side for the c# lib?
      /// Returns:
      ///     <ProductLineSalesRep>
      ///       <TupleID>141</TupleID>
      ///       <ProductLineID>191</ProductLineID>
      ///       <SalesRepID>40</SalesRepID>
      ///       <IsPrimary>false</IsPrimary>
      ///     </ProductLineSalesRep>
      Post["/{id:int}/salesreps/add/{srid:int}"] = _ =>
      {
        try
        {
          DbProductLineSalesRep item = DbProductLineSalesRep.CreateLinkage(
              productLineID: _.id,
              salesRepID: _.srid,
              isPrimary: false
            );
          return NancyExtensions.Respond(item);
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      #endregion


      #region DbProductLineSalesRep - Update

      /// Modifies the sales rep / product line link fields.
      /// 
      /// URL Parameters:
      ///   product line id
      ///   sales rep id
      /// Query Parameters:
      ///   TP.Object.ProductLineSalesRep => serialized object with new data
      /// Returns:
      ///     <ProductLineSalesRep>
      ///       <TupleID>141</TupleID>
      ///       <ProductLineID>191</ProductLineID>
      ///       <SalesRepID>40</SalesRepID>
      ///       <IsPrimary>false</IsPrimary>
      ///     </ProductLineSalesRep>
      Post["/{id:int}/salesreps/update/{srid:int}"] = _ =>
      {
        try
        {
          DbProductLineSalesRep item = NancyExtensions.GetDbObjectParameter<DbProductLineSalesRep>(Request.Form, "TP.Object.ProductLineSalesRep");
          if (item.ProductLineID != _.id)
          {
            throw new ArgumentException("url id does not match given object!");
          }
          if (item.SalesRepID != _.srid)
          {
            throw new ArgumentException("url id does not match given object!");
          }
          return NancyExtensions.Respond(item.UpdateLinkage());
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      #endregion


      #region DbProductLineSalesRep - Delete

      /// Deletes the given sales rep / product line link. Delete will only be
      /// successful if the object has not been modified from the current
      /// version in the database.
      /// 
      /// URL Parameters:
      ///   product line id
      ///   sales rep id
      /// Query Parameters:
      ///   TP.Object.ProductLineSalesRep => serialized object with current data
      /// Returns:
      ///   serialized SuccessResponse or ErrorResponse
      Post["/{id:int}/salesreps/delete/{srid:int}"] = _ =>
      {
        try
        {
          DbProductLineSalesRep item = NancyExtensions.GetDbObjectParameter<DbProductLineSalesRep>(Request.Form, "TP.Object.ProductLineSalesRep");
          if (item.ProductLineID != _.id)
          {
            throw new ArgumentException("url id does not match given object!");
          }
          if (item.SalesRepID != _.srid)
          {
            throw new ArgumentException("url id does not match given object!");
          }
          if (item.DeleteLinkage())
          {
            return NancyExtensions.Respond(new TP.Base.SuccessResponse("line/rep entry deleted"));
          }
          else
          {
            return NancyExtensions.Respond(new TP.Base.ErrorResponse("delete failed"));
          }
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      #endregion


      #region DbProductLineSubstitution - Read

      /// Returns a substitution object with the id of the product line that
      /// this gets substituted to (i.e., all purchase orders are created
      /// with the PurchasingProductLineID's product line prefix). Note that
      /// this will return a 404 if there is not substitution.
      /// 
      /// URL Parameters:
      ///   product line id
      /// Query Parameters:
      ///   preload:
      ///     productline
      /// Returns:
      ///   <ProductLineSubstitution>
      ///     <TupleID>21</TupleID>
      ///     <ProductLineID>191</ProductLineID>
      ///     <PurchasingProductLineID>509</PurchasingProductLineID>
      ///   </ProductLineSubstitution>
      Get["/{id:int}/substitutes/to"] = _ =>
      {
        try
        {
          DbProductLineSubstitution item = DbProductLineSubstitution.GetByProductLineID(_.id, NancyExtensions.GetPreloads(Request.Query));
          return NancyExtensions.Respond(item);
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      /// Returns a list of substitutions that this product line is responsible
      /// for when it comes to purchase orders (i.e., all ProductLineID info
      /// returned is a pointer to a product line that gets consumed into this
      /// product line's purchase order).
      /// 
      /// URL Parameters:
      ///   product line id
      /// Query Parameters:
      ///   preload:
      ///     productline
      /// Returns:
      ///   <ArrayOfProductLineSubstitution>
      ///     <ProductLineSubstitution>
      ///       <TupleID>21</TupleID>
      ///       <ProductLineID>191</ProductLineID>
      ///       <PurchasingProductLineID>509</PurchasingProductLineID>
      ///     </ProductLineSubstitution>
      ///     ...
      ///   </ArrayOfBaseProductLine>
      Get["/{id:int}/substitutes/from"] = _ =>
      {
        try
        {
          IEnumerable<DbProductLineSubstitution> list = DbProductLineSubstitution.GetByPurchasingProductLineID(_.id, NancyExtensions.GetPreloads(Request.Query));
          return NancyExtensions.Respond(list);
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      #endregion


      #region DbProductLineSubstitution - Create

      /// Define a product line as a second-tier when it comes to purchasing.
      /// Any POs created for items will use the first-tier product line code
      /// to match up the correct PO. E.g., D-W => STA, any items under D-W
      /// will always map to a STA order.
      /// 
      /// URL Parameters:
      ///   product line id
      ///   purchasing product line id
      /// Query Parameters:
      ///   none -- note: this might need to have an explicit content length zero set client side for the c# lib?
      /// Returns:
      ///     <ProductLineSubstitution>
      ///       <TupleID>21</TupleID>
      ///       <ProductLineID>191</ProductLineID>
      ///       <PurchasingProductLineID>509</PurchasingProductLineID>
      ///     </ProductLineSubstitution>
      Post["/{id:int}/substitutes/to/create/{pplid:int}"] = _ =>
      {
        try
        {
          DbProductLineSubstitution item = DbProductLineSubstitution.CreateSubstitution(_.id, _.pplid);
          return NancyExtensions.Respond(item);
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      #endregion


      #region DbProductLineSubstitution - Update

      /// Modify an existing product line subsititution. Really the only thing
      /// to change is the PurchasingProductLineID endpoint. This can have
      /// weird effects if used while there's an open PO.
      /// 
      /// URL Parameters:
      ///   product line id
      /// Query Parameters:
      ///   TP.Object.ProductLineSubstitution => serialized object with new data
      /// Returns:
      ///     <ProductLineSubstitution>
      ///       <TupleID>21</TupleID>
      ///       <ProductLineID>191</ProductLineID>
      ///       <PurchasingProductLineID>509</PurchasingProductLineID>
      ///     </ProductLineSubstitution>
      Post["/{id:int}/substitutes/to/update"] = _ =>
      {
        try
        {
          DbProductLineSubstitution item = NancyExtensions.GetDbObjectParameter<DbProductLineSubstitution>(Request.Form, "TP.Object.ProductLineSubstitution");
          if (item.ProductLineID != _.id)
          {
            throw new ArgumentException("url id does not match given object!");
          }
          // 
          return NancyExtensions.Respond(item.UpdateSubstitution());
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      #endregion


      #region DbProductLineSubstitution - Delete

      /// Deletes the given purchasing subsititution. This will have very bad
      /// effects if used while an open PO exists (item will basically be
      /// orphaned onto that PO and not modifiable)
      /// 
      /// URL Parameters:
      ///   product line id
      /// Query Parameters:
      ///   TP.Object.ProductLineSubstitution => serialized object with current data
      /// Returns:
      ///   serialized SuccessResponse or ErrorResponse
      Post["/{id:int}/substitutes/to/delete/{pplid:int}"] = _ =>
      {
        try
        {
          DbProductLineSubstitution item = NancyExtensions.GetDbObjectParameter<DbProductLineSubstitution>(Request.Form, "TP.Object.ProductLineSubstitution");
          if (item.ProductLineID != _.id)
          {
            throw new ArgumentException("url id does not match given object!");
          }
          if (item.PurchasingProductLineID != _.pplid)
          {
            throw new ArgumentException("url id does not match given object!");
          }
          if (item.DeleteSubstitution())
          {
            return NancyExtensions.Respond(new TP.Base.SuccessResponse("substitution entry deleted"));
          }
          else
          {
            return NancyExtensions.Respond(new TP.Base.ErrorResponse("delete failed"));
          }
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      #endregion


      #region DbProductLineDropshipCharge - Read

      /// Returns a list of dropship charge objects. The AddressZoningType
      /// property is unique among these. Note that is will always load
      /// the corresponding ChargeMatrix sub-table.
      /// 
      /// URL Parameters:
      ///   product line id
      /// Query Parameters:
      ///   none
      /// Returns:
      ///   <ArrayOfProductLineDropshipCharge>
      ///     <ProductLineDropshipCharge>
      ///       <TupleID>13</TupleID>
      ///       <ProductLineID>191</ProductLineID>
      ///       <AddressZoningType>Commercial</AddressZoningType>
      ///       <AdderAmount>5.0000</AdderAmount>
      ///       <AdderCalculationType>Percentage</AdderCalculationType>
      ///       ...
      ///       <ChargeMatrix>
      ///         <ProductLineDropshipMatrix>
      ///           <ProductLineDropshipMatrixLine>
      ///             <TupleID>25</TupleID>
      ///             <ProductLineID>191</ProductLineID>
      ///             <CalculationType>Flat</CalculationType>
      ///             <ThresholdCost>250.0100</ThresholdCost>
      ///             <Amount>15.5000</Amount>
      ///           </ProductLineDropshipMatrixLine>
      ///           ...
      ///         </ProductLineDropshipMatrix>
      ///       </ChargeMatrix>
      ///     </ProductLineContactInfo>
      ///     ...
      ///   </ArrayOfProductLineContactInfo>
      Get["/{id:int}/dropshipcharges"] = _ =>
      {
        try
        {
          IEnumerable<DbProductLineDropshipCharge> list = DbProductLineDropshipCharge.GetByProductLineID(_.id, NancyExtensions.GetPreloads(Request.Query));
          return NancyExtensions.Respond(list);
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      #endregion


      #region DbProductLineDropshipCharge - Create

      /// Define a product line dropship charge for a type of ship-to zoning.
      /// This doesn't work super flexibly, and should probably be rewritten
      /// at some point. Note that a ProductLineDropshipCharge must exist to
      /// create a ProductLineDropshipMatrix, even if there isn't any data in
      /// the charge line. Also, the ProductLineDropshipMatrix is shared
      /// between all ProductLineDropshipCharge tuples.
      /// 
      /// URL Parameters:
      ///   product line id
      ///   (int)AddressZoningType
      /// Query Parameters:
      ///   none -- note: this might need to have an explicit content length zero set client side for the c# lib?
      /// Returns:
      ///     <ProductLineDropshipCharge>
      ///       <TupleID>68</TupleID>
      ///       <ProductLineID>80</ProductLineID>
      ///       <AdderAmount i:nil="true" />
      ///       <AdderCalculationType>Percentage</AdderCalculationType>
      ///       <AdderMaximum i:nil="true" />
      ///       <AdderThreshold i:nil="true" />
      ///       <AddressZoningType>Commercial</AddressZoningType>
      ///       <LiftGateCharge i:nil="true" />
      ///       <LiftGateWeightThreshold i:nil="true" />
      ///     </ProductLineDropshipCharge>
      Post["/{id:int}/dropshipcharges/create/{zone:int}"] = _ =>
      {
        try
        {
          if (!Enum.IsDefined(typeof(TP.Object.AddressZoningType), (int)_.zone))
          {
            throw new ArgumentException("invalid address zoning type id");
          }
          DbProductLineDropshipCharge item = DbProductLineDropshipCharge.CreateChargeLine(
              productLineID: _.id,
              zoning: (TP.Object.AddressZoningType)_.zone
            );
          return NancyExtensions.Respond(item);
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      #endregion


      #region DbProductLineDropshipCharge - Update

      /// Modify an existing product line dropship charge for a specific type
      /// of zoning. Note that the ProductLineDropshipMatrix will probably be
      /// included in the object, but does not get updated. Should it?
      /// 
      /// URL Parameters:
      ///   product line id
      ///   (int)AddressZoningType
      /// Query Parameters:
      ///   TP.Object.ProductLineDropshipCharge => serialized object with new data
      /// Returns:
      ///     <ProductLineDropshipCharge>
      ///       <TupleID>68</TupleID>
      ///       <ProductLineID>80</ProductLineID>
      ///       <AdderAmount>5.0000</AdderAmount>
      ///       <AdderCalculationType>Percentage</AdderCalculationType>
      ///       <AdderMaximum i:nil="true" />
      ///       <AdderThreshold i:nil="true" />
      ///       <AddressZoningType>Commercial</AddressZoningType>
      ///       <LiftGateCharge i:nil="true" />
      ///       <LiftGateWeightThreshold i:nil="true" />
      ///     </ProductLineDropshipCharge>
      Post["/{id:int}/dropshipcharges/update/{zone:int}"] = _ =>
      {
        try
        {
          if (!Enum.IsDefined(typeof(TP.Object.AddressZoningType), (int)_.zone))
          {
            throw new ArgumentException("invalid address zoning type id");
          }
          DbProductLineDropshipCharge item = NancyExtensions.GetDbObjectParameter<DbProductLineDropshipCharge>(Request.Form, "TP.Object.ProductLineDropshipCharge");
          if (item.ProductLineID != _.id)
          {
            throw new ArgumentException("url id does not match given object!");
          }
          if (item.AddressZoningType != (TP.Object.AddressZoningType)_.zone)
          {
            throw new ArgumentException("url id does not match given object!");
          }
          return NancyExtensions.Respond(item.UpdateDropshipCharges());
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      #endregion


      #region DbProductLineDropshipMatrix - Create/Update/Delete (Overwrite)

      /// Overwrite the entire product line's dropship fee matrix with the new
      /// data. Threshold/Cost values not supplied are assumed to be deleted.
      /// Note that thresholds must be increasing. No validation on costs. The
      /// last row given should probably have a ridiculously high threshold
      /// (INT_MAX probably), otherwise there is an implied no-charge level at
      /// the top.
      /// 
      /// It kind of sucks that the <ProductLineDropshipMatrix> tag gets duped
      /// because of the list encapsulation thing.
      /// 
      /// URL Parameters:
      ///   product line id
      /// Query Parameters:
      ///   (int)CalculationType
      ///   Threshold1
      ///   Cost1
      ///   Threshold2
      ///   Cost2
      ///   Threshold3
      ///   Cost3
      ///   Threshold4
      ///   Cost4
      ///   Threshold5
      ///   Cost5
      /// Returns:
      ///     <ProductLineDropshipMatrix>
      ///       <ProductLineDropshipMatrix>
      ///         <ProductLineDropshipMatrixLine><Amount>13.5000</Amount>
      ///           <CalculationType>Percentage</CalculationType>
      ///           <ProductLineID>80</ProductLineID>
      ///           <ThresholdCost>49.9900</ThresholdCost>
      ///           <TupleID>147</TupleID>
      ///         </ProductLineDropshipMatrixLine>
      ///         <ProductLineDropshipMatrixLine>
      ///           <Amount>26.7500</Amount>
      ///           <CalculationType>Percentage</CalculationType>
      ///           <ProductLineID>80</ProductLineID>
      ///           <ThresholdCost>99.9900</ThresholdCost>
      ///           <TupleID>148</TupleID>
      ///         </ProductLineDropshipMatrixLine>
      ///       </ProductLineDropshipMatrix>
      ///     </ProductLineDropshipMatrix>
      Post["/{id:int}/dropshipcharges/matrix/overwrite"] = _ =>
      {
        try
        {
          int calcTypeID = NancyExtensions.GetBasicParameter<int>(Request.Form, "CalculationType", -1);
          if (!Enum.IsDefined(typeof(TP.Object.DropshipChargeCalculationType), calcTypeID))
          {
            throw new ArgumentException("invalid dropship charge id");
          }
          List<Tuple<decimal, decimal>> values = NancyExtensions.GetPairedValueList(Request.Form, "Threshold", "Cost", Decimal.MinValue, Decimal.MinValue);
          DbProductLineDropshipMatrix item = DbProductLineDropshipMatrix.OverwriteMatrix(
              productLineID: _.id,
              calcType: (TP.Object.DropshipChargeCalculationType)calcTypeID,
              values: values
            );
          return NancyExtensions.Respond(item);
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      #endregion


      #region DbProductLineCostField - Read

      /// Returns a list of the current cost fields for a product line.
      /// 
      /// URL Parameters:
      ///   product line id
      /// Query Parameters:
      ///   none
      /// Returns:
      ///   <ProductLineCostField>
      ///     <TupleID>1</TupleID>
      ///     <ProductLineID>1</ProductLineID>
      ///     <Name>Platinum</Name>
      ///     <Description>platinum pricing tier, see mfg price folder</Description>
      ///   </ProductLineDropshipCharge>
      Get["/{id:int}/costfields/active"] = _ =>
      {
        try
        {
          IEnumerable<DbProductLineCostField> list = DbProductLineCostField.GetActiveByProductLineID(_.id, NancyExtensions.GetPreloads(Request.Query));
          return NancyExtensions.Respond(list);
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      /// Returns a list of all cost fields for a product line, including ones
      /// that were previously deleted.
      /// 
      /// URL Parameters:
      ///   product line id
      /// Query Parameters:
      ///   none
      /// Returns:
      ///   <ProductLineCostField>
      ///     <TupleID>1</TupleID>
      ///     <ProductLineID>1</ProductLineID>
      ///     <Name>Platinum</Name>
      ///     <Description>platinum pricing tier, see mfg price folder</Description>
      ///   </ProductLineDropshipCharge>
      Get["/{id:int}/costfields/all"] = _ =>
      {
        try
        {
          IEnumerable<DbProductLineCostField> list = DbProductLineCostField.GetByProductLineID(_.id, NancyExtensions.GetPreloads(Request.Query));
          return NancyExtensions.Respond(list);
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      #endregion


      #region DbProductLineCostField - Create

      /// Define a product line cost field, this will show up for all items in
      /// the product line. Very little validation on this, since I'm not 100%
      /// sure how it will be used. XXX: in the future, we might want to add
      /// some attributes to this, maybe what to display on hover-over, ways
      /// to use it in calculations, etc.
      /// 
      /// URL Parameters:
      ///   product line id
      ///   field name/type
      /// Query Parameters:
      ///   Description
      /// Returns:
      ///   <ProductLineCostField>
      ///     <TupleID>1</TupleID>
      ///     <ProductLineID>1</ProductLineID>
      ///     <Name>Platinum</Name>
      ///     <Description>platinum pricing tier, see mfg price folder</Description>
      ///   </ProductLineDropshipCharge>
      Post["/{id:int}/costfields/create/{type}"] = _ =>
      {
        try
        {
          DbProductLineCostField item = DbProductLineCostField.CreateNewCostField(
              productLineID: _.id,
              name: _.type,
              desc: NancyExtensions.GetBasicParameter<string>(Request.Form, "Description", string.Empty)
            );
          return NancyExtensions.Respond(item);
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      #endregion


      #region DbProductLineCostField - Update

      /// Modify an existing product line cost field. Just the description so
      /// far, but this will probably increase.
      /// 
      /// URL Parameters:
      ///   product line id
      ///   field name/type
      /// Query Parameters:
      ///   TP.Object.ProductLineCostField => serialized object with new data
      /// Returns:
      ///   <ProductLineCostField>
      ///     <TupleID>1</TupleID>
      ///     <ProductLineID>1</ProductLineID>
      ///     <Name>Platinum</Name>
      ///     <Description>platinum pricing tier, see mfg price folder</Description>
      ///   </ProductLineDropshipCharge>
      Post["/{id:int}/costfields/update/{type}"] = _ =>
      {
        try
        {
          DbProductLineCostField item = NancyExtensions.GetDbObjectParameter<DbProductLineCostField>(Request.Form, "TP.Object.ProductLineCostField");
          if (item.ProductLineID != _.id)
          {
            throw new ArgumentException("url id does not match given object!");
          }
          if (item.Name != _.type)
          {
            throw new ArgumentException("url id does not match given object!");
          }
          return NancyExtensions.Respond(item.UpdateCostField());
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      #endregion

    }
  }

  // SalesRepCRUD handles:
  //   OBJECT                           READ  CREATE  UPDATE  DELETE
  //   DbSalesRep                         Y     Y       Y       N
  //   DbProductLineSalesRep              Y     Y       N       N
  public class SalesRepCRUD : NancyModule
  {
    public SalesRepCRUD() : base("/salesrep")
    {
      #region DbSalesRep - Read

      /// Returns a specific sales rep by id, probably from a product line's
      /// list of sales reps.
      /// 
      /// URL Parameters:
      ///   sales rep id
      /// Query Parameters:
      ///   preload:
      ///     productlines
      ///     reps
      /// Returns:
      ///   <SalesRep>
      ///     <SalesRepID>95</SalesRepID>
      ///     <FirstName>Frank</FirstName>
      ///     <LastName>Bednarz</LastName>
      ///     ...
      ///   </SalesRep>
      Get["/{id:int}"] = _ =>
      {
        try
        {
          DbSalesRep item = DbSalesRep.GetByID(_.id, NancyExtensions.GetPreloads(Request.Query));
          return NancyExtensions.Respond(item);
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      #endregion


      #region DbSalesRep - Create

      /// Creates a new sales rep stub. Not connected to any product lines,
      /// that needs to be done separately.
      /// 
      /// URL Parameters:
      ///   none
      /// Query Parameters:
      ///   FirstName
      ///   LastName
      /// Returns:
      ///   <SalesRep>
      ///     <SalesRepID>95</SalesRepID>
      ///     <FirstName>Frank</FirstName>
      ///     <LastName>Bednarz</LastName>
      ///     ...
      ///   </SalesRep>
      Post["/create"] = _ =>
      {
        try
        {
          DbSalesRep item = DbSalesRep.CreateSalesRep(
              firstName: NancyExtensions.GetBasicParameter<string>(Request.Form, "FirstName"),
              lastName: NancyExtensions.GetBasicParameter<string>(Request.Form, "LastName")
            );
          return NancyExtensions.Respond(item);
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      #endregion


      #region DbSalesRep - Update

      /// Modifies an existing sales rep, given a new serialized version of
      /// the sales rep object. Returns the updated, current version of the
      /// object in the database.
      /// 
      /// URL Parameters:
      ///   sales rep id
      /// Query Parameters:
      ///   TP.Object.SalesRep => serialized object with new data
      /// Returns:
      ///   <SalesRep>
      ///     <SalesRepID>95</SalesRepID>
      ///     <FirstName>Frank</FirstName>
      ///     <LastName>Bednarz</LastName>
      ///     ...
      ///   </SalesRep>
      Post["/update/{id:int}"] = _ =>
      {
        try
        {
          DbSalesRep item = NancyExtensions.GetDbObjectParameter<DbSalesRep>(Request.Form, "TP.Object.SalesRep");
          if (item.SalesRepID != _.id)
          {
            throw new ArgumentException("url id does not match given object!");
          }
          return NancyExtensions.Respond(item.UpdateSalesRep());
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      #endregion


      #region DbProductLineSalesRep - Read

      /// Returns a list of product line / sales rep many-to-many link objects
      /// for all the lines that this rep is associated with.
      /// 
      /// URL Parameters:
      ///   sales rep id
      /// Query Parameters:
      ///   preload:
      ///     productlines
      /// Returns:
      ///   <ArrayOfProductLineSalesRep>
      ///     <ProductLineSalesRep>
      ///       <TupleID>141</TupleID>
      ///       <ProductLineID>191</ProductLineID>
      ///       <SalesRepID>40</SalesRepID>
      ///       <IsPrimary>false</IsPrimary>
      ///     </ProductLineSalesRep>
      ///     ...
      ///   </ProductLineSalesRep>
      Get["/{id:int}/lines"] = _ =>
      {
        try
        {
          IEnumerable<DbProductLineSalesRep> list = DbProductLineSalesRep.GetBySalesRepID(_.id, NancyExtensions.GetPreloads(Request.Query));
          return NancyExtensions.Respond(list);
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      #endregion

      
      #region DbProductLineSalesRep - Create

      /// Attaches a sales rep to a product line. Always defaults to creating
      /// as a non-primary, update required to change this.
      /// 
      /// URL Parameters:
      ///   sales rep id
      ///   product line id
      /// Query Parameters:
      ///   none -- note: this might need to have an explicit content length zero set client side for the c# lib?
      /// Returns:
      ///     <ProductLineSalesRep>
      ///       <TupleID>141</TupleID>
      ///       <ProductLineID>191</ProductLineID>
      ///       <SalesRepID>40</SalesRepID>
      ///       <IsPrimary>false</IsPrimary>
      ///     </ProductLineSalesRep>
      Post["/{id:int}/lines/add/{plid:int}"] = _ =>
      {
        try
        {
          DbProductLineSalesRep item = DbProductLineSalesRep.CreateLinkage(
              productLineID: _.plid,
              salesRepID: _.id,
              isPrimary: false
            );
          return NancyExtensions.Respond(item);
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      #endregion


      #region DbProductLineSalesRep - Update

      /// Modifies the sales rep / product line link fields.
      /// 
      /// URL Parameters:
      ///   sales rep id
      ///   product line id
      /// Query Parameters:
      ///   TP.Object.ProductLineSalesRep => serialized object with new data
      /// Returns:
      ///     <ProductLineSalesRep>
      ///       <TupleID>141</TupleID>
      ///       <ProductLineID>191</ProductLineID>
      ///       <SalesRepID>40</SalesRepID>
      ///       <IsPrimary>false</IsPrimary>
      ///     </ProductLineSalesRep>
      Post["/{id:int}/lines/update/{plid:int}"] = _ =>
      {
        try
        {
          DbProductLineSalesRep item = NancyExtensions.GetDbObjectParameter<DbProductLineSalesRep>(Request.Form, "TP.Object.ProductLineSalesRep");
          if (item.SalesRepID != _.id)
          {
            throw new ArgumentException("url id does not match given object!");
          }
          if (item.ProductLineID != _.plid)
          {
            throw new ArgumentException("url id does not match given object!");
          }
          return NancyExtensions.Respond(item.UpdateLinkage());
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      #endregion


      #region DbProductLineSalesRep - Delete

      /// Deletes the given sales rep / product line link. Delete will only be
      /// successful if the object has not been modified from the current
      /// version in the database.
      /// 
      /// URL Parameters:
      ///   sales rep id
      ///   product line id
      /// Query Parameters:
      ///   TP.Object.ProductLineSalesRep => serialized object with current data
      /// Returns:
      ///   serialized SuccessResponse or ErrorResponse
      Post["/{id:int}/lines/delete/{plid:int}"] = _ =>
      {
        try
        {
          DbProductLineSalesRep item = NancyExtensions.GetDbObjectParameter<DbProductLineSalesRep>(Request.Form, "TP.Object.ProductLineSalesRep");
          if (item.SalesRepID != _.srid)
          {
            throw new ArgumentException("url id does not match given object!");
          }
          if (item.ProductLineID != _.id)
          {
            throw new ArgumentException("url id does not match given object!");
          }
          if (item.DeleteLinkage())
          {
            return NancyExtensions.Respond(new TP.Base.SuccessResponse("line/rep entry deleted"));
          }
          else
          {
            return NancyExtensions.Respond(new TP.Base.ErrorResponse("delete failed"));
          }
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      #endregion
    }
  }

  // PersonCRUD handles:
  //   OBJECT                           READ  CREATE  UPDATE  DELETE
  //   DbPerson                           Y     Y       Y       N
  public class PersonCRUD : NancyModule
  {
    public PersonCRUD() : base("/person")
    {
      #region DbPerson - Read

      /// Returns a specific person by id
      /// 
      /// URL Parameters:
      ///   person id
      /// Query Parameters:
      ///   preload:
      ///     TODO: DbPerson preloads
      /// Returns:
      ///   <Person>
      ///     <PersonID>2</PersonID>
      ///     <NTUsername>briandonorfio</NTUsername>
      ///     <ShortName>Brian</ShortName>
      ///     ...
      ///   </Person>
      Get["/{id:int}"] = _ =>
      {
        try
        {
          DbPerson item = DbPerson.GetByID(_.id, NancyExtensions.GetPreloads(Request.Query));
          return NancyExtensions.Respond(item);
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      /// Returns a specific person by their windows username (all lowercase,
      /// but this gets checked).
      /// 
      /// URL Parameters:
      ///   person username
      /// Query Parameters:
      ///   preload:
      ///     TODO: DbPerson preloads
      /// Returns:
      ///   <Person>
      ///     <PersonID>2</PersonID>
      ///     <NTUsername>briandonorfio</NTUsername>
      ///     <ShortName>Brian</ShortName>
      ///     ...
      ///   </Person>
      Get["/{ntuser}"] = _ =>
      {
        try
        {
          DbPerson item = DbPerson.GetByNTUsername(_.ntuser, NancyExtensions.GetPreloads(Request.Query));
          return NancyExtensions.Respond(item);
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      /// Returns all users in the database. Possibly expand this to options
      /// for non-deleted or something?
      /// 
      /// URL Parameters:
      ///   none
      /// Query Parameters:
      ///   preload:
      ///     TODO: DbPerson preloads
      /// Returns:
      ///   <ArrayOfPerson>
      ///     <Person>
      ///       <PersonID>2</PersonID>
      ///       <NTUsername>briandonorfio</NTUsername>
      ///       <ShortName>Brian</ShortName>
      ///       ...
      ///     </Person>
      ///     ...
      ///   </ArrayOfPerson>
      Get["/list/"] = _ =>
      {
        try
        {
          IEnumerable<DbPerson> list = DbPerson.GetAll(NancyExtensions.GetPreloads(Request.Query));
          return NancyExtensions.Respond(list);
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      #endregion


      #region DbPerson - Create

      /// Creates a new user stub. Permissions, etc. need to be entered.
      /// 
      /// URL Parameters:
      ///   none
      /// Query Parameters:
      ///   NTusername
      ///   ShortName
      /// Returns:
      ///   <Person>
      ///     <PersonID>2</PersonID>
      ///     <NTUsername>briandonorfio</NTUsername>
      ///     <ShortName>Brian</ShortName>
      ///     ...
      ///   </Person>
      Post["/create"] = _ =>
      {
        try
        {
          DbPerson item = DbPerson.CreatePerson(
              ntUsername: NancyExtensions.GetBasicParameter<string>(Request.Form, "NTUsername"),
              shortName: NancyExtensions.GetBasicParameter<string>(Request.Form, "ShortName")
            );
          return NancyExtensions.Respond(item);
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      #endregion


      #region DbPerson - Update

      /// Modify an existing user. Really just deleted flag?
      /// 
      /// URL Parameters:
      ///   person id
      /// Query Parameters:
      ///   TP.Object.Person => serialized object with new data
      /// Returns:
      ///   <Person>
      ///     <PersonID>2</PersonID>
      ///     <NTUsername>briandonorfio</NTUsername>
      ///     <ShortName>Brian</ShortName>
      ///     ...
      ///   </Person>
      Post["/update/{id:int}"] = _ =>
      {
        try
        {
          DbPerson item = NancyExtensions.GetDbObjectParameter<DbPerson>(Request.Form, "TP.Object.Person");
          if (item.PersonID != _.id)
          {
            throw new ArgumentException("url id does not match given object!");
          }
          return NancyExtensions.Respond(item.UpdatePerson());
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      #endregion
    }
  }

  // ComputerCRUD handles:
  //   OBJECT                           READ  CREATE  UPDATE  DELETE
  //   DbComputer                         Y     Y       Y       N
  public class ComputerCRUD : NancyModule
  {
    public ComputerCRUD() : base("/computer")
    {
      #region DbComputer - Read

      /// Returns a specific computer by id
      /// 
      /// URL Parameters:
      ///   computer id
      /// Query Parameters:
      ///   none
      /// Returns:
      ///   <Computer>
      ///     <ComputerID>2</ComputerID>
      ///     <ComputerName>briandonorfio</ComputerName>
      ///   </Computer>
      Get["/{id:int}"] = _ =>
      {
        try
        {
          DbComputer item = DbComputer.GetByID(_.id, NancyExtensions.GetPreloads(Request.Query));
          return NancyExtensions.Respond(item);
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      /// Returns a specific computer by its windows (netbios) name
      /// 
      /// URL Parameters:
      ///   computer name
      /// Query Parameters:
      ///   none
      /// Returns:
      ///   <Computer>
      ///     <ComputerID>2</ComputerID>
      ///     <ComputerName>briandonorfio</ComputerName>
      ///   </Computer>
      Get["/{name}"] = _ =>
      {
        try
        {
          DbComputer item = DbComputer.GetByComputerName(_.name, NancyExtensions.GetPreloads(Request.Query));
          return NancyExtensions.Respond(item);
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      /// Returns all computers in the database.
      /// 
      /// URL Parameters:
      ///   none
      /// Query Parameters:
      ///   none
      /// Returns:
      ///   <ArrayOfComputer>
      ///     <Computer>
      ///       <ComputerID>2</ComputerID>
      ///       <ComputerName>briandonorfio</ComputerName>
      ///     </Computer>
      ///     ...
      ///   </ArrayOfComputer>
      Get["/list/"] = _ =>
      {
        try
        {
          IEnumerable<DbComputer> list = DbComputer.GetAll(NancyExtensions.GetPreloads(Request.Query));
          return NancyExtensions.Respond(list);
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      #endregion


      #region DbComputer - Create

      /// Creates a new computer stub.
      /// 
      /// URL Parameters:
      ///   none
      /// Query Parameters:
      ///   ComputerName
      /// Returns:
      ///   <Computer>
      ///     <ComputerID>2</ComputerID>
      ///     <ComputerName>briandonorfio</ComputerName>
      ///   </Computer>
      Post["/create"] = _ =>
      {
        try
        {
          DbComputer item = DbComputer.CreateComputer(
              computerName: NancyExtensions.GetBasicParameter<string>(Request.Form, "ComputerName")
            );
          return NancyExtensions.Respond(item);
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      #endregion
    }
  }

  // ItemCRUD handles:
  //   OBJECT                           READ  CREATE  UPDATE  DELETE
  //   DbItem                             Y     Y       Y       N
  //   DbItemFreightEstimate              Y     N       N       N
  //   DbItemFreightActual                Y     N       N       N
  //   DbItemQuantity                     Y     N       N       N
  //   DbItemCost                         Y     Y       N       N
  //   DbItemPrice                        Y     Y       N       N
  public class ItemCRUD : NancyModule
  {
    public ItemCRUD() : base("/item")
    {
      #region DbItem - Read

      /// Returns an item object, based on the item id.
      /// 
      /// URL Parameters:
      ///   item id
      /// Query Parameters:
      ///   preload:
      ///     TODO
      /// Returns:
      ///   <Item>
      ///     <ItemID>9565</ItemID>
      ///     <ItemNumber>DEA28-127</ItemNumber>
      ///     <ItemDescription>111 IN. X 1/2 IN. X 3 TPI TIMB</ItemDescription>
      ///     <AdditionalInfo>IMPORTED FROM JOE Bs SPREADSHEET 1/11/2012</AdditionalInfo>
      ///     ...
      ///   </Item>
      Get["/{id:int}"] = _ =>
      {
        try
        {
          DbItem item = DbItem.GetByID(_.id, NancyExtensions.GetPreloads(Request.Query));
          return NancyExtensions.Respond(item);
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      /// Returns an item object, based on the item number.
      /// 
      /// URL Parameters:
      ///   item id
      /// Query Parameters:
      ///   preload:
      ///     TODO
      /// Returns:
      ///   <Item>
      ///     <ItemID>9565</ItemID>
      ///     <ItemNumber>DEA28-127</ItemNumber>
      ///     <ItemDescription>111 IN. X 1/2 IN. X 3 TPI TIMB</ItemDescription>
      ///     <AdditionalInfo>IMPORTED FROM JOE Bs SPREADSHEET 1/11/2012</AdditionalInfo>
      ///     ...
      ///   </Item>
      Get["/{item}"] = _ =>
      {
        try
        {
          DbItem item = DbItem.GetByItemNumber(_.item, NancyExtensions.GetPreloads(Request.Query));
          return NancyExtensions.Respond(item);
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      /// Returns a list of BaseItem objects, based on the item number
      /// prefix (generally product line).
      /// 
      /// URL Parameters:
      ///   prefix
      /// Query Parameters:
      ///   none
      /// Returns:
      ///   <ArrayOfBaseItem>
      ///     <BaseItem>
      ///       <ItemID>1703</ItemID>
      ///       <ItemNumber>SAI00501</ItemNumber>
      ///     </BaseItem>
      ///     ...
      ///   </ArrayOfBaseItem>
      Get["/list/{prefix:length(3,6)}"] = _ =>
      {
        try
        {
          IEnumerable<DbBaseItem> list = DbItem.ListByPrefix(_.prefix);
          return NancyExtensions.Respond(list);
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      #endregion


      #region DbItem - Create

      /// Creates a new item in the database, including any other default
      /// required sub-objects (e.g., quantity in 000 whse, component map).
      /// 
      /// URL Parameters:
      ///   none
      /// Query Parameters:
      ///   ItemNumber
      ///   ItemDescription
      /// Returns:
      ///   <Item>
      ///     <ItemID>9565</ItemID>
      ///     <ItemNumber>DEA28-127</ItemNumber>
      ///     <ItemDescription>111 IN. X 1/2 IN. X 3 TPI TIMB</ItemDescription>
      ///     <AdditionalInfo>IMPORTED FROM JOE Bs SPREADSHEET 1/11/2012</AdditionalInfo>
      ///     ...
      ///   </Item>
      Post["/create"] = _ =>
      {
        try
        {
          DbItem item = DbItem.CreateItem(
              itemNumber: NancyExtensions.GetBasicParameter<string>(Request.Form, "ItemNumber"),
              itemDescription: NancyExtensions.GetBasicParameter<string>(Request.Form, "ItemDescription")
            );
          return NancyExtensions.Respond(item);
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      #endregion


      #region DbItem - Update

      /// Edits field(s) for a given item, note that edits do not propogate down
      /// to preloaded sub-objects.
      /// 
      /// URL Parameters:
      ///   item id
      /// Query Parameters:
      ///   TP.Object.Item => serialized object with new data
      /// Returns:
      ///   <Item>
      ///     <ItemID>9565</ItemID>
      ///     <ItemNumber>DEA28-127</ItemNumber>
      ///     <ItemDescription>111 IN. X 1/2 IN. X 3 TPI TIMB</ItemDescription>
      ///     <AdditionalInfo>IMPORTED FROM JOE Bs SPREADSHEET 1/11/2012</AdditionalInfo>
      ///     ...
      ///   </Item>
      Post["/update/{id:int}"] = _ =>
      {
        try
        {
          DbItem item = NancyExtensions.GetDbObjectParameter<DbItem>(Request.Form, "TP.Object.Item");
          if (item.ItemID != _.id)
          {
            throw new ArgumentException("url id does not match given object!");
          }
          return NancyExtensions.Respond(item.UpdateItem());
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      #endregion


      #region DbItemFreightEstimate - Read

      /// Returns all freight estimates for an item.
      /// 
      /// URL Parameters:
      ///   item id
      /// Query Parameters:
      ///   none
      /// Returns:
      ///   <ArrayOfItemFreightEstimate>
      ///     <ItemFreightEstimate>
      ///       <TupleID>11656</TupleID>
      ///       <ItemID>6585</ItemID>
      ///       <PostalCode>89005</PostalCode>
      ///       <Cost>11.4100</Cost>
      ///       <Service>U</Service>
      ///       <LastUpdated>2014-03-17T15:50:00</LastUpdated>
      ///     </ItemFreightEstimate>
      ///     ...
      ///   </ArrayOfItemFreightEstimate>
      Get["/{id:int}/freight/estimate"] = _ =>
      {
        try
        {
          IEnumerable<DbItemFreightEstimate> list = DbItemFreightEstimate.GetByItemID(_.id, NancyExtensions.GetPreloads(Request.Query));
          return NancyExtensions.Respond(list);
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      #endregion


      #region DbItemFreightEstimate - Delete

      /// Deletes the given freight estimate.
      /// 
      /// URL Parameters:
      ///   item id
      ///   freight estimate tuple id
      /// Query Parameters:
      ///   TP.Object.ItemFreightEstimate => serialized object with current data
      /// Returns:
      ///   serialized SuccessResponse or ErrorResponse
      Post["/{id:int}/freight/estimate/delete/{fid:int}"] = _ =>
      {
        try
        {
          DbItemFreightEstimate item = NancyExtensions.GetDbObjectParameter<DbItemFreightEstimate>(Request.Form, "TP.Object.ItemFreightEstimate");
          if (item.ItemID != _.id)
          {
            throw new ArgumentException("url id does not match given object!");
          }
          if (item.TupleID != _.fid)
          {
            throw new ArgumentException("url id does not match given object!");
          }
          if (item.DeleteFreightEstimate())
          {
            return NancyExtensions.Respond(new TP.Base.SuccessResponse("freight estimate deleted"));
          }
          else
          {
            return NancyExtensions.Respond(new TP.Base.ErrorResponse("delete failed"));
          }
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      #endregion


      #region DbItemFreightActual - Read

      /// Returns all actual freight values for this item. We might want this
      /// to change to "within the past X months/years", if there ends up
      /// being too much.
      /// 
      /// URL Parameters:
      ///   item id
      /// Query Parameters:
      ///   none
      /// Returns:
      ///   <ArrayOfItemFreightActual>
      ///     <ItemFreightActual>
      ///       <TupleID>253393</TupleID>
      ///       <ItemID>9987</ItemID>
      ///       <PostalCode>60031</PostalCode>
      ///       <Cost>37.9100</Cost>
      ///       ...
      ///     </ItemFreightActual>
      ///     ...
      ///   </ArrayOfItemFreightActual>
      Get["/{id:int}/freight/actual"] = _ =>
      {
        try
        {
          IEnumerable<DbItemFreightActual> list = DbItemFreightActual.GetByItemID(_.id, NancyExtensions.GetPreloads(Request.Query));
          return NancyExtensions.Respond(list);
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      #endregion


      #region DbItemQuantity - Read

      /// Returns quantity information for an item's Mas warehouses.
      /// 
      /// URL Parameters:
      ///   item id
      /// Query Parameters:
      ///   none
      /// Returns:
      ///   <ArrayOfItemQuantity>
      ///     <ItemQuantity>
      ///       <ItemID>6585</ItemID>
      ///       <OnHand>2</OnHand>
      ///       <WarehouseCode>000</WarehouseCode>
      ///       <OnPurchaseOrder>0</OnPurchaseOrder>
      ///       <OnSalesOrder>1</OnSalesOrder>
      ///       <OnBackOrder>0</OnBackOrder>
      ///     </ItemQuantity>
      ///     ...
      ///   </ArrayOfItemQuantity>
      Get["/{id:int}/quantity"] = _ =>
      {
        try
        {
          IEnumerable<DbItemQuantity> list = DbItemQuantity.GetByItemID(_.id, NancyExtensions.GetPreloads(Request.Query));
          return NancyExtensions.Respond(list);
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      #endregion


      #region DbItemCost - Read

      /// Returns current cost information for an item. Note that this will
      /// return all cost types that exist, even if they've been deleted or
      /// whatever in the product line. Not sure if we want to change this
      /// functionality or what.
      /// 
      /// URL Parameters:
      ///   item id
      /// Query Parameters:
      ///   none
      /// Returns:
      ///   <ArrayOfItemCost>
      ///     <ItemCost>
      ///       <TupleID>51243</TupleID>
      ///       <ItemID>6585</ItemID>
      ///       <CostType>Std Cost</CostType>
      ///       <CostValue>25.0000</CostValue>
      ///       <TimeEffective>2009-08-26T11:40:21.983</TimeEffective>
      ///       <PersonID>6</PersonID>
      ///       <ApplicationModuleID>PoinvInventoryMaintenance</ApplicationModuleID>
      ///     </ItemCost>
      ///     ...
      ///   </ArrayOfItemCost>
      Get["/{id:int}/cost/"] = _ =>
      {
        try
        {
          IEnumerable<DbItemCost> list = DbItemCost.GetCurrentByItemID(_.id, NancyExtensions.GetPreloads(Request.Query));
          return NancyExtensions.Respond(list);
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      /// Returns the current specified cost for the given item. Current cost
      /// types are "Std Cost", "New Cost", "List Price", and "MAPP". New
      /// types can be added per productline, see DbProductLineCostField.
      /// 
      /// URL Parameters:
      ///   item id
      ///   cost type string
      /// Query Parameters:
      ///   none
      /// Returns:
      ///   <ItemCost>
      ///     <TupleID>51243</TupleID>
      ///     <ItemID>6585</ItemID>
      ///     <CostType>Std Cost</CostType>
      ///     <CostValue>25.0000</CostValue>
      ///     <TimeEffective>2009-08-26T11:40:21.983</TimeEffective>
      ///     <PersonID>6</PersonID>
      ///     <ApplicationModuleID>PoinvInventoryMaintenance</ApplicationModuleID>
      ///   </ItemCost>
      Get["/{id:int}/cost/{type}"] = _ =>
      {
        try
        {
          DbItemCost item = DbItemCost.GetSpecifiedCurrentByItemID(_.type, _.id, NancyExtensions.GetPreloads(Request.Query));
          return NancyExtensions.Respond(item);
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      /// Returns cost history for the given item. Optionally, the query can
      /// be limited by either date or record count with query parameters,
      /// defaulting to all if the limit is not set. Note that the first
      /// records for each type will be the current values.
      /// 
      /// URL Parameters:
      ///   item id
      /// Query Parameters:
      ///   DateLimit - in a parsable format like YYYY-MM-DD
      ///   RecordLimit
      /// Returns:
      ///   <ArrayOfItemCost>
      ///     <ItemCost>
      ///       <TupleID>51243</TupleID>
      ///       <ItemID>6585</ItemID>
      ///       <CostType>Std Cost</CostType>
      ///       <CostValue>25.0000</CostValue>
      ///       <TimeEffective>2009-08-26T11:40:21.983</TimeEffective>
      ///       <PersonID>6</PersonID>
      ///       <ApplicationModuleID>PoinvInventoryMaintenance</ApplicationModuleID>
      ///     </ItemCost>
      ///     ...
      ///   </ArrayOfItemCost>
      Get["/{id:int}/cost/history"] = _ =>
      {
        try
        {
          IEnumerable<DbItemCost> list = DbItemCost.GetHistoryByItemID(
              id: _.id,
              dateLimit: NancyExtensions.GetBasicParameter<DateTime>(Request.Query, "DateLimit", new DateTime(2000, 1, 1)),
              recordLimit: NancyExtensions.GetBasicParameter<int>(Request.Query, "RecordLimit", -1)
            );
          return NancyExtensions.Respond(list);
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      /// Returns cost history for the given item, but only for the specifed
      /// type of cost. Optionally, the query can be limited by either date or
      /// record count with query parameters, defaulting to all if the limit
      /// is not set. Note that the first record is the current value, so to
      /// get the previous value, use RecordLimit=2 and select the 2nd.
      /// 
      /// URL Parameters:
      ///   item id
      /// Query Parameters:
      ///   DateLimit - in a parsable format like YYYY-MM-DD
      ///   RecordLimit
      /// Returns:
      ///   <ArrayOfItemCost>
      ///     <ItemCost>
      ///       <TupleID>51243</TupleID>
      ///       <ItemID>6585</ItemID>
      ///       <CostType>Std Cost</CostType>
      ///       <CostValue>25.0000</CostValue>
      ///       <TimeEffective>2009-08-26T11:40:21.983</TimeEffective>
      ///       <PersonID>6</PersonID>
      ///       <ApplicationModuleID>PoinvInventoryMaintenance</ApplicationModuleID>
      ///     </ItemCost>
      ///     ...
      ///   </ArrayOfItemCost>
      Get["/{id:int}/cost/history/{type}"] = _ =>
      {
        try
        {
          IEnumerable<DbItemCost> list = DbItemCost.GetSpecifiedHistoryByItemID(
              type: _.type,
              id: _.id,
              dateLimit: NancyExtensions.GetBasicParameter<DateTime>(Request.Query, "DateLimit", new DateTime(2000,1,1)),
              recordLimit: NancyExtensions.GetBasicParameter<int>(Request.Query, "RecordLimit", -1)
            );
          return NancyExtensions.Respond(list);
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      /// Returns the previous specified cost for the given item. Really just
      /// a convenience method, so you don't have to mess around with history
      /// and RecordLimit stuff.
      /// 
      /// URL Parameters:
      ///   item id
      ///   cost type string
      /// Query Parameters:
      ///   none
      /// Returns:
      ///   <ItemCost>
      ///     <TupleID>51243</TupleID>
      ///     <ItemID>6585</ItemID>
      ///     <CostType>Std Cost</CostType>
      ///     <CostValue>25.0000</CostValue>
      ///     <TimeEffective>2009-08-26T11:40:21.983</TimeEffective>
      ///     <PersonID>6</PersonID>
      ///     <ApplicationModuleID>PoinvInventoryMaintenance</ApplicationModuleID>
      ///   </ItemCost>
      Get["/{id:int}/cost/history/{type}/previous"] = _ =>
      {
        try
        {
          IEnumerable<DbItemCost> list = DbItemCost.GetSpecifiedHistoryByItemID(
              type: _.type,
              id: _.id,
              dateLimit: NancyExtensions.GetBasicParameter<DateTime>(Request.Query, "DateLimit", new DateTime(2000, 1, 1)),
              recordLimit: 2
            );
          DbItemCost ret = list.Count() == 2 ? list.Last() : null;
          return NancyExtensions.Respond(ret);
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      #endregion


      #region DbItemCost - Create

      /// Updates/creates the cost for an item. Note that this doesn't check
      /// whether the cost type is actually valid (listed as a cost with an
      /// associated ProductLineCostField), it just takes anything.
      /// 
      /// URL Parameters:
      ///   item id
      ///   cost type - string like "Std Cost", "MAPP", etc
      /// Query Parameters:
      ///   CostValue
      ///   PersonID
      ///   ApplicationModuleID
      /// Returns:
      ///   <ItemCost>
      ///     <TupleID>51243</TupleID>
      ///     <ItemID>6585</ItemID>
      ///     <CostType>Std Cost</CostType>
      ///     <CostValue>25.0000</CostValue>
      ///     <TimeEffective>2009-08-26T11:40:21.983</TimeEffective>
      ///     <PersonID>6</PersonID>
      ///     <ApplicationModuleID>PoinvInventoryMaintenance</ApplicationModuleID>
      ///   </ItemCost>
      Post["/{id:int}/cost/{type}/update"] = _ =>
      {
        try
        {
          int app = NancyExtensions.GetBasicParameter<int>(Request.Form, "ApplicationModule");
          if (!Enum.IsDefined(typeof(TP.Object.ApplicationModule), app))
          {
            throw new ArgumentException("invalid application module id");
          }
          DbItemCost item = DbItemCost.CreateItemCost(
              itemID: _.id,
              costType: _.type,
              costValue: NancyExtensions.GetBasicParameter<decimal>(Request.Form, "CostValue"),
              personID: NancyExtensions.GetBasicParameter<int>(Request.Form, "PersonID"), // XXX: at some point, we might want authentication here?
              applicationModule: (TP.Object.ApplicationModule)app
            );
          return NancyExtensions.Respond(item);
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      #endregion


      #region DbItemPrice - Read

      /// Returns current price information for an item, across all sales
      /// channels.
      /// 
      /// URL Parameters:
      ///   item id
      /// Query Parameters:
      ///   none
      /// Returns:
      ///   <ArrayOfItemCost>
      ///     <ItemPrice>
      ///       <TupleID>8</TupleID>
      ///       <ItemID>5321</ItemID>
      ///       <SalesChannelID>1</SalesChannelID>
      ///       <TimeEffective>2014-04-01T14:15:29.01</TimeEffective>
      ///       <PersonID>2</PersonID>
      ///       <ApplicationModuleID>PoinvInventoryMaintenance</ApplicationModuleID>
      ///       <MaximumQuantity1>100</MaximumQuantity1>
      ///       ...
      ///       <PriceValue1>9.9900</PriceValue1>
      ///       ...
      ///     </ItemPrice>
      ///     ...
      ///   </ArrayOfItemCost>
      Get["/{id:int}/price"] = _ =>
      {
        try
        {
          IEnumerable<DbItemPrice> list = DbItemPrice.GetCurrentByItemID(_.id, NancyExtensions.GetPreloads(Request.Query));
          return NancyExtensions.Respond(list);
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      /// Returns current price information for an item for a specified sales
      /// channel.
      /// 
      /// URL Parameters:
      ///   item id
      ///   sales channel id
      /// Query Parameters:
      ///   none
      /// Returns:
      ///   <ItemPrice>
      ///     <TupleID>8</TupleID>
      ///     <ItemID>5321</ItemID>
      ///     <SalesChannelID>1</SalesChannelID>
      ///     <TimeEffective>2014-04-01T14:15:29.01</TimeEffective>
      ///     <PersonID>2</PersonID>
      ///     <ApplicationModuleID>PoinvInventoryMaintenance</ApplicationModuleID>
      ///     <MaximumQuantity1>100</MaximumQuantity1>
      ///     ...
      ///     <PriceValue1>9.9900</PriceValue1>
      ///     ...
      ///   </ItemPrice>
      Get["/{id:int}/price/channel/{chanid:int}"] = _ =>
      {
        try
        {
          DbItemPrice item = DbItemPrice.GetChannelCurrentByItemID(_.chanid, _.id, NancyExtensions.GetPreloads(Request.Query));
          return NancyExtensions.Respond(item);
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      /// Returns historical price information for an item, across all sales
      /// channels.
      /// 
      /// URL Parameters:
      ///   item id
      /// Query Parameters:
      ///   DateLimit - optional, format like YYYY-MM-DD
      ///   RecordLimit - optional
      /// Returns:
      ///   <ArrayOfItemCost>
      ///     <ItemPrice>
      ///       <TupleID>8</TupleID>
      ///       <ItemID>5321</ItemID>
      ///       <SalesChannelID>1</SalesChannelID>
      ///       <TimeEffective>2014-04-01T14:15:29.01</TimeEffective>
      ///       <PersonID>2</PersonID>
      ///       <ApplicationModuleID>PoinvInventoryMaintenance</ApplicationModuleID>
      ///       <MaximumQuantity1>100</MaximumQuantity1>
      ///       ...
      ///       <PriceValue1>9.9900</PriceValue1>
      ///       ...
      ///     </ItemPrice>
      ///     ...
      ///   </ArrayOfItemCost>
      Get["/{id:int}/price/history"] = _ =>
      {
        try
        {
          IEnumerable<DbItemPrice> list = DbItemPrice.GetHistoryByItemID(
              id: _.id,
              dateLimit: NancyExtensions.GetBasicParameter<DateTime>(Request.Query, "DateLimit", new DateTime(2000, 1, 1)),
              recordLimit: NancyExtensions.GetBasicParameter<int>(Request.Query, "RecordLimit", -1)
            );
          return NancyExtensions.Respond(list);
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      /// Returns historical price information for an item for the specified
      /// sales channel.
      /// 
      /// URL Parameters:
      ///   item id
      ///   sales channel id
      /// Query Parameters:
      ///   DateLimit - optional, format like YYYY-MM-DD
      ///   RecordLimit - optional
      /// Returns:
      ///   <ArrayOfItemCost>
      ///     <ItemPrice>
      ///       <TupleID>8</TupleID>
      ///       <ItemID>5321</ItemID>
      ///       <SalesChannelID>1</SalesChannelID>
      ///       <TimeEffective>2014-04-01T14:15:29.01</TimeEffective>
      ///       <PersonID>2</PersonID>
      ///       <ApplicationModuleID>PoinvInventoryMaintenance</ApplicationModuleID>
      ///       <MaximumQuantity1>100</MaximumQuantity1>
      ///       ...
      ///       <PriceValue1>9.9900</PriceValue1>
      ///       ...
      ///     </ItemPrice>
      ///     ...
      ///   </ArrayOfItemCost>
      Get["/{id:int}/price/channel/{chanid:int}/history"] = _ =>
      {
        try
        {
          IEnumerable<DbItemPrice> list = DbItemPrice.GetChannelHistoryByItemID(
              chanID: _.chanid,
              id: _.id,
              dateLimit: NancyExtensions.GetBasicParameter<DateTime>(Request.Query, "DateLimit", new DateTime(2000, 1, 1)),
              recordLimit: NancyExtensions.GetBasicParameter<int>(Request.Query, "RecordLimit", -1)
            );
          return NancyExtensions.Respond(list);
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      #endregion


      #region DbItemPrice - Create

      /// Updates/creates the price for an item.
      /// 
      /// URL Parameters:
      ///   item id
      ///   channel id
      /// Query Parameters:
      ///   Break1 - required
      ///   Price1 - required
      ///   Break2 - optional, if Break1 == 999999
      ///   Price2 - optional, if Break1 == 999999
      ///   Break3 - optional, if Break2 == 999999
      ///   Price3 - optional, if Break2 == 999999
      ///   Break4 - optional, if Break3 == 999999
      ///   Price4 - optional, if Break3 == 999999
      ///   Break5 - optional, if Break4 == 999999
      ///   Price5 - optional, if Break4 == 999999
      ///   PersonID - required
      ///   ApplicationModuleID - required
      /// Returns:
      ///   <ItemPrice>
      ///     <TupleID>8</TupleID>
      ///     <ItemID>5321</ItemID>
      ///     <SalesChannelID>1</SalesChannelID>
      ///     <TimeEffective>2014-04-01T14:15:29.01</TimeEffective>
      ///     <PersonID>2</PersonID>
      ///     <ApplicationModuleID>PoinvInventoryMaintenance</ApplicationModuleID>
      ///     <MaximumQuantity1>100</MaximumQuantity1>
      ///     <MaximumQuantity2>999999</MaximumQuantity2>
      ///     <MaximumQuantity3>0</MaximumQuantity3>
      ///     <MaximumQuantity4>0</MaximumQuantity4>
      ///     <MaximumQuantity5>0</MaximumQuantity5>
      ///     <PriceValue1>9.9900</PriceValue1>
      ///     <PriceValue2>8.9900</PriceValue2>
      ///     <PriceValue3>0.0000</PriceValue3>
      ///     <PriceValue4>0.0000</PriceValue4>
      ///     <PriceValue5>0.0000</PriceValue5>
      ///   </ItemPrice>
      Post["/{id:int}/price/channel/{chanid:int}"] = _ =>
      {
        try
        {
          int app = NancyExtensions.GetBasicParameter<int>(Request.Form, "ApplicationModule");
          if (!Enum.IsDefined(typeof(TP.Object.ApplicationModule), app))
          {
            throw new ArgumentException("invalid application module id");
          }
          DbItemPrice item = DbItemPrice.CreateItemPrice(
              itemID: _.id,
              salesChannelID: _.chanid,
              q1: NancyExtensions.GetBasicParameter<int>(Request.Form, "Break1"),
              p1: NancyExtensions.GetBasicParameter<decimal>(Request.Form, "Price1"),
              q2: NancyExtensions.GetBasicParameter<int>(Request.Form, "Break2"),
              p2: NancyExtensions.GetBasicParameter<decimal>(Request.Form, "Price2"),
              q3: NancyExtensions.GetBasicParameter<int>(Request.Form, "Break3"),
              p3: NancyExtensions.GetBasicParameter<decimal>(Request.Form, "Price3"),
              q4: NancyExtensions.GetBasicParameter<int>(Request.Form, "Break4"),
              p4: NancyExtensions.GetBasicParameter<decimal>(Request.Form, "Price4"),
              q5: NancyExtensions.GetBasicParameter<int>(Request.Form, "Break5"),
              p5: NancyExtensions.GetBasicParameter<decimal>(Request.Form, "Price5"),
              personID: NancyExtensions.GetBasicParameter<int>(Request.Form, "PersonID"),
              applicationModule: (TP.Object.ApplicationModule)app
            );
          return NancyExtensions.Respond(item);
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      #endregion
    }
  }

  // ComponentCRUD handles:
  //   OBJECT                           READ  CREATE  UPDATE  DELETE
  //   DbComponent                        Y     Y       Y       N
  //   DbItemComponent                    Y     Y       Y       Y
  //   DbBarcode                          Y     N       N       N
  //   DbInventoryLocationContent         Y     Y       Y       Y
  public class ComponentCRUD : NancyModule
  {
    public ComponentCRUD() : base("/component")
    {
      #region DbComponent - Read

      /// Returns a single component given the id.
      /// 
      /// URL Parameters:
      ///   component id
      /// Query Parameters:
      ///   preload:
      ///     barcodes
      /// Returns:
      ///   <Component>
      ///     <ComponentID>9100</ComponentID>
      ///     <ComponentName>MAA164095-8</ComponentName>
      ///     <Weight>0.28</Weight>
      ///     <Length>11.75</Length>
      ///     ...
      ///   </Component>
      Get["/{id:int}"] = _ =>
      {
        try
        {
          DbComponent item = DbComponent.GetByID(_.id, NancyExtensions.GetPreloads(Request.Query));
          return NancyExtensions.Respond(item);
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      Get["/{id:int}/name"] = _ =>
      {
        try
        {
          DbComponent item = DbComponent.GetByID(_.id, NancyExtensions.GetPreloads(Request.Query));
          if (item == null) { return NancyExtensions.Respond(item); }
          return NancyExtensions.Respond(new TP.Base.SuccessResponse(item.ComponentName));
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      Get["/{name}"] = _ =>
      {
        try
        {
          DbComponent item = DbComponent.GetByName(_.name, NancyExtensions.GetPreloads(Request.Query));
          if (item == null) { return NancyExtensions.Respond(item); }
          return NancyExtensions.Respond(item);
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      #endregion


      #region DbComponent - Create

      /// Creates a new component and links it to the given item with an
      /// ItemComponent tuple.
      /// 
      /// URL Parameters:
      ///   none
      /// Query Parameters:
      ///   ComponentName
      ///   ItemID
      /// Returns:
      ///   <Component>
      ///     <ComponentID>9100</ComponentID>
      ///     <ComponentName>MAA164095-8</ComponentName>
      ///     <Weight>0.28</Weight>
      ///     <Length>11.75</Length>
      ///     ...
      ///   </Component>
      Post["/create"] = _ =>
      {
        try
        {
          DbComponent item = DbComponent.CreateComponent(
              componentName: NancyExtensions.GetBasicParameter<string>(Request.Form, "ComponentName"),
              attachedToItemID: NancyExtensions.GetBasicParameter<int>(Request.Form, "ItemID")
            );
          return NancyExtensions.Respond(item);
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      #endregion


      #region DbComponent - Update

      /// Updates the given component's miscellaneous fields
      /// 
      /// URL Parameters:
      ///   component id
      /// Query Parameters:
      ///   TP.Object.Component => serialized object with new data
      /// Returns:
      ///   <Component>
      ///     <ComponentID>9100</ComponentID>
      ///     <ComponentName>MAA164095-8</ComponentName>
      ///     <Weight>0.28</Weight>
      ///     <Length>11.75</Length>
      ///     ...
      ///   </Component>
      Post["/update/{id:int}"] = _ =>
      {
        try
        {
          DbComponent item = NancyExtensions.GetDbObjectParameter<DbComponent>(Request.Form, "TP.Object.Component");
          if (item.ComponentID != _.id)
          {
            throw new ArgumentException("url id does not match given object!");
          }
          return NancyExtensions.Respond(item.UpdateComponent());
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      #endregion


      #region DbItemComponent - Read

      /// Returns a list of ItemComponent tuples for the given item. Using the
      /// components preload is probably really useful here, since there isn't
      /// much to these.
      /// 
      /// URL Parameters:
      ///   item id
      /// Query Parameters:
      ///   preload:
      ///     components
      ///       barcodes
      /// Returns:
      ///   <ArrayOfItemComponent>
      ///     <ItemComponent>
      ///       <ComponentID>11552</ComponentID>
      ///       <ItemID>21535</ItemID>
      ///       <Quantity>1</Quantity>
      ///     </ItemComponent>
      ///     ...
      ///   </ArrayOfItemComponent>
      Get["/item/{id:int}"] = _ =>
      {
        try
        {
          IEnumerable<DbItemComponent> list = DbItemComponent.GetByItemID(_.id, NancyExtensions.GetPreloads(Request.Query));
          return NancyExtensions.Respond(list);
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      /// Returns a list of ItemComponent tuples for the given component.
      /// 
      /// URL Parameters:
      ///   component id
      /// Query Parameters:
      ///   preload:
      ///     components
      ///       barcodes
      /// Returns:
      ///   <ArrayOfItemComponent>
      ///     <ItemComponent>
      ///       <ComponentID>11552</ComponentID>
      ///       <ItemID>21535</ItemID>
      ///       <Quantity>1</Quantity>
      ///     </ItemComponent>
      ///     ...
      ///   </ArrayOfItemComponent>
      Get["/{id:int}/items"] = _ =>
      {
        try
        {
          IEnumerable<DbItemComponent> list = DbItemComponent.GetByComponentID(_.id, NancyExtensions.GetPreloads(Request.Query));
          return NancyExtensions.Respond(list);
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      #endregion


      #region DbItemComponent - Create

      /// Links an existing component to an existing item
      /// 
      /// URL Parameters:
      ///   component id
      ///   item id
      /// Query Parameters:
      ///   Quantity (optional, defaults to 1)
      /// Returns:
      ///   <ItemComponent>
      ///     <ComponentID>11552</ComponentID>
      ///     <ItemID>21535</ItemID>
      ///     <Quantity>1</Quantity>
      ///   </ItemComponent>
      Post["/{id:int}/link/item/{iid:int}"] = _ =>
      {
        try
        {
          bool success = DbItemComponent.CreateItemComponentLink(
              componentID: _.id,
              itemID: _.iid,
              quantity: NancyExtensions.GetBasicParameter<int>(Request.Form, "Quantity", 1)
            );
          if (success)
          {
            DbItemComponent ic = DbItemComponent.GetByPrimaryKey(_.iid, _.id);
            return NancyExtensions.Respond(ic);
          }
          else
          {
            return NancyExtensions.Respond("error linking");
          }
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      #endregion


      #region DbItemComponent - Update

      /// Modifies the given ItemComponent fields (qty only so far).
      /// 
      /// URL Parameters:
      ///   component id
      ///   item id
      /// Query Parameters:
      ///   TP.Object.ItemComponent => serialized object with new data
      /// Returns:
      ///   <ItemComponent>
      ///     <ComponentID>11552</ComponentID>
      ///     <ItemID>21535</ItemID>
      ///     <Quantity>1</Quantity>
      ///   </ItemComponent>
      Post["/{id:int}/item/{iid:int}/update"] = _ =>
      {
        try
        {
          DbItemComponent item = NancyExtensions.GetDbObjectParameter<DbItemComponent>(Request.Form, "TP.Object.ItemComponent");
          if (item.ComponentID != _.id)
          {
            throw new ArgumentException("url id does not match given object!");
          }
          if (item.ItemID != _.iid)
          {
            throw new ArgumentException("url id does not match given object!");
          }
          return NancyExtensions.Respond(item.UpdateItemComponentLink());
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      #endregion


      #region DbItemComponent - Delete

      /// Deletes the given ItemComponent mapping. Note that if this is the
      /// only mapping left for the component, it and any sub-values (barcode)
      /// will be deleted as well.
      /// 
      /// URL Parameters:
      ///   component id
      ///   item id
      /// Query Parameters:
      ///   TP.Object.ItemComponent => serialized object with current data
      /// Returns:
      ///   serialized SuccessResponse or ErrorResponse
      Post["/{id:int}/unlink/item/{iid:int}"] = _ =>
      {
        try
        {
          DbItemComponent item = NancyExtensions.GetDbObjectParameter<DbItemComponent>(Request.Form, "TP.Object.ItemComponent");
          if (item.ComponentID != _.id)
          {
            throw new ArgumentException("url id does not match given object!");
          }
          if (item.ItemID != _.iid)
          {
            throw new ArgumentException("url id does not match given object!");
          }
          if (item.DeleteItemComponentLink())
          {
            return NancyExtensions.Respond(new TP.Base.SuccessResponse("item/component deleted"));
          }
          else
          {
            return NancyExtensions.Respond(new TP.Base.ErrorResponse("delete failed"));
          }
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      #endregion


      #region DbInventoryLocationContent - Read

      /// Returns a list of InventoryLocationContent objects showing where and
      /// how many of this component exist in the warehouse. Optionally
      /// filterable by warehouse id and location type id. XXX: Any way to get
      /// the routes into a single method?
      /// 
      /// URL Parameters:
      ///   component id
      ///   warehouse id - optional
      ///   location type id - optional
      /// Query Parameters:
      ///   preload:
      ///     location
      ///     component
      /// Returns:
      ///   <ArrayOfInventoryLocationContent>
      ///     <InventoryLocationContent>
      ///       <LocationID>8651</LocationID>
      ///       <ComponentID>31736</ComponentID>
      ///       <MasterID>8651</MasterID>
      ///       <Quantity>8></Quantity>
      ///       <AssociatedTaskType i:nil="true"/>
      ///       <AssociatedTaskID i:nil="true"/>
      ///     </InventoryLocationContent>
      ///     ...
      ///   </ArrayOfInventoryLocationContent>
      Get["/{id:int}/locations"] = _ =>
      {
        try
        {
          IEnumerable<DbInventoryLocationContent> list = DbInventoryLocationContent.GetByComponentID(_.id, NancyExtensions.GetPreloads(Request.Query));
          return NancyExtensions.Respond(list);
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };
      Get["/{id:int}/locations/warehouse/{wid:int}"] = _ =>
      {
        IEnumerable<DbInventoryLocationContent> list1 = DbInventoryLocationContent.GetByComponentID(_.id, NancyExtensions.GetPreloads(Request.Query));
        IEnumerable<DbInventoryLocationContent> list2 = list1.Where(i => DbInventoryLocation.GetByID(i.LocationID).WarehouseID == _.wid);
        return NancyExtensions.Respond(list2);
      };
      Get["/{id:int}/locations/warehouse/{wid:int}/locationtype/{ltid:int}"] = _ =>
      {
        IEnumerable<DbInventoryLocationContent> list1 = DbInventoryLocationContent.GetByComponentID(_.id, NancyExtensions.GetPreloads(Request.Query));
        IEnumerable<DbInventoryLocationContent> list2 = list1.Where(i => { var l = DbInventoryLocation.GetByID(i.LocationID); return l.WarehouseID == _.wid && l.LocationTypeID == _.ltid; });
        return NancyExtensions.Respond(list2);
      };

      Get["/{id:int}/location/{lid:int}"] = _ =>
      {
        try
        {
          IEnumerable<DbInventoryLocationContent> list = DbInventoryLocationContent.GetByComponentID(_.id, NancyExtensions.GetPreloads(Request.Query));
          DbInventoryLocationContent item = list.SingleOrDefault(lc => lc.LocationID == _.lid);
          return NancyExtensions.Respond(item);
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      #endregion


      #region DbInventoryLocationContent - Update

      //Post["/{id:int}/location/{lid:int}/inventoried"] = _ =>
      //{
      //  try
      //  {
      //    DbInventoryLocationContent.UpdateInventoriedDate(_.lid, _.cid, NancyExtensions.GetBasicParameter<int>(Request.Form, "Quantity"));
      //    return NancyExtensions.Respond(new TP.Base.SuccessResponse("success"));
      //  }
      //  catch (Exception ex)
      //  {
      //    return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
      //  }
      //};

      #endregion


      #region DbBarcode - Read

      /// Returns a list of barcode objects for a given component.
      /// 
      /// URL Parameters:
      ///   component id
      /// Query Parameters:
      ///   none
      /// Returns:
      ///   <ArrayOfBarcode>
      ///     <Barcode>
      ///       <Code>099555937664</Code>
      ///       <ComponentID>14302</ComponentID>
      ///       <IsInternal>false</IsInternal>
      ///       <PrintByStandardPack>false</PrintByStandardPack>
      ///       <Quantity>1</Quantity>
      ///     </Barcode>
      ///     ...
      ///   </ArrayOfBarcode>
      Get["/{id:int}/barcodes"] = _ =>
      {
        try
        {
          IEnumerable<DbBarcode> list = DbBarcode.GetByComponentID(_.id, NancyExtensions.GetPreloads(Request.Query));
          return NancyExtensions.Respond(list);
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      #endregion


      #region DbInventoryLocationContent - Create/Update/Delete (Merged)

      /// Adds or removes an item from a location in the warehouse, without an
      /// associated task. E.g., receiving.
      /// 
      /// URL Parameters:
      ///   component id
      ///   location id
      /// Query Parameters:
      ///   LocationStart        - id of beginning of location span
      ///   LocationEnd          - id of end of location span
      ///   ComponentID          - component
      ///   DeltaQuantity        - quantity change, must be positive
      ///   TransactionReference - optional, use to describe reason for movement
      ///   Comment              - optional, extended info about movement
      ///   NTUsername           - optional but probably required, who is doing this
      ///   ComputerName         - optional but probably required, native computer
      ///   RDPClientName        - optional but probably required
      /// Returns:
      ///   serialized SuccessResponse or ErrorResponse
      Post["/{id:int}/locations/{lid:int}/change"] = _ =>
      {
        try
        {
          int cid = NancyExtensions.GetBasicParameter<int>(Request.Form, "ComponentID");
          int ls = NancyExtensions.GetBasicParameter<int>(Request.Form, "LocationStart");
          if (cid != _.id)
          {
            throw new ArgumentException("url id does not match given request body!");
          }
          if (ls != _.lid)
          {
            throw new ArgumentException("url id does not match given request body!");
          }
          bool ret = DbInventoryLocationContent.AddOrRemove(
              startLocationID: ls,
              endLocationID: NancyExtensions.GetBasicParameter<int>(Request.Form, "LocationEnd"),
              componentID: cid,
              deltaQuantity: NancyExtensions.GetBasicParameter<int>(Request.Form, "DeltaQuantity"),
              transactionReference: NancyExtensions.GetBasicParameter<string>(Request.Form, "TransactionReference"),
              comment: NancyExtensions.GetBasicParameter<string>(Request.Form, "Comment"),
              personNTUsername: NancyExtensions.GetBasicParameter<string>(Request.Form, "NTUsername"),
              computerName: NancyExtensions.GetBasicParameter<string>(Request.Form, "ComputerName"),
              rdpClientName: NancyExtensions.GetBasicParameter<string>(Request.Form, "RDPClientName")
            );
          if (ret)
          {
            return NancyExtensions.Respond(new TP.Base.SuccessResponse("success"));
          }
          else
          {
            return NancyExtensions.Respond(new TP.Base.ErrorResponse("this won't ever happen"));
          }
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      /// Moves a component from one location to another within the warehouse,
      /// without an associated task. E.g., putting, consolidation, etc.
      /// 
      /// URL Parameters:
      ///   component id
      ///   from location id
      /// Query Parameters:
      ///   FromLocationStart    - id of beginning of from location span
      ///   FromLocationEnd      - id of end of from location span
      ///   ToLocationStart      - id of beginning of to location span
      ///   ToLocationEnd        - id of end of to location span
      ///   ComponentID          - component
      ///   DeltaQuantity        - quantity change, must be positive
      ///   TransactionReference - optional, use to describe reason for movement
      ///   Comment              - optional, extended info about movement
      ///   NTUsername           - optional but probably required, who is doing this
      ///   ComputerName         - optional but probably required, native computer
      ///   RDPClientName        - optional but probably required
      /// Returns:
      ///   serialized SuccessResponse or ErrorResponse
      Post["/{id:int}/locations/{lid:int}/move"] = _ =>
      {
        try
        {
          int cid = NancyExtensions.GetBasicParameter<int>(Request.Form, "ComponentID");
          int fls = NancyExtensions.GetBasicParameter<int>(Request.Form, "FromLocationStart");
          if (cid != _.id)
          {
            throw new ArgumentException("url id does not match given request body!");
          }
          if (fls != _.lid)
          {
            throw new ArgumentException("url id does not match given request body!");
          }
          bool ret = DbInventoryLocationContent.MoveFromTo(
              fromStartLocationID: fls,
              fromEndLocationID: NancyExtensions.GetBasicParameter<int>(Request.Form, "FromLocationEnd"),
              toStartLocationID: NancyExtensions.GetBasicParameter<int>(Request.Form, "ToLocationStart"),
              toEndLocationID: NancyExtensions.GetBasicParameter<int>(Request.Form, "ToLocationEnd"),
              componentID: cid,
              deltaQuantity: NancyExtensions.GetBasicParameter<int>(Request.Form, "DeltaQuantity"),
              transactionReference: NancyExtensions.GetBasicParameter<string>(Request.Form, "TransactionReference"),
              comment: NancyExtensions.GetBasicParameter<string>(Request.Form, "Comment"),
              personNTUsername: NancyExtensions.GetBasicParameter<string>(Request.Form, "NTUsername"),
              computerName: NancyExtensions.GetBasicParameter<string>(Request.Form, "ComputerName"),
              rdpClientName: NancyExtensions.GetBasicParameter<string>(Request.Form, "RDPClientName")
            );
          if (ret)
          {
            return NancyExtensions.Respond(new TP.Base.SuccessResponse("success"));
          }
          else
          {
            return NancyExtensions.Respond(new TP.Base.ErrorResponse("this won't ever happen"));
          }
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      #endregion
    }
  }

  // BarcodeCRUD handles:
  //   DbBarcode                          Y     Y       Y       Y
  public class BarcodeCRUD : NancyModule
  {
    public BarcodeCRUD() : base("/barcode")
    {
      #region DbBarcode - Read

      /// Returns a barcode object by barcode, probably only if we have
      /// scanner integration or something?
      /// 
      /// URL Parameters:
      ///   code
      /// Query Parameters:
      ///   none
      /// Returns:
      ///   <Barcode>
      ///     <Code>099555937664</Code>
      ///     <ComponentID>14302</ComponentID>
      ///     <IsInternal>false</IsInternal>
      ///     <PrintByStandardPack>false</PrintByStandardPack>
      ///     <Quantity>1</Quantity>
      ///   </Barcode>
      Get["/{code}"] = _ =>
      {
        try
        {
          DbBarcode item = DbBarcode.GetByCode(_.code, NancyExtensions.GetPreloads(Request.Query));
          return NancyExtensions.Respond(item);
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      /// Returns a list of barcode objects for a given component.
      /// 
      /// URL Parameters:
      ///   component id
      /// Query Parameters:
      ///   none
      /// Returns:
      ///   <ArrayOfBarcode>
      ///     <Barcode>
      ///       <Code>099555937664</Code>
      ///       <ComponentID>14302</ComponentID>
      ///       <IsInternal>false</IsInternal>
      ///       <PrintByStandardPack>false</PrintByStandardPack>
      ///       <Quantity>1</Quantity>
      ///     </Barcode>
      ///     ...
      ///   </ArrayOfBarcode>
      Get["/component/{id:int}"] = _ =>
      {
        try
        {
          IEnumerable<DbBarcode> list = DbBarcode.GetByComponentID(_.id, NancyExtensions.GetPreloads(Request.Query));
          return NancyExtensions.Respond(list);
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      #endregion


      #region DbBarcode - Create

      /// Creates a new barcode, associated with the given component. Quantity
      /// is optional, and defaults to 1. Note that quantity is not locked to
      /// the type of barcode, UPCs can be for multiple, 14-digit master pack
      /// barcodes can be for a single sellable unit.
      /// 
      /// URL Parameters:
      ///   none
      /// Query Parameters:
      ///   Code
      ///   Component
      ///   Quantity (optional, defaults to 1)
      /// Returns:
      ///   <Barcode>
      ///     <Code>099555937664</Code>
      ///     <ComponentID>14302</ComponentID>
      ///     <IsInternal>false</IsInternal>
      ///     <PrintByStandardPack>false</PrintByStandardPack>
      ///     <Quantity>1</Quantity>
      ///   </Barcode>
      Post["/create"] = _ =>
      {
        try
        {
          DbBarcode item = DbBarcode.CreateBarcode(
              code: NancyExtensions.GetBasicParameter<string>(Request.Form, "Code"),
              componentID: NancyExtensions.GetBasicParameter<int>(Request.Form, "ComponentID"),
              quantity: NancyExtensions.GetBasicParameter<int>(Request.Form, "Quantity", 1)
            );
          return NancyExtensions.Respond(item);
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      /// Creates a new internal-only sequential barcode.
      /// 
      /// URL Parameters:
      ///   none
      /// Query Parameters:
      ///   Code
      ///   Component
      ///   Quantity (optional, defaults to 1)
      /// Returns:
      ///   <Barcode>
      ///     <Code>TP0000000259</Code>
      ///     <ComponentID>14302</ComponentID>
      ///     <IsInternal>true</IsInternal>
      ///     <PrintByStandardPack>false</PrintByStandardPack>
      ///     <Quantity>1</Quantity>
      ///   </Barcode>
      Post["/create/TP"] = _ =>
      {
        try
        {
          DbBarcode item = DbBarcode.CreateToolsPlusBarcode(
              componentID: NancyExtensions.GetBasicParameter<int>(Request.Form, "ComponentID")
            );
          return NancyExtensions.Respond(item);
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      #endregion


      #region DbBarcode - Update

      /// Modifies the barcode fields (associated component, quantity, etc).
      /// 
      /// URL Parameters:
      ///   code
      /// Query Parameters:
      ///   TP.Object.Barcode => serialized object with new data
      /// Returns:
      ///   <Barcode>
      ///     <Code>TP0000000259</Code>
      ///     <ComponentID>14302</ComponentID>
      ///     <IsInternal>true</IsInternal>
      ///     <PrintByStandardPack>false</PrintByStandardPack>
      ///     <Quantity>1</Quantity>
      ///   </Barcode>
      Post["/update/{code}"] = _ =>
      {
        try
        {
          DbBarcode item = NancyExtensions.GetDbObjectParameter<DbBarcode>(Request.Form, "TP.Object.Barcode");
          if (item.Code != _.code)
          {
            throw new ArgumentException("url id does not match given object!");
          }
          return NancyExtensions.Respond(item.UpdateBarcode());
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      #endregion


      #region DbBarcode - Delete

      /// Deletes the given barcode.
      /// 
      /// URL Parameters:
      ///   code
      /// Query Parameters:
      ///   TP.Object.Barcode => serialized object with current data
      /// Returns:
      ///   serialized SuccessResponse or ErrorResponse
      Post["/delete/{code}"] = _ =>
      {
        try
        {
          DbBarcode item = NancyExtensions.GetDbObjectParameter<DbBarcode>(Request.Form, "TP.Object.Barcode");
          if (item.Code != _.code)
          {
            throw new ArgumentException("url id does not match given object!");
          }
          if (item.DeleteBarcode())
          {
            return NancyExtensions.Respond(new TP.Base.SuccessResponse("barcode deleted"));
          }
          else
          {
            return NancyExtensions.Respond(new TP.Base.ErrorResponse("delete failed"));
          }
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      #endregion
    }
  }

  // PurchaseOrderCRUD handles:
  //   OBJECT                           READ  CREATE  UPDATE  DELETE
  //   DbPurchaseOrder                    Y     Y       Y       Y
  //   DbPurchaseOrderItem                Y     Y       Y       Y
  //   DbPurchaseOrderTracking            Y     Y       Y       N
  public class PurchaseOrderCRUD : NancyModule
  {
    public PurchaseOrderCRUD() : base("/purchaseorder/")
    {

      #region DbPurchaseOrder - Read

      /// Returns a purchase order object, given its internal id.
      /// 
      /// URL Parameters:
      ///   purchase order id
      /// Query Parameters:
      ///   preload:
      ///     vendor
      ///     items
      ///     tracking
      /// Returns:
      ///   <PurchaseOrder>
      ///     <PurchaseOrderID>32171</PurchaseOrderID>
      ///     <PurchaseOrderNo>A35000</PurchaseOrderNo>
      ///     <VendorNo>MLW</VendorNo>
      ///     <CoopVendorNo>MLW</CoopVendorNo>
      ///     <ProductLine>MLW</ProductLine>
      ///     ...
      ///   </PurchaseOrder>
      Get["/{id:int}"] = _ =>
      {
        try
        {
          DbPurchaseOrder item = DbPurchaseOrder.GetByID(_.id, NancyExtensions.GetPreloads(Request.Query));
          return NancyExtensions.Respond(item);
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      /// Returns a purchase order object, given its generated purchase order
      /// number for Mas. This is probably not really a good lookup method,
      /// since purchase order numbers might change from the current method of
      /// ^A\d{5,6}$ naming, and a number only PO like Mas assigns would be
      /// conflicting with the id # lookup. Depending on how much use this
      /// method gets, we might want to get rid of it.
      /// 
      /// URL Parameters:
      ///   purchase order number
      /// Query Parameters:
      ///   preload:
      ///     vendor
      ///     items
      ///     tracking
      /// Returns:
      ///   <PurchaseOrder>
      ///     <PurchaseOrderID>32171</PurchaseOrderID>
      ///     <PurchaseOrderNo>A35000</PurchaseOrderNo>
      ///     <VendorNo>MLW</VendorNo>
      ///     <CoopVendorNo>MLW</CoopVendorNo>
      ///     <ProductLine>MLW</ProductLine>
      ///     ...
      ///   </PurchaseOrder>
      Get["/{po}"] = _ =>
      {
        try
        {
          DbPurchaseOrder item = DbPurchaseOrder.GetByPurchaseOrderNo(_.po, NancyExtensions.GetPreloads(Request.Query));
          return NancyExtensions.Respond(item);
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      /// Returns a list of purchase order objects. This will return any
      /// active (non-exported) purchase orders that the given item is able to
      /// be added to (matching vendor/coop/line). For regular items, this
      /// should be 0 or 1 objects, since we don't currently support adding to
      /// multiple POs through the GUI. For dropship items, however, it can be
      /// any number.
      /// 
      /// URL Parameters:
      ///   item id
      /// Query Parameters:
      ///   preload:
      ///     vendor
      ///     items
      ///     tracking
      /// Returns:
      ///   <ArrayOfPurchaseOrder>
      ///     <PurchaseOrder>
      ///       <PurchaseOrderID>32171</PurchaseOrderID>
      ///       <PurchaseOrderNo>A35000</PurchaseOrderNo>
      ///       <VendorNo>MLW</VendorNo>
      ///       <CoopVendorNo>MLW</CoopVendorNo>
      ///       <ProductLine>MLW</ProductLine>
      ///       ...
      ///     </PurchaseOrder>
      ///     ...
      ///   </ArrayOfPurchaseOrder>
      Get["/item/{id:int}"] = _ =>
      {
        try
        {
          IEnumerable<DbPurchaseOrder> list = DbPurchaseOrder.GetForItem(_.id, NancyExtensions.GetPreloads(Request.Query));
          return NancyExtensions.Respond(list);
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      /// Returns a list of all active purchase orders (non-exported).
      /// 
      /// URL Parameters:
      ///   none
      /// Query Parameters:
      ///   preload:
      ///     vendor
      ///     items
      ///     tracking
      /// Returns:
      ///   <ArrayOfPurchaseOrder>
      ///     <PurchaseOrder>
      ///       <PurchaseOrderID>32171</PurchaseOrderID>
      ///       <PurchaseOrderNo>A35000</PurchaseOrderNo>
      ///       <VendorNo>MLW</VendorNo>
      ///       <CoopVendorNo>MLW</CoopVendorNo>
      ///       <ProductLine>MLW</ProductLine>
      ///       ...
      ///     </PurchaseOrder>
      ///     ...
      ///   </ArrayOfPurchaseOrder>
      Get["/active"] = _ =>
      {
        try
        {
          IEnumerable<DbPurchaseOrder> list = DbPurchaseOrder.GetAllActive(NancyExtensions.GetPreloads(Request.Query));
          return NancyExtensions.Respond(list);
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      #endregion


      #region DbPurchaseOrder - Create

      /// Creates a new purchase order for stock, with a single item on it.
      /// 
      /// URL Parameters:
      ///   none
      /// Query Parameters:
      ///   Code
      ///   Component
      ///   Quantity (optional, defaults to 1)
      /// Returns:
      ///   <PurchaseOrder>
      ///     <PurchaseOrderID>32171</PurchaseOrderID>
      ///     <PurchaseOrderNo>A35000</PurchaseOrderNo>
      ///     <VendorNo>MLW</VendorNo>
      ///     <CoopVendorNo>MLW</CoopVendorNo>
      ///     <ProductLine>MLW</ProductLine>
      ///     ...
      ///   </PurchaseOrder>
      Post["/create/stock"] = _ =>
      {
        try
        {
          DbPurchaseOrder item = DbPurchaseOrder.CreatePurchaseOrderForStock(
              initialItemID: NancyExtensions.GetBasicParameter<int>(Request.Form, "ItemID"),
              initialItemQuantity: NancyExtensions.GetBasicParameter<int>(Request.Form, "Quantity")
            );
          return NancyExtensions.Respond(item);
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      #endregion


      #region DbPurchaseOrder - Update

      /// Modifies the purchase order's fields. This isn't really a great way
      /// to do it, since there's some dependent fields in the purchase order,
      /// but it'll work for now.
      /// 
      /// URL Parameters:
      ///   purchase order id
      /// Query Parameters:
      ///   TP.Object.PurchaseOrder => serialized object with new data
      /// Returns:
      ///   <PurchaseOrder>
      ///     <PurchaseOrderID>32171</PurchaseOrderID>
      ///     <PurchaseOrderNo>A35000</PurchaseOrderNo>
      ///     <VendorNo>MLW</VendorNo>
      ///     <CoopVendorNo>MLW</CoopVendorNo>
      ///     <ProductLine>MLW</ProductLine>
      ///     ...
      ///   </PurchaseOrder>
      Post["/update/{id:int}"] = _ =>
      {
        try
        {
          DbPurchaseOrder item = NancyExtensions.GetDbObjectParameter<DbPurchaseOrder>(Request.Form, "TP.Object.PurchaseOrder");
          if (item.PurchaseOrderID != _.id)
          {
            throw new ArgumentException("url id does not match given object!");
          }
          return NancyExtensions.Respond(item.UpdatePurchaseOrder());
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      #endregion


      #region DbPurchaseOrder - Delete

      /// Deletes the given purchase order. Note that currently, the purchase
      /// order must be completely empty of items and not exported.
      /// 
      /// URL Parameters:
      ///   purchase order id
      /// Query Parameters:
      ///   TP.Object.PurchaseOrder => serialized object with current data
      /// Returns:
      ///   serialized SuccessResponse or ErrorResponse
      Post["/delete/{id:int}"] = _ =>
      {
        try
        {
          DbPurchaseOrder item = NancyExtensions.GetDbObjectParameter<DbPurchaseOrder>(Request.Form, "TP.Object.PurchaseOrder");
          if (item.PurchaseOrderID != _.id)
          {
            throw new ArgumentException("url id does not match given object!");
          }
          if (item.DeletePurchaseOrder())
          {
            return NancyExtensions.Respond(new TP.Base.SuccessResponse("purchase order deleted"));
          }
          else
          {
            return NancyExtensions.Respond(new TP.Base.ErrorResponse("delete failed"));
          }
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      #endregion


      #region DbPurchaseOrderItem - Read

      /// Returns a list of the given purchase order's items.
      /// 
      /// URL Parameters:
      ///   purchase order id
      /// Query Parameters:
      ///   preload:
      ///     item
      ///     purchaseorder
      /// Returns:
      ///   <ArrayOfPurchaseOrderItem>
      ///     <PurchaseOrderItem>
      ///       <TupleID>1156</TupleID>
      ///       <HeaderID>30665</HeaderID>
      ///       <ItemNumber>REE02093</ItemNumber>
      ///       <Quantity>6</Quantity>
      ///     </PurchaseOrderItem>
      ///     ...
      ///   </ArrayOfPurchaseOrderItem>
      Get["/{id:int}/itemlines"] = _ =>
      {
        try
        {
          IEnumerable<DbPurchaseOrderItem> list = DbPurchaseOrderItem.GetByPurchaseOrderID(_.id, NancyExtensions.GetPreloads(Request.Query));
          return NancyExtensions.Respond(list);
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      /// Returns a list of current active purchase order lines that the item
      /// is associated with. For a potential vendor purchase order that the
      /// item is not on, but could be, see /purchaseorder/item/{id}. Now, the
      /// number of results should be 0 or 1 for stock items, and 0+ for
      /// dropship items, but this is a GUI limitation only.
      /// 
      /// URL Parameters:
      ///   item id
      /// Query Parameters:
      ///   preload:
      ///     item
      ///     purchaseorder
      /// Returns:
      ///   <ArrayOfPurchaseOrderItem>
      ///     <PurchaseOrderItem>
      ///       <TupleID>1156</TupleID>
      ///       <HeaderID>30665</HeaderID>
      ///       <ItemNumber>REE02093</ItemNumber>
      ///       <Quantity>6</Quantity>
      ///     </PurchaseOrderItem>
      ///     ...
      ///   </ArrayOfPurchaseOrderItem>
      Get["/itemlines/item/{id:int}/active"] = _ =>
      {
        try
        {
          IEnumerable<DbPurchaseOrderItem> list = DbPurchaseOrderItem.GetCurrentByItemID(_.id, NancyExtensions.GetPreloads(Request.Query));
          return NancyExtensions.Respond(list);
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      #endregion


      #region DbPurchaseOrderItem - Create

      /// Adds an item to a purchase order.
      /// 
      /// URL Parameters:
      ///   purchase order id
      ///   item id
      /// Query Parameters:
      ///   Quantity
      /// Returns:
      ///   <PurchaseOrderItem>
      ///     <TupleID>1156</TupleID>
      ///     <HeaderID>30665</HeaderID>
      ///     <ItemNumber>REE02093</ItemNumber>
      ///     <Quantity>6</Quantity>
      ///   </PurchaseOrderItem>
      Post["/{id:int}/itemlines/add/{itemid:int}"] = _ =>
      {
        try
        {
          DbPurchaseOrderItem item = DbPurchaseOrderItem.AddItemToPurchaseOrder(
              poID: _.id,
              itemID: _.itemid,
              quantity: NancyExtensions.GetBasicParameter<int>(Request.Form, "Quantity")
            );
          return NancyExtensions.Respond(item);
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      #endregion


      #region DbPurchaseOrderItem - Update

      /// Modifies the purchase order line item. Currently the only option is
      /// quantity, but we may move comments and costs into here eventually,
      /// if we move forward with multiple stock POs.
      /// 
      /// URL Parameters:
      ///   purchase order id
      ///   item id
      /// Query Parameters:
      ///   TP.Object.PurchaseOrderItem => serialized object with new data
      /// Returns:
      ///   <PurchaseOrderItem>
      ///     <TupleID>1156</TupleID>
      ///     <HeaderID>30665</HeaderID>
      ///     <ItemNumber>REE02093</ItemNumber>
      ///     <Quantity>6</Quantity>
      ///   </PurchaseOrderItem>
      Post["/{id:int}/itemlines/update/{itemid:int}"] = _ =>
      {
        try
        {
          DbPurchaseOrderItem item = NancyExtensions.GetDbObjectParameter<DbPurchaseOrderItem>(Request.Form, "TP.Object.PurchaseOrderItem");
          if (item.HeaderID != _.id)
          {
            throw new ArgumentException("url id does not match given object!");
          }
          if (item.ItemID != _.itemid)
          {
            throw new ArgumentException("url id does not match given object!");
          }
          return NancyExtensions.Respond(item.UpdatePurchaseOrderItem());
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      #endregion


      #region DbPurchaseOrderItem - Delete

      /// Deletes the given purchase order item, removing it from the PO.
      /// 
      /// URL Parameters:
      ///   purchase order id
      ///   item id
      /// Query Parameters:
      ///   TP.Object.PurchaseOrderItem => serialized object with current data
      /// Returns:
      ///   serialized SuccessResponse or ErrorResponse
      Post["/{id:int}/itemlines/remove/{itemid:int}"] = _ =>
      {
        DbPurchaseOrderItem item = NancyExtensions.GetDbObjectParameter<DbPurchaseOrderItem>(Request.Form, "TP.Object.PurchaseOrderItem");
        if (item.HeaderID != _.id)
        {
          throw new ArgumentException("url id does not match given object!");
        }
        if (item.ItemID != _.itemid)
        {
          throw new ArgumentException("url id does not match given object!");
        }
        if (item.RemovePurchaseOrderItem())
        {
          return NancyExtensions.Respond(new TP.Base.SuccessResponse("purchase order item removed"));
        }
        else
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse("delete failed"));
        }
      };

      #endregion


      #region DbPurchaseOrderTracking - Read

      /// Returns tracking information for a purchase order's items. Currently
      /// only WMH/JPW dropship orders, but more could be added in the future.
      /// 
      /// URL Parameters:
      ///   purchase order id
      /// Query Parameters:
      ///   preload:
      ///     purchaseorder
      /// Returns:
      ///   <ArrayOfPurchaseOrderTracking>
      ///     <PurchaseOrderTracking>
      ///       <TupleID>620</TupleID>
      ///       <HeaderID>32349</HeaderID>
      ///       <TrackingID>4719970</TrackingID>
      ///       <Carrier>SOUTHEASTERN FREIGHT LINES INC</Carrier>
      ///       <IsCustomerNotified>false</IsCustomerNotified>
      ///     </PurchaseOrderTracking>
      ///     ...
      ///   </ArrayOfPurchaseOrderTracking>
      Get["/{id:int}/tracking"] = _ =>
      {
        try
        {
          IEnumerable<DbPurchaseOrderTracking> list = DbPurchaseOrderTracking.GetByPurchaseOrderID(_.id, NancyExtensions.GetPreloads(Request.Query));
          return NancyExtensions.Respond(list);
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      #endregion


      #region DbPurchaseOrderTracking - Create

      /// Returns tracking information for a purchase order's items. Currently
      /// only WMH/JPW dropship orders, but more could be added in the future.
      /// 
      /// URL Parameters:
      ///   purchase order id
      /// Query Parameters:
      ///   TrackingID
      ///   Carrier
      ///   CustomerNotified, defaults to false (valid values are "[Ff]alse" and "[Tt]rue", looks like 1/0 doesn't parse)
      /// Returns:
      ///   <PurchaseOrderTracking>
      ///     <TupleID>620</TupleID>
      ///     <HeaderID>32349</HeaderID>
      ///     <TrackingID>4719970</TrackingID>
      ///     <Carrier>SOUTHEASTERN FREIGHT LINES INC</Carrier>
      ///     <IsCustomerNotified>false</IsCustomerNotified>
      ///   </PurchaseOrderTracking>
      Post["/{id:int}/tracking/add"] = _ =>
      {
        try
        {
          DbPurchaseOrderTracking item = DbPurchaseOrderTracking.AddPackageToPurchaseOrder(
              poID: _.id,
              trackingID: NancyExtensions.GetBasicParameter<string>(Request.Form, "TrackingID"),
              carrier: NancyExtensions.GetBasicParameter<string>(Request.Form, "Carrier"),
              customerNotified: NancyExtensions.GetBasicParameter<bool>(Request.Form, "CustomerNotified", false)
            );
          return NancyExtensions.Respond(item);
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      #endregion


      #region DbPurchaseOrderTracking - Update

      /// Modifies the tracking info for the given package line.
      /// 
      /// URL Parameters:
      ///   purchase order id
      ///   tracking id
      /// Query Parameters:
      ///   TP.Object.PurchaseOrderTracking => serialized object with new data
      /// Returns:
      ///   <PurchaseOrderTracking>
      ///     <TupleID>620</TupleID>
      ///     <HeaderID>32349</HeaderID>
      ///     <TrackingID>4719970</TrackingID>
      ///     <Carrier>SOUTHEASTERN FREIGHT LINES INC</Carrier>
      ///     <IsCustomerNotified>true</IsCustomerNotified>
      ///   </PurchaseOrderTracking>
      Post["/{id:int}/tracking/update/{tid:int}"] = _ =>
      {
        try
        {
          DbPurchaseOrderTracking item = NancyExtensions.GetDbObjectParameter<DbPurchaseOrderTracking>(Request.Form, "TP.Object.PurchaseOrderTracking");
          if (item.HeaderID != _.id)
          {
            throw new ArgumentException("url id does not match given object!");
          }
          if (item.TupleID != _.tid)
          {
            throw new ArgumentException("url id does not match given object!");
          }
          return NancyExtensions.Respond(item.UpdatePurchaseOrderTracking());
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      #endregion

    }
  }

  // MasObjectCRUD handles:
  //   OBJECT                           READ  CREATE  UPDATE  DELETE
  //   DbVendor                           Y     Y       Y       N
  //   DbPurchaseTerms                    Y     Y       N       N
  //   DbPurchaseOrderShipToAddress       Y     Y       N       N
  //   DbMasWarehouse                     Y     Y       N       N
  public class MasObjectCRUD : NancyModule
  {
    public MasObjectCRUD() : base() {

      #region DbVendor - Read

      /// Returns the complete vendor information, specified by the Mas vendor
      /// number.
      /// 
      /// URL Parameters:
      ///   vendor number
      /// Query Parameters:
      ///   preload:
      ///     terms
      /// Returns:
      ///   <Vendor>
      ///     <VendorNo>BLAC01</VendorNo>
      ///     <Name>Black &amp; Decker (U.S) Inc</Name>
      ///     <AddressLine1>P.O. Box 223516</AddressLine1>
      ///     ...
      ///   </Vendor>
      Get["/vendor/{vendorno:length(3,7)}"] = _ =>
      {
        try
        {
          DbVendor item = DbVendor.GetByID(_.vendorno, NancyExtensions.GetPreloads(Request.Query));
          return NancyExtensions.Respond(item);
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      #endregion


      #region DbVendor - Create

      /// Creates a brand new vendor, really just the stub only. Other fields
      /// will be filled in with an update. Otherwise, ready for import into
      /// Mas. Requires the vendor number and vendor name.
      /// 
      /// URL Parameters:
      ///   none
      /// Query Parameters:
      ///   VendorNo: the ^[-_A-Z0-9]{1,7}$ vendor code
      ///   VendorName: vendor name, max 30 chars
      /// Returns:
      ///   <Vendor>
      ///     <VendorNo>BLAC01</VendorNo>
      ///     <Name>Black &amp; Decker (U.S) Inc</Name>
      ///     <AddressLine1>P.O. Box 223516</AddressLine1>
      ///     ...
      ///   </Vendor>
      Post["/vendor/create"] = _ =>
      {
        try
        {
          DbVendor item = DbVendor.CreateVendor(
              vendorNo: NancyExtensions.GetBasicParameter<string>(Request.Form, "VendorNo"),
              vendorName: NancyExtensions.GetBasicParameter<string>(Request.Form, "VendorName")
            );
          return NancyExtensions.Respond(item);
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      #endregion


      #region DbVendor - Update

      /// Modifies an existing vendor, given a new serialized version of the
      /// vendor object. Returns the updated, current version of the object in
      /// the database.
      /// 
      /// URL Parameters:
      ///   vendor number
      /// Query Parameters:
      ///   TP.Object.Vendor => serialized object with new data
      /// Returns:
      ///   <Vendor>
      ///     <VendorNo>BLAC01</VendorNo>
      ///     <Name>Black &amp; Decker (U.S) Inc</Name>
      ///     <AddressLine1>P.O. Box 223516</AddressLine1>
      ///     ...
      ///   </Vendor>
      Post["/vendor/update/{vendorno:length(3,7)}"] = _ =>
      {
        try
        {
          DbVendor item = NancyExtensions.GetDbObjectParameter<DbVendor>(Request.Form, "TP.Object.Vendor");
          if (item.VendorNo != _.vendorno)
          {
            throw new ArgumentException("url id does not match given object!");
          }
          return NancyExtensions.Respond(item.UpdateVendor());
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      #endregion
    

      #region DbPurchaseTerms - Read

      /// Returns a purchase term object with all information relating to a
      /// purchase order's terms, or the default terms associated with a
      /// vendor.
      /// 
      /// URL Parameters:
      ///   product line id
      /// Query Parameters:
      ///   none
      /// Returns:
      ///   <PurchaseTerms>
      ///     <Code>00</Code>
      ///     <Description>"SPECIAL TERMS" *SEE COMMENTS*</Description>
      ///     <DaysBeforeDue>0</DaysBeforeDue>
      ///     ...
      ///   </PurchaseTerms>
      Get["/purchaseterms/{code:length(2)}"] = _ =>
      {
        try
        {
          DbPurchaseTerms item = DbPurchaseTerms.GetByID(_.code, NancyExtensions.GetPreloads(Request.Query));
          return NancyExtensions.Respond(item);
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      #endregion


      #region DbPurchaseTerms - Create

      /// Creates a brand new purchase term, full information. Note that there
      /// is no update here (why?), so everything should be correct this time
      /// through, no fixing.
      /// 
      /// URL Parameters:
      ///   none
      /// Query Parameters:
      ///   Code: the 2 digit (or letter) terms code
      ///   Description: text to identify, this is pretty important to get correct as far as readability is concerned
      ///   DaysBeforeDue: days to pay invoice, integer
      ///   MinAllowedDays: can't pay the invoice before this time? integer
      ///   DiscountAllowedDays: days to pay and receive a discount (see rate), integer
      ///   MinDiscountAllowedDays: can't pay and receive a discount before this time? integer
      ///   Rate: the percentage discount rate, decimal (but probably an integer value)
      /// Returns:
      ///   <PurchaseTerms>
      ///     <Code>00</Code>
      ///     <Description>"SPECIAL TERMS" *SEE COMMENTS*</Description>
      ///     <DaysBeforeDue>0</DaysBeforeDue>
      ///     ...
      ///   </PurchaseTerms>
      Post["/purchaseterms/create"] = _ =>
      {
        try
        {
          DbPurchaseTerms item = DbPurchaseTerms.CreateNewPurchaseTerms(
              code: NancyExtensions.GetBasicParameter<string>(Request.Form, "Code"),
              desc: NancyExtensions.GetBasicParameter<string>(Request.Form, "Description"),
              daysBeforeDue: NancyExtensions.GetBasicParameter<int>(Request.Form, "DaysBeforeDue"),
              minAllowedDays: NancyExtensions.GetBasicParameter<int>(Request.Form, "MinAllowedDays"),
              discAllowedDays: NancyExtensions.GetBasicParameter<int>(Request.Form, "DiscountAllowedDays"),
              minDiscAllowedDays: NancyExtensions.GetBasicParameter<int>(Request.Form, "MinDiscountAllowedDays"),
              rate: NancyExtensions.GetBasicParameter<decimal>(Request.Form, "Rate")
            );
          return NancyExtensions.Respond(item);
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      #endregion


      #region DbPurchaseOrderShipToAddress - Read

      /// Returns a purchase order ship-to address.
      /// 
      /// URL Parameters:
      ///   ship-to code
      /// Query Parameters:
      ///   none
      /// Returns:
      ///   <PurchaseOrderShipToAddress>
      ///     <ShipToCode>0003</ShipToCode>
      ///     <ShipToCodeDesc>Scott Rd</ShipToCodeDesc>
      ///     <ShipToAddress1>60 Scott Rd</ShipToAddress1>
      ///     ...
      ///   </PurchaseOrderShipToAddress>
      Get["/shipto/{code:length(4)}"] = _ =>
      {
        try
        {
          DbPurchaseOrderShipToAddress item = DbPurchaseOrderShipToAddress.GetByID(_.code, NancyExtensions.GetPreloads(Request.Query));
          return NancyExtensions.Respond(item);
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      #endregion


      #region DbPurchaseOrderShipToAddress - Create

      /// Create a new purchase order ship-to address, ready for exporting to
      /// Mas.
      /// 
      /// URL Parameters:
      ///   none
      /// Query Parameters:
      ///   ShipToCode
      ///   ShipToCodeDesc, optional all the way down?
      ///   ShipToAddress1
      ///   ShipToAddress2
      ///   ShipToAddress3
      ///   ShipToCity
      ///   ShipToState
      ///   ShipToPostalCode
      ///   ShipToCountryCode
      /// Returns:
      ///   <PurchaseOrderShipToAddress>
      ///     <ShipToCode>0003</ShipToCode>
      ///     <ShipToCodeDesc>Scott Rd</ShipToCodeDesc>
      ///     <ShipToAddress1>60 Scott Rd</ShipToAddress1>
      ///     ...
      ///   </PurchaseOrderShipToAddress>
      Post["/shipto/create"] = _ =>
      {
        try
        {
          DbPurchaseOrderShipToAddress item = DbPurchaseOrderShipToAddress.CreateShipToAddress(
              code: NancyExtensions.GetBasicParameter<string>(Request.Form, "ShipToCode"),
              desc: NancyExtensions.GetBasicParameter<string>(Request.Form, "ShipToCodeDesc"),
              name: NancyExtensions.GetBasicParameter<string>(Request.Form, "ShipToName"),
              addr1: NancyExtensions.GetBasicParameter<string>(Request.Form, "ShipToAddress1"),
              addr2: NancyExtensions.GetBasicParameter<string>(Request.Form, "ShipToAddress2"),
              addr3: NancyExtensions.GetBasicParameter<string>(Request.Form, "ShipToAddress3"),
              city: NancyExtensions.GetBasicParameter<string>(Request.Form, "ShipToCity"),
              state: NancyExtensions.GetBasicParameter<string>(Request.Form, "ShipToState"),
              postalCode: NancyExtensions.GetBasicParameter<string>(Request.Form, "ShipToPostalCode"),
              countryCode: NancyExtensions.GetBasicParameter<string>(Request.Form, "ShipToCountryCode")
            );
          return NancyExtensions.Respond(item);
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      #endregion


      #region DbMasWarehouse - Read

      /// Returns mas warehouse information. Note that only part of the table
      /// is mirrored.
      /// 
      /// URL Parameters:
      ///   warehouse code
      /// Query Parameters:
      ///   none
      /// Returns:
      ///   <MasWarehouse>
      ///     <WarehouseCode>000</WarehouseCode>
      ///     <Description>MAIN WAREHOUSE</Description>
      ///     <Name>Tools Plus, Inc</Name>
      ///     <Address1>153 Meadow St</Address1>
      ///     ...
      ///   </MasWarehouse>
      Get["/maswarehouse/{code:length(3)}"] = _ =>
      {
        try
        {
          DbMasWarehouse item = DbMasWarehouse.GetByCode(_.code, NancyExtensions.GetPreloads(Request.Query));
          return NancyExtensions.Respond(item);
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      #endregion


      #region DbMasWarehouse - Create

      /// Creates a new mas warehouse. Probably won't use this too often.
      /// 
      /// URL Parameters:
      ///   none
      /// Query Parameters:
      ///   Code, required
      ///   Description, optional here and below
      ///   Name
      ///   Address1
      ///   Address2
      ///   Address3
      ///   City
      ///   State
      ///   PostalCode
      ///   CountryCode
      /// Returns:
      ///   <MasWarehouse>
      ///     <WarehouseCode>000</WarehouseCode>
      ///     <Description>MAIN WAREHOUSE</Description>
      ///     <Name>Tools Plus, Inc</Name>
      ///     <Address1>153 Meadow St</Address1>
      ///     ...
      ///   </MasWarehouse>
      Post["/maswarehouse/create"] = _ =>
      {
        try
        {
          DbMasWarehouse item = DbMasWarehouse.CreateNewMasWarehouse(
              code: NancyExtensions.GetBasicParameter<string>(Request.Form, "Code"),
              desc: NancyExtensions.GetBasicParameter<string>(Request.Form, "Description"),
              name: NancyExtensions.GetBasicParameter<string>(Request.Form, "Name"),
              addr1: NancyExtensions.GetBasicParameter<string>(Request.Form, "Address1"),
              addr2: NancyExtensions.GetBasicParameter<string>(Request.Form, "Address2"),
              addr3: NancyExtensions.GetBasicParameter<string>(Request.Form, "Address3"),
              city: NancyExtensions.GetBasicParameter<string>(Request.Form, "City"),
              state: NancyExtensions.GetBasicParameter<string>(Request.Form, "State"),
              postalCode: NancyExtensions.GetBasicParameter<string>(Request.Form, "PostalCode"),
              countryCode: NancyExtensions.GetBasicParameter<string>(Request.Form, "CountryCode")
            );
          return NancyExtensions.Respond(item);
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      #endregion
    }
  }

  // InventoryLocationCRUD handles:
  //   OBJECT                           READ  CREATE  UPDATE  DELETE
  //   DbInventoryLocation                Y     N       N       N
  //   DbInventoryLocationContent         Y     Y       Y       Y
  public class InventoryLocationCRUD : NancyModule
  {
    public InventoryLocationCRUD() : base("/invlocation/") {

      #region DbInventoryLocation - Read

      /// Returns the location object with the given id.
      /// 
      /// URL Parameters:
      ///   location id
      /// Query Parameters:
      ///   preload:
      ///     contents
      /// Returns:
      ///   <InventoryLocation>
      ///     <LocationID>8529</LocationID>
      ///     <LocationTypeID>2</LocationTypeID>
      ///     <WarehouseID>5</WarehouseID>
      ///     <Division1>1</Division1>
      ///     ...
      ///   </InventoryLocation>
      Get["/{id:int}"] = _ =>
      {
        try
        {
          DbInventoryLocation item = DbInventoryLocation.GetByID(_.id, NancyExtensions.GetPreloads(Request.Query));
          return NancyExtensions.Respond(item);
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      /// Returns the location object with the given warehouse barcode. Format
      /// is ^LOC\d{9}$ where the digits are the zero-padded location id.
      /// 
      /// URL Parameters:
      ///   location id
      /// Query Parameters:
      ///   preload:
      ///     contents
      /// Returns:
      ///   <InventoryLocation>
      ///     <LocationID>8529</LocationID>
      ///     <LocationTypeID>2</LocationTypeID>
      ///     <WarehouseID>5</WarehouseID>
      ///     <Division1>1</Division1>
      ///     ...
      ///   </InventoryLocation>
      Get["/{bc}"] = _ =>
      {
        try
        {
          DbInventoryLocation item = DbInventoryLocation.GetByBarcode(_.bc, NancyExtensions.GetPreloads(Request.Query));
          return NancyExtensions.Respond(item);
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      // XXX: deleteme -- LocationName part of serialized InventoryLocation
      Get["/{id:int}/name"] = _ =>
      {
        try
        {
          DbInventoryLocation item = DbInventoryLocation.GetByID(_.id, NancyExtensions.GetPreloads(Request.Query));
          if (item == null) { return NancyExtensions.Respond(item); }
          return NancyExtensions.Respond(new TP.Base.SuccessResponse(string.Format("{0} {1}", item.LocationTypeName, item.LocationName)));
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };
      // XXX: deleteme -- LocationName part of serialized InventoryLocation
      Get["/{id:int}/name/short"] = _ =>
      {
        try
        {
          DbInventoryLocation item = DbInventoryLocation.GetByID(_.id, NancyExtensions.GetPreloads(Request.Query));
          if (item == null) { return NancyExtensions.Respond(item); }
          return NancyExtensions.Respond(new TP.Base.SuccessResponse(string.Format("{0}", item.LocationName)));
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      Get["/{id:int}/span/endpoint"] = _ =>
      {
        try
        {
          DbInventoryLocation item = DbInventoryLocation.GetSpanEndpoint(_.id, NancyExtensions.GetPreloads(Request.Query));
          return NancyExtensions.Respond(item);
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      Get["/{id:int}/span/to/{id2:int}"] = _ =>
      {
        try
        {
          if (DbInventoryLocation.IsSpanValid(_.id, _.id2))
          {
            return NancyExtensions.Respond(new TP.Base.SuccessResponse("spannable"));
          }
          else
          {
            return NancyExtensions.Respond(new TP.Base.ErrorResponse("not spannable"));
          }
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      #endregion


      #region DbInventoryLocationContent - Read

      /// Returns a list of InventoryLocationContent objects showing the
      /// contents of the given location.
      /// 
      /// URL Parameters:
      ///   location id
      /// Query Parameters:
      ///   preload:
      ///     location
      ///     component
      /// Returns:
      ///   <ArrayOfInventoryLocationContent>
      ///     <InventoryLocationContent>
      ///       <LocationID>8651</LocationID>
      ///       <ComponentID>31736</ComponentID>
      ///       <MasterID>8651</MasterID>
      ///       <Quantity>8></Quantity>
      ///       <AssociatedTaskType i:nil="true"/>
      ///       <AssociatedTaskID i:nil="true"/>
      ///     </InventoryLocationContent>
      ///     ...
      ///   </ArrayOfInventoryLocationContent>
      Get["/{id:int}/contents"] = _ =>
      {
        try
        {
          IEnumerable<DbInventoryLocationContent> list = DbInventoryLocationContent.GetByLocationID(_.id, NancyExtensions.GetPreloads(Request.Query));
          return NancyExtensions.Respond(list);
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      Get["/{id:int}/contents/inventoryrequired/limit/{limit:int}/skip/{skip:int}"] = _ =>
      {
        try
        {
          IEnumerable<DbInventoryLocationContent> list = DbInventoryLocationContent.GetByLocationID(_.id, NancyExtensions.GetPreloads(Request.Query));
          list = list.OrderBy(c => c.LastInventoriedDate).Skip((int)_.skip).Take((int)_.limit);
          return NancyExtensions.Respond(list);
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      #endregion


      #region DbInventoryLocationContent - Create/Update/Delete (Merged)

      /// Adds or removes an item from a location in the warehouse, without an
      /// associated task. E.g., receiving.
      /// 
      /// URL Parameters:
      ///   component id
      ///   location id
      /// Query Parameters:
      ///   LocationStart        - id of beginning of location span
      ///   LocationEnd          - id of end of location span
      ///   ComponentID          - component
      ///   DeltaQuantity        - quantity change, must be positive
      ///   TransactionReference - optional, use to describe reason for movement
      ///   Comment              - optional, extended info about movement
      ///   NTUsername           - optional but probably required, who is doing this
      ///   ComputerName         - optional but probably required, native computer
      ///   RDPClientName        - optional but probably required
      /// Returns:
      ///   serialized SuccessResponse or ErrorResponse
      Post["/{id:int}/contents/{cid:int}/change"] = _ =>
      {
        try
        {
          int ls = NancyExtensions.GetBasicParameter<int>(Request.Form, "LocationStart");
          int cid = NancyExtensions.GetBasicParameter<int>(Request.Form, "ComponentID");
          if (ls != _.id)
          {
            throw new ArgumentException("url id does not match given request body!");
          }
          if (cid != _.cid)
          {
            throw new ArgumentException("url id does not match given request body!");
          }
          bool ret = DbInventoryLocationContent.AddOrRemove(
              startLocationID: ls,
              endLocationID: NancyExtensions.GetBasicParameter<int>(Request.Form, "LocationEnd"),
              componentID: cid,
              deltaQuantity: NancyExtensions.GetBasicParameter<int>(Request.Form, "DeltaQuantity"),
              transactionReference: NancyExtensions.GetBasicParameter<string>(Request.Form, "TransactionReference"),
              comment: NancyExtensions.GetBasicParameter<string>(Request.Form, "Comment"),
              personNTUsername: NancyExtensions.GetBasicParameter<string>(Request.Form, "NTUsername"),
              computerName: NancyExtensions.GetBasicParameter<string>(Request.Form, "ComputerName"),
              rdpClientName: NancyExtensions.GetBasicParameter<string>(Request.Form, "RDPClientName")
            );
          if (ret)
          {
            return NancyExtensions.Respond(new TP.Base.SuccessResponse("success"));
          }
          else
          {
            return NancyExtensions.Respond(new TP.Base.ErrorResponse("this won't ever happen"));
          }
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      /// Moves a component from one location to another within the warehouse,
      /// without an associated task. E.g., putting, consolidation, etc.
      /// 
      /// URL Parameters:
      ///   component id
      ///   from location id
      /// Query Parameters:
      ///   FromLocationStart    - id of beginning of from location span
      ///   FromLocationEnd      - id of end of from location span
      ///   ToLocationStart      - id of beginning of to location span
      ///   ToLocationEnd        - id of end of to location span
      ///   ComponentID          - component
      ///   DeltaQuantity        - quantity change, must be positive
      ///   TransactionReference - optional, use to describe reason for movement
      ///   Comment              - optional, extended info about movement
      ///   NTUsername           - optional but probably required, who is doing this
      ///   ComputerName         - optional but probably required, native computer
      ///   RDPClientName        - optional but probably required
      /// Returns:
      ///   serialized SuccessResponse or ErrorResponse
      Post["/{id:int}/contents/{cid:int}/move"] = _ =>
      {
        try
        {
          int fls = NancyExtensions.GetBasicParameter<int>(Request.Form, "FromLocationStart");
          int cid = NancyExtensions.GetBasicParameter<int>(Request.Form, "ComponentID");
          if (fls != _.id)
          {
            throw new ArgumentException("url id does not match given request body!");
          }
          if (cid != _.cid)
          {
            throw new ArgumentException("url id does not match given request body!");
          }
          bool ret = DbInventoryLocationContent.MoveFromTo(
              fromStartLocationID: fls,
              fromEndLocationID: NancyExtensions.GetBasicParameter<int>(Request.Form, "FromLocationEnd"),
              toStartLocationID: NancyExtensions.GetBasicParameter<int>(Request.Form, "ToLocationStart"),
              toEndLocationID: NancyExtensions.GetBasicParameter<int>(Request.Form, "ToLocationEnd"),
              componentID: cid,
              deltaQuantity: NancyExtensions.GetBasicParameter<int>(Request.Form, "DeltaQuantity"),
              transactionReference: NancyExtensions.GetBasicParameter<string>(Request.Form, "TransactionReference"),
              comment: NancyExtensions.GetBasicParameter<string>(Request.Form, "Comment"),
              personNTUsername: NancyExtensions.GetBasicParameter<string>(Request.Form, "NTUsername"),
              computerName: NancyExtensions.GetBasicParameter<string>(Request.Form, "ComputerName"),
              rdpClientName: NancyExtensions.GetBasicParameter<string>(Request.Form, "RDPClientName")
            );
          if (ret)
          {
            return NancyExtensions.Respond(new TP.Base.SuccessResponse("success"));
          }
          else
          {
            return NancyExtensions.Respond(new TP.Base.ErrorResponse("this won't ever happen"));
          }
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      #endregion
    }
  }

  // TaskCRUD handles:
  //   OBJECT                           READ  CREATE  UPDATE  DELETE
  //   DbSalesOrderTask                   Y     N       N       N
  //   DbSalesOrderTaskLine               Y     N       Y       N
  //   DbTransferTask                     Y     N       N       N
  //   DbTransferTaskLine                 Y     N       Y       N
  //   DbWillCallTask                     Y     N       N       N
  //   DbWillCallTaskLine                 Y     N       Y       N
  public class TaskCRUD : NancyModule
  {
    public TaskCRUD() : base("/task")
    {
      #region Db(SalesOrder|Transfer|WillCall)Task - Read

      /// Returns the task object with the given id.
      /// 
      /// URL Parameters:
      ///   type - string "salesorder", "transfer", "willcall"
      ///   task id
      /// Query Parameters:
      ///   preload:
      ///     lines -- applicable to all
      /// Returns:
      ///   serialized one of <SalesOrderTask>, <TransferTask>, <WillCallTask>
      ///   
      ///   <SalesOrderTask>
      ///     <SalesOrderTaskID>164299</SalesOrderTaskID>
      ///     <WarehouseID>5</WarehouseID>
      ///     <SalesOrderNo>0067782</SalesOrderNo>
      ///     <CurrentStatus>6</CurrentStatus>
      ///     <CurrentHoldReason>1</CurrentHoldReason>
      ///     <ShipVia>UPS NEXTDAY AIR</ShipVia>
      ///     ...
      ///   </SalesOrderTask>
      ///   
      ///   <TransferTask>
      ///     <TransferTaskID>11652</TransferTaskID>
      ///     <WarehouseID>5</WarehouseID>
      ///     <InitialDate>2014-01-22T00:00:00</InitialDate>
      ///     <SequenceNo>8</SequenceNo>
      ///     <TransferNo>T140122H</TransferNo>
      ///     <DestinationWarehouseID>1</DestinationWarehouseID>
      ///     ...
      ///   </TransferTask>
      ///   
      ///   <WillCallTask>
      ///     <WillCallTaskID>2402</WillCallTaskID>
      ///     <WarehouseID>5</WarehouseID>
      ///     <MASTransactionID>0221739</MASTransactionID>
      ///     <CreatedByUserID>72</CreatedByUserID>
      ///     <TimeInserted>2014-04-10T11:00:40.157</TimeInserted>
      ///   </WillCallTask>
      Get["/{type}/id/{id:int}"] = _ =>
      {
        try
        {
          switch ((string)_.type)
          {
            case "salesorder":
              {
                DbSalesOrderTask item = DbSalesOrderTask.GetByID(_.id, NancyExtensions.GetPreloads(Request.Query));
                return NancyExtensions.Respond(item);
              }
            case "transfer":
              {
                DbTransferTask item = DbTransferTask.GetByID(_.id, NancyExtensions.GetPreloads(Request.Query));
                return NancyExtensions.Respond(item);
              }
            case "willcall":
              {
                DbWillCallTask item = DbWillCallTask.GetByID(_.id, NancyExtensions.GetPreloads(Request.Query));
                return NancyExtensions.Respond(item);
              }
            default:
              throw new ArgumentException("invalid type in url");
          }
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      /// Returns the task object with the given friendly transaction
      /// identifier (sales order number, transfer number, will call
      /// transaction number).
      /// 
      /// URL Parameters:
      ///   type - string "salesorder", "transfer", "willcall"
      ///   task identifier
      /// Query Parameters:
      ///   preload:
      ///     lines -- applicable to all
      /// Returns:
      ///   serialized one of <SalesOrderTask>, <TransferTask>, <WillCallTask>, see above
      Get["/{type}/no/{number}"] = _ =>
      {
        try
        {
          switch ((string)_.type)
          {
            case "salesorder":
              {
                DbSalesOrderTask item = DbSalesOrderTask.GetBySalesOrderNo(_.number, NancyExtensions.GetPreloads(Request.Query));
                return NancyExtensions.Respond(item);
              }
            case "transfer":
              {
                DbTransferTask item = DbTransferTask.GetByTransferNo(_.number, NancyExtensions.GetPreloads(Request.Query));
                return NancyExtensions.Respond(item);
              }
            case "willcall":
              {
                DbWillCallTask item = DbWillCallTask.GetByMASTransactionNo(_.number, NancyExtensions.GetPreloads(Request.Query));
                return NancyExtensions.Respond(item);
              }
            default:
              throw new ArgumentException("invalid type in url");
          }
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      #endregion


      #region Db(SalesOrder|Transfer|WillCall)TaskLine - Read

      /// Returns the lines associated with the given task.
      /// 
      /// URL Parameters:
      ///   type - string "salesorder", "transfer", "willcall"
      ///   task id
      /// Query Parameters:
      ///   preload:
      ///     component -- applicable to all
      /// Returns:
      ///   serialized one of <SalesOrderTaskLine>, <TransferTaskLine>, <WillCallTaskLine>
      ///   
      ///   <ArrayOfSalesOrderTaskLine>
      ///     <SalesOrderTaskLine>
      ///       <TupleID>252223</TupleID>
      ///       <SalesOrderTaskID>164299</SalesOrderTaskID>
      ///       <ComponentID>25383</ComponentID>
      ///       <QuantityOrdered>2</QuantityOrdered>
      ///       <QuantityBackordered>0</QuantityBackordered>
      ///       <QuantityPicked>2</QuantityPicked>
      ///       <QuantityPacked>2</QuantityPacked>
      ///       <QuantityShipped>2</QuantityShipped>
      ///     </SalesOrderTaskLine>
      ///     ...
      ///   </ArrayOfSalesOrderTaskLine>
      ///   
      ///   <ArrayOfTransferTaskLine>
      ///     <TransferTaskLine>
      ///       <TupleID>108490</TupleID>
      ///       <TransferTaskID>11652</TransferTaskID>
      ///       <ComponentID>2979</ComponentID>
      ///       <QuantityRequested>1</QuantityRequested>
      ///       <QuantityPicked>1</QuantityPicked>
      ///       <QuantityPacked>1</QuantityPacked>
      ///       <QuantityTransported>1</QuantityTransported>
      ///       <QuantityFilled>1</QuantityFilled>
      ///     </TransferTaskLine>
      ///     ...
      ///   </ArrayOfTransferTaskLine>
      ///   
      ///   <ArrayOfWillCallTaskLine>
      ///     <WillCallTaskLine>
      ///       <TupleID>2937</TupleID>
      ///       <WillCallTaskID>2402</WillCallTaskID>
      ///       <ComponentID>34757</ComponentID>
      ///       <QuantityRequired>1</QuantityRequired>
      ///       <QuantityPicked>1</QuantityPicked>
      ///       <QuantityPacked>1</QuantityPacked>
      ///       <QuantityFilled>1</QuantityFilled>
      ///     </WillCallTaskLine>
      ///     ...
      ///   </ArrayOfWillCallTaskLine>
      Get["/{type}/id/{id:int}/lines"] = _ =>
      {
        try
        {
          switch ((string)_.type)
          {
            case "salesorder":
              {
                IEnumerable<DbSalesOrderTaskLine> list = DbSalesOrderTaskLine.GetByTaskID(_.id, NancyExtensions.GetPreloads(Request.Query));
                return NancyExtensions.Respond(list);
              }
            case "transfer":
              {
                IEnumerable<DbTransferTaskLine> list = DbTransferTaskLine.GetByTaskID(_.id, NancyExtensions.GetPreloads(Request.Query));
                return NancyExtensions.Respond(list);
              }
            case "willcall":
              {
                IEnumerable<DbWillCallTaskLine> list = DbWillCallTaskLine.GetByTaskID(_.id, NancyExtensions.GetPreloads(Request.Query));
                return NancyExtensions.Respond(list);
              }
            default:
              throw new ArgumentException("invalid type in url");
          }
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };
      Get["/{type}/no/{number}/lines"] = _ =>
      {
        try
        {
          switch ((string)_.type)
          {
            case "salesorder":
              {
                DbSalesOrderTask header = DbSalesOrderTask.GetBySalesOrderNo(_.number, NancyExtensions.GetPreloads(Request.Query));
                IEnumerable<DbSalesOrderTaskLine> list = DbSalesOrderTaskLine.GetByTaskID(header.SalesOrderTaskID, NancyExtensions.GetPreloads(Request.Query));
                return NancyExtensions.Respond(list);
              }
            case "transfer":
              {
                DbTransferTask header = DbTransferTask.GetByTransferNo(_.number, NancyExtensions.GetPreloads(Request.Query));
                IEnumerable<DbTransferTaskLine> list = DbTransferTaskLine.GetByTaskID(header.TransferTaskID, NancyExtensions.GetPreloads(Request.Query));
                return NancyExtensions.Respond(list);
              }
            case "willcall":
              {
                DbWillCallTask header = DbWillCallTask.GetByMASTransactionNo(_.number, NancyExtensions.GetPreloads(Request.Query));
                IEnumerable<DbWillCallTaskLine> list = DbWillCallTaskLine.GetByTaskID(header.WillCallTaskID, NancyExtensions.GetPreloads(Request.Query));
                return NancyExtensions.Respond(list);
              }
            default:
              throw new ArgumentException("invalid type in url");
          }
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      /// Returns a list of task lines that still have outstanding picking to
      /// do. Note that this returns ONLY the outstanding lines for each
      /// order, lines where picking has been fulfilled are not included, so
      /// we can't assume 1 line here means the component is solo on the
      /// order.
      /// NOTE: this assumes warehouse 5 (scott rd), which is dumb.
      /// 
      /// URL Parameters:
      ///   type - string "salesorder", "transfer", "willcall"
      ///   task id
      /// Query Parameters:
      ///   none
      /// Returns:
      ///   serialized array of one of <SalesOrderTaskLine>, <TransferTaskLine>, <WillCallTaskLine>, see above
      Get["/{type}/picking"] = _ =>
      {
        try
        {
          switch ((string)_.type)
          {
            case "salesorder":
              {
                IEnumerable<DbSalesOrderTaskLine> list = DbSalesOrderTaskLine.GetPickingRequired();
                return NancyExtensions.Respond(list);
              }
            case "transfer":
              {
                IEnumerable<DbTransferTaskLine> list = DbTransferTaskLine.GetPickingRequired();
                return NancyExtensions.Respond(list);
              }
            case "willcall":
              {
                IEnumerable<DbWillCallTaskLine> list = DbWillCallTaskLine.GetPickingRequired();
                return NancyExtensions.Respond(list);
              }
            default:
              throw new ArgumentException("invalid type in url");
          }
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      /// Returns a list of task lines that still have outstanding picking to
      /// do. Note that this returns ONLY the outstanding lines for each
      /// order, lines where picking has been fulfilled are not included, so
      /// we can't assume 1 line here means the component is solo on the
      /// order.
      /// NOTE: lots of preloads are automatically included here to assist
      /// with the filtering required, currently:
      ///   task header
      ///   component
      ///   locationcontent
      ///   location
      /// 
      /// URL Parameters:
      ///   type - string "salesorder", "transfer", "willcall"
      ///   task id
      /// Query Parameters:
      ///   none
      /// Returns:
      ///   serialized array of one of <SalesOrderTaskLine>, <TransferTaskLine>, <WillCallTaskLine>, see above
      Get["/{type}/picking/warehouse/{wid:int}/locationtype/{ltid:int}"] = _ =>
      {
        try
        {
          switch ((string)_.type)
          {
            case "salesorder":
              {
                IEnumerable<DbSalesOrderTaskLine> list = DbSalesOrderTaskLine.GetPickingRequired(_.wid, new string[] { "component", "locations" });
                Dictionary<int, object> incompletable = new Dictionary<int, object>();
                foreach (var tl in list)
                {
                  if (incompletable.ContainsKey(tl.HeaderID) == false && tl.QuantityLeftToPick > tl.Component.Locations.Sum(ilc => ilc.Quantity ?? 0))
                  {
                    incompletable.Add(tl.HeaderID, null);
                  }
                }
                IEnumerable<DbSalesOrderTaskLine> retval = list.Where(tl => incompletable.ContainsKey(tl.HeaderID) == false && tl.Component.Locations.Any(ilc => ilc.InventoryLocation.WarehouseID == _.wid && ilc.InventoryLocation.LocationTypeID == _.ltid));
                return NancyExtensions.Respond(retval);
              }
            case "transfer":
              {
                IEnumerable<DbTransferTaskLine> list = DbTransferTaskLine.GetPickingRequired(_.wid, new string[] { "component", "locations" });
                IEnumerable<DbTransferTaskLine> retval = list.Where(tl => tl.Component.Locations.Any(ilc => ilc.InventoryLocation.WarehouseID == _.wid && ilc.InventoryLocation.LocationTypeID == _.ltid));
                return NancyExtensions.Respond(retval);
              }
            case "willcall":
              {
                IEnumerable<DbWillCallTaskLine> list = DbWillCallTaskLine.GetPickingRequired(_.wid, new string[] { "component", "locations" });
                IEnumerable<DbWillCallTaskLine> retval = list.Where(tl => tl.Component.Locations.Any(ilc => ilc.InventoryLocation.WarehouseID == _.wid && ilc.InventoryLocation.LocationTypeID == _.ltid));
                return NancyExtensions.Respond(retval);
              }
            default:
              throw new ArgumentException("invalid type in url");
          }
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      #endregion


      #region Db(SalesOrder|Transfer|WillCall)TaskLine - Update (inventory movement)

      /// Moves a component from one location to another within the warehouse,
      /// associating the move with the given task line. E.g., picking.
      /// 
      /// URL Parameters:
      ///   type - string "salesorder", "transfer", "willcall"
      ///   task id
      ///   task line id
      /// Query Parameters:
      ///   preload:
      ///     lines -- applicable to all
      /// Returns:
      ///   serialized SuccessResponse or ErrorResponse
      Post["/{type}/id/{id:int}/lines/{lineid:int}/move"] = _ =>
      {
        try
        {
          IWarehouseTaskLine tl;
          if (NancyExtensions.HasParameter(Request.Form, "PreviousPickedQuantity"))
          {
            int ppq = NancyExtensions.GetBasicParameter<int>(Request.Form, "PreviousPickedQuantity", -1);
            switch ((string)_.type)
            {
              case "salesorder":
                tl = DbSalesOrderTaskLine.GetByID(_.lineid);
                break;
              case "transfer":
                tl = DbTransferTaskLine.GetByID(_.lineid);
                break;
              case "willcall":
                tl = DbWillCallTaskLine.GetByID(_.lineid);
                break;
              default:
                throw new ArgumentException("invalid type in url");
            }
            if (ppq != tl.QuantityPicked)
            {
              throw new ArgumentException("quantity picked does not match, task line may have been modified!");
            }
          }
          else
          {
            switch ((string)_.type)
            {
              case "salesorder":
                tl = NancyExtensions.GetDbObjectParameter<DbSalesOrderTaskLine>(Request.Form, "TP.Object.SalesOrderTaskLine");
                break;
              case "transfer":
                tl = NancyExtensions.GetDbObjectParameter<DbTransferTaskLine>(Request.Form, "TP.Object.TransferTaskLine");
                break;
              case "willcall":
                tl = NancyExtensions.GetDbObjectParameter<DbWillCallTaskLine>(Request.Form, "TP.Object.WillCallTaskLine");
                break;
              default:
                throw new ArgumentException("invalid type in url");
            }
          }
          if (tl.HeaderID != _.id)
          {
            throw new ArgumentException("url id does not match given request body!");
          }
          if (tl.LineID != _.lineid)
          {
            throw new ArgumentException("url id does not match given request body!");
          }

          bool ret = DbInventoryLocationContent.MoveFromTo(
              fromStartLocationID: NancyExtensions.GetBasicParameter<int>(Request.Form, "FromLocationStart"),
              fromEndLocationID: NancyExtensions.GetBasicParameter<int>(Request.Form, "FromLocationEnd"),
              toStartLocationID: NancyExtensions.GetBasicParameter<int>(Request.Form, "ToLocationStart"),
              toEndLocationID: NancyExtensions.GetBasicParameter<int>(Request.Form, "ToLocationEnd"),
              taskLine: tl,
              deltaQuantity: NancyExtensions.GetBasicParameter<int>(Request.Form, "DeltaQuantity"),
              comment: NancyExtensions.GetBasicParameter<string>(Request.Form, "Comment"),
              personNTUsername: NancyExtensions.GetBasicParameter<string>(Request.Form, "NTUsername"),
              computerName: NancyExtensions.GetBasicParameter<string>(Request.Form, "ComputerName"),
              rdpClientName: NancyExtensions.GetBasicParameter<string>(Request.Form, "RDPClientName")
            );
          if (ret)
          {
            return NancyExtensions.Respond(new TP.Base.SuccessResponse("success"));
          }
          else
          {
            return NancyExtensions.Respond(new TP.Base.ErrorResponse("this won't ever happen"));
          }
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };
      Post["/{type}/no/{number}/lines/{lineid:int}/move"] = _ =>
      {
        try
        {
          IWarehouseTaskLine tl;
          if (NancyExtensions.HasParameter(Request.Form, "PreviousPickedQuantity"))
          {
            int ppq = NancyExtensions.GetBasicParameter<int>(Request.Form, "PreviousPickedQuantity", -1);
            switch ((string)_.type)
            {
              case "salesorder":
                tl = DbSalesOrderTaskLine.GetByID(_.lineid);
                break;
              case "transfer":
                tl = DbTransferTaskLine.GetByID(_.lineid);
                break;
              case "willcall":
                tl = DbWillCallTaskLine.GetByID(_.lineid);
                break;
              default:
                throw new ArgumentException("invalid type in url");
            }
            if (ppq != tl.QuantityPicked)
            {
              throw new ArgumentException("quantity picked does not match, task line may have been modified!");
            }
          }
          else
          {
            switch ((string)_.type)
            {
              case "salesorder":
                tl = NancyExtensions.GetDbObjectParameter<DbSalesOrderTaskLine>(Request.Form, "TP.Object.SalesOrderTaskLine");
                break;
              case "transfer":
                tl = NancyExtensions.GetDbObjectParameter<DbTransferTaskLine>(Request.Form, "TP.Object.TransferTaskLine");
                break;
              case "willcall":
                tl = NancyExtensions.GetDbObjectParameter<DbWillCallTaskLine>(Request.Form, "TP.Object.WillCallTaskLine");
                break;
              default:
                throw new ArgumentException("invalid type in url");
            }
          }
          if (tl.Header(null).TransactionReference != _.number)
          {
            throw new ArgumentException("url id does not match given request body!");
          }
          if (tl.LineID != _.lineid)
          {
            throw new ArgumentException("url id does not match given request body!");
          }

          bool ret = DbInventoryLocationContent.MoveFromTo(
              fromStartLocationID: NancyExtensions.GetBasicParameter<int>(Request.Form, "FromLocationStart"),
              fromEndLocationID: NancyExtensions.GetBasicParameter<int>(Request.Form, "FromLocationEnd"),
              toStartLocationID: NancyExtensions.GetBasicParameter<int>(Request.Form, "ToLocationStart"),
              toEndLocationID: NancyExtensions.GetBasicParameter<int>(Request.Form, "ToLocationEnd"),
              taskLine: tl,
              deltaQuantity: NancyExtensions.GetBasicParameter<int>(Request.Form, "DeltaQuantity"),
              comment: NancyExtensions.GetBasicParameter<string>(Request.Form, "Comment"),
              personNTUsername: NancyExtensions.GetBasicParameter<string>(Request.Form, "NTUsername"),
              computerName: NancyExtensions.GetBasicParameter<string>(Request.Form, "ComputerName"),
              rdpClientName: NancyExtensions.GetBasicParameter<string>(Request.Form, "RDPClientName")
            );
          if (ret)
          {
            return NancyExtensions.Respond(new TP.Base.SuccessResponse("success"));
          }
          else
          {
            return NancyExtensions.Respond(new TP.Base.ErrorResponse("this won't ever happen"));
          }
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      /// Add or remove an item from a warehouse location, associating the
      /// change with a given task line. E.g., willcall fill.
      /// 
      /// URL Parameters:
      ///   type - string "salesorder", "transfer", "willcall"
      ///   task id
      ///   task line id
      /// Query Parameters:
      ///   preload:
      ///     lines -- applicable to all
      /// Returns:
      ///   serialized SuccessResponse or ErrorResponse
      Post["/{type}/id/{id:int}/lines/{lineid:int}/change"] = _ =>
      {
        try
        {
          IWarehouseTaskLine tl;
          switch ((string)_.type)
          {
            case "salesorder":
              tl = NancyExtensions.GetDbObjectParameter<DbSalesOrderTaskLine>(Request.Form, "TP.Object.SalesOrderTaskLine");
              break;
            case "transfer":
              tl = NancyExtensions.GetDbObjectParameter<DbTransferTaskLine>(Request.Form, "TP.Object.TransferTaskLine");
              break;
            case "willcall":
              tl = NancyExtensions.GetDbObjectParameter<DbWillCallTaskLine>(Request.Form, "TP.Object.WillCallTaskLine");
              break;
            default:
              throw new ArgumentException("invalid type in url");
          }
          if (tl.HeaderID != _.id)
          {
            throw new ArgumentException("url id does not match given request body!");
          }
          if (tl.LineID != _.lineid)
          {
            throw new ArgumentException("url id does not match given request body!");
          }

          bool ret = DbInventoryLocationContent.MoveFromTo(
              fromStartLocationID: NancyExtensions.GetBasicParameter<int>(Request.Form, "FromLocationStart"),
              fromEndLocationID: NancyExtensions.GetBasicParameter<int>(Request.Form, "FromLocationEnd"),
              toStartLocationID: NancyExtensions.GetBasicParameter<int>(Request.Form, "ToLocationStart"),
              toEndLocationID: NancyExtensions.GetBasicParameter<int>(Request.Form, "ToLocationEnd"),
              taskLine: tl,
              deltaQuantity: NancyExtensions.GetBasicParameter<int>(Request.Form, "DeltaQuantity"),
              comment: NancyExtensions.GetBasicParameter<string>(Request.Form, "Comment"),
              personNTUsername: NancyExtensions.GetBasicParameter<string>(Request.Form, "NTUsername"),
              computerName: NancyExtensions.GetBasicParameter<string>(Request.Form, "ComputerName"),
              rdpClientName: NancyExtensions.GetBasicParameter<string>(Request.Form, "RDPClientName")
            );
          if (ret)
          {
            return NancyExtensions.Respond(new TP.Base.SuccessResponse("success"));
          }
          else
          {
            return NancyExtensions.Respond(new TP.Base.ErrorResponse("this won't ever happen"));
          }
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };
      Post["/{type}/no/{number}/lines/{lineid:int}/change"] = _ =>
      {
        try
        {
          IWarehouseTaskLine tl;
          switch ((string)_.type)
          {
            case "salesorder":
              tl = NancyExtensions.GetDbObjectParameter<DbSalesOrderTaskLine>(Request.Form, "TP.Object.SalesOrderTaskLine");
              break;
            case "transfer":
              tl = NancyExtensions.GetDbObjectParameter<DbTransferTaskLine>(Request.Form, "TP.Object.TransferTaskLine");
              break;
            case "willcall":
              tl = NancyExtensions.GetDbObjectParameter<DbWillCallTaskLine>(Request.Form, "TP.Object.WillCallTaskLine");
              break;
            default:
              throw new ArgumentException("invalid type in url");
          }
          if (tl.Header(null).TransactionReference != _.number)
          {
            throw new ArgumentException("url id does not match given request body!");
          }
          if (tl.LineID != _.lineid)
          {
            throw new ArgumentException("url id does not match given request body!");
          }

          bool ret = DbInventoryLocationContent.MoveFromTo(
              fromStartLocationID: NancyExtensions.GetBasicParameter<int>(Request.Form, "FromLocationStart"),
              fromEndLocationID: NancyExtensions.GetBasicParameter<int>(Request.Form, "FromLocationEnd"),
              toStartLocationID: NancyExtensions.GetBasicParameter<int>(Request.Form, "ToLocationStart"),
              toEndLocationID: NancyExtensions.GetBasicParameter<int>(Request.Form, "ToLocationEnd"),
              taskLine: tl,
              deltaQuantity: NancyExtensions.GetBasicParameter<int>(Request.Form, "DeltaQuantity"),
              comment: NancyExtensions.GetBasicParameter<string>(Request.Form, "Comment"),
              personNTUsername: NancyExtensions.GetBasicParameter<string>(Request.Form, "NTUsername"),
              computerName: NancyExtensions.GetBasicParameter<string>(Request.Form, "ComputerName"),
              rdpClientName: NancyExtensions.GetBasicParameter<string>(Request.Form, "RDPClientName")
            );
          if (ret)
          {
            return NancyExtensions.Respond(new TP.Base.SuccessResponse("success"));
          }
          else
          {
            return NancyExtensions.Respond(new TP.Base.ErrorResponse("this won't ever happen"));
          }
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new TP.Base.ErrorResponse(ex));
        }
      };

      #endregion

    }
  }
}
