using ExcelDna.Integration;
using MarketplaceWebServiceProducts;
using MarketplaceWebServiceProducts.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GetProduct
{
    public class GetProduct
    {
        [ExcelFunction(Description = "Get Product Information by ASIN.")]
        public static void GetMatchingProduct()
        {
            // アクティブシート変更
            XlCall.Excel(XlCall.xlcWorkbookActivate, "Amazon MWS アカウント設定");

            Product getProduct = null;
            string ASIN = string.Empty;
            string SellerId = new ExcelReference(4, 2).GetValue().ToString().Trim();
            string MarketplaceId = new ExcelReference(5, 2).GetValue().ToString().Trim();
            string MWSAuthToken = new ExcelReference(6, 2).GetValue().ToString().Trim();
            string AccessKeyId = "AccessKey";
            string SecretKeyId = "SecretKey";
            string ApplicationVersion = "1.0.0.0";
            string ApplicationName = "AB_Soft";
            string strWidth = string.Empty;
            string strHeight = string.Empty;
            string strLength = string.Empty;
            string strWeight = string.Empty;
            string strSize = string.Empty;

            MarketplaceWebServiceProductsConfig config = new MarketplaceWebServiceProductsConfig();
            config.ServiceURL = "https://mws.amazonservices.jp";

            MarketplaceWebServiceProductsClient client = new MarketplaceWebServiceProductsClient(
                                                                ApplicationName,
                                                                ApplicationVersion,
                                                                AccessKeyId,
                                                                SecretKeyId,
                                                                config);
            GetMatchingProductRequest request = new GetMatchingProductRequest();
            request.SellerId = SellerId;
            request.MarketplaceId = MarketplaceId;

            // アクティブシート変更
            XlCall.Excel(XlCall.xlcWorkbookActivate, "商品検索");
            ASIN = new ExcelReference(3, 2).GetValue().ToString().Trim();
            ASINListType asinListType = new ASINListType();
            asinListType.ASIN.Add(ASIN);
            request.ASINList = asinListType;
            request.MWSAuthToken = MWSAuthToken;

            GetMatchingProductResponse response = client.GetMatchingProduct(request);
            if (response.IsSetGetMatchingProductResult())
            {
                List<GetMatchingProductResult> getMatchingProductResultList = response.GetMatchingProductResult;
                getProduct = getMatchingProductResultList[0].Product;
                System.Xml.XmlElement elements = (System.Xml.XmlElement)getProduct.AttributeSets.Any[0];
                foreach (System.Xml.XmlElement element in elements)
                {
                    switch (element.LocalName)
                    {
                        case "Title":                                                                           // Title
                            XlCall.Excel(XlCall.xlcFormula, element.InnerText, new ExcelReference(6, 2));
                            break;
                        case "BandMaterialType":                                                                // BandMaterialType
                            XlCall.Excel(XlCall.xlcFormula, element.InnerText, new ExcelReference(7, 2));
                            break;
                        case "Binding":                                                                         // Binding
                            XlCall.Excel(XlCall.xlcFormula, element.InnerText, new ExcelReference(7, 4));
                            break;
                        case "Brand":                                                                           // Brand
                            XlCall.Excel(XlCall.xlcFormula, element.InnerText, new ExcelReference(8, 2));
                            break;
                        case "ClaspType":                                                                       // ClaspType
                            XlCall.Excel(XlCall.xlcFormula, element.InnerText, new ExcelReference(8, 4));
                            break;
                        case "Color":                                                                           // Color
                            XlCall.Excel(XlCall.xlcFormula, element.InnerText, new ExcelReference(9, 2));
                            break;
                        case "Label":                                                                           // Label
                            XlCall.Excel(XlCall.xlcFormula, element.InnerText, new ExcelReference(10, 2));
                            break;
                        case "ListPrice":                                                                    // CurrencyCode
                            XlCall.Excel(XlCall.xlcFormula, element.ChildNodes[0].InnerText, new ExcelReference(10, 4));
                            break;
                        case "Manufacturer":                                                                    // Manufacturer
                            XlCall.Excel(XlCall.xlcFormula, element.InnerText, new ExcelReference(11, 2));
                            break;
                        case "Model":                                                                           // Model
                            XlCall.Excel(XlCall.xlcFormula, element.InnerText, new ExcelReference(11, 4));
                            break;
                        case "PackageDimensions":                                                               // PackageDimensions
                            if (element.ChildNodes[0] != null)
                            {
                                strWidth = element.ChildNodes[0].InnerText;
                            }
                            else
                            {
                                strWidth = "0";
                            }
                            if (element.ChildNodes[1] != null)
                            {
                                strHeight = element.ChildNodes[0].InnerText;
                            }
                            else
                            {
                                strHeight = "0";
                            }
                            if (element.ChildNodes[2] != null)
                            {
                                strLength = element.ChildNodes[2].InnerText;
                            }
                            else
                            {
                                strLength = "0";
                            }
                            if (element.ChildNodes[3] != null)
                            {
                                strWeight = element.ChildNodes[3].InnerText;
                            }
                            else
                            {
                                strWeight = "";
                            }
                            strSize = strWidth + " x " + strHeight + " x " + strLength;
                            break;
                        case "PackageQuantity":                                                                 // PackageQuantity
                            XlCall.Excel(XlCall.xlcFormula, element.InnerText, new ExcelReference(12, 4));
                            break;
                        case "PartNumber":                                                                      // PartNumber
                            XlCall.Excel(XlCall.xlcFormula, element.InnerText, new ExcelReference(13, 2));
                            break;
                        case "ProductGroup":                                                                    // ProductGroup
                            XlCall.Excel(XlCall.xlcFormula, element.InnerText, new ExcelReference(13, 4));
                            break;
                        case "ProductTypeName":                                                                 // ProductTypeName
                            XlCall.Excel(XlCall.xlcFormula, element.InnerText, new ExcelReference(14, 2));
                            break;
                        case "Publisher":                                                                       // Publisher
                            XlCall.Excel(XlCall.xlcFormula, element.InnerText, new ExcelReference(14, 4));
                            break;
                        case "ReleaseDate":                                                                     // ReleaseDate
                            XlCall.Excel(XlCall.xlcFormula, element.InnerText, new ExcelReference(15, 2));
                            break;
                    }
                }
                // サイズ有りの場合設定
                if (strSize != "")
                {
                    XlCall.Excel(XlCall.xlcFormula, strSize, new ExcelReference(12, 2));
                }
                // 重さありの場合設定
                if (strWeight != "")
                {
                    XlCall.Excel(XlCall.xlcFormula, strWeight, new ExcelReference(9, 4));
                }
                // ランキング設定
                XlCall.Excel(XlCall.xlcFormula, getProduct.SalesRankings.SalesRank[0].Rank, new ExcelReference(15, 4));
            }
        }
    }
}
