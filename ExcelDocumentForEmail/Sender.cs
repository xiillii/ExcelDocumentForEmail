using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Office2010.ExcelAc;
using ExcelDocumentForEmail.Entities;
using IronXL;
using IronXL.Options;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using System.Data;
using System.Linq;
using System.Xml.Serialization;

namespace ExcelDocumentForEmail
{
    public static class Sender
    {
        [FunctionName("ExcelSender")]
        public static async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Function, "get", "post", Route = null)] HttpRequest req,
            ILogger log, ExecutionContext context)
        {
            var config = new ConfigurationBuilder()
                .SetBasePath(context.FunctionAppDirectory)
                .AddJsonFile("local.settings.json", optional: true, reloadOnChange: true)
                .AddEnvironmentVariables()
                .AddUserSecrets(Assembly.GetExecutingAssembly(), true)
                .Build();

            License.LicenseKey = config.GetSection("Licenses:IronXL").Value ?? string.Empty;

            var products = await GetProductsFromDatabase(config.GetConnectionString("AzureTest"));

            var document = CreateExcelDocument(products);


            return new OkObjectResult(new { success = true, file = document });
        }


        private static async Task<List<Product>> GetProductsFromDatabase(string connectionString)
        {
            var products = new List<Product>();

            await using var conn = new SqlConnection(connectionString);
            await conn.OpenAsync();

            var sql =
                "select ProductId, Name, ProductNumber, Color, StandardCost, ListPrice, Size, Weight, ProductCategoryId, ProductModelId, SellStartDate, SellEndDate, DiscontinuedDate, ThumbNailPhoto, ThumbNailPhotoFileName, Rowguid, ModifiedDate from [saleslt].[Product]";

            try
            {
                await using var cmd = new SqlCommand(sql, conn);
                var reader = await cmd.ExecuteReaderAsync();
                while (reader.Read())
                {
                    products.Add(new Product
                    {
                        ProductId = reader.GetInt32(0),
                        Name = reader.GetString(1),
                        ProductNumber = reader.GetString(2),
                        Color = reader.GetString(3),
                        StandardCost = reader.GetDecimal(4),
                    });
                }
            }
            catch (Exception e)
            {
                Debug.WriteLine(e);
            }
            finally
            {
                await conn.CloseAsync();
            }

            

            return products;
        }

        private static string CreateExcelDocument(List<Product> products)
        {
            var result = "";

            try
            {
                var excelFile = WorkBook.Load(ToDataSet(products),
                        new CreatingOptions { DefaultFileFormat = ExcelFileFormat.XLSX });

                var bytes = excelFile.ToByteArray();
                result = Convert.ToBase64String(bytes);
            }
            catch (Exception e)
            {

                Debug.WriteLine(e);
            }

            return result;
        }

        private static DataSet ToDataSet(List<Product> list)
        {
            var elementType = typeof(Product);
            var ds = ToDataSetFromArray(list.ToArray());

            var selectedColumns = new[] { "ProductId", "Name", "ProductNumber", "Color", "StandardCost" };
            var t = new DataView(ds.Tables[0]).ToTable(false, selectedColumns);

            var rowHeader = t.NewRow();

            //Agregar columnas con los nombres y tipo de las propiedades al DataSet
            foreach (var propInfo in elementType.GetProperties())
            {
                var ColType = Nullable.GetUnderlyingType(propInfo.PropertyType) ?? propInfo.PropertyType;

                //Agregar los nombres de los campos como parte de la primera fila
                //ya que la libería IronXL no permite indicar considerar los header de las tablas del dataset
                switch (propInfo.Name)
                {
                    case "ApplicationName":
                        rowHeader[propInfo.Name] = "Description";
                        break;
                    case "NotificationTypeName":
                        rowHeader[propInfo.Name] = "Type";
                        break;
                    case "SendDate":
                        rowHeader[propInfo.Name] = "Send date";
                        break;
                    case "StatusName":
                        rowHeader[propInfo.Name] = "Status";
                        break;
                    case "TotalEmail":
                        rowHeader[propInfo.Name] = "Email sent";
                        break;
                    case "CountryName":
                        rowHeader[propInfo.Name] = "Country";
                        break;
                }
            }

            t.Rows.InsertAt(rowHeader, 0);

            t.AcceptChanges();
            var dsRet = new DataSet();
            dsRet.Tables.Add(t);

            return dsRet;
        }

        private static DataSet ToDataSetFromArray(object[] listDto)
        {
            var result = new DataSet();
            var xmlSerializer = new XmlSerializer(listDto.GetType());
            var sw = new StringWriter();
            xmlSerializer.Serialize(sw, listDto);
            var reader = new StringReader(sw.ToString());
            result.ReadXml(reader);

            return result;
        }
    }
}
