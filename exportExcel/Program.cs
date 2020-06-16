using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using MongoDB.Bson;
using MongoDB.Driver;
using Excel = Microsoft.Office.Interop.Excel;

namespace exportExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            var connectionString = new MongoUrl("mongodb://admin:admin123@ds239009.mlab.com:39009/heroku_0g0g5g6c?replicaSet=rs-ds239009&retryWrites=false");
            MongoClient dbClient = new MongoClient(connectionString);
            //MongoClient dbClient = new MongoClient();
            var database = dbClient.GetDatabase("heroku_0g0g5g6c");
            //var database = dbClient.GetDatabase("NNNdb");
            var cart = database.GetCollection<BsonDocument>("cart");
            var product_colletion = database.GetCollection<BsonDocument>("stock");
            List<string> header = new List<string>();
            var pk = "";
            var file = "";
            var pList = product_colletion.Find(new BsonDocument()).Limit(1).Sort(Builders<BsonDocument>.Sort.Descending("_id")).ToList();
            foreach (var d in pList)
            {
                pk = d["primaryKey"].ToString();
                file = d["filename"].ToString();
                var text = d["header"].ToString();
                text = text.Substring(1, text.Length - 2);
                header = text.Replace(", ", ",").Split(',').ToList();
            }

            var indexPK = header.IndexOf(pk) + 1;
            var indexQty = 0;
            foreach (var h in header)
            {
                if (h.ToLower().Contains("quantity"))
                {
                    indexQty = header.IndexOf(h) + 1;
                }
            }

            var filter = Builders<BsonDocument>.Filter.Eq("status", "confirmed");
            var update = Builders<BsonDocument>.Update.Set("status", "shipping");
            var document = cart.Find(filter).ToList();
            if (!document.Any())
            {
                Console.WriteLine("No order(s) confirmed yet...");
                Console.WriteLine("Saving " + file.Substring(file.LastIndexOf('/') + 1) + "...");
                save_excel(file, indexPK, indexQty);
            }
            else
            {
                List<Dictionary<string, object>> aproduct = new List<Dictionary<string, object>>();
                Dictionary<string, int> product = new Dictionary<string, int>();
                foreach (BsonDocument doc in document)
                {
                    var test = doc["carts"].ToBsonDocument();
                    aproduct.Add(test.ToDictionary());
                }

                foreach (var d in aproduct)
                {
                    foreach (var k in d)
                    {
                        if (product.ContainsKey(k.Key))
                        {
                            product[k.Key] += (int)k.Value;
                        }
                        else
                        {
                            product.Add(k.Key, (int)k.Value);
                        }

                    }
                }
                Console.WriteLine("Generating Excel File...Please Wait Patiently...");
                update_excel(file, indexPK, indexQty, product);
                cart.UpdateMany(filter, update);
            }
        }

        public static void update_excel(string file, int primary, int qty, Dictionary<string, int> dict)
        {
            try

            {
                Excel.Application excel = new Excel.Application();
                Excel.Workbook sheet = excel.Workbooks.Open(file);
                Excel.Worksheet x = excel.Sheets["GBD_Asia"] as Excel.Worksheet;
                Excel.Range userrange = x.UsedRange;
                int countRecords = userrange.Rows.Count;
                for (int row = 2; row <= countRecords; row++)
                {
                    foreach (var item in dict)
                    {
                        if ((string)(x.Cells[row, primary] as Excel.Range).Value == item.Key)
                        {
                            x.Cells[row, qty] = item.Value;
                        }
                    }
                }
                Excel.Range rng = x.Cells[2, primary] as Excel.Range;
                rng.Select();

                GC.Collect();
                GC.WaitForPendingFinalizers();
                Marshal.ReleaseComObject(userrange);
                Marshal.ReleaseComObject(x);

                sheet.Close(true, Type.Missing, Type.Missing);
                Marshal.ReleaseComObject(sheet);
                excel.Quit();
                Marshal.ReleaseComObject(excel);

            }
            catch (Exception exHandle)

            {

                Console.WriteLine("Exception: " + exHandle.Message);

                Console.ReadLine();

            }
}

        public static void save_excel(string file, int primary, int qty)
        {
            try

            {
                Excel.Application excel = new Excel.Application();
                Excel.Workbook sheet = excel.Workbooks.Open(file);
                Excel.Worksheet x = excel.Sheets["GBD_Asia"] as Excel.Worksheet;
                Excel.Range userrange = x.UsedRange;

                x.Cells[2, qty] = 0;
                x.Cells[2, qty] = "";

                Excel.Range rng = x.Cells[2, primary] as Excel.Range;
                rng.Select();

                GC.Collect();
                GC.WaitForPendingFinalizers();
                Marshal.ReleaseComObject(userrange);
                Marshal.ReleaseComObject(x);
                sheet.Close(true, Type.Missing, Type.Missing);
                Marshal.ReleaseComObject(sheet);
                excel.Quit();
                Marshal.ReleaseComObject(excel);
            }
            catch (Exception exHandle)

            {

                Console.WriteLine("Exception: " + exHandle.Message);

                Console.ReadLine();

            }
        }

    }
}
