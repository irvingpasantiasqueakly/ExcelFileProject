using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Data.OleDb;
using System.Data;
using System.Configuration;
using System.Data.SqlClient;
using excelFileProject.Models;
using System.IO;

namespace ExcelProjectFile.Controllers
{
    public class HomeController : Controller
    {
        SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["con"].ConnectionString);
        OleDbConnection Econ;

        public ActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public ActionResult Index(HttpPostedFileBase file)
        {
            string filename = Guid.NewGuid() + Path.GetExtension(file.FileName);
            string filepath = "/excelFolder/" + filename;
            file.SaveAs(Path.Combine(Server.MapPath("/excelFolder"), filename));

            string ext = Path.GetExtension(filename.ToString());

            if (ext.Equals(".xlsx") || ext.Equals(".xls"))
            {
                if (InsertExcelData(filepath, filename))
                    return View("Succesful");
                else
                    return View("NoChanges");
            }
            else
                return View("Error");

        }

        private Boolean InsertExcelData(string filepath, string filename)
        {
            string fullpath = Server.MapPath("/excelFolder/") + filename;
            ExcelConn(fullpath);
            string queryExcel = string.Format("SELECT * FROM [{0}]", "Sheet1$");
            Boolean insert = false;

            if (InsertRegion(queryExcel))
            { insert = true; }

            if (InsertCurrency(queryExcel))
            { insert = true; }

            if (InsertfoundSymbol(queryExcel))
            { insert = true; }

            if (InsertCountry(queryExcel))
            { insert = true; }

            if (InsertMaster(queryExcel))
            { insert = true; }

            return insert;
        }

        /*METHODS*/

        private void ExcelConn(string filepath)
        {
            string constr = string.Format(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='Excel 12.0 Xml;HDR=YES;'", filepath);
            Econ = new OleDbConnection(constr);
        }

        private List<string> verifyList(List<string> Excel, List<string> Database)
        {
            var list3 = Excel.Except(Database);
            var list4 = Database.Except(Excel);
            var resultList = list3.Concat(list4).ToList();

            return resultList;
        }

        /* GET DATA FROM DATABASE*/

        private List<string> RegionListDB()
        {
            List<string> regionList = new List<string>();
            string query = "SELECT regions FROM dbo.region";

            conn.Open();
            SqlCommand Econ = new SqlCommand(query, conn);
            SqlDataReader dr = Econ.ExecuteReader();
            try
            {
                while (dr.Read())
                {
                    regionList.Add(dr.GetString(0).Trim().ToString());
                }
            }
            catch (Exception e)
            {
                ViewBag.Error = "ERROR: " + e;
            }
            finally
            {
                conn.Close();
            }
            return regionList;
        }

        private List<string> CurrencyListDB()
        {
            List<string> currencyList = new List<string>();
            string query = "SELECT [currencies] FROM dbo.[currency]";

            conn.Open();
            SqlCommand Econ = new SqlCommand(query, conn);
            SqlDataReader dr = Econ.ExecuteReader();
            try
            {
                while (dr.Read())
                {
                    currencyList.Add(dr.GetString(0).Trim().ToString());
                }
            }
            catch (Exception e)
            {
                ViewBag.Error = "ERROR: " + e;
            }
            finally
            {
                conn.Close();
            }
            return currencyList;
        }

        private List<string> FundSymbolListDB()
        {
            List<string> fundSymbolList = new List<string>();
            string query = "SELECT fundSymbols FROM dbo.fundSymbol";

            conn.Open();
            SqlCommand Econ = new SqlCommand(query, conn);
            SqlDataReader dr = Econ.ExecuteReader();
            try
            {
                while (dr.Read())
                {
                    fundSymbolList.Add(dr.GetString(0).Trim().ToString());
                }
            }
            catch (Exception e)
            {
                ViewBag.Error = "ERROR: " + e;
            }
            finally
            {
                conn.Close();
            }
            return fundSymbolList;
        }

        private int RegionIDDB(string region)
        {
            string query = "SELECT id FROM dbo.region WHERE regions = '" + region + "';";
            int id = 0;

            conn.Open();
            SqlCommand Econ = new SqlCommand(query, conn);
            SqlDataReader dr = Econ.ExecuteReader();
            try
            {
                while (dr.Read())
                {
                    id = dr.GetInt32(0);
                }
            }
            catch (Exception e)
            {
                ViewBag.Error = "ERROR: " + e;
            }
            finally
            {
                conn.Close();
            }
            return id;
        }

        private int CurrencyIDDB(string currency)
        {
            string query = "SELECT [id] FROM dbo.[currency] WHERE currencies = '" + currency + "';";
            int id = 0;

            conn.Open();
            SqlCommand Econ = new SqlCommand(query, conn);
            SqlDataReader dr = Econ.ExecuteReader();
            try
            {
                while (dr.Read())
                {
                    id = dr.GetInt32(0);
                }
            }
            catch (Exception e)
            {
                ViewBag.Error = "ERROR: " + e;
            }
            finally
            {
                conn.Close();
            }
            return id;
        }

        private int FundSymbolIDDB(string fundSymbol)
        {
            string query = "SELECT [id] FROM [excelFileProject].[dbo].[fundSymbol] WHERE [fundSymbols] = '" + fundSymbol + "';";
            int id = 0;

            conn.Open();
            SqlCommand Econ = new SqlCommand(query, conn);
            SqlDataReader dr = Econ.ExecuteReader();
            try
            {
                while (dr.Read())
                {
                    id = dr.GetInt32(0);
                }
            }
            catch (Exception e)
            {
                ViewBag.Error = "ERROR: " + e;
            }
            finally
            {
                conn.Close();
            }
            return id;
        }

        private int CountryIDDB(string country, int fundSymbol_id)
        {
            string query = "SELECT [id] FROM [excelFileProject].[dbo].[country] WHERE [countries] = '" + country + "' AND [fundSymbol_id] = " + fundSymbol_id + ";";
            int id = 0;

            conn.Open();
            SqlCommand Econ = new SqlCommand(query, conn);
            SqlDataReader dr = Econ.ExecuteReader();
            try
            {
                while (dr.Read())
                {
                    id = dr.GetInt32(0);
                }
            }
            catch (Exception e)
            {
                ViewBag.Error = "ERROR: " + e;
            }
            finally
            {
                conn.Close();
            }
            return id;
        }

        private int ExistCountryDB(string countries, string countryCodes, int region_id, int currency_id, int fundSymbol_id)
        {
            string query = "SELECT id FROM [dbo].[country] WHERE [countries] = '" + countries + "' AND [countryCodes] = '" + countryCodes + "' AND [region_id] = " + region_id + " AND [currency_id] = " + currency_id + " AND [fundSymbol_id] = " + fundSymbol_id + ";";
            int id = 0;

            conn.Open();
            SqlCommand Econ = new SqlCommand(query, conn);
            SqlDataReader dr = Econ.ExecuteReader();
            try
            {
                while (dr.Read())
                {
                    id = dr.GetInt32(0);
                }
            }
            catch (Exception e)
            {
                ViewBag.Error = "ERROR: " + e;
            }
            finally
            {
                conn.Close();
            }
            return id;
        }

        private int ExistMasterDB(DateTime date, string countries, string countryCodes, string foundSymbol)
        {
            string query = "SELECT M.id,c.countries,c.countryCodes,f.fundSymbols FROM [excelFileProject].[dbo].[master] AS M"
                + " JOIN [excelFileProject].[dbo].[country] AS C ON M.country_id = C.id"
                + " JOIN [excelFileProject].[dbo].[fundSymbol] AS f ON c.fundSymbol_id = f.id"
                + " WHERE M.[date] = '" + date + "' AND C.countries = '" + countries + "' AND C.countryCodes = '" + countryCodes + "' AND F.fundSymbols = '" + foundSymbol + "';";
            int id = 0;

            conn.Open();
            SqlCommand Econ = new SqlCommand(query, conn);
            SqlDataReader dr = Econ.ExecuteReader();
            try
            {
                while (dr.Read())
                {
                    id = dr.GetInt32(0);
                }
            }
            catch (Exception e)
            {
                ViewBag.Error = "ERROR: " + e;
            }
            finally
            {
                conn.Close();
            }
            return id;
        }

        /*INSERT DATA IN DATABASE*/

        private Boolean InsertRegion(string queryExcel)
        {
            Boolean insertRegion = false;
            List<string> regionListExcel = new List<string>();
            List<string> regionList = new List<string>();

            Econ.Open();
            OleDbCommand Ecom = new OleDbCommand(queryExcel, Econ);
            OleDbDataReader dr = Ecom.ExecuteReader();
            try
            {
                while (dr.Read())
                {
                    regionListExcel.Add(dr.GetString(2).Trim().ToString());
                }
            }
            catch (Exception e)
            {
                ViewBag.Error = "ERROR: " + e;
            }
            finally
            {
                Econ.Close();
            }

            regionList = verifyList(regionListExcel, RegionListDB());

            if (regionList.Count > 0)
            {
                using (excelFileProjectEntities ctx = new excelFileProjectEntities())
                {
                    try
                    {
                        foreach (var list in regionList)
                        {
                            excelFileProject.Models.region region = new excelFileProject.Models.region()
                            {
                                regions = list.Trim().ToString()
                            };
                            ctx.regions.Add(region);
                            ctx.SaveChanges();
                            insertRegion = true;
                        }
                    }
                    catch (Exception e)
                    {
                        ViewBag.Error = "ERORR: " + e;
                        regionListExcel.Clear();
                        regionList.Clear();
                    }
                };
            }
            return insertRegion;
        }

        private Boolean InsertCurrency(string queryExcel)
        {
            Boolean insert = false;
            List<string> currencyListExcel = new List<string>();
            List<string> currencyList = new List<string>();

            Econ.Open();
            OleDbCommand Ecom = new OleDbCommand(queryExcel, Econ);
            OleDbDataReader dr = Ecom.ExecuteReader();
            try
            {
                while (dr.Read())
                {
                    currencyListExcel.Add(dr.GetString(5).Trim().ToString());
                }
            }
            catch (Exception e)
            {
                ViewBag.Error = "ERROR: " + e;
            }
            finally
            {
                Econ.Close();
            }

            currencyList = verifyList(currencyListExcel, CurrencyListDB());

            if (currencyList.Count > 0)
            {
                using (excelFileProjectEntities ctx = new excelFileProjectEntities())
                {
                    try
                    {
                        foreach (var list in currencyList)
                        {
                            excelFileProject.Models.currency currency = new excelFileProject.Models.currency()
                            {
                                currencies = list.Trim().ToString()
                            };
                            ctx.currencies.Add(currency);
                            ctx.SaveChanges();
                            insert = true;
                        }
                    }
                    catch (Exception e)
                    {
                        ViewBag.Error = "ERORR: " + e;
                    }
                };
            }
            return insert;
        }

        private Boolean InsertfoundSymbol(string queryExcel)
        {
            Boolean insert = false;
            List<string> fundSymbolListExcel = new List<string>();
            List<string> fundSymbolList = new List<string>();

            Econ.Open();
            OleDbCommand Ecom = new OleDbCommand(queryExcel, Econ);
            OleDbDataReader dr = Ecom.ExecuteReader();
            try
            {
                while (dr.Read())
                {
                    fundSymbolListExcel.Add(dr.GetString(1).Trim().ToString());
                }
            }
            catch (Exception e)
            {
                ViewBag.Error = "ERROR: " + e;
            }
            finally
            {
                Econ.Close();
            }

            fundSymbolList = verifyList(fundSymbolListExcel, FundSymbolListDB());

            if (fundSymbolList.Count > 0)
            {
                using (excelFileProjectEntities ctx = new excelFileProjectEntities())
                {
                    try
                    {
                        foreach (var list in fundSymbolList)
                        {
                            excelFileProject.Models.fundSymbol fundSymbol = new excelFileProject.Models.fundSymbol()
                            {
                                fundSymbols = list.Trim().ToString()
                            };
                            ctx.fundSymbols.Add(fundSymbol);
                            ctx.SaveChanges();
                            insert = true;
                        }
                    }
                    catch (Exception e)
                    {
                        ViewBag.Error = "ERORR: " + e;
                    }
                };
            }
            return insert;
        }

        private Boolean InsertCountry(string queryExcel)
        {
            Boolean insert = false;

            Econ.Open();
            OleDbCommand Ecom = new OleDbCommand(queryExcel, Econ);
            OleDbDataReader dr = Ecom.ExecuteReader();
            try
            {
                while (dr.Read())
                {
                    string cou = (dr.GetString(3).Trim().ToString());
                    string couCode = (dr.GetString(4).Trim().ToString());
                    int reg_id = RegionIDDB(dr.GetString(2).Trim().ToString());
                    int curr_id = CurrencyIDDB(dr.GetString(5).Trim().ToString());
                    int fundSymb_id = FundSymbolIDDB(dr.GetString(1).Trim().ToString());

                    int id = ExistCountryDB(cou, couCode, reg_id, curr_id, fundSymb_id);
                    if (id == 0)
                    {
                        using (excelFileProjectEntities ctx = new excelFileProjectEntities())
                        {
                            try
                            {
                                excelFileProject.Models.country countrya = new excelFileProject.Models.country()
                                {
                                    countries = cou.Trim().ToString(),
                                    countryCodes = couCode.Trim().ToString(),
                                    region_id = reg_id,
                                    currency_id = curr_id,
                                    fundSymbol_id = fundSymb_id
                                };
                                ctx.countries.Add(countrya);
                                ctx.SaveChanges();
                                insert = true;
                            }
                            catch (Exception e)
                            {
                                ViewBag.Error = "ERORR: " + e;
                            }
                        };
                    }
                }
            }
            catch (Exception e)
            {
                ViewBag.Error = "ERROR: " + e;
            }
            finally
            {
                Econ.Close();
            }
            return insert;
        }

        private Boolean InsertMaster(string queryExcel)
        {
            Boolean insert = false;

            Econ.Open();
            OleDbCommand Ecom = new OleDbCommand(queryExcel, Econ);
            OleDbDataReader dr = Ecom.ExecuteReader();
            try
            {
                while (dr.Read())
                {
                    DateTime dat = (dr.GetDateTime(0));
                    string fundSymbol = (dr.GetString(1).Trim().ToString());
                    string cou = (dr.GetString(3).Trim().ToString());
                    string couCod = (dr.GetString(4).Trim().ToString());
                    double portfolioWeighta = (dr.GetDouble(6));
                    double effectiveDurationa = (dr.GetDouble(7));
                    double embeddedIncomeYielda = (dr.GetDouble(8));
                    double yieldToMaturitya = (dr.GetDouble(9));
                    int cou_id = CountryIDDB(cou, FundSymbolIDDB(fundSymbol));

                    int id = ExistMasterDB(dat, cou, couCod, fundSymbol);
                    if (id == 0)
                    {
                        using (excelFileProjectEntities ctx = new excelFileProjectEntities())
                        {
                            try
                            {
                                excelFileProject.Models.master mastera = new excelFileProject.Models.master()
                                {
                                    date = dat,
                                    country_id = cou_id,
                                    portfolioWeight = portfolioWeighta,
                                    effectiveDuration = effectiveDurationa,
                                    embeddedIncomeYield = embeddedIncomeYielda,
                                    yieldToMaturity = yieldToMaturitya

                                };
                                ctx.masters.Add(mastera);
                                ctx.SaveChanges();
                                insert = true;
                            }
                            catch (Exception e)
                            {
                                ViewBag.Error = "ERORR: " + e;
                            }
                        };
                    }
                }
            }
            catch (Exception e)
            {
                ViewBag.Error = "ERROR: " + e;
            }
            finally
            {
                Econ.Close();
            }
            return insert;
        }
    }
}