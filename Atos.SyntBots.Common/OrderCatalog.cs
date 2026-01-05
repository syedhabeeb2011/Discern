/*
 * Created by SharpDevelop.
 * User: v.chakradhar.ark
 * Date: 11/4/2019
 * Time: 5:43 AM
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Collections.Generic;

namespace Atos.SyntBots.Common
{
    /// <summary>
    /// Description of OrderCatalog.
    /// </summary>
    public class OrderCatalog
    {
        public OrderCatalog()
        {
        }

        public DataTable getOrderCatalogdata(string fileName)
        {
            //string cs = @"provider=Microsoft.ACE.OLEDB.12.0;Data Source='"+fileName+"';Extended Properties=Excel 12.0;";
            string cs = @"provider=Microsoft.Jet.OLEDB.4.0;Data Source='" + fileName + "';Extended Properties=Excel 8.0;";
            DataTable dt = new DataTable();
            OleDbConnection con = new OleDbConnection(cs);
            OleDbCommand cmd = new OleDbCommand("Select * from [Master Sheet$]", con);
            OleDbDataAdapter oda = new OleDbDataAdapter(cmd);
            oda.Fill(dt);
            return dt;
        }

        public string GetPowerChartChargeDescription(DataTable dt, string accessHIMChargeDescription)
        {
            string catalog_display = (from DataRow dr in dt.Rows
                                      where (string)dr["CHARGE_DESCRIPTION"] == accessHIMChargeDescription.Trim()
                                      select (string)dr["CATALOG_DISPLAY"]).FirstOrDefault();

            return catalog_display;
        }

        public string GetAccessHIMChargeDescription(DataTable dt, string powerchartChargeDescription)
        {
            string catalog_display = (from DataRow dr in dt.Rows
                                      where (string)dr["CATALOG_DISPLAY"] == powerchartChargeDescription.Trim()
                                      select (string)dr["CHARGE_DESCRIPTION"]).FirstOrDefault();

            if (string.IsNullOrEmpty(catalog_display))
            {
                catalog_display = (from DataRow dr in dt.Rows
                                   where (string)dr["CATALOG_DISPLAY2"] == powerchartChargeDescription.Trim()
                                   select (string)dr["CHARGE_DESCRIPTION"]).FirstOrDefault();

            }
            return catalog_display;
        }

        public IList<string> GetAllAccessHIMChargeDescriptionsForCPT(DataTable dt, string cptCode)
        {
            var catalog_descriptionList = (from DataRow dr in dt.Rows
                                           where (string)dr["CPT_Code"] == cptCode
                                           select (string)dr["CHARGE_DESCRIPTION"]).ToList();

            return catalog_descriptionList;
        }

        public IList<string> GetAllAccessHIMChargeDescriptionsForCPT(DataTable dt, IList<string> cptCodes)
        {
            var query = dt.AsEnumerable().
                                    Select(item => new
                                    {
                                        CPTCode = item.Field<string>("CPT_CODE"),
                                        CatalogDisplay = item.Field<string>("CATALOG_DISPLAY"),
                                        ChargeDescription = item.Field<string>("CHARGE_DESCRIPTION"),
                                        HCPCSCode = item.Field<string>("HCPCS_CODE"),
                                    });

            var query1 = cptCodes.Join(query, x => x, y => y.CPTCode, (x, y) => y.ChargeDescription).ToList();

            if (query1.Count == 0)
            {
                query1 = cptCodes.Join(query, x => x, y => y.HCPCSCode, (x, y) => y.ChargeDescription).ToList();
            }

            return query1;
        }

        public IList<string> GetAllPowerChartChargeDescriptionsForCPT(DataTable dt, string cptCode)
        {
            var catalog_descriptionList = (from DataRow dr in dt.Rows
                                           where (string)dr["CPT_Code"] == cptCode
                                           select (string)dr["CATALOG_DISPLAY"]).ToList();

            return catalog_descriptionList;
        }

        public IList<string> GetAllCPTForOrderDescription(DataTable dt, string chargedescription)
        {

            var cptCodesList = (from DataRow dr in dt.Rows
                                where (string)dr["CATALOG_DISPLAY"] == chargedescription.Trim()
                                select (string)dr["CPT_CODE"]).ToList();

            if (cptCodesList.Count == 0)
            {
                cptCodesList = (from DataRow dr in dt.Rows
                                where (string)dr["CATALOG_DISPLAY2"] == chargedescription.Trim()
                                select (string)dr["CPT_CODE"]).ToList();
            }

            if (cptCodesList.Count == 0)
            {
                cptCodesList = (from DataRow dr in dt.Rows
                                where (string)dr["CATALOG_DISPLAY"] == chargedescription.Trim()
                                select (string)dr["HCPCS_CODE"]).ToList();
            }

            if (cptCodesList.Count == 0)
            {
                cptCodesList = (from DataRow dr in dt.Rows
                                where (string)dr["CATALOG_DISPLAY2"] == chargedescription.Trim()
                                select (string)dr["HCPCS_CODE"]).ToList();
            }

            return cptCodesList;

        }

    }
}
