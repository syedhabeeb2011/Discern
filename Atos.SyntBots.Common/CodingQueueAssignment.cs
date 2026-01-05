/*
 * Created by SharpDevelop.
 * User: v.chakradhar.ark
 * Date: 11/6/2019
 * Time: 9:29 AM
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using System.Data;
using System.Data.OleDb;
using System.Linq;

namespace Atos.SyntBots.Common
{
    /// <summary>
    /// Description of CodingQueueAssignment.
    /// </summary>
    public class CodingQueueAssignment
    {
        public CodingQueueAssignment()
        {
        }

        public DataTable GetCodingQueueAssignmentData(string fileName)
        {
            string cs = @"provider=Microsoft.ACE.OLEDB.12.0;Data Source='" + fileName + "';Extended Properties=Excel 12.0;";
            DataTable EmployeeTable = new DataTable();
            OleDbConnection con = new OleDbConnection(cs);
            OleDbCommand cmd = new OleDbCommand("Select * from [Sheet1$]", con);
            OleDbDataAdapter oda = new OleDbDataAdapter(cmd);
            oda.Fill(EmployeeTable);
            return EmployeeTable;
        }

        public DataRow GetAssignmentDetailsByLocation(DataTable dt, string location)
        {
            DataRow catalog_display = (from DataRow dr in dt.Rows
                                       where (string)dr["Location  (Cerner)"] == location
                                       select dr).FirstOrDefault();

            return catalog_display;
        }




    }
}
