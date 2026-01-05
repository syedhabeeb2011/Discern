/*
 * Created by SharpDevelop.
 * User: v.chakradhar.ark
 * Date: 11/5/2019
 * Time: 5:40 AM
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;

namespace Atos.SyntBots.Common.Entity
{
	/// <summary>
	/// Description of Charge.
	/// </summary>
	public class Charge
    {
        public string OrderDescription;
        public string FIN;
        public string ChargeDescription;
        public string ChargeDate;
        public string CPT4;
        public string Price;
        public string ActivityType;
        public string ChargeType;
        public string Status;
    }
    public class Order
    {
        public string FIN;
        public string OrderDescription;
        public string OrderDate;
        public string CPT4;
        public string DischargeDate;
        public string Comments;
        public string OrderStatus;
        public string ResultDate;
        public string Organization;
        
    }
}
