using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DiscernDequeService
{


    public class HeaderBt
    {
        public string botAdapterOper { get; set; }
        public string userid { get; set; }
        public string correlationid { get; set; }
        public string sopexecutionId { get; set; }
        public string stepid { get; set; }
        public string appcode { get; set; }
        public string botprocess { get; set; }
        public string skipCredentialCheck { get; set; }
        public string nodeCredCheck { get; set; }
        public string sessionReuse { get; set; }
        public int timeoutseconds { get; set; }
    }

    public class BotInputData
    {
        public string policyNumber { get; set; }
        public int botTimeOut { get; set; }
    }

    public class Rules
    {
        public string maxage { get; set; }
    }

    public class RootObject
    {
        public HeaderBt header { get; set; }
        public BotInputData botInputData { get; set; }
        public Rules rules { get; set; }
    }
    public class Rootobject2
    {
        public Sopexecjson2 sopExecJSON { get; set; }
    }


    public class Rootobject
    {
        public Sopexecjson sopExecJSON { get; set; }
    }

    public class Sopexecjson
    {
        public string sopName { get; set; }
        public Execsopparamsdata execSOPParamsData { get; set; }
        public Requestheaderdata requestHeaderData { get; set; }
        public Requestrulesdata requestRulesData { get; set; }
    }

    public class Sopexecjson2
    {
        public HeaderBt requestHeaderData { get; set; }

        public BotInputData botInputData { get; set; }

        public Rules requestRulesData { get; set; }
    }

    public class Execsopparamsdata
    {
        public string mrnValue { get; set; }
    }

    public class Requestheaderdata
    {
        public string correlationid { get; set; }
        public string userid { get; set; }
        public string userrole { get; set; }
        public string stepid { get; set; }
        public string appcode { get; set; }
        public string botprocess { get; set; }
        public string timeoutseconds { get; set; }
    }

    public class Requestrulesdata
    {
        public string maxage { get; set; }
    }

}
