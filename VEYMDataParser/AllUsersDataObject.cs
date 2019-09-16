using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VEYMDataParser
{
    class AllUsersDataObject
    {
        //Values are HT
        public partial class Value
        {
            public List<object> businessPhones { get; set; }
            public string displayName { get; set; }
            public string givenName { get; set; }
            public string jobTitle { get; set; }
            public string mail { get; set; }
            public string mobilePhone { get; set; }
            public string officeLocation { get; set; }
            public string preferredLanguage { get; set; }
            public string surname { get; set; }
            public string userPrincipalName { get; set; }
            public string id { get; set; }
        }

        public partial class RootObject
        {
            [JsonProperty("@odata.context")]
            public string context { get; set; }
            [JsonProperty("@odata.nextLink")]
            public string nextLink { get; set; }
            public List<Value> value { get; set; }
        }
    }
}

