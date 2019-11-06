using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace VEYMDataParser
{
    class AllUsersDataObjectBETA
    {
        public partial class Identity
        {
            public string signInType { get; set; }
            public string issuer { get; set; }
            public string issuerAssignedId { get; set; }
        }

        public partial class OnPremisesExtensionAttributes
        {
            public object extensionAttribute1 { get; set; }
            public object extensionAttribute2 { get; set; }
            public object extensionAttribute3 { get; set; }
            public object extensionAttribute4 { get; set; }
            public object extensionAttribute5 { get; set; }
            public object extensionAttribute6 { get; set; }
            public object extensionAttribute7 { get; set; }
            public object extensionAttribute8 { get; set; }
            public object extensionAttribute9 { get; set; }
            public object extensionAttribute10 { get; set; }
            public object extensionAttribute11 { get; set; }
            public object extensionAttribute12 { get; set; }
            public object extensionAttribute13 { get; set; }
            public object extensionAttribute14 { get; set; }
            public object extensionAttribute15 { get; set; }
        }

        public partial class PasswordProfile
        {
            public object password { get; set; }
            public bool forceChangePasswordNextSignIn { get; set; }
            public bool forceChangePasswordNextSignInWithMfa { get; set; }
        }

        public partial class Value
        {
            public string id { get; set; }
            public object deletedDateTime { get; set; }
            public bool accountEnabled { get; set; }
            public object ageGroup { get; set; }
            public List<object> businessPhones { get; set; }
            public string city { get; set; }
            public DateTime createdDateTime { get; set; }
            public string creationType { get; set; }
            public object companyName { get; set; }
            public object consentProvidedForMinor { get; set; }
            public string country { get; set; }
            public string department { get; set; }
            public string displayName { get; set; }
            public object employeeId { get; set; }
            public object faxNumber { get; set; }
            public string givenName { get; set; }
            public List<object> imAddresses { get; set; }
            public object isResourceAccount { get; set; }
            public string jobTitle { get; set; }
            public object legalAgeGroupClassification { get; set; }
            public string mail { get; set; }
            public string mailNickname { get; set; }
            public string mobilePhone { get; set; }
            public object onPremisesDistinguishedName { get; set; }
            public string officeLocation { get; set; }
            public object onPremisesDomainName { get; set; }
            public object onPremisesImmutableId { get; set; }
            public object onPremisesLastSyncDateTime { get; set; }
            public object onPremisesSecurityIdentifier { get; set; }
            public object onPremisesSamAccountName { get; set; }
            public object onPremisesSyncEnabled { get; set; }
            public object onPremisesUserPrincipalName { get; set; }
            public List<object> otherMails { get; set; }
            public string passwordPolicies { get; set; }
            public object postalCode { get; set; }
            public object preferredDataLocation { get; set; }
            public string preferredLanguage { get; set; }
            public List<string> proxyAddresses { get; set; }
            public DateTime refreshTokensValidFromDateTime { get; set; }
            public bool? showInAddressList { get; set; }
            public DateTime signInSessionsValidFromDateTime { get; set; }
            public string state { get; set; }
            public object streetAddress { get; set; }
            public string surname { get; set; }
            public string usageLocation { get; set; }
            public string userPrincipalName { get; set; }
            public string externalUserState { get; set; }
            public DateTime? externalUserStateChangeDateTime { get; set; }
            public string userType { get; set; }
            [JsonProperty("extension_4d982a3099ee47359aed5ac368c6d277_Chapter")]
            public string chapter { get; set; }
            [JsonProperty("extension_4d982a3099ee47359aed5ac368c6d277_MemberID_Legacy")]
            public string memberID { get; set; }
            [JsonProperty("extension_4d982a3099ee47359aed5ac368c6d277_Rank")]
            public string rank { get; set; }
            [JsonProperty("extension_4d982a3099ee47359aed5ac368c6d277_League")]
            public string leauge { get; set; }
            public List<object> assignedLicenses { get; set; }
            public List<object> assignedPlans { get; set; }
            public List<object> deviceKeys { get; set; }
            public List<Identity> identities { get; set; }
            public OnPremisesExtensionAttributes onPremisesExtensionAttributes { get; set; }
            public List<object> onPremisesProvisioningErrors { get; set; }
            public PasswordProfile passwordProfile { get; set; }
            public List<object> provisionedPlans { get; set; }
            public string extension_4d982a3099ee47359aed5ac368c6d277_Region { get; set; }
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