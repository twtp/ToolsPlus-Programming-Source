using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace JetDotComInterface
{
    public static class UPS_API
    {
        public struct UPS_Credentials
        {
            public string APILicenseNumber;
            public string APIUserId;
            public string APIPassword;
            public string APIAccountNo;
            public string APIProductionBaseURL;
            public string APITestBaseURL;
            public string APIStreetAddressValidationSuffix;
            public string APIAccessRequestXML;
        }
        public static UPS_Credentials API_Credentials = new UPS_Credentials();

        public static UPS_Credentials GenerateLoginCredentialsXML()
        {
            API_Credentials = new UPS_Credentials();
            List<string> APIData = Connectivity.SQLCalls.sqlQuery("SELECT APIKey,APIName,APIUrl,MerchantID,SecretCode FROM APIKeys WHERE ID=5");
            
            if (APIData.Count > 0)
            {
                API_Credentials.APILicenseNumber = APIData[0].Split('|')[0];
                API_Credentials.APIUserId = APIData[0].Split('|')[1];
                API_Credentials.APIPassword = APIData[0].Split('|')[4];
                API_Credentials.APIProductionBaseURL = APIData[0].Split('|')[2];
                API_Credentials.APITestBaseURL = "https://wwwcie.ups.com:443";
                API_Credentials.APIAccountNo = APIData[0].Split('|')[3];   
                API_Credentials.APIStreetAddressValidationSuffix = "/ups.app/xml/XAV";
            }
            else
            {
                //error email or such...

            }

            string loginCreds = "<?xml version=\"1.0\"?>\r\n";
            loginCreds += "<AccessRequest xml:lang='en-US'>\r\n";
            loginCreds += " <AccessLicenseNumber>" + API_Credentials.APILicenseNumber + "</AccessLicenseNumber>\r\n";
            loginCreds += " <UserId>" + API_Credentials.APIUserId + "</UserId>\r\n";
            loginCreds += " <Password>" + API_Credentials.APIPassword + "</Password>\r\n";
            loginCreds += "</AccessRequest>\r\n";

            API_Credentials.APIAccessRequestXML = loginCreds;

            return API_Credentials;
        }
        public static string GenerateAddressValidationXML(List<string> Address, string Zipcode,string CountryCode)
        {
            return GenerateAddressValidationXML(Address, Zipcode, CountryCode, "", "", "", "");
        }
        public static string GenerateAddressValidationXML(List<string> Address, string Zipcode, string CountryCode, string ConsigneeName, string BuildingName, string PoliticalDivisionName, string PoliticalDivisionCode)
        {
            string avRequest = "<?xml version=\"1.0\"?>\r\n";
            avRequest += "<AddressValidationRequest xml:lang='en-US'>\r\n";
            avRequest += " <Request>\r\n";
            avRequest += "  <TransactionReference>\r\n";
            avRequest += "   <CustomerContext/><XpciVersion>1.0001</XpciVersion>\r\n";
            avRequest += "  </TransactionReference>\r\n";
            avRequest += "  <RequestAction>XAV</RequestAction>\r\n";
            avRequest += "  <RequestOption>3</RequestOption>\r\n";
            avRequest += " </Request>\r\n";
            avRequest += " <MaximumListSize>3</MaximumListSize>\r\n";
            avRequest += " <AddressKeyFormat>\r\n";

            if (ConsigneeName != "")
            {
                avRequest += "  <ConsigneeName>" + ConsigneeName + "</ConsigneeName>\r\n";
            }
            if (BuildingName != "")
            {
                avRequest += "  <BuildingName>" + BuildingName + "</BuildingName>\r\n";
            }
            foreach (string aLine in Address)
            {
                avRequest += "  <AddressLine>" + aLine + "</AddressLine>\r\n";
            }
            if (PoliticalDivisionName != "")
            {
                avRequest += "  <PoliticalDivision2>" + PoliticalDivisionName + "</PoliticalDivision2>\r\n";
            }
            if (PoliticalDivisionCode != "")
            {
                avRequest += "  <PoliticalDivision1>" + PoliticalDivisionCode + "</PoliticalDivision1>\r\n";
            }
            avRequest += "  <PostcodePrimaryLow>" + Zipcode + "</PostcodePrimaryLow>\r\n";
            if (CountryCode != "")
            {
                avRequest += "  <CountryCode>" + CountryCode + "</CountryCode>\r\n";
            }
            avRequest += " </AddressKeyFormat>\r\n";
            avRequest += "</AddressValidationRequest>\r\n";

            return avRequest;

        }

    }
}
