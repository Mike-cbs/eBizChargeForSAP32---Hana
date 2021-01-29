using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.IO;
using System.Security.Cryptography;
using System.Net;

using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace EBizEMV
{
    //*****************************************

    public class DeviceList
    {
        // public DeviceDetails();

        public string type { get; set; }
        public string limit { get; set; }
        public string offset { get; set; }
        public Device[] data { get; set; }
        public string total { get; set; }

    }


    public class DeviceDetails
    {
        // public DeviceDetails();

        public string make { get; set; }
        public string model { get; set; }
        public string revision { get; set; }
        public string serial { get; set; }
    }
    public class TerminalConfig
    {
        //  public TerminalConfig();

        public bool enable_emv { get; set; }
        public bool enable_debit_msr { get; set; }
        public bool tip_adjust { get; set; }
        public bool enable_contactless { get; set; }
    }

    public class Terminalcapabilities
    {
        //  public TerminalConfig();

        public bool emv { get; set; }
        public bool swipe { get; set; }
        public bool contactless { get; set; }
        public bool debit { get; set; }
        public bool signature_capture { get; set; }

    }

    public class DeviceSettings
    {

        public int timeout { get; set; }
        public bool enable_standalone { get; set; }
        public bool share_device { get; set; }
    }
    public class Device
    {

        public DeviceSettings settings { get; set; }
        public DateTime expiration { get; set; }
        public string pairing_code { get; set; }
        public string name { get; set; }
        public string status { get; set; }
        public string terminal_type { get; set; }
        public string type { get; set; }
        public string key { get; set; }
        public string apikeyid { get; set; }
        public string apikey_masked { get; set; }
        public bool update_available { get; set; }


        public TerminalConfig terminal_config { get; set; }
        public DeviceDetails terminal_info { get; set; }
        public Terminalcapabilities capabilities { get; set; }



    }
    public class AmountDetail
    {
        //  public AmountDetail();

        public string tax { get; set; }
        public string nontaxable { get; set; }
        public string cash { get; set; }
        public string tip { get; set; }
        public string discount { get; set; }
        public string shipping { get; set; }
        public string duty { get; set; }
        public bool enable_partialauth { get; set; }
    }
    public class LineItem
    {
        //    public LineItem();

        public string sku { get; set; }
        public string name { get; set; }
        public string description { get; set; }
        public string cost { get; set; }
        public string qty { get; set; }
        public string taxamount { get; set; }
        public string taxable { get; set; }
        public string taxrate { get; set; }
        public string discountamount { get; set; }
        public string discountrate { get; set; }
        public string um { get; set; }

    }

    public class Address
    {
        //   public Address();

        public string firstname { get; set; }
        public string lastname { get; set; }
        public string company { get; set; }
        public string street { get; set; }
        public string street2 { get; set; }
        public string city { get; set; }
        public string state { get; set; }
        public string postalcode { get; set; }
        public string country { get; set; }
        public string phone { get; set; }
        public string fax { get; set; }
    }
    public class TransactionRequest
    {

        public bool prompt_signature { get; set; }
        public bool prompt_cashback { get; set; }
        public bool prompt_tip { get; set; }
        public bool manual_key { get; set; }
        public LineItem[] lineitems { get; set; }
        public string geolocation { get; set; }
        public string clientip { get; set; }
        public string customerid { get; set; }
        public string email { get; set; }
        public Address shipping_address { get; set; }
        public string custom_flow { get; set; }
        public Address billing_address { get; set; }
        public string table { get; set; }
        public string comments { get; set; }
        public string description { get; set; }
        public string orderid { get; set; }
        public string invoice { get; set; }
        public string ponum { get; set; }
        public AmountDetail amount_detail { get; set; }
        public string amount { get; set; }
        public string command { get; set; }
        public string devicekey { get; set; }
        public string clerk { get; set; }
        public bool save_card { get; set; }
    }


    public class TransactionResult1
    {

        public string key { get; set; }
        public string refnum { get; set; }
        public string is_duplicate { get; set; }
        public string result_code { get; set; }
        public string result { get; set; }
        public string auth_amount { get; set; }
        public string authcode { get; set; }
        public string error { get; set; }
        public string error_code { get; set; }
        public string proc_refnum { get; set; }
        public AvsResult avs { get; set; }
        public CvcResult cvc { get; set; }
        public CreditCard creditcard { get; set; }
    }

    public class PaymentResponse
    {
        public string key { get; set; }
        public string expiration { get; set; }
        public string status { get; set; }
        public string type { get; set; }


    }

    public class DeviceTransactionResult
    {
        public string type { get; set; }

        public string status { get; set; }
        public string key { get; set; }


        public DeviceTransaction transaction { get; set; }
    }

    public class DeviceTransaction
    {
        public string type { get; set; }
        public string key { get; set; }

        public string refnum { get; set; }

        public string is_duplicate { get; set; }
        public string result_code { get; set; }
        public string result { get; set; }
        public string authcode { get; set; }


        public string error { get; set; }
        public string error_code { get; set; }
        public string iccdata { get; set; }
        public string proc_refnum { get; set; }
        public string auth_amount { get; set; }
        public string token { get; set; }

        public AvsResult avs { get; set; }
        public CvcResult cvc { get; set; }
        public CreditCard creditcard { get; set; }
    }

    public class AvsResult
    {
        //   public AvsResult();

        public string result_code { get; set; }
        public string result { get; set; }
    }
    public class CvcResult
    {
        //    public CvcResult();

        public string result_code { get; set; }
        public string result { get; set; }
    }

    public class CreditCard
    {
        //   public CreditCard();

        public string type { get; set; }
        public string number { get; set; }
        public string cardholder { get; set; }
        public string entrymode { get; set; }
        public string category_code { get; set; }
        public string avs_street { get; set; }
        public string avs_postalcode { get; set; }
        public string verification { get; set; }
        public string signature { get; set; }
        public string token { get; set; }
    }
    //-----------------------------------------------------

    //***************************************
    public static class EMVHelper
    {

        //-----------------------------------------------------------------------------------------------------------------------------------------------------------
        public static string GetDevices(string Url, string DeviceKey, string ApiKey, string ApiPin)
        {
            string prmAuthorization = BuildAuthHeader(ApiKey, ApiPin);
            string strrequest = "/paymentengine/devices/" + DeviceKey;

            string response = SubmitRequest(Url, strrequest, "GET", prmAuthorization);
            return response;


        }
        //-----------------------------------------------------------------------------------------------------------------------------------------------------------

        //-----------------------------------------------------------------------------------------------------------------------------------------------------------
        public static string GetDeviceList(string Url, string ApiKey, string ApiPin)
        {
            string prmAuthorization = BuildAuthHeader(ApiKey, ApiPin);
            string strrequest = "/paymentengine/devices/";

            string response = SubmitRequest(Url, strrequest, "GET", prmAuthorization);
            return response;


        }
        //-----------------------------------------------------------------------------------------------------------------------------------------------------------

        public static string Payrequest(string Url, string DeviceKey, string ApiKey, string ApiPin, TransactionRequest payreq, string tranCommand)
        {
            string prmAuthorization = BuildAuthHeader(ApiKey, ApiPin);
            string response = "";
            try
            {
                payreq.devicekey = DeviceKey; // "sa_R6g6wrjLeqdyKEcxpdL7qPsYfkkdc";
                payreq.command = tranCommand;

                JObject o = (JObject)JToken.FromObject(payreq);

                string strJsonReq1 = o.ToString();

                //  string strJsonReq = "{\"devicekey\": \"sa_R6g6wrjLeqdyKEcxpdL7qPsYfkkdc\",\"command\": \"cc:sale\",\"amount\":\"8.88\"}";

                Url = Url + "/paymentengine/payrequests";

                 response = SubmitPOSTRequest(Url, strJsonReq1, "POST", prmAuthorization);
            }
            catch
            {

            }
            return response;


        }
        //-----------------------------------------------------------------------------------------------------------------------------------------------------------
        //-----------------------------------------------------------------------------------------------------------------------------------------------------------
        public static string GetPayrequestResult(string Url, string DeviceKey, string ApiKey, string ApiPin, string payrequestResponse)
        {
            string prmAuthorization = BuildAuthHeader(ApiKey, ApiPin);

            PaymentResponse account = JsonConvert.DeserializeObject<PaymentResponse>(payrequestResponse);

            string strrequest = "/paymentengine/payrequests/" + account.key;

            string response = SubmitRequest(Url, strrequest, "GET", prmAuthorization);
            return response;


        }
        //-----------------------------------------------------------------------------------------------------------------------------------------------------------

        public static String BuildAuthHeader(String key, String pin)
        {

            // create seed for hash
            String Seed = Guid.NewGuid().ToString();

            // create hash of apikey, seed and pin
            String Hash = GenerateHash(key + Seed + pin);

            // assemble auth string,  api key is the username,  hash is the password
            String authInfo = key + ":s2/" + Seed + "/" + Hash;
            authInfo = "Basic " + Convert.ToBase64String(Encoding.Default.GetBytes(authInfo));

            return authInfo;
        }
        //-----------------------------------------------------------------------------------------------------------------

        public static String GenerateHash(string input)
        {
            // Create a new instance of the SHA256CryptoServiceProvider object.
            SHA256 shaHasher = SHA256.Create();

            // Convert the input string to a byte array and compute the hash.
            byte[] data = shaHasher.ComputeHash(Encoding.Default.GetBytes(input));

            // Create a new Stringbuilder to collect the bytes
            // and create a string.
            StringBuilder sBuilder = new StringBuilder();

            // Loop through each byte of the hashed data 
            // and format each one as a hexadecimal string.
            for (int i = 0; i < data.Length; i++)
            {
                sBuilder.Append(data[i].ToString("x2"));
            }

            // Return the hexadecimal string.
            return sBuilder.ToString();
        }

        //------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        public static string SubmitRequest(string prmUrl, string prmRequest, string prmRequestMethod, string prmAuthorization)
        {


            HttpWebRequest req = (HttpWebRequest)WebRequest.Create(prmUrl + prmRequest);
            req.ContentType = "application/json;charset=utf-8";

            //   req.Accept = "text/html";

            req.Headers["Authorization"] = prmAuthorization;

            req.Method = prmRequestMethod;


            //     Console.WriteLine("prmUrl + prmRequest: " + prmUrl + prmRequest);

            string strresult = "";

            using (HttpWebResponse response = req.GetResponse() as HttpWebResponse)
            {
                if (response.StatusCode != HttpStatusCode.OK)
                    throw new Exception(String.Format(
                    "Server error (HTTP {0}: {1}).",
                    response.StatusCode,
                    response.StatusDescription));
                Stream stream1 = response.GetResponseStream();
                StreamReader sr = new StreamReader(stream1);
                strresult = sr.ReadToEnd();

            }

            return strresult;
        }
        //------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        //------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        public static string SubmitPOSTRequest(string requestUrl, string prmRequest, string prmRequestMethod, string prmAuthorization)
        {
            string strsb = "";

            try
            {
                string webAddr = requestUrl;

                var httpWebRequest = (HttpWebRequest)WebRequest.Create(webAddr);
                httpWebRequest.ContentType = "application/json; charset=utf-8";
                httpWebRequest.Method = "POST";
                httpWebRequest.Headers["Authorization"] = prmAuthorization;

                using (var streamWriter = new StreamWriter(httpWebRequest.GetRequestStream()))
                {
                    string json = prmRequest;

                    streamWriter.Write(json);
                    streamWriter.Flush();
                }
                var httpResponse = (HttpWebResponse)httpWebRequest.GetResponse();
                using (var streamReader = new StreamReader(httpResponse.GetResponseStream()))
                {
                    var responseText = streamReader.ReadToEnd();
                    Console.WriteLine(responseText);

                    PaymentResponse account = JsonConvert.DeserializeObject<PaymentResponse>(responseText);

                //    Console.WriteLine("key: " + account.key);
                    strsb = responseText;
                }
            }
            catch (WebException ex)
            {
               // Console.WriteLine("Errors: " + ex.Message);
                strsb = "Errors: " + ex.Message;
            }
            return strsb;


            //----------------------------------------------------
        }
    }
}
