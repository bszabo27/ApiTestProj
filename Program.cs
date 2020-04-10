using Google.Apis.Auth.OAuth2;
using Google.Apis.Drive.v2;
using Google.Apis.Services;
using System;
using System.Security.Cryptography.X509Certificates;
using Google.Apis.Util.Store;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;


namespace ApiTestProj
{
    class Program
    {
        static void Main(string[] args)
        {
            //Google Auth Methods for Existing GoogleAPI Projects
            //Generate at console.developers.google.com-- > credentials

            //API Key: 
            //A generated key, registered in the GoogleAPI project, anyone who brings the key (http header or queryparam) have access to the hoster's public data, or "link-shared" data.
            //As of 2020.04.08 it's not working for writes anymore, since all nd every write requires Authentication, and API Key is not proper Authentication.

            //OAuth:
            //Auditible permissions system, with pre-configured permissions, reviewed by google. If the permissions dont include Restricted or Sensitive data, the review process is less painful.
            //Otherwise it takes weeks and possibly a bag of money. 
            //This is when an app tries to access a resouce of the owner (liek a spreadsheet) and a popup asks for agreement from the user on the private data authorizations (these ae base on the scopes.)
            //More at: https://support.google.com/cloud/answer/9110914?hl=en

            //Service Accounts
            //It's basically an non-user dependent google account, so it is supposed to be used by an App exlusively as a resource, so no user permissions are needed, as the data does not belong to a user, thus: "Service" Account.
            //Keep in mind, any fiels created trough a Service Account belongs to the Service Account, not to you, meaning you dond have read permissions as "kovacs.bela@gmail.com", unless its not provided trough the Service Account.
            //A workaround is to create a google doc, and share it with a Service Account.
            //Since Service Account generated / permitted files are not user dependent, access any rights can be acquired only by defining the Scopes. In this example we use Drive.all

            //Additional useful links:
            // https://www.gmass.co/blog/google-sheets-api-v4-bullshit/
            // https://developers.google.com/identity/protocols/oauth2/service-account  
            // https://developers.google.com/sheets/api/samples/writing

            //Generated and donwloaded Service Account key
            var serviceAccount_keyFile = @"c:\apitest-273518-982e93166b41.p12";    
            var serviceAccount_EmailAddress = "postmanapitestacc@apitest-273518.iam.gserviceaccount.com"; 
            
            //Define scopes for later use. We can use all Drive scopes beacuse its a Service Acc.
            //Otherwise we would need to go trugh a 3-4 week long review process by google to get special permissions for those operations.
            string[] scopes = new string[] { DriveService.Scope.Drive };
            
            //Reading API key. We should use the newer JSON key, but I accidentaly downloaded p12 and I am a lazy potáta.
            var certi = new X509Certificate2(serviceAccount_keyFile, "notasecret", X509KeyStorageFlags.Exportable);
            var credential = new ServiceAccountCredential(new ServiceAccountCredential.Initializer(serviceAccount_EmailAddress)
            {
                Scopes = scopes
            }.FromCertificate(certi));

            //Dont need this ATM, we just want a spreadsheet, no need to overexpose.
            //DriveService service = new DriveService(new BaseClientService.Initializer() { 
            //    HttpClientInitializer = credential,
            //    ApplicationName = "PostmanApitestAccTesterC"
            //});

            var service = new SheetsService(new BaseClientService.Initializer() { 
               HttpClientInitializer = credential,
                ApplicationName = "PostmanApitestAccTesterC"
            });

            //Spreadsheet ID from browser, and other identifiers for the data destination.
            String spreadsheet = "13gJJGDBkSHMXdVo0CaKhreKi7R6r5FZDW9tlGgWwbKQ";
            String sheetandcell = "Munkalap1!C4";  

            //DataDims
            ValueRange valueRange = new ValueRange();
            valueRange.MajorDimension = "COLUMNS";

            //And the data
            var oblist = new List<object>() { "Jig Rulz" };
            valueRange.Values = new List<IList<object>> { oblist };

            //Filling params and firing request
            SpreadsheetsResource.ValuesResource.UpdateRequest update = service.Spreadsheets.Values.Update(valueRange, spreadsheet, sheetandcell);
            update.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
            UpdateValuesResponse result = update.Execute();         


            Console.ReadLine();

        }
    }
}
