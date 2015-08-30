
using System;
using System.IO;
using System.ServiceModel.Description;
using System.Net;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Client;


namespace ExportSolution
{
    
    public class Session
    {
        //Assigning initial values
        public Session()
        {
            UserName = "david.karapetyan@praemiumMIC.onmicrosoft.com";
            Password = "David2015";
            SolutionName = "HelloWorld";
            OrganizationUri = new Uri("https://praemiummic.api.crm5.dynamics.com/XRMServices/2011/Organization.svc");
        }

        //Properties
        public string UserName { get; set; }
        public string Password { get; set; }
        public string SolutionName { get; set; }
        public Uri OrganizationUri { get; set; }

        public Session(string name, string password, string solName, Uri orgUri)
        {
            UserName = name;
            Password = password;
            SolutionName = solName;
            OrganizationUri = orgUri;
        }
    }

    public class Program
    {
        //Creates string which contains the date
        public static string DateToString()
        {
            string fileDate = DateTime.Now.Date.ToShortDateString();
            string formatedDate = fileDate.Replace("/", "-");
            string fileTime = DateTime.Now.ToShortTimeString().Replace(" ", "-");
            string formatedTime = fileTime.ToString().Replace(":", "-");
            string finalformat = formatedDate + "_" + formatedTime;
            return finalformat;
        }

        //Gets credentials for CRM access
        public static ClientCredentials GetCredentials(Session userSession)
        {
            ClientCredentials credentials = new ClientCredentials();
            credentials.UserName.UserName = userSession.UserName;
            credentials.UserName.Password = userSession.Password;
            credentials.Windows.ClientCredential = CredentialCache.DefaultNetworkCredentials;
            return credentials;
        }
        private static ClientCredentials GetDeviceCredentials()
        {
            return Microsoft.Crm.Services.Utility.DeviceIdManager.LoadOrRegisterDevice();
        }

        //Start of Main
        static public void Main(string[] args)
        {
            Console.WriteLine("The process has started \nTrying to Export Solution");
            //Create new user and gets credentials to login and capture desired solution
            Session loginUser = new Session();
            ClientCredentials credentials = GetCredentials(loginUser);
                                
            using (OrganizationServiceProxy serviceProxy = new OrganizationServiceProxy(loginUser.OrganizationUri,null, credentials, GetDeviceCredentials()))
            {
                string outputDir = @"C:\temp\";

                //Creates the Export Request
                ExportSolutionRequest exportRequest = new ExportSolutionRequest();
                exportRequest.Managed = true;
                exportRequest.SolutionName = loginUser.SolutionName;

                
                ExportSolutionResponse exportResponse = (ExportSolutionResponse)serviceProxy.Execute(exportRequest);

                //Handles the response
                byte[] exportXml = exportResponse.ExportSolutionFile;
                string filename = loginUser.SolutionName + "_" + DateToString() + ".zip";
                File.WriteAllBytes(outputDir + filename, exportXml);

                Console.WriteLine("Solution Successfully Exported to {0}", outputDir + filename);

            }

        }
        //End of Main

       
    }

}

