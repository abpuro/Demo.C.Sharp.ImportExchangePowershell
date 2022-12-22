using Demo.RunExchangeOnlinePowershell.Models;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Security;

namespace Demo.RunExchangeOnlinePowershell.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;
        private readonly IWebHostEnvironment _hostingEnvironment;


        public HomeController(ILogger<HomeController> logger, IWebHostEnvironment host )
        {
            _logger = logger;
            _hostingEnvironment = host; 
        }

        public IActionResult Index()
        {

            InitialSessionState iss = InitialSessionState.CreateDefault();
            iss.ExecutionPolicy = Microsoft.PowerShell.ExecutionPolicy.Bypass;
            iss.ImportPSModule(new[] { "ExchangeOnlineManagement" });

            Runspace rs = RunspaceFactory.CreateRunspace(iss);
            try
            {

                rs.Open();
                PowerShell ps = PowerShell.Create();
                ps.Runspace = rs;

                //Collection<PSObject> results = ps.AddCommand("Connect-ExchangeOnline").AddParameter("UserPrincipalName", "admin@m365devsub.onmicrosoft.com").Invoke();
                Collection<PSObject> results = ps.AddCommand("Get-Module").Invoke();



                if (ps.HadErrors)
                {
                    foreach (var error in ps.Streams.Error)
                    {

                    }
                }
                foreach (var result in results)
                {

                }

            }
            catch (Exception ex)
            {
                //log.LogInformation(ex.Message);
            }
            finally
            {
                rs.Close();
            }
















            ////InitialSessionState iss = InitialSessionState.CreateDefault();
            ////iss.ImportPSModule(new[] { "ExchangeOnlineManagement" });

            ////Runspace rs = RunspaceFactory.CreateRunspace(iss);
            ////rs.Open();
            ////PowerShell ps = PowerShell.Create();
            ////ps.Runspace = rs;

            ////ps.AddCommand("Connect-ExchangeOnline").AddParameter("UserPrincipalName", "admin@m365devsub.onmicrosoft.com").Invoke();
            ////rs.Close();

            //try
            //{
            //    Collection<PSObject> userList = null;
            //    // Create Initial Session State for runspace.
            //    InitialSessionState initialSession = InitialSessionState.CreateDefault2();

            //    initialSession.ImportPSModule(new[] { "MSOnline" });
            //    // Create credential object.
            //    SecureString ss = new SecureString();
            //    var credential = new PSCredential("admin@m365devsub.onmicrosoft.com", ConvertToSecureString("Arvind@1234"));
            //    // Create command to connect office 365.
            //    Command connectCommand = new Command("Connect-MsolService");
            //    connectCommand.Parameters.Add((new CommandParameter("Credential", credential)));
            //    // Create command to get office 365 users.
            //    Command getUserCommand = new Command("Get-MsolUser");
            //    using (Runspace psRunSpace = RunspaceFactory.CreateRunspace(initialSession))
            //    {
            //        // Open runspace.
            //        psRunSpace.Open();

            //        //Iterate through each command and executes it.
            //        foreach (var com in new Command[] { connectCommand, getUserCommand })
            //        {
            //            var pipe = psRunSpace.CreatePipeline();
            //            pipe.Commands.Add(com);
            //            // Execute command and generate results and errors (if any).
            //            Collection<PSObject> results = pipe.Invoke();
            //            var error = pipe.Error.ReadToEnd();
            //            if (error.Count > 0 && com == connectCommand)
            //            {

            //                return Content(error[0].ToString());
            //            }
            //            if (error.Count > 0 && com == getUserCommand)
            //            {

            //                return Content(error[0].ToString());
            //            }
            //            else
            //            {
            //                userList = results;
            //            }
            //        }
            //        // Close the runspace.
            //        psRunSpace.Close();
            //    }

            //}
            //catch (Exception)
            //{
            //    throw;
            //}
            string moduleName = "None Loaded";
            if (iss.Modules.Count > 0)
            {
                moduleName = iss.Modules[0].Name;
            }
            return Content(moduleName);
        }

        public IActionResult Privacy()
        {
            return View();
        }

        private SecureString ConvertToSecureString(string password)
        {
            if (password == null)
                throw new ArgumentNullException("password");

            var securePassword = new SecureString();

            foreach (char c in password)
                securePassword.AppendChar(c);

            securePassword.MakeReadOnly();
            return securePassword;
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}