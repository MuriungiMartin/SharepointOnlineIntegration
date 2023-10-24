using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Authentication;
using System.Security;
using System.Net.NetworkInformation;
using System.Net;

namespace SharepointOnlineIntegration.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class SharepointController : ControllerBase
    {
        public IConfiguration configuration = new ConfigurationBuilder()
.SetBasePath(Directory.GetCurrentDirectory())
.AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
.Build();

        [HttpGet("ListItems")]
        public  IActionResult GetListItems(string listName)
        {
            Config.DefaultResponse response = new Config.DefaultResponse();

            try
            {
                //string connectionString = configuration.GetConnectionString("DefaultConnection");
                var siteUrl = configuration["Sharepoint:SiteUrl"];
                var userName = configuration["Sharepoint:UserName"];
                var password = configuration["Sharepoint:Password"];
                var domain = configuration["Sharepoint:Domain"];
                var securePassword = new SecureString();
                foreach (char c in password)
                {
                    securePassword.AppendChar(c);
                }
                var onlineCredentials = new  NetworkCredential(userName, securePassword);
                using (var clientContext = new ClientContext(siteUrl))
                {
                    clientContext.Credentials = onlineCredentials;
                    var web = clientContext.Web;
                    var list = web.Lists.GetByTitle(listName);
                    var items = list.GetItems(CamlQuery.CreateAllItemsQuery());
                    clientContext.Load(items);
                    clientContext.ExecuteQuery();
                    var result = new List<object>();
                    foreach (var item in items)
                    {
                        result.Add(new
                        {
                            Id = item.Id,
                            Title = item["Title"]
                        });
                    }
                    return Ok(new Config.DefaultResponse(200, 1, "Success", result));
                }
            }
            catch (Exception ex)
            {
                return StatusCode(StatusCodes.Status500InternalServerError, new Config.DefaultResponse(500, 0, ex.Message));
            }
        }

        public void ConnectToSharePointOnline(string siteCollUrl, string userName, string password)
        {

            //Namespace: It belongs to Microsoft.SharePoint.Client
            ClientContext ctx = new ClientContext(siteCollUrl);

            // Namespace: It belongs to System.Security
            SecureString secureString = new SecureString();
            password.ToList().ForEach(secureString.AppendChar);

            // Namespace: It belongs to Microsoft.SharePoint.Client
            ctx.Credentials = new NetworkCredential(userName, secureString);

            // Namespace: It belongs to Microsoft.SharePoint.Client
            Site mySite = ctx.Site;

            ctx.Load(mySite);
            ctx.ExecuteQuery();

            Console.WriteLine(mySite.Url.ToString());
        }

        //Upload file to SharePoint Online Document Library
        public void UploadFileToSharePointOnline(string siteCollUrl, string userName, string password, string sourceFilePath, string targetLibraryPath)
        {
            // Namespace: It belongs to Microsoft.SharePoint.Client
            ClientContext ctx = new ClientContext(siteCollUrl);
            // Namespace: It belongs to System.Security
            SecureString secureString = new SecureString();
            password.ToList().ForEach(secureString.AppendChar);
            // Namespace: It belongs to Microsoft.SharePoint.Client
            ctx.Credentials = new NetworkCredential(userName, secureString);
            // Namespace: It belongs to Microsoft.SharePoint.Client
            Web web = ctx.Web;
            // Namespace: It belongs to Microsoft.SharePoint.Client
            FileCreationInformation newFile = new FileCreationInformation();
            //Get the file name from source file path
            string fileName = System.IO.Path.GetFileName(sourceFilePath);
            //Assign to content byte[] i.e. documentStream
            newFile.Content = System.IO.File.ReadAllBytes(sourceFilePath);
            //Allow owerwrite of document
            newFile.Overwrite = true;
            //Upload URL
            newFile.Url = siteCollUrl + targetLibraryPath + fileName;
            //Upload document
            Microsoft.SharePoint.Client.File uploadFile = web.GetFolderByServerRelativeUrl(targetLibraryPath).Files.Add(newFile);
            //Update the metadata for a field having name "DocType"
            uploadFile.ListItemAllFields["DocType"] = "Sample Document";
            //Update the metadata for a field having name "Project"
            uploadFile.ListItemAllFields["Project"] = "SharePoint Online Development";
            //Update the metadata for a field having name "DocumentType"
            uploadFile.ListItemAllFields["DocumentType"] = "Word Document";
            //Update the metadata for a field having name "DocumentDescription"
            uploadFile.ListItemAllFields["DocumentDescription"] = "How to Upload a document to SharePoint Online Document Library";
            //Update the metadata for a field having name "DocumentOwner"
            uploadFile.ListItemAllFields["DocumentOwner"] = "TSInfo Technologies";
            //Update the metadata for a field having name "DocumentStatus"
            uploadFile.ListItemAllFields["DocumentStatus"] = "Draft";
            //Update the metadata for a field having name "DocumentID"
            uploadFile.ListItemAllFields["DocumentID"] = "TSInfo-DOC-0001";
        }
    }

}

