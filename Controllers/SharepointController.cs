using System;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using SharepointOnlineIntegration;
using System.Text.Json;
using Newtonsoft.Json;
using System.Text.Json.Nodes;
using Microsoft.Extensions.Configuration;
using System.Net.Http.Headers;
using SharepointOnlineIntegration.Models;
using System.Net;
using Newtonsoft.Json.Linq;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint;

public class AccessTokenResponse
{
    [JsonProperty("token_type")]
    public string TokenType { get; set; }

    [JsonProperty("expires_in")]
    public int ExpiresIn { get; set; }

    [JsonProperty("not_before")]
    public long NotBefore { get; set; }

    [JsonProperty("expires_on")]
    public long ExpiresOn { get; set; }

    [JsonProperty("resource")]
    public string Resource { get; set; }

    [JsonProperty("access_token")]
    public string AccessToken { get; set; }
}

[ApiController]
[Route("[controller]")]
public class SharepointController : ControllerBase
{
    public IConfiguration configuration = new ConfigurationBuilder()
.SetBasePath(Directory.GetCurrentDirectory())
.AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
.Build();

    //get variable from config appsettings


    [HttpPost("token")]
    public async Task<IActionResult> GetAccessToken()
    {
        Config.DefaultResponse response = new Config.DefaultResponse();
       

        try
        {
            
            var grantType = "client_credentials";
            var clientId = "";
            var clientSecret = "";
            var resource = "";
            var tenantId = "";
           var url = $"https://accounts.accesscontrol.windows.net/{tenantId}/tokens/OAuth/2";

            using (var httpClient = new HttpClient())
            {
                var requestContent = $"grant_type={grantType}&client_id={clientId}&client_secret={clientSecret}&resource={resource}";

                using (var httpContent = new StringContent(requestContent))
                {
                    httpContent.Headers.Clear();
                    httpContent.Headers.Add("Content-Type", "application/x-www-form-urlencoded");

                    using (var httpResponse = await httpClient.PostAsync(url, httpContent))
                    {
                        var responseText = await httpResponse.Content.ReadAsStringAsync();
                        //get as Json
                        string decodedString = System.Text.RegularExpressions.Regex.Unescape(responseText);
                        AccessTokenResponse responseObject = JsonConvert.DeserializeObject<AccessTokenResponse>(decodedString);
                        



                        if (httpResponse.IsSuccessStatusCode)
                        {
                           


                            return Ok(responseObject?.AccessToken);
                        }
                        else
                        {
                            response = new Config.DefaultResponse(500,"Failed to fetch token", responseObject);

                            return BadRequest(response);
                        }
                    }
                }
            }
        }
        catch (Exception ex)
        {
            var exceptionResponse = new AccessTokenResponse
            {
               
            };

            return StatusCode(500, exceptionResponse);
        }
    }
    [HttpGet("list")]
    public async Task<IActionResult> GetList(string list)
    {
        Config.DefaultResponse response = new Config.DefaultResponse();
        try
        {
            string siteUrl = "";

            var accessTokenResult = await GetAccessToken();

            if (accessTokenResult is ObjectResult accessTokenObjectResult && accessTokenObjectResult.Value is string accessToken)
            {
                var url = $"{siteUrl}/_api/web/lists/getbytitle('{list}')/items";

                using (var httpClient = new HttpClient())
                {
                    httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
                    httpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

                    using (var httpResponse = await httpClient.GetAsync(url))
                    {
                        var responseText = await httpResponse.Content.ReadAsStringAsync();

                        if (httpResponse.IsSuccessStatusCode)
                        {
                            var responseJson = JsonArray.Parse(responseText);
                            return Ok(responseJson);
                        }
                        else
                        {
                            response = new Config.DefaultResponse(500,"Failed to fetch list", responseText);
                            return BadRequest(response);
                        }
                    }
                }
            }
            else
            {
                return BadRequest("Failed to obtain access token");
            }
        }
        catch
        {
            var exceptionResponse = new Config.DefaultResponse(500,"Failed to fetch list", "Uwongo");
            return StatusCode(500, exceptionResponse);
        }
    }

    [HttpPost("file")]
    public async Task<IActionResult> UploadFile([FromBody] FileUpload Upload)
    {
        Config.DefaultResponse response = new Config.DefaultResponse();
        try
        {
            string siteUrl = "";

            var accessTokenResult = await GetAccessToken();

            if (accessTokenResult is ObjectResult accessTokenObjectResult && accessTokenObjectResult.Value is string accessToken)
            {
                string libraryName = "Shared Documents";
                string folderPath = Upload.FolderName;

                string[] folders = folderPath.Trim('/').Split('/');
                var firstFolderPathSet = folders[0] + "/" + folders[1];
                var firstFolderUrl = $"{siteUrl}/_api/web/getfolderbyserverrelativeurl('{libraryName}/{firstFolderPathSet}')/Exists";
                var secondFolderPathSet = folders[0] + "/" + folders[1] + "/" + folders[2];
                var secondFolderUrl = $"{siteUrl}/_api/web/getfolderbyserverrelativeurl('{libraryName}/{secondFolderPathSet}')/Exists";
                var thirdFolderPathSet = folders[0] + "/" + folders[1] + "/" + folders[2] + "/" + folders[3];
                var thirdFolderUrl = $"{siteUrl}/_api/web/getfolderbyserverrelativeurl('{libraryName}/{thirdFolderPathSet}')/Exists";
                var fourthFolderPathSet = folders[0] + "/" + folders[1] + "/" + folders[2] + "/" + folders[3] + "/" + folders[4];
                var fourthFolderUrl = $"{siteUrl}/_api/web/getfolderbyserverrelativeurl('{libraryName}/{fourthFolderPathSet}')/Exists";

                using (var httpClient = new HttpClient())
                {
                    httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
                    httpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                    var checkResponse = await httpClient.GetAsync(firstFolderUrl);
                    if (checkResponse.IsSuccessStatusCode)
                    {
                        var responseText = await checkResponse.Content.ReadAsStringAsync();
                        JObject jsonResponse = JObject.Parse(responseText);
                        bool extractedValue = jsonResponse["value"].Value<bool>();
                        if (!extractedValue)
                        {
                            var createFolderUrl = $"{siteUrl}/_api/web/folders/add('{libraryName}/{firstFolderPathSet}')";
                            using (var createResponse = await httpClient.PostAsync(createFolderUrl, null))
                            {
                                if (!createResponse.IsSuccessStatusCode)
                                {
                                    response = new Config.DefaultResponse(500,"Failed to create folder", "Failed to create folder");
                                    return BadRequest(response);
                                }
                            }
                        }
                    }
                    else
                    {
                        response = new Config.DefaultResponse(500,"Failed", "Failed to check if folder exists");
                        return BadRequest(response);
                    }
                }
                //second folder
                using (var httpClient = new HttpClient())
                {
                    httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
                    httpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                    var checkResponse = await httpClient.GetAsync(secondFolderUrl);
                    if (checkResponse.IsSuccessStatusCode)
                    {
                        var responseText = await checkResponse.Content.ReadAsStringAsync();
                        JObject jsonResponse = JObject.Parse(responseText);
                        bool extractedValue = jsonResponse["value"].Value<bool>();
                        if (!extractedValue)
                        {
                            var createFolderUrl = $"{siteUrl}/_api/web/folders/add('{libraryName}/{secondFolderPathSet}')";
                            using (var createResponse = await httpClient.PostAsync(createFolderUrl, null))
                            {
                                if (!createResponse.IsSuccessStatusCode)
                                {
                                    response = new Config.DefaultResponse(500,"Failed to create folder", "Failed to create folder");
                                    return BadRequest(response);
                                }
                            }
                        }
                    }
                    else
                    {
                        response = new Config.DefaultResponse(500,"Failed", "Failed to check if folder exists");
                        return BadRequest(response);
                    }
                }
                //third folder
                using (var httpClient = new HttpClient())
                {
                    httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
                    httpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                    var checkResponse = await httpClient.GetAsync(thirdFolderUrl);
                    if (checkResponse.IsSuccessStatusCode)
                    {
                        var responseText = await checkResponse.Content.ReadAsStringAsync();
                        JObject jsonResponse = JObject.Parse(responseText);
                        bool extractedValue = jsonResponse["value"].Value<bool>();
                        if (!extractedValue)
                        {
                            var createFolderUrl = $"{siteUrl}/_api/web/folders/add('{libraryName}/{thirdFolderPathSet}')";
                            using (var createResponse = await httpClient.PostAsync(createFolderUrl, null))
                            {
                                if (!createResponse.IsSuccessStatusCode)
                                {
                                    response = new Config.DefaultResponse(500,"Failed to create folder", "Failed to create folder");
                                    return BadRequest(response);
                                }
                            }
                        }
                    }
                    else
                    {
                        response = new Config.DefaultResponse(500,"Failed", "Failed to check if folder exists");
                        return BadRequest(response);
                    }
                }
                //fourth folder
                using (var httpClient = new HttpClient())
                {
                    httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
                    httpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                    var checkResponse = await httpClient.GetAsync(fourthFolderUrl);
                    if (checkResponse.IsSuccessStatusCode)
                    {
                        var responseText = await checkResponse.Content.ReadAsStringAsync();
                        JObject jsonResponse = JObject.Parse(responseText);
                        bool extractedValue = jsonResponse["value"].Value<bool>();
                        if (!extractedValue)
                        {
                            var createFolderUrl = $"{siteUrl}/_api/web/folders/add('{libraryName}/{fourthFolderPathSet}')";
                            using (var createResponse = await httpClient.PostAsync(createFolderUrl, null))
                            {
                                if (!createResponse.IsSuccessStatusCode)
                                {
                                    response = new Config.DefaultResponse(500,"Failed to create folder", "Failed to create folder");
                                    return BadRequest(response);
                                }
                            }
                        }
                    }
                    else
                    {
                        response = new Config.DefaultResponse(500,"Failed", "Failed to check if folder exists");
                        return BadRequest(response);
                    }
                }
                    string fileName = Upload.FileName;
                //var uploadUrl = $"{currentFolderUrl}/files/add(url='{fileName}',overwrite=true)";
                var uploadUrl = $"{siteUrl}/_api/web/getfolderbyserverrelativeurl('{libraryName}{folderPath}')/files/add(overwrite=true,url='{fileName}')";

                //Write filecontent to local file


                using (var httpClient = new HttpClient())
                {
                    httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
                    httpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

                    byte[] fileBytes = Convert.FromBase64String(Upload.FileContent);

                    ByteArrayContent fileContent = new ByteArrayContent(fileBytes);
                    fileContent.Headers.ContentType = new MediaTypeHeaderValue("application/octet-stream");

                    using (var httpResponse = await httpClient.PostAsync(uploadUrl, fileContent))
                    {
                        var responseText = await httpResponse.Content.ReadAsStringAsync();

                        if (httpResponse.IsSuccessStatusCode)
                        {

                            var fileUrl= siteUrl + "/" + libraryName +folderPath + "/" + fileName;
                            var myData = new
                            {
                                Name = fileName,
                                fileUrl = fileUrl
                            };
                            response = new Config.DefaultResponse(200, "Success", myData);
                            return Ok(response);
                        }
                        else
                        {
                            response = new Config.DefaultResponse(500,"Failed to upload file", responseText);
                            return BadRequest(response);
                        }
                    }
                }
            }
            else
            {
                return BadRequest("Failed to obtain access token");
            }
        }
        catch (Exception ex)
        {
           var exceptionResponse = new Config.DefaultResponse(500,"Failed to upload file", ex.Message);
            return StatusCode(500, exceptionResponse);
        }
    }

    //Downdload file
    [HttpPost("download")]
    public async Task<IActionResult> DownloadFile([FromBody] FileUpload Upload)
    {
        Config.DefaultResponse response = new Config.DefaultResponse();
        try
        {
            string siteUrl = "";
            var RelativeFilePath = "/sites/NavisionBusinessCentral/Shared Documents" + Upload.FolderName + "/" + Upload.FileName;
            var apiUrl = $"{siteUrl}/_api/web/GetFileByServerRelativePath(decodedurl='{RelativeFilePath}')/$value";

            var accessTokenResult = await GetAccessToken();

            if (accessTokenResult is ObjectResult accessTokenObjectResult && accessTokenObjectResult.Value is string accessToken)
            {
                using (var httpClient = new HttpClient())
                {
                    httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
                    httpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                    var downloadResponse = await httpClient.GetAsync(apiUrl);
                    var responseText = downloadResponse.Content.ReadAsStringAsync();

                    if (downloadResponse.IsSuccessStatusCode)
                    {
                        var filecontent = await downloadResponse.Content.ReadAsByteArrayAsync();
                        return Ok(filecontent);
                    }
                    else
                    {
                         
                        response = new Config.DefaultResponse(500, "Failed to download", responseText);
                        return BadRequest(response);
                    }

                }

            }
            else
            {
                response = new Config.DefaultResponse(500, "Failed", "Something Wrong occured");
                return BadRequest(response);
            }
        }
        catch
        {
            response = new Config.DefaultResponse(500, "Failed", "Something Wrong occured");
            return BadRequest(response);
        }
    }

    //delete file
    [HttpPost("file:delete")]
    public async Task<IActionResult> DeleteFile([FromBody] FileUpload Upload)
    {
        Config.DefaultResponse response = new Config.DefaultResponse();
        try
        {
            string siteUrl = "";
            var relativeFilePath = "/sites/NavisionBusinessCentral/Shared Documents" + Upload.FolderName + "/" + Upload.FileName;
            var apiUrl = $"{siteUrl}/_api/web/GetFileByServerRelativePath(decodedurl='{relativeFilePath}')";

            var accessTokenResult = await GetAccessToken();

            if (accessTokenResult is ObjectResult accessTokenObjectResult && accessTokenObjectResult.Value is string accessToken)
            {
                using (var httpClient = new HttpClient())
                {
                    httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
                    httpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

                    // Send HTTP DELETE request to delete the file
                    var deleteResponse = await httpClient.DeleteAsync(apiUrl);

                    if (deleteResponse.IsSuccessStatusCode)
                    {
                        var responseText = await deleteResponse.Content.ReadAsStringAsync();
                        response = new Config.DefaultResponse(200, "Success", "File deleted successfully");
                        return Ok(response);
                    }
                    else
                    {
                        var responseText = await deleteResponse.Content.ReadAsStringAsync();
                        response = new Config.DefaultResponse(500, "Failed to delete file", responseText);
                        return BadRequest(response);
                    }
                }
            }
            else
            {
                response = new Config.DefaultResponse(500, "Failed", "Something wrong occurred");
                return BadRequest(response);
            }
        }
        catch
        {
            response = new Config.DefaultResponse(500, "Failed", "Something wrong occurred");
            return BadRequest(response);
        }
    }


}
