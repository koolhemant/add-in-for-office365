using Microsoft.SharePoint.Client;
using SPSignicatDemoWeb.DocumentService;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace SPSignicatDemoWeb.Controllers
{
    public class HomeController : Controller
    {
        private static readonly string PASSWORD = "<INSERT PASSWORD FOR SERVICE>";
        private static readonly string SERVICE = "shared";

        public ActionResult Ping()
        {
            return View();
        }

        [SharePointContextFilter]
        public ActionResult Index(string SPHostURL, string SPListItemId, string SPListId, string SPSource, string SPListURLDir, string SPItemUrl)
        {
            // Let's save some tokens and stuff for later
            Response.SetCookie(new HttpCookie("SPAppToken", Request.Form["SPAppToken"]));
            Response.SetCookie(new HttpCookie("SPHostURL", SPHostURL));

            // Add selected item info to the viewbag so we can pass it to
            // the Sign method when the user has selected signature method
            ViewBag.SPHostURL = SPHostURL;
            ViewBag.SPListItemId = SPListItemId;
            ViewBag.SPListId = SPListId;
            ViewBag.SPSource = SPSource;
            ViewBag.SPListURLDir = SPListURLDir;
            ViewBag.SPItemUrl = SPItemUrl;
            return View();
        }

        [SharePointContextFilter]
        public ActionResult Sign(string SPHostURL, string SPListItemId, string SPListId, string SPSource, string SPListURLDir, string SPItemUrl, string Method)
        {
            // Prepare the document
            var spContext = SharePointContextProvider.Current.GetSharePointContext(HttpContext);
            var document = GetDocumentForSigningFromSPList(spContext, SPListId, SPListItemId);

            // Prepare the request to the Signicat document service
            var request = GetRequest(document, Method, SPListId, SPSource, SPListURLDir);

            // Send the request to Signicat
            createrequestresponse response;
            using (var client = new DocumentEndPointClient())
            {
                response = client.createRequest(request);
            }

            // And then forward the end user to the signature process URL
            String signHereUrl =
                String.Format("https://{0}.signicat.com/std/docaction/{1}?request_id={2}&task_id={3}&artifact={4}",
                              "preprod",
                              "signicat",
                              response.requestid[0],
                              request.request[0].task[0].id,
                              response.artifact);
            return Redirect(signHereUrl);
        }

        public ActionResult Return(string rid, string SPListId, string documentName, string SPSource, string SPListURLDir)
        {
            // Get the tokens
            string spAppToken = Request.Cookies["SPAppToken"].Value;
            string spHostURL = Request.Cookies["SPHostURL"].Value;

            using (var client = new DocumentEndPointClient())
            {
                var getStatusResponse = client.getStatus(new getstatusrequest
                {
                    password = PASSWORD,
                    service = SERVICE,
                    requestid = new string[] { rid }
                });

                // If there's no signature result available, we assume it's been cancelled
                // and just send the user back to the document list he/she came from
                if (getStatusResponse.First().documentstatus == null)
                {
                    return Redirect(SPSource);
                }

                // Otherwise, we get the URI to the SDO (signed document object)
                string resultUri = getStatusResponse.First().documentstatus.First().resulturi;

                // Let's just use an old school HTTP WebClient with Basic authentication
                // to download the signed document from Signicat Session Data Storage
                byte[] sdo;
                string contentType;
                using (var webClient = new System.Net.WebClient())
                {
                    webClient.Credentials = new System.Net.NetworkCredential(SERVICE, PASSWORD);
                    sdo = webClient.DownloadData(resultUri);
                    contentType = webClient.ResponseHeaders["Content-Type"];
                }

                // Now, let's upload it to the same document list as the original
                SharePointContextToken contextToken = TokenHelper.ReadAndValidateContextToken(spAppToken, Request.Url.Authority);
                string accessToken = TokenHelper.GetAccessToken(contextToken, new Uri(spHostURL).Authority).AccessToken;
                using (var clientContext = TokenHelper.GetClientContextWithAccessToken(spHostURL.ToString(), accessToken))
                {
                    if (clientContext != null)
                    {
                        List list = clientContext.Web.Lists.GetById(Guid.Parse(SPListId));
                        var fileCreationInformation = new FileCreationInformation
                        {
                            Content = sdo,
                            Overwrite = true,
                            Url = spHostURL + SPListURLDir + "/" + documentName + " - SIGNED" + GetFileExtensionForContentType(contentType)
                        };
                        var uploadFile = list.RootFolder.Files.Add(fileCreationInformation);
                        uploadFile.ListItemAllFields.Update();
                        clientContext.ExecuteQuery();
                    }
                }
            }
            return Redirect(SPSource);
        }

        // Grabs a document from a Sharepoint list and turns it into a "provided document",
        // i.e. a document which is to provided to Signicat in a web service request
        // and used as the original in a signing process
        private provideddocument GetDocumentForSigningFromSPList(SharePointContext spContext, string SPListId, string SPListItemId)
        {
            // Create a provided document, prepared for signing with basic properties
            var document = new provideddocument
            {
                externalreference = SPListItemId,
                id = "doc_1",
                mimetype = "application/pdf",
                signtextentry = "I have read and understood the document which has been presented to me, and I agree to its contents with my signature."
            };

            // Now, let's get the name and data of the document from the Sharepoint list
            using (var clientContext = spContext.CreateUserClientContextForSPHost())
            {
                if (clientContext != null)
                {
                    // Get the list item representing the document
                    List list = clientContext.Web.Lists.GetById(Guid.Parse(SPListId));
                    ListItem listItem = list.GetItemById(SPListItemId);
                    clientContext.Load(listItem, li => li.DisplayName, li => li.File);
                    clientContext.ExecuteQuery();

                    // Set the name of the document
                    document.description = listItem.DisplayName;

                    // Now get the binary contents
                    Microsoft.SharePoint.Client.File itemFile = listItem.File;
                    ClientResult<Stream> fileStream = listItem.File.OpenBinaryStream();
                    clientContext.Load(itemFile);
                    clientContext.ExecuteQuery();
                    using (var memoryStream = new MemoryStream())
                    {
                        fileStream.Value.CopyTo(memoryStream);
                        document.data = memoryStream.ToArray();
                    }
                }
            }
            return document;
        }

        // Creates a "document order" request, i.e. a mapping to the object model provided
        // by the DocumentService, a Signicat web service for administering signature processes.
        private createrequestrequest GetRequest(provideddocument document, string method, string spListId, string spSource, string spListUrlDir)
        {
            string queryString =
                "?rid=${requestId}&" +
                string.Format("SPListId={0}&documentName={1}&SPSource={2}&SPListURLDir={3}",
                Url.Encode(spListId), Url.Encode(document.description), Url.Encode(spSource), Url.Encode(spListUrlDir));

            var request = new createrequestrequest
            {
                password = PASSWORD,
                service = SERVICE,
                request = new request[]
                {
                    new request
                    {
                        profile = "slim",
                        language = "en",
                        document = new document[]
                        {
                            document
                        },
                        task = new task[]
                        {
                            new task
                            {
                                bundle = false,
                                bundleSpecified = true,
                                id = "task_1",
                                ontaskcomplete = Url.Action("Return", "Home", null, Request.Url.Scheme) + queryString,
                                ontaskcancel = Url.Action("Return", "Home", null, Request.Url.Scheme) + queryString,
                                ontaskpostpone = Url.Action("Return", "Home", null, Request.Url.Scheme) + queryString,
                                documentaction = new documentaction[]
                                {
                                    new documentaction
                                    {
                                        type = documentactiontype.sign,
                                        documentref = document.id,
                                        sendresulttoarchive = true,
                                        sendresulttoarchiveSpecified = true
                                    }
                                },
                                signature = new signature[]
                                {
                                    new signature
                                    {
                                        method = new string[]
                                        {
                                            method
                                        }
                                    }
                                },
                                authentication = new authentication { artifact = true, artifactSpecified = true }
                            }
                        }
                    }
                }
            };
            return request;
        }

        // Well this might seem a weird. Usually the SDO format and extension
        // is already known in advance, but for various hacky demo reasons
        // we kind of have to figure it out here.
        private string GetFileExtensionForContentType(string contentType)
        {
            if (contentType.StartsWith("application/x-ltv-sdo"))
            {
                return ".xml";
            }
            return ".pdf";
        }
    }
}
