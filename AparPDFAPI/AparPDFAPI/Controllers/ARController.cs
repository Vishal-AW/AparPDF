using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using System.IO;
using System.Web;
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.html;
using iTextSharp.text.html.simpleparser;
using System.Data;
using System.Configuration;
using System.Security;
using Microsoft.SharePoint;
using System.Collections.Specialized;
//using Microsoft.SharePoint.Utilities;
using Microsoft.SharePoint.Client;
using System.Web.Script.Serialization;
using Newtonsoft.Json.Linq;
using System.Net.Sockets;
using AparPDFAPI.Core.Utilities;

namespace AparPDFAPI.Controllers
{
    public class ARController : ApiController
    {
        [HttpGet]
        [Route("api/PDFImageAR/GetPDFGenrate/{ActionName}/{LoginName}/{EmailID}/{Comments}/{LID}/{DisplayNo}/{TotalApprover}/{CurrentApprover}/{Position}/{designation}")]
        public string GetPDFMyAction(string ActionName, string LoginName, string EmailID, string Comments, string LID, string DisplayNo, int TotalApprover, int CurrentApprover,int Position,string designation)
        {
            string returnmessage = "done";
            try
            {
                
                string Action = ActionName;
                string User = LoginName;
                string POnum = LID;
                string DispayPO = DisplayNo.Replace('@', '/');
               // string DispayPO = DisplayNo.Insert(2, "/").Insert(5, "-").Insert(8, "/");
                string footer = "DOCUMENT ARE SIGNED DIGITALLY, HENCE NO PHYSICAL SIGNATURE REQUIRED.";
                int B_P = Position;
                string SPToken = GetToken();
                string login = "sp.admin@apar.com";
                string PurchaseText = ConfigurationManager.AppSettings["ARText"];
                string WSiteName = ConfigurationManager.AppSettings["ARSiteName"];
                string Nextpagetext = ConfigurationManager.AppSettings["ARNewPage"];
                string designationName = "";

                if (designation == "Creator")
                {
                    designationName = ConfigurationManager.AppSettings["Creator"];
                }else if (designation == "Approver")
                {
                    designationName = ConfigurationManager.AppSettings["Approver"];
                }
                else if (designation == "Analyst")
                {
                    designationName = ConfigurationManager.AppSettings["analyst"];
                }
                else if (designation == "MD")
                {
                    designationName = ConfigurationManager.AppSettings["MD"];
                }
                else
                {
                    designationName = "";
                }


                using (var contextimage = new ClientContext(ConfigurationManager.AppSettings["MainSite"]))
                {
                    using (var context = new ClientContext(ConfigurationManager.AppSettings["ARSite"]))
                    {

                        #region Get Data From List
                        //string password = "zpsllhcvdfbfhgmk";
                        string password = ConfigurationManager.AppSettings["Password"];
                        SecureString sec_pass = new SecureString();
                        Array.ForEach(password.ToArray(), sec_pass.AppendChar);
                        sec_pass.MakeReadOnly();
                        context.Credentials = new SharePointOnlineCredentials(login, sec_pass);

                        contextimage.Credentials = new SharePointOnlineCredentials(login, sec_pass);

                        /***********Image*************/

                        var listdata = contextimage.Web.Lists.GetByTitle("EmployeSignature");

                        var listdataLogolist = contextimage.Web.Lists.GetByTitle("logoimages");


                        var items = listdata.GetItems(CamlQuery.CreateAllItemsQuery());
                        CamlQuery query = new CamlQuery();
                        query.ViewXml = "<View><Query><Where><Eq><FieldRef Name='EmployeeCode_x003a_Employee_x0020_Email' /><Value Type='Lookup'>" + EmailID + "</Value></Eq></Where></Query><OrderBy><FieldRef Name='FileLeafRef' /></OrderBy></View>";
                        ListItemCollection listitem = listdata.GetItems(query);

                        //// logo code
                        ///
                        var itemslogo = listdataLogolist.GetItems(CamlQuery.CreateAllItemsQuery());
                        CamlQuery querylogo = new CamlQuery();
                        query.ViewXml = "<View><Query><Where><Eq><FieldRef Name='FileLeafRef' /><Value Type='File'>logo.png</Value></Eq></Where></Query><OrderBy><FieldRef Name='FileLeafRef' /></OrderBy></View>";
                        ListItemCollection listdataLogo = listdataLogolist.GetItems(query);
                        contextimage.Load(listdataLogo);
                        contextimage.ExecuteQuery();
                        var signimagelogo = "";
                        foreach (var oListItem in listdataLogo)
                        {
                            signimagelogo = oListItem["FileLeafRef"].ToString();
                        }

                        /// logo code end



                        contextimage.Load(listitem);
                        contextimage.ExecuteQuery();
                        var signimage = "";
                        foreach (var oListItem in listitem)
                        {
                            signimage = oListItem["FileLeafRef"].ToString();
                        }

                        /************End************/

                        var subsitelistdata = context.Web.Lists.GetByTitle("ARDocument");


                        // var itemss = subsitelistdata.GetItems(CamlQuery.CreateAllItemsQuery());
                        CamlQuery query1 = new CamlQuery();

                        //query1.ViewXml = "<View><Query><Where><Eq><FieldRef Name='POReferenceNumber' LookupId='FALSE'/><Value Type='Lookup'>"+POnum+"</Value></Eq></Where></Query></view>";

                        query1.ViewXml = "<View><Query><Where><Eq><FieldRef Name='LID' LookupId='TRUE'/><Value Type='Lookup'>" + POnum + "</Value></Eq></Where></Query></view>";

                        ListItemCollection listitem1 = subsitelistdata.GetItems(query1);


                        context.Load(listitem1);
                        context.ExecuteQuery();
                        var docnm = "";
                        var path = "";
                        int noticline = B_P;

                        if (TotalApprover > 5)
                        {
                            if (CurrentApprover > 5)
                            {
                                B_P = B_P + 650;
                            }

                        }
                        int Xvalue = 0;
                        if (CurrentApprover == 1 || CurrentApprover == 6)
                        {
                            Xvalue = 15;
                        }
                        else if (CurrentApprover == 2 || CurrentApprover == 7)
                        {
                            Xvalue = 130;

                        }
                        else if (CurrentApprover == 3 || CurrentApprover == 8)
                        {
                            Xvalue = 245;

                        }
                        else if (CurrentApprover == 4 || CurrentApprover == 9)
                        {
                            Xvalue = 360;

                        }
                        else if (CurrentApprover == 5 || CurrentApprover == 10)
                        {
                            Xvalue = 475;
                        }


                        foreach (var oListItem1 in listitem1)
                        {
                            //var FileType = oListItem1["FileType"].ToString();
                            docnm = oListItem1["FileLeafRef"].ToString();
                            var docId = oListItem1["ID"].ToString();
                         //   var FileTypename = oListItem1["FileType"].ToString();
                            string ImgName = docnm;
                            int lastIndex = ImgName.LastIndexOf('.');
                            var Filenm = ImgName.Substring(0, lastIndex);
                            var Exttype = oListItem1["File_x0020_Type"].ToString();

                            // check current document is PDF or not
                            if (Exttype != "pdf" && Exttype != "PDF")
                            {

                                using (MemoryStream ms = new MemoryStream())
                                {
                                    // Document document = new Document(PageSize.A4, 25, 25, 30, 30);
                                    Document document = new Document(PageSize.A4.Rotate());

                                    PdfWriter writer = PdfWriter.GetInstance(document, ms);

                                    document.Open();

                                    //document.Add(new Paragraph("Hello World"));
                                    var docimg = WSiteName+ "/ARDocument/" + docnm + "";
                                    var docimg1 = WSiteName + "/ARDocument/" + Filenm + ".pdf";
                                    var fileimagetype = context.Web.GetFileByServerRelativeUrl(docimg);

                                    //var file = context.Web.GetFileByServerRelativeUrl(docimg);
                                    context.ExecuteQuery();
                                    //context.Load(writer);
                                    ClientResult<System.IO.Stream> Imagedata1 = fileimagetype.OpenBinaryStream();
                                    System.IO.MemoryStream imageStream1 = new System.IO.MemoryStream();
                                    context.ExecuteQuery();

                                    Imagedata1.Value.CopyTo(imageStream1);
                                    byte[] imgarray1 = imageStream1.ToArray();

                                    iTextSharp.text.Image sigimage1 = iTextSharp.text.Image.GetInstance(imgarray1);
                                    //sigimage1.SetAbsolutePosition(0, 0);
                                    //sigimage1.ScaleAbsolute(0,0);
                                    //sigimage1.ScaleToFit(150, 80);
                                    sigimage1.ScaleAbsolute(400f, 370f);
                                    //sigimage1.ScalePercent(95f);
                                    sigimage1.Alignment = iTextSharp.text.Image.ALIGN_CENTER;

                                    //sigimage1.HasAbsolutePosition();
                                    document.Add(sigimage1);
                                    

                                    document.Close();

                                    writer.Close();

                                    Microsoft.SharePoint.Client.ListItem oListItem = subsitelistdata.GetItemById(docId);

                                    oListItem.DeleteObject();

                                    context.ExecuteQuery();

                                    Web web = context.Web;
                                    Microsoft.SharePoint.Client.List list = web.Lists.GetByTitle("ARDocument");

                                    var fileCreationInformation = new FileCreationInformation();
                                    byte[] array1 = ms.ToArray();
                                    fileCreationInformation.Content = array1;
                                    fileCreationInformation.Overwrite = true;
                                    //fileCreationInformation.Url = documentListURL + documentName;
                                    fileCreationInformation.Url = docimg1;
                                    Microsoft.SharePoint.Client.File uploadFile = list.RootFolder.Files.Add(fileCreationInformation);
                                    uploadFile.ListItemAllFields["LID"] = POnum;
                                    // uploadFile.ListItemAllFields["ExpVoucherNo"] = 27;
                                    uploadFile.ListItemAllFields.Update();
                                    context.ExecuteQuery();

                                    path = WSiteName + "/ARDocument/" + Filenm + ".pdf";

                                }
                            }

                            else
                            {
                                path = WSiteName + "/ARDocument/" + docnm + "";
                            }

                            var file = context.Web.GetFileByServerRelativeUrl(path);

                            var image = "/EmployeSignature/" + signimage + "";
                            var fileimage = contextimage.Web.GetFileByServerRelativeUrl(image);
                            var imagelogo = "/logoimages/" + signimagelogo + "";
                            var fileimageLogo = contextimage.Web.GetFileByServerRelativeUrl(imagelogo);

                            context.Load(file);
                            context.ExecuteQuery();

                            contextimage.ExecuteQuery();
                            ClientResult<System.IO.Stream> data = file.OpenBinaryStream();
                            ClientResult<System.IO.Stream> Imagedata = fileimage.OpenBinaryStream();
                            ClientResult<System.IO.Stream> ImagedataLogo = fileimageLogo.OpenBinaryStream();

                            context.Load(file);
                            context.ExecuteQuery();
                            contextimage.ExecuteQuery();


                            #endregion

                             

                            int[] PO_X = new int[10];
                            int[] PO_Y = new int[10];

                            int[] OT_X = new int[10];
                            int[] OT_Y = new int[10];

                          
                            PO_X[0] = Xvalue; PO_Y[0] = B_P + 60;
                            PO_X[1] = Xvalue; PO_Y[1] = B_P + 18;
                            PO_X[2] = Xvalue; PO_Y[2] = B_P + 12;
                            PO_X[3] = Xvalue; PO_Y[3] = B_P + 8;
                            PO_X[4] = Xvalue; PO_Y[4] = B_P + 4;
                            PO_X[5] = Xvalue; PO_Y[5] = noticline;

                            System.IO.MemoryStream outputStream = new System.IO.MemoryStream();
                            System.IO.MemoryStream imageStream = new System.IO.MemoryStream();
                            System.IO.MemoryStream imageStreamLogo = new System.IO.MemoryStream();

                            string textPDF = string.Empty;
                            using (System.IO.MemoryStream mStream = new System.IO.MemoryStream())
                            {
                                if (data != null)
                                {
                                    String pathout = "";
                                    data.Value.CopyTo(mStream);
                                    byte[] array = mStream.ToArray();

                                    if (CurrentApprover == 6)
                                    {
                                        PdfReader readertemp = new PdfReader(array);

                                        array = AddDocumentPages(array, readertemp.NumberOfPages);
                                    }

                                    Imagedata.Value.CopyTo(imageStream);
                                    byte[] imgarray = imageStream.ToArray();


                                    ImagedataLogo.Value.CopyTo(imageStreamLogo);
                                    byte[] imgarrayLogo = imageStreamLogo.ToArray();

                                    PdfReader reader = new PdfReader(array);

                                    //select three pages from the original document
                                    // reader.SelectPages("2");
                                    int n = reader.NumberOfPages;

                                    //create PdfStamper object to write to get the pages from reader 
                                    PdfStamper stamper = new PdfStamper(reader, outputStream);

                                    PdfContentByte pbover = stamper.GetOverContent(n);

                                    if (CurrentApprover == 1)
                                    {
                                        for (int y = 1; y < reader.NumberOfPages + 1; y++)
                                        {
                                            pbover = stamper.GetOverContent(y);
                                            var lable = FontFactory.GetFont("Arial", 8, Color.BLACK);

                                            iTextSharp.text.Image logo = iTextSharp.text.Image.GetInstance(imgarrayLogo);
                                            logo.SetAbsolutePosition(15, 780);
                                            pbover.AddImage(logo);
                                            ColumnText.ShowTextAligned(pbover, Element.ALIGN_LEFT, new Phrase(new Chunk("AR refrence no " + DispayPO, lable)), 445, 780, 0);

                                        }
                                    }

                                    if (CurrentApprover == 6)
                                    {
                                         
                                            pbover = stamper.GetOverContent(n);
                                            var lable = FontFactory.GetFont("Arial", 8, Color.BLACK);

                                            iTextSharp.text.Image logo = iTextSharp.text.Image.GetInstance(imgarrayLogo);
                                            logo.SetAbsolutePosition(15, 780);
                                            pbover.AddImage(logo);
                                            ColumnText.ShowTextAligned(pbover, Element.ALIGN_LEFT, new Phrase(new Chunk("AR refrence no " + DispayPO, lable)), 445, 780, 0);

                                        
                                    }

                                    pbover = stamper.GetOverContent(n);



                                    DateTime dateTime = DateTime.Now;

                                    var blackListTextFont = FontFactory.GetFont("Arial", 4, Color.BLACK);



                                    ColumnText.ShowTextAligned(pbover, Element.ALIGN_LEFT, new Phrase(new Chunk(PurchaseText, blackListTextFont)), PO_X[0], PO_Y[0], 0);
                                    if(ActionName != "Rejected")
                                    {
                                        iTextSharp.text.Image sigimage = iTextSharp.text.Image.GetInstance(imgarray);
                                        sigimage.SetAbsolutePosition(PO_X[1], PO_Y[1]);
                                        pbover.AddImage(sigimage);

                                    }
                                    else
                                    {
                                        ColumnText.ShowTextAligned(pbover, Element.ALIGN_LEFT, new Phrase(new Chunk("Rejected", FontFactory.GetFont("Arial", 20, Color.BLACK))), PO_X[1], PO_Y[1] + 10, 0);
                                        ColumnText.ShowTextAligned(pbover, Element.ALIGN_LEFT, new Phrase(new Chunk("Rejected By  " + User, FontFactory.GetFont("Arial", 8, Color.BLACK))), 450, 770, 0);

                                    }

                                    ColumnText.ShowTextAligned(pbover, Element.ALIGN_LEFT, new Phrase(new Chunk(User, blackListTextFont)), PO_X[2], PO_Y[2], 0);
                                    ColumnText.ShowTextAligned(pbover, Element.ALIGN_LEFT, new Phrase(new Chunk(Convert.ToString(dateTime), blackListTextFont)), PO_X[3], PO_Y[3], 0);
                                    ColumnText.ShowTextAligned(pbover, Element.ALIGN_LEFT, new Phrase(new Chunk(designationName, blackListTextFont)), PO_X[4], PO_Y[4], 0);
                                    if (CurrentApprover == 1)
                                    {
                                        ColumnText.ShowTextAligned(pbover, Element.ALIGN_LEFT, new Phrase(new Chunk(footer, blackListTextFont)), PO_X[5], PO_Y[5], 0);
                                    }


                                   

                                    PdfContentByte pbunder = stamper.GetUnderContent(n);

                                    if (CurrentApprover == 6)
                                    {
                                        int m = n;
                                        PdfContentByte pbover1 = stamper.GetOverContent(m);
                                        var lable = FontFactory.GetFont("Arial", 8, Color.BLACK);
                                        ColumnText.ShowTextAligned(pbover1, Element.ALIGN_LEFT, new Phrase(new Chunk(Nextpagetext + (m-1) , lable)), 445, 790, 0);

                                        PdfContentByte pbunder1 = stamper.GetUnderContent(m);
                                    }

                                    stamper.Close();

                                    
                                   

                                    #region Update PDF Code
                                    byte[] array1 = outputStream.ToArray();
                                    HttpWebRequest endpointRequest = (HttpWebRequest)HttpWebRequest.Create(ConfigurationManager.AppSettings["ARSite"].ToString() + "/_api/web/GetFolderByServerRelativeUrl('ARDocument')/Files/add(url='" + ImgName + "',overwrite=true)");
                                    endpointRequest.Method = "POST";
                                    endpointRequest.Headers.Add("binaryStringRequestBody", "true");
                                    endpointRequest.Headers.Add("Authorization", "Bearer " + SPToken);
                                    endpointRequest.GetRequestStream().Write(array1, 0, array1.Length);
                                    HttpWebResponse endpointresponse = (HttpWebResponse)endpointRequest.GetResponse();
                                    endpointresponse.Close();
                                    #endregion



                                }


                            }




                        }
                    }
                }
            }
            catch (Exception ex)
            {
                returnmessage = ex.ToString ();
            }


            return returnmessage;
        }

        private static byte[] AddDocumentPages(byte[] pdf, int pages)
        {
            var reader = new PdfReader(pdf);
            System.IO.MemoryStream ms = new System.IO.MemoryStream();
            PdfStamper stamper = new PdfStamper(reader, ms);
            var numberofPages = reader.NumberOfPages;
            var rectangle = reader.GetPageSize(1);
            stamper.InsertPage(numberofPages + 1, rectangle);
            reader.Close();
            stamper.Close();
            ms.Flush();
            return ms.GetBuffer();

        }


        public static string GetToken()
        {

            Uri webUri = new Uri(ConfigurationManager.AppSettings["URLName"]);

            string realm = TokenHelper.GetRealmFromTargetUrl(webUri);

            var SharePointPrincipalId = "00000003-0000-0ff1-ce00-000000000000";
            var token = TokenHelper.GetAppOnlyAccessToken(SharePointPrincipalId, webUri.Authority, realm).AccessToken;


            return token;
        }

    }
}
