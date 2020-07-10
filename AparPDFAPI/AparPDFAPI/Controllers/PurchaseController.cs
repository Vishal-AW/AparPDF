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

namespace AparPDFAPI.Controllers
{
    public class PurchaseController : ApiController
    {
        [HttpGet]
        [Route("api/PDFImageNew/GetPDFGenrate/{ActionName}/{LoginName}/{EmailID}/{Comments}/{PONumber}/{PODisplayNo}/{UserType}/{letterheadType}")]
        public string GetPDFMyAction(string ActionName, string LoginName, string EmailID, string Comments, string PONumber, string PODisplayNo, string UserType, string letterheadType)
        {

            string Action = ActionName;
            string User = LoginName;
            string POnum = PONumber;
            string DispayPO = PODisplayNo;
            string footer = "DOCUMENT ARE SIGNED DIGITALLY, HENCE NO PHYSICAL SIGNATURE REQUIRED.";
            string letterhead = letterheadType;

           

            string login = "sp.admin@apar.com"; //give your username here  
            string PurchaseText = ConfigurationManager.AppSettings["PurchaseText"];


            using (var contextimage = new ClientContext("https://aparindltd.sharepoint.com"))
            {


                using (var context = new ClientContext("https://aparindltd.sharepoint.com/PurchaseOrder"))
                {




                    #region Get Data From List
                    //string password = "zpsllhcvdfbfhgmk";
                    string password = "kzqxvzgmkgwmpjmp";
                    SecureString sec_pass = new SecureString();
                    Array.ForEach(password.ToArray(), sec_pass.AppendChar);
                    sec_pass.MakeReadOnly();
                    context.Credentials = new SharePointOnlineCredentials(login, sec_pass);

                    contextimage.Credentials = new SharePointOnlineCredentials(login, sec_pass);

                    /***********Image*************/

                    var listdata = contextimage.Web.Lists.GetByTitle("EmployeSignature");


                    var items = listdata.GetItems(CamlQuery.CreateAllItemsQuery());
                    CamlQuery query = new CamlQuery();
                    //query.ViewXml = "<View><Query><Eq><FieldRef Name='EmployeeUserName' LookupId='TRUE'/><Value Type='User'>" + User + "</Value></Eq></Query><OrderBy><FieldRef Name='FileLeafRef' /></OrderBy></View><ViewFields><FieldRef Name='ID' /><FieldRef Name='FileLeafRef' /><FieldRef Name='FileDirRef' /></ViewFields><QueryOptions><ViewAttributes Scope='Recursive' /><OptimizeFor>FolderUrls</OptimizeFor></QueryOptions>";
                    //query.ViewXml = "<View><Query><Eq><FieldRef Name='EmployeeCode_x003a_Employee_x0020_Email' /><Value Type='Lookup'>kiran.sawant@apar.com</Value></Eq></Query><OrderBy><FieldRef Name='FileLeafRef' /></OrderBy></View><ViewFields><FieldRef Name='ID' /><FieldRef Name='FileLeafRef' /><FieldRef Name='FileDirRef' /></ViewFields><QueryOptions><ViewAttributes Scope='Recursive' /><OptimizeFor>FolderUrls</OptimizeFor></QueryOptions>";
                    query.ViewXml = "<View><Query><Where><Eq><FieldRef Name='EmployeeCode_x003a_Employee_x0020_Email' /><Value Type='Lookup'>" + EmailID + "</Value></Eq></Where></Query><OrderBy><FieldRef Name='FileLeafRef' /></OrderBy></View>";


                    ListItemCollection listitem = listdata.GetItems(query);


                    contextimage.Load(listitem);
                    contextimage.ExecuteQuery();
                    var signimage = "";
                    foreach (var oListItem in listitem)
                    {
                        signimage = oListItem["FileLeafRef"].ToString();
                    }

                    /************End************/

                    var subsitelistdata = context.Web.Lists.GetByTitle("PurchaseDocuments");


                    // var itemss = subsitelistdata.GetItems(CamlQuery.CreateAllItemsQuery());
                    CamlQuery query1 = new CamlQuery();

                    //query1.ViewXml = "<View><Query><Where><Eq><FieldRef Name='POReferenceNumber' LookupId='FALSE'/><Value Type='Lookup'>"+POnum+"</Value></Eq></Where></Query></view>";

                    query1.ViewXml = "<View><Query><Where><Eq><FieldRef Name='POReferenceNumber' LookupId='TRUE'/><Value Type='Lookup'>" + POnum + "</Value></Eq></Where></Query></view>";

                    ListItemCollection listitem1 = subsitelistdata.GetItems(query1);


                    context.Load(listitem1);
                    context.ExecuteQuery();
                    var docnm = "";
                    var path = "";
                    foreach (var oListItem1 in listitem1)
                    {
                        var FileType = oListItem1["FileType"].ToString();
                        docnm = oListItem1["FileLeafRef"].ToString();
                        var docId = oListItem1["ID"].ToString();
                        var FileTypename = oListItem1["FileType"].ToString();


                        string ImgName = docnm;
                        int lastIndex = ImgName.LastIndexOf('.');
                        var Filenm = ImgName.Substring(0, lastIndex);
                       
                        var Exttype = oListItem1["File_x0020_Type"].ToString();
                       

                        if (Exttype != "pdf" && Exttype != "PDF")
                        {

                            using (MemoryStream ms = new MemoryStream())
                            {
                                // Document document = new Document(PageSize.A4, 25, 25, 30, 30);
                                Document document = new Document(PageSize.A4.Rotate());

                                PdfWriter writer = PdfWriter.GetInstance(document, ms);

                                document.Open();

                                //document.Add(new Paragraph("Hello World"));
                                var docimg = "/PurchaseOrder/PurchaseDocuments/" + docnm + "";
                                var docimg1 = "/PurchaseOrder/PurchaseDocuments/" + Filenm + ".pdf";
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


                                //HttpContext.Current.Response.ContentType = "pdf/application";

                                //HttpContext.Current.Response.AddHeader("content-disposition",
                                //"attachment;filename=First PDF document.pdf");
                                //HttpContext.Current.Response.OutputStream.Write(ms.GetBuffer(), 0, ms.GetBuffer().Length);

                                string siteURL = "https://aparindltd.sharepoint.com";
                                string documentListName = "PurchaseDocuments";
                                string documentListURL = "https://aparindltd.sharepoint.com/PurchaseOrder/PurchaseDocuments/";
                                //string documentName = "11111_Airnet.pdf";


                                Web web = context.Web;
                                Microsoft.SharePoint.Client.List list = web.Lists.GetByTitle("PurchaseDocuments");

                                var fileCreationInformation = new FileCreationInformation();
                                byte[] array1 = ms.ToArray();
                                fileCreationInformation.Content = array1;
                                fileCreationInformation.Overwrite = true;
                                //fileCreationInformation.Url = documentListURL + documentName;
                                fileCreationInformation.Url = docimg1;
                                Microsoft.SharePoint.Client.File uploadFile = list.RootFolder.Files.Add(fileCreationInformation);
                                uploadFile.ListItemAllFields["FileType"] = FileTypename;
                                uploadFile.ListItemAllFields["POReferenceNumber"] = POnum;
                                // uploadFile.ListItemAllFields["ExpVoucherNo"] = 27;
                                uploadFile.ListItemAllFields.Update();
                                context.ExecuteQuery();

                                path = "/PurchaseOrder/PurchaseDocuments/" + Filenm + ".pdf";






                            }
                        }

                        else
                        {


                            path = "/PurchaseOrder/PurchaseDocuments/" + docnm + "";



                        }






                        //path = "/PurchaseOrder/PurchaseDocuments/" + docnm + "";

                        var file = context.Web.GetFileByServerRelativeUrl(path);
                        //var path = "/PurchaseOrder/PurchaseDocuments/11111_Airnet.pdf";

                        //var file = context.Web.GetFileByServerRelativeUrl(path);




                        var image = "/EmployeSignature/" + signimage + "";
                        var fileimage = contextimage.Web.GetFileByServerRelativeUrl(image);

                        //var image = "/PurchaseOrder/PurchaseDocuments/image.jpg";
                        //var fileimage = context.Web.GetFileByServerRelativeUrl(image);

                        context.Load(file);
                        context.ExecuteQuery();


                        contextimage.ExecuteQuery();

                        ClientResult<System.IO.Stream> data = file.OpenBinaryStream();

                        ClientResult<System.IO.Stream> Imagedata = fileimage.OpenBinaryStream();


                        context.Load(file);
                        context.ExecuteQuery();
                        contextimage.ExecuteQuery();
                        #endregion

                        int[] PO_X = new int[10];
                        int[] PO_Y = new int[10];

                        int[] OT_X = new int[10];
                        int[] OT_Y = new int[10];

                        #region ENIDept
                        if (letterhead == "ENIDept")
                        {
                            if (UserType == "Head")
                            {
                                if (FileType == "PO")
                                {
                                    PO_X[0] = 15; PO_Y[0] = 216;
                                    PO_X[1] = 15; PO_Y[1] = 173;
                                    PO_X[2] = 15; PO_Y[2] = 162;
                                    PO_X[3] = 15; PO_Y[3] = 153;
                                    PO_X[4] = 15; PO_Y[4] = 145;
                                }
                                else
                                {
                                    PO_X[0] = 15; PO_Y[0] = 235;
                                    PO_X[1] = 15; PO_Y[1] = 225;
                                    PO_X[2] = 15; PO_Y[2] = 182;
                                    PO_X[3] = 15; PO_Y[3] = 171;
                                    PO_X[4] = 15; PO_Y[4] = 162;
                                    PO_X[5] = 15; PO_Y[5] = 153;
                                    PO_X[6] = 15; PO_Y[6] = 145;

                                }
                            }
                            if (UserType == "Head1")
                            {
                                if (FileType == "PO")
                                {
                                    PO_X[0] = 130; PO_Y[0] = 216;
                                    PO_X[1] = 130; PO_Y[1] = 173;
                                    PO_X[2] = 130; PO_Y[2] = 162;
                                    PO_X[3] = 130; PO_Y[3] = 153;
                                    PO_X[4] = 130; PO_Y[4] = 145;
                                }
                                else
                                {
                                    PO_X[0] = 130; PO_Y[0] = 235;
                                    PO_X[1] = 130; PO_Y[1] = 225;
                                    PO_X[2] = 130; PO_Y[2] = 182;
                                    PO_X[3] = 130; PO_Y[3] = 171;
                                    PO_X[4] = 130; PO_Y[4] = 162;
                                    PO_X[5] = 130; PO_Y[5] = 153;
                                    PO_X[6] = 130; PO_Y[6] = 145;

                                }
                            }

                            if (UserType == "Head2")
                            {
                                if (FileType == "PO")
                                {
                                    PO_X[0] = 245; PO_Y[0] = 216;
                                    PO_X[1] = 245; PO_Y[1] = 173;
                                    PO_X[2] = 245; PO_Y[2] = 162;
                                    PO_X[3] = 245; PO_Y[3] = 153;
                                    PO_X[4] = 245; PO_Y[4] = 145;
                                }
                                else
                                {
                                    PO_X[0] = 245; PO_Y[0] = 235;
                                    PO_X[1] = 245; PO_Y[1] = 225;
                                    PO_X[2] = 245; PO_Y[2] = 182;
                                    PO_X[3] = 245; PO_Y[3] = 171;
                                    PO_X[4] = 245; PO_Y[4] = 162;
                                    PO_X[5] = 245; PO_Y[5] = 153;
                                    PO_X[6] = 245; PO_Y[6] = 145;

                                }
                                }

                            if (UserType == "Head3")
                            {
                                if (FileType == "PO")
                                {
                                    PO_X[0] = 360; PO_Y[0] = 216;
                                    PO_X[1] = 360; PO_Y[1] = 173;
                                    PO_X[2] = 360; PO_Y[2] = 162;
                                    PO_X[3] = 360; PO_Y[3] = 153;
                                    PO_X[4] = 360; PO_Y[4] = 145;
                                }
                                else
                                {
                                    PO_X[0] = 360; PO_Y[0] = 235;
                                    PO_X[1] = 360; PO_Y[1] = 225;
                                    PO_X[2] = 360; PO_Y[2] = 182;
                                    PO_X[3] = 360; PO_Y[3] = 171;
                                    PO_X[4] = 360; PO_Y[4] = 162;
                                    PO_X[5] = 360; PO_Y[5] = 153;
                                    PO_X[6] = 360; PO_Y[6] = 145;

                                }
                            }

                            if (UserType == "Head4")
                            {
                                if (FileType == "PO")
                                {
                                    PO_X[0] = 475; PO_Y[0] = 216;
                                    PO_X[1] = 475; PO_Y[1] = 173;
                                    PO_X[2] = 475; PO_Y[2] = 162;
                                    PO_X[3] = 475; PO_Y[3] = 153;
                                    PO_X[4] = 476; PO_Y[4] = 145;
                                }
                                else
                                {
                                    PO_X[0] = 475; PO_Y[0] = 235;
                                    PO_X[1] = 475; PO_Y[1] = 225;
                                    PO_X[2] = 475; PO_Y[2] = 182;
                                    PO_X[3] = 475; PO_Y[3] = 171;
                                    PO_X[4] = 475; PO_Y[4] = 162;
                                    PO_X[5] = 475; PO_Y[5] = 153;
                                    PO_X[6] = 475; PO_Y[6] = 145;

                                }
                            }

                        }


                        #endregion
                 
                        #region ENI
                        if (letterhead == "ENI")
                        {
                            if (UserType == "Head")
                            {
                                int x = 15;
                                if (FileType == "PO")
                                { 
                                    PO_X[0] = x; PO_Y[0] = 177;
                                    PO_X[1] = x; PO_Y[1] = 134;
                                    PO_X[2] = x; PO_Y[2] = 126;
                                    PO_X[3] = x; PO_Y[3] = 117;
                                    PO_X[4] = x; PO_Y[4] = 110;
                                }
                                else
                                {
                                    
                                    PO_X[0] = x; PO_Y[0] = 195;
                                    PO_X[1] = x; PO_Y[1] = 188;
                                    PO_X[2] = x; PO_Y[2] = 145;
                                    PO_X[3] = x; PO_Y[3] = 134;
                                    PO_X[4] = x; PO_Y[4] = 126;
                                    PO_X[5] = x; PO_Y[5] = 117;
                                    PO_X[6] = x; PO_Y[6] = 110;

                                }
                            }
                            if (UserType == "Head1")
                            {
                                int x = 130;
                                if (FileType == "PO")
                                {
                                    PO_X[0] = x; PO_Y[0] = 177;
                                    PO_X[1] = x; PO_Y[1] = 134;
                                    PO_X[2] = x; PO_Y[2] = 126;
                                    PO_X[3] = x; PO_Y[3] = 117;
                                    PO_X[4] = x; PO_Y[4] = 110;
                                }
                                else
                                {

                                    PO_X[0] = x; PO_Y[0] = 195;
                                    PO_X[1] = x; PO_Y[1] = 188;
                                    PO_X[2] = x; PO_Y[2] = 145;
                                    PO_X[3] = x; PO_Y[3] = 134;
                                    PO_X[4] = x; PO_Y[4] = 126;
                                    PO_X[5] = x; PO_Y[5] = 117;
                                    PO_X[6] = x; PO_Y[6] = 110;

                                }
                            }
                            if (UserType == "Head2")
                            {
                                int x = 245;
                                if (FileType == "PO")
                                {
                                    PO_X[0] = x; PO_Y[0] = 177;
                                    PO_X[1] = x; PO_Y[1] = 134;
                                    PO_X[2] = x; PO_Y[2] = 126;
                                    PO_X[3] = x; PO_Y[3] = 117;
                                    PO_X[4] = x; PO_Y[4] = 110;
                                }
                                else
                                {

                                    PO_X[0] = x; PO_Y[0] = 195;
                                    PO_X[1] = x; PO_Y[1] = 188;
                                    PO_X[2] = x; PO_Y[2] = 145;
                                    PO_X[3] = x; PO_Y[3] = 134;
                                    PO_X[4] = x; PO_Y[4] = 126;
                                    PO_X[5] = x; PO_Y[5] = 117;
                                    PO_X[6] = x; PO_Y[6] = 110;

                                }
                            }
                            if (UserType == "Head3")
                            {
                                int x = 360;
                                if (FileType == "PO")
                                {
                                    PO_X[0] = x; PO_Y[0] = 177;
                                    PO_X[1] = x; PO_Y[1] = 134;
                                    PO_X[2] = x; PO_Y[2] = 126;
                                    PO_X[3] = x; PO_Y[3] = 117;
                                    PO_X[4] = x; PO_Y[4] = 110;
                                }
                                else
                                {

                                    PO_X[0] = x; PO_Y[0] = 195;
                                    PO_X[1] = x; PO_Y[1] = 188;
                                    PO_X[2] = x; PO_Y[2] = 145;
                                    PO_X[3] = x; PO_Y[3] = 134;
                                    PO_X[4] = x; PO_Y[4] = 126;
                                    PO_X[5] = x; PO_Y[5] = 117;
                                    PO_X[6] = x; PO_Y[6] = 110;

                                }
                            }
                       
                            if (UserType == "Head4")
                            {
                                int x = 475;
                                if (FileType == "PO")
                                {
                                    PO_X[0] = x; PO_Y[0] = 177;
                                    PO_X[1] = x; PO_Y[1] = 134;
                                    PO_X[2] = x; PO_Y[2] = 126;
                                    PO_X[3] = x; PO_Y[3] = 117;
                                    PO_X[4] = x; PO_Y[4] = 110;
                                }
                                else
                                {

                                    PO_X[0] = x; PO_Y[0] = 195;
                                    PO_X[1] = x; PO_Y[1] = 188;
                                    PO_X[2] = x; PO_Y[2] = 145;
                                    PO_X[3] = x; PO_Y[3] = 134;
                                    PO_X[4] = x; PO_Y[4] = 126;
                                    PO_X[5] = x; PO_Y[5] = 117;
                                    PO_X[6] = x; PO_Y[6] = 110;

                                }
                            }
                        }


                        #endregion

                        #region CP
                        if (letterhead == "CP")
                        {
                            if (UserType == "Head")
                            {
                                int x = 15;
                                if (FileType == "PO")
                                {
                                    PO_X[0] = x; PO_Y[0] = 177;
                                    PO_X[1] = x; PO_Y[1] = 134;
                                    PO_X[2] = x; PO_Y[2] = 126;
                                    PO_X[3] = x; PO_Y[3] = 117;
                                    PO_X[4] = x; PO_Y[4] = 110;
                                }
                                else
                                {

                                    PO_X[0] = x; PO_Y[0] = 195;
                                    PO_X[1] = x; PO_Y[1] = 188;
                                    PO_X[2] = x; PO_Y[2] = 145;
                                    PO_X[3] = x; PO_Y[3] = 134;
                                    PO_X[4] = x; PO_Y[4] = 126;
                                    PO_X[5] = x; PO_Y[5] = 117;
                                    PO_X[6] = x; PO_Y[6] = 110;

                                }
                            }
                            if (UserType == "Head1")
                            {
                                int x = 130;
                                if (FileType == "PO")
                                {
                                    PO_X[0] = x; PO_Y[0] = 177;
                                    PO_X[1] = x; PO_Y[1] = 134;
                                    PO_X[2] = x; PO_Y[2] = 126;
                                    PO_X[3] = x; PO_Y[3] = 117;
                                    PO_X[4] = x; PO_Y[4] = 110;
                                }
                                else
                                {

                                    PO_X[0] = x; PO_Y[0] = 195;
                                    PO_X[1] = x; PO_Y[1] = 188;
                                    PO_X[2] = x; PO_Y[2] = 145;
                                    PO_X[3] = x; PO_Y[3] = 134;
                                    PO_X[4] = x; PO_Y[4] = 126;
                                    PO_X[5] = x; PO_Y[5] = 117;
                                    PO_X[6] = x; PO_Y[6] = 110;

                                }
                            }
                            if (UserType == "Head2")
                            {
                                int x = 245;
                                if (FileType == "PO")
                                {
                                    PO_X[0] = x; PO_Y[0] = 177;
                                    PO_X[1] = x; PO_Y[1] = 134;
                                    PO_X[2] = x; PO_Y[2] = 126;
                                    PO_X[3] = x; PO_Y[3] = 117;
                                    PO_X[4] = x; PO_Y[4] = 110;
                                }
                                else
                                {

                                    PO_X[0] = x; PO_Y[0] = 195;
                                    PO_X[1] = x; PO_Y[1] = 188;
                                    PO_X[2] = x; PO_Y[2] = 145;
                                    PO_X[3] = x; PO_Y[3] = 134;
                                    PO_X[4] = x; PO_Y[4] = 126;
                                    PO_X[5] = x; PO_Y[5] = 117;
                                    PO_X[6] = x; PO_Y[6] = 110;

                                }
                            }
                            if (UserType == "Head3")
                            {
                                int x = 360;
                                if (FileType == "PO")
                                {
                                    PO_X[0] = x; PO_Y[0] = 177;
                                    PO_X[1] = x; PO_Y[1] = 134;
                                    PO_X[2] = x; PO_Y[2] = 126;
                                    PO_X[3] = x; PO_Y[3] = 117;
                                    PO_X[4] = x; PO_Y[4] = 110;
                                }
                                else
                                {

                                    PO_X[0] = x; PO_Y[0] = 195;
                                    PO_X[1] = x; PO_Y[1] = 188;
                                    PO_X[2] = x; PO_Y[2] = 145;
                                    PO_X[3] = x; PO_Y[3] = 134;
                                    PO_X[4] = x; PO_Y[4] = 126;
                                    PO_X[5] = x; PO_Y[5] = 117;
                                    PO_X[6] = x; PO_Y[6] = 110;

                                }
                            }

                            if (UserType == "Head4")
                            {
                                int x = 475;
                                if (FileType == "PO")
                                {
                                    PO_X[0] = x; PO_Y[0] = 177;
                                    PO_X[1] = x; PO_Y[1] = 134;
                                    PO_X[2] = x; PO_Y[2] = 126;
                                    PO_X[3] = x; PO_Y[3] = 117;
                                    PO_X[4] = x; PO_Y[4] = 110;
                                }
                                else
                                {

                                    PO_X[0] = x; PO_Y[0] = 195;
                                    PO_X[1] = x; PO_Y[1] = 188;
                                    PO_X[2] = x; PO_Y[2] = 145;
                                    PO_X[3] = x; PO_Y[3] = 134;
                                    PO_X[4] = x; PO_Y[4] = 126;
                                    PO_X[5] = x; PO_Y[5] = 117;
                                    PO_X[6] = x; PO_Y[6] = 110;

                                }
                            }
                        }


                        #endregion

                        #region Silvassa
                        if (letterhead == "Silvassa")
                        {
                            if (UserType == "Head")
                            {
                                int x = 15;
                                if (FileType == "PO")
                                {
                                    PO_X[0] = x; PO_Y[0] = 177;
                                    PO_X[1] = x; PO_Y[1] = 134;
                                    PO_X[2] = x; PO_Y[2] = 126;
                                    PO_X[3] = x; PO_Y[3] = 117;
                                    PO_X[4] = x; PO_Y[4] = 110;
                                }
                                else
                                {

                                    PO_X[0] = x; PO_Y[0] = 195;
                                    PO_X[1] = x; PO_Y[1] = 188;
                                    PO_X[2] = x; PO_Y[2] = 145;
                                    PO_X[3] = x; PO_Y[3] = 134;
                                    PO_X[4] = x; PO_Y[4] = 126;
                                    PO_X[5] = x; PO_Y[5] = 117;
                                    PO_X[6] = x; PO_Y[6] = 110;

                                }
                            }
                            if (UserType == "Head1")
                            {
                                int x = 130;
                                if (FileType == "PO")
                                {
                                    PO_X[0] = x; PO_Y[0] = 177;
                                    PO_X[1] = x; PO_Y[1] = 134;
                                    PO_X[2] = x; PO_Y[2] = 126;
                                    PO_X[3] = x; PO_Y[3] = 117;
                                    PO_X[4] = x; PO_Y[4] = 110;
                                }
                                else
                                {

                                    PO_X[0] = x; PO_Y[0] = 195;
                                    PO_X[1] = x; PO_Y[1] = 188;
                                    PO_X[2] = x; PO_Y[2] = 145;
                                    PO_X[3] = x; PO_Y[3] = 134;
                                    PO_X[4] = x; PO_Y[4] = 126;
                                    PO_X[5] = x; PO_Y[5] = 117;
                                    PO_X[6] = x; PO_Y[6] = 110;

                                }
                            }
                            if (UserType == "Head2")
                            {
                                int x = 245;
                                if (FileType == "PO")
                                {
                                    PO_X[0] = x; PO_Y[0] = 177;
                                    PO_X[1] = x; PO_Y[1] = 134;
                                    PO_X[2] = x; PO_Y[2] = 126;
                                    PO_X[3] = x; PO_Y[3] = 117;
                                    PO_X[4] = x; PO_Y[4] = 110;
                                }
                                else
                                {

                                    PO_X[0] = x; PO_Y[0] = 195;
                                    PO_X[1] = x; PO_Y[1] = 188;
                                    PO_X[2] = x; PO_Y[2] = 145;
                                    PO_X[3] = x; PO_Y[3] = 134;
                                    PO_X[4] = x; PO_Y[4] = 126;
                                    PO_X[5] = x; PO_Y[5] = 117;
                                    PO_X[6] = x; PO_Y[6] = 110;

                                }
                            }
                            if (UserType == "Head3")
                            {
                                int x = 360;
                                if (FileType == "PO")
                                {
                                    PO_X[0] = x; PO_Y[0] = 177;
                                    PO_X[1] = x; PO_Y[1] = 134;
                                    PO_X[2] = x; PO_Y[2] = 126;
                                    PO_X[3] = x; PO_Y[3] = 117;
                                    PO_X[4] = x; PO_Y[4] = 110;
                                }
                                else
                                {

                                    PO_X[0] = x; PO_Y[0] = 195;
                                    PO_X[1] = x; PO_Y[1] = 188;
                                    PO_X[2] = x; PO_Y[2] = 145;
                                    PO_X[3] = x; PO_Y[3] = 134;
                                    PO_X[4] = x; PO_Y[4] = 126;
                                    PO_X[5] = x; PO_Y[5] = 117;
                                    PO_X[6] = x; PO_Y[6] = 110;

                                }
                            }

                            if (UserType == "Head4")
                            {
                                int x = 475;
                                if (FileType == "PO")
                                {
                                    PO_X[0] = x; PO_Y[0] = 177;
                                    PO_X[1] = x; PO_Y[1] = 134;
                                    PO_X[2] = x; PO_Y[2] = 126;
                                    PO_X[3] = x; PO_Y[3] = 117;
                                    PO_X[4] = x; PO_Y[4] = 110;
                                }
                                else
                                {

                                    PO_X[0] = x; PO_Y[0] = 195;
                                    PO_X[1] = x; PO_Y[1] = 188;
                                    PO_X[2] = x; PO_Y[2] = 145;
                                    PO_X[3] = x; PO_Y[3] = 134;
                                    PO_X[4] = x; PO_Y[4] = 126;
                                    PO_X[5] = x; PO_Y[5] = 117;
                                    PO_X[6] = x; PO_Y[6] = 110;

                                }
                            }
                        }


                        #endregion



                        System.IO.MemoryStream outputStream = new System.IO.MemoryStream();
                        System.IO.MemoryStream imageStream = new System.IO.MemoryStream();

                        string textPDF = string.Empty;
                        using (System.IO.MemoryStream mStream = new System.IO.MemoryStream())
                        {
                            if (data != null)
                            {
                                String pathout = "";
                                data.Value.CopyTo(mStream);
                                byte[] array = mStream.ToArray();

                                Imagedata.Value.CopyTo(imageStream);
                                byte[] imgarray = imageStream.ToArray();



                                PdfReader reader = new PdfReader(array);

                                //select three pages from the original document
                                // reader.SelectPages("2");
                                int n = reader.NumberOfPages;

                                //create PdfStamper object to write to get the pages from reader 
                                PdfStamper stamper = new PdfStamper(reader, outputStream);
                                // PdfContentByte from stamper to add content to the pages over the original content
                            for (int i = 1; i <= n; i++)
                            {
                                   PdfContentByte pbover = stamper.GetOverContent(i);
                                //add content to the page using ColumnText

                                DateTime dateTime = DateTime.Now;

                                var blackListTextFont = FontFactory.GetFont("Arial", 6, Color.BLACK);
                                    // add image


                                    if (FileType == "PO")
                                    {
                                        ColumnText.ShowTextAligned(pbover, Element.ALIGN_LEFT, new Phrase(new Chunk(PurchaseText, blackListTextFont)), PO_X[0], PO_Y[0], 0);
                                        iTextSharp.text.Image sigimage = iTextSharp.text.Image.GetInstance(imgarray);
                                        sigimage.SetAbsolutePosition(PO_X[1], PO_Y[1]);
                                        pbover.AddImage(sigimage);
                                        ColumnText.ShowTextAligned(pbover, Element.ALIGN_LEFT, new Phrase(new Chunk(User, blackListTextFont)), PO_X[2], PO_Y[2], 0);
                                        ColumnText.ShowTextAligned(pbover, Element.ALIGN_LEFT, new Phrase(new Chunk(Convert.ToString(dateTime), blackListTextFont)), PO_X[3], PO_Y[3], 0);
                                        if (UserType == "Head")
                                        {
                                            ColumnText.ShowTextAligned(pbover, Element.ALIGN_LEFT, new Phrase(new Chunk(footer, blackListTextFont)), PO_X[4], PO_Y[4], 0);
                                        }
                                    }
                                    else
                                    {
                                        ColumnText.ShowTextAligned(pbover, Element.ALIGN_LEFT, new Phrase(new Chunk(DispayPO, blackListTextFont)), PO_X[0], PO_Y[0], 0);
                                        ColumnText.ShowTextAligned(pbover, Element.ALIGN_LEFT, new Phrase(new Chunk(ActionName, blackListTextFont)), PO_X[1], PO_Y[1], 0);
                                        iTextSharp.text.Image sigimage = iTextSharp.text.Image.GetInstance(imgarray);
                                        sigimage.SetAbsolutePosition(PO_X[2], PO_Y[2]);
                                        pbover.AddImage(sigimage);
                                        ColumnText.ShowTextAligned(pbover, Element.ALIGN_LEFT, new Phrase(new Chunk(User, blackListTextFont)), PO_X[3], PO_Y[3], 0);
                                        ColumnText.ShowTextAligned(pbover, Element.ALIGN_LEFT, new Phrase(new Chunk(Comments, blackListTextFont)), PO_X[4], PO_Y[4], 0);
                                        ColumnText.ShowTextAligned(pbover, Element.ALIGN_LEFT, new Phrase(new Chunk(Convert.ToString(dateTime), blackListTextFont)), PO_X[5], PO_Y[5], 0);
                                        if (UserType == "Head")
                                        {
                                            ColumnText.ShowTextAligned(pbover, Element.ALIGN_LEFT, new Phrase(new Chunk(footer, blackListTextFont)), PO_X[6], PO_Y[6], 0);
                                        }
                                    }
                                
 

                                 
                                   
                                    PdfContentByte pbunder = stamper.GetUnderContent(i);
                                   
                            }
                                stamper.Close();
                                //close the stamper



                                // Update PDF Code
                                #region Update PDF Code

                                string siteURL = "https://aparindltd.sharepoint.com";
                                string documentListName = "PurchaseDocuments";
                                string documentListURL = "https://aparindltd.sharepoint.com/PurchaseOrder/PurchaseDocuments/";
                                //string documentName = "11111_Airnet.pdf";


                                Web web = context.Web;
                                Microsoft.SharePoint.Client.List list = web.Lists.GetByTitle("PurchaseDocuments");

                                var fileCreationInformation = new FileCreationInformation();
                                byte[] array1 = outputStream.ToArray();
                                fileCreationInformation.Content = array1;
                                fileCreationInformation.Overwrite = true;
                                //fileCreationInformation.Url = documentListURL + documentName;
                                fileCreationInformation.Url = path;
                                Microsoft.SharePoint.Client.File uploadFile = list.RootFolder.Files.Add(fileCreationInformation);
                                //   uploadFile.ListItemAllFields["Action"] = "Favourites";
                                uploadFile.ListItemAllFields.Update();
                                context.ExecuteQuery();

                                #endregion



                            }


                        }
                    }
                    return ActionName;
                }
            }


        }


    }
}
