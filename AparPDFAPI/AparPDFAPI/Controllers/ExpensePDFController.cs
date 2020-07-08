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
    public class ExpensePDFController : ApiController
    {
        [HttpGet]
        [Route("api/ExpensePDFImage/GenerateExpensePDF/{ActionName}/{LoginName}/{EmailID}/{Comments}/{PONumber}/{PODisplayNo}/{UserType}")]
        public string GetPDFMyAction(string ActionName, string LoginName, string EmailID, string Comments, string PONumber, string PODisplayNo, string UserType)
        {
            string Action = ActionName;
            //LoginName = "SP User2";
            // PONumber = "62";
            //PONumber = "32";
            string User = LoginName;
            string POnum = PONumber;

            string Pono = PODisplayNo;
            string res = Pono.Replace("@", "/");

            //string DispayPO = PODisplayNo;
            string DispayPO = res;
            // string UserType = "Head";
            string footer = "DOCUMENT ARE SIGNED DIGITALLY, HENCE NO PHYSICAL SIGNATURE REQUIRED.";

            string login = "sp.admin@apar.com"; //give your username here  


            using (var contextimage = new ClientContext("https://aparindltd.sharepoint.com"))
            {


                using (var context = new ClientContext("https://aparindltd.sharepoint.com/ExpenseVoucher"))
                {




                    #region Get Data From List
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

                    var subsitelistdata = context.Web.Lists.GetByTitle("ExpenseDocuments");


                    // var itemss = subsitelistdata.GetItems(CamlQuery.CreateAllItemsQuery());
                    CamlQuery query1 = new CamlQuery();

                    //query1.ViewXml = "<View><Query><Where><Eq><FieldRef Name='POReferenceNumber' LookupId='FALSE'/><Value Type='Lookup'>"+POnum+"</Value></Eq></Where></Query></view>";

                    query1.ViewXml = "<View><Query><Where><Eq><FieldRef Name='ExpVoucherNo' LookupId='TRUE'/><Value Type='Lookup'>" + POnum + "</Value></Eq></Where></Query></view>";

                    ListItemCollection listitem1 = subsitelistdata.GetItems(query1);


                    context.Load(listitem1);
                    context.ExecuteQuery();
                    var docnm = "";
                    var path = "";
                    foreach (var oListItem1 in listitem1)
                    {
                        //var FileType = oListItem1["FileType"].ToString();
                        docnm = oListItem1["FileLeafRef"].ToString();

                        var docId = oListItem1["ID"].ToString();
                        var Rowno = oListItem1["RowNumber"].ToString();
                        string[] Type = docnm.Split('.');

                        var Filenm = Type[0].ToString();
                        var Exttype = Type[1].ToString();

                        if (Exttype != "pdf")
                        {

                            using (MemoryStream ms = new MemoryStream())
                            {
                                // Document document = new Document(PageSize.A4, 25, 25, 30, 30);
                                Document document = new Document(PageSize.A4.Rotate());

                                PdfWriter writer = PdfWriter.GetInstance(document, ms);

                                document.Open();

                                //document.Add(new Paragraph("Hello World"));
                                var docimg = "/ExpenseVoucher/ExpenseDocuments/" + docnm + "";
                                var docimg1 = "/ExpenseVoucher/ExpenseDocuments/" + Filenm + ".pdf";
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
                                sigimage1.ScaleAbsolute(400f,370f);
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
                                string documentListName = "ExpenseDocuments";
                                string documentListURL = "https://aparindltd.sharepoint.com/ExpenseVoucher/ExpenseDocuments/";
                                //string documentName = "11111_Airnet.pdf";


                                Web web = context.Web;
                                Microsoft.SharePoint.Client.List list = web.Lists.GetByTitle("ExpenseDocuments");

                                var fileCreationInformation = new FileCreationInformation();
                                byte[] array1 = ms.ToArray();
                                fileCreationInformation.Content = array1;
                                fileCreationInformation.Overwrite = true;
                                //fileCreationInformation.Url = documentListURL + documentName;
                                fileCreationInformation.Url = docimg1;
                                Microsoft.SharePoint.Client.File uploadFile = list.RootFolder.Files.Add(fileCreationInformation);
                                uploadFile.ListItemAllFields["ExpVoucherNo"] = POnum;
                                uploadFile.ListItemAllFields["RowNumber"] = Rowno;
                               // uploadFile.ListItemAllFields["ExpVoucherNo"] = 27;
                                uploadFile.ListItemAllFields.Update();
                                context.ExecuteQuery();

                                path = "/ExpenseVoucher/ExpenseDocuments/" + Filenm + ".pdf";






                            }
                        }

                        else
                        {


                            path = "/ExpenseVoucher/ExpenseDocuments/" + docnm + "";

                           

                        }


                        

                        //file = context.Web.GetFileByServerRelativeUrl(path);



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


                        //TcpClient socketConnection;
                        //socketConnection = new TcpClient("localhost", 49335);  	
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
                                PdfContentByte pbover = stamper.GetOverContent(n);
                                //add content to the page using ColumnText

                                DateTime dateTime = DateTime.Now;


                                // add image

                                if (UserType == "FunctionalHead")
                                {


                                    ColumnText.ShowTextAligned(pbover, Element.ALIGN_LEFT, new Phrase(DispayPO), 20,180,0);
                                    ColumnText.ShowTextAligned(pbover, Element.ALIGN_LEFT, new Phrase(ActionName), 20,165,0);
                                    iTextSharp.text.Image sigimage = iTextSharp.text.Image.GetInstance(imgarray);
                                    sigimage.SetAbsolutePosition(20,115);
                                    pbover.AddImage(sigimage);
                                    ColumnText.ShowTextAligned(pbover, Element.ALIGN_LEFT, new Phrase(User), 20,90, 0);
                                    ColumnText.ShowTextAligned(pbover, Element.ALIGN_LEFT, new Phrase(Comments), 20,73, 0);
                                    ColumnText.ShowTextAligned(pbover, Element.ALIGN_LEFT, new Phrase(Convert.ToString(dateTime)), 20,50, 0);
                                    ColumnText.ShowTextAligned(pbover, Element.ALIGN_LEFT, new Phrase(footer), 20, 35, 0);


                                }

                                if (UserType == "InternalAuditor")
                                {

                                    //DispayPO
                                    ColumnText.ShowTextAligned(pbover, Element.ALIGN_LEFT, new Phrase(DispayPO), 200, 180, 0);                                    
                                    ColumnText.ShowTextAligned(pbover, Element.ALIGN_LEFT, new Phrase(ActionName), 200, 165, 0);
                                    iTextSharp.text.Image sigimage = iTextSharp.text.Image.GetInstance(imgarray);
                                    sigimage.SetAbsolutePosition(200, 115);
                                    pbover.AddImage(sigimage);
                                    ColumnText.ShowTextAligned(pbover, Element.ALIGN_LEFT, new Phrase(User), 200, 90, 0);
                                    ColumnText.ShowTextAligned(pbover, Element.ALIGN_LEFT, new Phrase(Comments), 200, 73, 0);
                                    ColumnText.ShowTextAligned(pbover, Element.ALIGN_LEFT, new Phrase(Convert.ToString(dateTime)), 200, 50, 0);

                                }

                                if (UserType == "Accountant")
                                {

                                    //DispayPO
                                    ColumnText.ShowTextAligned(pbover, Element.ALIGN_LEFT, new Phrase(DispayPO), 400, 180, 0);                                    
                                    ColumnText.ShowTextAligned(pbover, Element.ALIGN_LEFT, new Phrase(ActionName), 400, 165, 0);
                                    iTextSharp.text.Image sigimage = iTextSharp.text.Image.GetInstance(imgarray);
                                    sigimage.SetAbsolutePosition(400, 115);
                                    pbover.AddImage(sigimage);
                                    ColumnText.ShowTextAligned(pbover, Element.ALIGN_LEFT, new Phrase(User), 400, 90, 0);
                                    ColumnText.ShowTextAligned(pbover, Element.ALIGN_LEFT, new Phrase(Comments), 400, 73, 0);
                                    ColumnText.ShowTextAligned(pbover, Element.ALIGN_LEFT, new Phrase(Convert.ToString(dateTime)), 400, 50, 0);

                                }




                                // PdfContentByte from stamper to add content to the pages under the original content
                                PdfContentByte pbunder = stamper.GetUnderContent(n);
                                //close the stamper
                                stamper.Close();


                                // Update PDF Code
                                #region Update PDF Code

                                string siteURL = "https://aparindltd.sharepoint.com";
                                string documentListName = "ExpenseDocuments";
                                string documentListURL = "https://aparindltd.sharepoint.com/ExpenseVoucher/ExpenseDocuments/";
                                //string documentName = "11111_Airnet.pdf";


                                Web web = context.Web;
                                Microsoft.SharePoint.Client.List list = web.Lists.GetByTitle("ExpenseDocuments");

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




        [HttpGet]
        [Route("api/PDF/GetUser/{UserName}/{Password}")]
        public string MyAction(string UserName, string Password)
        {
            return Convert.ToString(UserName);
            // return UserName;

        }



    }
}
