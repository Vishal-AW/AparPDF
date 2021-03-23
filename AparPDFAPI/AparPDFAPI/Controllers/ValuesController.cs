using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Drawing;
using QRCoder;
using System.Web.Http;
using System.IO;

namespace AparPDFAPI.Controllers
{
    public class ValuesController : ApiController
    {

        public class ASD
        {
            public string FirstName { get; set; }
            public string LastName { get; set; }
        }


        // GET api/values
        public byte[] Get(string BarCodeString)
        {
            Image img = GetImage(BarCodeString);
            byte[] image;
            using (MemoryStream ms = new MemoryStream())
            {
                img.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
                image = ms.ToArray();
            }

            return image;
        }

        public Image GetImage(string str)
        {
            QRCodeGenerator qrGenerator = new QRCodeGenerator();
            QRCodeData qrCodeData = qrGenerator.CreateQrCode(str, QRCodeGenerator.ECCLevel.Q);
            QRCoder.QRCode qrCode = new QRCoder.QRCode(qrCodeData);
            Bitmap qrCodeImage = qrCode.GetGraphic(20);
            Bitmap resized = new Bitmap(qrCodeImage, new Size(300, 300));
            return resized;
        }





        // GET api/values/5
        public string Get(int id)
        {
            return "value";
        }

        // POST api/values
        public ASD Post([FromBody] ASD aSD)
        {
            ASD sD = new ASD();
            sD.FirstName = aSD.FirstName + "Change NAme";
            sD.LastName = aSD.LastName + "Change";
            return sD;
        }

        // PUT api/values/5
        public void Put(int id, [FromBody]string value)
        {
        }

        // DELETE api/values/5
        public void Delete(int id)
        {
        }
    }
}
