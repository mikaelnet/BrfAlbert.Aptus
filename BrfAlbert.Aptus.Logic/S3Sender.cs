using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Amazon.S3;
using Amazon.S3.Model;

namespace BrfAlbert.Aptus.Logic
{
    public class S3Sender
    {
        public string BucketName { get; set; }


        public void SendFile(string key, Stream stream)
        {
            using (var client = new AmazonS3Client())
            {
                var request = new PutObjectRequest();
                request.BucketName = BucketName;
                request.Key = key;
                request.InputStream = stream;

                client.PutObject(request);
            }
        }
    }
}
