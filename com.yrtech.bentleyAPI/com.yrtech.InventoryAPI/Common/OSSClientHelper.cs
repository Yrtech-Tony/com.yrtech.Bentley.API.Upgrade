﻿using Aliyun.OSS;
using Aliyun.OSS.Util;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Web;

namespace com.yrtech.InventoryAPI.Common
{
    public class OSSClientHelper
    {
        protected const string accessid = "LTAI5tQE74NUsWQhxTdS78S7";
        protected const string accessKey = "QeEoL0ERxEOUaReCxZhMKY5oFjLFuV";
        protected const string endpoin = "http://oss-cn-beijing.aliyuncs.com";
        protected const string bucket = "yrsurvey";
        public const string OSS_BASE_URL = "http://yrsurvey.oss-cn-beijing.aliyuncs.com/";
        public static bool UploadOSSFile(string key, Stream fileStream,long length)
        {
            try
            {
                string md5 = OssUtils.ComputeContentMd5(fileStream, length);
                //创建上传Object的Metadata 
                ObjectMetadata objectMetadata = new ObjectMetadata() {
                    ContentMd5 = md5
                };

                OssClient ossClient = new OssClient(endpoin, accessid, accessKey);
                var result = ossClient.PutObject(bucket, key, fileStream, objectMetadata);
                return true;
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());
                return false;
            }
        }
        public static void GetObject(string key, string fileToDownload)
        {
            OssClient ossClient = new OssClient(endpoin, accessid, accessKey);
            var o = ossClient.GetObject(bucket, key);
            using (var requestStream = o.Content)
            {
                byte[] buf = new byte[1024];
                var fs = File.Open(fileToDownload, FileMode.OpenOrCreate);
                var len = 0;
                while ((len = requestStream.Read(buf, 0, 1024)) != 0)
                {
                    fs.Write(buf, 0, len);
                }
                fs.Close();
            }
        }
    }
}