using System;
using System.IO;
using System.IO.Compression;
using System.Text;

namespace TestImageOnSqlserver
{
    /// <summary>  
    /// 字符压缩类  
    /// </summary>  
    public static class CompressHelper
    {
        /// <summary>
        /// 压缩String
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        public static string CompressStringByString(string str)
        {
            //byte[] Bytes = CompressString(str);
            //Bytes = Encoding.UTF8.GetBytes(Convert.ToBase64String(Bytes));
            //Bytes = Convert.FromBase64String(Convert.ToBase64String(Bytes));
            //return Encoding.UTF8.GetString(Bytes, 0, Bytes.Length);

            byte[] Bytes = CompressString(str);
            return Convert.ToBase64String(Bytes);
        }

        /// <summary>
        /// 解压String
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        public static string DeCompressStringByString(string str)
        {
            byte[] Bytes = Convert.FromBase64String(str);
            return DeCompressStringByBytes(Bytes);
        }

        /// <summary>  
        /// 压缩字符串
        /// </summary>  
        /// <param name="str"></param>  
        /// <returns></returns>  
        public static byte[] CompressString(string str)
        {
            return CompressBytes(Encoding.UTF8.GetBytes(str));
        }

        /// <summary>  
        /// 压缩二进制  
        /// </summary>  
        /// <param name="str"></param>  
        /// <returns></returns>  
        public static byte[] CompressBytes(byte[] str)
        {
            var ms = new MemoryStream(str) { Position = 0 };
            var outms = new MemoryStream();
            using (var deflateStream = new DeflateStream(outms, CompressionMode.Compress, true))
            {
                var buf = new byte[1024];
                int len;
                while ((len = ms.Read(buf, 0, buf.Length)) > 0)
                    deflateStream.Write(buf, 0, len);
            }
            return outms.ToArray();
        }
        /// <summary>  
        /// 解压字符串  
        /// </summary>  
        /// <param name="str"></param>  
        /// <returns></returns>  
        public static string DeCompressStringByBytes(byte[] str)
        {
            byte[] Bytes = DecompressBytes(str);
            return Encoding.UTF8.GetString(Bytes, 0, Bytes.Length);
        }
        /// <summary>  
        /// 解压二进制  
        /// </summary>  
        /// <param name="str"></param>  
        /// <returns></returns>  
        public static byte[] DecompressBytes(byte[] str)
        {
            var ms = new MemoryStream(str) { Position = 0 };
            var outms = new MemoryStream();
            using (var deflateStream = new DeflateStream(ms, CompressionMode.Decompress, true))
            {
                var buf = new byte[1024];
                int len;
                while ((len = deflateStream.Read(buf, 0, buf.Length)) > 0)
                    outms.Write(buf, 0, len);
            }
            return outms.ToArray();
        }
    }
}
