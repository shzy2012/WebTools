using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net;

namespace Project.Utility {
    public class DataTypeConverter {
        /// <summary>
        /// byte[]转16进制格式string
        /// </summary>
        /// <param name="bytes"></param>
        /// <returns></returns>
        public static string ByteToHexString(byte[] bytes) {
            string str = string.Empty;
            if (bytes != null) {
                StringBuilder sb = new StringBuilder();
                for (int i = 0; i < bytes.Length; i++) {
                    sb.Append(bytes[i].ToString("X2"));
                }
                str = sb.ToString();
            }
            return str;
        }

        /// <summary>
        /// int转byte[]
        /// </summary>
        /// <param name="i"></param>
        /// <returns></returns>
        public static byte[] IntToByte(int i) {
            byte[] bytes = new byte[4];
            bytes[0] = (byte)(i);
            bytes[1] = (byte)(i >> 8);
            bytes[2] = (byte)(i >> 16);
            bytes[3] = (byte)(i >> 24);
            return bytes;
        }

        /// <summary>
        /// byte[]转int
        /// </summary>
        /// <param name="bt"></param>
        /// <returns></returns>
        public static int ByteToInt(byte[] bytes) {
            return (int)(bytes[0] | bytes[1] << 8 | bytes[2] << 16 | bytes[3] << 24);
        }

        /// <summary>
        /// int16转byte[]
        /// </summary>
        /// <param name="i"></param>
        /// <returns></returns>
        public static byte[] Int16ToByte(Int16 i) {
            byte[] bytes = new byte[2];
            bytes[0] = (byte)(i);
            bytes[1] = (byte)(i >> 8);
            return bytes;
        }

        /// <summary>
        /// byte[]转int
        /// </summary>
        /// <param name="bytes"></param>
        /// <returns></returns>
        public static Int16 ByteToInt16(byte[] bytes) {
            return (short)(bytes[0] | bytes[1] << 8);
        }

        /// <summary>
        /// IPV4转化为Bytes
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        public static byte[] IPV4ToByte(string address) {
            return IPAddress.Parse(address).GetAddressBytes();
        }

        /// <summary>
        /// Bytes转化为IPV4
        /// </summary>
        /// <param name="bytes"></param>
        /// <returns></returns>
        public static string ByteToIPV4(byte[] bytes) {
            return new IPAddress(bytes).ToString();
        }

        /// <summary>
        /// UnixTimeStamp
        /// </summary>
        /// <param name="dt"></param>
        /// <returns></returns>
        public static double DateTimeToUTC(DateTime dt) {
            return (dt - TimeZone.CurrentTimeZone.ToLocalTime(new DateTime(1970, 1, 1))).TotalSeconds;
        }

        /// <summary>
        /// UnixStampToDateTime
        /// </summary>
        /// <param name="d"></param>
        /// <returns></returns>
        public static DateTime UTCToDateTime(double d) {
            return TimeZone.CurrentTimeZone.ToLocalTime(new DateTime(1970, 1, 1).AddSeconds(d));
        }

        /// <summary>
        /// 将string类型转换成DateTime时间格式
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="format"></param>
        /// <returns></returns>
        public static DateTime StringToDateTime(string dt, string format) {
            return DateTime.ParseExact(dt, format, System.Globalization.CultureInfo.CurrentCulture);
        }


        /// <summary>
        /// UnixTimeStampLong
        /// </summary>
        /// <param name="dt"></param>
        /// <returns></returns>
        public static double DateTimeToUTCLong(DateTime dt) {
            return (dt - TimeZone.CurrentTimeZone.ToLocalTime(new DateTime(1970, 1, 1))).TotalMilliseconds;
        }

        /// <summary>
        /// LongUnixStampToDateTime
        /// </summary>
        /// <param name="d"></param>
        /// <returns></returns>
        public static DateTime LongUTCToDateTime(double d) {
            return TimeZone.CurrentTimeZone.ToLocalTime(new DateTime(1970, 1, 1).AddMilliseconds(d));
        }

        /// <summary>
        /// 合并俩个数组
        /// </summary>
        /// <param name="array1"></param>
        /// <param name="array2"></param>
        public static List<string> MergeArray(string[] array1, string[] array2) {
            List<string> list = new List<string>();
            for (int i = 0; i < array1.Length; i++) {
                list.Add(array1[i]);
            }

            for (int j = 0; j < array2.Length; j++) {
                if (!ArrayContains(array2[j], array1)) {
                    list.Add(array2[j]);
                }
            }

            return list;
        }

        ///<summary>
        ///包含在数组中
        /// </summary>
        /// <param name="s"></param>
        /// <param name="array2"></param>
        public static bool ArrayContains(string s, string[] array) {
            for (int i = 0; i < array.Length; i++) {
                if (array[i] == s) {
                    return true;
                }
            }
            return false;
        }

        ///<summary>
        ///包含在数组中
        /// </summary>
        /// <param name="s"></param>
        /// <param name="array2"></param>
        public static bool ArrayContains(int s, int[] array) {
            for (int i = 0; i < array.Length; i++) {
                if (array[i] == s) {
                    return true;
                }
            }
            return false;
        }

        /// <summary>
        /// 连接数组
        /// </summary>
        /// <param name="array"></param>
        /// <param name="c"></param>
        /// <returns></returns>
        public static string Concat(string[] array, char c) {
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < array.Length; i++) {
                sb.Append(array[i]);
                if (i < array.Length - 1) {
                    sb.Append(c);
                }
            }
            return sb.ToString();
        }

        /// <summary>
        /// Convert String to Int
        /// </summary>
        /// <param name="s"></param>
        /// <returns></returns>
        public static int? StringToInt(string s) {

            int value;
            if (int.TryParse(s, out value)) {
                return value;
            }

            return null;
        }
    }
}
