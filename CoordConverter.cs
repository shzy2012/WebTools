using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Project.Utility {
    /// <summary>
    /// 各地图API坐标系统比较与转换; 
    /// WGS84坐标系：即地球坐标系，国际上通用的坐标系。设备一般包含GPS芯片或者北斗芯片获取的经纬度为WGS84地理坐标系
    /// 谷歌地图采用的是WGS84地理坐标系（中国范围除外）
    /// GCJ02坐标系：即火星坐标系，是由中国国家测绘局制订的地理信息系统的坐标系统。由WGS84坐标系经加密后的坐标系
    /// 谷歌中国地图和搜搜中国地图采用的是GCJ02地理坐标系; BD09坐标系：即百度坐标系，GCJ02坐标系经加密后的坐标系
    /// 搜狗坐标系、图吧坐标系等，也是在GCJ02基础上加密而成的 
    /// </summary>
    public class CoordConverter {
        private static double pi = 3.1415926535897932384626;
        /// <summary>
        /// 椭球的偏心率
        /// </summary>
        private static double ee = 0.00669342162296594323;
        /// <summary>
        /// 卫星椭球坐标投影到平面地图坐标系的投影因子
        /// </summary>
        private static double a = 6378245.0;

        /// <summary>
        /// wgs84->gcj02
        /// </summary>
        /// <param name="lat"></param>
        /// <param name="lng"></param>
        /// <returns></returns>
        public static DotArg Wgs84ToGcj02(double lat, double lng) {
            if (Abroad(lat, lng)) {
                return new DotArg(lat, lng);
            }
            DotArg d = Delta(lat, lng);
            d.Lat += lat;
            d.Lng += lng;
            return d;
        }

        /// <summary>
        /// gcj02->wgs84
        /// </summary>
        /// <param name="lat"></param>
        /// <param name="lng"></param>
        /// <returns></returns>
        public static DotArg Gcj02ToWgs84(double lat, double lng) {
            if (Abroad(lat, lng)) {
                return new DotArg(lat, lng);
            }
            DotArg d = Delta(lat, lng);
            d.Lat = lat - d.Lat;
            d.Lng = lng - d.Lng;
            return d;
        }

        /// <summary>
        /// gcj02->bd09
        /// </summary>
        /// <param name="lat"></param>
        /// <param name="lng"></param>
        public static DotArg Gcj02ToBd09(double lat, double lng) {
            double x = lng, y = lat;
            double z = Math.Sqrt(x * x + y * y) + 0.00002 * Math.Sin(y * pi);
            double theta = Math.Atan2(y, x) + 0.000003 * Math.Cos(x * pi);
            double bdLng = z * Math.Cos(theta) + 0.0065;
            double bdLat = z * Math.Sin(theta) + 0.006;
            return new DotArg(bdLat, bdLng);
        }

        /// <summary>
        /// bd09->gcj02
        /// </summary>
        /// <param name="lat"></param>
        /// <param name="lng"></param>
        /// <returns></returns>
        public static DotArg Bd09ToGcj02(double lat, double lng) {
            double x = lng - 0.0065, y = lat - 0.006;
            double z = Math.Sqrt(x * x + y * y) - 0.00002 * Math.Sin(y * pi);
            double theta = Math.Atan2(y, x) - 0.000003 * Math.Cos(x * pi);
            double gcjLng = z * Math.Cos(theta);
            double gcjLat = z * Math.Sin(theta);
            return new DotArg(gcjLat, gcjLng);
        }

        /// <summary>
        /// 修正坐标
        /// </summary>
        /// <param name="lat"></param>
        /// <param name="lng"></param>
        /// <returns></returns>
        private static DotArg Delta(double lat, double lng) {
            double dLat = ReviseFormLat(lng - 105.0, lat - 35.0);
            double dLng = ReviseFormLng(lng - 105.0, lat - 35.0);
            double radLat = lat / 180.0 * pi;
            double magic = Math.Sin(radLat);
            magic = 1 - ee * magic * magic;
            double sqrtMagic = Math.Sqrt(magic);
            dLat = (dLat * 180.0) / ((a * (1 - ee)) / (magic * sqrtMagic) * pi);
            dLng = (dLng * 180.0) / (a / sqrtMagic * Math.Cos(radLat) * pi);

            return new DotArg(dLat, dLng);
        }

        /// <summary>
        /// 国外
        /// </summary>
        /// <param name="lat"></param>
        /// <param name="lng"></param>
        /// <returns></returns>
        private static bool Abroad(double lat, double lng) {
            if (lng < 72.004 || lng > 137.8347)
                return true;
            if (lat < 0.8293 || lat > 55.8271)
                return true;
            return false;
        }

        /// <summary>
        /// 修正lat
        /// </summary>
        /// <param name="x"></param>
        /// <param name="y"></param>
        /// <returns></returns>
        private static double ReviseFormLat(double x, double y) {
            double ret = -100.0 + 2.0 * x + 3.0 * y + 0.2 * y * y + 0.1 * x * y + 0.2 * Math.Sqrt(Math.Abs(x));
            ret += (20.0 * Math.Sin(6.0 * x * pi) + 20.0 * Math.Sin(2.0 * x * pi)) * 2.0 / 3.0;
            ret += (20.0 * Math.Sin(y * pi) + 40.0 * Math.Sin(y / 3.0 * pi)) * 2.0 / 3.0;
            ret += (160.0 * Math.Sin(y / 12.0 * pi) + 320 * Math.Sin(y * pi / 30.0)) * 2.0 / 3.0;
            return ret;
        }

        /// <summary>
        /// 修正lng
        /// </summary>
        /// <param name="x"></param>
        /// <param name="y"></param>
        /// <returns></returns>
        private static double ReviseFormLng(double x, double y) {
            double ret = 300.0 + x + 2.0 * y + 0.1 * x * x + 0.1 * x * y + 0.1 * Math.Sqrt(Math.Abs(x));
            ret += (20.0 * Math.Sin(6.0 * x * pi) + 20.0 * Math.Sin(2.0 * x * pi)) * 2.0 / 3.0;
            ret += (20.0 * Math.Sin(x * pi) + 40.0 * Math.Sin(x / 3.0 * pi)) * 2.0 / 3.0;
            ret += (150.0 * Math.Sin(x / 12.0 * pi) + 300.0 * Math.Sin(x / 30.0 * pi)) * 2.0 / 3.0;
            return ret;
        }
    }
    
    
public class DotArg {
        /// <summary>
        /// 构造函数
        /// </summary>
        public DotArg() { }
        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="lat"></param>
        /// <param name="lng"></param>
        public DotArg(double lat, double lng) {
            Lat = lat;
            Lng = lng;
        }

        /// <summary>
        /// 经度
        /// </summary>
        public double Lat { get; set; }

        /// <summary>
        /// 维度
        /// </summary>
        public double Lng { get; set; }

        /// <summary>
        /// 精度
        /// </summary>
        public double Accuracy { get; set; }

        /// <summary>
        /// 定位类型
        /// </summary>
        public LocaTypeEnum LocaType { get; set; }

        private string desc = "";
        /// <summary>
        /// 具体地址
        /// </summary>
        public string Desc {
            get {
                return desc;
            }
            set {
                desc = value;
            }
        }

        /// <summary>
        /// tostring
        /// </summary>
        /// <returns></returns>
        public override string ToString() {
            return string.Concat(Lat, ',', Lng);
        }
    }


    /// <summary>
    /// 定位类型
    /// </summary>
    public enum LocaTypeEnum {
        /// <summary>
        /// gps
        /// </summary>
        GPS = 0,
        /// <summary>
        /// lbs
        /// </summary>
        LBS = 1,
        /// <summary>
        /// wifi
        /// </summary>
        WIFI = 2,
        /// <summary>
        /// last
        /// </summary>
        LAST = 3,
        /// <summary>
        /// most
        /// </summary>
        MOST = 4,
        /// <summary>
        /// 无效点
        /// </summary>
        Invalid = 5
    }
}



