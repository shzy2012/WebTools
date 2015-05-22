namespace VAU.Web.CommonCode
{
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Web;
    using System.Xml;

    /// <summary>
    /// Data Source from the /Config Folder
    /// </summary>
    public class CommonXmlSource
    {
        #region Config String

        private static readonly string DeliveryPath =
            HttpContext.Current.Server.MapPath(System.Web.HttpContext.Current.Request.ApplicationPath)
            + @"\Config\Delivery.xml";

        private static readonly string InspectorListPath =
            HttpContext.Current.Server.MapPath(System.Web.HttpContext.Current.Request.ApplicationPath)
            + @"\Config\InspectorList.xml";

        private static readonly string OrderTypeEmailPath =
            HttpContext.Current.Server.MapPath(System.Web.HttpContext.Current.Request.ApplicationPath)
            + @"\Config\OrderTypeEmail.xml";

        private static readonly string BOMStatueTypeEmailPath =
            HttpContext.Current.Server.MapPath(System.Web.HttpContext.Current.Request.ApplicationPath)
            + @"\Config\BOMStatueTypeEmail.xml";

        private static readonly string WarehouseTypeCollectionPath =
            HttpContext.Current.Server.MapPath(System.Web.HttpContext.Current.Request.ApplicationPath)
            + @"\Config\WarehouseTypeCollection.xml";

        private static readonly string BookingTypeEmailPath =
            HttpContext.Current.Server.MapPath(System.Web.HttpContext.Current.Request.ApplicationPath)
            + @"\Config\BookingTypeEmail.xml";

        private static readonly string ForwarderTelexReleaseEmailPath =
            HttpContext.Current.Server.MapPath(System.Web.HttpContext.Current.Request.ApplicationPath)
            + @"\Config\ForwarderTelexReleaseEmail.xml";

        private static readonly string SampleReviewMessageEmailPath =
            HttpContext.Current.Server.MapPath(System.Web.HttpContext.Current.Request.ApplicationPath)
            + @"\Config\SampleReviewMessageEmail.xml";

        private static readonly string PoDepositEmailPath = 
            HttpContext.Current.Server.MapPath(System.Web.HttpContext.Current.Request.ApplicationPath)
            + @"\Config\PoDepositEmail.xml";

        private static XmlDocument doc = new XmlDocument();
        #endregion

        public CommonXmlSource()
        {
        }

        /// <summary>
        /// Get RequestedBY data from XML, the path : /Config/Delivery.xml
        /// </summary>
        /// <returns></returns>
        public static List<string> GetDelivery()
        {
            List<string> list = new List<string>();
            list.Add("[Select]");
            if (System.Web.HttpContext.Current.Request.ApplicationPath != null)
            {
                if (!File.Exists(DeliveryPath))
                {
                    return list;
                }

                doc.Load(DeliveryPath);
                XmlNode root = doc.SelectSingleNode("Groups");
                if (root != null)
                {
                    for (int i = 0; i < root.ChildNodes.Count; i++)
                    {
                        var child = root.ChildNodes[i];
                        list.AddRange(
                            from XmlElement item in child.ChildNodes select child.Name + " - " + item.InnerText);
                    }
                }
            }

            return list;
        }

        public static IList<string> GetInspector()
        {
            List<string> list = new List<string>();
            if (System.Web.HttpContext.Current.Request.ApplicationPath != null)
            {
                doc.Load(InspectorListPath);
                XmlNode root = doc.SelectSingleNode("Root");
                if (root != null && root.SelectNodes("Inspector") != null)
                {
                    var xmlNodeList = root.SelectNodes("Inspector");
                    if (xmlNodeList != null)
                    {
                        list.AddRange(from XmlElement item in xmlNodeList select item.InnerText);
                    }
                }
            }

            return list;
        }

        public static List<string> GetEmailAddress(string type)
        {
            List<string> list = new List<string>();
            if (System.Web.HttpContext.Current.Request.ApplicationPath != null)
            {
                doc.Load(OrderTypeEmailPath);
                XmlNode root = doc.SelectSingleNode("Root");
                var nodes = root.FirstChild.SelectNodes(type);
                var flag = true;
                bool.TryParse(root.Attributes[0].Value, out flag);
                if (!flag)
                {
                    nodes = root.LastChild.SelectNodes(type);
                }

                if (nodes != null && nodes.Count > 0)
                {
                    if (nodes[0].HasChildNodes)
                    {
                        list.AddRange(from XmlElement item in nodes[0] select item.InnerText);
                    }
                }
            }

            return list;
        }

        public static List<string> GetBOMStatueAddress(string status, string ps = null)
        {
            List<string> list = new List<string>();
            if (System.Web.HttpContext.Current.Request.ApplicationPath != null)
            {
                doc.Load(BOMStatueTypeEmailPath);
                XmlNode root = doc.SelectSingleNode("Root");
                var nodes = root.FirstChild.SelectNodes(status);
                var flag = true;
                bool.TryParse(root.Attributes[0].Value, out flag); // Read flag: who receive the email
                if (!flag)
                {
                    nodes = root.LastChild.SelectNodes(status);
                }

                if (nodes != null && nodes.Count > 0)
                {
                    if (nodes[0].HasChildNodes)
                    {
                        var psnodes = nodes[0].SelectNodes(ps);
                        if (psnodes != null && psnodes.Count > 0)
                        {
                            if (psnodes[0].HasChildNodes)
                            {
                                list.AddRange(from XmlElement item in psnodes[0] select item.InnerText);
                            }
                        }
                    }
                }
            }

            return list;
        }

        public static List<string> GetWHNew()
        {
            List<string> list = new List<string>();

            doc.Load(WarehouseTypeCollectionPath);
            var nodes = doc.SelectSingleNode("/root/types");
            if (nodes == null || !nodes.HasChildNodes)
            {
                return list;
            }

            foreach (XmlNode item in nodes.ChildNodes)
            {
                var attribute = item.Attributes.Count;
                if (attribute > 0)
                {
                    list.Add(item.InnerText);
                }
            }

            return list;
        }

        public static List<string> GetWHEdit()
        {
            List<string> list = new List<string>();

            doc.Load(WarehouseTypeCollectionPath);
            var nodes = doc.SelectSingleNode("/root/types");
            if (nodes == null || !nodes.HasChildNodes)
            {
                return list;
            }

            list.AddRange(from XmlNode item in nodes.ChildNodes select item.InnerText);

            return list;
        }

        public static List<string> GetBookingEmailAddress(string type)
        {
            List<string> list = new List<string>();

            doc.Load(BookingTypeEmailPath);
            var root = doc.SelectSingleNode("/root");
            var nodepath = "/root/AddressForUs/" + type;
            var flag = true;
            bool.TryParse(root.Attributes["SendToUs"].Value, out flag);
            if (!flag)
            {
                nodepath = "/root/AddressForUSA/" + type;
            }

            var nodes = doc.SelectSingleNode(nodepath);
            if (nodes == null || !nodes.HasChildNodes)
            {
                return list;
            }

            list.AddRange(from XmlNode item in nodes.ChildNodes select item.InnerText);

            return list;
        }

        public static List<string> GetBookingNodeList()
        {
            List<string> list = new List<string>();

            doc.Load(BookingTypeEmailPath);
            var nodepath = "/root/AddressForUs";
            var root = doc.SelectSingleNode("/root");
            var flag = true;
            bool.TryParse(root.Attributes["SendToUs"].Value, out flag);
            if (!flag)
            {
                nodepath = "/root/AddressForUSA";
            }

            var nodes = doc.SelectSingleNode(nodepath);
            if (nodes == null || !nodes.HasChildNodes)
            {
                return list;
            }

            list.AddRange(from XmlNode item in nodes.ChildNodes select item.Name);

            return list;
        }

        public static List<string> GetForwarderTelexReleaseEmailList(string type)
        {
            List<string> list = new List<string>();

            doc.Load(ForwarderTelexReleaseEmailPath);
            var nodepath = "/root/AddressForUs/" + type;
            var root = doc.SelectSingleNode("/root");
            var flag = true;
            bool.TryParse(root.Attributes["SendToUs"].Value, out flag);
            if (!flag)
            {
                nodepath = "/root/AddressForUSA/" + type;
            }

            var nodes = doc.SelectSingleNode(nodepath);
            if (nodes == null || !nodes.HasChildNodes)
            {
                return list;
            }

            list.AddRange(from XmlNode item in nodes.ChildNodes select item.InnerText);

            return list;
        }

        public static List<string> GetSampleReviewMessageEmailAddress(string type)
        {
            List<string> list = new List<string>();

            doc.Load(SampleReviewMessageEmailPath);
            var root = doc.SelectSingleNode("/root");
            var nodepath = "/root/AddressForUs/" + type;
            var flag = true;
            bool.TryParse(root.Attributes["SendToUs"].Value, out flag);
            if (!flag)
            {
                nodepath = "/root/AddressForUSA/" + type;
            }

            var nodes = doc.SelectSingleNode(nodepath);
            if (nodes == null || !nodes.HasChildNodes)
            {
                return list;
            }

            list.AddRange(from XmlNode item in nodes.ChildNodes select item.InnerText);

            return list;
        }

        public static List<string> GetSampleReviewMessageNodeList()
        {
            List<string> list = new List<string>();

            doc.Load(SampleReviewMessageEmailPath);
            var nodepath = "/root/AddressForUs";
            var root = doc.SelectSingleNode("/root");
            var flag = true;
            bool.TryParse(root.Attributes["SendToUs"].Value, out flag);
            if (!flag)
            {
                nodepath = "/root/AddressForUSA";
            }

            var nodes = doc.SelectSingleNode(nodepath);
            if (nodes == null || !nodes.HasChildNodes)
            {
                return list;
            }

            list.AddRange(from XmlNode item in nodes.ChildNodes select item.Name);

            return list;
        }

        public static List<string> GetDepositEmailAddress(string type)
        {
            List<string> list = new List<string>();
            if (System.Web.HttpContext.Current.Request.ApplicationPath != null)
            {
                doc.Load(PoDepositEmailPath);
                XmlNode root = doc.SelectSingleNode("Root");
                var nodes = root.FirstChild.SelectNodes(type);
                var flag = true;
                bool.TryParse(root.Attributes[0].Value, out flag);
                if (!flag)
                {
                    nodes = root.LastChild.SelectNodes(type);
                }

                if (nodes != null && nodes.Count > 0)
                {
                    if (nodes[0].HasChildNodes)
                    {
                        list.AddRange(from XmlElement item in nodes[0] select item.InnerText);
                    }
                }
            }

            return list;
        }
    }
}
