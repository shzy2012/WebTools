namespace VAU.Web.CommonCode
{
    using System;
    using System.Security.Cryptography;
    using System.Text;

    /// <summary>
    /// Summary description for EncrytyString
    /// </summary>
    public class EncrytyString
    {
        private static readonly string Vaukey = "222222222222222222";

        /// <summary>
        /// Encrypt the given string using the specified key.
        /// </summary>
        /// <param name="strToEncrypt">The string to be encrypted.</param>
        /// <param name="strKey">The encryption key.</param>
        /// <returns>The encrypted string.</returns>
        public static string Encrypt(string strToEncrypt)
        {
            try
            {
                TripleDESCryptoServiceProvider objDESCrypto = new TripleDESCryptoServiceProvider();
                MD5CryptoServiceProvider objHashMD5 = new MD5CryptoServiceProvider();
                byte[] byteHash, byteBuff;
                string strTempKey = Vaukey;
                byteHash = objHashMD5.ComputeHash(ASCIIEncoding.UTF8.GetBytes(strTempKey));
                objHashMD5 = null;
                objDESCrypto.Key = byteHash;

                // CBC, CFB
                objDESCrypto.Mode = CipherMode.CBC;
                byteBuff = ASCIIEncoding.UTF8.GetBytes(strToEncrypt);
                return
                    Convert.ToBase64String(
                        objDESCrypto.CreateEncryptor().TransformFinalBlock(byteBuff, 0, byteBuff.Length));
            }
            catch (Exception ex)
            {
                return "Wrong Input. " + ex.Message;
            }
        }

        /// <summary>
        /// Decrypt the given string using the specified key.
        /// </summary>
        /// <param name="strEncrypted">The string to be decrypted.</param>
        /// <param name="strKey">The decryption key.</param>
        /// <returns>The decrypted string.</returns>
        public static string Decrypt(string strEncrypted)
        {
            try
            {
                TripleDESCryptoServiceProvider objDESCrypto = new TripleDESCryptoServiceProvider();
                MD5CryptoServiceProvider objHashMD5 = new MD5CryptoServiceProvider();
                byte[] byteHash, byteBuff;
                string strTempKey = Vaukey;
                byteHash = objHashMD5.ComputeHash(ASCIIEncoding.UTF8.GetBytes(strTempKey));
                objHashMD5 = null;
                objDESCrypto.Key = byteHash;

                // CBC, CFB
                objDESCrypto.Mode = CipherMode.CBC;
                byteBuff = Convert.FromBase64String(strEncrypted);
                string strDecrypted =
                    ASCIIEncoding.UTF8.GetString(
                        objDESCrypto.CreateDecryptor().TransformFinalBlock(byteBuff, 0, byteBuff.Length));
                objDESCrypto = null;
                return strDecrypted;
            }
            catch (Exception ex)
            {
                return "Wrong Input. " + ex.Message;
            }
        }

        public static string NewEncrypt(string strToEncrypt)
        {
            string encryptKey = "2222222222222222222222222";

            DESCryptoServiceProvider descsp = new DESCryptoServiceProvider();

            byte[] key = Encoding.Unicode.GetBytes(encryptKey);

            byte[] data = Encoding.Unicode.GetBytes(strToEncrypt);

            System.IO.MemoryStream mstream = new System.IO.MemoryStream();

            CryptoStream cstream = new CryptoStream(mstream, descsp.CreateEncryptor(key, key), CryptoStreamMode.Write);

            cstream.Write(data, 0, data.Length);

            cstream.FlushFinalBlock();
            var encrypted = mstream.ToArray();
            char[] chars = new char[encrypted.Length / sizeof(char)];
            System.Buffer.BlockCopy(encrypted, 0, chars, 0, encrypted.Length);
            return new string(chars);
        }

        public static string NewDecrypt(string strEncrypted)
        {
            string encryptKey = "222222222222222222222";

            DESCryptoServiceProvider descsp = new DESCryptoServiceProvider();

            byte[] key = Encoding.Unicode.GetBytes(encryptKey);

            byte[] data = new byte[strEncrypted.Length * sizeof(char)];
            System.Buffer.BlockCopy(strEncrypted.ToCharArray(), 0, data, 0, data.Length);

            System.IO.MemoryStream mstream = new System.IO.MemoryStream();

            CryptoStream cstream = new CryptoStream(mstream, descsp.CreateDecryptor(key, key), CryptoStreamMode.Write);

            cstream.Write(data, 0, data.Length);

            cstream.FlushFinalBlock();

            return Encoding.Unicode.GetString(mstream.ToArray());
        }
    }
}
