namespace VAU.Common
{
    using System;

    public class Money
    {
        private static readonly string CnNumber = "零壹贰叁肆伍陆柒捌玖";

        private static readonly string CnUnit = "分角元拾佰仟万拾佰仟亿拾佰仟兆拾佰仟";

        private static readonly string[] EnSmallNumber =
        {
            string.Empty, "ONE", "TWO", "THREE", "FOUR", "FIVE", "SIX", "SEVEN",
            "EIGHT", "NINE", "TEN", "ELEVEN", "TWELVE", "THIRTEEN",
            "FOURTEEN", "FIFTEEN", "SIXTEEN", "SEVENTEEN", "EIGHTEEN",
            "NINETEEN"
        };

        private static readonly string[] EnLargeNumber =
        {
            "TWENTY", "THIRTY", "FORTY", "FIFTY", "SIXTY", "SEVENTY",
            "EIGHTY", "NINETY"
        };

        private static readonly string[] EnUnit = { string.Empty, "THOUSAND", "MILLION", "BILLION", "TRILLION" };

        //// 以下是货币金额中文大写转换方法
        public static string GetCnString(string moneyString)
        {
            string[] tmpString = moneyString.Split('.');
            //// 默认为整数
            string intString = moneyString;
            //// 保存小数部分字串
            string decString = string.Empty;
            //// 保存中文大写字串
            string rmbCapital = string.Empty;
            int k;
            int j;
            int n;

            if (tmpString.Length > 1)
            {
                //// 取整数部分
                intString = tmpString[0];
                //// 取小数部分
                decString = tmpString[1];
            }

            decString += "00";
            //// 保留两位小数位
            decString = decString.Substring(0, 2);
            intString += decString;

            try
            {
                k = intString.Length - 1;
                if (k > 0 && k < 18)
                {
                    for (int i = 0; i <= k; i++)
                    {
                        j = intString[i] - 48;
                        //// rmbCapital = rmbCapital + cnNumber[j] + cnUnit[k-i];     // 供调试用的直接转换
                        n = i + 1 >= k ? intString[k] - 48 : intString[i + 1] - 48;
                        if (j == 0)
                        {
                            if (k - i == 2 || k - i == 6 || k - i == 10 || k - i == 14)
                            {
                                rmbCapital += CnUnit[k - i];
                            }
                            else
                            {
                                if (n != 0)
                                {
                                    rmbCapital += CnNumber[j];
                                }
                            }
                        }
                        else
                        {
                            rmbCapital = rmbCapital + CnNumber[j] + CnUnit[k - i];
                        }
                    }

                    rmbCapital = rmbCapital.Replace("兆亿万", "兆");
                    rmbCapital = rmbCapital.Replace("兆亿", "兆");
                    rmbCapital = rmbCapital.Replace("亿万", "亿");
                    rmbCapital = rmbCapital.TrimStart('元');
                    rmbCapital = rmbCapital.TrimStart('零');

                    return rmbCapital;
                }
                //// 超出转换范围时，返回零长字串
                return string.Empty;
            }
            catch
            {
                //// 含有非数值字符时，返回零长字串
                return string.Empty;
            }
        }

        //// 以下是货币金额英文大写转换方法
        public static string GetEnString(string moneyString)
        {
            string[] tmpString = moneyString.Split('.');
            //// 默认为整数
            string intString = moneyString;
            //// 保存小数部分字串
            string decString = string.Empty;
            //// 保存英文大写字串
            string engCapital = string.Empty;
            string strBuff1;
            string strBuff2;
            string strBuff3;
            int curPoint;
            int i1;
            int i2;
            int i3;
            int k;
            int n;

            if (tmpString.Length > 1)
            {
                //// 取整数部分
                intString = tmpString[0];
                //// 取小数部分
                decString = tmpString[1];
            }

            decString += "00";
            //// 保留两位小数位
            decString = decString.Substring(0, 2);

            try
            {
                // 以下处理整数部分
                curPoint = intString.Length - 1;
                if (curPoint >= 0 && curPoint < 15)
                {
                    k = 0;
                    while (curPoint >= 0)
                    {
                        strBuff1 = string.Empty;
                        strBuff2 = string.Empty;
                        strBuff3 = string.Empty;
                        if (curPoint >= 2)
                        {
                            n = int.Parse(intString.Substring(curPoint - 2, 3));
                            if (n != 0)
                            {
                                //// 取佰位数值
                                i1 = n / 100;
                                //// 取拾位数值
                                i2 = (n - (i1 * 100)) / 10;
                                //// 取个位数值
                                i3 = n - (i1 * 100) - (i2 * 10);
                                if (i1 != 0)
                                {
                                    strBuff1 = EnSmallNumber[i1] + " HUNDRED ";
                                }

                                if (i2 != 0)
                                {
                                    if (i2 == 1)
                                    {
                                        strBuff2 = EnSmallNumber[(i2 * 10) + i3] + " ";
                                    }
                                    else
                                    {
                                        strBuff2 = EnLargeNumber[i2 - 2] + " ";
                                        if (i3 != 0)
                                        {
                                            strBuff3 = EnSmallNumber[i3] + " ";
                                        }
                                    }
                                }
                                else
                                {
                                    if (i3 != 0)
                                    {
                                        strBuff3 = EnSmallNumber[i3] + " ";
                                    }
                                }

                                engCapital = strBuff1 + strBuff2 + strBuff3 + EnUnit[k] + " " + engCapital;
                            }
                        }
                        else
                        {
                            n = int.Parse(intString.Substring(0, curPoint + 1));
                            if (n != 0)
                            {
                                i2 = n / 10; // 取拾位数值
                                i3 = n - (i2 * 10); // 取个位数值
                                if (i2 != 0)
                                {
                                    if (i2 == 1)
                                    {
                                        strBuff2 = EnSmallNumber[(i2 * 10) + i3] + " ";
                                    }
                                    else
                                    {
                                        strBuff2 = EnLargeNumber[i2 - 2] + " ";
                                        if (i3 != 0)
                                        {
                                            strBuff3 = EnSmallNumber[i3] + " ";
                                        }
                                    }
                                }
                                else
                                {
                                    if (i3 != 0)
                                    {
                                        strBuff3 = EnSmallNumber[i3] + " ";
                                    }
                                }

                                engCapital = strBuff2 + strBuff3 + EnUnit[k] + " " + engCapital;
                            }
                        }

                        ++k;
                        curPoint -= 3;
                    }

                    engCapital = engCapital.TrimEnd();
                }

                // 以下处理小数部分
                strBuff2 = string.Empty;
                strBuff3 = string.Empty;
                n = int.Parse(decString);
                if (n != 0)
                {
                    i2 = n / 10; // 取拾位数值
                    i3 = n - (i2 * 10); // 取个位数值
                    if (i2 != 0)
                    {
                        if (i2 == 1)
                        {
                            strBuff2 = EnSmallNumber[(i2 * 10) + i3] + " ";
                        }
                        else
                        {
                            strBuff2 = EnLargeNumber[i2 - 2] + " ";
                            if (i3 != 0)
                            {
                                strBuff3 = EnSmallNumber[i3] + " ";
                            }
                        }
                    }
                    else
                    {
                        if (i3 != 0)
                        {
                            strBuff3 = EnSmallNumber[i3] + " ";
                        }
                    }

                    // 将小数字串追加到整数字串后
                    if (engCapital.Length > 0)
                    {
                        engCapital = engCapital + " AND CENTS " + strBuff2 + strBuff3; // 有整数部分时
                    }
                    else
                    {
                        engCapital = "CENTS " + strBuff2 + strBuff3; // 只有小数部分时
                    }
                }

                engCapital = engCapital.TrimEnd();
                return engCapital;
            }
            catch (Exception)
            {
                return string.Empty; // 含非数字字符时，返回零长字串
            }
        }
    }
}
