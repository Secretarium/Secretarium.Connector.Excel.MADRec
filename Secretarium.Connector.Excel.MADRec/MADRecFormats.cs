using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace Secretarium.Excel
{
    class MADRecFormats
    {
        public static List<string> MADREC_FIELDS = new List<string> { "LEI", "BIC", "EMIR", "NACE", "IF", "GK", "PERMID" };

        public static HashSet<string> ISO_3166_1 = new HashSet<string> {
            "AD", "AE", "AF", "AG", "AI", "AL", "AM", "AO", "AQ", "AR", "AS", "AT", "AU", "AW", "AX", "AZ", "BA", "BB", "BD", "BE", "BF",
            "BG", "BH", "BI", "BJ", "BL", "BM", "BN", "BO", "BQ", "BR", "BS", "BT", "BV", "BW", "BY", "BZ", "CA", "CC", "CD", "CF", "CG",
            "CH", "CI", "CK", "CL", "CM", "CN", "CO", "CR", "CU", "CV", "CW", "CX", "CY", "CZ", "DE", "DJ", "DK", "DM", "DO", "DZ", "EC",
            "EE", "EG", "EH", "ER", "ES", "ET", "FI", "FJ", "FK", "FM", "FO", "FR", "GA", "GB", "GD", "GE", "GF", "GG", "GH", "GI", "GL",
            "GM", "GN", "GP", "GQ", "GR", "GS", "GT", "GU", "GW", "GY", "HK", "HM", "HN", "HR", "HT", "HU", "ID", "IE", "IL", "IM", "IN",
            "IO", "IQ", "IR", "IS", "IT", "JE", "JM", "JO", "JP", "KE", "KG", "KH", "KI", "KM", "KN", "KP", "KR", "KW", "KY", "KZ", "LA",
            "LB", "LC", "LI", "LK", "LR", "LS", "LT", "LU", "LV", "LY", "MA", "MC", "MD", "ME", "MF", "MG", "MH", "MK", "ML", "MM", "MN",
            "MO", "MP", "MQ", "MR", "MS", "MT", "MU", "MV", "MW", "MX", "MY", "MZ", "NA", "NC", "NE", "NF", "NG", "NI", "NL", "NO", "NP",
            "NR", "NU", "NZ", "OM", "PA", "PE", "PF", "PG", "PH", "PK", "PL", "PM", "PN", "PR", "PS", "PT", "PW", "PY", "QA", "RE", "RO",
            "RS", "RU", "RW", "SA", "SB", "SC", "SD", "SE", "SG", "SH", "SI", "SJ", "SK", "SL", "SM", "SN", "SO", "SR", "SS", "ST", "SV",
            "SX", "SY", "SZ", "TC", "TD", "TF", "TG", "TH", "TJ", "TK", "TL", "TM", "TN", "TO", "TR", "TT", "TV", "TW", "TZ", "UA", "UG",
            "UM", "US", "UY", "UZ", "VA", "VC", "VE", "VG", "VI", "VN", "VU", "WF", "WS", "YE", "YT", "ZA", "ZM", "ZW" };
        public static Regex LEI = new Regex(@"^[0-9A-Z]{4}[0-9A-Z]{14}\d{2}$");
        public static Regex BIC = new Regex(@"^[0-9A-Z]{4}([A-Z]{2})[0-9A-Z]{2}([0-9A-Z]{3})?$");
        public static List<string> EMIR = new List<string> { "FC", "NFC-", "NFC+" };
        public static Regex NACE = new Regex(@"^\d{1,2}(\.\d{1,2})?$");
        public static Regex IF = new Regex(@"^(true|false|y|n)$i");
        public static Regex GK = new Regex(@"^[1-9]\d{0,5}$");
        public static Regex PERMID = new Regex(@"^\d{10}$");

        public static int ISO_7064_MOD_97_10(string s)
        {
            int p = 0;
            for (var i = 0; i < s.Length; i++)
            {
                var c = (int)s[i];
                p = c + (c >= 65 ? p * 100 - 55 : p * 10 - 48); /* 'A' is 65, '0' is 48 */
                if (p > 1000000) p %= 97;
            }
            return p % 97;
        }

        public static bool Verify(string field, ref object value, out string error)
        {
            error = "";

            switch (field)
            {
                case "LEI":
                    var x = value as string;
                    if (x == null)
                    {
                        error = "expecting a string";
                        return false;
                    }
                    x = x.ToUpper();
                    var m = LEI.Match(x);
                    if(!m.Success)
                    {
                        error = "format is incorrect (ISO 17442)";
                        return false;
                    }
                    if (ISO_7064_MOD_97_10(x) != 1)
                    {
                        error = "incorrect check sum";
                        return false;
                    }

                    value = x;
                    return true;

                case "BIC":
                    x = value as string;
                    if (x == null)
                    {
                        error = "expecting a string";
                        return false;
                    }
                    x = x.ToUpper();
                    m = BIC.Match(x);
                    if (!m.Success || m.Groups.Count < 2)
                    {
                        error = "format is incorrect";
                        return false;
                    }
                    if (!ISO_3166_1.Contains(m.Groups[1].Value))
                    {
                        error = m.Groups[1].Value + " is not a valid country (ISO 3166-1)";
                        return false;
                    }
                    if (x.Length == 8)
                    {
                        x += "XXX";
                    }

                    value = x;
                    return true;

                case "EMIR":
                    x = value as string;
                    if (x == null)
                    {
                        error = "expecting a string";
                        return false;
                    }
                    x = x.ToUpper();
                    if ((x.Length != 2 && x.Length != 4) || !EMIR.Contains(x))
                    {
                        error = "format is incorrect";
                        return false;
                    }

                    value = x;
                    return true;

                case "NACE":
                    x = value as string;
                    if (x == null)
                    {
                        error = "expecting a string";
                        return false;
                    }
                    m = NACE.Match(x);
                    if (!m.Success)
                    {
                        error = "format is incorrect";
                        return false;
                    }

                    value = x;
                    return true;

                case "IF":
                    if (value is bool b)
                        return true;

                    x = value as string;
                    if (x == null)
                    {
                        error = "expecting a string";
                        return false;
                    }
                    m = IF.Match(x);
                    if (!m.Success)
                    {
                        error = "format is incorrect";
                        return false;
                    }
                    x = x.ToUpper();

                    value = x == "TRUE" || x == "Y";
                    return true;

                case "GK":
                    x = value as string;
                    if (x == null)
                    {
                        error = "expecting a string";
                        return false;
                    }
                    m = GK.Match(x);
                    if (!m.Success)
                    {
                        error = "format is incorrect";
                        return false;
                    }

                    value = x;
                    return true;

                case "PERMID":
                    x = value as string;
                    if (x == null)
                    {
                        error = "expecting a string";
                        return false;
                    }
                    m = PERMID.Match(x);
                    if (!m.Success)
                    {
                        error = "format is incorrect";
                        return false;
                    }

                    value = x;
                    return true;
            }

            return false;
        }
    }
}
