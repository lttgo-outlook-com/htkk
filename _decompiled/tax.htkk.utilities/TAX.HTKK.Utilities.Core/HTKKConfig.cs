using System.Collections.Generic;
using System.Configuration;
using System.Windows.Forms;

namespace TAX.HTKK.Utilities.Core;

public class HTKKConfig
{
	private static string _APPPATH;

	private static string _LOGPATH;

	private static string[] _TOAN_TU;

	private static string _PAGE_SIZE;

	private static string _PATH_TEMP;

	private static string _PRODUCT_NAME;

	private static string _VERSION;

	private static string _APP_PREFIX;

	private static string _ENABLE_LOGGING;

	private static string _MA_DICH_VU;

	private static string _TEN_DICH_VU;

	private static string _THONG_TIN_NHA_CUNG_CAP;

	private static string _PHIEN_BAN_TO_KHAI_XML;

	private static string _INTERFACE_TEMPLATES;

	private static string _XMLSCHEMA;

	private static string _REPORT_TEMPLATES;

	private static string _BANGKEPREF;

	private static string _INTERFACE_INI;

	private static string _DATAFILES;

	private static string _HISTORY;

	private static int _MAX_ROW_EXCEL = 0;

	private static Dictionary<string, string> _CELL_DATA = new Dictionary<string, string>();

	private static string _URLUPDATE;

	public static string APPPATH
	{
		get
		{
			if (string.IsNullOrEmpty(_APPPATH))
			{
				_APPPATH = Application.StartupPath;
			}
			return _APPPATH;
		}
	}

	public static string LOGPATH
	{
		get
		{
			if (string.IsNullOrEmpty(_LOGPATH))
			{
				_LOGPATH = Application.StartupPath + "\\HTKK_Log";
			}
			return _LOGPATH;
		}
	}

	public static string[] TOAN_TU
	{
		get
		{
			if (_TOAN_TU == null || _TOAN_TU.Length == 0)
			{
				_TOAN_TU = ConfigurationManager.AppSettings["TOAN_TU"].Split(new char[1] { ',' });
			}
			return _TOAN_TU;
		}
	}

	public static string PAGE_SIZE
	{
		get
		{
			if (string.IsNullOrEmpty(_PAGE_SIZE))
			{
				_PAGE_SIZE = ConfigurationManager.AppSettings["PAGE_SIZE"];
			}
			return _PAGE_SIZE;
		}
	}

	public static string PATH_TEMP
	{
		get
		{
			if (string.IsNullOrEmpty(_PATH_TEMP))
			{
				_PATH_TEMP = ConfigurationManager.AppSettings["PATH_TEMP"];
			}
			return _PATH_TEMP;
		}
	}

	public static string PRODUCT_NAME
	{
		get
		{
			if (string.IsNullOrEmpty(_PRODUCT_NAME))
			{
				_PRODUCT_NAME = ConfigurationManager.AppSettings["PRODUCT_NAME"];
			}
			return _PRODUCT_NAME;
		}
	}

	public static string VERSION
	{
		get
		{
			if (string.IsNullOrEmpty(_VERSION))
			{
				_VERSION = ConfigurationManager.AppSettings["VERSION"];
			}
			return _VERSION;
		}
	}

	public static string APP_PREFIX
	{
		get
		{
			if (string.IsNullOrEmpty(_APP_PREFIX))
			{
				_APP_PREFIX = ConfigurationManager.AppSettings["APP_PREFIX"];
			}
			return _APP_PREFIX;
		}
	}

	public static string ENABLE_LOGGING
	{
		get
		{
			if (string.IsNullOrEmpty(_ENABLE_LOGGING))
			{
				_ENABLE_LOGGING = ConfigurationManager.AppSettings["ENABLE_LOGGING"];
			}
			return _ENABLE_LOGGING;
		}
	}

	public static string MA_DICH_VU
	{
		get
		{
			if (string.IsNullOrEmpty(_MA_DICH_VU))
			{
				_MA_DICH_VU = ConfigurationManager.AppSettings["MA_DICH_VU"];
			}
			return _MA_DICH_VU;
		}
	}

	public static string TEN_DICH_VU
	{
		get
		{
			if (string.IsNullOrEmpty(_TEN_DICH_VU))
			{
				_TEN_DICH_VU = ConfigurationManager.AppSettings["TEN_DICH_VU"];
			}
			return _TEN_DICH_VU;
		}
	}

	public static string THONG_TIN_NHA_CUNG_CAP
	{
		get
		{
			if (string.IsNullOrEmpty(_THONG_TIN_NHA_CUNG_CAP))
			{
				_THONG_TIN_NHA_CUNG_CAP = ConfigurationManager.AppSettings["THONG_TIN_NHA_CUNG_CAP"];
			}
			return _THONG_TIN_NHA_CUNG_CAP;
		}
	}

	public static string PHIEN_BAN_TO_KHAI_XML
	{
		get
		{
			if (string.IsNullOrEmpty(_PHIEN_BAN_TO_KHAI_XML))
			{
				_PHIEN_BAN_TO_KHAI_XML = ConfigurationManager.AppSettings["PHIEN_BAN_TO_KHAI_XML"];
			}
			return _PHIEN_BAN_TO_KHAI_XML;
		}
	}

	public static string INTERFACE_TEMPLATES
	{
		get
		{
			if (string.IsNullOrEmpty(_INTERFACE_TEMPLATES))
			{
				_INTERFACE_TEMPLATES = ConfigurationManager.AppSettings["InterfaceTemplates"];
			}
			return _INTERFACE_TEMPLATES;
		}
	}

	public static string XMLSCHEMA
	{
		get
		{
			if (string.IsNullOrEmpty(_XMLSCHEMA))
			{
				_XMLSCHEMA = ConfigurationManager.AppSettings["XmlSchema"];
			}
			return _XMLSCHEMA;
		}
	}

	public static string REPORT_TEMPLATES
	{
		get
		{
			if (string.IsNullOrEmpty(_REPORT_TEMPLATES))
			{
				_REPORT_TEMPLATES = ConfigurationManager.AppSettings["ReportTemplates"];
			}
			return _REPORT_TEMPLATES;
		}
	}

	public static string BANGKEPREF
	{
		get
		{
			if (string.IsNullOrEmpty(_BANGKEPREF))
			{
				_BANGKEPREF = ConfigurationManager.AppSettings["BangKePref"];
			}
			return _BANGKEPREF;
		}
	}

	public static string INTERFACE_INI
	{
		get
		{
			if (string.IsNullOrEmpty(_INTERFACE_INI))
			{
				_INTERFACE_INI = ConfigurationManager.AppSettings["InterfaceIni"];
			}
			return _INTERFACE_INI;
		}
	}

	public static string DATAFILES
	{
		get
		{
			if (string.IsNullOrEmpty(_DATAFILES))
			{
				_DATAFILES = ConfigurationManager.AppSettings["DataFiles"];
			}
			return _DATAFILES;
		}
	}

	public static string HISTORY
	{
		get
		{
			if (string.IsNullOrEmpty(_HISTORY))
			{
				_HISTORY = ConfigurationManager.AppSettings["History"];
			}
			return _HISTORY;
		}
	}

	public static int MAX_ROW_EXCEL
	{
		get
		{
			if (_MAX_ROW_EXCEL == 0)
			{
				_MAX_ROW_EXCEL = int.Parse(ConfigurationManager.AppSettings["MAX_ROW_EXCEL"]);
			}
			return _MAX_ROW_EXCEL;
		}
	}

	public static Dictionary<string, string> CELL_DATA
	{
		get
		{
			if (_CELL_DATA == null || _CELL_DATA.Count == 0)
			{
				string[] array = new string[130]
				{
					"A", "B", "C", "D", "E", "F", "G", "H", "I", "J",
					"K", "L", "M", "N", "O", "P", "Q", "R", "S", "T",
					"U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD",
					"AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL", "AM", "AN",
					"AO", "AP", "AQ", "AR", "AS", "AT", "AU", "AV", "AW", "AX",
					"AY", "AZ", "BA", "BB", "BC", "BD", "BE", "BF", "BG", "BH",
					"BI", "BJ", "BK", "BL", "BM", "BN", "BO", "BP", "BQ", "BR",
					"BS", "BT", "BU", "BV", "BW", "BX", "BY", "BZ", "CA", "CB",
					"CC", "CD", "CE", "CF", "CG", "CH", "CI", "CJ", "CK", "CL",
					"CM", "CN", "CO", "CP", "CQ", "CR", "CS", "CT", "CU", "CV",
					"CW", "CX", "CY", "CZ", "DA", "DB", "DC", "DD", "DE", "DF",
					"DG", "DH", "DI", "DJ", "DK", "DL", "DM", "DN", "DO", "DP",
					"DQ", "DR", "DS", "DT", "DU", "DV", "DW", "DX", "DY", "DZ"
				};
				for (int i = 0; i < array.Length; i++)
				{
					_CELL_DATA.Add(array[i], (i + 1).ToString());
				}
			}
			return _CELL_DATA;
		}
	}

	public static string URLUPDATE
	{
		get
		{
			if (string.IsNullOrEmpty(_URLUPDATE))
			{
				_URLUPDATE = ConfigurationManager.AppSettings["UrlAutoUpdates"];
			}
			return _URLUPDATE;
		}
	}
}
