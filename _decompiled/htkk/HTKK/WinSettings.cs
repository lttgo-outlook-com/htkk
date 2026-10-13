using System.Collections.Generic;
using System.Data;
using System.Xml;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Entity;

namespace HTKK;

public class WinSettings
{
	public static DataSet dsTasks = new DataSet();

	public static short OnLoad;

	public static string NewVersionOnServer = string.Empty;

	public static string CurentVersionOnClient = "4.0.0";

	public static Dictionary<string, int> DicSheet = new Dictionary<string, int>();

	public static object InstanceClass;

	public static bool bTimKiem = false;

	public static string MenuID = string.Empty;

	public static string KieuKy = string.Empty;

	public static string VersionApp = "";

	public static string ExportingData = "";

	public static bool bCanhBaoDo = false;

	public static bool bCanhBaoVang = false;

	public static bool isReadDataXML = false;

	public static HTKKCMCSheetInfo KHBS_LanTruoc = new HTKKCMCSheetInfo();

	public static HTKKCMCToKhaiInfo ToKhaiInfo_BoSungLanTruoc = new HTKKCMCToKhaiInfo();

	public static List<HTKKCMCKHBS> KHBS = new List<HTKKCMCKHBS>();

	public static HTKKCMCToKhaiInfo ToKhaiInfo = new HTKKCMCToKhaiInfo();

	public static XmlNode xmlMenuNode;

	public static HTKKHeaderInfo HTKKHeaderInfo = new HTKKHeaderInfo();

	public static HTKKPeriodInfo TTPeriodInfo = new HTKKPeriodInfo();

	public static HTKKCMCTTinChungInfo TTinChung = new HTKKCMCTTinChungInfo();

	public static PrintInfo PrintInfo = new PrintInfo();

	public static frmHTKKSystems myHTKKSystems;

	public static frmWinToKhai myWinToKhai;

	public static Dictionary<HTKKSheetInfo, DictProcessInfo> DictRunWorkerAsync = new Dictionary<HTKKSheetInfo, DictProcessInfo>();

	public static Dictionary<int, string> DictSaveRunWorkerAsync = new Dictionary<int, string>();

	public static SortedDictionary<int, string> DictSortData = new SortedDictionary<int, string>();

	public static XmlDataDocument xmlDocMenu = new XmlDataDocument();

	public static XmlDocument XmlMessage = new XmlDocument();

	public static XmlNodeList XmlNodeConfigCell = null;

	public static int treeIndex = -1;

	public static Dictionary<int, Dictionary<string, List<DictInfo>>> WinListMessageErorr = new Dictionary<int, Dictionary<string, List<DictInfo>>>();

	public static bool isActionPrint = false;

	public static bool chuyenMSTKhac = false;

	public static Dictionary<string, string> DictTypeValidate = new Dictionary<string, string>
	{
		{ "T001", "ValidateInt" },
		{ "T002", "ValidateNumberBetween" },
		{ "T003", "ValidateFormula" },
		{ "T004", "ValidateDouble" },
		{ "T005", "ValidateStringChecklength" },
		{ "T006", "ValidateEmail" },
		{ "T007", "ValidateMaSoThue" },
		{ "T008", "ValidateCMTND" },
		{ "T009", "ValidateBangNhau" },
		{ "T010", "ValidateLonHon" },
		{ "T011", "ValidateNhoHon" },
		{ "T012", "ValidateDuplicate" }
	};

	public static Dictionary<string, string> DictTTinChung = new Dictionary<string, string>
	{
		{ "maDVu", "22_20" },
		{ "tenDVu", "22_21" },
		{ "pbanDVu", "22_22" },
		{ "ttinNhaCCapDVu", "22_23" },
		{ "maTKhai", "22_24" },
		{ "tenTKhai", "22_25" },
		{ "moTaBMau", "22_26" },
		{ "pbanTKhaiXML", "22_27" },
		{ "loaiTKhai", "22_28" },
		{ "soLan", "22_29" },
		{ "kieuKy", "22_30" },
		{ "kyKKhai", "22_31" },
		{ "kyKKhaiTuNgay", "22_32" },
		{ "kyKKhaiDenNgay", "22_33" },
		{ "kyKKhaiTuThang", "22_34" },
		{ "kyKKhaiDenThang", "22_35" },
		{ "maCQTNoiNop", "22_36" },
		{ "tenCQTNoiNop", "22_37" },
		{ "ngayLapTKhai", "22_38" },
		{ "nguoiKy", "22_41" },
		{ "ngayKy", "22_42" },
		{ "nganhNgheKD", "22_43" },
		{ "mst", "3_2" },
		{ "tenNNT", "3_3" },
		{ "dchiNNT", "3_4" },
		{ "phuongXa", "" },
		{ "maHuyenNNT", "" },
		{ "tenXaPhuongNNT", "3_5" },
		{ "maTinhNNT", "" },
		{ "tenTinhNNT", "3_6" },
		{ "dthoaiNNT", "3_7" },
		{ "faxNNT", "3_8" },
		{ "emailNNT", "3_9" },
		{ "SoTKNganHang", "22_2" },
		{ "TenDonViChuQuan", "22_3" },
		{ "MaSoThueDonViChuQuan", "22_4" },
		{ "mstDLyThue", "22_5" },
		{ "tenDLyThue", "22_6" },
		{ "dchiDLyThue", "22_7" },
		{ "maHuyenDLyThue", "" },
		{ "tenHuyenDLyThue", "" },
		{ "tenXaPhuongDLyThue", "22_8" },
		{ "maTinhDLyThue", "" },
		{ "tenTinhDLyThue", "22_9" },
		{ "dthoaiDLyThue", "22_10" },
		{ "faxDLyThue", "22_11" },
		{ "emailDLyThue", "22_12" },
		{ "soHDongDLyThue", "22_13" },
		{ "ngayKyHDDLyThue", "22_14" },
		{ "tenNVienDLyThue", "22_15" },
		{ "cchiHNghe", "22_16" },
		{ "maCQTCapCuc", "22_44" },
		{ "tenCQTCapCuc", "22_45" },
		{ "thang", "22_46" },
		{ "quy", "22_47" },
		{ "nam", "22_48" },
		{ "maDonViTinh", "22_49" },
		{ "tenDonViTinh", "22_50" },
		{ "ngay", "22_51" },
		{ "ngayLapTKhaiBS", "22_52" }
	};

	public static Dictionary<string, string> DictMapCotrols = new Dictionary<string, string>
	{
		{ "M", "UcThang" },
		{ "D", "UcNgay" },
		{ "Q", "UcQuy" },
		{ "Y", "UcNam" },
		{ "K", "UcKy" },
		{ "D_Y", "UcNgay" },
		{ "D_D", "UcNgay" },
		{ "M_Y", "UcThang" },
		{ "KTN", "UcThang" },
		{ "CD", "UcNgayCD" },
		{ "DT", "UcNgayCD" },
		{ "KTN_Y", "UcNam" },
		{ "CD_Y", "UcNam" },
		{ "DT_Y", "UcNam" }
	};
}
