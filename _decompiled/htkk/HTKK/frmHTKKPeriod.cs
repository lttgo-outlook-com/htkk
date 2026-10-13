using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Windows.Forms.Layout;
using System.Xml;
using FarPoint.Win.Spread;
using FarPoint.Win.Spread.Model;
using TAX.HTKK.Business;
using TAX.HTKK.Utilities.Control;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace HTKK;

public class frmHTKKPeriod : Form
{
	private static XmlNode xmlNode;

	private static XmlNode xmlNodeGoc;

	private static string menuID = "";

	private static string periodTypeValue = "";

	private static string parentMenuID = "";

	private static string startDate = "";

	private int ckbPadL = 5;

	private int ctrPadR = 10;

	private string strKieuKy;

	private string strDanhMucNgheFile;

	private int defaultYear;

	private int defaultMonth;

	private int defaultDay;

	private string strDataFileToKhaiChinh;

	private string strDataFileToKHBS;

	private string strVersionXML;

	private string strSheetNameHeader;

	private string strSheetNameKHBS;

	private static frmHTKKSystems frmHTKKSystems;

	private XmlNodeList elemList;

	private string strFileNameExtend;

	private string fullNameDataFileToKhaiChinh;

	private List<CheckBox> listCheckBox = new List<CheckBox>();

	private List<MaskedTextBox> listMaskedTextBox = new List<MaskedTextBox>();

	private List<ComboBox> listComboBox = new List<ComboBox>();

	private List<RadioButton> listRadioButton = new List<RadioButton>();

	private int numberColCheckBox = 2;

	private bool isCheckOptions = false;

	private int widthPeriod = 0;

	private List<KeyValuePair<string, string>> dataCboCQTCapCuc = null;

	private List<KeyValuePair<string, string>> dataCboCQTChiCuc = null;

	private List<KeyValuePair<string, string>> dataCboTinh = null;

	private List<KeyValuePair<string, string>> dataCboHuyen = null;

	private Dictionary<string, string> dataCboDADT = new Dictionary<string, string>();

	private string _maCQTCapCuc = "";

	private string _tenCQTCapCuc = "";

	private string _maCQTNoiNop = "";

	private string _tenCQTNoiNop = "";

	private IContainer components = null;

	private RadioButton OptTKThang;

	private RadioButton OptTKLanPS;

	private CheckBox chkDauTho;

	private CheckBox chkKhiThien;

	private CheckBox chkCondensate;

	private Panel pnVT3Button;

	private Button btnDongY;

	private Button cmdClose;

	private Panel pnVT2_PhuLucKeKhai;

	private FpSpread axfpsPhuLuc;

	private SheetView axfpsPhuLuc_Sheet1;

	private CheckBox chkChonPhuLucKeKhaiAll;

	private Panel f;

	private Panel pnVT1;

	private Panel pnVT1_1CheckBox;

	private CheckBox chkTKLanPS;

	private CheckBox ckNam;

	private CheckBox chkTKhaiLanXB;

	private CheckBox chkTKKy;

	private CheckBox chkTKQuy;

	private CheckBox chkTuThangDenThang;

	private CheckBox chkTkhaiThang;

	private Panel pnVT1_4LanDauBoXung;

	private Panel pnQuyetToan_LienTuc;

	private CheckBox chkQTNamDau_KhongHoatDongLienTuc;

	private CheckBox chkQTTungNam_HoatDongLienTuc;

	private Panel pnNgayLap;

	private MaskedTextBox mskNgayLapKHBS;

	private Label lblNgayLapKHBS;

	private Panel pnRDLanDauBoXung;

	private RadioButton OptBosung;

	private Panel pnLan;

	private Label lblLan;

	private RadioButton OptChinhthuc;

	private Panel pnVT1_2NgayThangQuyNam;

	private Panel pnKy;

	private ComboBox cbbKy;

	private Label lblKy;

	private Panel pnNam;

	private MaskedTextBox mskYear;

	private Label lblYear;

	private Panel pnNgay;

	private Label lblNgay;

	private MaskedTextBox mskDay;

	private Panel pnQuy;

	private ComboBox cbbQuy;

	private Label lblQuy;

	private Panel pnThang;

	private MaskedTextBox mskMonth;

	private Label lblMonth;

	private Panel pnVT1_5DanhMucNghe;

	private Label lblNganhKD;

	private ComboBox cboNganhKD;

	private Panel pnVT1_3TuNgayDenNgay;

	private Panel pnDenThang;

	private MaskedTextBox mskDenThang;

	private Label lblDenThang;

	private Panel pnLanXuatBan;

	private Label lblLanXuat;

	private Panel pnDenNgay;

	private MaskedTextBox mskDenNgay;

	private Label lblDenNgay;

	private Panel pnTuThang;

	private MaskedTextBox mskTuThang;

	private Label lblTuThang;

	private Panel pnTuNgay;

	private MaskedTextBox mskTuNgay;

	private Label lblTuNgay;

	private PictureBox pbHeader;

	private TextBox txtLanXuatBan;

	private TextBox txtLan;

	private CheckBox chkToChucKhaiThay;

	private CheckBox chk_CoSoThanhLap;

	private Label lbNgayPS;

	private MaskedTextBox mskNgayPS;

	private CheckBox cbxLNDC;

	private CheckBox cbxCTDC;

	private CheckBox chkTuQuyDenQuy;

	private Panel pnMaHSDNHT;

	private TextBox txtMaHSDNHT_DADT;

	private Label label1;

	private Label lblChuDADT;

	private Panel pnThuHo;

	private CheckBox chkThuHo;

	private ComboBox cbbMaHSDADT;

	private Panel pnPPhapNopThue;

	private CheckBox chkMSTCaNhan;

	private CheckBox chkMSTToChuc;

	private Panel pnMSTCaNhanToChuc;

	private Panel pnTinhCQT;

	private ComboBox cboCQTChiCuc;

	private ComboBox cboCQTCuc;

	private Panel pnKhaiTheoPLuatDanSuThue;

	private CheckBox chkKhaiTheoPLuatThue;

	private CheckBox chkKhaiTheoPLuatDanSu;

	private ComboBox cboTinh;

	private Label lblCQTChiCuc;

	private Label lblCQTCuc;

	private Label lblTinh;

	private Panel pnChuyenNhuong;

	private CheckBox chkDoanhNghiepChuyenNhuong;

	private CheckBox chkToChucCaNhanChuyenNhuong;

	private Panel pnCCT;

	private Panel pnCQTCuc;

	private Panel pnTinh;

	private Panel pnLoaiKyHoan;

	private ComboBox cboKieuKy;

	private Label label2;

	private Panel pnTruongHopQT;

	private ComboBox cboTruongHopQT;

	private Label label3;

	private Panel pnRDLanDauBoXung1;

	private Label label4;

	private ComboBox cboLoaiToKhai;

	private Panel pnLan1;

	private TextBox txtLan1;

	private Label label5;

	private Panel pnNgayLap1;

	private MaskedTextBox mskNgayLapKHBS1;

	private Label lblNgayLapKHBS1;

	private ComboBox cbMaHSDNHT_DADT;

	private TextBox txtMst;

	private Label lblMst;

	private Panel pnMst;

	private Panel pnLoaiTien;

	private ComboBox cboLoaiTien;

	private Label lblLoaiTien;

	private Panel pnThongBao;

	private MaskedTextBox mskNgayTb;

	private Label lblNgayTb;

	private TextBox txtSoTb;

	private Label lblSoTB;

	private Panel panel1;

	private ComboBox cboXa;

	private Label lblXa;

	private Label lblNgayPsThuNhap;

	public frmHTKKPeriod(frmHTKKSystems _frmHTKKSystems, XmlNode _xmlNode, int _frmThemPhuLuc)
	{
		frmHTKKSystems = _frmHTKKSystems;
		xmlNode = _xmlNode;
		xmlNodeGoc = _xmlNode;
		SetDataFile();
		InitializeComponent();
		ControlInit(_xmlNode);
		FormInit();
	}

	private void frmHTKKPeriod_Load(object sender, EventArgs e)
	{
		//IL_080e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0815: Expected O, but got Unknown
		//IL_0101: Unknown result type (might be due to invalid IL or missing references)
		//IL_0108: Expected O, but got Unknown
		//IL_0855: Unknown result type (might be due to invalid IL or missing references)
		//IL_085c: Expected O, but got Unknown
		//IL_0148: Unknown result type (might be due to invalid IL or missing references)
		//IL_014f: Expected O, but got Unknown
		//IL_1123: Unknown result type (might be due to invalid IL or missing references)
		//IL_112a: Expected O, but got Unknown
		//IL_12fa: Unknown result type (might be due to invalid IL or missing references)
		//IL_1301: Expected O, but got Unknown
		//IL_1243: Unknown result type (might be due to invalid IL or missing references)
		//IL_124d: Expected O, but got Unknown
		//IL_138b: Unknown result type (might be due to invalid IL or missing references)
		//IL_1395: Expected O, but got Unknown
		try
		{
			if (WinSettings.MenuID == "T6")
			{
				WinFnProcess.RenameDatafileIfError();
			}
			WinFnProcess.Createa_PL_KHBS();
			WinFnProcess.FixDatafileConfig_03_3C_TNDN_TT80();
			WinFnProcess.FixFileNameKHBS_01_GTGT_TT80();
			WinFnProcess.FixFileNameKHBS_02_TTDB_TT80();
			WinFnProcess.FixFileNameThangToiNgay("02_PBVMT_TT80", "251");
			WinFnProcess.FixFileNameThangToiNgay("02_PH_TT80", "251");
			WinFnProcess.FixFileNameThangToiNgay("02_TAIN_TT80", "251");
			WinFnProcess.FixFileNameThangToiNgay("02_1_TAIN_TT80", "251");
			WinFnProcess.FixFileNameNamToiNgay("04_TNDN_TT80", "251");
			WinFnProcess.FixFileNameNamToiNgay("04_2_MGT_TNDN_TT80", "251");
			WinFnProcess.RemoveCellById("04_TNDN_TT80", "251", "AI_21,AJ_21,S_27,CR_27");
			WinFnProcess.FixFileDauKhiLoaiTien();
			WinFnProcess.FixPathStatic("02_GTGT", "241", "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/TTDuAnDauTu/diaChiThucHienDuAn", "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/TTDuAnDauTu/soNha");
			WinFnProcess.ConvertCatalogue_BVMT_CF();
			string dataFileName = "03_TNDN~03_1A_TNDN~03_1B_TNDN~03_1C_TNDN~03_2A_TNDN~03_2B_TNDN~03_3A_TNDN~03_3B_TNDN~03_3C_TNDN~03_4_TNDN~03_5_TNDN~03_6_TNDN~03_7_TNDN~03_8_TNDN~03_9_TD_TNDN~03_02_TNDN~03_01_ND20_TNDN~03_02_ND20_TNDN~03_03_ND20_TNDN~03_04_ND20_TNDN~KHBS_03_TNDN";
			string versionXmlOld = "231";
			string versionXml = "239";
			WinFnProcess.ConverDataVersion(dataFileName, versionXmlOld, versionXml);
			try
			{
				string text = HTKKConfig.INTERFACE_INI + "\\ConvertVersion.xml";
				XmlDocument val = new XmlDocument();
				if (File.Exists(text))
				{
					val.Load(text);
					XmlNodeList elementsByTagName = val.GetElementsByTagName("Item");
					foreach (XmlNode item in elementsByTagName)
					{
						XmlNode val2 = item;
						string value = ((XmlNode)val2.Attributes["DataFileName"]).Value;
						string value2 = ((XmlNode)val2.Attributes["VersionXmlOld"]).Value;
						string value3 = ((XmlNode)val2.Attributes["VersionXmlNew"]).Value;
						string value4 = ((XmlNode)val2.Attributes["ConvertType"]).Value;
						string versionXmlKHBSOld = ((val2.Attributes["VersionXmlKHBSOld"] != null) ? ((XmlNode)val2.Attributes["VersionXmlKHBSOld"]).Value : "");
						string versionXmlKHBSNew = ((val2.Attributes["VersionXmlKHBSNew"] != null) ? ((XmlNode)val2.Attributes["VersionXmlKHBSNew"]).Value : "");
						int insertSection = ((val2.Attributes["InsertSection"] != null) ? Convert.ToInt32(((XmlNode)val2.Attributes["InsertSection"]).Value) : (-1));
						int restart = ((val2.Attributes["Restart"] != null) ? Convert.ToInt32(((XmlNode)val2.Attributes["Restart"]).Value) : 0);
						int versionApp = ((val2.Attributes["VersionApp"] != null) ? Convert.ToInt32(((XmlNode)val2.Attributes["VersionApp"]).Value) : (-1));
						bool noRemoveAfterBK = ((val2.Attributes["NoRemoveAfterBK"] != null && ((XmlNode)val2.Attributes["NoRemoveAfterBK"]).Value == "1") ? true : false);
						string prefix = ((val2.Attributes["Prefix"] != null) ? ((XmlNode)val2.Attributes["Prefix"]).Value : "");
						WinFnProcess.ConverDataVersionCommon(value, value2, value3, value4, insertSection, restart, versionXmlKHBSOld, versionXmlKHBSNew, versionApp, noRemoveAfterBK, prefix);
					}
				}
			}
			catch (Exception ex)
			{
				HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			}
			WinFnProcess.FixStt("01_9_MGT_GTGT_TT80", "268", "pathStt");
			WinFnProcess.FixStt("01_4_CNKD_TT40", "268", "pathStt");
			int versionConvert = 518;
			WinFnProcess.FixStt("04_2_MGT_GTGT_TT80", "268", "pathStt", versionConvert);
			WinFnProcess.FixStt("01_4_TTS_TT40", "268", "pathStt", versionConvert);
			WinFnProcess.FixStt("01_8_MGT_GTGT_TT80", "268", "pathStt", versionConvert);
			WinFnProcess.FixStt("04_1_MGT_GTGT_TT80", "268", "pathStt", versionConvert);
			WinFnProcess.FixStt("01_3_CNKD_TT40", "268", "pathStt", versionConvert);
			WinFnProcess.FixStt("01_3_TTS_TT40", "268", "pathStt", versionConvert);
			dataFileName = "03_TNDN~03_1A_TNDN~03_1B_TNDN~03_1C_TNDN~03_2A_TNDN~03_2B_TNDN~03_3A_TNDN~03_3B_TNDN~03_3C_TNDN~03_4_TNDN~03_5_TNDN~03_6_TNDN~03_7_TNDN~03_8_TNDN~03_9_TD_TNDN~03_02_TNDN~03_01_ND20_TNDN~03_02_ND20_TNDN~03_03_ND20_TNDN~03_04_ND20_TNDN";
			versionXmlOld = "231";
			versionXml = "239";
			WinFnProcess.ConverDataVersionKHBS(dataFileName, versionXmlOld, versionXml);
			dataFileName = "KHBS_03_TNDN";
			versionXmlOld = "228";
			versionXml = "239";
			WinFnProcess.ConverDataVersionKHBS(dataFileName, versionXmlOld, versionXml);
			dataFileName = "01_LPMB_TT80";
			versionXmlOld = "272";
			versionXml = "272";
			WinFnProcess.ConverData01LPMB_TT80(dataFileName, versionXmlOld, versionXml);
			dataFileName = "02_TNDN~KHBS_02_TNDN";
			versionXmlOld = "207";
			versionXml = "239";
			WinFnProcess.ConverDataVersion(dataFileName, versionXmlOld, versionXml);
			dataFileName = "02_TNDN";
			versionXmlOld = "207";
			versionXml = "239";
			WinFnProcess.ConverDataVersionKHBS(dataFileName, versionXmlOld, versionXml);
			dataFileName = "KHBS_02_TNDN";
			versionXmlOld = "214";
			versionXml = "239";
			WinFnProcess.ConverDataVersionKHBS(dataFileName, versionXmlOld, versionXml);
			dataFileName = "02_TNDN~02_1_MGT_TNDN~KHBS_02_TNDN";
			versionXmlOld = "239";
			versionXml = "250";
			WinFnProcess.ConverDataVersion(dataFileName, versionXmlOld, versionXml);
			dataFileName = "02_TNDN~02_1_MGT_TNDN";
			versionXmlOld = "239";
			versionXml = "250";
			WinFnProcess.ConverDataVersionKHBS(dataFileName, versionXmlOld, versionXml);
			dataFileName = "KHBS_02_TNDN";
			versionXmlOld = "239";
			versionXml = "250";
			WinFnProcess.ConverDataVersionKHBS(dataFileName, versionXmlOld, versionXml);
			dataFileName = "04_TNDN~KHBS_04_TNDN";
			versionXmlOld = "208";
			versionXml = "239";
			WinFnProcess.ConverDataVersion(dataFileName, versionXmlOld, versionXml);
			dataFileName = "04_TNDN~KHBS_04_TNDN";
			versionXmlOld = "208";
			versionXml = "239";
			WinFnProcess.ConverDataVersionKHBS(dataFileName, versionXmlOld, versionXml);
			dataFileName = "04_TNDN~04_1_MGT_TNDN~KHBS_04_TNDN";
			versionXmlOld = "239";
			versionXml = "250";
			WinFnProcess.ConverDataVersion(dataFileName, versionXmlOld, versionXml);
			dataFileName = "04_TNDN~04_1_MGT_TNDN~KHBS_04_TNDN";
			versionXmlOld = "239";
			versionXml = "250";
			WinFnProcess.ConverDataVersionKHBS(dataFileName, versionXmlOld, versionXml);
			dataFileName = "01_5_GTGT~01_6_GTGT~01_7_GTGT~01_1_TD_GTGT~01_2_TD_GTGT";
			versionXmlOld = "212";
			versionXml = "240";
			WinFnProcess.ConverDataVersion(dataFileName, versionXmlOld, versionXml);
			WinFnProcess.ConverDataVersionKHBS(dataFileName, versionXmlOld, versionXml);
			dataFileName = "01_GTGT";
			versionXmlOld = "212";
			versionXml = "240";
			WinFnProcess.ConverDataVersion01GTGT(dataFileName, versionXmlOld, versionXml);
			dataFileName = "KHBS_01_GTGT";
			versionXmlOld = "228";
			versionXml = "240";
			WinFnProcess.ConverDataVersionKHBS(dataFileName, versionXmlOld, versionXml);
			dataFileName = "02_GTGT";
			versionXmlOld = "212";
			versionXml = "241";
			WinFnProcess.ConverDataVersion02GTGT(dataFileName, versionXmlOld, versionXml);
			dataFileName = "KHBS_02_GTGT";
			versionXmlOld = "204";
			versionXml = "241";
			WinFnProcess.ConverDataVersionKHBS(dataFileName, versionXmlOld, versionXml);
			dataFileName = "04_GTGT";
			versionXmlOld = "212";
			versionXml = "241";
			WinFnProcess.ConverDataVersion04GTGT(dataFileName, versionXmlOld, versionXml);
			dataFileName = "KHBS_04_GTGT";
			versionXmlOld = "204";
			versionXml = "241";
			WinFnProcess.ConverDataVersionKHBS(dataFileName, versionXmlOld, versionXml);
			dataFileName = "01_DNHT";
			versionXmlOld = "222";
			versionXml = "244";
			WinFnProcess.ConverDataVersion01DNHT(dataFileName, versionXmlOld, versionXml);
			dataFileName = "05_QTT_TNCN~05_2_QTT_TNCN~05_3_QTT_TNCN";
			versionXmlOld = "216";
			versionXml = "242";
			WinFnProcess.ConverDataVersion(dataFileName, versionXmlOld, versionXml);
			WinFnProcess.ConverDataVersionKHBS(dataFileName, versionXmlOld, versionXml);
			dataFileName = "05_1_QTT_TNCN";
			versionXmlOld = "216";
			versionXml = "242";
			WinFnProcess.ConverDataVersion05_1_QTT(dataFileName, versionXmlOld, versionXml);
			dataFileName = "200_01_CDKT_HN";
			versionXmlOld = "226";
			versionXml = "243";
			WinFnProcess.ConverDataVersion200_HN_CDKT(dataFileName, versionXmlOld, versionXml);
			dataFileName = "200_02_SXKD_HN";
			versionXmlOld = "226";
			versionXml = "243";
			WinFnProcess.ConverDataVersion200_HN_HDSXKD(dataFileName, versionXmlOld, versionXml);
			dataFileName = "200_03_LCTTTT_HN~200_04_LCTTGT_HN";
			versionXmlOld = "226";
			versionXml = "243";
			WinFnProcess.ConverDataVersion(dataFileName, versionXmlOld, versionXml);
			dataFileName = "03_TNDN~03_1A_TNDN~03_1B_TNDN~03_1C_TNDN~03_2A_TNDN~03_2B_TNDN~03_3A_TNDN~03_3B_TNDN~03_3C_TNDN~03_4_TNDN~03_5_TNDN~03_6_TNDN~03_7_TNDN~03_8_TNDN~03_9_TD_TNDN~03_02_TNDN~03_01_ND20_TNDN~03_02_ND20_TNDN~03_03_ND20_TNDN~03_04_ND20_TNDN~03_01_ND132_TNDN~03_02_ND132_TNDN~03_03_ND132_TNDN~03_04_ND132_TNDN~03_1_MGT_TNDN~KHBS_03_TNDN";
			versionXmlOld = "239";
			versionXml = "245";
			WinFnProcess.ConverDataVersion(dataFileName, versionXmlOld, versionXml);
			dataFileName = "03_TNDN~03_1A_TNDN~03_1B_TNDN~03_1C_TNDN~03_2A_TNDN~03_2B_TNDN~03_3A_TNDN~03_3B_TNDN~03_3C_TNDN~03_4_TNDN~03_5_TNDN~03_6_TNDN~03_7_TNDN~03_8_TNDN~03_9_TD_TNDN~03_02_TNDN~03_01_ND20_TNDN~03_02_ND20_TNDN~03_03_ND20_TNDN~03_04_ND20_TNDN~03_01_ND132_TNDN~03_02_ND132_TNDN~03_03_ND132_TNDN~03_04_ND132_TNDN~03_1_MGT_TNDN";
			versionXmlOld = "239";
			versionXml = "245";
			WinFnProcess.ConverDataVersionKHBS(dataFileName, versionXmlOld, versionXml);
			convertDL_MGT_ND92();
			checkSuDungLoaiToKhai();
			WinFnProcess.FixDataFile_01_6_CSSX_GTGT_TT80();
			WinFnProcess.ConvertDataCommon();
			WinFnProcess.FixDataFile_03_8_TNDN_TT80();
			WinFnProcess.FixDataFile_01_DNHUY();
			WinFnProcess.FixPhuLucVTNN("01_1_VTNN_TT80", "251");
			WinFnProcess.FixPhuLucVTNN("01_2_VTNN_TT80", "251");
			WinFnProcess.FixPhuLucVTNN("01_3_VTNN_TT80", "251");
			WinFnProcess.Fix05QTTSuaDuLieu("05_1_QTT_TNCN_TT80", "279", "###Thẻ CCCD###", "###Thẻ CCCD/Số định danh cá nhân###");
			WinFnProcess.Fix05QTTSuaDuLieu("05_2_QTT_TNCN_TT80", "279", "###Thẻ CCCD###", "###Thẻ CCCD/Số định danh cá nhân###");
			WinFnProcess.Fix05QTTSuaDuLieu("05_3_QTT_TNCN_TT80", "279", "###Thẻ CCCD###", "###Thẻ CCCD/Số định danh cá nhân###");
			WinFnProcess.ReplaceDataFileDSNopThayTNCN("Ds_tien_nop_TNCN");
			try
			{
				string text = HTKKConfig.APPPATH + "\\ReplaceDataFile.xml";
				XmlDocument val = new XmlDocument();
				if (File.Exists(text))
				{
					val.Load(text);
					XmlNodeList elementsByTagName = val.GetElementsByTagName("Item");
					foreach (XmlNode item2 in elementsByTagName)
					{
						XmlNode val2 = item2;
						string value = ((XmlNode)val2.Attributes["DataFileName"]).Value;
						string value5 = ((XmlNode)val2.Attributes["VersionXml"]).Value;
						string value4 = ((XmlNode)val2.Attributes["ConvertType"]).Value;
						string tablePath = ((val2.Attributes["TablePath"] != null) ? ((XmlNode)val2.Attributes["TablePath"]).Value : "");
						string value6 = ((XmlNode)val2.Attributes["TxtOld"]).Value;
						string value7 = ((XmlNode)val2.Attributes["TxtNew"]).Value;
						int versionApp = ((val2.Attributes["VersionApp"] != null) ? Convert.ToInt32(((XmlNode)val2.Attributes["VersionApp"]).Value) : (-1));
						WinFnProcess.ReplaceDataFileCommon(value, value5, value4, versionApp, tablePath, value6, value7);
					}
				}
			}
			catch (Exception ex)
			{
				HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			}
			WinFnProcess.FixStaticToDynamic03TNDN_Tk(555, "292");
			WinFnProcess.ReplaceDataFileStaticCommon("03_TNDN_TT80", "292", 557, "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Header/ct04_ten", "Value", "");
			WinFnProcess.ReplaceDataFileStaticCommon("03_TNDN_TT80", "292", 557, "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Header/ct04_ma", "Value", "");
			WinFnProcess.ReplaceDataFileDinamicCommon("03_3A_TNDN_TT80", "292", 557, "PLuc_03_3A_TNDN/NoiDungPL03_3A", 7, "");
			WinFnProcess.ReplaceDataFileDinamicCommon("03_3A_TNDN_TT80", "292", 557, "PLuc_03_3A_TNDN/NoiDungPL03_3A", 28, "");
			WinFnProcess.ReplaceDataFileDinamicCommon("03_3B_TNDN_TT80", "292", 557, "PLuc_03_3B_TNDN/NoiDungPL03_3B", 8, "");
			WinFnProcess.ReplaceDataFileDinamicCommon("03_3B_TNDN_TT80", "292", 557, "PLuc_03_3B_TNDN/NoiDungPL03_3B", 31, "");
			WinSettings.KieuKy = strKieuKy;
			switch (strKieuKy)
			{
			case "1":
				SetLayoutToKhaiBoSung();
				break;
			case "2":
				SetLayoutToKhaiChinhThuc();
				break;
			case "3":
				SetLayoutToKhaiCoDanhMucNganhNghe();
				break;
			case "4":
				SetLayoutToKhaiCoPhuLuc();
				break;
			case "5":
				SetLayoutToKhaiKhongCoDanhMucNganhNghe();
				break;
			case "6":
				SetLayoutToKhaiKhongCoPhuLuc();
				break;
			case "7":
				SetLayoutToKhaiKyQuyTuNgayDenNgay();
				break;
			case "8":
				if (menuID == "47" || menuID == "K7")
				{
					SetLayoutToKhaiLanCTLN();
				}
				else if ("~O1~O2~O3~O4~O5~".IndexOf("~" + menuID + "~") >= 0)
				{
					SetLayoutToKhaiLanPhatSinhMst();
				}
				else
				{
					SetLayoutToKhaiLanPhatSinh();
				}
				break;
			case "9":
				SetLayoutToKhaiLanXuatBanThang();
				break;
			case "10":
				SetLayoutToKhaiNam();
				break;
			case "11":
				SetLayoutToKhaiNamTuNgayDenNgay();
				break;
			case "26":
				SetLayoutToKhaiNamTuNgayDenNgayBCTC();
				break;
			case "38":
				SetLayoutToKhaiQuyBCTC();
				break;
			case "12":
				SetLayoutToKhaiNamTuThangDenThang();
				break;
			case "25":
				SetLayoutToKhaiNamTuThangDenThangNone();
				break;
			case "24":
				SetLayoutToKhaiNamNone();
				break;
			case "13":
				SetLayoutToKhaiNamLanPhatSinh();
				break;
			case "14":
				SetLayoutToKhaiQuy();
				break;
			case "15":
				SetLayoutToKhaiQuyTuNgayDenNgay();
				break;
			case "16":
				SetLayoutToKhaiQuyTuThangDenThang();
				break;
			case "17":
				SetLayoutToKhaiThang();
				break;
			case "18":
				SetLayoutToKhaiThangLanPhatSinh();
				break;
			case "19":
				SetLayoutToKhaiThangQuy();
				break;
			case "27":
				SetLayoutToKhaiThangQuyKhongDMNN();
				break;
			case "20":
				SetLayoutToKhaiThangQuyTuNgayDenNgay();
				break;
			case "21":
				SetLayoutToKhaiThangQuyLanPhatSinh();
				break;
			case "22":
				SetLayoutToKhaiTuNgayDenNgay();
				break;
			case "28":
				SetLayoutToKhaiNamQuy_BSTuThangDenThangDmnnPl();
				break;
			case "29":
				SetLayoutToKhaiNamQuy();
				break;
			case "23":
				SetLayoutThemPhuLuc();
				break;
			case "30":
				SetLayoutToKhaiQuyBSTuThangDenThang();
				break;
			case "31":
				SetLayoutToKhaiKhongCoPhuLuc();
				break;
			case "32":
				SetLayoutToKhaiQuyBSTuThangDenThangOnly();
				break;
			case "33":
				SetLayoutToKhaiThangLanPS_BS_DMNN_PL();
				break;
			case "34":
				SetLayoutToKhaiNam_BS_TuThang_DenThang_DMNN_PL();
				break;
			case "36":
				SetLayoutToKhaiNamCoNgayLapKHBS();
				break;
			case "0":
				SetLayoutToKhaiKhongCoForm();
				if (menuID == "10" || menuID == "67")
				{
					LoadDefaultInfor();
				}
				((Control)this).Hide();
				return;
			case "35":
				SetLayoutToKhaiNamThangQuyLanPhatSinh();
				break;
			case "37":
				SetLayoutToKhaiNamKyTuNgayDenNgay();
				break;
			case "42":
				SetLayoutToQDDSThueMSTToChucCaNhanNamThangQuy();
				break;
			case "39":
				SetLayoutToKhaiDNHT();
				break;
			case "43":
				SetLayoutToKhai01HT();
				break;
			case "44":
				SetLayoutToKhai02HT();
				break;
			}
			SetDefaultValues();
			if (menuID == "66" || menuID == "52")
			{
				((Control)this).Hide();
				return;
			}
			if (menuID == "01" || menuID == "I5")
			{
				XmlDataDocument xmlDataDocument = new XmlDataDocument();
				FileStream fileStream = new FileStream(HTKKConfig.INTERFACE_INI + "\\Catalogue_DM_NNKD.xml", FileMode.Open, FileAccess.Read);
				((XmlDocument)xmlDataDocument).Load((Stream)fileStream);
				XmlNodeList elementsByTagName2 = ((XmlDocument)xmlDataDocument).GetElementsByTagName("Item");
				Dictionary<string, string> dictionary = new Dictionary<string, string>();
				foreach (XmlNode item3 in elementsByTagName2)
				{
					XmlNode val3 = item3;
					string value8 = ((XmlNode)val3.Attributes["Value"]).Value;
					string[] array = Regex.Split(value8, "###");
					if (array.Length < 3 || !array[0].Equals("01_GTGT_240"))
					{
						continue;
					}
					if (ConvertHelpers.ConvertToInt(((Control)mskYear).Text) >= 2021)
					{
						if (ConvertHelpers.ConvertToInt(array[1]) < 7)
						{
							dictionary.Add(array[1], array[2]);
						}
					}
					else if (ConvertHelpers.ConvertToInt(array[1]) >= 7)
					{
						dictionary.Add(array[1], array[2]);
					}
				}
				if (dictionary != null && dictionary.Count > 0)
				{
					cboNganhKD.DataSource = (object)new BindingSource((object)dictionary, (string)null);
					((ListControl)cboNganhKD).DisplayMember = "Value";
					((ListControl)cboNganhKD).ValueMember = "Key";
					((Control)cboNganhKD).Refresh();
				}
			}
			if ("~K2~K3~J7~E6~E7~D7~R6~X5~R9~E9~".IndexOf("~" + menuID + "~") >= 0)
			{
				XmlDataDocument xmlDataDocument = new XmlDataDocument();
				FileStream fileStream2 = new FileStream(HTKKConfig.INTERFACE_INI + "\\DM_LOAI_TIEN_DK.xml", FileMode.Open, FileAccess.Read);
				((XmlDocument)xmlDataDocument).Load((Stream)fileStream2);
				XmlNodeList elementsByTagName3 = ((XmlDocument)xmlDataDocument).GetElementsByTagName("Item");
				Dictionary<string, string> dictionary2 = new Dictionary<string, string>();
				foreach (XmlNode item4 in elementsByTagName3)
				{
					XmlNode val3 = item4;
					string value8 = ((XmlNode)val3.Attributes["Value"]).Value;
					string[] array = Regex.Split(value8, "###");
					dictionary2.Add(array[0], array[1]);
				}
				if (dictionary2 != null && dictionary2.Count > 0)
				{
					cboLoaiTien.DataSource = (object)new BindingSource((object)dictionary2, (string)null);
					((ListControl)cboLoaiTien).DisplayMember = "Value";
					((ListControl)cboLoaiTien).ValueMember = "Key";
					((Control)cboLoaiTien).Refresh();
					((ListControl)cboLoaiTien).SelectedIndex = 0;
				}
			}
			((Form)this).Location = new Point((Screen.PrimaryScreen.Bounds.Width - ((Form)this).Size.Width) / 2, (Screen.PrimaryScreen.Bounds.Height - ((Form)this).Size.Height) / 2);
			NopTheoTinhCQT();
			TruongHopQuyetToan();
			getTTDuAnDT();
			if (((Control)pnVT2_PhuLucKeKhai).Visible)
			{
				if (menuID == "Q9")
				{
					((Control)pnTinhCQT).Enabled = false;
				}
				LoadPhuLucSpread();
				if (menuID == "Y7")
				{
					bool enabled = (bool)axfpsPhuLuc.ActiveSheet.Cells[1, 0].Value;
					((Control)pnTinhCQT).Enabled = enabled;
					((Control)pnCQTCuc).Enabled = false;
					((Control)pnCCT).Enabled = false;
					if ((bool)axfpsPhuLuc.ActiveSheet.Cells[1, 0].Value)
					{
						((Control)chkTKQuy).Enabled = false;
						chkTKQuy.Checked = false;
					}
					else
					{
						((Control)chkTKQuy).Enabled = true;
					}
					if ((bool)axfpsPhuLuc.ActiveSheet.Cells[0, 0].Value)
					{
						((Control)chkTKLanPS).Enabled = false;
						chkTKLanPS.Checked = false;
					}
					else
					{
						((Control)chkTKLanPS).Enabled = true;
					}
				}
			}
			if (menuID == "Q5" || menuID == "Q6")
			{
				((Control)pnVT2_PhuLucKeKhai).Visible = false;
			}
			if (menuID == "X1" || menuID == "X2")
			{
				((Control)chkTKKy).Text = "Bán niên độ";
				((Control)chkTKQuy).Text = "Quý";
				Dictionary<string, string> dictionary3 = new Dictionary<string, string>();
				dictionary3.Add("1", "6 tháng đầu kỳ");
				dictionary3.Add("2", "6 tháng cuối kỳ");
				bindingDataToComboBox(cbbKy, dictionary3);
				((Control)cbbKy).Width = 100;
				((Control)pnKy).Width = 155;
				((Control)pnVT2_PhuLucKeKhai).Visible = true;
			}
			Resize();
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void LoadPhuLucSpread()
	{
		try
		{
			GetPeriodInfo(WinSettings.TTPeriodInfo);
			strFileNameExtend = WinFnProcess.GetDataFileNameExtend(WinSettings.TTPeriodInfo);
			WinSettings.DicSheet = WinFnProcess.GetPhuLucDaKeKhai(((Control)pnVT2_PhuLucKeKhai).Visible, strFileNameExtend);
			WinSpreadProcess.BindDataToPhuLucSpread(axfpsPhuLuc.Sheets[0], elemList, chkChonPhuLucKeKhaiAll);
			WinDoWorkEvent.event_Leave(axfpsPhuLuc, listCheckBox, listMaskedTextBox, listComboBox, listRadioButton, "LoadPhuLucSpread");
			if (menuID.Equals("68") || menuID.Equals("14") || menuID.Equals("13") || menuID.Equals("18") || menuID.Equals("62") || menuID.Equals("65") || menuID == "Q9" || menuID == "Y7")
			{
				((Control)chkChonPhuLucKeKhaiAll).Enabled = false;
			}
			int row = 20;
			if (menuID == "N8")
			{
				if (ConvertHelpers.ConvertToInt(((Control)mskYear).Text) != 2021)
				{
					axfpsPhuLuc.Sheets[0].Cells[row, 0].Value = false;
					axfpsPhuLuc.Sheets[0].Cells[row, 0].Locked = true;
				}
				else
				{
					axfpsPhuLuc.Sheets[0].Cells[row, 0].Locked = false;
				}
			}
			else if (menuID == "Q9")
			{
				if (chkTKQuy.Checked)
				{
					axfpsPhuLuc.Sheets[0].Cells[1, 0].Locked = true;
					axfpsPhuLuc.Sheets[0].Cells[1, 0].Value = false;
					axfpsPhuLuc.Sheets[0].Cells[0, 0].Value = true;
				}
				if (((Control)pnTinhCQT).Enabled)
				{
					axfpsPhuLuc.Sheets[0].Cells[1, 0].Value = ((Control)pnTinhCQT).Enabled;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	private void SetLayoutToKhaiBoSung()
	{
		ShowDetailpnVT1_1CheckBox(chkTKhaiLanXBShow: false, chkTkhaiThangShow: true, chkTKKyShow: false, chkTKQuyShow: true, ckNamShow: false, chkTuThangDenThangShow: false, chkTKLanPSShow: false, chkTuQuyDenQuyShow: false);
		chkTkhaiThang.Checked = true;
		ResetpnVT1_2NgayThangQuyNam();
		ShowDetailpnVT1_3TuNgayDenNgay(pnLanXuatBanShow: false, pnTuNgayShow: false, pnDenNgayShow: false, pnTuThangShow: false, pnDenThangShow: false);
		ShowDetailpnVT1_4LanDauBoXung(pnQuyetToan_LienTuc: false, pnRDLanDauBoXungShow: true, pnNgayLapShow: false);
		ShowDetailpnVT1_5DanhMucNghe(pnVT1_5DanhMucNgheShow: false);
		ShowDetailpnVT2_PhuLucKeKhai(pnVT2_PhuLucKeKhaiShow: false);
		Resize();
	}

	private void SetLayoutToKhaiChinhThuc()
	{
		ShowDetailpnVT1_1CheckBox(chkTKhaiLanXBShow: false, chkTkhaiThangShow: true, chkTKKyShow: false, chkTKQuyShow: true, ckNamShow: false, chkTuThangDenThangShow: false, chkTKLanPSShow: false, chkTuQuyDenQuyShow: false);
		chkTkhaiThang.Checked = true;
		ResetpnVT1_2NgayThangQuyNam();
		ShowDetailpnVT1_3TuNgayDenNgay(pnLanXuatBanShow: false, pnTuNgayShow: false, pnDenNgayShow: false, pnTuThangShow: false, pnDenThangShow: false);
		ShowDetailpnVT1_4LanDauBoXung(pnQuyetToan_LienTuc: false, pnRDLanDauBoXungShow: true, pnNgayLapShow: false);
		ShowDetailpnVT1_5DanhMucNghe(pnVT1_5DanhMucNgheShow: false);
		ShowDetailpnVT2_PhuLucKeKhai(pnVT2_PhuLucKeKhaiShow: false);
		Resize();
	}

	private void SetLayoutToKhaiCoDanhMucNganhNghe()
	{
		ShowDetailpnVT1_1CheckBox(chkTKhaiLanXBShow: false, chkTkhaiThangShow: false, chkTKKyShow: false, chkTKQuyShow: false, ckNamShow: false, chkTuThangDenThangShow: false, chkTKLanPSShow: false, chkTuQuyDenQuyShow: false);
		ShowDetailpnVT1_2NgayThangQuyNam(pnNgayShow: false, pnThangShow: false, pnKyShow: false, pnQuyShow: true, pnNamShow: true);
		ShowDetailpnVT1_3TuNgayDenNgay(pnLanXuatBanShow: false, pnTuNgayShow: false, pnDenNgayShow: false, pnTuThangShow: false, pnDenThangShow: false);
		ShowDetailpnVT1_4LanDauBoXung(pnQuyetToan_LienTuc: false, pnRDLanDauBoXungShow: true, pnNgayLapShow: false);
		ShowDetailpnVT1_5DanhMucNghe(pnVT1_5DanhMucNgheShow: true);
		ShowDetailpnVT2_PhuLucKeKhai(pnVT2_PhuLucKeKhaiShow: true);
		Resize();
	}

	private void SetLayoutToKhaiCoPhuLuc()
	{
		ShowDetailpnVT1_1CheckBox(chkTKhaiLanXBShow: false, chkTkhaiThangShow: true, chkTKKyShow: false, chkTKQuyShow: true, ckNamShow: false, chkTuThangDenThangShow: false, chkTKLanPSShow: false, chkTuQuyDenQuyShow: false);
		ShowDetailpnVT1_3TuNgayDenNgay(pnLanXuatBanShow: false, pnTuNgayShow: false, pnDenNgayShow: false, pnTuThangShow: false, pnDenThangShow: false);
		ShowDetailpnVT1_4LanDauBoXung(pnQuyetToan_LienTuc: false, pnRDLanDauBoXungShow: true, pnNgayLapShow: false);
		ShowDetailpnVT1_5DanhMucNghe(pnVT1_5DanhMucNgheShow: true);
		ShowDetailpnVT2_PhuLucKeKhai(pnVT2_PhuLucKeKhaiShow: true);
		chkTkhaiThang.Checked = true;
	}

	private void SetLayoutToKhaiKhongCoDanhMucNganhNghe()
	{
		ShowDetailpnVT1_1CheckBox(chkTKhaiLanXBShow: false, chkTkhaiThangShow: true, chkTKKyShow: false, chkTKQuyShow: true, ckNamShow: false, chkTuThangDenThangShow: false, chkTKLanPSShow: false, chkTuQuyDenQuyShow: false);
		chkTkhaiThang.Checked = true;
		ResetpnVT1_2NgayThangQuyNam();
		ShowDetailpnVT1_3TuNgayDenNgay(pnLanXuatBanShow: false, pnTuNgayShow: false, pnDenNgayShow: false, pnTuThangShow: false, pnDenThangShow: false);
		ShowDetailpnVT1_4LanDauBoXung(pnQuyetToan_LienTuc: false, pnRDLanDauBoXungShow: true, pnNgayLapShow: false);
		ShowDetailpnVT1_5DanhMucNghe(pnVT1_5DanhMucNgheShow: false);
		ShowDetailpnVT2_PhuLucKeKhai(pnVT2_PhuLucKeKhaiShow: true);
		Resize();
	}

	private void SetLayoutToKhaiKhongCoPhuLuc()
	{
		ShowDetailpnVT1_1CheckBox(chkTKhaiLanXBShow: false, chkTkhaiThangShow: true, chkTKKyShow: false, chkTKQuyShow: true, ckNamShow: false, chkTuThangDenThangShow: false, chkTKLanPSShow: false, chkTuQuyDenQuyShow: false);
		chkTkhaiThang.Checked = true;
		ResetpnVT1_2NgayThangQuyNam();
		if (menuID == "02" || menuID == "D2")
		{
			ShowDetailpnMaHSDADT(pnMaHSDADT: true);
			if (menuID == "02")
			{
				((Control)txtMaHSDNHT_DADT).Visible = true;
				((Control)cbMaHSDNHT_DADT).Visible = false;
			}
		}
		ShowDetailpnVT1_3TuNgayDenNgay(pnLanXuatBanShow: false, pnTuNgayShow: false, pnDenNgayShow: false, pnTuThangShow: false, pnDenThangShow: false);
		ShowDetailpnVT1_4LanDauBoXung(pnQuyetToan_LienTuc: false, pnRDLanDauBoXungShow: true, pnNgayLapShow: false);
		ShowDetailpnVT1_5DanhMucNghe(pnVT1_5DanhMucNgheShow: false);
		ShowDetailpnVT2_PhuLucKeKhai(pnVT2_PhuLucKeKhaiShow: false);
		Resize();
	}

	private void SetLayoutToKhaiKyQuyTuNgayDenNgay()
	{
		ShowDetailpnVT1_1CheckBox(chkTKhaiLanXBShow: false, chkTkhaiThangShow: false, chkTKKyShow: true, chkTKQuyShow: true, ckNamShow: false, chkTuThangDenThangShow: false, chkTKLanPSShow: false, chkTuQuyDenQuyShow: false);
		if (menuID == "13")
		{
			chkTKKy.Checked = true;
		}
		else
		{
			chkTKQuy.Checked = true;
		}
		ResetpnVT1_2NgayThangQuyNam();
		ShowDetailpnVT1_3TuNgayDenNgay(pnLanXuatBanShow: false, pnTuNgayShow: true, pnDenNgayShow: true, pnTuThangShow: false, pnDenThangShow: false);
		ShowDetailpnVT1_4LanDauBoXung(pnQuyetToan_LienTuc: false, pnRDLanDauBoXungShow: false, pnNgayLapShow: false);
		ShowDetailpnVT1_5DanhMucNghe(pnVT1_5DanhMucNgheShow: false);
		ShowDetailpnVT2_PhuLucKeKhai(pnVT2_PhuLucKeKhaiShow: false);
		Resize();
	}

	private void SetLayoutToKhaiLanPhatSinh()
	{
		ShowDetailpnVT1_1CheckBox(chkTKhaiLanXBShow: false, chkTkhaiThangShow: false, chkTKKyShow: false, chkTKQuyShow: false, ckNamShow: false, chkTuThangDenThangShow: false, chkTKLanPSShow: false, chkTuQuyDenQuyShow: false);
		ShowDetailpnVT1_2NgayThangQuyNam(pnNgayShow: true, pnThangShow: true, pnKyShow: false, pnQuyShow: false, pnNamShow: true);
		if ("~K2~K3~J7~E6~E7~D7~R6~X5~R9~E9~".IndexOf("~" + menuID + "~") >= 0)
		{
			ShowDetailpnLoaiTien(pnLoaiTien: true);
		}
		if (menuID == "T6" || menuID == "U2" || menuID == "Y5")
		{
			ShowDetailpnVT1_3TuNgayDenNgay(pnLanXuatBanShow: false, pnTuNgayShow: true, pnDenNgayShow: false, pnTuThangShow: false, pnDenThangShow: false);
			((Control)lblTuNgay).Text = "Ngày hợp đồng";
			((Control)pnTuNgay).Width = 300;
			((Control)mskTuNgay).Location = new Point(((Control)mskTuNgay).Location.X + 40, ((Control)mskTuNgay).Location.Y);
		}
		else
		{
			ShowDetailpnVT1_3TuNgayDenNgay(pnLanXuatBanShow: false, pnTuNgayShow: false, pnDenNgayShow: false, pnTuThangShow: false, pnDenThangShow: false);
		}
		string text = "~M1~M2~G3~";
		if (menuID == "07" || menuID == "27" || menuID == "91" || menuID == "64" || menuID == "B6" || menuID == "B7" || menuID == "G1" || menuID == "M5" || menuID == "N6" || menuID == "M6" || menuID == "O8" || menuID == "B2" || menuID == "O9" || menuID == "N2" || text.IndexOf("~" + menuID + "~") >= 0 || menuID == "P1" || menuID == "P6" || menuID == "X7" || menuID.Equals("Y9"))
		{
			ShowDetailpnVT1_4LanDauBoXung(pnQuyetToan_LienTuc: false, pnRDLanDauBoXungShow: false, pnNgayLapShow: false);
		}
		else
		{
			ShowDetailpnVT1_4LanDauBoXung(pnQuyetToan_LienTuc: false, pnRDLanDauBoXungShow: true, pnNgayLapShow: false);
		}
		if (menuID == "M8" || menuID == "M9")
		{
			ShowDetailpnVT1_5DanhMucNghe(pnVT1_5DanhMucNgheShow: true);
		}
		else
		{
			ShowDetailpnVT1_5DanhMucNghe(pnVT1_5DanhMucNgheShow: false);
		}
		if (menuID == "73" || menuID == "N2")
		{
			ShowDetailpnVT2_PhuLucKeKhai(pnVT2_PhuLucKeKhaiShow: true);
		}
		else
		{
			ShowDetailpnVT2_PhuLucKeKhai(pnVT2_PhuLucKeKhaiShow: false);
		}
		if (menuID == "G4" || menuID == "N5" || menuID == "H5" || menuID == "O7" || menuID == "P5" || menuID == "U6")
		{
			ShowDetailpnVT1_4LanDauBoXung(pnQuyetToan_LienTuc: false, pnRDLanDauBoXungShow: true, pnNgayLapShow: false);
			((Control)OptChinhthuc).Text = "Lần đầu";
			((Control)OptBosung).Text = "Thay thế";
		}
		if ("U8~".Contains(WinSettings.MenuID))
		{
			((Control)OptChinhthuc).Text = "Lần đầu";
			((Control)OptBosung).Text = "Thay thế";
		}
		if (menuID == "P7")
		{
			((Control)OptChinhthuc).Text = "Thông báo lần đầu";
			((Control)OptBosung).Text = "Thông báo thay đổi";
			((Control)pnNgayLap).Visible = false;
		}
		if (menuID == "T6" || menuID == "U2" || menuID == "Y5")
		{
			ShowDetailpnVT1_5DanhMucNghe(pnVT1_5DanhMucNgheShow: true);
			((Control)lblNganhKD).Text = "Hình thức kê khai";
		}
		if (menuID == "U1")
		{
			ShowDetailpnVT1_5DanhMucNghe(pnVT1_5DanhMucNgheShow: true);
			((Control)lblNganhKD).Text = "Hình thức kê khai";
			((Control)pnChuyenNhuong).Visible = true;
			((Control)chkToChucCaNhanChuyenNhuong).Text = "Bất động sản";
			((Control)chkDoanhNghiepChuyenNhuong).Text = "Tài sản khác";
		}
		if (menuID == "Y3")
		{
			ShowDetailpnVT1_5DanhMucNghe(pnVT1_5DanhMucNgheShow: true);
			((Control)lblNganhKD).Text = "Hình thức kê khai";
			((Control)pnChuyenNhuong).Visible = true;
			((Control)chkToChucCaNhanChuyenNhuong).Text = "Tổ chức áp dụng phương pháp tỷ lệ trên doanh thu theo quy định của pháp luật về thuế TNDN có hoạt động chuyển nhượng bất động sản";
			((Control)chkToChucCaNhanChuyenNhuong).Height = 100;
			((Control)chkDoanhNghiepChuyenNhuong).Text = "Cơ quan, tổ chức bán hoặc bán đấu giá bất động sản, tài sản khác là tài sản bảo đảm thi hành án, xử lý nợ";
		}
		if (menuID == "Z1")
		{
			ShowDetailpnVT1_5DanhMucNghe(pnVT1_5DanhMucNgheShow: true);
			((Control)lblNganhKD).Text = "Loại thu nhập";
		}
		if (menuID == "P1" || menuID == "P6" || menuID == "X7" || menuID.Equals("Y9"))
		{
			((Control)pnVT1_3TuNgayDenNgay).Visible = true;
			((Control)pnLanXuatBan).Visible = true;
			((Control)lblLanXuat).Text = "Lần gửi";
			((Control)mskDay).Location = new Point(((Control)mskDay).Location.X + 17, ((Control)mskDay).Location.Y);
			((Control)lblNgay).Text = "Ngày lập";
			((Control)pnThang).Location = new Point(((Control)pnThang).Location.X + 2, ((Control)pnThang).Location.Y);
		}
		if (menuID == "P1" || menuID == "P6" || menuID == "X7" || menuID.Equals("Y9"))
		{
			ShowDetailpnMst(pnMst: true);
			((Control)lblMst).Text = "Số chứng từ nộp NSNN";
			((Control)lblMst).AutoSize = true;
			((Control)txtMst).Location = new Point(((Control)txtMst).Location.X - 110, ((Control)txtMst).Location.Y);
			TextBox obj = txtMst;
			((Control)obj).Width = ((Control)obj).Width + 100;
		}
		Resize();
	}

	private void SetLayoutToKhaiLanPhatSinhMst()
	{
		ShowDetailpnVT1_1CheckBox(chkTKhaiLanXBShow: false, chkTkhaiThangShow: false, chkTKKyShow: false, chkTKQuyShow: false, ckNamShow: false, chkTuThangDenThangShow: false, chkTKLanPSShow: false, chkTuQuyDenQuyShow: false);
		ShowDetailpnVT1_2NgayThangQuyNam(pnNgayShow: true, pnThangShow: true, pnKyShow: false, pnQuyShow: false, pnNamShow: true);
		ShowDetailpnVT1_3TuNgayDenNgay(pnLanXuatBanShow: false, pnTuNgayShow: false, pnDenNgayShow: false, pnTuThangShow: false, pnDenThangShow: false);
		ShowDetailpnVT1_4LanDauBoXung(pnQuyetToan_LienTuc: false, pnRDLanDauBoXungShow: true, pnNgayLapShow: false);
		ShowDetailpnVT1_5DanhMucNghe(pnVT1_5DanhMucNgheShow: false);
		ShowDetailpnMst(pnMst: true);
		ShowDetailpnVT2_PhuLucKeKhai(pnVT2_PhuLucKeKhaiShow: false);
		Resize();
	}

	private void SetLayoutToKhaiLanXuatBanThang()
	{
		ShowDetailpnVT1_1CheckBox(chkTKhaiLanXBShow: true, chkTkhaiThangShow: true, chkTKKyShow: false, chkTKQuyShow: false, ckNamShow: false, chkTuThangDenThangShow: false, chkTKLanPSShow: false, chkTuQuyDenQuyShow: false);
		ShowDetailpnVT1_3TuNgayDenNgay(pnLanXuatBanShow: false, pnTuNgayShow: false, pnDenNgayShow: false, pnTuThangShow: false, pnDenThangShow: false);
		ShowDetailpnVT1_4LanDauBoXung(pnQuyetToan_LienTuc: false, pnRDLanDauBoXungShow: true, pnNgayLapShow: false);
		ShowDetailpnVT1_5DanhMucNghe(pnVT1_5DanhMucNgheShow: true);
		if ("~K2~K3~J7~E6~E7~D7~R6~X5~R9~E9~".IndexOf("~" + menuID + "~") >= 0)
		{
			ShowDetailpnLoaiTien(pnLoaiTien: true);
		}
		if (menuID == "R6")
		{
			((Control)pnChuyenNhuong).Visible = true;
			((Control)chkToChucCaNhanChuyenNhuong).Text = "Hợp đồng dầu khí (PSC)";
			((Control)chkDoanhNghiepChuyenNhuong).Text = "Hiệp đinh liên chính phủ (VSP)";
		}
		ShowDetailpnVT2_PhuLucKeKhai(pnVT2_PhuLucKeKhaiShow: true);
		chkTKhaiLanXB.Checked = true;
	}

	private void SetLayoutToKhaiLanCTLN()
	{
		ShowDetailpnVT1_1CheckBox(chkTKhaiLanXBShow: false, chkTkhaiThangShow: false, chkTKKyShow: false, chkTKQuyShow: false, ckNamShow: false, chkTuThangDenThangShow: false, chkTKLanPSShow: false, chkTuQuyDenQuyShow: false);
		ShowDetailpnVT1_2NgayThangQuyNam(pnNgayShow: true, pnThangShow: true, pnKyShow: false, pnQuyShow: false, pnNamShow: true);
		ShowDetailpnVT1_3TuNgayDenNgay(pnLanXuatBanShow: false, pnTuNgayShow: false, pnDenNgayShow: false, pnTuThangShow: false, pnDenThangShow: false);
		ShowDetailpnVT1_4LanDauBoXung(pnQuyetToan_LienTuc: true, pnRDLanDauBoXungShow: true, pnNgayLapShow: false);
		ShowDetailpnVT1_5DanhMucNghe(pnVT1_5DanhMucNgheShow: false);
		ShowDetailpnVT2_PhuLucKeKhai(pnVT2_PhuLucKeKhaiShow: false);
		Resize();
	}

	private void SetLayoutToKhaiNam()
	{
		ShowDetailpnVT1_1CheckBox(chkTKhaiLanXBShow: false, chkTkhaiThangShow: false, chkTKKyShow: false, chkTKQuyShow: false, ckNamShow: false, chkTuThangDenThangShow: false, chkTKLanPSShow: false, chkTuQuyDenQuyShow: false);
		ShowDetailpnVT1_2NgayThangQuyNam(pnNgayShow: false, pnThangShow: false, pnKyShow: false, pnQuyShow: false, pnNamShow: true);
		ShowDetailpnVT1_3TuNgayDenNgay(pnLanXuatBanShow: false, pnTuNgayShow: false, pnDenNgayShow: false, pnTuThangShow: false, pnDenThangShow: false);
		ShowDetailpnVT1_4LanDauBoXung(pnQuyetToan_LienTuc: false, pnRDLanDauBoXungShow: false, pnNgayLapShow: false);
		ShowDetailpnVT1_5DanhMucNghe(pnVT1_5DanhMucNgheShow: false);
		ShowDetailpnVT2_PhuLucKeKhai(pnVT2_PhuLucKeKhaiShow: true);
		if (menuID == "95")
		{
			ShowDetailpnVT1_4LanDauBoXung(pnQuyetToan_LienTuc: false, pnRDLanDauBoXungShow: true, pnNgayLapShow: false);
			((Control)OptChinhthuc).Visible = false;
			((Control)OptBosung).Visible = false;
			((Control)pnNgayLap).Visible = false;
			((Control)pnLan).Visible = true;
			ShowDetailpnVT2_PhuLucKeKhai(pnVT2_PhuLucKeKhaiShow: false);
		}
		Resize();
	}

	private void SetLayoutToKhaiNamTuNgayDenNgay()
	{
		ShowDetailpnVT1_1CheckBox(chkTKhaiLanXBShow: false, chkTkhaiThangShow: false, chkTKKyShow: false, chkTKQuyShow: false, ckNamShow: false, chkTuThangDenThangShow: false, chkTKLanPSShow: false, chkTuQuyDenQuyShow: false);
		ShowDetailpnVT1_2NgayThangQuyNam(pnNgayShow: false, pnThangShow: false, pnKyShow: false, pnQuyShow: false, pnNamShow: true);
		ShowDetailpnVT1_3TuNgayDenNgay(pnLanXuatBanShow: false, pnTuNgayShow: true, pnDenNgayShow: true, pnTuThangShow: false, pnDenThangShow: false);
		ShowDetailpnVT1_4LanDauBoXung(pnQuyetToan_LienTuc: false, pnRDLanDauBoXungShow: true, pnNgayLapShow: false);
		if (menuID != "83" && menuID != "C6" && menuID != "C7" && menuID != "Q7")
		{
			if (menuID == "N8" || menuID == "J3" || menuID == "H4" || menuID == "Q5" || menuID == "Q6")
			{
				ShowDetailpnVT1_5DanhMucNghe(pnVT1_5DanhMucNgheShow: false);
			}
			else
			{
				ShowDetailpnVT1_5DanhMucNghe(pnVT1_5DanhMucNgheShow: true);
			}
			ShowDetailpnVT2_PhuLucKeKhai(pnVT2_PhuLucKeKhaiShow: true);
		}
		else
		{
			ShowDetailpnVT1_5DanhMucNghe(pnVT1_5DanhMucNgheShow: false);
			ShowDetailpnVT2_PhuLucKeKhai(pnVT2_PhuLucKeKhaiShow: false);
		}
		if (menuID == "Q7")
		{
			ShowDetailpnVT1_4LanDauBoXung(pnQuyetToan_LienTuc: false, pnRDLanDauBoXungShow: false, pnNgayLapShow: false);
		}
		Resize();
	}

	private void SetLayoutToKhaiNamTuNgayDenNgayBCTC()
	{
		ShowDetailpnVT1_1CheckBox(chkTKhaiLanXBShow: false, chkTkhaiThangShow: false, chkTKKyShow: false, chkTKQuyShow: false, ckNamShow: false, chkTuThangDenThangShow: false, chkTKLanPSShow: false, chkTuQuyDenQuyShow: false);
		ShowDetailpnVT1_2NgayThangQuyNam(pnNgayShow: false, pnThangShow: false, pnKyShow: false, pnQuyShow: false, pnNamShow: true);
		ShowDetailpnVT1_3TuNgayDenNgay(pnLanXuatBanShow: false, pnTuNgayShow: true, pnDenNgayShow: true, pnTuThangShow: false, pnDenThangShow: false);
		if ("78,58".Contains(menuID) || "U9".Equals(menuID))
		{
			ShowDetailpnVT1_4LanDauBoXung(pnQuyetToan_LienTuc: true, pnRDLanDauBoXungShow: false, pnNgayLapShow: false);
			ShowDetailpnVT1_5DanhMucNghe(pnVT1_5DanhMucNgheShow: false);
		}
		else
		{
			ShowDetailpnVT1_4LanDauBoXung(pnQuyetToan_LienTuc: false, pnRDLanDauBoXungShow: false, pnNgayLapShow: false);
			ShowDetailpnVT1_5DanhMucNghe(pnVT1_5DanhMucNgheShow: false);
		}
		ShowDetailpnVT2_PhuLucKeKhai(pnVT2_PhuLucKeKhaiShow: true);
		Resize();
	}

	private void SetLayoutToKhaiQuyBCTC()
	{
		ShowDetailpnVT1_1CheckBox(chkTKhaiLanXBShow: false, chkTkhaiThangShow: false, chkTKKyShow: false, chkTKQuyShow: false, ckNamShow: false, chkTuThangDenThangShow: false, chkTKLanPSShow: false, chkTuQuyDenQuyShow: false);
		ShowDetailpnVT1_2NgayThangQuyNam(pnNgayShow: false, pnThangShow: false, pnKyShow: false, pnQuyShow: true, pnNamShow: true);
		ShowDetailpnVT1_3TuNgayDenNgay(pnLanXuatBanShow: false, pnTuNgayShow: true, pnDenNgayShow: true, pnTuThangShow: false, pnDenThangShow: false);
		ShowDetailpnVT1_4LanDauBoXung(pnQuyetToan_LienTuc: false, pnRDLanDauBoXungShow: false, pnNgayLapShow: false);
		ShowDetailpnVT1_5DanhMucNghe(pnVT1_5DanhMucNgheShow: false);
		ShowDetailpnVT2_PhuLucKeKhai(pnVT2_PhuLucKeKhaiShow: true);
		Resize();
	}

	private void SetLayoutToKhaiDNHT()
	{
		ShowDetailpnVT1_1CheckBox(chkTKhaiLanXBShow: false, chkTkhaiThangShow: false, chkTKKyShow: false, chkTKQuyShow: false, ckNamShow: false, chkTuThangDenThangShow: true, chkTKLanPSShow: false, chkTuQuyDenQuyShow: true);
		ShowDetailpnVT1_2NgayThangQuyNam(pnNgayShow: false, pnThangShow: false, pnKyShow: false, pnQuyShow: false, pnNamShow: false);
		ShowDetailpnVT1_3TuNgayDenNgay(pnLanXuatBanShow: false, pnTuNgayShow: false, pnDenNgayShow: false, pnTuThangShow: true, pnDenThangShow: true);
		ShowDetailpnVT1_4LanDauBoXung(pnQuyetToan_LienTuc: false, pnRDLanDauBoXungShow: false, pnNgayLapShow: false);
		ShowDetailpnVT1_5DanhMucNghe(pnVT1_5DanhMucNgheShow: false);
		if (menuID == "A4" || menuID == "99")
		{
			FormProcess.BeginControlUpdate((Control)(object)WinSettings.myHTKKSystems);
			((Control)pnMaHSDNHT).Visible = true;
			((Control)pnVT1_5DanhMucNghe).Visible = true;
			XMLMessage xMLMessage = new XMLMessage("../Project/Message.xml");
			((Control)lblNganhKD).Text = xMLMessage.GetMessage("0522");
			FormProcess.EndControlUpdate((Control)(object)WinSettings.myHTKKSystems);
		}
		ShowDetailpnVT2_PhuLucKeKhai(pnVT2_PhuLucKeKhaiShow: true);
		chkTuThangDenThang.Checked = true;
		Resize();
	}

	private void SetLayoutToKhai01HT()
	{
		//IL_0093: Unknown result type (might be due to invalid IL or missing references)
		//IL_009a: Expected O, but got Unknown
		//IL_0172: Unknown result type (might be due to invalid IL or missing references)
		//IL_017c: Expected O, but got Unknown
		ShowDetailpnVT1_1CheckBox(chkTKhaiLanXBShow: false, chkTkhaiThangShow: false, chkTKKyShow: false, chkTKQuyShow: false, ckNamShow: false, chkTuThangDenThangShow: false, chkTKLanPSShow: false, chkTuQuyDenQuyShow: false);
		ShowDetailpnVT1_2NgayThangQuyNam(pnNgayShow: false, pnThangShow: false, pnKyShow: false, pnQuyShow: false, pnNamShow: false);
		ShowDetailpnVT1_3TuNgayDenNgay(pnLanXuatBanShow: false, pnTuNgayShow: false, pnDenNgayShow: false, pnTuThangShow: true, pnDenThangShow: true);
		ShowDetailpnVT1_4LanDauBoXung(pnQuyetToan_LienTuc: false, pnRDLanDauBoXungShow: false, pnNgayLapShow: false);
		ShowDetailpnVT1_5DanhMucNghe(pnVT1_5DanhMucNgheShow: false);
		((Control)pnLoaiKyHoan).Visible = true;
		XmlDataDocument xmlDataDocument = new XmlDataDocument();
		FileStream fileStream = new FileStream(HTKKConfig.INTERFACE_INI + "\\DM_KIEU_KY_HT.xml", FileMode.Open, FileAccess.Read);
		((XmlDocument)xmlDataDocument).Load((Stream)fileStream);
		XmlNodeList elementsByTagName = ((XmlDocument)xmlDataDocument).GetElementsByTagName("Item");
		Dictionary<string, string> dictionary = new Dictionary<string, string>();
		foreach (XmlNode item in elementsByTagName)
		{
			XmlNode val = item;
			string value = ((XmlNode)val.Attributes["Value"]).Value;
			string[] array = Regex.Split(value, "###");
			if (array.Length >= 2)
			{
				if (menuID != "H6")
				{
					dictionary.Add(array[0], array[1]);
				}
				else if (value.StartsWith("1"))
				{
					dictionary.Add(array[0], array[1]);
				}
			}
		}
		if (dictionary != null && dictionary.Count > 0)
		{
			cboKieuKy.DataSource = (object)new BindingSource((object)dictionary, (string)null);
			((ListControl)cboKieuKy).DisplayMember = "Value";
			((ListControl)cboKieuKy).ValueMember = "Key";
			((Control)cboKieuKy).Refresh();
		}
		if (menuID == "A4" || menuID == "99")
		{
			FormProcess.BeginControlUpdate((Control)(object)WinSettings.myHTKKSystems);
			((Control)pnMaHSDNHT).Visible = true;
			((Control)pnVT1_5DanhMucNghe).Visible = true;
			XMLMessage xMLMessage = new XMLMessage("../Project/Message.xml");
			((Control)lblNganhKD).Text = xMLMessage.GetMessage("0522");
			FormProcess.EndControlUpdate((Control)(object)WinSettings.myHTKKSystems);
		}
		if (menuID == "99")
		{
			ShowDetailpnVT2_PhuLucKeKhai(pnVT2_PhuLucKeKhaiShow: true);
		}
		else
		{
			ShowDetailpnVT2_PhuLucKeKhai(pnVT2_PhuLucKeKhaiShow: false);
		}
		Resize();
	}

	private void SetLayoutToKhai02HT()
	{
		//IL_0093: Unknown result type (might be due to invalid IL or missing references)
		//IL_009a: Expected O, but got Unknown
		//IL_0188: Unknown result type (might be due to invalid IL or missing references)
		//IL_0192: Expected O, but got Unknown
		ShowDetailpnVT1_1CheckBox(chkTKhaiLanXBShow: false, chkTkhaiThangShow: false, chkTKKyShow: false, chkTKQuyShow: false, ckNamShow: false, chkTuThangDenThangShow: false, chkTKLanPSShow: false, chkTuQuyDenQuyShow: false);
		ShowDetailpnVT1_2NgayThangQuyNam(pnNgayShow: false, pnThangShow: false, pnKyShow: false, pnQuyShow: false, pnNamShow: false);
		ShowDetailpnVT1_3TuNgayDenNgay(pnLanXuatBanShow: false, pnTuNgayShow: false, pnDenNgayShow: false, pnTuThangShow: true, pnDenThangShow: true);
		ShowDetailpnVT1_4LanDauBoXung(pnQuyetToan_LienTuc: false, pnRDLanDauBoXungShow: false, pnNgayLapShow: false);
		ShowDetailpnVT1_5DanhMucNghe(pnVT1_5DanhMucNgheShow: false);
		((Control)pnLoaiKyHoan).Visible = true;
		XmlDataDocument xmlDataDocument = new XmlDataDocument();
		FileStream fileStream = new FileStream(HTKKConfig.INTERFACE_INI + "\\DM_KIEU_KY_HT_02.xml", FileMode.Open, FileAccess.Read);
		((XmlDocument)xmlDataDocument).Load((Stream)fileStream);
		XmlNodeList elementsByTagName = ((XmlDocument)xmlDataDocument).GetElementsByTagName("Item");
		Dictionary<string, string> dictionary = new Dictionary<string, string>();
		foreach (XmlNode item in elementsByTagName)
		{
			XmlNode val = item;
			string value = ((XmlNode)val.Attributes["Value"]).Value;
			string[] array = Regex.Split(value, "###");
			if (array.Length >= 2)
			{
				if (menuID != "H6")
				{
					dictionary.Add(array[0], array[1]);
				}
				else if (value.StartsWith("1"))
				{
					dictionary.Add(array[0], array[1]);
				}
			}
		}
		FormProcess.BeginControlUpdate((Control)(object)WinSettings.myHTKKSystems);
		FormProcess.EndControlUpdate((Control)(object)WinSettings.myHTKKSystems);
		if (dictionary != null && dictionary.Count > 0)
		{
			cboKieuKy.DataSource = (object)new BindingSource((object)dictionary, (string)null);
			((ListControl)cboKieuKy).DisplayMember = "Value";
			((ListControl)cboKieuKy).ValueMember = "Key";
			((Control)cboKieuKy).Refresh();
		}
		ShowDetailpnVT2_PhuLucKeKhai(pnVT2_PhuLucKeKhaiShow: true);
		Resize();
	}

	private void SetLayoutToKhaiNamTuThangDenThang()
	{
		ShowDetailpnVT1_1CheckBox(chkTKhaiLanXBShow: false, chkTkhaiThangShow: false, chkTKKyShow: false, chkTKQuyShow: false, ckNamShow: false, chkTuThangDenThangShow: false, chkTKLanPSShow: false, chkTuQuyDenQuyShow: false);
		ShowDetailpnVT1_2NgayThangQuyNam(pnNgayShow: false, pnThangShow: false, pnKyShow: false, pnQuyShow: false, pnNamShow: true);
		ShowDetailpnVT1_3TuNgayDenNgay(pnLanXuatBanShow: false, pnTuNgayShow: false, pnDenNgayShow: false, pnTuThangShow: true, pnDenThangShow: true);
		ShowDetailpnVT1_4LanDauBoXung(pnQuyetToan_LienTuc: false, pnRDLanDauBoXungShow: true, pnNgayLapShow: false);
		ShowDetailpnVT1_5DanhMucNghe(pnVT1_5DanhMucNgheShow: false);
		if (menuID == "97" || menuID == "H4" || menuID == "C8")
		{
			ShowDetailpnVT2_PhuLucKeKhai(pnVT2_PhuLucKeKhaiShow: true);
		}
		else
		{
			ShowDetailpnVT2_PhuLucKeKhai(pnVT2_PhuLucKeKhaiShow: false);
		}
		if (menuID == "L9")
		{
			ShowDetailpnVT1_3TuNgayDenNgay(pnLanXuatBanShow: false, pnTuNgayShow: false, pnDenNgayShow: false, pnTuThangShow: false, pnDenThangShow: false);
		}
		if (menuID == "B8")
		{
			((Control)txtSoTb).Enabled = !OptChinhthuc.Checked;
			((Control)mskNgayTb).Enabled = !OptChinhthuc.Checked;
			if (OptChinhthuc.Checked)
			{
				((Control)txtSoTb).Text = "";
				((Control)mskNgayTb).Text = "";
			}
		}
		Resize();
	}

	private void SetLayoutToKhaiNamQuy_BSTuThangDenThangDmnnPl()
	{
		ShowDetailpnVT1_1CheckBox(chkTKhaiLanXBShow: false, chkTkhaiThangShow: false, chkTKKyShow: false, chkTKQuyShow: true, ckNamShow: true, chkTuThangDenThangShow: false, chkTKLanPSShow: false, chkTuQuyDenQuyShow: false);
		ckNam.Checked = true;
		ShowDetailpnVT1_2NgayThangQuyNam(pnNgayShow: false, pnThangShow: false, pnKyShow: false, pnQuyShow: false, pnNamShow: true);
		ShowDetailpnVT1_3TuNgayDenNgay(pnLanXuatBanShow: false, pnTuNgayShow: false, pnDenNgayShow: false, pnTuThangShow: true, pnDenThangShow: true);
		ShowDetailpnVT1_4LanDauBoXung(pnQuyetToan_LienTuc: false, pnRDLanDauBoXungShow: true, pnNgayLapShow: false);
		ShowDetailpnVT1_5DanhMucNghe(pnVT1_5DanhMucNgheShow: true);
		ShowDetailpnVT2_PhuLucKeKhai(pnVT2_PhuLucKeKhaiShow: true);
		Resize();
	}

	private void SetLayoutToKhaiNamQuy()
	{
		ShowDetailpnVT1_1CheckBox(chkTKhaiLanXBShow: false, chkTkhaiThangShow: false, chkTKKyShow: false, chkTKQuyShow: true, ckNamShow: true, chkTuThangDenThangShow: false, chkTKLanPSShow: false, chkTuQuyDenQuyShow: false);
		ckNam.Checked = true;
		ShowDetailpnVT1_2NgayThangQuyNam(pnNgayShow: false, pnThangShow: false, pnKyShow: false, pnQuyShow: false, pnNamShow: true);
		ShowDetailpnVT1_3TuNgayDenNgay(pnLanXuatBanShow: false, pnTuNgayShow: false, pnDenNgayShow: false, pnTuThangShow: false, pnDenThangShow: false);
		ShowDetailpnVT1_4LanDauBoXung(pnQuyetToan_LienTuc: false, pnRDLanDauBoXungShow: true, pnNgayLapShow: false);
		ShowDetailpnVT1_5DanhMucNghe(pnVT1_5DanhMucNgheShow: false);
		ShowDetailpnVT2_PhuLucKeKhai(pnVT2_PhuLucKeKhaiShow: false);
		Resize();
	}

	private void SetLayoutToKhaiQuyBSTuThangDenThang()
	{
		ShowDetailpnVT1_1CheckBox(chkTKhaiLanXBShow: false, chkTkhaiThangShow: false, chkTKKyShow: false, chkTKQuyShow: false, ckNamShow: false, chkTuThangDenThangShow: false, chkTKLanPSShow: false, chkTuQuyDenQuyShow: false);
		ShowDetailpnVT1_2NgayThangQuyNam(pnNgayShow: false, pnThangShow: false, pnKyShow: false, pnQuyShow: true, pnNamShow: true);
		ShowDetailpnVT1_3TuNgayDenNgay(pnLanXuatBanShow: false, pnTuNgayShow: false, pnDenNgayShow: false, pnTuThangShow: true, pnDenThangShow: true);
		ShowDetailpnVT1_4LanDauBoXung(pnQuyetToan_LienTuc: false, pnRDLanDauBoXungShow: true, pnNgayLapShow: false);
		ShowDetailpnVT1_5DanhMucNghe(pnVT1_5DanhMucNgheShow: false);
		ShowDetailpnVT2_PhuLucKeKhai(pnVT2_PhuLucKeKhaiShow: false);
		Resize();
	}

	private void SetLayoutToKhaiKhongCoForm()
	{
		ShowDetailpnVT1_1CheckBox(chkTKhaiLanXBShow: false, chkTkhaiThangShow: false, chkTKKyShow: false, chkTKQuyShow: false, ckNamShow: false, chkTuThangDenThangShow: false, chkTKLanPSShow: false, chkTuQuyDenQuyShow: false);
		ShowDetailpnVT1_2NgayThangQuyNam(pnNgayShow: true, pnThangShow: true, pnKyShow: false, pnQuyShow: false, pnNamShow: true);
		ShowDetailpnVT1_3TuNgayDenNgay(pnLanXuatBanShow: false, pnTuNgayShow: false, pnDenNgayShow: false, pnTuThangShow: false, pnDenThangShow: false);
		ShowDetailpnVT1_4LanDauBoXung(pnQuyetToan_LienTuc: false, pnRDLanDauBoXungShow: false, pnNgayLapShow: false);
		ShowDetailpnVT1_5DanhMucNghe(pnVT1_5DanhMucNgheShow: false);
		ShowDetailpnVT2_PhuLucKeKhai(pnVT2_PhuLucKeKhaiShow: false);
	}

	private void SetLayoutToKhaiQuyBSTuThangDenThangOnly()
	{
		ShowDetailpnVT1_1CheckBox(chkTKhaiLanXBShow: false, chkTkhaiThangShow: false, chkTKKyShow: false, chkTKQuyShow: false, ckNamShow: false, chkTuThangDenThangShow: false, chkTKLanPSShow: false, chkTuQuyDenQuyShow: false);
		ShowDetailpnVT1_2NgayThangQuyNam(pnNgayShow: false, pnThangShow: false, pnKyShow: false, pnQuyShow: true, pnNamShow: true);
		ShowDetailpnVT1_3TuNgayDenNgay(pnLanXuatBanShow: false, pnTuNgayShow: true, pnDenNgayShow: true, pnTuThangShow: false, pnDenThangShow: false);
		ShowDetailpnVT1_4LanDauBoXung(pnQuyetToan_LienTuc: false, pnRDLanDauBoXungShow: false, pnNgayLapShow: false);
		ShowDetailpnVT1_5DanhMucNghe(pnVT1_5DanhMucNgheShow: false);
		ShowDetailpnVT2_PhuLucKeKhai(pnVT2_PhuLucKeKhaiShow: false);
		Resize();
	}

	private void SetLayoutToKhaiThangLanPS_BS_DMNN_PL()
	{
		ShowDetailpnVT1_1CheckBox(chkTKhaiLanXBShow: false, chkTkhaiThangShow: true, chkTKKyShow: false, chkTKQuyShow: false, ckNamShow: false, chkTuThangDenThangShow: false, chkTKLanPSShow: true, chkTuQuyDenQuyShow: false);
		chkTkhaiThang.Checked = true;
		ShowDetailpnVT1_2NgayThangQuyNam(pnNgayShow: false, pnThangShow: true, pnKyShow: false, pnQuyShow: false, pnNamShow: true);
		ShowDetailpnVT1_3TuNgayDenNgay(pnLanXuatBanShow: false, pnTuNgayShow: false, pnDenNgayShow: false, pnTuThangShow: false, pnDenThangShow: false);
		ShowDetailpnVT1_4LanDauBoXung(pnQuyetToan_LienTuc: false, pnRDLanDauBoXungShow: true, pnNgayLapShow: false);
		ShowDetailpnVT1_5DanhMucNghe(pnVT1_5DanhMucNgheShow: true);
		ShowDetailpnVT2_PhuLucKeKhai(pnVT2_PhuLucKeKhaiShow: true);
		Resize();
	}

	private void SetLayoutToKhaiNam_BS_TuThang_DenThang_DMNN_PL()
	{
		ShowDetailpnVT1_1CheckBox(chkTKhaiLanXBShow: false, chkTkhaiThangShow: false, chkTKKyShow: false, chkTKQuyShow: false, ckNamShow: false, chkTuThangDenThangShow: false, chkTKLanPSShow: false, chkTuQuyDenQuyShow: false);
		ckNam.Checked = true;
		ShowDetailpnVT1_2NgayThangQuyNam(pnNgayShow: false, pnThangShow: false, pnKyShow: false, pnQuyShow: false, pnNamShow: true);
		ShowDetailpnVT1_3TuNgayDenNgay(pnLanXuatBanShow: false, pnTuNgayShow: false, pnDenNgayShow: false, pnTuThangShow: true, pnDenThangShow: true);
		ShowDetailpnVT1_4LanDauBoXung(pnQuyetToan_LienTuc: false, pnRDLanDauBoXungShow: true, pnNgayLapShow: false);
		ShowDetailpnVT1_5DanhMucNghe(pnVT1_5DanhMucNgheShow: true);
		if ("~K2~K3~J7~E6~E7~D7~R6~X5~R9~E9~".IndexOf("~" + menuID + "~") >= 0)
		{
			ShowDetailpnLoaiTien(pnLoaiTien: true);
		}
		if (menuID == "E3")
		{
			ShowDetailpnVT2_PhuLucKeKhai(pnVT2_PhuLucKeKhaiShow: false);
		}
		else
		{
			ShowDetailpnVT2_PhuLucKeKhai(pnVT2_PhuLucKeKhaiShow: true);
		}
		Resize();
	}

	private void SetLayoutToKhaiQTTNCN()
	{
		try
		{
			((Control)pnVT1_1CheckBox).Visible = false;
			((Control)pnQuyetToan_LienTuc).Visible = false;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void SetLayoutToKhaiNamTuThangDenThangNone()
	{
		ShowDetailpnVT1_1CheckBox(chkTKhaiLanXBShow: false, chkTkhaiThangShow: false, chkTKKyShow: false, chkTKQuyShow: false, ckNamShow: false, chkTuThangDenThangShow: false, chkTKLanPSShow: false, chkTuQuyDenQuyShow: false);
		ShowDetailpnVT1_2NgayThangQuyNam(pnNgayShow: false, pnThangShow: false, pnKyShow: false, pnQuyShow: false, pnNamShow: true);
		ShowDetailpnVT1_3TuNgayDenNgay(pnLanXuatBanShow: false, pnTuNgayShow: false, pnDenNgayShow: false, pnTuThangShow: true, pnDenThangShow: true);
		ShowDetailpnVT1_4LanDauBoXung(pnQuyetToan_LienTuc: false, pnRDLanDauBoXungShow: false, pnNgayLapShow: false);
		ShowDetailpnVT1_5DanhMucNghe(pnVT1_5DanhMucNgheShow: false);
		ShowDetailpnVT2_PhuLucKeKhai(pnVT2_PhuLucKeKhaiShow: false);
		Resize();
	}

	private void SetLayoutToKhaiNamNone()
	{
		ShowDetailpnVT1_1CheckBox(chkTKhaiLanXBShow: false, chkTkhaiThangShow: false, chkTKKyShow: false, chkTKQuyShow: false, ckNamShow: false, chkTuThangDenThangShow: false, chkTKLanPSShow: false, chkTuQuyDenQuyShow: false);
		ShowDetailpnVT1_2NgayThangQuyNam(pnNgayShow: false, pnThangShow: false, pnKyShow: false, pnQuyShow: false, pnNamShow: true);
		ShowDetailpnVT1_3TuNgayDenNgay(pnLanXuatBanShow: false, pnTuNgayShow: false, pnDenNgayShow: false, pnTuThangShow: false, pnDenThangShow: false);
		if (menuID == "B2" || menuID == "O9" || menuID == "M1")
		{
			ShowDetailpnVT1_4LanDauBoXung(pnQuyetToan_LienTuc: false, pnRDLanDauBoXungShow: false, pnNgayLapShow: false);
		}
		else
		{
			ShowDetailpnVT1_4LanDauBoXung(pnQuyetToan_LienTuc: false, pnRDLanDauBoXungShow: true, pnNgayLapShow: false);
		}
		ShowDetailpnVT1_5DanhMucNghe(pnVT1_5DanhMucNgheShow: false);
		if (menuID == "29")
		{
			ShowDetailpnVT2_PhuLucKeKhai(pnVT2_PhuLucKeKhaiShow: true);
		}
		else
		{
			ShowDetailpnVT2_PhuLucKeKhai(pnVT2_PhuLucKeKhaiShow: false);
		}
		if (menuID == "L9" || menuID == "I1" || menuID == "I2")
		{
			((Control)pnVT1_3TuNgayDenNgay).Visible = true;
			((Control)pnLanXuatBan).Visible = true;
			((Control)lblLanXuat).Text = "Số lần nộp";
		}
		Resize();
	}

	private void SetLayoutToKhaiNamCoNgayLapKHBS()
	{
		ShowDetailpnVT1_1CheckBox(chkTKhaiLanXBShow: false, chkTkhaiThangShow: false, chkTKKyShow: false, chkTKQuyShow: false, ckNamShow: false, chkTuThangDenThangShow: false, chkTKLanPSShow: false, chkTuQuyDenQuyShow: false);
		ShowDetailpnVT1_2NgayThangQuyNam(pnNgayShow: false, pnThangShow: false, pnKyShow: false, pnQuyShow: false, pnNamShow: true);
		ShowDetailpnVT1_3TuNgayDenNgay(pnLanXuatBanShow: false, pnTuNgayShow: false, pnDenNgayShow: false, pnTuThangShow: false, pnDenThangShow: false);
		if (menuID == "63" || menuID == "L1")
		{
			ShowDetailpnVT1_4LanDauBoXung(pnQuyetToan_LienTuc: true, pnRDLanDauBoXungShow: true, pnNgayLapShow: false);
			((Control)pnQuyetToan_LienTuc).Height = ((Control)chk_CoSoThanhLap).Height + 10;
			((Control)chk_CoSoThanhLap).Visible = true;
			chk_CoSoThanhLap.Checked = false;
			((Control)chkQTNamDau_KhongHoatDongLienTuc).Visible = false;
			((Control)chkQTTungNam_HoatDongLienTuc).Visible = false;
			if (menuID == "L1")
			{
				((Control)chk_CoSoThanhLap).Text = "Khai cho ĐĐKD mới thành lập trong năm";
			}
		}
		else
		{
			ShowDetailpnVT1_4LanDauBoXung(pnQuyetToan_LienTuc: false, pnRDLanDauBoXungShow: true, pnNgayLapShow: false);
		}
		ShowDetailpnVT1_5DanhMucNghe(pnVT1_5DanhMucNgheShow: false);
		ShowDetailpnVT2_PhuLucKeKhai(pnVT2_PhuLucKeKhaiShow: false);
		Resize();
	}

	private void SetLayoutToKhaiNamLanPhatSinh()
	{
		ShowDetailpnVT1_1CheckBox(chkTKhaiLanXBShow: false, chkTkhaiThangShow: false, chkTKKyShow: false, chkTKQuyShow: false, ckNamShow: true, chkTuThangDenThangShow: false, chkTKLanPSShow: true, chkTuQuyDenQuyShow: false);
		ckNam.Checked = true;
		ResetpnVT1_2NgayThangQuyNam();
		ShowDetailpnVT1_3TuNgayDenNgay(pnLanXuatBanShow: false, pnTuNgayShow: false, pnDenNgayShow: false, pnTuThangShow: false, pnDenThangShow: false);
		ShowDetailpnVT1_4LanDauBoXung(pnQuyetToan_LienTuc: false, pnRDLanDauBoXungShow: true, pnNgayLapShow: false);
		ShowDetailpnVT1_5DanhMucNghe(pnVT1_5DanhMucNgheShow: false);
		if (menuID == "55" || menuID == "T5")
		{
			ShowDetailpnVT2_PhuLucKeKhai(pnVT2_PhuLucKeKhaiShow: true);
		}
		else
		{
			ShowDetailpnVT2_PhuLucKeKhai(pnVT2_PhuLucKeKhaiShow: false);
		}
		if (menuID == "T5" || menuID == "Y4")
		{
			ShowDetailpnVT1_3TuNgayDenNgay(pnLanXuatBanShow: false, pnTuNgayShow: true, pnDenNgayShow: true, pnTuThangShow: false, pnDenThangShow: false);
		}
		Resize();
	}

	private void SetLayoutToKhaiQuy()
	{
		ShowDetailpnVT1_1CheckBox(chkTKhaiLanXBShow: false, chkTkhaiThangShow: false, chkTKKyShow: false, chkTKQuyShow: false, ckNamShow: false, chkTuThangDenThangShow: false, chkTKLanPSShow: false, chkTuQuyDenQuyShow: false);
		ShowDetailpnVT1_2NgayThangQuyNam(pnNgayShow: false, pnThangShow: false, pnKyShow: false, pnQuyShow: true, pnNamShow: true);
		ShowDetailpnVT1_3TuNgayDenNgay(pnLanXuatBanShow: false, pnTuNgayShow: false, pnDenNgayShow: false, pnTuThangShow: false, pnDenThangShow: false);
		if (menuID == "66" || menuID == "B1")
		{
			ShowDetailpnVT1_4LanDauBoXung(pnQuyetToan_LienTuc: false, pnRDLanDauBoXungShow: false, pnNgayLapShow: false);
			ShowDetailpnVT1_5DanhMucNghe(pnVT1_5DanhMucNgheShow: false);
			ShowDetailpnVT2_PhuLucKeKhai(pnVT2_PhuLucKeKhaiShow: false);
		}
		else
		{
			ShowDetailpnVT1_4LanDauBoXung(pnQuyetToan_LienTuc: false, pnRDLanDauBoXungShow: true, pnNgayLapShow: false);
			if (menuID == "J2" || menuID == "K5" || menuID == "J4" || menuID == "N7" || menuID == "R7")
			{
				ShowDetailpnVT1_5DanhMucNghe(pnVT1_5DanhMucNgheShow: false);
			}
			else
			{
				ShowDetailpnVT1_5DanhMucNghe(pnVT1_5DanhMucNgheShow: true);
			}
			if (menuID == "E9" || menuID == "N7")
			{
				ShowDetailpnVT2_PhuLucKeKhai(pnVT2_PhuLucKeKhaiShow: false);
			}
			else
			{
				ShowDetailpnVT2_PhuLucKeKhai(pnVT2_PhuLucKeKhaiShow: true);
			}
			if ("~K2~K3~J7~E6~E7~D7~R6~X5~R9~E9~".IndexOf("~" + menuID + "~") >= 0)
			{
				ShowDetailpnLoaiTien(pnLoaiTien: true);
			}
		}
		Resize();
	}

	private void SetLayoutToKhaiQuyTuNgayDenNgay()
	{
		ShowDetailpnVT1_1CheckBox(chkTKhaiLanXBShow: false, chkTkhaiThangShow: false, chkTKKyShow: false, chkTKQuyShow: false, ckNamShow: false, chkTuThangDenThangShow: false, chkTKLanPSShow: false, chkTuQuyDenQuyShow: false);
		ShowDetailpnVT1_2NgayThangQuyNam(pnNgayShow: false, pnThangShow: false, pnKyShow: false, pnQuyShow: true, pnNamShow: true);
		ShowDetailpnVT1_3TuNgayDenNgay(pnLanXuatBanShow: false, pnTuNgayShow: true, pnDenNgayShow: true, pnTuThangShow: false, pnDenThangShow: false);
		if (menuID == "38")
		{
			ShowDetailpnVT1_4LanDauBoXung(pnQuyetToan_LienTuc: false, pnRDLanDauBoXungShow: true, pnNgayLapShow: false);
		}
		else
		{
			ShowDetailpnVT1_4LanDauBoXung(pnQuyetToan_LienTuc: false, pnRDLanDauBoXungShow: false, pnNgayLapShow: false);
		}
		ShowDetailpnVT1_5DanhMucNghe(pnVT1_5DanhMucNgheShow: false);
		if (menuID == "13" || menuID == "14" || menuID == "B1" || menuID == "B7")
		{
			ShowDetailpnVT2_PhuLucKeKhai(pnVT2_PhuLucKeKhaiShow: false);
		}
		else
		{
			ShowDetailpnVT2_PhuLucKeKhai(pnVT2_PhuLucKeKhaiShow: true);
		}
		Resize();
	}

	private void SetLayoutToKhaiQuyTuThangDenThang()
	{
		ShowDetailpnVT1_1CheckBox(chkTKhaiLanXBShow: false, chkTkhaiThangShow: false, chkTKKyShow: false, chkTKQuyShow: true, ckNamShow: false, chkTuThangDenThangShow: true, chkTKLanPSShow: false, chkTuQuyDenQuyShow: false);
		ShowDetailpnVT1_3TuNgayDenNgay(pnLanXuatBanShow: false, pnTuNgayShow: false, pnDenNgayShow: false, pnTuThangShow: false, pnDenThangShow: false);
		ShowDetailpnVT1_4LanDauBoXung(pnQuyetToan_LienTuc: false, pnRDLanDauBoXungShow: true, pnNgayLapShow: false);
		ShowDetailpnVT1_5DanhMucNghe(pnVT1_5DanhMucNgheShow: false);
		ShowDetailpnVT2_PhuLucKeKhai(pnVT2_PhuLucKeKhaiShow: false);
		chkTKQuy.Checked = true;
	}

	private void SetLayoutToKhaiThang()
	{
		ShowDetailpnVT1_1CheckBox(chkTKhaiLanXBShow: false, chkTkhaiThangShow: false, chkTKKyShow: false, chkTKQuyShow: false, ckNamShow: false, chkTuThangDenThangShow: false, chkTKLanPSShow: false, chkTuQuyDenQuyShow: false);
		ShowDetailpnVT1_2NgayThangQuyNam(pnNgayShow: false, pnThangShow: true, pnKyShow: false, pnQuyShow: false, pnNamShow: true);
		ShowDetailpnVT1_3TuNgayDenNgay(pnLanXuatBanShow: false, pnTuNgayShow: false, pnDenNgayShow: false, pnTuThangShow: false, pnDenThangShow: false);
		if ("T3;".Contains(menuID + ";"))
		{
			ShowDetailpnVT1_4LanDauBoXung(pnQuyetToan_LienTuc: false, pnRDLanDauBoXungShow: false, pnNgayLapShow: false);
		}
		else
		{
			ShowDetailpnVT1_4LanDauBoXung(pnQuyetToan_LienTuc: false, pnRDLanDauBoXungShow: true, pnNgayLapShow: false);
		}
		ShowDetailpnVT1_5DanhMucNghe(pnVT1_5DanhMucNgheShow: false);
		if (menuID == "G9" || menuID == "U7")
		{
			ShowDetailpnVT2_PhuLucKeKhai(pnVT2_PhuLucKeKhaiShow: true);
		}
		else
		{
			ShowDetailpnVT2_PhuLucKeKhai(pnVT2_PhuLucKeKhaiShow: false);
		}
		if ("U8~".Contains(WinSettings.MenuID))
		{
			((Control)OptChinhthuc).Text = "Lần đầu";
			((Control)OptBosung).Text = "Thay thế";
		}
		if ("U8".Equals(WinSettings.MenuID))
		{
			ShowDetailpnMaHSDADT(pnMaHSDADT: true);
			((Control)label1).Text = "File";
			((Control)lblChuDADT).Visible = false;
			List<int> list = new List<int>();
			for (int i = 1; i <= 10; i++)
			{
				list.Add(i);
			}
			cbMaHSDNHT_DADT.DataSource = list;
		}
		Resize();
	}

	private void SetLayoutToKhaiThangLanPhatSinh()
	{
		ShowDetailpnVT1_1CheckBox(chkTKhaiLanXBShow: false, chkTkhaiThangShow: true, chkTKKyShow: false, chkTKQuyShow: false, ckNamShow: false, chkTuThangDenThangShow: false, chkTKLanPSShow: true, chkTuQuyDenQuyShow: false);
		ShowDetailpnVT1_3TuNgayDenNgay(pnLanXuatBanShow: false, pnTuNgayShow: false, pnDenNgayShow: false, pnTuThangShow: false, pnDenThangShow: false);
		ShowDetailpnVT1_4LanDauBoXung(pnQuyetToan_LienTuc: false, pnRDLanDauBoXungShow: true, pnNgayLapShow: false);
		ShowDetailpnVT1_5DanhMucNghe(pnVT1_5DanhMucNgheShow: false);
		if (menuID == "90" || menuID == "N9")
		{
			ShowDetailpnVT2_PhuLucKeKhai(pnVT2_PhuLucKeKhaiShow: true);
		}
		else
		{
			ShowDetailpnVT2_PhuLucKeKhai(pnVT2_PhuLucKeKhaiShow: false);
		}
		chkTkhaiThang.Checked = true;
		if (menuID == "G3")
		{
			((Control)chkTKLanPS).Text = "Lần đầu";
			((Control)chkTkhaiThang).Text = "Định kỳ";
			ShowDetailpnVT1_4LanDauBoXung(pnQuyetToan_LienTuc: false, pnRDLanDauBoXungShow: false, pnNgayLapShow: false);
			((Control)pnVT1_3TuNgayDenNgay).Visible = true;
			((Control)pnLanXuatBan).Visible = true;
			((Control)lblLanXuat).Text = "Lần gửi";
		}
	}

	private void SetLayoutToKhaiThangQuy()
	{
		ShowDetailpnVT1_1CheckBox(chkTKhaiLanXBShow: false, chkTkhaiThangShow: true, chkTKKyShow: false, chkTKQuyShow: true, ckNamShow: false, chkTuThangDenThangShow: false, chkTKLanPSShow: false, chkTuQuyDenQuyShow: false);
		chkTkhaiThang.Checked = true;
		ResetpnVT1_2NgayThangQuyNam();
		ShowDetailpnVT1_3TuNgayDenNgay(pnLanXuatBanShow: false, pnTuNgayShow: false, pnDenNgayShow: false, pnTuThangShow: false, pnDenThangShow: false);
		ShowDetailpnVT1_4LanDauBoXung(pnQuyetToan_LienTuc: false, pnRDLanDauBoXungShow: true, pnNgayLapShow: false);
		if (menuID == "L7" || menuID == "L8" || menuID == "Y6" || menuID == "Z8")
		{
			ShowDetailpnVT1_5DanhMucNghe(pnVT1_5DanhMucNgheShow: false);
		}
		else
		{
			ShowDetailpnVT1_5DanhMucNghe(pnVT1_5DanhMucNgheShow: true);
		}
		ShowDetailpnVT2_PhuLucKeKhai(pnVT2_PhuLucKeKhaiShow: true);
		Resize();
	}

	private void SetLayoutToKhaiThangQuyKhongDMNN()
	{
		ShowDetailpnVT1_1CheckBox(chkTKhaiLanXBShow: false, chkTkhaiThangShow: true, chkTKKyShow: false, chkTKQuyShow: true, ckNamShow: false, chkTuThangDenThangShow: false, chkTKLanPSShow: false, chkTuQuyDenQuyShow: false);
		chkTkhaiThang.Checked = true;
		ResetpnVT1_2NgayThangQuyNam();
		ShowDetailpnVT1_3TuNgayDenNgay(pnLanXuatBanShow: false, pnTuNgayShow: false, pnDenNgayShow: false, pnTuThangShow: false, pnDenThangShow: false);
		ShowDetailpnVT1_4LanDauBoXung(pnQuyetToan_LienTuc: false, pnRDLanDauBoXungShow: true, pnNgayLapShow: false);
		ShowDetailpnVT1_5DanhMucNghe(pnVT1_5DanhMucNgheShow: false);
		ShowDetailpnVT2_PhuLucKeKhai(pnVT2_PhuLucKeKhaiShow: true);
		Resize();
	}

	private void SetLayoutToKhaiThangQuyTuNgayDenNgay()
	{
		ShowDetailpnVT1_1CheckBox(chkTKhaiLanXBShow: false, chkTkhaiThangShow: true, chkTKKyShow: false, chkTKQuyShow: true, ckNamShow: false, chkTuThangDenThangShow: false, chkTKLanPSShow: false, chkTuQuyDenQuyShow: false);
		if (menuID.Equals("M4"))
		{
			ShowDetailpnVT1_3TuNgayDenNgay(pnLanXuatBanShow: false, pnTuNgayShow: false, pnDenNgayShow: false, pnTuThangShow: false, pnDenThangShow: false);
			((Control)pnVT1_3TuNgayDenNgay).Visible = true;
			((Control)pnLanXuatBan).Visible = true;
			((Control)lblLanXuat).Text = "Tờ khai số";
		}
		else
		{
			ShowDetailpnVT1_3TuNgayDenNgay(pnLanXuatBanShow: false, pnTuNgayShow: true, pnDenNgayShow: true, pnTuThangShow: false, pnDenThangShow: false);
		}
		ShowDetailpnVT1_4LanDauBoXung(pnQuyetToan_LienTuc: false, pnRDLanDauBoXungShow: false, pnNgayLapShow: false);
		ShowDetailpnVT1_5DanhMucNghe(pnVT1_5DanhMucNgheShow: false);
		ShowDetailpnVT2_PhuLucKeKhai(pnVT2_PhuLucKeKhaiShow: true);
		chkTKQuy.Checked = true;
	}

	private void SetLayoutToKhaiThangQuyLanPhatSinh()
	{
		ShowDetailpnVT1_1CheckBox(chkTKhaiLanXBShow: false, chkTkhaiThangShow: true, chkTKKyShow: false, chkTKQuyShow: true, ckNamShow: false, chkTuThangDenThangShow: false, chkTKLanPSShow: true, chkTuQuyDenQuyShow: false);
		ShowDetailpnVT1_3TuNgayDenNgay(pnLanXuatBanShow: false, pnTuNgayShow: false, pnDenNgayShow: false, pnTuThangShow: false, pnDenThangShow: false);
		ShowDetailpnVT1_4LanDauBoXung(pnQuyetToan_LienTuc: false, pnRDLanDauBoXungShow: true, pnNgayLapShow: false);
		ShowDetailpnVT1_5DanhMucNghe(pnVT1_5DanhMucNgheShow: false);
		if (menuID == "71" || menuID == "D4" || WinSettings.MenuID.Equals("Z3"))
		{
			((Control)pnThuHo).Visible = true;
		}
		ShowDetailpnVT2_PhuLucKeKhai(pnVT2_PhuLucKeKhaiShow: false);
		chkTkhaiThang.Checked = true;
	}

	private void SetLayoutToKhaiTuNgayDenNgay()
	{
		ShowDetailpnVT1_1CheckBox(chkTKhaiLanXBShow: false, chkTkhaiThangShow: false, chkTKKyShow: false, chkTKQuyShow: false, ckNamShow: false, chkTuThangDenThangShow: false, chkTKLanPSShow: false, chkTuQuyDenQuyShow: false);
		ShowDetailpnVT1_2NgayThangQuyNam(pnNgayShow: false, pnThangShow: false, pnKyShow: false, pnQuyShow: false, pnNamShow: false);
		ShowDetailpnVT1_3TuNgayDenNgay(pnLanXuatBanShow: false, pnTuNgayShow: true, pnDenNgayShow: true, pnTuThangShow: false, pnDenThangShow: false);
		ShowDetailpnVT1_4LanDauBoXung(pnQuyetToan_LienTuc: false, pnRDLanDauBoXungShow: true, pnNgayLapShow: false);
		ShowDetailpnVT1_5DanhMucNghe(pnVT1_5DanhMucNgheShow: false);
		ShowDetailpnVT2_PhuLucKeKhai(pnVT2_PhuLucKeKhaiShow: true);
		Resize();
	}

	private void SetLayoutThemPhuLuc()
	{
		ShowDetailpnVT1_1CheckBox(chkTKhaiLanXBShow: false, chkTkhaiThangShow: false, chkTKKyShow: false, chkTKQuyShow: false, ckNamShow: false, chkTuThangDenThangShow: false, chkTKLanPSShow: false, chkTuQuyDenQuyShow: false);
		ShowDetailpnVT1_2NgayThangQuyNam(pnNgayShow: false, pnThangShow: false, pnKyShow: false, pnQuyShow: false, pnNamShow: false);
		ShowDetailpnVT1_3TuNgayDenNgay(pnLanXuatBanShow: false, pnTuNgayShow: false, pnDenNgayShow: false, pnTuThangShow: false, pnDenThangShow: false);
		ShowDetailpnVT1_4LanDauBoXung(pnQuyetToan_LienTuc: false, pnRDLanDauBoXungShow: false, pnNgayLapShow: false);
		ShowDetailpnVT1_5DanhMucNghe(pnVT1_5DanhMucNgheShow: false);
		((Control)pnVT1).Visible = false;
		ShowDetailpnVT2_PhuLucKeKhai(pnVT2_PhuLucKeKhaiShow: true);
		Resize();
	}

	private void ResetpnVT1_1CheckBox(string nameControl, bool value, Panel pn)
	{
		//IL_0020: Unknown result type (might be due to invalid IL or missing references)
		//IL_0026: Expected O, but got Unknown
		//IL_00a4: Unknown result type (might be due to invalid IL or missing references)
		//IL_00aa: Expected O, but got Unknown
		//IL_004a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0050: Expected O, but got Unknown
		//IL_011a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0120: Expected O, but got Unknown
		//IL_0152: Unknown result type (might be due to invalid IL or missing references)
		//IL_0158: Expected O, but got Unknown
		if (value)
		{
			foreach (Control item in (ArrangedElementCollection)((Control)pn).Controls)
			{
				Control val = item;
				if (val.Name != nameControl && val is CheckBox)
				{
					CheckBox val2 = (CheckBox)val;
					val2.Checked = false;
				}
			}
		}
		if (value)
		{
			return;
		}
		int num = 0;
		foreach (Control item2 in (ArrangedElementCollection)((Control)pn).Controls)
		{
			Control val = item2;
			if (val.Visible && val is CheckBox)
			{
				num++;
			}
		}
		if (num != 2)
		{
			return;
		}
		foreach (Control item3 in (ArrangedElementCollection)((Control)pn).Controls)
		{
			Control val = item3;
			if ((val.Name != nameControl) & val.Visible & (((object)val).GetType().ToString() == "System.Windows.Forms.Checkbox"))
			{
				CheckBox val2 = (CheckBox)val;
				val2.Checked = true;
			}
		}
	}

	private void ResetpnVT1_2NgayThangQuyNam()
	{
		//IL_0035: Unknown result type (might be due to invalid IL or missing references)
		//IL_003b: Expected O, but got Unknown
		FormProcess.BeginControlUpdate((Control)(object)WinSettings.myHTKKSystems);
		ShowDetailpnVT1_2NgayThangQuyNam(pnNgayShow: false, pnThangShow: false, pnKyShow: false, pnQuyShow: false, pnNamShow: false);
		foreach (CheckBox item in (ArrangedElementCollection)((Control)pnVT1_1CheckBox).Controls)
		{
			CheckBox val = item;
			if (!val.Checked)
			{
				continue;
			}
			switch (((Control)val).Name)
			{
			case "chkTKhaiLanXB":
				if (!((Control)pnVT1_2NgayThangQuyNam).Visible)
				{
					((Control)pnVT1_2NgayThangQuyNam).Visible = true;
				}
				((Control)pnNgay).Visible = true;
				((Control)pnThang).Visible = true;
				((Control)pnNam).Visible = true;
				if (menuID == "Q8" || menuID == "U5")
				{
					((Control)pnVT1_3TuNgayDenNgay).Visible = false;
				}
				if (strKieuKy == "9")
				{
					if (menuID == "J8" || menuID == "J7" || menuID == "R6")
					{
						if (!((Control)pnVT1_3TuNgayDenNgay).Visible)
						{
							((Control)pnVT1_3TuNgayDenNgay).Visible = true;
						}
						((Control)pnLanXuatBan).Visible = false;
					}
					else
					{
						if (!((Control)pnVT1_3TuNgayDenNgay).Visible)
						{
							((Control)pnVT1_3TuNgayDenNgay).Visible = true;
						}
						((Control)pnLanXuatBan).Visible = true;
					}
				}
				if (strKieuKy == "35")
				{
					ShowDetailpnVT1_3TuNgayDenNgay(pnLanXuatBanShow: false, pnTuNgayShow: false, pnDenNgayShow: false, pnTuThangShow: false, pnDenThangShow: false);
				}
				break;
			case "chkTkhaiThang":
				if (!((Control)pnVT1_2NgayThangQuyNam).Visible)
				{
					((Control)pnVT1_2NgayThangQuyNam).Visible = true;
				}
				((Control)pnThang).Visible = true;
				((Control)pnNam).Visible = true;
				if (strKieuKy == "9" || strKieuKy == "35" || strKieuKy == "42")
				{
					ShowDetailpnVT1_3TuNgayDenNgay(pnLanXuatBanShow: false, pnTuNgayShow: false, pnDenNgayShow: false, pnTuThangShow: false, pnDenThangShow: false);
				}
				break;
			case "chkTKKy":
				if (!((Control)pnVT1_2NgayThangQuyNam).Visible)
				{
					((Control)pnVT1_2NgayThangQuyNam).Visible = true;
				}
				((Control)pnKy).Visible = true;
				((Control)pnNam).Visible = true;
				if (strKieuKy == "37")
				{
					((Control)pnNam).Visible = false;
					((Control)pnVT1_2NgayThangQuyNam).Visible = true;
					ShowDetailpnVT1_2NgayThangQuyNam(pnNgayShow: false, pnThangShow: false, pnKyShow: false, pnQuyShow: false, pnNamShow: false);
					((Control)pnVT1_3TuNgayDenNgay).Visible = true;
					ShowDetailpnVT1_3TuNgayDenNgay(pnLanXuatBanShow: false, pnTuNgayShow: true, pnDenNgayShow: true, pnTuThangShow: false, pnDenThangShow: false);
				}
				if (strKieuKy == "42")
				{
					if (menuID == "Q8" || menuID == "U5")
					{
						((Control)pnNam).Visible = false;
						((Control)pnVT1_2NgayThangQuyNam).Visible = true;
						ShowDetailpnVT1_2NgayThangQuyNam(pnNgayShow: false, pnThangShow: false, pnKyShow: true, pnQuyShow: false, pnNamShow: true);
						((Control)pnVT1_3TuNgayDenNgay).Visible = true;
						ShowDetailpnVT1_3TuNgayDenNgay(pnLanXuatBanShow: false, pnTuNgayShow: false, pnDenNgayShow: false, pnTuThangShow: true, pnDenThangShow: true);
						((Control)mskTuNgay).Text = "";
						((Control)mskDenNgay).Text = "";
					}
					else
					{
						((Control)pnNam).Visible = false;
						((Control)pnVT1_2NgayThangQuyNam).Visible = true;
						ShowDetailpnVT1_2NgayThangQuyNam(pnNgayShow: false, pnThangShow: false, pnKyShow: false, pnQuyShow: false, pnNamShow: false);
						((Control)pnVT1_3TuNgayDenNgay).Visible = true;
						ShowDetailpnVT1_3TuNgayDenNgay(pnLanXuatBanShow: false, pnTuNgayShow: true, pnDenNgayShow: true, pnTuThangShow: false, pnDenThangShow: false);
						((Control)mskTuNgay).Text = "";
						((Control)mskDenNgay).Text = "";
					}
				}
				break;
			case "chkTKQuy":
				if (!((Control)pnVT1_2NgayThangQuyNam).Visible)
				{
					((Control)pnVT1_2NgayThangQuyNam).Visible = true;
				}
				((Control)pnQuy).Visible = true;
				((Control)pnNam).Visible = true;
				if (strKieuKy == "16")
				{
					ShowDetailpnVT1_3TuNgayDenNgay(pnLanXuatBanShow: false, pnTuNgayShow: false, pnDenNgayShow: false, pnTuThangShow: false, pnDenThangShow: false);
					((Control)pnQuyetToan_LienTuc).Visible = false;
					((Control)pnRDLanDauBoXung).Location = new Point(((Control)pnRDLanDauBoXung).Location.X, ((Control)pnQuyetToan_LienTuc).Location.Y + 1);
					((Control)pnNgayLap).Location = new Point(((Control)pnNgayLap).Location.X, ((Control)pnRDLanDauBoXung).Location.Y + 1);
				}
				if (strKieuKy == "35")
				{
					ShowDetailpnVT1_3TuNgayDenNgay(pnLanXuatBanShow: false, pnTuNgayShow: false, pnDenNgayShow: false, pnTuThangShow: false, pnDenThangShow: false);
				}
				if (strKieuKy == "42")
				{
					if (menuID == "Q8" || menuID == "U5")
					{
						((Control)pnVT1_3TuNgayDenNgay).Visible = true;
					}
					ShowDetailpnVT1_3TuNgayDenNgay(pnLanXuatBanShow: false, pnTuNgayShow: false, pnDenNgayShow: false, pnTuThangShow: true, pnDenThangShow: true);
				}
				break;
			case "ckNam":
				if (!((Control)pnVT1_2NgayThangQuyNam).Visible)
				{
					((Control)pnVT1_2NgayThangQuyNam).Visible = true;
				}
				((Control)pnNam).Visible = true;
				if (strKieuKy == "35")
				{
					((Control)pnVT1_3TuNgayDenNgay).Visible = true;
					ShowDetailpnVT1_3TuNgayDenNgay(pnLanXuatBanShow: false, pnTuNgayShow: false, pnDenNgayShow: false, pnTuThangShow: true, pnDenThangShow: true);
				}
				if (strKieuKy == "42")
				{
					((Control)pnNam).Visible = true;
					((Control)pnVT1_2NgayThangQuyNam).Visible = true;
					ShowDetailpnVT1_2NgayThangQuyNam(pnNgayShow: false, pnThangShow: false, pnKyShow: false, pnQuyShow: false, pnNamShow: true);
					((Control)pnVT1_3TuNgayDenNgay).Visible = false;
					ShowDetailpnVT1_3TuNgayDenNgay(pnLanXuatBanShow: false, pnTuNgayShow: false, pnDenNgayShow: false, pnTuThangShow: false, pnDenThangShow: false);
				}
				if (strKieuKy == "37")
				{
					((Control)pnNam).Visible = true;
					((Control)pnVT1_2NgayThangQuyNam).Visible = true;
					ShowDetailpnVT1_2NgayThangQuyNam(pnNgayShow: false, pnThangShow: false, pnKyShow: false, pnQuyShow: false, pnNamShow: true);
					((Control)pnVT1_3TuNgayDenNgay).Visible = false;
					ShowDetailpnVT1_3TuNgayDenNgay(pnLanXuatBanShow: false, pnTuNgayShow: false, pnDenNgayShow: false, pnTuThangShow: false, pnDenThangShow: false);
				}
				break;
			case "chkTuThangDenThang":
				if (!((Control)pnVT1_3TuNgayDenNgay).Visible)
				{
					((Control)pnVT1_3TuNgayDenNgay).Visible = true;
				}
				((Control)pnTuThang).Visible = true;
				((Control)pnDenThang).Visible = true;
				if (strKieuKy == "16")
				{
					ShowDetailpnVT1_2NgayThangQuyNam(pnNgayShow: false, pnThangShow: false, pnKyShow: false, pnQuyShow: false, pnNamShow: false);
					((Control)pnQuyetToan_LienTuc).Visible = true;
				}
				break;
			case "chkTKLanPS":
				if (!((Control)pnVT1_2NgayThangQuyNam).Visible)
				{
					((Control)pnVT1_2NgayThangQuyNam).Visible = true;
				}
				((Control)pnNgay).Visible = true;
				((Control)pnThang).Visible = true;
				((Control)pnNam).Visible = true;
				if (strKieuKy == "35")
				{
					ShowDetailpnVT1_3TuNgayDenNgay(pnLanXuatBanShow: false, pnTuNgayShow: false, pnDenNgayShow: false, pnTuThangShow: false, pnDenThangShow: false);
				}
				break;
			}
			return;
		}
		FormProcess.EndControlUpdate((Control)(object)WinSettings.myHTKKSystems);
	}

	private void CheckedChanged(object sender, EventArgs e)
	{
		//IL_0002: Unknown result type (might be due to invalid IL or missing references)
		//IL_0008: Expected O, but got Unknown
		//IL_001b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0021: Expected O, but got Unknown
		//IL_0069: Unknown result type (might be due to invalid IL or missing references)
		//IL_006f: Expected O, but got Unknown
		try
		{
			CheckBox val = new CheckBox();
			if (sender is CheckBox)
			{
				val = (CheckBox)sender;
			}
			if (!val.Checked)
			{
				return;
			}
			ResetpnVT1_1CheckBox(((Control)val).Name, val.Checked, pnVT1_1CheckBox);
			int num = 0;
			foreach (CheckBox item in (ArrangedElementCollection)((Control)pnVT1_1CheckBox).Controls)
			{
				CheckBox val2 = item;
				if (((Control)val2).Visible)
				{
					num++;
				}
			}
			ResetpnVT1_2NgayThangQuyNam();
			if (menuID == "28")
			{
				((Control)chkToChucKhaiThay).Enabled = true;
			}
			if (menuID == "G5")
			{
				((Control)chkToChucKhaiThay).Visible = false;
			}
			if (((Control)val).Name == "chkTKQuy" && val.Checked)
			{
				if (menuID == "G5")
				{
					((Control)pnVT1_3TuNgayDenNgay).Visible = true;
					ShowDetailpnVT1_3TuNgayDenNgay(pnLanXuatBanShow: false, pnTuNgayShow: false, pnDenNgayShow: false, pnTuThangShow: true, pnDenThangShow: true);
				}
				if (menuID == "L2")
				{
					((Control)pnVT1_3TuNgayDenNgay).Visible = true;
					ShowDetailpnVT1_3TuNgayDenNgay(pnLanXuatBanShow: false, pnTuNgayShow: false, pnDenNgayShow: false, pnTuThangShow: true, pnDenThangShow: true);
				}
				if (menuID == "G7")
				{
					((Control)pnVT1_3TuNgayDenNgay).Visible = true;
					ShowDetailpnVT1_3TuNgayDenNgay(pnLanXuatBanShow: false, pnTuNgayShow: false, pnDenNgayShow: false, pnTuThangShow: true, pnDenThangShow: true);
				}
				if (periodTypeValue.Equals("12"))
				{
					if (DateTime.Now.Month <= 12 && DateTime.Now.Month > 9)
					{
						((ListControl)cbbQuy).SelectedIndex = 3;
						((Control)mskYear).Text = DateTime.Now.Year.ToString();
					}
					if (DateTime.Now.Month <= 9 && DateTime.Now.Month > 6)
					{
						((ListControl)cbbQuy).SelectedIndex = 2;
						((Control)mskYear).Text = DateTime.Now.Year.ToString();
					}
					if (DateTime.Now.Month <= 6 && DateTime.Now.Month > 3)
					{
						((ListControl)cbbQuy).SelectedIndex = 1;
						((Control)mskYear).Text = DateTime.Now.Year.ToString();
					}
					if (DateTime.Now.Month <= 3 && DateTime.Now.Month > 0)
					{
						((ListControl)cbbQuy).SelectedIndex = 0;
						((Control)mskYear).Text = DateTime.Now.Year.ToString();
					}
				}
				else
				{
					if (DateTime.Now.Month <= 12 && DateTime.Now.Month > 9)
					{
						((ListControl)cbbQuy).SelectedIndex = 2;
						((Control)mskYear).Text = DateTime.Now.Year.ToString();
					}
					if (DateTime.Now.Month <= 9 && DateTime.Now.Month > 6)
					{
						((ListControl)cbbQuy).SelectedIndex = 1;
						((Control)mskYear).Text = DateTime.Now.Year.ToString();
					}
					if (DateTime.Now.Month <= 6 && DateTime.Now.Month > 3)
					{
						((ListControl)cbbQuy).SelectedIndex = 0;
						((Control)mskYear).Text = DateTime.Now.Year.ToString();
					}
					if (DateTime.Now.Month <= 3 && DateTime.Now.Month > 0)
					{
						((ListControl)cbbQuy).SelectedIndex = 3;
						((Control)mskYear).Text = (DateTime.Now.Year - 1).ToString();
					}
				}
				if (periodTypeValue.Equals("11"))
				{
					int iQuy = int.Parse(((Control)cbbQuy).Text);
					int iNam = int.Parse(((Control)mskYear).Text);
					DateTime ngayDauQuy = GeneralFunctions.GetNgayDauQuy(iQuy, iNam, 1, 1);
					DateTime ngayCuoiQuy = GeneralFunctions.GetNgayCuoiQuy(iQuy, iNam, 1, 1);
					((Control)mskTuNgay).Text = ngayDauQuy.ToString("dd/MM/yyyy");
					((Control)mskDenNgay).Text = ngayCuoiQuy.ToString("dd/MM/yyyy");
				}
				else
				{
					((Control)mskTuThang).Text = GeneralFunctions.QuarterToFirstMonth(((ListControl)cbbQuy).SelectedIndex + 1).ToString().PadLeft(2, '0') + "/" + ((Control)mskYear).Text;
					((Control)mskDenThang).Text = GeneralFunctions.QuarterToLastMonth(((ListControl)cbbQuy).SelectedIndex + 1).ToString().PadLeft(2, '0') + "/" + ((Control)mskYear).Text;
				}
				loadDMNN_01GTGT();
			}
			if (((Control)val).Name == "chkTkhaiThang" && val.Checked)
			{
				int num2 = defaultMonth;
				int num3 = defaultYear;
				if (num2 == 1)
				{
					num2 = 12;
					num3--;
				}
				else
				{
					num2--;
				}
				((Control)mskMonth).Text = num2.ToString().PadLeft(2, '0');
				((Control)mskYear).Text = num3.ToString();
				if (menuID == "L2")
				{
					((Control)pnVT1_3TuNgayDenNgay).Visible = false;
					ShowDetailpnVT1_3TuNgayDenNgay(pnLanXuatBanShow: false, pnTuNgayShow: false, pnDenNgayShow: false, pnTuThangShow: false, pnDenThangShow: false);
				}
				if (strKieuKy == "33")
				{
					if (chkTkhaiThang.Checked)
					{
						((Control)pnVT2_PhuLucKeKhai).Visible = true;
					}
				}
				else if (strKieuKy == "20")
				{
					((Control)mskTuNgay).Text = "01/" + num2.ToString().PadLeft(2, '0') + "/" + num3;
					DateTime dateTime = new DateTime(num3, num2, 1).AddMonths(1).AddDays(-1.0);
					((Control)mskDenNgay).Text = dateTime.ToString("dd/MM/yyyy");
				}
				loadDMNN_01GTGT();
			}
			if (((Control)val).Name == "chkTKLanPS" && val.Checked)
			{
				((Control)mskMonth).Text = defaultMonth.ToString().PadLeft(2, '0');
				((Control)mskYear).Text = defaultYear.ToString().PadLeft(2, '0');
				((Control)mskDay).Text = defaultDay.ToString().PadLeft(2, '0');
				if (strKieuKy == "33" && chkTKLanPS.Checked)
				{
					((Control)pnVT2_PhuLucKeKhai).Visible = false;
					Resize();
				}
				if (menuID == "T5" || menuID == "Y4")
				{
					ShowDetailpnVT1_3TuNgayDenNgay(pnLanXuatBanShow: false, pnTuNgayShow: false, pnDenNgayShow: false, pnTuThangShow: false, pnDenThangShow: false);
					((Control)pnTruongHopQT).Visible = false;
				}
			}
			if (((Control)val).Name == "chkTKhaiLanXB" && val.Checked)
			{
				((Control)mskMonth).Text = defaultMonth.ToString().PadLeft(2, '0');
				((Control)mskYear).Text = defaultYear.ToString().PadLeft(2, '0');
				((Control)mskDay).Text = defaultDay.ToString().PadLeft(2, '0');
			}
			if (((Control)val).Name == "ckNam" && val.Checked)
			{
				if (!val.Checked || strKieuKy == "37")
				{
				}
				if (strKieuKy == "35")
				{
					((Control)mskYear).Text = (DateTime.Now.Year - 1).ToString();
					((Control)mskTuThang).Text = "01/" + ((Control)mskYear).Text;
					((Control)mskDenThang).Text = "12/" + ((Control)mskYear).Text;
				}
				if (strKieuKy == "42")
				{
					((Control)mskYear).Text = (DateTime.Now.Year - 1).ToString();
					((Control)mskTuThang).Text = "01/" + ((Control)mskYear).Text;
					((Control)mskDenThang).Text = "12/" + ((Control)mskYear).Text;
				}
				if (periodTypeValue.Equals("3"))
				{
					((Control)mskYear).Text = (defaultYear - 1).ToString();
				}
				else if (periodTypeValue.Equals("8"))
				{
					((Control)mskYear).Text = (DateTime.Now.Year - 1).ToString();
					((Control)mskTuThang).Text = "01/" + ((Control)mskYear).Text;
					((Control)mskDenThang).Text = "12/" + ((Control)mskYear).Text;
				}
				else if (periodTypeValue.Equals("12"))
				{
					((Control)mskYear).Text = defaultYear.ToString();
				}
				if (menuID == "T5" || menuID == "Y4")
				{
					((Control)pnVT1_3TuNgayDenNgay).Visible = true;
					ShowDetailpnVT1_3TuNgayDenNgay(pnLanXuatBanShow: false, pnTuNgayShow: true, pnDenNgayShow: true, pnTuThangShow: false, pnDenThangShow: false);
					((Control)mskYear).Text = (defaultYear - 1).ToString();
					((Control)pnTruongHopQT).Visible = true;
				}
			}
			if (((Control)val).Name == "chkTKKy" && val.Checked)
			{
				if (strKieuKy == "7")
				{
					if (menuID == "X1" || menuID == "X2")
					{
						if (DateTime.Now.Month <= 6 && DateTime.Now.Month > 0)
						{
							((ListControl)cbbKy).SelectedIndex = 0;
						}
						if (DateTime.Now.Month <= 12 && DateTime.Now.Month > 6)
						{
							((ListControl)cbbKy).SelectedIndex = 1;
						}
					}
					else
					{
						if (DateTime.Now.Month <= 12 && DateTime.Now.Month > 6)
						{
							((ListControl)cbbKy).SelectedIndex = 1;
							((Control)mskYear).Text = DateTime.Now.Year.ToString();
							((Control)mskTuNgay).Text = "01/07" + ((Control)mskYear).Text;
							((Control)mskDenNgay).Text = DateTime.DaysInMonth(DateTime.Now.Year, 12) + "/12/" + ((Control)mskYear).Text;
						}
						if (DateTime.Now.Month <= 6 && DateTime.Now.Month > 0)
						{
							((ListControl)cbbKy).SelectedIndex = 0;
							((Control)mskYear).Text = DateTime.Now.Year.ToString();
							((Control)mskTuNgay).Text = "01/01" + ((Control)mskYear).Text;
							((Control)mskDenNgay).Text = DateTime.DaysInMonth(DateTime.Now.Year, 6) + "/06/" + ((Control)mskYear).Text;
						}
					}
				}
				if (strKieuKy == "42")
				{
					if (DateTime.Now.Month <= 12 && DateTime.Now.Month > 6)
					{
						((ListControl)cbbKy).SelectedIndex = 1;
						((Control)mskYear).Text = DateTime.Now.Year.ToString();
					}
					if (DateTime.Now.Month <= 6 && DateTime.Now.Month > 0)
					{
						((ListControl)cbbKy).SelectedIndex = 0;
						((Control)mskYear).Text = DateTime.Now.Year.ToString();
					}
				}
			}
			if (menuID == "A4" || menuID == "99" || menuID == "H6")
			{
				((Control)mskTuThang).Text = "";
				((Control)mskDenThang).Text = "";
				if (chkTuThangDenThang.Checked)
				{
					((Control)lblTuThang).Text = "Từ tháng";
					((Control)lblDenThang).Text = "Đến tháng";
					mskTuThang.Mask = "00/0000";
					mskDenThang.Mask = "00/0000";
				}
				else
				{
					((Control)lblTuThang).Text = "Từ quý";
					((Control)lblDenThang).Text = "Đến quý";
					mskTuThang.Mask = "0/0000";
					mskDenThang.Mask = "0/0000";
				}
			}
			SetDefaultView();
			if (menuID == "79" || menuID == "T8" || menuID == "R5")
			{
				if (chkTKLanPS.Checked)
				{
					((Control)pnVT2_PhuLucKeKhai).Visible = false;
				}
				else
				{
					((Control)pnVT2_PhuLucKeKhai).Visible = true;
				}
			}
			if (((Control)pnVT2_PhuLucKeKhai).Visible)
			{
				LoadPhuLucSpread();
			}
			if ((menuID == "65" || menuID == "13") && ((Control)mskTuNgay).Visible && ((Control)mskDenNgay).Visible)
			{
				if (((Control)val).Name == "chkTKKy" && val.Checked)
				{
					((Control)mskTuNgay).Enabled = false;
					((Control)mskDenNgay).Enabled = false;
				}
				else
				{
					((Control)mskTuNgay).Enabled = true;
					((Control)mskDenNgay).Enabled = true;
				}
			}
			if (menuID == "R9")
			{
				((Control)pnLanXuatBan).Visible = true;
			}
			UpdatePeriodByVerSion();
			OptChinhthuc.Checked = true;
			if (ConvertHelpers.ConvertToInt(((Control)mskYear).Text) < 2021)
			{
				if (((Control)cbMaHSDNHT_DADT).Visible)
				{
					((Control)cbMaHSDNHT_DADT).Enabled = false;
					((Control)cbMaHSDNHT_DADT).Text = "";
				}
			}
			else if (((Control)cbMaHSDNHT_DADT).Visible)
			{
				((Control)cbMaHSDNHT_DADT).Enabled = true;
			}
			if (ConvertHelpers.ConvertToInt(((Control)mskYear).Text) < 2021)
			{
				if (((Control)txtMaHSDNHT_DADT).Visible)
				{
					((Control)txtMaHSDNHT_DADT).Enabled = false;
					((Control)txtMaHSDNHT_DADT).Text = "";
				}
			}
			else if (((Control)txtMaHSDNHT_DADT).Visible)
			{
				((Control)txtMaHSDNHT_DADT).Enabled = true;
			}
			chkToChucCaNhanDoanhNghiepChuyenNhuong_CheckedChanged();
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void Close_Click(object sender, EventArgs e)
	{
		((Form)this).Close();
	}

	private void Resize()
	{
		FormProcess.BeginControlUpdate((Control)(object)WinSettings.myHTKKSystems);
		if (strKieuKy == "9")
		{
			ArrangeVerticalPanel(pnVT1_5DanhMucNghe);
			((Control)pnVT1_5DanhMucNghe).Location = new Point(((Control)pnVT1).Location.X, ((Control)pnVT1).Location.Y);
		}
		ArrangeHorizonPanel(pnPPhapNopThue);
		ArrangeHorizonPanel(pnVT1_1CheckBox);
		ArrangeHorizonPanel(pnVT1_2NgayThangQuyNam);
		ArrangeHorizonPanel(pnVT1_3TuNgayDenNgay);
		ArrangeVerticalPanel(pnVT1_2NgayThangQuyNam);
		if (strKieuKy != "9")
		{
			ArrangeVerticalPanel(pnVT1_5DanhMucNghe);
		}
		ArrangeVerticalPanel(pnVT1_4LanDauBoXung);
		if (menuID == "95")
		{
			ArrangeVerticalPanel(pnRDLanDauBoXung);
		}
		((Control)pnPPhapNopThue).Location = new Point(((Control)pnVT1).Location.X + 7, ((Control)pnVT1_5DanhMucNghe).Location.Y + 1);
		((Control)pnTinhCQT).Location = new Point(((Control)pnVT1).Location.X, ((Control)pnPPhapNopThue).Location.Y + 1);
		((Control)pnMSTCaNhanToChuc).Location = new Point(((Control)pnVT1).Location.X, ((Control)pnTinhCQT).Location.Y + 1);
		((Control)pnVT1_1CheckBox).Location = new Point(((Control)pnVT1).Location.X, ((Control)pnMSTCaNhanToChuc).Location.Y + 1);
		if (menuID == "99" || menuID == "H6")
		{
			((Control)pnLoaiKyHoan).Location = new Point(((Control)pnVT1).Location.X, ((Control)pnTinhCQT).Location.Y + 1);
		}
		if (menuID == "Q4")
		{
			((Control)pnLoaiKyHoan).Location = new Point(((Control)pnVT1).Location.X, ((Control)pnTinhCQT).Location.Y + 1);
		}
		int num = (((Control)pnVT1).Width - ((Control)pnVT1_2NgayThangQuyNam).Width) / 2;
		if (menuID == "Y7" && chkTKLanPS.Checked)
		{
			((Control)lblNgayPsThuNhap).Visible = true;
			((Control)lblNgayPsThuNhap).Location = new Point(((Control)pnVT1_1CheckBox).Location.X + 35, ((Control)pnVT1_1CheckBox).Location.Y + 1);
			((Control)pnVT1_2NgayThangQuyNam).Location = new Point(((Control)pnVT1).Location.X + num, ((Control)lblNgayPsThuNhap).Location.Y + 1);
		}
		else
		{
			((Control)lblNgayPsThuNhap).Visible = false;
			((Control)pnVT1_2NgayThangQuyNam).Location = new Point(((Control)pnVT1).Location.X + num, ((Control)pnVT1_1CheckBox).Location.Y + 1);
		}
		int num2 = (((Control)pnVT1).Width - ((Control)pnVT1_3TuNgayDenNgay).Width) / 2;
		((Control)pnVT1_3TuNgayDenNgay).Location = new Point(((Control)pnVT1).Location.X + num2 + ctrPadR, ((Control)pnVT1_2NgayThangQuyNam).Location.Y + 1);
		((Control)pnVT1_4LanDauBoXung).Location = new Point(((Control)pnVT1_4LanDauBoXung).Location.X, ((Control)pnVT1_3TuNgayDenNgay).Location.Y + 1);
		if (strKieuKy != "9")
		{
			((Control)pnVT1_5DanhMucNghe).Location = new Point(((Control)pnVT1_5DanhMucNghe).Location.X, ((Control)pnVT1_4LanDauBoXung).Location.Y + 1);
		}
		if (menuID == "02" || menuID == "D2")
		{
			((Control)pnMaHSDNHT).Location = new Point(((Control)pnMaHSDNHT).Location.X, ((Control)pnVT1_4LanDauBoXung).Location.Y - 1);
		}
		if (menuID == "U8")
		{
			((Control)pnMaHSDNHT).Location = new Point(((Control)pnMaHSDNHT).Location.X, ((Control)pnVT1_4LanDauBoXung).Location.Y - 1);
			((Control)pnMaHSDNHT).Height = 30;
		}
		if ("~O1~O2~O3~O4~O5~".IndexOf("~" + menuID + "~") >= 0)
		{
			((Control)pnMst).Location = new Point(((Control)pnVT1_2NgayThangQuyNam).Location.X - 30, ((Control)pnVT1_2NgayThangQuyNam).Location.Y + 1);
			((Control)pnVT1_4LanDauBoXung).Location = new Point(((Control)pnMst).Location.X, ((Control)pnMst).Location.Y + ((Control)pnMst).Height + 1);
		}
		if (menuID == "P6" || menuID == "P1" || menuID == "X7" || menuID.Equals("Y9"))
		{
			((Control)pnMst).Location = new Point(((Control)pnVT1_3TuNgayDenNgay).Location.X - 90, ((Control)pnVT1_3TuNgayDenNgay).Location.Y + 1);
		}
		if ("~K2~K3~J7~E6~E7~D7~R6~X5~R9~E9~".IndexOf("~" + menuID + "~") >= 0)
		{
			if ("J7" == menuID || "R6" == menuID || "X5" == menuID)
			{
				((Control)pnLoaiTien).Location = new Point(((Control)pnVT1_2NgayThangQuyNam).Location.X - 30, ((Control)pnVT1_2NgayThangQuyNam).Location.Y + 1);
			}
			else if ("E6" == menuID || "E7" == menuID || "D7" == menuID || "E9" == menuID)
			{
				((Control)pnLoaiTien).Location = new Point(((Control)pnVT1_5DanhMucNghe).Location.X + 40, ((Control)pnVT1_5DanhMucNghe).Location.Y + ((Control)pnVT1_5DanhMucNghe).Height + 1);
			}
			else
			{
				((Control)pnLoaiTien).Location = new Point(((Control)pnVT1_3TuNgayDenNgay).Location.X - 30, ((Control)pnVT1_3TuNgayDenNgay).Location.Y + 1);
			}
			if (menuID == "R6")
			{
				((Control)pnChuyenNhuong).Location = new Point(((Control)pnLoaiTien).Location.X, ((Control)pnLoaiTien).Location.Y + 60);
				((Control)pnVT1_4LanDauBoXung).Location = new Point(((Control)pnChuyenNhuong).Location.X, ((Control)pnChuyenNhuong).Location.Y + ((Control)pnChuyenNhuong).Height + 1);
			}
			else
			{
				((Control)pnVT1_4LanDauBoXung).Location = new Point(((Control)pnLoaiTien).Location.X, ((Control)pnLoaiTien).Location.Y + ((Control)pnLoaiTien).Height + 1);
			}
		}
		if (menuID == "A4" || menuID == "99")
		{
			((Control)pnVT1_5DanhMucNghe).Location = new Point(((Control)pnVT1_5DanhMucNghe).Location.X, ((Control)pnVT1_3TuNgayDenNgay).Location.Y + 1);
			((Control)pnMaHSDNHT).Location = new Point(((Control)pnMaHSDNHT).Location.X, ((Control)pnVT1_3TuNgayDenNgay).Location.Y + 2);
		}
		if (menuID == "B8")
		{
			ShowDetailpnSoTb(pnSoTb: true);
			((Control)pnThongBao).Location = new Point(((Control)pnMaHSDNHT).Location.X, ((Control)pnVT1_3TuNgayDenNgay).Location.Y + 2);
		}
		if (menuID == "71" || menuID == "D4" || WinSettings.MenuID.Equals("Z3"))
		{
			((Control)pnThuHo).Location = new Point(((Control)pnThuHo).Location.X, ((Control)pnVT1_4LanDauBoXung).Location.Y - 1);
		}
		if (menuID == "U1")
		{
			((Control)pnChuyenNhuong).Location = new Point(((Control)pnVT1_5DanhMucNghe).Location.X, ((Control)pnVT1_5DanhMucNghe).Location.Y + 60);
		}
		if (menuID == "Y3")
		{
			((Control)pnChuyenNhuong).Location = new Point(((Control)pnVT1_5DanhMucNghe).Location.X, ((Control)pnVT1_5DanhMucNghe).Location.Y + 60);
		}
		ArrangeVerticalPanel(pnVT1);
		ArrangeVerticalPanel(f);
		if (menuID == "95")
		{
			ArrangeHorizonPanel(pnRDLanDauBoXung);
			((Control)pnRDLanDauBoXung).Location = new Point(((Control)pnVT1_2NgayThangQuyNam).Location.X, ((Control)pnRDLanDauBoXung).Location.Y);
		}
		((Form)this).Size = new Size(((Control)pnVT1).Size.Width + ctrPadR + 4, ((Control)f).Size.Height + ((Control)pnDenNgay).Location.Y + ((Control)pbHeader).Height + 4);
		((Control)f).Width = ((Form)this).Size.Width - 4;
		if (widthPeriod == 0)
		{
			widthPeriod = ((Form)this).Size.Width;
			((Control)WinSettings.myHTKKSystems.panelPeriod).Width = ((Form)this).Size.Width;
		}
		else
		{
			((Control)WinSettings.myHTKKSystems.panelPeriod).Width = widthPeriod;
		}
		((Control)WinSettings.myHTKKSystems.panelPeriod).Height = ((Control)f).Height + ((Control)pbHeader).Height + 4;
		FormProcess.EndControlUpdate((Control)(object)WinSettings.myHTKKSystems);
	}

	private void OptChinhthuc_CheckedChanged(object sender, EventArgs e)
	{
		((Control)pnLan).Visible = false;
		((Control)pnNgayLap).Visible = false;
		if (menuID == "B8")
		{
			((Control)txtSoTb).Enabled = !OptChinhthuc.Checked;
			((Control)mskNgayTb).Enabled = !OptChinhthuc.Checked;
			if (OptChinhthuc.Checked)
			{
				((Control)txtSoTb).Text = "";
				((Control)mskNgayTb).Text = "";
			}
		}
		if (strKieuKy == "12" || strKieuKy == "21" || strKieuKy == "6" || strKieuKy == "18" || strKieuKy == "24" || strKieuKy == "2" || strKieuKy == "29" || strKieuKy == "8" || strKieuKy == "30" || strKieuKy == "31" || strKieuKy == "36" || strKieuKy == "13")
		{
			if (menuID == "H4")
			{
				((Control)pnVT2_PhuLucKeKhai).Visible = true;
				LoadPhuLucSpread();
			}
			else
			{
				((Control)pnVT2_PhuLucKeKhai).Visible = false;
			}
		}
		else if (strKieuKy == "33")
		{
			if (chkTkhaiThang.Checked)
			{
				((Control)pnVT2_PhuLucKeKhai).Visible = true;
				LoadPhuLucSpread();
			}
			else
			{
				((Control)pnVT2_PhuLucKeKhai).Visible = false;
			}
		}
		else if (OptChinhthuc.Checked || "30,I5,K9,K1,M7,K6,K8,73,55,71,31,P2,X6".Contains(menuID))
		{
			if (elemList.Count > 2)
			{
				((Control)pnVT2_PhuLucKeKhai).Visible = true;
				LoadPhuLucSpread();
			}
		}
		else if (strKieuKy != "42")
		{
			axfpsPhuLucLockedSetAllValues(value: false);
		}
		ArrangeVerticalPanel(pnVT1_4LanDauBoXung);
		isCheckOptions = true;
		UpdatePeriodByVerSion();
	}

	private void axfpsPhuLucLockedSetAllValues(bool value)
	{
		try
		{
			XmlNodeList val = (xmlNode.SelectSingleNode("Validity").HasChildNodes ? xmlNode.SelectSingleNode("Validity").ChildNodes : null);
			if (!value && val != null && val.Count > 2)
			{
				if (axfpsPhuLuc.Sheets[0].RowCount > val.Count)
				{
					axfpsPhuLuc.Sheets[0].RowCount = val.Count;
				}
				for (int i = 0; i < axfpsPhuLuc.Sheets[0].RowCount; i++)
				{
					axfpsPhuLuc.Sheets[0].Cells[i, 0].Value = value;
					axfpsPhuLuc.Sheets[0].Cells[i, 0].Locked = value;
				}
				chkChonPhuLucKeKhaiAll.Checked = value;
				((Control)chkChonPhuLucKeKhaiAll).Enabled = !value;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void OptBosung_CheckedChanged(object sender, EventArgs e)
	{
		try
		{
			((Control)pnLan).Visible = true;
			if (menuID == "G4" || menuID == "N5" || menuID == "H5" || menuID == "O7" || menuID == "P5" || menuID == "U6" || "U8~".Contains(WinSettings.MenuID))
			{
				((Control)pnLan).Visible = false;
			}
			((Control)pnNgayLap).Visible = true;
			if (strKieuKy == "5" || strKieuKy == "35" || "30,I5,K9,K1,M7,K6,K8,73,55,71,31,P2,X6".Contains(menuID) || strKieuKy == "42")
			{
				((Control)pnVT2_PhuLucKeKhai).Visible = true;
			}
			else if (elemList != null && elemList.Count > 3)
			{
				((Control)pnVT2_PhuLucKeKhai).Visible = true;
			}
			else
			{
				((Control)pnVT2_PhuLucKeKhai).Visible = false;
			}
			((Control)pnRDLanDauBoXung).Location = new Point(((Control)pnRDLanDauBoXung).Location.X, ((Control)pnQuyetToan_LienTuc).Location.Y + 1);
			((Control)pnNgayLap).Location = new Point(((Control)pnNgayLap).Location.X, ((Control)pnRDLanDauBoXung).Location.Y + 1);
			ArrangeVerticalPanel(pnVT1_4LanDauBoXung);
			if (OptBosung.Checked)
			{
				GetPeriodInfo(WinSettings.TTPeriodInfo);
				if (!(menuID == "G4") && !(menuID == "N5") && !(menuID == "H5") && !(menuID == "O7") && !(menuID == "P5") && !(menuID == "U6") && !"U8~".Contains(WinSettings.MenuID))
				{
					strFileNameExtend = WinFnProcess.GetDataFileNameExtend(WinSettings.TTPeriodInfo);
					fullNameDataFileToKhaiChinh = WinDoWorkEvent.GetFileNameBySheet(WinSettings.TTPeriodInfo, strDataFileToKhaiChinh, HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst + "\\", strVersionXML);
					((Control)txtLan).Text = WinFnProcess.GetSoLanKeKhai(fullNameDataFileToKhaiChinh, strFileNameExtend);
				}
				((Control)mskNgayLapKHBS).Text = DateTime.Now.Day.ToString().PadLeft(2, '0') + "/" + DateTime.Now.Month.ToString().PadLeft(2, '0') + "/" + DateTime.Now.Year;
				if (strKieuKy == "24" || strKieuKy == "25" || strKieuKy == "5" || (strKieuKy == "12" && menuID != "77" && menuID != "87" && menuID != "88" && menuID != "97") || strKieuKy == "2" || strKieuKy == "29" || strKieuKy == "30" || strKieuKy == "31" || strKieuKy == "35" || menuID == "G4" || menuID == "N5" || menuID == "H5" || menuID == "O7" || menuID == "P5" || menuID == "U6" || menuID == "L6" || strKieuKy == "37" || strKieuKy == "42" || menuID == "P2" || menuID == "P7")
				{
					((Control)pnNgayLap).Visible = false;
				}
				else
				{
					((Control)pnNgayLap).Visible = true;
				}
				LoadPhuLucSpread();
			}
			ArrangeVerticalPanel(pnVT1_4LanDauBoXung);
			isCheckOptions = true;
			if (!(menuID == "G4") && !(menuID == "N5") && !(menuID == "H5") && !(menuID == "O7") && !(menuID == "P5") && !(menuID == "U6") && !"U8~".Contains(WinSettings.MenuID))
			{
				UpdatePeriodByVerSion();
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void ControlInit(XmlNode _xmlNode)
	{
		//IL_0001: Unknown result type (might be due to invalid IL or missing references)
		//IL_0007: Expected O, but got Unknown
		//IL_0035: Unknown result type (might be due to invalid IL or missing references)
		//IL_003b: Expected O, but got Unknown
		//IL_0197: Unknown result type (might be due to invalid IL or missing references)
		//IL_019e: Expected O, but got Unknown
		//IL_02cd: Unknown result type (might be due to invalid IL or missing references)
		//IL_02d4: Expected O, but got Unknown
		//IL_0494: Unknown result type (might be due to invalid IL or missing references)
		//IL_049e: Expected O, but got Unknown
		XmlDocument val = new XmlDocument();
		val.LoadXml(_xmlNode.OuterXml);
		XmlNodeList val2 = ((XmlNode)val).SelectNodes("/Menu");
		foreach (XmlNode item in val2)
		{
			XmlNode val3 = item;
			menuID = ((XmlNode)val3.Attributes["ID"]).Value;
			parentMenuID = ((XmlNode)val3.Attributes["ParentID"]).Value;
			periodTypeValue = ((XmlNode)val3.Attributes["PeriodTypeValue"]).Value;
			string value = ((XmlNode)val3.Attributes["KieuKy"]).Value;
			if (val3.Attributes["Class"] != null && val3.Attributes["VersionXML"] != null)
			{
				string text = ((XmlNode)val3.Attributes["Class"]).Value + HTKKConfig.APP_PREFIX + ((XmlNode)val3.Attributes["VersionXML"]).Value.Replace(".", "");
				if (text.Length > 0)
				{
					WinSettings.InstanceClass = WinDoWorkEvent.GetInstance(text);
				}
			}
			WinSettings.TTPeriodInfo.KieuKy = value;
			strKieuKy = value;
		}
		val2 = ((XmlNode)val).SelectNodes("/Menu/Validity");
		foreach (XmlNode item2 in val2)
		{
			XmlNode val4 = item2;
			startDate = ((val4.Attributes["StartDate"] == null) ? "" : ((XmlNode)val4.Attributes["StartDate"]).Value);
		}
		val2 = ((XmlNode)val).SelectNodes("/Menu/Validity/Sheet");
		XmlNode val5 = val2[0];
		strDanhMucNgheFile = ((XmlNode)val5.Attributes["DataFile"]).Value;
		XmlDataDocument xmlDataDocument = new XmlDataDocument();
		if (menuID == "01" || menuID == "I5")
		{
			strDanhMucNgheFile += "_240";
		}
		FileStream fileStream = new FileStream(HTKKConfig.INTERFACE_INI + "\\Catalogue_DM_NNKD.xml", FileMode.Open, FileAccess.Read);
		((XmlDocument)xmlDataDocument).Load((Stream)fileStream);
		XmlNodeList elementsByTagName = ((XmlDocument)xmlDataDocument).GetElementsByTagName("Item");
		Dictionary<string, string> dictionary = new Dictionary<string, string>();
		foreach (XmlNode item3 in elementsByTagName)
		{
			XmlNode val6 = item3;
			if (strKieuKy == "26" || strKieuKy == "38")
			{
				((Control)lblNganhKD).Text = "Đơn vị tính";
				dictionary = new Dictionary<string, string>();
				dictionary.Add("VND", "Việt Nam Đồng");
				dictionary.Add("NT", "Ngoại tệ");
				continue;
			}
			string value2 = ((XmlNode)val6.Attributes["Value"]).Value;
			string[] array = Regex.Split(value2, "###");
			if (array.Length < 3 || !(strDanhMucNgheFile == array[0]))
			{
				continue;
			}
			if (menuID == "40" || menuID == "08")
			{
				List<string> list = new List<string>();
				List<string> list2 = new List<string>();
				list2.Add("1");
				list2.Add("2");
				list2.Add("3");
				list = list2;
				if (list.Contains(array[1]))
				{
					dictionary.Add(array[1], array[2]);
				}
			}
			else
			{
				dictionary.Add(array[1], array[2]);
			}
		}
		if (dictionary != null && dictionary.Count > 0)
		{
			cboNganhKD.DataSource = (object)new BindingSource((object)dictionary, (string)null);
			((ListControl)cboNganhKD).DisplayMember = "Value";
			((ListControl)cboNganhKD).ValueMember = "Key";
			((Control)cboNganhKD).Refresh();
		}
		axfpsPhuLuc.Sheets[0].Columns[0].Width = 400f;
		axfpsPhuLuc.Sheets[0].Columns.Count = 1;
		WinSpreadProcess.WinSetupSpread(axfpsPhuLuc);
		GridLine horizontalGridLine = new GridLine(GridLineType.Flat, Color.Empty);
		GridLine verticalGridLine = new GridLine(GridLineType.Flat, Color.Empty);
		axfpsPhuLuc.Sheets[0].HorizontalGridLine = horizontalGridLine;
		axfpsPhuLuc.Sheets[0].VerticalGridLine = verticalGridLine;
		axfpsPhuLuc.Sheets[0].RowHeader.Visible = false;
		axfpsPhuLuc.Sheets[0].ColumnHeader.Visible = false;
		axfpsPhuLuc.BorderStyle = (BorderStyle)1;
		axfpsPhuLuc.RowSplitBoxPolicy = SplitBoxPolicy.Never;
		axfpsPhuLuc.HorizontalScrollBarPolicy = ScrollBarPolicy.Never;
		axfpsPhuLuc.Sheets[0].HorizontalGridLine = new GridLine();
		axfpsPhuLuc.FocusRenderer = null;
		GetAllControls((Control)(object)this);
	}

	private void GetAllControls(Control container)
	{
		//IL_0025: Unknown result type (might be due to invalid IL or missing references)
		//IL_002b: Expected O, but got Unknown
		//IL_004b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0055: Expected O, but got Unknown
		//IL_006f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0079: Expected O, but got Unknown
		//IL_0093: Unknown result type (might be due to invalid IL or missing references)
		//IL_009d: Expected O, but got Unknown
		//IL_00b7: Unknown result type (might be due to invalid IL or missing references)
		//IL_00c1: Expected O, but got Unknown
		SetWidthNgayThangNam((Control)(object)pnVT1_2NgayThangQuyNam);
		foreach (Control item in (ArrangedElementCollection)container.Controls)
		{
			Control val = item;
			GetAllControls(val);
			if (val is MaskedTextBox)
			{
				listMaskedTextBox.Add((MaskedTextBox)val);
			}
			else if (val is CheckBox)
			{
				listCheckBox.Add((CheckBox)val);
			}
			else if (val is ComboBox)
			{
				listComboBox.Add((ComboBox)val);
			}
			else if (val is RadioButton)
			{
				listRadioButton.Add((RadioButton)val);
			}
		}
	}

	private void FormInit()
	{
		((Control)pnVT1_1CheckBox).Size = new Size(360, ((Control)pnVT1_1CheckBox).Size.Height);
		((Control)pnVT1_2NgayThangQuyNam).Size = new Size(360, ((Control)pnVT1_1CheckBox).Size.Height);
		((Control)pnVT1_3TuNgayDenNgay).Size = new Size(360, ((Control)pnVT1_3TuNgayDenNgay).Size.Height);
		((Control)pnVT1_5DanhMucNghe).Size = new Size(360, ((Control)pnVT1_5DanhMucNghe).Size.Height);
		((Control)pnVT1_4LanDauBoXung).Size = new Size(360, ((Control)pnVT1_4LanDauBoXung).Size.Height);
		((Control)pnVT1).Size = new Size(360, ((Control)pnVT1).Size.Height);
		((Control)pnVT2_PhuLucKeKhai).Size = new Size(360, ((Control)pnVT2_PhuLucKeKhai).Size.Height);
	}

	private int GetPanelHeight(Panel pn)
	{
		if (((Control)pn).Visible)
		{
			return 0;
		}
		return ((Control)pn).Size.Height;
	}

	private void ArrangeHorizonPanel(Panel pn)
	{
		//IL_0031: Unknown result type (might be due to invalid IL or missing references)
		//IL_0038: Expected O, but got Unknown
		//IL_01f1: Unknown result type (might be due to invalid IL or missing references)
		//IL_01f8: Expected O, but got Unknown
		FormProcess.BeginControlUpdate((Control)(object)pn);
		List<ControlProperties> list = new List<ControlProperties>();
		int num = 0;
		int num2 = 0;
		ControlProperties item = default(ControlProperties);
		foreach (Control item2 in (ArrangedElementCollection)((Control)pn).Controls)
		{
			Control val = item2;
			if (val.Visible)
			{
				item.Name = val.Name;
				item.X = val.Location.X;
				item.Y = val.Location.Y;
				list.Add(item);
				num++;
				num2 += val.Width;
			}
		}
		if (num == 0)
		{
			((Control)pn).Visible = false;
			return;
		}
		((Control)pn).Visible = true;
		((Control)pn).Size = new Size(num2 + ckbPadL, ((Control)pn).Height);
		if (((Control)pn).Name == "pnVT1_1CheckBox")
		{
			int width = ((Control)pn).Parent.Width;
			((Control)pn).Width = width;
		}
		for (int i = 0; i < list.Count - 1; i++)
		{
			for (int j = i + 1; j < list.Count; j++)
			{
				if (list[i].X > list[j].X)
				{
					item = list[i];
					list[i] = list[j];
					list[j] = item;
				}
			}
		}
		Dictionary<int, int> dictionary = new Dictionary<int, int>();
		int num3 = numberColCheckBox;
		for (int i = 0; i < list.Count; i++)
		{
			foreach (Control item3 in (ArrangedElementCollection)((Control)pn).Controls)
			{
				Control val = item3;
				if (!(val.Name == list[i].Name))
				{
					continue;
				}
				if (val is CheckBox)
				{
					if (dictionary.Count != 0)
					{
						if (dictionary.ContainsKey(list[i].Y))
						{
							int num4 = dictionary[list[i].Y];
							dictionary.Remove(list[i].Y);
							dictionary.Add(list[i].Y, num4 + 1);
						}
						else
						{
							dictionary.Add(list[i].Y, 0);
						}
					}
					else
					{
						dictionary.Add(list[i].Y, 0);
					}
					int num5 = dictionary[list[i].Y];
					if (val.Name == "chkToChucKhaiThay")
					{
						val.Location = new Point(((Control)pn).Width / 4, list[i].Y);
					}
					else if (val.Name == "chkTuQuyDenQuy")
					{
						int num6 = num5 * ((((Control)pn).Width - 2 * ckbPadL) / num3) + ckbPadL;
						val.Location = new Point(num6 + 50, list[i].Y);
					}
					else
					{
						int num6 = num5 * ((((Control)pn).Width - 2 * ckbPadL) / num3) + ckbPadL;
						val.Location = new Point(num6, list[i].Y);
					}
				}
				else
				{
					val.Location = new Point(i * (((Control)pn).Width / num), val.Location.Y);
				}
			}
		}
		FormProcess.EndControlUpdate((Control)(object)pn);
	}

	private void ArrangeVerticalPanel(Panel pn)
	{
		//IL_002f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0035: Expected O, but got Unknown
		//IL_020e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0214: Expected O, but got Unknown
		//IL_016a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0170: Expected O, but got Unknown
		FormProcess.BeginControlUpdate((Control)(object)pn);
		List<ControlProperties> list = new List<ControlProperties>();
		int num = 0;
		ControlProperties item = default(ControlProperties);
		foreach (Control item2 in (ArrangedElementCollection)((Control)pn).Controls)
		{
			Control val = item2;
			if (val.Visible)
			{
				item.Name = val.Name;
				item.Y = val.Location.Y;
				list.Add(item);
				num++;
			}
		}
		for (int i = 0; i < list.Count - 1; i++)
		{
			for (int j = i + 1; j < list.Count; j++)
			{
				if (list[i].Y > list[j].Y)
				{
					item = list[i];
					list[i] = list[j];
					list[j] = item;
				}
			}
		}
		int num2 = 0;
		if (((Control)pn).Name == "pnVT1_2NgayThangQuyNam")
		{
			for (int i = 0; i < list.Count; i++)
			{
				foreach (Control item3 in (ArrangedElementCollection)((Control)pn).Controls)
				{
					Control val = item3;
					if (val.Name == list[i].Name)
					{
						num2 = val.Height + ckbPadL;
					}
				}
			}
		}
		else
		{
			for (int i = 0; i < list.Count; i++)
			{
				foreach (Control item4 in (ArrangedElementCollection)((Control)pn).Controls)
				{
					Control val = item4;
					if (val.Name == list[i].Name)
					{
						val.Location = new Point(val.Location.X, num2);
						num2 = num2 + val.Height + ckbPadL;
					}
				}
			}
		}
		((Control)pn).Size = new Size(((Control)pn).Size.Width, num2);
		FormProcess.EndControlUpdate((Control)(object)pn);
	}

	private void ShowDetailpnVT1_1CheckBox(bool chkTKhaiLanXBShow, bool chkTkhaiThangShow, bool chkTKKyShow, bool chkTKQuyShow, bool ckNamShow, bool chkTuThangDenThangShow, bool chkTKLanPSShow, bool chkTuQuyDenQuyShow)
	{
		try
		{
			((Control)chkTKhaiLanXB).Visible = chkTKhaiLanXBShow;
			((Control)chkTkhaiThang).Visible = chkTkhaiThangShow;
			((Control)chkTKKy).Visible = chkTKKyShow;
			((Control)chkTKQuy).Visible = chkTKQuyShow;
			((Control)ckNam).Visible = ckNamShow;
			((Control)chkTuThangDenThang).Visible = chkTuThangDenThangShow;
			((Control)chkTuQuyDenQuy).Visible = chkTuQuyDenQuyShow;
			((Control)chkTKLanPS).Visible = chkTKLanPSShow;
			if (!chkTKhaiLanXBShow && !chkTkhaiThangShow && chkTKKyShow && !chkTKQuyShow && ckNamShow && !chkTuThangDenThangShow && !chkTKLanPSShow && !chkTuQuyDenQuyShow)
			{
				((Control)ckNam).Location = new Point(((Control)chkTKhaiLanXB).Location.X, ((Control)ckNam).Location.Y);
				((Control)chkTKKy).Location = new Point(((Control)ckNam).Location.X + 1, ((Control)ckNam).Location.Y);
				numberColCheckBox = 2;
			}
			else if (chkTkhaiThangShow && chkTKQuyShow && ckNamShow && chkTKLanPSShow)
			{
				numberColCheckBox = 2;
				Control parent = ((Control)ckNam).Parent;
				parent.Height = parent.Height * 2 + ctrPadR - 5;
				((Control)ckNam).Location = new Point(((Control)chkTKhaiLanXB).Location.X, ((Control)ckNam).Location.Y);
				((Control)chkTKQuy).Location = new Point(((Control)ckNam).Location.X + 1, ((Control)ckNam).Location.Y);
				((Control)chkTkhaiThang).Location = new Point(((Control)ckNam).Location.X, ((Control)ckNam).Location.Y + ((Control)chkTkhaiThang).Height + ctrPadR - 3);
				((Control)chkTKLanPS).Location = new Point(((Control)ckNam).Location.X + 1, ((Control)ckNam).Location.Y + ((Control)chkTkhaiThang).Height + ctrPadR - 3);
				if (strKieuKy == "35")
				{
					((Control)pnVT1_1CheckBox).Height = parent.Height + ((Control)chkToChucKhaiThay).Height + ctrPadR - 5;
					((Control)pnVT1_1CheckBox).Controls.Add((Control)(object)chkToChucKhaiThay);
					((Control)chkToChucKhaiThay).Visible = true;
					((Control)chkToChucKhaiThay).Location = new Point(((Control)ckNam).Location.X, ((Control)chkTkhaiThang).Location.Y + ((Control)chkToChucKhaiThay).Height + ctrPadR - 3);
				}
				if (strKieuKy == "42")
				{
					((Control)pnVT1_1CheckBox).Height = parent.Height + ((Control)chkToChucKhaiThay).Height + ctrPadR - 5;
					((Control)pnVT1_1CheckBox).Controls.Add((Control)(object)chkToChucKhaiThay);
				}
			}
			else
			{
				numberColCheckBox = 3;
				int y = ((Control)chkTKhaiLanXB).Location.Y;
				((Control)chkTkhaiThang).Location = new Point(((Control)chkTKhaiLanXB).Location.X + 1, y);
				((Control)chkTKKy).Location = new Point(((Control)chkTkhaiThang).Location.X + 1, y);
				((Control)chkTKQuy).Location = new Point(((Control)chkTKKy).Location.X + 1, y);
				((Control)ckNam).Location = new Point(((Control)chkTKQuy).Location.X + 1, y);
				((Control)chkTuThangDenThang).Location = new Point(((Control)ckNam).Location.X + 1, y);
				((Control)chkTKLanPS).Location = new Point(((Control)chkTuThangDenThang).Location.X + 1, y);
				((Control)chkTuQuyDenQuy).Location = new Point(((Control)chkTKLanPS).Location.X + 1, y);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void ShowDetailpnVT1_2NgayThangQuyNam(bool pnNgayShow, bool pnThangShow, bool pnKyShow, bool pnQuyShow, bool pnNamShow)
	{
		FormProcess.BeginControlUpdate((Control)(object)pnVT1_2NgayThangQuyNam);
		((Control)pnNam).Visible = pnNamShow;
		((Control)pnNgay).Visible = pnNgayShow;
		((Control)pnKy).Visible = pnKyShow;
		((Control)pnQuy).Visible = pnQuyShow;
		((Control)pnThang).Visible = pnThangShow;
		((Control)pnThang).Location = new Point(((Control)pnNgay).Location.X + 1, ((Control)pnThang).Location.Y);
		((Control)pnKy).Location = new Point(((Control)pnThang).Location.X + 1, ((Control)pnKy).Location.Y);
		((Control)pnQuy).Location = new Point(((Control)pnKy).Location.X + 1, ((Control)pnQuy).Location.Y);
		((Control)pnNam).Location = new Point(((Control)pnQuy).Location.X + 1, ((Control)pnNam).Location.Y);
		FormProcess.EndControlUpdate((Control)(object)pnVT1_2NgayThangQuyNam);
	}

	private void ShowDetailpnVT1_3TuNgayDenNgay(bool pnLanXuatBanShow, bool pnTuNgayShow, bool pnDenNgayShow, bool pnTuThangShow, bool pnDenThangShow)
	{
		FormProcess.BeginControlUpdate((Control)(object)WinSettings.myHTKKSystems);
		((Control)pnLanXuatBan).Visible = pnLanXuatBanShow;
		((Control)pnTuNgay).Visible = pnTuNgayShow;
		((Control)pnDenNgay).Visible = pnDenNgayShow;
		((Control)pnTuThang).Visible = pnTuThangShow;
		((Control)pnDenThang).Visible = pnDenThangShow;
		((Control)pnTuNgay).Location = new Point(((Control)pnLanXuatBan).Location.X + 1, ((Control)pnTuNgay).Location.Y);
		((Control)pnDenNgay).Location = new Point(((Control)pnTuNgay).Location.X + 1, ((Control)pnDenNgay).Location.Y);
		((Control)pnTuThang).Location = new Point(((Control)pnDenNgay).Location.X + 1, ((Control)pnTuThang).Location.Y);
		((Control)pnDenThang).Location = new Point(((Control)pnTuThang).Location.X + 1, ((Control)pnDenThang).Location.Y);
		FormProcess.EndControlUpdate((Control)(object)WinSettings.myHTKKSystems);
	}

	private void ShowDetailpnVT1_5DanhMucNghe(bool pnVT1_5DanhMucNgheShow)
	{
		((Control)pnVT1_5DanhMucNghe).Visible = pnVT1_5DanhMucNgheShow;
	}

	private void ShowDetailpnMst(bool pnMst)
	{
		((Control)this.pnMst).Visible = pnMst;
	}

	private void ShowDetailpnLoaiTien(bool pnLoaiTien)
	{
		((Control)this.pnLoaiTien).Visible = pnLoaiTien;
	}

	private void ShowDetailpnMaHSDADT(bool pnMaHSDADT)
	{
		((Control)pnMaHSDNHT).Visible = pnMaHSDADT;
	}

	private void ShowDetailpnSoTb(bool pnSoTb)
	{
		((Control)pnThongBao).Visible = pnSoTb;
	}

	private void ShowDetailpnVT1_4LanDauBoXung(bool pnQuyetToan_LienTuc, bool pnRDLanDauBoXungShow, bool pnNgayLapShow)
	{
		((Control)this.pnQuyetToan_LienTuc).Visible = pnQuyetToan_LienTuc;
		if ("N7~".IndexOf(WinSettings.MenuID) < 0)
		{
			((Control)pnRDLanDauBoXung).Visible = pnRDLanDauBoXungShow;
		}
		if (menuID == "47" || menuID == "K7")
		{
			((Control)this.pnQuyetToan_LienTuc).Height = 20;
			((Control)this.pnQuyetToan_LienTuc).Visible = true;
			((Control)cbxCTDC).Visible = true;
			((Control)cbxLNDC).Visible = true;
			((Control)chkQTTungNam_HoatDongLienTuc).Visible = false;
			((Control)chkQTNamDau_KhongHoatDongLienTuc).Visible = false;
		}
		((Control)pnNgayLap).Visible = pnNgayLapShow;
		((Control)pnRDLanDauBoXung).Location = new Point(((Control)pnRDLanDauBoXung).Location.X, ((Control)this.pnQuyetToan_LienTuc).Location.Y + 1);
		((Control)pnNgayLap).Location = new Point(((Control)pnNgayLap).Location.X, ((Control)pnRDLanDauBoXung).Location.Y + 1);
	}

	private void ShowDetailpnVT1(bool pnVT1_1CheckBoxShow, bool pnVT1_2NgayThangQuyNamShow, bool pnVT1_3TuNgayDenNgayShow, bool pnVT1_5DanhMucNgheShow, bool pnVT1_4LanDauBoXungShow)
	{
		((Control)pnVT1_1CheckBox).Visible = pnVT1_1CheckBoxShow;
		((Control)pnVT1_2NgayThangQuyNam).Visible = pnVT1_2NgayThangQuyNamShow;
		((Control)pnVT1_3TuNgayDenNgay).Visible = pnVT1_3TuNgayDenNgayShow;
		((Control)pnVT1_5DanhMucNghe).Visible = pnVT1_5DanhMucNgheShow;
		((Control)pnVT1_4LanDauBoXung).Visible = pnVT1_4LanDauBoXungShow;
	}

	private void ShowDetailpnVT2_PhuLucKeKhai(bool pnVT2_PhuLucKeKhaiShow)
	{
		((Control)pnVT2_PhuLucKeKhai).Visible = pnVT2_PhuLucKeKhaiShow;
		((Control)pnVT2_PhuLucKeKhai).Location = new Point(((Control)pnVT1).Location.X, ((Control)pnQuy).Location.Y + 1);
		((Control)pnVT3Button).Location = new Point(((Control)pnVT2_PhuLucKeKhai).Location.X, ((Control)pnNam).Location.Y + 1);
	}

	private void ShowDetailpnVonChuyenNhuong(bool showChkToChucCaNhanChuyenNhuong, bool showChkDoanhNghiepChuyenNhuong)
	{
		if (!showChkToChucCaNhanChuyenNhuong && !showChkDoanhNghiepChuyenNhuong)
		{
			((Control)pnChuyenNhuong).Visible = false;
			chkToChucCaNhanChuyenNhuong.Checked = false;
			chkDoanhNghiepChuyenNhuong.Checked = false;
			return;
		}
		((Control)pnChuyenNhuong).Visible = true;
		if (showChkToChucCaNhanChuyenNhuong)
		{
			chkToChucCaNhanChuyenNhuong.Checked = true;
			((Control)chkToChucCaNhanChuyenNhuong).Enabled = true;
			chkDoanhNghiepChuyenNhuong.Checked = false;
			((Control)chkDoanhNghiepChuyenNhuong).Enabled = true;
		}
		else
		{
			chkDoanhNghiepChuyenNhuong.Checked = true;
			((Control)chkDoanhNghiepChuyenNhuong).Enabled = true;
			chkToChucCaNhanChuyenNhuong.Checked = false;
			((Control)chkToChucCaNhanChuyenNhuong).Enabled = true;
		}
		((Control)pnChuyenNhuong).Location = new Point(((Control)pnVT1_2NgayThangQuyNam).Location.X, ((Control)pnVT1_2NgayThangQuyNam).Location.Y - ((Control)pnVT1_2NgayThangQuyNam).Height + ctrPadR - 3);
	}

	private void ShowDetailpnTinhCQT(bool cboTinhShow, bool cboXaShow, bool cboCQTCucShow, bool cboCQTChiCucShow, string typeNopCQT)
	{
		try
		{
			if (cboTinhShow || cboCQTCucShow || cboCQTChiCucShow)
			{
				((Control)pnTinhCQT).Visible = true;
			}
			else
			{
				((Control)pnTinhCQT).Visible = false;
			}
			((Control)cboTinh).Visible = cboTinhShow;
			((Control)lblTinh).Visible = cboTinhShow;
			((Control)panel1).Visible = cboXaShow;
			((Control)cboXa).Visible = cboXaShow;
			((Control)lblXa).Visible = cboXaShow;
			((Control)cboCQTCuc).Visible = cboCQTCucShow;
			((Control)lblCQTCuc).Visible = cboCQTCucShow;
			((Control)cboCQTChiCuc).Visible = cboCQTChiCucShow;
			((Control)lblCQTChiCuc).Visible = cboCQTChiCucShow;
			if (cboCQTCucShow || cboCQTChiCucShow)
			{
				getCQTCapCuc();
				defaultCQT();
			}
			if (cboTinhShow)
			{
				Dictionary<string, string> dictionary = new Dictionary<string, string>((cboXaShow && WinSettings.ToKhaiInfo.NopTheoTinhCQT == "3") ? HTKKWinDictionaryDiaDanh.dicTinhTPMoi : HTKKWinDictionaryDiaDanh.dicTinhTP);
				dataCboTinh = dictionary.ToList();
				List<string> list = new List<string>();
				list.Add("Chọn tỉnh");
				list.AddRange(dictionary.Keys.ToList());
				cboTinh.DataSource = list;
			}
			if (WinSettings.ToKhaiInfo.NopTheoTinhCQT == "3")
			{
				loadDefaultTinhCqtTheoDiaBan();
			}
			if (typeNopCQT.Equals("2"))
			{
				((Control)pnCCT).Location = new Point(((Control)panel1).Location.X, ((Control)panel1).Location.Y + 2);
				((Control)pnCQTCuc).Location = new Point(((Control)pnTinh).Location.X, ((Control)pnTinh).Location.Y);
				((Control)pnTinhCQT).Height = ((Control)pnTinhCQT).Height - ((Control)pnTinh).Height - ((Control)panel1).Height - 20;
			}
			else if (typeNopCQT.Equals("1"))
			{
				((Control)pnCCT).Location = new Point(((Control)pnCQTCuc).Location.X, ((Control)pnCQTCuc).Location.Y + 10);
				((Control)pnCQTCuc).Location = new Point(((Control)panel1).Location.X, ((Control)panel1).Location.Y + 5);
				((Control)pnTinhCQT).Height = ((Control)pnTinhCQT).Height - ((Control)pnTinh).Height - 20;
			}
			else
			{
				((Control)pnCQTCuc).Location = new Point(((Control)pnCQTCuc).Location.X, ((Control)pnCQTCuc).Location.Y + 5);
				((Control)pnTinhCQT).Height = ((Control)pnTinhCQT).Height - 20;
			}
			Resize();
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void loadDefaultTinhCqtTheoDiaBan()
	{
		if (!string.IsNullOrEmpty(WinSettings.HTKKHeaderInfo.tenTinhNNT) && ((Control)cboTinh).Visible)
		{
			((Control)cboTinh).Text = WinSettings.HTKKHeaderInfo.tenTinhNNT;
			cboXa.DataSource = HTKKWinDictionaryDiaDanh.getDanhMucPhuongXaMoi(WinSettings.HTKKHeaderInfo.maTinhNNT);
			cboCQTCuc.DataSource = GeneralFunctions.loadCqtCapCucTheoDiaBanXa(WinSettings.HTKKHeaderInfo.maXaPhuongNNT).Values.ToList();
			cboCQTChiCuc.DataSource = GeneralFunctions.loadCqtCapChiCucTheoDiaBanXa(WinSettings.HTKKHeaderInfo.maXaPhuongNNT).Values.ToList();
		}
		((Control)cboXa).Text = WinSettings.HTKKHeaderInfo.tenXaPhuongNNT;
		((Control)cboCQTCuc).Text = WinSettings.HTKKHeaderInfo.tenCQTCapCuc;
		((Control)cboCQTChiCuc).Text = _tenCQTNoiNop;
		WinSettings.TTPeriodInfo.TenTinh = WinSettings.HTKKHeaderInfo.tenTinhNNT;
		WinSettings.TTPeriodInfo.MaTinh = WinSettings.HTKKHeaderInfo.maTinhNNT;
		WinSettings.TTPeriodInfo.TenXa = WinSettings.HTKKHeaderInfo.tenXaPhuongNNT;
		WinSettings.TTPeriodInfo.MaXa = WinSettings.HTKKHeaderInfo.maXaPhuongNNT;
		WinSettings.TTPeriodInfo.tenCQTCapCuc = WinSettings.HTKKHeaderInfo.tenCQTCapCuc;
		WinSettings.TTPeriodInfo.maCQTCapCuc = WinSettings.HTKKHeaderInfo.maCQTCapCuc;
		WinSettings.TTPeriodInfo.tenCQTQuanLy = _tenCQTNoiNop;
		WinSettings.TTPeriodInfo.maCQTQuanLy = _maCQTNoiNop;
	}

	private void CheckedQTChanged(object sender, EventArgs e)
	{
		//IL_0002: Unknown result type (might be due to invalid IL or missing references)
		//IL_0008: Expected O, but got Unknown
		//IL_0019: Unknown result type (might be due to invalid IL or missing references)
		//IL_001f: Expected O, but got Unknown
		try
		{
			CheckBox val = new CheckBox();
			if (sender is CheckBox)
			{
				val = (CheckBox)sender;
			}
			ResetpnVT1_1CheckBox(((Control)val).Name, val.Checked, pnQuyetToan_LienTuc);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void chkChonPhuLucKeKhaiAll_CheckedChanged(object sender, EventArgs e)
	{
	}

	private void axfpsPhuLucSetAllValues(bool value)
	{
		for (int i = 0; i < axfpsPhuLuc.Sheets[0].RowCount; i++)
		{
			if (!axfpsPhuLuc.Sheets[0].Cells[i, 0].Locked)
			{
				axfpsPhuLuc.Sheets[0].Cells[i, 0].Value = value;
			}
		}
		chkChonPhuLucKeKhaiAll.Checked = value;
	}

	private void axfpsPhuLucLockAllValues(bool value)
	{
		for (int i = 0; i < axfpsPhuLuc.Sheets[0].RowCount; i++)
		{
			axfpsPhuLuc.Sheets[0].Cells[i, 0].Locked = value;
		}
		((Control)chkChonPhuLucKeKhaiAll).Enabled = !value;
	}

	private void mskDay_Leave(object sender, EventArgs e)
	{
		try
		{
			((Control)mskDay).Text = ((Control)mskDay).Text.PadLeft(2, '0');
			if (((Control)pnVT2_PhuLucKeKhai).Visible && !((Control)mskDay).Text.Equals(WinSettings.TTPeriodInfo.Ngay))
			{
				LoadPhuLucSpread();
			}
			UpdatePeriodByVerSion();
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void mskMonth_Leave(object sender, EventArgs e)
	{
		try
		{
			if (((Control)mskMonth).Text.Length <= 0)
			{
				return;
			}
			if (((Control)pnVT2_PhuLucKeKhai).Visible)
			{
				((Control)mskMonth).Text = ((Control)mskMonth).Text.PadLeft(2, '0');
				if (!((Control)mskMonth).Text.Equals(WinSettings.TTPeriodInfo.Thang))
				{
					LoadPhuLucSpread();
				}
			}
			else
			{
				((Control)mskMonth).Text = ((Control)mskMonth).Text.PadLeft(2, '0');
			}
			if (periodTypeValue.Equals("11"))
			{
				int year = int.Parse(((Control)mskYear).Text);
				if (chkTkhaiThang.Checked)
				{
					int num = int.Parse(((Control)mskMonth).Text);
					((Control)mskTuNgay).Text = "01/" + num.ToString().PadLeft(2, '0') + "/" + year;
					DateTime dateTime = new DateTime(year, num + 1, 1).AddDays(-1.0);
					((Control)mskDenNgay).Text = dateTime.ToString("dd/MM/yyyy");
				}
			}
			if (!((Control)mskMonth).Text.Equals(WinSettings.TTPeriodInfo.Thang))
			{
				loadDMNN_01GTGT();
			}
			WinDoWorkEvent.event_Leave(axfpsPhuLuc, listCheckBox, listMaskedTextBox, listComboBox, listRadioButton, "mskMonth_Leave");
			UpdatePeriodByVerSion();
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void Period_31()
	{
		try
		{
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void mskYear_Leave(object sender, EventArgs e)
	{
		try
		{
			if (((Control)mskYear).Text.Length <= 0)
			{
				return;
			}
			string ngayBatDauNamTaiChinh = WinSettings.HTKKHeaderInfo.NgayBatDauNamTaiChinh;
			if (HTKKFormater.IsYear(((Control)mskYear).Text))
			{
				DateUtils dateUtils = new DateUtils();
				DateTime value = dateUtils.ToDate(((Control)mskYear).Text, "YYYY").Value;
				((Control)mskYear).Text = dateUtils.ToString(value, "YYYY");
			}
			if ("P8,P9,Q1,Q2,Q3".Contains(menuID))
			{
				axfpsPhuLucSetAllValues(value: true);
			}
			if (menuID.Equals("31"))
			{
				Period_31();
			}
			if (menuID.Equals("28"))
			{
				if (ckNam.Checked && !((Control)mskYear).Text.Equals(WinSettings.TTPeriodInfo.Nam))
				{
					((Control)mskTuThang).Text = "01/" + ((Control)mskYear).Text;
					((Control)mskDenThang).Text = "02/" + ((Control)mskYear).Text;
				}
			}
			else if (ckNam.Checked && !chkTKQuy.Checked && ((Control)mskTuThang).Visible && ((Control)mskDenThang).Visible)
			{
				((Control)mskTuThang).Text = "01/" + ((Control)mskYear).Text;
				((Control)mskDenThang).Text = "12/" + ((Control)mskYear).Text;
			}
			if (((Control)mskTuNgay).Visible && ((Control)mskDenNgay).Visible && !ngayBatDauNamTaiChinh.Equals(""))
			{
				if ("03,78,58,53,54,A2,A5,A6,A7,B3,B4,C1,C2,C3,C4,C5,N8,X3".Contains(menuID) || "N8,C6,C7,T5,H4,P8,P9,Q2,U3,U4,X4,Y4".Contains(menuID) || "U9".Equals(menuID))
				{
					string[] array = ngayBatDauNamTaiChinh.Split(new char[1] { '/' });
					((Control)mskTuNgay).Text = ngayBatDauNamTaiChinh + "/" + ((Control)mskYear).Text;
					DateUtils dateUtils2 = new DateUtils();
					((Control)mskDenNgay).Text = dateUtils2.ToString(GeneralFunctions.NgayCuoiNamTaiChinh(ConvertHelpers.ConvertToInt(((Control)mskYear).Text), ConvertHelpers.ConvertToInt(array[1]), ConvertHelpers.ConvertToInt(array[0])), "DDMMYYYY");
				}
				else if (strKieuKy.Equals("15"))
				{
					int month = GeneralFunctions.QuarterToFirstMonth(((ListControl)cbbQuy).SelectedIndex + 1);
					int month2 = GeneralFunctions.QuarterToLastMonth(((ListControl)cbbQuy).SelectedIndex + 1);
					if (int.TryParse(((Control)mskYear).Text.Trim(), out var result))
					{
						DateTime dateTime = new DateTime(result, month, 1);
						DateTime dateTime2 = new DateTime(result, month2, DateTime.DaysInMonth(result, month2));
						((Control)mskTuNgay).Text = dateTime.ToString("dd/MM/yyyy");
						((Control)mskDenNgay).Text = dateTime2.ToString("dd/MM/yyyy");
					}
				}
				else
				{
					((Control)mskTuNgay).Text = "01/01/" + ((Control)mskYear).Text;
					((Control)mskDenNgay).Text = "31/12/" + ((Control)mskYear).Text;
				}
			}
			if (((Control)mskTuThang).Visible && ((Control)mskDenThang).Visible)
			{
				if (WinSettings.TTPeriodInfo.KieuKy.Equals("Q"))
				{
					((Control)mskTuThang).Text = GeneralFunctions.QuarterToFirstMonth(((ListControl)cbbQuy).SelectedIndex + 1).ToString().PadLeft(2, '0') + "/" + ((Control)mskYear).Text;
					((Control)mskDenThang).Text = GeneralFunctions.QuarterToLastMonth(((ListControl)cbbQuy).SelectedIndex + 1).ToString().PadLeft(2, '0') + "/" + ((Control)mskYear).Text;
				}
				else
				{
					((Control)mskTuThang).Text = "01/" + ((Control)mskYear).Text;
					((Control)mskDenThang).Text = "12/" + ((Control)mskYear).Text;
				}
			}
			if (periodTypeValue.Equals("11"))
			{
				int num = int.Parse(((Control)mskYear).Text);
				if (strKieuKy == "38")
				{
					if (ConvertHelpers.ConvertToInt(((Control)cbbQuy).Text) > 0 && ConvertHelpers.ConvertToInt(((Control)mskYear).Text) > 0)
					{
						setQuyByNgayBatDauTC(ConvertHelpers.ConvertToInt(((Control)cbbQuy).Text), ConvertHelpers.ConvertToInt(((Control)mskYear).Text));
					}
				}
				else if (chkTkhaiThang.Checked)
				{
					int num2 = int.Parse(((Control)mskMonth).Text);
					((Control)mskTuNgay).Text = "01/" + num2.ToString().PadLeft(2, '0') + "/" + num;
					DateTime dateTime3 = new DateTime(num, num2 + 1, 1).AddDays(-1.0);
					((Control)mskDenNgay).Text = dateTime3.ToString("dd/MM/yyyy");
				}
				else if (chkTKKy.Checked)
				{
					if (((ListControl)cbbKy).SelectedIndex == 1)
					{
						((Control)mskTuNgay).Text = "01/07" + ((Control)mskYear).Text;
						((Control)mskDenNgay).Text = DateTime.DaysInMonth(ConvertHelpers.ConvertToInt(((Control)mskYear).Text), 12) + "/12/" + ((Control)mskYear).Text;
					}
					if (((ListControl)cbbKy).SelectedIndex == 0)
					{
						((Control)mskTuNgay).Text = "01/01" + ((Control)mskYear).Text;
						((Control)mskDenNgay).Text = DateTime.DaysInMonth(ConvertHelpers.ConvertToInt(((Control)mskYear).Text), 6) + "/06/" + ((Control)mskYear).Text;
					}
				}
				else
				{
					int iQuy = int.Parse(((Control)cbbQuy).Text);
					DateTime ngayDauQuy = GeneralFunctions.GetNgayDauQuy(iQuy, num, 1, 1);
					DateTime ngayCuoiQuy = GeneralFunctions.GetNgayCuoiQuy(iQuy, num, 1, 1);
					((Control)mskTuNgay).Text = ngayDauQuy.ToString("dd/MM/yyyy");
					((Control)mskDenNgay).Text = ngayCuoiQuy.ToString("dd/MM/yyyy");
				}
			}
			if (menuID == "X1" || menuID == "X2")
			{
				thayDoiTuNgayDenNgayTheoKy();
				if (((Control)cbbQuy).Visible && !string.IsNullOrEmpty(((Control)mskYear).Text) && ConvertHelpers.ConvertToInt(((Control)cbbQuy).Text) > 0 && ConvertHelpers.ConvertToInt(((Control)mskYear).Text) > 0)
				{
					setQuyByNgayBatDauTC(ConvertHelpers.ConvertToInt(((Control)cbbQuy).Text), ConvertHelpers.ConvertToInt(((Control)mskYear).Text));
				}
			}
			if (((Control)pnVT2_PhuLucKeKhai).Visible && !((Control)mskYear).Text.Equals(WinSettings.TTPeriodInfo.Nam))
			{
				LoadPhuLucSpread();
			}
			SetPeriodCustom();
			UpdatePeriodByVerSion();
			if (!((Control)mskYear).Text.Equals(WinSettings.TTPeriodInfo.Nam))
			{
				loadDMNN_01GTGT();
			}
			if (ConvertHelpers.ConvertToInt(((Control)mskYear).Text) < 2021)
			{
				if (((Control)cbMaHSDNHT_DADT).Visible)
				{
					((Control)cbMaHSDNHT_DADT).Enabled = false;
					((Control)cbMaHSDNHT_DADT).Text = "";
				}
			}
			else if (((Control)cbMaHSDNHT_DADT).Visible)
			{
				((Control)cbMaHSDNHT_DADT).Enabled = true;
			}
			if (ConvertHelpers.ConvertToInt(((Control)mskYear).Text) < 2021)
			{
				if (((Control)txtMaHSDNHT_DADT).Visible)
				{
					((Control)txtMaHSDNHT_DADT).Enabled = false;
					((Control)txtMaHSDNHT_DADT).Text = "";
				}
			}
			else if (((Control)txtMaHSDNHT_DADT).Visible)
			{
				((Control)txtMaHSDNHT_DADT).Enabled = true;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void mskTuNgay_Leave(object sender, EventArgs e)
	{
		try
		{
			bool success;
			DateTime dateTime = HTKKFormater.ConvertToDateTime(((Control)mskTuNgay).Text, out success);
			if (success)
			{
				((Control)mskTuNgay).Text = HTKKFormater.FormatDdmmyyyy(((Control)mskTuNgay).Text);
				if (((Control)pnVT2_PhuLucKeKhai).Visible && !((Control)mskTuNgay).Text.Equals(WinSettings.TTPeriodInfo.TuNgay))
				{
					LoadPhuLucSpread();
				}
			}
			UpdatePeriodByVerSion();
			CheckNextDenThangTheoQuyetToan();
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void mskDenNgay_Leave(object sender, EventArgs e)
	{
		try
		{
			bool success;
			DateTime dateTime = HTKKFormater.ConvertToDateTime(((Control)mskDenNgay).Text, out success);
			if (success)
			{
				((Control)mskDenNgay).Text = HTKKFormater.FormatDdmmyyyy(((Control)mskDenNgay).Text);
				if (((Control)pnVT2_PhuLucKeKhai).Visible && !((Control)mskDenNgay).Text.Equals(WinSettings.TTPeriodInfo.DenNgay))
				{
					LoadPhuLucSpread();
				}
			}
			UpdatePeriodByVerSion();
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void mskTuThang_Leave(object sender, EventArgs e)
	{
		try
		{
			string text = "";
			text = ((!(menuID == "Q4")) ? HTKKFormater.FormatDdmmyyyy("01/" + ((Control)mskTuThang).Text) : HTKKFormater.FormatDdmmyyyy("01/" + ((DateTime.Now.Month.ToString().Length == 1) ? ("0" + DateTime.Now.Month) : DateTime.Now.Month.ToString()) + "/" + ((Control)mskTuThang).Text));
			if (text.Length == 10)
			{
				if (menuID == "99" || menuID == "H6" || menuID == "Q4")
				{
					string text2 = ((ListControl)cboKieuKy).SelectedValue.ToString();
					if (text2 == "5")
					{
						((Control)mskTuThang).Text = text.Substring(6, 4);
					}
					switch (text2)
					{
					case "1":
						((Control)mskTuThang).Text = text.Substring(3, 7);
						break;
					case "2":
						((Control)mskTuThang).Text = text.Substring(4, 6);
						break;
					case "3":
						((Control)mskTuThang).Text = text.Substring(3, 7);
						break;
					case "4":
						((Control)mskTuThang).Text = text.Substring(4, 6);
						break;
					}
				}
				else if (chkTuQuyDenQuy.Checked)
				{
					((Control)mskTuThang).Text = text.Substring(4, 6);
				}
				else
				{
					((Control)mskTuThang).Text = text.Substring(3, 7);
				}
			}
			if (menuID == "A4" || menuID == "99")
			{
				string[] array = ((Control)mskDenThang).Text.Split(new char[1] { '/' });
				if (array != null && array.Length > 1)
				{
					if (((ListControl)cboNganhKD).SelectedValue.Equals("0039"))
					{
						if (ConvertHelpers.ConvertToInt(array[1]) >= 2021)
						{
							((Control)cbMaHSDNHT_DADT).Enabled = true;
						}
						else
						{
							((Control)cbMaHSDNHT_DADT).Enabled = false;
							((Control)cbMaHSDNHT_DADT).Text = "";
						}
					}
					else
					{
						((Control)cbMaHSDNHT_DADT).Enabled = false;
						((Control)cbMaHSDNHT_DADT).Text = "";
					}
				}
			}
			LoadPhuLucSpread();
			CheckNextDenThangTheoQuyetToan();
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void mskDenThang_Leave(object sender, EventArgs e)
	{
		try
		{
			string text = "";
			text = ((!(menuID == "Q4")) ? HTKKFormater.FormatDdmmyyyy("01/" + ((Control)mskDenThang).Text) : HTKKFormater.FormatDdmmyyyy("01/" + ((DateTime.Now.Month.ToString().Length == 1) ? ("0" + DateTime.Now.Month) : DateTime.Now.Month.ToString()) + "/" + ((Control)mskDenThang).Text));
			if (text.Length == 10)
			{
				if (menuID == "99" || menuID == "H6" || menuID == "Q4")
				{
					string text2 = ((ListControl)cboKieuKy).SelectedValue.ToString();
					if (text2 == "5")
					{
						((Control)mskDenThang).Text = text.Substring(6, 4);
					}
					switch (text2)
					{
					case "1":
						((Control)mskDenThang).Text = text.Substring(3, 7);
						break;
					case "2":
						((Control)mskDenThang).Text = text.Substring(4, 6);
						break;
					case "3":
						((Control)mskDenThang).Text = text.Substring(4, 6);
						break;
					case "4":
						((Control)mskDenThang).Text = text.Substring(3, 7);
						break;
					}
				}
				else if (chkTuQuyDenQuy.Checked)
				{
					((Control)mskDenThang).Text = text.Substring(4, 6);
				}
				else
				{
					((Control)mskDenThang).Text = text.Substring(3, 7);
				}
			}
			if (menuID == "A4" || menuID == "99")
			{
				string[] array = ((Control)mskDenThang).Text.Split(new char[1] { '/' });
				if (array != null && array.Length > 1)
				{
					if (((ListControl)cboNganhKD).SelectedValue.Equals("0039"))
					{
						if (ConvertHelpers.ConvertToInt(array[1]) >= 2021)
						{
							((Control)cbMaHSDNHT_DADT).Enabled = true;
						}
						else
						{
							((Control)cbMaHSDNHT_DADT).Enabled = false;
							((Control)cbMaHSDNHT_DADT).Text = "";
						}
					}
					else
					{
						((Control)cbMaHSDNHT_DADT).Enabled = false;
						((Control)cbMaHSDNHT_DADT).Text = "";
					}
				}
			}
			LoadPhuLucSpread();
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void mskNgayLapKHBS_Leave(object sender, EventArgs e)
	{
		try
		{
			if (!HTKKFormater.isNullString(((Control)mskNgayLapKHBS).Text.Replace("/", "").Replace("_", "")))
			{
				string text = HTKKFormater.FormatDdmmyyyy(((Control)mskNgayLapKHBS).Text);
				((Control)mskNgayLapKHBS).Text = text;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private bool CheckTuNgayDenNgay()
	{
		//IL_006c: Unknown result type (might be due to invalid IL or missing references)
		//IL_00be: Unknown result type (might be due to invalid IL or missing references)
		//IL_0106: Unknown result type (might be due to invalid IL or missing references)
		//IL_0263: Unknown result type (might be due to invalid IL or missing references)
		//IL_0381: Unknown result type (might be due to invalid IL or missing references)
		//IL_0315: Unknown result type (might be due to invalid IL or missing references)
		//IL_08bd: Unknown result type (might be due to invalid IL or missing references)
		//IL_0526: Unknown result type (might be due to invalid IL or missing references)
		//IL_074b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0573: Unknown result type (might be due to invalid IL or missing references)
		//IL_0798: Unknown result type (might be due to invalid IL or missing references)
		//IL_0640: Unknown result type (might be due to invalid IL or missing references)
		//IL_05e4: Unknown result type (might be due to invalid IL or missing references)
		bool success = true;
		try
		{
			string kieuKy = WinSettings.TTPeriodInfo.KieuKy;
			if (((Control)pnTuNgay).Visible & ((Control)pnDenNgay).Visible)
			{
				DateTime dateTime = HTKKFormater.ConvertToDateTime(((Control)mskTuNgay).Text, out success);
				if (!success)
				{
					frmMessage frmMessage = new frmMessage("0071", "3", "1", "Thông báo");
					((Form)frmMessage).ShowDialog();
					((Control)mskTuNgay).Focus();
					return success;
				}
				DateTime dateTime2 = HTKKFormater.ConvertToDateTime(((Control)mskDenNgay).Text, out success);
				if (!success)
				{
					frmMessage frmMessage = new frmMessage("0071", "3", "1", "Thông báo");
					((Form)frmMessage).ShowDialog();
					((Control)mskDenNgay).Focus();
					return success;
				}
				if (dateTime > dateTime2)
				{
					frmMessage frmMessage = new frmMessage("0069", "3", "1", "Thông báo");
					((Form)frmMessage).ShowDialog();
					((Control)mskDenNgay).Focus();
					success = false;
					return success;
				}
				string[] array = ((Control)mskTuNgay).Text.Split(new char[1] { '/' });
				string[] array2 = ((Control)mskDenNgay).Text.Split(new char[1] { '/' });
				if (DateUtils.GetSoThang(array[1].PadLeft(2, '0') + array[2], array2[1].PadLeft(2, '0') + array2[2]) + 1 > 15 && menuID != "30" && menuID != "82" && menuID != "80" && menuID != "G7" && menuID != "C9" && menuID != "D6" && menuID != "Q4" && WinSettings.MenuID != "X9" && menuID != "Y1")
				{
					frmMessage frmMessage = new frmMessage("0068", "3", "1", "Thông báo");
					((Form)frmMessage).ShowDialog();
					((Control)mskDenNgay).Focus();
					success = false;
					return success;
				}
				if ((menuID.Equals("03") || menuID.Equals("83")) && (ConvertHelpers.ConvertToInt(((Control)mskYear).Text) < ConvertHelpers.ConvertToInt(array[2]) || ConvertHelpers.ConvertToInt(((Control)mskYear).Text) > ConvertHelpers.ConvertToInt(array2[2])))
				{
					frmMessage frmMessage = new frmMessage("0480", "3", "1", "Thông báo");
					((Form)frmMessage).ShowDialog();
					((Control)mskDenNgay).Focus();
					success = false;
					return success;
				}
				if (dateTime > DateTime.Now && !(menuID == "G7"))
				{
					frmMessage frmMessage = new frmMessage("0401", "3", "1", "Thông báo");
					((Form)frmMessage).ShowDialog();
					((Control)mskDenNgay).Focus();
					success = false;
					return success;
				}
				if (menuID.Equals("68") || menuID.Equals("14") || menuID.Equals("13") || (menuID.Equals("65") && ((Control)cbbQuy).Visible) || menuID.Equals("18") || menuID.Equals("38") || menuID.Equals("62"))
				{
					if (kieuKy.Equals("Q"))
					{
						if (((Control)mskTuNgay).Visible && ((Control)mskDenNgay).Visible)
						{
							dateTime = ConvertHelpers.ConvertToDateTime(((Control)mskTuNgay).Text, "dd/MM/yyyy");
							dateTime2 = ConvertHelpers.ConvertToDateTime(((Control)mskDenNgay).Text, "dd/MM/yyyy");
							DateTime ngayDauQuy = GeneralFunctions.GetNgayDauQuy(ConvertHelpers.ConvertToInt(((Control)cbbQuy).Text), ConvertHelpers.ConvertToInt(((Control)mskYear).Text), 1, 1);
							DateTime ngayCuoiQuy = GeneralFunctions.GetNgayCuoiQuy(ConvertHelpers.ConvertToInt(((Control)cbbQuy).Text), ConvertHelpers.ConvertToInt(((Control)mskYear).Text), 1, 1);
							if (DateTime.Compare(dateTime2, ngayCuoiQuy) > 0)
							{
								frmMessage frmMessage = new frmMessage("0255", "3", "1", "Thông báo");
								((Form)frmMessage).ShowDialog();
								((Control)mskDenNgay).Focus();
								success = false;
								return success;
							}
							if (DateTime.Compare(dateTime, ngayDauQuy) < 0)
							{
								frmMessage frmMessage = new frmMessage("0256", "3", "1", "Thông báo");
								((Form)frmMessage).ShowDialog();
								((Control)mskTuNgay).Focus();
								success = false;
								return success;
							}
							if (menuID.Equals("62"))
							{
								if (DateTime.Compare(dateTime, new DateTime(2014, 1, 1)) < 0)
								{
									frmMessage frmMessage = new frmMessage("0416", "3", "1", "Thông báo");
									((Form)frmMessage).ShowDialog();
									((Control)mskTuNgay).Focus();
									success = false;
									return success;
								}
							}
							else if (DateTime.Compare(dateTime, new DateTime(2011, 1, 1)) < 0)
							{
								frmMessage frmMessage = new frmMessage("0257", "3", "1", "Thông báo");
								((Form)frmMessage).ShowDialog();
								((Control)mskTuNgay).Focus();
								success = false;
								return success;
							}
						}
					}
					else if (kieuKy.Equals("M") && ((Control)mskTuNgay).Visible && ((Control)mskDenNgay).Visible)
					{
						dateTime = ConvertHelpers.ConvertToDateTime(((Control)mskTuNgay).Text, "dd/MM/yyyy");
						dateTime2 = ConvertHelpers.ConvertToDateTime(((Control)mskDenNgay).Text, "dd/MM/yyyy");
						DateTime ngayDauQuy = new DateTime(ConvertHelpers.ConvertToInt(((Control)mskYear).Text), ConvertHelpers.ConvertToInt(((Control)mskMonth).Text), 1);
						DateTime ngayCuoiQuy = ngayDauQuy.AddMonths(1).AddDays(-1.0);
						if (DateTime.Compare(dateTime2, ngayCuoiQuy) > 0)
						{
							frmMessage frmMessage = new frmMessage("0319", "3", "1", "Thông báo");
							((Form)frmMessage).ShowDialog();
							((Control)mskDenNgay).Focus();
							success = false;
							return success;
						}
						if (DateTime.Compare(dateTime, ngayDauQuy) < 0)
						{
							frmMessage frmMessage = new frmMessage("0318", "3", "1", "Thông báo");
							((Form)frmMessage).ShowDialog();
							((Control)mskDenNgay).Focus();
							success = false;
							return success;
						}
					}
				}
				if (menuID.Equals("80") || menuID.Equals("82") || menuID.Equals("03") || menuID.Equals("83") || menuID.Equals("C9") || menuID.Equals("D6") || menuID.Equals("X9") || menuID.Equals("Y1"))
				{
					DateTime t = new DateTime(dateTime2.Year, dateTime2.Month, 1);
					DateTime t2 = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1);
					if (DateTime.Compare(t2, t) < 0)
					{
						frmMessage frmMessage = new frmMessage("0247", "3", "1", "Thông báo");
						((Form)frmMessage).ShowDialog();
						((Control)mskDenNgay).Focus();
						success = false;
						return success;
					}
				}
			}
			return success;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return success;
		}
	}

	private bool CheckTuThangDenThang()
	{
		//IL_00af: Unknown result type (might be due to invalid IL or missing references)
		//IL_010b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0153: Unknown result type (might be due to invalid IL or missing references)
		//IL_01df: Unknown result type (might be due to invalid IL or missing references)
		//IL_0321: Unknown result type (might be due to invalid IL or missing references)
		//IL_0284: Unknown result type (might be due to invalid IL or missing references)
		//IL_03a5: Unknown result type (might be due to invalid IL or missing references)
		//IL_041f: Unknown result type (might be due to invalid IL or missing references)
		//IL_074d: Unknown result type (might be due to invalid IL or missing references)
		//IL_057f: Unknown result type (might be due to invalid IL or missing references)
		//IL_065a: Unknown result type (might be due to invalid IL or missing references)
		//IL_07da: Unknown result type (might be due to invalid IL or missing references)
		//IL_06fb: Unknown result type (might be due to invalid IL or missing references)
		bool success = true;
		try
		{
			if (((Control)pnTuThang).Visible && ((Control)pnDenThang).Visible && !chkTuQuyDenQuy.Checked && menuID != "99" && menuID != "H6" && menuID != "Q4")
			{
				DateTime dateTime = HTKKFormater.ConvertToDateTime("01/" + ((Control)mskTuThang).Text, out success);
				if (!success)
				{
					frmMessage frmMessage = new frmMessage("0018", "3", "1", "Thông báo");
					((Form)frmMessage).ShowDialog();
					((Control)mskTuThang).Focus();
					return success;
				}
				DateTime dateTime2 = HTKKFormater.ConvertToDateTime("01/" + ((Control)mskDenThang).Text, out success);
				if (!success)
				{
					frmMessage frmMessage = new frmMessage("0018", "3", "1", "Thông báo");
					((Form)frmMessage).ShowDialog();
					((Control)mskDenThang).Focus();
					return success;
				}
				if (dateTime > dateTime2)
				{
					frmMessage frmMessage = new frmMessage("0097", "3", "1", "Thông báo");
					((Form)frmMessage).ShowDialog();
					((Control)mskDenThang).Focus();
					success = false;
					return success;
				}
				if (menuID == "44")
				{
					string text = ((Control)mskTuThang).Text;
					string text2 = ((Control)mskDenThang).Text;
					string text3 = ((Control)mskYear).Text;
					if (!text.Contains(text3))
					{
						frmMessage frmMessage = new frmMessage("0336", "3", "1", "Thông báo");
						((Form)frmMessage).ShowDialog();
						((Control)mskTuThang).Focus();
						return false;
					}
					int num = (ConvertHelpers.ConvertToDateTime(text2, "MM/yyyy").Year - ConvertHelpers.ConvertToDateTime(text, "MM/yyyy").Year) * 12 + ConvertHelpers.ConvertToDateTime(text2, "MM/yyyy").Month - ConvertHelpers.ConvertToDateTime(text, "MM/yyyy").Month + 1;
					if (num > 15)
					{
						frmMessage frmMessage = new frmMessage("0335", "3", "1", "Thông báo");
						((Form)frmMessage).ShowDialog();
						((Control)mskDenThang).Focus();
						return false;
					}
				}
				if ("52,57,95,39,33,37,29".Contains(menuID))
				{
					if (DateUtils.GetSoThang(((Control)mskTuThang).Text.Replace("/", ""), ((Control)mskDenThang).Text.Replace("/", "")) + 1 > 12)
					{
						frmMessage frmMessage = new frmMessage("0339", "3", "1", "Thông báo");
						((Form)frmMessage).ShowDialog();
						((Control)mskDenThang).Focus();
						success = false;
						return success;
					}
					string[] array = ((Control)mskTuThang).Text.Split(new char[1] { '/' });
					if (array.Length == 2 && !array[1].Equals(((Control)mskYear).Text))
					{
						frmMessage frmMessage = new frmMessage("0336", "3", "1", "Thông báo");
						((Form)frmMessage).ShowDialog();
						((Control)mskTuThang).Focus();
						success = false;
						return success;
					}
					if (menuID == "33" && !dateTime2.Year.ToString().Equals(((Control)mskYear).Text))
					{
						frmMessage frmMessage = new frmMessage("0435", "3", "1", "Thông báo");
						((Form)frmMessage).ShowDialog();
						((Control)mskTuThang).Focus();
						success = false;
						return success;
					}
				}
				else if (((Control)mskTuThang).Visible && ((Control)mskDenThang).Visible)
				{
					if (menuID.Equals("35"))
					{
						DateTime ngayDauQuy = GeneralFunctions.GetNgayDauQuy(ConvertHelpers.ConvertToInt(((Control)cbbQuy).Text), ConvertHelpers.ConvertToInt(((Control)mskYear).Text), 1, 1);
						DateTime ngayCuoiQuy = GeneralFunctions.GetNgayCuoiQuy(ConvertHelpers.ConvertToInt(((Control)cbbQuy).Text), ConvertHelpers.ConvertToInt(((Control)mskYear).Text), 1, 1);
						DateTime t = ConvertHelpers.ConvertToDateTime("01/" + ((Control)mskTuThang).Text, "dd/MM/yyyy");
						DateTime t2 = ConvertHelpers.ConvertToDateTime("01/" + ((Control)mskDenThang).Text, "dd/MM/yyyy").AddMonths(1).AddDays(-1.0);
						if (DateTime.Compare(t, ngayDauQuy) < 0 || DateTime.Compare(t2, ngayCuoiQuy) > 0)
						{
							frmMessage frmMessage = new frmMessage("0365", "3", "1", "Thông báo");
							((Form)frmMessage).ShowDialog();
							((Control)mskDenThang).Focus();
							success = false;
							return success;
						}
					}
					else if (menuID != "A4" && menuID != "99" && menuID != "H6" && menuID != "Q4")
					{
						if (DateUtils.GetSoThang(((Control)mskTuThang).Text.Replace("/", ""), ((Control)mskDenThang).Text.Replace("/", "")) + 1 > 15)
						{
							frmMessage frmMessage = new frmMessage("0335", "3", "1", "Thông báo");
							((Form)frmMessage).ShowDialog();
							((Control)mskDenThang).Focus();
							success = false;
							return success;
						}
						string[] array = ((Control)mskTuThang).Text.Split(new char[1] { '/' });
						if (array.Length == 2 && ConvertHelpers.ConvertToInt(((Control)mskYear).Text) > 0 && !array[1].Equals(((Control)mskYear).Text))
						{
							frmMessage frmMessage = new frmMessage("0336", "3", "1", "Thông báo");
							((Form)frmMessage).ShowDialog();
							((Control)mskTuThang).Focus();
							success = false;
							return success;
						}
					}
				}
				if (dateTime > DateTime.Today)
				{
					frmMessage frmMessage = new frmMessage("0470", "3", "1", "Thông báo");
					((Form)frmMessage).ShowDialog();
					((Control)mskTuThang).Focus();
					success = false;
					return success;
				}
				if (dateTime2 > DateTime.Today && menuID != "28" && menuID != "G5" && menuID != "G7")
				{
					frmMessage frmMessage = new frmMessage("0471", "3", "1", "Thông báo");
					((Form)frmMessage).ShowDialog();
					((Control)mskDenThang).Focus();
					success = false;
					return success;
				}
				return success;
			}
			if (chkTuQuyDenQuy.Checked)
			{
			}
			return true;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return success;
		}
	}

	private bool CheckNgayThangNam()
	{
		//IL_008c: Unknown result type (might be due to invalid IL or missing references)
		//IL_010b: Unknown result type (might be due to invalid IL or missing references)
		//IL_018b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0251: Unknown result type (might be due to invalid IL or missing references)
		bool success = true;
		try
		{
			string text = "01";
			string text2 = "01";
			if (!((Control)pnNgay).Visible & !((Control)pnThang).Visible & !((Control)pnNam).Visible)
			{
				return true;
			}
			if (ConvertHelpers.ConvertToInt(((Control)mskYear).Text) < 2000)
			{
				frmMessage frmMessage = new frmMessage("0067", "3", "3", "Thông báo");
				((Form)frmMessage).ShowDialog();
				((Control)mskYear).Focus();
				return false;
			}
			if (((Control)mskMonth).Text.Trim().Length > 0 && ((Control)mskMonth).Visible && !HTKKFormater.IsMonth(((Control)mskMonth).Text))
			{
				frmMessage frmMessage = new frmMessage("0062", "3", "3", "Thông báo");
				((Form)frmMessage).ShowDialog();
				((Control)mskMonth).Focus();
				return false;
			}
			if (((Control)mskYear).Text.Trim().Length > 0 && ((Control)mskYear).Visible && !HTKKFormater.IsYear(((Control)mskYear).Text))
			{
				frmMessage frmMessage = new frmMessage("0062", "3", "3", "Thông báo");
				((Form)frmMessage).ShowDialog();
				((Control)mskYear).Focus();
				return false;
			}
			if (((Control)pnNgay).Visible)
			{
				text = ((Control)mskDay).Text;
			}
			if (((Control)pnThang).Visible)
			{
				text2 = ((Control)mskMonth).Text;
			}
			DateTime dateTime = HTKKFormater.ConvertToDateTime(text + '/' + text2 + '/' + ((Control)mskYear).Text, out success);
			if (!success)
			{
				frmMessage frmMessage = new frmMessage("0071", "3", "1", "Thông báo");
				((Form)frmMessage).ShowDialog();
				((Control)mskDay).Focus();
			}
			return success;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return success;
		}
	}

	private void btnDongY_Click(object sender, EventArgs e)
	{
		//IL_0088: Unknown result type (might be due to invalid IL or missing references)
		//IL_013c: Unknown result type (might be due to invalid IL or missing references)
		//IL_01f1: Unknown result type (might be due to invalid IL or missing references)
		//IL_02f8: Unknown result type (might be due to invalid IL or missing references)
		//IL_02a5: Unknown result type (might be due to invalid IL or missing references)
		//IL_04da: Unknown result type (might be due to invalid IL or missing references)
		//IL_03ad: Unknown result type (might be due to invalid IL or missing references)
		//IL_043a: Unknown result type (might be due to invalid IL or missing references)
		//IL_05a2: Unknown result type (might be due to invalid IL or missing references)
		//IL_06cd: Unknown result type (might be due to invalid IL or missing references)
		//IL_061d: Unknown result type (might be due to invalid IL or missing references)
		//IL_077b: Unknown result type (might be due to invalid IL or missing references)
		//IL_07ba: Unknown result type (might be due to invalid IL or missing references)
		//IL_0854: Unknown result type (might be due to invalid IL or missing references)
		//IL_08d7: Unknown result type (might be due to invalid IL or missing references)
		//IL_0933: Unknown result type (might be due to invalid IL or missing references)
		//IL_0b71: Unknown result type (might be due to invalid IL or missing references)
		//IL_0d33: Unknown result type (might be due to invalid IL or missing references)
		//IL_0bf3: Unknown result type (might be due to invalid IL or missing references)
		//IL_0a1e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0ed3: Unknown result type (might be due to invalid IL or missing references)
		//IL_0d98: Unknown result type (might be due to invalid IL or missing references)
		//IL_0c85: Unknown result type (might be due to invalid IL or missing references)
		//IL_0aa3: Unknown result type (might be due to invalid IL or missing references)
		//IL_0e2a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0fd5: Unknown result type (might be due to invalid IL or missing references)
		//IL_1031: Unknown result type (might be due to invalid IL or missing references)
		//IL_11b1: Unknown result type (might be due to invalid IL or missing references)
		//IL_120d: Unknown result type (might be due to invalid IL or missing references)
		//IL_1271: Unknown result type (might be due to invalid IL or missing references)
		//IL_1699: Unknown result type (might be due to invalid IL or missing references)
		//IL_1917: Unknown result type (might be due to invalid IL or missing references)
		//IL_179c: Unknown result type (might be due to invalid IL or missing references)
		//IL_1c4b: Unknown result type (might be due to invalid IL or missing references)
		//IL_1ae9: Unknown result type (might be due to invalid IL or missing references)
		//IL_189f: Unknown result type (might be due to invalid IL or missing references)
		//IL_1cd2: Unknown result type (might be due to invalid IL or missing references)
		//IL_1b6c: Unknown result type (might be due to invalid IL or missing references)
		//IL_19ae: Unknown result type (might be due to invalid IL or missing references)
		//IL_1d13: Unknown result type (might be due to invalid IL or missing references)
		//IL_1a29: Unknown result type (might be due to invalid IL or missing references)
		//IL_132c: Unknown result type (might be due to invalid IL or missing references)
		//IL_1bdb: Unknown result type (might be due to invalid IL or missing references)
		//IL_1ded: Unknown result type (might be due to invalid IL or missing references)
		//IL_1e46: Unknown result type (might be due to invalid IL or missing references)
		//IL_21f3: Unknown result type (might be due to invalid IL or missing references)
		//IL_1f1a: Unknown result type (might be due to invalid IL or missing references)
		//IL_2289: Unknown result type (might be due to invalid IL or missing references)
		//IL_20ec: Unknown result type (might be due to invalid IL or missing references)
		//IL_26f0: Unknown result type (might be due to invalid IL or missing references)
		//IL_2159: Unknown result type (might be due to invalid IL or missing references)
		//IL_1fb0: Unknown result type (might be due to invalid IL or missing references)
		//IL_233c: Unknown result type (might be due to invalid IL or missing references)
		//IL_2b9c: Unknown result type (might be due to invalid IL or missing references)
		//IL_27a4: Unknown result type (might be due to invalid IL or missing references)
		//IL_2465: Unknown result type (might be due to invalid IL or missing references)
		//IL_2038: Unknown result type (might be due to invalid IL or missing references)
		//IL_2c50: Unknown result type (might be due to invalid IL or missing references)
		//IL_28cd: Unknown result type (might be due to invalid IL or missing references)
		//IL_258e: Unknown result type (might be due to invalid IL or missing references)
		//IL_23b1: Unknown result type (might be due to invalid IL or missing references)
		//IL_2f38: Unknown result type (might be due to invalid IL or missing references)
		//IL_2d79: Unknown result type (might be due to invalid IL or missing references)
		//IL_29f7: Unknown result type (might be due to invalid IL or missing references)
		//IL_2819: Unknown result type (might be due to invalid IL or missing references)
		//IL_24da: Unknown result type (might be due to invalid IL or missing references)
		//IL_325f: Unknown result type (might be due to invalid IL or missing references)
		//IL_2cc5: Unknown result type (might be due to invalid IL or missing references)
		//IL_2943: Unknown result type (might be due to invalid IL or missing references)
		//IL_2665: Unknown result type (might be due to invalid IL or missing references)
		//IL_34f1: Unknown result type (might be due to invalid IL or missing references)
		//IL_2e50: Unknown result type (might be due to invalid IL or missing references)
		//IL_2acf: Unknown result type (might be due to invalid IL or missing references)
		//IL_3631: Unknown result type (might be due to invalid IL or missing references)
		//IL_3566: Unknown result type (might be due to invalid IL or missing references)
		//IL_2fec: Unknown result type (might be due to invalid IL or missing references)
		//IL_3965: Unknown result type (might be due to invalid IL or missing references)
		//IL_375a: Unknown result type (might be due to invalid IL or missing references)
		//IL_3313: Unknown result type (might be due to invalid IL or missing references)
		//IL_3115: Unknown result type (might be due to invalid IL or missing references)
		//IL_3c36: Unknown result type (might be due to invalid IL or missing references)
		//IL_36a6: Unknown result type (might be due to invalid IL or missing references)
		//IL_33c7: Unknown result type (might be due to invalid IL or missing references)
		//IL_3061: Unknown result type (might be due to invalid IL or missing references)
		//IL_3ca1: Unknown result type (might be due to invalid IL or missing references)
		//IL_38d8: Unknown result type (might be due to invalid IL or missing references)
		//IL_37cf: Unknown result type (might be due to invalid IL or missing references)
		//IL_318b: Unknown result type (might be due to invalid IL or missing references)
		//IL_3bca: Unknown result type (might be due to invalid IL or missing references)
		//IL_3a30: Unknown result type (might be due to invalid IL or missing references)
		//IL_343c: Unknown result type (might be due to invalid IL or missing references)
		//IL_3b16: Unknown result type (might be due to invalid IL or missing references)
		//IL_3fb4: Unknown result type (might be due to invalid IL or missing references)
		//IL_4004: Unknown result type (might be due to invalid IL or missing references)
		//IL_3da1: Unknown result type (might be due to invalid IL or missing references)
		//IL_40f8: Unknown result type (might be due to invalid IL or missing references)
		//IL_4170: Unknown result type (might be due to invalid IL or missing references)
		//IL_3ee4: Unknown result type (might be due to invalid IL or missing references)
		//IL_41b9: Unknown result type (might be due to invalid IL or missing references)
		//IL_3e67: Unknown result type (might be due to invalid IL or missing references)
		//IL_4256: Unknown result type (might be due to invalid IL or missing references)
		//IL_4358: Unknown result type (might be due to invalid IL or missing references)
		//IL_454c: Unknown result type (might be due to invalid IL or missing references)
		//IL_44bf: Unknown result type (might be due to invalid IL or missing references)
		//IL_4fd2: Unknown result type (might be due to invalid IL or missing references)
		//IL_50c9: Unknown result type (might be due to invalid IL or missing references)
		//IL_5062: Unknown result type (might be due to invalid IL or missing references)
		//IL_465d: Unknown result type (might be due to invalid IL or missing references)
		//IL_513c: Unknown result type (might be due to invalid IL or missing references)
		//IL_51ab: Unknown result type (might be due to invalid IL or missing references)
		//IL_53f1: Unknown result type (might be due to invalid IL or missing references)
		//IL_524c: Unknown result type (might be due to invalid IL or missing references)
		//IL_4cf4: Unknown result type (might be due to invalid IL or missing references)
		//IL_538b: Unknown result type (might be due to invalid IL or missing references)
		//IL_5475: Unknown result type (might be due to invalid IL or missing references)
		//IL_52af: Unknown result type (might be due to invalid IL or missing references)
		//IL_4913: Unknown result type (might be due to invalid IL or missing references)
		//IL_56aa: Unknown result type (might be due to invalid IL or missing references)
		//IL_4ab8: Unknown result type (might be due to invalid IL or missing references)
		//IL_5714: Unknown result type (might be due to invalid IL or missing references)
		//IL_5500: Unknown result type (might be due to invalid IL or missing references)
		//IL_5790: Unknown result type (might be due to invalid IL or missing references)
		//IL_5876: Unknown result type (might be due to invalid IL or missing references)
		//IL_5563: Unknown result type (might be due to invalid IL or missing references)
		//IL_4ed5: Unknown result type (might be due to invalid IL or missing references)
		//IL_591d: Unknown result type (might be due to invalid IL or missing references)
		//IL_5989: Unknown result type (might be due to invalid IL or missing references)
		//IL_5b87: Unknown result type (might be due to invalid IL or missing references)
		//IL_5bf0: Unknown result type (might be due to invalid IL or missing references)
		//IL_5a4d: Unknown result type (might be due to invalid IL or missing references)
		//IL_5c59: Unknown result type (might be due to invalid IL or missing references)
		//IL_5af1: Unknown result type (might be due to invalid IL or missing references)
		//IL_5aae: Unknown result type (might be due to invalid IL or missing references)
		//IL_5f81: Unknown result type (might be due to invalid IL or missing references)
		//IL_5d39: Unknown result type (might be due to invalid IL or missing references)
		//IL_5e19: Unknown result type (might be due to invalid IL or missing references)
		//IL_62db: Unknown result type (might be due to invalid IL or missing references)
		//IL_6080: Unknown result type (might be due to invalid IL or missing references)
		//IL_60e5: Unknown result type (might be due to invalid IL or missing references)
		//IL_63c2: Unknown result type (might be due to invalid IL or missing references)
		//IL_6141: Unknown result type (might be due to invalid IL or missing references)
		//IL_669e: Unknown result type (might be due to invalid IL or missing references)
		//IL_643b: Unknown result type (might be due to invalid IL or missing references)
		//IL_6726: Unknown result type (might be due to invalid IL or missing references)
		//IL_64ab: Unknown result type (might be due to invalid IL or missing references)
		//IL_685a: Unknown result type (might be due to invalid IL or missing references)
		//IL_678f: Unknown result type (might be due to invalid IL or missing references)
		//IL_6b45: Unknown result type (might be due to invalid IL or missing references)
		//IL_6903: Unknown result type (might be due to invalid IL or missing references)
		//IL_6bdd: Unknown result type (might be due to invalid IL or missing references)
		//IL_69b8: Unknown result type (might be due to invalid IL or missing references)
		//IL_6ecc: Unknown result type (might be due to invalid IL or missing references)
		//IL_6daa: Unknown result type (might be due to invalid IL or missing references)
		//IL_6a31: Unknown result type (might be due to invalid IL or missing references)
		//IL_6e1e: Unknown result type (might be due to invalid IL or missing references)
		//IL_6c5e: Unknown result type (might be due to invalid IL or missing references)
		//IL_6a97: Unknown result type (might be due to invalid IL or missing references)
		//IL_6e75: Unknown result type (might be due to invalid IL or missing references)
		//IL_71c4: Unknown result type (might be due to invalid IL or missing references)
		//IL_7182: Unknown result type (might be due to invalid IL or missing references)
		//IL_6d08: Unknown result type (might be due to invalid IL or missing references)
		//IL_7124: Unknown result type (might be due to invalid IL or missing references)
		//IL_7090: Unknown result type (might be due to invalid IL or missing references)
		try
		{
			((Control)btnDongY).Focus();
			if (!CheckValidpnVT1_1CheckBox())
			{
				return;
			}
			if ("~J7~J8~K2~K3~J9~".IndexOf("~" + menuID + "~") > -1 && ConvertHelpers.ConvertToInt(((Control)mskYear).Text) > 2025)
			{
				frmMessage frmMessage = new frmMessage("0671", "3", "", "Thông báo");
				((Form)frmMessage).ShowDialog();
				return;
			}
			if (menuID == "Y4")
			{
				DateTime dateTime = new DateTime(2026, 10, 1);
				if (chkTKLanPS.Checked)
				{
					DateTime dateTime2 = new DateTime(ConvertHelpers.ConvertToInt(((Control)mskYear).Text), ConvertHelpers.ConvertToInt(((Control)mskMonth).Text), ConvertHelpers.ConvertToInt(((Control)mskDay).Text));
					if (dateTime2 < dateTime)
					{
						frmMessage frmMessage2 = new frmMessage("0697", "3", "", "Thông báo");
						((Form)frmMessage2).ShowDialog();
						return;
					}
				}
			}
			if (menuID == "U2" || menuID == "U1")
			{
				DateTime dateTime = new DateTime(2026, 10, 1);
				DateTime dateTime2 = new DateTime(ConvertHelpers.ConvertToInt(((Control)mskYear).Text), ConvertHelpers.ConvertToInt(((Control)mskMonth).Text), ConvertHelpers.ConvertToInt(((Control)mskDay).Text));
				if (dateTime2 >= dateTime)
				{
					frmMessage frmMessage2 = new frmMessage("0701", "3", "", "Thông báo");
					((Form)frmMessage2).ShowDialog();
					return;
				}
			}
			if (menuID == "T5")
			{
				DateTime dateTime = new DateTime(2026, 10, 1);
				if (chkTKLanPS.Checked)
				{
					DateTime dateTime2 = new DateTime(ConvertHelpers.ConvertToInt(((Control)mskYear).Text), ConvertHelpers.ConvertToInt(((Control)mskMonth).Text), ConvertHelpers.ConvertToInt(((Control)mskDay).Text));
					if (dateTime2 >= dateTime)
					{
						frmMessage frmMessage2 = new frmMessage("0701", "3", "", "Thông báo");
						((Form)frmMessage2).ShowDialog();
						return;
					}
				}
				else
				{
					DateTime dateTime3 = ConvertHelpers.ConvertToDateTime(((Control)mskDenNgay).Text, "dd/MM/yyyy");
					if (dateTime3 >= dateTime)
					{
						frmMessage frmMessage2 = new frmMessage("0701", "3", "", "Thông báo");
						((Form)frmMessage2).ShowDialog();
						return;
					}
				}
			}
			if (menuID == "R5")
			{
				DateTime dateTime4 = new DateTime(2026, 2, 1);
				DateTime dateTime = new DateTime(2026, 2, 14);
				if (chkTkhaiThang.Checked)
				{
					DateTime dateTime2 = new DateTime(ConvertHelpers.ConvertToInt(((Control)mskYear).Text), ConvertHelpers.ConvertToInt(((Control)mskMonth).Text), 1);
					if (dateTime2 < dateTime4)
					{
						frmMessage frmMessage2 = new frmMessage("0672", "3", "", "Thông báo");
						((Form)frmMessage2).ShowDialog();
						return;
					}
				}
				else if (chkTKLanPS.Checked)
				{
					DateTime dateTime2 = new DateTime(ConvertHelpers.ConvertToInt(((Control)mskYear).Text), ConvertHelpers.ConvertToInt(((Control)mskMonth).Text), ConvertHelpers.ConvertToInt(((Control)mskDay).Text));
					if (dateTime2 < dateTime)
					{
						frmMessage frmMessage2 = new frmMessage("0673", "3", "", "Thông báo");
						((Form)frmMessage2).ShowDialog();
						return;
					}
				}
			}
			if (menuID == "T4")
			{
				DateTime dateTime2 = new DateTime(ConvertHelpers.ConvertToInt(((Control)mskYear).Text), ConvertHelpers.ConvertToInt(((Control)mskMonth).Text), ConvertHelpers.ConvertToInt(((Control)mskDay).Text));
				DateTime dateTime = ConvertHelpers.ConvertToDateTime("17/03/2026", "dd/MM/yyyy");
				if (dateTime2 >= dateTime)
				{
					frmMessage frmMessage2 = new frmMessage("0674", "3", "", "Thông báo");
					((Form)frmMessage2).ShowDialog();
					return;
				}
			}
			if (menuID == "X8" || menuID == "Y6")
			{
				DateTime dateTime4 = new DateTime(2026, 7, 1);
				DateTime dateTime5 = new DateTime(2026, 3, 1);
				if (chkTkhaiThang.Checked)
				{
					DateTime dateTime2 = new DateTime(ConvertHelpers.ConvertToInt(((Control)mskYear).Text), ConvertHelpers.ConvertToInt(((Control)mskMonth).Text), 1);
					if (dateTime2 < dateTime4)
					{
						frmMessage frmMessage2 = new frmMessage("0703", "3", "", "Thông báo");
						((Form)frmMessage2).ShowDialog();
						return;
					}
				}
				else if (chkTKQuy.Checked)
				{
					DateTime dateTime2 = new DateTime(ConvertHelpers.ConvertToInt(((Control)mskYear).Text), ConvertHelpers.ConvertToInt(((Control)cbbQuy).Text), 1);
					if (dateTime2 < dateTime5)
					{
						frmMessage frmMessage2 = new frmMessage("0704", "3", "", "Thông báo");
						((Form)frmMessage2).ShowDialog();
						return;
					}
				}
			}
			if (menuID == "R4" && ((chkTKQuy.Checked && cbbQuy.SelectedItem.ToIntOrDefault() == 4) || (chkTkhaiThang.Checked && ((Control)mskMonth).Text == "12")) && !(bool)axfpsPhuLuc.Sheets[0].Cells[0, 0].Value)
			{
				frmMessage frmMessage = new frmMessage("0368", "3", "", "Thông báo");
				((Form)frmMessage).ShowDialog();
				return;
			}
			if (menuID == "X8" && ((chkTKQuy.Checked && cbbQuy.SelectedItem.ToIntOrDefault() == 4) || (chkTkhaiThang.Checked && ((Control)mskMonth).Text == "12")) && !(bool)axfpsPhuLuc.Sheets[0].Cells[0, 0].Value)
			{
				frmMessage frmMessage = new frmMessage("0368", "3", "", "Thông báo");
				((Form)frmMessage).ShowDialog();
				return;
			}
			if (menuID == "95")
			{
				frmMessage frmMessage = new frmMessage("0651", "3", "", "Thông báo");
				((Form)frmMessage).ShowDialog();
				((Control)btnDongY).Focus();
				return;
			}
			if (menuID == "P2")
			{
				DateTime t = new DateTime(2025, 6, 30);
				if (DateTime.Compare(t, new DateTime(ConvertHelpers.ConvertToInt(((Control)mskYear).Text), ConvertHelpers.ConvertToInt(((Control)mskMonth).Text), 1)) > 0)
				{
					frmMessage frmMessage2 = new frmMessage("0653", "3", "3", "Thông báo");
					((Form)frmMessage2).ShowDialog();
					((Control)mskMonth).Focus();
					return;
				}
			}
			if (OptBosung.Checked)
			{
				if (HTKKFormater.isNullString(((Control)mskNgayLapKHBS).Text.Replace("/", "").Replace("_", "")))
				{
					frmMessage frmMessage = new frmMessage("0071", "3", "", "Thông báo");
					((Form)frmMessage).ShowDialog();
					((Control)mskNgayLapKHBS).Focus();
					return;
				}
				if (ConvertHelpers.ConvertToDateTime(((Control)mskNgayLapKHBS).Text, "dd/MM/yyyy") > DateTime.Now)
				{
					frmMessage frmMessage = new frmMessage("0224", "3", "", "Thông báo");
					((Form)frmMessage).ShowDialog();
					((Control)mskNgayLapKHBS).Focus();
					return;
				}
			}
			if (menuID.Equals("Q8") && Convert.ToBoolean(axfpsPhuLuc.Sheets[0].Cells[1, 0].Value))
			{
				if (WinSettings.TTPeriodInfo.KieuKy == "M")
				{
					if (Convert.ToInt32(WinSettings.TTPeriodInfo.Nam) != 2026 || (Convert.ToInt32(WinSettings.TTPeriodInfo.Thang) != 1 && Convert.ToInt32(WinSettings.TTPeriodInfo.Thang) != 2))
					{
						frmMessage frmMessage = new frmMessage("0667", "3", "3", "Thông báo");
						((Form)frmMessage).ShowDialog();
						return;
					}
				}
				else if (WinSettings.TTPeriodInfo.KieuKy == "Q" && (Convert.ToInt32(WinSettings.TTPeriodInfo.Nam) != 2026 || Convert.ToInt32(WinSettings.TTPeriodInfo.Quy) != 1))
				{
					frmMessage frmMessage = new frmMessage("0668", "3", "3", "Thông báo");
					((Form)frmMessage).ShowDialog();
					return;
				}
			}
			if (menuID == "Q8")
			{
				DateTime dateTime4 = new DateTime(2026, 5, 1);
				DateTime dateTime5 = new DateTime(2026, 1, 1);
				DateTime dateTime6 = new DateTime(2026, 4, 30);
				if (WinSettings.TTPeriodInfo.KieuKy == "M")
				{
					DateTime dateTime2 = new DateTime(ConvertHelpers.ConvertToInt(((Control)mskYear).Text), ConvertHelpers.ConvertToInt(((Control)mskMonth).Text), 1);
					if (dateTime2 > dateTime4)
					{
						frmMessage frmMessage2 = new frmMessage("0671", "3", "", "Thông báo");
						((Form)frmMessage2).ShowDialog();
						return;
					}
				}
				if (WinSettings.TTPeriodInfo.KieuKy == "Q")
				{
					DateTime dateTime2 = new DateTime(ConvertHelpers.ConvertToInt(((Control)mskYear).Text), ConvertHelpers.ConvertToInt(((Control)cbbQuy).Text), 1);
					if (dateTime2 > dateTime5)
					{
						frmMessage frmMessage = new frmMessage("0671", "3", "3", "Thông báo");
						((Form)frmMessage).ShowDialog();
						return;
					}
				}
				if (WinSettings.TTPeriodInfo.KieuKy == "D")
				{
					DateTime dateTime2 = new DateTime(ConvertHelpers.ConvertToInt(((Control)mskYear).Text), ConvertHelpers.ConvertToInt(((Control)mskMonth).Text), ConvertHelpers.ConvertToInt(((Control)mskDay).Text));
					if (dateTime2 > dateTime6)
					{
						frmMessage frmMessage2 = new frmMessage("0671", "3", "", "Thông báo");
						((Form)frmMessage2).ShowDialog();
						return;
					}
				}
			}
			if (menuID == "U5")
			{
				DateTime dateTime4 = new DateTime(2026, 1, 1);
				if (WinSettings.TTPeriodInfo.KieuKy == "M")
				{
					DateTime dateTime2 = new DateTime(ConvertHelpers.ConvertToInt(((Control)mskYear).Text), ConvertHelpers.ConvertToInt(((Control)mskMonth).Text), 1);
					if (dateTime2 < dateTime4)
					{
						frmMessage frmMessage2 = new frmMessage("0677", "3", "", "Thông báo");
						((Form)frmMessage2).ShowDialog();
						return;
					}
				}
				if (WinSettings.TTPeriodInfo.KieuKy == "Q" && Convert.ToInt32(WinSettings.TTPeriodInfo.Nam) < 2026)
				{
					frmMessage frmMessage = new frmMessage("0678", "3", "3", "Thông báo");
					((Form)frmMessage).ShowDialog();
					return;
				}
				if (WinSettings.TTPeriodInfo.KieuKy == "D")
				{
					DateTime dateTime2 = new DateTime(ConvertHelpers.ConvertToInt(((Control)mskYear).Text), ConvertHelpers.ConvertToInt(((Control)mskMonth).Text), ConvertHelpers.ConvertToInt(((Control)mskDay).Text));
					if (dateTime2 < dateTime4)
					{
						frmMessage frmMessage2 = new frmMessage("0683", "3", "", "Thông báo");
						((Form)frmMessage2).ShowDialog();
						return;
					}
				}
			}
			if (menuID == "B8" && !HTKKFormater.isNullString(((Control)mskNgayTb).Text.Replace("/", "").Replace("_", "")) && ConvertHelpers.ConvertToDateTime(((Control)mskNgayTb).Text, "dd/MM/yyyy") > DateTime.Now)
			{
				frmMessage frmMessage = new frmMessage("0642", "3", "", "Thông báo");
				((Form)frmMessage).ShowDialog();
				((Control)mskNgayTb).Focus();
				return;
			}
			if (WinSettings.TTPeriodInfo.dicVar.ContainsKey("CoSoMoiThanhLap"))
			{
				WinSettings.TTPeriodInfo.dicVar.Remove("CoSoMoiThanhLap");
			}
			if (WinSettings.TTPeriodInfo.dicVar.ContainsKey("NgayThanhLap"))
			{
				WinSettings.TTPeriodInfo.dicVar.Remove("NgayThanhLap");
			}
			if (chk_CoSoThanhLap.Checked)
			{
				WinSettings.TTPeriodInfo.dicVar.Add("CoSoMoiThanhLap", "true");
				if (HTKKFormater.isNullString(((Control)mskNgayPS).Text.Replace("/", "").Replace("_", "")))
				{
					frmMessage frmMessage = new frmMessage("0071", "3", "", "Thông báo");
					((Form)frmMessage).ShowDialog();
					((Control)mskNgayPS).Focus();
					return;
				}
				if (ConvertHelpers.ConvertToDateTime(((Control)mskNgayPS).Text, "dd/MM/yyyy") > DateTime.Now)
				{
					frmMessage frmMessage = new frmMessage("0223", "3", "", "Thông báo");
					((Form)frmMessage).ShowDialog();
					((Control)mskNgayPS).Focus();
					return;
				}
				WinSettings.TTPeriodInfo.dicVar.Add("NgayThanhLap", ((Control)mskNgayPS).Text);
			}
			if (!WinSettings.KieuKy.Equals("0") && !ValidateBaseControl())
			{
				return;
			}
			GetPeriodInfo(WinSettings.TTPeriodInfo);
			if ("N7~".IndexOf(WinSettings.MenuID) >= 0)
			{
				WinSettings.TTPeriodInfo.NgayLapKHBS = ((Control)mskNgayLapKHBS1).Text.Trim();
				WinSettings.TTPeriodInfo.SoLan = ((Control)txtLan1).Text.Trim();
				WinSettings.TTPeriodInfo.LoaiToKhai = ((((ListControl)cboLoaiToKhai).SelectedIndex == 0) ? "C" : ((((ListControl)cboLoaiToKhai).SelectedIndex == 1) ? "T" : ""));
				if (WinSettings.TTPeriodInfo.LoaiToKhai != "C" && ((Control)mskNgayLapKHBS1).Visible)
				{
					if (HTKKFormater.isNullString(((Control)mskNgayLapKHBS1).Text.Replace("/", "").Replace("_", "")))
					{
						frmMessage frmMessage = new frmMessage("0071", "3", "", "Thông báo");
						((Form)frmMessage).ShowDialog();
						((Control)mskNgayLapKHBS1).Focus();
						return;
					}
					if (ConvertHelpers.ConvertToDateTime(((Control)mskNgayLapKHBS1).Text, "dd/MM/yyyy") > DateTime.Now)
					{
						frmMessage frmMessage = new frmMessage("0224", "3", "", "Thông báo");
						((Form)frmMessage).ShowDialog();
						((Control)mskNgayLapKHBS1).Focus();
						return;
					}
				}
				if (((Control)txtLan1).Text.Trim().Length == 0 && ((Control)txtLan1).Visible)
				{
					frmMessage frmMessage = new frmMessage("0604", "3", "1", "Thông báo");
					((Form)frmMessage).ShowDialog();
					((Control)txtLan).Focus();
					return;
				}
				if (((Control)txtLan1).Text.Trim().Length > 0 && ((Control)txtLan1).Visible && (!GeneralFunctions.IsNumeric(((Control)txtLan1).Text.Trim()) || (GeneralFunctions.IsNumeric(((Control)txtLan1).Text.Trim()) && Convert.ToInt32(((Control)txtLan1).Text.Trim()) <= 0)))
				{
					frmMessage frmMessage = new frmMessage("0605", "3", "1", "Thông báo");
					((Form)frmMessage).ShowDialog();
					((Control)txtLan).Focus();
					return;
				}
			}
			if (menuID == "21" && WinSettings.TTPeriodInfo.LoaiToKhai == "B")
			{
				string valueInXml = GeneralFunctions.GetValueInXml("../DataFiles/" + WinSettings.HTKKHeaderInfo.mst + "/03_TNDN_DK_" + WinSettings.TTPeriodInfo.kyKKhai.Replace("/", "") + "_" + WinSettings.TTPeriodInfo.VersionXML.Replace(".", "") + ".xml", "P_53");
				int num = 0;
				DateTime result = DateTime.MinValue;
				DateTime result2 = DateTime.MinValue;
				DateTime.TryParseExact(WinSettings.TTPeriodInfo.kyKKhai, "dd/MM/yyyy", CultureInfo.CurrentCulture, DateTimeStyles.None, out result);
				DateTime.TryParseExact(WinSettings.TTPeriodInfo.NgayLapKHBS, "dd/MM/yyyy", CultureInfo.CurrentCulture, DateTimeStyles.None, out result2);
				num = ((valueInXml == WinSettings.HTKKHeaderInfo.mst) ? (WinSettings.TTPeriodInfo.SoNgayNopChamKHBS = (int)(DateTime.ParseExact(WinSettings.TTPeriodInfo.kyKKhai, "dd/MM/yyyy", CultureInfo.CurrentCulture, DateTimeStyles.None).AddMonths(6) - DateTime.ParseExact(WinSettings.TTPeriodInfo.kyKKhai, "dd/MM/yyyy", CultureInfo.CurrentCulture, DateTimeStyles.None)).TotalDays) : (WinSettings.TTPeriodInfo.SoNgayNopChamKHBS = 10));
				DateTime dateTime7 = result.AddDays(num);
				if (dateTime7.DayOfWeek == DayOfWeek.Saturday)
				{
					dateTime7 = dateTime7.AddDays(2.0);
				}
				else if (dateTime7.DayOfWeek == DayOfWeek.Sunday)
				{
					dateTime7 = dateTime7.AddDays(1.0);
				}
			}
			if (menuID == "03")
			{
				if (WinSettings.TTPeriodInfo.dicVar.ContainsKey("NGAY_BAT_DAU_NAM_TAI_CHINH"))
				{
					WinSettings.TTPeriodInfo.dicVar.Remove("NGAY_BAT_DAU_NAM_TAI_CHINH");
				}
				WinSettings.TTPeriodInfo.dicVar.Add("NGAY_BAT_DAU_NAM_TAI_CHINH", WinSettings.HTKKHeaderInfo.NgayBatDauNamTaiChinh);
				if (ConvertHelpers.ConvertToInt(((Control)mskYear).Text) < 2017 && ((bool)axfpsPhuLuc.Sheets[0].Cells[18, 0].Value || (bool)axfpsPhuLuc.Sheets[0].Cells[15, 0].Value || (bool)axfpsPhuLuc.Sheets[0].Cells[16, 0].Value || (bool)axfpsPhuLuc.Sheets[0].Cells[17, 0].Value))
				{
					frmMessage frmMessage2 = new frmMessage("0482", "3", "", "Thông báo");
					((Form)frmMessage2).ShowDialog();
					return;
				}
				if (ConvertHelpers.ConvertToInt(((Control)mskYear).Text) > 2019 && ((bool)axfpsPhuLuc.Sheets[0].Cells[18, 0].Value || (bool)axfpsPhuLuc.Sheets[0].Cells[15, 0].Value || (bool)axfpsPhuLuc.Sheets[0].Cells[16, 0].Value || (bool)axfpsPhuLuc.Sheets[0].Cells[17, 0].Value))
				{
					frmMessage frmMessage2 = new frmMessage("0515", "3", "", "Thông báo");
					((Form)frmMessage2).ShowDialog();
					return;
				}
				if (ConvertHelpers.ConvertToInt(((Control)mskYear).Text) < 2019 && ((bool)axfpsPhuLuc.Sheets[0].Cells[19, 0].Value || (bool)axfpsPhuLuc.Sheets[0].Cells[20, 0].Value || (bool)axfpsPhuLuc.Sheets[0].Cells[21, 0].Value || (bool)axfpsPhuLuc.Sheets[0].Cells[22, 0].Value))
				{
					frmMessage frmMessage2 = new frmMessage("0506", "3", "", "Thông báo");
					((Form)frmMessage2).ShowDialog();
					return;
				}
				if (ConvertHelpers.ConvertToInt(((Control)mskYear).Text) >= 2017 && (bool)axfpsPhuLuc.Sheets[0].Cells[11, 0].Value)
				{
					frmMessage frmMessage2 = new frmMessage("0483", "3", "", "Thông báo");
					((Form)frmMessage2).ShowDialog();
					return;
				}
				if ((ConvertHelpers.ConvertToInt(((Control)mskYear).Text) < 2020 || ConvertHelpers.ConvertToInt(((Control)mskYear).Text) > 2020) && (bool)axfpsPhuLuc.Sheets[0].Cells[23, 0].Value)
				{
					frmMessage frmMessage2 = new frmMessage("0503", "3", "", "Thông báo");
					((Form)frmMessage2).ShowDialog();
					return;
				}
				if (ConvertHelpers.ConvertToInt(((Control)mskYear).Text) < 2019 && (bool)axfpsPhuLuc.Sheets[0].Cells[24, 0].Value)
				{
					frmMessage frmMessage2 = new frmMessage("0528", "3", "", "Thông báo");
					((Form)frmMessage2).ShowDialog();
					return;
				}
			}
			if (!string.IsNullOrEmpty(WinSettings.TTPeriodInfo.MaQuyetToan) && !CheckTruongHopQuyetToan())
			{
				return;
			}
			if ((menuID == "71" || menuID == "D4" || WinSettings.MenuID.Equals("Z3")) && ConvertHelpers.ConvertToInt(((Control)mskYear).Text) < 2021 && chkThuHo.Checked)
			{
				frmMessage frmMessage2 = new frmMessage("0516", "3", "", "Thông báo");
				((Form)frmMessage2).ShowDialog();
				return;
			}
			if (menuID == "02")
			{
				if (ConvertHelpers.ConvertToInt(((Control)mskYear).Text) >= 2021 && ((Control)txtMaHSDNHT_DADT).Text.Trim().Equals(""))
				{
					frmMessage frmMessage2 = new frmMessage("0517", "3", "", "Thông báo");
					((Form)frmMessage2).ShowDialog();
					((Control)txtMaHSDNHT_DADT).Focus();
					return;
				}
				if (ConvertHelpers.ConvertToInt(((Control)mskYear).Text) < 2021 && !((Control)txtMaHSDNHT_DADT).Text.Trim().Equals(""))
				{
					frmMessage frmMessage2 = new frmMessage("0518", "3", "", "Thông báo");
					((Form)frmMessage2).ShowDialog();
					((Control)txtMaHSDNHT_DADT).Focus();
					return;
				}
			}
			if (menuID == "D2" && ((Control)cbMaHSDNHT_DADT).Text.Trim().Equals(""))
			{
				frmMessage frmMessage2 = new frmMessage("0597", "3", "", "Thông báo");
				((Form)frmMessage2).ShowDialog();
				((Control)cbMaHSDNHT_DADT).Focus();
				return;
			}
			if (menuID == "U7")
			{
				string text = "~0100100417~0100231226998~0104009085~0106356736~2300130532~";
				if (text.IndexOf("~" + WinSettings.HTKKHeaderInfo.mst + "~") <= -1)
				{
					frmMessage frmMessage2 = new frmMessage("0689", "3", "", "Thông báo");
					((Form)frmMessage2).ShowDialog();
					return;
				}
				if (!WinSettings.TTPeriodInfo.maCQTQuanLy.Equals("82700"))
				{
					frmMessage frmMessage2 = new frmMessage("0686", "3", "", "Thông báo");
					((Form)frmMessage2).ShowDialog();
					((Control)cbMaHSDNHT_DADT).Focus();
					return;
				}
			}
			if ((menuID == "D2" || menuID == "99") && !((Control)cbMaHSDNHT_DADT).Text.Trim().Equals(""))
			{
				if (WinSettings.TTPeriodInfo.maCQTQuanLy.Equals("82500") || WinSettings.TTPeriodInfo.maCQTQuanLy.Equals("82700") || !WinSettings.TTPeriodInfo.maCQTQuanLy.Substring(3, 2).Equals("00"))
				{
					frmMessage frmMessage2 = new frmMessage("0685", "3", "", "Thông báo");
					((Form)frmMessage2).ShowDialog();
					((Control)cbMaHSDNHT_DADT).Focus();
					return;
				}
				if (!getMaCQTQuanLyDADT(((Control)cbMaHSDNHT_DADT).Text).Equals(WinSettings.TTPeriodInfo.maCQTQuanLy))
				{
					frmMessage frmMessage2 = new frmMessage("0635", "3", "", "Thông báo");
					((Form)frmMessage2).ShowDialog();
					((Control)cbMaHSDNHT_DADT).Focus();
					return;
				}
			}
			if (menuID == "73" || menuID == "55")
			{
				if ((ConvertHelpers.ConvertToInt(((Control)mskYear).Text) < 2020 || ConvertHelpers.ConvertToInt(((Control)mskYear).Text) > 2020) && (bool)axfpsPhuLuc.Sheets[0].Cells[0, 0].Value)
				{
					frmMessage frmMessage2 = new frmMessage("0503", "3", "", "Thông báo");
					((Form)frmMessage2).ShowDialog();
					return;
				}
				if (ConvertHelpers.ConvertToInt(((Control)mskYear).Text) < 2021 || ConvertHelpers.ConvertToInt(((Control)mskYear).Text) > 2021)
				{
					if ((bool)axfpsPhuLuc.Sheets[0].Cells[1, 0].Value)
					{
						frmMessage frmMessage2 = new frmMessage("0541", "3", "", "Thông báo");
						((Form)frmMessage2).ShowDialog();
						return;
					}
				}
				else if (menuID == "55" && ckNam.Checked && (bool)axfpsPhuLuc.Sheets[0].Cells[1, 0].Value)
				{
					frmMessage frmMessage2 = new frmMessage("0542", "3", "", "Thông báo");
					((Form)frmMessage2).ShowDialog();
					return;
				}
			}
			if (menuID == "T5")
			{
				if (ConvertHelpers.ConvertToInt(((Control)mskYear).Text) < 2021 || ConvertHelpers.ConvertToInt(((Control)mskYear).Text) > 2021)
				{
					if ((bool)axfpsPhuLuc.Sheets[0].Cells[0, 0].Value)
					{
						frmMessage frmMessage2 = new frmMessage("0541", "3", "", "Thông báo");
						((Form)frmMessage2).ShowDialog();
						return;
					}
				}
				else if (!ckNam.Checked && (bool)axfpsPhuLuc.Sheets[0].Cells[0, 0].Value)
				{
					frmMessage frmMessage2 = new frmMessage("0616", "3", "", "Thông báo");
					((Form)frmMessage2).ShowDialog();
					return;
				}
			}
			if (menuID == "H6" && !string.IsNullOrEmpty(((Control)mskTuThang).Text) && DateTime.ParseExact("01/" + ((Control)mskTuThang).Text, "dd/MM/yyyy", CultureInfo.CurrentCulture).Year < 2022)
			{
				frmMessage frmMessage2 = new frmMessage("0577", "3", "", "Thông báo");
				((Form)frmMessage2).ShowDialog();
				((Control)mskTuThang).Focus();
				return;
			}
			if (menuID == "28")
			{
				if (ckNam.Checked)
				{
					if ((bool)axfpsPhuLuc.Sheets[0].Cells[1, 0].Value)
					{
						frmMessage frmMessage2 = new frmMessage("0543", "3", "", "Thông báo");
						((Form)frmMessage2).ShowDialog();
						return;
					}
				}
				else if (chkTKQuy.Checked)
				{
					if (ConvertHelpers.ConvertToInt(((Control)mskYear).Text) < 2021 || ConvertHelpers.ConvertToInt(((Control)mskYear).Text) > 2021)
					{
						if ((bool)axfpsPhuLuc.Sheets[0].Cells[1, 0].Value)
						{
							frmMessage frmMessage2 = new frmMessage("0544", "3", "", "Thông báo");
							((Form)frmMessage2).ShowDialog();
							return;
						}
					}
					else if ((bool)axfpsPhuLuc.Sheets[0].Cells[1, 0].Value && ConvertHelpers.ConvertToInt(((Control)cbbQuy).Text) < 3)
					{
						frmMessage frmMessage2 = new frmMessage("0544", "3", "", "Thông báo");
						((Form)frmMessage2).ShowDialog();
						return;
					}
				}
				else if (chkTkhaiThang.Checked)
				{
					if (ConvertHelpers.ConvertToInt(((Control)mskYear).Text) < 2021 || ConvertHelpers.ConvertToInt(((Control)mskYear).Text) > 2021)
					{
						if ((bool)axfpsPhuLuc.Sheets[0].Cells[1, 0].Value)
						{
							frmMessage frmMessage2 = new frmMessage("0545", "3", "", "Thông báo");
							((Form)frmMessage2).ShowDialog();
							return;
						}
					}
					else if ((bool)axfpsPhuLuc.Sheets[0].Cells[1, 0].Value && ConvertHelpers.ConvertToInt(((Control)mskMonth).Text) < 7)
					{
						frmMessage frmMessage2 = new frmMessage("0545", "3", "", "Thông báo");
						((Form)frmMessage2).ShowDialog();
						return;
					}
				}
				else if (chkTKLanPS.Checked)
				{
					if (ConvertHelpers.ConvertToInt(((Control)mskYear).Text) < 2021 || ConvertHelpers.ConvertToInt(((Control)mskYear).Text) > 2021)
					{
						if ((bool)axfpsPhuLuc.Sheets[0].Cells[1, 0].Value)
						{
							frmMessage frmMessage2 = new frmMessage("0546", "3", "", "Thông báo");
							((Form)frmMessage2).ShowDialog();
							return;
						}
					}
					else
					{
						DateTime dateTime8 = new DateTime(2021, 7, 1);
						DateTime dateTime9 = new DateTime(2021, 12, 31);
						DateTime dateTime10 = new DateTime(ConvertHelpers.ConvertToInt(((Control)mskYear).Text), ConvertHelpers.ConvertToInt(((Control)mskMonth).Text), ConvertHelpers.ConvertToInt(((Control)mskDay).Text));
						if ((bool)axfpsPhuLuc.Sheets[0].Cells[1, 0].Value && (dateTime10 < dateTime8 || dateTime10 > dateTime9))
						{
							frmMessage frmMessage2 = new frmMessage("0546", "3", "", "Thông báo");
							((Form)frmMessage2).ShowDialog();
							return;
						}
					}
				}
			}
			if (menuID == "71")
			{
				if (chkThuHo.Checked)
				{
					if ((bool)axfpsPhuLuc.Sheets[0].Cells[0, 0].Value)
					{
						frmMessage frmMessage2 = new frmMessage("0552", "3", "", "Thông báo");
						((Form)frmMessage2).ShowDialog();
						return;
					}
				}
				else if (chkTKQuy.Checked)
				{
					if (ConvertHelpers.ConvertToInt(((Control)mskYear).Text) < 2021 || ConvertHelpers.ConvertToInt(((Control)mskYear).Text) > 2021)
					{
						if ((bool)axfpsPhuLuc.Sheets[0].Cells[0, 0].Value)
						{
							frmMessage frmMessage2 = new frmMessage("0549", "3", "", "Thông báo");
							((Form)frmMessage2).ShowDialog();
							return;
						}
					}
					else if ((bool)axfpsPhuLuc.Sheets[0].Cells[0, 0].Value && ConvertHelpers.ConvertToInt(((Control)cbbQuy).Text) < 4)
					{
						frmMessage frmMessage2 = new frmMessage("0549", "3", "", "Thông báo");
						((Form)frmMessage2).ShowDialog();
						return;
					}
				}
				else if (chkTkhaiThang.Checked)
				{
					if (ConvertHelpers.ConvertToInt(((Control)mskYear).Text) < 2021 || ConvertHelpers.ConvertToInt(((Control)mskYear).Text) > 2021)
					{
						if ((bool)axfpsPhuLuc.Sheets[0].Cells[0, 0].Value)
						{
							frmMessage frmMessage2 = new frmMessage("0550", "3", "", "Thông báo");
							((Form)frmMessage2).ShowDialog();
							return;
						}
					}
					else if ((bool)axfpsPhuLuc.Sheets[0].Cells[0, 0].Value && ConvertHelpers.ConvertToInt(((Control)mskMonth).Text) < 11)
					{
						frmMessage frmMessage2 = new frmMessage("0550", "3", "", "Thông báo");
						((Form)frmMessage2).ShowDialog();
						return;
					}
				}
				else if (chkTKLanPS.Checked)
				{
					if (ConvertHelpers.ConvertToInt(((Control)mskYear).Text) < 2021 || ConvertHelpers.ConvertToInt(((Control)mskYear).Text) > 2021)
					{
						if ((bool)axfpsPhuLuc.Sheets[0].Cells[0, 0].Value)
						{
							frmMessage frmMessage2 = new frmMessage("0551", "3", "", "Thông báo");
							((Form)frmMessage2).ShowDialog();
							return;
						}
					}
					else
					{
						DateTime dateTime8 = new DateTime(2021, 11, 1);
						DateTime dateTime9 = new DateTime(2021, 12, 31);
						DateTime dateTime10 = new DateTime(ConvertHelpers.ConvertToInt(((Control)mskYear).Text), ConvertHelpers.ConvertToInt(((Control)mskMonth).Text), ConvertHelpers.ConvertToInt(((Control)mskDay).Text));
						if ((bool)axfpsPhuLuc.Sheets[0].Cells[0, 0].Value && (dateTime10 < dateTime8 || dateTime10 > dateTime9))
						{
							frmMessage frmMessage2 = new frmMessage("0551", "3", "", "Thông báo");
							((Form)frmMessage2).ShowDialog();
							return;
						}
					}
				}
			}
			if (menuID == "D4")
			{
				if (chkTKQuy.Checked)
				{
					if ((ConvertHelpers.ConvertToInt(((Control)mskYear).Text) < 2022 || ConvertHelpers.ConvertToInt(((Control)mskYear).Text) > 2022) && (bool)axfpsPhuLuc.Sheets[0].Cells[0, 0].Value)
					{
						frmMessage frmMessage2 = new frmMessage("0599", "3", "", "Thông báo");
						((Form)frmMessage2).ShowDialog();
						return;
					}
				}
				else if (chkTkhaiThang.Checked)
				{
					if (ConvertHelpers.ConvertToInt(((Control)mskYear).Text) < 2022 || ConvertHelpers.ConvertToInt(((Control)mskYear).Text) > 2022)
					{
						if ((bool)axfpsPhuLuc.Sheets[0].Cells[0, 0].Value)
						{
							frmMessage frmMessage2 = new frmMessage("0600", "3", "", "Thông báo");
							((Form)frmMessage2).ShowDialog();
							return;
						}
					}
					else if ((bool)axfpsPhuLuc.Sheets[0].Cells[0, 0].Value && ConvertHelpers.ConvertToInt(((Control)mskMonth).Text) < 2)
					{
						frmMessage frmMessage2 = new frmMessage("0600", "3", "", "Thông báo");
						((Form)frmMessage2).ShowDialog();
						return;
					}
				}
				else if (chkTKLanPS.Checked)
				{
					if (ConvertHelpers.ConvertToInt(((Control)mskYear).Text) < 2022 || ConvertHelpers.ConvertToInt(((Control)mskYear).Text) > 2022)
					{
						if ((bool)axfpsPhuLuc.Sheets[0].Cells[0, 0].Value)
						{
							frmMessage frmMessage2 = new frmMessage("0601", "3", "", "Thông báo");
							((Form)frmMessage2).ShowDialog();
							return;
						}
					}
					else
					{
						DateTime dateTime8 = new DateTime(2022, 2, 1);
						DateTime dateTime9 = new DateTime(2022, 12, 31);
						DateTime dateTime10 = new DateTime(ConvertHelpers.ConvertToInt(((Control)mskYear).Text), ConvertHelpers.ConvertToInt(((Control)mskMonth).Text), ConvertHelpers.ConvertToInt(((Control)mskDay).Text));
						if ((bool)axfpsPhuLuc.Sheets[0].Cells[0, 0].Value && (dateTime10 < dateTime8 || dateTime10 > dateTime9))
						{
							frmMessage frmMessage2 = new frmMessage("0601", "3", "", "Thông báo");
							((Form)frmMessage2).ShowDialog();
							return;
						}
					}
				}
			}
			if (menuID == "01" || menuID == "I5")
			{
				if (!"07".Equals(ConvertHelpers.ConvertToString(((ListControl)cboNganhKD).SelectedValue)) && !"00".Equals(ConvertHelpers.ConvertToString(((ListControl)cboNganhKD).SelectedValue)) && !"1701".Equals(ConvertHelpers.ConvertToString(((ListControl)cboNganhKD).SelectedValue)))
				{
					if ((bool)axfpsPhuLuc.Sheets[0].Cells[5, 0].Value)
					{
						frmMessage frmMessage2 = new frmMessage("0553", "3", "", "Thông báo");
						((Form)frmMessage2).ShowDialog();
						return;
					}
				}
				else if (chkTKQuy.Checked)
				{
					if (ConvertHelpers.ConvertToInt(((Control)mskYear).Text) < 2021 || ConvertHelpers.ConvertToInt(((Control)mskYear).Text) > 2021)
					{
						if ((bool)axfpsPhuLuc.Sheets[0].Cells[5, 0].Value)
						{
							frmMessage frmMessage2 = new frmMessage("0549", "3", "", "Thông báo");
							((Form)frmMessage2).ShowDialog();
							return;
						}
					}
					else if ((bool)axfpsPhuLuc.Sheets[0].Cells[5, 0].Value && ConvertHelpers.ConvertToInt(((Control)cbbQuy).Text) < 4)
					{
						frmMessage frmMessage2 = new frmMessage("0549", "3", "", "Thông báo");
						((Form)frmMessage2).ShowDialog();
						return;
					}
				}
				else if (chkTkhaiThang.Checked)
				{
					if (ConvertHelpers.ConvertToInt(((Control)mskYear).Text) < 2021 || ConvertHelpers.ConvertToInt(((Control)mskYear).Text) > 2021)
					{
						if ((bool)axfpsPhuLuc.Sheets[0].Cells[5, 0].Value)
						{
							frmMessage frmMessage2 = new frmMessage("0550", "3", "", "Thông báo");
							((Form)frmMessage2).ShowDialog();
							return;
						}
					}
					else if ((bool)axfpsPhuLuc.Sheets[0].Cells[5, 0].Value && ConvertHelpers.ConvertToInt(((Control)mskMonth).Text) < 11)
					{
						frmMessage frmMessage2 = new frmMessage("0550", "3", "", "Thông báo");
						((Form)frmMessage2).ShowDialog();
						return;
					}
				}
			}
			if (menuID == "D1")
			{
				if (!"00".Equals(ConvertHelpers.ConvertToString(((ListControl)cboNganhKD).SelectedValue)) && !"06".Equals(ConvertHelpers.ConvertToString(((ListControl)cboNganhKD).SelectedValue)) && !"05".Equals(ConvertHelpers.ConvertToString(((ListControl)cboNganhKD).SelectedValue)))
				{
					if ((bool)axfpsPhuLuc.Sheets[0].Cells[4, 0].Value)
					{
						frmMessage frmMessage2 = new frmMessage("0614", "3", "", "Thông báo");
						((Form)frmMessage2).ShowDialog();
						return;
					}
				}
				else if (chkTKQuy.Checked)
				{
					if ((ConvertHelpers.ConvertToInt(((Control)mskYear).Text) < 2022 || ConvertHelpers.ConvertToInt(((Control)mskYear).Text) > 2022) && (bool)axfpsPhuLuc.Sheets[0].Cells[4, 0].Value)
					{
						frmMessage frmMessage2 = new frmMessage("0599", "3", "", "Thông báo");
						((Form)frmMessage2).ShowDialog();
						return;
					}
				}
				else if (chkTkhaiThang.Checked)
				{
					if (ConvertHelpers.ConvertToInt(((Control)mskYear).Text) < 2022 || ConvertHelpers.ConvertToInt(((Control)mskYear).Text) > 2022)
					{
						if ((bool)axfpsPhuLuc.Sheets[0].Cells[4, 0].Value)
						{
							frmMessage frmMessage2 = new frmMessage("0600", "3", "", "Thông báo");
							((Form)frmMessage2).ShowDialog();
							return;
						}
					}
					else if ((bool)axfpsPhuLuc.Sheets[0].Cells[4, 0].Value && ConvertHelpers.ConvertToInt(((Control)mskMonth).Text) < 2)
					{
						frmMessage frmMessage2 = new frmMessage("0600", "3", "", "Thông báo");
						((Form)frmMessage2).ShowDialog();
						return;
					}
				}
			}
			if (menuID == "38")
			{
				if (ConvertHelpers.ConvertToInt(((Control)mskYear).Text) < 2021 || ConvertHelpers.ConvertToInt(((Control)mskYear).Text) > 2021)
				{
					if ((bool)axfpsPhuLuc.Sheets[0].Cells[0, 0].Value)
					{
						frmMessage frmMessage2 = new frmMessage("0544", "3", "", "Thông báo");
						((Form)frmMessage2).ShowDialog();
						return;
					}
				}
				else if ((bool)axfpsPhuLuc.Sheets[0].Cells[0, 0].Value && ConvertHelpers.ConvertToInt(((Control)cbbQuy).Text) < 3)
				{
					frmMessage frmMessage2 = new frmMessage("0544", "3", "", "Thông báo");
					((Form)frmMessage2).ShowDialog();
					return;
				}
			}
			if (menuID == "31")
			{
				if (chkTKQuy.Checked)
				{
					if (ConvertHelpers.ConvertToInt(((Control)mskYear).Text) < 2021 || ConvertHelpers.ConvertToInt(((Control)mskYear).Text) > 2021)
					{
						if ((bool)axfpsPhuLuc.Sheets[0].Cells[1, 0].Value)
						{
							frmMessage frmMessage2 = new frmMessage("0544", "3", "", "Thông báo");
							((Form)frmMessage2).ShowDialog();
							return;
						}
					}
					else if ((bool)axfpsPhuLuc.Sheets[0].Cells[1, 0].Value && ConvertHelpers.ConvertToInt(((Control)cbbQuy).Text) < 3)
					{
						frmMessage frmMessage2 = new frmMessage("0544", "3", "", "Thông báo");
						((Form)frmMessage2).ShowDialog();
						return;
					}
				}
				else if (chkTkhaiThang.Checked)
				{
					if (ConvertHelpers.ConvertToInt(((Control)mskYear).Text) < 2021 || ConvertHelpers.ConvertToInt(((Control)mskYear).Text) > 2021)
					{
						if ((bool)axfpsPhuLuc.Sheets[0].Cells[1, 0].Value)
						{
							frmMessage frmMessage2 = new frmMessage("0545", "3", "", "Thông báo");
							((Form)frmMessage2).ShowDialog();
							return;
						}
					}
					else if ((bool)axfpsPhuLuc.Sheets[0].Cells[1, 0].Value && ConvertHelpers.ConvertToInt(((Control)mskMonth).Text) < 7)
					{
						frmMessage frmMessage2 = new frmMessage("0545", "3", "", "Thông báo");
						((Form)frmMessage2).ShowDialog();
						return;
					}
				}
			}
			if (menuID == "T8" || menuID == "T9")
			{
				DateTime dateTime4 = new DateTime(2026, 2, 1);
				DateTime dateTime11 = new DateTime(2026, 3, 1);
				DateTime dateTime = new DateTime(2026, 2, 14);
				if (chkTkhaiThang.Checked)
				{
					DateTime dateTime2 = new DateTime(ConvertHelpers.ConvertToInt(((Control)mskYear).Text), ConvertHelpers.ConvertToInt(((Control)mskMonth).Text), 1);
					if ((dateTime2 >= dateTime4 && OptChinhthuc.Checked) || (dateTime2 >= dateTime11 && OptBosung.Checked))
					{
						frmMessage frmMessage2 = new frmMessage("0671", "3", "", "Thông báo");
						((Form)frmMessage2).ShowDialog();
						return;
					}
				}
				else if (chkTKLanPS.Checked)
				{
					DateTime dateTime2 = new DateTime(ConvertHelpers.ConvertToInt(((Control)mskYear).Text), ConvertHelpers.ConvertToInt(((Control)mskMonth).Text), ConvertHelpers.ConvertToInt(((Control)mskDay).Text));
					if (dateTime2 >= dateTime)
					{
						frmMessage frmMessage2 = new frmMessage("0671", "3", "", "Thông báo");
						((Form)frmMessage2).ShowDialog();
						return;
					}
				}
			}
			if (menuID == "30")
			{
				if (ckNam.Checked)
				{
					if ((ConvertHelpers.ConvertToInt(((Control)mskYear).Text) < 2021 || ConvertHelpers.ConvertToInt(((Control)mskYear).Text) > 2021) && (bool)axfpsPhuLuc.Sheets[0].Cells[1, 0].Value)
					{
						frmMessage frmMessage2 = new frmMessage("0548", "3", "", "Thông báo");
						((Form)frmMessage2).ShowDialog();
						return;
					}
				}
				else if (chkTKQuy.Checked)
				{
					DateTime dateTime12 = ConvertHelpers.ConvertToDateTime(((Control)mskTuNgay).Text, "dd/MM/yyyy");
					DateTime dateTime3 = ConvertHelpers.ConvertToDateTime(((Control)mskDenNgay).Text, "dd/MM/yyyy");
					DateTime dateTime8 = new DateTime(2021, 7, 1);
					DateTime dateTime9 = new DateTime(2021, 12, 31);
					if ((dateTime3 < dateTime8 || dateTime12 > dateTime9) && (bool)axfpsPhuLuc.Sheets[0].Cells[1, 0].Value)
					{
						frmMessage frmMessage2 = new frmMessage("0547", "3", "", "Thông báo");
						((Form)frmMessage2).ShowDialog();
						return;
					}
				}
			}
			if (menuID == "29" && (ConvertHelpers.ConvertToInt(((Control)mskYear).Text) < 2021 || ConvertHelpers.ConvertToInt(((Control)mskYear).Text) > 2021) && (bool)axfpsPhuLuc.Sheets[0].Cells[0, 0].Value)
			{
				frmMessage frmMessage2 = new frmMessage("0548", "3", "", "Thông báo");
				((Form)frmMessage2).ShowDialog();
				return;
			}
			if ("78~A1~A2~R4~C9~D6~".Contains(menuID))
			{
				string text2 = GeneralFunctions.CheckToKhaiHetHieuLuc(menuID, WinSettings.TTPeriodInfo);
				if (!string.IsNullOrEmpty(text2))
				{
					frmMessage frmMessage2 = new frmMessage("", "Tờ khai hiệu lực đến " + text2, "3", "3", "Thông báo");
					((Form)frmMessage2).ShowDialog();
					return;
				}
			}
			if ("U9~X1~X2~".Contains(menuID))
			{
				string text3 = GeneralFunctions.CheckToKhaiBatDauCoHieuLuc(menuID, WinSettings.TTPeriodInfo);
				if (!string.IsNullOrEmpty(text3))
				{
					frmMessage frmMessage2 = new frmMessage("", "Tờ khai hiệu lực từ " + text3, "3", "3", "Thông báo");
					((Form)frmMessage2).ShowDialog();
					return;
				}
			}
			if (!WinSettings.KieuKy.Equals("0") && !CheckPeriod(WinSettings.TTPeriodInfo))
			{
				return;
			}
			if (menuID == "G4" || menuID == "N5" || menuID == "H5" || menuID == "O7" || menuID == "P5" || menuID == "U6" || "U8~".Contains(WinSettings.MenuID))
			{
				if (isToKhaiLanDau(WinSettings.TTPeriodInfo) && OptChinhthuc.Checked)
				{
					frmMessage frmMessage2 = new frmMessage("0529", "3", "", "Thông báo");
					((Form)frmMessage2).ShowDialog();
					return;
				}
				if (OptBosung.Checked)
				{
					if ("U8~".Contains(WinSettings.MenuID))
					{
						DateTime dateTKLanDau = getDateTKLanDau();
						DateTime dateTime13 = new DateTime(ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Nam), ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Thang), ConvertHelpers.ConvertToInt(string.IsNullOrEmpty(WinSettings.TTPeriodInfo.Ngay) ? "01" : WinSettings.TTPeriodInfo.Ngay));
						if (dateTime13 < dateTKLanDau)
						{
							frmMessage frmMessage2 = new frmMessage("0530", "3", "", "Thông báo");
							((Form)frmMessage2).ShowDialog();
							return;
						}
					}
					else
					{
						DateTime dateTKLanDau = getDateTKLanDau();
						DateTime dateTime13 = new DateTime(ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Nam), ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Thang), ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Ngay));
						if (dateTime13 < dateTKLanDau)
						{
							frmMessage frmMessage2 = new frmMessage("0530", "3", "", "Thông báo");
							((Form)frmMessage2).ShowDialog();
							return;
						}
					}
				}
			}
			if (menuID == "M1")
			{
				DateTime t2 = new DateTime(ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Nam), ConvertHelpers.ConvertToInt(DateTime.Now.Month), ConvertHelpers.ConvertToInt(DateTime.Now.Day));
				if (DateTime.Compare(t2: new DateTime(ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Nam), ConvertHelpers.ConvertToInt("2"), ConvertHelpers.ConvertToInt("1")), t1: t2) >= 0)
				{
					frmMessage frmMessage2 = new frmMessage("0633", "3", "", "Thông báo");
					((Form)frmMessage2).ShowDialog();
					return;
				}
				if (ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Nam) < DateTime.Now.Year)
				{
					frmMessage frmMessage2 = new frmMessage("0638", "3", "", "Thông báo");
					((Form)frmMessage2).ShowDialog();
					return;
				}
			}
			if ("~O1~O2~O3~O4~O5~".IndexOf("~" + menuID + "~") >= 0 || "O6".Equals(menuID))
			{
				string s = WinSettings.TTPeriodInfo.Ngay + "/" + WinSettings.TTPeriodInfo.Thang + "/" + WinSettings.TTPeriodInfo.Nam;
				DateTime t2 = DateTime.ParseExact(s, "dd/MM/yyyy", CultureInfo.CurrentCulture);
				DateTime t3 = DateTime.ParseExact("31/10/2022", "dd/MM/yyyy", CultureInfo.CurrentCulture);
				if (DateTime.Compare(t2, t3) < 0)
				{
					frmMessage frmMessage2 = new frmMessage("0637", "3", "", "Thông báo");
					((Form)frmMessage2).ShowDialog();
					((Control)mskDay).Focus();
					return;
				}
			}
			if ("~O1~O2~O3~O4~O5~".IndexOf("~" + menuID + "~") >= 0)
			{
				if (HTKKFormater.isNullString(((Control)txtMst).Text))
				{
					frmMessage frmMessage2 = new frmMessage("0051", "3", "1", "Thông báo");
					((Form)frmMessage2).ShowDialog();
					((Control)txtMst).Focus();
					return;
				}
				if (!ValidateBase.CheckMaSoThue(((Control)txtMst).Text))
				{
					frmMessage frmMessage2 = new frmMessage("0047", "3", "1", "Thông báo");
					((Form)frmMessage2).ShowDialog();
					((Control)txtMst).Focus();
					return;
				}
			}
			if ((menuID == "P6" || menuID == "P1" || menuID == "X7" || menuID.Equals("Y9")) && HTKKFormater.isNullString(((Control)txtMst).Text))
			{
				frmMessage frmMessage2 = new frmMessage("0017", "3", "1", "Thông báo");
				((Form)frmMessage2).ShowDialog();
				((Control)txtMst).Focus();
				return;
			}
			WinDoWorkEvent.event_Leave(axfpsPhuLuc, listCheckBox, listMaskedTextBox, listComboBox, listRadioButton, "btnDongY_Click");
			strFileNameExtend = WinFnProcess.GetDataFileNameExtend(WinSettings.TTPeriodInfo);
			WinSettings.DicSheet = WinFnProcess.GetPhuLucDaKeKhai(((Control)pnVT2_PhuLucKeKhai).Visible, strFileNameExtend);
			fullNameDataFileToKhaiChinh = WinDoWorkEvent.GetFileNameBySheet(WinSettings.TTPeriodInfo, strDataFileToKhaiChinh, HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst + "\\", strVersionXML);
			bool flag = false;
			if (WinSettings.TTPeriodInfo.KieuKy.Equals("M"))
			{
				if (ExistTokhaiThangQuy(strDataFileToKhaiChinh, isThang: true, WinSettings.TTPeriodInfo))
				{
					frmMessage frmMessage2 = new frmMessage("0126", "3", "1", "Thông báo");
					((Form)frmMessage2).ShowDialog();
					return;
				}
			}
			else if (WinSettings.TTPeriodInfo.KieuKy.Equals("Q") && !menuID.Equals("10") && !menuID.Equals("67") && !menuID.Equals("42") && !menuID.Equals("A4") && !menuID.Equals("99") && !menuID.Equals("H6"))
			{
				flag = ExistTokhaiThangQuy(strDataFileToKhaiChinh, isThang: false, WinSettings.TTPeriodInfo);
				if (menuID == "L2" || menuID == "L7" || menuID == "Z8")
				{
					DateTime dateTime14 = ConvertHelpers.ConvertToDateTime("01/02/2026", "dd/MM/yyyy");
					DateTime dateTime15 = ConvertHelpers.ConvertToDateTime("01/0" + WinSettings.TTPeriodInfo.kyKKhai, "dd/MM/yyyy");
					if (dateTime15 >= dateTime14)
					{
						flag = false;
					}
				}
				if (flag)
				{
					frmMessage frmMessage2 = new frmMessage("0125", "3", "1", "Thông báo");
					((Form)frmMessage2).ShowDialog();
					return;
				}
			}
			if (menuID == "99" && ((ListControl)cboNganhKD).SelectedValue != null && !((ListControl)cboNganhKD).SelectedValue.Equals("0039") && WinSettings.TTPeriodInfo.maCQTQuanLy != _maCQTNoiNop)
			{
				frmMessage frmMessage2 = new frmMessage("0649", "3", "1", "Thông báo");
				((Form)frmMessage2).ShowDialog();
				return;
			}
			if (menuID == "A4" || menuID == "99" || menuID == "H6")
			{
				if (((ListControl)cboNganhKD).SelectedValue != null && ((ListControl)cboNganhKD).SelectedValue.Equals("0039"))
				{
					string[] array = ((Control)mskDenThang).Text.Split(new char[1] { '/' });
					if (array != null && array.Length > 1 && ConvertHelpers.ConvertToInt(array[1]) >= 2021 && ((Control)cbMaHSDNHT_DADT).Text.Trim().Equals(""))
					{
						frmMessage frmMessage2 = new frmMessage("0517", "3", "", "Thông báo");
						((Form)frmMessage2).ShowDialog();
						((Control)cbMaHSDNHT_DADT).Focus();
						return;
					}
				}
				string path = HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst;
				string[] files = Directory.GetFiles(path, (WinSettings.MenuID != "H6") ? "01_DNHT_*.xml" : "01A_DNHT_*.xml");
				bool flag2 = true;
				if (chkTuQuyDenQuy.Checked)
				{
					string text4 = ((Control)mskTuThang).Text.Split(new char[1] { '/' })[0];
					string text5 = ((Control)mskTuThang).Text.Split(new char[1] { '/' })[1];
					string text6 = ((Control)mskDenThang).Text.Split(new char[1] { '/' })[0];
					string text7 = ((Control)mskDenThang).Text.Split(new char[1] { '/' })[1];
					for (int i = 0; i < files.Length; i++)
					{
						string text8 = files[i];
						string text9 = "";
						text9 = ((!WinSettings.TTPeriodInfo.MaHSDuAnDT.Equals("")) ? ("01_DNHT_" + WinSettings.TTPeriodInfo.MaHSDuAnDT + "_Q") : ((WinSettings.MenuID != "H6") ? "01_DNHT_Q" : "01A_DNHT_Q"));
						if (text8.Contains(text9))
						{
							text8 = text8.Replace(text9, "_");
							string[] array2 = text8.Split(new char[1] { '_' });
							string text10 = array2[^3].Substring(0, 2);
							string text11 = array2[^3].Substring(2, 4);
							string text12 = array2[^2].Substring(0, 2);
							string text13 = array2[^2].Substring(2, 4);
							if ((ConvertHelpers.ConvertToInt(text5 + text4) != ConvertHelpers.ConvertToInt(text11 + text10) || ConvertHelpers.ConvertToInt(text7 + text6) != ConvertHelpers.ConvertToInt(text13 + text12)) && ConvertHelpers.ConvertToInt(text5 + text4) <= ConvertHelpers.ConvertToInt(text13 + text12) && ConvertHelpers.ConvertToInt(text7 + text6) >= ConvertHelpers.ConvertToInt(text11 + text10))
							{
								flag2 = false;
								frmMessage frmMessage2 = new frmMessage("0466", "3", "1", "Thông báo");
								((Form)frmMessage2).ShowDialog();
								return;
							}
						}
						text9 = "";
						text9 = ((!WinSettings.TTPeriodInfo.MaHSDuAnDT.Equals("")) ? ("01_DNHT_" + WinSettings.TTPeriodInfo.MaHSDuAnDT + "_") : ((WinSettings.MenuID != "H6") ? "01_DNHT_" : "01A_DNHT_"));
						if (text8.Contains(text9) && !text8.Contains("Q"))
						{
							text8 = text8.Replace(text9, "_");
							string[] array2 = text8.Split(new char[1] { '_' });
							string text14 = array2[^3].Substring(0, 2);
							string text11 = array2[^3].Substring(2, 4);
							string text15 = array2[^2].Substring(0, 2);
							string text13 = array2[^2].Substring(2, 4);
							if (ConvertHelpers.ConvertToInt(text5 + (ConvertHelpers.ConvertToInt(text4) * 3 - 2).ToString().PadLeft(2, '0')) <= ConvertHelpers.ConvertToInt(text13 + text15) && ConvertHelpers.ConvertToInt(text7 + (ConvertHelpers.ConvertToInt(text6) * 3).ToString().PadLeft(2, '0')) >= ConvertHelpers.ConvertToInt(text11 + text14))
							{
								flag2 = false;
								frmMessage frmMessage2 = new frmMessage("0468", "3", "1", "Thông báo");
								((Form)frmMessage2).ShowDialog();
								return;
							}
						}
					}
				}
				else
				{
					string text16 = ((Control)mskTuThang).Text.Split(new char[1] { '/' })[0];
					string text5 = ((Control)mskTuThang).Text.Split(new char[1] { '/' })[1];
					string text17 = ((Control)mskDenThang).Text.Split(new char[1] { '/' })[0];
					string text7 = ((Control)mskDenThang).Text.Split(new char[1] { '/' })[1];
					for (int i = 0; i < files.Length; i++)
					{
						string text8 = files[i];
						string text9 = "";
						text9 = ((!WinSettings.TTPeriodInfo.MaHSDuAnDT.Equals("")) ? ("01_DNHT_" + WinSettings.TTPeriodInfo.MaHSDuAnDT + "_Q") : ((WinSettings.MenuID != "H6") ? "01_DNHT_Q" : "01A_DNHT_Q"));
						if (text8.Contains(text9))
						{
							text8 = text8.Replace(text9, "_");
							string[] array2 = text8.Split(new char[1] { '_' });
							string text10 = array2[^3].Substring(0, 2);
							string text11 = array2[^3].Substring(2, 4);
							string text12 = array2[^2].Substring(0, 2);
							string text13 = array2[^2].Substring(2, 4);
							if (ConvertHelpers.ConvertToInt(text5 + text16) <= ConvertHelpers.ConvertToInt(text13 + (ConvertHelpers.ConvertToInt(text12) * 3 - 2)) && ConvertHelpers.ConvertToInt(text7 + text17) >= ConvertHelpers.ConvertToInt(text11 + ConvertHelpers.ConvertToInt(text12) * 3))
							{
								flag2 = false;
								frmMessage frmMessage2 = new frmMessage("0467", "3", "1", "Thông báo");
								((Form)frmMessage2).ShowDialog();
								return;
							}
						}
						text9 = ((!WinSettings.TTPeriodInfo.MaHSDuAnDT.Equals("")) ? ("01_DNHT_" + WinSettings.TTPeriodInfo.MaHSDuAnDT + "_") : ((WinSettings.MenuID != "H6") ? "01_DNHT_" : "01A_DNHT_"));
						if (text8.Contains(text9) && !text8.Contains("Q"))
						{
							text8 = text8.Replace(text9, "_");
							if (menuID.Equals("H6"))
							{
								text8 = text8.Replace("TT", "");
							}
							string[] array2 = text8.Split(new char[1] { '_' });
							string text14 = array2[^3].Substring(0, 2);
							string text11 = array2[^3].Substring(2, 4);
							string text15 = array2[^2].Substring(0, 2);
							string text13 = array2[^2].Substring(2, 4);
							if ((ConvertHelpers.ConvertToInt(text5 + text16) != ConvertHelpers.ConvertToInt(text11 + text14) || ConvertHelpers.ConvertToInt(text7 + text17) != ConvertHelpers.ConvertToInt(text13 + text15)) && ConvertHelpers.ConvertToInt(text5 + text16) <= ConvertHelpers.ConvertToInt(text13 + text15) && ConvertHelpers.ConvertToInt(text7 + text17) >= ConvertHelpers.ConvertToInt(text11 + text14))
							{
								flag2 = false;
								frmMessage frmMessage2 = new frmMessage("0469", "3", "1", "Thông báo");
								((Form)frmMessage2).ShowDialog();
								return;
							}
						}
					}
				}
				if (flag2)
				{
				}
			}
			if (menuID == "L9" || menuID == "I1" || menuID == "I2" || menuID == "G3" || menuID == "P1" || menuID == "P6" || menuID == "X7" || menuID.Equals("Y9"))
			{
				if (string.IsNullOrEmpty(((Control)txtLanXuatBan).Text))
				{
					frmMessage frmMessage2 = new frmMessage("0604", "3", "1", "Thông báo");
					((Form)frmMessage2).ShowDialog();
					return;
				}
				if ((!string.IsNullOrEmpty(((Control)txtLanXuatBan).Text) && !GeneralFunctions.IsNumeric(((Control)txtLanXuatBan).Text)) || (!string.IsNullOrEmpty(((Control)txtLanXuatBan).Text) && GeneralFunctions.IsNumeric(((Control)txtLanXuatBan).Text) && Convert.ToInt32(((Control)txtLanXuatBan).Text) <= 0))
				{
					frmMessage frmMessage2 = new frmMessage("0605", "3", "1", "Thông báo");
					((Form)frmMessage2).ShowDialog();
					return;
				}
			}
			if (((Control)cboNganhKD).Visible && ((Control)pnVT1_5DanhMucNghe).Visible && string.IsNullOrEmpty(((Control)cboNganhKD).Text))
			{
				frmMessage frmMessage2 = new frmMessage("0596", "3", "1", "Thông báo");
				((Form)frmMessage2).ShowDialog();
				((Control)cboNganhKD).Focus();
				return;
			}
			if (((Control)pnTruongHopQT).Visible && ((Control)cboTruongHopQT).Visible && string.IsNullOrEmpty(((Control)cboTruongHopQT).Text))
			{
				frmMessage frmMessage2 = new frmMessage("0596", "3", "1", "Thông báo");
				((Form)frmMessage2).ShowDialog();
				((Control)cboTruongHopQT).Focus();
				return;
			}
			if (menuID == "G6" && WinSettings.TTPeriodInfo.maCQTQuanLy == WinSettings.TTPeriodInfo.maCQTCapCuc)
			{
				frmMessage frmMessage2 = new frmMessage("0571", "3", "1", "Thông báo");
				((Form)frmMessage2).ShowDialog();
				return;
			}
			if (menuID == "G5")
			{
				if (Convert.ToBoolean(axfpsPhuLuc.Sheets[0].Cells[0, 0].Value) && Convert.ToBoolean(axfpsPhuLuc.Sheets[0].Cells[1, 0].Value))
				{
					frmMessage frmMessage2 = new frmMessage("0532", "3", "1", "Thông báo");
					((Form)frmMessage2).ShowDialog();
					return;
				}
				if (!ckNam.Checked && !chkTkhaiThang.Checked && !chkTKQuy.Checked && !chkTKLanPS.Checked)
				{
					frmMessage frmMessage2 = new frmMessage("0062", "3", "1", "Thông báo");
					((Form)frmMessage2).ShowDialog();
					return;
				}
			}
			if (WinSettings.ToKhaiInfo.NopTheoTinhCQT == "3" && (string.IsNullOrEmpty(WinSettings.TTPeriodInfo.MaTinh) || string.IsNullOrEmpty(WinSettings.TTPeriodInfo.TenTinh) || string.IsNullOrEmpty(WinSettings.TTPeriodInfo.MaXa) || string.IsNullOrEmpty(WinSettings.TTPeriodInfo.TenXa) || string.IsNullOrEmpty(WinSettings.TTPeriodInfo.maCQTCapCuc) || string.IsNullOrEmpty(WinSettings.TTPeriodInfo.tenCQTCapCuc) || string.IsNullOrEmpty(WinSettings.TTPeriodInfo.maCQTQuanLy) || string.IsNullOrEmpty(WinSettings.TTPeriodInfo.tenCQTQuanLy)))
			{
				frmMessage frmMessage2 = new frmMessage("0652", "3", "1", "Thông báo");
				((Form)frmMessage2).ShowDialog();
				return;
			}
			if (menuID == "G7")
			{
				if (WinSettings.TTPeriodInfo.maCQTQuanLy == WinSettings.TTPeriodInfo.maCQTCapCuc)
				{
					frmMessage frmMessage2 = new frmMessage("0571", "3", "1", "Thông báo");
					((Form)frmMessage2).ShowDialog();
					return;
				}
				if (Convert.ToBoolean(axfpsPhuLuc.Sheets[0].Cells[0, 0].Value) && Convert.ToBoolean(axfpsPhuLuc.Sheets[0].Cells[1, 0].Value))
				{
					frmMessage frmMessage2 = new frmMessage("0532", "3", "1", "Thông báo");
					((Form)frmMessage2).ShowDialog();
					return;
				}
				if (Convert.ToBoolean(axfpsPhuLuc.Sheets[0].Cells[0, 0].Value) && (WinSettings.TTPeriodInfo.KieuKy == "M" || WinSettings.TTPeriodInfo.KieuKy == "Q"))
				{
					frmMessage frmMessage2 = new frmMessage("0570", "3", "1", "Thông báo");
					((Form)frmMessage2).ShowDialog();
					return;
				}
				if (!ckNam.Checked && !chkTKKy.Checked && !chkTkhaiThang.Checked && !chkTKQuy.Checked)
				{
					frmMessage frmMessage2 = new frmMessage("0062", "3", "1", "Thông báo");
					((Form)frmMessage2).ShowDialog();
					return;
				}
			}
			if (menuID == "T8" || menuID == "R5")
			{
				if (WinSettings.TTPeriodInfo.KieuKy != "M")
				{
					axfpsPhuLuc.Sheets[0].Cells[0, 0].Value = false;
					axfpsPhuLuc.Sheets[0].Cells[1, 0].Value = false;
					axfpsPhuLuc.Sheets[0].Cells[2, 0].Value = false;
				}
				if (((ListControl)cboNganhKD).SelectedValue.ToString() != "03" && Convert.ToBoolean(axfpsPhuLuc.Sheets[0].Cells[1, 0].Value))
				{
					frmMessage frmMessage2 = new frmMessage("0572", "3", "3", "Thông báo");
					((Form)frmMessage2).ShowDialog();
					return;
				}
				if (((ListControl)cboNganhKD).SelectedValue.ToString() == "04" && WinSettings.TTPeriodInfo.maCQTQuanLy != _maCQTNoiNop)
				{
					frmMessage frmMessage2 = new frmMessage("0630", "3", "1", "Thông báo");
					((Form)frmMessage2).ShowDialog();
					return;
				}
				if (((ListControl)cboNganhKD).SelectedValue.ToString() != "04" && Convert.ToBoolean(axfpsPhuLuc.Sheets[0].Cells[2, 0].Value))
				{
					frmMessage frmMessage2 = new frmMessage("0631", "3", "3", "Thông báo");
					((Form)frmMessage2).ShowDialog();
					return;
				}
			}
			if (menuID == "L2" || menuID == "L7" || menuID == "Z8")
			{
				DateTime dateTime16 = ConvertHelpers.ConvertToDateTime("01/04/2026", "dd/MM/yyyy");
				DateTime dateTime15 = ConvertHelpers.ConvertToDateTime("01/" + WinSettings.TTPeriodInfo.kyKKhai, "dd/MM/yyyy");
				if (WinSettings.TTPeriodInfo.KieuKy == "M" && dateTime15 >= dateTime16 && WinSettings.TTPeriodInfo.LoaiToKhai.Equals("C"))
				{
					frmMessage frmMessage2 = new frmMessage("0682", "3", "1", "Thông báo");
					((Form)frmMessage2).ShowDialog();
					return;
				}
			}
			if (menuID == "T6" || menuID == "U2" || menuID == "Y5")
			{
				if (WinSettings.TTPeriodInfo.TieuMucHachToanVal == "2" && WinSettings.TTPeriodInfo.maCQTQuanLy != _maCQTNoiNop)
				{
					frmMessage frmMessage2 = new frmMessage("0584", "3", "1", "Thông báo");
					((Form)frmMessage2).ShowDialog();
					return;
				}
				WinSettings.TTPeriodInfo.TuNgay = ((Control)mskTuNgay).Text;
				if (string.IsNullOrEmpty(WinSettings.TTPeriodInfo.TuNgay) || WinSettings.TTPeriodInfo.TuNgay.Length != 10)
				{
					frmMessage frmMessage2 = new frmMessage("0681", "3", "", "Thông báo");
					((Form)frmMessage2).ShowDialog();
					return;
				}
				DateTime dateTime2 = new DateTime(ConvertHelpers.ConvertToInt(((Control)mskYear).Text), ConvertHelpers.ConvertToInt(((Control)mskMonth).Text), ConvertHelpers.ConvertToInt(((Control)mskDay).Text));
				DateTime dateTime17 = ConvertHelpers.ConvertToDateTime("17/03/2026", "dd/MM/yyyy");
				DateTime dateTime18 = ConvertHelpers.ConvertToDateTime("15/12/2025", "dd/MM/yyyy");
				DateTime dateTime19 = ConvertHelpers.ConvertToDateTime(WinSettings.TTPeriodInfo.TuNgay, "dd/MM/yyyy");
				if (menuID == "Y5" && dateTime2 < dateTime19)
				{
					frmMessage frmMessage2 = new frmMessage("0698", "3", "", "Thông báo");
					((Form)frmMessage2).ShowDialog();
					return;
				}
				if (dateTime2 >= dateTime17 && dateTime19 >= dateTime18)
				{
					if (menuID == "T6")
					{
						frmMessage frmMessage2 = new frmMessage("0680", "3", "", "Thông báo");
						((Form)frmMessage2).ShowDialog();
						return;
					}
				}
				else if (menuID == "U2")
				{
					frmMessage frmMessage2 = new frmMessage("0679", "3", "", "Thông báo");
					((Form)frmMessage2).ShowDialog();
					return;
				}
			}
			if (menuID == "D1" || WinSettings.MenuID.Equals("Z5"))
			{
				if (WinSettings.TTPeriodInfo.TieuMucHachToanVal == "00")
				{
					if (WinSettings.TTPeriodInfo.maCQTQuanLy != _maCQTNoiNop)
					{
						frmMessage frmMessage2 = new frmMessage("0611", "3", "1", "Thông báo");
						((Form)frmMessage2).ShowDialog();
						return;
					}
				}
				else if (WinSettings.TTPeriodInfo.TieuMucHachToanVal == "01")
				{
					if (WinSettings.TTPeriodInfo.maCQTQuanLy != _maCQTNoiNop)
					{
						frmMessage frmMessage2 = new frmMessage("0612", "3", "1", "Thông báo");
						((Form)frmMessage2).ShowDialog();
						return;
					}
				}
				else if (WinSettings.TTPeriodInfo.TieuMucHachToanVal == "05")
				{
					if (WinSettings.TTPeriodInfo.maCQTQuanLy != _maCQTNoiNop)
					{
						frmMessage frmMessage2 = new frmMessage("0608", "3", "1", "Thông báo");
						((Form)frmMessage2).ShowDialog();
						return;
					}
				}
				else if (WinSettings.TTPeriodInfo.TieuMucHachToanVal == "02")
				{
					if ((WinSettings.TTPeriodInfo.maCQTQuanLy.Length > 3 && _maCQTNoiNop.Length > 3 && WinSettings.TTPeriodInfo.maCQTQuanLy.Substring(0, 3) == _maCQTNoiNop.Substring(0, 3)) || (WinSettings.TTPeriodInfo.maCQTQuanLy.Length > 3 && _maCQTNoiNop.Length > 3 && WinSettings.TTPeriodInfo.MaTinh == _maCQTNoiNop.Substring(0, 3)))
					{
						frmMessage frmMessage2 = new frmMessage("0609", "3", "1", "Thông báo");
						((Form)frmMessage2).ShowDialog();
						return;
					}
				}
				else if (WinSettings.TTPeriodInfo.TieuMucHachToanVal == "06" && ((WinSettings.TTPeriodInfo.maCQTQuanLy.Length > 3 && _maCQTNoiNop.Length > 3 && WinSettings.TTPeriodInfo.maCQTQuanLy.Substring(0, 3) == _maCQTNoiNop.Substring(0, 3)) || (WinSettings.TTPeriodInfo.maCQTQuanLy.Length > 3 && _maCQTNoiNop.Length > 3 && WinSettings.TTPeriodInfo.MaTinh == _maCQTNoiNop.Substring(0, 3))))
				{
					frmMessage frmMessage2 = new frmMessage("0610", "3", "1", "Thông báo");
					((Form)frmMessage2).ShowDialog();
					return;
				}
			}
			if (WinSettings.TTPeriodInfo.LoaiToKhai.Equals("C") && WinSettings.TTPeriodInfo.KieuKy.Equals("Y"))
			{
				if (WinSettings.TTPeriodInfo.menuID.Equals("93") || WinSettings.TTPeriodInfo.menuID.Equals("89") || WinSettings.TTPeriodInfo.menuID.Equals("44") || WinSettings.TTPeriodInfo.menuID.Equals("42"))
				{
					string text18 = "";
					if (WinSettings.TTPeriodInfo.TieuMucHachToanVal.Equals("1"))
					{
						text18 = "DT";
					}
					else if (WinSettings.TTPeriodInfo.TieuMucHachToanVal.Equals("2"))
					{
						text18 = "KTN";
					}
					else if (WinSettings.TTPeriodInfo.TieuMucHachToanVal.Equals("3"))
					{
						text18 = "CD";
					}
					if (CheckKyKKTrung(strDataFileToKhaiChinh + "_" + text18, WinSettings.TTPeriodInfo))
					{
						frmMessage frmMessage2 = new frmMessage("0340", "3", "1", "Thông báo");
						((Form)frmMessage2).ShowDialog();
						return;
					}
				}
				else if (!WinSettings.TTPeriodInfo.menuID.Equals("41") && !WinSettings.TTPeriodInfo.menuID.Equals("37") && !WinSettings.TTPeriodInfo.menuID.Equals("76"))
				{
					if ("03,78,58,53,54,A1,A3,H1,H3,A8,H8,A2,A5,A6,A7,B3,B4,C1,C2,C3,C4,C5,X3".Contains(WinSettings.TTPeriodInfo.menuID) || "U9".Equals(WinSettings.TTPeriodInfo.menuID) || "X1".Equals(WinSettings.TTPeriodInfo.menuID) || "X2".Equals(WinSettings.TTPeriodInfo.menuID))
					{
						if (CheckKyKKTrungNgay(strDataFileToKhaiChinh, WinSettings.TTPeriodInfo))
						{
							frmMessage frmMessage2 = new frmMessage("0344", "3", "1", "Thông báo");
							((Form)frmMessage2).ShowDialog();
							return;
						}
					}
					else if ("N8,C6,C7,T5,H4".Contains(WinSettings.TTPeriodInfo.menuID))
					{
						if (CheckKyKKTrungNgay1(strDataFileToKhaiChinh, WinSettings.TTPeriodInfo))
						{
							frmMessage frmMessage2 = new frmMessage("0344", "3", "1", "Thông báo");
							((Form)frmMessage2).ShowDialog();
							return;
						}
					}
					else if (CheckKyKKTrung(strDataFileToKhaiChinh, WinSettings.TTPeriodInfo) && !WinSettings.TTPeriodInfo.menuID.Equals("39"))
					{
						frmMessage frmMessage2 = new frmMessage("0340", "3", "1", "Thông báo");
						((Form)frmMessage2).ShowDialog();
						return;
					}
				}
			}
			else if (WinSettings.TTPeriodInfo.LoaiToKhai.Equals("B") && WinSettings.TTPeriodInfo.KieuKy.Equals("Y"))
			{
				if (WinSettings.TTPeriodInfo.menuID.Equals("93") || WinSettings.TTPeriodInfo.menuID.Equals("89") || WinSettings.TTPeriodInfo.menuID.Equals("44") || WinSettings.TTPeriodInfo.menuID.Equals("42"))
				{
					string text18 = "";
					if (WinSettings.TTPeriodInfo.TieuMucHachToanVal.Equals("1"))
					{
						text18 = "DT";
					}
					else if (WinSettings.TTPeriodInfo.TieuMucHachToanVal.Equals("2"))
					{
						text18 = "KTN";
					}
					else if (WinSettings.TTPeriodInfo.TieuMucHachToanVal.Equals("3"))
					{
						text18 = "CD";
					}
					if (CheckKyKKTrung("bs" + WinSettings.TTPeriodInfo.SoLan + strDataFileToKhaiChinh + "_" + text18, WinSettings.TTPeriodInfo))
					{
						frmMessage frmMessage2 = new frmMessage("0340", "3", "1", "Thông báo");
						((Form)frmMessage2).ShowDialog();
						return;
					}
				}
				else if (!WinSettings.TTPeriodInfo.menuID.Equals("41") && !WinSettings.TTPeriodInfo.menuID.Equals("37") && !WinSettings.TTPeriodInfo.menuID.Equals("76"))
				{
					if (WinSettings.TTPeriodInfo.menuID.Equals("03") || WinSettings.TTPeriodInfo.menuID.Equals("N8"))
					{
						if (CheckKyKKTrungNgay("bs" + WinSettings.TTPeriodInfo.SoLan + strDataFileToKhaiChinh, WinSettings.TTPeriodInfo))
						{
							frmMessage frmMessage2 = new frmMessage("0344", "3", "1", "Thông báo");
							((Form)frmMessage2).ShowDialog();
							return;
						}
					}
					else if (WinSettings.TTPeriodInfo.menuID.Equals("C6"))
					{
						if (CheckKyKKTrungNgay("bs" + WinSettings.TTPeriodInfo.SoLan + strDataFileToKhaiChinh, WinSettings.TTPeriodInfo))
						{
							frmMessage frmMessage2 = new frmMessage("0344", "3", "1", "Thông báo");
							((Form)frmMessage2).ShowDialog();
							return;
						}
					}
					else if (CheckKyKKTrung("bs" + WinSettings.TTPeriodInfo.SoLan + strDataFileToKhaiChinh, WinSettings.TTPeriodInfo) && !WinSettings.TTPeriodInfo.menuID.Equals("39"))
					{
						frmMessage frmMessage2 = new frmMessage("0340", "3", "1", "Thông báo");
						((Form)frmMessage2).ShowDialog();
						return;
					}
				}
			}
			if (!checkCqtNoiNop() || (WinSettings.TTPeriodInfo.LoaiToKhai.Equals("B") && "I5~K1~K6~K8~K9~M7~N1~N9~O8~O9~".IndexOf(WinSettings.MenuID) < 0 && !WinSettings.TTPeriodInfo.menuID.Equals("G4") && !WinSettings.TTPeriodInfo.menuID.Equals("N5") && !WinSettings.TTPeriodInfo.menuID.Equals("H5") && !WinSettings.TTPeriodInfo.menuID.Equals("O7") && !WinSettings.TTPeriodInfo.menuID.Equals("P5") && !WinSettings.TTPeriodInfo.menuID.Equals("U6") && !WinSettings.TTPeriodInfo.menuID.Equals("P7") && !"U8~".Contains(WinSettings.MenuID) && !CheckPeriodKhbs(WinSettings.TTPeriodInfo, strDataFileToKhaiChinh)))
			{
				return;
			}
			if (menuID == "D5" || menuID == "Z4")
			{
				HTKKCMCTTinChungInfo headerInfo = GeneralFunctions.GetHeaderInfo(WinSettings.HTKKHeaderInfo.mst, removeMaCQT: false);
				if (headerInfo.maCQTCapCuc != "82500" && !string.IsNullOrEmpty(WinSettings.TTPeriodInfo.maCQTCapCuc) && headerInfo.maCQTCapCuc.Substring(0, 3) == WinSettings.TTPeriodInfo.maCQTCapCuc.Substring(0, 3))
				{
					frmMessage frmMessage2 = new frmMessage("0624", "3", "", "Thông báo");
					((Form)frmMessage2).ShowDialog();
					return;
				}
			}
			if (menuID == "D4" || WinSettings.MenuID.Equals("Z3"))
			{
				if (WinSettings.TTPeriodInfo.KieuKy == "D" && chkThuHo.Checked)
				{
					frmMessage frmMessage2 = new frmMessage("0593", "3", "", "Thông báo");
					((Form)frmMessage2).ShowDialog();
					return;
				}
				HTKKCMCTTinChungInfo headerInfo = GeneralFunctions.GetHeaderInfo(WinSettings.HTKKHeaderInfo.mst);
				if (chkThuHo.Checked && WinSettings.TTPeriodInfo.maCQTQuanLy != headerInfo.maCQTNoiNop)
				{
					frmMessage frmMessage2 = new frmMessage("0594", "3", "", "Thông báo");
					((Form)frmMessage2).ShowDialog();
					return;
				}
			}
			if (menuID == "G9")
			{
				HTKKCMCTTinChungInfo headerInfo = GeneralFunctions.GetHeaderInfo(WinSettings.HTKKHeaderInfo.mst);
				if ((Convert.ToBoolean(axfpsPhuLuc.Sheets[0].Cells[0, 0].Value) || Convert.ToBoolean(axfpsPhuLuc.Sheets[0].Cells[1, 0].Value)) && WinSettings.TTPeriodInfo.maCQTCapCuc != headerInfo.maCQTCapCuc)
				{
					frmMessage frmMessage2 = new frmMessage("0613", "3", "1", "Thông báo");
					((Form)frmMessage2).ShowDialog();
					return;
				}
			}
			if (menuID == "T8" || menuID == "R5")
			{
				if (((ListControl)cboNganhKD).SelectedValue.ToString() != "03" && Convert.ToBoolean(axfpsPhuLuc.Sheets[0].Cells[1, 0].Value))
				{
					frmMessage frmMessage2 = new frmMessage("0572", "3", "3", "Thông báo");
					((Form)frmMessage2).ShowDialog();
					return;
				}
				if (WinSettings.TTPeriodInfo.KieuKy == "D" && (((ListControl)cboNganhKD).SelectedValue.ToString() == "03" || ((ListControl)cboNganhKD).SelectedValue.ToString() == "02") && Convert.ToBoolean(axfpsPhuLuc.Sheets[0].Cells[1, 0].Value))
				{
					frmMessage frmMessage2 = new frmMessage("0586", "3", "3", "Thông báo");
					((Form)frmMessage2).ShowDialog();
					return;
				}
				HTKKCMCTTinChungInfo headerInfo = GeneralFunctions.GetHeaderInfo(WinSettings.HTKKHeaderInfo.mst);
				if (((ListControl)cboNganhKD).SelectedValue.ToString() == "03" && WinSettings.TTPeriodInfo.maCQTCapCuc != headerInfo.maCQTCapCuc)
				{
					frmMessage frmMessage2 = new frmMessage("0607", "3", "3", "Thông báo");
					((Form)frmMessage2).ShowDialog();
					return;
				}
				if (((ListControl)cboNganhKD).SelectedValue.ToString() == "02" && WinSettings.TTPeriodInfo.maCQTCapCuc != headerInfo.maCQTCapCuc)
				{
					frmMessage frmMessage2 = new frmMessage("0606", "3", "3", "Thông báo");
					((Form)frmMessage2).ShowDialog();
					return;
				}
				WinSettings.TTPeriodInfo.NganhNgheKinhDoanh = ((Control)cboNganhKD).Text.Trim();
			}
			if (menuID.Equals("N8"))
			{
				int row = 20;
				if (Convert.ToInt32(WinSettings.TTPeriodInfo.Nam) != 2021 && Convert.ToBoolean(axfpsPhuLuc.Sheets[0].Cells[row, 0].Value))
				{
					frmMessage frmMessage = new frmMessage("0669", "3", "3", "Thông báo");
					((Form)frmMessage).ShowDialog();
					return;
				}
			}
			if (menuID == "Q9" || menuID == "Y7")
			{
				if (chkTKLanPS.Checked && !Convert.ToBoolean(axfpsPhuLuc.Sheets[0].Cells[1, 0].Value))
				{
					frmMessage frmMessage2 = new frmMessage("0666", "3", "3", "Thông báo");
					((Form)frmMessage2).ShowDialog();
					return;
				}
				if (!Convert.ToBoolean(axfpsPhuLuc.Sheets[0].Cells[0, 0].Value) && !Convert.ToBoolean(axfpsPhuLuc.Sheets[0].Cells[1, 0].Value))
				{
					frmMessage frmMessage2 = new frmMessage("0670", "3", "3", "Thông báo");
					((Form)frmMessage2).ShowDialog();
					return;
				}
				if (Convert.ToBoolean(axfpsPhuLuc.Sheets[0].Cells[1, 0].Value) && (WinSettings.TTPeriodInfo.maCQTQuanLy.Equals("82700") || WinSettings.TTPeriodInfo.maCQTQuanLy.Equals("82500") || WinSettings.TTPeriodInfo.maCQTQuanLy == WinSettings.TTPeriodInfo.maCQTCapCuc))
				{
					frmMessage frmMessage2 = new frmMessage("0688", "3", "3", "Thông báo");
					((Form)frmMessage2).ShowDialog();
					return;
				}
			}
			if (menuID == "R9")
			{
				WinSettings.TTPeriodInfo.LanXuatBan = ((Control)txtLanXuatBan).Text;
			}
			if (menuID == "U1" && WinSettings.TTPeriodInfo.TieuMucHachToanVal == "2" && WinSettings.TTPeriodInfo.ToChucCNDNVonChuyenNhuong == 0)
			{
				frmMessage frmMessage2 = new frmMessage("0675", "3", "3", "Thông báo");
				((Form)frmMessage2).ShowDialog();
				return;
			}
			if (menuID == "Y3")
			{
				if (WinSettings.TTPeriodInfo.TieuMucHachToanVal == "1" && WinSettings.TTPeriodInfo.ToChucCNDNVonChuyenNhuong == 0)
				{
					frmMessage frmMessage2 = new frmMessage("0694", "3", "3", "Thông báo");
					((Form)frmMessage2).ShowDialog();
					return;
				}
				if (WinSettings.TTPeriodInfo.TieuMucHachToanVal == "2" && WinSettings.TTPeriodInfo.ToChucCNDNVonChuyenNhuong == 0)
				{
					frmMessage frmMessage2 = new frmMessage("0695", "3", "3", "Thông báo");
					((Form)frmMessage2).ShowDialog();
					return;
				}
			}
			if (menuID == "R6" && WinSettings.TTPeriodInfo.ToChucCNDNVonChuyenNhuong == 0)
			{
				frmMessage frmMessage2 = new frmMessage("0676", "3", "3", "Thông báo");
				((Form)frmMessage2).ShowDialog();
				return;
			}
			if (WinSettings.TTPeriodInfo.LoaiToKhai == "T" && menuID == "U8")
			{
				string sFolder = HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst + "\\";
				string menuChildenAttrById = WinFnProcess.getMenuChildenAttrById(WinSettings.MenuID, 0, "DataFile");
				string text19 = HTKKConfig.INTERFACE_TEMPLATES + "\\xml\\" + menuChildenAttrById + "_" + WinSettings.ToKhaiInfo.VersionXML.Replace(".", "") + ".xml";
				string fileNameBySheet = WinDoWorkEvent.GetFileNameBySheet(WinSettings.TTPeriodInfo, menuChildenAttrById, sFolder, WinSettings.ToKhaiInfo.VersionXML);
				Match match = Regex.Match(fileNameBySheet, "tt[^_]*_");
				if (match.Success)
				{
					string numberInString = GeneralFunctions.GetNumberInString(match.Value);
					if (!string.IsNullOrEmpty(numberInString))
					{
						fileNameBySheet = ((!(numberInString == "1")) ? fileNameBySheet.Replace("tt" + numberInString + "_", "tt" + (Convert.ToInt32(numberInString) - 1) + "_") : fileNameBySheet.Replace("tt" + numberInString + "_", ""));
						if (!File.Exists(fileNameBySheet))
						{
							frmMessage frmMessage2 = new frmMessage("0690", "3", "3", "Thông báo");
							((Form)frmMessage2).ShowDialog();
							return;
						}
					}
				}
			}
			if ((menuID == "R9" || menuID == "R6") && checkTenFileKhongPhuThuocDVT(menuID, "01/07/2026"))
			{
				frmMessage frmMessage2 = new frmMessage((WinSettings.TTPeriodInfo.KieuKy == "D") ? "0691" : "0692", "3", "3", "Thông báo");
				((Form)frmMessage2).ShowDialog();
				return;
			}
			if (menuID == "E9" && checkTenFileKhongPhuThuocDVT(menuID, "30/09/2026"))
			{
				frmMessage frmMessage2 = new frmMessage("0693", "3", "3", "Thông báo");
				((Form)frmMessage2).ShowDialog();
				return;
			}
			if (menuID == "Y7")
			{
				frmMessage frmMessage3 = new frmMessage("0699", "3", "", "Thông báo");
				((Form)frmMessage3).ShowDialog();
			}
			WinFnProcess.OpenToKhai((Form)(object)this, axfpsPhuLuc.Sheets[0], ((Control)pnVT2_PhuLucKeKhai).Visible, WinSettings.TTPeriodInfo);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private bool checkTenFileKhongPhuThuocDVT(string menuId, string gioiHan)
	{
		string sFolder = HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst + "\\";
		string menuChildenAttrById = WinFnProcess.getMenuChildenAttrById(WinSettings.MenuID, 0, "DataFile");
		string text = HTKKConfig.INTERFACE_TEMPLATES + "\\xml\\" + menuChildenAttrById + "_" + WinSettings.ToKhaiInfo.VersionXML.Replace(".", "") + ".xml";
		string text2 = WinDoWorkEvent.GetFileNameBySheet(WinSettings.TTPeriodInfo, menuChildenAttrById, sFolder, WinSettings.ToKhaiInfo.VersionXML);
		DateTime dateTime = DateTime.ParseExact(gioiHan, "dd/MM/yyyy", CultureInfo.CurrentCulture);
		DateTime dateTime2 = DateTime.ParseExact(WinSettings.TTPeriodInfo.DenNgay, "dd/MM/yyyy", CultureInfo.CurrentCulture);
		if (dateTime2 >= dateTime)
		{
			return false;
		}
		if (text2.Contains("VND"))
		{
			text2 = text2.Replace("VND", "USD");
		}
		else if (text2.Contains("USD"))
		{
			text2 = text2.Replace("USD", "VND");
		}
		if (File.Exists(text2))
		{
			return true;
		}
		return false;
	}

	private void SetDataFile()
	{
		//IL_0002: Unknown result type (might be due to invalid IL or missing references)
		//IL_0008: Expected O, but got Unknown
		try
		{
			XmlDocument val = new XmlDocument();
			val.LoadXml(xmlNode.OuterXml);
			elemList = val.GetElementsByTagName("Sheet");
			XmlNode val2 = elemList[0];
			strDataFileToKhaiChinh = ((val2.Attributes["DataFile"] == null) ? "" : ((XmlNode)val2.Attributes["DataFile"]).Value);
			strVersionXML = ((val2.ParentNode.ParentNode.Attributes["VersionXML"] == null) ? "" : ((XmlNode)val2.ParentNode.ParentNode.Attributes["VersionXML"]).Value);
			strSheetNameHeader = ((val2.Attributes["SheetName"] == null) ? "" : ((XmlNode)val2.Attributes["SheetName"]).Value);
			val2 = elemList[elemList.Count - 1];
			strDataFileToKHBS = ((val2.Attributes["DataFile"] == null) ? "" : ((XmlNode)val2.Attributes["DataFile"]).Value);
			strSheetNameKHBS = ((val2.Attributes["SheetName"] == null) ? "" : ((XmlNode)val2.Attributes["SheetName"]).Value);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void SetDefaultValues()
	{
		LoadDefaultInfor();
		SetDefaultView();
	}

	private void LoadDefaultInfor()
	{
		try
		{
			defaultDay = DateTime.Now.Day;
			defaultMonth = DateTime.Now.Month;
			defaultYear = DateTime.Now.Year;
			int num = defaultDay;
			int num2 = defaultMonth;
			int num3 = defaultYear;
			if (periodTypeValue.Equals("1"))
			{
				if (num2 == 1)
				{
					num2 = 12;
					num3--;
				}
				else
				{
					num2--;
				}
				((Control)mskMonth).Text = num2.ToString().PadLeft(2, '0');
				((Control)mskYear).Text = num3.ToString();
			}
			else if (periodTypeValue.Equals("2"))
			{
				if (DateTime.Now.Month <= 12 && DateTime.Now.Month > 9)
				{
					((ListControl)cbbQuy).SelectedIndex = 2;
					((Control)mskYear).Text = defaultYear.ToString();
				}
				if (DateTime.Now.Month <= 9 && DateTime.Now.Month > 6)
				{
					((ListControl)cbbQuy).SelectedIndex = 1;
					((Control)mskYear).Text = defaultYear.ToString();
				}
				if (DateTime.Now.Month <= 6 && DateTime.Now.Month > 3)
				{
					((ListControl)cbbQuy).SelectedIndex = 0;
					((Control)mskYear).Text = defaultYear.ToString();
				}
				if (DateTime.Now.Month <= 3 && DateTime.Now.Month > 0)
				{
					((ListControl)cbbQuy).SelectedIndex = 3;
					((Control)mskYear).Text = (defaultYear - 1).ToString();
				}
			}
			else if (periodTypeValue.Equals("4"))
			{
				if (DateTime.Now.Month <= 12 && DateTime.Now.Month > 9)
				{
					((ListControl)cbbQuy).SelectedIndex = 2;
					((Control)mskYear).Text = defaultYear.ToString();
				}
				if (DateTime.Now.Month <= 9 && DateTime.Now.Month > 6)
				{
					((ListControl)cbbQuy).SelectedIndex = 1;
					((Control)mskYear).Text = defaultYear.ToString();
				}
				if (DateTime.Now.Month <= 6 && DateTime.Now.Month > 3)
				{
					((ListControl)cbbQuy).SelectedIndex = 0;
					((Control)mskYear).Text = defaultYear.ToString();
				}
				if (DateTime.Now.Month <= 3 && DateTime.Now.Month > 0)
				{
					((ListControl)cbbQuy).SelectedIndex = 3;
					((Control)mskYear).Text = (defaultYear - 1).ToString();
				}
			}
			else if (periodTypeValue.Equals("3"))
			{
				if (menuID.Equals("63"))
				{
					((Control)mskYear).Text = defaultYear.ToString();
				}
				else
				{
					((Control)mskYear).Text = (defaultYear - 1).ToString();
				}
			}
			else if (periodTypeValue.Equals("6"))
			{
				int num4 = 1;
				int num5 = 1;
				num3 = DateUtils.GetNamHienTai(num4, num5);
				num3--;
				DateTime dateTime = new DateTime(num3, num4, num5);
				DateTime dateTime2 = dateTime.AddMonths(12).AddDays(-1.0);
				((Control)mskYear).Text = num3.ToString();
				((Control)mskTuNgay).Text = dateTime.ToString("dd/MM/yyyy");
				((Control)mskDenNgay).Text = dateTime2.ToString("dd/MM/yyyy");
			}
			else if (periodTypeValue.Equals("7"))
			{
				string text = WinSettings.HTKKHeaderInfo.NgayBatDauNamTaiChinh;
				if (string.IsNullOrEmpty(text))
				{
					text = "01/01";
				}
				string[] array = text.Split(new char[1] { '/' });
				int num4 = short.Parse(array[0]);
				int num5 = short.Parse(array[1]);
				num3 = DateUtils.GetNamHienTai(num4, num5);
				num3--;
				DateTime dateTime = new DateTime(num3, num5, num4);
				DateTime dateTime2 = dateTime.AddMonths(12).AddDays(-1.0);
				((Control)mskYear).Text = num3.ToString();
				((Control)mskTuNgay).Text = dateTime.ToString("dd/MM/yyyy", CultureInfo.CurrentCulture);
				((Control)mskDenNgay).Text = dateTime2.ToString("dd/MM/yyyy");
			}
			else if (periodTypeValue.Equals("5"))
			{
				((Control)mskYear).Text = defaultYear.ToString();
				((Control)mskMonth).Text = defaultMonth.ToString().PadLeft(2, '0');
				((Control)mskDay).Text = defaultDay.ToString().PadLeft(2, '0');
			}
			else if (periodTypeValue.Equals("13"))
			{
				((Control)mskYear).Text = defaultYear.ToString();
				((Control)mskMonth).Text = defaultMonth.ToString().PadLeft(2, '0');
				((Control)mskDay).Text = defaultDay.ToString().PadLeft(2, '0');
			}
			else if (!periodTypeValue.Equals("14"))
			{
				if (periodTypeValue.Equals("8"))
				{
					num3 = defaultYear;
					num3--;
					((Control)mskYear).Text = num3.ToString();
					((Control)mskTuThang).Text = "01/" + num3;
					((Control)mskDenThang).Text = "12/" + num3;
				}
				else if (periodTypeValue.Equals("9"))
				{
					((Control)mskYear).Text = (defaultYear - 1).ToString();
					((Control)txtLan).Text = "1";
				}
				else if (periodTypeValue.Equals("10"))
				{
					if (DateTime.Now.Month <= 12 && DateTime.Now.Month > 9)
					{
						((ListControl)cbbQuy).SelectedIndex = 2;
						((Control)mskYear).Text = defaultYear.ToString();
					}
					if (DateTime.Now.Month <= 9 && DateTime.Now.Month > 6)
					{
						((ListControl)cbbQuy).SelectedIndex = 1;
						((Control)mskYear).Text = defaultYear.ToString();
					}
					if (DateTime.Now.Month <= 6 && DateTime.Now.Month > 3)
					{
						((ListControl)cbbQuy).SelectedIndex = 0;
						((Control)mskYear).Text = defaultYear.ToString();
					}
					if (DateTime.Now.Month <= 3 && DateTime.Now.Month > 0)
					{
						((ListControl)cbbQuy).SelectedIndex = 3;
						((Control)mskYear).Text = (defaultYear - 1).ToString();
					}
					int iQuy = int.Parse(((Control)cbbQuy).Text);
					int iNam = int.Parse(((Control)mskYear).Text);
					DateTime ngayDauQuy = GeneralFunctions.GetNgayDauQuy(iQuy, iNam, 1, 1);
					DateTime ngayCuoiQuy = GeneralFunctions.GetNgayCuoiQuy(iQuy, iNam, 1, 1);
					((Control)mskTuThang).Text = ngayDauQuy.Month.ToString().PadLeft(2, '0') + ngayDauQuy.Year;
					((Control)mskDenThang).Text = ngayCuoiQuy.Month.ToString().PadLeft(2, '0') + ngayCuoiQuy.Year;
				}
				else if (periodTypeValue.Equals("11"))
				{
					string text2 = ((WinSettings.xmlMenuNode.Attributes["ParentID"] == null) ? "" : ((XmlNode)WinSettings.xmlMenuNode.Attributes["ParentID"]).Value);
					if (text2 == "101_6")
					{
						int iThangBdNamTc = ConvertHelpers.ConvertToInt(WinSettings.HTKKHeaderInfo.NgayBatDauNamTaiChinh.Split(new char[1] { '/' })[1]);
						int iNgayBdNamTc = ConvertHelpers.ConvertToInt(WinSettings.HTKKHeaderInfo.NgayBatDauNamTaiChinh.Split(new char[1] { '/' })[0]);
						int num6 = DateTime.Now.Month - 3;
						if (num6 <= 0)
						{
							num6 = 12;
							defaultYear--;
						}
						((ListControl)cbbQuy).SelectedIndex = 1;
						((Control)mskYear).Text = defaultYear.ToString();
						List<string> quyByNamTC = GeneralFunctions.GetQuyByNamTC(num6, defaultYear, WinSettings.HTKKHeaderInfo.NgayBatDauNamTaiChinh);
						if (quyByNamTC.Count > 0)
						{
							((Control)cbbQuy).Text = quyByNamTC[0];
							((Control)mskYear).Text = quyByNamTC[1];
						}
						int iQuy = int.Parse(((Control)cbbQuy).Text);
						int iNam = int.Parse(((Control)mskYear).Text);
						DateTime ngayDauQuy = GeneralFunctions.GetNgayDauQuy(iQuy, iNam, iNgayBdNamTc, iThangBdNamTc);
						DateTime ngayCuoiQuy = GeneralFunctions.GetNgayCuoiQuy(iQuy, iNam, iNgayBdNamTc, iThangBdNamTc);
						((Control)mskTuNgay).Text = ngayDauQuy.ToString("dd/MM/yyyy");
						((Control)mskDenNgay).Text = ngayCuoiQuy.ToString("dd/MM/yyyy");
					}
					else
					{
						if (DateTime.Now.Month <= 12 && DateTime.Now.Month > 9)
						{
							((ListControl)cbbQuy).SelectedIndex = 2;
							((Control)mskYear).Text = defaultYear.ToString();
						}
						if (DateTime.Now.Month <= 9 && DateTime.Now.Month > 6)
						{
							((ListControl)cbbQuy).SelectedIndex = 1;
							((Control)mskYear).Text = defaultYear.ToString();
						}
						if (DateTime.Now.Month <= 6 && DateTime.Now.Month > 3)
						{
							((ListControl)cbbQuy).SelectedIndex = 0;
							((Control)mskYear).Text = defaultYear.ToString();
						}
						if (DateTime.Now.Month <= 3 && DateTime.Now.Month > 0)
						{
							((ListControl)cbbQuy).SelectedIndex = 3;
							((Control)mskYear).Text = (defaultYear - 1).ToString();
						}
						int iQuy = int.Parse(((Control)cbbQuy).Text);
						int iNam = int.Parse(((Control)mskYear).Text);
						if (WinSettings.MenuID == "13" && strKieuKy == "7")
						{
							if (((ListControl)cbbKy).SelectedIndex == 1)
							{
								((Control)mskTuNgay).Text = "01/07" + ((Control)mskYear).Text;
								((Control)mskDenNgay).Text = DateTime.DaysInMonth(ConvertHelpers.ConvertToInt(((Control)mskYear).Text), 12) + "/12/" + ((Control)mskYear).Text;
							}
							if (((ListControl)cbbKy).SelectedIndex == 0)
							{
								((Control)mskTuNgay).Text = "01/01" + ((Control)mskYear).Text;
								((Control)mskDenNgay).Text = DateTime.DaysInMonth(ConvertHelpers.ConvertToInt(((Control)mskYear).Text), 6) + "/06/" + ((Control)mskYear).Text;
							}
						}
						else
						{
							DateTime ngayDauQuy = GeneralFunctions.GetNgayDauQuy(iQuy, iNam, 1, 1);
							DateTime ngayCuoiQuy = GeneralFunctions.GetNgayCuoiQuy(iQuy, iNam, 1, 1);
							((Control)mskTuNgay).Text = ngayDauQuy.ToString("dd/MM/yyyy");
							((Control)mskDenNgay).Text = ngayCuoiQuy.ToString("dd/MM/yyyy");
						}
					}
				}
				else if (periodTypeValue.Equals("12"))
				{
					((Control)mskYear).Text = defaultYear.ToString();
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
		SetPeriodCustom();
	}

	private void SetDefaultView()
	{
		((Control)mskDay).Text = ((Control)mskDay).Text.PadLeft(2, '0');
		((Control)mskMonth).Text = ((Control)mskMonth).Text.PadLeft(2, '0');
		if ((menuID == "65" || menuID == "13") && ((Control)mskTuNgay).Visible && ((Control)mskDenNgay).Visible)
		{
			if (((Control)chkTKKy).Visible && chkTKKy.Checked)
			{
				((Control)mskTuNgay).Enabled = false;
				((Control)mskDenNgay).Enabled = false;
			}
			else
			{
				((Control)mskTuNgay).Enabled = true;
				((Control)mskDenNgay).Enabled = true;
			}
		}
	}

	private void GetPeriodInfo(HTKKPeriodInfo periodInfo)
	{
		try
		{
			periodInfo.menuID = menuID;
			periodInfo.Nam = ((Control)mskYear).Text.Trim();
			if (((Control)mskTuNgay).Visible)
			{
				periodInfo.TuNgay = ((Control)mskTuNgay).Text.Trim();
			}
			else
			{
				periodInfo.TuNgay = "";
			}
			if (((Control)mskDenNgay).Visible)
			{
				periodInfo.DenNgay = ((Control)mskDenNgay).Text.Trim();
			}
			else
			{
				periodInfo.DenNgay = "";
			}
			if (((Control)mskTuThang).Visible)
			{
				periodInfo.TuThang = ((Control)mskTuThang).Text.Trim();
			}
			else
			{
				periodInfo.TuThang = "";
			}
			if (((Control)mskDenThang).Visible)
			{
				periodInfo.DenThang = ((Control)mskDenThang).Text.Trim();
			}
			else
			{
				periodInfo.DenThang = "";
			}
			if (((Control)mskMonth).Visible)
			{
				periodInfo.Thang = ((Control)mskMonth).Text.Trim();
			}
			else
			{
				periodInfo.Thang = "";
			}
			if (((Control)cbbQuy).Visible)
			{
				periodInfo.Quy = ((cbbQuy.SelectedItem != null) ? cbbQuy.SelectedItem.ToString() : "1");
			}
			else if (((Control)cbbKy).Visible)
			{
				periodInfo.Quy = cbbKy.SelectedItem.ToString();
			}
			else
			{
				periodInfo.Quy = "";
			}
			if (((Control)mskDay).Visible)
			{
				periodInfo.Ngay = ((Control)mskDay).Text.ToString();
			}
			else
			{
				periodInfo.Ngay = "";
			}
			if (((Control)cbMaHSDNHT_DADT).Visible && ((Control)cbMaHSDNHT_DADT).Text.Trim().Length > 0)
			{
				periodInfo.MaHSDuAnDT = ((Control)cbMaHSDNHT_DADT).Text.Trim();
			}
			if (((Control)pnMst).Visible)
			{
				if ("~O1~O2~O3~O4~O5~".IndexOf("~" + menuID + "~") >= 0)
				{
					periodInfo.MaHSDuAnDT = ((Control)txtMst).Text.Trim();
				}
				if (menuID == "P6" || menuID == "P1" || menuID == "X7" || menuID.Equals("Y9"))
				{
					periodInfo.MaHSDuAnDT = ((Control)txtMst).Text.Trim();
				}
			}
			if (((Control)pnLoaiTien).Visible && "~K2~K3~J7~E6~E7~D7~R6~X5~R9~E9~".IndexOf("~" + menuID + "~") >= 0 && cboLoaiTien.SelectedItem != null)
			{
				periodInfo.DonViTinhMa = ((ListControl)cboLoaiTien).SelectedValue.ToString();
				periodInfo.DonViTinhTen = ((Control)cboLoaiTien).Text;
			}
			if (((Control)pnThongBao).Visible)
			{
				periodInfo.SoTB = ((Control)txtSoTb).Text;
				periodInfo.NgayTb = ((Control)mskNgayTb).Text;
			}
			if (((Control)txtMaHSDNHT_DADT).Visible && ((Control)txtMaHSDNHT_DADT).Text.Trim().Length > 0)
			{
				periodInfo.MaHSDuAnDT = ((Control)txtMaHSDNHT_DADT).Text.Trim();
			}
			if (((Control)chkThuHo).Visible)
			{
				if (chkThuHo.Checked)
				{
					periodInfo.ThuHo = "TH";
				}
				else
				{
					periodInfo.ThuHo = "";
				}
			}
			if (((Control)cboTruongHopQT).Visible)
			{
				periodInfo.TenQuyetToan = ((Control)cboTruongHopQT).Text.Trim();
				if (((ListControl)cboTruongHopQT).SelectedValue != null)
				{
					periodInfo.MaQuyetToan = ((ListControl)cboTruongHopQT).SelectedValue.ToString().Trim();
				}
			}
			if (OptBosung.Checked)
			{
				periodInfo.NgayLapKHBS = ((Control)mskNgayLapKHBS).Text.ToString();
			}
			else
			{
				periodInfo.NgayLapKHBS = "";
			}
			if (!string.IsNullOrEmpty(cboNganhKD.SelectedText))
			{
				periodInfo.NganhNgheKinhDoanh = cboNganhKD.SelectedText;
			}
			if (((Control)pnQuyetToan_LienTuc).Visible)
			{
				if (chkQTTungNam_HoatDongLienTuc.Checked)
				{
					periodInfo.DoanhNghiepHoatDongLienTuc = "LT";
				}
				if (chkQTNamDau_KhongHoatDongLienTuc.Checked)
				{
					periodInfo.DoanhNghiepHoatDongLienTuc = "KL";
				}
			}
			else
			{
				periodInfo.DoanhNghiepHoatDongLienTuc = string.Empty;
			}
			if (OptChinhthuc.Checked)
			{
				if (menuID.Equals("95"))
				{
					periodInfo.SoLan = ((Control)txtLan).Text;
					periodInfo.LoaiToKhai = "C";
				}
				else
				{
					periodInfo.SoLan = "";
					periodInfo.LoaiToKhai = "C";
				}
			}
			else
			{
				periodInfo.SoLan = ((Control)txtLan).Text.Trim();
				periodInfo.LoaiToKhai = "B";
			}
			if (chkToChucKhaiThay.Checked)
			{
				periodInfo.ToChucKhaiThay = "1";
			}
			else
			{
				periodInfo.ToChucKhaiThay = "0";
			}
			if (((Control)pnChuyenNhuong).Visible)
			{
				periodInfo.ToChucCNDNVonChuyenNhuong = (chkToChucCaNhanChuyenNhuong.Checked ? 1 : (chkDoanhNghiepChuyenNhuong.Checked ? 2 : 0));
			}
			if (menuID.Equals("G4") || menuID.Equals("N5") || menuID.Equals("H5") || menuID.Equals("O7") || menuID.Equals("P5") || menuID.Equals("U6") || "U8~".Contains(WinSettings.MenuID))
			{
				if (OptChinhthuc.Checked)
				{
					periodInfo.LoaiToKhai = "C";
				}
				else
				{
					periodInfo.LoaiToKhai = "T";
				}
			}
			switch (strKieuKy)
			{
			case "10":
				periodInfo.kyKKhai = periodInfo.Nam;
				periodInfo.KieuKy = "Y";
				if (menuID == "95")
				{
					periodInfo.SoLan = ((Control)txtLan).Text;
				}
				if (menuID == "19" || menuID == "22" || menuID == "95")
				{
					periodInfo.kyKKhai = periodInfo.Nam;
					DateTime dateTime = new DateTime(int.Parse(periodInfo.Nam), 1, 1);
					DateTime dateTime2 = dateTime.AddYears(1).AddDays(-1.0);
					periodInfo.TuNgay = dateTime.Day.ToString().PadLeft(2, '0') + "/" + dateTime.Month.ToString().PadLeft(2, '0') + "/" + dateTime.Year;
					periodInfo.DenNgay = dateTime2.Day.ToString().PadLeft(2, '0') + "/" + dateTime2.Month.ToString().PadLeft(2, '0') + "/" + dateTime2.Year;
				}
				break;
			case "12":
				periodInfo.kyKKhai = periodInfo.Nam;
				periodInfo.KieuKy = "Y";
				periodInfo.TuThang = ((Control)mskTuThang).Text;
				periodInfo.DenThang = ((Control)mskDenThang).Text;
				break;
			case "24":
			{
				periodInfo.kyKKhai = periodInfo.Nam;
				DateTime dateTime = new DateTime(int.Parse(periodInfo.Nam), 1, 1);
				DateTime dateTime2 = dateTime.AddYears(1).AddDays(-1.0);
				periodInfo.TuNgay = dateTime.Day.ToString().PadLeft(2, '0') + "/" + dateTime.Month.ToString().PadLeft(2, '0') + "/" + dateTime.Year;
				periodInfo.DenNgay = dateTime2.Day.ToString().PadLeft(2, '0') + "/" + dateTime2.Month.ToString().PadLeft(2, '0') + "/" + dateTime2.Year;
				periodInfo.KieuKy = "Y";
				if (menuID == "L9" || menuID == "I1" || menuID == "I2")
				{
					periodInfo.LanXuatBan = ((Control)txtLanXuatBan).Text.Trim();
				}
				break;
			}
			case "5":
				if (chkTkhaiThang.Checked)
				{
					periodInfo.KieuKy = "M";
					int num2 = ConvertHelpers.ConvertToInt(((Control)mskYear).Text.Trim());
					int num3 = ConvertHelpers.ConvertToInt(((Control)mskMonth).Text.Trim());
					if (num2 != 0 && num3 != 0)
					{
						DateTime dateTime3 = new DateTime(num2, num3, 1);
						DateTime dateTime4 = dateTime3.AddMonths(1).AddDays(-1.0);
						periodInfo.TuNgay = dateTime3.Day.ToString().PadLeft(2, '0') + "/" + dateTime3.Month.ToString().PadLeft(2, '0') + "/" + dateTime3.Year;
						periodInfo.DenNgay = dateTime4.Day.ToString().PadLeft(2, '0') + "/" + dateTime4.Month.ToString().PadLeft(2, '0') + "/" + dateTime4.Year;
						periodInfo.Thang = ((Control)mskMonth).Text.Trim().PadLeft(2, '0');
						periodInfo.Nam = ((Control)mskYear).Text.Trim();
						periodInfo.kyKKhai = periodInfo.Thang + "/" + periodInfo.Nam;
					}
				}
				else if (chkTKQuy.Checked)
				{
					periodInfo.KieuKy = "Q";
					if (cbbQuy.SelectedItem != null && !HTKKFormater.isNullString(((Control)mskYear).Text))
					{
						periodInfo.Quy = cbbQuy.SelectedItem.ToString();
						periodInfo.Nam = ((Control)mskYear).Text.Trim();
						periodInfo.kyKKhai = periodInfo.Quy + "/" + periodInfo.Nam;
						int year = ConvertHelpers.ConvertToInt(periodInfo.Nam);
						int num = ConvertHelpers.ConvertToInt(periodInfo.Quy);
						int month2 = num * 3 - 2;
						int num4 = num * 3;
						periodInfo.TuNgay = new DateTime(year, month2, 1).ToString("dd/MM/yyyy");
						periodInfo.DenNgay = new DateTime(year, num4, DateTime.DaysInMonth(year, num4)).ToString("dd/MM/yyyy");
					}
				}
				break;
			case "19":
				if (chkTkhaiThang.Checked)
				{
					periodInfo.KieuKy = "M";
					int num2 = ConvertHelpers.ConvertToInt(((Control)mskYear).Text.Trim());
					int num3 = ConvertHelpers.ConvertToInt(((Control)mskMonth).Text.Trim());
					if (num2 == 0 || num3 == 0)
					{
						break;
					}
					if (num2 == 0)
					{
						num2 = DateTime.Now.Year;
					}
					DateTime dateTime3 = new DateTime(num2, num3, 1);
					DateTime dateTime4 = dateTime3.AddMonths(1).AddDays(-1.0);
					periodInfo.TuNgay = dateTime3.Day.ToString().PadLeft(2, '0') + "/" + dateTime3.Month.ToString().PadLeft(2, '0') + "/" + dateTime3.Year;
					periodInfo.DenNgay = dateTime4.Day.ToString().PadLeft(2, '0') + "/" + dateTime4.Month.ToString().PadLeft(2, '0') + "/" + dateTime4.Year;
					periodInfo.kyKKhai = ((Control)mskMonth).Text.Trim().ToString().PadLeft(2, '0') + "/" + ((Control)mskYear).Text.Trim();
				}
				else if (chkTKQuy.Checked)
				{
					periodInfo.KieuKy = "Q";
					if (cbbQuy.SelectedItem == null || HTKKFormater.isNullString(((Control)mskYear).Text.Trim()))
					{
						break;
					}
					periodInfo.kyKKhai = cbbQuy.SelectedItem.ToString() + "/" + ((Control)mskYear).Text.Trim();
					int year = ConvertHelpers.ConvertToInt(periodInfo.Nam);
					int num = ConvertHelpers.ConvertToInt(periodInfo.Quy);
					int month2 = num * 3 - 2;
					int num4 = num * 3;
					periodInfo.TuNgay = new DateTime(year, month2, 1).ToString("dd/MM/yyyy");
					periodInfo.DenNgay = new DateTime(year, month2, 1).AddMonths(3).AddDays(-1.0).ToString("dd/MM/yyyy");
				}
				if (((Control)cboNganhKD).Visible)
				{
					periodInfo.NganhNgheKinhDoanh = ((Control)cboNganhKD).Text;
					periodInfo.TieuMucHachToanVal = ((((ListControl)cboNganhKD).SelectedValue != null) ? ((ListControl)cboNganhKD).SelectedValue.ToString() : "");
				}
				break;
			case "27":
				if (chkTkhaiThang.Checked)
				{
					periodInfo.KieuKy = "M";
					int num2 = ((!((Control)mskYear).Text.Trim().Equals("0")) ? int.Parse(((Control)mskYear).Text.Trim()) : DateTime.Now.Year);
					int num3 = int.Parse(((Control)mskMonth).Text.Trim());
					DateTime dateTime3 = new DateTime(num2, num3, 1);
					DateTime dateTime4 = dateTime3.AddMonths(1).AddDays(-1.0);
					periodInfo.TuNgay = dateTime3.Day.ToString().PadLeft(2, '0') + "/" + dateTime3.Month.ToString().PadLeft(2, '0') + "/" + dateTime3.Year;
					periodInfo.DenNgay = dateTime4.Day.ToString().PadLeft(2, '0') + "/" + dateTime4.Month.ToString().PadLeft(2, '0') + "/" + dateTime4.Year;
					periodInfo.kyKKhai = ((Control)mskMonth).Text.Trim().ToString().PadLeft(2, '0') + "/" + ((Control)mskYear).Text.Trim();
				}
				else if (chkTKQuy.Checked)
				{
					periodInfo.KieuKy = "Q";
					periodInfo.kyKKhai = ((cbbQuy.SelectedItem != null) ? (cbbQuy.SelectedItem.ToString() + "/" + ((Control)mskYear).Text.Trim()) : "1");
					int year = ConvertHelpers.ConvertToInt(periodInfo.Nam);
					int num = ConvertHelpers.ConvertToInt(periodInfo.Quy);
					int month2 = num * 3 - 2;
					int num4 = num * 3;
					periodInfo.TuNgay = new DateTime(year, month2, 1).ToString("dd/MM/yyyy");
					periodInfo.DenNgay = new DateTime(year, month2, 1).AddMonths(3).AddDays(-1.0).ToString("dd/MM/yyyy");
				}
				break;
			case "11":
				periodInfo.kyKKhai = periodInfo.Nam;
				periodInfo.KieuKy = "Y";
				if (((Control)cboNganhKD).Visible)
				{
					periodInfo.NganhNgheKinhDoanh = ((Control)cboNganhKD).Text;
					periodInfo.TieuMucHachToanVal = ((((ListControl)cboNganhKD).SelectedValue != null) ? ((ListControl)cboNganhKD).SelectedValue.ToString() : "");
				}
				break;
			case "22":
				if (menuID == "80" || menuID == "82" || menuID == "C9" || menuID == "D6" || menuID.Equals("X9") || menuID == "Y1")
				{
					periodInfo.Nam = ((Control)mskDenNgay).Text.Trim().Split(new char[1] { '/' })[2];
				}
				periodInfo.kyKKhai = periodInfo.Nam;
				periodInfo.TuNgay = ((Control)mskTuNgay).Text;
				periodInfo.DenNgay = ((Control)mskDenNgay).Text;
				periodInfo.KieuKy = "D_D";
				break;
			case "26":
				periodInfo.KieuKy = "Y";
				if (ConvertHelpers.ConvertToInt(((Control)mskYear).Text.Trim()) != 0)
				{
					periodInfo.TuNgay = ((Control)mskTuNgay).Text;
					periodInfo.DenNgay = ((Control)mskDenNgay).Text;
					periodInfo.kyKKhai = periodInfo.Nam;
					periodInfo.DonViTinhTen = ((Control)cboNganhKD).Text;
					if (((ListControl)cboNganhKD).SelectedValue != null)
					{
						periodInfo.DonViTinhMa = ((ListControl)cboNganhKD).SelectedValue.ToString();
					}
				}
				break;
			case "25":
				periodInfo.kyKKhai = periodInfo.Nam;
				periodInfo.TuThang = ((Control)mskTuThang).Text;
				periodInfo.DenThang = ((Control)mskDenThang).Text;
				periodInfo.KieuKy = "Y";
				break;
			case "39":
				periodInfo.TuThang = ((Control)mskTuThang).Text;
				periodInfo.DenThang = ((Control)mskDenThang).Text;
				periodInfo.kyKKhai = ((Control)mskTuThang).Text;
				periodInfo.Nam = ((Control)mskTuThang).Text.Split(new char[1] { '/' })[1];
				if (chkTuQuyDenQuy.Checked)
				{
					periodInfo.KieuKy = "Q";
					periodInfo.Quy = ((Control)mskTuThang).Text.Split(new char[1] { '/' })[0];
				}
				else
				{
					periodInfo.KieuKy = "M";
					periodInfo.Thang = ((Control)mskTuThang).Text.Split(new char[1] { '/' })[0];
				}
				if (((Control)cboNganhKD).Visible)
				{
					periodInfo.NganhNgheKinhDoanh = ((Control)cboNganhKD).Text;
					periodInfo.TieuMucHachToanVal = ((((ListControl)cboNganhKD).SelectedValue != null) ? ((ListControl)cboNganhKD).SelectedValue.ToString() : "");
				}
				break;
			case "43":
				periodInfo.TuThang = ((Control)mskTuThang).Text;
				periodInfo.DenThang = ((Control)mskDenThang).Text;
				periodInfo.kyKKhai = ((Control)mskDenThang).Text;
				periodInfo.Nam = ((Control)mskTuThang).Text.Split(new char[1] { '/' })[1];
				switch (periodInfo.KieuKyHoanThue = ((ListControl)cboKieuKy).SelectedValue.ToString())
				{
				case "1":
					periodInfo.KieuKy = "M";
					periodInfo.Thang = ((Control)mskDenThang).Text.Split(new char[1] { '/' })[0];
					break;
				case "2":
					periodInfo.KieuKy = "Q";
					periodInfo.Quy = ((Control)mskDenThang).Text.Split(new char[1] { '/' })[0];
					break;
				case "3":
					periodInfo.KieuKy = "Q";
					periodInfo.Quy = ((Control)mskDenThang).Text.Split(new char[1] { '/' })[0];
					break;
				case "4":
					periodInfo.KieuKy = "M";
					periodInfo.Thang = ((Control)mskDenThang).Text.Split(new char[1] { '/' })[0];
					break;
				}
				if (((Control)cboNganhKD).Visible)
				{
					periodInfo.NganhNgheKinhDoanh = ((Control)cboNganhKD).Text;
					periodInfo.TieuMucHachToanVal = ((((ListControl)cboNganhKD).SelectedValue != null) ? ((ListControl)cboNganhKD).SelectedValue.ToString() : "");
				}
				break;
			case "44":
				periodInfo.TuNam = "";
				periodInfo.DenNam = "";
				periodInfo.kyKKhai = "";
				periodInfo.Nam = "";
				if (cboKieuKy != null && ((ListControl)cboKieuKy).SelectedValue != null)
				{
					string text2 = (periodInfo.KieuKyHoanThue = ((ListControl)cboKieuKy).SelectedValue.ToString());
					if (text2 == "5")
					{
						periodInfo.KieuKy = "Y";
						mskTuThang.Mask = "0000";
						mskDenThang.Mask = "0000";
						periodInfo.TuNam = ((Control)mskTuThang).Text;
						periodInfo.DenNam = ((Control)mskDenThang).Text;
						((Control)lblTuThang).Text = "Từ năm";
						((Control)lblDenThang).Text = "Đến năm";
						periodInfo.Nam = ((Control)mskDenThang).Text;
					}
					if (text2 == "6")
					{
						periodInfo.KieuKy = "D_D";
						periodInfo.Ngay = ((Control)mskNgayPS).Text;
					}
					switch (text2)
					{
					case "1":
						periodInfo.KieuKy = "M";
						periodInfo.Thang = ((Control)mskDenThang).Text.Split(new char[1] { '/' })[0];
						break;
					case "2":
						periodInfo.KieuKy = "Q";
						periodInfo.Quy = ((Control)mskDenThang).Text.Split(new char[1] { '/' })[0];
						break;
					case "3":
						periodInfo.KieuKy = "Q";
						periodInfo.Quy = ((Control)mskDenThang).Text.Split(new char[1] { '/' })[0];
						break;
					case "4":
						periodInfo.KieuKy = "M";
						periodInfo.Thang = ((Control)mskDenThang).Text.Split(new char[1] { '/' })[0];
						break;
					}
					if (((Control)cboNganhKD).Visible)
					{
						periodInfo.NganhNgheKinhDoanh = ((Control)cboNganhKD).Text;
						periodInfo.TieuMucHachToanVal = ((((ListControl)cboNganhKD).SelectedValue != null) ? ((ListControl)cboNganhKD).SelectedValue.ToString() : "");
					}
				}
				break;
			case "6":
				if (chkTkhaiThang.Checked)
				{
					periodInfo.KieuKy = "M";
					int num2 = ConvertHelpers.ConvertToInt(((Control)mskYear).Text.Trim());
					int num3 = ConvertHelpers.ConvertToInt(((Control)mskMonth).Text.Trim());
					if (num2 != 0 && num3 != 0)
					{
						DateTime dateTime3 = new DateTime(num2, num3, 1);
						DateTime dateTime4 = dateTime3.AddMonths(1).AddDays(-1.0);
						periodInfo.TuNgay = dateTime3.Day.ToString().PadLeft(2, '0') + "/" + dateTime3.Month.ToString().PadLeft(2, '0') + "/" + dateTime3.Year;
						periodInfo.DenNgay = dateTime4.Day.ToString().PadLeft(2, '0') + "/" + dateTime4.Month.ToString().PadLeft(2, '0') + "/" + dateTime4.Year;
						periodInfo.kyKKhai = ((Control)mskMonth).Text.Trim().ToString().PadLeft(2, '0') + "/" + ((Control)mskYear).Text.Trim();
					}
				}
				else if (chkTKQuy.Checked)
				{
					periodInfo.KieuKy = "Q";
					int year = ConvertHelpers.ConvertToInt(periodInfo.Nam);
					int num = ConvertHelpers.ConvertToInt(periodInfo.Quy);
					int month2 = num * 3 - 2;
					int num4 = num * 3;
					if (menuID != "96")
					{
						periodInfo.TuNgay = new DateTime(year, month2, 1).ToString("dd/MM/yyyy");
						periodInfo.DenNgay = new DateTime(year, month2, 1).AddMonths(3).AddDays(-1.0).ToString("dd/MM/yyyy");
					}
					if (menuID == "L2")
					{
						DateTime dateTime5 = ConvertHelpers.ConvertToDateTime("01/" + periodInfo.TuThang, "dd/MM/yyyy");
						DateTime dateTime6 = ConvertHelpers.ConvertToDateTime("01/" + periodInfo.DenThang, "dd/MM/yyyy").AddMonths(1).AddDays(-1.0);
						periodInfo.TuNgay = dateTime5.ToString("dd/MM/yyyy");
						periodInfo.DenNgay = dateTime6.ToString("dd/MM/yyyy");
					}
					periodInfo.kyKKhai = cbbQuy.SelectedItem.ToString() + "/" + ((Control)mskYear).Text.Trim();
				}
				break;
			case "8":
				periodInfo.KieuKy = "D";
				periodInfo.Ngay = ((Control)mskDay).Text.Trim().PadLeft(2, '0');
				periodInfo.Thang = ((Control)mskMonth).Text.Trim().PadLeft(2, '0');
				periodInfo.Nam = ((Control)mskYear).Text.Trim();
				periodInfo.kyKKhai = periodInfo.Ngay + "/" + periodInfo.Thang + "/" + periodInfo.Nam;
				periodInfo.TuNgay = periodInfo.kyKKhai;
				periodInfo.DenNgay = periodInfo.kyKKhai;
				periodInfo.ToKhaiPhatSinh = 1;
				if (menuID == "47" || menuID == "K7")
				{
					periodInfo.CTDC = cbxCTDC.Checked;
					periodInfo.LNDC = cbxLNDC.Checked;
					periodInfo.NgayChiaCTLN = ((Control)mskDay).Text + "/" + ((Control)mskMonth).Text + "/" + ((Control)mskYear).Text;
				}
				if (((Control)cboNganhKD).Visible)
				{
					periodInfo.NganhNgheKinhDoanh = ((Control)cboNganhKD).Text;
					periodInfo.TieuMucHachToanVal = ((((ListControl)cboNganhKD).SelectedValue != null) ? ((ListControl)cboNganhKD).SelectedValue.ToString() : "");
				}
				if (menuID == "P1" || menuID == "P6" || menuID == "X7" || menuID.Equals("Y9"))
				{
					periodInfo.LanXuatBan = ((Control)txtLanXuatBan).Text.Trim();
					periodInfo.SoLan = ((Control)txtLanXuatBan).Text.Trim();
				}
				break;
			case "13":
				if (chkTKLanPS.Checked)
				{
					periodInfo.KieuKy = "D";
					periodInfo.Ngay = ((Control)mskDay).Text.Trim().PadLeft(2, '0');
					periodInfo.Thang = ((Control)mskMonth).Text.Trim().PadLeft(2, '0');
					periodInfo.Nam = ((Control)mskYear).Text.Trim();
					periodInfo.kyKKhai = periodInfo.Ngay + "/" + periodInfo.Thang + "/" + periodInfo.Nam;
					periodInfo.TuNgay = periodInfo.Ngay + "/" + periodInfo.Thang + "/" + periodInfo.Nam;
					periodInfo.DenNgay = periodInfo.Ngay + "/" + periodInfo.Thang + "/" + periodInfo.Nam;
					periodInfo.ToKhaiPhatSinh = 1;
				}
				else if (menuID != "T5" && menuID != "Y4")
				{
					periodInfo.KieuKy = "Y";
					string text4 = (periodInfo.kyKKhai = periodInfo.Nam);
					periodInfo.Nam = text4;
					periodInfo.TuNgay = "01/01/" + periodInfo.Nam;
					periodInfo.DenNgay = "31/12/" + periodInfo.Nam;
				}
				else
				{
					periodInfo.KieuKy = "Y";
					periodInfo.kyKKhai = ((Control)mskYear).Text.Trim();
				}
				break;
			case "18":
				if (chkTKLanPS.Checked)
				{
					periodInfo.KieuKy = "D";
					periodInfo.TuNgay = ((Control)mskDay).Text.Trim().PadLeft(2, '0') + "/" + ((Control)mskMonth).Text.Trim().PadLeft(2, '0') + "/" + ((Control)mskYear).Text.Trim();
					periodInfo.DenNgay = ((Control)mskDay).Text.Trim().PadLeft(2, '0') + "/" + ((Control)mskMonth).Text.Trim().PadLeft(2, '0') + "/" + ((Control)mskYear).Text.Trim();
					periodInfo.Ngay = ((Control)mskDay).Text.Trim().PadLeft(2, '0');
					periodInfo.Thang = ((Control)mskMonth).Text.Trim().PadLeft(2, '0');
					periodInfo.Nam = ((Control)mskYear).Text.Trim();
					periodInfo.kyKKhai = periodInfo.Ngay + "/" + periodInfo.Thang + "/" + periodInfo.Nam;
				}
				else if (chkTkhaiThang.Checked)
				{
					periodInfo.KieuKy = "M";
					int num2 = ConvertHelpers.ConvertToInt(((Control)mskYear).Text.Trim());
					int num3 = ConvertHelpers.ConvertToInt(((Control)mskMonth).Text.Trim());
					if (num2 == 0 || num3 == 0)
					{
						break;
					}
					DateTime dateTime3 = new DateTime(num2, num3, 1);
					DateTime dateTime4 = dateTime3.AddMonths(1).AddDays(-1.0);
					periodInfo.TuNgay = dateTime3.Day.ToString().PadLeft(2, '0') + "/" + dateTime3.Month.ToString().PadLeft(2, '0') + "/" + dateTime3.Year;
					periodInfo.DenNgay = dateTime4.Day.ToString().PadLeft(2, '0') + "/" + dateTime4.Month.ToString().PadLeft(2, '0') + "/" + dateTime4.Year;
					periodInfo.Thang = ((Control)mskMonth).Text.Trim().PadLeft(2, '0');
					periodInfo.Nam = ((Control)mskYear).Text.Trim();
					periodInfo.kyKKhai = periodInfo.Thang + "/" + periodInfo.Nam;
				}
				if (menuID == "G3")
				{
					periodInfo.LanXuatBan = ((Control)txtLanXuatBan).Text.Trim();
				}
				break;
			case "21":
				if (chkTKLanPS.Checked)
				{
					periodInfo.KieuKy = "D";
					periodInfo.TuNgay = ((Control)mskDay).Text.Trim().PadLeft(2, '0') + "/" + ((Control)mskMonth).Text.Trim().PadLeft(2, '0') + "/" + ((Control)mskYear).Text.Trim();
					periodInfo.DenNgay = ((Control)mskDay).Text.Trim().PadLeft(2, '0') + "/" + ((Control)mskMonth).Text.Trim().PadLeft(2, '0') + "/" + ((Control)mskYear).Text.Trim();
					periodInfo.Ngay = ((Control)mskDay).Text.Trim().PadLeft(2, '0');
					periodInfo.Thang = ((Control)mskMonth).Text.Trim().PadLeft(2, '0');
					periodInfo.Nam = ((Control)mskYear).Text.Trim();
					periodInfo.kyKKhai = periodInfo.Ngay + "/" + periodInfo.Thang + "/" + periodInfo.Nam;
				}
				else if (chkTkhaiThang.Checked)
				{
					periodInfo.KieuKy = "M";
					int num2 = ConvertHelpers.ConvertToInt(((Control)mskYear).Text.Trim());
					int num3 = ConvertHelpers.ConvertToInt(((Control)mskMonth).Text.Trim());
					if (num2 != 0 && num3 != 0)
					{
						DateTime dateTime3 = new DateTime(num2, num3, 1);
						DateTime dateTime4 = dateTime3.AddMonths(1).AddDays(-1.0);
						periodInfo.TuNgay = dateTime3.Day.ToString().PadLeft(2, '0') + "/" + dateTime3.Month.ToString().PadLeft(2, '0') + "/" + dateTime3.Year;
						periodInfo.DenNgay = dateTime4.Day.ToString().PadLeft(2, '0') + "/" + dateTime4.Month.ToString().PadLeft(2, '0') + "/" + dateTime4.Year;
						periodInfo.Thang = ((Control)mskMonth).Text.Trim().PadLeft(2, '0');
						periodInfo.Nam = ((Control)mskYear).Text.Trim();
						periodInfo.kyKKhai = periodInfo.Thang + "/" + periodInfo.Nam;
					}
				}
				else if (chkTKQuy.Checked)
				{
					periodInfo.KieuKy = "Q";
					int year = ConvertHelpers.ConvertToInt(periodInfo.Nam);
					int num = ConvertHelpers.ConvertToInt(periodInfo.Quy);
					int month2 = num * 3 - 2;
					int num4 = month2 + 2;
					periodInfo.TuNgay = new DateTime(year, month2, 1).ToString("dd/MM/yyyy");
					periodInfo.DenNgay = new DateTime(year, num4, DateTime.DaysInMonth(year, num4)).ToString("dd/MM/yyyy");
					periodInfo.Quy = cbbQuy.SelectedItem.ToString();
					periodInfo.Nam = ((Control)mskYear).Text.Trim();
					periodInfo.kyKKhai = periodInfo.Quy + "/" + periodInfo.Nam;
				}
				break;
			case "14":
			{
				periodInfo.KieuKy = "Q";
				periodInfo.Quy = cbbQuy.SelectedItem.ToString();
				periodInfo.Nam = ((Control)mskYear).Text.Trim();
				int num5 = ConvertHelpers.ConvertToInt(periodInfo.Quy) * 3 - 2;
				int month3 = num5 + 2;
				if (!string.IsNullOrEmpty(periodInfo.Nam))
				{
					periodInfo.TuNgay = new DateTime(ConvertHelpers.ConvertToInt(periodInfo.Nam), num5, 1).ToString("dd/MM/yyyy");
					periodInfo.DenNgay = new DateTime(ConvertHelpers.ConvertToInt(periodInfo.Nam), month3, DateTime.DaysInMonth(ConvertHelpers.ConvertToInt(periodInfo.Nam), month3)).ToString("dd/MM/yyyy");
				}
				periodInfo.kyKKhai = periodInfo.Quy + "/" + periodInfo.Nam;
				if (((Control)cboNganhKD).Visible)
				{
					periodInfo.NganhNgheKinhDoanh = ((Control)cboNganhKD).Text;
					periodInfo.TieuMucHachToanVal = ((ListControl)cboNganhKD).SelectedValue.ToString();
				}
				break;
			}
			case "17":
			{
				periodInfo.KieuKy = "M";
				periodInfo.Thang = ((Control)mskMonth).Text.Trim().PadLeft(2, '0');
				periodInfo.Nam = ((Control)mskYear).Text.Trim();
				periodInfo.kyKKhai = periodInfo.Thang + "/" + periodInfo.Nam;
				int month = ConvertHelpers.ConvertToInt(periodInfo.Thang);
				int year = ConvertHelpers.ConvertToInt(periodInfo.Nam);
				periodInfo.TuNgay = new DateTime(year, month, 1).ToString("dd/MM/yyyy");
				periodInfo.DenNgay = new DateTime(year, month, DateTime.DaysInMonth(year, month)).ToString("dd/MM/yyyy");
				break;
			}
			case "9":
				if (chkTKhaiLanXB.Checked)
				{
					periodInfo.KieuKy = "D";
					if (!(menuID == "J8") && !(menuID == "J7") && !(menuID == "R6"))
					{
						periodInfo.LanXuatBan = ((Control)txtLanXuatBan).Text.Trim();
					}
					periodInfo.Ngay = ((Control)mskDay).Text.Trim().PadLeft(2, '0');
					periodInfo.Thang = ((Control)mskMonth).Text.Trim().PadLeft(2, '0');
					periodInfo.Nam = ((Control)mskYear).Text.Trim();
					periodInfo.kyKKhai = periodInfo.Ngay + "/" + periodInfo.Thang + "/" + periodInfo.Nam;
					periodInfo.TuNgay = periodInfo.kyKKhai;
					periodInfo.DenNgay = periodInfo.kyKKhai;
				}
				else if (chkTkhaiThang.Checked)
				{
					periodInfo.KieuKy = "M";
					periodInfo.Thang = ((Control)mskMonth).Text.Trim();
					periodInfo.Nam = ((Control)mskYear).Text.Trim();
					periodInfo.kyKKhai = periodInfo.Thang + "/" + periodInfo.Nam;
					int month = ConvertHelpers.ConvertToInt(periodInfo.Thang);
					int year = ConvertHelpers.ConvertToInt(periodInfo.Nam);
					periodInfo.TuNgay = new DateTime(year, month, 1).ToString("dd/MM/yyyy");
					periodInfo.DenNgay = new DateTime(year, month, DateTime.DaysInMonth(year, month)).ToString("dd/MM/yyyy");
				}
				if (cboNganhKD.SelectedItem != null)
				{
					periodInfo.NganhNgheKinhDoanh = ((Control)cboNganhKD).Text;
					periodInfo.TieuMucHachToanVal = ((ListControl)cboNganhKD).SelectedValue.ToString();
				}
				break;
			case "28":
				if (ckNam.Checked)
				{
					periodInfo.kyKKhai = ((Control)mskYear).Text.Trim();
					periodInfo.TuThang = ((Control)mskTuThang).Text;
					periodInfo.DenThang = ((Control)mskDenThang).Text;
					periodInfo.KieuKy = "Y";
					periodInfo.NganhNgheKinhDoanh = ((Control)cboNganhKD).Text;
					periodInfo.TieuMucHachToanVal = ((ListControl)cboNganhKD).SelectedValue.ToString();
				}
				else
				{
					periodInfo.kyKKhai = ((cbbQuy.SelectedItem != null) ? (cbbQuy.SelectedItem.ToString() + "/" + ((Control)mskYear).Text.Trim()) : string.Empty);
					periodInfo.TuThang = ((Control)mskTuThang).Text;
					periodInfo.DenThang = ((Control)mskDenThang).Text;
					periodInfo.KieuKy = "Q";
					periodInfo.NganhNgheKinhDoanh = ((Control)cboNganhKD).Text;
					periodInfo.TieuMucHachToanVal = ((ListControl)cboNganhKD).SelectedValue.ToString();
				}
				break;
			case "30":
			{
				periodInfo.KieuKy = "Q";
				periodInfo.Quy = ConvertHelpers.ConvertToInt(cbbQuy.SelectedItem.ToString()).ToString();
				periodInfo.Nam = ((Control)mskYear).Text.Trim();
				periodInfo.kyKKhai = periodInfo.Quy + "/" + periodInfo.Nam;
				int year2 = ConvertHelpers.ConvertToInt(periodInfo.Nam);
				int num6 = ConvertHelpers.ConvertToInt(periodInfo.Quy);
				if (((Control)mskTuThang).Visible && ((Control)mskDenThang).Visible)
				{
					periodInfo.TuNgay = "01/" + ((Control)mskTuThang).Text;
					periodInfo.DenNgay = DateTime.DaysInMonth(ConvertHelpers.ConvertToInt(periodInfo.Nam), ConvertHelpers.ConvertToInt(((Control)mskDenThang).Text.Split(new char[1] { '/' })[0])) + "/" + ((Control)mskDenThang).Text;
					periodInfo.TuThang = ((Control)mskTuThang).Text;
					periodInfo.DenThang = ((Control)mskDenThang).Text;
				}
				else
				{
					periodInfo.TuNgay = new DateTime(year2, num6 * 3 - 2, 1).ToString("dd/MM/yyyy");
					periodInfo.DenNgay = new DateTime(year2, num6 * 3, DateTime.DaysInMonth(year2, num6 * 3)).ToString("dd/MM/yyyy");
					periodInfo.TuThang = new DateTime(year2, num6 * 3 - 2, 1).ToString("MM/yyyy");
					periodInfo.DenThang = new DateTime(year2, num6 * 3, DateTime.DaysInMonth(year2, num6 * 3)).ToString("MM/yyyy");
				}
				break;
			}
			case "31":
				if (chkTkhaiThang.Checked)
				{
					periodInfo.KieuKy = "M";
					periodInfo.Thang = ((Control)mskMonth).Text.Trim().ToString().PadLeft(2, '0');
					periodInfo.Nam = ((Control)mskYear).Text.Trim().ToString();
					periodInfo.kyKKhai = periodInfo.Thang + "/" + periodInfo.Nam;
					periodInfo.TuNgay = "01/" + periodInfo.Thang + "/" + periodInfo.Nam;
					int month = ConvertHelpers.ConvertToInt(periodInfo.Thang);
					int year = ((ConvertHelpers.ConvertToInt(periodInfo.Nam) == 0) ? DateTime.Now.Year : ConvertHelpers.ConvertToInt(periodInfo.Nam));
					periodInfo.DenNgay = new DateTime(year, month, DateTime.DaysInMonth(year, month)).ToString("dd/MM/yyyy");
				}
				else if (chkTKQuy.Checked)
				{
					periodInfo.KieuKy = "Q";
					periodInfo.Quy = ((cbbQuy.SelectedItem != null) ? cbbQuy.SelectedItem.ToString() : "1");
					periodInfo.Nam = ((Control)mskYear).Text.Trim();
					periodInfo.kyKKhai = periodInfo.Quy + "/" + periodInfo.Nam;
					int year = ConvertHelpers.ConvertToInt(periodInfo.Nam);
					int num = ConvertHelpers.ConvertToInt(periodInfo.Quy);
					periodInfo.TuNgay = new DateTime(year, num * 3 - 2, 1).ToString("dd/MM/yyyy");
					periodInfo.DenNgay = new DateTime(year, num * 3, DateTime.DaysInMonth(year, num * 3)).ToString("dd/MM/yyyy");
					if (menuID != "32" && menuID != "34")
					{
						periodInfo.TuThang = new DateTime(year, num * 3 - 2, 1).ToString("MM/yyyy");
						periodInfo.DenThang = new DateTime(year, num * 3, DateTime.DaysInMonth(year, num * 3)).ToString("MM/yyyy");
					}
				}
				break;
			case "32":
				periodInfo.KieuKy = "Q";
				periodInfo.Quy = cbbQuy.SelectedItem.ToString();
				periodInfo.Nam = ((Control)mskYear).Text.Trim();
				periodInfo.kyKKhai = periodInfo.Quy + "/" + periodInfo.Nam;
				periodInfo.TuNgay = ((Control)mskTuNgay).Text.Trim();
				periodInfo.DenNgay = ((Control)mskDenNgay).Text.Trim();
				break;
			case "33":
				if (chkTkhaiThang.Checked)
				{
					periodInfo.KieuKy = "M";
					periodInfo.Thang = ((Control)mskMonth).Text.Trim().ToString().PadLeft(2, '0');
					periodInfo.Nam = ((((Control)mskYear).Text.Trim().ToString() == "0") ? "1900" : ((Control)mskYear).Text.Trim());
					periodInfo.kyKKhai = periodInfo.Thang + "/" + periodInfo.Nam;
					periodInfo.TuNgay = "01/" + periodInfo.kyKKhai;
					periodInfo.DenNgay = DateTime.DaysInMonth(ConvertHelpers.ConvertToInt(periodInfo.Nam), ConvertHelpers.ConvertToInt(periodInfo.Thang)).ToString().PadLeft(2, '0') + "/" + periodInfo.kyKKhai;
				}
				else
				{
					periodInfo.KieuKy = "D";
					periodInfo.Ngay = ((Control)mskDay).Text.Trim().PadLeft(2, '0');
					periodInfo.Thang = ((Control)mskMonth).Text.Trim().PadLeft(2, '0');
					periodInfo.Nam = ((Control)mskYear).Text.Trim();
					periodInfo.kyKKhai = periodInfo.Ngay + "/" + periodInfo.Thang + "/" + periodInfo.Nam;
					periodInfo.TuNgay = periodInfo.kyKKhai;
					periodInfo.DenNgay = periodInfo.kyKKhai;
				}
				periodInfo.NganhNgheKinhDoanh = ((Control)cboNganhKD).Text;
				periodInfo.TieuMucHachToanVal = ((ListControl)cboNganhKD).SelectedValue.ToString();
				break;
			case "20":
				if (chkTkhaiThang.Checked)
				{
					periodInfo.KieuKy = "M";
					periodInfo.Thang = ((Control)mskMonth).Text.Trim().ToString().PadLeft(2, '0');
					periodInfo.Nam = ((Control)mskYear).Text.Trim().ToString();
					periodInfo.kyKKhai = periodInfo.Thang + "/" + periodInfo.Nam;
					periodInfo.TuNgay = ((Control)mskTuNgay).Text.Trim();
					periodInfo.DenNgay = ((Control)mskDenNgay).Text.Trim();
					if (menuID == "M4")
					{
						periodInfo.LanXuatBan = ((Control)txtLanXuatBan).Text.Trim();
					}
					break;
				}
				periodInfo.KieuKy = "Q";
				periodInfo.Ngay = ((Control)mskDay).Text.Trim().PadLeft(2, '0');
				periodInfo.Thang = ((Control)mskMonth).Text.Trim().PadLeft(2, '0');
				periodInfo.Nam = ((Control)mskYear).Text.Trim();
				periodInfo.TuNgay = ((Control)mskTuNgay).Text.Trim();
				periodInfo.DenNgay = ((Control)mskDenNgay).Text.Trim();
				periodInfo.kyKKhai = periodInfo.Quy + "/" + periodInfo.Nam;
				if (menuID == "M4")
				{
					periodInfo.LanXuatBan = ((Control)txtLanXuatBan).Text.Trim();
				}
				break;
			case "34":
				periodInfo.KieuKy = "Y";
				periodInfo.Nam = ((Control)mskYear).Text.Trim();
				periodInfo.TuThang = ((Control)mskTuThang).Text.Trim();
				periodInfo.DenThang = ((Control)mskDenThang).Text.Trim();
				periodInfo.kyKKhai = ((Control)mskYear).Text.Trim();
				periodInfo.NganhNgheKinhDoanh = ((Control)cboNganhKD).Text;
				periodInfo.TieuMucHachToanVal = ((ListControl)cboNganhKD).SelectedValue.ToString();
				break;
			case "36":
				periodInfo.KieuKy = "Y";
				periodInfo.Nam = ((Control)mskYear).Text.Trim();
				periodInfo.kyKKhai = ((Control)mskYear).Text.Trim();
				if (menuID == "63")
				{
					periodInfo.TuNgay = new DateTime(ConvertHelpers.ConvertToInt(periodInfo.Nam), 1, 1).ToString("dd/MM/yyyy");
					periodInfo.DenNgay = new DateTime(ConvertHelpers.ConvertToInt(periodInfo.Nam), 12, 31).ToString("dd/MM/yyyy");
				}
				break;
			case "0":
			{
				periodInfo.LoaiToKhai = "C";
				periodInfo.KieuKy = "D";
				periodInfo.Ngay = DateTime.Now.Day.ToString().PadLeft(2, '0');
				periodInfo.Thang = DateTime.Now.Month.ToString().PadLeft(2, '0');
				periodInfo.Nam = DateTime.Now.Year.ToString();
				string text4 = (periodInfo.kyKKhai = periodInfo.Ngay + "/" + periodInfo.Thang + "/" + periodInfo.Nam);
				text4 = (periodInfo.DenNgay = text4);
				periodInfo.TuNgay = text4;
				if (menuID == "10" || menuID == "67")
				{
					periodInfo.KieuKy = "Q";
					periodInfo.Quy = ((Control)cbbQuy).Text.Trim();
					periodInfo.Nam = ((Control)mskYear).Text.Trim();
					periodInfo.TuNgay = ((Control)mskDay).Text + "/" + ((Control)mskMonth).Text + "/" + ((Control)mskYear).Text;
					periodInfo.DenNgay = ((Control)mskDay).Text + "/" + ((Control)mskMonth).Text + "/" + ((Control)mskYear).Text;
				}
				break;
			}
			case "35":
				if (chkTKLanPS.Checked)
				{
					if (HTKKFormater.isNullString(((Control)mskDay).Text) || HTKKFormater.isNullString(((Control)mskMonth).Text) || HTKKFormater.isNullString(((Control)mskYear).Text))
					{
						return;
					}
					periodInfo.KieuKy = "D";
					periodInfo.TuNgay = ((Control)mskDay).Text.Trim().PadLeft(2, '0') + "/" + ((Control)mskMonth).Text.Trim().PadLeft(2, '0') + "/" + ((Control)mskYear).Text.Trim();
					periodInfo.DenNgay = ((Control)mskDay).Text.Trim().PadLeft(2, '0') + "/" + ((Control)mskMonth).Text.Trim().PadLeft(2, '0') + "/" + ((Control)mskYear).Text.Trim();
					periodInfo.Ngay = ((Control)mskDay).Text.Trim().PadLeft(2, '0');
					periodInfo.Thang = ((Control)mskMonth).Text.Trim().PadLeft(2, '0');
					periodInfo.Nam = ((Control)mskYear).Text.Trim();
					periodInfo.kyKKhai = periodInfo.Ngay + "/" + periodInfo.Thang + "/" + periodInfo.Nam;
				}
				else if (chkTkhaiThang.Checked)
				{
					if (HTKKFormater.isNullString(((Control)mskMonth).Text) || HTKKFormater.isNullString(((Control)mskYear).Text))
					{
						return;
					}
					periodInfo.KieuKy = "M";
					periodInfo.Thang = ((Control)mskMonth).Text.Trim().PadLeft(2, '0');
					periodInfo.Nam = ((Control)mskYear).Text.Trim();
					periodInfo.kyKKhai = periodInfo.Thang + "/" + periodInfo.Nam;
					periodInfo.TuNgay = "01/" + periodInfo.kyKKhai;
					periodInfo.DenNgay = DateTime.DaysInMonth(ConvertHelpers.ConvertToInt(periodInfo.Nam), ConvertHelpers.ConvertToInt(periodInfo.Thang)).ToString().PadLeft(2, '0') + "/" + periodInfo.kyKKhai;
				}
				else if (chkTKQuy.Checked)
				{
					if (HTKKFormater.isNullString(((Control)cbbQuy).Text) || HTKKFormater.isNullString(((Control)mskYear).Text))
					{
						return;
					}
					periodInfo.KieuKy = "Q";
					periodInfo.Quy = cbbQuy.SelectedItem.ToString();
					periodInfo.Nam = ((Control)mskYear).Text.Trim();
					periodInfo.kyKKhai = periodInfo.Quy + "/" + periodInfo.Nam;
					int year = ConvertHelpers.ConvertToInt(periodInfo.Nam);
					int num = ConvertHelpers.ConvertToInt(periodInfo.Quy);
					periodInfo.TuNgay = new DateTime(year, num * 3 - 2, 1).ToString("dd/MM/yyyy");
					periodInfo.DenNgay = new DateTime(year, num * 3, DateTime.DaysInMonth(year, num * 3)).ToString("dd/MM/yyyy");
				}
				else if (ckNam.Checked)
				{
					if (HTKKFormater.isNullString(((Control)mskYear).Text) || HTKKFormater.isNullString(((Control)mskTuThang).Text) || HTKKFormater.isNullString(((Control)mskDenThang).Text))
					{
						return;
					}
					periodInfo.KieuKy = "Y";
					periodInfo.Nam = ((Control)mskYear).Text.Trim();
					periodInfo.kyKKhai = ((Control)mskYear).Text.Trim();
					periodInfo.TuThang = ((Control)mskTuThang).Text;
					periodInfo.DenThang = ((Control)mskDenThang).Text;
				}
				break;
			case "42":
				if (chkTKKy.Checked)
				{
					periodInfo.KieuKy = "D_D";
					periodInfo.TuNgay = ((Control)mskTuNgay).Text.Trim();
					periodInfo.DenNgay = ((Control)mskDenNgay).Text.Trim();
					periodInfo.kyKKhai = periodInfo.TuNgay;
					if (!HTKKFormater.isNullString(((Control)mskTuNgay).Text.Trim().Replace("/", "").Replace("_", "")))
					{
						periodInfo.Nam = ((Control)mskTuNgay).Text.Trim().Replace("/", "").Substring(4, 4);
					}
				}
				else if (chkTkhaiThang.Checked)
				{
					if (HTKKFormater.isNullString(((Control)mskMonth).Text) || HTKKFormater.isNullString(((Control)mskYear).Text))
					{
						return;
					}
					periodInfo.KieuKy = "M";
					periodInfo.Thang = ((Control)mskMonth).Text.Trim().PadLeft(2, '0');
					periodInfo.Nam = ((Control)mskYear).Text.Trim();
					periodInfo.kyKKhai = periodInfo.Thang + "/" + periodInfo.Nam;
					periodInfo.TuNgay = "01/" + periodInfo.kyKKhai;
					periodInfo.DenNgay = DateTime.DaysInMonth(ConvertHelpers.ConvertToInt(periodInfo.Nam), ConvertHelpers.ConvertToInt(periodInfo.Thang)).ToString().PadLeft(2, '0') + "/" + periodInfo.kyKKhai;
				}
				else if (chkTKQuy.Checked)
				{
					if (HTKKFormater.isNullString(((Control)cbbQuy).Text) || HTKKFormater.isNullString(((Control)mskYear).Text))
					{
						return;
					}
					periodInfo.KieuKy = "Q";
					periodInfo.Quy = cbbQuy.SelectedItem.ToString();
					periodInfo.Nam = ((Control)mskYear).Text.Trim();
					periodInfo.kyKKhai = periodInfo.Quy + "/" + periodInfo.Nam;
					int year = ConvertHelpers.ConvertToInt(periodInfo.Nam);
					int num = ConvertHelpers.ConvertToInt(periodInfo.Quy);
					periodInfo.TuNgay = new DateTime(year, num * 3 - 2, 1).ToString("dd/MM/yyyy");
					periodInfo.DenNgay = new DateTime(year, num * 3, DateTime.DaysInMonth(year, num * 3)).ToString("dd/MM/yyyy");
					periodInfo.TuThang = ((Control)mskTuThang).Text;
					periodInfo.DenThang = ((Control)mskDenThang).Text;
					bool visible = ((Control)pnVT1_3TuNgayDenNgay).Visible;
					bool visible2 = ((Control)mskTuThang).Visible;
					bool visible3 = ((Control)mskDenThang).Visible;
				}
				else if (ckNam.Checked)
				{
					if (HTKKFormater.isNullString(((Control)mskYear).Text) || HTKKFormater.isNullString(((Control)mskTuThang).Text) || HTKKFormater.isNullString(((Control)mskDenThang).Text))
					{
						return;
					}
					periodInfo.KieuKy = "Y";
					periodInfo.Nam = ((Control)mskYear).Text.Trim();
					periodInfo.kyKKhai = ((Control)mskYear).Text.Trim();
				}
				break;
			case "15":
				periodInfo.KieuKy = "Q";
				periodInfo.Nam = ((Control)mskYear).Text.Trim();
				periodInfo.TuNgay = ((Control)mskTuNgay).Text.Trim();
				periodInfo.DenNgay = ((Control)mskDenNgay).Text.Trim();
				periodInfo.kyKKhai = periodInfo.Quy + "/" + periodInfo.Nam;
				break;
			case "29":
				if (chkTKQuy.Checked)
				{
					periodInfo.KieuKy = "Q";
					periodInfo.Quy = ConvertHelpers.ConvertToInt(cbbQuy.SelectedItem).ToString();
					periodInfo.Nam = ((Control)mskYear).Text.Trim();
					periodInfo.kyKKhai = periodInfo.Quy + "/" + periodInfo.Nam;
					if (periodInfo.menuID == "46")
					{
						if (periodInfo.Nam == "0")
						{
							periodInfo.Nam = DateTime.Now.Year.ToString();
						}
						periodInfo.TuNgay = ConvertHelpers.GetNgayDauQuy(ConvertHelpers.ConvertToInt(cbbQuy.SelectedItem), ConvertHelpers.ConvertToInt(((Control)mskYear).Text.Trim()), 1, 1).ToString("dd/MM/yyyy");
						periodInfo.DenNgay = ConvertHelpers.GetNgayCuoiQuy(ConvertHelpers.ConvertToInt(cbbQuy.SelectedItem), ConvertHelpers.ConvertToInt(((Control)mskYear).Text.Trim()), 1, 1).ToString("dd/MM/yyyy");
					}
					break;
				}
				periodInfo.KieuKy = "Y";
				periodInfo.Nam = ((Control)mskYear).Text.Trim();
				periodInfo.kyKKhai = ((Control)mskYear).Text.Trim();
				if (periodInfo.menuID == "46")
				{
					if (periodInfo.Nam == "0")
					{
						periodInfo.Nam = DateTime.Now.Year.ToString();
					}
					DateTime dateTime = new DateTime(int.Parse(periodInfo.Nam), 1, 1);
					DateTime dateTime2 = dateTime.AddYears(1).AddDays(-1.0);
					periodInfo.TuNgay = dateTime.Day.ToString().PadLeft(2, '0') + "/" + dateTime.Month.ToString().PadLeft(2, '0') + "/" + dateTime.Year;
					periodInfo.DenNgay = dateTime2.Day.ToString().PadLeft(2, '0') + "/" + dateTime2.Month.ToString().PadLeft(2, '0') + "/" + dateTime2.Year;
				}
				break;
			case "37":
				if (chkTKKy.Checked)
				{
					periodInfo.KieuKy = "Q";
					periodInfo.TuNgay = ((Control)mskTuNgay).Text.Trim();
					periodInfo.DenNgay = ((Control)mskDenNgay).Text.Trim();
					periodInfo.kyKKhai = periodInfo.TuNgay;
					if (!HTKKFormater.isNullString(((Control)mskTuNgay).Text.Trim().Replace("/", "").Replace("_", "")))
					{
						periodInfo.Nam = ((Control)mskTuNgay).Text.Trim().Replace("/", "").Substring(4, 4);
					}
				}
				if (ckNam.Checked)
				{
					periodInfo.KieuKy = "Y";
					periodInfo.Nam = ((Control)mskYear).Text.Trim();
					periodInfo.kyKKhai = periodInfo.Nam;
					DateTime dateTime = new DateTime(int.Parse(periodInfo.Nam), 1, 1);
					DateTime dateTime2 = dateTime.AddYears(1).AddDays(-1.0);
					periodInfo.TuNgay = dateTime.Day.ToString().PadLeft(2, '0') + "/" + dateTime.Month.ToString().PadLeft(2, '0') + "/" + dateTime.Year;
					periodInfo.DenNgay = dateTime2.Day.ToString().PadLeft(2, '0') + "/" + dateTime2.Month.ToString().PadLeft(2, '0') + "/" + dateTime2.Year;
				}
				break;
			case "7":
				if (chkTKKy.Checked)
				{
					periodInfo.KieuKy = "K";
				}
				else
				{
					periodInfo.KieuKy = "Q";
				}
				periodInfo.Nam = ((Control)mskYear).Text.Trim();
				periodInfo.TuNgay = ((Control)mskTuNgay).Text.Trim();
				periodInfo.DenNgay = ((Control)mskDenNgay).Text.Trim();
				periodInfo.kyKKhai = periodInfo.Quy + "/" + periodInfo.Nam;
				break;
			case "38":
				periodInfo.KieuKy = "Q";
				periodInfo.Quy = cbbQuy.SelectedItem.ToString();
				periodInfo.Nam = ((Control)mskYear).Text.Trim();
				if (!string.IsNullOrEmpty(((Control)mskTuNgay).Text.Trim().Trim(new char[1] { '/' })))
				{
					periodInfo.TuNgay = ((Control)mskTuNgay).Text.Trim();
					periodInfo.DenNgay = ((Control)mskDenNgay).Text.Trim();
				}
				periodInfo.kyKKhai = periodInfo.Quy + "/" + periodInfo.Nam;
				periodInfo.DonViTinhTen = ((Control)cboNganhKD).Text;
				periodInfo.DonViTinhMa = ((ListControl)cboNganhKD).SelectedValue.ToString();
				break;
			}
			if (menuID == "R9")
			{
				periodInfo.LanXuatBan = ((Control)txtLanXuatBan).Text;
			}
			if (menuID == "X1" || menuID == "X2")
			{
				periodInfo.kyKKhai = ((ListControl)cbbKy).SelectedIndex.ToString();
			}
			if (((Control)cbMaHSDNHT_DADT).Visible && "U8~".Contains(WinSettings.MenuID))
			{
				periodInfo.LanXuatBan = ((Control)cbMaHSDNHT_DADT).Text;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void bindingDataToComboBox(ComboBox combobox, Dictionary<string, string> data)
	{
		//IL_0004: Unknown result type (might be due to invalid IL or missing references)
		//IL_000e: Expected O, but got Unknown
		combobox.DataSource = (object)new BindingSource((object)data, (string)null);
		((ListControl)combobox).DisplayMember = "Value";
		((ListControl)combobox).ValueMember = "Key";
		((Control)combobox).Refresh();
	}

	private void cbbQuy_SelectedIndexChanged(object sender, EventArgs e)
	{
		try
		{
			((Control)mskTuThang).Text = GeneralFunctions.QuarterToFirstMonth(((ListControl)cbbQuy).SelectedIndex + 1).ToString().PadLeft(2, '0') + "/" + ((Control)mskYear).Text;
			((Control)mskDenThang).Text = GeneralFunctions.QuarterToLastMonth(((ListControl)cbbQuy).SelectedIndex + 1).ToString().PadLeft(2, '0') + "/" + ((Control)mskYear).Text;
			if (strKieuKy == "32" || strKieuKy == "15" || strKieuKy == "20" || strKieuKy == "7")
			{
				int month = GeneralFunctions.QuarterToFirstMonth(((ListControl)cbbQuy).SelectedIndex + 1);
				int month2 = GeneralFunctions.QuarterToLastMonth(((ListControl)cbbQuy).SelectedIndex + 1);
				if (int.TryParse(((Control)mskYear).Text.Trim(), out var result))
				{
					DateTime dateTime = new DateTime(result, month, 1);
					DateTime dateTime2 = new DateTime(result, month2, DateTime.DaysInMonth(result, month2));
					((Control)mskTuNgay).Text = dateTime.ToString("dd/MM/yyyy");
					((Control)mskDenNgay).Text = dateTime2.ToString("dd/MM/yyyy");
				}
			}
			if (strKieuKy == "38" && ConvertHelpers.ConvertToInt(((Control)cbbQuy).Text) > 0 && ConvertHelpers.ConvertToInt(((Control)mskYear).Text) > 0)
			{
				setQuyByNgayBatDauTC(ConvertHelpers.ConvertToInt(((Control)cbbQuy).Text), ConvertHelpers.ConvertToInt(((Control)mskYear).Text));
			}
			if ((menuID == "X1" || menuID == "X2") && ConvertHelpers.ConvertToInt(((Control)cbbQuy).Text) > 0 && ConvertHelpers.ConvertToInt(((Control)mskYear).Text) > 0)
			{
				setQuyByNgayBatDauTC(ConvertHelpers.ConvertToInt(((Control)cbbQuy).Text), ConvertHelpers.ConvertToInt(((Control)mskYear).Text));
			}
			if (((Control)pnVT2_PhuLucKeKhai).Visible)
			{
				LoadPhuLucSpread();
			}
			if (menuID == "31")
			{
				Period_31();
			}
			loadDMNN_01GTGT();
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void axfpsPhuLuc_ButtonClicked(object sender, EditorNotifyEventArgs e)
	{
		try
		{
			int num = -1;
			int num2 = 2;
			for (int i = 0; i < axfpsPhuLuc.ActiveSheet.RowCount; i++)
			{
				if (axfpsPhuLuc.ActiveSheet.Cells[i, e.Column].Value.ToString().ToLower() == "false")
				{
					num2 = 0;
					break;
				}
				num2 = 1;
			}
			for (int i = 0; i < axfpsPhuLuc.ActiveSheet.RowCount; i++)
			{
				if (axfpsPhuLuc.ActiveSheet.Cells[i, e.Column].Value.ToString().ToLower() == "true")
				{
					num = i;
				}
			}
			if (menuID.Equals("18") || menuID.Equals("68") || menuID.Equals("Q9") || menuID == "Y7")
			{
				((Control)chkChonPhuLucKeKhaiAll).Enabled = false;
				if (e.Row == 1)
				{
					axfpsPhuLuc.ActiveSheet.Cells[0, 0].Value = false;
				}
				else
				{
					axfpsPhuLuc.ActiveSheet.Cells[1, 0].Value = false;
				}
			}
			else if ("78,53,54,58,A1,H1,A8,H8,A2,A5,A6,B3,B4,B5,U9,X1".Contains(menuID))
			{
				((Control)chkChonPhuLucKeKhaiAll).Enabled = false;
				if (e.Row == 1)
				{
					if (!axfpsPhuLuc.ActiveSheet.Cells[2, 0].Locked && axfpsPhuLuc.ActiveSheet.Cells[1, 0].Value.ToString().ToLower() == "true")
					{
						axfpsPhuLuc.ActiveSheet.Cells[2, 0].Value = false;
					}
				}
				else if (e.Row == 2 && !axfpsPhuLuc.ActiveSheet.Cells[1, 0].Locked && axfpsPhuLuc.ActiveSheet.Cells[2, 0].Value.ToString().ToLower() == "true")
				{
					axfpsPhuLuc.ActiveSheet.Cells[1, 0].Value = false;
				}
			}
			else if (menuID.Equals("13") || menuID.Equals("14") || menuID.Equals("62"))
			{
				((Control)chkChonPhuLucKeKhaiAll).Enabled = false;
				for (int i = 0; i < axfpsPhuLuc.ActiveSheet.RowCount; i++)
				{
					if (i == num || num == -1)
					{
						axfpsPhuLuc.ActiveSheet.Cells[i, 0].Locked = false;
					}
					else
					{
						axfpsPhuLuc.ActiveSheet.Cells[i, 0].Locked = true;
					}
				}
			}
			else
			{
				switch (num2)
				{
				case 2:
					((Control)chkChonPhuLucKeKhaiAll).Enabled = false;
					chkChonPhuLucKeKhaiAll.Checked = true;
					break;
				case 1:
					((Control)chkChonPhuLucKeKhaiAll).Enabled = true;
					chkChonPhuLucKeKhaiAll.Checked = true;
					break;
				default:
					((Control)chkChonPhuLucKeKhaiAll).Enabled = true;
					chkChonPhuLucKeKhaiAll.Checked = false;
					break;
				}
			}
			if (menuID.Equals("Q9"))
			{
				bool enabled = (bool)axfpsPhuLuc.ActiveSheet.Cells[1, 0].Value;
				((Control)pnTinhCQT).Enabled = enabled;
			}
			if (menuID == "Y7")
			{
				bool enabled = (bool)axfpsPhuLuc.ActiveSheet.Cells[1, 0].Value;
				((Control)pnTinhCQT).Enabled = enabled;
				((Control)pnCQTCuc).Enabled = false;
				((Control)pnCCT).Enabled = false;
				if ((bool)axfpsPhuLuc.ActiveSheet.Cells[1, 0].Value)
				{
					((Control)chkTKQuy).Enabled = false;
					chkTKQuy.Checked = false;
				}
				else
				{
					((Control)chkTKQuy).Enabled = true;
				}
				if ((bool)axfpsPhuLuc.ActiveSheet.Cells[0, 0].Value)
				{
					((Control)chkTKLanPS).Enabled = false;
					chkTKLanPS.Checked = false;
				}
				else
				{
					((Control)chkTKLanPS).Enabled = true;
				}
			}
			if (axfpsPhuLuc.ActiveSheet.Cells[e.Row, e.Column].Value.ToString().ToLower() == "false")
			{
				chkChonPhuLucKeKhaiAll.Checked = false;
				return;
			}
			for (int i = 0; i < axfpsPhuLuc.ActiveSheet.RowCount; i++)
			{
				if (axfpsPhuLuc.ActiveSheet.Cells[i, e.Column].Value.ToString().ToLower() == "false")
				{
					chkChonPhuLucKeKhaiAll.Checked = false;
					return;
				}
			}
			chkChonPhuLucKeKhaiAll.Checked = true;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void chkChonPhuLucKeKhaiAll_Click(object sender, EventArgs e)
	{
		try
		{
			if (chkChonPhuLucKeKhaiAll.Checked)
			{
				axfpsPhuLucSetAllValues(value: true);
			}
			else
			{
				axfpsPhuLucSetAllValues(value: false);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void cbbQuy_MouseLeave(object sender, EventArgs e)
	{
		if (menuID == "31")
		{
			Period_31();
		}
	}

	private void cbbQuy_Leave(object sender, EventArgs e)
	{
		if (menuID == "31")
		{
			Period_31();
		}
	}

	private bool ValidateBaseControl()
	{
		//IL_004a: Unknown result type (might be due to invalid IL or missing references)
		//IL_00a7: Unknown result type (might be due to invalid IL or missing references)
		//IL_0109: Unknown result type (might be due to invalid IL or missing references)
		//IL_016b: Unknown result type (might be due to invalid IL or missing references)
		//IL_01cd: Unknown result type (might be due to invalid IL or missing references)
		//IL_022f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0291: Unknown result type (might be due to invalid IL or missing references)
		//IL_02f3: Unknown result type (might be due to invalid IL or missing references)
		//IL_0355: Unknown result type (might be due to invalid IL or missing references)
		//IL_03ed: Unknown result type (might be due to invalid IL or missing references)
		if (((Control)mskMonth).Text.Trim().Length == 0 && ((Control)mskMonth).Visible)
		{
			frmMessage frmMessage = new frmMessage("0017", "3", "1", "Thông báo");
			((Form)frmMessage).ShowDialog();
			((Control)mskMonth).Focus();
			return false;
		}
		if (((Control)cbbQuy).Text.Length == 0 && ((Control)cbbQuy).Visible)
		{
			frmMessage frmMessage = new frmMessage("0017", "3", "1", "Thông báo");
			((Form)frmMessage).ShowDialog();
			((Control)cbbQuy).Focus();
			return false;
		}
		if (((Control)mskDay).Text.Trim().Length == 0 && ((Control)mskDay).Visible)
		{
			frmMessage frmMessage = new frmMessage("0017", "3", "1", "Thông báo");
			((Form)frmMessage).ShowDialog();
			((Control)mskDay).Focus();
			return false;
		}
		if (((Control)mskYear).Text.Trim().Length == 0 && ((Control)mskYear).Visible)
		{
			frmMessage frmMessage = new frmMessage("0017", "3", "1", "Thông báo");
			((Form)frmMessage).ShowDialog();
			((Control)mskYear).Focus();
			return false;
		}
		if (((Control)mskTuNgay).Text.Trim().Length == 0 && ((Control)mskTuNgay).Visible)
		{
			frmMessage frmMessage = new frmMessage("0017", "3", "1", "Thông báo");
			((Form)frmMessage).ShowDialog();
			((Control)mskTuNgay).Focus();
			return false;
		}
		if (((Control)mskDenNgay).Text.Trim().Length == 0 && ((Control)mskDenNgay).Visible)
		{
			frmMessage frmMessage = new frmMessage("0017", "3", "1", "Thông báo");
			((Form)frmMessage).ShowDialog();
			((Control)mskDenNgay).Focus();
			return false;
		}
		if (((Control)mskTuThang).Text.Trim().Length == 0 && ((Control)mskTuThang).Visible)
		{
			frmMessage frmMessage = new frmMessage("0017", "3", "1", "Thông báo");
			((Form)frmMessage).ShowDialog();
			((Control)mskTuThang).Focus();
			return false;
		}
		if (((Control)mskDenThang).Text.Trim().Length == 0 && ((Control)mskDenThang).Visible)
		{
			frmMessage frmMessage = new frmMessage("0017", "3", "1", "Thông báo");
			((Form)frmMessage).ShowDialog();
			((Control)mskDenThang).Focus();
			return false;
		}
		if (((Control)txtLan).Text.Trim().Length == 0 && ((Control)txtLan).Visible)
		{
			frmMessage frmMessage = new frmMessage("0017", "3", "1", "Thông báo");
			((Form)frmMessage).ShowDialog();
			((Control)txtLan).Focus();
			return false;
		}
		if (((Control)txtLanXuatBan).Text.Trim().Length == 0 && ((Control)txtLanXuatBan).Visible && (!(menuID == "R9") || !(((Control)cboNganhKD).Text == "Khí thiên nhiên")))
		{
			frmMessage frmMessage = new frmMessage("0017", "3", "1", "Thông báo");
			((Form)frmMessage).ShowDialog();
			((Control)txtLanXuatBan).Focus();
			return false;
		}
		return true;
	}

	private bool CheckValidpnVT1_1CheckBox()
	{
		//IL_004b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0051: Expected O, but got Unknown
		//IL_0070: Unknown result type (might be due to invalid IL or missing references)
		//IL_007a: Expected O, but got Unknown
		//IL_007c: Unknown result type (might be due to invalid IL or missing references)
		//IL_014e: Unknown result type (might be due to invalid IL or missing references)
		try
		{
			if (pnVT1_1CheckBox == null || !((Control)pnVT1_1CheckBox).Visible)
			{
				return true;
			}
			List<CheckBox> list = new List<CheckBox>();
			int num = 0;
			foreach (Control item in (ArrangedElementCollection)((Control)pnVT1_1CheckBox).Controls)
			{
				Control val = item;
				if (val is CheckBox && val.Visible)
				{
					list.Add((CheckBox)val);
					if (((CheckBox)val).Checked)
					{
						num++;
					}
				}
			}
			if (list.Count >= 2 && num == 0)
			{
				CheckBox val2 = (from c in list
					orderby ((Control)c).Location.Y, ((Control)c).Location.X
					select c).FirstOrDefault();
				frmMessage frmMessage = new frmMessage("0062", "3", "1", "Thông báo");
				((Form)frmMessage).ShowDialog();
				if (val2 != null)
				{
					((Control)val2).Focus();
				}
				return false;
			}
			return true;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return false;
		}
	}

	private bool ExistTokhaiThangQuy(string sDataFile, bool isThang, HTKKPeriodInfo periodInfo)
	{
		bool result = false;
		string text = "";
		if (ConvertHelpers.ConvertToInt(periodInfo.Nam) >= 2021)
		{
			if (periodInfo.menuID.Equals("01") || periodInfo.menuID.Equals("D1") || periodInfo.menuID.Equals("I5") || WinSettings.MenuID.Equals("Z5"))
			{
				text = periodInfo.TieuMucHachToanVal;
			}
			else if (periodInfo.menuID.Equals("02") || periodInfo.menuID.Equals("D2"))
			{
				text = periodInfo.MaHSDuAnDT;
			}
			else if (periodInfo.menuID.Equals("71") || periodInfo.menuID.Equals("D4") || WinSettings.MenuID.Equals("Z3"))
			{
				text = periodInfo.ThuHo;
			}
		}
		try
		{
			string thang = periodInfo.Thang;
			int num = ConvertHelpers.ConvertToInt(periodInfo.Quy);
			string nam = periodInfo.Nam;
			string text2 = "";
			string value = "";
			string value2 = "";
			string path = HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst;
			string[] files = Directory.GetFiles(path, "*.xml");
			if (isThang)
			{
				text2 = (text.Equals("") ? ((!thang.Equals("01") && !thang.Equals("02") && !thang.Equals("03")) ? ((!thang.Equals("04") && !thang.Equals("05") && !thang.Equals("06")) ? ((!thang.Equals("07") && !thang.Equals("08") && !thang.Equals("09")) ? (sDataFile + ((!string.IsNullOrEmpty(WinDoWorkEvent.GetNameNopTheoTinhCQT())) ? ("_" + WinDoWorkEvent.GetNameNopTheoTinhCQT()) : "") + "_Q04" + nam + "_" + periodInfo.VersionXML.Replace(".", "")) : (sDataFile + ((!string.IsNullOrEmpty(WinDoWorkEvent.GetNameNopTheoTinhCQT())) ? ("_" + WinDoWorkEvent.GetNameNopTheoTinhCQT()) : "") + "_Q03" + nam + "_" + periodInfo.VersionXML.Replace(".", ""))) : (sDataFile + ((!string.IsNullOrEmpty(WinDoWorkEvent.GetNameNopTheoTinhCQT())) ? ("_" + WinDoWorkEvent.GetNameNopTheoTinhCQT()) : "") + "_Q02" + nam + "_" + periodInfo.VersionXML.Replace(".", ""))) : (sDataFile + ((!string.IsNullOrEmpty(WinDoWorkEvent.GetNameNopTheoTinhCQT())) ? ("_" + WinDoWorkEvent.GetNameNopTheoTinhCQT()) : "") + "_Q01" + nam + "_" + periodInfo.VersionXML.Replace(".", ""))) : ((!thang.Equals("01") && !thang.Equals("02") && !thang.Equals("03")) ? ((!thang.Equals("04") && !thang.Equals("05") && !thang.Equals("06")) ? ((!thang.Equals("07") && !thang.Equals("08") && !thang.Equals("09")) ? (sDataFile + ((!string.IsNullOrEmpty(WinDoWorkEvent.GetNameNopTheoTinhCQT())) ? ("_" + WinDoWorkEvent.GetNameNopTheoTinhCQT()) : "") + "_" + text + "_Q04" + nam + "_" + periodInfo.VersionXML.Replace(".", "")) : (sDataFile + ((!string.IsNullOrEmpty(WinDoWorkEvent.GetNameNopTheoTinhCQT())) ? ("_" + WinDoWorkEvent.GetNameNopTheoTinhCQT()) : "") + "_" + text + "_Q03" + nam + "_" + periodInfo.VersionXML.Replace(".", ""))) : (sDataFile + ((!string.IsNullOrEmpty(WinDoWorkEvent.GetNameNopTheoTinhCQT())) ? ("_" + WinDoWorkEvent.GetNameNopTheoTinhCQT()) : "") + "_" + text + "_Q02" + nam + "_" + periodInfo.VersionXML.Replace(".", ""))) : (sDataFile + ((!string.IsNullOrEmpty(WinDoWorkEvent.GetNameNopTheoTinhCQT())) ? ("_" + WinDoWorkEvent.GetNameNopTheoTinhCQT()) : "") + "_" + text + "_Q01" + nam + "_" + periodInfo.VersionXML.Replace(".", ""))));
			}
			else if (text.Equals(""))
			{
				switch (num)
				{
				case 1:
					text2 = sDataFile + ((!string.IsNullOrEmpty(WinDoWorkEvent.GetNameNopTheoTinhCQT())) ? ("_" + WinDoWorkEvent.GetNameNopTheoTinhCQT()) : "") + "_01" + nam + "_" + periodInfo.VersionXML.Replace(".", "");
					value = sDataFile + ((!string.IsNullOrEmpty(WinDoWorkEvent.GetNameNopTheoTinhCQT())) ? ("_" + WinDoWorkEvent.GetNameNopTheoTinhCQT()) : "") + "_02" + nam + "_" + periodInfo.VersionXML.Replace(".", "");
					value2 = sDataFile + ((!string.IsNullOrEmpty(WinDoWorkEvent.GetNameNopTheoTinhCQT())) ? ("_" + WinDoWorkEvent.GetNameNopTheoTinhCQT()) : "") + "_03" + nam + "_" + periodInfo.VersionXML.Replace(".", "");
					break;
				case 2:
					text2 = sDataFile + ((!string.IsNullOrEmpty(WinDoWorkEvent.GetNameNopTheoTinhCQT())) ? ("_" + WinDoWorkEvent.GetNameNopTheoTinhCQT()) : "") + "_04" + nam + "_" + periodInfo.VersionXML.Replace(".", "");
					value = sDataFile + ((!string.IsNullOrEmpty(WinDoWorkEvent.GetNameNopTheoTinhCQT())) ? ("_" + WinDoWorkEvent.GetNameNopTheoTinhCQT()) : "") + "_05" + nam + "_" + periodInfo.VersionXML.Replace(".", "");
					value2 = sDataFile + ((!string.IsNullOrEmpty(WinDoWorkEvent.GetNameNopTheoTinhCQT())) ? ("_" + WinDoWorkEvent.GetNameNopTheoTinhCQT()) : "") + "_06" + nam + "_" + periodInfo.VersionXML.Replace(".", "");
					break;
				case 3:
					text2 = sDataFile + ((!string.IsNullOrEmpty(WinDoWorkEvent.GetNameNopTheoTinhCQT())) ? ("_" + WinDoWorkEvent.GetNameNopTheoTinhCQT()) : "") + "_07" + nam + "_" + periodInfo.VersionXML.Replace(".", "");
					value = sDataFile + ((!string.IsNullOrEmpty(WinDoWorkEvent.GetNameNopTheoTinhCQT())) ? ("_" + WinDoWorkEvent.GetNameNopTheoTinhCQT()) : "") + "_08" + nam + "_" + periodInfo.VersionXML.Replace(".", "");
					value2 = sDataFile + ((!string.IsNullOrEmpty(WinDoWorkEvent.GetNameNopTheoTinhCQT())) ? ("_" + WinDoWorkEvent.GetNameNopTheoTinhCQT()) : "") + "_09" + nam + "_" + periodInfo.VersionXML.Replace(".", "");
					break;
				default:
					text2 = sDataFile + ((!string.IsNullOrEmpty(WinDoWorkEvent.GetNameNopTheoTinhCQT())) ? ("_" + WinDoWorkEvent.GetNameNopTheoTinhCQT()) : "") + "_10" + nam + "_" + periodInfo.VersionXML.Replace(".", "");
					value = sDataFile + ((!string.IsNullOrEmpty(WinDoWorkEvent.GetNameNopTheoTinhCQT())) ? ("_" + WinDoWorkEvent.GetNameNopTheoTinhCQT()) : "") + "_11" + nam + "_" + periodInfo.VersionXML.Replace(".", "");
					value2 = sDataFile + ((!string.IsNullOrEmpty(WinDoWorkEvent.GetNameNopTheoTinhCQT())) ? ("_" + WinDoWorkEvent.GetNameNopTheoTinhCQT()) : "") + "_12" + nam + "_" + periodInfo.VersionXML.Replace(".", "");
					break;
				}
			}
			else
			{
				switch (num)
				{
				case 1:
					text2 = sDataFile + ((!string.IsNullOrEmpty(WinDoWorkEvent.GetNameNopTheoTinhCQT())) ? ("_" + WinDoWorkEvent.GetNameNopTheoTinhCQT()) : "") + "_" + text + "_01" + nam + "_" + periodInfo.VersionXML.Replace(".", "");
					value = sDataFile + ((!string.IsNullOrEmpty(WinDoWorkEvent.GetNameNopTheoTinhCQT())) ? ("_" + WinDoWorkEvent.GetNameNopTheoTinhCQT()) : "") + "_" + text + "_02" + nam + "_" + periodInfo.VersionXML.Replace(".", "");
					value2 = sDataFile + ((!string.IsNullOrEmpty(WinDoWorkEvent.GetNameNopTheoTinhCQT())) ? ("_" + WinDoWorkEvent.GetNameNopTheoTinhCQT()) : "") + "_" + text + "_03" + nam + "_" + periodInfo.VersionXML.Replace(".", "");
					break;
				case 2:
					text2 = sDataFile + ((!string.IsNullOrEmpty(WinDoWorkEvent.GetNameNopTheoTinhCQT())) ? ("_" + WinDoWorkEvent.GetNameNopTheoTinhCQT()) : "") + "_" + text + "_04" + nam + "_" + periodInfo.VersionXML.Replace(".", "");
					value = sDataFile + ((!string.IsNullOrEmpty(WinDoWorkEvent.GetNameNopTheoTinhCQT())) ? ("_" + WinDoWorkEvent.GetNameNopTheoTinhCQT()) : "") + "_" + text + "_05" + nam + "_" + periodInfo.VersionXML.Replace(".", "");
					value2 = sDataFile + ((!string.IsNullOrEmpty(WinDoWorkEvent.GetNameNopTheoTinhCQT())) ? ("_" + WinDoWorkEvent.GetNameNopTheoTinhCQT()) : "") + "_" + text + "_06" + nam + "_" + periodInfo.VersionXML.Replace(".", "");
					break;
				case 3:
					text2 = sDataFile + ((!string.IsNullOrEmpty(WinDoWorkEvent.GetNameNopTheoTinhCQT())) ? ("_" + WinDoWorkEvent.GetNameNopTheoTinhCQT()) : "") + "_" + text + "_07" + nam + "_" + periodInfo.VersionXML.Replace(".", "");
					value = sDataFile + ((!string.IsNullOrEmpty(WinDoWorkEvent.GetNameNopTheoTinhCQT())) ? ("_" + WinDoWorkEvent.GetNameNopTheoTinhCQT()) : "") + "_" + text + "_08" + nam + "_" + periodInfo.VersionXML.Replace(".", "");
					value2 = sDataFile + ((!string.IsNullOrEmpty(WinDoWorkEvent.GetNameNopTheoTinhCQT())) ? ("_" + WinDoWorkEvent.GetNameNopTheoTinhCQT()) : "") + "_" + text + "_09" + nam + "_" + periodInfo.VersionXML.Replace(".", "");
					break;
				default:
					text2 = sDataFile + ((!string.IsNullOrEmpty(WinDoWorkEvent.GetNameNopTheoTinhCQT())) ? ("_" + WinDoWorkEvent.GetNameNopTheoTinhCQT()) : "") + "_" + text + "_10" + nam + "_" + periodInfo.VersionXML.Replace(".", "");
					value = sDataFile + ((!string.IsNullOrEmpty(WinDoWorkEvent.GetNameNopTheoTinhCQT())) ? ("_" + WinDoWorkEvent.GetNameNopTheoTinhCQT()) : "") + "_" + text + "_11" + nam + "_" + periodInfo.VersionXML.Replace(".", "");
					value2 = sDataFile + ((!string.IsNullOrEmpty(WinDoWorkEvent.GetNameNopTheoTinhCQT())) ? ("_" + WinDoWorkEvent.GetNameNopTheoTinhCQT()) : "") + "_" + text + "_12" + nam + "_" + periodInfo.VersionXML.Replace(".", "");
					break;
				}
			}
			for (int i = 0; i < files.Length; i++)
			{
				if (isThang)
				{
					if (files[i].IndexOf(text2) > 0)
					{
						result = true;
						break;
					}
				}
				else if (files[i].IndexOf(text2) > 0 || files[i].IndexOf(value) > 0 || files[i].IndexOf(value2) > 0)
				{
					result = true;
					break;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return result;
	}

	private bool CheckKyKKTrung(string sDataFile, HTKKPeriodInfo periodInfo)
	{
		bool result = false;
		try
		{
			string tuThang = periodInfo.TuThang;
			string denThang = periodInfo.DenThang;
			string nam = periodInfo.Nam;
			string path = HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst;
			string[] files = Directory.GetFiles(path, "*.xml");
			string toMonth = "";
			string text = "";
			string text2 = "";
			for (int i = 0; i < files.Length; i++)
			{
				text2 = files[i].Substring(files[i].LastIndexOf('\\') + 1);
				if (text2.Length > 26 && sDataFile.Equals(text2.Substring(0, text2.Length - 27)))
				{
					string[] array = text2.Substring(sDataFile.Length + 6, 13).Split(new char[1] { '_' });
					if (array != null && array.Length > 1)
					{
						toMonth = array[0];
						text = array[1];
					}
					int soThang = DateUtils.GetSoThang(tuThang.Replace("/", ""), toMonth);
					int soThang2 = DateUtils.GetSoThang(text, denThang.Replace("/", ""));
					if (soThang == 0 && soThang2 == 0)
					{
						result = false;
						break;
					}
					if (soThang * soThang2 > 0 && text2.Substring(sDataFile.Length + 1, 4).Equals(nam))
					{
						result = true;
						break;
					}
					soThang = DateUtils.GetSoThang(tuThang.Replace("/", ""), toMonth);
					soThang2 = DateUtils.GetSoThang(tuThang.Replace("/", ""), text);
					if (soThang * soThang2 <= 0 && text2.Substring(sDataFile.Length + 1, 4).Equals(nam))
					{
						result = true;
						break;
					}
					soThang = DateUtils.GetSoThang(denThang.Replace("/", ""), toMonth);
					soThang2 = DateUtils.GetSoThang(denThang.Replace("/", ""), text);
					if (soThang * soThang2 <= 0 && text2.Substring(sDataFile.Length + 1, 4).Equals(nam))
					{
						result = true;
						break;
					}
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			result = false;
		}
		return result;
	}

	private bool CheckKyKKTrungNgay(string sDataFile, HTKKPeriodInfo periodInfo)
	{
		bool result = false;
		try
		{
			string tuNgay = periodInfo.TuNgay;
			string denNgay = periodInfo.DenNgay;
			string nam = periodInfo.Nam;
			string text = "";
			string text2 = "";
			int num = 0;
			int num2 = 0;
			string path = HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst;
			string[] files = Directory.GetFiles(path, "*.xml");
			string text3 = "";
			string text4 = "";
			for (int i = 0; i < files.Length; i++)
			{
				text3 = files[i].Substring(files[i].LastIndexOf('\\') + 1);
				if (text3.Length <= 31 || text3.IndexOf(sDataFile) != 0)
				{
					continue;
				}
				string[] array = text3.Replace(sDataFile + "_", "").Split(new char[1] { '_' });
				if (array != null && array.Length > 2)
				{
					text4 = array[0];
					text = array[1];
					text2 = array[2];
					DateTime dateTime = ConvertHelpers.ConvertToDateTime(HTKKFormater.FormatStringToDate(text2), "dd/MM/yyyy");
					DateTime toDate = ConvertHelpers.ConvertToDateTime(HTKKFormater.FormatStringToDate(text), "dd/MM/yyyy");
					DateTime fromDate = ConvertHelpers.ConvertToDateTime(HTKKFormater.FormatStringToDate(tuNgay), "dd/MM/yyyy");
					DateTime dateTime2 = ConvertHelpers.ConvertToDateTime(HTKKFormater.FormatStringToDate(denNgay), "dd/MM/yyyy");
					num = DateUtils.GetSoNgay(fromDate, toDate);
					num2 = DateUtils.GetSoNgay(dateTime, dateTime2);
					if (num != 0)
					{
						num /= Math.Abs(num);
					}
					if (num2 != 0)
					{
						num2 /= Math.Abs(num2);
					}
					if (num == 0 && num2 == 0 && text4.Equals(nam))
					{
						result = false;
						break;
					}
					if (num * num2 > 0 && text4.Equals(nam))
					{
						result = true;
						break;
					}
					num = DateUtils.GetSoNgay(fromDate, toDate);
					num2 = DateUtils.GetSoNgay(fromDate, dateTime);
					if (num != 0)
					{
						num /= Math.Abs(num);
					}
					if (num2 != 0)
					{
						num2 /= Math.Abs(num2);
					}
					if (num * num2 <= 0 && text4.Equals(nam))
					{
						result = true;
						break;
					}
					num = DateUtils.GetSoNgay(dateTime2, toDate);
					num2 = DateUtils.GetSoNgay(dateTime2, dateTime);
					if (num != 0)
					{
						num /= Math.Abs(num);
					}
					if (num2 != 0)
					{
						num2 /= Math.Abs(num2);
					}
					if (num * num2 <= 0 && text4.Equals(nam))
					{
						result = true;
						break;
					}
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			result = false;
		}
		return result;
	}

	private bool CheckKyKKTrungNgay1(string sDataFile, HTKKPeriodInfo periodInfo)
	{
		bool result = false;
		try
		{
			string tuNgay = periodInfo.TuNgay;
			string denNgay = periodInfo.DenNgay;
			string text = "";
			string text2 = "";
			int num = 0;
			int num2 = 0;
			string path = HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst;
			sDataFile += ((!string.IsNullOrEmpty(periodInfo.maCQTQuanLy)) ? ("_CQTQL" + periodInfo.maCQTQuanLy) : "");
			string[] files = Directory.GetFiles(path, "*.xml");
			string text3 = "";
			for (int i = 0; i < files.Length; i++)
			{
				text3 = files[i].Substring(files[i].LastIndexOf('\\') + 1);
				if (text3.Length <= 31 || text3.IndexOf(sDataFile) != 0)
				{
					continue;
				}
				string[] array = text3.Replace(sDataFile + "_", "").Split(new char[1] { '_' });
				if (array != null && array.Length > 2)
				{
					text = array[^3];
					text2 = array[^2];
					DateTime dateTime = ConvertHelpers.ConvertToDateTime(HTKKFormater.FormatStringToDate(text2), "dd/MM/yyyy");
					DateTime toDate = ConvertHelpers.ConvertToDateTime(HTKKFormater.FormatStringToDate(text), "dd/MM/yyyy");
					DateTime fromDate = ConvertHelpers.ConvertToDateTime(HTKKFormater.FormatStringToDate(tuNgay), "dd/MM/yyyy");
					DateTime dateTime2 = ConvertHelpers.ConvertToDateTime(HTKKFormater.FormatStringToDate(denNgay), "dd/MM/yyyy");
					num = DateUtils.GetSoNgay(fromDate, toDate);
					num2 = DateUtils.GetSoNgay(dateTime, dateTime2);
					if (num != 0)
					{
						num /= Math.Abs(num);
					}
					if (num2 != 0)
					{
						num2 /= Math.Abs(num2);
					}
					if (num == 0 && num2 == 0)
					{
						result = false;
						break;
					}
					if (num * num2 > 0)
					{
						result = true;
						break;
					}
					num = DateUtils.GetSoNgay(fromDate, toDate);
					num2 = DateUtils.GetSoNgay(fromDate, dateTime);
					if (num != 0)
					{
						num /= Math.Abs(num);
					}
					if (num2 != 0)
					{
						num2 /= Math.Abs(num2);
					}
					if (num * num2 <= 0)
					{
						result = true;
						break;
					}
					num = DateUtils.GetSoNgay(dateTime2, toDate);
					num2 = DateUtils.GetSoNgay(dateTime2, dateTime);
					if (num != 0)
					{
						num /= Math.Abs(num);
					}
					if (num2 != 0)
					{
						num2 /= Math.Abs(num2);
					}
					if (num * num2 <= 0)
					{
						result = true;
						break;
					}
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			result = false;
		}
		return result;
	}

	private bool CheckPeriod(HTKKPeriodInfo periodInfo)
	{
		//IL_0133: Unknown result type (might be due to invalid IL or missing references)
		//IL_01ea: Unknown result type (might be due to invalid IL or missing references)
		//IL_20d1: Unknown result type (might be due to invalid IL or missing references)
		//IL_02fa: Unknown result type (might be due to invalid IL or missing references)
		//IL_0c9f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0ab9: Unknown result type (might be due to invalid IL or missing references)
		//IL_0412: Unknown result type (might be due to invalid IL or missing references)
		//IL_0352: Unknown result type (might be due to invalid IL or missing references)
		//IL_0bcd: Unknown result type (might be due to invalid IL or missing references)
		//IL_0b11: Unknown result type (might be due to invalid IL or missing references)
		//IL_046a: Unknown result type (might be due to invalid IL or missing references)
		//IL_22eb: Unknown result type (might be due to invalid IL or missing references)
		//IL_0d22: Unknown result type (might be due to invalid IL or missing references)
		//IL_0c25: Unknown result type (might be due to invalid IL or missing references)
		//IL_051b: Unknown result type (might be due to invalid IL or missing references)
		//IL_2346: Unknown result type (might be due to invalid IL or missing references)
		//IL_091d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0578: Unknown result type (might be due to invalid IL or missing references)
		//IL_2217: Unknown result type (might be due to invalid IL or missing references)
		//IL_05c2: Unknown result type (might be due to invalid IL or missing references)
		//IL_26f7: Unknown result type (might be due to invalid IL or missing references)
		//IL_2493: Unknown result type (might be due to invalid IL or missing references)
		//IL_0f5e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0da0: Unknown result type (might be due to invalid IL or missing references)
		//IL_070c: Unknown result type (might be due to invalid IL or missing references)
		//IL_27d5: Unknown result type (might be due to invalid IL or missing references)
		//IL_24ee: Unknown result type (might be due to invalid IL or missing references)
		//IL_21b2: Unknown result type (might be due to invalid IL or missing references)
		//IL_0ded: Unknown result type (might be due to invalid IL or missing references)
		//IL_0762: Unknown result type (might be due to invalid IL or missing references)
		//IL_064f: Unknown result type (might be due to invalid IL or missing references)
		//IL_28b3: Unknown result type (might be due to invalid IL or missing references)
		//IL_23bf: Unknown result type (might be due to invalid IL or missing references)
		//IL_09b5: Unknown result type (might be due to invalid IL or missing references)
		//IL_081a: Unknown result type (might be due to invalid IL or missing references)
		//IL_2991: Unknown result type (might be due to invalid IL or missing references)
		//IL_0e6b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0867: Unknown result type (might be due to invalid IL or missing references)
		//IL_2af3: Unknown result type (might be due to invalid IL or missing references)
		//IL_2a6f: Unknown result type (might be due to invalid IL or missing references)
		//IL_2567: Unknown result type (might be due to invalid IL or missing references)
		//IL_2413: Unknown result type (might be due to invalid IL or missing references)
		//IL_1022: Unknown result type (might be due to invalid IL or missing references)
		//IL_25bb: Unknown result type (might be due to invalid IL or missing references)
		//IL_10b2: Unknown result type (might be due to invalid IL or missing references)
		//IL_1208: Unknown result type (might be due to invalid IL or missing references)
		//IL_193b: Unknown result type (might be due to invalid IL or missing references)
		//IL_137b: Unknown result type (might be due to invalid IL or missing references)
		//IL_16ed: Unknown result type (might be due to invalid IL or missing references)
		//IL_1d6f: Unknown result type (might be due to invalid IL or missing references)
		//IL_19a8: Unknown result type (might be due to invalid IL or missing references)
		//IL_1743: Unknown result type (might be due to invalid IL or missing references)
		//IL_1db5: Unknown result type (might be due to invalid IL or missing references)
		//IL_180c: Unknown result type (might be due to invalid IL or missing references)
		//IL_13e8: Unknown result type (might be due to invalid IL or missing references)
		//IL_1af7: Unknown result type (might be due to invalid IL or missing references)
		//IL_1b46: Unknown result type (might be due to invalid IL or missing references)
		//IL_1537: Unknown result type (might be due to invalid IL or missing references)
		//IL_1586: Unknown result type (might be due to invalid IL or missing references)
		//IL_1bf1: Unknown result type (might be due to invalid IL or missing references)
		//IL_1631: Unknown result type (might be due to invalid IL or missing references)
		//IL_2009: Unknown result type (might be due to invalid IL or missing references)
		bool success = true;
		string ngay = periodInfo.Ngay;
		string text = periodInfo.Thang;
		string quy = periodInfo.Quy;
		string nam = periodInfo.Nam;
		string kieuKy = periodInfo.KieuKy;
		string text2 = "";
		string tieuMucHachToanVal = periodInfo.TieuMucHachToanVal;
		try
		{
			DateTime today = DateTime.Today;
			if (!CheckTuNgayDenNgay())
			{
				return false;
			}
			if (!CheckTuThangDenThang())
			{
				return false;
			}
			if (!CheckNgayThangNam())
			{
				return false;
			}
			if (kieuKy.Equals("D"))
			{
				text2 = ngay + "/" + text + "/" + nam;
				if (DateTime.Compare(new DateTime(ConvertHelpers.ConvertToInt(nam), ConvertHelpers.ConvertToInt(text), ConvertHelpers.ConvertToInt(ngay)), today) > 0 && !(menuID == "R5"))
				{
					frmMessage frmMessage = new frmMessage("0223", "3", "3", "Thông báo");
					((Form)frmMessage).ShowDialog();
					success = false;
					return success;
				}
			}
			if (kieuKy.Equals("Y") && menuID == "Q4")
			{
				text2 = "01/01/" + nam;
				string kieuKyHoanThue = periodInfo.KieuKyHoanThue;
				if (kieuKyHoanThue == "5" && ConvertHelpers.ConvertToInt(((Control)mskTuThang).Text) > ConvertHelpers.ConvertToInt(((Control)mskDenThang).Text))
				{
					frmMessage frmMessage = new frmMessage("0658", "3", "3", "Thông báo");
					((Form)frmMessage).ShowDialog();
					((Control)mskTuThang).Focus();
					success = false;
					return success;
				}
			}
			else if (kieuKy.Equals("M"))
			{
				text2 = "01/" + text + "/" + nam;
				if (menuID == "99" && tieuMucHachToanVal == "0010")
				{
					if (periodInfo.KieuKyHoanThue == "1")
					{
						DateTime dateTime = DateTime.ParseExact(periodInfo.TuThang, "MM/yyyy", null);
						DateTime dateTime2 = DateTime.ParseExact(periodInfo.DenThang, "MM/yyyy", null);
						if (dateTime2.Month - dateTime.Month + 1 + (dateTime2.Year - dateTime.Year) * 12 < 12)
						{
							frmMessage frmMessage2 = new frmMessage("0657", "3", "1", "Thông báo");
							((Form)frmMessage2).ShowDialog();
							((Control)mskDenThang).Focus();
							return false;
						}
						if (dateTime < DateTime.ParseExact("07/2025", "MM/yyyy", null))
						{
							frmMessage frmMessage2 = new frmMessage("0656", "3", "1", "Thông báo");
							((Form)frmMessage2).ShowDialog();
							((Control)mskDenThang).Focus();
							return false;
						}
					}
					else if (periodInfo.KieuKyHoanThue == "4")
					{
						DateTime dateTime = DateTime.ParseExact("0" + periodInfo.TuThang, "MM/yyyy", null);
						DateTime dateTime2 = DateTime.ParseExact(periodInfo.DenThang, "MM/yyyy", null);
						if (dateTime2.Month - (dateTime.Month - 1) * 3 + 1 + (dateTime2.Year - dateTime.Year) * 12 < 12)
						{
							frmMessage frmMessage2 = new frmMessage("0657", "3", "1", "Thông báo");
							((Form)frmMessage2).ShowDialog();
							((Control)mskDenThang).Focus();
							return false;
						}
						if (dateTime < DateTime.ParseExact("03/2025", "MM/yyyy", null))
						{
							frmMessage frmMessage2 = new frmMessage("0656", "3", "1", "Thông báo");
							((Form)frmMessage2).ShowDialog();
							((Control)mskDenThang).Focus();
							return false;
						}
					}
				}
				if (menuID == "99" || menuID == "H6")
				{
					string kieuKyHoanThue = periodInfo.KieuKyHoanThue;
					if (kieuKyHoanThue == "1")
					{
						DateTime dateTime3 = HTKKFormater.ConvertToDateTime("01/" + ((Control)mskTuThang).Text, out success);
						if (!success)
						{
							frmMessage frmMessage2 = new frmMessage("0018", "3", "1", "Thông báo");
							((Form)frmMessage2).ShowDialog();
							((Control)mskTuThang).Focus();
							return success;
						}
						DateTime dateTime4 = HTKKFormater.ConvertToDateTime("01/" + ((Control)mskDenThang).Text, out success);
						if (!success)
						{
							frmMessage frmMessage2 = new frmMessage("0018", "3", "1", "Thông báo");
							((Form)frmMessage2).ShowDialog();
							((Control)mskDenThang).Focus();
							return success;
						}
						if (dateTime3 > dateTime4)
						{
							frmMessage frmMessage2 = new frmMessage("0097", "3", "1", "Thông báo");
							((Form)frmMessage2).ShowDialog();
							((Control)mskDenThang).Focus();
							success = false;
							return success;
						}
						if (menuID == "H6")
						{
							DateTime t = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1);
							if (DateTime.Compare(t, dateTime4) < 0)
							{
								frmMessage frmMessage2 = new frmMessage("0247", "3", "1", "Thông báo");
								((Form)frmMessage2).ShowDialog();
								((Control)mskDenNgay).Focus();
								success = false;
								return success;
							}
						}
					}
					else if (kieuKyHoanThue == "4")
					{
						if (ConvertHelpers.ConvertToInt(((Control)mskTuThang).Text.Split(new char[1] { '/' })[0]) < 1 || ConvertHelpers.ConvertToInt(((Control)mskTuThang).Text.Split(new char[1] { '/' })[0]) > 4)
						{
							frmMessage frmMessage = new frmMessage("0018", "3", "3", "Thông báo");
							((Form)frmMessage).ShowDialog();
							((Control)mskTuThang).Focus();
							success = false;
							return success;
						}
						if (((Control)mskTuThang).Text.Length < 6)
						{
							frmMessage frmMessage = new frmMessage("0018", "3", "3", "Thông báo");
							((Form)frmMessage).ShowDialog();
							((Control)mskTuThang).Focus();
							success = false;
							return success;
						}
						DateTime ngayCuoiQuy = GeneralFunctions.GetNgayCuoiQuy(ConvertHelpers.ConvertToInt(((Control)mskTuThang).Text.Split(new char[1] { '/' })[0]), ConvertHelpers.ConvertToInt(((Control)mskTuThang).Text.Split(new char[1] { '/' })[1]), 1, 1);
						DateTime dateTime4 = HTKKFormater.ConvertToDateTime("01/" + ((Control)mskDenThang).Text, out success);
						if (!success)
						{
							frmMessage frmMessage2 = new frmMessage("0018", "3", "1", "Thông báo");
							((Form)frmMessage2).ShowDialog();
							((Control)mskDenThang).Focus();
							return success;
						}
						if (DateTime.Compare(dateTime4, ngayCuoiQuy) < 0)
						{
							frmMessage frmMessage = new frmMessage("0589", "3", "3", "Thông báo");
							((Form)frmMessage).ShowDialog();
							((Control)mskTuThang).Focus();
							success = false;
							return success;
						}
					}
				}
				else
				{
					if (menuID == "A4")
					{
						text = ((Control)mskDenThang).Text.Split(new char[1] { '/' })[0];
					}
					if (ConvertHelpers.ConvertToInt(nam) > today.Year)
					{
						if (!(menuID == "R5"))
						{
							frmMessage frmMessage = new frmMessage("0044", "3", "3", "Thông báo");
							((Form)frmMessage).ShowDialog();
							success = false;
							return success;
						}
					}
					else if (ConvertHelpers.ConvertToInt(nam) == today.Year && ConvertHelpers.ConvertToInt(text) > today.Month && menuID != "G7" && !(menuID == "R5"))
					{
						frmMessage frmMessage = new frmMessage("0044", "3", "3", "Thông báo");
						((Form)frmMessage).ShowDialog();
						success = false;
						return success;
					}
				}
			}
			else if (kieuKy.Equals("Q"))
			{
				if (menuID == "99" && tieuMucHachToanVal == "0010")
				{
					if (periodInfo.KieuKyHoanThue == "2")
					{
						DateTime dateTime = DateTime.ParseExact("0" + periodInfo.TuThang, "MM/yyyy", null);
						DateTime dateTime2 = DateTime.ParseExact("0" + periodInfo.DenThang, "MM/yyyy", null);
						if (dateTime2.Month - dateTime.Month + 1 + (dateTime2.Year - dateTime.Year) * 4 < 4)
						{
							frmMessage frmMessage2 = new frmMessage("0657", "3", "1", "Thông báo");
							((Form)frmMessage2).ShowDialog();
							((Control)mskDenThang).Focus();
							return false;
						}
						if (dateTime < DateTime.ParseExact("03/2025", "MM/yyyy", null))
						{
							frmMessage frmMessage2 = new frmMessage("0656", "3", "1", "Thông báo");
							((Form)frmMessage2).ShowDialog();
							((Control)mskDenThang).Focus();
							return false;
						}
					}
					else if (periodInfo.KieuKyHoanThue == "3")
					{
						DateTime dateTime = DateTime.ParseExact(periodInfo.TuThang, "MM/yyyy", null);
						DateTime dateTime2 = DateTime.ParseExact("0" + periodInfo.DenThang, "MM/yyyy", null);
						if (dateTime2.Month * 3 - dateTime.Month + (dateTime2.Year - dateTime.Year) * 12 < 12)
						{
							frmMessage frmMessage2 = new frmMessage("0657", "3", "1", "Thông báo");
							((Form)frmMessage2).ShowDialog();
							((Control)mskDenThang).Focus();
							return false;
						}
						if (dateTime < DateTime.ParseExact("07/2025", "MM/yyyy", null))
						{
							frmMessage frmMessage2 = new frmMessage("0656", "3", "1", "Thông báo");
							((Form)frmMessage2).ShowDialog();
							((Control)mskDenThang).Focus();
							return false;
						}
					}
				}
				if (menuID == "30")
				{
					text2 = periodInfo.TuNgay;
					if (ConvertHelpers.ConvertToDateTime(text2, "dd/MM/yyyy") > DateTime.Now)
					{
						frmMessage frmMessage = new frmMessage("0401", "3", "1", "Thông báo");
						((Form)frmMessage).ShowDialog();
						return false;
					}
				}
				else
				{
					text2 = GeneralFunctions.GetNgayDauQuy(ConvertHelpers.ConvertToInt(quy), ConvertHelpers.ConvertToInt(nam), 1, 1).ToString("dd/MM/yyyy");
				}
				if (menuID.Equals("13") && ConvertHelpers.ConvertToInt(nam) >= 2017)
				{
					frmMessage frmMessage = new frmMessage("0487", "3", "1", "Thông báo");
					((Form)frmMessage).ShowDialog();
					return false;
				}
				if (menuID.Equals("65"))
				{
					if ((ConvertHelpers.ConvertToInt(nam) == 2014 && ConvertHelpers.ConvertToInt(quy) < 3) || ConvertHelpers.ConvertToInt(nam) < 2014)
					{
						frmMessage frmMessage = new frmMessage("0316", "3", "1", "Thông báo");
						((Form)frmMessage).ShowDialog();
						return false;
					}
					if (ConvertHelpers.ConvertToInt(nam) > DateTime.Today.Year)
					{
						frmMessage frmMessage = new frmMessage("0188", "3", "1", "Thông báo");
						((Form)frmMessage).ShowDialog();
						return false;
					}
					if (ConvertHelpers.ConvertToInt(nam) == DateTime.Today.Year)
					{
						DateTime ngayDauQuy = GeneralFunctions.GetNgayDauQuy(ConvertHelpers.ConvertToInt(quy), ConvertHelpers.ConvertToInt(nam), 1, 1);
						if (DateTime.Compare(ngayDauQuy, DateTime.Today) > 0)
						{
							frmMessage frmMessage = new frmMessage("0188", "3", "1", "Thông báo");
							((Form)frmMessage).ShowDialog();
							return false;
						}
					}
				}
				else if (!menuID.Equals("A1") && !"X1".Equals(WinSettings.TTPeriodInfo.menuID) && !menuID.Equals("A4") && !menuID.Equals("99") && !menuID.Equals("H6") && !menuID.Equals("A3") && !menuID.Equals("46") && !menuID.Equals("30") && DateTime.Compare(GeneralFunctions.GetNgayDauQuy(ConvertHelpers.ConvertToInt(quy), ConvertHelpers.ConvertToInt(nam), 1, 1), today) > 0)
				{
					frmMessage frmMessage = new frmMessage("0603", "3", "3", "Thông báo");
					((Form)frmMessage).ShowDialog();
					success = false;
					return success;
				}
				if ((menuID.Equals("01") || menuID.Equals("02") || menuID.Equals("04") || menuID.Equals("71") || menuID.Equals("I5")) && ((ConvertHelpers.ConvertToInt(nam) == 2013 && ConvertHelpers.ConvertToInt(quy) < 3) || ConvertHelpers.ConvertToInt(nam) < 2013))
				{
					frmMessage frmMessage = new frmMessage("0272", "3", "3", "Thông báo");
					((Form)frmMessage).ShowDialog();
					return false;
				}
				if ((menuID.Equals("11") || menuID.Equals("12")) && ((ConvertHelpers.ConvertToInt(quy) >= 4 && ConvertHelpers.ConvertToInt(nam) == 2014) || ConvertHelpers.ConvertToInt(nam) > 2014))
				{
					frmMessage frmMessage = new frmMessage("0341", "3", "3", "Thông báo");
					((Form)frmMessage).ShowDialog();
					((Control)cbbQuy).Focus();
					success = false;
					return success;
				}
				if (((Control)mskTuThang).Visible && ((Control)mskDenThang).Visible && !chkTuQuyDenQuy.Checked && !(menuID == "99") && !(menuID == "H6"))
				{
					DateTime ngayDauQuy2 = GeneralFunctions.GetNgayDauQuy(ConvertHelpers.ConvertToInt(quy), ConvertHelpers.ConvertToInt(nam), 1, 1);
					DateTime ngayCuoiQuy2 = GeneralFunctions.GetNgayCuoiQuy(ConvertHelpers.ConvertToInt(quy), ConvertHelpers.ConvertToInt(nam), 1, 1);
					DateTime t2 = ConvertHelpers.ConvertToDateTime("01/" + ((Control)mskTuThang).Text, "dd/MM/yyyy");
					DateTime t3 = ConvertHelpers.ConvertToDateTime("01/" + ((Control)mskDenThang).Text, "dd/MM/yyyy").AddMonths(1).AddDays(-1.0);
					if (DateTime.Compare(ngayDauQuy2, t2) > 0 || DateTime.Compare(ngayCuoiQuy2, t3) < 0)
					{
						frmMessage frmMessage = new frmMessage("0365", "3", "3", "Thông báo");
						((Form)frmMessage).ShowDialog();
						((Control)mskTuThang).Focus();
						success = false;
						return success;
					}
				}
				if ((menuID == "99" || menuID == "H6") && ((Control)mskTuThang).Visible && ((Control)mskDenThang).Visible)
				{
					string kieuKyHoanThue = periodInfo.KieuKyHoanThue;
					if (kieuKyHoanThue == "2")
					{
						if (ConvertHelpers.ConvertToInt(((Control)mskTuThang).Text.Split(new char[1] { '/' })[0]) < 1 || ConvertHelpers.ConvertToInt(((Control)mskTuThang).Text.Split(new char[1] { '/' })[0]) > 4 || ConvertHelpers.ConvertToInt(((Control)mskDenThang).Text.Split(new char[1] { '/' })[0]) < 1 || ConvertHelpers.ConvertToInt(((Control)mskDenThang).Text.Split(new char[1] { '/' })[0]) > 4)
						{
							frmMessage frmMessage = new frmMessage("0018", "3", "3", "Thông báo");
							((Form)frmMessage).ShowDialog();
							((Control)mskTuThang).Focus();
							success = false;
							return success;
						}
						if (((Control)mskTuThang).Text.Length < 6 || ((Control)mskDenThang).Text.Length < 6)
						{
							frmMessage frmMessage = new frmMessage("0018", "3", "3", "Thông báo");
							((Form)frmMessage).ShowDialog();
							((Control)mskTuThang).Focus();
							success = false;
							return success;
						}
						DateTime ngayDauQuy3 = GeneralFunctions.GetNgayDauQuy(ConvertHelpers.ConvertToInt(((Control)mskTuThang).Text.Split(new char[1] { '/' })[0]), ConvertHelpers.ConvertToInt(((Control)mskTuThang).Text.Split(new char[1] { '/' })[1]), 1, 1);
						DateTime ngayDauQuy4 = GeneralFunctions.GetNgayDauQuy(ConvertHelpers.ConvertToInt(((Control)mskDenThang).Text.Split(new char[1] { '/' })[0]), ConvertHelpers.ConvertToInt(((Control)mskDenThang).Text.Split(new char[1] { '/' })[1]), 1, 1);
						if (ConvertHelpers.ConvertToInt(((Control)mskTuThang).Text.Split(new char[1] { '/' })[0]) > 4 || ConvertHelpers.ConvertToInt(((Control)mskDenThang).Text.Split(new char[1] { '/' })[0]) > 4)
						{
							frmMessage frmMessage = new frmMessage("0018", "3", "3", "Thông báo");
							((Form)frmMessage).ShowDialog();
							((Control)mskTuThang).Focus();
							success = false;
							return success;
						}
						if (DateTime.Compare(ngayDauQuy4, ngayDauQuy3) < 0)
						{
							frmMessage frmMessage = new frmMessage("0463", "3", "3", "Thông báo");
							((Form)frmMessage).ShowDialog();
							((Control)mskTuThang).Focus();
							success = false;
							return success;
						}
						int iQuy = ((DateTime.Now.Month < 3) ? 1 : ((DateTime.Now.Month < 7) ? 2 : ((DateTime.Now.Month < 10) ? 3 : 4)));
						DateTime ngayDauQuy5 = GeneralFunctions.GetNgayDauQuy(iQuy, DateTime.Now.Year, 1, 1);
						if (DateTime.Compare(ngayDauQuy5, ngayDauQuy4) < 0)
						{
							frmMessage frmMessage = new frmMessage("0464", "3", "3", "Thông báo");
							((Form)frmMessage).ShowDialog();
							((Control)mskDenThang).Focus();
							success = false;
							return success;
						}
					}
					else if (kieuKyHoanThue == "3")
					{
						if (ConvertHelpers.ConvertToInt(((Control)mskDenThang).Text.Split(new char[1] { '/' })[0]) < 1 || ConvertHelpers.ConvertToInt(((Control)mskDenThang).Text.Split(new char[1] { '/' })[0]) > 4)
						{
							frmMessage frmMessage = new frmMessage("0018", "3", "3", "Thông báo");
							((Form)frmMessage).ShowDialog();
							((Control)mskTuThang).Focus();
							success = false;
							return success;
						}
						if (((Control)mskDenThang).Text.Length < 6)
						{
							frmMessage frmMessage = new frmMessage("0018", "3", "3", "Thông báo");
							((Form)frmMessage).ShowDialog();
							((Control)mskTuThang).Focus();
							success = false;
							return success;
						}
						DateTime ngayDauQuy4 = GeneralFunctions.GetNgayCuoiQuy(ConvertHelpers.ConvertToInt(((Control)mskDenThang).Text.Split(new char[1] { '/' })[0]), ConvertHelpers.ConvertToInt(((Control)mskDenThang).Text.Split(new char[1] { '/' })[1]), 1, 1);
						DateTime t4 = ConvertHelpers.ConvertToDateTime("01/" + ((Control)mskTuThang).Text, "dd/MM/yyyy");
						if (DateTime.Compare(ngayDauQuy4, t4) < 0)
						{
							frmMessage frmMessage = new frmMessage("0588", "3", "3", "Thông báo");
							((Form)frmMessage).ShowDialog();
							((Control)mskTuThang).Focus();
							success = false;
							return success;
						}
					}
				}
				if (((Control)mskTuThang).Visible && ((Control)mskDenThang).Visible && chkTuQuyDenQuy.Checked)
				{
					if (ConvertHelpers.ConvertToInt(((Control)mskTuThang).Text.Split(new char[1] { '/' })[0]) < 1 || ConvertHelpers.ConvertToInt(((Control)mskTuThang).Text.Split(new char[1] { '/' })[0]) > 4 || ConvertHelpers.ConvertToInt(((Control)mskDenThang).Text.Split(new char[1] { '/' })[0]) < 1 || ConvertHelpers.ConvertToInt(((Control)mskDenThang).Text.Split(new char[1] { '/' })[0]) > 4)
					{
						frmMessage frmMessage = new frmMessage("0018", "3", "3", "Thông báo");
						((Form)frmMessage).ShowDialog();
						((Control)mskTuThang).Focus();
						success = false;
						return success;
					}
					if (((Control)mskTuThang).Text.Length < 6 || ((Control)mskDenThang).Text.Length < 6)
					{
						frmMessage frmMessage = new frmMessage("0018", "3", "3", "Thông báo");
						((Form)frmMessage).ShowDialog();
						((Control)mskTuThang).Focus();
						success = false;
						return success;
					}
					DateTime ngayDauQuy3 = GeneralFunctions.GetNgayDauQuy(ConvertHelpers.ConvertToInt(((Control)mskTuThang).Text.Split(new char[1] { '/' })[0]), ConvertHelpers.ConvertToInt(((Control)mskTuThang).Text.Split(new char[1] { '/' })[1]), 1, 1);
					DateTime ngayDauQuy4 = GeneralFunctions.GetNgayDauQuy(ConvertHelpers.ConvertToInt(((Control)mskDenThang).Text.Split(new char[1] { '/' })[0]), ConvertHelpers.ConvertToInt(((Control)mskDenThang).Text.Split(new char[1] { '/' })[1]), 1, 1);
					if (ConvertHelpers.ConvertToInt(((Control)mskTuThang).Text.Split(new char[1] { '/' })[0]) > 4 || ConvertHelpers.ConvertToInt(((Control)mskDenThang).Text.Split(new char[1] { '/' })[0]) > 4)
					{
						frmMessage frmMessage = new frmMessage("0018", "3", "3", "Thông báo");
						((Form)frmMessage).ShowDialog();
						((Control)mskTuThang).Focus();
						success = false;
						return success;
					}
					if (DateTime.Compare(ngayDauQuy4, ngayDauQuy3) < 0)
					{
						frmMessage frmMessage = new frmMessage("0463", "3", "3", "Thông báo");
						((Form)frmMessage).ShowDialog();
						((Control)mskTuThang).Focus();
						success = false;
						return success;
					}
					int iQuy = ((DateTime.Now.Month < 3) ? 1 : ((DateTime.Now.Month < 7) ? 2 : ((DateTime.Now.Month < 10) ? 3 : 4)));
					DateTime ngayDauQuy5 = GeneralFunctions.GetNgayDauQuy(iQuy, DateTime.Now.Year, 1, 1);
					if (DateTime.Compare(ngayDauQuy5, ngayDauQuy4) < 0)
					{
						frmMessage frmMessage = new frmMessage("0464", "3", "3", "Thông báo");
						((Form)frmMessage).ShowDialog();
						((Control)mskDenThang).Focus();
						success = false;
						return success;
					}
				}
				if (((Control)mskTuNgay).Visible && ((Control)mskDenNgay).Visible && ("A1,H1,A3,H3,A8,H8,A9,H9,B5".Contains(menuID) || "X1".Equals(WinSettings.TTPeriodInfo.menuID) || "X2".Equals(WinSettings.TTPeriodInfo.menuID)))
				{
					int iNgayBdNamTc = ConvertHelpers.ConvertToInt(WinSettings.HTKKHeaderInfo.NgayBatDauNamTaiChinh.Split(new char[1] { '/' })[0]);
					int iThangBdNamTc = ConvertHelpers.ConvertToInt(WinSettings.HTKKHeaderInfo.NgayBatDauNamTaiChinh.Split(new char[1] { '/' })[1]);
					DateTime ngayDauQuy2 = GeneralFunctions.GetNgayDauQuy(ConvertHelpers.ConvertToInt(quy), ConvertHelpers.ConvertToInt(nam), iNgayBdNamTc, iThangBdNamTc);
					DateTime ngayCuoiQuy2 = GeneralFunctions.GetNgayCuoiQuy(ConvertHelpers.ConvertToInt(quy), ConvertHelpers.ConvertToInt(nam), iNgayBdNamTc, iThangBdNamTc);
					DateTime t2 = ConvertHelpers.ConvertToDateTime(((Control)mskTuNgay).Text, "dd/MM/yyyy");
					DateTime t3 = ConvertHelpers.ConvertToDateTime(((Control)mskDenNgay).Text, "dd/MM/yyyy");
					if (DateTime.Compare(ngayDauQuy2, t2) > 0 || DateTime.Compare(ngayCuoiQuy2, t3) < 0)
					{
						frmMessage frmMessage = new frmMessage("0453", "3", "3", "Thông báo");
						((Form)frmMessage).ShowDialog();
						success = false;
						return success;
					}
					if (DateTime.Compare(DateTime.Now, t2) < 0)
					{
						frmMessage frmMessage = new frmMessage("0401", "3", "3", "Thông báo");
						((Form)frmMessage).ShowDialog();
						success = false;
						return success;
					}
					XmlNode val = (WinSettings.xmlMenuNode.HasChildNodes ? WinSettings.xmlMenuNode.ChildNodes[0] : null);
					XmlNode val2 = ((val != null && val.HasChildNodes) ? val.ChildNodes[0] : null);
					if (val2 != null)
					{
						string text3 = ((val2.Attributes["DataFile"] != null) ? ((XmlNode)val2.Attributes["DataFile"]).Value : "");
						if (!HTKKFormater.isNullString(text3))
						{
							string path = HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst;
							DirectoryInfo directoryInfo = new DirectoryInfo(path);
							if (directoryInfo.Exists)
							{
								FileInfo[] files = directoryInfo.GetFiles();
								if (files.Length > 0)
								{
									FileInfo[] array = files;
									foreach (FileInfo fileInfo in array)
									{
										if (!fileInfo.Name.Contains(text3))
										{
											continue;
										}
										string text4 = fileInfo.Name.Replace(text3 + "_", "").Replace("_" + WinSettings.TTPeriodInfo.VersionXML.Replace(".", "") + ".xml", "");
										if (text4.Count((char f) => f == '_') >= 2)
										{
											string[] array2 = text4.Split(new char[1] { '_' });
											DateTime dateTime5 = ConvertHelpers.ConvertToDateTime(HTKKFormater.FormatStringToDate(array2[1]), "dd/MM/yyyy");
											DateTime dateTime6 = ConvertHelpers.ConvertToDateTime(HTKKFormater.FormatStringToDate(array2[2]), "dd/MM/yyyy");
											if ((DateTime.Compare(dateTime5, t2) > 0 && DateTime.Compare(t3, dateTime5) > 0) || (DateTime.Compare(dateTime6, t2) > 0 && DateTime.Compare(t3, dateTime6) > 0))
											{
												frmMessage frmMessage = new frmMessage("0344", "3", "3", "Thông báo");
												((Form)frmMessage).ShowDialog();
												success = false;
												return success;
											}
										}
									}
								}
							}
						}
					}
				}
			}
			else if (kieuKy.Equals("Y"))
			{
				text2 = "31/12/" + nam;
				if (ConvertHelpers.ConvertToInt(nam) > today.Year && !menuID.Equals("46") && menuID != "63" && menuID != "K4")
				{
					frmMessage frmMessage = new frmMessage("0063", "3", "3", "Thông báo");
					((Form)frmMessage).ShowDialog();
					success = false;
					return success;
				}
				if (menuID == "63" || menuID == "L1")
				{
					bool flag = false;
					string text5 = "";
					if (periodInfo.dicVar.Count > 0)
					{
						flag = Convert.ToBoolean(periodInfo.dicVar["CoSoMoiThanhLap"]);
						text5 = (string)periodInfo.dicVar["NgayThanhLap"];
						if (flag && !nam.Equals(ConvertHelpers.ConvertToDateTime(text5, "dd/MM/yyyy").ToString("yyyy")))
						{
							frmMessage frmMessage = new frmMessage("0443", "3", "3", "Thông báo");
							((Form)frmMessage).ShowDialog();
							success = false;
							return success;
						}
					}
				}
				if (menuID == "63" && ConvertHelpers.ConvertToInt(periodInfo.Nam) < 2017)
				{
					frmMessage frmMessage = new frmMessage("0461", "3", "3", "Thông báo");
					((Form)frmMessage).ShowDialog();
					success = false;
					return success;
				}
			}
			else if (kieuKy.Equals("D_D"))
			{
				text2 = periodInfo.TuNgay;
			}
			else if (kieuKy.Equals("K"))
			{
				text2 = ((ConvertHelpers.ConvertToInt(quy) < 2) ? ("01/01/" + nam) : ("01/06/" + nam));
				if (menuID.Equals("13"))
				{
					if (ConvertHelpers.ConvertToInt(nam) < 2017)
					{
						frmMessage frmMessage = new frmMessage("0488", "3", "3", "Thông báo");
						((Form)frmMessage).ShowDialog();
						((Control)cbbQuy).Focus();
						success = false;
						return success;
					}
					if (ConvertHelpers.ConvertToInt(nam) > DateTime.Today.Year)
					{
						frmMessage frmMessage = new frmMessage("0188", "3", "1", "Thông báo");
						((Form)frmMessage).ShowDialog();
						return false;
					}
					if (ConvertHelpers.ConvertToInt(nam) == DateTime.Today.Year)
					{
						if (ConvertHelpers.ConvertToInt(quy) == 1 && 1 > DateTime.Today.Month)
						{
							frmMessage frmMessage = new frmMessage("0188", "3", "1", "Thông báo");
							((Form)frmMessage).ShowDialog();
							return false;
						}
						if (ConvertHelpers.ConvertToInt(quy) == 2 && 7 > DateTime.Today.Month)
						{
							frmMessage frmMessage = new frmMessage("0188", "3", "1", "Thông báo");
							((Form)frmMessage).ShowDialog();
							return false;
						}
					}
				}
				if (menuID.Equals("65"))
				{
					if (ConvertHelpers.ConvertToInt(nam) > 2014 || (ConvertHelpers.ConvertToInt(nam) == 2014 && ConvertHelpers.ConvertToInt(quy) >= 2))
					{
						frmMessage frmMessage = new frmMessage("0315", "3", "3", "Thông báo");
						((Form)frmMessage).ShowDialog();
						((Control)cbbQuy).Focus();
						success = false;
						return success;
					}
					if (ConvertHelpers.ConvertToInt(nam) > DateTime.Today.Year)
					{
						frmMessage frmMessage = new frmMessage("0188", "3", "1", "Thông báo");
						((Form)frmMessage).ShowDialog();
						return false;
					}
					if (ConvertHelpers.ConvertToInt(nam) == DateTime.Today.Year)
					{
						if (ConvertHelpers.ConvertToInt(quy) == 1 && 1 > DateTime.Today.Month)
						{
							frmMessage frmMessage = new frmMessage("0188", "3", "1", "Thông báo");
							((Form)frmMessage).ShowDialog();
							return false;
						}
						if (ConvertHelpers.ConvertToInt(quy) == 2 && 7 > DateTime.Today.Month)
						{
							frmMessage frmMessage = new frmMessage("0188", "3", "1", "Thông báo");
							((Form)frmMessage).ShowDialog();
							return false;
						}
					}
				}
			}
			else
			{
				success = true;
			}
			if (!CheckPeriodBctc(periodInfo, axfpsPhuLuc))
			{
				success = false;
			}
			if (!CheckHLToKhaiDKTT36(periodInfo))
			{
				success = false;
			}
			if (!CheckHLToKhaiDKTT61(periodInfo))
			{
				success = false;
			}
			if (!CheckPeriodTncn(periodInfo))
			{
				success = false;
			}
			if (!CheckHLToKhaiTT80(periodInfo))
			{
				success = false;
			}
			if (menuID == "N5")
			{
				string nam2 = periodInfo.Nam;
				string thang = periodInfo.Thang;
				string ngay2 = periodInfo.Ngay;
				DateTime dateTime7 = ConvertHelpers.ConvertToDateTime("01/01/2022", "dd/MM/yyyy");
				DateTime dateTime8 = ConvertHelpers.ConvertToDateTime("30/09/2022", "dd/MM/yyyy");
				DateTime dateTime9 = ConvertHelpers.ConvertToDateTime(ngay2 + "/" + thang + "/" + nam2, "dd/MM/yyyy");
				if (dateTime9 > dateTime8 || dateTime9 < dateTime7)
				{
					frmMessage frmMessage = new frmMessage("0622", "3", "3", "Thông báo");
					((Form)frmMessage).ShowDialog();
					success = false;
				}
			}
			if (menuID == "H5")
			{
				string nam2 = periodInfo.Nam;
				string thang = periodInfo.Thang;
				string ngay2 = periodInfo.Ngay;
				DateTime dateTime7 = ConvertHelpers.ConvertToDateTime("14/04/2023", "dd/MM/yyyy");
				DateTime dateTime8 = ConvertHelpers.ConvertToDateTime("30/09/2023", "dd/MM/yyyy");
				DateTime dateTime9 = ConvertHelpers.ConvertToDateTime(ngay2 + "/" + thang + "/" + nam2, "dd/MM/yyyy");
				if (dateTime9 > dateTime8 || dateTime9 < dateTime7)
				{
					frmMessage frmMessage = new frmMessage("0634", "3", "3", "Thông báo");
					((Form)frmMessage).ShowDialog();
					success = false;
				}
			}
			if (menuID == "O7")
			{
				string nam2 = periodInfo.Nam;
				string thang = periodInfo.Thang;
				string ngay2 = periodInfo.Ngay;
				DateTime dateTime7 = ConvertHelpers.ConvertToDateTime("17/06/2024", "dd/MM/yyyy");
				DateTime dateTime8 = ConvertHelpers.ConvertToDateTime("30/09/2024", "dd/MM/yyyy");
				DateTime dateTime9 = ConvertHelpers.ConvertToDateTime(ngay2 + "/" + thang + "/" + nam2, "dd/MM/yyyy");
				if (dateTime9 > dateTime8 || dateTime9 < dateTime7)
				{
					frmMessage frmMessage = new frmMessage("0646", "3", "3", "Thông báo");
					((Form)frmMessage).ShowDialog();
					success = false;
				}
			}
			if (menuID == "P5")
			{
				string nam2 = periodInfo.Nam;
				string thang = periodInfo.Thang;
				string ngay2 = periodInfo.Ngay;
				DateTime dateTime7 = ConvertHelpers.ConvertToDateTime("02/04/2025", "dd/MM/yyyy");
				DateTime dateTime8 = ConvertHelpers.ConvertToDateTime("30/05/2025", "dd/MM/yyyy");
				DateTime dateTime9 = ConvertHelpers.ConvertToDateTime(ngay2 + "/" + thang + "/" + nam2, "dd/MM/yyyy");
				if (dateTime9 > dateTime8 || dateTime9 < dateTime7)
				{
					frmMessage frmMessage = new frmMessage("0650", "3", "3", "Thông báo");
					((Form)frmMessage).ShowDialog();
					success = false;
				}
			}
			if (menuID == "U6")
			{
				string nam2 = periodInfo.Nam;
				string thang = periodInfo.Thang;
				string ngay2 = periodInfo.Ngay;
				DateTime dateTime7 = ConvertHelpers.ConvertToDateTime("27/06/2026", "dd/MM/yyyy");
				DateTime dateTime8 = ConvertHelpers.ConvertToDateTime("02/11/2026", "dd/MM/yyyy");
				DateTime dateTime9 = ConvertHelpers.ConvertToDateTime(ngay2 + "/" + thang + "/" + nam2, "dd/MM/yyyy");
				if (dateTime9 > dateTime8 || dateTime9 < dateTime7)
				{
					frmMessage frmMessage = new frmMessage("0687", "3", "3", "Thông báo");
					((Form)frmMessage).ShowDialog();
					success = false;
				}
			}
			if (!HTKKFormater.isNullString(CheckHieuLuc(text2)) && success)
			{
				string value = ((XmlNode)((XmlNode)WinSettings.XmlMessage).SelectNodes("/Message/Infor[@ID='0441']")[0].Attributes["Msg"]).Value;
				frmMessage frmMessage = new frmMessage("", value + " " + CheckHieuLuc(text2), "3", "3", "Thông báo");
				((Form)frmMessage).ShowDialog();
				success = false;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return success;
	}

	private bool CheckPeriodBctc(HTKKPeriodInfo periodInfo, FpSpread winFpSpread)
	{
		//IL_0130: Unknown result type (might be due to invalid IL or missing references)
		//IL_03e7: Unknown result type (might be due to invalid IL or missing references)
		//IL_0626: Unknown result type (might be due to invalid IL or missing references)
		//IL_0190: Unknown result type (might be due to invalid IL or missing references)
		//IL_067f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0461: Unknown result type (might be due to invalid IL or missing references)
		//IL_06d8: Unknown result type (might be due to invalid IL or missing references)
		//IL_01ea: Unknown result type (might be due to invalid IL or missing references)
		//IL_0554: Unknown result type (might be due to invalid IL or missing references)
		//IL_02c1: Unknown result type (might be due to invalid IL or missing references)
		//IL_05cd: Unknown result type (might be due to invalid IL or missing references)
		//IL_033a: Unknown result type (might be due to invalid IL or missing references)
		bool result = true;
		try
		{
			string nam = periodInfo.Nam;
			string tuNgay = periodInfo.TuNgay;
			string[] array = tuNgay.Split(new char[1] { '/' });
			string denNgay = periodInfo.DenNgay;
			string[] array2 = denNgay.Split(new char[1] { '/' });
			string text = periodInfo.menuID;
			string text2 = startDate;
			string ngayBatDauNamTaiChinh = WinSettings.HTKKHeaderInfo.NgayBatDauNamTaiChinh;
			string[] array3 = ngayBatDauNamTaiChinh.Split(new char[1] { '/' });
			int iNgayTc;
			int iThangTc;
			if (array3.Length == 2)
			{
				iNgayTc = ConvertHelpers.ConvertToInt(array3[0]);
				iThangTc = ConvertHelpers.ConvertToInt(array3[1]);
			}
			else
			{
				iThangTc = 1;
				iNgayTc = 1;
			}
			DateTime value = GeneralFunctions.NgayCuoiNamTaiChinh(ConvertHelpers.ConvertToInt(nam), iThangTc, iNgayTc);
			if ("78,58".Contains(text) || "U9".Equals(text))
			{
				if (ConvertHelpers.ConvertToDateTime(text2, "dd/MM/yyyy").CompareTo(value) > 0)
				{
					frmMessage frmMessage = new frmMessage("0431", "3", "3", "Thông báo");
					((Form)frmMessage).ShowDialog();
					result = false;
					return result;
				}
				if (ConvertHelpers.ConvertToInt(nam) < ConvertHelpers.ConvertToInt(array[2]) || ConvertHelpers.ConvertToInt(nam) > ConvertHelpers.ConvertToInt(array2[2]))
				{
					frmMessage frmMessage = new frmMessage("0396", "3", "3", "Thông báo");
					((Form)frmMessage).ShowDialog();
					result = false;
					return result;
				}
				if (!periodInfo.DoanhNghiepHoatDongLienTuc.Equals("LT") && !periodInfo.DoanhNghiepHoatDongLienTuc.Equals("KL"))
				{
					frmMessage frmMessage = new frmMessage("0397", "3", "3", "Thông báo");
					((Form)frmMessage).ShowDialog();
					result = false;
					return result;
				}
				int num = (ConvertHelpers.ConvertToDateTime(denNgay, "dd/MM/yyyy").Year - ConvertHelpers.ConvertToDateTime(tuNgay, "dd/MM/yyyy").Year) * 12 + ConvertHelpers.ConvertToDateTime(denNgay, "dd/MM/yyyy").Month - ConvertHelpers.ConvertToDateTime(tuNgay, "dd/MM/yyyy").Month + 1;
				if (num > 12 && num <= 15 && ConvertHelpers.ConvertToDateTime(tuNgay, "dd/MM/yyyy").Month == 10 && ConvertHelpers.ConvertToInt(nam) != ConvertHelpers.ConvertToDateTime(tuNgay, "dd/MM/yyyy").Year + 1)
				{
					frmMessage frmMessage = new frmMessage("0476", "3", "3", "Thông báo");
					((Form)frmMessage).ShowDialog();
					result = false;
					return result;
				}
				if (num <= 12 && ConvertHelpers.ConvertToInt(nam) != ConvertHelpers.ConvertToDateTime(tuNgay, "dd/MM/yyyy").Year && ConvertHelpers.ConvertToInt(nam) != ConvertHelpers.ConvertToDateTime(denNgay, "dd/MM/yyyy").Year)
				{
					frmMessage frmMessage = new frmMessage("0477", "3", "3", "Thông báo");
					((Form)frmMessage).ShowDialog();
					result = false;
					return result;
				}
			}
			if ("C1,C2,C3,C4,C5,X3".Contains(text) && ConvertHelpers.ConvertToDateTime(text2, "dd/MM/yyyy").CompareTo(ConvertHelpers.ConvertToDateTime(tuNgay, "dd/MM/yyyy")) > 0)
			{
				string value2 = ((XmlNode)((XmlNode)WinSettings.XmlMessage).SelectNodes("/Message/Infor[@ID='0441']")[0].Attributes["Msg"]).Value;
				frmMessage frmMessage = new frmMessage("", value2 + " " + text2, "3", "3", "Thông báo");
				((Form)frmMessage).ShowDialog();
				result = false;
				return result;
			}
			if ("53,54,B3,B4,A2,A5,A6,A7,C1,C2,C3,C4,C5,X3".Contains(text))
			{
				if (ConvertHelpers.ConvertToInt(nam) < ConvertHelpers.ConvertToInt(array[2]) || ConvertHelpers.ConvertToInt(nam) > ConvertHelpers.ConvertToInt(array2[2]))
				{
					frmMessage frmMessage = new frmMessage("0396", "3", "3", "Thông báo");
					((Form)frmMessage).ShowDialog();
					result = false;
					return result;
				}
				int num = (ConvertHelpers.ConvertToDateTime(denNgay, "dd/MM/yyyy").Year - ConvertHelpers.ConvertToDateTime(tuNgay, "dd/MM/yyyy").Year) * 12 + ConvertHelpers.ConvertToDateTime(denNgay, "dd/MM/yyyy").Month - ConvertHelpers.ConvertToDateTime(tuNgay, "dd/MM/yyyy").Month + 1;
				if (num > 12 && num <= 15 && ConvertHelpers.ConvertToDateTime(tuNgay, "dd/MM/yyyy").Month == 10 && ConvertHelpers.ConvertToInt(nam) != ConvertHelpers.ConvertToDateTime(tuNgay, "dd/MM/yyyy").Year && ConvertHelpers.ConvertToInt(nam) != ConvertHelpers.ConvertToDateTime(denNgay, "dd/MM/yyyy").Year)
				{
					frmMessage frmMessage = new frmMessage("0476", "3", "3", "Thông báo");
					((Form)frmMessage).ShowDialog();
					result = false;
					return result;
				}
				if (num <= 12 && ConvertHelpers.ConvertToInt(nam) != ConvertHelpers.ConvertToDateTime(tuNgay, "dd/MM/yyyy").Year && ConvertHelpers.ConvertToInt(nam) != ConvertHelpers.ConvertToDateTime(denNgay, "dd/MM/yyyy").Year)
				{
					frmMessage frmMessage = new frmMessage("0477", "3", "3", "Thông báo");
					((Form)frmMessage).ShowDialog();
					result = false;
					return result;
				}
			}
			if (text.Equals("69") && ConvertHelpers.ConvertToInt(nam) >= 2015)
			{
				frmMessage frmMessage = new frmMessage("0358", "3", "3", "Thông báo");
				((Form)frmMessage).ShowDialog();
				result = false;
				return result;
			}
			if (text.Equals("22") && ConvertHelpers.ConvertToInt(nam) >= 2016)
			{
				frmMessage frmMessage = new frmMessage("0494", "3", "3", "Thông báo");
				((Form)frmMessage).ShowDialog();
				result = false;
				return result;
			}
			if (text.Equals("19") && ConvertHelpers.ConvertToInt(nam) >= 2018)
			{
				frmMessage frmMessage = new frmMessage("0501", "3", "3", "Thông báo");
				((Form)frmMessage).ShowDialog();
				result = false;
				return result;
			}
			result = WinDoWorkEvent.CheckChonPlBctc(periodInfo, winFpSpread);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return result;
	}

	private bool CheckHLToKhaiDKTT36(HTKKPeriodInfo periodInfo)
	{
		//IL_01b0: Unknown result type (might be due to invalid IL or missing references)
		//IL_012b: Unknown result type (might be due to invalid IL or missing references)
		//IL_00ca: Unknown result type (might be due to invalid IL or missing references)
		//IL_0215: Unknown result type (might be due to invalid IL or missing references)
		//IL_030e: Unknown result type (might be due to invalid IL or missing references)
		//IL_02a3: Unknown result type (might be due to invalid IL or missing references)
		bool result = true;
		try
		{
			string kieuKy = periodInfo.KieuKy;
			string text = periodInfo.menuID;
			string nam = periodInfo.Nam;
			string thang = periodInfo.Thang;
			string ngay = periodInfo.Ngay;
			string quy = periodInfo.Quy;
			string obj = "01/01/2016";
			DateTime t = ConvertHelpers.ConvertToDateTime(obj, "dd/MM/yyyy");
			if (text.Equals("08") || text.Equals("40"))
			{
				if (kieuKy.Equals("M"))
				{
					if (DateTime.Compare(t, new DateTime(ConvertHelpers.ConvertToInt(nam), ConvertHelpers.ConvertToInt(thang), 1)) > 0)
					{
						frmMessage frmMessage = new frmMessage("0424", "3", "3", "Thông báo");
						((Form)frmMessage).ShowDialog();
						result = false;
						return result;
					}
				}
				else if (t.CompareTo(new DateTime(ConvertHelpers.ConvertToInt(nam), ConvertHelpers.ConvertToInt(thang), ConvertHelpers.ConvertToInt(ngay))) > 0)
				{
					frmMessage frmMessage = new frmMessage("0423", "3", "3", "Thông báo");
					((Form)frmMessage).ShowDialog();
					result = false;
					return result;
				}
			}
			if (text.Equals("21"))
			{
				t = ConvertHelpers.ConvertToDateTime("11/11/2016", "dd/MM/yyyy");
				if (t.CompareTo(new DateTime(ConvertHelpers.ConvertToInt(nam), ConvertHelpers.ConvertToInt(thang), ConvertHelpers.ConvertToInt(ngay))) > 0)
				{
					frmMessage frmMessage = new frmMessage("0462", "3", "3", "Thông báo");
					((Form)frmMessage).ShowDialog();
					result = false;
					return result;
				}
			}
			if (text.Equals("44") && t.CompareTo(new DateTime(ConvertHelpers.ConvertToInt(nam), 1, 1)) > 0)
			{
				frmMessage frmMessage = new frmMessage("0425", "3", "3", "Thông báo");
				((Form)frmMessage).ShowDialog();
				result = false;
				return result;
			}
			if (text.Equals("42") || text.Equals("46"))
			{
				if (kieuKy.Equals("Y"))
				{
					if (t.CompareTo(new DateTime(ConvertHelpers.ConvertToInt(nam), 1, 1)) > 0)
					{
						frmMessage frmMessage = new frmMessage("0425", "3", "3", "Thông báo");
						((Form)frmMessage).ShowDialog();
						result = false;
						return result;
					}
				}
				else if (ConvertHelpers.ConvertToDateTime("01/04/2016", "dd/MM/yyyy").CompareTo(new DateTime(ConvertHelpers.ConvertToInt(nam), ConvertHelpers.ConvertToInt(quy), 1)) > 0)
				{
					frmMessage frmMessage = new frmMessage("0426", "3", "3", "Thông báo");
					((Form)frmMessage).ShowDialog();
					result = false;
					return result;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return result;
	}

	private bool CheckHLToKhaiTT80(HTKKPeriodInfo periodInfo)
	{
		//IL_00f1: Unknown result type (might be due to invalid IL or missing references)
		//IL_047f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0162: Unknown result type (might be due to invalid IL or missing references)
		//IL_06b9: Unknown result type (might be due to invalid IL or missing references)
		//IL_04ed: Unknown result type (might be due to invalid IL or missing references)
		//IL_071b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0554: Unknown result type (might be due to invalid IL or missing references)
		//IL_02d3: Unknown result type (might be due to invalid IL or missing references)
		//IL_05c9: Unknown result type (might be due to invalid IL or missing references)
		//IL_062b: Unknown result type (might be due to invalid IL or missing references)
		//IL_025a: Unknown result type (might be due to invalid IL or missing references)
		//IL_01fd: Unknown result type (might be due to invalid IL or missing references)
		//IL_03e9: Unknown result type (might be due to invalid IL or missing references)
		//IL_0398: Unknown result type (might be due to invalid IL or missing references)
		bool result = true;
		try
		{
			string kieuKy = periodInfo.KieuKy;
			string text = periodInfo.menuID;
			string nam = periodInfo.Nam;
			string thang = periodInfo.Thang;
			string ngay = periodInfo.Ngay;
			string quy = periodInfo.Quy;
			string tuNgay = periodInfo.TuNgay;
			string denNgay = periodInfo.DenNgay;
			string obj = "01/01/2022";
			string obj2 = "01/01/2021";
			DateTime t = ConvertHelpers.ConvertToDateTime(obj, "dd/MM/yyyy");
			DateTime dateTime = ConvertHelpers.ConvertToDateTime(obj2, "dd/MM/yyyy");
			if ("~G5~G6~G7~G8~G9~L1~L2~L3~L4~L5~L6~L7~L8~L9~T8~T9~H2~H4~D1~D2~D3~D4~D5~D6~T2~T3~T4~T5~T6~T7~I1~I2~I3~I4~I7~I8~I9~J1~J5~J6~J7~J8~J9~M8~M9~K2~K3~K4~K5~K7~N2~N8~B8~B9~C7~C9~C6~C8~J3~D7~D8~D9~E1~E2~E3~E4~E5~E6~E7~E8~J2~J4~E9~R9~R6~U2~X9~Y1~Y6~Z1~Z3~Y5~Z4~Z8~Z5~W1~".Contains("~" + menuID + "~"))
			{
				if (kieuKy.Equals("M"))
				{
					if (DateTime.Compare(t, new DateTime(ConvertHelpers.ConvertToInt(nam), ConvertHelpers.ConvertToInt(thang), 1)) > 0)
					{
						frmMessage frmMessage = new frmMessage("0577", "3", "3", "Thông báo");
						((Form)frmMessage).ShowDialog();
						result = false;
						return result;
					}
				}
				else if (kieuKy.Equals("Q"))
				{
					if (t.CompareTo(new DateTime(ConvertHelpers.ConvertToInt(nam), ConvertHelpers.ConvertToInt(quy), 1)) > 0)
					{
						frmMessage frmMessage = new frmMessage("0579", "3", "3", "Thông báo");
						((Form)frmMessage).ShowDialog();
						result = false;
						return result;
					}
				}
				else if (kieuKy.Equals("Y"))
				{
					if (menuID == "G5" || menuID == "L1")
					{
						if (t.CompareTo(new DateTime(ConvertHelpers.ConvertToInt(nam), 1, 1)) > 0)
						{
							frmMessage frmMessage = new frmMessage("0621", "3", "3", "Thông báo");
							((Form)frmMessage).ShowDialog();
							result = false;
							return result;
						}
					}
					else if (t.AddYears(-1).CompareTo(new DateTime(ConvertHelpers.ConvertToInt(nam), 1, 1)) > 0)
					{
						frmMessage frmMessage = new frmMessage("0578", "3", "3", "Thông báo");
						((Form)frmMessage).ShowDialog();
						result = false;
						return result;
					}
				}
				else if (kieuKy.Equals("D"))
				{
					if (t.CompareTo(new DateTime(ConvertHelpers.ConvertToInt(nam), ConvertHelpers.ConvertToInt(thang), ConvertHelpers.ConvertToInt(ngay))) > 0)
					{
						frmMessage frmMessage = new frmMessage("0576", "3", "3", "Thông báo");
						((Form)frmMessage).ShowDialog();
						result = false;
						return result;
					}
				}
				else if (kieuKy.Equals("D_D"))
				{
					if (menuID == "C9" || menuID == "D6" || WinSettings.MenuID == "X9" || WinSettings.MenuID == "Y1")
					{
						if (t.AddYears(-1).CompareTo(ConvertHelpers.ConvertToDateTime(denNgay, "dd/MM/yyyy")) > 0)
						{
							frmMessage frmMessage = new frmMessage("0619", "3", "3", "Thông báo");
							((Form)frmMessage).ShowDialog();
							result = false;
							return result;
						}
					}
					else if (t.CompareTo(ConvertHelpers.ConvertToDateTime(tuNgay, "dd/MM/yyyy")) > 0)
					{
						frmMessage frmMessage = new frmMessage("0576", "3", "3", "Thông báo");
						((Form)frmMessage).ShowDialog();
						result = false;
						return result;
					}
				}
			}
			if ("~28~29~30~31~01~02~04~71~72~73~55~56~35~32~34~39~70~81~06~79~86~87~85~88~63~90~08~21~40~46~47~94~96~97~".Contains("~" + menuID + "~"))
			{
				if (kieuKy.Equals("M"))
				{
					if (DateTime.Compare(t, new DateTime(ConvertHelpers.ConvertToInt(nam), ConvertHelpers.ConvertToInt(thang), 1)) <= 0)
					{
						frmMessage frmMessage = new frmMessage("0581", "3", "3", "Thông báo");
						((Form)frmMessage).ShowDialog();
						result = false;
						return result;
					}
				}
				else if (kieuKy.Equals("Q"))
				{
					if (t.CompareTo(new DateTime(ConvertHelpers.ConvertToInt(nam), ConvertHelpers.ConvertToInt(quy), 1)) <= 0)
					{
						frmMessage frmMessage = new frmMessage("0583", "3", "3", "Thông báo");
						((Form)frmMessage).ShowDialog();
						result = false;
						return result;
					}
				}
				else if (kieuKy.Equals("Y"))
				{
					if (t.CompareTo(new DateTime(ConvertHelpers.ConvertToInt(nam), 1, 1)) <= 0)
					{
						frmMessage frmMessage = new frmMessage("0582", "3", "3", "Thông báo");
						((Form)frmMessage).ShowDialog();
						result = false;
						return result;
					}
				}
				else if (kieuKy.Equals("D"))
				{
					if (t.CompareTo(new DateTime(ConvertHelpers.ConvertToInt(nam), ConvertHelpers.ConvertToInt(thang), ConvertHelpers.ConvertToInt(ngay))) <= 0)
					{
						frmMessage frmMessage = new frmMessage("0580", "3", "3", "Thông báo");
						((Form)frmMessage).ShowDialog();
						result = false;
						return result;
					}
				}
				else if (kieuKy.Equals("D_D") && t.CompareTo(ConvertHelpers.ConvertToDateTime(denNgay, "dd/MM/yyyy")) <= 0)
				{
					frmMessage frmMessage = new frmMessage("0580", "3", "3", "Thông báo");
					((Form)frmMessage).ShowDialog();
					result = false;
					return result;
				}
			}
			if ("~03~37~33~80~82~77~44~42~83~87~88~55~".Contains("~" + menuID + "~"))
			{
				if (kieuKy.Equals("Y"))
				{
					if (dateTime.CompareTo(new DateTime(ConvertHelpers.ConvertToInt(nam), 1, 1)) <= 0)
					{
						frmMessage frmMessage = new frmMessage("0592", "3", "3", "Thông báo");
						((Form)frmMessage).ShowDialog();
						result = false;
						return result;
					}
				}
				else if (kieuKy.Equals("D_D") && dateTime.CompareTo(ConvertHelpers.ConvertToDateTime(denNgay, "dd/MM/yyyy")) <= 0)
				{
					frmMessage frmMessage = new frmMessage("0592", "3", "3", "Thông báo");
					((Form)frmMessage).ShowDialog();
					result = false;
					return result;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return result;
	}

	private bool CheckHLToKhaiDKTT61(HTKKPeriodInfo periodInfo)
	{
		//IL_00b2: Unknown result type (might be due to invalid IL or missing references)
		//IL_0124: Unknown result type (might be due to invalid IL or missing references)
		bool result = true;
		try
		{
			string kieuKy = periodInfo.KieuKy;
			string text = periodInfo.menuID;
			string nam = periodInfo.Nam;
			string thang = periodInfo.Thang;
			string ngay = periodInfo.Ngay;
			string quy = periodInfo.Quy;
			string obj = "01/05/2016";
			string obj2 = "01/01/2016";
			string tuNgay = periodInfo.TuNgay;
			DateTime dateTime = ConvertHelpers.ConvertToDateTime(obj, "dd/MM/yyyy");
			DateTime dateTime2 = ConvertHelpers.ConvertToDateTime(obj2, "dd/MM/yyyy");
			if (text.Equals("83"))
			{
				if (ConvertHelpers.ConvertToDateTime(tuNgay, "dd/MM/yyyy") < dateTime2)
				{
					frmMessage frmMessage = new frmMessage("0412", "3", "3", "Thông báo");
					((Form)frmMessage).ShowDialog();
					result = false;
					return result;
				}
			}
			else if (text.Equals("47") && dateTime > new DateTime(ConvertHelpers.ConvertToInt(nam), ConvertHelpers.ConvertToInt(thang), ConvertHelpers.ConvertToInt(ngay)))
			{
				frmMessage frmMessage = new frmMessage("0447", "3", "3", "Thông báo");
				((Form)frmMessage).ShowDialog();
				result = false;
				return result;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return result;
	}

	private bool CheckPeriodKhbs(HTKKPeriodInfo periodInfo, string strDataFileToKhaiChinh)
	{
		//IL_0165: Unknown result type (might be due to invalid IL or missing references)
		//IL_025f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0215: Unknown result type (might be due to invalid IL or missing references)
		//IL_06d9: Unknown result type (might be due to invalid IL or missing references)
		//IL_044a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0451: Expected O, but got Unknown
		//IL_047f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0486: Expected O, but got Unknown
		//IL_0693: Unknown result type (might be due to invalid IL or missing references)
		bool result = true;
		try
		{
			string text = "~02~04~71~72~11~12~73~03~70~80~81~82~06~77~05~86~87~88~90~94~96~97~98~99~93~92~89~85~55~56~A3~79~08~21~40~42~44~63~34~83~47~28~29~30~31~G5~G6~G7~G8~I5~K1~K6~K8~38~~L1~L2~L3~L4~L5~L6~L7~L8~L9~T2~T3~T4~T5~T6~T7~T8~T9~H2~D1~D2~D3~D4~D5~D6~D7~D8~D9~E1~E2~E3~E4~E5~E6~E7~E8~E9~I1~I2~I3~I4~I7~I8~I9~J1~J2~J3~J4~J5~J6~J7~J8~J9~M8~M9~K2~K3~K4~K5~K7~N8~B8~B9~C6~C7~C8~C9~N3~N4~G9~P2~P7~Q6~Q5~Q9~R4~R9~R6~U2~U1~U7~X5~X8~X6~X9~Y1~Y6~Y3~Y4~Z1~Z3~Y5~Z4~Z8~Z5~W1~~O1~O2~O3~O4~O5~O6~";
			string text2 = WinDoWorkEvent.GetFileNameToKhaiCt(WinSettings.TTPeriodInfo, strDataFileToKhaiChinh, HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst + "\\", WinSettings.ToKhaiInfo.VersionXML);
			if (WinSettings.TTPeriodInfo.menuID == "P7")
			{
				text2 = text2.Replace("_" + WinSettings.TTPeriodInfo.Ngay + WinSettings.TTPeriodInfo.Thang + WinSettings.TTPeriodInfo.Nam, "");
			}
			DateTime hanNopToKhai = GeneralFunctions.GetHanNopToKhai(WinSettings.TTPeriodInfo, WinSettings.HTKKHeaderInfo);
			DateTime dateTime = DateTime.ParseExact(periodInfo.NgayLapKHBS, "d/M/yyyy", null);
			DateTime now = DateTime.Now;
			TimeSpan timeSpan = now - dateTime;
			if (periodInfo.LoaiToKhai.Equals("B") && (text.IndexOf("~" + periodInfo.menuID + "~") > -1 || periodInfo.menuID.Equals("01")))
			{
				if (!File.Exists(text2))
				{
					if (periodInfo.menuID == "P7")
					{
						return result;
					}
					frmMessage frmMessage = new frmMessage("0107", "3", "3", "Thông báo");
					((Form)frmMessage).ShowDialog();
					result = false;
					return result;
				}
				if (ConvertHelpers.ConvertToInt(periodInfo.SoLan) > 1)
				{
					string[] array = text2.Split(new char[1] { '\\' });
					string path = text2.Replace(array[^1], "bs" + ConvertHelpers.ConvertToString(ConvertHelpers.ConvertToInt(periodInfo.SoLan) - 1) + "_" + array[^1]);
					if (!File.Exists(path))
					{
						frmMessage frmMessage = new frmMessage("0296", "3", "3", "Thông báo");
						((Form)frmMessage).ShowDialog();
						result = false;
						return result;
					}
				}
				if (timeSpan.TotalDays < 0.0)
				{
					frmMessage frmMessage = new frmMessage("0224", "3", "3", "Thông báo");
					((Form)frmMessage).ShowDialog();
					return false;
				}
				if (periodInfo.KieuKy.Equals("D"))
				{
					DateTime dateTime2 = new DateTime(int.Parse(periodInfo.Nam), int.Parse(periodInfo.Thang), int.Parse(periodInfo.Ngay));
					if (menuID.Equals("98") || menuID.Equals("92") || menuID.Equals("08") || menuID.Equals("40") || menuID.Equals("K2") || menuID.Equals("K3") || menuID.Equals("J9") || menuID.Equals("S1") || menuID.Equals("J8") || menuID.Equals("J7") || menuID.Equals("R6"))
					{
						dateTime2 = dateTime2.AddDays(35.0);
					}
					else if (menuID.Equals("47") || menuID == "K7")
					{
						dateTime2 = DateTime.ParseExact(periodInfo.NgayChiaCTLN, "dd/MM/yyyy", CultureInfo.CurrentCulture);
						DateTime dateTime3 = new DateTime(2020, 12, 5);
						dateTime2 = ((!(dateTime2 >= dateTime3)) ? dateTime2.AddDays(45.0) : dateTime2.AddDays(10.0));
					}
					else
					{
						dateTime2 = dateTime2.AddDays(10.0);
					}
					string text3 = "DMNgayLe.xml";
					XmlDocument val = new XmlDocument();
					val.Load(text3);
					XmlNode val2 = ((XmlNode)val).SelectSingleNode("Root");
					foreach (XmlNode item in val2)
					{
						XmlNode val3 = item;
						string value = ((XmlNode)val3.Attributes["KieuKy"]).Value;
						string value2 = ((XmlNode)val3.Attributes["HanNop"]).Value;
						string[] array2 = value2.Split(new char[1] { '/' });
						string value3 = ((XmlNode)val3.Attributes["GiaTri"]).Value;
						string[] array3 = value3.Split(new char[1] { '/' });
						DateTime dateTime4 = new DateTime(ConvertHelpers.ConvertToInt(array2[2]), ConvertHelpers.ConvertToInt(array2[1]), ConvertHelpers.ConvertToInt(array2[0]));
						DateTime dateTime5 = new DateTime(ConvertHelpers.ConvertToInt(array3[2]), ConvertHelpers.ConvertToInt(array3[1]), ConvertHelpers.ConvertToInt(array3[0]));
						if (periodInfo.KieuKy.Equals("D") && value.Equals("D"))
						{
							if (dateTime2 == dateTime4)
							{
								dateTime2 = dateTime5;
								break;
							}
						}
						else if (periodInfo.KieuKy.Equals("Q") && value.Equals("Q") && dateTime2 == dateTime4)
						{
							dateTime2 = dateTime5;
							break;
						}
					}
					if (dateTime2.DayOfWeek == DayOfWeek.Saturday)
					{
						dateTime2 = dateTime2.AddDays(2.0);
					}
					else if (dateTime2.DayOfWeek == DayOfWeek.Sunday)
					{
						dateTime2 = dateTime2.AddDays(1.0);
					}
					if (dateTime > dateTime2.AddYears(10))
					{
						frmMessage frmMessage = new frmMessage("0507", "3", "3", "Thông báo");
						((Form)frmMessage).ShowDialog();
						return false;
					}
				}
				else if (dateTime > hanNopToKhai.AddYears(10))
				{
					frmMessage frmMessage = new frmMessage("0507", "3", "3", "Thông báo");
					((Form)frmMessage).ShowDialog();
					return false;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return result;
	}

	private void txtLan_Leave(object sender, EventArgs e)
	{
		try
		{
			((Control)txtLan).Text = string.Concat(ConvertHelpers.ConvertToInt(((Control)txtLan).Text));
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void txtLan_TextChanged(object sender, EventArgs e)
	{
		try
		{
			if (((Control)txtLan).Text.Equals("0"))
			{
				((Control)txtLan).Text = "1";
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void txtLan_KeyPress(object sender, KeyPressEventArgs e)
	{
		try
		{
			e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != '\b';
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void txtLanXuatBan_Leave(object sender, EventArgs e)
	{
		try
		{
			if (((Control)pnVT2_PhuLucKeKhai).Visible && !((Control)txtLanXuatBan).Text.Equals(WinSettings.TTPeriodInfo.LanXuatBan))
			{
				LoadPhuLucSpread();
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void cboNganhKD_SelectedIndexChanged(object sender, EventArgs e)
	{
		//IL_0444: Unknown result type (might be due to invalid IL or missing references)
		//IL_04f6: Unknown result type (might be due to invalid IL or missing references)
		try
		{
			if (((Control)pnVT2_PhuLucKeKhai).Visible && !((ListControl)cboNganhKD).SelectedValue.Equals(WinSettings.TTPeriodInfo.DonViTinhMa))
			{
				if (((Control)pnVT1_5DanhMucNghe).Visible && (menuID == "D1" || WinSettings.MenuID.Equals("Z5")))
				{
					WinSettings.TTPeriodInfo.TieuMucHachToanVal = ((ListControl)cboNganhKD).SelectedValue.ToString();
				}
				LoadPhuLucSpread();
			}
			if (menuID == "40" || menuID == "08" || menuID == "K2" || menuID == "K3" || menuID == "J9" || menuID == "S1" || menuID == "J8" || menuID == "J7" || menuID == "R6" || menuID == "R9")
			{
				if (((ListControl)cboNganhKD).SelectedValue.ToString() == "2")
				{
					((Control)chkTKhaiLanXB).Enabled = false;
					chkTKhaiLanXB.Checked = false;
					chkTkhaiThang.Checked = true;
					((Control)chkTkhaiThang).Enabled = true;
				}
				else
				{
					((Control)chkTKhaiLanXB).Enabled = true;
					if (chkTkhaiThang.Checked)
					{
						chkTKhaiLanXB.Checked = true;
						chkTkhaiThang.Checked = false;
					}
					((Control)chkTkhaiThang).Enabled = false;
				}
			}
			if (menuID == "R6")
			{
				string text = ((ListControl)cboNganhKD).SelectedValue.ToStr();
				if (text == "4" || text == "5" || text == "6")
				{
					((Control)chkToChucCaNhanChuyenNhuong).Enabled = false;
					chkToChucCaNhanChuyenNhuong.Checked = false;
				}
				else
				{
					((Control)chkToChucCaNhanChuyenNhuong).Enabled = true;
				}
			}
			if (menuID == "A4" || menuID == "99")
			{
				if (((ListControl)cboNganhKD).SelectedValue.ToString() == "0039")
				{
					string[] array = ((Control)mskDenThang).Text.Split(new char[1] { '/' });
					if (array != null && array.Length > 1)
					{
						if (ConvertHelpers.ConvertToInt(array[1]) >= 2021)
						{
							((Control)cbMaHSDNHT_DADT).Enabled = true;
						}
						else
						{
							((Control)cbMaHSDNHT_DADT).Enabled = false;
							((Control)cbMaHSDNHT_DADT).Text = "";
						}
					}
					lockTinhXaCqt(isLock: true);
					loadDefaultTinhCqtTheoDiaBan();
				}
				else
				{
					((Control)cbMaHSDNHT_DADT).Enabled = false;
					((Control)cbMaHSDNHT_DADT).Text = "";
					lockTinhXaCqt(isLock: false);
				}
			}
			if (menuID == "T8" || menuID == "R5")
			{
				if (((ListControl)cboNganhKD).SelectedValue.ToString() != "03" && Convert.ToBoolean(axfpsPhuLuc.Sheets[0].Cells[1, 0].Value))
				{
					frmMessage frmMessage = new frmMessage("0572", "3", "3", "Thông báo");
					((Form)frmMessage).ShowDialog();
				}
				if (WinSettings.TTPeriodInfo.KieuKy == "D" && (((ListControl)cboNganhKD).SelectedValue.ToString() == "03" || ((ListControl)cboNganhKD).SelectedValue.ToString() == "02") && Convert.ToBoolean(axfpsPhuLuc.Sheets[0].Cells[1, 0].Value))
				{
					frmMessage frmMessage = new frmMessage("0586", "3", "3", "Thông báo");
					((Form)frmMessage).ShowDialog();
				}
			}
			if (menuID == "U1")
			{
				if (((ListControl)cboNganhKD).SelectedIndex == 0)
				{
					((Control)pnTinhCQT).Enabled = true;
					((Control)pnChuyenNhuong).Enabled = false;
					chkToChucCaNhanChuyenNhuong.Checked = false;
					chkDoanhNghiepChuyenNhuong.Checked = false;
				}
				else
				{
					loadDefaultTinhCqtTheoDiaBan();
					((Control)pnTinhCQT).Enabled = false;
					((Control)pnChuyenNhuong).Enabled = true;
				}
			}
			if (menuID == "Y3")
			{
				if (((ListControl)cboNganhKD).SelectedIndex == 0)
				{
					((Control)pnTinhCQT).Enabled = true;
					((Control)pnChuyenNhuong).Enabled = true;
					((Control)chkToChucCaNhanChuyenNhuong).Enabled = true;
					chkToChucCaNhanChuyenNhuong.Checked = false;
				}
				else
				{
					loadDefaultTinhCqtTheoDiaBan();
					((Control)pnTinhCQT).Enabled = false;
					((Control)pnChuyenNhuong).Enabled = true;
					((Control)chkToChucCaNhanChuyenNhuong).Enabled = false;
					chkToChucCaNhanChuyenNhuong.Checked = false;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void SetLayoutToKhaiNamThangQuyLanPhatSinh()
	{
		ShowDetailpnVT1_1CheckBox(chkTKhaiLanXBShow: false, chkTkhaiThangShow: true, chkTKKyShow: false, chkTKQuyShow: true, ckNamShow: true, chkTuThangDenThangShow: false, chkTKLanPSShow: true, chkTuQuyDenQuyShow: false);
		ShowDetailpnVT1_2NgayThangQuyNam(pnNgayShow: false, pnThangShow: false, pnKyShow: false, pnQuyShow: false, pnNamShow: true);
		ShowDetailpnVT1_3TuNgayDenNgay(pnLanXuatBanShow: false, pnTuNgayShow: false, pnDenNgayShow: false, pnTuThangShow: true, pnDenThangShow: true);
		ShowDetailpnVT1_4LanDauBoXung(pnQuyetToan_LienTuc: false, pnRDLanDauBoXungShow: true, pnNgayLapShow: false);
		ShowDetailpnVT1_5DanhMucNghe(pnVT1_5DanhMucNgheShow: false);
		ShowDetailpnVT2_PhuLucKeKhai(pnVT2_PhuLucKeKhaiShow: true);
		ckNam.Checked = true;
	}

	private void SetLayoutToQDDSThueMSTToChucCaNhanNamThangQuy()
	{
		ShowDetailpnVT1_1CheckBox(chkTKhaiLanXBShow: false, chkTkhaiThangShow: true, chkTKKyShow: true, chkTKQuyShow: true, ckNamShow: true, chkTuThangDenThangShow: false, chkTKLanPSShow: false, chkTuQuyDenQuyShow: false);
		ShowDetailpnVT1_2NgayThangQuyNam(pnNgayShow: false, pnThangShow: false, pnKyShow: false, pnQuyShow: false, pnNamShow: true);
		ShowDetailpnVT1_3TuNgayDenNgay(pnLanXuatBanShow: false, pnTuNgayShow: false, pnDenNgayShow: false, pnTuThangShow: true, pnDenThangShow: true);
		ShowDetailpnVT1_4LanDauBoXung(pnQuyetToan_LienTuc: false, pnRDLanDauBoXungShow: true, pnNgayLapShow: false);
		ShowDetailpnVT1_5DanhMucNghe(pnVT1_5DanhMucNgheShow: false);
		ShowDetailpnVT2_PhuLucKeKhai(pnVT2_PhuLucKeKhaiShow: true);
		if (menuID.Equals("Q8") || menuID.Equals("U5"))
		{
			ShowDetailpnVT1_1CheckBox(chkTKhaiLanXBShow: true, chkTkhaiThangShow: true, chkTKKyShow: true, chkTKQuyShow: true, ckNamShow: true, chkTuThangDenThangShow: false, chkTKLanPSShow: false, chkTuQuyDenQuyShow: false);
			((Control)chkTKKy).Text = "Tờ khai kỳ 6 tháng";
			((Control)chkTKhaiLanXB).Text = "TK lần phát sinh";
		}
		else
		{
			((Control)chkTKKy).Text = "Tờ khai theo kỳ thanh toán";
		}
		((Control)pnVT1_1CheckBox).Size = new Size(360, 50);
		ckNam.Checked = true;
		((Control)ckNam).Location = new Point(((Control)chkTKhaiLanXB).Location.X, ((Control)ckNam).Location.Y);
		((Control)chkTKKy).Location = new Point(((Control)ckNam).Location.X + 20, ((Control)ckNam).Location.Y);
		((Control)chkTKQuy).Location = new Point(((Control)ckNam).Location.X + 20, ((Control)ckNam).Location.Y + ((Control)chkTkhaiThang).Height + ctrPadR - 3);
		((Control)chkTkhaiThang).Location = new Point(((Control)ckNam).Location.X, ((Control)ckNam).Location.Y + ((Control)chkTkhaiThang).Height + ctrPadR - 3);
		Resize();
	}

	private bool CheckHieuLucTK()
	{
		bool result = true;
		try
		{
			string text = startDate;
			string kieuKy = WinSettings.TTPeriodInfo.KieuKy;
			string[] array = text.Split(new char[1] { '/' });
			if (menuID.Equals("33") || menuID.Equals("35") || menuID.Equals("37") || menuID.Equals("38") || menuID.Equals("39"))
			{
				if (menuID.Equals("33") || menuID.Equals("37") || menuID.Equals("39"))
				{
					string[] array2 = ((Control)mskTuThang).Text.Split(new char[1] { '/' });
					DateTime t = new DateTime(ConvertHelpers.ConvertToInt(array2[1]), ConvertHelpers.ConvertToInt(array2[0]), 1);
					DateTime t2 = new DateTime(ConvertHelpers.ConvertToInt(array[2]), ConvertHelpers.ConvertToInt(array[1]), ConvertHelpers.ConvertToInt(array[0]));
					if (DateTime.Compare(t, t2) < 0)
					{
						result = false;
					}
				}
				else
				{
					DateTime t = GeneralFunctions.GetNgayDauQuy(ConvertHelpers.ConvertToInt(((Control)cbbQuy).Text), ConvertHelpers.ConvertToInt(((Control)mskYear).Text), 1, 1);
					DateTime t2 = new DateTime(ConvertHelpers.ConvertToInt(array[2]), ConvertHelpers.ConvertToInt(array[1]), ConvertHelpers.ConvertToInt(array[0]));
					if (DateTime.Compare(t, t2) < 0)
					{
						result = false;
					}
				}
			}
			else if (menuID.Equals("29"))
			{
				if (ConvertHelpers.ConvertToInt(((Control)mskYear).Text) < ConvertHelpers.ConvertToInt(array[2]))
				{
					result = false;
				}
			}
			else if (menuID.Equals("79"))
			{
				if (kieuKy.Equals("M"))
				{
					if (!HTKKFormater.isNullString(((Control)mskYear).Text) && !HTKKFormater.isNullString(((Control)mskMonth).Text))
					{
						DateTime t = new DateTime(ConvertHelpers.ConvertToInt(((Control)mskYear).Text), ConvertHelpers.ConvertToInt(((Control)mskMonth).Text), 1);
						DateTime t2 = new DateTime(ConvertHelpers.ConvertToInt(array[2]), ConvertHelpers.ConvertToInt(array[1]), ConvertHelpers.ConvertToInt(array[0]));
						if (DateTime.Compare(t, t2) < 0)
						{
							result = false;
						}
					}
				}
				else if (kieuKy.Equals("D") && !HTKKFormater.isNullString(((Control)mskYear).Text) && !HTKKFormater.isNullString(((Control)mskMonth).Text) && !HTKKFormater.isNullString(((Control)mskDay).Text))
				{
					DateTime t = new DateTime(ConvertHelpers.ConvertToInt(((Control)mskYear).Text), ConvertHelpers.ConvertToInt(((Control)mskMonth).Text), ConvertHelpers.ConvertToInt(((Control)mskDay).Text));
					DateTime t2 = new DateTime(ConvertHelpers.ConvertToInt(array[2]), ConvertHelpers.ConvertToInt(array[1]), ConvertHelpers.ConvertToInt(array[0]));
					if (DateTime.Compare(t, t2) < 0)
					{
						result = false;
					}
				}
			}
			else if (kieuKy.Equals("Y"))
			{
				if (ConvertHelpers.ConvertToInt(((Control)mskYear).Text) < ConvertHelpers.ConvertToInt(array[2]))
				{
					result = false;
				}
				if (((Control)mskTuThang).Visible && ((Control)mskDenThang).Visible && !HTKKFormater.isNullString(((Control)mskTuThang).Text) && !HTKKFormater.isNullString(((Control)mskDenThang).Text))
				{
					string[] array2 = ((Control)mskTuThang).Text.Split(new char[1] { '/' });
					DateTime t = new DateTime(ConvertHelpers.ConvertToInt(array2[1]), ConvertHelpers.ConvertToInt(array2[0]), 1);
					DateTime t2 = new DateTime(ConvertHelpers.ConvertToInt(array[2]), 1, 1);
					if (DateTime.Compare(t, t2) < 0)
					{
						result = false;
					}
				}
			}
			else if (kieuKy.Equals("M"))
			{
				if (!HTKKFormater.isNullString(((Control)mskYear).Text) && !HTKKFormater.isNullString(((Control)mskMonth).Text))
				{
					DateTime t = new DateTime(ConvertHelpers.ConvertToInt(((Control)mskYear).Text), ConvertHelpers.ConvertToInt(((Control)mskMonth).Text), 1);
					DateTime t2 = new DateTime(ConvertHelpers.ConvertToInt(array[2]), ConvertHelpers.ConvertToInt(array[1]), ConvertHelpers.ConvertToInt(array[0]));
					if (DateTime.Compare(t, t2) < 0)
					{
						result = false;
					}
				}
			}
			else if (kieuKy.Equals("Q"))
			{
				if (!HTKKFormater.isNullString(((Control)mskYear).Text) && !HTKKFormater.isNullString(((Control)cbbQuy).Text))
				{
					DateTime t = GeneralFunctions.GetNgayDauQuy(ConvertHelpers.ConvertToInt(((Control)cbbQuy).Text), ConvertHelpers.ConvertToInt(((Control)mskYear).Text), 1, 1);
					DateTime t2 = new DateTime(ConvertHelpers.ConvertToInt(array[2]), ConvertHelpers.ConvertToInt(array[1]), ConvertHelpers.ConvertToInt(array[0]));
					if (DateTime.Compare(t, t2) < 0)
					{
						result = false;
					}
				}
			}
			else if (kieuKy.Equals("D") && !HTKKFormater.isNullString(((Control)mskYear).Text) && !HTKKFormater.isNullString(((Control)mskMonth).Text) && !HTKKFormater.isNullString(((Control)mskDay).Text))
			{
				DateTime t = new DateTime(ConvertHelpers.ConvertToInt(((Control)mskYear).Text), ConvertHelpers.ConvertToInt(((Control)mskMonth).Text), ConvertHelpers.ConvertToInt(((Control)mskDay).Text));
				DateTime t2 = new DateTime(ConvertHelpers.ConvertToInt(array[2]), ConvertHelpers.ConvertToInt(array[1]), ConvertHelpers.ConvertToInt(array[0]));
				if (menuID.Equals("28"))
				{
					if (DateTime.Compare(t, new DateTime(2016, 2, 1)) < 0)
					{
						result = false;
					}
				}
				else if (DateTime.Compare(t, t2) < 0)
				{
					result = false;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
		return result;
	}

	private bool CheckPeriodTncn(HTKKPeriodInfo periodInfo)
	{
		//IL_00b8: Unknown result type (might be due to invalid IL or missing references)
		//IL_0132: Unknown result type (might be due to invalid IL or missing references)
		//IL_010d: Unknown result type (might be due to invalid IL or missing references)
		//IL_02a1: Unknown result type (might be due to invalid IL or missing references)
		//IL_027c: Unknown result type (might be due to invalid IL or missing references)
		//IL_023d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0218: Unknown result type (might be due to invalid IL or missing references)
		//IL_01c2: Unknown result type (might be due to invalid IL or missing references)
		//IL_019d: Unknown result type (might be due to invalid IL or missing references)
		bool flag = true;
		try
		{
			string kieuKy = periodInfo.KieuKy;
			string text = "~28~29~30~31~32~33~34~35~37~38~39~";
			if (text.IndexOf("~" + menuID + "~") > -1)
			{
				flag = CheckHieuLucTK();
				if (!flag)
				{
					if (menuID.Equals("33") || menuID.Equals("37") || menuID.Equals("29") || menuID.Equals("39"))
					{
						frmMessage frmMessage = new frmMessage("0407", "3", "3", "Thông báo");
						((Form)frmMessage).ShowDialog();
					}
					else if (kieuKy.Equals("M"))
					{
						if (menuID.Equals("28"))
						{
							frmMessage frmMessage = new frmMessage("0409", "3", "3", "Thông báo");
							((Form)frmMessage).ShowDialog();
						}
						else
						{
							frmMessage frmMessage = new frmMessage("0404", "3", "3", "Thông báo");
							((Form)frmMessage).ShowDialog();
						}
					}
					else if (kieuKy.Equals("Q"))
					{
						if (menuID.Equals("28") || menuID.Equals("38"))
						{
							frmMessage frmMessage = new frmMessage("0410", "3", "3", "Thông báo");
							((Form)frmMessage).ShowDialog();
						}
						else
						{
							frmMessage frmMessage = new frmMessage("0405", "3", "3", "Thông báo");
							((Form)frmMessage).ShowDialog();
						}
					}
					else if (kieuKy.Equals("D"))
					{
						if (menuID.Equals("28"))
						{
							frmMessage frmMessage = new frmMessage("0442", "3", "3", "Thông báo");
							((Form)frmMessage).ShowDialog();
						}
						else
						{
							frmMessage frmMessage = new frmMessage("0406", "3", "3", "Thông báo");
							((Form)frmMessage).ShowDialog();
						}
					}
					else if (menuID.Equals("28"))
					{
						frmMessage frmMessage = new frmMessage("0413", "3", "3", "Thông báo");
						((Form)frmMessage).ShowDialog();
					}
					else
					{
						frmMessage frmMessage = new frmMessage("0371", "3", "3", "Thông báo");
						((Form)frmMessage).ShowDialog();
					}
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
		return flag;
	}

	private void SetLayoutToKhaiNamKyTuNgayDenNgay()
	{
		ShowDetailpnVT1_1CheckBox(chkTKhaiLanXBShow: false, chkTkhaiThangShow: false, chkTKKyShow: true, chkTKQuyShow: false, ckNamShow: true, chkTuThangDenThangShow: false, chkTKLanPSShow: false, chkTuQuyDenQuyShow: false);
		ShowDetailpnVT1_2NgayThangQuyNam(pnNgayShow: true, pnThangShow: false, pnKyShow: false, pnQuyShow: false, pnNamShow: false);
		ShowDetailpnVT1_3TuNgayDenNgay(pnLanXuatBanShow: false, pnTuNgayShow: false, pnDenNgayShow: false, pnTuThangShow: false, pnDenThangShow: false);
		ShowDetailpnVT1_4LanDauBoXung(pnQuyetToan_LienTuc: false, pnRDLanDauBoXungShow: true, pnNgayLapShow: false);
		ShowDetailpnVT1_5DanhMucNghe(pnVT1_5DanhMucNgheShow: false);
		ShowDetailpnVT2_PhuLucKeKhai(pnVT2_PhuLucKeKhaiShow: true);
		ckNam.Checked = true;
		((Control)chkTKKy).Text = "Tờ khai theo kỳ thanh toán";
	}

	private void cbbKy_SelectedIndexChanged(object sender, EventArgs e)
	{
		try
		{
			if (menuID == "X1" || menuID == "X2")
			{
				thayDoiTuNgayDenNgayTheoKy();
			}
			else if (strKieuKy == "7" && !string.IsNullOrEmpty(((Control)mskYear).Text))
			{
				if (((ListControl)cbbKy).SelectedIndex == 1)
				{
					((Control)mskTuNgay).Text = "01/07" + ((Control)mskYear).Text;
					((Control)mskDenNgay).Text = DateTime.DaysInMonth(ConvertHelpers.ConvertToInt(((Control)mskYear).Text), 12) + "/12/" + ((Control)mskYear).Text;
				}
				if (((ListControl)cbbKy).SelectedIndex == 0)
				{
					((Control)mskTuNgay).Text = "01/01" + ((Control)mskYear).Text;
					((Control)mskDenNgay).Text = DateTime.DaysInMonth(ConvertHelpers.ConvertToInt(((Control)mskYear).Text), 6) + "/06/" + ((Control)mskYear).Text;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	private void thayDoiTuNgayDenNgayTheoKy()
	{
		if ((menuID == "X1" || menuID == "X2") && chkTKKy.Checked && !string.IsNullOrEmpty(((Control)mskYear).Text))
		{
			if (((ListControl)cbbKy).SelectedIndex == 0)
			{
				DateTime ngayBatDauTheoTaiChinh = GeneralFunctions.ngayBatDauTheoTaiChinh(WinSettings.HTKKHeaderInfo, ((Control)mskYear).Text);
				((Control)mskTuNgay).Text = ngayBatDauTheoTaiChinh.ToString("dd/MM/yyyy");
				((Control)mskDenNgay).Text = GeneralFunctions.ngayKetThuc6ThangDauTheoTaiChinh(ngayBatDauTheoTaiChinh).ToString("dd/MM/yyyy");
			}
			if (((ListControl)cbbKy).SelectedIndex == 1)
			{
				DateTime ngayCuoiNamTaiChinh = GeneralFunctions.NgayCuoiNamTaiChinh(WinSettings.HTKKHeaderInfo, ((Control)mskYear).Text);
				((Control)mskTuNgay).Text = GeneralFunctions.ngayBatDau6ThangCuoiNam_TheoTaiChinh(ngayCuoiNamTaiChinh).ToString("dd/MM/yyyy");
				((Control)mskDenNgay).Text = ngayCuoiNamTaiChinh.ToString("dd/MM/yyyy");
			}
		}
	}

	private void chk_CoSoThanhLap_CheckedChanged(object sender, EventArgs e)
	{
		if (chk_CoSoThanhLap.Checked)
		{
			((Control)lbNgayPS).Visible = true;
			((Control)mskNgayPS).Visible = true;
			if (menuID == "L1")
			{
				((Control)lbNgayPS).Location = new Point(((Control)chk_CoSoThanhLap).Location.X + 90, ((Control)chk_CoSoThanhLap).Location.Y + 30);
				((Control)mskNgayPS).Location = new Point(((Control)chk_CoSoThanhLap).Location.X + 190, ((Control)chk_CoSoThanhLap).Location.Y + 27);
				((Control)pnQuyetToan_LienTuc).Height = ((Control)chk_CoSoThanhLap).Height + 40;
				Resize();
			}
			if (WinSettings.TTPeriodInfo.dicVar.ContainsKey("CoSoMoiThanhLap"))
			{
				WinSettings.TTPeriodInfo.dicVar.Remove("CoSoMoiThanhLap");
			}
			WinSettings.TTPeriodInfo.dicVar.Add("CoSoMoiThanhLap", "true");
			if (WinSettings.TTPeriodInfo.dicVar.ContainsKey("NgayThanhLap"))
			{
				WinSettings.TTPeriodInfo.dicVar.Remove("NgayThanhLap");
			}
			WinSettings.TTPeriodInfo.dicVar.Add("NgayThanhLap", ((Control)mskNgayPS).Text);
		}
		else
		{
			((Control)lbNgayPS).Visible = false;
			((Control)mskNgayPS).Visible = false;
			if (menuID == "L1")
			{
				((Control)pnQuyetToan_LienTuc).Height = ((Control)chk_CoSoThanhLap).Height + 10;
				((Control)mskNgayPS).Text = "";
				Resize();
			}
		}
		SetPeriodCustom();
	}

	private void mskNgayPS_Leave(object sender, EventArgs e)
	{
		if (!HTKKFormater.isNullString(((Control)mskNgayPS).Text.Replace("/", "").Replace("_", "")))
		{
			string text = HTKKFormater.FormatDdmmyyyy(((Control)mskNgayPS).Text);
			((Control)mskNgayPS).Text = text;
		}
	}

	private void SetPeriodCustom()
	{
		try
		{
			GetPeriodInfo(WinSettings.TTPeriodInfo);
			string dataFileNameExtend = WinFnProcess.GetDataFileNameExtend(WinSettings.TTPeriodInfo);
			string versionXML = "";
			if (WinSettings.TTPeriodInfo.VersionXML != null && WinSettings.TTPeriodInfo.VersionXML != "")
			{
				versionXML = WinSettings.TTPeriodInfo.VersionXML;
			}
			string fileNameBySheet = WinDoWorkEvent.GetFileNameBySheet(WinSettings.TTPeriodInfo, strDataFileToKhaiChinh, HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst + "\\", versionXML);
			if (menuID == "10" || menuID == "67")
			{
				if (File.Exists(fileNameBySheet))
				{
					string valueByAttributes = HTKKFile.GetValueByAttributes(fileNameBySheet, "/Sections/Section/Cells/Cell[@Path='HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ngayKy']");
					if (menuID == "67")
					{
						valueByAttributes = HTKKFile.GetValueByAttributes(fileNameBySheet, "/Sections/Section/Cells/Cell[@Path='HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ngayBCao']");
					}
					if (!HTKKFormater.isNullString(valueByAttributes))
					{
						valueByAttributes = valueByAttributes.Replace("/", "");
						((Control)mskDay).Text = valueByAttributes.Substring(0, 2);
						((Control)mskMonth).Text = valueByAttributes.Substring(2, 2);
						((Control)mskYear).Text = valueByAttributes.Substring(4, 4);
					}
					else
					{
						((Control)mskDay).Text = DateTime.Now.Day.ToString().PadLeft(2, '0');
						((Control)mskMonth).Text = DateTime.Now.Month.ToString().PadLeft(2, '0');
						((Control)mskYear).Text = DateTime.Now.Year.ToString();
					}
				}
				else
				{
					((Control)mskDay).Text = DateTime.Now.Day.ToString().PadLeft(2, '0');
					((Control)mskMonth).Text = DateTime.Now.Month.ToString().PadLeft(2, '0');
					((Control)mskYear).Text = DateTime.Now.Year.ToString();
				}
			}
			if (menuID == "D2")
			{
				((Control)cbMaHSDNHT_DADT).Enabled = true;
			}
			if (menuID == "02")
			{
				((Control)txtMaHSDNHT_DADT).Enabled = true;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private string CheckHieuLuc(string ngayHlForm)
	{
		string result = "";
		try
		{
			DateTime dateTime = ConvertHelpers.ConvertToDateTime(ngayHlForm, "dd/MM/yyyy");
			if (WinSettings.xmlMenuNode.HasChildNodes)
			{
				XmlNode val = WinSettings.xmlMenuNode.ChildNodes[0];
				if (val.Attributes["StartDate"] != null)
				{
					string text = HTKKFormater.toDateDDMMYYYY(((XmlNode)val.Attributes["StartDate"]).Value, isMess: false);
					if (HTKKFormater.isNullString(text))
					{
						return "";
					}
					DateTime dateTime2 = ConvertHelpers.ConvertToDateTime(text, "dd/MM/yyyy");
					if ((dateTime - dateTime2).TotalDays < 0.0)
					{
						result = ((XmlNode)val.Attributes["StartDate"]).Value;
					}
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return result;
	}

	private void UpdatePeriodByVerSion()
	{
		try
		{
			XmlNode xmlMenuNode = WinSettings.xmlMenuNode;
			WinSettings.xmlMenuNode = xmlNodeGoc;
			GetPeriodInfo(WinSettings.TTPeriodInfo);
			bool flag = GeneralFunctions.checkVersionDateEffect(ref WinSettings.xmlMenuNode, WinSettings.TTPeriodInfo, WinDoWorkEvent.getMenuNode(WinSettings.MenuID));
			xmlNode = WinSettings.xmlMenuNode;
			if (xmlNode != null)
			{
				WinSettings.TTPeriodInfo.VersionXML = ((xmlNode.Attributes["VersionXML"] == null) ? "" : ((XmlNode)xmlNode.Attributes["VersionXML"]).Value.Trim());
				WinSettings.ToKhaiInfo.VersionXML = ((xmlNode.Attributes["VersionXML"] == null) ? "" : ((XmlNode)xmlNode.Attributes["VersionXML"]).Value.Trim());
				WinSettings.ToKhaiInfo.VersionKBSXML = ((xmlNode.Attributes["VersionKHBS"] == null) ? "" : ((XmlNode)xmlNode.Attributes["VersionKHBS"]).Value.Trim());
				WinSettings.ToKhaiInfo.KieuKHBS = ((xmlNode.Attributes["KieuKHBS"] == null) ? "" : ((XmlNode)xmlNode.Attributes["KieuKHBS"]).Value.Trim());
			}
			if (flag)
			{
				SetDataFile();
				ControlInit(xmlNode);
				FormInit();
			}
			if (WinSettings.xmlMenuNode.ChildNodes[0] == null)
			{
				return;
			}
			XmlNodeList val = (WinSettings.xmlMenuNode.ChildNodes[0].HasChildNodes ? WinSettings.xmlMenuNode.ChildNodes[0].ChildNodes : null);
			if (WinSettings.ToKhaiInfo.KieuKHBS != "3")
			{
				if ((val == null || val.Count <= 2 || (OptBosung.Checked && !HTKKFormater.isNullString(WinSettings.ToKhaiInfo.KieuKHBS))) && (WinSettings.MenuID != "33" || WinSettings.MenuID != "37") && !"30,I5,K9,K1,M7,K6,K8,73,55,71,31,P2,X6".Contains(menuID))
				{
					((Control)pnVT2_PhuLucKeKhai).Visible = false;
				}
				else if ((WinSettings.MenuID == "79" && WinSettings.TTPeriodInfo.KieuKy == "D") || WinSettings.MenuID == "N7" || WinSettings.MenuID == "P2")
				{
					((Control)pnVT2_PhuLucKeKhai).Visible = false;
				}
				else
				{
					((Control)pnVT2_PhuLucKeKhai).Visible = true;
					LoadPhuLucSpread();
					if (WinSettings.MenuID == "33" || WinSettings.MenuID == "37")
					{
						((Control)pnVT2_PhuLucKeKhai).Visible = false;
					}
				}
			}
			else
			{
				if (val == null || val.Count <= 3)
				{
					((Control)pnVT2_PhuLucKeKhai).Visible = false;
				}
				else if ((WinSettings.MenuID == "T8" || WinSettings.MenuID == "R5") && WinSettings.TTPeriodInfo.KieuKy == "D")
				{
					((Control)pnVT2_PhuLucKeKhai).Visible = false;
				}
				else
				{
					((Control)pnVT2_PhuLucKeKhai).Visible = true;
					LoadPhuLucSpread();
					if (WinSettings.MenuID == "B9" || WinSettings.MenuID == "B8" || WinSettings.MenuID.Equals("W1"))
					{
						((Control)pnVT2_PhuLucKeKhai).Visible = false;
					}
				}
				if (OptBosung.Checked)
				{
					((Control)pnNgayLap).Visible = true;
				}
			}
			if (menuID == "Q5" || menuID == "Q6")
			{
				((Control)pnVT2_PhuLucKeKhai).Visible = false;
			}
			Resize();
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void SetWidthNgayThangNam(Control pCtl)
	{
		//IL_001d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0023: Expected O, but got Unknown
		//IL_009b: Unknown result type (might be due to invalid IL or missing references)
		//IL_00a1: Expected O, but got Unknown
		try
		{
			int num = 0;
			int num2 = 0;
			foreach (Control item in (ArrangedElementCollection)pCtl.Controls)
			{
				Control val = item;
				if (val.Width > num)
				{
					num = val.Width;
				}
				if (val.Height > num2)
				{
					num2 = val.Height;
				}
			}
			foreach (Control item2 in (ArrangedElementCollection)pCtl.Controls)
			{
				Control val = item2;
				val.Width = num;
				val.Height = num2;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void setQuyByNgayBatDauTC(int quy, int nam)
	{
		try
		{
			int iNgayBdNamTc = ConvertHelpers.ConvertToInt(WinSettings.HTKKHeaderInfo.NgayBatDauNamTaiChinh.Split(new char[1] { '/' })[0]);
			int iThangBdNamTc = ConvertHelpers.ConvertToInt(WinSettings.HTKKHeaderInfo.NgayBatDauNamTaiChinh.Split(new char[1] { '/' })[1]);
			DateTime ngayDauQuy = GeneralFunctions.GetNgayDauQuy(quy, nam, iNgayBdNamTc, iThangBdNamTc);
			DateTime ngayCuoiQuy = GeneralFunctions.GetNgayCuoiQuy(quy, nam, iNgayBdNamTc, iThangBdNamTc);
			((Control)mskTuNgay).Text = ngayDauQuy.ToString("dd/MM/yyyy");
			((Control)mskDenNgay).Text = ngayCuoiQuy.ToString("dd/MM/yyyy");
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void cboNganhKD_KeyPress(object sender, KeyPressEventArgs e)
	{
		e.Handled = true;
	}

	private void cbbKy_KeyPress(object sender, KeyPressEventArgs e)
	{
		e.Handled = true;
	}

	private void cbbQuy_KeyPress(object sender, KeyPressEventArgs e)
	{
		e.Handled = true;
	}

	private void frmHTKKPeriod_Shown(object sender, EventArgs e)
	{
		//IL_007a: Unknown result type (might be due to invalid IL or missing references)
		if ("0" == WinSettings.KieuKy || menuID == "66" || menuID == "52")
		{
			btnDongY_Click(sender, e);
		}
		if (menuID == "Y6")
		{
			frmMessage frmMessage = new frmMessage("0696", "3", "", "Thông báo");
			((Form)frmMessage).ShowDialog();
		}
	}

	private void loadDMNN_01GTGT()
	{
		//IL_007e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0085: Expected O, but got Unknown
		//IL_0225: Unknown result type (might be due to invalid IL or missing references)
		//IL_022f: Expected O, but got Unknown
		if (!(menuID == "01") && !(menuID == "I5"))
		{
			return;
		}
		XmlDataDocument xmlDataDocument = new XmlDataDocument();
		FileStream fileStream = new FileStream(HTKKConfig.INTERFACE_INI + "\\Catalogue_DM_NNKD.xml", FileMode.Open, FileAccess.Read);
		((XmlDocument)xmlDataDocument).Load((Stream)fileStream);
		XmlNodeList elementsByTagName = ((XmlDocument)xmlDataDocument).GetElementsByTagName("Item");
		Dictionary<string, string> dictionary = new Dictionary<string, string>();
		foreach (XmlNode item in elementsByTagName)
		{
			XmlNode val = item;
			string value = ((XmlNode)val.Attributes["Value"]).Value;
			string[] array = Regex.Split(value, "###");
			if (array.Length < 3 || !array[0].Equals("01_GTGT_240"))
			{
				continue;
			}
			if (chkTkhaiThang.Checked)
			{
				if (ConvertHelpers.ConvertToInt(((Control)mskYear).Text) >= 2021)
				{
					if (ConvertHelpers.ConvertToInt(array[1]) < 7)
					{
						dictionary.Add(array[1], array[2]);
					}
				}
				else if (ConvertHelpers.ConvertToInt(array[1]) >= 7)
				{
					dictionary.Add(array[1], array[2]);
				}
			}
			else if (ConvertHelpers.ConvertToInt(((Control)mskYear).Text) >= 2021)
			{
				if (ConvertHelpers.ConvertToInt(array[1]) < 7)
				{
					dictionary.Add(array[1], array[2]);
				}
			}
			else if (ConvertHelpers.ConvertToInt(array[1]) >= 7)
			{
				dictionary.Add(array[1], array[2]);
			}
		}
		if (dictionary != null && dictionary.Count > 0)
		{
			cboNganhKD.DataSource = (object)new BindingSource((object)dictionary, (string)null);
			((ListControl)cboNganhKD).DisplayMember = "Value";
			((ListControl)cboNganhKD).ValueMember = "Key";
			((Control)cboNganhKD).Refresh();
		}
	}

	private void label3_Click(object sender, EventArgs e)
	{
	}

	private void setTenChuDADT(string fileName)
	{
		try
		{
			string text = HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst;
			string pathFileXml = text + "\\" + fileName + ".xml";
			((Control)lblChuDADT).Text = "Tên dự án đầu tư: " + HTKKFile.GetValueByAttributes(pathFileXml, "/Sections/Section/Cells/Cell[@Path='HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/TTDuAnDauTu/tenDuAn']");
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private string maDuAnDT(string maToKhai)
	{
		try
		{
			return "";
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return "";
		}
	}

	private void cbMaHSDNHT_DADT_KeyPress(object sender, KeyPressEventArgs e)
	{
		try
		{
			e.Handled = true;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void cbMaHSDNHT_DADT_TextChanged(object sender, EventArgs e)
	{
		try
		{
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void txtMaHSDNHT_DADT_Leave(object sender, EventArgs e)
	{
		try
		{
			string tenChuDADT = "";
			if (!((Control)txtMaHSDNHT_DADT).Text.Equals(""))
			{
				if (ConvertHelpers.ConvertToDouble(((Control)txtMaHSDNHT_DADT).Text) == 0.0)
				{
					((Control)txtMaHSDNHT_DADT).Text = "1";
				}
				else
				{
					((Control)txtMaHSDNHT_DADT).Text = ConvertHelpers.ConvertToString(ConvertHelpers.ConvertToDouble(((Control)txtMaHSDNHT_DADT).Text));
				}
			}
			if (WinSettings.MenuID == "A4")
			{
				if (chkTuThangDenThang.Checked)
				{
					tenChuDADT = "01_DNHT_" + ((Control)txtMaHSDNHT_DADT).Text + "_" + ((Control)mskTuThang).Text.Replace("/", "") + "_" + ((Control)mskDenThang).Text.Replace("/", "") + "_" + WinSettings.TTPeriodInfo.VersionXML.Replace(".", "");
				}
				if (chkTuQuyDenQuy.Checked)
				{
					tenChuDADT = "01_DNHT_" + ((Control)txtMaHSDNHT_DADT).Text + "_Q0" + ((Control)mskTuThang).Text.Replace("/", "") + "_0" + ((Control)mskDenThang).Text.Replace("/", "") + "_" + WinSettings.TTPeriodInfo.VersionXML.Replace(".", "");
				}
			}
			else if (WinSettings.MenuID == "99")
			{
				switch (((ListControl)cboKieuKy).SelectedValue.ToString())
				{
				case "1":
					tenChuDADT = "01_DNHT_" + ((Control)txtMaHSDNHT_DADT).Text + "_TT" + ((Control)mskTuThang).Text.Replace("/", "").PadLeft(6, '0') + "_" + ((Control)mskDenThang).Text.Replace("/", "").PadLeft(6, '0') + "_" + WinSettings.TTPeriodInfo.VersionXML.Replace(".", "");
					break;
				case "2":
					tenChuDADT = "01_DNHT_" + ((Control)txtMaHSDNHT_DADT).Text + "_QQ" + ((Control)mskTuThang).Text.Replace("/", "").PadLeft(6, '0') + "_" + ((Control)mskDenThang).Text.Replace("/", "").PadLeft(6, '0') + "_" + WinSettings.TTPeriodInfo.VersionXML.Replace(".", "");
					break;
				case "3":
					tenChuDADT = "01_DNHT_" + ((Control)txtMaHSDNHT_DADT).Text + "_TQ" + ((Control)mskTuThang).Text.Replace("/", "").PadLeft(6, '0') + "_" + ((Control)mskDenThang).Text.Replace("/", "").PadLeft(6, '0') + "_" + WinSettings.TTPeriodInfo.VersionXML.Replace(".", "");
					break;
				case "4":
					tenChuDADT = "01_DNHT_" + ((Control)txtMaHSDNHT_DADT).Text + "_QT" + ((Control)mskTuThang).Text.Replace("/", "").PadLeft(6, '0') + "_" + ((Control)mskDenThang).Text.Replace("/", "").PadLeft(6, '0') + "_" + WinSettings.TTPeriodInfo.VersionXML.Replace(".", "");
					break;
				}
			}
			else if (WinSettings.MenuID == "Q4")
			{
				string text = ((ListControl)cboKieuKy).SelectedValue.ToString();
				if (text == "5")
				{
					tenChuDADT = "02_HT_" + ((Control)txtMaHSDNHT_DADT).Text + "_YY" + ((Control)mskTuThang).Text.Replace("/", "").PadLeft(6, '0') + "_" + ((Control)mskDenThang).Text.Replace("/", "").PadLeft(6, '0') + "_" + WinSettings.TTPeriodInfo.VersionXML.Replace(".", "");
				}
				switch (text)
				{
				case "1":
					tenChuDADT = "02_HT_" + ((Control)txtMaHSDNHT_DADT).Text + "_TT" + ((Control)mskTuThang).Text.Replace("/", "").PadLeft(6, '0') + "_" + ((Control)mskDenThang).Text.Replace("/", "").PadLeft(6, '0') + "_" + WinSettings.TTPeriodInfo.VersionXML.Replace(".", "");
					break;
				case "2":
					tenChuDADT = "02_HT_" + ((Control)txtMaHSDNHT_DADT).Text + "_QQ" + ((Control)mskTuThang).Text.Replace("/", "").PadLeft(6, '0') + "_" + ((Control)mskDenThang).Text.Replace("/", "").PadLeft(6, '0') + "_" + WinSettings.TTPeriodInfo.VersionXML.Replace(".", "");
					break;
				case "3":
					tenChuDADT = "02_HT_" + ((Control)txtMaHSDNHT_DADT).Text + "_TQ" + ((Control)mskTuThang).Text.Replace("/", "").PadLeft(6, '0') + "_" + ((Control)mskDenThang).Text.Replace("/", "").PadLeft(6, '0') + "_" + WinSettings.TTPeriodInfo.VersionXML.Replace(".", "");
					break;
				case "4":
					tenChuDADT = "02_HT_" + ((Control)txtMaHSDNHT_DADT).Text + "_QT" + ((Control)mskTuThang).Text.Replace("/", "").PadLeft(6, '0') + "_" + ((Control)mskDenThang).Text.Replace("/", "").PadLeft(6, '0') + "_" + WinSettings.TTPeriodInfo.VersionXML.Replace(".", "");
					break;
				case "6":
					tenChuDADT = "02_HT_" + ((Control)txtMaHSDNHT_DADT).Text + "_QT" + ((Control)mskTuNgay).Text.Replace("/", "") + "_" + ((Control)mskDenNgay).Text.Replace("/", "") + "_" + WinSettings.TTPeriodInfo.VersionXML.Replace(".", "");
					break;
				}
			}
			else if (WinSettings.MenuID == "02")
			{
				if (((Control)txtMaHSDNHT_DADT).Text.Equals("0"))
				{
					((Control)txtMaHSDNHT_DADT).Text = "1";
				}
				if (chkTkhaiThang.Checked)
				{
					tenChuDADT = "02_GTGT_" + ((Control)txtMaHSDNHT_DADT).Text + "_" + ((Control)mskMonth).Text + ((Control)mskYear).Text + "_" + WinSettings.TTPeriodInfo.VersionXML.Replace(".", "");
				}
				if (chkTKQuy.Checked)
				{
					tenChuDADT = "02_GTGT_" + ((Control)txtMaHSDNHT_DADT).Text + "_Q" + $"{((Control)cbbQuy).Text}{((Control)mskYear).Text}".PadLeft(6, '0') + "_" + WinSettings.TTPeriodInfo.VersionXML.Replace(".", "");
				}
			}
			else if (WinSettings.MenuID == "D2")
			{
				if (((Control)txtMaHSDNHT_DADT).Text.Equals("0"))
				{
					((Control)txtMaHSDNHT_DADT).Text = "1";
				}
				if (chkTkhaiThang.Checked)
				{
					tenChuDADT = "02_GTGT_TT80_" + ((Control)txtMaHSDNHT_DADT).Text + "_" + ((Control)mskMonth).Text + ((Control)mskYear).Text + "_" + WinSettings.TTPeriodInfo.VersionXML.Replace(".", "");
				}
				if (chkTKQuy.Checked)
				{
					tenChuDADT = "02_GTGT_TT80_" + ((Control)txtMaHSDNHT_DADT).Text + "_Q" + $"{((Control)cbbQuy).Text}{((Control)mskYear).Text}".PadLeft(6, '0') + "_" + WinSettings.TTPeriodInfo.VersionXML.Replace(".", "");
				}
			}
			setTenChuDADT(tenChuDADT);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void txtMaHSDNHT_DADT_KeyPress(object sender, KeyPressEventArgs e)
	{
		try
		{
			e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != '\b';
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void txtMaHSDNHT_DADT_TextChanged(object sender, EventArgs e)
	{
		try
		{
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private bool isToKhaiLanDau(HTKKPeriodInfo periodInfor)
	{
		try
		{
			bool result = false;
			DirectoryInfo directoryInfo = new DirectoryInfo(HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst);
			string text = "Giay_GHNT_ND52";
			if (menuID == "N5")
			{
				text = "Giay_GHNT_ND34";
			}
			if (menuID == "H5")
			{
				text = "Giay_GHNT_ND12";
			}
			if (menuID == "O7")
			{
				text = "Giay_GHNT_ND64";
			}
			if (menuID == "P5")
			{
				text = "Giay_GHNT_ND82";
			}
			if (menuID == "U6")
			{
				text = "Giay_GHNT_ND245";
			}
			FileInfo[] files = directoryInfo.GetFiles(text + "*.xml");
			foreach (FileInfo fileInfo in files)
			{
				string text2 = fileInfo.Name.Replace(text + "_", "");
				string[] array = text2.Split(new char[1] { '_' });
				if (!array[0].Equals(periodInfor.Ngay + periodInfor.Thang + periodInfor.Nam))
				{
					result = true;
					break;
				}
			}
			return result;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return false;
		}
	}

	private DateTime getDateTKLanDau()
	{
		try
		{
			DateTime result = DateTime.MinValue;
			DirectoryInfo directoryInfo = new DirectoryInfo(HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst);
			string text = "Giay_GHNT_ND52";
			if (menuID == "N5")
			{
				text = "Giay_GHNT_ND34";
			}
			if (menuID == "H5")
			{
				text = "Giay_GHNT_ND12";
			}
			if (menuID == "O7")
			{
				text = "Giay_GHNT_ND64";
			}
			if (menuID == "P5")
			{
				text = "Giay_GHNT_ND82";
			}
			if (menuID == "U6")
			{
				text = "Giay_GHNT_ND245";
			}
			FileInfo[] files = directoryInfo.GetFiles(text + "*.xml");
			foreach (FileInfo fileInfo in files)
			{
				string text2 = fileInfo.Name.Replace(text + "_", "");
				string[] array = text2.Split(new char[1] { '_' });
				string text3 = array[0];
				if (text3.Length == 8)
				{
					result = new DateTime(ConvertHelpers.ConvertToInt(text3.Substring(4, 4)), ConvertHelpers.ConvertToInt(text3.Substring(2, 2)), ConvertHelpers.ConvertToInt(text3.Substring(0, 2)));
				}
			}
			return result;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return DateTime.MinValue;
		}
	}

	private void getCQTCapCuc()
	{
		_maCQTCapCuc = WinSettings.HTKKHeaderInfo.maCQTCapCuc;
		_tenCQTCapCuc = WinSettings.HTKKHeaderInfo.tenCQTCapCuc;
		_maCQTNoiNop = WinSettings.HTKKHeaderInfo.maCQTNoiNop;
		_tenCQTNoiNop = WinSettings.HTKKHeaderInfo.tenCQTNoiNop;
		Dictionary<string, string> dictionary = new Dictionary<string, string>();
		foreach (KeyValuePair<string, string> item in HTKKWinDictionaryDiaDanh.CoQuanThueCapCuc)
		{
			HTKKFile.AddOrUpdateDicStr(dictionary, item.Key, item.Value);
		}
		if (cboCQTCuc.DataSource != null)
		{
			return;
		}
		Dictionary<string, string> dictionary2 = new Dictionary<string, string>(dictionary);
		dataCboCQTCapCuc = dictionary2.ToList();
		List<string> list = new List<string>();
		if (((Control)cboTinh).Visible)
		{
			list.Add("Chọn chi cục thuế khu vực");
		}
		list.AddRange(dictionary2.Keys.ToList());
		cboCQTCuc.DataSource = list;
		if (WinSettings.ToKhaiInfo.NopTheoTinhCQT == "2")
		{
			int num = dataCboCQTCapCuc.FindIndex((KeyValuePair<string, string> x) => x.Value == WinSettings.HTKKHeaderInfo.maCQTCapCuc);
			if (num >= 0)
			{
				((ListControl)cboCQTCuc).SelectedIndex = num;
			}
		}
	}

	private void cboCQTCuc_SelectedIndexChanged(object sender, EventArgs e)
	{
		string tenCQTCuc = ((Control)cboCQTCuc).Text.Trim();
		if (WinSettings.ToKhaiInfo.NopTheoTinhCQT == "3")
		{
			if (tenCQTCuc == "Chọn chi cục thuế khu vực")
			{
				ResetCQTTinhToEmpty(resetTinh: false, resetXa: false, resetCQTCapCuc: true, resetCQTQuanLy: true);
				return;
			}
			ResetCQTTinhToEmpty(resetTinh: false, resetXa: false, resetCQTCapCuc: false, resetCQTQuanLy: true);
			if (tenCQTCuc == "Chi cục Thuế doanh nghiệp lớn")
			{
				WinSettings.TTPeriodInfo.tenCQTCapCuc = ((Control)cboCQTCuc).Text;
				WinSettings.TTPeriodInfo.maCQTCapCuc = "82500";
			}
			else if (tenCQTCuc == "Chi cục Thuế Thương mại điện tử")
			{
				WinSettings.TTPeriodInfo.tenCQTCapCuc = ((Control)cboCQTCuc).Text;
				WinSettings.TTPeriodInfo.maCQTCapCuc = "82700";
			}
			else
			{
				CQTDiaBan cQTDiaBan = HTKKWinDictionaryDiaDanh._cqtDiaBan.FirstOrDefault((CQTDiaBan x) => x.MaXa == WinSettings.TTPeriodInfo.MaXa);
				if (cQTDiaBan != null)
				{
					string maCQTCapCuc = cQTDiaBan.MaCQT.Substring(0, 3) + "00";
					WinSettings.TTPeriodInfo.tenCQTCapCuc = ((Control)cboCQTCuc).Text;
					WinSettings.TTPeriodInfo.maCQTCapCuc = maCQTCapCuc;
				}
				else
				{
					WinSettings.TTPeriodInfo.tenCQTCapCuc = "";
					WinSettings.TTPeriodInfo.maCQTCapCuc = "";
				}
			}
			cboCQTChiCuc.DataSource = GeneralFunctions.loadCqtCapChiCucTheoDiaBanXa(WinSettings.TTPeriodInfo.maCQTCapCuc, WinSettings.TTPeriodInfo.MaXa).Values.ToList();
			WinSettings.TTinChung.maCQTCapCuc = WinSettings.TTPeriodInfo.maCQTCapCuc;
			WinSettings.TTinChung.tenCQTCapCuc = WinSettings.TTPeriodInfo.tenCQTCapCuc;
			_maCQTCapCuc = WinSettings.HTKKHeaderInfo.maCQTCapCuc;
			return;
		}
		KeyValuePair<string, string> keyValuePair = HTKKWinDictionaryDiaDanh.CoQuanThueCapCuc.FirstOrDefault((KeyValuePair<string, string> x) => x.Key == tenCQTCuc);
		if (!string.IsNullOrEmpty(keyValuePair.Value))
		{
			if (WinSettings.ToKhaiInfo.NopTheoTinhCQT != "3")
			{
				ResetCQTTinhToEmpty(resetTinh: true, resetXa: true, resetCQTCapCuc: true, resetCQTQuanLy: true);
			}
			string value = keyValuePair.Value;
			string key = keyValuePair.Key;
			WinSettings.TTPeriodInfo.tenCQTCapCuc = key;
			WinSettings.TTPeriodInfo.maCQTCapCuc = value;
			WinSettings.HTKKHeaderInfo.maCQTCapCuc = value;
			WinSettings.HTKKHeaderInfo.tenCQTCapCuc = key;
			if (((Control)cboCQTChiCuc).Visible)
			{
				Dictionary<string, string> dictionary = new Dictionary<string, string>(HTKKWinDictionaryDiaDanh.getCoQuanThueQuanLybyMaCoQuanThueCapCuc(value));
				dataCboCQTChiCuc = dictionary.ToList();
				List<string> list = new List<string>();
				if (((Control)cboTinh).Visible)
				{
					list.Add("Chọn cơ quan thuế");
				}
				list.AddRange(dictionary.Values.ToList());
				cboCQTChiCuc.DataSource = list;
			}
			else
			{
				WinSettings.TTPeriodInfo.tenCQTQuanLy = "";
				WinSettings.TTPeriodInfo.maCQTQuanLy = "";
				WinSettings.HTKKHeaderInfo.maCQTNoiNop = _maCQTNoiNop;
				WinSettings.HTKKHeaderInfo.tenCQTNoiNop = _tenCQTNoiNop;
			}
			LoadPhuLucSpread();
		}
		if (cboTinh.Items.Count > 0 && ((Control)cboCQTCuc).Text != "Chọn chi cục thuế khu vực" && WinSettings.ToKhaiInfo.NopTheoTinhCQT != "3")
		{
			((ListControl)cboTinh).SelectedIndex = 0;
		}
	}

	private void cboCQTChiCuc_SelectedIndexChanged(object sender, EventArgs e)
	{
		try
		{
			string tenCQTChiCuc = ((Control)cboCQTChiCuc).Text.Trim();
			if (tenCQTChiCuc == "Chọn cơ quan thuế")
			{
				WinSettings.TTPeriodInfo.maCQTQuanLy = "";
				WinSettings.TTPeriodInfo.tenCQTQuanLy = "";
				return;
			}
			if (tenCQTChiCuc == "Chi cục Thuế doanh nghiệp lớn")
			{
				WinSettings.TTPeriodInfo.tenCQTCapCuc = ((Control)cboCQTCuc).Text;
				WinSettings.TTPeriodInfo.maCQTCapCuc = "82500";
				WinSettings.TTPeriodInfo.maCQTQuanLy = "82500";
				WinSettings.TTPeriodInfo.tenCQTQuanLy = ((Control)cboCQTCuc).Text;
				LoadPhuLucSpread();
			}
			else if (tenCQTChiCuc == "Chi cục Thuế Thương mại điện tử")
			{
				WinSettings.TTPeriodInfo.tenCQTCapCuc = ((Control)cboCQTCuc).Text;
				WinSettings.TTPeriodInfo.maCQTCapCuc = "82700";
				WinSettings.TTPeriodInfo.maCQTQuanLy = "82700";
				WinSettings.TTPeriodInfo.tenCQTQuanLy = ((Control)cboCQTCuc).Text;
				LoadPhuLucSpread();
			}
			else
			{
				CQTDiaBan cQTDiaBan = HTKKWinDictionaryDiaDanh._cqtDiaBan.FirstOrDefault((CQTDiaBan x) => x.MaXa == WinSettings.TTPeriodInfo.MaXa && x.TenCQT == tenCQTChiCuc);
				if (cQTDiaBan != null && !string.IsNullOrEmpty(cQTDiaBan.MaCQT))
				{
					WinSettings.TTPeriodInfo.maCQTQuanLy = cQTDiaBan.MaCQT;
					WinSettings.TTPeriodInfo.tenCQTQuanLy = cQTDiaBan.TenCQT;
					LoadPhuLucSpread();
				}
				else
				{
					CQTDiaBan cQTDiaBan2 = HTKKWinDictionaryDiaDanh._cqtDiaBan.FirstOrDefault((CQTDiaBan x) => x.MaCQT == WinSettings.TTPeriodInfo.maCQTCapCuc && x.TenCQT == tenCQTChiCuc);
					if (cQTDiaBan2 != null && !string.IsNullOrEmpty(cQTDiaBan2.MaCQT))
					{
						WinSettings.TTPeriodInfo.maCQTQuanLy = cQTDiaBan2.MaCQT;
						WinSettings.TTPeriodInfo.tenCQTQuanLy = cQTDiaBan2.TenCQT;
						LoadPhuLucSpread();
					}
					else
					{
						WinSettings.TTPeriodInfo.maCQTQuanLy = "";
						WinSettings.TTPeriodInfo.tenCQTQuanLy = "";
					}
				}
			}
			WinSettings.HTKKHeaderInfo.maCQTNoiNop = WinSettings.TTPeriodInfo.maCQTQuanLy;
			WinSettings.HTKKHeaderInfo.tenCQTNoiNop = WinSettings.TTPeriodInfo.tenCQTQuanLy;
		}
		catch (Exception)
		{
		}
	}

	private void cboTinh_SelectedIndexChanged(object sender, EventArgs e)
	{
		try
		{
			string tenTinh = ((Control)cboTinh).Text.Trim();
			if (tenTinh == "Chọn tỉnh")
			{
				ResetCQTTinhToEmpty(resetTinh: true, resetXa: true, resetCQTCapCuc: true, resetCQTQuanLy: true);
				return;
			}
			KeyValuePair<string, string> keyValuePair = HTKKWinDictionaryDiaDanh.dicTinhTP.FirstOrDefault((KeyValuePair<string, string> x) => x.Key == tenTinh);
			if (WinSettings.ToKhaiInfo.NopTheoTinhCQT == "3")
			{
				keyValuePair = HTKKWinDictionaryDiaDanh.dicTinhTPMoi.FirstOrDefault((KeyValuePair<string, string> x) => x.Key == tenTinh);
			}
			if (!string.IsNullOrEmpty(keyValuePair.Value))
			{
				WinSettings.TTPeriodInfo.MaTinh = keyValuePair.Value;
				WinSettings.TTPeriodInfo.TenTinh = keyValuePair.Key;
				ResetCQTTinhToEmpty(resetTinh: false, resetXa: true, resetCQTCapCuc: true, resetCQTQuanLy: true);
				if (WinSettings.ToKhaiInfo.NopTheoTinhCQT == "3")
				{
					if (((Control)cboXa).Visible)
					{
						string value = keyValuePair.Value;
						string[] danhMucPhuongXaMoi = HTKKWinDictionaryDiaDanh.getDanhMucPhuongXaMoi(value);
						cboXa.DataSource = danhMucPhuongXaMoi;
					}
				}
				else
				{
					ResetCQTTinhToEmpty(resetTinh: true, resetXa: true, resetCQTCapCuc: true, resetCQTQuanLy: true);
					string maCQTCapCuc = keyValuePair.Value + "00";
					if (HTKKWinDictionaryDiaDanh.CoQuanThueCapCuc.ContainsValue(maCQTCapCuc))
					{
						KeyValuePair<string, string> keyValuePair2 = HTKKWinDictionaryDiaDanh.CoQuanThueCapCuc.FirstOrDefault((KeyValuePair<string, string> x) => x.Value == maCQTCapCuc);
						if (!string.IsNullOrEmpty(keyValuePair2.Value))
						{
							WinSettings.HTKKHeaderInfo.maCQTCapCuc = maCQTCapCuc;
							WinSettings.HTKKHeaderInfo.tenCQTCapCuc = keyValuePair2.Key;
							WinSettings.TTinChung.maCQTCapCuc = maCQTCapCuc;
							WinSettings.TTinChung.tenCQTCapCuc = keyValuePair2.Key;
							if ("D1,N8,G9,H2,H4,T8,E8,L8,L7,N2,G2,D2,Q5,Q6,U6,Y6,Z8,Z5".Contains(WinSettings.MenuID))
							{
								int selectedIndex = cboCQTCuc.Items.IndexOf((object)keyValuePair2.Key);
								((ListControl)cboCQTCuc).SelectedIndex = selectedIndex;
							}
						}
					}
					string text = keyValuePair.Value + "00";
					if (HTKKWinDictionaryDiaDanh.CoQuanThueQuanLy.ContainsKey(text))
					{
						WinSettings.HTKKHeaderInfo.maCQTNoiNop = text;
						WinSettings.HTKKHeaderInfo.tenCQTNoiNop = HTKKWinDictionaryDiaDanh.CoQuanThueQuanLy[text];
						WinSettings.TTinChung.maCQTNoiNop = text;
						WinSettings.TTinChung.tenCQTNoiNop = HTKKWinDictionaryDiaDanh.CoQuanThueQuanLy[text];
					}
				}
				LoadPhuLucSpread();
			}
			if (WinSettings.ToKhaiInfo.NopTheoTinhCQT != "3")
			{
				if (cboCQTCuc.Items.Count > 0)
				{
					((ListControl)cboCQTCuc).SelectedIndex = 0;
				}
				if (cboCQTChiCuc.Items.Count > 0)
				{
					((ListControl)cboCQTChiCuc).SelectedIndex = 0;
				}
			}
		}
		catch
		{
		}
	}

	public void ResetCQTTinhToEmpty(bool resetTinh, bool resetXa, bool resetCQTCapCuc, bool resetCQTQuanLy)
	{
		if (resetTinh)
		{
			WinSettings.TTPeriodInfo.MaTinh = "";
			WinSettings.TTPeriodInfo.TenTinh = "";
			WinSettings.TTPeriodInfo.MaXa = "";
			WinSettings.TTPeriodInfo.TenXa = "";
			WinSettings.TTPeriodInfo.maCQTCapCuc = "";
			WinSettings.TTPeriodInfo.tenCQTCapCuc = "";
			WinSettings.TTPeriodInfo.maCQTQuanLy = "";
			WinSettings.TTPeriodInfo.tenCQTQuanLy = "";
			WinSettings.TTinChung.maCQTCapCuc = "";
			WinSettings.TTinChung.tenCQTCapCuc = "";
			WinSettings.TTinChung.maCQTNoiNop = "";
			WinSettings.TTinChung.tenCQTNoiNop = "";
			cboTinh.DataSource = new string[1] { "Chọn tỉnh" };
			cboXa.DataSource = new string[1] { "Chọn xã/phường/đặc khu" };
			cboCQTCuc.DataSource = new string[1] { "Chọn chi cục thuế khu vực" };
			cboCQTChiCuc.DataSource = new string[1] { "Chọn cơ quan thuế" };
		}
		if (resetXa)
		{
			WinSettings.TTPeriodInfo.MaXa = "";
			WinSettings.TTPeriodInfo.TenXa = "";
			WinSettings.TTPeriodInfo.maCQTCapCuc = "";
			WinSettings.TTPeriodInfo.tenCQTCapCuc = "";
			WinSettings.TTPeriodInfo.maCQTQuanLy = "";
			WinSettings.TTPeriodInfo.tenCQTQuanLy = "";
			WinSettings.TTinChung.maCQTCapCuc = "";
			WinSettings.TTinChung.tenCQTCapCuc = "";
			WinSettings.TTinChung.maCQTNoiNop = "";
			WinSettings.TTinChung.tenCQTNoiNop = "";
			cboXa.DataSource = new string[1] { "Chọn xã/phường/đặc khu" };
			cboCQTCuc.DataSource = new string[1] { "Chọn chi cục thuế khu vực" };
			cboCQTChiCuc.DataSource = new string[1] { "Chọn cơ quan thuế" };
		}
		if (resetCQTCapCuc)
		{
			WinSettings.TTPeriodInfo.maCQTCapCuc = "";
			WinSettings.TTPeriodInfo.tenCQTCapCuc = "";
			WinSettings.TTPeriodInfo.maCQTQuanLy = "";
			WinSettings.TTPeriodInfo.tenCQTQuanLy = "";
			WinSettings.TTinChung.maCQTCapCuc = "";
			WinSettings.TTinChung.tenCQTCapCuc = "";
			WinSettings.TTinChung.maCQTNoiNop = "";
			WinSettings.TTinChung.tenCQTNoiNop = "";
			cboCQTCuc.DataSource = new string[1] { "Chọn chi cục thuế khu vực" };
			cboCQTChiCuc.DataSource = new string[1] { "Chọn cơ quan thuế" };
		}
		if (resetCQTQuanLy)
		{
			WinSettings.TTPeriodInfo.maCQTQuanLy = "";
			WinSettings.TTPeriodInfo.tenCQTQuanLy = "";
			WinSettings.TTinChung.maCQTNoiNop = "";
			WinSettings.TTinChung.tenCQTNoiNop = "";
			cboCQTChiCuc.DataSource = new string[1] { "Chọn cơ quan thuế" };
		}
	}

	private void defaultCQT()
	{
		try
		{
			int num;
			if (!string.IsNullOrEmpty(WinSettings.HTKKHeaderInfo.tenCQTCapCuc) && dataCboCQTCapCuc != null)
			{
				num = dataCboCQTCapCuc.FindIndex((KeyValuePair<string, string> x) => x.Value == _maCQTCapCuc);
				if (WinSettings.ToKhaiInfo.NopTheoTinhCQT == "2")
				{
					if (num >= 0)
					{
						((ListControl)cboCQTCuc).SelectedIndex = num;
					}
				}
				else if (num >= 0)
				{
					((ListControl)cboCQTCuc).SelectedIndex = num + 1;
				}
			}
			if (string.IsNullOrEmpty(WinSettings.HTKKHeaderInfo.maCQTNoiNop) || dataCboCQTChiCuc == null)
			{
				return;
			}
			num = dataCboCQTChiCuc.FindIndex((KeyValuePair<string, string> x) => x.Key == _maCQTNoiNop);
			if (WinSettings.ToKhaiInfo.NopTheoTinhCQT == "2")
			{
				if (num >= 0)
				{
					((ListControl)cboCQTChiCuc).SelectedIndex = num;
				}
			}
			else if (num >= 0)
			{
				((ListControl)cboCQTChiCuc).SelectedIndex = num + 1;
			}
		}
		catch (Exception)
		{
		}
	}

	private void chkMSTCaNhanToChuc_CheckedChanged(object sender, EventArgs e)
	{
		//IL_0002: Unknown result type (might be due to invalid IL or missing references)
		//IL_003c: Unknown result type (might be due to invalid IL or missing references)
		if (((Control)(CheckBox)sender).Name == "chkMSTCaNhan" && chkMSTCaNhan.Checked)
		{
			chkMSTToChuc.Checked = false;
		}
		if (((Control)(CheckBox)sender).Name == "chkMSTToChuc" && chkMSTToChuc.Checked)
		{
			chkMSTCaNhan.Checked = false;
		}
	}

	private void NopTheoTinhCQT()
	{
		WinSettings.ToKhaiInfo.KieuKHBS = ((WinSettings.xmlMenuNode.Attributes["KieuKHBS"] != null) ? ((XmlNode)WinSettings.xmlMenuNode.Attributes["KieuKHBS"]).Value : "");
		if (WinSettings.xmlMenuNode.Attributes["NopTheoTinhCQT"] == null)
		{
			WinSettings.ToKhaiInfo.NopTheoTinhCQT = "";
			return;
		}
		string text = ((XmlNode)WinSettings.xmlMenuNode.Attributes["NopTheoTinhCQT"]).Value.ToString();
		switch (text)
		{
		case "1":
			WinSettings.ToKhaiInfo.NopTheoTinhCQT = "1";
			ShowDetailpnTinhCQT(cboTinhShow: true, cboXaShow: false, cboCQTCucShow: true, cboCQTChiCucShow: true, text);
			break;
		case "2":
			WinSettings.ToKhaiInfo.NopTheoTinhCQT = "2";
			ShowDetailpnTinhCQT(cboTinhShow: false, cboXaShow: false, cboCQTCucShow: true, cboCQTChiCucShow: true, text);
			break;
		case "3":
			WinSettings.ToKhaiInfo.NopTheoTinhCQT = "3";
			ShowDetailpnTinhCQT(cboTinhShow: true, cboXaShow: true, cboCQTCucShow: true, cboCQTChiCucShow: true, text);
			break;
		}
	}

	private void chkToChucCaNhanDoanhNghiepChuyenNhuong_CheckedChanged()
	{
		if (chkToChucCaNhanChuyenNhuong.Checked)
		{
			chkDoanhNghiepChuyenNhuong.Checked = false;
		}
		if (chkDoanhNghiepChuyenNhuong.Checked)
		{
			chkToChucCaNhanChuyenNhuong.Checked = false;
		}
	}

	private void chkThuHo_CheckedChanged(object sender, EventArgs e)
	{
		try
		{
			SheetView sheetView = axfpsPhuLuc.Sheets[0];
			sheetView.Cells[0, 0].Value = false;
			sheetView.Cells[0, 0].Locked = true;
			chkChonPhuLucKeKhaiAll.Checked = false;
			if (menuID == "71")
			{
				if (chkThuHo.Checked)
				{
					return;
				}
				if (chkTKLanPS.Checked || chkTkhaiThang.Checked)
				{
					if (((Control)mskYear).Text.Equals("2021") && ConvertHelpers.ConvertToInt(((Control)mskMonth).Text) >= 11 && ConvertHelpers.ConvertToInt(((Control)mskMonth).Text) <= 12)
					{
						sheetView.Cells[0, 0].Value = false;
						sheetView.Cells[0, 0].Locked = false;
						chkChonPhuLucKeKhaiAll.Checked = false;
					}
				}
				else if (chkTKQuy.Checked && ((Control)mskYear).Text.Equals("2021") && ConvertHelpers.ConvertToInt(((Control)cbbQuy).Text) == 4)
				{
					sheetView.Cells[0, 0].Value = false;
					sheetView.Cells[0, 0].Locked = false;
					chkChonPhuLucKeKhaiAll.Checked = false;
				}
			}
			else
			{
				if (!(menuID == "D4") && !WinSettings.MenuID.Equals("Z3"))
				{
					return;
				}
				if (chkTKLanPS.Checked || chkTkhaiThang.Checked)
				{
					if (((Control)mskYear).Text.Equals("2022") && ConvertHelpers.ConvertToInt(((Control)mskMonth).Text) >= 2 && ConvertHelpers.ConvertToInt(((Control)mskMonth).Text) <= 12)
					{
						sheetView.Cells[0, 0].Value = false;
						sheetView.Cells[0, 0].Locked = false;
						chkChonPhuLucKeKhaiAll.Checked = false;
					}
				}
				else if (chkTKQuy.Checked && ((Control)mskYear).Text.Equals("2022"))
				{
					sheetView.Cells[0, 0].Value = false;
					sheetView.Cells[0, 0].Locked = false;
					chkChonPhuLucKeKhaiAll.Checked = false;
				}
				LoadPhuLucSpread();
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private bool checkCqtNoiNop()
	{
		//IL_015a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0094: Unknown result type (might be due to invalid IL or missing references)
		//IL_00ce: Unknown result type (might be due to invalid IL or missing references)
		//IL_01d0: Unknown result type (might be due to invalid IL or missing references)
		//IL_0260: Unknown result type (might be due to invalid IL or missing references)
		try
		{
			bool result = true;
			if (WinSettings.ToKhaiInfo.NopTheoTinhCQT == "1")
			{
				if ((string.IsNullOrEmpty(WinSettings.TTPeriodInfo.maCQTQuanLy) && string.IsNullOrEmpty(WinSettings.TTPeriodInfo.MaTinh)) || (string.IsNullOrEmpty(WinSettings.TTPeriodInfo.tenCQTQuanLy) && string.IsNullOrEmpty(WinSettings.TTPeriodInfo.TenTinh)))
				{
					frmMessage frmMessage = new frmMessage("0554", "3", "1", "Thông báo");
					((Form)frmMessage).ShowDialog();
					result = false;
				}
				if (!string.IsNullOrEmpty(WinSettings.TTPeriodInfo.MaTinh))
				{
					frmMessage frmMessage = new frmMessage("0590", "3", "1", "Thông báo");
					((Form)frmMessage).ShowDialog();
				}
				if (!(menuID == "D4"))
				{
				}
			}
			if (WinSettings.ToKhaiInfo.NopTheoTinhCQT == "2")
			{
				if (string.IsNullOrEmpty(WinSettings.TTPeriodInfo.maCQTQuanLy) || string.IsNullOrEmpty(WinSettings.TTPeriodInfo.tenCQTQuanLy))
				{
					frmMessage frmMessage = new frmMessage("0536", "3", "1", "Thông báo");
					((Form)frmMessage).ShowDialog();
					result = false;
				}
				if (menuID == "G5" && WinSettings.TTPeriodInfo.KieuKy == "Y" && WinSettings.TTPeriodInfo.maCQTQuanLy != _maCQTNoiNop)
				{
					frmMessage frmMessage = new frmMessage("0567", "3", "1", "Thông báo");
					((Form)frmMessage).ShowDialog();
					result = false;
				}
				if (menuID == "D2" && WinSettings.TTPeriodInfo.maCQTQuanLy.Length == 5 && WinSettings.TTPeriodInfo.maCQTQuanLy.Substring(3, 2) != "00" && WinSettings.TTPeriodInfo.maCQTQuanLy != _maCQTNoiNop)
				{
					frmMessage frmMessage = new frmMessage("0598", "3", "1", "Thông báo");
					((Form)frmMessage).ShowDialog();
					result = false;
				}
			}
			return result;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return false;
		}
	}

	private void convertDL_MGT_ND92()
	{
		try
		{
			string text = "";
			string text2 = "";
			string text3 = "";
			text = "01_GTGT~01_5_GTGT~01_6_GTGT~01_7_GTGT~01_1_TD_GTGT~01_2_TD_GTGT";
			text2 = "240";
			text3 = "250";
			WinFnProcess.ConverDataVersion(text, text2, text3);
			text = "01_GTGT~01_5_GTGT~01_6_GTGT~01_7_GTGT~01_1_TD_GTGT~01_2_TD_GTGT~KHBS_01_GTGT";
			text2 = "240";
			text3 = "250";
			WinFnProcess.ConverDataVersionKHBS(text, text2, text3);
			text = "04_GTGT";
			text2 = "241";
			text3 = "250";
			WinFnProcess.ConverDataVersion(text, text2, text3);
			text = "04_GTGT~KHBS_04_GTGT";
			text2 = "241";
			text3 = "250";
			WinFnProcess.ConverDataVersionKHBS(text, text2, text3);
			text = "01_CNKD~01_1_CNKD";
			text2 = "216";
			text3 = "250";
			WinFnProcess.ConverDataVersion(text, text2, text3);
			WinFnProcess.ConverDataVersionKHBS(text, text2, text3);
			text = "01BC_SDHD_CNKD";
			text2 = "216";
			text3 = "250";
			WinFnProcess.ConverDataVersion(text, text2, text3);
			WinFnProcess.ConverDataVersionKHBS(text, text2, text3);
			text = "01_XSBHDC~01_1_XSBHDC";
			text2 = "216";
			text3 = "250";
			WinFnProcess.ConverDataVersion(text, text2, text3);
			WinFnProcess.ConverDataVersionKHBS(text, text2, text3);
			text = "01_TTS~01_1_TTS";
			text2 = "216";
			text3 = "250";
			WinFnProcess.ConverDataVersion(text, text2, text3);
			WinFnProcess.ConverDataVersionKHBS(text, text2, text3);
			text = "01_TKN_BHXSDC";
			text2 = "216";
			text3 = "250";
			WinFnProcess.ConverDataVersion(text, text2, text3);
			WinFnProcess.ConverDataVersionKHBS(text, text2, text3);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void cboCQTCuc_KeyPress(object sender, KeyPressEventArgs e)
	{
		e.Handled = true;
	}

	private void cboTinh_KeyPress(object sender, KeyPressEventArgs e)
	{
		e.Handled = true;
	}

	private void cboCQTChiCuc_KeyPress(object sender, KeyPressEventArgs e)
	{
		e.Handled = true;
	}

	private void cboKieuKy_SelectedIndexChanged(object sender, EventArgs e)
	{
		try
		{
			string text = ((ListControl)cboKieuKy).SelectedValue.ToString();
			switch (text)
			{
			case "1":
				((Control)mskTuThang).Text = "";
				((Control)mskDenThang).Text = "";
				((Control)lblTuThang).Text = "Từ tháng";
				((Control)lblDenThang).Text = "Đến tháng";
				mskTuThang.Mask = "00/0000";
				mskDenThang.Mask = "00/0000";
				WinSettings.TTPeriodInfo.KieuKyHoanThue = "1";
				WinSettings.TTPeriodInfo.KieuKy = "M";
				break;
			case "2":
				((Control)lblTuThang).Text = "Từ quý";
				((Control)lblDenThang).Text = "Đến quý";
				mskTuThang.Mask = "0/0000";
				mskDenThang.Mask = "0/0000";
				WinSettings.TTPeriodInfo.KieuKyHoanThue = "2";
				WinSettings.TTPeriodInfo.KieuKy = "Q";
				break;
			case "3":
				((Control)lblTuThang).Text = "Từ tháng";
				((Control)lblDenThang).Text = "Đến quý";
				mskTuThang.Mask = "00/0000";
				mskDenThang.Mask = "0/0000";
				WinSettings.TTPeriodInfo.KieuKyHoanThue = "3";
				WinSettings.TTPeriodInfo.KieuKy = "Q";
				break;
			case "4":
				((Control)lblTuThang).Text = "Từ quý";
				((Control)lblDenThang).Text = "Đến tháng";
				mskTuThang.Mask = "0/0000";
				mskDenThang.Mask = "00/0000";
				WinSettings.TTPeriodInfo.KieuKyHoanThue = "4";
				WinSettings.TTPeriodInfo.KieuKy = "M";
				break;
			}
			if (text == "5")
			{
				((Control)mskTuThang).Text = "";
				((Control)mskDenThang).Text = "";
				((Control)lblTuThang).Text = "Từ năm";
				((Control)lblDenThang).Text = "Đến năm";
				mskTuThang.Mask = "0000";
				mskDenThang.Mask = "0000";
				WinSettings.TTPeriodInfo.KieuKyHoanThue = "5";
				WinSettings.TTPeriodInfo.KieuKy = "Y";
				ShowDetailpnVT1_3TuNgayDenNgay(pnLanXuatBanShow: false, pnTuNgayShow: false, pnDenNgayShow: false, pnTuThangShow: true, pnDenThangShow: true);
				Resize();
			}
			if (text == "6")
			{
				((Control)mskTuNgay).Text = "";
				((Control)mskDenNgay).Text = "";
				WinSettings.TTPeriodInfo.KieuKyHoanThue = "6";
				WinSettings.TTPeriodInfo.KieuKy = "D_D";
				ShowDetailpnVT1_3TuNgayDenNgay(pnLanXuatBanShow: false, pnTuNgayShow: true, pnDenNgayShow: true, pnTuThangShow: false, pnDenThangShow: false);
				Resize();
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void cboKieuKy_KeyPress(object sender, KeyPressEventArgs e)
	{
		e.Handled = true;
	}

	private void cboTruongHopQT_KeyPress(object sender, KeyPressEventArgs e)
	{
		e.Handled = true;
	}

	private void txtLanXuatBan_KeyPress(object sender, KeyPressEventArgs e)
	{
		if (menuID == "M4" || menuID == "G3" || menuID == "P1" || menuID == "P6" || menuID == "X7" || menuID.Equals("Y9"))
		{
			e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != '\b';
		}
	}

	private void txtLanXuatBan_TextChanged(object sender, EventArgs e)
	{
		if ((menuID == "M4" || menuID == "G3" || menuID == "P1" || menuID == "P6" || menuID == "X7" || menuID.Equals("Y9")) && ((Control)txtLanXuatBan).Text.Equals("0"))
		{
			((Control)txtLanXuatBan).Text = "1";
		}
	}

	private void TruongHopQuyetToan()
	{
		WinSettings.TTPeriodInfo.TenQuyetToan = "";
		WinSettings.TTPeriodInfo.MaQuyetToan = "";
		string text = ((WinSettings.xmlMenuNode.Attributes["TruongHopQuyetToan"] != null) ? ((XmlNode)WinSettings.xmlMenuNode.Attributes["TruongHopQuyetToan"]).Value.Trim() : "");
		if (text == "1")
		{
			((Control)pnTruongHopQT).Visible = true;
			SetDataTruongHopQuyetToan();
		}
		else
		{
			((Control)pnTruongHopQT).Visible = false;
		}
		((Control)pnTruongHopQT).Location = new Point(((Control)pnTinhCQT).Location.X, ((Control)pnTinhCQT).Location.Y + 1);
		Resize();
	}

	private void SetDataTruongHopQuyetToan()
	{
		//IL_0025: Unknown result type (might be due to invalid IL or missing references)
		//IL_002f: Expected O, but got Unknown
		Dictionary<string, string> dataTruongHopQuyetToan = HTKKWinDictionaryDiaDanh.GetDataTruongHopQuyetToan(WinSettings.MenuID);
		if (dataTruongHopQuyetToan.Count > 0)
		{
			cboTruongHopQT.DataSource = (object)new BindingSource((object)dataTruongHopQuyetToan, (string)null);
			((ListControl)cboTruongHopQT).DisplayMember = "Value";
			((ListControl)cboTruongHopQT).ValueMember = "Key";
			((Control)cboTruongHopQT).Refresh();
		}
	}

	private void cboTruongHopQT_SelectedIndexChanged(object sender, EventArgs e)
	{
		try
		{
			string text = ((Control)cboTruongHopQT).Text.Trim();
			string text2 = ((ListControl)cboTruongHopQT).SelectedValue.ToString().Trim();
			if (!string.IsNullOrEmpty(text2) && !string.IsNullOrEmpty(text) && text2 != text)
			{
				string maQuyetToan = text2.Replace(text, "").Replace("[", "").Replace("]", "")
					.Replace(" ", "")
					.Replace(",", "");
				WinSettings.TTPeriodInfo.TenQuyetToan = text;
				WinSettings.TTPeriodInfo.MaQuyetToan = maQuyetToan;
			}
			CheckDongDenThangTheoQuyetToan();
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private bool CheckTruongHopQuyetToan()
	{
		//IL_0238: Unknown result type (might be due to invalid IL or missing references)
		//IL_046c: Unknown result type (might be due to invalid IL or missing references)
		//IL_064e: Unknown result type (might be due to invalid IL or missing references)
		//IL_01c7: Unknown result type (might be due to invalid IL or missing references)
		//IL_07f8: Unknown result type (might be due to invalid IL or missing references)
		//IL_03fb: Unknown result type (might be due to invalid IL or missing references)
		//IL_0b46: Unknown result type (might be due to invalid IL or missing references)
		//IL_09c9: Unknown result type (might be due to invalid IL or missing references)
		//IL_0bf2: Unknown result type (might be due to invalid IL or missing references)
		//IL_0ce1: Unknown result type (might be due to invalid IL or missing references)
		//IL_0af5: Unknown result type (might be due to invalid IL or missing references)
		try
		{
			if (!string.IsNullOrEmpty(((Control)mskTuThang).Text.Trim()) && !string.IsNullOrEmpty(((Control)mskDenThang).Text.Trim()) && !string.IsNullOrEmpty(WinSettings.TTPeriodInfo.MaQuyetToan) && ((Control)mskTuThang).Visible && ((Control)mskDenThang).Visible)
			{
				DateTime dateTime = ConvertHelpers.ConvertToDateTime("01/" + ((Control)mskTuThang).Text.Trim(), "dd/MM/yyyy");
				DateTime dateTime2 = ConvertHelpers.ConvertToDateTime("01/" + ((Control)mskDenThang).Text.Trim(), "dd/MM/yyyy");
				XmlNode nodeTruongHopQuyetToan = GetNodeTruongHopQuyetToan(WinSettings.MenuID, WinSettings.TTPeriodInfo.MaQuyetToan);
				string text = ((nodeTruongHopQuyetToan.Attributes["NamTaiChinh"] == null) ? "0" : ((XmlNode)nodeTruongHopQuyetToan.Attributes["NamTaiChinh"]).Value.Trim());
				if (text == "0")
				{
					if (nodeTruongHopQuyetToan != null && nodeTruongHopQuyetToan.Attributes["TuDenPhaiThuocNamQT"] != null)
					{
						string text2 = ((XmlNode)nodeTruongHopQuyetToan.Attributes["TuDenPhaiThuocNamQT"]).Value.Trim();
						if (text2 == "1" && (dateTime.Year != Convert.ToInt32(((Control)mskYear).Text) || dateTime2.Year != Convert.ToInt32(((Control)mskYear).Text)))
						{
							frmMessage frmMessage = new frmMessage("0617", "3", "", "Thông báo");
							((Form)frmMessage).ShowDialog();
							return false;
						}
					}
				}
				else if (dateTime.Year > Convert.ToInt32(((Control)mskYear).Text) || dateTime2.Year < Convert.ToInt32(((Control)mskYear).Text))
				{
					frmMessage frmMessage = new frmMessage("0617", "3", "", "Thông báo");
					((Form)frmMessage).ShowDialog();
					return false;
				}
			}
			if (!string.IsNullOrEmpty(((Control)mskTuNgay).Text.Trim()) && !string.IsNullOrEmpty(((Control)mskDenNgay).Text.Trim()) && !string.IsNullOrEmpty(WinSettings.TTPeriodInfo.MaQuyetToan) && ((Control)mskTuNgay).Visible && ((Control)mskDenNgay).Visible)
			{
				DateTime dateTime3 = ConvertHelpers.ConvertToDateTime(((Control)mskTuNgay).Text.Trim(), "dd/MM/yyyy");
				DateTime dateTime4 = ConvertHelpers.ConvertToDateTime(((Control)mskDenNgay).Text.Trim(), "dd/MM/yyyy");
				XmlNode nodeTruongHopQuyetToan = GetNodeTruongHopQuyetToan(WinSettings.MenuID, WinSettings.TTPeriodInfo.MaQuyetToan);
				string text = ((nodeTruongHopQuyetToan.Attributes["NamTaiChinh"] == null) ? "0" : ((XmlNode)nodeTruongHopQuyetToan.Attributes["NamTaiChinh"]).Value.Trim());
				if (text == "0")
				{
					if (nodeTruongHopQuyetToan != null && nodeTruongHopQuyetToan.Attributes["TuDenPhaiThuocNamQT"] != null)
					{
						string text2 = ((XmlNode)nodeTruongHopQuyetToan.Attributes["TuDenPhaiThuocNamQT"]).Value.Trim();
						if (text2 == "1" && (dateTime3.Year != Convert.ToInt32(((Control)mskYear).Text) || dateTime4.Year != Convert.ToInt32(((Control)mskYear).Text)))
						{
							frmMessage frmMessage = new frmMessage("0617", "3", "", "Thông báo");
							((Form)frmMessage).ShowDialog();
							return false;
						}
					}
				}
				else if (dateTime3.Year > Convert.ToInt32(((Control)mskYear).Text) || dateTime4.Year < Convert.ToInt32(((Control)mskYear).Text))
				{
					frmMessage frmMessage = new frmMessage("0617", "3", "", "Thông báo");
					((Form)frmMessage).ShowDialog();
					return false;
				}
			}
			if (!string.IsNullOrEmpty(((Control)mskTuThang).Text.Trim()) && !string.IsNullOrEmpty(((Control)mskDenThang).Text.Trim()) && !string.IsNullOrEmpty(WinSettings.TTPeriodInfo.MaQuyetToan) && ((Control)mskTuThang).Visible && ((Control)mskDenThang).Visible)
			{
				int num = (ConvertHelpers.ConvertToDateTime(((Control)mskDenThang).Text.Trim(), "MM/yyyy").Year - ConvertHelpers.ConvertToDateTime(((Control)mskTuThang).Text.Trim(), "MM/yyyy").Year) * 12 + ConvertHelpers.ConvertToDateTime(((Control)mskDenThang).Text.Trim(), "MM/yyyy").Month - ConvertHelpers.ConvertToDateTime(((Control)mskTuThang).Text.Trim(), "MM/yyyy").Month + 1;
				XmlNode nodeTruongHopQuyetToan = GetNodeTruongHopQuyetToan(WinSettings.MenuID, WinSettings.TTPeriodInfo.MaQuyetToan);
				if (nodeTruongHopQuyetToan != null && nodeTruongHopQuyetToan.Attributes["KhongQuaSoThang"] != null)
				{
					int num2 = Convert.ToInt32(((XmlNode)nodeTruongHopQuyetToan.Attributes["KhongQuaSoThang"]).Value.Trim());
					if (num > num2)
					{
						XMLMessage xMLMessage = new XMLMessage("../Project/Message.xml");
						string message = xMLMessage.GetMessage("0618");
						frmMessage frmMessage = new frmMessage("", message + " " + num2 + " tháng", "3", "", "Thông báo");
						((Form)frmMessage).ShowDialog();
						return false;
					}
				}
			}
			if (!string.IsNullOrEmpty(((Control)mskTuNgay).Text.Trim()) && !string.IsNullOrEmpty(((Control)mskDenNgay).Text.Trim()) && !string.IsNullOrEmpty(WinSettings.TTPeriodInfo.MaQuyetToan) && ((Control)mskTuNgay).Visible && ((Control)mskDenNgay).Visible)
			{
				DateTime dateTime3 = ConvertHelpers.ConvertToDateTime(((Control)mskTuNgay).Text.Trim(), "dd/MM/yyyy");
				DateTime dateTime4 = ConvertHelpers.ConvertToDateTime(((Control)mskDenNgay).Text.Trim(), "dd/MM/yyyy");
				int num = (dateTime4.Year - dateTime3.Year) * 12 + dateTime4.Month - dateTime3.Month + 1;
				XmlNode nodeTruongHopQuyetToan = GetNodeTruongHopQuyetToan(WinSettings.MenuID, WinSettings.TTPeriodInfo.MaQuyetToan);
				if (nodeTruongHopQuyetToan != null && nodeTruongHopQuyetToan.Attributes["KhongQuaSoThang"] != null)
				{
					int num2 = Convert.ToInt32(((XmlNode)nodeTruongHopQuyetToan.Attributes["KhongQuaSoThang"]).Value.Trim());
					if (num > num2)
					{
						XMLMessage xMLMessage = new XMLMessage("../Project/Message.xml");
						string message = xMLMessage.GetMessage("0618");
						frmMessage frmMessage = new frmMessage("", message + " " + num2 + " tháng", "3", "", "Thông báo");
						((Form)frmMessage).ShowDialog();
						return false;
					}
				}
			}
			if (!string.IsNullOrEmpty(((Control)mskTuNgay).Text.Trim()) && !string.IsNullOrEmpty(((Control)mskDenNgay).Text.Trim()) && !string.IsNullOrEmpty(WinSettings.TTPeriodInfo.MaQuyetToan) && ((Control)mskTuNgay).Visible && ((Control)mskDenNgay).Visible)
			{
				DateTime dateTime3 = ConvertHelpers.ConvertToDateTime(((Control)mskTuNgay).Text.Trim(), "dd/MM/yyyy");
				DateTime dateTime4 = ConvertHelpers.ConvertToDateTime(((Control)mskDenNgay).Text.Trim(), "dd/MM/yyyy");
				XmlNode nodeTruongHopQuyetToan = GetNodeTruongHopQuyetToan(WinSettings.MenuID, WinSettings.TTPeriodInfo.MaQuyetToan);
				string text = ((nodeTruongHopQuyetToan.Attributes["NamTaiChinh"] == null) ? "0" : ((XmlNode)nodeTruongHopQuyetToan.Attributes["NamTaiChinh"]).Value.Trim());
				if (menuID != "C9" && WinSettings.MenuID != "X9" && menuID != "D6" && menuID != "Y1" && text == "0")
				{
					if (menuID == "N8")
					{
						if (dateTime3.Year != Convert.ToInt32(((Control)mskYear).Text) && dateTime4.Year != Convert.ToInt32(((Control)mskYear).Text))
						{
							frmMessage frmMessage = new frmMessage("0617", "3", "", "Thông báo");
							((Form)frmMessage).ShowDialog();
							return false;
						}
						string text3 = ((nodeTruongHopQuyetToan.Attributes["MaQT"] == null) ? "" : ((XmlNode)nodeTruongHopQuyetToan.Attributes["MaQT"]).Value.Trim());
						if (text3 == "01")
						{
							string ngayBatDauNamTaiChinh = WinSettings.HTKKHeaderInfo.NgayBatDauNamTaiChinh;
							string[] array = ngayBatDauNamTaiChinh.Split(new char[1] { '/' });
							DateUtils dateUtils = new DateUtils();
							string text4 = dateUtils.ToString(GeneralFunctions.NgayCuoiNamTaiChinh(ConvertHelpers.ConvertToInt(((Control)mskYear).Text), ConvertHelpers.ConvertToInt(array[1]), ConvertHelpers.ConvertToInt(array[0])), "DDMMYYYY");
							string text5 = ((Control)mskDenNgay).Text;
							if (text5.Length >= 5 && text5.Substring(0, 5).Replace("/", "") != text4.Substring(0, 4))
							{
								frmMessage frmMessage = new frmMessage("0617", "3", "", "Thông báo");
								((Form)frmMessage).ShowDialog();
								return false;
							}
						}
					}
					else if (dateTime3.Year != Convert.ToInt32(((Control)mskYear).Text))
					{
						frmMessage frmMessage = new frmMessage("0617", "3", "", "Thông báo");
						((Form)frmMessage).ShowDialog();
						return false;
					}
				}
				if (menuID == "N8" || menuID == "T5" || menuID == "Y4")
				{
					if (dateTime3.Year != Convert.ToInt32(((Control)mskYear).Text) && dateTime4.Year != Convert.ToInt32(((Control)mskYear).Text))
					{
						frmMessage frmMessage = new frmMessage("0617", "3", "", "Thông báo");
						((Form)frmMessage).ShowDialog();
						return false;
					}
					if (WinSettings.TTPeriodInfo.MaQuyetToan == "01")
					{
						string ngayBatDauNamTaiChinh = WinSettings.HTKKHeaderInfo.NgayBatDauNamTaiChinh;
						string[] array = ngayBatDauNamTaiChinh.Split(new char[1] { '/' });
						DateUtils dateUtils = new DateUtils();
						string text4 = dateUtils.ToString(GeneralFunctions.NgayCuoiNamTaiChinh(ConvertHelpers.ConvertToInt(((Control)mskYear).Text), ConvertHelpers.ConvertToInt(array[1]), ConvertHelpers.ConvertToInt(array[0])), "DDMMYYYY");
						string text5 = ((Control)mskDenNgay).Text;
						if (text5.Length >= 5 && text5.Substring(0, 5) != text4.Substring(0, 5))
						{
							frmMessage frmMessage = new frmMessage("0700", "3", "", "Thông báo");
							((Form)frmMessage).ShowDialog();
							return false;
						}
					}
				}
			}
			return true;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return true;
		}
	}

	public void CheckDongDenThangTheoQuyetToan()
	{
		try
		{
			if (string.IsNullOrEmpty(WinSettings.TTPeriodInfo.MaQuyetToan))
			{
				return;
			}
			XmlNode nodeTruongHopQuyetToan = GetNodeTruongHopQuyetToan(WinSettings.MenuID, WinSettings.TTPeriodInfo.MaQuyetToan);
			if (nodeTruongHopQuyetToan == null || nodeTruongHopQuyetToan.Attributes["DongDenThang"] == null)
			{
				return;
			}
			string text = ((XmlNode)nodeTruongHopQuyetToan.Attributes["DongDenThang"]).Value.Trim();
			string text2 = ((XmlNode)nodeTruongHopQuyetToan.Attributes["Tron12ThangLT"]).Value.Trim();
			if (text == "1")
			{
				if (text2 == "0")
				{
					if (((Control)mskDenNgay).Visible)
					{
						if ("N8,C6,C7,T5,H4,P8,P9,Q2,U3,U4,X4,Y4".Contains(menuID))
						{
							string ngayBatDauNamTaiChinh = WinSettings.HTKKHeaderInfo.NgayBatDauNamTaiChinh;
							string[] array = ngayBatDauNamTaiChinh.Split(new char[1] { '/' });
							((Control)mskTuNgay).Text = ngayBatDauNamTaiChinh + "/" + ((Control)mskYear).Text;
							DateUtils dateUtils = new DateUtils();
							((Control)mskDenNgay).Text = dateUtils.ToString(GeneralFunctions.NgayCuoiNamTaiChinh(ConvertHelpers.ConvertToInt(((Control)mskYear).Text), ConvertHelpers.ConvertToInt(array[1]), ConvertHelpers.ConvertToInt(array[0])), "DDMMYYYY");
						}
						else
						{
							((Control)mskDenNgay).Text = "31/12/" + ((Control)mskYear).Text;
						}
					}
					if (((Control)mskDenThang).Visible)
					{
						((Control)mskDenThang).Text = "12/" + ((Control)mskYear).Text;
					}
				}
				if (((Control)mskDenNgay).Visible)
				{
					((Control)mskDenNgay).Enabled = false;
				}
				if (((Control)mskDenThang).Visible)
				{
					((Control)mskDenThang).Enabled = false;
				}
			}
			else
			{
				if (((Control)mskDenNgay).Visible)
				{
					((Control)mskDenNgay).Enabled = true;
				}
				if (((Control)mskDenThang).Visible)
				{
					((Control)mskDenThang).Enabled = true;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void CheckNextDenThangTheoQuyetToan()
	{
		try
		{
			if (string.IsNullOrEmpty(WinSettings.TTPeriodInfo.MaQuyetToan))
			{
				return;
			}
			XmlNode nodeTruongHopQuyetToan = GetNodeTruongHopQuyetToan(WinSettings.MenuID, WinSettings.TTPeriodInfo.MaQuyetToan);
			if (nodeTruongHopQuyetToan == null || nodeTruongHopQuyetToan.Attributes["Tron12ThangLT"] == null)
			{
				return;
			}
			string text = ((XmlNode)nodeTruongHopQuyetToan.Attributes["Tron12ThangLT"]).Value.Trim();
			if (text == "1")
			{
				if (((Control)mskDenNgay).Visible && !string.IsNullOrEmpty(((Control)mskTuNgay).Text))
				{
					DateTime dateTime = DateTime.ParseExact(((Control)mskTuNgay).Text, "dd/MM/yyyy", CultureInfo.CurrentCulture);
					((Control)mskDenNgay).Text = dateTime.AddMonths(11).ToString("dd/MM/yyyy");
				}
				if (((Control)mskDenThang).Visible && !string.IsNullOrEmpty(((Control)mskTuThang).Text))
				{
					DateTime dateTime2 = DateTime.ParseExact(((Control)mskTuThang).Text, "MM/yyyy", CultureInfo.CurrentCulture);
					((Control)mskDenThang).Text = dateTime2.AddMonths(11).ToString("MM/yyyy");
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public XmlNode GetNodeTruongHopQuyetToan(string menuId, string maQuyetToan)
	{
		try
		{
			XmlDataDocument xmlDataDocument = new XmlDataDocument();
			FileStream fileStream = new FileStream(HTKKConfig.INTERFACE_INI + "\\DM_TruongHopQuyetToan.xml", FileMode.Open, FileAccess.Read);
			((XmlDocument)xmlDataDocument).Load((Stream)fileStream);
			XmlNodeList elementsByTagName = ((XmlDocument)xmlDataDocument).GetElementsByTagName("Item");
			return ((IEnumerable)elementsByTagName).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode x) => ((XmlNode)x.Attributes["MenuId"]).Value.Trim() == WinSettings.MenuID && ((XmlNode)x.Attributes["MaQT"]).Value.Trim() == WinSettings.TTPeriodInfo.MaQuyetToan));
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	private void cboNganhKD_KeyDown(object sender, KeyEventArgs e)
	{
		e.Handled = true;
	}

	private void cboTinh_KeyDown(object sender, KeyEventArgs e)
	{
		e.Handled = true;
	}

	private void cboCQTCuc_KeyDown(object sender, KeyEventArgs e)
	{
		e.Handled = true;
	}

	private void cboCQTChiCuc_KeyDown(object sender, KeyEventArgs e)
	{
		e.Handled = true;
	}

	private void cboTruongHopQT_KeyDown(object sender, KeyEventArgs e)
	{
		e.Handled = true;
	}

	private void cboLoaiToKhai_SelectedIndexChanged(object sender, EventArgs e)
	{
		WinSettings.TTPeriodInfo.LoaiToKhai = ((((ListControl)cboLoaiToKhai).SelectedIndex == 0) ? "C" : ((((ListControl)cboLoaiToKhai).SelectedIndex == 1) ? "T" : ""));
		if (((ListControl)cboLoaiToKhai).SelectedIndex != 0)
		{
			((Control)pnLan1).Visible = true;
			((Control)mskNgayLapKHBS1).Text = DateTime.Now.Day.ToString().PadLeft(2, '0') + "/" + DateTime.Now.Month.ToString().PadLeft(2, '0') + "/" + DateTime.Now.Year;
			WinSettings.TTPeriodInfo.NgayLapKHBS = ((Control)mskNgayLapKHBS1).Text.Trim();
			WinSettings.TTPeriodInfo.SoLan = ((Control)txtLan1).Text.Trim();
		}
		else
		{
			((Control)pnLan1).Visible = false;
			((Control)pnNgayLap1).Visible = false;
			WinSettings.TTPeriodInfo.SoLan = "";
			WinSettings.TTPeriodInfo.NgayLapKHBS = "";
		}
		((Control)lblNgayLapKHBS1).Text = ((((ListControl)cboLoaiToKhai).SelectedIndex == 1) ? "Thay thế" : "Ngày lập KHBS");
		strFileNameExtend = WinFnProcess.GetDataFileNameExtend(WinSettings.TTPeriodInfo);
		WinSettings.DicSheet = WinFnProcess.GetPhuLucDaKeKhai(((Control)pnVT2_PhuLucKeKhai).Visible, strFileNameExtend);
		WinSpreadProcess.BindDataToPhuLucSpread(axfpsPhuLuc.Sheets[0], elemList, chkChonPhuLucKeKhaiAll);
	}

	public void checkSuDungLoaiToKhai()
	{
		if ("N7~".IndexOf(WinSettings.MenuID) >= 0)
		{
			((Control)pnRDLanDauBoXung1).Visible = true;
			((Control)pnRDLanDauBoXung).Visible = false;
			((Control)pnRDLanDauBoXung1).Location = new Point(((Control)pnRDLanDauBoXung1).Location.X, ((Control)pnQuyetToan_LienTuc).Location.Y + 1);
			((Control)pnRDLanDauBoXung1).Size = new Size(((Control)pnRDLanDauBoXung1).Size.Width, ((Control)pnNgayLap1).Size.Height + 2);
			((ListControl)cboLoaiToKhai).SelectedIndex = 0;
			strFileNameExtend = WinFnProcess.GetDataFileNameExtend(WinSettings.TTPeriodInfo);
			WinSettings.DicSheet = WinFnProcess.GetPhuLucDaKeKhai(((Control)pnVT2_PhuLucKeKhai).Visible, strFileNameExtend);
			WinSpreadProcess.BindDataToPhuLucSpread(axfpsPhuLuc.Sheets[0], elemList, chkChonPhuLucKeKhaiAll);
		}
	}

	private void txtLan1_TextChanged(object sender, EventArgs e)
	{
		try
		{
			WinSettings.TTPeriodInfo.SoLan = ((Control)txtLan1).Text.Trim();
			strFileNameExtend = WinFnProcess.GetDataFileNameExtend(WinSettings.TTPeriodInfo);
			WinSettings.DicSheet = WinFnProcess.GetPhuLucDaKeKhai(((Control)pnVT2_PhuLucKeKhai).Visible, strFileNameExtend);
			WinSpreadProcess.BindDataToPhuLucSpread(axfpsPhuLuc.Sheets[0], elemList, chkChonPhuLucKeKhaiAll);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void mskNgayLapKHBS1_TextChanged(object sender, EventArgs e)
	{
		WinSettings.TTPeriodInfo.NgayLapKHBS = ((Control)mskNgayLapKHBS1).Text.Trim();
	}

	private void cboLoaiToKhai_KeyPress(object sender, KeyPressEventArgs e)
	{
		e.Handled = true;
	}

	private void mskTuThang_TextChanged(object sender, EventArgs e)
	{
		CheckNextDenThangTheoQuyetToan();
	}

	private void getTTDuAnDT()
	{
		//IL_00cb: Unknown result type (might be due to invalid IL or missing references)
		//IL_00d2: Expected O, but got Unknown
		if (WinSettings.MenuID != "D2" && WinSettings.MenuID != "99")
		{
			return;
		}
		dataCboDADT.Clear();
		try
		{
			string text = HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst;
			string path = text + "\\DK_Ma_du_an_DT_100.xml";
			int num = HTKKFile.ColumnLettersToIndex("F");
			int num2 = HTKKFile.ColumnLettersToIndex("I");
			if (File.Exists(path))
			{
				XmlDataDocument xmlDataDocument = new XmlDataDocument();
				FileStream fileStream = new FileStream(path, FileMode.Open, FileAccess.Read);
				((XmlDocument)xmlDataDocument).Load((Stream)fileStream);
				XmlNodeList elementsByTagName = ((XmlDocument)xmlDataDocument).GetElementsByTagName("RowData");
				foreach (XmlNode item in elementsByTagName)
				{
					XmlNode val = item;
					string innerText = val.InnerText;
					string[] array = Regex.Split(innerText, "###");
					bool flag = ConvertHelpers.ConvertToBool(array[0]);
					string key = array[1];
					string value = array[3];
					if (!dataCboDADT.ContainsKey(key) && flag)
					{
						dataCboDADT.Add(key, value);
					}
				}
			}
			cbMaHSDNHT_DADT.DataSource = dataCboDADT.Keys.ToList();
			((Control)cbMaHSDNHT_DADT).Text = "";
			((Control)lblChuDADT).Text = "Tên dự án đầu tư: ";
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private string getMaCQTQuanLyDADT(string maDADT)
	{
		//IL_007a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0081: Expected O, but got Unknown
		try
		{
			string text = HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst;
			string path = text + "\\DK_Ma_du_an_DT_100.xml";
			string result = "";
			if (File.Exists(path))
			{
				XmlDataDocument xmlDataDocument = new XmlDataDocument();
				FileStream fileStream = new FileStream(path, FileMode.Open, FileAccess.Read);
				((XmlDocument)xmlDataDocument).Load((Stream)fileStream);
				XmlNodeList elementsByTagName = ((XmlDocument)xmlDataDocument).GetElementsByTagName("RowData");
				foreach (XmlNode item in elementsByTagName)
				{
					XmlNode val = item;
					string innerText = val.InnerText;
					string[] array = Regex.Split(innerText, "###");
					bool flag = ConvertHelpers.ConvertToBool(array[0]);
					string text2 = array[1];
					if (text2.Equals(maDADT))
					{
						result = array[18];
						break;
					}
				}
			}
			return result;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return "";
		}
	}

	private void cbMaHSDNHT_DADT_SelectedIndexChanged(object sender, EventArgs e)
	{
		try
		{
			string text = "";
			if (!((Control)cbMaHSDNHT_DADT).Text.Equals(""))
			{
				if (ConvertHelpers.ConvertToDouble(((Control)cbMaHSDNHT_DADT).Text) == 0.0)
				{
					((Control)cbMaHSDNHT_DADT).Text = "";
				}
				else
				{
					((Control)cbMaHSDNHT_DADT).Text = ConvertHelpers.ConvertToString(ConvertHelpers.ConvertToDouble(((Control)cbMaHSDNHT_DADT).Text));
				}
			}
			((Control)lblChuDADT).Text = "Tên dự án đầu tư: " + (dataCboDADT.ContainsKey(((Control)cbMaHSDNHT_DADT).Text) ? dataCboDADT[((Control)cbMaHSDNHT_DADT).Text] : "");
			if ("U8~".Contains(WinSettings.MenuID))
			{
				WinSettings.TTPeriodInfo.LanXuatBan = ((Control)cbMaHSDNHT_DADT).Text;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void cbMaHSDNHT_DADT_KeyDown(object sender, KeyEventArgs e)
	{
		e.Handled = true;
	}

	private void txtMst_KeyPress(object sender, KeyPressEventArgs e)
	{
		if (!(menuID == "P6") && !(menuID == "P1") && !(menuID == "X7") && !menuID.Equals("Y9") && !char.IsDigit(e.KeyChar) && !char.IsControl(e.KeyChar))
		{
			e.Handled = true;
		}
	}

	private void cboLoaiTien_KeyPress(object sender, KeyPressEventArgs e)
	{
		e.Handled = true;
	}

	private void cboLoaiTien_SelectedIndexChanged(object sender, EventArgs e)
	{
		try
		{
			if (((Control)pnLoaiTien).Visible && !((ListControl)cboLoaiTien).SelectedValue.Equals(WinSettings.TTPeriodInfo.DonViTinhMa))
			{
				LoadPhuLucSpread();
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void mskNgayTb_Leave(object sender, EventArgs e)
	{
		try
		{
			if (!HTKKFormater.isNullString(((Control)mskNgayTb).Text.Replace("/", "").Replace("_", "")))
			{
				string text = HTKKFormater.FormatDdmmyyyy(((Control)mskNgayTb).Text);
				((Control)mskNgayTb).Text = text;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void cboXa_SelectedIndexChanged(object sender, EventArgs e)
	{
		string tenTinh = ((Control)cboTinh).Text.Trim();
		string text = ((Control)cboXa).Text.Trim();
		if (text == "Chọn phường xã")
		{
			ResetCQTTinhToEmpty(resetTinh: false, resetXa: true, resetCQTCapCuc: true, resetCQTQuanLy: true);
			return;
		}
		string maXa = HTKKWinDictionaryDiaDanh.getMaPhuongXaMoi(text, HTKKWinDictionaryDiaDanh.dicTinhTPMoi.FirstOrDefault((KeyValuePair<string, string> x) => x.Key == tenTinh).Value);
		if (string.IsNullOrEmpty(maXa))
		{
			return;
		}
		WinSettings.TTPeriodInfo.MaXa = maXa;
		WinSettings.TTPeriodInfo.TenXa = text;
		ResetCQTTinhToEmpty(resetTinh: false, resetXa: false, resetCQTCapCuc: true, resetCQTQuanLy: true);
		cboCQTCuc.DataSource = GeneralFunctions.loadCqtCapCucTheoDiaBanXa(maXa).Values.ToList();
		cboCQTChiCuc.DataSource = GeneralFunctions.loadCqtCapChiCucTheoDiaBanXa(maXa).Values.ToList();
		if (menuID == "Y7")
		{
			CQTDiaBan cQTDiaBan = HTKKWinDictionaryDiaDanh._cqtDiaBan.Find((CQTDiaBan x) => x.MaXa == maXa);
			string maCuc = cQTDiaBan.MaCQT.Substring(0, 3) + "00";
			CQTDiaBan cQTDiaBan2 = HTKKWinDictionaryDiaDanh._cqtDiaBan.Find((CQTDiaBan x) => x.MaCQT == maCuc);
			((Control)cboCQTCuc).Text = cQTDiaBan2.TenCQT;
			((Control)cboCQTChiCuc).Text = cQTDiaBan.TenCQT;
		}
	}

	private void lockTinhXaCqt(bool isLock)
	{
		((Control)cboTinh).Enabled = isLock;
		((Control)cboXa).Enabled = isLock;
		((Control)cboCQTCuc).Enabled = isLock;
		((Control)cboCQTChiCuc).Enabled = isLock;
	}

	private void chkToChucCaNhanChuyenNhuong_CheckedChanged(object sender, EventArgs e)
	{
		if (chkToChucCaNhanChuyenNhuong.Checked)
		{
			chkDoanhNghiepChuyenNhuong.Checked = false;
		}
	}

	private void chkDoanhNghiepChuyenNhuong_CheckedChanged(object sender, EventArgs e)
	{
		if (chkDoanhNghiepChuyenNhuong.Checked)
		{
			chkToChucCaNhanChuyenNhuong.Checked = false;
		}
	}

	protected override void Dispose(bool disposing)
	{
		if (disposing && components != null)
		{
			components.Dispose();
		}
		((Form)this).Dispose(disposing);
	}

	private void InitializeComponent()
	{
		//IL_0012: Unknown result type (might be due to invalid IL or missing references)
		//IL_001c: Expected O, but got Unknown
		//IL_001d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0027: Expected O, but got Unknown
		//IL_0028: Unknown result type (might be due to invalid IL or missing references)
		//IL_0032: Expected O, but got Unknown
		//IL_0033: Unknown result type (might be due to invalid IL or missing references)
		//IL_003d: Expected O, but got Unknown
		//IL_003e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0048: Expected O, but got Unknown
		//IL_0049: Unknown result type (might be due to invalid IL or missing references)
		//IL_0053: Expected O, but got Unknown
		//IL_0054: Unknown result type (might be due to invalid IL or missing references)
		//IL_005e: Expected O, but got Unknown
		//IL_005f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0069: Expected O, but got Unknown
		//IL_006a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0074: Expected O, but got Unknown
		//IL_008b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0095: Expected O, but got Unknown
		//IL_0096: Unknown result type (might be due to invalid IL or missing references)
		//IL_00a0: Expected O, but got Unknown
		//IL_00a1: Unknown result type (might be due to invalid IL or missing references)
		//IL_00ab: Expected O, but got Unknown
		//IL_00ac: Unknown result type (might be due to invalid IL or missing references)
		//IL_00b6: Expected O, but got Unknown
		//IL_00b7: Unknown result type (might be due to invalid IL or missing references)
		//IL_00c1: Expected O, but got Unknown
		//IL_00c2: Unknown result type (might be due to invalid IL or missing references)
		//IL_00cc: Expected O, but got Unknown
		//IL_00cd: Unknown result type (might be due to invalid IL or missing references)
		//IL_00d7: Expected O, but got Unknown
		//IL_00d8: Unknown result type (might be due to invalid IL or missing references)
		//IL_00e2: Expected O, but got Unknown
		//IL_00e3: Unknown result type (might be due to invalid IL or missing references)
		//IL_00ed: Expected O, but got Unknown
		//IL_00ee: Unknown result type (might be due to invalid IL or missing references)
		//IL_00f8: Expected O, but got Unknown
		//IL_00f9: Unknown result type (might be due to invalid IL or missing references)
		//IL_0103: Expected O, but got Unknown
		//IL_0104: Unknown result type (might be due to invalid IL or missing references)
		//IL_010e: Expected O, but got Unknown
		//IL_010f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0119: Expected O, but got Unknown
		//IL_011a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0124: Expected O, but got Unknown
		//IL_0125: Unknown result type (might be due to invalid IL or missing references)
		//IL_012f: Expected O, but got Unknown
		//IL_0130: Unknown result type (might be due to invalid IL or missing references)
		//IL_013a: Expected O, but got Unknown
		//IL_013b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0145: Expected O, but got Unknown
		//IL_0146: Unknown result type (might be due to invalid IL or missing references)
		//IL_0150: Expected O, but got Unknown
		//IL_0151: Unknown result type (might be due to invalid IL or missing references)
		//IL_015b: Expected O, but got Unknown
		//IL_015c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0166: Expected O, but got Unknown
		//IL_0167: Unknown result type (might be due to invalid IL or missing references)
		//IL_0171: Expected O, but got Unknown
		//IL_0172: Unknown result type (might be due to invalid IL or missing references)
		//IL_017c: Expected O, but got Unknown
		//IL_017d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0187: Expected O, but got Unknown
		//IL_0188: Unknown result type (might be due to invalid IL or missing references)
		//IL_0192: Expected O, but got Unknown
		//IL_0193: Unknown result type (might be due to invalid IL or missing references)
		//IL_019d: Expected O, but got Unknown
		//IL_019e: Unknown result type (might be due to invalid IL or missing references)
		//IL_01a8: Expected O, but got Unknown
		//IL_01a9: Unknown result type (might be due to invalid IL or missing references)
		//IL_01b3: Expected O, but got Unknown
		//IL_01b4: Unknown result type (might be due to invalid IL or missing references)
		//IL_01be: Expected O, but got Unknown
		//IL_01bf: Unknown result type (might be due to invalid IL or missing references)
		//IL_01c9: Expected O, but got Unknown
		//IL_01ca: Unknown result type (might be due to invalid IL or missing references)
		//IL_01d4: Expected O, but got Unknown
		//IL_01d5: Unknown result type (might be due to invalid IL or missing references)
		//IL_01df: Expected O, but got Unknown
		//IL_01e0: Unknown result type (might be due to invalid IL or missing references)
		//IL_01ea: Expected O, but got Unknown
		//IL_01eb: Unknown result type (might be due to invalid IL or missing references)
		//IL_01f5: Expected O, but got Unknown
		//IL_01f6: Unknown result type (might be due to invalid IL or missing references)
		//IL_0200: Expected O, but got Unknown
		//IL_0201: Unknown result type (might be due to invalid IL or missing references)
		//IL_020b: Expected O, but got Unknown
		//IL_020c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0216: Expected O, but got Unknown
		//IL_0217: Unknown result type (might be due to invalid IL or missing references)
		//IL_0221: Expected O, but got Unknown
		//IL_0222: Unknown result type (might be due to invalid IL or missing references)
		//IL_022c: Expected O, but got Unknown
		//IL_022d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0237: Expected O, but got Unknown
		//IL_0238: Unknown result type (might be due to invalid IL or missing references)
		//IL_0242: Expected O, but got Unknown
		//IL_0243: Unknown result type (might be due to invalid IL or missing references)
		//IL_024d: Expected O, but got Unknown
		//IL_024e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0258: Expected O, but got Unknown
		//IL_0259: Unknown result type (might be due to invalid IL or missing references)
		//IL_0263: Expected O, but got Unknown
		//IL_0264: Unknown result type (might be due to invalid IL or missing references)
		//IL_026e: Expected O, but got Unknown
		//IL_026f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0279: Expected O, but got Unknown
		//IL_027a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0284: Expected O, but got Unknown
		//IL_0285: Unknown result type (might be due to invalid IL or missing references)
		//IL_028f: Expected O, but got Unknown
		//IL_0290: Unknown result type (might be due to invalid IL or missing references)
		//IL_029a: Expected O, but got Unknown
		//IL_029b: Unknown result type (might be due to invalid IL or missing references)
		//IL_02a5: Expected O, but got Unknown
		//IL_02a6: Unknown result type (might be due to invalid IL or missing references)
		//IL_02b0: Expected O, but got Unknown
		//IL_02b1: Unknown result type (might be due to invalid IL or missing references)
		//IL_02bb: Expected O, but got Unknown
		//IL_02bc: Unknown result type (might be due to invalid IL or missing references)
		//IL_02c6: Expected O, but got Unknown
		//IL_02c7: Unknown result type (might be due to invalid IL or missing references)
		//IL_02d1: Expected O, but got Unknown
		//IL_02d2: Unknown result type (might be due to invalid IL or missing references)
		//IL_02dc: Expected O, but got Unknown
		//IL_02dd: Unknown result type (might be due to invalid IL or missing references)
		//IL_02e7: Expected O, but got Unknown
		//IL_02e8: Unknown result type (might be due to invalid IL or missing references)
		//IL_02f2: Expected O, but got Unknown
		//IL_02f3: Unknown result type (might be due to invalid IL or missing references)
		//IL_02fd: Expected O, but got Unknown
		//IL_02fe: Unknown result type (might be due to invalid IL or missing references)
		//IL_0308: Expected O, but got Unknown
		//IL_0309: Unknown result type (might be due to invalid IL or missing references)
		//IL_0313: Expected O, but got Unknown
		//IL_0314: Unknown result type (might be due to invalid IL or missing references)
		//IL_031e: Expected O, but got Unknown
		//IL_031f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0329: Expected O, but got Unknown
		//IL_032a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0334: Expected O, but got Unknown
		//IL_0335: Unknown result type (might be due to invalid IL or missing references)
		//IL_033f: Expected O, but got Unknown
		//IL_0340: Unknown result type (might be due to invalid IL or missing references)
		//IL_034a: Expected O, but got Unknown
		//IL_034b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0355: Expected O, but got Unknown
		//IL_0356: Unknown result type (might be due to invalid IL or missing references)
		//IL_0360: Expected O, but got Unknown
		//IL_0361: Unknown result type (might be due to invalid IL or missing references)
		//IL_036b: Expected O, but got Unknown
		//IL_036c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0376: Expected O, but got Unknown
		//IL_0377: Unknown result type (might be due to invalid IL or missing references)
		//IL_0381: Expected O, but got Unknown
		//IL_0382: Unknown result type (might be due to invalid IL or missing references)
		//IL_038c: Expected O, but got Unknown
		//IL_038d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0397: Expected O, but got Unknown
		//IL_0398: Unknown result type (might be due to invalid IL or missing references)
		//IL_03a2: Expected O, but got Unknown
		//IL_03a3: Unknown result type (might be due to invalid IL or missing references)
		//IL_03ad: Expected O, but got Unknown
		//IL_03ae: Unknown result type (might be due to invalid IL or missing references)
		//IL_03b8: Expected O, but got Unknown
		//IL_03b9: Unknown result type (might be due to invalid IL or missing references)
		//IL_03c3: Expected O, but got Unknown
		//IL_03c4: Unknown result type (might be due to invalid IL or missing references)
		//IL_03ce: Expected O, but got Unknown
		//IL_03cf: Unknown result type (might be due to invalid IL or missing references)
		//IL_03d9: Expected O, but got Unknown
		//IL_03da: Unknown result type (might be due to invalid IL or missing references)
		//IL_03e4: Expected O, but got Unknown
		//IL_03e5: Unknown result type (might be due to invalid IL or missing references)
		//IL_03ef: Expected O, but got Unknown
		//IL_03f0: Unknown result type (might be due to invalid IL or missing references)
		//IL_03fa: Expected O, but got Unknown
		//IL_03fb: Unknown result type (might be due to invalid IL or missing references)
		//IL_0405: Expected O, but got Unknown
		//IL_0406: Unknown result type (might be due to invalid IL or missing references)
		//IL_0410: Expected O, but got Unknown
		//IL_0411: Unknown result type (might be due to invalid IL or missing references)
		//IL_041b: Expected O, but got Unknown
		//IL_041c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0426: Expected O, but got Unknown
		//IL_0427: Unknown result type (might be due to invalid IL or missing references)
		//IL_0431: Expected O, but got Unknown
		//IL_0432: Unknown result type (might be due to invalid IL or missing references)
		//IL_043c: Expected O, but got Unknown
		//IL_043d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0447: Expected O, but got Unknown
		//IL_0448: Unknown result type (might be due to invalid IL or missing references)
		//IL_0452: Expected O, but got Unknown
		//IL_0453: Unknown result type (might be due to invalid IL or missing references)
		//IL_045d: Expected O, but got Unknown
		//IL_045e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0468: Expected O, but got Unknown
		//IL_0469: Unknown result type (might be due to invalid IL or missing references)
		//IL_0473: Expected O, but got Unknown
		//IL_0474: Unknown result type (might be due to invalid IL or missing references)
		//IL_047e: Expected O, but got Unknown
		//IL_047f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0489: Expected O, but got Unknown
		//IL_048a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0494: Expected O, but got Unknown
		//IL_0495: Unknown result type (might be due to invalid IL or missing references)
		//IL_049f: Expected O, but got Unknown
		//IL_04a0: Unknown result type (might be due to invalid IL or missing references)
		//IL_04aa: Expected O, but got Unknown
		//IL_04ab: Unknown result type (might be due to invalid IL or missing references)
		//IL_04b5: Expected O, but got Unknown
		//IL_04b6: Unknown result type (might be due to invalid IL or missing references)
		//IL_04c0: Expected O, but got Unknown
		//IL_04c1: Unknown result type (might be due to invalid IL or missing references)
		//IL_04cb: Expected O, but got Unknown
		//IL_04cc: Unknown result type (might be due to invalid IL or missing references)
		//IL_04d6: Expected O, but got Unknown
		//IL_04d7: Unknown result type (might be due to invalid IL or missing references)
		//IL_04e1: Expected O, but got Unknown
		//IL_04e2: Unknown result type (might be due to invalid IL or missing references)
		//IL_04ec: Expected O, but got Unknown
		//IL_04ed: Unknown result type (might be due to invalid IL or missing references)
		//IL_04f7: Expected O, but got Unknown
		//IL_04f8: Unknown result type (might be due to invalid IL or missing references)
		//IL_0502: Expected O, but got Unknown
		//IL_0503: Unknown result type (might be due to invalid IL or missing references)
		//IL_050d: Expected O, but got Unknown
		//IL_050e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0518: Expected O, but got Unknown
		//IL_0519: Unknown result type (might be due to invalid IL or missing references)
		//IL_0523: Expected O, but got Unknown
		//IL_0524: Unknown result type (might be due to invalid IL or missing references)
		//IL_052e: Expected O, but got Unknown
		//IL_052f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0539: Expected O, but got Unknown
		//IL_053a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0544: Expected O, but got Unknown
		//IL_0545: Unknown result type (might be due to invalid IL or missing references)
		//IL_054f: Expected O, but got Unknown
		//IL_0550: Unknown result type (might be due to invalid IL or missing references)
		//IL_055a: Expected O, but got Unknown
		//IL_055b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0565: Expected O, but got Unknown
		//IL_0566: Unknown result type (might be due to invalid IL or missing references)
		//IL_0570: Expected O, but got Unknown
		//IL_0571: Unknown result type (might be due to invalid IL or missing references)
		//IL_057b: Expected O, but got Unknown
		//IL_057c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0586: Expected O, but got Unknown
		//IL_0587: Unknown result type (might be due to invalid IL or missing references)
		//IL_0591: Expected O, but got Unknown
		//IL_0592: Unknown result type (might be due to invalid IL or missing references)
		//IL_059c: Expected O, but got Unknown
		//IL_059d: Unknown result type (might be due to invalid IL or missing references)
		//IL_05a7: Expected O, but got Unknown
		//IL_05a8: Unknown result type (might be due to invalid IL or missing references)
		//IL_05b2: Expected O, but got Unknown
		//IL_05b3: Unknown result type (might be due to invalid IL or missing references)
		//IL_05bd: Expected O, but got Unknown
		//IL_05be: Unknown result type (might be due to invalid IL or missing references)
		//IL_05c8: Expected O, but got Unknown
		//IL_05c9: Unknown result type (might be due to invalid IL or missing references)
		//IL_05d3: Expected O, but got Unknown
		//IL_05d4: Unknown result type (might be due to invalid IL or missing references)
		//IL_05de: Expected O, but got Unknown
		//IL_153a: Unknown result type (might be due to invalid IL or missing references)
		//IL_1544: Expected O, but got Unknown
		//IL_1bea: Unknown result type (might be due to invalid IL or missing references)
		//IL_1bf4: Expected O, but got Unknown
		//IL_1c48: Unknown result type (might be due to invalid IL or missing references)
		//IL_1cc6: Unknown result type (might be due to invalid IL or missing references)
		//IL_1d2c: Unknown result type (might be due to invalid IL or missing references)
		//IL_1d36: Expected O, but got Unknown
		//IL_1d44: Unknown result type (might be due to invalid IL or missing references)
		//IL_1d4e: Expected O, but got Unknown
		//IL_1d7a: Unknown result type (might be due to invalid IL or missing references)
		//IL_1ed9: Unknown result type (might be due to invalid IL or missing references)
		//IL_1ee3: Expected O, but got Unknown
		//IL_1f9d: Unknown result type (might be due to invalid IL or missing references)
		//IL_2157: Unknown result type (might be due to invalid IL or missing references)
		//IL_2328: Unknown result type (might be due to invalid IL or missing references)
		//IL_2491: Unknown result type (might be due to invalid IL or missing references)
		//IL_249b: Expected O, but got Unknown
		//IL_24a9: Unknown result type (might be due to invalid IL or missing references)
		//IL_24b3: Expected O, but got Unknown
		//IL_260d: Unknown result type (might be due to invalid IL or missing references)
		//IL_2617: Expected O, but got Unknown
		//IL_2625: Unknown result type (might be due to invalid IL or missing references)
		//IL_262f: Expected O, but got Unknown
		//IL_27a8: Unknown result type (might be due to invalid IL or missing references)
		//IL_27b2: Expected O, but got Unknown
		//IL_27c0: Unknown result type (might be due to invalid IL or missing references)
		//IL_27ca: Expected O, but got Unknown
		//IL_2882: Unknown result type (might be due to invalid IL or missing references)
		//IL_2a07: Unknown result type (might be due to invalid IL or missing references)
		//IL_2b2c: Unknown result type (might be due to invalid IL or missing references)
		//IL_3946: Unknown result type (might be due to invalid IL or missing references)
		//IL_3950: Expected O, but got Unknown
		//IL_3b06: Unknown result type (might be due to invalid IL or missing references)
		//IL_3c63: Unknown result type (might be due to invalid IL or missing references)
		//IL_3c6d: Expected O, but got Unknown
		//IL_40c4: Unknown result type (might be due to invalid IL or missing references)
		//IL_40ce: Expected O, but got Unknown
		//IL_43ff: Unknown result type (might be due to invalid IL or missing references)
		//IL_4409: Expected O, but got Unknown
		//IL_4417: Unknown result type (might be due to invalid IL or missing references)
		//IL_4421: Expected O, but got Unknown
		//IL_4571: Unknown result type (might be due to invalid IL or missing references)
		//IL_457b: Expected O, but got Unknown
		//IL_4589: Unknown result type (might be due to invalid IL or missing references)
		//IL_4593: Expected O, but got Unknown
		//IL_46a7: Unknown result type (might be due to invalid IL or missing references)
		//IL_46b1: Expected O, but got Unknown
		//IL_4767: Unknown result type (might be due to invalid IL or missing references)
		//IL_48d8: Unknown result type (might be due to invalid IL or missing references)
		//IL_4b6c: Unknown result type (might be due to invalid IL or missing references)
		//IL_4b76: Expected O, but got Unknown
		//IL_507b: Unknown result type (might be due to invalid IL or missing references)
		//IL_5085: Expected O, but got Unknown
		//IL_50a5: Unknown result type (might be due to invalid IL or missing references)
		//IL_5375: Unknown result type (might be due to invalid IL or missing references)
		//IL_537f: Expected O, but got Unknown
		ComponentResourceManager componentResourceManager = new ComponentResourceManager(typeof(frmHTKKPeriod));
		OptTKThang = new RadioButton();
		OptTKLanPS = new RadioButton();
		chkDauTho = new CheckBox();
		chkKhiThien = new CheckBox();
		chkCondensate = new CheckBox();
		pnVT3Button = new Panel();
		btnDongY = new Button();
		cmdClose = new Button();
		pnVT2_PhuLucKeKhai = new Panel();
		axfpsPhuLuc = new FpSpread();
		axfpsPhuLuc_Sheet1 = new SheetView();
		chkChonPhuLucKeKhaiAll = new CheckBox();
		f = new Panel();
		pnThongBao = new Panel();
		mskNgayTb = new MaskedTextBox();
		lblNgayTb = new Label();
		txtSoTb = new TextBox();
		lblSoTB = new Label();
		pnMst = new Panel();
		lblMst = new Label();
		txtMst = new TextBox();
		pnVT1 = new Panel();
		pnRDLanDauBoXung1 = new Panel();
		pnNgayLap1 = new Panel();
		mskNgayLapKHBS1 = new MaskedTextBox();
		lblNgayLapKHBS1 = new Label();
		pnLan1 = new Panel();
		txtLan1 = new TextBox();
		label5 = new Label();
		label4 = new Label();
		cboLoaiToKhai = new ComboBox();
		pnTruongHopQT = new Panel();
		cboTruongHopQT = new ComboBox();
		label3 = new Label();
		pnLoaiKyHoan = new Panel();
		cboKieuKy = new ComboBox();
		label2 = new Label();
		pnChuyenNhuong = new Panel();
		chkDoanhNghiepChuyenNhuong = new CheckBox();
		chkToChucCaNhanChuyenNhuong = new CheckBox();
		pnKhaiTheoPLuatDanSuThue = new Panel();
		chkKhaiTheoPLuatThue = new CheckBox();
		chkKhaiTheoPLuatDanSu = new CheckBox();
		pnTinhCQT = new Panel();
		pnCCT = new Panel();
		cboCQTChiCuc = new ComboBox();
		lblCQTChiCuc = new Label();
		pnCQTCuc = new Panel();
		cboCQTCuc = new ComboBox();
		lblCQTCuc = new Label();
		pnTinh = new Panel();
		cboTinh = new ComboBox();
		lblTinh = new Label();
		pnMSTCaNhanToChuc = new Panel();
		chkMSTToChuc = new CheckBox();
		chkMSTCaNhan = new CheckBox();
		pnPPhapNopThue = new Panel();
		pnVT1_1CheckBox = new Panel();
		chkTKLanPS = new CheckBox();
		ckNam = new CheckBox();
		chkTKhaiLanXB = new CheckBox();
		chkTKKy = new CheckBox();
		chkTKQuy = new CheckBox();
		chkTuThangDenThang = new CheckBox();
		chkTuQuyDenQuy = new CheckBox();
		chkTkhaiThang = new CheckBox();
		pnVT1_4LanDauBoXung = new Panel();
		pnQuyetToan_LienTuc = new Panel();
		cbxCTDC = new CheckBox();
		chkQTTungNam_HoatDongLienTuc = new CheckBox();
		lbNgayPS = new Label();
		cbxLNDC = new CheckBox();
		mskNgayPS = new MaskedTextBox();
		chk_CoSoThanhLap = new CheckBox();
		chkQTNamDau_KhongHoatDongLienTuc = new CheckBox();
		pnNgayLap = new Panel();
		mskNgayLapKHBS = new MaskedTextBox();
		lblNgayLapKHBS = new Label();
		pnRDLanDauBoXung = new Panel();
		OptBosung = new RadioButton();
		pnLan = new Panel();
		txtLan = new TextBox();
		lblLan = new Label();
		OptChinhthuc = new RadioButton();
		pnVT1_2NgayThangQuyNam = new Panel();
		pnKy = new Panel();
		cbbKy = new ComboBox();
		lblKy = new Label();
		pnNam = new Panel();
		mskYear = new MaskedTextBox();
		lblYear = new Label();
		pnNgay = new Panel();
		lblNgay = new Label();
		mskDay = new MaskedTextBox();
		pnQuy = new Panel();
		cbbQuy = new ComboBox();
		lblQuy = new Label();
		pnThang = new Panel();
		mskMonth = new MaskedTextBox();
		lblMonth = new Label();
		pnVT1_5DanhMucNghe = new Panel();
		lblNganhKD = new Label();
		cboNganhKD = new ComboBox();
		pnMaHSDNHT = new Panel();
		cbMaHSDNHT_DADT = new ComboBox();
		lblChuDADT = new Label();
		txtMaHSDNHT_DADT = new TextBox();
		label1 = new Label();
		pnThuHo = new Panel();
		chkThuHo = new CheckBox();
		pnVT1_3TuNgayDenNgay = new Panel();
		pnDenThang = new Panel();
		mskDenThang = new MaskedTextBox();
		lblDenThang = new Label();
		pnLanXuatBan = new Panel();
		txtLanXuatBan = new TextBox();
		lblLanXuat = new Label();
		pnDenNgay = new Panel();
		mskDenNgay = new MaskedTextBox();
		lblDenNgay = new Label();
		pnTuThang = new Panel();
		mskTuThang = new MaskedTextBox();
		lblTuThang = new Label();
		pnTuNgay = new Panel();
		mskTuNgay = new MaskedTextBox();
		lblTuNgay = new Label();
		pbHeader = new PictureBox();
		chkToChucKhaiThay = new CheckBox();
		pnLoaiTien = new Panel();
		lblLoaiTien = new Label();
		cboLoaiTien = new ComboBox();
		panel1 = new Panel();
		cboXa = new ComboBox();
		lblXa = new Label();
		lblNgayPsThuNhap = new Label();
		((Control)pnVT3Button).SuspendLayout();
		((Control)pnVT2_PhuLucKeKhai).SuspendLayout();
		((ISupportInitialize)axfpsPhuLuc).BeginInit();
		((ISupportInitialize)axfpsPhuLuc_Sheet1).BeginInit();
		((Control)f).SuspendLayout();
		((Control)pnThongBao).SuspendLayout();
		((Control)pnMst).SuspendLayout();
		((Control)pnVT1).SuspendLayout();
		((Control)pnRDLanDauBoXung1).SuspendLayout();
		((Control)pnNgayLap1).SuspendLayout();
		((Control)pnLan1).SuspendLayout();
		((Control)pnTruongHopQT).SuspendLayout();
		((Control)pnLoaiKyHoan).SuspendLayout();
		((Control)pnChuyenNhuong).SuspendLayout();
		((Control)pnKhaiTheoPLuatDanSuThue).SuspendLayout();
		((Control)pnTinhCQT).SuspendLayout();
		((Control)pnCCT).SuspendLayout();
		((Control)pnCQTCuc).SuspendLayout();
		((Control)pnTinh).SuspendLayout();
		((Control)pnMSTCaNhanToChuc).SuspendLayout();
		((Control)pnPPhapNopThue).SuspendLayout();
		((Control)pnVT1_1CheckBox).SuspendLayout();
		((Control)pnVT1_4LanDauBoXung).SuspendLayout();
		((Control)pnQuyetToan_LienTuc).SuspendLayout();
		((Control)pnNgayLap).SuspendLayout();
		((Control)pnRDLanDauBoXung).SuspendLayout();
		((Control)pnLan).SuspendLayout();
		((Control)pnVT1_2NgayThangQuyNam).SuspendLayout();
		((Control)pnKy).SuspendLayout();
		((Control)pnNam).SuspendLayout();
		((Control)pnNgay).SuspendLayout();
		((Control)pnQuy).SuspendLayout();
		((Control)pnThang).SuspendLayout();
		((Control)pnVT1_5DanhMucNghe).SuspendLayout();
		((Control)pnMaHSDNHT).SuspendLayout();
		((Control)pnThuHo).SuspendLayout();
		((Control)pnVT1_3TuNgayDenNgay).SuspendLayout();
		((Control)pnDenThang).SuspendLayout();
		((Control)pnLanXuatBan).SuspendLayout();
		((Control)pnDenNgay).SuspendLayout();
		((Control)pnTuThang).SuspendLayout();
		((Control)pnTuNgay).SuspendLayout();
		((ISupportInitialize)pbHeader).BeginInit();
		((Control)pnLoaiTien).SuspendLayout();
		((Control)panel1).SuspendLayout();
		((Control)this).SuspendLayout();
		((Control)OptTKThang).AutoSize = true;
		((Control)OptTKThang).Location = new Point(12, 737);
		((Control)OptTKThang).Name = "OptTKThang";
		((Control)OptTKThang).Size = new Size(91, 17);
		((Control)OptTKThang).TabIndex = 25;
		((Control)OptTKThang).Text = "Tờ khai tháng";
		((ButtonBase)OptTKThang).UseVisualStyleBackColor = true;
		((Control)OptTKLanPS).AutoSize = true;
		((Control)OptTKLanPS).Location = new Point(185, 737);
		((Control)OptTKLanPS).Name = "OptTKLanPS";
		((Control)OptTKLanPS).Size = new Size(124, 17);
		((Control)OptTKLanPS).TabIndex = 26;
		((Control)OptTKLanPS).Text = "Tờ khai lần phát sinh";
		((ButtonBase)OptTKLanPS).UseVisualStyleBackColor = true;
		((Control)chkDauTho).AutoSize = true;
		((Control)chkDauTho).Location = new Point(0, 769);
		((Control)chkDauTho).Name = "chkDauTho";
		((Control)chkDauTho).Size = new Size(64, 17);
		((Control)chkDauTho).TabIndex = 33;
		((Control)chkDauTho).Text = "Dầu thô";
		((ButtonBase)chkDauTho).UseVisualStyleBackColor = true;
		((Control)chkKhiThien).AutoSize = true;
		((Control)chkKhiThien).Location = new Point(156, 769);
		((Control)chkKhiThien).Name = "chkKhiThien";
		((Control)chkKhiThien).Size = new Size(98, 17);
		((Control)chkKhiThien).TabIndex = 34;
		((Control)chkKhiThien).Text = "Khí thiên nhiên";
		((ButtonBase)chkKhiThien).UseVisualStyleBackColor = true;
		((Control)chkCondensate).AutoSize = true;
		((Control)chkCondensate).Location = new Point(69, 769);
		((Control)chkCondensate).Name = "chkCondensate";
		((Control)chkCondensate).Size = new Size(83, 17);
		((Control)chkCondensate).TabIndex = 35;
		((Control)chkCondensate).Text = "Condensate";
		((ButtonBase)chkCondensate).UseVisualStyleBackColor = true;
		((Control)pnVT3Button).Controls.Add((Control)(object)btnDongY);
		((Control)pnVT3Button).Controls.Add((Control)(object)cmdClose);
		((Control)pnVT3Button).Location = new Point(4, 673);
		((Control)pnVT3Button).Name = "pnVT3Button";
		((Control)pnVT3Button).Size = new Size(372, 28);
		((Control)pnVT3Button).TabIndex = 1;
		((Control)btnDongY).Location = new Point(103, 2);
		((Control)btnDongY).Name = "btnDongY";
		((Control)btnDongY).Size = new Size(80, 25);
		((Control)btnDongY).TabIndex = 8;
		((Control)btnDongY).Text = "Đồn&g ý";
		((ButtonBase)btnDongY).UseVisualStyleBackColor = true;
		((Control)btnDongY).Click += btnDongY_Click;
		((Control)cmdClose).Location = new Point(188, 2);
		((Control)cmdClose).Name = "cmdClose";
		((Control)cmdClose).Size = new Size(80, 25);
		((Control)cmdClose).TabIndex = 6;
		((Control)cmdClose).Text = "Đó&ng";
		((ButtonBase)cmdClose).UseVisualStyleBackColor = true;
		((Control)cmdClose).Click += Close_Click;
		pnVT2_PhuLucKeKhai.BorderStyle = (BorderStyle)1;
		((Control)pnVT2_PhuLucKeKhai).Controls.Add((Control)(object)axfpsPhuLuc);
		((Control)pnVT2_PhuLucKeKhai).Controls.Add((Control)(object)chkChonPhuLucKeKhaiAll);
		((Control)pnVT2_PhuLucKeKhai).Location = new Point(4, 572);
		((Control)pnVT2_PhuLucKeKhai).Name = "pnVT2_PhuLucKeKhai";
		((Control)pnVT2_PhuLucKeKhai).Size = new Size(372, 95);
		((Control)pnVT2_PhuLucKeKhai).TabIndex = 2;
		((Control)axfpsPhuLuc).AccessibleDescription = "axfpsPhuLuc, Sheet1, Row 0, Column 0, ";
		((Control)axfpsPhuLuc).Location = new Point(4, 26);
		axfpsPhuLuc.Name = "axfpsPhuLuc";
		axfpsPhuLuc.Sheets.AddRange(new SheetView[1] { axfpsPhuLuc_Sheet1 });
		((Control)axfpsPhuLuc).Size = new Size(350, 63);
		((Control)axfpsPhuLuc).TabIndex = 2;
		axfpsPhuLuc.ButtonClicked += axfpsPhuLuc_ButtonClicked;
		axfpsPhuLuc_Sheet1.Reset();
		axfpsPhuLuc_Sheet1.SheetName = "Sheet1";
		axfpsPhuLuc_Sheet1.ReferenceStyle = ReferenceStyle.R1C1;
		axfpsPhuLuc_Sheet1.ColumnCount = 1;
		axfpsPhuLuc_Sheet1.ColumnFooter.DefaultStyle.HorizontalAlignment = CellHorizontalAlignment.General;
		axfpsPhuLuc_Sheet1.ColumnFooter.DefaultStyle.NoteIndicatorColor = Color.Red;
		axfpsPhuLuc_Sheet1.ColumnFooter.DefaultStyle.Parent = "ColumnHeaderDefaultEnhanced";
		axfpsPhuLuc_Sheet1.ColumnFooter.DefaultStyle.VerticalAlignment = CellVerticalAlignment.General;
		axfpsPhuLuc_Sheet1.ColumnFooterSheetCornerStyle.HorizontalAlignment = CellHorizontalAlignment.General;
		axfpsPhuLuc_Sheet1.ColumnFooterSheetCornerStyle.NoteIndicatorColor = Color.Red;
		axfpsPhuLuc_Sheet1.ColumnFooterSheetCornerStyle.Parent = "CornerFooterDefaultEnhanced";
		axfpsPhuLuc_Sheet1.ColumnFooterSheetCornerStyle.VerticalAlignment = CellVerticalAlignment.General;
		axfpsPhuLuc_Sheet1.ColumnHeader.DefaultStyle.HorizontalAlignment = CellHorizontalAlignment.General;
		axfpsPhuLuc_Sheet1.ColumnHeader.DefaultStyle.NoteIndicatorColor = Color.Red;
		axfpsPhuLuc_Sheet1.ColumnHeader.DefaultStyle.Parent = "ColumnHeaderDefaultEnhanced";
		axfpsPhuLuc_Sheet1.ColumnHeader.DefaultStyle.VerticalAlignment = CellVerticalAlignment.General;
		axfpsPhuLuc_Sheet1.Columns.Get(0).Width = 308f;
		axfpsPhuLuc_Sheet1.DefaultStyle.HorizontalAlignment = CellHorizontalAlignment.General;
		axfpsPhuLuc_Sheet1.DefaultStyle.NoteIndicatorColor = Color.Red;
		axfpsPhuLuc_Sheet1.DefaultStyle.Parent = "DataAreaDefault";
		axfpsPhuLuc_Sheet1.DefaultStyle.VerticalAlignment = CellVerticalAlignment.General;
		axfpsPhuLuc_Sheet1.FilterBar.DefaultStyle.HorizontalAlignment = CellHorizontalAlignment.General;
		axfpsPhuLuc_Sheet1.FilterBar.DefaultStyle.NoteIndicatorColor = Color.Red;
		axfpsPhuLuc_Sheet1.FilterBar.DefaultStyle.Parent = "FilterBarDefaultEnhanced";
		axfpsPhuLuc_Sheet1.FilterBar.DefaultStyle.VerticalAlignment = CellVerticalAlignment.General;
		axfpsPhuLuc_Sheet1.FilterBarHeaderStyle.HorizontalAlignment = CellHorizontalAlignment.General;
		axfpsPhuLuc_Sheet1.FilterBarHeaderStyle.NoteIndicatorColor = Color.Red;
		axfpsPhuLuc_Sheet1.FilterBarHeaderStyle.Parent = "RowHeaderDefaultEnhanced";
		axfpsPhuLuc_Sheet1.FilterBarHeaderStyle.VerticalAlignment = CellVerticalAlignment.General;
		axfpsPhuLuc_Sheet1.RowHeader.Columns.Default.Resizable = false;
		axfpsPhuLuc_Sheet1.RowHeader.DefaultStyle.HorizontalAlignment = CellHorizontalAlignment.General;
		axfpsPhuLuc_Sheet1.RowHeader.DefaultStyle.NoteIndicatorColor = Color.Red;
		axfpsPhuLuc_Sheet1.RowHeader.DefaultStyle.Parent = "RowHeaderDefaultEnhanced";
		axfpsPhuLuc_Sheet1.RowHeader.DefaultStyle.VerticalAlignment = CellVerticalAlignment.General;
		axfpsPhuLuc_Sheet1.SheetCornerStyle.HorizontalAlignment = CellHorizontalAlignment.General;
		axfpsPhuLuc_Sheet1.SheetCornerStyle.NoteIndicatorColor = Color.Red;
		axfpsPhuLuc_Sheet1.SheetCornerStyle.Parent = "CornerDefaultEnhanced";
		axfpsPhuLuc_Sheet1.SheetCornerStyle.VerticalAlignment = CellVerticalAlignment.General;
		axfpsPhuLuc_Sheet1.ReferenceStyle = ReferenceStyle.A1;
		((Control)chkChonPhuLucKeKhaiAll).AutoSize = true;
		((Control)chkChonPhuLucKeKhaiAll).Location = new Point(8, 5);
		((Control)chkChonPhuLucKeKhaiAll).Name = "chkChonPhuLucKeKhaiAll";
		((Control)chkChonPhuLucKeKhaiAll).Size = new Size(127, 17);
		((Control)chkChonPhuLucKeKhaiAll).TabIndex = 1;
		((Control)chkChonPhuLucKeKhaiAll).Text = "Chọn phụ lục kê khai";
		((ButtonBase)chkChonPhuLucKeKhaiAll).UseVisualStyleBackColor = true;
		chkChonPhuLucKeKhaiAll.CheckedChanged += chkChonPhuLucKeKhaiAll_CheckedChanged;
		((Control)chkChonPhuLucKeKhaiAll).Click += chkChonPhuLucKeKhaiAll_Click;
		((Control)f).Controls.Add((Control)(object)pnVT1);
		((Control)f).Controls.Add((Control)(object)pnVT2_PhuLucKeKhai);
		((Control)f).Controls.Add((Control)(object)pnVT3Button);
		((Control)f).Location = new Point(3, 25);
		((Control)f).Name = "f";
		((Control)f).Size = new Size(1080, 709);
		((Control)f).TabIndex = 54;
		((Control)pnThongBao).Controls.Add((Control)(object)mskNgayTb);
		((Control)pnThongBao).Controls.Add((Control)(object)lblNgayTb);
		((Control)pnThongBao).Controls.Add((Control)(object)txtSoTb);
		((Control)pnThongBao).Controls.Add((Control)(object)lblSoTB);
		((Control)pnThongBao).Location = new Point(396, 459);
		((Control)pnThongBao).Name = "pnThongBao";
		((Control)pnThongBao).Size = new Size(372, 59);
		((Control)pnThongBao).TabIndex = 164;
		((Control)pnThongBao).Visible = false;
		((Control)mskNgayTb).Location = new Point(177, 31);
		mskNgayTb.Mask = "00/00/0000";
		((Control)mskNgayTb).Name = "mskNgayTb";
		((Control)mskNgayTb).Size = new Size(90, 20);
		((Control)mskNgayTb).TabIndex = 57;
		mskNgayTb.TextAlign = (HorizontalAlignment)2;
		mskNgayTb.ValidatingType = typeof(DateTime);
		((Control)mskNgayTb).Leave += mskNgayTb_Leave;
		((Control)lblNgayTb).AutoSize = true;
		((Control)lblNgayTb).Location = new Point(5, 34);
		((Control)lblNgayTb).Name = "lblNgayTb";
		((Control)lblNgayTb).Size = new Size(166, 13);
		((Control)lblNgayTb).TabIndex = 53;
		((Control)lblNgayTb).Text = "Ngày thông báo giải trình bổ sung";
		((Control)txtSoTb).Location = new Point(175, 5);
		((TextBoxBase)txtSoTb).MaxLength = 12;
		((Control)txtSoTb).Name = "txtSoTb";
		((Control)txtSoTb).Size = new Size(92, 20);
		((Control)txtSoTb).TabIndex = 52;
		((Control)lblSoTB).AutoSize = true;
		((Control)lblSoTB).Location = new Point(4, 8);
		((Control)lblSoTB).Name = "lblSoTB";
		((Control)lblSoTB).Size = new Size(154, 13);
		((Control)lblSoTB).TabIndex = 51;
		((Control)lblSoTB).Text = "Số thông báo giải trình bổ sung";
		((Control)pnMst).Controls.Add((Control)(object)lblMst);
		((Control)pnMst).Controls.Add((Control)(object)txtMst);
		((Control)pnMst).Location = new Point(687, 457);
		((Control)pnMst).Name = "pnMst";
		((Control)pnMst).Size = new Size(356, 34);
		((Control)pnMst).TabIndex = 55;
		((Control)pnMst).Visible = false;
		((Control)lblMst).AutoSize = true;
		((Control)lblMst).Location = new Point(7, 10);
		((Control)lblMst).Name = "lblMst";
		((Control)lblMst).Size = new Size(232, 13);
		((Control)lblMst).TabIndex = 0;
		((Control)lblMst).Text = "MST Doanh nghiệp/ Đơn vị sự nghiệp công lập";
		((Control)txtMst).Location = new Point(250, 5);
		((Control)txtMst).Name = "txtMst";
		((Control)txtMst).Size = new Size(92, 20);
		((Control)txtMst).TabIndex = 1;
		((Control)txtMst).KeyPress += new KeyPressEventHandler(txtMst_KeyPress);
		pnVT1.BorderStyle = (BorderStyle)1;
		((Control)pnVT1).Controls.Add((Control)(object)lblNgayPsThuNhap);
		((Control)pnVT1).Controls.Add((Control)(object)pnThongBao);
		((Control)pnVT1).Controls.Add((Control)(object)pnMst);
		((Control)pnVT1).Controls.Add((Control)(object)pnLoaiTien);
		((Control)pnVT1).Controls.Add((Control)(object)pnRDLanDauBoXung1);
		((Control)pnVT1).Controls.Add((Control)(object)pnTruongHopQT);
		((Control)pnVT1).Controls.Add((Control)(object)pnLoaiKyHoan);
		((Control)pnVT1).Controls.Add((Control)(object)pnChuyenNhuong);
		((Control)pnVT1).Controls.Add((Control)(object)pnKhaiTheoPLuatDanSuThue);
		((Control)pnVT1).Controls.Add((Control)(object)pnTinhCQT);
		((Control)pnVT1).Controls.Add((Control)(object)pnMSTCaNhanToChuc);
		((Control)pnVT1).Controls.Add((Control)(object)pnPPhapNopThue);
		((Control)pnVT1).Controls.Add((Control)(object)pnVT1_1CheckBox);
		((Control)pnVT1).Controls.Add((Control)(object)pnVT1_4LanDauBoXung);
		((Control)pnVT1).Controls.Add((Control)(object)pnVT1_2NgayThangQuyNam);
		((Control)pnVT1).Controls.Add((Control)(object)pnVT1_5DanhMucNghe);
		((Control)pnVT1).Controls.Add((Control)(object)pnMaHSDNHT);
		((Control)pnVT1).Controls.Add((Control)(object)pnThuHo);
		((Control)pnVT1).Controls.Add((Control)(object)pnVT1_3TuNgayDenNgay);
		((Control)pnVT1).Location = new Point(0, 3);
		((Control)pnVT1).Name = "pnVT1";
		((Control)pnVT1).Size = new Size(1062, 552);
		((Control)pnVT1).TabIndex = 53;
		((Control)pnRDLanDauBoXung1).Controls.Add((Control)(object)pnNgayLap1);
		((Control)pnRDLanDauBoXung1).Controls.Add((Control)(object)pnLan1);
		((Control)pnRDLanDauBoXung1).Controls.Add((Control)(object)label4);
		((Control)pnRDLanDauBoXung1).Controls.Add((Control)(object)cboLoaiToKhai);
		((Control)pnRDLanDauBoXung1).Location = new Point(2, 32);
		((Control)pnRDLanDauBoXung1).Name = "pnRDLanDauBoXung1";
		((Control)pnRDLanDauBoXung1).Size = new Size(324, 62);
		((Control)pnRDLanDauBoXung1).TabIndex = 162;
		((Control)pnRDLanDauBoXung1).Visible = false;
		((Control)pnNgayLap1).Controls.Add((Control)(object)mskNgayLapKHBS1);
		((Control)pnNgayLap1).Controls.Add((Control)(object)lblNgayLapKHBS1);
		((Control)pnNgayLap1).Location = new Point(105, 32);
		((Control)pnNgayLap1).Name = "pnNgayLap1";
		((Control)pnNgayLap1).Size = new Size(171, 26);
		((Control)pnNgayLap1).TabIndex = 57;
		((Control)pnNgayLap1).Visible = false;
		((Control)mskNgayLapKHBS1).Location = new Point(87, 3);
		mskNgayLapKHBS1.Mask = "00/00/0000";
		((Control)mskNgayLapKHBS1).Name = "mskNgayLapKHBS1";
		((Control)mskNgayLapKHBS1).Size = new Size(70, 20);
		((Control)mskNgayLapKHBS1).TabIndex = 55;
		mskNgayLapKHBS1.TextAlign = (HorizontalAlignment)2;
		mskNgayLapKHBS1.ValidatingType = typeof(DateTime);
		((Control)mskNgayLapKHBS1).TextChanged += mskNgayLapKHBS1_TextChanged;
		((Control)lblNgayLapKHBS1).AutoSize = true;
		((Control)lblNgayLapKHBS1).Location = new Point(4, 6);
		((Control)lblNgayLapKHBS1).Name = "lblNgayLapKHBS1";
		((Control)lblNgayLapKHBS1).Size = new Size(81, 13);
		((Control)lblNgayLapKHBS1).TabIndex = 56;
		((Control)lblNgayLapKHBS1).Text = "Ngày lập KHBS";
		((Control)pnLan1).Controls.Add((Control)(object)txtLan1);
		((Control)pnLan1).Controls.Add((Control)(object)label5);
		((Control)pnLan1).Location = new Point(232, 3);
		((Control)pnLan1).Name = "pnLan1";
		((Control)pnLan1).Size = new Size(86, 25);
		((Control)pnLan1).TabIndex = 47;
		((Control)pnLan1).Visible = false;
		((Control)txtLan1).Location = new Point(35, 3);
		((TextBoxBase)txtLan1).MaxLength = 2;
		((Control)txtLan1).Name = "txtLan1";
		((Control)txtLan1).Size = new Size(30, 20);
		((Control)txtLan1).TabIndex = 52;
		((Control)txtLan1).TextChanged += txtLan1_TextChanged;
		((Control)label5).AutoSize = true;
		((Control)label5).Location = new Point(4, 8);
		((Control)label5).Name = "label5";
		((Control)label5).Size = new Size(25, 13);
		((Control)label5).TabIndex = 51;
		((Control)label5).Text = "Lần";
		((Control)label4).AutoSize = true;
		((Control)label4).Location = new Point(12, 8);
		((Control)label4).Name = "label4";
		((Control)label4).Size = new Size(62, 13);
		((Control)label4).TabIndex = 1;
		((Control)label4).Text = "Loại tờ khai";
		((ListControl)cboLoaiToKhai).FormattingEnabled = true;
		cboLoaiToKhai.Items.AddRange(new object[2] { "Lần đầu", "Thay thế" });
		((Control)cboLoaiToKhai).Location = new Point(105, 5);
		((Control)cboLoaiToKhai).Name = "cboLoaiToKhai";
		((Control)cboLoaiToKhai).Size = new Size(118, 21);
		((Control)cboLoaiToKhai).TabIndex = 0;
		cboLoaiToKhai.SelectedIndexChanged += cboLoaiToKhai_SelectedIndexChanged;
		((Control)cboLoaiToKhai).KeyPress += new KeyPressEventHandler(cboLoaiToKhai_KeyPress);
		((Control)pnTruongHopQT).Controls.Add((Control)(object)cboTruongHopQT);
		((Control)pnTruongHopQT).Controls.Add((Control)(object)label3);
		((Control)pnTruongHopQT).Location = new Point(597, 704);
		((Control)pnTruongHopQT).Margin = new Padding(4, 5, 4, 5);
		((Control)pnTruongHopQT).Name = "pnTruongHopQT";
		((Control)pnTruongHopQT).Size = new Size(372, 52);
		((Control)pnTruongHopQT).TabIndex = 161;
		((Control)pnTruongHopQT).Visible = false;
		((ListControl)cboTruongHopQT).FormattingEnabled = true;
		((Control)cboTruongHopQT).Location = new Point(14, 30);
		((Control)cboTruongHopQT).Margin = new Padding(4, 5, 4, 5);
		((Control)cboTruongHopQT).Name = "cboTruongHopQT";
		((Control)cboTruongHopQT).Size = new Size(340, 28);
		((Control)cboTruongHopQT).TabIndex = 1;
		cboTruongHopQT.SelectedIndexChanged += cboTruongHopQT_SelectedIndexChanged;
		((Control)cboTruongHopQT).KeyDown += new KeyEventHandler(cboTruongHopQT_KeyDown);
		((Control)cboTruongHopQT).KeyPress += new KeyPressEventHandler(cboTruongHopQT_KeyPress);
		((Control)label3).AutoSize = true;
		((Control)label3).Location = new Point(12, 7);
		((Control)label3).Margin = new Padding(4, 0, 4, 0);
		((Control)label3).Name = "label3";
		((Control)label3).Size = new Size(173, 20);
		((Control)label3).TabIndex = 0;
		((Control)label3).Text = "Trường hợp quyết toán:";
		((Control)pnLoaiKyHoan).Controls.Add((Control)(object)cboKieuKy);
		((Control)pnLoaiKyHoan).Controls.Add((Control)(object)label2);
		((Control)pnLoaiKyHoan).Location = new Point(411, 503);
		((Control)pnLoaiKyHoan).Name = "pnLoaiKyHoan";
		((Control)pnLoaiKyHoan).Size = new Size(349, 39);
		((Control)pnLoaiKyHoan).TabIndex = 160;
		((Control)pnLoaiKyHoan).Visible = false;
		((ListControl)cboKieuKy).FormattingEnabled = true;
		((Control)cboKieuKy).Location = new Point(67, 9);
		((Control)cboKieuKy).Name = "cboKieuKy";
		((Control)cboKieuKy).Size = new Size(275, 21);
		((Control)cboKieuKy).TabIndex = 1;
		cboKieuKy.SelectedIndexChanged += cboKieuKy_SelectedIndexChanged;
		((Control)cboKieuKy).KeyPress += new KeyPressEventHandler(cboKieuKy_KeyPress);
		((Control)label2).AutoSize = true;
		((Control)label2).Location = new Point(9, 11);
		((Control)label2).Name = "label2";
		((Control)label2).Size = new Size(42, 13);
		((Control)label2).TabIndex = 0;
		((Control)label2).Text = "Kiểu kỳ";
		((Control)pnChuyenNhuong).Controls.Add((Control)(object)chkDoanhNghiepChuyenNhuong);
		((Control)pnChuyenNhuong).Controls.Add((Control)(object)chkToChucCaNhanChuyenNhuong);
		((Control)pnChuyenNhuong).Location = new Point(398, 288);
		((Control)pnChuyenNhuong).Margin = new Padding(3, 0, 3, 0);
		((Control)pnChuyenNhuong).Name = "pnChuyenNhuong";
		((Control)pnChuyenNhuong).Size = new Size(372, 50);
		((Control)pnChuyenNhuong).TabIndex = 60;
		((Control)pnChuyenNhuong).Visible = false;
		((Control)chkDoanhNghiepChuyenNhuong).AutoSize = true;
		((Control)chkDoanhNghiepChuyenNhuong).Location = new Point(4, 27);
		((Control)chkDoanhNghiepChuyenNhuong).Name = "chkDoanhNghiepChuyenNhuong";
		((Control)chkDoanhNghiepChuyenNhuong).Size = new Size(234, 17);
		((Control)chkDoanhNghiepChuyenNhuong).TabIndex = 57;
		((Control)chkDoanhNghiepChuyenNhuong).Text = "Doanh nghiệp có vốn được chuyển nhượng";
		((ButtonBase)chkDoanhNghiepChuyenNhuong).UseVisualStyleBackColor = true;
		chkDoanhNghiepChuyenNhuong.CheckedChanged += chkDoanhNghiepChuyenNhuong_CheckedChanged;
		((Control)chkToChucCaNhanChuyenNhuong).AutoSize = true;
		((Control)chkToChucCaNhanChuyenNhuong).Location = new Point(4, 5);
		((Control)chkToChucCaNhanChuyenNhuong).Name = "chkToChucCaNhanChuyenNhuong";
		((Control)chkToChucCaNhanChuyenNhuong).Size = new Size(235, 17);
		((Control)chkToChucCaNhanChuyenNhuong).TabIndex = 56;
		((Control)chkToChucCaNhanChuyenNhuong).Text = "Tổ chức/cá nhân nhận vốn chuyển nhượng";
		((ButtonBase)chkToChucCaNhanChuyenNhuong).UseVisualStyleBackColor = true;
		chkToChucCaNhanChuyenNhuong.CheckedChanged += chkToChucCaNhanChuyenNhuong_CheckedChanged;
		((Control)pnKhaiTheoPLuatDanSuThue).Controls.Add((Control)(object)chkKhaiTheoPLuatThue);
		((Control)pnKhaiTheoPLuatDanSuThue).Controls.Add((Control)(object)chkKhaiTheoPLuatDanSu);
		((Control)pnKhaiTheoPLuatDanSuThue).Location = new Point(473, 15);
		((Control)pnKhaiTheoPLuatDanSuThue).Margin = new Padding(3, 0, 3, 0);
		((Control)pnKhaiTheoPLuatDanSuThue).Name = "pnKhaiTheoPLuatDanSuThue";
		((Control)pnKhaiTheoPLuatDanSuThue).Size = new Size(364, 88);
		((Control)pnKhaiTheoPLuatDanSuThue).TabIndex = 60;
		((Control)pnKhaiTheoPLuatDanSuThue).Visible = false;
		((Control)chkKhaiTheoPLuatThue).Location = new Point(3, 55);
		((Control)chkKhaiTheoPLuatThue).Name = "chkKhaiTheoPLuatThue";
		((Control)chkKhaiTheoPLuatThue).Size = new Size(333, 33);
		((Control)chkKhaiTheoPLuatThue).TabIndex = 57;
		((Control)chkKhaiTheoPLuatThue).Text = "Doanh nghiệp, tổ chức kinh tế khai thuế thay, nộp thuế thay theo pháp luật thuế";
		((ButtonBase)chkKhaiTheoPLuatThue).UseVisualStyleBackColor = true;
		chkKhaiTheoPLuatThue.CheckedChanged += CheckedChanged;
		((Control)chkKhaiTheoPLuatDanSu).Location = new Point(4, 2);
		((Control)chkKhaiTheoPLuatDanSu).Name = "chkKhaiTheoPLuatDanSu";
		((Control)chkKhaiTheoPLuatDanSu).Size = new Size(357, 55);
		((Control)chkKhaiTheoPLuatDanSu).TabIndex = 56;
		((Control)chkKhaiTheoPLuatDanSu).Text = "Cá nhân cho thuê tài sản trực tiếp khai thuế/ Tổ chức, cá nhân khai thuế thay, nộp thuế thay cho cá nhân ủy quyền theo quy định của pháp luật dân sự";
		((ButtonBase)chkKhaiTheoPLuatDanSu).UseVisualStyleBackColor = true;
		chkKhaiTheoPLuatDanSu.CheckedChanged += CheckedChanged;
		((Control)pnTinhCQT).Controls.Add((Control)(object)panel1);
		((Control)pnTinhCQT).Controls.Add((Control)(object)pnCCT);
		((Control)pnTinhCQT).Controls.Add((Control)(object)pnCQTCuc);
		((Control)pnTinhCQT).Controls.Add((Control)(object)pnTinh);
		((Control)pnTinhCQT).Location = new Point(398, 338);
		((Control)pnTinhCQT).Margin = new Padding(3, 0, 3, 0);
		((Control)pnTinhCQT).Name = "pnTinhCQT";
		((Control)pnTinhCQT).Size = new Size(372, 141);
		((Control)pnTinhCQT).TabIndex = 60;
		((Control)pnTinhCQT).Visible = false;
		((Control)pnCCT).Controls.Add((Control)(object)cboCQTChiCuc);
		((Control)pnCCT).Controls.Add((Control)(object)lblCQTChiCuc);
		((Control)pnCCT).Location = new Point(4, 91);
		((Control)pnCCT).Name = "pnCCT";
		((Control)pnCCT).Size = new Size(358, 25);
		((Control)pnCCT).TabIndex = 11;
		cboCQTChiCuc.DropDownWidth = 500;
		((ListControl)cboCQTChiCuc).FormattingEnabled = true;
		((Control)cboCQTChiCuc).Location = new Point(108, 3);
		((Control)cboCQTChiCuc).Name = "cboCQTChiCuc";
		((Control)cboCQTChiCuc).Size = new Size(238, 21);
		((Control)cboCQTChiCuc).TabIndex = 1;
		((Control)cboCQTChiCuc).Text = "Cơ quan thuế chi cục";
		cboCQTChiCuc.SelectedIndexChanged += cboCQTChiCuc_SelectedIndexChanged;
		((Control)cboCQTChiCuc).KeyDown += new KeyEventHandler(cboCQTChiCuc_KeyDown);
		((Control)cboCQTChiCuc).KeyPress += new KeyPressEventHandler(cboCQTChiCuc_KeyPress);
		((Control)lblCQTChiCuc).AutoSize = true;
		((Control)lblCQTChiCuc).Location = new Point(3, 5);
		((Control)lblCQTChiCuc).Name = "lblCQTChiCuc";
		((Control)lblCQTChiCuc).Size = new Size(67, 13);
		((Control)lblCQTChiCuc).TabIndex = 5;
		((Control)lblCQTChiCuc).Text = "CQT nơi nộp";
		((Control)pnCQTCuc).Controls.Add((Control)(object)cboCQTCuc);
		((Control)pnCQTCuc).Controls.Add((Control)(object)lblCQTCuc);
		((Control)pnCQTCuc).Location = new Point(4, 57);
		((Control)pnCQTCuc).Name = "pnCQTCuc";
		((Control)pnCQTCuc).Size = new Size(358, 31);
		((Control)pnCQTCuc).TabIndex = 10;
		((ListControl)cboCQTCuc).FormattingEnabled = true;
		((Control)cboCQTCuc).Location = new Point(108, 3);
		((Control)cboCQTCuc).Name = "cboCQTCuc";
		((Control)cboCQTCuc).Size = new Size(238, 21);
		((Control)cboCQTCuc).TabIndex = 0;
		cboCQTCuc.SelectedIndexChanged += cboCQTCuc_SelectedIndexChanged;
		((Control)cboCQTCuc).KeyDown += new KeyEventHandler(cboCQTCuc_KeyDown);
		((Control)cboCQTCuc).KeyPress += new KeyPressEventHandler(cboCQTCuc_KeyPress);
		((Control)lblCQTCuc).AutoSize = true;
		((Control)lblCQTCuc).Location = new Point(2, 5);
		((Control)lblCQTCuc).Name = "lblCQTCuc";
		((Control)lblCQTCuc).Size = new Size(78, 13);
		((Control)lblCQTCuc).TabIndex = 4;
		((Control)lblCQTCuc).Text = "Thuế tỉnh/CCT";
		((Control)pnTinh).Controls.Add((Control)(object)cboTinh);
		((Control)pnTinh).Controls.Add((Control)(object)lblTinh);
		((Control)pnTinh).Location = new Point(4, 6);
		((Control)pnTinh).Name = "pnTinh";
		((Control)pnTinh).Size = new Size(358, 25);
		((Control)pnTinh).TabIndex = 8;
		((ListControl)cboTinh).FormattingEnabled = true;
		((Control)cboTinh).Location = new Point(173, 3);
		((Control)cboTinh).Name = "cboTinh";
		((Control)cboTinh).Size = new Size(173, 21);
		((Control)cboTinh).TabIndex = 2;
		((Control)cboTinh).Text = "Chọn tỉnh";
		((Control)cboTinh).Visible = false;
		cboTinh.SelectedIndexChanged += cboTinh_SelectedIndexChanged;
		((Control)cboTinh).KeyDown += new KeyEventHandler(cboTinh_KeyDown);
		((Control)cboTinh).KeyPress += new KeyPressEventHandler(cboTinh_KeyPress);
		((Control)lblTinh).AutoSize = true;
		((Control)lblTinh).Location = new Point(3, 6);
		((Control)lblTinh).Name = "lblTinh";
		((Control)lblTinh).Size = new Size(28, 13);
		((Control)lblTinh).TabIndex = 3;
		((Control)lblTinh).Text = "Tỉnh";
		((Control)pnMSTCaNhanToChuc).Controls.Add((Control)(object)chkMSTToChuc);
		((Control)pnMSTCaNhanToChuc).Controls.Add((Control)(object)chkMSTCaNhan);
		((Control)pnMSTCaNhanToChuc).Location = new Point(398, 249);
		((Control)pnMSTCaNhanToChuc).Margin = new Padding(3, 0, 3, 0);
		((Control)pnMSTCaNhanToChuc).Name = "pnMSTCaNhanToChuc";
		((Control)pnMSTCaNhanToChuc).Size = new Size(372, 26);
		((Control)pnMSTCaNhanToChuc).TabIndex = 59;
		((Control)pnMSTCaNhanToChuc).Visible = false;
		((Control)chkMSTToChuc).AutoSize = true;
		((Control)chkMSTToChuc).Location = new Point(179, 5);
		((Control)chkMSTToChuc).Name = "chkMSTToChuc";
		((Control)chkMSTToChuc).Size = new Size(118, 17);
		((Control)chkMSTToChuc).TabIndex = 57;
		((Control)chkMSTToChuc).Text = "Mã số thuế tổ chức";
		((ButtonBase)chkMSTToChuc).UseVisualStyleBackColor = true;
		chkMSTToChuc.CheckedChanged += CheckedChanged;
		((Control)chkMSTCaNhan).AutoSize = true;
		((Control)chkMSTCaNhan).Location = new Point(4, 5);
		((Control)chkMSTCaNhan).Name = "chkMSTCaNhan";
		((Control)chkMSTCaNhan).Size = new Size(121, 17);
		((Control)chkMSTCaNhan).TabIndex = 56;
		((Control)chkMSTCaNhan).Text = "Mã số thuế cá nhân";
		((ButtonBase)chkMSTCaNhan).UseVisualStyleBackColor = true;
		chkMSTCaNhan.CheckedChanged += CheckedChanged;
		((Control)pnPPhapNopThue).Location = new Point(3, -6);
		((Control)pnPPhapNopThue).Margin = new Padding(3, 0, 3, 0);
		((Control)pnPPhapNopThue).Name = "pnPPhapNopThue";
		((Control)pnPPhapNopThue).Size = new Size(350, 34);
		((Control)pnPPhapNopThue).TabIndex = 59;
		((Control)pnPPhapNopThue).Visible = false;
		((Control)pnVT1_1CheckBox).Controls.Add((Control)(object)chkTKLanPS);
		((Control)pnVT1_1CheckBox).Controls.Add((Control)(object)ckNam);
		((Control)pnVT1_1CheckBox).Controls.Add((Control)(object)chkTKhaiLanXB);
		((Control)pnVT1_1CheckBox).Controls.Add((Control)(object)chkTKKy);
		((Control)pnVT1_1CheckBox).Controls.Add((Control)(object)chkTKQuy);
		((Control)pnVT1_1CheckBox).Controls.Add((Control)(object)chkTuThangDenThang);
		((Control)pnVT1_1CheckBox).Controls.Add((Control)(object)chkTuQuyDenQuy);
		((Control)pnVT1_1CheckBox).Controls.Add((Control)(object)chkTkhaiThang);
		((Control)pnVT1_1CheckBox).Location = new Point(2, 126);
		((Control)pnVT1_1CheckBox).Margin = new Padding(3, 0, 3, 0);
		((Control)pnVT1_1CheckBox).Name = "pnVT1_1CheckBox";
		((Control)pnVT1_1CheckBox).Size = new Size(920, 23);
		((Control)pnVT1_1CheckBox).TabIndex = 2;
		((Control)chkTKLanPS).AutoSize = true;
		((Control)chkTKLanPS).Location = new Point(633, 3);
		((Control)chkTKLanPS).Name = "chkTKLanPS";
		((Control)chkTKLanPS).Size = new Size(125, 17);
		((Control)chkTKLanPS).TabIndex = 57;
		((Control)chkTKLanPS).Text = "Tờ khai lần phát sinh";
		((ButtonBase)chkTKLanPS).UseVisualStyleBackColor = true;
		chkTKLanPS.CheckedChanged += CheckedChanged;
		((Control)ckNam).AutoSize = true;
		((Control)ckNam).Location = new Point(110, 4);
		((Control)ckNam).Name = "ckNam";
		((Control)ckNam).Size = new Size(87, 17);
		((Control)ckNam).TabIndex = 56;
		((Control)ckNam).Text = "Tờ khai Năm";
		((ButtonBase)ckNam).UseVisualStyleBackColor = true;
		ckNam.CheckedChanged += CheckedChanged;
		((Control)chkTKhaiLanXB).AutoSize = true;
		((Control)chkTKhaiLanXB).Location = new Point(9, 3);
		((Control)chkTKhaiLanXB).Name = "chkTKhaiLanXB";
		((Control)chkTKhaiLanXB).Size = new Size(101, 17);
		((Control)chkTKhaiLanXB).TabIndex = 42;
		((Control)chkTKhaiLanXB).Text = "TK lần xuất bán";
		((ButtonBase)chkTKhaiLanXB).UseVisualStyleBackColor = true;
		chkTKhaiLanXB.CheckedChanged += CheckedChanged;
		((Control)chkTKKy).AutoSize = true;
		((Control)chkTKKy).Location = new Point(199, 4);
		((Control)chkTKKy).Name = "chkTKKy";
		((Control)chkTKKy).Size = new Size(76, 17);
		((Control)chkTKKy).TabIndex = 41;
		((Control)chkTKKy).Text = "Tờ khai kỳ";
		((ButtonBase)chkTKKy).UseVisualStyleBackColor = true;
		chkTKKy.CheckedChanged += CheckedChanged;
		((Control)chkTKQuy).AutoSize = true;
		((Control)chkTKQuy).Location = new Point(276, 4);
		((Control)chkTKQuy).Name = "chkTKQuy";
		((Control)chkTKQuy).Size = new Size(82, 17);
		((Control)chkTKQuy).TabIndex = 40;
		((Control)chkTKQuy).Text = "Tờ khai quý";
		((ButtonBase)chkTKQuy).UseVisualStyleBackColor = true;
		chkTKQuy.CheckedChanged += CheckedChanged;
		((Control)chkTuThangDenThang).AutoSize = true;
		((Control)chkTuThangDenThang).Location = new Point(465, 4);
		((Control)chkTuThangDenThang).Name = "chkTuThangDenThang";
		((Control)chkTuThangDenThang).Size = new Size(156, 17);
		((Control)chkTuThangDenThang).TabIndex = 39;
		((Control)chkTuThangDenThang).Text = "Tờ khai từ tháng đến tháng";
		((ButtonBase)chkTuThangDenThang).UseVisualStyleBackColor = true;
		chkTuThangDenThang.CheckedChanged += CheckedChanged;
		((Control)chkTuQuyDenQuy).AutoSize = true;
		((Control)chkTuQuyDenQuy).Location = new Point(781, 3);
		((Control)chkTuQuyDenQuy).Name = "chkTuQuyDenQuy";
		((Control)chkTuQuyDenQuy).Size = new Size(136, 17);
		((Control)chkTuQuyDenQuy).TabIndex = 58;
		((Control)chkTuQuyDenQuy).Text = "Tờ khai từ quý đến quý";
		((ButtonBase)chkTuQuyDenQuy).UseVisualStyleBackColor = true;
		chkTuQuyDenQuy.CheckedChanged += CheckedChanged;
		((Control)chkTkhaiThang).AutoSize = true;
		((Control)chkTkhaiThang).Location = new Point(363, 3);
		((Control)chkTkhaiThang).Name = "chkTkhaiThang";
		((Control)chkTkhaiThang).Size = new Size(92, 17);
		((Control)chkTkhaiThang).TabIndex = 37;
		((Control)chkTkhaiThang).Text = "Tờ khai tháng";
		((ButtonBase)chkTkhaiThang).UseVisualStyleBackColor = true;
		chkTkhaiThang.CheckedChanged += CheckedChanged;
		((Control)pnVT1_4LanDauBoXung).Controls.Add((Control)(object)pnQuyetToan_LienTuc);
		((Control)pnVT1_4LanDauBoXung).Controls.Add((Control)(object)pnNgayLap);
		((Control)pnVT1_4LanDauBoXung).Controls.Add((Control)(object)pnRDLanDauBoXung);
		((Control)pnVT1_4LanDauBoXung).Location = new Point(2, 242);
		((Control)pnVT1_4LanDauBoXung).Name = "pnVT1_4LanDauBoXung";
		((Control)pnVT1_4LanDauBoXung).Size = new Size(372, 135);
		((Control)pnVT1_4LanDauBoXung).TabIndex = 51;
		((Control)pnQuyetToan_LienTuc).Controls.Add((Control)(object)cbxCTDC);
		((Control)pnQuyetToan_LienTuc).Controls.Add((Control)(object)chkQTTungNam_HoatDongLienTuc);
		((Control)pnQuyetToan_LienTuc).Controls.Add((Control)(object)lbNgayPS);
		((Control)pnQuyetToan_LienTuc).Controls.Add((Control)(object)cbxLNDC);
		((Control)pnQuyetToan_LienTuc).Controls.Add((Control)(object)mskNgayPS);
		((Control)pnQuyetToan_LienTuc).Controls.Add((Control)(object)chk_CoSoThanhLap);
		((Control)pnQuyetToan_LienTuc).Controls.Add((Control)(object)chkQTNamDau_KhongHoatDongLienTuc);
		((Control)pnQuyetToan_LienTuc).ImeMode = (ImeMode)1;
		((Control)pnQuyetToan_LienTuc).Location = new Point(3, 0);
		((Control)pnQuyetToan_LienTuc).Name = "pnQuyetToan_LienTuc";
		((Control)pnQuyetToan_LienTuc).Size = new Size(321, 49);
		((Control)pnQuyetToan_LienTuc).TabIndex = 55;
		((Control)pnQuyetToan_LienTuc).Visible = false;
		((Control)cbxCTDC).AutoSize = true;
		cbxCTDC.Checked = true;
		cbxCTDC.CheckState = (CheckState)1;
		((Control)cbxCTDC).Location = new Point(7, 5);
		((Control)cbxCTDC).Name = "cbxCTDC";
		((Control)cbxCTDC).Size = new Size(108, 17);
		((Control)cbxCTDC).TabIndex = 58;
		((Control)cbxCTDC).Text = "Cổ tức được chia";
		((ButtonBase)cbxCTDC).UseVisualStyleBackColor = true;
		((Control)cbxCTDC).Visible = false;
		cbxCTDC.CheckedChanged += CheckedQTChanged;
		((Control)chkQTTungNam_HoatDongLienTuc).AutoSize = true;
		chkQTTungNam_HoatDongLienTuc.Checked = true;
		chkQTTungNam_HoatDongLienTuc.CheckState = (CheckState)1;
		((Control)chkQTTungNam_HoatDongLienTuc).Location = new Point(7, 5);
		((Control)chkQTTungNam_HoatDongLienTuc).Name = "chkQTTungNam_HoatDongLienTuc";
		((Control)chkQTTungNam_HoatDongLienTuc).Size = new Size(266, 17);
		((Control)chkQTTungNam_HoatDongLienTuc).TabIndex = 55;
		((Control)chkQTTungNam_HoatDongLienTuc).Text = "Doanh nghiệp đáp ứng giả định hoạt động liên tục";
		((ButtonBase)chkQTTungNam_HoatDongLienTuc).UseVisualStyleBackColor = true;
		chkQTTungNam_HoatDongLienTuc.CheckedChanged += CheckedQTChanged;
		((Control)lbNgayPS).AutoSize = true;
		((Control)lbNgayPS).Location = new Point(146, 7);
		((Control)lbNgayPS).Name = "lbNgayPS";
		((Control)lbNgayPS).Size = new Size(78, 13);
		((Control)lbNgayPS).TabIndex = 57;
		((Control)lbNgayPS).Text = "Ngày phát sinh";
		((Control)lbNgayPS).Visible = false;
		((Control)cbxLNDC).AutoSize = true;
		((Control)cbxLNDC).Location = new Point(194, 5);
		((Control)cbxLNDC).Name = "cbxLNDC";
		((Control)cbxLNDC).Size = new Size(124, 17);
		((Control)cbxLNDC).TabIndex = 59;
		((Control)cbxLNDC).Text = "Lợi nhuận được chia";
		((ButtonBase)cbxLNDC).UseVisualStyleBackColor = true;
		((Control)cbxLNDC).Visible = false;
		cbxLNDC.CheckedChanged += CheckedQTChanged;
		((Control)mskNgayPS).Location = new Point(229, 4);
		mskNgayPS.Mask = "00/00/0000";
		((Control)mskNgayPS).Name = "mskNgayPS";
		((Control)mskNgayPS).Size = new Size(70, 20);
		((Control)mskNgayPS).TabIndex = 57;
		mskNgayPS.TextAlign = (HorizontalAlignment)2;
		mskNgayPS.ValidatingType = typeof(DateTime);
		((Control)mskNgayPS).Visible = false;
		((Control)mskNgayPS).Leave += mskNgayPS_Leave;
		((Control)chk_CoSoThanhLap).AutoSize = true;
		((Control)chk_CoSoThanhLap).Location = new Point(8, 6);
		((Control)chk_CoSoThanhLap).Name = "chk_CoSoThanhLap";
		((Control)chk_CoSoThanhLap).Size = new Size(119, 17);
		((Control)chk_CoSoThanhLap).TabIndex = 57;
		((Control)chk_CoSoThanhLap).Text = "Cơ sở mới thành lập";
		((ButtonBase)chk_CoSoThanhLap).UseVisualStyleBackColor = true;
		((Control)chk_CoSoThanhLap).Visible = false;
		chk_CoSoThanhLap.CheckedChanged += chk_CoSoThanhLap_CheckedChanged;
		((Control)chkQTNamDau_KhongHoatDongLienTuc).AutoSize = true;
		((Control)chkQTNamDau_KhongHoatDongLienTuc).Location = new Point(7, 26);
		((Control)chkQTNamDau_KhongHoatDongLienTuc).Name = "chkQTNamDau_KhongHoatDongLienTuc";
		((Control)chkQTNamDau_KhongHoatDongLienTuc).Size = new Size(299, 17);
		((Control)chkQTNamDau_KhongHoatDongLienTuc).TabIndex = 56;
		((Control)chkQTNamDau_KhongHoatDongLienTuc).Text = "Doanh nghiệp không đáp ứng giả định hoạt động liên tục";
		((ButtonBase)chkQTNamDau_KhongHoatDongLienTuc).UseVisualStyleBackColor = true;
		chkQTNamDau_KhongHoatDongLienTuc.CheckedChanged += CheckedQTChanged;
		((Control)pnNgayLap).Controls.Add((Control)(object)mskNgayLapKHBS);
		((Control)pnNgayLap).Controls.Add((Control)(object)lblNgayLapKHBS);
		((Control)pnNgayLap).Location = new Point(115, 105);
		((Control)pnNgayLap).Name = "pnNgayLap";
		((Control)pnNgayLap).Size = new Size(171, 26);
		((Control)pnNgayLap).TabIndex = 56;
		((Control)pnNgayLap).Visible = false;
		((Control)mskNgayLapKHBS).Location = new Point(87, 3);
		mskNgayLapKHBS.Mask = "00/00/0000";
		((Control)mskNgayLapKHBS).Name = "mskNgayLapKHBS";
		((Control)mskNgayLapKHBS).Size = new Size(70, 20);
		((Control)mskNgayLapKHBS).TabIndex = 55;
		mskNgayLapKHBS.TextAlign = (HorizontalAlignment)2;
		mskNgayLapKHBS.ValidatingType = typeof(DateTime);
		((Control)mskNgayLapKHBS).Leave += mskNgayLapKHBS_Leave;
		((Control)lblNgayLapKHBS).AutoSize = true;
		((Control)lblNgayLapKHBS).Location = new Point(4, 6);
		((Control)lblNgayLapKHBS).Name = "lblNgayLapKHBS";
		((Control)lblNgayLapKHBS).Size = new Size(81, 13);
		((Control)lblNgayLapKHBS).TabIndex = 56;
		((Control)lblNgayLapKHBS).Text = "Ngày lập KHBS";
		((Control)pnRDLanDauBoXung).Controls.Add((Control)(object)OptBosung);
		((Control)pnRDLanDauBoXung).Controls.Add((Control)(object)pnLan);
		((Control)pnRDLanDauBoXung).Controls.Add((Control)(object)OptChinhthuc);
		((Control)pnRDLanDauBoXung).Location = new Point(3, 52);
		((Control)pnRDLanDauBoXung).Name = "pnRDLanDauBoXung";
		((Control)pnRDLanDauBoXung).Size = new Size(321, 48);
		((Control)pnRDLanDauBoXung).TabIndex = 52;
		((Control)OptBosung).AutoSize = true;
		((Control)OptBosung).Location = new Point(7, 26);
		((Control)OptBosung).Name = "OptBosung";
		((Control)OptBosung).Size = new Size(102, 17);
		((Control)OptBosung).TabIndex = 50;
		((Control)OptBosung).Text = "Tờ khai bổ sung";
		((ButtonBase)OptBosung).UseVisualStyleBackColor = true;
		OptBosung.CheckedChanged += OptBosung_CheckedChanged;
		((Control)pnLan).Controls.Add((Control)(object)txtLan);
		((Control)pnLan).Controls.Add((Control)(object)lblLan);
		((Control)pnLan).Location = new Point(163, 21);
		((Control)pnLan).Name = "pnLan";
		((Control)pnLan).Size = new Size(80, 25);
		((Control)pnLan).TabIndex = 46;
		((Control)pnLan).Visible = false;
		((Control)txtLan).Location = new Point(35, 3);
		((TextBoxBase)txtLan).MaxLength = 2;
		((Control)txtLan).Name = "txtLan";
		((Control)txtLan).Size = new Size(30, 20);
		((Control)txtLan).TabIndex = 52;
		((Control)txtLan).TextChanged += txtLan_TextChanged;
		((Control)txtLan).KeyPress += new KeyPressEventHandler(txtLan_KeyPress);
		((Control)txtLan).Leave += txtLan_Leave;
		((Control)lblLan).AutoSize = true;
		((Control)lblLan).Location = new Point(4, 8);
		((Control)lblLan).Name = "lblLan";
		((Control)lblLan).Size = new Size(25, 13);
		((Control)lblLan).TabIndex = 51;
		((Control)lblLan).Text = "Lần";
		((Control)OptChinhthuc).AutoSize = true;
		OptChinhthuc.Checked = true;
		((Control)OptChinhthuc).Location = new Point(7, 3);
		((Control)OptChinhthuc).Name = "OptChinhthuc";
		((Control)OptChinhthuc).Size = new Size(100, 17);
		((Control)OptChinhthuc).TabIndex = 41;
		OptChinhthuc.TabStop = true;
		((Control)OptChinhthuc).Text = "Tờ khai lần đầu";
		((ButtonBase)OptChinhthuc).UseVisualStyleBackColor = true;
		OptChinhthuc.CheckedChanged += OptChinhthuc_CheckedChanged;
		((Control)pnVT1_2NgayThangQuyNam).Controls.Add((Control)(object)pnKy);
		((Control)pnVT1_2NgayThangQuyNam).Controls.Add((Control)(object)pnNam);
		((Control)pnVT1_2NgayThangQuyNam).Controls.Add((Control)(object)pnNgay);
		((Control)pnVT1_2NgayThangQuyNam).Controls.Add((Control)(object)pnQuy);
		((Control)pnVT1_2NgayThangQuyNam).Controls.Add((Control)(object)pnThang);
		((Control)pnVT1_2NgayThangQuyNam).Location = new Point(3, 159);
		((Control)pnVT1_2NgayThangQuyNam).Margin = new Padding(3, 0, 3, 0);
		((Control)pnVT1_2NgayThangQuyNam).Name = "pnVT1_2NgayThangQuyNam";
		((Control)pnVT1_2NgayThangQuyNam).Size = new Size(768, 38);
		((Control)pnVT1_2NgayThangQuyNam).TabIndex = 7;
		((Control)pnKy).Controls.Add((Control)(object)cbbKy);
		((Control)pnKy).Controls.Add((Control)(object)lblKy);
		((Control)pnKy).Location = new Point(276, 2);
		((Control)pnKy).Name = "pnKy";
		((Control)pnKy).Size = new Size(74, 30);
		((Control)pnKy).TabIndex = 47;
		((ListControl)cbbKy).FormattingEnabled = true;
		cbbKy.Items.AddRange(new object[2] { "1", "2" });
		((Control)cbbKy).Location = new Point(24, 3);
		((Control)cbbKy).Name = "cbbKy";
		((Control)cbbKy).Size = new Size(40, 21);
		((Control)cbbKy).TabIndex = 44;
		cbbKy.SelectedIndexChanged += cbbKy_SelectedIndexChanged;
		((Control)cbbKy).KeyPress += new KeyPressEventHandler(cbbKy_KeyPress);
		((Control)lblKy).AutoSize = true;
		((Control)lblKy).Location = new Point(3, 6);
		((Control)lblKy).Name = "lblKy";
		((Control)lblKy).Size = new Size(19, 13);
		((Control)lblKy).TabIndex = 43;
		((Control)lblKy).Text = "Kỳ";
		((Control)pnNam).Controls.Add((Control)(object)mskYear);
		((Control)pnNam).Controls.Add((Control)(object)lblYear);
		((Control)pnNam).Location = new Point(575, 2);
		((Control)pnNam).Name = "pnNam";
		((Control)pnNam).Size = new Size(96, 30);
		((Control)pnNam).TabIndex = 47;
		((Control)mskYear).Location = new Point(36, 3);
		mskYear.Mask = "0000";
		((Control)mskYear).Name = "mskYear";
		((Control)mskYear).Size = new Size(50, 20);
		((Control)mskYear).TabIndex = 54;
		mskYear.TextAlign = (HorizontalAlignment)2;
		((Control)mskYear).Leave += mskYear_Leave;
		((Control)lblYear).AutoSize = true;
		((Control)lblYear).Location = new Point(3, 6);
		((Control)lblYear).Name = "lblYear";
		((Control)lblYear).Size = new Size(29, 13);
		((Control)lblYear).TabIndex = 17;
		((Control)lblYear).Text = "Năm";
		((Control)pnNgay).Controls.Add((Control)(object)lblNgay);
		((Control)pnNgay).Controls.Add((Control)(object)mskDay);
		((Control)pnNgay).Location = new Point(5, 2);
		((Control)pnNgay).Name = "pnNgay";
		((Control)pnNgay).Size = new Size(83, 30);
		((Control)pnNgay).TabIndex = 45;
		((Control)lblNgay).AutoSize = true;
		((Control)lblNgay).Location = new Point(3, 6);
		((Control)lblNgay).Name = "lblNgay";
		((Control)lblNgay).Size = new Size(32, 13);
		((Control)lblNgay).TabIndex = 11;
		((Control)lblNgay).Text = "Ngày";
		((Control)mskDay).Location = new Point(35, 3);
		mskDay.Mask = "00";
		((Control)mskDay).Name = "mskDay";
		((Control)mskDay).Size = new Size(40, 20);
		((Control)mskDay).TabIndex = 52;
		mskDay.TextAlign = (HorizontalAlignment)2;
		((Control)mskDay).Leave += mskDay_Leave;
		((Control)pnQuy).Controls.Add((Control)(object)cbbQuy);
		((Control)pnQuy).Controls.Add((Control)(object)lblQuy);
		((Control)pnQuy).Location = new Point(443, 2);
		((Control)pnQuy).Name = "pnQuy";
		((Control)pnQuy).Size = new Size(87, 30);
		((Control)pnQuy).TabIndex = 46;
		((Control)cbbQuy).BackColor = SystemColors.Window;
		((ListControl)cbbQuy).FormattingEnabled = true;
		cbbQuy.Items.AddRange(new object[4] { "1", "2", "3", "4" });
		((Control)cbbQuy).Location = new Point(32, 3);
		((Control)cbbQuy).Name = "cbbQuy";
		((Control)cbbQuy).Size = new Size(40, 21);
		((Control)cbbQuy).TabIndex = 44;
		cbbQuy.SelectedIndexChanged += cbbQuy_SelectedIndexChanged;
		((Control)cbbQuy).KeyPress += new KeyPressEventHandler(cbbQuy_KeyPress);
		((Control)lblQuy).AutoSize = true;
		((Control)lblQuy).Location = new Point(3, 6);
		((Control)lblQuy).Name = "lblQuy";
		((Control)lblQuy).Size = new Size(26, 13);
		((Control)lblQuy).TabIndex = 43;
		((Control)lblQuy).Text = "Quý";
		((Control)pnThang).Controls.Add((Control)(object)mskMonth);
		((Control)pnThang).Controls.Add((Control)(object)lblMonth);
		((Control)pnThang).Location = new Point(139, 2);
		((Control)pnThang).Name = "pnThang";
		((Control)pnThang).Size = new Size(90, 30);
		((Control)pnThang).TabIndex = 46;
		((Control)mskMonth).Location = new Point(43, 3);
		mskMonth.Mask = "00";
		((Control)mskMonth).Name = "mskMonth";
		((Control)mskMonth).Size = new Size(40, 20);
		((Control)mskMonth).TabIndex = 53;
		mskMonth.TextAlign = (HorizontalAlignment)2;
		((Control)mskMonth).Leave += mskMonth_Leave;
		((Control)lblMonth).AutoSize = true;
		((Control)lblMonth).Location = new Point(3, 6);
		((Control)lblMonth).Name = "lblMonth";
		((Control)lblMonth).Size = new Size(38, 13);
		((Control)lblMonth).TabIndex = 16;
		((Control)lblMonth).Text = "Tháng";
		((Control)pnVT1_5DanhMucNghe).Controls.Add((Control)(object)lblNganhKD);
		((Control)pnVT1_5DanhMucNghe).Controls.Add((Control)(object)cboNganhKD);
		((Control)pnVT1_5DanhMucNghe).Location = new Point(2, 383);
		((Control)pnVT1_5DanhMucNghe).Name = "pnVT1_5DanhMucNghe";
		((Control)pnVT1_5DanhMucNghe).Size = new Size(372, 43);
		((Control)pnVT1_5DanhMucNghe).TabIndex = 51;
		((Control)lblNganhKD).AutoSize = true;
		((Control)lblNganhKD).Location = new Point(1, 5);
		((Control)lblNganhKD).Name = "lblNganhKD";
		((Control)lblNganhKD).Size = new Size(116, 13);
		((Control)lblNganhKD).TabIndex = 23;
		((Control)lblNganhKD).Text = "Danh mục ngành nghề";
		cboNganhKD.DropDownWidth = 340;
		((ListControl)cboNganhKD).FormattingEnabled = true;
		((Control)cboNganhKD).Location = new Point(5, 20);
		((Control)cboNganhKD).Name = "cboNganhKD";
		((Control)cboNganhKD).Size = new Size(340, 21);
		((Control)cboNganhKD).TabIndex = 24;
		cboNganhKD.SelectedIndexChanged += cboNganhKD_SelectedIndexChanged;
		((Control)cboNganhKD).KeyDown += new KeyEventHandler(cboNganhKD_KeyDown);
		((Control)cboNganhKD).KeyPress += new KeyPressEventHandler(cboNganhKD_KeyPress);
		((Control)pnMaHSDNHT).Controls.Add((Control)(object)cbMaHSDNHT_DADT);
		((Control)pnMaHSDNHT).Controls.Add((Control)(object)lblChuDADT);
		((Control)pnMaHSDNHT).Controls.Add((Control)(object)txtMaHSDNHT_DADT);
		((Control)pnMaHSDNHT).Controls.Add((Control)(object)label1);
		((Control)pnMaHSDNHT).Location = new Point(4, 432);
		((Control)pnMaHSDNHT).Name = "pnMaHSDNHT";
		((Control)pnMaHSDNHT).Size = new Size(372, 59);
		((Control)pnMaHSDNHT).TabIndex = 159;
		((Control)pnMaHSDNHT).Visible = false;
		((ListControl)cbMaHSDNHT_DADT).FormattingEnabled = true;
		((Control)cbMaHSDNHT_DADT).Location = new Point(120, 3);
		((Control)cbMaHSDNHT_DADT).Name = "cbMaHSDNHT_DADT";
		((Control)cbMaHSDNHT_DADT).Size = new Size(223, 21);
		((Control)cbMaHSDNHT_DADT).TabIndex = 54;
		cbMaHSDNHT_DADT.SelectedIndexChanged += cbMaHSDNHT_DADT_SelectedIndexChanged;
		((Control)cbMaHSDNHT_DADT).TextChanged += cbMaHSDNHT_DADT_TextChanged;
		((Control)cbMaHSDNHT_DADT).KeyDown += new KeyEventHandler(cbMaHSDNHT_DADT_KeyDown);
		((Control)cbMaHSDNHT_DADT).KeyPress += new KeyPressEventHandler(cbMaHSDNHT_DADT_KeyPress);
		((Control)lblChuDADT).AutoSize = true;
		((Control)lblChuDADT).Location = new Point(5, 34);
		((Control)lblChuDADT).Name = "lblChuDADT";
		((Control)lblChuDADT).Size = new Size(90, 13);
		((Control)lblChuDADT).TabIndex = 53;
		((Control)lblChuDADT).Text = "Tên dự án đầu tư";
		((Control)lblChuDADT).Click += label3_Click;
		((Control)txtMaHSDNHT_DADT).Enabled = false;
		((Control)txtMaHSDNHT_DADT).Location = new Point(120, 5);
		((TextBoxBase)txtMaHSDNHT_DADT).MaxLength = 12;
		((Control)txtMaHSDNHT_DADT).Name = "txtMaHSDNHT_DADT";
		((Control)txtMaHSDNHT_DADT).Size = new Size(81, 20);
		((Control)txtMaHSDNHT_DADT).TabIndex = 52;
		((Control)txtMaHSDNHT_DADT).Visible = false;
		((Control)txtMaHSDNHT_DADT).TextChanged += txtMaHSDNHT_DADT_TextChanged;
		((Control)txtMaHSDNHT_DADT).KeyPress += new KeyPressEventHandler(txtMaHSDNHT_DADT_KeyPress);
		((Control)txtMaHSDNHT_DADT).Leave += txtMaHSDNHT_DADT_Leave;
		((Control)label1).AutoSize = true;
		((Control)label1).Location = new Point(4, 8);
		((Control)label1).Name = "label1";
		((Control)label1).Size = new Size(107, 13);
		((Control)label1).TabIndex = 51;
		((Control)label1).Text = "Mã hồ sơ khai DAĐT";
		((Control)pnThuHo).Controls.Add((Control)(object)chkThuHo);
		((Control)pnThuHo).Location = new Point(3, 503);
		((Control)pnThuHo).Margin = new Padding(3, 0, 3, 0);
		((Control)pnThuHo).Name = "pnThuHo";
		((Control)pnThuHo).Size = new Size(372, 38);
		((Control)pnThuHo).TabIndex = 58;
		((Control)pnThuHo).Visible = false;
		((Control)chkThuHo).AutoSize = true;
		((Control)chkThuHo).Location = new Point(5, 10);
		((Control)chkThuHo).Name = "chkThuHo";
		((Control)chkThuHo).Size = new Size(240, 17);
		((Control)chkThuHo).TabIndex = 57;
		((Control)chkThuHo).Text = "Hoạt động thu hộ CQNN có thẩm quyền giao";
		((ButtonBase)chkThuHo).UseVisualStyleBackColor = true;
		chkThuHo.CheckedChanged += chkThuHo_CheckedChanged;
		((Control)pnVT1_3TuNgayDenNgay).Controls.Add((Control)(object)pnDenThang);
		((Control)pnVT1_3TuNgayDenNgay).Controls.Add((Control)(object)pnLanXuatBan);
		((Control)pnVT1_3TuNgayDenNgay).Controls.Add((Control)(object)pnDenNgay);
		((Control)pnVT1_3TuNgayDenNgay).Controls.Add((Control)(object)pnTuThang);
		((Control)pnVT1_3TuNgayDenNgay).Controls.Add((Control)(object)pnTuNgay);
		((Control)pnVT1_3TuNgayDenNgay).Location = new Point(2, 208);
		((Control)pnVT1_3TuNgayDenNgay).Margin = new Padding(3, 0, 3, 0);
		((Control)pnVT1_3TuNgayDenNgay).Name = "pnVT1_3TuNgayDenNgay";
		((Control)pnVT1_3TuNgayDenNgay).Size = new Size(768, 30);
		((Control)pnVT1_3TuNgayDenNgay).TabIndex = 48;
		((Control)pnDenThang).Controls.Add((Control)(object)mskDenThang);
		((Control)pnDenThang).Controls.Add((Control)(object)lblDenThang);
		((Control)pnDenThang).Location = new Point(584, 1);
		((Control)pnDenThang).Name = "pnDenThang";
		((Control)pnDenThang).Size = new Size(119, 25);
		((Control)pnDenThang).TabIndex = 48;
		((Control)mskDenThang).Location = new Point(62, 3);
		mskDenThang.Mask = "00/0000";
		((Control)mskDenThang).Name = "mskDenThang";
		((Control)mskDenThang).Size = new Size(50, 20);
		((Control)mskDenThang).TabIndex = 57;
		mskDenThang.TextAlign = (HorizontalAlignment)2;
		((Control)mskDenThang).Leave += mskDenThang_Leave;
		((Control)lblDenThang).AutoSize = true;
		((Control)lblDenThang).Location = new Point(3, 6);
		((Control)lblDenThang).Name = "lblDenThang";
		((Control)lblDenThang).Size = new Size(57, 13);
		((Control)lblDenThang).TabIndex = 9;
		((Control)lblDenThang).Text = "Đến tháng";
		((Control)pnLanXuatBan).Controls.Add((Control)(object)txtLanXuatBan);
		((Control)pnLanXuatBan).Controls.Add((Control)(object)lblLanXuat);
		((Control)pnLanXuatBan).Location = new Point(2, 3);
		((Control)pnLanXuatBan).Name = "pnLanXuatBan";
		((Control)pnLanXuatBan).Size = new Size(175, 25);
		((Control)pnLanXuatBan).TabIndex = 47;
		((Control)txtLanXuatBan).Location = new Point(72, 3);
		((TextBoxBase)txtLanXuatBan).MaxLength = 255;
		((Control)txtLanXuatBan).Name = "txtLanXuatBan";
		((Control)txtLanXuatBan).Size = new Size(100, 20);
		((Control)txtLanXuatBan).TabIndex = 20;
		((Control)txtLanXuatBan).TextChanged += txtLanXuatBan_TextChanged;
		((Control)txtLanXuatBan).KeyPress += new KeyPressEventHandler(txtLanXuatBan_KeyPress);
		((Control)txtLanXuatBan).Leave += txtLanXuatBan_Leave;
		((Control)lblLanXuat).AutoSize = true;
		((Control)lblLanXuat).Location = new Point(3, 6);
		((Control)lblLanXuat).Name = "lblLanXuat";
		((Control)lblLanXuat).Size = new Size(69, 13);
		((Control)lblLanXuat).TabIndex = 19;
		((Control)lblLanXuat).Text = "Lần xuất bán";
		((Control)pnDenNgay).Controls.Add((Control)(object)mskDenNgay);
		((Control)pnDenNgay).Controls.Add((Control)(object)lblDenNgay);
		((Control)pnDenNgay).Location = new Point(321, 3);
		((Control)pnDenNgay).Name = "pnDenNgay";
		((Control)pnDenNgay).Size = new Size(134, 25);
		((Control)pnDenNgay).TabIndex = 46;
		((Control)mskDenNgay).Location = new Point(58, 3);
		mskDenNgay.Mask = "00/00/0000";
		((Control)mskDenNgay).Name = "mskDenNgay";
		((Control)mskDenNgay).Size = new Size(70, 20);
		((Control)mskDenNgay).TabIndex = 55;
		mskDenNgay.TextAlign = (HorizontalAlignment)2;
		mskDenNgay.ValidatingType = typeof(DateTime);
		((Control)mskDenNgay).Leave += mskDenNgay_Leave;
		((Control)lblDenNgay).AutoSize = true;
		((Control)lblDenNgay).Location = new Point(3, 6);
		((Control)lblDenNgay).Name = "lblDenNgay";
		((Control)lblDenNgay).Size = new Size(53, 13);
		((Control)lblDenNgay).TabIndex = 9;
		((Control)lblDenNgay).Text = "Đến ngày";
		((Control)pnTuThang).Controls.Add((Control)(object)mskTuThang);
		((Control)pnTuThang).Controls.Add((Control)(object)lblTuThang);
		((Control)pnTuThang).Location = new Point(465, 2);
		((Control)pnTuThang).Name = "pnTuThang";
		((Control)pnTuThang).Size = new Size(113, 25);
		((Control)pnTuThang).TabIndex = 47;
		((Control)mskTuThang).Location = new Point(54, 3);
		mskTuThang.Mask = "00/0000";
		((Control)mskTuThang).Name = "mskTuThang";
		((Control)mskTuThang).Size = new Size(50, 20);
		((Control)mskTuThang).TabIndex = 56;
		mskTuThang.TextAlign = (HorizontalAlignment)2;
		((Control)mskTuThang).TextChanged += mskTuThang_TextChanged;
		((Control)mskTuThang).Leave += mskTuThang_Leave;
		((Control)lblTuThang).AutoSize = true;
		((Control)lblTuThang).Location = new Point(3, 6);
		((Control)lblTuThang).Name = "lblTuThang";
		((Control)lblTuThang).Size = new Size(50, 13);
		((Control)lblTuThang).TabIndex = 7;
		((Control)lblTuThang).Text = "Từ tháng";
		((Control)pnTuNgay).Controls.Add((Control)(object)mskTuNgay);
		((Control)pnTuNgay).Controls.Add((Control)(object)lblTuNgay);
		((Control)pnTuNgay).Location = new Point(180, 3);
		((Control)pnTuNgay).Name = "pnTuNgay";
		((Control)pnTuNgay).Size = new Size(129, 25);
		((Control)pnTuNgay).TabIndex = 45;
		((Control)mskTuNgay).Location = new Point(51, 3);
		mskTuNgay.Mask = "00/00/0000";
		((Control)mskTuNgay).Name = "mskTuNgay";
		((Control)mskTuNgay).Size = new Size(70, 20);
		((Control)mskTuNgay).TabIndex = 54;
		mskTuNgay.TextAlign = (HorizontalAlignment)2;
		mskTuNgay.ValidatingType = typeof(DateTime);
		((Control)mskTuNgay).Leave += mskTuNgay_Leave;
		((Control)lblTuNgay).AutoSize = true;
		((Control)lblTuNgay).Location = new Point(3, 6);
		((Control)lblTuNgay).Name = "lblTuNgay";
		((Control)lblTuNgay).Size = new Size(46, 13);
		((Control)lblTuNgay).TabIndex = 7;
		((Control)lblTuNgay).Text = "Từ ngày";
		((Control)pbHeader).BackgroundImageLayout = (ImageLayout)3;
		((Control)pbHeader).Dock = (DockStyle)1;
		pbHeader.Image = (Image)componentResourceManager.GetObject("pbHeader.Image");
		((Control)pbHeader).Location = new Point(0, 0);
		((Control)pbHeader).Margin = new Padding(10, 3, 10, 3);
		((Control)pbHeader).Name = "pbHeader";
		((Control)pbHeader).Size = new Size(1095, 20);
		pbHeader.TabIndex = 0;
		pbHeader.TabStop = false;
		((Control)chkToChucKhaiThay).AutoSize = true;
		((Control)chkToChucKhaiThay).Location = new Point(273, 769);
		((Control)chkToChucKhaiThay).Name = "chkToChucKhaiThay";
		((Control)chkToChucKhaiThay).Size = new Size(112, 17);
		((Control)chkToChucKhaiThay).TabIndex = 55;
		((Control)chkToChucKhaiThay).Text = "Tổ chức khai thay";
		((ButtonBase)chkToChucKhaiThay).UseVisualStyleBackColor = true;
		((Control)chkToChucKhaiThay).Visible = false;
		((Control)lblNgayPsThuNhap).AutoSize = true;
		((Control)lblNgayPsThuNhap).Location = new Point(828, 400);
		((Control)lblNgayPsThuNhap).Name = "lblNgayPsThuNhap";
		((Control)lblNgayPsThuNhap).Size = new Size(123, 13);
		((Control)lblNgayPsThuNhap).TabIndex = 54;
		((Control)lblNgayPsThuNhap).Text = "Ngày phát sinh thu nhập";
		((Control)lblNgayPsThuNhap).Visible = false;
		((Control)pnLoaiTien).Controls.Add((Control)(object)cboLoaiTien);
		((Control)pnLoaiTien).Controls.Add((Control)(object)lblLoaiTien);
		((Control)pnLoaiTien).Location = new Point(767, 498);
		((Control)pnLoaiTien).Name = "pnLoaiTien";
		((Control)pnLoaiTien).Size = new Size(236, 29);
		((Control)pnLoaiTien).TabIndex = 163;
		((Control)pnLoaiTien).Visible = false;
		((Control)lblLoaiTien).AutoSize = true;
		((Control)lblLoaiTien).Location = new Point(7, 7);
		((Control)lblLoaiTien).Name = "lblLoaiTien";
		((Control)lblLoaiTien).Size = new Size(80, 13);
		((Control)lblLoaiTien).TabIndex = 0;
		((Control)lblLoaiTien).Text = "Loại đơn vị tiền";
		((ListControl)cboLoaiTien).FormattingEnabled = true;
		((Control)cboLoaiTien).Location = new Point(94, 4);
		((Control)cboLoaiTien).Name = "cboLoaiTien";
		((Control)cboLoaiTien).Size = new Size(134, 21);
		((Control)cboLoaiTien).TabIndex = 1;
		((Control)cboLoaiTien).Text = "Đồng Việt Nam";
		cboLoaiTien.SelectedIndexChanged += cboLoaiTien_SelectedIndexChanged;
		((Control)cboLoaiTien).KeyPress += new KeyPressEventHandler(cboLoaiTien_KeyPress);
		((Control)panel1).Controls.Add((Control)(object)cboXa);
		((Control)panel1).Controls.Add((Control)(object)lblXa);
		((Control)panel1).Location = new Point(4, 33);
		((Control)panel1).Name = "panel1";
		((Control)panel1).Size = new Size(358, 25);
		((Control)panel1).TabIndex = 9;
		((ListControl)cboXa).FormattingEnabled = true;
		((Control)cboXa).Location = new Point(173, 3);
		((Control)cboXa).Name = "cboXa";
		((Control)cboXa).Size = new Size(173, 21);
		((Control)cboXa).TabIndex = 2;
		((Control)cboXa).Text = "Chọn xã";
		((Control)cboXa).Visible = false;
		cboXa.SelectedIndexChanged += cboXa_SelectedIndexChanged;
		((Control)lblXa).AutoSize = true;
		((Control)lblXa).Location = new Point(3, 6);
		((Control)lblXa).Name = "lblXa";
		((Control)lblXa).Size = new Size(108, 13);
		((Control)lblXa).TabIndex = 3;
		((Control)lblXa).Text = "Xã/Phường/Đặc khu";
		((Form)this).AcceptButton = (IButtonControl)(object)btnDongY;
		((ContainerControl)this).AutoScaleDimensions = new SizeF(6f, 13f);
		((ContainerControl)this).AutoScaleMode = (AutoScaleMode)1;
		((Control)this).BackgroundImageLayout = (ImageLayout)3;
		((Form)this).ClientSize = new Size(1095, 663);
		((Control)this).Controls.Add((Control)(object)chkToChucKhaiThay);
		((Control)this).Controls.Add((Control)(object)f);
		((Control)this).Controls.Add((Control)(object)chkCondensate);
		((Control)this).Controls.Add((Control)(object)chkKhiThien);
		((Control)this).Controls.Add((Control)(object)chkDauTho);
		((Control)this).Controls.Add((Control)(object)OptTKLanPS);
		((Control)this).Controls.Add((Control)(object)OptTKThang);
		((Control)this).Controls.Add((Control)(object)pbHeader);
		((Form)this).FormBorderStyle = (FormBorderStyle)0;
		((Form)this).MaximizeBox = false;
		((Form)this).MinimizeBox = false;
		((Control)this).Name = "frmHTKKPeriod";
		((Form)this).StartPosition = (FormStartPosition)1;
		((Control)this).Text = "Chọn kỳ tính thuế";
		((Form)this).Load += frmHTKKPeriod_Load;
		((Form)this).Shown += frmHTKKPeriod_Shown;
		((Control)pnVT3Button).ResumeLayout(false);
		((Control)pnVT2_PhuLucKeKhai).ResumeLayout(false);
		((Control)pnVT2_PhuLucKeKhai).PerformLayout();
		((ISupportInitialize)axfpsPhuLuc).EndInit();
		((ISupportInitialize)axfpsPhuLuc_Sheet1).EndInit();
		((Control)f).ResumeLayout(false);
		((Control)pnThongBao).ResumeLayout(false);
		((Control)pnThongBao).PerformLayout();
		((Control)pnMst).ResumeLayout(false);
		((Control)pnMst).PerformLayout();
		((Control)pnVT1).ResumeLayout(false);
		((Control)pnRDLanDauBoXung1).ResumeLayout(false);
		((Control)pnRDLanDauBoXung1).PerformLayout();
		((Control)pnNgayLap1).ResumeLayout(false);
		((Control)pnNgayLap1).PerformLayout();
		((Control)pnLan1).ResumeLayout(false);
		((Control)pnLan1).PerformLayout();
		((Control)pnTruongHopQT).ResumeLayout(false);
		((Control)pnTruongHopQT).PerformLayout();
		((Control)pnLoaiKyHoan).ResumeLayout(false);
		((Control)pnLoaiKyHoan).PerformLayout();
		((Control)pnChuyenNhuong).ResumeLayout(false);
		((Control)pnChuyenNhuong).PerformLayout();
		((Control)pnKhaiTheoPLuatDanSuThue).ResumeLayout(false);
		((Control)pnTinhCQT).ResumeLayout(false);
		((Control)pnCCT).ResumeLayout(false);
		((Control)pnCCT).PerformLayout();
		((Control)pnCQTCuc).ResumeLayout(false);
		((Control)pnCQTCuc).PerformLayout();
		((Control)pnTinh).ResumeLayout(false);
		((Control)pnTinh).PerformLayout();
		((Control)pnMSTCaNhanToChuc).ResumeLayout(false);
		((Control)pnMSTCaNhanToChuc).PerformLayout();
		((Control)pnPPhapNopThue).ResumeLayout(false);
		((Control)pnVT1_1CheckBox).ResumeLayout(false);
		((Control)pnVT1_1CheckBox).PerformLayout();
		((Control)pnVT1_4LanDauBoXung).ResumeLayout(false);
		((Control)pnQuyetToan_LienTuc).ResumeLayout(false);
		((Control)pnQuyetToan_LienTuc).PerformLayout();
		((Control)pnNgayLap).ResumeLayout(false);
		((Control)pnNgayLap).PerformLayout();
		((Control)pnRDLanDauBoXung).ResumeLayout(false);
		((Control)pnRDLanDauBoXung).PerformLayout();
		((Control)pnLan).ResumeLayout(false);
		((Control)pnLan).PerformLayout();
		((Control)pnVT1_2NgayThangQuyNam).ResumeLayout(false);
		((Control)pnKy).ResumeLayout(false);
		((Control)pnKy).PerformLayout();
		((Control)pnNam).ResumeLayout(false);
		((Control)pnNam).PerformLayout();
		((Control)pnNgay).ResumeLayout(false);
		((Control)pnNgay).PerformLayout();
		((Control)pnQuy).ResumeLayout(false);
		((Control)pnQuy).PerformLayout();
		((Control)pnThang).ResumeLayout(false);
		((Control)pnThang).PerformLayout();
		((Control)pnVT1_5DanhMucNghe).ResumeLayout(false);
		((Control)pnVT1_5DanhMucNghe).PerformLayout();
		((Control)pnMaHSDNHT).ResumeLayout(false);
		((Control)pnMaHSDNHT).PerformLayout();
		((Control)pnThuHo).ResumeLayout(false);
		((Control)pnThuHo).PerformLayout();
		((Control)pnVT1_3TuNgayDenNgay).ResumeLayout(false);
		((Control)pnDenThang).ResumeLayout(false);
		((Control)pnDenThang).PerformLayout();
		((Control)pnLanXuatBan).ResumeLayout(false);
		((Control)pnLanXuatBan).PerformLayout();
		((Control)pnDenNgay).ResumeLayout(false);
		((Control)pnDenNgay).PerformLayout();
		((Control)pnTuThang).ResumeLayout(false);
		((Control)pnTuThang).PerformLayout();
		((Control)pnTuNgay).ResumeLayout(false);
		((Control)pnTuNgay).PerformLayout();
		((ISupportInitialize)pbHeader).EndInit();
		((Control)pnLoaiTien).ResumeLayout(false);
		((Control)pnLoaiTien).PerformLayout();
		((Control)panel1).ResumeLayout(false);
		((Control)panel1).PerformLayout();
		((Control)this).ResumeLayout(false);
		((Control)this).PerformLayout();
	}
}
