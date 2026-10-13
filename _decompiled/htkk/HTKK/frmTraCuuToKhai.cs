using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Windows.Forms.Layout;
using System.Xml;
using TAX.HTKK.Business;
using TAX.HTKK.Utilities.Control;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;
using TAX.HTKK.Utilities.Implement;

namespace HTKK;

public class frmTraCuuToKhai : Form
{
	private static Dictionary<string, string> dicSacThue = new Dictionary<string, string>();

	private static Dictionary<string, HTKKTraCuuInfo> dictData = new Dictionary<string, HTKKTraCuuInfo>();

	private static Dictionary<string, HTKKTraCuuInfo> dictDataFull = new Dictionary<string, HTKKTraCuuInfo>();

	private static XmlNode _XmlNode;

	private static DataTable ListDataTableFile = new DataTable();

	private List<HTKKTraCuuInfo> listDataFileResult;

	private bool isEventContent = false;

	private bool isFisrtLoad = true;

	private string strDropTK = "";

	private string loaiKyKeKhai = "";

	private IContainer components = null;

	private GroupBox groupBox1;

	private GroupBox groupBox2;

	private ComboBox DropLoaiToKhai;

	private GroupBox groupBox3;

	private Button btnDelete;

	private Button cmdExit;

	private Button btnOpen;

	private Button btnTraCuu;

	private DataGridView grv_listTable;

	private Panel panelLoadControls;

	private CheckBox chkAll;

	private Label label2;

	private ComboBox DropSacThue;

	private Label label1;

	private Panel panel1;

	private Label label3;

	private Panel panel2;

	private Label label4;

	private Label label5;

	private DataGridViewTextBoxColumn Column8;

	private DataGridViewCheckBoxColumn Column5;

	private DataGridViewTextBoxColumn Column7;

	private DataGridViewTextBoxColumn Column6;

	private DataGridViewTextBoxColumn VersionAPP;

	private DataGridViewTextBoxColumn VersionXML;

	private DataGridViewTextBoxColumn Column2;

	private DataGridViewTextBoxColumn LoaiTK;

	private DataGridViewTextBoxColumn Column1;

	private DataGridViewTextBoxColumn lanBS;

	private DataGridViewTextBoxColumn Column3;

	private DataGridViewTextBoxColumn Column4;

	private DataGridViewTextBoxColumn DataFile;

	public frmTraCuuToKhai(XmlNode xmlNode)
	{
		//IL_0142: Unknown result type (might be due to invalid IL or missing references)
		//IL_0148: Expected O, but got Unknown
		//IL_09cb: Unknown result type (might be due to invalid IL or missing references)
		//IL_09d1: Expected O, but got Unknown
		//IL_0184: Unknown result type (might be due to invalid IL or missing references)
		//IL_018b: Expected O, but got Unknown
		//IL_0a0d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0a14: Expected O, but got Unknown
		InitializeComponent();
		_XmlNode = xmlNode;
		XmlNodeList xnList = ((XmlNode)WinSettings.xmlDocMenu).SelectNodes("/Root/Menu");
		dicSacThue = HTKKFactory.GetInstanceHTKKTraCuu().GetDicSacThue(xnList);
		dictData = HTKKFactory.GetInstanceHTKKTraCuu().GetDictDataByMapFileXml("Map.xml");
		dictDataFull = dictData;
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
		string dataFileName = "03_TNDN~03_1A_TNDN~03_1B_TNDN~03_1C_TNDN~03_2A_TNDN~03_2B_TNDN~03_3A_TNDN~03_3B_TNDN~03_3C_TNDN~03_4_TNDN~03_5_TNDN~03_6_TNDN~03_7_TNDN~03_8_TNDN~03_9_TD_TNDN~03_02_TNDN~03_01_ND20_TNDN~03_02_ND20_TNDN~03_03_ND20_TNDN~03_04_ND20_TNDN~KHBS_03_TNDN";
		string versionXmlOld = "231";
		string versionXml = "239";
		WinFnProcess.ConverDataVersion(dataFileName, versionXmlOld, versionXml);
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
		dataFileName = "02_TNDN~02_1_MGT_TNDN~KHBS_02_TNDN";
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
		WinFnProcess.FixPhuLucVTNN("01_1_VTNN_TT80", "251");
		WinFnProcess.FixPhuLucVTNN("01_2_VTNN_TT80", "251");
		WinFnProcess.FixPhuLucVTNN("01_3_VTNN_TT80", "251");
		WinFnProcess.Fix05QTTSuaDuLieu("05_1_QTT_TNCN_TT80", "279", "###Thẻ CCCD###", "###Thẻ CCCD/Số định danh cá nhân###");
		WinFnProcess.Fix05QTTSuaDuLieu("05_2_QTT_TNCN_TT80", "279", "###Thẻ CCCD###", "###Thẻ CCCD/Số định danh cá nhân###");
		WinFnProcess.Fix05QTTSuaDuLieu("05_3_QTT_TNCN_TT80", "279", "###Thẻ CCCD###", "###Thẻ CCCD/Số định danh cá nhân###");
		WinFnProcess.ReplaceDataFileDSNopThayTNCN("Ds_tien_nop_TNCN");
		WinFnProcess.FixStaticToDynamic03TNDN_Tk(555, "292");
		WinFnProcess.ReplaceDataFileStaticCommon("03_TNDN_TT80", "292", 557, "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Header/ct04_ten", "Value", "");
		WinFnProcess.ReplaceDataFileStaticCommon("03_TNDN_TT80", "292", 557, "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Header/ct04_ma", "Value", "");
		WinFnProcess.ReplaceDataFileDinamicCommon("03_3A_TNDN_TT80", "292", 557, "PLuc_03_3A_TNDN/NoiDungPL03_3A", 7, "");
		WinFnProcess.ReplaceDataFileDinamicCommon("03_3A_TNDN_TT80", "292", 557, "PLuc_03_3A_TNDN/NoiDungPL03_3A", 28, "");
		WinFnProcess.ReplaceDataFileDinamicCommon("03_3B_TNDN_TT80", "292", 557, "PLuc_03_3B_TNDN/NoiDungPL03_3B", 8, "");
		WinFnProcess.ReplaceDataFileDinamicCommon("03_3B_TNDN_TT80", "292", 557, "PLuc_03_3B_TNDN/NoiDungPL03_3B", 31, "");
		convertDL_MGT_ND92();
		WinFnProcess.FixDataFile_01_6_CSSX_GTGT_TT80();
		WinFnProcess.ConvertDataCommon();
		WinFnProcess.FixDataFile_03_8_TNDN_TT80();
		WinFnProcess.FixDataFile_01_DNHUY();
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
		if (dicSacThue.Count > 0)
		{
			LoadDropSacThue(dicSacThue);
		}
		if (dictData.Count > 0)
		{
			LoadDropLoaiToKhai(dictData);
			isFisrtLoad = false;
			ListDataTableFile = HTKKFile.GetDataTableFilesInDirectory(HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst);
		}
	}

	private void frmTraCuuToKhai_Load(object sender, EventArgs e)
	{
		((Control)DropSacThue).Select();
		Control val = ((Control)panelLoadControls).Controls[0];
		Control val2 = FormProcess.FindControl(((Control)panelLoadControls).Controls[0], "txtBegin");
		Control val3 = FormProcess.FindControl(((Control)panelLoadControls).Controls[0], "txtEnd");
		val2.TabIndex = 4;
		val3.TabIndex = 5;
		Control val4 = FormProcess.FindControl(((Control)panelLoadControls).Controls[0], "txtLanXuatBan");
		string beginValue = ((val2 == null) ? string.Empty : val2.Text);
		string endValue = ((val3 == null) ? string.Empty : val3.Text);
		string lan = ((val4 == null) ? string.Empty : val4.Text);
		string loaiToKhai = ((ListControl)DropLoaiToKhai).SelectedValue.ToString();
		LoadGridView(loaiToKhai, beginValue, endValue, lan);
	}

	private void LoadDropSacThue(Dictionary<string, string> data)
	{
		//IL_00b6: Unknown result type (might be due to invalid IL or missing references)
		//IL_00c0: Expected O, but got Unknown
		try
		{
			DropSacThue.Items.Clear();
			Dictionary<string, string> dictionary = new Dictionary<string, string>();
			dictionary.Add("-1", "Tất cả");
			DropSacThue.Items.Add((object)new
			{
				Text = "Tất cả",
				Value = "-1"
			});
			foreach (KeyValuePair<string, string> datum in data)
			{
				dictionary.Add(datum.Key, datum.Value.ToString().Split(new char[1] { '$' })[0]);
			}
			DropSacThue.DataSource = (object)new BindingSource((object)dictionary, (string)null);
			((ListControl)DropSacThue).DisplayMember = "Value";
			((ListControl)DropSacThue).ValueMember = "Key";
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void LoadDropLoaiToKhai(Dictionary<string, HTKKTraCuuInfo> data)
	{
		//IL_00ce: Unknown result type (might be due to invalid IL or missing references)
		//IL_00d8: Expected O, but got Unknown
		try
		{
			DropLoaiToKhai.DataSource = null;
			DropLoaiToKhai.Items.Clear();
			Dictionary<string, string> dictionary = new Dictionary<string, string>();
			if (data.Count > 0)
			{
				if (isFisrtLoad)
				{
					dictionary.Add("-1", "Tất cả");
					DropLoaiToKhai.Items.Add((object)new
					{
						Text = "Tất cả",
						Value = "-1"
					});
				}
				foreach (KeyValuePair<string, HTKKTraCuuInfo> datum in data)
				{
					dictionary.Add(datum.Key, datum.Value.Caption);
				}
				DropLoaiToKhai.DataSource = (object)new BindingSource((object)dictionary, (string)null);
				((ListControl)DropLoaiToKhai).DisplayMember = "Value";
				((ListControl)DropLoaiToKhai).ValueMember = "Key";
			}
			if (!HTKKFormater.isNullString(strDropTK))
			{
				DropLoaiToKhai.Items.Add((object)new
				{
					Text = strDropTK,
					Value = "-2"
				});
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void DropLoaiToKhai_KeyPress(object sender, KeyPressEventArgs e)
	{
		e.Handled = true;
	}

	private void DropLoaiToKhai_SelectedIndexChanged(object sender, EventArgs e)
	{
		//IL_0125: Unknown result type (might be due to invalid IL or missing references)
		//IL_012c: Expected O, but got Unknown
		try
		{
			string text = "";
			string text2 = "";
			if (((ListControl)DropLoaiToKhai).SelectedValue == null)
			{
				return;
			}
			string key = ((ListControl)DropLoaiToKhai).SelectedValue.ToString();
			bool flag = true;
			if (dictData.ContainsKey(key))
			{
				HTKKTraCuuInfo hTKKTraCuuInfo = dictData[key];
				text2 = hTKKTraCuuInfo.ID;
				text = hTKKTraCuuInfo.LoaiTK;
				if (hTKKTraCuuInfo.Day == "0" && hTKKTraCuuInfo.Month == "0" && hTKKTraCuuInfo.Year == "0")
				{
					flag = false;
				}
				FormProcess.LoadUserControl((Control)(object)panelLoadControls, WinSettings.DictMapCotrols[hTKKTraCuuInfo.LoaiTK], null);
				string[] array = Regex.Split("40;402;08;082;42;44;", ";");
				for (int i = 0; i < array.Length; i++)
				{
					if (array.Equals(hTKKTraCuuInfo.ID))
					{
						Assembly assembly = Assembly.Load("HTKK");
						Type type = assembly.GetType("HTKK.UcLanXuatBan", throwOnError: true, ignoreCase: true);
						UserControl val = (UserControl)Activator.CreateInstance(type, null);
						((Control)val).Dock = (DockStyle)2;
						((Control)panelLoadControls).Controls.Add((Control)(object)val);
						break;
					}
				}
			}
			else
			{
				FormProcess.LoadUserControl((Control)(object)panelLoadControls, WinSettings.DictMapCotrols["Y"], null);
			}
			if (!HTKKFormater.isNullString(text))
			{
				loadKyTraCuu(text);
			}
			if (!flag)
			{
				SetValueControlByName("1990", "txtBegin", ((Control)panelLoadControls).Controls[0]);
				SetValueControlByName("2220", "txtEnd", ((Control)panelLoadControls).Controls[0]);
				SetDisplayKy(((Control)panelLoadControls).Controls[0], display: false);
			}
			else
			{
				SetDisplayKy(((Control)panelLoadControls).Controls[0], display: true);
			}
			if (text2 == "P7")
			{
				SetDisplayKy(((Control)panelLoadControls).Controls[0], display: false);
			}
			loaiKyKeKhai = text;
			Control val2 = ((Control)panelLoadControls).Controls[0];
			Control val3 = FormProcess.FindControl(((Control)panelLoadControls).Controls[0], "txtBegin");
			Control val4 = FormProcess.FindControl(((Control)panelLoadControls).Controls[0], "txtEnd");
			Control val5 = FormProcess.FindControl(((Control)panelLoadControls).Controls[0], "txtLanXuatBan");
			string beginValue = ((val3 == null) ? string.Empty : val3.Text);
			string endValue = ((val4 == null) ? string.Empty : val4.Text);
			string lan = ((val5 == null) ? string.Empty : val5.Text);
			string loaiToKhai = ((ListControl)DropLoaiToKhai).SelectedValue.ToString();
			LoadGridView(loaiToKhai, beginValue, endValue, lan);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void LoadGridView(string loaiToKhai, string beginValue, string endValue, string lan)
	{
		//IL_038a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0391: Expected O, but got Unknown
		try
		{
			int num = 0;
			Dictionary<string, HTKKTraCuuInfo> dictionary = new Dictionary<string, HTKKTraCuuInfo>();
			if (dictData.ContainsKey(loaiToKhai))
			{
				dictionary.Add(loaiToKhai, dictData[loaiToKhai]);
			}
			else
			{
				dictionary = dictData;
			}
			ListDataTableFile = HTKKFile.GetDataTableFilesInDirectory(HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst);
			List<HTKKTraCuuInfo> listDataSelect = HTKKFactory.GetInstanceHTKKTraCuu().GetListDataSelect(ListDataTableFile, dictionary);
			orderByListDataFile(listDataSelect);
			listDataFileResult = new List<HTKKTraCuuInfo>();
			for (int i = 0; i < listDataSelect.Count; i++)
			{
				HTKKTraCuuInfo hTKKTraCuuInfo = listDataSelect[i];
				bool flag = true;
				for (int j = 0; j < hTKKTraCuuInfo.ListDateNumber.Count; j++)
				{
					string value = beginValue;
					string value2 = endValue;
					if (hTKKTraCuuInfo.KieuKy == "Q" && beginValue.IndexOf("/") > -1 && endValue.IndexOf("/") > -1)
					{
						string[] array = beginValue.Split(new char[1] { '/' });
						value = GeneralFunctions.GetNgayDauQuy(ConvertHelpers.ConvertToInt(array[0]), ConvertHelpers.ConvertToInt(array[1]), 1, 1).ToString("dd/MM/yyyy");
						array = endValue.Split(new char[1] { '/' });
						value2 = GeneralFunctions.GetNgayCuoiQuy(ConvertHelpers.ConvertToInt(array[0]), ConvertHelpers.ConvertToInt(array[1]), 1, 1).ToString("dd/MM/yyyy");
					}
					else if (hTKKTraCuuInfo.KieuKy == "K" && beginValue.IndexOf("/") > -1 && endValue.IndexOf("/") > -1)
					{
						string[] array = beginValue.Split(new char[1] { '/' });
						value = GeneralFunctions.GetNgayDauQuy(ConvertHelpers.ConvertToInt(array[0]) * 2 - 1, ConvertHelpers.ConvertToInt(array[1]), 1, 1).ToString("dd/MM/yyyy");
						array = endValue.Split(new char[1] { '/' });
						value2 = GeneralFunctions.GetNgayCuoiQuy(ConvertHelpers.ConvertToInt(array[0]) * 2, ConvertHelpers.ConvertToInt(array[1]), 1, 1).ToString("dd/MM/yyyy");
					}
					if (hTKKTraCuuInfo.ListDateNumber[j] < HTKKFormater.FormatStringToDateNumberBegin(value) || hTKKTraCuuInfo.ListDateNumber[j] > HTKKFormater.FormatStringToDateNumberEnd(value2))
					{
						flag = false;
						break;
					}
					if (!HTKKFormater.isNullString(lan) && lan != hTKKTraCuuInfo.LanXuatBan)
					{
						flag = false;
						break;
					}
				}
				if (flag)
				{
					num++;
					hTKKTraCuuInfo.STT = num.ToString();
					hTKKTraCuuInfo.CheckValue = "false".ToString();
					listDataFileResult.Add(hTKKTraCuuInfo);
				}
			}
			DataTable dataTable = HTKKFactory.GetInstanceHTKKTraCuu().GetDataTable(listDataFileResult);
			DataRow[] array2 = dataTable.Select();
			foreach (DataRow dataRow in array2)
			{
				try
				{
					if (string.IsNullOrEmpty(dataRow["ThueKhauTru"].ToString()) && string.IsNullOrEmpty(dataRow["ThuePhaiNop"].ToString()))
					{
						continue;
					}
					XmlDocument val = new XmlDocument();
					string text = HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst + "\\" + dataRow["FileName"].ToString();
					if (!File.Exists(text))
					{
						continue;
					}
					val.LoadXml(HTKKFile.LoadFile(text));
					if (!string.IsNullOrEmpty(dataRow["ThueKhauTru"].ToString()))
					{
						XmlNodeList val2 = ((XmlNode)val).SelectNodes("/Sections/Section/Cells/Cell[@Path='" + dataRow["ThueKhauTru"].ToString() + "']");
						if (val2.Count != 0 && val2[0].Attributes["Value"] != null)
						{
							CultureInfo provider = CultureInfo.CreateSpecificCulture("el-GR");
							string value3 = ((XmlNode)val2[0].Attributes["Value"]).Value;
							value3 = string.Format(provider, "{0:0,0}", new object[1] { ConvertHelpers.ConvertToDouble1(value3) });
							if (value3.Length <= 2 && value3.IndexOf("0") == 0)
							{
								value3 = value3.Substring(1);
							}
							if (value3.Length <= 3 && (value3.IndexOf("-0") == 0 || value3.IndexOf("(0") == 0))
							{
								value3 = "(" + value3.Substring(2) + ")";
							}
							if (value3.IndexOf("-") == 0)
							{
								value3 = "(" + value3.Substring(1) + ")";
							}
							dataRow["ThueKhauTru"] = value3;
						}
					}
					if (string.IsNullOrEmpty(dataRow["ThuePhaiNop"].ToString()))
					{
						continue;
					}
					XmlNodeList val3 = ((XmlNode)val).SelectNodes("/Sections/Section/Cells/Cell[@Path='" + dataRow["ThuePhaiNop"].ToString() + "']");
					if (val3.Count != 0 && val3[0].Attributes["Value"] != null)
					{
						CultureInfo provider = CultureInfo.CreateSpecificCulture("el-GR");
						string value3 = ((XmlNode)val3[0].Attributes["Value"]).Value;
						value3 = string.Format(provider, "{0:0,0}", new object[1] { ConvertHelpers.ConvertToDouble1(value3) });
						if (value3.Length <= 2 && value3.IndexOf("0") == 0)
						{
							value3 = value3.Substring(1);
						}
						if (value3.Length <= 3 && (value3.IndexOf("-0") == 0 || value3.IndexOf("(0") == 0))
						{
							value3 = "(" + value3.Substring(2) + ")";
						}
						if (value3.IndexOf("-") == 0)
						{
							value3 = "(" + value3.Substring(1) + ")";
						}
						dataRow["ThuePhaiNop"] = value3;
					}
				}
				catch (Exception ex)
				{
					HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
				}
			}
			dataTable.AcceptChanges();
			grv_listTable.AutoGenerateColumns = false;
			grv_listTable.DataSource = dataTable;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void btnTraCuu_Click(object sender, EventArgs e)
	{
		//IL_023b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0142: Unknown result type (might be due to invalid IL or missing references)
		//IL_0290: Unknown result type (might be due to invalid IL or missing references)
		//IL_01e0: Unknown result type (might be due to invalid IL or missing references)
		//IL_0195: Unknown result type (might be due to invalid IL or missing references)
		//IL_0699: Unknown result type (might be due to invalid IL or missing references)
		//IL_0347: Unknown result type (might be due to invalid IL or missing references)
		//IL_0394: Unknown result type (might be due to invalid IL or missing references)
		//IL_042b: Unknown result type (might be due to invalid IL or missing references)
		//IL_03d9: Unknown result type (might be due to invalid IL or missing references)
		//IL_048b: Unknown result type (might be due to invalid IL or missing references)
		//IL_05c7: Unknown result type (might be due to invalid IL or missing references)
		//IL_04d9: Unknown result type (might be due to invalid IL or missing references)
		//IL_060a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0539: Unknown result type (might be due to invalid IL or missing references)
		//IL_064f: Unknown result type (might be due to invalid IL or missing references)
		//IL_057f: Unknown result type (might be due to invalid IL or missing references)
		try
		{
			Control val = ((Control)panelLoadControls).Controls[0];
			Control val2 = FormProcess.FindControl(((Control)panelLoadControls).Controls[0], "txtBegin");
			Control val3 = FormProcess.FindControl(((Control)panelLoadControls).Controls[0], "txtEnd");
			Control val4 = FormProcess.FindControl(((Control)panelLoadControls).Controls[0], "txtLanXuatBan");
			if (((ArrangedElementCollection)((Control)panelLoadControls).Controls).Count >= 2)
			{
				val4 = FormProcess.FindControl(((Control)panelLoadControls).Controls[1], "txtLanXuatBan");
			}
			string text = ((val2 == null) ? string.Empty : val2.Text);
			string text2 = ((val3 == null) ? string.Empty : val3.Text);
			string lan = ((val4 == null) ? string.Empty : val4.Text);
			string text3 = ((ListControl)DropLoaiToKhai).SelectedValue.ToString();
			if (text3.Equals("-1"))
			{
				if (val2 != null && string.IsNullOrEmpty(val2.Text.Trim()))
				{
					frmMessage frmMessage = new frmMessage("0017", "3", "1", "Thông báo");
					((Form)frmMessage).ShowDialog();
					val2.Focus();
					return;
				}
				if (val3 != null && string.IsNullOrEmpty(val3.Text.Trim()))
				{
					frmMessage frmMessage = new frmMessage("0017", "3", "1", "Thông báo");
					((Form)frmMessage).ShowDialog();
					val3.Focus();
					return;
				}
				if (ConvertHelpers.ConvertToInt(text) > ConvertHelpers.ConvertToInt(text2))
				{
					frmMessage frmMessage = new frmMessage("0093", "3", "1", "Thông báo");
					((Form)frmMessage).ShowDialog();
					val2.Focus();
					return;
				}
			}
			else
			{
				if (val2 != null && string.IsNullOrEmpty(val2.Text.Trim()))
				{
					frmMessage frmMessage = new frmMessage("0017", "3", "1", "Thông báo");
					((Form)frmMessage).ShowDialog();
					val2.Focus();
					return;
				}
				if (val3 != null && string.IsNullOrEmpty(val3.Text.Trim()))
				{
					frmMessage frmMessage = new frmMessage("0017", "3", "1", "Thông báo");
					((Form)frmMessage).ShowDialog();
					val3.Focus();
					return;
				}
				bool success = true;
				switch (loaiKyKeKhai)
				{
				case "M":
				{
					DateTime t = HTKKFormater.ConvertToDateTime("01/" + text, out success);
					if (!success)
					{
						frmMessage frmMessage = new frmMessage("0090", "3", "1", "Thông báo");
						((Form)frmMessage).ShowDialog();
						val2.Focus();
						return;
					}
					DateTime t2 = HTKKFormater.ConvertToDateTime("01/" + text2, out success);
					if (!success)
					{
						frmMessage frmMessage = new frmMessage("0090", "3", "1", "Thông báo");
						((Form)frmMessage).ShowDialog();
						val3.Focus();
						return;
					}
					if (DateTime.Compare(t, t2) > 0)
					{
						frmMessage frmMessage = new frmMessage("0097", "3", "1", "Thông báo");
						((Form)frmMessage).ShowDialog();
						val2.Focus();
						return;
					}
					break;
				}
				case "Q":
				{
					DateTime t = HTKKFormater.ConvertToDateTime("01/" + text, out success);
					if (!success)
					{
						frmMessage frmMessage = new frmMessage("0456", "3", "1", "Thông báo");
						((Form)frmMessage).ShowDialog();
						val2.Focus();
						return;
					}
					string[] array = text.Split(new char[1] { '/' });
					if (ConvertHelpers.ConvertToInt(array[0]) > 4)
					{
						frmMessage frmMessage = new frmMessage("0091", "3", "1", "Thông báo");
						((Form)frmMessage).ShowDialog();
						val2.Focus();
						return;
					}
					DateTime t2 = HTKKFormater.ConvertToDateTime("01/" + text2, out success);
					if (!success)
					{
						frmMessage frmMessage = new frmMessage("0456", "3", "1", "Thông báo");
						((Form)frmMessage).ShowDialog();
						val3.Focus();
						return;
					}
					array = text2.Split(new char[1] { '/' });
					if (ConvertHelpers.ConvertToInt(array[0]) > 4)
					{
						frmMessage frmMessage = new frmMessage("0091", "3", "1", "Thông báo");
						((Form)frmMessage).ShowDialog();
						val3.Focus();
						return;
					}
					if (DateTime.Compare(t, t2) > 0)
					{
						frmMessage frmMessage = new frmMessage("0098", "3", "1", "Thông báo");
						((Form)frmMessage).ShowDialog();
						val2.Focus();
						return;
					}
					break;
				}
				case "D":
				case "D_D":
				case "D_Y":
				{
					DateTime t = HTKKFormater.ConvertToDateTime(text, out success);
					if (!success)
					{
						frmMessage frmMessage = new frmMessage("0071", "3", "1", "Thông báo");
						((Form)frmMessage).ShowDialog();
						val2.Focus();
						return;
					}
					DateTime t2 = HTKKFormater.ConvertToDateTime(text2, out success);
					if (!success)
					{
						frmMessage frmMessage = new frmMessage("0071", "3", "1", "Thông báo");
						((Form)frmMessage).ShowDialog();
						val3.Focus();
						return;
					}
					if (DateTime.Compare(t, t2) > 0)
					{
						frmMessage frmMessage = new frmMessage("0309", "3", "1", "Thông báo");
						((Form)frmMessage).ShowDialog();
						val2.Focus();
						return;
					}
					break;
				}
				default:
					if (ConvertHelpers.ConvertToInt(text) > ConvertHelpers.ConvertToInt(text2))
					{
						frmMessage frmMessage = new frmMessage("0093", "3", "1", "Thông báo");
						((Form)frmMessage).ShowDialog();
						val2.Focus();
						return;
					}
					break;
				}
			}
			LoadGridView(text3, text, text2, lan);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void btnOpen_Click(object sender, EventArgs e)
	{
		try
		{
			if (((BaseCollection)grv_listTable.SelectedCells).Count <= 0)
			{
				return;
			}
			string fileName = grv_listTable.SelectedCells[3].Value.ToString();
			HTKKTraCuuInfo hTKKTraCuuInfo = listDataFileResult[int.Parse(grv_listTable.SelectedCells[0].Value.ToString()) - 1];
			WinSettings.TTPeriodInfo = GetPeriodInfo(hTKKTraCuuInfo, fileName);
			XmlNodeList val = ((XmlNode)WinSettings.xmlDocMenu).SelectNodes("/Root/Menu[@ID='" + WinSettings.TTPeriodInfo.menuID + "']");
			WinSettings.xmlMenuNode = val[0];
			WinSettings.MenuID = WinSettings.TTPeriodInfo.menuID;
			WinSettings.ToKhaiInfo.NopTheoTinhCQT = hTKKTraCuuInfo.NopTheoTinhCQT;
			WinSettings.KieuKy = ((WinSettings.xmlMenuNode.Attributes["KieuKy"] == null) ? "" : ((XmlNode)WinSettings.xmlMenuNode.Attributes["KieuKy"]).Value);
			if (!string.IsNullOrEmpty(WinSettings.ToKhaiInfo.NopTheoTinhCQT))
			{
				WinSettings.HTKKHeaderInfo.maCQTCapCuc = WinSettings.TTPeriodInfo.maCQTCapCuc;
				WinSettings.HTKKHeaderInfo.tenCQTCapCuc = WinSettings.TTPeriodInfo.tenCQTCapCuc;
				WinSettings.HTKKHeaderInfo.maCQTNoiNop = WinSettings.TTPeriodInfo.maCQTQuanLy;
				WinSettings.HTKKHeaderInfo.tenCQTNoiNop = WinSettings.TTPeriodInfo.tenCQTQuanLy;
				WinSettings.TTinChung.maCQTCapCuc = WinSettings.TTPeriodInfo.maCQTCapCuc;
				WinSettings.TTinChung.tenCQTCapCuc = WinSettings.TTPeriodInfo.tenCQTCapCuc;
				WinSettings.TTinChung.maCQTNoiNop = WinSettings.TTPeriodInfo.maCQTQuanLy;
				WinSettings.TTinChung.tenCQTNoiNop = WinSettings.TTPeriodInfo.tenCQTQuanLy;
			}
			if ("66;67;09;10;60;61;G2;".IndexOf(WinSettings.TTPeriodInfo.menuID) > -1)
			{
				DateTime dateEffectByAttributes = GeneralFunctions.getDateEffectByAttributes(WinSettings.xmlMenuNode, "StartDate");
				string text = ((WinSettings.xmlMenuNode.Attributes["VersionXML"] == null) ? "" : ((XmlNode)WinSettings.xmlMenuNode.Attributes["VersionXML"]).Value.Replace(".", ""));
				if (WinSettings.TTPeriodInfo.VersionXML == text)
				{
					WinSettings.TTPeriodInfo.Nam = dateEffectByAttributes.ToString("yyyy");
				}
				else
				{
					bool flag = GeneralFunctions.checkVersionEffectByVerXml(ref WinSettings.xmlMenuNode, WinDoWorkEvent.getMenuNode(WinSettings.TTPeriodInfo.menuID), WinSettings.TTPeriodInfo.VersionXML);
					dateEffectByAttributes = GeneralFunctions.getDateEffectByAttributes(WinSettings.xmlMenuNode, "StartDate");
					text = ((WinSettings.xmlMenuNode.Attributes["VersionXML"] == null) ? "" : ((XmlNode)WinSettings.xmlMenuNode.Attributes["VersionXML"]).Value.Replace(".", ""));
					if (WinSettings.TTPeriodInfo.VersionXML == text)
					{
						WinSettings.TTPeriodInfo.Nam = dateEffectByAttributes.ToString("yyyy");
					}
				}
			}
			else
			{
				bool flag = GeneralFunctions.checkVersionDateEffect(ref WinSettings.xmlMenuNode, WinSettings.TTPeriodInfo, WinDoWorkEvent.getMenuNode(WinSettings.TTPeriodInfo.menuID));
			}
			string dataFileNameExtend = WinFnProcess.GetDataFileNameExtend(WinSettings.TTPeriodInfo);
			WinSettings.DicSheet = WinFnProcess.GetPhuLucDaKeKhai(bPhuLucKeKhaiVisible: true, dataFileNameExtend);
			if (WinSettings.xmlMenuNode.Attributes["Class"] != null && WinSettings.xmlMenuNode.Attributes["VersionXML"] != null)
			{
				string text2 = ((XmlNode)WinSettings.xmlMenuNode.Attributes["Class"]).Value + HTKKConfig.APP_PREFIX + ((XmlNode)WinSettings.xmlMenuNode.Attributes["VersionXML"]).Value.Replace(".", "");
				if (text2.Length > 0)
				{
					WinSettings.InstanceClass = WinDoWorkEvent.GetInstance(text2);
				}
			}
			WinSettings.myHTKKSystems.frmChild = (Control)(object)WinSettings.myHTKKSystems.panelPeriod;
			WinFnProcess.OpenToKhai((Form)(object)this, null, isCoLayThemPhuLuc: false, WinSettings.TTPeriodInfo);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void btnDelete_Click(object sender, EventArgs e)
	{
		//IL_001d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0053: Unknown result type (might be due to invalid IL or missing references)
		//IL_0059: Expected O, but got Unknown
		//IL_0641: Unknown result type (might be due to invalid IL or missing references)
		//IL_0100: Unknown result type (might be due to invalid IL or missing references)
		//IL_0106: Expected O, but got Unknown
		//IL_023e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0245: Expected O, but got Unknown
		//IL_0278: Unknown result type (might be due to invalid IL or missing references)
		//IL_027f: Expected O, but got Unknown
		try
		{
			frmMessage frmMessage = new frmMessage("0094", "2", "2", "Thông báo");
			((Form)frmMessage).ShowDialog();
			if (!frmMessage.Yes)
			{
				return;
			}
			bool flag = false;
			foreach (DataGridViewRow item in (IEnumerable)grv_listTable.Rows)
			{
				DataGridViewRow val = item;
				if (val.Cells[1].Value != null && val.Cells[1].Value.ToString() == "True")
				{
					flag = true;
					break;
				}
			}
			if (flag)
			{
				Dictionary<string, string> dictionary = new Dictionary<string, string>();
				foreach (DataGridViewRow item2 in (IEnumerable)grv_listTable.Rows)
				{
					DataGridViewRow val = item2;
					if (val.Cells[1].Value == null || !(val.Cells[1].Value.ToString() == "True"))
					{
						continue;
					}
					string text = val.Cells[3].Value.ToString();
					string[] array = new string[2];
					List<string> list = new List<string>();
					string text2 = val.Cells["DataFile"].Value.ToString();
					if (text.IndexOf(text2) > -1)
					{
						array = Regex.Split(text, text2);
					}
					string text3 = val.Cells[2].Value.ToString().Trim();
					if (text3.Length == 3)
					{
						text3 = text3.Substring(0, text3.Length - 1);
					}
					if (text3.Length == 4)
					{
						text3 = text3.Substring(0, text3.Length - 2);
					}
					XmlNodeList val2 = ((XmlNode)WinSettings.xmlDocMenu).SelectNodes("/Root/Menu[@ID='" + text3 + "']");
					XmlNode val3 = val2[0];
					XmlDocument val4 = new XmlDocument();
					val4.LoadXml(val3.OuterXml);
					XmlNodeList elementsByTagName = val4.GetElementsByTagName("Sheet");
					foreach (XmlNode item3 in elementsByTagName)
					{
						XmlNode val5 = item3;
						if (array[0].StartsWith("bs") && ((XmlNode)val5.Attributes["DataFile"]).Value.IndexOf("KHBS") > -1)
						{
							string oldValue = "_" + ((val3.Attributes["VersionXML"] == null) ? "" : ((XmlNode)val3.Attributes["VersionXML"]).Value).Replace(".", "") + ".xml";
							string newValue = "_" + ((val3.Attributes["VersionKHBS"] == null) ? "" : ((XmlNode)val3.Attributes["VersionKHBS"]).Value).Replace(".", "") + ".xml";
							HTKKFile.AddOrUpdateDicStr(dictionary, array[0].ToString() + ((XmlNode)val5.Attributes["DataFile"]).Value + array[1].Replace(oldValue, newValue), val.Cells[2].Value.ToString());
						}
						else
						{
							string text4 = ((val5.Attributes["DataFile"] == null) ? "#####" : ((XmlNode)val5.Attributes["DataFile"]).Value);
							HTKKFile.AddOrUpdateDicStr(dictionary, array[0].ToString() + text4 + array[1], val.Cells[2].Value.ToString());
						}
					}
				}
				if (dictionary.Count <= 0)
				{
					return;
				}
				foreach (KeyValuePair<string, string> item4 in dictionary)
				{
					if (File.Exists(HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst + "\\" + item4.Key))
					{
						HTKKFile.DeleteFile(HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst + "\\" + item4.Key);
					}
				}
				Control val6 = ((Control)panelLoadControls).Controls[0];
				Control val7 = FormProcess.FindControl(((Control)panelLoadControls).Controls[0], "txtBegin");
				Control val8 = FormProcess.FindControl(((Control)panelLoadControls).Controls[0], "txtEnd");
				Control val9 = FormProcess.FindControl(((Control)panelLoadControls).Controls[0], "txtLanXuatBan");
				string beginValue = ((val7 == null) ? string.Empty : val7.Text);
				string endValue = ((val8 == null) ? string.Empty : val8.Text);
				string lan = ((val9 == null) ? string.Empty : val9.Text);
				string loaiToKhai = ((ListControl)DropLoaiToKhai).SelectedValue.ToString();
				LoadGridView(loaiToKhai, beginValue, endValue, lan);
			}
			else
			{
				MessageBox.Show("Bạn chưa chọn tờ khai.");
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void cmdExit_Click(object sender, EventArgs e)
	{
		try
		{
			WinSettings.myHTKKSystems.formName = "";
			((Form)this).Close();
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private bool CheckVersion(string fileName)
	{
		//IL_0004: Unknown result type (might be due to invalid IL or missing references)
		//IL_000a: Expected O, but got Unknown
		try
		{
			bool result = true;
			XmlDocument val = new XmlDocument();
			val.LoadXml(HTKKFile.LoadFile(HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst + "\\" + fileName));
			if (((XmlNode)((XmlNode)val).LastChild.Attributes["Version"]).Value != null && ((XmlNode)((XmlNode)val).LastChild.Attributes["Version"]).Value != HTKKConfig.VERSION && !Directory.Exists(HTKKConfig.HISTORY + "\\" + ((XmlNode)((XmlNode)val).LastChild.Attributes["Version"]).Value))
			{
				result = false;
			}
			return result;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return false;
		}
	}

	private HTKKPeriodInfo GetPeriodInfo(HTKKTraCuuInfo resultChecked, string fileName)
	{
		HTKKPeriodInfo hTKKPeriodInfo = new HTKKPeriodInfo();
		hTKKPeriodInfo.menuID = resultChecked.ID;
		if (resultChecked.ID.Length > 2)
		{
			hTKKPeriodInfo.menuID = resultChecked.ID.Substring(0, 2);
		}
		hTKKPeriodInfo.Nam = resultChecked.Year;
		hTKKPeriodInfo.TuNgay = resultChecked.TuNgay;
		hTKKPeriodInfo.DenNgay = resultChecked.DenNgay;
		hTKKPeriodInfo.TuThang = resultChecked.TuThang;
		hTKKPeriodInfo.DenThang = resultChecked.DenThang;
		hTKKPeriodInfo.Thang = resultChecked.Thang;
		hTKKPeriodInfo.Quy = resultChecked.Quy;
		hTKKPeriodInfo.Ngay = resultChecked.Ngay;
		hTKKPeriodInfo.NgayLapKHBS = resultChecked.NgayLapKHBS;
		hTKKPeriodInfo.DoanhNghiepHoatDongLienTuc = resultChecked.DoanhNghiepHoatDongLienTuc;
		hTKKPeriodInfo.TieuMucHachToanVal = resultChecked.TieuMucHachToan;
		hTKKPeriodInfo.NganhNgheKinhDoanh = resultChecked.TieuMucHachToan;
		if (resultChecked.ID == "99" || resultChecked.ID.Substring(0, 2) == "R5")
		{
			hTKKPeriodInfo.NganhNgheKinhDoanh = resultChecked.NganhNgheKD;
		}
		hTKKPeriodInfo.LanXuatBan = resultChecked.LanXuatBan;
		hTKKPeriodInfo.tenCQTQuanLy = resultChecked.tenCQTQuanLy;
		hTKKPeriodInfo.maCQTQuanLy = resultChecked.maCQTQuanLy;
		hTKKPeriodInfo.KieuKyHoanThue = resultChecked.KieuKyHoanThue;
		hTKKPeriodInfo.ThuHo = resultChecked.ThuHo;
		hTKKPeriodInfo.MaHSDuAnDT = resultChecked.MaHSDuAnDT;
		if (resultChecked.ID == "78" || resultChecked.ID == "53")
		{
			hTKKPeriodInfo.KieuKy = "Y";
		}
		else if (resultChecked.ID == "A1" || resultChecked.ID == "A3")
		{
			hTKKPeriodInfo.KieuKy = "Q";
		}
		else
		{
			hTKKPeriodInfo.KieuKy = resultChecked.KieuKy;
		}
		if (resultChecked.ID == "30" && hTKKPeriodInfo.KieuKy == "Y")
		{
			hTKKPeriodInfo.TuNgay = "01/01/" + hTKKPeriodInfo.Nam;
			hTKKPeriodInfo.DenNgay = "31/12/" + hTKKPeriodInfo.Nam;
		}
		if (resultChecked.ID == "301" && hTKKPeriodInfo.KieuKy == "D_D")
		{
			hTKKPeriodInfo.KieuKy = "Q";
		}
		hTKKPeriodInfo.kyKKhai = resultChecked.KyKeKhai;
		hTKKPeriodInfo.LoaiToKhai = resultChecked.LoaiTK;
		hTKKPeriodInfo.VersionXML = resultChecked.VersionXML;
		if ((resultChecked.ID == "631" || resultChecked.ID == "L11") && !HTKKFormater.isNullString(resultChecked.Ngay))
		{
			if (hTKKPeriodInfo.dicVar.ContainsKey("CoSoMoiThanhLap"))
			{
				hTKKPeriodInfo.dicVar.Remove("CoSoMoiThanhLap");
			}
			hTKKPeriodInfo.dicVar.Add("CoSoMoiThanhLap", "true");
			if (hTKKPeriodInfo.dicVar.ContainsKey("NgayThanhLap"))
			{
				hTKKPeriodInfo.dicVar.Remove("NgayThanhLap");
			}
			hTKKPeriodInfo.dicVar.Add("NgayThanhLap", resultChecked.Ngay);
		}
		if (hTKKPeriodInfo.menuID == "47" || hTKKPeriodInfo.menuID == "K7")
		{
			string value = "";
			resultChecked.dicVar.TryGetValue("A2_CTDC", out value);
			if (ConvertHelpers.ConvertToBool(value))
			{
				hTKKPeriodInfo.CTDC = true;
				hTKKPeriodInfo.LNDC = false;
			}
			else
			{
				hTKKPeriodInfo.CTDC = false;
				hTKKPeriodInfo.LNDC = true;
			}
			hTKKPeriodInfo.NgayChiaCTLN = resultChecked.KyKeKhai;
		}
		hTKKPeriodInfo.SoLan = resultChecked.Lan;
		hTKKPeriodInfo.DonViTinhMa = resultChecked.DonViTinhMa;
		hTKKPeriodInfo.DonViTinhTen = resultChecked.DonViTinhTen;
		hTKKPeriodInfo.MaTinh = resultChecked.maTinh;
		hTKKPeriodInfo.ToChucCNDNVonChuyenNhuong = resultChecked.ToChucCNDNVonChuyenNhuong;
		hTKKPeriodInfo.maCQTCapCuc = resultChecked.maCQTCapCuc;
		hTKKPeriodInfo.tenCQTCapCuc = resultChecked.tenCQTCapCuc;
		hTKKPeriodInfo.TenTinh = resultChecked.tenTinh;
		hTKKPeriodInfo.TenXa = resultChecked.tenXa;
		hTKKPeriodInfo.MaXa = resultChecked.maXa;
		return hTKKPeriodInfo;
	}

	private void chkAll_CheckedChanged(object sender, EventArgs e)
	{
		isEventContent = false;
	}

	private void grv_listTable_CellContentClick(object sender, DataGridViewCellEventArgs e)
	{
		//IL_00b7: Unknown result type (might be due to invalid IL or missing references)
		//IL_00bd: Expected O, but got Unknown
		bool flag = false;
		isEventContent = true;
		if (e.ColumnIndex != 1)
		{
			return;
		}
		DataGridViewRow val = grv_listTable.Rows[e.RowIndex];
		if (val.Cells[e.ColumnIndex].Value.ToString() == "false")
		{
			val.Cells[e.ColumnIndex].Value = "True";
		}
		else
		{
			val.Cells[e.ColumnIndex].Value = "false";
		}
		foreach (DataGridViewRow item in (IEnumerable)grv_listTable.Rows)
		{
			DataGridViewRow val2 = item;
			if (val2.Cells[e.ColumnIndex].Value.ToString() == "false")
			{
				flag = true;
				break;
			}
		}
		if (flag)
		{
			chkAll.Checked = false;
		}
		else
		{
			chkAll.Checked = true;
		}
	}

	private void chkAll_Click(object sender, EventArgs e)
	{
		//IL_00a1: Unknown result type (might be due to invalid IL or missing references)
		//IL_00a7: Expected O, but got Unknown
		//IL_0040: Unknown result type (might be due to invalid IL or missing references)
		//IL_0046: Expected O, but got Unknown
		if (isEventContent)
		{
			return;
		}
		if (chkAll.Checked)
		{
			foreach (DataGridViewRow item in (IEnumerable)grv_listTable.Rows)
			{
				DataGridViewRow val = item;
				val.Cells[1].Value = "True";
			}
		}
		else
		{
			foreach (DataGridViewRow item2 in (IEnumerable)grv_listTable.Rows)
			{
				DataGridViewRow val = item2;
				val.Cells[1].Value = "false";
			}
		}
		isEventContent = false;
	}

	private void grv_listTable_CellClick(object sender, DataGridViewCellEventArgs e)
	{
	}

	private void DropSacThue_SelectedIndexChanged(object sender, EventArgs e)
	{
		try
		{
			if (((ListControl)DropSacThue).SelectedValue == null)
			{
				return;
			}
			string key = ((ListControl)DropSacThue).SelectedValue.ToString();
			if (dicSacThue.ContainsKey(key))
			{
				string text = dicSacThue[key];
				if (text.IndexOf('$') == -1)
				{
					return;
				}
				text = text.Substring(text.IndexOf('$'));
				isFisrtLoad = false;
				dictData = new Dictionary<string, HTKKTraCuuInfo>();
				foreach (KeyValuePair<string, HTKKTraCuuInfo> item in dictDataFull)
				{
					string value = item.Key.Substring(0, 2);
					if (text.IndexOf(value) > 0)
					{
						dictData.Add(item.Key, item.Value);
					}
				}
				LoadDropLoaiToKhai(dictData);
				ListDataTableFile = HTKKFile.GetDataTableFilesInDirectory(HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst);
			}
			else
			{
				isFisrtLoad = true;
				dictData = dictDataFull;
				if (dictData.Count > 0)
				{
					LoadDropLoaiToKhai(dictData);
					ListDataTableFile = HTKKFile.GetDataTableFilesInDirectory(HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst);
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void DropSacThue_KeyPress(object sender, KeyPressEventArgs e)
	{
		e.Handled = true;
	}

	private void DropLoaiToKhai_Enter(object sender, EventArgs e)
	{
	}

	private void loadKyTraCuu(string loaiTK)
	{
		try
		{
			Control val = FormProcess.FindControl(((Control)panelLoadControls).Controls[0], "txtBegin");
			Control val2 = FormProcess.FindControl(((Control)panelLoadControls).Controls[0], "txtEnd");
			string value = "";
			string value2 = "";
			switch (loaiTK)
			{
			case "M":
				value = DateTime.Now.ToString("MM") + "/" + (DateTime.Now.Year - 1);
				value2 = DateTime.Now.ToString("MM") + "/" + DateTime.Now.Year;
				break;
			case "Q":
				if (DateTime.Now.Month <= 12 && DateTime.Now.Month > 9)
				{
					value = "04/" + (DateTime.Now.Year - 1);
					value2 = "04/" + DateTime.Now.Year;
				}
				if (DateTime.Now.Month <= 9 && DateTime.Now.Month > 6)
				{
					value = "03/" + (DateTime.Now.Year - 1);
					value2 = "03/" + DateTime.Now.Year;
				}
				if (DateTime.Now.Month <= 6 && DateTime.Now.Month > 3)
				{
					value = "02/" + (DateTime.Now.Year - 1);
					value2 = "02/" + DateTime.Now.Year;
				}
				if (DateTime.Now.Month <= 3 && DateTime.Now.Month > 0)
				{
					value = "01/" + (DateTime.Now.Year - 1);
					value2 = "01/" + DateTime.Now.Year;
				}
				break;
			case "K":
				if (DateTime.Now.Month <= 12 && DateTime.Now.Month > 6)
				{
					value = "02/" + (DateTime.Now.Year - 1);
					value2 = "02/" + DateTime.Now.Year;
				}
				if (DateTime.Now.Month <= 6 && DateTime.Now.Month > 0)
				{
					value = "01/" + (DateTime.Now.Year - 1);
					value2 = "01/" + DateTime.Now.Year;
				}
				break;
			case "D":
				value = DateTime.Now.ToString("dd") + "/" + DateTime.Now.ToString("MM") + "/" + (DateTime.Now.Year - 1);
				value2 = DateTime.Now.ToString("dd/MM/yyyy");
				break;
			case "D_D":
				value = DateTime.Now.ToString("dd") + "/" + DateTime.Now.ToString("MM") + "/" + (DateTime.Now.Year - 1);
				value2 = DateTime.Now.ToString("dd/MM/yyyy");
				break;
			case "D_Y":
				value = DateTime.Now.ToString("dd") + "/" + DateTime.Now.ToString("MM") + "/" + (DateTime.Now.Year - 1);
				value2 = DateTime.Now.ToString("dd/MM/yyyy");
				break;
			default:
				value = (DateTime.Now.Year - 1).ToString();
				value2 = DateTime.Now.Year.ToString();
				break;
			}
			SetValueControlByName(value, "txtBegin", ((Control)panelLoadControls).Controls[0]);
			SetValueControlByName(value2, "txtEnd", ((Control)panelLoadControls).Controls[0]);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void SetValueControlByName(string value, string nameCtl, Control parent)
	{
		//IL_0017: Unknown result type (might be due to invalid IL or missing references)
		//IL_001d: Expected O, but got Unknown
		try
		{
			foreach (Control item in (ArrangedElementCollection)parent.Controls)
			{
				Control val = item;
				if (!HTKKFormater.isNullString(nameCtl) && val.Name.Equals(nameCtl) && val is MaskedTextBox)
				{
					val.Text = value;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void SetDisplayKy(Control parent, bool display)
	{
		//IL_0017: Unknown result type (might be due to invalid IL or missing references)
		//IL_001d: Expected O, but got Unknown
		try
		{
			foreach (Control item in (ArrangedElementCollection)parent.Controls)
			{
				Control val = item;
				val.Visible = display;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void orderByListDataFile(List<HTKKTraCuuInfo> listDataTableFileSelect)
	{
		int num = 0;
		try
		{
			for (int i = 0; i < listDataTableFileSelect.Count; i++)
			{
				for (int j = i; j < listDataTableFileSelect.Count; j++)
				{
					num = j;
					if (listDataTableFileSelect[i].ListDateNumber.Count != 0 && (listDataTableFileSelect[j].ListDateNumber.Count == 0 || listDataTableFileSelect[i].ListDateNumber[0] > listDataTableFileSelect[j].ListDateNumber[0]))
					{
						HTKKTraCuuInfo hTKKTraCuuInfo = new HTKKTraCuuInfo();
						hTKKTraCuuInfo = listDataTableFileSelect[j];
						listDataTableFileSelect[j] = listDataTableFileSelect[i];
						listDataTableFileSelect[i] = hTKKTraCuuInfo;
					}
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
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
		//IL_0001: Unknown result type (might be due to invalid IL or missing references)
		//IL_0007: Expected O, but got Unknown
		//IL_0007: Unknown result type (might be due to invalid IL or missing references)
		//IL_000d: Expected O, but got Unknown
		//IL_000d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0013: Expected O, but got Unknown
		//IL_0013: Unknown result type (might be due to invalid IL or missing references)
		//IL_0019: Expected O, but got Unknown
		//IL_0019: Unknown result type (might be due to invalid IL or missing references)
		//IL_0020: Expected O, but got Unknown
		//IL_0032: Unknown result type (might be due to invalid IL or missing references)
		//IL_003c: Expected O, but got Unknown
		//IL_003d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0047: Expected O, but got Unknown
		//IL_0048: Unknown result type (might be due to invalid IL or missing references)
		//IL_0052: Expected O, but got Unknown
		//IL_0053: Unknown result type (might be due to invalid IL or missing references)
		//IL_005d: Expected O, but got Unknown
		//IL_005e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0068: Expected O, but got Unknown
		//IL_0069: Unknown result type (might be due to invalid IL or missing references)
		//IL_0073: Expected O, but got Unknown
		//IL_0074: Unknown result type (might be due to invalid IL or missing references)
		//IL_007e: Expected O, but got Unknown
		//IL_007f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0089: Expected O, but got Unknown
		//IL_008a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0094: Expected O, but got Unknown
		//IL_0095: Unknown result type (might be due to invalid IL or missing references)
		//IL_009f: Expected O, but got Unknown
		//IL_00a0: Unknown result type (might be due to invalid IL or missing references)
		//IL_00aa: Expected O, but got Unknown
		//IL_00ab: Unknown result type (might be due to invalid IL or missing references)
		//IL_00b5: Expected O, but got Unknown
		//IL_00b6: Unknown result type (might be due to invalid IL or missing references)
		//IL_00c0: Expected O, but got Unknown
		//IL_00c1: Unknown result type (might be due to invalid IL or missing references)
		//IL_00cb: Expected O, but got Unknown
		//IL_00cc: Unknown result type (might be due to invalid IL or missing references)
		//IL_00d6: Expected O, but got Unknown
		//IL_00d7: Unknown result type (might be due to invalid IL or missing references)
		//IL_00e1: Expected O, but got Unknown
		//IL_00e2: Unknown result type (might be due to invalid IL or missing references)
		//IL_00ec: Expected O, but got Unknown
		//IL_00ed: Unknown result type (might be due to invalid IL or missing references)
		//IL_00f7: Expected O, but got Unknown
		//IL_00f8: Unknown result type (might be due to invalid IL or missing references)
		//IL_0102: Expected O, but got Unknown
		//IL_0103: Unknown result type (might be due to invalid IL or missing references)
		//IL_010d: Expected O, but got Unknown
		//IL_010e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0118: Expected O, but got Unknown
		//IL_0119: Unknown result type (might be due to invalid IL or missing references)
		//IL_0123: Expected O, but got Unknown
		//IL_0124: Unknown result type (might be due to invalid IL or missing references)
		//IL_012e: Expected O, but got Unknown
		//IL_012f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0139: Expected O, but got Unknown
		//IL_013a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0144: Expected O, but got Unknown
		//IL_0145: Unknown result type (might be due to invalid IL or missing references)
		//IL_014f: Expected O, but got Unknown
		//IL_0150: Unknown result type (might be due to invalid IL or missing references)
		//IL_015a: Expected O, but got Unknown
		//IL_015b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0165: Expected O, but got Unknown
		//IL_0166: Unknown result type (might be due to invalid IL or missing references)
		//IL_0170: Expected O, but got Unknown
		//IL_0171: Unknown result type (might be due to invalid IL or missing references)
		//IL_017b: Expected O, but got Unknown
		//IL_017c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0186: Expected O, but got Unknown
		//IL_0187: Unknown result type (might be due to invalid IL or missing references)
		//IL_0191: Expected O, but got Unknown
		//IL_0388: Unknown result type (might be due to invalid IL or missing references)
		//IL_0392: Expected O, but got Unknown
		//IL_048e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0498: Expected O, but got Unknown
		//IL_083a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0844: Expected O, but got Unknown
		//IL_0852: Unknown result type (might be due to invalid IL or missing references)
		//IL_085c: Expected O, but got Unknown
		//IL_0e85: Unknown result type (might be due to invalid IL or missing references)
		//IL_0e8f: Expected O, but got Unknown
		//IL_0f2f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0f39: Expected O, but got Unknown
		DataGridViewCellStyle val = new DataGridViewCellStyle();
		DataGridViewCellStyle val2 = new DataGridViewCellStyle();
		DataGridViewCellStyle val3 = new DataGridViewCellStyle();
		DataGridViewCellStyle val4 = new DataGridViewCellStyle();
		DataGridViewCellStyle val5 = new DataGridViewCellStyle();
		ComponentResourceManager componentResourceManager = new ComponentResourceManager(typeof(frmTraCuuToKhai));
		groupBox1 = new GroupBox();
		label2 = new Label();
		DropSacThue = new ComboBox();
		label1 = new Label();
		DropLoaiToKhai = new ComboBox();
		groupBox2 = new GroupBox();
		panelLoadControls = new Panel();
		groupBox3 = new GroupBox();
		chkAll = new CheckBox();
		grv_listTable = new DataGridView();
		Column8 = new DataGridViewTextBoxColumn();
		Column5 = new DataGridViewCheckBoxColumn();
		Column7 = new DataGridViewTextBoxColumn();
		Column6 = new DataGridViewTextBoxColumn();
		VersionAPP = new DataGridViewTextBoxColumn();
		VersionXML = new DataGridViewTextBoxColumn();
		Column2 = new DataGridViewTextBoxColumn();
		LoaiTK = new DataGridViewTextBoxColumn();
		Column1 = new DataGridViewTextBoxColumn();
		lanBS = new DataGridViewTextBoxColumn();
		Column3 = new DataGridViewTextBoxColumn();
		Column4 = new DataGridViewTextBoxColumn();
		DataFile = new DataGridViewTextBoxColumn();
		btnDelete = new Button();
		cmdExit = new Button();
		btnOpen = new Button();
		btnTraCuu = new Button();
		panel1 = new Panel();
		label3 = new Label();
		panel2 = new Panel();
		label5 = new Label();
		label4 = new Label();
		((Control)groupBox1).SuspendLayout();
		((Control)groupBox2).SuspendLayout();
		((Control)groupBox3).SuspendLayout();
		((ISupportInitialize)grv_listTable).BeginInit();
		((Control)panel1).SuspendLayout();
		((Control)panel2).SuspendLayout();
		((Control)this).SuspendLayout();
		((Control)groupBox1).Controls.Add((Control)(object)label2);
		((Control)groupBox1).Controls.Add((Control)(object)DropSacThue);
		((Control)groupBox1).Controls.Add((Control)(object)label1);
		((Control)groupBox1).Controls.Add((Control)(object)DropLoaiToKhai);
		((Control)groupBox1).Location = new Point(9, 27);
		((Control)groupBox1).Name = "groupBox1";
		((Control)groupBox1).Size = new Size(477, 74);
		((Control)groupBox1).TabIndex = 0;
		groupBox1.TabStop = false;
		((Control)groupBox1).Text = "Chọn loại tờ khai";
		((Control)label2).AutoSize = true;
		((Control)label2).Location = new Point(24, 22);
		((Control)label2).Name = "label2";
		((Control)label2).Size = new Size(65, 13);
		((Control)label2).TabIndex = 3;
		((Control)label2).Text = "Loại tờ khai:";
		((ListControl)DropSacThue).FormattingEnabled = true;
		((Control)DropSacThue).Location = new Point(88, 17);
		((Control)DropSacThue).Name = "DropSacThue";
		((Control)DropSacThue).Size = new Size(369, 21);
		((Control)DropSacThue).TabIndex = 1;
		DropSacThue.SelectedIndexChanged += DropSacThue_SelectedIndexChanged;
		((Control)DropSacThue).KeyPress += new KeyPressEventHandler(DropSacThue_KeyPress);
		((Control)label1).AutoSize = true;
		((Control)label1).Location = new Point(24, 49);
		((Control)label1).Name = "label1";
		((Control)label1).Size = new Size(64, 13);
		((Control)label1).TabIndex = 1;
		((Control)label1).Text = "Tên tờ khai:";
		((ListControl)DropLoaiToKhai).FormattingEnabled = true;
		((Control)DropLoaiToKhai).Location = new Point(88, 45);
		((Control)DropLoaiToKhai).Name = "DropLoaiToKhai";
		((Control)DropLoaiToKhai).Size = new Size(369, 21);
		((Control)DropLoaiToKhai).TabIndex = 2;
		DropLoaiToKhai.SelectedIndexChanged += DropLoaiToKhai_SelectedIndexChanged;
		((Control)DropLoaiToKhai).Enter += DropLoaiToKhai_Enter;
		((Control)DropLoaiToKhai).KeyPress += new KeyPressEventHandler(DropLoaiToKhai_KeyPress);
		((Control)groupBox2).Controls.Add((Control)(object)panelLoadControls);
		((Control)groupBox2).Location = new Point(495, 26);
		((Control)groupBox2).Name = "groupBox2";
		((Control)groupBox2).Size = new Size(522, 75);
		((Control)groupBox2).TabIndex = 3;
		groupBox2.TabStop = false;
		((Control)groupBox2).Text = "Chọn kỳ tính thuế";
		((Control)panelLoadControls).Dock = (DockStyle)5;
		((Control)panelLoadControls).Location = new Point(3, 16);
		((Control)panelLoadControls).Name = "panelLoadControls";
		((Control)panelLoadControls).Size = new Size(516, 56);
		((Control)panelLoadControls).TabIndex = 0;
		((Control)groupBox3).Anchor = (AnchorStyles)15;
		((Control)groupBox3).Controls.Add((Control)(object)chkAll);
		((Control)groupBox3).Controls.Add((Control)(object)grv_listTable);
		((Control)groupBox3).Location = new Point(10, 105);
		((Control)groupBox3).Name = "groupBox3";
		((Control)groupBox3).Size = new Size(1007, 349);
		((Control)groupBox3).TabIndex = 10;
		groupBox3.TabStop = false;
		((Control)groupBox3).Text = "Kết quả";
		((Control)chkAll).AutoSize = true;
		((Control)chkAll).Location = new Point(68, 22);
		((Control)chkAll).Name = "chkAll";
		((Control)chkAll).Size = new Size(15, 14);
		((Control)chkAll).TabIndex = 21;
		((ButtonBase)chkAll).UseVisualStyleBackColor = true;
		chkAll.CheckedChanged += chkAll_CheckedChanged;
		((Control)chkAll).Click += chkAll_Click;
		grv_listTable.AllowUserToAddRows = false;
		grv_listTable.AllowUserToDeleteRows = false;
		grv_listTable.AllowUserToResizeColumns = false;
		grv_listTable.AllowUserToResizeRows = false;
		grv_listTable.AutoSizeRowsMode = (DataGridViewAutoSizeRowsMode)7;
		grv_listTable.BackgroundColor = Color.White;
		grv_listTable.ColumnHeadersHeightSizeMode = (DataGridViewColumnHeadersHeightSizeMode)2;
		grv_listTable.Columns.AddRange((DataGridViewColumn[])(object)new DataGridViewColumn[13]
		{
			(DataGridViewColumn)Column8,
			(DataGridViewColumn)Column5,
			(DataGridViewColumn)Column7,
			(DataGridViewColumn)Column6,
			(DataGridViewColumn)VersionAPP,
			(DataGridViewColumn)VersionXML,
			(DataGridViewColumn)Column2,
			(DataGridViewColumn)LoaiTK,
			(DataGridViewColumn)Column1,
			(DataGridViewColumn)lanBS,
			(DataGridViewColumn)Column3,
			(DataGridViewColumn)Column4,
			(DataGridViewColumn)DataFile
		});
		((Control)grv_listTable).Location = new Point(8, 18);
		((Control)grv_listTable).Name = "grv_listTable";
		grv_listTable.RowHeadersVisible = false;
		val.WrapMode = (DataGridViewTriState)1;
		grv_listTable.RowsDefaultCellStyle = val;
		grv_listTable.SelectionMode = (DataGridViewSelectionMode)1;
		((Control)grv_listTable).Size = new Size(993, 325);
		((Control)grv_listTable).TabIndex = 20;
		grv_listTable.CellClick += new DataGridViewCellEventHandler(grv_listTable_CellClick);
		grv_listTable.CellContentClick += new DataGridViewCellEventHandler(grv_listTable_CellContentClick);
		((DataGridViewColumn)Column8).DataPropertyName = "STT";
		val2.Alignment = (DataGridViewContentAlignment)32;
		((DataGridViewBand)Column8).DefaultCellStyle = val2;
		((DataGridViewColumn)Column8).HeaderText = "STT";
		((DataGridViewColumn)Column8).Name = "Column8";
		((DataGridViewColumn)Column8).Width = 40;
		((DataGridViewColumn)Column5).DataPropertyName = "CheckValue";
		((DataGridViewColumn)Column5).HeaderText = "";
		((DataGridViewColumn)Column5).Name = "Column5";
		((DataGridViewColumn)Column5).Width = 50;
		((DataGridViewColumn)Column7).DataPropertyName = "ID";
		((DataGridViewColumn)Column7).HeaderText = "ID";
		((DataGridViewColumn)Column7).Name = "Column7";
		((DataGridViewBand)Column7).Visible = false;
		((DataGridViewColumn)Column6).DataPropertyName = "FileName";
		((DataGridViewColumn)Column6).HeaderText = "FileName";
		((DataGridViewColumn)Column6).Name = "Column6";
		((DataGridViewBand)Column6).Visible = false;
		((DataGridViewColumn)VersionAPP).DataPropertyName = "VersionAPP";
		((DataGridViewColumn)VersionAPP).HeaderText = "Phiên bản ứng dụng";
		((DataGridViewColumn)VersionAPP).Name = "VersionAPP";
		((DataGridViewBand)VersionAPP).Visible = false;
		((DataGridViewColumn)VersionXML).DataPropertyName = "VersionXML";
		((DataGridViewColumn)VersionXML).HeaderText = "Phiên bản XML";
		((DataGridViewColumn)VersionXML).Name = "VersionXML";
		((DataGridViewBand)VersionXML).Visible = false;
		((DataGridViewColumn)Column2).AutoSizeMode = (DataGridViewAutoSizeColumnMode)16;
		((DataGridViewColumn)Column2).DataPropertyName = "Caption";
		((DataGridViewColumn)Column2).HeaderText = "Tên tờ khai";
		((DataGridViewColumn)Column2).Name = "Column2";
		((DataGridViewBand)Column2).ReadOnly = true;
		((DataGridViewColumn)LoaiTK).DataPropertyName = "LoaiTK";
		((DataGridViewColumn)LoaiTK).HeaderText = "Loại tờ khai";
		((DataGridViewColumn)LoaiTK).Name = "LoaiTK";
		((DataGridViewBand)LoaiTK).Visible = false;
		((DataGridViewColumn)Column1).DataPropertyName = "KyTinhThueShow";
		val3.Alignment = (DataGridViewContentAlignment)32;
		((DataGridViewBand)Column1).DefaultCellStyle = val3;
		((DataGridViewColumn)Column1).HeaderText = "Kỳ tính thuế";
		((DataGridViewColumn)Column1).Name = "Column1";
		((DataGridViewColumn)Column1).Width = 150;
		((DataGridViewColumn)lanBS).DataPropertyName = "Lan";
		((DataGridViewColumn)lanBS).HeaderText = "Lần bổ sung";
		((DataGridViewColumn)lanBS).Name = "lanBS";
		((DataGridViewColumn)lanBS).Width = 95;
		((DataGridViewColumn)Column3).DataPropertyName = "ThueKhauTru";
		val4.Alignment = (DataGridViewContentAlignment)64;
		((DataGridViewBand)Column3).DefaultCellStyle = val4;
		((DataGridViewColumn)Column3).HeaderText = "Thuế khấu trừ";
		((DataGridViewColumn)Column3).Name = "Column3";
		((DataGridViewColumn)Column3).Width = 120;
		((DataGridViewColumn)Column4).DataPropertyName = "ThuePhaiNop";
		val5.Alignment = (DataGridViewContentAlignment)64;
		((DataGridViewBand)Column4).DefaultCellStyle = val5;
		((DataGridViewColumn)Column4).HeaderText = "Thuế phải nộp";
		((DataGridViewColumn)Column4).Name = "Column4";
		((DataGridViewColumn)Column4).Width = 120;
		((DataGridViewColumn)DataFile).DataPropertyName = "DataFile";
		((DataGridViewColumn)DataFile).HeaderText = "DataFile";
		((DataGridViewColumn)DataFile).Name = "DataFile";
		((DataGridViewBand)DataFile).Visible = false;
		((Control)btnDelete).Anchor = (AnchorStyles)10;
		((ButtonBase)btnDelete).ImageAlign = (ContentAlignment)16;
		((Control)btnDelete).Location = new Point(827, 461);
		((Control)btnDelete).Name = "btnDelete";
		((Control)btnDelete).Size = new Size(91, 30);
		((Control)btnDelete).TabIndex = 8;
		((Control)btnDelete).Text = "&Xóa";
		((ButtonBase)btnDelete).UseVisualStyleBackColor = true;
		((Control)btnDelete).Click += btnDelete_Click;
		((Control)cmdExit).Anchor = (AnchorStyles)10;
		((ButtonBase)cmdExit).ImageAlign = (ContentAlignment)16;
		((Control)cmdExit).Location = new Point(923, 461);
		((Control)cmdExit).Name = "cmdExit";
		((Control)cmdExit).Size = new Size(91, 30);
		((Control)cmdExit).TabIndex = 9;
		((Control)cmdExit).Text = "&Thoát";
		((ButtonBase)cmdExit).UseVisualStyleBackColor = true;
		((Control)cmdExit).Click += cmdExit_Click;
		((Control)btnOpen).Anchor = (AnchorStyles)10;
		((Control)btnOpen).Location = new Point(731, 461);
		((Control)btnOpen).Name = "btnOpen";
		((Control)btnOpen).Size = new Size(91, 30);
		((Control)btnOpen).TabIndex = 7;
		((Control)btnOpen).Text = "&Mở";
		((ButtonBase)btnOpen).UseVisualStyleBackColor = true;
		((Control)btnOpen).Click += btnOpen_Click;
		((Control)btnTraCuu).Anchor = (AnchorStyles)10;
		((Control)btnTraCuu).Location = new Point(634, 461);
		((Control)btnTraCuu).Name = "btnTraCuu";
		((Control)btnTraCuu).Size = new Size(91, 30);
		((Control)btnTraCuu).TabIndex = 6;
		((Control)btnTraCuu).Text = "Tra &cứu";
		((ButtonBase)btnTraCuu).UseVisualStyleBackColor = true;
		((Control)btnTraCuu).Click += btnTraCuu_Click;
		((Control)panel1).BackgroundImage = (Image)componentResourceManager.GetObject("panel1.BackgroundImage");
		((Control)panel1).BackgroundImageLayout = (ImageLayout)3;
		((Control)panel1).Controls.Add((Control)(object)label3);
		((Control)panel1).Location = new Point(2, 0);
		((Control)panel1).Name = "panel1";
		((Control)panel1).Size = new Size(1015, 21);
		((Control)panel1).TabIndex = 25;
		((Control)label3).AutoSize = true;
		((Control)label3).BackColor = Color.Transparent;
		((Control)label3).Font = new Font("Tahoma", 9.75f, (FontStyle)1, (GraphicsUnit)3, (byte)0);
		((Control)label3).ForeColor = Color.White;
		((Control)label3).Location = new Point(13, 3);
		((Control)label3).Name = "label3";
		((Control)label3).Size = new Size(168, 16);
		((Control)label3).TabIndex = 0;
		((Control)label3).Text = "Tra cứu thông tin tờ khai";
		((Control)panel2).Controls.Add((Control)(object)label5);
		((Control)panel2).Controls.Add((Control)(object)label4);
		((Control)panel2).Location = new Point(14, 460);
		((Control)panel2).Name = "panel2";
		((Control)panel2).Size = new Size(469, 44);
		((Control)panel2).TabIndex = 26;
		((Control)label5).AutoSize = true;
		((Control)label5).Location = new Point(5, 24);
		((Control)label5).Name = "label5";
		((Control)label5).Size = new Size(275, 13);
		((Control)label5).TabIndex = 1;
		((Control)label5).Text = "(*) Xóa: Ứng dụng hỗ trợ xóa các tờ khai được tích chọn";
		((Control)label4).AutoSize = true;
		((Control)label4).Location = new Point(5, 6);
		((Control)label4).Name = "label4";
		((Control)label4).Size = new Size(261, 13);
		((Control)label4).TabIndex = 0;
		((Control)label4).Text = "(*) Mở:  Ứng dụng hỗ trợ mở một tờ khai được bôi màu";
		((ContainerControl)this).AutoScaleDimensions = new SizeF(6f, 13f);
		((ContainerControl)this).AutoScaleMode = (AutoScaleMode)1;
		((Form)this).ClientSize = new Size(1025, 510);
		((Form)this).ControlBox = false;
		((Control)this).Controls.Add((Control)(object)panel2);
		((Control)this).Controls.Add((Control)(object)panel1);
		((Control)this).Controls.Add((Control)(object)btnDelete);
		((Control)this).Controls.Add((Control)(object)cmdExit);
		((Control)this).Controls.Add((Control)(object)btnOpen);
		((Control)this).Controls.Add((Control)(object)btnTraCuu);
		((Control)this).Controls.Add((Control)(object)groupBox3);
		((Control)this).Controls.Add((Control)(object)groupBox2);
		((Control)this).Controls.Add((Control)(object)groupBox1);
		((Form)this).FormBorderStyle = (FormBorderStyle)0;
		((Control)this).Name = "frmTraCuuToKhai";
		((Form)this).StartPosition = (FormStartPosition)4;
		((Control)this).Text = "`";
		((Form)this).Load += frmTraCuuToKhai_Load;
		((Control)groupBox1).ResumeLayout(false);
		((Control)groupBox1).PerformLayout();
		((Control)groupBox2).ResumeLayout(false);
		((Control)groupBox3).ResumeLayout(false);
		((Control)groupBox3).PerformLayout();
		((ISupportInitialize)grv_listTable).EndInit();
		((Control)panel1).ResumeLayout(false);
		((Control)panel1).PerformLayout();
		((Control)panel2).ResumeLayout(false);
		((Control)panel2).PerformLayout();
		((Control)this).ResumeLayout(false);
	}
}
