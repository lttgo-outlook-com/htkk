using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Security;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;
using FarPoint.Excel;
using FarPoint.Win.Spread;
using TAX.HTKK.Business;
using TAX.HTKK.Utilities.Control;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace HTKK;

public class WinFnProcess
{
	public static void WinLoadTTinChungInfo(string maSothue, FpSpread winFpSpread)
	{
		try
		{
			winFpSpread.ActiveSheetIndex = winFpSpread.Sheets.Count - 1;
			PropertyInfo[] properties = typeof(HTKKCMCTTinChungInfo).GetProperties();
			foreach (PropertyInfo propertyInfo in properties)
			{
				if (WinSettings.DictTTinChung.ContainsKey(propertyInfo.Name))
				{
					object value = propertyInfo.GetValue(WinSettings.TTinChung, null);
					if (propertyInfo.Name == "mst")
					{
						WinSpreadProcess.FillMstHeader(winFpSpread.ActiveSheet, value.ToString());
					}
					else
					{
						WinSpreadProcess.WinSetValueByIndex(winFpSpread.ActiveSheet, WinSettings.DictTTinChung[propertyInfo.Name], value.ToString());
					}
				}
			}
			if (winFpSpread != null && winFpSpread.Sheets.Count > 0)
			{
				winFpSpread.Sheets[0].RecalculateAll();
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static bool WinLoadFromFile(FpSpread winFpSpread, string TemplateFile)
	{
		//IL_0287: Unknown result type (might be due to invalid IL or missing references)
		try
		{
			bool flag = false;
			((Control)winFpSpread).Controls.Clear();
			Stopwatch stopwatch = Stopwatch.StartNew();
			stopwatch.Reset();
			stopwatch.Start();
			flag = winFpSpread.OpenExcel(TemplateFile, ExcelOpenFlags.DoNotRecalculateAfterLoad);
			if (WinSettings.ToKhaiInfo.KieuKHBS == "3" && WinSettings.TTPeriodInfo.LoaiToKhai == "B" && flag)
			{
				string text = "";
				FpSpread fpSpread = new FpSpread();
				text = ("I7,I4,C7,I8,C8".Contains(WinSettings.MenuID) ? TemplateFile.Replace(WinSettings.ToKhaiInfo.FolderName, "KHBS_TT80_DVT") : ((!"D1,N8,G9,H2,H4,T8,E8,L8,L7,N2,G2,R5,Y6,Z8,Z5".Contains(WinSettings.MenuID)) ? TemplateFile.Replace(WinSettings.ToKhaiInfo.FolderName, "KHBS_TT80") : TemplateFile.Replace(WinSettings.ToKhaiInfo.FolderName, "KHBS_DIA_BAN")));
				text = text.Replace(WinSettings.ToKhaiInfo.VersionXML.Replace(".", "") + ".xls", WinSettings.ToKhaiInfo.VersionKBSXML.Replace(".", "") + ".xls");
				if (File.Exists(text) && fpSpread.OpenExcel(text, ExcelOpenFlags.DoNotRecalculateAfterLoad))
				{
					winFpSpread.Sheets[winFpSpread.Sheets.Count - 2] = fpSpread.Sheets[1];
					winFpSpread.Sheets[winFpSpread.Sheets.Count - 3] = fpSpread.Sheets[0];
					fpSpread = null;
				}
			}
			stopwatch.Stop();
			long elapsedMilliseconds = stopwatch.ElapsedMilliseconds;
			HTKKFile.LogRun(MethodBase.GetCurrentMethod().Name, "Load Template - " + elapsedMilliseconds);
			if (flag & (WinSettings.ToKhaiInfo.FolderName != "HEADER"))
			{
				WinLoadTTinChungInfo(WinSettings.HTKKHeaderInfo.mst, winFpSpread);
			}
			else
			{
				winFpSpread.Sheets[0].Cells[1, 19].Text = WinSettings.HTKKHeaderInfo.mst;
			}
			return flag;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			frmMessage frmMessage = new frmMessage("0105", "3", "3", "Thông báo");
			((Form)frmMessage).ShowDialog();
			return false;
		}
	}

	public static HTKKCMCCellInfo GetStaticCell(int iRow, int iCol, List<HTKKCMCCellInfo> ListStaticCells)
	{
		HTKKCMCCellInfo result = null;
		try
		{
			foreach (HTKKCMCCellInfo ListStaticCell in ListStaticCells)
			{
				if (ListStaticCell.CellTemplate.RowIndex == iRow + 1 && ListStaticCell.CellTemplate.ColIndex == iCol + 1)
				{
					result = ListStaticCell;
					break;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
		return result;
	}

	public static HTKKCMCSheetInfo GetHTKKSheetInfo(HTKKCMCToKhaiInfo ToKhaiInfo, XmlNodeList elemList, int i, bool getTemplate, HTKKPeriodInfo periodInfo)
	{
		try
		{
			string text = ((elemList[i].Attributes["DataFile"] == null) ? "" : ((XmlNode)elemList[i].Attributes["DataFile"]).Value);
			HTKKCMCSheetInfo hTKKCMCSheetInfo = new HTKKCMCSheetInfo();
			hTKKCMCSheetInfo.SheetID = ((elemList[i].Attributes["ID"] == null) ? "" : ((XmlNode)elemList[i].Attributes["ID"]).Value);
			hTKKCMCSheetInfo.SheetName = ((elemList[i].Attributes["SheetName"] == null) ? "" : ((XmlNode)elemList[i].Attributes["SheetName"]).Value);
			hTKKCMCSheetInfo.Caption = ((elemList[i].Attributes["Caption"] == null) ? "" : ((XmlNode)elemList[i].Attributes["Caption"]).Value);
			hTKKCMCSheetInfo.DataFile = text;
			hTKKCMCSheetInfo.Active = ((elemList[i].Attributes["Active"] == null) ? "" : ((XmlNode)elemList[i].Attributes["Active"]).Value);
			hTKKCMCSheetInfo.SheetIndex = short.Parse(i.ToString());
			hTKKCMCSheetInfo.ShowSearchButton = ((elemList[i].Attributes["ShowSearchButton"] == null) ? "" : ((XmlNode)elemList[i].Attributes["ShowSearchButton"]).Value);
			hTKKCMCSheetInfo.ShowTaiBangKeButton = ((elemList[i].Attributes["ShowTaiBangKeButton"] == null) ? "" : ((XmlNode)elemList[i].Attributes["ShowTaiBangKeButton"]).Value);
			hTKKCMCSheetInfo.strPrintBarcode = ((elemList[i].Attributes["PrintBarCode"] == null) ? "" : ((XmlNode)elemList[i].Attributes["PrintBarCode"]).Value);
			hTKKCMCSheetInfo.XMLnode = ((elemList[i].Attributes["XMLnode"] == null) ? "" : ((XmlNode)elemList[i].Attributes["XMLnode"]).Value);
			hTKKCMCSheetInfo.CoDieuChinhSoLieu = ((elemList[i].Attributes["CoDieuChinhSoLieu"] == null) ? (-1) : int.Parse(((XmlNode)elemList[i].Attributes["CoDieuChinhSoLieu"]).Value));
			hTKKCMCSheetInfo.RowTTDLT = ((elemList[i].Attributes["RowTTDLT"] == null) ? "" : ((XmlNode)elemList[i].Attributes["RowTTDLT"]).Value);
			hTKKCMCSheetInfo.RowBreakPage = ((elemList[i].Attributes["RowBreakPage"] == null) ? "" : ((XmlNode)elemList[i].Attributes["RowBreakPage"]).Value);
			hTKKCMCSheetInfo.TaiDuLieuLon = ((elemList[i].Attributes["TaiDuLieuLon"] == null) ? "" : ((XmlNode)elemList[i].Attributes["TaiDuLieuLon"]).Value);
			hTKKCMCSheetInfo.TaiDuLieuRatLon = ((elemList[i].Attributes["TaiDuLieuRatLon"] == null) ? "" : ((XmlNode)elemList[i].Attributes["TaiDuLieuRatLon"]).Value);
			HTKKCMCSheetInfo hTKKCMCSheetInfo2 = hTKKCMCSheetInfo;
			bool flag = false;
			string text2 = "";
			string fileNameBySheetNew = WinDoWorkEvent.GetFileNameBySheetNew(WinSettings.TTPeriodInfo, text, HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst + "\\", ToKhaiInfo.VersionXML);
			if (ToKhaiInfo.MenuID.Equals("100_1"))
			{
				text2 = HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst + "\\Header.xml";
			}
			else if (periodInfo.LoaiToKhai == "B" && WinSettings.ToKhaiInfo.KieuKHBS.Equals(""))
			{
				text2 = (("N7~M1~M2~C6~C7~T5~H4~O1~O2~O3~O4~O5~O6~Q4~P7~U1~U8~X5~E9~R9~Y3~Y4~Z1~".IndexOf(WinSettings.MenuID) < 0) ? WinDoWorkEvent.GetFileNameLoadKhbs(WinSettings.TTPeriodInfo, text, HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst + "\\", isLoadSavedFile: true) : WinDoWorkEvent.GetFileNameLoadKhbsNew(WinSettings.TTPeriodInfo, fileNameBySheetNew, HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst + "\\", isLoadSavedFile: true));
				if (!File.Exists(text2))
				{
					text2 = (("N7~".IndexOf(WinSettings.MenuID) >= 0) ? GeneralFunctions.GetXMLFileName(ToKhaiInfo.InterfaceTemplateXml, text, ToKhaiInfo.VersionKBSXML, "xml") : WinDoWorkEvent.GetFileNameLoadKhbsNew(WinSettings.TTPeriodInfo, fileNameBySheetNew, HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst + "\\", isLoadSavedFile: false));
				}
			}
			else if (periodInfo.LoaiToKhai == "B" && "30,I5,K9,K1,M7,K6,K8,73,55,71,31,P2,X6".Contains(WinSettings.MenuID) && i > 0)
			{
				text2 = (("N7~M1~M2~C6~C7~T5~H4~O1~O2~O3~O4~O5~O6~Q4~P7~U1~U8~X5~E9~R9~Y3~Y4~Z1~".IndexOf(WinSettings.MenuID) < 0) ? WinDoWorkEvent.GetFileNameLoadKhbs(WinSettings.TTPeriodInfo, text, HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst + "\\", isLoadSavedFile: true) : WinDoWorkEvent.GetFileNameLoadKhbsNew(WinSettings.TTPeriodInfo, fileNameBySheetNew, HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst + "\\", isLoadSavedFile: true));
				if (!File.Exists(text2))
				{
					text2 = (("N7~M1~M2~C6~C7~T5~H4~O1~O2~O3~O4~O5~O6~Q4~P7~U1~U8~X5~E9~R9~Y3~Y4~Z1~".IndexOf(WinSettings.MenuID) < 0) ? WinDoWorkEvent.GetFileNameLoadKhbs(WinSettings.TTPeriodInfo, text, HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst + "\\", isLoadSavedFile: false) : WinDoWorkEvent.GetFileNameLoadKhbsNew(WinSettings.TTPeriodInfo, fileNameBySheetNew, HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst + "\\", isLoadSavedFile: false));
					if (!File.Exists(text2))
					{
						text2 = ((!text.Contains("KHBS")) ? GeneralFunctions.GetXMLFileName(ToKhaiInfo.InterfaceTemplateXml, text, ToKhaiInfo.VersionXML, "xml") : GeneralFunctions.GetXMLFileName(ToKhaiInfo.InterfaceTemplateXml, text, ToKhaiInfo.VersionKBSXML, "xml"));
						flag = true;
					}
				}
			}
			else if (periodInfo.LoaiToKhai == "T")
			{
				text2 = (("N7~M1~M2~C6~C7~T5~H4~O1~O2~O3~O4~O5~O6~Q4~P7~U1~U8~X5~E9~R9~Y3~Y4~Z1~".IndexOf(WinSettings.MenuID) < 0) ? WinDoWorkEvent.GetFileNameLoadKhbs(WinSettings.TTPeriodInfo, text, HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst + "\\", isLoadSavedFile: true) : WinDoWorkEvent.GetFileNameLoadKhbsNew(WinSettings.TTPeriodInfo, fileNameBySheetNew, HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst + "\\", isLoadSavedFile: true));
				if (!File.Exists(text2))
				{
					text2 = ((!(WinSettings.MenuID != "N7")) ? GeneralFunctions.GetXMLFileName(ToKhaiInfo.InterfaceTemplateXml, text, ToKhaiInfo.VersionKBSXML, "xml") : WinDoWorkEvent.GetFileNameBySheet(WinSettings.TTPeriodInfo, text, HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst + "\\", ToKhaiInfo.VersionXML));
				}
			}
			else
			{
				text2 = WinDoWorkEvent.GetFileNameBySheet(WinSettings.TTPeriodInfo, text, HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst + "\\", ToKhaiInfo.VersionXML);
			}
			if (getTemplate)
			{
				string version = ToKhaiInfo.VersionXML;
				if (hTKKCMCSheetInfo2.SheetID.Equals("KHBS"))
				{
					version = ToKhaiInfo.VersionKBSXML;
				}
				if (hTKKCMCSheetInfo2.SheetID.Equals("01-KHBS"))
				{
					version = ToKhaiInfo.VersionKBSXML;
				}
				if (hTKKCMCSheetInfo2.SheetID.Equals("01-1-KHBS"))
				{
					version = ToKhaiInfo.VersionKBSXML;
				}
				if (hTKKCMCSheetInfo2.SheetID.Equals("01-KHBS"))
				{
					if ("I7,I4,C7,I8,C8".Contains(WinSettings.MenuID))
					{
						hTKKCMCSheetInfo2.GetDetailSheetInfo(GeneralFunctions.GetXMLFileName(ToKhaiInfo.InterfaceTemplateXml, "KHBS_01_TT80_DVT", version, "xml"));
					}
					else if ("D1,N8,G9,H2,H4,T8,E8,L8,L7,N2,G2,R5,Y6,Z8,Z5".Contains(WinSettings.MenuID))
					{
						hTKKCMCSheetInfo2.GetDetailSheetInfo(GeneralFunctions.GetXMLFileName(ToKhaiInfo.InterfaceTemplateXml, "KHBS_01_TT80_DIA_BAN", version, "xml"));
					}
					else
					{
						hTKKCMCSheetInfo2.GetDetailSheetInfo(GeneralFunctions.GetXMLFileName(ToKhaiInfo.InterfaceTemplateXml, "KHBS_01_TT80", version, "xml"));
					}
				}
				else if (hTKKCMCSheetInfo2.SheetID.Equals("01-1-KHBS"))
				{
					if ("I7,I4,C7,I8,C8".Contains(WinSettings.MenuID))
					{
						hTKKCMCSheetInfo2.GetDetailSheetInfo(GeneralFunctions.GetXMLFileName(ToKhaiInfo.InterfaceTemplateXml, "KHBS_01_1_TT80_DVT", version, "xml"));
					}
					else
					{
						hTKKCMCSheetInfo2.GetDetailSheetInfo(GeneralFunctions.GetXMLFileName(ToKhaiInfo.InterfaceTemplateXml, "KHBS_01_1_TT80", version, "xml"));
					}
				}
				else
				{
					hTKKCMCSheetInfo2.GetDetailSheetInfo(GeneralFunctions.GetXMLFileName(ToKhaiInfo.InterfaceTemplateXml, text, version, "xml"));
				}
				hTKKCMCSheetInfo2.IsReadDataFile = false;
			}
			else if (File.Exists(text2))
			{
				hTKKCMCSheetInfo2.IsReadDataFile = true;
				if (flag)
				{
					hTKKCMCSheetInfo2.IsReadDataFile = false;
				}
				hTKKCMCSheetInfo2.GetDetailSheetInfo(text2);
				HTKKCMCXML hTKKCMCXML = new HTKKCMCXML(text2);
				XmlNode val = ((XmlNode)hTKKCMCXML.XDocument).SelectSingleNode("/Sections");
				if (((i == 0) & WinSettings.bTimKiem) && periodInfo.TieuMucHachToanVal == "" && ((val.Attributes["tieuMucHachToan"] != null) & (periodInfo.TieuMucHachToanVal.Length > 0)))
				{
					periodInfo.TieuMucHachToanVal = ((XmlNode)val.Attributes["tieuMucHachToan"]).Value;
				}
				if (HTKKFormater.isNullString(periodInfo.NgayLapKHBS) && hTKKCMCSheetInfo2.SheetID.Contains("KHBS") && periodInfo.LoaiToKhai == "B" && WinSettings.ToKhaiInfo.KieuKHBS != "3")
				{
					string valueByAttributes = HTKKFile.GetValueByAttributes(text2, "/Sections/Section/Cells/Cell[@Path='Path1']");
					if (WinSettings.MenuID == "01" || WinSettings.MenuID == "I5")
					{
						valueByAttributes = HTKKFile.GetValueByAttributes(text2, "/Sections/Section/Cells/Cell[@Path='PathB47']");
					}
					periodInfo.NgayLapKHBS = valueByAttributes;
					WinSettings.TTPeriodInfo.NgayLapKHBS = valueByAttributes;
				}
				if (HTKKFormater.isNullString(periodInfo.NgayLapKHBS) && hTKKCMCSheetInfo2.SheetName == "01-KHBS" && periodInfo.LoaiToKhai == "B" && WinSettings.ToKhaiInfo.KieuKHBS == "3")
				{
					string valueByAttributes = HTKKFile.GetValueByAttributes(text2, "/Sections/Section/Cells/Cell[@Path='PL01_KHBS/Muc_A/Muc_I/Muc_3/ct3a_1']");
					if (string.IsNullOrEmpty(valueByAttributes))
					{
						periodInfo.NgayLapKHBS = valueByAttributes;
						WinSettings.TTPeriodInfo.NgayLapKHBS = valueByAttributes;
					}
				}
			}
			else if (periodInfo.LoaiToKhai == "B")
			{
				hTKKCMCSheetInfo2.IsReadDataFile = false;
				if (((XmlNode)WinSettings.xmlMenuNode.Attributes["KieuKHBS"]).Value != "3")
				{
					if (i == 0)
					{
						hTKKCMCSheetInfo2 = new HTKKCMCSheetInfo(WinSettings.KHBS_LanTruoc);
					}
					else
					{
						hTKKCMCSheetInfo2.GetDetailSheetInfo(GeneralFunctions.GetXMLFileName(ToKhaiInfo.InterfaceTemplateXml, text, ToKhaiInfo.VersionKBSXML, "xml"));
					}
				}
				else if (!hTKKCMCSheetInfo2.SheetID.Contains("KHBS"))
				{
					HTKKCMCSheetInfo hTKKCMCSheetInfo3 = ReadSheet_LanTruoc((short)i);
					if (hTKKCMCSheetInfo3 != null)
					{
						hTKKCMCSheetInfo2 = new HTKKCMCSheetInfo(hTKKCMCSheetInfo3);
					}
					else
					{
						hTKKCMCSheetInfo2.GetDetailSheetInfo(GeneralFunctions.GetXMLFileName(ToKhaiInfo.InterfaceTemplateXml, text, ToKhaiInfo.VersionXML, "xml"));
					}
				}
				else if (hTKKCMCSheetInfo2.SheetID.Equals("01-KHBS"))
				{
					if ("I7,I4,C7,I8,C8".Contains(WinSettings.MenuID))
					{
						hTKKCMCSheetInfo2.GetDetailSheetInfo(GeneralFunctions.GetXMLFileName(ToKhaiInfo.InterfaceTemplateXml, "KHBS_01_TT80_DVT", ToKhaiInfo.VersionKBSXML, "xml"));
					}
					else if ("D1,N8,G9,H2,H4,T8,E8,L8,L7,N2,G2,R5,Y6,Z8,Z5".Contains(WinSettings.MenuID))
					{
						hTKKCMCSheetInfo2.GetDetailSheetInfo(GeneralFunctions.GetXMLFileName(ToKhaiInfo.InterfaceTemplateXml, "KHBS_01_TT80_DIA_BAN", ToKhaiInfo.VersionKBSXML, "xml"));
					}
					else
					{
						hTKKCMCSheetInfo2.GetDetailSheetInfo(GeneralFunctions.GetXMLFileName(ToKhaiInfo.InterfaceTemplateXml, "KHBS_01_TT80", ToKhaiInfo.VersionKBSXML, "xml"));
					}
				}
				else if (hTKKCMCSheetInfo2.SheetID.Equals("01-1-KHBS"))
				{
					if ("I7,I4,C7,I8,C8".Contains(WinSettings.MenuID))
					{
						hTKKCMCSheetInfo2.GetDetailSheetInfo(GeneralFunctions.GetXMLFileName(ToKhaiInfo.InterfaceTemplateXml, "KHBS_01_1_TT80_DVT", ToKhaiInfo.VersionKBSXML, "xml"));
					}
					else
					{
						hTKKCMCSheetInfo2.GetDetailSheetInfo(GeneralFunctions.GetXMLFileName(ToKhaiInfo.InterfaceTemplateXml, "KHBS_01_1_TT80", ToKhaiInfo.VersionKBSXML, "xml"));
					}
				}
				else
				{
					hTKKCMCSheetInfo2.GetDetailSheetInfo(GeneralFunctions.GetXMLFileName(ToKhaiInfo.InterfaceTemplateXml, text, ToKhaiInfo.VersionKBSXML, "xml"));
				}
			}
			else
			{
				hTKKCMCSheetInfo2.GetDetailSheetInfo(GeneralFunctions.GetXMLFileName(ToKhaiInfo.InterfaceTemplateXml, text, ToKhaiInfo.VersionXML, "xml"));
				hTKKCMCSheetInfo2.IsReadDataFile = false;
			}
			return hTKKCMCSheetInfo2;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public static HTKKCMCToKhaiInfo GetToKhaiInfo(XmlNode xmlNode, HTKKPeriodInfo periodInfo)
	{
		//IL_000e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0014: Expected O, but got Unknown
		//IL_0106: Unknown result type (might be due to invalid IL or missing references)
		//IL_010d: Expected O, but got Unknown
		//IL_038a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0391: Expected O, but got Unknown
		HTKKCMCToKhaiInfo hTKKCMCToKhaiInfo = new HTKKCMCToKhaiInfo();
		try
		{
			hTKKCMCToKhaiInfo = new HTKKCMCToKhaiInfo();
			XmlDocument val = new XmlDocument();
			val.LoadXml(xmlNode.OuterXml);
			hTKKCMCToKhaiInfo.MenuID = ((XmlNode)xmlNode.Attributes["ID"]).Value;
			hTKKCMCToKhaiInfo.ParentID = ((xmlNode.Attributes["ParentID"] == null) ? "" : ((XmlNode)xmlNode.Attributes["ParentID"]).Value);
			hTKKCMCToKhaiInfo.SignNode = ((xmlNode.Attributes["SignNode"] == null) ? "" : ((XmlNode)xmlNode.Attributes["SignNode"]).Value);
			hTKKCMCToKhaiInfo.NopTheoTinhCQT = ((xmlNode.Attributes["NopTheoTinhCQT"] == null) ? "" : ((XmlNode)xmlNode.Attributes["NopTheoTinhCQT"]).Value);
			XmlNodeList val2 = ((XmlNode)val).SelectNodes("/Menu");
			string empty = string.Empty;
			foreach (XmlNode item in val2)
			{
				XmlNode val3 = item;
				string value = ((XmlNode)val3.Attributes["Caption"]).Value;
				hTKKCMCToKhaiInfo.Caption = value;
				string moTaBieuMau = ((val3.Attributes["MoTaBieuMau"] == null) ? "" : ((XmlNode)val3.Attributes["MoTaBieuMau"]).Value);
				hTKKCMCToKhaiInfo.MoTaBieuMau = moTaBieuMau;
				string maFile = ((val3.Attributes["MaFile"] == null) ? "" : ((XmlNode)val3.Attributes["MaFile"]).Value);
				hTKKCMCToKhaiInfo.MaFile = maFile;
				string maToKhai = ((val3.Attributes["MaToKhai"] == null) ? "" : ((XmlNode)val3.Attributes["MaToKhai"]).Value);
				hTKKCMCToKhaiInfo.MaToKhai = maToKhai;
				hTKKCMCToKhaiInfo.UseDate = ((val3.Attributes["UseDate"] == null) ? "" : ((XmlNode)val3.Attributes["UseDate"]).Value);
				HTKKCMCToKhaiInfo toKhaiInfo = WinSettings.ToKhaiInfo;
				string kieuKHBS = (hTKKCMCToKhaiInfo.KieuKHBS = ((val3.Attributes["KieuKHBS"] == null) ? "" : ((XmlNode)val3.Attributes["KieuKHBS"]).Value));
				toKhaiInfo.KieuKHBS = kieuKHBS;
				if (!string.IsNullOrEmpty(periodInfo.VersionXML))
				{
					hTKKCMCToKhaiInfo.VersionXML = periodInfo.VersionXML;
				}
				string versionXML = ((val3.Attributes["VersionXML"] == null) ? "" : ((XmlNode)val3.Attributes["VersionXML"]).Value);
				hTKKCMCToKhaiInfo.VersionXML = versionXML;
				string versionKBSXML = ((val3.Attributes["VersionKHBS"] == null) ? "" : ((XmlNode)val3.Attributes["VersionKHBS"]).Value);
				hTKKCMCToKhaiInfo.VersionKBSXML = versionKBSXML;
				string tenToKhai = ((val3.Attributes["TenToKhai"] == null) ? "" : ((XmlNode)val3.Attributes["TenToKhai"]).Value);
				hTKKCMCToKhaiInfo.TenToKhai = tenToKhai;
				empty = ((val3.Attributes["XsdSchema"] == null) ? string.Empty : ((XmlNode)val3.Attributes["XsdSchema"]).Value);
				val2 = ((XmlNode)val).SelectNodes("/Menu/Validity");
				foreach (XmlNode item2 in val2)
				{
					XmlNode val4 = item2;
					hTKKCMCToKhaiInfo.FolderName = ((val4.Attributes["FolderName"] == null) ? "" : ((XmlNode)val4.Attributes["FolderName"]).Value);
					hTKKCMCToKhaiInfo.StartDate = ((val4.Attributes["StartDate"] == null) ? "" : ((XmlNode)val4.Attributes["StartDate"]).Value);
					hTKKCMCToKhaiInfo.EndDate = ((val4.Attributes["EndDate"] == null) ? "" : ((XmlNode)val4.Attributes["EndDate"]).Value);
				}
				if (hTKKCMCToKhaiInfo.MenuID != "100_1")
				{
					hTKKCMCToKhaiInfo.DataFileNameSaveExtend = GetDataFileNameExtend(periodInfo);
					hTKKCMCToKhaiInfo.FileNameExport = GetFileNameToSave(hTKKCMCToKhaiInfo, periodInfo, WinSettings.HTKKHeaderInfo);
				}
				hTKKCMCToKhaiInfo.InterfaceIni = HTKKConfig.INTERFACE_INI + "\\" + hTKKCMCToKhaiInfo.FolderName;
				hTKKCMCToKhaiInfo.InterfaceTemplateExcel = GetInterfaceTemplateExcel(hTKKCMCToKhaiInfo);
				hTKKCMCToKhaiInfo.InterfaceTemplateValidate = GetInterfaceTemplateValidate(hTKKCMCToKhaiInfo);
				hTKKCMCToKhaiInfo.InterfaceTemplateXml = GetInterfaceTemplateXml(hTKKCMCToKhaiInfo);
				hTKKCMCToKhaiInfo.InterfaceTemplateKBSXml = GetInterfaceTemplateKBSXml(hTKKCMCToKhaiInfo);
				hTKKCMCToKhaiInfo.ReportTemplates = GetReportTemplatesExcel(hTKKCMCToKhaiInfo);
				hTKKCMCToKhaiInfo.Version = HTKKConfig.VERSION;
				WinSettings.TTPeriodInfo.VersionXML = hTKKCMCToKhaiInfo.VersionXML;
				WinSettings.ToKhaiInfo.VersionXML = hTKKCMCToKhaiInfo.VersionXML;
				WinSettings.ToKhaiInfo.VersionKBSXML = hTKKCMCToKhaiInfo.VersionKBSXML;
				if (!periodInfo.LoaiToKhai.Equals("C") && "31;30;39;33;28;M2".IndexOf(WinSettings.MenuID) == -1 && hTKKCMCToKhaiInfo.KieuKHBS != "")
				{
					hTKKCMCToKhaiInfo.XsdFileName = (string.IsNullOrEmpty(empty) ? string.Empty : string.Format("{0}\\KHBS_{1}_{2}.xsd", HTKKConfig.XMLSCHEMA, empty, hTKKCMCToKhaiInfo.VersionKBSXML.Replace(".", "")));
				}
				else
				{
					hTKKCMCToKhaiInfo.XsdFileName = (string.IsNullOrEmpty(empty) ? string.Empty : string.Format("{0}\\{1}_{2}.xsd", HTKKConfig.XMLSCHEMA, empty, hTKKCMCToKhaiInfo.VersionXML.Replace(".", "")));
				}
				GetHTKKSheetInfos(hTKKCMCToKhaiInfo, getTemplate: false, periodInfo);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return hTKKCMCToKhaiInfo;
	}

	public static void GetHTKKSheetInfos(HTKKCMCToKhaiInfo toKhaiInfo, bool getTemplate, HTKKPeriodInfo periodInfo)
	{
		//IL_0004: Unknown result type (might be due to invalid IL or missing references)
		//IL_000a: Expected O, but got Unknown
		try
		{
			XmlNodeList val = null;
			XmlDocument val2 = new XmlDocument();
			val2.LoadXml(WinSettings.xmlMenuNode.OuterXml);
			val = val2.GetElementsByTagName("Sheet");
			for (int i = 0; i < val.Count; i++)
			{
				if (IsWorking(((XmlNode)val[i].Attributes["SheetName"]).Value))
				{
					if (getTemplate || WinSettings.ToKhaiInfo.TaiDuLieuRatLon != "1")
					{
						HTKKCMCSheetInfo hTKKSheetInfo = GetHTKKSheetInfo(toKhaiInfo, val, i, getTemplate, periodInfo);
						toKhaiInfo.SheetInfos.Add(hTKKSheetInfo);
					}
					else
					{
						HTKKCMCSheetInfo hTKKSheetInfo = null;
						hTKKSheetInfo = ((i == 0) ? GetHTKKSheetInfo(toKhaiInfo, val, i, getTemplate, periodInfo) : ((!(periodInfo.LoaiToKhai == "B") || !(WinSettings.ToKhaiInfo.KieuKHBS == "3") || i < val.Count - 2) ? GetHTKKSheetInfo(toKhaiInfo, val, i, getTemplate: true, periodInfo) : GetHTKKSheetInfo(toKhaiInfo, val, i, getTemplate, periodInfo)));
						toKhaiInfo.SheetInfos.Add(hTKKSheetInfo);
					}
				}
				else
				{
					toKhaiInfo.SheetInfos.Add(null);
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static bool IsWorking(string sheetName)
	{
		if (WinSettings.DicSheet.ContainsKey(sheetName))
		{
			return true;
		}
		return false;
	}

	public static void BindToKhaiInfoToTTinChung()
	{
		try
		{
			WinSettings.TTinChung.maDVu = HTKKConfig.MA_DICH_VU;
			WinSettings.TTinChung.tenDVu = HTKKConfig.TEN_DICH_VU;
			WinSettings.TTinChung.pbanDVu = HTKKConfig.VERSION;
			WinSettings.TTinChung.ttinNhaCCapDVu = HTKKConfig.THONG_TIN_NHA_CUNG_CAP;
			WinSettings.TTinChung.maTKhai = WinSettings.ToKhaiInfo.MaToKhai;
			WinSettings.TTinChung.tenTKhai = WinSettings.ToKhaiInfo.TenToKhai;
			WinSettings.TTinChung.moTaBMau = WinSettings.ToKhaiInfo.MoTaBieuMau;
			if (WinSettings.TTPeriodInfo.LoaiToKhai == "B")
			{
				WinSettings.TTinChung.pbanTKhaiXML = WinSettings.ToKhaiInfo.VersionKBSXML;
			}
			else
			{
				WinSettings.TTinChung.pbanTKhaiXML = WinSettings.ToKhaiInfo.VersionXML;
			}
			WinSettings.TTinChung.loaiTKhai = WinSettings.TTPeriodInfo.LoaiToKhai;
			WinSettings.TTinChung.soLan = WinSettings.TTPeriodInfo.SoLan;
			WinSettings.TTinChung.kieuKy = WinSettings.TTPeriodInfo.KieuKy;
			WinSettings.TTinChung.kyKKhai = WinSettings.TTPeriodInfo.kyKKhai;
			WinSettings.TTinChung.kyKKhaiTuNgay = WinSettings.TTPeriodInfo.TuNgay;
			WinSettings.TTinChung.kyKKhaiDenNgay = WinSettings.TTPeriodInfo.DenNgay;
			WinSettings.TTinChung.kyKKhaiTuThang = WinSettings.TTPeriodInfo.TuThang;
			WinSettings.TTinChung.kyKKhaiDenThang = WinSettings.TTPeriodInfo.DenThang;
			WinSettings.TTinChung.maCQTNoiNop = WinSettings.HTKKHeaderInfo.maCQTNoiNop;
			WinSettings.TTinChung.tenCQTNoiNop = WinSettings.HTKKHeaderInfo.tenCQTNoiNop;
			WinSettings.TTinChung.maCQTCapCuc = WinSettings.HTKKHeaderInfo.maCQTCapCuc;
			WinSettings.TTinChung.tenCQTCapCuc = WinSettings.HTKKHeaderInfo.tenCQTCapCuc;
			WinSettings.TTinChung.maDonViTinh = WinSettings.TTPeriodInfo.DonViTinhMa;
			WinSettings.TTinChung.tenDonViTinh = WinSettings.TTPeriodInfo.DonViTinhTen;
			WinSettings.TTinChung.ngay = WinSettings.TTPeriodInfo.Ngay;
			WinSettings.TTinChung.ngayLapTKhai = "";
			WinSettings.TTinChung.nguoiKy = WinSettings.HTKKHeaderInfo.nguoiKy;
			WinSettings.TTinChung.ngayKy = DateTime.Now.ToString("dd/MM/yyyy");
			WinSettings.TTinChung.nganhNgheKD = WinSettings.HTKKHeaderInfo.nganhNgheKD;
			WinSettings.TTinChung.mst = WinSettings.HTKKHeaderInfo.mst;
			WinSettings.TTinChung.tenNNT = WinSettings.HTKKHeaderInfo.tenNNT;
			WinSettings.TTinChung.dchiNNT = WinSettings.HTKKHeaderInfo.dchiNNT;
			WinSettings.TTinChung.phuongXa = "";
			WinSettings.TTinChung.maHuyenNNT = WinSettings.HTKKHeaderInfo.maHuyenNNT;
			WinSettings.TTinChung.tenXaPhuongNNT = WinSettings.HTKKHeaderInfo.tenXaPhuongNNT;
			WinSettings.TTinChung.maXaPhuongNNT = WinSettings.HTKKHeaderInfo.maXaPhuongNNT;
			WinSettings.TTinChung.maTinhNNT = WinSettings.HTKKHeaderInfo.maTinhNNT;
			WinSettings.TTinChung.tenTinhNNT = WinSettings.HTKKHeaderInfo.tenTinhNNT;
			WinSettings.TTinChung.dthoaiNNT = WinSettings.HTKKHeaderInfo.dthoaiNNT;
			WinSettings.TTinChung.faxNNT = WinSettings.HTKKHeaderInfo.faxNNT;
			WinSettings.TTinChung.emailNNT = WinSettings.HTKKHeaderInfo.emailNNT;
			WinSettings.TTinChung.mstDLyThue = WinSettings.HTKKHeaderInfo.mstDLyThue;
			WinSettings.TTinChung.tenDLyThue = WinSettings.HTKKHeaderInfo.tenDLyThue;
			WinSettings.TTinChung.dchiDLyThue = WinSettings.HTKKHeaderInfo.dchiDLyThue;
			WinSettings.TTinChung.maHuyenDLyThue = WinSettings.HTKKHeaderInfo.maHuyenDLyThue;
			WinSettings.TTinChung.tenHuyenDLyThue = WinSettings.HTKKHeaderInfo.tenHuyenDLyThue;
			WinSettings.TTinChung.tenXaPhuongDLyThue = WinSettings.HTKKHeaderInfo.tenXaPhuongDLyThue;
			WinSettings.TTinChung.maXaPhuongDlyThue = WinSettings.HTKKHeaderInfo.maXaPhuongDlyThue;
			List<string> list = new List<string>();
			list.Add("73");
			list.Add("56");
			List<string> list2 = list;
			if (list2.Contains(WinSettings.MenuID))
			{
				WinSettings.TTinChung.maHuyenDLyThue = WinSettings.HTKKHeaderInfo.maHuyenDLyThue;
				WinSettings.TTinChung.maTinhDLyThue = WinSettings.HTKKHeaderInfo.tenTinhDLyThue;
			}
			WinSettings.TTinChung.tenTinhDLyThue = WinSettings.HTKKHeaderInfo.tenTinhDLyThue;
			WinSettings.TTinChung.maTinhDLyThue = WinSettings.HTKKHeaderInfo.maTinhDLyThue;
			WinSettings.TTinChung.dthoaiDLyThue = WinSettings.HTKKHeaderInfo.dthoaiDLyThue;
			WinSettings.TTinChung.faxDLyThue = WinSettings.HTKKHeaderInfo.faxDLyThue;
			WinSettings.TTinChung.emailDLyThue = WinSettings.HTKKHeaderInfo.emailDLyThue;
			WinSettings.TTinChung.soHDongDLyThue = WinSettings.HTKKHeaderInfo.soHDongDLyThue;
			WinSettings.TTinChung.ngayKyHDDLyThue = WinSettings.HTKKHeaderInfo.ngayKyHDDLyThue;
			WinSettings.TTinChung.tenNVienDLyThue = WinSettings.HTKKHeaderInfo.tenNVienDLyThue;
			WinSettings.TTinChung.cchiHNghe = WinSettings.HTKKHeaderInfo.cchiHNghe;
			WinSettings.TTinChung.thang = WinSettings.TTPeriodInfo.Thang;
			WinSettings.TTinChung.quy = WinSettings.TTPeriodInfo.Quy;
			WinSettings.TTinChung.nam = WinSettings.TTPeriodInfo.Nam;
			WinSettings.TTinChung.SoTKNganHang = WinSettings.HTKKHeaderInfo.SoTKNganHang;
			WinSettings.TTinChung.TenDonViChuQuan = WinSettings.HTKKHeaderInfo.TenDonViChuQuan;
			WinSettings.TTinChung.MaSoThueDonViChuQuan = WinSettings.HTKKHeaderInfo.MaSoThueDonViChuQuan;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw;
		}
	}

	public static void BindTTChungToToKhaiInfo()
	{
		try
		{
			WinSettings.ToKhaiInfo.MaToKhai = WinSettings.TTinChung.maTKhai;
			WinSettings.ToKhaiInfo.TenToKhai = WinSettings.TTinChung.tenTKhai;
			WinSettings.ToKhaiInfo.MoTaBieuMau = WinSettings.TTinChung.moTaBMau;
			WinSettings.ToKhaiInfo.VersionXMLOld = WinSettings.TTinChung.pbanTKhaiXML;
			WinSettings.TTPeriodInfo.LoaiToKhai = WinSettings.TTinChung.loaiTKhai;
			WinSettings.TTPeriodInfo.SoLan = WinSettings.TTinChung.soLan;
			WinSettings.TTPeriodInfo.KieuKy = WinSettings.TTinChung.kieuKy;
			WinSettings.TTPeriodInfo.Nam = WinSettings.TTinChung.kyKKhai;
			WinSettings.TTPeriodInfo.TuNgay = WinSettings.TTinChung.kyKKhaiTuNgay;
			WinSettings.TTPeriodInfo.DenNgay = WinSettings.TTinChung.kyKKhaiDenNgay;
			WinSettings.TTPeriodInfo.TuThang = WinSettings.TTinChung.kyKKhaiTuThang;
			WinSettings.TTPeriodInfo.DenThang = WinSettings.TTinChung.kyKKhaiDenThang;
			WinSettings.HTKKHeaderInfo.maCQTNoiNop = WinSettings.TTinChung.maCQTNoiNop;
			WinSettings.HTKKHeaderInfo.tenCQTNoiNop = WinSettings.TTinChung.tenCQTNoiNop;
			WinSettings.HTKKHeaderInfo.maCQTCapCuc = WinSettings.TTinChung.maCQTCapCuc;
			WinSettings.HTKKHeaderInfo.tenCQTCapCuc = WinSettings.TTinChung.tenCQTCapCuc;
			WinSettings.TTinChung.ngayLapTKhai = "";
			WinSettings.HTKKHeaderInfo.SoTKNganHang = WinSettings.TTinChung.SoTKNganHang;
			WinSettings.HTKKHeaderInfo.TenDonViChuQuan = WinSettings.TTinChung.TenDonViChuQuan;
			WinSettings.HTKKHeaderInfo.MaSoThueDonViChuQuan = WinSettings.TTinChung.MaSoThueDonViChuQuan;
			WinSettings.TTinChung.nguoiKy = "";
			WinSettings.TTinChung.ngayKy = DateTime.Now.ToString("dd/MM/yyyy");
			WinSettings.HTKKHeaderInfo.nganhNgheKD = WinSettings.TTinChung.nganhNgheKD;
			WinSettings.HTKKHeaderInfo.mst = WinSettings.TTinChung.mst;
			WinSettings.HTKKHeaderInfo.tenNNT = WinSettings.TTinChung.tenNNT;
			WinSettings.HTKKHeaderInfo.dchiNNT = WinSettings.TTinChung.dchiNNT;
			WinSettings.TTinChung.phuongXa = "";
			WinSettings.TTinChung.maHuyenNNT = WinSettings.TTinChung.maHuyenNNT;
			WinSettings.HTKKHeaderInfo.tenXaPhuongNNT = WinSettings.TTinChung.tenXaPhuongNNT;
			WinSettings.HTKKHeaderInfo.maXaPhuongNNT = WinSettings.TTinChung.maXaPhuongNNT;
			WinSettings.HTKKHeaderInfo.tenHuyenNNT = WinSettings.TTinChung.tenHuyenNNT;
			WinSettings.TTinChung.maTinhNNT = WinSettings.TTinChung.maTinhNNT;
			WinSettings.HTKKHeaderInfo.tenTinhNNT = WinSettings.TTinChung.tenTinhNNT;
			WinSettings.HTKKHeaderInfo.dthoaiNNT = WinSettings.TTinChung.dthoaiNNT;
			WinSettings.HTKKHeaderInfo.faxNNT = WinSettings.TTinChung.faxNNT;
			WinSettings.HTKKHeaderInfo.emailNNT = WinSettings.TTinChung.emailNNT;
			WinSettings.HTKKHeaderInfo.mstDLyThue = WinSettings.TTinChung.mstDLyThue;
			WinSettings.HTKKHeaderInfo.tenDLyThue = WinSettings.TTinChung.tenDLyThue;
			WinSettings.HTKKHeaderInfo.dchiDLyThue = WinSettings.TTinChung.dchiDLyThue;
			WinSettings.TTinChung.maHuyenDLyThue = WinSettings.TTinChung.maHuyenDLyThue;
			WinSettings.HTKKHeaderInfo.tenHuyenDLyThue = WinSettings.TTinChung.tenHuyenDLyThue;
			WinSettings.HTKKHeaderInfo.tenXaPhuongDLyThue = WinSettings.TTinChung.tenXaPhuongDLyThue;
			WinSettings.HTKKHeaderInfo.maXaPhuongDlyThue = WinSettings.TTinChung.maXaPhuongDlyThue;
			WinSettings.TTinChung.maTinhDLyThue = WinSettings.TTinChung.maTinhDLyThue;
			WinSettings.HTKKHeaderInfo.tenTinhDLyThue = WinSettings.TTinChung.tenTinhDLyThue;
			WinSettings.HTKKHeaderInfo.dthoaiDLyThue = WinSettings.TTinChung.dthoaiDLyThue;
			WinSettings.HTKKHeaderInfo.faxDLyThue = WinSettings.TTinChung.faxDLyThue;
			WinSettings.HTKKHeaderInfo.emailDLyThue = WinSettings.TTinChung.emailDLyThue;
			WinSettings.HTKKHeaderInfo.soHDongDLyThue = WinSettings.TTinChung.soHDongDLyThue;
			WinSettings.HTKKHeaderInfo.ngayKyHDDLyThue = WinSettings.TTinChung.ngayKyHDDLyThue;
			WinSettings.HTKKHeaderInfo.tenNVienDLyThue = WinSettings.TTinChung.tenNVienDLyThue;
			WinSettings.HTKKHeaderInfo.cchiHNghe = WinSettings.TTinChung.cchiHNghe;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static string GetFileNameToSave(HTKKCMCToKhaiInfo ToKhaiInfo, HTKKPeriodInfo periodInfo, HTKKHeaderInfo HeaderInfo)
	{
		string empty = string.Empty;
		try
		{
			string tenCQTCapCuc = HeaderInfo.tenCQTCapCuc;
			string text = WinSettings.HTKKHeaderInfo.mst.PadRight(13, '0').Replace("-", "");
			string text2 = "L" + periodInfo.SoLan.PadLeft(2, '0');
			empty = tenCQTCapCuc.Substring(0, 3) + "-" + text + "-" + ToKhaiInfo.MaFile;
			if (periodInfo.KieuKy == "11")
			{
				empty = empty + "-Y" + periodInfo.Nam;
			}
			if ((periodInfo.TuThang != null) & (periodInfo.TuThang != ""))
			{
				empty = empty + "-" + periodInfo.TuThang.Replace("/", "");
			}
			if ((periodInfo.DenThang != null) & (periodInfo.DenThang != ""))
			{
				empty += periodInfo.DenThang.Replace("/", "");
			}
			if ((periodInfo.TuNgay != null) & (periodInfo.TuNgay != ""))
			{
				empty += periodInfo.TuNgay.Replace("/", "");
			}
			if ((periodInfo.DenNgay != null) & (periodInfo.DenNgay != ""))
			{
				empty += periodInfo.DenNgay.Replace("/", "");
			}
			return empty + "-" + text2;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return string.Empty;
		}
	}

	public static string GetDataFileNameExtend(HTKKPeriodInfo periodInfo)
	{
		string text = "";
		try
		{
			string text2 = "";
			string versionXML = periodInfo.VersionXML;
			string thang = periodInfo.Thang;
			string ngay = periodInfo.Ngay;
			string quy = periodInfo.Quy;
			string nam = periodInfo.Nam;
			string menuID = periodInfo.menuID;
			versionXML = ((!string.IsNullOrEmpty(versionXML)) ? versionXML.Replace(".", "") : string.Empty);
			string text3 = "~98~92~08~40~42~44~M8~M9~K2~K3~J9~J8~J7~E2~E3~E4~E5~E6~E7~E9~D7~D8~D9~R9~R6~S1~";
			string text4 = "~J3~Q7~Q5~Q6~";
			if (text3.IndexOf("~" + periodInfo.menuID + "~") > 0)
			{
				if (periodInfo.TieuMucHachToanVal.Equals("1"))
				{
					text2 = "DT";
				}
				else if (periodInfo.TieuMucHachToanVal.Equals("2"))
				{
					text2 = "KTN";
				}
				else if (periodInfo.TieuMucHachToanVal.Equals("3"))
				{
					text2 = "CD";
				}
			}
			if (periodInfo.KieuKy == "M")
			{
				if (text3.IndexOf("~" + periodInfo.menuID + "~") > 0)
				{
					if (periodInfo.menuID == "R9")
					{
						return ((!string.IsNullOrEmpty(periodInfo.DonViTinhMa)) ? ("_" + periodInfo.DonViTinhMa) : "") + text2 + "_L" + periodInfo.LanXuatBan + "_" + periodInfo.Thang + periodInfo.Nam + "_" + versionXML;
					}
					return ((!string.IsNullOrEmpty(periodInfo.DonViTinhMa)) ? ("_" + periodInfo.DonViTinhMa) : "") + text2 + "_" + periodInfo.Thang + periodInfo.Nam + "_" + versionXML;
				}
				return periodInfo.Thang + periodInfo.Nam + "_" + versionXML;
			}
			if (periodInfo.KieuKy == "Q")
			{
				if (!HTKKFormater.isNullString(periodInfo.TuThang) && !HTKKFormater.isNullString(periodInfo.DenThang))
				{
					if (periodInfo.menuID.Equals("42") || text3.IndexOf("~" + periodInfo.menuID + "~") > 0)
					{
						return text2 + "Q0" + periodInfo.Quy + periodInfo.Nam + "_" + periodInfo.TuThang.Replace("/", "") + "_" + periodInfo.DenThang.Replace("/", "") + "_" + versionXML;
					}
					return "Q0" + periodInfo.Quy + periodInfo.Nam + "_" + periodInfo.TuThang.Replace("/", "") + "_" + periodInfo.DenThang.Replace("/", "") + "_" + versionXML;
				}
				if (periodInfo.menuID.Equals("A1") || "X1".Equals(WinSettings.TTPeriodInfo.menuID))
				{
					return periodInfo.Nam + "_" + periodInfo.TuNgay.Replace("/", "") + "_" + periodInfo.DenNgay.Replace("/", "") + "_" + periodInfo.DonViTinhMa + "_" + versionXML;
				}
				if (periodInfo.menuID.Equals("A1") || periodInfo.menuID.Equals("A3") || "X1".Equals(WinSettings.TTPeriodInfo.menuID))
				{
					return "Q0" + periodInfo.Quy + periodInfo.Nam + "_" + versionXML;
				}
				if (periodInfo.menuID.Equals("10") || periodInfo.menuID.Equals("67"))
				{
					return "0" + periodInfo.Quy + periodInfo.Nam + "_" + versionXML;
				}
				return "Q0" + periodInfo.Quy + periodInfo.Nam + "_" + versionXML;
			}
			if (periodInfo.KieuKy == "Y")
			{
				if (periodInfo.menuID.Equals("93") || periodInfo.menuID.Equals("89") || periodInfo.menuID.Equals("44") || periodInfo.menuID.Equals("42") || text3.IndexOf("~" + periodInfo.menuID + "~") > 0)
				{
					return ((!string.IsNullOrEmpty(periodInfo.DonViTinhMa)) ? ("_" + periodInfo.DonViTinhMa) : "") + text2 + "_" + periodInfo.Nam + "_" + periodInfo.TuThang.Replace("/", "") + "_" + periodInfo.DenThang.Replace("/", "") + "_" + versionXML;
				}
				if (periodInfo.menuID.Equals("95"))
				{
					return "L" + periodInfo.SoLan + "_" + periodInfo.Nam + "_" + versionXML;
				}
				if (!HTKKFormater.isNullString(periodInfo.TuThang) && !HTKKFormater.isNullString(periodInfo.DenThang))
				{
					return periodInfo.Nam + "_" + periodInfo.TuThang.Replace("/", "") + "_" + periodInfo.DenThang.Replace("/", "") + "_" + versionXML;
				}
				if (periodInfo.menuID.Equals("78") || WinSettings.MenuID.Equals("U9"))
				{
					return periodInfo.Nam + "_" + periodInfo.TuNgay.Replace("/", "") + "_" + periodInfo.DenNgay.Replace("/", "") + "_" + periodInfo.DonViTinhMa + "_" + versionXML;
				}
				if (periodInfo.menuID.Equals("53") || periodInfo.menuID.Equals("83"))
				{
					return periodInfo.Nam + "_" + periodInfo.TuNgay.Replace("/", "") + "_" + periodInfo.DenNgay.Replace("/", "") + "_" + versionXML;
				}
				if (periodInfo.menuID.Equals("E8"))
				{
					return periodInfo.TieuMucHachToanVal + "_" + periodInfo.Nam + "_" + periodInfo.TuNgay.Replace("/", "") + "_" + periodInfo.DenNgay.Replace("/", "") + "_" + versionXML;
				}
				if (periodInfo.menuID.Equals("K8") || periodInfo.menuID.Equals("03") || periodInfo.menuID.Equals("N8") || text4.IndexOf("~" + periodInfo.menuID + "~") >= 0)
				{
					return periodInfo.Nam + "_" + periodInfo.TuNgay.Replace("/", "") + "_" + periodInfo.DenNgay.Replace("/", "") + "_" + versionXML;
				}
				return periodInfo.Nam + "_" + versionXML;
			}
			if (periodInfo.KieuKy == "D")
			{
				if (!HTKKFormater.isNullString(periodInfo.LanXuatBan) && !HTKKFormater.isNullString(periodInfo.TieuMucHachToanVal))
				{
					return text2 + "_L" + periodInfo.LanXuatBan + "_" + periodInfo.Ngay + periodInfo.Thang + periodInfo.Nam + "_" + versionXML;
				}
				if (HTKKFormater.isNullString(periodInfo.LanXuatBan) && !HTKKFormater.isNullString(periodInfo.TieuMucHachToanVal))
				{
					return ((!string.IsNullOrEmpty(periodInfo.DonViTinhMa)) ? ("_" + periodInfo.DonViTinhMa) : "") + text2 + "_" + periodInfo.Ngay + periodInfo.Thang + periodInfo.Nam + "_" + versionXML;
				}
				if (menuID == "47" || menuID == "K7")
				{
					return periodInfo.Ngay + periodInfo.Thang + periodInfo.Nam + "_" + (periodInfo.CTDC ? "CTDC" : "LNDC") + "_" + versionXML;
				}
				return periodInfo.Ngay + periodInfo.Thang + periodInfo.Nam + versionXML;
			}
			if (periodInfo.KieuKy == "D_D")
			{
				return periodInfo.TuNgay.Replace("/", "") + "_" + periodInfo.DenNgay.Replace("/", "") + "_" + versionXML;
			}
			return "_" + versionXML;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return string.Empty;
		}
	}

	public static string GetDataFileNameExtend(string strDataFileNameSaveExtend, string strVersionXML, string strDataFile)
	{
		string empty = string.Empty;
		try
		{
			if (strDataFile == "HEADER")
			{
				return $"{HTKKConfig.DATAFILES}\\{WinSettings.HTKKHeaderInfo.mst}\\{strDataFile}.xml";
			}
			return HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst + "\\" + HTKKConfig.VERSION + "_" + strVersionXML + "_" + strDataFile + "_" + strDataFileNameSaveExtend + ".xml";
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return string.Empty;
		}
	}

	public static List<string> GetQuyByThang(string strThang, string strNam)
	{
		List<string> list = new List<string>();
		switch (strThang)
		{
		case "T_0001":
			list.Add("Q_0001" + strNam);
			break;
		case "T_0002":
			list.Add("Q_0001" + strNam);
			break;
		case "T_0003":
			list.Add("Q_0001" + strNam);
			break;
		case "T_0004":
			list.Add("Q_0002" + strNam);
			break;
		case "T_0005":
			list.Add("Q_0002" + strNam);
			break;
		case "T_0006":
			list.Add("Q_0002" + strNam);
			break;
		case "T_0007":
			list.Add("Q_0003" + strNam);
			break;
		case "T_0008":
			list.Add("Q_0003" + strNam);
			break;
		case "T_0009":
			list.Add("Q_0003" + strNam);
			break;
		case "T_0010":
			list.Add("Q_0004" + strNam);
			break;
		case "T_0011":
			list.Add("Q_0004" + strNam);
			break;
		case "T_0012":
			list.Add("Q_0004" + strNam);
			break;
		default:
			return list;
		}
		return list;
	}

	public static string GetInterfaceTemplateExcel(HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		string empty = string.Empty;
		try
		{
			if (ToKhaiInfo.FolderName == "HEADER")
			{
				return HTKKConfig.INTERFACE_TEMPLATES + "\\Excel\\" + ToKhaiInfo.FolderName + ".xls";
			}
			return HTKKConfig.INTERFACE_TEMPLATES + "\\Excel\\" + ToKhaiInfo.FolderName + "_" + ToKhaiInfo.VersionXML.Replace(".", "") + ".xls";
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return string.Empty;
		}
	}

	public static string GetInterfaceTemplateKBSXml(HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		string empty = string.Empty;
		try
		{
			return HTKKConfig.INTERFACE_TEMPLATES + "\\Xml\\";
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return string.Empty;
		}
	}

	public static string GetReportTemplatesExcel(HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		string empty = string.Empty;
		try
		{
			if (ToKhaiInfo.FolderName == "HEADER")
			{
				return HTKKConfig.REPORT_TEMPLATES + "\\Excel\\" + ToKhaiInfo.FolderName + ".xls";
			}
			return HTKKConfig.REPORT_TEMPLATES + "\\Excel\\" + ToKhaiInfo.FolderName + "_" + ToKhaiInfo.VersionXML.Replace(".", "") + ".xls";
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return string.Empty;
		}
	}

	public static string GetInterfaceTemplateValidate(HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		string empty = string.Empty;
		try
		{
			if (ToKhaiInfo.FolderName == "HEADER")
			{
				return HTKKConfig.INTERFACE_TEMPLATES + "\\Excel\\" + ToKhaiInfo.FolderName + ".xls";
			}
			return HTKKConfig.INTERFACE_TEMPLATES + "\\Excel\\" + ToKhaiInfo.FolderName + "_" + ToKhaiInfo.VersionXML.Replace(".", "") + ".xls";
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return string.Empty;
		}
	}

	public static string GetInterfaceTemplateXml(HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		string empty = string.Empty;
		try
		{
			if (ToKhaiInfo.FolderName == "HEADER")
			{
				return HTKKConfig.INTERFACE_TEMPLATES + "\\Xml\\";
			}
			return HTKKConfig.INTERFACE_TEMPLATES + "\\Xml\\";
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return string.Empty;
		}
	}

	public static string GetHeaderStringBarcode(string sttMaVach, string tongsomavach, HTKKPeriodInfo periodInfo)
	{
		string empty = string.Empty;
		try
		{
			empty = ((!periodInfo.LoaiToKhai.Equals("C")) ? (empty + "bs") : (empty + "aa"));
			empty += WinSettings.ToKhaiInfo.Version;
			empty += WinSettings.ToKhaiInfo.MaToKhai;
			empty += WinSettings.TTinChung.mst.PadRight(13, ' ');
			empty += WinSettings.TTinChung.kyKKhai.PadRight(6, '0');
			empty += HTKKFile.GetSession(6);
			empty += sttMaVach.PadLeft(3, '0');
			empty += tongsomavach.PadLeft(3, '0');
			empty += WinSettings.HTKKHeaderInfo.NgayBatDauNamTaiChinh;
			return empty + WinSettings.ToKhaiInfo.StartDate;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return string.Empty;
		}
	}

	public static string getfileNameLanTruoc_KHBS(string fullNameDataFileToKhaiChinh)
	{
		string result = string.Empty;
		try
		{
			string path = fullNameDataFileToKhaiChinh.Substring(0, fullNameDataFileToKhaiChinh.LastIndexOf('\\'));
			string[] files = Directory.GetFiles(path, "*.xml");
			string text = fullNameDataFileToKhaiChinh.Substring(0, fullNameDataFileToKhaiChinh.LastIndexOf('_'));
			text = text.Substring(0, text.LastIndexOf('_'));
			string s = fullNameDataFileToKhaiChinh.Replace(".xml", "").Substring(fullNameDataFileToKhaiChinh.LastIndexOf('_') + 1);
			for (int i = 0; i < files.Length; i++)
			{
				if (files[i].IndexOf(text) >= 0)
				{
					string s2 = files[i].Replace(".xml", "").Substring(files[i].LastIndexOf('_') + 1);
					if (int.Parse(s) == int.Parse(s2) + 1)
					{
						result = files[i];
						break;
					}
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			result = string.Empty;
		}
		return result;
	}

	public static string getfileNameKHBS(string fullNameDataFileToKhaiChinh)
	{
		string result = string.Empty;
		try
		{
			string path = fullNameDataFileToKhaiChinh.Substring(0, fullNameDataFileToKhaiChinh.LastIndexOf('\\'));
			string[] files = Directory.GetFiles(path, "*.xml");
			string text = fullNameDataFileToKhaiChinh.Substring(0, fullNameDataFileToKhaiChinh.LastIndexOf('_'));
			text = text.Substring(0, text.LastIndexOf('_'));
			string s = fullNameDataFileToKhaiChinh.Replace(".xml", "").Substring(fullNameDataFileToKhaiChinh.LastIndexOf('_') + 1);
			for (int i = 0; i < files.Length; i++)
			{
				if (files[i].IndexOf(text) >= 0)
				{
					string s2 = files[i].Replace(".xml", "").Substring(files[i].LastIndexOf('_') + 1);
					if (int.Parse(s) == int.Parse(s2))
					{
						result = Path.GetFileName(files[i]).Replace(".xml", "");
						break;
					}
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			result = string.Empty;
		}
		return result;
	}

	public static string GetSoLanKeKhai(string fullNameDataFileToKhaiChinh, string strFileNameExtend)
	{
		int num = 0;
		try
		{
			string text = "";
			string empty = string.Empty;
			string text2 = fullNameDataFileToKhaiChinh.Substring(fullNameDataFileToKhaiChinh.LastIndexOf('\\') + 1);
			string text3 = text2.Substring(0, text2.LastIndexOf('_'));
			if (text3.Contains("bs_"))
			{
				text3 = text3.Replace("bs_", "");
			}
			string path = fullNameDataFileToKhaiChinh.Substring(0, fullNameDataFileToKhaiChinh.LastIndexOf('\\'));
			string[] files = Directory.GetFiles(path, "*.xml");
			for (int i = 0; i < files.Length; i++)
			{
				text = files[i].Substring(files[i].LastIndexOf('\\') + 1);
				if (text.IndexOf(text3) > 0 && text.Substring(0, 2).ToUpper().Equals("BS"))
				{
					string text4 = text.Substring(0, text.IndexOf('_'));
					text4 = text4.ToUpper();
					text4 = text4.Replace("BS", "");
					int num2 = ConvertHelpers.ConvertToInt(text4);
					if (num < num2)
					{
						num = num2;
					}
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return num.ToString();
	}

	public static void GetMCT(string pathFile, string IDTokhai)
	{
		//IL_0001: Unknown result type (might be due to invalid IL or missing references)
		//IL_0007: Expected O, but got Unknown
		//IL_009a: Unknown result type (might be due to invalid IL or missing references)
		//IL_00a0: Expected O, but got Unknown
		XmlDocument val = new XmlDocument();
		val.Load(pathFile);
		XmlNodeList val2 = ((XmlNode)val.DocumentElement).SelectNodes("/Root/Map");
		for (int i = 0; i < val2.Count; i++)
		{
			WinSettings.KHBS = new List<HTKKCMCKHBS>();
			if (!(IDTokhai == ((XmlNode)val2[i].Attributes["ID"]).Value) || val2[i].Attributes["type"] != null)
			{
				continue;
			}
			{
				foreach (XmlNode item in val2[i])
				{
					XmlNode val3 = item;
					HTKKCMCKHBS hTKKCMCKHBS = new HTKKCMCKHBS();
					string value = ((XmlNode)val3.Attributes["CellID"]).Value;
					string[] array = value.Split(new string[1] { "_" }, StringSplitOptions.None);
					hTKKCMCKHBS.Cell.ColIndex = HTKKFile.ColumnLettersToIndex(array[0]);
					hTKKCMCKHBS.Cell.RowIndex = Convert.ToInt32(array[1]);
					hTKKCMCKHBS.MaChiTieu = ((XmlNode)val3.Attributes["MCT"]).Value;
					hTKKCMCKHBS.TenChiTieu = ((XmlNode)val3.Attributes["Caption"]).Value;
					hTKKCMCKHBS.DieuChinhTang = ((XmlNode)val3.Attributes["DieuChinhTang"]).Value;
					WinSettings.KHBS.Add(hTKKCMCKHBS);
				}
				break;
			}
		}
	}

	public static XmlNode GetMCT01(string pathFile, string IDTokhai)
	{
		//IL_000e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0014: Expected O, but got Unknown
		XmlDocument val = new XmlDocument();
		val.Load(pathFile);
		XmlNodeList source = ((XmlNode)val.DocumentElement).SelectNodes("/Root/Map");
		XmlNode val2 = ((IEnumerable)source).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode x) => ((XmlNode)x.Attributes["ID"]).Value == IDTokhai));
		return (val2 != null) ? val2 : null;
	}

	public static void OpenToKhai(Form frmForm, SheetView sView, bool isCoLayThemPhuLuc, HTKKPeriodInfo periodInfo)
	{
		//IL_0036: Unknown result type (might be due to invalid IL or missing references)
		//IL_003d: Expected O, but got Unknown
		//IL_0f07: Unknown result type (might be due to invalid IL or missing references)
		try
		{
			WinSettings.ToKhaiInfo.TaiDuLieuRatLon = WinDoWorkEvent.checkTaiDuLieuLon(WinSettings.MenuID, "TaiDuLieuRatLon");
			WinSettings.ToKhaiInfo.TaiDuLieuLon = WinDoWorkEvent.checkTaiDuLieuLon(WinSettings.MenuID, "TaiDuLieuLon");
			XmlDocument val = new XmlDocument();
			val.LoadXml(WinSettings.xmlMenuNode.OuterXml);
			XmlNodeList elementsByTagName = val.GetElementsByTagName("Sheet");
			XmlNode val2 = elementsByTagName[0];
			string text = ((val2.Attributes["DataFile"] == null) ? "" : ((XmlNode)val2.Attributes["DataFile"]).Value);
			string versionXML = ((periodInfo.VersionXML == null || !(periodInfo.VersionXML != "")) ? ((val2.ParentNode.ParentNode.Attributes["VersionXML"] == null) ? "" : ((XmlNode)val2.ParentNode.ParentNode.Attributes["VersionXML"]).Value) : periodInfo.VersionXML);
			string text2 = ((val2.ParentNode.ParentNode.Attributes["VersionKHBS"] == null) ? "" : ((XmlNode)val2.ParentNode.ParentNode.Attributes["VersionKHBS"]).Value);
			string key = ((val2.Attributes["SheetName"] == null) ? "" : ((XmlNode)val2.Attributes["SheetName"]).Value);
			val2 = elementsByTagName[elementsByTagName.Count - 1];
			string text3 = ((val2.Attributes["DataFile"] == null) ? "" : ((XmlNode)val2.Attributes["DataFile"]).Value);
			string key2 = ((val2.Attributes["SheetName"] == null) ? "" : ((XmlNode)val2.Attributes["SheetName"]).Value);
			string dataFileNameExtend = GetDataFileNameExtend(periodInfo);
			string fileNameBySheet = WinDoWorkEvent.GetFileNameBySheet(WinSettings.TTPeriodInfo, text, HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst + "\\", versionXML);
			if ("N7~M1~M2~C6~C7~T5~H4~O1~O2~O3~O4~O5~O6~Q4~P7~U1~U8~X5~E9~R9~Y3~Y4~Z1~".IndexOf(WinSettings.MenuID) >= 0)
			{
				string fileNameLoadKhbsNew = WinDoWorkEvent.GetFileNameLoadKhbsNew(WinSettings.TTPeriodInfo, fileNameBySheet, HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst + "\\", isLoadSavedFile: true);
			}
			else
			{
				string fileNameLoadKhbsNew = WinDoWorkEvent.GetFileNameLoadKhbs(WinSettings.TTPeriodInfo, text, HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst + "\\", isLoadSavedFile: true);
			}
			if (periodInfo.LoaiToKhai == "B" && ((XmlNode)WinSettings.xmlMenuNode.Attributes["KieuKHBS"]).Value != "3")
			{
				if (WinSettings.xmlMenuNode.Attributes["KieuKHBS"] != null)
				{
					if (!((XmlNode)WinSettings.xmlMenuNode.Attributes["KieuKHBS"]).Value.Equals(""))
					{
						if ("30,I5,K9,K1,M7,K6,K8,73,55,71,31,P2,X6".Contains(WinSettings.MenuID))
						{
							List<string> phuLucThemMoi = GetPhuLucThemMoi(sView, elementsByTagName);
							if (phuLucThemMoi.Count > 0)
							{
								WinSettings.DicSheet = new Dictionary<string, int>();
								WinSettings.DicSheet.Add(key, 2);
								foreach (string item in phuLucThemMoi)
								{
									WinSettings.DicSheet.Add(item, 2);
								}
							}
							WinSettings.DicSheet.Add(key2, 2);
						}
						else
						{
							WinSettings.DicSheet = new Dictionary<string, int>();
							WinSettings.DicSheet.Add(key, 2);
							WinSettings.DicSheet.Add(key2, 2);
						}
					}
					else if (isCoLayThemPhuLuc)
					{
						List<string> phuLucThemMoi = GetPhuLucThemMoi(sView, elementsByTagName);
						foreach (string item2 in phuLucThemMoi)
						{
							WinSettings.DicSheet.Add(item2, 2);
						}
					}
				}
				GetMCT("MapMCT.xml", text);
				string text4 = "";
				text4 = (("N7~M1~M2~C6~C7~T5~H4~O1~O2~O3~O4~O5~O6~Q4~P7~U1~U8~X5~E9~R9~Y3~Y4~Z1~".IndexOf(WinSettings.MenuID) < 0) ? WinDoWorkEvent.GetFileNameLoadKhbs(WinSettings.TTPeriodInfo, text, HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst + "\\", isLoadSavedFile: true) : WinDoWorkEvent.GetFileNameLoadKhbsNew(WinSettings.TTPeriodInfo, fileNameBySheet, HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst + "\\", isLoadSavedFile: true));
				string text5 = "";
				text5 = (("N7~M1~M2~C6~C7~T5~H4~O1~O2~O3~O4~O5~O6~Q4~P7~U1~U8~X5~E9~R9~Y3~Y4~Z1~".IndexOf(WinSettings.MenuID) < 0) ? WinDoWorkEvent.GetFileNameLoadKhbs(WinSettings.TTPeriodInfo, text, HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst + "\\", isLoadSavedFile: false) : WinDoWorkEvent.GetFileNameLoadKhbsNew(WinSettings.TTPeriodInfo, fileNameBySheet, HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst + "\\", isLoadSavedFile: false));
				short num = 0;
				HTKKCMCSheetInfo hTKKCMCSheetInfo = new HTKKCMCSheetInfo();
				hTKKCMCSheetInfo.SheetID = ((elementsByTagName[(int)num].Attributes["ID"] == null) ? "" : ((XmlNode)elementsByTagName[(int)num].Attributes["ID"]).Value);
				hTKKCMCSheetInfo.SheetName = ((elementsByTagName[(int)num].Attributes["SheetName"] == null) ? "" : ((XmlNode)elementsByTagName[(int)num].Attributes["SheetName"]).Value);
				hTKKCMCSheetInfo.Caption = ((elementsByTagName[(int)num].Attributes["Caption"] == null) ? "" : ((XmlNode)elementsByTagName[(int)num].Attributes["Caption"]).Value);
				hTKKCMCSheetInfo.DataFile = text;
				hTKKCMCSheetInfo.Active = ((elementsByTagName[(int)num].Attributes["Active"] == null) ? "" : ((XmlNode)elementsByTagName[(int)num].Attributes["Active"]).Value);
				hTKKCMCSheetInfo.SheetIndex = num;
				hTKKCMCSheetInfo.ShowSearchButton = ((elementsByTagName[(int)num].Attributes["ShowSearchButton"] == null) ? "" : ((XmlNode)elementsByTagName[(int)num].Attributes["ShowSearchButton"]).Value);
				hTKKCMCSheetInfo.ShowTaiBangKeButton = ((elementsByTagName[(int)num].Attributes["ShowTaiBangKeButton"] == null) ? "" : ((XmlNode)elementsByTagName[(int)num].Attributes["ShowTaiBangKeButton"]).Value);
				hTKKCMCSheetInfo.strPrintBarcode = ((elementsByTagName[(int)num].Attributes["PrintBarCode"] == null) ? "" : ((XmlNode)elementsByTagName[(int)num].Attributes["PrintBarCode"]).Value);
				hTKKCMCSheetInfo.XMLnode = ((elementsByTagName[(int)num].Attributes["XMLnode"] == null) ? "" : ((XmlNode)elementsByTagName[(int)num].Attributes["XMLnode"]).Value);
				hTKKCMCSheetInfo.CoDieuChinhSoLieu = ((elementsByTagName[(int)num].Attributes["CoDieuChinhSoLieu"] == null) ? (-1) : int.Parse(((XmlNode)elementsByTagName[(int)num].Attributes["CoDieuChinhSoLieu"]).Value));
				hTKKCMCSheetInfo.RowTTDLT = ((elementsByTagName[(int)num].Attributes["RowTTDLT"] == null) ? "" : ((XmlNode)elementsByTagName[(int)num].Attributes["RowTTDLT"]).Value);
				hTKKCMCSheetInfo.TaiDuLieuLon = ((elementsByTagName[(int)num].Attributes["TaiDuLieuLon"] == null) ? "" : ((XmlNode)elementsByTagName[(int)num].Attributes["TaiDuLieuLon"]).Value);
				hTKKCMCSheetInfo.TaiDuLieuRatLon = ((elementsByTagName[(int)num].Attributes["TaiDuLieuRatLon"] == null) ? "" : ((XmlNode)elementsByTagName[(int)num].Attributes["TaiDuLieuRatLon"]).Value);
				WinSettings.KHBS_LanTruoc = hTKKCMCSheetInfo;
				if (periodInfo.menuID == "P7")
				{
					if (!File.Exists(text4))
					{
						text4 = GeneralFunctions.GetXMLFileName(HTKKConfig.INTERFACE_TEMPLATES + "\\xml", WinSettings.KHBS_LanTruoc.DataFile, WinSettings.ToKhaiInfo.VersionXML, "xml");
					}
					if (!File.Exists(text5))
					{
						text5 = GeneralFunctions.GetXMLFileName(HTKKConfig.INTERFACE_TEMPLATES + "\\xml", WinSettings.KHBS_LanTruoc.DataFile, WinSettings.ToKhaiInfo.VersionXML, "xml");
					}
				}
				WinSettings.KHBS_LanTruoc.GetDetailSheetInfo(text4);
				WinSettings.ToKhaiInfo_BoSungLanTruoc = new HTKKCMCToKhaiInfo();
				WinSettings.ToKhaiInfo_BoSungLanTruoc.SheetInfos.Add(new HTKKCMCSheetInfo());
				WinSettings.ToKhaiInfo_BoSungLanTruoc.SheetInfos[0].GetDetailSheetInfo(text5);
			}
			else if (periodInfo.LoaiToKhai == "B" && ((XmlNode)WinSettings.xmlMenuNode.Attributes["KieuKHBS"]).Value == "3")
			{
				List<string> phuLucThemMoi = GetPhuLucThemMoi(sView, elementsByTagName);
				if (isCoLayThemPhuLuc && phuLucThemMoi.Count > 0)
				{
					foreach (string item3 in phuLucThemMoi)
					{
						WinSettings.DicSheet.Add(item3, 2);
					}
				}
				WinSettings.DicSheet.Add(key2, 2);
				val2 = elementsByTagName[elementsByTagName.Count - 2];
				text3 = ((val2.Attributes["DataFile"] == null) ? "" : ((XmlNode)val2.Attributes["DataFile"]).Value);
				key2 = ((val2.Attributes["SheetName"] == null) ? "" : ((XmlNode)val2.Attributes["SheetName"]).Value);
				WinSettings.DicSheet.Add(key2, 2);
			}
			else if (isCoLayThemPhuLuc)
			{
				List<string> phuLucThemMoi = GetPhuLucThemMoi(sView, elementsByTagName);
				foreach (string item4 in phuLucThemMoi)
				{
					WinSettings.DicSheet.Add(item4, 2);
				}
			}
			WinSettings.ToKhaiInfo = GetToKhaiInfo(WinSettings.xmlMenuNode, periodInfo);
			BindToKhaiInfoToTTinChung();
			if (periodInfo.menuID == "G5" || periodInfo.menuID == "G7" || periodInfo.menuID == "G8" || periodInfo.menuID == "G6" || periodInfo.menuID == "L1")
			{
				if (string.IsNullOrEmpty(WinSettings.TTinChung.kyKKhaiTuNgay) || string.IsNullOrEmpty(WinSettings.TTinChung.kyKKhaiDenNgay))
				{
					DateTime ngayDauKyKeKhai = GeneralFunctions.GetNgayDauKyKeKhai(WinSettings.TTinChung, WinSettings.TTPeriodInfo);
					DateTime ngayCuoiKyKeKhai = GeneralFunctions.GetNgayCuoiKyKeKhai(WinSettings.TTinChung, WinSettings.TTPeriodInfo);
					string kyKKhaiTuNgay = ((!string.IsNullOrEmpty(WinSettings.TTinChung.kyKKhaiTuThang)) ? ("01/" + SecurityElement.Escape(WinSettings.TTinChung.kyKKhaiTuThang)) : SecurityElement.Escape(ngayDauKyKeKhai.ToString("dd/MM/yyyy")));
					DateTime dateTime = (string.IsNullOrEmpty(WinSettings.TTinChung.kyKKhaiDenThang) ? DateTime.MinValue : DateTime.ParseExact("01/" + WinSettings.TTinChung.kyKKhaiDenThang, "dd/MM/yyyy", CultureInfo.CurrentCulture));
					DateTime dateTime2 = (string.IsNullOrEmpty(WinSettings.TTinChung.kyKKhaiDenThang) ? DateTime.MinValue : dateTime.AddMonths(1).AddDays(-1.0));
					string kyKKhaiDenNgay = ((!string.IsNullOrEmpty(WinSettings.TTinChung.kyKKhaiDenThang)) ? dateTime2.ToString("dd/MM/yyyy") : SecurityElement.Escape(ngayCuoiKyKeKhai.ToString("dd/MM/yyyy")));
					WinSettings.TTinChung.kyKKhaiTuNgay = kyKKhaiTuNgay;
					WinSettings.TTinChung.kyKKhaiDenNgay = kyKKhaiDenNgay;
				}
				else if (!string.IsNullOrEmpty(WinSettings.TTinChung.kyKKhaiTuThang) || !string.IsNullOrEmpty(WinSettings.TTinChung.kyKKhaiDenThang))
				{
					DateTime ngayDauKyKeKhai = GeneralFunctions.GetNgayDauKyKeKhai(WinSettings.TTinChung, WinSettings.TTPeriodInfo);
					DateTime ngayCuoiKyKeKhai = GeneralFunctions.GetNgayCuoiKyKeKhai(WinSettings.TTinChung, WinSettings.TTPeriodInfo);
					string kyKKhaiTuNgay = ((!string.IsNullOrEmpty(WinSettings.TTinChung.kyKKhaiTuThang)) ? ("01/" + SecurityElement.Escape(WinSettings.TTinChung.kyKKhaiTuThang)) : SecurityElement.Escape(ngayDauKyKeKhai.ToString("dd/MM/yyyy")));
					DateTime dateTime2 = DateTime.ParseExact("01/" + WinSettings.TTinChung.kyKKhaiDenThang, "dd/MM/yyyy", CultureInfo.CurrentCulture).AddMonths(1).AddDays(-1.0);
					string kyKKhaiDenNgay = ((!string.IsNullOrEmpty(WinSettings.TTinChung.kyKKhaiDenThang)) ? dateTime2.ToString("dd/MM/yyyy") : SecurityElement.Escape(ngayCuoiKyKeKhai.ToString("dd/MM/yyyy")));
					WinSettings.TTinChung.kyKKhaiTuNgay = kyKKhaiTuNgay;
					WinSettings.TTinChung.kyKKhaiDenNgay = kyKKhaiDenNgay;
				}
			}
			frmWinToKhai frmWinToKhai2 = new frmWinToKhai(WinSettings.xmlMenuNode);
			if (frmForm != null)
			{
				if (frmForm is frmHTKKPeriod)
				{
					((Control)frmForm).Hide();
					((Control)WinSettings.myHTKKSystems.panelPeriod).Size = new Size(1, 1);
					((Control)WinSettings.myHTKKSystems.panelPeriod).Location = new Point(((Control)WinSettings.myHTKKSystems).Width, ((Control)WinSettings.myHTKKSystems).Height);
					((Control)WinSettings.myHTKKSystems.panelPeriod).Visible = false;
					((Control)WinSettings.myHTKKSystems).Hide();
				}
				if (frmForm is frmTraCuuToKhai)
				{
					((Control)WinSettings.myHTKKSystems).Visible = false;
				}
			}
			((Form)frmWinToKhai2).ShowDialog();
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static HTKKCMCSheetInfo ReadSheet_LanTruoc(int sheetIndex)
	{
		//IL_0001: Unknown result type (might be due to invalid IL or missing references)
		//IL_0007: Expected O, but got Unknown
		XmlDocument val = new XmlDocument();
		val.LoadXml(WinSettings.xmlMenuNode.OuterXml);
		XmlNodeList elementsByTagName = val.GetElementsByTagName("Sheet");
		XmlNode val2 = elementsByTagName[sheetIndex];
		string text = ((val2.Attributes["DataFile"] == null) ? "" : ((XmlNode)val2.Attributes["DataFile"]).Value);
		string text2 = "";
		if ("N7~M1~M2~C6~C7~T5~H4~O1~O2~O3~O4~O5~O6~Q4~P7~U1~U8~X5~E9~R9~Y3~Y4~Z1~".IndexOf(WinSettings.MenuID) >= 0)
		{
			string text3 = "";
			string sDataFile = WinDoWorkEvent.GetFileNameBySheetNew(versionXML: (WinSettings.TTPeriodInfo.VersionXML == null || !(WinSettings.TTPeriodInfo.VersionXML != "")) ? ((val2.ParentNode.ParentNode.Attributes["VersionXML"] == null) ? "" : ((XmlNode)val2.ParentNode.ParentNode.Attributes["VersionXML"]).Value) : WinSettings.TTPeriodInfo.VersionXML, periodInfor: WinSettings.TTPeriodInfo, dataFile: text, sFolder: HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst + "\\");
			text2 = WinDoWorkEvent.GetFileNameLoadKhbsNew(WinSettings.TTPeriodInfo, sDataFile, HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst + "\\", isLoadSavedFile: false);
		}
		else
		{
			text2 = WinDoWorkEvent.GetFileNameLoadKhbs(WinSettings.TTPeriodInfo, text, HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst + "\\", isLoadSavedFile: false);
		}
		if (!File.Exists(text2))
		{
			return null;
		}
		HTKKCMCSheetInfo hTKKCMCSheetInfo = new HTKKCMCSheetInfo();
		hTKKCMCSheetInfo.SheetID = ((elementsByTagName[sheetIndex].Attributes["ID"] == null) ? "" : ((XmlNode)elementsByTagName[sheetIndex].Attributes["ID"]).Value);
		hTKKCMCSheetInfo.SheetName = ((elementsByTagName[sheetIndex].Attributes["SheetName"] == null) ? "" : ((XmlNode)elementsByTagName[sheetIndex].Attributes["SheetName"]).Value);
		hTKKCMCSheetInfo.Caption = ((elementsByTagName[sheetIndex].Attributes["Caption"] == null) ? "" : ((XmlNode)elementsByTagName[sheetIndex].Attributes["Caption"]).Value);
		hTKKCMCSheetInfo.DataFile = text;
		hTKKCMCSheetInfo.Active = ((elementsByTagName[sheetIndex].Attributes["Active"] == null) ? "" : ((XmlNode)elementsByTagName[sheetIndex].Attributes["Active"]).Value);
		hTKKCMCSheetInfo.SheetIndex = (short)sheetIndex;
		hTKKCMCSheetInfo.ShowSearchButton = ((elementsByTagName[sheetIndex].Attributes["ShowSearchButton"] == null) ? "" : ((XmlNode)elementsByTagName[sheetIndex].Attributes["ShowSearchButton"]).Value);
		hTKKCMCSheetInfo.ShowTaiBangKeButton = ((elementsByTagName[sheetIndex].Attributes["ShowTaiBangKeButton"] == null) ? "" : ((XmlNode)elementsByTagName[sheetIndex].Attributes["ShowTaiBangKeButton"]).Value);
		hTKKCMCSheetInfo.strPrintBarcode = ((elementsByTagName[sheetIndex].Attributes["PrintBarCode"] == null) ? "" : ((XmlNode)elementsByTagName[sheetIndex].Attributes["PrintBarCode"]).Value);
		hTKKCMCSheetInfo.XMLnode = ((elementsByTagName[sheetIndex].Attributes["XMLnode"] == null) ? "" : ((XmlNode)elementsByTagName[sheetIndex].Attributes["XMLnode"]).Value);
		hTKKCMCSheetInfo.CoDieuChinhSoLieu = ((elementsByTagName[sheetIndex].Attributes["CoDieuChinhSoLieu"] == null) ? (-1) : int.Parse(((XmlNode)elementsByTagName[sheetIndex].Attributes["CoDieuChinhSoLieu"]).Value));
		hTKKCMCSheetInfo.RowTTDLT = ((elementsByTagName[sheetIndex].Attributes["RowTTDLT"] == null) ? "" : ((XmlNode)elementsByTagName[sheetIndex].Attributes["RowTTDLT"]).Value);
		hTKKCMCSheetInfo.TaiDuLieuLon = ((elementsByTagName[sheetIndex].Attributes["TaiDuLieuLon"] == null) ? "" : ((XmlNode)elementsByTagName[sheetIndex].Attributes["TaiDuLieuLon"]).Value.Trim());
		hTKKCMCSheetInfo.TaiDuLieuRatLon = ((elementsByTagName[sheetIndex].Attributes["TaiDuLieuRatLon"] == null) ? "" : ((XmlNode)elementsByTagName[sheetIndex].Attributes["TaiDuLieuRatLon"]).Value.Trim());
		HTKKCMCSheetInfo hTKKCMCSheetInfo2 = hTKKCMCSheetInfo;
		hTKKCMCSheetInfo2.GetDetailSheetInfo(text2);
		return hTKKCMCSheetInfo2;
	}

	private static List<string> GetPhuLucThemMoi(SheetView sView, XmlNodeList elemList)
	{
		List<string> list = new List<string>();
		try
		{
			for (int i = 1; i < elemList.Count; i++)
			{
				string text = ((elemList[i].Attributes["SheetName"] == null) ? null : ((XmlNode)elemList[i].Attributes["SheetName"]).Value);
				if (elemList[i].Attributes["Active"] != null && !text.Contains("KHBS") && !text.Contains("01-KHBS") && !text.Contains("01-1/KHBS"))
				{
					if (sView == null && WinSettings.DicSheet.ContainsKey(text))
					{
						list.Add(text);
					}
					else if (sView != null && !WinSettings.DicSheet.ContainsKey(text) && sView.Cells[i - 1, 0].Value != null && (bool)sView.Cells[i - 1, 0].Value)
					{
						list.Add(text);
					}
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			list = new List<string>();
		}
		return list;
	}

	public static Dictionary<string, int> GetPhuLucDaKeKhai(bool bPhuLucKeKhaiVisible, string strFileNameExtend)
	{
		//IL_0008: Unknown result type (might be due to invalid IL or missing references)
		//IL_000e: Expected O, but got Unknown
		Dictionary<string, int> dictionary = new Dictionary<string, int>();
		try
		{
			XmlDocument val = new XmlDocument();
			val.LoadXml(WinSettings.xmlMenuNode.OuterXml);
			XmlNode val2 = ((XmlNode)val).SelectSingleNode("/Menu");
			string versionXML = WinSettings.TTPeriodInfo.VersionXML;
			XmlNode val3 = val2.FirstChild.ChildNodes[0];
			string dataFile = ((val3.Attributes["DataFile"] == null) ? "" : ((XmlNode)val3.Attributes["DataFile"]).Value);
			string text = ((val3.Attributes["DataFile"] == null) ? "" : ((XmlNode)val3.Attributes["DataFile"]).Value);
			string key = ((val3.Attributes["SheetName"] == null) ? "" : ((XmlNode)val3.Attributes["SheetName"]).Value);
			dictionary.Add(key, 1);
			for (int i = 1; i < val2.FirstChild.ChildNodes.Count; i++)
			{
				val3 = val2.FirstChild.ChildNodes[i];
				string text2 = ((val3.Attributes["ID"] == null) ? "" : ((XmlNode)val3.Attributes["ID"]).Value);
				text = ((val3.Attributes["DataFile"] == null) ? "" : ((XmlNode)val3.Attributes["DataFile"]).Value);
				key = ((val3.Attributes["SheetName"] == null) ? "" : ((XmlNode)val3.Attributes["SheetName"]).Value);
				if (key.Contains("KHBS") || key.Contains("01-KHBS") || !(!key.Contains("01-1/KHBS") & (key != "27MT-TNCN")))
				{
					continue;
				}
				switch (key)
				{
				case "PL 01-1/TTĐB":
				case "01-2/TTĐB":
				case "01-3/TTĐB":
					if (WinSettings.TTPeriodInfo.KieuKy == "D")
					{
						continue;
					}
					break;
				}
				string fileNameBySheet = WinDoWorkEvent.GetFileNameBySheet(WinSettings.TTPeriodInfo, text, HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst + "\\", versionXML);
				string path = ((WinSettings.MenuID != "N7") ? fileNameBySheet.Replace("bs1_", "") : fileNameBySheet);
				string fileNameBySheet2 = WinDoWorkEvent.GetFileNameBySheet(WinSettings.TTPeriodInfo, dataFile, HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst + "\\", versionXML);
				if (File.Exists(fileNameBySheet))
				{
					dictionary.Add(key, 1);
				}
				else if (File.Exists(path) && !File.Exists(fileNameBySheet2))
				{
					dictionary.Add(key, 1);
				}
				else if (!bPhuLucKeKhaiVisible)
				{
					dictionary.Add(key, 2);
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			dictionary = null;
		}
		return dictionary;
	}

	public static void ConverDataVb6(string pathDataFileOld, out string listFileError)
	{
		//IL_0149: Unknown result type (might be due to invalid IL or missing references)
		//IL_0150: Expected O, but got Unknown
		//IL_0157: Unknown result type (might be due to invalid IL or missing references)
		//IL_015e: Expected O, but got Unknown
		//IL_01df: Unknown result type (might be due to invalid IL or missing references)
		//IL_01e6: Expected O, but got Unknown
		//IL_04c4: Unknown result type (might be due to invalid IL or missing references)
		//IL_04cb: Expected O, but got Unknown
		//IL_0969: Unknown result type (might be due to invalid IL or missing references)
		//IL_0970: Expected O, but got Unknown
		//IL_08da: Unknown result type (might be due to invalid IL or missing references)
		//IL_08e1: Expected O, but got Unknown
		//IL_0aad: Unknown result type (might be due to invalid IL or missing references)
		//IL_0ab4: Expected O, but got Unknown
		//IL_0ad0: Unknown result type (might be due to invalid IL or missing references)
		//IL_0ad7: Expected O, but got Unknown
		//IL_02a7: Unknown result type (might be due to invalid IL or missing references)
		//IL_02ae: Expected O, but got Unknown
		//IL_0b1b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0b22: Expected O, but got Unknown
		//IL_11af: Unknown result type (might be due to invalid IL or missing references)
		//IL_11b6: Expected O, but got Unknown
		//IL_0b6d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0b74: Expected O, but got Unknown
		//IL_11e9: Unknown result type (might be due to invalid IL or missing references)
		//IL_11f0: Expected O, but got Unknown
		//IL_0bb4: Unknown result type (might be due to invalid IL or missing references)
		//IL_0bbb: Expected O, but got Unknown
		//IL_0c94: Unknown result type (might be due to invalid IL or missing references)
		//IL_0c9b: Expected O, but got Unknown
		//IL_395a: Unknown result type (might be due to invalid IL or missing references)
		//IL_3961: Expected O, but got Unknown
		//IL_12f2: Unknown result type (might be due to invalid IL or missing references)
		//IL_12f9: Expected O, but got Unknown
		//IL_0bce: Unknown result type (might be due to invalid IL or missing references)
		//IL_0bd5: Expected O, but got Unknown
		//IL_134c: Unknown result type (might be due to invalid IL or missing references)
		//IL_1353: Expected O, but got Unknown
		//IL_14c2: Unknown result type (might be due to invalid IL or missing references)
		//IL_14c9: Expected O, but got Unknown
		//IL_2a36: Unknown result type (might be due to invalid IL or missing references)
		//IL_2a3d: Expected O, but got Unknown
		//IL_1636: Unknown result type (might be due to invalid IL or missing references)
		//IL_163d: Expected O, but got Unknown
		//IL_22d8: Unknown result type (might be due to invalid IL or missing references)
		//IL_22df: Expected O, but got Unknown
		//IL_2a7e: Unknown result type (might be due to invalid IL or missing references)
		//IL_2a85: Expected O, but got Unknown
		//IL_2dc2: Unknown result type (might be due to invalid IL or missing references)
		//IL_2dc9: Expected O, but got Unknown
		//IL_1670: Unknown result type (might be due to invalid IL or missing references)
		//IL_1677: Expected O, but got Unknown
		//IL_1763: Unknown result type (might be due to invalid IL or missing references)
		//IL_176a: Expected O, but got Unknown
		//IL_22fe: Unknown result type (might be due to invalid IL or missing references)
		//IL_2305: Expected O, but got Unknown
		//IL_2a9f: Unknown result type (might be due to invalid IL or missing references)
		//IL_2aa6: Expected O, but got Unknown
		//IL_2de8: Unknown result type (might be due to invalid IL or missing references)
		//IL_2def: Expected O, but got Unknown
		//IL_2356: Unknown result type (might be due to invalid IL or missing references)
		//IL_235d: Expected O, but got Unknown
		//IL_1892: Unknown result type (might be due to invalid IL or missing references)
		//IL_1899: Expected O, but got Unknown
		//IL_2f2a: Unknown result type (might be due to invalid IL or missing references)
		//IL_2f31: Expected O, but got Unknown
		//IL_2e6a: Unknown result type (might be due to invalid IL or missing references)
		//IL_2e71: Expected O, but got Unknown
		//IL_20e6: Unknown result type (might be due to invalid IL or missing references)
		//IL_20ed: Expected O, but got Unknown
		//IL_18b5: Unknown result type (might be due to invalid IL or missing references)
		//IL_18bc: Expected O, but got Unknown
		//IL_1ada: Unknown result type (might be due to invalid IL or missing references)
		//IL_1ae1: Expected O, but got Unknown
		//IL_1b1a: Unknown result type (might be due to invalid IL or missing references)
		//IL_1b21: Expected O, but got Unknown
		//IL_1d2a: Unknown result type (might be due to invalid IL or missing references)
		//IL_1d31: Expected O, but got Unknown
		//IL_1ded: Unknown result type (might be due to invalid IL or missing references)
		//IL_1df4: Expected O, but got Unknown
		//IL_1e10: Unknown result type (might be due to invalid IL or missing references)
		//IL_1e17: Expected O, but got Unknown
		string text = "";
		listFileError = "";
		string empty = string.Empty;
		string empty2 = string.Empty;
		string empty3 = string.Empty;
		string empty4 = string.Empty;
		string empty5 = string.Empty;
		string empty6 = string.Empty;
		string empty7 = string.Empty;
		string empty8 = string.Empty;
		string empty9 = string.Empty;
		string empty10 = string.Empty;
		string empty11 = string.Empty;
		string text2 = string.Empty;
		string text3 = string.Empty;
		string text4 = string.Empty;
		string empty12 = string.Empty;
		string text5 = string.Empty;
		int num = 0;
		string text6 = "01;94;96;97;03";
		try
		{
			int num2 = pathDataFileOld.LastIndexOf("\\");
			string text7 = pathDataFileOld.Substring(num2 + 1, pathDataFileOld.Length - num2 - 1).Replace(".xml", "");
			text = text7;
			if (text7.IndexOf("_") == -1 || text7.IndexOf("Header_01") > -1)
			{
				return;
			}
			string text8 = text7;
			if (text7.IndexOf("bs") == 0)
			{
				text8 = text7.Replace(text7.Split(new char[1] { '_' })[0] + "_", "");
			}
			XmlDocument val = new XmlDocument();
			XmlDataDocument xmlDataDocument = new XmlDataDocument();
			XmlDocument val2 = new XmlDocument();
			val2.LoadXml(HTKKFile.LoadFile("ConverDataFile.xml"));
			xmlDataDocument = WinSettings.xmlDocMenu;
			XmlNodeList val3 = ((XmlNode)xmlDataDocument).SelectNodes("/Root/Menu/Validity/Sheet");
			XmlNodeList val4 = null;
			if (HTKKFormater.isNullString(text3) || HTKKFormater.isNullString(text3) || HTKKFormater.isNullString(text4) || HTKKFormater.isNullString(text5))
			{
				val4 = ((XmlNode)val2).SelectNodes("/Root/Map");
				foreach (XmlNode item3 in val4)
				{
					XmlNode val5 = item3;
					if (val5.Attributes["DataFileOld"] == null || text8.IndexOf(((XmlNode)val5.Attributes["DataFileOld"]).Value) != 0)
					{
						continue;
					}
					text2 = ((val5.Attributes["DataFileOld"] == null) ? "" : ((XmlNode)val5.Attributes["DataFileOld"]).Value);
					text3 = ((val5.Attributes["DataFile"] == null) ? "" : ((XmlNode)val5.Attributes["DataFile"]).Value);
					foreach (XmlNode item4 in val3)
					{
						XmlNode val6 = item4;
						if (val6.Attributes["DataFile"] != null && (text3 + "_").IndexOf(((XmlNode)val6.Attributes["DataFile"]).Value.ToString() + "_") == 0)
						{
							text4 = ((val6.ParentNode.ParentNode.Attributes["ID"] == null) ? "" : ((XmlNode)val6.ParentNode.ParentNode.Attributes["ID"]).Value);
							text5 = ((val6.ParentNode.ParentNode.Attributes["VersionXML"] == null) ? "" : ((XmlNode)val6.ParentNode.ParentNode.Attributes["VersionXML"]).Value.Replace(".", ""));
							if (text7.IndexOf("KHBS") >= 0)
							{
								text5 = ((val6.ParentNode.ParentNode.Attributes["VersionKHBS"] == null) ? "" : ((XmlNode)val6.ParentNode.ParentNode.Attributes["VersionKHBS"]).Value.Replace(".", ""));
							}
							break;
						}
					}
					break;
				}
			}
			if (HTKKFormater.isNullString(text3) || HTKKFormater.isNullString(text3) || HTKKFormater.isNullString(text4) || HTKKFormater.isNullString(text5))
			{
				foreach (XmlNode item5 in val3)
				{
					XmlNode val5 = item5;
					int num3 = 0;
					if (val5.Attributes["DataFile"] != null && ((XmlNode)val5.Attributes["DataFile"]).Value == "02_TNDN_TT36_DK")
					{
						num3 = 1;
					}
					if (val5.Attributes["DataFile"] == null || text8.IndexOf(((XmlNode)val5.Attributes["DataFile"]).Value + "_") != 0 || HTKKFormater.isNullString(((XmlNode)val5.Attributes["DataFile"]).Value) || ((XmlNode)val5.Attributes["DataFile"]).Value.Equals("KHBS"))
					{
						continue;
					}
					string text9 = text8.Replace(((XmlNode)val5.Attributes["DataFile"]).Value + "_", "");
					if ((!text9.StartsWith("DT") && !text9.StartsWith("CTDC") && !text9.StartsWith("LNDC") && !text9.StartsWith("L") && !text9.StartsWith("KTN") && !text9.StartsWith("CD") && !text9.StartsWith("Q") && text9.Length > 2 && ConvertHelpers.ConvertToInt(text9.Substring(0, 2)) == 0 && !text9.StartsWith("DK")) || (text9.StartsWith("DK") && ((XmlNode)val5.Attributes["DataFile"]).Value.IndexOf("DK") == -1))
					{
						continue;
					}
					text2 = ((val5.Attributes["DataFile"] == null) ? "" : ((XmlNode)val5.Attributes["DataFile"]).Value);
					text3 = ((val5.Attributes["DataFile"] == null) ? "" : ((XmlNode)val5.Attributes["DataFile"]).Value);
					text4 = ((val5.ParentNode.ParentNode.Attributes["ID"] == null) ? "" : ((XmlNode)val5.ParentNode.ParentNode.Attributes["ID"]).Value);
					text5 = ((val5.ParentNode.ParentNode.Attributes["VersionXML"] == null) ? "" : ((XmlNode)val5.ParentNode.ParentNode.Attributes["VersionXML"]).Value.Replace(".", ""));
					if (text7.IndexOf("KHBS") >= 0)
					{
						text5 = ((val5.ParentNode.ParentNode.Attributes["VersionKHBS"] == null) ? "" : ((XmlNode)val5.ParentNode.ParentNode.Attributes["VersionKHBS"]).Value.Replace(".", ""));
					}
					break;
				}
			}
			if (HTKKFormater.isNullString(text3) || HTKKFormater.isNullString(text3) || HTKKFormater.isNullString(text4) || HTKKFormater.isNullString(text5))
			{
				listFileError = listFileError + text.PadRight(50, ' ') + "Không tìm thấy phiên bản tờ khai trên ứng dụng mới.";
				return;
			}
			val4 = ((XmlNode)val2).SelectNodes("/Root/Map[@DataFile='" + text3 + "']");
			int num4 = 0;
			if (text4 == "01")
			{
				foreach (XmlNode item6 in val4)
				{
					XmlNode val7 = item6;
					if (val7.Attributes["ROW_NCKT_HTKK"] != null)
					{
						num4 = ConvertHelpers.ConvertToInt(((XmlNode)val7.Attributes["ROW_NCKT_HTKK"]).Value);
						break;
					}
				}
			}
			Dictionary<string, string> dictionary = new Dictionary<string, string>();
			foreach (XmlNode item7 in val4)
			{
				XmlNode val7 = item7;
				if (val7.Attributes["CELL_NEW"] != null && val7.Attributes["CELL_OLD"] != null && !HTKKFormater.isNullString(((XmlNode)val7.Attributes["CELL_NEW"]).Value) && !HTKKFormater.isNullString(((XmlNode)val7.Attributes["CELL_OLD"]).Value))
				{
					HTKKFile.AddOrUpdateDicStr(dictionary, ((XmlNode)val7.Attributes["CELL_NEW"]).Value, ((XmlNode)val7.Attributes["CELL_OLD"]).Value);
				}
			}
			empty11 = text7.Substring(text7.IndexOf(text2));
			empty11 = text7.Replace(text2 + "_", "");
			empty12 = HTKKConfig.INTERFACE_TEMPLATES + "\\xml\\" + text3 + "_" + text5 + ".xml";
			XmlDocument val8 = new XmlDocument();
			if (!File.Exists(empty12))
			{
				return;
			}
			val8.Load(empty12);
			XmlDocument val9 = new XmlDocument();
			val9.Load(pathDataFileOld);
			XmlNodeList val10 = ((XmlNode)val8).SelectNodes("/Sections/Section");
			XmlNodeList val11 = ((XmlNode)val9).SelectNodes("/Sections/Section");
			int num5 = 0;
			int num6 = 0;
			int num7 = 0;
			foreach (XmlNode item8 in val10)
			{
				XmlNode val5 = item8;
				if (((XmlNode)val5.Attributes["Dynamic"]).Value == "0")
				{
					List<XmlNode> list = new List<XmlNode>();
					foreach (XmlNode item9 in val11)
					{
						XmlNode val12 = item9;
						if (!(((XmlNode)val12.Attributes["Dynamic"]).Value == "0"))
						{
							continue;
						}
						foreach (XmlNode item10 in val12)
						{
							XmlNode val13 = item10;
							foreach (XmlNode item11 in val13)
							{
								XmlNode item = item11;
								list.Add(item);
							}
						}
					}
					XmlNodeList childNodes = val5.ChildNodes[0].ChildNodes;
					foreach (XmlNode item12 in childNodes)
					{
						XmlNode item = item12;
						if (text4 == "30" && text3.IndexOf("01_1_TTS") > -1 && ((XmlNode)item.Attributes["CellID"]).Value == "B_15")
						{
							((XmlNode)item.Attributes["Value"]).Value = "VB6";
						}
						string key = "";
						string value = "";
						string value2 = ((XmlNode)item.Attributes["CellID"]).Value;
						if (value2.IndexOf("_") > -1)
						{
							key = ((ConvertHelpers.ConvertToInt(value2.Split(new char[1] { '_' })[1]) + num4 - num5 <= num) ? (value2.Split(new char[1] { '_' })[0] + "_" + (ConvertHelpers.ConvertToInt(value2.Split(new char[1] { '_' })[1]) + num4)) : (value2.Split(new char[1] { '_' })[0] + "_" + (ConvertHelpers.ConvertToInt(value2.Split(new char[1] { '_' })[1]) + num4 - num5)));
						}
						dictionary.TryGetValue(key, out value);
						if (!HTKKFormater.isNullString(value))
						{
							for (int i = 0; i < list.Count; i++)
							{
								string[] array = ((XmlNode)list[i].Attributes["CellID"]).Value.ToString().Split(new char[1] { '_' });
								string text10 = array[0] + "_" + (ConvertHelpers.ConvertToInt(array[1]) + num4);
								string[] array2 = value.Split(new char[1] { '_' });
								string text11 = array2[0] + "_" + (ConvertHelpers.ConvertToInt(array2[1]) - num4 + num5);
								if (text11 == text10)
								{
									string text12 = ((XmlNode)list[i].Attributes["Value"]).Value;
									if (text12.IndexOf("{") > -1 && text3 == "01_1_BC26_AC")
									{
										string key2 = text12.Split(new char[1] { '{' })[1].Replace("}", "");
										text12 = HTKKWinDictionaryDiaDanh.CoQuanThueQuanLy[key2];
									}
									if (text12.IndexOf(" - ") > 0)
									{
										text12 = text12.Substring(text12.IndexOf(" - ") + 3);
									}
									if (text12.IndexOf("         ") > -1 && Regex.Split(text12, "         ").Length > 2)
									{
										text12 = Regex.Split(text12, "         ")[0];
									}
									((XmlNode)item.Attributes["Value"]).Value = text12;
									break;
								}
							}
							continue;
						}
						int num8 = 0;
						bool flag = false;
						for (num8 = 0; num8 < list.Count; num8++)
						{
							string[] array = ((XmlNode)list[num8].Attributes["CellID"]).Value.ToString().Split(new char[1] { '_' });
							string text10 = array[0] + "_" + (ConvertHelpers.ConvertToInt(array[1]) + num4);
							if (((XmlNode)item.Attributes["CellID"]).Value == text10)
							{
								string text12 = ((XmlNode)list[num8].Attributes["Value"]).Value;
								if (text12.IndexOf(" - ") > 0)
								{
									text12 = text12.Substring(text12.IndexOf(" - ") + 3);
								}
								((XmlNode)item.Attributes["Value"]).Value = text12;
								flag = true;
								break;
							}
						}
					}
					continue;
				}
				XmlNode val14 = null;
				XmlNode val15 = null;
				int num9 = 0;
				foreach (XmlNode childNode in val5.ChildNodes)
				{
					XmlNode val16 = childNode;
					if (val16.Name == "Location")
					{
						foreach (XmlNode childNode2 in val16.ChildNodes)
						{
							XmlNode val17 = childNode2;
							if (val17.Name == "DataLocationTemplateRowStart")
							{
								num9 = ConvertHelpers.ConvertToInt(val17.InnerText);
							}
						}
					}
					if (val16.Name == "RowInfo")
					{
						val14 = val16;
					}
					if (val16.Name == "Header")
					{
						val15 = val16;
					}
				}
				XmlNode val18 = null;
				val18 = ((val15 == null) ? val14.FirstChild : val15.FirstChild);
				XmlNode val19 = null;
				foreach (XmlNode item13 in val11)
				{
					XmlNode val20 = item13;
					XmlNode val21 = null;
					if (val20.HasChildNodes)
					{
						val21 = val20.FirstChild;
					}
					if (val21 == null || !val21.HasChildNodes)
					{
						continue;
					}
					foreach (XmlNode childNode3 in val21.ChildNodes)
					{
						XmlNode val16 = childNode3;
						if (val15 != null)
						{
							if (((XmlNode)val16.Attributes["CellID"]).Value == ((XmlNode)val18.Attributes["CellID"]).Value)
							{
								val19 = val21.ParentNode;
								break;
							}
							continue;
						}
						string text10 = ((XmlNode)val18.Attributes["CellID"]).Value.ToString() + "_" + (num9 - num4);
						if (!(((XmlNode)val16.Attributes["CellID"]).Value == text10))
						{
							continue;
						}
						val19 = val21.ParentNode;
						break;
					}
				}
				XmlNodeList val22 = null;
				if (val19 == null || !val19.HasChildNodes)
				{
					continue;
				}
				val22 = val19.ChildNodes;
				int num10 = 0;
				int num11 = 0;
				List<XmlNode> list2 = new List<XmlNode>();
				foreach (XmlNode item14 in val22)
				{
					XmlNode val23 = item14;
					if ((((XmlNode)val23.LastChild.Attributes["Value"]).Value == "x" && text6.IndexOf(text4) > -1) || (val23.LastChild.Attributes["PageBreak"] != null && ((XmlNode)val23.LastChild.Attributes["PageBreak"]).Value == "1" && text6.IndexOf(text4) > -1) || (val15 != null && text6.IndexOf(text4) == -1))
					{
						num6++;
						num7++;
						XmlNode item2 = val5.Clone();
						list2.Add(item2);
					}
				}
				if (val15 != null && list2.Count == 0)
				{
					listFileError = listFileError + text.PadRight(50, ' ') + "Không có nhà máy thủy điện.";
					return;
				}
				if (list2.Count > 0)
				{
					for (int j = 0; j < list2.Count; j++)
					{
						foreach (XmlNode childNode4 in list2[j].ChildNodes)
						{
							XmlNode val16 = childNode4;
							if (val16.Name == "Location")
							{
								foreach (XmlNode childNode5 in val16.ChildNodes)
								{
									XmlNode val17 = childNode5;
									if (val17.Name == "DataLocationTemplateRowStart")
									{
										num9 = ConvertHelpers.ConvertToInt(val17.InnerText);
									}
								}
							}
							if (val16.Name == "RowInfo")
							{
								val14 = val16;
							}
							if (val16.Name == "Header")
							{
								val15 = val16;
							}
						}
						int num12 = 0;
						string text13 = "";
						string text10 = "";
						foreach (XmlNode item15 in val22)
						{
							XmlNode val23 = item15;
							text10 = ((XmlNode)val23.FirstChild.Attributes["CellID"]).Value.ToString().Split(new char[1] { '_' })[1];
							if (ConvertHelpers.ConvertToInt(text10) <= num11)
							{
								continue;
							}
							num11 = ConvertHelpers.ConvertToInt(text10);
							num = num11;
							if ((((XmlNode)val23.LastChild.Attributes["Value"]).Value == "x" && text6.IndexOf(text4) > -1) || (val23.LastChild.Attributes["PageBreak"] != null && ((XmlNode)val23.LastChild.Attributes["PageBreak"]).Value == "1" && text6.IndexOf(text4) > -1) || (text6.IndexOf(text4) == -1 && val23.HasChildNodes))
							{
								num12++;
								foreach (XmlNode childNode6 in val23.ChildNodes)
								{
									XmlNode val24 = childNode6;
									foreach (XmlNode childNode7 in val15.ChildNodes)
									{
										XmlNode val25 = childNode7;
										text10 = ((XmlNode)val24.Attributes["CellID"]).Value.ToString().Split(new char[1] { '_' })[0] + "_" + (ConvertHelpers.ConvertToInt(((XmlNode)val24.Attributes["CellID"]).Value.ToString().Split(new char[1] { '_' })[1]) + num10);
										if (((XmlNode)val25.Attributes["CellID"]).Value.ToString() == text10)
										{
											string text12 = ((XmlNode)val24.Attributes["Value"]).Value;
											if (text12.IndexOf(" - ") > 0)
											{
												text12 = text12.Substring(text12.IndexOf(" - ") + 3);
											}
											if (text12.IndexOf("         ") > -1 && Regex.Split(text12, "         ").Length > 2)
											{
												text12 = Regex.Split(text12, "         ")[0];
											}
											((XmlNode)val25.Attributes["Value"]).Value = text12;
										}
									}
								}
								if (num12 != 2)
								{
									continue;
								}
								if (j < list2.Count - 1)
								{
									XmlNode val26 = list2[j + 1];
									XmlNode val27 = null;
									XmlNode val28 = null;
									XmlNode val29 = null;
									XmlNode val30 = null;
									XmlNode val31 = null;
									foreach (XmlNode childNode8 in val26.ChildNodes)
									{
										XmlNode val16 = childNode8;
										if (val16.Name == "Location")
										{
											foreach (XmlNode childNode9 in val16.ChildNodes)
											{
												XmlNode val17 = childNode9;
												if (val17.Name == "TableTemplateRowStart")
												{
													val27 = val17;
												}
												if (val17.Name == "TableReportRowStart")
												{
													val28 = val17;
												}
												if (val17.Name == "DataLocationTemplateRowStart")
												{
													val29 = val17;
												}
												if (val17.Name == "DataLocationReportRowStart")
												{
													val30 = val17;
												}
											}
										}
										if (val16.Name == "RowInfo")
										{
											val14 = val16;
										}
										if (val16.Name == "Header")
										{
											val31 = val16;
										}
									}
									if (val27 != null)
									{
										val27.InnerText = (3 + num10 + num5 + ConvertHelpers.ConvertToInt(val27.InnerText)).ToString();
									}
									if (val28 != null)
									{
										val28.InnerText = (3 + num10 + num5 + ConvertHelpers.ConvertToInt(val28.InnerText)).ToString();
									}
									if (val29 != null)
									{
										val29.InnerText = (3 + num10 + num5 + ConvertHelpers.ConvertToInt(val29.InnerText)).ToString();
									}
									if (val30 != null)
									{
										val30.InnerText = (3 + num10 + num5 + ConvertHelpers.ConvertToInt(val30.InnerText)).ToString();
									}
									string text14 = "";
									if (val31 != null)
									{
										foreach (XmlNode childNode10 in val31.ChildNodes)
										{
											XmlNode val32 = childNode10;
											text14 = ((XmlNode)val32.Attributes["CellID"]).Value.ToString().Split(new char[1] { '_' })[0];
											((XmlNode)val32.Attributes["CellID"]).Value = text14 + "_" + (ConvertHelpers.ConvertToInt(val29.InnerText) - 1);
										}
										foreach (XmlNode childNode11 in val23.ChildNodes)
										{
											XmlNode val24 = childNode11;
											foreach (XmlNode childNode12 in val31.ChildNodes)
											{
												XmlNode val25 = childNode12;
												text10 = ((XmlNode)val24.Attributes["CellID"]).Value.ToString().Split(new char[1] { '_' })[0] + "_" + (ConvertHelpers.ConvertToInt(((XmlNode)val24.Attributes["CellID"]).Value.ToString().Split(new char[1] { '_' })[1]) + num10);
												if (((XmlNode)val25.Attributes["CellID"]).Value.ToString() == text10)
												{
													((XmlNode)val25.Attributes["Value"]).Value = ((XmlNode)val24.Attributes["Value"]).Value;
												}
											}
										}
									}
									text14 = ((XmlNode)val26.Attributes["TableSTTLocation"]).Value.ToString().Split(new char[1] { '_' })[0];
									((XmlNode)val26.Attributes["TableSTTLocation"]).Value = text14 + "_" + (ConvertHelpers.ConvertToInt(val29.InnerText) - 1);
									text14 = ((XmlNode)val26.Attributes["TableReportSTTLocation"]).Value.ToString().Split(new char[1] { '_' })[0];
									((XmlNode)val26.Attributes["TableReportSTTLocation"]).Value = text14 + "_" + (int.Parse(val28.InnerText) + 1);
								}
								num12 = 0;
								break;
							}
							if (num12 == 2)
							{
								continue;
							}
							text13 = "";
							int num13 = ConvertHelpers.ConvertToInt(((XmlNode)val23.FirstChild.Attributes["CellID"]).Value.Split(new char[1] { '_' })[1]);
							int num14 = ConvertHelpers.ConvertToInt(((XmlNode)val15.FirstChild.Attributes["CellID"]).Value.Split(new char[1] { '_' })[1]);
							foreach (XmlNode childNode13 in val23.ChildNodes)
							{
								XmlNode item = childNode13;
								string text12 = ((XmlNode)item.Attributes["Value"]).Value;
								if (text12.IndexOf("         ") > -1 && Regex.Split(text12, "         ").Length > 2)
								{
									text12 = Regex.Split(text12, "         ")[0];
								}
								text13 = text13 + text12 + "###";
							}
							num10++;
							text13 = text13.Substring(0, text13.LastIndexOf("###"));
							XmlNode lastChild = list2[j].LastChild;
							lastChild.InnerText = text13;
							if (lastChild.Attributes["RowDisplay"] != null)
							{
								((XmlNode)lastChild.Attributes["RowDisplay"]).Value = "1";
							}
							XmlNode item2 = lastChild.Clone();
							list2[j].InsertAfter(item2, lastChild);
						}
						if (num10 > 0)
						{
							list2[j].RemoveChild(list2[j].LastChild);
						}
					}
					num5 = ((num7 != 0) ? (num5 + num7 * 2 + num10 - 1 - 2) : (num5 + num10 - 1));
					foreach (XmlNode childNode14 in val5.ParentNode.ChildNodes)
					{
						XmlNode val33 = childNode14;
						XmlNode val31 = null;
						foreach (XmlNode childNode15 in val33.ChildNodes)
						{
							XmlNode val16 = childNode15;
							if (!(val16.Name == "Header") && !(val16.Name == "Cells"))
							{
								continue;
							}
							foreach (XmlNode childNode16 in val16.ChildNodes)
							{
								XmlNode item = childNode16;
								string obj = ((XmlNode)item.Attributes["CellID"]).Value.ToString().Split(new char[1] { '_' })[1];
								if (num9 - num5 >= ConvertHelpers.ConvertToInt(obj))
								{
									continue;
								}
								((XmlNode)item.Attributes["CellID"]).Value = ((XmlNode)item.Attributes["CellID"]).Value.ToString().Split(new char[1] { '_' })[0] + "_" + (ConvertHelpers.ConvertToInt(((XmlNode)item.Attributes["CellID"]).Value.ToString().Split(new char[1] { '_' })[1]) + num5);
								((XmlNode)item.Attributes["CellID2"]).Value = ((XmlNode)item.Attributes["CellID2"]).Value.ToString().Split(new char[1] { '_' })[0] + "_" + (ConvertHelpers.ConvertToInt(((XmlNode)item.Attributes["CellID2"]).Value.ToString().Split(new char[1] { '_' })[1]) + num5);
								if (item.Attributes["CellID3"] != null)
								{
									((XmlNode)item.Attributes["CellID3"]).Value = ((XmlNode)item.Attributes["CellID3"]).Value.ToString().Split(new char[1] { '_' })[0] + "_" + (ConvertHelpers.ConvertToInt(((XmlNode)item.Attributes["CellID3"]).Value.ToString().Split(new char[1] { '_' })[1]) + num5);
								}
								if (item.Attributes["ChildCell"] != null)
								{
									((XmlNode)item.Attributes["ChildCell"]).Value = ((XmlNode)item.Attributes["ChildCell"]).Value.ToString().Split(new char[1] { '_' })[0] + "_" + (ConvertHelpers.ConvertToInt(((XmlNode)item.Attributes["ChildCell"]).Value.ToString().Split(new char[1] { '_' })[1]) + num5);
								}
								if (item.Attributes["SelectedValue"] != null)
								{
									((XmlNode)item.Attributes["SelectedValue"]).Value = ((XmlNode)item.Attributes["SelectedValue"]).Value.ToString().Split(new char[1] { '_' })[0] + "_" + (ConvertHelpers.ConvertToInt(((XmlNode)item.Attributes["SelectedValue"]).Value.ToString().Split(new char[1] { '_' })[1]) + num5);
								}
								if (item.Attributes["ParentCell"] != null)
								{
									((XmlNode)item.Attributes["ParentCell"]).Value = ((XmlNode)item.Attributes["ParentCell"]).Value.ToString().Split(new char[1] { '_' })[0] + "_" + (ConvertHelpers.ConvertToInt(((XmlNode)item.Attributes["ParentCell"]).Value.ToString().Split(new char[1] { '_' })[1]) + num5);
								}
								if (val16.Name == "Header")
								{
									string text14 = "";
									if (val33.Attributes["TableSTTLocation"] != null)
									{
										text14 = ((XmlNode)val33.Attributes["TableSTTLocation"]).Value.ToString().Split(new char[1] { '_' })[0];
										((XmlNode)val33.Attributes["TableSTTLocation"]).Value = text14 + "_" + (ConvertHelpers.ConvertToInt(((XmlNode)val33.Attributes["TableSTTLocation"]).Value.ToString().Split(new char[1] { '_' })[1]) + num5);
									}
									if (val33.Attributes["TableReportSTTLocation"] != null)
									{
										((XmlNode)val33.Attributes["TableReportSTTLocation"]).Value = text14 + "_" + (int.Parse(((XmlNode)val33.Attributes["TableReportSTTLocation"]).Value.ToString().Split(new char[1] { '_' })[1]) + num5);
									}
								}
							}
						}
					}
					for (int j = list2.Count - 1; j >= 0; j--)
					{
						val5.ParentNode.InsertAfter(list2[j], val5);
					}
					val5.ParentNode.RemoveChild(val5);
					continue;
				}
				num5 = val22.Count - 1;
				if (num < num9 - num5)
				{
					num = num9 - num5;
				}
				foreach (XmlNode item16 in val22)
				{
					XmlNode val23 = item16;
					string text13 = "";
					XmlNodeList childNodes2 = val5.ChildNodes[1].ChildNodes;
					if (childNodes2 == null)
					{
						continue;
					}
					foreach (XmlNode item17 in childNodes2)
					{
						XmlNode val34 = item17;
						bool flag2 = false;
						foreach (XmlNode item18 in val23)
						{
							XmlNode item = item18;
							if (val34.Attributes["CellID"] != null && item.Attributes["CellID"] != null && ((XmlNode)item.Attributes["CellID"]).Value.IndexOf("_") > 0 && ((XmlNode)val34.Attributes["CellID"]).Value == ((XmlNode)item.Attributes["CellID"]).Value.Split(new char[1] { '_' })[0])
							{
								string text12 = ((XmlNode)item.Attributes["Value"]).Value;
								if (text12.IndexOf(" - ") > 0)
								{
									text12 = text12.Substring(text12.IndexOf(" - ") + 3);
								}
								if (text12.IndexOf("         ") > -1 && Regex.Split(text12, "         ").Length > 2)
								{
									text12 = Regex.Split(text12, "         ")[0];
								}
								text13 = text13 + text12.Replace("###", "") + "###";
								flag2 = true;
								break;
							}
						}
						if (flag2)
						{
							continue;
						}
						if (val34.Attributes["Value"] != null)
						{
							string text12 = ((XmlNode)val34.Attributes["Value"]).Value;
							if (text12.IndexOf("         ") > -1 && Regex.Split(text12, "         ").Length > 2)
							{
								text12 = Regex.Split(text12, "         ")[0];
							}
							text13 = text13 + text12.Replace("###", "") + "###";
						}
						else
						{
							text13 += "###";
						}
					}
					text13 = text13.Substring(0, text13.LastIndexOf("###"));
					XmlNode lastChild = val5.LastChild;
					lastChild.InnerText = text13;
					XmlNode item2 = lastChild.Clone();
					val5.InsertAfter(item2, lastChild);
				}
				val5.RemoveChild(val5.LastChild);
				foreach (XmlNode childNode17 in val5.ParentNode.ChildNodes)
				{
					XmlNode val33 = childNode17;
					XmlNode val31 = null;
					foreach (XmlNode childNode18 in val33.ChildNodes)
					{
						XmlNode val16 = childNode18;
						if (val16.Name == "Location" && val16.SelectNodes("DataLocationTemplateRowStart")[0] != null && ConvertHelpers.ConvertToInt(val16.SelectNodes("DataLocationTemplateRowStart")[0].InnerText) > num9)
						{
							foreach (XmlNode childNode19 in val16.ChildNodes)
							{
								XmlNode val35 = childNode19;
								if (ConvertHelpers.ConvertToInt(val35.InnerXml) > 0)
								{
									val35.InnerXml = (ConvertHelpers.ConvertToInt(val35.InnerXml) + num5).ToString();
								}
							}
						}
						if (!(val16.Name == "Header") && !(val16.Name == "Cells"))
						{
							continue;
						}
						foreach (XmlNode childNode20 in val16.ChildNodes)
						{
							XmlNode item = childNode20;
							if (item.Attributes["CellID"] == null || ((XmlNode)item.Attributes["CellID"]).Value.IndexOf("_") <= -1)
							{
								continue;
							}
							string obj = ((XmlNode)item.Attributes["CellID"]).Value.ToString().Split(new char[1] { '_' })[1];
							if (num9 < ConvertHelpers.ConvertToInt(obj))
							{
								if (item.Attributes["CellID"] != null && ((XmlNode)item.Attributes["CellID"]).Value.IndexOf("_") > -1)
								{
									((XmlNode)item.Attributes["CellID"]).Value = ((XmlNode)item.Attributes["CellID"]).Value.ToString().Split(new char[1] { '_' })[0] + "_" + (ConvertHelpers.ConvertToInt(((XmlNode)item.Attributes["CellID"]).Value.ToString().Split(new char[1] { '_' })[1]) + num5);
								}
								if (item.Attributes["CellID2"] != null && ((XmlNode)item.Attributes["CellID2"]).Value.IndexOf("_") > -1)
								{
									((XmlNode)item.Attributes["CellID2"]).Value = ((XmlNode)item.Attributes["CellID2"]).Value.ToString().Split(new char[1] { '_' })[0] + "_" + (ConvertHelpers.ConvertToInt(((XmlNode)item.Attributes["CellID2"]).Value.ToString().Split(new char[1] { '_' })[1]) + num5);
								}
								if (item.Attributes["CellID3"] != null)
								{
									((XmlNode)item.Attributes["CellID3"]).Value = ((XmlNode)item.Attributes["CellID3"]).Value.ToString().Split(new char[1] { '_' })[0] + "_" + (ConvertHelpers.ConvertToInt(((XmlNode)item.Attributes["CellID3"]).Value.ToString().Split(new char[1] { '_' })[1]) + num5);
								}
								if (item.Attributes["ChildCell"] != null)
								{
									((XmlNode)item.Attributes["ChildCell"]).Value = ((XmlNode)item.Attributes["ChildCell"]).Value.ToString().Split(new char[1] { '_' })[0] + "_" + (ConvertHelpers.ConvertToInt(((XmlNode)item.Attributes["ChildCell"]).Value.ToString().Split(new char[1] { '_' })[1]) + num5);
								}
								if (item.Attributes["SelectedValue"] != null)
								{
									((XmlNode)item.Attributes["SelectedValue"]).Value = ((XmlNode)item.Attributes["SelectedValue"]).Value.ToString().Split(new char[1] { '_' })[0] + "_" + (ConvertHelpers.ConvertToInt(((XmlNode)item.Attributes["SelectedValue"]).Value.ToString().Split(new char[1] { '_' })[1]) + num5);
								}
								if (item.Attributes["ParentCell"] != null)
								{
									((XmlNode)item.Attributes["ParentCell"]).Value = ((XmlNode)item.Attributes["ParentCell"]).Value.ToString().Split(new char[1] { '_' })[0] + "_" + (ConvertHelpers.ConvertToInt(((XmlNode)item.Attributes["ParentCell"]).Value.ToString().Split(new char[1] { '_' })[1]) + num5);
								}
								if (val16.Name == "Header")
								{
									string text14 = ((XmlNode)val33.Attributes["TableSTTLocation"]).Value.ToString().Split(new char[1] { '_' })[0];
									((XmlNode)val33.Attributes["TableSTTLocation"]).Value = text14 + "_" + (ConvertHelpers.ConvertToInt(((XmlNode)val33.Attributes["TableSTTLocation"]).Value.ToString().Split(new char[1] { '_' })[1]) + num5);
									((XmlNode)val33.Attributes["TableReportSTTLocation"]).Value = text14 + "_" + (int.Parse(((XmlNode)val33.Attributes["TableReportSTTLocation"]).Value.ToString().Split(new char[1] { '_' })[1]) + num5);
								}
							}
						}
					}
				}
			}
			val8.DocumentElement.SetAttribute("Version", WinSettings.VersionApp);
			string text15 = HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst + "\\" + text7.Replace(text2, text3) + "_" + text5.Replace(".", "") + ".xml";
			if ("66;10;67;18;09;65;68;13;14;62;38".IndexOf(text4) > -1 || text4 == "35" || text4 == "60" || text4 == "61")
			{
				if ("66;67;09;10;60;61;G2;".IndexOf(text4) > -1)
				{
					text15 = HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst + "\\" + text3 + "_" + text5.Replace(".", "") + ".xml";
				}
				else
				{
					string text16 = text15.Substring(text15.IndexOf(text3));
					text16 = text16.Replace(text3 + "_", "");
					text15 = ((!text16.StartsWith("T")) ? text15.Replace(text16, "Q" + text16) : text15.Replace(text16, text16.Substring(1)));
				}
			}
			if (text4 == "63")
			{
				string text17 = text15.Split(new char[1] { '\\' })[text15.Split(new char[1] { '\\' }).Length - 1];
				if (text17.Length > 27)
				{
					string newValue = text17.Substring(0, 13) + "CSM_" + text17.Substring(13, 16);
					text15 = text15.Replace(text17, newValue);
				}
			}
			if ("A1,H1,A3,H3,A8,H8,A9,H9".Contains(text4) || "X1".Equals(WinSettings.TTPeriodInfo.menuID))
			{
				text15 = HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst + "\\" + text7.Replace(text2 + "_", text3 + "_Q") + "_" + text5.Replace(".", "") + ".xml";
			}
			XmlTextWriter val36 = new XmlTextWriter(text15, Encoding.UTF8);
			try
			{
				val36.Formatting = (Formatting)1;
				val8.Save((XmlWriter)(object)val36);
			}
			finally
			{
				((IDisposable)val36)?.Dispose();
			}
		}
		catch (Exception ex)
		{
			listFileError = listFileError + text.PadRight(50, ' ') + "Lỗi chuyển đổi.";
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			HTKKFile.LogRun(MethodBase.GetCurrentMethod().Name, "Loi conver TK: " + pathDataFileOld);
		}
	}

	public static void ConverDataVersion(string dataFileName, string versionXmlOld, string versionXml)
	{
		//IL_0164: Unknown result type (might be due to invalid IL or missing references)
		//IL_016b: Expected O, but got Unknown
		//IL_01d1: Unknown result type (might be due to invalid IL or missing references)
		//IL_01d8: Expected O, but got Unknown
		string text = string.Empty;
		string empty = string.Empty;
		string empty2 = string.Empty;
		string empty3 = string.Empty;
		string empty4 = string.Empty;
		string empty5 = string.Empty;
		try
		{
			string[] array = dataFileName.Split(new char[1] { '~' });
			DirectoryInfo directoryInfo = new DirectoryInfo(HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst);
			string path = HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst + "\\BK";
			if (!Directory.Exists(path))
			{
				Directory.CreateDirectory(path);
			}
			string[] array2 = array;
			foreach (string text2 in array2)
			{
				FileInfo[] files = directoryInfo.GetFiles(text2 + "_*.xml");
				foreach (FileInfo fileInfo in files)
				{
					empty = fileInfo.Name;
					text = fileInfo.FullName;
					empty2 = empty.Replace(text2 + "_", "");
					empty3 = text2 + "_" + empty2.Replace("_" + versionXmlOld + ".xml", "_" + versionXml + ".xml");
					empty4 = fileInfo.DirectoryName + "\\BK\\" + empty;
					if (!File.Exists(fileInfo.DirectoryName + "\\" + empty3))
					{
						XmlDocument val = new XmlDocument();
						val.Load(text);
						val.DocumentElement.SetAttribute("Version", WinSettings.VersionApp);
						string text3 = HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst + "\\" + empty3;
						XmlTextWriter val2 = new XmlTextWriter(text3, Encoding.UTF8);
						try
						{
							val2.Formatting = (Formatting)1;
							val.Save((XmlWriter)(object)val2);
						}
						finally
						{
							((IDisposable)val2)?.Dispose();
						}
						Directory.Move(text, empty4);
					}
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			HTKKFile.LogRun(MethodBase.GetCurrentMethod().Name, "Loi conver TK: " + text);
		}
	}

	public static void ConverDataVersion02GTGT(string dataFileName, string versionXmlOld, string versionXml)
	{
		//IL_01c2: Unknown result type (might be due to invalid IL or missing references)
		//IL_01c9: Expected O, but got Unknown
		//IL_01f3: Unknown result type (might be due to invalid IL or missing references)
		//IL_01fa: Expected O, but got Unknown
		//IL_0322: Unknown result type (might be due to invalid IL or missing references)
		//IL_0329: Expected O, but got Unknown
		string text = string.Empty;
		string empty = string.Empty;
		string empty2 = string.Empty;
		string empty3 = string.Empty;
		string empty4 = string.Empty;
		string empty5 = string.Empty;
		string empty6 = string.Empty;
		try
		{
			string[] array = dataFileName.Split(new char[1] { '~' });
			DirectoryInfo directoryInfo = new DirectoryInfo(HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst);
			string path = HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst + "\\BK";
			if (!Directory.Exists(path))
			{
				Directory.CreateDirectory(path);
			}
			empty6 = HTKKConfig.INTERFACE_TEMPLATES + "\\xml\\" + dataFileName + "_" + versionXml + ".xml";
			string[] array2 = array;
			foreach (string text2 in array2)
			{
				FileInfo[] files = directoryInfo.GetFiles("*" + text2 + "*.xml");
				foreach (FileInfo fileInfo in files)
				{
					if (fileInfo.Name.Contains("KHBS"))
					{
						continue;
					}
					empty = fileInfo.Name;
					text = fileInfo.FullName;
					empty2 = empty.Replace(text2 + "_", "");
					empty3 = empty.Replace("_" + versionXmlOld + ".xml", "_" + versionXml + ".xml");
					empty4 = fileInfo.DirectoryName + "\\BK\\" + empty;
					if (File.Exists(fileInfo.DirectoryName + "\\" + empty3))
					{
						continue;
					}
					XmlDocument val = new XmlDocument();
					if (!File.Exists(empty6))
					{
						return;
					}
					val.Load(empty6);
					XmlNodeList val2 = ((XmlNode)val).SelectNodes("/Sections/Section");
					XmlDocument val3 = new XmlDocument();
					val3.Load(text);
					XmlNodeList val4 = ((XmlNode)val3).SelectNodes("/Sections/Section");
					XmlNode val5 = val4.Item(1);
					if (val5.ChildNodes[0].ChildNodes.Count == 15)
					{
						XmlNode val6 = val2.Item(1).ChildNodes[0];
						XmlNodeList childNodes = val6.ChildNodes;
						for (int num = 9; num >= 0; num--)
						{
							XmlNode val7 = childNodes.Item(num);
							XmlNode val8 = val7.Clone();
							XmlNode val9 = val3.ImportNode(val8, true);
							val5.ChildNodes[0].InsertBefore(val9, val5.ChildNodes[0].FirstChild);
						}
					}
					val3.DocumentElement.SetAttribute("Version", WinSettings.VersionApp);
					string text3 = HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst + "\\" + empty3;
					XmlTextWriter val10 = new XmlTextWriter(text3, Encoding.UTF8);
					try
					{
						val10.Formatting = (Formatting)1;
						val3.Save((XmlWriter)(object)val10);
					}
					finally
					{
						((IDisposable)val10)?.Dispose();
					}
					Directory.Move(text, empty4);
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			HTKKFile.LogRun(MethodBase.GetCurrentMethod().Name, "Loi conver TK: " + text);
		}
	}

	public static void ConverDataVersionCommon(string dataFileName, string versionXmlOld, string versionXmlNew, string convertType, int insertSection, int restart, string versionXmlKHBSOld, string versionXmlKHBSNew, int versionApp, bool noRemoveAfterBK, string prefix)
	{
		//IL_247e: Unknown result type (might be due to invalid IL or missing references)
		//IL_2485: Expected O, but got Unknown
		//IL_2572: Unknown result type (might be due to invalid IL or missing references)
		//IL_2579: Expected O, but got Unknown
		//IL_06f2: Unknown result type (might be due to invalid IL or missing references)
		//IL_06f9: Expected O, but got Unknown
		//IL_06f9: Unknown result type (might be due to invalid IL or missing references)
		//IL_0700: Expected O, but got Unknown
		//IL_033b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0342: Expected O, but got Unknown
		//IL_07b8: Unknown result type (might be due to invalid IL or missing references)
		//IL_07bf: Expected O, but got Unknown
		//IL_0439: Unknown result type (might be due to invalid IL or missing references)
		//IL_0440: Expected O, but got Unknown
		//IL_0871: Unknown result type (might be due to invalid IL or missing references)
		//IL_0878: Expected O, but got Unknown
		//IL_04ef: Unknown result type (might be due to invalid IL or missing references)
		//IL_04f6: Expected O, but got Unknown
		//IL_23cb: Unknown result type (might be due to invalid IL or missing references)
		//IL_23d2: Expected O, but got Unknown
		//IL_09f8: Unknown result type (might be due to invalid IL or missing references)
		//IL_09ff: Expected O, but got Unknown
		//IL_1149: Unknown result type (might be due to invalid IL or missing references)
		//IL_1150: Expected O, but got Unknown
		//IL_0c99: Unknown result type (might be due to invalid IL or missing references)
		//IL_0ca0: Expected O, but got Unknown
		//IL_0d06: Unknown result type (might be due to invalid IL or missing references)
		//IL_0d10: Expected O, but got Unknown
		//IL_1ab9: Unknown result type (might be due to invalid IL or missing references)
		//IL_1ac3: Expected O, but got Unknown
		//IL_1ef8: Unknown result type (might be due to invalid IL or missing references)
		//IL_1f02: Expected O, but got Unknown
		string text = string.Empty;
		string empty = string.Empty;
		string empty2 = string.Empty;
		string text2 = string.Empty;
		string empty3 = string.Empty;
		string empty4 = string.Empty;
		string empty5 = string.Empty;
		try
		{
			string[] array = dataFileName.Split(new char[1] { '~' });
			DirectoryInfo directoryInfo = new DirectoryInfo(HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst);
			string path = HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst + "\\BK";
			if (!Directory.Exists(path))
			{
				Directory.CreateDirectory(path);
			}
			string[] array2 = array;
			foreach (string text3 in array2)
			{
				empty5 = HTKKConfig.INTERFACE_TEMPLATES + "\\xml\\" + text3 + "_" + versionXmlNew + ".xml";
				bool flag = false;
				FileInfo[] files = directoryInfo.GetFiles("*" + text3 + "*.xml");
				foreach (FileInfo fileInfo in files)
				{
					if (!fileInfo.Name.EndsWith(versionXmlOld + ".xml"))
					{
						continue;
					}
					if (fileInfo.Name.IndexOf("03_01_ND132_TNDN_TT80") >= 0)
					{
						DeleteCellSameSInfoStatic(fileInfo);
					}
					string text6;
					if (convertType == "2" || convertType == "3")
					{
						empty = fileInfo.Name;
						text = fileInfo.FullName;
						empty2 = empty.Replace(text3 + "_", "");
						for (int k = 0; k < 2; k++)
						{
							string text4 = versionXmlOld;
							string text5 = versionXmlNew;
							if (k == 1 && (string.IsNullOrEmpty(versionXmlKHBSOld) || string.IsNullOrEmpty(versionXmlKHBSNew)))
							{
								continue;
							}
							if (k == 1)
							{
								text4 = versionXmlKHBSOld;
								text5 = versionXmlKHBSNew;
							}
							if ((k == 0 && fileInfo.Name.Contains("KHBS")) || (k == 1 && !fileInfo.Name.Contains("KHBS")))
							{
								continue;
							}
							text2 = empty.Replace("_" + text4 + ".xml", "_" + text5 + ".xml");
							empty3 = fileInfo.DirectoryName + "\\BK\\" + empty.Replace("_" + text5 + ".xml", "_" + text4 + ".xml");
							if ((!string.IsNullOrEmpty(prefix) && text2.IndexOf(prefix) >= 0) || (empty.EndsWith(versionXmlNew + ".xml") && !string.IsNullOrEmpty(prefix) && text2.IndexOf(prefix) >= 0) || (empty.EndsWith(versionXmlNew + ".xml") && string.IsNullOrEmpty(prefix)))
							{
								continue;
							}
							XmlDocument val = new XmlDocument();
							val.Load(text);
							text6 = HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst + "\\" + text2;
							XmlNodeList val2 = ((XmlNode)val).SelectNodes("/Sections");
							if (val2.Count > 0 && versionApp > 0)
							{
								string value = ((XmlNode)val2[0].Attributes["Version"]).Value;
								value = value.Replace(".", "");
								if (Convert.ToInt32(value) >= versionApp)
								{
									continue;
								}
							}
							XmlTextWriter val3;
							if (!string.IsNullOrEmpty(prefix) && text2.StartsWith("bs"))
							{
								string text7 = text.Replace(text3, text3 + prefix);
								val3 = new XmlTextWriter(text7, Encoding.UTF8);
								try
								{
									val3.Formatting = (Formatting)1;
									val.Save((XmlWriter)(object)val3);
								}
								finally
								{
									((IDisposable)val3)?.Dispose();
								}
							}
							if (!File.Exists(empty3))
							{
								if (noRemoveAfterBK)
								{
									File.Copy(text, empty3, overwrite: true);
								}
								else
								{
									Directory.Move(text, empty3);
								}
							}
							else if (noRemoveAfterBK)
							{
								File.Copy(text, empty3, overwrite: true);
							}
							else
							{
								File.Delete(empty3);
								Directory.Move(text, empty3);
							}
							val.DocumentElement.SetAttribute("Version", GeneralFunctions.Chen1KyTuGiuaCacTu(versionApp.ToString()));
							val3 = new XmlTextWriter(text6, Encoding.UTF8);
							try
							{
								val3.Formatting = (Formatting)1;
								val.Save((XmlWriter)(object)val3);
							}
							finally
							{
								((IDisposable)val3)?.Dispose();
							}
						}
					}
					if ((!flag && convertType == "5") || convertType == "6")
					{
						empty = fileInfo.Name;
						text = fileInfo.FullName;
						empty2 = empty.Replace(text3 + "_", "");
						for (int k = 0; k < 2; k++)
						{
							empty5 = HTKKConfig.INTERFACE_TEMPLATES + "\\xml\\" + text3 + "_" + versionXmlNew + ".xml";
							string text4 = versionXmlOld;
							string text5 = versionXmlNew;
							if (k == 1 && (string.IsNullOrEmpty(versionXmlKHBSOld) || string.IsNullOrEmpty(versionXmlKHBSNew)))
							{
								continue;
							}
							if (k == 1)
							{
								text4 = versionXmlKHBSOld;
								text5 = versionXmlKHBSNew;
							}
							if ((k == 0 && fileInfo.Name.Contains("KHBS")) || (k == 1 && !fileInfo.Name.Contains("KHBS")))
							{
								continue;
							}
							text2 = empty.Replace("_" + text4 + ".xml", "_" + text5 + ".xml");
							empty3 = fileInfo.DirectoryName + "\\BK\\" + empty.Replace("_" + text5 + ".xml", "_" + text4 + ".xml");
							if (!string.IsNullOrEmpty(prefix) && text2.IndexOf(prefix) >= 0)
							{
								continue;
							}
							XmlDocument val4 = new XmlDocument();
							XmlDocument val = new XmlDocument();
							val.Load(text);
							XmlNodeList val2 = ((XmlNode)val).SelectNodes("/Sections");
							if (val2.Count > 0 && versionApp > 0)
							{
								string value = ((XmlNode)val2[0].Attributes["Version"]).Value;
								value = value.Replace(".", "");
								if (Convert.ToInt32(value) >= versionApp)
								{
									continue;
								}
							}
							XmlTextWriter val3;
							if (!string.IsNullOrEmpty(prefix) && text2.StartsWith("bs"))
							{
								string text7 = text.Replace(text3, text3 + prefix);
								val3 = new XmlTextWriter(text7, Encoding.UTF8);
								try
								{
									val3.Formatting = (Formatting)1;
									val.Save((XmlWriter)(object)val3);
								}
								finally
								{
									((IDisposable)val3)?.Dispose();
								}
							}
							string value2 = "";
							if (fileInfo.Name.Contains("KHBS"))
							{
								value2 = TimMenuIdTuDatafile(text3);
							}
							if (k != 1 || !fileInfo.Name.Contains("KHBS") || fileInfo.Name.IndexOf("KHBS_01_1") >= 0 || "D1,N8,G9,H2,H4,T8,E8,L8,L7,N2,G2,R5,Y6,Z8,Z5".Contains(value2))
							{
								XmlNodeList val5 = ((XmlNode)val).SelectNodes("/Sections/Section");
								int num = 0;
								XmlDocument val6 = new XmlDocument();
								if (fileInfo.Name.Contains("KHBS"))
								{
									empty5 = (("D1,N8,G9,H2,H4,T8,E8,L8,L7,N2,G2,R5,Y6,Z8,Z5".Contains(value2) && !fileInfo.Name.Contains("KHBS_01_1")) ? (HTKKConfig.INTERFACE_TEMPLATES + "\\xml\\KHBS_01_TT80_DIA_BAN_" + versionXmlNew + ".xml") : ((!fileInfo.Name.Contains("KHBS_01_1")) ? (HTKKConfig.INTERFACE_TEMPLATES + "\\xml\\KHBS_01_TT80_" + versionXmlNew + ".xml") : (HTKKConfig.INTERFACE_TEMPLATES + "\\xml\\KHBS_01_1_TT80_" + versionXmlNew + ".xml")));
								}
								if (!File.Exists(empty5))
								{
									continue;
								}
								val6.Load(empty5);
								XmlNodeList val7 = ((XmlNode)val6).SelectNodes("/Sections/Section");
								for (int l = 0; l < val7.Count; l++)
								{
									if (l != insertSection || val7.Count == val5.Count)
									{
										continue;
									}
									XmlNode val8 = val.ImportNode(val7[l].Clone(), true);
									int num2 = 0;
									int num3 = 0;
									if (val8.Attributes["Dynamic"] != null && ((XmlNode)val8.Attributes["Dynamic"]).Value == "1")
									{
										foreach (XmlNode item in val5)
										{
											XmlNode val9 = item;
											num2++;
											if (val9.Attributes["Dynamic"] != null && ((XmlNode)val9.Attributes["Dynamic"]).Value == "1")
											{
												XmlNodeList val10 = val9.SelectNodes("RowData");
												num3 = num3 + val10.Count - 1;
											}
											if (num2 >= insertSection)
											{
												break;
											}
										}
										XmlNode val11 = val8.SelectSingleNode("Location");
										string innerText = val11.SelectSingleNode("TableTemplateRowStart").InnerText;
										string innerText2 = val11.SelectSingleNode("TableReportRowStart").InnerText;
										string innerText3 = val11.SelectSingleNode("DataLocationTemplateRowStart").InnerText;
										string innerText4 = val11.SelectSingleNode("DataLocationReportRowStart").InnerText;
										val11.SelectSingleNode("TableTemplateRowStart").InnerText = string.Concat(Convert.ToInt32(innerText) + num3);
										val11.SelectSingleNode("TableReportRowStart").InnerText = string.Concat(Convert.ToInt32(innerText2) + num3);
										val11.SelectSingleNode("DataLocationTemplateRowStart").InnerText = string.Concat(Convert.ToInt32(innerText3) + num3);
										val11.SelectSingleNode("DataLocationReportRowStart").InnerText = string.Concat(Convert.ToInt32(innerText4) + num3);
									}
									val5[l].ParentNode.InsertBefore(val8, val5[l].ParentNode.ChildNodes[insertSection]);
									num++;
								}
								val5 = ((XmlNode)val).SelectNodes("/Sections/Section");
								int num4 = 0;
								for (int l = 0; l < val7.Count; l++)
								{
									if (val7[l].Attributes["Dynamic"] != null && ((XmlNode)val7[l].Attributes["Dynamic"]).Value == "0")
									{
										XmlNode val9 = val5[num4];
										for (int m = 0; m < val7[l].ChildNodes.Count; m++)
										{
											int num5 = 0;
											foreach (XmlNode childNode in val7[l].ChildNodes[m].ChildNodes)
											{
												XmlNode val12 = childNode;
												string pathNew = ((XmlNode)val12.Attributes["Path"]).Value;
												if (string.IsNullOrEmpty(pathNew))
												{
													continue;
												}
												foreach (XmlAttribute item2 in (XmlNamedNodeMap)val12.Attributes)
												{
													XmlAttribute attr = item2;
													string attrNew = ((val12.Attributes[((XmlNode)attr).Name] == null) ? "" : ((XmlNode)val12.Attributes[((XmlNode)attr).Name]).Value);
													if (string.IsNullOrEmpty(attrNew) || ((XmlNode)attr).Name == "Value" || pathNew == attrNew)
													{
														continue;
													}
													if (val9 == null)
													{
														if (!flag)
														{
															HTKKFile.Log("Lỗi convert dataFileName" + dataFileName + "fileName" + empty + "elemListNew" + val7.Count + "elemListOld" + val5.Count + "versionXmlOld" + versionXmlOld + "versionXmlNew" + versionXmlNew + "convertType" + convertType + "insertSection" + insertSection + "restart" + restart + "versionXmlKHBSOld" + versionXmlKHBSOld + "versionXmlKHBSNew" + versionXmlKHBSNew + "versionApp" + versionApp + "pathNew" + pathNew);
															flag = true;
														}
														break;
													}
													XmlNode val13 = ((IEnumerable)val9.ChildNodes[m]).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode x) => x.Attributes["Path"] != null && ((XmlNode)x.Attributes["Path"]).Value == pathNew));
													XmlNode val14 = ((IEnumerable)val9.ChildNodes[m]).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode x) => x.Attributes["Path"] != null && ((XmlNode)x.Attributes["Path"]).Value == pathNew && x.Attributes[((XmlNode)attr).Name] != null && ((XmlNode)x.Attributes[((XmlNode)attr).Name]).Value == attrNew));
													if (val13 == null)
													{
														XmlNode val8 = val.ImportNode(val12.Clone(), true);
														val9.ChildNodes[m].InsertAfter(val8, val9.ChildNodes[m].ChildNodes[num5 - 1]);
														num++;
													}
													else if (val14 == null && val13.Attributes[((XmlNode)attr).Name] != null)
													{
														((XmlNode)val13.Attributes[((XmlNode)attr).Name]).Value = ((XmlNode)val12.Attributes[((XmlNode)attr).Name]).Value;
														num++;
													}
												}
												num5++;
											}
										}
										num4++;
										continue;
									}
									string text8 = ((val7[l].Attributes["Dynamic"] != null && ((XmlNode)val7[l].Attributes["Dynamic"]).Value == "1" && val7[l].Attributes["TableName"] != null) ? ((XmlNode)val7[l].Attributes["TableName"]).Value : "");
									foreach (XmlNode item3 in val5)
									{
										XmlNode val9 = item3;
										string text9 = ((val9.Attributes["Dynamic"] != null && ((XmlNode)val9.Attributes["Dynamic"]).Value == "1" && val9.Attributes["TableName"] != null) ? ((XmlNode)val9.Attributes["TableName"]).Value : "");
										if (string.IsNullOrEmpty(text8) || !(text8 == text9))
										{
											continue;
										}
										num4++;
										List<int> listIndexAddNew = new List<int>();
										List<string> list = new List<string>();
										for (int m = 0; m < val7[l].ChildNodes.Count; m++)
										{
											XmlNode val15 = val7[l].ChildNodes[m];
											if (val15 != null && val15.Name == "Location")
											{
												for (int num6 = 0; num6 < val15.ChildNodes.Count; num6++)
												{
													XmlNode val12 = val15.ChildNodes[num6];
													XmlNode val13 = val9.ChildNodes[m].ChildNodes[num6];
													if ((val13 != null && val12 != null && val13.InnerText == "-1" && val12.InnerText != val13.InnerText) || (val13.InnerText == "0" && val12.InnerText != val13.InnerText))
													{
														val13.InnerText = val12.InnerText;
													}
												}
												if (fileInfo.Name.IndexOf("01_1_TBVMT_TT80") >= 0 && val9.ChildNodes[m].ParentNode.Attributes["TableName"] != null && ((XmlNode)val9.ChildNodes[m].ParentNode.Attributes["TableName"]).Value == "01_1_TBVMT_4" && val9.ChildNodes[m].ChildNodes[0].Name == "TableTemplateRowStart" && val9.ChildNodes[m].ChildNodes[2].Name == "DataLocationTemplateRowStart" && val5[num4 - 2].ChildNodes[0].ChildNodes[0].Attributes["CellID"] != null && ((XmlNode)val5[num4 - 2].ChildNodes[0].ChildNodes[0].Attributes["CellID"]).Value.IndexOf(val9.ChildNodes[m].ChildNodes[0].InnerText) >= 0)
												{
													int num7 = Convert.ToInt32(val5[num4 - 3].ChildNodes[m].ChildNodes[0].InnerText);
													int num8 = Convert.ToInt32(val5[num4 - 3].ChildNodes[m].ChildNodes[1].InnerText);
													int num9 = Convert.ToInt32(val5[num4 - 3].ChildNodes.Count - 4);
													((XmlNode)val5[num4 - 2].ChildNodes[0].ChildNodes[0].Attributes["CellID"]).Value = "AL_" + (num7 + num9 + 7);
													val9.ChildNodes[m].ChildNodes[0].InnerText = (num7 + num9 + 8).ToString();
													val9.ChildNodes[m].ChildNodes[2].InnerText = (num7 + num9 + 10).ToString();
													((XmlNode)val5[num4 - 2].ChildNodes[0].ChildNodes[0].Attributes["CellID2"]).Value = "AL_" + (num8 + num9 + 8);
													val9.ChildNodes[m].ChildNodes[1].InnerText = (num8 + num9 + 9).ToString();
													val9.ChildNodes[m].ChildNodes[3].InnerText = (num8 + num9 + 11).ToString();
													num++;
												}
												if (fileInfo.Name.IndexOf("01_2_TBVMT_TT80") >= 0 && val9.ChildNodes[m].ParentNode.Attributes["TableName"] != null && ((XmlNode)val9.ChildNodes[m].ParentNode.Attributes["TableName"]).Value == "01_2_TBVMT_3" && ((val9.ChildNodes[m].ChildNodes[0].Name == "TableTemplateRowStart" && val9.ChildNodes[m].ChildNodes[2].Name == "DataLocationTemplateRowStart" && val5[num4 - 1].ChildNodes[0].Attributes["CellID"] != null && ((XmlNode)val5[num4 - 1].ChildNodes[0].ChildNodes[0].Attributes["CellID"]).Value.IndexOf(val9.ChildNodes[m].ChildNodes[0].InnerText) >= 0) || Convert.ToInt32(val9.ChildNodes[m].ChildNodes[0].InnerText) + 1 == Convert.ToInt32(val9.ChildNodes[m].ChildNodes[2].InnerText)))
												{
													int num7 = Convert.ToInt32(val5[num4 - 2].ChildNodes[m].ChildNodes[0].InnerText);
													int num8 = Convert.ToInt32(val5[num4 - 2].ChildNodes[m].ChildNodes[1].InnerText);
													int num9 = Convert.ToInt32(val5[num4 - 2].ChildNodes.Count - 4);
													((XmlNode)val5[num4 - 1].ChildNodes[0].ChildNodes[0].Attributes["CellID"]).Value = "AF_" + (num7 + num9 + 5);
													val9.ChildNodes[m].ChildNodes[0].InnerText = (num7 + num9 + 6).ToString();
													val9.ChildNodes[m].ChildNodes[2].InnerText = (num7 + num9 + 8).ToString();
													((XmlNode)val5[num4 - 1].ChildNodes[0].ChildNodes[0].Attributes["CellID2"]).Value = "AF_" + (num8 + num9 + 12);
													val9.ChildNodes[m].ChildNodes[1].InnerText = (num8 + num9 + 13).ToString();
													val9.ChildNodes[m].ChildNodes[3].InnerText = (num8 + num9 + 15).ToString();
													num++;
												}
											}
											if (val15 != null && val15.Name == "Header")
											{
												if (fileInfo.Name.IndexOf("01_2_TBVMT_TT80") >= 0)
												{
													continue;
												}
												for (int num5 = 0; num5 < val15.ChildNodes.Count; num5++)
												{
													XmlNode val12 = val15.ChildNodes[num5];
													string cellIdNew = ((XmlNode)val12.Attributes["CellID"]).Value;
													string pathNew2 = ((XmlNode)val12.Attributes["Path"]).Value;
													foreach (XmlAttribute item4 in (XmlNamedNodeMap)val12.Attributes)
													{
														XmlAttribute attr2 = item4;
														string attrNew2 = ((val12.Attributes[((XmlNode)attr2).Name] == null) ? "" : ((XmlNode)val12.Attributes[((XmlNode)attr2).Name]).Value);
														if (string.IsNullOrEmpty(attrNew2) || ((XmlNode)attr2).Name.IndexOf("Value") >= 0 || attrNew2 == cellIdNew)
														{
															continue;
														}
														XmlNode val13 = ((IEnumerable)val9.ChildNodes[m]).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode x) => x.Attributes["CellID"] != null && ((XmlNode)x.Attributes["CellID"]).Value == cellIdNew));
														XmlNode val14 = ((IEnumerable)val9.ChildNodes[m]).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode x) => x.Attributes["CellID"] != null && ((XmlNode)x.Attributes["CellID"]).Value == cellIdNew && x.Attributes[((XmlNode)attr2).Name] != null && ((XmlNode)x.Attributes[((XmlNode)attr2).Name]).Value == attrNew2));
														int num10 = 0;
														if (val13 == null)
														{
															val13 = ((IEnumerable)val9.ChildNodes[m]).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode x) => x.Attributes["Path"] != null && ((XmlNode)x.Attributes["Path"]).Value == pathNew2));
															if (val13 != null && val12 != null)
															{
																num10 = TinhOffset(val12, val13);
															}
														}
														if (val14 == null && num10 != 0)
														{
															string attrNewOffset = SetOffsetForCellIdOld(attrNew2, ((XmlNode)attr2).Name, num10);
															val14 = ((IEnumerable)val9.ChildNodes[m]).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode x) => x.Attributes["Path"] != null && ((XmlNode)x.Attributes["Path"]).Value == pathNew2 && x.Attributes[((XmlNode)attr2).Name] != null && ((XmlNode)x.Attributes[((XmlNode)attr2).Name]).Value == attrNewOffset));
														}
														if (val13 == null)
														{
															XmlNode val8 = val.ImportNode(val12.Clone(), true);
															val9.ChildNodes[m].InsertAfter(val8, val9.ChildNodes[m].ChildNodes[num5 - 1]);
															num++;
														}
														else if (val14 == null && val13.Attributes[((XmlNode)attr2).Name] != null)
														{
															((XmlNode)val13.Attributes[((XmlNode)attr2).Name]).Value = ((XmlNode)val12.Attributes[((XmlNode)attr2).Name]).Value;
															num++;
														}
													}
												}
											}
											if (val15 != null && val15.Name == "RowInfo")
											{
												listIndexAddNew = GetViTriThayDoi(val7[l], val9);
												val9.ChildNodes[m].InnerXml = val7[l].ChildNodes[m].InnerXml;
												num++;
											}
											if (val15 == null || !(val15.Name == "Footer") || fileInfo.Name.IndexOf("01_2_TBVMT_TT80") >= 0)
											{
												continue;
											}
											for (int num5 = 0; num5 < val15.ChildNodes.Count; num5++)
											{
												XmlNode val12 = val15.ChildNodes[num5];
												string cellIdNew2 = ((XmlNode)val12.Attributes["CellID"]).Value;
												string pathNew3 = ((XmlNode)val12.Attributes["Path"]).Value;
												foreach (XmlAttribute item5 in (XmlNamedNodeMap)val12.Attributes)
												{
													XmlAttribute attr3 = item5;
													string attrNew3 = ((val12.Attributes[((XmlNode)attr3).Name] == null) ? "" : ((XmlNode)val12.Attributes[((XmlNode)attr3).Name]).Value);
													if (string.IsNullOrEmpty(attrNew3) || ((XmlNode)attr3).Name.IndexOf("Value") >= 0 || attrNew3 == cellIdNew2)
													{
														continue;
													}
													XmlNode val13 = ((IEnumerable)val9.ChildNodes[m]).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode x) => x.Attributes["CellID"] != null && ((XmlNode)x.Attributes["CellID"]).Value == cellIdNew2));
													XmlNode val14 = ((IEnumerable)val9.ChildNodes[m]).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode x) => x.Attributes["CellID"] != null && ((XmlNode)x.Attributes["CellID"]).Value == cellIdNew2 && x.Attributes[((XmlNode)attr3).Name] != null && ((XmlNode)x.Attributes[((XmlNode)attr3).Name]).Value == attrNew3));
													int num10 = 0;
													if (val13 == null)
													{
														val13 = ((IEnumerable)val9.ChildNodes[m]).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode x) => x.Attributes["Path"] != null && ((XmlNode)x.Attributes["Path"]).Value == pathNew3));
														if (val13 != null && val12 != null)
														{
															num10 = TinhOffset(val12, val13);
														}
													}
													if (val14 == null && num10 != 0)
													{
														string attrNewOffset2 = SetOffsetForCellIdOld(attrNew3, ((XmlNode)attr3).Name, num10);
														val14 = ((IEnumerable)val9.ChildNodes[m]).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode x) => x.Attributes["Path"] != null && ((XmlNode)x.Attributes["Path"]).Value == pathNew3 && x.Attributes[((XmlNode)attr3).Name] != null && ((XmlNode)x.Attributes[((XmlNode)attr3).Name]).Value == attrNewOffset2));
													}
													if (val13 == null)
													{
														XmlNode val8 = val.ImportNode(val12.Clone(), true);
														val9.ChildNodes[m].InsertAfter(val8, val9.ChildNodes[m].ChildNodes[num5 - 1]);
														num++;
													}
													else if (val14 == null && val13.Attributes[((XmlNode)attr3).Name] != null)
													{
														((XmlNode)val13.Attributes[((XmlNode)attr3).Name]).Value = ((XmlNode)val12.Attributes[((XmlNode)attr3).Name]).Value;
														num++;
													}
												}
											}
										}
										if (num > 0 && (fileInfo.Name.IndexOf("01_6_GTGT_DVVT_TT80") < 0 || !(versionXmlNew == "283")) && (fileInfo.Name.IndexOf("01_6_GTGT_CSSX_TT80") < 0 || !(versionXmlNew == "283")))
										{
											InsertRowsValue(val7[l], val9, listIndexAddNew);
										}
									}
								}
								FixRowStatic(val6, ref val);
							}
							if (flag)
							{
								break;
							}
							if (!File.Exists(empty3))
							{
								if (noRemoveAfterBK)
								{
									File.Copy(text, empty3, overwrite: true);
								}
								else
								{
									Directory.Move(text, empty3);
								}
							}
							else if (noRemoveAfterBK)
							{
								File.Copy(text, empty3, overwrite: true);
							}
							else
							{
								File.Delete(empty3);
								Directory.Move(text, empty3);
							}
							val.DocumentElement.SetAttribute("Version", GeneralFunctions.Chen1KyTuGiuaCacTu(versionApp.ToString()));
							text6 = HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst + "\\" + text2;
							val3 = new XmlTextWriter(text6, Encoding.UTF8);
							try
							{
								val3.Formatting = (Formatting)1;
								val.Save((XmlWriter)(object)val3);
							}
							finally
							{
								((IDisposable)val3)?.Dispose();
							}
						}
					}
					if (flag || !(convertType == "6"))
					{
						continue;
					}
					empty = fileInfo.Name;
					text6 = HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst + "\\" + text2;
					bool dataChange = false;
					XmlDocument val16 = new XmlDocument();
					val16.Load(text6);
					List<MapDiaBan> mapDiaBan = getMapDiaBan("Catalogue_Dia_Ban_Tinh_Map.xml");
					val16 = ThayDoiMaTenDiaBanTinhStatic(val16, mapDiaBan, HTKKWinDictionaryDiaDanh.dicTinhTPMoi, "100", ref dataChange);
					val16 = ThayDoiMaTenDiaBanTinhDinamic(val16, mapDiaBan, HTKKWinDictionaryDiaDanh.dicTinhTPMoi, "100", ref dataChange);
					List<MapDiaBan> mapDiaBan2 = getMapDiaBan("Catalogue_Dia_Ban_Xa_Map.xml");
					val16 = ThayDoiMaTenDiaBanXaStatic(val16, mapDiaBan2, HTKKWinDictionaryDiaDanh._diaBanXa, "101", ref dataChange);
					val16 = ThayDoiMaTenDiaBanXaDinamic(val16, mapDiaBan2, HTKKWinDictionaryDiaDanh._diaBanXa, "101", ref dataChange);
					val16 = XoaHuyenDinamic(val16, "6", ref dataChange);
					val16 = XoaHuyenStatic(val16, "6", ref dataChange);
					ThayDoiTenCqtDinamic(val16, "8", ref dataChange);
					ThayDoiTenCqtDinamic(val16, "9", ref dataChange);
					ThayDoiTenCqtStatic(val16, "8", ref dataChange);
					ThayDoiTenCqtStatic(val16, "9", ref dataChange);
					if (dataChange)
					{
						XmlTextWriter val3 = new XmlTextWriter(text6, Encoding.UTF8);
						try
						{
							val3.Formatting = (Formatting)1;
							val16.Save((XmlWriter)(object)val3);
						}
						finally
						{
							((IDisposable)val3)?.Dispose();
						}
					}
					string text10 = renameFileDiaBan(mapDiaBan, text6);
					if (text10 != text6)
					{
						File.Move(text6, text10);
					}
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			HTKKFile.LogRun(MethodBase.GetCurrentMethod().Name, "Loi conver TK: " + text);
		}
	}

	public static string renameFileDiaBan(List<MapDiaBan> mapDiaBan, string path)
	{
		if (path.IndexOf("TINH") >= 0)
		{
			string[] array = path.Split(new char[1] { '_' });
			string text = "";
			string maTinhCu = "";
			string[] array2 = array;
			foreach (string text2 in array2)
			{
				if (text2.IndexOf("TINH") >= 0)
				{
					maTinhCu = text2.Replace("TINH", "");
					MapDiaBan mapDiaBan2 = mapDiaBan.Find((MapDiaBan x) => x.MaCu.IndexOf(maTinhCu) >= 0);
					if (mapDiaBan2 != null)
					{
						text = mapDiaBan2.MaMoi;
						break;
					}
				}
			}
			if (!string.IsNullOrEmpty(text))
			{
				path = path.Replace("TINH" + maTinhCu, "TINH" + text);
			}
		}
		return path;
	}

	public static XmlDocument ThayDoiMaTenDiaBanXaDinamic(XmlDocument doc, List<MapDiaBan> mapDiaBan, List<DiaBanXa> diaBan, string controltypeNew, ref bool dataChange)
	{
		//IL_0028: Unknown result type (might be due to invalid IL or missing references)
		//IL_002e: Expected O, but got Unknown
		try
		{
			string text = "/Sections/Section";
			string text2 = ".//RowInfo/Cell";
			XmlNodeList val = ((XmlNode)doc).SelectNodes(text);
			foreach (XmlNode item in val)
			{
				XmlNode val2 = item;
				XmlNodeList cellNodes = val2.SelectNodes(text2);
				int indexTen = -1;
				int indexMa = -1;
				if (timViTriTraCon(cellNodes, controltypeNew, ref indexTen, ref indexMa))
				{
					thayTheDuLieuXa(val2, mapDiaBan, diaBan, ref dataChange, indexTen, indexMa);
				}
			}
			return doc;
		}
		catch (Exception ex)
		{
			HTKKFile.LogRun(MethodBase.GetCurrentMethod().Name, ex.Message);
			return null;
		}
	}

	public static XmlDocument ThayDoiTenCqtDinamic(XmlDocument doc, string controltypeNew, ref bool dataChange)
	{
		//IL_0028: Unknown result type (might be due to invalid IL or missing references)
		//IL_002e: Expected O, but got Unknown
		try
		{
			string text = "/Sections/Section";
			string text2 = ".//RowInfo/Cell";
			XmlNodeList val = ((XmlNode)doc).SelectNodes(text);
			foreach (XmlNode item in val)
			{
				XmlNode val2 = item;
				XmlNodeList cellNodes = val2.SelectNodes(text2);
				int indexTen = -1;
				int indexMa = -1;
				if (timViTriTraCon(cellNodes, controltypeNew, ref indexTen, ref indexMa))
				{
					thayTheDuLieuCqt(val2, ref dataChange, indexTen, indexMa);
				}
			}
			return doc;
		}
		catch (Exception ex)
		{
			HTKKFile.LogRun(MethodBase.GetCurrentMethod().Name, ex.Message);
			return null;
		}
	}

	public static XmlNode thayTheDuLieuCqt(XmlNode section, ref bool dataChange, int indexTenCu, int indexMaCu)
	{
		//IL_0024: Unknown result type (might be due to invalid IL or missing references)
		//IL_002a: Expected O, but got Unknown
		string text = ".//RowData";
		XmlNodeList val = section.SelectNodes(text);
		foreach (XmlNode item in val)
		{
			XmlNode val2 = item;
			List<string> list = GeneralFunctions.SplitString2(val2.InnerText, "###");
			if (list.Count <= indexTenCu || list.Count <= indexMaCu)
			{
				continue;
			}
			string text2 = list[indexTenCu];
			string maCu = list[indexMaCu];
			if (string.IsNullOrEmpty(maCu))
			{
				continue;
			}
			CQTDiaBan cQTDiaBan = HTKKWinDictionaryDiaDanh._cqtDiaBan.Find((CQTDiaBan x) => x.MaCQT == maCu);
			if (cQTDiaBan != null)
			{
				string tenCQT = cQTDiaBan.TenCQT;
				if (!(list[indexTenCu] == tenCQT))
				{
					list[indexTenCu] = tenCQT;
					val2.InnerText = GeneralFunctions.ListToString(list, "###");
					dataChange = true;
				}
			}
		}
		return section;
	}

	public static XmlDocument ThayDoiTenCqtStatic(XmlDocument doc, string controltype, ref bool dataChange)
	{
		//IL_0040: Unknown result type (might be due to invalid IL or missing references)
		//IL_0047: Expected O, but got Unknown
		//IL_0078: Unknown result type (might be due to invalid IL or missing references)
		//IL_007f: Expected O, but got Unknown
		//IL_0203: Unknown result type (might be due to invalid IL or missing references)
		//IL_020a: Expected O, but got Unknown
		try
		{
			string text = "Value";
			string text2 = "SelectedValue";
			string text3 = "/Sections/Section";
			string text4 = ".//Cells/Cell";
			string text5 = ".//Header/Cell";
			XmlNodeList val = ((XmlNode)doc).SelectNodes(text3);
			foreach (XmlNode item in val)
			{
				XmlNode val2 = item;
				XmlNodeList val3 = val2.SelectNodes(text4 + "[@Controltype='" + controltype + "']");
				foreach (XmlNode item2 in val3)
				{
					XmlNode val4 = item2;
					string value = ((XmlNode)val4.Attributes[text]).Value;
					string value2 = ((XmlNode)val4.Attributes[text2]).Value;
					if (string.IsNullOrEmpty(value2))
					{
						continue;
					}
					XmlNodeList val5 = val2.SelectNodes(text4 + "[@CellID='" + value2 + "']");
					if (val5 == null || val5.Count != 1)
					{
						continue;
					}
					string maCu = ((XmlNode)val5[0].Attributes[text]).Value;
					if (!string.IsNullOrEmpty(maCu))
					{
						CQTDiaBan cQTDiaBan = HTKKWinDictionaryDiaDanh._cqtDiaBan.Find((CQTDiaBan x) => x.MaCQT == maCu);
						if (cQTDiaBan != null && !(((XmlNode)val4.Attributes[text]).Value == cQTDiaBan.TenCQT))
						{
							((XmlNode)val4.Attributes[text]).Value = cQTDiaBan.TenCQT;
							dataChange = true;
						}
					}
				}
				XmlNodeList val6 = val2.SelectNodes(text5 + "[@Controltype='" + controltype + "']");
				foreach (XmlNode item3 in val6)
				{
					XmlNode val4 = item3;
					string value = ((XmlNode)val4.Attributes[text]).Value;
					string value2 = ((XmlNode)val4.Attributes[text2]).Value;
					if (string.IsNullOrEmpty(value2))
					{
						continue;
					}
					XmlNodeList val7 = val2.SelectNodes(text5 + "[@CellID='" + value2 + "']");
					if (val7 == null || val7.Count != 1)
					{
						continue;
					}
					string maCu2 = ((XmlNode)val7[0].Attributes[text]).Value;
					if (!string.IsNullOrEmpty(maCu2))
					{
						CQTDiaBan cQTDiaBan = HTKKWinDictionaryDiaDanh._cqtDiaBan.Find((CQTDiaBan x) => x.MaCQT == maCu2);
						if (cQTDiaBan != null && !(((XmlNode)val4.Attributes[text]).Value == cQTDiaBan.TenCQT))
						{
							((XmlNode)val4.Attributes[text]).Value = cQTDiaBan.TenCQT;
							dataChange = true;
						}
					}
				}
			}
			return doc;
		}
		catch (Exception ex)
		{
			HTKKFile.LogRun(MethodBase.GetCurrentMethod().Name, ex.Message);
			return null;
		}
	}

	public static XmlDocument ThayDoiMaTenDiaBanTinhDinamic(XmlDocument doc, List<MapDiaBan> mapDiaBan, Dictionary<string, string> diaBanMoi, string controltypeNew, ref bool dataChange)
	{
		//IL_0028: Unknown result type (might be due to invalid IL or missing references)
		//IL_002e: Expected O, but got Unknown
		try
		{
			string text = "/Sections/Section";
			string text2 = ".//RowInfo/Cell";
			XmlNodeList val = ((XmlNode)doc).SelectNodes(text);
			foreach (XmlNode item in val)
			{
				XmlNode val2 = item;
				XmlNodeList cellNodes = val2.SelectNodes(text2);
				int indexTen = -1;
				int indexMa = -1;
				if (timViTriTraCon(cellNodes, controltypeNew, ref indexTen, ref indexMa))
				{
					thayTheDuLieuTinh(val2, mapDiaBan, diaBanMoi, ref dataChange, indexTen, indexMa);
				}
			}
			return doc;
		}
		catch (Exception ex)
		{
			HTKKFile.LogRun(MethodBase.GetCurrentMethod().Name, ex.Message);
			return null;
		}
	}

	public static XmlDocument XoaHuyenDinamic(XmlDocument doc, string controltypeNew, ref bool dataChange)
	{
		//IL_002b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0031: Expected O, but got Unknown
		//IL_0080: Unknown result type (might be due to invalid IL or missing references)
		//IL_0087: Expected O, but got Unknown
		try
		{
			string text = "/Sections/Section";
			string text2 = ".//RowInfo/Cell";
			XmlNodeList val = ((XmlNode)doc).SelectNodes(text);
			foreach (XmlNode item in val)
			{
				XmlNode val2 = item;
				XmlNodeList cellNodes = val2.SelectNodes(text2);
				int indexTen = -1;
				int indexMa = -1;
				if (!timViTriTraCon(cellNodes, controltypeNew, ref indexTen, ref indexMa))
				{
					continue;
				}
				string text3 = ".//RowData";
				XmlNodeList val3 = val2.SelectNodes(text3);
				foreach (XmlNode item2 in val3)
				{
					XmlNode val4 = item2;
					List<string> list = GeneralFunctions.SplitString2(val4.InnerText, "###");
					if (list.Count > indexTen && list.Count > indexMa)
					{
						list[indexMa] = "";
						list[indexTen] = "";
						val4.InnerText = GeneralFunctions.ListToString(list, "###");
						dataChange = true;
					}
				}
			}
			return doc;
		}
		catch (Exception ex)
		{
			HTKKFile.LogRun(MethodBase.GetCurrentMethod().Name, ex.Message);
			return null;
		}
	}

	public static XmlNode thayTheDuLieuXa(XmlNode section, List<MapDiaBan> mapDiaBan, List<DiaBanXa> diaBan, ref bool dataChange, int indexTen, int indexMa)
	{
		//IL_0024: Unknown result type (might be due to invalid IL or missing references)
		//IL_002a: Expected O, but got Unknown
		string text = ".//RowData";
		XmlNodeList val = section.SelectNodes(text);
		foreach (XmlNode item in val)
		{
			XmlNode val2 = item;
			List<string> list = GeneralFunctions.SplitString2(val2.InnerText, "###");
			if (list.Count <= indexTen || list.Count <= indexMa)
			{
				continue;
			}
			string text2 = list[indexTen];
			string ma = list[indexMa];
			if (string.IsNullOrEmpty(ma))
			{
				continue;
			}
			List<MapDiaBan> list2 = mapDiaBan.FindAll((MapDiaBan x) => x.MaCu.Contains(ma));
			if (list2 != null && list2.Count == 1)
			{
				string maMoi = list2[0].MaMoi;
				DiaBanXa diaBanXa = diaBan.Find((DiaBanXa x) => x.MaXa == maMoi);
				if (diaBanXa != null && !(list[indexMa] == maMoi))
				{
					list[indexMa] = maMoi;
					list[indexTen] = diaBanXa.TenXa;
					val2.InnerText = GeneralFunctions.ListToString(list, "###");
					dataChange = true;
				}
			}
		}
		return section;
	}

	public static XmlNode thayTheDuLieuTinh(XmlNode section, List<MapDiaBan> mapDiaBan, Dictionary<string, string> diaBanMoi, ref bool dataChange, int indexTen, int indexMa)
	{
		//IL_0024: Unknown result type (might be due to invalid IL or missing references)
		//IL_002a: Expected O, but got Unknown
		string text = ".//RowData";
		XmlNodeList val = section.SelectNodes(text);
		foreach (XmlNode item in val)
		{
			XmlNode val2 = item;
			List<string> list = GeneralFunctions.SplitString2(val2.InnerText, "###");
			if (list.Count <= indexTen || list.Count <= indexMa)
			{
				continue;
			}
			string text2 = list[indexTen];
			string ma = list[indexMa];
			if (string.IsNullOrEmpty(ma))
			{
				continue;
			}
			List<MapDiaBan> list2 = mapDiaBan.FindAll((MapDiaBan x) => x.MaCu.Contains(ma));
			if (list2 != null && list2.Count == 1)
			{
				string maMoi = list2[0].MaMoi;
				KeyValuePair<string, string> keyValuePair = diaBanMoi.FirstOrDefault((KeyValuePair<string, string> x) => x.Value == maMoi);
				if (!(list[indexMa] == maMoi))
				{
					list[indexMa] = maMoi;
					list[indexTen] = keyValuePair.Key;
					val2.InnerText = GeneralFunctions.ListToString(list, "###");
					dataChange = true;
				}
			}
		}
		return section;
	}

	public static bool timViTriTraCon(XmlNodeList cellNodes, string controltypeNew, ref int indexTen, ref int indexMa)
	{
		string text = "SelectedValue";
		string text2 = "CellID";
		string text3 = "Controltype";
		string text4 = "";
		for (int i = 0; i < cellNodes.Count; i++)
		{
			XmlNode val = cellNodes[i];
			string text5 = ((val.Attributes[text2] != null) ? ((XmlNode)val.Attributes[text2]).Value : "");
			string text6 = ((val.Attributes[text3] != null) ? ((XmlNode)val.Attributes[text3]).Value : "");
			if (text6 == controltypeNew)
			{
				indexTen = i;
				text4 = ((val.Attributes[text] != null) ? ((XmlNode)val.Attributes[text]).Value : "");
			}
			else if (!string.IsNullOrEmpty(text4) && text5 == text4)
			{
				indexMa = i;
			}
		}
		if (indexTen < 0 || indexMa < 0)
		{
			return false;
		}
		return true;
	}

	public static XmlDocument XoaHuyenStatic(XmlDocument doc, string controltype, ref bool dataChange)
	{
		//IL_0040: Unknown result type (might be due to invalid IL or missing references)
		//IL_0047: Expected O, but got Unknown
		//IL_0078: Unknown result type (might be due to invalid IL or missing references)
		//IL_007f: Expected O, but got Unknown
		//IL_01a5: Unknown result type (might be due to invalid IL or missing references)
		//IL_01ac: Expected O, but got Unknown
		try
		{
			string text = "Value";
			string text2 = "SelectedValue";
			string text3 = "/Sections/Section";
			string text4 = ".//Cells/Cell";
			string text5 = ".//Header/Cell";
			XmlNodeList val = ((XmlNode)doc).SelectNodes(text3);
			foreach (XmlNode item in val)
			{
				XmlNode val2 = item;
				XmlNodeList val3 = val2.SelectNodes(text4 + "[@Controltype='" + controltype + "']");
				foreach (XmlNode item2 in val3)
				{
					XmlNode val4 = item2;
					((XmlNode)val4.Attributes[text]).Value = "";
					string value = ((XmlNode)val4.Attributes[text2]).Value;
					if (string.IsNullOrEmpty(value))
					{
						continue;
					}
					XmlNodeList val5 = val2.SelectNodes(text4 + "[@CellID='" + value + "']");
					if (val5 != null && val5.Count == 1)
					{
						string value2 = ((XmlNode)val5[0].Attributes[text]).Value;
						if (!string.IsNullOrEmpty(value2))
						{
							((XmlNode)val5[0].Attributes[text]).Value = "";
							dataChange = true;
						}
					}
				}
				XmlNodeList val6 = val2.SelectNodes(text5 + "[@Controltype='" + controltype + "']");
				foreach (XmlNode item3 in val6)
				{
					XmlNode val4 = item3;
					((XmlNode)val4.Attributes[text]).Value = "";
					string value = ((XmlNode)val4.Attributes[text2]).Value;
					if (string.IsNullOrEmpty(value))
					{
						continue;
					}
					XmlNodeList val5 = val2.SelectNodes(text5 + "[@CellID='" + value + "']");
					if (val5 != null && val5.Count == 1)
					{
						string value2 = ((XmlNode)val5[0].Attributes[text]).Value;
						if (!string.IsNullOrEmpty(value2))
						{
							((XmlNode)val5[0].Attributes[text]).Value = "";
							dataChange = true;
						}
					}
				}
			}
			return doc;
		}
		catch (Exception ex)
		{
			HTKKFile.LogRun(MethodBase.GetCurrentMethod().Name, ex.Message);
			return null;
		}
	}

	public static XmlDocument ThayDoiMaTenDiaBanXaStatic(XmlDocument doc, List<MapDiaBan> mapDiaBan, List<DiaBanXa> diaBan, string controltype, ref bool dataChange)
	{
		//IL_0040: Unknown result type (might be due to invalid IL or missing references)
		//IL_0047: Expected O, but got Unknown
		//IL_0078: Unknown result type (might be due to invalid IL or missing references)
		//IL_007f: Expected O, but got Unknown
		//IL_0273: Unknown result type (might be due to invalid IL or missing references)
		//IL_027a: Expected O, but got Unknown
		try
		{
			string text = "Value";
			string text2 = "SelectedValue";
			string text3 = "/Sections/Section";
			string text4 = ".//Cells/Cell";
			string text5 = ".//Header/Cell";
			XmlNodeList val = ((XmlNode)doc).SelectNodes(text3);
			foreach (XmlNode item in val)
			{
				XmlNode val2 = item;
				XmlNodeList val3 = val2.SelectNodes(text4 + "[@Controltype='" + controltype + "']");
				foreach (XmlNode item2 in val3)
				{
					XmlNode val4 = item2;
					string value = ((XmlNode)val4.Attributes[text]).Value;
					string value2 = ((XmlNode)val4.Attributes[text2]).Value;
					if (string.IsNullOrEmpty(value2))
					{
						continue;
					}
					XmlNodeList val5 = val2.SelectNodes(text4 + "[@CellID='" + value2 + "']");
					if (val5 == null || val5.Count != 1)
					{
						continue;
					}
					string maCu = ((XmlNode)val5[0].Attributes[text]).Value;
					if (string.IsNullOrEmpty(maCu))
					{
						continue;
					}
					List<MapDiaBan> list = mapDiaBan.FindAll((MapDiaBan x) => x.MaCu.Contains(maCu));
					if (list != null && list.Count == 1)
					{
						string maMoi = list[0].MaMoi;
						DiaBanXa diaBanXa = diaBan.Find((DiaBanXa x) => x.MaXa == maMoi);
						if (diaBanXa != null && !(((XmlNode)val5[0].Attributes[text]).Value == list[0].MaMoi))
						{
							((XmlNode)val5[0].Attributes[text]).Value = list[0].MaMoi;
							((XmlNode)val4.Attributes[text]).Value = diaBanXa.TenXa;
							dataChange = true;
						}
					}
				}
				XmlNodeList val6 = val2.SelectNodes(text5 + "[@Controltype='" + controltype + "']");
				foreach (XmlNode item3 in val6)
				{
					XmlNode val4 = item3;
					string value = ((XmlNode)val4.Attributes[text]).Value;
					string value2 = ((XmlNode)val4.Attributes[text2]).Value;
					if (string.IsNullOrEmpty(value2))
					{
						continue;
					}
					XmlNodeList val5 = val2.SelectNodes(text5 + "[@CellID='" + value2 + "']");
					if (val5 == null || val5.Count != 1)
					{
						continue;
					}
					string maCu2 = ((XmlNode)val5[0].Attributes[text]).Value;
					if (string.IsNullOrEmpty(maCu2))
					{
						continue;
					}
					List<MapDiaBan> list = mapDiaBan.FindAll((MapDiaBan x) => x.MaCu.Contains(maCu2));
					if (list != null && list.Count == 1)
					{
						string maMoi2 = list[0].MaMoi;
						DiaBanXa diaBanXa = diaBan.Find((DiaBanXa x) => x.MaXa == maMoi2);
						if (diaBanXa != null && !(((XmlNode)val5[0].Attributes[text]).Value == list[0].MaMoi))
						{
							((XmlNode)val5[0].Attributes[text]).Value = list[0].MaMoi;
							((XmlNode)val4.Attributes[text]).Value = diaBanXa.TenXa;
							dataChange = true;
						}
					}
				}
			}
			return doc;
		}
		catch (Exception ex)
		{
			HTKKFile.LogRun(MethodBase.GetCurrentMethod().Name, ex.Message);
			return null;
		}
	}

	public static XmlDocument ThayDoiMaTenDiaBanTinhStatic(XmlDocument doc, List<MapDiaBan> mapDiaBan, Dictionary<string, string> diaBanMoi, string controltype, ref bool dataChange)
	{
		//IL_0040: Unknown result type (might be due to invalid IL or missing references)
		//IL_0047: Expected O, but got Unknown
		//IL_0078: Unknown result type (might be due to invalid IL or missing references)
		//IL_007f: Expected O, but got Unknown
		//IL_025f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0266: Expected O, but got Unknown
		try
		{
			string text = "Value";
			string text2 = "SelectedValue";
			string text3 = "/Sections/Section";
			string text4 = ".//Cells/Cell";
			string text5 = ".//Header/Cell";
			XmlNodeList val = ((XmlNode)doc).SelectNodes(text3);
			foreach (XmlNode item in val)
			{
				XmlNode val2 = item;
				XmlNodeList val3 = val2.SelectNodes(text4 + "[@Controltype='" + controltype + "']");
				foreach (XmlNode item2 in val3)
				{
					XmlNode val4 = item2;
					string value = ((XmlNode)val4.Attributes[text]).Value;
					string value2 = ((XmlNode)val4.Attributes[text2]).Value;
					if (string.IsNullOrEmpty(value2))
					{
						continue;
					}
					XmlNodeList val5 = val2.SelectNodes(text4 + "[@CellID='" + value2 + "']");
					if (val5 == null || val5.Count != 1)
					{
						continue;
					}
					string maCu = ((XmlNode)val5[0].Attributes[text]).Value;
					if (string.IsNullOrEmpty(maCu))
					{
						continue;
					}
					List<MapDiaBan> list = mapDiaBan.FindAll((MapDiaBan x) => x.MaCu.Contains(maCu));
					if (list != null && list.Count == 1)
					{
						string maMoi = list[0].MaMoi;
						KeyValuePair<string, string> keyValuePair = diaBanMoi.FirstOrDefault((KeyValuePair<string, string> x) => x.Value == maMoi);
						if (!(((XmlNode)val5[0].Attributes[text]).Value == maMoi))
						{
							((XmlNode)val5[0].Attributes[text]).Value = maMoi;
							((XmlNode)val4.Attributes[text]).Value = keyValuePair.Key;
							dataChange = true;
						}
					}
				}
				XmlNodeList val6 = val2.SelectNodes(text5 + "[@Controltype='" + controltype + "']");
				foreach (XmlNode item3 in val6)
				{
					XmlNode val4 = item3;
					string value = ((XmlNode)val4.Attributes[text]).Value;
					string value2 = ((XmlNode)val4.Attributes[text2]).Value;
					if (string.IsNullOrEmpty(value2))
					{
						continue;
					}
					XmlNodeList val5 = val2.SelectNodes(text5 + "[@CellID='" + value2 + "']");
					if (val5 == null || val5.Count != 1)
					{
						continue;
					}
					string maCu2 = ((XmlNode)val5[0].Attributes[text]).Value;
					if (string.IsNullOrEmpty(maCu2))
					{
						continue;
					}
					List<MapDiaBan> list = mapDiaBan.FindAll((MapDiaBan x) => x.MaCu.Contains(maCu2));
					if (list != null && list.Count == 1)
					{
						string maMoi2 = list[0].MaMoi;
						KeyValuePair<string, string> keyValuePair = diaBanMoi.FirstOrDefault((KeyValuePair<string, string> x) => x.Value == maMoi2);
						if (!(((XmlNode)val5[0].Attributes[text]).Value == maMoi2))
						{
							((XmlNode)val5[0].Attributes[text]).Value = maMoi2;
							((XmlNode)val4.Attributes[text]).Value = keyValuePair.Key;
							dataChange = true;
						}
					}
				}
			}
			return doc;
		}
		catch (Exception ex)
		{
			HTKKFile.LogRun(MethodBase.GetCurrentMethod().Name, ex.Message);
			return null;
		}
	}

	public static List<MapDiaBan> getMapDiaBan(string pathFileMap)
	{
		string text = HTKKConfig.INTERFACE_INI + "\\" + pathFileMap;
		try
		{
			XDocument val = XDocument.Load(text);
			return (from item in ((XContainer)val).Descendants(XName.op_Implicit("Item"))
				select new MapDiaBan
				{
					MaCu = item.Attribute(XName.op_Implicit("MaCu")).Value,
					MaMoi = item.Attribute(XName.op_Implicit("MaMoi")).Value
				}).ToList();
		}
		catch (Exception)
		{
			HTKKFile.LogRun(MethodBase.GetCurrentMethod().Name, "Loi conver TK: " + text);
			return null;
		}
	}

	public static void FixRowStatic(XmlDocument xmlTemplate, ref XmlDocument xmlDataFile)
	{
		//IL_01aa: Unknown result type (might be due to invalid IL or missing references)
		//IL_01b1: Expected O, but got Unknown
		//IL_01cd: Unknown result type (might be due to invalid IL or missing references)
		//IL_01d4: Expected O, but got Unknown
		if (xmlTemplate == null || xmlDataFile == null)
		{
			return;
		}
		XmlNodeList val = ((XmlNode)xmlDataFile).SelectNodes("/Sections/Section");
		for (int i = 0; i < val.Count - 1; i++)
		{
			XmlNode val2 = val[i];
			if (val2.Attributes["Dynamic"] == null || !(((XmlNode)val2.Attributes["Dynamic"]).Value == "1"))
			{
				continue;
			}
			XmlNodeList val3 = val2.SelectNodes("RowData");
			XmlNodeList val4 = val2.SelectNodes("Location/DataLocationTemplateRowStart");
			int num = Convert.ToInt32(val4[0].InnerText);
			XmlNodeList val5 = val2.SelectNodes("Location/DataLocationReportRowStart");
			int num2 = Convert.ToInt32(val5[0].InnerText);
			int num3 = 1;
			if (val3.Count > 0)
			{
				num3 = ((val3[0].Attributes["RowDisplay"] == null) ? 1 : Convert.ToInt32(((XmlNode)val3[0].Attributes["RowDisplay"]).Value));
			}
			int num4 = val3.Count - num3;
			if (num4 <= 0)
			{
				continue;
			}
			for (int j = i + 1; j < val.Count; j++)
			{
				XmlNode val6 = val[j];
				if (val6.Attributes["Dynamic"] == null || !(((XmlNode)val6.Attributes["Dynamic"]).Value == "0"))
				{
					continue;
				}
				foreach (XmlNode childNode in val6.ChildNodes)
				{
					XmlNode val7 = childNode;
					foreach (XmlNode childNode2 in val7.ChildNodes)
					{
						XmlNode val8 = childNode2;
						string path = ((val8.Attributes["Path"] != null) ? ((XmlNode)val8.Attributes["Path"]).Value : "");
						if (string.IsNullOrEmpty(path))
						{
							continue;
						}
						XmlNodeList val9 = ((XmlNode)xmlTemplate).SelectNodes("/Sections/Section/Cells/Cell");
						if (val9 == null || val9.Count == 0)
						{
							continue;
						}
						XmlNode val10 = ((IEnumerable)val9).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode x) => x.Attributes["Path"] != null && ((XmlNode)x.Attributes["Path"]).Value == path));
						if (val10 == null)
						{
							continue;
						}
						string value = ((XmlNode)val10.Attributes["CellID"]).Value;
						string[] array = value.Split(new char[1] { '_' });
						if (array.Count() > 1)
						{
							int num5 = Convert.ToInt32(array[1]);
							int num6 = num5 + num4;
							if (num5 < num)
							{
								num6 = num5;
							}
							string value2 = array[0] + "_" + num6;
							((XmlNode)val10.Attributes["CellID"]).Value = value2;
							((XmlNode)val8.Attributes["CellID"]).Value = value2;
						}
						string value3 = ((XmlNode)val10.Attributes["CellID2"]).Value;
						string[] array2 = value3.Split(new char[1] { '_' });
						if (array2.Count() > 1)
						{
							int num5 = Convert.ToInt32(array2[1]);
							int num6 = num5 + num4;
							if (num5 < num2)
							{
								num6 = num5;
							}
							string value2 = array2[0] + "_" + num6;
							((XmlNode)val10.Attributes["CellID2"]).Value = value2;
							((XmlNode)val8.Attributes["CellID2"]).Value = value2;
						}
					}
				}
			}
		}
	}

	public static int TinhOffset(XmlNode itemNew, XmlNode itemOld)
	{
		try
		{
			int result = 0;
			if (itemOld != null && itemNew != null)
			{
				string text = ((itemNew.Attributes["CellID"] == null) ? "" : ((XmlNode)itemNew.Attributes["CellID"]).Value);
				string text2 = ((itemOld.Attributes["CellID"] == null) ? "" : ((XmlNode)itemOld.Attributes["CellID"]).Value);
				if (!string.IsNullOrEmpty(text) && !string.IsNullOrEmpty(text2))
				{
					string[] array = text.Split(new char[1] { '_' });
					string[] array2 = text2.Split(new char[1] { '_' });
					result = Convert.ToInt32(array2[1]) - Convert.ToInt32(array[1]);
				}
			}
			return result;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return 0;
		}
	}

	public static string SetOffsetForCellIdOld(string attrValue, string attrName, int offset)
	{
		try
		{
			if (attrName.IndexOf("CellID") >= 0)
			{
				string[] array = attrValue.Split(new char[1] { '_' });
				int num = Convert.ToInt32(array[1]);
				num += offset;
				return array[0] + "_" + num;
			}
			return attrValue;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return attrValue;
		}
	}

	public static void RenameDatafileIfError()
	{
		//IL_00a7: Unknown result type (might be due to invalid IL or missing references)
		//IL_00ad: Expected O, but got Unknown
		//IL_02dd: Unknown result type (might be due to invalid IL or missing references)
		//IL_02e4: Expected O, but got Unknown
		//IL_0396: Unknown result type (might be due to invalid IL or missing references)
		//IL_039d: Expected O, but got Unknown
		//IL_01e6: Unknown result type (might be due to invalid IL or missing references)
		//IL_01ed: Expected O, but got Unknown
		try
		{
			string path = HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst + "\\BK";
			if (!Directory.Exists(path))
			{
				Directory.CreateDirectory(path);
			}
			DirectoryInfo directoryInfo = new DirectoryInfo(HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst);
			FileInfo[] files = directoryInfo.GetFiles("*05_TNDN_TT80*.xml");
			foreach (FileInfo fileInfo in files)
			{
				if (fileInfo.Name.IndexOf("CN1") >= 0 || fileInfo.Name.IndexOf("CN2") >= 0)
				{
					continue;
				}
				XmlDocument val = new XmlDocument();
				val.Load(fileInfo.FullName);
				XmlNodeList val2 = ((XmlNode)val).SelectNodes("/Sections/Section/Cells/Cell");
				if (val2 != null && val2.Count == 0)
				{
					continue;
				}
				XmlNode val3 = ((IEnumerable)val2).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode x) => x.Attributes["Path"] != null && ((XmlNode)x.Attributes["Path"]).Value == "path_AA_22"));
				if (val3 == null)
				{
					continue;
				}
				string text = "CN" + ((XmlNode)val3.Attributes["Value"]).Value;
				string[] array = fileInfo.Name.Split(new char[1] { '_' });
				string text2 = "";
				for (int num = 0; num < array.Length; num++)
				{
					if (num == array.Length - 2)
					{
						text2 = text2 + text + "_";
					}
					text2 = text2 + array[num] + "_";
				}
				text2 = text2.Substring(0, text2.Length - 1);
				XmlTextWriter val4 = new XmlTextWriter(fileInfo.DirectoryName + "\\" + text2, Encoding.UTF8);
				try
				{
					val4.Formatting = (Formatting)1;
					val.Save((XmlWriter)(object)val4);
					try
					{
						string destDirName = fileInfo.DirectoryName + "\\BK\\" + fileInfo.Name;
						Directory.Move(fileInfo.FullName, destDirName);
					}
					catch
					{
					}
				}
				finally
				{
					((IDisposable)val4)?.Dispose();
				}
				if (fileInfo.Name.IndexOf("bs") < 0)
				{
					continue;
				}
				string text3 = fileInfo.Name.Substring(0, 3);
				FileInfo[] files2 = directoryInfo.GetFiles(fileInfo.Name.Substring(0, 3) + "*KHBS*05_TNDN_TT80*.xml");
				foreach (FileInfo fileInfo2 in files2)
				{
					if (fileInfo2.Name.IndexOf("CN1") >= 0 || fileInfo2.Name.IndexOf("CN2") >= 0)
					{
						continue;
					}
					XmlDocument val5 = new XmlDocument();
					val5.Load(fileInfo2.FullName);
					string[] array2 = fileInfo2.Name.Split(new char[1] { '_' });
					string text4 = "";
					for (int num = 0; num < array2.Length; num++)
					{
						if (num == array2.Length - 2)
						{
							text4 = text4 + text + "_";
						}
						text4 = text4 + array2[num] + "_";
					}
					text4 = text4.Substring(0, text4.Length - 1);
					val4 = new XmlTextWriter(fileInfo2.DirectoryName + "\\" + text4, Encoding.UTF8);
					try
					{
						val4.Formatting = (Formatting)1;
						val5.Save((XmlWriter)(object)val4);
						try
						{
							string destDirName = fileInfo2.DirectoryName + "\\BK\\" + fileInfo2.Name;
							Directory.Move(fileInfo2.FullName, destDirName);
						}
						catch
						{
						}
					}
					finally
					{
						((IDisposable)val4)?.Dispose();
					}
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void UpdateRestartConvertData(string dataFileName, int restartFrom, int restartTo)
	{
		//IL_0044: Unknown result type (might be due to invalid IL or missing references)
		//IL_004a: Expected O, but got Unknown
		//IL_00d7: Unknown result type (might be due to invalid IL or missing references)
		//IL_00de: Expected O, but got Unknown
		try
		{
			if (restartFrom == 0)
			{
				return;
			}
			string text = HTKKConfig.INTERFACE_INI + "\\ConvertVersion.xml";
			XmlDocument val = new XmlDocument();
			if (!File.Exists(text))
			{
				return;
			}
			val.Load(text);
			XmlNodeList elementsByTagName = val.GetElementsByTagName("Item");
			XmlNode val2 = ((IEnumerable)elementsByTagName).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode x) => x.Attributes["DataFileName"] != null && ((XmlNode)x.Attributes["DataFileName"]).Value.Trim() == dataFileName && x.Attributes["Restart"] != null && ((XmlNode)x.Attributes["Restart"]).Value.Trim() == restartFrom.ToString()));
			if (val2 != null && val2.Attributes["Restart"] != null)
			{
				((XmlNode)val2.Attributes["Restart"]).Value = restartTo.ToString();
			}
			XmlTextWriter val3 = new XmlTextWriter(text, Encoding.UTF8);
			try
			{
				val3.Formatting = (Formatting)1;
				val.Save((XmlWriter)(object)val3);
			}
			finally
			{
				((IDisposable)val3)?.Dispose();
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static XmlNode InsertRowsValue(XmlNode elemNew, XmlNode elemOld, List<int> listIndexAddNew)
	{
		try
		{
			if (listIndexAddNew.Count == 0)
			{
				return elemOld;
			}
			XmlNode val = ((IEnumerable)elemNew.ChildNodes).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode x) => x.Name == "RowData"));
			List<XmlNode> list = (from XmlNode x in (IEnumerable)elemOld.ChildNodes
				where x.Name == "RowData"
				select x).ToList();
			foreach (XmlNode item in list)
			{
				List<string> list2 = GeneralFunctions.SplitString2(val.InnerText, "###");
				List<string> list3 = GeneralFunctions.SplitString2(item.InnerText, "###");
				List<string> list4 = new List<string>();
				int num = 0;
				for (int num2 = 0; num2 < list2.Count; num2++)
				{
					if (!listIndexAddNew.Contains(num2))
					{
						list4.Add(list3[num]);
						num++;
					}
					else
					{
						list4.Add(list2[num2]);
					}
				}
				item.InnerText = GeneralFunctions.ListToString(list4, "###");
			}
			return elemOld;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public static List<int> GetViTriThayDoi(XmlNode elemNew, XmlNode elemOld)
	{
		//IL_0024: Unknown result type (might be due to invalid IL or missing references)
		//IL_002a: Expected O, but got Unknown
		//IL_006b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0072: Expected O, but got Unknown
		List<int> list = new List<int>();
		try
		{
			int num = 0;
			foreach (XmlNode childNode in elemNew.ChildNodes)
			{
				XmlNode val = childNode;
				if (val != null && val.Name == "RowInfo")
				{
					int num2 = 0;
					foreach (XmlNode childNode2 in val.ChildNodes)
					{
						XmlNode val2 = childNode2;
						string cellIdNew = ((val2.Attributes["CellID"] != null) ? ((XmlNode)val2.Attributes["CellID"]).Value.Trim() : "");
						if (!string.IsNullOrEmpty(cellIdNew))
						{
							string pathNew = ((val2.Attributes["Path"] != null) ? ((XmlNode)val2.Attributes["Path"]).Value.Trim() : "");
							XmlNode val3 = ((IEnumerable)elemOld.ChildNodes[num]).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode x) => x.Attributes["CellID"] != null && ((XmlNode)x.Attributes["CellID"]).Value.Trim() == cellIdNew));
							XmlNode val4 = ((IEnumerable)elemOld.ChildNodes[num]).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode x) => x.Attributes["Path"] != null && ((XmlNode)x.Attributes["Path"]).Value.Trim() == pathNew));
							if ((val3 == null || val3.Attributes["CellID"] == null) && val4 == null)
							{
								list.Add(num2);
							}
							if (val4 == null)
							{
								list.Add(num2);
							}
						}
						num2++;
					}
				}
				num++;
			}
			return list;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return list;
		}
	}

	public static void ConverDataVersion01DNHT(string dataFileName, string versionXmlOld, string versionXml)
	{
		//IL_01c2: Unknown result type (might be due to invalid IL or missing references)
		//IL_01c9: Expected O, but got Unknown
		//IL_01f3: Unknown result type (might be due to invalid IL or missing references)
		//IL_01fa: Expected O, but got Unknown
		//IL_07a6: Unknown result type (might be due to invalid IL or missing references)
		//IL_07ad: Expected O, but got Unknown
		string text = string.Empty;
		string empty = string.Empty;
		string empty2 = string.Empty;
		string empty3 = string.Empty;
		string empty4 = string.Empty;
		string empty5 = string.Empty;
		string empty6 = string.Empty;
		try
		{
			string[] array = dataFileName.Split(new char[1] { '~' });
			DirectoryInfo directoryInfo = new DirectoryInfo(HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst);
			string path = HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst + "\\BK";
			if (!Directory.Exists(path))
			{
				Directory.CreateDirectory(path);
			}
			empty6 = HTKKConfig.INTERFACE_TEMPLATES + "\\xml\\" + dataFileName + "_" + versionXml + ".xml";
			string[] array2 = array;
			foreach (string text2 in array2)
			{
				FileInfo[] files = directoryInfo.GetFiles("*" + text2 + "*.xml");
				foreach (FileInfo fileInfo in files)
				{
					if (fileInfo.Name.Contains("KHBS"))
					{
						continue;
					}
					empty = fileInfo.Name;
					text = fileInfo.FullName;
					empty2 = empty.Replace(text2 + "_", "");
					empty3 = empty.Replace("_" + versionXmlOld + ".xml", "_" + versionXml + ".xml");
					empty4 = fileInfo.DirectoryName + "\\BK\\" + empty;
					if (File.Exists(fileInfo.DirectoryName + "\\" + empty3))
					{
						continue;
					}
					XmlDocument val = new XmlDocument();
					if (!File.Exists(empty6))
					{
						return;
					}
					val.Load(empty6);
					XmlNodeList val2 = ((XmlNode)val).SelectNodes("/Sections/Section");
					XmlDocument val3 = new XmlDocument();
					val3.Load(text);
					XmlNodeList val4 = ((XmlNode)val3).SelectNodes("/Sections/Section");
					XmlNode val5 = val4.Item(1);
					if (val5.ChildNodes[0].ChildNodes.Count == 23)
					{
						XmlNode val6 = val2.Item(1).ChildNodes[0];
						XmlNodeList childNodes = val6.ChildNodes;
						for (int k = 0; k <= 9; k++)
						{
							XmlNode val7 = childNodes.Item(k + 23);
							XmlNode val8 = val7.Clone();
							XmlNode val9 = val3.ImportNode(val8, true);
							val5.ChildNodes[0].InsertAfter(val9, val5.ChildNodes[0].LastChild);
						}
						val5 = val4.Item(2);
						val6 = val5.ChildNodes[0];
						childNodes = val6.ChildNodes;
						XmlNode val10 = null;
						for (int k = 0; k <= 11; k++)
						{
							val10 = childNodes.Item(k);
							string value = ((XmlNode)val10.Attributes["CellID2"]).Value;
							if (!value.Equals(""))
							{
								string[] array3 = value.Split(new char[1] { '_' });
								((XmlNode)val10.Attributes["CellID2"]).Value = array3[0] + "_" + (ConvertHelpers.ConvertToInt(array3[1]) + 23);
							}
						}
						val5 = val4.Item(3);
						val6 = val5.ChildNodes[0];
						childNodes = val6.ChildNodes;
						val10 = childNodes.Item(1);
						string innerXml = val10.InnerXml;
						if (!innerXml.Equals(""))
						{
							val10.InnerXml = ConvertHelpers.ConvertToString(ConvertHelpers.ConvertToInt(innerXml) + 23);
						}
						val10 = childNodes.Item(3);
						innerXml = val10.InnerXml;
						if (!innerXml.Equals(""))
						{
							val10.InnerXml = ConvertHelpers.ConvertToString(ConvertHelpers.ConvertToInt(innerXml) + 23);
						}
						val5 = val4.Item(4);
						val6 = val5.ChildNodes[0];
						childNodes = val6.ChildNodes;
						for (int k = 0; k <= 1; k++)
						{
							val10 = childNodes.Item(k);
							string value = ((XmlNode)val10.Attributes["CellID2"]).Value;
							if (!value.Equals(""))
							{
								string[] array3 = value.Split(new char[1] { '_' });
								((XmlNode)val10.Attributes["CellID2"]).Value = array3[0] + "_" + (ConvertHelpers.ConvertToInt(array3[1]) + 23);
							}
						}
						val5 = val4.Item(5);
						val6 = val5.ChildNodes[0];
						childNodes = val6.ChildNodes;
						for (int k = 0; k <= 11; k++)
						{
							val10 = childNodes.Item(k);
							string value = ((XmlNode)val10.Attributes["CellID2"]).Value;
							if (!value.Equals(""))
							{
								string[] array3 = value.Split(new char[1] { '_' });
								((XmlNode)val10.Attributes["CellID2"]).Value = array3[0] + "_" + (ConvertHelpers.ConvertToInt(array3[1]) + 23);
							}
						}
						val5 = val4.Item(6);
						val6 = val5.ChildNodes[0];
						childNodes = val6.ChildNodes;
						val10 = childNodes.Item(1);
						innerXml = val10.InnerXml;
						if (!innerXml.Equals(""))
						{
							val10.InnerXml = ConvertHelpers.ConvertToString(ConvertHelpers.ConvertToInt(innerXml) + 23);
						}
						val10 = childNodes.Item(3);
						innerXml = val10.InnerXml;
						if (!innerXml.Equals(""))
						{
							val10.InnerXml = ConvertHelpers.ConvertToString(ConvertHelpers.ConvertToInt(innerXml) + 23);
						}
						val5 = val4.Item(7);
						val6 = val5.ChildNodes[0];
						childNodes = val6.ChildNodes;
						for (int k = 0; k <= 4; k++)
						{
							val10 = childNodes.Item(k);
							string value = ((XmlNode)val10.Attributes["CellID2"]).Value;
							if (!value.Equals(""))
							{
								string[] array3 = value.Split(new char[1] { '_' });
								((XmlNode)val10.Attributes["CellID2"]).Value = array3[0] + "_" + (ConvertHelpers.ConvertToInt(array3[1]) + 23);
							}
						}
					}
					val3.DocumentElement.SetAttribute("Version", WinSettings.VersionApp);
					string text3 = HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst + "\\" + empty3;
					XmlTextWriter val11 = new XmlTextWriter(text3, Encoding.UTF8);
					try
					{
						val11.Formatting = (Formatting)1;
						val3.Save((XmlWriter)(object)val11);
					}
					finally
					{
						((IDisposable)val11)?.Dispose();
					}
					Directory.Move(text, empty4);
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			HTKKFile.LogRun(MethodBase.GetCurrentMethod().Name, "Loi conver TK: " + text);
		}
	}

	public static void ConverDataVersion04GTGT(string dataFileName, string versionXmlOld, string versionXml)
	{
		//IL_01c2: Unknown result type (might be due to invalid IL or missing references)
		//IL_01c9: Expected O, but got Unknown
		//IL_01f3: Unknown result type (might be due to invalid IL or missing references)
		//IL_01fa: Expected O, but got Unknown
		//IL_0303: Unknown result type (might be due to invalid IL or missing references)
		//IL_030a: Expected O, but got Unknown
		string text = string.Empty;
		string empty = string.Empty;
		string empty2 = string.Empty;
		string empty3 = string.Empty;
		string empty4 = string.Empty;
		string empty5 = string.Empty;
		string empty6 = string.Empty;
		try
		{
			string[] array = dataFileName.Split(new char[1] { '~' });
			DirectoryInfo directoryInfo = new DirectoryInfo(HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst);
			string path = HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst + "\\BK";
			if (!Directory.Exists(path))
			{
				Directory.CreateDirectory(path);
			}
			empty6 = HTKKConfig.INTERFACE_TEMPLATES + "\\xml\\" + dataFileName + "_" + versionXml + ".xml";
			string[] array2 = array;
			foreach (string text2 in array2)
			{
				FileInfo[] files = directoryInfo.GetFiles("*" + text2 + "*.xml");
				foreach (FileInfo fileInfo in files)
				{
					if (fileInfo.Name.Contains("KHBS"))
					{
						continue;
					}
					empty = fileInfo.Name;
					text = fileInfo.FullName;
					empty2 = empty.Replace(text2 + "_", "");
					empty3 = empty.Replace("_" + versionXmlOld + ".xml", "_" + versionXml + ".xml");
					empty4 = fileInfo.DirectoryName + "\\BK\\" + empty;
					if (!File.Exists(fileInfo.DirectoryName + "\\" + empty3))
					{
						XmlDocument val = new XmlDocument();
						if (!File.Exists(empty6))
						{
							return;
						}
						val.Load(empty6);
						XmlNodeList val2 = ((XmlNode)val).SelectNodes("/Sections/Section");
						XmlDocument val3 = new XmlDocument();
						val3.Load(text);
						XmlNodeList val4 = ((XmlNode)val3).SelectNodes("/Sections/Section");
						XmlNode val5 = val4.Item(2);
						if (val5.ChildNodes[0].ChildNodes.Count == 9)
						{
							XmlNode val6 = val2.Item(2).ChildNodes[0];
							XmlNodeList childNodes = val6.ChildNodes;
							XmlNode val7 = childNodes.Item(9);
							XmlNode val8 = val7.Clone();
							XmlNode val9 = val3.ImportNode(val8, true);
							val5.ChildNodes[0].InsertAfter(val9, val5.ChildNodes[0].LastChild);
						}
						val3.DocumentElement.SetAttribute("Version", WinSettings.VersionApp);
						string text3 = HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst + "\\" + empty3;
						XmlTextWriter val10 = new XmlTextWriter(text3, Encoding.UTF8);
						try
						{
							val10.Formatting = (Formatting)1;
							val3.Save((XmlWriter)(object)val10);
						}
						finally
						{
							((IDisposable)val10)?.Dispose();
						}
						Directory.Move(text, empty4);
					}
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			HTKKFile.LogRun(MethodBase.GetCurrentMethod().Name, "Loi conver TK: " + text);
		}
	}

	public static void ConverDataVersion01GTGT(string dataFileName, string versionXmlOld, string versionXml)
	{
		//IL_01c2: Unknown result type (might be due to invalid IL or missing references)
		//IL_01c9: Expected O, but got Unknown
		//IL_01f3: Unknown result type (might be due to invalid IL or missing references)
		//IL_01fa: Expected O, but got Unknown
		//IL_042e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0435: Expected O, but got Unknown
		string text = string.Empty;
		string empty = string.Empty;
		string empty2 = string.Empty;
		string empty3 = string.Empty;
		string empty4 = string.Empty;
		string empty5 = string.Empty;
		string empty6 = string.Empty;
		try
		{
			string[] array = dataFileName.Split(new char[1] { '~' });
			DirectoryInfo directoryInfo = new DirectoryInfo(HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst);
			string path = HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst + "\\BK";
			if (!Directory.Exists(path))
			{
				Directory.CreateDirectory(path);
			}
			empty6 = HTKKConfig.INTERFACE_TEMPLATES + "\\xml\\" + dataFileName + "_" + versionXml + ".xml";
			string[] array2 = array;
			foreach (string text2 in array2)
			{
				FileInfo[] files = directoryInfo.GetFiles("*" + text2 + "*.xml");
				foreach (FileInfo fileInfo in files)
				{
					if (fileInfo.Name.Contains("KHBS"))
					{
						continue;
					}
					empty = fileInfo.Name;
					text = fileInfo.FullName;
					empty2 = empty.Replace(text2 + "_", "");
					empty3 = empty.Replace("_" + versionXmlOld + ".xml", "_" + versionXml + ".xml");
					empty4 = fileInfo.DirectoryName + "\\BK\\" + empty;
					if (!File.Exists(fileInfo.DirectoryName + "\\" + empty3))
					{
						XmlDocument val = new XmlDocument();
						if (!File.Exists(empty6))
						{
							return;
						}
						val.Load(empty6);
						XmlNodeList val2 = ((XmlNode)val).SelectNodes("/Sections/Section");
						XmlDocument val3 = new XmlDocument();
						val3.Load(text);
						XmlNodeList val4 = ((XmlNode)val3).SelectNodes("/Sections/Section");
						string valueByAttributes = HTKKFile.GetValueByAttributes(text, "/Sections/Section/Cells/Cell[@Path='HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/tieuMucHachToan']");
						string value = "";
						string value2 = "";
						XMLMessage xMLMessage = new XMLMessage("../Project/Message.xml");
						if (valueByAttributes.Equals("1701"))
						{
							value = "07";
							value2 = xMLMessage.GetMessage("0519");
						}
						else if (valueByAttributes.Equals("1704"))
						{
							value = "08";
							value2 = xMLMessage.GetMessage("0520");
						}
						else if (valueByAttributes.Equals("1705"))
						{
							value = "09";
							value2 = xMLMessage.GetMessage("0521");
						}
						XmlNode val5 = val4.Item(2);
						if (val5.ChildNodes[0].ChildNodes.Count == 12)
						{
							XmlNode val6 = val2.Item(2).ChildNodes[0];
							XmlNodeList childNodes = val6.ChildNodes;
							XmlNode val7 = childNodes.Item(12);
							XmlNode val8 = val7.Clone();
							XmlNode val9 = val3.ImportNode(val8, true);
							((XmlNode)val9.Attributes["Value"]).Value = value2;
							val5.ChildNodes[0].InsertAfter(val9, val5.ChildNodes[0].LastChild);
							val7 = childNodes.Item(13);
							val8 = val7.Clone();
							val9 = val3.ImportNode(val8, true);
							((XmlNode)val9.Attributes["Value"]).Value = value;
							val5.ChildNodes[0].InsertAfter(val9, val5.ChildNodes[0].LastChild);
						}
						val3.DocumentElement.SetAttribute("Version", WinSettings.VersionApp);
						string text3 = HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst + "\\" + empty3;
						XmlTextWriter val10 = new XmlTextWriter(text3, Encoding.UTF8);
						try
						{
							val10.Formatting = (Formatting)1;
							val3.Save((XmlWriter)(object)val10);
						}
						finally
						{
							((IDisposable)val10)?.Dispose();
						}
						Directory.Move(text, empty4);
					}
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			HTKKFile.LogRun(MethodBase.GetCurrentMethod().Name, "Loi conver TK: " + text);
		}
	}

	public static void ConverDataVersion05_1_QTT(string dataFileName, string versionXmlOld, string versionXml)
	{
		//IL_01c2: Unknown result type (might be due to invalid IL or missing references)
		//IL_01c9: Expected O, but got Unknown
		//IL_01f3: Unknown result type (might be due to invalid IL or missing references)
		//IL_01fa: Expected O, but got Unknown
		//IL_02f8: Unknown result type (might be due to invalid IL or missing references)
		//IL_02ff: Expected O, but got Unknown
		string text = string.Empty;
		string empty = string.Empty;
		string empty2 = string.Empty;
		string empty3 = string.Empty;
		string empty4 = string.Empty;
		string empty5 = string.Empty;
		string empty6 = string.Empty;
		try
		{
			string[] array = dataFileName.Split(new char[1] { '~' });
			DirectoryInfo directoryInfo = new DirectoryInfo(HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst);
			string path = HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst + "\\BK";
			if (!Directory.Exists(path))
			{
				Directory.CreateDirectory(path);
			}
			empty6 = HTKKConfig.INTERFACE_TEMPLATES + "\\xml\\" + dataFileName + "_" + versionXml + ".xml";
			string[] array2 = array;
			foreach (string text2 in array2)
			{
				FileInfo[] files = directoryInfo.GetFiles("*" + text2 + "*.xml");
				foreach (FileInfo fileInfo in files)
				{
					if (fileInfo.Name.Contains("KHBS"))
					{
						continue;
					}
					empty = fileInfo.Name;
					text = fileInfo.FullName;
					empty2 = empty.Replace(text2 + "_", "");
					empty3 = empty.Replace("_" + versionXmlOld + ".xml", "_" + versionXml + ".xml");
					empty4 = fileInfo.DirectoryName + "\\BK\\" + empty;
					if (!File.Exists(fileInfo.DirectoryName + "\\" + empty3))
					{
						XmlDocument val = new XmlDocument();
						if (!File.Exists(empty6))
						{
							return;
						}
						val.Load(empty6);
						XmlNodeList val2 = ((XmlNode)val).SelectNodes("/Sections/Section");
						XmlDocument val3 = new XmlDocument();
						val3.Load(text);
						XmlNodeList val4 = ((XmlNode)val3).SelectNodes("/Sections/Section");
						XmlNode val5 = val4.Item(2);
						if (val5.ChildNodes[0].ChildNodes.Count == 14)
						{
							XmlNode val6 = val2.Item(2).ChildNodes[0];
							XmlNode lastChild = val6.LastChild;
							XmlNode val7 = lastChild.Clone();
							XmlNode val8 = val3.ImportNode(val7, true);
							val5.ChildNodes[0].InsertAfter(val8, val5.ChildNodes[0].LastChild);
						}
						val3.DocumentElement.SetAttribute("Version", WinSettings.VersionApp);
						string text3 = HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst + "\\" + empty3;
						XmlTextWriter val9 = new XmlTextWriter(text3, Encoding.UTF8);
						try
						{
							val9.Formatting = (Formatting)1;
							val3.Save((XmlWriter)(object)val9);
						}
						finally
						{
							((IDisposable)val9)?.Dispose();
						}
						Directory.Move(text, empty4);
					}
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			HTKKFile.LogRun(MethodBase.GetCurrentMethod().Name, "Loi conver TK: " + text);
		}
	}

	public static void ConverDataVersion200_HN_CDKT(string dataFileName, string versionXmlOld, string versionXml)
	{
		//IL_019e: Unknown result type (might be due to invalid IL or missing references)
		//IL_01a5: Expected O, but got Unknown
		//IL_01cf: Unknown result type (might be due to invalid IL or missing references)
		//IL_01d6: Expected O, but got Unknown
		//IL_04fc: Unknown result type (might be due to invalid IL or missing references)
		//IL_0503: Expected O, but got Unknown
		string text = string.Empty;
		string empty = string.Empty;
		string empty2 = string.Empty;
		string empty3 = string.Empty;
		string empty4 = string.Empty;
		string empty5 = string.Empty;
		string empty6 = string.Empty;
		try
		{
			string[] array = dataFileName.Split(new char[1] { '~' });
			DirectoryInfo directoryInfo = new DirectoryInfo(HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst);
			string path = HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst + "\\BK";
			if (!Directory.Exists(path))
			{
				Directory.CreateDirectory(path);
			}
			empty6 = HTKKConfig.INTERFACE_TEMPLATES + "\\xml\\" + dataFileName + "_" + versionXml + ".xml";
			string[] array2 = array;
			foreach (string text2 in array2)
			{
				FileInfo[] files = directoryInfo.GetFiles(text2 + "*.xml");
				foreach (FileInfo fileInfo in files)
				{
					empty = fileInfo.Name;
					text = fileInfo.FullName;
					empty2 = empty.Replace(text2 + "_", "");
					empty3 = empty.Replace("_" + versionXmlOld + ".xml", "_" + versionXml + ".xml");
					empty4 = fileInfo.DirectoryName + "\\BK\\" + empty;
					if (File.Exists(fileInfo.DirectoryName + "\\" + empty3))
					{
						continue;
					}
					XmlDocument val = new XmlDocument();
					if (!File.Exists(empty6))
					{
						return;
					}
					val.Load(empty6);
					XmlNodeList val2 = ((XmlNode)val).SelectNodes("/Sections/Section");
					XmlDocument val3 = new XmlDocument();
					val3.Load(text);
					XmlNodeList val4 = ((XmlNode)val3).SelectNodes("/Sections/Section");
					XmlNode val5 = null;
					val5 = val4.Item(2);
					if (val5.ChildNodes[0].ChildNodes.Count != 346)
					{
						continue;
					}
					val5 = val4.Item(3);
					if (val5.ChildNodes[0].ChildNodes.Count != 219)
					{
						continue;
					}
					for (int k = 0; k < val2.Count; k++)
					{
						XmlNode val6 = val2.Item(k);
						val5 = val4.Item(k);
						for (int l = 0; l < val6.ChildNodes[0].ChildNodes.Count; l++)
						{
							XmlNodeList childNodes = val6.ChildNodes[0].ChildNodes;
							XmlNode val7 = childNodes.Item(l);
							if (val7.Attributes["Path"] != null)
							{
								for (int m = 0; m < val5.ChildNodes[0].ChildNodes.Count; m++)
								{
									XmlNodeList childNodes2 = val5.ChildNodes[0].ChildNodes;
									XmlNode val8 = childNodes2.Item(m);
									if (val8.Attributes["Path"] != null && ((XmlNode)val7.Attributes["Path"]).Value == ((XmlNode)val8.Attributes["Path"]).Value)
									{
										((XmlNode)val7.Attributes["Value"]).Value = ((XmlNode)val8.Attributes["Value"]).Value;
										break;
									}
								}
								continue;
							}
							for (int m = 0; m < val5.ChildNodes[0].ChildNodes.Count; m++)
							{
								XmlNodeList childNodes2 = val5.ChildNodes[0].ChildNodes;
								XmlNode val8 = childNodes2.Item(m);
								if (((XmlNode)val7.Attributes["CellID"]).Value == ((XmlNode)val8.Attributes["CellID"]).Value)
								{
									((XmlNode)val7.Attributes["Value"]).Value = ((XmlNode)val8.Attributes["Value"]).Value;
									break;
								}
							}
						}
					}
					val.DocumentElement.SetAttribute("Version", WinSettings.VersionApp);
					string text3 = HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst + "\\" + empty3;
					XmlTextWriter val9 = new XmlTextWriter(text3, Encoding.UTF8);
					try
					{
						val9.Formatting = (Formatting)1;
						val.Save((XmlWriter)(object)val9);
					}
					finally
					{
						((IDisposable)val9)?.Dispose();
					}
					Directory.Move(text, empty4);
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			HTKKFile.LogRun(MethodBase.GetCurrentMethod().Name, "Loi conver TK: " + text);
		}
	}

	public static void ConverDataVersion200_HN_HDSXKD(string dataFileName, string versionXmlOld, string versionXml)
	{
		//IL_019e: Unknown result type (might be due to invalid IL or missing references)
		//IL_01a5: Expected O, but got Unknown
		//IL_01cf: Unknown result type (might be due to invalid IL or missing references)
		//IL_01d6: Expected O, but got Unknown
		//IL_04de: Unknown result type (might be due to invalid IL or missing references)
		//IL_04e5: Expected O, but got Unknown
		string text = string.Empty;
		string empty = string.Empty;
		string empty2 = string.Empty;
		string empty3 = string.Empty;
		string empty4 = string.Empty;
		string empty5 = string.Empty;
		string empty6 = string.Empty;
		try
		{
			string[] array = dataFileName.Split(new char[1] { '~' });
			DirectoryInfo directoryInfo = new DirectoryInfo(HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst);
			string path = HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst + "\\BK";
			if (!Directory.Exists(path))
			{
				Directory.CreateDirectory(path);
			}
			empty6 = HTKKConfig.INTERFACE_TEMPLATES + "\\xml\\" + dataFileName + "_" + versionXml + ".xml";
			string[] array2 = array;
			foreach (string text2 in array2)
			{
				FileInfo[] files = directoryInfo.GetFiles(text2 + "*.xml");
				foreach (FileInfo fileInfo in files)
				{
					empty = fileInfo.Name;
					text = fileInfo.FullName;
					empty2 = empty.Replace(text2 + "_", "");
					empty3 = empty.Replace("_" + versionXmlOld + ".xml", "_" + versionXml + ".xml");
					empty4 = fileInfo.DirectoryName + "\\BK\\" + empty;
					if (File.Exists(fileInfo.DirectoryName + "\\" + empty3))
					{
						continue;
					}
					XmlDocument val = new XmlDocument();
					if (!File.Exists(empty6))
					{
						return;
					}
					val.Load(empty6);
					XmlNodeList val2 = ((XmlNode)val).SelectNodes("/Sections/Section");
					XmlDocument val3 = new XmlDocument();
					val3.Load(text);
					XmlNodeList val4 = ((XmlNode)val3).SelectNodes("/Sections/Section");
					XmlNode val5 = null;
					val5 = val4.Item(0);
					if (val5.ChildNodes[0].ChildNodes.Count != 65)
					{
						continue;
					}
					for (int k = 0; k < val2.Count; k++)
					{
						XmlNode val6 = val2.Item(k);
						val5 = val4.Item(k);
						for (int l = 0; l < val6.ChildNodes[0].ChildNodes.Count; l++)
						{
							XmlNodeList childNodes = val6.ChildNodes[0].ChildNodes;
							XmlNode val7 = childNodes.Item(l);
							if (val7.Attributes["Path"] != null)
							{
								for (int m = 0; m < val5.ChildNodes[0].ChildNodes.Count; m++)
								{
									XmlNodeList childNodes2 = val5.ChildNodes[0].ChildNodes;
									XmlNode val8 = childNodes2.Item(m);
									if (val8.Attributes["Path"] != null && ((XmlNode)val7.Attributes["Path"]).Value == "C_17" && ((XmlNode)val8.Attributes["Path"]).Value == "B_17")
									{
										((XmlNode)val7.Attributes["Value"]).Value = ((XmlNode)val8.Attributes["Value"]).Value;
										break;
									}
								}
								continue;
							}
							for (int m = 0; m < val5.ChildNodes[0].ChildNodes.Count; m++)
							{
								XmlNodeList childNodes2 = val5.ChildNodes[0].ChildNodes;
								XmlNode val8 = childNodes2.Item(m);
								if (((XmlNode)val7.Attributes["CellID"]).Value == ((XmlNode)val8.Attributes["CellID"]).Value)
								{
									((XmlNode)val7.Attributes["Value"]).Value = ((XmlNode)val8.Attributes["Value"]).Value;
									break;
								}
							}
						}
					}
					val.DocumentElement.SetAttribute("Version", WinSettings.VersionApp);
					string text3 = HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst + "\\" + empty3;
					XmlTextWriter val9 = new XmlTextWriter(text3, Encoding.UTF8);
					try
					{
						val9.Formatting = (Formatting)1;
						val.Save((XmlWriter)(object)val9);
					}
					finally
					{
						((IDisposable)val9)?.Dispose();
					}
					Directory.Move(text, empty4);
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			HTKKFile.LogRun(MethodBase.GetCurrentMethod().Name, "Loi conver TK: " + text);
		}
	}

	public static void ConverDataVersionKHBS(string dataFileName, string versionXmlOld, string versionXml)
	{
		//IL_015d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0164: Expected O, but got Unknown
		//IL_01ca: Unknown result type (might be due to invalid IL or missing references)
		//IL_01d1: Expected O, but got Unknown
		string text = string.Empty;
		string empty = string.Empty;
		string empty2 = string.Empty;
		string empty3 = string.Empty;
		string empty4 = string.Empty;
		string empty5 = string.Empty;
		try
		{
			string[] array = dataFileName.Split(new char[1] { '~' });
			DirectoryInfo directoryInfo = new DirectoryInfo(HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst);
			string path = HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst + "\\BK";
			if (!Directory.Exists(path))
			{
				Directory.CreateDirectory(path);
			}
			string[] array2 = array;
			foreach (string text2 in array2)
			{
				FileInfo[] files = directoryInfo.GetFiles("*_" + text2 + "*.xml");
				foreach (FileInfo fileInfo in files)
				{
					empty = fileInfo.Name;
					text = fileInfo.FullName;
					empty2 = empty.Replace(text2 + "_", "");
					empty3 = empty.Replace("_" + versionXmlOld + ".xml", "_" + versionXml + ".xml");
					empty4 = fileInfo.DirectoryName + "\\BK\\" + empty;
					if (!File.Exists(fileInfo.DirectoryName + "\\" + empty3))
					{
						XmlDocument val = new XmlDocument();
						val.Load(text);
						val.DocumentElement.SetAttribute("Version", WinSettings.VersionApp);
						string text3 = HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst + "\\" + empty3;
						XmlTextWriter val2 = new XmlTextWriter(text3, Encoding.UTF8);
						try
						{
							val2.Formatting = (Formatting)1;
							val.Save((XmlWriter)(object)val2);
						}
						finally
						{
							((IDisposable)val2)?.Dispose();
						}
						Directory.Move(text, empty4);
					}
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			HTKKFile.LogRun(MethodBase.GetCurrentMethod().Name, "Loi conver TK: " + text);
		}
	}

	public static void ConverData01LPMB_TT80(string dataFileName, string versionXmlOld, string versionXml)
	{
		//IL_00d4: Unknown result type (might be due to invalid IL or missing references)
		//IL_00db: Expected O, but got Unknown
		//IL_01a8: Unknown result type (might be due to invalid IL or missing references)
		//IL_01af: Expected O, but got Unknown
		//IL_0323: Unknown result type (might be due to invalid IL or missing references)
		//IL_032a: Expected O, but got Unknown
		string text = string.Empty;
		string empty = string.Empty;
		string empty2 = string.Empty;
		string empty3 = string.Empty;
		string empty4 = string.Empty;
		string empty5 = string.Empty;
		try
		{
			string[] array = dataFileName.Split(new char[1] { '~' });
			DirectoryInfo directoryInfo = new DirectoryInfo(HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst);
			string path = HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst + "\\BK";
			if (!Directory.Exists(path))
			{
				Directory.CreateDirectory(path);
			}
			string text2 = HTKKConfig.INTERFACE_TEMPLATES + "\\xml\\" + dataFileName + "_" + versionXml + ".xml";
			XmlDocument val = new XmlDocument();
			if (!File.Exists(text2))
			{
				return;
			}
			val.Load(text2);
			XmlNodeList val2 = ((XmlNode)val).SelectNodes("/Sections/Section");
			string[] array2 = array;
			foreach (string text3 in array2)
			{
				FileInfo[] files = directoryInfo.GetFiles("*" + text3 + "*.xml");
				foreach (FileInfo fileInfo in files)
				{
					empty = fileInfo.Name;
					text = fileInfo.FullName;
					empty2 = empty.Replace(text3 + "_", "");
					empty3 = empty.Replace("_" + versionXmlOld + ".xml", "_" + versionXml + ".xml");
					empty4 = fileInfo.DirectoryName + "\\BK\\" + empty;
					XmlDocument val3 = new XmlDocument();
					val3.Load(text);
					XmlNodeList val4 = ((XmlNode)val3).SelectNodes("/Sections/Section");
					if (val4.Count == 5 && val4.Item(1).ChildNodes[0].ChildNodes.Count == 17)
					{
						XmlNode val5 = val2.Item(1);
						XmlNode val6 = val5.ChildNodes[0];
						XmlNodeList childNodes = val6.ChildNodes;
						XmlNode val7 = childNodes.Item(1);
						XmlNode val8 = val4.Item(1);
						XmlNode val9 = val8.ChildNodes[0];
						XmlNodeList childNodes2 = val9.ChildNodes;
						XmlNode val10 = childNodes2.Item(1);
						string value = ((XmlNode)val7.Attributes["Path"]).Value;
						((XmlNode)val10.Attributes["Path"]).Value = value;
						XmlNode val11 = childNodes.Item(2);
						XmlNode val12 = val11.Clone();
						XmlNode val13 = val3.ImportNode(val12, true);
						val9.InsertAfter(val13, val9.ChildNodes[1]);
						val3.DocumentElement.SetAttribute("Version", WinSettings.VersionApp);
						string text4 = HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst + "\\" + empty3;
						XmlTextWriter val14 = new XmlTextWriter(text4, Encoding.UTF8);
						try
						{
							val14.Formatting = (Formatting)1;
							val3.Save((XmlWriter)(object)val14);
						}
						finally
						{
							((IDisposable)val14)?.Dispose();
						}
					}
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			HTKKFile.LogRun(MethodBase.GetCurrentMethod().Name, "Loi conver TK: " + text);
		}
	}

	public static void ConvertDataCommon()
	{
		//IL_0002: Unknown result type (might be due to invalid IL or missing references)
		//IL_0008: Expected O, but got Unknown
		//IL_0053: Unknown result type (might be due to invalid IL or missing references)
		//IL_0059: Expected O, but got Unknown
		try
		{
			XmlDocument val = new XmlDocument();
			val.Load("ConvertDataCommon.xml");
			XmlNodeList val2 = ((XmlNode)val).SelectNodes("/Sections/Section/Items/Item");
			if (val2 == null || val2.Count == 0)
			{
				return;
			}
			foreach (XmlNode item in val2)
			{
				XmlNode val3 = item;
				try
				{
					int num = ((val3.Attributes["Type"] != null) ? Convert.ToInt32(((XmlNode)val3.Attributes["Type"]).Value.Trim()) : (-1));
					int num2 = ((val3.Attributes["Dynamic"] != null) ? Convert.ToInt32(((XmlNode)val3.Attributes["Dynamic"]).Value.Trim()) : (-1));
					int sheetIndexConfig = ((val3.Attributes["SheetIndex"] != null) ? Convert.ToInt32(((XmlNode)val3.Attributes["SheetIndex"]).Value.Trim()) : (-1));
					string id = ((val3.Attributes["ID"] != null) ? ((XmlNode)val3.Attributes["ID"]).Value.Trim() : "");
					if (num == -1)
					{
						continue;
					}
					string versionXML = "";
					string versionKHBS = "";
					string dataFileMenu = "";
					GetInfoInMenu(id, sheetIndexConfig, ref versionXML, ref versionKHBS, ref dataFileMenu);
					if (string.IsNullOrEmpty(versionXML) || string.IsNullOrEmpty(dataFileMenu))
					{
						continue;
					}
					for (int i = 0; i <= 1; i++)
					{
						List<FileInfo> allPathDatafile = GetAllPathDatafile(dataFileMenu, (i == 0) ? versionXML : versionKHBS);
						if (allPathDatafile == null || allPathDatafile.Count == 0)
						{
							continue;
						}
						foreach (FileInfo item2 in allPathDatafile)
						{
							if (num == 1)
							{
								ConvertDataHeaderToSubline(id, val3, num, sheetIndexConfig, item2);
							}
							if (num == 2)
							{
								ConvertDataCQTChiCucTuCQTCapCucCungDong(id, val3, num, sheetIndexConfig, item2);
							}
							if (num == 3)
							{
								ConvertDataTinhHuyenTuCQTCungDong(id, val3, num, sheetIndexConfig, item2);
							}
							if (num == 4)
							{
								ConvertDataLayTieuMucTheoHH(id, val3, num, sheetIndexConfig, item2);
							}
							if (num == 5)
							{
								ConvertDataSetDataStatic(id, val3, num, sheetIndexConfig, item2);
							}
							if (num == 6)
							{
								SetAttributeDataStatic(id, val3, sheetIndexConfig, item2);
							}
						}
					}
				}
				catch (Exception ex)
				{
					HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void ConvertDataHeaderToSubline(string id, XmlNode itemConfig, int typeConfig, int sheetIndexConfig, FileInfo file)
	{
		//IL_00b7: Unknown result type (might be due to invalid IL or missing references)
		//IL_00bd: Expected O, but got Unknown
		//IL_03c6: Unknown result type (might be due to invalid IL or missing references)
		//IL_03cd: Expected O, but got Unknown
		try
		{
			string tablePathConfig = ((itemConfig.Attributes["TablePath"] != null) ? ((XmlNode)itemConfig.Attributes["TablePath"]).Value.Trim() : "");
			string headerPathConfig = ((itemConfig.Attributes["HeaderPath"] != null) ? ((XmlNode)itemConfig.Attributes["HeaderPath"]).Value.Trim() : "");
			string text = ((itemConfig.Attributes["Subline"] != null) ? ((XmlNode)itemConfig.Attributes["Subline"]).Value.Trim() : "");
			XmlDocument val = new XmlDocument();
			val.Load(file.FullName);
			XmlNodeList val2 = ((XmlNode)val).SelectNodes("/Sections/Section");
			if (val2 == null)
			{
				return;
			}
			IEnumerable<XmlNode> enumerable = from XmlNode x in (IEnumerable)val2
				where x.Attributes["TablePath"] != null && ((XmlNode)x.Attributes["TablePath"]).Value.Trim() == tablePathConfig
				select x;
			if (enumerable == null || enumerable.Count() == 0)
			{
				return;
			}
			bool flag = false;
			foreach (XmlNode item in enumerable)
			{
				IEnumerable<XmlNode> enumerable2 = from XmlNode x in (IEnumerable)item
					where x.Name == "RowData"
					select x;
				if (enumerable2 == null || enumerable2.Count() == 0)
				{
					continue;
				}
				XmlNode source = ((IEnumerable)item).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode x) => x.Name == "Header"));
				XmlNode val3 = ((IEnumerable)source).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode x) => x.Attributes["Path"] != null && ((XmlNode)x.Attributes["Path"]).Value.Trim() == headerPathConfig));
				if (val3 == null || val3.Attributes["Value"] == null)
				{
					continue;
				}
				XmlNode val4 = ((IEnumerable)item).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode x) => x.Name == "RowInfo"));
				string text2 = ((XmlNode)val3.Attributes["Value"]).Value.Trim();
				for (int num = 0; num < val4.ChildNodes.Count; num++)
				{
					if (val4.ChildNodes[num] == null || val4.ChildNodes[num].Attributes["Path"] == null || !(((XmlNode)val4.ChildNodes[num].Attributes["Path"]).Value.Trim() == text))
					{
						continue;
					}
					foreach (XmlNode item2 in enumerable2)
					{
						List<string> list = GeneralFunctions.SplitString2(item2.InnerText, "###");
						if (list.Count > num && list[num] != text2)
						{
							list[num] = text2;
							flag = true;
						}
						item2.InnerText = GeneralFunctions.ListToString(list, "###");
					}
				}
			}
			if (!flag)
			{
				return;
			}
			XmlTextWriter val5 = new XmlTextWriter(file.FullName, Encoding.UTF8);
			try
			{
				val5.Formatting = (Formatting)1;
				val.Save((XmlWriter)(object)val5);
			}
			finally
			{
				((IDisposable)val5)?.Dispose();
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void ConvertDataCQTChiCucTuCQTCapCucCungDong(string id, XmlNode itemConfig, int typeConfig, int sheetIndexConfig, FileInfo file)
	{
		//IL_00fe: Unknown result type (might be due to invalid IL or missing references)
		//IL_0104: Expected O, but got Unknown
		//IL_03eb: Unknown result type (might be due to invalid IL or missing references)
		//IL_03f2: Expected O, but got Unknown
		try
		{
			string tablePathConfig = ((itemConfig.Attributes["TablePath"] != null) ? ((XmlNode)itemConfig.Attributes["TablePath"]).Value.Trim() : "");
			string maCQTCapCucPath = ((itemConfig.Attributes["MaCQTCapCucPath"] != null) ? ((XmlNode)itemConfig.Attributes["MaCQTCapCucPath"]).Value.Trim() : "");
			string maCQTCapChiCucPath = ((itemConfig.Attributes["MaCQTCapChiCucPath"] != null) ? ((XmlNode)itemConfig.Attributes["MaCQTCapChiCucPath"]).Value.Trim() : "");
			string tenCQTCapChiCucPath = ((itemConfig.Attributes["TenCQTCapChiCucPath"] != null) ? ((XmlNode)itemConfig.Attributes["TenCQTCapChiCucPath"]).Value.Trim() : "");
			XmlDocument val = new XmlDocument();
			val.Load(file.FullName);
			XmlNodeList val2 = ((XmlNode)val).SelectNodes("/Sections/Section");
			if (val2 == null)
			{
				return;
			}
			IEnumerable<XmlNode> enumerable = from XmlNode x in (IEnumerable)val2
				where x.Attributes["TablePath"] != null && ((XmlNode)x.Attributes["TablePath"]).Value.Trim() == tablePathConfig
				select x;
			if (enumerable == null || enumerable.Count() == 0)
			{
				return;
			}
			bool flag = false;
			foreach (XmlNode item in enumerable)
			{
				IEnumerable<XmlNode> enumerable2 = from XmlNode x in (IEnumerable)item
					where x.Name == "RowData"
					select x;
				if (enumerable2 == null || enumerable2.Count() == 0)
				{
					continue;
				}
				XmlNode val3 = ((IEnumerable)item).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode x) => x.Name == "RowInfo"));
				if (val3 == null)
				{
					continue;
				}
				int num = ((IEnumerable)val3.ChildNodes).Cast<XmlNode>().ToList().FindIndex((XmlNode x) => x.Attributes["Path"] != null && ((XmlNode)x.Attributes["Path"]).Value.Trim() == maCQTCapCucPath);
				int num2 = ((IEnumerable)val3.ChildNodes).Cast<XmlNode>().ToList().FindIndex((XmlNode x) => x.Attributes["Path"] != null && ((XmlNode)x.Attributes["Path"]).Value.Trim() == maCQTCapChiCucPath);
				int index = ((IEnumerable)val3.ChildNodes).Cast<XmlNode>().ToList().FindIndex((XmlNode x) => x.Attributes["Path"] != null && ((XmlNode)x.Attributes["Path"]).Value.Trim() == tenCQTCapChiCucPath);
				if (num < 0)
				{
					continue;
				}
				foreach (XmlNode item2 in enumerable2)
				{
					List<string> list = GeneralFunctions.SplitString2(item2.InnerText, "###");
					string text = list[num];
					if (string.IsNullOrEmpty(text) || text == "0")
					{
						continue;
					}
					string text2 = text.Substring(0, 3) + "00";
					if (num2 >= 0)
					{
						if (list[num2] != text2)
						{
							list[num2] = text2;
							string value = HTKKWinDictionaryDiaDanh.CoQuanThueQuanLy[text2];
							list[index] = value;
							flag = true;
						}
						item2.InnerText = GeneralFunctions.ListToString(list, "###");
					}
				}
			}
			if (!flag)
			{
				return;
			}
			XmlTextWriter val4 = new XmlTextWriter(file.FullName, Encoding.UTF8);
			try
			{
				val4.Formatting = (Formatting)1;
				val.Save((XmlWriter)(object)val4);
			}
			finally
			{
				((IDisposable)val4)?.Dispose();
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void ConvertDataTinhHuyenTuCQTCungDong(string id, XmlNode itemConfig, int typeConfig, int sheetIndexConfig, FileInfo file)
	{
		//IL_01b8: Unknown result type (might be due to invalid IL or missing references)
		//IL_01be: Expected O, but got Unknown
		//IL_06ab: Unknown result type (might be due to invalid IL or missing references)
		//IL_06b2: Expected O, but got Unknown
		try
		{
			string tablePathConfig = ((itemConfig.Attributes["TablePath"] != null) ? ((XmlNode)itemConfig.Attributes["TablePath"]).Value.Trim() : "");
			string maCQTCapCucPath = ((itemConfig.Attributes["MaCQTCapCucPath"] != null) ? ((XmlNode)itemConfig.Attributes["MaCQTCapCucPath"]).Value.Trim() : "");
			string maCQTCapChiCucPath = ((itemConfig.Attributes["MaCQTCapChiCucPath"] != null) ? ((XmlNode)itemConfig.Attributes["MaCQTCapChiCucPath"]).Value.Trim() : "");
			string tenTinhPath = ((itemConfig.Attributes["TenTinhPath"] != null) ? ((XmlNode)itemConfig.Attributes["TenTinhPath"]).Value.Trim() : "");
			string maTinhPath = ((itemConfig.Attributes["MaTinhPath"] != null) ? ((XmlNode)itemConfig.Attributes["MaTinhPath"]).Value.Trim() : "");
			string tenHuyenPath = ((itemConfig.Attributes["TenHuyenPath"] != null) ? ((XmlNode)itemConfig.Attributes["TenHuyenPath"]).Value.Trim() : "");
			string maHuyenPath = ((itemConfig.Attributes["MaHuyenPath"] != null) ? ((XmlNode)itemConfig.Attributes["MaHuyenPath"]).Value.Trim() : "");
			XmlDocument val = new XmlDocument();
			val.Load(file.FullName);
			XmlNodeList val2 = ((XmlNode)val).SelectNodes("/Sections/Section");
			if (val2 == null)
			{
				return;
			}
			IEnumerable<XmlNode> enumerable = from XmlNode x in (IEnumerable)val2
				where x.Attributes["TablePath"] != null && ((XmlNode)x.Attributes["TablePath"]).Value.Trim() == tablePathConfig
				select x;
			if (enumerable == null || enumerable.Count() == 0)
			{
				return;
			}
			bool flag = false;
			foreach (XmlNode item in enumerable)
			{
				IEnumerable<XmlNode> enumerable2 = from XmlNode x in (IEnumerable)item
					where x.Name == "RowData"
					select x;
				if (enumerable2 == null || enumerable2.Count() == 0)
				{
					continue;
				}
				XmlNode val3 = ((IEnumerable)item).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode x) => x.Name == "RowInfo"));
				if (val3 == null)
				{
					continue;
				}
				int num = ((IEnumerable)val3.ChildNodes).Cast<XmlNode>().ToList().FindIndex((XmlNode x) => x.Attributes["Path"] != null && ((XmlNode)x.Attributes["Path"]).Value.Trim() == maCQTCapCucPath);
				int num2 = ((IEnumerable)val3.ChildNodes).Cast<XmlNode>().ToList().FindIndex((XmlNode x) => x.Attributes["Path"] != null && ((XmlNode)x.Attributes["Path"]).Value.Trim() == maCQTCapChiCucPath);
				int num3 = ((IEnumerable)val3.ChildNodes).Cast<XmlNode>().ToList().FindIndex((XmlNode x) => x.Attributes["Path"] != null && ((XmlNode)x.Attributes["Path"]).Value.Trim() == tenTinhPath);
				int num4 = ((IEnumerable)val3.ChildNodes).Cast<XmlNode>().ToList().FindIndex((XmlNode x) => x.Attributes["Path"] != null && ((XmlNode)x.Attributes["Path"]).Value.Trim() == maTinhPath);
				int num5 = ((IEnumerable)val3.ChildNodes).Cast<XmlNode>().ToList().FindIndex((XmlNode x) => x.Attributes["Path"] != null && ((XmlNode)x.Attributes["Path"]).Value.Trim() == tenHuyenPath);
				int num6 = ((IEnumerable)val3.ChildNodes).Cast<XmlNode>().ToList().FindIndex((XmlNode x) => x.Attributes["Path"] != null && ((XmlNode)x.Attributes["Path"]).Value.Trim() == maHuyenPath);
				if (num < 0)
				{
					continue;
				}
				foreach (XmlNode item2 in enumerable2)
				{
					List<string> list = GeneralFunctions.SplitString2(item2.InnerText, "###");
					string text = list[num];
					if (string.IsNullOrEmpty(text) || text == "0")
					{
						continue;
					}
					string text2 = ((num2 < 0) ? "" : list[num2]);
					text = text.Replace(".", "");
					text2 = text2.Replace(".", "");
					string maTinh = text.Substring(0, 3);
					string text3 = "";
					string text4 = "";
					string text5 = "";
					if (!string.IsNullOrEmpty(maTinh))
					{
						KeyValuePair<string, string> keyValuePair = HTKKWinDictionaryDiaDanh.dicTinhTP.FirstOrDefault((KeyValuePair<string, string> x) => x.Value == maTinh);
						if (!string.IsNullOrEmpty(keyValuePair.Key))
						{
							text3 = keyValuePair.Key;
						}
					}
					if (!string.IsNullOrEmpty(text2))
					{
						text4 = (HTKKWinDictionaryDiaDanh.dicQuanHuyen.ContainsKey(text2) ? text2 : "");
						if (!string.IsNullOrEmpty(text4))
						{
							text5 = HTKKWinDictionaryDiaDanh.dicQuanHuyen[text4];
						}
					}
					if (num3 >= 0 && list[num3] != text3)
					{
						list[num3] = text3;
						flag = true;
					}
					if (num4 >= 0 && list[num4] != maTinh)
					{
						list[num4] = maTinh;
						flag = true;
					}
					if (num5 >= 0 && list[num5] != text5)
					{
						list[num5] = text5;
						flag = true;
					}
					if (num6 >= 0 && list[num6] != text4)
					{
						list[num6] = text4;
						flag = true;
					}
					item2.InnerText = GeneralFunctions.ListToString(list, "###");
				}
			}
			if (!flag)
			{
				return;
			}
			XmlTextWriter val4 = new XmlTextWriter(file.FullName, Encoding.UTF8);
			try
			{
				val4.Formatting = (Formatting)1;
				val.Save((XmlWriter)(object)val4);
			}
			finally
			{
				((IDisposable)val4)?.Dispose();
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void ConvertDataLayTieuMucTheoHH(string id, XmlNode itemConfig, int typeConfig, int sheetIndexConfig, FileInfo file)
	{
		//IL_0109: Unknown result type (might be due to invalid IL or missing references)
		//IL_010f: Expected O, but got Unknown
		//IL_0379: Unknown result type (might be due to invalid IL or missing references)
		//IL_0380: Expected O, but got Unknown
		try
		{
			string tablePathConfig = ((itemConfig.Attributes["TablePath"] != null) ? ((XmlNode)itemConfig.Attributes["TablePath"]).Value.Trim() : "");
			string maHangHoaPath = ((itemConfig.Attributes["MaHangHoaPath"] != null) ? ((XmlNode)itemConfig.Attributes["MaHangHoaPath"]).Value.Trim() : "");
			string text = ((itemConfig.Attributes["LoaiHangHoa"] != null) ? ((XmlNode)itemConfig.Attributes["LoaiHangHoa"]).Value.Trim() : "");
			string maTieuMucPath = ((itemConfig.Attributes["MaTieuMucPath"] != null) ? ((XmlNode)itemConfig.Attributes["MaTieuMucPath"]).Value.Trim() : "");
			if (string.IsNullOrEmpty(text))
			{
				return;
			}
			XmlDocument val = new XmlDocument();
			val.Load(file.FullName);
			XmlNodeList val2 = ((XmlNode)val).SelectNodes("/Sections/Section");
			if (val2 == null)
			{
				return;
			}
			IEnumerable<XmlNode> enumerable = from XmlNode x in (IEnumerable)val2
				where x.Attributes["TablePath"] != null && ((XmlNode)x.Attributes["TablePath"]).Value.Trim() == tablePathConfig
				select x;
			if (enumerable == null || enumerable.Count() == 0)
			{
				return;
			}
			bool flag = false;
			foreach (XmlNode item in enumerable)
			{
				IEnumerable<XmlNode> enumerable2 = from XmlNode x in (IEnumerable)item
					where x.Name == "RowData"
					select x;
				if (enumerable2 == null || enumerable2.Count() == 0)
				{
					continue;
				}
				XmlNode val3 = ((IEnumerable)item).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode x) => x.Name == "RowInfo"));
				if (val3 == null)
				{
					continue;
				}
				int index = ((IEnumerable)val3.ChildNodes).Cast<XmlNode>().ToList().FindIndex((XmlNode x) => x.Attributes["Path"] != null && ((XmlNode)x.Attributes["Path"]).Value.Trim() == maHangHoaPath);
				int num = ((IEnumerable)val3.ChildNodes).Cast<XmlNode>().ToList().FindIndex((XmlNode x) => x.Attributes["Path"] != null && ((XmlNode)x.Attributes["Path"]).Value.Trim() == maTieuMucPath);
				foreach (XmlNode item2 in enumerable2)
				{
					List<string> list = GeneralFunctions.SplitString2(item2.InnerText, "###");
					string text2 = list[index];
					if (!string.IsNullOrEmpty(text2))
					{
						string maTieuMucByMaHH = GeneralFunctions.GetMaTieuMucByMaHH(text2, text);
						if (num >= 0 && list[num] != maTieuMucByMaHH)
						{
							list[num] = maTieuMucByMaHH;
							flag = true;
						}
						item2.InnerText = GeneralFunctions.ListToString(list, "###");
					}
				}
			}
			if (!flag)
			{
				return;
			}
			XmlTextWriter val4 = new XmlTextWriter(file.FullName, Encoding.UTF8);
			try
			{
				val4.Formatting = (Formatting)1;
				val.Save((XmlWriter)(object)val4);
			}
			finally
			{
				((IDisposable)val4)?.Dispose();
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void ConvertDataSetDataStatic(string id, XmlNode itemConfig, int typeConfig, int sheetIndexConfig, FileInfo file)
	{
		//IL_00d1: Unknown result type (might be due to invalid IL or missing references)
		//IL_00d7: Expected O, but got Unknown
		//IL_02ee: Unknown result type (might be due to invalid IL or missing references)
		//IL_02f5: Expected O, but got Unknown
		try
		{
			string tablePathConfig = ((itemConfig.Attributes["TablePath"] != null) ? ((XmlNode)itemConfig.Attributes["TablePath"]).Value.Trim() : "");
			string MCTPath = ((itemConfig.Attributes["MCTPath"] != null) ? ((XmlNode)itemConfig.Attributes["MCTPath"]).Value.Trim() : "");
			string text = ((itemConfig.Attributes["ValueStatic"] != null) ? ((XmlNode)itemConfig.Attributes["ValueStatic"]).Value.Trim() : "");
			if (string.IsNullOrEmpty(MCTPath))
			{
				return;
			}
			XmlDocument val = new XmlDocument();
			val.Load(file.FullName);
			XmlNodeList val2 = ((XmlNode)val).SelectNodes("/Sections/Section");
			if (val2 == null)
			{
				return;
			}
			IEnumerable<XmlNode> enumerable = from XmlNode x in (IEnumerable)val2
				where x.Attributes["TablePath"] != null && ((XmlNode)x.Attributes["TablePath"]).Value.Trim() == tablePathConfig
				select x;
			if (enumerable == null || enumerable.Count() == 0)
			{
				return;
			}
			bool flag = false;
			foreach (XmlNode item in enumerable)
			{
				IEnumerable<XmlNode> enumerable2 = from XmlNode x in (IEnumerable)item
					where x.Name == "RowData"
					select x;
				if (enumerable2 == null || enumerable2.Count() == 0)
				{
					continue;
				}
				XmlNode val3 = ((IEnumerable)item).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode x) => x.Name == "RowInfo"));
				if (val3 == null)
				{
					continue;
				}
				int num = ((IEnumerable)val3.ChildNodes).Cast<XmlNode>().ToList().FindIndex((XmlNode x) => x.Attributes["Path"] != null && ((XmlNode)x.Attributes["Path"]).Value.Trim() == MCTPath);
				foreach (XmlNode item2 in enumerable2)
				{
					List<string> list = GeneralFunctions.SplitString2(item2.InnerText, "###");
					string text2 = list[num];
					if (num >= 0 && list[num] != text)
					{
						list[num] = text;
						flag = true;
					}
					item2.InnerText = GeneralFunctions.ListToString(list, "###");
				}
			}
			if (!flag)
			{
				return;
			}
			XmlTextWriter val4 = new XmlTextWriter(file.FullName, Encoding.UTF8);
			try
			{
				val4.Formatting = (Formatting)1;
				val.Save((XmlWriter)(object)val4);
			}
			finally
			{
				((IDisposable)val4)?.Dispose();
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static List<FileInfo> GetAllPathDatafile(string dataFileMenu, string versionXML)
	{
		List<FileInfo> list = new List<FileInfo>();
		try
		{
			string text = HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst;
			DirectoryInfo directoryInfo = new DirectoryInfo(HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst);
			FileInfo[] files = directoryInfo.GetFiles("*" + dataFileMenu + "*.xml");
			foreach (FileInfo fileInfo in files)
			{
				if (fileInfo.Name.EndsWith(versionXML + ".xml"))
				{
					list.Add(fileInfo);
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return list;
	}

	public static void GetInfoInMenu(string id, int sheetIndexConfig, ref string versionXML, ref string versionKHBS, ref string dataFileMenu)
	{
		//IL_002d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0033: Expected O, but got Unknown
		//IL_011b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0122: Expected O, but got Unknown
		try
		{
			string text = "menu.xml";
			if (!File.Exists(text))
			{
				return;
			}
			XmlDocument val = new XmlDocument();
			val.LoadXml(HTKKFile.LoadFile(text));
			XmlNodeList source = ((XmlNode)val).SelectNodes("/Root/Menu");
			XmlNode val2 = ((IEnumerable)source).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode x) => x.Attributes["ID"] != null && ((XmlNode)x.Attributes["ID"]).Value.Trim() == id));
			if (val2 == null)
			{
				return;
			}
			versionXML = ((val2.Attributes["VersionXML"] != null) ? ((XmlNode)val2.Attributes["VersionXML"]).Value.Replace(".", "") : "");
			versionKHBS = ((val2.Attributes["VersionKHBS"] != null) ? ((XmlNode)val2.Attributes["VersionKHBS"]).Value.Replace(".", "") : "");
			foreach (XmlNode childNode in val2.ChildNodes)
			{
				XmlNode val3 = childNode;
				for (int num = 0; num < val3.ChildNodes.Count; num++)
				{
					if (num == sheetIndexConfig)
					{
						dataFileMenu = ((val3.ChildNodes[num].Attributes["DataFile"] != null) ? ((XmlNode)val3.ChildNodes[num].Attributes["DataFile"]).Value.Trim() : "");
						break;
					}
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static bool SetToAllRowData(ref IEnumerable<XmlNode> RowData, int colIndex, string value)
	{
		bool result = false;
		try
		{
			foreach (XmlNode RowDatum in RowData)
			{
				List<string> list = GeneralFunctions.SplitString2(RowDatum.InnerText, "###");
				if (list.Count > colIndex && list[colIndex] != value)
				{
					list[colIndex] = value;
					result = true;
				}
				RowDatum.InnerText = GeneralFunctions.ListToString(list, "###");
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return result;
	}

	public static void SetAttributeDataStatic(string id, XmlNode itemConfig, int sheetIndexConfig, FileInfo file)
	{
		//IL_00c8: Unknown result type (might be due to invalid IL or missing references)
		//IL_00ce: Expected O, but got Unknown
		//IL_01ef: Unknown result type (might be due to invalid IL or missing references)
		//IL_01f6: Expected O, but got Unknown
		try
		{
			string MCTPath = ((itemConfig.Attributes["MCTPath"] != null) ? ((XmlNode)itemConfig.Attributes["MCTPath"]).Value.Trim() : "");
			string text = ((itemConfig.Attributes["Attribute"] != null) ? ((XmlNode)itemConfig.Attributes["Attribute"]).Value.Trim() : "");
			string value = ((itemConfig.Attributes["ValueStatic"] != null) ? ((XmlNode)itemConfig.Attributes["ValueStatic"]).Value.Trim() : "");
			if (string.IsNullOrEmpty(MCTPath))
			{
				return;
			}
			XmlDocument val = new XmlDocument();
			val.Load(file.FullName);
			XmlNodeList val2 = ((XmlNode)val).SelectNodes("/Sections/Section/Cells/Cell");
			if (val2 == null || val2.Count == 0)
			{
				return;
			}
			IEnumerable<XmlNode> enumerable = from XmlNode x in (IEnumerable)val2
				where x.Attributes["Path"] != null && ((XmlNode)x.Attributes["Path"]).Value.Trim() == MCTPath
				select x;
			if (enumerable == null || enumerable.Count() == 0)
			{
				return;
			}
			bool flag = false;
			foreach (XmlNode item in enumerable)
			{
				if (item.Attributes[text] != null)
				{
					((XmlNode)item.Attributes[text]).Value = value;
				}
				else
				{
					XmlDocument ownerDocument = item.OwnerDocument;
					XmlAttribute val3 = ownerDocument.CreateAttribute(text);
					((XmlNode)val3).Value = value;
					item.Attributes.Append(val3);
				}
				flag = true;
			}
			if (!flag)
			{
				return;
			}
			XmlTextWriter val4 = new XmlTextWriter(file.FullName, Encoding.UTF8);
			try
			{
				val4.Formatting = (Formatting)1;
				val.Save((XmlWriter)(object)val4);
			}
			finally
			{
				((IDisposable)val4)?.Dispose();
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void DeleteCellSameSInfoStatic(FileInfo file)
	{
		//IL_0002: Unknown result type (might be due to invalid IL or missing references)
		//IL_0008: Expected O, but got Unknown
		//IL_0056: Unknown result type (might be due to invalid IL or missing references)
		//IL_005c: Expected O, but got Unknown
		//IL_018d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0194: Expected O, but got Unknown
		try
		{
			XmlDocument val = new XmlDocument();
			val.Load(file.FullName);
			XmlNodeList val2 = ((XmlNode)val).SelectNodes("/Sections/Section/Cells/Cell");
			if (val2 == null || val2.Count == 0)
			{
				return;
			}
			bool flag = false;
			foreach (XmlNode item in val2)
			{
				XmlNode val3 = item;
				string path = ((val3.Attributes["Path"] != null) ? ((XmlNode)val3.Attributes["Path"]).Value : "");
				if (string.IsNullOrEmpty(path))
				{
					continue;
				}
				List<XmlNode> list = (from XmlNode x in (IEnumerable)val2
					where x.Attributes["Path"] != null && ((XmlNode)x.Attributes["Path"]).Value.Trim() == path
					select x).ToList();
				if (list == null || list.Count() <= 1)
				{
					continue;
				}
				for (int num = 1; num < list.Count; num++)
				{
					if (list != null && list[num] != null && val3.ParentNode != null)
					{
						val3.ParentNode.RemoveChild(list[num]);
					}
				}
				flag = true;
			}
			if (!flag)
			{
				return;
			}
			XmlTextWriter val4 = new XmlTextWriter(file.FullName, Encoding.UTF8);
			try
			{
				val4.Formatting = (Formatting)1;
				val.Save((XmlWriter)(object)val4);
			}
			finally
			{
				((IDisposable)val4)?.Dispose();
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void FixDatafileConfig_03_3C_TNDN_TT80()
	{
		//IL_002d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0034: Expected O, but got Unknown
		//IL_0096: Unknown result type (might be due to invalid IL or missing references)
		//IL_009d: Expected O, but got Unknown
		//IL_0234: Unknown result type (might be due to invalid IL or missing references)
		//IL_023b: Expected O, but got Unknown
		//IL_00b9: Unknown result type (might be due to invalid IL or missing references)
		//IL_00c0: Expected O, but got Unknown
		try
		{
			string dataFileMenu = "03_3C_TNDN_TT80";
			string versionXML = "251";
			List<FileInfo> allPathDatafile = GetAllPathDatafile(dataFileMenu, versionXML);
			foreach (FileInfo item in allPathDatafile)
			{
				XmlDocument val = new XmlDocument();
				val.LoadXml(HTKKFile.LoadFile(item.FullName));
				XmlNodeList val2 = ((XmlNode)val).SelectNodes("/Sections/Section");
				if (val2 == null || val2.Count == 0)
				{
					continue;
				}
				int num = 23;
				int num2 = 74;
				bool flag = false;
				foreach (XmlNode item2 in val2)
				{
					XmlNode val3 = item2;
					foreach (XmlNode childNode in val3.ChildNodes)
					{
						XmlNode val4 = childNode;
						if (!(val4.Name == "Header"))
						{
							continue;
						}
						XmlNode val5 = ((IEnumerable)val4.ChildNodes).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode x) => x.Attributes["Path"] != null && ((XmlNode)x.Attributes["Path"]).Value.Trim() == "path_tieu_muc"));
						if (val5 != null)
						{
							if (((XmlNode)val5.Attributes["CellID"]).Value.IndexOf(num.ToString()) < 0)
							{
								flag = true;
							}
							((XmlNode)val5.Attributes["CellID"]).Value = "T_" + num;
							((XmlNode)val5.Attributes["Value"]).Value = "1052";
							((XmlNode)val5.Attributes["DefaultValue"]).Value = "1052";
							num += num2;
						}
					}
				}
				if (flag)
				{
					XmlTextWriter val6 = new XmlTextWriter(item.FullName, Encoding.UTF8);
					try
					{
						val6.Formatting = (Formatting)1;
						val.Save((XmlWriter)(object)val6);
					}
					finally
					{
						((IDisposable)val6)?.Dispose();
					}
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void FixDataFile_01_6_CSSX_GTGT_TT80()
	{
		//IL_0034: Unknown result type (might be due to invalid IL or missing references)
		//IL_003b: Expected O, but got Unknown
		//IL_0096: Unknown result type (might be due to invalid IL or missing references)
		//IL_009d: Expected O, but got Unknown
		//IL_0445: Unknown result type (might be due to invalid IL or missing references)
		//IL_044c: Expected O, but got Unknown
		try
		{
			string dataFileMenu = "01_6_GTGT_CSSX_TT80";
			string versionXML = "251";
			string text = "Pluc_01_6_GTGT_CSSX/NoiDungBangPhanBo";
			List<FileInfo> allPathDatafile = GetAllPathDatafile(dataFileMenu, versionXML);
			foreach (FileInfo item in allPathDatafile)
			{
				XmlDocument val = new XmlDocument();
				val.LoadXml(HTKKFile.LoadFile(item.FullName));
				XmlNodeList val2 = ((XmlNode)val).SelectNodes("/Sections/Section");
				if (val2 == null || val2.Count == 0)
				{
					continue;
				}
				bool flag = false;
				foreach (XmlNode item2 in val2)
				{
					XmlNode val3 = item2;
					if (val3.Attributes["TablePath"] == null || !(((XmlNode)val3.Attributes["TablePath"]).Value.Trim() == text))
					{
						continue;
					}
					IEnumerable<XmlNode> enumerable = from XmlNode x in (IEnumerable)val3
						where x.Name == "RowData"
						select x;
					if (enumerable == null || enumerable.Count() == 0)
					{
						continue;
					}
					XmlNode val4 = ((IEnumerable)val3).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode x) => x.Name == "RowInfo"));
					int num = -1;
					int num2 = -1;
					for (int num3 = 0; num3 < val4.ChildNodes.Count; num3++)
					{
						if (val4.ChildNodes[num3] != null && val4.ChildNodes[num3].Attributes["Path"] != null && ((XmlNode)val4.ChildNodes[num3].Attributes["Path"]).Value.Trim() == "ct17")
						{
							num = num3;
						}
						if (val4.ChildNodes[num3] != null && val4.ChildNodes[num3].Attributes["Path"] != null && ((XmlNode)val4.ChildNodes[num3].Attributes["Path"]).Value.Trim() == "ct21")
						{
							num2 = num3;
						}
						if (val4.ChildNodes[num3] == null || val4.ChildNodes[num3].Attributes["Path"] == null || !(((XmlNode)val4.ChildNodes[num3].Attributes["Path"]).Value.Trim() == "path_21") || num < 0 || num2 < 0)
						{
							continue;
						}
						foreach (XmlNode item3 in enumerable)
						{
							List<string> list = GeneralFunctions.SplitString2(item3.InnerText, "###");
							if (list.Count > num3)
							{
								string text2 = list[num];
								string text3 = list[num2];
								string text4 = list[num3];
								if (text4 == "0" && (text2 != "0" || text3 != "0"))
								{
									list[num3] = ((text2 != "0") ? text2 : text3);
									flag = true;
								}
							}
							item3.InnerText = GeneralFunctions.ListToString(list, "###");
						}
					}
				}
				if (flag)
				{
					XmlTextWriter val5 = new XmlTextWriter(item.FullName, Encoding.UTF8);
					try
					{
						val5.Formatting = (Formatting)1;
						val.Save((XmlWriter)(object)val5);
					}
					finally
					{
						((IDisposable)val5)?.Dispose();
					}
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void Createa_PL_KHBS()
	{
		//IL_0002: Unknown result type (might be due to invalid IL or missing references)
		//IL_0008: Expected O, but got Unknown
		//IL_0053: Unknown result type (might be due to invalid IL or missing references)
		//IL_0059: Expected O, but got Unknown
		//IL_0568: Unknown result type (might be due to invalid IL or missing references)
		//IL_056f: Expected O, but got Unknown
		try
		{
			XmlDocument val = new XmlDocument();
			val.Load("CreateDataFileKHBSMiss.xml");
			XmlNodeList val2 = ((XmlNode)val).SelectNodes("/Sections/Section/Items/Item");
			if (val2 == null || val2.Count == 0)
			{
				return;
			}
			foreach (XmlNode item in val2)
			{
				XmlNode val3 = item;
				string text = "_" + ((val3.Attributes["DataFile"] != null) ? ((XmlNode)val3.Attributes["DataFile"]).Value.Trim() : "");
				string text2 = ((val3.Attributes["VersionXML"] != null) ? ((XmlNode)val3.Attributes["VersionXML"]).Value.Trim() : "");
				string text3 = ((val3.Attributes["VersionKHBS"] != null) ? ((XmlNode)val3.Attributes["VersionKHBS"]).Value.Trim() : "");
				string text4 = ((val3.Attributes["IsConvert"] != null) ? ((XmlNode)val3.Attributes["IsConvert"]).Value.Trim() : "");
				string text5 = ((val3.Attributes["PL_01_KHBS"] != null) ? ((XmlNode)val3.Attributes["PL_01_KHBS"]).Value.Trim() : "");
				string text6 = ((val3.Attributes["PL_01_1_KHBS"] != null) ? ((XmlNode)val3.Attributes["PL_01_1_KHBS"]).Value.Trim() : "");
				if (!text4.Equals("0"))
				{
					continue;
				}
				List<FileInfo> allPathDatafile = GetAllPathDatafile(text, text2);
				foreach (FileInfo item2 in allPathDatafile)
				{
					if (item2.Name.StartsWith("bs") && item2.Name.IndexOf("_VIEW") < 0 && item2.Name.EndsWith(text2 + ".xml") && item2.Name.IndexOf("_TT80") < 0 && item2.Name.IndexOf("_KHBS") < 0)
					{
						string text7 = item2.Name.Substring(0, item2.Name.IndexOf(text));
						string text8 = item2.Name.Substring(item2.Name.IndexOf(text) + text.Length + 1);
						string text9 = "";
						string text10 = "";
						text9 = text7 + "_" + text5 + "_" + text8;
						text9 = text9.Replace("_" + text2, "_" + text3);
						text10 = text7 + "_" + text6 + "_" + text8;
						text10 = text10.Replace("_" + text2, "_" + text3);
						if (!File.Exists(HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst + "\\" + text9))
						{
							string text11 = "";
							text11 = ((!"D1,N8,G9,H2,H4,T8,E8,L8,L7,N2,G2,R5,Y6,Z8,Z5".Contains(WinSettings.MenuID)) ? (HTKKConfig.INTERFACE_TEMPLATES + "\\xml\\KHBS_01_TT80_" + text3 + ".xml") : (HTKKConfig.INTERFACE_TEMPLATES + "\\xml\\KHBS_01_TT80_DIA_BAN_" + text3 + ".xml"));
							string destFileName = HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst + "\\" + text9;
							File.Copy(text11, destFileName);
						}
						if (!File.Exists(HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst + "\\" + text10))
						{
							string text11 = HTKKConfig.INTERFACE_TEMPLATES + "\\xml\\KHBS_01_1_TT80_" + text3 + ".xml";
							string destFileName = HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst + "\\" + text10;
							File.Copy(text11, destFileName);
						}
					}
				}
				((XmlNode)val3.Attributes["IsConvert"]).Value = "1";
			}
			XmlTextWriter val4 = new XmlTextWriter("CreateDataFileKHBSMiss.xml", Encoding.UTF8);
			try
			{
				val4.Formatting = (Formatting)1;
				val.Save((XmlWriter)(object)val4);
			}
			finally
			{
				((IDisposable)val4)?.Dispose();
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void FixDataFile_03_8_TNDN_TT80()
	{
		//IL_003c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0043: Expected O, but got Unknown
		//IL_009e: Unknown result type (might be due to invalid IL or missing references)
		//IL_00a5: Expected O, but got Unknown
		//IL_03d4: Unknown result type (might be due to invalid IL or missing references)
		//IL_03db: Expected O, but got Unknown
		try
		{
			string dataFileMenu = "03_8_TNDN_TT80";
			string versionXML = "251";
			string text = "PLuc_03_8_TNDN/PhanBoThueTNDN/NoiDungBangPhanBo";
			string text2 = "1052";
			List<FileInfo> allPathDatafile = GetAllPathDatafile(dataFileMenu, versionXML);
			foreach (FileInfo item in allPathDatafile)
			{
				XmlDocument val = new XmlDocument();
				val.LoadXml(HTKKFile.LoadFile(item.FullName));
				XmlNodeList val2 = ((XmlNode)val).SelectNodes("/Sections/Section");
				if (val2 == null || val2.Count == 0)
				{
					continue;
				}
				bool flag = false;
				foreach (XmlNode item2 in val2)
				{
					XmlNode val3 = item2;
					if (val3.Attributes["TablePath"] == null || !(((XmlNode)val3.Attributes["TablePath"]).Value.Trim() == text))
					{
						continue;
					}
					IEnumerable<XmlNode> enumerable = from XmlNode x in (IEnumerable)val3
						where x.Name == "RowData"
						select x;
					if (enumerable == null || enumerable.Count() == 0)
					{
						continue;
					}
					XmlNode val4 = ((IEnumerable)val3).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode x) => x.Name == "RowInfo"));
					int num = -1;
					int num2 = -1;
					for (int num3 = 0; num3 < val4.ChildNodes.Count; num3++)
					{
						if (val4.ChildNodes[num3] != null && val4.ChildNodes[num3].Attributes["Path"] != null && ((XmlNode)val4.ChildNodes[num3].Attributes["Path"]).Value.Trim() == "path_ma_tieu_muc")
						{
							num = num3;
						}
						if (val4.ChildNodes[num3] != null && val4.ChildNodes[num3].Attributes["Path"] != null && ((XmlNode)val4.ChildNodes[num3].Attributes["Path"]).Value.Trim() == "path_ten_tieu_muc")
						{
							num2 = num3;
						}
						if (num != num3 && num2 != num3)
						{
							continue;
						}
						foreach (XmlNode item3 in enumerable)
						{
							List<string> list = GeneralFunctions.SplitString2(item3.InnerText, "###");
							if (list.Count > num3)
							{
								string text3 = ((num >= 0) ? list[num] : "");
								string text4 = ((num2 >= 0) ? list[num2] : "");
								string text5 = list[num3];
								if ((num >= 0 && text3 != text2) || (num2 >= 0 && text4 != text2))
								{
									list[num3] = text2;
									flag = true;
								}
							}
							item3.InnerText = GeneralFunctions.ListToString(list, "###");
						}
					}
				}
				if (flag)
				{
					XmlTextWriter val5 = new XmlTextWriter(item.FullName, Encoding.UTF8);
					try
					{
						val5.Formatting = (Formatting)1;
						val.Save((XmlWriter)(object)val5);
					}
					finally
					{
						((IDisposable)val5)?.Dispose();
					}
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void FixDataFile_01_DNHUY()
	{
		//IL_002d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0034: Expected O, but got Unknown
		//IL_008b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0092: Expected O, but got Unknown
		//IL_0136: Unknown result type (might be due to invalid IL or missing references)
		//IL_013d: Expected O, but got Unknown
		try
		{
			string dataFileMenu = "01_DNHUY_TT80";
			string versionXML = "252";
			List<FileInfo> allPathDatafile = GetAllPathDatafile(dataFileMenu, versionXML);
			foreach (FileInfo item in allPathDatafile)
			{
				XmlDocument val = new XmlDocument();
				val.LoadXml(HTKKFile.LoadFile(item.FullName));
				XmlNodeList val2 = ((XmlNode)val).SelectNodes("/Sections/Section");
				if (val2 == null || val2.Count == 0)
				{
					continue;
				}
				bool flag = false;
				foreach (XmlNode item2 in val2)
				{
					XmlNode val3 = item2;
					XmlNode val4 = ((IEnumerable)val3.FirstChild).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode x) => x.Attributes["CellID"] != null && ((XmlNode)x.Attributes["CellID"]).Value == "V_21"));
					if (val4 != null)
					{
						((XmlNode)val4.Attributes["CellID"]).Value = "H_21";
						flag = true;
					}
				}
				if (flag)
				{
					XmlTextWriter val5 = new XmlTextWriter(item.FullName, Encoding.UTF8);
					try
					{
						val5.Formatting = (Formatting)1;
						val.Save((XmlWriter)(object)val5);
					}
					finally
					{
						((IDisposable)val5)?.Dispose();
					}
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void FixFileNameKHBS_01_GTGT_TT80()
	{
		//IL_005a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0061: Expected O, but got Unknown
		//IL_00b9: Unknown result type (might be due to invalid IL or missing references)
		//IL_00c0: Expected O, but got Unknown
		//IL_01a5: Unknown result type (might be due to invalid IL or missing references)
		//IL_01ac: Expected O, but got Unknown
		try
		{
			string text = "KHBS_01_02_GTGT_TT80~KHBS_01_1_02_GTGT_TT80";
			string versionXML = "251";
			string[] array = text.Split(new char[1] { '~' });
			string[] array2 = array;
			foreach (string text2 in array2)
			{
				List<FileInfo> allPathDatafile = GetAllPathDatafile(text2, versionXML);
				foreach (FileInfo item in allPathDatafile)
				{
					XmlDocument val = new XmlDocument();
					val.LoadXml(HTKKFile.LoadFile(item.FullName));
					XmlNodeList val2 = ((XmlNode)val).SelectNodes("/Sections/Section");
					if (val2 == null || val2.Count == 0)
					{
						continue;
					}
					bool flag = false;
					foreach (XmlNode item2 in val2)
					{
						XmlNode val3 = item2;
						XmlNode val4 = ((IEnumerable)val3.FirstChild).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode x) => x.Attributes["Value"] != null && ((XmlNode)x.Attributes["Value"]).Value == "01/GTGT"));
						if (val4 != null)
						{
							flag = true;
							break;
						}
					}
					if (!flag)
					{
						continue;
					}
					string text3 = item.DirectoryName + "\\BK\\" + item.Name;
					if (!File.Exists(text3))
					{
						Directory.Move(item.FullName, text3);
					}
					else
					{
						Directory.Delete(item.FullName);
					}
					string text4 = item.FullName.Replace("02_GTGT_TT80", "01_GTGT_TT80");
					XmlTextWriter val5 = new XmlTextWriter(text4, Encoding.UTF8);
					try
					{
						val5.Formatting = (Formatting)1;
						val.Save((XmlWriter)(object)val5);
					}
					finally
					{
						((IDisposable)val5)?.Dispose();
					}
					string[] array3 = item.Name.Split(new char[1] { '_' });
					string dataFileMenu = array3[0] + "_02_GTGT_TT80";
					List<FileInfo> allPathDatafile2 = GetAllPathDatafile(dataFileMenu, versionXML);
					foreach (FileInfo item3 in allPathDatafile2)
					{
						if (text2.IndexOf("KHBS_01_02_GTGT_TT80") >= 0)
						{
							string text5 = item3.DirectoryName + "\\" + item3.Name.Replace(array3[0], array3[0] + "_KHBS_01");
							if (File.Exists(text5))
							{
								continue;
							}
							string sourceFileName = HTKKConfig.INTERFACE_TEMPLATES + "\\xml\\KHBS_01_TT80_268.xml";
							File.Copy(sourceFileName, text5, overwrite: false);
						}
						if (text2.IndexOf("KHBS_01_1_02_GTGT_TT80") >= 0)
						{
							string text5 = item3.DirectoryName + "\\" + item3.Name.Replace(array3[0], array3[0] + "_KHBS_01_1");
							if (!File.Exists(text5))
							{
								string sourceFileName = HTKKConfig.INTERFACE_TEMPLATES + "\\xml\\KHBS_01_1_TT80_268.xml";
								File.Copy(sourceFileName, text5, overwrite: false);
							}
						}
					}
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void FixFileNameKHBS_02_TTDB_TT80()
	{
		//IL_005a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0061: Expected O, but got Unknown
		//IL_00b9: Unknown result type (might be due to invalid IL or missing references)
		//IL_00c0: Expected O, but got Unknown
		//IL_01f8: Unknown result type (might be due to invalid IL or missing references)
		//IL_01ff: Expected O, but got Unknown
		try
		{
			string text = "KHBS_01_TTDB_TT80~KHBS_01_1_TTDB_TT80";
			string versionXML = "251";
			string[] array = text.Split(new char[1] { '~' });
			string[] array2 = array;
			foreach (string text2 in array2)
			{
				List<FileInfo> allPathDatafile = GetAllPathDatafile(text2, versionXML);
				foreach (FileInfo item in allPathDatafile)
				{
					XmlDocument val = new XmlDocument();
					val.LoadXml(HTKKFile.LoadFile(item.FullName));
					XmlNodeList val2 = ((XmlNode)val).SelectNodes("/Sections/Section");
					if (val2 == null || val2.Count == 0)
					{
						continue;
					}
					bool flag = false;
					foreach (XmlNode item2 in val2)
					{
						XmlNode val3 = item2;
						XmlNode val4 = ((IEnumerable)val3.FirstChild).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode x) => x.Attributes["Value"] != null && ((XmlNode)x.Attributes["Value"]).Value == "02/TTĐB"));
						if (val4 != null)
						{
							flag = true;
							break;
						}
					}
					if (!flag)
					{
						continue;
					}
					string text3 = item.DirectoryName + "\\BK\\" + item.Name;
					if (!File.Exists(text3))
					{
						Directory.Move(item.FullName, text3);
					}
					else
					{
						Directory.Delete(item.FullName);
					}
					string text4 = "";
					if (item.FullName.IndexOf("KHBS_01_TTDB_TT80") >= 0)
					{
						text4 = item.FullName.Replace("KHBS_01_TTDB_TT80", "KHBS_01_02_TTDB_TT80");
					}
					if (item.FullName.IndexOf("KHBS_01_1_TTDB_TT80") >= 0)
					{
						text4 = item.FullName.Replace("KHBS_01_1_TTDB_TT80", "KHBS_01_1_02_TTDB_TT80");
					}
					XmlTextWriter val5 = new XmlTextWriter(text4, Encoding.UTF8);
					try
					{
						val5.Formatting = (Formatting)1;
						val.Save((XmlWriter)(object)val5);
					}
					finally
					{
						((IDisposable)val5)?.Dispose();
					}
					string[] array3 = item.Name.Split(new char[1] { '_' });
					string dataFileMenu = array3[0] + "_01_TTDB_TT80";
					List<FileInfo> allPathDatafile2 = GetAllPathDatafile(dataFileMenu, "261");
					foreach (FileInfo item3 in allPathDatafile2)
					{
						if (text2.IndexOf("KHBS_01_TTDB_TT80") >= 0)
						{
							string text5 = item3.DirectoryName + "\\" + item3.Name.Replace(array3[0], array3[0] + "_KHBS_01");
							text5 = text5.Replace("01_TTDB_TT80", "TTDB_TT80");
							if (File.Exists(text5))
							{
								continue;
							}
							string sourceFileName = HTKKConfig.INTERFACE_TEMPLATES + "\\xml\\KHBS_01_TT80_261.xml";
							File.Copy(sourceFileName, text5, overwrite: false);
						}
						if (text2.IndexOf("KHBS_01_1_TTDB_TT80") >= 0)
						{
							string text5 = item3.DirectoryName + "\\" + item3.Name.Replace(array3[0], array3[0] + "_KHBS_01_1");
							text5 = text5.Replace("01_TTDB_TT80", "TTDB_TT80");
							if (!File.Exists(text5))
							{
								string sourceFileName = HTKKConfig.INTERFACE_TEMPLATES + "\\xml\\KHBS_01_1_TT80_261.xml";
								File.Copy(sourceFileName, text5, overwrite: false);
							}
						}
					}
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void FixFileNameThangToiNgay(string dataFiles, string versionXml)
	{
		//IL_017c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0183: Expected O, but got Unknown
		//IL_01de: Unknown result type (might be due to invalid IL or missing references)
		//IL_01e5: Expected O, but got Unknown
		try
		{
			string[] array = dataFiles.Split(new char[1] { '~' });
			string[] array2 = array;
			foreach (string dataFileMenu in array2)
			{
				List<FileInfo> allPathDatafile = GetAllPathDatafile(dataFileMenu, versionXml);
				foreach (FileInfo item in allPathDatafile)
				{
					if (item.Name.IndexOf("_CD_") >= 0 || item.Name.IndexOf("_DT_") >= 0 || item.Name.IndexOf("_KTN_") >= 0)
					{
						continue;
					}
					string[] array3 = item.FullName.Split(new char[1] { '_' });
					if (array3.Length < 3)
					{
						continue;
					}
					string text = array3[^3];
					string text2 = array3[^2];
					if (text.Length == 6 && text2.Length == 6)
					{
						text = "01" + text;
						text2 = "01" + text2;
						DateTime dateTime = DateTime.ParseExact(text2, "ddMMyyyy", CultureInfo.CurrentCulture).AddMonths(1).AddDays(-1.0);
						array3[^3] = text;
						array3[^2] = dateTime.ToString("ddMMyyyy");
						string text3 = GeneralFunctions.ListToString(array3.ToList(), "_");
						XmlDocument val = new XmlDocument();
						val.LoadXml(HTKKFile.LoadFile(item.FullName));
						string text4 = item.DirectoryName + "\\BK\\" + item.Name;
						if (!File.Exists(text4))
						{
							Directory.Move(item.FullName, text4);
						}
						else
						{
							Directory.Delete(item.FullName);
						}
						XmlTextWriter val2 = new XmlTextWriter(text3, Encoding.UTF8);
						try
						{
							val2.Formatting = (Formatting)1;
							val.Save((XmlWriter)(object)val2);
						}
						finally
						{
							((IDisposable)val2)?.Dispose();
						}
					}
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void FixFileNameNamToiNgay(string dataFiles, string versionXml)
	{
		//IL_00f1: Unknown result type (might be due to invalid IL or missing references)
		//IL_00f8: Expected O, but got Unknown
		//IL_0153: Unknown result type (might be due to invalid IL or missing references)
		//IL_015a: Expected O, but got Unknown
		try
		{
			string[] array = dataFiles.Split(new char[1] { '~' });
			string[] array2 = array;
			foreach (string dataFileMenu in array2)
			{
				List<FileInfo> allPathDatafile = GetAllPathDatafile(dataFileMenu, versionXml);
				foreach (FileInfo item in allPathDatafile)
				{
					string[] array3 = item.FullName.Split(new char[1] { '_' });
					if (array3.Length < 3)
					{
						continue;
					}
					string text = array3[^2];
					if (text.Trim().Length == 4)
					{
						array3[^2] = text + "_0101" + text + "_3112" + text;
						string text2 = GeneralFunctions.ListToString(array3.ToList(), "_");
						XmlDocument val = new XmlDocument();
						val.LoadXml(HTKKFile.LoadFile(item.FullName));
						string text3 = item.DirectoryName + "\\BK\\" + item.Name;
						if (!File.Exists(text3))
						{
							Directory.Move(item.FullName, text3);
						}
						else
						{
							Directory.Delete(item.FullName);
						}
						XmlTextWriter val2 = new XmlTextWriter(text2, Encoding.UTF8);
						try
						{
							val2.Formatting = (Formatting)1;
							val.Save((XmlWriter)(object)val2);
						}
						finally
						{
							((IDisposable)val2)?.Dispose();
						}
					}
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void FixFileDauKhiLoaiTien()
	{
		//IL_01da: Unknown result type (might be due to invalid IL or missing references)
		//IL_01e1: Expected O, but got Unknown
		//IL_01e1: Unknown result type (might be due to invalid IL or missing references)
		//IL_01e8: Expected O, but got Unknown
		//IL_0313: Unknown result type (might be due to invalid IL or missing references)
		//IL_031a: Expected O, but got Unknown
		//IL_03b0: Unknown result type (might be due to invalid IL or missing references)
		//IL_03b7: Expected O, but got Unknown
		//IL_042e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0435: Expected O, but got Unknown
		try
		{
			string text = "";
			string[] array = "~K2~K3~J7~E6~E7~D7~R6~X5~R9~E9~".Split(new char[1] { '~' });
			string[] array2 = array;
			foreach (string text2 in array2)
			{
				switch (text2)
				{
				case "K2":
					text = "01_1_TAIN_TT80_DK~KHBS_01_01_TAIN_TT80_DK~KHBS_01_1_01_TAIN_TT80_DK";
					break;
				case "K3":
					text = "01_1_TNDN_TT80_DK~KHBS_01_01_TNDN_TT80_DK~KHBS_01_1_01_TNDN_TT80_DK";
					break;
				case "J7":
					text = "KHBS_01_01_LNCN_PSC_TT80_DK~KHBS_01_1_01_LNCN_PSC_TT80_DK";
					break;
				case "E6":
					text = "02_2_TAIN_TT80_DK~02_1_TAIN_TT80_DK~KHBS_01_02_TAIN_DK_TT80~KHBS_01_1_02_TAIN_DK_TT80";
					break;
				case "E7":
					text = "02_1_TNDN_TT80_DK~KHBS_01_02_TNDN_TT80_DK~KHBS_01_1_02_TNDN_TT80_DK";
					break;
				case "D7":
					text = "02_1_LNCN_PSC_TT80_DK~KHBS_01_02_LNCN_PSC_TT80_DK~KHBS_01_1_02_LNCN_PSC_TT80_DK";
					break;
				case "R9":
					text = "01_1_TK_DK_ND373~KHBS_01_01_TK_DK_ND373~KHBS_01_1_01_TK_DK_ND373";
					break;
				case "E9":
					text = "KHBS_01_01_PTHU_DK_TT80~KHBS_01_1_01_PTHU_DK_TT80";
					break;
				}
				if (text2 == null || text2.Equals(""))
				{
					continue;
				}
				string versionXML = "";
				string versionKHBS = "";
				string dataFileMenu = "";
				string text3 = "";
				GetInfoInMenu(text2, 0, ref versionXML, ref versionKHBS, ref dataFileMenu);
				List<FileInfo> allPathDatafile = GetAllPathDatafile(dataFileMenu, versionXML);
				foreach (FileInfo item in allPathDatafile)
				{
					if (item.Name.Contains("VND") || item.Name.Contains("USD") || item.Name.Contains("KHBS"))
					{
						continue;
					}
					XmlDocument val = new XmlDocument();
					XmlDocument val2 = new XmlDocument();
					val.LoadXml(HTKKFile.LoadFile(item.FullName));
					XmlNodeList val3 = ((XmlNode)val).SelectNodes("/Sections");
					if (text2 != "R9" && val3.Count > 0)
					{
						string value = ((XmlNode)val3[0].Attributes["Version"]).Value;
						value = value.Replace(".", "");
						if (Convert.ToInt32(value) >= 513)
						{
							continue;
						}
					}
					text3 = HTKKFile.GetValueByAttributes(item.FullName, "/Sections/Section/Cells/Cell[@Path='HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ma_DonVi']");
					string text4 = item.FullName.Replace(dataFileMenu, dataFileMenu + "_" + text3);
					string text5 = item.DirectoryName + "\\BK\\" + item.Name;
					if (!File.Exists(text5))
					{
						File.Move(item.FullName, text5);
					}
					else
					{
						File.Delete(item.FullName);
					}
					val.DocumentElement.SetAttribute("Version", WinSettings.VersionApp);
					XmlTextWriter val4 = new XmlTextWriter(text4, Encoding.UTF8);
					try
					{
						val4.Formatting = (Formatting)1;
						val.Save((XmlWriter)(object)val4);
					}
					finally
					{
						((IDisposable)val4)?.Dispose();
					}
					string[] array3 = text.Split(new char[1] { '~' });
					string[] array4 = array3;
					foreach (string text6 in array4)
					{
						string text7 = item.FullName.Replace(dataFileMenu, text6);
						string text8 = text7.Replace(text6, text6 + "_" + text3);
						string text9 = item.Name.Replace(dataFileMenu, text6);
						val2 = new XmlDocument();
						if (File.Exists(text7))
						{
							val2.LoadXml(HTKKFile.LoadFile(text7));
							string text10 = item.DirectoryName + "\\BK\\" + text9;
							if (!File.Exists(text10))
							{
								File.Move(text7, text10);
							}
							else
							{
								File.Delete(text7);
							}
							val2.DocumentElement.SetAttribute("Version", WinSettings.VersionApp);
							val4 = new XmlTextWriter(text8, Encoding.UTF8);
							try
							{
								val4.Formatting = (Formatting)1;
								val2.Save((XmlWriter)(object)val4);
							}
							finally
							{
								((IDisposable)val4)?.Dispose();
							}
						}
					}
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void FixPathStatic(string dataFiles, string versionXml, string pathOld, string pathNew)
	{
		//IL_005d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0064: Expected O, but got Unknown
		//IL_0121: Unknown result type (might be due to invalid IL or missing references)
		//IL_0128: Expected O, but got Unknown
		try
		{
			string[] array = dataFiles.Split(new char[1] { '~' });
			string[] array2 = array;
			foreach (string dataFileMenu in array2)
			{
				List<FileInfo> allPathDatafile = GetAllPathDatafile(dataFileMenu, versionXml);
				foreach (FileInfo item in allPathDatafile)
				{
					XmlDocument val = new XmlDocument();
					val.LoadXml(HTKKFile.LoadFile(item.FullName));
					XmlNodeList source = ((XmlNode)val).SelectNodes("Sections/Section/Cells/Cell");
					XmlNode val2 = ((IEnumerable)source).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode x) => x.Attributes["Path"] != null && ((XmlNode)x.Attributes["Path"]).Value == pathOld));
					if (val2 != null)
					{
						((XmlNode)val2.Attributes["Path"]).Value = pathNew;
						string text = item.DirectoryName + "\\BK\\" + item.Name;
						if (!File.Exists(text))
						{
							Directory.Move(item.FullName, text);
						}
						else
						{
							Directory.Delete(item.FullName);
						}
						XmlTextWriter val3 = new XmlTextWriter(item.FullName, Encoding.UTF8);
						try
						{
							val3.Formatting = (Formatting)1;
							val.Save((XmlWriter)(object)val3);
						}
						finally
						{
							((IDisposable)val3)?.Dispose();
						}
					}
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void ConvertCatalogue_BVMT_CF()
	{
		//IL_0029: Unknown result type (might be due to invalid IL or missing references)
		//IL_002f: Expected O, but got Unknown
		//IL_00dd: Unknown result type (might be due to invalid IL or missing references)
		//IL_00e4: Expected O, but got Unknown
		//IL_0187: Unknown result type (might be due to invalid IL or missing references)
		//IL_018e: Expected O, but got Unknown
		try
		{
			List<FileInfo> allPathDatafile = GetAllPathDatafile("Catalogue_BVMT", "CF");
			foreach (FileInfo item in allPathDatafile)
			{
				XmlDocument val = new XmlDocument();
				val.LoadXml(HTKKFile.LoadFile(item.FullName));
				XmlNodeList val2 = ((XmlNode)val).SelectNodes("/dmconfig/dmcfs/dmcf/cells/cell");
				if (val2 == null || val2.Count == 0)
				{
					continue;
				}
				XmlNode val3 = ((IEnumerable)val2).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode x) => x.Attributes["field"] != null && ((XmlNode)x.Attributes["field"]).Value == "giatoithieu"));
				if (val3 != null)
				{
					continue;
				}
				XmlNode val4 = ((XmlNode)val).SelectSingleNode("/dmconfig/dmcfs/dmcf/cells");
				DirectoryInfo directoryInfo = new DirectoryInfo(HTKKConfig.INTERFACE_INI);
				FileInfo[] files = directoryInfo.GetFiles(item.Name);
				foreach (FileInfo fileInfo in files)
				{
					XmlDocument val5 = new XmlDocument();
					val5.LoadXml(HTKKFile.LoadFile(fileInfo.FullName));
					XmlNode val6 = ((XmlNode)val5).SelectSingleNode("/dmconfig/dmcfs/dmcf/cells");
					if (val6 != null)
					{
						val4.InnerXml = val6.InnerXml;
						break;
					}
				}
				string text = item.DirectoryName + "\\BK\\" + item.Name;
				if (!File.Exists(text))
				{
					Directory.Move(item.FullName, text);
				}
				else
				{
					Directory.Delete(item.FullName);
				}
				XmlTextWriter val7 = new XmlTextWriter(item.FullName, Encoding.UTF8);
				try
				{
					val7.Formatting = (Formatting)1;
					val.Save((XmlWriter)(object)val7);
				}
				finally
				{
					((IDisposable)val7)?.Dispose();
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void RemoveCellById(string dataFiles, string versionXml, string cellIds)
	{
		//IL_0063: Unknown result type (might be due to invalid IL or missing references)
		//IL_006a: Expected O, but got Unknown
		//IL_01ba: Unknown result type (might be due to invalid IL or missing references)
		//IL_01c1: Expected O, but got Unknown
		try
		{
			string[] array = dataFiles.Split(new char[1] { '~' });
			string[] array2 = array;
			foreach (string dataFileMenu in array2)
			{
				List<FileInfo> allPathDatafile = GetAllPathDatafile(dataFileMenu, versionXml);
				foreach (FileInfo item in allPathDatafile)
				{
					string[] array3 = cellIds.Split(new char[1] { ',' });
					XmlDocument val = new XmlDocument();
					val.LoadXml(HTKKFile.LoadFile(item.FullName));
					XmlNodeList val2 = ((XmlNode)val).SelectNodes("/Sections/Section/Cells/Cell");
					if (val2 == null || val2.Count == 0)
					{
						continue;
					}
					bool flag = false;
					string[] array4 = array3;
					foreach (string cellId in array4)
					{
						IEnumerable<XmlNode> source = ((IEnumerable)val2).Cast<XmlNode>();
						Func<XmlNode, bool> predicate = (XmlNode x) => x.Attributes["CellID"] != null && ((XmlNode)x.Attributes["CellID"]).Value == cellId;
						List<XmlNode> list = source.Where(predicate).ToList();
						if (list != null && list.Count > 1)
						{
							list[list.Count() - 1].ParentNode.RemoveChild(list[list.Count() - 1]);
							flag = true;
						}
					}
					if (flag)
					{
						string text = item.DirectoryName + "\\BK\\" + item.Name;
						if (!File.Exists(text))
						{
							Directory.Move(item.FullName, text);
						}
						else
						{
							Directory.Delete(item.FullName);
						}
						XmlTextWriter val3 = new XmlTextWriter(item.FullName, Encoding.UTF8);
						try
						{
							val3.Formatting = (Formatting)1;
							val.Save((XmlWriter)(object)val3);
						}
						finally
						{
							((IDisposable)val3)?.Dispose();
						}
					}
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void FixStt(string dataFiles, string versionXml, string pathStt, int versionConvert = 0)
	{
		//IL_004b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0052: Expected O, but got Unknown
		//IL_0179: Unknown result type (might be due to invalid IL or missing references)
		//IL_0180: Expected O, but got Unknown
		try
		{
			string[] array = dataFiles.Split(new char[1] { '~' });
			string[] array2 = array;
			foreach (string dataFileMenu in array2)
			{
				List<FileInfo> allPathDatafile = GetAllPathDatafile(dataFileMenu, versionXml);
				foreach (FileInfo item in allPathDatafile)
				{
					XmlDocument val = new XmlDocument();
					val.LoadXml(HTKKFile.LoadFile(item.FullName));
					XmlNodeList val2 = ((XmlNode)val).SelectNodes("/Sections");
					if (val2.Count > 0 && versionConvert > 0)
					{
						string value = ((XmlNode)val2[0].Attributes["Version"]).Value;
						value = value.Replace(".", "");
						if (Convert.ToInt32(value) > versionConvert)
						{
							continue;
						}
					}
					XmlNodeList val3 = ((XmlNode)val).SelectNodes("/Sections/Section/RowData");
					if (val3 != null && val3.Count >= 2)
					{
						for (int j = 0; j < val3.Count; j++)
						{
							XmlNode val4 = val3[j];
							List<string> list = GeneralFunctions.SplitString2(val4.InnerText, "###");
							list[0] = string.Concat(j + 1);
							val4.InnerText = GeneralFunctions.ListToString(list, "###");
						}
						XmlTextWriter val5 = new XmlTextWriter(item.FullName, Encoding.UTF8);
						try
						{
							val5.Formatting = (Formatting)1;
							val.Save((XmlWriter)(object)val5);
						}
						finally
						{
							((IDisposable)val5)?.Dispose();
						}
					}
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void FixPhuLucVTNN(string dataFile, string versionXml)
	{
		//IL_0021: Unknown result type (might be due to invalid IL or missing references)
		//IL_0027: Expected O, but got Unknown
		//IL_0082: Unknown result type (might be due to invalid IL or missing references)
		//IL_0089: Expected O, but got Unknown
		//IL_00fa: Unknown result type (might be due to invalid IL or missing references)
		//IL_0101: Expected O, but got Unknown
		//IL_0545: Unknown result type (might be due to invalid IL or missing references)
		//IL_054c: Expected O, but got Unknown
		//IL_0137: Unknown result type (might be due to invalid IL or missing references)
		//IL_013e: Expected O, but got Unknown
		//IL_0228: Unknown result type (might be due to invalid IL or missing references)
		//IL_022f: Expected O, but got Unknown
		try
		{
			List<FileInfo> allPathDatafile = GetAllPathDatafile(dataFile, versionXml);
			foreach (FileInfo item in allPathDatafile)
			{
				XmlDocument val = new XmlDocument();
				val.LoadXml(HTKKFile.LoadFile(item.FullName));
				XmlNodeList val2 = ((XmlNode)val).SelectNodes("/Sections/Section");
				if (val2 == null || val2.Count == 0)
				{
					continue;
				}
				bool flag = false;
				string text = "";
				foreach (XmlNode item2 in val2)
				{
					XmlNode val3 = item2;
					if (val3.Attributes["TaiBangKe"] == null)
					{
						XmlDocument ownerDocument = val3.OwnerDocument;
						XmlAttribute val4 = ownerDocument.CreateAttribute("TaiBangKe");
						((XmlNode)val4).Value = "1";
						val3.Attributes.Append(val4);
						flag = true;
					}
					foreach (XmlNode childNode in val3.ChildNodes)
					{
						XmlNode val5 = childNode;
						if (val5.Name == "Location")
						{
							foreach (XmlNode childNode2 in val5.ChildNodes)
							{
								XmlNode val6 = childNode2;
								if (val6.Name == "DataLocationTaiBangKeRowStart")
								{
									val6.InnerText = "23";
								}
								if (val6.Name == "DataLocationTemplateRowStart")
								{
									text = val6.InnerText;
								}
							}
						}
						if (!(val5.Name == "Header") || val5.ChildNodes.Count != 4)
						{
							continue;
						}
						string value = "";
						string text2 = "";
						string text3 = "";
						string text4 = "";
						string text5 = "";
						foreach (XmlNode childNode3 in val5.ChildNodes)
						{
							XmlNode val7 = childNode3;
							if (((XmlNode)val7.Attributes["Path"]).Value == "ct04")
							{
								text2 = ((XmlNode)val7.Attributes["Value"]).Value;
							}
							if (((XmlNode)val7.Attributes["Path"]).Value == "ct05_ten")
							{
								text3 = ((XmlNode)val7.Attributes["Value"]).Value;
							}
							if (((XmlNode)val7.Attributes["Path"]).Value == "ct05_ma")
							{
								text4 = ((XmlNode)val7.Attributes["Value"]).Value;
							}
							if (((XmlNode)val7.Attributes["Path"]).Value == "ct06")
							{
								text5 = ((XmlNode)val7.Attributes["Value"]).Value;
							}
						}
						if (!text2.IsNullOrEmpty())
						{
							value = text2 + "___" + text3 + "___" + text4 + "___" + text5;
						}
						XmlNode val8 = val5.ChildNodes[0].CloneNode(false);
						if (!text.IsNullOrEmpty())
						{
							((XmlNode)val8.Attributes["CellID"]).Value = "AD_" + (Convert.ToInt32(text) - 1);
							((XmlNode)val8.Attributes["Path"]).Value = "path_AC";
							((XmlNode)val8.Attributes["Value"]).Value = value;
							XmlDocument ownerDocument2 = val8.OwnerDocument;
							XmlAttribute val9 = ownerDocument2.CreateAttribute("BKMaGroupByCol");
							((XmlNode)val9).Value = "Q";
							val5.ChildNodes[0].ParentNode.AppendChild(val8);
						}
					}
				}
				if (flag)
				{
					string text6 = item.DirectoryName + "\\BK\\" + item.Name;
					if (!File.Exists(text6))
					{
						Directory.Move(item.FullName, text6);
					}
					else
					{
						Directory.Delete(text6);
					}
					XmlTextWriter val10 = new XmlTextWriter(item.FullName, Encoding.UTF8);
					try
					{
						val10.Formatting = (Formatting)1;
						val.Save((XmlWriter)(object)val10);
					}
					finally
					{
						((IDisposable)val10)?.Dispose();
					}
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void FixDataLocationTemplateRowStart(XmlNodeList Sections, string tablePath, int khoangCach)
	{
		//IL_0017: Unknown result type (might be due to invalid IL or missing references)
		//IL_001d: Expected O, but got Unknown
		//IL_0075: Unknown result type (might be due to invalid IL or missing references)
		//IL_007b: Expected O, but got Unknown
		//IL_00af: Unknown result type (might be due to invalid IL or missing references)
		//IL_00b5: Expected O, but got Unknown
		try
		{
			foreach (XmlNode Section in Sections)
			{
				XmlNode val = Section;
				if (val.Attributes["TablePath"] == null || !(((XmlNode)val.Attributes["TablePath"]).Value == tablePath))
				{
					continue;
				}
				foreach (XmlNode childNode in val.ChildNodes)
				{
					XmlNode val2 = childNode;
					if (!(val2.Name == "Location"))
					{
						continue;
					}
					foreach (XmlNode childNode2 in val2.ChildNodes)
					{
						XmlNode val3 = childNode2;
						if (val3.Name == "DataLocationTemplateRowStart")
						{
							val3.InnerText = (Convert.ToInt32(val3.InnerText) + khoangCach).ToString();
						}
						if (val3.Name == "TableTemplateRowStart")
						{
							val3.InnerText = (Convert.ToInt32(val3.InnerText) + khoangCach).ToString();
						}
					}
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void FixTangGiamCellTinh(XmlDocument xmlDatafile, int nguong, int khoangCach)
	{
		//IL_0047: Unknown result type (might be due to invalid IL or missing references)
		//IL_004d: Expected O, but got Unknown
		try
		{
			string text = "CellID";
			XmlNodeList val = ((XmlNode)xmlDatafile).SelectNodes("/Sections/Section/Cells/Cell");
			if (val == null || val.Count == 0)
			{
				return;
			}
			foreach (XmlNode item in val)
			{
				XmlNode val2 = item;
				if (val2.Attributes[text] == null || ((XmlNode)val2.Attributes[text]).Value.IndexOf("_") < 0)
				{
					continue;
				}
				string value = ((XmlNode)val2.Attributes[text]).Value;
				string[] array = value.Split(new char[1] { '_' });
				if (array.Length > 1)
				{
					int num = Convert.ToInt32(array[1]);
					if (num > nguong)
					{
						string value2 = array[0] + "_" + (num + khoangCach);
						((XmlNode)val2.Attributes[text]).Value = value2;
					}
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void Fix05QTTSuaDuLieu(string dataFile, string versionXml, string txtCu, string txtMoi)
	{
		//IL_0021: Unknown result type (might be due to invalid IL or missing references)
		//IL_0027: Expected O, but got Unknown
		//IL_0078: Unknown result type (might be due to invalid IL or missing references)
		//IL_007f: Expected O, but got Unknown
		//IL_0143: Unknown result type (might be due to invalid IL or missing references)
		//IL_014a: Expected O, but got Unknown
		try
		{
			List<FileInfo> allPathDatafile = GetAllPathDatafile(dataFile, versionXml);
			foreach (FileInfo item in allPathDatafile)
			{
				XmlDocument val = new XmlDocument();
				val.LoadXml(HTKKFile.LoadFile(item.FullName));
				XmlNodeList val2 = ((XmlNode)val).SelectNodes("/Sections/Section/RowData");
				if (val2 == null || val2.Count == 0)
				{
					continue;
				}
				bool flag = false;
				foreach (XmlNode item2 in val2)
				{
					XmlNode val3 = item2;
					if (val3.InnerText.IndexOf(txtMoi) >= 0)
					{
						break;
					}
					if (val3.InnerText.IndexOf(txtCu) >= 0)
					{
						val3.InnerText = val3.InnerText.Replace(txtCu, txtMoi);
						flag = true;
					}
				}
				if (flag)
				{
					string text = item.DirectoryName + "\\BK\\" + item.Name;
					if (!File.Exists(text))
					{
						Directory.Move(item.FullName, text);
					}
					XmlTextWriter val4 = new XmlTextWriter(item.FullName, Encoding.UTF8);
					try
					{
						val4.Formatting = (Formatting)1;
						val.Save((XmlWriter)(object)val4);
					}
					finally
					{
						((IDisposable)val4)?.Dispose();
					}
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void ReplaceDataFileCommon(string dataFileName, string versionXml, string convertType, int versionApp, string tablePath, string txtOld, string txtNew)
	{
		//IL_0123: Unknown result type (might be due to invalid IL or missing references)
		//IL_012a: Expected O, but got Unknown
		//IL_0185: Unknown result type (might be due to invalid IL or missing references)
		//IL_018c: Expected O, but got Unknown
		//IL_0360: Unknown result type (might be due to invalid IL or missing references)
		//IL_0367: Expected O, but got Unknown
		//IL_01ea: Unknown result type (might be due to invalid IL or missing references)
		//IL_01f1: Expected O, but got Unknown
		string empty = string.Empty;
		string empty2 = string.Empty;
		string empty3 = string.Empty;
		string empty4 = string.Empty;
		string empty5 = string.Empty;
		try
		{
			string[] array = dataFileName.Split(new char[1] { '~' });
			DirectoryInfo directoryInfo = new DirectoryInfo(HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst);
			string path = HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst + "\\BK";
			if (!Directory.Exists(path))
			{
				Directory.CreateDirectory(path);
			}
			string[] array2 = array;
			foreach (string text in array2)
			{
				FileInfo[] files = directoryInfo.GetFiles("*" + text + "*.xml");
				foreach (FileInfo fileInfo in files)
				{
					if (!fileInfo.Name.EndsWith(versionXml + ".xml") || !(convertType == "1"))
					{
						continue;
					}
					empty2 = fileInfo.Name;
					empty4 = fileInfo.DirectoryName + "\\BK\\" + empty2;
					XmlDocument val = new XmlDocument();
					val.LoadXml(HTKKFile.LoadFile(fileInfo.FullName));
					XmlNodeList val2 = ((XmlNode)val).SelectNodes("/Sections/Section");
					if (val2 == null || val2.Count == 0)
					{
						continue;
					}
					bool flag = false;
					foreach (XmlNode item in val2)
					{
						XmlNode val3 = item;
						if (val3.Attributes["TablePath"] == null || ((XmlNode)val3.Attributes["TablePath"]).Value != tablePath)
						{
							continue;
						}
						foreach (XmlNode childNode in val3.ChildNodes)
						{
							XmlNode val4 = childNode;
							if (val4.Name.IndexOf("RowData") >= 0 && val4.InnerText.ToLower().Trim().IndexOf(txtOld.ToLower().Trim()) >= 0 && val4.InnerText.ToLower().Trim().IndexOf(txtNew.ToLower().Trim()) < 0)
							{
								val4.InnerText = val4.InnerText.Replace(txtOld, txtNew);
								flag = true;
							}
						}
					}
					if (flag)
					{
						if (!File.Exists(empty4))
						{
							File.Copy(fileInfo.FullName, empty4, overwrite: true);
						}
						string text2 = HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst + "\\" + empty3;
						XmlTextWriter val5 = new XmlTextWriter(fileInfo.FullName, Encoding.UTF8);
						try
						{
							val5.Formatting = (Formatting)1;
							val.Save((XmlWriter)(object)val5);
						}
						finally
						{
							((IDisposable)val5)?.Dispose();
						}
					}
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			HTKKFile.LogRun(MethodBase.GetCurrentMethod().Name, "Loi replace datafile: " + empty);
		}
	}

	public static void ReplaceDataFileDSNopThayTNCN(string dataFileName)
	{
		//IL_012b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0132: Expected O, but got Unknown
		//IL_0205: Unknown result type (might be due to invalid IL or missing references)
		//IL_020c: Expected O, but got Unknown
		string text = string.Empty;
		string empty = string.Empty;
		string empty2 = string.Empty;
		string empty3 = string.Empty;
		string empty4 = string.Empty;
		try
		{
			string[] array = dataFileName.Split(new char[1] { '~' });
			DirectoryInfo directoryInfo = new DirectoryInfo(HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst);
			string path = HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst + "\\BK";
			if (!Directory.Exists(path))
			{
				Directory.CreateDirectory(path);
			}
			string[] array2 = array;
			foreach (string text2 in array2)
			{
				FileInfo[] files = directoryInfo.GetFiles("*" + text2 + "*.xml");
				foreach (FileInfo fileInfo in files)
				{
					empty = fileInfo.Name;
					text = fileInfo.Name;
					empty4 = empty.Replace(text2, "");
					string[] array3 = empty4.Split(new char[1] { '_' });
					if (array3.Length == 4)
					{
						empty3 = fileInfo.DirectoryName + "\\BK\\" + empty;
						XmlDocument val = new XmlDocument();
						val.LoadXml(HTKKFile.LoadFile(fileInfo.FullName));
						XmlNodeList val2 = ((XmlNode)val).SelectNodes("/Sections/Section");
						string valueByAttributes = HTKKFile.GetValueByAttributes(fileInfo.FullName, "/Sections/Section/Cells/Cell[@Path='HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/Header/soChungTu']");
						empty2 = dataFileName + "_" + array3[1] + "_" + valueByAttributes + "_" + array3[2] + "_" + array3[3];
						string text3 = HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst + "\\" + empty2;
						XmlTextWriter val3 = new XmlTextWriter(text3, Encoding.UTF8);
						try
						{
							val3.Formatting = (Formatting)1;
							val.Save((XmlWriter)(object)val3);
						}
						finally
						{
							((IDisposable)val3)?.Dispose();
						}
						if (File.Exists(empty3))
						{
							File.Delete(empty3);
						}
						if (!File.Exists(empty3))
						{
							File.Move(fileInfo.FullName, empty3);
						}
					}
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			HTKKFile.LogRun(MethodBase.GetCurrentMethod().Name, "Loi replace datafile: " + text);
		}
	}

	public static string TimMenuIdTuDatafile(string dataFile)
	{
		XDocument val = XDocument.Load("Menu.xml");
		XElement val2 = ((XContainer)val).Descendants(XName.op_Implicit("Menu")).FirstOrDefault((Func<XElement, bool>)((XElement m) => ((XContainer)m).Descendants(XName.op_Implicit("Sheet")).Any((XElement s) => (string)s.Attribute(XName.op_Implicit("DataFile")) == dataFile)));
		if (val2 != null)
		{
			return val2.Attribute(XName.op_Implicit("ID")).Value;
		}
		return "";
	}

	public static HTKKCMCSheetInfo GetDataFromDatafileLarge(int sheetIndex, bool getOnlyTemplate)
	{
		//IL_0580: Unknown result type (might be due to invalid IL or missing references)
		//IL_0587: Invalid comparison between Unknown and I4
		//IL_05b5: Unknown result type (might be due to invalid IL or missing references)
		//IL_05bb: Invalid comparison between Unknown and I4
		//IL_05ff: Unknown result type (might be due to invalid IL or missing references)
		//IL_0605: Invalid comparison between Unknown and I4
		//IL_02e3: Unknown result type (might be due to invalid IL or missing references)
		//IL_02e9: Invalid comparison between Unknown and I4
		//IL_0626: Unknown result type (might be due to invalid IL or missing references)
		//IL_062d: Expected O, but got Unknown
		//IL_03bd: Unknown result type (might be due to invalid IL or missing references)
		//IL_03c3: Invalid comparison between Unknown and I4
		//IL_04e2: Unknown result type (might be due to invalid IL or missing references)
		//IL_04e8: Invalid comparison between Unknown and I4
		try
		{
			string sFolder = HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst + "\\";
			string menuChildenAttrById = getMenuChildenAttrById(WinSettings.MenuID, sheetIndex, "DataFile");
			string infaceTemplatePath = HTKKConfig.INTERFACE_TEMPLATES + "\\xml\\" + menuChildenAttrById + "_" + WinSettings.ToKhaiInfo.VersionXML.Replace(".", "") + ".xml";
			string text = WinDoWorkEvent.GetFileNameBySheet(WinSettings.TTPeriodInfo, menuChildenAttrById, sFolder, WinSettings.ToKhaiInfo.VersionXML);
			HTKKCMCSheetInfo hTKKCMCSheetInfo = new HTKKCMCSheetInfo();
			hTKKCMCSheetInfo.GetDetailSheetInfo(infaceTemplatePath);
			hTKKCMCSheetInfo = addInfoMenuToSheetInfo(hTKKCMCSheetInfo, sheetIndex);
			if (WinSettings.TTPeriodInfo.LoaiToKhai == "B" && !File.Exists(text) && !getOnlyTemplate)
			{
				Match match = Regex.Match(text, "bs[^_]*_");
				if (!match.Success)
				{
					return hTKKCMCSheetInfo;
				}
				string numberInString = GeneralFunctions.GetNumberInString(match.Value);
				if (string.IsNullOrEmpty(numberInString))
				{
					return hTKKCMCSheetInfo;
				}
				text = ((!(numberInString == "1")) ? text.Replace("bs" + numberInString + "_", "bs" + (Convert.ToInt32(numberInString) - 1) + "_") : text.Replace("bs" + numberInString + "_", ""));
			}
			if (WinSettings.TTPeriodInfo.LoaiToKhai == "T" && !File.Exists(text))
			{
				Match match = Regex.Match(text, "tt[^_]*_");
				if (!match.Success)
				{
					return hTKKCMCSheetInfo;
				}
				string numberInString = GeneralFunctions.GetNumberInString(match.Value);
				if (string.IsNullOrEmpty(numberInString))
				{
					return hTKKCMCSheetInfo;
				}
				text = ((!(numberInString == "1")) ? text.Replace("tt" + numberInString + "_", "tt" + (Convert.ToInt32(numberInString) - 1) + "_") : text.Replace("tt" + numberInString + "_", ""));
			}
			if (File.Exists(text) && !getOnlyTemplate)
			{
				if (hTKKCMCSheetInfo.Tables.Count == 0)
				{
					return hTKKCMCSheetInfo;
				}
				hTKKCMCSheetInfo.StaticCells.Clear();
				XmlReader val = XmlReader.Create(text);
				bool flag = false;
				bool flag2 = false;
				bool flag3 = false;
				HTKKCMCTableInfo hTKKCMCTableInfo = null;
				while (val.Read())
				{
					if ((int)val.NodeType == 1 && val.Name == "Section")
					{
						flag2 = true;
						flag3 = false;
						string tablePath = val.GetAttribute("TablePath");
						if (!string.IsNullOrEmpty(tablePath))
						{
							hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables.Find((HTKKCMCTableInfo x) => x.TablePath == tablePath);
							if (hTKKCMCTableInfo == null)
							{
								continue;
							}
							hTKKCMCTableInfo.RowsValues.Clear();
						}
						if (hTKKCMCTableInfo != null && tablePath == hTKKCMCTableInfo.TablePath)
						{
							flag3 = true;
						}
					}
					if (flag2 && flag3 && (int)val.NodeType == 1 && hTKKCMCTableInfo != null)
					{
						if (val.Name == "TableTemplateRowStart")
						{
							hTKKCMCTableInfo.Location.TableTemplateRowStart = Convert.ToInt32(val.ReadElementContentAsString());
						}
						if (val.Name == "TableReportRowStart")
						{
							hTKKCMCTableInfo.Location.TableReportRowStart = Convert.ToInt32(val.ReadElementContentAsString());
						}
						if (val.Name == "DataLocationTemplateRowStart")
						{
							hTKKCMCTableInfo.Location.DataLocationTemplateRowStart = Convert.ToInt32(val.ReadElementContentAsString());
						}
						if (val.Name == "DataLocationReportRowStart")
						{
							hTKKCMCTableInfo.Location.DataLocationReportRowStart = Convert.ToInt32(val.ReadElementContentAsString());
						}
						if (val.Name == "DataLocationTaiBangKeRowStart")
						{
							hTKKCMCTableInfo.Location.DataLocationTaiBangKeRowStart = Convert.ToInt32(val.ReadElementContentAsString());
						}
					}
					if (flag2 && flag3 && (int)val.NodeType == 1 && val.Name == "RowData" && hTKKCMCTableInfo != null)
					{
						try
						{
							HTKKCMCRowValue hTKKCMCRowValue = new HTKKCMCRowValue();
							string text2 = val.ReadInnerXml();
							string[] source = text2.Split(new string[1] { "###" }, StringSplitOptions.None);
							hTKKCMCRowValue.RowValue = source.ToList();
							hTKKCMCTableInfo.RowsValues.Add(hTKKCMCRowValue);
						}
						catch (Exception ex)
						{
							HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
							string text3 = val.ReadInnerXml();
						}
					}
					if ((int)val.NodeType == 15 && val.Name == "Section" && hTKKCMCTableInfo != null)
					{
						flag2 = false;
						flag3 = false;
					}
					if ((int)val.NodeType == 1 && val.Name == "Section")
					{
						string attribute = val.GetAttribute("Dynamic");
						flag = attribute == "0";
					}
					if (flag && (int)val.NodeType == 1 && val.Name == "Cell")
					{
						XmlDocument val2 = new XmlDocument();
						XmlNode val3 = val2.ReadNode(val);
						if (val3 != null)
						{
							hTKKCMCSheetInfo.StaticCells.Add(hTKKCMCSheetInfo.GetStaticCellInfo(val3));
						}
					}
				}
				val.Close();
			}
			return hTKKCMCSheetInfo;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public static HTKKCMCSheetInfo addInfoMenuToSheetInfo(HTKKCMCSheetInfo sInfo, int sheetIndex)
	{
		sInfo.SheetName = getMenuChildenAttrById(WinSettings.MenuID, sheetIndex, "SheetName");
		sInfo.TaiDuLieuRatLon = getMenuChildenAttrById(WinSettings.MenuID, sheetIndex, "TaiDuLieuRatLon");
		sInfo.DataFile = getMenuChildenAttrById(WinSettings.MenuID, sheetIndex, "DataFile");
		sInfo.Active = getMenuChildenAttrById(WinSettings.MenuID, sheetIndex, "Active");
		sInfo.XMLnode = getMenuChildenAttrById(WinSettings.MenuID, sheetIndex, "XMLnode");
		sInfo.RowTTDLT = getMenuChildenAttrById(WinSettings.MenuID, sheetIndex, "RowTTDLT");
		return sInfo;
	}

	public static string getMenuHeaderAttrById(string id, string getNameAttr)
	{
		string text = "Menu.xml";
		XDocument val = XDocument.Load(text);
		XElement val2 = ((XContainer)val).Descendants(XName.op_Implicit("Menu")).FirstOrDefault((Func<XElement, bool>)((XElement m) => (string)m.Attribute(XName.op_Implicit("ID")) == id));
		if (val2 != null)
		{
			return (string)val2.Attribute(XName.op_Implicit(getNameAttr));
		}
		return "";
	}

	public static string getMenuChildenAttrById(string id, int sheetIndex, string getNameAttr)
	{
		string text = "Menu.xml";
		XDocument val = XDocument.Load(text);
		XElement val2 = ((XContainer)val).Descendants(XName.op_Implicit("Menu")).FirstOrDefault((Func<XElement, bool>)((XElement m) => (string)m.Attribute(XName.op_Implicit("ID")) == id));
		if (val2 != null)
		{
			List<XElement> list = ((XContainer)val2).Descendants(XName.op_Implicit("Sheet")).ToList();
			if (sheetIndex < list.Count)
			{
				return (string)list[sheetIndex].Attribute(XName.op_Implicit(getNameAttr));
			}
		}
		return "";
	}

	public static void FixStaticToDynamic03TNDN_Tk(int versionApp, string versionXmlNew)
	{
		//IL_0077: Unknown result type (might be due to invalid IL or missing references)
		//IL_007d: Expected O, but got Unknown
		//IL_0139: Unknown result type (might be due to invalid IL or missing references)
		//IL_0140: Expected O, but got Unknown
		//IL_01bd: Unknown result type (might be due to invalid IL or missing references)
		//IL_01c4: Expected O, but got Unknown
		//IL_0305: Unknown result type (might be due to invalid IL or missing references)
		//IL_030c: Expected O, but got Unknown
		//IL_02a7: Unknown result type (might be due to invalid IL or missing references)
		//IL_02ae: Expected O, but got Unknown
		//IL_0484: Unknown result type (might be due to invalid IL or missing references)
		//IL_048b: Expected O, but got Unknown
		//IL_0366: Unknown result type (might be due to invalid IL or missing references)
		//IL_036d: Expected O, but got Unknown
		DirectoryInfo directoryInfo = new DirectoryInfo(HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst);
		FileInfo[] files = directoryInfo.GetFiles("*03_TNDN_TT80*.xml");
		foreach (FileInfo fileInfo in files)
		{
			if (!fileInfo.Name.EndsWith(versionXmlNew + ".xml") || fileInfo.Name.Contains("KHBS"))
			{
				continue;
			}
			XmlDocument val = new XmlDocument();
			val.Load(fileInfo.FullName);
			bool flag = false;
			string valueByAttributes = HTKKFile.GetValueByAttributes(fileInfo.FullName, "/Sections/Section/Cells/Cell[@Path='HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ctC7']");
			if (string.IsNullOrEmpty(valueByAttributes))
			{
				break;
			}
			XmlNodeList val2 = ((XmlNode)val).SelectNodes("/Sections/Section");
			string text = "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/dsThMienThue/thMienThue";
			string valueByAttributes2 = HTKKFile.GetValueByAttributes(fileInfo.FullName, "/Sections/Section/Cells/Cell[@Path='HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ten_loaiMien']");
			string valueByAttributes3 = HTKKFile.GetValueByAttributes(fileInfo.FullName, "/Sections/Section/Cells/Cell[@Path='HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ma_loaiMien']");
			string valueByAttributes4 = HTKKFile.GetValueByAttributes(fileInfo.FullName, "/Sections/Section/Cells/Cell[@Path='HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ctC2']");
			XmlNodeList val3 = ((XmlNode)val).SelectNodes("/Sections/Section/Cells/Cell[@Path='HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ten_loaiMien']");
			if (val3 != null && val3.Count > 0)
			{
				foreach (XmlNode item in val3)
				{
					XmlNode val4 = item;
					val4.ParentNode.RemoveChild(val4);
					flag = true;
				}
			}
			XmlNodeList val5 = ((XmlNode)val).SelectNodes("/Sections/Section/Cells/Cell[@Path='HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ma_loaiMien']");
			if (val5 != null && val5.Count > 0)
			{
				foreach (XmlNode item2 in val5)
				{
					XmlNode val4 = item2;
					val4.ParentNode.RemoveChild(val4);
					flag = true;
				}
			}
			XmlNodeList source = ((XmlNode)val).SelectNodes("Sections/Section/Cells/Cell");
			XmlNode val6 = ((IEnumerable)source).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode x) => x.Attributes["Path"] != null && ((XmlNode)x.Attributes["Path"]).Value == "HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ctC7_thuNhap"));
			if (val6 == null)
			{
				continue;
			}
			((XmlNode)val6.Attributes["Value"]).Value = valueByAttributes;
			XmlNodeList val7 = ((XmlNode)val).SelectNodes("/Sections/Section/Cells/Cell[@Path='HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/ctC7']");
			if (val7 != null && val7.Count > 0)
			{
				foreach (XmlNode item3 in val7)
				{
					XmlNode val4 = item3;
					val4.ParentNode.RemoveChild(val4);
					flag = true;
				}
			}
			foreach (XmlNode item4 in val2)
			{
				XmlNode val8 = item4;
				if (val8.Attributes["TablePath"] == null || !(((XmlNode)val8.Attributes["TablePath"]).Value == text))
				{
					continue;
				}
				foreach (XmlNode childNode in val8.ChildNodes)
				{
					XmlNode val9 = childNode;
					if (val9.Name.Contains("RowData"))
					{
						val9.InnerText = valueByAttributes2 + "###" + valueByAttributes4 + "###" + valueByAttributes3;
						flag = true;
						break;
					}
				}
			}
			if (flag)
			{
				string text2 = fileInfo.DirectoryName + "/BK/du_lieu_" + fileInfo.Name;
				Directory.CreateDirectory(Path.GetDirectoryName(text2));
				if (!File.Exists(text2))
				{
					File.Copy(fileInfo.FullName, text2, overwrite: true);
				}
				XmlTextWriter val10 = new XmlTextWriter(fileInfo.FullName, Encoding.UTF8);
				try
				{
					val10.Formatting = (Formatting)1;
					val.Save((XmlWriter)(object)val10);
				}
				finally
				{
					((IDisposable)val10)?.Dispose();
				}
			}
		}
	}

	public static void ReplaceDataFileStaticCommon(string dataFileName, string versionXml, int versionApp, string path, string Attributes, string valueNew)
	{
		//IL_00ef: Unknown result type (might be due to invalid IL or missing references)
		//IL_00f6: Expected O, but got Unknown
		//IL_026b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0272: Expected O, but got Unknown
		string empty = string.Empty;
		string empty2 = string.Empty;
		string empty3 = string.Empty;
		string empty4 = string.Empty;
		string empty5 = string.Empty;
		try
		{
			DirectoryInfo directoryInfo = new DirectoryInfo(HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst);
			string path2 = HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst + "\\BK";
			if (!Directory.Exists(path2))
			{
				Directory.CreateDirectory(path2);
			}
			FileInfo[] files = directoryInfo.GetFiles("*" + dataFileName + "*.xml");
			foreach (FileInfo fileInfo in files)
			{
				if (!fileInfo.Name.EndsWith(versionXml + ".xml"))
				{
					continue;
				}
				empty2 = fileInfo.Name;
				empty4 = fileInfo.DirectoryName + "\\BK\\" + empty2;
				XmlDocument val = new XmlDocument();
				val.LoadXml(HTKKFile.LoadFile(fileInfo.FullName));
				XmlNode val2 = ((XmlNode)val).SelectSingleNode("/Sections");
				if (val2 != null && versionApp > 0)
				{
					string value = ((XmlNode)val2.Attributes["Version"]).Value.Replace(".", "");
					if (!string.IsNullOrEmpty(value) && Convert.ToInt32(value) >= versionApp)
					{
						continue;
					}
				}
				XmlNodeList val3 = ((XmlNode)val).SelectNodes("/Sections/Section");
				if (val3 == null || val3.Count == 0)
				{
					continue;
				}
				XmlNodeList source = ((XmlNode)val).SelectNodes("Sections/Section/Cells/Cell");
				XmlNode val4 = ((IEnumerable)source).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode x) => x.Attributes["Path"] != null && ((XmlNode)x.Attributes["Path"]).Value == path));
				if (val4 != null)
				{
					((XmlNode)val4.Attributes[Attributes]).Value = valueNew;
					if (!File.Exists(empty4))
					{
						File.Copy(fileInfo.FullName, empty4, overwrite: true);
					}
					string text = HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst + "\\" + empty3;
					XmlTextWriter val5 = new XmlTextWriter(fileInfo.FullName, Encoding.UTF8);
					try
					{
						val5.Formatting = (Formatting)1;
						val.Save((XmlWriter)(object)val5);
					}
					finally
					{
						((IDisposable)val5)?.Dispose();
					}
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			HTKKFile.LogRun(MethodBase.GetCurrentMethod().Name, "Loi replace datafile: " + empty);
		}
	}

	public static void ReplaceDataFileDinamicCommon(string dataFileName, string versionXml, int versionApp, string tablePath, int indexCol, string valueNew)
	{
		//IL_00dd: Unknown result type (might be due to invalid IL or missing references)
		//IL_00e4: Expected O, but got Unknown
		//IL_013f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0146: Expected O, but got Unknown
		//IL_02d0: Unknown result type (might be due to invalid IL or missing references)
		//IL_02d7: Expected O, but got Unknown
		//IL_01a0: Unknown result type (might be due to invalid IL or missing references)
		//IL_01a7: Expected O, but got Unknown
		string empty = string.Empty;
		string empty2 = string.Empty;
		string empty3 = string.Empty;
		string empty4 = string.Empty;
		string empty5 = string.Empty;
		try
		{
			DirectoryInfo directoryInfo = new DirectoryInfo(HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst);
			string path = HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst + "\\BK";
			if (!Directory.Exists(path))
			{
				Directory.CreateDirectory(path);
			}
			FileInfo[] files = directoryInfo.GetFiles("*" + dataFileName + "*.xml");
			foreach (FileInfo fileInfo in files)
			{
				if (!fileInfo.Name.EndsWith(versionXml + ".xml"))
				{
					continue;
				}
				empty2 = fileInfo.Name;
				empty4 = fileInfo.DirectoryName + "\\BK\\" + empty2;
				XmlDocument val = new XmlDocument();
				val.LoadXml(HTKKFile.LoadFile(fileInfo.FullName));
				XmlNodeList val2 = ((XmlNode)val).SelectNodes("/Sections/Section");
				if (val2 == null || val2.Count == 0)
				{
					continue;
				}
				bool flag = false;
				foreach (XmlNode item in val2)
				{
					XmlNode val3 = item;
					if (val3.Attributes["TablePath"] == null || ((XmlNode)val3.Attributes["TablePath"]).Value != tablePath)
					{
						continue;
					}
					foreach (XmlNode childNode in val3.ChildNodes)
					{
						XmlNode val4 = childNode;
						if (val4.Name.IndexOf("RowData") >= 0)
						{
							List<string> list = GeneralFunctions.SplitString2(val4.InnerText, "###");
							list[indexCol] = valueNew;
							val4.InnerText = GeneralFunctions.ListToString(list, "###");
							flag = true;
						}
					}
				}
				if (flag)
				{
					if (!File.Exists(empty4))
					{
						File.Copy(fileInfo.FullName, empty4, overwrite: true);
					}
					string text = HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst + "\\" + empty3;
					XmlTextWriter val5 = new XmlTextWriter(fileInfo.FullName, Encoding.UTF8);
					try
					{
						val5.Formatting = (Formatting)1;
						val.Save((XmlWriter)(object)val5);
					}
					finally
					{
						((IDisposable)val5)?.Dispose();
					}
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			HTKKFile.LogRun(MethodBase.GetCurrentMethod().Name, "Loi replace datafile: " + empty);
		}
	}
}
