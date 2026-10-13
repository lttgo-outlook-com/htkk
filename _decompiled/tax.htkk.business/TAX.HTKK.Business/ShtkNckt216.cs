using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Control;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class ShtkNckt216
{
	private const string HO_CHIEU = "5";

	private const string CMND = "3";

	private const string CCCD = "4";

	private const string CA_NHAN = "02";

	private const string TO_CHUC = "01";

	private int idxColC = HTKKFile.ColumnLettersToIndex("C");

	private int idxColD = HTKKFile.ColumnLettersToIndex("D");

	private int idxColF = HTKKFile.ColumnLettersToIndex("F");

	private int idxColG = HTKKFile.ColumnLettersToIndex("G");

	private int idxColH = HTKKFile.ColumnLettersToIndex("H");

	private int idxColI = HTKKFile.ColumnLettersToIndex("I");

	private int idxColK = HTKKFile.ColumnLettersToIndex("K");

	private int idxColM = HTKKFile.ColumnLettersToIndex("M");

	private int idxColN = HTKKFile.ColumnLettersToIndex("N");

	private int idxColP = HTKKFile.ColumnLettersToIndex("P");

	private int idxColQ = HTKKFile.ColumnLettersToIndex("Q");

	private int idxColR = HTKKFile.ColumnLettersToIndex("R");

	private int idxColS = HTKKFile.ColumnLettersToIndex("S");

	private int idxTableColC = 1;

	private int idxTableColD = 2;

	private int idxTableColF = 3;

	private int idxTableColH = 4;

	private int idxTableColI = 5;

	private int idxTableColK = 6;

	private int idxTableColM = 7;

	private int idxTableColN = 8;

	private int idxTableColO = 9;

	private int idxTableColP = 10;

	public void AutoBindData_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			int[] array = GeneralFunctions.ArrayCellId("E_38");
			SheetView sheetView = winFpSpread.Sheets[0];
			if (PeriodInfo.LanXuatBan != "")
			{
				sheetView.Cells[array[1], array[0]].Text = PeriodInfo.LanXuatBan;
			}
			int[] array2 = GeneralFunctions.ArrayCellId("O_36");
			if (PeriodInfo.KieuKy == "M")
			{
				sheetView.Cells[array2[1], array2[0]].Text = "DK";
			}
			else
			{
				sheetView.Cells[array2[1], array2[0]].Text = "LD";
			}
			int[] array3 = GeneralFunctions.ArrayCellId("P_36");
			sheetView.Cells[array3[1], array3[0]].Text = PeriodInfo.kyKKhai;
			int[] array4 = GeneralFunctions.ArrayCellId("E_40");
			if (string.IsNullOrEmpty(sheetView.Cells[array4[1], array4[0]].Text))
			{
				sheetView.Cells[array4[1], array4[0]].Text = GetMaNH(TTinChung.mst);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void CellChange_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void ButtonClick_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, int idxRow, int idxCol)
	{
		try
		{
			if (idxRow == 0 && idxCol == 2)
			{
				ShowBangKe("Bang_Ke_01_SHTK.xls");
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	private void ShowBangKe(string sFileName)
	{
		//IL_005e: Unknown result type (might be due to invalid IL or missing references)
		try
		{
			string text = Path.GetTempPath() + sFileName;
			File.Copy(HTKKConfig.INTERFACE_TEMPLATES + "\\excel\\" + sFileName, text, overwrite: true);
			Process.Start(text);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			frmMessage frmMessage = new frmMessage("0433", "3", "", "Thông báo");
			((Form)frmMessage).ShowDialog();
		}
	}

	public bool isValidDuLieuTK(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			HTKKFile.deleteLogErrorBigData();
			bool result = true;
			Validate_0(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
			foreach (KeyValuePair<int, int> item in levelWarning)
			{
				int key = item.Key;
				switch (item.Value)
				{
				case 2:
					winFpSpread.TabStrip[key].BackColor = Color.Pink;
					break;
				case 1:
					if (winFpSpread.TabStrip[key].BackColor == Color.White)
					{
						winFpSpread.TabStrip[key].BackColor = Color.Yellow;
					}
					break;
				}
			}
			return result;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void Validate_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			int value = 0;
			levelWarning.TryGetValue(0, out value);
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			int num = HTKKFile.ColumnLettersToIndex("T");
			int[] array = GeneralFunctions.ArrayCellId("E_14");
			string text = sheetView.Cells[array[1], array[0]].Text;
			DateTime dateTime = ConvertHelpers.ConvertToDateTime(text, "dd/MM/yyyy");
			if (dateTime > DateTime.MinValue)
			{
				HTKKFile.RemoveValidateBackGround1(sheetView, array[1], array[0]);
			}
			string text2 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[14, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text3 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[15, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text4 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[16, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text5 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[17, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text6 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[20, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text7 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[21, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text8 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[22, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text9 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[24, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text10 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[23, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text11 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[25, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text12 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[26, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text13 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[32, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text14 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[35, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text15 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[36, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text16 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[37, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text17 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[38, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text18 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[39, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			int[] array2 = GeneralFunctions.ArrayCellId("E_40");
			string text19 = sheetView.Cells[array2[1], array2[0]].Text;
			HTKKFile.RemoveValidate(sheetView, array2[1], array2[0]);
			if (string.IsNullOrEmpty(text19))
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, array2[1], array2[0], text2);
				value = 2;
			}
			else if (text19.Length != 8 || !GeneralFunctions.IsNumeric(text19))
			{
				HTKKFile.SetNodeBackGroundPink(sheetView, array2[1], array2[0], text14);
				value = 2;
			}
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			HTKKCMCDataLocation location = hTKKCMCTableInfo.Location;
			int num2 = location.DataLocationTemplateRowStart - 1;
			List<HTKKCMCRowValue> rowsValues = hTKKCMCTableInfo.RowsValues;
			HTKKCMCRowInfo rowsInfo = hTKKCMCTableInfo.RowsInfo;
			Dictionary<string, string> dictionary = new Dictionary<string, string>();
			for (int i = 0; i < rowsValues.Count; i++)
			{
				string[] array3 = rowsValues[i].RowValue.ToArray();
				HTKKFile.RemoveValidate(sheetView, num2 + i, idxColC - 1);
				HTKKFile.RemoveValidate(sheetView, num2 + i, idxColD - 1);
				HTKKFile.RemoveValidate(sheetView, num2 + i, idxColF - 1);
				HTKKFile.RemoveValidate(sheetView, num2 + i, idxColH - 1);
				HTKKFile.RemoveValidate(sheetView, num2 + i, idxColI - 1);
				HTKKFile.RemoveValidate(sheetView, num2 + i, idxColK - 1);
				HTKKFile.RemoveValidate(sheetView, num2 + i, idxColM - 1);
				HTKKFile.RemoveValidate(sheetView, num2 + i, idxColN - 1);
				if (Convert.ToInt32(sheetView.Cells[num2 + i, num - 1].Value) == 0)
				{
					continue;
				}
				bool flag = false;
				bool flag2 = false;
				string text20 = array3[idxTableColF];
				string value2 = array3[idxTableColD];
				string text21 = array3[idxTableColC];
				if (string.IsNullOrEmpty(text21))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxColC - 1, text2, i + 1);
					value = 2;
				}
				else if (!ValidateBase.CheckMaSoThueCCCD(text21))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxColC - 1, text5, i + 1);
					value = 2;
				}
				string text22 = array3[idxTableColO];
				if (string.IsNullOrEmpty(value2) && !string.IsNullOrEmpty(text20))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxColD - 1, text2, i + 1);
					value = 2;
				}
				if (!string.IsNullOrEmpty(value2) && string.IsNullOrEmpty(text22))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxColD - 1, text11, i + 1);
					value = 2;
				}
				else
				{
					string text23 = array3[idxTableColP];
					if (text23.Equals("01"))
					{
						if (text22 == "5" || text22 == "3" || text22 == "4")
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxColD - 1, text15, i + 1);
							value = 2;
						}
					}
					else if (text23.Equals("02") && !(text22 == "5") && !(text22 == "3") && !(text22 == "4"))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxColD - 1, text16, i + 1);
						value = 2;
					}
				}
				if (!string.IsNullOrEmpty(value2) && string.IsNullOrEmpty(text20))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxColF - 1, text2, i + 1);
					value = 2;
				}
				else if (!string.IsNullOrEmpty(text20))
				{
					if (text22.Equals("3"))
					{
						if (text20.Length != 9 && text20.Length != 12)
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxColF - 1, text6, i + 1);
							value = 2;
						}
						else if (!HTKKFormater.IsNumber(text20))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxColF - 1, text8, i + 1);
							value = 2;
						}
					}
					else if (text22.Equals("4"))
					{
						if (text20.Length != 12)
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxColF - 1, text7, i + 1);
							value = 2;
						}
						else if (!HTKKFormater.IsNumber(text20))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxColF - 1, text8, i + 1);
							value = 2;
						}
					}
					else if (text22.Equals("5") && text20.Length > 20)
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxColF - 1, text17, i + 1);
						value = 2;
					}
				}
				string value3 = array3[idxTableColH];
				if (string.IsNullOrEmpty(value3))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxColH - 1, text2, i + 1);
					value = 2;
				}
				string value4 = array3[idxTableColI];
				if (string.IsNullOrEmpty(value4))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxColI - 1, text2, i + 1);
					value = 2;
				}
				bool flag3 = true;
				string text24 = array3[idxTableColK];
				DateTime dateTime2 = ConvertHelpers.ConvertToDateTime(text24, "dd/MM/yyyy");
				if (string.IsNullOrEmpty(text24))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxColK - 1, text2, i + 1);
					value = 2;
					flag3 = false;
				}
				else if (dateTime2 == DateTime.MinValue)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxColK - 1, text3, i + 1);
					value = 2;
					flag3 = false;
				}
				else if (dateTime2 > DateTime.Now)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxColK - 1, text4, i + 1);
					value = 2;
					flag3 = false;
				}
				string text25 = array3[idxTableColM];
				DateTime dateTime3 = ConvertHelpers.ConvertToDateTime(text25, "dd/MM/yyyy");
				if (string.IsNullOrEmpty(text25))
				{
					flag3 = false;
				}
				else if (dateTime3 == DateTime.MinValue)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxColM - 1, text3, i + 1);
					value = 2;
					flag3 = false;
				}
				else if (dateTime3 > DateTime.Now)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxColM - 1, text4, i + 1);
					value = 2;
					flag3 = false;
				}
				if (flag3 && dateTime3.CompareTo(dateTime2) <= 0)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxColK - 1, text13, i + 1);
					HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxColM - 1, text13, i + 1);
					value = 2;
				}
				string value5 = array3[idxTableColN];
				string value6 = array3[idxTableColP];
				if (string.IsNullOrEmpty(value5))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxColN - 1, text2, i + 1);
					value = 2;
				}
				else if (string.IsNullOrEmpty(value6))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num2 + i, idxColN - 1, text11, i + 1);
					value = 2;
				}
			}
			switch (value)
			{
			case 2:
				sheetView.FpSpread.TabStrip[0].BackColor = Color.Pink;
				break;
			case 1:
				sheetView.FpSpread.TabStrip[0].BackColor = Color.Yellow;
				break;
			default:
				sheetView.FpSpread.TabStrip[0].BackColor = Color.White;
				break;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void ExportXML_0(HTKKPeriodInfo PeriodInfo, XmlNode xmlNode)
	{
		//IL_0053: Unknown result type (might be due to invalid IL or missing references)
		//IL_0059: Expected O, but got Unknown
		//IL_0176: Unknown result type (might be due to invalid IL or missing references)
		//IL_017c: Expected O, but got Unknown
		//IL_026a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0270: Expected O, but got Unknown
		try
		{
			string path = "HSoKhaiThue/CTieuTKhaiChinh/NOIDUNG_DANHSACH_CTIET/ROW_CTIET/SOHIEU_TK";
			setValueMaHoa(xmlNode, path);
			string text = "HSoKhaiThue/CTieuTKhaiChinh/NOIDUNG_DANHSACH_CTIET/ROW_CTIET/NGAY_MO";
			XmlNodeList val = null;
			val = xmlNode.SelectNodes(text);
			if (val != null && val.Count > 0)
			{
				foreach (XmlNode item in val)
				{
					XmlNode val2 = item;
					string text2 = val2.InnerText.Trim();
					if (text2.Length == 10)
					{
						string[] array = text2.Split(new char[1] { '-' });
						val2.InnerText = array[2] + "/" + array[1] + "/" + array[0];
					}
					if (string.IsNullOrEmpty(text2))
					{
						val2.Attributes.Remove(val2.Attributes["xsi:nil"]);
					}
				}
			}
			string text3 = "HSoKhaiThue/CTieuTKhaiChinh/NOIDUNG_DANHSACH_CTIET/ROW_CTIET/NGAY_DONG";
			val = xmlNode.SelectNodes(text3);
			if (val != null && val.Count > 0)
			{
				foreach (XmlNode item2 in val)
				{
					XmlNode val2 = item2;
					string text2 = val2.InnerText.Trim();
					if (text2.Length == 10)
					{
						string[] array = text2.Split(new char[1] { '-' });
						val2.InnerText = array[2] + "/" + array[1] + "/" + array[0];
					}
				}
			}
			string text4 = "HSoKhaiThue/CTieuTKhaiChinh/NOIDUNG_DANHSACH_CTIET/ROW_CTIET/MST";
			val = xmlNode.SelectNodes(text4);
			if (val == null || val.Count <= 0)
			{
				return;
			}
			foreach (XmlNode item3 in val)
			{
				XmlNode val2 = item3;
				string text2 = val2.InnerText.Trim();
				if (string.IsNullOrEmpty(text2))
				{
					val2.Attributes.Remove(val2.Attributes["xsi:nil"]);
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void setValueMaHoa(XmlNode xmlNode, string path)
	{
		//IL_003e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0044: Expected O, but got Unknown
		try
		{
			XmlNodeList val = null;
			string pathCert = "..\\Pictures\\TCT_Tk.cer";
			val = xmlNode.SelectNodes(path);
			if (val == null || val.Count <= 0)
			{
				return;
			}
			foreach (XmlNode item in val)
			{
				XmlNode val2 = item;
				if (!string.IsNullOrEmpty(val2.InnerText.Trim()))
				{
					string innerText = HTKKSecurity.MaHoaDuLieu(val2.InnerText.Trim(), pathCert);
					val2.InnerText = innerText;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public string GetMaNH(string mst)
	{
		//IL_0051: Unknown result type (might be due to invalid IL or missing references)
		//IL_0058: Expected O, but got Unknown
		string result = "";
		XmlDataDocument xmlDataDocument = new XmlDataDocument();
		FileStream fileStream = new FileStream(HTKKConfig.INTERFACE_INI + "\\Catalogue_DM_MA_NH.xml", FileMode.Open, FileAccess.Read);
		((XmlDocument)xmlDataDocument).Load((Stream)fileStream);
		XmlNodeList elementsByTagName = ((XmlDocument)xmlDataDocument).GetElementsByTagName("Item");
		Dictionary<string, string> dictionary = new Dictionary<string, string>();
		foreach (XmlNode item in elementsByTagName)
		{
			XmlNode val = item;
			string value = ((XmlNode)val.Attributes["Value"]).Value;
			string[] array = Regex.Split(value, "###");
			if (array.Length >= 3 && mst == array[0])
			{
				result = array[1];
				break;
			}
		}
		return result;
	}

	public void FinishImportExcel(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo)
	{
		GeneralFunctions.fpSpReindexingTableTemplate(winFpSpread.Sheets[0], ToKhaiInfo.SheetInfos[0].Tables[0]);
	}

	public void FinishImportXML(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			int num = HTKKFile.ColumnLettersToIndex("D");
			int num2 = HTKKFile.ColumnLettersToIndex("O");
			int num3 = HTKKFile.ColumnLettersToIndex("N");
			int num4 = HTKKFile.ColumnLettersToIndex("P");
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				Dictionary<string, string> dMFromCombobox = GeneralFunctions.GetDMFromCombobox(sheetView, dataLocationTemplateRowStart + i, num);
				Dictionary<string, string> dMFromCombobox2 = GeneralFunctions.GetDMFromCombobox(sheetView, dataLocationTemplateRowStart + i, num3);
				string text = sheetView.Cells[dataLocationTemplateRowStart + i - 1, num2 - 1].Text.Trim();
				if (!string.IsNullOrEmpty(text) && dMFromCombobox.ContainsKey(text))
				{
					sheetView.Cells[dataLocationTemplateRowStart + i - 1, num - 1].Text = dMFromCombobox[text];
				}
				string text2 = sheetView.Cells[dataLocationTemplateRowStart + i - 1, num4 - 1].Text.Trim();
				if (!string.IsNullOrEmpty(text2) && dMFromCombobox2.ContainsKey(text2))
				{
					sheetView.Cells[dataLocationTemplateRowStart + i - 1, num3 - 1].Text = dMFromCombobox2[text2];
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}
}
