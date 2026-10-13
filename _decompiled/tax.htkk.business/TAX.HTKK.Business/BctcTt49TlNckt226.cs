using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using System.Xml;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Control;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class BctcTt49TlNckt226
{
	private HTKKCMCToKhaiInfo tKhaiInfoOld = new HTKKCMCToKhaiInfo();

	public void ExportXML_0(HTKKPeriodInfo PeriodInfo, XmlNode xmlNode)
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

	public void ExportXML_1(HTKKPeriodInfo PeriodInfo, XmlNode xmlNode)
	{
		try
		{
			XmlNode val = xmlNode.SelectSingleNode("PL_KQHDSXKD/PhanKy/ngayLap");
			if (val != null)
			{
				string innerText = val.InnerText;
				if (innerText.Length >= 8)
				{
					val.InnerText = innerText.Substring(4, 4) + "-" + innerText.Substring(2, 1).PadLeft(2, '0') + "-" + innerText.Substring(0, 1).PadLeft(2, '0');
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void ExportXML_2(HTKKPeriodInfo PeriodInfo, XmlNode xmlNode)
	{
		try
		{
			XmlNode val = xmlNode.SelectSingleNode("PL_LCTTTT/PhanKy/ngayLap");
			if (val != null)
			{
				string innerText = val.InnerText;
				if (innerText.Length >= 8)
				{
					val.InnerText = innerText.Substring(4, 4) + "-" + innerText.Substring(2, 1).PadLeft(2, '0') + "-" + innerText.Substring(0, 1).PadLeft(2, '0');
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public bool ValidatePeriod_ALL(HTKKPeriodInfo PeriodInfo)
	{
		bool result = true;
		try
		{
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			result = false;
		}
		return result;
	}

	public bool ValidatePeriod_axfpsPhuLuc(FpSpread winFpSpread)
	{
		//IL_0095: Unknown result type (might be due to invalid IL or missing references)
		bool result = true;
		try
		{
			if ((bool)winFpSpread.Sheets[0].Cells[2, 0].Value == (bool)winFpSpread.Sheets[0].Cells[1, 0].Value && (bool)winFpSpread.Sheets[0].Cells[1, 0].Value)
			{
				frmMessage frmMessage = new frmMessage("0314", "2", "", "Thông báo");
				((Form)frmMessage).ShowDialog();
				return frmMessage.Yes;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			result = false;
		}
		return result;
	}

	public void AutoBindData_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
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

	public void AutoBindData_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			string text = HTKKConfig.DATAFILES + "\\" + TTinChung.mst + "\\" + ToKhaiInfo.SheetInfos[1].DataFile;
			int num = int.Parse(PeriodInfo.Nam);
			text = text + "_" + num + "_" + PeriodInfo.TuNgay.Split(new char[1] { '/' })[0] + PeriodInfo.TuNgay.Split(new char[1] { '/' })[1] + num + "_" + PeriodInfo.DenNgay.Split(new char[1] { '/' })[0] + PeriodInfo.DenNgay.Split(new char[1] { '/' })[1] + num;
			text = text + "_" + ToKhaiInfo.VersionXML.Replace(".", "").Replace("_", "") + ".xml";
			bool flag = false;
			if (File.Exists(text))
			{
				flag = true;
			}
			SheetView sView = winFpSpread.Sheets[1];
			int[] array = GeneralFunctions.ArrayCellId("B_6");
			string valueCustom = GeneralFunctions.GetValueCustom(sView, array[1] + 1, array[0] + 1, "2");
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void AutoBindData_2(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			string text = HTKKConfig.DATAFILES + "\\" + TTinChung.mst + "\\" + ToKhaiInfo.SheetInfos[2].DataFile;
			int num = int.Parse(PeriodInfo.Nam);
			text = text + "_" + num + "_" + PeriodInfo.TuNgay.Split(new char[1] { '/' })[0] + PeriodInfo.TuNgay.Split(new char[1] { '/' })[1] + num + "_" + PeriodInfo.DenNgay.Split(new char[1] { '/' })[0] + PeriodInfo.DenNgay.Split(new char[1] { '/' })[1] + num;
			text = text + "_" + ToKhaiInfo.VersionXML.Replace(".", "").Replace("_", "") + ".xml";
			bool flag = false;
			if (File.Exists(text))
			{
				flag = true;
			}
			SheetView sView = winFpSpread.Sheets[2];
			int[] array = GeneralFunctions.ArrayCellId("B_35");
			string valueCustom = GeneralFunctions.GetValueCustom(sView, array[1] + 1, array[0] + 1, "2");
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void ButtonClick_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, long idxRow, long idxCol)
	{
		try
		{
			SheetView sView = winFpSpread.Sheets[0];
			int[] array = GeneralFunctions.ArrayCellId("B_10");
			string valueCustom = GeneralFunctions.GetValueCustom(sView, array[1] + 1, array[0] + 1, "2");
			if (idxCol == array[0] && idxRow == array[1])
			{
				getSoKyTruoc(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo, 0, isReadDataFile: false, Convert.ToBoolean(valueCustom), callIsCheck: true);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void ButtonClick_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, long idxRow, long idxCol)
	{
		try
		{
			SheetView sView = winFpSpread.Sheets[1];
			int[] array = GeneralFunctions.ArrayCellId("B_6");
			string valueCustom = GeneralFunctions.GetValueCustom(sView, array[1] + 1, array[0] + 1, "2");
			if (idxCol == array[0] && idxRow == array[1])
			{
				getSoKyTruoc(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo, 1, isReadDataFile: false, Convert.ToBoolean(valueCustom), callIsCheck: true);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void ButtonClick_2(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, long idxRow, long idxCol)
	{
		try
		{
			SheetView sView = winFpSpread.Sheets[2];
			int[] array = GeneralFunctions.ArrayCellId("B_35");
			string valueCustom = GeneralFunctions.GetValueCustom(sView, array[1] + 1, array[0] + 1, "2");
			if (idxCol == array[0] && idxRow == array[1])
			{
				getSoKyTruoc(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo, 2, isReadDataFile: false, Convert.ToBoolean(valueCustom), callIsCheck: true);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public static void CellChange_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
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

	public static void CellChange_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			string text = "I###+R###";
			if ((int)iCol == HTKKFile.ColumnLettersToIndex("I"))
			{
				winFpSpread.ActiveSheet.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("M") - 1].Formula = text.Replace("###", iRow.ToString());
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public static void CellChange_2(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
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

	public void isValidDuLieuTK(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			if (ToKhaiInfo.SheetInfos[0] != null)
			{
				Validate_0(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo);
			}
			if (ToKhaiInfo.SheetInfos[1] != null)
			{
				Validate_1(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo);
			}
			if (ToKhaiInfo.SheetInfos[2] != null)
			{
				Validate_2(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo);
			}
			foreach (KeyValuePair<int, int> item in levelWarning)
			{
				int key = item.Key;
				int value = item.Value;
				if ((value != 2 || !(winFpSpread.TabStrip[key].BackColor == Color.Pink)) && (value != 1 || !(winFpSpread.TabStrip[key].BackColor == Color.Yellow)))
				{
					if (value == 2 && winFpSpread.TabStrip[key].BackColor != Color.Pink)
					{
						winFpSpread.TabStrip[key].BackColor = Color.Pink;
					}
					else if (value == 1 && winFpSpread.TabStrip[key].BackColor == Color.White)
					{
						winFpSpread.TabStrip[key].BackColor = Color.Yellow;
					}
					else if (value == 0)
					{
						winFpSpread.TabStrip[key].BackColor = Color.White;
					}
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void Validate_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
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

	public void Validate_1(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			SheetView sView = winFpSpread.Sheets[1];
			int[] array = GeneralFunctions.ArrayCellId("B_6");
			string valueCustom = GeneralFunctions.GetValueCustom(sView, array[1] + 1, array[0] + 1, "2");
			checkSoNamTruoc(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo, 1, Convert.ToBoolean(valueCustom));
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void Validate_2(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			SheetView sView = winFpSpread.Sheets[2];
			int[] array = GeneralFunctions.ArrayCellId("B_35");
			string valueCustom = GeneralFunctions.GetValueCustom(sView, array[1] + 1, array[0] + 1, "2");
			checkSoNamTruoc(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo, 2, Convert.ToBoolean(valueCustom));
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	private void getSoKyTruoc(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, int sheetIndex, bool isReadDataFile, bool isCheckLayDl, bool callIsCheck)
	{
		try
		{
			List<string> pathDataFileOld = getPathDataFileOld(ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo, sheetIndex);
			string text = pathDataFileOld[0];
			string text2 = pathDataFileOld[1];
			string text3 = pathDataFileOld[2];
			string text4 = text3;
			SheetView sheetView = winFpSpread.Sheets[sheetIndex];
			int num = 0;
			int num2 = 0;
			int num3 = 0;
			int num4 = 0;
			int num5 = 0;
			int num6 = 0;
			int num7 = 0;
			int num8 = 0;
			string text5 = "";
			if (sheetIndex == 0)
			{
				num = HTKKFile.ColumnLettersToIndex("R");
				num2 = HTKKFile.ColumnLettersToIndex("W");
				num3 = HTKKFile.ColumnLettersToIndex("AC");
			}
			if (sheetIndex == 2 || sheetIndex == 3)
			{
				text4 = text;
				num = HTKKFile.ColumnLettersToIndex("R");
				num2 = HTKKFile.ColumnLettersToIndex("W");
				num3 = HTKKFile.ColumnLettersToIndex("AC");
			}
			if (sheetIndex == 1)
			{
				text4 = text;
				num = HTKKFile.ColumnLettersToIndex("I");
				num2 = HTKKFile.ColumnLettersToIndex("J");
				num3 = HTKKFile.ColumnLettersToIndex("P");
				num4 = HTKKFile.ColumnLettersToIndex("M");
				num5 = HTKKFile.ColumnLettersToIndex("R");
				num7 = HTKKFile.ColumnLettersToIndex("N");
				num8 = HTKKFile.ColumnLettersToIndex("S");
				text5 = "I###+R###";
			}
			if (File.Exists(text4) && isCheckLayDl)
			{
				HTKKCMCSheetInfo hTKKCMCSheetInfo = new HTKKCMCSheetInfo();
				hTKKCMCSheetInfo.GetDetailSheetInfo(text4);
				for (int i = 1; i < hTKKCMCSheetInfo.StaticCells.Count; i++)
				{
					if (hTKKCMCSheetInfo.StaticCells[i].CellTemplate.ColIndex == num)
					{
						sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num3 - 1].Value = ConvertHelpers.ConvertToDouble(hTKKCMCSheetInfo.StaticCells[i].Value);
						if (!isReadDataFile)
						{
							sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num2 - 1].Value = ConvertHelpers.ConvertToDouble(hTKKCMCSheetInfo.StaticCells[i].Value);
						}
					}
				}
			}
			else
			{
				for (int i = 1; i < ToKhaiInfo.SheetInfos[sheetIndex].StaticCells.Count; i++)
				{
					if (ToKhaiInfo.SheetInfos[sheetIndex].StaticCells[i].CellTemplate.ColIndex == num && !isReadDataFile)
					{
						sheetView.Cells[ToKhaiInfo.SheetInfos[sheetIndex].StaticCells[i].CellTemplate.RowIndex - 1, num2 - 1].Value = 0;
					}
				}
			}
			if (File.Exists(text))
			{
				HTKKCMCSheetInfo hTKKCMCSheetInfo = new HTKKCMCSheetInfo();
				hTKKCMCSheetInfo.GetDetailSheetInfo(text);
				for (int i = 1; i < hTKKCMCSheetInfo.StaticCells.Count; i++)
				{
					if (sheetIndex == 1 && hTKKCMCSheetInfo.StaticCells[i].CellTemplate.ColIndex == num4)
					{
						sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num8 - 1].Value = ConvertHelpers.ConvertToDouble(hTKKCMCSheetInfo.StaticCells[i].Value);
						if (!isReadDataFile)
						{
							sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num7 - 1].Value = ConvertHelpers.ConvertToDouble(hTKKCMCSheetInfo.StaticCells[i].Value);
						}
					}
				}
				return;
			}
			for (int i = 1; i < ToKhaiInfo.SheetInfos[sheetIndex].StaticCells.Count; i++)
			{
				if (sheetIndex == 1 && ToKhaiInfo.SheetInfos[sheetIndex].StaticCells[i].CellTemplate.ColIndex == num4 && !isReadDataFile)
				{
					sheetView.Cells[ToKhaiInfo.SheetInfos[sheetIndex].StaticCells[i].CellTemplate.RowIndex - 1, num7 - 1].Value = 0;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private List<string> getPathDataFileOld(HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, int sheetIndex)
	{
		List<string> list = new List<string>();
		try
		{
			string text = HTKKConfig.DATAFILES + "\\" + TTinChung.mst + "\\" + ToKhaiInfo.SheetInfos[sheetIndex].DataFile;
			string item = text;
			string item2 = text;
			string text2 = "";
			string text3 = "";
			if (HTKKFormater.isNullString(PeriodInfo.TuNgay) || HTKKFormater.isNullString(PeriodInfo.DenNgay))
			{
				list.Add(text);
				return list;
			}
			text2 = PeriodInfo.TuNgay.Split(new char[1] { '/' })[0] + PeriodInfo.TuNgay.Split(new char[1] { '/' })[1] + (ConvertHelpers.ConvertToInt(PeriodInfo.TuNgay.Split(new char[1] { '/' })[2]) - 1);
			text3 = PeriodInfo.DenNgay.Split(new char[1] { '/' })[0] + PeriodInfo.DenNgay.Split(new char[1] { '/' })[1] + (ConvertHelpers.ConvertToInt(PeriodInfo.DenNgay.Split(new char[1] { '/' })[2]) - 1);
			string text4 = HTKKConfig.DATAFILES + "\\" + HeaderInfo.mst;
			DirectoryInfo directoryInfo = new DirectoryInfo(text4);
			if (sheetIndex == 0)
			{
				string value = ToKhaiInfo.SheetInfos[sheetIndex].DataFile.Replace("_DD", "").Replace("_TL", "") + "_" + (ConvertHelpers.ConvertToInt(PeriodInfo.Nam) - 1) + "_";
				if (directoryInfo.Exists)
				{
					FileInfo[] files = directoryInfo.GetFiles();
					if (files.Length > 0)
					{
						FileInfo[] array = files;
						foreach (FileInfo fileInfo in array)
						{
							if (fileInfo.Name.Contains(value))
							{
								item2 = text4 + "\\" + fileInfo.Name;
								break;
							}
						}
					}
				}
			}
			if (sheetIndex == 1 || sheetIndex == 2)
			{
				text = text + "_Q0" + PeriodInfo.Quy + (ConvertHelpers.ConvertToInt(PeriodInfo.Nam) - 1) + "_" + text2 + "_" + text3 + "_" + ToKhaiInfo.VersionXML.Replace(".", "").Replace("_", "") + ".xml";
				int num = ConvertHelpers.ConvertToInt(PeriodInfo.Quy);
				int num2 = ConvertHelpers.ConvertToInt(PeriodInfo.Nam);
				int month = GeneralFunctions.GetNgayDauQuy(num, num2, 1, 1).Month;
				Dictionary<string, string> quyNamByThangNam = GeneralFunctions.GetQuyNamByThangNam(month, num2, HeaderInfo.NgayBatDauNamTaiChinh);
				if (num == 1)
				{
					num = 4;
					num2--;
				}
				else
				{
					num--;
				}
				string value2 = "";
				quyNamByThangNam.TryGetValue(num.ToString(), out value2);
				if (num2 == ConvertHelpers.ConvertToInt(value2))
				{
					string value3 = ToKhaiInfo.SheetInfos[sheetIndex].DataFile + "_Q0" + num + value2;
					string value4 = ToKhaiInfo.VersionXML.Replace(".", "").Replace("_", "") + ".xml";
					if (directoryInfo.Exists)
					{
						FileInfo[] files = directoryInfo.GetFiles();
						if (files.Length > 0)
						{
							FileInfo[] array = files;
							foreach (FileInfo fileInfo in array)
							{
								if (fileInfo.Name.Contains(value3) && fileInfo.Name.Contains(value4))
								{
									item = text4 + "\\" + fileInfo.Name;
									break;
								}
							}
						}
					}
				}
			}
			list.Add(text);
			list.Add(item);
			list.Add(item2);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return list;
	}

	private void checkSoNamTruoc(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, int sheetIndex, bool isCheck)
	{
		try
		{
			int num = 0;
			SheetView sheetView = winFpSpread.Sheets[sheetIndex];
			List<string> pathDataFileOld = getPathDataFileOld(ToKhaiInfo, TTinChung, PeriodInfo, HeaderInfo, sheetIndex);
			int num2 = 0;
			int num3 = 0;
			int num4 = 0;
			if (sheetIndex == 0 || sheetIndex == 2)
			{
				num2 = HTKKFile.ColumnLettersToIndex("R");
				num3 = HTKKFile.ColumnLettersToIndex("W");
				num4 = HTKKFile.ColumnLettersToIndex("AC");
				if (pathDataFileOld == null || (pathDataFileOld != null && !File.Exists(pathDataFileOld[0])))
				{
					return;
				}
			}
			if (sheetIndex == 1)
			{
				num2 = HTKKFile.ColumnLettersToIndex("I");
				num3 = HTKKFile.ColumnLettersToIndex("J");
				num4 = HTKKFile.ColumnLettersToIndex("N");
			}
			string text = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[25, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			string text2 = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Cells[52, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[sheetIndex];
			for (int i = 1; i < hTKKCMCSheetInfo.StaticCells.Count; i++)
			{
				if (!File.Exists(pathDataFileOld[0]) || hTKKCMCSheetInfo.StaticCells[i].CellTemplate.ColIndex != num3 || sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num3 - 1].Locked)
				{
					continue;
				}
				string valueCustom = GeneralFunctions.GetValueCustom(sheetView, hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex, num3, "NumberCellType");
				object value = sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num4 - 1].Value;
				if (ConvertHelpers.ConvertToDouble(valueCustom) != ConvertHelpers.ConvertToDouble(value) && isCheck)
				{
					if (sheetView.Cells[hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num3 - 1].BackColor != Color.Pink)
					{
						num = 1;
						HTKKFile.SetNodeBackGroundYellow(sheetView, hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num3 - 1, text);
					}
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex - 1, num3 - 1);
				}
			}
			if ((num != 2 || !(winFpSpread.TabStrip[sheetIndex].BackColor == Color.Pink)) && (num != 1 || !(winFpSpread.TabStrip[sheetIndex].BackColor == Color.Yellow)))
			{
				if (num == 2 && winFpSpread.TabStrip[sheetIndex].BackColor != Color.Pink)
				{
					winFpSpread.TabStrip[sheetIndex].BackColor = Color.Pink;
				}
				else if (num == 1 && winFpSpread.TabStrip[sheetIndex].BackColor == Color.White)
				{
					winFpSpread.TabStrip[sheetIndex].BackColor = Color.Yellow;
				}
				else if (num == 0)
				{
					winFpSpread.TabStrip[sheetIndex].BackColor = Color.White;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void FinishImportXML(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			sheetView.Cells[GeneralFunctions.ArrayCellId("S_30")[1], GeneralFunctions.ArrayCellId("S_30")[0]].Value = TTinChung.kyKKhaiTuNgay;
			sheetView.Cells[GeneralFunctions.ArrayCellId("U_30")[1], GeneralFunctions.ArrayCellId("U_30")[0]].Value = TTinChung.kyKKhaiDenNgay;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}
}
