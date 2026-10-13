using System;
using System.Collections.Generic;
using System.Drawing;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml;
using FarPoint.Win.Spread;
using FarPoint.Win.Spread.CellType;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace TAX.HTKK.Business;

public class AcBk01Nckt201
{
	private Dictionary<string, string> dicDMHoaDonTmp = new Dictionary<string, string>();

	private Dictionary<string, string[]> dicDMHoaDon = new Dictionary<string, string[]>();

	private List<string> listHD = new List<string>();

	private string pathCellNguoiKy = "HSoThueDTu/HSoKhaiThue/TTinChung/TTinTKhaiThue/TKhaiThue/nguoiKy";

	private string pathCellNgayKy = "HSoThueDTu/HSoKhaiThue/TTinChung/TTinTKhaiThue/TKhaiThue/ngayKy";

	private int idxCol_C => HTKKFile.ColumnLettersToIndex("C") - 1;

	private int idxCol_D => HTKKFile.ColumnLettersToIndex("D") - 1;

	private int idxCol_O => HTKKFile.ColumnLettersToIndex("O") - 1;

	private int idxCol_S => HTKKFile.ColumnLettersToIndex("S") - 1;

	private int idxCol_U => HTKKFile.ColumnLettersToIndex("U") - 1;

	private int idxCol_V => HTKKFile.ColumnLettersToIndex("V") - 1;

	private int idxCol_Y => HTKKFile.ColumnLettersToIndex("Y") - 1;

	private int idxCol_AA => HTKKFile.ColumnLettersToIndex("AA") - 1;

	private int idxCol_AI => HTKKFile.ColumnLettersToIndex("AI") - 1;

	public void AutoBindData_0(SheetView sView, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathCellNgayKy));
			if (hTKKCMCCellInfo != null)
			{
				sView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Text = PeriodInfo.kyKKhai;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void KeyPress(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, int idxSheet, int idxRow, int idxCol, KeyPressEventArgs e)
	{
		try
		{
			if (idxSheet == 0)
			{
				SheetView sheetView = winFpSpread.Sheets[0];
				HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
				HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
				int num = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
				int num2 = num + hTKKCMCTableInfo.RowsValues.Count - 1;
				if (idxRow >= num && idxRow <= num2 && (idxCol == idxCol_V || idxCol == idxCol_Y))
				{
					GeneralFunctions.CompareKeycode(e, 1);
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void CellEditchange_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			int num = dataLocationTemplateRowStart + hTKKCMCTableInfo.RowsValues.Count - 1;
			if (iRow < dataLocationTemplateRowStart || iRow > num || iCol - 1 != idxCol_C)
			{
				return;
			}
			string text = sheetView.Cells[(int)iRow - 1, idxCol_C].Text;
			if (HTKKFormater.isNullString(text))
			{
				sheetView.Cells[(int)iRow - 1, idxCol_D].ResetValue();
				sheetView.Cells[(int)iRow - 1, idxCol_O].ResetValue();
				sheetView.Cells[(int)iRow - 1, idxCol_AI].ResetValue();
				return;
			}
			string[] array = dicDMHoaDon[text];
			sheetView.Cells[(int)iRow - 1, idxCol_D].Text = array[2];
			sheetView.Cells[(int)iRow - 1, idxCol_AI].Text = array[0];
			if (!HTKKFormater.isNullString(array[2]))
			{
				sheetView.Cells[(int)iRow - 1, idxCol_O].Text = text;
			}
			else
			{
				sheetView.Cells[(int)iRow - 1, idxCol_O].ResetValue();
			}
			string text2 = sheetView.Cells[(int)iRow - 1, idxCol_AI].Text;
			if (text2.Equals("0") || text2.Equals("4"))
			{
				TextCellType textCellType = new TextCellType();
				textCellType.WordWrap = true;
				TextCellType textCellType2 = textCellType;
				textCellType2.MaxLength = 13;
				sheetView.Cells[(int)iRow - 1, idxCol_O].CellType = textCellType2;
			}
			else
			{
				TextCellType textCellType3 = new TextCellType();
				textCellType3.WordWrap = true;
				TextCellType textCellType2 = textCellType3;
				textCellType2.MaxLength = 20;
				sheetView.Cells[(int)iRow - 1, idxCol_O].CellType = textCellType2;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public void LeaveCell_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, long iRow, long iCol, long iNewRow, long iNewCol)
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

	public void CellChange_0(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, long iRow, long iCol)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int num = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
			int num2 = num + hTKKCMCTableInfo.RowsValues.Count - 1;
			if (iRow - 1 >= num && iRow - 1 <= num2 && (iCol - 1 == idxCol_V || iCol - 1 == idxCol_Y))
			{
				int num3 = ConvertHelpers.ConvertToInt(sheetView.Cells[(int)iRow - 1, idxCol_V].Text);
				int num4 = ConvertHelpers.ConvertToInt(sheetView.Cells[(int)iRow - 1, idxCol_Y].Text);
				if ((num3 == 0 && num4 == 0) || num3 > num4)
				{
					sheetView.Cells[(int)iRow - 1, idxCol_AA].Value = 0;
				}
				else
				{
					sheetView.Cells[(int)iRow - 1, idxCol_AA].Value = num4 - num3 + 1;
				}
			}
			string text = "";
			if (iRow - 1 < num || iRow - 1 > num2)
			{
				return;
			}
			if (iCol - 1 == idxCol_V || iCol - 1 == idxCol_Y)
			{
				text = sheetView.Cells[(int)iRow - 1, (int)iCol - 1].Text;
				sheetView.Cells[(int)iRow - 1, (int)iCol - 1].Text = GeneralFunctions.FormatSoHD(text);
			}
			if (iCol - 1 == idxCol_S)
			{
				text = sheetView.Cells[(int)iRow - 1, (int)iCol - 1].Text;
				sheetView.Cells[(int)iRow - 1, (int)iCol - 1].Text = GeneralFunctions.FormatKyHieu(text);
			}
			if (iCol - 1 == idxCol_O)
			{
				string text2 = sheetView.Cells[(int)iRow - 1, HTKKFile.ColumnLettersToIndex("C") - 1].Text;
				string text3 = "";
				if (!string.IsNullOrEmpty(text2))
				{
					text3 = dicDMHoaDon[text2][0];
				}
				if (!text3.Equals("3"))
				{
					string text4 = sheetView.Cells[(int)iRow - 1, idxCol_O].Text;
					sheetView.Cells[(int)iRow - 1, idxCol_O].Text = text4.ToUpper();
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public bool isValidDuLieuTK(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo, Dictionary<int, int> levelWarning)
	{
		try
		{
			Validate_0(winFpSpread, ToKhaiInfo, TTinChung, PeriodInfo, levelWarning);
			return true;
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
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int num = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
			HTKKCMCRowInfo rowsInfo = hTKKCMCTableInfo.RowsInfo;
			string message = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_32"));
			string message2 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_33"));
			string message3 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_34"));
			string message4 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_35"));
			string message5 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_36"));
			string message6 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_37"));
			string message7 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_38"));
			string message8 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_39"));
			string message9 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_40"));
			string message10 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_41"));
			string message11 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_48"));
			string message12 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_51"));
			string message13 = HTKKFile.GetMessage(sheetView, GeneralFunctions.ArrayCellId("E_55"));
			string[] array = GeneralFunctions.checkTrung(winFpSpread, num, "V", "Y", "B", "D", "O", "S", "", isMST: false, 1);
			string[] array2 = new string[rowsInfo.Cells.Count];
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				array2[0] = sheetView.Cells[num + i, idxCol_C].Text.Trim();
				array2[1] = sheetView.Cells[num + i, idxCol_D].Text.Trim();
				array2[2] = sheetView.Cells[num + i, idxCol_O].Text.Trim();
				array2[3] = sheetView.Cells[num + i, idxCol_S].Text.Trim();
				array2[4] = sheetView.Cells[num + i, idxCol_V].Text.Trim();
				array2[5] = sheetView.Cells[num + i, idxCol_Y].Text.Trim();
				if (!GeneralFunctions.HaveToValidateRow(rowsInfo, array2))
				{
					HTKKFile.RemoveValidate(sheetView, num + i, idxCol_C);
					HTKKFile.RemoveValidate(sheetView, num + i, idxCol_D);
					HTKKFile.RemoveValidate(sheetView, num + i, idxCol_O);
					HTKKFile.RemoveValidate(sheetView, num + i, idxCol_S);
					HTKKFile.RemoveValidate(sheetView, num + i, idxCol_V);
					HTKKFile.RemoveValidate(sheetView, num + i, idxCol_Y);
					continue;
				}
				HTKKFile.RemoveValidate(sheetView, num + i, idxCol_C);
				HTKKFile.RemoveValidate(sheetView, num + i, idxCol_D);
				HTKKFile.RemoveValidate(sheetView, num + i, idxCol_O);
				HTKKFile.RemoveValidate(sheetView, num + i, idxCol_S);
				HTKKFile.RemoveValidate(sheetView, num + i, idxCol_V);
				HTKKFile.RemoveValidate(sheetView, num + i, idxCol_Y);
				string text = sheetView.Cells[num + i, idxCol_AI].Text;
				GeneralFunctions.strMauSoHD_01GTKT = array2[0];
				if (HTKKFormater.isNullString(array2[0]))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, idxCol_C, message7);
					value = 2;
				}
				if (HTKKFormater.isNullString(array2[1]))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, idxCol_D, message7);
					value = 2;
				}
				if (HTKKFormater.isNullString(array2[2]))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, idxCol_O, message7);
					value = 2;
				}
				else
				{
					string text2 = array2[0];
					if (text2 != "TT120")
					{
						string text3 = GeneralFunctions.CheckMauSoHD(array2[2], text, text2);
						if (text3.Equals("1"))
						{
							HTKKFile.SetNodeBackGroundPink(sheetView, num + i, idxCol_O, message);
							value = 2;
						}
						else if (text3.Equals("2"))
						{
							if (array2[3].Length > 0)
							{
								if (array2[3].Substring(array2[3].Length - 1).Equals("E"))
								{
									if (GeneralFunctions.CheckSoLienHDDT(array2[2], text, text2, array2[3]).Equals("3"))
									{
										HTKKFile.SetNodeBackGroundYellow(sheetView, num + i, idxCol_O, message13);
										value = 1;
									}
									else
									{
										HTKKFile.RemoveValidate(sheetView, num + i, idxCol_O);
									}
								}
								else
								{
									HTKKFile.SetNodeBackGroundPink(sheetView, num + i, idxCol_O, message3);
									value = 2;
								}
							}
						}
						else if (array2[3].Length > 0)
						{
							if (array2[3].Substring(array2[3].Length - 1).Equals("E"))
							{
								if (GeneralFunctions.CheckSoLienHDDT(array2[2], text, text2, array2[3]).Equals("3"))
								{
									HTKKFile.SetNodeBackGroundYellow(sheetView, num + i, idxCol_O, message13);
									value = 1;
								}
								else
								{
									HTKKFile.RemoveValidate(sheetView, num + i, idxCol_O);
								}
							}
							else
							{
								HTKKFile.RemoveValidate(sheetView, num + i, idxCol_O);
							}
						}
					}
				}
				if (HTKKFormater.isNullString(array2[3]))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, idxCol_S, message7);
					value = 2;
				}
				else if (!string.IsNullOrEmpty(array2[0]) && array2[0] != "TT120")
				{
					string text4 = GeneralFunctions.CheckSoHD(array2[3], text);
					if (text4.Equals("1"))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, idxCol_S, message2);
						value = 2;
					}
					else if (text4.Equals("2"))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, idxCol_S, message5);
						value = 2;
					}
					else
					{
						HTKKFile.RemoveValidate(sheetView, num + i, idxCol_S);
					}
				}
				bool flag = false;
				if (HTKKFormater.isNullString(array2[4]) && HTKKFormater.isNullString(array2[5]))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, idxCol_V, message7);
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, idxCol_Y, message7);
					value = 2;
					flag = true;
				}
				else if (HTKKFormater.isNullString(array2[4]) && !HTKKFormater.isNullString(array2[5]))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, idxCol_V, message7);
					HTKKFile.RemoveValidate(sheetView, num + i, idxCol_Y);
					value = 2;
					flag = true;
				}
				else if (!HTKKFormater.isNullString(array2[4]) && HTKKFormater.isNullString(array2[5]))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, num + i, idxCol_Y, message7);
					HTKKFile.RemoveValidate(sheetView, num + i, idxCol_V);
					value = 2;
					flag = true;
				}
				else
				{
					double num2 = ConvertHelpers.ConvertToDouble(sheetView.Cells[num + i, idxCol_V].Text);
					double num3 = ConvertHelpers.ConvertToDouble(sheetView.Cells[num + i, idxCol_Y].Text);
					if ((num3 - num2 >= 0.0 && num3 > 0.0) || (num2 == 0.0 && num3 == 0.0))
					{
						HTKKFile.RemoveValidate(sheetView, num + i, idxCol_V);
						HTKKFile.RemoveValidate(sheetView, num + i, idxCol_Y);
					}
					else
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, idxCol_V, message6);
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, idxCol_Y, message6);
						value = 2;
						flag = true;
					}
				}
				if (array.Length <= 1)
				{
					continue;
				}
				for (int j = 0; j < array.Length; j++)
				{
					if (num + i == Convert.ToInt32(array[j]))
					{
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, idxCol_V, message10);
						HTKKFile.SetNodeBackGroundPink(sheetView, num + i, idxCol_Y, message10);
						value = 2;
						break;
					}
				}
			}
			HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathCellNguoiKy));
			if (hTKKCMCCellInfo != null)
			{
				string str = sheetView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Text.Trim();
				if (HTKKFormater.isNullString(str))
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1, message7);
					value = 2;
				}
				else
				{
					HTKKFile.RemoveValidate(sheetView, hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1);
				}
			}
			switch (value)
			{
			case 2:
				winFpSpread.TabStrip[0].BackColor = Color.Pink;
				break;
			case 1:
				winFpSpread.TabStrip[0].BackColor = Color.Yellow;
				break;
			default:
				winFpSpread.TabStrip[0].BackColor = Color.White;
				break;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public Dictionary<string, string> GetDMHoaDon(string pathFile)
	{
		//IL_0032: Unknown result type (might be due to invalid IL or missing references)
		//IL_0038: Expected O, but got Unknown
		//IL_0069: Unknown result type (might be due to invalid IL or missing references)
		//IL_0070: Expected O, but got Unknown
		try
		{
			if (dicDMHoaDonTmp.Count > 0)
			{
				return dicDMHoaDonTmp;
			}
			Dictionary<string, string[]> dictionary = new Dictionary<string, string[]>();
			Dictionary<string, string> dictionary2 = new Dictionary<string, string>();
			XmlDocument val = new XmlDocument();
			val.Load(pathFile);
			XmlNodeList elementsByTagName = val.GetElementsByTagName("Item");
			listHD = new List<string>();
			foreach (XmlNode item in elementsByTagName)
			{
				XmlNode val2 = item;
				string value = ((XmlNode)val2.Attributes["Value"]).Value;
				string[] array = Regex.Split(value, "###");
				string key = array[1];
				string value2 = array[1];
				dictionary2.Add(key, value2);
				dictionary.Add(key, array);
				listHD.Add(value);
			}
			dicDMHoaDon = dictionary;
			dicDMHoaDonTmp = dictionary2;
			return dictionary2;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public void ExportHeaderXML_0(HTKKCMCToKhaiInfo ToKhaiInfo, HTKKCMCTTinChungInfo TTinChung, HTKKPeriodInfo PeriodInfo, XmlNode xmlNode)
	{
		try
		{
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void FinishImportXML(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		try
		{
			SheetView sheetView = winFpSpread.Sheets[0];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[0];
			HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[0];
			int num = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
			string text = "";
			string text2 = "";
			int num2 = HTKKFile.ColumnLettersToIndex("C");
			int num3 = HTKKFile.ColumnLettersToIndex("O");
			bool flag = true;
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				text2 = sheetView.Cells[num + i, num3 - 1].Text;
				flag = true;
				for (int j = 0; j < listHD.Count; j++)
				{
					string input = listHD[j];
					string[] array = Regex.Split(input, "###");
					if (text2.IndexOf(array[1], 0) == 0)
					{
						sheetView.Cells[num + i, idxCol_AI].Text = array[0];
						flag = false;
						break;
					}
				}
				if (flag && !string.IsNullOrEmpty(text2))
				{
					sheetView.Cells[num + i, idxCol_AI].Text = "3";
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}
}
