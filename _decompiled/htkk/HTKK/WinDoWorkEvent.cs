using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Security;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;
using System.Xml.Schema;
using CA;
using FarPoint.Win.Spread;
using FarPoint.Win.Spread.CellType;
using HTKKDanhmuc;
using Ionic.Zip;
using TAX.HTKK.Business;
using TAX.HTKK.ComputeHash;
using TAX.HTKK.Utilities.Control;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace HTKK;

public class WinDoWorkEvent
{
	private static string soLanTT = "";

	public static int firstSheetErr = -1;

	public static HashSet<string> nillableElements = new HashSet<string>();

	private static string strSession = HTKKFile.GetSession(6);

	public static bool WinDoWorkSave(FpSpread winFpSpread)
	{
		try
		{
			for (int i = 0; i < WinSettings.ToKhaiInfo.SheetInfos.Count; i++)
			{
				if (WinSettings.ToKhaiInfo.SheetInfos[i] != null && WinSettings.DicSheet.ContainsKey(WinSettings.ToKhaiInfo.SheetInfos[i].SheetName))
				{
					WinDoWorkSheetSave(winFpSpread.Sheets[i], WinSettings.ToKhaiInfo.SheetInfos[i]);
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return false;
		}
		return true;
	}

	public static string WinDoWorkExportXML(FpSpread winFpSpread, string fileName)
	{
		//IL_08d8: Unknown result type (might be due to invalid IL or missing references)
		//IL_08df: Expected O, but got Unknown
		string result = "";
		try
		{
			string empty = string.Empty;
			WinSettings.ExportingData = "";
			if (WinSettings.MenuID == "G4" || WinSettings.MenuID == "N5" || WinSettings.MenuID == "H5" || WinSettings.MenuID == "O7" || WinSettings.MenuID == "P5" || WinSettings.MenuID == "U6" || "U8~".Contains(WinSettings.MenuID))
			{
				soLanTT = winFpSpread.Sheets[0].Cells[1, HTKKFile.ColumnLettersToIndex("T") - 1].Text;
			}
			XmlDocument val = WinDoWorkGetExportSheetXML(winFpSpread.Sheets[0], WinSettings.ToKhaiInfo.SheetInfos[0]);
			if (WinSettings.MenuID == "M4")
			{
				XmlNode val2 = ((XmlNode)val).SelectSingleNode("TKhai");
				XmlAttribute val3 = val.CreateAttribute("xmlns:xsi");
				((XmlNode)val3).Value = "http://www.w3.org/2001/XMLSchema-instance";
				((XmlNamedNodeMap)val2.Attributes).SetNamedItem((XmlNode)(object)val3);
				val3 = val.CreateAttribute("xmlns");
				((XmlNode)val3).Value = "http://kekhaithue.gdt.gov.vn/TKhaiThue";
				((XmlNamedNodeMap)val2.Attributes).SetNamedItem((XmlNode)(object)val3);
				val3 = val.CreateAttribute("xmlns:n2");
				((XmlNode)val3).Value = "http://www.altova.com/samplexml/other-namespace";
				((XmlNamedNodeMap)val2.Attributes).SetNamedItem((XmlNode)(object)val3);
				val3 = val.CreateAttribute("xmlns:ds");
				((XmlNode)val3).Value = "http://www.w3.org/2000/09/xmldsig#";
				((XmlNamedNodeMap)val2.Attributes).SetNamedItem((XmlNode)(object)val3);
			}
			else
			{
				XmlNode val2 = ((XmlNode)val).SelectSingleNode("HSoThueDTu");
				XmlAttribute val3 = val.CreateAttribute("xmlns:xsi");
				((XmlNode)val3).Value = "http://www.w3.org/2001/XMLSchema-instance";
				((XmlNamedNodeMap)val2.Attributes).SetNamedItem((XmlNode)(object)val3);
				val3 = val.CreateAttribute("xmlns");
				((XmlNode)val3).Value = "http://kekhaithue.gdt.gov.vn/TKhaiThue";
				((XmlNamedNodeMap)val2.Attributes).SetNamedItem((XmlNode)(object)val3);
			}
			if (HTKKFormater.isNullString(WinSettings.HTKKHeaderInfo.mstDLyThue) && WinSettings.MenuID != "M4")
			{
				XmlNode parentNode = ((XmlNode)val).SelectSingleNode("HSoThueDTu/HSoKhaiThue/TTinChung/TTinTKhaiThue/DLyThue").ParentNode;
				XmlNode val4 = parentNode.SelectSingleNode("DLyThue");
				if (val4 != null)
				{
					XmlNode val5 = val4.SelectSingleNode("mstDLyThue");
					XmlNode val6 = val5.OwnerDocument.CreateNode((XmlNodeType)2, "xsi:nil", "http://www.w3.org/2001/XMLSchema-instance");
					val6.Value = "true";
					if (val5.InnerText == "")
					{
						((XmlNamedNodeMap)val5.Attributes).SetNamedItem(val6);
					}
					val5 = val4.SelectSingleNode("ngayKyHDDLyThue");
					val6 = val5.OwnerDocument.CreateNode((XmlNodeType)2, "xsi:nil", "http://www.w3.org/2001/XMLSchema-instance");
					val6.Value = "true";
					if (val5.InnerText == "")
					{
						((XmlNamedNodeMap)val5.Attributes).SetNamedItem(val6);
					}
				}
				if (val4 != null && WinSettings.MenuID != "M6" && WinSettings.MenuID != "O8" && WinSettings.MenuID != "B2" && WinSettings.MenuID != "O9")
				{
					parentNode.RemoveChild(val4);
				}
			}
			XmlNode documentElement = (XmlNode)(object)val.DocumentElement;
			if (WinSettings.MenuID == "M4")
			{
				xmlExportTTinChungHDDT(val);
			}
			else
			{
				xmlExportTTinChung(val);
			}
			ExportHeaderXML((XmlNode)(object)val, 0);
			ExportXML(documentElement, 0);
			XmlNode val7 = null;
			val7 = ((!(WinSettings.MenuID == "M4")) ? ((XmlNode)val).SelectSingleNode("HSoThueDTu/HSoKhaiThue/PLuc") : ((XmlNode)val).SelectSingleNode("TKhai/DLTKhai/PLuc"));
			for (int i = 1; i < WinSettings.ToKhaiInfo.SheetInfos.Count; i++)
			{
				if (WinSettings.ToKhaiInfo.SheetInfos[i] != null && ((WinSettings.MenuID == "99" && i == 2) || WinSettings.MenuID == "Q6" || WinSettings.MenuID == "Q5"))
				{
					WinSettings.ToKhaiInfo.SheetInfos[i].CoDuLieu = true;
				}
				if (WinSettings.ToKhaiInfo.SheetInfos[i] != null && (WinSettings.ToKhaiInfo.SheetInfos[i].CoDuLieu || ((WinSettings.xmlMenuNode.Attributes["ParentID"] == null) ? "" : ((XmlNode)WinSettings.xmlMenuNode.Attributes["ParentID"]).Value) == "101_6" || WinSettings.TTinChung.loaiTKhai == "B") && WinSettings.DicSheet.ContainsKey(WinSettings.ToKhaiInfo.SheetInfos[i].SheetName) && WinSettings.ToKhaiInfo.SheetInfos[i].SheetName != "Header")
				{
					XmlDocument val8 = WinDoWorkGetExportSheetXML(winFpSpread.Sheets[i], WinSettings.ToKhaiInfo.SheetInfos[i]);
					if (WinSettings.ToKhaiInfo.KieuKHBS == "3" && WinSettings.TTPeriodInfo.LoaiToKhai == "B" && (i == WinSettings.ToKhaiInfo.SheetInfos.Count - 1 || i == WinSettings.ToKhaiInfo.SheetInfos.Count - 2))
					{
						xmlExportTTinChungKHBS(val8);
					}
					HTKKFile.LogRun("ExportXML", "WinDoWorkGetExportSheetXML");
					XmlElement documentElement2 = val8.DocumentElement;
					ExportXML((XmlNode)(object)documentElement2, i);
					if (WinSettings.MenuID.Equals("N8") && i == 8)
					{
						ExportTableF7XML(winFpSpread, (XmlNode)(object)documentElement2, 8);
					}
					if (WinSettings.TTPeriodInfo.LoaiToKhai == "B" && "30,I5,K9,K1,M7,K6,K8,73,55,71,31,P2,X6".Contains(WinSettings.MenuID) && WinSettings.ToKhaiInfo.SheetInfos[i].SheetName == "KHBS")
					{
						val7.InsertBefore(val7.OwnerDocument.ImportNode((XmlNode)(object)documentElement2, true), val7.FirstChild);
						continue;
					}
					XmlNode val9 = val.ImportNode((XmlNode)(object)documentElement2, true);
					HTKKFile.LogRun("ExportXML", "ImportNode " + i);
					val7.AppendChild(val9);
					HTKKFile.LogRun("ExportXML", "AppendChild " + i);
				}
			}
			if (val7 != null && !val7.HasChildNodes)
			{
				val7.ParentNode.RemoveChild(val7);
			}
			string text = "tmpExport.xml";
			string text2 = HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst;
			string text3 = text2 + "\\" + text;
			string text4 = text2 + "\\log_valid.txt";
			if (File.Exists(text4))
			{
				HTKKFile.DeleteFile(text4);
			}
			val.Save(text3);
			HTKKFile.LogRun("ExportXML", "Save");
			if (string.IsNullOrEmpty(WinSettings.ToKhaiInfo.XsdFileName) || (!string.IsNullOrEmpty(WinSettings.ToKhaiInfo.XsdFileName) & ValidateXmlAgainstXsd(text3, WinSettings.ToKhaiInfo.XsdFileName)))
			{
				HTKKFile.LogRun("ExportXML", "ValidateXmlAgainstXsd");
				try
				{
					if (!(WinSettings.MenuID == "33") || WinSettings.ToKhaiInfo.SheetInfos.Count <= 1 || WinSettings.ToKhaiInfo.SheetInfos[1] == null || WinSettings.ToKhaiInfo.SheetInfos[1].Tables[0].RowsValues.Count <= 200000)
					{
						XmlNode val10 = ((XmlNode)val).SelectSingleNode("/HSoThueDTu/HSoKhaiThue");
						if (val10 != null)
						{
							XmlDocument val11 = new XmlDocument();
							val11.LoadXml(val10.OuterXml);
							XmlNode val12 = ((XmlNode)val11).SelectSingleNode("/HSoKhaiThue/TTinChung/TTinDVu");
							if (val12 != null)
							{
								val12.ParentNode.RemoveChild(val12);
							}
							string text5 = HashGenerator.GenerateKey(((XmlNode)val11).InnerXml, "HTKK.CMCSoft");
							if (!string.IsNullOrEmpty(text5))
							{
								((XmlNode)val).SelectSingleNode("/HSoThueDTu/HSoKhaiThue/TTinChung/TTinDVu/ttinNhaCCapDVu").InnerText = text5;
							}
						}
						HTKKFile.LogRun("ExportXML", "Process EnCrypt");
					}
				}
				catch (Exception ex)
				{
					HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
				}
				val.Save(fileName);
				if (WinSettings.TTPeriodInfo.LoaiToKhai == "B" && "Q5,Q6".Contains(WinSettings.MenuID))
				{
					GeneralFunctions.ExportXmlKHBSSuaDonViTien(fileName, winFpSpread, WinSettings.ToKhaiInfo, 1, "path_ten_don_vi_tien_khbs", "path_ma_don_vi_tien_khbs");
				}
				HTKKFile.LogRun("ExportXML", "Save Encrypt");
			}
			else if (WinSettings.MenuID == "33" && WinSettings.ToKhaiInfo.SheetInfos.Count > 1 && WinSettings.ToKhaiInfo.SheetInfos[1] != null && WinSettings.ToKhaiInfo.SheetInfos[1].Tables[0].RowsValues.Count > 200000)
			{
				string text6 = File.ReadAllText(text3);
				string value = text6.Substring(0, text6.IndexOf("<PLuc_05_1_BK_QTT>") + "<PLuc_05_1_BK_QTT>".Length) + Environment.NewLine;
				string value2 = text6.Substring(text6.LastIndexOf("<PLuc_05_1_BK_QTT>") + "<PLuc_05_1_BK_QTT>".Length);
				StringBuilder stringBuilder = new StringBuilder();
				stringBuilder.Append(value);
				stringBuilder.Append(WinSettings.ExportingData);
				stringBuilder.Append(value2);
				File.WriteAllText(text3, stringBuilder.ToString(), Encoding.UTF8);
				if (string.IsNullOrEmpty(WinSettings.ToKhaiInfo.XsdFileName) || (!string.IsNullOrEmpty(WinSettings.ToKhaiInfo.XsdFileName) & ValidateXmlAgainstXsd(text3, WinSettings.ToKhaiInfo.XsdFileName)))
				{
					File.WriteAllText(fileName, stringBuilder.ToString(), Encoding.UTF8);
				}
				else
				{
					result = "0279";
				}
			}
			else
			{
				result = "0279";
			}
		}
		catch (Exception ex2)
		{
			result = "0281";
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex2);
		}
		return result;
	}

	public static IEnumerable<XElement> StreamItem(string uri)
	{
		XmlReader reader = XmlReader.Create(uri);
		try
		{
			reader.MoveToContent();
			while (reader.Read())
			{
				if ((int)reader.NodeType == 1)
				{
					XNode obj = XNode.ReadFrom(reader);
					yield return (XElement)(object)((obj is XElement) ? obj : null);
				}
			}
		}
		finally
		{
			((IDisposable)reader)?.Dispose();
		}
	}

	public static bool ValidateXmlAgainstXsd(string xmlPath, string xsdPath)
	{
		//IL_0096: Unknown result type (might be due to invalid IL or missing references)
		//IL_009c: Expected O, but got Unknown
		//IL_00b8: Unknown result type (might be due to invalid IL or missing references)
		//IL_00be: Unknown result type (might be due to invalid IL or missing references)
		//IL_00c7: Unknown result type (might be due to invalid IL or missing references)
		//IL_00cd: Unknown result type (might be due to invalid IL or missing references)
		//IL_00dc: Unknown result type (might be due to invalid IL or missing references)
		//IL_00e6: Expected O, but got Unknown
		string text = "log_valid.txt";
		try
		{
			string text2 = HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst;
			if (File.Exists(HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst + "\\log_valid.txt"))
			{
				File.Delete(HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst + "\\log_valid.txt");
			}
			if (!File.Exists(xsdPath))
			{
				HTKKFile.saveFile(text2, text, "Chưa có file XSD");
				return false;
			}
			XmlReaderSettings val = new XmlReaderSettings();
			val.Schemas.Add("http://kekhaithue.gdt.gov.vn/TKhaiThue", xsdPath);
			val.ValidationType = (ValidationType)4;
			val.ValidationFlags = (XmlSchemaValidationFlags)(val.ValidationFlags | 1);
			val.ValidationFlags = (XmlSchemaValidationFlags)(val.ValidationFlags | 4);
			val.ValidationEventHandler += new ValidationEventHandler(ValidationCallBack);
			XmlReader val2 = XmlReader.Create(xmlPath, val);
			while (val2.Read())
			{
			}
			val2.Close();
			text2 = text2 + "\\" + text;
			if (File.Exists(text2))
			{
				return false;
			}
			return true;
		}
		catch (Exception ex)
		{
			string text2 = HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst;
			HTKKFile.saveFile(text2, text, ex.Message);
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return false;
		}
	}

	private static void ValidationCallBack(object sender, ValidationEventArgs args)
	{
		//IL_0022: Unknown result type (might be due to invalid IL or missing references)
		//IL_0028: Invalid comparison between Unknown and I4
		string fileName = "log_valid.txt";
		string pathFolder = HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst;
		if ((int)args.Severity == 1)
		{
			HTKKFile.saveFile(pathFolder, fileName, args.Message);
		}
		else
		{
			HTKKFile.saveFile(pathFolder, fileName, args.Message);
		}
	}

	public static void xmlExportTTinChung(XmlDocument xmlExportContent)
	{
		try
		{
			XmlNode val = ((XmlNode)xmlExportContent).SelectSingleNode("HSoThueDTu/HSoKhaiThue/TTinChung");
			xmlExportTTinChungProcessNode(val);
			if ("L4~L7~L3~L8~B9~J9~K5~M9~M8~D9~E5~K4~E3~E1~E2~K2~K3~E9~E6~E4~P7~R9~R6~U7~X9~Y1~Y2~Y3~Y4~Y5~Y6~Y7~Y8~Y9~Z8~W1~".IndexOf(WinSettings.MenuID + "~") >= 0)
			{
				XmlNode val2 = val.SelectSingleNode("TTinTKhaiThue/NNT/tenXaNNT");
				val2.InnerText = WinSettings.HTKKHeaderInfo.tenXaPhuongNNT;
				XmlNode val3 = val.SelectSingleNode("TTinTKhaiThue/NNT/maXaNNT");
				val3.InnerText = WinSettings.HTKKHeaderInfo.maXaPhuongNNT;
			}
			XmlNode val4 = val.SelectSingleNode("TTinTKhaiThue/TKhaiThue/ngayKy");
			if (val4 != null)
			{
				string text = val4.InnerText;
				if (WinSettings.MenuID == "07")
				{
					text = WinSettings.TTPeriodInfo.kyKKhai;
				}
				if (text.Length >= 8)
				{
					string[] array = text.Split(new char[1] { '/' });
					if (array.Length == 3)
					{
						if (array[0].Length == 1)
						{
							array[0] = "0" + array[0];
						}
						if (array[1].Length == 1)
						{
							array[1] = "0" + array[1];
						}
						if (array[2].Length == 2)
						{
							array[2] = "20" + array[2];
						}
					}
					val4.InnerText = array[2] + "-" + array[1] + "-" + array[0];
				}
				if (HTKKFormater.isNullString(val4.InnerText))
				{
					val4.InnerText = DateTime.Now.ToString("yyyy-MM-dd");
					val4.Attributes.Remove(val4.Attributes["xsi:nil"]);
				}
			}
			XmlNode val5 = val.SelectSingleNode("TTinTKhaiThue/ngayKy");
			if (val5 != null)
			{
				string text = val5.InnerText;
				if (WinSettings.MenuID == "07")
				{
					text = WinSettings.TTPeriodInfo.kyKKhai;
				}
				if (text.Length >= 8)
				{
					string[] array = text.Split(new char[1] { '/' });
					if (array.Length == 3)
					{
						if (array[0].Length == 1)
						{
							array[0] = "0" + array[0];
						}
						if (array[1].Length == 1)
						{
							array[1] = "0" + array[1];
						}
						if (array[2].Length == 2)
						{
							array[2] = "20" + array[2];
						}
					}
					val5.InnerText = array[2] + "-" + array[1] + "-" + array[0];
				}
				if (HTKKFormater.isNullString(val5.InnerText))
				{
					val5.InnerText = DateTime.Now.ToString("yyyy-MM-dd");
					val5.Attributes.Remove(val5.Attributes["xsi:nil"]);
				}
			}
			XmlNode val6 = val.SelectSingleNode("TTinTKhaiThue/TKhaiThue/ngayLapTKhai");
			if (val6 != null)
			{
				val6.InnerText = DateTime.Now.Year + "-" + DateTime.Now.Month.ToString().PadLeft(2, '0') + "-" + DateTime.Now.Day.ToString().PadLeft(2, '0');
			}
			XmlNode val7 = val.SelectSingleNode("TTinTKhaiThue/TKhaiThue/soLan");
			if (val7 != null && string.IsNullOrEmpty(val7.InnerText))
			{
				val7.InnerText = "0";
			}
			XmlNode val8 = val.SelectSingleNode("TTinTKhaiThue/DLyThue/ngayKyHDDLyThue");
			if (val8 != null)
			{
				string innerText = val8.InnerText;
				if (innerText.Length >= 8)
				{
					val8.InnerText = innerText.Substring(6, 4) + "-" + innerText.Substring(3, 2).PadLeft(2, '0') + "-" + innerText.Substring(0, 2).PadLeft(2, '0');
				}
				if (string.IsNullOrEmpty(innerText))
				{
					XmlNode val9 = val8.OwnerDocument.CreateNode((XmlNodeType)2, "xsi:nil", "http://www.w3.org/2001/XMLSchema-instance");
					val9.Value = "true";
					((XmlNamedNodeMap)val8.Attributes).SetNamedItem(val9);
					val8.Normalize();
				}
			}
			XmlNode val10 = val.SelectSingleNode("TTinTKhaiThue/NNT/mst");
			if (val10 != null && val10.InnerXml.Trim().Length == 13)
			{
				val10.InnerXml = val10.InnerXml.Substring(0, 10) + "-" + val10.InnerXml.Substring(10, 3);
			}
			val10 = val.SelectSingleNode("TTinTKhaiThue/DLyThue/mstDLyThue");
			if (val10 != null && val10.InnerXml.Trim().Length == 13)
			{
				val10.InnerXml = val10.InnerXml.Substring(0, 10) + "-" + val10.InnerXml.Substring(10, 3);
			}
			if (WinSettings.MenuID == "B2" || WinSettings.MenuID == "O9" || WinSettings.MenuID == "M6" || WinSettings.MenuID == "O8")
			{
				val10 = val.SelectSingleNode("TTinTKhaiThue/mst");
				if (val10 != null && val10.InnerXml.Trim().Length == 13)
				{
					val10.InnerXml = val10.InnerXml.Substring(0, 10) + "-" + val10.InnerXml.Substring(10, 3);
				}
				XmlNode val11 = val.SelectSingleNode("TTinTKhaiThue/ngayNhanTKhai");
				XmlNode val9 = val11.OwnerDocument.CreateNode((XmlNodeType)2, "xsi:nil", "http://www.w3.org/2001/XMLSchema-instance");
				val9.Value = "true";
				if (val11.InnerText == "")
				{
					((XmlNamedNodeMap)val11.Attributes).SetNamedItem(val9);
				}
				val11 = val.SelectSingleNode("TTinTKhaiThue/nhanKQQuaBuuDien");
				val9 = val11.OwnerDocument.CreateNode((XmlNodeType)2, "xsi:nil", "http://www.w3.org/2001/XMLSchema-instance");
				val9.Value = "true";
				if (val11.InnerText == "")
				{
					((XmlNamedNodeMap)val11.Attributes).SetNamedItem(val9);
				}
			}
			if (WinSettings.KieuKy.Equals("22"))
			{
				val.SelectSingleNode("TTinTKhaiThue/TKhaiThue/KyKKhaiThue/kieuKy").InnerText = "Y";
			}
			if ((WinSettings.MenuID == "G7" || WinSettings.MenuID == "Q4") && WinSettings.TTPeriodInfo.KieuKy.Equals("D_D"))
			{
				val.SelectSingleNode("TTinTKhaiThue/TKhaiThue/KyKKhaiThue/kieuKy").InnerText = "D";
			}
			if (WinSettings.MenuID == "10" || WinSettings.MenuID == "67")
			{
				val.SelectSingleNode("TTinTKhaiThue/TKhaiThue/KyKKhaiThue/kieuKy").InnerText = "D";
			}
			if ((WinSettings.MenuID == "G4" || WinSettings.MenuID == "N5" || WinSettings.MenuID == "H5" || WinSettings.MenuID == "O7" || WinSettings.MenuID == "P5" || WinSettings.MenuID == "U6" || "U8~".Contains(WinSettings.MenuID)) && WinSettings.TTPeriodInfo.LoaiToKhai == "T")
			{
				val.SelectSingleNode("TTinTKhaiThue/TKhaiThue/loaiTKhai").InnerText = "B";
				val.SelectSingleNode("TTinTKhaiThue/TKhaiThue/soLan").InnerText = soLanTT;
			}
			if (WinSettings.MenuID == "G3")
			{
				if (WinSettings.TTPeriodInfo.KieuKy.Equals("M"))
				{
					val.SelectSingleNode("TTinTKhaiThue/TKhaiThue/KyKKhaiThue/kieuKy").InnerText = "DK";
					val.SelectSingleNode("TTinTKhaiThue/TKhaiThue/KyKKhaiThue/kyKKhaiTuThang").InnerXml = SecurityElement.Escape(WinSettings.TTPeriodInfo.kyKKhai);
					val.SelectSingleNode("TTinTKhaiThue/TKhaiThue/KyKKhaiThue/kyKKhaiDenThang").InnerXml = SecurityElement.Escape(WinSettings.TTPeriodInfo.kyKKhai);
				}
				else if (WinSettings.TTPeriodInfo.KieuKy.Equals("D"))
				{
					val.SelectSingleNode("TTinTKhaiThue/TKhaiThue/KyKKhaiThue/kieuKy").InnerText = "LD";
				}
				val.SelectSingleNode("TTinTKhaiThue/TKhaiThue/soLan").InnerText = WinSettings.TTPeriodInfo.LanXuatBan;
			}
			if (WinSettings.MenuID == "N2" || WinSettings.MenuID == "N6")
			{
				val.SelectSingleNode("TTinTKhaiThue/TKhaiThue/KyKKhaiThue/kieuKy").InnerText = "M";
				val.SelectSingleNode("TTinTKhaiThue/TKhaiThue/KyKKhaiThue/kyKKhai").InnerText = WinSettings.TTPeriodInfo.kyKKhai.Substring(3);
			}
			if ((WinSettings.MenuID == "G7" || WinSettings.MenuID == "G8" || WinSettings.MenuID == "G5" || WinSettings.MenuID == "G6" || WinSettings.MenuID == "L1") && (string.IsNullOrEmpty(val.SelectSingleNode("TTinTKhaiThue/TKhaiThue/KyKKhaiThue/kyKKhaiTuNgay").InnerText) || string.IsNullOrEmpty(val.SelectSingleNode("TTinTKhaiThue/TKhaiThue/KyKKhaiThue/kyKKhaiDenNgay").InnerText)))
			{
				DateTime ngayDauKyKeKhai = GeneralFunctions.GetNgayDauKyKeKhai(WinSettings.TTinChung, WinSettings.TTPeriodInfo);
				string text2 = ((!string.IsNullOrEmpty(WinSettings.TTinChung.kyKKhaiTuThang)) ? ("01/" + SecurityElement.Escape(WinSettings.TTinChung.kyKKhaiTuThang)) : SecurityElement.Escape(ngayDauKyKeKhai.ToString("dd/MM/yyyy")));
				val.SelectSingleNode("TTinTKhaiThue/TKhaiThue/KyKKhaiThue/kyKKhaiTuNgay").InnerXml = text2;
				DateTime ngayCuoiKyKeKhai = GeneralFunctions.GetNgayCuoiKyKeKhai(WinSettings.TTinChung, WinSettings.TTPeriodInfo);
				DateTime dateTime = (string.IsNullOrEmpty(WinSettings.TTinChung.kyKKhaiDenThang) ? DateTime.MinValue : DateTime.ParseExact("01/" + WinSettings.TTinChung.kyKKhaiDenThang, "dd/MM/yyyy", CultureInfo.CurrentCulture));
				DateTime dateTime2 = (string.IsNullOrEmpty(WinSettings.TTinChung.kyKKhaiDenThang) ? DateTime.MinValue : dateTime.AddMonths(1).AddDays(-1.0));
				string text3 = ((!string.IsNullOrEmpty(WinSettings.TTinChung.kyKKhaiDenThang)) ? dateTime2.ToString("dd/MM/yyyy") : SecurityElement.Escape(ngayCuoiKyKeKhai.ToString("dd/MM/yyyy")));
				val.SelectSingleNode("TTinTKhaiThue/TKhaiThue/KyKKhaiThue/kyKKhaiDenNgay").InnerXml = text3;
				WinSettings.TTinChung.kyKKhaiTuNgay = text2;
				WinSettings.TTinChung.kyKKhaiDenNgay = text3;
			}
			if ("H6" == WinSettings.MenuID)
			{
				XmlNode val12 = val.SelectSingleNode("TTinTKhaiThue/TKhaiThue/ngayKy");
				if (val12 != null && !string.IsNullOrEmpty(val12.InnerText))
				{
					XmlNode val13 = val.SelectSingleNode("TTinTKhaiThue/TKhaiThue/KyKKhaiThue/kyKKhai");
					val13.InnerXml = DateTime.ParseExact(val12.InnerText.Trim(), "yyyy-MM-dd", CultureInfo.CurrentCulture).ToString("MM/yyyy");
				}
			}
			if ("Q4" == WinSettings.MenuID)
			{
				val.SelectSingleNode("TTinTKhaiThue/TKhaiThue/KyKKhaiThue/kyKKhai").InnerText = "";
				if (WinSettings.TTPeriodInfo.KieuKy == "Y" && (string.IsNullOrEmpty(val.SelectSingleNode("TTinTKhaiThue/TKhaiThue/KyKKhaiThue/kyKKhaiTuNgay").InnerText) || string.IsNullOrEmpty(val.SelectSingleNode("TTinTKhaiThue/TKhaiThue/KyKKhaiThue/kyKKhaiDenNgay").InnerText)))
				{
					string innerText2 = "01/01/" + WinSettings.TTPeriodInfo.TuThang;
					DateTime dateTime3 = new DateTime(Convert.ToInt32(WinSettings.TTPeriodInfo.DenThang), 1, 1).AddYears(1).AddDays(-1.0);
					val.SelectSingleNode("TTinTKhaiThue/TKhaiThue/KyKKhaiThue/kyKKhaiTuNgay").InnerText = innerText2;
					val.SelectSingleNode("TTinTKhaiThue/TKhaiThue/KyKKhaiThue/kyKKhaiDenNgay").InnerText = dateTime3.ToString("dd/MM/yyyy");
				}
			}
			if (WinSettings.MenuID == "P7" || WinSettings.MenuID == "Q7" || WinSettings.MenuID == "Q6" || WinSettings.MenuID == "Q5")
			{
				val.SelectSingleNode("TTinTKhaiThue/TKhaiThue/tenCQTNoiNop").InnerText = "Chi cục Thuế doanh nghiệp lớn";
				val.SelectSingleNode("TTinTKhaiThue/TKhaiThue/maCQTNoiNop").InnerText = "82500";
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static XmlNode xmlExportTTinChungKHBS(XmlDocument xmlExportContent)
	{
		XmlNode val = null;
		try
		{
			val = ((XmlNode)xmlExportContent).SelectSingleNode("PL01_1_KHBS/Header");
			if (val == null)
			{
				val = ((XmlNode)xmlExportContent).SelectSingleNode("PL01_KHBS/Header");
			}
			if (val == null)
			{
				return val;
			}
			val.SelectSingleNode("maTKhai").InnerXml = SecurityElement.Escape(WinSettings.TTinChung.maTKhai);
			val.SelectSingleNode("tenTKhai").InnerXml = SecurityElement.Escape(WinSettings.TTinChung.tenTKhai);
			val.SelectSingleNode("KyKKhaiThue/kieuKy").InnerXml = SecurityElement.Escape(WinSettings.TTinChung.kieuKy);
			val.SelectSingleNode("KyKKhaiThue/kyKKhai").InnerXml = SecurityElement.Escape(WinSettings.TTinChung.kyKKhai);
			val.SelectSingleNode("KyKKhaiThue/kyKKhaiTuThang").InnerXml = SecurityElement.Escape(WinSettings.TTinChung.kyKKhaiTuThang);
			val.SelectSingleNode("KyKKhaiThue/kyKKhaiDenThang").InnerXml = SecurityElement.Escape(WinSettings.TTinChung.kyKKhaiDenThang);
			val.SelectSingleNode("soLan").InnerXml = SecurityElement.Escape(WinSettings.TTinChung.soLan);
			DateTime ngayDauKyKeKhai = GeneralFunctions.GetNgayDauKyKeKhai(WinSettings.TTinChung, WinSettings.TTPeriodInfo);
			string innerXml = ((!string.IsNullOrEmpty(WinSettings.TTinChung.kyKKhaiTuNgay)) ? SecurityElement.Escape(WinSettings.TTinChung.kyKKhaiTuNgay) : SecurityElement.Escape(ngayDauKyKeKhai.ToString("dd/MM/yyyy")));
			val.SelectSingleNode("KyKKhaiThue/kyKKhaiTuNgay").InnerXml = innerXml;
			DateTime ngayCuoiKyKeKhai = GeneralFunctions.GetNgayCuoiKyKeKhai(WinSettings.TTinChung, WinSettings.TTPeriodInfo);
			string innerXml2 = ((!string.IsNullOrEmpty(WinSettings.TTinChung.kyKKhaiDenNgay)) ? SecurityElement.Escape(WinSettings.TTinChung.kyKKhaiDenNgay) : SecurityElement.Escape(ngayCuoiKyKeKhai.ToString("dd/MM/yyyy")));
			val.SelectSingleNode("KyKKhaiThue/kyKKhaiDenNgay").InnerXml = innerXml2;
			if (WinSettings.TTinChung.mst.Trim().Length == 13)
			{
				val.SelectSingleNode("mst").InnerXml = SecurityElement.Escape(WinSettings.TTinChung.mst.Substring(0, 10) + "-" + WinSettings.TTinChung.mst.Substring(10, 3));
			}
			else
			{
				val.SelectSingleNode("mst").InnerXml = SecurityElement.Escape(WinSettings.TTinChung.mst);
			}
			val.SelectSingleNode("tenNNT").InnerXml = SecurityElement.Escape(WinSettings.TTinChung.tenNNT);
			if (WinSettings.TTinChung.mstDLyThue.Trim().Length == 13)
			{
				val.SelectSingleNode("mstDLyThue").InnerXml = SecurityElement.Escape(WinSettings.TTinChung.mstDLyThue.Substring(0, 10) + "-" + WinSettings.TTinChung.mstDLyThue.Substring(10, 3));
			}
			else
			{
				val.SelectSingleNode("mstDLyThue").InnerXml = SecurityElement.Escape(WinSettings.TTinChung.mstDLyThue);
			}
			val.SelectSingleNode("tenDLyThue").InnerXml = SecurityElement.Escape(WinSettings.TTinChung.tenDLyThue);
			val.SelectSingleNode("soHDongDLyThue").InnerXml = SecurityElement.Escape(WinSettings.TTinChung.soHDongDLyThue);
			val.SelectSingleNode("ngayKyHDDLyThue").InnerXml = (string.IsNullOrEmpty(WinSettings.TTinChung.ngayKyHDDLyThue) ? "" : DateTime.ParseExact(WinSettings.TTinChung.ngayKyHDDLyThue, "dd/MM/yyyy", CultureInfo.CurrentCulture).ToString("yyyy-MM-dd"));
			XmlNode mCT = WinFnProcess.GetMCT01("MapMCT.xml", WinSettings.ToKhaiInfo.MenuID);
			string text = ((mCT.Attributes["LoaiTienPath"] != null) ? ((XmlNode)mCT.Attributes["LoaiTienPath"]).Value.Trim() : "");
			if (string.IsNullOrEmpty(text))
			{
				if (val.ParentNode != null && val.ParentNode.SelectSingleNode("ma_DonViTien") != null)
				{
					val.ParentNode.SelectSingleNode("ma_DonViTien").InnerXml = "VND";
					val.ParentNode.SelectSingleNode("ten_DonViTien").InnerXml = "Đồng Việt Nam";
				}
			}
			else
			{
				foreach (HTKKCMCCellInfo staticCell in WinSettings.ToKhaiInfo.SheetInfos[0].StaticCells)
				{
					if (staticCell != null && staticCell.Path == text)
					{
						val.ParentNode.SelectSingleNode("ma_DonViTien").InnerXml = ((staticCell.Value == "USD") ? "USD" : "VND");
						val.ParentNode.SelectSingleNode("ten_DonViTien").InnerXml = ((staticCell.Value == "USD") ? "Đô la Mỹ" : "Đồng Việt Nam");
						break;
					}
				}
			}
			XmlNode val2 = val.SelectSingleNode("mstDLyThue");
			XmlNode val3 = val2.OwnerDocument.CreateNode((XmlNodeType)2, "xsi:nil", "http://www.w3.org/2001/XMLSchema-instance");
			val3.Value = "true";
			if (val2.InnerText == "")
			{
				((XmlNamedNodeMap)val2.Attributes).SetNamedItem(val3);
			}
			val2 = val.SelectSingleNode("ngayKyHDDLyThue");
			val3 = val2.OwnerDocument.CreateNode((XmlNodeType)2, "xsi:nil", "http://www.w3.org/2001/XMLSchema-instance");
			val3.Value = "true";
			if (val2.InnerText == "")
			{
				((XmlNamedNodeMap)val2.Attributes).SetNamedItem(val3);
			}
			if (WinSettings.MenuID == "G7" && WinSettings.TTinChung.kieuKy == "D_D")
			{
				val.SelectSingleNode("KyKKhaiThue/kieuKy").InnerXml = "D";
			}
			if ((WinSettings.MenuID == "C9" || WinSettings.MenuID == "D6" || WinSettings.MenuID == "X9" || WinSettings.MenuID == "Y1") && WinSettings.TTinChung.kieuKy == "D_D")
			{
				val.SelectSingleNode("KyKKhaiThue/kieuKy").InnerXml = "Y";
			}
			return val;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return val;
	}

	public static void xmlExportTTinChungHDDT(XmlDocument xmlExportContent)
	{
		try
		{
			XmlNode val = ((XmlNode)xmlExportContent).SelectSingleNode("TKhai/DLTKhai/TTChung");
			xmlExportTTinChungProcessNodeHDDT(val);
			XmlNode val2 = val.SelectSingleNode("NLap");
			if (val2 != null)
			{
				string innerText = val2.InnerText;
				if (innerText.Length >= 8)
				{
					string[] array = innerText.Split(new char[1] { '-' });
					if (array.Length == 3)
					{
						val2.InnerText = array[2] + "/" + array[1] + "/" + array[0];
					}
				}
			}
			XmlNode val3 = val.SelectSingleNode("MST");
			if (val3 != null && val3.InnerXml.Trim().Length == 13)
			{
				val3.InnerXml = val3.InnerXml.Substring(0, 10) + "-" + val3.InnerXml.Substring(10, 3);
			}
			val3 = val.SelectSingleNode("MSTDLThue");
			if (val3 != null && val3.InnerXml.Trim().Length == 13)
			{
				val3.InnerXml = val3.InnerXml.Substring(0, 10) + "-" + val3.InnerXml.Substring(10, 3);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void xmlExportTTinChungProcessNode(XmlNode xChildNode)
	{
		//IL_039c: Unknown result type (might be due to invalid IL or missing references)
		//IL_03a3: Expected O, but got Unknown
		try
		{
			bool flag = false;
			string text = "";
			PropertyInfo[] properties = typeof(HTKKCMCTTinChungInfo).GetProperties();
			foreach (PropertyInfo propertyInfo in properties)
			{
				if (!(xChildNode.Name == propertyInfo.Name))
				{
					continue;
				}
				if ((xChildNode.Name == "kyKKhaiTuNgay" || xChildNode.Name == "kyKKhaiDenNgay") && WinSettings.MenuID == "39")
				{
					break;
				}
				if ((xChildNode.Name == "tenNVienDLyThue" || xChildNode.Name == "cchiHNghe") && (WinSettings.MenuID == "91" || WinSettings.MenuID == "10"))
				{
					continue;
				}
				if ((HTKKFormater.isNullString(xChildNode.InnerXml) && WinSettings.MenuID == "07") || WinSettings.MenuID != "07")
				{
					text = propertyInfo.GetValue(WinSettings.TTinChung, null).ToString();
					if (propertyInfo.Name == "pbanTKhaiXML" && text.IsNullOrEmpty())
					{
						text = WinSettings.ToKhaiInfo.VersionXML;
					}
					if (!SecurityElement.IsValidText(text))
					{
						text = SecurityElement.Escape(text);
					}
					xChildNode.InnerXml = SecurityElement.Escape(propertyInfo.GetValue(WinSettings.TTinChung, null).ToString());
					if (propertyInfo.Name == "pbanTKhaiXML")
					{
						xChildNode.InnerXml = text;
						if (WinSettings.MenuID == "Q9")
						{
							DateTime dateTime = new DateTime(2026, 5, 1);
							string innerXml = "";
							string text2 = "";
							if (WinSettings.TTPeriodInfo.KieuKy == "D")
							{
								text2 = WinSettings.TTPeriodInfo.kyKKhai;
								DateTime dateTime2 = DateTime.ParseExact(text2, "dd/MM/yyyy", CultureInfo.CurrentCulture);
								innerXml = ((!(dateTime2 < dateTime)) ? "2.9.8" : "2.8.9");
							}
							if (WinSettings.TTPeriodInfo.KieuKy == "M")
							{
								text2 = "01/" + WinSettings.TTPeriodInfo.kyKKhai;
								DateTime dateTime2 = DateTime.ParseExact(text2, "dd/MM/yyyy", CultureInfo.CurrentCulture);
								innerXml = ((!(dateTime2 < dateTime)) ? "2.9.8" : "2.8.9");
							}
							if (WinSettings.TTPeriodInfo.KieuKy == "Q")
							{
								innerXml = ((WinSettings.TTPeriodInfo.Nam == "2026" && WinSettings.TTPeriodInfo.Quy == "1") ? "2.8.9" : ((Convert.ToInt32(WinSettings.TTPeriodInfo.Nam) >= 2026) ? "2.9.8" : "2.8.9"));
							}
							xChildNode.InnerXml = innerXml;
						}
					}
				}
				flag = true;
				break;
			}
			if (flag)
			{
				return;
			}
			foreach (XmlNode childNode in xChildNode.ChildNodes)
			{
				XmlNode xChildNode2 = childNode;
				xmlExportTTinChungProcessNode(xChildNode2);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void xmlExportTTinChungProcessNodeHDDT(XmlNode xChildNode)
	{
		//IL_001a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0020: Expected O, but got Unknown
		try
		{
			foreach (XmlNode childNode in xChildNode.ChildNodes)
			{
				XmlNode val = childNode;
				if (val.Name == "PBan")
				{
					val.InnerXml = SecurityElement.Escape(WinSettings.TTinChung.pbanTKhaiXML);
				}
				if (val.Name == "MSo")
				{
					val.InnerXml = SecurityElement.Escape(WinSettings.TTinChung.moTaBMau);
				}
				if (val.Name == "Ten")
				{
					val.InnerXml = SecurityElement.Escape(WinSettings.TTinChung.tenTKhai);
				}
				if (val.Name == "LKTThue")
				{
					val.InnerXml = SecurityElement.Escape(WinSettings.TTinChung.kieuKy);
				}
				if (val.Name == "KTThue")
				{
					val.InnerXml = SecurityElement.Escape(WinSettings.TTinChung.kyKKhai);
				}
				if (val.Name == "TNNT")
				{
					val.InnerXml = SecurityElement.Escape(WinSettings.TTinChung.tenNNT);
				}
				if (val.Name == "MST")
				{
					val.InnerXml = SecurityElement.Escape(WinSettings.TTinChung.mst);
				}
				if (val.Name == "TDLThue")
				{
					val.InnerXml = SecurityElement.Escape(WinSettings.TTinChung.tenDLyThue);
				}
				if (val.Name == "MSTDLThue")
				{
					val.InnerXml = SecurityElement.Escape(WinSettings.TTinChung.mstDLyThue);
				}
				if (val.Name == "DVTTe")
				{
					val.InnerXml = SecurityElement.Escape("VND");
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static XmlDocument WinDoWorkGetExportSheetXML(SheetView sView, HTKKCMCSheetInfo sInfo)
	{
		//IL_0001: Unknown result type (might be due to invalid IL or missing references)
		//IL_0007: Expected O, but got Unknown
		XmlDocument val = new XmlDocument();
		try
		{
			string interfaceTemplateXml = WinSettings.ToKhaiInfo.InterfaceTemplateXml;
			interfaceTemplateXml = ((sInfo.SheetName.Equals("KHBS") && !WinSettings.ToKhaiInfo.KieuKHBS.Equals("")) ? GetFileNameKHBSTheoThongTu(interfaceTemplateXml) : (sInfo.SheetID.Equals("01-KHBS") ? ("I7,I4,C7,I8,C8".Contains(WinSettings.MenuID) ? (WinSettings.ToKhaiInfo.InterfaceTemplateXml + "KHBS_01_TT80_DVT_xml.xml") : ((!"D1,N8,G9,H2,H4,T8,E8,L8,L7,N2,G2,R5,Y6,Z8,Z5".Contains(WinSettings.MenuID)) ? (WinSettings.ToKhaiInfo.InterfaceTemplateXml + "KHBS_01_TT80_xml.xml") : (WinSettings.ToKhaiInfo.InterfaceTemplateXml + "KHBS_01_DIA_BAN_xml.xml"))) : ((!sInfo.SheetID.Equals("01-1-KHBS")) ? string.Format("{0}_xml_{1}.xml", WinSettings.ToKhaiInfo.InterfaceTemplateXml + sInfo.DataFile, WinSettings.ToKhaiInfo.VersionXML.Replace(".", "")) : ((!"I7,I4,C7,I8,C8".Contains(WinSettings.MenuID)) ? (WinSettings.ToKhaiInfo.InterfaceTemplateXml + "KHBS_01_1_TT80_xml.xml") : (WinSettings.ToKhaiInfo.InterfaceTemplateXml + "KHBS_01_1_TT80_DVT_xml.xml")))));
			val.Load(interfaceTemplateXml);
			for (int i = 0; i < sInfo.StaticCells.Count; i++)
			{
				if (sInfo.StaticCells[i].Path == null || sInfo.StaticCells[i].Path.Length <= 0 || sInfo.StaticCells[i].CellTemplate == null)
				{
					continue;
				}
				try
				{
					XmlNode val2 = ((XmlNode)val).SelectSingleNode(sInfo.StaticCells[i].Path);
					if (val2 == null)
					{
						continue;
					}
					if (sInfo.StaticCells[i].Controltype == "16" && (WinSettings.MenuID == "78" || WinSettings.MenuID == "A1" || WinSettings.MenuID == "A3" || "X1".Equals(WinSettings.TTPeriodInfo.menuID)))
					{
						sInfo.StaticCells[i].DefaultValue = "0,00";
					}
					string text = sInfo.StaticCells[i].Controltype;
					string defaultValue = sInfo.StaticCells[i].DefaultValue;
					if ("I7,I4,C7,I8,C8".Contains(WinSettings.MenuID) || sInfo.SheetName == "01-KHBS" || sInfo.SheetName == "01-1/KHBS")
					{
						string controlTypeFromSView_01_1_KHBS = GetControlTypeFromSView_01_1_KHBS(sView, sInfo.StaticCells[i].CellTemplate.RowIndex, sInfo.StaticCells[i].CellTemplate.ColIndex, ref defaultValue);
						if (!string.IsNullOrEmpty(controlTypeFromSView_01_1_KHBS))
						{
							text = controlTypeFromSView_01_1_KHBS;
						}
					}
					string text2 = GeneralFunctions.GetValueExportXML(sView, text, sInfo.StaticCells[i].CellTemplate.RowIndex, sInfo.StaticCells[i].CellTemplate.ColIndex, defaultValue, sInfo.SheetName).ToString().Trim();
					if (text == "26" && text2.Trim().Length == 13)
					{
						text2 = text2.Substring(0, 10) + "-" + text2.Substring(10, 3);
					}
					if (!string.IsNullOrEmpty(text2))
					{
						val2.InnerText = text2;
					}
					if (text == "2")
					{
						if (WinSettings.MenuID == "63" || WinSettings.MenuID == "L1" || WinSettings.MenuID == "A4" || WinSettings.MenuID == "H6")
						{
							val2.InnerXml = ((val2.InnerXml == "1" || val2.InnerXml.ToLower() == "true") ? "true" : "false");
						}
						if (WinSettings.MenuID == "78" || "U9".Equals(WinSettings.MenuID))
						{
							val2.InnerXml = val2.InnerXml.Equals("1").ToString().ToLower();
						}
					}
					if (!string.IsNullOrEmpty(val2.InnerXml))
					{
						continue;
					}
					switch (text)
					{
					default:
						if (text == "94")
						{
							break;
						}
						goto end_IL_01f4;
					case "14":
					case "26":
					case "33":
					case "96":
						break;
					}
					XmlNode val3 = val2.OwnerDocument.CreateNode((XmlNodeType)2, "xsi:nil", "http://www.w3.org/2001/XMLSchema-instance");
					val3.Value = "true";
					((XmlNamedNodeMap)val2.Attributes).SetNamedItem(val3);
					end_IL_01f4:;
				}
				catch (Exception ex)
				{
					HTKKFile.Log(sInfo.StaticCells[i].Path);
					throw ex;
				}
			}
			if (sInfo.Tables != null && sInfo.Tables.Count > 0)
			{
				for (int i = 0; i < sInfo.Tables.Count; i++)
				{
					string tableID = sInfo.Tables[i].TableID;
					string tableName = sInfo.Tables[i].TableName;
					HTKKCMCTableInfo hTKKCMCTableInfo = sInfo.Tables[i];
					if (!(tableID == "1"))
					{
						continue;
					}
					XmlNode val4 = ((XmlNode)val).SelectSingleNode(hTKKCMCTableInfo.TablePath);
					XmlNode xmlTableTemplateNode = val4.Clone();
					ExportDataTableToTableNode(sView, val4, hTKKCMCTableInfo, hTKKCMCTableInfo, sInfo.SheetIndex);
					if (val4.Attributes["id"] != null)
					{
						((XmlNode)val4.Attributes["id"]).Value = "ID_" + tableID;
					}
					if (val4.Attributes["ID"] != null)
					{
						((XmlNode)val4.Attributes["ID"]).Value = "ID_" + tableID;
					}
					for (int j = 0; j < sInfo.Tables.Count; j++)
					{
						if ((tableName == sInfo.Tables[j].TableName) & (tableID != sInfo.Tables[j].TableID))
						{
							InsertExportTableNode(sView, ((XmlNode)val).SelectSingleNode(hTKKCMCTableInfo.TablePath), xmlTableTemplateNode, sInfo.Tables[j], hTKKCMCTableInfo, sInfo.SheetIndex);
						}
					}
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			val = null;
		}
		return val;
	}

	private static string GetFileNameKHBSTheoThongTu(string xmlFilePath)
	{
		switch (WinSettings.ToKhaiInfo.MaFile)
		{
		case "01_GTGT":
			xmlFilePath += "KHBS_01_GTGT_xml.xml";
			break;
		case "02_GTGT":
		case "03_GTGT":
		case "04_GTGT":
		case "05_GTGT":
		case "01A_TNDN":
		case "01B_TNDN":
		case "02_TNDN":
		case "01_NTNN":
		case "03_NTNN":
		case "01_TAIN":
		case "01_TTDB":
		case "01_TTDB_TT195":
		case "01_BVMT":
		case "01_TBVMT":
		case "01_TD_GTGT":
		case "01A_TNDN_DK":
		case "01B_TNDN_DK":
		case "01_TAIN_DK":
		case "02_TAIN_DK":
		case "03_TD_TAIN":
		case "01_BCTL_DK":
		case "01_PHLP":
		case "04_TNDN":
		case "01_MBAI":
		case "03_TNDN_DK":
		case "02_TNDN_DK":
		case "01_TNDN_DK":
		case "01_QT_LNCL":
		case "01_CTLNDC":
		case "01_AC":
			xmlFilePath += "KHBS_TT156_xml.xml";
			break;
		default:
			xmlFilePath += "KHBS_QT_TT156_xml.xml";
			break;
		}
		return xmlFilePath;
	}

	public static void ExportDataTableToTableNode(SheetView sView, XmlNode xmlTableNode, HTKKCMCTableInfo tableInfo, HTKKCMCTableInfo rootTableInfo, int sheetIndex = -1)
	{
		try
		{
			ExportDataStaticTableToXMLNode(sView, xmlTableNode, tableInfo.Header, sheetIndex);
			ExportDataStaticTableToXMLNode(sView, xmlTableNode, tableInfo.Footer, sheetIndex);
			if (rootTableInfo.RowsValues.Count <= 0)
			{
				return;
			}
			List<XmlNode> list = new List<XmlNode>();
			XmlNode parentNode;
			XmlNode val;
			if (tableInfo.RowsInfo.RowPath == null || tableInfo.RowsInfo.RowPath.Length <= 0)
			{
				parentNode = xmlTableNode.ParentNode;
				val = xmlTableNode.Clone();
			}
			else
			{
				parentNode = xmlTableNode.SelectSingleNode(tableInfo.RowsInfo.RowPath).ParentNode;
				val = xmlTableNode.SelectSingleNode(tableInfo.RowsInfo.RowPath);
			}
			string rowDataDefault = tableInfo.GetRowDataDefault();
			bool flag = false;
			StringBuilder stringBuilder = new StringBuilder();
			XmlNode val2 = val.Clone();
			for (int i = 0; i < tableInfo.RowsValues.Count; i++)
			{
				if (!(rowDataDefault != tableInfo.GetRowData(sView, i)) && i != 0)
				{
					continue;
				}
				if (WinSettings.MenuID == "33" && WinSettings.ToKhaiInfo.SheetInfos.Count > 1 && WinSettings.ToKhaiInfo.SheetInfos[1] != null && WinSettings.ToKhaiInfo.SheetInfos[1].Tables[0].RowsValues.Count > 200000)
				{
					flag = true;
				}
				else
				{
					val2 = val.Clone();
				}
				string text = ((val2.Attributes["minOccurs"] == null) ? "" : ((XmlNode)val2.Attributes["minOccurs"]).Value);
				bool flag2 = true;
				for (int j = 0; j < tableInfo.RowsInfo.Cells.Count; j++)
				{
					if (tableInfo.RowsInfo.Cells[j].Path == null || tableInfo.RowsInfo.Cells[j].Path.Length <= 0)
					{
						continue;
					}
					XmlNode val3 = val2.SelectSingleNode(tableInfo.RowsInfo.Cells[j].Path);
					if (val3 == null)
					{
						continue;
					}
					string obj = ((val3.Attributes["nillable"] == null) ? "" : ((XmlNode)val3.Attributes["nillable"]).Value);
					string text2 = tableInfo.RowsInfo.Cells[j].Controltype;
					string defaultValue = tableInfo.RowsInfo.Cells[j].DefaultValue;
					object value = sView.Cells[tableInfo.Location.DataLocationTemplateRowStart + i - 1, tableInfo.RowsInfo.Cells[j].CellTemplate.ColIndex - 1].Value;
					if ("I7,I4,C7,I8,C8".Contains(WinSettings.MenuID) || (WinSettings.ToKhaiInfo.SheetInfos[sheetIndex] != null && WinSettings.ToKhaiInfo.SheetInfos[sheetIndex].SheetName == "01-KHBS") || (WinSettings.ToKhaiInfo.SheetInfos[sheetIndex] != null && WinSettings.ToKhaiInfo.SheetInfos[sheetIndex].SheetName == "01-1/KHBS"))
					{
						string controlTypeFromSView_01_1_KHBS = GetControlTypeFromSView_01_1_KHBS(sView, tableInfo.Location.DataLocationTemplateRowStart + i, tableInfo.RowsInfo.Cells[j].CellTemplate.ColIndex, ref defaultValue);
						if (!string.IsNullOrEmpty(controlTypeFromSView_01_1_KHBS))
						{
							text2 = controlTypeFromSView_01_1_KHBS;
						}
					}
					string value2 = GeneralFunctions.GetValueExportXML(sView, text2, tableInfo.Location.DataLocationTemplateRowStart + i, tableInfo.RowsInfo.Cells[j].CellTemplate.ColIndex, defaultValue, WinSettings.ToKhaiInfo.SheetInfos[sheetIndex].SheetName).Trim();
					value2 = RemoveValueSpecial(tableInfo.RowsInfo.Cells[j], value2);
					if (text2 == "26" && value2.Trim().Length == 13)
					{
						value2 = value2.Substring(0, 10) + "-" + value2.Substring(10, 3);
					}
					if (ConvertHelpers.ConvertToBool(obj))
					{
						val3.Attributes.Remove(val3.Attributes["nillable"]);
					}
					if (!string.IsNullOrEmpty(value2))
					{
						if (!value2.Equals(defaultValue))
						{
							flag2 = false;
						}
						if (defaultValue == "0,00" && !value2.Contains("."))
						{
							val3.InnerText = value2;
						}
						else
						{
							val3.InnerText = value2;
						}
					}
					else
					{
						if (ConvertHelpers.ConvertToBool(obj))
						{
							XmlNode val4 = val3.OwnerDocument.CreateNode((XmlNodeType)2, "xsi:nil", "http://www.w3.org/2001/XMLSchema-instance");
							val4.Value = "true";
							((XmlNamedNodeMap)val3.Attributes).SetNamedItem(val4);
						}
						if (text2 != "14" && text2 != "26" && text2 != "33" && text2 != "96" && text2 != "94")
						{
							val3.InnerXml = value2;
						}
					}
					switch (text2)
					{
					case "2":
						if (WinSettings.MenuID == "33" || WinSettings.MenuID == "A4" || WinSettings.MenuID == "99" || WinSettings.MenuID == "H6" || WinSettings.MenuID == "B2" || WinSettings.MenuID == "O9" || WinSettings.MenuID == "M6" || WinSettings.MenuID == "O8")
						{
							val3.InnerXml = ((val3.InnerXml == "1") ? "true" : "false");
						}
						break;
					}
					if (!string.IsNullOrEmpty(val3.InnerXml.Trim()))
					{
						continue;
					}
					switch (text2)
					{
					default:
						if (!(text2 == "94"))
						{
							continue;
						}
						break;
					case "14":
					case "26":
					case "33":
					case "96":
						break;
					}
					XmlNode val5 = val3.OwnerDocument.CreateNode((XmlNodeType)2, "xsi:nil", "http://www.w3.org/2001/XMLSchema-instance");
					val5.Value = "true";
					((XmlNamedNodeMap)val3.Attributes).SetNamedItem(val5);
				}
				if (val2.Attributes["id"] != null)
				{
					((XmlNode)val2.Attributes["id"]).Value = "ID_" + (i + 1);
				}
				if (val2.Attributes["ID"] != null)
				{
					((XmlNode)val2.Attributes["ID"]).Value = "ID_" + (i + 1);
				}
				if (!flag)
				{
					if (text == "0")
					{
						if (!flag2)
						{
							val2.Attributes.Remove(val2.Attributes["minOccurs"]);
							list.Add(val2);
						}
					}
					else
					{
						list.Add(val2);
					}
				}
				stringBuilder.Append(val2.OuterXml.Replace("><", ">" + Environment.NewLine + "<") + Environment.NewLine);
			}
			WinSettings.ExportingData = stringBuilder.ToString();
			int num = 0;
			if (parentNode == null)
			{
				return;
			}
			for (int k = 0; k < parentNode.ChildNodes.Count; k++)
			{
				if (tableInfo.RowsInfo.RowPath != null && (WinSettings.MenuID == "H2" || WinSettings.MenuID == "T8" || WinSettings.MenuID == "T9" || WinSettings.MenuID == "G9" || WinSettings.MenuID == "D1" || WinSettings.MenuID == "H4" || WinSettings.MenuID == "N8" || WinSettings.MenuID == "Q8" || WinSettings.MenuID == "R5" || WinSettings.MenuID == "R7" || WinSettings.MenuID == "U5" || WinSettings.MenuID.Equals("Z5")))
				{
					if (parentNode.ChildNodes[k].Name == tableInfo.RowsInfo.RowPath)
					{
						num = k;
					}
				}
				else if (parentNode.ChildNodes[k].Name == xmlTableNode.Name)
				{
					num = k;
				}
			}
			int num2 = 0;
			for (int j = list.Count - 1; j >= 0; j--)
			{
				num2++;
				parentNode.InsertBefore(list[j], parentNode.ChildNodes[num]);
			}
			if (WinSettings.MenuID == "B2" || WinSettings.MenuID == "O9")
			{
				parentNode.RemoveChild(parentNode.ChildNodes[parentNode.ChildNodes.Count - 1]);
			}
			else if (WinSettings.MenuID == "03" && tableInfo.TableName == "PL_TT03")
			{
				parentNode.RemoveChild(parentNode.ChildNodes[parentNode.ChildNodes.Count - 7]);
			}
			else if ((WinSettings.MenuID == "28" || WinSettings.MenuID == "38" || WinSettings.MenuID == "31" || WinSettings.MenuID == "30" || WinSettings.MenuID == "29") && tableInfo.TableName.Contains("PLuc_01_1_PL_CNKD"))
			{
				parentNode.RemoveChild(parentNode.ChildNodes[num + num2]);
			}
			else if (WinSettings.MenuID == "M6" || WinSettings.MenuID == "O8" || WinSettings.MenuID == "H2" || WinSettings.MenuID == "T8" || WinSettings.MenuID == "T9" || WinSettings.MenuID == "I4" || WinSettings.MenuID == "I7" || WinSettings.MenuID == "D1" || WinSettings.MenuID == "G9" || WinSettings.MenuID == "N8" || WinSettings.MenuID == "B8" || WinSettings.MenuID == "H4" || WinSettings.MenuID == "C7" || WinSettings.MenuID == "C8" || WinSettings.MenuID == "E8" || WinSettings.MenuID == "I8" || "~N2~N3~N7~99~Q4~Q5~R1~Q6~Q2~Q9~Q8~R7~R5~U5~Y7~Z5~".IndexOf("~" + WinSettings.MenuID + "~") > -1)
			{
				parentNode.RemoveChild(parentNode.ChildNodes[num + num2]);
			}
			else
			{
				parentNode.RemoveChild(parentNode.ChildNodes[num2]);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static string RemoveValueSpecial(HTKKCMCTableCellInfo cell, string value)
	{
		if (WinSettings.MenuID == "B2" || WinSettings.MenuID == "O9")
		{
			if (cell.Path == "DChokhau/ct19_ma" && value == "999")
			{
				return "";
			}
			if (cell.Path == "DChokhau/ct20_ma" && value == "99999")
			{
				return "";
			}
			if (cell.Path == "DChokhau/ct21_ma" && value == "9999999")
			{
				return "";
			}
		}
		if (WinSettings.MenuID == "M6" || WinSettings.MenuID == "O8")
		{
			if (cell.Path == "DCThuongTru/ct26_ma" && value == "999")
			{
				return "";
			}
			if (cell.Path == "DCThuongTru/ct27_ma" && value == "99999")
			{
				return "";
			}
			if (cell.Path == "DCThuongTru/ct28_ma" && value == "9999999")
			{
				return "";
			}
			if (cell.Path == "DCHienTai/ct30_ma" && value == "999")
			{
				return "";
			}
			if (cell.Path == "DCHienTai/ct31_ma" && value == "99999")
			{
				return "";
			}
			if (cell.Path == "DCHienTai/ct32_ma" && value == "9999999")
			{
				return "";
			}
			if (cell.Path == "ct18_ma" && value == "999")
			{
				return "";
			}
			if (cell.Path == "ct19_ma" && value == "99999")
			{
				return "";
			}
			if (cell.Path == "ct20_ma" && value == "9999999")
			{
				return "";
			}
		}
		return value;
	}

	public static string GetControlTypeFromSView_01_1_KHBS(SheetView sView, int row, int col, ref string defaultValue)
	{
		try
		{
			NumberCellType numberCellType = new NumberCellType();
			CurrencyCellType currencyCellType = new CurrencyCellType();
			TextCellType textCellType = new TextCellType();
			CheckBoxCellType checkBoxCellType = new CheckBoxCellType();
			MaskCellType maskCellType = new MaskCellType();
			ICellType cellType = sView.Cells[row - 1, col - 1].CellType;
			object value = sView.Cells[row - 1, col - 1].Value;
			if ((object)cellType.GetType() == numberCellType.GetType())
			{
				int decimalPlaces = ((NumberCellType)cellType).DecimalPlaces;
				defaultValue = "0,";
				for (int i = 0; i < decimalPlaces; i++)
				{
					defaultValue += "0";
				}
				if (defaultValue.Length == 2)
				{
					defaultValue = "0";
				}
				return "16";
			}
			if ((object)cellType.GetType() == currencyCellType.GetType())
			{
				int decimalPlaces = ((CurrencyCellType)cellType).DecimalPlaces;
				defaultValue = "0,";
				for (int i = 0; i < decimalPlaces; i++)
				{
					defaultValue += "0";
				}
				if (defaultValue.Length == 2)
				{
					defaultValue = "0";
				}
				return "16";
			}
			if ((object)cellType.GetType() == textCellType.GetType())
			{
				return "0";
			}
			if ((object)cellType.GetType() == maskCellType.GetType())
			{
				return "14";
			}
			if ((object)cellType.GetType() == checkBoxCellType.GetType())
			{
				return "2";
			}
			return "";
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return "";
		}
	}

	public static void InsertExportTableNode(SheetView sView, XmlNode xmlTableNode, XmlNode xmlTableTemplateNode, HTKKCMCTableInfo tableInfo, HTKKCMCTableInfo rootTableInfo, int sheetIndex = -1)
	{
		//IL_001e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0024: Expected O, but got Unknown
		try
		{
			XmlNode val = xmlTableTemplateNode.Clone();
			foreach (XmlAttribute item in (XmlNamedNodeMap)val.Attributes)
			{
				XmlAttribute val2 = item;
				if (((XmlNode)val2).Name == "minOccurs")
				{
					val.Attributes.Remove(val2);
					break;
				}
			}
			if (WinSettings.MenuID == "M4")
			{
				if (xmlTableNode.ParentNode != null && xmlTableNode.ParentNode.Name == "NDTKhai")
				{
					xmlTableNode.ParentNode.InsertBefore(val, xmlTableNode.ParentNode.ChildNodes[xmlTableNode.ParentNode.ChildNodes.Count - 2]);
				}
			}
			else if (tableInfo.TablePath == "PL01_1_KHBS/Muc_A/DSachHSo/BKeHSo" && xmlTableNode.ParentNode.Name == "DSachHSo")
			{
				xmlTableNode.ParentNode.InsertBefore(val, xmlTableNode.ParentNode.ChildNodes[xmlTableNode.ParentNode.ChildNodes.Count - 1]);
			}
			else if (tableInfo.TablePath == "PL01_KHBS/Muc_A/Muc_I/Muc_2/DSachPLuc/BKePLuc" && xmlTableNode.ParentNode.Name == "DSachPLuc")
			{
				xmlTableNode.ParentNode.InsertBefore(val, xmlTableNode.ParentNode.ChildNodes[xmlTableNode.ParentNode.ChildNodes.Count - 1]);
			}
			else if (WinSettings.MenuID == "H2" || WinSettings.MenuID == "H4")
			{
				if (tableInfo.TablePath == "PLuc_01_1_TAIN/TenNhaMay" && WinSettings.MenuID == "H2")
				{
					xmlTableNode.ParentNode.InsertBefore(val, xmlTableNode.ParentNode.ChildNodes[xmlTableNode.ParentNode.ChildNodes.Count - 2]);
				}
				else if (tableInfo.TablePath == "PLuc_01_1_TAIN/TenNhaMay" && WinSettings.MenuID == "H4")
				{
					xmlTableNode.ParentNode.InsertBefore(val, xmlTableNode.ParentNode.ChildNodes[xmlTableNode.ParentNode.ChildNodes.Count - 4]);
				}
				else if (xmlTableNode.ParentNode != null)
				{
					xmlTableNode.ParentNode.AppendChild(val);
				}
			}
			else if (WinSettings.MenuID == "G9")
			{
				if (tableInfo.TablePath == "PLuc_01_2_TBVMT/DN_DVPT/CTiet_DN_DVPT")
				{
					xmlTableNode.ParentNode.InsertBefore(val, xmlTableNode.ParentNode.ChildNodes[xmlTableNode.ParentNode.ChildNodes.Count - 10]);
				}
				else if (tableInfo.TablePath == "PLuc_01_2_TBVMT/TongTheoTinh/CTietTheoTinh")
				{
					xmlTableNode.ParentNode.InsertBefore(val, xmlTableNode.ParentNode.ChildNodes[xmlTableNode.ParentNode.ChildNodes.Count - 3]);
				}
				else if (tableInfo.TablePath == "PLuc_01_1_TBVMT/CoSoKhaiThac/CTietCoSoKhaiThac")
				{
					xmlTableNode.ParentNode.InsertBefore(val, xmlTableNode.ParentNode.ChildNodes[xmlTableNode.ParentNode.ChildNodes.Count - 9]);
				}
				else if (tableInfo.TablePath == "PLuc_01_1_TBVMT/TongTheoTinh/CTietTheoTinh")
				{
					xmlTableNode.ParentNode.InsertBefore(val, xmlTableNode.ParentNode.ChildNodes[xmlTableNode.ParentNode.ChildNodes.Count - 3]);
				}
				else if (xmlTableNode.ParentNode != null)
				{
					xmlTableNode.ParentNode.AppendChild(val);
				}
			}
			else if (WinSettings.MenuID == "D1" || WinSettings.MenuID.Equals("Z5"))
			{
				if (tableInfo.TablePath == "PLuc_01_2_GTGT/TenNhaMay")
				{
					xmlTableNode.ParentNode.InsertBefore(val, xmlTableNode.ParentNode.ChildNodes[xmlTableNode.ParentNode.ChildNodes.Count - 2]);
				}
				else if (xmlTableNode.ParentNode != null)
				{
					xmlTableNode.ParentNode.AppendChild(val);
				}
			}
			else if (tableInfo.TablePath == "PL_01-1_TBH_TB/Tentochuc" && WinSettings.MenuID == "R7")
			{
				xmlTableNode.ParentNode.InsertBefore(val, xmlTableNode.ParentNode.ChildNodes[xmlTableNode.ParentNode.ChildNodes.Count - 1]);
			}
			else if ("~U5~".IndexOf("~" + WinSettings.MenuID + "~") > -1)
			{
				InsertAfterLastNodeByXPath(xmlTableNode.ParentNode, val, xmlTableNode.Name);
			}
			else if (xmlTableNode.ParentNode != null)
			{
				xmlTableNode.ParentNode.AppendChild(val);
			}
			ExportDataTableToTableNode(sView, val, tableInfo, rootTableInfo, sheetIndex);
			if (val.Attributes["id"] != null)
			{
				((XmlNode)val.Attributes["id"]).Value = "ID_" + tableInfo.TableID;
			}
			if (val.Attributes["ID"] != null)
			{
				((XmlNode)val.Attributes["ID"]).Value = "ID_" + tableInfo.TableID;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void ExportDataStaticTableToXMLNode(SheetView sView, XmlNode xmlExportTableNode, HTKKCMCStaticOfTable staticOfTable, int sheetIndex = -1)
	{
		if (xmlExportTableNode == null)
		{
			return;
		}
		foreach (HTKKCMCCellInfo staticCell in staticOfTable.StaticCells)
		{
			if (staticCell.Path == null || staticCell.Path.Length <= 0)
			{
				continue;
			}
			XmlNode val = xmlExportTableNode.SelectSingleNode(staticCell.Path);
			if (val == null)
			{
				continue;
			}
			string text = staticCell.Controltype;
			string defaultValue = staticCell.DefaultValue;
			if ("I7,I4,C7,I8,C8".Contains(WinSettings.MenuID) || (WinSettings.ToKhaiInfo.SheetInfos[sheetIndex] != null && WinSettings.ToKhaiInfo.SheetInfos[sheetIndex].SheetName == "01-KHBS") || (WinSettings.ToKhaiInfo.SheetInfos[sheetIndex] != null && WinSettings.ToKhaiInfo.SheetInfos[sheetIndex].SheetName == "01-1/KHBS"))
			{
				string controlTypeFromSView_01_1_KHBS = GetControlTypeFromSView_01_1_KHBS(sView, staticCell.CellTemplate.RowIndex, staticCell.CellTemplate.ColIndex, ref defaultValue);
				if (!string.IsNullOrEmpty(controlTypeFromSView_01_1_KHBS))
				{
					text = controlTypeFromSView_01_1_KHBS;
				}
			}
			string text2 = GeneralFunctions.GetValueExportXML(sView, text, staticCell.CellTemplate.RowIndex, staticCell.CellTemplate.ColIndex, defaultValue, WinSettings.ToKhaiInfo.SheetInfos[sheetIndex].SheetName).Trim();
			if (text == "26" && text2.Trim().Length == 13)
			{
				text2 = text2.Substring(0, 10) + "-" + text2.Substring(10, 3);
			}
			if (!string.IsNullOrEmpty(text2))
			{
				val.InnerText = text2;
			}
			string text3 = text;
			if (text3 == null || !(text3 == "2"))
			{
			}
			if (!string.IsNullOrEmpty(val.InnerXml.Trim()))
			{
				continue;
			}
			switch (text)
			{
			default:
				if (!(text == "94"))
				{
					continue;
				}
				break;
			case "14":
			case "26":
			case "33":
			case "96":
				break;
			}
			XmlNode val2 = val.OwnerDocument.CreateNode((XmlNodeType)2, "xsi:nil", "http://www.w3.org/2001/XMLSchema-instance");
			val2.Value = "true";
			((XmlNamedNodeMap)val.Attributes).SetNamedItem(val2);
		}
	}

	public static void WinDoWorkKiemTra(FpSpread winFpSpread)
	{
		if (File.Exists("ProgressIndicatorTest.exe"))
		{
			Process.Start("ProgressIndicatorTest.exe");
		}
		GC.Collect();
		GC.WaitForPendingFinalizers();
		winFpSpread.SuspendLayout();
		HTKKFile.LogRun("Validate", "Start validate data");
		int num = 0;
		int num2 = 1;
		int num3 = -1;
		int row = -1;
		int column = -1;
		int num4 = 0;
		int num5 = 0;
		int[] array = GeneralFunctions.ArrayCellId("BA_1");
		int column2 = HTKKFile.ColumnLettersToIndex("BA");
		HTKKFile.deleteLogErrorBigData();
		object obj = null;
		Dictionary<int, int> levelWarning = new Dictionary<int, int>();
		Dictionary<object, object> dictionary = new Dictionary<object, object>();
		bool flag = true;
		try
		{
			WinSpreadProcess.reloadBc26AcPl310(winFpSpread);
			if (WinSettings.MenuID == "A4")
			{
				Cell cell = null;
				if (WinSettings.TTPeriodInfo.dicVar.ContainsKey("CELL_CURRENT"))
				{
					cell = (Cell)WinSettings.TTPeriodInfo.dicVar["CELL_CURRENT"];
				}
				if (cell != null)
				{
					CellChange(winFpSpread, winFpSpread.ActiveSheetIndex, cell.Row.Index + 1, cell.Column.Index + 1);
				}
			}
			for (int i = 0; i < winFpSpread.TabStrip.Count; i++)
			{
				if (i < WinSettings.ToKhaiInfo.SheetInfos.Count && WinSettings.ToKhaiInfo.SheetInfos[i] != null)
				{
					WinSettings.ToKhaiInfo.SheetInfos[i].isValidData = true;
				}
				winFpSpread.TabStrip[i].BackColor = Color.White;
			}
			SheetView sheetView = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1];
			for (int i = 11; i < sheetView.NonEmptyRowCount; i++)
			{
				try
				{
					if (!(sheetView.Cells[array[1], array[0]].Text.Trim() == "1"))
					{
						continue;
					}
					string formula = sheetView.Cells[i, num2 - 1].Formula;
					string text = sheetView.Cells[i, num2].Text;
					string formula2 = sheetView.Cells[i, num2 + 1].Formula;
					string text2 = sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("S") - 1].Text;
					if (string.IsNullOrEmpty(formula) && string.IsNullOrEmpty(text) && string.IsNullOrEmpty(formula2) && string.IsNullOrEmpty(text2))
					{
						break;
					}
					if (HTKKFormater.isNullString(formula2))
					{
						continue;
					}
					Color backColor = sheetView.Cells[i, num2 + 1].BackColor;
					string text3 = ((!string.IsNullOrEmpty(formula)) ? formula : ((!formula2.Contains("&")) ? formula2 : ""));
					int num6 = GeneralFunctions.DetectIndexOfSheet(WinSettings.ToKhaiInfo.SheetInfos, text3);
					if (num6 < 0 || WinSettings.ToKhaiInfo.SheetInfos[num6] == null)
					{
						continue;
					}
					string[] array2 = GeneralFunctions.GetCellValild(text3).Split(new char[1] { '#' });
					if (array2.Length != 3)
					{
						continue;
					}
					int row2 = Convert.ToInt32(array2[1]);
					int num7 = GeneralFunctions.DetectIndexOfTable(WinSettings.ToKhaiInfo.SheetInfos[num6], row2);
					int num10;
					string[] array3;
					if (num7 >= 0)
					{
						HTKKCMCSheetInfo hTKKCMCSheetInfo = WinSettings.ToKhaiInfo.SheetInfos[num6];
						string tablePath = hTKKCMCSheetInfo.Tables[num7].TablePath;
						List<HTKKCMCTableInfo> list = hTKKCMCSheetInfo.Tables.Where((HTKKCMCTableInfo x) => x.TablePath == tablePath).ToList();
						foreach (HTKKCMCTableInfo item in list)
						{
							foreach (HTKKCMCCellInfo staticCell in item.Header.StaticCells)
							{
								HTKKFile.RemoveValidate(winFpSpread.Sheets[num6], staticCell.CellTemplate.RowIndex - 1, staticCell.CellTemplate.ColIndex - 1);
							}
							int dataLocationTemplateRowStart = item.Location.DataLocationTemplateRowStart;
							for (int num8 = 0; num8 < item.RowsValues.Count; num8++)
							{
								array3 = formula2.Split(new char[1] { '&' });
								foreach (string cellFomular in array3)
								{
									string[] array4 = GeneralFunctions.GetCellValild(cellFomular).Split(new char[1] { '#' });
									if (array4.Length == 3)
									{
										num10 = HTKKFile.ColumnLettersToIndex(array4[2]);
										if (winFpSpread.Sheets[num6].Cells[dataLocationTemplateRowStart + num8 - 1, num10 - 1].NoteStyle != NoteStyle.Hidden)
										{
											HTKKFile.RemoveValidate(winFpSpread.Sheets[num6], dataLocationTemplateRowStart + num8 - 1, num10 - 1);
											winFpSpread.Sheets[num6].Cells[dataLocationTemplateRowStart + num8 - 1, num10 - 1].BackColor = backColor;
										}
									}
								}
							}
						}
						continue;
					}
					num10 = HTKKFile.ColumnLettersToIndex(array2[2]);
					int num11 = int.Parse(array2[1]);
					if (winFpSpread.Sheets[num6].Cells[num11 - 1, num10 - 1].NoteStyle != NoteStyle.Hidden)
					{
						HTKKFile.RemoveValidate(winFpSpread.Sheets[num6], num11 - 1, num10 - 1);
						winFpSpread.Sheets[num6].Cells[num11 - 1, num10 - 1].BackColor = backColor;
					}
					array3 = formula2.Split(new char[1] { '&' });
					foreach (string cellFomular in array3)
					{
						string[] array5 = GeneralFunctions.GetCellValild(cellFomular).Split(new char[1] { '#' });
						if (array5.Length == 3)
						{
							int num12 = HTKKFile.ColumnLettersToIndex(array5[2]);
							int num13 = int.Parse(array5[1]);
							if (winFpSpread.Sheets[num6].Cells[num13 - 1, num12 - 1].NoteStyle != NoteStyle.Hidden)
							{
								HTKKFile.RemoveValidate(winFpSpread.Sheets[num6], num13 - 1, num12 - 1);
								winFpSpread.Sheets[num6].Cells[num13 - 1, num12 - 1].BackColor = backColor;
							}
						}
					}
				}
				catch (Exception ex)
				{
					HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
				}
			}
			if (WinSettings.ToKhaiInfo.TaiDuLieuRatLon != "1")
			{
				for (int i = 11; i < sheetView.NonEmptyRowCount; i++)
				{
					string formula = sheetView.Cells[i, num2 - 1].Formula;
					string text = sheetView.Cells[i, num2].Text;
					string formula2 = sheetView.Cells[i, num2 + 1].Formula;
					string text2 = sheetView.Cells[i, HTKKFile.ColumnLettersToIndex("S") - 1].Text;
					Color backColor = sheetView.Cells[i, num2 + 1].BackColor;
					if (i == 45)
					{
						num5 = i;
					}
					if (sheetView.Cells[array[1], array[0]].Text.Trim() == "1")
					{
						if (string.IsNullOrEmpty(formula) && string.IsNullOrEmpty(text) && string.IsNullOrEmpty(formula2) && string.IsNullOrEmpty(text2))
						{
							break;
						}
						if (HTKKFormater.isNullString(formula2))
						{
							continue;
						}
						string text3 = ((!string.IsNullOrEmpty(formula)) ? formula : ((!formula2.Contains("&")) ? formula2 : ""));
						string[] array2 = GeneralFunctions.GetCellValild(text3).Split(new char[1] { '#' });
						int num7 = -1;
						object value = 0;
						if (array2.Length != 3)
						{
							continue;
						}
						int row2 = Convert.ToInt32(array2[1]);
						int index = winFpSpread.Sheets.Count - 1;
						string formula3 = sheetView.Cells[i, num2].Formula;
						int num6 = GeneralFunctions.DetectIndexOfSheet(WinSettings.ToKhaiInfo.SheetInfos, text3);
						if (num6 < 0 || WinSettings.ToKhaiInfo.SheetInfos[num6] == null)
						{
							continue;
						}
						num7 = GeneralFunctions.DetectIndexOfTable(WinSettings.ToKhaiInfo.SheetInfos[num6], row2);
						if (num7 >= 0)
						{
							HTKKCMCSheetInfo hTKKCMCSheetInfo = WinSettings.ToKhaiInfo.SheetInfos[num6];
							string tablePath2 = hTKKCMCSheetInfo.Tables[num7].TablePath;
							List<HTKKCMCTableInfo> list = hTKKCMCSheetInfo.Tables.Where((HTKKCMCTableInfo x) => x.TablePath == tablePath2).ToList();
							foreach (HTKKCMCTableInfo item2 in list)
							{
								int dataLocationTemplateRowStart = item2.Location.DataLocationTemplateRowStart;
								for (int num8 = 0; num8 < item2.RowsValues.Count; num8++)
								{
									try
									{
										value = 0;
										string text4 = text;
										if (!string.IsNullOrEmpty(formula3))
										{
											string formula4 = GeneralFunctions.FormulaToTemplate(formula3, row2);
											string formula5 = GeneralFunctions.SetFormulaNextRow(formula4, dataLocationTemplateRowStart + num8);
											winFpSpread.Sheets[index].Cells[i, column2].Formula = formula5;
											text4 = winFpSpread.Sheets[index].Cells[i, column2].Text;
										}
										if (text4 == "1" || string.IsNullOrEmpty(text4) || (string.IsNullOrEmpty(formula3) && text == "0"))
										{
											continue;
										}
										string formula6 = GeneralFunctions.FormulaToTemplate(formula2, row2);
										string text5 = GeneralFunctions.SetFormulaNextRow(formula6, dataLocationTemplateRowStart + num8);
										string[] array3 = text5.Split(new char[1] { '&' });
										foreach (string cellFomular in array3)
										{
											string[] array6 = GeneralFunctions.GetCellValild(cellFomular).Split(new char[1] { '#' });
											string text6;
											if (text4 == "0")
											{
												text6 = sheetView.Cells[i, num2 + 3].Text;
												if (flag)
												{
													flag = false;
												}
												for (int num14 = 0; num14 < winFpSpread.Sheets.Count - 1; num14++)
												{
													num4 = 0;
													if (!(winFpSpread.Sheets[num14].SheetName == array6[0]) || !winFpSpread.Sheets[num14].Visible)
													{
														continue;
													}
													if (num3 == -1)
													{
														num3 = num14;
														column = HTKKFile.ColumnLettersToIndex(array6[2]) - 1;
														row = int.Parse(array6[1]) - 1;
													}
													if (HTKKFormater.isNullString(text2) || text2.Trim() == "Y")
													{
														WinSettings.ToKhaiInfo.SheetInfos[num14].isValidData = false;
														if (num4 == 0 || num4 == 1)
														{
															num4 = 2;
														}
														HTKKFile.SetNodeBackGroundPink(winFpSpread.Sheets[num14], int.Parse(array6[1]) - 1, HTKKFile.ColumnLettersToIndex(array6[2]) - 1, text6, int.Parse(array6[1]));
														dictionary.TryGetValue(cellFomular, out value);
														if (ConvertHelpers.ConvertToInt(value) < 2)
														{
															HTKKFile.AddOrUpdate(dictionary, cellFomular, 2);
														}
													}
													else
													{
														if (num4 == 0)
														{
															num4 = 1;
														}
														dictionary.TryGetValue(cellFomular, out value);
														if (ConvertHelpers.ConvertToInt(value) < 1)
														{
															HTKKFile.AddOrUpdate(dictionary, cellFomular, 1);
															HTKKFile.SetNodeBackGroundYellow(winFpSpread.Sheets[num14], int.Parse(array6[1]) - 1, HTKKFile.ColumnLettersToIndex(array6[2]) - 1, text6);
														}
													}
													switch (num4)
													{
													case 2:
														winFpSpread.TabStrip[num14].BackColor = Color.Pink;
														break;
													case 1:
														winFpSpread.TabStrip[num14].BackColor = Color.LemonChiffon;
														break;
													}
													if (levelWarning.Count > 0)
													{
														bool flag2 = false;
														foreach (KeyValuePair<int, int> item3 in levelWarning)
														{
															int key = item3.Key;
															int value2 = item3.Value;
															if (key == num14)
															{
																flag2 = true;
																if (value2 < num4)
																{
																	levelWarning.Remove(key);
																	levelWarning.Add(key, num4);
																	break;
																}
															}
														}
														if (!flag2)
														{
															levelWarning.Add(num14, num4);
														}
													}
													else
													{
														levelWarning.Add(num14, num4);
													}
													break;
												}
											}
											string text7;
											if (text == "2")
											{
												text7 = winFpSpread.Sheets[num6].Cells[int.Parse(array6[1]) - 1, HTKKFile.ColumnLettersToIndex(array6[2]) - 1].Text.Trim();
												if (!string.IsNullOrEmpty(text7) && !ValidateBase.CheckMaSoThue(text7))
												{
													text6 = sheetView.Cells[i, num2 + 3].Text;
													if (flag)
													{
														flag = false;
													}
													for (int num14 = 0; num14 < winFpSpread.Sheets.Count - 1; num14++)
													{
														num4 = 0;
														if (!(winFpSpread.Sheets[num14].SheetName == array6[0]) || !winFpSpread.Sheets[num14].Visible)
														{
															continue;
														}
														if (num3 == -1)
														{
															num3 = num14;
															column = HTKKFile.ColumnLettersToIndex(array6[2]) - 1;
															row = int.Parse(array6[1]) - 1;
														}
														if (HTKKFormater.isNullString(text2) || text2.Trim() == "Y")
														{
															WinSettings.ToKhaiInfo.SheetInfos[num14].isValidData = false;
															if (num4 == 0 || num4 == 1)
															{
																num4 = 2;
															}
															HTKKFile.SetNodeBackGroundPink(winFpSpread.Sheets[num14], int.Parse(array6[1]) - 1, HTKKFile.ColumnLettersToIndex(array6[2]) - 1, text6, int.Parse(array6[1]));
															dictionary.TryGetValue(cellFomular, out value);
															if (ConvertHelpers.ConvertToInt(value) < 2)
															{
																HTKKFile.AddOrUpdate(dictionary, cellFomular, 2);
															}
														}
														else
														{
															if (num4 == 0)
															{
																num4 = 1;
															}
															dictionary.TryGetValue(cellFomular, out value);
															if (ConvertHelpers.ConvertToInt(value) < 1)
															{
																HTKKFile.AddOrUpdate(dictionary, cellFomular, 1);
																HTKKFile.SetNodeBackGroundYellow(winFpSpread.Sheets[num14], int.Parse(array6[1]) - 1, HTKKFile.ColumnLettersToIndex(array6[2]) - 1, text6);
															}
														}
														switch (num4)
														{
														case 2:
															winFpSpread.TabStrip[num14].BackColor = Color.Pink;
															break;
														case 1:
															winFpSpread.TabStrip[num14].BackColor = Color.LemonChiffon;
															break;
														}
														if (levelWarning.Count > 0)
														{
															bool flag2 = false;
															foreach (KeyValuePair<int, int> item4 in levelWarning)
															{
																int key = item4.Key;
																int value2 = item4.Value;
																if (key == num14)
																{
																	flag2 = true;
																	if (value2 < num4)
																	{
																		levelWarning.Remove(key);
																		levelWarning.Add(key, num4);
																		break;
																	}
																}
															}
															if (!flag2)
															{
																levelWarning.Add(num14, num4);
															}
														}
														else
														{
															levelWarning.Add(num14, num4);
														}
														break;
													}
												}
											}
											if (!(text == "3"))
											{
												continue;
											}
											text7 = winFpSpread.Sheets[num6].Cells[int.Parse(array6[1]) - 1, HTKKFile.ColumnLettersToIndex(array6[2]) - 1].Text.Trim();
											if (string.IsNullOrEmpty(text7) || ValidateBase.CheckMaSoThueCCCD(text7))
											{
												continue;
											}
											text6 = sheetView.Cells[i, num2 + 3].Text;
											if (flag)
											{
												flag = false;
											}
											for (int num14 = 0; num14 < winFpSpread.Sheets.Count - 1; num14++)
											{
												num4 = 0;
												if (!(winFpSpread.Sheets[num14].SheetName == array6[0]) || !winFpSpread.Sheets[num14].Visible)
												{
													continue;
												}
												if (num3 == -1)
												{
													num3 = num14;
													column = HTKKFile.ColumnLettersToIndex(array6[2]) - 1;
													row = int.Parse(array6[1]) - 1;
												}
												if (HTKKFormater.isNullString(text2) || text2.Trim() == "Y")
												{
													WinSettings.ToKhaiInfo.SheetInfos[num14].isValidData = false;
													if (num4 == 0 || num4 == 1)
													{
														num4 = 2;
													}
													HTKKFile.SetNodeBackGroundPink(winFpSpread.Sheets[num14], int.Parse(array6[1]) - 1, HTKKFile.ColumnLettersToIndex(array6[2]) - 1, text6, int.Parse(array6[1]));
													dictionary.TryGetValue(cellFomular, out value);
													if (ConvertHelpers.ConvertToInt(value) < 2)
													{
														HTKKFile.AddOrUpdate(dictionary, cellFomular, 2);
													}
												}
												else
												{
													if (num4 == 0)
													{
														num4 = 1;
													}
													dictionary.TryGetValue(cellFomular, out value);
													if (ConvertHelpers.ConvertToInt(value) < 1)
													{
														HTKKFile.AddOrUpdate(dictionary, cellFomular, 1);
														HTKKFile.SetNodeBackGroundYellow(winFpSpread.Sheets[num14], int.Parse(array6[1]) - 1, HTKKFile.ColumnLettersToIndex(array6[2]) - 1, text6);
													}
												}
												switch (num4)
												{
												case 2:
													winFpSpread.TabStrip[num14].BackColor = Color.Pink;
													break;
												case 1:
													winFpSpread.TabStrip[num14].BackColor = Color.LemonChiffon;
													break;
												}
												if (levelWarning.Count > 0)
												{
													bool flag2 = false;
													foreach (KeyValuePair<int, int> item5 in levelWarning)
													{
														int key = item5.Key;
														int value2 = item5.Value;
														if (key == num14)
														{
															flag2 = true;
															if (value2 < num4)
															{
																levelWarning.Remove(key);
																levelWarning.Add(key, num4);
																break;
															}
														}
													}
													if (!flag2)
													{
														levelWarning.Add(num14, num4);
													}
												}
												else
												{
													levelWarning.Add(num14, num4);
												}
												break;
											}
										}
									}
									catch (Exception ex)
									{
										HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
									}
								}
							}
							continue;
						}
						switch (text)
						{
						case "0":
						{
							string text6 = sheetView.Cells[i, num2 + 3].Text;
							if (flag)
							{
								flag = false;
							}
							for (int num14 = 0; num14 < winFpSpread.Sheets.Count - 1; num14++)
							{
								num4 = 0;
								if (!(winFpSpread.Sheets[num14].SheetName == array2[0]) || !winFpSpread.Sheets[num14].Visible)
								{
									continue;
								}
								if (num3 == -1)
								{
									num3 = num14;
									column = HTKKFile.ColumnLettersToIndex(array2[2]) - 1;
									row = int.Parse(array2[1]) - 1;
								}
								string[] array3 = formula2.Split(new char[1] { '&' });
								foreach (string cellFomular in array3)
								{
									string[] array6 = GeneralFunctions.GetCellValild(cellFomular).Split(new char[1] { '#' });
									if (HTKKFormater.isNullString(text2) || text2.Trim() == "Y")
									{
										WinSettings.ToKhaiInfo.SheetInfos[num14].isValidData = false;
										if (num4 == 0 || num4 == 1)
										{
											num4 = 2;
										}
										HTKKFile.SetNodeBackGroundPink(winFpSpread.Sheets[num14], int.Parse(array6[1]) - 1, HTKKFile.ColumnLettersToIndex(array6[2]) - 1, text6);
										dictionary.TryGetValue(cellFomular, out value);
										if (ConvertHelpers.ConvertToInt(value) < 2)
										{
											HTKKFile.AddOrUpdate(dictionary, cellFomular, 2);
										}
									}
									else
									{
										if (num4 == 0)
										{
											num4 = 1;
										}
										dictionary.TryGetValue(cellFomular, out value);
										if (ConvertHelpers.ConvertToInt(value) < 1)
										{
											HTKKFile.AddOrUpdate(dictionary, cellFomular, 1);
											HTKKFile.SetNodeBackGroundYellow(winFpSpread.Sheets[num14], int.Parse(array6[1]) - 1, HTKKFile.ColumnLettersToIndex(array6[2]) - 1, text6);
										}
									}
								}
								switch (num4)
								{
								case 2:
									winFpSpread.TabStrip[num14].BackColor = Color.Pink;
									break;
								case 1:
									winFpSpread.TabStrip[num14].BackColor = Color.LemonChiffon;
									break;
								}
								if (levelWarning.Count > 0)
								{
									bool flag2 = false;
									foreach (KeyValuePair<int, int> item6 in levelWarning)
									{
										int key = item6.Key;
										int value2 = item6.Value;
										if (key == num14)
										{
											flag2 = true;
											if (value2 < num4)
											{
												levelWarning.Remove(key);
												levelWarning.Add(key, num4);
												break;
											}
										}
									}
									if (!flag2)
									{
										levelWarning.Add(num14, num4);
									}
								}
								else
								{
									levelWarning.Add(num14, num4);
								}
								break;
							}
							break;
						}
						case "2":
						{
							string text7 = winFpSpread.Sheets[num6].Cells[int.Parse(array2[1]) - 1, HTKKFile.ColumnLettersToIndex(array2[2]) - 1].Text.Trim();
							if (string.IsNullOrEmpty(text7) || ValidateBase.CheckMaSoThue(text7))
							{
								break;
							}
							string text6 = sheetView.Cells[i, num2 + 3].Text;
							if (flag)
							{
								flag = false;
							}
							for (int num14 = 0; num14 < winFpSpread.Sheets.Count - 1; num14++)
							{
								num4 = 0;
								if (!(winFpSpread.Sheets[num14].SheetName == array2[0]) || !winFpSpread.Sheets[num14].Visible)
								{
									continue;
								}
								if (num3 == -1)
								{
									num3 = num14;
									column = HTKKFile.ColumnLettersToIndex(array2[2]) - 1;
									row = int.Parse(array2[1]) - 1;
								}
								if (HTKKFormater.isNullString(text2) || text2.Trim() == "Y")
								{
									WinSettings.ToKhaiInfo.SheetInfos[num14].isValidData = false;
									if (num4 == 0 || num4 == 1)
									{
										num4 = 2;
									}
									HTKKFile.SetNodeBackGroundPink(winFpSpread.Sheets[num14], int.Parse(array2[1]) - 1, HTKKFile.ColumnLettersToIndex(array2[2]) - 1, text6);
									dictionary.TryGetValue(formula2, out value);
									if (ConvertHelpers.ConvertToInt(value) < 2)
									{
										HTKKFile.AddOrUpdate(dictionary, formula2, 2);
									}
								}
								else
								{
									if (num4 == 0)
									{
										num4 = 1;
									}
									dictionary.TryGetValue(formula2, out value);
									if (ConvertHelpers.ConvertToInt(value) < 1)
									{
										HTKKFile.AddOrUpdate(dictionary, formula2, 1);
										HTKKFile.SetNodeBackGroundYellow(winFpSpread.Sheets[num14], int.Parse(array2[1]) - 1, HTKKFile.ColumnLettersToIndex(array2[2]) - 1, text6);
									}
								}
								switch (num4)
								{
								case 2:
									winFpSpread.TabStrip[num14].BackColor = Color.Pink;
									break;
								case 1:
									winFpSpread.TabStrip[num14].BackColor = Color.LemonChiffon;
									break;
								}
								if (levelWarning.Count > 0)
								{
									bool flag2 = false;
									foreach (KeyValuePair<int, int> item7 in levelWarning)
									{
										int key = item7.Key;
										int value2 = item7.Value;
										if (key == num14)
										{
											flag2 = true;
											if (value2 < num4)
											{
												levelWarning.Remove(key);
												levelWarning.Add(key, num4);
												break;
											}
										}
									}
									if (!flag2)
									{
										levelWarning.Add(num14, num4);
									}
								}
								else
								{
									levelWarning.Add(num14, num4);
								}
								break;
							}
							break;
						}
						case "3":
						{
							string text7 = winFpSpread.Sheets[num6].Cells[int.Parse(array2[1]) - 1, HTKKFile.ColumnLettersToIndex(array2[2]) - 1].Text.Trim();
							if (string.IsNullOrEmpty(text7) || ValidateBase.CheckMaSoThueCCCD(text7))
							{
								break;
							}
							string text6 = sheetView.Cells[i, num2 + 3].Text;
							if (flag)
							{
								flag = false;
							}
							for (int num14 = 0; num14 < winFpSpread.Sheets.Count - 1; num14++)
							{
								num4 = 0;
								if (!(winFpSpread.Sheets[num14].SheetName == array2[0]) || !winFpSpread.Sheets[num14].Visible)
								{
									continue;
								}
								if (num3 == -1)
								{
									num3 = num14;
									column = HTKKFile.ColumnLettersToIndex(array2[2]) - 1;
									row = int.Parse(array2[1]) - 1;
								}
								if (HTKKFormater.isNullString(text2) || text2.Trim() == "Y")
								{
									WinSettings.ToKhaiInfo.SheetInfos[num14].isValidData = false;
									if (num4 == 0 || num4 == 1)
									{
										num4 = 2;
									}
									HTKKFile.SetNodeBackGroundPink(winFpSpread.Sheets[num14], int.Parse(array2[1]) - 1, HTKKFile.ColumnLettersToIndex(array2[2]) - 1, text6);
									dictionary.TryGetValue(formula2, out value);
									if (ConvertHelpers.ConvertToInt(value) < 2)
									{
										HTKKFile.AddOrUpdate(dictionary, formula2, 2);
									}
								}
								else
								{
									if (num4 == 0)
									{
										num4 = 1;
									}
									dictionary.TryGetValue(formula2, out value);
									if (ConvertHelpers.ConvertToInt(value) < 1)
									{
										HTKKFile.AddOrUpdate(dictionary, formula2, 1);
										HTKKFile.SetNodeBackGroundYellow(winFpSpread.Sheets[num14], int.Parse(array2[1]) - 1, HTKKFile.ColumnLettersToIndex(array2[2]) - 1, text6);
									}
								}
								switch (num4)
								{
								case 2:
									winFpSpread.TabStrip[num14].BackColor = Color.Pink;
									break;
								case 1:
									winFpSpread.TabStrip[num14].BackColor = Color.LemonChiffon;
									break;
								}
								if (levelWarning.Count > 0)
								{
									bool flag2 = false;
									foreach (KeyValuePair<int, int> item8 in levelWarning)
									{
										int key = item8.Key;
										int value2 = item8.Value;
										if (key == num14)
										{
											flag2 = true;
											if (value2 < num4)
											{
												levelWarning.Remove(key);
												levelWarning.Add(key, num4);
												break;
											}
										}
									}
									if (!flag2)
									{
										levelWarning.Add(num14, num4);
									}
								}
								else
								{
									levelWarning.Add(num14, num4);
								}
								break;
							}
							break;
						}
						}
					}
					else
					{
						if (HTKKFormater.isNullString(formula2))
						{
							continue;
						}
						string[] array7 = GeneralFunctions.GetCellValild(formula2).Split(new char[1] { '#' });
						object value = 0;
						if (text == "0")
						{
							string text6 = sheetView.Cells[i, num2 + 3].Text;
							if (flag)
							{
								flag = false;
							}
							for (int num14 = 0; num14 < winFpSpread.Sheets.Count - 1; num14++)
							{
								num4 = 0;
								if (!(winFpSpread.Sheets[num14].SheetName == array7[0]) || !winFpSpread.Sheets[num14].Visible)
								{
									continue;
								}
								if (num3 == -1)
								{
									num3 = num14;
									column = HTKKFile.ColumnLettersToIndex(array7[2]) - 1;
									row = int.Parse(array7[1]) - 1;
								}
								if (HTKKFormater.isNullString(text2) || text2.Trim() == "Y")
								{
									WinSettings.ToKhaiInfo.SheetInfos[num14].isValidData = false;
									if (num4 == 0 || num4 == 1)
									{
										num4 = 2;
									}
									HTKKFile.SetNodeBackGroundPink(winFpSpread.Sheets[num14], int.Parse(array7[1]) - 1, HTKKFile.ColumnLettersToIndex(array7[2]) - 1, text6);
									dictionary.TryGetValue(formula2, out value);
									if (ConvertHelpers.ConvertToInt(value) < 2)
									{
										HTKKFile.AddOrUpdate(dictionary, formula2, 2);
									}
								}
								else
								{
									if (num4 == 0)
									{
										num4 = 1;
									}
									dictionary.TryGetValue(formula2, out value);
									if (ConvertHelpers.ConvertToInt(value) < 1)
									{
										HTKKFile.AddOrUpdate(dictionary, formula2, 1);
										HTKKFile.SetNodeBackGroundYellow(winFpSpread.Sheets[num14], int.Parse(array7[1]) - 1, HTKKFile.ColumnLettersToIndex(array7[2]) - 1, text6);
									}
								}
								switch (num4)
								{
								case 2:
									winFpSpread.TabStrip[num14].BackColor = Color.Pink;
									break;
								case 1:
									winFpSpread.TabStrip[num14].BackColor = Color.LemonChiffon;
									break;
								}
								if (levelWarning.Count > 0)
								{
									bool flag2 = false;
									foreach (KeyValuePair<int, int> item9 in levelWarning)
									{
										int key = item9.Key;
										int value2 = item9.Value;
										if (key == num14)
										{
											flag2 = true;
											if (value2 < num4)
											{
												levelWarning.Remove(key);
												levelWarning.Add(key, num4);
												break;
											}
										}
									}
									if (!flag2)
									{
										levelWarning.Add(num14, num4);
									}
								}
								else
								{
									levelWarning.Add(num14, num4);
								}
								break;
							}
							continue;
						}
						for (int num14 = 0; num14 < winFpSpread.Sheets.Count - 1; num14++)
						{
							if (winFpSpread.Sheets[num14].SheetName == array7[0] && winFpSpread.Sheets[num14].Visible)
							{
								dictionary.TryGetValue(formula2, out value);
								if (ConvertHelpers.ConvertToInt(value) == 0)
								{
									HTKKFile.RemoveValidate(winFpSpread.Sheets[num14], int.Parse(array7[1]) - 1, HTKKFile.ColumnLettersToIndex(array7[2]) - 1);
									winFpSpread.Sheets[num14].Cells[int.Parse(array7[1]) - 1, HTKKFile.ColumnLettersToIndex(array7[2]) - 1].BackColor = backColor;
									break;
								}
							}
						}
					}
				}
			}
			HTKKFile.LogRun("Validate", "Start update value");
			if (WinSettings.ToKhaiInfo.TaiDuLieuRatLon != "1")
			{
				UpdateObjectValue(winFpSpread, levelWarning);
			}
			HTKKFile.LogRun("Validate", "End update value");
			if (flag)
			{
				for (int num14 = 0; num14 < winFpSpread.Sheets.Count - 1; num14++)
				{
					winFpSpread.TabStrip[num14].BackColor = Color.White;
				}
			}
			for (int i = 0; i < winFpSpread.TabStrip.Count; i++)
			{
				num4 = 0;
				levelWarning.TryGetValue(i, out num4);
				switch (num4)
				{
				case 2:
					winFpSpread.TabStrip[i].BackColor = Color.Pink;
					break;
				case 1:
					winFpSpread.TabStrip[i].BackColor = Color.LemonChiffon;
					break;
				default:
					winFpSpread.TabStrip[i].BackColor = Color.White;
					break;
				}
			}
			if ("A1,H1,A3,H3,A8,H8,A2,A5,A6,A7,A9,H9,53,54,58,78,69,19,20,22,B3,B4,B5,C1,C2,C3,C4,C5,P8,P9,Q1,Q2,Q3,U3,U4,X4,X3,X2,X1,U9".Contains(WinSettings.MenuID) && winFpSpread.Sheets.Count > 4 && winFpSpread.Sheets[2].Visible && winFpSpread.Sheets[3].Visible)
			{
				if (levelWarning.ContainsKey(2))
				{
					levelWarning.Remove(2);
					levelWarning.Add(2, 2);
				}
				if (levelWarning.ContainsKey(3))
				{
					levelWarning.Remove(3);
					levelWarning.Add(3, 2);
				}
			}
			MethodInfo[] methods = GetMethods();
			bool flag3 = false;
			for (int i = 0; i < methods.Length; i++)
			{
				if (methods[i].Name.IndexOf("isValidDuLieuTK") == -1)
				{
					continue;
				}
				ParameterInfo[] parameters = methods[i].GetParameters();
				object[] array8 = new object[parameters.Length];
				for (int num14 = 0; num14 < parameters.Length; num14++)
				{
					switch (parameters[num14].Name)
					{
					case "winFpSpread":
						array8[num14] = winFpSpread;
						break;
					case "ToKhaiInfo":
						array8[num14] = WinSettings.ToKhaiInfo;
						break;
					case "TTinChung":
						array8[num14] = WinSettings.TTinChung;
						break;
					case "PeriodInfo":
						array8[num14] = WinSettings.TTPeriodInfo;
						break;
					case "HeaderInfo":
						array8[num14] = WinSettings.HTKKHeaderInfo;
						break;
					case "levelWarning":
						array8[num14] = levelWarning;
						break;
					case "output":
						array8[num14] = obj;
						break;
					}
					flag3 = true;
				}
				methods[i].Invoke(WinSettings.InstanceClass, array8);
				obj = array8[^1];
				if (WinSettings.MenuID == "57")
				{
					object key2 = "CHECK_CHON";
					HTKKFile.AddOrUpdate(WinSettings.ToKhaiInfo.dicObject, key2, obj);
				}
			}
			if (WinSettings.ToKhaiInfo.TaiDuLieuRatLon != "1")
			{
				GeneralFunctions.CheckComboboxEmptyCode(winFpSpread, WinSettings.ToKhaiInfo, ref levelWarning);
			}
			HTKKFile.LogRun("Validate", "End validate data");
			winFpSpread.ResumeLayout();
			firstSheetErr = -1;
			if (num3 != -1)
			{
				winFpSpread.ActiveSheetIndex = num3;
				winFpSpread.ActiveSheet.SetActiveCell(row, column);
				firstSheetErr = num3;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		finally
		{
			Process[] processesByName = Process.GetProcessesByName("ProgressIndicatorTest");
			foreach (Process process in processesByName)
			{
				process.Kill();
			}
		}
	}

	public static void UpdateObjectValue(FpSpread winFpSpread, Dictionary<int, int> levelWarning)
	{
		try
		{
			for (int i = 0; i < WinSettings.ToKhaiInfo.SheetInfos.Count; i++)
			{
				HTKKCMCSheetInfo hTKKCMCSheetInfo = WinSettings.ToKhaiInfo.SheetInfos[i];
				if (hTKKCMCSheetInfo == null)
				{
					continue;
				}
				WinSettings.ToKhaiInfo.SheetInfos[i].CoDuLieu = false;
				List<HTKKCMCCellInfo> staticCells = hTKKCMCSheetInfo.StaticCells;
				for (int j = 0; j < staticCells.Count; j++)
				{
					HTKKCMCCellInfo hTKKCMCCellInfo = staticCells[j];
					if (hTKKCMCCellInfo.CellTemplate.RowIndex > winFpSpread.Sheets[i].NonEmptyRowCount || hTKKCMCCellInfo.CellTemplate.ColIndex > winFpSpread.Sheets[i].NonEmptyColumnCount)
					{
						continue;
					}
					if (hTKKCMCCellInfo.Controltype == "14")
					{
						string text = CheckFormatDdmmyyyy(winFpSpread, i, hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1, levelWarning);
						winFpSpread.Sheets[i].Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Text = text;
					}
					if (hTKKCMCCellInfo.Controltype == "104")
					{
						string text = CheckFormatDdmm(winFpSpread, i, hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1, levelWarning);
						winFpSpread.Sheets[i].Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Text = text;
					}
					if (hTKKCMCCellInfo.Controltype == "109")
					{
						string text = CheckFormatYyyy(winFpSpread, i, hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1, levelWarning);
						winFpSpread.Sheets[i].Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Text = text;
					}
					hTKKCMCCellInfo.Value = ConvertHelpers.ConvertCellValueToText(winFpSpread.Sheets[i], hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1);
					if (winFpSpread.Sheets[i].Rows[hTKKCMCCellInfo.CellTemplate.RowIndex - 1].Visible && !(hTKKCMCCellInfo.Path == "PL01_KHBS/NDUNGKHAIBOSUNGDIEUCHINH/TongHopDCTangGiamPNop/ct2") && !(hTKKCMCCellInfo.Path == "PL01_KHBS/NDUNGKHAIBOSUNGDIEUCHINH/TongHopDCTangGiamPNop/ct40/ct2") && !(hTKKCMCCellInfo.Path == "PL01_KHBS/NDUNGKHAIBOSUNGDIEUCHINH/TongHopDCTangGiamPNop/ct43/ct2") && !(hTKKCMCCellInfo.Path == "PL01_KHBS/NDUNGKHAIBOSUNGDIEUCHINH/TongHopDCTangGiamPNop/ct43/ct3") && !(hTKKCMCCellInfo.Path == "PL01_KHBS/NDUNGKHAIBOSUNGDIEUCHINH/TongHopDCTangGiamPNop/ct3"))
					{
						if (hTKKCMCCellInfo.Controltype == "16")
						{
							if (ConvertHelpers.ConvertToDouble(hTKKCMCCellInfo.Value) > 0.0)
							{
								WinSettings.ToKhaiInfo.SheetInfos[i].CoDuLieu = true;
							}
						}
						else if (hTKKCMCCellInfo.Controltype == "2")
						{
							if (((!(WinSettings.MenuID == "44") && !(WinSettings.MenuID == "42")) || !(hTKKCMCCellInfo.Encode == "0")) && ((!HTKKFormater.isNullString(hTKKCMCCellInfo.DefaultValue) && ConvertHelpers.ConvertToBool(hTKKCMCCellInfo.DefaultValue) != ConvertHelpers.ConvertToBool(hTKKCMCCellInfo.Value)) || (HTKKFormater.isNullString(hTKKCMCCellInfo.DefaultValue) && ConvertHelpers.ConvertToBool(hTKKCMCCellInfo.Value))))
							{
								WinSettings.ToKhaiInfo.SheetInfos[i].CoDuLieu = true;
							}
						}
						else if (hTKKCMCCellInfo.Controltype == "14")
						{
							string text = CheckFormatDdmmyyyy(winFpSpread, i, hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1, levelWarning);
							winFpSpread.Sheets[i].Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Text = text;
						}
						else if (hTKKCMCCellInfo.Controltype == "104")
						{
							string text = CheckFormatDdmm(winFpSpread, i, hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1, levelWarning);
							winFpSpread.Sheets[i].Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Text = text;
						}
						else if (hTKKCMCCellInfo.Controltype == "109")
						{
							string text = CheckFormatYyyy(winFpSpread, i, hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1, levelWarning);
							winFpSpread.Sheets[i].Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Text = text;
						}
						else if (!(hTKKCMCCellInfo.Controltype == "15"))
						{
							if (WinSettings.MenuID == "03" && hTKKCMCCellInfo.CellTemplate.ColIndex == 3 && "03_2A_TNDN;03_2B_TNDN".IndexOf(WinSettings.ToKhaiInfo.SheetInfos[i].DataFile) > -1)
							{
								continue;
							}
							if ((!HTKKFormater.isNullString(hTKKCMCCellInfo.DefaultValue) && !hTKKCMCCellInfo.DefaultValue.Equals(hTKKCMCCellInfo.Value)) || (HTKKFormater.isNullString(hTKKCMCCellInfo.DefaultValue) && !HTKKFormater.isNullString(hTKKCMCCellInfo.Value)))
							{
								WinSettings.ToKhaiInfo.SheetInfos[i].CoDuLieu = true;
							}
						}
					}
					switch (hTKKCMCCellInfo.Path)
					{
					case "HSoThueDTu/HSoKhaiThue/TTinChung/TTinTKhaiThue/TKhaiThue/ngayKy":
						WinSettings.TTinChung.ngayKy = hTKKCMCCellInfo.Value;
						break;
					case "HSoThueDTu/HSoKhaiThue/TTinChung/TTinTKhaiThue/TKhaiThue/nguoiKy":
						WinSettings.TTinChung.nguoiKy = hTKKCMCCellInfo.Value;
						break;
					case "HSoThueDTu/HSoKhaiThue/TTinChung/TTinTKhaiThue/DLyThue/NVienDLy/tenNVienDLyThue":
						WinSettings.TTinChung.tenNVienDLyThue = hTKKCMCCellInfo.Value;
						break;
					case "HSoThueDTu/HSoKhaiThue/TTinChung/TTinTKhaiThue/DLyThue/NVienDLy/cchiHNghe":
						WinSettings.TTinChung.cchiHNghe = hTKKCMCCellInfo.Value;
						break;
					case "HSoThueDTu/HSoKhaiThue/TTinChung/TTinTKhaiThue/ngayKy":
						WinSettings.TTinChung.ngayKy = hTKKCMCCellInfo.Value;
						break;
					case "HSoThueDTu/HSoKhaiThue/TTinChung/TTinTKhaiThue/nguoiKy":
						WinSettings.TTinChung.nguoiKy = hTKKCMCCellInfo.Value;
						break;
					case "HSoThueDTu/HSoKhaiThue/TTinChung/TTinTKhaiThue/NNT/phuongXa":
						WinSettings.TTinChung.phuongXa = hTKKCMCCellInfo.Value;
						break;
					case "HSoThueDTu/HSoKhaiThue/TTinChung/TTinTKhaiThue/NNT/dchiNNT":
						WinSettings.TTinChung.dchiNNT = hTKKCMCCellInfo.Value;
						break;
					case "HSoThueDTu/HSoKhaiThue/TTinChung/TTinTKhaiThue/NNT/maHuyenNNT":
						WinSettings.TTinChung.maHuyenNNT = hTKKCMCCellInfo.Value;
						break;
					case "HSoThueDTu/HSoKhaiThue/TTinChung/TTinTKhaiThue/NNT/tenHuyenNNT":
						WinSettings.TTinChung.tenHuyenNNT = hTKKCMCCellInfo.Value;
						break;
					case "HSoThueDTu/HSoKhaiThue/TTinChung/TTinTKhaiThue/NNT/maTinhNNT":
						WinSettings.TTinChung.maTinhNNT = hTKKCMCCellInfo.Value;
						break;
					case "HSoThueDTu/HSoKhaiThue/TTinChung/TTinTKhaiThue/NNT/tenTinhNNT":
						WinSettings.TTinChung.tenTinhNNT = hTKKCMCCellInfo.Value;
						break;
					case "HSoThueDTu/HSoKhaiThue/TTinChung/TTinTKhaiThue/NNT/dthoaiNNT":
						WinSettings.TTinChung.dthoaiNNT = hTKKCMCCellInfo.Value;
						break;
					case "HSoThueDTu/HSoKhaiThue/TTinChung/TTinTKhaiThue/NNT/faxNNT":
						WinSettings.TTinChung.faxNNT = hTKKCMCCellInfo.Value;
						break;
					case "HSoThueDTu/HSoKhaiThue/TTinChung/TTinTKhaiThue/NNT/emailNNT":
						WinSettings.TTinChung.emailNNT = hTKKCMCCellInfo.Value;
						break;
					}
				}
				while (true)
				{
					List<HTKKCMCTableInfo> tables = hTKKCMCSheetInfo.Tables;
					int j = 0;
					while (j < tables.Count)
					{
						bool flag = true;
						HTKKCMCTableInfo hTKKCMCTableInfo = tables[j];
						if (hTKKCMCTableInfo.Header.Length > 0)
						{
							for (int k = 0; k < hTKKCMCTableInfo.Header.StaticCells.Count; k++)
							{
								hTKKCMCTableInfo.Header.StaticCells[k].Value = ConvertHelpers.ConvertCellValueToTextWithControlType(winFpSpread.Sheets[i], hTKKCMCTableInfo.Header.StaticCells[k].CellTemplate.RowIndex - 1, hTKKCMCTableInfo.Header.StaticCells[k].CellTemplate.ColIndex - 1, hTKKCMCTableInfo.Header.StaticCells[k].Controltype);
								if (hTKKCMCTableInfo.Header.StaticCells[k].CellTemplate.ColIndex == 2)
								{
									continue;
								}
								if (hTKKCMCTableInfo.Header.StaticCells[k].Controltype == "16")
								{
									if (ConvertHelpers.ConvertToDouble(hTKKCMCTableInfo.Header.StaticCells[k].Value) != 0.0)
									{
										WinSettings.ToKhaiInfo.SheetInfos[i].CoDuLieu = true;
										flag = false;
									}
								}
								else if (hTKKCMCTableInfo.Header.StaticCells[k].Controltype == "2")
								{
									if ((!HTKKFormater.isNullString(hTKKCMCTableInfo.Header.StaticCells[k].DefaultValue) && ConvertHelpers.ConvertToBool(hTKKCMCTableInfo.Header.StaticCells[k].DefaultValue) != ConvertHelpers.ConvertToBool(hTKKCMCTableInfo.Header.StaticCells[k].Value)) || (HTKKFormater.isNullString(hTKKCMCTableInfo.Header.StaticCells[k].DefaultValue) && ConvertHelpers.ConvertToBool(hTKKCMCTableInfo.Header.StaticCells[k].Value)))
									{
										WinSettings.ToKhaiInfo.SheetInfos[i].CoDuLieu = true;
										flag = false;
									}
								}
								else if (hTKKCMCTableInfo.Header.StaticCells[k].Controltype == "14")
								{
									string text = CheckFormatDdmmyyyy(winFpSpread, i, hTKKCMCTableInfo.Header.StaticCells[k].CellTemplate.RowIndex - 1, hTKKCMCTableInfo.Header.StaticCells[k].CellTemplate.ColIndex - 1, levelWarning);
									winFpSpread.Sheets[i].Cells[hTKKCMCTableInfo.Header.StaticCells[k].CellTemplate.RowIndex - 1, hTKKCMCTableInfo.Header.StaticCells[k].CellTemplate.ColIndex - 1].Text = text;
									if ((!HTKKFormater.isNullString(hTKKCMCTableInfo.Header.StaticCells[k].DefaultValue) && !hTKKCMCTableInfo.Header.StaticCells[k].DefaultValue.Equals(hTKKCMCTableInfo.Header.StaticCells[k].Value)) || (HTKKFormater.isNullString(hTKKCMCTableInfo.Header.StaticCells[k].DefaultValue) && !HTKKFormater.isNullString(hTKKCMCTableInfo.Header.StaticCells[k].Value)))
									{
										WinSettings.ToKhaiInfo.SheetInfos[i].CoDuLieu = true;
										flag = false;
									}
								}
								else if (hTKKCMCTableInfo.Header.StaticCells[k].Controltype == "104")
								{
									string text = CheckFormatDdmm(winFpSpread, i, hTKKCMCTableInfo.Header.StaticCells[k].CellTemplate.RowIndex - 1, hTKKCMCTableInfo.Header.StaticCells[k].CellTemplate.ColIndex - 1, levelWarning);
									winFpSpread.Sheets[i].Cells[hTKKCMCTableInfo.Header.StaticCells[k].CellTemplate.RowIndex - 1, hTKKCMCTableInfo.Header.StaticCells[k].CellTemplate.ColIndex - 1].Text = text;
									if ((!HTKKFormater.isNullString(hTKKCMCTableInfo.Header.StaticCells[k].DefaultValue) && !hTKKCMCTableInfo.Header.StaticCells[k].DefaultValue.Equals(hTKKCMCTableInfo.Header.StaticCells[k].Value)) || (HTKKFormater.isNullString(hTKKCMCTableInfo.Header.StaticCells[k].DefaultValue) && !HTKKFormater.isNullString(hTKKCMCTableInfo.Header.StaticCells[k].Value)))
									{
										WinSettings.ToKhaiInfo.SheetInfos[i].CoDuLieu = true;
										flag = false;
									}
								}
								else if (hTKKCMCTableInfo.Header.StaticCells[k].Controltype == "109")
								{
									string text = CheckFormatYyyy(winFpSpread, i, hTKKCMCTableInfo.Header.StaticCells[k].CellTemplate.RowIndex - 1, hTKKCMCTableInfo.Header.StaticCells[k].CellTemplate.ColIndex - 1, levelWarning);
									winFpSpread.Sheets[i].Cells[hTKKCMCTableInfo.Header.StaticCells[k].CellTemplate.RowIndex - 1, hTKKCMCTableInfo.Header.StaticCells[k].CellTemplate.ColIndex - 1].Text = text;
									if ((!HTKKFormater.isNullString(hTKKCMCTableInfo.Header.StaticCells[k].DefaultValue) && !hTKKCMCTableInfo.Header.StaticCells[k].DefaultValue.Equals(hTKKCMCTableInfo.Header.StaticCells[k].Value)) || (HTKKFormater.isNullString(hTKKCMCTableInfo.Header.StaticCells[k].DefaultValue) && !HTKKFormater.isNullString(hTKKCMCTableInfo.Header.StaticCells[k].Value)))
									{
										WinSettings.ToKhaiInfo.SheetInfos[i].CoDuLieu = true;
										flag = false;
									}
								}
								else if ((!HTKKFormater.isNullString(hTKKCMCTableInfo.Header.StaticCells[k].DefaultValue) && !hTKKCMCTableInfo.Header.StaticCells[k].DefaultValue.Equals(hTKKCMCTableInfo.Header.StaticCells[k].Value)) || (HTKKFormater.isNullString(hTKKCMCTableInfo.Header.StaticCells[k].DefaultValue) && !HTKKFormater.isNullString(hTKKCMCTableInfo.Header.StaticCells[k].Value)))
								{
									WinSettings.ToKhaiInfo.SheetInfos[i].CoDuLieu = true;
									flag = false;
								}
							}
						}
						if (hTKKCMCTableInfo.Footer.Length > 0)
						{
							for (int k = 0; k < hTKKCMCTableInfo.Footer.StaticCells.Count; k++)
							{
								hTKKCMCTableInfo.Footer.StaticCells[k].Value = ConvertHelpers.ConvertCellValueToTextWithControlType(winFpSpread.Sheets[i], hTKKCMCTableInfo.Footer.StaticCells[k].CellTemplate.RowIndex - 1, hTKKCMCTableInfo.Footer.StaticCells[k].CellTemplate.ColIndex - 1, hTKKCMCTableInfo.Footer.StaticCells[k].Controltype);
								if (hTKKCMCTableInfo.Footer.StaticCells[k].Controltype == "14")
								{
									string text = CheckFormatDdmmyyyy(winFpSpread, i, hTKKCMCTableInfo.Footer.StaticCells[k].CellTemplate.RowIndex - 1, hTKKCMCTableInfo.Footer.StaticCells[k].CellTemplate.ColIndex - 1, levelWarning);
									winFpSpread.Sheets[i].Cells[hTKKCMCTableInfo.Footer.StaticCells[k].CellTemplate.RowIndex - 1, hTKKCMCTableInfo.Footer.StaticCells[k].CellTemplate.ColIndex - 1].Text = text;
								}
								if (hTKKCMCTableInfo.Footer.StaticCells[k].Controltype == "104")
								{
									string text = CheckFormatDdmm(winFpSpread, i, hTKKCMCTableInfo.Footer.StaticCells[k].CellTemplate.RowIndex - 1, hTKKCMCTableInfo.Footer.StaticCells[k].CellTemplate.ColIndex - 1, levelWarning);
									winFpSpread.Sheets[i].Cells[hTKKCMCTableInfo.Footer.StaticCells[k].CellTemplate.RowIndex - 1, hTKKCMCTableInfo.Footer.StaticCells[k].CellTemplate.ColIndex - 1].Text = text;
								}
								if (hTKKCMCTableInfo.Footer.StaticCells[k].Controltype == "109")
								{
									string text = CheckFormatYyyy(winFpSpread, i, hTKKCMCTableInfo.Footer.StaticCells[k].CellTemplate.RowIndex - 1, hTKKCMCTableInfo.Footer.StaticCells[k].CellTemplate.ColIndex - 1, levelWarning);
									winFpSpread.Sheets[i].Cells[hTKKCMCTableInfo.Footer.StaticCells[k].CellTemplate.RowIndex - 1, hTKKCMCTableInfo.Footer.StaticCells[k].CellTemplate.ColIndex - 1].Text = text;
								}
								if ((!HTKKFormater.isNullString(hTKKCMCTableInfo.Footer.StaticCells[k].DefaultValue) && !hTKKCMCTableInfo.Footer.StaticCells[k].DefaultValue.Equals(hTKKCMCTableInfo.Footer.StaticCells[k].Value)) || (HTKKFormater.isNullString(hTKKCMCTableInfo.Footer.StaticCells[k].DefaultValue) && !HTKKFormater.isNullString(hTKKCMCTableInfo.Footer.StaticCells[k].Value)))
								{
									WinSettings.ToKhaiInfo.SheetInfos[i].CoDuLieu = true;
									flag = false;
								}
							}
						}
						if (hTKKCMCTableInfo.Header.StaticCells.Count <= 0 || hTKKCMCTableInfo.RowsInfo.Cells.Count != 0 || !flag || !(hTKKCMCSheetInfo.Tables[j].TableID != "1"))
						{
							HTKKCMCRowInfo rowsInfo = hTKKCMCTableInfo.RowsInfo;
							Cell cell = null;
							for (int k = 0; k < hTKKCMCTableInfo.RowsValues.Count; k++)
							{
								bool flag2 = true;
								HTKKCMCRowValue hTKKCMCRowValue = hTKKCMCTableInfo.RowsValues[k];
								int count = hTKKCMCRowValue.RowValue.Count;
								for (int l = 0; l < count; l++)
								{
									List<string> rowValue = hTKKCMCRowValue.RowValue;
									int count2 = hTKKCMCTableInfo.RowsInfo.Cells.Count;
									if (l >= count2 || hTKKCMCTableInfo.Location.DataLocationTemplateRowStart + k - 1 > winFpSpread.Sheets[i].NonEmptyRowCount || hTKKCMCTableInfo.RowsInfo.Cells[l].CellTemplate.ColIndex - 1 > winFpSpread.Sheets[i].NonEmptyColumnCount)
									{
										continue;
									}
									rowValue[l] = ConvertHelpers.ConvertCellValueToText(winFpSpread.Sheets[i], hTKKCMCTableInfo.Location.DataLocationTemplateRowStart + k - 1, hTKKCMCTableInfo.RowsInfo.Cells[l].CellTemplate.ColIndex - 1);
									cell = winFpSpread.Sheets[i].Cells[hTKKCMCTableInfo.Location.DataLocationTemplateRowStart + k - 1, hTKKCMCTableInfo.RowsInfo.Cells[l].CellTemplate.ColIndex - 1];
									if (cell.CellType.ToString() == "MaskCellType")
									{
										ICellType cellType = cell.CellType;
										MaskCellType maskCellType = (MaskCellType)cellType;
										if (maskCellType.Mask == "00/00/0000")
										{
											string text = CheckFormatDdmmyyyy(winFpSpread, i, hTKKCMCTableInfo.Location.DataLocationTemplateRowStart + k - 1, hTKKCMCTableInfo.RowsInfo.Cells[l].CellTemplate.ColIndex - 1, levelWarning, hTKKCMCTableInfo.iColumnSTT);
											cell.Text = text;
										}
										if (rowsInfo.Cells[l].Controltype == "104")
										{
											string text = CheckFormatDdmm(winFpSpread, i, hTKKCMCTableInfo.Location.DataLocationTemplateRowStart + k - 1, hTKKCMCTableInfo.RowsInfo.Cells[l].CellTemplate.ColIndex - 1, levelWarning);
											cell.Text = text;
										}
										if (rowsInfo.Cells[l].Controltype == "109")
										{
											string text = CheckFormatYyyy(winFpSpread, i, hTKKCMCTableInfo.Location.DataLocationTemplateRowStart + k - 1, hTKKCMCTableInfo.RowsInfo.Cells[l].CellTemplate.ColIndex - 1, levelWarning);
											cell.Text = text;
										}
									}
									string text2 = ConvertHelpers.ConvertToString(rowsInfo.Cells[l].DefaultValue);
									if (l < rowValue.Count && !text2.Equals("NA"))
									{
										if (rowsInfo.Cells[l].Controltype == "16")
										{
											text2 = ConvertHelpers.ConvertToDouble(rowsInfo.Cells[l].DefaultValue).ToString();
										}
										else if (rowsInfo.Cells[l].Controltype == "2")
										{
											text2 = (ConvertHelpers.ConvertToBool(text2) ? "True" : "False");
										}
										if (!string.IsNullOrEmpty(rowValue[l]) && text2 != null && text2 != rowValue[l] && rowsInfo.Cells[l].CellTemplate.ColIndex != 2)
										{
											flag2 = false;
											WinSettings.ToKhaiInfo.SheetInfos[i].CoDuLieu = true;
										}
									}
								}
								if (flag2)
								{
									int frRow = hTKKCMCSheetInfo.Tables[j].Location.DataLocationTemplateRowStart - 1 + k;
									if (hTKKCMCSheetInfo.Tables[j].RowsValues.Count > 1)
									{
										hTKKCMCSheetInfo.Tables[j].RowsValues.RemoveAt(k);
										WinSpreadProcess.RemoveDataRows(winFpSpread.Sheets[i], frRow, 1, hTKKCMCSheetInfo);
										WinSpreadProcess.fpSpReindexingTableTemplate(winFpSpread.Sheets[i], hTKKCMCSheetInfo.Tables[j]);
										k--;
									}
									if (!(hTKKCMCSheetInfo.Tables[j].TableID != "1") || hTKKCMCSheetInfo.Tables[j].RowsValues.Count != 1 || GeneralFunctions.HaveToValidateRow(hTKKCMCSheetInfo.Tables[j].RowsInfo, hTKKCMCSheetInfo.Tables[j].RowsValues[0].RowValue.ToArray()) || hTKKCMCSheetInfo.Tables[j].RowsInfo.Cells.Count <= 0 || !flag)
									{
										continue;
									}
									goto IL_1b1a;
								}
								WinSettings.ToKhaiInfo.SheetInfos[i].CoDuLieu = true;
							}
							j++;
							continue;
						}
						goto IL_1571;
					}
					break;
					IL_1b1a:
					WinSpreadProcess.RemoveTable(winFpSpread.Sheets[i], j, hTKKCMCSheetInfo);
					WinSpreadProcess.fpSpReindexingSheetTemplate(winFpSpread.Sheets[i], hTKKCMCSheetInfo);
					continue;
					IL_1571:
					WinSpreadProcess.RemoveTable(winFpSpread.Sheets[i], j, hTKKCMCSheetInfo);
					WinSpreadProcess.fpSpReindexingSheetTemplate(winFpSpread.Sheets[i], hTKKCMCSheetInfo);
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void WinDoWorkPrint(FpSpread winFpSpread, HTKKPeriodInfo periodInfo)
	{
		int num = 0;
		string text = "";
		try
		{
			List<string> list = new List<string>();
			list.Add(GetPrefix(WinSettings.ToKhaiInfo, periodInfo, WinSettings.HTKKHeaderInfo));
			text = GetPrefixFirstBarcode(WinSettings.ToKhaiInfo, periodInfo, WinSettings.HTKKHeaderInfo);
			Dictionary<string, string> dictionary = new Dictionary<string, string>();
			int num2 = 1;
			string text2 = text + WinDoWorkGetBarcodeFromSheet(winFpSpread.Sheets[0], WinSettings.ToKhaiInfo.SheetInfos[0]);
			dictionary.Add("0", ConvertHelpers.UNICODETOTCVN(text2));
			if (periodInfo.menuID != "33")
			{
				for (short num3 = 1; num3 < WinSettings.ToKhaiInfo.SheetInfos.Count; num3++)
				{
					if (WinSettings.ToKhaiInfo.SheetInfos[num3] != null && ((periodInfo.menuID != "37" && WinSettings.DicSheet.ContainsKey(WinSettings.ToKhaiInfo.SheetInfos[num3].SheetName)) || (periodInfo.menuID == "37" && WinSettings.ToKhaiInfo.SheetInfos[num3].CoDuLieu)))
					{
						string empty = string.Empty;
						num = num3;
						empty = WinDoWorkGetBarcodeFromSheet(winFpSpread.Sheets[num3], WinSettings.ToKhaiInfo.SheetInfos[num3]);
						text2 += empty;
						empty = ConvertHelpers.UNICODETOTCVN(empty);
						if (WinSettings.ToKhaiInfo.SheetInfos[num3].CoDuLieu)
						{
							dictionary.Add(num2.ToString(), empty);
							num2++;
						}
					}
				}
			}
			text2 = ConvertHelpers.UNICODETOTCVN(text2);
			HTKKFile.Log("BACODE: " + text2);
			list.Add(text2);
			frmHTKKPrintOption form = new frmHTKKPrintOption(list.ToArray(), winFpSpread, dictionary);
			FormProcess.AddControl((Form)(object)form, (Control)(object)WinSettings.myWinToKhai.panelChild);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static string CheckFormatDateOnChange(string sDate)
	{
		//IL_01fa: Unknown result type (might be due to invalid IL or missing references)
		string text = sDate;
		bool flag = true;
		try
		{
			string text2 = "";
			string text3 = "";
			string text4 = "";
			if (!sDate.Trim().Equals("") || sDate.Trim().Length > 0)
			{
				string[] array = sDate.Split(new char[1] { '/' });
				text2 = array[0];
				text3 = array[1];
				text4 = array[2].Replace(".", "");
				if (ConvertHelpers.ConvertToInt(text2) >= 1 && ConvertHelpers.ConvertToInt(text2) <= 31)
				{
					text2 = text2.PadLeft(2, '0');
				}
				else
				{
					flag = false;
				}
				if (ConvertHelpers.ConvertToInt(text3) >= 1 && ConvertHelpers.ConvertToInt(text3) <= 12)
				{
					text3 = text3.PadLeft(2, '0');
				}
				else
				{
					flag = false;
				}
				if (ConvertHelpers.ConvertToInt(text4) >= 0 && ConvertHelpers.ConvertToInt(text4) <= 9999)
				{
					if (ConvertHelpers.ConvertToInt(text4) >= 0 && ConvertHelpers.ConvertToInt(text4) <= 999)
					{
						text4 = ConvertHelpers.ConvertToString(2000 + ConvertHelpers.ConvertToInt(text4));
					}
					if (ConvertHelpers.ConvertToInt(text4) < 1900)
					{
						flag = false;
					}
				}
				else
				{
					flag = false;
				}
				text = text2 + "/" + text3 + "/" + text4;
				DateTime result = DateTime.MinValue;
				DateTime.TryParseExact(text, "dd/MM/yyyy", CultureInfo.CurrentCulture, DateTimeStyles.None, out result);
				if (result == DateTime.MinValue)
				{
					flag = false;
				}
			}
			if (!flag)
			{
				frmMessage frmMessage = new frmMessage("0403", "3", "1", "Thông báo");
				((Form)frmMessage).ShowDialog();
				text = "";
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return text;
	}

	public static string CheckFormatDdmmyyyy(FpSpread winFpSpread, int activeSheet, int iRow, int iCol, Dictionary<int, int> levelWarning, int colSttOnSViewOfTable = -1)
	{
		string text = winFpSpread.Sheets[activeSheet].Cells[iRow, iCol].Text;
		string text2 = text;
		try
		{
			string rowSTT = "";
			if (colSttOnSViewOfTable > 0 && !string.IsNullOrEmpty(winFpSpread.Sheets[activeSheet].Cells[iRow, colSttOnSViewOfTable - 1].Text))
			{
				rowSTT = winFpSpread.Sheets[activeSheet].Cells[iRow, colSttOnSViewOfTable - 1].Text;
			}
			string text3 = "";
			string text4 = "";
			string text5 = "";
			if (!text.Trim().Equals("") || text.Trim().Length > 0)
			{
				string[] array = text.Split(new char[1] { '/' });
				text3 = array[0];
				text4 = array[1];
				text5 = array[2].Replace(".", "");
				int value;
				if (ConvertHelpers.ConvertToInt(text3) < 1 || ConvertHelpers.ConvertToInt(text3) > 31)
				{
					HTKKFile.SetNodeBackGroundPink2(winFpSpread.Sheets[activeSheet], iRow, iCol, HTKKFile.GetMessage("0403"), rowSTT);
					WinSettings.ToKhaiInfo.SheetInfos[activeSheet].isValidData = false;
					if (levelWarning.TryGetValue(activeSheet, out value))
					{
						levelWarning[activeSheet] = 2;
					}
					else
					{
						levelWarning.Add(activeSheet, 2);
					}
					return text2;
				}
				text3 = text3.PadLeft(2, '0');
				if (ConvertHelpers.ConvertToInt(text4) < 1 || ConvertHelpers.ConvertToInt(text4) > 12)
				{
					HTKKFile.SetNodeBackGroundPink2(winFpSpread.Sheets[activeSheet], iRow, iCol, HTKKFile.GetMessage("0403"), rowSTT);
					WinSettings.ToKhaiInfo.SheetInfos[activeSheet].isValidData = false;
					if (levelWarning.TryGetValue(activeSheet, out value))
					{
						levelWarning[activeSheet] = 2;
					}
					else
					{
						levelWarning.Add(activeSheet, 2);
					}
					return text2;
				}
				text4 = text4.PadLeft(2, '0');
				if (ConvertHelpers.ConvertToInt(text5) < 0 || ConvertHelpers.ConvertToInt(text5) > 9999)
				{
					HTKKFile.SetNodeBackGroundPink2(winFpSpread.Sheets[activeSheet], iRow, iCol, HTKKFile.GetMessage("0403"), rowSTT);
					WinSettings.ToKhaiInfo.SheetInfos[activeSheet].isValidData = false;
					if (levelWarning.TryGetValue(activeSheet, out value))
					{
						levelWarning[activeSheet] = 2;
					}
					else
					{
						levelWarning.Add(activeSheet, 2);
					}
					return text2;
				}
				if (ConvertHelpers.ConvertToInt(text5) >= 0 && ConvertHelpers.ConvertToInt(text5) <= 999)
				{
					text5 = ConvertHelpers.ConvertToString(2000 + ConvertHelpers.ConvertToInt(text5));
				}
				if (ConvertHelpers.ConvertToInt(text5) < 1900)
				{
					HTKKFile.SetNodeBackGroundPink2(winFpSpread.Sheets[activeSheet], iRow, iCol, HTKKFile.GetMessage("0403"), rowSTT);
					WinSettings.ToKhaiInfo.SheetInfos[activeSheet].isValidData = false;
					if (levelWarning.TryGetValue(activeSheet, out value))
					{
						levelWarning[activeSheet] = 2;
					}
					else
					{
						levelWarning.Add(activeSheet, 2);
					}
					return text2;
				}
				text2 = text3 + "/" + text4 + "/" + text5;
				DateTime result = DateTime.MinValue;
				DateTime.TryParseExact(text2, "dd/MM/yyyy", CultureInfo.CurrentCulture, DateTimeStyles.None, out result);
				if (result == DateTime.MinValue)
				{
					HTKKFile.SetNodeBackGroundPink2(winFpSpread.Sheets[activeSheet], iRow, iCol, HTKKFile.GetMessage("0403"), rowSTT);
					WinSettings.ToKhaiInfo.SheetInfos[activeSheet].isValidData = false;
					if (levelWarning.TryGetValue(activeSheet, out value))
					{
						levelWarning[activeSheet] = 2;
					}
					else
					{
						levelWarning.Add(activeSheet, 2);
					}
					return text2;
				}
				HTKKFile.RemoveValidate(winFpSpread.Sheets[activeSheet], iRow, iCol);
			}
			else if (winFpSpread.Sheets[activeSheet].Cells[iRow, iCol].Note.Equals(HTKKFile.GetMessage("0403")))
			{
				HTKKFile.RemoveValidate(winFpSpread.Sheets[activeSheet], iRow, iCol);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return text2;
	}

	public static string CheckFormatDdmm(FpSpread winFpSpread, int activeSheet, int iRow, int iCol, Dictionary<int, int> levelWarning)
	{
		string text = winFpSpread.Sheets[activeSheet].Cells[iRow, iCol].Text;
		string result = text;
		try
		{
			string text2 = "";
			string text3 = "";
			if (!text.Trim().Equals("") || text.Trim().Length > 0)
			{
				int value;
				if (text.IndexOf(".") >= 0)
				{
					HTKKFile.SetNodeBackGroundPink(winFpSpread.Sheets[activeSheet], iRow, iCol, HTKKFile.GetMessage("0660"));
					WinSettings.ToKhaiInfo.SheetInfos[activeSheet].isValidData = false;
					if (levelWarning.TryGetValue(activeSheet, out value))
					{
						levelWarning[activeSheet] = 2;
					}
					else
					{
						levelWarning.Add(activeSheet, 2);
					}
					return result;
				}
				string[] array = text.Split(new char[1] { '/' });
				text2 = array[0];
				text3 = array[1];
				if (ConvertHelpers.ConvertToInt(text2) < 1 || ConvertHelpers.ConvertToInt(text2) > 31)
				{
					HTKKFile.SetNodeBackGroundPink(winFpSpread.Sheets[activeSheet], iRow, iCol, HTKKFile.GetMessage("0403"));
					WinSettings.ToKhaiInfo.SheetInfos[activeSheet].isValidData = false;
					if (levelWarning.TryGetValue(activeSheet, out value))
					{
						levelWarning[activeSheet] = 2;
					}
					else
					{
						levelWarning.Add(activeSheet, 2);
					}
					return result;
				}
				text2 = text2.PadLeft(2, '0');
				if (ConvertHelpers.ConvertToInt(text3) < 1 || ConvertHelpers.ConvertToInt(text3) > 12)
				{
					HTKKFile.SetNodeBackGroundPink(winFpSpread.Sheets[activeSheet], iRow, iCol, HTKKFile.GetMessage("0451"));
					WinSettings.ToKhaiInfo.SheetInfos[activeSheet].isValidData = false;
					if (levelWarning.TryGetValue(activeSheet, out value))
					{
						levelWarning[activeSheet] = 2;
					}
					else
					{
						levelWarning.Add(activeSheet, 2);
					}
					return result;
				}
				text3 = text3.PadLeft(2, '0');
				if (ConvertHelpers.ConvertToInt(text3) == 2 && ConvertHelpers.ConvertToInt(text2) > 29)
				{
					HTKKFile.SetNodeBackGroundPink(winFpSpread.Sheets[activeSheet], iRow, iCol, HTKKFile.GetMessage("0403"));
					WinSettings.ToKhaiInfo.SheetInfos[activeSheet].isValidData = false;
					if (levelWarning.TryGetValue(activeSheet, out value))
					{
						levelWarning[activeSheet] = 2;
					}
					else
					{
						levelWarning.Add(activeSheet, 2);
					}
					return result;
				}
				if (ConvertHelpers.ConvertToInt(text3) != 2)
				{
					string s = text2 + "/" + text3 + "/" + DateTime.Now.Year;
					DateTime result2 = DateTime.MinValue;
					DateTime.TryParseExact(s, "dd/MM/yyyy", CultureInfo.CurrentCulture, DateTimeStyles.None, out result2);
					if (result2 == DateTime.MinValue)
					{
						HTKKFile.SetNodeBackGroundPink(winFpSpread.Sheets[activeSheet], iRow, iCol, HTKKFile.GetMessage("0403"));
						WinSettings.ToKhaiInfo.SheetInfos[activeSheet].isValidData = false;
						if (levelWarning.TryGetValue(activeSheet, out value))
						{
							levelWarning[activeSheet] = 2;
						}
						else
						{
							levelWarning.Add(activeSheet, 2);
						}
						return result;
					}
				}
				result = text2 + "/" + text3;
				HTKKFile.RemoveValidate(winFpSpread.Sheets[activeSheet], iRow, iCol);
			}
			else if (winFpSpread.Sheets[activeSheet].Cells[iRow, iCol].Note.Equals(HTKKFile.GetMessage("0403")))
			{
				HTKKFile.RemoveValidate(winFpSpread.Sheets[activeSheet], iRow, iCol);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return result;
	}

	public static string CheckFormatMMyyyy(FpSpread winFpSpread, int activeSheet, LeaveCellEventArgs e, Dictionary<int, int> levelWarning)
	{
		//IL_00f3: Unknown result type (might be due to invalid IL or missing references)
		//IL_02a2: Unknown result type (might be due to invalid IL or missing references)
		//IL_0209: Unknown result type (might be due to invalid IL or missing references)
		string text = winFpSpread.Sheets[activeSheet].Cells[e.Row, e.Column].Text;
		string result = text;
		try
		{
			string text2 = "";
			string text3 = "";
			string text4 = "";
			if (!text.Trim().Equals("") || text.Trim().Length > 0)
			{
				string[] array = text.Split(new char[1] { '/' });
				text3 = array[0];
				text4 = array[1].Replace(".", "");
				int value;
				if (ConvertHelpers.ConvertToInt(text3) < 1 || ConvertHelpers.ConvertToInt(text3) > 12)
				{
					frmMessage frmMessage = new frmMessage("0451", "3", "1", "Thông báo");
					((Form)frmMessage).ShowDialog();
					e.Cancel = true;
					winFpSpread.Sheets[activeSheet].SetActiveCell(e.Row, e.Column);
					WinSettings.ToKhaiInfo.SheetInfos[activeSheet].isValidData = false;
					if (levelWarning.TryGetValue(activeSheet, out value))
					{
						levelWarning[activeSheet] = 2;
					}
					else
					{
						levelWarning.Add(activeSheet, 2);
					}
					return result;
				}
				text3 = text3.PadLeft(2, '0');
				if (ConvertHelpers.ConvertToInt(text4) < 0 || ConvertHelpers.ConvertToInt(text4) > 9999)
				{
					frmMessage frmMessage = new frmMessage("0451", "3", "1", "Thông báo");
					((Form)frmMessage).ShowDialog();
					e.Cancel = true;
					winFpSpread.Sheets[activeSheet].SetActiveCell(e.Row, e.Column);
					WinSettings.ToKhaiInfo.SheetInfos[activeSheet].isValidData = false;
					if (levelWarning.TryGetValue(activeSheet, out value))
					{
						levelWarning[activeSheet] = 2;
					}
					else
					{
						levelWarning.Add(activeSheet, 2);
					}
					return result;
				}
				if (ConvertHelpers.ConvertToInt(text4) >= 0 && ConvertHelpers.ConvertToInt(text4) <= 999)
				{
					text4 = ConvertHelpers.ConvertToString(2000 + ConvertHelpers.ConvertToInt(text4));
				}
				if (ConvertHelpers.ConvertToInt(text4) < 1900)
				{
					frmMessage frmMessage = new frmMessage("0451", "3", "1", "Thông báo");
					((Form)frmMessage).ShowDialog();
					e.Cancel = true;
					winFpSpread.Sheets[activeSheet].SetActiveCell(e.Row, e.Column);
					WinSettings.ToKhaiInfo.SheetInfos[activeSheet].isValidData = false;
					if (levelWarning.TryGetValue(activeSheet, out value))
					{
						levelWarning[activeSheet] = 2;
					}
					else
					{
						levelWarning.Add(activeSheet, 2);
					}
					return result;
				}
				result = text3 + "/" + text4;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return result;
	}

	public static string CheckFormatYyyy(FpSpread winFpSpread, int activeSheet, int iRow, int iCol, Dictionary<int, int> levelWarning)
	{
		//IL_017b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0108: Unknown result type (might be due to invalid IL or missing references)
		string text = winFpSpread.Sheets[activeSheet].Cells[iRow, iCol].Text;
		string result = text;
		try
		{
			string text2 = "";
			string text3 = "";
			string text4 = "";
			if (!text.Trim().Equals("") || text.Trim().Length > 0)
			{
				text4 = text;
				int value;
				if (ConvertHelpers.ConvertToInt(text4) < 0 || ConvertHelpers.ConvertToInt(text4) > 9999)
				{
					frmMessage frmMessage = new frmMessage("0451", "3", "1", "Thông báo");
					((Form)frmMessage).ShowDialog();
					WinSettings.ToKhaiInfo.SheetInfos[activeSheet].isValidData = false;
					if (levelWarning.TryGetValue(activeSheet, out value))
					{
						levelWarning[activeSheet] = 2;
					}
					else
					{
						levelWarning.Add(activeSheet, 2);
					}
					return result;
				}
				if (ConvertHelpers.ConvertToInt(text4) >= 0 && ConvertHelpers.ConvertToInt(text4) <= 999)
				{
					text4 = ConvertHelpers.ConvertToString(2000 + ConvertHelpers.ConvertToInt(text4));
				}
				if (ConvertHelpers.ConvertToInt(text4) < 1900)
				{
					frmMessage frmMessage = new frmMessage("0451", "3", "1", "Thông báo");
					((Form)frmMessage).ShowDialog();
					WinSettings.ToKhaiInfo.SheetInfos[activeSheet].isValidData = false;
					if (levelWarning.TryGetValue(activeSheet, out value))
					{
						levelWarning[activeSheet] = 2;
					}
					else
					{
						levelWarning.Add(activeSheet, 2);
					}
					return result;
				}
				result = text4;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return result;
	}

	public static List<string> getListBarcode(string strPrintBarcode, string strHeaderBarcode)
	{
		List<string> list = new List<string>();
		try
		{
			int num = strPrintBarcode.Length / 200 + 1;
			int num2 = strPrintBarcode.Length / num;
			int num3 = strPrintBarcode.Length - num * num2;
			for (int i = 0; i < num3; i++)
			{
				list.Add(strHeaderBarcode + strPrintBarcode.Substring(i * (num2 + 1), num2 + 1));
			}
			for (int i = num3; i < num; i++)
			{
				list.Add(strHeaderBarcode + strPrintBarcode.Substring(num3 * (num2 + 1) + (i - num3) * num2, num2));
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return list;
	}

	public static string[] getMucLucNganSach()
	{
		try
		{
			string menuID = WinSettings.TTPeriodInfo.menuID;
			string tieuMucHachToanVal = WinSettings.TTPeriodInfo.TieuMucHachToanVal;
			string[] result = null;
			XMLMessage xMLMessage = new XMLMessage("../Project/Message.xml");
			if (menuID.Equals("15") || menuID.Equals("16"))
			{
				result = new string[2]
				{
					xMLMessage.GetMessage("0127"),
					xMLMessage.GetMessage("0178")
				};
			}
			else if (menuID.Equals("37"))
			{
				result = new string[2]
				{
					xMLMessage.GetMessage("0127"),
					xMLMessage.GetMessage("0128")
				};
			}
			else if (menuID.Equals("48") || menuID.Equals("49"))
			{
				result = new string[2]
				{
					xMLMessage.GetMessage("0127"),
					xMLMessage.GetMessage("0190")
				};
			}
			else if (menuID.Equals("50") || menuID.Equals("51"))
			{
				result = new string[6]
				{
					xMLMessage.GetMessage("0127"),
					xMLMessage.GetMessage("0179"),
					xMLMessage.GetMessage("0180"),
					xMLMessage.GetMessage("0181"),
					xMLMessage.GetMessage("0182"),
					xMLMessage.GetMessage("0287")
				};
			}
			else if (menuID.Equals("36") || menuID.Equals("35"))
			{
				result = new string[2]
				{
					xMLMessage.GetMessage("0127"),
					xMLMessage.GetMessage("0134")
				};
			}
			else if (menuID.Equals("31"))
			{
				result = new string[3]
				{
					xMLMessage.GetMessage("0127"),
					xMLMessage.GetMessage("0366"),
					xMLMessage.GetMessage("0367")
				};
			}
			else if (menuID.Equals("70"))
			{
				result = new string[3]
				{
					xMLMessage.GetMessage("0127"),
					xMLMessage.GetMessage("0191"),
					xMLMessage.GetMessage("0192")
				};
			}
			else if (menuID.Equals("81"))
			{
				result = new string[2]
				{
					xMLMessage.GetMessage("0127"),
					xMLMessage.GetMessage("0258")
				};
			}
			else if (menuID.Equals("82"))
			{
				result = new string[2]
				{
					xMLMessage.GetMessage("0127"),
					xMLMessage.GetMessage("0289")
				};
			}
			else if (menuID.Equals("04"))
			{
				result = new string[2]
				{
					xMLMessage.GetMessage("0127"),
					xMLMessage.GetMessage("0193")
				};
			}
			else if ((menuID.Equals("05") || menuID.Equals("79")) && !tieuMucHachToanVal.Equals("0"))
			{
				if (!WinSettings.TTPeriodInfo.TieuMucHachToanVal.Equals("0000") && menuID.Equals("79"))
				{
					result = new string[2]
					{
						xMLMessage.GetMessage("0127"),
						xMLMessage.GetMessage("0194")
					};
				}
				else if (!menuID.Equals("79"))
				{
					result = new string[2]
					{
						xMLMessage.GetMessage("0127"),
						xMLMessage.GetMessage("0194")
					};
				}
			}
			else if (menuID.Equals("11"))
			{
				if (tieuMucHachToanVal.Equals("1052"))
				{
					result = new string[2]
					{
						xMLMessage.GetMessage("0127"),
						xMLMessage.GetMessage("0195")
					};
				}
				else if (tieuMucHachToanVal.Equals("1056"))
				{
					result = new string[2]
					{
						xMLMessage.GetMessage("0127"),
						xMLMessage.GetMessage("0196")
					};
				}
				else if (tieuMucHachToanVal.Equals("1057"))
				{
					result = new string[2]
					{
						xMLMessage.GetMessage("0127"),
						xMLMessage.GetMessage("0197")
					};
				}
			}
			else if (menuID.Equals("12"))
			{
				if (tieuMucHachToanVal.Equals("1052"))
				{
					result = new string[2]
					{
						xMLMessage.GetMessage("0127"),
						xMLMessage.GetMessage("0198")
					};
				}
				else if (tieuMucHachToanVal.Equals("1056"))
				{
					result = new string[2]
					{
						xMLMessage.GetMessage("0127"),
						xMLMessage.GetMessage("0199")
					};
				}
				else if (tieuMucHachToanVal.Equals("1057"))
				{
					result = new string[2]
					{
						xMLMessage.GetMessage("0127"),
						xMLMessage.GetMessage("0200")
					};
				}
			}
			else if (menuID.Equals("03"))
			{
				if (tieuMucHachToanVal.Equals("1052"))
				{
					result = new string[2]
					{
						xMLMessage.GetMessage("0127"),
						xMLMessage.GetMessage("0249")
					};
				}
				else if (tieuMucHachToanVal.Equals("1056"))
				{
					result = new string[2]
					{
						xMLMessage.GetMessage("0127"),
						xMLMessage.GetMessage("0250")
					};
				}
				else if (tieuMucHachToanVal.Equals("1057"))
				{
					result = new string[2]
					{
						xMLMessage.GetMessage("0127"),
						xMLMessage.GetMessage("0251")
					};
				}
			}
			else if (menuID.Equals("01"))
			{
				result = ((!tieuMucHachToanVal.Equals("1704") && !tieuMucHachToanVal.Equals("05") && !tieuMucHachToanVal.Equals("08")) ? ((!tieuMucHachToanVal.Equals("1705") && !tieuMucHachToanVal.Equals("01") && !tieuMucHachToanVal.Equals("09")) ? new string[2]
				{
					xMLMessage.GetMessage("0127"),
					xMLMessage.GetMessage("0201")
				} : new string[2]
				{
					xMLMessage.GetMessage("0127"),
					xMLMessage.GetMessage("0203")
				}) : new string[2]
				{
					xMLMessage.GetMessage("0127"),
					xMLMessage.GetMessage("0202")
				});
			}
			else if (menuID.Equals("73"))
			{
				result = new string[2]
				{
					xMLMessage.GetMessage("0127"),
					xMLMessage.GetMessage("0348")
				};
			}
			else if (menuID.Equals("56"))
			{
				result = new string[2]
				{
					xMLMessage.GetMessage("0127"),
					xMLMessage.GetMessage("0355")
				};
			}
			else if (menuID.Equals("59"))
			{
				result = new string[5]
				{
					xMLMessage.GetMessage("0231"),
					xMLMessage.GetMessage("0232"),
					xMLMessage.GetMessage("0233"),
					xMLMessage.GetMessage("0234"),
					xMLMessage.GetMessage("0232")
				};
			}
			else if (menuID.Equals("98"))
			{
				if (tieuMucHachToanVal.Equals("1"))
				{
					result = new string[2]
					{
						xMLMessage.GetMessage("0127"),
						xMLMessage.GetMessage("0300")
					};
				}
				else if (tieuMucHachToanVal.Equals("2"))
				{
					result = new string[2]
					{
						xMLMessage.GetMessage("0127"),
						xMLMessage.GetMessage("0302")
					};
				}
				else if (tieuMucHachToanVal.Equals("3"))
				{
					result = new string[2]
					{
						xMLMessage.GetMessage("0127"),
						xMLMessage.GetMessage("0301")
					};
				}
			}
			else if (menuID.Equals("08"))
			{
				if (tieuMucHachToanVal.Equals("1"))
				{
					result = new string[2]
					{
						xMLMessage.GetMessage("0127"),
						xMLMessage.GetMessage("0420")
					};
				}
				else if (tieuMucHachToanVal.Equals("2"))
				{
					result = new string[2]
					{
						xMLMessage.GetMessage("0127"),
						xMLMessage.GetMessage("0422")
					};
				}
				else if (tieuMucHachToanVal.Equals("3"))
				{
					result = new string[2]
					{
						xMLMessage.GetMessage("0127"),
						xMLMessage.GetMessage("0421")
					};
				}
			}
			else if (menuID.Equals("92") || menuID.Equals("40"))
			{
				if (tieuMucHachToanVal.Equals("1"))
				{
					result = new string[2]
					{
						xMLMessage.GetMessage("0127"),
						xMLMessage.GetMessage("0303")
					};
				}
				else if (tieuMucHachToanVal.Equals("2"))
				{
					result = new string[2]
					{
						xMLMessage.GetMessage("0127"),
						xMLMessage.GetMessage("0305")
					};
				}
				else if (tieuMucHachToanVal.Equals("3"))
				{
					result = new string[2]
					{
						xMLMessage.GetMessage("0127"),
						xMLMessage.GetMessage("0304")
					};
				}
			}
			else if (menuID.Equals("93"))
			{
				if (tieuMucHachToanVal.Equals("1"))
				{
					result = new string[2]
					{
						xMLMessage.GetMessage("0127"),
						xMLMessage.GetMessage("0326")
					};
				}
				else if (tieuMucHachToanVal.Equals("2"))
				{
					result = new string[2]
					{
						xMLMessage.GetMessage("0127"),
						xMLMessage.GetMessage("0328")
					};
				}
				else if (tieuMucHachToanVal.Equals("3"))
				{
					result = new string[2]
					{
						xMLMessage.GetMessage("0127"),
						xMLMessage.GetMessage("0327")
					};
				}
			}
			else if (menuID.Equals("44"))
			{
				if (tieuMucHachToanVal.Equals("1"))
				{
					result = new string[2]
					{
						xMLMessage.GetMessage("0127"),
						xMLMessage.GetMessage("0417")
					};
				}
				else if (tieuMucHachToanVal.Equals("2"))
				{
					result = new string[2]
					{
						xMLMessage.GetMessage("0127"),
						xMLMessage.GetMessage("0419")
					};
				}
				else if (tieuMucHachToanVal.Equals("3"))
				{
					result = new string[2]
					{
						xMLMessage.GetMessage("0127"),
						xMLMessage.GetMessage("0418")
					};
				}
			}
			else if (menuID.Equals("89") || menuID.Equals("42"))
			{
				if (tieuMucHachToanVal.Equals("1"))
				{
					result = new string[2]
					{
						xMLMessage.GetMessage("0127"),
						xMLMessage.GetMessage("0329")
					};
				}
				else if (tieuMucHachToanVal.Equals("2"))
				{
					result = new string[2]
					{
						xMLMessage.GetMessage("0127"),
						xMLMessage.GetMessage("0331")
					};
				}
				else if (tieuMucHachToanVal.Equals("3"))
				{
					result = new string[2]
					{
						xMLMessage.GetMessage("0127"),
						xMLMessage.GetMessage("0330")
					};
				}
			}
			else if (menuID.Equals("32"))
			{
				result = new string[2]
				{
					xMLMessage.GetMessage("0127"),
					xMLMessage.GetMessage("0359")
				};
			}
			else if (menuID.Equals("34"))
			{
				result = new string[7]
				{
					xMLMessage.GetMessage("0127"),
					xMLMessage.GetMessage("0360"),
					xMLMessage.GetMessage("0496"),
					xMLMessage.GetMessage("0361"),
					xMLMessage.GetMessage("0362"),
					xMLMessage.GetMessage("0363"),
					xMLMessage.GetMessage("0495")
				};
			}
			else if (menuID.Equals("96"))
			{
				result = new string[2]
				{
					xMLMessage.GetMessage("0127"),
					xMLMessage.GetMessage("0308")
				};
			}
			else if (menuID.Equals("25"))
			{
				result = new string[3]
				{
					xMLMessage.GetMessage("0127"),
					xMLMessage.GetMessage("0322"),
					xMLMessage.GetMessage("0323")
				};
			}
			else if (menuID.Equals("26"))
			{
				result = new string[3]
				{
					xMLMessage.GetMessage("0127"),
					xMLMessage.GetMessage("0320"),
					xMLMessage.GetMessage("0321")
				};
			}
			else if (menuID.Equals("80"))
			{
				result = new string[3]
				{
					xMLMessage.GetMessage("0127"),
					xMLMessage.GetMessage("0333"),
					xMLMessage.GetMessage("0334")
				};
			}
			else if (menuID.Equals("29"))
			{
				result = new string[2]
				{
					xMLMessage.GetMessage("0127"),
					xMLMessage.GetMessage("0374")
				};
			}
			else if (menuID.Equals("21"))
			{
				result = new string[2]
				{
					xMLMessage.GetMessage("0127"),
					xMLMessage.GetMessage("0444")
				};
			}
			else if (menuID.Equals("47"))
			{
				result = new string[1] { xMLMessage.GetMessage("0448") };
			}
			else if (menuID.Equals("94"))
			{
				result = new string[2]
				{
					xMLMessage.GetMessage("0127"),
					xMLMessage.GetMessage("0193")
				};
			}
			else if (menuID.Equals("96"))
			{
				result = new string[2]
				{
					xMLMessage.GetMessage("0127"),
					xMLMessage.GetMessage("0454")
				};
			}
			else if (menuID.Equals("97"))
			{
				result = new string[2]
				{
					xMLMessage.GetMessage("0127"),
					xMLMessage.GetMessage("0455")
				};
			}
			return result;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public static void ExportXML(XmlNode xmlNode, int index)
	{
		try
		{
			MethodInfo method = GetMethod("ExportXML_" + index);
			if ((object)method == null)
			{
				return;
			}
			ParameterInfo[] parameters = method.GetParameters();
			object[] array = new object[parameters.Length];
			for (int i = 0; i < parameters.Length; i++)
			{
				switch (parameters[i].Name)
				{
				case "ToKhaiInfo":
					array[i] = WinSettings.ToKhaiInfo;
					break;
				case "TTinChung":
					array[i] = WinSettings.TTinChung;
					break;
				case "PeriodInfo":
					array[i] = WinSettings.TTPeriodInfo;
					break;
				case "xmlNode":
					array[i] = xmlNode;
					break;
				case "header":
					array[i] = WinSettings.HTKKHeaderInfo;
					break;
				}
			}
			method.Invoke(WinSettings.InstanceClass, array);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void ExportHeaderXML(XmlNode xmlNode, int index)
	{
		try
		{
			MethodInfo method = GetMethod("ExportHeaderXML_" + index);
			if ((object)method == null)
			{
				return;
			}
			ParameterInfo[] parameters = method.GetParameters();
			object[] array = new object[parameters.Length];
			for (int i = 0; i < parameters.Length; i++)
			{
				switch (parameters[i].Name)
				{
				case "ToKhaiInfo":
					array[i] = WinSettings.ToKhaiInfo;
					break;
				case "TTinChung":
					array[i] = WinSettings.TTinChung;
					break;
				case "PeriodInfo":
					array[i] = WinSettings.TTPeriodInfo;
					break;
				case "xmlNode":
					array[i] = xmlNode;
					break;
				}
			}
			method.Invoke(WinSettings.InstanceClass, array);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void ExportTableF7XML(FpSpread winFpSpread, XmlNode xmlNode, int index)
	{
		try
		{
			MethodInfo method = GetMethod("ExportTableF7XML_" + index);
			if ((object)method == null)
			{
				return;
			}
			ParameterInfo[] parameters = method.GetParameters();
			object[] array = new object[parameters.Length];
			for (int i = 0; i < parameters.Length; i++)
			{
				switch (parameters[i].Name)
				{
				case "winFpSpread":
					array[i] = winFpSpread;
					break;
				case "ToKhaiInfo":
					array[i] = WinSettings.ToKhaiInfo;
					break;
				case "TTinChung":
					array[i] = WinSettings.TTinChung;
					break;
				case "PeriodInfo":
					array[i] = WinSettings.TTPeriodInfo;
					break;
				case "xmlNode":
					array[i] = xmlNode;
					break;
				}
			}
			method.Invoke(WinSettings.InstanceClass, array);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void ButtonClick(FpSpread winFpSpread, int idxRow, int idxCol)
	{
		try
		{
			MethodInfo method = GetMethod("ButtonClick_" + winFpSpread.ActiveSheetIndex);
			if ((object)method == null)
			{
				winFpSpread.ResumeLayout();
				return;
			}
			ParameterInfo[] parameters = method.GetParameters();
			object[] array = new object[parameters.Length];
			for (int i = 0; i < parameters.Length; i++)
			{
				switch (parameters[i].Name)
				{
				case "winFpSpread":
					array[i] = winFpSpread;
					break;
				case "ToKhaiInfo":
					array[i] = WinSettings.ToKhaiInfo;
					break;
				case "TTinChung":
					array[i] = WinSettings.TTinChung;
					break;
				case "PeriodInfo":
					array[i] = WinSettings.TTPeriodInfo;
					break;
				case "DicSheet":
					array[i] = WinSettings.DicSheet;
					break;
				case "HeaderInfo":
					array[i] = WinSettings.HTKKHeaderInfo;
					break;
				case "idxRow":
					array[i] = idxRow;
					break;
				case "idxCol":
					array[i] = idxCol;
					break;
				case "Obj":
					array[i] = null;
					break;
				}
			}
			method.Invoke(WinSettings.InstanceClass, array);
			if (WinSettings.MenuID == "95")
			{
				List<HTKKCMCRowValue> list = (List<HTKKCMCRowValue>)array[parameters.Length - 1];
				if (list != null && list.Count > 0)
				{
					WinSpreadProcess.InserRowTKDkNPT(winFpSpread.Sheets[0], WinSettings.ToKhaiInfo.SheetInfos[0], list);
					for (int i = 0; i < WinSettings.ToKhaiInfo.SheetInfos[0].Tables.Count; i++)
					{
						WinSpreadProcess.fpSpReindexingTableTemplate(winFpSpread.ActiveSheet, WinSettings.ToKhaiInfo.SheetInfos[0].Tables[i]);
					}
				}
			}
			else if ((WinSettings.MenuID == "01" || WinSettings.MenuID == "I5") && ConvertHelpers.ConvertToBool(array[parameters.Length - 1]))
			{
				if (winFpSpread.Sheets[1].Visible)
				{
					WinSpreadProcess.ClearData(winFpSpread, 1, isShowMess: false);
				}
				if (winFpSpread.Sheets[3].Visible)
				{
					WinSpreadProcess.ClearData(winFpSpread, 3, isShowMess: false);
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void CellChange(FpSpread winFpSpread, int sheetIndex, long iRow, long iCol)
	{
		try
		{
			MethodInfo method = GetMethod("CellChange_" + sheetIndex);
			if ((object)method == null)
			{
				return;
			}
			ParameterInfo[] parameters = method.GetParameters();
			object[] array = new object[parameters.Length];
			for (int i = 0; i < parameters.Length; i++)
			{
				switch (parameters[i].Name)
				{
				case "winFpSpread":
					array[i] = winFpSpread;
					break;
				case "ToKhaiInfo":
					array[i] = WinSettings.ToKhaiInfo;
					break;
				case "TTinChung":
					array[i] = WinSettings.TTinChung;
					break;
				case "PeriodInfo":
					array[i] = WinSettings.TTPeriodInfo;
					break;
				case "iRow":
					array[i] = iRow;
					break;
				case "iCol":
					array[i] = iCol;
					break;
				}
			}
			method.Invoke(WinSettings.InstanceClass, array);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void CellChangeBeforClickTab(FpSpread winFpSpread, int sheetIndex, long iRow, long iCol)
	{
		try
		{
			if (!(WinSettings.MenuID != "33"))
			{
				CellChange(winFpSpread, sheetIndex, iRow, iCol);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void CellEditchange(FpSpread winFpSpread, int sheetIndex, long iRow, long iCol)
	{
		try
		{
			MethodInfo method = GetMethod("CellEditchange_" + sheetIndex);
			if ((object)method == null)
			{
				return;
			}
			ParameterInfo[] parameters = method.GetParameters();
			object[] array = new object[parameters.Length];
			for (int i = 0; i < parameters.Length; i++)
			{
				switch (parameters[i].Name)
				{
				case "winFpSpread":
					array[i] = winFpSpread;
					break;
				case "ToKhaiInfo":
					array[i] = WinSettings.ToKhaiInfo;
					break;
				case "TTinChung":
					array[i] = WinSettings.TTinChung;
					break;
				case "PeriodInfo":
					array[i] = WinSettings.TTPeriodInfo;
					break;
				case "iRow":
					array[i] = iRow;
					break;
				case "iCol":
					array[i] = iCol;
					break;
				}
			}
			method.Invoke(WinSettings.InstanceClass, array);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void CellClick(FpSpread winFpSpread, int sheetIndex, long iRow, long iCol)
	{
		try
		{
			MethodInfo method = GetMethod("CellClick_" + sheetIndex);
			if ((object)method == null)
			{
				return;
			}
			ParameterInfo[] parameters = method.GetParameters();
			object[] array = new object[parameters.Length];
			for (int i = 0; i < parameters.Length; i++)
			{
				switch (parameters[i].Name)
				{
				case "winFpSpread":
					array[i] = winFpSpread;
					break;
				case "ToKhaiInfo":
					array[i] = WinSettings.ToKhaiInfo;
					break;
				case "TTinChung":
					array[i] = WinSettings.TTinChung;
					break;
				case "PeriodInfo":
					array[i] = WinSettings.TTPeriodInfo;
					break;
				case "iRow":
					array[i] = iRow;
					break;
				case "iCol":
					array[i] = iCol;
					break;
				}
			}
			method.Invoke(WinSettings.InstanceClass, array);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void LeaveCell(FpSpread winFpSpread, int sheetIndex, long iRow, long iCol, long iNewRow, long iNewCol)
	{
		try
		{
			MethodInfo method = GetMethod("LeaveCell_" + sheetIndex);
			if ((object)method == null)
			{
				return;
			}
			ParameterInfo[] parameters = method.GetParameters();
			object[] array = new object[parameters.Length];
			for (int i = 0; i < parameters.Length; i++)
			{
				switch (parameters[i].Name)
				{
				case "winFpSpread":
					array[i] = winFpSpread;
					break;
				case "ToKhaiInfo":
					array[i] = WinSettings.ToKhaiInfo;
					break;
				case "TTinChung":
					array[i] = WinSettings.TTinChung;
					break;
				case "PeriodInfo":
					array[i] = WinSettings.TTPeriodInfo;
					break;
				case "iRow":
					array[i] = iRow;
					break;
				case "iCol":
					array[i] = iCol;
					break;
				case "iNewRow":
					array[i] = iNewRow;
					break;
				case "iNewCol":
					array[i] = iNewCol;
					break;
				}
			}
			method.Invoke(WinSettings.InstanceClass, array);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void EnterCell(FpSpread winFpSpread, int sheetIndex, long iRow, long iCol)
	{
		try
		{
			MethodInfo method = GetMethod("EnterCell_" + sheetIndex);
			if ((object)method == null)
			{
				return;
			}
			ParameterInfo[] parameters = method.GetParameters();
			object[] array = new object[parameters.Length];
			for (int i = 0; i < parameters.Length; i++)
			{
				switch (parameters[i].Name)
				{
				case "winFpSpread":
					array[i] = winFpSpread;
					break;
				case "ToKhaiInfo":
					array[i] = WinSettings.ToKhaiInfo;
					break;
				case "TTinChung":
					array[i] = WinSettings.TTinChung;
					break;
				case "PeriodInfo":
					array[i] = WinSettings.TTPeriodInfo;
					break;
				case "iRow":
					array[i] = iRow;
					break;
				case "iCol":
					array[i] = iCol;
					break;
				}
			}
			method.Invoke(WinSettings.InstanceClass, array);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static bool SetActivePrintSheet(FpSpread winFpSpread, int sheetIndex)
	{
		try
		{
			MethodInfo method = GetMethod("SetActivePrintSheet_" + sheetIndex);
			if ((object)method == null)
			{
				return true;
			}
			ParameterInfo[] parameters = method.GetParameters();
			object[] array = new object[parameters.Length];
			for (int i = 0; i < parameters.Length; i++)
			{
				switch (parameters[i].Name)
				{
				case "winFpSpread":
					array[i] = winFpSpread;
					break;
				case "ToKhaiInfo":
					array[i] = WinSettings.ToKhaiInfo;
					break;
				case "TTinChung":
					array[i] = WinSettings.TTinChung;
					break;
				case "PeriodInfo":
					array[i] = WinSettings.TTPeriodInfo;
					break;
				case "output":
					array[i] = null;
					break;
				}
			}
			method.Invoke(WinSettings.InstanceClass, array);
			return ConvertHelpers.ConvertToBool(array[^1]);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return false;
		}
	}

	public static void KeyPress(FpSpread winFpSpread, int idxSheet, int idxRow, int idxCol, KeyPressEventArgs e, Keys keyData = (Keys)0)
	{
		//IL_0179: Unknown result type (might be due to invalid IL or missing references)
		try
		{
			MethodInfo method = GetMethod("KeyPress");
			if ((object)method == null)
			{
				return;
			}
			ParameterInfo[] parameters = method.GetParameters();
			object[] array = new object[parameters.Length];
			for (int i = 0; i < parameters.Length; i++)
			{
				switch (parameters[i].Name)
				{
				case "winFpSpread":
					array[i] = winFpSpread;
					break;
				case "ToKhaiInfo":
					array[i] = WinSettings.ToKhaiInfo;
					break;
				case "TTinChung":
					array[i] = WinSettings.TTinChung;
					break;
				case "PeriodInfo":
					array[i] = WinSettings.TTPeriodInfo;
					break;
				case "HeaderInfo":
					array[i] = WinSettings.HTKKHeaderInfo;
					break;
				case "idxSheet":
					array[i] = idxSheet;
					break;
				case "idxRow":
					array[i] = idxRow;
					break;
				case "idxCol":
					array[i] = idxCol;
					break;
				case "e":
					array[i] = e;
					break;
				case "keyData":
					array[i] = keyData;
					break;
				}
			}
			method.Invoke(WinSettings.InstanceClass, array);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void AutoBindData(FpSpread winFpSpread, int sheetIndex)
	{
		try
		{
			MethodInfo method = GetMethod("AutoBindData_" + sheetIndex);
			if ((object)method == null)
			{
				return;
			}
			ParameterInfo[] parameters = method.GetParameters();
			object[] array = new object[parameters.Length];
			for (int i = 0; i < parameters.Length; i++)
			{
				switch (parameters[i].Name)
				{
				case "winFpSpread":
					array[i] = winFpSpread;
					break;
				case "ToKhaiInfo":
					array[i] = WinSettings.ToKhaiInfo;
					break;
				case "TTinChung":
					array[i] = WinSettings.TTinChung;
					break;
				case "PeriodInfo":
					array[i] = WinSettings.TTPeriodInfo;
					break;
				case "HeaderInfo":
					array[i] = WinSettings.HTKKHeaderInfo;
					break;
				case "DicSheet":
					array[i] = WinSettings.DicSheet;
					break;
				case "sView":
					array[i] = winFpSpread.Sheets[sheetIndex];
					break;
				}
			}
			method.Invoke(WinSettings.InstanceClass, array);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void PreBindData(FpSpread winFpSpread)
	{
		try
		{
			MethodInfo method = GetMethod("PreBindData");
			if ((object)method == null)
			{
				return;
			}
			ParameterInfo[] parameters = method.GetParameters();
			object[] array = new object[parameters.Length];
			for (int i = 0; i < parameters.Length; i++)
			{
				switch (parameters[i].Name)
				{
				case "winFpSpread":
					array[i] = winFpSpread;
					break;
				case "ToKhaiInfo":
					array[i] = WinSettings.ToKhaiInfo;
					break;
				case "TTinChung":
					array[i] = WinSettings.TTinChung;
					break;
				case "PeriodInfo":
					array[i] = WinSettings.TTPeriodInfo;
					break;
				case "HeaderInfo":
					array[i] = WinSettings.HTKKHeaderInfo;
					break;
				}
			}
			method.Invoke(WinSettings.InstanceClass, array);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void UpdateFormula(FpSpread winFpSpread, int sheetIndex)
	{
		try
		{
			if ((object)GetMethod("UpdateFormula_" + sheetIndex) == null)
			{
				return;
			}
			MethodInfo method = GetMethod("UpdateFormula_" + sheetIndex);
			ParameterInfo[] parameters = method.GetParameters();
			object[] array = new object[parameters.Length];
			for (int i = 0; i < parameters.Length; i++)
			{
				switch (parameters[i].Name)
				{
				case "winFpSpread":
					array[i] = winFpSpread;
					break;
				case "ToKhaiInfo":
					array[i] = WinSettings.ToKhaiInfo;
					break;
				case "TTinChung":
					array[i] = WinSettings.TTinChung;
					break;
				case "PeriodInfo":
					array[i] = WinSettings.TTPeriodInfo;
					break;
				case "KHBS":
					array[i] = WinSettings.KHBS;
					break;
				}
			}
			method.Invoke(WinSettings.InstanceClass, array);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static MethodInfo GetMethod(string strMethod)
	{
		try
		{
			Type type = WinSettings.InstanceClass.GetType();
			object obj = Activator.CreateInstance(type);
			return type.GetMethod(strMethod);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public static MethodInfo[] GetMethods()
	{
		try
		{
			Type type = WinSettings.InstanceClass.GetType();
			object obj = Activator.CreateInstance(type);
			return type.GetMethods();
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public static object GetInstance(string strClassName)
	{
		try
		{
			strClassName = strClassName.Replace("-", "");
			strClassName = strClassName.Replace("_", "");
			strClassName = strClassName.Replace(".", "");
			Type type = Type.GetType("TAX.HTKK.Business." + strClassName + ",TAX.HTKK.Business");
			return Activator.CreateInstance(type);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public static string GetDefaultYear()
	{
		object obj;
		try
		{
			MethodInfo method = GetMethod("GetDefaultYear");
			obj = method.Invoke(WinSettings.InstanceClass, null).ToString();
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			obj = null;
		}
		return obj.ToString();
	}

	public static bool ValidatePeriod_mskYear()
	{
		bool flag = false;
		try
		{
			MethodInfo method = GetMethod("ValidatePeriod_mskYear");
			if ((object)method == null)
			{
				return true;
			}
			flag = (bool)method.Invoke(parameters: new object[1] { WinSettings.TTPeriodInfo }, obj: WinSettings.InstanceClass);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			flag = false;
		}
		return flag;
	}

	public static bool ValidatePeriod_axfpsPhuLuc(FpSpread winFpSpread)
	{
		bool flag = false;
		try
		{
			MethodInfo method = GetMethod("ValidatePeriod_axfpsPhuLuc");
			if ((object)method == null)
			{
				return true;
			}
			ParameterInfo[] parameters = method.GetParameters();
			object[] array = new object[parameters.Length];
			for (int i = 0; i < parameters.Length; i++)
			{
				switch (parameters[i].Name)
				{
				case "winFpSpread":
					array[i] = winFpSpread;
					break;
				case "TTinChung":
					array[i] = WinSettings.TTinChung;
					break;
				case "PeriodInfo":
					array[i] = WinSettings.TTPeriodInfo;
					break;
				case "HeaderInfo":
					array[i] = WinSettings.HTKKHeaderInfo;
					break;
				}
			}
			flag = (bool)method.Invoke(WinSettings.InstanceClass, array);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			flag = false;
		}
		return flag;
	}

	public static bool ValidatePeriod_ALL()
	{
		bool flag = false;
		try
		{
			MethodInfo method = GetMethod("ValidatePeriod_ALL");
			if ((object)method == null)
			{
				return true;
			}
			ParameterInfo[] parameters = method.GetParameters();
			object[] array = new object[parameters.Length];
			for (int i = 0; i < parameters.Length; i++)
			{
				switch (parameters[i].Name)
				{
				case "DicSheet":
					array[i] = WinSettings.DicSheet;
					break;
				case "TTinChung":
					array[i] = WinSettings.TTinChung;
					break;
				case "PeriodInfo":
					array[i] = WinSettings.TTPeriodInfo;
					break;
				case "HeaderInfo":
					array[i] = WinSettings.HTKKHeaderInfo;
					break;
				}
			}
			flag = (bool)method.Invoke(WinSettings.InstanceClass, array);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			flag = false;
		}
		return flag;
	}

	public static bool CheckThangQuyTuNgayDenNgay(string fullNameDataFileToKhaiChinh, string strFileNameExtend, ref string returnMess)
	{
		try
		{
			bool flag = false;
			MethodInfo method = GetMethod("CheckThangQuyTuNgayDenNgay");
			if ((object)method != null)
			{
				ParameterInfo[] parameters = method.GetParameters();
				object[] array = new object[parameters.Length];
				int num = -1;
				for (int i = 0; i < parameters.Length; i++)
				{
					switch (parameters[i].Name)
					{
					case "fullNameDataFileToKhaiChinh":
						array[i] = fullNameDataFileToKhaiChinh;
						break;
					case "strFileNameExtend":
						array[i] = strFileNameExtend;
						break;
					case "periodInfo":
						array[i] = WinSettings.TTPeriodInfo;
						break;
					case "errorMess":
						array[i] = returnMess;
						num = i;
						break;
					}
				}
				flag = (bool)method.Invoke(WinSettings.InstanceClass, array);
				if (num != -1)
				{
					returnMess = array[num].ToString();
				}
				return flag;
			}
			return true;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return false;
		}
	}

	public static bool CheckDuplicateMonInYear(string fileName, string afterName)
	{
		bool flag = false;
		try
		{
			MethodInfo method = GetMethod("CheckDuplicateMonInYear");
			if ((object)method != null)
			{
				return (bool)method.Invoke(parameters: new object[2] { fileName, afterName }, obj: WinSettings.InstanceClass);
			}
			return true;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return false;
		}
	}

	public static void WinDoWorkDelete(FpSpread winFpSpread, Dictionary<string, int> listSheet, int activeSheet)
	{
		try
		{
			if (activeSheet == 0 || WinSettings.MenuID == "33" || WinSettings.MenuID == "B9" || WinSettings.MenuID == "B8" || WinSettings.MenuID == "37" || WinSettings.MenuID == "P2" || WinSettings.MenuID.Equals("W1") || WinSettings.MenuID == "N7" || ((WinSettings.MenuID == "T8" || WinSettings.MenuID == "R5") && activeSheet == 3) || (WinSettings.TTPeriodInfo.LoaiToKhai == "B" && !"30,I5,K9,K1,M7,K6,K8,73,55,71,31,P2,X6".Contains(WinSettings.MenuID) && WinSettings.ToKhaiInfo.KieuKHBS != "3") || winFpSpread.ActiveSheet.SheetName == "KHBS" || winFpSpread.ActiveSheet.SheetName == "01-KHBS" || winFpSpread.ActiveSheet.SheetName == "01-1/KHBS")
			{
				for (int i = 0; i < WinSettings.ToKhaiInfo.SheetInfos.Count; i++)
				{
					if (WinSettings.ToKhaiInfo.SheetInfos[i] != null)
					{
						string fileNameBySheet = GetFileNameBySheet(WinSettings.TTPeriodInfo, WinSettings.ToKhaiInfo.SheetInfos[i].DataFile, HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst + "\\", WinSettings.ToKhaiInfo.VersionXML);
						if (File.Exists(fileNameBySheet))
						{
							HTKKFile.DeleteFile(fileNameBySheet);
							WinSettings.ToKhaiInfo.SheetInfos[i] = null;
						}
					}
				}
			}
			else if (WinSettings.MenuID == "G9")
			{
				WinSettings.DicSheet.Remove(WinSettings.ToKhaiInfo.SheetInfos[activeSheet].SheetName);
				string fileNameBySheet = GetFileNameBySheet(WinSettings.TTPeriodInfo, WinSettings.ToKhaiInfo.SheetInfos[activeSheet].DataFile, HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst + "\\", WinSettings.ToKhaiInfo.VersionXML);
				HTKKCMCSheetInfo hTKKCMCSheetInfo = WinSettings.ToKhaiInfo.SheetInfos[activeSheet];
				WinSpreadProcess.ResetColorSheetDefault(winFpSpread, WinSettings.ToKhaiInfo, activeSheet);
				List<HTKKCMCTableInfo> tablesF7 = hTKKCMCSheetInfo.Tables.Where((HTKKCMCTableInfo x) => x.Header != null && x.Header.StaticCells.Count > 0 && x.TableID == "1").ToList();
				int i2;
				for (i2 = tablesF7.Count - 1; i2 >= 0; i2--)
				{
					HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables.FirstOrDefault((HTKKCMCTableInfo x) => x.TableID == tablesF7[i2].TableID && x.TableName == tablesF7[i2].TableName);
					GeneralFunctions.DeleteTableGroup(winFpSpread.ActiveSheet, hTKKCMCSheetInfo, hTKKCMCTableInfo.TableName);
				}
				List<HTKKCMCTableInfo> tablesF8 = hTKKCMCSheetInfo.Tables.Where((HTKKCMCTableInfo x) => x.Header == null || x.Header.StaticCells.Count == 0).ToList();
				int i3;
				for (i3 = tablesF8.Count - 1; i3 >= 0; i3--)
				{
					int num = hTKKCMCSheetInfo.Tables.FindIndex((HTKKCMCTableInfo x) => x.TableID == tablesF8[i3].TableID && x.TableName == tablesF8[i3].TableName);
					int dataLocationTemplateRowStart = hTKKCMCSheetInfo.Tables[num].Location.DataLocationTemplateRowStart;
					if (hTKKCMCSheetInfo.Tables[num].RowsValues.Count > 1)
					{
						GeneralFunctions.RemoveDataRows(winFpSpread.ActiveSheet, num, dataLocationTemplateRowStart - 1, hTKKCMCSheetInfo.Tables[num].RowsValues.Count - 1, hTKKCMCSheetInfo);
					}
					foreach (HTKKCMCTableCellInfo cell in hTKKCMCSheetInfo.Tables[num].RowsInfo.Cells)
					{
						winFpSpread.ActiveSheet.Cells[dataLocationTemplateRowStart - 1, cell.CellTemplate.ColIndex - 1].Value = cell.DefaultValue;
					}
				}
				WinSettings.ToKhaiInfo.SheetInfos[activeSheet] = null;
				RefreshSheet(winFpSpread, activeSheet, 2);
				if (File.Exists(fileNameBySheet))
				{
					HTKKFile.DeleteFile(fileNameBySheet);
				}
			}
			else
			{
				WinSettings.DicSheet.Remove(WinSettings.ToKhaiInfo.SheetInfos[activeSheet].SheetName);
				string fileNameBySheet = GetFileNameBySheet(WinSettings.TTPeriodInfo, WinSettings.ToKhaiInfo.SheetInfos[activeSheet].DataFile, HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst + "\\", WinSettings.ToKhaiInfo.VersionXML);
				HTKKCMCSheetInfo hTKKCMCSheetInfo = WinSettings.ToKhaiInfo.SheetInfos[activeSheet];
				WinSpreadProcess.ResetColorSheetDefault(winFpSpread, WinSettings.ToKhaiInfo, activeSheet);
				for (int i = 0; i < hTKKCMCSheetInfo.Tables.Count; i++)
				{
					HTKKCMCTableInfo hTKKCMCTableInfo2 = hTKKCMCSheetInfo.Tables[i];
					if (i > 0 && hTKKCMCSheetInfo.Tables[i].TableName == hTKKCMCSheetInfo.Tables[i - 1].TableName)
					{
						WinSpreadProcess.RemoveTable(winFpSpread.ActiveSheet, i, hTKKCMCSheetInfo);
						i--;
						continue;
					}
					WinSpreadProcess.RemoveDataRows(winFpSpread.ActiveSheet, i, hTKKCMCSheetInfo.Tables[i].Location.DataLocationTemplateRowStart, hTKKCMCTableInfo2.RowsValues.Count - 1, hTKKCMCSheetInfo);
					for (int num2 = 0; num2 < hTKKCMCSheetInfo.Tables[i].RowsValues.Count; num2++)
					{
						for (int num3 = 0; num3 < hTKKCMCSheetInfo.Tables[i].RowsInfo.Cells.Count; num3++)
						{
							WinSpreadProcess.fpSpBindDataCell(winFpSpread.ActiveSheet, hTKKCMCSheetInfo.Tables[i].Location.DataLocationTemplateRowStart + num2, hTKKCMCSheetInfo.Tables[i].RowsInfo.Cells[num3].CellTemplate.ColIndex, hTKKCMCSheetInfo.Tables[i].RowsInfo.Cells[num3].Controltype, hTKKCMCSheetInfo.Tables[i].RowsInfo.Cells[num3].DefaultValue);
						}
					}
				}
				WinSettings.ToKhaiInfo.SheetInfos[activeSheet] = null;
				RefreshSheet(winFpSpread, activeSheet, 2);
				if (File.Exists(fileNameBySheet))
				{
					HTKKFile.DeleteFile(fileNameBySheet);
				}
			}
			if (activeSheet == 1 && WinSettings.MenuID == "73")
			{
				winFpSpread.Sheets[0].Cells[72, HTKKFile.ColumnLettersToIndex("AW") - 1].Formula = "BF72";
			}
			if (activeSheet == 2 && WinSettings.MenuID == "73")
			{
				winFpSpread.Sheets[0].Cells[72, HTKKFile.ColumnLettersToIndex("AW") - 1].Formula = "BF72";
			}
			if (activeSheet == 1 && WinSettings.MenuID == "55")
			{
				int num4 = WinSettings.ToKhaiInfo.SheetInfos[0].Tables[0].RowsValues.Count - 1;
				HTKKCMCSheetInfo hTKKCMCSheetInfo = WinSettings.ToKhaiInfo.SheetInfos[0];
				HTKKCMCTableInfo hTKKCMCTableInfo3 = hTKKCMCSheetInfo.Tables[0];
				int num5 = hTKKCMCTableInfo3.Location.DataLocationTemplateRowStart - 1;
				for (int i = 0; i < hTKKCMCTableInfo3.RowsValues.Count; i++)
				{
					winFpSpread.Sheets[0].Cells[num5 + i, HTKKFile.ColumnLettersToIndex("CF") - 1].Formula = "ROUND(Y" + (num5 + i + 1) + "+AX" + (num5 + i + 1) + "+BW" + (num5 + i + 1) + ",0)";
				}
			}
			if (activeSheet == 2 && WinSettings.MenuID == "55")
			{
				int num4 = WinSettings.ToKhaiInfo.SheetInfos[0].Tables[0].RowsValues.Count - 1;
				HTKKCMCSheetInfo hTKKCMCSheetInfo = WinSettings.ToKhaiInfo.SheetInfos[0];
				HTKKCMCTableInfo hTKKCMCTableInfo3 = hTKKCMCSheetInfo.Tables[0];
				int num5 = hTKKCMCTableInfo3.Location.DataLocationTemplateRowStart - 1;
				for (int i = 0; i < hTKKCMCTableInfo3.RowsValues.Count; i++)
				{
					winFpSpread.Sheets[0].Cells[num5 + i, HTKKFile.ColumnLettersToIndex("CF") - 1].Formula = "ROUND(Y" + (num5 + i + 1) + "+AX" + (num5 + i + 1) + "+BW" + (num5 + i + 1) + ",0)";
				}
			}
			if (activeSheet == 1 && WinSettings.MenuID == "T5")
			{
				int num4 = WinSettings.ToKhaiInfo.SheetInfos[0].Tables[0].RowsValues.Count - 1;
				HTKKCMCSheetInfo hTKKCMCSheetInfo = WinSettings.ToKhaiInfo.SheetInfos[0];
				HTKKCMCTableInfo hTKKCMCTableInfo3 = hTKKCMCSheetInfo.Tables[0];
				int num5 = hTKKCMCTableInfo3.Location.DataLocationTemplateRowStart - 1;
				for (int i = 0; i < hTKKCMCTableInfo3.RowsValues.Count; i++)
				{
				}
			}
			if (activeSheet == 1 && WinSettings.MenuID == "29")
			{
				winFpSpread.Sheets[0].Cells[56, HTKKFile.ColumnLettersToIndex("O") - 1].Locked = true;
				winFpSpread.Sheets[0].Cells[57, HTKKFile.ColumnLettersToIndex("O") - 1].Locked = true;
				winFpSpread.Sheets[0].Cells[58, HTKKFile.ColumnLettersToIndex("O") - 1].Locked = true;
				winFpSpread.Sheets[0].Cells[56, HTKKFile.ColumnLettersToIndex("O") - 1].Formula = "ROUND(K57*0.05,0)";
				winFpSpread.Sheets[0].Cells[57, HTKKFile.ColumnLettersToIndex("O") - 1].Formula = "ROUND(K58*0.05,0)";
				winFpSpread.Sheets[0].Cells[58, HTKKFile.ColumnLettersToIndex("O") - 1].Formula = "ROUND(K59*0.05,0)";
			}
			if (activeSheet == 1 && WinSettings.MenuID == "71")
			{
				winFpSpread.Sheets[0].Cells[42, HTKKFile.ColumnLettersToIndex("Z") - 1].Locked = true;
				winFpSpread.Sheets[0].Cells[43, HTKKFile.ColumnLettersToIndex("Z") - 1].Locked = true;
				winFpSpread.Sheets[0].Cells[44, HTKKFile.ColumnLettersToIndex("Z") - 1].Locked = true;
				winFpSpread.Sheets[0].Cells[45, HTKKFile.ColumnLettersToIndex("Z") - 1].Locked = true;
				winFpSpread.Sheets[0].Cells[42, HTKKFile.ColumnLettersToIndex("Z") - 1].Formula = "ROUND(Q43*1/100,0)";
				winFpSpread.Sheets[0].Cells[43, HTKKFile.ColumnLettersToIndex("Z") - 1].Formula = "ROUND(Q44*5/100,0)";
				winFpSpread.Sheets[0].Cells[44, HTKKFile.ColumnLettersToIndex("Z") - 1].Formula = "ROUND(Q45*3/100,0)";
				winFpSpread.Sheets[0].Cells[45, HTKKFile.ColumnLettersToIndex("Z") - 1].Formula = "ROUND(Q46*2/100,0)";
			}
			HTKKCMCSheetInfo hTKKCMCSheetInfo2;
			if (activeSheet == 1 && WinSettings.MenuID == "38")
			{
				SheetView sheetView = winFpSpread.Sheets[0];
				hTKKCMCSheetInfo2 = WinSettings.ToKhaiInfo.SheetInfos[0];
				HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo2.Tables[0];
				int num6 = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1;
				int num7 = num6 + hTKKCMCTableInfo.RowsValues.Count - 1;
				string formula = "";
				for (int i = num7 + 10; i < num7 + 14; i++)
				{
					if (i == num7 + 10)
					{
						formula = "ROUND(K" + i + "*0.005,0)";
					}
					if (i == num7 + 11)
					{
						formula = "ROUND(K" + i + "*0.02,0)";
					}
					if (i == num7 + 12)
					{
						formula = "ROUND(K" + i + "*0.015,0)";
					}
					if (i == num7 + 13)
					{
						formula = "ROUND(K" + i + "*0.01,0)";
					}
					sheetView.Cells[i - 1, HTKKFile.ColumnLettersToIndex("R") - 1].Formula = formula;
				}
			}
			if (activeSheet != 2 || !(WinSettings.MenuID == "28"))
			{
				return;
			}
			SheetView sheetView2 = winFpSpread.Sheets[0];
			hTKKCMCSheetInfo2 = WinSettings.ToKhaiInfo.SheetInfos[0];
			sheetView2.Cells[91, HTKKFile.ColumnLettersToIndex("P") - 1].Locked = true;
			sheetView2.Cells[93, HTKKFile.ColumnLettersToIndex("P") - 1].Locked = true;
			sheetView2.Cells[95, HTKKFile.ColumnLettersToIndex("P") - 1].Locked = true;
			sheetView2.Cells[97, HTKKFile.ColumnLettersToIndex("P") - 1].Locked = true;
			sheetView2.Cells[91, HTKKFile.ColumnLettersToIndex("X") - 1].Locked = true;
			sheetView2.Cells[93, HTKKFile.ColumnLettersToIndex("X") - 1].Locked = true;
			sheetView2.Cells[95, HTKKFile.ColumnLettersToIndex("X") - 1].Locked = true;
			sheetView2.Cells[97, HTKKFile.ColumnLettersToIndex("X") - 1].Locked = true;
			for (int num8 = 1; num8 < hTKKCMCSheetInfo2.Tables.Count - 1; num8++)
			{
				HTKKCMCTableInfo hTKKCMCTableInfo4 = hTKKCMCSheetInfo2.Tables[num8];
				int num6 = hTKKCMCTableInfo4.Location.DataLocationTemplateRowStart - 1;
				string[] array = new string[hTKKCMCTableInfo4.RowsInfo.Cells.Count];
				for (int num2 = 0; num2 < hTKKCMCTableInfo4.RowsValues.Count; num2++)
				{
					sheetView2.Cells[num6 + num2, HTKKFile.ColumnLettersToIndex("W") - 1].Locked = true;
				}
			}
			int num9 = HTKKFile.ColumnLettersToIndex("L");
			int num10 = HTKKFile.ColumnLettersToIndex("T");
			sheetView2.Cells[91, HTKKFile.ColumnLettersToIndex("P") - 1].Formula = "ROUND(L92*0.01,0)";
			sheetView2.Cells[93, HTKKFile.ColumnLettersToIndex("P") - 1].Formula = "ROUND(L94*0.05,0)";
			sheetView2.Cells[95, HTKKFile.ColumnLettersToIndex("P") - 1].Formula = "ROUND(L96*0.03,0)";
			sheetView2.Cells[97, HTKKFile.ColumnLettersToIndex("P") - 1].Formula = "ROUND(L98*0.02,0)";
			sheetView2.Cells[91, HTKKFile.ColumnLettersToIndex("X") - 1].Formula = "ROUND(T92*0.005,0)";
			sheetView2.Cells[93, HTKKFile.ColumnLettersToIndex("X") - 1].Formula = "ROUND(T94*0.02,0)";
			sheetView2.Cells[95, HTKKFile.ColumnLettersToIndex("X") - 1].Formula = "ROUND(T96*0.015,0)";
			sheetView2.Cells[97, HTKKFile.ColumnLettersToIndex("X") - 1].Formula = "ROUND(T98*0.01,0)";
			for (int num8 = 1; num8 < hTKKCMCSheetInfo2.Tables.Count; num8++)
			{
				HTKKCMCTableInfo hTKKCMCTableInfo4 = hTKKCMCSheetInfo2.Tables[num8];
				int num6 = hTKKCMCTableInfo4.Location.DataLocationTemplateRowStart - 1;
				int num7 = num6 + hTKKCMCTableInfo4.RowsValues.Count - 1;
				List<HTKKCMCRowValue> rowsValues = hTKKCMCTableInfo4.RowsValues;
				for (int i = 0; i < rowsValues.Count; i++)
				{
					if (num8 == 1)
					{
						sheetView2.Cells[num6 + i, HTKKFile.ColumnLettersToIndex("W") - 1].Formula = "ROUND(P" + (num6 + i + 1) + "* U" + (num6 + i + 1) + " / 100, 0)";
					}
					if (num8 == 2)
					{
						sheetView2.Cells[num6 + i, HTKKFile.ColumnLettersToIndex("W") - 1].Formula = "ROUND(M" + (num6 + i + 1) + "*Q" + (num6 + i + 1) + "*U" + (num6 + i + 1) + "/100,0)";
					}
					if (num8 == 3)
					{
						sheetView2.Cells[num6 + i, HTKKFile.ColumnLettersToIndex("W") - 1].Formula = "ROUND(M" + (num6 + i + 1) + "*Q" + (num6 + i + 1) + ",0)";
					}
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void RefreshSheet(FpSpread winFpSpread, int activeSheet, int actionType)
	{
		try
		{
			MethodInfo method = GetMethod("Refresh_" + activeSheet);
			if ((object)method == null)
			{
				return;
			}
			ParameterInfo[] parameters = method.GetParameters();
			object[] array = new object[parameters.Length];
			for (int i = 0; i < parameters.Length; i++)
			{
				switch (parameters[i].Name)
				{
				case "winFpSpread":
					array[i] = winFpSpread;
					break;
				case "actionType":
					array[i] = actionType;
					break;
				case "ToKhaiInfo":
					array[i] = WinSettings.ToKhaiInfo;
					break;
				case "PeriodInfo":
					array[i] = WinSettings.TTPeriodInfo;
					break;
				}
			}
			method.Invoke(WinSettings.InstanceClass, array);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static List<item> getItemsDm(string loaiDm, Dictionary<string, object> dicParam = null)
	{
		List<item> result = null;
		try
		{
			MethodInfo method = GetMethod("GetItemDm_" + loaiDm);
			if ((object)method == null)
			{
				return null;
			}
			ParameterInfo[] parameters = method.GetParameters();
			object[] array = new object[parameters.Length];
			for (int i = 0; i < parameters.Length; i++)
			{
				string name = parameters[i].Name;
				if (name != null && name == "dicParam")
				{
					array[i] = dicParam;
				}
			}
			result = (List<item>)method.Invoke(WinSettings.InstanceClass, array);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return result;
	}

	public static void SheetTabClick(FpSpread winFpSpread, int activeSheet)
	{
		try
		{
			MethodInfo method = GetMethod("SheetTabClick");
			if ((object)method == null)
			{
				return;
			}
			ParameterInfo[] parameters = method.GetParameters();
			object[] array = new object[parameters.Length];
			for (int i = 0; i < parameters.Length; i++)
			{
				switch (parameters[i].Name)
				{
				case "winFpSpread":
					array[i] = winFpSpread;
					break;
				case "ToKhaiInfo":
					array[i] = WinSettings.ToKhaiInfo;
					break;
				}
			}
			method.Invoke(WinSettings.InstanceClass, array);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void ClearData(FpSpread winFpSpread, int activeSheet)
	{
		try
		{
			MethodInfo method = GetMethod("ClearData_" + activeSheet);
			if ((object)method == null)
			{
				return;
			}
			ParameterInfo[] parameters = method.GetParameters();
			object[] array = new object[parameters.Length];
			for (int i = 0; i < parameters.Length; i++)
			{
				switch (parameters[i].Name)
				{
				case "winFpSpread":
					array[i] = winFpSpread;
					break;
				case "ToKhaiInfo":
					array[i] = WinSettings.ToKhaiInfo;
					break;
				}
			}
			method.Invoke(WinSettings.InstanceClass, array);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void ProcessXMLNode(HTKKCMCSheetInfo sInfo, XmlNode sheetNode, bool checkKHBS)
	{
		//IL_0aae: Unknown result type (might be due to invalid IL or missing references)
		//IL_0ab5: Expected O, but got Unknown
		try
		{
			int num = 0;
			for (int i = 0; i < sInfo.StaticCells.Count; i++)
			{
				if (sInfo.StaticCells[i].Path == null)
				{
					continue;
				}
				if (sInfo.StaticCells[i].Path == "PL_01_ND20/Muc_IV/IV_3/Dong_1a/ct06")
				{
					num = i;
				}
				if (sInfo.StaticCells[i].Path.Length <= 0 || sInfo.StaticCells[i].CellTemplate == null)
				{
					continue;
				}
				if (sInfo.StaticCells[i].Path.IndexOf("giamDoc") > 0)
				{
					num = 0;
				}
				if ((sInfo.StaticCells[i].Path.IndexOf("CDKT_HoatDongKhongLienTuc") >= 0 && WinSettings.TTPeriodInfo.DoanhNghiepHoatDongLienTuc == "LT") || (sInfo.StaticCells[i].Path.IndexOf("CDKT_HoatDongLienTuc") >= 0 && WinSettings.TTPeriodInfo.DoanhNghiepHoatDongLienTuc == "KL"))
				{
					continue;
				}
				XmlNode val = sheetNode.SelectSingleNode(sInfo.StaticCells[i].Path.Remove(0, sInfo.StaticCells[i].Path.IndexOf('/') + 1));
				if (val != null)
				{
					if (sInfo.StaticCells[i].Controltype != null)
					{
						ReChangeValue(val, sInfo.StaticCells[i].Controltype);
					}
					sInfo.StaticCells[i].Value = val.InnerXml;
				}
				if (checkKHBS && HTKKFile.ColumnIndexToLetters(sInfo.StaticCells[i].CellTemplate.ColIndex) + sInfo.StaticCells[i].CellTemplate.RowIndex == "BI5" && sInfo.StaticCells[i].CellTemplate.RowIndex == 5)
				{
					sInfo.StaticCells[i].Value = sInfo.StaticCells[i - 1].Value;
				}
			}
			int count = sInfo.Tables.Count;
			for (int i = count - 1; i >= 0; i--)
			{
				if (sInfo.Tables[i].TableID != "1")
				{
					sInfo.Tables.Remove(sInfo.Tables[i]);
				}
			}
			List<HTKKCMCTableInfo> list = new List<HTKKCMCTableInfo>();
			count = sInfo.Tables.Count;
			int num2 = 0;
			int num3 = 0;
			bool flag = false;
			for (int i = 0; i < count; i++)
			{
				num2 = 0;
				num3 = 0;
				flag = false;
				sInfo.Tables[i].RowsValues.Clear();
				XmlNodeList val2 = sheetNode.SelectNodes(sInfo.Tables[i].TablePath.Remove(0, sInfo.Tables[i].TablePath.IndexOf('/') + 1));
				int num4 = 0;
				num4 = ((sInfo.Tables[i].RowsInfo.RowPath == null) ? 1 : val2.Count);
				if (num4 > 1)
				{
					flag = true;
				}
				int num5 = 0;
				for (int j = 0; j < num4; j++)
				{
					HTKKCMCTableInfo hTKKCMCTableInfo = new HTKKCMCTableInfo(sInfo.Tables[i], 2);
					hTKKCMCTableInfo.RowsValues.Clear();
					hTKKCMCTableInfo.TableID = (j + 1).ToString();
					if (hTKKCMCTableInfo.CellTableSTT != null)
					{
						hTKKCMCTableInfo.CellTableSTT.Value = (j + 1).ToString();
					}
					for (int k = 0; k < hTKKCMCTableInfo.Header.StaticCells.Count; k++)
					{
						if (hTKKCMCTableInfo.Header.StaticCells[k].Path != null)
						{
							if (hTKKCMCTableInfo.Header.StaticCells[k].Path.Length <= 0 || hTKKCMCTableInfo.Header.StaticCells[k].CellTemplate == null)
							{
								continue;
							}
							string path = hTKKCMCTableInfo.Header.StaticCells[k].Path;
							XmlNode val = val2[j].SelectSingleNode(path);
							if (val != null)
							{
								if (hTKKCMCTableInfo.Header.StaticCells[k].Controltype != null)
								{
									ReChangeValue(val, hTKKCMCTableInfo.Header.StaticCells[k].Controltype);
								}
								hTKKCMCTableInfo.Header.StaticCells[k].Value = val.InnerXml;
							}
							else if (hTKKCMCTableInfo.Header.StaticCells[k].CellTemplate.ColIndex == hTKKCMCTableInfo.iColumnSTT)
							{
								hTKKCMCTableInfo.Header.StaticCells[k].Value = (j + 1).ToString();
								hTKKCMCTableInfo.CellTableSTT.Value = (j + 1).ToString();
							}
							else
							{
								if (!(WinSettings.MenuID == "N8") || !(sInfo.SheetID == "03-3D"))
								{
									continue;
								}
								string text = "congNghe_sanPham/ten";
								if (!path.Contains(text))
								{
									continue;
								}
								XmlNode val3 = val2[j].SelectNodes(text)[ConvertHelpers.ConvertToInt(path.Substring(path.Length - 1, 1))];
								if (val3 != null)
								{
									if (hTKKCMCTableInfo.Header.StaticCells[k].Controltype != null)
									{
										ReChangeValue(val, hTKKCMCTableInfo.Header.StaticCells[k].Controltype);
									}
									hTKKCMCTableInfo.Header.StaticCells[k].Value = val3.InnerXml;
								}
							}
						}
						else if (hTKKCMCTableInfo.Header.StaticCells[k].CellTemplate.ColIndex == hTKKCMCTableInfo.iColumnSTT)
						{
							hTKKCMCTableInfo.Header.StaticCells[k].Value = (j + 1).ToString();
						}
					}
					XmlNodeList val4 = null;
					if (sInfo.Tables[i].RowsInfo.RowPath != null)
					{
						val4 = ((sInfo.Tables[i].RowsInfo.RowPath.Length <= 0) ? null : val2[j].SelectNodes(hTKKCMCTableInfo.RowsInfo.RowPath));
					}
					else if (val2[j] != null)
					{
						if (WinSettings.MenuID == "B2" || WinSettings.MenuID == "O9")
						{
							val4 = val2[j].ParentNode.SelectNodes("DSachCNDangkyMST");
						}
						else if (WinSettings.MenuID == "M6" || WinSettings.MenuID == "O8" || "~T9~T8~L4~L3~H2~N3~D1~B9~N8~N7~01~03~94~96~97~J2~99~Q9~R7~Y7~Z5~W1~".IndexOf("~" + WinSettings.MenuID + "~") > -1)
						{
							if (sInfo.Tables[i].TablePath != null)
							{
								string[] array = sInfo.Tables[i].TablePath.Split(new char[1] { '/' });
								val4 = val2[j].ParentNode.SelectNodes(array[^1]);
							}
							else
							{
								val4 = val2[j].ParentNode.ChildNodes;
							}
						}
						else if (sInfo.SheetID == "01-KHBS" || sInfo.SheetID == "01-1-KHBS")
						{
							if (sInfo.Tables[i].TablePath != null)
							{
								string[] array = sInfo.Tables[i].TablePath.Split(new char[1] { '/' });
								val4 = val2[j].ParentNode.SelectNodes(array[^1]);
							}
							else
							{
								val4 = val2[j].ParentNode.ChildNodes;
							}
						}
						else if (WinSettings.MenuID == "G5")
						{
							if (sInfo.Tables[i].TablePath != null)
							{
								string[] array = sInfo.Tables[i].TablePath.Split(new char[1] { '/' });
								val4 = val2[j].ParentNode.SelectNodes(array[^1]);
							}
							else
							{
								val4 = val2[j].ParentNode.ChildNodes;
							}
						}
						else if (sInfo.Tables[i].TablePath != null)
						{
							string[] array = sInfo.Tables[i].TablePath.Split(new char[1] { '/' });
							val4 = val2[j].ParentNode.SelectNodes(array[^1]);
						}
						else
						{
							val4 = val2[j].ParentNode.ChildNodes;
						}
					}
					int num6 = 0;
					if (val4 != null)
					{
						foreach (XmlNode item in val4)
						{
							XmlNode val5 = item;
							num6++;
							HTKKCMCRowValue hTKKCMCRowValue = new HTKKCMCRowValue();
							for (int l = 0; l < hTKKCMCTableInfo.RowsInfo.Cells.Count; l++)
							{
								if (hTKKCMCTableInfo.RowsInfo.Cells[l].Path != null)
								{
									if (hTKKCMCTableInfo.RowsInfo.Cells[l].Path.Length > 0)
									{
										XmlNode val = val5.SelectSingleNode(hTKKCMCTableInfo.RowsInfo.Cells[l].Path);
										if (val != null)
										{
											if (hTKKCMCTableInfo.RowsInfo.Cells[l].Controltype != null)
											{
												ReChangeValue(val, hTKKCMCTableInfo.RowsInfo.Cells[l].Controltype);
											}
											hTKKCMCRowValue.RowValue.Add(val.InnerText);
										}
										else if (hTKKCMCTableInfo.RowsInfo.Cells[l].CellTemplate.ColIndex == hTKKCMCTableInfo.iColumnSTT)
										{
											hTKKCMCRowValue.RowValue.Add(j + 1 + "." + num6);
										}
										else
										{
											hTKKCMCRowValue.RowValue.Add(hTKKCMCTableInfo.RowsInfo.Cells[l].DefaultValue);
										}
									}
									else
									{
										hTKKCMCRowValue.RowValue.Add(hTKKCMCTableInfo.RowsInfo.Cells[l].DefaultValue);
									}
								}
								else if (hTKKCMCTableInfo.RowsInfo.Cells[l].CellTemplate.ColIndex == hTKKCMCTableInfo.iColumnSTT)
								{
									hTKKCMCRowValue.RowValue.Add(j + 1 + "." + num6);
								}
								else
								{
									hTKKCMCRowValue.RowValue.Add(hTKKCMCTableInfo.RowsInfo.Cells[l].DefaultValue);
								}
							}
							if (hTKKCMCRowValue.RowValue.Count == hTKKCMCTableInfo.RowsInfo.Cells.Count && !sInfo.Tables[i].IsEqualDefaultValue(hTKKCMCRowValue))
							{
								hTKKCMCTableInfo.RowsValues.Add(hTKKCMCRowValue);
							}
						}
					}
					if (hTKKCMCTableInfo.RowsValues.Count == 0 && hTKKCMCTableInfo.RowsInfo.Cells.Count > 0)
					{
						hTKKCMCTableInfo.RowsValues.Add(hTKKCMCTableInfo.getDefaultRowValue());
					}
					if (hTKKCMCTableInfo.RowsValues.Count - 1 > 0)
					{
						hTKKCMCTableInfo.IncrementFooter(hTKKCMCTableInfo.RowsValues.Count - 1);
					}
					for (int k = 0; k < hTKKCMCTableInfo.Footer.StaticCells.Count; k++)
					{
						if (hTKKCMCTableInfo.Footer.StaticCells[k].Path == null || hTKKCMCTableInfo.Footer.StaticCells[k].Path.Length <= 0 || hTKKCMCTableInfo.Footer.StaticCells[k].CellTemplate == null)
						{
							continue;
						}
						string path = hTKKCMCTableInfo.Footer.StaticCells[k].Path;
						XmlNode val = val2[j].SelectSingleNode(path);
						if (val != null)
						{
							if (hTKKCMCTableInfo.Footer.StaticCells[k].Controltype != null)
							{
								ReChangeValue(val, hTKKCMCTableInfo.Footer.StaticCells[k].Controltype);
							}
							hTKKCMCTableInfo.Footer.StaticCells[k].Value = val.InnerXml;
						}
					}
					if (i < count - 1)
					{
						for (int k = i + 1; k < count; k++)
						{
							if (WinSettings.MenuID == "H2" || WinSettings.MenuID == "G9" || WinSettings.MenuID == "H4")
							{
								if (hTKKCMCTableInfo.RowsInfo != null && hTKKCMCTableInfo.RowsInfo.RowPath != "")
								{
									if (j == 0)
									{
										sInfo.Tables[k].Increment(hTKKCMCTableInfo.RowsValues.Count - 1);
										sInfo.Tables[k].IncrementFooter(hTKKCMCTableInfo.RowsValues.Count - 1);
										num3 += hTKKCMCTableInfo.RowsValues.Count - 1;
									}
									else
									{
										sInfo.Tables[k].Increment(hTKKCMCTableInfo.RowsValues.Count + 1);
										sInfo.Tables[k].IncrementFooter(hTKKCMCTableInfo.RowsValues.Count + 1);
									}
								}
								else
								{
									sInfo.Tables[k].Increment(hTKKCMCTableInfo.RowsValues.Count - 1);
									sInfo.Tables[k].IncrementFooter(hTKKCMCTableInfo.RowsValues.Count - 1);
								}
							}
							else if (sInfo.SheetID == "01-KHBS" || sInfo.SheetID == "01-1-KHBS")
							{
								if (j == 0)
								{
									sInfo.Tables[k].Increment(hTKKCMCTableInfo.RowsValues.Count - 1);
									num3 += hTKKCMCTableInfo.RowsValues.Count - 1;
								}
								else
								{
									sInfo.Tables[k].Increment(hTKKCMCTableInfo.RowsValues.Count + 1);
								}
							}
							else
							{
								sInfo.Tables[k].Increment(hTKKCMCTableInfo.RowsValues.Count - 1);
							}
						}
					}
					if (j != 0)
					{
						num2 += hTKKCMCTableInfo.TableLength;
					}
					hTKKCMCTableInfo.Increment(num2);
					hTKKCMCTableInfo.IncrementFooter(num2);
					list.Add(hTKKCMCTableInfo);
					if (hTKKCMCTableInfo.RowsValues.Count > 0)
					{
						num2 = ((hTKKCMCTableInfo.TableLength > 0) ? (num2 + hTKKCMCTableInfo.RowsValues.Count) : (num2 + (hTKKCMCTableInfo.RowsValues.Count - 1)));
						if (hTKKCMCTableInfo.RowsInfo.Cells.Count == 0)
						{
							num2--;
						}
						if (hTKKCMCTableInfo.Header.StaticCells.Count > 0)
						{
							num5++;
						}
					}
				}
				if (WinSettings.ToKhaiInfo.KieuKHBS == "3" && (sInfo.SheetName == "01-1/KHBS" || sInfo.SheetName == "01-KHBS"))
				{
					if (num5 > 1)
					{
						num5--;
					}
					if (num2 - num5 != 0)
					{
						if ("I7,I4,C7,I8,C8".Contains(WinSettings.MenuID))
						{
							sInfo.IncrementStaticCell(sInfo.Tables[i].Location.DataLocationTemplateRowStart, num2);
						}
						else if (sInfo.SheetName == "01-KHBS")
						{
							sInfo.IncrementStaticCell(sInfo.Tables[i].Location.DataLocationTemplateRowStart, num2);
						}
						else
						{
							sInfo.IncrementStaticCell(sInfo.Tables[i].Location.DataLocationTemplateRowStart, num2 - 1);
						}
					}
				}
				else if (WinSettings.MenuID == "H2" || WinSettings.MenuID == "G9" || WinSettings.MenuID == "H4" || WinSettings.MenuID == "N8")
				{
					if (flag)
					{
						if (num5 > 1)
						{
							num5--;
						}
						sInfo.IncrementStaticCell(sInfo.Tables[i].Location.DataLocationTemplateRowStart, num2 - num5 + num3);
					}
					else
					{
						if (num5 > 1)
						{
							num5--;
						}
						sInfo.IncrementStaticCell(sInfo.Tables[i].Location.DataLocationTemplateRowStart, num2 - num5);
					}
				}
				else if (WinSettings.MenuID == "D1" || WinSettings.MenuID.Equals("Z5"))
				{
					if (num5 > 1)
					{
						num5 -= sInfo.Tables[i].TableLength;
						num5--;
					}
					sInfo.IncrementStaticCell(sInfo.Tables[i].Location.DataLocationTemplateRowStart, num2 - num5);
				}
				else
				{
					if (num5 > 1)
					{
						num5--;
					}
					sInfo.IncrementStaticCell(sInfo.Tables[i].Location.DataLocationTemplateRowStart, num2 - num5);
				}
			}
			sInfo.Tables = list;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static bool WinDoWorkImportFormXml(FpSpread winFpSpread, string fileName)
	{
		//IL_00ec: Unknown result type (might be due to invalid IL or missing references)
		//IL_00f2: Expected O, but got Unknown
		//IL_01a9: Unknown result type (might be due to invalid IL or missing references)
		//IL_02b2: Unknown result type (might be due to invalid IL or missing references)
		//IL_03d5: Unknown result type (might be due to invalid IL or missing references)
		//IL_04a0: Unknown result type (might be due to invalid IL or missing references)
		//IL_0446: Unknown result type (might be due to invalid IL or missing references)
		//IL_0bf0: Unknown result type (might be due to invalid IL or missing references)
		//IL_0bf7: Expected O, but got Unknown
		//IL_10b5: Unknown result type (might be due to invalid IL or missing references)
		//IL_10bc: Expected O, but got Unknown
		//IL_096e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0a36: Unknown result type (might be due to invalid IL or missing references)
		try
		{
			string text = (string.IsNullOrEmpty(fileName) ? "" : HTKKFile.LoadFile(fileName));
			if (text.IndexOf("xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\"") != -1)
			{
				text = text.Replace("xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\"", "");
			}
			if (text.IndexOf("xmlns=\"http://kekhaithue.gdt.gov.vn/TKhaiThue\"") != -1)
			{
				text = text.Replace("xmlns=\"http://kekhaithue.gdt.gov.vn/TKhaiThue\"", "");
			}
			if (text.IndexOf("xsi:nil=\"true\"") != -1)
			{
				text = text.Replace("xsi:nil=\"true\"", "");
			}
			if (WinSettings.MenuID == "M4")
			{
				if (text.IndexOf("xmlns:n2=\"http://www.altova.com/samplexml/other-namespace\"") != -1)
				{
					text = text.Replace("xmlns:n2=\"http://www.altova.com/samplexml/other-namespace\"", "");
				}
				if (text.IndexOf("xmlns:ds=\"http://www.w3.org/2000/09/xmldsig#\"") != -1)
				{
					text = text.Replace("xmlns:ds=\"http://www.w3.org/2000/09/xmldsig#\"", "");
				}
			}
			XmlDocument val = new XmlDocument();
			val.LoadXml(text);
			if (WinSettings.MenuID == "78" || "U9".Equals(WinSettings.MenuID))
			{
				string text2 = "";
				bool flag = ConvertHelpers.ConvertToBool(((XmlNode)val).SelectSingleNode("HSoThueDTu/HSoKhaiThue/CTieuTKhaiChinh/hoatDongLienTuc").InnerText.Trim());
				string text3 = "";
				text3 = ((!flag) ? "KL" : "LT");
				if (!WinSettings.TTPeriodInfo.DoanhNghiepHoatDongLienTuc.Equals(text3))
				{
					winFpSpread.ResumeLayout();
					frmMessage frmMessage = new frmMessage("0278", "3", "3", "Thông báo");
					((Form)frmMessage).ShowDialog();
					return false;
				}
			}
			XmlNode val2 = ((XmlNode)val).SelectSingleNode("HSoThueDTu/HSoKhaiThue/TTinChung/TTinTKhaiThue/NNT/mst");
			if (val2 == null && (WinSettings.MenuID == "B2" || WinSettings.MenuID == "O9" || WinSettings.MenuID == "M6" || WinSettings.MenuID == "O8"))
			{
				val2 = ((XmlNode)val).SelectSingleNode("HSoThueDTu/HSoKhaiThue/TTinChung/TTinTKhaiThue/mst");
			}
			if (val2 == null && WinSettings.MenuID == "M4")
			{
				val2 = ((XmlNode)val).SelectSingleNode("TKhai/DLTKhai/TTChung/MST");
			}
			if (val2.InnerText.Trim().Replace("-", "") != WinSettings.TTinChung.mst)
			{
				winFpSpread.ResumeLayout();
				frmMessage frmMessage = new frmMessage("0437", "3", "3", "Thông báo");
				((Form)frmMessage).ShowDialog();
				return false;
			}
			XmlNode val3 = ((XmlNode)val).SelectSingleNode("HSoThueDTu/HSoKhaiThue/TTinChung/TTinTKhaiThue/TKhaiThue/maTKhai");
			if (val3 == null && (WinSettings.MenuID == "B2" || WinSettings.MenuID == "O9" || WinSettings.MenuID == "M6" || WinSettings.MenuID == "O8"))
			{
				val3 = ((XmlNode)val).SelectSingleNode("HSoThueDTu/HSoKhaiThue/TTinChung/TTinTKhaiThue/maTKhai");
			}
			if (val3 == null && WinSettings.MenuID == "M4")
			{
				val3 = ((XmlNode)val).SelectSingleNode("TKhai/DLTKhai/TTChung/MSo");
			}
			if (WinSettings.MenuID == "68" && WinSettings.TTPeriodInfo.KieuKy == "Q")
			{
				if (val3.InnerText.Trim() != "102")
				{
					winFpSpread.ResumeLayout();
					frmMessage frmMessage = new frmMessage("0350", "3", "3", "Thông báo");
					((Form)frmMessage).ShowDialog();
					return false;
				}
			}
			else if (WinSettings.MenuID == "M4")
			{
				if (val3.InnerText.Trim() != "03/DL-HĐĐT")
				{
					winFpSpread.ResumeLayout();
					frmMessage frmMessage = new frmMessage("0539", "3", "3", "Thông báo");
					((Form)frmMessage).ShowDialog();
					return false;
				}
			}
			else if (val3.InnerText.Trim() != WinSettings.TTinChung.maTKhai)
			{
				winFpSpread.ResumeLayout();
				frmMessage frmMessage = new frmMessage("0350", "3", "3", "Thông báo");
				((Form)frmMessage).ShowDialog();
				return false;
			}
			if ("66;67;09;10;60;61;G2;".IndexOf(WinSettings.MenuID) == -1 && WinSettings.MenuID != "B2" && WinSettings.MenuID != "O9" && WinSettings.MenuID != "M4" && WinSettings.MenuID != "M6" && WinSettings.MenuID != "O8")
			{
				string text4 = ((XmlNode)val).SelectSingleNode("HSoThueDTu/HSoKhaiThue/TTinChung/TTinTKhaiThue/TKhaiThue/KyKKhaiThue/kieuKy").InnerText.Trim();
				string text5 = ((XmlNode)val).SelectSingleNode("HSoThueDTu/HSoKhaiThue/TTinChung/TTinTKhaiThue/TKhaiThue/KyKKhaiThue/kyKKhai").InnerText.Trim();
				string text6 = ((XmlNode)val).SelectSingleNode("HSoThueDTu/HSoKhaiThue/TTinChung/TTinTKhaiThue/TKhaiThue/KyKKhaiThue/kyKKhaiTuNgay").InnerText.Trim();
				string text7 = ((XmlNode)val).SelectSingleNode("HSoThueDTu/HSoKhaiThue/TTinChung/TTinTKhaiThue/TKhaiThue/KyKKhaiThue/kyKKhaiDenNgay").InnerText.Trim();
				string text8 = ((XmlNode)val).SelectSingleNode("HSoThueDTu/HSoKhaiThue/TTinChung/TTinTKhaiThue/TKhaiThue/KyKKhaiThue/kyKKhaiTuThang").InnerText.Trim();
				string text9 = ((XmlNode)val).SelectSingleNode("HSoThueDTu/HSoKhaiThue/TTinChung/TTinTKhaiThue/TKhaiThue/KyKKhaiThue/kyKKhaiDenThang").InnerText.Trim();
				string text10 = ((XmlNode)val).SelectSingleNode("HSoThueDTu/HSoKhaiThue/TTinChung/TTinTKhaiThue/TKhaiThue/loaiTKhai").InnerText.Trim();
				if ((WinSettings.MenuID == "G4" || WinSettings.MenuID == "N5" || WinSettings.MenuID == "H5" || WinSettings.MenuID == "O7" || WinSettings.MenuID == "P5" || WinSettings.MenuID == "U6" || "U8~".Contains(WinSettings.MenuID)) && text10 == "B")
				{
					text10 = "T";
				}
				if (!(WinSettings.MenuID == "A4") || text4 == "Q")
				{
				}
				if (WinSettings.MenuID == "G7" && text4 == "D")
				{
					text4 = "D_D";
				}
				if (WinSettings.MenuID == "G3")
				{
					if (text4 == "DK")
					{
						text4 = "M";
					}
					if (text4 == "LD")
					{
						text4 = "D";
					}
				}
				if (WinSettings.MenuID == "N2" || WinSettings.MenuID == "N6")
				{
					if (text4 == "M")
					{
						text4 = "D";
					}
					text5 = WinSettings.TTinChung.kyKKhai;
				}
				if (!(WinSettings.MenuID == "66") && !(WinSettings.MenuID == "67") && !(WinSettings.MenuID == "09") && !(WinSettings.MenuID == "10") && !(WinSettings.MenuID == "60") && !(WinSettings.MenuID == "61") && !(WinSettings.MenuID == "G2") && (!(WinSettings.MenuID == "G7") || !(text4 == "D_D")))
				{
					if ((WinSettings.MenuID != "29" && WinSettings.MenuID != "10" && WinSettings.MenuID != "G3" && WinSettings.MenuID != "35" && WinSettings.MenuID != "Q4" && (text4 != (WinSettings.TTinChung.kieuKy.Equals("D_D") ? "Y" : WinSettings.TTinChung.kieuKy) || text6 != WinSettings.TTinChung.kyKKhaiTuNgay || text7 != WinSettings.TTinChung.kyKKhaiDenNgay || text8 != WinSettings.TTinChung.kyKKhaiTuThang || text9 != WinSettings.TTinChung.kyKKhaiDenThang)) || ("78;A1;A3;53;19,22;64;65;66;67;68;18;27;91;07;13;09;10;14;60;61;62;17;G2;U9;X2;X1;".IndexOf(WinSettings.MenuID) == -1 && text10 != WinSettings.TTPeriodInfo.LoaiToKhai && !WinSettings.MenuID.Equals("95")))
					{
						winFpSpread.ResumeLayout();
						frmMessage frmMessage = new frmMessage("0437", "3", "3", "Thông báo");
						((Form)frmMessage).ShowDialog();
						return false;
					}
					if ((!(WinSettings.MenuID == "30") || !(WinSettings.TTinChung.kieuKy == "Q")) && !(WinSettings.MenuID == "A4") && !(WinSettings.MenuID == "99") && !(WinSettings.MenuID == "H6") && text5 != WinSettings.TTinChung.kyKKhai)
					{
						winFpSpread.ResumeLayout();
						frmMessage frmMessage = new frmMessage("0437", "3", "3", "Thông báo");
						((Form)frmMessage).ShowDialog();
						return false;
					}
				}
			}
			XmlNode documentElement = (XmlNode)(object)val.DocumentElement;
			StartImportXML(documentElement, 0);
			if (WinSettings.MenuID == "M4")
			{
				ProcessXMLNode(WinSettings.ToKhaiInfo.SheetInfos[0], ((XmlNode)val).SelectSingleNode("TKhai"), checkKHBS: false);
			}
			else
			{
				ProcessXMLNode(WinSettings.ToKhaiInfo.SheetInfos[0], ((XmlNode)val).SelectSingleNode("HSoThueDTu"), checkKHBS: false);
			}
			ProcessSelectedCell(WinSettings.ToKhaiInfo.SheetInfos[0]);
			ProcessParentCell(WinSettings.ToKhaiInfo.SheetInfos[0]);
			WinSettings.ToKhaiInfo.SheetInfos[0].IsReadDataFile = true;
			XmlNode val4 = ((XmlNode)val).SelectSingleNode("HSoThueDTu/HSoKhaiThue/PLuc");
			XmlNodeList val5 = null;
			if (val4 != null && val4.HasChildNodes)
			{
				val5 = val4.ChildNodes;
			}
			for (int i = 1; i < WinSettings.ToKhaiInfo.SheetInfos.Count; i++)
			{
				if (WinSettings.ToKhaiInfo.SheetInfos[i] != null && WinSettings.DicSheet.ContainsKey(WinSettings.ToKhaiInfo.SheetInfos[i].SheetName))
				{
					WinSettings.DicSheet.Remove(WinSettings.ToKhaiInfo.SheetInfos[i].SheetName);
				}
				WinSettings.ToKhaiInfo.SheetInfos[i] = null;
			}
			if (val5 != null && val5.Count > 0)
			{
				XmlNodeList val6 = null;
				XmlDocument val7 = new XmlDocument();
				val7.LoadXml(WinSettings.xmlMenuNode.OuterXml);
				val6 = val7.GetElementsByTagName("Sheet");
				List<HTKKSheetInMenu> allSheetInMenu = getAllSheetInMenu(val6);
				if (val5.Count > 0)
				{
					for (int i = 0; i < val5.Count; i++)
					{
						XmlNode val8 = val5[i];
						string name = val8.Name;
						for (int j = 1; j < allSheetInMenu.Count; j++)
						{
							string text11 = "";
							string text12 = "";
							bool checkKHBS = false;
							if (WinSettings.TTPeriodInfo.LoaiToKhai == "B")
							{
								if ("30,I5,K9,K1,M7,K6,K8,73,55,71,31,P2,X6".Contains(WinSettings.MenuID))
								{
									text11 = ((j <= 1) ? GetFileNameKHBSTheoThongTu(WinSettings.ToKhaiInfo.InterfaceTemplateXml) : string.Format("{0}{1}_xml_{2}.xml", WinSettings.ToKhaiInfo.InterfaceTemplateXml, allSheetInMenu[j - 1].dataFile, WinSettings.ToKhaiInfo.VersionXML.Replace(".", "")));
									checkKHBS = WinSettings.ToKhaiInfo.KieuKHBS == "2";
								}
								else if (WinSettings.MenuID == "N7")
								{
									if (j == allSheetInMenu.Count - 1)
									{
										continue;
									}
									text11 = string.Format("{0}{1}_xml_{2}.xml", WinSettings.ToKhaiInfo.InterfaceTemplateXml, allSheetInMenu[j].dataFile, WinSettings.ToKhaiInfo.VersionXML.Replace(".", ""));
								}
								else
								{
									if (!(WinSettings.ToKhaiInfo.KieuKHBS == "3") && WinSettings.ToKhaiInfo.KieuKHBS != "" && j != allSheetInMenu.Count - 1)
									{
										continue;
									}
									text11 = ((!(WinSettings.ToKhaiInfo.KieuKHBS == "3")) ? GetFileNameKHBSTheoThongTu(WinSettings.ToKhaiInfo.InterfaceTemplateXml) : ((allSheetInMenu[j].sheetName == "01-1/KHBS") ? ((!"I7,I4,C7,I8,C8".Contains(WinSettings.MenuID)) ? string.Format("{0}{1}.xml", WinSettings.ToKhaiInfo.InterfaceTemplateXml, "KHBS_01_1_TT80_xml") : string.Format("{0}{1}.xml", WinSettings.ToKhaiInfo.InterfaceTemplateXml, "KHBS_01_1_TT80_DVT_xml")) : ((!(allSheetInMenu[j].sheetName == "01-KHBS")) ? string.Format("{0}{1}_xml_{2}.xml", WinSettings.ToKhaiInfo.InterfaceTemplateXml, allSheetInMenu[j].dataFile, WinSettings.ToKhaiInfo.VersionXML.Replace(".", "")) : ("I7,I4,C7,I8,C8".Contains(WinSettings.MenuID) ? string.Format("{0}{1}.xml", WinSettings.ToKhaiInfo.InterfaceTemplateXml, "KHBS_01_TT80_DVT_xml") : ((!"D1,N8,G9,H2,H4,T8,E8,L8,L7,N2,G2,R5,Y6,Z8,Z5".Contains(WinSettings.MenuID)) ? string.Format("{0}{1}.xml", WinSettings.ToKhaiInfo.InterfaceTemplateXml, "KHBS_01_TT80_xml") : string.Format("{0}{1}.xml", WinSettings.ToKhaiInfo.InterfaceTemplateXml, "KHBS_01_DIA_BAN_xml"))))));
									checkKHBS = WinSettings.ToKhaiInfo.KieuKHBS == "2";
								}
							}
							if (WinSettings.TTPeriodInfo.LoaiToKhai == "C" || "31;30;39;33;28;M2".IndexOf(WinSettings.MenuID) > -1 || (WinSettings.TTPeriodInfo.LoaiToKhai == "T" && WinSettings.MenuID == "N7"))
							{
								if (WinSettings.ToKhaiInfo.KieuKHBS == "3")
								{
									if (j == allSheetInMenu.Count - 1 || j == allSheetInMenu.Count - 2)
									{
										continue;
									}
								}
								else if (j == allSheetInMenu.Count - 1)
								{
									continue;
								}
								text11 = string.Format("{0}{1}_xml_{2}.xml", WinSettings.ToKhaiInfo.InterfaceTemplateXml, allSheetInMenu[j].dataFile, WinSettings.ToKhaiInfo.VersionXML.Replace(".", ""));
							}
							XmlDocument val9 = new XmlDocument();
							val9.Load(text11);
							text12 = ((WinSettings.TTPeriodInfo.LoaiToKhai == "C" || "31;30;39;33;28;M2".IndexOf(WinSettings.MenuID) > -1 || (WinSettings.TTPeriodInfo.LoaiToKhai == "T" && WinSettings.MenuID == "N7")) ? ((XmlNode)val9).ChildNodes[0].Name : ((WinSettings.TTPeriodInfo.LoaiToKhai == "B" && WinSettings.MenuID == "N7") ? ((XmlNode)val9).ChildNodes[0].Name : (("30,I5,K9,K1,M7,K6,K8,73,55,71,31,P2,X6".Contains(WinSettings.MenuID) && j > 1) ? ((XmlNode)val9).ChildNodes[0].Name : ((!(WinSettings.ToKhaiInfo.KieuKHBS == "3")) ? ((XmlNode)val9).ChildNodes[1].Name : ((XmlNode)val9).ChildNodes[0].Name))));
							if (!(name == text12))
							{
								continue;
							}
							if ("30,I5,K9,K1,M7,K6,K8,73,55,71,31,P2,X6".Contains(WinSettings.MenuID) && WinSettings.TTinChung.loaiTKhai == "B")
							{
								if (j > 1)
								{
									HTKKCMCSheetInfo hTKKSheetInfo = WinFnProcess.GetHTKKSheetInfo(WinSettings.ToKhaiInfo, val6, allSheetInMenu[j - 1].stt, getTemplate: true, WinSettings.TTPeriodInfo);
									WinSettings.ToKhaiInfo.SheetInfos[allSheetInMenu[j - 1].stt] = hTKKSheetInfo;
									WinSettings.DicSheet.Add(((XmlNode)val6[allSheetInMenu[j - 1].stt].Attributes["SheetName"]).Value, 2);
									WinSettings.ToKhaiInfo.SheetInfos[allSheetInMenu[j - 1].stt].IsReadDataFile = true;
									ProcessXMLNode(WinSettings.ToKhaiInfo.SheetInfos[allSheetInMenu[j - 1].stt], val5[i], checkKHBS);
									ProcessSelectedCell(WinSettings.ToKhaiInfo.SheetInfos[allSheetInMenu[j - 1].stt]);
									ProcessParentCell(WinSettings.ToKhaiInfo.SheetInfos[allSheetInMenu[j - 1].stt]);
								}
								else
								{
									HTKKCMCSheetInfo hTKKSheetInfo = WinFnProcess.GetHTKKSheetInfo(WinSettings.ToKhaiInfo, val6, allSheetInMenu[allSheetInMenu.Count - 1].stt, getTemplate: true, WinSettings.TTPeriodInfo);
									WinSettings.ToKhaiInfo.SheetInfos[allSheetInMenu[allSheetInMenu.Count - 1].stt] = hTKKSheetInfo;
									WinSettings.DicSheet.Add(((XmlNode)val6[allSheetInMenu[allSheetInMenu.Count - 1].stt].Attributes["SheetName"]).Value, 2);
									WinSettings.ToKhaiInfo.SheetInfos[allSheetInMenu[allSheetInMenu.Count - 1].stt].IsReadDataFile = true;
									ProcessXMLNode(WinSettings.ToKhaiInfo.SheetInfos[allSheetInMenu[allSheetInMenu.Count - 1].stt], val5[i], checkKHBS);
									ProcessSelectedCell(WinSettings.ToKhaiInfo.SheetInfos[allSheetInMenu[allSheetInMenu.Count - 1].stt]);
									ProcessParentCell(WinSettings.ToKhaiInfo.SheetInfos[allSheetInMenu[allSheetInMenu.Count - 1].stt]);
								}
							}
							else if (WinSettings.ToKhaiInfo.KieuKHBS == "3" && WinSettings.TTinChung.loaiTKhai == "B")
							{
								HTKKCMCSheetInfo hTKKSheetInfo = WinFnProcess.GetHTKKSheetInfo(WinSettings.ToKhaiInfo, val6, allSheetInMenu[j].stt, getTemplate: true, WinSettings.TTPeriodInfo);
								WinSettings.ToKhaiInfo.SheetInfos[allSheetInMenu[j].stt] = hTKKSheetInfo;
								WinSettings.DicSheet.Add(((XmlNode)val6[allSheetInMenu[j].stt].Attributes["SheetName"]).Value, 2);
								WinSettings.ToKhaiInfo.SheetInfos[allSheetInMenu[j].stt].IsReadDataFile = true;
								ProcessXMLNode(WinSettings.ToKhaiInfo.SheetInfos[allSheetInMenu[j].stt], val5[i], checkKHBS);
								ProcessSelectedCell(WinSettings.ToKhaiInfo.SheetInfos[allSheetInMenu[j].stt]);
								ProcessParentCell(WinSettings.ToKhaiInfo.SheetInfos[allSheetInMenu[j].stt]);
							}
							else
							{
								HTKKCMCSheetInfo hTKKSheetInfo = WinFnProcess.GetHTKKSheetInfo(WinSettings.ToKhaiInfo, val6, allSheetInMenu[j].stt, getTemplate: true, WinSettings.TTPeriodInfo);
								WinSettings.ToKhaiInfo.SheetInfos[allSheetInMenu[j].stt] = hTKKSheetInfo;
								WinSettings.DicSheet.Add(((XmlNode)val6[allSheetInMenu[j].stt].Attributes["SheetName"]).Value, 2);
								WinSettings.ToKhaiInfo.SheetInfos[allSheetInMenu[j].stt].IsReadDataFile = true;
								ProcessXMLNode(WinSettings.ToKhaiInfo.SheetInfos[allSheetInMenu[j].stt], val5[i], checkKHBS);
								ProcessSelectedCell(WinSettings.ToKhaiInfo.SheetInfos[allSheetInMenu[j].stt]);
								ProcessParentCell(WinSettings.ToKhaiInfo.SheetInfos[allSheetInMenu[j].stt]);
							}
						}
					}
				}
			}
			return true;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return false;
		}
	}

	public static void FinishImportXML(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		MethodInfo[] methods = GetMethods();
		bool flag = false;
		for (int i = 0; i < methods.Length; i++)
		{
			if (methods[i].Name.IndexOf("FinishImportXML") == -1)
			{
				continue;
			}
			ParameterInfo[] parameters = methods[i].GetParameters();
			object[] array = new object[parameters.Length];
			for (int j = 0; j < parameters.Length; j++)
			{
				switch (parameters[j].Name)
				{
				case "winFpSpread":
					array[j] = winFpSpread;
					break;
				case "ToKhaiInfo":
					array[j] = WinSettings.ToKhaiInfo;
					break;
				case "TTinChung":
					array[j] = WinSettings.TTinChung;
					break;
				case "PeriodInfo":
					array[j] = WinSettings.TTPeriodInfo;
					break;
				case "HeaderInfo":
					array[j] = WinSettings.HTKKHeaderInfo;
					break;
				}
			}
			methods[i].Invoke(WinSettings.InstanceClass, array);
			break;
		}
	}

	public static void StartImportXML(XmlNode xmlNode, int index)
	{
		try
		{
			MethodInfo method = GetMethod("StartImportXML_" + index);
			if ((object)method == null)
			{
				return;
			}
			ParameterInfo[] parameters = method.GetParameters();
			object[] array = new object[parameters.Length];
			for (int i = 0; i < parameters.Length; i++)
			{
				switch (parameters[i].Name)
				{
				case "ToKhaiInfo":
					array[i] = WinSettings.ToKhaiInfo;
					break;
				case "TTinChung":
					array[i] = WinSettings.TTinChung;
					break;
				case "PeriodInfo":
					array[i] = WinSettings.TTPeriodInfo;
					break;
				case "xmlNode":
					array[i] = xmlNode;
					break;
				}
			}
			method.Invoke(WinSettings.InstanceClass, array);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void FinishImportExcel(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo)
	{
		MethodInfo[] methods = GetMethods();
		bool flag = false;
		for (int i = 0; i < methods.Length; i++)
		{
			if (methods[i].Name.IndexOf("FinishImportExcel") == -1)
			{
				continue;
			}
			ParameterInfo[] parameters = methods[i].GetParameters();
			object[] array = new object[parameters.Length];
			for (int j = 0; j < parameters.Length; j++)
			{
				switch (parameters[j].Name)
				{
				case "winFpSpread":
					array[j] = winFpSpread;
					break;
				case "ToKhaiInfo":
					array[j] = WinSettings.ToKhaiInfo;
					break;
				case "TTinChung":
					array[j] = WinSettings.TTinChung;
					break;
				case "PeriodInfo":
					array[j] = WinSettings.TTPeriodInfo;
					break;
				case "HeaderInfo":
					array[j] = WinSettings.HTKKHeaderInfo;
					break;
				}
			}
			methods[i].Invoke(WinSettings.InstanceClass, array);
			break;
		}
	}

	public static void ProcessSelectedCell(HTKKCMCSheetInfo sInfo)
	{
		try
		{
			List<HTKKCMCCellInfo> lstCbxCell = sInfo.StaticCells.FindAll((HTKKCMCCellInfo cell) => cell.CellSelectedValue != null);
			int i;
			for (i = 0; i < lstCbxCell.Count; i++)
			{
				HTKKCMCCellInfo hTKKCMCCellInfo = sInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.CellTemplate.ColIndex == lstCbxCell[i].CellSelectedValue.ColIndex && cell.CellTemplate.RowIndex == lstCbxCell[i].CellSelectedValue.RowIndex);
				if (hTKKCMCCellInfo != null)
				{
					GetValueFromSelectedCell(lstCbxCell[i], hTKKCMCCellInfo);
				}
				if (lstCbxCell[i].ParentCell != null)
				{
					HTKKCMCCellInfo hTKKCMCCellInfo2 = sInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.CellSelectedValue != null && cell.CellSelectedValue.ColIndex == lstCbxCell[i].ParentCell.ColIndex && cell.CellSelectedValue.RowIndex == lstCbxCell[i].ParentCell.RowIndex);
					if (hTKKCMCCellInfo2 != null && hTKKCMCCellInfo != null)
					{
						GetValueFromChildCell(hTKKCMCCellInfo2, hTKKCMCCellInfo2.Controltype, hTKKCMCCellInfo.Value);
					}
				}
			}
			if (!(WinSettings.ToKhaiInfo.SheetInfos[sInfo.SheetIndex].TaiDuLieuLon != "1") && !(WinSettings.ToKhaiInfo.SheetInfos[sInfo.SheetIndex].TaiDuLieuRatLon != "1"))
			{
				return;
			}
			for (int num = 0; num < sInfo.Tables.Count; num++)
			{
				HTKKCMCTableInfo hTKKCMCTableInfo = sInfo.Tables[num];
				for (int num2 = 0; num2 < hTKKCMCTableInfo.Header.StaticCells.Count; num2++)
				{
					if (hTKKCMCTableInfo.Header.StaticCells[num2].CellSelectedValue != null)
					{
						GetValueFromSelectedCell(hTKKCMCTableInfo.Header.StaticCells[num2], hTKKCMCTableInfo.Header.StaticCells[num2].CellSelectedValue);
					}
				}
				Dictionary<int, int[]> dictionary = new Dictionary<int, int[]>();
				int[] array = new int[3];
				for (int num3 = 0; num3 < hTKKCMCTableInfo.RowsInfo.Cells.Count; num3++)
				{
					HTKKCMCTableCellInfo hTKKCMCTableCellInfo = hTKKCMCTableInfo.RowsInfo.Cells[num3];
					try
					{
						if (string.IsNullOrEmpty(hTKKCMCTableCellInfo.indexColumnSelectedValue) && string.IsNullOrEmpty(hTKKCMCTableCellInfo.ChildValue))
						{
							continue;
						}
						array = new int[3];
						if (hTKKCMCTableCellInfo.Controltype != null && hTKKCMCTableCellInfo.Controltype.Length >= 3 && hTKKCMCTableCellInfo.Controltype.Substring(0, 3) == "DC_" && hTKKCMCTableCellInfo.Controltype.IndexOf("_M_") < 0)
						{
							array[0] = 0;
						}
						else if (hTKKCMCTableCellInfo.Controltype != null && hTKKCMCTableCellInfo.Controltype.Length >= 3 && hTKKCMCTableCellInfo.Controltype.Substring(0, 3) == "DC_" && hTKKCMCTableCellInfo.Controltype.IndexOf("_M_") >= 0)
						{
							array[0] = 0;
						}
						else
						{
							array[0] = Convert.ToInt32(hTKKCMCTableCellInfo.Controltype);
						}
						bool flag = false;
						bool flag2 = false;
						for (int num4 = 0; num4 < hTKKCMCTableInfo.RowsInfo.Cells.Count; num4++)
						{
							if (!flag && !string.IsNullOrEmpty(hTKKCMCTableCellInfo.indexColumnSelectedValue))
							{
								if (hTKKCMCTableInfo.RowsInfo.Cells[num4].CellTemplate.ColIndex == HTKKFile.ColumnLettersToIndex(hTKKCMCTableCellInfo.indexColumnSelectedValue))
								{
									array[1] = num4;
									flag = true;
								}
								else
								{
									array[1] = -1;
								}
							}
							if (!flag2 && !string.IsNullOrEmpty(hTKKCMCTableCellInfo.ChildValue))
							{
								if (hTKKCMCTableInfo.RowsInfo.Cells[num4].CellTemplate.ColIndex == HTKKFile.ColumnLettersToIndex(hTKKCMCTableCellInfo.ChildValue))
								{
									array[2] = num4;
									flag2 = true;
								}
								else
								{
									array[2] = -1;
								}
							}
							if (flag && flag2)
							{
								break;
							}
						}
						dictionary.Add(num3, array);
					}
					catch (Exception ex)
					{
						HTKKFile.Log(MethodBase.GetCurrentMethod().Name + " path " + hTKKCMCTableCellInfo.Path + " Controltype " + hTKKCMCTableCellInfo.Controltype, ex);
					}
				}
				for (int num5 = 0; num5 < hTKKCMCTableInfo.RowsValues.Count; num5++)
				{
					for (int num3 = 0; num3 < hTKKCMCTableInfo.RowsInfo.Cells.Count; num3++)
					{
						if (!dictionary.ContainsKey(num3))
						{
							continue;
						}
						HTKKCMCTableCellInfo hTKKCMCTableCellInfo2 = hTKKCMCTableInfo.RowsInfo.Cells[num3];
						int[] array2 = dictionary[num3];
						if (array2[1] != -1)
						{
							if (hTKKCMCTableCellInfo2 != null && hTKKCMCTableCellInfo2.Controltype.IndexOf("DC_") >= 0)
							{
								GetValueFromSelectedCell(hTKKCMCTableInfo.RowsValues[num5], hTKKCMCTableCellInfo2.Controltype, num3, array2[1]);
							}
							else
							{
								GetValueFromSelectedCell(hTKKCMCTableInfo.RowsValues[num5], array2[0].ToString(), num3, array2[1]);
							}
							if (hTKKCMCTableInfo.RowsValues[num5].RowValue[array2[1]].IsNulOrEmpty() && array2[2] != -1)
							{
								GetValueFromChildCell(hTKKCMCTableInfo.RowsValues[num5], array2[0].ToString(), num3, array2[2], array2[1]);
							}
						}
						else if (array2[2] != -1)
						{
							GetValueFromChildCell(hTKKCMCTableInfo.RowsValues[num5], array2[0].ToString(), num3, array2[2], array2[1]);
						}
					}
				}
				for (int num2 = 0; num2 < hTKKCMCTableInfo.Footer.StaticCells.Count; num2++)
				{
					if (hTKKCMCTableInfo.Footer.StaticCells[num2].CellSelectedValue != null)
					{
						GetValueFromSelectedCell(hTKKCMCTableInfo.Footer.StaticCells[num2], hTKKCMCTableInfo.Footer.StaticCells[num2].CellSelectedValue);
					}
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void ProcessParentCell(HTKKCMCSheetInfo sInfo)
	{
		try
		{
			List<HTKKCMCCellInfo> list = sInfo.StaticCells.Where((HTKKCMCCellInfo cell) => cell.ParentCell != null).ToList();
			foreach (HTKKCMCCellInfo childentNameCell in list)
			{
				try
				{
					if (childentNameCell.ParentCell == null)
					{
						continue;
					}
					HTKKCMCCellInfo hTKKCMCCellInfo = sInfo.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.CellTemplate.RowIndex == childentNameCell.ParentCell.RowIndex && x.CellTemplate.ColIndex == childentNameCell.ParentCell.ColIndex);
					if (hTKKCMCCellInfo != null && childentNameCell.CellSelectedValue != null)
					{
						HTKKCMCCellInfo hTKKCMCCellInfo2 = sInfo.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.CellTemplate.RowIndex == childentNameCell.CellSelectedValue.RowIndex && x.CellTemplate.ColIndex == childentNameCell.CellSelectedValue.ColIndex);
						if (hTKKCMCCellInfo2 != null && hTKKCMCCellInfo.Value == hTKKCMCCellInfo.DefaultValue && !string.IsNullOrEmpty(hTKKCMCCellInfo2.Value))
						{
							GetParentCodeFromChildCell(hTKKCMCCellInfo, childentNameCell.Controltype, hTKKCMCCellInfo2.Value);
						}
					}
					if (hTKKCMCCellInfo != null && childentNameCell.ParentCell != null)
					{
						HTKKCMCCellInfo hTKKCMCCellInfo3 = sInfo.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.CellSelectedValue != null && x.CellSelectedValue.RowIndex == childentNameCell.ParentCell.RowIndex && x.CellSelectedValue.ColIndex == childentNameCell.ParentCell.ColIndex);
						if (hTKKCMCCellInfo3 != null && !string.IsNullOrEmpty(hTKKCMCCellInfo.Value))
						{
							GetParentNameFromChildCell(hTKKCMCCellInfo3, childentNameCell.Controltype, hTKKCMCCellInfo.Value);
						}
					}
				}
				catch (Exception ex)
				{
					HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
				}
			}
			for (int num = 0; num < sInfo.Tables.Count; num++)
			{
				HTKKCMCTableInfo hTKKCMCTableInfo = sInfo.Tables[num];
				List<HTKKCMCCellInfo> list2 = hTKKCMCTableInfo.Header.StaticCells.Where((HTKKCMCCellInfo cell) => cell.ParentCell != null).ToList();
				foreach (HTKKCMCCellInfo childentNameCell2 in list2)
				{
					try
					{
						if (childentNameCell2.ParentCell == null)
						{
							continue;
						}
						HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCTableInfo.Header.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.CellTemplate.RowIndex == childentNameCell2.ParentCell.RowIndex && x.CellTemplate.ColIndex == childentNameCell2.ParentCell.ColIndex);
						if (hTKKCMCCellInfo != null && childentNameCell2.CellSelectedValue != null)
						{
							HTKKCMCCellInfo hTKKCMCCellInfo2 = hTKKCMCTableInfo.Header.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.CellTemplate.RowIndex == childentNameCell2.CellSelectedValue.RowIndex && x.CellTemplate.ColIndex == childentNameCell2.CellSelectedValue.ColIndex);
							if (hTKKCMCCellInfo2 != null && hTKKCMCCellInfo.Value == hTKKCMCCellInfo.DefaultValue)
							{
								GetParentCodeFromChildCell(hTKKCMCCellInfo, childentNameCell2.Controltype, hTKKCMCCellInfo2.Value);
							}
						}
						if (hTKKCMCCellInfo != null && childentNameCell2.ParentCell != null)
						{
							HTKKCMCCellInfo hTKKCMCCellInfo3 = hTKKCMCTableInfo.Header.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.CellSelectedValue != null && x.CellSelectedValue.RowIndex == childentNameCell2.ParentCell.RowIndex && x.CellSelectedValue.ColIndex == childentNameCell2.ParentCell.ColIndex);
							if (hTKKCMCCellInfo3 != null)
							{
								GetParentNameFromChildCell(hTKKCMCCellInfo3, childentNameCell2.Controltype, hTKKCMCCellInfo.Value);
							}
						}
					}
					catch (Exception ex)
					{
						HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
					}
				}
				for (int num2 = 0; num2 < hTKKCMCTableInfo.RowsValues.Count; num2++)
				{
					try
					{
						List<HTKKCMCTableCellInfo> list3 = hTKKCMCTableInfo.RowsInfo.Cells.Where((HTKKCMCTableCellInfo x) => x.ParentValue != null).ToList();
						foreach (HTKKCMCTableCellInfo childentNameCell3 in list3)
						{
							if (string.IsNullOrEmpty(childentNameCell3.ParentValue))
							{
								continue;
							}
							int num3 = hTKKCMCTableInfo.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.CellTemplate.ColIndex == HTKKFile.ColumnLettersToIndex(childentNameCell3.ParentValue));
							string text = "";
							if (num3 >= 0 && !string.IsNullOrEmpty(childentNameCell3.indexColumnSelectedValue))
							{
								int num4 = hTKKCMCTableInfo.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.CellTemplate.ColIndex == HTKKFile.ColumnLettersToIndex(childentNameCell3.indexColumnSelectedValue));
								if (num4 >= 0 && !string.IsNullOrEmpty(hTKKCMCTableInfo.RowsValues[num2].RowValue[num4]))
								{
									text = hTKKCMCTableInfo.RowsValues[num2].RowValue[num4];
									GetParentCodeFromChildCellInTable(hTKKCMCTableInfo.RowsValues[num2], childentNameCell3.Controltype, num3, hTKKCMCTableInfo.RowsValues[num2].RowValue[num4]);
								}
							}
							if (num3 >= 0 && childentNameCell3.ParentValue != null)
							{
								int num5 = hTKKCMCTableInfo.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.indexColumnSelectedValue != null && HTKKFile.ColumnLettersToIndex(x.indexColumnSelectedValue) == HTKKFile.ColumnLettersToIndex(childentNameCell3.ParentValue));
								if (num5 >= 0 && !string.IsNullOrEmpty(text))
								{
									GetParentNameFromChildCellInTable(hTKKCMCTableInfo.RowsValues[num2], childentNameCell3.Controltype, num5, text);
								}
							}
						}
					}
					catch (Exception ex)
					{
						HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
					}
				}
				List<HTKKCMCCellInfo> list4 = hTKKCMCTableInfo.Footer.StaticCells.Where((HTKKCMCCellInfo cell) => cell.ParentCell != null).ToList();
				foreach (HTKKCMCCellInfo childentNameCell4 in list4)
				{
					try
					{
						if (childentNameCell4.ParentCell == null)
						{
							continue;
						}
						HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCTableInfo.Footer.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.CellTemplate.RowIndex == childentNameCell4.ParentCell.RowIndex && x.CellTemplate.ColIndex == childentNameCell4.ParentCell.ColIndex);
						if (hTKKCMCCellInfo != null && childentNameCell4.CellSelectedValue != null)
						{
							HTKKCMCCellInfo hTKKCMCCellInfo2 = hTKKCMCTableInfo.Footer.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.CellTemplate.RowIndex == childentNameCell4.CellSelectedValue.RowIndex && x.CellTemplate.ColIndex == childentNameCell4.CellSelectedValue.ColIndex);
							if (hTKKCMCCellInfo2 != null && hTKKCMCCellInfo.Value == hTKKCMCCellInfo.DefaultValue && !string.IsNullOrEmpty(hTKKCMCCellInfo2.Value))
							{
								GetParentCodeFromChildCell(hTKKCMCCellInfo, childentNameCell4.Controltype, hTKKCMCCellInfo2.Value);
							}
						}
						if (hTKKCMCCellInfo != null && childentNameCell4.ParentCell != null)
						{
							HTKKCMCCellInfo hTKKCMCCellInfo3 = hTKKCMCTableInfo.Footer.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.CellSelectedValue != null && x.CellSelectedValue.RowIndex == childentNameCell4.ParentCell.RowIndex && x.CellSelectedValue.ColIndex == childentNameCell4.ParentCell.ColIndex);
							if (hTKKCMCCellInfo3 != null && !string.IsNullOrEmpty(hTKKCMCCellInfo.Value))
							{
								GetParentNameFromChildCell(hTKKCMCCellInfo3, childentNameCell4.Controltype, hTKKCMCCellInfo.Value);
							}
						}
					}
					catch (Exception ex)
					{
						HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
					}
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void GetValueFromSelectedCell(HTKKCMCCellInfo StaticCell, HTKKCMCCellInfo ChildCell)
	{
		string value = StaticCell.Value;
		string value2 = ChildCell.Value;
		try
		{
			if (StaticCell.Controltype == "8")
			{
				if (!(WinSettings.ToKhaiInfo.FolderName.ToLower() == "header") && !"D1,N8,G9,H2,H4,T8,E8,L8,L7,N2,G2,D2,Q5,Q6,U6,Y6,Z8,Z5".Contains(WinSettings.MenuID))
				{
					StaticCell.Value = HTKKWinDictionaryDiaDanh.GetTenCoQuanThueCapCuc(value2, WinSettings.ToKhaiInfo.FolderName.ToLower() == "header");
				}
			}
			else if (StaticCell.Controltype == "9")
			{
				if (!(WinSettings.ToKhaiInfo.FolderName.ToLower() == "header") && !"D1,N8,G9,H2,H4,T8,E8,L8,L7,N2,G2,D2,Q5,Q6,U6,Y6,Z8,Z5".Contains(WinSettings.MenuID))
				{
					StaticCell.Value = HTKKWinDictionaryDiaDanh.GetTenCoQuanThueQuanLy(value2);
				}
			}
			else
			{
				if (StaticCell.Controltype == "35")
				{
					Dictionary<string, string> danhMuc = WinSpreadProcess.GetDanhMuc("GetDMThueBVMT");
					if (!danhMuc.ContainsValue(value2))
					{
						return;
					}
					{
						foreach (KeyValuePair<string, string> item in danhMuc)
						{
							if (item.Value.Equals(value2))
							{
								StaticCell.Value = item.Key;
								break;
							}
						}
						return;
					}
				}
				if (StaticCell.Controltype == "11")
				{
					return;
				}
				if (StaticCell.Controltype == "29")
				{
					Dictionary<string, string> dictionary = Functions.getDictionary("Catalogue_DM_Lydo_GH_TT16.xml");
					if (!dictionary.ContainsValue(value2))
					{
						return;
					}
					{
						foreach (KeyValuePair<string, string> item2 in dictionary)
						{
							if (item2.Value.Equals(value2))
							{
								StaticCell.Value = item2.Key;
								break;
							}
						}
						return;
					}
				}
				if (!(StaticCell.Controltype == "25"))
				{
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void GetValueFromSelectedCell(HTKKCMCCellInfo StaticCell, HTKKCMCCellBasicInfo ChildCell)
	{
		string value = StaticCell.Value;
		string value2 = ChildCell.Value;
		try
		{
			if (!value2.IsNulOrEmpty() || value.IsNulOrEmpty() || StaticCell.Controltype == "56" || StaticCell.Controltype == "4" || StaticCell.Controltype == "67" || StaticCell.Controltype == "3" || StaticCell.Controltype == "5" || StaticCell.Controltype == "6" || StaticCell.Controltype == "7")
			{
			}
			if (value2.IsNulOrEmpty() || !value.IsNulOrEmpty())
			{
				return;
			}
			if (StaticCell.Controltype == "8")
			{
				if (!(WinSettings.ToKhaiInfo.FolderName.ToLower() == "header") && !"D1,N8,G9,H2,H4,T8,E8,L8,L7,N2,G2,D2,Q5,Q6,U6,Y6,Z8,Z5".Contains(WinSettings.MenuID))
				{
					StaticCell.Value = HTKKWinDictionaryDiaDanh.GetTenCoQuanThueCapCuc(value2, WinSettings.ToKhaiInfo.FolderName.ToLower() == "header");
				}
			}
			else if (StaticCell.Controltype == "9")
			{
				if (!(WinSettings.ToKhaiInfo.FolderName.ToLower() == "header") && !"D1,N8,G9,H2,H4,T8,E8,L8,L7,N2,G2,D2,Q5,Q6,U6,Y6,Z8,Z5".Contains(WinSettings.MenuID))
				{
					StaticCell.Value = HTKKWinDictionaryDiaDanh.GetTenCoQuanThueQuanLy(value2);
				}
			}
			else
			{
				if (StaticCell.Controltype == "35")
				{
					Dictionary<string, string> danhMuc = WinSpreadProcess.GetDanhMuc("GetDMThueBVMT");
					if (!danhMuc.ContainsValue(value2))
					{
						return;
					}
					{
						foreach (KeyValuePair<string, string> item in danhMuc)
						{
							if (item.Value.Equals(value2))
							{
								StaticCell.Value = item.Key;
								break;
							}
						}
						return;
					}
				}
				if (StaticCell.Controltype == "11")
				{
					return;
				}
				if (StaticCell.Controltype == "29")
				{
					Dictionary<string, string> dictionary = Functions.getDictionary("Catalogue_DM_Lydo_GH_TT16.xml");
					if (!dictionary.ContainsValue(value2))
					{
						return;
					}
					{
						foreach (KeyValuePair<string, string> item2 in dictionary)
						{
							if (item2.Value.Equals(value2))
							{
								StaticCell.Value = item2.Key;
								break;
							}
						}
						return;
					}
				}
				if (!(StaticCell.Controltype == "25"))
				{
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void GetValueFromSelectedCell(HTKKCMCRowValue tableCellRowValue, string strControltype, int idxValue, int idxSelectedValue)
	{
		try
		{
			Dictionary<string, string> dicPhuongXa;
			switch (strControltype)
			{
			case "8":
				if (!(WinSettings.ToKhaiInfo.NopTheoTinhCQT == "3") && !"D1,N8,G9,H2,H4,T8,E8,L8,L7,N2,G2,D2,Q5,Q6,U6,Y6,Z8,Z5".Contains(WinSettings.MenuID))
				{
					string tenCoQuanThueQuanLy = HTKKWinDictionaryDiaDanh.GetTenCoQuanThueCapCuc(tableCellRowValue.RowValue[idxSelectedValue], WinSettings.ToKhaiInfo.FolderName.ToLower() == "header");
					if (!tenCoQuanThueQuanLy.IsNulOrEmpty())
					{
						tableCellRowValue.RowValue[idxValue] = tenCoQuanThueQuanLy;
					}
				}
				return;
			case "9":
				if (!(WinSettings.ToKhaiInfo.NopTheoTinhCQT == "3") && !"D1,N8,G9,H2,H4,T8,E8,L8,L7,N2,G2,D2,Q5,Q6,U6,Y6,Z8,Z5".Contains(WinSettings.MenuID))
				{
					string tenCoQuanThueQuanLy = HTKKWinDictionaryDiaDanh.GetTenCoQuanThueQuanLy(tableCellRowValue.RowValue[idxSelectedValue]);
					if (!tenCoQuanThueQuanLy.IsNulOrEmpty())
					{
						tableCellRowValue.RowValue[idxValue] = tenCoQuanThueQuanLy;
					}
				}
				return;
			case "29":
				return;
			case "35":
				dicPhuongXa = WinSpreadProcess.GetDanhMuc("GetDMThueBVMT");
				if (!dicPhuongXa.ContainsKey(tableCellRowValue.RowValue[idxSelectedValue]))
				{
					return;
				}
				{
					foreach (KeyValuePair<string, string> item in dicPhuongXa)
					{
						if (item.Key.Equals(tableCellRowValue.RowValue[idxSelectedValue]))
						{
							tableCellRowValue.RowValue[idxValue] = item.Value;
							break;
						}
					}
					return;
				}
			case "11":
				return;
			case "28":
			{
				Dictionary<string, string> dictionary2 = WinSpreadProcess.DMGhiChuTemRuou();
				if (dictionary2 != null && dictionary2.Keys.Count > 0)
				{
					string text = tableCellRowValue.RowValue[idxSelectedValue];
					if (!string.IsNullOrEmpty(text) && dictionary2.ContainsKey(text))
					{
						tableCellRowValue.RowValue[idxValue] = dictionary2[text];
					}
				}
				return;
			}
			case "73":
			{
				Dictionary<string, string> dictionary5 = WinSpreadProcess.DMLoaiTienTe();
				if (dictionary5 != null && dictionary5.Keys.Count > 0)
				{
					string text = tableCellRowValue.RowValue[idxSelectedValue];
					if (!string.IsNullOrEmpty(text) && dictionary5.ContainsKey(text))
					{
						tableCellRowValue.RowValue[idxValue] = dictionary5[text];
					}
				}
				return;
			}
			case "80":
			{
				Dictionary<string, string> dictionary6 = WinSpreadProcess.DMThueSuat();
				if (dictionary6 != null && dictionary6.Keys.Count > 0)
				{
					string text = tableCellRowValue.RowValue[idxSelectedValue];
					if (!string.IsNullOrEmpty(text) && dictionary6.ContainsKey(text))
					{
						tableCellRowValue.RowValue[idxValue] = dictionary6[text];
					}
				}
				return;
			}
			case "72":
			{
				Dictionary<string, string> dictionary3 = WinSpreadProcess.DMMo();
				if (dictionary3 != null && dictionary3.Keys.Count > 0)
				{
					string text = tableCellRowValue.RowValue[idxSelectedValue];
					if (!string.IsNullOrEmpty(text) && dictionary3.ContainsKey(text))
					{
						tableCellRowValue.RowValue[idxValue] = dictionary3[text];
					}
				}
				return;
			}
			case "81":
			{
				Dictionary<string, string> dictionary4 = WinSpreadProcess.DMLoaiDat();
				if (dictionary4 != null && dictionary4.Keys.Count > 0)
				{
					string text = tableCellRowValue.RowValue[idxSelectedValue];
					if (!string.IsNullOrEmpty(text) && dictionary4.ContainsKey(text))
					{
						tableCellRowValue.RowValue[idxValue] = dictionary4[text];
					}
				}
				return;
			}
			case "82":
			{
				Dictionary<string, string> dictionary = WinSpreadProcess.DMHangDat();
				if (dictionary != null && dictionary.Keys.Count > 0)
				{
					string text = tableCellRowValue.RowValue[idxSelectedValue];
					if (!string.IsNullOrEmpty(text) && dictionary.ContainsKey(text))
					{
						tableCellRowValue.RowValue[idxValue] = dictionary[text];
					}
				}
				return;
			}
			case "22":
			{
				foreach (KeyValuePair<string, string> item2 in WinSpreadProcess.DMPhiLephi())
				{
					if (item2.Key == tableCellRowValue.RowValue[idxSelectedValue])
					{
						tableCellRowValue.RowValue[idxValue] = item2.Value;
						break;
					}
				}
				return;
			}
			case "88":
			{
				foreach (KeyValuePair<string, string> item3 in WinSpreadProcess.DMPhiLephiTt63())
				{
					if (item3.Key == tableCellRowValue.RowValue[idxSelectedValue])
					{
						tableCellRowValue.RowValue[idxValue] = item3.Value;
						break;
					}
				}
				return;
			}
			case "43":
			{
				foreach (KeyValuePair<string, string> item4 in WinSpreadProcess.GetDanhMuc("GetDMLoaiThongTinThayDoi"))
				{
					if (item4.Key == tableCellRowValue.RowValue[idxSelectedValue])
					{
						tableCellRowValue.RowValue[idxValue] = item4.Value;
						break;
					}
				}
				return;
			}
			case "58":
			{
				foreach (KeyValuePair<string, string> item5 in WinSpreadProcess.GetDanhMuc("GetDMLoaiThongTinThayDoiBLP"))
				{
					if (item5.Key == tableCellRowValue.RowValue[idxSelectedValue])
					{
						tableCellRowValue.RowValue[idxValue] = item5.Value;
						break;
					}
				}
				return;
			}
			case "20":
			{
				foreach (KeyValuePair<string, string> item6 in WinSpreadProcess.GetDanhMuc("GetDMKhoangSan"))
				{
					if (item6.Key == tableCellRowValue.RowValue[idxSelectedValue])
					{
						tableCellRowValue.RowValue[idxValue] = item6.Value;
						break;
					}
				}
				return;
			}
			case "62":
			{
				foreach (KeyValuePair<string, string> item7 in WinSpreadProcess.GetDanhMuc("GetDMKSBVMT"))
				{
					if (item7.Key == tableCellRowValue.RowValue[idxSelectedValue])
					{
						tableCellRowValue.RowValue[idxValue] = item7.Value;
						break;
					}
				}
				return;
			}
			case "53":
			{
				foreach (KeyValuePair<string, string> item8 in WinSpreadProcess.GetDMYKienKiemToan())
				{
					if (item8.Key == tableCellRowValue.RowValue[idxSelectedValue])
					{
						tableCellRowValue.RowValue[idxValue] = item8.Value;
						break;
					}
				}
				return;
			}
			case "90":
			{
				foreach (KeyValuePair<string, string> item9 in WinSpreadProcess.GetMaTieuMuc())
				{
					if (item9.Key == tableCellRowValue.RowValue[idxSelectedValue])
					{
						tableCellRowValue.RowValue[idxValue] = item9.Value;
						break;
					}
				}
				return;
			}
			case "91":
			{
				foreach (KeyValuePair<string, string> item10 in WinSpreadProcess.GetDMNganhCoThuNhapCap())
				{
					if (item10.Key == tableCellRowValue.RowValue[idxSelectedValue])
					{
						tableCellRowValue.RowValue[idxValue] = item10.Value;
						break;
					}
				}
				return;
			}
			case "92":
			{
				foreach (KeyValuePair<string, string> item11 in WinSpreadProcess.GetDMChuong())
				{
					if (item11.Key == tableCellRowValue.RowValue[idxSelectedValue])
					{
						tableCellRowValue.RowValue[idxValue] = item11.Value;
						break;
					}
				}
				return;
			}
			case "93":
			{
				foreach (KeyValuePair<string, string> item12 in WinSpreadProcess.GetDMNganHang())
				{
					if (item12.Key == tableCellRowValue.RowValue[idxSelectedValue])
					{
						tableCellRowValue.RowValue[idxValue] = item12.Value;
						break;
					}
				}
				return;
			}
			case "106":
			{
				foreach (KeyValuePair<string, string> item13 in WinSpreadProcess.GetDMDynamic("DM_NganHang_Moi.xml"))
				{
					if (item13.Key == tableCellRowValue.RowValue[idxSelectedValue])
					{
						tableCellRowValue.RowValue[idxValue] = item13.Value;
						break;
					}
				}
				return;
			}
			case "107":
			{
				foreach (KeyValuePair<string, string> item14 in WinSpreadProcess.GetDMDynamic("DM_Kho_Bac.xml"))
				{
					if (item14.Key == tableCellRowValue.RowValue[idxSelectedValue])
					{
						tableCellRowValue.RowValue[idxValue] = item14.Value;
						break;
					}
				}
				return;
			}
			case "108":
			{
				foreach (KeyValuePair<string, string> item15 in WinSpreadProcess.GetDMDynamic("DM_CQHQ.xml"))
				{
					if (item15.Key == tableCellRowValue.RowValue[idxSelectedValue])
					{
						tableCellRowValue.RowValue[idxValue] = item15.Value;
						break;
					}
				}
				return;
			}
			case "98":
			{
				foreach (KeyValuePair<string, string> item16 in WinSpreadProcess.GetDMNganHangKBNN())
				{
					if (item16.Key == tableCellRowValue.RowValue[idxSelectedValue])
					{
						tableCellRowValue.RowValue[idxValue] = item16.Value;
						break;
					}
				}
				return;
			}
			case "51":
			{
				foreach (KeyValuePair<string, string> item17 in HTKKWinDictionaryDiaDanh.KyPhatSinhHoanThue)
				{
					if (item17.Key == tableCellRowValue.RowValue[idxSelectedValue])
					{
						tableCellRowValue.RowValue[idxValue] = item17.Value;
						break;
					}
				}
				return;
			}
			case "39":
			{
				foreach (KeyValuePair<string, string> item18 in WinSpreadProcess.GetDanhMuc("GetDMThueTieuThuDB"))
				{
					if (item18.Key == tableCellRowValue.RowValue[idxSelectedValue])
					{
						tableCellRowValue.RowValue[idxValue] = item18.Value;
						break;
					}
				}
				return;
			}
			case "42":
			{
				foreach (KeyValuePair<string, string> item19 in WinSpreadProcess.GetDanhMuc("GetDichVuThueTieuThuDB"))
				{
					if (item19.Key == tableCellRowValue.RowValue[idxSelectedValue])
					{
						tableCellRowValue.RowValue[idxValue] = item19.Value;
						break;
					}
				}
				return;
			}
			case "36":
				dicPhuongXa = WinSpreadProcess.GetDanhMuc("GetDMThan");
				if (!dicPhuongXa.ContainsKey(tableCellRowValue.RowValue[idxSelectedValue]))
				{
					return;
				}
				{
					foreach (KeyValuePair<string, string> item20 in dicPhuongXa)
					{
						if (item20.Key.Equals(tableCellRowValue.RowValue[idxSelectedValue]))
						{
							tableCellRowValue.RowValue[idxValue] = item20.Value;
							break;
						}
					}
					return;
				}
			case "71":
				dicPhuongXa = WinSpreadProcess.GetDanhMuc("GetDMXangDau");
				if (!dicPhuongXa.ContainsKey(tableCellRowValue.RowValue[idxSelectedValue]))
				{
					return;
				}
				{
					foreach (KeyValuePair<string, string> item21 in dicPhuongXa)
					{
						if (item21.Key.Equals(tableCellRowValue.RowValue[idxSelectedValue]))
						{
							tableCellRowValue.RowValue[idxValue] = item21.Value;
							break;
						}
					}
					return;
				}
			case "3":
				dicPhuongXa = HTKKWinDictionaryDiaDanh.QuanHe;
				if (!dicPhuongXa.ContainsValue(tableCellRowValue.RowValue[idxSelectedValue]))
				{
					return;
				}
				{
					foreach (KeyValuePair<string, string> item22 in dicPhuongXa)
					{
						if (item22.Value.Equals(tableCellRowValue.RowValue[idxSelectedValue]))
						{
							tableCellRowValue.RowValue[idxValue] = item22.Key;
							break;
						}
					}
					return;
				}
			case "4":
				dicPhuongXa = HTKKWinDictionaryDiaDanh.QuocTich;
				if (!dicPhuongXa.ContainsValue(tableCellRowValue.RowValue[idxSelectedValue]))
				{
					return;
				}
				{
					foreach (KeyValuePair<string, string> item23 in dicPhuongXa)
					{
						if (item23.Value.Equals(tableCellRowValue.RowValue[idxSelectedValue]))
						{
							tableCellRowValue.RowValue[idxValue] = item23.Key;
							break;
						}
					}
					return;
				}
			case "5":
				dicPhuongXa = HTKKWinDictionaryDiaDanh.dicTinhTP;
				if (!dicPhuongXa.ContainsValue(tableCellRowValue.RowValue[idxSelectedValue]))
				{
					return;
				}
				{
					foreach (KeyValuePair<string, string> item24 in dicPhuongXa)
					{
						if (item24.Value.Equals(tableCellRowValue.RowValue[idxSelectedValue]))
						{
							tableCellRowValue.RowValue[idxValue] = item24.Key;
							break;
						}
					}
					return;
				}
			case "103":
				dicPhuongXa = HTKKWinDictionaryDiaDanh.dicThueTinh;
				if (!dicPhuongXa.ContainsValue(tableCellRowValue.RowValue[idxSelectedValue]))
				{
					return;
				}
				{
					foreach (KeyValuePair<string, string> item25 in dicPhuongXa)
					{
						if (item25.Value.Equals(tableCellRowValue.RowValue[idxSelectedValue]))
						{
							tableCellRowValue.RowValue[idxValue] = item25.Key;
							break;
						}
					}
					return;
				}
			case "100":
				dicPhuongXa = HTKKWinDictionaryDiaDanh.dicTinhTPMoi;
				if (!dicPhuongXa.ContainsValue(tableCellRowValue.RowValue[idxSelectedValue]))
				{
					return;
				}
				{
					foreach (KeyValuePair<string, string> item26 in dicPhuongXa)
					{
						if (item26.Value.Equals(tableCellRowValue.RowValue[idxSelectedValue]))
						{
							tableCellRowValue.RowValue[idxValue] = item26.Key;
							break;
						}
					}
					return;
				}
			case "6":
				dicPhuongXa = HTKKWinDictionaryDiaDanh.dicQuanHuyen;
				if (!dicPhuongXa.ContainsKey(tableCellRowValue.RowValue[idxSelectedValue]))
				{
					return;
				}
				{
					foreach (KeyValuePair<string, string> item27 in dicPhuongXa)
					{
						if (item27.Key.Equals(tableCellRowValue.RowValue[idxSelectedValue]))
						{
							tableCellRowValue.RowValue[idxValue] = item27.Value;
							break;
						}
					}
					return;
				}
			case "7":
				dicPhuongXa = HTKKWinDictionaryDiaDanh.dicPhuongXa;
				if (!dicPhuongXa.ContainsKey(tableCellRowValue.RowValue[idxSelectedValue]))
				{
					return;
				}
				{
					foreach (KeyValuePair<string, string> item28 in dicPhuongXa)
					{
						if (item28.Key.Equals(tableCellRowValue.RowValue[idxSelectedValue]))
						{
							tableCellRowValue.RowValue[idxValue] = item28.Value;
							break;
						}
					}
					return;
				}
			case "101":
				return;
			case null:
				return;
			}
			if (strControltype.Length <= 3 || !(strControltype.Substring(0, 3) == "DC_"))
			{
				return;
			}
			dicPhuongXa = WinSpreadProcess.GetDMDungChung(strControltype);
			if (tableCellRowValue.RowValue.Count > idxSelectedValue)
			{
				if (!dicPhuongXa.ContainsKey(tableCellRowValue.RowValue[idxSelectedValue]))
				{
					return;
				}
				{
					foreach (KeyValuePair<string, string> item29 in dicPhuongXa)
					{
						if (item29.Key.Equals(tableCellRowValue.RowValue[idxSelectedValue]))
						{
							tableCellRowValue.RowValue[idxValue] = item29.Value;
							break;
						}
					}
					return;
				}
			}
			HTKKFile.Log("kiem tra lai " + strControltype + " idxSelectedValue " + idxSelectedValue);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public static void GetValueFromChildCell(HTKKCMCRowValue tableCellRowValue, string strControltype, int idxCell, int idxChildCell, int idxSelectedCell)
	{
		try
		{
			switch (strControltype)
			{
			case "25":
				tableCellRowValue.RowValue[idxCell] = tableCellRowValue.RowValue[idxChildCell];
				break;
			case "8":
			{
				if (tableCellRowValue.RowValue[idxChildCell].IsNulOrEmpty())
				{
					break;
				}
				string key = HTKKWinDictionaryDiaDanh.CoQuanThueQuanLy.SingleOrDefault((KeyValuePair<string, string> p) => p.Value.Contains(tableCellRowValue.RowValue[idxChildCell])).Key;
				if (key != null)
				{
					if (key.Length > 3)
					{
						tableCellRowValue.RowValue[idxCell] = HTKKWinDictionaryDiaDanh.GetTenCoQuanThueCapCuc(key.Substring(0, 3) + "00", WinSettings.ToKhaiInfo.FolderName.ToLower() == "header");
					}
					break;
				}
				string text = tableCellRowValue.RowValue[idxChildCell];
				if (!string.IsNullOrEmpty(text))
				{
					string maCqtMoiByTenCqt = HTKKWinDictionaryDiaDanh.getMaCqtMoiByTenCqt(text);
					if (!string.IsNullOrEmpty(maCqtMoiByTenCqt))
					{
						string text2 = maCqtMoiByTenCqt.Substring(0, 3) + "00";
						tableCellRowValue.RowValue[idxCell] = HTKKWinDictionaryDiaDanh.getTenCqtMoiByMaCqt(text2);
						tableCellRowValue.RowValue[idxSelectedCell] = text2;
					}
				}
				break;
			}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public static void GetValueFromChildCell(HTKKCMCCellInfo cellinfo, string strControltype, string childValue)
	{
		try
		{
			if (strControltype != null && strControltype == "8" && !(WinSettings.ToKhaiInfo.FolderName.ToLower() == "header") && !"D1,N8,G9,H2,H4,T8,E8,L8,L7,N2,G2,D2,Q5,Q6,U6,Y6,Z8,Z5".Contains(WinSettings.MenuID))
			{
				childValue = ((childValue.Length >= 3) ? (childValue.Substring(0, 3) + "00") : string.Empty);
				string key = HTKKWinDictionaryDiaDanh.CoQuanThueCapCuc.FirstOrDefault((KeyValuePair<string, string> p) => p.Value == childValue).Key;
				cellinfo.Value = key;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public static void GetParentCodeFromChildCell(HTKKCMCCellInfo cellinfo, string strControltype, string childValue)
	{
		try
		{
			switch (strControltype)
			{
			case "9":
				childValue = ((childValue.Length >= 3) ? (childValue.Substring(0, 3) + "00") : string.Empty);
				if (string.IsNullOrEmpty(cellinfo.Value))
				{
					cellinfo.Value = childValue;
				}
				break;
			case "6":
				childValue = ((childValue.Length >= 3) ? childValue.Substring(0, 3) : string.Empty);
				if (string.IsNullOrEmpty(cellinfo.Value))
				{
					cellinfo.Value = childValue;
				}
				break;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void GetParentNameFromChildCell(HTKKCMCCellInfo cellinfo, string strControltype, string code)
	{
		try
		{
			string text = "";
			switch (strControltype)
			{
			case "9":
				text = HTKKWinDictionaryDiaDanh.CoQuanThueCapCuc.FirstOrDefault((KeyValuePair<string, string> p) => p.Value == code).Key;
				if (!string.IsNullOrEmpty(text))
				{
					text = text.Split(new char[1] { '-' })[1].Trim();
					if (string.IsNullOrEmpty(cellinfo.Value))
					{
						cellinfo.Value = text;
					}
				}
				break;
			case "6":
				code = ((code.Length >= 3) ? code.Substring(0, 3) : string.Empty);
				text = HTKKWinDictionaryDiaDanh.dicTinhTP.FirstOrDefault((KeyValuePair<string, string> p) => p.Value == code).Key;
				if (!string.IsNullOrEmpty(text) && string.IsNullOrEmpty(cellinfo.Value))
				{
					cellinfo.Value = text;
				}
				break;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void GetParentCodeFromChildCellInTable(HTKKCMCRowValue tableCellRowValue, string strControltype, int idxValue, string childenCode)
	{
		try
		{
			if (strControltype == "9")
			{
				string text = ((childenCode.Length >= 3) ? (childenCode.Substring(0, 3) + "00") : string.Empty);
				if (!text.IsNulOrEmpty() && string.IsNullOrEmpty(tableCellRowValue.RowValue[idxValue]))
				{
					tableCellRowValue.RowValue[idxValue] = text;
				}
			}
			if (strControltype == "6")
			{
				string text = ((childenCode.Length >= 3) ? childenCode.Substring(0, 3) : string.Empty);
				if (!text.IsNulOrEmpty() && string.IsNullOrEmpty(tableCellRowValue.RowValue[idxValue]))
				{
					tableCellRowValue.RowValue[idxValue] = text;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void GetParentNameFromChildCellInTable(HTKKCMCRowValue tableCellRowValue, string strControltype, int idxValue, string parentCode)
	{
		try
		{
			string text = "";
			if (strControltype != null && strControltype == "6")
			{
				parentCode = ((parentCode.Length >= 3) ? parentCode.Substring(0, 3) : string.Empty);
				text = HTKKWinDictionaryDiaDanh.dicTinhTP.FirstOrDefault((KeyValuePair<string, string> p) => p.Value == parentCode.Substring(0, 3)).Key;
				if (!string.IsNullOrEmpty(text))
				{
					tableCellRowValue.RowValue[idxValue] = text;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void WinDoWorkBackup(string pathFile)
	{
		//IL_0044: Unknown result type (might be due to invalid IL or missing references)
		try
		{
			using ZipFile zipFile = new ZipFile();
			zipFile.StatusMessageTextWriter = Console.Out;
			zipFile.AddDirectory(HTKKConfig.DATAFILES);
			zipFile.Save(pathFile);
			frmMessage frmMessage = new frmMessage("0007", "3", "1", "Thông báo");
			((Form)frmMessage).ShowDialog();
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void WinDoWorkRestore(string pathFile)
	{
		//IL_0032: Unknown result type (might be due to invalid IL or missing references)
		try
		{
			using ZipFile zipFile = ZipFile.Read(pathFile);
			zipFile.ExtractAll(HTKKConfig.DATAFILES, ExtractExistingFileAction.OverwriteSilently);
			frmMessage frmMessage = new frmMessage("0008", "3", "1", "Thông báo");
			((Form)frmMessage).ShowDialog();
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void ChangeValue(string value, string controlType)
	{
		try
		{
			if (controlType == "2")
			{
				value = ((!(value == "True")) ? "0" : "1");
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void ReChangeValue(XmlNode xNode, string controlType)
	{
		try
		{
			if (controlType == "2")
			{
				if (xNode.InnerXml == "1" || xNode.InnerXml.ToUpper() == "TRUE" || xNode.InnerXml.ToUpper() == "X")
				{
					xNode.InnerXml = "True";
				}
				else
				{
					xNode.InnerXml = "False";
				}
			}
			else if (controlType == "14")
			{
				DateTime result = default(DateTime);
				if (DateTime.TryParseExact(xNode.InnerXml, "yyyy-MM-dd", CultureInfo.CurrentCulture, DateTimeStyles.None, out result))
				{
					xNode.InnerXml = result.ToString("dd/MM/yyyy");
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void WinDoWorkSheetSave(SheetView sView, HTKKCMCSheetInfo sInfo)
	{
		//IL_04a9: Unknown result type (might be due to invalid IL or missing references)
		//IL_04b0: Expected O, but got Unknown
		//IL_0b25: Unknown result type (might be due to invalid IL or missing references)
		//IL_0b2c: Expected O, but got Unknown
		try
		{
			string text = "";
			text = (WinSettings.ToKhaiInfo.MenuID.Equals("100_1") ? (WinSettings.ToKhaiInfo.InterfaceTemplateXml + sInfo.DataFile + ".xml") : ((WinSettings.TTPeriodInfo.LoaiToKhai.Equals("C") || sInfo.SheetName.Equals("Tờ khai")) ? (WinSettings.ToKhaiInfo.InterfaceTemplateXml + sInfo.DataFile + "_" + WinSettings.ToKhaiInfo.VersionXML.Replace(".", "") + ".xml") : ((WinSettings.ToKhaiInfo.KieuKHBS == "3") ? (sInfo.SheetID.Equals("01-KHBS") ? ("I7,I4,C7,I8,C8".Contains(WinSettings.MenuID) ? (WinSettings.ToKhaiInfo.InterfaceTemplateXml + "KHBS_01_TT80_DVT_" + WinSettings.ToKhaiInfo.VersionKBSXML.Replace(".", "") + ".xml") : ((!"D1,N8,G9,H2,H4,T8,E8,L8,L7,N2,G2,R5,Y6,Z8,Z5".Contains(WinSettings.MenuID)) ? (WinSettings.ToKhaiInfo.InterfaceTemplateXml + "KHBS_01_TT80_" + WinSettings.ToKhaiInfo.VersionKBSXML.Replace(".", "") + ".xml") : (WinSettings.ToKhaiInfo.InterfaceTemplateXml + "KHBS_01_TT80_DIA_BAN_" + WinSettings.ToKhaiInfo.VersionKBSXML.Replace(".", "") + ".xml"))) : ((!sInfo.SheetID.Equals("01-1-KHBS")) ? (WinSettings.ToKhaiInfo.InterfaceTemplateXml + sInfo.DataFile + "_" + WinSettings.ToKhaiInfo.VersionXML.Replace(".", "") + ".xml") : ((!"I7,I4,C7,I8,C8".Contains(WinSettings.MenuID)) ? (WinSettings.ToKhaiInfo.InterfaceTemplateXml + "KHBS_01_1_TT80_" + WinSettings.ToKhaiInfo.VersionKBSXML.Replace(".", "") + ".xml") : (WinSettings.ToKhaiInfo.InterfaceTemplateXml + "KHBS_01_1_TT80_DVT_" + WinSettings.ToKhaiInfo.VersionKBSXML.Replace(".", "") + ".xml")))) : ((!"30,I5,K9,K1,M7,K6,K8,73,55,71,31,P2,X6".Contains(WinSettings.MenuID) || sInfo.SheetName.Equals("KHBS")) ? (WinSettings.ToKhaiInfo.InterfaceTemplateXml + sInfo.DataFile + "_" + WinSettings.ToKhaiInfo.VersionKBSXML.Replace(".", "") + ".xml") : (WinSettings.ToKhaiInfo.InterfaceTemplateXml + sInfo.DataFile + "_" + WinSettings.ToKhaiInfo.VersionXML.Replace(".", "") + ".xml")))));
			HTKKCMCXML hTKKCMCXML = new HTKKCMCXML(text);
			string sheetName = sInfo.SheetName;
			XmlNodeList childNodes = ((XmlNode)hTKKCMCXML.XDocument).SelectSingleNode("/Sections").ChildNodes;
			List<XmlNode> list = new List<XmlNode>();
			for (int i = 0; i < childNodes.Count; i++)
			{
				list.Add(childNodes[i]);
			}
			for (int i = 0; i < list.Count; i++)
			{
				try
				{
					XmlNode xmlTableNode = list[i];
					if (!(xmlTableNode.Name == "Section"))
					{
						continue;
					}
					if (((XmlNode)xmlTableNode.Attributes["Dynamic"]).Value == "0")
					{
						XmlNodeList childNodes2 = xmlTableNode.SelectSingleNode("Cells").ChildNodes;
						foreach (XmlNode item in childNodes2)
						{
							XmlNode val = item;
							try
							{
								if (val.Attributes["Path"] == null)
								{
									continue;
								}
								string value = ((XmlNode)val.Attributes["Path"]).Value;
								for (int j = 0; j < sInfo.StaticCells.Count; j++)
								{
									try
									{
										if (value == sInfo.StaticCells[j].Path)
										{
											string valueCustom = GeneralFunctions.GetValueCustom(sView, sInfo.StaticCells[j].CellTemplate.RowIndex, sInfo.StaticCells[j].CellTemplate.ColIndex, sInfo.StaticCells[j].Controltype);
											((XmlNode)val.Attributes["CellID"]).Value = HTKKFile.ColumnIndexToLetters(sInfo.StaticCells[j].CellTemplate.ColIndex) + "_" + sInfo.StaticCells[j].CellTemplate.RowIndex;
											if (sInfo.StaticCells[j].CellReport != null)
											{
												((XmlNode)val.Attributes["CellID2"]).Value = HTKKFile.ColumnIndexToLetters(sInfo.StaticCells[j].CellReport.ColIndex) + "_" + sInfo.StaticCells[j].CellReport.RowIndex;
											}
											if (sInfo.StaticCells[j].ChildCell != null)
											{
												((XmlNode)val.Attributes["ChildCell"]).Value = HTKKFile.ColumnIndexToLetters(sInfo.StaticCells[j].ChildCell.ColIndex) + "_" + sInfo.StaticCells[j].ChildCell.RowIndex;
											}
											if (sInfo.StaticCells[j].CellSelectedValue != null)
											{
												((XmlNode)val.Attributes["SelectedValue"]).Value = HTKKFile.ColumnIndexToLetters(sInfo.StaticCells[j].CellSelectedValue.ColIndex) + "_" + sInfo.StaticCells[j].CellSelectedValue.RowIndex;
											}
											if (sInfo.StaticCells[j].ParentCell != null)
											{
												((XmlNode)val.Attributes["ParentCell"]).Value = HTKKFile.ColumnIndexToLetters(sInfo.StaticCells[j].ParentCell.ColIndex) + "_" + sInfo.StaticCells[j].ParentCell.RowIndex;
											}
											if (!HTKKFormater.isNullString(valueCustom))
											{
												((XmlNode)val.Attributes["Value"]).Value = valueCustom;
												break;
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
						continue;
					}
					string value2 = ((XmlNode)xmlTableNode.Attributes["id"]).Value;
					string value3 = ((XmlNode)xmlTableNode.Attributes["TableName"]).Value;
					HTKKCMCTableInfo table = sInfo.GetTable(value2, value3);
					HTKKCMCTableInfo hTKKCMCTableInfo = table.Clone();
					if (hTKKCMCTableInfo.CellTableSTT != null || hTKKCMCTableInfo.Header.StaticCells.Count > 0 || hTKKCMCTableInfo.Footer.StaticCells.Count > 0)
					{
						XmlNode val2 = xmlTableNode.SelectSingleNode("Location");
						if (val2 != null)
						{
							XmlNode val3 = val2.SelectSingleNode("DataLocationTemplateRowStart");
							if (val3 != null)
							{
								hTKKCMCTableInfo.Location.DataLocationTemplateRowStart = ConvertHelpers.ConvertToInt(val3.InnerText.Trim());
							}
						}
					}
					XmlNode xmlTableTemplateNode = xmlTableNode.Clone();
					if (table.CellTableSTT != null || hTKKCMCTableInfo.Header.StaticCells.Count > 0 || hTKKCMCTableInfo.Footer.StaticCells.Count > 0)
					{
						hTKKCMCXML.BindDataTableToTableNode(sView, xmlTableNode, table, hTKKCMCTableInfo);
					}
					else
					{
						hTKKCMCXML.BindDataTableToTableNode(sView, xmlTableNode, table, table);
					}
					for (int j = 0; j < sInfo.Tables.Count; j++)
					{
						if ((value3 == sInfo.Tables[j].TableName) & (value2 != sInfo.Tables[j].TableID))
						{
							hTKKCMCXML.InsertTableNode(sView, ref xmlTableNode, xmlTableTemplateNode, sInfo.Tables[j], hTKKCMCTableInfo);
						}
					}
				}
				catch (Exception ex)
				{
					HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
				}
			}
			hTKKCMCXML.XDocument.DocumentElement.SetAttribute("Version", WinSettings.ToKhaiInfo.Version);
			string text2 = "";
			text2 = ((!WinSettings.ToKhaiInfo.MenuID.Equals("100_1")) ? GetFileNameBySheet(WinSettings.TTPeriodInfo, sInfo.DataFile, HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst + "\\", WinSettings.ToKhaiInfo.VersionXML) : (HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst + "\\" + sInfo.DataFile + ".xml"));
			XmlTextWriter val4 = new XmlTextWriter(text2, Encoding.UTF8);
			try
			{
				val4.Formatting = (Formatting)1;
				hTKKCMCXML.XDocument.Save((XmlWriter)(object)val4);
			}
			finally
			{
				((IDisposable)val4)?.Dispose();
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public static string WinDoWorkGetBarcodeFromSheet(SheetView sView, HTKKCMCSheetInfo sInfo)
	{
		//IL_035e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0365: Expected O, but got Unknown
		try
		{
			string text = "<S" + sInfo.SheetID + ">";
			string pathFile = string.Format("{0}{1}_{2}.xml", WinSettings.ToKhaiInfo.InterfaceTemplateXml, sInfo.DataFile, WinSettings.ToKhaiInfo.VersionXML.Replace(".", ""));
			if (WinSettings.ToKhaiInfo.KieuKHBS == "3")
			{
				if (sInfo.SheetID == "01-1-KHBS")
				{
					pathFile = ((!"I7,I4,C7,I8,C8".Contains(WinSettings.MenuID)) ? string.Format("{0}{1}_{2}.xml", WinSettings.ToKhaiInfo.InterfaceTemplateXml, "KHBS_01_1_TT80", WinSettings.ToKhaiInfo.VersionKBSXML.Replace(".", "")) : string.Format("{0}{1}_{2}.xml", WinSettings.ToKhaiInfo.InterfaceTemplateXml, "KHBS_01_1_TT80_DVT", WinSettings.ToKhaiInfo.VersionKBSXML.Replace(".", "")));
				}
				else if (sInfo.SheetID == "01-KHBS")
				{
					pathFile = ("I7,I4,C7,I8,C8".Contains(WinSettings.MenuID) ? string.Format("{0}{1}_{2}.xml", WinSettings.ToKhaiInfo.InterfaceTemplateXml, "KHBS_01_TT80_DVT", WinSettings.ToKhaiInfo.VersionKBSXML.Replace(".", "")) : ((!"D1,N8,G9,H2,H4,T8,E8,L8,L7,N2,G2,R5,Y6,Z8,Z5".Contains(WinSettings.MenuID)) ? string.Format("{0}{1}_{2}.xml", WinSettings.ToKhaiInfo.InterfaceTemplateXml, "KHBS_01_TT80", WinSettings.ToKhaiInfo.VersionKBSXML.Replace(".", "")) : string.Format("{0}{1}_{2}.xml", WinSettings.ToKhaiInfo.InterfaceTemplateXml, "KHBS_01_TT80_DIA_BAN", WinSettings.ToKhaiInfo.VersionKBSXML.Replace(".", ""))));
				}
			}
			else if (sInfo.DataFile.IndexOf("KHBS") > -1)
			{
				pathFile = string.Format("{0}{1}_{2}.xml", WinSettings.ToKhaiInfo.InterfaceTemplateXml, sInfo.DataFile, WinSettings.ToKhaiInfo.VersionKBSXML.Replace(".", ""));
			}
			HTKKCMCXML hTKKCMCXML = new HTKKCMCXML(pathFile);
			string sheetName = sInfo.SheetName;
			XmlNodeList childNodes = ((XmlNode)hTKKCMCXML.XDocument).SelectSingleNode("/Sections").ChildNodes;
			List<XmlNode> list = new List<XmlNode>();
			for (int i = 0; i < childNodes.Count; i++)
			{
				list.Add(childNodes[i]);
			}
			for (int i = 0; i < list.Count; i++)
			{
				XmlNode val = list[i];
				if (!(val.Name == "Section"))
				{
					continue;
				}
				if (((XmlNode)val.Attributes["Dynamic"]).Value == "0")
				{
					XmlNodeList val2 = val.SelectSingleNode("Cells").SelectNodes("Cell[@Encode='1']");
					if (val2.Count <= 0)
					{
						continue;
					}
					text += "<S>";
					foreach (XmlNode item in val2)
					{
						XmlNode val3 = item;
						string strCellPath = ((val3.Attributes["Path"] != null) ? ((XmlNode)val3.Attributes["Path"]).Value : "");
						string text2 = "";
						if (!string.IsNullOrEmpty(strCellPath))
						{
							HTKKCMCCellInfo hTKKCMCCellInfo = sInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path != null && cell.Path.Equals(strCellPath));
							if (hTKKCMCCellInfo != null)
							{
								text2 = GeneralFunctions.GetValueCustomMV(sView, hTKKCMCCellInfo.CellTemplate.RowIndex, hTKKCMCCellInfo.CellTemplate.ColIndex, hTKKCMCCellInfo.Controltype) + "~";
							}
						}
						text += text2;
					}
					if (text.Length > 0 && text[text.Length - 1] == '~')
					{
						text = text.Substring(0, text.Length - 1);
					}
					text += "</S>";
				}
				else
				{
					string value = ((XmlNode)val.Attributes["id"]).Value;
					string value2 = ((XmlNode)val.Attributes["TableName"]).Value;
					HTKKCMCTableInfo table = sInfo.GetTable(value, value2);
					text = text + "<S>" + GetBarcodeFromTable(sView, val, table, table, sInfo) + "</S>";
				}
			}
			return text + "</S" + sInfo.SheetID + ">";
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public static string GetBarcodeFromTable(SheetView sView, XmlNode xmlTableNode, HTKKCMCTableInfo tableInfo, HTKKCMCTableInfo rootTableInfo, HTKKCMCSheetInfo sInfo)
	{
		try
		{
			string text = string.Empty;
			if (xmlTableNode.SelectSingleNode("Header") != null)
			{
				XmlNodeList val = xmlTableNode.SelectSingleNode("Header").SelectNodes("Cell[@Encode='1']");
				if (val != null)
				{
					for (int i = 0; i < val.Count; i++)
					{
						string strCellPath = ((val[i].Attributes["Path"] != null) ? ((XmlNode)val[i].Attributes["Path"]).Value : "");
						string text2 = "";
						if (!string.IsNullOrEmpty(strCellPath))
						{
							HTKKCMCCellInfo hTKKCMCCellInfo = tableInfo.Header.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path != null && cell.Path.Equals(strCellPath));
							if (hTKKCMCCellInfo != null)
							{
								string value = hTKKCMCCellInfo.Value;
								text2 = ((!string.IsNullOrEmpty(value)) ? (value + "~") : "~");
							}
						}
						text += text2;
					}
				}
			}
			XmlNodeList val2 = xmlTableNode.SelectNodes("RowInfo/Cell[@Encode='1']");
			List<HTKKCMCRowValue> rowsValues = tableInfo.RowsValues;
			for (int i = 0; i < tableInfo.RowsValues.Count; i++)
			{
				List<string> rowValue = tableInfo.RowsValues[i].RowValue;
				if (WinSettings.MenuID == "57" && !ConvertHelpers.ConvertToBool(rowValue[0]))
				{
					continue;
				}
				for (int num = 0; num < val2.Count; num++)
				{
					string strCellPath2 = ((val2[num].Attributes["Path"] != null) ? ((XmlNode)val2[num].Attributes["Path"]).Value : "");
					string text2 = "";
					if (!string.IsNullOrEmpty(strCellPath2))
					{
						HTKKCMCTableCellInfo hTKKCMCTableCellInfo = tableInfo.RowsInfo.Cells.Find((HTKKCMCTableCellInfo cell) => cell.Path != null && cell.Path.Equals(strCellPath2));
						text2 = GeneralFunctions.GetValueCustom(sView, tableInfo.Location.DataLocationTemplateRowStart + i, hTKKCMCTableCellInfo.CellTemplate.ColIndex, hTKKCMCTableCellInfo.Controltype) + "~";
					}
					text += text2;
				}
			}
			if (xmlTableNode.SelectSingleNode("Footer") != null)
			{
				XmlNodeList val3 = xmlTableNode.SelectSingleNode("Footer").SelectNodes("Cell[@Encode='1']");
				if (val3 != null)
				{
					for (int i = 0; i < val3.Count; i++)
					{
						string strCellPath3 = ((val3[i].Attributes["Path"] != null) ? ((XmlNode)val3[i].Attributes["Path"]).Value : "");
						string text2 = "";
						if (!string.IsNullOrEmpty(strCellPath3))
						{
							HTKKCMCTableCellInfo hTKKCMCTableCellInfo = tableInfo.RowsInfo.Cells.Find((HTKKCMCTableCellInfo cell) => cell.Path != null && cell.Path.Equals(strCellPath3));
							text2 = GeneralFunctions.GetValueCustom(sView, hTKKCMCTableCellInfo.CellTemplate.ColIndex, hTKKCMCTableCellInfo.CellTemplate.ColIndex, hTKKCMCTableCellInfo.Controltype) + "~";
						}
						text += text2;
					}
				}
			}
			if (text.Contains("True") || text.Contains("False"))
			{
				text = text.Replace("True", "x").Replace("False", "");
			}
			return (text.Length - 1 > 0) ? text.Substring(0, text.Length - 1) : string.Empty;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return string.Empty;
		}
	}

	public static void WinDoWorkSignXml()
	{
		//IL_004f: Unknown result type (might be due to invalid IL or missing references)
		//IL_007b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0081: Expected O, but got Unknown
		//IL_0073: Unknown result type (might be due to invalid IL or missing references)
		//IL_00a2: Unknown result type (might be due to invalid IL or missing references)
		//IL_00a8: Invalid comparison between Unknown and I4
		//IL_00e6: Unknown result type (might be due to invalid IL or missing references)
		//IL_00ed: Expected O, but got Unknown
		//IL_026c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0152: Unknown result type (might be due to invalid IL or missing references)
		//IL_0159: Expected O, but got Unknown
		//IL_01ac: Unknown result type (might be due to invalid IL or missing references)
		//IL_01b3: Expected O, but got Unknown
		//IL_01e8: Unknown result type (might be due to invalid IL or missing references)
		//IL_01ef: Expected O, but got Unknown
		//IL_022c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0232: Invalid comparison between Unknown and I4
		//IL_0147: Unknown result type (might be due to invalid IL or missing references)
		//IL_0257: Unknown result type (might be due to invalid IL or missing references)
		try
		{
			while (true)
			{
				X509Certificate2 certificate = CA_FPT.GetCertificate();
				if (certificate == null)
				{
					break;
				}
				if (string.IsNullOrEmpty(WinSettings.HTKKHeaderInfo.mst) | (WinSettings.HTKKHeaderInfo.mst != CA_FPT.GetMaSoThueOfCA(certificate)))
				{
					MessageBox.Show("Mã số thuế đăng nhập khác với mã số thuế của CA. Hãy kiểm tra lại.", "Thông báo");
					continue;
				}
				if (!CA_FPT.CheckValidDate(certificate, DateTime.Now))
				{
					MessageBox.Show("CA đã hết hiệu lực. Hãy kiểm tra lại.", "Thông báo");
					continue;
				}
				OpenFileDialog val = new OpenFileDialog();
				((FileDialog)val).Filter = "Xml file (.xml) |*.xml";
				val.Multiselect = false;
				((FileDialog)val).Title = "Chọn file cần ký";
				while ((int)((CommonDialog)val).ShowDialog() == 1)
				{
					XmlDsigEnveloping xmlDsigEnveloping = new XmlDsigEnveloping(c14: true);
					string fileName = ((FileDialog)val).FileName;
					string text = fileName.Substring(fileName.LastIndexOf('\\') + 1).Replace(".xml", "");
					XmlDocument val2 = new XmlDocument();
					val2.Load(fileName);
					XmlNodeList elementsByTagName = val2.GetElementsByTagName("HSoKhaiThue");
					if (elementsByTagName.Count > 0)
					{
						XmlNodeList elementsByTagName2 = val2.GetElementsByTagName("CKyDTu");
						if (elementsByTagName2.Count <= 0)
						{
							XmlDocument val3 = new XmlDocument();
							XmlAttribute val4 = val2.CreateAttribute("id");
							((XmlNode)val4).Value = "_NODE_TO_SIGN";
							XmlNode val5 = elementsByTagName[0];
							val5.Attributes.Append(val4);
							val3.LoadXml(val5.OuterXml);
							string text2 = xmlDsigEnveloping.SignXml(val3, certificate, "CKyDTu");
							XmlDocument val6 = new XmlDocument();
							val6.LoadXml(text2);
							XmlNode val7 = val2.ImportNode(val6.GetElementsByTagName("CKyDTu")[0], true);
							((XmlNode)val2.DocumentElement).AppendChild(val7);
							SaveFileDialog val8 = new SaveFileDialog();
							((FileDialog)val8).Title = "Lưu file";
							((FileDialog)val8).DefaultExt = "xml";
							((FileDialog)val8).Filter = "Xml file (.xml) |*.xml";
							((FileDialog)val8).FileName = text + "_signed.xml";
							if ((int)((CommonDialog)val8).ShowDialog() == 1)
							{
								val2.Save(((FileDialog)val8).FileName);
								MessageBox.Show("Ký điện tử thành công.", "Thông báo");
							}
							break;
						}
						MessageBox.Show("File bạn chọn là file đã ký. Hãy chọn lại.", "Thông báo");
					}
					else
					{
						MessageBox.Show("File xml dữ liệu không đúng cấu trúc. Hãy chọn lại.", "Thông báo");
					}
				}
				break;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void SetFPSpreadExpExcel(FpSpread winFpForm, FpSpread winFpExp, List<HTKKCMCSheetInfo> sheetInfos)
	{
		try
		{
			for (int i = 0; i < sheetInfos.Count; i++)
			{
				if (sheetInfos[i] == null)
				{
					continue;
				}
				SheetView sheetView = winFpExp.Sheets[i];
				sheetView.Protect = false;
				List<HTKKCMCCellInfo> staticCells = sheetInfos[i].StaticCells;
				for (int j = 0; j < staticCells.Count; j++)
				{
					SetCellExportExcel(sheetView, staticCells[j].CellTemplate.RowIndex, staticCells[j].CellTemplate.ColIndex, staticCells[j].Controltype, staticCells[j].Value, staticCells[j].DefaultValue);
				}
				List<HTKKCMCTableInfo> tables = sheetInfos[i].Tables;
				sheetView.FpSpread.SuspendLayout();
				sheetView.AutoCalculation = false;
				for (int j = 0; j < tables.Count; j++)
				{
					for (int k = 0; k < tables[j].Header.StaticCells.Count; k++)
					{
						SetCellExportExcel(sheetView, tables[j].Header.StaticCells[k].CellTemplate.RowIndex, tables[j].Header.StaticCells[k].CellTemplate.ColIndex, tables[j].Header.StaticCells[k].Controltype, tables[j].Header.StaticCells[k].Value, tables[j].Header.StaticCells[k].DefaultValue);
					}
					HTKKCMCSheetInfo hTKKCMCSheetInfo = sheetInfos[i];
					if ("05-1BK-QTT-TNCN33;05-2BK-QTT-TNCN33;05-3BK-QTT-TNCN33;05-DS-TNCN39;BC26/AC68;02TH95;01-1/BK-XSBHĐC31".Split(new char[1] { ';' }).Contains(hTKKCMCSheetInfo.SheetName + WinSettings.ToKhaiInfo.MenuID) || WinSettings.ToKhaiInfo.SheetInfos[hTKKCMCSheetInfo.SheetIndex].TaiDuLieuLon == "1")
					{
						WinSpreadProcess.BindDataLargeExport(sheetView, hTKKCMCSheetInfo, j, null);
					}
					else
					{
						for (int k = 0; k < tables[j].RowsValues.Count; k++)
						{
							for (int l = 0; l < tables[j].RowsInfo.Cells.Count; l++)
							{
								SetCellExportExcel(sheetView, tables[j].Location.DataLocationTemplateRowStart + k, tables[j].RowsInfo.Cells[l].CellTemplate.ColIndex, tables[j].RowsInfo.Cells[l].Controltype, tables[j].RowsValues[k].RowValue[l], tables[j].RowsInfo.Cells[l].DefaultValue);
							}
						}
					}
					for (int k = 0; k < tables[j].Footer.StaticCells.Count; k++)
					{
						SetCellExportExcel(sheetView, tables[j].Footer.StaticCells[k].CellTemplate.RowIndex, tables[j].Footer.StaticCells[k].CellTemplate.ColIndex, tables[j].Footer.StaticCells[k].Controltype, tables[j].Footer.StaticCells[k].Value, tables[j].Footer.StaticCells[k].DefaultValue);
					}
				}
			}
			MethodInfo method = GetMethod("SetFPSpreadExp");
			if ((object)method != null)
			{
				method.Invoke(parameters: new object[6]
				{
					winFpExp,
					sheetInfos,
					WinSettings.TTPeriodInfo,
					WinSettings.ToKhaiInfo,
					WinSettings.TTinChung,
					WinSettings.HTKKHeaderInfo
				}, obj: WinSettings.InstanceClass);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void SetFPSpreadExpExcelAfterExport(FpSpread winFpForm, FpSpread winFpExp, List<HTKKCMCSheetInfo> sheetInfos)
	{
		try
		{
			for (int i = 0; i < sheetInfos.Count; i++)
			{
				if (sheetInfos[i] == null)
				{
					continue;
				}
				SheetView sheetView = winFpExp.Sheets[i];
				sheetView.Protect = true;
				sheetView.FpSpread.SuspendLayout();
				sheetView.AutoCalculation = false;
				List<HTKKCMCCellInfo> staticCells = sheetInfos[i].StaticCells;
				for (int j = 0; j < staticCells.Count; j++)
				{
					SetCellAfterExportExcel(sheetView, staticCells[j].CellTemplate.RowIndex, staticCells[j].CellTemplate.ColIndex, staticCells[j].Controltype, staticCells[j].Value, staticCells[j].DefaultValue);
				}
				List<HTKKCMCTableInfo> tables = sheetInfos[i].Tables;
				for (int j = 0; j < tables.Count; j++)
				{
					for (int k = 0; k < tables[j].Header.StaticCells.Count; k++)
					{
						SetCellAfterExportExcel(sheetView, tables[j].Header.StaticCells[k].CellTemplate.RowIndex, tables[j].Header.StaticCells[k].CellTemplate.ColIndex, tables[j].Header.StaticCells[k].Controltype, tables[j].Header.StaticCells[k].Value, tables[j].Header.StaticCells[k].DefaultValue);
					}
					HTKKCMCSheetInfo hTKKCMCSheetInfo = sheetInfos[i];
					if ("05-1BK-QTT-TNCN33;05-2BK-QTT-TNCN33;05-3BK-QTT-TNCN33;05-DS-TNCN39;BC26/AC68;02TH95;01-1/BK-XSBHĐC31".Split(new char[1] { ';' }).Contains(hTKKCMCSheetInfo.SheetName + WinSettings.ToKhaiInfo.MenuID) || WinSettings.ToKhaiInfo.SheetInfos[hTKKCMCSheetInfo.SheetIndex].TaiDuLieuLon == "1")
					{
						WinSpreadProcess.BindDataLargeAfterExport(sheetView, hTKKCMCSheetInfo, j, null);
					}
					else
					{
						for (int k = 0; k < tables[j].RowsValues.Count; k++)
						{
							for (int l = 0; l < tables[j].RowsInfo.Cells.Count; l++)
							{
								SetCellAfterExportExcel(sheetView, tables[j].Location.DataLocationTemplateRowStart + k, tables[j].RowsInfo.Cells[l].CellTemplate.ColIndex, tables[j].RowsInfo.Cells[l].Controltype, tables[j].RowsValues[k].RowValue[l], tables[j].RowsInfo.Cells[l].DefaultValue);
							}
						}
					}
					for (int k = 0; k < tables[j].Footer.StaticCells.Count; k++)
					{
						SetCellAfterExportExcel(sheetView, tables[j].Footer.StaticCells[k].CellTemplate.RowIndex, tables[j].Footer.StaticCells[k].CellTemplate.ColIndex, tables[j].Footer.StaticCells[k].Controltype, tables[j].Footer.StaticCells[k].Value, tables[j].Footer.StaticCells[k].DefaultValue);
					}
				}
				sheetView.FpSpread.ResumeLayout();
				sheetView.AutoCalculation = true;
			}
			MethodInfo method = GetMethod("SetFPSpreadExp");
			if ((object)method != null)
			{
				method.Invoke(parameters: new object[6]
				{
					winFpExp,
					sheetInfos,
					WinSettings.TTPeriodInfo,
					WinSettings.ToKhaiInfo,
					WinSettings.TTinChung,
					WinSettings.HTKKHeaderInfo
				}, obj: WinSettings.InstanceClass);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static List<HTKKSheetInMenu> getAllSheetInMenu(XmlNodeList nodeSheet)
	{
		List<HTKKSheetInMenu> list = new List<HTKKSheetInMenu>();
		for (int i = 0; i < nodeSheet.Count; i++)
		{
			HTKKSheetInMenu hTKKSheetInMenu = new HTKKSheetInMenu();
			hTKKSheetInMenu.stt = i;
			hTKKSheetInMenu.dataFile = ((nodeSheet[i].Attributes["ID"] == null) ? "" : ((XmlNode)nodeSheet[i].Attributes["ID"]).Value);
			hTKKSheetInMenu.sheetName = ((nodeSheet[i].Attributes["SheetName"] == null) ? "" : ((XmlNode)nodeSheet[i].Attributes["SheetName"]).Value);
			hTKKSheetInMenu.Active = ((nodeSheet[i].Attributes["Active"] == null) ? "" : ((XmlNode)nodeSheet[i].Attributes["Active"]).Value);
			hTKKSheetInMenu.dataFile = ((nodeSheet[i].Attributes["DataFile"] == null) ? "" : ((XmlNode)nodeSheet[i].Attributes["DataFile"]).Value);
			hTKKSheetInMenu.version = ((nodeSheet[i].ParentNode.Attributes["VersionXML"] == null) ? "" : ((XmlNode)nodeSheet[i].ParentNode.Attributes["VersionXML"]).Value);
			list.Add(hTKKSheetInMenu);
		}
		return list;
	}

	public static void CreateButton(FpSpread winFpSpread, int sheetIndex)
	{
		try
		{
			MethodInfo method = GetMethod("CreateButton_" + sheetIndex);
			if ((object)method == null)
			{
				return;
			}
			ParameterInfo[] parameters = method.GetParameters();
			object[] array = new object[parameters.Length];
			for (int i = 0; i < parameters.Length; i++)
			{
				switch (parameters[i].Name)
				{
				case "winFpSpread":
					array[i] = winFpSpread;
					break;
				case "ToKhaiInfo":
					array[i] = WinSettings.ToKhaiInfo;
					break;
				case "TTinChung":
					array[i] = WinSettings.TTinChung;
					break;
				case "PeriodInfo":
					array[i] = WinSettings.TTPeriodInfo;
					break;
				case "HeaderInfo":
					array[i] = WinSettings.HTKKHeaderInfo;
					break;
				case "sView":
					array[i] = winFpSpread.Sheets[sheetIndex];
					break;
				}
			}
			method.Invoke(WinSettings.InstanceClass, array);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static bool event_Leave(FpSpread axfpsPhuLuc, List<CheckBox> listCheckBox, List<MaskedTextBox> listMaskedTextBox, List<ComboBox> listCombobox, List<RadioButton> listRadioButton, string strFunction)
	{
		bool flag = false;
		try
		{
			MethodInfo method = GetMethod(strFunction);
			if ((object)method == null)
			{
				return true;
			}
			flag = (bool)method.Invoke(parameters: new object[5] { axfpsPhuLuc, listCheckBox, listMaskedTextBox, listCombobox, listRadioButton }, obj: WinSettings.InstanceClass);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			flag = false;
		}
		return flag;
	}

	public static XmlNodeList getMenuNode(string strMenuID)
	{
		string path = "VersionToKhai.XML";
		XmlDataDocument xmlDataDocument = new XmlDataDocument();
		FileStream fileStream = new FileStream(path, FileMode.Open, FileAccess.Read);
		((XmlDocument)xmlDataDocument).Load((Stream)fileStream);
		return ((XmlNode)xmlDataDocument).SelectNodes("/Root/Menu[@ID='" + strMenuID + "']");
	}

	public static void WinDoWorkConvert(string pathFile)
	{
		//IL_01ce: Unknown result type (might be due to invalid IL or missing references)
		//IL_024c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0206: Unknown result type (might be due to invalid IL or missing references)
		try
		{
			using ZipFile zipFile = ZipFile.Read(pathFile);
			HTKKFile.Delete_Folder(HTKKConfig.DATAFILES + "Old");
			HTKKFile.Create_Folder(HTKKConfig.DATAFILES + "Old");
			zipFile.ExtractAll(HTKKConfig.DATAFILES + "Old", ExtractExistingFileAction.OverwriteSilently);
			DirectoryInfo directoryInfo = new DirectoryInfo(HTKKConfig.DATAFILES + "Old");
			DirectoryInfo[] directories = directoryInfo.GetDirectories();
			bool flag = true;
			string text = "";
			string text2 = HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst + "\\log_convert.txt";
			HTKKFile.DeleteFile(text2);
			DirectoryInfo[] array = directories;
			foreach (DirectoryInfo directoryInfo2 in array)
			{
				if (directoryInfo2.Name != WinSettings.HTKKHeaderInfo.mst)
				{
					continue;
				}
				FileInfo[] files = directoryInfo2.GetFiles("*.xml");
				foreach (FileInfo fileInfo in files)
				{
					string fullName = fileInfo.FullName;
					try
					{
						string listFileError = "";
						WinFnProcess.ConverDataVb6(fullName, out listFileError);
						if (!HTKKFormater.isNullString(listFileError))
						{
							text = text + listFileError + "\r\n";
						}
					}
					catch (Exception)
					{
						flag = false;
						break;
					}
					string text3 = "";
				}
			}
			if (!HTKKFormater.isNullString(text))
			{
				HTKKFile.saveFile(HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst, "log_convert.txt", text);
			}
			HTKKFile.Delete_Folder(HTKKConfig.DATAFILES + "Old");
			if (!flag)
			{
				frmMessage frmMessage = new frmMessage("0439", "3", "1", "Thông báo");
				((Form)frmMessage).ShowDialog();
			}
			else if (File.Exists(text2))
			{
				frmMessage frmMessage = new frmMessage("0457", "2", "2", "Thông báo");
				((Form)frmMessage).ShowDialog();
				if (frmMessage.Yes)
				{
					Process.Start("notepad.exe", text2);
				}
			}
			else
			{
				frmMessage frmMessage = new frmMessage("0440", "3", "1", "Thông báo");
				((Form)frmMessage).ShowDialog();
			}
		}
		catch (Exception ex2)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex2);
		}
	}

	public static string getFileName(HTKKCMCToKhaiInfo ToKhaiInfo, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			string text = "";
			string tenCQTCapCuc = HeaderInfo.tenCQTCapCuc;
			string mst = HeaderInfo.mst;
			mst = mst.Trim();
			string soLan = PeriodInfo.SoLan;
			string kyKeKhai = GetKyKeKhai(PeriodInfo);
			kyKeKhai = kyKeKhai.Replace("/", "");
			if (mst.Length == 10)
			{
				mst += "000";
			}
			else if (mst.Length == 14)
			{
				mst = mst.Replace("-", "");
			}
			string text2 = ToKhaiInfo.MaFile;
			string text3 = ((ConvertHelpers.ConvertToInt(PeriodInfo.SoLan) <= 0) ? "L00" : ((ConvertHelpers.ConvertToInt(PeriodInfo.SoLan) >= 10) ? ("L" + soLan) : ("L0" + soLan)));
			if ((ToKhaiInfo.MenuID.Equals("G4") || ToKhaiInfo.MenuID.Equals("N5") || ToKhaiInfo.MenuID.Equals("H5") || ToKhaiInfo.MenuID.Equals("O7") || ToKhaiInfo.MenuID.Equals("P5") || ToKhaiInfo.MenuID.Equals("U6") || "U8~".Contains(ToKhaiInfo.MenuID)) && PeriodInfo.LoaiToKhai.Equals("T"))
			{
				string value = ToKhaiInfo.SheetInfos[0].StaticCells[1].Value;
				text3 = ((ConvertHelpers.ConvertToInt(value) >= 10) ? ("L" + value) : ("L0" + value));
			}
			if (ToKhaiInfo.MenuID.Equals("95"))
			{
				text2 = "02TH";
				text3 = ((PeriodInfo.SoLan.Length == 1) ? ("L00" + PeriodInfo.SoLan) : ((PeriodInfo.SoLan.Length != 2) ? ("L" + PeriodInfo.SoLan) : ("L0" + PeriodInfo.SoLan)));
				text = tenCQTCapCuc.Substring(0, 3) + "-" + mst + "-" + text2 + "-Y" + PeriodInfo.Nam + "-" + text3;
			}
			else if (ToKhaiInfo.MenuID.Equals("29"))
			{
				text = tenCQTCapCuc.Substring(0, 3) + "-" + mst + "-" + text2 + "-" + PeriodInfo.KieuKy + kyKeKhai + "-" + text3;
			}
			else if (ToKhaiInfo.MenuID.Equals("39"))
			{
				text = tenCQTCapCuc.Substring(0, 3) + "-" + mst + "-" + text2 + "-" + PeriodInfo.TuThang.Replace("/", "") + PeriodInfo.DenThang.Replace("/", "");
			}
			else if (ToKhaiInfo.MenuID.Equals("33"))
			{
				text = tenCQTCapCuc.Substring(0, 3) + "-" + mst + "-" + text2 + "-" + PeriodInfo.TuThang.Replace("/", "") + PeriodInfo.DenThang.Replace("/", "") + "-" + text3;
			}
			else if (ToKhaiInfo.MenuID.Equals("37"))
			{
				text = tenCQTCapCuc.Substring(0, 3) + "-" + mst + "-" + text2 + "-" + PeriodInfo.TuThang.Replace("/", "") + PeriodInfo.DenThang.Replace("/", "") + "-" + text3;
			}
			else if ("57,52".Contains(ToKhaiInfo.MenuID))
			{
				text = tenCQTCapCuc.Substring(0, 3) + "-" + mst + "-" + text2 + "-" + PeriodInfo.TuNgay.Replace("/", "") + PeriodInfo.DenNgay.Replace("/", "") + "-" + text3;
			}
			if (ToKhaiInfo.MenuID.Equals("66") || ToKhaiInfo.MenuID.Equals("61") || ToKhaiInfo.MenuID.Equals("67") || ToKhaiInfo.MenuID.Equals("G2"))
			{
				text = tenCQTCapCuc.Substring(0, 3) + "-" + mst + "-" + text2 + "-D" + PeriodInfo.TuNgay.Replace("/", "") + "-" + text3;
			}
			else if (ToKhaiInfo.MenuID.Equals("78"))
			{
				text = tenCQTCapCuc.Substring(0, 3) + "-" + mst + "-200_BCTC-" + PeriodInfo.TuNgay.Replace("/", "") + "_" + PeriodInfo.DenNgay.Replace("/", "") + "_" + PeriodInfo.DonViTinhMa + "-" + text3;
			}
			else if (ToKhaiInfo.MenuID.Equals("U9"))
			{
				text = tenCQTCapCuc.Substring(0, 3) + "-" + mst + "-99_BCTC-" + PeriodInfo.TuNgay.Replace("/", "") + "_" + PeriodInfo.DenNgay.Replace("/", "") + "_" + PeriodInfo.DonViTinhMa + "-" + text3;
			}
			else if (ToKhaiInfo.MenuID.Equals("A1") || ToKhaiInfo.MenuID.Equals("A3") || ToKhaiInfo.MenuID.Equals("X2") || "X1".Equals(WinSettings.TTPeriodInfo.menuID))
			{
				text = tenCQTCapCuc.Substring(0, 3) + "-" + mst + "-" + text2 + "-Q0" + PeriodInfo.Quy + PeriodInfo.Nam + "-" + text3;
			}
			else if (ToKhaiInfo.MenuID.Equals("53"))
			{
				text = tenCQTCapCuc.Substring(0, 3) + "-" + mst + "-49_BCTC-" + PeriodInfo.TuNgay.Replace("/", "") + "_" + PeriodInfo.DenNgay.Replace("/", "") + "-" + text3;
			}
			else if (!ToKhaiInfo.MenuID.Equals("10"))
			{
				text = ((ToKhaiInfo.MenuID.Equals("30") && PeriodInfo.KieuKy.Equals("Q")) ? (tenCQTCapCuc.Substring(0, 3) + "-" + mst + "-" + text2 + "-K" + PeriodInfo.TuNgay.Replace("/", "") + "_" + PeriodInfo.DenNgay.Replace("/", "") + "-" + text3) : (ToKhaiInfo.MenuID.Equals("U8") ? (tenCQTCapCuc.Substring(0, 3) + "-" + mst + "-" + text2 + "-" + PeriodInfo.KieuKy + kyKeKhai + "-F" + PeriodInfo.LanXuatBan + "-" + PeriodInfo.LoaiToKhai) : ((!"~K2~K3~J7~E6~E7~D7~R6~X5~R9~E9~".Contains(ToKhaiInfo.MenuID)) ? (tenCQTCapCuc.Substring(0, 3) + "-" + mst + "-" + text2 + "-" + PeriodInfo.KieuKy + kyKeKhai + "-" + text3) : ((!(ToKhaiInfo.MenuID == "R9")) ? (tenCQTCapCuc.Substring(0, 3) + "-" + mst + "-" + text2 + "-" + PeriodInfo.DonViTinhMa + "-" + PeriodInfo.KieuKy + kyKeKhai + "-" + text3) : (tenCQTCapCuc.Substring(0, 3) + "-" + mst + "-" + text2 + "-" + PeriodInfo.DonViTinhMa + "-LXB" + PeriodInfo.LanXuatBan.Replace("/", "").Replace("\\", "") + "-" + PeriodInfo.KieuKy + kyKeKhai + "-" + text3)))));
			}
			else
			{
				string value2 = "";
				GeneralFunctions.dicVar.TryGetValue("NgayTB03BLP", out value2);
				if (HTKKFormater.isNullString(value2))
				{
					value2 = DateTime.Now.ToString("ddMMyyyy");
				}
				text = tenCQTCapCuc.Substring(0, 3) + "-" + mst + "-03_TBAC_BLP-D" + value2.Replace("/", "") + "-" + text3;
			}
			if (text.Length > 5)
			{
				text = text.Substring(4);
			}
			return text;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return "";
		}
	}

	public static string GetKyKeKhai(HTKKPeriodInfo PeriodInfo)
	{
		try
		{
			string result = "";
			string menuID = PeriodInfo.menuID;
			if (PeriodInfo.KieuKy.Equals("D"))
			{
				result = PeriodInfo.Ngay + "/" + PeriodInfo.Thang + "/" + PeriodInfo.Nam;
			}
			else if (!menuID.Equals("01") && !menuID.Equals("02") && !menuID.Equals("04") && !menuID.Equals("71") && !menuID.Equals("36") && !menuID.Equals("25") && !menuID.Equals("68") && !menuID.Equals("14") && !menuID.Equals("96") && !menuID.Equals("94") && !menuID.Equals("18") && !menuID.Equals("38") && !menuID.Equals("32") && !menuID.Equals("34") && !menuID.Equals("28") && !menuID.Equals("31") && !menuID.Equals("62") && !menuID.Equals("D3") && !menuID.Equals("D2") && !menuID.Equals("I5") && !menuID.Equals("K9"))
			{
				result = (menuID.Equals("30") ? (PeriodInfo.KieuKy.Equals("Y") ? PeriodInfo.Nam : (("01,02,03".IndexOf(PeriodInfo.TuNgay.Split(new char[1] { '/' })[1]) > 0) ? ("1/" + PeriodInfo.TuNgay.Split(new char[1] { '/' })[2]) : (("04,05,06".IndexOf(PeriodInfo.TuNgay.Split(new char[1] { '/' })[1]) > 0) ? ("2/" + PeriodInfo.TuNgay.Split(new char[1] { '/' })[2]) : (("07,08,09".IndexOf(PeriodInfo.TuNgay.Split(new char[1] { '/' })[1]) <= 0) ? ("4/" + PeriodInfo.TuNgay.Split(new char[1] { '/' })[2]) : ("3/" + PeriodInfo.TuNgay.Split(new char[1] { '/' })[2]))))) : (PeriodInfo.KieuKy.Equals("M") ? (PeriodInfo.Thang + "/" + PeriodInfo.Nam) : (PeriodInfo.KieuKy.Equals("Q") ? (PeriodInfo.Quy + "/" + PeriodInfo.Nam) : (PeriodInfo.KieuKy.Equals("D") ? (PeriodInfo.Ngay + "/" + PeriodInfo.Thang + "/" + PeriodInfo.Nam) : ((!PeriodInfo.KieuKy.Equals("D_D")) ? PeriodInfo.Nam : (PeriodInfo.TuNgay + "_" + PeriodInfo.DenNgay))))));
			}
			else if (PeriodInfo.KieuKy.Equals("M"))
			{
				result = PeriodInfo.Thang + "/" + PeriodInfo.Nam;
			}
			else if (PeriodInfo.KieuKy.Equals("Q"))
			{
				result = PeriodInfo.Quy + "/" + PeriodInfo.Nam;
			}
			else if (PeriodInfo.KieuKy.Equals("Y"))
			{
				result = PeriodInfo.Nam;
			}
			else if (PeriodInfo.KieuKy.Equals("D"))
			{
				result = PeriodInfo.Ngay + "/" + PeriodInfo.Thang + "/" + PeriodInfo.Nam;
			}
			return result;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return "";
		}
	}

	public static string GetFileNameExcel(HTKKCMCToKhaiInfo ToKhaiInfo, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			string menuID = PeriodInfo.menuID;
			string folderName = ToKhaiInfo.FolderName;
			folderName = folderName.Trim().ToUpper();
			string text = "";
			string result;
			if ("52,57,95,39,33,37".Contains(PeriodInfo.menuID))
			{
				string text2 = HeaderInfo.mst;
				if (text2.Trim().Length == 10)
				{
					text2 = text2.Trim() + "000";
				}
				else if (text2.Trim().Length == 10)
				{
					text2 = text2.Trim().Replace("-", "");
				}
				text = folderName switch
				{
					"05_TNCN" => "05TL", 
					"05QTT_TNCN" => "05QT", 
					"05_DS_TNCN" => "05DS", 
					"02_TNCN_BH" => "02BH", 
					"02_TNCN_BHDC" => "02BH", 
					"02_TNCN_XS" => "02XS", 
					"06_TNCN10" => "06KK", 
					"09_TNCN" => "09CN", 
					"02_QTT_TNCN" => "02QT", 
					"08B_TNCN" => "08BN", 
					"16_TH_DKNPT" => "02TH", 
					_ => "", 
				};
				string text3;
				if (PeriodInfo.LoaiToKhai.Equals("C") && !WinSettings.MenuID.Equals("95"))
				{
					text3 = "L00";
				}
				else
				{
					string soLan = PeriodInfo.SoLan;
					text3 = (folderName.Equals("16_TH_DKNPT") ? ((soLan.Length == 1) ? ("L00" + soLan) : ((soLan.Length != 2) ? ("L" + soLan) : ("L0" + soLan))) : ((soLan.Length != 1) ? ("L" + soLan) : ("L0" + soLan)));
				}
				result = (folderName.Equals("16_TH_DKNPT") ? (text2 + "-" + text + "-Y" + PeriodInfo.Nam + "-" + text3) : (folderName.Equals("05_DS_TNCN") ? (text2 + "-" + text + "-" + PeriodInfo.TuThang.Replace("/", "") + PeriodInfo.DenThang.Replace("/", "")) : ((!folderName.Equals("05QTT_TNCN") && !folderName.Equals("02_QTT_TNCN")) ? (text2 + "-" + text + "-" + PeriodInfo.TuNgay.Replace("/", "") + PeriodInfo.DenNgay.Replace("/", "") + "-" + text3) : (text2 + "-" + text + "-" + PeriodInfo.TuThang.Replace("/", "") + PeriodInfo.DenThang.Replace("/", "") + "-" + text3))));
			}
			else if (menuID.Equals("68") || menuID.Equals("65"))
			{
				string text2 = HeaderInfo.mst;
				if (text2.Trim().Length == 10)
				{
					text2 = text2.Trim() + "000";
				}
				else if (text2.Trim().Length == 14)
				{
					text2 = text2.Trim().Replace("-", "");
				}
				if (folderName.Equals("BC26_AC"))
				{
					text = "BC26AC";
				}
				else if (folderName.Equals("01_AC"))
				{
					text = "BC01AC";
				}
				result = text2 + "-" + text + "-0" + PeriodInfo.Quy + PeriodInfo.Nam;
			}
			else if (menuID.Equals("14"))
			{
				result = "BC7_AC_0" + PeriodInfo.Quy + PeriodInfo.Nam;
			}
			else if (ToKhaiInfo.MenuID.Equals("U8"))
			{
				string text2 = HeaderInfo.mst;
				if (text2.Trim().Length == 10)
				{
					text2 = text2.Trim() + "000";
				}
				else if (text2.Trim().Length == 14)
				{
					text2 = text2.Trim().Replace("-", "");
				}
				string kyKeKhai = GetKyKeKhai(PeriodInfo);
				kyKeKhai = kyKeKhai.Replace("/", "");
				result = text2 + "-" + folderName + "-" + PeriodInfo.KieuKy + kyKeKhai + "-F" + PeriodInfo.LanXuatBan + "-" + PeriodInfo.LoaiToKhai;
			}
			else
			{
				result = folderName + "_" + PeriodInfo.Nam;
			}
			return result;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return "";
		}
	}

	public static string GetPrefix(HTKKCMCToKhaiInfo ToKhaiInfo, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			string text = ((!PeriodInfo.LoaiToKhai.Equals("C")) ? ("bs" + "5.7.9".Replace(".", "") + ToKhaiInfo.MenuID) : ("aa" + "5.7.9".Replace(".", "") + ToKhaiInfo.MenuID));
			string mst = HeaderInfo.mst;
			text += mst.PadRight(13, ' ');
			if (PeriodInfo.KieuKy.Equals("M"))
			{
				text = text + PeriodInfo.Thang + PeriodInfo.Nam;
			}
			else if (PeriodInfo.KieuKy.Equals("Q") || PeriodInfo.KieuKy.Equals("K"))
			{
				text = ((!(WinSettings.TTPeriodInfo.menuID == "30")) ? (text + "0" + PeriodInfo.Quy + PeriodInfo.Nam) : (text + "00" + PeriodInfo.Nam));
			}
			else if (PeriodInfo.KieuKy.Equals("Y"))
			{
				text = text + "00" + PeriodInfo.Nam;
			}
			else
			{
				string text2 = (string.IsNullOrEmpty(PeriodInfo.Thang) ? "00" : PeriodInfo.Thang);
				text = text + text2 + PeriodInfo.Nam;
			}
			text += strSession;
			text += "###";
			return text + "$$$";
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return "";
		}
	}

	public static string GetPrefixFirstBarcode(HTKKCMCToKhaiInfo ToKhaiInfo, HTKKPeriodInfo PeriodInfo, HTKKHeaderInfo HeaderInfo)
	{
		try
		{
			string text = "";
			string useDate = ToKhaiInfo.UseDate;
			return (!useDate.Equals("1")) ? (HeaderInfo.NgayBatDauNamTaiChinh + ToKhaiInfo.StartDate) : (HeaderInfo.NgayBatDauNamTaiChinh + ToKhaiInfo.StartDate + PeriodInfo.TuNgay + PeriodInfo.DenNgay);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return "";
		}
	}

	public static void SetUpDataKhbs(FpSpread winSpread, HTKKPeriodInfo PeriodInfo, bool isClickTongHopSNNC, HTKKHeaderInfo headerInfo = null)
	{
		try
		{
			if (WinSettings.ToKhaiInfo.KieuKHBS != "3")
			{
				string menuID = PeriodInfo.menuID;
				int soNgayNC = GeneralFunctions.SoNgayNopCham(PeriodInfo, headerInfo);
				DateTime hanNopToKhai = GeneralFunctions.GetHanNopToKhai(PeriodInfo, headerInfo);
				string ngayLapKHBS = PeriodInfo.NgayLapKHBS;
				int index = winSpread.Sheets.Count - 2;
				int num = 17;
				int num2 = num - 2;
				if (WinSettings.ToKhaiInfo.SheetInfos[index].Tables.Count > 0)
				{
					num2 += ((WinSettings.ToKhaiInfo.SheetInfos[index].Tables[0].RowsValues != null && WinSettings.ToKhaiInfo.SheetInfos[index].Tables[0].RowsValues.Count == 0) ? 1 : WinSettings.ToKhaiInfo.SheetInfos[index].Tables[0].RowsValues.Count);
					num2 += ((WinSettings.ToKhaiInfo.SheetInfos[index].Tables[1].RowsValues != null && WinSettings.ToKhaiInfo.SheetInfos[index].Tables[1].RowsValues.Count == 0) ? 1 : WinSettings.ToKhaiInfo.SheetInfos[index].Tables[1].RowsValues.Count);
				}
				string formulaTienPNC = GetFormulaTienPNC(0, 0.0, "BH" + num2, soNgayNC, PeriodInfo);
				if (isClickTongHopSNNC)
				{
					winSpread.Sheets[index].Cells[19, HTKKFile.ColumnLettersToIndex("BE") - 1].Formula = "BD5";
					winSpread.Sheets[index].Cells[20, HTKKFile.ColumnLettersToIndex("BE") - 1].Formula = "BF5";
				}
				winSpread.Sheets[index].Cells[4, HTKKFile.ColumnLettersToIndex("BG") - 1].Value = ngayLapKHBS;
				if (!menuID.Equals("63"))
				{
					winSpread.Sheets[index].Cells[4, HTKKFile.ColumnLettersToIndex("BD") - 1].Value = soNgayNC.ToString();
				}
				if (!menuID.Equals("02") && !menuID.Equals("63"))
				{
					if (menuID.Equals("72"))
					{
						winSpread.Sheets[index].Cells[20, HTKKFile.ColumnLettersToIndex("BE") - 1].Formula = formulaTienPNC;
					}
					else if (menuID.Equals("01") || menuID.Equals("I5"))
					{
						winSpread.Sheets[index].Cells[20, HTKKFile.ColumnLettersToIndex("BE") - 1].Formula = formulaTienPNC;
					}
					else
					{
						winSpread.Sheets[index].Cells[20, HTKKFile.ColumnLettersToIndex("BE") - 1].Formula = formulaTienPNC;
					}
				}
			}
			else
			{
				int[] array = GeneralFunctions.ArrayCellId("AA_36");
				string menuID = PeriodInfo.menuID;
				int soNgayNC = GeneralFunctions.SoNgayNopCham(PeriodInfo, headerInfo);
				int index = WinSettings.ToKhaiInfo.SheetInfos.Count - 2;
				winSpread.Sheets[index].Cells[array[1], array[0]].Value = soNgayNC;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public static string GetFormulaTienPNC(int countRowDyn, double soTien, string strColRow, int soNgayNC, HTKKPeriodInfo periodInfor)
	{
		try
		{
			string text = "";
			string text2 = "~08~44~21~40~42~98~99~93~92~89~";
			DateTime hanNopToKhai = GeneralFunctions.GetHanNopToKhai(periodInfor);
			if (!periodInfor.NgayLapKHBS.Equals(""))
			{
				DateTime dateTime = ConvertHelpers.ConvertToDateTime(periodInfor.NgayLapKHBS, "d/M/yyyy");
			}
			DateTime t = new DateTime(2013, 7, 1);
			DateTime dateTime2 = new DateTime(2013, 9, 28);
			DateTime dateTime3 = new DateTime(2014, 12, 31);
			DateTime dateTime4 = new DateTime(2016, 6, 30);
			if (DateTime.Compare(hanNopToKhai, t) < 0)
			{
				int days = (dateTime2 - hanNopToKhai).Days;
				int days2 = (dateTime3 - dateTime2).Days;
				int days3 = (dateTime4 - dateTime3).Days;
				text = ((soNgayNC > DateUtils.GetSoNgay(hanNopToKhai, dateTime4)) ? (days + "*" + strColRow + "*0.05/100 + " + days2 + "*" + strColRow + "*0.07/100 + " + days3 + "*" + strColRow + "*0.05/100 + " + (soNgayNC - days - days2 - days3) + "*" + strColRow + "*0.03/100") : ((soNgayNC <= DateUtils.GetSoNgay(hanNopToKhai, dateTime4) && soNgayNC > DateUtils.GetSoNgay(hanNopToKhai, dateTime3)) ? (days + "*" + strColRow + "*0.05/100 + " + days2 + "*" + strColRow + "*0.07/100 + " + (soNgayNC - days - days2) + "*" + strColRow + "*0.05/100") : ((soNgayNC > DateUtils.GetSoNgay(hanNopToKhai, dateTime3) || soNgayNC <= DateUtils.GetSoNgay(hanNopToKhai, dateTime2)) ? (soNgayNC + "*" + strColRow + "*0.05/100") : (days + "*" + strColRow + "*0.05/100 + " + (soNgayNC - days) + "*" + strColRow + "*0.07/100"))));
			}
			else if (DateTime.Compare(hanNopToKhai, t) >= 0 && DateTime.Compare(hanNopToKhai, dateTime3) <= 0)
			{
				int days;
				int days2;
				if (DateTime.Compare(hanNopToKhai, dateTime2) > 0)
				{
					days = 0;
					days2 = DateUtils.GetSoNgay(hanNopToKhai, dateTime3);
				}
				else
				{
					days = DateUtils.GetSoNgay(hanNopToKhai, dateTime2) + 1;
					days2 = DateUtils.GetSoNgay(dateTime2, dateTime3);
				}
				int days3 = DateUtils.GetSoNgay(dateTime3, dateTime4);
				text = ((soNgayNC > DateUtils.GetSoNgay(hanNopToKhai, dateTime4)) ? (days + "*" + strColRow + "*0.07/100 + " + days2 + "*" + strColRow + "*0.07/100 + " + days3 + "*" + strColRow + "*0.05/100 + " + (soNgayNC - days - days2 - days3) + "*" + strColRow + "*0.03/100") : ((soNgayNC <= DateUtils.GetSoNgay(hanNopToKhai, dateTime4) && soNgayNC > DateUtils.GetSoNgay(hanNopToKhai, dateTime3)) ? (days + "*" + strColRow + "*0.07/100 + " + days2 + "*" + strColRow + "*0.07/100 + " + (soNgayNC - days - days2) + "*" + strColRow + "*0.05/100") : ((soNgayNC > DateUtils.GetSoNgay(hanNopToKhai, dateTime3) || soNgayNC <= DateUtils.GetSoNgay(hanNopToKhai, dateTime2)) ? (soNgayNC + "*" + strColRow + "*0.05/100") : ((soNgayNC > 90) ? (days + "*" + strColRow + "*0.07/100 + " + (soNgayNC - days) + "*" + strColRow + "*0.07/100") : (days + "*" + strColRow + "*0.05/100 + " + (soNgayNC - days) + "*" + strColRow + "*0.05/100")))));
			}
			else if (DateTime.Compare(hanNopToKhai, dateTime3) > 0 && DateTime.Compare(hanNopToKhai, dateTime4) <= 0)
			{
				int days3 = DateUtils.GetSoNgay(hanNopToKhai, dateTime4);
				text = ((soNgayNC <= DateUtils.GetSoNgay(hanNopToKhai, dateTime4)) ? (soNgayNC + "*" + strColRow + "*0.05/100") : (days3 + "*" + strColRow + "*0.05/100 + " + (soNgayNC - days3) + "*" + strColRow + "*0.03/100"));
			}
			else
			{
				text = soNgayNC + "*" + strColRow + "*0.03/100";
			}
			if (text2.IndexOf("~" + periodInfor.menuID + "~") > 0)
			{
				return "IF(" + text + ">0,ROUND(" + text + ",2),0)";
			}
			return "IF(" + text + ">0,ROUND(" + text + ",0),0)";
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public static string GetFileNameBySheet(HTKKPeriodInfo periodInfor, string dataFile, string sFolder, string versionXML)
	{
		try
		{
			if ("N7~M1~M2~C6~C7~T5~H4~O1~O2~O3~O4~O5~O6~Q4~P7~U1~U8~X5~E9~R9~Y3~Y4~Z1~".IndexOf(WinSettings.MenuID) >= 0)
			{
				return GetFileNameBySheetNew(periodInfor, dataFile, sFolder, versionXML);
			}
			string text = "";
			string text2 = "";
			string text3 = "~98~92~08~40~42~44~M8~M9~K2~K3~J9~J8~J7~E2~E3~E4~E5~E6~E7~E9~D7~D8~D9~R9~R6~S1~";
			string text4 = "~J3~Q7~Q5~Q6~";
			versionXML = ((!(periodInfor.LoaiToKhai == "B") || !dataFile.StartsWith("KHBS")) ? versionXML.Replace(".", "") : WinSettings.ToKhaiInfo.VersionKBSXML.ToString().Replace(".", ""));
			if ("66;67;09;10;60;61;G2;".IndexOf(WinSettings.MenuID) > -1)
			{
				return sFolder + dataFile + "_" + versionXML + ".xml";
			}
			if (text3.IndexOf("~" + periodInfor.menuID + "~") > 0)
			{
				if (periodInfor.TieuMucHachToanVal.Equals("1"))
				{
					text2 = "DT";
				}
				else if (periodInfor.TieuMucHachToanVal.Equals("2"))
				{
					text2 = "KTN";
				}
				else if (periodInfor.TieuMucHachToanVal.Equals("3"))
				{
					text2 = "CD";
				}
				else if (periodInfor.TieuMucHachToanVal.Equals("4"))
				{
					text2 = "KDDT";
				}
				else if (periodInfor.TieuMucHachToanVal.Equals("5"))
				{
					text2 = "KDKTN";
				}
				else if (periodInfor.TieuMucHachToanVal.Equals("6"))
				{
					text2 = "KDCD";
				}
			}
			if (periodInfor.LoaiToKhai.Equals("C"))
			{
				if (periodInfor.KieuKy == "M")
				{
					text = ((text3.IndexOf("~" + periodInfor.menuID + "~") > 0) ? ((periodInfor.menuID == "R9") ? (sFolder + dataFile + ((!string.IsNullOrEmpty(periodInfor.DonViTinhMa)) ? ("_" + periodInfor.DonViTinhMa) : "") + "_" + text2 + "_L" + periodInfor.LanXuatBan + "_" + periodInfor.Thang + periodInfor.Nam + "_" + versionXML + ".xml") : ((!(periodInfor.menuID == "R6")) ? (sFolder + dataFile + ((!string.IsNullOrEmpty(periodInfor.DonViTinhMa)) ? ("_" + periodInfor.DonViTinhMa) : "") + "_" + text2 + "_" + periodInfor.Thang + periodInfor.Nam + "_" + versionXML + ".xml") : (sFolder + dataFile + ((!string.IsNullOrEmpty(periodInfor.DonViTinhMa)) ? ("_" + periodInfor.DonViTinhMa) : "") + "_" + text2 + "_LHD" + periodInfor.ToChucCNDNVonChuyenNhuong + "_" + periodInfor.Thang + periodInfor.Nam + "_" + versionXML + ".xml"))) : ((periodInfor.menuID == "A4") ? ((!periodInfor.MaHSDuAnDT.Equals("")) ? (sFolder + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.MaHSDuAnDT + "_" + periodInfor.TuThang.Replace("/", "") + "_" + periodInfor.DenThang.Replace("/", "") + "_" + versionXML + ".xml") : (sFolder + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.TuThang.Replace("/", "") + "_" + periodInfor.DenThang.Replace("/", "") + "_" + versionXML + ".xml")) : ((!(periodInfor.menuID == "99") && !(periodInfor.menuID == "H6")) ? ((!(periodInfor.menuID == "01") && !(periodInfor.menuID == "I5")) ? ((!(periodInfor.menuID == "D1") && !WinSettings.MenuID.Equals("Z5")) ? ((!(periodInfor.menuID == "02") && !(periodInfor.menuID == "D2")) ? ((!(periodInfor.menuID == "71") && !(periodInfor.menuID == "D4") && !WinSettings.MenuID.Equals("Z3")) ? ((!(periodInfor.menuID == "M4") && !(periodInfor.menuID == "G3")) ? ((!(periodInfor.menuID == "R5")) ? (sFolder + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.Thang + periodInfor.Nam + "_" + versionXML + ".xml") : (sFolder + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.TieuMucHachToanVal + "_" + periodInfor.Thang + periodInfor.Nam + "_" + versionXML + ".xml")) : (sFolder + dataFile + "_L" + periodInfor.LanXuatBan + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.Thang + periodInfor.Nam + "_" + versionXML + ".xml")) : ((ConvertHelpers.ConvertToInt(periodInfor.Nam) < 2021 || periodInfor.ThuHo.Equals("")) ? (sFolder + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.Thang + periodInfor.Nam + "_" + versionXML + ".xml") : (sFolder + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.ThuHo + "_" + periodInfor.Thang + periodInfor.Nam + "_" + versionXML + ".xml"))) : ((ConvertHelpers.ConvertToInt(periodInfor.Nam) < 2021) ? (sFolder + dataFile + "_" + periodInfor.Thang + periodInfor.Nam + "_" + versionXML + ".xml") : (sFolder + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.MaHSDuAnDT + "_" + periodInfor.Thang + periodInfor.Nam + "_" + versionXML + ".xml"))) : (sFolder + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.TieuMucHachToanVal + "_" + periodInfor.Thang + periodInfor.Nam + "_" + versionXML + ".xml")) : ((ConvertHelpers.ConvertToInt(periodInfor.Nam) < 2021) ? (sFolder + dataFile + "_" + periodInfor.Thang + periodInfor.Nam + "_" + versionXML + ".xml") : (sFolder + dataFile + "_" + periodInfor.TieuMucHachToanVal + "_" + periodInfor.Thang + periodInfor.Nam + "_" + versionXML + ".xml"))) : ((!periodInfor.MaHSDuAnDT.Equals("")) ? (sFolder + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.MaHSDuAnDT + ((!string.IsNullOrEmpty(GetNameTheoKieuKyHoanThue(periodInfor.KieuKyHoanThue))) ? ("_" + GetNameTheoKieuKyHoanThue(periodInfor.KieuKyHoanThue)) : "") + periodInfor.TuThang.Replace("/", "") + "_" + periodInfor.DenThang.Replace("/", "") + "_" + versionXML + ".xml") : (sFolder + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + ((!string.IsNullOrEmpty(GetNameTheoKieuKyHoanThue(periodInfor.KieuKyHoanThue))) ? ("_" + GetNameTheoKieuKyHoanThue(periodInfor.KieuKyHoanThue)) : "") + periodInfor.TuThang.Replace("/", "").PadLeft(6, '0') + "_" + periodInfor.DenThang.Replace("/", "").PadLeft(6, '0') + "_" + versionXML + ".xml")))));
				}
				else if (periodInfor.KieuKy == "Q")
				{
					text = ((!HTKKFormater.isNullString(periodInfor.TuThang) && !HTKKFormater.isNullString(periodInfor.DenThang)) ? ((!periodInfor.menuID.Equals("42") && text3.IndexOf("~" + periodInfor.menuID + "~") <= 0) ? (periodInfor.menuID.Equals("A4") ? ((!periodInfor.MaHSDuAnDT.Equals("")) ? (sFolder + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.MaHSDuAnDT + "_Q0" + periodInfor.TuThang.Replace("/", "") + "_0" + periodInfor.DenThang.Replace("/", "") + "_" + versionXML + ".xml") : (sFolder + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_Q0" + periodInfor.TuThang.Replace("/", "") + "_0" + periodInfor.DenThang.Replace("/", "") + "_" + versionXML + ".xml")) : ((!periodInfor.menuID.Equals("99") && !periodInfor.menuID.Equals("H6")) ? (sFolder + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_Q" + $"{periodInfor.Quy}{periodInfor.Nam}".PadLeft(6, '0') + "_" + periodInfor.TuThang.Replace("/", "") + "_" + periodInfor.DenThang.Replace("/", "") + "_" + versionXML + ".xml") : ((!periodInfor.MaHSDuAnDT.Equals("")) ? (sFolder + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.MaHSDuAnDT + ((!string.IsNullOrEmpty(GetNameTheoKieuKyHoanThue(periodInfor.KieuKyHoanThue))) ? ("_" + GetNameTheoKieuKyHoanThue(periodInfor.KieuKyHoanThue)) : "") + periodInfor.TuThang.Replace("/", "").PadLeft(6, '0') + "_" + periodInfor.DenThang.Replace("/", "").PadLeft(6, '0') + "_" + versionXML + ".xml") : (sFolder + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + ((!string.IsNullOrEmpty(GetNameTheoKieuKyHoanThue(periodInfor.KieuKyHoanThue))) ? ("_" + GetNameTheoKieuKyHoanThue(periodInfor.KieuKyHoanThue)) : "") + periodInfor.TuThang.Replace("/", "").PadLeft(6, '0') + "_" + periodInfor.DenThang.Replace("/", "").PadLeft(6, '0') + "_" + versionXML + ".xml")))) : (sFolder + dataFile + "_" + text2 + "_Q" + $"{periodInfor.Quy}{periodInfor.Nam}".PadLeft(6, '0') + "_" + periodInfor.TuThang.Replace("/", "") + "_" + periodInfor.DenThang.Replace("/", "") + "_" + versionXML + ".xml")) : ((text3.IndexOf("~" + periodInfor.menuID + "~") > 0) ? (sFolder + dataFile + "_" + text2 + "_Q" + $"{periodInfor.Quy}{periodInfor.Nam}".PadLeft(6, '0') + "_" + versionXML + ".xml") : ((!"A1,H1,A3,H3,A8,H8,A9,H9,B5,Q1,Q3".Contains(periodInfor.menuID) && !periodInfor.menuID.Equals("X2") && !"X1".Equals(WinSettings.TTPeriodInfo.menuID)) ? (periodInfor.menuID.Equals("30") ? (sFolder + dataFile + "_" + periodInfor.TuNgay.Replace("/", "") + "_" + periodInfor.DenNgay.Replace("/", "") + "_" + versionXML + ".xml") : ((!periodInfor.menuID.Equals("01") && !periodInfor.menuID.Equals("I5")) ? ((!periodInfor.menuID.Equals("D1") && !WinSettings.MenuID.Equals("Z5")) ? ((!periodInfor.menuID.Equals("02") && !periodInfor.menuID.Equals("D2")) ? ((!periodInfor.menuID.Equals("71") && !periodInfor.menuID.Equals("D4") && !WinSettings.MenuID.Equals("Z3")) ? ((!periodInfor.menuID.Equals("M4")) ? (sFolder + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_Q" + $"{periodInfor.Quy}{periodInfor.Nam}".PadLeft(6, '0') + "_" + versionXML + ".xml") : (sFolder + dataFile + "_L" + periodInfor.LanXuatBan + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_Q" + $"{periodInfor.Quy}{periodInfor.Nam}".PadLeft(6, '0') + "_" + versionXML + ".xml")) : ((ConvertHelpers.ConvertToInt(periodInfor.Nam) < 2021 || periodInfor.ThuHo.Equals("")) ? (sFolder + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_Q" + $"{periodInfor.Quy}{periodInfor.Nam}".PadLeft(6, '0') + "_" + versionXML + ".xml") : (sFolder + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.ThuHo + "_Q" + $"{periodInfor.Quy}{periodInfor.Nam}".PadLeft(6, '0') + "_" + versionXML + ".xml"))) : ((ConvertHelpers.ConvertToInt(periodInfor.Nam) < 2021) ? (sFolder + dataFile + "_Q" + $"{periodInfor.Quy}{periodInfor.Nam}".PadLeft(6, '0') + "_" + versionXML + ".xml") : (sFolder + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.MaHSDuAnDT + "_Q" + $"{periodInfor.Quy}{periodInfor.Nam}".PadLeft(6, '0') + "_" + versionXML + ".xml"))) : (sFolder + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.TieuMucHachToanVal + "_Q" + $"{periodInfor.Quy}{periodInfor.Nam}".PadLeft(6, '0') + "_" + versionXML + ".xml")) : ((ConvertHelpers.ConvertToInt(periodInfor.Nam) < 2021) ? (sFolder + dataFile + "_Q" + $"{periodInfor.Quy}{periodInfor.Nam}".PadLeft(6, '0') + "_" + versionXML + ".xml") : (sFolder + dataFile + "_" + periodInfor.TieuMucHachToanVal + "_Q" + $"{periodInfor.Quy}{periodInfor.Nam}".PadLeft(6, '0') + "_" + versionXML + ".xml")))) : (sFolder + dataFile + "_Q" + $"{periodInfor.Quy}{periodInfor.Nam}".PadLeft(6, '0') + "_" + periodInfor.TuNgay.Replace("/", "") + "_" + periodInfor.DenNgay.Replace("/", "") + "_" + versionXML + ".xml"))));
				}
				else if (!(periodInfor.KieuKy == "Y"))
				{
					text = ((periodInfor.KieuKy == "D") ? ((!HTKKFormater.isNullString(periodInfor.LanXuatBan) && !HTKKFormater.isNullString(periodInfor.TieuMucHachToanVal) && !HTKKFormater.isNullString(text2)) ? (sFolder + dataFile + ((!string.IsNullOrEmpty(periodInfor.DonViTinhMa)) ? ("_" + periodInfor.DonViTinhMa) : "") + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + text2 + "_L" + periodInfor.LanXuatBan + "_" + periodInfor.Ngay + periodInfor.Thang + periodInfor.Nam + "_" + versionXML + ".xml") : ((HTKKFormater.isNullString(periodInfor.LanXuatBan) && !HTKKFormater.isNullString(periodInfor.TieuMucHachToanVal) && !HTKKFormater.isNullString(text2)) ? ((!periodInfor.menuID.Equals("R6")) ? (sFolder + dataFile + ((!string.IsNullOrEmpty(periodInfor.DonViTinhMa)) ? ("_" + periodInfor.DonViTinhMa) : "") + "_" + text2 + "_" + periodInfor.Ngay + periodInfor.Thang + periodInfor.Nam + "_" + versionXML + ".xml") : (sFolder + dataFile + ((!string.IsNullOrEmpty(periodInfor.DonViTinhMa)) ? ("_" + periodInfor.DonViTinhMa) : "") + "_" + text2 + "_LHD" + periodInfor.ToChucCNDNVonChuyenNhuong + "_" + periodInfor.Ngay + periodInfor.Thang + periodInfor.Nam + "_" + versionXML + ".xml")) : (periodInfor.menuID.Equals("09") ? (sFolder + dataFile + "_" + versionXML + ".xml") : ((!periodInfor.menuID.Equals("47") && !periodInfor.menuID.Equals("K7")) ? ((!periodInfor.menuID.Equals("71") && !periodInfor.menuID.Equals("D4") && !WinSettings.MenuID.Equals("Z3")) ? (periodInfor.menuID.Equals("M5") ? (sFolder + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.Ngay + periodInfor.Thang + periodInfor.Nam + "_" + versionXML + ".xml") : ((!periodInfor.menuID.Equals("T6") && !periodInfor.menuID.Equals("U2")) ? (periodInfor.menuID.Equals("Y5") ? (sFolder + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_CN" + periodInfor.TieuMucHachToanVal + "_" + periodInfor.TuNgay.Replace("/", "") + "_" + periodInfor.Ngay + periodInfor.Thang + periodInfor.Nam + "_" + versionXML + ".xml") : (periodInfor.menuID.Equals("G3") ? (sFolder + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_L" + periodInfor.LanXuatBan + "_" + periodInfor.Ngay + periodInfor.Thang + periodInfor.Nam + "_" + versionXML + ".xml") : ((!periodInfor.menuID.Equals("P6") && !periodInfor.menuID.Equals("P1") && !periodInfor.menuID.Equals("X7") && !periodInfor.menuID.Equals("Y9")) ? ((!periodInfor.menuID.Equals("R5")) ? (sFolder + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.Ngay + periodInfor.Thang + periodInfor.Nam + "_" + versionXML + ".xml") : (sFolder + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.TieuMucHachToanVal + "_" + periodInfor.Ngay + periodInfor.Thang + periodInfor.Nam + "_" + versionXML + ".xml")) : (sFolder + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_L" + periodInfor.LanXuatBan + "_" + periodInfor.MaHSDuAnDT + "_" + periodInfor.Ngay + periodInfor.Thang + periodInfor.Nam + "_" + versionXML + ".xml")))) : (sFolder + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_CN" + periodInfor.TieuMucHachToanVal + "_" + periodInfor.Ngay + periodInfor.Thang + periodInfor.Nam + "_" + versionXML + ".xml"))) : ((ConvertHelpers.ConvertToInt(periodInfor.Nam) < 2021 || periodInfor.ThuHo.Equals("")) ? (sFolder + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.Ngay + periodInfor.Thang + periodInfor.Nam + "_" + versionXML + ".xml") : (sFolder + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.ThuHo + "_" + periodInfor.Ngay + periodInfor.Thang + periodInfor.Nam + "_" + versionXML + ".xml"))) : (sFolder + dataFile + "_" + periodInfor.Ngay + periodInfor.Thang + periodInfor.Nam + "_" + (periodInfor.CTDC ? "CTDC" : "LNDC") + "_" + versionXML + ".xml"))))) : ((periodInfor.KieuKy == "D_D") ? (sFolder + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.TuNgay.Replace("/", "") + "_" + periodInfor.DenNgay.Replace("/", "") + "_" + versionXML + ".xml") : ((!(periodInfor.KieuKy == "K")) ? (sFolder + dataFile) : (sFolder + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_K" + $"{periodInfor.Quy}{periodInfor.Nam}".PadLeft(6, '0') + "_" + versionXML + ".xml"))));
				}
				else if (!HTKKFormater.isNullString(periodInfor.TuThang) && !HTKKFormater.isNullString(periodInfor.DenThang) && !periodInfor.menuID.Equals("42") && !periodInfor.menuID.Equals("44") && text3.IndexOf("~" + periodInfor.menuID + "~") <= 0)
				{
					text = sFolder + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.Nam + "_" + periodInfor.TuThang.Replace("/", "") + "_" + periodInfor.DenThang.Replace("/", "") + "_" + versionXML + ".xml";
				}
				else if (periodInfor.menuID.Equals("93") || periodInfor.menuID.Equals("89") || periodInfor.menuID.Equals("44") || periodInfor.menuID.Equals("42") || text3.IndexOf("~" + periodInfor.menuID + "~") > 0)
				{
					text = sFolder + dataFile + ((!string.IsNullOrEmpty(periodInfor.DonViTinhMa)) ? ("_" + periodInfor.DonViTinhMa) : "") + "_" + text2 + "_" + periodInfor.Nam + "_" + periodInfor.TuThang.Replace("/", "") + "_" + periodInfor.DenThang.Replace("/", "") + "_" + versionXML + ".xml";
				}
				else if (periodInfor.menuID.Equals("95"))
				{
					text = sFolder + dataFile + "_L" + periodInfor.SoLan + "_" + periodInfor.Nam + "_" + versionXML + ".xml";
				}
				else if ("78,58".Contains(periodInfor.menuID) || "U9".Equals(periodInfor.menuID))
				{
					text = sFolder + dataFile + "_" + periodInfor.Nam + "_" + periodInfor.TuNgay.Replace("/", "") + "_" + periodInfor.DenNgay.Replace("/", "") + "_" + periodInfor.DonViTinhMa + "_" + versionXML + ".xml";
				}
				else if ("A2,A5,A6,A7,B3,B4".Contains(periodInfor.menuID))
				{
					text = sFolder + dataFile + "_" + periodInfor.Nam + "_" + periodInfor.TuNgay.Replace("/", "") + "_" + periodInfor.DenNgay.Replace("/", "") + "_" + versionXML + ".xml";
				}
				else if ("C1,C2,C3,C4,C5,P8,P9,Q2,U3,U4,X3,X4".Contains(periodInfor.menuID))
				{
					text = sFolder + dataFile + "_" + periodInfor.Nam + "_" + periodInfor.TuNgay.Replace("/", "") + "_" + periodInfor.DenNgay.Replace("/", "") + "_" + versionXML + ".xml";
				}
				else if (periodInfor.menuID.Equals("K8") || periodInfor.menuID.Equals("03") || periodInfor.menuID.Equals("83") || periodInfor.menuID.Equals("N8") || text4.IndexOf("~" + periodInfor.menuID + "~") >= 0)
				{
					text = sFolder + dataFile + "_" + periodInfor.Nam + "_" + periodInfor.TuNgay.Replace("/", "") + "_" + periodInfor.DenNgay.Replace("/", "") + "_" + versionXML + ".xml";
				}
				else if (periodInfor.menuID.Equals("E8"))
				{
					text = sFolder + dataFile + "_" + periodInfor.TieuMucHachToanVal + "_" + periodInfor.Nam + "_" + periodInfor.TuNgay.Replace("/", "") + "_" + periodInfor.DenNgay.Replace("/", "") + "_" + versionXML + ".xml";
				}
				else if ("53,54".Contains(periodInfor.menuID))
				{
					text = sFolder + dataFile + "_" + periodInfor.Nam + "_" + periodInfor.TuNgay.Replace("/", "") + "_" + periodInfor.DenNgay.Replace("/", "") + "_" + versionXML + ".xml";
				}
				else if (periodInfor.menuID.Equals("63"))
				{
					text = sFolder + dataFile + "_" + periodInfor.Nam;
					bool flag = false;
					string text5 = "";
					if (periodInfor.dicVar.Count > 0)
					{
						flag = Convert.ToBoolean(periodInfor.dicVar["CoSoMoiThanhLap"]);
						text5 = (string)periodInfor.dicVar["NgayThanhLap"];
						if (flag)
						{
							text = text + "_CSM_" + text5.Replace("/", "");
						}
					}
					text = text + "_" + versionXML + ".xml";
				}
				else if (!periodInfor.menuID.Equals("L1"))
				{
					text = ((!periodInfor.menuID.Equals("I2") && !periodInfor.menuID.Equals("L9") && !periodInfor.menuID.Equals("I1")) ? (sFolder + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.Nam + "_" + versionXML + ".xml") : (sFolder + dataFile + "_L" + periodInfor.LanXuatBan + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.Nam + "_" + versionXML + ".xml"));
				}
				else
				{
					text = sFolder + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.Nam;
					bool flag = false;
					string text5 = "";
					if (periodInfor.dicVar.Count > 0)
					{
						flag = Convert.ToBoolean(periodInfor.dicVar["CoSoMoiThanhLap"]);
						text5 = (string)periodInfor.dicVar["NgayThanhLap"];
						if (flag)
						{
							text = text + "_CSM_" + text5.Replace("/", "");
						}
					}
					text = text + "_" + versionXML + ".xml";
				}
			}
			else if (periodInfor.KieuKy == "M")
			{
				text = ((text3.IndexOf("~" + periodInfor.menuID + "~") > 0) ? ((periodInfor.menuID == "R9") ? (sFolder + "bs" + periodInfor.SoLan + "_" + dataFile + ((!string.IsNullOrEmpty(periodInfor.DonViTinhMa)) ? ("_" + periodInfor.DonViTinhMa) : "") + "_" + text2 + "_L" + periodInfor.LanXuatBan + "_" + periodInfor.Thang + periodInfor.Nam + "_" + versionXML + ".xml") : ((!(periodInfor.menuID == "R6")) ? (sFolder + "bs" + periodInfor.SoLan + "_" + dataFile + ((!string.IsNullOrEmpty(periodInfor.DonViTinhMa)) ? ("_" + periodInfor.DonViTinhMa) : "") + "_" + text2 + "_" + periodInfor.Thang + periodInfor.Nam + "_" + versionXML + ".xml") : (sFolder + "bs" + periodInfor.SoLan + "_" + dataFile + ((!string.IsNullOrEmpty(periodInfor.DonViTinhMa)) ? ("_" + periodInfor.DonViTinhMa) : "") + "_" + text2 + "_LHD" + periodInfor.ToChucCNDNVonChuyenNhuong + "_" + periodInfor.Thang + periodInfor.Nam + "_" + versionXML + ".xml"))) : ((!periodInfor.menuID.Equals("01") && !periodInfor.menuID.Equals("I5")) ? ((!periodInfor.menuID.Equals("D1") && !WinSettings.MenuID.Equals("Z5")) ? ((!periodInfor.menuID.Equals("02") && !periodInfor.menuID.Equals("D2")) ? ((!periodInfor.menuID.Equals("71") && !periodInfor.menuID.Equals("D4") && !WinSettings.MenuID.Equals("Z3")) ? ((!periodInfor.menuID.Equals("R5")) ? (sFolder + "bs" + periodInfor.SoLan + "_" + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.Thang + periodInfor.Nam + "_" + versionXML + ".xml") : (sFolder + "bs" + periodInfor.SoLan + "_" + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.TieuMucHachToanVal + "_" + periodInfor.Thang + periodInfor.Nam + "_" + versionXML + ".xml")) : ((ConvertHelpers.ConvertToInt(periodInfor.Nam) < 2021 || periodInfor.ThuHo.Equals("")) ? (sFolder + "bs" + periodInfor.SoLan + "_" + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.Thang + periodInfor.Nam + "_" + versionXML + ".xml") : (sFolder + "bs" + periodInfor.SoLan + "_" + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.ThuHo + "_" + periodInfor.Thang + periodInfor.Nam + "_" + versionXML + ".xml"))) : ((ConvertHelpers.ConvertToInt(periodInfor.Nam) < 2021) ? (sFolder + "bs" + periodInfor.SoLan + "_" + dataFile + "_" + periodInfor.Thang + periodInfor.Nam + "_" + versionXML + ".xml") : (sFolder + "bs" + periodInfor.SoLan + "_" + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.MaHSDuAnDT + "_" + periodInfor.Thang + periodInfor.Nam + "_" + versionXML + ".xml"))) : (sFolder + "bs" + periodInfor.SoLan + "_" + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.TieuMucHachToanVal + "_" + periodInfor.Thang + periodInfor.Nam + "_" + versionXML + ".xml")) : ((ConvertHelpers.ConvertToInt(periodInfor.Nam) < 2021) ? (sFolder + "bs" + periodInfor.SoLan + "_" + dataFile + "_" + periodInfor.Thang + periodInfor.Nam + "_" + versionXML + ".xml") : (sFolder + "bs" + periodInfor.SoLan + "_" + dataFile + "_" + periodInfor.TieuMucHachToanVal + "_" + periodInfor.Thang + periodInfor.Nam + "_" + versionXML + ".xml"))));
			}
			else if (periodInfor.KieuKy == "Q")
			{
				text = ((!HTKKFormater.isNullString(periodInfor.TuThang) && !HTKKFormater.isNullString(periodInfor.DenThang)) ? ((!periodInfor.menuID.Equals("42") && text3.IndexOf("~" + periodInfor.menuID + "~") <= 0) ? (sFolder + "bs" + periodInfor.SoLan + "_" + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_Q0" + periodInfor.Quy + periodInfor.Nam + "_" + periodInfor.TuThang.Replace("/", "") + "_" + periodInfor.DenThang.Replace("/", "") + "_" + versionXML + ".xml") : (sFolder + "bs" + periodInfor.SoLan + "_" + dataFile + "_" + text2 + "_Q0" + periodInfor.Quy + periodInfor.Nam + "_" + periodInfor.TuThang.Replace("/", "") + "_" + periodInfor.DenThang.Replace("/", "") + "_" + versionXML + ".xml")) : ((text3.IndexOf("~" + periodInfor.menuID + "~") > 0) ? (sFolder + "bs" + periodInfor.SoLan + "_" + dataFile + "_" + text2 + "_Q0" + periodInfor.Quy + periodInfor.Nam + "_" + versionXML + ".xml") : ((!periodInfor.menuID.Equals("A1") && !periodInfor.menuID.Equals("A3") && !periodInfor.menuID.Equals("X2") && !"X1".Equals(WinSettings.TTPeriodInfo.menuID)) ? (periodInfor.menuID.Equals("30") ? (sFolder + "bs" + periodInfor.SoLan + "_" + dataFile + "_" + periodInfor.TuNgay.Replace("/", "") + "_" + periodInfor.DenNgay.Replace("/", "") + "_" + versionXML + ".xml") : ((!periodInfor.menuID.Equals("01") && !periodInfor.menuID.Equals("I5")) ? ((!periodInfor.menuID.Equals("D1") && !WinSettings.MenuID.Equals("Z5")) ? ((!periodInfor.menuID.Equals("02") && !periodInfor.menuID.Equals("D2")) ? ((!periodInfor.menuID.Equals("71") && !periodInfor.menuID.Equals("D4") && !WinSettings.MenuID.Equals("Z3")) ? (sFolder + "bs" + periodInfor.SoLan + "_" + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_Q0" + periodInfor.Quy + periodInfor.Nam + "_" + versionXML + ".xml") : ((ConvertHelpers.ConvertToInt(periodInfor.Nam) < 2021 || periodInfor.ThuHo.Equals("")) ? (sFolder + "bs" + periodInfor.SoLan + "_" + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_Q0" + periodInfor.Quy + periodInfor.Nam + "_" + versionXML + ".xml") : (sFolder + "bs" + periodInfor.SoLan + "_" + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.ThuHo + "_Q0" + periodInfor.Quy + periodInfor.Nam + "_" + versionXML + ".xml"))) : ((ConvertHelpers.ConvertToInt(periodInfor.Nam) < 2021) ? (sFolder + "bs" + periodInfor.SoLan + "_" + dataFile + "_Q0" + periodInfor.Quy + periodInfor.Nam + "_" + versionXML + ".xml") : (sFolder + "bs" + periodInfor.SoLan + "_" + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.MaHSDuAnDT + "_Q0" + periodInfor.Quy + periodInfor.Nam + "_" + versionXML + ".xml"))) : (sFolder + "bs" + periodInfor.SoLan + "_" + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.TieuMucHachToanVal + "_Q0" + periodInfor.Quy + periodInfor.Nam + "_" + versionXML + ".xml")) : ((ConvertHelpers.ConvertToInt(periodInfor.Nam) < 2021) ? (sFolder + "bs" + periodInfor.SoLan + "_" + dataFile + "_Q0" + periodInfor.Quy + periodInfor.Nam + "_" + versionXML + ".xml") : (sFolder + "bs" + periodInfor.SoLan + "_" + dataFile + "_" + periodInfor.TieuMucHachToanVal + "_Q0" + periodInfor.Quy + periodInfor.Nam + "_" + versionXML + ".xml")))) : (sFolder + "bs" + periodInfor.SoLan + "_" + dataFile + "_Q0" + periodInfor.Quy + periodInfor.Nam + versionXML + ".xml"))));
			}
			else if (!(periodInfor.KieuKy == "Y"))
			{
				text = ((periodInfor.KieuKy == "D") ? ((!periodInfor.menuID.Equals("G4") && !periodInfor.menuID.Equals("N5") && !periodInfor.menuID.Equals("H5") && !periodInfor.menuID.Equals("O7") && !periodInfor.menuID.Equals("P5") && !periodInfor.menuID.Equals("U6") && !"U8~".Contains(periodInfor.menuID)) ? ((!HTKKFormater.isNullString(periodInfor.LanXuatBan) && !HTKKFormater.isNullString(periodInfor.TieuMucHachToanVal) && !HTKKFormater.isNullString(text2)) ? (sFolder + "bs" + periodInfor.SoLan + "_" + dataFile + ((!string.IsNullOrEmpty(periodInfor.DonViTinhMa)) ? ("_" + periodInfor.DonViTinhMa) : "") + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + text2 + "_L" + periodInfor.LanXuatBan + "_" + periodInfor.Ngay + periodInfor.Thang + periodInfor.Nam + "_" + versionXML + ".xml") : ((HTKKFormater.isNullString(periodInfor.LanXuatBan) && !HTKKFormater.isNullString(periodInfor.TieuMucHachToanVal) && !HTKKFormater.isNullString(text2)) ? ((!periodInfor.menuID.Equals("R6")) ? (sFolder + "bs" + periodInfor.SoLan + "_" + dataFile + ((!string.IsNullOrEmpty(periodInfor.DonViTinhMa)) ? ("_" + periodInfor.DonViTinhMa) : "") + "_" + text2 + "_" + periodInfor.Ngay + periodInfor.Thang + periodInfor.Nam + "_" + versionXML + ".xml") : (sFolder + "bs" + periodInfor.SoLan + "_" + dataFile + ((!string.IsNullOrEmpty(periodInfor.DonViTinhMa)) ? ("_" + periodInfor.DonViTinhMa) : "") + "_" + text2 + "_LHD" + periodInfor.ToChucCNDNVonChuyenNhuong + "_" + periodInfor.Ngay + periodInfor.Thang + periodInfor.Nam + "_" + versionXML + ".xml")) : ((!periodInfor.menuID.Equals("47") && !periodInfor.menuID.Equals("K7")) ? ((!periodInfor.menuID.Equals("T6") && !periodInfor.menuID.Equals("U2")) ? (periodInfor.menuID.Equals("Y5") ? (sFolder + "bs" + periodInfor.SoLan + "_" + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_CN" + periodInfor.TieuMucHachToanVal + "_" + periodInfor.TuNgay.Replace("/", "") + "_" + periodInfor.Ngay + periodInfor.Thang + periodInfor.Nam + "_" + versionXML + ".xml") : ((!periodInfor.menuID.Equals("71") && !periodInfor.menuID.Equals("D4") && !WinSettings.MenuID.Equals("Z3")) ? ((!periodInfor.menuID.Equals("R5")) ? (sFolder + "bs" + periodInfor.SoLan + "_" + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.Ngay + periodInfor.Thang + periodInfor.Nam + "_" + versionXML + ".xml") : (sFolder + "bs" + periodInfor.SoLan + "_" + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.TieuMucHachToanVal + "_" + periodInfor.Ngay + periodInfor.Thang + periodInfor.Nam + "_" + versionXML + ".xml")) : ((ConvertHelpers.ConvertToInt(periodInfor.Nam) < 2021 || periodInfor.ThuHo.Equals("")) ? (sFolder + "bs" + periodInfor.SoLan + "_" + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.Ngay + periodInfor.Thang + periodInfor.Nam + "_" + versionXML + ".xml") : (sFolder + "bs" + periodInfor.SoLan + "_" + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.ThuHo + "_" + periodInfor.Ngay + periodInfor.Thang + periodInfor.Nam + "_" + versionXML + ".xml")))) : (sFolder + "bs" + periodInfor.SoLan + "_" + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_CN" + periodInfor.TieuMucHachToanVal + "_" + periodInfor.Ngay + periodInfor.Thang + periodInfor.Nam + "_" + versionXML + ".xml")) : (sFolder + "bs" + periodInfor.SoLan + "_" + dataFile + "_" + periodInfor.Ngay + periodInfor.Thang + periodInfor.Nam + "_" + (periodInfor.CTDC ? "CTDC" : "LNDC") + "_" + versionXML + ".xml")))) : (sFolder + "bs_" + dataFile + "_" + periodInfor.Ngay + periodInfor.Thang + periodInfor.Nam + "_" + versionXML + ".xml")) : ((periodInfor.KieuKy == "D_D") ? (sFolder + "bs" + periodInfor.SoLan + "_" + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.TuNgay.Replace("/", "") + "_" + periodInfor.DenNgay.Replace("/", "") + "_" + versionXML + ".xml") : ((!(periodInfor.KieuKy == "K")) ? (sFolder + "bs" + periodInfor.SoLan + "_" + dataFile) : (sFolder + "bs" + periodInfor.SoLan + "_" + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_K" + $"{periodInfor.Quy}{periodInfor.Nam}".PadLeft(6, '0') + "_" + versionXML + ".xml"))));
			}
			else if (!HTKKFormater.isNullString(periodInfor.TuThang) && !HTKKFormater.isNullString(periodInfor.DenThang) && !periodInfor.menuID.Equals("42") && !periodInfor.menuID.Equals("44") && text3.IndexOf("~" + periodInfor.menuID + "~") <= 0)
			{
				text = sFolder + "bs" + periodInfor.SoLan + "_" + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.Nam + "_" + periodInfor.TuThang.Replace("/", "") + "_" + periodInfor.DenThang.Replace("/", "") + "_" + versionXML + ".xml";
			}
			else if (periodInfor.menuID.Equals("93") || periodInfor.menuID.Equals("89") || periodInfor.menuID.Equals("44") || periodInfor.menuID.Equals("42") || text3.IndexOf("~" + periodInfor.menuID + "~") > 0)
			{
				text = sFolder + "bs" + periodInfor.SoLan + "_" + dataFile + ((!string.IsNullOrEmpty(periodInfor.DonViTinhMa)) ? ("_" + periodInfor.DonViTinhMa) : "") + "_" + text2 + "_" + periodInfor.Nam + "_" + periodInfor.TuThang.Replace("/", "") + "_" + periodInfor.DenThang.Replace("/", "") + "_" + versionXML + ".xml";
			}
			else if (periodInfor.menuID.Equals("95"))
			{
				text = sFolder + "bs" + periodInfor.SoLan + "_" + dataFile + "_L" + periodInfor.SoLan + "_" + periodInfor.Nam + "_" + versionXML + ".xml";
			}
			else if (periodInfor.menuID.Equals("78") || periodInfor.menuID.Equals("U9"))
			{
				text = sFolder + "bs" + periodInfor.SoLan + "_" + dataFile + "_" + periodInfor.Nam + "_" + periodInfor.TuNgay.Replace("/", "") + "_" + periodInfor.DenNgay.Replace("/", "") + "_" + periodInfor.DonViTinhMa + "_" + versionXML + ".xml";
			}
			else if (periodInfor.menuID.Equals("K8") || periodInfor.menuID.Equals("03") || periodInfor.menuID.Equals("83") || periodInfor.menuID.Equals("N8") || text4.IndexOf("~" + periodInfor.menuID + "~") >= 0)
			{
				text = sFolder + "bs" + periodInfor.SoLan + "_" + dataFile + "_" + periodInfor.Nam + "_" + periodInfor.TuNgay.Replace("/", "") + "_" + periodInfor.DenNgay.Replace("/", "") + "_" + versionXML + ".xml";
			}
			else if (periodInfor.menuID.Equals("E8"))
			{
				text = sFolder + "bs" + periodInfor.SoLan + "_" + dataFile + "_" + periodInfor.TieuMucHachToanVal + "_" + periodInfor.Nam + "_" + periodInfor.TuNgay.Replace("/", "") + "_" + periodInfor.DenNgay.Replace("/", "") + "_" + versionXML + ".xml";
			}
			else if (periodInfor.menuID.Equals("53"))
			{
				text = sFolder + "bs" + periodInfor.SoLan + "_" + dataFile + "_" + periodInfor.Nam + "_" + periodInfor.TuNgay.Replace("/", "") + "_" + periodInfor.DenNgay.Replace("/", "") + "_" + versionXML + ".xml";
			}
			else if (periodInfor.menuID.Equals("63"))
			{
				text = sFolder + "bs" + periodInfor.SoLan + "_" + dataFile + "_" + periodInfor.Nam;
				bool flag = false;
				string text5 = "";
				if (periodInfor.dicVar.Count > 0)
				{
					flag = Convert.ToBoolean(periodInfor.dicVar["CoSoMoiThanhLap"]);
					text5 = (string)periodInfor.dicVar["NgayThanhLap"];
					if (flag)
					{
						text = text + "_CSM_" + text5.Replace("/", "");
					}
				}
				text = text + "_" + versionXML + ".xml";
			}
			else if (!periodInfor.menuID.Equals("L1"))
			{
				text = ((!periodInfor.menuID.Equals("I1") && !periodInfor.menuID.Equals("L9") && !periodInfor.menuID.Equals("I2")) ? (sFolder + "bs" + periodInfor.SoLan + "_" + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.Nam + "_" + versionXML + ".xml") : (sFolder + "bs" + periodInfor.SoLan + "_" + dataFile + "_L" + periodInfor.LanXuatBan + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.Nam + "_" + versionXML + ".xml"));
			}
			else
			{
				text = sFolder + "bs" + periodInfor.SoLan + "_" + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.Nam;
				bool flag = false;
				string text5 = "";
				if (periodInfor.dicVar.Count > 0)
				{
					flag = Convert.ToBoolean(periodInfor.dicVar["CoSoMoiThanhLap"]);
					text5 = (string)periodInfor.dicVar["NgayThanhLap"];
					if (flag)
					{
						text = text + "_CSM_" + text5.Replace("/", "");
					}
				}
				text = text + "_" + versionXML + ".xml";
			}
			return text;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public static string GetNameNopTheoTinhCQT()
	{
		if (!string.IsNullOrEmpty(WinSettings.ToKhaiInfo.NopTheoTinhCQT))
		{
			if (WinSettings.ToKhaiInfo.NopTheoTinhCQT == "1")
			{
				if (!string.IsNullOrEmpty(WinSettings.TTPeriodInfo.MaTinh))
				{
					return "TINH" + WinSettings.TTPeriodInfo.MaTinh;
				}
				if (!string.IsNullOrEmpty(WinSettings.TTPeriodInfo.maCQTQuanLy))
				{
					return "CQTQL" + WinSettings.TTPeriodInfo.maCQTQuanLy;
				}
			}
			if (WinSettings.ToKhaiInfo.NopTheoTinhCQT == "2" && !string.IsNullOrEmpty(WinSettings.TTPeriodInfo.maCQTQuanLy))
			{
				return "CQTQL" + WinSettings.TTPeriodInfo.maCQTQuanLy;
			}
			if (WinSettings.ToKhaiInfo.NopTheoTinhCQT == "3" && !string.IsNullOrEmpty(WinSettings.TTPeriodInfo.maCQTQuanLy))
			{
				return "CQTQL" + WinSettings.TTPeriodInfo.maCQTQuanLy;
			}
		}
		return "";
	}

	public static string GetNameTheoKieuKyHoanThue(string kieuKyHT)
	{
		string result = "";
		switch (kieuKyHT)
		{
		case "1":
			result = "TT";
			break;
		case "3":
			result = "TQ";
			break;
		case "4":
			result = "QT";
			break;
		case "2":
			result = "QQ";
			break;
		}
		return result;
	}

	public static string GetFileNameToKhaiCt(HTKKPeriodInfo periodInfor, string dataFile, string sFolder, string versionXML)
	{
		try
		{
			if ("N7~M1~M2~C6~C7~T5~H4~O1~O2~O3~O4~O5~O6~Q4~P7~U1~U8~X5~E9~R9~Y3~Y4~Z1~".IndexOf(WinSettings.MenuID) >= 0)
			{
				return GetFileNameBySheetNew(periodInfor, dataFile, sFolder, versionXML, isLoadSavedFile: false);
			}
			string text = "";
			string text2 = "";
			string text3 = "~98~92~08~40~42~44~M8~M9~K2~K3~J9~J8~J7~E2~E3~E4~E5~E6~E7~E9~D7~D8~D9~R9~R6~S1~";
			versionXML = versionXML.Replace(".", "");
			if (text3.IndexOf("~" + periodInfor.menuID + "~") > 0)
			{
				if (periodInfor.TieuMucHachToanVal.Equals("1"))
				{
					text2 = "DT";
				}
				else if (periodInfor.TieuMucHachToanVal.Equals("2"))
				{
					text2 = "KTN";
				}
				else if (periodInfor.TieuMucHachToanVal.Equals("3"))
				{
					text2 = "CD";
				}
				else if (periodInfor.TieuMucHachToanVal.Equals("4"))
				{
					text2 = "KDDT";
				}
				else if (periodInfor.TieuMucHachToanVal.Equals("5"))
				{
					text2 = "KDKTN";
				}
				else if (periodInfor.TieuMucHachToanVal.Equals("6"))
				{
					text2 = "KDCD";
				}
			}
			if (periodInfor.KieuKy == "M")
			{
				text = ((text3.IndexOf("~" + periodInfor.menuID + "~") > 0) ? ((periodInfor.menuID == "R9") ? (sFolder + dataFile + ((!string.IsNullOrEmpty(periodInfor.DonViTinhMa)) ? ("_" + periodInfor.DonViTinhMa) : "") + "_" + text2 + "_L" + periodInfor.LanXuatBan + "_" + periodInfor.Thang + periodInfor.Nam + "_" + versionXML + ".xml") : ((!(periodInfor.menuID == "R6")) ? (sFolder + dataFile + ((!string.IsNullOrEmpty(periodInfor.DonViTinhMa)) ? ("_" + periodInfor.DonViTinhMa) : "") + "_" + text2 + "_" + periodInfor.Thang + periodInfor.Nam + "_" + versionXML + ".xml") : (sFolder + dataFile + ((!string.IsNullOrEmpty(periodInfor.DonViTinhMa)) ? ("_" + periodInfor.DonViTinhMa) : "") + "_" + text2 + "_LHD" + periodInfor.ToChucCNDNVonChuyenNhuong + "_" + periodInfor.Thang + periodInfor.Nam + "_" + versionXML + ".xml"))) : ((!periodInfor.menuID.Equals("01") && !periodInfor.menuID.Equals("I5")) ? ((!periodInfor.menuID.Equals("D1") && !WinSettings.MenuID.Equals("Z5")) ? ((!periodInfor.menuID.Equals("02") && !periodInfor.menuID.Equals("D2")) ? ((!periodInfor.menuID.Equals("71") && !periodInfor.menuID.Equals("D4") && !WinSettings.MenuID.Equals("Z3")) ? ((!periodInfor.menuID.Equals("R5")) ? (sFolder + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.Thang + periodInfor.Nam + "_" + versionXML + ".xml") : (sFolder + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.TieuMucHachToanVal + "_" + periodInfor.Thang + periodInfor.Nam + "_" + versionXML + ".xml")) : ((ConvertHelpers.ConvertToInt(periodInfor.Nam) < 2021 || periodInfor.ThuHo.Equals("")) ? (sFolder + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.Thang + periodInfor.Nam + "_" + versionXML + ".xml") : (sFolder + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.ThuHo + "_" + periodInfor.Thang + periodInfor.Nam + "_" + versionXML + ".xml"))) : ((ConvertHelpers.ConvertToInt(periodInfor.Nam) < 2021) ? (sFolder + dataFile + "_" + periodInfor.Thang + periodInfor.Nam + "_" + versionXML + ".xml") : (sFolder + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.MaHSDuAnDT + "_" + periodInfor.Thang + periodInfor.Nam + "_" + versionXML + ".xml"))) : (sFolder + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.TieuMucHachToanVal + "_" + periodInfor.Thang + periodInfor.Nam + "_" + versionXML + ".xml")) : ((ConvertHelpers.ConvertToInt(periodInfor.Nam) < 2021) ? (sFolder + dataFile + "_" + periodInfor.Thang + periodInfor.Nam + "_" + versionXML + ".xml") : (sFolder + dataFile + "_" + periodInfor.TieuMucHachToanVal + "_" + periodInfor.Thang + periodInfor.Nam + "_" + versionXML + ".xml"))));
			}
			else if (periodInfor.KieuKy == "Q")
			{
				text = ((!HTKKFormater.isNullString(periodInfor.TuThang) && !HTKKFormater.isNullString(periodInfor.DenThang)) ? ((!periodInfor.menuID.Equals("42") && text3.IndexOf("~" + periodInfor.menuID + "~") <= 0) ? (periodInfor.menuID.Equals("A4") ? ((!periodInfor.MaHSDuAnDT.Equals("")) ? (sFolder + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.MaHSDuAnDT + "_Q0" + periodInfor.TuThang.Replace("/", "") + "_0" + periodInfor.DenThang.Replace("/", "") + "_" + versionXML + ".xml") : (sFolder + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_Q0" + periodInfor.TuThang.Replace("/", "") + "_0" + periodInfor.DenThang.Replace("/", "") + "_" + versionXML + ".xml")) : ((!periodInfor.menuID.Equals("99") && !periodInfor.menuID.Equals("H6")) ? (sFolder + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_Q" + $"{periodInfor.Quy}{periodInfor.Nam}".PadLeft(6, '0') + "_" + periodInfor.TuThang.Replace("/", "") + "_" + periodInfor.DenThang.Replace("/", "") + "_" + versionXML + ".xml") : ((!periodInfor.MaHSDuAnDT.Equals("")) ? (sFolder + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.MaHSDuAnDT + ((!string.IsNullOrEmpty(GetNameTheoKieuKyHoanThue(periodInfor.KieuKyHoanThue))) ? ("_" + GetNameTheoKieuKyHoanThue(periodInfor.KieuKyHoanThue)) : "") + periodInfor.TuThang.Replace("/", "").PadLeft(6, '0') + "_" + periodInfor.DenThang.Replace("/", "").PadLeft(6, '0') + "_" + versionXML + ".xml") : (sFolder + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + ((!string.IsNullOrEmpty(GetNameTheoKieuKyHoanThue(periodInfor.KieuKyHoanThue))) ? ("_" + GetNameTheoKieuKyHoanThue(periodInfor.KieuKyHoanThue)) : "") + periodInfor.TuThang.Replace("/", "").PadLeft(6, '0') + "_" + periodInfor.DenThang.Replace("/", "").PadLeft(6, '0') + "_" + versionXML + ".xml")))) : (sFolder + dataFile + "_" + text2 + "_Q" + $"{periodInfor.Quy}{periodInfor.Nam}".PadLeft(6, '0') + "_" + periodInfor.TuThang.Replace("/", "") + "_" + periodInfor.DenThang.Replace("/", "") + "_" + versionXML + ".xml")) : ((text3.IndexOf("~" + periodInfor.menuID + "~") > 0) ? (sFolder + dataFile + "_" + text2 + "_Q" + $"{periodInfor.Quy}{periodInfor.Nam}".PadLeft(6, '0') + "_" + versionXML + ".xml") : ((!periodInfor.menuID.Equals("A1") && !periodInfor.menuID.Equals("A3") && !periodInfor.menuID.Equals("X2") && !"X1".Equals(WinSettings.TTPeriodInfo.menuID)) ? ((!periodInfor.menuID.Equals("01") && !periodInfor.menuID.Equals("I5")) ? ((!periodInfor.menuID.Equals("D1") && !WinSettings.MenuID.Equals("Z5")) ? ((!periodInfor.menuID.Equals("02") && !periodInfor.menuID.Equals("D2")) ? ((!periodInfor.menuID.Equals("71") && !periodInfor.menuID.Equals("D4") && !WinSettings.MenuID.Equals("Z3")) ? (sFolder + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_Q" + $"{periodInfor.Quy}{periodInfor.Nam}".PadLeft(6, '0') + "_" + versionXML + ".xml") : ((ConvertHelpers.ConvertToInt(periodInfor.Nam) < 2021 || periodInfor.ThuHo.Equals("")) ? (sFolder + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_Q" + $"{periodInfor.Quy}{periodInfor.Nam}".PadLeft(6, '0') + "_" + versionXML + ".xml") : (sFolder + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.ThuHo + "_Q" + $"{periodInfor.Quy}{periodInfor.Nam}".PadLeft(6, '0') + "_" + versionXML + ".xml"))) : ((ConvertHelpers.ConvertToInt(periodInfor.Nam) < 2021) ? (sFolder + dataFile + "_Q" + $"{periodInfor.Quy}{periodInfor.Nam}".PadLeft(6, '0') + "_" + versionXML + ".xml") : (sFolder + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.MaHSDuAnDT + "_Q" + $"{periodInfor.Quy}{periodInfor.Nam}".PadLeft(6, '0') + "_" + versionXML + ".xml"))) : (sFolder + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.TieuMucHachToanVal + "_Q" + $"{periodInfor.Quy}{periodInfor.Nam}".PadLeft(6, '0') + "_" + versionXML + ".xml")) : ((ConvertHelpers.ConvertToInt(periodInfor.Nam) < 2021) ? (sFolder + dataFile + "_Q" + $"{periodInfor.Quy}{periodInfor.Nam}".PadLeft(6, '0') + "_" + versionXML + ".xml") : (sFolder + dataFile + "_" + periodInfor.TieuMucHachToanVal + "_Q" + $"{periodInfor.Quy}{periodInfor.Nam}".PadLeft(6, '0') + "_" + versionXML + ".xml"))) : (sFolder + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_Q" + $"{periodInfor.Quy}{periodInfor.Nam}".PadLeft(6, '0') + "_" + versionXML + ".xml"))));
			}
			else if (!(periodInfor.KieuKy == "Y"))
			{
				text = ((periodInfor.KieuKy == "D") ? ((!HTKKFormater.isNullString(periodInfor.LanXuatBan) && !HTKKFormater.isNullString(periodInfor.TieuMucHachToanVal) && !HTKKFormater.isNullString(text2)) ? (sFolder + dataFile + ((!string.IsNullOrEmpty(periodInfor.DonViTinhMa)) ? ("_" + periodInfor.DonViTinhMa) : "") + "_" + text2 + "_L" + periodInfor.LanXuatBan + "_" + periodInfor.Ngay + periodInfor.Thang + periodInfor.Nam + "_" + versionXML + ".xml") : ((HTKKFormater.isNullString(periodInfor.LanXuatBan) && !HTKKFormater.isNullString(periodInfor.TieuMucHachToanVal) && !HTKKFormater.isNullString(text2)) ? ((!periodInfor.menuID.Equals("R6")) ? (sFolder + dataFile + ((!string.IsNullOrEmpty(periodInfor.DonViTinhMa)) ? ("_" + periodInfor.DonViTinhMa) : "") + "_" + text2 + "_" + periodInfor.Ngay + periodInfor.Thang + periodInfor.Nam + "_" + versionXML + ".xml") : (sFolder + dataFile + ((!string.IsNullOrEmpty(periodInfor.DonViTinhMa)) ? ("_" + periodInfor.DonViTinhMa) : "") + "_" + text2 + "_LHD" + periodInfor.ToChucCNDNVonChuyenNhuong + "_" + periodInfor.Ngay + periodInfor.Thang + periodInfor.Nam + "_" + versionXML + ".xml")) : ((!periodInfor.menuID.Equals("47") && !periodInfor.menuID.Equals("K7")) ? ((!periodInfor.menuID.Equals("71") && !periodInfor.menuID.Equals("D4") && !WinSettings.MenuID.Equals("Z3")) ? ((!periodInfor.menuID.Equals("T6") && !periodInfor.menuID.Equals("U2")) ? (periodInfor.menuID.Equals("Y5") ? (sFolder + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_CN" + periodInfor.TieuMucHachToanVal + "_" + periodInfor.TuNgay.Replace("/", "") + "_" + periodInfor.Ngay + periodInfor.Thang + periodInfor.Nam + "_" + versionXML + ".xml") : (("~O1~O2~O3~O4~O5~".IndexOf("~" + periodInfor.menuID + "~") >= 0) ? (sFolder + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + ((!string.IsNullOrEmpty(periodInfor.MaHSDuAnDT.Trim())) ? ("_" + periodInfor.MaHSDuAnDT.Trim()) : "") + "_" + periodInfor.Ngay + periodInfor.Thang + periodInfor.Nam + "_" + versionXML + ".xml") : (periodInfor.menuID.Equals("R5") ? (sFolder + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.TieuMucHachToanVal + "_" + periodInfor.Ngay + periodInfor.Thang + periodInfor.Nam + "_" + versionXML + ".xml") : ((!(periodInfor.menuID == "U1") && !(periodInfor.menuID == "Y3")) ? (sFolder + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.Ngay + periodInfor.Thang + periodInfor.Nam + "_" + versionXML + ".xml") : (sFolder + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_NN" + periodInfor.TieuMucHachToanVal + "_L" + periodInfor.ToChucCNDNVonChuyenNhuong + "_" + periodInfor.Ngay + periodInfor.Thang + periodInfor.Nam + "_" + versionXML + ".xml"))))) : (sFolder + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_CN" + periodInfor.TieuMucHachToanVal + "_" + periodInfor.Ngay + periodInfor.Thang + periodInfor.Nam + "_" + versionXML + ".xml")) : ((ConvertHelpers.ConvertToInt(periodInfor.Nam) < 2021 || periodInfor.ThuHo.Equals("")) ? (sFolder + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.Ngay + periodInfor.Thang + periodInfor.Nam + "_" + versionXML + ".xml") : (sFolder + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.ThuHo + "_" + periodInfor.Ngay + periodInfor.Thang + periodInfor.Nam + "_" + versionXML + ".xml"))) : (sFolder + dataFile + "_" + periodInfor.Ngay + periodInfor.Thang + periodInfor.Nam + "_" + (periodInfor.CTDC ? "CTDC" : "LNDC") + "_" + versionXML + ".xml")))) : ((!(periodInfor.KieuKy == "D_D")) ? (sFolder + dataFile) : (sFolder + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.TuNgay.Replace("/", "") + "_" + periodInfor.DenNgay.Replace("/", "") + "_" + versionXML + ".xml")));
			}
			else if (WinSettings.KieuKy.Equals("12"))
			{
				text = sFolder + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.Nam + "_" + periodInfor.TuThang.Replace("/", "") + "_" + periodInfor.DenThang.Replace("/", "") + "_" + versionXML + ".xml";
			}
			else if (periodInfor.menuID.Equals("93") || periodInfor.menuID.Equals("89") || periodInfor.menuID.Equals("44") || periodInfor.menuID.Equals("42") || text3.IndexOf("~" + periodInfor.menuID + "~") > 0)
			{
				text = sFolder + dataFile + ((!string.IsNullOrEmpty(periodInfor.DonViTinhMa)) ? ("_" + periodInfor.DonViTinhMa) : "") + "_" + text2 + "_" + periodInfor.Nam + "_" + periodInfor.TuThang.Replace("/", "") + "_" + periodInfor.DenThang.Replace("/", "") + "_" + versionXML + ".xml";
			}
			else if (periodInfor.menuID.Equals("95"))
			{
				text = sFolder + dataFile + "_L" + periodInfor.SoLan + "_" + periodInfor.Nam + "_" + versionXML + ".xml";
			}
			else if (periodInfor.menuID.Equals("G5"))
			{
				text = sFolder + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.Nam + "_" + periodInfor.TuThang.Replace("/", "") + "_" + periodInfor.DenThang.Replace("/", "") + "_" + versionXML + ".xml";
			}
			else if (periodInfor.menuID.Equals("63"))
			{
				text = sFolder + dataFile + "_" + periodInfor.Nam;
				bool flag = false;
				string text4 = "";
				if (periodInfor.dicVar.Count > 0)
				{
					flag = Convert.ToBoolean(periodInfor.dicVar["CoSoMoiThanhLap"]);
					text4 = (string)periodInfor.dicVar["NgayThanhLap"];
					if (flag)
					{
						text = text + "_CSM_" + text4.Replace("/", "");
					}
				}
				text = text + "_" + versionXML + ".xml";
			}
			else if (!periodInfor.menuID.Equals("L1"))
			{
				text = ((!periodInfor.menuID.Equals("K8") && !periodInfor.menuID.Equals("03") && !periodInfor.menuID.Equals("N8") && !periodInfor.menuID.Equals("J3") && !periodInfor.menuID.Equals("Q6") && !periodInfor.menuID.Equals("Q5")) ? (("C6~C7~T5~H4~Y4".IndexOf(periodInfor.menuID) >= 0) ? (sFolder + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.Nam + "_" + periodInfor.TuNgay.Replace("/", "") + "_" + periodInfor.DenNgay.Replace("/", "") + "_" + versionXML + ".xml") : ((!periodInfor.menuID.Equals("78") && !periodInfor.menuID.Equals("U9")) ? (periodInfor.menuID.Equals("53") ? (sFolder + dataFile + "_" + periodInfor.Nam + "_" + periodInfor.TuNgay.Replace("/", "") + "_" + periodInfor.DenNgay.Replace("/", "") + "_" + versionXML + ".xml") : (periodInfor.menuID.Equals("83") ? (sFolder + dataFile + "_" + periodInfor.Nam + "_" + periodInfor.TuNgay.Replace("/", "") + "_" + periodInfor.DenNgay.Replace("/", "") + "_" + versionXML + ".xml") : (periodInfor.menuID.Equals("E8") ? (sFolder + dataFile + "_" + periodInfor.TieuMucHachToanVal + "_" + periodInfor.Nam + "_" + periodInfor.TuNgay.Replace("/", "") + "_" + periodInfor.DenNgay.Replace("/", "") + "_" + versionXML + ".xml") : ((!periodInfor.menuID.Equals("I1") && !periodInfor.menuID.Equals("L9") && !periodInfor.menuID.Equals("I2")) ? (sFolder + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.Nam + "_" + versionXML + ".xml") : (sFolder + dataFile + "_L" + periodInfor.LanXuatBan + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.Nam + "_" + versionXML + ".xml"))))) : (sFolder + dataFile + "_" + periodInfor.Nam + "_" + periodInfor.TuNgay.Replace("/", "") + "_" + periodInfor.DenNgay.Replace("/", "") + "_" + periodInfor.DonViTinhMa + "_" + versionXML + ".xml"))) : (sFolder + dataFile + "_" + periodInfor.Nam + "_" + periodInfor.TuNgay.Replace("/", "") + "_" + periodInfor.DenNgay.Replace("/", "") + "_" + versionXML + ".xml"));
			}
			else
			{
				text = sFolder + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.Nam;
				bool flag = false;
				string text4 = "";
				if (periodInfor.dicVar.Count > 0)
				{
					flag = Convert.ToBoolean(periodInfor.dicVar["CoSoMoiThanhLap"]);
					text4 = (string)periodInfor.dicVar["NgayThanhLap"];
					if (flag)
					{
						text = text + "_CSM_" + text4.Replace("/", "");
					}
				}
				text = text + "_" + versionXML + ".xml";
			}
			return text;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public static Dictionary<string, string> GetCatalogue(string pathFile, string funName)
	{
		Dictionary<string, string> dictionary = new Dictionary<string, string>();
		try
		{
			if (HTKKFormater.isNullString(pathFile))
			{
				pathFile = HTKKConfig.INTERFACE_INI;
			}
			MethodInfo method = GetMethod(funName);
			if ((object)method != null)
			{
				ParameterInfo[] parameters = method.GetParameters();
				object[] array = new object[parameters.Length];
				for (int i = 0; i < parameters.Length; i++)
				{
					switch (parameters[i].Name)
					{
					case "pathFile":
						array[i] = pathFile;
						break;
					case "PeriodInfo":
						array[i] = WinSettings.TTPeriodInfo;
						break;
					}
				}
				return (Dictionary<string, string>)method.Invoke(WinSettings.InstanceClass, array);
			}
			return null;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public static Dictionary<string, string> GetCatalogueNew(List<PathDm> pathData, string funName)
	{
		Dictionary<string, string> dictionary = new Dictionary<string, string>();
		try
		{
			MethodInfo method = GetMethod(funName + "New");
			if ((object)method != null)
			{
				ParameterInfo[] parameters = method.GetParameters();
				object[] array = new object[parameters.Length];
				for (int i = 0; i < parameters.Length; i++)
				{
					switch (parameters[i].Name)
					{
					case "pathDms":
						array[i] = pathData;
						break;
					case "PeriodInfo":
						array[i] = WinSettings.TTPeriodInfo;
						break;
					}
				}
				return (Dictionary<string, string>)method.Invoke(WinSettings.InstanceClass, array);
			}
			return null;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public static string GetFileNameLoadKhbs(HTKKPeriodInfo periodInfor, string sDataFile, string sFolder, bool isLoadSavedFile)
	{
		try
		{
			string text = periodInfor.VersionXML;
			string text2 = "~98~92~08~40~42~44~M8~M9~K2~K3~J9~J8~J7~E2~E3~E4~E5~E6~E7~E9~D7~D8~D9~R9~R6~S1~";
			string text3 = "~J3~Q7~Q5~Q6~";
			string text4 = "";
			string text5 = "";
			int num = ConvertHelpers.ConvertToInt(periodInfor.SoLan);
			if (sDataFile.Contains("KHBS"))
			{
				text = WinSettings.ToKhaiInfo.VersionKBSXML;
			}
			else
			{
				num--;
			}
			string text6 = ConvertHelpers.ConvertToString(num);
			text = text.Replace(".", "");
			if (text2.IndexOf("~" + periodInfor.menuID + "~") > 0)
			{
				if (periodInfor.TieuMucHachToanVal.Equals("1"))
				{
					text4 = "DT";
				}
				else if (periodInfor.TieuMucHachToanVal.Equals("2"))
				{
					text4 = "KTN";
				}
				else if (periodInfor.TieuMucHachToanVal.Equals("3"))
				{
					text4 = "CD";
				}
				else if (periodInfor.TieuMucHachToanVal.Equals("4"))
				{
					text4 = "KDDT";
				}
				else if (periodInfor.TieuMucHachToanVal.Equals("5"))
				{
					text4 = "KDKTN";
				}
				else if (periodInfor.TieuMucHachToanVal.Equals("6"))
				{
					text4 = "KDCD";
				}
			}
			if (periodInfor.KieuKy == "M")
			{
				if (text2.IndexOf("~" + periodInfor.menuID + "~") > 0)
				{
					text5 = ((periodInfor.menuID == "R9") ? ((!(periodInfor.SoLan == "1")) ? (sFolder + "bs" + text6 + "_" + sDataFile + ((!string.IsNullOrEmpty(periodInfor.DonViTinhMa)) ? ("_" + periodInfor.DonViTinhMa) : "") + "_" + text4 + "_L" + periodInfor.LanXuatBan + "_" + periodInfor.Thang + periodInfor.Nam + "_" + text + ".xml") : (sFolder + sDataFile + ((!string.IsNullOrEmpty(periodInfor.DonViTinhMa)) ? ("_" + periodInfor.DonViTinhMa) : "") + "_" + text4 + "_L" + periodInfor.LanXuatBan + "_" + periodInfor.Thang + periodInfor.Nam + "_" + text + ".xml")) : ((periodInfor.menuID == "R6") ? ((!(periodInfor.SoLan == "1")) ? (sFolder + "bs" + text6 + "_" + sDataFile + ((!string.IsNullOrEmpty(periodInfor.DonViTinhMa)) ? ("_" + periodInfor.DonViTinhMa) : "") + "_" + text4 + "_LHD" + periodInfor.ToChucCNDNVonChuyenNhuong + "_" + periodInfor.Thang + periodInfor.Nam + "_" + text + ".xml") : (sFolder + sDataFile + ((!string.IsNullOrEmpty(periodInfor.DonViTinhMa)) ? ("_" + periodInfor.DonViTinhMa) : "") + "_" + text4 + "_LHD" + periodInfor.ToChucCNDNVonChuyenNhuong + "_" + periodInfor.Thang + periodInfor.Nam + "_" + text + ".xml")) : ((!(periodInfor.SoLan == "1")) ? (sFolder + "bs" + text6 + "_" + sDataFile + ((!string.IsNullOrEmpty(periodInfor.DonViTinhMa)) ? ("_" + periodInfor.DonViTinhMa) : "") + "_" + text4 + "_" + periodInfor.Thang + periodInfor.Nam + "_" + text + ".xml") : (sFolder + sDataFile + ((!string.IsNullOrEmpty(periodInfor.DonViTinhMa)) ? ("_" + periodInfor.DonViTinhMa) : "") + "_" + text4 + "_" + periodInfor.Thang + periodInfor.Nam + "_" + text + ".xml"))));
				}
				else
				{
					string text7 = "";
					text7 = ((!periodInfor.menuID.Equals("01") && !periodInfor.menuID.Equals("I5")) ? ((!periodInfor.menuID.Equals("D1") && !WinSettings.MenuID.Equals("Z5")) ? ((!periodInfor.menuID.Equals("02") && !periodInfor.menuID.Equals("D2")) ? ((periodInfor.menuID.Equals("71") || periodInfor.menuID.Equals("D4") || WinSettings.MenuID.Equals("Z3")) ? ((ConvertHelpers.ConvertToInt(periodInfor.Nam) < 2021 || periodInfor.ThuHo.Equals("")) ? (sFolder + "bs" + periodInfor.SoLan + "_" + sDataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.Thang + periodInfor.Nam + "_" + text + ".xml") : (sFolder + "bs" + periodInfor.SoLan + "_" + sDataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.ThuHo + "_" + periodInfor.Thang + periodInfor.Nam + "_" + text + ".xml")) : ((!periodInfor.menuID.Equals("R5")) ? (sFolder + "bs" + periodInfor.SoLan + "_" + sDataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.Thang + periodInfor.Nam + "_" + text + ".xml") : (sFolder + "bs" + periodInfor.SoLan + "_" + sDataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.TieuMucHachToanVal + "_" + periodInfor.Thang + periodInfor.Nam + "_" + text + ".xml"))) : ((ConvertHelpers.ConvertToInt(periodInfor.Nam) < 2021) ? (sFolder + "bs" + periodInfor.SoLan + "_" + sDataFile + "_" + periodInfor.Thang + periodInfor.Nam + "_" + text + ".xml") : (sFolder + "bs" + periodInfor.SoLan + "_" + sDataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.MaHSDuAnDT + "_" + periodInfor.Thang + periodInfor.Nam + "_" + text + ".xml"))) : (sFolder + "bs" + periodInfor.SoLan + "_" + sDataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.TieuMucHachToanVal + "_" + periodInfor.Thang + periodInfor.Nam + "_" + text + ".xml")) : ((ConvertHelpers.ConvertToInt(periodInfor.Nam) < 2021) ? (sFolder + "bs" + periodInfor.SoLan + "_" + sDataFile + "_" + periodInfor.Thang + periodInfor.Nam + "_" + text + ".xml") : (sFolder + "bs" + periodInfor.SoLan + "_" + sDataFile + "_" + periodInfor.TieuMucHachToanVal + "_" + periodInfor.Thang + periodInfor.Nam + "_" + text + ".xml")));
					text5 = ((periodInfor.SoLan == "1") ? ((File.Exists(text7) && isLoadSavedFile) ? text7 : ((!periodInfor.menuID.Equals("01") && !periodInfor.menuID.Equals("I5")) ? ((!periodInfor.menuID.Equals("D1") && !WinSettings.MenuID.Equals("Z5")) ? ((!periodInfor.menuID.Equals("02") && !periodInfor.menuID.Equals("D2")) ? ((!periodInfor.menuID.Equals("71") && !periodInfor.menuID.Equals("D4") && !WinSettings.MenuID.Equals("Z3")) ? ((!periodInfor.menuID.Equals("R5")) ? (sFolder + sDataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.Thang + periodInfor.Nam + "_" + text + ".xml") : (sFolder + sDataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.TieuMucHachToanVal + "_" + periodInfor.Thang + periodInfor.Nam + "_" + text + ".xml")) : ((ConvertHelpers.ConvertToInt(periodInfor.Nam) < 2021 || periodInfor.ThuHo.Equals("")) ? (sFolder + sDataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.Thang + periodInfor.Nam + "_" + text + ".xml") : (sFolder + sDataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.ThuHo + "_" + periodInfor.Thang + periodInfor.Nam + "_" + text + ".xml"))) : ((ConvertHelpers.ConvertToInt(periodInfor.Nam) < 2021) ? (sFolder + sDataFile + "_" + periodInfor.Thang + periodInfor.Nam + "_" + text + ".xml") : (sFolder + sDataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.MaHSDuAnDT + "_" + periodInfor.Thang + periodInfor.Nam + "_" + text + ".xml"))) : (sFolder + sDataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.TieuMucHachToanVal + "_" + periodInfor.Thang + periodInfor.Nam + "_" + text + ".xml")) : ((ConvertHelpers.ConvertToInt(periodInfor.Nam) < 2021) ? (sFolder + sDataFile + "_" + periodInfor.Thang + periodInfor.Nam + "_" + text + ".xml") : (sFolder + sDataFile + "_" + periodInfor.TieuMucHachToanVal + "_" + periodInfor.Thang + periodInfor.Nam + "_" + text + ".xml")))) : ((File.Exists(text7) && isLoadSavedFile) ? text7 : ((!periodInfor.menuID.Equals("01") && !periodInfor.menuID.Equals("I5")) ? ((!periodInfor.menuID.Equals("D1") && !WinSettings.MenuID.Equals("Z5")) ? ((!periodInfor.menuID.Equals("02") && !periodInfor.menuID.Equals("D2")) ? ((!periodInfor.menuID.Equals("71") && !periodInfor.menuID.Equals("D4") && !WinSettings.MenuID.Equals("Z3")) ? ((!periodInfor.menuID.Equals("R5")) ? (sFolder + "bs" + text6 + "_" + sDataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.Thang + periodInfor.Nam + "_" + text + ".xml") : (sFolder + "bs" + text6 + "_" + sDataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.TieuMucHachToanVal + "_" + periodInfor.Thang + periodInfor.Nam + "_" + text + ".xml")) : ((ConvertHelpers.ConvertToInt(periodInfor.Nam) < 2021 || periodInfor.ThuHo.Equals("")) ? (sFolder + "bs" + text6 + "_" + sDataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.Thang + periodInfor.Nam + "_" + text + ".xml") : (sFolder + "bs" + text6 + "_" + sDataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.ThuHo + "_" + periodInfor.Thang + periodInfor.Nam + "_" + text + ".xml"))) : ((ConvertHelpers.ConvertToInt(periodInfor.Nam) < 2021) ? (sFolder + "bs" + text6 + "_" + sDataFile + "_" + periodInfor.Thang + periodInfor.Nam + "_" + text + ".xml") : (sFolder + "bs" + text6 + "_" + sDataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.MaHSDuAnDT + "_" + periodInfor.Thang + periodInfor.Nam + "_" + text + ".xml"))) : (sFolder + "bs" + text6 + "_" + sDataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.TieuMucHachToanVal + "_" + periodInfor.Thang + periodInfor.Nam + "_" + text + ".xml")) : ((ConvertHelpers.ConvertToInt(periodInfor.Nam) < 2021) ? (sFolder + "bs" + text6 + "_" + sDataFile + "_" + periodInfor.Thang + periodInfor.Nam + "_" + text + ".xml") : (sFolder + "bs" + text6 + "_" + sDataFile + "_" + periodInfor.TieuMucHachToanVal + "_" + periodInfor.Thang + periodInfor.Nam + "_" + text + ".xml")))));
				}
			}
			else if (periodInfor.KieuKy == "Q")
			{
				if (!HTKKFormater.isNullString(periodInfor.TuThang) && !HTKKFormater.isNullString(periodInfor.DenThang))
				{
					if (periodInfor.menuID.Equals("42") || text2.IndexOf("~" + periodInfor.menuID + "~") > 0)
					{
						string text7 = sFolder + "bs" + periodInfor.SoLan + "_" + sDataFile + "_" + text4 + "_Q0" + periodInfor.Quy + periodInfor.Nam + "_" + periodInfor.TuThang.Replace("/", "") + "_" + periodInfor.DenThang.Replace("/", "") + "_" + text + ".xml";
						text5 = ((periodInfor.SoLan == "1") ? ((!File.Exists(text7) || !isLoadSavedFile) ? (sFolder + sDataFile + "_" + text4 + "_Q0" + periodInfor.Quy + periodInfor.Nam + "_" + periodInfor.TuThang.Replace("/", "") + "_" + periodInfor.DenThang.Replace("/", "") + "_" + text + ".xml") : text7) : ((!File.Exists(text7) || !isLoadSavedFile) ? (sFolder + "bs" + text6 + "_" + sDataFile + "_" + text4 + "_Q0" + periodInfor.Quy + periodInfor.Nam + "_" + periodInfor.TuThang.Replace("/", "") + "_" + periodInfor.DenThang.Replace("/", "") + "_" + text + ".xml") : text7));
					}
					else if (periodInfor.menuID.Equals("A4") || periodInfor.menuID.Equals("99") || periodInfor.menuID.Equals("H6"))
					{
						string text7 = "";
						text7 = ((!periodInfor.MaHSDuAnDT.Equals("")) ? (sFolder + "bs" + periodInfor.SoLan + "_" + sDataFile + "_" + periodInfor.MaHSDuAnDT + "_Q0" + periodInfor.TuThang.Replace("/", "") + "_0" + periodInfor.DenThang.Replace("/", "") + "_" + text + ".xml") : (sFolder + "bs" + periodInfor.SoLan + "_" + sDataFile + "_Q0" + periodInfor.TuThang.Replace("/", "") + "_0" + periodInfor.DenThang.Replace("/", "") + "_" + text + ".xml"));
						text5 = ((periodInfor.SoLan == "1") ? ((File.Exists(text7) && isLoadSavedFile) ? text7 : ((!periodInfor.MaHSDuAnDT.Equals("")) ? (sFolder + sDataFile + "_" + periodInfor.MaHSDuAnDT + "_Q0" + periodInfor.TuThang.Replace("/", "") + "_0" + periodInfor.DenThang.Replace("/", "") + "_" + text + ".xml") : (sFolder + sDataFile + "_Q0" + periodInfor.TuThang.Replace("/", "") + "_0" + periodInfor.DenThang.Replace("/", "") + "_" + text + ".xml"))) : ((!File.Exists(text7) || !isLoadSavedFile) ? (sFolder + "bs" + text6 + "_" + sDataFile + "_Q0" + periodInfor.TuThang.Replace("/", "") + "_0" + periodInfor.DenThang.Replace("/", "") + "_" + text + ".xml") : text7));
					}
					else if (periodInfor.SoLan == "1" && periodInfor.menuID == "35")
					{
						text5 = sFolder + "bs" + periodInfor.SoLan + "_" + sDataFile + "_Q0" + periodInfor.Quy + periodInfor.Nam + "_" + periodInfor.TuThang.Replace("/", "") + "_" + periodInfor.DenThang.Replace("/", "") + "_" + text + ".xml";
						if (!File.Exists(text5))
						{
							text5 = sFolder + sDataFile + "_Q0" + periodInfor.Quy + periodInfor.Nam + "_" + periodInfor.TuThang.Replace("/", "") + "_" + periodInfor.DenThang.Replace("/", "") + "_" + text + ".xml";
						}
					}
					else if (periodInfor.SoLan == "1")
					{
						string text7 = (text7 = sFolder + "bs" + periodInfor.SoLan + "_" + sDataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_Q0" + periodInfor.Quy + periodInfor.Nam + "_" + periodInfor.TuThang.Replace("/", "") + "_" + periodInfor.DenThang.Replace("/", "") + "_" + text + ".xml");
						text5 = ((!File.Exists(text7) || !isLoadSavedFile) ? (sFolder + sDataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_Q0" + periodInfor.Quy + periodInfor.Nam + "_" + periodInfor.TuThang.Replace("/", "") + "_" + periodInfor.DenThang.Replace("/", "") + "_" + text + ".xml") : text7);
					}
					else
					{
						string text7 = sFolder + "bs" + periodInfor.SoLan + "_" + sDataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_Q0" + periodInfor.Quy + periodInfor.Nam + "_" + periodInfor.TuThang.Replace("/", "") + "_" + periodInfor.DenThang.Replace("/", "") + "_" + text + ".xml";
						text5 = ((!File.Exists(text7) || !isLoadSavedFile) ? (sFolder + "bs" + text6 + "_" + sDataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_Q0" + periodInfor.Quy + periodInfor.Nam + "_" + periodInfor.TuThang.Replace("/", "") + "_" + periodInfor.DenThang.Replace("/", "") + "_" + text + ".xml") : text7);
					}
				}
				else if (text2.IndexOf("~" + periodInfor.menuID + "~") > 0)
				{
					string text7 = sFolder + "bs" + periodInfor.SoLan + "_" + sDataFile + "_" + text4 + "_Q0" + periodInfor.Quy + periodInfor.Nam + "_" + text + ".xml";
					text5 = ((periodInfor.SoLan == "1") ? ((!File.Exists(text7) || !isLoadSavedFile) ? (sFolder + sDataFile + "_" + text4 + "_Q0" + periodInfor.Quy + periodInfor.Nam + "_" + text + ".xml") : text7) : ((!File.Exists(text7) || !isLoadSavedFile) ? (sFolder + "bs" + text6 + "_" + sDataFile + "_" + text4 + "_Q0" + periodInfor.Quy + periodInfor.Nam + "_" + text + ".xml") : text7));
				}
				else if (periodInfor.menuID == "30")
				{
					string text7 = sFolder + "bs" + periodInfor.SoLan + "_" + sDataFile + "_" + periodInfor.TuNgay.Replace("/", "") + "_" + periodInfor.DenNgay.Replace("/", "") + "_" + text + ".xml";
					text5 = ((periodInfor.SoLan == "1" && isLoadSavedFile) ? ((!File.Exists(text7) || !isLoadSavedFile) ? (sFolder + sDataFile + "_" + periodInfor.TuNgay.Replace("/", "") + "_" + periodInfor.DenNgay.Replace("/", "") + "_" + text + ".xml") : text7) : ((!File.Exists(text5) || !isLoadSavedFile) ? (sFolder + "bs" + text6 + "_" + sDataFile + "_" + periodInfor.TuNgay.Replace("/", "") + "_" + periodInfor.DenNgay.Replace("/", "") + "_" + text + ".xml") : text7));
				}
				else if (periodInfor.menuID == "01" || periodInfor.menuID == "I5")
				{
					string text7 = "";
					text7 = ((ConvertHelpers.ConvertToInt(periodInfor.Nam) < 2021) ? (sFolder + "bs" + periodInfor.SoLan + "_" + sDataFile + "_Q0" + periodInfor.Quy + periodInfor.Nam + "_" + text + ".xml") : (sFolder + "bs" + periodInfor.SoLan + "_" + sDataFile + "_" + periodInfor.TieuMucHachToanVal + "_Q0" + periodInfor.Quy + periodInfor.Nam + "_" + text + ".xml"));
					text5 = ((periodInfor.SoLan == "1") ? ((File.Exists(text7) && isLoadSavedFile) ? text7 : ((ConvertHelpers.ConvertToInt(periodInfor.Nam) < 2021) ? (sFolder + sDataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_Q0" + periodInfor.Quy + periodInfor.Nam + "_" + text + ".xml") : (sFolder + sDataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.TieuMucHachToanVal + "_Q0" + periodInfor.Quy + periodInfor.Nam + "_" + text + ".xml"))) : ((File.Exists(text7) && isLoadSavedFile) ? text7 : ((ConvertHelpers.ConvertToInt(periodInfor.Nam) < 2021) ? (sFolder + "bs" + text6 + "_" + sDataFile + "_Q0" + periodInfor.Quy + periodInfor.Nam + "_" + text + ".xml") : (sFolder + "bs" + text6 + "_" + sDataFile + "_" + periodInfor.TieuMucHachToanVal + "_Q0" + periodInfor.Quy + periodInfor.Nam + "_" + text + ".xml"))));
				}
				else if (periodInfor.menuID == "D1" || WinSettings.MenuID.Equals("Z5"))
				{
					string text7 = "";
					text7 = sFolder + "bs" + periodInfor.SoLan + "_" + sDataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.TieuMucHachToanVal + "_Q0" + periodInfor.Quy + periodInfor.Nam + "_" + text + ".xml";
					text5 = ((periodInfor.SoLan == "1") ? ((!File.Exists(text7) || !isLoadSavedFile) ? (sFolder + sDataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.TieuMucHachToanVal + "_Q0" + periodInfor.Quy + periodInfor.Nam + "_" + text + ".xml") : text7) : ((!File.Exists(text7) || !isLoadSavedFile) ? (sFolder + "bs" + text6 + "_" + sDataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.TieuMucHachToanVal + "_Q0" + periodInfor.Quy + periodInfor.Nam + "_" + text + ".xml") : text7));
				}
				else if (periodInfor.menuID == "02" || periodInfor.menuID == "D2")
				{
					string text7 = "";
					text7 = ((ConvertHelpers.ConvertToInt(periodInfor.Nam) < 2021) ? (sFolder + "bs" + periodInfor.SoLan + "_" + sDataFile + "_Q0" + periodInfor.Quy + periodInfor.Nam + "_" + text + ".xml") : (sFolder + "bs" + periodInfor.SoLan + "_" + sDataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.MaHSDuAnDT + "_Q0" + periodInfor.Quy + periodInfor.Nam + "_" + text + ".xml"));
					text5 = ((periodInfor.SoLan == "1") ? ((File.Exists(text7) && isLoadSavedFile) ? text7 : ((ConvertHelpers.ConvertToInt(periodInfor.Nam) < 2021) ? (sFolder + sDataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_Q0" + periodInfor.Quy + periodInfor.Nam + "_" + text + ".xml") : (sFolder + sDataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.MaHSDuAnDT + "_Q0" + periodInfor.Quy + periodInfor.Nam + "_" + text + ".xml"))) : ((File.Exists(text7) && isLoadSavedFile) ? text7 : ((ConvertHelpers.ConvertToInt(periodInfor.Nam) < 2021) ? (sFolder + "bs" + text6 + "_" + sDataFile + "_Q0" + periodInfor.Quy + periodInfor.Nam + "_" + text + ".xml") : (sFolder + "bs" + text6 + "_" + sDataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.MaHSDuAnDT + "_Q0" + periodInfor.Quy + periodInfor.Nam + "_" + text + ".xml"))));
				}
				else if (periodInfor.menuID == "71" || periodInfor.menuID == "D4" || WinSettings.MenuID.Equals("Z3"))
				{
					string text7 = "";
					text7 = ((ConvertHelpers.ConvertToInt(periodInfor.Nam) < 2021 || periodInfor.ThuHo.Equals("")) ? (sFolder + "bs" + periodInfor.SoLan + "_" + sDataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_Q0" + periodInfor.Quy + periodInfor.Nam + "_" + text + ".xml") : (sFolder + "bs" + periodInfor.SoLan + "_" + sDataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.ThuHo + "_Q0" + periodInfor.Quy + periodInfor.Nam + "_" + text + ".xml"));
					text5 = ((periodInfor.SoLan == "1") ? ((File.Exists(text7) && isLoadSavedFile) ? text7 : ((ConvertHelpers.ConvertToInt(periodInfor.Nam) < 2021 || periodInfor.ThuHo.Equals("")) ? (sFolder + sDataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_Q0" + periodInfor.Quy + periodInfor.Nam + "_" + text + ".xml") : (sFolder + sDataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.ThuHo + "_Q0" + periodInfor.Quy + periodInfor.Nam + "_" + text + ".xml"))) : ((File.Exists(text7) && isLoadSavedFile) ? text7 : ((ConvertHelpers.ConvertToInt(periodInfor.Nam) < 2021 || periodInfor.ThuHo.Equals("")) ? (sFolder + "bs" + text6 + "_" + sDataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_Q0" + periodInfor.Quy + periodInfor.Nam + "_" + text + ".xml") : (sFolder + "bs" + text6 + "_" + sDataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.ThuHo + "_Q0" + periodInfor.Quy + periodInfor.Nam + "_" + text + ".xml"))));
				}
				else
				{
					string text7 = sFolder + "bs" + periodInfor.SoLan + "_" + sDataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_Q0" + periodInfor.Quy + periodInfor.Nam + "_" + text + ".xml";
					text5 = ((periodInfor.SoLan == "1") ? ((!File.Exists(text7) || !isLoadSavedFile) ? (sFolder + sDataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_Q0" + periodInfor.Quy + periodInfor.Nam + "_" + text + ".xml") : text7) : ((!File.Exists(text7) || !isLoadSavedFile) ? (sFolder + "bs" + text6 + "_" + sDataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_Q0" + periodInfor.Quy + periodInfor.Nam + "_" + text + ".xml") : text7));
				}
			}
			else if (periodInfor.KieuKy == "Y")
			{
				if (WinSettings.KieuKy.Equals("12") || WinSettings.KieuKy.Equals("35"))
				{
					string text7 = sFolder + "bs" + periodInfor.SoLan + "_" + sDataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.Nam + "_" + periodInfor.TuThang.Replace("/", "") + "_" + periodInfor.DenThang.Replace("/", "") + "_" + text + ".xml";
					text5 = ((periodInfor.SoLan == "1") ? ((!File.Exists(text7) || !isLoadSavedFile) ? (sFolder + sDataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.Nam + "_" + periodInfor.TuThang.Replace("/", "") + "_" + periodInfor.DenThang.Replace("/", "") + "_" + text + ".xml") : text7) : ((!File.Exists(text7) || !isLoadSavedFile) ? (sFolder + "bs" + text6 + "_" + sDataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.Nam + "_" + periodInfor.TuThang.Replace("/", "") + "_" + periodInfor.DenThang.Replace("/", "") + "_" + text + ".xml") : text7));
				}
				else if (periodInfor.menuID.Equals("93") || periodInfor.menuID.Equals("89") || periodInfor.menuID.Equals("44") || periodInfor.menuID.Equals("42") || text2.IndexOf("~" + periodInfor.menuID + "~") > 0)
				{
					string text7 = sFolder + "bs" + periodInfor.SoLan + "_" + sDataFile + ((!string.IsNullOrEmpty(periodInfor.DonViTinhMa)) ? ("_" + periodInfor.DonViTinhMa) : "") + "_" + text4 + "_" + periodInfor.Nam + "_" + periodInfor.TuThang.Replace("/", "") + "_" + periodInfor.DenThang.Replace("/", "") + "_" + text + ".xml";
					text5 = ((periodInfor.SoLan == "1") ? ((!File.Exists(text7) || !isLoadSavedFile) ? (sFolder + sDataFile + ((!string.IsNullOrEmpty(periodInfor.DonViTinhMa)) ? ("_" + periodInfor.DonViTinhMa) : "") + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + text4 + "_" + periodInfor.Nam + "_" + periodInfor.TuThang.Replace("/", "") + "_" + periodInfor.DenThang.Replace("/", "") + "_" + text + ".xml") : text7) : ((!File.Exists(text7) || !isLoadSavedFile) ? (sFolder + "bs" + text6 + "_" + sDataFile + ((!string.IsNullOrEmpty(periodInfor.DonViTinhMa)) ? ("_" + periodInfor.DonViTinhMa) : "") + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + text4 + "_" + periodInfor.Nam + "_" + periodInfor.TuThang.Replace("/", "") + "_" + periodInfor.DenThang.Replace("/", "") + "_" + text + ".xml") : text7));
				}
				else if (periodInfor.menuID.Equals("63"))
				{
					text5 = sFolder + "bs" + periodInfor.SoLan + "_" + sDataFile + "_" + periodInfor.Nam;
					bool flag = false;
					string text8 = "";
					if (periodInfor.dicVar.Count > 0)
					{
						flag = Convert.ToBoolean(periodInfor.dicVar["CoSoMoiThanhLap"]);
						text8 = (string)periodInfor.dicVar["NgayThanhLap"];
						if (flag)
						{
							text5 = text5 + "_CSM_" + text8.Replace("/", "");
						}
					}
					text5 = text5 + "_" + text + ".xml";
					if (periodInfor.SoLan == "1")
					{
						if (!File.Exists(text5) || !isLoadSavedFile)
						{
							text5 = text5.Replace("bs" + periodInfor.SoLan + "_", "");
						}
					}
					else if (!File.Exists(text5) || !isLoadSavedFile)
					{
						text5 = text5.Replace("bs" + periodInfor.SoLan + "_", "bs" + text6 + "_");
					}
				}
				else if (periodInfor.menuID.Equals("L1"))
				{
					text5 = sFolder + "bs" + periodInfor.SoLan + "_" + sDataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.Nam;
					bool flag = false;
					string text8 = "";
					if (periodInfor.dicVar.Count > 0)
					{
						flag = Convert.ToBoolean(periodInfor.dicVar["CoSoMoiThanhLap"]);
						text8 = (string)periodInfor.dicVar["NgayThanhLap"];
						if (flag)
						{
							text5 = text5 + "_CSM_" + text8.Replace("/", "");
						}
					}
					text5 = text5 + "_" + text + ".xml";
					if (periodInfor.SoLan == "1")
					{
						if (!File.Exists(text5) || !isLoadSavedFile)
						{
							text5 = text5.Replace("bs" + periodInfor.SoLan + "_", "");
						}
					}
					else if (!File.Exists(text5) || !isLoadSavedFile)
					{
						text5 = text5.Replace("bs" + periodInfor.SoLan + "_", "bs" + text6 + "_");
					}
				}
				else if (!periodInfor.menuID.Equals("95"))
				{
					string text7 = "";
					text7 = ((!periodInfor.menuID.Equals("K8") && !periodInfor.menuID.Equals("03") && !periodInfor.menuID.Equals("83") && !periodInfor.menuID.Equals("N8") && text3.IndexOf("~" + periodInfor.menuID + "~") < 0) ? (periodInfor.menuID.Equals("E8") ? (sFolder + "bs" + periodInfor.SoLan + "_" + sDataFile + "_" + periodInfor.TieuMucHachToanVal + "_" + periodInfor.Nam + "_" + periodInfor.TuNgay.Replace("/", "") + "_" + periodInfor.DenNgay.Replace("/", "") + "_" + text + ".xml") : ((!periodInfor.menuID.Equals("I1") && !periodInfor.menuID.Equals("L9") && !periodInfor.menuID.Equals("I2")) ? (sFolder + "bs" + periodInfor.SoLan + "_" + sDataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.Nam + "_" + text + ".xml") : (sFolder + "bs" + periodInfor.SoLan + "_" + sDataFile + "_L" + periodInfor.LanXuatBan + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.Nam + "_" + text + ".xml"))) : (sFolder + "bs" + periodInfor.SoLan + "_" + sDataFile + "_" + periodInfor.Nam + "_" + periodInfor.TuNgay.Replace("/", "") + "_" + periodInfor.DenNgay.Replace("/", "") + "_" + text + ".xml"));
					if (!(periodInfor.SoLan == "1"))
					{
						text5 = ((File.Exists(text7) && isLoadSavedFile) ? text7 : ((!"30,I5,K9,K1,M7,K6,K8,73,55,71,31,P2,X6".Contains(WinSettings.MenuID)) ? (periodInfor.menuID.Equals("83") ? (sFolder + "bs" + text6 + "_" + sDataFile + "_" + periodInfor.Nam + "_" + periodInfor.TuNgay.Replace("/", "") + "_" + periodInfor.DenNgay.Replace("/", "") + "_" + text + ".xml") : (periodInfor.menuID.Equals("E8") ? (sFolder + "bs" + text6 + "_" + sDataFile + "_" + periodInfor.TieuMucHachToanVal + "_" + periodInfor.Nam + "_" + periodInfor.TuNgay.Replace("/", "") + "_" + periodInfor.DenNgay.Replace("/", "") + "_" + text + ".xml") : ((!periodInfor.menuID.Equals("K8") && !periodInfor.menuID.Equals("03") && !periodInfor.menuID.Equals("83") && !periodInfor.menuID.Equals("N8") && text3.IndexOf("~" + periodInfor.menuID + "~") < 0) ? (periodInfor.menuID.Equals("E8") ? (sFolder + "bs" + text6 + "_" + sDataFile + "_" + periodInfor.TieuMucHachToanVal + "_" + periodInfor.Nam + "_" + periodInfor.TuNgay.Replace("/", "") + "_" + periodInfor.DenNgay.Replace("/", "") + "_" + text + ".xml") : ((!periodInfor.menuID.Equals("I1") && !periodInfor.menuID.Equals("L9") && !periodInfor.menuID.Equals("I2")) ? (sFolder + "bs" + text6 + "_" + sDataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.Nam + "_" + text + ".xml") : (sFolder + "bs" + text6 + "_" + sDataFile + "_L" + periodInfor.LanXuatBan + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.Nam + "_" + text + ".xml"))) : (sFolder + "bs" + text6 + "_" + sDataFile + "_" + periodInfor.Nam + "_" + periodInfor.TuNgay.Replace("/", "") + "_" + periodInfor.DenNgay.Replace("/", "") + "_" + text + ".xml")))) : ((!periodInfor.menuID.Equals("55")) ? (sFolder + "bs" + text6 + "_" + sDataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.Nam + "_" + periodInfor.TuNgay.Replace("/", "") + "_" + periodInfor.DenNgay.Replace("/", "") + "_" + text + ".xml") : (sFolder + "bs" + text6 + "_" + sDataFile + "_" + periodInfor.Nam + "_" + text + ".xml"))));
					}
					else if (File.Exists(text7) && isLoadSavedFile)
					{
						text5 = text7;
					}
					else if ("30,I5,K9,K1,M7,K6,K8,73,55,71,31,P2,X6".Contains(WinSettings.MenuID))
					{
						text5 = ((!periodInfor.menuID.Equals("55")) ? (sFolder + sDataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.Nam + "_" + periodInfor.TuNgay.Replace("/", "") + "_" + periodInfor.DenNgay.Replace("/", "") + "_" + text + ".xml") : (sFolder + sDataFile + "_" + periodInfor.Nam + "_" + text + ".xml"));
					}
					else if (periodInfor.menuID.Equals("K8") || periodInfor.menuID.Equals("03") || periodInfor.menuID.Equals("83") || periodInfor.menuID.Equals("N8") || text3.IndexOf("~" + periodInfor.menuID + "~") >= 0)
					{
						text5 = sFolder + sDataFile + "_" + periodInfor.Nam + "_" + periodInfor.TuNgay.Replace("/", "") + "_" + periodInfor.DenNgay.Replace("/", "") + "_" + text + ".xml";
					}
					else if (periodInfor.menuID.Equals("E8"))
					{
						text5 = sFolder + sDataFile + "_" + periodInfor.TieuMucHachToanVal + "_" + periodInfor.Nam + "_" + periodInfor.TuNgay.Replace("/", "") + "_" + periodInfor.DenNgay.Replace("/", "") + "_" + text + ".xml";
					}
					else if (periodInfor.menuID.Equals("I1") || periodInfor.menuID.Equals("L9") || periodInfor.menuID.Equals("I2"))
					{
						text5 = sFolder + sDataFile + "_L" + periodInfor.LanXuatBan + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.Nam + "_" + text + ".xml";
					}
					else if (!"30,I5,K9,K1,M7,K6,K8,73,55,71,31,P2,X6".Contains(WinSettings.MenuID))
					{
						text5 = sFolder + sDataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.Nam + "_" + text + ".xml";
					}
				}
			}
			else if (periodInfor.KieuKy == "D")
			{
				if (periodInfor.menuID.Equals("G4") || periodInfor.menuID.Equals("N5") || periodInfor.menuID.Equals("H5") || periodInfor.menuID.Equals("O7") || periodInfor.menuID.Equals("P5") || periodInfor.menuID.Equals("U6") || "U8~".Contains(periodInfor.menuID))
				{
					string text7 = sFolder + "bs_" + sDataFile + "_" + periodInfor.Ngay + periodInfor.Thang + periodInfor.Nam + "_" + text + ".xml";
					if (File.Exists(text7) && isLoadSavedFile)
					{
						text5 = text7;
					}
					else
					{
						bool flag2 = true;
						DirectoryInfo directoryInfo = new DirectoryInfo(HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst);
						FileInfo[] files = directoryInfo.GetFiles("bs_" + sDataFile + "*.xml");
						int num2 = 0;
						if (num2 < files.Length)
						{
							FileInfo fileInfo = files[num2];
							flag2 = false;
						}
						if (flag2)
						{
							files = directoryInfo.GetFiles(sDataFile + "*.xml");
							foreach (FileInfo fileInfo in files)
							{
								text5 = fileInfo.FullName;
							}
						}
						else
						{
							string text9 = "";
							string obj = periodInfor.Nam + periodInfor.Thang + periodInfor.Ngay;
							files = directoryInfo.GetFiles("bs_" + sDataFile + "*.xml");
							foreach (FileInfo fileInfo in files)
							{
								string name = fileInfo.Name;
								string obj2 = "";
								string[] array = name.Split(new char[1] { '_' });
								double num3 = 0.0;
								if (array.Length == 6 && array[4].Length == 8)
								{
									obj2 = array[4].Substring(4, 4) + array[4].Substring(2, 2) + array[4].Substring(0, 2);
								}
								if (ConvertHelpers.ConvertToDouble(obj) > ConvertHelpers.ConvertToDouble(obj2) && num3 < ConvertHelpers.ConvertToDouble(obj2))
								{
									num3 = ConvertHelpers.ConvertToDouble(obj2);
									text5 = fileInfo.FullName;
								}
							}
						}
					}
				}
				else if (!HTKKFormater.isNullString(periodInfor.LanXuatBan) && !HTKKFormater.isNullString(periodInfor.TieuMucHachToanVal) && !HTKKFormater.isNullString(text4))
				{
					string text7 = sFolder + "bs" + periodInfor.SoLan + "_" + sDataFile + ((!string.IsNullOrEmpty(periodInfor.DonViTinhMa)) ? ("_" + periodInfor.DonViTinhMa) : "") + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + text4 + "_L" + periodInfor.LanXuatBan + "_" + periodInfor.Ngay + periodInfor.Thang + periodInfor.Nam + "_" + text + ".xml";
					text5 = ((periodInfor.SoLan == "1") ? ((!File.Exists(text7) || !isLoadSavedFile) ? (sFolder + sDataFile + ((!string.IsNullOrEmpty(periodInfor.DonViTinhMa)) ? ("_" + periodInfor.DonViTinhMa) : "") + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + text4 + "_L" + periodInfor.LanXuatBan + "_" + periodInfor.Ngay + periodInfor.Thang + periodInfor.Nam + "_" + text + ".xml") : text7) : ((!File.Exists(text7) || !isLoadSavedFile) ? (sFolder + "bs" + text6 + "_" + sDataFile + ((!string.IsNullOrEmpty(periodInfor.DonViTinhMa)) ? ("_" + periodInfor.DonViTinhMa) : "") + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + text4 + "_L" + periodInfor.LanXuatBan + "_" + periodInfor.Ngay + periodInfor.Thang + periodInfor.Nam + "_" + text + ".xml") : text7));
				}
				else if (HTKKFormater.isNullString(periodInfor.LanXuatBan) && !HTKKFormater.isNullString(periodInfor.TieuMucHachToanVal) && !HTKKFormater.isNullString(text4))
				{
					if (periodInfor.menuID.Equals("R6"))
					{
						string text7 = sFolder + "bs" + periodInfor.SoLan + "_" + sDataFile + ((!string.IsNullOrEmpty(periodInfor.DonViTinhMa)) ? ("_" + periodInfor.DonViTinhMa) : "") + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + text4 + "_LHD" + periodInfor.ToChucCNDNVonChuyenNhuong + "_" + periodInfor.Ngay + periodInfor.Thang + periodInfor.Nam + "_" + text + ".xml";
						text5 = ((periodInfor.SoLan == "1") ? ((!File.Exists(text7) || !isLoadSavedFile) ? (sFolder + sDataFile + ((!string.IsNullOrEmpty(periodInfor.DonViTinhMa)) ? ("_" + periodInfor.DonViTinhMa) : "") + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + text4 + "_LHD" + periodInfor.ToChucCNDNVonChuyenNhuong + "_" + periodInfor.Ngay + periodInfor.Thang + periodInfor.Nam + "_" + text + ".xml") : text7) : ((!File.Exists(text7) || !isLoadSavedFile) ? (sFolder + "bs" + num + "_" + sDataFile + ((!string.IsNullOrEmpty(periodInfor.DonViTinhMa)) ? ("_" + periodInfor.DonViTinhMa) : "") + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + text4 + "_LHD" + periodInfor.ToChucCNDNVonChuyenNhuong + "_" + periodInfor.Ngay + periodInfor.Thang + periodInfor.Nam + "_" + text + ".xml") : text7));
					}
					else
					{
						string text7 = sFolder + "bs" + periodInfor.SoLan + "_" + sDataFile + ((!string.IsNullOrEmpty(periodInfor.DonViTinhMa)) ? ("_" + periodInfor.DonViTinhMa) : "") + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + text4 + "_" + periodInfor.Ngay + periodInfor.Thang + periodInfor.Nam + "_" + text + ".xml";
						text5 = ((periodInfor.SoLan == "1") ? ((!File.Exists(text7) || !isLoadSavedFile) ? (sFolder + sDataFile + ((!string.IsNullOrEmpty(periodInfor.DonViTinhMa)) ? ("_" + periodInfor.DonViTinhMa) : "") + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + text4 + "_" + periodInfor.Ngay + periodInfor.Thang + periodInfor.Nam + "_" + text + ".xml") : text7) : ((!File.Exists(text7) || !isLoadSavedFile) ? (sFolder + "bs" + num + "_" + sDataFile + ((!string.IsNullOrEmpty(periodInfor.DonViTinhMa)) ? ("_" + periodInfor.DonViTinhMa) : "") + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + text4 + "_" + periodInfor.Ngay + periodInfor.Thang + periodInfor.Nam + "_" + text + ".xml") : text7));
					}
				}
				else
				{
					string text7 = sFolder + "bs" + periodInfor.SoLan + "_" + sDataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.Ngay + periodInfor.Thang + periodInfor.Nam + "_" + text + ".xml";
					if (periodInfor.menuID.Equals("47") || periodInfor.menuID.Equals("K7"))
					{
						text7 = sFolder + "bs" + periodInfor.SoLan + "_" + sDataFile + "_" + periodInfor.Ngay + periodInfor.Thang + periodInfor.Nam + "_" + (periodInfor.CTDC ? "CTDC" : "LNDC") + "_" + text + ".xml";
					}
					if (periodInfor.menuID.Equals("R5"))
					{
						text7 = sFolder + "bs" + periodInfor.SoLan + "_" + sDataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.TieuMucHachToanVal + "_" + periodInfor.Ngay + periodInfor.Thang + periodInfor.Nam + "_" + text + ".xml";
					}
					text5 = ((periodInfor.SoLan == "1") ? ((File.Exists(text7) && isLoadSavedFile) ? text7 : ((!periodInfor.menuID.Equals("47") && !periodInfor.menuID.Equals("K7")) ? ((!periodInfor.menuID.Equals("71") && !periodInfor.menuID.Equals("D4") && !WinSettings.MenuID.Equals("Z3")) ? ((!periodInfor.menuID.Equals("T6") && !periodInfor.menuID.Equals("U2")) ? (periodInfor.menuID.Equals("Y5") ? (sFolder + sDataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_CN" + periodInfor.TieuMucHachToanVal + "_" + periodInfor.TuNgay.Replace("/", "") + "_" + periodInfor.Ngay + periodInfor.Thang + periodInfor.Nam + "_" + text + ".xml") : ((!periodInfor.menuID.Equals("R5")) ? (sFolder + sDataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.Ngay + periodInfor.Thang + periodInfor.Nam + "_" + text + ".xml") : (sFolder + sDataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.TieuMucHachToanVal + "_" + periodInfor.Ngay + periodInfor.Thang + periodInfor.Nam + "_" + text + ".xml"))) : (sFolder + sDataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_CN" + periodInfor.TieuMucHachToanVal + "_" + periodInfor.Ngay + periodInfor.Thang + periodInfor.Nam + "_" + text + ".xml")) : ((ConvertHelpers.ConvertToInt(periodInfor.Nam) < 2021 || periodInfor.ThuHo.Equals("")) ? (sFolder + sDataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.Ngay + periodInfor.Thang + periodInfor.Nam + "_" + text + ".xml") : (sFolder + sDataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.ThuHo + "_" + periodInfor.Ngay + periodInfor.Thang + periodInfor.Nam + "_" + text + ".xml"))) : (sFolder + sDataFile + "_" + periodInfor.Ngay + periodInfor.Thang + periodInfor.Nam + "_" + (periodInfor.CTDC ? "CTDC" : "LNDC") + "_" + text + ".xml"))) : ((File.Exists(text7) && isLoadSavedFile) ? text7 : ((!periodInfor.menuID.Equals("47") && !periodInfor.menuID.Equals("K7")) ? ((!periodInfor.menuID.Equals("71") && !periodInfor.menuID.Equals("D4") && !WinSettings.MenuID.Equals("Z3")) ? ((!periodInfor.menuID.Equals("T6") && !periodInfor.menuID.Equals("U2")) ? (periodInfor.menuID.Equals("Y5") ? (sFolder + "bs" + num + "_" + sDataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_CN" + periodInfor.TieuMucHachToanVal + "_" + periodInfor.TuNgay.Replace("/", "") + "_" + periodInfor.Ngay + periodInfor.Thang + periodInfor.Nam + "_" + text + ".xml") : ((!periodInfor.menuID.Equals("R5")) ? (sFolder + "bs" + num + "_" + sDataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.Ngay + periodInfor.Thang + periodInfor.Nam + "_" + text + ".xml") : (sFolder + "bs" + num + "_" + sDataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.TieuMucHachToanVal + "_" + periodInfor.Ngay + periodInfor.Thang + periodInfor.Nam + "_" + text + ".xml"))) : (sFolder + "bs" + num + "_" + sDataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_CN" + periodInfor.TieuMucHachToanVal + "_" + periodInfor.Ngay + periodInfor.Thang + periodInfor.Nam + "_" + text + ".xml")) : ((ConvertHelpers.ConvertToInt(periodInfor.Nam) < 2021 || periodInfor.ThuHo.Equals("")) ? (sFolder + "bs" + num + "_" + sDataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.Ngay + periodInfor.Thang + periodInfor.Nam + "_" + text + ".xml") : (sFolder + "bs" + num + "_" + sDataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.ThuHo + "_" + periodInfor.Ngay + periodInfor.Thang + periodInfor.Nam + "_" + text + ".xml"))) : (sFolder + "bs" + num + "_" + sDataFile + "_" + periodInfor.Ngay + periodInfor.Thang + periodInfor.Nam + "_" + (periodInfor.CTDC ? "CTDC" : "LNDC") + "_" + text + ".xml"))));
				}
			}
			else if (periodInfor.KieuKy == "D_D")
			{
				string text7 = sFolder + "bs" + periodInfor.SoLan + "_" + sDataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.TuNgay.Replace("/", "") + "_" + periodInfor.DenNgay.Replace("/", "") + "_" + text + ".xml";
				text5 = ((periodInfor.SoLan == "1") ? ((!File.Exists(text7) || !isLoadSavedFile) ? (sFolder + sDataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.TuNgay.Replace("/", "") + "_" + periodInfor.DenNgay.Replace("/", "") + "_" + text + ".xml") : text7) : ((!File.Exists(text7) || !isLoadSavedFile) ? (sFolder + "bs" + num + "_" + sDataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.TuNgay.Replace("/", "") + "_" + periodInfor.DenNgay.Replace("/", "") + "_" + text + ".xml") : text7));
			}
			else
			{
				text5 = sFolder + "bs" + num + "_" + sDataFile;
			}
			return text5;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public static bool CheckChonPlBctc(HTKKPeriodInfo periodInfo, FpSpread winFpSpread)
	{
		//IL_0099: Unknown result type (might be due to invalid IL or missing references)
		//IL_0166: Unknown result type (might be due to invalid IL or missing references)
		//IL_0252: Unknown result type (might be due to invalid IL or missing references)
		//IL_01ba: Unknown result type (might be due to invalid IL or missing references)
		//IL_0392: Unknown result type (might be due to invalid IL or missing references)
		//IL_02a6: Unknown result type (might be due to invalid IL or missing references)
		//IL_042a: Unknown result type (might be due to invalid IL or missing references)
		//IL_02fa: Unknown result type (might be due to invalid IL or missing references)
		//IL_05e6: Unknown result type (might be due to invalid IL or missing references)
		//IL_047e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0a44: Unknown result type (might be due to invalid IL or missing references)
		//IL_074e: Unknown result type (might be due to invalid IL or missing references)
		//IL_063a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0ab1: Unknown result type (might be due to invalid IL or missing references)
		//IL_090a: Unknown result type (might be due to invalid IL or missing references)
		//IL_07a2: Unknown result type (might be due to invalid IL or missing references)
		//IL_0b05: Unknown result type (might be due to invalid IL or missing references)
		//IL_095e: Unknown result type (might be due to invalid IL or missing references)
		//IL_04fa: Unknown result type (might be due to invalid IL or missing references)
		//IL_0cda: Unknown result type (might be due to invalid IL or missing references)
		//IL_06b6: Unknown result type (might be due to invalid IL or missing references)
		//IL_054e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0d9b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0d2e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0b9d: Unknown result type (might be due to invalid IL or missing references)
		//IL_081e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0def: Unknown result type (might be due to invalid IL or missing references)
		//IL_0bf1: Unknown result type (might be due to invalid IL or missing references)
		//IL_09da: Unknown result type (might be due to invalid IL or missing references)
		//IL_0872: Unknown result type (might be due to invalid IL or missing references)
		//IL_0ed2: Unknown result type (might be due to invalid IL or missing references)
		//IL_0e76: Unknown result type (might be due to invalid IL or missing references)
		//IL_0c6d: Unknown result type (might be due to invalid IL or missing references)
		//IL_1383: Unknown result type (might be due to invalid IL or missing references)
		//IL_152a: Unknown result type (might be due to invalid IL or missing references)
		//IL_159f: Unknown result type (might be due to invalid IL or missing references)
		//IL_11e3: Unknown result type (might be due to invalid IL or missing references)
		//IL_0fcc: Unknown result type (might be due to invalid IL or missing references)
		//IL_1405: Unknown result type (might be due to invalid IL or missing references)
		//IL_1266: Unknown result type (might be due to invalid IL or missing references)
		//IL_148e: Unknown result type (might be due to invalid IL or missing references)
		//IL_12d7: Unknown result type (might be due to invalid IL or missing references)
		//IL_1175: Unknown result type (might be due to invalid IL or missing references)
		bool result = true;
		try
		{
			string menuID = periodInfo.menuID;
			string nam = periodInfo.Nam;
			if ("69,19,22".Contains(menuID) && (bool)winFpSpread.Sheets[0].Cells[1, 0].Value && (bool)winFpSpread.Sheets[0].Cells[2, 0].Value)
			{
				frmMessage frmMessage = new frmMessage("0314", "2", "2", "Thông báo");
				((Form)frmMessage).ShowDialog();
				if (!frmMessage.Yes)
				{
					result = false;
					return result;
				}
			}
			if ("78,53,54,58,A1,H1,A8,H8".Contains(menuID) || "U9".Equals(menuID) || "X1".Equals(WinSettings.TTPeriodInfo.menuID))
			{
				if ((bool)winFpSpread.Sheets[0].Cells[1, 0].Value && (bool)winFpSpread.Sheets[0].Cells[2, 0].Value)
				{
					frmMessage frmMessage = new frmMessage("0395", "3", "3", "Thông báo");
					((Form)frmMessage).ShowDialog();
					result = false;
					return result;
				}
				if (!(bool)winFpSpread.Sheets[0].Cells[0, 0].Value)
				{
					frmMessage frmMessage = new frmMessage("0475", "3", "3", "Thông báo");
					((Form)frmMessage).ShowDialog();
					result = false;
					return result;
				}
			}
			if ("A2,A5".Contains(menuID))
			{
				if ((bool)winFpSpread.Sheets[0].Cells[1, 0].Value && (bool)winFpSpread.Sheets[0].Cells[2, 0].Value)
				{
					frmMessage frmMessage = new frmMessage("0395", "3", "3", "Thông báo");
					((Form)frmMessage).ShowDialog();
					result = false;
					return result;
				}
				if (!(bool)winFpSpread.Sheets[0].Cells[0, 0].Value)
				{
					frmMessage frmMessage = new frmMessage("0485", "3", "3", "Thông báo");
					((Form)frmMessage).ShowDialog();
					result = false;
					return result;
				}
				if (!(bool)winFpSpread.Sheets[0].Cells[3, 0].Value)
				{
					frmMessage frmMessage = new frmMessage("0485", "3", "3", "Thông báo");
					((Form)frmMessage).ShowDialog();
					result = false;
					return result;
				}
			}
			if ("P9,Q1,Q2,Q3".Contains(menuID) && (bool)winFpSpread.Sheets[0].Cells[1, 0].Value && (bool)winFpSpread.Sheets[0].Cells[2, 0].Value)
			{
				frmMessage frmMessage = new frmMessage("0395", "3", "3", "Thông báo");
				((Form)frmMessage).ShowDialog();
				result = false;
				return result;
			}
			if ("B3".Contains(menuID))
			{
				if ((bool)winFpSpread.Sheets[0].Cells[1, 0].Value && (bool)winFpSpread.Sheets[0].Cells[2, 0].Value)
				{
					frmMessage frmMessage = new frmMessage("0395", "3", "3", "Thông báo");
					((Form)frmMessage).ShowDialog();
					result = false;
					return result;
				}
				if (!(bool)winFpSpread.Sheets[0].Cells[0, 0].Value)
				{
					frmMessage frmMessage = new frmMessage("0491", "3", "3", "Thông báo");
					((Form)frmMessage).ShowDialog();
					result = false;
					return result;
				}
				if (!(bool)winFpSpread.Sheets[0].Cells[1, 0].Value && !(bool)winFpSpread.Sheets[0].Cells[2, 0].Value)
				{
					frmMessage frmMessage = new frmMessage("0491", "3", "3", "Thông báo");
					((Form)frmMessage).ShowDialog();
					result = false;
					return result;
				}
				if (!(bool)winFpSpread.Sheets[0].Cells[3, 0].Value)
				{
					frmMessage frmMessage = new frmMessage("0491", "3", "3", "Thông báo");
					((Form)frmMessage).ShowDialog();
					result = false;
					return result;
				}
			}
			if ("B4".Contains(menuID))
			{
				if ((bool)winFpSpread.Sheets[0].Cells[1, 0].Value && (bool)winFpSpread.Sheets[0].Cells[2, 0].Value)
				{
					frmMessage frmMessage = new frmMessage("0395", "3", "3", "Thông báo");
					((Form)frmMessage).ShowDialog();
					result = false;
					return result;
				}
				if (!(bool)winFpSpread.Sheets[0].Cells[0, 0].Value)
				{
					frmMessage frmMessage = new frmMessage("0492", "3", "3", "Thông báo");
					((Form)frmMessage).ShowDialog();
					result = false;
					return result;
				}
				if (!(bool)winFpSpread.Sheets[0].Cells[1, 0].Value && !(bool)winFpSpread.Sheets[0].Cells[2, 0].Value)
				{
					frmMessage frmMessage = new frmMessage("0492", "3", "3", "Thông báo");
					((Form)frmMessage).ShowDialog();
					result = false;
					return result;
				}
			}
			if ("B5".Contains(menuID))
			{
				if ((bool)winFpSpread.Sheets[0].Cells[1, 0].Value && (bool)winFpSpread.Sheets[0].Cells[2, 0].Value)
				{
					frmMessage frmMessage = new frmMessage("0395", "3", "3", "Thông báo");
					((Form)frmMessage).ShowDialog();
					result = false;
					return result;
				}
				if (!(bool)winFpSpread.Sheets[0].Cells[0, 0].Value)
				{
					frmMessage frmMessage = new frmMessage("0493", "3", "3", "Thông báo");
					((Form)frmMessage).ShowDialog();
					result = false;
					return result;
				}
				if (!(bool)winFpSpread.Sheets[0].Cells[1, 0].Value && !(bool)winFpSpread.Sheets[0].Cells[2, 0].Value)
				{
					frmMessage frmMessage = new frmMessage("0493", "3", "3", "Thông báo");
					((Form)frmMessage).ShowDialog();
					result = false;
					return result;
				}
				if (!(bool)winFpSpread.Sheets[0].Cells[3, 0].Value)
				{
					frmMessage frmMessage = new frmMessage("0493", "3", "3", "Thông báo");
					((Form)frmMessage).ShowDialog();
					result = false;
					return result;
				}
			}
			if ("A6".Contains(menuID))
			{
				if ((bool)winFpSpread.Sheets[0].Cells[1, 0].Value && (bool)winFpSpread.Sheets[0].Cells[2, 0].Value)
				{
					frmMessage frmMessage = new frmMessage("0395", "3", "3", "Thông báo");
					((Form)frmMessage).ShowDialog();
					result = false;
					return result;
				}
				if (!(bool)winFpSpread.Sheets[0].Cells[0, 0].Value)
				{
					frmMessage frmMessage = new frmMessage("0486", "3", "3", "Thông báo");
					((Form)frmMessage).ShowDialog();
					result = false;
					return result;
				}
				if (!(bool)winFpSpread.Sheets[0].Cells[1, 0].Value && !(bool)winFpSpread.Sheets[0].Cells[2, 0].Value)
				{
					frmMessage frmMessage = new frmMessage("0478", "3", "3", "Thông báo");
					((Form)frmMessage).ShowDialog();
					result = false;
					return result;
				}
			}
			if ("A7".Contains(menuID) && !(bool)winFpSpread.Sheets[0].Cells[0, 0].Value)
			{
				frmMessage frmMessage = new frmMessage("0484", "3", "3", "Thông báo");
				((Form)frmMessage).ShowDialog();
				result = false;
				return result;
			}
			if ("A3,H3,A9,H9".Contains(menuID))
			{
				if (!(bool)winFpSpread.Sheets[0].Cells[0, 0].Value)
				{
					frmMessage frmMessage = new frmMessage("0475", "3", "3", "Thông báo");
					((Form)frmMessage).ShowDialog();
					result = false;
					return result;
				}
				if (!(bool)winFpSpread.Sheets[0].Cells[1, 0].Value)
				{
					frmMessage frmMessage = new frmMessage("0479", "3", "3", "Thông báo");
					((Form)frmMessage).ShowDialog();
					result = false;
					return result;
				}
			}
			if ("C1".Contains(menuID))
			{
				if ((bool)winFpSpread.Sheets[0].Cells[1, 0].Value && (bool)winFpSpread.Sheets[0].Cells[2, 0].Value)
				{
					frmMessage frmMessage = new frmMessage("0395", "3", "3", "Thông báo");
					((Form)frmMessage).ShowDialog();
					result = false;
					return result;
				}
				if (!(bool)winFpSpread.Sheets[0].Cells[0, 0].Value)
				{
					frmMessage frmMessage = new frmMessage("0498", "3", "3", "Thông báo");
					((Form)frmMessage).ShowDialog();
					result = false;
					return result;
				}
				if (!(bool)winFpSpread.Sheets[0].Cells[1, 0].Value && !(bool)winFpSpread.Sheets[0].Cells[2, 0].Value)
				{
					frmMessage frmMessage = new frmMessage("0498", "3", "3", "Thông báo");
					((Form)frmMessage).ShowDialog();
					result = false;
					return result;
				}
			}
			if ("C2".Contains(menuID))
			{
				if (!(bool)winFpSpread.Sheets[0].Cells[0, 0].Value)
				{
					frmMessage frmMessage = new frmMessage("0498", "3", "3", "Thông báo");
					((Form)frmMessage).ShowDialog();
					result = false;
					return result;
				}
				if (!(bool)winFpSpread.Sheets[0].Cells[1, 0].Value)
				{
					frmMessage frmMessage = new frmMessage("0498", "3", "3", "Thông báo");
					((Form)frmMessage).ShowDialog();
					result = false;
					return result;
				}
			}
			if ("C3,C4,C5".Contains(menuID))
			{
				if (!(bool)winFpSpread.Sheets[0].Cells[0, 0].Value)
				{
					frmMessage frmMessage = new frmMessage("0499", "3", "3", "Thông báo");
					((Form)frmMessage).ShowDialog();
					result = false;
					return result;
				}
				if (!(bool)winFpSpread.Sheets[0].Cells[1, 0].Value)
				{
					frmMessage frmMessage = new frmMessage("0499", "3", "3", "Thông báo");
					((Form)frmMessage).ShowDialog();
					result = false;
					return result;
				}
			}
			if (menuID.Equals("03"))
			{
				if (ConvertHelpers.ConvertToInt(nam) <= 2015)
				{
					if ((bool)winFpSpread.Sheets[0].Cells[14, 0].Value)
					{
						frmMessage frmMessage = new frmMessage("0415", "3", "3", "Thông báo");
						((Form)frmMessage).ShowDialog();
						result = false;
						return result;
					}
				}
				else if ((bool)winFpSpread.Sheets[0].Cells[10, 0].Value)
				{
					frmMessage frmMessage = new frmMessage("0414", "3", "3", "Thông báo");
					((Form)frmMessage).ShowDialog();
					result = false;
					return result;
				}
				if ((ConvertHelpers.ConvertToInt(nam) < 2017 || ConvertHelpers.ConvertToInt(nam) > 2019) && ((bool)winFpSpread.Sheets[0].Cells[18, 0].Value || (bool)winFpSpread.Sheets[0].Cells[15, 0].Value || (bool)winFpSpread.Sheets[0].Cells[16, 0].Value || (bool)winFpSpread.Sheets[0].Cells[17, 0].Value))
				{
					frmMessage frmMessage = new frmMessage("0482", "3", "", "Thông báo");
					((Form)frmMessage).ShowDialog();
					result = false;
					return result;
				}
				if (ConvertHelpers.ConvertToInt(nam) == 2019)
				{
					bool flag = false;
					bool flag2 = false;
					if ((bool)winFpSpread.Sheets[0].Cells[18, 0].Value || (bool)winFpSpread.Sheets[0].Cells[15, 0].Value || (bool)winFpSpread.Sheets[0].Cells[16, 0].Value || (bool)winFpSpread.Sheets[0].Cells[17, 0].Value)
					{
						flag = true;
					}
					if ((bool)winFpSpread.Sheets[0].Cells[19, 0].Value || (bool)winFpSpread.Sheets[0].Cells[20, 0].Value || (bool)winFpSpread.Sheets[0].Cells[21, 0].Value || (bool)winFpSpread.Sheets[0].Cells[22, 0].Value)
					{
						flag2 = true;
					}
					if (flag && flag2)
					{
						frmMessage frmMessage = new frmMessage("0525", "3", "", "Thông báo");
						((Form)frmMessage).ShowDialog();
						result = false;
						return result;
					}
				}
				if (ConvertHelpers.ConvertToInt(nam) >= 2017 && (bool)winFpSpread.Sheets[0].Cells[11, 0].Value)
				{
					frmMessage frmMessage = new frmMessage("0483", "3", "", "Thông báo");
					((Form)frmMessage).ShowDialog();
					result = false;
					return result;
				}
				if ((ConvertHelpers.ConvertToInt(nam) < 2020 || ConvertHelpers.ConvertToInt(nam) > 2020) && (bool)winFpSpread.Sheets[0].Cells[23, 0].Value)
				{
					frmMessage frmMessage = new frmMessage("0503", "3", "", "Thông báo");
					((Form)frmMessage).ShowDialog();
					result = false;
					return result;
				}
				if (ConvertHelpers.ConvertToInt(nam) < 2019 && (bool)winFpSpread.Sheets[0].Cells[24, 0].Value)
				{
					frmMessage frmMessage = new frmMessage("0528", "3", "", "Thông báo");
					((Form)frmMessage).ShowDialog();
					result = false;
					return result;
				}
			}
			if (menuID == "73" || menuID == "55")
			{
				if ((ConvertHelpers.ConvertToInt(nam) < 2020 || ConvertHelpers.ConvertToInt(nam) > 2020) && (bool)winFpSpread.Sheets[0].Cells[0, 0].Value)
				{
					frmMessage frmMessage = new frmMessage("0503", "3", "", "Thông báo");
					((Form)frmMessage).ShowDialog();
					result = false;
					return result;
				}
				if (ConvertHelpers.ConvertToInt(nam) < 2021 || ConvertHelpers.ConvertToInt(nam) > 2021)
				{
					if ((bool)winFpSpread.Sheets[0].Cells[1, 0].Value)
					{
						frmMessage frmMessage = new frmMessage("0541", "3", "", "Thông báo");
						((Form)frmMessage).ShowDialog();
						result = false;
						return result;
					}
				}
				else if (menuID == "55" && periodInfo.KieuKy == "Y" && (bool)winFpSpread.Sheets[0].Cells[1, 0].Value)
				{
					frmMessage frmMessage = new frmMessage("0542", "3", "", "Thông báo");
					((Form)frmMessage).ShowDialog();
					result = false;
					return result;
				}
			}
			if (menuID == "T5")
			{
				if (ConvertHelpers.ConvertToInt(nam) < 2021 || ConvertHelpers.ConvertToInt(nam) > 2021)
				{
					if ((bool)winFpSpread.Sheets[0].Cells[0, 0].Value)
					{
						frmMessage frmMessage = new frmMessage("0541", "3", "", "Thông báo");
						((Form)frmMessage).ShowDialog();
						result = false;
						return result;
					}
				}
				else if (periodInfo.KieuKy == "D" && (bool)winFpSpread.Sheets[0].Cells[0, 0].Value)
				{
					frmMessage frmMessage = new frmMessage("0616", "3", "", "Thông báo");
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

	public static void CellFormatImport(FpSpread winFpSpread, int sheetIndex, long iRow, long iCol)
	{
		try
		{
			MethodInfo method = GetMethod("CellFormatImport");
			if ((object)method == null)
			{
				return;
			}
			ParameterInfo[] parameters = method.GetParameters();
			object[] array = new object[parameters.Length];
			for (int i = 0; i < parameters.Length; i++)
			{
				switch (parameters[i].Name)
				{
				case "winFpSpread":
					array[i] = winFpSpread;
					break;
				case "ToKhaiInfo":
					array[i] = WinSettings.ToKhaiInfo;
					break;
				case "TTinChung":
					array[i] = WinSettings.TTinChung;
					break;
				case "PeriodInfo":
					array[i] = WinSettings.TTPeriodInfo;
					break;
				case "sheetIndex":
					array[i] = sheetIndex;
					break;
				case "iRow":
					array[i] = iRow;
					break;
				case "iCol":
					array[i] = iCol;
					break;
				}
			}
			method.Invoke(WinSettings.InstanceClass, array);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void FormatReport_01_1_KHBS(SheetView sview_Report, int iRow_Report, int iCol_Report, ref string controlType_Report, ref string value, ref string defaultValue_Report, int sheetIndex_KHBS = 0, HTKKCMCTableInfo tableSInfo = null)
	{
		try
		{
			HTKKCMCSheetInfo hTKKCMCSheetInfo = WinSettings.ToKhaiInfo.SheetInfos[sheetIndex_KHBS];
			if (hTKKCMCSheetInfo == null || ConvertHelpers._dicCellType_01_1_KHBS == null)
			{
				return;
			}
			XmlNode mCT = WinFnProcess.GetMCT01("MapMCT.xml", WinSettings.ToKhaiInfo.MenuID);
			if (!"I7,I4,C7,I8,C8".Contains(WinSettings.MenuID) && mCT.Attributes["LoaiTienPath"] == null && !(hTKKCMCSheetInfo.SheetName == "01-1/KHBS") && !(hTKKCMCSheetInfo.SheetName == "01-KHBS"))
			{
				return;
			}
			ICellType cellType = null;
			if (tableSInfo != null)
			{
				int dataLocationTemplateRowStart = tableSInfo.Location.DataLocationTemplateRowStart;
				int dataLocationReportRowStart = tableSInfo.Location.DataLocationReportRowStart;
				int num = dataLocationReportRowStart - dataLocationTemplateRowStart;
				int num2 = iRow_Report - num;
				HTKKCMCTableCellInfo hTKKCMCTableCellInfo = tableSInfo.RowsInfo.Cells.FirstOrDefault((HTKKCMCTableCellInfo x) => x.CellReport.ColIndex == iCol_Report);
				int colIndex = hTKKCMCTableCellInfo.CellTemplate.ColIndex;
				List<string> list = new List<string>();
				list.Add(sheetIndex_KHBS.ToString());
				list.Add(num2.ToString());
				list.Add(colIndex.ToString());
				string key = Functions.FormatKey(list, "___");
				if (ConvertHelpers._dicCellType_01_1_KHBS.ContainsKey(key))
				{
					CellTypeKHBS cellTypeKHBS = ConvertHelpers._dicCellType_01_1_KHBS[key];
					cellType = cellTypeKHBS.iCellType;
				}
			}
			else
			{
				HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCSheetInfo.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.CellReport.RowIndex == iRow_Report && x.CellReport.ColIndex == iCol_Report);
				if (hTKKCMCCellInfo == null)
				{
					return;
				}
				int num2 = hTKKCMCCellInfo.CellTemplate.RowIndex;
				int colIndex = hTKKCMCCellInfo.CellTemplate.ColIndex;
				List<string> list = new List<string>();
				list.Add(sheetIndex_KHBS.ToString());
				list.Add(num2.ToString());
				list.Add(colIndex.ToString());
				string key = Functions.FormatKey(list, "___");
				if (ConvertHelpers._dicCellType_01_1_KHBS.ContainsKey(key))
				{
					CellTypeKHBS cellTypeKHBS = ConvertHelpers._dicCellType_01_1_KHBS[key];
					cellType = cellTypeKHBS.iCellType;
				}
			}
			if (cellType == null)
			{
				return;
			}
			sview_Report.Cells[iRow_Report - 1, iCol_Report - 1].CellType = cellType;
			NumberCellType numberCellType = new NumberCellType();
			CurrencyCellType currencyCellType = new CurrencyCellType();
			TextCellType textCellType = new TextCellType();
			CheckBoxCellType checkBoxCellType = new CheckBoxCellType();
			MaskCellType maskCellType = new MaskCellType();
			if ((object)cellType.GetType() == numberCellType.GetType())
			{
				int decimalPlaces = ((NumberCellType)cellType).DecimalPlaces;
				defaultValue_Report = "0,";
				for (int num3 = 0; num3 < decimalPlaces; num3++)
				{
					defaultValue_Report += "0";
				}
				if (defaultValue_Report.Length == 2)
				{
					defaultValue_Report = "0";
				}
				controlType_Report = "16";
				int digits = ((defaultValue_Report.Length > 2) ? (defaultValue_Report.Length - 2) : 0);
				if (!string.IsNullOrEmpty(value))
				{
					value = Math.Round(Convert.ToDouble(value), digits).ToString();
				}
			}
			else if ((object)cellType.GetType() == currencyCellType.GetType())
			{
				int decimalPlaces = ((CurrencyCellType)cellType).DecimalPlaces;
				defaultValue_Report = "0,";
				for (int num3 = 0; num3 < decimalPlaces; num3++)
				{
					defaultValue_Report += "0";
				}
				if (defaultValue_Report.Length == 2)
				{
					defaultValue_Report = "0";
				}
				controlType_Report = "16";
				int digits = ((defaultValue_Report.Length > 2) ? (defaultValue_Report.Length - 2) : 0);
				if (!string.IsNullOrEmpty(value))
				{
					value = Math.Round(Convert.ToDouble(value), digits).ToString();
				}
			}
			else if ((object)cellType.GetType() == textCellType.GetType())
			{
				controlType_Report = "0";
			}
			else if ((object)cellType.GetType() == maskCellType.GetType())
			{
				controlType_Report = "14";
			}
			else if ((object)cellType.GetType() == checkBoxCellType.GetType())
			{
				controlType_Report = "2";
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void SetCellExport(SheetView sview, int iRow, int iCol, string cellType, string value, string defaultValue, int sheetIndex = 0, HTKKCMCTableInfo tableSInfo = null)
	{
		try
		{
			FormatReport_01_1_KHBS(sview, iRow, iCol, ref cellType, ref value, ref defaultValue, sheetIndex, tableSInfo);
			string text = value;
			int num = 0;
			if (iRow == 48 && iCol == HTKKFile.ColumnLettersToIndex("W"))
			{
				num = 1;
			}
			TextCellType textCellType = new TextCellType();
			textCellType.WordWrap = true;
			textCellType.MaxLength = 500;
			sview.Cells[iRow - 1, iCol - 1].Formula = "";
			sview.Cells[iRow - 1, iCol - 1].CellType = textCellType;
			switch (cellType)
			{
			case "2":
				text = ((!(text.ToUpper() == "TRUE") && !(text == "1") && !(text.ToUpper() == "X")) ? "" : "x");
				break;
			case "14":
				text = HTKKFormater.toDateDDMMYYYY(text, isMess: false);
				break;
			case "96":
				text = HTKKFormater.toTimeHHMMSS(text, isMess: false);
				break;
			case "94":
				text = HTKKFormater.toDateAndTimeDDMMYYYYHHMMSS(text, isMess: false);
				break;
			case "19":
				text = HTKKFormater.toDateMMYYYY(text, isMess: false);
				break;
			case "16":
			{
				if (ConvertHelpers.ConvertToDouble(text) == 0.0 && WinSettings.isActionPrint)
				{
					sview.Cells[iRow - 1, iCol - 1].Text = "";
					return;
				}
				int num2 = 0;
				string text2 = "";
				if ((WinSettings.MenuID == "78" || "U9".Equals(WinSettings.MenuID)) && WinSettings.TTPeriodInfo.DonViTinhMa == "NT" && iRow != 38)
				{
					CultureInfo provider = CultureInfo.CreateSpecificCulture("da-DK");
					text2 = "{0:0,0.00}";
					text = string.Format(provider, text2, new object[1] { ConvertHelpers.ConvertToDouble(text) });
					if (text.IndexOf("-") == 0)
					{
						text = "(" + text.Substring(1) + ")";
					}
				}
				else if (HTKKFormater.isNullString(defaultValue) || defaultValue == "0")
				{
					CultureInfo provider = CultureInfo.CreateSpecificCulture("el-GR");
					if (ConvertHelpers.ConvertToDouble1(text) == 0.0)
					{
						sview.Cells[iRow - 1, iCol - 1].Text = "0";
						return;
					}
					text = Math.Round(ConvertHelpers.ConvertToDouble(text), 0).ToString();
					text = string.Format(provider, "{0:0,0}", new object[1] { ConvertHelpers.ConvertToDouble1(text) });
					if (text.Length <= 2 && text.IndexOf("0") == 0)
					{
						text = text.Substring(1);
					}
					if (text.Length <= 3 && (text.IndexOf("-0") == 0 || text.IndexOf("(0") == 0))
					{
						text = "(" + text.Substring(2) + ")";
					}
					if (text.IndexOf("-") == 0)
					{
						text = "(" + text.Substring(1) + ")";
					}
				}
				else if (defaultValue.IndexOf(".") > -1)
				{
					if (ConvertHelpers.ConvertToDouble1(text) == 0.0)
					{
						sview.Cells[iRow - 1, iCol - 1].Text = defaultValue;
						return;
					}
					num2 = defaultValue.Substring(defaultValue.LastIndexOf(".") + 1).Length;
					CultureInfo provider = CultureInfo.CreateSpecificCulture("da-DK");
					text2 = "{{0:0,0." + HTKKFile.GetStrByLengh("0", num2) + "}";
					text = string.Format(provider, text2, new object[1] { ConvertHelpers.ConvertToDouble(text) });
					if (text.IndexOf("-") == 0)
					{
						text = "(" + text.Substring(1) + ")";
					}
				}
				else if (defaultValue.IndexOf(",") > -1)
				{
					if (ConvertHelpers.ConvertToDouble1(text) == 0.0)
					{
						sview.Cells[iRow - 1, iCol - 1].Text = defaultValue;
						return;
					}
					num2 = defaultValue.Substring(defaultValue.LastIndexOf(",") + 1).Length;
					CultureInfo provider = CultureInfo.CreateSpecificCulture("da-DK");
					text2 = "{0:0,0." + HTKKFile.GetStrByLengh("0", num2) + "}";
					text = string.Format(provider, text2, new object[1] { ConvertHelpers.ConvertToDouble(text) });
					text = ((ConvertHelpers.ConvertToDouble(text.Split(new char[1] { ',' })[0].Replace(".", "")) != 0.0) ? text.TrimStart(new char[1] { '0' }).Replace("," + defaultValue.Split(new char[1] { ',' })[1], "") : ("0," + text.Split(new char[1] { ',' })[1]));
					if (text.IndexOf("-") == 0)
					{
						text = "(" + text.Substring(1) + ")";
					}
					sview.Cells[iRow - 1, iCol - 1].Text = text;
				}
				break;
			}
			}
			if (!string.IsNullOrEmpty(text) && text.IndexOf(",") > 0 && WinSettings.MenuID == "38" && ConvertHelpers.ConvertToInt(text.Split(new char[1] { ',' })[1].Replace(")", "")) == 0)
			{
				text = text.Split(new char[1] { ',' })[0];
				if (text.StartsWith("("))
				{
					text += ")";
				}
			}
			sview.Cells[iRow - 1, iCol - 1].Text = text;
			Row row = sview.Rows[iRow - 1];
			float preferredHeight = row.GetPreferredHeight();
			if (WinSettings.MenuID == "68")
			{
				if (text.Length > 30)
				{
					float num3 = preferredHeight / row.Height;
					preferredHeight += num3 * 5f;
					row.Height = preferredHeight;
				}
			}
			else
			{
				if (WinSettings.MenuID == "33" && (iRow == 61 || iRow == 62))
				{
					return;
				}
				if (WinSettings.MenuID == "N7")
				{
					if (row.Height < preferredHeight)
					{
						float num3 = preferredHeight / row.Height;
						preferredHeight += num3 * 8f;
						row.Height = preferredHeight;
					}
				}
				else
				{
					row.Height = preferredHeight;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void SetCellExportExcel(SheetView sview, int iRow, int iCol, string cellType, string value, string defaultValue)
	{
		try
		{
			string text = value;
			switch (cellType)
			{
			case "2":
				text = ((!(text.ToUpper() == "TRUE") && !(text == "1") && !(text.ToUpper() == "X")) ? "" : "x");
				sview.Cells[iRow - 1, iCol - 1].CellType = new TextCellType();
				sview.Cells[iRow - 1, iCol - 1].Text = text;
				break;
			case "14":
				text = HTKKFormater.toDateDDMMYYYY(text, isMess: false);
				sview.Cells[iRow - 1, iCol - 1].CellType = new TextCellType();
				sview.Cells[iRow - 1, iCol - 1].Text = text;
				break;
			case "96":
				text = HTKKFormater.toTimeHHMMSS(text, isMess: false);
				sview.Cells[iRow - 1, iCol - 1].CellType = new TextCellType();
				sview.Cells[iRow - 1, iCol - 1].Text = text;
				break;
			case "94":
				text = HTKKFormater.toDateAndTimeDDMMYYYYHHMMSS(text, isMess: false);
				sview.Cells[iRow - 1, iCol - 1].CellType = new TextCellType();
				sview.Cells[iRow - 1, iCol - 1].Text = text;
				break;
			case "19":
				text = HTKKFormater.toDateMMYYYY(text, isMess: false);
				sview.Cells[iRow - 1, iCol - 1].CellType = new TextCellType();
				sview.Cells[iRow - 1, iCol - 1].Text = text;
				break;
			case "15":
				sview.Cells[iRow - 1, iCol - 1].CellType = new TextCellType();
				sview.Cells[iRow - 1, iCol - 1].Value = "";
				break;
			}
			Row row = sview.Rows[iRow - 1];
			float preferredHeight = row.GetPreferredHeight();
			if (row.Height < preferredHeight - 5f || row.Height > preferredHeight + 5f)
			{
				row.Height = preferredHeight;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void SetCellAfterExportExcel(SheetView sview, int iRow, int iCol, string cellType, string value, string defaultValue)
	{
		try
		{
			switch (cellType)
			{
			case "2":
				sview.Cells[iRow - 1, iCol - 1].CellType = new CheckBoxCellType();
				sview.Cells[iRow - 1, iCol - 1].Text = value;
				break;
			case "14":
			{
				MaskCellType maskCellType5 = new MaskCellType();
				maskCellType5.Mask = "00/00/0000";
				maskCellType5.MaskChar = '.';
				sview.Cells[iRow - 1, iCol - 1].CellType = maskCellType5;
				sview.Cells[iRow - 1, iCol - 1].Text = value;
				break;
			}
			case "96":
			{
				MaskCellType maskCellType4 = new MaskCellType();
				maskCellType4.Mask = "00:00:00";
				maskCellType4.MaskChar = '.';
				sview.Cells[iRow - 1, iCol - 1].CellType = maskCellType4;
				sview.Cells[iRow - 1, iCol - 1].Text = value;
				break;
			}
			case "94":
			{
				MaskCellType maskCellType3 = new MaskCellType();
				maskCellType3.Mask = "00/00/0000 00:00:00";
				maskCellType3.MaskChar = '.';
				sview.Cells[iRow - 1, iCol - 1].CellType = maskCellType3;
				sview.Cells[iRow - 1, iCol - 1].Text = value;
				break;
			}
			case "19":
			{
				MaskCellType maskCellType2 = new MaskCellType();
				maskCellType2.Mask = "00/0000";
				maskCellType2.MaskChar = '.';
				sview.Cells[iRow - 1, iCol - 1].CellType = maskCellType2;
				sview.Cells[iRow - 1, iCol - 1].Text = value;
				break;
			}
			case "109":
			{
				MaskCellType maskCellType = new MaskCellType();
				maskCellType.Mask = "0000";
				maskCellType.MaskChar = '.';
				sview.Cells[iRow - 1, iCol - 1].CellType = maskCellType;
				sview.Cells[iRow - 1, iCol - 1].Text = value;
				break;
			}
			case "15":
			{
				ButtonCellType buttonCellType = new ButtonCellType();
				buttonCellType.TwoState = false;
				buttonCellType.Text = defaultValue;
				buttonCellType.ShadowSize = 0;
				buttonCellType.TextColor = Color.Blue;
				sview.Cells[iRow - 1, iCol - 1].CellType = buttonCellType;
				break;
			}
			}
			Row row = sview.Rows[iRow - 1];
			float preferredHeight = row.GetPreferredHeight();
			if (row.Height < preferredHeight - 5f || row.Height > preferredHeight + 5f)
			{
				row.Height = preferredHeight;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void SetFPSpreadExp(FpSpread winFpExp, List<HTKKCMCSheetInfo> sheetInfos)
	{
		try
		{
			MethodInfo method = GetMethod("SetFPSpreadExp");
			if ((object)method != null)
			{
				method.Invoke(parameters: new object[6]
				{
					winFpExp,
					sheetInfos,
					WinSettings.TTPeriodInfo,
					WinSettings.ToKhaiInfo,
					WinSettings.TTinChung,
					WinSettings.HTKKHeaderInfo
				}, obj: WinSettings.InstanceClass);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static string GetFileNameBySheetNew(HTKKPeriodInfo periodInfor, string dataFile, string sFolder, string versionXML, bool isLoadSavedFile = true)
	{
		try
		{
			string text = "";
			string text2 = "";
			string text3 = "~98~92~08~40~42~44~M8~M9~K2~K3~J9~J8~J7~E2~E3~E4~E5~E6~E7~E9~D7~D8~D9~R9~R6~S1~";
			string text4 = "~J3~Q7~Q5~Q6~";
			versionXML = ((!(periodInfor.LoaiToKhai == "B") || !dataFile.StartsWith("KHBS")) ? versionXML.Replace(".", "") : WinSettings.ToKhaiInfo.VersionKBSXML.ToString().Replace(".", ""));
			if ("66;67;09;10;60;61;G2;".IndexOf(WinSettings.MenuID) > -1)
			{
				return sFolder + dataFile + "_" + versionXML + ".xml";
			}
			if (text3.IndexOf("~" + periodInfor.menuID + "~") > 0)
			{
				if (periodInfor.TieuMucHachToanVal.Equals("1"))
				{
					text2 = "DT";
				}
				else if (periodInfor.TieuMucHachToanVal.Equals("2"))
				{
					text2 = "KTN";
				}
				else if (periodInfor.TieuMucHachToanVal.Equals("3"))
				{
					text2 = "CD";
				}
				else if (periodInfor.TieuMucHachToanVal.Equals("4"))
				{
					text2 = "KDDT";
				}
				else if (periodInfor.TieuMucHachToanVal.Equals("5"))
				{
					text2 = "KDKTN";
				}
				else if (periodInfor.TieuMucHachToanVal.Equals("6"))
				{
					text2 = "KDCD";
				}
			}
			if (periodInfor.KieuKy == "M")
			{
				text = ((text3.IndexOf("~" + periodInfor.menuID + "~") > 0) ? (sFolder + dataFile + ((!string.IsNullOrEmpty(periodInfor.DonViTinhMa)) ? ("_" + periodInfor.DonViTinhMa) : "") + "_" + text2 + "_" + periodInfor.Thang + periodInfor.Nam + "_" + versionXML + ".xml") : ((periodInfor.menuID == "A4") ? ((!periodInfor.MaHSDuAnDT.Equals("")) ? (sFolder + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.MaHSDuAnDT + "_" + periodInfor.TuThang.Replace("/", "") + "_" + periodInfor.DenThang.Replace("/", "") + "_" + versionXML + ".xml") : (sFolder + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.TuThang.Replace("/", "") + "_" + periodInfor.DenThang.Replace("/", "") + "_" + versionXML + ".xml")) : ((!(periodInfor.menuID == "99") && !(periodInfor.menuID == "H6")) ? ((!(periodInfor.menuID == "01") && !(periodInfor.menuID == "I5")) ? ((!(periodInfor.menuID == "D1") && !WinSettings.MenuID.Equals("Z5")) ? ((!(periodInfor.menuID == "02") && !(periodInfor.menuID == "D2")) ? ((!(periodInfor.menuID == "71") && !(periodInfor.menuID == "D4") && !WinSettings.MenuID.Equals("Z3")) ? ((!(periodInfor.menuID == "M4") && !(periodInfor.menuID == "G3")) ? ((!"U8~".Contains(periodInfor.menuID)) ? (sFolder + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.Thang + periodInfor.Nam + "_" + versionXML + ".xml") : (sFolder + dataFile + "_L" + periodInfor.LanXuatBan + "_" + periodInfor.Thang + periodInfor.Nam + "_" + versionXML + ".xml")) : (sFolder + dataFile + "_L" + periodInfor.LanXuatBan + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.Thang + periodInfor.Nam + "_" + versionXML + ".xml")) : ((ConvertHelpers.ConvertToInt(periodInfor.Nam) < 2021 || periodInfor.ThuHo.Equals("")) ? (sFolder + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.Thang + periodInfor.Nam + "_" + versionXML + ".xml") : (sFolder + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.ThuHo + "_" + periodInfor.Thang + periodInfor.Nam + "_" + versionXML + ".xml"))) : ((ConvertHelpers.ConvertToInt(periodInfor.Nam) < 2021) ? (sFolder + dataFile + "_" + periodInfor.Thang + periodInfor.Nam + "_" + versionXML + ".xml") : (sFolder + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.MaHSDuAnDT + "_" + periodInfor.Thang + periodInfor.Nam + "_" + versionXML + ".xml"))) : (sFolder + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.TieuMucHachToanVal + "_" + periodInfor.Thang + periodInfor.Nam + "_" + versionXML + ".xml")) : ((ConvertHelpers.ConvertToInt(periodInfor.Nam) < 2021) ? (sFolder + dataFile + "_" + periodInfor.Thang + periodInfor.Nam + "_" + versionXML + ".xml") : (sFolder + dataFile + "_" + periodInfor.TieuMucHachToanVal + "_" + periodInfor.Thang + periodInfor.Nam + "_" + versionXML + ".xml"))) : ((!periodInfor.MaHSDuAnDT.Equals("")) ? (sFolder + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.MaHSDuAnDT + ((!string.IsNullOrEmpty(GetNameTheoKieuKyHoanThue(periodInfor.KieuKyHoanThue))) ? ("_" + GetNameTheoKieuKyHoanThue(periodInfor.KieuKyHoanThue)) : "") + periodInfor.TuThang.Replace("/", "") + "_" + periodInfor.DenThang.Replace("/", "") + "_" + versionXML + ".xml") : (sFolder + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + ((!string.IsNullOrEmpty(GetNameTheoKieuKyHoanThue(periodInfor.KieuKyHoanThue))) ? ("_" + GetNameTheoKieuKyHoanThue(periodInfor.KieuKyHoanThue)) : "") + periodInfor.TuThang.Replace("/", "").PadLeft(6, '0') + "_" + periodInfor.DenThang.Replace("/", "").PadLeft(6, '0') + "_" + versionXML + ".xml")))));
			}
			else if (periodInfor.KieuKy == "Q")
			{
				text = ((!HTKKFormater.isNullString(periodInfor.TuThang) && !HTKKFormater.isNullString(periodInfor.DenThang)) ? ((!periodInfor.menuID.Equals("42") && text3.IndexOf("~" + periodInfor.menuID + "~") <= 0) ? (periodInfor.menuID.Equals("A4") ? ((!periodInfor.MaHSDuAnDT.Equals("")) ? (sFolder + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.MaHSDuAnDT + "_Q0" + periodInfor.TuThang.Replace("/", "") + "_0" + periodInfor.DenThang.Replace("/", "") + "_" + versionXML + ".xml") : (sFolder + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_Q0" + periodInfor.TuThang.Replace("/", "") + "_0" + periodInfor.DenThang.Replace("/", "") + "_" + versionXML + ".xml")) : ((!periodInfor.menuID.Equals("99") && !periodInfor.menuID.Equals("H6")) ? (sFolder + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_Q" + $"{periodInfor.Quy}{periodInfor.Nam}".PadLeft(6, '0') + "_" + periodInfor.TuThang.Replace("/", "") + "_" + periodInfor.DenThang.Replace("/", "") + "_" + versionXML + ".xml") : ((!periodInfor.MaHSDuAnDT.Equals("")) ? (sFolder + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.MaHSDuAnDT + ((!string.IsNullOrEmpty(GetNameTheoKieuKyHoanThue(periodInfor.KieuKyHoanThue))) ? ("_" + GetNameTheoKieuKyHoanThue(periodInfor.KieuKyHoanThue)) : "") + periodInfor.TuThang.Replace("/", "").PadLeft(6, '0') + "_" + periodInfor.DenThang.Replace("/", "").PadLeft(6, '0') + "_" + versionXML + ".xml") : (sFolder + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + ((!string.IsNullOrEmpty(GetNameTheoKieuKyHoanThue(periodInfor.KieuKyHoanThue))) ? ("_" + GetNameTheoKieuKyHoanThue(periodInfor.KieuKyHoanThue)) : "") + periodInfor.TuThang.Replace("/", "").PadLeft(6, '0') + "_" + periodInfor.DenThang.Replace("/", "").PadLeft(6, '0') + "_" + versionXML + ".xml")))) : (sFolder + dataFile + "_" + text2 + "_Q" + $"{periodInfor.Quy}{periodInfor.Nam}".PadLeft(6, '0') + "_" + periodInfor.TuThang.Replace("/", "") + "_" + periodInfor.DenThang.Replace("/", "") + "_" + versionXML + ".xml")) : ((text3.IndexOf("~" + periodInfor.menuID + "~") > 0) ? (sFolder + dataFile + ((!string.IsNullOrEmpty(periodInfor.DonViTinhMa)) ? ("_" + periodInfor.DonViTinhMa) : "") + "_" + text2 + "_Q" + $"{periodInfor.Quy}{periodInfor.Nam}".PadLeft(6, '0') + "_" + versionXML + ".xml") : ((!"A1,H1,A3,H3,A8,H8,A9,H9,B5".Contains(periodInfor.menuID) && !periodInfor.menuID.Equals("X2") && !"X1".Equals(WinSettings.TTPeriodInfo.menuID)) ? (periodInfor.menuID.Equals("30") ? (sFolder + dataFile + "_" + periodInfor.TuNgay.Replace("/", "") + "_" + periodInfor.DenNgay.Replace("/", "") + "_" + versionXML + ".xml") : ((!periodInfor.menuID.Equals("01") && !periodInfor.menuID.Equals("I5")) ? ((!periodInfor.menuID.Equals("D1") && !WinSettings.MenuID.Equals("Z5")) ? ((!periodInfor.menuID.Equals("02") && !periodInfor.menuID.Equals("D2")) ? ((!periodInfor.menuID.Equals("71") && !periodInfor.menuID.Equals("D4") && !WinSettings.MenuID.Equals("Z3")) ? ((!periodInfor.menuID.Equals("M4")) ? (sFolder + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_Q" + $"{periodInfor.Quy}{periodInfor.Nam}".PadLeft(6, '0') + "_" + versionXML + ".xml") : (sFolder + dataFile + "_L" + periodInfor.LanXuatBan + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_Q" + $"{periodInfor.Quy}{periodInfor.Nam}".PadLeft(6, '0') + "_" + versionXML + ".xml")) : ((ConvertHelpers.ConvertToInt(periodInfor.Nam) < 2021 || periodInfor.ThuHo.Equals("")) ? (sFolder + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_Q" + $"{periodInfor.Quy}{periodInfor.Nam}".PadLeft(6, '0') + "_" + versionXML + ".xml") : (sFolder + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.ThuHo + "_Q" + $"{periodInfor.Quy}{periodInfor.Nam}".PadLeft(6, '0') + "_" + versionXML + ".xml"))) : ((ConvertHelpers.ConvertToInt(periodInfor.Nam) < 2021) ? (sFolder + dataFile + "_Q" + $"{periodInfor.Quy}{periodInfor.Nam}".PadLeft(6, '0') + "_" + versionXML + ".xml") : (sFolder + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.MaHSDuAnDT + "_Q" + $"{periodInfor.Quy}{periodInfor.Nam}".PadLeft(6, '0') + "_" + versionXML + ".xml"))) : (sFolder + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.TieuMucHachToanVal + "_Q" + $"{periodInfor.Quy}{periodInfor.Nam}".PadLeft(6, '0') + "_" + versionXML + ".xml")) : ((ConvertHelpers.ConvertToInt(periodInfor.Nam) < 2021) ? (sFolder + dataFile + "_Q" + $"{periodInfor.Quy}{periodInfor.Nam}".PadLeft(6, '0') + "_" + versionXML + ".xml") : (sFolder + dataFile + "_" + periodInfor.TieuMucHachToanVal + "_Q" + $"{periodInfor.Quy}{periodInfor.Nam}".PadLeft(6, '0') + "_" + versionXML + ".xml")))) : (sFolder + dataFile + "_Q" + $"{periodInfor.Quy}{periodInfor.Nam}".PadLeft(6, '0') + "_" + periodInfor.TuNgay.Replace("/", "") + "_" + periodInfor.DenNgay.Replace("/", "") + "_" + versionXML + ".xml"))));
			}
			else if (periodInfor.KieuKy == "Y")
			{
				if (!HTKKFormater.isNullString(periodInfor.TuThang) && !HTKKFormater.isNullString(periodInfor.DenThang) && !periodInfor.menuID.Equals("42") && !periodInfor.menuID.Equals("44") && text3.IndexOf("~" + periodInfor.menuID + "~") <= 0)
				{
					text = ((!(periodInfor.menuID == "Q4")) ? (sFolder + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.Nam + "_" + periodInfor.TuThang.Replace("/", "") + "_" + periodInfor.DenThang.Replace("/", "") + "_" + versionXML + ".xml") : (sFolder + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.TuThang.Replace("/", "") + "_" + periodInfor.DenThang.Replace("/", "") + "_" + versionXML + ".xml"));
				}
				else if (periodInfor.menuID.Equals("93") || periodInfor.menuID.Equals("89") || periodInfor.menuID.Equals("44") || periodInfor.menuID.Equals("42") || text3.IndexOf("~" + periodInfor.menuID + "~") > 0)
				{
					text = sFolder + dataFile + ((!string.IsNullOrEmpty(periodInfor.DonViTinhMa)) ? ("_" + periodInfor.DonViTinhMa) : "") + "_" + text2 + "_" + periodInfor.Nam + "_" + periodInfor.TuThang.Replace("/", "") + "_" + periodInfor.DenThang.Replace("/", "") + "_" + versionXML + ".xml";
				}
				else if (periodInfor.menuID.Equals("95"))
				{
					text = sFolder + dataFile + "_L" + periodInfor.SoLan + "_" + periodInfor.Nam + "_" + versionXML + ".xml";
				}
				else if ("78,58".Contains(periodInfor.menuID) || "U9".Equals(periodInfor.menuID))
				{
					text = sFolder + dataFile + "_" + periodInfor.Nam + "_" + periodInfor.TuNgay.Replace("/", "") + "_" + periodInfor.DenNgay.Replace("/", "") + "_" + periodInfor.DonViTinhMa + "_" + versionXML + ".xml";
				}
				else if ("A2,A5,A6,A7,B3,B4".Contains(periodInfor.menuID))
				{
					text = sFolder + dataFile + "_" + periodInfor.Nam + "_" + periodInfor.TuNgay.Replace("/", "") + "_" + periodInfor.DenNgay.Replace("/", "") + "_" + versionXML + ".xml";
				}
				else if ("C1,C2,C3,C4,C5,X3,X4".Contains(periodInfor.menuID))
				{
					text = sFolder + dataFile + "_" + periodInfor.Nam + "_" + periodInfor.TuNgay.Replace("/", "") + "_" + periodInfor.DenNgay.Replace("/", "") + "_" + versionXML + ".xml";
				}
				else if (periodInfor.menuID.Equals("K8") || periodInfor.menuID.Equals("03") || periodInfor.menuID.Equals("83") || periodInfor.menuID.Equals("N8") || text4.IndexOf("~" + periodInfor.menuID + "~") >= 0)
				{
					text = sFolder + dataFile + "_" + periodInfor.Nam + "_" + periodInfor.TuNgay.Replace("/", "") + "_" + periodInfor.DenNgay.Replace("/", "") + "_" + versionXML + ".xml";
				}
				else if (periodInfor.menuID.Equals("E8"))
				{
					text = sFolder + dataFile + "_" + periodInfor.TieuMucHachToanVal + "_" + periodInfor.Nam + "_" + periodInfor.TuNgay.Replace("/", "") + "_" + periodInfor.DenNgay.Replace("/", "") + "_" + versionXML + ".xml";
				}
				else if ("53,54".Contains(periodInfor.menuID))
				{
					text = sFolder + dataFile + "_" + periodInfor.Nam + "_" + periodInfor.TuNgay.Replace("/", "") + "_" + periodInfor.DenNgay.Replace("/", "") + "_" + versionXML + ".xml";
				}
				else if (periodInfor.menuID.Equals("63"))
				{
					text = sFolder + dataFile + "_" + periodInfor.Nam;
					bool flag = false;
					string text5 = "";
					if (periodInfor.dicVar.Count > 0)
					{
						flag = Convert.ToBoolean(periodInfor.dicVar["CoSoMoiThanhLap"]);
						text5 = (string)periodInfor.dicVar["NgayThanhLap"];
						if (flag)
						{
							text = text + "_CSM_" + text5.Replace("/", "");
						}
					}
					text = text + "_" + versionXML + ".xml";
				}
				else if (!periodInfor.menuID.Equals("L1"))
				{
					text = ((!periodInfor.menuID.Equals("I2") && !periodInfor.menuID.Equals("L9") && !periodInfor.menuID.Equals("I1")) ? (("C6~C7~T5~H4~Y4".IndexOf(periodInfor.menuID) < 0) ? (sFolder + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.Nam + "_" + versionXML + ".xml") : (sFolder + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.Nam + "_" + periodInfor.TuNgay.Replace("/", "") + "_" + periodInfor.DenNgay.Replace("/", "") + "_" + versionXML + ".xml")) : (sFolder + dataFile + "_L" + periodInfor.LanXuatBan + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.Nam + "_" + versionXML + ".xml"));
				}
				else
				{
					text = sFolder + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.Nam;
					bool flag = false;
					string text5 = "";
					if (periodInfor.dicVar.Count > 0)
					{
						flag = Convert.ToBoolean(periodInfor.dicVar["CoSoMoiThanhLap"]);
						text5 = (string)periodInfor.dicVar["NgayThanhLap"];
						if (flag)
						{
							text = text + "_CSM_" + text5.Replace("/", "");
						}
					}
					text = text + "_" + versionXML + ".xml";
				}
			}
			else if (!(periodInfor.KieuKy == "D"))
			{
				text = ((periodInfor.KieuKy == "D_D") ? (sFolder + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.TuNgay.Replace("/", "") + "_" + periodInfor.DenNgay.Replace("/", "") + "_" + versionXML + ".xml") : ((!(periodInfor.KieuKy == "K")) ? (sFolder + dataFile) : (sFolder + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_K" + $"{periodInfor.Quy}{periodInfor.Nam}".PadLeft(6, '0') + "_" + versionXML + ".xml")));
			}
			else
			{
				text = ((!HTKKFormater.isNullString(periodInfor.LanXuatBan) && !HTKKFormater.isNullString(periodInfor.TieuMucHachToanVal) && !HTKKFormater.isNullString(text2)) ? (sFolder + dataFile + ((!string.IsNullOrEmpty(periodInfor.DonViTinhMa)) ? ("_" + periodInfor.DonViTinhMa) : "") + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + text2 + "_L" + GeneralFunctions.removeKyTuDacBietFileName(periodInfor.LanXuatBan) + "_" + periodInfor.Ngay + periodInfor.Thang + periodInfor.Nam + "_" + versionXML + ".xml") : ((HTKKFormater.isNullString(periodInfor.LanXuatBan) && !HTKKFormater.isNullString(periodInfor.TieuMucHachToanVal) && !HTKKFormater.isNullString(text2)) ? (sFolder + dataFile + ((!string.IsNullOrEmpty(periodInfor.DonViTinhMa)) ? ("_" + periodInfor.DonViTinhMa) : "") + "_" + text2 + "_" + periodInfor.Ngay + periodInfor.Thang + periodInfor.Nam + "_" + versionXML + ".xml") : (periodInfor.menuID.Equals("09") ? (sFolder + dataFile + "_" + versionXML + ".xml") : ((!periodInfor.menuID.Equals("47") && !periodInfor.menuID.Equals("K7")) ? ((!periodInfor.menuID.Equals("71") && !periodInfor.menuID.Equals("D4") && !WinSettings.MenuID.Equals("Z3")) ? (periodInfor.menuID.Equals("M5") ? (sFolder + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.Ngay + periodInfor.Thang + periodInfor.Nam + "_" + versionXML + ".xml") : ((!periodInfor.menuID.Equals("T6") && !periodInfor.menuID.Equals("U2")) ? (periodInfor.menuID.Equals("Y5") ? (sFolder + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_CN" + periodInfor.TieuMucHachToanVal + "_" + periodInfor.TuNgay.Replace("/", "") + "_" + periodInfor.Ngay + periodInfor.Thang + periodInfor.Nam + "_" + versionXML + ".xml") : (periodInfor.menuID.Equals("G3") ? (sFolder + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_L" + periodInfor.LanXuatBan + "_" + periodInfor.Ngay + periodInfor.Thang + periodInfor.Nam + "_" + versionXML + ".xml") : ((!periodInfor.menuID.Equals("P6") && !periodInfor.menuID.Equals("P1") && !periodInfor.menuID.Equals("X7") && !periodInfor.menuID.Equals("Y9")) ? (("~O1~O2~O3~O4~O5~".IndexOf("~" + periodInfor.menuID + "~") >= 0) ? (sFolder + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + ((!string.IsNullOrEmpty(periodInfor.MaHSDuAnDT.Trim())) ? ("_" + periodInfor.MaHSDuAnDT.Trim()) : "") + "_" + periodInfor.Ngay + periodInfor.Thang + periodInfor.Nam + "_" + versionXML + ".xml") : ((periodInfor.menuID == "U1" || periodInfor.menuID == "Y3") ? (sFolder + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_NN" + periodInfor.TieuMucHachToanVal + "_L" + periodInfor.ToChucCNDNVonChuyenNhuong + "_" + periodInfor.Ngay + periodInfor.Thang + periodInfor.Nam + "_" + versionXML + ".xml") : ((!"~K2~K3~J7~E6~E7~D7~R6~X5~R9~E9~".Contains(periodInfor.menuID)) ? (sFolder + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + neuCoNganhNgheKD() + "_" + periodInfor.Ngay + periodInfor.Thang + periodInfor.Nam + "_" + versionXML + ".xml") : (sFolder + dataFile + "_" + periodInfor.DonViTinhMa + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.Ngay + periodInfor.Thang + periodInfor.Nam + "_" + versionXML + ".xml")))) : (sFolder + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_L" + periodInfor.LanXuatBan + "_" + periodInfor.MaHSDuAnDT + "_" + periodInfor.Ngay + periodInfor.Thang + periodInfor.Nam + "_" + versionXML + ".xml")))) : (sFolder + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_CN" + periodInfor.TieuMucHachToanVal + "_" + periodInfor.Ngay + periodInfor.Thang + periodInfor.Nam + "_" + versionXML + ".xml"))) : ((ConvertHelpers.ConvertToInt(periodInfor.Nam) < 2021 || periodInfor.ThuHo.Equals("")) ? (sFolder + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.Ngay + periodInfor.Thang + periodInfor.Nam + "_" + versionXML + ".xml") : (sFolder + dataFile + ((!string.IsNullOrEmpty(GetNameNopTheoTinhCQT())) ? ("_" + GetNameNopTheoTinhCQT()) : "") + "_" + periodInfor.ThuHo + "_" + periodInfor.Ngay + periodInfor.Thang + periodInfor.Nam + "_" + versionXML + ".xml"))) : (sFolder + dataFile + "_" + periodInfor.Ngay + periodInfor.Thang + periodInfor.Nam + "_" + (periodInfor.CTDC ? "CTDC" : "LNDC") + "_" + versionXML + ".xml")))));
				if (periodInfor.menuID == "P7")
				{
					text = text.Replace("_" + periodInfor.Ngay + periodInfor.Thang + periodInfor.Nam, "");
				}
			}
			if (periodInfor.LoaiToKhai.Equals("B") || periodInfor.LoaiToKhai.Equals("T"))
			{
				text = GetFileNameLoadKhbsNew(periodInfor, text, sFolder, isLoadSavedFile);
			}
			return text;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	public static string neuCoNganhNgheKD()
	{
		if (string.IsNullOrEmpty(WinSettings.TTPeriodInfo.NganhNgheKinhDoanh) || string.IsNullOrEmpty(WinSettings.TTPeriodInfo.TieuMucHachToanVal))
		{
			return "";
		}
		string versionXML = "";
		string versionKHBS = "";
		string dataFileMenu = "";
		WinFnProcess.GetInfoInMenu(WinSettings.MenuID, 0, ref versionXML, ref versionKHBS, ref dataFileMenu);
		string catalogueDMNNKDItem = GeneralFunctions.GetCatalogueDMNNKDItem(dataFileMenu);
		if (!string.IsNullOrEmpty(catalogueDMNNKDItem))
		{
			string tieuMucHachToanVal = WinSettings.TTPeriodInfo.TieuMucHachToanVal;
			string[] array = catalogueDMNNKDItem.Split(new string[1] { "###" }, StringSplitOptions.None);
			if (array.Length > 2)
			{
				WinSettings.TTPeriodInfo.NganhNgheKinhDoanh = array[2];
				return "_NN" + tieuMucHachToanVal;
			}
		}
		return "";
	}

	public static string GetFileNameLoadKhbsNew(HTKKPeriodInfo periodInfor, string sDataFile, string sFolder, bool isLoadSavedFile)
	{
		if (string.IsNullOrEmpty(sDataFile))
		{
			return "";
		}
		string text = sDataFile.Split(new char[1] { '\\' })[^1];
		string result = sFolder + text;
		if (isLoadSavedFile && sDataFile.IndexOf("bs" + periodInfor.SoLan) >= 0)
		{
			return result;
		}
		if (isLoadSavedFile && sDataFile.IndexOf("tt" + periodInfor.SoLan) >= 0)
		{
			return result;
		}
		try
		{
			if (periodInfor.LoaiToKhai.Equals("B"))
			{
				if (isLoadSavedFile)
				{
					result = ((!string.IsNullOrEmpty(periodInfor.SoLan)) ? (sFolder + "bs" + periodInfor.SoLan + "_" + text) : (sFolder + "bs1_" + text));
				}
				else
				{
					if (string.IsNullOrEmpty(periodInfor.SoLan))
					{
						return result;
					}
					if (periodInfor.SoLan != "1" && "N7~".IndexOf(WinSettings.MenuID) <= 0)
					{
						text = text.Substring(4, text.Length - 4);
						text = "bs" + (int.Parse(periodInfor.SoLan) - 1) + "_" + text;
						result = sFolder + text;
					}
					else
					{
						text = text.Replace("bs1_", "");
						result = sFolder + text;
					}
				}
			}
			if (periodInfor.LoaiToKhai.Equals("T"))
			{
				if (isLoadSavedFile)
				{
					result = ((!string.IsNullOrEmpty(periodInfor.SoLan)) ? (sFolder + "tt" + periodInfor.SoLan + "_" + text) : (sFolder + "tt1_" + text));
				}
				else
				{
					if ("U8~".Contains(periodInfor.menuID))
					{
						if (!string.IsNullOrEmpty(periodInfor.SoLan) && periodInfor.SoLan != "1" && "N7~".IndexOf(WinSettings.MenuID) <= 0)
						{
							text = text.Substring(4, text.Length - 4);
							text = "tt" + (int.Parse(periodInfor.SoLan) - 1) + "_" + text;
							result = sFolder + text;
						}
						else
						{
							text = text.Replace("tt1_", "");
							result = sFolder + text;
						}
					}
					if (string.IsNullOrEmpty(periodInfor.SoLan))
					{
						return result;
					}
					if (!string.IsNullOrEmpty(periodInfor.SoLan) && periodInfor.SoLan != "1" && "N7~".IndexOf(WinSettings.MenuID) <= 0)
					{
						text = text.Substring(4, text.Length - 4);
						text = "tt" + (int.Parse(periodInfor.SoLan) - 1) + "_" + text;
						result = sFolder + text;
					}
					else
					{
						text = text.Replace("tt1_", "");
						result = sFolder + text;
					}
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

	public static List<string> getAllPathOfTable(string templateXmlPath, ref string rootNode, ref string childNodeOfRoot, ref string childNodeOfRoot2)
	{
		//IL_0037: Unknown result type (might be due to invalid IL or missing references)
		//IL_003d: Expected O, but got Unknown
		//IL_0083: Unknown result type (might be due to invalid IL or missing references)
		//IL_008a: Expected O, but got Unknown
		//IL_00a6: Unknown result type (might be due to invalid IL or missing references)
		//IL_00ac: Invalid comparison between Unknown and I4
		//IL_00d6: Unknown result type (might be due to invalid IL or missing references)
		//IL_00dd: Expected O, but got Unknown
		//IL_00f9: Unknown result type (might be due to invalid IL or missing references)
		//IL_00ff: Invalid comparison between Unknown and I4
		//IL_0123: Unknown result type (might be due to invalid IL or missing references)
		//IL_012a: Expected O, but got Unknown
		//IL_012d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0133: Invalid comparison between Unknown and I4
		List<string> list = new List<string>();
		string text = File.ReadAllText(templateXmlPath, Encoding.UTF8).Trim();
		if (!text.StartsWith("<?xml"))
		{
			text = "<?xml version=\"1.0\" encoding=\"utf-8\"?>\n" + text;
		}
		XmlDocument val = new XmlDocument();
		val.LoadXml(text);
		XmlElement documentElement = val.DocumentElement;
		if (documentElement == null)
		{
			return list;
		}
		rootNode = ((XmlNode)documentElement).Name;
		foreach (XmlNode childNode in ((XmlNode)documentElement).ChildNodes)
		{
			XmlNode val2 = childNode;
			if (string.IsNullOrEmpty(childNodeOfRoot))
			{
				childNodeOfRoot = val2.Name;
			}
			if ((int)val2.NodeType != 1)
			{
				continue;
			}
			foreach (XmlNode childNode2 in val2.ChildNodes)
			{
				XmlNode val3 = childNode2;
				if (string.IsNullOrEmpty(childNodeOfRoot2))
				{
					childNodeOfRoot2 = val3.Name;
				}
				if ((int)val3.NodeType != 1)
				{
					continue;
				}
				foreach (XmlNode childNode3 in val3.ChildNodes)
				{
					XmlNode val4 = childNode3;
					if ((int)val4.NodeType == 1)
					{
						list.Add(val4.Name);
					}
				}
				break;
			}
		}
		return list;
	}

	public static string WinDoWorkExportXML_Large(FpSpread winFpSpread, List<HTKKCMCSheetInfo> sheetInfos, string outputXmlPath)
	{
		//IL_001e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0024: Expected O, but got Unknown
		//IL_007c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0083: Expected O, but got Unknown
		//IL_0093: Unknown result type (might be due to invalid IL or missing references)
		//IL_009a: Expected O, but got Unknown
		string text = "";
		try
		{
			if (File.Exists(outputXmlPath))
			{
				File.Delete(outputXmlPath);
			}
			XmlDocument val = new XmlDocument();
			for (int i = 0; i < sheetInfos.Count; i++)
			{
				HTKKCMCSheetInfo hTKKCMCSheetInfo = sheetInfos[i];
				if (hTKKCMCSheetInfo != null && hTKKCMCSheetInfo.TaiDuLieuRatLon != "1")
				{
					val = WinDoWorkGetExportSheetXML(winFpSpread.Sheets[i], hTKKCMCSheetInfo);
					if (i > 0)
					{
						XmlDocument val2 = new XmlDocument();
						val2.Load(outputXmlPath);
						XmlNamespaceManager val3 = new XmlNamespaceManager(val2.NameTable);
						val3.AddNamespace("ns", "http://kekhaithue.gdt.gov.vn/TKhaiThue");
						XmlNode val4 = ((XmlNode)val2).SelectSingleNode("//ns:PLuc", val3);
						XmlNode documentElement = (XmlNode)(object)val.DocumentElement;
						XmlNode val5 = val2.ImportNode(documentElement, true);
						val4.AppendChild(val5);
						if (WinSettings.ToKhaiInfo.KieuKHBS == "3" && WinSettings.TTPeriodInfo.LoaiToKhai == "B" && i >= sheetInfos.Count - 2)
						{
							xmlExportTTinChungKHBS(val2);
						}
						val2.Save(outputXmlPath);
					}
				}
				if (i == 0)
				{
					if (hTKKCMCSheetInfo != null && hTKKCMCSheetInfo.TaiDuLieuRatLon == "1")
					{
						text += exportToKhaiDuLieuRatLon(outputXmlPath, hTKKCMCSheetInfo);
						continue;
					}
					val = exportToKhaiVaThongTinChung(val);
					val.Save(outputXmlPath);
					if (WinSettings.MenuID == "P2")
					{
						XmlNode val6 = ((XmlNode)val).SelectSingleNode("//DLyThue");
						if (val6 != null)
						{
							XmlNode parentNode = val6.ParentNode;
							parentNode.RemoveChild(val6);
						}
					}
					val.Save(outputXmlPath);
				}
				if (sheetInfos[i] != null && sheetInfos[i].TaiDuLieuRatLon == "1")
				{
					string text2 = "tmpExport_Pluc.xml";
					string text3 = HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst;
					string text4 = text3 + "\\" + text2;
					string text5 = "tmpExport_Full.xml";
					string text6 = HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst;
					string text7 = text6 + "\\" + text5;
					text += exportPhuLucDuLieuRatLonDynamic(text4, hTKKCMCSheetInfo);
					text += setStaticXml(text4, hTKKCMCSheetInfo);
					text += ghepPhuLucVaoToKhai(outputXmlPath, text4, text7);
					File.Delete(text4);
					File.Delete(outputXmlPath);
					File.Move(text7, outputXmlPath);
				}
			}
			return text;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return text + ex.Message;
		}
	}

	public static string exportToKhaiDuLieuRatLon(string outputXmlPath, HTKKCMCSheetInfo sInfo)
	{
		//IL_006d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0073: Expected O, but got Unknown
		//IL_01b9: Unknown result type (might be due to invalid IL or missing references)
		//IL_01c0: Expected O, but got Unknown
		try
		{
			if (sInfo == null)
			{
				return "Không có thông tin sheet tờ khai.";
			}
			string text = string.Format("{0}{1}_xml_{2}.xml", WinSettings.ToKhaiInfo.InterfaceTemplateXml, sInfo.DataFile, WinSettings.ToKhaiInfo.VersionXML.Replace(".", ""));
			if (!File.Exists(text))
			{
				return "Không tìm thấy template XML: " + text;
			}
			XmlDocument val = new XmlDocument();
			val.PreserveWhitespace = true;
			val.Load(text);
			setStaticXmlForLargeTemplate(val, sInfo.StaticCells);
			soLanTT = WinSettings.TTinChung.soLan;
			val = exportToKhaiVaThongTinChung(val);
			Dictionary<string, List<HTKKCMCTableInfo>> dictionary = new Dictionary<string, List<HTKKCMCTableInfo>>(StringComparer.Ordinal);
			if (sInfo.Tables != null)
			{
				foreach (HTKKCMCTableInfo table in sInfo.Tables)
				{
					if (table != null && !string.IsNullOrEmpty(table.TablePath))
					{
						string key = normalizeXmlPath(table.TablePath);
						if (!dictionary.ContainsKey(key))
						{
							dictionary.Add(key, new List<HTKKCMCTableInfo>());
						}
						dictionary[key].Add(table);
					}
				}
			}
			foreach (string key2 in dictionary.Keys)
			{
				if (((XmlNode)val).SelectSingleNode(key2) == null)
				{
					return "Không tìm thấy TablePath '" + key2 + "' trong template XML.";
				}
			}
			XmlWriterSettings val2 = new XmlWriterSettings();
			val2.Encoding = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false);
			val2.Indent = true;
			val2.CloseOutput = true;
			XmlWriter val3 = XmlWriter.Create(outputXmlPath, val2);
			try
			{
				val3.WriteStartDocument();
				writeLargeToKhaiNode((XmlNode)(object)val.DocumentElement, val3, "", dictionary, "");
				val3.WriteEndDocument();
			}
			finally
			{
				((IDisposable)val3)?.Dispose();
			}
			return "";
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return ex.Message;
		}
	}

	private static void setStaticXmlForLargeTemplate(XmlDocument doc, List<HTKKCMCCellInfo> staticCells)
	{
		if (doc == null || staticCells == null)
		{
			return;
		}
		foreach (HTKKCMCCellInfo staticCell in staticCells)
		{
			if (staticCell == null || string.IsNullOrEmpty(staticCell.Path))
			{
				continue;
			}
			XmlNode val = ((XmlNode)doc).SelectSingleNode(staticCell.Path);
			if (val != null)
			{
				string text = ((staticCell.Value == null) ? "" : staticCell.Value.Trim());
				if (staticCell.Controltype == "16" && text.Length == 0)
				{
					text = "0";
				}
				else if (staticCell.Controltype == "2")
				{
					text = ((text == "1" || text.ToLower() == "x" || text.ToLower() == "true") ? "1" : "0");
				}
				else if (staticCell.Controltype == "26" && text.Length == 13)
				{
					text = text.Substring(0, 10) + "-" + text.Substring(10, 3);
				}
				val.InnerText = text;
				setNilAttributeByControlType(val, staticCell.Controltype, text.Length == 0);
			}
		}
	}

	private static void setNilAttributeByControlType(XmlNode node, string controlType, bool isEmpty)
	{
		if (node == null || node.Attributes == null)
		{
			return;
		}
		XmlAttribute val = node.Attributes["nil", "http://www.w3.org/2001/XMLSchema-instance"];
		int num;
		switch (controlType)
		{
		default:
			num = ((controlType == "94") ? 1 : 0);
			break;
		case "14":
		case "26":
		case "33":
		case "96":
			num = 1;
			break;
		}
		bool flag = (byte)num != 0;
		if (isEmpty && flag)
		{
			if (val == null)
			{
				val = node.OwnerDocument.CreateAttribute("xsi", "nil", "http://www.w3.org/2001/XMLSchema-instance");
				((XmlNamedNodeMap)node.Attributes).SetNamedItem((XmlNode)(object)val);
			}
			((XmlNode)val).Value = "true";
		}
		else if (val != null)
		{
			node.Attributes.Remove(val);
		}
	}

	private static void writeLargeToKhaiNode(XmlNode node, XmlWriter writer, string parentPath, Dictionary<string, List<HTKKCMCTableInfo>> tablesByPath, string inheritedDefaultNamespace)
	{
		//IL_0014: Unknown result type (might be due to invalid IL or missing references)
		//IL_001a: Invalid comparison between Unknown and I4
		//IL_0090: Unknown result type (might be due to invalid IL or missing references)
		//IL_0096: Expected O, but got Unknown
		if (node == null)
		{
			return;
		}
		if ((int)node.NodeType != 1)
		{
			node.WriteTo(writer);
			return;
		}
		string text = normalizeXmlPath(parentPath + "/" + node.Name);
		if (tablesByPath.TryGetValue(text, out var value))
		{
			writeLargeTable(node, writer, value, inheritedDefaultNamespace);
			return;
		}
		string effectiveDefaultNamespace = getEffectiveDefaultNamespace(node, inheritedDefaultNamespace);
		writeLargeStartElement(node, writer, effectiveDefaultNamespace);
		foreach (XmlNode childNode in node.ChildNodes)
		{
			XmlNode node2 = childNode;
			writeLargeToKhaiNode(node2, writer, text, tablesByPath, effectiveDefaultNamespace);
		}
		writer.WriteEndElement();
	}

	private static void writeLargeTable(XmlNode tableTemplate, XmlWriter writer, List<HTKKCMCTableInfo> tables, string inheritedDefaultNamespace)
	{
		if (tables == null || tables.Count == 0)
		{
			return;
		}
		int num = 1;
		foreach (HTKKCMCTableInfo table in tables)
		{
			if (table == null || table.RowsInfo == null)
			{
				continue;
			}
			if (string.IsNullOrEmpty(table.RowsInfo.RowPath))
			{
				foreach (HTKKCMCRowValue rowsValue in table.RowsValues)
				{
					XmlNode val = tableTemplate.CloneNode(true);
					setLargeRowValue(val, table, rowsValue);
					setLargeTableId(val, num.ToString());
					num++;
					writeXmlNodeWithDefaultNamespace(val, writer, inheritedDefaultNamespace);
				}
				continue;
			}
			XmlNode val2 = tableTemplate.CloneNode(true);
			setLargeTableId(val2, num.ToString());
			num++;
			XmlNode val3 = val2.SelectSingleNode(table.RowsInfo.RowPath);
			if (val3 == null)
			{
				throw new InvalidOperationException("Không tìm thấy RowPath '" + table.RowsInfo.RowPath + "' trong TablePath '" + table.TablePath + "'.");
			}
			writeLargeTableContainer(val2, val3, writer, table, inheritedDefaultNamespace);
		}
	}

	private static void writeLargeTableContainer(XmlNode node, XmlNode rowTemplate, XmlWriter writer, HTKKCMCTableInfo table, string inheritedDefaultNamespace)
	{
		//IL_006c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0072: Invalid comparison between Unknown and I4
		//IL_00ae: Unknown result type (might be due to invalid IL or missing references)
		//IL_00b4: Expected O, but got Unknown
		if (object.ReferenceEquals(node, rowTemplate))
		{
			foreach (HTKKCMCRowValue rowsValue in table.RowsValues)
			{
				XmlNode val = rowTemplate.CloneNode(true);
				setLargeRowValue(val, table, rowsValue);
				writeXmlNodeWithDefaultNamespace(val, writer, inheritedDefaultNamespace);
			}
			return;
		}
		if ((int)node.NodeType != 1)
		{
			node.WriteTo(writer);
			return;
		}
		string effectiveDefaultNamespace = getEffectiveDefaultNamespace(node, inheritedDefaultNamespace);
		writeLargeStartElement(node, writer, effectiveDefaultNamespace);
		foreach (XmlNode childNode in node.ChildNodes)
		{
			XmlNode node2 = childNode;
			writeLargeTableContainer(node2, rowTemplate, writer, table, effectiveDefaultNamespace);
		}
		writer.WriteEndElement();
	}

	private static void writeXmlNodeWithDefaultNamespace(XmlNode node, XmlWriter writer, string inheritedDefaultNamespace)
	{
		//IL_000f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0015: Invalid comparison between Unknown and I4
		//IL_004c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0052: Expected O, but got Unknown
		if (node == null)
		{
			return;
		}
		if ((int)node.NodeType != 1)
		{
			node.WriteTo(writer);
			return;
		}
		string effectiveDefaultNamespace = getEffectiveDefaultNamespace(node, inheritedDefaultNamespace);
		writeLargeStartElement(node, writer, effectiveDefaultNamespace);
		foreach (XmlNode childNode in node.ChildNodes)
		{
			XmlNode node2 = childNode;
			writeXmlNodeWithDefaultNamespace(node2, writer, effectiveDefaultNamespace);
		}
		writer.WriteEndElement();
	}

	private static string getEffectiveDefaultNamespace(XmlNode node, string inheritedDefaultNamespace)
	{
		if (node != null && node.Attributes != null && node.Attributes["xmlns"] != null)
		{
			return ((XmlNode)node.Attributes["xmlns"]).Value;
		}
		return (inheritedDefaultNamespace == null) ? "" : inheritedDefaultNamespace;
	}

	private static void writeLargeStartElement(XmlNode node, XmlWriter writer, string defaultNamespace)
	{
		//IL_0068: Unknown result type (might be due to invalid IL or missing references)
		//IL_006e: Expected O, but got Unknown
		if (string.IsNullOrEmpty(node.Prefix))
		{
			writer.WriteStartElement("", node.LocalName, defaultNamespace);
		}
		else
		{
			writer.WriteStartElement(node.Prefix, node.LocalName, node.NamespaceURI);
		}
		if (node.Attributes == null)
		{
			return;
		}
		foreach (XmlAttribute item in (XmlNamedNodeMap)node.Attributes)
		{
			XmlAttribute val = item;
			if (!(((XmlNode)val).Name == "xmlns"))
			{
				((XmlNode)val).WriteTo(writer);
			}
		}
	}

	private static void setLargeRowValue(XmlNode rowNode, HTKKCMCTableInfo table, HTKKCMCRowValue row)
	{
		if (rowNode == null || table == null || table.RowsInfo == null || row == null)
		{
			return;
		}
		for (int i = 0; i < table.RowsInfo.Cells.Count; i++)
		{
			HTKKCMCTableCellInfo hTKKCMCTableCellInfo = table.RowsInfo.Cells[i];
			if (hTKKCMCTableCellInfo == null || string.IsNullOrEmpty(hTKKCMCTableCellInfo.Path))
			{
				continue;
			}
			XmlNode val = rowNode.SelectSingleNode(hTKKCMCTableCellInfo.Path);
			if (val != null)
			{
				string text = ((i < row.RowValue.Count && row.RowValue[i] != null) ? row.RowValue[i].Trim() : "");
				if (hTKKCMCTableCellInfo.Controltype == "16" && text.Length == 0)
				{
					text = "0";
				}
				else if (hTKKCMCTableCellInfo.Controltype == "2")
				{
					text = ((text == "1" || text.ToLower() == "x" || text.ToLower() == "true") ? "1" : "0");
				}
				else if (hTKKCMCTableCellInfo.Controltype == "26" && text.Length == 13)
				{
					text = text.Substring(0, 10) + "-" + text.Substring(10, 3);
				}
				text = RemoveValueSpecial(hTKKCMCTableCellInfo, text);
				XmlAttribute val2 = ((val.Attributes == null) ? null : val.Attributes["nillable"]);
				bool flag = val2 != null && ConvertHelpers.ConvertToBool(((XmlNode)val2).Value);
				if (val2 != null)
				{
					val.Attributes.Remove(val2);
				}
				val.InnerText = text;
				setNilAttributeByControlType(val, hTKKCMCTableCellInfo.Controltype, text.Length == 0);
				if (text.Length == 0 && flag && val.Attributes["nil", "http://www.w3.org/2001/XMLSchema-instance"] == null)
				{
					XmlAttribute val3 = val.OwnerDocument.CreateAttribute("xsi", "nil", "http://www.w3.org/2001/XMLSchema-instance");
					((XmlNode)val3).Value = "true";
					((XmlNamedNodeMap)val.Attributes).SetNamedItem((XmlNode)(object)val3);
				}
			}
		}
	}

	private static void setLargeTableId(XmlNode node, string tableId)
	{
		if (node != null && node.Attributes != null && !string.IsNullOrEmpty(tableId))
		{
			if (node.Attributes["minOccurs"] != null)
			{
				node.Attributes.Remove(node.Attributes["minOccurs"]);
			}
			if (node.Attributes["id"] != null)
			{
				((XmlNode)node.Attributes["id"]).Value = "ID_" + tableId;
			}
			if (node.Attributes["ID"] != null)
			{
				((XmlNode)node.Attributes["ID"]).Value = "ID_" + tableId;
			}
		}
	}

	private static string normalizeXmlPath(string path)
	{
		if (string.IsNullOrEmpty(path))
		{
			return "";
		}
		string text = path.Replace("\\", "/").Trim();
		while (text.StartsWith("/"))
		{
			text = text.Substring(1);
		}
		while (text.EndsWith("/"))
		{
			text = text.Substring(0, text.Length - 1);
		}
		return text;
	}

	public static string exportPhuLucDuLieuRatLonDynamic(string outputXmlPath, HTKKCMCSheetInfo sInfo)
	{
		//IL_0092: Unknown result type (might be due to invalid IL or missing references)
		//IL_0098: Expected O, but got Unknown
		//IL_017a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0181: Expected O, but got Unknown
		try
		{
			if (sInfo == null || sInfo.Tables == null || sInfo.Tables.Count == 0)
			{
				return "";
			}
			string text = sInfo.DataFile + "_xml_" + WinSettings.ToKhaiInfo.VersionXML.Replace(".", "") + ".xml";
			string text2 = WinSettings.ToKhaiInfo.InterfaceTemplateXml + text;
			if (!File.Exists(text2))
			{
				return "Không tìm thấy template XML: " + text2;
			}
			XmlDocument val = new XmlDocument();
			val.PreserveWhitespace = true;
			val.Load(text2);
			Dictionary<string, List<HTKKCMCTableInfo>> dictionary = new Dictionary<string, List<HTKKCMCTableInfo>>(StringComparer.Ordinal);
			foreach (HTKKCMCTableInfo table in sInfo.Tables)
			{
				if (table != null && !string.IsNullOrEmpty(table.TablePath))
				{
					string text3 = resolveLargeTemplatePath(val, table.TablePath);
					if (string.IsNullOrEmpty(text3))
					{
						return "Không tìm thấy TablePath '" + table.TablePath + "' trong template XML.";
					}
					if (!dictionary.ContainsKey(text3))
					{
						dictionary.Add(text3, new List<HTKKCMCTableInfo>());
					}
					dictionary[text3].Add(table);
				}
			}
			XmlWriterSettings val2 = new XmlWriterSettings();
			val2.Encoding = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false);
			val2.Indent = true;
			val2.CloseOutput = true;
			XmlWriter val3 = XmlWriter.Create(outputXmlPath, val2);
			try
			{
				val3.WriteStartDocument();
				writeLargeToKhaiNode((XmlNode)(object)val.DocumentElement, val3, "", dictionary, "");
				val3.WriteEndDocument();
			}
			finally
			{
				((IDisposable)val3)?.Dispose();
			}
			return "";
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return ex.Message;
		}
	}

	private static string resolveLargeTemplatePath(XmlDocument template, string configuredPath)
	{
		//IL_00aa: Unknown result type (might be due to invalid IL or missing references)
		//IL_00b1: Expected O, but got Unknown
		//IL_00b4: Unknown result type (might be due to invalid IL or missing references)
		//IL_00ba: Invalid comparison between Unknown and I4
		if (template == null || template.DocumentElement == null)
		{
			return "";
		}
		string text = normalizeXmlPath(configuredPath);
		string[] array = text.Split(new char[1] { '/' }, StringSplitOptions.RemoveEmptyEntries);
		if (array.Length == 0 || !isEquivalentXmlName(((XmlNode)template.DocumentElement).Name, array[0]))
		{
			return "";
		}
		XmlNode val = (XmlNode)(object)template.DocumentElement;
		string text2 = val.Name;
		for (int i = 1; i < array.Length; i++)
		{
			XmlNode val2 = null;
			XmlNode val3 = null;
			int num = 0;
			foreach (XmlNode childNode in val.ChildNodes)
			{
				XmlNode val4 = childNode;
				if ((int)val4.NodeType == 1)
				{
					if (string.Equals(val4.Name, array[i], StringComparison.Ordinal))
					{
						val2 = val4;
						break;
					}
					if (isEquivalentXmlName(val4.Name, array[i]))
					{
						val3 = val4;
						num++;
					}
				}
			}
			val = ((val2 != null) ? val2 : ((num == 1) ? val3 : null));
			if (val == null)
			{
				return "";
			}
			text2 = text2 + "/" + val.Name;
		}
		return text2;
	}

	private static bool isEquivalentXmlName(string left, string right)
	{
		if (string.Equals(left, right, StringComparison.Ordinal))
		{
			return true;
		}
		if (left == null || right == null)
		{
			return false;
		}
		return string.Equals(left.Replace("-", "_").Replace("_", ""), right.Replace("-", "_").Replace("_", ""), StringComparison.OrdinalIgnoreCase);
	}

	public static string exportPhuLucDuLieuRatLon(string outputXmlPath, HTKKCMCSheetInfo sInfo)
	{
		//IL_00f4: Unknown result type (might be due to invalid IL or missing references)
		//IL_00fb: Expected O, but got Unknown
		//IL_011b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0121: Invalid comparison between Unknown and I4
		//IL_0165: Unknown result type (might be due to invalid IL or missing references)
		//IL_016c: Invalid comparison between Unknown and I4
		//IL_01a3: Unknown result type (might be due to invalid IL or missing references)
		//IL_01a9: Invalid comparison between Unknown and I4
		//IL_0475: Unknown result type (might be due to invalid IL or missing references)
		//IL_047b: Invalid comparison between Unknown and I4
		//IL_051f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0525: Invalid comparison between Unknown and I4
		//IL_054f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0556: Invalid comparison between Unknown and I4
		string text = "";
		try
		{
			if (sInfo == null || sInfo.Tables.Count == 0)
			{
				return text;
			}
			string dataFile = sInfo.DataFile;
			string text2 = dataFile + "_xml_" + WinSettings.ToKhaiInfo.VersionXML.Replace(".", "") + ".xml";
			string text3 = WinSettings.ToKhaiInfo.InterfaceTemplateXml + text2;
			Dictionary<string, string[]> tableFromSInfo = GeneralFunctions.getTableFromSInfo(sInfo);
			if (tableFromSInfo.Count == 0)
			{
				return text;
			}
			string text4 = "";
			using (Dictionary<string, string[]>.Enumerator enumerator = tableFromSInfo.GetEnumerator())
			{
				if (enumerator.MoveNext())
				{
					text4 = enumerator.Current.Value[0];
				}
			}
			XmlReader reader = XmlReader.Create(text3);
			try
			{
				XmlWriterSettings val = new XmlWriterSettings();
				val.Indent = true;
				XmlWriter val2 = XmlWriter.Create(outputXmlPath, val);
				try
				{
					while (reader.Read())
					{
						if ((int)reader.NodeType == 1 && reader.Name == text4)
						{
							val2.WriteStartElement(reader.Name);
							continue;
						}
						if ((int)reader.NodeType == 15 && reader.Name == text4)
						{
							val2.WriteEndElement();
							break;
						}
						if ((int)reader.NodeType == 1 && tableFromSInfo.Any((KeyValuePair<string, string[]> s) => s.Key == reader.Name))
						{
							string sectionName = reader.Name;
							string[] array = tableFromSInfo[sectionName];
							string rootName = array[0];
							string rowName = array[2];
							HTKKCMCTableInfo hTKKCMCTableInfo = sInfo.Tables.Find((HTKKCMCTableInfo x) => x.TablePath.Contains(sectionName) && x.TablePath.Contains(rootName) && x.TablePath.Contains(rowName));
							val2.WriteStartElement(sectionName);
							if (hTKKCMCTableInfo != null)
							{
								List<string> tagsFromXml = getTagsFromXml(text3, rootName, sectionName, rowName);
								foreach (HTKKCMCRowValue rowsValue in hTKKCMCTableInfo.RowsValues)
								{
									val2.WriteStartElement(rowName);
									foreach (string tag in tagsFromXml)
									{
										List<HTKKCMCTableCellInfo> cells = hTKKCMCTableInfo.RowsInfo.Cells;
										Predicate<HTKKCMCTableCellInfo> match = (HTKKCMCTableCellInfo c) => c.Path == tag;
										int num = cells.FindIndex(match);
										if (num >= 0)
										{
											HTKKCMCTableCellInfo hTKKCMCTableCellInfo = hTKKCMCTableInfo.RowsInfo.Cells[num];
											if (hTKKCMCTableCellInfo.Controltype == "16" && string.IsNullOrEmpty(rowsValue.RowValue[num]))
											{
												rowsValue.RowValue[num] = "0";
											}
											if (hTKKCMCTableCellInfo.Controltype == "2")
											{
												rowsValue.RowValue[num] = ((rowsValue.RowValue[num] == "1" || rowsValue.RowValue[num].ToLower() == "x") ? "true" : "false");
											}
											val2.WriteElementString(tag, rowsValue.RowValue[num]);
										}
									}
									val2.WriteEndElement();
								}
							}
							val2.WriteEndElement();
							reader.Skip();
						}
						else if ((int)reader.NodeType == 1)
						{
							val2.WriteStartElement(reader.Name);
							if (reader.HasAttributes)
							{
								while (reader.MoveToNextAttribute())
								{
									val2.WriteAttributeString(reader.Name, reader.Value);
								}
								reader.MoveToElement();
							}
							if (reader.IsEmptyElement)
							{
								val2.WriteEndElement();
							}
						}
						else if ((int)reader.NodeType == 3)
						{
							val2.WriteString(reader.Value);
						}
						else if ((int)reader.NodeType == 15)
						{
							val2.WriteEndElement();
						}
					}
				}
				finally
				{
					((IDisposable)val2)?.Dispose();
				}
			}
			finally
			{
				if (reader != null)
				{
					((IDisposable)reader).Dispose();
				}
			}
			return text;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return text + ex.Message;
		}
	}

	public static List<string> getTagsFromXml(string templateXmlPath, string root, string key, string childen)
	{
		//IL_0008: Unknown result type (might be due to invalid IL or missing references)
		//IL_000e: Expected O, but got Unknown
		//IL_0056: Unknown result type (might be due to invalid IL or missing references)
		//IL_005d: Expected O, but got Unknown
		//IL_0076: Unknown result type (might be due to invalid IL or missing references)
		//IL_007d: Expected O, but got Unknown
		List<string> list = new List<string>();
		try
		{
			XmlDocument val = new XmlDocument();
			val.Load(templateXmlPath);
			XmlNode val2 = ((XmlNode)val).SelectSingleNode("//" + root);
			XmlNode val3 = val2.SelectSingleNode(key);
			if (val3 != null)
			{
				foreach (XmlNode item in val3.SelectNodes(childen))
				{
					XmlNode val4 = item;
					foreach (XmlNode childNode in val4.ChildNodes)
					{
						XmlNode val5 = childNode;
						list.Add(val5.Name);
					}
				}
			}
			return list;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return new List<string>();
		}
	}

	public static string setStaticXml(string outputXmlPath, HTKKCMCSheetInfo sInfo)
	{
		//IL_0008: Unknown result type (might be due to invalid IL or missing references)
		//IL_000e: Expected O, but got Unknown
		try
		{
			string result = "";
			XmlDocument val = new XmlDocument();
			val.Load(outputXmlPath);
			foreach (HTKKCMCCellInfo cell in sInfo.StaticCells)
			{
				XmlNode val2 = ((XmlNode)val).SelectSingleNode(cell.Path);
				if (val2 == null)
				{
					continue;
				}
				HTKKCMCCellInfo hTKKCMCCellInfo = sInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Path == cell.Path);
				if (hTKKCMCCellInfo != null)
				{
					if (hTKKCMCCellInfo.Controltype == "16" && string.IsNullOrEmpty(hTKKCMCCellInfo.Value))
					{
						hTKKCMCCellInfo.Value = "0";
					}
					if (cell.Controltype == "2")
					{
						hTKKCMCCellInfo.Value = ((hTKKCMCCellInfo.Value == "1" || hTKKCMCCellInfo.Value == "x" || hTKKCMCCellInfo.Value == "true") ? "True" : "False");
					}
					val2.InnerText = hTKKCMCCellInfo.Value;
				}
			}
			if (File.Exists(outputXmlPath))
			{
				File.SetAttributes(outputXmlPath, FileAttributes.Normal);
			}
			val.Save(outputXmlPath);
			return result;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return ex.Message;
		}
	}

	public static string ghepPhuLucVaoToKhai(string pathXmlToKhai, string pathXmlPhuLuc, string pathOutput)
	{
		try
		{
			using (StreamReader streamReader = new StreamReader(pathXmlToKhai))
			{
				using StreamWriter streamWriter = new StreamWriter(pathOutput);
				string text;
				while ((text = streamReader.ReadLine()) != null)
				{
					if (text.Trim().StartsWith("<PLuc>"))
					{
						streamWriter.WriteLine("<PLuc>");
						using StreamReader streamReader2 = new StreamReader(pathXmlPhuLuc);
						string text2;
						while ((text2 = streamReader2.ReadLine()) != null)
						{
							if (!text2.Trim().StartsWith("<?xml"))
							{
								streamWriter.WriteLine(text2);
							}
						}
					}
					else
					{
						streamWriter.WriteLine(text);
					}
				}
			}
			return "";
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return ex.Message;
		}
	}

	private static void RemoveAllNamespaces(XmlNode node)
	{
		//IL_0033: Unknown result type (might be due to invalid IL or missing references)
		//IL_0039: Expected O, but got Unknown
		XmlElement val = (XmlElement)(object)((node is XmlElement) ? node : null);
		if (val != null)
		{
			val.RemoveAttribute("xmlns");
		}
		foreach (XmlNode childNode in node.ChildNodes)
		{
			XmlNode node2 = childNode;
			RemoveAllNamespaces(node2);
		}
	}

	public static string getStatic(string xmlFilePath, string rootNodeName, string childNodeOfRoot)
	{
		//IL_0001: Unknown result type (might be due to invalid IL or missing references)
		//IL_0007: Expected O, but got Unknown
		//IL_000f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0015: Expected O, but got Unknown
		//IL_0046: Unknown result type (might be due to invalid IL or missing references)
		//IL_004d: Expected O, but got Unknown
		XmlDocument val = new XmlDocument();
		val.Load(xmlFilePath);
		XmlDocument val2 = new XmlDocument();
		XmlElement val3 = val2.CreateElement(rootNodeName);
		((XmlNode)val2).AppendChild((XmlNode)(object)val3);
		XmlNode documentElement = (XmlNode)(object)val.DocumentElement;
		bool flag = false;
		foreach (XmlNode childNode in documentElement.ChildNodes)
		{
			XmlNode val4 = childNode;
			if (flag)
			{
				XmlNode val5 = val2.ImportNode(val4, true);
				((XmlNode)val3).AppendChild(val5);
			}
			if (val4.Name == childNodeOfRoot)
			{
				flag = true;
			}
		}
		string innerXml = ((XmlNode)val2).InnerXml;
		innerXml = innerXml.Replace("<" + rootNodeName + ">", "");
		innerXml = innerXml.Replace("</" + rootNodeName + ">", "");
		innerXml = innerXml.Replace(" ", "");
		return innerXml.Replace("\t", "");
	}

	public static void exportFooterOfTable(string xmlFilePath, HTKKCMCSheetInfo sInfo)
	{
		//IL_0001: Unknown result type (might be due to invalid IL or missing references)
		//IL_0007: Expected O, but got Unknown
		//IL_0021: Unknown result type (might be due to invalid IL or missing references)
		//IL_0027: Expected O, but got Unknown
		XmlDocument val = new XmlDocument();
		val.Load(xmlFilePath);
		string namespaceURI = ((XmlNode)val.DocumentElement).NamespaceURI;
		XmlNamespaceManager val2 = new XmlNamespaceManager(val.NameTable);
		val2.AddNamespace("ns", namespaceURI);
		for (int i = 0; i < sInfo.StaticCells.Count; i++)
		{
			if (sInfo.StaticCells[i].Path == null || sInfo.StaticCells[i].Path.Length <= 0 || sInfo.StaticCells[i].CellTemplate == null)
			{
				continue;
			}
			string path = sInfo.StaticCells[i].Path;
			string[] source = path.Split(new char[1] { '/' });
			string text = GeneralFunctions.ListToString(source.ToList(), "/ns:");
			XmlNode val3 = ((XmlNode)val).SelectSingleNode("//ns:" + text, val2);
			if (val3 != null)
			{
				string value = sInfo.StaticCells[i].Value;
				if (!string.IsNullOrEmpty(value))
				{
					val3.InnerText = value;
				}
			}
		}
		val.Save(xmlFilePath);
	}

	public static XmlDocument exportToKhaiVaThongTinChung(XmlDocument xmlExportContent)
	{
		XmlNode val = ((XmlNode)xmlExportContent).SelectSingleNode("HSoThueDTu");
		XmlAttribute val2 = xmlExportContent.CreateAttribute("xmlns:xsi");
		((XmlNode)val2).Value = "http://www.w3.org/2001/XMLSchema-instance";
		((XmlNamedNodeMap)val.Attributes).SetNamedItem((XmlNode)(object)val2);
		val2 = xmlExportContent.CreateAttribute("xmlns");
		((XmlNode)val2).Value = "http://kekhaithue.gdt.gov.vn/TKhaiThue";
		((XmlNamedNodeMap)val.Attributes).SetNamedItem((XmlNode)(object)val2);
		if (HTKKFormater.isNullString(WinSettings.HTKKHeaderInfo.mstDLyThue))
		{
			XmlNode parentNode = ((XmlNode)xmlExportContent).SelectSingleNode("HSoThueDTu/HSoKhaiThue/TTinChung/TTinTKhaiThue/DLyThue").ParentNode;
			XmlNode val3 = parentNode.SelectSingleNode("DLyThue");
			if (val3 != null)
			{
				XmlNode val4 = val3.SelectSingleNode("mstDLyThue");
				XmlNode val5 = val4.OwnerDocument.CreateNode((XmlNodeType)2, "xsi:nil", "http://www.w3.org/2001/XMLSchema-instance");
				val5.Value = "true";
				if (val4.InnerText == "")
				{
					((XmlNamedNodeMap)val4.Attributes).SetNamedItem(val5);
				}
				val4 = val3.SelectSingleNode("ngayKyHDDLyThue");
				val5 = val4.OwnerDocument.CreateNode((XmlNodeType)2, "xsi:nil", "http://www.w3.org/2001/XMLSchema-instance");
				val5.Value = "true";
				if (val4.InnerText == "")
				{
					((XmlNamedNodeMap)val4.Attributes).SetNamedItem(val5);
				}
			}
			if (val3 != null)
			{
				parentNode.RemoveChild(val3);
			}
		}
		XmlNode documentElement = (XmlNode)(object)xmlExportContent.DocumentElement;
		xmlExportTTinChung(xmlExportContent);
		ExportHeaderXML((XmlNode)(object)xmlExportContent, 0);
		ExportXML(documentElement, 0);
		return xmlExportContent;
	}

	public static string WinDoWorkSave_Large1(FpSpread winFpSpread, List<HTKKCMCSheetInfo> sheetInfos)
	{
		//IL_0397: Unknown result type (might be due to invalid IL or missing references)
		//IL_039d: Expected O, but got Unknown
		//IL_03a5: Unknown result type (might be due to invalid IL or missing references)
		//IL_03ac: Expected O, but got Unknown
		string text = "";
		try
		{
			for (int i = 0; i < sheetInfos.Count; i++)
			{
				HTKKCMCSheetInfo hTKKCMCSheetInfo = sheetInfos[i];
				if (hTKKCMCSheetInfo == null)
				{
					continue;
				}
				if (hTKKCMCSheetInfo.TaiDuLieuRatLon != "1")
				{
					WinDoWorkSheetSave(winFpSpread.Sheets[i], WinSettings.ToKhaiInfo.SheetInfos[i]);
					continue;
				}
				string sFolder = HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst + "\\";
				string fileNameBySheet = GetFileNameBySheet(WinSettings.TTPeriodInfo, hTKKCMCSheetInfo.DataFile, sFolder, WinSettings.ToKhaiInfo.VersionXML);
				if (File.Exists(fileNameBySheet))
				{
					File.Delete(fileNameBySheet);
				}
				string text2 = hTKKCMCSheetInfo.DataFile + "_" + WinSettings.ToKhaiInfo.VersionXML.Replace(".", "") + ".xml";
				string filePath = WinSettings.ToKhaiInfo.InterfaceTemplateXml + text2;
				saveStaticCell(hTKKCMCSheetInfo, filePath, fileNameBySheet);
				XDocument val = XDocument.Load(fileNameBySheet);
				List<XElement> source = ((XContainer)val).Descendants(XName.op_Implicit("Section")).ToList();
				foreach (HTKKCMCTableInfo table in hTKKCMCSheetInfo.Tables)
				{
					Func<XElement, bool> predicate = (XElement s) => (string)s.Attribute(XName.op_Implicit("TablePath")) == table.TablePath;
					XElement val2 = source.FirstOrDefault(predicate);
					if (val2 == null)
					{
						text += "Không tìm thấy Section cần chèn!";
						return text;
					}
					XElement val3 = ((XContainer)val2).Element(XName.op_Implicit("Location"));
					XElement val4 = ((XContainer)val3).Element(XName.op_Implicit("TableTemplateRowStart"));
					if (val4 != null)
					{
						val4.Value = table.Location.TableTemplateRowStart.ToString();
					}
					XElement val5 = ((XContainer)val3).Element(XName.op_Implicit("TableReportRowStart"));
					if (val5 != null)
					{
						val5.Value = table.Location.TableReportRowStart.ToString();
					}
					XElement val6 = ((XContainer)val3).Element(XName.op_Implicit("DataLocationTemplateRowStart"));
					if (val6 != null)
					{
						val6.Value = table.Location.DataLocationTemplateRowStart.ToString();
					}
					XElement val7 = ((XContainer)val3).Element(XName.op_Implicit("DataLocationTaiBangKeRowStart"));
					if (val7 != null)
					{
						val7.Value = table.Location.DataLocationTaiBangKeRowStart.ToString();
					}
					Extensions.Remove<XElement>(((XContainer)val2).Elements(XName.op_Implicit("RowData")));
					try
					{
						HTKKCMCRowInfo rowsInfo = table.RowsInfo;
						List<HTKKCMCRowValue> rowsValues = table.RowsValues;
						int num = 1;
						foreach (HTKKCMCRowValue item in rowsValues)
						{
							try
							{
								string text3 = "";
								for (int num2 = 0; num2 < rowsInfo.Cells.Count; num2++)
								{
									text3 = text3 + item.RowValue[num2] + "###";
								}
								if (text3.Length > 0)
								{
									text3 = text3.Substring(0, text3.Length - "###".Length);
									XElement val8 = new XElement(XName.op_Implicit("RowData"), new object[2]
									{
										(object)new XAttribute(XName.op_Implicit("RowDisplay"), (object)"1"),
										text3
									});
									((XContainer)val2).Add((object)val8);
								}
							}
							catch (Exception ex)
							{
								HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
								text += ex.Message;
							}
							num++;
						}
					}
					catch (Exception ex)
					{
						HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
						text += ex.Message;
					}
				}
				val.Save(fileNameBySheet);
			}
			return text;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return text + ex.Message;
		}
	}

	public static void saveStaticCell(HTKKCMCSheetInfo sInfo, string filePath, string outputXmlPath)
	{
		XDocument val = XDocument.Load(filePath);
		List<XElement> source = (from section in ((XContainer)val).Descendants(XName.op_Implicit("Section"))
			where (string)section.Attribute(XName.op_Implicit("Dynamic")) == "0"
			select section).SelectMany((XElement section) => ((XContainer)section).Descendants(XName.op_Implicit("Cell"))).ToList();
		foreach (HTKKCMCCellInfo cell in sInfo.StaticCells)
		{
			Func<XElement, bool> predicate = (XElement xmlCell) => (string)xmlCell.Attribute(XName.op_Implicit("Path")) == cell.Path;
			XElement val2 = source.FirstOrDefault(predicate);
			if (val2 != null)
			{
				if (val2.Attribute(XName.op_Implicit("CellID")) != null && cell.CellTemplate != null)
				{
					ReplaceRowAttrStatic(val2, "CellID", cell.CellTemplate.RowIndex);
				}
				if (val2.Attribute(XName.op_Implicit("CellID2")) != null && cell.CellReport != null)
				{
					ReplaceRowAttrStatic(val2, "CellID2", cell.CellReport.RowIndex);
				}
				if (val2.Attribute(XName.op_Implicit("Value")) != null)
				{
					val2.SetAttributeValue(XName.op_Implicit("Value"), (object)(cell.Value ?? string.Empty));
				}
			}
		}
		val.Save(outputXmlPath);
	}

	public static XElement ReplaceRowAttrStatic(XElement targetCell, string attr, int row)
	{
		if (targetCell.Attribute(XName.op_Implicit(attr)) != null)
		{
			string text = (string)targetCell.Attribute(XName.op_Implicit(attr));
			string[] array = text.Split(new char[1] { '_' });
			string text2 = array[0] + "_" + row;
			targetCell.SetAttributeValue(XName.op_Implicit(attr), (object)text2);
		}
		return targetCell;
	}

	public static string checkTaiDuLieuLon(string id, string getNameAttr)
	{
		string text = "Menu.xml";
		XDocument val = XDocument.Load(text);
		XElement val2 = ((XContainer)val).Descendants(XName.op_Implicit("Menu")).FirstOrDefault((Func<XElement, bool>)((XElement m) => (string)m.Attribute(XName.op_Implicit("ID")) == id));
		if (val2 != null)
		{
			List<XElement> list = ((XContainer)val2).Descendants(XName.op_Implicit("Sheet")).ToList();
			for (int num = 0; num < list.Count; num++)
			{
				string text2 = (string)list[num].Attribute(XName.op_Implicit(getNameAttr));
				if (!string.IsNullOrEmpty(text2) && text2 == "1")
				{
					return text2;
				}
			}
		}
		return "";
	}

	public static void UpadatePage(FpSpread winFpSpread, int sheetIndex, int page)
	{
		try
		{
			MethodInfo method = GetMethod("UpadatePage_" + sheetIndex);
			if ((object)method == null)
			{
				return;
			}
			ParameterInfo[] parameters = method.GetParameters();
			object[] array = new object[parameters.Length];
			for (int i = 0; i < parameters.Length; i++)
			{
				switch (parameters[i].Name)
				{
				case "winFpSpread":
					array[i] = winFpSpread;
					break;
				case "ToKhaiInfo":
					array[i] = WinSettings.ToKhaiInfo;
					break;
				case "TTinChung":
					array[i] = WinSettings.TTinChung;
					break;
				case "PeriodInfo":
					array[i] = WinSettings.TTPeriodInfo;
					break;
				case "HeaderInfo":
					array[i] = WinSettings.HTKKHeaderInfo;
					break;
				case "DicSheet":
					array[i] = WinSettings.DicSheet;
					break;
				case "sView":
					array[i] = winFpSpread.Sheets[sheetIndex];
					break;
				case "page":
					array[i] = page;
					break;
				}
			}
			method.Invoke(WinSettings.InstanceClass, array);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void SheetTabChange(FpSpread winFpSpread, int sheetIndex)
	{
		try
		{
			MethodInfo method = GetMethod("SheetTabChange_" + sheetIndex);
			if ((object)method == null)
			{
				return;
			}
			ParameterInfo[] parameters = method.GetParameters();
			object[] array = new object[parameters.Length];
			for (int i = 0; i < parameters.Length; i++)
			{
				switch (parameters[i].Name)
				{
				case "winFpSpread":
					array[i] = winFpSpread;
					break;
				case "ToKhaiInfo":
					array[i] = WinSettings.ToKhaiInfo;
					break;
				case "TTinChung":
					array[i] = WinSettings.TTinChung;
					break;
				case "PeriodInfo":
					array[i] = WinSettings.TTPeriodInfo;
					break;
				case "HeaderInfo":
					array[i] = WinSettings.HTKKHeaderInfo;
					break;
				case "DicSheet":
					array[i] = WinSettings.DicSheet;
					break;
				case "sView":
					array[i] = winFpSpread.Sheets[sheetIndex];
					break;
				case "sheetIndex":
					array[i] = sheetIndex;
					break;
				}
			}
			method.Invoke(WinSettings.InstanceClass, array);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private static void InsertAfterLastNodeByXPath(XmlNode parent, XmlNode newNode, string nodeName)
	{
		if (parent == null || newNode == null)
		{
			return;
		}
		XmlNode val = parent.SelectSingleNode(nodeName + "[last()]");
		if (val != null)
		{
			if (val.NextSibling != null)
			{
				parent.InsertBefore(newNode, val.NextSibling);
			}
			else
			{
				parent.AppendChild(newNode);
			}
		}
		else
		{
			parent.AppendChild(newNode);
		}
	}

	public static List<HTKKCMCSheetInfo> WinDoWorkImportXmlToSheetInfos(string xmlFilePath)
	{
		//IL_0099: Unknown result type (might be due to invalid IL or missing references)
		try
		{
			if (string.IsNullOrEmpty(xmlFilePath) || xmlFilePath.Trim().Length == 0)
			{
				throw new ArgumentException("Duong dan file XML khong duoc de trong.", "xmlFilePath");
			}
			if (!File.Exists(xmlFilePath))
			{
				throw new FileNotFoundException("Khong tim thay file XML.", xmlFilePath);
			}
			if (WinSettings.ToKhaiInfo == null || WinSettings.ToKhaiInfo.SheetInfos == null)
			{
				throw new InvalidOperationException("SheetInfos chua duoc khoi tao.");
			}
			XmlDocument val = LoadXmlWithoutNamespaces(xmlFilePath);
			XmlNode documentElement = (XmlNode)(object)val.DocumentElement;
			if (documentElement == null)
			{
				throw new XmlException("File XML khong co node goc.");
			}
			StartImportXML(documentElement, 0);
			List<HTKKCMCSheetInfo> sheetInfos = WinSettings.ToKhaiInfo.SheetInfos;
			foreach (HTKKCMCSheetInfo item in sheetInfos)
			{
				if (item != null)
				{
					XmlNode importSheetNode = GetImportSheetNode(documentElement, item);
					if (importSheetNode != null)
					{
						ProcessXMLNode(item, importSheetNode, checkKHBS: false);
						ProcessSelectedCell(item);
						ProcessParentCell(item);
						item.IsReadDataFile = true;
						item.CoDuLieu = true;
					}
				}
			}
			return sheetInfos;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	private static XmlDocument LoadXmlWithoutNamespaces(string xmlFilePath)
	{
		//IL_0002: Unknown result type (might be due to invalid IL or missing references)
		//IL_0008: Expected O, but got Unknown
		//IL_0060: Unknown result type (might be due to invalid IL or missing references)
		//IL_0066: Expected O, but got Unknown
		//IL_004e: Unknown result type (might be due to invalid IL or missing references)
		XmlTextReader val = new XmlTextReader(xmlFilePath);
		XDocument val2;
		try
		{
			val.ProhibitDtd = true;
			val.XmlResolver = null;
			val2 = XDocument.Load((XmlReader)(object)val, (LoadOptions)1);
		}
		finally
		{
			((IDisposable)val)?.Dispose();
		}
		if (val2.Root == null)
		{
			throw new XmlException("File XML khong co node goc.");
		}
		XElement val3 = RemoveXmlNamespaces(val2.Root);
		XmlDocument val4 = new XmlDocument();
		val4.XmlResolver = null;
		val4.LoadXml(((XNode)val3).ToString((SaveOptions)1));
		return val4;
	}

	private static XElement RemoveXmlNamespaces(XElement element)
	{
		//IL_0097: Unknown result type (might be due to invalid IL or missing references)
		//IL_009d: Expected O, but got Unknown
		return new XElement(XName.op_Implicit(element.Name.LocalName), new object[2]
		{
			(from attribute in element.Attributes()
				where !attribute.IsNamespaceDeclaration
				select attribute).Select((Func<XAttribute, XAttribute>)((XAttribute attribute) => new XAttribute(XName.op_Implicit(attribute.Name.LocalName), (object)attribute.Value))),
			from node in ((XContainer)element).Nodes()
				select (XNode)((node is XElement) ? ((object)RemoveXmlNamespaces((XElement)node)) : ((object)node))
		});
	}

	private static XmlNode GetImportSheetNode(XmlNode xmlRoot, HTKKCMCSheetInfo sheetInfo)
	{
		//IL_01d1: Unknown result type (might be due to invalid IL or missing references)
		//IL_01d8: Expected O, but got Unknown
		List<string> list = new List<string>();
		if (sheetInfo.StaticCells != null)
		{
			list.AddRange(from cell in sheetInfo.StaticCells
				where cell != null && !string.IsNullOrEmpty(cell.Path) && cell.Path.Trim().Length > 0 && cell.Path.Contains("/")
				select cell.Path);
		}
		if (sheetInfo.Tables != null)
		{
			list.AddRange(from table in sheetInfo.Tables
				where table != null && !string.IsNullOrEmpty(table.TablePath) && table.TablePath.Trim().Length > 0
				select table.TablePath);
		}
		if (!string.IsNullOrEmpty(sheetInfo.XMLnode) && sheetInfo.XMLnode.Trim().Length > 0)
		{
			list.Add(sheetInfo.XMLnode);
		}
		if (list.Count == 0)
		{
			return xmlRoot;
		}
		XmlNode result = null;
		int num = -1;
		IEnumerable<string> enumerable = (from name in list.Select(GetFirstXmlPathNode)
			where !string.IsNullOrEmpty(name)
			select name).Distinct();
		foreach (string item in enumerable)
		{
			List<XmlNode> list2 = new List<XmlNode>();
			if (xmlRoot.Name == item)
			{
				list2.Add(xmlRoot);
			}
			XmlNodeList val = xmlRoot.SelectNodes(".//*");
			if (val != null)
			{
				foreach (XmlNode item2 in val)
				{
					XmlNode val2 = item2;
					if (val2.Name == item)
					{
						list2.Add(val2);
					}
				}
			}
			foreach (XmlNode item3 in list2)
			{
				int num2 = 0;
				foreach (string item4 in list)
				{
					if (!(GetFirstXmlPathNode(item4) != item3.Name))
					{
						string text = RemoveFirstXmlPathNode(item4);
						if (text.Length == 0 || item3.SelectSingleNode(text) != null)
						{
							num2++;
						}
					}
				}
				if (num2 > num)
				{
					num = num2;
					result = item3;
				}
			}
		}
		return result;
	}

	private static string GetFirstXmlPathNode(string xmlPath)
	{
		if (string.IsNullOrEmpty(xmlPath))
		{
			return string.Empty;
		}
		string[] array = xmlPath.Trim(new char[1] { '/' }).Split(new char[1] { '/' }, StringSplitOptions.RemoveEmptyEntries);
		return (array.Length == 0) ? string.Empty : array[0];
	}

	private static string RemoveFirstXmlPathNode(string xmlPath)
	{
		if (string.IsNullOrEmpty(xmlPath))
		{
			return string.Empty;
		}
		string text = xmlPath.Trim(new char[1] { '/' });
		int num = text.IndexOf('/');
		return (num < 0) ? string.Empty : text.Substring(num + 1);
	}

	public static void removeNilAttrIfValueOk(string inputPath, string outputPath)
	{
		XDocument val = XDocument.Load(inputPath);
		XNamespace val2 = XNamespace.op_Implicit("http://www.w3.org/2001/XMLSchema-instance");
		foreach (XElement item in ((XContainer)val).Descendants())
		{
			XAttribute val3 = item.Attribute(val2 + "nil");
			if (val3 != null && val3.Value == "true" && !string.IsNullOrEmpty(item.Value))
			{
				val3.Remove();
			}
		}
		val.Save(outputPath);
	}
}
