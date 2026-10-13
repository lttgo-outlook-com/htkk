using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Security;
using System.Reflection;
using System.Security.Cryptography.X509Certificates;
using System.Security.Principal;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;
using FileTransfer;
using Foundation;
using HTKK.Properties;
using Ionic.Zip;
using TAX.HTKK.Business;
using TAX.HTKK.Utilities.Control;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Implement;

namespace HTKK.Forms;

public class frmFormStart : Form
{
	private BackgroundWorker bw = new BackgroundWorker();

	private string lastest_version_code;

	private DownloadFileInfo[] down_file_info_list;

	private CurrVersionMgt version_mgt = null;

	private string rootPath = "";

	private string rootUpdate = "";

	private string admin_serv_uri = HTKKConfig.URLUPDATE;

	private string isProxy = string.Empty;

	private string isUser = string.Empty;

	private string userName = string.Empty;

	private string pass = string.Empty;

	private string appver = string.Empty;

	private string timeExcute = string.Empty;

	private string proxy = string.Empty;

	private string port = string.Empty;

	private string tenCQTCapCuc = string.Empty;

	private string maCQTCapCuc = string.Empty;

	private string tenCQTNoiNop = string.Empty;

	private string maCQTNoiNop = string.Empty;

	private List<string> lstMaSoThue;

	private IContainer components = null;

	private Panel panel_login;

	private Panel panel1;

	private Label label1;

	private PictureBox pictureBox1;

	private ComboBox cboMaSoThue;

	private Label label2;

	private Button btnThoat;

	private Button btnDongY;

	private Button btnMaSoMoi;

	private Button btnXoa;

	private Label lblNguoiNopThue;

	private Panel panel_tenNNT;

	private Label lblTitle_Version_bg;

	private Label label3;

	public static bool InVisualStudio()
	{
		Process[] processes = Process.GetProcesses();
		Process[] array = processes;
		foreach (Process process in array)
		{
			if (process.ProcessName.Equals("devenv"))
			{
				return true;
			}
		}
		return false;
	}

	internal static void InstalarFuente(string NombreFnt, string RutaFnt)
	{
		try
		{
			string comando = $"copy /Y \"{RutaFnt}\" \"%WINDIR%\\Fonts\" ";
			HTKKFile.EjecutarCMD(comando);
			FileInfo fileInfo = new FileInfo(RutaFnt);
			comando = $"reg add \"HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Fonts\" /v \"{NombreFnt}\" /t REG_SZ /d {fileInfo.Name} /f";
			HTKKFile.EjecutarCMD(comando);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public frmFormStart()
	{
		//IL_00c7: Unknown result type (might be due to invalid IL or missing references)
		//IL_00cd: Expected O, but got Unknown
		//IL_01b4: Unknown result type (might be due to invalid IL or missing references)
		//IL_01bb: Expected O, but got Unknown
		//IL_03a7: Unknown result type (might be due to invalid IL or missing references)
		//IL_03ae: Expected O, but got Unknown
		//IL_0200: Unknown result type (might be due to invalid IL or missing references)
		//IL_0206: Invalid comparison between Unknown and I4
		try
		{
			XmlDocument val = new XmlDocument();
			if (File.Exists("Config.xml"))
			{
				val.LoadXml(HTKKFile.LoadFile("Config.xml"));
				XmlNode val2 = ((XmlNode)val).SelectNodes("/Root/RegisterTBarCode5")[0];
				if (val2 != null && !ConvertHelpers.ConvertToBool(val2.InnerText))
				{
					WindowsIdentity current = WindowsIdentity.GetCurrent();
					WindowsPrincipal windowsPrincipal = new WindowsPrincipal(current);
					if (windowsPrincipal.IsInRole(WindowsBuiltInRole.Administrator))
					{
						try
						{
							string text = HTKKConfig.APPPATH + "\\VNARIAL.TTF";
							if (File.Exists(text))
							{
								InstalarFuente("VNARIAL.GV", text);
							}
						}
						catch (Exception ex)
						{
							HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
						}
						HTKKFile.RegisterDll("TBarCode5.ocx");
						val2.InnerText = "True";
						XmlTextWriter val3 = new XmlTextWriter("Config.xml", Encoding.UTF8);
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
					else if ((int)MessageBox.Show("Ứng dụng yêu cầu quyền Administrator cho lần chạy đầu tiên để hoàn tất cài đặt, chọn Yes để chạy ứng dụng với quyền Administrator.", "Cảnh báo", (MessageBoxButtons)4, (MessageBoxIcon)48, (MessageBoxDefaultButton)0, (MessageBoxOptions)131072) == 6)
					{
						ProcessStartInfo startInfo = new ProcessStartInfo
						{
							UseShellExecute = true,
							WorkingDirectory = Environment.CurrentDirectory,
							FileName = Application.ExecutablePath,
							Arguments = "/force",
							Verb = "runas"
						};
						try
						{
							Process.Start(startInfo);
							Environment.Exit(1);
						}
						catch
						{
							return;
						}
					}
					else
					{
						Environment.Exit(1);
					}
				}
			}
			InitializeComponent();
			GetDataMaSoThue();
			rootPath = HTKKConfig.APPPATH.Substring(0, HTKKConfig.APPPATH.LastIndexOf("\\"));
			rootUpdate = rootPath + "\\AutoUpdate";
			string path = rootUpdate + "\\AppSchedulerCf.xml";
			XmlNode val4 = ((XmlNode)val).SelectNodes("/Root/NewInstall")[0];
			if ((val4 != null && ConvertHelpers.ConvertToBool(val4.InnerText)) || !File.Exists(path))
			{
				File.WriteAllBytes(Path.Combine(Path.GetTempPath(), "AutoUpdate.zip"), Resources.AutoUpdate);
				if (File.Exists(Path.Combine(Path.GetTempPath(), "AutoUpdate.zip")))
				{
					using ZipFile zipFile = ZipFile.Read(Path.Combine(Path.GetTempPath(), "AutoUpdate.zip"));
					zipFile.ExtractAll(rootPath, ExtractExistingFileAction.OverwriteSilently);
					HTKKFile.Log("Ghi de thu muc AutoUpdate thanh cong.");
				}
				val4.InnerText = "False";
				XmlTextWriter val3 = new XmlTextWriter("Config.xml", Encoding.UTF8);
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
			version_mgt = new CurrVersionMgt(rootPath, rootPath + "\\version");
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void frmFormStart_Load(object sender, EventArgs e)
	{
		//IL_0135: Unknown result type (might be due to invalid IL or missing references)
		//IL_013f: Expected O, but got Unknown
		//IL_0164: Unknown result type (might be due to invalid IL or missing references)
		//IL_018f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0195: Expected O, but got Unknown
		try
		{
			if (WinSettings.chuyenMSTKhac)
			{
				((Control)btnThoat).Text = "Quay lại";
				if (!HTKKFormater.isNullString(WinSettings.HTKKHeaderInfo.mst) && WinSettings.HTKKHeaderInfo.mst.Trim() == ((Control)cboMaSoThue).Text.Trim())
				{
					((Control)btnXoa).Enabled = false;
				}
				else
				{
					((Control)btnXoa).Enabled = true;
				}
			}
			else
			{
				((Control)btnThoat).Text = "Thoát";
				((Control)btnXoa).Enabled = true;
			}
			((Control)panel_login).Left = (((Form)this).ClientSize.Width - ((Control)panel_login).Width) / 2;
			((Control)panel_login).Top = (((Form)this).ClientSize.Height - ((Control)panel_login).Height) / 2;
			WinSettings.VersionApp = GetVersionApp();
			((Control)this).Text = "Hỗ trợ kê khai thuế -  Phiên bản " + WinSettings.VersionApp;
			((Control)lblTitle_Version_bg).Text = "Phiên bản " + WinSettings.VersionApp;
			((Form)this).FormClosing += new FormClosingEventHandler(frmFormStart_FormClosing);
			int num = (((Control)panel_tenNNT).Width - ((Control)lblNguoiNopThue).Width) / 2;
			((Control)panel_tenNNT).Padding = new Padding(num, 3, 0, 2);
			if (File.Exists(rootPath + "\\Rollback\\Rollback.xml"))
			{
				XmlDocument val = new XmlDocument();
				val.Load(rootPath + "\\Rollback\\Rollback.xml");
				if (((XmlNode)val).SelectNodes("/Data/Rollback/Status").Count > 0 && ((XmlNode)val).SelectNodes("/Data/Rollback/Status")[0].InnerText != "1")
				{
					Process.Start(rootUpdate + "\\Client_VersionUpdate.exe");
					Environment.Exit(0);
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private string GetVersionApp()
	{
		//IL_0001: Unknown result type (might be due to invalid IL or missing references)
		//IL_0007: Expected O, but got Unknown
		XmlDocument val = new XmlDocument();
		string text = HTKKConfig.APPPATH.Substring(0, HTKKConfig.APPPATH.LastIndexOf("\\"));
		val.Load(text + "\\AutoUpdate\\AppSchedulerCf.xml");
		return ((XmlNode)val).SelectNodes("/AppSchedule/Task/AppVersion")[0].InnerText;
	}

	private void frmFormStart_FormClosing(object sender, FormClosingEventArgs e)
	{
		Application.Exit();
	}

	private void frmFormStart_Resize(object sender, EventArgs e)
	{
		((Control)panel_login).Left = (((Form)this).ClientSize.Width - ((Control)panel_login).Width) / 2;
		((Control)panel_login).Top = (((Form)this).ClientSize.Height - ((Control)panel_login).Height) / 2;
		FormProcess.SetConfigForProgcessBar((Control)(object)this);
	}

	private void GetDataMaSoThue()
	{
		try
		{
			lstMaSoThue = new List<string>();
			string text = "..\\DataFiles";
			if (!Directory.Exists(text))
			{
				Directory.CreateDirectory(text);
			}
			string[] directories = Directory.GetDirectories(text);
			foreach (string text2 in directories)
			{
				lstMaSoThue.Add(text2.Remove(0, text.Length + 1));
			}
			cboMaSoThue.DataSource = lstMaSoThue;
			if (lstMaSoThue != null && lstMaSoThue.Count > 0)
			{
				((ListControl)cboMaSoThue).SelectedIndex = 0;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void cboMaSoThue_SelectedIndexChanged(object sender, EventArgs e)
	{
		//IL_00e1: Unknown result type (might be due to invalid IL or missing references)
		try
		{
			string path = string.Concat("..\\DataFiles\\", ((ListControl)cboMaSoThue).SelectedValue, "\\Header.xml");
			string text = "";
			if (File.Exists(path))
			{
				XmlDataDocument xmlDataDocument = new XmlDataDocument();
				FileStream fileStream = new FileStream(path, FileMode.Open, FileAccess.Read);
				((XmlDocument)xmlDataDocument).Load((Stream)fileStream);
				XmlNodeList elementsByTagName = ((XmlDocument)xmlDataDocument).GetElementsByTagName("Cell");
				XmlNode val = elementsByTagName[14];
				text = ((XmlNode)val.Attributes["Value"]).Value;
				text = text.Replace("&", "&&");
				((Control)lblNguoiNopThue).Text = text;
				WinSettings.HTKKHeaderInfo.tenNNT = ((XmlNode)val.Attributes["Value"]).Value;
				int num = (((Control)panel_login).Width - ((Control)lblNguoiNopThue).Width) / 2;
				((Control)panel_tenNNT).Padding = new Padding(num, 3, 0, 2);
				fileStream.Close();
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void btnDongY_Click(object sender, EventArgs e)
	{
		//IL_0040: Unknown result type (might be due to invalid IL or missing references)
		//IL_0089: Unknown result type (might be due to invalid IL or missing references)
		try
		{
			WinSettings.chuyenMSTKhac = false;
			if (HTKKFormater.isNullString(((Control)cboMaSoThue).Text))
			{
				frmMessage frmMessage = new frmMessage("0051", "3", "1", "Thông báo");
				((Form)frmMessage).ShowDialog();
				((Control)cboMaSoThue).Focus();
				return;
			}
			if (!ValidateBase.CheckMaSoThueCCCD(((Control)cboMaSoThue).Text))
			{
				frmMessage frmMessage = new frmMessage("0047", "3", "1", "Thông báo");
				((Form)frmMessage).ShowDialog();
				((Control)cboMaSoThue).Focus();
				return;
			}
			string path = "Menu.XML";
			((XmlNode)WinSettings.xmlDocMenu).RemoveAll();
			FileStream fileStream = new FileStream(path, FileMode.Open, FileAccess.Read);
			((XmlDocument)WinSettings.xmlDocMenu).Load((Stream)fileStream);
			if (lstMaSoThue != null && lstMaSoThue.Count > 0)
			{
				if (lstMaSoThue.Contains(((Control)cboMaSoThue).Text))
				{
					WinSettings.HTKKHeaderInfo.mst = ((Control)cboMaSoThue).Text;
					FixHeader(((Control)cboMaSoThue).Text);
					HTKKFactory.GetInstanceHTKKHeader().SetPropertyHeader(ref WinSettings.HTKKHeaderInfo, ((Control)cboMaSoThue).Text);
					if (WinSettings.HTKKHeaderInfo == null)
					{
						return;
					}
					if (Directory.Exists(rootUpdate) && !WinSettings.chuyenMSTKhac)
					{
						new Thread((ThreadStart)delegate
						{
							//IL_0018: Unknown result type (might be due to invalid IL or missing references)
							//IL_001e: Expected O, but got Unknown
							//IL_0364: Unknown result type (might be due to invalid IL or missing references)
							//IL_0305: Unknown result type (might be due to invalid IL or missing references)
							//IL_03c0: Unknown result type (might be due to invalid IL or missing references)
							//IL_03c7: Expected O, but got Unknown
							//IL_0473: Unknown result type (might be due to invalid IL or missing references)
							Thread.CurrentThread.IsBackground = true;
							Thread.Sleep(2000);
							XmlDocument val = new XmlDocument();
							string text = rootUpdate + "\\AppSchedulerCf.xml";
							if (File.Exists(text))
							{
								val.Load(text);
								if (((XmlNode)val).SelectNodes("/AppSchedule/Task/UrlAutoUpdates").Count > 0)
								{
									admin_serv_uri = TAX.HTKK.Utilities.Core.HTKKSecurity.Decrypt(((XmlNode)val).SelectNodes("/AppSchedule/Task/UrlAutoUpdates")[0].InnerText);
								}
								if (((XmlNode)val).SelectNodes("/AppSchedule/Task/UserName").Count > 0)
								{
									userName = TAX.HTKK.Utilities.Core.HTKKSecurity.Decrypt(((XmlNode)val).SelectNodes("/AppSchedule/Task/UserName")[0].InnerText);
								}
								if (((XmlNode)val).SelectNodes("/AppSchedule/Task/PassWord").Count > 0)
								{
									pass = TAX.HTKK.Utilities.Core.HTKKSecurity.Decrypt(((XmlNode)val).SelectNodes("/AppSchedule/Task/PassWord")[0].InnerText);
								}
								if (((XmlNode)val).SelectNodes("/AppSchedule/Task/Proxy").Count > 0)
								{
									proxy = TAX.HTKK.Utilities.Core.HTKKSecurity.Decrypt(((XmlNode)val).SelectNodes("/AppSchedule/Task/Proxy")[0].InnerText);
								}
								if (((XmlNode)val).SelectNodes("/AppSchedule/Task/Port").Count > 0)
								{
									port = TAX.HTKK.Utilities.Core.HTKKSecurity.Decrypt(((XmlNode)val).SelectNodes("/AppSchedule/Task/Port")[0].InnerText);
								}
								if (((XmlNode)val).SelectNodes("/AppSchedule/Task/IsProxy").Count > 0)
								{
									isProxy = ((XmlNode)val).SelectNodes("/AppSchedule/Task/IsProxy")[0].InnerText;
								}
								if (((XmlNode)val).SelectNodes("/AppSchedule/Task/IsUser").Count > 0)
								{
									isUser = ((XmlNode)val).SelectNodes("/AppSchedule/Task/IsUser")[0].InnerText;
								}
								string text2 = string.Empty;
								if (((XmlNode)val).SelectNodes("/AppSchedule/Task/AppVersion").Count > 0)
								{
									text2 = ((XmlNode)val).SelectNodes("/AppSchedule/Task/AppVersion")[0].InnerText;
								}
								if (((XmlNode)val).SelectNodes("/AppSchedule/Task/Status").Count > 0 && ((XmlNode)val).SelectNodes("/AppSchedule/Task/Status")[0].InnerText == "1")
								{
									File.WriteAllText(rootUpdate + "\\mst.dat", WinSettings.HTKKHeaderInfo.mst);
									bool flag = isNewVersion(text2);
									WinSettings.CurentVersionOnClient = text2;
									WinSettings.NewVersionOnServer = lastest_version_code;
									if (flag)
									{
										frmMessage frmMessage2 = new frmMessage("", "Đã có phiên bản " + lastest_version_code + " mới hơn phiên bản " + text2 + " hiện tại.\r\nBạn muốn cập nhật không?", "5", "2", "Thông báo");
										((Form)frmMessage2).ShowDialog();
										if (frmMessage2.Yes)
										{
											Process.Start(rootUpdate + "\\Client_VersionUpdate.exe");
											Environment.Exit(0);
										}
									}
								}
								else
								{
									frmMessage frmMessage2 = new frmMessage("", "Để chương trình hoạt động tốt.\r\nBạn vui lòng kích hoạt tính năng cập nhật phiên bản mới.\r\nBạn muốn kích hoạt không?", "5", "2", "Thông báo");
									((Form)frmMessage2).ShowDialog();
									if (frmMessage2.Yes && ((XmlNode)val).SelectNodes("/AppSchedule/Task/Status").Count > 0)
									{
										((XmlNode)val).SelectNodes("/AppSchedule/Task/Status")[0].InnerText = "1";
										XmlTextWriter val2 = new XmlTextWriter(text, Encoding.UTF8);
										try
										{
											val2.Formatting = (Formatting)1;
											val.Save((XmlWriter)(object)val2);
											bool flag = isNewVersion(text2);
											if (!string.IsNullOrEmpty(lastest_version_code))
											{
												WinSettings.CurentVersionOnClient = text2;
												WinSettings.NewVersionOnServer = lastest_version_code;
												if (flag)
												{
													frmMessage frmMessage3 = new frmMessage("", "Đã có phiên bản " + lastest_version_code + " mới hơn phiên bản " + text2 + " hiện tại.\r\nBạn muốn cập nhật không?", "5", "2", "Thông báo");
													((Form)frmMessage3).ShowDialog();
													if (frmMessage3.Yes)
													{
														Process.Start(rootUpdate + "\\Client_VersionUpdate.exe");
														Environment.Exit(0);
													}
												}
											}
										}
										finally
										{
											((IDisposable)val2)?.Dispose();
										}
									}
								}
							}
						}).Start();
					}
					((Control)this).Hide();
					((Form)this).ShowIcon = false;
					FormProcess.ReLoadfrmSystems();
				}
				else
				{
					ThemMSTMoi(((Control)cboMaSoThue).Text.Trim());
				}
			}
			else
			{
				ThemMSTMoi(((Control)cboMaSoThue).Text.Trim());
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void ThemMSTMoi(string strMST)
	{
		//IL_00f2: Unknown result type (might be due to invalid IL or missing references)
		try
		{
			if (!File.Exists(HTKKConfig.DATAFILES + "\\" + ((Control)cboMaSoThue).Text))
			{
				Directory.CreateDirectory(HTKKConfig.DATAFILES + "\\" + ((Control)cboMaSoThue).Text);
			}
			XmlNodeList val = ((XmlNode)WinSettings.xmlDocMenu).SelectNodes("/Root/Menu[@ID='100_1']");
			WinSettings.xmlMenuNode = val[0];
			Functions.ResetInfo(WinSettings.DicSheet, WinSettings.TTPeriodInfo, WinSettings.ToKhaiInfo);
			WinSettings.InstanceClass = WinDoWorkEvent.GetInstance(((XmlNode)WinSettings.xmlMenuNode.Attributes["Class"]).Value.ToString());
			WinSettings.HTKKHeaderInfo.mst = ((Control)cboMaSoThue).Text;
			WinSettings.ToKhaiInfo = WinFnProcess.GetToKhaiInfo(WinSettings.xmlMenuNode, WinSettings.TTPeriodInfo);
			WinSettings.XmlMessage.LoadXml(HTKKFile.LoadFile("Message.xml"));
			frmWinToKhai frmWinToKhai = new frmWinToKhai(WinSettings.xmlMenuNode);
			((Control)this).Hide();
			((Form)frmWinToKhai).ShowDialog();
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void btnMaSoMoi_Click(object sender, EventArgs e)
	{
		((Control)cboMaSoThue).Text = string.Empty;
	}

	private void btnThoat_Click(object sender, EventArgs e)
	{
		if (WinSettings.chuyenMSTKhac)
		{
			btnDongY_Click(sender, e);
		}
		else
		{
			Application.Exit();
		}
	}

	private void btnXoa_Click(object sender, EventArgs e)
	{
		//IL_005b: Unknown result type (might be due to invalid IL or missing references)
		//IL_011e: Unknown result type (might be due to invalid IL or missing references)
		//IL_00f6: Unknown result type (might be due to invalid IL or missing references)
		try
		{
			string text = "";
			bool flag = false;
			int index = 0;
			int num = ((ListControl)cboMaSoThue).SelectedIndex;
			text = ((Control)cboMaSoThue).Text;
			if (text.Trim().Equals(""))
			{
				frmMessage frmMessage = new frmMessage("0051", "3", "1", "Thông báo");
				((Form)frmMessage).ShowDialog();
				return;
			}
			if (lstMaSoThue != null && lstMaSoThue.Count > 0)
			{
				for (int i = 0; i < lstMaSoThue.Count; i++)
				{
					if (text.Equals(lstMaSoThue[i]))
					{
						flag = true;
						index = i;
						break;
					}
				}
			}
			if (!flag)
			{
				frmMessage frmMessage = new frmMessage("0048", "3", "1", "Thông báo");
				((Form)frmMessage).ShowDialog();
				return;
			}
			frmMessage frmMessage2 = new frmMessage("0049", "2", "2", "Cảnh báo");
			((Form)frmMessage2).ShowDialog();
			if (frmMessage2.Yes)
			{
				lstMaSoThue.RemoveAt(index);
				cboMaSoThue.DataSource = null;
				cboMaSoThue.DataSource = lstMaSoThue;
				if (num > lstMaSoThue.Count - 1)
				{
					num = lstMaSoThue.Count - 1;
				}
				((ListControl)cboMaSoThue).SelectedIndex = num;
				if (Directory.Exists(HTKKConfig.DATAFILES + "\\" + text))
				{
					Directory.Delete(HTKKConfig.DATAFILES + "\\" + text, recursive: true);
				}
				if (((Control)cboMaSoThue).Text.Equals(""))
				{
					((Control)btnXoa).Enabled = false;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void btExit_Click(object sender, EventArgs e)
	{
		Application.Exit();
	}

	private bool IsConnectServer()
	{
		try
		{
			ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
			HttpWebRequest httpWebRequest = (HttpWebRequest)WebRequest.Create(admin_serv_uri);
			httpWebRequest.Timeout = 25000;
			httpWebRequest.ReadWriteTimeout = 25000;
			httpWebRequest.Credentials = CredentialCache.DefaultNetworkCredentials;
			if (isProxy == "1")
			{
				WebProxy webProxy = new WebProxy("http://" + proxy + ":" + port, BypassOnLocal: true);
				if (isUser == "1" && !string.IsNullOrEmpty(userName))
				{
					string[] array = userName.Split(new char[1] { '\\' });
					if (array.Length > 1)
					{
						webProxy.Credentials = new NetworkCredential(array[1], pass, array[0]);
					}
					else
					{
						webProxy.Credentials = new NetworkCredential(array[1], pass);
					}
				}
				httpWebRequest.Proxy = webProxy;
			}
			HttpWebResponse httpWebResponse = ((httpWebRequest != null) ? ((HttpWebResponse)httpWebRequest.GetResponse()) : null);
			if (httpWebResponse != null && httpWebResponse.StatusCode == HttpStatusCode.OK)
			{
				httpWebRequest = null;
				httpWebResponse.Close();
				return true;
			}
			httpWebRequest = null;
			httpWebResponse.Close();
			return false;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return false;
		}
	}

	private bool isNewVersion(string curr_version)
	{
		//IL_004c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0052: Expected O, but got Unknown
		ServicePointManager.ServerCertificateValidationCallback = (object sender, X509Certificate certificate, X509Chain chain, SslPolicyErrors sslPolicyErrors) => true;
		bool flag = IsConnectServer();
		if (flag)
		{
			string hDDSerialNo = Foundation.HTKKSecurity.GetHDDSerialNo();
			string mst = WinSettings.HTKKHeaderInfo.mst;
			XmlDocument val = new XmlDocument();
			val.Load("../DataFiles/" + mst + "/HEADER.xml");
			tenCQTCapCuc = ((XmlNode)((XmlNode)val).SelectSingleNode("/Sections/Section/Cells/Cell[@Path='p28']").Attributes["Value"]).InnerText;
			maCQTCapCuc = ((XmlNode)((XmlNode)val).SelectSingleNode("/Sections/Section/Cells/Cell[@Path='p32']").Attributes["Value"]).InnerText;
			tenCQTNoiNop = ((XmlNode)((XmlNode)val).SelectSingleNode("/Sections/Section/Cells/Cell[@Path='p29']").Attributes["Value"]).InnerText;
			maCQTNoiNop = ((XmlNode)((XmlNode)val).SelectSingleNode("/Sections/Section/Cells/Cell[@Path='p33']").Attributes["Value"]).InnerText;
			if (version_mgt.CheckLastestVersion(hDDSerialNo, mst, curr_version, tenCQTCapCuc, maCQTCapCuc, tenCQTNoiNop, maCQTNoiNop, admin_serv_uri, out lastest_version_code, out down_file_info_list) != 0)
			{
				flag = false;
			}
			if (curr_version != null && lastest_version_code != null && curr_version.Replace(".", "").ToIntOrDefault() >= lastest_version_code.Replace(".", "").ToIntOrDefault())
			{
				flag = false;
			}
			return flag;
		}
		curr_version = "Không tìm thây";
		return flag;
	}

	private void cboMaSoThue_TextChanged(object sender, EventArgs e)
	{
		try
		{
			if (((Control)cboMaSoThue).Text.Equals(""))
			{
				((Control)lblNguoiNopThue).Text = "";
				((Control)btnXoa).Enabled = false;
			}
			else if (WinSettings.chuyenMSTKhac)
			{
				if (!HTKKFormater.isNullString(WinSettings.HTKKHeaderInfo.mst) && WinSettings.HTKKHeaderInfo.mst.Trim() == ((Control)cboMaSoThue).Text.Trim())
				{
					((Control)btnXoa).Enabled = false;
				}
				else
				{
					((Control)btnXoa).Enabled = true;
				}
			}
			else
			{
				((Control)btnXoa).Enabled = true;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void cboMaSoThue_KeyPress(object sender, KeyPressEventArgs e)
	{
		if (!char.IsDigit(e.KeyChar) && !char.IsControl(e.KeyChar))
		{
			e.Handled = true;
		}
	}

	private void frmFormStart_Move(object sender, EventArgs e)
	{
		FormProcess.SetConfigForProgcessBar((Control)(object)this);
	}

	public void FixHeader(string maSoThue)
	{
		//IL_006a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0070: Expected O, but got Unknown
		//IL_1e5f: Unknown result type (might be due to invalid IL or missing references)
		//IL_1e66: Expected O, but got Unknown
		//IL_0a76: Unknown result type (might be due to invalid IL or missing references)
		//IL_0a7d: Expected O, but got Unknown
		//IL_1c09: Unknown result type (might be due to invalid IL or missing references)
		//IL_1c10: Expected O, but got Unknown
		try
		{
			string text = HTKKConfig.DATAFILES + "\\" + maSoThue + "\\HEADER.xml";
			string text2 = HTKKConfig.DATAFILES + "\\" + maSoThue + "\\BK\\HEADER.xml";
			string path = HTKKConfig.DATAFILES + "\\" + maSoThue + "\\BK";
			if (!Directory.Exists(path))
			{
				Directory.CreateDirectory(path);
			}
			if (!File.Exists(text))
			{
				return;
			}
			XmlDocument val = new XmlDocument();
			val.LoadXml(HTKKFile.LoadFile(text));
			XmlNodeList elementsByTagName = val.GetElementsByTagName("Cell");
			XmlNode val2 = ((IEnumerable)elementsByTagName).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode x) => ((XmlNode)x.Attributes["PropertyInfo"]).Value == "maTinhNNT"));
			if (val2 == null)
			{
				XmlNode val3 = ((IEnumerable)elementsByTagName).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode x) => ((XmlNode)x.Attributes["PropertyInfo"]).Value == "tenTinhNNT"));
				((XmlNode)val3.Attributes["CellID"]).Value = "C_6";
				((XmlNode)val3.Attributes["Controltype"]).Value = "5";
				XmlAttribute val4 = val.CreateAttribute("SelectedValue");
				((XmlNode)val4).Value = "S_6";
				val3.Attributes.Append(val4);
				XmlAttribute val5 = val.CreateAttribute("ChildCell");
				((XmlNode)val5).Value = "C_8";
				val3.Attributes.Append(val5);
				XmlNode val6 = ((IEnumerable)elementsByTagName).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode x) => ((XmlNode)x.Attributes["PropertyInfo"]).Value == "tenHuyenNNT"));
				((XmlNode)val6.Attributes["Controltype"]).Value = "6";
				XmlAttribute val7 = val.CreateAttribute("SelectedValue");
				((XmlNode)val7).Value = "S_8";
				val6.Attributes.Append(val7);
				XmlAttribute val8 = val.CreateAttribute("ParentCell");
				((XmlNode)val8).Value = "S_6";
				val6.Attributes.Append(val8);
				XmlNode val9 = ((IEnumerable)elementsByTagName).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode x) => ((XmlNode)x.Attributes["PropertyInfo"]).Value == "dchiNNT"));
				((XmlNode)val9.Attributes["CellID"]).Value = "C_10";
				XmlNode val10 = ((IEnumerable)elementsByTagName).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode x) => ((XmlNode)x.Attributes["PropertyInfo"]).Value == "tenTinhDLyThue"));
				((XmlNode)val10.Attributes["CellID"]).Value = "C_43";
				((XmlNode)val10.Attributes["Controltype"]).Value = "5";
				XmlAttribute val11 = val.CreateAttribute("SelectedValue");
				((XmlNode)val11).Value = "S_43";
				val10.Attributes.Append(val11);
				XmlAttribute val12 = val.CreateAttribute("ChildCell");
				((XmlNode)val12).Value = "C_45";
				val10.Attributes.Append(val12);
				XmlNode val13 = ((IEnumerable)elementsByTagName).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode x) => ((XmlNode)x.Attributes["PropertyInfo"]).Value == "tenHuyenDLyThue"));
				((XmlNode)val13.Attributes["Controltype"]).Value = "6";
				XmlAttribute val14 = val.CreateAttribute("SelectedValue");
				((XmlNode)val14).Value = "S_45";
				val13.Attributes.Append(val14);
				XmlAttribute val15 = val.CreateAttribute("ParentCell");
				((XmlNode)val15).Value = "S_43";
				val13.Attributes.Append(val15);
				XmlNode val16 = ((IEnumerable)elementsByTagName).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode x) => ((XmlNode)x.Attributes["PropertyInfo"]).Value == "dchiDLyThue"));
				((XmlNode)val16.Attributes["CellID"]).Value = "C_47";
				XmlNode val17 = ((IEnumerable)elementsByTagName).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode x) => ((XmlNode)x.Attributes["PropertyInfo"]).Value == "tenNNT"));
				XmlNode val18 = val17.Clone();
				((XmlNode)val18.Attributes["CellID"]).Value = "S_6";
				((XmlNode)val18.Attributes["CellID2"]).Value = "V_51";
				((XmlNode)val18.Attributes["Path"]).Value = "p50";
				((XmlNode)val18.Attributes["PropertyInfo"]).Value = "maTinhNNT";
				((XmlNode)val18.Attributes["Value"]).Value = "";
				XmlNode val19 = val.ImportNode(val18, true);
				elementsByTagName[0].ParentNode.InsertAfter(val19, elementsByTagName[0].ParentNode.LastChild);
				XmlNode val20 = val17.Clone();
				((XmlNode)val20.Attributes["CellID"]).Value = "S_8";
				((XmlNode)val20.Attributes["CellID2"]).Value = "V_52";
				((XmlNode)val20.Attributes["Path"]).Value = "p51";
				((XmlNode)val20.Attributes["PropertyInfo"]).Value = "maHuyenNNT";
				((XmlNode)val20.Attributes["Value"]).Value = "";
				XmlNode val21 = val.ImportNode(val20, true);
				elementsByTagName[0].ParentNode.InsertAfter(val21, elementsByTagName[0].ParentNode.LastChild);
				XmlNode val22 = val17.Clone();
				((XmlNode)val22.Attributes["CellID"]).Value = "S_43";
				((XmlNode)val22.Attributes["CellID2"]).Value = "V_53";
				((XmlNode)val22.Attributes["Path"]).Value = "p52";
				((XmlNode)val22.Attributes["PropertyInfo"]).Value = "maTinhDLyThue";
				((XmlNode)val22.Attributes["Value"]).Value = "";
				XmlNode val23 = val.ImportNode(val22, true);
				elementsByTagName[0].ParentNode.InsertAfter(val23, elementsByTagName[0].ParentNode.LastChild);
				XmlNode val24 = val17.Clone();
				((XmlNode)val24.Attributes["CellID"]).Value = "S_45";
				((XmlNode)val24.Attributes["CellID2"]).Value = "V_54";
				((XmlNode)val24.Attributes["Path"]).Value = "p53";
				((XmlNode)val24.Attributes["PropertyInfo"]).Value = "maHuyenDLyThue";
				((XmlNode)val24.Attributes["Value"]).Value = "";
				XmlNode val25 = val.ImportNode(val24, true);
				elementsByTagName[0].ParentNode.InsertAfter(val25, elementsByTagName[0].ParentNode.LastChild);
				string valueTenTinh = ((XmlNode)val3.Attributes["Value"]).Value;
				KeyValuePair<string, string> keyValuePair = HTKKWinDictionaryDiaDanh.dicTinhTP.FirstOrDefault((KeyValuePair<string, string> x) => x.Key.ToLower().Trim() == valueTenTinh.ToLower().Trim());
				if (!string.IsNullOrEmpty(keyValuePair.Value))
				{
					XmlNode val26 = ((IEnumerable)elementsByTagName).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode x) => ((XmlNode)x.Attributes["PropertyInfo"]).Value == "maTinhNNT"));
					if (val26 != null)
					{
						((XmlNode)val26.Attributes["Value"]).Value = keyValuePair.Value;
					}
				}
				else
				{
					((XmlNode)val3.Attributes["Value"]).Value = "";
				}
				string valueTenHuyen = ((XmlNode)val6.Attributes["Value"]).Value;
				KeyValuePair<string, string> keyValuePair2 = HTKKWinDictionaryDiaDanh.dicQuanHuyen.FirstOrDefault((KeyValuePair<string, string> x) => x.Value.ToLower().Trim() == valueTenHuyen.ToLower().Trim());
				if (!string.IsNullOrEmpty(keyValuePair2.Value))
				{
					XmlNode val27 = ((IEnumerable)elementsByTagName).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode x) => ((XmlNode)x.Attributes["PropertyInfo"]).Value == "maHuyenNNT"));
					if (val27 != null)
					{
						((XmlNode)val27.Attributes["Value"]).Value = keyValuePair2.Key;
					}
				}
				else
				{
					((XmlNode)val6.Attributes["Value"]).Value = "";
				}
				string valueTenTinhDLThue = ((XmlNode)val10.Attributes["Value"]).Value;
				KeyValuePair<string, string> keyValuePair3 = HTKKWinDictionaryDiaDanh.dicTinhTP.FirstOrDefault((KeyValuePair<string, string> x) => x.Key.ToLower().Trim() == valueTenTinhDLThue.ToLower().Trim());
				if (!string.IsNullOrEmpty(keyValuePair3.Value))
				{
					XmlNode val28 = ((IEnumerable)elementsByTagName).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode x) => ((XmlNode)x.Attributes["PropertyInfo"]).Value == "maTinhDLyThue"));
					if (val28 != null)
					{
						((XmlNode)val28.Attributes["Value"]).Value = keyValuePair3.Value;
					}
				}
				else
				{
					((XmlNode)val10.Attributes["Value"]).Value = "";
				}
				string valueTenHuyenDLThue = ((XmlNode)val13.Attributes["Value"]).Value;
				KeyValuePair<string, string> keyValuePair4 = HTKKWinDictionaryDiaDanh.dicQuanHuyen.FirstOrDefault((KeyValuePair<string, string> x) => x.Value.ToLower().Trim() == valueTenHuyenDLThue.ToLower().Trim());
				if (!string.IsNullOrEmpty(keyValuePair4.Value))
				{
					XmlNode val29 = ((IEnumerable)elementsByTagName).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode x) => ((XmlNode)x.Attributes["PropertyInfo"]).Value == "maHuyenDLyThue"));
					if (val29 != null)
					{
						((XmlNode)val29.Attributes["Value"]).Value = keyValuePair4.Key;
					}
				}
				else
				{
					((XmlNode)val13.Attributes["Value"]).Value = "";
				}
				if (!File.Exists(text2))
				{
					Directory.Move(text, text2);
				}
				else
				{
					File.Delete(text2);
					Directory.Move(text, text2);
				}
				XmlTextWriter val30 = new XmlTextWriter(text, Encoding.UTF8);
				try
				{
					val30.Formatting = (Formatting)1;
					val.Save((XmlWriter)(object)val30);
				}
				finally
				{
					((IDisposable)val30)?.Dispose();
				}
			}
			val2 = ((IEnumerable)elementsByTagName).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode x) => ((XmlNode)x.Attributes["PropertyInfo"]).Value == "maXaPhuongNNT"));
			XmlNode val32;
			if (val2 == null)
			{
				List<MapDiaBan> mapDiaBan = WinFnProcess.getMapDiaBan("Catalogue_Dia_Ban_Tinh_Map.xml");
				XmlNode val31 = ((IEnumerable)elementsByTagName).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode x) => ((XmlNode)x.Attributes["PropertyInfo"]).Value == "tenHuyenNNT"));
				((XmlNode)val31.Attributes["Value"]).Value = "";
				((XmlNode)val31.Attributes["SelectedValue"]).Value = "";
				((XmlNode)val31.Attributes["ParentCell"]).Value = "";
				((XmlNode)val31.Attributes["CellID2"]).Value = "";
				XmlNode itemMaTinh = ((IEnumerable)elementsByTagName).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode x) => ((XmlNode)x.Attributes["PropertyInfo"]).Value == "maTinhNNT"));
				string maTinh = ((XmlNode)itemMaTinh.Attributes["Value"]).Value;
				((XmlNode)itemMaTinh.Attributes["Value"]).Value = mapDiaBan.FirstOrDefault((MapDiaBan x) => x.MaCu.Contains(maTinh)).MaMoi ?? maTinh;
				val32 = ((IEnumerable)elementsByTagName).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode x) => ((XmlNode)x.Attributes["PropertyInfo"]).Value == "tenTinhNNT"));
				((XmlNode)val32.Attributes["Controltype"]).Value = "100";
				((XmlNode)val32.Attributes["ChildCell"]).Value = "C_10";
				((XmlNode)val32.Attributes["Value"]).Value = HTKKWinDictionaryDiaDanh.dicTinhTPMoi.FirstOrDefault((KeyValuePair<string, string> x) => x.Value == itemMaTinh.Value).Key ?? ((XmlNode)val32.Attributes["Value"]).Value;
				val32 = ((IEnumerable)elementsByTagName).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode x) => ((XmlNode)x.Attributes["PropertyInfo"]).Value == "dchiNNT"));
				((XmlNode)val32.Attributes["CellID"]).Value = "C_16";
				val32 = ((IEnumerable)elementsByTagName).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode x) => ((XmlNode)x.Attributes["PropertyInfo"]).Value == "dthoaiNNT"));
				((XmlNode)val32.Attributes["CellID"]).Value = "C_18";
				val32 = ((IEnumerable)elementsByTagName).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode x) => ((XmlNode)x.Attributes["PropertyInfo"]).Value == "faxNNT"));
				((XmlNode)val32.Attributes["CellID"]).Value = "C_20";
				val32 = ((IEnumerable)elementsByTagName).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode x) => ((XmlNode)x.Attributes["PropertyInfo"]).Value == "emailNNT"));
				((XmlNode)val32.Attributes["CellID"]).Value = "C_22";
				val32 = ((IEnumerable)elementsByTagName).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode x) => ((XmlNode)x.Attributes["PropertyInfo"]).Value == "nganhNgheKD"));
				((XmlNode)val32.Attributes["CellID"]).Value = "C_24";
				val32 = ((IEnumerable)elementsByTagName).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode x) => ((XmlNode)x.Attributes["PropertyInfo"]).Value == "SoTKNganHang"));
				((XmlNode)val32.Attributes["CellID"]).Value = "C_26";
				val32 = ((IEnumerable)elementsByTagName).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode x) => ((XmlNode)x.Attributes["PropertyInfo"]).Value == "NgayBatDauNamTaiChinh"));
				((XmlNode)val32.Attributes["CellID"]).Value = "C_30";
				val32 = ((IEnumerable)elementsByTagName).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode x) => ((XmlNode)x.Attributes["PropertyInfo"]).Value == "nguoiKy"));
				((XmlNode)val32.Attributes["CellID"]).Value = "C_32";
				XmlNode val33 = ((IEnumerable)elementsByTagName).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode x) => ((XmlNode)x.Attributes["PropertyInfo"]).Value == "maCQTCapCuc"));
				string ma = ((XmlNode)val33.Attributes["Value"]).Value;
				val32 = ((IEnumerable)elementsByTagName).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode x) => ((XmlNode)x.Attributes["PropertyInfo"]).Value == "tenCQTCapCuc"));
				((XmlNode)val32.Attributes["CellID"]).Value = "C_34";
				((XmlNode)val32.Attributes["ChildCell"]).Value = "C_36";
				((XmlNode)val32.Attributes["SelectedValue"]).Value = "T_34";
				((XmlNode)val32.Attributes["Value"]).Value = HTKKWinDictionaryDiaDanh._cqtDiaBan.FirstOrDefault((CQTDiaBan x) => x.MaCQT == ma).TenCQT ?? ((XmlNode)val32.Attributes["Value"]).Value;
				val33 = ((IEnumerable)elementsByTagName).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode x) => ((XmlNode)x.Attributes["PropertyInfo"]).Value == "maCQTNoiNop"));
				ma = ((XmlNode)val33.Attributes["Value"]).Value;
				val32 = ((IEnumerable)elementsByTagName).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode x) => ((XmlNode)x.Attributes["PropertyInfo"]).Value == "tenCQTNoiNop"));
				((XmlNode)val32.Attributes["CellID"]).Value = "C_36";
				((XmlNode)val32.Attributes["ParentCell"]).Value = "T_34";
				((XmlNode)val32.Attributes["SelectedValue"]).Value = "T_36";
				((XmlNode)val32.Attributes["Value"]).Value = HTKKWinDictionaryDiaDanh._cqtDiaBan.FirstOrDefault((CQTDiaBan x) => x.MaCQT == ma).TenCQT ?? ((XmlNode)val32.Attributes["Value"]).Value;
				val32 = ((IEnumerable)elementsByTagName).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode x) => ((XmlNode)x.Attributes["PropertyInfo"]).Value == "maCQTCapCuc"));
				((XmlNode)val32.Attributes["CellID"]).Value = "T_34";
				val32 = ((IEnumerable)elementsByTagName).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode x) => ((XmlNode)x.Attributes["PropertyInfo"]).Value == "maCQTNoiNop"));
				((XmlNode)val32.Attributes["CellID"]).Value = "T_36";
				val32 = ((IEnumerable)elementsByTagName).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode x) => ((XmlNode)x.Attributes["PropertyInfo"]).Value == "TenDonViChuQuan"));
				((XmlNode)val32.Attributes["CellID"]).Value = "C_38";
				val32 = ((IEnumerable)elementsByTagName).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode x) => ((XmlNode)x.Attributes["PropertyInfo"]).Value == "MaSoThueDonViChuQuan"));
				((XmlNode)val32.Attributes["CellID"]).Value = "C_40";
				val32 = ((IEnumerable)elementsByTagName).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode x) => ((XmlNode)x.Attributes["PropertyInfo"]).Value == "ThongTinDaiLy"));
				((XmlNode)val32.Attributes["CellID"]).Value = "C_43";
				val32 = ((IEnumerable)elementsByTagName).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode x) => ((XmlNode)x.Attributes["PropertyInfo"]).Value == "mstDLyThue"));
				((XmlNode)val32.Attributes["CellID"]).Value = "C_45";
				val32 = ((IEnumerable)elementsByTagName).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode x) => ((XmlNode)x.Attributes["PropertyInfo"]).Value == "tenDLyThue"));
				((XmlNode)val32.Attributes["CellID"]).Value = "C_47";
				val32 = ((IEnumerable)elementsByTagName).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode x) => ((XmlNode)x.Attributes["PropertyInfo"]).Value == "dchiDLyThue"));
				((XmlNode)val32.Attributes["CellID"]).Value = "C_59";
				val32 = ((IEnumerable)elementsByTagName).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode x) => ((XmlNode)x.Attributes["PropertyInfo"]).Value == "tenHuyenDLyThue"));
				((XmlNode)val32.Attributes["CellID"]).Value = "C_51";
				val32 = ((IEnumerable)elementsByTagName).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode x) => ((XmlNode)x.Attributes["PropertyInfo"]).Value == "tenTinhDLyThue"));
				((XmlNode)val32.Attributes["Controltype"]).Value = "100";
				((XmlNode)val32.Attributes["CellID"]).Value = "C_49";
				((XmlNode)val32.Attributes["SelectedValue"]).Value = "S_49";
				((XmlNode)val32.Attributes["ChildCell"]).Value = "C_53";
				val32 = ((IEnumerable)elementsByTagName).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode x) => ((XmlNode)x.Attributes["PropertyInfo"]).Value == "dthoaiDLyThue"));
				((XmlNode)val32.Attributes["CellID"]).Value = "C_61";
				val32 = ((IEnumerable)elementsByTagName).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode x) => ((XmlNode)x.Attributes["PropertyInfo"]).Value == "faxDLyThue"));
				((XmlNode)val32.Attributes["CellID"]).Value = "C_63";
				val32 = ((IEnumerable)elementsByTagName).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode x) => ((XmlNode)x.Attributes["PropertyInfo"]).Value == "emailDLyThue"));
				((XmlNode)val32.Attributes["CellID"]).Value = "C_65";
				val32 = ((IEnumerable)elementsByTagName).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode x) => ((XmlNode)x.Attributes["PropertyInfo"]).Value == "soHDongDLyThue"));
				((XmlNode)val32.Attributes["CellID"]).Value = "C_67";
				val32 = ((IEnumerable)elementsByTagName).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode x) => ((XmlNode)x.Attributes["PropertyInfo"]).Value == "ngayKyHDDLyThue"));
				((XmlNode)val32.Attributes["CellID"]).Value = "C_69";
				val32 = ((IEnumerable)elementsByTagName).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode x) => ((XmlNode)x.Attributes["PropertyInfo"]).Value == "tenNVienDLyThue"));
				((XmlNode)val32.Attributes["CellID"]).Value = "C_71";
				val32 = ((IEnumerable)elementsByTagName).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode x) => ((XmlNode)x.Attributes["PropertyInfo"]).Value == "cchiHNghe"));
				((XmlNode)val32.Attributes["CellID"]).Value = "C_73";
				val32 = ((IEnumerable)elementsByTagName).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode x) => ((XmlNode)x.Attributes["PropertyInfo"]).Value == "ThongTinDonViChuQuan"));
				((XmlNode)val32.Attributes["CellID"]).Value = "P_38";
				val32 = ((IEnumerable)elementsByTagName).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode x) => ((XmlNode)x.Attributes["PropertyInfo"]).Value == "maTinhDLyThue"));
				((XmlNode)val32.Attributes["CellID"]).Value = "S_49";
				val32 = ((IEnumerable)elementsByTagName).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode x) => ((XmlNode)x.Attributes["PropertyInfo"]).Value == "maHuyenDLyThue"));
				((XmlNode)val32.Attributes["CellID"]).Value = "S_51";
				XmlNode val17 = ((IEnumerable)elementsByTagName).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode x) => ((XmlNode)x.Attributes["PropertyInfo"]).Value == "tenHuyenNNT"));
				val32 = val17.Clone();
				((XmlNode)val32.Attributes["CellID"]).Value = "C_10";
				((XmlNode)val32.Attributes["CellID2"]).Value = "C_5";
				((XmlNode)val32.Attributes["Path"]).Value = "p54";
				((XmlNode)val32.Attributes["PropertyInfo"]).Value = "tenXaPhuongNNT";
				((XmlNode)val32.Attributes["Value"]).Value = "";
				((XmlNode)val32.Attributes["Controltype"]).Value = "101";
				((XmlNode)val32.Attributes["SelectedValue"]).Value = "S_10";
				((XmlNode)val32.Attributes["ParentCell"]).Value = "S_6";
				XmlNode val34 = val.ImportNode(val32, true);
				elementsByTagName[0].ParentNode.InsertAfter(val34, elementsByTagName[0].ParentNode.LastChild);
				val17 = ((IEnumerable)elementsByTagName).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode x) => ((XmlNode)x.Attributes["PropertyInfo"]).Value == "tenHuyenDLyThue"));
				val32 = val17.Clone();
				((XmlNode)val32.Attributes["CellID"]).Value = "C_53";
				((XmlNode)val32.Attributes["CellID2"]).Value = "V_8";
				((XmlNode)val32.Attributes["Path"]).Value = "p55";
				((XmlNode)val32.Attributes["PropertyInfo"]).Value = "tenXaPhuongDLyThue";
				((XmlNode)val32.Attributes["Value"]).Value = "";
				((XmlNode)val32.Attributes["Controltype"]).Value = "101";
				((XmlNode)val32.Attributes["SelectedValue"]).Value = "S_53";
				((XmlNode)val32.Attributes["ParentCell"]).Value = "S_49";
				val34 = val.ImportNode(val32, true);
				elementsByTagName[0].ParentNode.InsertAfter(val34, elementsByTagName[0].ParentNode.LastChild);
				val17 = ((IEnumerable)elementsByTagName).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode x) => ((XmlNode)x.Attributes["PropertyInfo"]).Value == "maHuyenNNT"));
				val32 = val17.Clone();
				((XmlNode)val32.Attributes["CellID"]).Value = "S_10";
				((XmlNode)val32.Attributes["CellID2"]).Value = "";
				((XmlNode)val32.Attributes["Path"]).Value = "p56";
				((XmlNode)val32.Attributes["PropertyInfo"]).Value = "maXaPhuongNNT";
				((XmlNode)val32.Attributes["Value"]).Value = "";
				val34 = val.ImportNode(val32, true);
				elementsByTagName[0].ParentNode.InsertAfter(val34, elementsByTagName[0].ParentNode.LastChild);
				val17 = ((IEnumerable)elementsByTagName).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode x) => ((XmlNode)x.Attributes["PropertyInfo"]).Value == "maHuyenNNT"));
				val32 = val17.Clone();
				((XmlNode)val32.Attributes["CellID"]).Value = "S_53";
				((XmlNode)val32.Attributes["CellID2"]).Value = "";
				((XmlNode)val32.Attributes["Path"]).Value = "p57";
				((XmlNode)val32.Attributes["PropertyInfo"]).Value = "maXaPhuongDLyThue";
				((XmlNode)val32.Attributes["Value"]).Value = "";
				val34 = val.ImportNode(val32, true);
				elementsByTagName[0].ParentNode.InsertAfter(val34, elementsByTagName[0].ParentNode.LastChild);
				if (!File.Exists(text2))
				{
					Directory.Move(text, text2);
				}
				else
				{
					File.Delete(text2);
					Directory.Move(text, text2);
				}
				XmlTextWriter val30 = new XmlTextWriter(text, Encoding.UTF8);
				try
				{
					val30.Formatting = (Formatting)1;
					val.Save((XmlWriter)(object)val30);
					return;
				}
				finally
				{
					((IDisposable)val30)?.Dispose();
				}
			}
			int num = 0;
			val32 = ((IEnumerable)elementsByTagName).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode x) => ((XmlNode)x.Attributes["PropertyInfo"]).Value == "tenXaPhuongNNT"));
			if (((XmlNode)val32.Attributes["CellID2"]).Value == "")
			{
				((XmlNode)val32.Attributes["CellID2"]).Value = "C_5";
				num = 1;
			}
			val32 = ((IEnumerable)elementsByTagName).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode x) => ((XmlNode)x.Attributes["PropertyInfo"]).Value == "tenXaPhuongDLyThue"));
			if (((XmlNode)val32.Attributes["CellID2"]).Value == "")
			{
				((XmlNode)val32.Attributes["CellID2"]).Value = "V_8";
				num = 1;
			}
			val32 = ((IEnumerable)elementsByTagName).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode x) => ((XmlNode)x.Attributes["PropertyInfo"]).Value == "tenHuyenNNT"));
			if (((XmlNode)val32.Attributes["CellID2"]).Value != "")
			{
				((XmlNode)val32.Attributes["CellID2"]).Value = "";
				num = 1;
			}
			val32 = ((IEnumerable)elementsByTagName).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode x) => ((XmlNode)x.Attributes["PropertyInfo"]).Value == "tenHuyenDLyThue"));
			if (((XmlNode)val32.Attributes["CellID2"]).Value != "")
			{
				((XmlNode)val32.Attributes["CellID2"]).Value = "";
				num = 1;
			}
			if (num == 1)
			{
				if (!File.Exists(text2))
				{
					Directory.Move(text, text2);
				}
				else
				{
					File.Delete(text2);
					Directory.Move(text, text2);
				}
				XmlTextWriter val30 = new XmlTextWriter(text, Encoding.UTF8);
				try
				{
					val30.Formatting = (Formatting)1;
					val.Save((XmlWriter)(object)val30);
					return;
				}
				finally
				{
					((IDisposable)val30)?.Dispose();
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public Dictionary<string, string> LoadMappingFromXml(string xmlPath)
	{
		Dictionary<string, string> dictionary = new Dictionary<string, string>();
		XDocument val = XDocument.Load(xmlPath);
		IEnumerable<XElement> enumerable = ((XContainer)val).Descendants(XName.op_Implicit("Item"));
		foreach (XElement item in enumerable)
		{
			string text = (string)item.Attribute(XName.op_Implicit("MaCu"));
			string value = (string)item.Attribute(XName.op_Implicit("MaMoi"));
			if (string.IsNullOrEmpty(text) || string.IsNullOrEmpty(value))
			{
				continue;
			}
			string[] array = text.Split(new char[1] { ',' });
			string[] array2 = array;
			foreach (string text2 in array2)
			{
				string key = text2.Trim();
				if (!dictionary.ContainsKey(key))
				{
					dictionary[key] = value;
				}
			}
		}
		return dictionary;
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
		//IL_0075: Unknown result type (might be due to invalid IL or missing references)
		//IL_007f: Expected O, but got Unknown
		//IL_0080: Unknown result type (might be due to invalid IL or missing references)
		//IL_008a: Expected O, but got Unknown
		//IL_008b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0095: Expected O, but got Unknown
		//IL_0096: Unknown result type (might be due to invalid IL or missing references)
		//IL_00a0: Expected O, but got Unknown
		//IL_00a1: Unknown result type (might be due to invalid IL or missing references)
		//IL_00ab: Expected O, but got Unknown
		//IL_025e: Unknown result type (might be due to invalid IL or missing references)
		//IL_02ce: Unknown result type (might be due to invalid IL or missing references)
		//IL_02d8: Expected O, but got Unknown
		//IL_06a3: Unknown result type (might be due to invalid IL or missing references)
		//IL_06ad: Expected O, but got Unknown
		//IL_06cc: Unknown result type (might be due to invalid IL or missing references)
		//IL_06d6: Expected O, but got Unknown
		//IL_07ec: Unknown result type (might be due to invalid IL or missing references)
		//IL_07f6: Expected O, but got Unknown
		//IL_08af: Unknown result type (might be due to invalid IL or missing references)
		//IL_08b9: Expected O, but got Unknown
		//IL_0951: Unknown result type (might be due to invalid IL or missing references)
		//IL_095b: Expected O, but got Unknown
		//IL_0992: Unknown result type (might be due to invalid IL or missing references)
		//IL_0a29: Unknown result type (might be due to invalid IL or missing references)
		//IL_0a33: Expected O, but got Unknown
		//IL_0a9c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0aa6: Expected O, but got Unknown
		//IL_0ad7: Unknown result type (might be due to invalid IL or missing references)
		//IL_0ae1: Expected O, but got Unknown
		ComponentResourceManager componentResourceManager = new ComponentResourceManager(typeof(frmFormStart));
		panel_login = new Panel();
		panel_tenNNT = new Panel();
		lblNguoiNopThue = new Label();
		btnThoat = new Button();
		btnDongY = new Button();
		btnMaSoMoi = new Button();
		btnXoa = new Button();
		label2 = new Label();
		cboMaSoThue = new ComboBox();
		pictureBox1 = new PictureBox();
		panel1 = new Panel();
		label1 = new Label();
		lblTitle_Version_bg = new Label();
		label3 = new Label();
		((Control)panel_login).SuspendLayout();
		((Control)panel_tenNNT).SuspendLayout();
		((ISupportInitialize)pictureBox1).BeginInit();
		((Control)panel1).SuspendLayout();
		((Control)this).SuspendLayout();
		((Control)panel_login).AutoSize = true;
		panel_login.BorderStyle = (BorderStyle)1;
		((Control)panel_login).Controls.Add((Control)(object)panel_tenNNT);
		((Control)panel_login).Controls.Add((Control)(object)btnThoat);
		((Control)panel_login).Controls.Add((Control)(object)btnDongY);
		((Control)panel_login).Controls.Add((Control)(object)btnMaSoMoi);
		((Control)panel_login).Controls.Add((Control)(object)btnXoa);
		((Control)panel_login).Controls.Add((Control)(object)label2);
		((Control)panel_login).Controls.Add((Control)(object)cboMaSoThue);
		((Control)panel_login).Controls.Add((Control)(object)pictureBox1);
		((Control)panel_login).Controls.Add((Control)(object)panel1);
		((Control)panel_login).Location = new Point(311, 154);
		((Control)panel_login).Name = "panel_login";
		((Control)panel_login).Size = new Size(371, 226);
		((Control)panel_login).TabIndex = 0;
		((Control)panel_tenNNT).Controls.Add((Control)(object)lblNguoiNopThue);
		((Control)panel_tenNNT).Dock = (DockStyle)1;
		((Control)panel_tenNNT).Location = new Point(0, 115);
		((Control)panel_tenNNT).Margin = new Padding(0);
		((Control)panel_tenNNT).Name = "panel_tenNNT";
		((Control)panel_tenNNT).Size = new Size(369, 32);
		((Control)panel_tenNNT).TabIndex = 13;
		((Control)lblNguoiNopThue).Anchor = (AnchorStyles)15;
		((Control)lblNguoiNopThue).AutoSize = true;
		((Control)lblNguoiNopThue).Font = new Font("Microsoft Sans Serif", 8.25f, (FontStyle)1, (GraphicsUnit)3, (byte)0);
		((Control)lblNguoiNopThue).Location = new Point(1, 10);
		((Control)lblNguoiNopThue).Name = "lblNguoiNopThue";
		((Control)lblNguoiNopThue).Size = new Size(0, 13);
		((Control)lblNguoiNopThue).TabIndex = 12;
		lblNguoiNopThue.TextAlign = (ContentAlignment)32;
		((ButtonBase)btnThoat).FlatStyle = (FlatStyle)3;
		((ButtonBase)btnThoat).Image = (Image)(object)Resources._16_logout;
		((ButtonBase)btnThoat).ImageAlign = (ContentAlignment)16;
		((Control)btnThoat).Location = new Point(272, 192);
		((Control)btnThoat).Name = "btnThoat";
		((Control)btnThoat).Size = new Size(85, 25);
		((Control)btnThoat).TabIndex = 11;
		((Control)btnThoat).Text = "&Thoát";
		((ButtonBase)btnThoat).UseVisualStyleBackColor = true;
		((Control)btnThoat).Click += btnThoat_Click;
		((ButtonBase)btnDongY).ImageAlign = (ContentAlignment)16;
		((Control)btnDongY).Location = new Point(185, 192);
		((Control)btnDongY).Name = "btnDongY";
		((Control)btnDongY).Size = new Size(85, 25);
		((Control)btnDongY).TabIndex = 10;
		((Control)btnDongY).Text = "Đồ&ng ý";
		((ButtonBase)btnDongY).UseVisualStyleBackColor = true;
		((Control)btnDongY).Click += btnDongY_Click;
		((ButtonBase)btnMaSoMoi).ImageAlign = (ContentAlignment)16;
		((Control)btnMaSoMoi).Location = new Point(98, 192);
		((Control)btnMaSoMoi).Name = "btnMaSoMoi";
		((Control)btnMaSoMoi).Size = new Size(85, 25);
		((Control)btnMaSoMoi).TabIndex = 9;
		((Control)btnMaSoMoi).Text = "&Mã số mới";
		((ButtonBase)btnMaSoMoi).UseVisualStyleBackColor = true;
		((Control)btnMaSoMoi).Click += btnMaSoMoi_Click;
		((ButtonBase)btnXoa).ImageAlign = (ContentAlignment)16;
		((Control)btnXoa).Location = new Point(11, 192);
		((Control)btnXoa).Name = "btnXoa";
		((Control)btnXoa).Size = new Size(85, 25);
		((Control)btnXoa).TabIndex = 8;
		((Control)btnXoa).Text = "&Xóa";
		((ButtonBase)btnXoa).UseVisualStyleBackColor = true;
		((Control)btnXoa).Click += btnXoa_Click;
		((Control)label2).AutoSize = true;
		((Control)label2).Location = new Point(65, 157);
		((Control)label2).Name = "label2";
		((Control)label2).Size = new Size(60, 13);
		((Control)label2).TabIndex = 4;
		((Control)label2).Text = "Mã số thuế";
		((ListControl)cboMaSoThue).FormattingEnabled = true;
		((Control)cboMaSoThue).Location = new Point(131, 154);
		cboMaSoThue.MaxLength = 13;
		((Control)cboMaSoThue).Name = "cboMaSoThue";
		((Control)cboMaSoThue).Size = new Size(196, 21);
		((Control)cboMaSoThue).TabIndex = 3;
		cboMaSoThue.SelectedIndexChanged += cboMaSoThue_SelectedIndexChanged;
		((Control)cboMaSoThue).TextChanged += cboMaSoThue_TextChanged;
		((Control)cboMaSoThue).KeyPress += new KeyPressEventHandler(cboMaSoThue_KeyPress);
		((Control)pictureBox1).Dock = (DockStyle)1;
		pictureBox1.Image = (Image)componentResourceManager.GetObject("pictureBox1.Image");
		((Control)pictureBox1).Location = new Point(0, 26);
		((Control)pictureBox1).Name = "pictureBox1";
		((Control)pictureBox1).Size = new Size(369, 89);
		pictureBox1.TabIndex = 1;
		pictureBox1.TabStop = false;
		((Control)panel1).BackColor = Color.FromArgb(192, 192, 255);
		((Control)panel1).Controls.Add((Control)(object)label1);
		((Control)panel1).Dock = (DockStyle)1;
		((Control)panel1).Location = new Point(0, 0);
		((Control)panel1).Name = "panel1";
		((Control)panel1).Size = new Size(369, 26);
		((Control)panel1).TabIndex = 0;
		((Control)label1).AutoSize = true;
		((Control)label1).BackColor = Color.Transparent;
		((Control)label1).Font = new Font("Calibri", 11.25f, (FontStyle)1, (GraphicsUnit)3, (byte)0);
		((Control)label1).ForeColor = Color.White;
		((Control)label1).Location = new Point(0, 3);
		((Control)label1).Name = "label1";
		((Control)label1).Size = new Size(157, 18);
		((Control)label1).TabIndex = 0;
		((Control)label1).Text = "Đăng nhập vào hệ thống";
		label1.TextAlign = (ContentAlignment)32;
		((Control)lblTitle_Version_bg).Anchor = (AnchorStyles)10;
		((Control)lblTitle_Version_bg).AutoSize = true;
		((Control)lblTitle_Version_bg).BackColor = Color.Transparent;
		((Control)lblTitle_Version_bg).Font = new Font("Microsoft Sans Serif", 10f, (FontStyle)1, (GraphicsUnit)3, (byte)0);
		((Control)lblTitle_Version_bg).ForeColor = Color.DarkRed;
		((Control)lblTitle_Version_bg).Location = new Point(903, 500);
		((Control)lblTitle_Version_bg).Name = "lblTitle_Version_bg";
		((Control)lblTitle_Version_bg).Size = new Size(128, 17);
		((Control)lblTitle_Version_bg).TabIndex = 3;
		((Control)lblTitle_Version_bg).Text = "Phiên bản: 4.0.0";
		((Control)label3).BackColor = Color.Transparent;
		((Control)label3).Font = new Font("Microsoft Sans Serif", 9f, (FontStyle)1, (GraphicsUnit)0, (byte)0);
		((Control)label3).ForeColor = Color.White;
		((Control)label3).Location = new Point(129, 534);
		((Control)label3).Margin = new Padding(2, 0, 2, 0);
		((Control)label3).Name = "label3";
		((Control)label3).Size = new Size(740, 41);
		((Control)label3).TabIndex = 4;
		((Control)label3).Text = "Bản quyền thuộc Cục Thuế - Bộ Tài chính\r\nĐịa chỉ: Số 1A Nguyễn Công Trứ, Phường Hai Bà Trưng, Thành phố Hà Nội; Website: http://www.gdt.gov.vn\r\nPhần mềm cung cấp miễn phí cho người nộp thuế\r\n";
		label3.TextAlign = (ContentAlignment)32;
		((Form)this).AcceptButton = (IButtonControl)(object)btnDongY;
		((ContainerControl)this).AutoScaleDimensions = new SizeF(6f, 13f);
		((ContainerControl)this).AutoScaleMode = (AutoScaleMode)1;
		((Control)this).BackgroundImage = (Image)componentResourceManager.GetObject("$this.BackgroundImage");
		((Control)this).BackgroundImageLayout = (ImageLayout)3;
		((Form)this).ClientSize = new Size(1028, 582);
		((Control)this).Controls.Add((Control)(object)label3);
		((Control)this).Controls.Add((Control)(object)lblTitle_Version_bg);
		((Control)this).Controls.Add((Control)(object)panel_login);
		((Form)this).FormBorderStyle = (FormBorderStyle)1;
		((Form)this).Icon = (Icon)componentResourceManager.GetObject("$this.Icon");
		((Form)this).MaximizeBox = false;
		((Control)this).Name = "frmFormStart";
		((Form)this).StartPosition = (FormStartPosition)1;
		((Control)this).Text = "frmFormStart";
		((Form)this).FormClosing += new FormClosingEventHandler(frmFormStart_FormClosing);
		((Form)this).Load += frmFormStart_Load;
		((Control)this).Move += frmFormStart_Move;
		((Control)this).Resize += frmFormStart_Resize;
		((Control)panel_login).ResumeLayout(false);
		((Control)panel_login).PerformLayout();
		((Control)panel_tenNNT).ResumeLayout(false);
		((Control)panel_tenNNT).PerformLayout();
		((ISupportInitialize)pictureBox1).EndInit();
		((Control)panel1).ResumeLayout(false);
		((Control)panel1).PerformLayout();
		((Control)this).ResumeLayout(false);
		((Control)this).PerformLayout();
	}
}
