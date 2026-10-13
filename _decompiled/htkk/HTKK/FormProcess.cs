using System;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Text;
using System.Windows.Forms;
using System.Windows.Forms.Layout;
using System.Xml;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Implement;

namespace HTKK;

public class FormProcess
{
	public static void LoadForm(Form objForm, Panel panelContainer)
	{
		try
		{
			((Control)panelContainer).Controls.Clear();
			objForm.TopLevel = false;
			((Control)panelContainer).Controls.Add((Control)(object)objForm);
			objForm.FormBorderStyle = (FormBorderStyle)0;
			((Control)objForm).Dock = (DockStyle)5;
			((Control)panelContainer).Visible = true;
			((Control)objForm).Show();
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void LoadControl(Control container, UserControl u)
	{
		try
		{
			container.Controls.Clear();
			((Control)u).Dock = (DockStyle)5;
			container.Controls.Add((Control)(object)u);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void LoadUserControl(Control container, string controlName, object[] parameterList)
	{
		//IL_0034: Unknown result type (might be due to invalid IL or missing references)
		//IL_003a: Expected O, but got Unknown
		try
		{
			container.Controls.Clear();
			Assembly assembly = Assembly.Load("HTKK");
			Type type = assembly.GetType("HTKK." + controlName, throwOnError: true, ignoreCase: true);
			UserControl val = (UserControl)Activator.CreateInstance(type, parameterList);
			((Control)val).Dock = (DockStyle)1;
			container.Controls.Add((Control)(object)val);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static Control FindControl(Control parent, string ctlName)
	{
		//IL_0017: Unknown result type (might be due to invalid IL or missing references)
		//IL_001d: Expected O, but got Unknown
		try
		{
			foreach (Control item in (ArrangedElementCollection)parent.Controls)
			{
				Control val = item;
				if (val.Name.Equals(ctlName))
				{
					return val;
				}
				FindControl(val, ctlName);
			}
			return null;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public static void SetReadOnlyControlTexBox(bool value, Control parent)
	{
		//IL_0017: Unknown result type (might be due to invalid IL or missing references)
		//IL_001d: Expected O, but got Unknown
		try
		{
			foreach (Control item in (ArrangedElementCollection)parent.Controls)
			{
				Control val = item;
				if (val.Name.IndexOf("txt") != -1)
				{
					TextBox val2 = (TextBox)(object)((val is TextBox) ? val : null);
					((TextBoxBase)val2).ReadOnly = value;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void SetValueControlTexBox(string value, Control parent)
	{
		//IL_0017: Unknown result type (might be due to invalid IL or missing references)
		//IL_001d: Expected O, but got Unknown
		try
		{
			foreach (Control item in (ArrangedElementCollection)parent.Controls)
			{
				Control val = item;
				if (val.Name.IndexOf("txt") != -1)
				{
					TextBox val2 = (TextBox)(object)((val is TextBox) ? val : null);
					((Control)val2).Text = value;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void ReLoadfrmSystems()
	{
		//IL_0139: Unknown result type (might be due to invalid IL or missing references)
		try
		{
			if (WinSettings.myWinToKhai != null)
			{
				((Control)WinSettings.myWinToKhai).Visible = false;
			}
			if (WinSettings.myHTKKSystems != null)
			{
				HTKKFactory.GetInstanceHTKKHeader().SetPropertyHeader(ref WinSettings.HTKKHeaderInfo, WinSettings.HTKKHeaderInfo.mst);
				((ToolStripItem)WinSettings.myHTKKSystems.toolStripStatusMST).Text = WinSettings.HTKKHeaderInfo.mst + " : " + WinSettings.HTKKHeaderInfo.tenNNT.Replace("&", "&&");
				if (WinSettings.myHTKKSystems.formName == "frmTraCuuToKhai")
				{
					((Control)WinSettings.myHTKKSystems).Visible = true;
					return;
				}
				((Control)WinSettings.myHTKKSystems.panelMenu).Visible = true;
				((Control)WinSettings.myHTKKSystems.panelPeriod).Size = new Size(1, 1);
				((Control)WinSettings.myHTKKSystems.panelPeriod).Location = new Point(((Control)WinSettings.myHTKKSystems).Width, ((Control)WinSettings.myHTKKSystems).Height);
				WinSettings.myHTKKSystems.panelPeriod.BorderStyle = (BorderStyle)1;
				((Control)WinSettings.myHTKKSystems.panelPeriod).Visible = false;
				((Control)WinSettings.myHTKKSystems).Show();
			}
			else
			{
				frmHTKKSystems frmHTKKSystems2 = new frmHTKKSystems();
				((Form)frmHTKKSystems2).ShowDialog();
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void AddControl(Form form, Control control)
	{
		//IL_0077: Unknown result type (might be due to invalid IL or missing references)
		try
		{
			BeginControlUpdate(control);
			control.Controls.Clear();
			control.Visible = true;
			form.TopLevel = false;
			control.Controls.AddRange((Control[])(object)new Control[1] { (Control)form });
			((Control)form).Show();
			control.Width = ((Control)form).Width;
			control.Height = ((Control)form).Height;
			((Control)form).Dock = (DockStyle)5;
			if (control is Panel)
			{
				((Panel)control).BorderStyle = (BorderStyle)1;
			}
			ResizeForm(control);
			EndControlUpdate(control);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void ChildFormClosed(Form form, Control control)
	{
		//IL_0071: Unknown result type (might be due to invalid IL or missing references)
		try
		{
			if (form != null)
			{
				control.Size = new Size(1, 1);
				if (control.Parent != null)
				{
					control.Location = new Point(control.Parent.Width, control.Parent.Height);
				}
				else
				{
					control.Location = new Point(0, 0);
				}
				control.Visible = false;
				((Control)form).Show();
			}
			else
			{
				frmHTKKSystems frmHTKKSystems2 = new frmHTKKSystems();
				((Form)frmHTKKSystems2).ShowDialog();
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void ResizeForm(Control control)
	{
		try
		{
			BeginControlUpdate(control);
			int x = (control.Parent.Width - control.Width) / 2;
			int y = (control.Parent.Height - control.Height) / 2;
			control.Location = new Point(x, y);
			EndControlUpdate(control);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void BeginControlUpdate(Control control)
	{
		//IL_0013: Unknown result type (might be due to invalid IL or missing references)
		//IL_0018: Unknown result type (might be due to invalid IL or missing references)
		Message val = Message.Create(control.Handle, 11, IntPtr.Zero, IntPtr.Zero);
		NativeWindow val2 = NativeWindow.FromHandle(control.Handle);
		val2.DefWndProc(ref val);
	}

	public static void EndControlUpdate(Control control)
	{
		//IL_0018: Unknown result type (might be due to invalid IL or missing references)
		//IL_001d: Unknown result type (might be due to invalid IL or missing references)
		IntPtr intPtr = new IntPtr(1);
		Message val = Message.Create(control.Handle, 11, intPtr, IntPtr.Zero);
		NativeWindow val2 = NativeWindow.FromHandle(control.Handle);
		val2.DefWndProc(ref val);
		control.Invalidate();
		control.Refresh();
	}

	public static void SuspendsChild(Control pCtl)
	{
		//IL_0017: Unknown result type (might be due to invalid IL or missing references)
		//IL_001d: Expected O, but got Unknown
		try
		{
			foreach (Control item in (ArrangedElementCollection)pCtl.Controls)
			{
				Control val = item;
				if (val.HasChildren)
				{
					SuspendsChild(val);
				}
				BeginControlUpdate(val);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void ResumesChild(Control pCtl)
	{
		//IL_0017: Unknown result type (might be due to invalid IL or missing references)
		//IL_001d: Expected O, but got Unknown
		try
		{
			foreach (Control item in (ArrangedElementCollection)pCtl.Controls)
			{
				Control val = item;
				if (val.HasChildren)
				{
					ResumesChild(val);
				}
				EndControlUpdate(val);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void SetConfigForProgcessBar(Control control)
	{
		//IL_0019: Unknown result type (might be due to invalid IL or missing references)
		//IL_001f: Expected O, but got Unknown
		//IL_013c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0143: Expected O, but got Unknown
		try
		{
			if (!File.Exists("Config.xml"))
			{
				return;
			}
			XmlDocument val = new XmlDocument();
			val.LoadXml(HTKKFile.LoadFile("Config.xml"));
			XmlNode val2 = ((XmlNode)val).SelectNodes("/Root/SizeAPP")[0];
			if (val2 == null)
			{
				XmlNode val3 = val.CreateNode((XmlNodeType)1, "SizeAPP", (string)null);
				XmlNode val4 = ((XmlNode)val).SelectNodes("/Root")[0];
				if (val4 != null)
				{
					val4.AppendChild(val3);
				}
			}
			val2 = ((XmlNode)val).SelectNodes("/Root/SizeAPP")[0];
			if (val2 != null)
			{
				string text = control.Width.ToString();
				string text2 = control.Height.ToString();
				val2.InnerText = control.Location.X + "~" + control.Location.Y + "~" + text + "~" + text2;
				XmlTextWriter val5 = new XmlTextWriter("Config.xml", Encoding.UTF8);
				try
				{
					val5.Formatting = (Formatting)1;
					val.Save((XmlWriter)(object)val5);
					return;
				}
				finally
				{
					((IDisposable)val5)?.Dispose();
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}
}
