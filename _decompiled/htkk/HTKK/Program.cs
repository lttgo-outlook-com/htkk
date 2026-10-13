using System;
using System.Diagnostics;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using HTKK.Forms;
using TAX.HTKK.Utilities.Core;

namespace HTKK;

internal static class Program
{
	private const int SPI_SETKEYBOARDCUES = 4107;

	private const int SPIF_SENDWININICHANGE = 2;

	[DllImport("user32.dll")]
	private static extern int SystemParametersInfo(int uAction, int uParam, int lpvParam, int fuWinIni);

	public static Process PriorProcess()
	{
		Process currentProcess = Process.GetCurrentProcess();
		Process[] processesByName = Process.GetProcessesByName(currentProcess.ProcessName);
		Process[] array = processesByName;
		foreach (Process process in array)
		{
			if (process.Id != currentProcess.Id && process.MainModule.FileName == currentProcess.MainModule.FileName)
			{
				return process;
			}
		}
		return null;
	}

	[STAThread]
	private static void Main(string[] args)
	{
		//IL_003d: Unknown result type (might be due to invalid IL or missing references)
		HTKKFile.LogRun("APP START", "--------------------------------------------   START   ------------------------------------------------");
		if (PriorProcess() != null && args.Length == 0)
		{
			MessageBox.Show("Bạn đang chạy ứng dụng này rồi", "Thông báo", (MessageBoxButtons)0, (MessageBoxIcon)16, (MessageBoxDefaultButton)0, (MessageBoxOptions)131072);
			return;
		}
		Thread.CurrentThread.CurrentCulture = CultureInfo.InvariantCulture;
		Application.EnableVisualStyles();
		Application.SetCompatibleTextRenderingDefault(false);
		HTKKWinDictionaryDiaDanh hTKKWinDictionaryDiaDanh = new HTKKWinDictionaryDiaDanh();
		SystemParametersInfo(4107, 0, 1, 0);
		Application.Run((Form)(object)new frmFormStart());
		SystemParametersInfo(4107, 0, 0, 0);
	}
}
