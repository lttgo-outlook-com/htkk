using System;
using System.Collections.Generic;

namespace HTKKDanhmuc;

[Serializable]
public class PathDm
{
	public bool dmNew { get; set; }

	public string pathDataDm { get; set; }

	public string pathSaveInt { get; set; }

	public string pathSaveDataFile { get; set; }

	public List<string> pathDonGiaTinh { get; set; }

	public string pathSaveIntOld { get; set; }

	public string pathSaveDataFileOld { get; set; }

	public PathDm(string pDataDm, string pSaveInt, string pSaveDataFile, List<string> pathDGTinh)
	{
		dmNew = true;
		pathDataDm = pDataDm;
		pathSaveInt = pSaveInt;
		pathSaveDataFile = pSaveDataFile;
		pathDonGiaTinh = pathDGTinh;
	}

	public PathDm(string pSaveIntOld, string pSaveDataFileOld)
	{
		dmNew = false;
		pathSaveIntOld = pSaveIntOld;
		pathSaveDataFileOld = pSaveDataFileOld;
	}
}
