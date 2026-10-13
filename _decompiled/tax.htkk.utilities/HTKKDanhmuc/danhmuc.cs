using System;
using System.Collections.Generic;

namespace HTKKDanhmuc;

[Serializable]
public class danhmuc
{
	public string type { get; set; }

	public List<item> items { get; set; }
}
