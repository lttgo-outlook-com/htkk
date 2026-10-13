namespace TAX.HTKK.Utilities.Implement;

public class HTKKFactory
{
	private static HTKKHeaderImpl _HTKKHeaderImpl;

	private static HTKKToKhaiImpl _HTKKToKhaiImpl;

	private static HTKKSheetImpl _HTKKSheetImpl;

	private static HTKKTraCuuImpl _HTKKTraCuuImpl;

	public static HTKKHeaderImpl GetInstanceHTKKHeader()
	{
		return _HTKKHeaderImpl ?? (_HTKKHeaderImpl = new HTKKHeaderImpl());
	}

	public static HTKKToKhaiImpl GetInstanceHTKKToKhai()
	{
		return _HTKKToKhaiImpl ?? (_HTKKToKhaiImpl = new HTKKToKhaiImpl());
	}

	public static HTKKSheetImpl GetInstanceHTKKSheet()
	{
		return _HTKKSheetImpl ?? (_HTKKSheetImpl = new HTKKSheetImpl());
	}

	public static HTKKTraCuuImpl GetInstanceHTKKTraCuu()
	{
		return _HTKKTraCuuImpl ?? (_HTKKTraCuuImpl = new HTKKTraCuuImpl());
	}
}
