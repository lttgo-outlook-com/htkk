using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Reflection;
using System.Xml;
using System.Xml.Serialization;

namespace TAX.HTKK.Utilities.Core;

public class HTKKBase<T> where T : new()
{
	public static void XmlToObject(string xmlContent, ref T info)
	{
		//IL_0002: Unknown result type (might be due to invalid IL or missing references)
		//IL_0008: Expected O, but got Unknown
		try
		{
			XmlDocument val = new XmlDocument();
			val.LoadXml(xmlContent);
			XmlNodeList elementsByTagName = val.GetElementsByTagName("Cell");
			Dictionary<string, string> dictionary = new Dictionary<string, string>();
			string value = ((XmlNode)val.GetElementsByTagName("Sections")[0].Attributes["Version"]).Value;
			for (int i = 0; i < elementsByTagName.Count; i++)
			{
				string key = ((XmlNode)elementsByTagName[i].Attributes["CellID"]).Value + "," + ((XmlNode)elementsByTagName[i].Attributes["CellID2"]).Value;
				string value2 = ((XmlNode)elementsByTagName[i].Attributes["Value"]).Value;
				dictionary.Add(key, value2);
			}
			Type typeFromHandle = typeof(T);
			PropertyInfo[] properties = typeFromHandle.GetProperties();
			PropertyInfo[] array = properties;
			foreach (PropertyInfo propertyInfo in array)
			{
				if (propertyInfo.Name == "Version")
				{
					propertyInfo.SetValue(info, value, null);
				}
				else if ((object)propertyInfo.PropertyType == typeof(Dictionary<string, string>))
				{
					propertyInfo.SetValue(info, dictionary, null);
				}
				else
				{
					propertyInfo.SetValue(info, string.Empty, null);
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static string ObjectToXml(T info)
	{
		//IL_0030: Unknown result type (might be due to invalid IL or missing references)
		//IL_0036: Expected O, but got Unknown
		//IL_003e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0045: Expected O, but got Unknown
		try
		{
			string result = string.Empty;
			if (info == null)
			{
				return string.Empty;
			}
			XmlSerializer val = new XmlSerializer(typeof(T));
			using (StringWriter stringWriter = new StringWriter())
			{
				XmlWriterSettings val2 = new XmlWriterSettings();
				val2.Indent = true;
				XmlWriter val3 = XmlWriter.Create((TextWriter)stringWriter, val2);
				try
				{
					val.Serialize(val3, (object)info);
					result = stringWriter.ToString();
				}
				finally
				{
					((IDisposable)val3)?.Dispose();
				}
			}
			return result;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return string.Empty;
		}
	}

	public static T XmlToObject(string xmlContent)
	{
		//IL_0014: Unknown result type (might be due to invalid IL or missing references)
		//IL_001a: Expected O, but got Unknown
		T result = default(T);
		try
		{
			XmlSerializer val = new XmlSerializer(typeof(T));
			using TextReader textReader = new StringReader(xmlContent);
			result = (T)val.Deserialize(textReader);
			return result;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return result;
	}

	public static List<T> ToCollection(DataTable dt)
	{
		List<T> list = new List<T>();
		try
		{
			Type typeFromHandle = typeof(T);
			PropertyInfo[] properties = typeFromHandle.GetProperties();
			List<DataColumn> list2 = new List<DataColumn>();
			foreach (DataRow row in dt.Rows)
			{
				foreach (DataColumn column in dt.Columns)
				{
					list2.Add(column);
				}
			}
			foreach (DataRow row2 in dt.Rows)
			{
				T val = (T)Activator.CreateInstance(typeFromHandle);
				PropertyInfo[] array = properties;
				foreach (PropertyInfo pc in array)
				{
					try
					{
						DataColumn dataColumn = list2.Find((DataColumn c) => c.ColumnName == pc.Name);
						if (dataColumn != null)
						{
							pc.SetValue(val, row2[pc.Name], null);
						}
					}
					catch
					{
					}
				}
				list.Add(val);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return list;
	}

	public static DataTable ToDataTable(List<T> data)
	{
		PropertyDescriptorCollection properties = TypeDescriptor.GetProperties(typeof(T));
		object[] array = new object[properties.Count];
		using DataTable dataTable = new DataTable();
		try
		{
			long num = properties.Count;
			for (int i = 0; i < num; i++)
			{
				PropertyDescriptor propertyDescriptor = properties[i];
				dataTable.Columns.Add(propertyDescriptor.Name, propertyDescriptor.PropertyType);
			}
			foreach (T datum in data)
			{
				long num2 = array.Length;
				for (int i = 0; i < num2; i++)
				{
					array[i] = properties[i].GetValue(datum);
				}
				dataTable.Rows.Add(array);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return dataTable;
	}

	public static DataSet ToDataSet(List<T> list)
	{
		Type typeFromHandle = typeof(T);
		using DataSet dataSet = new DataSet();
		using (DataTable dataTable = new DataTable())
		{
			try
			{
				dataSet.Tables.Add(dataTable);
				PropertyInfo[] properties = typeFromHandle.GetProperties();
				PropertyInfo[] array = properties;
				foreach (PropertyInfo propertyInfo in array)
				{
					Type propertyType = propertyInfo.PropertyType;
					Type type = Nullable.GetUnderlyingType(propertyType) ?? propertyType;
					dataTable.Columns.Add(propertyInfo.Name, type);
				}
				foreach (T item in list)
				{
					DataRow dataRow = dataTable.NewRow();
					array = properties;
					foreach (PropertyInfo propertyInfo in array)
					{
						dataRow[propertyInfo.Name] = propertyInfo.GetValue(item, null) ?? DBNull.Value;
					}
					dataTable.Rows.Add(dataRow);
				}
			}
			catch (Exception ex)
			{
				HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			}
		}
		return dataSet;
	}
}
