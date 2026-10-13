using System.IO;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Formatters.Binary;
using System.Xml.Serialization;

namespace HTKKSerializer;

public class Serializer
{
	public T Deserialize<T>(string input) where T : class
	{
		//IL_000b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0011: Expected O, but got Unknown
		XmlSerializer val = new XmlSerializer(typeof(T));
		using StringReader stringReader = new StringReader(input);
		return (T)val.Deserialize((TextReader)stringReader);
	}

	public string Serialize<T>(T ObjectToSerialize)
	{
		//IL_000e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0014: Expected O, but got Unknown
		XmlSerializer val = new XmlSerializer(ObjectToSerialize.GetType());
		using StringWriter stringWriter = new StringWriter();
		val.Serialize((TextWriter)stringWriter, (object)ObjectToSerialize);
		return stringWriter.ToString();
	}

	public static T DeepCopy<T>(T other)
	{
		using MemoryStream memoryStream = new MemoryStream();
		BinaryFormatter binaryFormatter = new BinaryFormatter();
		binaryFormatter.Context = new StreamingContext(StreamingContextStates.Clone);
		binaryFormatter.Serialize(memoryStream, other);
		memoryStream.Position = 0L;
		return (T)binaryFormatter.Deserialize(memoryStream);
	}
}
