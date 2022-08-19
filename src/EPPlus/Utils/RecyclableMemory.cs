using System.IO;
using System.Threading;

namespace OfficeOpenXml.Utils
{
	/// <summary>
	/// Handles the Recyclable Memory stream for supported and unsupported target frameworks.
	/// </summary>
	public static class RecyclableMemory
	{
		private static Microsoft.IO.RecyclableMemoryStreamManager _memoryManager;
		private static bool _dataInitialized = false;
		private static object _dataLock = new object();

		private static Microsoft.IO.RecyclableMemoryStreamManager MemoryManager
		{
			get
			{
				return LazyInitializer.EnsureInitialized(ref _memoryManager, ref _dataInitialized, ref _dataLock);
			}
		}

		/// <summary>
		/// Sets the RecyclableMemorytreamsManager to manage pools
		/// </summary>
		/// <param name="recyclableMemoryStreamManager">The memory manager</param>
		public static void SetRecyclableMemoryStreamManager(Microsoft.IO.RecyclableMemoryStreamManager recyclableMemoryStreamManager)
		{
			_dataInitialized = recyclableMemoryStreamManager is object;
			_memoryManager = recyclableMemoryStreamManager;
		}
		/// <summary>
		/// Get a new memory stream.
		/// </summary>
		/// <returns>A MemoryStream</returns>
		internal static MemoryStream GetStream()
		{

			return MemoryManager.GetStream();
		}

		/// <summary>
		/// Get a new memory stream initiated with a byte-array
		/// </summary>
		/// <returns>A MemoryStream</returns>
		internal static MemoryStream GetStream(byte[] array)
		{
			return MemoryManager.GetStream(array);
		}
		/// <summary>
		/// Get a new memory stream initiated with a byte-array
		/// </summary>
		/// <param name="capacity">The initial size of the internal array</param>
		/// <returns>A MemoryStream</returns>
		internal static MemoryStream GetStream(int capacity)
		{
			return MemoryManager.GetStream(null, capacity);
		}
	}
}
