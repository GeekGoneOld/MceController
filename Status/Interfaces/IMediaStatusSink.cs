using System.Runtime.InteropServices;

[ComImport, Guid("075FC453-F236-41DA-B90D-9FBB8BBDC101"), InterfaceType(ComInterfaceType.InterfaceIsDual)]
internal interface IMediaStatusSink {
	[DispId(1)]
	void Initialize();

	[DispId(2)]
	IMediaStatusSession CreateSession();
}