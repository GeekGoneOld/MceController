using System.Runtime.InteropServices;
using VmcController.MceState;

[ComImport, Guid("A70D81F2-C9D2-4053-AF0E-CDEA39BDD1AD"), InterfaceType(ComInterfaceType.InterfaceIsDual)]
public interface IMediaStatusSession {
	[DispId(1)]
	void MediaStatusChange(MediaState.MEDIASTATUSPROPERTYTAG[] Tags, object[] Properties);

	[DispId(2)]
	void Close();
}