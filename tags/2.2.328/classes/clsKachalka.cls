VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsKachalka"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' To Form!!! --->  Private WithEvents Kachalka As clsKachalka

Event Authenticate(hWnd As Long, Username As String, Password As String)
Event Complete(ByVal Status As kach_tlb.BINDSTATUS, ByVal StatusText As String)
Event Progress(ByVal Progress As Long, ByVal ProgressMax As Long, ByVal Status As Long, ByVal StatusText As String, Cancel As Boolean)
Event DataAvailable(ByVal EventType As kach_tlb.BSCF, ByVal Data As String, ByVal DataFormat As Long)
Event GetOptions(Options As kach_tlb.BINDF)

Event Error(ErrorCode, ErrorString)


Public Method As kach_tlb.BINDVERB
Public RequestData As String


Implements IBindStatusCallback
Implements IAuthenticate

Private Offset As Long
Private ReturnCode(11) As Byte

Private Declare Function GetDesktopWindow Lib "user32" () As Long

Private Sub IAuthenticate_Authenticate(hWnd As Long, pszUsername As Long, pszPassword As Long)
Dim Username As String, Password As String
    hWnd = GetDesktopWindow
    RaiseEvent Authenticate(hWnd, Username, Password)
    If hWnd = 0 Then
        pszUsername = StrPtr(Username)
        PutMem4 ByVal VarPtr(Username), 0
        pszPassword = StrPtr(Password)
        PutMem4 ByVal VarPtr(Password), 0
    End If
End Sub

Private Sub IBindStatusCallback_GetBindInfo(grfBINDF As kach_tlb.BINDF, pbindinfo As kach_tlb.BINDINFO)
    grfBINDF = BINDF_RESYNCHRONIZE Or BINDF_GETFROMCACHE_IF_NET_FAIL
    RaiseEvent GetOptions(grfBINDF)
    pbindinfo.dwBindVerb = Me.Method
    If StrPtr(Me.RequestData) Then
        pbindinfo.stgmedData.TYMED = TYMED_ISTREAM
        Set pbindinfo.stgmedData.pstm = CreateStreamOnHGlobal(0, 1)
        pbindinfo.cbstgmedData = Len(Me.RequestData)
        pbindinfo.stgmedData.pstm.Write Me.RequestData, pbindinfo.cbstgmedData
        pbindinfo.stgmedData.pstm.Seek 0, STREAM_SEEK_SET
    End If
End Sub

Private Function IBindStatusCallback_GetPriority() As Long
Stop 'афаик не нужен
End Function

Private Sub IBindStatusCallback_OnDataAvailable(ByVal grfBSCF As kach_tlb.BSCF, ByVal dwSize As Long, pformatetc As kach_tlb.FORMATETC, pstgmed As kach_tlb.STGMEDIUM)
    dwSize = dwSize - Offset
Dim buf As String: buf = String(dwSize, 0)
    If pstgmed.TYMED <> TYMED_ISTREAM Then Error 13
    If dwSize <> pstgmed.pstm.Read(buf, dwSize) Then Error 7
    RaiseEvent DataAvailable(grfBSCF, buf, VarPtr(pformatetc))
    Offset = Offset + dwSize
End Sub

Private Sub IBindStatusCallback_OnLowResource(ByVal Reserved As Long)
'не нужен
End Sub

Private Sub IBindStatusCallback_OnObjectAvailable(riid As kach_tlb.UUID, ByVal pUnk As stdole.IUnknown)
'не нужен
End Sub

Private Sub IBindStatusCallback_OnProgress(ByVal ulProgress As Long, ByVal ulProgressMax As Long, ByVal ulStatusCode As kach_tlb.BINDSTATUS, ByVal szStatusText As Long)
Dim Cancel As Boolean
RaiseEvent Progress(ulProgress, ulProgressMax, ulStatusCode, SysAllocString(szStatusText), Cancel)
If Cancel Then 'whoo-hoo
    Dim RetAddr As Long, pRetAddr As Long
    pRetAddr = VarPtr(ulProgress) - 8
    PutMem4 ReturnCode(1), E_ABORT
    GetMem4 ByVal pRetAddr, RetAddr
    PutMem4 ByVal pRetAddr, VarPtr(ReturnCode(0))
    PutMem4 ReturnCode(6), RetAddr
End If
End Sub

Private Sub IBindStatusCallback_OnStartBinding(ByVal dwReserved As Long, ByVal pib As kach_tlb.IBinding)
Offset = 0
End Sub

Private Sub IBindStatusCallback_OnStopBinding(ByVal hrStatus As Long, ByVal szStatusText As Long)
RaiseEvent Complete(hrStatus, SysAllocString(szStatusText))
End Sub




Public Sub DownloadToFile(ByVal URL As String, ByVal FileName As String)
Dim Result As Long
    Result = URLDownloadToFile(Nothing, URL, FileName, 0, Me)
    If Result Then RaiseEvent Error(0, Result): Exit Sub
End Sub

Public Function DownloadToString(ByVal URL As String) As String
Dim Result As Long, pStream As IStream

''' ADDED BY RG
If InStr(Command, "/offline") Then DownloadToString = "offline": Exit Function
''' END ADDED

    Result = URLOpenBlockingStream(Nothing, URL, pStream, 0, Me)
    If Result Then RaiseEvent Error(0, Result): Exit Function
    Dim S As STATSTG
    pStream.stat S, STATFLAG_NONAME
    If S.cbSizeHi Then Error 7
    DownloadToString = String(S.cbSize, 0)
    If S.cbSize <> pStream.Read(DownloadToString, S.cbSize) Then Error 7
End Function

Public Sub DownloadPush(ByVal URL As String)
Dim Result As Long, pStream As IStream
    Result = URLOpenStream(Nothing, URL, 0, Me)
    If Result Then RaiseEvent Error(0, Result): Exit Sub
End Sub

Private Sub Class_Initialize()
    ReturnCode(0) = &HB8
    ReturnCode(5) = &HBA
    ReturnCode(10) = &HFF
    ReturnCode(11) = &HE2
End Sub

Function LookupStatus(ByVal ulStatusCode As kach_tlb.BINDSTATUS) As String
If ulStatusCode <= 0 Then
    LookupStatus = Hex(ulStatusCode)
Else
    LookupStatus = Choose(ulStatusCode, _
        "BINDSTATUS_FINDINGRESOURCE", "BINDSTATUS_CONNECTING", _
        "BINDSTATUS_REDIRECTING", "BINDSTATUS_BEGINDOWNLOADDATA", _
        "BINDSTATUS_DOWNLOADINGDATA", "BINDSTATUS_ENDDOWNLOADDATA", _
        "BINDSTATUS_BEGINDOWNLOADCOMPONENTS", "BINDSTATUS_INSTALLINGCOMPONENTS", _
        "BINDSTATUS_ENDDOWNLOADCOMPONENTS", "BINDSTATUS_USINGCACHEDCOPY", _
        "BINDSTATUS_SENDINGREQUEST", "BINDSTATUS_CLASSIDAVAILABLE", _
        "BINDSTATUS_MIMETYPEAVAILABLE", "BINDSTATUS_CACHEFILENAMEAVAILABLE", _
        "BINDSTATUS_BEGINSYNCOPERATION", "BINDSTATUS_ENDSYNCOPERATION", _
        "BINDSTATUS_BEGINUPLOADDATA", "BINDSTATUS_UPLOADINGDATA", _
        "BINDSTATUS_ENDUPLOADDATA", "BINDSTATUS_PROTOCOLCLASSID", _
        "BINDSTATUS_ENCODING", "BINDSTATUS_VERIFIEDMIMETYPEAVAILABLE", _
        "BINDSTATUS_CLASSINSTALLLOCATION", "BINDSTATUS_DECODING", _
        "BINDSTATUS_LOADINGMIMEHANDLER", "BINDSTATUS_CONTENTDISPOSITIONATTACH", _
        "BINDSTATUS_FILTERREPORTMIMETYPE", "BINDSTATUS_CLSIDCANINSTANTIATE", _
        "BINDSTATUS_IUNKNOWNAVAILABLE", "BINDSTATUS_DIRECTBIND", _
        "BINDSTATUS_RAWMIMETYPE", "BINDSTATUS_PROXYDETECTING", _
        "BINDSTATUS_ACCEPTRANGES", "BINDSTATUS_COOKIE_SENT", _
        "BINDSTATUS_COMPACT_POLICY_RECEIVED", "BINDSTATUS_COOKIE_SUPPRESSED", _
        "BINDSTATUS_COOKIE_STATE_UNKNOWN", "BINDSTATUS_COOKIE_STATE_ACCEPT", _
        "BINDSTATUS_COOKIE_STATE_REJECT", "BINDSTATUS_COOKIE_STATE_PROMPT", _
        "BINDSTATUS_COOKIE_STATE_LEASH", "BINDSTATUS_COOKIE_STATE_DOWNGRADE", _
        "BINDSTATUS_POLICY_HREF", "BINDSTATUS_P3P_HEADER", _
        "BINDSTATUS_SESSION_COOKIE_RECEIVED", "BINDSTATUS_PERSISTENT_COOKIE_RECEIVED", _
        "BINDSTATUS_SESSION_COOKIES_ALLOWED")
End If
End Function


