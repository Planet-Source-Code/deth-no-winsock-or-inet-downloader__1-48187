VERSION 5.00
Begin VB.UserControl Download 
   BackStyle       =   0  'Transparent
   CanGetFocus     =   0   'False
   ClientHeight    =   885
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1260
   HasDC           =   0   'False
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   885
   ScaleWidth      =   1260
   ToolboxBitmap   =   "Download.ctx":0000
   Begin VB.Image Image1 
      Height          =   480
      Left            =   0
      Picture         =   "Download.ctx":0312
      Top             =   0
      Width           =   480
   End
End
Attribute VB_Name = "Download"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Event DownloadProgress(ByVal Key As String, ByVal Progress As Long, ByVal ProgressMax As Long, ByVal StatusNumber As AsyncStatusCodeConstants, ByVal StatusText As String)
Event DownloadError(ByVal Key As String, ByVal Code As Long, ByVal Description As String)
Event DownloadComplete(ByVal Key As String, ByVal Value As Variant, ByVal FileType As AsyncTypeConstants)

Sub BeginDownload(ByVal Path As String, Optional ByVal Key As String = vbNullString, Optional FileType As AsyncTypeConstants = vbAsyncTypeByteArray, Optional DownloadType As AsyncReadConstants = vbAsyncReadForceUpdate)

    On Error GoTo Err_Handle

    UserControl.AsyncRead Path, FileType, Key, DownloadType

Exit Sub

Err_Handle:
    RaiseEvent DownloadError(Key, Err.Number, Err.Description)

End Sub

Sub Cancel(Optional ByVal Key As String = vbNullString)

    On Error Resume Next
        UserControl.CancelAsyncRead Key

End Sub

Private Sub UserControl_AsyncReadComplete(AsyncProp As AsyncProperty)

    On Error Resume Next
      Dim CodeString As String

        With AsyncProp

            If .StatusCode = vbAsyncStatusCodeError Then
                If .Status = "" Then
                    CodeString = GetStatusCode(.StatusCode)
                  Else
                    CodeString = .Status
                End If
                RaiseEvent DownloadError(.PropertyName, .StatusCode, CodeString)
              Else
                RaiseEvent DownloadComplete(.PropertyName, .Value, .AsyncType)
            End If
        End With

End Sub

Private Sub UserControl_AsyncReadProgress(AsyncProp As AsyncProperty)

    On Error Resume Next
      Dim CodeString As String

        With AsyncProp

            If .Status = "" Then
                CodeString = GetStatusCode(.StatusCode)
              Else
                CodeString = .Status
            End If

            If .StatusCode = vbAsyncStatusCodeError Then
                RaiseEvent DownloadError(.PropertyName, .StatusCode, CodeString)
              Else
                RaiseEvent DownloadProgress(.PropertyName, .BytesRead, .BytesMax, .StatusCode, CodeString)
            End If
        End With

End Sub

Private Sub UserControl_Resize()

    On Error Resume Next
        UserControl.Width = Image1.Width
        UserControl.Height = UserControl.Height

End Sub

Function GetStatusCode(ByVal Code As AsyncStatusCodeConstants) As String

    If Code = vbAsyncStatusCodeBeginDownloadData Then
        GetStatusCode = "Download Initialized."
      ElseIf Code = vbAsyncStatusCodeBeginSyncOperation Then
        GetStatusCode = "Synchronous Download Has Started."
      ElseIf Code = vbAsyncStatusCodeCacheFileNameAvailable Then
        GetStatusCode = "Local Cache File Is Available."
      ElseIf Code = vbAsyncStatusCodeConnecting Then
        GetStatusCode = "Connecting To Resource."
      ElseIf Code = vbAsyncStatusCodeDownloadingData Then
        GetStatusCode = "Download In Progress."
      ElseIf Code = vbAsyncStatusCodeEndDownloadData Then
        GetStatusCode = "Download Complete."
      ElseIf Code = vbAsyncStatusCodeEndSyncOperation Then
        GetStatusCode = "Synchronous Download Complete."
      ElseIf Code = vbAsyncStatusCodeError Then
        GetStatusCode = "An Error Has Occurred."
      ElseIf Code = vbAsyncStatusCodeFindingResource Then
        GetStatusCode = "Finding Resource."
      ElseIf Code = vbAsyncStatusCodeMIMETypeAvailable Then
        GetStatusCode = "MIME Type Is Available."
      ElseIf Code = vbAsyncStatusCodeRedirecting Then
        GetStatusCode = "Redirecting."
      ElseIf Code = vbAsyncStatusCodeSendingRequest Then
        GetStatusCode = "Sending Request."
      ElseIf Code = vbAsyncStatusCodeUsingCachedCopy Then
        GetStatusCode = "Using Cached Copy."
      Else
        GetStatusCode = "Unknown."
    End If

End Function

