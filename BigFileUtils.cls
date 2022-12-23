VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BigFileUtils"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'==================================================================================================
' Class to support read/write of files >2GB
' Based on http://support.microsoft.com/kb/189981
'==================================================================================================
Option Explicit

Public Enum W32F_Errors
  W32F_UNKNOWN_ERROR = 45600
  W32F_FILE_ALREADY_OPEN
  W32F_PROBLEM_OPENING_FILE
  W32F_FILE_ALREADY_CLOSED
  W32F_Problem_seeking
End Enum

Private Const W32F_SOURCE = "Win32File Object"

Private Const GENERIC_WRITE = &H40000000
Private Const GENERIC_READ = &H80000000
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const CREATE_ALWAYS = 2
Private Const OPEN_ALWAYS = 4
Private Const INVALID_HANDLE_VALUE = -1

Private Const FILE_SHARE_DELETE = 4
Private Const FILE_SHARE_READ = 1
Private Const FILE_SHARE_WRITE = 2


Private Const FILE_BEGIN = 0, FILE_CURRENT = 1, FILE_END = 2

Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000

Private Declare Function FormatMessage Lib "kernel32" _
        Alias "FormatMessageA" (ByVal dwFlags As Long, _
                                lpSource As Long, _
                                ByVal dwMessageId As Long, _
                                ByVal dwLanguageId As Long, _
                                ByVal lpBuffer As String, _
                                ByVal nSize As Long, _
                                Arguments As Any) As Long

Private Declare Function ReadFile Lib "kernel32" _
                        (ByVal hFile As Long, _
                         lpBuffer As Any, _
                         ByVal nNumberOfBytesToRead As Long, _
                         lpNumberOfBytesRead As Long, _
                         ByVal lpOverlapped As Long) As Long

Private Declare Function CloseHandle Lib "kernel32" _
                        (ByVal hObject As Long) As Long

Private Declare Function WriteFile Lib "kernel32" _
                        (ByVal hFile As Long, _
                         lpBuffer As Any, _
                         ByVal nNumberOfBytesToWrite As Long, _
                         lpNumberOfBytesWritten As Long, _
                         ByVal lpOverlapped As Long) As Long

Private Declare Function CreateFile Lib "kernel32" _
        Alias "CreateFileA" (ByVal lpFileName As String, _
                             ByVal dwDesiredAccess As Long, _
                             ByVal dwShareMode As Long, _
                             ByVal lpSecurityAttributes As Long, _
                             ByVal dwCreationDisposition As Long, _
                             ByVal dwFlagsAndAttributes As Long, _
                             ByVal hTemplateFile As Long) As Long

Private Declare Function SetFilePointer Lib "kernel32" _
                        (ByVal hFile As Long, _
                         ByVal lDistanceToMove As Long, _
                         lpDistanceToMoveHigh As Long, _
                         ByVal dwMoveMethod As Long) As Long

Private Declare Function FlushFileBuffers Lib "kernel32" _
                        (ByVal hFile As Long) As Long

Private hFile As Long                                   '<- File handle of open file
Private sFName As String                                '<- Open file name
Private fAutoFlush As Boolean                           '<- Auto flush flag


Private Type ByteLong                                   '<- Long stored in byte array
    Value(3) As Byte
End Type
 
Private Type TypedLong                                  '<- Long type for Byte array conversion
    Value As Long
End Type

Private Type ByteInteger                            '<- Integer stored in byte array
    Value(2) As Byte
End Type
 
Private Type TypedInteger                           '<- Integer type for Byte array conversion
    Value As Integer
End Type


'==================================================================================================
' Get the currently assigned file handle
'==================================================================================================
Public Property Get FileHandle() As Long
  If hFile = INVALID_HANDLE_VALUE Then
    RaiseError W32F_FILE_ALREADY_CLOSED
  End If
  FileHandle = hFile
End Property

'==================================================================================================
' Get the currently open file name
'==================================================================================================
Public Property Get FileName() As String
  If hFile = INVALID_HANDLE_VALUE Then
    RaiseError W32F_FILE_ALREADY_CLOSED
  End If
  FileName = sFName
End Property

'==================================================================================================
' Returns a boolean to indicate whether the file is open.
'==================================================================================================
Public Property Get IsOpen() As Boolean
  IsOpen = hFile <> INVALID_HANDLE_VALUE
End Property

'==================================================================================================
' Flag to enable/disable autoflush
'==================================================================================================
Public Property Get AutoFlush() As Boolean
  If hFile = INVALID_HANDLE_VALUE Then
    RaiseError W32F_FILE_ALREADY_CLOSED
  End If
  AutoFlush = fAutoFlush
End Property

Public Property Let AutoFlush(ByVal NewVal As Boolean)
  If hFile = INVALID_HANDLE_VALUE Then
    RaiseError W32F_FILE_ALREADY_CLOSED
  End If
  fAutoFlush = NewVal
End Property

'==================================================================================================
' Opens the file specified by the sFileName argument.
' File name can be relative or rooted
'==================================================================================================
Public Sub OpenFile(ByVal sFileName As String)
  If hFile <> INVALID_HANDLE_VALUE Then
    RaiseError W32F_FILE_ALREADY_OPEN, sFName
  End If
  
  hFile = CreateFile(sFileName, GENERIC_WRITE Or GENERIC_READ, _
                     FILE_SHARE_DELETE Or FILE_SHARE_READ Or FILE_SHARE_WRITE, _
                     0, OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL, 0)
                     
  If hFile = INVALID_HANDLE_VALUE Then
    RaiseError W32F_PROBLEM_OPENING_FILE, sFileName
  End If
  
  sFName = sFileName
End Sub

'==================================================================================================
' Closes the currently open file.
'==================================================================================================
Public Sub CloseFile()
  If hFile = INVALID_HANDLE_VALUE Then
    RaiseError W32F_FILE_ALREADY_CLOSED
  End If
  CloseHandle hFile
  sFName = ""
  fAutoFlush = False
  hFile = INVALID_HANDLE_VALUE
End Sub

'==================================================================================================
' Reads ByteCount bytes and returns them in a Variant byte array and moves the pointer.
'==================================================================================================
Public Function ReadBytes(ByVal ByteCount As Long) As Variant

    Dim BytesRead As Long, Bytes() As Byte
    Dim DataBytes() As Byte
    
    If hFile = INVALID_HANDLE_VALUE Then
        RaiseError W32F_FILE_ALREADY_CLOSED
    End If
    
    ReDim DataBytes(0 To ByteCount - 1)
    ReadFile hFile, DataBytes(0), ByteCount, BytesRead, 0
    ReadBytes = DataBytes
    
End Function

'==================================================================================================
' Reads a single Byte and returns it and moves the pointer.
'==================================================================================================
Public Function ReadByte() As Byte
    Dim DataByte As Byte                            '<- Byte value read
    Dim BytesRead As Long                           '<- Number of bytes read
    
    If hFile = INVALID_HANDLE_VALUE Then
        RaiseError W32F_FILE_ALREADY_CLOSED
    End If
    
    ReadFile hFile, DataByte, 1, BytesRead, 0
    ReadByte = DataByte
    
End Function

'==================================================================================================
' Reads a 2 byte integer and returns it and moves the pointer.
'==================================================================================================
Public Function ReadInteger() As Integer
    Dim DataBytes(1) As Byte                         '<- 2 Byte values read
    Dim DataVal As Integer                          '<- Resulting integer
    Dim BytesRead As Long                           '<- Number of bytes read
    
    If hFile = INVALID_HANDLE_VALUE Then
        RaiseError W32F_FILE_ALREADY_CLOSED
    End If
    
    ReadFile hFile, DataBytes(0), 2, BytesRead, 0
    
    DataVal = ByteArrayToInteger(DataBytes)
    
    ReadInteger = DataVal
    
End Function

'==================================================================================================
' Reads a 4 byte long and returns it and moves the pointer.
'==================================================================================================
Public Function ReadLong() As Long
    Dim DataBytes(3) As Byte                        '<- 2 Byte values read
    Dim DataVal As Long                             '<- Resulting long
    Dim BytesRead As Long                           '<- Number of bytes read
    
    If hFile = INVALID_HANDLE_VALUE Then
        RaiseError W32F_FILE_ALREADY_CLOSED
    End If
    
    ReadFile hFile, DataBytes(0), 4, BytesRead, 0
    
    DataVal = ByteArrayToLong(DataBytes)

    ReadLong = DataVal
    
End Function

'==================================================================================================
' Reads a N byte string and returns it and moves the pointer.
'==================================================================================================
Public Function ReadString(length As Long) As String
    Dim DataBytes() As Byte                         '<- N Byte values read
    Dim DataVal As Integer                          '<- Resulting integer
    Dim BytesRead As Long                           '<- Number of bytes read
    
    If hFile = INVALID_HANDLE_VALUE Then
        RaiseError W32F_FILE_ALREADY_CLOSED
    End If
    
    ReDim DataBytes(length - 1)
    
    ReadFile hFile, DataBytes(0), length, BytesRead, 0
    
    ReadString = StrConv(DataBytes, vbUnicode)
    
End Function


'==================================================================================================
' Writes the contents of the byte array to the current position in the file and moves the pointer.
'==================================================================================================
Public Sub WriteBytes(DataBytes() As Byte)
Dim fSuccess As Long, BytesToWrite As Long, BytesWritten As Long
  If hFile = INVALID_HANDLE_VALUE Then
    RaiseError W32F_FILE_ALREADY_CLOSED
  End If
  BytesToWrite = UBound(DataBytes) - LBound(DataBytes) + 1
  fSuccess = WriteFile(hFile, DataBytes(LBound(DataBytes)), _
                       BytesToWrite, BytesWritten, 0)
  If fAutoFlush Then Flush
End Sub

'==================================================================================================
' Writes a single byte value to the current position in the file and move the pointer.
'==================================================================================================
Public Sub WriteByte(DataValue As Byte)
    Dim fSuccess As Long
    Dim BytesToWrite As Long
    Dim BytesWritten As Long
    
    '___ Error check ______________________________________________________________________________
    If hFile = INVALID_HANDLE_VALUE Then
        RaiseError W32F_FILE_ALREADY_CLOSED
    End If
  
    '___ Convert the data _________________________________________________________________________
    BytesToWrite = 1
    
    '___ Write the data ___________________________________________________________________________
    fSuccess = WriteFile(hFile, DataValue, _
                         BytesToWrite, BytesWritten, 0)
                         
    If fAutoFlush Then Flush
    
End Sub

'==================================================================================================
' Writes an integer value to the current position in the file and move the pointer.
'==================================================================================================
Public Sub WriteInteger(DataValue As Integer)
    Dim fSuccess As Long
    Dim BytesToWrite As Long
    Dim BytesWritten As Long
    Dim DataBytes() As Byte                        '<- Byte data for integer
    
    '___ Error check ______________________________________________________________________________
    If hFile = INVALID_HANDLE_VALUE Then
        RaiseError W32F_FILE_ALREADY_CLOSED
    End If
  
    '___ Convert the data _________________________________________________________________________
    DataBytes = IntegerToByteArray(DataValue)
    BytesToWrite = 2
    
    '___ Write the data ___________________________________________________________________________
    fSuccess = WriteFile(hFile, DataBytes(LBound(DataBytes)), _
                         BytesToWrite, BytesWritten, 0)
                         
    If fAutoFlush Then Flush
    
End Sub

'==================================================================================================
' Writes a long value to the current position in the file and move the pointer.
'==================================================================================================
Public Sub WriteLong(DataValue As Long)
    Dim fSuccess As Long
    Dim BytesToWrite As Long
    Dim BytesWritten As Long
    Dim DataBytes() As Byte                        '<- Byte data for long
    
    '___ Error check ______________________________________________________________________________
    If hFile = INVALID_HANDLE_VALUE Then
        RaiseError W32F_FILE_ALREADY_CLOSED
    End If
  
    '___ Convert the data _________________________________________________________________________
    DataBytes = LongToByteArray(DataValue)
    BytesToWrite = 4
    
    '___ Write the data ___________________________________________________________________________
    fSuccess = WriteFile(hFile, DataBytes(LBound(DataBytes)), _
                         BytesToWrite, BytesWritten, 0)
                         
    If fAutoFlush Then Flush
    
End Sub

'==================================================================================================
' Writes a string value to the current position in the file and move the pointer.
'==================================================================================================
Public Sub WriteString(DataValue As String)
    Dim fSuccess As Long
    Dim BytesToWrite As Long
    Dim BytesWritten As Long
    Dim DataBytes() As Byte                         '<- Byte data for string
    
    '___ Error check ______________________________________________________________________________
    If hFile = INVALID_HANDLE_VALUE Then
        RaiseError W32F_FILE_ALREADY_CLOSED
    End If
  
    '___ Convert the data _________________________________________________________________________
    DataBytes = StrConv(DataValue, vbFromUnicode)
    BytesToWrite = UBound(DataBytes) + 1
    
    '___ Write the data ___________________________________________________________________________
    fSuccess = WriteFile(hFile, DataBytes(LBound(DataBytes)), _
                         BytesToWrite, BytesWritten, 0)
                         
    If fAutoFlush Then Flush
    
End Sub

'==================================================================================================
' Forces Windows to flush the write cache.
'==================================================================================================
Public Sub Flush()
  If hFile = INVALID_HANDLE_VALUE Then
    RaiseError W32F_FILE_ALREADY_CLOSED
  End If
  FlushFileBuffers hFile
End Sub

'==================================================================================================
' Get the current absolute file pointer location.  Returns the location as two DWORDs which
' represent the higher 4 bytes and lower 4 bytes of the address.  VBA treats there 4 byte
' DWORDs as signed longs so a negative value on the lower represents a location between
' 2G (2,147,483,648) and 4G (4,294,967,296).
'==================================================================================================
Public Sub GetAbsolute(ByRef rtnHighPos As Long, ByRef rtnLowPos As Long)
    Dim LowPos As Long                              '<- Lower 4 bytes of position
    
    If hFile = INVALID_HANDLE_VALUE Then
        RaiseError W32F_FILE_ALREADY_CLOSED
    End If
  
    LowPos = 0
    rtnHighPos = 0
  
    rtnLowPos = SetFilePointer(hFile, LowPos, rtnHighPos, FILE_CURRENT)
  
End Sub

'==================================================================================================
' Moves the file pointer to the designated position from the beginning of the file. Though
' VBA treats the DWORDS as signed values, the API treats them as unsigned. Make the high-order
' argument non-zero to exceed 4GB. The low-order DWORD will be negative for values between
' 2GB and 4GB.
'==================================================================================================
Public Sub SeekAbsolute(ByVal HighPos As Long, ByVal LowPos As Long)
  If hFile = INVALID_HANDLE_VALUE Then
    RaiseError W32F_FILE_ALREADY_CLOSED
  End If
  LowPos = SetFilePointer(hFile, LowPos, HighPos, FILE_BEGIN)
End Sub

'==================================================================================================
' Moves the file pointer up to +/- 2GB (2,147,483,648) from the current
' location. You can rewrite this method to allow for
' offsets greater than 2GB by converting a 64-bit signed
' offset into two 32-bit values.
'==================================================================================================
Public Sub SeekRelative(ByVal offset As Long)
Dim TempLow As Long, TempErr As Long
  If hFile = INVALID_HANDLE_VALUE Then
    RaiseError W32F_FILE_ALREADY_CLOSED
  End If
  TempLow = SetFilePointer(hFile, offset, ByVal 0&, FILE_CURRENT)
  If TempLow = -1 Then
    TempErr = Err.LastDllError
    If TempErr Then
      RaiseError W32F_Problem_seeking, "Error " & TempErr & "." & _
                                       vbCrLf & CStr(TempErr)
    End If
  End If
End Sub


'==================================================================================================
' Init this object.
'==================================================================================================
Private Sub Class_Initialize()
  hFile = INVALID_HANDLE_VALUE
End Sub

'==================================================================================================
' Release this object.
'==================================================================================================
Private Sub Class_Terminate()
  If hFile <> INVALID_HANDLE_VALUE Then CloseHandle hFile
End Sub

'==================================================================================================
' Raise an error
'==================================================================================================
Private Sub RaiseError(ByVal ErrorCode As W32F_Errors, _
                       Optional sExtra)
Dim Win32Err As Long, Win32Text As String
  Win32Err = Err.LastDllError
  If Win32Err Then
    Win32Text = vbCrLf & "Error " & Win32Err & vbCrLf & _
                DecodeAPIErrors(Win32Err)
  End If
  Select Case ErrorCode
    Case W32F_FILE_ALREADY_OPEN
      Err.Raise W32F_FILE_ALREADY_OPEN, W32F_SOURCE, _
          "The file '" & sExtra & "' is already open." & Win32Text
    Case W32F_PROBLEM_OPENING_FILE
      Err.Raise W32F_PROBLEM_OPENING_FILE, W32F_SOURCE, _
          "Error opening '" & sExtra & "'." & Win32Text
    Case W32F_FILE_ALREADY_CLOSED
      Err.Raise W32F_FILE_ALREADY_CLOSED, W32F_SOURCE, _
          "There is no open file."
    Case W32F_Problem_seeking
      Err.Raise W32F_Problem_seeking, W32F_SOURCE, _
          "Seek Error." & vbCrLf & sExtra
    Case Else
      Err.Raise W32F_UNKNOWN_ERROR, W32F_SOURCE, _
         "Unknown error." & Win32Text
  End Select
End Sub

'==================================================================================================
' Decode erros from the API
'==================================================================================================
Private Function DecodeAPIErrors(ByVal ErrorCode As Long) As String
Dim sMessage As String, MessageLength As Long
  sMessage = Space$(256)
  MessageLength = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, 0&, _
                                ErrorCode, 0&, sMessage, 256&, 0&)
  If MessageLength > 0 Then
    DecodeAPIErrors = Left(sMessage, MessageLength)
  Else
    DecodeAPIErrors = "Unknown Error."
  End If
End Function
        
'==================================================================================================
' Convert long to 4-byte array
' From http://www.vbaexpress.com/kb/getarticle.php?kb_id=1052
'==================================================================================================
Private Function LongToByteArray(ByVal DataVal As Long) As Byte()
    
    Dim TempLongData As TypedLong                       '<- Long value used in conversion
    Dim TempByteData As ByteLong                        '<- 4-byte array used in conversion
    
    TempLongData.Value = DataVal
    LSet TempByteData = TempLongData
    
    LongToByteArray = TempByteData.Value
    
End Function
 
'==================================================================================================
' Convert 4-byte array to long
' Array passed ByRef - use temp variable to prevent modification to original data
' From http://www.vbaexpress.com/kb/getarticle.php?kb_id=1052
'==================================================================================================
Private Function ByteArrayToLong(ByRef ByteData() As Byte) As Long
    
    Const MaxBound As Long = 3                          '<- Max byte array bound
    Const MinBound As Long = 0                          '<- Min byte array bound
    
    Dim SizedByteData() As Byte                         '<- Properly bound input byte array
    Dim TempLongData As TypedLong                       '<- Long value used in conversion
    Dim TempByteData As ByteLong                        '<- 4-byte array used in conversion
    Dim ByteIdx As Long                                 '<- Byte array loop index
    
    SizedByteData = ByteData
    
    '___ Prevent subscript errors _________________________________________________________________
    If UBound(SizedByteData) < MaxBound Then
        ReDim Preserve SizedByteData(MaxBound)
    End If
    
    '___ Load array into ByteLong - LSet syntax can only be used on User Defined Types ________
    For ByteIdx = MinBound To MaxBound
        TempByteData.Value(ByteIdx) = SizedByteData(ByteIdx)
    Next
    
    LSet TempLongData = TempByteData
    
    ByteArrayToLong = TempLongData.Value
    
End Function

'==================================================================================================
' Convert long to 4-byte array
' From http://www.vbaexpress.com/kb/getarticle.php?kb_id=1052
'==================================================================================================
Private Function IntegerToByteArray(ByVal DataVal As Integer) As Byte()
    
    Dim TempIntData As TypedInteger                     '<- Integer value used in conversion
    Dim TempByteData As ByteInteger                     '<- 2-byte array used in conversion
    
    TempIntData.Value = DataVal
    LSet TempByteData = TempIntData
    
    IntegerToByteArray = TempByteData.Value
    
End Function
 
'==================================================================================================
' Convert 4-byte array to long
' Array passed ByRef - use temp variable to prevent modification to original data
' From http://www.vbaexpress.com/kb/getarticle.php?kb_id=1052
'==================================================================================================
Private Function ByteArrayToInteger(ByRef ByteData() As Byte) As Integer
    
    Const MaxBound As Long = 1                          '<- Max byte array bound
    Const MinBound As Long = 0                          '<- Min byte array bound
    
    Dim SizedByteData() As Byte                         '<- Properly bound input byte array
    Dim TempIntData As TypedInteger                     '<- Long value used in conversion
    Dim TempByteData As ByteInteger                     '<- 2-byte array used in conversion
    Dim ByteIdx As Long                                 '<- Byte array loop index
    
    SizedByteData = ByteData
    
    '___ Prevent subscript errors _________________________________________________________________
    If UBound(SizedByteData) < MaxBound Then
        ReDim Preserve SizedByteData(MaxBound)
    End If
    
    '___ Load array into ByteLong - LSet syntax can only be used on User Defined Types ________
    For ByteIdx = MinBound To MaxBound
        TempByteData.Value(ByteIdx) = SizedByteData(ByteIdx)
    Next
    
    LSet TempIntData = TempByteData
    
    ByteArrayToInteger = TempIntData.Value
    
End Function



