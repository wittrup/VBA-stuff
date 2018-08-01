Attribute VB_Name = "mHash"
Option Explicit

Sub Test()
    'run this to test md5, sha1, sha2/256, sha2/384, or sha2/512
    Dim sIn As String, sOut As String, bB64 As Boolean, sH As String
    
    'insert the text to hash within the sIn quotes
    'note that a private string could be joined to sIn at this point
    sIn = ""
    
    'select as required
    'bB64 = False   'output hex
    bB64 = True   'output base-64
    
    'enable any one
    'sH = MD5(sIn, bB64)
    'sH = SHA1(sIn, bB64)
    'sH = SHA256(sIn, bB64)
    'sH = SHA384(sIn, bB64)
    sH = SHA512(sIn, bB64)
    
    Debug.Print sH & vbNewLine & Len(sH) & " characters in length"
    MsgBox sH & vbNewLine & Len(sH) & " characters in length"
   
End Sub

Public Function MD5(ByVal sIn As String, Optional bB64 As Boolean = 0) As String
    'Set a reference to mscorlib 4.0 64-bit
    
    'Test with empty string input:
    'Hex:   d41d8cd98f00...etc
    'Base-64: 1B2M2Y8Asg...etc
        
    Dim oT As Object, oMD5 As Object
    Dim TextToHash() As Byte
    Dim bytes() As Byte
        
    Set oT = CreateObject("System.Text.UTF8Encoding")
    Set oMD5 = CreateObject("System.Security.Cryptography.MD5CryptoServiceProvider")
 
    TextToHash = oT.GetBytes_4(sIn)
    bytes = oMD5.ComputeHash_2((TextToHash))
 
    If bB64 = True Then
       MD5 = ConvToBase64String(bytes)
    Else
       MD5 = ConvToHexString(bytes)
    End If
        
    Set oT = Nothing
    Set oMD5 = Nothing

End Function

Public Function SHA1(sIn As String, Optional bB64 As Boolean = 0) As String
    'Set a reference to mscorlib 4.0 64-bit
    
    'Test with empty string input:
    '40 Hex:   da39a3ee5e6...etc
    '28 Base-64:   2jmj7l5rSw0yVb...etc
    
    Dim oT As Object, oSHA1 As Object
    Dim TextToHash() As Byte
    Dim bytes() As Byte
            
    Set oT = CreateObject("System.Text.UTF8Encoding")
    Set oSHA1 = CreateObject("System.Security.Cryptography.SHA1Managed")
    
    TextToHash = oT.GetBytes_4(sIn)
    bytes = oSHA1.ComputeHash_2((TextToHash))
        
    If bB64 = True Then
       SHA1 = ConvToBase64String(bytes)
    Else
       SHA1 = ConvToHexString(bytes)
    End If
            
    Set oT = Nothing
    Set oSHA1 = Nothing
    
End Function

Public Function SHA256(sIn As String, Optional bB64 As Boolean = 0) As String
    'Set a reference to mscorlib 4.0 64-bit
    
    'Test with empty string input:
    '64 Hex:   e3b0c44298f...etc
    '44 Base-64:   47DEQpj8HBSa+/...etc
    
    Dim oT As Object, oSHA256 As Object
    Dim TextToHash() As Byte, bytes() As Byte
    
    Set oT = CreateObject("System.Text.UTF8Encoding")
    Set oSHA256 = CreateObject("System.Security.Cryptography.SHA256Managed")
    
    TextToHash = oT.GetBytes_4(sIn)
    bytes = oSHA256.ComputeHash_2((TextToHash))
    
    If bB64 = True Then
       SHA256 = ConvToBase64String(bytes)
    Else
       SHA256 = ConvToHexString(bytes)
    End If
    
    Set oT = Nothing
    Set oSHA256 = Nothing
    
End Function

Public Function SHA384(sIn As String, Optional bB64 As Boolean = 0) As String
    'Set a reference to mscorlib 4.0 64-bit
    
    'Test with empty string input:
    '96 Hex:   38b060a751ac...etc
    '64 Base-64:   OLBgp1GsljhM2T...etc
    
    Dim oT As Object, oSHA384 As Object
    Dim TextToHash() As Byte, bytes() As Byte
    
    Set oT = CreateObject("System.Text.UTF8Encoding")
    Set oSHA384 = CreateObject("System.Security.Cryptography.SHA384Managed")
    
    TextToHash = oT.GetBytes_4(sIn)
    bytes = oSHA384.ComputeHash_2((TextToHash))
    
    If bB64 = True Then
       SHA384 = ConvToBase64String(bytes)
    Else
       SHA384 = ConvToHexString(bytes)
    End If
    
    Set oT = Nothing
    Set oSHA384 = Nothing
    
End Function

Public Function SHA512(sIn As String, Optional bB64 As Boolean = 0) As String
    'Set a reference to mscorlib 4.0 64-bit
    
    'Test with empty string input:
    '128 Hex:   cf83e1357eefb8bd...etc
    '88 Base-64:   z4PhNX7vuL3xVChQ...etc
    
    Dim oT As Object, oSHA512 As Object
    Dim TextToHash() As Byte, bytes() As Byte
    
    Set oT = CreateObject("System.Text.UTF8Encoding")
    Set oSHA512 = CreateObject("System.Security.Cryptography.SHA512Managed")
    
    TextToHash = oT.GetBytes_4(sIn)
    bytes = oSHA512.ComputeHash_2((TextToHash))
    
    If bB64 = True Then
       SHA512 = ConvToBase64String(bytes)
    Else
       SHA512 = ConvToHexString(bytes)
    End If
    
    Set oT = Nothing
    Set oSHA512 = Nothing
    
End Function

Private Function ConvToBase64String(vIn As Variant) As Variant

    Dim oD As Object
      
    Set oD = CreateObject("MSXML2.DOMDocument")
      With oD
        .LoadXML "<root />"
        .DocumentElement.DataType = "bin.base64"
        .DocumentElement.nodeTypedValue = vIn
      End With
    ConvToBase64String = Replace(oD.DocumentElement.text, vbLf, "")
    
    Set oD = Nothing

End Function

Private Function ConvToHexString(vIn As Variant) As Variant

    Dim oD As Object
      
    Set oD = CreateObject("MSXML2.DOMDocument")
      
      With oD
        .LoadXML "<root />"
        .DocumentElement.DataType = "bin.Hex"
        .DocumentElement.nodeTypedValue = vIn
      End With
    ConvToHexString = Replace(oD.DocumentElement.text, vbLf, "")
    
    Set oD = Nothing

End Function

Private Sub TestFileHashes()
    'run this to test the file hasher. Select or comment lines as necessary
    'enter your own paths for the files to test
    'Set a reference to mscorlib 4.0 64-bit
    
    Dim sFPath As String, b64 As Boolean, sH As String
    
    'set output type and path to target file
    'b64 = False   'output hex
    b64 = True     'output base-64
    sFPath = "C:\Users\Your Folder\Documents\test.txt"
    
    'enable any one line to test hash
    'sh=FileToMD5(sFPath, b64)
    'sh=FileToSHA1(sFPath, b64)
    'sh=FileToSHA256(sFPath, b64)
    'sH = FileToSHA384(sFPath, b64)
    sH = FileToSHA512(sFPath, b64)
    
    Debug.Print sH & vbNewLine & Len(sH) & " characters in length"
    MsgBox sH & vbNewLine & Len(sH) & " characters in length"

End Sub

Public Function FileToMD5(sFullPath As String, Optional bB64 As Boolean = False) As String
    'parameter full path with name of file returned in the function as an MD5 hash
    'Set a reference to mscorlib 4.0 64-bit
    
    Dim enc, bytes, outstr As String, pos As Integer
    
    Set enc = CreateObject("System.Security.Cryptography.MD5CryptoServiceProvider")
    'Convert the string to a byte array and hash it
    bytes = GetFileBytes(sFullPath)
    bytes = enc.ComputeHash_2((bytes))
    
    If bB64 = True Then
       FileToMD5 = ConvToBase64String(bytes)
    Else
       FileToMD5 = ConvToHexString(bytes)
    End If
        
    Set enc = Nothing

End Function

Public Function FileToSHA1(sFullPath As String, Optional bB64 As Boolean = False) As String
    'parameter full path with name of file returned in the function as an SHA1 hash
    'Set a reference to mscorlib 4.0 64-bit
    
    Dim enc, bytes, outstr As String, pos As Integer
    
    Set enc = CreateObject("System.Security.Cryptography.SHA1CryptoServiceProvider")
    'Convert the string to a byte array and hash it
    bytes = GetFileBytes(sFullPath) 'returned as a byte array
    bytes = enc.ComputeHash_2((bytes))
    
    If bB64 = True Then
       FileToSHA1 = ConvToBase64String(bytes)
    Else
       FileToSHA1 = ConvToHexString(bytes)
    End If
        
    Set enc = Nothing

End Function

Public Function FileToSHA256(sFullPath As String, Optional bB64 As Boolean = False) As String
    'parameter full path with name of file returned in the function as an SHA2-256 hash
    'Set a reference to mscorlib 4.0 64-bit
    
    Dim enc, bytes, outstr As String, pos As Integer
    
    Set enc = CreateObject("System.Security.Cryptography.SHA256Managed")
    'Convert the string to a byte array and hash it
    bytes = GetFileBytes(sFullPath) 'returned as a byte array
    bytes = enc.ComputeHash_2((bytes))
    
    If bB64 = True Then
       FileToSHA256 = ConvToBase64String(bytes)
    Else
       FileToSHA256 = ConvToHexString(bytes)
    End If
        
    Set enc = Nothing

End Function

Public Function FileToSHA384(sFullPath As String, Optional bB64 As Boolean = False) As String
    'parameter full path with name of file returned in the function as an SHA2-384 hash
    'Set a reference to mscorlib 4.0 64-bit
    
    Dim enc, bytes, outstr As String, pos As Integer
    
    Set enc = CreateObject("System.Security.Cryptography.SHA384Managed")
    'Convert the string to a byte array and hash it
    bytes = GetFileBytes(sFullPath) 'returned as a byte array
    bytes = enc.ComputeHash_2((bytes))
    
    If bB64 = True Then
       FileToSHA384 = ConvToBase64String(bytes)
    Else
       FileToSHA384 = ConvToHexString(bytes)
    End If
    
    Set enc = Nothing

End Function

Public Function FileToSHA512(sFullPath As String, Optional bB64 As Boolean = False) As String
    'parameter full path with name of file returned in the function as an SHA2-512 hash
    'Set a reference to mscorlib 4.0 64-bit
    
    Dim enc, bytes, outstr As String, pos As Integer
    
    Set enc = CreateObject("System.Security.Cryptography.SHA512Managed")
    'Convert the string to a byte array and hash it
    bytes = GetFileBytes(sFullPath) 'returned as a byte array
    bytes = enc.ComputeHash_2((bytes))
    
    If bB64 = True Then
       FileToSHA512 = ConvToBase64String(bytes)
    Else
       FileToSHA512 = ConvToHexString(bytes)
    End If
    
    Set enc = Nothing

End Function

Private Function GetFileBytes(ByVal sPath As String) As Byte()
    'makes byte array from file
    'Set a reference to mscorlib 4.0 64-bit
    
    Dim lngFileNum As Long, bytRtnVal() As Byte, bTest
    
    lngFileNum = FreeFile
    
    If LenB(Dir(sPath)) Then ''// Does file exist?
        
        Open sPath For Binary Access Read As lngFileNum
        
        'a zero length file content will give error 9 here
        
        ReDim bytRtnVal(0 To LOF(lngFileNum) - 1&) As Byte
        Get lngFileNum, , bytRtnVal
        Close lngFileNum
    Else
        Err.Raise 53 'File not found
    End If
    
    GetFileBytes = bytRtnVal
    
    Erase bytRtnVal

End Function


Function GetFileSize(sFilePath As String, nSize As Long) As Boolean
    'use this to test for a zero file size
    'takes full path as string in sFileSize
    'returns file size in bytes in nSize
    
    Dim fs As FileSystemObject, f As File
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    If fs.FileExists(sFilePath) Then
        Set f = fs.GetFile(sFilePath)
        nSize = f.Size
        GetFileSize = True
        Exit Function
    End If

End Function
