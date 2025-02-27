## Extract Zip Files through Windows Shell

Just a code snippet that shows a simple way to unzip...

The simplified method is shorter, but the original would make it easier for you to only extract certain files.

This is the complete code, just requires enabling my "Windows Development Library for twinBASIC" (WinDevLib) package in References->Available packages.

```vba
    Private Function ExtractByShell(pszZip As String, pszDest As String) As Long
        If PathFileExistsW(StrPtr(pszZip)) = 0 Then
            ExtractByShell = ERROR_FILE_NOT_FOUND
            Exit Function
        End If
        If PathFileExistsW(StrPtr(pszZip)) = 0 Then
            SHCreateDirectory 0, pszDest
        End If
        
        Dim siZip As IShellItem
        Dim siDest As IShellItem
        Dim siChild As IShellItem
        Dim pEnum As IEnumShellItems
        Dim pArray As IShellItemArray
        Dim pCopy As New FileOperation
        Dim pidl() As LongPtr
        Dim cPidl As Long
        Dim pPIDL As IPersistIDList
        Dim lRet As Long
        lRet = SHCreateItemFromParsingName(StrPtr(pszZip), Nothing, IID_IShellItem, siZip)
        lRet = SHCreateItemFromParsingName(StrPtr(pszDest), Nothing, IID_IShellItem, siDest)
        If (siZip Is Nothing) Or (siDest Is Nothing) Then
            ExtractByShell = ERROR_FILE_NOT_FOUND
            Exit Function
        End If
        lRet = siZip.BindToHandler(0, BHID_EnumItems, IID_IEnumShellItems, pEnum)
        If (pEnum Is Nothing) = False Then
            Do While pEnum.Next(1, siChild) = S_OK
                Set pPIDL = siChild
                ReDim Preserve pidl(cPidl)
                pPIDL.GetIDList pidl(cPidl)
                cPidl = cPidl + 1
            Loop
            If cPidl Then
                SHCreateShellItemArrayFromIDLists cPidl, VarPtr(pidl(0)), pArray
                If (pArray Is Nothing) = False Then
                    pCopy.CopyItems pArray, siDest
                    pCopy.PerformOperations
                    pCopy.GetAnyOperationsAborted lRet
                End If
                FreeIDListArray pidl, cPidl
            End If
        End If
        ExtractByShell = lRet
    End Function
    Private Function ExtractByShellSimplfied(pszZip As String, pszDest As String) As Long
        If PathFileExistsW(StrPtr(pszZip)) = 0 Then
            ExtractByShellSimplfied = ERROR_FILE_NOT_FOUND
            Exit Function
        End If
        If PathFileExistsW(StrPtr(pszZip)) = 0 Then
            SHCreateDirectory 0, pszDest
        End If
        
        Dim siZip As IShellItem
        Dim siDest As IShellItem
        Dim siChild As IShellItem
        Dim pEnum As IEnumShellItems
        Dim pCopy As New FileOperation
        Dim lRet As Long
        lRet = SHCreateItemFromParsingName(StrPtr(pszZip), Nothing, IID_IShellItem, siZip)
        lRet = SHCreateItemFromParsingName(StrPtr(pszDest), Nothing, IID_IShellItem, siDest)
        If (siZip Is Nothing) Or (siDest Is Nothing) Then
            ExtractByShellSimplfied = ERROR_FILE_NOT_FOUND
            Exit Function
        End If
        lRet = siZip.BindToHandler(0, BHID_EnumItems, IID_IEnumShellItems, pEnum)
        If (pEnum Is Nothing) = False Then
            pCopy.CopyItems pEnum, siDest
            pCopy.PerformOperations
            pCopy.GetAnyOperationsAborted lRet
        End If
        ExtractByShellSimplfied = lRet
    End Function
```
