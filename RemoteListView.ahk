;///////////////////////////////////////////////////////////////////////////////////////////
; RemoteListView.ahk v1.0.0-git
; Copyright (c) 2025 The-CoDingman
; https://github.com/The-CoDingman/RemoteListView.ahk
;
; MIT License
;
; Permission is hereby granted, free of charge, to any person obtaining a copy
; of this software and associated documentation files (the "Software"), to deal
; in the Software without restriction, including without limitation the rights
; to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
; copies of the Software, and to permit persons to whom the Software is
; furnished to do so, subject to the following conditions:
;
; The above copyright notice and this permission notice shall be included in all
; copies or substantial portions of the Software.
;
; THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
; IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
; FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
; AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
; LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
; OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
; SOFTWARE.
;///////////////////////////////////////////////////////////////////////////////////////////

#Requires AutoHotkey v2
class RemoteListView {
    ;Constants for ListView controls
    static WC_LISTVIEW := "SysListView32"

    ;Messages
    static WM_SETFOCUS             := 0x0007
    static LVM_GETITEMSTATE        := 0x102C
    static LVM_SETITEMSTATE        := 0x102B
    static LVM_ENSUREVISIBLE       := 0x1013
    static LVM_GETITEMCOUNT        := 0x1004
    static LVM_GETITEMTEXT         := 0x102D
    static LVM_FINDITEM            := 0x1053

    ;Item states
    static LVIS_FOCUSED := 0x0001
    static LVIS_SELECTED := 0x0002

    ;Item Masks
    static LVIF_TEXT := 0x0001
    static LVFI_STRING := 0x0002
    static LVFI_PARTIAL := 0x0008
    static LVFI_PARAM := 0x0001

    ;Memory constants
    static PAGE_READWRITE := 0x04
    static MEM_COMMIT := 0x1000
    
    static MEM_RELEASE := 0x8000
    static PROCESS_VM_OPERATION := 0x0008
    static PROCESS_VM_READ := 0x0010
    static PROCESS_VM_WRITE := 0x0020

    __New(LVHwnd) {
        this.LVHWnd := LVHWnd
    }
    
    FindItem(StartIndex := -1, SearchFlags := 0, SearchText := "", SearchColumn := 0, SearchLParam := 0) {
        ProcessId := WinGetPID("ahk_id " this.LVHwnd)
        hProcess := RemoteListView.OpenProcess(RemoteListView.PROCESS_VM_OPERATION | RemoteListView.PROCESS_VM_WRITE, false, ProcessId)
        if (!hProcess) {
            return -1
        }
        
        ;Allocate memory for LVFINDINFO structure
        LVFINDINFO_SIZE := A_PtrSize * 4
        pLVFINDINFO := RemoteListView.VirtualAllocEx(hProcess, 0, LVFINDINFO_SIZE, RemoteListView.MEM_COMMIT, RemoteListView.PAGE_READWRITE)
        if (!pLVFINDINFO) {
            RemoteListView.CloseHandle(hProcess)
            return -1
        }
        
        pText := 0
        if (SearchFlags & RemoteListView.LVFI_STRING) && (SearchText != "") {
            ;Allocate memory for text if doing text search
            textBuffer := Buffer(StrLen(SearchText) * 2 + 2, 0)
            StrPut(SearchText, textBuffer, "UTF-16")
            pText := RemoteListView.VirtualAllocEx(hProcess, 0, textBuffer.Size, RemoteListView.MEM_COMMIT, RemoteListView.PAGE_READWRITE)
            if (!pText) {
                RemoteListView.VirtualFreeEx(hProcess, pLVFINDINFO, 0, RemoteListView.MEM_RELEASE)
                RemoteListView.CloseHandle(hProcess)
                return -1
            }
            RemoteListView.WriteProcessMemory(hProcess, pText, textBuffer, textBuffer.Size)
        }
        
        ;Prepare LVFINDINFO structure
        lvFindInfo := Buffer(LVFINDINFO_SIZE, 0)
        NumPut("UInt", SearchFlags, lvFindInfo, 0) ; flags
        if (SearchFlags & RemoteListView.LVFI_STRING) {
            NumPut("Ptr", pText, lvFindInfo, A_PtrSize) ; psz (text to search)
        } else if (SearchFlags & RemoteListView.LVFI_PARAM) {
            NumPut("Ptr", SearchLParam, lvFindInfo, A_PtrSize * 2) ; lParam
        }
        NumPut("Int", SearchColumn, lvFindInfo, A_PtrSize * 3) ; lParam for column search
        
        ;Write to remote process and send message
        RemoteListView.WriteProcessMemory(hProcess, pLVFINDINFO, lvFindInfo, LVFINDINFO_SIZE)
        result := SendMessage(RemoteListView.LVM_FINDITEM, StartIndex, pLVFINDINFO,, this.LVHwnd)
        
        ;Cleanup
        if (pText) {
            RemoteListView.VirtualFreeEx(hProcess, pText, 0, RemoteListView.MEM_RELEASE)
        }
        RemoteListView.VirtualFreeEx(hProcess, pLVFINDINFO, 0, RemoteListView.MEM_RELEASE)
        RemoteListView.CloseHandle(hProcess)
        return result
    }

    FindItemByText(Text, StartIndex := -1, Column := 0, Occurrence := 1, PartialMatch := false) {
        Flags := RemoteListView.LVFI_STRING
        if (PartialMatch) {
            Flags |= RemoteListView.LVFI_PARTIAL
        }
        
        CurrentIndex := StartIndex
        CurrentOccurrence := 0
        Loop {
            FoundIndex := this.FindItem(CurrentIndex, Flags, Text, Column)
            if (FoundIndex = -1) {
                break
            }

            CurrentOccurrence++
            if (CurrentOccurrence = Occurrence) {
                return FoundIndex
            }
            CurrentIndex := FoundIndex
        }
        return -1
    }
    
    FindItemByLParam(lParamValue, StartIndex := -1, Occurrence := 1) {
        CurrentIndex := StartIndex
        CurrentOccurrence := 0
        Loop {
            FoundIndex := this.FindItem(CurrentIndex, RemoteListView.LVFI_PARAM, "", 0, lParamValue)
            if (FoundIndex = -1) {
                break
            }
            
            CurrentOccurrence++
            if (CurrentOccurrence = Occurrence) {
                return FoundIndex
            }            
            CurrentIndex := FoundIndex ; Start from next item
        }        
        return -1
    }
    
    GetItemText(ItemIndex, Column := 0) {
        ProcessId := WinGetPID("ahk_id " this.LVHwnd)
        hProcess := RemoteListView.OpenProcess(RemoteListView.PROCESS_VM_OPERATION | RemoteListView.PROCESS_VM_WRITE | RemoteListView.PROCESS_VM_READ, false, ProcessId)
        if (!hProcess) {
            return ""
        }

        ;Create buffer for text
        TEXT_BUFFER_SIZE := 1024
        pTextBuffer := RemoteListView.VirtualAllocEx(hProcess, 0, TEXT_BUFFER_SIZE, RemoteListView.MEM_COMMIT, RemoteListView.PAGE_READWRITE)
        if (!pTextBuffer) {
            RemoteListView.CloseHandle(hProcess)
            return ""
        }

        ;Empty buffer
        EmptyBuffer := Buffer(TEXT_BUFFER_SIZE, 0)
        RemoteListView.WriteProcessMemory(hProcess, pTextBuffer, EmptyBuffer, TEXT_BUFFER_SIZE)

        ;Build LVITEM structure in remote process memory
        LVITEM_SIZE := 56
        pLVITEM := RemoteListView.VirtualAllocEx(hProcess, 0, LVITEM_SIZE, RemoteListView.MEM_COMMIT, RemoteListView.PAGE_READWRITE)
        if (!pLVITEM) {
            RemoteListView.VirtualFreeEx(hProcess, pTextBuffer, 0, RemoteListView.MEM_RELEASE)
            RemoteListView.CloseHandle(hProcess)
            return ""
        }

        lvItem := Buffer(LVITEM_SIZE, 0)
        NumPut("UInt", RemoteListView.LVIF_TEXT, lvItem, 0) ;mask
        NumPut("Int", ItemIndex, lvItem, 4)                 ;iItem
        NumPut("Int", Column, lvItem, 8)                    ;iSubItem
        NumPut("Ptr", pTextBuffer, lvItem, 24)              ;pszText
        NumPut("Int", TEXT_BUFFER_SIZE, lvItem, 32)         ;cchTextMax
        RemoteListView.WriteProcessMemory(hProcess, pLVITEM, lvItem, LVITEM_SIZE)

        ;Query ListView for text
        SendMessage(RemoteListView.LVM_GETITEMTEXT, ItemIndex, pLVITEM,, this.LVHwnd)
        TextBuffer := Buffer(TEXT_BUFFER_SIZE, 0)
        ItemText := ""
        if (RemoteListView.ReadProcessMemory(hProcess, pTextBuffer, TextBuffer, TEXT_BUFFER_SIZE)) {
            ItemText := StrGet(TextBuffer.Ptr, "UTF-8")
        }

        ; Clean up
        RemoteListView.VirtualFreeEx(hProcess, pLVITEM, 0, RemoteListView.MEM_RELEASE)
        RemoteListView.VirtualFreeEx(hProcess, pTextBuffer, 0, RemoteListView.MEM_RELEASE)
        RemoteListView.CloseHandle(hProcess)
        return ItemText
    }

    SetSelection(ItemIndex, EnsureVisible := true, DefaultAction := true) {
        ;Set focus to the ListView control
        PreviousFocusedControl := DllCall("GetFocus", "Ptr")
        SendMessage(RemoteListView.WM_SETFOCUS, 0, 0, this.LVHwnd)

        ;Make sure the target item is visible
        if (EnsureVisible) {
            SendMessage(RemoteListView.LVM_ENSUREVISIBLE, ItemIndex, 0,, this.LVHwnd)
        }

        ;Build LVITEM structure in remote process memory
        ProcessId := WinGetPID("ahk_id " this.LVHwnd)
        hProcess := RemoteListView.OpenProcess(RemoteListView.PROCESS_VM_OPERATION | RemoteListView.PROCESS_VM_WRITE | RemoteListView.PROCESS_VM_READ, false, ProcessId)
        _lvi := RemoteListView.VirtualAllocEx(hProcess, 0, Size := 32, RemoteListView.MEM_COMMIT, RemoteListView.PAGE_READWRITE)
        lvi := Buffer(Size, 0)

        ;Clear selection
        NumPut("UInt", 0, lvi, 12)
        NumPut("UInt", RemoteListView.LVIS_SELECTED | RemoteListView.LVIS_FOCUSED, lvi, 16)
        RemoteListView.WriteProcessMemory(hProcess, _lvi, lvi, Size)
        SendMessage(RemoteListView.LVM_SETITEMSTATE, -1, _lvi,, this.LVHwnd)

        ;Set selection
        NumPut("UInt", RemoteListView.LVIS_SELECTED | RemoteListView.LVIS_FOCUSED, lvi, 12)
        NumPut("UInt", RemoteListView.LVIS_SELECTED | RemoteListView.LVIS_FOCUSED, lvi, 16)
        RemoteListView.WriteProcessMemory(hProcess, _lvi, lvi, Size)
        SendMessage(RemoteListView.LVM_SETITEMSTATE, ItemIndex, _lvi,, this.LVHwnd)

        ;Verify target has been selected
        IsSelected := SendMessage(RemoteListView.LVM_GETITEMSTATE, ItemIndex, RemoteListView.LVIS_SELECTED,, this.LVHwnd)
        Result := (IsSelected & RemoteListView.LVIS_SELECTED) != 0

        ; Cleanup
        RemoteListView.VirtualFreeEx(hProcess, _lvi, 0, RemoteListView.MEM_RELEASE)
        RemoteListView.CloseHandle(hProcess)
        if (PreviousFocusedControl) {
            DllCall("SetFocus", "ptr", PreviousFocusedControl)
        }
        return Result
    }

    static OpenProcess(dwDesiredAccess, bInheritHandle, dwProcessId) => DllCall("OpenProcess", "UInt", dwDesiredAccess, "Int", bInheritHandle, "UInt", dwProcessId, "Ptr")
    static CloseHandle(hObject) => DllCall("CloseHandle", "Ptr", hObject)
    static VirtualAllocEx(hProcess, lpAddress, dwSize, flAllocationType, flProtect) => DllCall("VirtualAllocEx", "Ptr", hProcess, "Ptr", lpAddress, "UInt", dwSize, "UInt", flAllocationType, "UInt", flProtect, "Ptr")
    static VirtualFreeEx(hProcess, lpAddress, dwSize, dwFreeType) => DllCall("VirtualFreeEx", "Ptr", hProcess, "Ptr", lpAddress, "UInt", dwSize, "UInt", dwFreeType)
    static WriteProcessMemory(hProcess, lpBaseAddress, lpBuffer, nSize) => DllCall("WriteProcessMemory", "Ptr", hProcess, "Ptr", lpBaseAddress, "Ptr", lpBuffer, "UInt", nSize, "UInt*", &lpNumberOfBytesWritten:=0)
    static ReadProcessMemory(hProcess, lpBaseAddress, lpBuffer, nSize) => DllCall("ReadProcessMemory", "Ptr", hProcess, "Ptr", lpBaseAddress, "Ptr", lpBuffer, "UInt", nSize, "UInt*", &lpNumberOfBytesRead:=0)
}
