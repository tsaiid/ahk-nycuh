#Requires AutoHotkey v2.0

class AHKHID {
    ;______________________________________
    ;Constants
    static DI_DEVTYPE := 4
    static DI_MSE_ID := 8
    static DI_MSE_NUMBEROFBUTTONS := 12
    static DI_MSE_SAMPLERATE := 16
    static DI_MSE_HASHORIZONTALWHEEL := 20
    static DI_KBD_TYPE := 8
    static DI_KBD_SUBTYPE := 12
    static DI_KBD_KEYBOARDMODE := 16
    static DI_KBD_NUMBEROFFUNCTIONKEYS := 20
    static DI_KBD_NUMBEROFINDICATORS := 24
    static DI_KBD_NUMBEROFKEYSTOTAL := 28
    static DI_HID_VENDORID := 8
    static DI_HID_PRODUCTID := 12
    static DI_HID_VERSIONNUMBER := 16
    static DI_HID_USAGEPAGE := 20 | 0x0100
    static DI_HID_USAGE := 22 | 0x0100

    static II_DEVTYPE := 0
    static II_DEVHANDLE := 8
    static II_MSE_FLAGS := (08 + A_PtrSize * 2) | 0x0100
    static II_MSE_BUTTONFLAGS := (12 + A_PtrSize * 2) | 0x0100
    static II_MSE_BUTTONDATA := (14 + A_PtrSize * 2) | 0x1100
    static II_MSE_RAWBUTTONS := (16 + A_PtrSize * 2)
    static II_MSE_LASTX := (20 + A_PtrSize * 2) | 0x1000
    static II_MSE_LASTY := (24 + A_PtrSize * 2) | 0x1000
    static II_MSE_EXTRAINFO := (28 + A_PtrSize * 2)
    static II_KBD_MAKECODE := (08 + A_PtrSize * 2) | 0x0100
    static II_KBD_FLAGS := (10 + A_PtrSize * 2) | 0x0100
    static II_KBD_VKEY := (14 + A_PtrSize * 2) | 0x0100
    static II_KBD_MSG := (16 + A_PtrSize * 2)
    static II_KBD_EXTRAINFO := (20 + A_PtrSize * 2)
    static II_HID_SIZE := (08 + A_PtrSize * 2)
    static II_HID_COUNT := (12 + A_PtrSize * 2)
    static II_HID_DATA := (16 + A_PtrSize * 2)

    static RIM_TYPEMOUSE := 0
    static RIM_TYPEKEYBOARD := 1
    static RIM_TYPEHID := 2

    static RIDEV_REMOVE := 0x00000001
    static RIDEV_EXCLUDE := 0x00000010
    static RIDEV_PAGEONLY := 0x00000020
    static RIDEV_NOLEGACY := 0x00000030
    static RIDEV_INPUTSINK := 0x00000100
    static RIDEV_CAPTUREMOUSE := 0x00000200
    static RIDEV_NOHOTKEYS := 0x00000200
    static RIDEV_APPKEYS := 0x00000400
    static RIDEV_EXINPUTSINK := 0x00001000
    static RIDEV_DEVNOTIFY := 0x00002000

    static RIM_INPUT := 0
    static RIM_INPUTSINK := 1

    static RID_INPUT := 0x10000003
    static RID_HEADER := 0x10000005

    ; Internal state
    ; 修改處 1: 初始為 0，而不是 unset，這樣可以直接用於 if 判斷
    static uHIDList := 0

    static Initialize(bRefresh := False) {
        ; 修改處 2: 移除 IsSet，改用屬性本身判斷 (如果已初始化為 Buffer，則為真；如果是 0，則為假)
        if AHKHID.uHIDList && !bRefresh
            return AHKHID.uHIDList

        iCount := 0
        r := DllCall("GetRawInputDeviceList", "Ptr", 0, "UInt*", &iCount, "UInt", A_PtrSize * 2)

        if (r == -1) {
            return -1
        }

        AHKHID.uHIDList := Buffer(iCount * (A_PtrSize * 2))
        r := DllCall("GetRawInputDeviceList", "Ptr", AHKHID.uHIDList, "UInt*", &iCount, "UInt", A_PtrSize * 2)

        if (r == -1) {
            return -1
        }
        return AHKHID.uHIDList
    }

    static GetDevCount() {
        iCount := 0
        r := DllCall("GetRawInputDeviceList", "Ptr", 0, "UInt*", &iCount, "UInt", A_PtrSize * 2)
        if (r == -1)
            return -1
        return iCount
    }

    static GetDevHandle(i) {
        ; i is 1-based index to match v1 logic
        return NumGet(AHKHID.Initialize(), (i - 1) * (A_PtrSize * 2), "Ptr")
    }

    static GetDevIndex(Handle) {
        loop AHKHID.GetDevCount()
            if (NumGet(AHKHID.Initialize(), (A_Index - 1) * (A_PtrSize * 2), "Ptr") = Handle)
                return A_Index
        return 0
    }

    static GetDevType(i, IsHandle := False) {
        offset := IsHandle ? ((AHKHID.GetDevIndex(i) - 1) * (A_PtrSize * 2)) : ((i - 1) * (A_PtrSize * 2))
        return NumGet(AHKHID.Initialize(), offset + A_PtrSize, "UInt")
    }

    static GetDevName(i, IsHandle := False) {
        h := IsHandle ? i : AHKHID.GetDevHandle(i)
        iLength := 0

        r := DllCall("GetRawInputDeviceInfo", "Ptr", h, "UInt", 0x20000007, "Ptr", 0, "UInt*", &iLength)
        if (r == -1)
            return ""

        s := Buffer((iLength + 1) * 2)
        r := DllCall("GetRawInputDeviceInfo", "Ptr", h, "UInt", 0x20000007, "Ptr", s, "UInt*", &iLength)
        if (r == -1)
            return ""

        return StrGet(s)
    }

    static GetDevInfo(i, Flag, IsHandle := False) {
        static uInfo := Buffer(0)
        static iLastHandle := 0

        h := IsHandle ? i : AHKHID.GetDevHandle(i)

        if (h = iLastHandle && uInfo.Size > 0) {
            ; Handle hasn't changed, return cached data
        } else {
            iLength := 0
            r := DllCall("GetRawInputDeviceInfo", "Ptr", h, "UInt", 0x2000000b, "Ptr", 0, "UInt*", &iLength)
            if (r == -1)
                return -1

            uInfo := Buffer(iLength)
            NumPut("UInt", iLength, uInfo, 0) ; Size of RIDI_DEVICEINFO

            r := DllCall("GetRawInputDeviceInfo", "Ptr", h, "UInt", 0x2000000b, "Ptr", uInfo, "UInt*", &iLength)
            if (r == -1)
                return -1

            iLastHandle := h
        }

        ; Retrieve data based on Flag type
        isShort := AHKHID.NumIsShort(&Flag)
        return NumGet(uInfo, Flag, isShort ? "UShort" : "UInt")
    }

    static Register(UsagePage := False, Usage := False, Handle := False, Flags := 0) {
        uDev := Buffer(8 + A_PtrSize)

        ; Check if hwnd needs to be null. RIDEV_REMOVE, RIDEV_EXCLUDE
        Handle := ((Flags & 0x00000001) || (Flags & 0x00000010)) ? 0 : Handle

        NumPut("UShort", UsagePage, uDev, 0)
        NumPut("UShort", Usage, uDev, 2)
        NumPut("UInt", Flags, uDev, 4)
        NumPut("Ptr", Handle, uDev, 8)

        r := DllCall("RegisterRawInputDevices", "Ptr", uDev, "UInt", 1, "UInt", 8 + A_PtrSize)

        if !r { ; RegisterRawInputDevices returns FALSE on failure
            return -1
        }
        return 0
    }

    static GetInputInfo(InputHandle, Flag) {
        static uRawInput := Buffer(0)
        static iLastHandle := 0

        if (InputHandle = iLastHandle && uRawInput.Size > 0) {
            ; Use cached
        } else {
            iSize := 0
            r := DllCall("GetRawInputData", "Ptr", InputHandle, "UInt", 0x10000003, "Ptr", 0, "UInt*", &iSize, "UInt",
                8 + A_PtrSize * 2)
            if (r == -1)
                return -1

            uRawInput := Buffer(iSize)
            r := DllCall("GetRawInputData", "Ptr", InputHandle, "UInt", 0x10000003, "Ptr", uRawInput, "UInt*", &iSize,
                "UInt", 8 + A_PtrSize * 2)
            if (r == -1)
                return -1

            iLastHandle := InputHandle
        }

        isShort := AHKHID.NumIsShort(&Flag)
        isSigned := AHKHID.NumIsSigned(&Flag)

        typeStr := isShort ? (isSigned ? "Short" : "UShort") : (isSigned ? "Int" : (Flag = 8 ? "Ptr" : "UInt"))

        return NumGet(uRawInput, Flag, typeStr)
    }

    static GetInputData(InputHandle) {
        iSize := 0
        r := DllCall("GetRawInputData", "Ptr", InputHandle, "UInt", 0x10000003, "Ptr", 0, "UInt*", &iSize, "UInt", 8 +
            A_PtrSize * 2)
        if (r == -1)
            return Buffer(0)

        uRawInput := Buffer(iSize)
        r := DllCall("GetRawInputData", "Ptr", InputHandle, "UInt", 0x10000003, "Ptr", uRawInput, "UInt*", &iSize,
            "UInt", 8 + A_PtrSize * 2)
        if (r == -1)
            return Buffer(0)

        ; Get size of each HID input and count
        iSizeBody := NumGet(uRawInput, 8 + A_PtrSize * 2 + 0, "UInt") ; ID_HID_SIZE
        iCountBody := NumGet(uRawInput, 8 + A_PtrSize * 2 + 4, "UInt") ; ID_HID_COUNT

        if (iSizeBody * iCountBody <= 0)
            return Buffer(0)

        uData := Buffer(iSizeBody * iCountBody)

        ; Calculate source address: uRawInput.Ptr + HeaderSize + HID specific header part
        srcPtr := uRawInput.Ptr + 8 + A_PtrSize * 2 + 8

        DllCall("RtlMoveMemory", "Ptr", uData.Ptr, "Ptr", srcPtr, "UInt", iSizeBody * iCountBody)

        return uData
    }

    static NumIsShort(&Flag) {
        if (Flag & 0x0100) {
            Flag ^= 0x0100
            return True
        }
        return False
    }

    static NumIsSigned(&Flag) {
        if (Flag & 0x1000) {
            Flag ^= 0x1000
            return True
        }
        return False
    }
}
