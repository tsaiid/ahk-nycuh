#Requires AutoHotkey v2.0

class G3PacsProbe {
    static ClassPrefix := "Afx:00400000:b:00000000:00000013:00000000"
    static Configs := [
        [3, 10, 29, 101],
        [3, 10, 31, 101],
        [3, 10, 33, 101],
        [3, 10, 35, 101],
        [3, 10, 37, 101],
        [3, 10, 39, 101],
        [3, 10, 41, 101],
        [3, 10, 43, 101],
        [3, 10, 45, 101],
        [3, 10, 47, 101],
        [3, 10, 49, 101],
        [3, 10, 51, 101],
        [3, 10, 53, 101],
        [3, 10, 55, 101],
        [5, 57, 29, 185],
        [5, 57, 31, 185],
        [5, 57, 37, 185],
        [5, 57, 39, 185],
        [5, 57, 41, 185],
        [5, 57, 43, 185],
        [5, 57, 45, 185],
        [5, 57, 47, 185],
        [5, 57, 49, 185],
        [5, 57, 51, 185],
        [5, 57, 53, 185],
    ]

    static GetPatternList() {
        patterns := []
        for cfg in this.Configs {
            pMap := Map()
            pName := "Pattern_" . cfg[1] . "_" . cfg[2] . "_" . cfg[3] . "_" . cfg[4]

            Loop 8 {
                offset := A_Index - 1
                focusClassNN := this.ClassPrefix . (cfg[1] + offset)
                imgClassNN := "ComboBox" . (cfg[2] + (offset * 6))
                imgGroupClassNN := "ComboBox" . (cfg[2] + 3 + (offset * 6))
                srsClassNN := "AfxWnd140u" . (cfg[3] + (offset * 3))
                descClassNN := "Button" . (cfg[4] + (offset * 5))
                pMap[focusClassNN] := {
                    img: imgClassNN,
                    imgGroup: imgGroupClassNN,
                    srs: srsClassNN,
                    desc: descClassNN,
                    type: pName
                }
            }
            patterns.Push({name: pName, map: pMap})
        }
        return patterns
    }

    static GetSeriesControlsForFocusClassNN(focusClassNN, hwnd, patterns := "") {
        match := this.GetSeriesMatchForFocusClassNN(focusClassNN, hwnd, patterns)
        return match ? match.candidate : false
    }

    static GetSrsControlForFocusClassNN(focusClassNN, hwnd, patterns := "") {
        controls := this.GetSeriesControlsForFocusClassNN(focusClassNN, hwnd, patterns)
        return controls ? controls.srs : ""
    }

    static GetSeriesMatchForFocusClassNN(focusClassNN, hwnd, patterns := "") {
        if (patterns == "")
            patterns := this.GetPatternList()

        for patternData in patterns {
            pMap := patternData.map
            if !pMap.Has(focusClassNN)
                continue

            candidate := pMap[focusClassNN]
            descText := ""
            if this.IsSeriesPatternMatch(candidate, focusClassNN, hwnd, &descText)
                return {candidate: candidate, name: patternData.name, desc: descText}
        }
        return false
    }

    static IsSeriesPatternMatch(candidate, focusClassNN, hwnd, &descText := "") {
        try {
            srsText := ControlGetText(candidate.srs, hwnd)
            if !InStr(srsText, "VMTool")
                return false

            descText := ControlGetText(candidate.desc, hwnd)
            if !RegExMatch(descText, "^\(\d+\)\s")
                return false
        } catch {
            return false
        }

        return this.IsSpatialControlMatch(candidate.srs, focusClassNN, hwnd)
            && this.IsSpatialControlMatch(candidate.img, focusClassNN, hwnd)
            && this.IsSpatialControlMatch(candidate.desc, focusClassNN, hwnd)
    }

    static IsSpatialControlMatch(childClassNN, focusClassNN, hwnd) {
        try {
            ControlGetPos(&childX, &childY, &childW, &childH, childClassNN, hwnd)
            ControlGetPos(&focusX, &focusY, &focusW, &focusH, focusClassNN, hwnd)
            if (childW <= 0 || childH <= 0 || focusW <= 0 || focusH <= 0)
                return false

            childBottom := childY + childH
            if (Abs(focusY - childBottom) > 50)
                return false

            childCenterX := childX + (childW / 2)
            return childCenterX >= (focusX - 10) && childCenterX <= (focusX + focusW + 10)
        }
        return false
    }
}
