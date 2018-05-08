

MultiStrReplace(needleArray, Haystack, replaceArray) {
    
    If needleArray.MaxIndex() == replaceArray.MaxIndex() then:
        Loop, % needleArray.MaxIndex() {
            Haystack := StrReplace(Haystack, needleArray[A_Index], replaceArray[A_Index])
        }
        Return Haystack
    }

StrReplace(Haystack, SearchText [, ReplaceText, OutputVarCount, Limit := -1])