pushHTML(sHtmlFragment)
{
m_sDescription = 
(
Version:1.0
StartHTML:aaaaaaaaaa 
EndHTML:bbbbbbbbbb
StartFragment:cccccccccc 
EndFragment:dddddddddd 
)
sContextStart = <HTML><BODY>
sContextEnd = </BODY></HTML>

   sData = %m_sDescription% %sContextStart% %sHtmlFragment% %sContextEnd%
   mylen := StrLen(m_sDescription) +4
   thelen := SubStr("0000000000" mylen, -9)
   StringReplace sData, sData, aaaaaaaaaa, %thelen%
   mylen := StrLen(sData) +6 ; was 4
   thelen := SubStr("0000000000" mylen, -9)
   StringReplace sData, sData, bbbbbbbbbb, %thelen%
   mylen :=  StrLen(m_sDescription . sContextStart) +2 
   thelen := SubStr("0000000000" mylen, -9)
   StringReplace sData, sData, cccccccccc, %thelen%
   mylen :=  StrLen(m_sDescription . sContextStart . sHtmlFragment) +6
   thelen := SubStr("0000000000" mylen, -9)
   StringReplace sData, sData, dddddddddd, %thelen%  

   Return, sData
}   


