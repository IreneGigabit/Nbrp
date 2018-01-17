<% 
'發明專利申請書
Function SpaceString()
SpaceString = "<w:p wsp:rsidR=""008515F6"" " & _
"wsp:rsidRDefault=""008515F6"" wsp:rsidP=""008515F6""><w:pPr><w:rPr><w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體""/><wx:font " & _
"wx:val=""新細明體""/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體"" w:hint=""fareast""/>" & _
"<wx:font wx:val=""新細明體""/></w:rPr></w:r></w:p>"
End function


Function DocHead_1()
DocHead_1 = "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>" & _
"<?mso-application progid=""Word.Document""?>" & _
"<w:wordDocument xmlns:aml=""http://schemas.microsoft.com/aml/2001/core"" xmlns:dt=""uuid:C2F41010-65B3-11d1-A29F-00AA00C14882"" xmlns:ve=""http://schemas.openxmlformats.org/markup-compatibility/2006"" " & _
"xmlns:o=""urn:schemas-microsoft-com:office:office"" xmlns:v=""urn:schemas-microsoft-com:vml"" xmlns:w10=""urn:schemas-microsoft-com:office:word"" " & _
"xmlns:w=""http://schemas.microsoft.com/office/word/2003/wordml"" xmlns:wx=""http://schemas.microsoft.com/office/word/2003/auxHint"" " & _
"xmlns:wsp=""http://schemas.microsoft.com/office/word/2003/wordml/sp2"" xmlns:sl=""http://schemas.microsoft.com/schemaLibrary/2003/core"" " & _
"w:macrosPresent=""no"" w:embeddedObjPresent=""no"" w:ocxPresent=""no"" xml:space=""preserve""><w:ignoreSubtree w:val=""http://schemas.microsoft.com/office/word/2003/wordml/sp2""/>" & _
"<w:fonts><w:defaultFonts w:ascii=""Times New Roman"" w:fareast=""新細明體"" w:h-ansi=""Times " & _
"New Roman"" w:cs=""Times New Roman""/><w:font w:name=""Times New Roman""><w:panose-1 w:val=""02020603050405020304""/><w:charset w:val=""00""/>" & _
"<w:family w:val=""Roman""/><w:pitch w:val=""variable""/><w:sig w:usb-0=""20002A87"" w:usb-1=""80000000"" w:usb-2=""00000008"" w:usb-3=""00000000"" " & _
"w:csb-0=""000001FF"" w:csb-1=""00000000""/></w:font><w:font w:name=""新細明體""><w:altName w:val=""PMingLiU""/><w:panose-1 w:val=""02020300000000000000""/>" & _
"<w:charset w:val=""88""/><w:family w:val=""Roman""/><w:pitch w:val=""variable""/><w:sig w:usb-0=""00000003"" w:usb-1=""080E0000"" w:usb-2=""00000016"" " & _
"w:usb-3=""00000000"" w:csb-0=""00100001"" w:csb-1=""00000000""/></w:font><w:font w:name=""Cambria Math""><w:panose-1 w:val=""02040503050406030204""/>" & _
"<w:charset w:val=""00""/><w:family w:val=""Roman""/><w:pitch w:val=""variable""/><w:sig w:usb-0=""E00002FF"" w:usb-1=""420024FF"" w:usb-2=""00000000"" " & _
"w:usb-3=""00000000"" w:csb-0=""0000019F"" w:csb-1=""00000000""/></w:font><w:font w:name=""Cambria""><w:panose-1 w:val=""02040503050406030204""/>" & _
"<w:charset w:val=""00""/><w:family w:val=""Roman""/><w:pitch w:val=""variable""/><w:sig w:usb-0=""E00002FF"" w:usb-1=""400004FF"" w:usb-2=""00000000"" " & _
"w:usb-3=""00000000"" w:csb-0=""0000019F"" w:csb-1=""00000000""/></w:font><w:font w:name=""Calibri""><w:panose-1 w:val=""020F0502020204030204""/>" & _
"<w:charset w:val=""00""/><w:family w:val=""Swiss""/><w:pitch w:val=""variable""/><w:sig w:usb-0=""E10002FF"" w:usb-1=""4000ACFF"" w:usb-2=""00000009"" " & _
"w:usb-3=""00000000"" w:csb-0=""0000019F"" w:csb-1=""00000000""/></w:font><w:font w:name=""@新細明體""><w:panose-1 w:val=""02020300000000000000""/>" & _
"<w:charset w:val=""88""/><w:family w:val=""Roman""/><w:pitch w:val=""variable""/><w:sig w:usb-0=""00000003"" w:usb-1=""080E0000"" w:usb-2=""00000016"" " & _
"w:usb-3=""00000000"" w:csb-0=""00100001"" w:csb-1=""00000000""/></w:font></w:fonts><w:lists>" & _
"<w:listDef w:listDefId=""0""><w:lsid w:val=""05ED05A0""/>" & _
"<w:plt w:val=""HybridMultilevel""/><w:tmpl w:val=""675C8C10""/><w:lvl w:ilvl=""0"" w:tplc=""75104D78""><w:start w:val=""1""/><w:lvlText " & _
"w:val=""【主張優先權】""/><w:lvlJc w:val=""left""/><w:pPr><w:ind w:left=""480"" w:hanging=""480""/></w:pPr><w:rPr><w:rFonts " & _
"w:ascii=""Times New Roman"" w:fareast=""新細明體"" w:h-ansi=""Times New Roman"" w:hint=""fareast""/><w:b w:val=""off""/><w:i w:val=""off""/>" & _
"<w:sz w:val=""24""/></w:rPr></w:lvl><w:lvl w:ilvl=""1"" w:tplc=""04090019""><w:start w:val=""1""/><w:nfc w:val=""30""/><w:lvlText w:val=""%2、""/>" & _
"<w:lvlJc w:val=""left""/><w:pPr><w:ind w:left=""960"" w:hanging=""480""/></w:pPr></w:lvl><w:lvl w:ilvl=""2"" w:tplc=""0409001B""><w:start " & _
"w:val=""1""/><w:nfc w:val=""2""/><w:lvlText w:val=""%3.""/><w:lvlJc w:val=""right""/><w:pPr><w:ind w:left=""1440"" w:hanging=""480""/>" & _
"</w:pPr></w:lvl><w:lvl w:ilvl=""3"" w:tplc=""0409000F""><w:start w:val=""1""/><w:lvlText w:val=""%4.""/><w:lvlJc w:val=""left""/><w:pPr>" & _
"<w:ind w:left=""1920"" w:hanging=""480""/></w:pPr></w:lvl><w:lvl w:ilvl=""4"" w:tplc=""04090019""><w:start w:val=""1""/><w:nfc w:val=""30""/>" & _
"<w:lvlText w:val=""%5、""/><w:lvlJc w:val=""left""/><w:pPr><w:ind w:left=""2400"" w:hanging=""480""/></w:pPr></w:lvl><w:lvl w:ilvl=""5"" " & _
"w:tplc=""0409001B""><w:start w:val=""1""/><w:nfc w:val=""2""/><w:lvlText w:val=""%6.""/><w:lvlJc w:val=""right""/><w:pPr><w:ind w:left=""2880"" " & _
"w:hanging=""480""/></w:pPr></w:lvl><w:lvl w:ilvl=""6"" w:tplc=""0409000F""><w:start w:val=""1""/><w:lvlText w:val=""%7.""/><w:lvlJc " & _
"w:val=""left""/><w:pPr><w:ind w:left=""3360"" w:hanging=""480""/></w:pPr></w:lvl><w:lvl w:ilvl=""7"" w:tplc=""04090019""><w:start " & _
"w:val=""1""/><w:nfc w:val=""30""/><w:lvlText w:val=""%8、""/><w:lvlJc w:val=""left""/><w:pPr><w:ind w:left=""3840"" w:hanging=""480""/>" & _
"</w:pPr></w:lvl><w:lvl w:ilvl=""8"" w:tplc=""0409001B""><w:start w:val=""1""/><w:nfc w:val=""2""/><w:lvlText w:val=""%9.""/><w:lvlJc " & _
"w:val=""right""/><w:pPr><w:ind w:left=""4320"" w:hanging=""480""/></w:pPr></w:lvl></w:listDef>" & _
"<w:listDef w:listDefId=""1""><w:lsid " & _
"w:val=""0E0A38EF""/><w:plt w:val=""HybridMultilevel""/><w:tmpl w:val=""2A02EAF2""/><w:lvl w:ilvl=""0"" w:tplc=""DD3253EE""><w:start " & _
"w:val=""1""/><w:lvlText w:val=""【代理人】""/><w:lvlJc w:val=""left""/><w:pPr><w:ind w:left=""480"" w:hanging=""480""/></w:pPr>" & _
"<w:rPr><w:rFonts w:ascii=""Times New Roman"" w:fareast=""新細明體"" w:h-ansi=""Times New Roman"" w:hint=""fareast""/><w:b w:val=""off""/>" & _
"<w:i w:val=""off""/><w:sz w:val=""24""/></w:rPr></w:lvl><w:lvl w:ilvl=""1"" w:tplc=""04090019""><w:start w:val=""1""/><w:nfc w:val=""30""/>" & _
"<w:lvlText w:val=""%2、""/><w:lvlJc w:val=""left""/><w:pPr><w:ind w:left=""960"" w:hanging=""480""/></w:pPr></w:lvl><w:lvl w:ilvl=""2"" " & _
"w:tplc=""0409001B""><w:start w:val=""1""/><w:nfc w:val=""2""/><w:lvlText w:val=""%3.""/><w:lvlJc w:val=""right""/><w:pPr><w:ind w:left=""1440"" " & _
"w:hanging=""480""/></w:pPr></w:lvl><w:lvl w:ilvl=""3"" w:tplc=""0409000F""><w:start w:val=""1""/><w:lvlText w:val=""%4.""/><w:lvlJc " & _
"w:val=""left""/><w:pPr><w:ind w:left=""1920"" w:hanging=""480""/></w:pPr></w:lvl><w:lvl w:ilvl=""4"" w:tplc=""04090019""><w:start " & _
"w:val=""1""/><w:nfc w:val=""30""/><w:lvlText w:val=""%5、""/><w:lvlJc w:val=""left""/><w:pPr><w:ind w:left=""2400"" w:hanging=""480""/>" & _
"</w:pPr></w:lvl><w:lvl w:ilvl=""5"" w:tplc=""0409001B""><w:start w:val=""1""/><w:nfc w:val=""2""/><w:lvlText w:val=""%6.""/><w:lvlJc " & _
"w:val=""right""/><w:pPr><w:ind w:left=""2880"" w:hanging=""480""/></w:pPr></w:lvl><w:lvl w:ilvl=""6"" w:tplc=""0409000F""><w:start " & _
"w:val=""1""/><w:lvlText w:val=""%7.""/><w:lvlJc w:val=""left""/><w:pPr><w:ind w:left=""3360"" w:hanging=""480""/></w:pPr></w:lvl><w:lvl " & _
"w:ilvl=""7"" w:tplc=""04090019""><w:start w:val=""1""/><w:nfc w:val=""30""/><w:lvlText w:val=""%8、""/><w:lvlJc w:val=""left""/><w:pPr>" & _
"<w:ind w:left=""3840"" w:hanging=""480""/></w:pPr></w:lvl><w:lvl w:ilvl=""8"" w:tplc=""0409001B""><w:start w:val=""1""/><w:nfc w:val=""2""/>" & _
"<w:lvlText w:val=""%9.""/><w:lvlJc w:val=""right""/><w:pPr><w:ind w:left=""4320"" w:hanging=""480""/></w:pPr></w:lvl></w:listDef>" & _

"<w:listDef w:listDefId=""2""><w:lsid w:val=""249D468C""/><w:plt w:val=""HybridMultilevel""/><w:tmpl w:val=""C270F68A""/><w:lvl w:ilvl=""0"" w:tplc=""07000BB0"">" & _
"<w:start w:val=""1""/><w:lvlText w:val=""【主張利用生物材料%1】""/><w:lvlJc w:val=""left""/><w:pPr><w:ind w:left=""480"" " & _
"w:hanging=""480""/></w:pPr><w:rPr><w:rFonts w:ascii=""Times New Roman"" w:fareast=""新細明體"" w:h-ansi=""Times New Roman"" w:hint=""fareast""/>" & _
"<w:b w:val=""off""/><w:i w:val=""off""/><w:sz w:val=""24""/></w:rPr></w:lvl><w:lvl w:ilvl=""1"" w:tplc=""04090019""><w:start w:val=""1""/>" & _
"<w:nfc w:val=""30""/><w:lvlText w:val=""%2、""/><w:lvlJc w:val=""left""/><w:pPr><w:ind w:left=""960"" w:hanging=""480""/></w:pPr></w:lvl>" & _
"<w:lvl w:ilvl=""2"" w:tplc=""0409001B""><w:start w:val=""1""/><w:nfc w:val=""2""/><w:lvlText w:val=""%3.""/><w:lvlJc w:val=""right""/>" & _
"<w:pPr><w:ind w:left=""1440"" w:hanging=""480""/></w:pPr></w:lvl><w:lvl w:ilvl=""3"" w:tplc=""0409000F""><w:start w:val=""1""/><w:lvlText " & _
"w:val=""%4.""/><w:lvlJc w:val=""left""/><w:pPr><w:ind w:left=""1920"" w:hanging=""480""/></w:pPr></w:lvl><w:lvl w:ilvl=""4"" w:tplc=""04090019"">" & _
"<w:start w:val=""1""/><w:nfc w:val=""30""/><w:lvlText w:val=""%5、""/><w:lvlJc w:val=""left""/><w:pPr><w:ind w:left=""2400"" w:hanging=""480""/>" & _
"</w:pPr></w:lvl><w:lvl w:ilvl=""5"" w:tplc=""0409001B""><w:start w:val=""1""/><w:nfc w:val=""2""/><w:lvlText w:val=""%6.""/><w:lvlJc " & _
"w:val=""right""/><w:pPr><w:ind w:left=""2880"" w:hanging=""480""/></w:pPr></w:lvl><w:lvl w:ilvl=""6"" w:tplc=""0409000F""><w:start " & _
"w:val=""1""/><w:lvlText w:val=""%7.""/><w:lvlJc w:val=""left""/><w:pPr><w:ind w:left=""3360"" w:hanging=""480""/></w:pPr></w:lvl><w:lvl " & _
"w:ilvl=""7"" w:tplc=""04090019""><w:start w:val=""1""/><w:nfc w:val=""30""/><w:lvlText w:val=""%8、""/><w:lvlJc w:val=""left""/><w:pPr>" & _
"<w:ind w:left=""3840"" w:hanging=""480""/></w:pPr></w:lvl><w:lvl w:ilvl=""8"" w:tplc=""0409001B""><w:start w:val=""1""/><w:nfc w:val=""2""/>" & _
"<w:lvlText w:val=""%9.""/><w:lvlJc w:val=""right""/><w:pPr><w:ind w:left=""4320"" w:hanging=""480""/></w:pPr></w:lvl></w:listDef>" & _

"<w:listDef w:listDefId=""3""><w:lsid w:val=""29A844E5""/><w:plt w:val=""HybridMultilevel""/><w:tmpl w:val=""128837EA""/><w:lvl w:ilvl=""0"" w:tplc=""0608C004"">" & _
"<w:start w:val=""1""/><w:lvlText w:val=""【主張優惠期%1】""/><w:lvlJc w:val=""left""/><w:pPr><w:ind w:left=""480"" w:hanging=""480""/>" & _
"</w:pPr><w:rPr><w:rFonts w:ascii=""Times New Roman"" w:fareast=""新細明體"" w:h-ansi=""Times New Roman"" w:hint=""fareast""/><w:b " & _
"w:val=""off""/><w:i w:val=""off""/><w:sz w:val=""24""/></w:rPr></w:lvl><w:lvl w:ilvl=""1"" w:tplc=""04090019""><w:start w:val=""1""/>" & _
"<w:nfc w:val=""30""/><w:lvlText w:val=""%2、""/><w:lvlJc w:val=""left""/><w:pPr><w:ind w:left=""960"" w:hanging=""480""/></w:pPr></w:lvl>" & _
"<w:lvl w:ilvl=""2"" w:tplc=""0409001B""><w:start w:val=""1""/><w:nfc w:val=""2""/><w:lvlText w:val=""%3.""/><w:lvlJc w:val=""right""/>" & _
"<w:pPr><w:ind w:left=""1440"" w:hanging=""480""/></w:pPr></w:lvl><w:lvl w:ilvl=""3"" w:tplc=""0409000F""><w:start w:val=""1""/><w:lvlText " & _
"w:val=""%4.""/><w:lvlJc w:val=""left""/><w:pPr><w:ind w:left=""1920"" w:hanging=""480""/></w:pPr></w:lvl><w:lvl w:ilvl=""4"" w:tplc=""04090019"">" & _
"<w:start w:val=""1""/><w:nfc w:val=""30""/><w:lvlText w:val=""%5、""/><w:lvlJc w:val=""left""/><w:pPr><w:ind w:left=""2400"" w:hanging=""480""/>" & _
"</w:pPr></w:lvl><w:lvl w:ilvl=""5"" w:tplc=""0409001B""><w:start w:val=""1""/><w:nfc w:val=""2""/><w:lvlText w:val=""%6.""/><w:lvlJc " & _
"w:val=""right""/><w:pPr><w:ind w:left=""2880"" w:hanging=""480""/></w:pPr></w:lvl><w:lvl w:ilvl=""6"" w:tplc=""0409000F""><w:start " & _
"w:val=""1""/><w:lvlText w:val=""%7.""/><w:lvlJc w:val=""left""/><w:pPr><w:ind w:left=""3360"" w:hanging=""480""/></w:pPr></w:lvl><w:lvl " & _
"w:ilvl=""7"" w:tplc=""04090019""><w:start w:val=""1""/><w:nfc w:val=""30""/><w:lvlText w:val=""%8、""/><w:lvlJc w:val=""left""/><w:pPr>" & _
"<w:ind w:left=""3840"" w:hanging=""480""/></w:pPr></w:lvl><w:lvl w:ilvl=""8"" w:tplc=""0409001B""><w:start w:val=""1""/><w:nfc w:val=""2""/>" & _
"<w:lvlText w:val=""%9.""/><w:lvlJc w:val=""right""/><w:pPr><w:ind w:left=""4320"" w:hanging=""480""/></w:pPr></w:lvl></w:listDef>" & _

"<w:listDef w:listDefId=""4""><w:lsid w:val=""3EAE5824""/><w:plt w:val=""HybridMultilevel""/><w:tmpl w:val=""ADC040D6""/><w:lvl w:ilvl=""0"" w:tplc=""C4801CB4"">" & _
"<w:start w:val=""1""/><w:lvlText w:val=""【申請人】""/><w:lvlJc w:val=""left""/><w:pPr><w:ind w:left=""480"" w:hanging=""480""/>" & _
"</w:pPr><w:rPr><w:rFonts w:ascii=""Times New Roman"" w:fareast=""新細明體"" w:h-ansi=""Times New Roman"" w:hint=""fareast""/><w:b " & _
"w:val=""off""/><w:i w:val=""off""/><w:sz w:val=""24""/></w:rPr></w:lvl><w:lvl w:ilvl=""1"" w:tplc=""04090019""><w:start w:val=""1""/>" & _
"<w:nfc w:val=""30""/><w:lvlText w:val=""%2、""/><w:lvlJc w:val=""left""/><w:pPr><w:ind w:left=""960"" w:hanging=""480""/></w:pPr></w:lvl>" & _
"<w:lvl w:ilvl=""2"" w:tplc=""0409001B""><w:start w:val=""1""/><w:nfc w:val=""2""/><w:lvlText w:val=""%3.""/><w:lvlJc w:val=""right""/>" & _
"<w:pPr><w:ind w:left=""1440"" w:hanging=""480""/></w:pPr></w:lvl><w:lvl w:ilvl=""3"" w:tplc=""0409000F""><w:start w:val=""1""/><w:lvlText " & _
"w:val=""%4.""/><w:lvlJc w:val=""left""/><w:pPr><w:ind w:left=""1920"" w:hanging=""480""/></w:pPr></w:lvl><w:lvl w:ilvl=""4"" w:tplc=""04090019"">" & _
"<w:start w:val=""1""/><w:nfc w:val=""30""/><w:lvlText w:val=""%5、""/><w:lvlJc w:val=""left""/><w:pPr><w:ind w:left=""2400"" w:hanging=""480""/>" & _
"</w:pPr></w:lvl><w:lvl w:ilvl=""5"" w:tplc=""0409001B""><w:start w:val=""1""/><w:nfc w:val=""2""/><w:lvlText w:val=""%6.""/><w:lvlJc " & _
"w:val=""right""/><w:pPr><w:ind w:left=""2880"" w:hanging=""480""/></w:pPr></w:lvl><w:lvl w:ilvl=""6"" w:tplc=""0409000F""><w:start " & _
"w:val=""1""/><w:lvlText w:val=""%7.""/><w:lvlJc w:val=""left""/><w:pPr><w:ind w:left=""3360"" w:hanging=""480""/></w:pPr></w:lvl><w:lvl " & _
"w:ilvl=""7"" w:tplc=""04090019""><w:start w:val=""1""/><w:nfc w:val=""30""/><w:lvlText w:val=""%8、""/><w:lvlJc w:val=""left""/><w:pPr>" & _
"<w:ind w:left=""3840"" w:hanging=""480""/></w:pPr></w:lvl><w:lvl w:ilvl=""8"" w:tplc=""0409001B""><w:start w:val=""1""/><w:nfc w:val=""2""/>" & _
"<w:lvlText w:val=""%9.""/><w:lvlJc w:val=""right""/><w:pPr><w:ind w:left=""4320"" w:hanging=""480""/></w:pPr></w:lvl></w:listDef>" & _
"<w:listDef w:listDefId=""5""><w:lsid w:val=""7B0A0C09""/><w:plt w:val=""HybridMultilevel""/><w:tmpl w:val=""D1903E14""/><w:lvl w:ilvl=""0"" w:tplc=""A5948AC2"">" & _
"<w:start w:val=""1""/><w:lvlText w:val=""【發明人】""/><w:lvlJc w:val=""left""/><w:pPr><w:ind w:left=""480"" w:hanging=""480""/>" & _
"</w:pPr><w:rPr><w:rFonts w:ascii=""Times New Roman"" w:fareast=""新細明體"" w:h-ansi=""Times New Roman"" w:hint=""fareast""/><w:b " & _
"w:val=""off""/><w:i w:val=""off""/><w:sz w:val=""24""/></w:rPr></w:lvl><w:lvl w:ilvl=""1"" w:tplc=""04090019""><w:start w:val=""1""/>" & _
"<w:nfc w:val=""30""/><w:lvlText w:val=""%2、""/><w:lvlJc w:val=""left""/><w:pPr><w:ind w:left=""960"" w:hanging=""480""/></w:pPr></w:lvl>" & _
"<w:lvl w:ilvl=""2"" w:tplc=""0409001B""><w:start w:val=""1""/><w:nfc w:val=""2""/><w:lvlText w:val=""%3.""/><w:lvlJc w:val=""right""/>" & _
"<w:pPr><w:ind w:left=""1440"" w:hanging=""480""/></w:pPr></w:lvl><w:lvl w:ilvl=""3"" w:tplc=""0409000F""><w:start w:val=""1""/><w:lvlText " & _
"w:val=""%4.""/><w:lvlJc w:val=""left""/><w:pPr><w:ind w:left=""1920"" w:hanging=""480""/></w:pPr></w:lvl><w:lvl w:ilvl=""4"" w:tplc=""04090019"">" & _
"<w:start w:val=""1""/><w:nfc w:val=""30""/><w:lvlText w:val=""%5、""/><w:lvlJc w:val=""left""/><w:pPr><w:ind w:left=""2400"" w:hanging=""480""/>" & _
"</w:pPr></w:lvl><w:lvl w:ilvl=""5"" w:tplc=""0409001B""><w:start w:val=""1""/><w:nfc w:val=""2""/><w:lvlText w:val=""%6.""/><w:lvlJc " & _
"w:val=""right""/><w:pPr><w:ind w:left=""2880"" w:hanging=""480""/></w:pPr></w:lvl><w:lvl w:ilvl=""6"" w:tplc=""0409000F""><w:start " & _
"w:val=""1""/><w:lvlText w:val=""%7.""/><w:lvlJc w:val=""left""/><w:pPr><w:ind w:left=""3360"" w:hanging=""480""/></w:pPr></w:lvl><w:lvl " & _
"w:ilvl=""7"" w:tplc=""04090019""><w:start w:val=""1""/><w:nfc w:val=""30""/><w:lvlText w:val=""%8、""/><w:lvlJc w:val=""left""/><w:pPr>" & _
"<w:ind w:left=""3840"" w:hanging=""480""/></w:pPr></w:lvl><w:lvl w:ilvl=""8"" w:tplc=""0409001B""><w:start w:val=""1""/><w:nfc w:val=""2""/>" & _
"<w:lvlText w:val=""%9.""/><w:lvlJc w:val=""right""/><w:pPr><w:ind w:left=""4320"" w:hanging=""480""/></w:pPr></w:lvl></w:listDef><w:list " & _
"w:ilfo=""1""><w:ilst w:val=""4""/></w:list><w:list w:ilfo=""2""><w:ilst w:val=""4""/><w:lvlOverride w:ilvl=""0""><w:startOverride " & _
"w:val=""1""/></w:lvlOverride><w:lvlOverride w:ilvl=""1""><w:startOverride w:val=""1""/></w:lvlOverride><w:lvlOverride w:ilvl=""2"">" & _
"<w:startOverride w:val=""1""/></w:lvlOverride><w:lvlOverride w:ilvl=""3""><w:startOverride w:val=""1""/></w:lvlOverride><w:lvlOverride " & _
"w:ilvl=""4""><w:startOverride w:val=""1""/></w:lvlOverride><w:lvlOverride w:ilvl=""5""><w:startOverride w:val=""1""/></w:lvlOverride>" & _
"<w:lvlOverride w:ilvl=""6""><w:startOverride w:val=""1""/></w:lvlOverride><w:lvlOverride w:ilvl=""7""><w:startOverride w:val=""1""/>" & _
"</w:lvlOverride><w:lvlOverride w:ilvl=""8""><w:startOverride w:val=""1""/></w:lvlOverride></w:list><w:list w:ilfo=""3""><w:ilst w:val=""1""/>" & _
"</w:list><w:list w:ilfo=""4""><w:ilst w:val=""1""/><w:lvlOverride w:ilvl=""0""><w:startOverride w:val=""1""/></w:lvlOverride><w:lvlOverride " & _
"w:ilvl=""1""><w:startOverride w:val=""1""/></w:lvlOverride><w:lvlOverride w:ilvl=""2""><w:startOverride w:val=""1""/></w:lvlOverride>" & _
"<w:lvlOverride w:ilvl=""3""><w:startOverride w:val=""1""/></w:lvlOverride><w:lvlOverride w:ilvl=""4""><w:startOverride w:val=""1""/>" & _
"</w:lvlOverride><w:lvlOverride w:ilvl=""5""><w:startOverride w:val=""1""/></w:lvlOverride><w:lvlOverride w:ilvl=""6""><w:startOverride " & _
"w:val=""1""/></w:lvlOverride><w:lvlOverride w:ilvl=""7""><w:startOverride w:val=""1""/></w:lvlOverride><w:lvlOverride w:ilvl=""8"">" & _
"<w:startOverride w:val=""1""/></w:lvlOverride></w:list><w:list w:ilfo=""5""><w:ilst w:val=""5""/></w:list><w:list w:ilfo=""6""><w:ilst " & _
"w:val=""5""/><w:lvlOverride w:ilvl=""0""><w:startOverride w:val=""1""/></w:lvlOverride><w:lvlOverride w:ilvl=""1""><w:startOverride " & _
"w:val=""1""/></w:lvlOverride><w:lvlOverride w:ilvl=""2""><w:startOverride w:val=""1""/></w:lvlOverride><w:lvlOverride w:ilvl=""3"">" & _
"<w:startOverride w:val=""1""/></w:lvlOverride><w:lvlOverride w:ilvl=""4""><w:startOverride w:val=""1""/></w:lvlOverride><w:lvlOverride " & _
"w:ilvl=""5""><w:startOverride w:val=""1""/></w:lvlOverride><w:lvlOverride w:ilvl=""6""><w:startOverride w:val=""1""/></w:lvlOverride>" & _
"<w:lvlOverride w:ilvl=""7""><w:startOverride w:val=""1""/></w:lvlOverride><w:lvlOverride w:ilvl=""8""><w:startOverride w:val=""1""/>" & _
"</w:lvlOverride></w:list><w:list w:ilfo=""7""><w:ilst w:val=""3""/></w:list><w:list w:ilfo=""8""><w:ilst w:val=""3""/><w:lvlOverride " & _
"w:ilvl=""0""><w:startOverride w:val=""1""/></w:lvlOverride><w:lvlOverride w:ilvl=""1""><w:startOverride w:val=""1""/></w:lvlOverride>" & _
"<w:lvlOverride w:ilvl=""2""><w:startOverride w:val=""1""/></w:lvlOverride><w:lvlOverride w:ilvl=""3""><w:startOverride w:val=""1""/>" & _
"</w:lvlOverride><w:lvlOverride w:ilvl=""4""><w:startOverride w:val=""1""/></w:lvlOverride><w:lvlOverride w:ilvl=""5""><w:startOverride " & _
"w:val=""1""/></w:lvlOverride><w:lvlOverride w:ilvl=""6""><w:startOverride w:val=""1""/></w:lvlOverride><w:lvlOverride w:ilvl=""7"">" & _
"<w:startOverride w:val=""1""/></w:lvlOverride><w:lvlOverride w:ilvl=""8""><w:startOverride w:val=""1""/></w:lvlOverride></w:list><w:list " & _
"w:ilfo=""9""><w:ilst w:val=""0""/></w:list><w:list w:ilfo=""10""><w:ilst w:val=""0""/><w:lvlOverride w:ilvl=""0""><w:startOverride " & _
"w:val=""1""/></w:lvlOverride><w:lvlOverride w:ilvl=""1""><w:startOverride w:val=""1""/></w:lvlOverride><w:lvlOverride w:ilvl=""2"">" & _
"<w:startOverride w:val=""1""/></w:lvlOverride><w:lvlOverride w:ilvl=""3""><w:startOverride w:val=""1""/></w:lvlOverride><w:lvlOverride " & _
"w:ilvl=""4""><w:startOverride w:val=""1""/></w:lvlOverride><w:lvlOverride w:ilvl=""5""><w:startOverride w:val=""1""/></w:lvlOverride>" & _
"<w:lvlOverride w:ilvl=""6""><w:startOverride w:val=""1""/></w:lvlOverride><w:lvlOverride w:ilvl=""7""><w:startOverride w:val=""1""/>" & _
"</w:lvlOverride><w:lvlOverride w:ilvl=""8""><w:startOverride w:val=""1""/></w:lvlOverride></w:list><w:list w:ilfo=""11""><w:ilst w:val=""2""/>" & _
"</w:list><w:list w:ilfo=""12""><w:ilst w:val=""2""/><w:lvlOverride w:ilvl=""0""><w:startOverride w:val=""1""/></w:lvlOverride><w:lvlOverride " & _
"w:ilvl=""1""><w:startOverride w:val=""1""/></w:lvlOverride><w:lvlOverride w:ilvl=""2""><w:startOverride w:val=""1""/></w:lvlOverride>" & _
"<w:lvlOverride w:ilvl=""3""><w:startOverride w:val=""1""/></w:lvlOverride><w:lvlOverride w:ilvl=""4""><w:startOverride w:val=""1""/>" & _
"</w:lvlOverride><w:lvlOverride w:ilvl=""5""><w:startOverride w:val=""1""/></w:lvlOverride><w:lvlOverride w:ilvl=""6""><w:startOverride " & _
"w:val=""1""/></w:lvlOverride><w:lvlOverride w:ilvl=""7""><w:startOverride w:val=""1""/></w:lvlOverride><w:lvlOverride w:ilvl=""8"">" & _
"<w:startOverride w:val=""1""/></w:lvlOverride></w:list></w:lists><w:styles><w:versionOfBuiltInStylenames w:val=""7""/><w:latentStyles " & _
"w:defLockedState=""off"" w:latentStyleCount=""267""><w:lsdException w:name=""Normal""/><w:lsdException w:name=""heading 1""/><w:lsdException " & _
"w:name=""heading 2""/><w:lsdException w:name=""heading 3""/><w:lsdException w:name=""heading 4""/><w:lsdException w:name=""heading " & _
"5""/><w:lsdException w:name=""heading 6""/><w:lsdException w:name=""heading 7""/><w:lsdException w:name=""heading 8""/><w:lsdException " & _
"w:name=""heading 9""/><w:lsdException w:name=""toc 1""/><w:lsdException w:name=""toc 2""/><w:lsdException w:name=""toc 3""/><w:lsdException " & _
"w:name=""toc 4""/><w:lsdException w:name=""toc 5""/><w:lsdException w:name=""toc 6""/><w:lsdException w:name=""toc 7""/><w:lsdException " & _
"w:name=""toc 8""/><w:lsdException w:name=""toc 9""/><w:lsdException w:name=""caption""/><w:lsdException w:name=""Title""/><w:lsdException " & _
"w:name=""Default Paragraph Font""/><w:lsdException w:name=""Subtitle""/><w:lsdException w:name=""Hyperlink""/><w:lsdException w:name=""Strong""/>" & _
"<w:lsdException w:name=""Emphasis""/><w:lsdException w:name=""Table Grid""/><w:lsdException w:name=""Placeholder Text""/><w:lsdException " & _
"w:name=""No Spacing""/><w:lsdException w:name=""Light Shading""/><w:lsdException w:name=""Light List""/><w:lsdException w:name=""Light " & _
"Grid""/><w:lsdException w:name=""Medium Shading 1""/><w:lsdException w:name=""Medium Shading 2""/><w:lsdException w:name=""Medium " & _
"List 1""/><w:lsdException w:name=""Medium List 2""/><w:lsdException w:name=""Medium Grid 1""/><w:lsdException w:name=""Medium Grid " & _
"2""/><w:lsdException w:name=""Medium Grid 3""/><w:lsdException w:name=""Dark List""/><w:lsdException w:name=""Colorful Shading""/>" & _
"<w:lsdException w:name=""Colorful List""/><w:lsdException w:name=""Colorful Grid""/><w:lsdException w:name=""Light Shading Accent 1""/>" & _
"<w:lsdException w:name=""Light List Accent 1""/><w:lsdException w:name=""Light Grid Accent 1""/><w:lsdException w:name=""Medium Shading " & _
"1 Accent 1""/><w:lsdException w:name=""Medium Shading 2 Accent 1""/><w:lsdException w:name=""Medium List 1 Accent 1""/><w:lsdException " & _
"w:name=""Revision""/><w:lsdException w:name=""List Paragraph""/><w:lsdException w:name=""Quote""/><w:lsdException w:name=""Intense " & _
"Quote""/><w:lsdException w:name=""Medium List 2 Accent 1""/><w:lsdException w:name=""Medium Grid 1 Accent 1""/><w:lsdException w:name=""Medium " & _
"Grid 2 Accent 1""/><w:lsdException w:name=""Medium Grid 3 Accent 1""/><w:lsdException w:name=""Dark List Accent 1""/><w:lsdException " & _
"w:name=""Colorful Shading Accent 1""/><w:lsdException w:name=""Colorful List Accent 1""/><w:lsdException w:name=""Colorful Grid Accent " & _
"1""/><w:lsdException w:name=""Light Shading Accent 2""/><w:lsdException w:name=""Light List Accent 2""/><w:lsdException w:name=""Light " & _
"Grid Accent 2""/><w:lsdException w:name=""Medium Shading 1 Accent 2""/><w:lsdException w:name=""Medium Shading 2 Accent 2""/><w:lsdException " & _
"w:name=""Medium List 1 Accent 2""/><w:lsdException w:name=""Medium List 2 Accent 2""/><w:lsdException w:name=""Medium Grid 1 Accent " & _
"2""/><w:lsdException w:name=""Medium Grid 2 Accent 2""/><w:lsdException w:name=""Medium Grid 3 Accent 2""/><w:lsdException w:name=""Dark " & _
"List Accent 2""/><w:lsdException w:name=""Colorful Shading Accent 2""/><w:lsdException w:name=""Colorful List Accent 2""/><w:lsdException " & _
"w:name=""Colorful Grid Accent 2""/><w:lsdException w:name=""Light Shading Accent 3""/><w:lsdException w:name=""Light List Accent 3""/>" & _
"<w:lsdException w:name=""Light Grid Accent 3""/><w:lsdException w:name=""Medium Shading 1 Accent 3""/><w:lsdException w:name=""Medium " & _
"Shading 2 Accent 3""/><w:lsdException w:name=""Medium List 1 Accent 3""/><w:lsdException w:name=""Medium List 2 Accent 3""/><w:lsdException " & _
"w:name=""Medium Grid 1 Accent 3""/><w:lsdException w:name=""Medium Grid 2 Accent 3""/><w:lsdException w:name=""Medium Grid 3 Accent " & _
"3""/><w:lsdException w:name=""Dark List Accent 3""/><w:lsdException w:name=""Colorful Shading Accent 3""/><w:lsdException w:name=""Colorful " & _
"List Accent 3""/><w:lsdException w:name=""Colorful Grid Accent 3""/><w:lsdException w:name=""Light Shading Accent 4""/><w:lsdException " & _
"w:name=""Light List Accent 4""/><w:lsdException w:name=""Light Grid Accent 4""/><w:lsdException w:name=""Medium Shading 1 Accent 4""/>" & _
"<w:lsdException w:name=""Medium Shading 2 Accent 4""/><w:lsdException w:name=""Medium List 1 Accent 4""/><w:lsdException w:name=""Medium " & _
"List 2 Accent 4""/><w:lsdException w:name=""Medium Grid 1 Accent 4""/><w:lsdException w:name=""Medium Grid 2 Accent 4""/><w:lsdException " & _
"w:name=""Medium Grid 3 Accent 4""/><w:lsdException w:name=""Dark List Accent 4""/><w:lsdException w:name=""Colorful Shading Accent " & _
"4""/><w:lsdException w:name=""Colorful List Accent 4""/><w:lsdException w:name=""Colorful Grid Accent 4""/><w:lsdException w:name=""Light " & _
"Shading Accent 5""/><w:lsdException w:name=""Light List Accent 5""/><w:lsdException w:name=""Light Grid Accent 5""/><w:lsdException " & _
"w:name=""Medium Shading 1 Accent 5""/><w:lsdException w:name=""Medium Shading 2 Accent 5""/><w:lsdException w:name=""Medium List 1 " & _
"Accent 5""/><w:lsdException w:name=""Medium List 2 Accent 5""/><w:lsdException w:name=""Medium Grid 1 Accent 5""/><w:lsdException " & _
"w:name=""Medium Grid 2 Accent 5""/><w:lsdException w:name=""Medium Grid 3 Accent 5""/><w:lsdException w:name=""Dark List Accent 5""/>" & _
"<w:lsdException w:name=""Colorful Shading Accent 5""/><w:lsdException w:name=""Colorful List Accent 5""/><w:lsdException w:name=""Colorful " & _
"Grid Accent 5""/><w:lsdException w:name=""Light Shading Accent 6""/><w:lsdException w:name=""Light List Accent 6""/><w:lsdException " & _
"w:name=""Light Grid Accent 6""/><w:lsdException w:name=""Medium Shading 1 Accent 6""/><w:lsdException w:name=""Medium Shading 2 Accent " & _
"6""/><w:lsdException w:name=""Medium List 1 Accent 6""/><w:lsdException w:name=""Medium List 2 Accent 6""/><w:lsdException w:name=""Medium " & _
"Grid 1 Accent 6""/><w:lsdException w:name=""Medium Grid 2 Accent 6""/><w:lsdException w:name=""Medium Grid 3 Accent 6""/><w:lsdException " & _
"w:name=""Dark List Accent 6""/><w:lsdException w:name=""Colorful Shading Accent 6""/><w:lsdException w:name=""Colorful List Accent " & _
"6""/><w:lsdException w:name=""Colorful Grid Accent 6""/><w:lsdException w:name=""Subtle Emphasis""/><w:lsdException w:name=""Intense " & _
"Emphasis""/><w:lsdException w:name=""Subtle Reference""/><w:lsdException w:name=""Intense Reference""/><w:lsdException w:name=""Book " & _
"Title""/><w:lsdException w:name=""Bibliography""/><w:lsdException w:name=""TOC Heading""/></w:latentStyles><w:style w:type=""paragraph"" " & _
"w:default=""on"" w:styleId=""a""><w:name w:val=""Normal""/><wx:uiName wx:val=""內文""/><w:rPr><w:rFonts w:ascii=""Calibri"" w:h-ansi=""Calibri"" " & _
"w:cs=""新細明體""/><wx:font wx:val=""Calibri""/><w:sz w:val=""24""/><w:sz-cs w:val=""24""/><w:lang w:val=""EN-US"" w:fareast=""ZH-TW"" " & _
"w:bidi=""AR-SA""/></w:rPr></w:style><w:style w:type=""character"" w:default=""on"" w:styleId=""a0""><w:name w:val=""Default Paragraph " & _
"Font""/><wx:uiName wx:val=""預設段落字型""/></w:style><w:style w:type=""table"" w:default=""on"" w:styleId=""a1""><w:name w:val=""Normal " & _
"Table""/><wx:uiName wx:val=""表格內文""/><w:rPr><wx:font wx:val=""Times New Roman""/><w:lang w:val=""EN-US"" w:fareast=""ZH-TW"" " & _
"w:bidi=""AR-SA""/></w:rPr><w:tblPr><w:tblInd w:w=""0"" w:type=""dxa""/><w:tblCellMar><w:top w:w=""0"" w:type=""dxa""/><w:left w:w=""108"" " & _
"w:type=""dxa""/><w:bottom w:w=""0"" w:type=""dxa""/><w:right w:w=""108"" w:type=""dxa""/></w:tblCellMar></w:tblPr></w:style><w:style " & _
"w:type=""list"" w:default=""on"" w:styleId=""a2""><w:name w:val=""No List""/><wx:uiName wx:val=""無清單""/></w:style><w:style w:type=""character"" " & _
"w:styleId=""a3""><w:name w:val=""Hyperlink""/><wx:uiName wx:val=""超連結""/><w:rPr><w:color w:val=""0000FF""/><w:u w:val=""single""/>" & _
"</w:rPr></w:style><w:style w:type=""character"" w:styleId=""a4""><w:name w:val=""FollowedHyperlink""/><wx:uiName wx:val=""已查閱的超連結""/>" & _
"<w:rPr><w:color w:val=""800080""/><w:u w:val=""single""/></w:rPr></w:style><w:style w:type=""paragraph"" w:styleId=""a5""><w:name w:val=""header""/>" & _
"<wx:uiName wx:val=""頁首""/><w:basedOn w:val=""a""/><w:link w:val=""a6""/><w:pPr><w:snapToGrid w:val=""off""/></w:pPr><w:rPr><wx:font " & _
"wx:val=""Calibri""/><w:sz w:val=""20""/><w:sz-cs w:val=""20""/></w:rPr></w:style><w:style w:type=""character"" w:styleId=""a6""><w:name " & _
"w:val=""頁首 字元""/><w:basedOn w:val=""a0""/><w:link w:val=""a5""/><w:locked/></w:style><w:style w:type=""paragraph"" w:styleId=""a7"">" & _
"<w:name w:val=""footer""/><wx:uiName wx:val=""頁尾""/><w:basedOn w:val=""a""/><w:link w:val=""a8""/><w:pPr><w:snapToGrid w:val=""off""/>" & _
"</w:pPr><w:rPr><wx:font wx:val=""Calibri""/><w:sz w:val=""20""/><w:sz-cs w:val=""20""/></w:rPr></w:style><w:style w:type=""character"" " & _
"w:styleId=""a8""><w:name w:val=""頁尾 字元""/><w:basedOn w:val=""a0""/><w:link w:val=""a7""/><w:locked/></w:style><w:style w:type=""paragraph"" " & _
"w:styleId=""a9""><w:name w:val=""Balloon Text""/><wx:uiName wx:val=""註解方塊文字""/><w:basedOn w:val=""a""/><w:link w:val=""aa""/>" & _
"<w:rPr><w:rFonts w:ascii=""Cambria"" w:h-ansi=""Cambria"" w:cs=""Times New Roman""/><wx:font wx:val=""Cambria""/><w:sz w:val=""20""/>" & _
"<w:sz-cs w:val=""20""/><w:lang/></w:rPr></w:style><w:style w:type=""character"" w:styleId=""aa""><w:name w:val=""註解方塊文字 " & _
"字元""/><w:link w:val=""a9""/><w:locked/><w:rPr><w:rFonts w:ascii=""Cambria"" w:h-ansi=""Cambria"" w:hint=""default""/></w:rPr></w:style>" & _
"<w:style w:type=""paragraph"" w:styleId=""ab""><w:name w:val=""Revision""/><wx:uiName wx:val=""修訂""/><w:rPr><w:rFonts w:ascii=""Calibri"" " & _
"w:h-ansi=""Calibri"" w:cs=""新細明體""/><wx:font wx:val=""Calibri""/><w:sz w:val=""24""/><w:sz-cs w:val=""24""/><w:lang w:val=""EN-US"" " & _
"w:fareast=""ZH-TW"" w:bidi=""AR-SA""/></w:rPr></w:style><w:style w:type=""paragraph"" w:styleId=""ac""><w:name w:val=""List Paragraph""/>" & _
"<wx:uiName wx:val=""清單段落""/><w:basedOn w:val=""a""/><w:pPr><w:ind w:left=""480""/></w:pPr><w:rPr><wx:font wx:val=""Calibri""/>" & _
"</w:rPr></w:style><w:style w:type=""paragraph"" w:styleId=""msochpdefault""><w:name w:val=""msochpdefault""/><w:basedOn w:val=""a""/>" & _
"<w:pPr><w:spacing w:before=""100"" w:before-autospacing=""on"" w:after=""100"" w:after-autospacing=""on""/></w:pPr><w:rPr><w:rFonts " & _
"w:ascii=""新細明體"" w:h-ansi=""新細明體""/><wx:font wx:val=""新細明體""/><w:sz w:val=""20""/><w:sz-cs w:val=""20""/></w:rPr>" & _
"</w:style></w:styles><w:shapeDefaults><o:shapedefaults v:ext=""edit"" spidmax=""3074""/><o:shapelayout v:ext=""edit""><o:idmap v:ext=""edit"" " & _
"data=""1""/></o:shapelayout></w:shapeDefaults><w:docPr><w:view w:val=""print""/><w:zoom w:percent=""90""/><w:doNotEmbedSystemFonts/>" & _
"<w:bordersDontSurroundHeader/><w:bordersDontSurroundFooter/><w:proofState w:grammar=""clean""/><w:defaultTabStop w:val=""480""/><w:characterSpacingControl " & _
"w:val=""CompressPunctuation""/><w:optimizeForBrowser/><w:targetScreenSz w:val=""1024x768""/><w:validateAgainstSchema/><w:saveInvalidXML " & _
"w:val=""off""/><w:ignoreMixedContent w:val=""off""/><w:alwaysShowPlaceholderText w:val=""off""/><w:hdrShapeDefaults><o:shapedefaults " & _
"v:ext=""edit"" spidmax=""3074""/></w:hdrShapeDefaults><w:footnotePr><w:footnote w:type=""separator""><w:p wsp:rsidR=""00BF474D"" wsp:rsidRDefault=""00BF474D"">" & _
"<w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type=""continuation-separator""><w:p wsp:rsidR=""00BF474D"" wsp:rsidRDefault=""00BF474D"">" & _
"<w:r><w:continuationSeparator/></w:r></w:p></w:footnote></w:footnotePr><w:endnotePr><w:endnote w:type=""separator""><w:p wsp:rsidR=""00BF474D"" " & _
"wsp:rsidRDefault=""00BF474D""><w:r><w:separator/></w:r></w:p></w:endnote><w:endnote w:type=""continuation-separator""><w:p wsp:rsidR=""00BF474D"" " & _
"wsp:rsidRDefault=""00BF474D""><w:r><w:continuationSeparator/></w:r></w:p></w:endnote></w:endnotePr><w:compat><w:breakWrappedTables/>" & _
"<w:useFELayout/></w:compat><wsp:rsids><wsp:rsidRoot wsp:val=""000940B1""/><wsp:rsid wsp:val=""000940B1""/><wsp:rsid wsp:val=""00097D83""/>" & _
"<wsp:rsid wsp:val=""000A07BC""/><wsp:rsid wsp:val=""001548CA""/><wsp:rsid wsp:val=""00173215""/><wsp:rsid wsp:val=""0017701A""/><wsp:rsid " & _
"wsp:val=""001B4EDE""/><wsp:rsid wsp:val=""001D5812""/><wsp:rsid wsp:val=""002563D4""/><wsp:rsid wsp:val=""002831B1""/><wsp:rsid wsp:val=""00290EC6""/>" & _
"<wsp:rsid wsp:val=""002E6DC4""/><wsp:rsid wsp:val=""003067C7""/><wsp:rsid wsp:val=""00325C57""/><wsp:rsid wsp:val=""003304A1""/><wsp:rsid " & _
"wsp:val=""0037287A""/><wsp:rsid wsp:val=""00397393""/><wsp:rsid wsp:val=""00397406""/><wsp:rsid wsp:val=""003D0D41""/><wsp:rsid wsp:val=""004542AC""/>" & _
"<wsp:rsid wsp:val=""004C1689""/><wsp:rsid wsp:val=""005029F5""/><wsp:rsid wsp:val=""00517763""/><wsp:rsid wsp:val=""00521150""/><wsp:rsid " & _
"wsp:val=""00534587""/><wsp:rsid wsp:val=""005476FC""/><wsp:rsid wsp:val=""005477FE""/><wsp:rsid wsp:val=""0057572F""/><wsp:rsid wsp:val=""007B726F""/>" & _
"<wsp:rsid wsp:val=""007F3823""/><wsp:rsid wsp:val=""00802186""/><wsp:rsid wsp:val=""008406A1""/><wsp:rsid wsp:val=""00840A0E""/><wsp:rsid " & _
"wsp:val=""00881A76""/><wsp:rsid wsp:val=""00890066""/><wsp:rsid wsp:val=""008E1167""/><wsp:rsid wsp:val=""008E4F69""/><wsp:rsid wsp:val=""00975CB3""/>" & _
"<wsp:rsid wsp:val=""009E422D""/><wsp:rsid wsp:val=""00A53646""/><wsp:rsid wsp:val=""00A55D33""/><wsp:rsid wsp:val=""00A73EC2""/><wsp:rsid " & _
"wsp:val=""00A90BF0""/><wsp:rsid wsp:val=""00AA0311""/><wsp:rsid wsp:val=""00B01C92""/><wsp:rsid wsp:val=""00B13187""/><wsp:rsid wsp:val=""00BA7D2C""/>" & _
"<wsp:rsid wsp:val=""00BB3E24""/><wsp:rsid wsp:val=""00BB5802""/><wsp:rsid wsp:val=""00BE731D""/><wsp:rsid wsp:val=""00BF474D""/><wsp:rsid " & _
"wsp:val=""00C1673F""/><wsp:rsid wsp:val=""00C41BA3""/><wsp:rsid wsp:val=""00C95ED8""/><wsp:rsid wsp:val=""00D8143C""/><wsp:rsid wsp:val=""00DA4A7C""/>" & _
"<wsp:rsid wsp:val=""00E30BC0""/><wsp:rsid wsp:val=""00E3365A""/><wsp:rsid wsp:val=""00E76F19""/><wsp:rsid wsp:val=""00E83593""/><wsp:rsid " & _
"wsp:val=""00E86378""/><wsp:rsid wsp:val=""00E91440""/><wsp:rsid wsp:val=""00EA7C1C""/><wsp:rsid wsp:val=""00ED4953""/><wsp:rsid wsp:val=""00EF3DB7""/>" & _
"<wsp:rsid wsp:val=""00EF5BA9""/><wsp:rsid wsp:val=""00F37F28""/><wsp:rsid wsp:val=""00FF7CD5""/></wsp:rsids></w:docPr><w:body>"
End function

'標題抬頭
Function DocBody_1()
DocBody_1 = "<w:p wsp:rsidR=""008515F6"" wsp:rsidRDefault=""008515F6"" wsp:rsidP=""008515F6""><w:pPr><w:jc w:val=""center""/><w:rPr>" & _
"<w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體""/><wx:font wx:val=""新細明體""/><w:sz w:val=""44""/><w:sz-cs w:val=""44""/>" & _
"</w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體"" w:hint=""fareast""/><wx:font wx:val=""新細明體""/>" & _
"<w:sz w:val=""44""/><w:sz-cs w:val=""44""/></w:rPr><w:t>【發明專利申請書】</w:t></w:r></w:p><w:p wsp:rsidR=""008515F6"" wsp:rsidRDefault=""008515F6"" " & _
"wsp:rsidP=""008515F6""><w:pPr><w:rPr><w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體""/><wx:font wx:val=""新細明體""/>" & _
"</w:rPr></w:pPr></w:p>"


End function

'案由,一併申請實體審查,事務所或申請人案件編號
Function DocBody_2()
DocBody_2 = "<w:p wsp:rsidR=""008515F6"" wsp:rsidRPr=""00565151"" wsp:rsidRDefault=""008515F6"" wsp:rsidP=""00565151""><w:pPr>" & _
"<w:overflowPunct w:val=""off""/><w:autoSpaceDE w:val=""off""/><w:autoSpaceDN w:val=""off""/><w:spacing w:line=""360"" w:line-rule=""at-least""/>" & _
"</w:pPr><w:r><w:rPr><w:rFonts w:ascii=""Times New Roman"" w:h-ansi=""新細明體"" w:hint=""fareast""/><wx:font wx:val=""新細明體""/>" & _
"</w:rPr><w:t>【案由】　　　　　　　　　　　</w:t></w:r><w:r wsp:rsidR=""00565151""><w:rPr><w:rFonts w:ascii=""Times " & _
"New Roman"" w:h-ansi=""Times New Roman"" w:cs=""Times New Roman"" w:hint=""fareast""/><wx:font wx:val=""Times New Roman""/></w:rPr>" & _
"<w:t>#case_no#</w:t></w:r></w:p><w:p wsp:rsidR=""008515F6"" wsp:rsidRDefault=""008515F6"" wsp:rsidP=""008515F6""><w:pPr><w:rPr><w:rFonts " & _
"w:ascii=""新細明體"" w:h-ansi=""新細明體""/><wx:font wx:val=""新細明體""/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=""Times " & _
"New Roman"" w:h-ansi=""新細明體"" w:hint=""fareast""/><wx:font wx:val=""新細明體""/></w:rPr><w:t>【一併申請實體審查】　　　　　</w:t>" & _
"</w:r><w:r wsp:rsidR=""00565151""><w:rPr><w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體"" w:hint=""fareast""/><wx:font " & _
"wx:val=""新細明體""/></w:rPr><w:t>#reality#</w:t></w:r>" & _
"</w:p><w:p wsp:rsidR=""008515F6"" wsp:rsidRDefault=""008515F6"" wsp:rsidP=""008515F6""><w:pPr><w:rPr><w:rFonts w:ascii=""新細明體"" " & _
"w:h-ansi=""新細明體""/><wx:font wx:val=""新細明體""/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=""Times New Roman"" w:h-ansi=""新細明體"" " & _
"w:hint=""fareast""/><wx:font wx:val=""新細明體""/></w:rPr><w:t>【事務所或申請人案件編號】　　</w:t></w:r><w:r wsp:rsidR=""00565151"">" & _
"<w:rPr><w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體"" w:hint=""fareast""/><wx:font wx:val=""新細明體""/></w:rPr><w:t>#seq#</w:t>" & _
"</w:r></w:p>"
End function

'中文發明名稱,英文發明名稱
Function DocBody_3()
DocBody_3 = "<w:p wsp:rsidR=""008515F6"" wsp:rsidRDefault=""008515F6"" " & _
"wsp:rsidP=""008515F6""><w:pPr><w:rPr><w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體""/><wx:font wx:val=""新細明體""/>" & _
"</w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=""Times New Roman"" w:h-ansi=""新細明體"" w:hint=""fareast""/><wx:font wx:val=""新細明體""/>" & _
"</w:rPr><w:t>【中文發明名稱】　　　　　　　</w:t></w:r><w:r wsp:rsidR=""00565151""><w:rPr><w:rFonts w:ascii=""新細明體"" " & _
"w:h-ansi=""新細明體"" w:hint=""fareast""/><wx:font wx:val=""新細明體""/></w:rPr><w:t>#cappl_name#</w:t></w:r></w:p><w:p wsp:rsidR=""008515F6"" " & _
"wsp:rsidRDefault=""008515F6"" wsp:rsidP=""008515F6""><w:pPr><w:rPr><w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體""/><wx:font " & _
"wx:val=""新細明體""/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=""Times New Roman"" w:h-ansi=""新細明體"" w:hint=""fareast""/>" & _
"<wx:font wx:val=""新細明體""/></w:rPr><w:t>【英文發明名稱】　　　　　　　</w:t></w:r><w:r wsp:rsidR=""00565151"">" & _
"<w:rPr><w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體"" w:hint=""fareast""/><wx:font wx:val=""新細明體""/></w:rPr><w:t>#eappl_name#</w:t>" & _
"</w:r></w:p>"
End function

'申請人迴圈
Function Dmp_apcust_data()
Dmp_apcust_data = "<w:p wsp:rsidR=""008515F6"" wsp:rsidRDefault=""00565151"" " & _
"wsp:rsidP=""00565151""><w:pPr><w:pStyle w:val=""ac""/><w:ind w:left=""0""/><w:rPr><w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體""/>" & _
"<wx:font wx:val=""新細明體""/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=""Times New Roman"" w:h-ansi=""新細明體"" w:hint=""fareast""/>" & _
"<wx:font wx:val=""新細明體""/></w:rPr><w:t>【申請人#apply_num#】</w:t></w:r></w:p><w:p " & _
"wsp:rsidR=""00565151"" wsp:rsidRDefault=""008515F6"" wsp:rsidP=""00B55A67""><w:pPr><w:tabs><w:tab w:val=""left"" w:pos=""8028""/></w:tabs>" & _
"<w:rPr><w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體"" w:hint=""fareast""/><wx:font wx:val=""新細明體""/></w:rPr></w:pPr>" & _
"<w:r><w:rPr><w:rFonts w:ascii=""Times New Roman"" w:h-ansi=""新細明體"" w:hint=""fareast""/><wx:font wx:val=""新細明體""/></w:rPr>" & _
"<w:t>　　【國籍】　　　　　　　　　</w:t></w:r><w:r wsp:rsidR=""00565151""><w:rPr><w:rFonts w:ascii=""新細明體"" " & _
"w:h-ansi=""新細明體"" w:hint=""fareast""/><wx:font wx:val=""新細明體""/></w:rPr><w:t>#ap_country#</w:t></w:r></w:p><w:p wsp:rsidR=""008515F6"" " & _
"wsp:rsidRDefault=""008515F6"" wsp:rsidP=""008515F6""><w:pPr><w:rPr><w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體""/><wx:font " & _
"wx:val=""新細明體""/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=""Times New Roman"" w:h-ansi=""新細明體"" w:hint=""fareast""/>" & _
"<wx:font wx:val=""新細明體""/></w:rPr><w:t>　　【#ap_cname1_title#】　　　　　　　</w:t>" & _
"</w:r><w:r wsp:rsidR=""00565151""><w:rPr><w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體"" w:hint=""fareast""/><wx:font " & _
"wx:val=""新細明體""/></w:rPr><w:t>#ap_cname1#</w:t></w:r></w:p><w:p wsp:rsidR=""00565151"" " & _
"wsp:rsidRDefault=""00565151"" wsp:rsidP=""008515F6""><w:pPr><w:rPr><w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體"" w:hint=""fareast""/>" & _
"<wx:font wx:val=""新細明體""/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=""Times New Roman"" w:h-ansi=""新細明體"" w:hint=""fareast""/>" & _
"<wx:font wx:val=""新細明體""/></w:rPr><w:t>　　【#ap_ename1_title#】　　　　　　　#ap_ename1#</w:t></w:r></w:p>"
End function

'代理人1
Function Agt_data_1()
Agt_data_1 = "<w:p wsp:rsidR=""008515F6"" wsp:rsidRDefault=""00565151"" wsp:rsidP=""00565151""><w:pPr><w:pStyle w:val=""ac""/><w:ind w:left=""0""/>" & _
"<w:rPr><w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體""/><wx:font wx:val=""新細明體""/></w:rPr></w:pPr><w:r><w:rPr>" & _
"<w:rFonts w:ascii=""Times New Roman"" w:h-ansi=""新細明體"" w:hint=""fareast""/><wx:font wx:val=""新細明體""/></w:rPr><w:t>【代理人1】</w:t></w:r></w:p><w:p wsp:rsidR=""008515F6"" wsp:rsidRDefault=""008515F6"" wsp:rsidP=""008515F6"">" & _
"<w:pPr><w:rPr><w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體""/><wx:font wx:val=""新細明體""/></w:rPr></w:pPr><w:r>" & _
"<w:rPr><w:rFonts w:ascii=""Times New Roman"" w:h-ansi=""新細明體"" w:hint=""fareast""/><wx:font wx:val=""新細明體""/></w:rPr>" & _
"<w:t>　　【中文姓名】　　　　　　　</w:t></w:r><w:r wsp:rsidR=""00565151""><w:rPr><w:rFonts w:ascii=""新細明體"" " & _
"w:h-ansi=""新細明體"" w:hint=""fareast""/><wx:font wx:val=""新細明體""/></w:rPr><w:t>#agt_name1#</w:t></w:r></w:p>"
End function

'代理人2
Function Agt_data_2()
Agt_data_2 = "<w:p wsp:rsidR=""008515F6"" wsp:rsidRDefault=""00565151"" wsp:rsidP=""00565151""><w:pPr><w:pStyle w:val=""ac""/><w:ind w:left=""0""/>" & _
"<w:rPr><w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體""/><wx:font wx:val=""新細明體""/></w:rPr></w:pPr><w:r><w:rPr>" & _
"<w:rFonts w:ascii=""Times New Roman"" w:h-ansi=""新細明體"" w:hint=""fareast""/><wx:font wx:val=""新細明體""/></w:rPr><w:t>【代理人2】</w:t></w:r></w:p><w:p wsp:rsidR=""008515F6"" wsp:rsidRDefault=""008515F6"" wsp:rsidP=""008515F6"">" & _
"<w:pPr><w:rPr><w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體""/><wx:font wx:val=""新細明體""/></w:rPr></w:pPr><w:r>" & _
"<w:rPr><w:rFonts w:ascii=""Times New Roman"" w:h-ansi=""新細明體"" w:hint=""fareast""/><wx:font wx:val=""新細明體""/></w:rPr>" & _
"<w:t>　　【中文姓名】　　　　　　　</w:t></w:r><w:r wsp:rsidR=""00565151""><w:rPr><w:rFonts w:ascii=""新細明體"" " & _
"w:h-ansi=""新細明體"" w:hint=""fareast""/><wx:font wx:val=""新細明體""/></w:rPr><w:t>#agt_name2#</w:t></w:r></w:p>"
End function


'發明人迴圈
Function Ant_data()
Ant_data = "<w:p wsp:rsidR=""008515F6"" wsp:rsidRDefault=""00565151"" wsp:rsidP=""00565151""><w:pPr><w:pStyle w:val=""ac""/><w:ind w:left=""0""/>" & _
"<w:rPr><w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體""/><wx:font wx:val=""新細明體""/></w:rPr></w:pPr><w:r><w:rPr>" & _
"<w:rFonts w:ascii=""Times New Roman"" w:h-ansi=""新細明體"" w:hint=""fareast""/><wx:font wx:val=""新細明體""/></w:rPr><w:t>【#ant_num#】</w:t></w:r></w:p><w:p wsp:rsidR=""008515F6"" wsp:rsidRDefault=""008515F6"" " & _
"wsp:rsidP=""008515F6""><w:pPr><w:rPr><w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體""/><wx:font wx:val=""新細明體""/>" & _
"</w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=""Times New Roman"" w:h-ansi=""新細明體"" w:hint=""fareast""/><wx:font wx:val=""新細明體""/>" & _
"</w:rPr><w:t>　　【國籍】　　　　　　　　　</w:t></w:r><w:r wsp:rsidR=""00565151""><w:rPr><w:rFonts w:ascii=""新細明體"" " & _
"w:h-ansi=""新細明體"" w:hint=""fareast""/><wx:font wx:val=""新細明體""/></w:rPr><w:t>#ant_country#</w:t></w:r></w:p><w:p wsp:rsidR=""008515F6"" " & _
"wsp:rsidRDefault=""008515F6"" wsp:rsidP=""008515F6""><w:pPr><w:rPr><w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體""/><wx:font " & _
"wx:val=""新細明體""/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=""Times New Roman"" w:h-ansi=""新細明體"" w:hint=""fareast""/>" & _
"<wx:font wx:val=""新細明體""/></w:rPr><w:t>　　【中文姓名】　　　　　　　</w:t></w:r><w:r wsp:rsidR=""00565151"">" & _
"<w:rPr><w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體"" w:hint=""fareast""/><wx:font wx:val=""新細明體""/></w:rPr><w:t>#ant_cname#</w:t></w:r></w:p><w:p wsp:rsidR=""00565151"" wsp:rsidRPr=""00565151"" " & _
"wsp:rsidRDefault=""00565151"" wsp:rsidP=""00565151""><w:pPr><w:rPr><w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體""/><wx:font " & _
"wx:val=""新細明體""/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=""Times New Roman"" w:h-ansi=""新細明體"" w:hint=""fareast""/>" & _
"<wx:font wx:val=""新細明體""/></w:rPr><w:t>　　【英文姓名】　　　　　　　</w:t></w:r><w:r><w:rPr><w:rFonts w:ascii=""新細明體"" " & _
"w:h-ansi=""新細明體"" w:hint=""fareast""/><wx:font wx:val=""新細明體""/></w:rPr><w:t>#ant_ename#</w:t></w:r></w:p>"
End function
'主張優惠期
function DocBody_6()
DocBody_6 = "<w:p wsp:rsidR=""00C667B1"" wsp:rsidRDefault=""00C667B1""><w:pPr><w:pStyle w:val=""ac""/><w:listPr><w:ilvl w:val=""0""/><w:ilfo w:val=""8""/>" & _
"<wx:t wx:val=""【主張優惠期1】""/><wx:font wx:val=""新細明體""/></w:listPr><w:overflowPunct w:val=""off""/><w:autoSpaceDE " & _
"w:val=""off""/><w:autoSpaceDN w:val=""off""/><w:spacing w:line=""360"" w:line-rule=""at-least""/></w:pPr></w:p>" & _
"<w:p wsp:rsidR=""00C667B1"" " & _
"wsp:rsidRDefault=""005B309B""><w:pPr><w:overflowPunct w:val=""off""/><w:autoSpaceDE w:val=""off""/><w:autoSpaceDN w:val=""off""/><w:spacing " & _
"w:line=""360"" w:line-rule=""at-least""/></w:pPr><w:r><w:rPr><w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體"" w:hint=""fareast""/>" & _
"<wx:font wx:val=""新細明體""/></w:rPr><w:t>　　【發生日期】　　　　　　　#exh_date#</w:t></w:r></w:p>" & _
"<w:p wsp:rsidR=""00C667B1"" " & _
"wsp:rsidRDefault=""005B309B""><w:pPr><w:overflowPunct w:val=""off""/><w:autoSpaceDE w:val=""off""/><w:autoSpaceDN w:val=""off""/><w:spacing " & _
"w:line=""360"" w:line-rule=""at-least""/></w:pPr><w:r><w:rPr><w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體"" w:hint=""fareast""/>" & _
"<wx:font wx:val=""新細明體""/></w:rPr><w:t>　　【因實驗而公開者】　　　　</w:t></w:r></w:p>" & _
"<w:p wsp:rsidR=""00C667B1"" " & _
"wsp:rsidRDefault=""005B309B""><w:pPr><w:overflowPunct w:val=""off""/><w:autoSpaceDE w:val=""off""/><w:autoSpaceDN w:val=""off""/><w:spacing " & _
"w:line=""360"" w:line-rule=""at-least""/><w:rPr><w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體""/><wx:font wx:val=""新細明體""/>" & _
"</w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體"" w:hint=""fareast""/><wx:font wx:val=""新細明體""/>" & _
"</w:rPr><w:t>　　【因於刊物發表者】　　　　</w:t></w:r></w:p>" & _
"<w:p wsp:rsidR=""00C667B1"" wsp:rsidRDefault=""005B309B"">" & _
"<w:pPr><w:overflowPunct w:val=""off""/><w:autoSpaceDE w:val=""off""/><w:autoSpaceDN w:val=""off""/><w:spacing w:line=""360"" w:line-rule=""at-least""/>" & _
"</w:pPr><w:r><w:rPr><w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體"" w:hint=""fareast""/><wx:font wx:val=""新細明體""/>" & _
"</w:rPr><w:t>　　【因陳列於政府主辦或認可之展覽會者】</w:t></w:r></w:p>"
End function
'優惠期事實 20170504 智慧局取消主張優惠期改用優惠期事實
function DocBody_6_1()
DocBody_6_1 = "<w:p wsp:rsidR=""00C667B1"" " & _
"wsp:rsidRDefault=""005B309B""><w:pPr><w:overflowPunct w:val=""off""/><w:autoSpaceDE w:val=""off""/><w:autoSpaceDN w:val=""off""/><w:spacing " & _
"w:line=""360"" w:line-rule=""at-least""/></w:pPr><w:r><w:rPr><w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體"" w:hint=""fareast""/>" & _
"<wx:font wx:val=""新細明體""/></w:rPr><w:t>【本案符合優惠期相關規定】　　是／否</w:t></w:r></w:p>" & _
"<w:p wsp:rsidR=""00C667B1"" " & _
"wsp:rsidRDefault=""005B309B""><w:pPr><w:overflowPunct w:val=""off""/><w:autoSpaceDE w:val=""off""/><w:autoSpaceDN w:val=""off""/><w:spacing " & _
"w:line=""360"" w:line-rule=""at-least""/></w:pPr><w:r><w:rPr><w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體"" w:hint=""fareast""/>" & _
"<wx:font wx:val=""新細明體""/></w:rPr><w:t>【優惠期事實1】</w:t></w:r></w:p>" & _
"<w:p wsp:rsidR=""00C667B1"" " & _
"wsp:rsidRDefault=""005B309B""><w:pPr><w:overflowPunct w:val=""off""/><w:autoSpaceDE w:val=""off""/><w:autoSpaceDN w:val=""off""/><w:spacing " & _
"w:line=""360"" w:line-rule=""at-least""/></w:pPr><w:r><w:rPr><w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體"" w:hint=""fareast""/>" & _
"<wx:font wx:val=""新細明體""/></w:rPr><w:t>　　【發生日期】　　　　　　　#exh_date#</w:t></w:r></w:p>" & _
"<w:p wsp:rsidR=""00C667B1"" " & _
"wsp:rsidRDefault=""005B309B""><w:pPr><w:overflowPunct w:val=""off""/><w:autoSpaceDE w:val=""off""/><w:autoSpaceDN w:val=""off""/><w:spacing " & _
"w:line=""360"" w:line-rule=""at-least""/></w:pPr><w:r><w:rPr><w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體"" w:hint=""fareast""/>" & _
"<wx:font wx:val=""新細明體""/></w:rPr><w:t>　　【公開事由】　　　　　　　</w:t></w:r></w:p>"
End function
'主張優先權迴圈
Function DocBody_7()
DocBody_7 = "<w:p wsp:rsidR=""008515F6"" wsp:rsidRDefault=""00CC1819"" wsp:rsidP=""00CC1819""><w:pPr><w:pStyle w:val=""ac""/><w:ind w:left=""0""/>" & _
"<w:rPr><w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體""/><wx:font wx:val=""新細明體""/></w:rPr></w:pPr><w:r><w:rPr>" & _
"<w:rFonts w:ascii=""Times New Roman"" w:h-ansi=""新細明體"" w:hint=""fareast""/><wx:font wx:val=""新細明體""/></w:rPr><w:t>【主張優先權#prior_num#】</w:t></w:r></w:p>" & _
"<w:p wsp:rsidR=""008515F6"" wsp:rsidRDefault=""008515F6"" wsp:rsidP=""008515F6"">" & _
"<w:pPr><w:rPr><w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體""/><wx:font wx:val=""新細明體""/></w:rPr></w:pPr><w:r>" & _
"<w:rPr><w:rFonts w:ascii=""Times New Roman"" w:h-ansi=""新細明體"" w:hint=""fareast""/><wx:font wx:val=""新細明體""/></w:rPr>" & _
"<w:t>　　【申請日】　　　　　　　　</w:t></w:r><w:r wsp:rsidR=""00CC1819""><w:rPr><w:rFonts w:ascii=""新細明體"" " & _
"w:h-ansi=""新細明體"" w:hint=""fareast""/><wx:font wx:val=""新細明體""/></w:rPr><w:t>#prior_date#</w:t></w:r></w:p>" & _
"<w:p wsp:rsidR=""008515F6"" " & _
"wsp:rsidRDefault=""008515F6"" wsp:rsidP=""008515F6""><w:pPr><w:rPr><w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體""/><wx:font " & _
"wx:val=""新細明體""/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=""Times New Roman"" w:h-ansi=""新細明體"" w:hint=""fareast""/>" & _
"<wx:font wx:val=""新細明體""/></w:rPr><w:t>　　【受理國家或地區】　　　　</w:t></w:r><w:r wsp:rsidR=""00A67C95"">" & _
"<w:rPr><w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體"" w:hint=""fareast""/><wx:font wx:val=""新細明體""/></w:rPr><w:t>#prior_country#</w:t>" & _
"</w:r></w:p>" & _
"<w:p wsp:rsidR=""008515F6"" wsp:rsidRDefault=""008515F6"" wsp:rsidP=""008515F6""><w:pPr><w:rPr><w:rFonts w:ascii=""新細明體"" " & _
"w:h-ansi=""新細明體""/><wx:font wx:val=""新細明體""/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=""Times New Roman"" w:h-ansi=""新細明體"" " & _
"w:hint=""fareast""/><wx:font wx:val=""新細明體""/></w:rPr><w:t>　　【申請案號】　　　　　　　</w:t></w:r><w:r wsp:rsidR=""00A67C95"">" & _
"<w:rPr><w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體"" w:hint=""fareast""/><wx:font wx:val=""新細明體""/></w:rPr><w:t>#prior_no#</w:t>" & _
"</w:r></w:p>"
END function
'主張優先權迴圈-JA
Function DocBody_7_1()
DocBody_7_1 = "<w:p wsp:rsidR=""008515F6"" wsp:rsidRDefault=""008515F6"" wsp:rsidP=""008515F6""><w:pPr><w:overflowPunct w:val=""off""/>" & _
"<w:autoSpaceDE w:val=""off""/><w:autoSpaceDN w:val=""off""/><w:spacing w:line=""360"" w:line-rule=""at-least""/></w:pPr><w:r><w:rPr>" & _
"<w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體"" w:hint=""fareast""/><wx:font wx:val=""新細明體""/></w:rPr><w:t>　　【專利類別】　　　　　　　</w:t>" & _
"</w:r><w:r wsp:rsidR=""00A67C95""><w:rPr><w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體"" w:hint=""fareast""/><wx:font " & _
"wx:val=""新細明體""/></w:rPr><w:t>#case1nm#</w:t></w:r></w:p>" & _
"<w:p wsp:rsidR=""008515F6"" wsp:rsidRDefault=""008515F6"" wsp:rsidP=""008515F6"">" & _
"<w:pPr><w:overflowPunct w:val=""off""/><w:autoSpaceDE w:val=""off""/><w:autoSpaceDN w:val=""off""/><w:spacing w:line=""360"" w:line-rule=""at-least""/>" & _
"<w:rPr><w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體""/><wx:font wx:val=""新細明體""/></w:rPr></w:pPr><w:r><w:rPr>" & _
"<w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體"" w:hint=""fareast""/><wx:font wx:val=""新細明體""/></w:rPr><w:t>　　【存取碼】　　　　　　　　</w:t>" & _
"</w:r><w:r wsp:rsidR=""00A67C95""><w:rPr><w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體"" w:hint=""fareast""/><wx:font " & _
"wx:val=""新細明體""/></w:rPr><w:t>#mprior_access#</w:t></w:r></w:p>"
end function
'主張優先權迴圈-KO
Function DocBody_7_2()
DocBody_7_2 = "<w:p wsp:rsidR=""008515F6"" wsp:rsidRDefault=""008515F6"" wsp:rsidP=""008515F6"">" & _
"<w:pPr><w:overflowPunct w:val=""off""/><w:autoSpaceDE w:val=""off""/><w:autoSpaceDN w:val=""off""/><w:spacing w:line=""360"" w:line-rule=""at-least""/>" & _
"<w:rPr><w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體""/><wx:font wx:val=""新細明體""/></w:rPr></w:pPr><w:r><w:rPr>" & _
"<w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體"" w:hint=""fareast""/><wx:font wx:val=""新細明體""/></w:rPr><w:t>　　【存取碼】　　　　　　　　</w:t>" & _
"</w:r><w:r wsp:rsidR=""00A67C95""><w:rPr><w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體"" w:hint=""fareast""/><wx:font " & _
"wx:val=""新細明體""/></w:rPr><w:t>#mprior_access#</w:t></w:r></w:p>"
end function
'主張利用生物材料
Function DocBody_7_2()
DocBody_7_2 = "<w:p wsp:rsidR=""00C667B1"" wsp:rsidRDefault=""00C667B1"">" & _
"<w:pPr><w:pStyle w:val=""ac""/><w:listPr><w:ilvl w:val=""0""/><w:ilfo w:val=""12""/><wx:t wx:val=""【主張利用生物材料1】""/>" & _
"<wx:font wx:val=""新細明體""/></w:listPr><w:overflowPunct w:val=""off""/><w:autoSpaceDE w:val=""off""/><w:autoSpaceDN w:val=""off""/>" & _
"<w:spacing w:line=""360"" w:line-rule=""at-least""/></w:pPr></w:p>" & _
"<w:p wsp:rsidR=""00C667B1"" wsp:rsidRDefault=""005B309B""><w:pPr>" & _
"<w:overflowPunct w:val=""off""/><w:autoSpaceDE w:val=""off""/><w:autoSpaceDN w:val=""off""/><w:spacing w:line=""360"" w:line-rule=""at-least""/>" & _
"</w:pPr><w:r><w:rPr><w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體"" w:hint=""fareast""/><wx:font wx:val=""新細明體""/>" & _
"</w:rPr><w:t>　　【寄存國家】　　　　　　　</w:t></w:r></w:p>" & _
"<w:p wsp:rsidR=""00C667B1"" wsp:rsidRDefault=""005B309B"">" & _
"<w:pPr><w:overflowPunct w:val=""off""/><w:autoSpaceDE w:val=""off""/><w:autoSpaceDN w:val=""off""/><w:spacing w:line=""360"" w:line-rule=""at-least""/>" & _
"</w:pPr><w:r><w:rPr><w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體"" w:hint=""fareast""/><wx:font wx:val=""新細明體""/>" & _
"</w:rPr><w:t>　　【寄存機構】　　　　　　　</w:t></w:r></w:p>" & _
"<w:p wsp:rsidR=""00C667B1"" wsp:rsidRDefault=""005B309B"">" & _
"<w:pPr><w:overflowPunct w:val=""off""/><w:autoSpaceDE w:val=""off""/><w:autoSpaceDN w:val=""off""/><w:spacing w:line=""360"" w:line-rule=""at-least""/>" & _
"</w:pPr><w:r><w:rPr><w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體"" w:hint=""fareast""/><wx:font wx:val=""新細明體""/>" & _
"</w:rPr><w:t>　　【寄存日期】　　　　　　　</w:t></w:r></w:p>" & _
"<w:p wsp:rsidR=""00C667B1"" wsp:rsidRDefault=""005B309B"">" & _
"<w:pPr><w:overflowPunct w:val=""off""/><w:autoSpaceDE w:val=""off""/><w:autoSpaceDN w:val=""off""/><w:spacing w:line=""360"" w:line-rule=""at-least""/>" & _
"</w:pPr><w:r><w:rPr><w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體"" w:hint=""fareast""/><wx:font wx:val=""新細明體""/>" & _
"</w:rPr><w:t>　　【寄存號碼】　　　　　　　</w:t></w:r></w:p>" 
end function
'生物材料不須寄存
Function DocBody_8()
DocBody_8 = "<w:p wsp:rsidR=""008515F6"" wsp:rsidRDefault=""008515F6"" wsp:rsidP=""008515F6""><w:pPr><w:rPr><w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體""/>" & _
"<wx:font wx:val=""新細明體""/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=""Times New Roman"" w:h-ansi=""新細明體"" w:hint=""fareast""/>" & _
"<wx:font wx:val=""新細明體""/></w:rPr><w:t>【生物材料不須寄存】　　　所屬技術領域中具有通常知識者易於獲得。</w:t>" & _
"</w:r></w:p>"
end function

'聲明本人就相同創作在申請本發明專利之同日-另申請新型專利
Function DocBody_81()
DocBody_81 = "<w:p wsp:rsidR=""008515F6"" wsp:rsidRDefault=""008515F6"" wsp:rsidP=""008515F6""><w:pPr><w:rPr><w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體""/>" & _
"<wx:font wx:val=""新細明體""/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=""Times New Roman"" w:h-ansi=""新細明體"" w:hint=""fareast""/>" & _
"<wx:font wx:val=""新細明體""/></w:rPr><w:t>【聲明本人就相同創作在申請本發明專利之同日-另申請新型專利】　#same_apply#</w:t>" & _
"</w:r></w:p>"
end function

'中文本資訊 外文本資訊 繳費資訊
Function DocBody_9()
DocBody_9 = "<w:p wsp:rsidR=""008515F6"" wsp:rsidRDefault=""008515F6"" " & _
"wsp:rsidP=""008515F6""><w:pPr><w:rPr><w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體""/><wx:font wx:val=""新細明體""/>" & _
"</w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=""Times New Roman"" w:h-ansi=""新細明體"" w:hint=""fareast""/><wx:font wx:val=""新細明體""/>" & _
"</w:rPr><w:t>【中文本資訊】</w:t></w:r></w:p>" & _
"<w:p wsp:rsidR=""008515F6"" wsp:rsidRDefault=""008515F6"" wsp:rsidP=""008515F6"">" & _
"<w:pPr><w:overflowPunct w:val=""off""/><w:autoSpaceDE w:val=""off""/><w:autoSpaceDN w:val=""off""/><w:rPr><w:rFonts w:ascii=""新細明體"" " & _
"w:h-ansi=""新細明體""/><wx:font wx:val=""新細明體""/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體"" " & _
"w:hint=""fareast""/><wx:font wx:val=""新細明體""/></w:rPr><w:t>　　【摘要頁數】　　　　　　　</w:t></w:r><w:r wsp:rsidR=""00B55A67"">" & _
"<w:rPr><w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體"" w:hint=""fareast""/><wx:font wx:val=""新細明體""/></w:rPr><w:t>0</w:t>" & _
"</w:r></w:p>" & _
"<w:p wsp:rsidR=""008515F6"" wsp:rsidRDefault=""008515F6"" wsp:rsidP=""008515F6""><w:pPr><w:overflowPunct w:val=""off""/>" & _
"<w:autoSpaceDE w:val=""off""/><w:autoSpaceDN w:val=""off""/><w:rPr><w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體""/><wx:font " & _
"wx:val=""新細明體""/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體"" w:hint=""fareast""/>" & _
"<wx:font wx:val=""新細明體""/></w:rPr><w:t>　　【說明書頁數】　　　　　　0</w:t></w:r></w:p>" & _
"<w:p wsp:rsidR=""008515F6"" " & _
"wsp:rsidRDefault=""008515F6"" wsp:rsidP=""008515F6""><w:pPr><w:overflowPunct w:val=""off""/><w:autoSpaceDE w:val=""off""/><w:autoSpaceDN " & _
"w:val=""off""/><w:rPr><w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體""/><wx:font wx:val=""新細明體""/></w:rPr></w:pPr>" & _
"<w:r><w:rPr><w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體"" w:hint=""fareast""/><wx:font wx:val=""新細明體""/></w:rPr>" & _
"<w:t>　　【申請專利範圍頁數】　　　</w:t></w:r><w:r wsp:rsidR=""00B55A67""><w:rPr><w:rFonts w:ascii=""新細明體"" " & _
"w:h-ansi=""新細明體"" w:hint=""fareast""/><wx:font wx:val=""新細明體""/></w:rPr><w:t>0</w:t></w:r></w:p>" & _
"<w:p wsp:rsidR=""008515F6"" " & _
"wsp:rsidRDefault=""008515F6"" wsp:rsidP=""008515F6""><w:pPr><w:overflowPunct w:val=""off""/><w:autoSpaceDE w:val=""off""/><w:autoSpaceDN " & _
"w:val=""off""/><w:rPr><w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體""/><wx:font wx:val=""新細明體""/></w:rPr></w:pPr>" & _
"<w:r><w:rPr><w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體"" w:hint=""fareast""/><wx:font wx:val=""新細明體""/></w:rPr>" & _
"<w:t>　　【圖式頁數】　　　　　　　</w:t></w:r><w:r wsp:rsidR=""00B55A67""><w:rPr><w:rFonts w:ascii=""新細明體"" " & _
"w:h-ansi=""新細明體"" w:hint=""fareast""/><wx:font wx:val=""新細明體""/></w:rPr><w:t>0</w:t></w:r></w:p>" & _
"<w:p wsp:rsidR=""008515F6"" " & _
"wsp:rsidRDefault=""008515F6"" wsp:rsidP=""008515F6""><w:pPr><w:overflowPunct w:val=""off""/><w:autoSpaceDE w:val=""off""/><w:autoSpaceDN " & _
"w:val=""off""/><w:rPr><w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體""/><wx:font wx:val=""新細明體""/></w:rPr></w:pPr>" & _
"<w:r><w:rPr><w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體"" w:hint=""fareast""/><wx:font wx:val=""新細明體""/></w:rPr>" & _
"<w:t>　　【頁數總計】　　　　　　　</w:t></w:r><w:r wsp:rsidR=""00B55A67""><w:rPr><w:rFonts w:ascii=""新細明體"" " & _
"w:h-ansi=""新細明體"" w:hint=""fareast""/><wx:font wx:val=""新細明體""/></w:rPr><w:t>0</w:t></w:r></w:p>" & _
"<w:p wsp:rsidR=""008515F6"" " & _
"wsp:rsidRDefault=""008515F6"" wsp:rsidP=""008515F6""><w:pPr><w:overflowPunct w:val=""off""/><w:autoSpaceDE w:val=""off""/><w:autoSpaceDN " & _
"w:val=""off""/><w:rPr><w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體""/><wx:font wx:val=""新細明體""/></w:rPr></w:pPr>" & _
"<w:r><w:rPr><w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體"" w:hint=""fareast""/><wx:font wx:val=""新細明體""/></w:rPr>" & _
"<w:t>　　【申請專利範圍項數】　　　0</w:t></w:r></w:p>" & _
"<w:p wsp:rsidR=""008515F6"" wsp:rsidRDefault=""008515F6"" wsp:rsidP=""008515F6"">" & _
"<w:pPr><w:overflowPunct w:val=""off""/><w:autoSpaceDE w:val=""off""/><w:autoSpaceDN w:val=""off""/><w:rPr><w:rFonts w:ascii=""新細明體"" " & _
"w:h-ansi=""新細明體""/><wx:font wx:val=""新細明體""/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體"" " & _
"w:hint=""fareast""/><wx:font wx:val=""新細明體""/></w:rPr><w:t>　　【圖式圖數】　　　　　　　</w:t></w:r><w:r wsp:rsidR=""00B55A67"">" & _
"<w:rPr><w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體"" w:hint=""fareast""/><wx:font wx:val=""新細明體""/></w:rPr><w:t>0</w:t>" & _
"</w:r></w:p>" & _
"<w:p wsp:rsidR=""008515F6"" wsp:rsidRDefault=""008515F6"" wsp:rsidP=""008515F6"">" & _
"<w:pPr><w:overflowPunct w:val=""off""/><w:autoSpaceDE w:val=""off""/><w:autoSpaceDN w:val=""off""/><w:rPr><w:rFonts w:ascii=""新細明體"" " & _
"w:h-ansi=""新細明體""/><wx:font wx:val=""新細明體""/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體"" " & _
"w:hint=""fareast""/><wx:font wx:val=""新細明體""/></w:rPr><w:t>　　【附英文摘要】　　　　　　　</w:t></w:r><w:r wsp:rsidR=""00B55A67"">" & _
"<w:rPr><w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體"" w:hint=""fareast""/><wx:font wx:val=""新細明體""/></w:rPr><w:t></w:t>" & _
"</w:r></w:p>" & _

"<w:p wsp:rsidR=""008515F6"" wsp:rsidRDefault=""008515F6"" wsp:rsidP=""008515F6""><w:pPr><w:rPr><w:rFonts w:ascii=""新細明體"" " & _
"w:h-ansi=""新細明體""/><wx:font wx:val=""新細明體""/></w:rPr></w:pPr></w:p>" & _

"<w:p wsp:rsidR=""008515F6"" wsp:rsidRDefault=""008515F6"" " & _
"wsp:rsidP=""008515F6""><w:pPr><w:rPr><w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體""/><wx:font wx:val=""新細明體""/>" & _
"</w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=""Times New Roman"" w:h-ansi=""新細明體"" w:hint=""fareast""/><wx:font wx:val=""新細明體""/>" & _
"</w:rPr><w:t>【外文本資訊】</w:t></w:r></w:p>" & _
"<w:p wsp:rsidR=""008515F6"" wsp:rsidRDefault=""008515F6"" wsp:rsidP=""008515F6"">" & _
"<w:pPr><w:rPr><w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體""/><wx:font wx:val=""新細明體""/></w:rPr></w:pPr><w:r>" & _
"<w:rPr><w:rFonts w:ascii=""Times New Roman"" w:h-ansi=""新細明體"" w:hint=""fareast""/><wx:font wx:val=""新細明體""/></w:rPr>" & _
"<w:t>　　【外文頁數總計】　　　　　</w:t></w:r><w:r><w:rPr><w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體"" " & _
"w:hint=""fareast""/><wx:font wx:val=""新細明體""/></w:rPr><w:t>0</w:t></w:r></w:p>" & _
"<w:p wsp:rsidR=""008515F6"" wsp:rsidRDefault=""008515F6"" " & _
"wsp:rsidP=""008515F6""><w:pPr><w:rPr><w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體""/><wx:font wx:val=""新細明體""/>" & _
"</w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=""Times New Roman"" w:h-ansi=""新細明體"" w:hint=""fareast""/><wx:font wx:val=""新細明體""/>" & _
"</w:rPr><w:t>　　【外文本種類】　　　　　　日文／英文／德文／韓文／法文／俄文／葡萄牙文／西班牙文／阿拉伯文</w:t></w:r></w:p>" & _
"<w:p wsp:rsidR=""008515F6"" wsp:rsidRDefault=""008515F6"" " & _
"wsp:rsidP=""008515F6""><w:pPr><w:rPr><w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體""/><wx:font wx:val=""新細明體""/>" & _
"</w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體"" w:hint=""fareast""/><wx:font wx:val=""新細明體""/>" & _
"</w:rPr><w:t></w:t></w:r></w:p>" & _
"<w:p wsp:rsidR=""008515F6"" wsp:rsidRDefault=""008515F6"" " & _
"wsp:rsidP=""008515F6""><w:pPr><w:rPr><w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體""/><wx:font wx:val=""新細明體""/>" & _
"</w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=""Times New Roman"" w:h-ansi=""新細明體"" w:hint=""fareast""/><wx:font wx:val=""新細明體""/>" & _
"</w:rPr><w:t>【簡體字本資訊】</w:t></w:r></w:p>" & _
"<w:p wsp:rsidR=""008515F6"" wsp:rsidRDefault=""008515F6"" wsp:rsidP=""008515F6"">" & _
"<w:pPr><w:rPr><w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體""/><wx:font wx:val=""新細明體""/></w:rPr></w:pPr><w:r>" & _
"<w:rPr><w:rFonts w:ascii=""Times New Roman"" w:h-ansi=""新細明體"" w:hint=""fareast""/><wx:font wx:val=""新細明體""/></w:rPr>" & _
"<w:t>　　【簡體字頁數總計】　　　　</w:t></w:r><w:r><w:rPr><w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體"" " & _
"w:hint=""fareast""/><wx:font wx:val=""新細明體""/></w:rPr><w:t>0</w:t></w:r></w:p>" & _
"<w:p wsp:rsidR=""008515F6"" wsp:rsidRDefault=""008515F6"" " & _
"wsp:rsidP=""008515F6""><w:pPr><w:rPr><w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體""/><wx:font wx:val=""新細明體""/>" & _
"</w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體"" w:hint=""fareast""/><wx:font wx:val=""新細明體""/>" & _
"</w:rPr><w:t></w:t></w:r></w:p>" & _
"<w:p wsp:rsidR=""008515F6"" wsp:rsidRDefault=""008515F6"" wsp:rsidP=""008515F6""><w:pPr><w:rPr><w:rFonts " & _
"w:ascii=""新細明體"" w:h-ansi=""新細明體""/><wx:font wx:val=""新細明體""/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=""Times " & _
"New Roman"" w:h-ansi=""新細明體"" w:hint=""fareast""/><wx:font wx:val=""新細明體""/></w:rPr><w:t>【繳費資訊】</w:t></w:r>" & _
"</w:p>" & _
"<w:p wsp:rsidR=""008515F6"" wsp:rsidRDefault=""008515F6"" wsp:rsidP=""008515F6""><w:pPr><w:rPr><w:rFonts w:ascii=""新細明體"" " & _
"w:h-ansi=""新細明體""/><wx:font wx:val=""新細明體""/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=""Times New Roman"" w:h-ansi=""新細明體"" " & _
"w:hint=""fareast""/><wx:font wx:val=""新細明體""/></w:rPr><w:t>　　【繳費金額】　　　　　　　</w:t></w:r><w:r wsp:rsidR=""00B55A67"">" & _
"<w:rPr><w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體"" w:hint=""fareast""/><wx:font wx:val=""新細明體""/></w:rPr><w:t>0</w:t>" & _
"</w:r></w:p>"
End function

'20170524 增加收據抬頭選項
Function Dmp_receipt_title()
Dmp_receipt_title = _
"<w:p wsp:rsidR=""008515F6"" wsp:rsidRDefault=""008515F6"" wsp:rsidP=""008515F6""><w:pPr><w:rPr><w:rFonts w:ascii=""新細明體"" " & _
"w:h-ansi=""新細明體""/><wx:font wx:val=""新細明體""/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=""Times New Roman"" w:h-ansi=""新細明體"" " & _
"w:hint=""fareast""/><wx:font wx:val=""新細明體""/></w:rPr><w:t>　　【收據抬頭】　　　　　　　</w:t></w:r><w:r wsp:rsidR=""00B55A67"">" & _
"<w:rPr><w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體"" w:hint=""fareast""/><wx:font wx:val=""新細明體""/></w:rPr><w:t>#rectitle_name#</w:t>" & _
"</w:r></w:p>"
End function

'附送書件&備註
Function DocBody_10()
DocBody_10 = "<w:p wsp:rsidR=""008515F6"" wsp:rsidRDefault=""008515F6"" " & _
"wsp:rsidP=""008515F6""><w:pPr><w:rPr><w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體""/><wx:font wx:val=""新細明體""/>" & _
"</w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=""Times New Roman"" w:h-ansi=""新細明體"" w:hint=""fareast""/><wx:font wx:val=""新細明體""/>" & _
"</w:rPr><w:t>【備註】　　　　　　　　　　　</w:t></w:r></w:p>" & _
SpaceString() & _
"<w:p wsp:rsidR=""008515F6"" wsp:rsidRDefault=""008515F6"" " & _
"wsp:rsidP=""008515F6""><w:pPr><w:rPr><w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體""/><wx:font wx:val=""新細明體""/>" & _
"</w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=""Times New Roman"" w:h-ansi=""新細明體"" w:hint=""fareast""/><wx:font wx:val=""新細明體""/>" & _
"</w:rPr><w:t>【附送書件】</w:t></w:r></w:p>" & _
"<w:p wsp:rsidR=""008515F6"" wsp:rsidRDefault=""008515F6"" wsp:rsidP=""008515F6""><w:pPr>" & _
"<w:rPr><w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體""/><wx:font wx:val=""新細明體""/></w:rPr></w:pPr><w:r><w:rPr>" & _
"<w:rFonts w:ascii=""Times New Roman"" w:h-ansi=""新細明體"" w:hint=""fareast""/><wx:font wx:val=""新細明體""/></w:rPr>" & _
"<w:t>　　【基本資料表】　　　　　　</w:t>" & _
"</w:r><w:r><w:rPr><w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體"" w:hint=""fareast""/><wx:font wx:val=""新細明體""/></w:rPr>" & _
"<w:t>#seq#-Contact.pdf</w:t></w:r></w:p>" & _
"<w:p wsp:rsidR=""008515F6"" wsp:rsidRDefault=""008515F6"" " & _
"wsp:rsidP=""008515F6""><w:pPr><w:rPr><w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體""/><wx:font wx:val=""新細明體""/>" & _
"</w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=""Times New Roman"" w:h-ansi=""新細明體"" w:hint=""fareast""/><wx:font wx:val=""新細明體""/>" & _
"</w:rPr>" & _
"<w:t>　　【發明摘要】　　　　　　　</w:t></w:r><w:r><w:rPr><w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體"" " & _
"w:hint=""fareast""/><wx:font wx:val=""新細明體""/></w:rPr>" & _
"<w:t>#seq#-desc_Abstract.pdf</w:t></w:r></w:p>" & _
"<w:p wsp:rsidR=""008515F6"" wsp:rsidRDefault=""008515F6"" wsp:rsidP=""008515F6""><w:pPr><w:rPr><w:rFonts " & _
"w:ascii=""新細明體"" w:h-ansi=""新細明體""/><wx:font wx:val=""新細明體""/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=""Times " & _
"New Roman"" w:h-ansi=""新細明體"" w:hint=""fareast""/><wx:font wx:val=""新細明體""/></w:rPr>" & _
"<w:t>　　【發明說明書】　　　　　　</w:t>" & _
"</w:r><w:r><w:rPr><w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體"" w:hint=""fareast""/><wx:font wx:val=""新細明體""/>" & _
"</w:rPr><w:t>#seq#-desc_Description.pdf</w:t></w:r></w:p>" & _
"<w:p wsp:rsidR=""008515F6"" " & _
"wsp:rsidRDefault=""008515F6"" wsp:rsidP=""008515F6""><w:pPr><w:rPr><w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體""/><wx:font " & _
"wx:val=""新細明體""/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=""Times New Roman"" w:h-ansi=""新細明體"" w:hint=""fareast""/>" & _
"<wx:font wx:val=""新細明體""/></w:rPr><w:t>　　【序列表】　　　　　　　　</w:t></w:r><w:r><w:rPr><w:rFonts w:ascii=""新細明體"" " & _
"w:h-ansi=""新細明體"" w:hint=""fareast""/><wx:font wx:val=""新細明體""/></w:rPr><w:t>#seq#-Squence.pdf</w:t></w:r></w:p>" & _

"<w:p wsp:rsidR=""008515F6"" " & _
"wsp:rsidRDefault=""008515F6"" wsp:rsidP=""008515F6""><w:pPr><w:rPr><w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體""/><wx:font " & _
"wx:val=""新細明體""/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=""Times New Roman"" w:h-ansi=""新細明體"" w:hint=""fareast""/>" & _
"<wx:font wx:val=""新細明體""/></w:rPr><w:t>　　【發明申請專利範圍】　　　</w:t></w:r><w:r><w:rPr><w:rFonts w:ascii=""新細明體"" " & _
"w:h-ansi=""新細明體"" w:hint=""fareast""/><wx:font wx:val=""新細明體""/></w:rPr><w:t>#seq#-desc_Claims.pdf</w:t></w:r></w:p>" & _
"<w:p wsp:rsidR=""008515F6"" wsp:rsidRDefault=""008515F6"" wsp:rsidP=""008515F6""><w:pPr><w:rPr><w:rFonts " & _
"w:ascii=""新細明體"" w:h-ansi=""新細明體""/><wx:font wx:val=""新細明體""/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=""Times " & _
"New Roman"" w:h-ansi=""新細明體"" w:hint=""fareast""/><wx:font wx:val=""新細明體""/></w:rPr><w:t>　　【發明圖式】　　　　　　　</w:t>" & _
"</w:r><w:r><w:rPr><w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體"" w:hint=""fareast""/><wx:font wx:val=""新細明體""/>" & _
"</w:rPr><w:t>#seq#-Drawings.pdf</w:t></w:r></w:p>" & _
"<w:p wsp:rsidR=""008515F6"" wsp:rsidRDefault=""008515F6"" " & _
"wsp:rsidP=""008515F6""><w:pPr><w:rPr><w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體""/><wx:font wx:val=""新細明體""/>" & _
"</w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=""Times New Roman"" w:h-ansi=""新細明體"" w:hint=""fareast""/><wx:font wx:val=""新細明體""/>" & _
"</w:rPr><w:t>　　【外文本】　　　　　　　　</w:t></w:r><w:r><w:rPr><w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體"" " & _
"w:hint=""fareast""/><wx:font wx:val=""新細明體""/></w:rPr><w:t>#seq#-ForeignAbstract.pdf</w:t></w:r></w:p>" & _
"<w:p wsp:rsidR=""008515F6"" wsp:rsidRDefault=""008515F6"" wsp:rsidP=""008515F6""><w:pPr><w:rPr><w:rFonts " & _
"w:ascii=""新細明體"" w:h-ansi=""新細明體""/><wx:font wx:val=""新細明體""/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=""Times " & _
"New Roman"" w:h-ansi=""新細明體"" w:hint=""fareast""/><wx:font wx:val=""新細明體""/></w:rPr><w:t>　　【外文本】　　　　　　　　</w:t>" & _
"</w:r><w:r><w:rPr><w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體"" w:hint=""fareast""/><wx:font wx:val=""新細明體""/>" & _
"</w:rPr><w:t>#seq#-ForeignDescription.pdf</w:t></w:r></w:p>" & _
"<w:p wsp:rsidR=""008515F6"" " & _
"wsp:rsidRDefault=""008515F6"" wsp:rsidP=""008515F6""><w:pPr><w:rPr><w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體""/><wx:font " & _
"wx:val=""新細明體""/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=""Times New Roman"" w:h-ansi=""新細明體"" w:hint=""fareast""/>" & _
"<wx:font wx:val=""新細明體""/></w:rPr><w:t>　　【外文本】　　　　　　　　</w:t></w:r><w:r><w:rPr><w:rFonts w:ascii=""新細明體"" " & _
"w:h-ansi=""新細明體"" w:hint=""fareast""/><wx:font wx:val=""新細明體""/></w:rPr><w:t>#seq#-ForeignClaims.pdf</w:t></w:r></w:p>" & _
"<w:p wsp:rsidR=""008515F6"" wsp:rsidRDefault=""008515F6"" wsp:rsidP=""008515F6""><w:pPr><w:rPr><w:rFonts " & _
"w:ascii=""新細明體"" w:h-ansi=""新細明體""/><wx:font wx:val=""新細明體""/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=""Times " & _
"New Roman"" w:h-ansi=""新細明體"" w:hint=""fareast""/><wx:font wx:val=""新細明體""/></w:rPr><w:t>　　【外文本】　　　　　　　　</w:t>" & _
"</w:r><w:r><w:rPr><w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體"" w:hint=""fareast""/><wx:font wx:val=""新細明體""/>" & _
"</w:rPr><w:t>#seq#-ForeignDrawings.pdf</w:t></w:r></w:p>" & _
"<w:p wsp:rsidR=""008515F6"" wsp:rsidRDefault=""008515F6"" wsp:rsidP=""008515F6""><w:pPr><w:rPr><w:rFonts " & _
"w:ascii=""新細明體"" w:h-ansi=""新細明體""/><wx:font wx:val=""新細明體""/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=""Times " & _
"New Roman"" w:h-ansi=""新細明體"" w:hint=""fareast""/><wx:font wx:val=""新細明體""/></w:rPr><w:t>　　【外文本】　　　　　　　　</w:t>" & _
"</w:r><w:r><w:rPr><w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體"" w:hint=""fareast""/><wx:font wx:val=""新細明體""/>" & _
"</w:rPr><w:t>#seq#-ForeignSpec.pdf</w:t></w:r></w:p>" & _
"<w:p wsp:rsidR=""008515F6"" wsp:rsidRDefault=""008515F6"" wsp:rsidP=""008515F6""><w:pPr><w:rPr><w:rFonts " & _
"w:ascii=""新細明體"" w:h-ansi=""新細明體""/><wx:font wx:val=""新細明體""/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=""Times " & _
"New Roman"" w:h-ansi=""新細明體"" w:hint=""fareast""/><wx:font wx:val=""新細明體""/></w:rPr><w:t>　　【簡體字本】　　　　　　　</w:t>" & _
"</w:r><w:r><w:rPr><w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體"" w:hint=""fareast""/><wx:font wx:val=""新細明體""/>" & _
"</w:rPr><w:t>#seq#-SimplifiedAbstract.pdf</w:t></w:r></w:p>" & _
"<w:p wsp:rsidR=""008515F6"" wsp:rsidRDefault=""008515F6"" wsp:rsidP=""008515F6""><w:pPr><w:rPr><w:rFonts " & _
"w:ascii=""新細明體"" w:h-ansi=""新細明體""/><wx:font wx:val=""新細明體""/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=""Times " & _
"New Roman"" w:h-ansi=""新細明體"" w:hint=""fareast""/><wx:font wx:val=""新細明體""/></w:rPr><w:t>　　【簡體字本】　　　　　　　</w:t>" & _
"</w:r><w:r><w:rPr><w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體"" w:hint=""fareast""/><wx:font wx:val=""新細明體""/>" & _
"</w:rPr><w:t>#seq#-SimplifiedDescription.pdf</w:t></w:r></w:p>" & _
"<w:p wsp:rsidR=""008515F6"" wsp:rsidRDefault=""008515F6"" wsp:rsidP=""008515F6""><w:pPr><w:rPr><w:rFonts " & _
"w:ascii=""新細明體"" w:h-ansi=""新細明體""/><wx:font wx:val=""新細明體""/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=""Times " & _
"New Roman"" w:h-ansi=""新細明體"" w:hint=""fareast""/><wx:font wx:val=""新細明體""/></w:rPr><w:t>　　【簡體字本】　　　　　　　</w:t>" & _
"</w:r><w:r><w:rPr><w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體"" w:hint=""fareast""/><wx:font wx:val=""新細明體""/>" & _
"</w:rPr><w:t>#seq#-SimplifiedClaims.pdf</w:t></w:r></w:p>" & _
"<w:p wsp:rsidR=""008515F6"" wsp:rsidRDefault=""008515F6"" wsp:rsidP=""008515F6""><w:pPr><w:rPr><w:rFonts " & _
"w:ascii=""新細明體"" w:h-ansi=""新細明體""/><wx:font wx:val=""新細明體""/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=""Times " & _
"New Roman"" w:h-ansi=""新細明體"" w:hint=""fareast""/><wx:font wx:val=""新細明體""/></w:rPr><w:t>　　【簡體字本】　　　　　　　</w:t>" & _
"</w:r><w:r><w:rPr><w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體"" w:hint=""fareast""/><wx:font wx:val=""新細明體""/>" & _
"</w:rPr><w:t>#seq#-SimplifiedDrawings.pdf</w:t></w:r></w:p>" & _
"<w:p wsp:rsidR=""008515F6"" wsp:rsidRDefault=""008515F6"" wsp:rsidP=""008515F6""><w:pPr><w:rPr><w:rFonts " & _
"w:ascii=""新細明體"" w:h-ansi=""新細明體""/><wx:font wx:val=""新細明體""/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=""Times " & _
"New Roman"" w:h-ansi=""新細明體"" w:hint=""fareast""/><wx:font wx:val=""新細明體""/></w:rPr><w:t>　　【簡體字本】　　　　　　　</w:t>" & _
"</w:r><w:r><w:rPr><w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體"" w:hint=""fareast""/><wx:font wx:val=""新細明體""/>" & _
"</w:rPr><w:t>#seq#-SimplifiedSpec.pdf</w:t></w:r></w:p>" & _
"<w:p wsp:rsidR=""008515F6"" wsp:rsidRDefault=""008515F6"" wsp:rsidP=""008515F6""><w:pPr><w:rPr><w:rFonts " & _
"w:ascii=""新細明體"" w:h-ansi=""新細明體""/><wx:font wx:val=""新細明體""/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=""Times " & _
"New Roman"" w:h-ansi=""新細明體"" w:hint=""fareast""/><wx:font wx:val=""新細明體""/></w:rPr><w:t>　　【國際優先權證明文件】　　</w:t>" & _
"</w:r><w:r><w:rPr><w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體"" w:hint=""fareast""/><wx:font wx:val=""新細明體""/>" & _
"</w:rPr><w:t>#seq#-Priority.pdf</w:t></w:r></w:p>" & _
"<w:p wsp:rsidR=""008515F6"" wsp:rsidRDefault=""008515F6"" wsp:rsidP=""008515F6""><w:pPr><w:rPr><w:rFonts " & _
"w:ascii=""新細明體"" w:h-ansi=""新細明體""/><wx:font wx:val=""新細明體""/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=""Times " & _
"New Roman"" w:h-ansi=""新細明體"" w:hint=""fareast""/><wx:font wx:val=""新細明體""/></w:rPr><w:t>　　【優惠期證明文件】　　　　</w:t>" & _
"</w:r><w:r><w:rPr><w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體"" w:hint=""fareast""/><wx:font wx:val=""新細明體""/>" & _
"</w:rPr><w:t>#seq#-ICExperiment.pdf</w:t></w:r></w:p>" & _
"<w:p wsp:rsidR=""008515F6"" wsp:rsidRDefault=""008515F6"" wsp:rsidP=""008515F6""><w:pPr><w:rPr><w:rFonts " & _
"w:ascii=""新細明體"" w:h-ansi=""新細明體""/><wx:font wx:val=""新細明體""/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=""Times " & _
"New Roman"" w:h-ansi=""新細明體"" w:hint=""fareast""/><wx:font wx:val=""新細明體""/></w:rPr><w:t>　　【優惠期證明文件】　　　　</w:t>" & _
"</w:r><w:r><w:rPr><w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體"" w:hint=""fareast""/><wx:font wx:val=""新細明體""/>" & _
"</w:rPr><w:t>#seq#-Exhibition.pdf</w:t></w:r></w:p>" & _
"<w:p wsp:rsidR=""008515F6"" wsp:rsidRDefault=""008515F6"" wsp:rsidP=""008515F6""><w:pPr><w:rPr><w:rFonts " & _
"w:ascii=""新細明體"" w:h-ansi=""新細明體""/><wx:font wx:val=""新細明體""/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=""Times " & _
"New Roman"" w:h-ansi=""新細明體"" w:hint=""fareast""/><wx:font wx:val=""新細明體""/></w:rPr><w:t>　　【委任書】　　　　　　　　</w:t>" & _
"</w:r><w:r><w:rPr><w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體"" w:hint=""fareast""/><wx:font wx:val=""新細明體""/>" & _
"</w:rPr><w:t>#seq#-PowerAttorney.pdf</w:t></w:r></w:p>" & _
"<w:p wsp:rsidR=""008515F6"" wsp:rsidRDefault=""008515F6"" wsp:rsidP=""008515F6""><w:pPr><w:rPr><w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體""/><wx:font " & _
"wx:val=""新細明體""/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=""Times New Roman"" w:h-ansi=""新細明體"" w:hint=""fareast""/>" & _
"<wx:font wx:val=""新細明體""/></w:rPr><w:t>　　【國內生物材料寄存證明文件】</w:t></w:r><w:r><w:rPr>" & _
"<w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體"" w:hint=""fareast""/><wx:font wx:val=""新細明體""/></w:rPr><w:t>#seq#-FIRDI99999.pdf</w:t></w:r></w:p>" & _
"<w:p wsp:rsidR=""008515F6"" wsp:rsidRDefault=""008515F6"" wsp:rsidP=""008515F6""><w:pPr><w:rPr><w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體""/><wx:font " & _
"wx:val=""新細明體""/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=""Times New Roman"" w:h-ansi=""新細明體"" w:hint=""fareast""/>" & _
"<wx:font wx:val=""新細明體""/></w:rPr><w:t>　　【國外生物材料寄存證明文件】</w:t></w:r><w:r><w:rPr>" & _
"<w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體"" w:hint=""fareast""/><wx:font wx:val=""新細明體""/></w:rPr><w:t>#seq#-ATCC99999.pdf</w:t></w:r></w:p>" & _
"<w:p wsp:rsidR=""008515F6"" wsp:rsidRDefault=""008515F6"" wsp:rsidP=""008515F6""><w:pPr><w:rPr><w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體""/><wx:font " & _
"wx:val=""新細明體""/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=""Times New Roman"" w:h-ansi=""新細明體"" w:hint=""fareast""/>" & _
"<wx:font wx:val=""新細明體""/></w:rPr><w:t>　　【生物材料為通常知識者易於獲得證明文件】</w:t></w:r><w:r><w:rPr>" & _
"<w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體"" w:hint=""fareast""/><wx:font wx:val=""新細明體""/></w:rPr><w:t>#seq#-EasilyObtained.pdf</w:t></w:r></w:p>" & _
"<w:p wsp:rsidR=""003D09F2"" wsp:rsidRDefault=""003D09F2"" wsp:rsidP=""003D09F2""><w:pPr><w:snapToGrid w:val=""off""/><w:spacing " & _
"w:line=""288"" w:line-rule=""auto""/></w:pPr><w:r><w:rPr><w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體"" w:hint=""fareast""/>" & _
"<wx:font wx:val=""新細明體""/><w:color w:val=""000000""/></w:rPr><w:t>　　【其他】</w:t></w:r></w:p>" & _
"<w:p wsp:rsidR=""003D09F2"" " & _
"wsp:rsidRDefault=""003D09F2"" wsp:rsidP=""003D09F2""><w:pPr><w:snapToGrid w:val=""off""/><w:spacing w:line=""288"" w:line-rule=""auto""/>" & _
"</w:pPr><w:r><w:rPr><w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體"" w:hint=""fareast""/><wx:font wx:val=""新細明體""/>" & _
"<w:color w:val=""000000""/></w:rPr><w:t>　　　【文件描述】　　　　　　</w:t></w:r></w:p>" & _
"<w:p wsp:rsidR=""003D09F2"" " & _
"wsp:rsidRDefault=""003D09F2"" wsp:rsidP=""003D09F2""><w:pPr><w:snapToGrid w:val=""off""/><w:spacing w:line=""288"" w:line-rule=""auto""/>" & _
"</w:pPr><w:r><w:rPr><w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體"" w:hint=""fareast""/><wx:font wx:val=""新細明體""/>" & _
"<w:color w:val=""000000""/></w:rPr><w:t>　　　【文件檔名】　　　　　　</w:t></w:r></w:p>" & _
"<w:p wsp:rsidR=""008515F6"" wsp:rsidRDefault=""008515F6"" " & _
"wsp:rsidP=""008515F6""><w:pPr><w:rPr><w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體""/><wx:font wx:val=""新細明體""/>" & _
"</w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=""Times New Roman"" w:h-ansi=""新細明體"" w:hint=""fareast""/><wx:font wx:val=""新細明體""/>" & _
"</w:rPr><w:t>【中文本原始檔】　　　　　　　</w:t></w:r><w:r><w:rPr>" & _
"<w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體"" w:hint=""fareast""/><wx:font wx:val=""新細明體""/></w:rPr><w:t>#seq#-desc.doc</w:t></w:r></w:p>" & _
"<w:p wsp:rsidR=""008515F6"" wsp:rsidRDefault=""008515F6"" wsp:rsidP=""008515F6"">" & _
"<w:pPr><w:rPr><w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體""/><wx:font wx:val=""新細明體""/></w:rPr></w:pPr><w:r>" & _
"<w:rPr><w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體"" w:hint=""fareast""/><wx:font wx:val=""新細明體""/></w:rPr><w:t></w:t>" & _
"</w:r></w:p>" & _
"<w:p wsp:rsidR=""008515F6"" wsp:rsidRDefault=""008515F6"" wsp:rsidP=""008515F6""><w:pPr><w:rPr><w:rFonts w:ascii=""新細明體"" " & _
"w:h-ansi=""新細明體""/><wx:font wx:val=""新細明體""/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=""Times New Roman"" w:h-ansi=""新細明體"" " & _
"w:hint=""fareast""/><wx:font wx:val=""新細明體""/></w:rPr><w:t>【本申請書所檢送之PDF檔或影像檔與原本或正本相同】</w:t></w:r></w:p>" & _

"<w:p wsp:rsidR=""008515F6"" wsp:rsidRDefault=""008515F6"" wsp:rsidP=""008515F6""><w:pPr><w:rPr><w:rFonts w:ascii=""新細明體"" " & _
"w:h-ansi=""新細明體""/><wx:font wx:val=""新細明體""/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=""Times New Roman"" w:h-ansi=""新細明體"" " & _
"w:hint=""fareast""/><wx:font wx:val=""新細明體""/></w:rPr><w:t>【申請人已詳閱申請須知所定個人資料保護注意事項-並已確認本申請案之附件-除基本資料表-委任書外-不包含應予保密之個人資料-其載有個人資料者-同意智慧財產局提供任何人以自動化或非自動化之方式閱覽或抄錄或攝影或影印.】</w:t></w:r></w:p>" & _

"<w:p wsp:rsidR=""000940B1"" " & _
"wsp:rsidRDefault=""00397406""><w:pPr><w:overflowPunct w:val=""off""/><w:autoSpaceDE w:val=""off""/><w:autoSpaceDN w:val=""off""/><w:spacing " & _
"w:line=""360"" w:line-rule=""at-least""/><w:rPr><w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體""/><wx:font wx:val=""新細明體""/>" & _
"</w:rPr><w:sectPr wsp:rsidR=""000940B1""><w:ftr w:type=""odd"">" & _
"<w:p wsp:rsidR=""000940B1"" wsp:rsidRDefault=""000940B1"" wsp:rsidP=""000940B1"">" & _
"<w:pPr><w:pStyle w:val=""a7""/><w:jc w:val=""center""/></w:pPr><w:r><w:rPr><w:rFonts w:hint=""fareast""/><wx:font wx:val=""新細明體""/>" & _
"</w:rPr><w:t>第</w:t></w:r><w:fldSimple w:instr="" PAGE   \* MERGEFORMAT ""><w:r wsp:rsidR=""00DE2999""><w:rPr><w:noProof/></w:rPr>" & _
"<w:t>1</w:t></w:r></w:fldSimple><w:r><w:rPr><w:rFonts w:hint=""fareast""/><wx:font wx:val=""新細明體""/></w:rPr><w:t>頁，共</w:t>" & _
"</w:r><w:fldSimple w:instr="" SECTIONPAGES  \* MERGEFORMAT ""><w:r wsp:rsidR=""00DE2999""><w:rPr><w:noProof/></w:rPr><w:t>3</w:t></w:r>" & _
"</w:fldSimple><w:r><w:rPr><w:rFonts w:hint=""fareast""/><wx:font wx:val=""新細明體""/></w:rPr><w:t>頁</w:t></w:r><w:r><w:rPr><w:rFonts " & _
"w:hint=""fareast""/></w:rPr><w:t>(</w:t></w:r><w:r wsp:rsidRPr=""000940B1""><w:rPr><w:rFonts w:hint=""fareast""/><wx:font wx:val=""新細明體""/>" & _
"</w:rPr><w:t>發明專利申請書</w:t></w:r><w:r><w:rPr><w:rFonts w:hint=""fareast""/></w:rPr><w:t>)</w:t></w:r></w:p></w:ftr><w:pgSz " & _
"w:w=""11906"" w:h=""16838""/><w:pgMar w:top=""1134"" w:right=""1134"" w:bottom=""1134"" w:left=""1134"" w:header=""851"" w:footer=""992"" " & _
"w:gutter=""0""/><w:cols w:space=""425""/><w:docGrid w:type=""lines"" w:line-pitch=""360""/></w:sectPr></w:pPr><w:r><w:rPr><w:rFonts " & _
"w:ascii=""新細明體"" w:h-ansi=""新細明體"" w:hint=""fareast""/><wx:font wx:val=""新細明體""/></w:rPr>" & _
"</w:r></w:p>"
END function


'基本資料表 個人資料 
Function DocBody_11()
DocBody_11 = "<w:p wsp:rsidR=""008515F6"" wsp:rsidRDefault=""008515F6"" wsp:rsidP=""008515F6""><w:pPr><w:rPr><w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體""/>" & _
"<wx:font wx:val=""新細明體""/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=""Times New Roman"" w:h-ansi=""新細明體"" w:hint=""fareast""/>" & _
"<wx:font wx:val=""新細明體""/></w:rPr><w:t>【基本資料】　　　</w:t>" & _
"</w:r></w:p><w:p wsp:rsidR=""008515F6"" wsp:rsidRDefault=""008515F6"" wsp:rsidP=""008515F6""><w:pPr><w:rPr><w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體""/>" & _
"<wx:font wx:val=""新細明體""/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=""Times New Roman"" w:h-ansi=""新細明體"" w:hint=""fareast""/>" & _
"<wx:font wx:val=""新細明體""/></w:rPr><w:t>【個人資料】　　　</w:t>" & _
"</w:r></w:p>"
End function


'申請人迴圈
Function Dmp_apcust_data_1()
Dmp_apcust_data_1 = "<w:p wsp:rsidR=""008515F6"" wsp:rsidRDefault=""00565151"" " & _
"wsp:rsidP=""00565151""><w:pPr><w:pStyle w:val=""ac""/><w:ind w:left=""0""/><w:rPr><w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體""/>" & _
"<wx:font wx:val=""新細明體""/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=""Times New Roman"" w:h-ansi=""新細明體"" w:hint=""fareast""/>" & _
"<wx:font wx:val=""新細明體""/></w:rPr><w:t>　　【申請人#apply_num#】</w:t></w:r></w:p>" & _
"<w:p wsp:rsidR=""00123906"" wsp:rsidRDefault=""00123906"" wsp:rsidP=""00123906"">" & _
"<w:pPr><w:snapToGrid w:val=""off""/><w:spacing w:line=""288"" w:line-rule=""auto""/></w:pPr><w:r><w:rPr><w:rFonts w:hint=""fareast""/>" & _
"<w:color w:val=""000000""/></w:rPr><w:t>　　　【國籍】　　　　　　</w:t></w:r><w:r><w:rPr><w:rFonts w:hint=""fareast""/>" & _
"</w:rPr><w:t>#ap_country#</w:t></w:r></w:p>" & _
"<w:p wsp:rsidR=""00123906"" " & _
"wsp:rsidRDefault=""00123906"" wsp:rsidP=""00123906""><w:pPr><w:snapToGrid w:val=""off""/><w:spacing w:line=""288"" w:line-rule=""auto""/>" & _
"</w:pPr><w:r><w:rPr><w:rFonts w:hint=""fareast""/><w:color w:val=""000000""/></w:rPr><w:t>　　　【身分種類】　　　　</w:t>" & _
"</w:r><w:r><w:rPr><w:rFonts w:hint=""fareast""/></w:rPr><w:t>#ap_class#</w:t></w:r></w:p>" 
end function
function Dmp_apcust_data_1_2()
Dmp_apcust_data_1_2 = "<w:p wsp:rsidR=""00123906"" wsp:rsidRDefault=""00123906"" wsp:rsidP=""00123906""><w:pPr><w:snapToGrid w:val=""off""/>" & _
"<w:spacing w:line=""288"" w:line-rule=""auto""/></w:pPr><w:r><w:rPr><w:rFonts w:hint=""fareast""/><w:color w:val=""000000""/></w:rPr>" & _
"<w:t>　　　【ID】　　　　　　　</w:t></w:r><w:r><w:rPr><w:rFonts w:hint=""fareast""/></w:rPr><w:t>#apcust_no#</w:t></w:r></w:p>" 
end function
function Dmp_apcust_data_1_3()
Dmp_apcust_data_1_3 = "<w:p wsp:rsidR=""00123906"" wsp:rsidRDefault=""00123906"" " & _
"wsp:rsidP=""00123906""><w:pPr><w:snapToGrid w:val=""off""/><w:spacing w:line=""288"" w:line-rule=""auto""/><w:rPr><w:rFonts w:hint=""fareast""/>" & _
"</w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:hint=""fareast""/><w:color w:val=""000000""/></w:rPr><w:t>　　　【#ap_cname1_title#】　　　　</w:t>" & _
"</w:r><w:r><w:rPr><w:rFonts w:hint=""fareast""/></w:rPr><w:t>#ap_cname1#</w:t></w:r></w:p>" & _
"<w:p wsp:rsidR=""00123906"" wsp:rsidRPr=""002F4BE0"" wsp:rsidRDefault=""00123906"" " & _
"wsp:rsidP=""00123906""><w:pPr><w:snapToGrid w:val=""off""/><w:spacing w:line=""288"" w:line-rule=""auto""/></w:pPr><w:r><w:rPr><w:rFonts " & _
"w:hint=""fareast""/><w:color w:val=""000000""/></w:rPr><w:t>　　　【#ap_ename1_title#】　　　　</w:t></w:r><w:r><w:rPr><w:rFonts " & _
"w:hint=""fareast""/></w:rPr><w:t>#ap_ename1#</w:t></w:r></w:p>" & _
"<w:p wsp:rsidR=""00123906"" wsp:rsidRDefault=""00123906"" wsp:rsidP=""00123906""><w:pPr><w:snapToGrid " & _
"w:val=""off""/><w:spacing w:line=""288"" w:line-rule=""auto""/></w:pPr><w:r><w:rPr><w:rFonts w:hint=""fareast""/><w:color w:val=""000000""/>" & _
"</w:rPr><w:t>　　　【居住國】　　　　　</w:t></w:r><w:r><w:rPr><w:rFonts w:ascii=""Times New " & _
"Roman"" w:h-ansi=""Times New Roman"" w:cs=""Times New Roman"" w:hint=""fareast""/><wx:font wx:val=""Times New Roman""/><w:color w:val=""000000""/>" & _
"</w:rPr><w:t>#ap_country#</w:t></w:r></w:p>" & _
"<w:p wsp:rsidR=""00123906"" wsp:rsidRDefault=""00123906"" wsp:rsidP=""00123906"">" & _
"<w:pPr><w:snapToGrid w:val=""off""/><w:spacing w:line=""288"" w:line-rule=""auto""/></w:pPr><w:r><w:rPr><w:rFonts w:hint=""fareast""/>" & _
"<w:color w:val=""000000""/></w:rPr><w:t>　　　【郵遞區號】　　　　</w:t></w:r><w:r><w:rPr><w:rFonts w:ascii=""Times New " & _
"Roman"" w:h-ansi=""Times New Roman"" w:cs=""Times New Roman"" w:hint=""fareast""/><wx:font wx:val=""Times New Roman""/><w:color w:val=""000000""/>" & _
"</w:rPr><w:t>#ap_zip#</w:t></w:r></w:p>" & _
"<w:p wsp:rsidR=""00123906"" wsp:rsidRDefault=""00123906"" wsp:rsidP=""00123906""><w:pPr><w:snapToGrid w:val=""off""/><w:spacing w:line=""288"" " & _
"w:line-rule=""auto""/><w:rPr><w:rFonts w:hint=""fareast""/><w:color w:val=""000000""/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:hint=""fareast""/>" & _
"<w:color w:val=""000000""/></w:rPr><w:t>　　　【中文地址】　　　　#ap_addr1##ap_addr2#</w:t></w:r></w:p>" & _
"<w:p wsp:rsidR=""00123906"" " & _
"wsp:rsidRDefault=""00123906"" wsp:rsidP=""00123906""><w:pPr><w:snapToGrid w:val=""off""/><w:spacing w:line=""288"" w:line-rule=""auto""/>" & _
"<w:rPr><w:rFonts w:ascii=""Times New Roman"" w:h-ansi=""Times New Roman"" w:cs=""Times New Roman"" w:hint=""fareast""/><wx:font wx:val=""Times " & _
"New Roman""/><w:color w:val=""000000""/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:hint=""fareast""/><w:color w:val=""000000""/></w:rPr>" & _
"<w:t>　　　【英文地址】　　　　</w:t></w:r><w:r><w:rPr><w:rFonts w:ascii=""Times New Roman"" w:h-ansi=""Times New Roman"" " & _
"w:cs=""Times New Roman"" w:hint=""fareast""/><wx:font wx:val=""Times New Roman""/><w:color w:val=""000000""/></w:rPr><w:t>#ap_eaddr1##ap_eaddr2##ap_eaddr3##ap_eaddr4#</w:t>" & _
"</w:r></w:p>" & _
"<w:p wsp:rsidR=""00123906"" wsp:rsidRDefault=""00123906"" wsp:rsidP=""00123906""><w:pPr><w:snapToGrid w:val=""off""/><w:spacing " & _
"w:line=""288"" w:line-rule=""auto""/></w:pPr><w:r><w:rPr><w:rFonts w:hint=""fareast""/><w:color w:val=""000000""/></w:rPr><w:t>　　　【代表人中文姓名】　#ap_crep#</w:t></w:r>" & _
"</w:p><w:p wsp:rsidR=""00123906"" wsp:rsidRDefault=""00123906"" wsp:rsidP=""00123906""><w:pPr><w:snapToGrid w:val=""off""/><w:spacing " & _
"w:line=""288"" w:line-rule=""auto""/><w:rPr><w:rFonts w:hint=""fareast""/><w:color w:val=""000000""/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts " & _
"w:hint=""fareast""/><w:color w:val=""000000""/></w:rPr><w:t>　　　【代表人英文姓名】　#ap_erep#</w:t></w:r></w:p>" & _
"<w:p wsp:rsidR=""00123906"" wsp:rsidRDefault=""00123906"" wsp:rsidP=""00123906""><w:pPr><w:snapToGrid w:val=""off""/><w:spacing " & _
"w:line=""288"" w:line-rule=""auto""/></w:pPr><w:r><w:rPr><w:rFonts w:hint=""fareast""/><w:color w:val=""000000""/></w:rPr><w:t>　　　【法定代理人ID】　　</w:t></w:r>" & _
"</w:p><w:p wsp:rsidR=""00123906"" wsp:rsidRDefault=""00123906"" wsp:rsidP=""00123906""><w:pPr><w:snapToGrid w:val=""off""/><w:spacing " & _
"w:line=""288"" w:line-rule=""auto""/><w:rPr><w:rFonts w:hint=""fareast""/><w:color w:val=""000000""/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts " & _
"w:hint=""fareast""/><w:color w:val=""000000""/></w:rPr><w:t>　　　【法定代理人中文姓名】</w:t></w:r></w:p>"
End function


'代理人1
Function Agt_data_3()
Agt_data_3 = "<w:p wsp:rsidR=""00754829"" wsp:rsidRDefault=""00754829"" wsp:rsidP=""00754829""><w:pPr><w:overflowPunct w:val=""off""/>" & _
"<w:autoSpaceDE w:val=""off""/><w:autoSpaceDN w:val=""off""/><w:spacing w:line=""360"" w:line-rule=""at-least""/></w:pPr><w:r><w:rPr>" & _
"<w:rFonts w:hint=""fareast""/></w:rPr><w:t>　　【代理人1】</w:t></w:r></w:p><w:p wsp:rsidR=""00123906"" wsp:rsidRDefault=""00123906"" " & _
"wsp:rsidP=""00123906""><w:pPr><w:snapToGrid w:val=""off""/><w:spacing w:line=""288"" w:line-rule=""auto""/></w:pPr><w:r><w:rPr><w:rFonts " & _
"w:hint=""fareast""/><w:color w:val=""000000""/></w:rPr><w:t>　　　【證書字號】　　　　</w:t></w:r><w:r><w:rPr><w:rFonts " & _
"w:ascii=""Times New Roman"" w:h-ansi=""Times New Roman"" w:cs=""Times New Roman"" w:hint=""fareast""/><wx:font wx:val=""Times New " & _
"Roman""/><w:color w:val=""000000""/></w:rPr><w:t>#agt_idno1#</w:t></w:r></w:p><w:p wsp:rsidR=""00123906"" wsp:rsidRDefault=""00123906"" " & _
"wsp:rsidP=""00123906""><w:pPr><w:snapToGrid w:val=""off""/><w:spacing w:line=""288"" w:line-rule=""auto""/></w:pPr><w:r><w:rPr><w:rFonts " & _
"w:hint=""fareast""/><w:color w:val=""000000""/></w:rPr><w:t>　　　【ID】　　　　　　　</w:t></w:r><w:r><w:rPr>" & _
"<w:rFonts w:ascii=""Times New Roman"" w:h-ansi=""Times New Roman"" w:cs=""Times New Roman"" w:hint=""fareast""/><wx:font wx:val=""Times " & _
"New Roman""/><w:color w:val=""000000""/></w:rPr><w:t>#agt_id1#</w:t></w:r></w:p><w:p wsp:rsidR=""00123906"" wsp:rsidRDefault=""00123906"" " & _
"wsp:rsidP=""00123906""><w:pPr><w:snapToGrid w:val=""off""/><w:spacing w:line=""288"" w:line-rule=""auto""/><w:rPr><w:rFonts w:hint=""fareast""/>" & _
"<w:color w:val=""000000""/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:hint=""fareast""/><w:color w:val=""000000""/></w:rPr><w:t>　　　【中文姓名】　　　　#agt_name1#</w:t>" & _
"</w:r></w:p><w:p wsp:rsidR=""00123906"" wsp:rsidRDefault=""00123906"" wsp:rsidP=""00123906""><w:pPr><w:snapToGrid w:val=""off""/><w:spacing " & _
"w:line=""288"" w:line-rule=""auto""/></w:pPr><w:r><w:rPr><w:rFonts w:hint=""fareast""/><w:color w:val=""000000""/></w:rPr><w:t>　　　【郵遞區號】　　　　</w:t>" & _
"</w:r><w:r><w:rPr><w:rFonts w:ascii=""Times New Roman"" w:h-ansi=""Times New Roman"" w:cs=""Times New Roman"" w:hint=""fareast""/><wx:font " & _
"wx:val=""Times New Roman""/><w:color w:val=""000000""/></w:rPr><w:t>#agt_zip#</w:t></w:r></w:p><w:p wsp:rsidR=""00123906"" wsp:rsidRDefault=""00123906"" wsp:rsidP=""00123906"">" & _
"<w:pPr><w:snapToGrid w:val=""off""/><w:spacing w:line=""288"" w:line-rule=""auto""/></w:pPr><w:r><w:rPr><w:rFonts w:hint=""fareast""/>" & _
"<w:color w:val=""000000""/></w:rPr><w:t>　　　【中文地址】　　　　#agt_addr#</w:t></w:r></w:p><w:p wsp:rsidR=""00123906"" wsp:rsidRDefault=""00123906"" " & _
"wsp:rsidP=""00123906""><w:pPr><w:snapToGrid w:val=""off""/><w:spacing w:line=""288"" w:line-rule=""auto""/><w:rPr><w:rFonts w:ascii=""Times " & _
"New Roman"" w:h-ansi=""Times New Roman"" w:cs=""Times New Roman"" w:hint=""fareast""/><wx:font wx:val=""Times New Roman""/><w:color " & _
"w:val=""000000""/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:hint=""fareast""/><w:color w:val=""000000""/></w:rPr><w:t>　　　【電話】　　　　　　#agt_tel#</w:t></w:r>" & _
"</w:p><w:p wsp:rsidR=""00123906"" wsp:rsidRDefault=""00123906"" wsp:rsidP=""00123906""><w:pPr><w:snapToGrid w:val=""off""/><w:spacing " & _
"w:line=""288"" w:line-rule=""auto""/><w:rPr><w:rFonts w:hint=""fareast""/><w:color w:val=""000000""/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts " & _
"w:hint=""fareast""/><w:color w:val=""000000""/></w:rPr><w:t>　　　【傳真】　　　　　　#agt_fax#</w:t></w:r></w:p><w:p wsp:rsidR=""00123906"" wsp:rsidRPr=""00FE23DF"" " & _
"wsp:rsidRDefault=""00123906"" wsp:rsidP=""00123906""><w:pPr><w:overflowPunct w:val=""off""/><w:autoSpaceDE w:val=""off""/><w:autoSpaceDN " & _
"w:val=""off""/><w:spacing w:line=""360"" w:line-rule=""at-least""/><w:rPr><w:rFonts w:ascii=""Calibri"" w:h-ansi=""Calibri"" w:hint=""fareast""/>" & _
"<wx:font wx:val=""Calibri""/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:hint=""fareast""/><w:color w:val=""000000""/></w:rPr><w:t>　　　【E-mail】　　　　　　</w:t>" & _
"</w:r><w:r wsp:rsidRPr=""00EF5BA9""><w:fldChar w:fldCharType=""begin""/></w:r><w:r wsp:rsidRPr=""00EF5BA9""><w:instrText> HYPERLINK " & _
"""mailto:siiplo@mail.saint-island.com.tw"" </w:instrText></w:r><w:r wsp:rsidRPr=""00EF5BA9""><w:fldChar w:fldCharType=""separate""/>" & _
"</w:r><w:r wsp:rsidRPr=""00EF5BA9""><w:t>siiplo@mail.saint-island.com.tw</w:t></w:r><w:r wsp:rsidRPr=""00EF5BA9""><w:fldChar w:fldCharType=""end""/>" & _
"</w:r></w:p>"
End function


'代理人2
Function Agt_data_4()
Agt_data_4 = "<w:p wsp:rsidR=""00754829"" wsp:rsidRDefault=""00754829"" wsp:rsidP=""00754829""><w:pPr><w:overflowPunct w:val=""off""/><w:autoSpaceDE " & _
"w:val=""off""/><w:autoSpaceDN w:val=""off""/><w:spacing w:line=""360"" w:line-rule=""at-least""/></w:pPr><w:r><w:rPr><w:rFonts w:hint=""fareast""/>" & _
"</w:rPr><w:t>　　【代理人2】</w:t></w:r></w:p><w:p wsp:rsidR=""00123906"" wsp:rsidRDefault=""00123906"" wsp:rsidP=""00123906"">" & _
"<w:pPr><w:snapToGrid w:val=""off""/><w:spacing w:line=""288"" w:line-rule=""auto""/></w:pPr><w:r><w:rPr><w:rFonts w:hint=""fareast""/>" & _
"<w:color w:val=""000000""/></w:rPr><w:t>　　　【證書字號】　　　　</w:t></w:r><w:r><w:rPr><w:rFonts w:ascii=""Times New " & _
"Roman"" w:h-ansi=""Times New Roman"" w:cs=""Times New Roman"" w:hint=""fareast""/><wx:font wx:val=""Times New Roman""/><w:color w:val=""000000""/>" & _
"</w:rPr><w:t>#agt_idno2#</w:t></w:r></w:p><w:p wsp:rsidR=""00123906"" wsp:rsidRDefault=""00123906"" wsp:rsidP=""00123906""><w:pPr><w:snapToGrid " & _
"w:val=""off""/><w:spacing w:line=""288"" w:line-rule=""auto""/></w:pPr><w:r><w:rPr><w:rFonts w:hint=""fareast""/><w:color w:val=""000000""/>" & _
"</w:rPr><w:t>　　　【ID】　　　　　　　</w:t></w:r><w:r><w:rPr><w:rFonts w:ascii=""Times New Roman"" w:h-ansi=""Times " & _
"New Roman"" w:cs=""Times New Roman"" w:hint=""fareast""/><wx:font wx:val=""Times New Roman""/><w:color w:val=""000000""/></w:rPr><w:t>#agt_id2#</w:t>" & _
"</w:r></w:p><w:p wsp:rsidR=""00123906"" wsp:rsidRDefault=""00123906"" wsp:rsidP=""00123906""><w:pPr><w:snapToGrid w:val=""off""/><w:spacing " & _
"w:line=""288"" w:line-rule=""auto""/><w:rPr><w:rFonts w:hint=""fareast""/><w:color w:val=""000000""/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts " & _
"w:hint=""fareast""/><w:color w:val=""000000""/></w:rPr><w:t>　　　【中文姓名】　　　　#agt_name2#</w:t></w:r></w:p><w:p " & _
"wsp:rsidR=""00123906"" wsp:rsidRDefault=""00123906"" wsp:rsidP=""00123906""><w:pPr><w:snapToGrid w:val=""off""/><w:spacing w:line=""288"" " & _
"w:line-rule=""auto""/></w:pPr><w:r><w:rPr><w:rFonts w:hint=""fareast""/><w:color w:val=""000000""/></w:rPr><w:t>　　　【郵遞區號】　　　　</w:t>" & _
"</w:r><w:r><w:rPr><w:rFonts w:ascii=""Times New Roman"" w:h-ansi=""Times New Roman"" w:cs=""Times New Roman""/><wx:font wx:val=""Times " & _
"New Roman""/><w:color w:val=""000000""/></w:rPr><w:t>#agt_zip#</w:t></w:r></w:p><w:p wsp:rsidR=""00123906"" wsp:rsidRDefault=""00123906"" " & _
"wsp:rsidP=""00123906""><w:pPr><w:snapToGrid w:val=""off""/><w:spacing w:line=""288"" w:line-rule=""auto""/></w:pPr><w:r><w:rPr><w:rFonts " & _
"w:hint=""fareast""/><w:color w:val=""000000""/></w:rPr><w:t>　　　【中文地址】　　　　#agt_addr#</w:t></w:r></w:p><w:p wsp:rsidR=""00123906"" wsp:rsidRDefault=""00123906"" " & _
"wsp:rsidP=""00123906""><w:pPr><w:snapToGrid w:val=""off""/><w:spacing w:line=""288"" w:line-rule=""auto""/><w:rPr><w:rFonts w:ascii=""Times " & _
"New Roman"" w:h-ansi=""Times New Roman"" w:cs=""Times New Roman"" w:hint=""fareast""/><wx:font wx:val=""Times New Roman""/><w:color " & _
"w:val=""000000""/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:hint=""fareast""/><w:color w:val=""000000""/></w:rPr><w:t>　　　【電話】　　　　　　#agt_tel#</w:t></w:r>" & _
"</w:p><w:p wsp:rsidR=""00123906"" wsp:rsidRDefault=""00123906"" wsp:rsidP=""00123906""><w:pPr><w:snapToGrid w:val=""off""/><w:spacing " & _
"w:line=""288"" w:line-rule=""auto""/><w:rPr><w:rFonts w:hint=""fareast""/><w:color w:val=""000000""/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts " & _
"w:hint=""fareast""/><w:color w:val=""000000""/></w:rPr><w:t>　　　【傳真】　　　　　　#agt_fax#</w:t></w:r></w:p><w:p wsp:rsidR=""00123906"" wsp:rsidRPr=""00FE23DF"" " & _
"wsp:rsidRDefault=""00123906"" wsp:rsidP=""00123906""><w:pPr><w:overflowPunct w:val=""off""/><w:autoSpaceDE w:val=""off""/><w:autoSpaceDN " & _
"w:val=""off""/><w:spacing w:line=""360"" w:line-rule=""at-least""/><w:rPr><w:rFonts w:ascii=""Calibri"" w:h-ansi=""Calibri"" w:hint=""fareast""/>" & _
"<wx:font wx:val=""Calibri""/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:hint=""fareast""/><w:color w:val=""000000""/></w:rPr><w:t>　　　【E-mail】　　　　　　</w:t>" & _
"</w:r><w:r wsp:rsidRPr=""00EF5BA9""><w:fldChar w:fldCharType=""begin""/></w:r><w:r wsp:rsidRPr=""00EF5BA9""><w:instrText> HYPERLINK " & _
"""mailto:siiplo@mail.saint-island.com.tw"" </w:instrText></w:r><w:r wsp:rsidRPr=""00EF5BA9""><w:fldChar w:fldCharType=""separate""/>" & _
"</w:r><w:r wsp:rsidRPr=""00EF5BA9""><w:t>siiplo@mail.saint-island.com.tw</w:t></w:r><w:r wsp:rsidRPr=""00EF5BA9""><w:fldChar w:fldCharType=""end""/>" & _
"</w:r></w:p>"



End function


'發明人迴圈
Function Ant_data_1()
Ant_data_1 = "<w:p wsp:rsidR=""008515F6"" wsp:rsidRDefault=""00565151"" " & _
"wsp:rsidP=""00565151""><w:pPr><w:pStyle w:val=""ac""/><w:ind w:left=""0""/><w:rPr><w:rFonts w:ascii=""新細明體"" w:h-ansi=""新細明體""/>" & _
"<wx:font wx:val=""新細明體""/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=""Times New Roman"" w:h-ansi=""新細明體"" w:hint=""fareast""/>" & _
"<wx:font wx:val=""新細明體""/></w:rPr><w:t>　　【#ant_num#】</w:t></w:r></w:p><w:p wsp:rsidR=""00123906"" wsp:rsidRDefault=""00123906"" wsp:rsidP=""00123906"">" & _
"<w:pPr><w:snapToGrid w:val=""off""/><w:spacing w:line=""288"" w:line-rule=""auto""/></w:pPr><w:r><w:rPr><w:rFonts w:hint=""fareast""/>" & _
"<w:color w:val=""000000""/></w:rPr><w:t>　　　【國籍】　　　　　　</w:t></w:r><w:r><w:rPr><w:rFonts w:ascii=""Times New " & _
"Roman"" w:h-ansi=""Times New Roman"" w:cs=""Times New Roman"" w:hint=""fareast""/><wx:font wx:val=""Times New Roman""/><w:color w:val=""000000""/>" & _
"</w:rPr><w:t>#ant_country#</w:t></w:r></w:p>"
END Function
'發明人迴圈
Function Ant_data_1_1()
Ant_data_1_1 = "<w:p " & _
"wsp:rsidR=""00123906"" wsp:rsidRDefault=""00123906"" wsp:rsidP=""00123906""><w:pPr><w:snapToGrid w:val=""off""/><w:spacing w:line=""288"" " & _
"w:line-rule=""auto""/></w:pPr><w:r><w:rPr><w:rFonts w:hint=""fareast""/><w:color w:val=""000000""/></w:rPr><w:t>　　　【ID】　　　　　　　</w:t></w:r><w:r><w:rPr><w:rFonts w:ascii=""Times New Roman"" w:h-ansi=""Times New Roman"" w:cs=""Times " & _
"New Roman"" w:hint=""fareast""/><wx:font wx:val=""Times New Roman""/><w:color w:val=""000000""/></w:rPr><w:t>#ant_id#</w:t></w:r></w:p>"
End function

Function Ant_data_1_2()
Ant_data_1_2 = "<w:p wsp:rsidR=""00123906"" wsp:rsidRDefault=""00123906"" " & _
"wsp:rsidP=""00123906""><w:pPr><w:snapToGrid w:val=""off""/><w:spacing w:line=""288"" w:line-rule=""auto""/><w:rPr><w:color w:val=""000000""/>" & _
"</w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:hint=""fareast""/><w:color w:val=""000000""/></w:rPr><w:t>　　　【中文姓名】　　　　#ant_cname#</w:t></w:r>" & _
"</w:p><w:p wsp:rsidR=""00754829"" wsp:rsidRPr=""00123906"" wsp:rsidRDefault=""00123906"" wsp:rsidP=""00123906""><w:pPr><w:snapToGrid " & _
"w:val=""off""/><w:spacing w:line=""288"" w:line-rule=""auto""/><w:rPr><w:color w:val=""000000""/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts " & _
"w:hint=""fareast""/><w:color w:val=""000000""/></w:rPr><w:t>　　　【英文姓名】　　　　#ant_ename#</w:t></w:r></w:p>"
End function


Function DocTail_1()
DocTail_1 = "<w:sectPr wsp:rsidR=""00397406"" wsp:rsidRPr=""008515F6"" wsp:rsidSect=""008515F6"">" & _
"<w:ftr w:type=""odd""><w:p wsp:rsidR=""00A97328"" wsp:rsidRPr=""00D01292"" wsp:rsidRDefault=""00945317"" wsp:rsidP=""00A97328""><w:pPr>" & _
"<w:pStyle w:val=""a7""/><w:jc w:val=""center""/></w:pPr><w:r><w:rPr><w:rFonts w:hint=""fareast""/><wx:font wx:val=""新細明體""/>" & _
"</w:rPr><w:t>第</w:t></w:r><w:fldSimple w:instr="" PAGE   \* MERGEFORMAT ""><w:r wsp:rsidR=""00A47B5D""><w:rPr><w:noProof/></w:rPr>" & _
"<w:t>1</w:t></w:r></w:fldSimple><w:r><w:rPr><w:rFonts w:hint=""fareast""/><wx:font wx:val=""新細明體""/></w:rPr><w:t>頁，共</w:t>" & _
"</w:r><w:fldSimple w:instr="" SECTIONPAGES  \* MERGEFORMAT ""><w:r wsp:rsidR=""00A47B5D""><w:rPr><w:noProof/></w:rPr><w:t>1</w:t></w:r>" & _
"</w:fldSimple><w:r><w:rPr><w:rFonts w:hint=""fareast""/><wx:font wx:val=""新細明體""/></w:rPr><w:t>頁</w:t></w:r><w:r><w:rPr><w:rFonts " & _
"w:hint=""fareast""/></w:rPr><w:t>(</w:t></w:r><w:r><w:rPr><w:rFonts w:hint=""fareast""/><wx:font wx:val=""新細明體""/></w:rPr>" & _
"<w:t>基本資料表</w:t></w:r><w:r><w:rPr><w:rFonts w:hint=""fareast""/></w:rPr><w:t>)</w:t></w:r></w:p></w:ftr><w:pgSz w:w=""11906"" " & _
"w:h=""16838""/><w:pgMar w:top=""1134"" w:right=""1134"" w:bottom=""1134"" w:left=""1134"" w:header=""851"" w:footer=""992"" w:gutter=""0""/>" & _
"<w:pgNumType w:start=""1""/><w:cols w:space=""425""/><w:docGrid w:line-pitch=""360""/></w:sectPr></w:body></w:wordDocument>"
End function



%>