Option Explicit

Public mMsg As String
Public AppTittle As String
Public UniversalChar As String          ' This Constant is a kind of trapping key for all vowels
Public TrappedCharNormal As String      ' Trapping normal Word's malfunctioning char like " ' ~ "
Public TrappedCharReserved As String    ' Trapping normal Word's reserved keyword

' Set the parameter
Public Const wdLayout = 3
Public CodeArray As Variant
Public VowelArray As Variant
Public FontRecognizer As Variant
Public FontCodeAlt As Variant
Public UVowels As String

Public Type CodeTable
    t_1CodeName As String
    t_2VowelList As String
    t_3DefaultFontName As String
    t_4FontRecognizer As String
    t_5FontConversion As String
    t_6FontUpperCase As String
End Type

Function isArray(arrObj As Variant) As Boolean
Attribute isArray.VB_ProcData.VB_Invoke_Func = " \n14"
    ' Check if a variant is an array or not
    On Error GoTo ErrHandler
    Dim myExt As Long
    myExt = Val(arrObj(0))
    isArray = True
ErrHandler:
End Function

Function GetUnicodeString() As String
Attribute GetUnicodeString.VB_ProcData.VB_Invoke_Func = " \n14"
    ' This function is no longer kept but I still would like it to be here for some folks if they want to diggest
    Dim iUnicode As Variant ' array to keep unicode char set
    Dim i As Long, iStr As String
    iUnicode = Array(225, 224, 7843, 227, 7841, 259, 7855, 7857, 7859, 7861, 7863, 226, 7845, 7847, 7849, _
        7851, 7853, 233, 232, 7867, 7869, 7865, 234, 7871, 7873, 7875, 7877, 7879, 237, 236, 7881, 297, 7883, _
        243, 242, 7887, 245, 7885, 244, 7889, 7891, 7893, 7895, 7897, 417, 7899, 7901, 7903, 7905, 7907, 250, _
        249, 7911, 361, 7909, 432, 7913, 7915, 7917, 7919, 7921, 253, 7923, 7927, 7929, 7925, 273, 193, 192, _
        7842, 195, 7840, 258, 7854, 7856, 7858, 7860, 7862, 194, 7844, 7846, 7848, 7850, 7852, 201, 200, 7866, _
        7868, 7864, 202, 7870, 7872, 7874, 7876, 7878, 205, 204, 7880, 296, 7882, 211, 210, 7886, 213, 7884, _
        212, 7888, 7890, 7892, 7894, 7896, 416, 7898, 7900, 7902, 7904, 7906, 218, 217, 7910, 360, 7908, 431, _
        7912, 7914, 7916, 7918, 7920, 221, 7922, 7926, 7928, 7924, 272)
    For i = 0 To UBound(iUnicode)
        iStr = iStr & "/" & ChrW(iUnicode(i))
    Next
    GetUnicodeString = Mid(iStr, 2)
End Function

Function SupportCodes() As Variant
Attribute SupportCodes.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim iCodeArr As Variant
    iCodeArr = Array("Unicode", "TCVN-ABC", "VNI")
    SupportCodes = iCodeArr
End Function

Function GetCodetable(iCodeName As Variant) As CodeTable
Attribute GetCodetable.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim iCode As CodeTable
    With iCode
        Select Case iCodeName
        Case "Unicode":
            ' this just for unicode -precompound
            .t_2VowelList = GetUnicodeString()
            .t_3DefaultFontName = "Times New Roman"
            .t_4FontRecognizer = "Times N/Tahoma/Arial/Courr"
            .t_5FontConversion = "Times New Roman/Arial/Tahoma/Verdana/Courier"
            .t_6FontUpperCase = "*NONE"
        Case "TCVN-ABC":
            .t_2VowelList = "¸/µ/¶/·/¹/¨/¾/»/¼/½/Æ/©/Ê/Ç/È/É/Ë/Ð/Ì/Î/Ï/Ñ/ª/Õ/Ò/Ó/Ô/Ö/Ý/×/Ø/Ü/Þ/ã/ß/á/â/ä/«/è/å/æ/ç/é/¬/í/ê/ë/ì/î/ó/ï/ñ/ò/ô/­/ø/õ/ö/÷/ù/ý/ú/û/ü/þ/®/¸/µ/¶/·/¹/¡/¾/»/¼/½/Æ/¢/Ê/Ç/È/É/Ë/Ð/Ì/Î/Ï/Ñ/£/Õ/Ò/Ó/Ô/Ö/Ý/×/Ø/Ü/Þ/ã/ß/á/â/ä/¤/è/å/æ/ç/é/¥/í/ê/ë/ì/î/ó/ï/ñ/ò/ô/¦/ø/õ/ö/÷/ù/ý/ú/û/ü/þ/§"
            .t_3DefaultFontName = ".VnTime"
            .t_4FontRecognizer = ".Vn"
            .t_5FontConversion = ".VnTime/.VnArial/.VnAvant/.VnBlack/.VnCourier"
            .t_6FontUpperCase = "*H"
        Case "VNI":
            .t_2VowelList = "aù/aø/aû/aõ/aï/aê/aé/aè/aú/aü/aë/aâ/aá/aà/aå/aã/aä/eù/eø/eû/eõ/eï/eâ/eá/eà/eå/eã/eä/í/ì/æ/ó/ò/où/oø/oû/oõ/oï/oâ/oá/oà/oå/oã/oä/ô/ôù/ôø/ôû/ôõ/ôï/uù/uø/uû/uõ/uï/ö/öù/öø/öû/öõ/öï/yù/yø/yû/yõ/î/ñ/AÙ/AØ/AÛ/AÕ/AÏ/AÊ/AÉ/AÈ/AÚ/AÜ/AË/AÂ/AÁ/AÀ/AÅ/AÃ/AÄ/EÙ/EØ/EÛ/EÕ/EÏ/EÂ/EÁ/EÀ/EÅ/EÃ/EÄ/Í/Ì/Æ/Ó/Ò/OÙ/OØ/OÛ/OÕ/OÏ/OÂ/OÁ/OÀ/OÅ/OÃ/OÄ/Ô/ÔÙ/ÔØ/ÔÛ/ÔÕ/ÔÏ/UÙ/UØ/UÛ/UÕ/UÏ/Ö/ÖÙ/ÖØ/ÖÛ/ÖÕ/ÖÏ/YÙ/YØ/YÛ/YÕ/Î/Ñ"
             .t_3DefaultFontName = "VNI-Times"
             .t_4FontRecognizer = "VNI"
             .t_5FontConversion = "VNI-Times/VNI-Helve/VNI-Helve/VNI-Helve/VNI-Couri"
             .t_6FontUpperCase = "*NONE"
        End Select
        iCode.t_1CodeName = iCodeName
    End With
    GetCodetable = iCode
End Function
