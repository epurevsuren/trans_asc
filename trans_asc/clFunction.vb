Imports System.IO

Public Class clFunction
    Public Function UpperCase(ByVal s As Object) As String
        Dim i As Integer
        Dim t As String = ""
        Dim st As String
        If Not (s Is Nothing Or s Is DBNull.Value) Then
            st = CStr(s)
            If st.Length > 0 Then
                For i = 0 To st.Length - 1
                    Select Case AscW(st(i))
                        Case 224 : t += ChrW(192) 'А
                        Case 225 : t += ChrW(193) 'Б
                        Case 226 : t += ChrW(194) 'В
                        Case 227 : t += ChrW(195) 'Г
                        Case 228 : t += ChrW(196) 'Д
                        Case 229 : t += ChrW(197) 'Е
                        Case 184 : t += ChrW(168) 'Ё
                        Case 230 : t += ChrW(198) 'Ж
                        Case 231 : t += ChrW(199) 'З
                        Case 232 : t += ChrW(200) 'И
                        Case 233 : t += ChrW(201) 'Й
                        Case 234 : t += ChrW(202) 'К
                        Case 235 : t += ChrW(203) 'Л
                        Case 236 : t += ChrW(204) 'М
                        Case 237 : t += ChrW(205) 'Н
                        Case 238 : t += ChrW(206) 'О
                        Case 186 : t += ChrW(170) 'Ө
                        Case 239 : t += ChrW(207) 'П
                        Case 240 : t += ChrW(208) 'Р
                        Case 241 : t += ChrW(209) 'С
                        Case 242 : t += ChrW(210) 'Т
                        Case 243 : t += ChrW(211) 'У
                        Case 191 : t += ChrW(175) 'Ү
                        Case 244 : t += ChrW(212) 'Ф
                        Case 245 : t += ChrW(213) 'Х
                        Case 246 : t += ChrW(214) 'Ц
                        Case 247 : t += ChrW(215) 'Ч
                        Case 248 : t += ChrW(216) 'Ш
                        Case 249 : t += ChrW(217) 'Щ
                        Case 250 : t += ChrW(218) 'Ъ
                        Case 251 : t += ChrW(219) 'Ы
                        Case 252 : t += ChrW(220) 'Ь
                        Case 253 : t += ChrW(221) 'Э
                        Case 254 : t += ChrW(222) 'Ю
                        Case 255 : t += ChrW(223) 'Я
                        Case Else : t += st(i)
                    End Select
                Next
            End If
        End If
        Return t
    End Function

    Shared Function Uni2Asc(ByVal s As Object) As String
        Dim i As Integer
        Dim t As String = ""
        Dim st As String
        If Not (s Is Nothing Or s Is DBNull.Value) Then
            st = CStr(s)
            If st.Length > 0 Then
                For i = 0 To st.Length - 1
                    Select Case AscW(st(i))
                        Case 1040 : t += ChrW(192) 'А
                        Case 1041 : t += ChrW(193) 'Б
                        Case 1042 : t += ChrW(194) 'В
                        Case 1043 : t += ChrW(195) 'Г
                        Case 1044 : t += ChrW(196) 'Д
                        Case 1045 : t += ChrW(197) 'Е
                        Case 1025 : t += ChrW(168) 'Ё
                        Case 1046 : t += ChrW(198) 'Ж
                        Case 1047 : t += ChrW(199) 'З
                        Case 1048 : t += ChrW(200) 'И
                        Case 1049 : t += ChrW(201) 'Й
                        Case 1050 : t += ChrW(202) 'К
                        Case 1051 : t += ChrW(203) 'Л
                        Case 1052 : t += ChrW(204) 'М
                        Case 1053 : t += ChrW(205) 'Н
                        Case 1054 : t += ChrW(206) 'О
                        Case 1256 : t += ChrW(170) 'Ө
                        Case 1055 : t += ChrW(207) 'П
                        Case 1056 : t += ChrW(208) 'Р
                        Case 1057 : t += ChrW(209) 'С
                        Case 1058 : t += ChrW(210) 'Т
                        Case 1059 : t += ChrW(211) 'У
                        Case 1198 : t += ChrW(175) 'Ү
                        Case 1060 : t += ChrW(212) 'Ф
                        Case 1061 : t += ChrW(213) 'Х
                        Case 1062 : t += ChrW(214) 'Ц
                        Case 1063 : t += ChrW(215) 'Ч
                        Case 1064 : t += ChrW(216) 'Ш
                        Case 1065 : t += ChrW(217) 'Щ
                        Case 1066 : t += ChrW(218) 'Ъ
                        Case 1067 : t += ChrW(219) 'Ы
                        Case 1068 : t += ChrW(220) 'Ь
                        Case 1069 : t += ChrW(221) 'Э
                        Case 1070 : t += ChrW(222) 'Ю
                        Case 1071 : t += ChrW(223) 'Я
                        Case 1072 : t += ChrW(224) 'а
                        Case 1073 : t += ChrW(225) 'б
                        Case 1074 : t += ChrW(226) 'в
                        Case 1075 : t += ChrW(227) 'г
                        Case 1076 : t += ChrW(228) 'д
                        Case 1077 : t += ChrW(229) 'е
                        Case 1105 : t += ChrW(184) 'ё
                        Case 1078 : t += ChrW(230) 'ж
                        Case 1079 : t += ChrW(231) 'з
                        Case 1080 : t += ChrW(232) 'и
                        Case 1081 : t += ChrW(233) 'й
                        Case 1082 : t += ChrW(234) 'к
                        Case 1083 : t += ChrW(235) 'л
                        Case 1084 : t += ChrW(236) 'м
                        Case 1085 : t += ChrW(237) 'н
                        Case 1086 : t += ChrW(238) 'о
                        Case 1257 : t += ChrW(186) 'ө
                        Case 1087 : t += ChrW(239) 'п
                        Case 1088 : t += ChrW(240) 'р
                        Case 1089 : t += ChrW(241) 'с
                        Case 1090 : t += ChrW(242) 'т
                        Case 1091 : t += ChrW(243) 'у
                        Case 1199 : t += ChrW(191) 'ү
                        Case 1092 : t += ChrW(244) 'ф
                        Case 1093 : t += ChrW(245) 'х
                        Case 1094 : t += ChrW(246) 'ц
                        Case 1095 : t += ChrW(247) 'ч
                        Case 1096 : t += ChrW(248) 'ш
                        Case 1097 : t += ChrW(249) 'щ
                        Case 1098 : t += ChrW(250) 'ъ
                        Case 1099 : t += ChrW(251) 'ы
                        Case 1100 : t += ChrW(252) 'ь
                        Case 1101 : t += ChrW(253) 'э
                        Case 1102 : t += ChrW(254) 'ю
                        Case 1103 : t += ChrW(255) 'я
                        Case Else : t += st(i)
                    End Select
                Next
            End If
        End If
        Return t
    End Function

    Shared Function Asc2Uni(ByVal s As Object) As String
        Dim i As Integer
        Dim t As String = ""
        Dim st As String
        If Not (s Is Nothing Or s Is DBNull.Value) Then
            st = CStr(s)
            If st.Length > 0 Then
                For i = 0 To st.Length - 1
                    Select Case AscW(st(i))
                        Case 192 : t += ChrW(1040) 'А
                        Case 193 : t += ChrW(1041) 'Б
                        Case 194 : t += ChrW(1042) 'В
                        Case 195 : t += ChrW(1043) 'Г
                        Case 196 : t += ChrW(1044) 'Д
                        Case 197 : t += ChrW(1045) 'Е
                        Case 168 : t += ChrW(1025) 'Ё
                        Case 198 : t += ChrW(1046) 'Ж
                        Case 199 : t += ChrW(1047) 'З
                        Case 200 : t += ChrW(1048) 'И
                        Case 201 : t += ChrW(1049) 'Й
                        Case 202 : t += ChrW(1050) 'К
                        Case 203 : t += ChrW(1051) 'Л
                        Case 204 : t += ChrW(1052) 'М
                        Case 205 : t += ChrW(1053) 'Н
                        Case 206 : t += ChrW(1054) 'О
                        Case 170 : t += ChrW(1256) 'Ө
                        Case 207 : t += ChrW(1055) 'П
                        Case 208 : t += ChrW(1056) 'Р
                        Case 209 : t += ChrW(1057) 'С
                        Case 210 : t += ChrW(1058) 'Т
                        Case 211 : t += ChrW(1059) 'У
                        Case 175 : t += ChrW(1198) 'Ү
                        Case 212 : t += ChrW(1060) 'Ф
                        Case 213 : t += ChrW(1061) 'Х
                        Case 214 : t += ChrW(1062) 'Ц
                        Case 215 : t += ChrW(1063) 'Ч
                        Case 216 : t += ChrW(1064) 'Ш
                        Case 217 : t += ChrW(1065) 'Щ
                        Case 218 : t += ChrW(1066) 'Ъ
                        Case 219 : t += ChrW(1067) 'Ы
                        Case 220 : t += ChrW(1068) 'Ь
                        Case 221 : t += ChrW(1069) 'Э
                        Case 222 : t += ChrW(1070) 'Ю
                        Case 223 : t += ChrW(1071) 'Я
                        Case 224 : t += ChrW(1072) 'а
                        Case 225 : t += ChrW(1073) 'б
                        Case 226 : t += ChrW(1074) 'в
                        Case 227 : t += ChrW(1075) 'г
                        Case 228 : t += ChrW(1076) 'д
                        Case 229 : t += ChrW(1077) 'е
                        Case 184 : t += ChrW(1105) 'ё
                        Case 230 : t += ChrW(1078) 'ж
                        Case 231 : t += ChrW(1079) 'з
                        Case 232 : t += ChrW(1080) 'и
                        Case 233 : t += ChrW(1081) 'й
                        Case 234 : t += ChrW(1082) 'к
                        Case 235 : t += ChrW(1083) 'л
                        Case 236 : t += ChrW(1084) 'м
                        Case 237 : t += ChrW(1085) 'н
                        Case 238 : t += ChrW(1086) 'о
                        Case 186 : t += ChrW(1257) 'ө
                        Case 239 : t += ChrW(1087) 'п
                        Case 240 : t += ChrW(1088) 'р
                        Case 241 : t += ChrW(1089) 'с
                        Case 242 : t += ChrW(1090) 'т
                        Case 243 : t += ChrW(1091) 'у
                        Case 191 : t += ChrW(1199) 'ү
                        Case 244 : t += ChrW(1092) 'ф
                        Case 245 : t += ChrW(1093) 'х
                        Case 246 : t += ChrW(1094) 'ц
                        Case 247 : t += ChrW(1095) 'ч
                        Case 248 : t += ChrW(1096) 'ш
                        Case 249 : t += ChrW(1097) 'щ
                        Case 250 : t += ChrW(1098) 'ъ
                        Case 251 : t += ChrW(1099) 'ы
                        Case 252 : t += ChrW(1100) 'ь
                        Case 253 : t += ChrW(1101) 'э
                        Case 254 : t += ChrW(1102) 'ю
                        Case 255 : t += ChrW(1103) 'я
                        Case Else : t += st(i)
                    End Select
                Next
            End If
        End If
        Return t
    End Function

    Function getStrKorpus(ByVal inp As Object) As String
        Dim ret As String = ""
        If Not (inp Is Nothing Or inp Is DBNull.Value) Then
            Select Case inp
                Case 0 : ret = "" 'No Korpus
                Case 1 : ret = "-" & ChrW(1040) 'А
                Case 2 : ret = "-" & ChrW(1041) 'Б
                Case 3 : ret = "-" & ChrW(1042) 'В
                Case 4 : ret = "-" & ChrW(1043) 'Г
                Case 5 : ret = "-" & ChrW(1044) 'Д
                Case 6 : ret = "-" & ChrW(1045) 'Е
                Case 7 : ret = "-" & ChrW(1025) 'Ё
                Case 8 : ret = "-" & ChrW(1046) 'Ж
                Case 9 : ret = "-" & ChrW(1047) 'З
                Case 10 : ret = "-" & ChrW(1048) 'И
                Case Else : ret = "-р байр, " & inp & "-р корпустай "
            End Select
        End If
        Return ret
    End Function

    Public Shared Sub WriteFileLog(ByVal str As String)
        'Try
        '    Dim strPath As String = Server.MapPath(".") & "\log\"
        '    Dim writer As New StreamWriter((strPath & "\" & DateTime.Now.ToString("yyyyMMdd") & ".log"), True)
        '    writer.WriteLine((DateTime.Now.ToString("HH:mm:ss.fff - ") & str))
        '    writer.Flush()
        '    writer.Close()
        'Catch ex As Exception
        'End Try
    End Sub
End Class
