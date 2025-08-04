<%
'######################################################################
'## easp.error.asp
'## -------------------------------------------------------------------
'## Feature     :   EasyASP Exception Class
'## Author      :   Coldstone(coldstone[at]qq.com)
'## Description :   Deal with the EasyASP Exception
'##
'######################################################################
Class EasyASP_Error
	Private b_redirect, b_continue, b_console
	Private i_errNum, i_delay, i_errLine, i_codeCache
	Private s_title, s_url, s_css, s_msg, s_funName
	Private o_err, a_detail, o_codeList
	Private e_err, e_conn, e_dom
	Private Sub Class_Initialize()
		i_errNum = ""
		i_delay = 5
		i_errLine = 0
		i_codeCache = 5
		s_title = Easp.Lang("error-title")
		b_redirect = False
		b_console = True
		b_continue = False
		s_url = "javascript:history.go(-1)"
		Set o_err = Server.CreateObject("Scripting.Dictionary")
		o_err.CompareMode = 1
	End Sub
	Private Sub Class_Terminate()
		If IsObject(o_codeList) Then Set o_codeList = Nothing
		Set o_err = Nothing
		If IsObject(e_err) Then Set e_err = Nothing
		If IsObject(e_conn) Then Set e_conn = Nothing
		If IsObject(e_dom) Then Set e_dom = Nothing
	End Sub
	'设置或读取自定义的错误代码和错误信息
	Public Default Property Get E(ByVal n)
		If IsNumeric(n) Then n = "E" & n
		If o_err.Exists(n) Then
			E = Join(o_err(n), "|")
		Else
			E = Easp.Lang("error-unkown")
		End If
	End Property
	Public Property Let E(ByVal n, ByVal s)
		Dim a_info, i_tmp
		If Easp.Has(n) And Easp.Has(s) Then
			If IsNumeric(n) Then n = "E" & n
			a_info = Split(s, "|")
			i_tmp = UBound(a_info)
			If i_tmp < 2 Then
				a_info = Split(s & String(2 - i_tmp, "|"), "|")
			End If
			o_err(n) = a_info
		End If
	End Property
	'取最后一次发生错误的代码
	Public Property Get LastError
		LastError = i_errNum
	End Property
	'设置和读取错误信息标题
	Public Property Get Title
		Title = s_title
	End Property
	Public Property Let Title(ByVal s)
		s_title = s
	End Property
	'设置显示错误信息时的详细信息替换参数
	Public Property Let Detail(ByVal arr)
		a_detail = arr
	End Property
	'设置和读取出错函数名
	Public Property Get FunctionName
		FunctionName = s_funName
	End Property
	Public Property Let FunctionName(ByVal String)
		s_funName = String
	End Property
	'设置和读取页面是否自动转向
	Public Property Get [Redirect]
		[Redirect] = b_redirect
	End Property
	Public Property Let [Redirect](ByVal b)
		b_redirect = b
	End Property
	'设置和读取Debug模式下出错后是否继续运行后面的代码
	'说明：普通模式下总是继续运行
	Public Property Get OnErrorContinue
		OnErrorContinue = b_continue
	End Property
	Public Property Let OnErrorContinue(ByVal bool)
		b_continue = bool
	End Property
	'设置和读取是否在控制台内显示详细错误信息
	Public Property Get ConsoleDetail
		ConsoleDetail = b_console
	End Property
	Public Property Let ConsoleDetail(ByVal bool)
		b_console = bool
	End Property
	'设置和读取发生错误后的跳转页地址
	'说明：如不设置此属性，则默认为返回前一页
	Public Property Get Url
		Url = s_url
	End Property
	Public Property Let Url(ByVal s)
		s_url = s
	End Property
	'设置和读取自动跳转页面等待时间（秒）
	Public Property Get Delay
		Delay = i_delay
	End Property
	Public Property Let Delay(ByVal i)
		i_delay = i
	End Property
	'设置和读取显示错误信息DIV的CSS样式名称
	Public Property Get ClassName
		ClassName = s_css
	End Property
	Public Property Let ClassName(ByVal s)
		s_css = s
	End Property
	'设置错误行号
	Public Property Let LineNumber(ByVal i)
		i_errLine = i
	End Property
	'设置要保存最后执行的语句的数量
	Public Property Let LastCodeCache(ByVal i)
		i_codeCache = i
	End Property
	
	'Dom和Connection错误
	Public Sub SetErrors(ByRef e, ByRef ec, ByRef ed)
		If IsObject(e) Then Set e_err = e
		If IsObject(ec) Then Set e_conn = ec
		If IsObject(ed) Then Set e_dom = ed
	End Sub
	
	'生成一个错误(常用于开发者错误模式)
	Public Sub Raise(ByVal n)
		If Easp.isN(n) Then Exit Sub
		If IsNumeric(n) Then n = "E" & n
		If Not IsObject(e_err) Then Set e_err = Err
		Dim b_consoleDetail, b_isEnd
		Dim msg
		'得到已定义错误信息
		msg = o_err(n)
		'如果是Debug模式，出错后是否继续运行
		b_isEnd = Easp.IIF(Easp.Debug , Not b_continue, False)
		'在控制台内输出错误信息
		InConsole msg, b_console
		If b_isEnd Then
			Easp.PrintEnd ShowErrorMsg(msg)
		Else
			Easp.Print ShowErrorMsg(msg)
		End If
		i_errNum = n
		s_msg = ""
	End Sub
	
	'立即抛出一个错误信息(常用于用户错误模式)
	Public Sub Throw(ByVal msg)
		Dim a_info, i_tmp
		If Easp.Has(msg) Then
			a_info = Split(msg, "|")
			i_tmp = UBound(a_info)
			If i_tmp < 2 Then
				a_info = Split(msg & String(2 - i_tmp, "|"), "|")
			End If
			Easp.Print ShowErrorMsg(a_info)
		End If
	End Sub
	
	Public Sub Inject(ByRef object_err, ByVal string_filePath, ByVal int_lineNumber, ByVal string_sourceCode)
		If Easp.Debug Then
			If Not IsObject(o_codeList) Then Set o_codeList = Easp.Json.NewArray
			Easp.Error.SetErrors object_err, Null, Null
			Easp.Error.LineNumber = int_lineNumber
			o_codeList.Add Array(string_sourceCode, string_filePath, int_lineNumber)
			If object_err.Number <> 0 Then
				Easp.Error.Throw "Microsoft VBScript 运行时错误"
				object_err.Clear
				i_errLine = 0
				If Not b_continue Then Easp.Exit
			End If
		End If
	End Sub
	
	'在控制台中抛出错误信息
	Public Sub Console(ByVal n)
	End Sub
	
	'控制台输出错误：
	Private Sub InConsole(ByVal msg, ByVal hasDetail)
	End Sub
	
	'显示错误信息框
	Private Function ShowErrorMsg(ByVal msg)
		Dim SB, key, s_ref, i, lines
		s_ref = Request.ServerVariables("HTTP_REFERER")
		Set SB = Easp.Str.StringBuilder()
		If Easp.IsN(s_css) Then
			s_css = "ea-error"
			SB.Append "<style>body{background-color:#f3f5f9;font-size:12px;color:#8B91A0;padding:20px;font-family:consolas,""Microsoft Yahei"";line-height:1}h1,h2,h3,h4,h5,h6{font-weight:100}.easp-error{border-radius:6px;box-shadow:3px 3px 3px 3px rgba(0,0,0,.08);-webkit-box-shadow:3px 3px 3px 3px rgba(0,0,0,.08);background-color:#fff;padding:2px 10px 20px 10px}.easp-error h2{color:#f79d3c;margin:30px 0}.easp-error h3{margin:20px 0;color:#b40605}.easp-error a{color:#b40605;text-decoration:none}.easp-error a:link{color:#1c84c6}.easp-error a:hover{color:#1ab394}.easp-error hr{margin:20px 0;border:0;border-top:1px solid #eee}</style>"
		End If
		SB.Append "<div id=""e-Error"" class="""
		SB.Append s_css
		SB.Append """>"
		SB.Append "<h3>"
		SB.Append s_title
		SB.Append "</h3><hr>"
		SB.Append "<h2><img src='/Images/404.gif' /></h2>"
		If Easp.Debug Then
			'显示详细错误信息
		Else
			'显示普通模式详细错误信息
		End If
		SB.Append "<br>"
		If Easp.Str.IsSame(s_ref, Easp.GetUrl("")) Or Easp.IsN(s_ref) Then
			b_redirect = False
			s_url = "javascript:location.reload(true)"
		End If
		If b_redirect Then
			SB.Append "页面将在<span id=""easp_timeoff"">"
			SB.Append i_delay
			SB.Append "</span>秒钟后跳转，如果浏览器没有正常跳转，"
		End If
		SB.Append "<a href="""
		SB.Append s_url
		SB.Append """>请点击此处"
		If Easp.Str.IsSame(s_url, "javascript:history.go(-1)") Then
			SB.Append "返回"
		ElseIf Easp.Str.IsSame(s_url, "javascript:location.reload(true)") Then
			SB.Append "刷新"
		Else
			SB.Append "继续"
		End If
		SB.Append "</a><br></div>"
		SB.Append "<script type=""text/javascript"">function toggle(id){var el = document.getElementById(id);var a = document.getElementById(id+""_m"");if(a.innerHTML==""[-]""){el.style.display = ""none"";a.innerHTML = ""[+]"";}else if(a.innerHTML==""[+]""){el.style.display = """";a.innerHTML = ""[-]"";}}"
		If b_redirect Then
			SB.Append "function timeMinus(){var el = document.getElementById(""easp_timeoff"");var timeLeft = parseInt(el.innerHTML);el.innerHTML = timeLeft - 1;} setInterval(timeMinus, 1000);"
			SB.Append "setTimeout(function(){"
			If Easp.Str.IsSame(Left(s_url,11), "javascript:") Then
				SB.Append Mid(s_url, 12)
			Else
				SB.Append "location.href='"
				SB.Append s_url
				SB.Append "'"
			End If
			SB.Append "},"
			SB.Append i_delay * 1000
			SB.Append ");"
		End If
		SB.Append "</script>"
		ShowErrorMsg = SB.ToString()
		Set SB = Nothing
	End Function
	
	'显示已定义的所有错误代码及信息，返回Json格式
	Public Function Defined()
		Defined = Easp.Str.ToString(o_err)
	End Function
End Class
%>