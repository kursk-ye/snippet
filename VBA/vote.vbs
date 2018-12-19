dim ip1
dim ip2
dim ip3
dim ip4
dim ip
dim ShellCMD
dim objShell

dim h1
dim h2
dim h3
dim h4
dim h5
dim h6
dim h7
dim h8
dim h9
dim h10
dim h11
dim url
dim postData
dim fakeIpHeader

dim count
dim i

i = 0
count = 10

ip1 = 10
ip2 = 2
ip3 = 252
ip4 = 1

h1 = "-H" & " " & chr(34) & "x-requested-with: XMLHttpRequest" & chr(34) & " "
h2 = "-H" & " " & chr(34) &  "Accept-Language: zh-cn"& chr(34) & " "
h3 = "-H" & " " & chr(34) &  "Referer: http://sgtv.sgcc.com.cn/publish/gwgsh/"& chr(34) & " "
h4 = "-H" & " " & chr(34) &  "Accept: */*"& chr(34) & " "
h5 = "-H" & " " & chr(34) &  "Content-Type: application/x-www-form-urlencoded"& chr(34) & " "
h6 = "-H" & " " & chr(34) &  "Accept-Encoding: gzip, deflate"& chr(34) & " "
h7 = "-H" & " " & chr(34) &  "User-Agent: Mozilla/5.0 (compatible; MSIE 9.0; Windows NT 6.1; WOW64; Trident/5.0)"& chr(34) & " "
h8 = "-H" & " " & chr(34) &  "Host: sgtv.sgcc.com.cn " & chr(34) & " "
h9 = "-H" & " " & chr(34) &  "Content-Length: 8"& chr(34) & " "
h10 = "-H" & " " & chr(34) &  "Connection: Keep-Alive"& chr(34) & " "
h11 = "-H" & " " & chr(34) &  "Pragma: no-cache"& chr(34) & " "

url = "http://sgtv.sgcc.com.cn/plus/zan.php?dopost=send"
postData = " -d " & " " & chr(34) &  "aid=8408"& chr(34) & " "


do while ip1 <= 10

do while ip2 <= 2

do while ip3 <= 254

    do while ip4 <= 254
    	
    	 		
    	 ip = ip1 & "." & ip2 & "." & ip3 & "." & ip4
       fakeIpHeader = "-H" & " " & chr(34) &  "x-forwarded-for: " & ip & chr(34) & " "
       ShellCMD = "curl" & " " & h1 & h2 & h3 & h4 & h5 & h6 & h7 & h8 & h9 & h10 & h11 &  fakeIpHeader & " " & url & postData

       Set objShell = WScript.CreateObject("WScript.Shell")
       strRun =  ShellCMD  & " -o  D:\software\vote_output.log"
       SESSION_ID=objShell.Run(strRun)

       wscript.echo ip
       wscript.echo strRun

       'wscript.sleep 500
			 i = i + 1       


    ip4 = ip4 + 1
    loop

    ip4 = 1
    ip3 = ip3 + 1
loop


ip3 = 1
ip2 = ip2 + 1
loop

ip2 = 1
ip1 = ip1 + 1
loop


wscript.echo "good work!"



Function AddQuotes(strInput)
AddQuotes = Chr(34) & strInput & Chr(34)
End Function