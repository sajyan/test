Dim ie
Set ie = CreateObject("InternetExplorer.Application")
ie.Visible = True 
call ie.Navigate("https://www.google.com/")

'ページが読み込まれるまで待機
Do While ie.Busy = True Or ie.readyState <> 4
    WScript.Sleep 100        
Loop

Dim doc
Set doc = ie.Document
Dim txt
Set txt = doc.getElementsByName("q")
txt.item(0).value = "corona Japan"

Dim btn
Set btn = doc.getElementsByName("btnK")
btn.item(0).click()

'ページが読み込まれるまで待機
Do While ie.Busy = True Or ie.readyState <> 4
    WScript.Sleep 100        
Loop
Set doc = ie.Document

Dim list
Do While True
    Set list = doc.getElementsByClassName("LC20lb")

    If Not list is Nothing Then
        If list.length > 0 Then
            Exit Do
        End If
    End If
    WScript.Sleep 100 
Loop

Dim item
For Each item In list
    WScript.Echo item.innerText
Next
ie.Quit