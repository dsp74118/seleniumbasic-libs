Public driver As New WebDriver
 
' Firefoxで某システムを開く
' - Firefox未起動だったら起動し、某システムのログイン画面を表示して自動ログイン
' - Firefox起動済みで、どこか他所のドメインに移動してたら、某システムのトップページに戻る
' - Firefox起動済みでログイン済みだったら何もしない
' - Firefox起動済みでセッションが切れてたら自動再ログイン
' @return 失敗したらfalse
Public Function OpenExampleCom() As Boolean
 
    Dim winCount As Integer
    Dim errNum As Integer
 
    OpenExampleCom = False
 
    On Error Resume Next
    winCount = driver.Windows.Count
    errNum = Err.Number
    On Error GoTo 0
    ' ToDo この辺りの処理は精査必要？
    If errNum = 57 Or winCount = 0 Then
        ' Firefoxを起動して、ページを開く
        driver.Start "firefox", "http://example.com/"
        driver.Get "/login"
        ' ログイン
        If AutoLogin = False then
            GoTo finally
        End If
    Else
        If driver.baseUrl <> "http://example.com" Then
            ' baseUrlが変わってたらセットし直し、トップページに移動。
            driver.baseUrl = "http://example.com/"
            driver.Get "/topPage"
        End If
        ' セッション判定
        ' ログアウトorタイムアウト画面ならログイン画面に移動。
        ' 対象のページに合わせて適切に書き換えること。
        If driver.FindElementByTag("h1").Text = "Logout" Or _
         driver.FindElementByTag("h1").Text = "TimeOut" Then
            driver.Get "/login"
        End If
        ' 今いるのがログイン画面ならログインする。
        If AutoLogin = False then
            GoTo Finally
        End If
    End If
 
    OpenExampleCom = True
 
finally:
    ' 事後処理。何かあればここに書く。
 
End Function
 
' 自動ログイン
' @return 失敗したらfalse
Private Function AutoLogin() as Boolean
 
    AutoLogin = True
 
    ' 今いるのがログイン画面ならログイン処理を行う。
    ' 対象のページに合わせて適切に書き換えること。
    If driver.FindElementByTag("h1").Text = "Login" Then
        driver.FindElementById("userid").SendKeys "dsp74118"
        driver.FindElementById("passwd").SendKeys "Password"
        driver.FindElementByTag("input").submit
        ' ログイン成否判定。対象のページに合わせて適切に書き換えること。
        If driver.FindElementByTag("h1").Text = "Login Error" Then
            MsgBox "ログインに失敗しました。"
            AutoLogin = False
        End If
    End If
 
End Function
