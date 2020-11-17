package main

import (
	"os/exec"
	"regexp"

	gomail "gopkg.in/gomail.v2"
)

func sendMail(from, to, cc, subject, body string) {
	m := gomail.NewMessage()
	m.SetHeader("From", from)
	m.SetHeader("To", to)
	m.SetHeader("Cc", cc)
	m.SetHeader("Subject", subject)
	m.SetBody("text/plain", body)

	d := gomail.NewDialer("<HOST>", 25, "<USER>", "<PASSWORD>") // ホスト名, ポート, ユーザー, パスワード
	if err := d.DialAndSend(m); err != nil {
		panic(err)
	}
}

func main() {
	// EXCELが30分以上起動しているかどうかをチェックする
	cmd := exec.Command("powershell.exe", "/c", "Get-Process -Name \"EXCEL\" | Where-Object {$_.StartTime -le (Get-Date).AddMinutes(-30)}")
	out, _ := cmd.Output()

	// outがtrueの場合
	if regexp.MustCompile(`EXCEL`).MatchString(string(out)) == true {
		// Mail通知
		from := "<sender>"
		to := "<destination>"
		cc := "<cc>"
		subject := "起動中プロセス"
		body := "30分以上起動したままのEXCELプロセスがあります。\n" + string(out)

		sendMail(from, to, cc, subject, body)
	}
}
