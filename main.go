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

	d := gomail.NewDialer("<HOST>", 25, "<USER>", "<PASSWORD>")
	if err := d.DialAndSend(m); err != nil {
		panic(err)
	}
}

func main() {
	// Check whether EXCEL has been running for more than 30 minutes
	cmd := exec.Command("powershell.exe", "/c", "Get-Process -Name \"EXCEL\" | Where-Object {$_.StartTime -le (Get-Date).AddMinutes(-30)}")
	out, _ := cmd.Output()

	if regexp.MustCompile(`EXCEL`).MatchString(string(out)) == true {
		from := "<sender>"
		to := "<destination>"
		cc := "<cc>"
		subject := "[WARNINGS] Process in progress"
		body := "There are EXCEL processes that remain running for more than 30 minutes.\n" + string(out)

		sendMail(from, to, cc, subject, body)
	}
}
