## Send email functionality for Exchange in Golang

### Credits: [andlabs](https://github.com/andlabs/ews)

`Note: Currently only a single attachment is supported`

### Example usage

```
package main

import (
	"encoding/base64"
	"fmt"
	"io/ioutil"
	"log"

	"github.com/arawal/ews/ews"
)

func main() {
	fileData, err := ioutil.ReadFile("sample.txt")
	fileEncoding := base64.StdEncoding.EncodeToString(fileData)

	var amd ews.AttachmentMetadata
	amd.Name = "some.txt"
	amd.Content = fileEncoding

	var email ews.EmailMetadata
	email.To = []string{"1@some.org", "2@some.org"}
	email.Subject = "sample"
	email.Body = "sample email body"
	email.Cc = []string{"3@some.org", "4@some.org"}
	email.Bcc = []string{"5@some.org", "6@some.org"}
	email.ReplyTo = "7@some.org"

	var creds ews.Credentials
	creds.Server = "https://outlook.office365.com/EWS/exchange.asmx"
	creds.Username = "yourusername"
	creds.Password = "yourpassword"

	resp, err := ews.SendEmailWithAttachment(creds, email, amd)
	fmt.Println(resp, err)
	return

	fileData, err = ioutil.ReadFile("sample.txt")
	if err != nil {
		log.Fatal(err)
	}

	fileEncoding = base64.StdEncoding.EncodeToString(fileData)
	fmt.Printf("File contents: %s", fileEncoding)
}
```