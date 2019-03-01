package ews

import (
	"encoding/xml"
)

type EmailMetadata struct {
	Action  string // SaveOnly or SendAndSaveCopy
	To      []string
	Cc      []string
	Bcc     []string
	ReplyTo string
	Subject string
	Body    string
	Type    string // Text or HTML
	Folder  string
}

type AttachmentMetadata struct {
	Name           string
	Content        string
	EmailID        string
	EmailChangeKey string
}

// https://msdn.microsoft.com/en-us/library/office/aa563009(v=exchg.140).aspx

type CreateItem struct {
	XMLName            struct{}          `xml:"m:CreateItem"`
	MessageDisposition string            `xml:"MessageDisposition,attr"`
	SavedItemFolderID  SavedItemFolderID `xml:"m:SavedItemFolderId"`
	Items              Messages          `xml:"m:Items"`
}

type Messages struct {
	Message []Message `xml:"t:Message"`
}

type SavedItemFolderID struct {
	DistinguishedFolderID DistinguishedFolderID `xml:"t:DistinguishedFolderId"`
}

type DistinguishedFolderID struct {
	ID string `xml:"Id,attr"`
}

type Message struct {
	ItemClass   string       `xml:"t:ItemClass"`
	Subject     string       `xml:"t:Subject"`
	Body        Body         `xml:"t:Body"`
	Attachments []Attachment `xml:"t:Attachments"`
	// Sender       OneMailbox   `xml:"t:Sender"`
	ToRecipients  XMailbox   `xml:"t:ToRecipients"`
	CcRecipients  XMailbox   `xml:"t:CcRecipients"`
	BccRecipients XMailbox   `xml:"t:BccRecipients"`
	ReplyTo       OneMailbox `xml:"t:ReplyTo"`
}

type Body struct {
	BodyType string `xml:"BodyType,attr"`
	Body     []byte `xml:",chardata"`
}

type OneMailbox struct {
	Mailbox Mailbox `xml:"t:Mailbox"`
}

type XMailbox struct {
	Mailbox []Mailbox `xml:"t:Mailbox"`
}

type Mailbox struct {
	EmailAddress string `xml:"t:EmailAddress"`
}

type FileAttachment struct {
	Name           string `xml:"t:Name"`
	IsInline       bool   `xml:"t:IsInline"`
	IsContactPhoto bool   `xml:"t:IsContactPhoto"`
	Content        string `xml:"t:Content"`
}

type Attachment struct {
	FileAttachment FileAttachment `xml:"t:FileAttachment"`
}

type ItemID struct {
	ID        string `xml:"Id,attr"`
	ChangeKey string `xml:"ChangeKey,attr"`
}

type ItemIDs struct {
	ItemID []ItemID `xml:"t:ItemId"`
}

type EmailAttachment struct {
	XMLName      xml.Name     `xml:"CreateAttachment"`
	Xmlns        string       `xml:"xmlns,attr"`
	XmlnsT       string       `xml:"xmlns:t,attr"`
	ParentItemID ItemID       `xml:"ParentItemId"`
	Attachments  []Attachment `xml:"Attachments"`
}

type SendItem struct {
	XMLName           xml.Name `xml:"SendItem"`
	Xmlns             string   `xml:"xmlns,attr"`
	SaveItemToFolder  bool     `xml:"SaveItemToFolder,attr"`
	ItemIds           ItemIDs  `xml:"ItemIds"`
	SavedItemFolderID string   `xml:"t:SavedItemFolderId"`
}

func BuildTextEmail(metadata EmailMetadata) ([]byte, error) {
	email := CreateItemXML(metadata)
	return xml.MarshalIndent(email, "", "  ")
}

func BuildAttachment(metadata AttachmentMetadata) ([]byte, error) {
	var emailAttachment EmailAttachment
	emailAttachment.Xmlns = "http://schemas.microsoft.com/exchange/services/2006/messages"
	emailAttachment.XmlnsT = "http://schemas.microsoft.com/exchange/services/2006/types"
	emailAttachment.ParentItemID.ID = metadata.EmailID
	emailAttachment.ParentItemID.ChangeKey = metadata.EmailChangeKey

	attachment := Attachment{FileAttachment{
		Name:    metadata.Name,
		Content: metadata.Content,
	}}
	emailAttachment.Attachments = []Attachment{attachment}

	return xml.MarshalIndent(emailAttachment, "", "  ")
}

func BuildSendSavedEmail(emailId, changeKey string) ([]byte, error) {
	var senditem SendItem
	senditem.Xmlns = "http://schemas.microsoft.com/exchange/services/2006/messages"
	senditem.SaveItemToFolder = true
	senditem.SavedItemFolderID = "sentitems"

	senditem.ItemIds = ItemIDs{
		ItemID: []ItemID{
			ItemID{
				ID:        emailId,
				ChangeKey: changeKey,
			},
		},
	}

	return xml.MarshalIndent(senditem, "", "  ")
}

func CreateItemXML(metadata EmailMetadata) CreateItem {
	c := new(CreateItem)
	c.MessageDisposition = metadata.Action
	c.SavedItemFolderID.DistinguishedFolderID.ID = metadata.Folder
	m := new(Message)
	m.ItemClass = "IPM.Note"
	m.Subject = metadata.Subject
	m.Body.BodyType = "Text"
	m.Body.Body = []byte(metadata.Body)

	// adding a sender requires "SendAs" on the account sending the message, even if it sending email from itself
	// m.Sender.Mailbox.EmailAddress = metadata.From

	mb := make([]Mailbox, len(metadata.To))
	for i, addr := range metadata.To {
		mb[i].EmailAddress = addr
	}
	m.ToRecipients.Mailbox = append(m.ToRecipients.Mailbox, mb...)

	mb = make([]Mailbox, len(metadata.Cc))
	for i, addr := range metadata.Cc {
		mb[i].EmailAddress = addr
	}
	m.CcRecipients.Mailbox = append(m.CcRecipients.Mailbox, mb...)

	mb = make([]Mailbox, len(metadata.Bcc))
	for i, addr := range metadata.Bcc {
		mb[i].EmailAddress = addr
	}
	m.BccRecipients.Mailbox = append(m.BccRecipients.Mailbox, mb...)

	m.ReplyTo.Mailbox.EmailAddress = metadata.ReplyTo
	c.Items.Message = append(c.Items.Message, *m)

	return *c
}
