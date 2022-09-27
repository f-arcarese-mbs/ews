package ews

import (
	"bytes"
	"encoding/json"
	"encoding/xml"
	"errors"
	"fmt"
	"io"
	"io/ioutil"
	"net/http"
	"net/url"
	"strconv"
	"strings"
	"time"
)

// https://msdn.microsoft.com/en-us/library/office/dd877045(v=exchg.140).aspx
// https://arvinddangra.wordpress.com/2011/09/29/send-email-using-exchange-smtp-and-ews-exchange-web-service/
// https://msdn.microsoft.com/en-us/library/office/dn789003(v=exchg.150).aspx

type Credentials struct {
	Server       string
	Username     string
	Password     string
	Tenant       string
	ClientID     string
	ClientSecret string
	GrantType    string
	Scope        string
	URL          string
}

type OAuthResponse struct {
	TokenType    string `json:"token_type"`
	ExpiresIn    int    `json:"expires_is"`
	ExtExpiresIn int    `json:"ext_expires_in"`
	AccessToken  string `json:"access_token"`
}

type CreateItemResponse struct {
	XMLName xml.Name `xml:"Envelope"`
	Body    struct {
		CreateItemResponseMessage struct {
			ResponseClass string `xml:"ResponseClass,attr"`
			ResponseCode  string `xml:"ResponseCode"`
			Items         struct {
				Message []struct {
					ItemID struct {
						ID        string `xml:"Id,attr"`
						ChangeKey string `xml:"ChangeKey,attr"`
					} `xml:"ItemId"`
				} `xml:"Message"`
			} `xml:"Items"`
		} `xml:"CreateItemResponse>ResponseMessages>CreateItemResponseMessage"`
	} `xml:"Body"`
}

type CreateAttachmentResponse struct {
	XMLName xml.Name `xml:"Envelope"`
	Body    struct {
		CreateAttachmentResponseMessage struct {
			ResponseClass string `xml:"ResponseClass,attr"`
			ResponseCode  string `xml:"ResponseCode"`
			Attachments   struct {
				FileAttachment []struct {
					AttachmentID struct {
						ID                string `xml:"Id,attr"`
						RootItemID        string `xml:"RootItemId,attr"`
						RootItemChangeKey string `xml:"RootItemChangeKey,attr"`
					} `xml:"AttachmentId"`
					LastModifiedTime string `xml:"LastModifiedTime"`
				} `xml:"FileAttachment"`
			} `xml:"Attachments"`
		} `xml:"CreateAttachmentResponse>ResponseMessages>CreateAttachmentResponseMessage"`
	} `xml:"Body"`
}

type SendItemResponse struct {
	XMLName xml.Name `xml:"Envelope"`
	Body    struct {
		SendItemResponse struct {
			ResponseMessages struct {
				SendItemResponseMessage struct {
					ResponseClass string `xml:"ResponseClass,attr"`
					ResponseCode  string `xml:"ResponseCode"`
				} `xml:"SendItemResponseMessage"`
			} `xml:"ResponseMessages"`
		} `xml:"SendItemResponse"`
	} `xml:"Body"`
}

func SendEmailWithAttachment(creds Credentials, emailMetadata EmailMetadata, attachmentMetadata AttachmentMetadata) (string, error) {
	emailMetadata.Action = "SaveOnly"
	emailMetadata.Folder = "sentitems"

	var err error
	_, attachmentMetadata.EmailID, attachmentMetadata.EmailChangeKey, err = IssueTextEmail(creds, emailMetadata)
	if err != nil {
		return "", err
	}

	attachmentMetadata.EmailChangeKey, err = IssueAttachment(creds, attachmentMetadata)
	if err != nil {
		return "", err
	}

	return IssueEmailWithAttachment(creds, attachmentMetadata)
}

func SendEmail(creds Credentials, metadata EmailMetadata) (string, error) {
	metadata.Action = "SendAndSaveCopy"
	if metadata.Folder == "" {
		metadata.Folder = "sentitems"
	}
	status, _, _, err := IssueTextEmail(creds, metadata)
	return status, err
}

func IssueEmailWithAttachment(creds Credentials, metadata AttachmentMetadata) (string, error) {
	sendXML, err := BuildSendSavedEmail(metadata.EmailID, metadata.EmailChangeKey)
	if err != nil {
		return "", err
	}

	resp, err := Issue(creds, sendXML)
	if err != nil {
		panic(err.Error())
	}

	var sendResp SendItemResponse
	err = xml.Unmarshal([]byte(resp), &sendResp)
	if err != nil {
		return "", err
	}

	if sendResp.Body.SendItemResponse.ResponseMessages.SendItemResponseMessage.ResponseClass != "Success" {
		return "", errors.New(sendResp.Body.SendItemResponse.ResponseMessages.SendItemResponseMessage.ResponseCode)
	}

	return sendResp.Body.SendItemResponse.ResponseMessages.SendItemResponseMessage.ResponseClass, nil
}

func IssueAttachment(creds Credentials, metadata AttachmentMetadata) (string, error) {
	attachmentXML, err := BuildAttachment(metadata)
	if err != nil {
		return "", err
	}

	resp, err := Issue(creds, attachmentXML)
	if err != nil {
		panic(err.Error())
	}

	var createAttachmentResponse CreateAttachmentResponse
	err = xml.Unmarshal([]byte(resp), &createAttachmentResponse)
	if err != nil {
		return "", err
	}

	return createAttachmentResponse.Body.CreateAttachmentResponseMessage.Attachments.FileAttachment[0].AttachmentID.RootItemChangeKey, nil
}

func IssueTextEmail(creds Credentials, metadata EmailMetadata) (string, string, string, error) {
	emailXML, err := BuildTextEmail(metadata)

	resp, err := Issue(creds, emailXML)
	if err != nil {
		return "", "", "", err
	}

	var createResponse CreateItemResponse
	err = xml.Unmarshal([]byte(resp), &createResponse)
	if err != nil {
		return "", "", "", err
	}

	if createResponse.Body.CreateItemResponseMessage.ResponseClass != "Success" {
		return "", "", "", errors.New(createResponse.Body.CreateItemResponseMessage.ResponseCode)
	}

	if metadata.Action == "SendAndSaveCopy" {
		return createResponse.Body.CreateItemResponseMessage.ResponseClass, "", "", nil
	}

	return createResponse.Body.CreateItemResponseMessage.ResponseClass, createResponse.Body.CreateItemResponseMessage.Items.Message[0].ItemID.ID, createResponse.Body.CreateItemResponseMessage.Items.Message[0].ItemID.ChangeKey, nil
}

func Issue(creds Credentials, body []byte) (string, error) {
	soapheader := `<?xml version="1.0" encoding="utf-8" ?>
<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages" xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
  <soap:Header>
    <t:RequestServerVersion Version="Exchange2007_SP1" />
   </soap:Header>
   <soap:Body>
`

	if creds.ClientID != "" {
		soapheader = `<?xml version="1.0" encoding="utf-8" ?>
<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages" xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
  <soap:Header>
	<t:RequestServerVersion Version="Exchange2007_SP1" />
	<t:ExchangeImpersonation>
	<t:ConnectingSID>
	  <t:PrimarySmtpAddress>` + creds.Username + `</t:PrimarySmtpAddress>
	</t:ConnectingSID>
  </t:ExchangeImpersonation>
  </soap:Header>
  <soap:Body>
`
	}

	b2 := []byte(soapheader)
	b2 = append(b2, body...)
	b2 = append(b2, "\n  </soap:Body>\n</soap:Envelope>"...)
	req, err := http.NewRequest("POST", creds.Server, bytes.NewReader(b2))
	if err != nil {
		return "", err
	}

	if creds.ClientID != "" {
		oauthResponse, err := requestOAuthToken(creds)

		if err != nil {
			fmt.Println(err.Error())
			return "", err
		}

		if oauthResponse.AccessToken == "" || oauthResponse.TokenType == "" {
			err := errors.New("No access token")
			fmt.Println(err.Error())
			return "", err
		}

		req.Header.Set("Authorization", oauthResponse.TokenType+" "+oauthResponse.AccessToken)
		req.Header.Set("X-PreferServerAffinity", "true")
		req.Header.Set("X-AnchorMailbox", creds.Username)
	} else {
		req.SetBasicAuth(creds.Username, creds.Password)
	}

	req.Header.Set("Content-Type", "text/xml")
	client := new(http.Client)
	client.CheckRedirect = func(req *http.Request, via []*http.Request) error { return http.ErrUseLastResponse }
	resp, err := client.Do(req)

	if err != nil {
		fmt.Println(err.Error())
		return "", err
	}
	defer resp.Body.Close()
	if resp.StatusCode == http.StatusOK {
		bodyBytes, err := ioutil.ReadAll(resp.Body)
		if err != nil {
			return "", err
		}
		bodyString := string(bodyBytes)
		return bodyString, nil
	}

	bodyBytes, err := ioutil.ReadAll(resp.Body)
	if err != nil {
		return "", err
	}
	bodyString := string(bodyBytes)
	fmt.Println(bodyString)
	return bodyString, errors.New(resp.Status)
}

func requestOAuthToken(creds Credentials) (OAuthResponse, error) {
	ret := OAuthResponse{}

	client := &http.Client{
		Timeout: time.Second * 30,
	}

	data := url.Values{}
	data.Set("grant_type", creds.GrantType)
	data.Set("client_id", creds.ClientID)
	data.Set("client_secret", creds.ClientSecret)
	data.Set("scope", creds.Scope)

	encodedData := data.Encode()
	req, err := http.NewRequest("GET", fmt.Sprintf(creds.URL, creds.Tenant), strings.NewReader(encodedData))

	if err != nil {
		fmt.Printf("Error %s", err.Error())
		return ret, err
	}

	req.Header.Add("Content-Type", "application/x-www-form-urlencoded")
	req.Header.Add("Content-Length", strconv.Itoa(len(data.Encode())))

	response, err := client.Do(req)
	if err != nil {
		fmt.Printf("Error %s", err.Error())
		return ret, err
	}

	defer response.Body.Close()

	b, err := io.ReadAll(response.Body)
	if err != nil {
		fmt.Printf("Error s%s", err.Error())
		return ret, err
	}

	err = json.Unmarshal(b, &ret)
	if err != nil {
		fmt.Printf("Error %s", err.Error())
		return ret, err
	}

	return ret, nil
}
