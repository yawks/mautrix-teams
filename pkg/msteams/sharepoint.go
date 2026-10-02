package msteams

import (
	"bytes"
	"context"
	"encoding/json"
	"fmt"
	"io"
	"net/http"
	"net/url"
	"path"
	"strings"
)

// UploadSharedFile uploads a document next to an existing Teams chat file and
// grants the chat participants read access using SharePoint's REST API. This
// uses the delegated SharePoint token minted from the Teams session; it does
// not require a customer-created Microsoft Graph application.
func (c *Client) UploadSharedFile(ctx context.Context, location SharedFile, name string, data []byte, recipients []string) (*SharedFile, error) {
	if location.SiteURL == "" || location.FileURL == "" {
		return nil, fmt.Errorf("SharePoint upload location is unavailable")
	}
	site, err := url.Parse(location.SiteURL)
	if err != nil || site.Host == "" {
		return nil, fmt.Errorf("parse SharePoint site URL: %w", err)
	}
	fileURL, err := url.Parse(location.FileURL)
	if err != nil || fileURL.Host == "" {
		return nil, fmt.Errorf("parse SharePoint file URL: %w", err)
	}
	name = path.Base(strings.TrimSpace(name))
	if name == "" || name == "." {
		return nil, fmt.Errorf("empty file name")
	}
	token, err := c.freshSharePointToken(ctx, site.Host)
	if err != nil {
		return nil, err
	}
	digest, err := c.sharePointDigest(ctx, strings.TrimRight(location.SiteURL, "/"), token)
	if err != nil {
		return nil, err
	}
	folder := path.Dir(fileURL.Path)
	type uploadResponse struct {
		D struct {
			UniqueID          string `json:"UniqueId"`
			ListID            string `json:"ListId"`
			ServerRelativeURL string `json:"ServerRelativeUrl"`
			Name              string `json:"Name"`
		} `json:"d"`
		UniqueID          string `json:"UniqueId"`
		ListID            string `json:"ListId"`
		ServerRelativeURL string `json:"ServerRelativeUrl"`
		Name              string `json:"Name"`
	}
	var uploaded uploadResponse
	actualName := name
	for suffix := 0; suffix <= 100; suffix++ {
		actualName = numberedFileName(name, suffix)
		endpoint := strings.TrimRight(location.SiteURL, "/") + "/_api/web/GetFolderByServerRelativeUrl('" +
			url.PathEscape(strings.ReplaceAll(folder, "'", "''")) + "')/Files/Add(url='" +
			url.PathEscape(strings.ReplaceAll(actualName, "'", "''")) + "',overwrite=false)?$select=ListId,UniqueId,ServerRelativeUrl,Name"
		uploaded = uploadResponse{}
		uploadErr := c.sharePointJSON(ctx, http.MethodPost, endpoint, token, digest, data, "application/octet-stream", &uploaded)
		if uploadErr == nil {
			break
		}
		if !sharePointFileAlreadyExists(uploadErr) {
			return nil, fmt.Errorf("upload SharePoint file: %w", uploadErr)
		}
		if suffix == 100 {
			return nil, fmt.Errorf("upload SharePoint file: no available name found for %s", name)
		}
	}
	itemID := firstNonEmpty(uploaded.UniqueID, uploaded.D.UniqueID)
	listID := firstNonEmpty(uploaded.ListID, uploaded.D.ListID)
	serverPath := firstNonEmpty(uploaded.ServerRelativeURL, uploaded.D.ServerRelativeURL)
	actualName = firstNonEmpty(uploaded.Name, uploaded.D.Name, actualName)
	if serverPath == "" {
		serverPath = path.Join(folder, actualName)
	}
	uploadedURL := site.Scheme + "://" + site.Host + serverPath
	if _, err := c.shareSharePointFile(ctx, location.SiteURL, uploadedURL, token, digest, recipients); err != nil {
		return nil, err
	}
	shareURL, shareID, err := c.createSharePointLink(ctx, location.SiteURL, listID, itemID, token, digest)
	if err != nil {
		return nil, err
	}
	siteID, err := c.sharePointSiteID(ctx, location.SiteURL, token)
	if err != nil {
		return nil, err
	}
	return &SharedFile{Name: actualName, ItemID: strings.Trim(itemID, "{}"), SiteURL: location.SiteURL, SiteID: siteID,
		FileURL: uploadedURL, ShareURL: shareURL, ShareID: shareID, Size: int64(len(data))}, nil
}

func (c *Client) sharePointSiteID(ctx context.Context, siteURL, token string) (string, error) {
	var result struct {
		ID string `json:"Id"`
		D  struct {
			ID string `json:"Id"`
		} `json:"d"`
	}
	if err := c.sharePointJSON(ctx, http.MethodGet, strings.TrimRight(siteURL, "/")+"/_api/site?$select=Id",
		token, "", nil, "application/json", &result); err != nil {
		return "", fmt.Errorf("get SharePoint site ID: %w", err)
	}
	id := strings.Trim(firstNonEmpty(result.ID, result.D.ID), "{}")
	if id == "" {
		return "", fmt.Errorf("SharePoint returned no site ID")
	}
	return id, nil
}

func (c *Client) createSharePointLink(ctx context.Context, siteURL, listID, itemID, token, digest string) (string, string, error) {
	if listID == "" || itemID == "" {
		return "", "", fmt.Errorf("SharePoint upload returned no list or item ID")
	}
	body, _ := json.Marshal(map[string]any{"request": map[string]any{
		"createLink": true,
		"settings": map[string]any{
			"allowAnonymousAccess": false, "linkKind": 2, "nav": "", "password": "",
			"restrictShareMembership": false, "role": 1, "scope": 1, "updatePassword": false,
		},
	}})
	endpoint := strings.TrimRight(siteURL, "/") + "/_api/web/Lists(@list)/GetItemByUniqueId(@item)/ShareLink?@list='" +
		url.QueryEscape(listID) + "'&@item='" + url.QueryEscape(strings.Trim(itemID, "{}")) + "'"
	type sharingLinkInfo struct {
		URL     any `json:"Url"`
		ShareID any `json:"ShareId"`
	}
	var result struct {
		SharingLinkInfo sharingLinkInfo `json:"sharingLinkInfo"`
		D               struct {
			SharingLinkInfo sharingLinkInfo `json:"sharingLinkInfo"`
			ShareLink       struct {
				SharingLinkInfo sharingLinkInfo `json:"sharingLinkInfo"`
			} `json:"ShareLink"`
		} `json:"d"`
	}
	if err := c.sharePointJSON(ctx, http.MethodPost, endpoint, token, digest, body, "application/json;odata=verbose", &result); err != nil {
		return "", "", fmt.Errorf("create SharePoint link: %w", err)
	}
	shareURL := firstNonEmpty(
		sharePointJSONText(result.SharingLinkInfo.URL),
		sharePointJSONText(result.D.SharingLinkInfo.URL),
		sharePointJSONText(result.D.ShareLink.SharingLinkInfo.URL),
	)
	if shareURL == "" {
		return "", "", fmt.Errorf("SharePoint returned no sharing link")
	}
	return shareURL, firstNonEmpty(
		sharePointJSONText(result.SharingLinkInfo.ShareID),
		sharePointJSONText(result.D.SharingLinkInfo.ShareID),
		sharePointJSONText(result.D.ShareLink.SharingLinkInfo.ShareID),
	), nil
}

func sharePointJSONText(value any) string {
	switch typed := value.(type) {
	case string:
		return typed
	case map[string]any:
		for _, key := range []string{"Value", "value"} {
			if text, ok := typed[key].(string); ok {
				return text
			}
		}
	}
	return ""
}

func numberedFileName(name string, suffix int) string {
	if suffix <= 0 {
		return name
	}
	ext := path.Ext(name)
	base := strings.TrimSuffix(name, ext)
	return fmt.Sprintf("%s (%d)%s", base, suffix, ext)
}

func sharePointFileAlreadyExists(err error) bool {
	if err == nil {
		return false
	}
	message := strings.ToLower(err.Error())
	return strings.Contains(message, "-2130575257") ||
		strings.Contains(message, "already exists") ||
		strings.Contains(message, "existe déjà") ||
		strings.Contains(message, "existe d\\u00e9j\\u00e0")
}

func (c *Client) sharePointDigest(ctx context.Context, siteURL, token string) (string, error) {
	var out struct {
		D struct {
			GetContextWebInformation struct {
				FormDigestValue string `json:"FormDigestValue"`
			} `json:"GetContextWebInformation"`
		} `json:"d"`
	}
	if err := c.sharePointJSON(ctx, http.MethodPost, siteURL+"/_api/contextinfo", token, "", nil, "application/json", &out); err != nil {
		return "", fmt.Errorf("get SharePoint request digest: %w", err)
	}
	if out.D.GetContextWebInformation.FormDigestValue == "" {
		return "", fmt.Errorf("SharePoint returned no request digest")
	}
	return out.D.GetContextWebInformation.FormDigestValue, nil
}

func (c *Client) shareSharePointFile(ctx context.Context, siteURL, fileURL, token, digest string, recipients []string) (string, error) {
	seen := make(map[string]bool)
	people := make([]map[string]string, 0, len(recipients))
	for _, recipient := range recipients {
		recipient = strings.TrimSpace(strings.ToLower(recipient))
		if recipient != "" && !seen[recipient] {
			seen[recipient] = true
			people = append(people, map[string]string{"Key": recipient})
		}
	}
	if len(people) == 0 {
		return "", fmt.Errorf("cannot share SharePoint file: no participant email address")
	}
	peopleJSON, _ := json.Marshal(people)
	body, _ := json.Marshal(map[string]any{
		"url": fileURL, "peoplePickerInput": string(peopleJSON), "roleValue": "role:1073741826",
		"groupId": 0, "propagateAcl": false, "sendEmail": false,
		"includeAnonymousLinkInEmail": false, "emailSubject": "", "emailBody": "", "useSimplifiedRoles": true,
	})
	var result struct {
		D struct {
			ShareObject struct {
				URL          string `json:"Url"`
				ErrorMessage string `json:"ErrorMessage"`
			} `json:"ShareObject"`
		} `json:"d"`
	}
	if err := c.sharePointJSON(ctx, http.MethodPost, strings.TrimRight(siteURL, "/")+"/_api/SP.Web.ShareObject",
		token, digest, body, "application/json;odata=verbose", &result); err != nil {
		return "", fmt.Errorf("share SharePoint file: %w", err)
	}
	if result.D.ShareObject.ErrorMessage != "" {
		return "", fmt.Errorf("share SharePoint file: %s", result.D.ShareObject.ErrorMessage)
	}
	return result.D.ShareObject.URL, nil
}

func (c *Client) sharePointJSON(ctx context.Context, method, endpoint, token, digest string, body []byte, contentType string, out any) error {
	req, err := http.NewRequestWithContext(ctx, method, endpoint, bytes.NewReader(body))
	if err != nil {
		return err
	}
	req.Header.Set("Authorization", "Bearer "+token)
	req.Header.Set("Accept", "application/json;odata=verbose")
	req.Header.Set("Content-Type", contentType)
	req.Header.Set("X-FORMS_BASED_AUTH_ACCEPTED", "f")
	req.Header.Set("User-Agent", c.cfg.UserAgent)
	if digest != "" {
		req.Header.Set("X-RequestDigest", digest)
	}
	resp, err := c.http.Do(req)
	if err != nil {
		return err
	}
	defer resp.Body.Close()
	responseBody, err := io.ReadAll(io.LimitReader(resp.Body, 2*1024*1024))
	if err != nil {
		return err
	}
	if resp.StatusCode >= 400 {
		return fmt.Errorf("server returned %d: %s", resp.StatusCode, strings.TrimSpace(string(responseBody)))
	}
	if out != nil && len(bytes.TrimSpace(responseBody)) > 0 {
		if err := json.Unmarshal(responseBody, out); err != nil {
			return fmt.Errorf("decode response: %w", err)
		}
	}
	return nil
}
