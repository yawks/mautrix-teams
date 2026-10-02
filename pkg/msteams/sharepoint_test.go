package msteams

import (
	"context"
	"encoding/json"
	"fmt"
	"net/http"
	"net/http/httptest"
	"strings"
	"testing"
	"time"
)

func TestUploadSharedFile(t *testing.T) {
	var uploadAttempts int
	var shared bool
	srv := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		if r.Header.Get("Authorization") != "Bearer sharepoint-token" {
			t.Errorf("missing SharePoint authorization")
		}
		switch {
		case strings.HasSuffix(r.URL.Path, "/_api/contextinfo"):
			_, _ = w.Write([]byte(`{"d":{"GetContextWebInformation":{"FormDigestValue":"digest"}}}`))
		case strings.Contains(r.URL.Path, "/Files/Add"):
			uploadAttempts++
			if r.Header.Get("X-RequestDigest") != "digest" {
				t.Errorf("missing request digest")
			}
			if uploadAttempts == 1 {
				w.WriteHeader(http.StatusBadRequest)
				_, _ = w.Write([]byte(`{"error":{"code":"-2130575257, Microsoft.SharePoint.SPException"}}`))
				return
			}
			_, _ = w.Write([]byte(`{"d":{"UniqueId":"{file-id}","ListId":"{list-id}","ServerRelativeUrl":"/personal/me/Documents/report (1).pdf","Name":"report (1).pdf"}}`))
		case strings.HasSuffix(r.URL.Path, "/_api/SP.Web.ShareObject"):
			shared = true
			_, _ = w.Write([]byte(`{"d":{"ShareObject":{"Url":"https://example.sharepoint.com/:w:/g/shared"}}}`))
		case strings.HasSuffix(r.URL.Path, "/ShareLink"):
			var body struct {
				Request struct {
					Settings struct {
						LinkKind int `json:"linkKind"`
						Scope    int `json:"scope"`
					} `json:"settings"`
				} `json:"request"`
			}
			if err := json.NewDecoder(r.Body).Decode(&body); err != nil || body.Request.Settings.LinkKind != 2 || body.Request.Settings.Scope != 1 {
				t.Errorf("unexpected ShareLink request: err=%v body=%+v", err, body)
			}
			_, _ = w.Write([]byte(`{"d":{"ShareLink":{"sharingLinkInfo":{"Url":{"Value":"https://example.sharepoint.com/:w:/g/shared"},"ShareId":"share-id"}}}}`))
		case strings.HasSuffix(r.URL.Path, "/_api/site"):
			_, _ = w.Write([]byte(`{"d":{"Id":"{site-id}"}}`))
		default:
			t.Fatalf("unexpected endpoint %s", r.URL.String())
		}
	}))
	defer srv.Close()

	c := newClientAt(t, srv.URL)
	c.sharePointAuth = map[string]*Token{
		strings.TrimPrefix(srv.URL, "http://"): {Value: "sharepoint-token", ExpiresAt: time.Now().Add(time.Hour)},
	}
	file, err := c.UploadSharedFile(context.Background(), SharedFile{
		SiteURL: srv.URL + "/personal/me", FileURL: srv.URL + "/personal/me/Documents/existing.pdf",
	}, "report.pdf", []byte("document"), []string{"reader@example.com"})
	if err != nil {
		t.Fatalf("UploadSharedFile: %v", err)
	}
	if uploadAttempts != 2 || !shared {
		t.Fatalf("uploadAttempts=%d shared=%v", uploadAttempts, shared)
	}
	if file.ItemID != "file-id" || file.Name != "report (1).pdf" ||
		!strings.HasSuffix(file.FileURL, "/Documents/report (1).pdf") || file.ShareURL != "https://example.sharepoint.com/:w:/g/shared" ||
		file.ShareID != "share-id" || file.SiteID != "site-id" {
		t.Fatalf("unexpected uploaded file: %+v", file)
	}
}

func TestNumberedFileName(t *testing.T) {
	if got := numberedFileName("archive.tar.gz", 2); got != "archive.tar (2).gz" {
		t.Fatalf("numberedFileName=%q", got)
	}
	if !sharePointFileAlreadyExists(fmt.Errorf(`server returned 400: {"code":"-2130575257"}`)) {
		t.Fatal("SharePoint collision was not recognized")
	}
}
