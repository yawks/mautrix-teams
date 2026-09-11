package msteams

import (
	"context"
	"encoding/json"
	"net/http"
	"net/http/httptest"
	"strings"
	"testing"
)

func TestChatPictureIsParsedFromBothThreadResponses(t *testing.T) {
	for _, field := range []string{"properties", "threadProperties"} {
		var raw rawConversation
		if err := json.Unmarshal([]byte(`{"id":"19:group@thread.v2","`+field+`":{"picture":"etag@https://example/image"}}`), &raw); err != nil {
			t.Fatal(err)
		}
		if got := convertRawConversation(&raw); got.Picture != "etag@https://example/image" {
			t.Fatalf("picture lost: %+v", got)
		}
	}
}

func TestPictureActivityEmitsChatUpdate(t *testing.T) {
	client := &Client{events: make(chan Event, 1)}
	raw, _ := json.Marshal(trouterMessageResource{From: "https://example/v1/users/ME/contacts/8:orgid:alice", ConversationLink: "https://example/v1/users/ME/conversations/19:group@thread.v2/messages/1", MessageType: "ThreadActivity/PictureUpdate"})
	client.handleEventMessage("NewMessage", raw)
	select {
	case e := <-client.events:
		if e.Type != EventTypeChatUpdate || e.ThreadID != "19:group@thread.v2" {
			t.Fatal(e)
		}
	default:
		t.Fatal("missing chat update")
	}
}

func TestFetchChatAvatarUsesImageServiceCookieAndDocumentURL(t *testing.T) {
	srv := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		if !strings.HasSuffix(r.URL.Path, "/threads/19:group@thread.v2/properties/pictureV2") {
			t.Errorf("path=%s", r.URL.Path)
		}
		if r.URL.Query().Get("documentUrl") != "https://example/image?a=b" {
			t.Errorf("wrong document URL")
		}
		if !strings.Contains(r.Header.Get("Cookie"), "Bearer=auth-value") || r.Header.Get("Authorization") != "" {
			t.Errorf("incorrect avatar authentication")
		}
		w.Header().Set("Content-Type", "image/png")
		_, _ = w.Write([]byte("image"))
	}))
	defer srv.Close()
	client, err := NewClient(ClientConfig{UserMRI: "8:orgid:me", AuthToken: "auth-value", Endpoints: Endpoints{MTBase: srv.URL}})
	if err != nil {
		t.Fatal(err)
	}
	defer client.Close()
	data, _, err := client.FetchChatAvatar(context.Background(), "19:group@thread.v2", "etag@https://example/image?a=b")
	if err != nil || string(data) != "image" {
		t.Fatalf("data=%s err=%v", data, err)
	}
}
