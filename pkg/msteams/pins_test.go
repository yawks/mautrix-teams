package msteams

import (
	"context"
	"encoding/json"
	"net/http"
	"net/http/httptest"
	"testing"
)

func TestPinAndUnpinMessageThreadProperty(t *testing.T) {
	const threadID = "19:thread@thread.v2"
	current := `[{"itemId":"old-message","itemType":"Message"}]`
	var updates []string
	srv := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		switch {
		case r.Method == http.MethodGet && r.URL.Path == "/v1/threads/"+threadID:
			_ = json.NewEncoder(w).Encode(map[string]any{
				"id": threadID, "properties": map[string]string{"pinnedItems": current},
			})
		case r.Method == http.MethodPut && r.URL.Path == "/v1/threads/"+threadID+"/properties":
			var body map[string]string
			if err := json.NewDecoder(r.Body).Decode(&body); err != nil {
				t.Fatal(err)
			}
			current = body["pinnedItems"]
			updates = append(updates, current)
			w.WriteHeader(http.StatusNoContent)
		default:
			t.Fatalf("unexpected request %s %s", r.Method, r.URL.String())
		}
	}))
	t.Cleanup(srv.Close)

	client := newClientAt(t, srv.URL)
	if err := client.PinMessage(context.Background(), threadID, "new-message"); err != nil {
		t.Fatalf("PinMessage: %v", err)
	}
	if err := client.UnpinMessage(context.Background(), threadID, "new-message"); err != nil {
		t.Fatalf("UnpinMessage: %v", err)
	}

	if len(updates) != 2 {
		t.Fatalf("got %d updates, want 2", len(updates))
	}
	if updates[0] != `[{"itemId":"new-message","itemType":"Message"},{"itemId":"old-message","itemType":"Message"}]` {
		t.Fatalf("pin payload = %s", updates[0])
	}
	if updates[1] != `[{"itemId":"old-message","itemType":"Message"}]` {
		t.Fatalf("unpin payload = %s", updates[1])
	}
}
