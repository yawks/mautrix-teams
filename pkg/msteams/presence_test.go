package msteams

import (
	"context"
	"encoding/json"
	"net/http"
	"net/http/httptest"
	"testing"
)

func TestGetPresences(t *testing.T) {
	server := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		if r.URL.Path != "/v1/presence/getpresence/" {
			t.Fatalf("path=%q", r.URL.Path)
		}
		if got := r.Header.Get("Authorization"); got != "Bearer csa-token" {
			t.Fatalf("authorization=%q", got)
		}
		var request []presenceRequest
		if err := json.NewDecoder(r.Body).Decode(&request); err != nil {
			t.Fatal(err)
		}
		if len(request) != 1 || request[0].MRI != "8:orgid:alice" {
			t.Fatalf("request=%+v", request)
		}
		_, _ = w.Write([]byte(`[{"mri":"8:orgid:alice","presence":{"availability":"Busy","activity":"InAMeeting"}}]`))
	}))
	defer server.Close()

	client, err := NewClient(ClientConfig{
		UserMRI:   "8:orgid:self",
		Endpoints: Endpoints{PresenceBase: server.URL},
	})
	if err != nil {
		t.Fatal(err)
	}
	client.auth = &Token{Value: "csa-token"}
	presences, err := client.GetPresences(context.Background(), []string{"8:orgid:alice", "8:orgid:alice", ""})
	if err != nil {
		t.Fatal(err)
	}
	if len(presences) != 1 || presences[0].Activity != "InAMeeting" {
		t.Fatalf("presences=%+v", presences)
	}
}
