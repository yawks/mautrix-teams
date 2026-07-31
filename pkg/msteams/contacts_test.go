// mautrix-teams - A Matrix-Microsoft Teams puppeting bridge.
// Copyright (C) 2026 Sandwich
//
// This program is free software: you can redistribute it and/or modify
// it under the terms of the GNU Affero General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.
//
// This program is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
// GNU Affero General Public License for more details.
//
// You should have received a copy of the GNU Affero General Public License
// along with this program.  If not, see <https://www.gnu.org/licenses/>.
package msteams

import (
	"context"
	"encoding/json"
	"net/http"
	"net/http/httptest"
	"testing"

	"github.com/rs/zerolog"
)

const listChatsFixture = `{
  "conversations": [
    {
      "id": "19:abc@thread.v2",
      "threadProperties": {"topic": "Team planning"},
      "members": [
        {"id": "8:orgid:alice", "role": "Admin"},
        {"id": "8:orgid:bob", "role": "User"}
      ]
    },
    {
      "id": "19:xyz@thread.tacv2",
      "threadProperties": {"topic": "General"},
      "members": [{"id": "8:orgid:alice"}]
    },
    {
      "id": "8:orgid:someone",
      "members": [
        {"id": "8:orgid:alice"},
        {"id": "8:orgid:someone"}
      ]
    }
  ]
}`

func TestListChats(t *testing.T) {
	srv := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		if got := r.Header.Get("Authentication"); got != "skypetoken=skype-value" {
			t.Errorf("missing skype auth: %q", got)
		}
		if r.URL.Path != "/v1/users/ME/conversations" {
			t.Errorf("wrong path: %q", r.URL.Path)
		}
		w.Header().Set("Content-Type", "application/json")
		_, _ = w.Write([]byte(listChatsFixture))
	}))
	t.Cleanup(srv.Close)

	c, err := NewClient(ClientConfig{
		UserMRI:    "8:orgid:alice",
		SkypeToken: "skype-value",
		Endpoints:  Endpoints{ChatSvcBase: srv.URL},
		Logger:     zerolog.Nop(),
	})
	if err != nil {
		t.Fatalf("NewClient: %v", err)
	}
	t.Cleanup(func() { _ = c.Close() })

	chats, err := c.ListChats(context.Background())
	if err != nil {
		t.Fatalf("ListChats: %v", err)
	}
	if len(chats) != 3 {
		t.Fatalf("got %d chats, want 3", len(chats))
	}
	cases := []struct {
		id      string
		kind    ChatType
		members int
	}{
		{"19:abc@thread.v2", ChatTypeGroup, 2},
		{"19:xyz@thread.tacv2", ChatTypeChannel, 1},
		{"8:orgid:someone", ChatType1on1, 2},
	}
	for i, tc := range cases {
		if chats[i].ID != tc.id {
			t.Errorf("chat[%d].ID=%q want %q", i, chats[i].ID, tc.id)
		}
		if chats[i].Type != tc.kind {
			t.Errorf("chat[%d].Type=%q want %q", i, chats[i].Type, tc.kind)
		}
		if len(chats[i].Members) != tc.members {
			t.Errorf("chat[%d].Members=%d want %d", i, len(chats[i].Members), tc.members)
		}
	}
}

func TestCreateGroupChat(t *testing.T) {
	srv := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		if r.Method != http.MethodPost || r.URL.Path != "/v1/threads" {
			t.Fatalf("unexpected request %s %s", r.Method, r.URL.Path)
		}
		if got := r.Header.Get("Authentication"); got != "skypetoken=skype-value" {
			t.Fatalf("missing skype auth: %q", got)
		}
		var body struct {
			Members    []struct{ ID, Role string } `json:"members"`
			Properties map[string]any              `json:"properties"`
		}
		if err := json.NewDecoder(r.Body).Decode(&body); err != nil {
			t.Fatal(err)
		}
		if len(body.Members) != 3 || body.Properties["topic"] != "Planning" {
			t.Fatalf("unexpected create payload: %+v", body)
		}
		w.Header().Set("Content-Type", "application/json")
		w.WriteHeader(http.StatusCreated)
		_, _ = w.Write([]byte(`{"id":"19:new@thread.v2","properties":{"topic":"Planning"},"members":[{"id":"8:orgid:me"},{"id":"8:orgid:alice"},{"id":"8:orgid:bob"}]}`))
	}))
	t.Cleanup(srv.Close)
	c, err := NewClient(ClientConfig{UserMRI: "8:orgid:me", SkypeToken: "skype-value", Endpoints: Endpoints{ChatSvcBase: srv.URL}, Logger: zerolog.Nop()})
	if err != nil {
		t.Fatal(err)
	}
	t.Cleanup(func() { _ = c.Close() })
	chat, err := c.CreateGroupChat(context.Background(), "Planning", []string{"8:orgid:alice", "8:orgid:bob"})
	if err != nil {
		t.Fatal(err)
	}
	if chat.ID != "19:new@thread.v2" || chat.Type != ChatTypeGroup {
		t.Fatalf("unexpected chat: %+v", chat)
	}
}

func TestLeaveGroupChat(t *testing.T) {
	hit := false
	srv := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		hit = true
		if r.Method != http.MethodDelete {
			t.Fatalf("method %q, want DELETE", r.Method)
		}
		if r.URL.Path != "/v1/threads/19:group@thread.v2/members/8:orgid:me" {
			t.Fatalf("unexpected path %q", r.URL.Path)
		}
		if got := r.Header.Get("Authentication"); got != "skypetoken=skype-value" {
			t.Fatalf("missing skype auth: %q", got)
		}
		w.WriteHeader(http.StatusOK)
	}))
	t.Cleanup(srv.Close)

	c, err := NewClient(ClientConfig{
		UserMRI:    "8:orgid:me",
		SkypeToken: "skype-value",
		Endpoints:  Endpoints{ChatSvcBase: srv.URL},
		Logger:     zerolog.Nop(),
	})
	if err != nil {
		t.Fatalf("NewClient: %v", err)
	}
	t.Cleanup(func() { _ = c.Close() })

	if err := c.LeaveGroupChat(context.Background(), "19:group@thread.v2"); err != nil {
		t.Fatalf("LeaveGroupChat: %v", err)
	}
	if !hit {
		t.Fatal("leave group did not hit the endpoint")
	}
}

func TestStartOneOnOneCreatesStickyThread(t *testing.T) {
	srv := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		if r.Method != http.MethodPost || r.URL.Path != "/v1/threads" {
			t.Fatalf("unexpected request %s %s", r.Method, r.URL.Path)
		}
		var body struct {
			Members    []struct{ ID, Role string } `json:"members"`
			Properties map[string]any              `json:"properties"`
		}
		if err := json.NewDecoder(r.Body).Decode(&body); err != nil {
			t.Fatal(err)
		}
		if len(body.Members) != 2 || body.Members[0].Role != "Admin" || body.Members[1].ID != "8:orgid:alice" {
			t.Fatalf("unexpected members: %+v", body.Members)
		}
		if body.Properties["uniquerosterthread"] != "true" || body.Properties["fixedRoster"] != "true" {
			t.Fatalf("unexpected properties: %+v", body.Properties)
		}
		w.WriteHeader(http.StatusCreated)
	}))
	t.Cleanup(srv.Close)
	c, err := NewClient(ClientConfig{
		UserMRI:    "8:orgid:me",
		SkypeToken: "skype-value",
		Endpoints:  Endpoints{ChatSvcBase: srv.URL},
		Logger:     zerolog.Nop(),
	})
	if err != nil {
		t.Fatal(err)
	}
	t.Cleanup(func() { _ = c.Close() })

	chat, err := c.StartOneOnOne(context.Background(), " 8:orgid:alice ")
	if err != nil {
		t.Fatal(err)
	}
	if chat.ID != "19:alice_me@unq.gbl.spaces" || chat.Type != ChatType1on1 {
		t.Fatalf("unexpected chat: %+v", chat)
	}
	if len(chat.Members) != 2 || chat.Members[0].MRI != "8:orgid:me" || chat.Members[1].MRI != "8:orgid:alice" {
		t.Fatalf("unexpected members: %+v", chat.Members)
	}
}

func TestClassifyChat(t *testing.T) {
	tests := []struct {
		name string
		r    rawConversation
		want ChatType
	}{
		{"channel", rawConversation{ID: "19:a@thread.tacv2"}, ChatTypeChannel},
		{"group", rawConversation{ID: "19:a@thread.v2"}, ChatTypeGroup},
		{"meeting", rawConversation{ID: "19:a@thread.v2", ThreadProperties: rawThreadProps{ChatType: "meeting"}}, ChatTypeMeeting},
		{"one_on_one", rawConversation{ID: "8:orgid:x"}, ChatType1on1},
		{"unique_roster", rawConversation{ID: "weird", ThreadProperties: rawThreadProps{UniqueRosterThread: "true"}}, ChatType1on1},
		{"fallback_group", rawConversation{ID: "weird"}, ChatTypeGroup},
	}
	for _, tc := range tests {
		if got := classifyChat(&tc.r); got != tc.want {
			t.Errorf("%s: got %q want %q", tc.name, got, tc.want)
		}
	}
}
