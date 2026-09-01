package msteams

import (
	"encoding/json"
	"testing"
)

func TestHandleEventMessageEmitsChatUpdateForMembershipChanges(t *testing.T) {
	client := &Client{events: make(chan Event, 1)}
	resource, err := json.Marshal(trouterMessageResource{
		ID:               "event-1",
		From:             "https://example.invalid/v1/users/ME/contacts/8:orgid:alice",
		ConversationLink: "https://example.invalid/v1/users/ME/conversations/19:group@thread.v2/messages/event-1",
		MessageType:      "ThreadActivity/AddMember",
		ComposeTime:      "2026-08-31T10:00:00.000Z",
	})
	if err != nil {
		t.Fatal(err)
	}

	client.handleEventMessage("NewMessage", resource)

	select {
	case event := <-client.events:
		if event.Type != EventTypeChatUpdate || event.ThreadID != "19:group@thread.v2" {
			t.Fatalf("unexpected event: %+v", event)
		}
	default:
		t.Fatal("membership change did not emit a chat update")
	}
}
