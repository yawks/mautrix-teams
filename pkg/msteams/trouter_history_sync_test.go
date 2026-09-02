package msteams

import (
	"testing"
	"time"
)

func TestMessageLossRequestsHistorySync(t *testing.T) {
	client := &Client{events: make(chan Event, 1)}
	before := time.Now().Add(-24*time.Hour - time.Minute)

	client.handleTrouterEvent([]byte(`{"name":"trouter.message_loss","args":[]}`))

	select {
	case event := <-client.events:
		if event.Type != EventTypeHistorySync {
			t.Fatalf("event type=%q, want %q", event.Type, EventTypeHistorySync)
		}
		if event.Timestamp.Before(before) || event.Timestamp.After(time.Now().Add(-23*time.Hour)) {
			t.Fatalf("unexpected repair window start: %s", event.Timestamp)
		}
	default:
		t.Fatal("message_loss did not request history repair")
	}
}
