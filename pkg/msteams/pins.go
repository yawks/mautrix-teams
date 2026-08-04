package msteams

import (
	"context"
	"encoding/json"
	"fmt"
)

const pinnedMessageItemType = "Message"

// GetPinnedMessages reads the shared pinnedItems thread property used by the
// Teams web client. The server stores the value as JSON inside a JSON string.
func (c *Client) GetPinnedMessages(ctx context.Context, threadID string) ([]PinnedConversationItem, error) {
	chat, err := c.GetChat(ctx, threadID)
	if err != nil {
		return nil, fmt.Errorf("get Teams pinned messages: %w", err)
	}
	items := make([]PinnedConversationItem, 0, len(chat.PinnedItems))
	for _, item := range chat.PinnedItems {
		if item.ItemID != "" && item.ItemType == pinnedMessageItemType {
			items = append(items, item)
		}
	}
	return items, nil
}

func (c *Client) PinMessage(ctx context.Context, threadID, messageID string) error {
	items, err := c.GetPinnedMessages(ctx, threadID)
	if err != nil {
		return err
	}
	next := make([]PinnedConversationItem, 0, len(items)+1)
	next = append(next, PinnedConversationItem{ItemID: messageID, ItemType: pinnedMessageItemType})
	for _, item := range items {
		if item.ItemID != messageID {
			next = append(next, item)
		}
	}
	return c.setPinnedMessages(ctx, threadID, next)
}

func (c *Client) UnpinMessage(ctx context.Context, threadID, messageID string) error {
	items, err := c.GetPinnedMessages(ctx, threadID)
	if err != nil {
		return err
	}
	next := make([]PinnedConversationItem, 0, len(items))
	for _, item := range items {
		if item.ItemID != messageID {
			next = append(next, item)
		}
	}
	return c.setPinnedMessages(ctx, threadID, next)
}

func (c *Client) setPinnedMessages(ctx context.Context, threadID string, items []PinnedConversationItem) error {
	value, err := json.Marshal(items)
	if err != nil {
		return err
	}
	if err := c.UpdateThreadProperties(ctx, threadID, map[string]string{"pinnedItems": string(value)}); err != nil {
		return fmt.Errorf("update Teams pinned messages: %w", err)
	}
	return nil
}
