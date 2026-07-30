package msteams

import (
	"context"
	"encoding/json"
	"fmt"
	"strings"
)

// Presence is the availability and current activity exposed by the Teams web
// presence service. Activity distinguishes Busy from InAMeeting and InACall.
type Presence struct {
	MRI          string `json:"mri"`
	Availability string `json:"availability"`
	Activity     string `json:"activity"`
	DeviceType   string `json:"deviceType,omitempty"`
}

type presenceRequest struct {
	MRI string `json:"mri"`
}

type presenceResponse struct {
	Presence []Presence `json:"presence"`
}

type presenceEnvelope struct {
	MRI      string   `json:"mri"`
	Presence Presence `json:"presence"`
}

// GetPresences returns presence for the requested Teams MRIs using the same
// delegated api.spaces.skype.com token as the official Teams web client.
func (c *Client) GetPresences(ctx context.Context, mris []string) ([]Presence, error) {
	request := make([]presenceRequest, 0, len(mris))
	seen := make(map[string]struct{}, len(mris))
	for _, mri := range mris {
		mri = strings.TrimSpace(mri)
		if mri == "" {
			continue
		}
		if _, exists := seen[mri]; exists {
			continue
		}
		seen[mri] = struct{}{}
		request = append(request, presenceRequest{MRI: mri})
	}
	if len(request) == 0 {
		return nil, nil
	}
	base := strings.TrimRight(c.presenceBase, "/")
	if base == "" {
		base = DefaultPresenceBase
	}
	var raw json.RawMessage
	if err := c.doJSON(ctx, "POST", base+"/v1/presence/getpresence/", AuthBearer, request, &raw); err != nil {
		return nil, err
	}
	var response presenceResponse
	if err := json.Unmarshal(raw, &response); err == nil && response.Presence != nil {
		return response.Presence, nil
	}
	var envelopes []presenceEnvelope
	if err := json.Unmarshal(raw, &envelopes); err == nil {
		presences := make([]Presence, 0, len(envelopes))
		for _, envelope := range envelopes {
			presence := envelope.Presence
			if presence.MRI == "" {
				presence.MRI = envelope.MRI
			}
			presences = append(presences, presence)
		}
		return presences, nil
	}
	var direct []Presence
	if err := json.Unmarshal(raw, &direct); err != nil {
		return nil, fmt.Errorf("decode presence response: %w", err)
	}
	return direct, nil
}
