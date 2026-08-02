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
	"bytes"
	"context"
	"encoding/json"
	"errors"
	"fmt"
	"io"
	"net/http"
	"net/url"
	"strings"
)

type conversationsResponse struct {
	Conversations []rawConversation `json:"conversations"`
}

type rawConversation struct {
	ID               string          `json:"id"`
	ThreadProperties rawThreadProps  `json:"threadProperties"`
	Properties       rawThreadProps  `json:"properties"`
	Members          []rawMember     `json:"members"`
	LastMessage      *rawMessageStub `json:"lastMessage"`
	Type             string          `json:"type"` // "Thread" for groups, missing for 1:1
}

// Mirrored under "threadProperties" by /conversations and "properties" by
// /threads/<id>; meeting subject is JSON-in-JSON under "meeting".
type rawThreadProps struct {
	Topic              string `json:"topic"`
	Description        string `json:"description"`
	ChatType           string `json:"chatType"` // meeting, group, or empty
	ThreadType         string `json:"threadType"`
	Meeting            string `json:"meeting"`
	UniqueRosterThread string `json:"uniquerosterthread"`
	ProductThreadType  string `json:"productThreadType"`
}

type rawMember struct {
	ID                string `json:"id"`
	MRI               string `json:"mri"`
	DisplayName       string `json:"displayName"`
	Email             string `json:"email"`
	UserPrincipalName string `json:"userPrincipalName"`
	Role              string `json:"role"`
}

type rawMessageStub struct {
	ID          string `json:"id"`
	ComposeTime string `json:"composetime"`
}

func (c *Client) ListChats(ctx context.Context) ([]Chat, error) {
	endpoint := c.chatSvcBaseURL() + "/v1/users/ME/conversations"
	params := url.Values{}
	params.Set("startTime", "0")
	params.Set("pageSize", "100")
	params.Set("view", "msnp24Equivalent")
	params.Set("targetType", "Passport|Skype|Lync|Thread|PSTN|Agent")
	var resp conversationsResponse
	if err := c.doJSON(ctx, "GET", endpoint+"?"+params.Encode(), AuthSkype, nil, &resp); err != nil {
		return nil, err
	}
	out := make([]Chat, 0, len(resp.Conversations))
	for _, conv := range resp.Conversations {
		out = append(out, convertRawConversation(&conv))
	}
	return out, nil
}

func (c *Client) listTeamsRequest(ctx context.Context, endpoint string) ([]byte, error) {
	req, err := http.NewRequestWithContext(ctx, "GET", endpoint, nil)
	if err != nil {
		return nil, err
	}
	c.tokenLock.RLock()
	bearer := ""
	skype := ""
	if c.csaAuth != nil {
		bearer = c.csaAuth.Value
	}
	if c.skype != nil {
		skype = c.skype.Value
	}
	c.tokenLock.RUnlock()
	if bearer == "" || skype == "" {
		return nil, ErrUnauthorized
	}
	req.Header.Set("Authorization", "Bearer "+bearer)
	req.Header.Set("X-Skypetoken", skype)
	req.Header.Set("Accept", "application/json")
	req.Header.Set("User-Agent", c.cfg.UserAgent)
	resp, err := c.http.Do(req)
	if err != nil {
		return nil, err
	}
	defer resp.Body.Close()
	body, _ := io.ReadAll(io.LimitReader(resp.Body, 4*1024*1024))
	if resp.StatusCode == http.StatusUnauthorized {
		c.log.Debug().Int("len", len(body)).Bytes("body", body).Msg("csa 401 body")
		return nil, ErrTokenExpired
	}
	if resp.StatusCode >= 400 {
		return nil, fmt.Errorf("ListTeams: %d %s", resp.StatusCode, string(body))
	}
	return body, nil
}

type teamsRosterResponse struct {
	Teams []rawTeam `json:"teams"`
}

type rawTeam struct {
	ID          string           `json:"id"`
	DisplayName string           `json:"displayName"`
	Description string           `json:"description"`
	PictureETag string           `json:"pictureETag"`
	IsArchived  bool             `json:"isArchived"`
	Channels    []rawTeamChannel `json:"channels"`
}

type rawTeamChannel struct {
	ID          string `json:"id"`
	DisplayName string `json:"displayName"`
	Description string `json:"description"`
	IsGeneral   bool   `json:"isGeneralChannel"`
	IsArchived  bool   `json:"isArchived"`
}

// ListTeams returns the user's joined teams along with the channels in each.
// The csa aggregator is regional (host comes from authz.regionGtms) and
// expects its own AAD audience (chatsvcagg.teams.microsoft.com), so we mint
// a second bearer via RefreshCsaToken and retry once if the cached token 401s.
func (c *Client) ListTeams(ctx context.Context) ([]Team, error) {
	base := c.csaBaseURL()
	if base == "" {
		base = "https://teams.microsoft.com/api/csa"
	}
	endpoint := base + "/api/v1/teams/users/me?isPrefetch=false&enableMembershipSummary=true"
	if IsConsumerTenant(c.cfg.TenantID) {
		// Consumer Teams has no concept of teams/channels: there's only
		// DMs and group chats. Short-circuit so we don't hit a 404.
		return nil, ErrNotImplemented
	}
	c.tokenLock.RLock()
	csaExp := c.csaAuth == nil || c.csaAuth.Expired()
	c.tokenLock.RUnlock()
	if csaExp {
		if err := c.RefreshCsaToken(ctx); err != nil {
			return nil, fmt.Errorf("refresh csa token for ListTeams: %w", err)
		}
	}
	body, err := c.listTeamsRequest(ctx, endpoint)
	if errors.Is(err, ErrTokenExpired) {
		if rerr := c.RefreshCsaToken(ctx); rerr != nil {
			if rerr2 := c.refreshAllTokens(ctx); rerr2 != nil {
				return nil, fmt.Errorf("refresh tokens for ListTeams: %w", rerr2)
			}
			if rerr2 := c.RefreshCsaToken(ctx); rerr2 != nil {
				return nil, fmt.Errorf("refresh csa token after full refresh: %w", rerr2)
			}
		}
		body, err = c.listTeamsRequest(ctx, endpoint)
	}
	if err != nil {
		return nil, err
	}
	var raw teamsRosterResponse
	if err := json.Unmarshal(body, &raw); err != nil {
		return nil, fmt.Errorf("decode teams roster: %w", err)
	}
	out := make([]Team, 0, len(raw.Teams))
	for _, t := range raw.Teams {
		if t.IsArchived {
			continue
		}
		team := Team{
			ID:          t.ID,
			DisplayName: t.DisplayName,
			Description: t.Description,
			PictureETag: t.PictureETag,
		}
		for _, ch := range t.Channels {
			if ch.IsArchived {
				continue
			}
			team.Channels = append(team.Channels, TeamChannel{
				ID:          ch.ID,
				DisplayName: ch.DisplayName,
				Description: ch.Description,
				IsGeneral:   ch.IsGeneral,
			})
		}
		out = append(out, team)
	}
	return out, nil
}

func (c *Client) GetChat(ctx context.Context, threadID string) (*Chat, error) {
	if threadID == "" {
		return nil, fmt.Errorf("empty thread id")
	}
	endpoint := c.chatSvcBaseURL() + "/v1/threads/" + url.PathEscape(threadID) + "?view=msnp24Equivalent"
	var resp rawConversation
	if err := c.doJSON(ctx, "GET", endpoint, AuthSkype, nil, &resp); err != nil {
		return nil, err
	}
	if resp.ID == "" {
		resp.ID = threadID
	}
	chat := convertRawConversation(&resp)
	return &chat, nil
}

type rawUserProfile struct {
	MRI               string `json:"mri"`
	ObjectID          string `json:"objectId"`
	DisplayName       string `json:"displayName"`
	GivenName         string `json:"givenName"`
	Surname           string `json:"surname"`
	Email             string `json:"email"`
	UserPrincipalName string `json:"userPrincipalName"`
	JobTitle          string `json:"jobTitle"`
	ImageURL          string `json:"profileImageUrl"`
	TenantName        string `json:"tenantName"`
	Type              string `json:"type"`
}

type fetchShortProfileResponse struct {
	Value         []rawUserProfile `json:"value"`
	ResolvedUsers []rawUserProfile `json:"resolvedUsers"`
}

const shortProfileQuery = "isMailAddress=false&canBeSmtpAddress=false&enableGuest=true&includeIBBarredUsers=true&skypeTeamsInfo=true&includeBots=true"

type Tenant struct {
	TenantID    string `json:"tenantId"`
	DisplayName string `json:"tenantName"`
	IsDefault   bool   `json:"isDefault"`
}

func (c *Client) FetchTenants(ctx context.Context) ([]Tenant, error) {
	endpoint := c.mtBaseURL() + "/beta/users/tenants"
	var raw []Tenant
	if err := c.doJSON(ctx, "GET", endpoint, AuthBearer, nil, &raw); err != nil {
		return nil, err
	}
	return raw, nil
}

// CurrentTenantName returns the display name of the organization that
// matches c.cfg.TenantID, falling back to the first tenant in the list.
// Empty string if the API is unavailable or reports nothing.
func (c *Client) CurrentTenantName(ctx context.Context) string {
	tenants, err := c.FetchTenants(ctx)
	if err != nil {
		c.log.Debug().Err(err).Msg("FetchTenants failed; space will fall back to generic label")
		return ""
	}
	if len(tenants) == 0 {
		c.log.Debug().Msg("FetchTenants returned empty; space will fall back to generic label")
		return ""
	}
	for _, t := range tenants {
		if t.TenantID == c.cfg.TenantID && t.DisplayName != "" {
			return t.DisplayName
		}
	}
	return tenants[0].DisplayName
}

func (c *Client) FetchShortProfiles(ctx context.Context, mris []string) ([]User, error) {
	if len(mris) == 0 {
		return nil, nil
	}
	endpoint := c.mtBaseURL() + "/beta/users/fetchShortProfile?" + shortProfileQuery
	var resp fetchShortProfileResponse
	if err := c.doJSON(ctx, "POST", endpoint, AuthBearer, mris, &resp); err != nil {
		return nil, err
	}
	rows := make([]rawUserProfile, 0, len(resp.Value)+len(resp.ResolvedUsers))
	rows = append(rows, resp.Value...)
	rows = append(rows, resp.ResolvedUsers...)
	out := make([]User, 0, len(rows))
	for _, r := range rows {
		out = append(out, profileToUser(&r))
	}
	return out, nil
}

func (c *Client) GetUser(ctx context.Context, mri string) (*User, error) {
	if mri == "" {
		return nil, fmt.Errorf("empty mri")
	}
	users, err := c.FetchShortProfiles(ctx, []string{mri})
	if err != nil {
		return nil, err
	}
	var profile *User
	for i := range users {
		if strings.EqualFold(users[i].MRI, mri) ||
			mriLookupObjectID(users[i].ObjectID) == mriLookupObjectID(mri) ||
			users[i].MRI == "" {
			p := users[i]
			profile = &p
			break
		}
	}
	// Enrich with substrate-cached profile (phones, dept, etc.). If we don't
	// have one yet but know the user's email, try a one-shot substrate query
	// so the first GetUser fills the cache for subsequent calls.
	if cached := c.CachedUserProfile(mri); cached != nil {
		if profile == nil {
			return cached, nil
		}
		mergeUserProfile(profile, cached)
	} else {
		// Try the loki/delve people-card first (postal address, full phones).
		// Fall back to substrate (search-by-email) when delve isn't available.
		if rich, err := c.FetchPersonCard(ctx, mri); err == nil && rich != nil {
			if profile == nil {
				c.CacheUserProfile(rich)
				return rich, nil
			}
			mergeUserProfile(profile, rich)
			c.CacheUserProfile(profile)
		} else if profile != nil && profile.Email != "" {
			if hits, err := c.SearchUsers(ctx, profile.Email); err == nil {
				for i := range hits {
					if strings.EqualFold(hits[i].MRI, mri) {
						mergeUserProfile(profile, &hits[i])
						break
					}
				}
			}
		}
	}
	if profile != nil {
		return profile, nil
	}
	return nil, ErrNotFound
}

func mergeUserProfile(dst, src *User) {
	if dst.JobTitle == "" {
		dst.JobTitle = src.JobTitle
	}
	if dst.Company == "" {
		dst.Company = src.Company
	}
	if dst.Department == "" {
		dst.Department = src.Department
	}
	if dst.Office == "" {
		dst.Office = src.Office
	}
	if dst.Email == "" {
		dst.Email = src.Email
	}
	if len(dst.Phones) == 0 {
		dst.Phones = src.Phones
	}
}

func profileToUser(r *rawUserProfile) User {
	return User{
		MRI:         firstNonEmpty(r.MRI, r.ObjectID),
		ObjectID:    r.ObjectID,
		DisplayName: firstNonEmpty(r.DisplayName, joinName(r.GivenName, r.Surname), r.UserPrincipalName, r.Email),
		Email:       firstNonEmpty(r.Email, r.UserPrincipalName),
		JobTitle:    r.JobTitle,
		AvatarURL:   r.ImageURL,
	}
}

// FetchAvatar downloads a user's profile picture using browser-style cookie
// auth (the asset endpoint rejects Authorization headers).
func (c *Client) FetchAvatar(ctx context.Context, mri string) ([]byte, string, error) {
	if mri == "" {
		return nil, "", fmt.Errorf("empty mri")
	}
	if c.cfg.UserMRI == "" {
		return nil, "", fmt.Errorf("client missing self mri")
	}
	selfOID := strings.TrimPrefix(c.cfg.UserMRI, "8:orgid:")
	endpoint := c.mtBaseURL() + "/beta/users/" + url.PathEscape(selfOID) + "/profilepicturev2/" + mri
	if err := c.ensureFreshTokens(ctx, true, false); err != nil {
		return nil, "", err
	}
	req, err := http.NewRequestWithContext(ctx, "GET", endpoint, nil)
	if err != nil {
		return nil, "", err
	}
	c.tokenLock.RLock()
	authToken := ""
	if c.auth != nil {
		authToken = c.auth.Value
	}
	c.tokenLock.RUnlock()
	if authToken == "" {
		return nil, "", ErrUnauthorized
	}
	req.Header.Set("Cookie", "authtoken=Bearer="+authToken+"&Origin=https://teams.microsoft.com")
	req.Header.Set("Referer", "https://teams.microsoft.com/")
	req.Header.Set("User-Agent", c.cfg.UserAgent)

	resp, err := c.http.Do(req)
	if err != nil {
		return nil, "", err
	}
	defer resp.Body.Close()
	if resp.StatusCode == http.StatusNotFound {
		return nil, "", ErrNotFound
	}
	if resp.StatusCode >= 400 {
		body, _ := io.ReadAll(io.LimitReader(resp.Body, 1024))
		return nil, "", fmt.Errorf("avatar fetch %s: %d %s", mri, resp.StatusCode, string(body))
	}
	data, err := io.ReadAll(io.LimitReader(resp.Body, 4*1024*1024))
	if err != nil {
		return nil, "", err
	}
	ct := resp.Header.Get("Content-Type")
	if ct == "" {
		ct = "image/jpeg"
	}
	return data, ct, nil
}

func joinName(first, last string) string {
	if first == "" {
		return last
	}
	if last == "" {
		return first
	}
	return first + " " + last
}

// SearchUsers queries the Microsoft 365 substrate "people picker" endpoint
// that Teams' web client uses for free-form user search. It needs a token
// scoped to outlook.office.com/search; one is minted via RefreshSearchToken.
func (c *Client) SearchUsers(ctx context.Context, query string) ([]User, error) {
	query = strings.TrimSpace(query)
	if query == "" {
		return nil, nil
	}
	if err := c.ensureSearchToken(ctx); err != nil {
		return nil, err
	}
	reqID := newUUIDv4()
	sessionID := newUUIDv4()
	raw, _ := json.Marshal(substrateRequest(query, reqID))
	req, err := http.NewRequestWithContext(ctx, "POST",
		"https://substrate.office.com/search/api/v1/suggestions?scenario=powerbar",
		bytes.NewReader(raw))
	if err != nil {
		return nil, err
	}
	req.Header.Set("Authorization", "Bearer "+c.searchTokenValue())
	req.Header.Set("Content-Type", "application/json")
	req.Header.Set("Accept", "application/json")
	req.Header.Set("X-AnchorMailbox", "Oid:"+strings.TrimPrefix(c.cfg.UserMRI, "8:orgid:")+"@"+c.cfg.TenantID)
	req.Header.Set("client-request-id", reqID)
	req.Header.Set("clientrequestid", reqID)
	req.Header.Set("client-session-id", sessionID)
	req.Header.Set("x-ms-request-id", reqID)
	req.Header.Set("x-ms-session-id", sessionID)
	req.Header.Set("x-client-version", "T2.1")
	req.Header.Set("Referer", "https://teams.microsoft.com/")
	resp, err := c.http.Do(req)
	if err != nil {
		return nil, err
	}
	defer resp.Body.Close()
	if resp.StatusCode == http.StatusUnauthorized {
		return nil, ErrTokenExpired
	}
	rb, _ := io.ReadAll(io.LimitReader(resp.Body, 256*1024))
	if resp.StatusCode >= 400 {
		return nil, fmt.Errorf("substrate search: %d %s", resp.StatusCode, string(rb))
	}
	users := parseSubstrateResponse(rb)
	for i := range users {
		c.CacheUserProfile(&users[i])
	}
	return users, nil
}

// FetchPersonCard hits the loki/delve people-card endpoint that powers Teams'
// "live persona card" popout. Returns a User with phones, postal address,
// office, manager metadata etc. Falls back to ErrNotImplemented when the
// delve token can't be minted (e.g. consumer accounts).
func (c *Client) FetchPersonCard(ctx context.Context, mri string) (*User, error) {
	if mri == "" {
		return nil, fmt.Errorf("empty mri")
	}
	if err := c.ensureDelveToken(ctx); err != nil {
		return nil, err
	}
	hostAppPersonaID, _ := json.Marshal(map[string]any{
		"userId":          mri,
		"isSharedChannel": false,
	})
	q := url.Values{}
	q.Set("hostAppPersonaId", string(hostAppPersonaID))
	q.Set("teamsMri", mri)
	base := c.delveBaseURL()
	if base == "" {
		base = "https://nam.loki.delve.office.com"
	}
	endpoint := base + "/api/v2/person?" + q.Encode()
	body, _ := json.Marshal(map[string]any{
		"X-ClientType":                "Teams",
		"X-ClientFeature":             "LivePersonaCard",
		"X-ClientArchitectureVersion": "v2",
		"X-ClientScenario":            "PersonaInfo",
	})
	req, err := http.NewRequestWithContext(ctx, "POST", endpoint, bytes.NewReader(body))
	if err != nil {
		return nil, err
	}
	req.Header.Set("Authorization", "Bearer "+c.delveTokenValue())
	req.Header.Set("Content-Type", "application/json")
	req.Header.Set("Accept", "application/json")
	resp, err := c.http.Do(req)
	if err != nil {
		return nil, err
	}
	defer resp.Body.Close()
	if resp.StatusCode == http.StatusUnauthorized {
		return nil, ErrTokenExpired
	}
	rb, _ := io.ReadAll(io.LimitReader(resp.Body, 256*1024))
	if resp.StatusCode >= 400 {
		return nil, fmt.Errorf("loki person: %d %s", resp.StatusCode, string(rb))
	}
	return parseLokiPerson(mri, rb)
}

func (c *Client) ensureDelveToken(ctx context.Context) error {
	c.tokenLock.RLock()
	tok := c.delveAuth
	c.tokenLock.RUnlock()
	if tok != nil && !tok.Expired() {
		return nil
	}
	return c.RefreshDelveToken(ctx)
}

func (c *Client) delveTokenValue() string {
	c.tokenLock.RLock()
	defer c.tokenLock.RUnlock()
	if c.delveAuth == nil {
		return ""
	}
	return c.delveAuth.Value
}

func parseLokiPerson(mri string, data []byte) (*User, error) {
	var out struct {
		Person struct {
			Names []struct {
				Value struct {
					DisplayName string `json:"displayName"`
				} `json:"value"`
			} `json:"names"`
			Phones []struct {
				Value struct {
					Type   string `json:"type"`
					Number string `json:"number"`
				} `json:"value"`
			} `json:"phones"`
			EmailAddresses []struct {
				Value struct {
					Address string `json:"address"`
				} `json:"value"`
			} `json:"emailAddresses"`
			PostalAddresses []struct {
				Value struct {
					Type   string `json:"type"`
					City   string `json:"city"`
					Street string `json:"street"`
				} `json:"value"`
			} `json:"postalAddresses"`
			WorkDetails []struct {
				Value struct {
					CompanyName string `json:"companyName"`
					JobTitle    string `json:"jobTitle"`
					Department  string `json:"department"`
					Office      string `json:"office"`
				} `json:"value"`
			} `json:"workDetails"`
			UserPrincipalName string `json:"userPrincipalName"`
		} `json:"person"`
	}
	if err := json.Unmarshal(data, &out); err != nil {
		return nil, fmt.Errorf("loki person: decode: %w", err)
	}
	u := &User{MRI: mri, Email: out.Person.UserPrincipalName}
	if len(out.Person.Names) > 0 {
		u.DisplayName = out.Person.Names[0].Value.DisplayName
	}
	if len(out.Person.WorkDetails) > 0 {
		w := out.Person.WorkDetails[0].Value
		u.Company, u.JobTitle, u.Department, u.Office = w.CompanyName, w.JobTitle, w.Department, w.Office
	}
	if u.Email == "" && len(out.Person.EmailAddresses) > 0 {
		u.Email = out.Person.EmailAddresses[0].Value.Address
	}
	for _, p := range out.Person.Phones {
		u.Phones = append(u.Phones, Phone{Type: p.Value.Type, Number: p.Value.Number})
	}
	if u.Office == "" && len(out.Person.PostalAddresses) > 0 {
		a := out.Person.PostalAddresses[0].Value
		if a.City != "" {
			u.Office = a.City
		}
	}
	return u, nil
}

func (c *Client) ensureSearchToken(ctx context.Context) error {
	c.tokenLock.RLock()
	tok := c.searchAuth
	c.tokenLock.RUnlock()
	if tok != nil && !tok.Expired() {
		return nil
	}
	return c.RefreshSearchToken(ctx)
}

func (c *Client) searchTokenValue() string {
	c.tokenLock.RLock()
	defer c.tokenLock.RUnlock()
	if c.searchAuth == nil {
		return ""
	}
	return c.searchAuth.Value
}

func substrateRequest(query, reqID string) map[string]any {
	return map[string]any{
		"EntityRequests": []any{
			map[string]any{
				"Query": map[string]any{
					"QueryString":           query,
					"DisplayQueryString":    query,
					"NormalizedQueryString": query,
				},
				"EntityType": "People",
				"Size":       10,
				"Fields": []string{
					"Id", "MRI", "DisplayName", "EmailAddresses", "PeopleType",
					"PeopleSubtype", "UserPrincipalName", "GivenName", "Surname",
					"JobTitle", "CompanyName", "Department", "Phones",
				},
				"Filter": map[string]any{
					"And": []any{
						map[string]any{"Or": []any{
							map[string]any{"Term": map[string]string{"PeopleType": "Person"}},
							map[string]any{"Term": map[string]string{"PeopleType": "Other"}},
						}},
						map[string]any{"Or": []any{
							map[string]any{"Term": map[string]string{"PeopleSubtype": "OrganizationUser"}},
							map[string]any{"Term": map[string]string{"PeopleSubtype": "MTOUser"}},
							map[string]any{"Term": map[string]string{"PeopleSubtype": "PersonalContact"}},
							map[string]any{"Term": map[string]string{"PeopleSubtype": "Guest"}},
						}},
						map[string]any{"Or": []any{
							map[string]any{"Term": map[string]string{"Flags": "NonHidden"}},
						}},
					},
				},
				"Provenances": []string{"Mailbox", "Directory"},
				"From":        0,
			},
		},
		"Scenario": map[string]any{
			"Name": "powerbar",
			"Dimensions": []any{map[string]string{
				"DimensionName":  "QueryType",
				"DimensionValue": "PeopleCentricSearch",
			}},
		},
		"Cvid":       reqID,
		"LogicalId":  reqID,
		"AppName":    "Microsoft Teams",
		"dataSource": "personScoped",
	}
}

func parseSubstrateResponse(data []byte) []User {
	var out struct {
		Groups []struct {
			Suggestions []struct {
				MRI               string   `json:"MRI"`
				DisplayName       string   `json:"DisplayName"`
				GivenName         string   `json:"GivenName"`
				Surname           string   `json:"Surname"`
				EmailAddresses    []string `json:"EmailAddresses"`
				UserPrincipalName string   `json:"UserPrincipalName"`
				JobTitle          string   `json:"JobTitle"`
				CompanyName       string   `json:"CompanyName"`
				Department        string   `json:"Department"`
				Phones            []struct {
					Number string `json:"Number"`
					Type   string `json:"Type"`
				} `json:"Phones"`
			} `json:"Suggestions"`
		} `json:"Groups"`
	}
	if err := json.Unmarshal(data, &out); err != nil {
		return nil
	}
	var users []User
	seen := map[string]bool{}
	for _, g := range out.Groups {
		for _, s := range g.Suggestions {
			if s.MRI == "" || seen[s.MRI] {
				continue
			}
			seen[s.MRI] = true
			email := s.UserPrincipalName
			if email == "" && len(s.EmailAddresses) > 0 {
				email = s.EmailAddresses[0]
			}
			phones := make([]Phone, 0, len(s.Phones))
			for _, p := range s.Phones {
				phones = append(phones, Phone{Type: p.Type, Number: p.Number})
			}
			users = append(users, User{
				MRI:         s.MRI,
				DisplayName: s.DisplayName,
				Email:       email,
				JobTitle:    s.JobTitle,
				Company:     s.CompanyName,
				Department:  s.Department,
				Phones:      phones,
			})
		}
	}
	return users
}

// StartOneOnOne creates (or resolves) the stable sticky thread for a DM.
// Sending directly to either the user's MRI or an unmaterialised calculated
// thread ID is rejected by the chat service.
func (c *Client) StartOneOnOne(ctx context.Context, targetMRI string) (*Chat, error) {
	targetMRI = strings.TrimSpace(targetMRI)
	if targetMRI == "" {
		return nil, fmt.Errorf("empty target MRI")
	}
	payload := map[string]any{
		"members": []map[string]any{
			{"id": c.cfg.UserMRI, "role": "Admin"},
			{"id": targetMRI, "role": "Admin"},
		},
		"properties": map[string]any{
			"threadType":         "chat",
			"chatFilesIndexId":   "2",
			"uniquerosterthread": "true",
			"fixedRoster":        "true",
		},
	}
	var raw rawConversation
	if err := c.doJSON(ctx, http.MethodPost, c.chatSvcBaseURL()+"/v1/threads", AuthSkype, payload, &raw); err != nil {
		return nil, fmt.Errorf("create one-to-one chat: %w", err)
	}
	threadID := raw.ID
	if threadID == "" {
		threadID = DM1on1ThreadID(c.cfg.UserMRI, targetMRI)
	}
	if threadID == "" {
		wanted := map[string]bool{
			strings.ToLower(c.cfg.UserMRI): true,
			strings.ToLower(targetMRI):     true,
		}
		chats, err := c.ListChats(ctx)
		if err != nil {
			return nil, fmt.Errorf("resolve created one-to-one chat: %w", err)
		}
		for i := len(chats) - 1; i >= 0; i-- {
			if chats[i].Type == ChatType1on1 && chatHasMembers(chats[i], wanted) {
				threadID = chats[i].ID
				break
			}
		}
	}
	if threadID == "" {
		return nil, fmt.Errorf("one-to-one chat was created but its thread id could not be resolved")
	}
	return &Chat{
		ID:   threadID,
		Type: ChatType1on1,
		Members: []Member{
			{MRI: c.cfg.UserMRI},
			{MRI: targetMRI},
		},
	}, nil
}

func (c *Client) CreateGroupChat(ctx context.Context, topic string, members []string) (*Chat, error) {
	topic = strings.TrimSpace(topic)
	if len(members) < 2 {
		return nil, fmt.Errorf("group chat requires at least two other members")
	}
	seen := make(map[string]bool, len(members)+1)
	requestMembers := make([]map[string]any, 0, len(members)+1)
	addMember := func(mri, role string) {
		mri = strings.TrimSpace(mri)
		if mri == "" || seen[strings.ToLower(mri)] {
			return
		}
		seen[strings.ToLower(mri)] = true
		requestMembers = append(requestMembers, map[string]any{"id": mri, "role": role})
	}
	addMember(c.cfg.UserMRI, "Admin")
	for _, mri := range members {
		addMember(mri, "User")
	}
	payload := map[string]any{
		"members": requestMembers,
		"properties": map[string]any{
			"threadType": "Chat", "templateType": "Chat", "chatType": "group", "topic": topic,
		},
	}
	var raw rawConversation
	if err := c.doJSON(ctx, http.MethodPost, c.chatSvcBaseURL()+"/v1/threads", AuthSkype, payload, &raw); err != nil {
		return nil, fmt.Errorf("create group chat: %w", err)
	}
	if raw.ID != "" {
		chat := convertRawConversation(&raw)
		chat.Type = ChatTypeGroup
		if chat.Topic == "" {
			chat.Topic = topic
		}
		return &chat, nil
	}
	// Some regional deployments return 201 with an empty body. Resolve the
	// newly-created thread from the refreshed conversation list.
	chats, err := c.ListChats(ctx)
	if err != nil {
		return nil, fmt.Errorf("resolve created group chat: %w", err)
	}
	for i := len(chats) - 1; i >= 0; i-- {
		if chats[i].Type == ChatTypeGroup && strings.EqualFold(strings.TrimSpace(chats[i].Topic), topic) && chatHasMembers(chats[i], seen) {
			return &chats[i], nil
		}
	}
	return nil, fmt.Errorf("group chat was created but its thread id could not be resolved")
}

// UpdateThreadProperties changes mutable metadata on a chat thread.
func (c *Client) UpdateThreadProperties(ctx context.Context, threadID string, properties map[string]string) error {
	threadID = strings.TrimSpace(threadID)
	if threadID == "" {
		return fmt.Errorf("empty thread id")
	}
	if len(properties) == 0 {
		return nil
	}
	endpoint := c.chatSvcBaseURL() + "/v1/threads/" + url.PathEscape(threadID) + "/properties"
	return c.doJSON(ctx, http.MethodPut, endpoint, AuthSkype, properties, nil)
}

// UpdateThreadTopic changes the display name of a group chat.
func (c *Client) UpdateThreadTopic(ctx context.Context, threadID, topic string) error {
	topic = strings.TrimSpace(topic)
	if topic == "" {
		return fmt.Errorf("empty thread topic")
	}
	return c.UpdateThreadProperties(ctx, threadID, map[string]string{"topic": topic})
}

// UpdateThreadDescription changes the description metadata of a chat thread.
func (c *Client) UpdateThreadDescription(ctx context.Context, threadID, description string) error {
	return c.UpdateThreadProperties(ctx, threadID, map[string]string{"description": strings.TrimSpace(description)})
}

// AddThreadMembers adds regular members to a group chat roster.
func (c *Client) AddThreadMembers(ctx context.Context, threadID string, memberMRIs []string) error {
	threadID = strings.TrimSpace(threadID)
	if threadID == "" {
		return fmt.Errorf("empty thread id")
	}
	seen := make(map[string]bool, len(memberMRIs))
	members := make([]map[string]string, 0, len(memberMRIs))
	for _, memberMRI := range memberMRIs {
		memberMRI = strings.TrimSpace(memberMRI)
		key := strings.ToLower(memberMRI)
		if memberMRI == "" || seen[key] {
			continue
		}
		seen[key] = true
		members = append(members, map[string]string{"id": memberMRI, "role": "User"})
	}
	if len(members) == 0 {
		return nil
	}
	endpoint := c.chatSvcBaseURL() + "/v1/threads/" + url.PathEscape(threadID) + "/members"
	return c.doJSON(ctx, http.MethodPost, endpoint, AuthSkype, map[string]any{"members": members}, nil)
}

// UpdateThreadMemberRole promotes or demotes an existing group chat member.
func (c *Client) UpdateThreadMemberRole(ctx context.Context, threadID, memberMRI, role string) error {
	threadID = strings.TrimSpace(threadID)
	memberMRI = strings.TrimSpace(memberMRI)
	if threadID == "" || memberMRI == "" {
		return fmt.Errorf("empty thread or member id")
	}
	if !strings.EqualFold(role, "Admin") && !strings.EqualFold(role, "User") {
		return fmt.Errorf("invalid thread member role %q", role)
	}
	if strings.EqualFold(role, "Admin") {
		role = "Admin"
	} else {
		role = "User"
	}
	endpoint := c.chatSvcBaseURL() + "/v1/threads/" + url.PathEscape(threadID) +
		"/members/" + url.PathEscape(memberMRI)
	return c.doJSON(ctx, http.MethodPut, endpoint, AuthSkype, map[string]string{"role": role}, nil)
}

// RemoveThreadMember removes a member from a group chat roster.
func (c *Client) RemoveThreadMember(ctx context.Context, threadID, memberMRI string) error {
	threadID = strings.TrimSpace(threadID)
	memberMRI = strings.TrimSpace(memberMRI)
	if threadID == "" || memberMRI == "" {
		return fmt.Errorf("empty thread or member id")
	}
	endpoint := c.chatSvcBaseURL() + "/v1/threads/" + url.PathEscape(threadID) +
		"/members/" + url.PathEscape(memberMRI)
	return c.doJSON(ctx, http.MethodDelete, endpoint, AuthSkype, nil, nil)
}

// LeaveGroupChat removes the authenticated user from a group chat roster.
func (c *Client) LeaveGroupChat(ctx context.Context, threadID string) error {
	return c.RemoveThreadMember(ctx, threadID, c.cfg.UserMRI)
}

func chatHasMembers(chat Chat, wanted map[string]bool) bool {
	found := make(map[string]bool, len(chat.Members))
	for _, member := range chat.Members {
		found[strings.ToLower(member.MRI)] = true
	}
	if len(found) != len(wanted) {
		return false
	}
	for mri := range wanted {
		if !found[mri] {
			return false
		}
	}
	return true
}

func convertRawConversation(r *rawConversation) Chat {
	c := Chat{
		ID:          r.ID,
		Topic:       firstNonEmpty(r.ThreadProperties.Topic, r.Properties.Topic, meetingSubject(&r.Properties), meetingSubject(&r.ThreadProperties)),
		Description: firstNonEmpty(r.ThreadProperties.Description, r.Properties.Description),
	}
	c.Type = classifyChat(r)
	for _, m := range r.Members {
		mri := m.MRI
		if mri == "" {
			mri = m.ID
		}
		if mri == "" {
			continue
		}
		c.Members = append(c.Members, Member{
			MRI: mri, DisplayName: m.DisplayName,
			Email: firstNonEmpty(m.Email, m.UserPrincipalName), Role: m.Role,
		})
	}
	if r.LastMessage != nil {
		c.LastUpdated = ParseTeamsTime(r.LastMessage.ComposeTime)
	}
	// /conversations omits members for 1:1 DMs - both peers are encoded in
	// the thread id itself.
	if len(c.Members) == 0 {
		if peers := peersFromThreadID(r.ID); len(peers) > 0 {
			for _, p := range peers {
				c.Members = append(c.Members, Member{MRI: p})
			}
		}
	}
	return c
}

func mriLookupObjectID(value string) string {
	value = strings.ToLower(strings.TrimSpace(value))
	value = strings.TrimPrefix(value, "8:orgid:")
	return strings.ReplaceAll(value, "-", "")
}

func peersFromThreadID(id string) []string {
	const unqSuffix = "@unq.gbl.spaces"
	if strings.HasSuffix(id, unqSuffix) && strings.HasPrefix(id, "19:") {
		body := id[len("19:") : len(id)-len(unqSuffix)]
		parts := strings.Split(body, "_")
		out := make([]string, 0, len(parts))
		for _, p := range parts {
			if strings.Count(p, "-") != 4 {
				continue
			}
			out = append(out, "8:orgid:"+p)
		}
		return out
	}
	if strings.HasPrefix(id, "8:") {
		return []string{id}
	}
	return nil
}

func isMeeting(p *rawThreadProps) bool {
	return strings.EqualFold(p.ChatType, "meeting") || strings.EqualFold(p.ThreadType, "meeting")
}

func meetingSubject(p *rawThreadProps) string {
	if p == nil || p.Meeting == "" {
		return ""
	}
	var m struct {
		Subject string `json:"subject"`
	}
	if json.Unmarshal([]byte(p.Meeting), &m) != nil {
		return ""
	}
	return m.Subject
}

func classifyChat(r *rawConversation) ChatType {
	if strings.HasSuffix(r.ID, "@thread.tacv2") {
		return ChatTypeChannel
	}
	if strings.HasSuffix(r.ID, "@thread.v2") {
		if isMeeting(&r.ThreadProperties) || isMeeting(&r.Properties) || strings.HasPrefix(r.ID, "19:meeting_") {
			return ChatTypeMeeting
		}
		return ChatTypeGroup
	}
	if strings.HasPrefix(r.ID, "8:") {
		return ChatType1on1
	}
	if r.ThreadProperties.UniqueRosterThread == "true" || r.ThreadProperties.ProductThreadType == "OneToOneChat" {
		return ChatType1on1
	}
	return ChatTypeGroup
}
