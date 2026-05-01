// src/http/types.ts

export interface GraphCollection<T> {
  value: T[];
  '@odata.nextLink'?: string;
  '@odata.count'?: number;
}

export interface User {
  id: string;
  userPrincipalName?: string;
  displayName?: string;
  mail?: string;
}

export interface Team {
  id: string;
  displayName: string;
  description?: string;
  isArchived?: boolean;
  visibility?: 'private' | 'public' | 'hiddenMembership';
}

export interface Channel {
  id: string;
  displayName: string;
  description?: string;
  membershipType?: 'standard' | 'private' | 'shared';
  webUrl?: string;
}

export interface Chat {
  id: string;
  topic?: string | null;
  chatType: 'oneOnOne' | 'group' | 'meeting' | 'unknownFutureValue';
  createdDateTime: string;
  lastUpdatedDateTime: string;
  webUrl?: string;
  members?: ChatMember[];
}

export interface ChatMember {
  id: string;
  displayName?: string;
  userId?: string;
  email?: string;
  roles?: string[];
}

export interface ItemBody {
  contentType: 'text' | 'html';
  content: string;
}

export interface ChatMessage {
  id: string;
  replyToId?: string | null;
  createdDateTime: string;
  lastModifiedDateTime?: string;
  from?: {
    user?: { id: string; displayName?: string; userIdentityType?: string };
  } | null;
  body: ItemBody;
  attachments?: ChatMessageAttachment[];
  mentions?: ChatMessageMention[];
  importance?: 'normal' | 'high' | 'urgent';
  webUrl?: string;
  subject?: string | null;
}

export interface ChatMessageAttachment {
  id?: string;
  contentType?: string;
  contentUrl?: string;
  name?: string;
  thumbnailUrl?: string;
}

export interface ChatMessageMention {
  id: number;
  mentionText?: string;
  mentioned?: {
    user?: { id: string; displayName?: string; userIdentityType?: string };
  };
}

export interface SearchHit {
  hitId: string;
  rank?: number;
  summary?: string;
  resource: ChatMessage & { '@odata.type'?: string };
}

export interface SearchHitsContainer {
  hits: SearchHit[];
  total?: number;
  moreResultsAvailable?: boolean;
}

export interface SearchResponse {
  value: Array<{
    searchTerms?: string[];
    hitsContainers: SearchHitsContainer[];
  }>;
}
