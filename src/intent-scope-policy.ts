export type Intent =
  | 'email'
  | 'calendar'
  | 'files'
  | 'people'
  | 'teams'
  | 'tasks'
  | 'search'
  | 'mixed'
  | 'sharepoint'
  | 'notes'
  | 'planner'
  | 'contacts'
  | 'meetings'
  | 'bookings'
  | 'insights';

export interface RequiredScopes {
  scopes: string[];
  workScopes: string[];
}

export const INTENT_REQUIRED_SCOPES: Record<Intent, RequiredScopes> = {
  email: { scopes: ['Mail.Read'], workScopes: [] },
  calendar: { scopes: ['Calendars.Read'], workScopes: [] },
  files: { scopes: ['Files.Read'], workScopes: ['Sites.Read.All'] },
  people: { scopes: ['People.Read', 'User.Read'], workScopes: ['User.Read.All'] },
  teams: { scopes: [], workScopes: ['Chat.Read', 'ChatMessage.Read'] },
  tasks: { scopes: ['Tasks.Read'], workScopes: [] },
  search: { scopes: ['Mail.Read', 'Files.Read'], workScopes: ['Sites.Read.All'] },
  sharepoint: { scopes: [], workScopes: ['Sites.Read.All', 'Sites.Selected'] },
  notes: { scopes: ['Notes.Read'], workScopes: [] },
  planner: { scopes: ['Tasks.Read'], workScopes: [] },
  contacts: { scopes: ['Contacts.Read'], workScopes: [] },
  meetings: { scopes: ['OnlineMeetings.Read'], workScopes: [] },
  bookings: { scopes: ['Bookings.Read.All'], workScopes: [] },
  insights: { scopes: [], workScopes: ['Sites.Read.All'] },
  mixed: { scopes: ['Mail.Read', 'Calendars.Read', 'Files.Read'], workScopes: [] },
};
