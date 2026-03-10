/**
 * Example questions the MCP server can answer 100% using Microsoft 365 interfaces.
 * Used by GET /capabilities and the get-example-questions tool.
 * Keep in sync with docs/example-questions.md when adding tools or capabilities.
 */

export interface ExampleQuestion {
  de: string;
  en: string;
}

export interface ExampleQuestionCategory {
  id: string;
  nameDe: string;
  nameEn: string;
  questions: ExampleQuestion[];
}

export const exampleQuestionsCategories: ExampleQuestionCategory[] = [
  {
    id: 'email',
    nameDe: 'E-Mail (Outlook)',
    nameEn: 'Email (Outlook)',
    questions: [
      { de: 'Zeig mir die letzten 10 E-Mails.', en: 'Show my last 10 emails.' },
      {
        de: 'Suche E-Mails von [Person] oder zum Betreff "[Thema]".',
        en: 'Search emails from [person] or with subject "[topic]".',
      },
      {
        de: 'Sende eine E-Mail an [Empfänger] mit Betreff "[…]" und Text "[…]".',
        en: 'Send an email to [recipient] with subject "[…]" and body "[…]".',
      },
      { de: 'Erstelle einen E-Mail-Entwurf.', en: 'Create an email draft.' },
      { de: 'Liste Anhänge einer E-Mail.', en: 'List attachments of an email.' },
      {
        de: 'Fasse den E-Mail-Verlauf zu "[Thema]" zusammen.',
        en: 'Summarize the email thread about "[topic]".',
      },
      {
        de: 'Welche offenen Punkte/Action-Items habe ich aus letzten E-Mails?',
        en: 'What action items do I have from recent emails?',
      },
    ],
  },
  {
    id: 'calendar',
    nameDe: 'Kalender',
    nameEn: 'Calendar',
    questions: [
      {
        de: 'Welche Termine habe ich heute / diese Woche?',
        en: 'What meetings do I have today / this week?',
      },
      { de: 'Finde freie Zeiten für ein Meeting.', en: 'Find available meeting times.' },
      {
        de: 'Erstelle einen Termin "[Titel]" am [Datum] von [Uhrzeit] bis [Uhrzeit].',
        en: 'Create a calendar event "[title]" on [date] from [time] to [time].',
      },
      { de: 'Zeig mir alle Termine mit [Person].', en: 'Show all meetings with [person].' },
      {
        de: 'Bereite mich auf mein Meeting mit [Team/Thema] vor.',
        en: 'Prepare me for my meeting with [team/topic].',
      },
    ],
  },
  {
    id: 'teams',
    nameDe: 'Teams & Chats',
    nameEn: 'Teams & Chats',
    questions: [
      { de: 'Zeig meine Teams.', en: 'List my Teams.' },
      {
        de: 'Zeig die letzten Chats / Nachrichten mit [Person].',
        en: 'Show recent chats / messages with [person].',
      },
      { de: 'Sende eine Chat-Nachricht in [Chat].', en: 'Send a chat message in [chat].' },
      {
        de: 'Gib mir eine Übersicht aller Interaktionen mit [Person] (E-Mails, Chats, Termine).',
        en: 'Give me an overview of all interactions with [person] (emails, chats, meetings).',
      },
    ],
  },
  {
    id: 'files',
    nameDe: 'Dateien (OneDrive / SharePoint)',
    nameEn: 'Files (OneDrive / SharePoint)',
    questions: [
      {
        de: 'Liste Dateien in [Ordner] oder im Stamm.',
        en: 'List files in [folder] or root.',
      },
      { de: 'Suche Dateien zu "[Thema]".', en: 'Search for files about "[topic]".' },
      { de: 'Lade den Inhalt von [Datei] herunter.', en: 'Download content of [file].' },
      {
        de: 'Welche Dateien hat [Person] mit mir geteilt?',
        en: 'What files did [person] share with me?',
      },
      {
        de: 'Zeig SharePoint-Standorte / Listen zu "[Suchbegriff]".',
        en: 'Show SharePoint sites / lists for "[search term]".',
      },
    ],
  },
  {
    id: 'search',
    nameDe: 'Suche (Microsoft 365)',
    nameEn: 'Search (Microsoft 365)',
    questions: [
      {
        de: 'Suche in allem (E-Mails, Dateien, Kalender, Teams) nach "[Suchbegriff]".',
        en: 'Search everything (email, files, calendar, Teams) for "[search term]".',
      },
      {
        de: 'Finde alles zu "[Projektname]" / "[Thema]".',
        en: 'Find everything about "[project name]" / "[topic]".',
      },
    ],
  },
  {
    id: 'tasks',
    nameDe: 'Aufgaben (To-Do / Planner)',
    nameEn: 'Tasks (To-Do / Planner)',
    questions: [
      {
        de: 'Zeig alle meine Aufgaben (To-Do und Planner).',
        en: 'Show all my tasks (To-Do and Planner).',
      },
      { de: 'Was muss ich noch erledigen?', en: 'What do I still need to do?' },
      {
        de: 'Erstelle eine Aufgabe "[Titel]" in [Liste].',
        en: 'Create a task "[title]" in [list].',
      },
      {
        de: 'Zeig Planner-Aufgaben für [Plan].',
        en: 'Show Planner tasks for [plan].',
      },
    ],
  },
  {
    id: 'contacts',
    nameDe: 'Kontakte',
    nameEn: 'Contacts',
    questions: [
      { de: 'Liste meine Outlook-Kontakte.', en: 'List my Outlook contacts.' },
      { de: 'Zeig Kontaktdaten zu [Name].', en: 'Show contact details for [name].' },
      {
        de: 'Wer sind unsere Ansprechpartner bei [Firma]?',
        en: 'Who are our contacts at [company]?',
      },
    ],
  },
  {
    id: 'onenote',
    nameDe: 'OneNote',
    nameEn: 'OneNote',
    questions: [
      {
        de: 'Liste meine OneNote-Notizbücher und Abschnitte.',
        en: 'List my OneNote notebooks and sections.',
      },
      {
        de: 'Zeig den Inhalt der Seite "[Seitentitel]".',
        en: 'Show the content of the page "[page title]".',
      },
      {
        de: 'Erstelle eine neue Seite in [Abschnitt].',
        en: 'Create a new page in [section].',
      },
    ],
  },
  {
    id: 'overview',
    nameDe: 'Woche & Überblick',
    nameEn: 'Week & Overview',
    questions: [
      {
        de: 'Was war meine Woche? (Zusammenfassung: Termine, E-Mails, Aufgaben.)',
        en: 'What did my week look like? (Summary: meetings, emails, tasks.)',
      },
      {
        de: 'Was braucht meine Aufmerksamkeit? (Flagged emails, überfällige Aufgaben.)',
        en: 'What needs my attention? (Flagged emails, overdue tasks.)',
      },
      {
        de: 'Gib mir einen Projekt-Überblick zu "[Projekt]". (Dateien, Termine, E-Mails, Aufgaben.)',
        en: 'Give me a project overview for "[project]". (Files, meetings, emails, tasks.)',
      },
    ],
  },
  {
    id: 'decisions',
    nameDe: 'Entscheidungen & Beziehungen',
    nameEn: 'Decisions & Relationships',
    questions: [
      {
        de: 'Welche Entscheidungen wurden zu "[Thema]" getroffen?',
        en: 'What decisions were made about "[topic]"?',
      },
      {
        de: 'Wie oft habe ich mit [Person] kommuniziert?',
        en: 'How often did I communicate with [person]?',
      },
      {
        de: 'Wer sind gemeinsame Kontakte von mir und [Person]?',
        en: 'Who are mutual connections between me and [person]?',
      },
    ],
  },
  {
    id: 'auth',
    nameDe: 'Authentifizierung',
    nameEn: 'Authentication',
    questions: [
      { de: 'Bin ich eingeloggt?', en: 'Am I logged in?' },
      { de: 'Mit welchem Konto bin ich verbunden?', en: 'Which account am I using?' },
    ],
  },
];
