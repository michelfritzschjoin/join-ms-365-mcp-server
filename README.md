<p align="center">
  <img src="https://img.shields.io/badge/Microsoft%20365-0078D4?style=for-the-badge&logo=microsoft&logoColor=white" alt="Microsoft 365">
  <img src="https://img.shields.io/badge/MCP%20Protocol-00A9CE?style=for-the-badge&logo=data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHZpZXdCb3g9IjAgMCAyNCAyNCI+PHBhdGggZmlsbD0id2hpdGUiIGQ9Ik0xMiAyTDIgN2wxMCA1IDEwLTV6Ii8+PC9zdmc+" alt="MCP">
  <img src="https://img.shields.io/badge/Docker-2496ED?style=for-the-badge&logo=docker&logoColor=white" alt="Docker">
  <img src="https://img.shields.io/badge/TypeScript-3178C6?style=for-the-badge&logo=typescript&logoColor=white" alt="TypeScript">
</p>

<h1 align="center">🚀 Join Microsoft 365 MCP Server</h1>

<p align="center">
  <strong>The Ultimate AI-Powered Gateway to Microsoft 365</strong><br>
  <strong>Der ultimative KI-gestützte Zugang zu Microsoft 365</strong>
</p>

<p align="center">
  A powerful Model Context Protocol (MCP) server enabling AI assistants to seamlessly interact with Microsoft 365 services through the Graph API.<br>
  Ein leistungsstarker Model Context Protocol (MCP) Server, der KI-Assistenten ermöglicht, nahtlos mit Microsoft 365-Diensten über die Graph API zu interagieren.
</p>

<p align="center">
  <a href="#-overview--übersicht">Overview</a> •
  <a href="#-quick-start--schnellstart">Quick Start</a> •
  <a href="#-features--funktionen">Features</a> •
  <a href="#-super-tools--super-tools">Super Tools</a> •
  <a href="#-tool-categories--tool-kategorien">Tools</a> •
  <a href="#-configuration--konfiguration">Configuration</a> •
  <a href="#-security--sicherheit">Security</a>
</p>

---

## 📋 Table of Contents / Inhaltsverzeichnis

- [Overview / Übersicht](#-overview--übersicht)
- [Quick Start / Schnellstart](#-quick-start--schnellstart)
- [Features / Funktionen](#-features--funktionen)
- [Super Tools / Super-Tools](#-super-tools--super-tools)
- [Tool Categories / Tool-Kategorien](#-tool-categories--tool-kategorien)
- [Intelligent Discovery System](#-intelligent-discovery-system)
- [Configuration / Konfiguration](#-configuration--konfiguration)
- [Authentication Methods / Authentifizierungsmethoden](#-authentication-methods--authentifizierungsmethoden)
- [Security & Compliance / Sicherheit & Compliance](#-security--compliance--sicherheit--compliance)
- [API Reference / API-Referenz](#-api-reference--api-referenz)

---

## 🌟 Overview / Übersicht

### English

The **Join Microsoft 365 MCP Server** transforms how AI assistants interact with Microsoft 365. Instead of simple API wrappers, it provides an **intelligent layer** that understands context, learns from usage patterns, and executes complex multi-step operations seamlessly.

**Key Innovations:**

- **Super Tools Mode**: Consolidates 126+ individual tools into 11 unified "Super-Tools" for easier LLM decision-making
- **Microsoft 365 Unified Search**: Primary search tool that searches across all M365 content and suggests specific tools to use next
- **Dual Timezone Display**: Shows both server local time and UTC for all calendar events and emails
- **Quick Summary Lists**: Comprehensive overview lists at the top of responses to ensure no item is overlooked
- **Intelligent Learning System**: Adapts and improves based on usage patterns
- **Read-Only Mode**: Safe exploration without write operations

### Deutsch

Der **Join Microsoft 365 MCP Server** revolutioniert die Art, wie KI-Assistenten mit Microsoft 365 interagieren. Statt einfacher API-Wrapper bietet er eine **intelligente Schicht**, die Kontext versteht, aus Nutzungsmustern lernt und komplexe Multi-Step-Operationen nahtlos ausführt.

**Wichtige Innovationen:**

- **Super-Tools-Modus**: Konsolidiert 126+ einzelne Tools zu 11 vereinheitlichten "Super-Tools" für einfachere LLM-Entscheidungen
- **Microsoft 365 Unified Search**: Primäres Suchtool, das alle M365-Inhalte durchsucht und spezifische Tools für die weitere Nutzung vorschlägt
- **Dual-Zeitzonen-Anzeige**: Zeigt sowohl Server-Lokalzeit als auch UTC für alle Kalendertermine und E-Mails
- **Schnellübersichtslisten**: Umfassende Übersichtslisten am Anfang von Antworten, damit kein Element übersehen wird
- **Intelligentes Lernsystem**: Passt sich an und verbessert sich basierend auf Nutzungsmustern
- **Read-Only-Modus**: Sichere Erkundung ohne Schreiboperationen

### Why Choose This Server? / Warum diesen Server wählen?

| Feature / Funktion                             | Traditional APIs / Traditionelle APIs             | Join MS365 MCP Server                                            |
| ---------------------------------------------- | ------------------------------------------------- | ---------------------------------------------------------------- |
| Context Understanding / Kontextverständnis     | ❌ None / Keines                                  | ✅ Deep semantic understanding / Tiefes semantisches Verständnis |
| Multi-step Operations / Multi-Step-Operationen | ❌ Manual orchestration / Manuelle Orchestrierung | ✅ Automatic chaining / Automatische Verkettung                  |
| Learning System / Lernsystem                   | ❌ Static / Statisch                              | ✅ Adaptive learning from usage / Adaptives Lernen aus Nutzung   |
| Natural Language / Natürliche Sprache          | ❌ Not supported / Nicht unterstützt              | ✅ Ask questions naturally / Fragen natürlich stellen            |
| Tool Consolidation / Tool-Konsolidierung       | ❌ 126+ individual tools / 126+ einzelne Tools    | ✅ 11 Super-Tools / 11 Super-Tools                               |
| Timezone Display / Zeitzonen-Anzeige           | ❌ Single timezone / Einzelne Zeitzone            | ✅ Local + UTC / Lokal + UTC                                     |
| Result Overview / Ergebnisübersicht            | ❌ Detailed only / Nur detailliert                | ✅ Quick summary + details / Schnellübersicht + Details          |

---

## ⚡ Quick Start / Schnellstart

### Prerequisites / Voraussetzungen

**English:**

- **Docker** and **Docker Compose**
- A Microsoft 365 account (personal, work, or school)
- Azure AD App Registration (for production use)

**Deutsch:**

- **Docker** und **Docker Compose**
- Ein Microsoft 365-Konto (privat, geschäftlich oder Schulkonto)
- Azure AD App-Registrierung (für Produktionseinsatz)

### 🐳 Docker Deployment / Docker-Bereitstellung

#### Option 1: Docker Compose (Recommended / Empfohlen)

```bash
# 1. Create configuration file / Konfigurationsdatei erstellen
cp stack.env.example stack.env

# 2. Configure your Azure AD credentials in stack.env
# Konfigurieren Sie Ihre Azure AD-Anmeldedaten in stack.env
nano stack.env

# 3. Start the server / Server starten
docker compose up -d

# For standalone mode (without Traefik):
# Für Standalone-Modus (ohne Traefik):
docker compose --profile standalone up -d
```

#### Option 2: Docker Run

```bash
# Pull the image / Image herunterladen
docker pull aijoin/join-ms-365-mcp-server:latest

# Run with environment variables / Mit Umgebungsvariablen ausführen
docker run -d \
  --name ms365-mcp \
  -p 3000:3000 \
  -e MS365_MCP_CLIENT_ID=your-client-id \
  -e MS365_MCP_TENANT_ID=your-tenant-id \
  -e MS365_MCP_USE_SUPER_TOOLS=true \
  -v ./data:/app/data \
  aijoin/join-ms-365-mcp-server:latest \
  --http 3000 -v
```

### MCP Client Integration / MCP-Client-Integration

**English:** Connect your AI assistant to the running server:

**Deutsch:** Verbinden Sie Ihren KI-Assistenten mit dem laufenden Server:

```json
{
  "mcpServers": {
    "ms365": {
      "url": "https://your-server.com/mcp",
      "transportType": "streamable-http"
    }
  }
}
```

### First Authentication / Erste Authentifizierung

**English:** Simply ask your AI assistant: _"Log me into Microsoft 365"_ - the server will guide you through device code authentication.

**Deutsch:** Fragen Sie einfach Ihren KI-Assistenten: _"Melde mich bei Microsoft 365 an"_ - der Server führt Sie durch die Device-Code-Authentifizierung.

---

## 🎯 Features / Funktionen

### Core Capabilities / Kernfunktionen

| Capability / Funktion            | Description / Beschreibung                                      |
| -------------------------------- | --------------------------------------------------------------- |
| **11 Super-Tools**               | Consolidated interface replacing 126+ individual tools          |
| **Microsoft 365 Unified Search** | Primary search tool across all M365 content                     |
| **90+ Individual Tools**         | Comprehensive coverage of Microsoft 365 services (classic mode) |
| **Intelligent Search**           | Cross-product search with semantic understanding                |
| **Deep Research**                | Multi-step reasoning for complex questions                      |
| **Learning System**              | Improves over time based on usage patterns                      |
| **Dual Timezone Display**        | Server local time + UTC for all dates/times                     |
| **Quick Summary Lists**          | Overview lists to ensure nothing is missed                      |
| **Download Links**               | Generate direct download links for files                        |
| **Microsoft Loop Support**       | Loop file detection and content extraction                      |
| **Read-Only Mode**               | Safe exploration without write operations                       |
| **Preset Filtering**             | Load only the tools you need                                    |

### Supported Microsoft 365 Services / Unterstützte Microsoft 365-Dienste

<table>
<tr>
<td width="25%">

**📧 Outlook**

- Email management / E-Mail-Verwaltung
- Folder organization / Ordnerorganisation
- Attachments / Anhänge
- Drafts / Entwürfe

</td>
<td width="25%">

**📅 Calendar**

- Events & meetings / Termine & Besprechungen
- Scheduling / Terminplanung
- Recurring events / Wiederkehrende Termine
- Meeting times / Besprechungszeiten

</td>
<td width="25%">

**📁 OneDrive**

- File operations / Dateioperationen
- Folder management / Ordnerverwaltung
- Sharing / Freigabe
- Download/Upload

</td>
<td width="25%">

**💬 Teams**

- Chats & messages / Chats & Nachrichten
- Channels / Kanäle
- Team management / Teamverwaltung
- Transcripts / Transkripte

</td>
</tr>
<tr>
<td>

**🌐 SharePoint**

- Sites & lists / Websites & Listen
- Document libraries / Dokumentbibliotheken
- Site search / Websitesuche
- Permissions / Berechtigungen

</td>
<td>

**📊 Excel**

- Worksheet operations / Arbeitsblattoperationen
- Range manipulation / Bereichsmanipulation
- Charts / Diagramme
- Formatting / Formatierung

</td>
<td>

**✅ Tasks**

- To-Do lists / Aufgabenlisten
- Planner tasks / Planner-Aufgaben
- Task assignment / Aufgabenverteilung
- Due dates / Fälligkeitsdaten

</td>
<td>

**📔 OneNote**

- Notebooks
- Sections & pages / Abschnitte & Seiten
- Content creation / Inhaltserstellung
- Search / Suche

</td>
<td>

**🔄 Microsoft Loop**

- Loop file detection / Loop-Datei-Erkennung
- Collaborative documents / Kollaborative Dokumente
- Content extraction / Inhalts-Extraktion
- Fluid format parsing / Fluid-Format-Parsing

</td>
</tr>
</table>

---

## 🚀 Super Tools / Super-Tools

### English

**Super-Tools Mode** consolidates 126+ individual tools into 11 unified "Super-Tools". Each Super-Tool accepts an `action` parameter to specify the operation, making it much easier for LLMs to choose the right tool.

**Enable Super-Tools Mode:**

```bash
# Via environment variable / Über Umgebungsvariable
MS365_MCP_USE_SUPER_TOOLS=true

# Or via Docker / Oder über Docker
docker run -d \
  -e MS365_MCP_USE_SUPER_TOOLS=true \
  aijoin/join-ms-365-mcp-server:latest \
  --http 3000
```

### Deutsch

**Super-Tools-Modus** konsolidiert 126+ einzelne Tools zu 11 vereinheitlichten "Super-Tools". Jedes Super-Tool akzeptiert einen `action`-Parameter zur Spezifikation der Operation, was es für LLMs viel einfacher macht, das richtige Tool zu wählen.

**Super-Tools-Modus aktivieren:**

```bash
# Über Umgebungsvariable
MS365_MCP_USE_SUPER_TOOLS=true

# Oder über Docker
docker run -d \
  -e MS365_MCP_USE_SUPER_TOOLS=true \
  aijoin/join-ms-365-mcp-server:latest \
  --http 3000
```

### Super-Tools List / Super-Tools-Liste

| #     | Tool         | Description / Beschreibung                                                                                                                                      |
| ----- | ------------ | --------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| **0** | `search`     | 🔍 **PRIMARY** - Microsoft 365 Unified Search across emails, calendar, files, SharePoint, Teams. Returns results and suggests which specific tools to use next. |
| 1     | `email`      | 📧 Unified email operations: list, get, folders, attachments, search, send, reply, delete, move                                                                 |
| 2     | `calendar`   | 📅 Calendar operations: list, get, view, calendars, create-event, update-event, delete-event                                                                    |
| 3     | `teams`      | 💬 Teams, Channels, Chats: list-teams, get-team, channels, channel-messages, chats, chat-messages                                                               |
| 4     | `files`      | 📁 OneDrive files: drives, list, get, download, search, root                                                                                                    |
| 5     | `tasks`      | ✅ To-Do & Planner: todo-lists, todo-tasks, planner-tasks, create-todo, update-todo, delete-todo                                                                |
| 6     | `contacts`   | 👥 Contacts & Users: list-contacts, get-contact, list-users, current-user                                                                                       |
| 7     | `meetings`   | 🎥 Online Meetings: list-meetings, get-meeting, transcripts, recordings                                                                                         |
| 8     | `sharepoint` | 🌐 SharePoint: search-sites, get-site, site-drives, site-lists                                                                                                  |
| 9     | `notes`      | 📔 OneNote: notebooks, sections, pages, page-content, search-pages                                                                                              |
| 10    | `assistant`  | 🤖 Smart operations: ask, search, my-day, my-week, person-info, project-overview, follow-ups                                                                    |

### Example Usage / Beispielverwendung

**English:**

```json
{
  "tool": "search",
  "arguments": {
    "query": "Project Alpha meeting notes",
    "entityTypes": ["message", "event", "driveItem"],
    "size": 10
  }
}
```

**Deutsch:**

```json
{
  "tool": "search",
  "arguments": {
    "query": "Projekt Alpha Besprechungsnotizen",
    "entityTypes": ["message", "event", "driveItem"],
    "size": 10
  }
}
```

### Read-Only Mode Support / Read-Only-Modus-Unterstützung

**English:** All Super-Tools respect the `READ_ONLY` environment variable. Write operations (send, create, update, delete) are automatically blocked with clear error messages when read-only mode is enabled.

**Deutsch:** Alle Super-Tools respektieren die `READ_ONLY` Umgebungsvariable. Schreiboperationen (send, create, update, delete) werden automatisch blockiert mit klaren Fehlermeldungen, wenn der Read-Only-Modus aktiviert ist.

```bash
# Enable read-only mode / Read-Only-Modus aktivieren
READ_ONLY=1
# or / oder
MS365_MCP_READ_ONLY=true
```

---

## 🕐 Dual Timezone Display / Dual-Zeitzonen-Anzeige

### English

All calendar events and emails now display **both server local time and UTC** for easy reference:

```
⏰ 10:30 (UTC: 09:30)
```

**Features:**

- Server local time (primary display)
- UTC time (for reference)
- Combined display format: `HH:MM (UTC: HH:MM)`
- ISO 8601 UTC timestamps in structured data

### Deutsch

Alle Kalendertermine und E-Mails zeigen jetzt **sowohl Server-Lokalzeit als auch UTC** zur einfachen Referenz:

```
⏰ 10:30 (UTC: 09:30)
```

**Funktionen:**

- Server-Lokalzeit (primäre Anzeige)
- UTC-Zeit (zur Referenz)
- Kombiniertes Anzeigeformat: `HH:MM (UTC: HH:MM)`
- ISO 8601 UTC-Zeitstempel in strukturierten Daten

### Example Output / Beispielausgabe

**Calendar Event / Kalendertermin:**

```
📅 Montag, 28.01.2026 (1 Termin)
──────────────────────────────────────────────────
💻 Kickoff-Briefing
   ⏰ 10:30 (UTC: 09:30) - 11:30 (UTC: 10:30) (1h)
   📍 Conference Room A
```

**Email / E-Mail:**

```
📬 Project Update
   ⏰ 09:15 (UTC: 08:15)
   👤 Von: Max Müller <max@example.com>
```

---

## 📋 Quick Summary Lists / Schnellübersichtslisten

### English

To ensure **no calendar event or email is overlooked**, all responses now include a **Quick Summary List** at the top, followed by the detailed view.

**Calendar Quick Summary:**

```
📋 SCHNELLÜBERSICHT ALLER TERMINE:
────────────────────────────────────────────────────────────
1. 💻 28.01.2026 10:30 (UTC: 09:30) | Kickoff-Briefing
2. 📍 28.01.2026 14:00 (UTC: 13:00) | Team Meeting
3. 💻 28.01.2026 16:30 (UTC: 15:30) | Client Call

═══════════════════════════════════════════════════════════
📖 DETAILANSICHT:
═══════════════════════════════════════════════════════════
```

**Email Quick Summary:**

```
📋 SCHNELLÜBERSICHT ALLER E-MAILS:
────────────────────────────────────────────────────────────
1. 📬📎 28.01.2026 09:15 (UTC: 08:15) | Max Müller | Projekt Update...
2. 📭 27.01.2026 18:30 (UTC: 17:30) | Anna Schmidt | Meeting Notizen

═══════════════════════════════════════════════════════════
📖 DETAILANSICHT:
═══════════════════════════════════════════════════════════
```

### Deutsch

Um sicherzustellen, dass **kein Kalendertermin oder E-Mail übersehen wird**, enthalten alle Antworten jetzt eine **Schnellübersichtsliste** am Anfang, gefolgt von der Detailansicht.

**Kalender-Schnellübersicht:**

```
📋 SCHNELLÜBERSICHT ALLER TERMINE:
────────────────────────────────────────────────────────────
1. 💻 28.01.2026 10:30 (UTC: 09:30) | Kickoff-Briefing
2. 📍 28.01.2026 14:00 (UTC: 13:00) | Team Meeting
3. 💻 28.01.2026 16:30 (UTC: 15:30) | Client Call

═══════════════════════════════════════════════════════════
📖 DETAILANSICHT:
═══════════════════════════════════════════════════════════
```

**E-Mail-Schnellübersicht:**

```
📋 SCHNELLÜBERSICHT ALLER E-MAILS:
────────────────────────────────────────────────────────────
1. 📬📎 28.01.2026 09:15 (UTC: 08:15) | Max Müller | Projekt Update...
2. 📭 27.01.2026 18:30 (UTC: 17:30) | Anna Schmidt | Meeting Notizen

═══════════════════════════════════════════════════════════
📖 DETAILANSICHT:
═══════════════════════════════════════════════════════════
```

---

## 🛠 Tool Categories / Tool-Kategorien

### 🔐 Authentication Tools / Authentifizierungs-Tools

Secure authentication with multi-account support / Sichere Authentifizierung mit Multi-Account-Unterstützung.

| Tool             | Description / Beschreibung                                                  | Notes / Hinweise                                                           |
| ---------------- | --------------------------------------------------------------------------- | -------------------------------------------------------------------------- |
| `login`          | Authenticate via device code flow / Authentifizierung über Device-Code-Flow | Required before using other tools / Erforderlich vor Nutzung anderer Tools |
| `logout`         | Log out from Microsoft account / Von Microsoft-Konto abmelden               | Clears cached tokens / Löscht gecachte Tokens                              |
| `verify-login`   | Check authentication status / Authentifizierungsstatus prüfen               | Non-interactive verification / Nicht-interaktive Verifizierung             |
| `list-accounts`  | List cached Microsoft accounts / Gecachte Microsoft-Konten auflisten        | Multi-account support / Multi-Account-Unterstützung                        |
| `select-account` | Switch between accounts / Zwischen Konten wechseln                          | Seamless account switching / Nahtloser Kontenwechsel                       |
| `remove-account` | Remove account from cache / Konto aus Cache entfernen                       | Clean up stored credentials / Gespeicherte Anmeldedaten bereinigen         |

### 📧 Email & Communication Tools / E-Mail- & Kommunikations-Tools

Complete email management and shared mailbox support / Vollständige E-Mail-Verwaltung und Shared-Mailbox-Unterstützung.

#### Personal Email / Persönliche E-Mail

| Tool                        | Description / Beschreibung                                   | Parameters                             |
| --------------------------- | ------------------------------------------------------------ | -------------------------------------- |
| `list-mail-messages`        | List emails with filtering / E-Mails mit Filterung auflisten | `top`, `filter`, `search`, `orderby`   |
| `get-mail-message`          | Get full email content / Vollständigen E-Mail-Inhalt abrufen | `messageId`                            |
| `send-mail`                 | Send new email / Neue E-Mail senden                          | `to`, `subject`, `body`, `attachments` |
| `create-draft-email`        | Create email draft / E-Mail-Entwurf erstellen                | `to`, `subject`, `body`                |
| `delete-mail-message`       | Delete email / E-Mail löschen                                | `messageId`                            |
| `move-mail-message`         | Move email to folder / E-Mail in Ordner verschieben          | `messageId`, `folderId`                |
| `list-mail-folders`         | List all mail folders / Alle E-Mail-Ordner auflisten         | -                                      |
| `list-mail-folder-messages` | List messages in folder / Nachrichten in Ordner auflisten    | `folderId`                             |

#### Super-Tool: `email`

**English:** Unified email operations with action-based interface:

```json
{
  "tool": "email",
  "arguments": {
    "action": "list",
    "top": 25,
    "search": "Project Alpha"
  }
}
```

**Available actions:** `list`, `get`, `folders`, `child-folders`, `attachments`, `search`, `send`, `reply`, `delete`, `move`

**Deutsch:** Vereinheitlichte E-Mail-Operationen mit aktionsbasierter Schnittstelle:

```json
{
  "tool": "email",
  "arguments": {
    "action": "list",
    "top": 25,
    "search": "Projekt Alpha"
  }
}
```

**Verfügbare Aktionen:** `list`, `get`, `folders`, `attachments`, `search`, `send`, `reply`, `delete`, `move`

### 📅 Calendar Tools / Kalender-Tools

Full calendar management with meeting scheduling / Vollständige Kalenderverwaltung mit Besprechungsplanung.

| Tool                    | Description / Beschreibung                         | Parameters                             |
| ----------------------- | -------------------------------------------------- | -------------------------------------- |
| `list-calendars`        | List all calendars / Alle Kalender auflisten       | -                                      |
| `list-calendar-events`  | List events / Termine auflisten                    | `top`, `filter`, `orderby`             |
| `get-calendar-event`    | Get event details / Termindetails abrufen          | `eventId`                              |
| `create-calendar-event` | Create new event / Neuen Termin erstellen          | `subject`, `start`, `end`, `attendees` |
| `update-calendar-event` | Update event / Termin aktualisieren                | `eventId`, `updates`                   |
| `delete-calendar-event` | Delete event / Termin löschen                      | `eventId`                              |
| `get-calendar-view`     | Get calendar view / Kalenderansicht abrufen        | `startDateTime`, `endDateTime`         |
| `find-meeting-times`    | Find available slots / Verfügbare Zeitslots finden | `attendees`, `duration`                |

#### Super-Tool: `calendar`

**English:** Unified calendar operations:

```json
{
  "tool": "calendar",
  "arguments": {
    "action": "view",
    "startDateTime": "2026-01-28T00:00:00Z",
    "endDateTime": "2026-01-29T00:00:00Z"
  }
}
```

**Available actions:** `list`, `get`, `view`, `calendars`, `specific-calendar`, `create-event`, `update-event`, `delete-event`

**Deutsch:** Vereinheitlichte Kalender-Operationen:

```json
{
  "tool": "calendar",
  "arguments": {
    "action": "view",
    "startDateTime": "2026-01-28T00:00:00Z",
    "endDateTime": "2026-01-29T00:00:00Z"
  }
}
```

**Verfügbare Aktionen:** `list`, `get`, `view`, `calendars`, `specific-calendar`, `create-event`, `update-event`, `delete-event`

### 📁 File & Drive Tools / Datei- & Laufwerk-Tools

OneDrive file management with upload/download capabilities / OneDrive-Dateiverwaltung mit Upload/Download-Funktionen.

| Tool                             | Description / Beschreibung                             | Parameters                        |
| -------------------------------- | ------------------------------------------------------ | --------------------------------- |
| `list-drives`                    | List available drives / Verfügbare Laufwerke auflisten | -                                 |
| `get-drive-root-item`            | Get drive root folder / Laufwerks-Stammordner abrufen  | `driveId`                         |
| `list-folder-files`              | List files in folder / Dateien im Ordner auflisten     | `folderId`, `top`                 |
| `download-onedrive-file-content` | Download file content / Dateiinhalt herunterladen      | `itemId`                          |
| `upload-file-content`            | Update file content / Dateiinhalt aktualisieren        | `itemId`, `content`               |
| `upload-new-file`                | Upload new file / Neue Datei hochladen                 | `folderId`, `fileName`, `content` |
| `delete-onedrive-file`           | Delete file / Datei löschen                            | `itemId`                          |

### 💬 Microsoft Teams Tools / Microsoft Teams-Tools

> **Note / Hinweis:** Requires `--org-mode` flag (work/school accounts only) / Erfordert `--org-mode` Flag (nur Geschäfts-/Schulkonten)

| Tool                    | Description / Beschreibung                                | Parameters                       |
| ----------------------- | --------------------------------------------------------- | -------------------------------- |
| `list-chats`            | List all chats / Alle Chats auflisten                     | `top`                            |
| `get-chat`              | Get chat details / Chat-Details abrufen                   | `chatId`                         |
| `list-chat-messages`    | List messages in chat / Nachrichten im Chat auflisten     | `chatId`, `top`                  |
| `send-chat-message`     | Send chat message / Chat-Nachricht senden                 | `chatId`, `content`              |
| `list-joined-teams`     | List teams you're in / Teams auflisten, in denen Sie sind | -                                |
| `list-team-channels`    | List team channels / Team-Kanäle auflisten                | `teamId`                         |
| `list-channel-messages` | List channel messages / Kanalnachrichten auflisten        | `teamId`, `channelId`            |
| `send-channel-message`  | Send channel message / Kanalnachricht senden              | `teamId`, `channelId`, `content` |

### 🔍 Search & Discovery Tools / Such- & Discovery-Tools

Powerful cross-product search capabilities / Leistungsstarke produktübergreifende Suchfunktionen.

#### Super-Tool: `search` (PRIMARY / PRIMÄR)

**English:** The **recommended first tool** for exploring Microsoft 365 content. Searches across emails, calendar, files, SharePoint, Teams, and suggests which specific tools to use next.

**Deutsch:** Das **empfohlene erste Tool** zur Erkundung von Microsoft 365-Inhalten. Durchsucht E-Mails, Kalender, Dateien, SharePoint, Teams und schlägt vor, welche spezifischen Tools als nächstes zu verwenden sind.

```json
{
  "tool": "search",
  "arguments": {
    "query": "Project Alpha meeting notes",
    "entityTypes": ["message", "event", "driveItem"],
    "size": 10
  }
}
```

**Entity Types:** `message`, `event`, `driveItem`, `site`, `list`, `listItem`, `chatMessage`, `person`

**Response includes:**

- Search results grouped by entity type
- Tool suggestions for next steps
- Total hits count
- Formatted results with metadata

### 🧠 Intelligent Compound Tools / Intelligente Verbund-Tools

These **intelligent tools** automatically chain multiple API calls to answer complex contextual questions / Diese **intelligenten Tools** verkettet automatisch mehrere API-Aufrufe, um komplexe kontextuelle Fragen zu beantworten.

| Tool                        | What It Does / Was es tut                                                          | Example Query / Beispielabfrage                                           |
| --------------------------- | ---------------------------------------------------------------------------------- | ------------------------------------------------------------------------- |
| `find-messages-with-person` | Find all Teams chats with a person / Findet alle Teams-Chats mit einer Person      | "What did I discuss with John?" / "Worüber habe ich mit John gesprochen?" |
| `find-emails-with-person`   | Find all email threads with a person / Findet alle E-Mail-Threads mit einer Person | "Show emails from Sarah" / "Zeige E-Mails von Sarah"                      |
| `find-meetings-with-person` | Find past & future meetings / Findet vergangene & zukünftige Besprechungen         | "When did I meet with Mike?" / "Wann habe ich mich mit Mike getroffen?"   |
| `discover-project`          | Find all project-related content / Findet alle projektbezogenen Inhalte            | "Everything about Project Apollo" / "Alles über Projekt Apollo"           |
| `discover-person`           | Comprehensive person profile / Umfassendes Personenprofil                          | "Who is John Smith?" / "Wer ist John Smith?"                              |
| `get-my-week-summary`       | Weekly productivity digest / Wöchentliche Produktivitätszusammenfassung            | "What did I accomplish this week?" / "Was habe ich diese Woche erreicht?" |

---

## 🔬 Intelligent Discovery System

### English

The server features a sophisticated **Search-First Strategy** with multiple intelligent components:

```
┌─────────────────────────────────────────────────────────────────┐
│                    User Question                                 │
└─────────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────────┐
│                    NLP Enhancer                                  │
│  • Entity extraction  • Intent classification                   │
│  • Synonym expansion  • Query refinement                        │
└─────────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────────┐
│                Search-First Strategy                             │
│  • Microsoft Search API (emails, files, chats, events)          │
│  • Learning-informed entity type selection                      │
│  • Automatic query refinement if no results                     │
└─────────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────────┐
│                 Entity Extractor                                 │
│  • Identifies sites, teams, users, files                        │
│  • Extracts relevant keywords                                   │
│  • Maps to specific product queries                             │
└─────────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────────┐
│                 Data Aggregator                                  │
│  • Deduplication  • Relevance sorting                           │
│  • LLM-optimized formatting  • Source tracking                  │
└─────────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────────┐
│                 Learning System                                  │
│  • Records successful patterns  • Updates confidence            │
│  • Learns entity type preferences  • User feedback              │
└─────────────────────────────────────────────────────────────────┘
```

### Deutsch

Der Server verfügt über eine ausgeklügelte **Search-First-Strategie** mit mehreren intelligenten Komponenten:

```
┌─────────────────────────────────────────────────────────────────┐
│                    Benutzerfrage                                 │
└─────────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────────┐
│                    NLP Enhancer                                  │
│  • Entitäts-Extraktion  • Intent-Klassifizierung               │
│  • Synonym-Erweiterung  • Abfrage-Verfeinerung                  │
└─────────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────────┐
│                Search-First-Strategie                           │
│  • Microsoft Search API (E-Mails, Dateien, Chats, Termine)      │
│  • Lernbasierte Entitätstyp-Auswahl                             │
│  • Automatische Abfrage-Verfeinerung bei keinen Ergebnissen     │
└─────────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────────┐
│                 Entity Extractor                                 │
│  • Identifiziert Websites, Teams, Benutzer, Dateien             │
│  • Extrahiert relevante Schlüsselwörter                         │
│  • Mappt auf spezifische Produktabfragen                       │
└─────────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────────┐
│                 Data Aggregator                                  │
│  • Deduplizierung  • Relevanz-Sortierung                        │
│  • LLM-optimierte Formatierung  • Quellen-Tracking              │
└─────────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────────┐
│                 Lernsystem                                       │
│  • Zeichnet erfolgreiche Muster auf  • Aktualisiert Konfidenz  │
│  • Lernt Entitätstyp-Präferenzen  • Benutzer-Feedback          │
└─────────────────────────────────────────────────────────────────┘
```

---

## ⚙ Configuration / Konfiguration

### Docker Command Options / Docker-Befehlsoptionen

| Option                      | Description / Beschreibung                                                                       | Example / Beispiel                 |
| --------------------------- | ------------------------------------------------------------------------------------------------ | ---------------------------------- |
| `--org-mode`                | Enable organization mode (Teams, SharePoint) / Organisationsmodus aktivieren (Teams, SharePoint) | `--org-mode`                       |
| `--read-only`               | Disable write operations / Schreiboperationen deaktivieren                                       | `--read-only`                      |
| `--http [port]`             | Start HTTP server (default: 3000) / HTTP-Server starten (Standard: 3000)                         | `--http 8080`                      |
| `--preset <name>`           | Load specific tool presets / Spezifische Tool-Presets laden                                      | `--preset mail,calendar`           |
| `--enabled-tools <pattern>` | Filter tools by regex / Tools nach Regex filtern                                                 | `--enabled-tools "excel\|contact"` |
| `--toon`                    | Enable TOON format (30-60% token savings) / TOON-Format aktivieren (30-60% Token-Ersparnis)      | `--toon`                           |
| `--discovery`               | Start with discovery tools only / Nur mit Discovery-Tools starten                                | `--discovery`                      |
| `--cloud <type>`            | Cloud environment (global/china) / Cloud-Umgebung (global/china)                                 | `--cloud china`                    |
| `-v`                        | Enable verbose logging / Ausführliches Logging aktivieren                                        | `-v`                               |

### Environment Variables / Umgebungsvariablen

| Variable                    | Description / Beschreibung                                                      | Default                     |
| --------------------------- | ------------------------------------------------------------------------------- | --------------------------- |
| `MS365_MCP_CLIENT_ID`       | Azure AD app client ID                                                          | **Required / Erforderlich** |
| `MS365_MCP_TENANT_ID`       | Azure AD tenant ID                                                              | `common`                    |
| `MS365_MCP_CLIENT_SECRET`   | Client secret (confidential apps) / Client-Geheimnis (vertrauliche Apps)        | -                           |
| `MS365_MCP_USE_SUPER_TOOLS` | Enable Super-Tools mode / Super-Tools-Modus aktivieren                          | `false`                     |
| `MS365_MCP_ORG_MODE`        | Enable organization mode / Organisationsmodus aktivieren                        | `false`                     |
| `MS365_MCP_OUTPUT_FORMAT`   | Output format (`json`/`toon`) / Ausgabeformat (`json`/`toon`)                   | `json`                      |
| `MS365_MCP_CLOUD_TYPE`      | Cloud environment / Cloud-Umgebung                                              | `global`                    |
| `MS365_MCP_KEYVAULT_URL`    | Azure Key Vault URL                                                             | -                           |
| `MS365_MCP_MAX_RESULTS`     | Maximum search results / Maximale Suchergebnisse                                | `500`                       |
| `MS365_MCP_ANONYMIZE_PII`   | Anonymize PII in knowledge base storage / PII in Wissensdatenbank anonymisieren | `true`                      |
| `READ_ONLY`                 | Enable read-only mode / Read-Only-Modus aktivieren                              | `false`                     |

> **Security Warning / Sicherheitswarnung**: Setting `MS365_MCP_ANONYMIZE_PII=false` disables the automatic removal of personally identifiable information (email addresses, phone numbers, IDs, etc.) before storing data in the knowledge base. This is **NOT recommended in production** and may violate GDPR/DSGVO compliance. Only disable for development or debugging purposes. / Das Setzen von `MS365_MCP_ANONYMIZE_PII=false` deaktiviert die automatische Entfernung von personenbezogenen Daten (E-Mail-Adressen, Telefonnummern, IDs, etc.) vor der Speicherung in der Wissensdatenbank. Dies wird **in der Produktion NICHT empfohlen** und kann gegen DSGVO-Compliance verstoßen. Nur für Entwicklung oder Debugging deaktivieren.
> | `LOG_LEVEL` | Logging level / Logging-Level | `info` |
> | `SILENT` | Disable console output / Konsolenausgabe deaktivieren | `false` |

### Docker Run Examples / Docker-Run-Beispiele

```bash
# Basic HTTP server / Grundlegender HTTP-Server
docker run -d -p 3000:3000 aijoin/join-ms-365-mcp-server:latest --http 3000

# Organization mode with Super-Tools / Organisationsmodus mit Super-Tools
docker run -d -p 3000:3000 \
  -e MS365_MCP_CLIENT_ID=your-client-id \
  -e MS365_MCP_TENANT_ID=your-tenant-id \
  -e MS365_MCP_USE_SUPER_TOOLS=true \
  aijoin/join-ms-365-mcp-server:latest \
  --http 3000 --org-mode -v

# Read-only mode with Super-Tools / Read-Only-Modus mit Super-Tools
docker run -d -p 3000:3000 \
  -e MS365_MCP_USE_SUPER_TOOLS=true \
  -e READ_ONLY=1 \
  aijoin/join-ms-365-mcp-server:latest \
  --http 3000
```

---

## 🔑 Authentication Methods / Authentifizierungsmethoden

### 1. Device Code Flow (Default / Standard)

**English:** Interactive authentication for users:

1. Call the `login` tool
2. Visit the provided URL and enter the code
3. Call `verify-login` to confirm

**Deutsch:** Interaktive Authentifizierung für Benutzer:

1. Rufen Sie das `login` Tool auf
2. Besuchen Sie die bereitgestellte URL und geben Sie den Code ein
3. Rufen Sie `verify-login` auf, um zu bestätigen

### 2. OAuth Authorization Code Flow (HTTP Mode)

**English:** For web applications and remote servers:

- Exposes OAuth endpoints at `/auth/*`
- Requires `Authorization: Bearer <token>` for MCP requests
- Supports MCP OAuth 2.1 with Dynamic Client Registration

**Deutsch:** Für Webanwendungen und Remote-Server:

- Stellt OAuth-Endpunkte unter `/auth/*` bereit
- Erfordert `Authorization: Bearer <token>` für MCP-Anfragen
- Unterstützt MCP OAuth 2.1 mit Dynamic Client Registration

### 3. Bring Your Own Token (BYOT)

**English:** For integration with existing OAuth systems:

```bash
docker run -d -p 3000:3000 \
  -e MS365_MCP_OAUTH_TOKEN=your_token \
  aijoin/join-ms-365-mcp-server:latest \
  --http 3000
```

**Deutsch:** Für die Integration mit bestehenden OAuth-Systemen:

```bash
docker run -d -p 3000:3000 \
  -e MS365_MCP_OAUTH_TOKEN=your_token \
  aijoin/join-ms-365-mcp-server:latest \
  --http 3000
```

---

## 🔒 Security & Compliance / Sicherheit & Compliance

### Security Features / Sicherheitsfunktionen

- ✅ **OAuth 2.1 / PKCE** - Secure token handling / Sichere Token-Verwaltung
- ✅ **Token validation** - Verified against Microsoft Graph / Gegen Microsoft Graph verifiziert
- ✅ **Secure storage** - Persistent volume for credentials / Persistenter Datenträger für Anmeldedaten
- ✅ **Read-only mode** - Safe exploration without modifications / Sichere Erkundung ohne Änderungen
- ✅ **Input validation** - Zod schema validation on all inputs / Zod-Schema-Validierung für alle Eingaben
- ✅ **Rate limiting** - Configurable request limits / Konfigurierbare Anfragelimits
- ✅ **HTTPS/TLS** - Traefik integration for production / Traefik-Integration für Produktion

### Compliance

- **ISO 27001** - Information security management / Informationssicherheitsmanagement
- **GDPR/DSGVO** - Data protection by design / Datenschutz durch Design
- **OWASP** - Security best practices / Sicherheitsbest Practices

---

## 📊 Query Dashboard / Abfrage-Dashboard

### English

The Query Dashboard provides a secure web interface to view and analyze all user queries. This feature enables auditing, analytics, and debugging of MCP tool usage.

**Features:**

- 📈 **Real-time Statistics** - Total queries, unique users, success rates
- 🔍 **Query Search & Filtering** - Filter by tool, user, date, status
- 📉 **Hourly Activity Charts** - Visual query distribution over 24 hours
- 🔒 **Password Protected** - Secure access via environment variable
- 📥 **GDPR Data Export** - Export user data for data portability
- 🗑️ **GDPR Erasure** - Delete user data (Right to be Forgotten)

### Deutsch

Das Query Dashboard bietet eine sichere Weboberfläche zur Anzeige und Analyse aller Benutzerabfragen. Diese Funktion ermöglicht Auditierung, Analysen und Debugging der MCP-Tool-Nutzung.

**Funktionen:**

- 📈 **Echtzeit-Statistiken** - Gesamtabfragen, eindeutige Benutzer, Erfolgsraten
- 🔍 **Abfrage-Suche & Filterung** - Nach Tool, Benutzer, Datum, Status filtern
- 📉 **Stündliche Aktivitätsdiagramme** - Visuelle Abfrageverteilung über 24 Stunden
- 🔒 **Passwortgeschützt** - Sicherer Zugriff über Umgebungsvariable
- 📥 **DSGVO-Datenexport** - Benutzerdaten für Datenportabilität exportieren
- 🗑️ **DSGVO-Löschung** - Benutzerdaten löschen (Recht auf Vergessenwerden)

### Enabling the Dashboard / Dashboard aktivieren

```bash
# In stack.env or docker-compose environment
# In stack.env oder docker-compose Umgebung
DASHBOARD_PASSWORD=your-secure-password-here
```

Access the dashboard at: `https://your-server.com/dashboard` / Zugriff auf das Dashboard unter: `https://your-server.com/dashboard`

---

## 📚 API Reference / API-Referenz

### MCP Client Configuration / MCP-Client-Konfiguration

#### OpenWebUI / Remote Clients

```json
{
  "mcpServers": {
    "ms365": {
      "url": "https://your-server.com/mcp",
      "transportType": "streamable-http"
    }
  }
}
```

#### Local Development / Lokale Entwicklung

```json
{
  "mcpServers": {
    "ms365": {
      "url": "http://localhost:3000/mcp",
      "transportType": "streamable-http"
    }
  }
}
```

### Tool Response Format / Tool-Antwortformat

All tools return MCP-compliant responses / Alle Tools geben MCP-konforme Antworten zurück:

```typescript
interface McpToolResult {
  content: Array<{
    type: 'text' | 'image' | 'resource';
    text?: string;
    data?: string;
    mimeType?: string;
    uri?: string;
  }>;
  isError?: boolean;
}
```

---

## 🐳 Docker Deployment / Docker-Bereitstellung

### Production with Traefik / Produktion mit Traefik

```yaml
# docker-compose.yml
services:
  ms365-mcp-server:
    image: aijoin/join-ms-365-mcp-server:latest
    container_name: ms365-mcp
    restart: unless-stopped
    env_file:
      - stack.env
    command: ['--http', '3000', '-v']
    volumes:
      - ./data:/app/data
    labels:
      - 'traefik.enable=true'
      - 'traefik.http.routers.ms365-mcp.rule=Host(`ms365-mcp.yourdomain.com`)'
      - 'traefik.http.routers.ms365-mcp.entrypoints=websecure'
      - 'traefik.http.routers.ms365-mcp.tls.certresolver=myresolver'
    networks:
      - web

networks:
  web:
    external: true
```

### Build from Source / Aus Quellcode erstellen

```bash
# Clone the repository / Repository klonen
git clone https://github.com/michelfritzschjoin/join-ms-365-mcp-server.git
cd join-ms-365-mcp-server

# Build the image / Image erstellen
docker build -t ms365-mcp-server .

# Run / Ausführen
docker run -p 3000:3000 ms365-mcp-server --http 3000
```

---

## 🤝 Contributing / Beitragen

**English:**

1. Fork the repository
2. Run `npm install`
3. Generate client: `npm run generate`
4. Make changes
5. Run verification: `npm run verify`
6. Submit PR

**Deutsch:**

1. Repository forken
2. `npm install` ausführen
3. Client generieren: `npm run generate`
4. Änderungen vornehmen
5. Verifizierung ausführen: `npm run verify`
6. PR einreichen

---

## 📄 License / Lizenz

All Rights Reserved © 2026 Join GmbH

---

## 📞 Support / Support

- 📋 [Issues](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/issues)

---

## ℹ️ Version Information / Versionsinformationen

**English:** The version is automatically read from `package.json` at runtime. The displayed version in the banner and CLI is always synchronized with the package version - no manual updates required.

**Deutsch:** Die Version wird automatisch zur Laufzeit aus `package.json` gelesen. Die angezeigte Version im Banner und CLI ist immer mit der Paketversion synchronisiert - keine manuellen Updates erforderlich.

---

<p align="center">
  <strong>Built with ❤️ by Join GmbH</strong><br>
  <strong>Mit ❤️ erstellt von Join GmbH</strong>
</p>
