# EZSheets

**EZSheets** is a collaborative spreadsheet with cloud sync, per-user permissions, sheet chat, snapshots, comments, import/export, and code-based sharing. _(or a Live Excel in a nutshell)_

The goal of the plugin is simple: allow multiple people to work on the same spreadsheet in real time, inside the game, without relying on an external tool for day-to-day tasks, venue organization, lists, internal tracking, logs, or general use.

---

## What the plugin does

EZSheets allows you to:

- create cloud-synced spreadsheets
- open existing spreadsheets from your account
- join spreadsheets through a shared code
- edit cells, tabs, and sheet structure
- apply permissions per user and per role preset
- use an integrated sheet chat
- use cell comments
- create snapshots for history and restoration
- favorite spreadsheets
- import and export spreadsheets
- track users seen in the sheet
- work with shared codes and temporary codes
- manage access, blocks, and member removal

In short, the plugin was designed to work as a persistent collaborative tool focused on practicality inside the game. _(You can say this is Discord+Excell baby too)_

---

## Main features

### 1. Discord sign-in

EZSheets uses Discord OAuth for login.

The plugin flow is:

1. you click **Sign In with Discord**
2. the browser opens the <ins>official</isn> Discord authorization page
3. after authorization, the browser redirects to a local callback on `127.0.0.1`
4. the plugin finishes the session and loads your spreadsheets

That means:

- the plugin **does not ask for your Discord password or e-mail directly** 
- authentication goes through Discord's official OAuth flow
- the plugin uses a local callback on your computer instead of asking you to log in manually inside the plugin window

In the current code, the client only requests the **`identify`** scope during OAuth. That reduces the amount of Discord data requested during login.

---

### 2. Cloud-synced spreadsheets

Each spreadsheet can be saved to and loaded from the cloud.

This allows you to:

- keep using the same spreadsheet across future sessions
- share access with other people
- preserve sheet state across computers and sessions
- collaborate on the same sheet as a group

The plugin also keeps local cache and session information to make repeated use more practical between logins.

---

### 3. Code-based sharing

EZSheets supports access through:

- **shared code**
- **temporary / unique code**, when applicable

This makes it easy to invite other people without requiring each user to have the permanent code of the sheet.

---

### 4. Roles, presets, and permissions

The sheet has more detailed access control than just viewer/editor.

The system includes:

- default join role (**Default Join Role**)
- permission presets
- per-user permission editing
- user blocking
- access removal
- visual summary of permissions and roles

Among the permissions present in the project are actions such as:

- edit sheet _(User can edit the sheet)_
- delete sheet _(User can delete the sheet)_
- edit permissions _(User can edit role permissions sheet, not applicable to its own role)_
- create tabs _(User can create new tabs in the sheet)_
- view history _(User can see the sheet activity history)_
- use comments _(User can leave popup comments in the sheet cells)_
- import sheet _(User can import a existing `.csv`/`.xls` sheet as a new tab)_
- invite / manage access _(User can view and generate acess codes for the sheet)_
- block users _(User can remove other users from the sheet)_
- administrative privileges _(User have almost full access on the sheet)_

This model lets the owner keep fine-grained control over what each person can or cannot do.

---

### 5. Sheet chat

Each spreadsheet has its own Live chat.

The chat includes:

- per-sheet history
- unread notification
- mentions notification
- popup / pop-out mode
- local clock in the interface
- integration with sheet members

---

### 6. Cell comments

EZSheets supports per-cell comments.

This is useful for:

- quick notes
- internal instructions
- temporary annotations
- extra context without changing the main cell value

---

### 7. Snapshots, history, and restoration

The snapshot system allows you to create restore points.

This helps in cases such as:

- editing mistakes
- major structural changes
- needing to return to a previous state
- basic audit/history tracking of sheet changes

The project also includes shared history and undo/redo structures at the document level.

---

### 8. Presets and templates

The plugin includes support for:

- permission presets
- tab presets
- spreadsheet templates

This speeds up creation of new sheets and reduces repeated manual setup.

---

### 9. Import and export

The project includes support for:

- local export to **Excel (`.xlsx`)**
- import from **Excel** and **CSV**, when permitted
- cloud import
- cloud save
- auto cloud save

This is important for:

- manual backup
- data migration
- hybrid use inside and outside the game

---

### 10. Presence and member context

The interface also includes features such as:

- **Users seen in this sheet**
- member list
- role colors
- visual access status
- helper windows such as **Find**, **Permissions**, **History**, **Blocklist**, and others

---

## Commands

The current project registers these commands:

- `/EZSheets`
- `/ss`
- `/sheet`

All of them open the main plugin window.

---

## Why a user can trust EZSheets

No third-party plugin should ask for blind trust. The best way to trust a plugin is to understand **what it does, what it stores, what it sends and what it does not do**.

### What helps build trust in this project

- the project is open-source
- login uses official Discord OAuth and does not collect anything besides your username
- the client currently uses only the `identify` scope in the discord login flow
- real backend secrets do not need to be stored in the user's plugin
- the plugin uses a public Supabase URL and publishable key, which is the normal for client applications
- the backend keeps sensitive secrets inside Supabase Edge Functions / secrets
- the server session flow uses its own separate token with expiration
- the server session token is persisted in the database as a **hash**, not as a raw lookup token

---

## Privacy and data

### What the plugin stores locally

The plugin may store local information such as:

- window layout and position
- sidebar state
- last opened sheet
- locally cached sheets
- last local export path
- favorite states
- chat read anchors
- mention toast anchors
- per-character session state
- session tokens and refresh token for login restoration
- user's display name
- user identifier

This data stays in the local plugin / Dalamud environment on the user's computer.

### What is stored in the backend

On the backend, the project is structured to handle data such as:

- spreadsheets
- sheet members
- roles and permissions
- snapshots / sheet history
- presence / runtime state
- chat messages
- EZSheets server sessions
- invite / access codes when applicable

### About the local callback

During login, the plugin opens a callback at:

`http://127.0.0.1:38473/callback/`

That means the browser returns the result to your own computer. The address `127.0.0.1` is local loopback.

---

## Backend security

The project includes a Supabase structure with:

- Edge Functions
- publishable key on the client
- service role on the backend
- server session tables
- authentication checks
- SQL policies and functions for per-user access

The `setup.sql` file shows a foundation with access control and functions for joining sheets by code and handling specific permissions.

In other words: security **does not depend only on the interface**. An important part of the logic is designed to live in the backend.

Even so, the recommendation remains:

- do not share the permanent code of your sheet
- do not give Admin or Relevant permissions to people you dont know well
- do not rely only on visual UI restrictions for access rules
- to join a sheet, only a code is necessary, do not trust any URL someone send to you

---

## Chat history and visibility

In the current state of the project, the backend logic was prepared to:

- limit history visibility by daily period in **EST** timezone _(Clear every 24 hours)_
- prevent a new user from seeing messages from before they joined the sheet

---

## What the plugin accesses on Discord

The plugin requests only:

- basic user identity through Discord OAuth

This is normally used to:

- identify the session owner
- show the display name
- associate membership and permissions in the backend

The plugin **will never** ask for your password, any information, a manual Discord token, or a text-field login inside the plugin UI.

---

## What the plugin accesses on the user's computer

To work as a Dalamud plugin, EZSheets needs access to:

- the plugin UI inside the game
- the default browser for OAuth login
- local plugin configuration storage
- local export for spreadsheets / files
- a temporary local callback to complete login

It does not need an extra driver, a dedicated background service, or installation outside the normal plugin flow.

---

## Installation for end users

### Via Dalamud custom repo

1. in Dalamud, open the experimental / custom repos area
2. add the URL: `https://raw.githubusercontent.com/seventity7/EZSheets/main/repo.json`
3. Save it
4. In the plugin list, search for `EZSheets`
5. install the plugin normally
6. open EZSheets with `/EZSheets`, `/ss`, or `/sheet`
7. click **Sign In with Discord**
8. authorize in the browser
9. return to the game
10. create or open a sheet

---

## Recommended usage for users

For safer use:

- do not store extremely sensitive information in the spreadsheet
- treat the sheet as an operational tool, not as a vault for private information
- use snapshots before major changes
- use local export as an additional backup
- review permissions before sharing codes
- remove access for members who no longer need to participate

---

## Limitations and notes

- EZSheets is a third-party tool, not an official Square Enix product
- it is also not an official Discord or Supabase product
- like any cloud-connected plugin, the experience depends on both the client and the published backend
- if the backend is offline, some online features may fail
- if you publish frequent updates, it is worth keeping a clear changelog and backup