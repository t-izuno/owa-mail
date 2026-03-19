#!/usr/bin/env node
/**
 * owa-mail: Office365 OWA メールクライアント
 *
 * Playwright でブラウザを起動してログイン → セッションを保存。
 * 以降は保存済みセッションのクッキーで OWA service.svc を叩く。
 *
 * Install: npm install && npm install -g .
 *          npm install -g @playwright/cli@latest
 *
 * Usage:
 *   owa-mail login
 *   owa-mail folders
 *   owa-mail list [--folder inbox] [--count 20] [--unread] [--from addr] [--subject text]
 *   owa-mail read <item_id>
 *   owa-mail search <query> [--folder inbox]
 *   owa-mail attachment <item_id> [--save-dir ./attachments]
 *   owa-mail mark-read <item_id> [<item_id> ...]
 */

import { basename, join } from "node:path";
import { execFileSync } from "node:child_process";
import { mkdirSync, readFileSync, writeFileSync } from "node:fs";
import { homedir } from "node:os";
import { argv, exit, stdout } from "node:process";

// ── OWA レスポンス型定義 ─────────────────────────────────────────────────────

interface OwaFolderId {
  Id: string;
  ChangeKey?: string;
}

interface OwaFolder {
  __type: string;
  FolderId: OwaFolderId;
  DisplayName: string;
  TotalCount: number;
  UnreadCount: number;
}

interface OwaMailbox {
  Name: string;
  EmailAddress: string;
}

interface OwaBody {
  Value: string;
  BodyType: string;
}

interface OwaAttachmentId {
  Id: string;
}

interface OwaAttachment {
  __type: string;
  AttachmentId: OwaAttachmentId;
  Name: string;
  Content?: string;
}

interface OwaMessage {
  ItemId: OwaFolderId;
  Subject: string;
  From: { Mailbox: OwaMailbox };
  ToRecipients?: Array<{ Mailbox: OwaMailbox }>;
  DateTimeReceived: string;
  IsRead: boolean;
  HasAttachments: boolean;
  Body?: OwaBody;
  Attachments?: OwaAttachment[];
}

interface OwaResponseItem {
  RootFolder?: {
    Folders?: OwaFolder[];
    Items?: OwaMessage[];
  };
  Items?: OwaMessage[];
  Attachments?: OwaAttachment[];
}

interface OwaResponse {
  Body?: {
    ResponseMessages?: {
      Items?: OwaResponseItem[];
    };
  };
}

// ── 出力型定義 ──────────────────────────────────────────────────────────────

interface FolderInfo {
  id: string | undefined;
  name: string;
  total: number;
  unread: number;
}

interface MailSummary {
  id: string | undefined;
  subject: string;
  from: string | undefined;
  from_email: string | undefined;
  date: string;
  read: boolean;
  has_attachment: boolean;
}

interface MailDetail {
  id: string;
  subject: string;
  from: string | undefined;
  from_email: string | undefined;
  to: string[];
  date: string;
  body: string;
  attachments: string[];
}

interface AttachmentResult {
  filename: string;
  path: string;
  size: number;
}

// ── セッション型定義 ────────────────────────────────────────────────────────

interface SessionCookie {
  name: string;
  value: string;
  domain: string;
}

interface Session {
  cookies: SessionCookie[];
}

// ── 定数 ─────────────────────────────────────────────────────────────────────

const OWA_BASE = "https://outlook.office.com/owa/";
const SERVICE_URL = "https://outlook.office.com/owa/service.svc";
const SESSION_PATH = join(homedir(), ".config", "owa-mail", "session.json");

// ── セッション管理 ────────────────────────────────────────────────────────────

function findPlaywrightCli(): string {
  try {
    execFileSync("playwright-cli", ["--help"], { stdio: "ignore" });
    return "playwright-cli";
  } catch {
    // npx 経由で試す
    try {
      execFileSync("npx", ["playwright-cli", "--help"], { stdio: "ignore" });
      return "npx";
    } catch {
      console.error(
        JSON.stringify({
          error: "playwright-cli が見つかりません",
          hint: "`npm install -g @playwright/cli@latest` を実行してください",
        }, null, 2)
      );
      exit(1);
    }
  }
}

function runPlaywrightCli(bin: string, args: string[]): string {
  const cmd = bin === "npx" ? ["playwright-cli", ...args] : args;
  return execFileSync(bin, cmd, {
    encoding: "utf8",
    stdio: ["inherit", "pipe", "inherit"],
    timeout: 300_000,
  });
}

function login(): void {
  const bin = findPlaywrightCli();

  console.error("ブラウザを起動します。Office365 にログインしてください。");

  // ブラウザを開いて OWA にアクセス
  runPlaywrightCli(bin, ["open", OWA_BASE, "--headed"]);

  console.error("ログイン後、セッションを保存します...");

  // セッション保存
  mkdirSync(join(homedir(), ".config", "owa-mail"), { recursive: true });
  runPlaywrightCli(bin, ["state-save", SESSION_PATH]);

  console.error(`セッションを保存しました: ${SESSION_PATH}`);
}

// ── OWA クライアント ──────────────────────────────────────────────────────────

function loadSession(): Session {
  try {
    return JSON.parse(readFileSync(SESSION_PATH, "utf8")) as Session;
  } catch {
    console.error(
      JSON.stringify({
        error: "セッションが見つかりません",
        hint: "先に `owa-mail login` を実行してください",
      }, null, 2)
    );
    exit(1);
  }
}

function buildHeaders(session: Session, action: string): HeadersInit {
  const cookies = session.cookies
    .filter((c) => c.domain.includes("outlook.office.com"))
    .map((c) => `${c.name}=${c.value}`)
    .join("; ");

  const canary = session.cookies.find(
    (c) => c.name === "X-OWA-CANARY" && c.domain.includes("outlook.office.com")
  )?.value;

  if (!canary) {
    console.error(
      JSON.stringify({
        error: "X-OWA-CANARY が見つかりません。セッションが期限切れの可能性があります",
        hint: "`owa-mail login` で再ログインしてください",
      }, null, 2)
    );
    exit(1);
  }

  return {
    Action: action,
    "Content-Type": "application/json; charset=utf-8",
    "X-OWA-CANARY": canary,
    "X-Requested-With": "XMLHttpRequest",
    Cookie: cookies,
  };
}

async function callService(session: Session, action: string, body: unknown): Promise<OwaResponse> {
  const headers = buildHeaders(session, action);
  const resp = await fetch(SERVICE_URL, {
    method: "POST",
    headers,
    body: JSON.stringify(body),
  });

  if (resp.status === 440 || resp.status === 401) {
    console.error(
      JSON.stringify({
        error: "セッション期限切れ",
        hint: "`owa-mail login` で再ログインしてください",
      }, null, 2)
    );
    exit(1);
  }

  if (!resp.ok) throw new Error(`HTTP ${resp.status}: ${resp.statusText}`);
  return resp.json() as Promise<OwaResponse>;
}

// レスポンスから Items 配列の先頭要素を取り出すヘルパー
function extractFirstItem(result: OwaResponse): OwaResponseItem {
  return result.Body?.ResponseMessages?.Items?.[0] ?? {};
}

// ── Exchange リクエスト共通ヘッダー ───────────────────────────────────────────

const EXCHANGE_HEADER = {
  __type: "JsonRequestHeaders:#Exchange",
  RequestServerVersion: "Exchange2013",
};

// ── フォルダ解決 ──────────────────────────────────────────────────────────────

const DISTINGUISHED: Record<string, string> = {
  inbox: "inbox",
  受信トレイ: "inbox",
  sent: "sentitems",
  送信済み: "sentitems",
  drafts: "drafts",
  下書き: "drafts",
  deleted: "deleteditems",
  ゴミ箱: "deleteditems",
  junk: "junkemail",
  迷惑メール: "junkemail",
};

async function resolveFolderIdObj(session: Session, name: string): Promise<Record<string, string>> {
  const key = name.toLowerCase();
  const distinguished = DISTINGUISHED[key];
  if (distinguished) {
    return { __type: "DistinguishedFolderId:#Exchange", Id: distinguished };
  }
  // 名前検索
  const list = await cmdFolders(session);
  const found = list.find((f) => f.name?.toLowerCase().includes(key));
  if (!found?.id) throw new Error(`フォルダが見つかりません: ${name}`);
  return { __type: "FolderId:#Exchange", Id: found.id };
}

// ── コマンド実装 ──────────────────────────────────────────────────────────────

async function cmdFolders(session: Session): Promise<FolderInfo[]> {
  const result = await callService(session, "FindFolder", {
    __type: "FindFolderJsonRequest:#Exchange",
    Header: EXCHANGE_HEADER,
    Body: {
      __type: "FindFolderRequest:#Exchange",
      FolderShape: { BaseShape: "Default" },
      Paging: null,
      ParentFolderIds: [{ __type: "DistinguishedFolderId:#Exchange", Id: "msgfolderroot" }],
      ReturnParentFolder: false,
      Traversal: "Deep",
    },
  });

  const folders = extractFirstItem(result).RootFolder?.Folders ?? [];

  return folders
    .filter((f) => f.__type === "Folder:#Exchange")
    .map((f) => ({
      id: f.FolderId?.Id,
      name: f.DisplayName,
      total: f.TotalCount,
      unread: f.UnreadCount,
    }));
}

function toMailSummary(m: OwaMessage): MailSummary {
  return {
    id: m.ItemId?.Id,
    subject: m.Subject,
    from: m.From?.Mailbox?.Name,
    from_email: m.From?.Mailbox?.EmailAddress,
    date: m.DateTimeReceived?.slice(0, 16).replace("T", " "),
    read: m.IsRead,
    has_attachment: m.HasAttachments,
  };
}

type MatchMode = "substring" | "prefix" | "exact";

interface ListFilters {
  unreadOnly: boolean;
  from?: string;
  fromMatch: "exact" | "contains";
  subject?: string;
  subjectMatch: MatchMode;
}

function buildRestriction(filters: ListFilters): unknown {
  const conditions: unknown[] = [];

  if (filters.unreadOnly) {
    conditions.push({
      __type: "IsEqualTo:#Exchange",
      FieldURI: { __type: "PropertyUri:#Exchange", FieldURI: "message:IsRead" },
      FieldURIOrConstant: { __type: "Constant:#Exchange", Value: "false" },
    });
  }

  if (filters.from) {
    if (filters.fromMatch === "contains") {
      conditions.push({
        __type: "Contains:#Exchange",
        ContainmentMode: "Substring",
        ContainmentComparison: "IgnoreCase",
        FieldURI: { __type: "PropertyUri:#Exchange", FieldURI: "message:From" },
        Constant: { __type: "Constant:#Exchange", Value: filters.from },
      });
    } else {
      conditions.push({
        __type: "IsEqualTo:#Exchange",
        FieldURI: { __type: "PropertyUri:#Exchange", FieldURI: "message:From" },
        FieldURIOrConstant: { __type: "Constant:#Exchange", Value: filters.from },
      });
    }
  }

  if (filters.subject) {
    if (filters.subjectMatch === "exact") {
      conditions.push({
        __type: "IsEqualTo:#Exchange",
        FieldURI: { __type: "PropertyUri:#Exchange", FieldURI: "item:Subject" },
        FieldURIOrConstant: { __type: "Constant:#Exchange", Value: filters.subject },
      });
    } else {
      conditions.push({
        __type: "Contains:#Exchange",
        ContainmentMode: filters.subjectMatch === "prefix" ? "Prefixed" : "Substring",
        ContainmentComparison: "IgnoreCase",
        FieldURI: { __type: "PropertyUri:#Exchange", FieldURI: "item:Subject" },
        Constant: { __type: "Constant:#Exchange", Value: filters.subject },
      });
    }
  }

  if (conditions.length === 0) return null;
  if (conditions.length === 1) return conditions[0];
  return {
    __type: "And:#Exchange",
    Items: conditions,
  };
}

async function cmdList(session: Session, folderName: string, count: number, filters: ListFilters): Promise<MailSummary[]> {
  const folderObj = await resolveFolderIdObj(session, folderName);

  const result = await callService(session, "FindItem", {
    __type: "FindItemJsonRequest:#Exchange",
    Header: EXCHANGE_HEADER,
    Body: {
      __type: "FindItemRequest:#Exchange",
      ItemShape: {
        BaseShape: "IdOnly",
        AdditionalProperties: [
          { __type: "PropertyUri:#Exchange", FieldURI: "item:Subject" },
          { __type: "PropertyUri:#Exchange", FieldURI: "message:From" },
          { __type: "PropertyUri:#Exchange", FieldURI: "item:DateTimeReceived" },
          { __type: "PropertyUri:#Exchange", FieldURI: "message:IsRead" },
          { __type: "PropertyUri:#Exchange", FieldURI: "item:HasAttachments" },
        ],
      },
      Paging: {
        __type: "IndexedPageView:#Exchange",
        BasePoint: "Beginning",
        Offset: 0,
        MaxEntriesReturned: count,
      },
      Restriction: buildRestriction(filters),
      SortOrder: [{
        Order: "Descending",
        FieldURI: { __type: "PropertyUri:#Exchange", FieldURI: "item:DateTimeReceived" },
      }],
      ParentFolderIds: [folderObj],
      Traversal: "Shallow",
    },
  });

  const messages = extractFirstItem(result).RootFolder?.Items ?? [];
  return messages.map(toMailSummary);
}

async function cmdRead(session: Session, itemId: string): Promise<MailDetail | { error: string }> {
  const result = await callService(session, "GetItem", {
    __type: "GetItemJsonRequest:#Exchange",
    Header: EXCHANGE_HEADER,
    Body: {
      __type: "GetItemRequest:#Exchange",
      ItemShape: {
        BaseShape: "Default",
        BodyType: "Text",
        AdditionalProperties: [
          { __type: "PropertyUri:#Exchange", FieldURI: "item:Attachments" },
        ],
      },
      ItemIds: [{ __type: "ItemId:#Exchange", Id: itemId }],
    },
  });

  const items = extractFirstItem(result).Items ?? [];
  const m = items[0];
  if (!m) return { error: `メッセージが見つかりません: ${itemId}` };

  return {
    id: itemId,
    subject: m.Subject,
    from: m.From?.Mailbox?.Name,
    from_email: m.From?.Mailbox?.EmailAddress,
    to: (m.ToRecipients ?? []).map((r) => r.Mailbox?.EmailAddress),
    date: m.DateTimeReceived?.slice(0, 16).replace("T", " "),
    body: m.Body?.Value ?? "(本文なし)",
    attachments: (m.Attachments ?? [])
      .filter((a) => a.__type?.startsWith("FileAttachment"))
      .map((a) => a.Name),
  };
}

async function cmdSearch(session: Session, query: string, folderName: string): Promise<MailSummary[]> {
  const folderObj = await resolveFolderIdObj(session, folderName);

  const result = await callService(session, "FindItem", {
    __type: "FindItemJsonRequest:#Exchange",
    Header: EXCHANGE_HEADER,
    Body: {
      __type: "FindItemRequest:#Exchange",
      ItemShape: {
        BaseShape: "IdOnly",
        AdditionalProperties: [
          { __type: "PropertyUri:#Exchange", FieldURI: "item:Subject" },
          { __type: "PropertyUri:#Exchange", FieldURI: "message:From" },
          { __type: "PropertyUri:#Exchange", FieldURI: "item:DateTimeReceived" },
          { __type: "PropertyUri:#Exchange", FieldURI: "message:IsRead" },
          { __type: "PropertyUri:#Exchange", FieldURI: "item:HasAttachments" },
        ],
      },
      Paging: {
        __type: "IndexedPageView:#Exchange",
        BasePoint: "Beginning",
        Offset: 0,
        MaxEntriesReturned: 50,
      },
      QueryString: query,
      ParentFolderIds: [folderObj],
      Traversal: "Shallow",
    },
  });

  const messages = extractFirstItem(result).RootFolder?.Items ?? [];
  return messages.map(toMailSummary);
}

async function cmdAttachment(session: Session, itemId: string, saveDir: string): Promise<AttachmentResult[] | [{ error: string }]> {
  const result = await callService(session, "GetItem", {
    __type: "GetItemJsonRequest:#Exchange",
    Header: EXCHANGE_HEADER,
    Body: {
      __type: "GetItemRequest:#Exchange",
      ItemShape: {
        BaseShape: "IdOnly",
        AdditionalProperties: [
          { __type: "PropertyUri:#Exchange", FieldURI: "item:Attachments" },
        ],
      },
      ItemIds: [{ __type: "ItemId:#Exchange", Id: itemId }],
    },
  });

  const items = extractFirstItem(result).Items ?? [];
  const firstItem = items[0];
  if (!firstItem) return [{ error: "メッセージが見つかりません" }];

  const attachments = (firstItem.Attachments ?? [])
    .filter((a) => a.__type?.startsWith("FileAttachment"));

  mkdirSync(saveDir, { recursive: true });
  const saved: AttachmentResult[] = [];

  for (const att of attachments) {
    const attId = att.AttachmentId?.Id;
    const safeName = basename(att.Name ?? "attachment");

    const attResult = await callService(session, "GetAttachment", {
      __type: "GetAttachmentJsonRequest:#Exchange",
      Header: EXCHANGE_HEADER,
      Body: {
        __type: "GetAttachmentRequest:#Exchange",
        AttachmentShape: { IncludeMimeContent: true },
        AttachmentIds: [{ __type: "RequestAttachmentId:#Exchange", Id: attId }],
      },
    });

    const attItems = extractFirstItem(attResult).Attachments ?? [];
    const attData = attItems[0];
    if (attData?.Content) {
      const content = Buffer.from(attData.Content, "base64");
      const outPath = join(saveDir, safeName);
      writeFileSync(outPath, content);
      saved.push({ filename: safeName, path: outPath, size: content.length });
    }
  }

  return saved;
}

async function cmdMarkRead(session: Session, itemIds: string[]): Promise<{ marked: string[] }> {
  const result = await callService(session, "UpdateItem", {
    __type: "UpdateItemJsonRequest:#Exchange",
    Header: EXCHANGE_HEADER,
    Body: {
      __type: "UpdateItemRequest:#Exchange",
      ConflictResolution: "AlwaysOverwrite",
      MessageDisposition: "SaveOnly",
      ItemChanges: itemIds.map((id) => ({
        __type: "ItemChange:#Exchange",
        ItemId: { __type: "ItemId:#Exchange", Id: id },
        Updates: [{
          __type: "SetItemField:#Exchange",
          FieldURI: { __type: "PropertyUri:#Exchange", FieldURI: "message:IsRead" },
          Message: {
            __type: "Message:#Exchange",
            IsRead: true,
          },
        }],
      })),
    },
  });

  // レスポンス確認（エラーがあれば例外で落ちる）
  const _ = extractFirstItem(result);
  return { marked: itemIds };
}

// ── CLI ───────────────────────────────────────────────────────────────────────

const BOOLEAN_FLAGS = new Set(["unread"]);

export function parseArgs(args: string[]) {
  const flags: Record<string, string | boolean> = {};
  const positional: string[] = [];

  for (let i = 0; i < args.length; i++) {
    const arg = args[i];
    if (arg === undefined) continue;

    if (arg.startsWith("--")) {
      const key = arg.slice(2);
      if (BOOLEAN_FLAGS.has(key)) {
        flags[key] = true;
      } else {
        const next = args[i + 1];
        if (next !== undefined && !next.startsWith("--")) {
          flags[key] = next;
          i++;
        } else {
          flags[key] = true;
        }
      }
    } else {
      positional.push(arg);
    }
  }
  return { cmd: positional[0], rest: positional.slice(1), flags };
}

function print(data: unknown) {
  stdout.write(JSON.stringify(data, null, 2) + "\n");
}

async function main() {
  const { cmd, rest, flags } = parseArgs(argv.slice(2));

  if (cmd === "login") {
    login();
    return;
  }

  if (!cmd || cmd === "help") {
    console.log(`
owa-mail <command> [options]

Commands:
  login                           ブラウザでログイン・セッション保存
  folders                         フォルダ一覧
  list [--folder inbox]           メール一覧
       [--count 20]
       [--unread]
       [--from addr]              差出人（デフォルト完全一致）
       [--from-match mode]        exact(default) / contains
       [--subject text]           件名フィルタ
       [--subject-match mode]     substring(default) / prefix / exact
  read <item_id>                  メール本文
  search <query>                  メール検索
         [--folder inbox]
  attachment <item_id>            添付ファイル取得
             [--save-dir ./attachments]
  mark-read <item_id> [...]       既読にする

Session: ${SESSION_PATH}
    `);
    return;
  }

  const session = loadSession();
  let result: unknown;

  switch (cmd) {
    case "folders":
      result = await cmdFolders(session);
      break;

    case "list": {
      const fromMatch = (flags["from-match"] as string | undefined) ?? "exact";
      if (fromMatch !== "exact" && fromMatch !== "contains") {
        console.error("--from-match は exact / contains のいずれか");
        exit(1);
      }
      const subjectMatch = (flags["subject-match"] as string | undefined) ?? "substring";
      if (subjectMatch !== "substring" && subjectMatch !== "prefix" && subjectMatch !== "exact") {
        console.error("--subject-match は substring / prefix / exact のいずれか");
        exit(1);
      }
      result = await cmdList(
        session,
        (flags.folder as string | undefined) ?? "inbox",
        Number(flags.count ?? 20),
        {
          unreadOnly: flags.unread === true,
          from: flags.from as string | undefined,
          fromMatch,
          subject: flags.subject as string | undefined,
          subjectMatch,
        }
      );
      break;
    }

    case "read":
      if (!rest[0]) { console.error("item_id が必要"); exit(1); }
      result = await cmdRead(session, rest[0]);
      break;

    case "search":
      if (!rest[0]) { console.error("query が必要"); exit(1); }
      result = await cmdSearch(session, rest[0], (flags.folder as string | undefined) ?? "inbox");
      break;

    case "attachment":
      if (!rest[0]) { console.error("item_id が必要"); exit(1); }
      result = await cmdAttachment(session, rest[0], (flags["save-dir"] as string | undefined) ?? "./attachments");
      break;

    case "mark-read":
      if (!rest[0]) { console.error("item_id が必要"); exit(1); }
      result = await cmdMarkRead(session, rest);
      break;

    default:
      console.error(`Unknown command: ${cmd}`);
      exit(1);
  }

  print(result);
}

// Node 実行・Bun 実行・Bun compile の単一実行ファイルを同じ判定で扱う。
const isDirectRun = (() => {
  const meta = import.meta as ImportMeta & { main?: boolean };
  return meta.main ?? argv[1]?.endsWith("index.js") ?? argv[1]?.endsWith("index.ts");
})();

if (isDirectRun) {
  main().catch((err: unknown) => {
    console.error(JSON.stringify({ error: String(err) }, null, 2));
    exit(1);
  });
}
