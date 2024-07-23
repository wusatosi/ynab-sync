import {
  ExecutionContext,
  ForwardableEmailMessage,
  Headers,
  HTMLRewriterElementContentHandlers,
  KVNamespace,
  R2Bucket,
  ReadableStream,
  Text
} from "@cloudflare/workers-types";

import type {
  FixedLengthStream, HTMLRewriter, Response, Request} from
  "@cloudflare/workers-types"

import {API} from "ynab";

import PostalMime  from "postal-mime";
import type {RawEmail, Email} from "postal-mime";

abstract class EmailParser {
  amount: number|undefined = undefined;
  date: Date|undefined = undefined;
  description: string|undefined = undefined;
  account: string|undefined = undefined;

  result(): Entry|undefined {
    const result = {
      amount : this.amount,
      title : this.description,
      account : this.account,
      postedDate : this.date
    };

    if (this.amount && this.account && this.date && this.description) {
      return result as Entry;
    } else {
      console.log("Email parsing failed", result);
      return undefined;
    }
  }
}

abstract class GenericStringParser extends EmailParser implements
    HTMLRewriterElementContentHandlers {

  abstract completeTextChunk(chunk: string): void;

  private working_text = "";
  text(text: Text) {
    this.working_text = this.working_text.concat(text.text);
    if (text.lastInTextNode) {
      // This is to exclude any open tags passed in
      // The debugging email server somehow inserts =\r\n in random places
      this.working_text = this.working_text.replaceAll("=\r\n", "")
                              .replaceAll(/<=[\s\S]*?>/g, "")
                              .trim();
      this.completeTextChunk(this.working_text);
      this.working_text = "";
    }
  }
}

abstract class GenericStringPairParser extends GenericStringParser {
  abstract completeTextPair(header: string, content: string): void;

  private last_message: string = "";
  completeTextChunk(chunk: string): void {
    this.completeTextPair(this.last_message, chunk);
    this.last_message = chunk;
  }
}

async function parseBOAEmail(email: CompoundEmail): Promise<Entry|undefined> {
  class BOAEmailParser extends GenericStringParser {
    completeTextChunk(chunk: string) {
      console.debug("Accepting new chunk", chunk);
      {
        // Match $5.72 => 5.72
        const match = chunk.match(/\$?(\d+\.\d{2})/);
        if (match && (match.length > 1)) {
          // milliunits format => -572
          // Note: rounding for floating point
          this.amount = Math.round(parseFloat(match[1]) * -1000);
          console.debug("Set amount", this.amount);
          return;
        }
      }
      {
        // Match ending in xxxx
        const match = chunk.match(/ending in (\d{4})/);
        if (match && (match.length > 1)) {
          this.account = match[1];
          console.debug("Set account", this.account);
          return;
        }
      }
      {
        // Match July 19, 2024
        const dateNum = Date.parse(chunk);
        if (dateNum && dateNum != Number.NaN) {
          this.date = new Date(dateNum);
          console.debug("Set date", this.date);
          return;
        }
      }
      {
        if (!this.description) {
          this.description = chunk;
          console.debug("Set description", this.description);
          return;
        }
      }
    }
  }

  const emailBody = email.html();
  if (emailBody === undefined)
    return undefined;

  const parser = new BOAEmailParser();
  await new HTMLRewriter()
      .on("b", parser)
      .transform(new Response(emailBody))
      .arrayBuffer();

  return parser.result();
}

async function parseChaseEmail(email: CompoundEmail): Promise<Entry|undefined> {
  class ChaseEmailParser extends GenericStringPairParser {
    completeTextPair(header: string, content: string) {
      if ("" in [header, content])
        return;

      switch (header) {
      case "Account ending in": {
        console.debug("Processing Account: ", content);
        // Match (...1234) => 1234
        const match = content.match(/\(\.\.\.(\d{4})\)/);
        if (match && (match.length > 1))
          this.account = match[1];
        break;
      }
      case "Made on": {
        console.debug("Processing date: ", content);
        // Match: e.g. Jul 15, 2024 at 7:02 PM ET
        const regex =
            /\b([A-Za-z]{3}) (\d{1,2}), (\d{4}) at (\d{1,2}):(\d{2}) (AM|PM) ([A-Za-z]{2,4})\b/;
        const match = content.match(regex);
        if (match) {
          const dateTime = match[0];
          // Match: e.g. Jul 15
          const date = dateTime.substring(0, dateTime.indexOf("at"));
          this.date = new Date(Date.parse(date));
        }
        break;
      }
      case "Description": {
        console.debug("Processing description: ", content);
        this.description = content;
        break;
      }
      case "Amount": {
        console.debug("Processing amount: ", content);
        // Match $23.45
        const match = content.match(/\$?(\d+\.\d{2})/);
        if (match && (match.length > 1))
          // milliunits format => -2345
          // Note: rounding for floating point
          this.amount = Math.round(parseFloat(match[1]) * -1000);
        break;
      }
      }
    }
  }

  const emailBody = email.html();
  if (emailBody === undefined)
    return undefined;

  const parser = new ChaseEmailParser();
  await new HTMLRewriter()
      .on("td", parser)
      .transform(new Response(emailBody))
      .arrayBuffer();

  return parser.result();
}

interface Entry {
  amount: number;
  title: string;
  account: string;
  postedDate: Date;
}

async function createYNABTransaction(accessToken: string, entry: Entry,
                                     account: UserAccount) {
  const memo = "Auto import through email alert via ynab-sync.";
  const isoString = entry.postedDate.toISOString();
  const dateString = isoString.substring(0, isoString.indexOf("T"));
  const importId = `ys:b0:${dateString}:${new Date().getUTCSeconds()}`;

  const api = new API(accessToken);
  try {
    const res = await api.transactions.createTransaction(account.budgetId, {
      transaction : {
        account_id : account.accountId,
        date : entry.postedDate.toISOString(),
        amount : entry.amount,
        payee_name : entry.title,
        memo : memo,
        cleared : "uncleared",
        approved : false,
        import_id : importId
      }
    });
    console.log("Response from ynab:", res);
  } catch (e) {
    console.warn("YNAB failed with exception:", e);
  }
}

interface UserInfo {
  accessToken: string;
  email: string;
}

async function getUserInfo(userTag: string,
                           env: YNABEnv): Promise<UserInfo|null> {
  const key = `${userTag}/auth`;
  const store = await env.YS_CONFIG.get(key);
  if (!store)
    return null;
  try {
    return JSON.parse(store);
  } catch (e) {
    console.warn("Malformed data in storage, key:", key, " got: ", store, e);
    throw e;
  }
}

interface UserAccount {
  budgetId: string;
  accountId: string;
  forwardAddress: boolean;
}

async function getUserAccount(userTag: string, bankAccount: string,
                              env: YNABEnv): Promise<UserAccount|null> {
  const key = `${userTag}/${bankAccount}`;
  const store = await env.YS_CONFIG.get(key);
  if (!store)
    return null;
  try {
    return JSON.parse(store);
  } catch (e) {
    console.warn("Malformed data in storage, key:", key, " got: ", store, e);
    throw e;
  }
}

function extractDomain(fromAddress: string) {
  const meta = parseEmailAddress(fromAddress);
  // This is for testing
  if (meta.domain === "wusatosi.com") {
    return meta.tag;
  } else {
    return meta.domain;
  }
}

class CompoundEmail {
  private underlyingEmail: ForwardableEmailMessage;
  private bodyStream: ReadableStream;
  private parsedEmail!: Email;

  private constructor(email: ForwardableEmailMessage, body: ReadableStream) {
    this.headers = email.headers;
    this.from = email.from;
    this.to = email.to;

    this.bodyStream = body;
    this.rawSize = email.rawSize;

    this.underlyingEmail = email;
  }

  static async create(email: ForwardableEmailMessage, body: ReadableStream) {
    const e = new CompoundEmail(email, body);
    e.parsedEmail = await PostalMime.parse(e.bodyStream as RawEmail);
    return e;
  }

  readonly rawSize: number;

  readonly headers: Headers;
  readonly from: string;
  readonly to: string;

  setReject(reason: string ) {
    this.underlyingEmail.setReject(reason);
  }
  async forward(rcptTo: string, header?:Headers|undefined) {
    this.underlyingEmail.forward(rcptTo, header);
  }

  html(): string|undefined { return this.parsedEmail.html; }
  messageId(): string { return this.parsedEmail.messageId; }
  subject(): string { return this.parsedEmail.subject || ""; }
};

interface EmailAddress {
  username: string;
  tag: string;
  domain: string;
}

function parseEmailAddress(addr: string): EmailAddress {
  const domainIdx = addr.lastIndexOf("@");
  const domain = addr.substring(domainIdx + 1);

  const front = addr.substring(0, domainIdx);
  const tagIdx = addr.lastIndexOf("+");

  var tag = "";
  var username = front;
  if (tagIdx != -1) {
    tag = front.substring(tagIdx + 1);
    username = front.substring(0, tagIdx);
  }

  return {username : username, tag : tag, domain : domain};
}

async function uploadEmailToR2(email: CompoundEmail,
                               rawEmail: ReadableStream<Uint8Array>,
                               bucket: R2Bucket) {
  const key = `${email.to}/${email.from}/${email.messageId()}`;

  async function uploadRaw() {
    const path = `${key}.eml`;
    console.log("Uploading raw email as", path);
    const stream = new FixedLengthStream(email.rawSize);
    rawEmail.pipeTo(stream.writable);
    await bucket.put(path, stream.readable, {
      httpMetadata : {
        contentType : "message/rfc822",
        contentDisposition : `${email.subject()}.eml`
      }
    });
    console.log("Raw Email uploaded successfully");
  }

  async function uploadHTML() {
    const path = `${key}.html`;
    console.log("Uploading html email as", path);

    const html = email.html();
    if (html === undefined) {
      console.log("Email do not have html part, abort upload.");
      return;
    }

    await bucket.put(path, html, {httpMetadata : {contentType : "text/html"}});
    console.log("HTML Email uploaded successfully");
  }

  return Promise.all([ uploadRaw(), uploadHTML() ]);
}

interface YNABEnv {
  YS_CONFIG: KVNamespace;
  YS_EMAIL_STORAGE: R2Bucket;
  YS_LOGS_STORAGE: R2Bucket;
}

interface WorkerEnv extends YNABEnv {
  BOUNCE_ADDRESS: string;
}

class LoggingHook {
  logs: string = "";
  private hook: {detach: ()=>void}|undefined;

  constructor() {
    const Hook = require('console-hook');
    this.hook = Hook().attach((method: string, args: any[]) => {
      this.logs = `${this.logs}\n[${method}] ${JSON.stringify(args)}`;
    });
  }

  async stopLogging() {
    this.hook?.detach();
    console.log(this.logs);
  }
}

async function handleYNABSync(rawEmail: ForwardableEmailMessage, env: YNABEnv) {
  const lService = new LoggingHook(); 
  try {
    // for logging service
    console.debug("Handling ynab-sync email", {
      from : rawEmail.from,
      to : rawEmail.to,
      subject : rawEmail.headers.get("Subject")
    });

    // Upload Email
    const [copyStream, ynabStream] = rawEmail.raw.tee();
    const email = await CompoundEmail.create(rawEmail, ynabStream);
    const uploadEmailProm = 
      uploadEmailToR2(email, copyStream, env.YS_EMAIL_STORAGE);

    const userTag = parseEmailAddress(rawEmail.to).username;
    const domain = extractDomain(rawEmail.from);
  
    // Grab user info
    const authInfo = await getUserInfo(userTag, env);
    if (!authInfo) {
      console.warn(`User auth with tag: ${userTag} not found`);
      rawEmail.setReject("Non-acceptable Address");
      return;
    }
  
    // Parse email
    let content: Entry|undefined = undefined;
    if (domain.endsWith("chase.com")) {
      content = await parseChaseEmail(email);
    } else if (domain.endsWith("bankofamerica.com")) {
      content = await parseBOAEmail(email);
    } else {
      console.log("Received from a non-whitelisted domain", domain);
      rawEmail.setReject("Non-Acceptable Origin");
      return;
    }
    if (!content) {
      console.log(
          `Cannot process email from ${rawEmail.from} to ${rawEmail.to}, subject:`,
          rawEmail.headers.get("Subject"));
      await rawEmail.forward(authInfo.email);
      return;
    } else {
      console.debug("Parsed object:", content);
    }
  
    // Grab account info
    const ynabInfo = await getUserAccount(userTag, content.account, env);
    if (!ynabInfo) {
      console.warn(
          `User with tag: ${userTag} and account ${content.account} not found`);
      // TODO: maybe I should block these...
      // email.setReject("Non-acceptable Address");
      await rawEmail.forward(authInfo.email);
      return;
    }
  
    // Post to ynab
    await createYNABTransaction(authInfo.accessToken, content, ynabInfo);
  
    // Handle forwarding
    if (ynabInfo.forwardAddress) {
      console.debug("Forward to email address as configured");
      await rawEmail.forward(authInfo.email);
    }

    // Wait till upload finish
    await uploadEmailProm;
  } finally {
    lService.stopLogging();
    await env.YS_LOGS_STORAGE.put(
      (new Date()).toISOString(), 
      lService.logs, 
      { httpMetadata: { contentType: "text/plain" } }
    );
  }
}

export default {
  async email(email: ForwardableEmailMessage, env: WorkerEnv,
              _ctx: ExecutionContext) {
    console.debug("Handling email", {
      from : email.from,
      to : email.to,
      subject : email.headers.get("Subject")
    });

    const addr = parseEmailAddress(email.to);
    console.debug("Parsed email address as:", addr);

    if (addr.domain === "s.kcibald.com") {
      console.log("Handling as YNAB sync");
      await handleYNABSync(email, env);
    } else {
      console.log("Forwarding...");
      await email.forward(env.BOUNCE_ADDRESS);
    }
  },

  async fetch(
      _request: Request,
      _env: WorkerEnv,
      _ctx: ExecutionContext,
      ):
      Promise<Response> { return Response.redirect("https://www.example.com/");}
}
