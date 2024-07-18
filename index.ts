import {
  ExecutionContext,
  ForwardableEmailMessage,
  HTMLRewriterElementContentHandlers,
  KVNamespace,
  R2Bucket,
  ReadableStream,
  Text
} from "@cloudflare/workers-types";

import type {
  FixedLengthStream, HTMLRewriter, Response} from "@cloudflare/workers-types"

import {API} from "ynab";

async function parseChaseEmail(emailBody: ReadableStream<Uint8Array>):
    Promise<Entry|undefined> {
  class ChaseEmailParser implements HTMLRewriterElementContentHandlers {

    amount: number|undefined = undefined;
    account: string|undefined = undefined;
    date: Date|undefined = undefined;
    description: string|undefined = undefined;

    handleNewText(header: string, content: string) {
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

    last_message: string = "";
    working_text = "";
    text(text: Text) {
      this.working_text = this.working_text.concat(text.text).trim();
      if (text.lastInTextNode) {
        // This is to exclude any open tags passed in
        // The debugging email server somehow inserts =\r\n in random places
        this.working_text =
            this.working_text.replace(/<=[\s\S]*?>/g, "").replace("=\r\n", "");

        this.handleNewText(this.last_message, this.working_text);

        this.last_message = this.working_text;
        this.working_text = "";
      }
    }

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
  const importId = `ys:b0:${entry.amount}:${new Date().getUTCSeconds()}`;

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
                           env: WorkerEnv): Promise<UserInfo|null> {
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
                              env: WorkerEnv): Promise<UserAccount|null> {
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

async function handleYNABSync(email: ForwardableEmailMessage,
                              body: ReadableStream<Uint8Array>,
                              env: WorkerEnv) {
  const emailMeta = parseEmailAddress(email.to)
  const userTag = emailMeta.tag;
  const domain = emailMeta.domain;

  const authInfo = await getUserInfo(userTag, env);
  if (!authInfo) {
    console.warn(`User auth with tag: ${userTag} not found`);
    email.setReject("Non-acceptable Address");
    return;
  }

  let content: Entry|undefined = undefined;
  // if (domain.endsWith("chase.com")) {
  content = await parseChaseEmail(body);
  // }

  if (!content) {
    console.log(
        `Cannot process email from ${email.from} to ${email.to}, subject:`,
        email.headers.get("Subject"));
    await email.forward(authInfo.email);
    return;
  }

  console.debug("Parsed object:", content);

  const ynabInfo = await getUserAccount(userTag, content.account, env);
  if (!ynabInfo) {
    console.warn(
        `User with tag: ${userTag} and account ${content.account} not found`);
    email.setReject("Non-acceptable Address");
    return;
  }

  await createYNABTransaction(authInfo.accessToken, content, ynabInfo);

  if (ynabInfo.forwardAddress) {
    console.debug("Forward to email address as configured");
    await email.forward(authInfo.email);
  }
}

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

async function uploadEmailToR2(email: ForwardableEmailMessage,
                               emailBody: ReadableStream<Uint8Array>,
                               bucket: R2Bucket) {
  const addr = parseEmailAddress(email.to);
  const key = `${addr.username}/${addr.tag}/${email.from}/${
      email.headers.get("Message-ID")}`;
  console.debug("Uploading email as", key);
  const stream = new FixedLengthStream(email.rawSize);
  emailBody.pipeTo(stream.writable);
  await bucket.put(key, stream.readable);
  console.debug("Email uploaded successfully");
}

interface WorkerEnv {
  BOUNCE_ADDRESS: string;
  YS_CONFIG: KVNamespace;
  YS_EMAIL_STORAGE: R2Bucket;
}

export default {
  async email(email: ForwardableEmailMessage, env: WorkerEnv,
              _ctx: ExecutionContext) {
    console.debug(`Handle email: from ${email.from} to ${email.to}`);

    const addr = parseEmailAddress(email.to);
    console.debug("Parsed email address as:", addr);

    if (addr.username === "ingest") {
      const [copyStream, ynabStream] = email.raw.tee();
      await Promise.all([
        uploadEmailToR2(email, copyStream, env.YS_EMAIL_STORAGE),
        handleYNABSync(email, ynabStream, env)
      ]);
    } else if (addr.username === "bounce") {
      await email.forward(env.BOUNCE_ADDRESS);
    } else {
      await uploadEmailToR2(email, email.raw, env.YS_EMAIL_STORAGE);
    }
  }
}
