import {
  ExecutionContext,
  ForwardableEmailMessage,
  HTMLRewriterElementContentHandlers,
  KVNamespace,
  Text
} from "@cloudflare/workers-types";

import type {HTMLRewriter, Response} from "@cloudflare/workers-types";

import {API} from "ynab";

async function parseChaseEmail(email: ForwardableEmailMessage):
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
      if (this.amount && this.account && this.date && this.description) {
        return {
          amount: this.amount, title: this.description, account: this.account,
              postedDate: this.date
        }
      } else {
        console.log(
            `Email parsing failed with: account: ${this.account}, title: ${
                this.description}, amount: ${this.amount}, date: ${this.date}`);
        return undefined;
      }
    }
  }

  const parser = new ChaseEmailParser();
  await new HTMLRewriter()
      .on("td", parser)
      .transform(new Response(email.raw))
      .arrayBuffer();

  return parser.result();
}

interface Entry {
  amount: number;
  title: string;
  account: string;
  postedDate: Date;
}

async function createYNABTransaction(entry: Entry, account: UserAccount,
                                     env: WorkerEnv) {
  const memo = "Auto import through email alert via ynab-sync.";
  const importId =
      `email-sync:b0:${entry.amount}:${new Date().getUTCSeconds()}`;

  const api = new API(env.YNAB_KEY);
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

interface UserAccount {
  budgetId: string;
  accountId: string;
  forwardAddress: string|null;
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

async function handleYNABSync(email: ForwardableEmailMessage, env: WorkerEnv) {
  const userTag = parseEmailAddress(email.to).tag;

  const content = await parseChaseEmail(email);
  if (!content) {
    console.warn(`Cannot process email from ${email.from} to ${email.to}`);
    email.setReject("Cannot parse email for sync");
    return;
  }
  console.debug(`Parsed object:`, content);

  const ynabInfo = await getUserAccount(userTag, content.account, env);
  if (!ynabInfo) {
    console.warn(
        `User with tag: ${userTag} and account ${content.account} not found`);
    email.setReject("Non-acceptable Address");
    return;
  }

  await createYNABTransaction(content, ynabInfo, env);

  if (ynabInfo.forwardAddress) {
    console.debug("Forward to email address as configured",
                  ynabInfo.forwardAddress);
    await email.forward(ynabInfo.forwardAddress);
  }
}

interface WorkerEnv {
  YNAB_KEY: string;
  YS_CONFIG: KVNamespace;
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

interface EmailAddress {
  username: string;
  tag: string;
  domain: string;
}

export default {
  async email(email: ForwardableEmailMessage, env: WorkerEnv,
              _ctx: ExecutionContext) {
    console.debug(`Handle email: from ${email.from} to ${email.to}`);

    const addr = parseEmailAddress(email.to);
    console.debug("Parsed email address as:", addr);
    if (addr.username === "ingest")
      await handleYNABSync(email, env);
  }
}
