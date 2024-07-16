import {
  ExecutionContext,
  ForwardableEmailMessage,
  HTMLRewriter,
  HTMLRewriterElementContentHandlers,
  Response,
  Text
} from "@cloudflare/workers-types"

async function parseChaseEmail(email: ForwardableEmailMessage):
    Promise<Entry|undefined> {

  const subject = email.headers.get("subject");
  if (!subject)
    return undefined;

  const matchSubject = (pattern: RegExp) => subject?.match(pattern)?.[1];

  const amount = matchSubject(/\$([0-9]+\.[0-9]{2})/);
  const account = matchSubject(/\(…([0-9]{4})\)/);

  if (!amount || !account)
    return undefined;

  var date: Date = new Date();
  var description = "";

  class ChaseEmailParser implements HTMLRewriterElementContentHandlers {
    last_message: string = "";

    amount: Number|undefined = undefined;
    account: string|undefined = undefined;
    date: Date|undefined = undefined;
    description: string|undefined = undefined;

    text(text: Text) {
      const tx = text.text.trim();
      if (tx === "")
        return;

      switch (tx) {
      case "Account ending in": {
        break;
      }
      case "Made on": {
        date = new Date(Date.parse(tx));
        break;
      }
      case "Description": {
        description = tx;
        break;
      }
      case "Amount": {
        break;
      }
      }

      this.last_message = tx;
    }
  }

  await new HTMLRewriter()
      .on("td", new ChaseEmailParser())
      .transform(new Response(email.raw))
      .arrayBuffer();

  return {
    amount: Number(amount), title: description, account: account,
        postedDate: date
  }
}

interface Entry {
  amount: number;
  title: string;
  account: string;
  postedDate: Date;
}

export default {
  async email(message: ForwardableEmailMessage, _env: {},
              _ctx: ExecutionContext) { const content = parseChaseEmail(email);}
}
