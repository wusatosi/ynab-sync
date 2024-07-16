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
  class ChaseEmailParser implements HTMLRewriterElementContentHandlers {
    last_message: string = "";

    amount: number|undefined = undefined;
    account: string|undefined = undefined;
    date: Date|undefined = undefined;
    description: string|undefined = undefined;

    handleNewText(text: string) {
      const tx = text.trim();
      if (tx === "")
        return;

      switch (tx) {
      case "Account ending in": {
        const match = tx.match(/\(\.\.\.(\d{4})\)/);
        if (match && (match.length > 1))
          this.account = match[1];
        break;
      }
      case "Made on": {
        this.date = new Date(Date.parse(tx));
        break;
      }
      case "Description": {
        this.description = tx;
        break;
      }
      case "Amount": {
        const match = tx.match(/\$?(\d+\.\d{2})/);
        if (match && (match.length > 1))
          this.amount = parseFloat(match[1]);
        break;
      }
      }

      this.last_message = tx;
    }

    working_text = "";
    text(text: Text) {
      this.working_text = this.working_text.concat(text.text);
      if (text.lastInTextNode) {
        this.handleNewText(this.working_text);
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

export default {
  async email(email: ForwardableEmailMessage, _env: {},
              _ctx: ExecutionContext) {
    const content = await parseChaseEmail(email);
    if (!content) {
      console.warn(`Cannot process email from ${email.from} to ${email.to}`);
      email.setReject("Cannot process");
    } else {
      console.debug(`Parsed email: from ${email.from} to ${email.to}`, content);
    }
  }
}
