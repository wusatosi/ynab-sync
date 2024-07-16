import {
  ExecutionContext,
  ForwardableEmailMessage,
  HTMLRewriterElementContentHandlers,
  Text
} from "@cloudflare/workers-types"

import type {HTMLRewriter, Response} from "@cloudflare/workers-types"

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
        // Match (...1234)
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
          this.amount = parseFloat(match[1]);
        break;
      }
      }
    }

    last_message: string = "";
    working_text = "";
    text(text: Text) {
      this.working_text = this.working_text.concat(text.text);
      this.working_text = this.working_text.trim();

      if (text.lastInTextNode) {
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
