export interface GmailMessageSnippet {
  id: string;
  threadId: string;
  snippet: string;
  subject: string;
  from: string;
  date: string;
  bodyText?: string;
  parsedTrade?: {
    ticker: string;
    quantity: number;
    price: number;
    action: 'BUY' | 'SELL';
    type: 'equity' | 'options';
  } | null;
}

// Base64Url encode helper for standard string mime-type
function base64urlEncode(str: string): string {
  return btoa(unescape(encodeURIComponent(str)))
    .replace(/\+/g, '-')
    .replace(/\//g, '_')
    .replace(/=+$/, '');
}

/**
 * Lists the user's recent messages matching a query (e.g. "broker", "Zerodha", "trade")
 */
export const listRecentEmails = async (
  accessToken: string,
  query: string = 'subject:(trade OR broker OR order OR transaction OR contract OR Zerodha OR Groww OR Fyers OR AngelOne OR Upstox)'
): Promise<GmailMessageSnippet[]> => {
  try {
    const listUrl = `https://gmail.googleapis.com/gmail/v1/users/me/messages?maxResults=8&q=${encodeURIComponent(query)}`;
    const listRes = await fetch(listUrl, {
      headers: {
        Authorization: `Bearer ${accessToken}`,
      },
    });

    if (!listRes.ok) {
      throw new Error(`Gmail API list failed: ${listRes.statusText}`);
    }

    const listData = await listRes.json();
    if (!listData.messages || listData.messages.length === 0) {
      return [];
    }

    const messages: GmailMessageSnippet[] = [];

    // Fetch details for each message
    for (const msgRef of listData.messages) {
      const detailUrl = `https://gmail.googleapis.com/gmail/v1/users/me/messages/${msgRef.id}`;
      const detailRes = await fetch(detailUrl, {
        headers: {
          Authorization: `Bearer ${accessToken}`,
        },
      });

      if (detailRes.ok) {
        const detailData = await detailRes.json();
        const headers = detailData.payload?.headers || [];
        
        const subject = headers.find((h: any) => h.name.toLowerCase() === 'subject')?.value || '(No Subject)';
        const from = headers.find((h: any) => h.name.toLowerCase() === 'from')?.value || 'Unknown Sender';
        const dateHeader = headers.find((h: any) => h.name.toLowerCase() === 'date')?.value || '';
        
        let bodyText = detailData.snippet || '';
        
        // Formulate a parsed action based on email contents for utility
        let parsedTrade = null;
        const textToAnalyze = `${subject} ${bodyText}`.toLowerCase();
        
        // Simple regex parser for common Indian stock/options trade formats
        // Example subject: "Executed: Buy 100 shares of IDFCFIRSTB"
        // Example text: "Your purchase of 50 shares of RELIANCE at Avg ₹2450.25 was filled"
        let ticker = 'NIFTY';
        let quantity = 50;
        let price = 150;
        let action: 'BUY' | 'SELL' = 'BUY';
        let isMatched = false;

        if (textToAnalyze.includes('buy') || textToAnalyze.includes('purchased') || textToAnalyze.includes('bought')) {
          action = 'BUY';
          isMatched = true;
        } else if (textToAnalyze.includes('sell') || textToAnalyze.includes('sold')) {
          action = 'SELL';
          isMatched = true;
        }

        // Ticker discovery
        const symbols = ['RELIANCE', 'TCS', 'INFY', 'HDFCBANK', 'ICICIBANK', 'NIFTY', 'TAX', 'SBIN', 'TATAMOTORS', 'SENSEX', 'ZOMATO', 'ITC'];
        for (const sym of symbols) {
          if (textToAnalyze.toUpperCase().includes(sym)) {
            ticker = sym;
            isMatched = true;
            break;
          }
        }

        // Extract numbers for quantity & price if matched
        const qtyMatch = textToAnalyze.match(/(?:qty|quantity|shares|number|vol|volume|lots?)\s*[:=]?\s*(\d+)/i) || textToAnalyze.match(/(\d+)\s*(?:shares|lots|qty)/i);
        if (qtyMatch && qtyMatch[1]) {
          quantity = parseInt(qtyMatch[1], 10);
        }

        const priceMatch = textToAnalyze.match(/(?:at|price|avg|rate|rs\.?|₹)\s*[:=]?\s*(?:rs|₹)?\s*(\d+(?:\.\d+)?)/i);
        if (priceMatch && priceMatch[1]) {
          price = parseFloat(priceMatch[1]);
        }

        if (isMatched) {
          parsedTrade = {
            ticker,
            quantity,
            price,
            action,
            type: textToAnalyze.includes('call') || textToAnalyze.includes('put') || textToAnalyze.includes('ce') || textToAnalyze.includes('pe') ? 'options' : 'equity' as any
          };
        }

        messages.push({
          id: detailData.id,
          threadId: detailData.threadId,
          snippet: detailData.snippet || '',
          subject,
          from,
          date: dateHeader ? new Date(dateHeader).toLocaleDateString() : 'N/A',
          bodyText,
          parsedTrade
        });
      }
    }

    return messages;
  } catch (error) {
    console.error('Error listing recent emails:', error);
    throw error;
  }
};

/**
 * Sends a custom trading journal report via Gmail
 */
export const sendTradingReport = async (
  accessToken: string,
  toEmail: string,
  subject: string,
  htmlContent: string
): Promise<void> => {
  try {
    const emailHeaderLines = [
      `To: ${toEmail}`,
      `Subject: ${subject}`,
      'Content-Type: text/html; charset=utf-8',
      'MIME-Version: 1.0',
      ''
    ];
    
    const emailStr = emailHeaderLines.join('\r\n') + htmlContent;
    const raw64 = base64urlEncode(emailStr);

    const sendRes = await fetch('https://gmail.googleapis.com/gmail/v1/users/me/messages/send', {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${accessToken}`,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        raw: raw64
      })
    });

    if (!sendRes.ok) {
      const errorText = await sendRes.text();
      throw new Error(`Failed to send email: ${errorText || sendRes.statusText}`);
    }
  } catch (error) {
    console.error('Error sending trading report:', error);
    throw error;
  }
};
