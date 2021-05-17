addEventListener("fetch", (event) => {
  event.respondWith(
    handleRequest(event.request).catch((err) => handleError(err))
  );
});

const corsHeaders = {
  "Access-Control-Allow-Headers": "*",
  "Access-Control-Allow-Methods": "POST",
  "Access-Control-Allow-Origin": "*",
};

async function handleRequest(request) {
  if (request.method === "OPTIONS") {
    return new Response("OK", {
      headers: corsHeaders,
    });
  }
  if (request.method === "POST") {
    const { userEmail, recipient, message } = await request.json();
    const token = await getToken();
    const response = await sendEmail(userEmail, recipient, message, token);
    return new Response("ok", {
      status: 200,
      headers: { "Content-Type": "text/html", ...corsHeaders },
    });
  }
  return new Response(
    "Bad Request, please use a Post request with Content-Type: application/JSON",
    { status: 400 }
  );
}

async function sendEmail(userEmail, recipient, message, token) {
  if (recipient === 'moritz@mortimerbaltus.com' || recipient === 'theo@mortimerbaltus.com') {
  recipientName = recipient.split("@")[0];
  emailRequest = await fetch(
    "https://graph.microsoft.com/v1.0/users/contact@mortimerbaltus.com/sendMail",
    {
      method: "POST",
      headers: new Headers({
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json",
      }),
      body: JSON.stringify({
        message: {
          subject: `New message for ${recipientName} from ${userEmail}
            `,
          body: {
            contentType: "HTML",
            content: `<h3>Hey ${recipientName}, you have a new message from: ${userEmail}</h3>
                            <hr/>
                            <p>${message}</p>
                            <hr/>
                            <footer><p>This email was sent automatically using the Contact Form on MortimerBaltus.com. However, you can just hit reply to reach ${userEmail}</p></footer>`,
          },
          toRecipients: [
            {
              emailAddress: {
                address: recipient,
              },
            },
          ],
          replyTo: [
            {
              emailAddress: {
                address: userEmail,
              },
            },
          ],
        },
      }),
    }
  ).then((res) => {
    if (!res.ok) {
      throw new Error(
        `Failed to send email! ${res.status}: ${res.statusText}.`
      );
    }
    return res;
  });
  return emailRequest;} else {
    throw new Error(`You have no permission to send an email to this repicient!`)
  }
}

async function getToken() {
  const tokenResponse = await fetch(
    `https://login.microsoftonline.com/45b420bd-dcbc-4abd-b11e-0c80c057b7db/oauth2/v2.0/token`,
    {
      method: "POST",
      headers: {
        "Content-type": "application/x-www-form-urlencoded; charset=utf-8",
      },
      body: new URLSearchParams({
        client_id: CLIENT_ID,
        scope: "https://graph.microsoft.com/.default",
        client_secret: CLIENT_SECRET,
        grant_type: "client_credentials",
      }),
    }
  )
    .then((res) => {
      if (!res.ok) {
        throw new Error(
          `Failed to get token! ${res.status}: ${res.statusText}.`
        );
      }
      return res;
    })
    .then((res) => res.json());
  return tokenResponse.access_token;
}

function handleError(err) {
  console.log("uncaught" + err);
  return new Response(err.message, {
    status: 500,
    headers: { "Content-Type": "text/html", ...corsHeaders },
  });
}
