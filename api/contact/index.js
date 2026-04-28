module.exports = async function (context, req) {
  const method = (req.method || "GET").toUpperCase();
  const timestamp = new Date().toISOString();

  context.log(`[contact] Hit at ${timestamp} with method=${method}`);

  if (method === "OPTIONS") {
    context.log("[contact] Responding to CORS preflight OPTIONS request.");
    return {
      status: 204,
      headers: {
        "Access-Control-Allow-Origin": "*",
        "Access-Control-Allow-Methods": "GET,POST,OPTIONS",
        "Access-Control-Allow-Headers": "Content-Type"
      }
    };
  }

  if (method === "GET") {
    context.log("[contact] Returning GET help payload for browser testing.");
    return {
      status: 200,
      headers: {
        "Content-Type": "application/json"
      },
      body: {
        ok: true,
        route: "/api/contact",
        message: "Contact endpoint is live. Send a POST with JSON to submit contact details.",
        expectedPostBody: {
          name: "string",
          email: "string",
          message: "string"
        },
        timestamp
      }
    };
  }

  if (method === "POST") {
    const payload = req.body || {};
    const name = (payload.name || "").toString().trim();
    const email = (payload.email || "").toString().trim();
    const message = (payload.message || "").toString().trim();

    context.log(
      `[contact] POST payload received name=${name || "(missing)"} email=${email || "(missing)"} messageLength=${message.length}`
    );

    if (!name || !email || !message) {
      context.log("[contact] Validation failed: missing one or more required fields.");
      return {
        status: 400,
        headers: {
          "Content-Type": "application/json"
        },
        body: {
          ok: false,
          error: "Missing required fields: name, email, message",
          timestamp
        }
      };
    }

    context.log("[contact] Validation passed. Returning success response.");
    return {
      status: 200,
      headers: {
        "Content-Type": "application/json"
      },
      body: {
        ok: true,
        message: "Contact request received.",
        received: {
          name,
          email,
          message
        },
        timestamp
      }
    };
  }

  context.log(`[contact] Method ${method} not allowed.`);
  return {
    status: 405,
    headers: {
      "Content-Type": "application/json"
    },
    body: {
      ok: false,
      error: `Method ${method} not allowed`,
      allowedMethods: ["GET", "POST", "OPTIONS"],
      timestamp
    }
  };
};
