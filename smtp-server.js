import "dotenv/config";
import express from "express";
import cors from "cors";
import nodemailer from "nodemailer";

const app = express();
const port = Number(process.env.SMTP_API_PORT || 3001);

app.use(cors());
app.use(express.json({ limit: "1mb" }));

function requireEnv(name) {
  const value = process.env[name];
  if (!value) {
    throw new Error(`Missing env variable: ${name}`);
  }
  return value;
}

function getTransporter() {
  const gmailUser = process.env.GMAIL_USER;
  const gmailAppPassword = process.env.GMAIL_APP_PASSWORD;

  if (gmailUser && gmailAppPassword) {
    return nodemailer.createTransport({
      service: "gmail",
      auth: {
        user: gmailUser,
        pass: gmailAppPassword
      }
    });
  }

  const host = requireEnv("SMTP_HOST");
  const portValue = Number(process.env.SMTP_PORT || 587);
  const user = requireEnv("SMTP_USER");
  const pass = requireEnv("SMTP_PASS");
  const secure = String(process.env.SMTP_SECURE || "false").toLowerCase() === "true";

  return nodemailer.createTransport({
    host,
    port: portValue,
    secure,
    auth: { user, pass }
  });
}

app.get("/api/health", (_req, res) => {
  res.json({ ok: true, service: "smtp-alert-api" });
});

app.post("/api/stock-alert", async (req, res) => {
  try {
    const apiKey = process.env.ALERT_API_KEY;
    if (apiKey) {
      const headerKey = req.header("x-alert-key");
      if (headerKey !== apiKey) {
        return res.status(401).json({ ok: false, error: "Unauthorized" });
      }
    }

    const { toEmail, subject, html } = req.body || {};
    if (!toEmail || !subject || !html) {
      return res.status(400).json({ ok: false, error: "Missing toEmail/subject/html" });
    }

    const fromEmail = process.env.SMTP_FROM || process.env.GMAIL_USER || requireEnv("SMTP_USER");
    const transporter = getTransporter();

    await transporter.sendMail({
      from: fromEmail,
      to: toEmail,
      subject,
      html
    });

    return res.json({ ok: true });
  } catch (error) {
    console.error("SMTP alert error", error);
    return res.status(500).json({ ok: false, error: "Failed to send email" });
  }
});

app.listen(port, () => {
  console.log(`SMTP alert API running on http://localhost:${port}`);
});
