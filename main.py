import logging
import os
from base64 import b64encode
from functools import lru_cache
from html import escape
from pathlib import Path
from time import perf_counter

import msal
import requests
from dotenv import load_dotenv
from fastapi import BackgroundTasks, FastAPI
from pydantic import BaseModel

load_dotenv()

if not logging.getLogger().handlers:
    logging.basicConfig(level=logging.INFO)

app = FastAPI()
logger = logging.getLogger("dreamit.leads")

CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
TENANT_ID = os.getenv("TENANT_ID")
SITE_ID = os.getenv("SITE_ID")

FROM_EMAIL = os.getenv("FROM_EMAIL")
TO_EMAIL = os.getenv("TO_EMAIL")
CC_EMAIL = os.getenv("CC_EMAIL")

AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPE = ["https://graph.microsoft.com/.default"]

BASE_DIR = Path(__file__).resolve().parent
LOGO_PATH = BASE_DIR / "logo.png"
INLINE_LOGO_CID = "dreamit-logo"
FALLBACK_LOGO_URL = "https://www.dreamitcs.com/wp-content/uploads/2023/05/logo.png"
REQUEST_TIMEOUT = (5, 20)


class Lead(BaseModel):
    firstName: str
    lastName: str
    email: str
    selectService: str
    messages: str


@lru_cache(maxsize=1)
def get_msal_client():
    return msal.ConfidentialClientApplication(
        CLIENT_ID,
        authority=AUTHORITY,
        client_credential=CLIENT_SECRET,
    )


def get_token():
    token = get_msal_client().acquire_token_for_client(scopes=SCOPE)

    if "access_token" not in token:
        raise Exception(f"Token Error: {token}")

    return token["access_token"]


def save_to_excel(token, lead):
    try:
        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
        }

        session_url = (
            f"https://graph.microsoft.com/v1.0/sites/{SITE_ID}/drive/root:"
            "/Recruitment/Leads.xlsx:/workbook/createSession"
        )
        session = requests.post(
            session_url,
            headers=headers,
            json={"persistChanges": True},
            timeout=REQUEST_TIMEOUT,
        )

        if session.status_code not in [200, 201]:
            return f"Session Error: {session.text}"

        session_id = session.json()["id"]
        headers["workbook-session-id"] = session_id

        insert_url = (
            f"https://graph.microsoft.com/v1.0/sites/{SITE_ID}/drive/root:"
            "/Recruitment/Leads.xlsx:/workbook/tables/Table1/rows/add"
        )
        data = {
            "values": [[
                lead.firstName,
                lead.lastName,
                lead.email,
                lead.selectService,
                lead.messages,
            ]]
        }

        res = requests.post(
            insert_url,
            headers=headers,
            json=data,
            timeout=REQUEST_TIMEOUT,
        )

        if res.status_code not in [200, 201]:
            return f"Excel Error: {res.text}"

        return "Excel Updated"
    except Exception as exc:
        return str(exc)


def build_recipients(addresses):
    if not addresses:
        return []

    if isinstance(addresses, str):
        addresses = [item.strip() for item in addresses.split(",") if item.strip()]

    recipients = []

    for address in addresses:
        cleaned = str(address).strip()
        if cleaned:
            recipients.append({"emailAddress": {"address": cleaned}})

    return recipients


@lru_cache(maxsize=1)
def get_inline_logo_attachment():
    if not LOGO_PATH.exists():
        return None

    with LOGO_PATH.open("rb") as logo_file:
        return {
            "@odata.type": "#microsoft.graph.fileAttachment",
            "name": LOGO_PATH.name,
            "contentType": "image/png",
            "contentBytes": b64encode(logo_file.read()).decode("utf-8"),
            "contentId": INLINE_LOGO_CID,
            "isInline": True,
        }


def format_html_text(value):
    cleaned = (value or "").strip()
    return escape(cleaned) if cleaned else "Not provided"


def format_html_message(value):
    cleaned = (value or "").strip()

    if not cleaned:
        return "Not provided"

    return escape(cleaned).replace("\n", "<br>")


def build_details_table(full_name, email, service, message):
    label_style = (
        "padding:10px 0; width:140px; vertical-align:top; "
        "font-size:14px; line-height:22px; font-weight:600; color:#102a43;"
    )
    value_style = (
        "padding:10px 0; vertical-align:top; "
        "font-size:15px; line-height:24px; color:#334e68;"
    )

    return f"""
        <table role="presentation" width="100%" cellpadding="0" cellspacing="0"
            style="width:100%; border-collapse:collapse; margin:0 0 24px 0;">
            <tr>
                <td style="{label_style}">Name</td>
                <td style="{value_style}">{full_name}</td>
            </tr>
            <tr>
                <td style="{label_style}">Email</td>
                <td style="{value_style}">{email}</td>
            </tr>
            <tr>
                <td style="{label_style}">Service</td>
                <td style="{value_style}">{service}</td>
            </tr>
            <tr>
                <td style="{label_style}">Message</td>
                <td style="{value_style}">{message}</td>
            </tr>
        </table>
    """


def build_email_layout(title, content_html, logo_src):
    return f"""
    <!DOCTYPE html>
    <html>
        <body style="margin:0; padding:0; background-color:#f3f6fb;">
            <table role="presentation" width="100%" cellpadding="0" cellspacing="0"
                style="width:100%; background-color:#f3f6fb; padding:32px 16px;">
                <tr>
                    <td align="center">
                        <table role="presentation" width="640" cellpadding="0" cellspacing="0"
                            style="width:100%; max-width:640px; background-color:#ffffff;
                            border:1px solid #d9e2ec; border-radius:16px;">
                            <tr>
                                <td style="padding:36px 40px;
                                    font-family:'Segoe UI', Arial, Helvetica, sans-serif;
                                    color:#243b53;">
                                    <img src="{logo_src}" alt="Dream IT Consulting Services"
                                        width="170"
                                        style="display:block; width:170px; max-width:100%;
                                        height:auto; border:0; outline:none; text-decoration:none;
                                        margin:0 0 28px 0;" />
                                    <div style="margin:0 0 12px 0; font-size:25px; line-height:34px;
                                        font-weight:700; color:#102a43;">
                                        {title}
                                    </div>
                                    {content_html}
                                    <div style="margin-top:28px; padding-top:18px;
                                        border-top:1px solid #e5edf5; font-size:13px;
                                        line-height:22px; color:#7b8794;">
                                        Dream IT Consulting Services Pvt. Ltd.<br>
                                        <a href="https://www.dreamitcs.com/"
                                            style="color:#0b66c3; text-decoration:none;">Website</a>
                                        <span style="color:#c5ced8;"> | </span>
                                        <a href="https://www.linkedin.com/company/dreamitcs/posts/?feedView=all"
                                            style="color:#0b66c3; text-decoration:none;">LinkedIn</a>
                                    </div>
                                </td>
                            </tr>
                        </table>
                    </td>
                </tr>
            </table>
        </body>
    </html>
    """


def build_lead_context(lead_data):
    full_name = format_html_text(
        " ".join(
            part.strip()
            for part in [lead_data.get("firstName"), lead_data.get("lastName")]
            if part and str(part).strip()
        )
    )
    first_name = format_html_text(lead_data.get("firstName"))
    email = format_html_text(lead_data.get("email"))
    service = format_html_text(lead_data.get("selectService"))
    message = format_html_message(lead_data.get("messages"))
    details_table = build_details_table(full_name, email, service, message)

    logo_attachment = get_inline_logo_attachment()
    attachments = [logo_attachment] if logo_attachment else None
    logo_src = f"cid:{INLINE_LOGO_CID}" if logo_attachment else FALLBACK_LOGO_URL

    return {
        "first_name": first_name,
        "details_table": details_table,
        "service": service,
        "logo_src": logo_src,
        "attachments": attachments,
    }


def send_email(token, subject, html, recipients, cc=None, attachments=None):
    url = f"https://graph.microsoft.com/v1.0/users/{FROM_EMAIL}/sendMail"

    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
    }

    message = {
        "subject": subject,
        "body": {
            "contentType": "HTML",
            "content": html,
        },
        "toRecipients": build_recipients(recipients),
        "ccRecipients": build_recipients(cc),
    }

    if attachments:
        message["attachments"] = attachments

    body = {"message": message}

    try:
        res = requests.post(
            url,
            headers=headers,
            json=body,
            timeout=REQUEST_TIMEOUT,
        )
    except requests.RequestException as exc:
        return f"Email Error: {exc}"

    if res.status_code != 202:
        return f"Email Error: {res.text}"

    return "Email Sent"


def send_lead_notifications(lead_data):
    started_at = perf_counter()

    try:
        token = get_token()
        context = build_lead_context(lead_data)

        admin_content = f"""
            <p style="margin:0 0 18px 0; font-size:15px; line-height:24px; color:#486581;">
                A new lead has been submitted through the website. The details are below.
            </p>
            {context["details_table"]}
            <p style="margin:0; font-size:14px; line-height:22px; color:#7b8794;">
                This notification was generated automatically from the lead form.
            </p>
        """

        admin_html = build_email_layout(
            "New Lead Notification",
            admin_content,
            context["logo_src"],
        )

        admin_status = send_email(
            token,
            "New Lead from Website",
            admin_html,
            [TO_EMAIL],
            CC_EMAIL,
            context["attachments"],
        )

        user_content = f"""
            <p style="margin:0 0 16px 0; font-size:15px; line-height:24px; color:#486581;">
                Dear {context["first_name"]},
            </p>
            <p style="margin:0 0 16px 0; font-size:15px; line-height:24px; color:#486581;">
                Thank you for contacting <strong>Dream IT Consulting Services Pvt. Ltd.</strong>.
            </p>
            <p style="margin:0 0 22px 0; font-size:15px; line-height:24px; color:#486581;">
                We have received your request regarding <strong>{context["service"]}</strong>.
                Our team will review it and get back to you shortly.
            </p>
            <div style="margin:0 0 10px 0; font-size:16px; line-height:24px;
                font-weight:600; color:#102a43;">
                Submitted Details
            </div>
            {context["details_table"]}
            <p style="margin:0 0 18px 0; font-size:15px; line-height:24px; color:#486581;">
                If you have any questions, simply reply to this email and our team will assist you.
            </p>
            <p style="margin:0; font-size:15px; line-height:24px; color:#486581;">
                <strong>Regards,</strong><br>
                Dream IT Consulting Services Pvt. Ltd.<br>
                <a href="mailto:connect@dreamitcs.com"
                    style="color:#0b66c3; text-decoration:none;">connect@dreamitcs.com</a>
            </p>
        """

        user_html = build_email_layout(
            "Thank You for Reaching Out",
            user_content,
            context["logo_src"],
        )

        user_status = send_email(
            token,
            "Thank You for Contacting Dream IT",
            user_html,
            [lead_data.get("email")],
            attachments=context["attachments"],
        )

        logger.info(
            "Lead background processing finished in %.2fs | admin=%s | user=%s",
            perf_counter() - started_at,
            admin_status,
            user_status,
        )
    except Exception:
        logger.exception("Lead background processing failed")


@app.post("/lead")
def create_lead(lead: Lead, background_tasks: BackgroundTasks):
    lead_data = lead.model_dump()
    background_tasks.add_task(send_lead_notifications, lead_data)

    return {
        "status": "accepted",
        "message": "Lead received. Email processing queued in background.",
        "admin_email": "queued",
        "user_email": "queued",
    }
