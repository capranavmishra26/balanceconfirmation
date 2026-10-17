"""Send email through the Outlook desktop app installed on this PC (COM)."""
from __future__ import annotations

from pathlib import Path

import pythoncom
import win32com.client

PR_SEND_USING_ACCOUNT = 64209  # DISPID of MailItem.SendUsingAccount


class Outlook:
    """Use as a context manager so COM is initialised on the calling thread."""

    def __enter__(self) -> "Outlook":
        pythoncom.CoInitialize()
        self.app = win32com.client.Dispatch("Outlook.Application")
        return self

    def __exit__(self, *exc) -> None:
        self.app = None
        pythoncom.CoUninitialize()

    def accounts(self) -> list[str]:
        return [a.SmtpAddress for a in self.app.Session.Accounts]

    def send(self, *, to: str, cc: str, subject: str, html_body: str,
             attachments: list[Path], from_account: str | None = None) -> None:
        mail = self.app.CreateItem(0)  # olMailItem
        mail.To = to
        if cc:
            mail.CC = cc
        mail.Subject = subject
        mail.HTMLBody = html_body
        for path in attachments:
            mail.Attachments.Add(str(Path(path).resolve()))
        if from_account:
            for acc in self.app.Session.Accounts:
                if acc.SmtpAddress.lower() == from_account.lower():
                    # Setting SendUsingAccount directly fails with late binding.
                    mail._oleobj_.Invoke(PR_SEND_USING_ACCOUNT, 0, 8, 0, acc)
                    break
        mail.Send()


def list_accounts() -> list[str]:
    try:
        with Outlook() as ol:
            return ol.accounts()
    except Exception:
        return []
