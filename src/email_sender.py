from pathlib import Path


class OutlookMailer:
    def __init__(self):
        self.outlook = None

    def check(self):
        try:
            import pythoncom
            import win32com.client

            pythoncom.CoInitialize()
            self.outlook = win32com.client.Dispatch("Outlook.Application")
            session = self.outlook.Session

            if session.Accounts.Count == 0:
                return False, "Outlook найден, но ни один аккаунт не подключён."

            accounts = []
            for i in range(1, session.Accounts.Count + 1):
                acc = session.Accounts.Item(i)
                address = getattr(acc, "SmtpAddress", "") or getattr(acc, "DisplayName", "")
                accounts.append(address)

            return True, f"Outlook готов. Аккаунты: {', '.join(accounts)}"
        except Exception as e:
            return False, f"Outlook не установлен или недоступен: {e}"

    def send(self, to_email, subject, body, attachment_path=None, display_only=False):
        try:
            import pythoncom
            import win32com.client

            pythoncom.CoInitialize()
            outlook = win32com.client.Dispatch("Outlook.Application")
            mail = outlook.CreateItem(0)

            mail.To = to_email
            mail.Subject = subject
            mail.Body = body

            if attachment_path:
                mail.Attachments.Add(str(Path(attachment_path).resolve()))

            if display_only:
                mail.Display()
                return "Письмо открыто в окне Outlook для проверки."

            mail.Send()
            return "Письмо успешно отправлено."
        except Exception as e:
            raise RuntimeError(f"Не удалось отправить письмо через Outlook: {e}")