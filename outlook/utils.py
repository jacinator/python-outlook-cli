from msgraph.generated.models.recipient import Recipient


def get_emails(recipients: list[Recipient] | None) -> list[str]:
    return [
        z for x in (recipients or ()) if (y := x.email_address) and (z := y.address)
    ]


def get_emails_str(recipients: list[Recipient] | None) -> str:
    if emails := get_emails(recipients):
        return ",".join(emails)
    return "NONE"


def get_from_str(from_: Recipient | None) -> str:
    if from_ and from_.email_address:
        name: str = from_.email_address.name or ""
        addr: str = from_.email_address.address or ""
        return f"{name} <{addr}>" if name else addr
    return "NONE"
