type Email = `${string}@${string}.${string}`;
type Status = "Sent" | "sent" | undefined | null | (string & {});

type EmailBatchEntry = { email: Email; rowNum: number };

type EmailRow = [Email, Status];
